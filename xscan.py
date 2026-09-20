"""A small PySide6 front end for Windows Image Acquisition (WIA)."""

from __future__ import annotations

import re
import sys
import json
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path

from PySide6.QtCore import Qt, QUrl
from PySide6.QtGui import QDesktopServices
from PySide6.QtWidgets import (
    QApplication,
    QComboBox,
    QFileDialog,
    QFormLayout,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMainWindow,
    QMessageBox,
    QPushButton,
    QVBoxLayout,
    QWidget,
)

try:
    import pythoncom
    import win32com.client
except ImportError:  # Allow the window to explain the missing optional dependency.
    pythoncom = None
    win32com = None


WIA_SCANNER_DEVICE_TYPE = 1
WIA_JPEG_FORMAT = "{B96B3CAE-0728-11D3-9D7B-0000F81EF32E}"
INVALID_FILENAME_CHARS = re.compile(r'[<>:"/\\|?*\x00-\x1f]')
CONFIG_DIR = Path.home() / "Tool_Config" / "xScan"
CONFIG_FILE = CONFIG_DIR / "setting.json"


@dataclass(frozen=True)
class Scanner:
    name: str
    device_id: str


def installed_scanners() -> list[Scanner]:
    """Return scanners registered with Windows Image Acquisition."""
    if win32com is None:
        raise RuntimeError("pywin32 is not installed. Run: pip install pywin32")

    manager = win32com.client.Dispatch("WIA.DeviceManager")
    scanners: list[Scanner] = []
    for index in range(1, manager.DeviceInfos.Count + 1):
        info = manager.DeviceInfos.Item(index)
        if info.Type != WIA_SCANNER_DEVICE_TYPE:
            continue
        try:
            name = str(info.Properties("Name").Value)
        except Exception:
            name = f"Scanner {index}"
        scanners.append(Scanner(name=name, device_id=str(info.DeviceID)))
    return scanners


def next_output_path(folder: Path, prefix: str) -> Path:
    """Build a readable, non-overwriting JPEG path."""
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    candidate = folder / f"{prefix}_{stamp}.jpg"
    suffix = 2
    while candidate.exists():
        candidate = folder / f"{prefix}_{stamp}_{suffix}.jpg"
        suffix += 1
    return candidate


def scan_to_file(scanner: Scanner, destination: Path) -> None:
    """Acquire the first scanner item and save it as a JPEG."""
    if win32com is None or pythoncom is None:
        raise RuntimeError("pywin32 is not installed. Run: pip install pywin32")

    pythoncom.CoInitialize()
    try:
        manager = win32com.client.Dispatch("WIA.DeviceManager")
        device_info = None
        for index in range(1, manager.DeviceInfos.Count + 1):
            candidate = manager.DeviceInfos.Item(index)
            if str(candidate.DeviceID) == scanner.device_id:
                device_info = candidate
                break
        if device_info is None:
            raise RuntimeError("The selected scanner is no longer available.")

        device = device_info.Connect()
        if device.Items.Count < 1:
            raise RuntimeError("The scanner did not expose a scan source.")

        dialog = win32com.client.Dispatch("WIA.CommonDialog")
        image = dialog.ShowTransfer(device.Items.Item(1), WIA_JPEG_FORMAT, False)
        if image is None:
            raise RuntimeError("The scan was cancelled.")
        image.SaveFile(str(destination))
    finally:
        pythoncom.CoUninitialize()


class ScannerWindow(QMainWindow):
    def __init__(self) -> None:
        super().__init__()
        self.scanners: list[Scanner] = []
        self.settings = self.load_settings()
        self.scan_count = 0
        self.count_offset = 0
        self._count_folder: Path | None = None
        self._initial_count = True
        self.setWindowTitle("xScan")
        self.setMinimumWidth(560)
        size_info = self.settings.get("sizeInfo")
        if isinstance(size_info, list) and len(size_info) == 4:
            self.setGeometry(*size_info)

        self.scanner_combo = QComboBox()
        self.refresh_button = QPushButton("Refresh")
        self.refresh_button.clicked.connect(self.refresh_scanners)

        scanner_row = QHBoxLayout()
        scanner_row.addWidget(self.scanner_combo, 1)
        scanner_row.addWidget(self.refresh_button)

        self.folder_edit = QLineEdit(
            str(self.settings.get("rootFolder", Path.home() / "Pictures" / "Scans"))
        )
        self.folder_edit.editingFinished.connect(self.refresh_count_from_folder)
        self.folder_button = QPushButton("Browse…")
        self.folder_button.clicked.connect(self.choose_folder)
        folder_row = QHBoxLayout()
        folder_row.addWidget(self.folder_edit, 1)
        folder_row.addWidget(self.folder_button)

        self.subfolder_edit = QLineEdit(str(self.settings.get("subfolder", "scan")))
        self.subfolder_edit.setPlaceholderText("e.g. invoice")
        self.subfolder_edit.editingFinished.connect(self.refresh_count_from_folder)

        form = QFormLayout()
        form.setLabelAlignment(Qt.AlignmentFlag.AlignRight)
        form.addRow("Scanner:", scanner_row)
        form.addRow("Root folder:", folder_row)
        form.addRow("Subfolder / prefix:", self.subfolder_edit)

        self.scan_button = QPushButton("Scan")
        self.scan_button.setDefault(True)
        self.scan_button.setMinimumHeight(38)
        self.scan_button.clicked.connect(self.start_scan)

        self.count_label = QLabel(f"Scans: {self.scan_count}")
        self.reset_button = QPushButton("Reset count")
        self.reset_button.clicked.connect(self.reset_count)
        self.open_folder_button = QPushButton("Open target folder")
        self.open_folder_button.clicked.connect(self.open_target_folder)

        actions_row = QHBoxLayout()
        actions_row.addWidget(self.count_label)
        actions_row.addWidget(self.reset_button)
        actions_row.addStretch()
        actions_row.addWidget(self.open_folder_button)

        self.status_label = QLabel("Ready")
        self.status_label.setWordWrap(True)

        layout = QVBoxLayout()
        layout.setContentsMargins(20, 20, 20, 20)
        layout.setSpacing(14)
        layout.addLayout(form)
        layout.addWidget(self.scan_button)
        layout.addLayout(actions_row)
        layout.addWidget(self.status_label)

        container = QWidget()
        container.setLayout(layout)
        self.setCentralWidget(container)
        self.refresh_count_from_folder()
        self.refresh_scanners()

    @staticmethod
    def load_settings() -> dict:
        try:
            with CONFIG_FILE.open("r", encoding="utf-8") as stream:
                data = json.load(stream)
            return data if isinstance(data, dict) else {}
        except (OSError, json.JSONDecodeError):
            return {}

    def save_settings(self) -> None:
        geometry = self.geometry()
        selected = (
            self.scanners[self.scanner_combo.currentIndex()].device_id
            if 0 <= self.scanner_combo.currentIndex() < len(self.scanners)
            else ""
        )
        data = {
            "sizeInfo": [
                geometry.x(),
                geometry.y(),
                geometry.width(),
                geometry.height(),
            ],
            "rootFolder": self.folder_edit.text().strip(),
            "subfolder": self.subfolder_edit.text().strip(),
            "scannerDeviceId": selected,
            "scanCount": self.scan_count,
            "scanCountOffset": self.count_offset,
            "scanCountFolder": str(self._count_folder) if self._count_folder else "",
        }
        try:
            CONFIG_DIR.mkdir(parents=True, exist_ok=True)
            with CONFIG_FILE.open("w", encoding="utf-8") as stream:
                json.dump(data, stream)
        except OSError as exc:
            self.status_label.setText(f"Could not save settings: {exc}")

    def closeEvent(self, event) -> None:
        self.save_settings()
        super().closeEvent(event)

    def choose_folder(self) -> None:
        chosen = QFileDialog.getExistingDirectory(
            self, "Choose scan folder", self.folder_edit.text()
        )
        if chosen:
            self.folder_edit.setText(chosen)
            self.refresh_count_from_folder()

    @staticmethod
    def count_jpg_files(folder: Path | None) -> int:
        if folder is None:
            return 0
        try:
            return sum(
                entry.is_file() and entry.suffix.lower() == ".jpg"
                for entry in folder.iterdir()
            )
        except OSError:
            return 0

    def refresh_count_from_folder(self) -> None:
        folder = self.target_folder(show_errors=False)
        if not self._initial_count and folder == self._count_folder:
            return
        saved_offset = self.settings.get("scanCountOffset", 0)
        self.count_offset = (
            saved_offset
            if self._initial_count
            and str(folder) == self.settings.get("scanCountFolder")
            and isinstance(saved_offset, int)
            and not isinstance(saved_offset, bool)
            else 0
        )
        self.scan_count = max(0, self.count_jpg_files(folder) + self.count_offset)
        self._count_folder = folder
        self._initial_count = False
        self.count_label.setText(f"Scans: {self.scan_count}")

    def reset_count(self) -> None:
        self.refresh_count_from_folder()
        self.count_offset = -self.count_jpg_files(self._count_folder)
        self.scan_count = 0
        self.count_label.setText("Scans: 0")
        self.save_settings()

    def target_folder(self, show_errors: bool = True) -> Path | None:
        subfolder = self.subfolder_edit.text().strip().rstrip(". ")
        folder_text = self.folder_edit.text().strip()
        if not folder_text:
            if show_errors:
                self.show_error("Choose a destination folder.")
            return None
        if (
            not subfolder
            or INVALID_FILENAME_CHARS.search(subfolder)
            or subfolder in {".", ".."}
        ):
            if show_errors:
                self.show_error(
                    'The subfolder cannot be empty or contain < > : " / \\ | ? *.'
                )
            return None
        return Path(folder_text).expanduser() / subfolder

    def open_target_folder(self) -> None:
        folder = self.target_folder()
        if folder is None:
            return
        self.refresh_count_from_folder()
        try:
            folder.mkdir(parents=True, exist_ok=True)
        except OSError as exc:
            self.show_error(f"Cannot use that folder:\n{exc}")
            return
        if not QDesktopServices.openUrl(QUrl.fromLocalFile(str(folder.resolve()))):
            self.show_error(f"Could not open the target folder:\n{folder}")

    def refresh_scanners(self) -> None:
        self.scanner_combo.clear()
        try:
            self.scanners = installed_scanners()
        except Exception as exc:
            self.scanners = []
            self.status_label.setText(str(exc))
        else:
            self.scanner_combo.addItems(scanner.name for scanner in self.scanners)
            saved_device_id = str(self.settings.get("scannerDeviceId", ""))
            for index, scanner in enumerate(self.scanners):
                if scanner.device_id == saved_device_id:
                    self.scanner_combo.setCurrentIndex(index)
                    break
            self.status_label.setText(
                f"Found {len(self.scanners)} scanner(s)."
                if self.scanners
                else "No WIA scanners found. Check the connection and click Refresh."
            )
        self.scan_button.setEnabled(bool(self.scanners))

    def start_scan(self) -> None:
        index = self.scanner_combo.currentIndex()
        if not (0 <= index < len(self.scanners)):
            self.show_error("Choose an available scanner.")
            return
        folder = self.target_folder()
        if folder is None:
            return
        self.refresh_count_from_folder()

        subfolder = folder.name
        try:
            folder.mkdir(parents=True, exist_ok=True)
            destination = next_output_path(folder, subfolder)
        except OSError as exc:
            self.show_error(f"Cannot use that folder:\n{exc}")
            return

        self.set_busy(True)
        self.save_settings()
        self.status_label.setText("Waiting for the scanner…")
        QApplication.processEvents()
        try:
            scan_to_file(self.scanners[index], destination)
        except Exception as exc:
            self.show_error(f"Scanning failed:\n{exc}")
        else:
            self.scan_count += 1
            self.count_label.setText(f"Scans: {self.scan_count}")
            self.save_settings()
            self.status_label.setText(f"Saved: {destination}")
            QMessageBox.information(self, "Scan complete", f"Saved to:\n{destination}")
        finally:
            self.set_busy(False)

    def set_busy(self, busy: bool) -> None:
        self.scan_button.setEnabled(not busy and bool(self.scanners))
        self.refresh_button.setEnabled(not busy)
        self.folder_button.setEnabled(not busy)
        self.open_folder_button.setEnabled(not busy)
        self.reset_button.setEnabled(not busy)

    def show_error(self, message: str) -> None:
        self.status_label.setText(message.replace("\n", " "))
        QMessageBox.critical(self, "xScan", message)


def main() -> int:
    app = QApplication(sys.argv)
    app.setApplicationName("xScan")
    window = ScannerWindow()
    window.show()
    return app.exec()


if __name__ == "__main__":
    raise SystemExit(main())
