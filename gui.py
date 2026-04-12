#!/usr/bin/env python3
"""Jingle browser GUI with category filters and playback device selection."""

from __future__ import annotations

import json
import sys
import ctypes
import hashlib
import os
import shutil
import subprocess
import wave
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Callable, Iterator


def _ensure_qt_logging_rules() -> None:
    """Suppress noisy Qt internal warnings from multimedia backend teardown."""
    suppress_rules = [
        "qt.core.qobject.connect.warning=false",
        "qt.multimedia.ffmpeg.warning=false",
    ]
    current = os.environ.get("QT_LOGGING_RULES", "").strip()
    existing = [part.strip() for part in current.split(";") if part.strip()]
    merged = list(existing)
    for rule in suppress_rules:
        if rule not in merged:
            merged.append(rule)
    os.environ["QT_LOGGING_RULES"] = ";".join(merged)


_ensure_qt_logging_rules()

_HERE = Path(sys._MEIPASS) if hasattr(sys, "_MEIPASS") else Path(__file__).resolve().parent  # type: ignore[attr-defined]


def _apply_windows_taskbar_icon(window: QMainWindow) -> None:
    if sys.platform != "win32":
        return
    icon_source = Path(sys.executable) if getattr(sys, "frozen", False) else (_HERE / "icon.ico")
    if not icon_source.exists():
        return
    hwnd = int(window.winId())
    shell32 = ctypes.windll.shell32
    user32 = ctypes.windll.user32
    large = ctypes.c_void_p()
    small = ctypes.c_void_p()
    extracted = shell32.ExtractIconExW(str(icon_source), 0, ctypes.byref(large), ctypes.byref(small), 1)
    if extracted <= 0:
        return
    WM_SETICON = 0x0080
    ICON_SMALL = 0
    ICON_BIG = 1
    GCLP_HICON = -14
    GCLP_HICONSM = -34
    if small.value:
        user32.SendMessageW(hwnd, WM_SETICON, ICON_SMALL, small.value)
        user32.SetClassLongPtrW(hwnd, GCLP_HICONSM, small.value)
    if large.value:
        user32.SendMessageW(hwnd, WM_SETICON, ICON_BIG, large.value)
        user32.SetClassLongPtrW(hwnd, GCLP_HICON, large.value)

from PyQt6.QtCore import (
    QEasingCurve,
    QEvent,
    QFileSystemWatcher,
    QObject,
    QPropertyAnimation,
    QSettings,
    QTimer,
    Qt,
    QStandardPaths,
    QUrl,
)
from PyQt6.QtGui import QAction, QCursor, QIcon, QKeyEvent, QKeySequence, QMouseEvent, QPixmap
from PyQt6.QtWidgets import (
    QApplication,
    QComboBox,
    QDialog,
    QDialogButtonBox,
    QFileDialog,
    QMenu,
    QGridLayout,
    QGraphicsOpacityEffect,
    QHBoxLayout,
    QLabel,
    QInputDialog,
    QKeySequenceEdit,
    QLineEdit,
    QMainWindow,
    QMessageBox,
    QPlainTextEdit,
    QPushButton,
    QSlider,
    QStatusBar,
    QHeaderView,
    QSizePolicy,
    QScrollArea,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)

_has_qt_multimedia = False
try:
    from PyQt6.QtMultimedia import QAudioOutput, QMediaDevices, QMediaPlayer

    _has_qt_multimedia = True
except ModuleNotFoundError:
    pass

AUDIO_EXTENSIONS = {
    ".wav",
    ".mp3",
    ".ogg",
    ".flac",
    ".m4a",
    ".aac",
    ".wma",
    ".aiff",
    ".aif",
}

DUPLICATE_FORMAT_PRIORITY = {
    ".wav": 0,
    ".flac": 1,
    ".aiff": 2,
    ".aif": 2,
    ".m4a": 3,
    ".aac": 4,
    ".ogg": 5,
    ".wma": 6,
    ".mp3": 7,
}

ORG_NAME = "JingleAllTheDay"
APP_NAME = "JingleAllTheDay"
APP_VERSION = "1.1.0.041226"


def _runtime_app_dir() -> Path:
    """Return the folder where the running app is located.

    In frozen builds this is the executable directory; in debug/dev this is the script directory.
    """
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent

DEFAULT_KEYBOARD_SHORTCUTS: dict[str, str] = {
    "rename": "F2",
    "delete": "Delete",
    "skip_previous": "Left",
    "skip_next": "Right",
    "select_up": "Up",
    "select_down": "Down",
}

MEDIA_PLAY_KEYS = tuple(
    key
    for key in (
        getattr(Qt.Key, "Key_MediaPlay", None),
        getattr(Qt.Key, "Key_AudioPlay", None),
    )
    if key is not None
)
MEDIA_PAUSE_KEYS = tuple(
    key
    for key in (
        getattr(Qt.Key, "Key_MediaPause", None),
        getattr(Qt.Key, "Key_AudioPause", None),
    )
    if key is not None
)
MEDIA_TOGGLE_PLAYBACK_KEYS = tuple(
    key
    for key in (
        getattr(Qt.Key, "Key_MediaTogglePlayPause", None),
        getattr(Qt.Key, "Key_AudioPlay", None),
    )
    if key is not None
)
MEDIA_NEXT_KEYS = tuple(
    key
    for key in (
        getattr(Qt.Key, "Key_MediaNext", None),
        getattr(Qt.Key, "Key_AudioForward", None),
    )
    if key is not None
)
MEDIA_PREVIOUS_KEYS = tuple(
    key
    for key in (
        getattr(Qt.Key, "Key_MediaPrevious", None),
        getattr(Qt.Key, "Key_AudioRewind", None),
    )
    if key is not None
)


def _normalize_tags(raw: str | list[str] | tuple[str, ...] | None) -> list[str]:
    if raw is None:
        return []

    if isinstance(raw, str):
        # Accept comma or semicolon separated tags.
        pieces = raw.replace(";", ",").split(",")
    else:
        pieces = [str(item) for item in raw]

    out: list[str] = []
    seen: set[str] = set()
    for piece in pieces:
        tag = piece.strip()
        if not tag:
            continue
        key = tag.casefold()
        if key in seen:
            continue
        seen.add(key)
        out.append(tag)
    return out


def _tags_to_text(tags: list[str]) -> str:
    return ", ".join(tags)


def _merge_tags(existing: list[str], incoming: list[str]) -> list[str]:
    out: list[str] = []
    seen: set[str] = set()
    for tag in [*existing, *incoming]:
        clean = tag.strip()
        if not clean:
            continue
        key = clean.casefold()
        if key in seen:
            continue
        seen.add(key)
        out.append(clean)
    return out


def _remove_tags(existing: list[str], incoming: list[str]) -> list[str]:
    remove_keys = {tag.casefold() for tag in incoming}
    if not remove_keys:
        return list(existing)
    return [tag for tag in existing if tag.casefold() not in remove_keys]


def _coerce_volume_percent(value: Any, default: int = 100) -> int:
    try:
        percent = int(round(float(value)))
    except (TypeError, ValueError):
        percent = default
    return max(0, min(100, percent))


def _chip_palette_for_tag_seed(tag_seed: str) -> tuple[str, str, str]:
    palettes = [
        ("#3b2f6b", "#5a4a96", "#4a3a82"),
        ("#204d4f", "#2f7073", "#296063"),
        ("#5a3e2b", "#7f5a3d", "#6d4f36"),
        ("#2f4f2f", "#4c7a4c", "#3f683f"),
        ("#5c2f4f", "#844571", "#6f3d60"),
        ("#2f3f6a", "#455d99", "#3a5185"),
        ("#4f4a2a", "#7a7340", "#686234"),
        ("#4a2f2f", "#734646", "#613b3b"),
    ]
    digest = hashlib.sha1(tag_seed.casefold().encode("utf-8")).digest()
    idx = digest[0] % len(palettes)
    return palettes[idx]


def _format_size_label(total_bytes: int) -> str:
    size = float(max(0, total_bytes))
    units = ["B", "KB", "MB", "GB", "TB"]
    unit_idx = 0
    while size >= 1000.0 and unit_idx < len(units) - 1:
        size /= 1000.0
        unit_idx += 1

    if unit_idx == 0:
        return f"{int(size)}{units[unit_idx]}"
    return f"{size:.2f}{units[unit_idx]}"


def _format_duration_hms(total_seconds: float) -> str:
    sec = max(0, int(round(total_seconds)))
    hours, rem = divmod(sec, 3600)
    minutes, seconds = divmod(rem, 60)
    return f"{hours:02d}:{minutes:02d}:{seconds:02d}"


_FFPROBE_PATH: str | None = None
_FFPROBE_CHECKED = False


def _get_ffprobe_path() -> str | None:
    global _FFPROBE_PATH, _FFPROBE_CHECKED
    if _FFPROBE_CHECKED:
        return _FFPROBE_PATH
    _FFPROBE_CHECKED = True
    _FFPROBE_PATH = shutil.which("ffprobe")
    return _FFPROBE_PATH


def _probe_duration_seconds(path: Path) -> float:
    suffix = path.suffix.lower()
    if suffix in {".wav", ".wave"}:
        try:
            with wave.open(str(path), "rb") as w:
                frames = w.getnframes()
                rate = w.getframerate()
                if rate > 0:
                    return float(frames) / float(rate)
        except (wave.Error, OSError):
            pass

    ffprobe = _get_ffprobe_path()
    if not ffprobe:
        return 0.0

    cmd = [
        ffprobe,
        "-v",
        "error",
        "-show_entries",
        "format=duration",
        "-of",
        "default=noprint_wrappers=1:nokey=1",
        str(path),
    ]
    try:
        run_kwargs: dict[str, Any] = {
            "capture_output": True,
            "text": True,
            "timeout": 2.0,
            "check": False,
        }
        if hasattr(subprocess, "CREATE_NO_WINDOW"):
            run_kwargs["creationflags"] = subprocess.CREATE_NO_WINDOW
        result = subprocess.run(cmd, **run_kwargs)
        if result.returncode != 0:
            return 0.0
        value = (result.stdout or "").strip()
        if not value:
            return 0.0
        return max(0.0, float(value))
    except (OSError, ValueError, subprocess.TimeoutExpired):
        return 0.0


@dataclass
class JingleRecord:
    path: Path
    categories: list[str]
    size_bytes: int = 0
    duration_seconds: float = 0.0

    @property
    def name(self) -> str:
        return self.path.stem

    @property
    def folder(self) -> str:
        parent = self.path.parent.name
        return parent if parent else str(self.path.parent)


class LibraryStore:
    def __init__(self, json_path: Path) -> None:
        self._json_path = json_path
        self._entries: dict[str, dict[str, Any]] = {}
        self._load()

    def _load(self) -> None:
        if not self._json_path.exists():
            self._entries = {}
            return
        try:
            payload = json.loads(self._json_path.read_text(encoding="utf-8"))
        except (OSError, json.JSONDecodeError):
            self._entries = {}
            return

        self._entries = self._entries_from_payload(payload)

    @staticmethod
    def _entries_from_payload(payload: object) -> dict[str, dict[str, Any]]:
        raw_items = payload.get("items", {}) if isinstance(payload, dict) else {}
        entries: dict[str, dict[str, Any]] = {}
        if isinstance(raw_items, dict):
            for path_key, info in raw_items.items():
                if not isinstance(path_key, str) or not isinstance(info, dict):
                    continue

                # Backward-compatible read: old shape used single category/subcategory strings.
                category_raw = info.get("categories")
                if category_raw is None:
                    category_raw = info.get("category", "")

                subcategory_raw = info.get("subcategories")
                if subcategory_raw is None:
                    subcategory_raw = info.get("subcategory", "")

                categories = _merge_tags(
                    _normalize_tags(category_raw),
                    _normalize_tags(subcategory_raw),
                )

                entry: dict[str, Any] = {
                    "categories": categories,
                }

                size_raw = info.get("size_bytes")
                duration_raw = info.get("duration_seconds")
                mtime_raw = info.get("mtime_ns")
                try:
                    if size_raw is not None:
                        entry["size_bytes"] = max(0, int(size_raw))
                except (TypeError, ValueError):
                    pass
                try:
                    if duration_raw is not None:
                        entry["duration_seconds"] = max(0.0, float(duration_raw))
                except (TypeError, ValueError):
                    pass
                try:
                    if mtime_raw is not None:
                        entry["mtime_ns"] = max(0, int(mtime_raw))
                except (TypeError, ValueError):
                    pass

                entries[path_key] = entry
        return entries

    def save(self) -> None:
        self._json_path.parent.mkdir(parents=True, exist_ok=True)
        payload = {"version": 4, "items": self._entries}
        self._json_path.write_text(json.dumps(payload, indent=2), encoding="utf-8")

    def export_to(self, destination: Path) -> None:
        destination.parent.mkdir(parents=True, exist_ok=True)
        payload = {"version": 4, "items": self._entries}
        destination.write_text(json.dumps(payload, indent=2), encoding="utf-8")

    def import_from(self, source: Path) -> None:
        payload = json.loads(source.read_text(encoding="utf-8"))
        self._entries = self._entries_from_payload(payload)

    def get(self, path: Path) -> list[str]:
        info = self._entries.get(str(path), {})
        categories = _normalize_tags(info.get("categories", []))
        return categories

    def set(self, path: Path, categories: list[str]) -> None:
        key = str(path)
        entry = dict(self._entries.get(key, {}))
        entry["categories"] = _normalize_tags(categories)
        self._entries[key] = entry

    def remove(self, path: Path) -> None:
        self._entries.pop(str(path), None)

    def rename(self, source: Path, destination: Path) -> None:
        source_key = str(source)
        destination_key = str(destination)
        entry = dict(self._entries.get(source_key, {}))
        if source_key in self._entries:
            self._entries.pop(source_key, None)
        if entry:
            self._entries[destination_key] = entry

    def get_media_cache(self, path: Path) -> tuple[int, float, int] | None:
        info = self._entries.get(str(path), {})
        try:
            size_bytes = int(info.get("size_bytes"))
            duration_seconds = float(info.get("duration_seconds"))
            mtime_ns = int(info.get("mtime_ns"))
        except (TypeError, ValueError):
            return None

        if size_bytes < 0 or duration_seconds < 0.0 or mtime_ns < 0:
            return None
        return size_bytes, duration_seconds, mtime_ns

    def set_media_cache(self, path: Path, size_bytes: int, duration_seconds: float, mtime_ns: int) -> None:
        key = str(path)
        entry = dict(self._entries.get(key, {}))
        entry["size_bytes"] = max(0, int(size_bytes))
        entry["duration_seconds"] = max(0.0, float(duration_seconds))
        entry["mtime_ns"] = max(0, int(mtime_ns))
        self._entries[key] = entry

    def iter_entries(self) -> Iterator[tuple[str, dict[str, Any]]]:
        for path_key, info in self._entries.items():
            if isinstance(path_key, str) and isinstance(info, dict):
                yield path_key, info

    def sync_with_files(self, files: list[Path]) -> None:
        keep = {str(path) for path in files}
        self._entries = {k: v for k, v in self._entries.items() if k in keep}
        for path in files:
            key = str(path)
            if key not in self._entries:
                self._entries[key] = {"categories": []}


class OptionsDialog(QDialog):
    def __init__(
        self,
        live_output_device: str,
        preview_output_device: str,
        live_volume_percent: int,
        preview_volume_percent: int,
        samples_dir: Path | None = None,
        parent: QWidget | None = None,
    ) -> None:
        super().__init__(parent)
        self.setWindowTitle("Options")
        self.setMinimumWidth(560)

        root = QVBoxLayout(self)

        folder_row = QHBoxLayout()
        folder_label = QLabel("Samples Folder")
        folder_label.setFixedWidth(120)
        folder_row.addWidget(folder_label)
        self._folder_edit = QLineEdit(str(samples_dir) if samples_dir else "")
        self._folder_edit.setReadOnly(True)
        self._folder_edit.setPlaceholderText("No folder selected")
        folder_row.addWidget(self._folder_edit)
        folder_browse_btn = QPushButton("Browse")
        folder_browse_btn.clicked.connect(self._on_browse_folder)
        folder_row.addWidget(folder_browse_btn)
        root.addLayout(folder_row)

        live_row = QHBoxLayout()
        live_label = QLabel("Live Device")
        live_label.setFixedWidth(100)
        live_row.addWidget(live_label)

        self._live_device_combo = QComboBox()
        self._live_device_combo.setMinimumWidth(340)
        self._live_device_combo.setToolTip("Audio output used in Live mode.")
        live_row.addWidget(self._live_device_combo)

        root.addLayout(live_row)

        preview_row = QHBoxLayout()
        preview_label = QLabel("Preview Device")
        preview_label.setFixedWidth(100)
        preview_row.addWidget(preview_label)

        self._preview_device_combo = QComboBox()
        self._preview_device_combo.setMinimumWidth(340)
        self._preview_device_combo.setToolTip("Audio output used in Preview mode.")
        preview_row.addWidget(self._preview_device_combo)

        root.addLayout(preview_row)

        live_volume_row = QHBoxLayout()
        live_volume_label = QLabel("Live Volume")
        live_volume_label.setFixedWidth(100)
        live_volume_row.addWidget(live_volume_label)

        self._live_volume_slider = QSlider(Qt.Orientation.Horizontal)
        self._live_volume_slider.setRange(0, 100)
        self._live_volume_slider.setPageStep(5)
        self._live_volume_slider.setValue(_coerce_volume_percent(live_volume_percent))
        self._live_volume_slider.setToolTip("Playback volume used in Live mode.")
        live_volume_row.addWidget(self._live_volume_slider)

        self._live_volume_value_label = QLabel()
        self._live_volume_value_label.setFixedWidth(44)
        live_volume_row.addWidget(self._live_volume_value_label)

        root.addLayout(live_volume_row)

        preview_volume_row = QHBoxLayout()
        preview_volume_label = QLabel("Preview Volume")
        preview_volume_label.setFixedWidth(100)
        preview_volume_row.addWidget(preview_volume_label)

        self._preview_volume_slider = QSlider(Qt.Orientation.Horizontal)
        self._preview_volume_slider.setRange(0, 100)
        self._preview_volume_slider.setPageStep(5)
        self._preview_volume_slider.setValue(_coerce_volume_percent(preview_volume_percent))
        self._preview_volume_slider.setToolTip("Playback volume used in Preview mode.")
        preview_volume_row.addWidget(self._preview_volume_slider)

        self._preview_volume_value_label = QLabel()
        self._preview_volume_value_label.setFixedWidth(44)
        preview_volume_row.addWidget(self._preview_volume_value_label)

        root.addLayout(preview_volume_row)

        refresh_btn = QPushButton("Refresh Devices")
        refresh_btn.clicked.connect(self._on_refresh_clicked)
        refresh_row = QHBoxLayout()
        refresh_row.addWidget(refresh_btn)
        refresh_row.addStretch()

        root.addLayout(refresh_row)

        buttons = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel
        )
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        root.addWidget(buttons)

        self._populate_devices(live_output_device, preview_output_device)
        self._live_volume_slider.valueChanged.connect(self._sync_volume_labels)
        self._preview_volume_slider.valueChanged.connect(self._sync_volume_labels)
        self._sync_volume_labels()

    def _populate_devices(self, live_selected: str, preview_selected: str) -> None:
        self._populate_device_combo(self._live_device_combo, live_selected)
        self._populate_device_combo(self._preview_device_combo, preview_selected)

    def _populate_device_combo(self, combo: QComboBox, selected_device: str) -> None:
        combo.blockSignals(True)
        combo.clear()

        default_name = ""
        if _has_qt_multimedia:
            try:
                default_name = QMediaDevices.defaultAudioOutput().description().strip()
                seen: set[str] = set()
                for device in QMediaDevices.audioOutputs():
                    name = device.description().strip()
                    if not name or name in seen:
                        continue
                    seen.add(name)
                    combo.addItem(name, name)
            except Exception:
                pass

        default_label = "System Default"
        if default_name:
            default_label = f"System Default ({default_name})"
        combo.insertItem(0, default_label, "")

        target = selected_device.strip()
        if target:
            idx = combo.findData(target)
            if idx >= 0:
                combo.setCurrentIndex(idx)
            else:
                combo.addItem(f"{target} (Unavailable)", target)
                combo.setCurrentIndex(combo.count() - 1)
        else:
            combo.setCurrentIndex(0)

        combo.blockSignals(False)

    def _on_refresh_clicked(self) -> None:
        live_current = self._live_device_combo.currentData()
        preview_current = self._preview_device_combo.currentData()
        live_selected = str(live_current).strip() if live_current is not None else ""
        preview_selected = str(preview_current).strip() if preview_current is not None else ""
        self._populate_devices(live_selected, preview_selected)

    def selected_devices(self) -> tuple[str, str]:
        live = self._live_device_combo.currentData()
        preview = self._preview_device_combo.currentData()
        live_value = str(live).strip() if live is not None else ""
        preview_value = str(preview).strip() if preview is not None else ""
        return live_value, preview_value

    def selected_volumes(self) -> tuple[int, int]:
        return (
            _coerce_volume_percent(self._live_volume_slider.value()),
            _coerce_volume_percent(self._preview_volume_slider.value()),
        )

    def selected_folder(self) -> Path | None:
        text = self._folder_edit.text().strip()
        if not text:
            return None
        p = Path(text)
        return p if p.exists() and p.is_dir() else None

    def _on_browse_folder(self) -> None:
        current = self._folder_edit.text().strip()
        start = current if current else str(Path.home())
        selected = QFileDialog.getExistingDirectory(self, "Choose Samples Folder", start)
        if not selected:
            return
        path = Path(selected)
        if path.exists() and path.is_dir():
            self._folder_edit.setText(str(path))

    def _sync_volume_labels(self) -> None:
        self._live_volume_value_label.setText(f"{self._live_volume_slider.value()}%")
        self._preview_volume_value_label.setText(f"{self._preview_volume_slider.value()}%")


class DeselectableTableWidget(QTableWidget):
    """Clears selection when clicking blank table whitespace."""

    def __init__(self, rows: int, columns: int) -> None:
        super().__init__(rows, columns)
        self._preserve_selection_callback: Callable[[], bool] | None = None

    def set_preserve_selection_callback(self, callback: Callable[[], bool]) -> None:
        self._preserve_selection_callback = callback

    def keyPressEvent(self, event: QKeyEvent | None) -> None:
        # Let arrow keys propagate to the parent window for global shortcuts
        if event is not None and event.key() in (
            Qt.Key.Key_Left,
            Qt.Key.Key_Right,
            Qt.Key.Key_Up,
            Qt.Key.Key_Down,
        ):
            event.ignore()
            return
        super().keyPressEvent(event)

    def mousePressEvent(self, event: QMouseEvent | None) -> None:
        if event is not None and event.button() == Qt.MouseButton.LeftButton:
            pt = event.position().toPoint()
            row = self.rowAt(pt.y())
            col = self.columnAt(pt.x())
            if row < 0 or col < 0:
                if self._preserve_selection_callback is not None and self._preserve_selection_callback():
                    event.accept()
                    return
                self.clearSelection()
                self.setCurrentCell(-1, -1)
                event.accept()
                return
        super().mousePressEvent(event)


class KeyboardShortcutsDialog(QDialog):
    def __init__(
        self,
        current_shortcuts: dict[str, str],
        default_shortcuts: dict[str, str],
        parent: QWidget | None = None,
    ) -> None:
        super().__init__(parent)
        self.setWindowTitle("Edit Keyboard Shortcuts")
        self.resize(520, 280)

        root = QVBoxLayout(self)

        help_label = QLabel("Set shortcut keys. Leave a field empty to disable that shortcut.")
        help_label.setWordWrap(True)
        root.addWidget(help_label)

        grid = QGridLayout()
        root.addLayout(grid)

        self._shortcut_editors: dict[str, QKeySequenceEdit] = {}
        self._row_labels: dict[str, str] = {}
        rows = [
            ("rename", "Rename"),
            ("delete", "Delete"),
            ("skip_previous", "Skip to Previous (while playing/paused)"),
            ("skip_next", "Skip to Next (while playing/paused)"),
            ("select_up", "Select Previous Row (when stopped)"),
            ("select_down", "Select Next Row (when stopped)"),
        ]
        for row_idx, (key, label_text) in enumerate(rows):
            label = QLabel(label_text)
            grid.addWidget(label, row_idx, 0)

            editor = QKeySequenceEdit()
            editor.setKeySequence(QKeySequence(current_shortcuts.get(key, default_shortcuts.get(key, ""))))
            grid.addWidget(editor, row_idx, 1)
            self._shortcut_editors[key] = editor
            self._row_labels[key] = label_text

            default_label = QLabel(f"Default: {default_shortcuts.get(key, '')}")
            default_label.setStyleSheet("color: #666;")
            grid.addWidget(default_label, row_idx, 2)

        controls_row = QHBoxLayout()
        reset_defaults_btn = QPushButton("Reset to Defaults")
        reset_defaults_btn.clicked.connect(self._reset_to_defaults)
        controls_row.addWidget(reset_defaults_btn)
        controls_row.addStretch()
        root.addLayout(controls_row)

        self._buttons = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel
        )
        self._buttons.accepted.connect(self._on_accept)
        self._buttons.rejected.connect(self.reject)
        root.addWidget(self._buttons)

        self._default_shortcuts = dict(default_shortcuts)

    def _reset_to_defaults(self) -> None:
        for key, editor in self._shortcut_editors.items():
            editor.setKeySequence(QKeySequence(self._default_shortcuts.get(key, "")))

    def selected_shortcuts(self) -> dict[str, str]:
        out: dict[str, str] = {}
        for key, editor in self._shortcut_editors.items():
            out[key] = editor.keySequence().toString(QKeySequence.SequenceFormat.PortableText)
        return out

    def _find_conflicts(self) -> list[tuple[str, list[str]]]:
        assignments: dict[str, list[str]] = {}
        for key, editor in self._shortcut_editors.items():
            seq_text = editor.keySequence().toString(QKeySequence.SequenceFormat.PortableText).strip()
            if not seq_text:
                continue
            assignments.setdefault(seq_text, []).append(self._row_labels.get(key, key))

        conflicts: list[tuple[str, list[str]]] = []
        for seq_text, labels in assignments.items():
            if len(labels) > 1:
                conflicts.append((seq_text, labels))
        return conflicts

    def _on_accept(self) -> None:
        conflicts = self._find_conflicts()
        if conflicts:
            lines = ["These shortcuts are assigned to multiple actions:", ""]
            for seq_text, labels in conflicts:
                lines.append(f"{seq_text}: {', '.join(labels)}")
            lines.append("")
            lines.append("Please resolve conflicts before saving.")
            QMessageBox.warning(self, "Shortcut Conflict", "\n".join(lines))
            return
        self.accept()


class AboutDialog(QDialog):
    def __init__(
        self,
        *,
        library_count: int,
        library_duration_seconds: float,
        library_size_bytes: int,
        revision_log_path: Path | None,
        parent: QWidget | None = None,
    ) -> None:
        super().__init__(parent)
        self.setWindowTitle("About JingleAllTheDay")
        self.setMinimumWidth(620)
        self._revision_log_path = revision_log_path

        icon_path = _HERE / "icon.png"
        self._icon_base = QPixmap(str(icon_path)) if icon_path.is_file() else QPixmap()

        root = QHBoxLayout(self)
        root.setContentsMargins(14, 14, 14, 14)
        root.setSpacing(16)

        icon_wrap = QWidget(self)
        icon_layout = QVBoxLayout(icon_wrap)
        icon_layout.setContentsMargins(0, 0, 0, 0)
        icon_layout.addStretch(1)

        self._icon_label = QLabel(icon_wrap)
        self._icon_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self._icon_label.setMinimumSize(1, 1)
        self._icon_label.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)
        icon_layout.addWidget(self._icon_label, 0, Qt.AlignmentFlag.AlignHCenter)
        icon_layout.addStretch(1)

        text_wrap = QWidget(self)
        text_layout = QVBoxLayout(text_wrap)
        text_layout.setContentsMargins(0, 0, 0, 0)
        text_layout.setSpacing(10)

        self._title_label = QLabel(APP_NAME)
        self._title_label.setStyleSheet("font-size: 20px; font-weight: 700;")
        text_layout.addWidget(self._title_label)

        self._version_label = QLabel(f"Version {APP_VERSION}")
        self._version_label.setAlignment(Qt.AlignmentFlag.AlignTop | Qt.AlignmentFlag.AlignLeft)
        intro_layout = QVBoxLayout()
        intro_layout.setContentsMargins(0, 0, 0, 0)
        intro_layout.setSpacing(2)
        intro_layout.addWidget(self._version_label)

        self._body_label = QLabel(
            "\nJingleAllTheDay is a desktop jingle library manager and playback board.\n\n"
            "You can scan and organize jingles, tag tracks by category, search and filter quickly, "
            "and trigger reliable playback in Live or Preview mode with selectable output devices.\n\n"
            "The app is built for fast on-air workflows, with keyboard shortcuts, batch tag tools, "
            "duplicate detection, and import/export of your tag database."
        )
        self._body_label.setWordWrap(True)
        self._body_label.setAlignment(Qt.AlignmentFlag.AlignTop | Qt.AlignmentFlag.AlignLeft)
        self._body_label.setMinimumWidth(430)
        intro_layout.addWidget(self._body_label)
        text_layout.addLayout(intro_layout)

        count_label = f"{max(0, int(library_count)):,}"
        duration_label = _format_duration_hms(max(0.0, float(library_duration_seconds)))
        size_label = _format_size_label(max(0, int(library_size_bytes)))

        self._library_summary_label = QLabel(
            "Current Library\n"
            f"Jingles: {count_label}\n"
            f"Total Duration: {duration_label}\n"
            f"Total Size: {size_label}"
        )
        self._library_summary_label.setWordWrap(True)
        self._library_summary_label.setAlignment(Qt.AlignmentFlag.AlignTop | Qt.AlignmentFlag.AlignLeft)
        text_layout.addWidget(self._library_summary_label)

        button_row = QHBoxLayout()
        button_row.addStretch(1)
        self._show_revision_history_btn = QPushButton("Show Revision History")
        self._show_revision_history_btn.setEnabled(
            self._revision_log_path is not None and self._revision_log_path.is_file()
        )
        if not self._show_revision_history_btn.isEnabled():
            self._show_revision_history_btn.setToolTip("rev.log was not found in the runtime application folder.")
        self._show_revision_history_btn.clicked.connect(self._on_show_revision_history)
        button_row.addWidget(self._show_revision_history_btn)
        text_layout.addLayout(button_row)

        root.addWidget(icon_wrap, 1)
        root.addWidget(text_wrap, 3)

        text_height = (
            self._title_label.sizeHint().height()
            + text_layout.spacing()
            + self._version_label.sizeHint().height()
            + text_layout.spacing()
            + self._body_label.sizeHint().height()
            + text_layout.spacing()
            + self._library_summary_label.sizeHint().height()
            + text_layout.spacing()
            + self._show_revision_history_btn.sizeHint().height()
        )
        frame_height = root.contentsMargins().top() + root.contentsMargins().bottom() + 8
        initial_height = max(220, text_height + frame_height)
        self.resize(640, initial_height)
        self._refresh_icon()
        self.setFixedSize(self.size())

    def _on_show_revision_history(self) -> None:
        if self._revision_log_path is None or not self._revision_log_path.is_file():
            QMessageBox.information(
                self,
                "Revision History",
                "rev.log was not found in the runtime application folder.",
            )
            return

        try:
            revision_text = self._revision_log_path.read_text(encoding="utf-8")
        except OSError as exc:
            QMessageBox.critical(
                self,
                "Revision History",
                f"Unable to read revision history file.\n\n{exc}",
            )
            return

        dialog = RevisionHistoryDialog(revision_text=revision_text, parent=self)
        dialog.exec()

    def _refresh_icon(self) -> None:
        if self._icon_base.isNull():
            self._icon_label.setText("No icon")
            return

        max_by_height = max(48, self.height() - 80)
        max_by_width = max(48, self.width() // 4)
        side_cap = max(100, self.height() // 3)
        side = max(48, min(max_by_height, max_by_width, side_cap))
        scaled = self._icon_base.scaled(
            side,
            side,
            Qt.AspectRatioMode.KeepAspectRatio,
            Qt.TransformationMode.SmoothTransformation,
        )
        self._icon_label.setText("")
        self._icon_label.setPixmap(scaled)


class RevisionHistoryDialog(QDialog):
    def __init__(self, *, revision_text: str, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self.setWindowTitle("Revision History")
        self.setMinimumSize(700, 420)

        root = QVBoxLayout(self)
        title = QLabel(f"{APP_NAME} {APP_VERSION}")
        title.setStyleSheet("font-weight: 700;")
        root.addWidget(title)

        history_view = QPlainTextEdit(self)
        history_view.setReadOnly(True)
        history_view.setPlainText(revision_text)
        root.addWidget(history_view, 1)

        buttons = QDialogButtonBox(QDialogButtonBox.StandardButton.Close)
        buttons.rejected.connect(self.reject)
        buttons.accepted.connect(self.accept)
        root.addWidget(buttons)


class MainWindow(QMainWindow):
    def __init__(self) -> None:
        super().__init__()
        self.setWindowTitle("JingleAllTheDay")
        self.resize(1200, 740)
        _icon_path = _HERE / "icon.png"
        if _icon_path.exists():
            self.setWindowIcon(QIcon(str(_icon_path)))

        app_data_location = QStandardPaths.writableLocation(
            QStandardPaths.StandardLocation.AppDataLocation
        )
        self._app_data_dir = Path(app_data_location)
        self._app_data_dir.mkdir(parents=True, exist_ok=True)
        self._settings = self._create_settings_store()
        self._output_device = str(self._settings.value("options/outputDevice", "")).strip()
        self._preview_output_device = str(self._settings.value("options/previewOutputDevice", "")).strip()
        self._live_volume_percent = _coerce_volume_percent(
            self._settings.value("options/liveVolumePercent", 100)
        )
        self._preview_volume_percent = _coerce_volume_percent(
            self._settings.value("options/previewVolumePercent", 100)
        )
        self._samples_dir: Path | None = self._load_samples_dir()
        self._auto_folder_tags: bool = self._load_auto_folder_tags()
        self._watch_library_changes: bool = self._load_watch_library_changes()
        self._keyboard_shortcuts = self._load_keyboard_shortcuts()

        self._rename_action: QAction | None = None
        self._delete_action: QAction | None = None

        library_path = self._app_data_dir / "jingle-library.json"
        self._store = LibraryStore(library_path)

        self._records: list[JingleRecord] = []
        self._visible_indices: list[int] = []
        self._updating_table = False
        self._is_rescanning = False

        self._player: QMediaPlayer | None = None
        self._audio_output: QAudioOutput | None = None
        self._is_muted = False
        self._slider_pressed = False
        self._playback_mode = "off"
        self._continuous_queue: list[int] = []
        self._continuous_queue_position = -1
        self._current_playing_name = ""
        self._is_preview_mode = False
        self._loop_breath_effect: QGraphicsOpacityEffect | None = None
        self._loop_breath_anim: QPropertyAnimation | None = None
        self._play_stop_breath_effect: QGraphicsOpacityEffect | None = None
        self._play_stop_breath_anim: QPropertyAnimation | None = None
        self._stop_btn_breath_effect: QGraphicsOpacityEffect | None = None
        self._stop_btn_breath_anim: QPropertyAnimation | None = None
        self._mode_live_breath_effect: QGraphicsOpacityEffect | None = None
        self._mode_live_breath_anim: QPropertyAnimation | None = None

        if _has_qt_multimedia:
            self._player = QMediaPlayer(self)
            self._audio_output = QAudioOutput(self)
            self._player.setAudioOutput(self._audio_output)
            self._apply_output_device()

        central = QWidget(self)
        self._central_widget = central
        app = QApplication.instance()
        if app is not None:
            app.installEventFilter(self)
        self.setCentralWidget(central)
        root = QVBoxLayout(central)
        root.setContentsMargins(8, 8, 8, 8)
        root.setSpacing(8)

        filter_row = QHBoxLayout()
        self._search_edit = QLineEdit()
        self._search_edit.setPlaceholderText("Search by jingle name, tags, or path")
        self._search_edit.textChanged.connect(self._apply_filters)
        filter_row.addWidget(self._search_edit, 2)

        self._search_scope_combo = QComboBox()
        self._search_scope_combo.addItem("Name + Path + Tag", "all")
        self._search_scope_combo.addItem("Name Only", "name")
        self._search_scope_combo.addItem("Tag Only", "tag")
        self._search_scope_combo.addItem("Path Only", "path")
        self._search_scope_combo.currentIndexChanged.connect(self._on_search_scope_changed)
        self._search_scope_combo.setToolTip("Choose what fields the search box matches.")
        filter_row.addWidget(self._search_scope_combo, 0)

        self._category_filter_edit = QLineEdit()
        self._category_filter_edit.setPlaceholderText("Filter categories (comma-separated)")
        self._category_filter_edit.textChanged.connect(self._apply_filters)
        filter_row.addWidget(self._category_filter_edit, 1)

        self._category_filter_mode = QComboBox()
        self._category_filter_mode.addItem("Match Any", "any")
        self._category_filter_mode.addItem("Match All", "all")
        self._category_filter_mode.setCurrentIndex(1)
        self._category_filter_mode.currentIndexChanged.connect(self._apply_filters)
        filter_row.addWidget(self._category_filter_mode, 0)

        root.addLayout(filter_row)

        chips_row = QHBoxLayout()
        chips_label = QLabel("Active Filters")
        chips_label.setFixedWidth(100)
        chips_row.addWidget(chips_label)

        self._chips_scroll = QScrollArea()
        self._chips_scroll.setWidgetResizable(True)
        self._chips_scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        self._chips_scroll.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        self._chips_scroll.setFrameShape(QScrollArea.Shape.NoFrame)
        self._chips_scroll.setMinimumHeight(34)
        self._chips_scroll.setMaximumHeight(40)

        self._chips_container = QWidget()
        self._chips_layout = QHBoxLayout(self._chips_container)
        self._chips_layout.setContentsMargins(0, 0, 0, 0)
        self._chips_layout.setSpacing(6)
        self._chips_layout.addStretch()
        self._chips_scroll.setWidget(self._chips_container)

        chips_row.addWidget(self._chips_scroll, 1)

        self._clear_filters_btn = QPushButton("Clear All")
        self._clear_filters_btn.setToolTip("Remove all active category filter tags")
        self._clear_filters_btn.clicked.connect(self._clear_all_filter_tags)
        self._clear_filters_btn.setEnabled(False)
        chips_row.addWidget(self._clear_filters_btn)

        root.addLayout(chips_row)

        bulk_grid = QGridLayout()
        bulk_grid.addWidget(QLabel("Set Categories"), 0, 0)
        self._bulk_category_edit = QLineEdit()
        self._bulk_category_edit.setPlaceholderText("Comma-separated, e.g. Holiday, Radio")
        bulk_grid.addWidget(self._bulk_category_edit, 0, 1)

        self._bulk_mode_combo = QComboBox()
        self._bulk_mode_combo.addItem("Replace tags", "replace")
        self._bulk_mode_combo.addItem("Append tags", "append")
        self._bulk_mode_combo.addItem("Remove tags", "remove")
        self._bulk_mode_combo.setCurrentIndex(1)
        self._bulk_mode_combo.setToolTip("Choose how bulk tags are applied to selected rows.")
        bulk_grid.addWidget(self._bulk_mode_combo, 0, 2)

        self._update_from_folders_selected_btn = QPushButton("From Folders (Selected)")
        self._update_from_folders_selected_btn.setToolTip(
            "Update selected rows from folder titles using preserve/overwrite mode."
        )
        self._update_from_folders_selected_btn.clicked.connect(
            self._on_update_selected_from_folders_clicked
        )
        bulk_grid.addWidget(self._update_from_folders_selected_btn, 0, 3)

        self._apply_selected_btn = QPushButton("Apply To Selected")
        self._apply_selected_btn.clicked.connect(self._on_apply_bulk_to_selected)
        bulk_grid.addWidget(self._apply_selected_btn, 0, 4)

        root.addLayout(bulk_grid)

        self._table = DeselectableTableWidget(0, 3)
        self._table.set_preserve_selection_callback(self._should_preserve_selected_row)
        self._table.setHorizontalHeaderLabels(["Jingle", "Categories", "Folder"])
        self._table.itemChanged.connect(self._on_table_item_changed)
        self._table.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
        self._table.setSelectionMode(QTableWidget.SelectionMode.ExtendedSelection)
        self._table.setAlternatingRowColors(True)
        self._table.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self._table.customContextMenuRequested.connect(self._on_table_context_menu_requested)
        header = self._table.horizontalHeader()
        if header is not None:
            header.setStretchLastSection(False)
            header.setSectionResizeMode(0, QHeaderView.ResizeMode.Stretch)
            header.setSectionResizeMode(1, QHeaderView.ResizeMode.Fixed)
            header.setSectionResizeMode(2, QHeaderView.ResizeMode.Fixed)

        self._table.setColumnWidth(1, 120)
        self._table.setColumnWidth(2, 160)
        self._table.itemSelectionChanged.connect(self._on_table_selection_changed)
        self._table.itemDoubleClicked.connect(self._on_table_item_double_clicked)
        root.addWidget(self._table, 1)

        playback_row = QHBoxLayout()
        self._play_btn = QPushButton("Play Selected")
        self._play_btn.setFocusPolicy(Qt.FocusPolicy.NoFocus)
        self._play_btn.clicked.connect(self._on_play_clicked)
        playback_row.addWidget(self._play_btn)
        self._set_play_button_state(False)

        self._stop_btn = QPushButton("Stop")
        self._stop_btn.setFocusPolicy(Qt.FocusPolicy.NoFocus)
        self._stop_btn.setStyleSheet(
            "QPushButton { background-color: #c62828; color: white; font-weight: bold; }"
            "QPushButton:hover { background-color: #d32f2f; }"
        )
        self._stop_btn.clicked.connect(self._on_stop_clicked)
        playback_row.addWidget(self._stop_btn)

        self._mute_btn = QPushButton("Mute")
        self._mute_btn.setFocusPolicy(Qt.FocusPolicy.NoFocus)
        self._mute_btn.clicked.connect(self._on_mute_clicked)
        playback_row.addWidget(self._mute_btn)
        self._refresh_mute_button_state()

        self._loop_btn = QPushButton("Loop Off")
        self._loop_btn.setFocusPolicy(Qt.FocusPolicy.NoFocus)
        self._loop_btn.clicked.connect(self._on_loop_clicked)
        playback_row.addWidget(self._loop_btn)
        self._refresh_playback_mode_button()

        self._mode_btn = QPushButton("Mode: Live")
        self._mode_btn.setFocusPolicy(Qt.FocusPolicy.NoFocus)
        self._mode_btn.setCheckable(True)
        self._mode_btn.toggled.connect(self._on_mode_toggled)
        playback_row.addWidget(self._mode_btn)
        self._refresh_mode_toggle_state(notify_if_disabled=False)

        self._position_slider = QSlider(Qt.Orientation.Horizontal)
        self._position_slider.setFocusPolicy(Qt.FocusPolicy.NoFocus)
        self._position_slider.setRange(0, 0)
        self._position_slider.sliderPressed.connect(self._on_slider_pressed)
        self._position_slider.sliderReleased.connect(self._on_slider_released)
        playback_row.addWidget(self._position_slider, 1)

        self._time_label = QLabel("00:00 / 00:00")
        self._time_label.setFixedWidth(110)
        playback_row.addWidget(self._time_label)

        self._volume_mode_label = QLabel("Live Vol")
        self._volume_mode_label.setFixedWidth(70)
        playback_row.addWidget(self._volume_mode_label)

        self._volume_slider = QSlider(Qt.Orientation.Horizontal)
        self._volume_slider.setFocusPolicy(Qt.FocusPolicy.NoFocus)
        self._volume_slider.setRange(0, 100)
        self._volume_slider.setPageStep(5)
        self._volume_slider.setFixedWidth(140)
        self._volume_slider.valueChanged.connect(self._on_volume_slider_changed)
        playback_row.addWidget(self._volume_slider)

        self._volume_value_label = QLabel("100%")
        self._volume_value_label.setFixedWidth(44)
        playback_row.addWidget(self._volume_value_label)

        self._refresh_volume_controls()

        root.addLayout(playback_row)

        self._status = QStatusBar(self)
        self.setStatusBar(self._status)

        self._library_watcher = QFileSystemWatcher(self)
        self._library_watcher.directoryChanged.connect(self._on_library_watch_path_changed)
        self._library_watcher.fileChanged.connect(self._on_library_watch_path_changed)
        self._watch_rescan_timer = QTimer(self)
        self._watch_rescan_timer.setSingleShot(True)
        self._watch_rescan_timer.timeout.connect(self._on_library_watch_rescan_timeout)

        self._build_menu()
        self._connect_player_signals()
        self._on_search_scope_changed()
        self._refresh_filter_chips([])

        # Defer the initial scan so the window can render immediately.
        QTimer.singleShot(0, self._maybe_run_first_time_setup)

    def eventFilter(self, watched: QObject | None, event: QEvent | None) -> bool:
        if (
            isinstance(watched, QWidget)
            and event is not None
            and event.type() == QEvent.Type.MouseButtonPress
            and isinstance(event, QMouseEvent)
            and event.button() == Qt.MouseButton.LeftButton
        ):
            widget = watched
            preserve_selection_widgets = (
                self._table,
                self._play_btn,
                self._stop_btn,
                self._mute_btn,
                self._loop_btn,
                self._mode_btn,
                self._bulk_category_edit,
                self._bulk_mode_combo,
                self._apply_selected_btn,
                self._update_from_folders_selected_btn,
            )
            preserve_selection = any(
                widget is candidate or candidate.isAncestorOf(widget)
                for candidate in preserve_selection_widgets
            )
            if widget.inherits("QMenuBar") or widget.inherits("QMenu"):
                preserve_selection = True
            if widget.window() is self and not preserve_selection:
                if self._should_preserve_selected_row():
                    return super().eventFilter(watched, event)
                self._table.clearSelection()
                self._table.setCurrentCell(-1, -1)
                self._refresh_status_summary()
        if event is not None and event.type() == QEvent.Type.KeyPress and isinstance(event, QKeyEvent):
            if self._handle_media_key_event(event):
                return True
        return super().eventFilter(watched, event)

    def _should_preserve_selected_row(self) -> bool:
        table = getattr(self, "_table", None)
        return (
            self._player is not None
            and self._player.playbackState() == QMediaPlayer.PlaybackState.PlayingState
            and table is not None
            and bool(table.selectedItems())
        )

    def _is_playback_active(self) -> bool:
        return (
            self._player is not None
            and self._player.playbackState()
            in (
                QMediaPlayer.PlaybackState.PlayingState,
                QMediaPlayer.PlaybackState.PausedState,
            )
        )

    def _handle_media_key_event(self, event: QKeyEvent) -> bool:
        if self._player is None:
            return False

        key = event.key()
        if key in MEDIA_TOGGLE_PLAYBACK_KEYS:
            self._toggle_play_pause()
            event.accept()
            return True
        if key in MEDIA_PLAY_KEYS:
            self._resume_or_start_playback()
            event.accept()
            return True
        if key in MEDIA_PAUSE_KEYS:
            self._pause_playback()
            event.accept()
            return True
        if key in MEDIA_NEXT_KEYS and self._is_playback_active():
            self._skip_to_next()
            event.accept()
            return True
        if key in MEDIA_PREVIOUS_KEYS and self._is_playback_active():
            self._skip_to_previous()
            event.accept()
            return True
        return False

    def _toggle_play_pause(self) -> None:
        if self._player is None:
            return
        state = self._player.playbackState()
        if state == QMediaPlayer.PlaybackState.PlayingState:
            self._player.pause()
            self._status.showMessage("Playback paused.")
            return
        if state == QMediaPlayer.PlaybackState.PausedState:
            self._player.play()
            self._status.showMessage("Playback resumed.")
            return
        self._on_play_clicked()

    def _resume_or_start_playback(self) -> None:
        if self._player is None:
            return
        if self._player.playbackState() == QMediaPlayer.PlaybackState.PausedState:
            self._player.play()
            self._status.showMessage("Playback resumed.")
            return
        if self._player.playbackState() != QMediaPlayer.PlaybackState.PlayingState:
            self._on_play_clicked()

    def _pause_playback(self) -> None:
        if self._player is None:
            return
        if self._player.playbackState() == QMediaPlayer.PlaybackState.PlayingState:
            self._player.pause()
            self._status.showMessage("Playback paused.")

    def _refresh_mute_button_state(self) -> None:
        if self._is_muted:
            self._mute_btn.setText("Unmute")
            self._mute_btn.setStyleSheet(
                "QPushButton { background-color: #546e7a; color: white; font-weight: bold; }"
                "QPushButton:hover { background-color: #607d8b; }"
            )
        else:
            self._mute_btn.setText("Mute")
            self._mute_btn.setStyleSheet(
                "QPushButton { background-color: #455a64; color: white; font-weight: bold; }"
                "QPushButton:hover { background-color: #546e7a; }"
            )

    def _set_muted(self, muted: bool) -> None:
        self._is_muted = muted
        if self._audio_output is not None:
            self._audio_output.setMuted(muted)
        self._refresh_mute_button_state()

    def _on_mute_clicked(self) -> None:
        if self._audio_output is None:
            self._status.showMessage("Playback unavailable: PyQt6 multimedia is not installed.")
            return
        self._set_muted(not self._is_muted)
        self._status.showMessage("Audio muted." if self._is_muted else "Audio unmuted.")

    def _active_volume_percent(self) -> int:
        if self._is_preview_mode and self._can_use_preview_mode():
            return self._preview_volume_percent
        return self._live_volume_percent

    def _save_volume_settings(self) -> None:
        self._settings.setValue("options/liveVolumePercent", self._live_volume_percent)
        self._settings.setValue("options/previewVolumePercent", self._preview_volume_percent)

    def _apply_active_volume(self) -> None:
        if self._audio_output is None:
            return
        self._audio_output.setVolume(self._active_volume_percent() / 100.0)

    def _refresh_volume_controls(self) -> None:
        if not hasattr(self, "_volume_slider"):
            return
        mode_text = "Preview Vol" if self._is_preview_mode and self._can_use_preview_mode() else "Live Vol"
        self._volume_mode_label.setText(mode_text)
        value = self._active_volume_percent()
        self._volume_slider.blockSignals(True)
        self._volume_slider.setValue(value)
        self._volume_slider.blockSignals(False)
        self._volume_value_label.setText(f"{value}%")
        self._volume_slider.setToolTip(
            "Adjust preview volume." if mode_text == "Preview Vol" else "Adjust live volume."
        )

    def _on_volume_slider_changed(self, value: int) -> None:
        percent = _coerce_volume_percent(value)
        if self._is_preview_mode and self._can_use_preview_mode():
            self._preview_volume_percent = percent
        else:
            self._live_volume_percent = percent
        self._volume_value_label.setText(f"{percent}%")
        self._save_volume_settings()
        self._apply_active_volume()

    def keyPressEvent(self, event: QKeyEvent | None) -> None:
        if event is None:
            super().keyPressEvent(event)
            return

        # Check if playback is currently active (playing or paused)
        if self._player is None:
            super().keyPressEvent(event)
            return

        is_playback_active = self._is_playback_active()

        # Handle arrow keys for skipping during active playback
        if is_playback_active:
            if self._event_matches_shortcut(event, "skip_previous"):
                self._skip_to_previous()
                event.accept()
                return
            elif self._event_matches_shortcut(event, "skip_next"):
                self._skip_to_next()
                event.accept()
                return
        else:
            # Handle arrow keys for table navigation when playback is not active
            if self._event_matches_shortcut(event, "select_up"):
                self._move_selection_up()
                event.accept()
                return
            elif self._event_matches_shortcut(event, "select_down"):
                self._move_selection_down()
                event.accept()
                return

        super().keyPressEvent(event)

    def _skip_to_previous(self) -> None:
        """Skip to the previous jingle in the visible list."""
        if not self._visible_indices:
            self._status.showMessage("No jingles available to skip.")
            return

        current_index = self._selected_record_index()
        if current_index is None:
            # No selection, skip to last visible jingle
            next_visible_index = self._visible_indices[-1]
        else:
            # Find current in visible list and go to previous
            try:
                visible_pos = self._visible_indices.index(current_index)
                if visible_pos > 0:
                    next_visible_index = self._visible_indices[visible_pos - 1]
                else:
                    # At the beginning, wrap to the end
                    next_visible_index = self._visible_indices[-1]
            except ValueError:
                # Current index not in visible list, go to last visible
                next_visible_index = self._visible_indices[-1]

        record = self._records[next_visible_index]
        if not record.path.exists():
            self._status.showMessage("Selected file no longer exists.")
            return

        # Select the row in the table
        visible_row = self._visible_row_for_record_index(next_visible_index)
        if visible_row >= 0:
            self._table.selectRow(visible_row)

        # In continuous mode, continue from the manually selected jingle.
        if self._playback_mode == "continuous":
            try:
                self._continuous_queue_position = self._continuous_queue.index(next_visible_index)
            except ValueError:
                pass

        self._play_record(next_visible_index)
        self._status.showMessage(f"Skipped to: {record.name}")

    def _skip_to_next(self) -> None:
        """Skip to the next jingle in the visible list."""
        if not self._visible_indices:
            self._status.showMessage("No jingles available to skip.")
            return

        current_index = self._selected_record_index()
        if current_index is None:
            # No selection, skip to first visible jingle
            next_visible_index = self._visible_indices[0]
        else:
            # Find current in visible list and go to next
            try:
                visible_pos = self._visible_indices.index(current_index)
                if visible_pos < len(self._visible_indices) - 1:
                    next_visible_index = self._visible_indices[visible_pos + 1]
                else:
                    # At the end, wrap to the beginning
                    next_visible_index = self._visible_indices[0]
            except ValueError:
                # Current index not in visible list, go to first visible
                next_visible_index = self._visible_indices[0]

        record = self._records[next_visible_index]
        if not record.path.exists():
            self._status.showMessage("Selected file no longer exists.")
            return

        # Select the row in the table
        visible_row = self._visible_row_for_record_index(next_visible_index)
        if visible_row >= 0:
            self._table.selectRow(visible_row)

        # In continuous mode, continue from the manually selected jingle.
        if self._playback_mode == "continuous":
            try:
                self._continuous_queue_position = self._continuous_queue.index(next_visible_index)
            except ValueError:
                pass

        self._play_record(next_visible_index)
        self._status.showMessage(f"Skipped to: {record.name}")

    def _move_selection_up(self) -> None:
        """Move the selected row up by one (when playback is not active)."""
        selected_ranges = self._table.selectedRanges()
        if not selected_ranges:
            # No selection, select the first row
            if self._visible_indices:
                self._table.selectRow(0)
            return

        current_row = selected_ranges[0].topRow()
        if current_row > 0:
            # Move up
            self._table.selectRow(current_row - 1)
        else:
            # At the top, wrap to the bottom
            if self._visible_indices:
                self._table.selectRow(len(self._visible_indices) - 1)

    def _move_selection_down(self) -> None:
        """Move the selected row down by one (when playback is not active)."""
        selected_ranges = self._table.selectedRanges()
        if not selected_ranges:
            # No selection, select the first row
            if self._visible_indices:
                self._table.selectRow(0)
            return

        current_row = selected_ranges[0].topRow()
        max_row = len(self._visible_indices) - 1
        if current_row < max_row:
            # Move down
            self._table.selectRow(current_row + 1)
        else:
            # At the bottom, wrap to the top
            self._table.selectRow(0)

    def _build_menu(self) -> None:
        menu_bar = self.menuBar()
        file_menu = menu_bar.addMenu("File")

        export_db_action = QAction("Export Tag Database...", self)
        export_db_action.triggered.connect(self._on_file_export_tag_database)
        file_menu.addAction(export_db_action)

        import_db_action = QAction("Import Tag Database...", self)
        import_db_action.triggered.connect(self._on_file_import_tag_database)
        file_menu.addAction(import_db_action)

        file_menu.addSeparator()

        rescan_action = QAction("Rescan", self)
        rescan_action.triggered.connect(self._rescan_library)
        file_menu.addAction(rescan_action)

        browse_action = QAction("Choose Samples Folder...", self)
        browse_action.triggered.connect(self._on_browse_folder)
        file_menu.addAction(browse_action)

        file_menu.addSeparator()

        exit_action = QAction("Exit", self)
        exit_action.triggered.connect(self.close)
        file_menu.addAction(exit_action)

        edit_menu = menu_bar.addMenu("Edit")

        rename_action = QAction("Rename", self)
        rename_action.triggered.connect(self._on_edit_rename)
        edit_menu.addAction(rename_action)
        self._rename_action = rename_action

        delete_action = QAction("Delete", self)
        delete_action.triggered.connect(self._on_edit_delete)
        edit_menu.addAction(delete_action)
        self._delete_action = delete_action

        edit_menu.addSeparator()

        copy_to_action = QAction("Copy-To", self)
        copy_to_action.triggered.connect(self._on_edit_copy_to)
        edit_menu.addAction(copy_to_action)

        move_to_action = QAction("Move-To", self)
        move_to_action.triggered.connect(self._on_edit_move_to)
        edit_menu.addAction(move_to_action)

        tools_menu = menu_bar.addMenu("Tools")
        options_action = QAction("Options", self)
        options_action.triggered.connect(self._on_open_options)
        tools_menu.addAction(options_action)
        edit_shortcuts_action = QAction("Edit Keyboard Shortcuts...", self)
        edit_shortcuts_action.triggered.connect(self._on_edit_keyboard_shortcuts)
        tools_menu.addAction(edit_shortcuts_action)
        tools_menu.addSeparator()
        update_from_folders_action = QAction("Update Categories from Folder Titles", self)
        update_from_folders_action.triggered.connect(self._on_tools_update_categories_from_folders)
        tools_menu.addAction(update_from_folders_action)
        find_duplicates_action = QAction("Find Duplicates", self)
        find_duplicates_action.triggered.connect(self._on_tools_find_duplicates)
        tools_menu.addAction(find_duplicates_action)
        clear_all_categories_action = QAction("Clear All Categories", self)
        clear_all_categories_action.triggered.connect(self._on_tools_clear_all_categories)
        tools_menu.addAction(clear_all_categories_action)
        tools_menu.addSeparator()
        self._auto_folder_tags_action = QAction("Auto-tag from Folders on Scan", self)
        self._auto_folder_tags_action.setCheckable(True)
        self._auto_folder_tags_action.setChecked(self._auto_folder_tags)
        self._auto_folder_tags_action.toggled.connect(self._on_auto_folder_tags_toggled)
        tools_menu.addAction(self._auto_folder_tags_action)

        self._watch_library_changes_action = QAction("Auto-Refresh on Library Changes", self)
        self._watch_library_changes_action.setCheckable(True)
        self._watch_library_changes_action.setChecked(self._watch_library_changes)
        self._watch_library_changes_action.toggled.connect(self._on_watch_library_changes_toggled)
        tools_menu.addAction(self._watch_library_changes_action)

        help_menu = menu_bar.addMenu("Help")
        about_action = QAction("About", self)
        about_action.triggered.connect(self._on_help_about)
        help_menu.addAction(about_action)

        self._apply_keyboard_shortcuts_to_actions()

    def _on_help_about(self) -> None:
        library_count = len(self._records)
        library_duration_seconds = sum(max(0.0, record.duration_seconds) for record in self._records)
        library_size_bytes = sum(max(0, record.size_bytes) for record in self._records)
        runtime_dir = _runtime_app_dir()
        revision_log_path = runtime_dir / "rev.log"
        resolved_revision_log = revision_log_path if revision_log_path.is_file() else None
        dialog = AboutDialog(
            library_count=library_count,
            library_duration_seconds=library_duration_seconds,
            library_size_bytes=library_size_bytes,
            revision_log_path=resolved_revision_log,
            parent=self,
        )
        dialog.exec()

    def _load_keyboard_shortcuts(self) -> dict[str, str]:
        shortcuts: dict[str, str] = dict(DEFAULT_KEYBOARD_SHORTCUTS)
        for key, default in DEFAULT_KEYBOARD_SHORTCUTS.items():
            value = str(self._settings.value(f"shortcuts/{key}", default)).strip()
            if value:
                shortcuts[key] = value
            else:
                shortcuts[key] = ""
        return shortcuts

    def _save_keyboard_shortcuts(self) -> None:
        for key, default in DEFAULT_KEYBOARD_SHORTCUTS.items():
            value = self._keyboard_shortcuts.get(key, default).strip()
            self._settings.setValue(f"shortcuts/{key}", value)

    def _shortcut_for(self, key: str) -> QKeySequence:
        value = self._keyboard_shortcuts.get(key, "")
        return QKeySequence(value)

    def _event_matches_shortcut(self, event: QKeyEvent, key: str) -> bool:
        seq = self._shortcut_for(key)
        if seq.isEmpty():
            return False
        event_seq = QKeySequence(event.keyCombination())
        return (
            event_seq.toString(QKeySequence.SequenceFormat.PortableText)
            == seq.toString(QKeySequence.SequenceFormat.PortableText)
        )

    def _apply_keyboard_shortcuts_to_actions(self) -> None:
        if self._rename_action is not None:
            self._rename_action.setShortcut(self._shortcut_for("rename"))
        if self._delete_action is not None:
            self._delete_action.setShortcut(self._shortcut_for("delete"))

    def _on_edit_keyboard_shortcuts(self) -> None:
        dialog = KeyboardShortcutsDialog(
            current_shortcuts=self._keyboard_shortcuts,
            default_shortcuts=DEFAULT_KEYBOARD_SHORTCUTS,
            parent=self,
        )
        if dialog.exec() != QDialog.DialogCode.Accepted:
            return

        selected = dialog.selected_shortcuts()
        self._keyboard_shortcuts = {
            key: selected.get(key, DEFAULT_KEYBOARD_SHORTCUTS[key]).strip()
            for key in DEFAULT_KEYBOARD_SHORTCUTS
        }
        self._save_keyboard_shortcuts()
        self._apply_keyboard_shortcuts_to_actions()
        self._status.showMessage("Keyboard shortcuts updated.")

    def _on_file_export_tag_database(self) -> None:
        default_name = "jingle-tags-backup.json"
        start_dir = str(self._samples_dir) if self._samples_dir else str(Path.home())
        start_path = str(Path(start_dir) / default_name)
        target_str, _ = QFileDialog.getSaveFileName(
            self,
            "Export Tag Database",
            start_path,
            "JSON Files (*.json);;All Files (*)",
        )
        if not target_str:
            return

        target_path = Path(target_str)
        if not target_path.suffix:
            target_path = target_path.with_suffix(".json")

        try:
            self._store.export_to(target_path)
        except OSError as exc:
            QMessageBox.critical(
                self,
                "Export Failed",
                f"Could not export tag database.\n\n{exc}",
            )
            self._status.showMessage("Tag database export failed.")
            return

        self._status.showMessage(f"Exported tag database to {target_path.name}")

    def _on_file_import_tag_database(self) -> None:
        source_str, _ = QFileDialog.getOpenFileName(
            self,
            "Import Tag Database",
            str(self._samples_dir) if self._samples_dir else str(Path.home()),
            "JSON Files (*.json);;All Files (*)",
        )
        if not source_str:
            return

        reply = QMessageBox.question(
            self,
            "Import Tag Database",
            "Importing will replace the current in-memory tag database and apply it to loaded jingles.\n\n"
            "Continue?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No,
        )
        if reply != QMessageBox.StandardButton.Yes:
            self._status.showMessage("Tag database import cancelled.")
            return

        source_path = Path(source_str)
        try:
            self._store.import_from(source_path)
        except (OSError, json.JSONDecodeError, UnicodeDecodeError) as exc:
            QMessageBox.critical(
                self,
                "Import Failed",
                f"Could not import tag database.\n\n{exc}",
            )
            self._status.showMessage("Tag database import failed.")
            return

        for record in self._records:
            record.categories = self._store.get(record.path)

        self._store.save()
        self._apply_filters()
        self._status.showMessage(f"Imported tag database from {source_path.name}")

    def _selected_single_record_index(self) -> int | None:
        selected_indices = self._selected_record_indices()
        if not selected_indices:
            self._status.showMessage("Select one jingle first.")
            return None
        if len(selected_indices) != 1:
            self._status.showMessage("Select exactly one jingle for this action.")
            return None
        return selected_indices[0]

    def _validate_new_filename(self, value: str) -> str | None:
        name = value.strip()
        if not name:
            return "Filename cannot be empty."

        forbidden_chars = '<>:"/\\|?*'
        if any(ch in forbidden_chars for ch in name):
            return "Filename contains illegal characters: < > : \" / \\ | ? *"

        if any(ord(ch) < 32 for ch in name):
            return "Filename contains illegal control characters."

        if name.endswith((" ", ".")):
            return "Filename cannot end with a space or period."

        stem = Path(name).stem.strip().upper()
        if stem in {
            "CON",
            "PRN",
            "AUX",
            "NUL",
            "COM1",
            "COM2",
            "COM3",
            "COM4",
            "COM5",
            "COM6",
            "COM7",
            "COM8",
            "COM9",
            "LPT1",
            "LPT2",
            "LPT3",
            "LPT4",
            "LPT5",
            "LPT6",
            "LPT7",
            "LPT8",
            "LPT9",
        }:
            return "Filename uses a reserved Windows device name."

        return None

    def _on_edit_rename(self) -> None:
        record_index = self._selected_single_record_index()
        if record_index is None:
            return

        record = self._records[record_index]
        source_path = record.path
        current_name = source_path.name

        while True:
            new_name, accepted = QInputDialog.getText(
                self,
                "Rename Jingle",
                "New filename:",
                text=current_name,
            )
            if not accepted:
                self._status.showMessage("Rename cancelled.")
                return

            candidate = new_name.strip()
            if not Path(candidate).suffix:
                candidate = f"{candidate}{source_path.suffix}"

            validation_error = self._validate_new_filename(candidate)
            if validation_error is not None:
                QMessageBox.warning(self, "Invalid Filename", validation_error)
                current_name = candidate
                continue

            destination_path = source_path.with_name(candidate)
            if destination_path == source_path:
                self._status.showMessage("Rename cancelled: filename is unchanged.")
                return

            if destination_path.exists():
                QMessageBox.warning(
                    self,
                    "Filename Exists",
                    "A file with that name already exists. Choose a different filename.",
                )
                current_name = destination_path.name
                continue

            try:
                source_path.rename(destination_path)
            except OSError as exc:
                QMessageBox.critical(self, "Rename Failed", f"Could not rename file.\n\n{exc}")
                self._status.showMessage("Rename failed.")
                return

            record.path = destination_path
            self._store.rename(source_path, destination_path)
            self._store.save()
            self._apply_filters()
            self._status.showMessage(f"Renamed to {destination_path.name}")
            return

    def _on_edit_delete(self) -> None:
        selected_indices = self._selected_record_indices()
        if not selected_indices:
            self._status.showMessage("Select one or more jingles first.")
            return

        count = len(selected_indices)
        reply = QMessageBox.question(
            self,
            "Delete Jingle(s)",
            f"Delete {count} selected file(s)? This cannot be undone.",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No,
        )
        if reply != QMessageBox.StandardButton.Yes:
            self._status.showMessage("Delete cancelled.")
            return

        removed_count = 0
        failed_paths: list[Path] = []
        for record_index in selected_indices:
            if record_index < 0 or record_index >= len(self._records):
                continue
            path = self._records[record_index].path
            try:
                path.unlink()
            except OSError:
                failed_paths.append(path)
                continue

            self._store.remove(path)
            removed_count += 1

        self._store.save()
        self._rescan_library()

        if failed_paths:
            QMessageBox.warning(
                self,
                "Delete Incomplete",
                f"Deleted {removed_count} file(s), but {len(failed_paths)} could not be deleted.",
            )
            self._status.showMessage(
                f"Deleted {removed_count} file(s); {len(failed_paths)} failed."
            )
            return

        self._status.showMessage(f"Deleted {removed_count} file(s).")

    def _copy_or_move_selected_jingle(self, move: bool) -> None:
        record_index = self._selected_single_record_index()
        if record_index is None:
            return

        record = self._records[record_index]
        source_path = record.path
        source_categories = list(record.categories)
        source_ext = source_path.suffix.lower()
        ext_label = source_ext[1:].upper() if source_ext.startswith(".") and len(source_ext) > 1 else "Source"
        ext_pattern = f"*{source_ext}" if source_ext else "*"
        file_filter = f"{ext_label} Files ({ext_pattern});;All Files (*)"

        operation_name = "Move" if move else "Copy"
        target_str, _ = QFileDialog.getSaveFileName(
            self,
            f"{operation_name} Jingle To",
            str(source_path),
            file_filter,
        )
        if not target_str:
            self._status.showMessage(f"{operation_name} cancelled.")
            return

        destination_path = Path(target_str)
        if destination_path.suffix == "":
            destination_path = destination_path.with_suffix(source_path.suffix)

        if destination_path == source_path:
            self._status.showMessage(f"{operation_name} cancelled: source and destination are the same.")
            return

        if destination_path.exists():
            overwrite = QMessageBox.question(
                self,
                f"{operation_name} Jingle",
                f"{destination_path.name} already exists. Overwrite it?",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                QMessageBox.StandardButton.No,
            )
            if overwrite != QMessageBox.StandardButton.Yes:
                self._status.showMessage(f"{operation_name} cancelled.")
                return

        try:
            destination_path.parent.mkdir(parents=True, exist_ok=True)
            shutil.copy2(source_path, destination_path)
            try:
                dest_stat = destination_path.stat()
                self._store.set_media_cache(
                    destination_path,
                    int(dest_stat.st_size),
                    record.duration_seconds,
                    int(dest_stat.st_mtime_ns),
                )
            except OSError:
                pass
            self._store.set(destination_path, source_categories)
            if move:
                source_path.unlink()
        except OSError as exc:
            QMessageBox.critical(self, f"{operation_name} Failed", f"Could not {operation_name.lower()} file.\n\n{exc}")
            self._status.showMessage(f"{operation_name} failed.")
            return

        if move:
            self._store.remove(source_path)
        self._store.save()

        self._rescan_library()
        self._status.showMessage(f"{operation_name} complete: {destination_path.name}")

    def _on_edit_copy_to(self) -> None:
        self._copy_or_move_selected_jingle(move=False)

    def _on_edit_move_to(self) -> None:
        self._copy_or_move_selected_jingle(move=True)

    def _on_tools_clear_all_categories(self) -> None:
        if not self._records:
            self._status.showMessage("No jingles loaded.")
            return

        reply = QMessageBox.question(
            self,
            "Clear All Categories",
            "This will remove all category tags from every loaded jingle.\n\n"
            "Are you sure you want to continue?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No,
        )
        if reply != QMessageBox.StandardButton.Yes:
            self._status.showMessage("Clear all categories cancelled.")
            return

        updated = 0
        for record in self._records:
            if record.categories:
                record.categories = []
                self._store.set(record.path, [])
                updated += 1

        self._store.save()
        self._apply_filters()
        self._status.showMessage(f"Cleared categories for {updated} jingle(s).")

    def _on_auto_folder_tags_toggled(self, checked: bool) -> None:
        self._auto_folder_tags = checked
        self._save_auto_folder_tags()
        state = "enabled" if checked else "disabled"
        self._status.showMessage(f"Auto-tag from folders on scan: {state}.")

    def _on_watch_library_changes_toggled(self, checked: bool) -> None:
        self._watch_library_changes = checked
        self._save_watch_library_changes()

        if not checked:
            self._watch_rescan_timer.stop()
            self._library_watcher.removePaths(self._library_watcher.directories())
            self._library_watcher.removePaths(self._library_watcher.files())
            self._status.showMessage("Auto-refresh on library changes: disabled.")
            return

        self._refresh_library_watcher_paths([record.path for record in self._records])
        self._status.showMessage("Auto-refresh on library changes: enabled.")

    def _on_tools_update_categories_from_folders(self) -> None:
        if not self._records:
            self._status.showMessage("No jingles loaded.")
            return

        preserve_existing = self._prompt_folder_update_mode()
        if preserve_existing is None:
            self._status.showMessage("Folder-derived category update cancelled.")
            return

        updated = self._apply_folder_titles_to_records(self._records, preserve_existing)
        mode_text = "preserved" if preserve_existing else "overwritten"
        self._status.showMessage(
            f"Updated categories from folder titles for {updated} jingle(s); existing tags {mode_text}."
        )

    def _duplicate_format_sort_key(self, path: Path) -> tuple[int, str, str]:
        suffix = path.suffix.lower()
        return (
            DUPLICATE_FORMAT_PRIORITY.get(suffix, 99),
            suffix,
            path.name.casefold(),
        )

    def _find_duplicate_audio_variants(self) -> list[list[Path]]:
        if self._samples_dir is None:
            return []

        grouped: dict[tuple[str, str], list[Path]] = {}
        for path in self._scan_audio_files(self._samples_dir):
            key = (str(path.parent).casefold(), path.stem.casefold())
            grouped.setdefault(key, []).append(path)

        duplicates: list[list[Path]] = []
        for variants in grouped.values():
            suffixes = {path.suffix.lower() for path in variants}
            if len(variants) < 2 or len(suffixes) < 2:
                continue

            ordered = sorted(variants, key=self._duplicate_format_sort_key)
            duplicates.append(ordered)

        return duplicates

    def _prompt_duplicate_resolution_mode(self, duplicate_count: int, removal_count: int, sample_paths: list[Path]) -> str | None:
        sample_keep = sample_paths[0]
        sample_remove = sample_paths[1:]
        sample_text = (
            f"\n\nExample:\nKeep: {sample_keep.name}\nRemove: "
            + ", ".join(path.name for path in sample_remove[:3])
        )

        box = QMessageBox(self)
        box.setWindowTitle("Find Duplicates")
        box.setIcon(QMessageBox.Icon.Question)
        box.setText(
            f"Found {duplicate_count} duplicate jingle name group(s) across multiple audio formats.\n\n"
            f"Automatic mode will remove {removal_count} lower-priority file(s), preferring WAV when available."
            f"{sample_text}\n\nHow would you like to resolve duplicates?"
        )
        auto_button = box.addButton("Automatic", QMessageBox.ButtonRole.AcceptRole)
        manual_button = box.addButton("Ask Me Each Time", QMessageBox.ButtonRole.ActionRole)
        cancel_button = box.addButton(QMessageBox.StandardButton.Cancel)
        box.setDefaultButton(auto_button)
        box.exec()

        clicked = box.clickedButton()
        if clicked is auto_button:
            return "auto"
        if clicked is manual_button:
            return "manual"
        if clicked is cancel_button:
            return None
        return None

    def _prompt_duplicate_keep_choice(self, variants: list[Path]) -> Path | None:
        folder_name = variants[0].parent.name or str(variants[0].parent)
        items = [path.name for path in variants]
        choice, accepted = QInputDialog.getItem(
            self,
            "Find Duplicates",
            "Choose which file to keep:\n\n"
            f"Folder: {folder_name}\n"
            f"Jingle: {variants[0].stem}",
            items,
            0,
            False,
        )
        if not accepted:
            return None
        for path in variants:
            if path.name == choice:
                return path
        return None

    def _build_duplicate_removal_plan(
        self, duplicates: list[list[Path]], resolution_mode: str
    ) -> list[tuple[Path, list[Path]]] | None:
        removal_plan: list[tuple[Path, list[Path]]] = []
        for variants in duplicates:
            keep_path = variants[0]
            if resolution_mode == "manual":
                selected_keep = self._prompt_duplicate_keep_choice(variants)
                if selected_keep is None:
                    return None
                keep_path = selected_keep
            remove_paths = [path for path in variants if path != keep_path]
            if remove_paths:
                removal_plan.append((keep_path, remove_paths))
        return removal_plan

    def _on_tools_find_duplicates(self) -> None:
        if self._samples_dir is None:
            self._status.showMessage("Choose a samples folder first.")
            return

        if (
            self._player is not None
            and self._player.playbackState() == QMediaPlayer.PlaybackState.PlayingState
        ):
            self._status.showMessage("Stop playback before running Find Duplicates.")
            return

        duplicates = self._find_duplicate_audio_variants()
        if not duplicates:
            self._status.showMessage("No duplicate audio-format variants were found.")
            return

        automatic_removal_count = sum(max(0, len(variants) - 1) for variants in duplicates)
        resolution_mode = self._prompt_duplicate_resolution_mode(
            len(duplicates),
            automatic_removal_count,
            duplicates[0],
        )
        if resolution_mode is None:
            self._status.showMessage("Find Duplicates cancelled.")
            return

        removal_plan = self._build_duplicate_removal_plan(duplicates, resolution_mode)
        if removal_plan is None:
            self._status.showMessage("Find Duplicates cancelled.")
            return

        removed_count = 0
        failed_paths: list[Path] = []
        for _keep_path, remove_paths in removal_plan:
            for path in remove_paths:
                try:
                    path.unlink()
                    removed_count += 1
                except OSError:
                    failed_paths.append(path)

        if removed_count > 0:
            self._rescan_library()

        if failed_paths:
            failed_count = len(failed_paths)
            QMessageBox.warning(
                self,
                "Find Duplicates",
                f"Removed {removed_count} duplicate file(s), but {failed_count} could not be deleted.",
            )
            self._status.showMessage(
                f"Removed {removed_count} duplicate file(s); {failed_count} failed to delete."
            )
            return

        if removed_count > 0:
            self._status.showMessage(f"Removed {removed_count} duplicate file(s).")
        else:
            self._status.showMessage("No duplicate files were removed.")

    def _on_update_selected_from_folders_clicked(self) -> None:
        selected_rows = sorted({idx.row() for idx in self._table.selectedIndexes()})
        if not selected_rows:
            self._status.showMessage("Select one or more rows first.")
            return

        selected_records: list[JingleRecord] = []
        seen: set[int] = set()
        for row in selected_rows:
            if row < 0 or row >= len(self._visible_indices):
                continue
            record_index = self._visible_indices[row]
            if record_index in seen:
                continue
            seen.add(record_index)
            selected_records.append(self._records[record_index])

        if not selected_records:
            self._status.showMessage("No valid selected rows found.")
            return

        preserve_existing = self._prompt_folder_update_mode()
        if preserve_existing is None:
            self._status.showMessage("Selected folder-derived category update cancelled.")
            return

        updated = self._apply_folder_titles_to_records(selected_records, preserve_existing)
        mode_text = "preserved" if preserve_existing else "overwritten"
        self._status.showMessage(
            f"Updated selected rows from folder titles for {updated} jingle(s); existing tags {mode_text}."
        )

    def _prompt_folder_update_mode(self) -> bool | None:
        """Return True=preserve, False=overwrite, None=cancel."""
        prompt = QMessageBox(self)
        prompt.setWindowTitle("Update Categories from Folder Titles")
        prompt.setIcon(QMessageBox.Icon.Question)
        prompt.setText(
            "How should folder-derived tags be applied?\n\n"
            "Folder tags are derived from the path under your selected Samples folder."
        )
        preserve_btn = prompt.addButton("Preserve Existing Tags", QMessageBox.ButtonRole.AcceptRole)
        overwrite_btn = prompt.addButton("Overwrite Tags", QMessageBox.ButtonRole.DestructiveRole)
        cancel_btn = prompt.addButton("Cancel", QMessageBox.ButtonRole.RejectRole)
        prompt.setDefaultButton(preserve_btn)
        prompt.exec()

        clicked = prompt.clickedButton()
        if clicked is preserve_btn:
            return True
        if clicked is overwrite_btn:
            return False
        if clicked is cancel_btn:
            return None
        return None

    def _derive_folder_tags(self, record: JingleRecord) -> list[str]:
        """Derive folder tags under sample root, excluding the sample root itself."""
        if self._samples_dir is None:
            return []
        try:
            rel_parent = record.path.relative_to(self._samples_dir).parent
            return [part.strip() for part in rel_parent.parts if part.strip()]
        except ValueError:
            # If the file is outside selected root unexpectedly, skip derivation.
            return []

    def _apply_folder_titles_to_records(
        self,
        records: list[JingleRecord],
        preserve_existing: bool,
    ) -> int:
        updated = 0
        for record in records:
            derived_tags = self._derive_folder_tags(record)
            if preserve_existing:
                new_categories = _merge_tags(record.categories, derived_tags)
            else:
                new_categories = _normalize_tags(derived_tags)

            if new_categories != record.categories:
                record.categories = new_categories
                self._store.set(record.path, new_categories)
                updated += 1

        self._store.save()
        self._apply_filters()
        return updated

    def _connect_player_signals(self) -> None:
        if self._player is None:
            self._play_btn.setEnabled(False)
            self._stop_btn.setEnabled(False)
            self._mute_btn.setEnabled(False)
            self._loop_btn.setEnabled(False)
            self._mode_btn.setEnabled(False)
            self._volume_slider.setEnabled(False)
            self._set_loop_breathing(False)
            self._status.showMessage("PyQt6 multimedia is not available. Playback is disabled.")
            return

        self._player.durationChanged.connect(self._on_duration_changed)
        self._player.positionChanged.connect(self._on_position_changed)
        self._player.playbackStateChanged.connect(self._on_playback_state_changed)
        self._player.mediaStatusChanged.connect(self._on_media_status_changed)
        self._apply_player_loop_mode()

    def _create_settings_store(self) -> QSettings:
        settings_path = self._app_data_dir / "settings.ini"
        settings = QSettings(str(settings_path), QSettings.Format.IniFormat)
        settings.setFallbacksEnabled(False)
        return settings

    def _load_samples_dir(self) -> Path | None:
        raw = str(self._settings.value("library/samplesDir", "")).strip()
        if raw:
            candidate = Path(raw)
            if candidate.exists() and candidate.is_dir():
                return candidate
        return None

    def _save_samples_dir(self) -> None:
        if self._samples_dir is not None:
            self._settings.setValue("library/samplesDir", str(self._samples_dir))

    def _load_auto_folder_tags(self) -> bool:
        return str(self._settings.value("library/autoFolderTags", "")).strip().lower() == "true"

    def _save_auto_folder_tags(self) -> None:
        self._settings.setValue("library/autoFolderTags", "true" if self._auto_folder_tags else "false")

    def _load_watch_library_changes(self) -> bool:
        raw = str(self._settings.value("library/watchLibraryChanges", "true")).strip().lower()
        return raw != "false"

    def _save_watch_library_changes(self) -> None:
        self._settings.setValue(
            "library/watchLibraryChanges",
            "true" if self._watch_library_changes else "false",
        )

    def _maybe_run_first_time_setup(self) -> None:
        is_first_run = str(self._settings.value("library/samplesDir", "")).strip() == ""
        if is_first_run:
            self._run_first_time_setup()
        else:
            loaded_count = self._load_cached_records_for_selected_root()
            if loaded_count > 0:
                self._status.showMessage(
                    f"Loaded {loaded_count} jingles from cache. Checking for library changes..."
                )
            QTimer.singleShot(0, self._rescan_library)

    def _run_first_time_setup(self) -> None:
        QMessageBox.information(
            self,
            "Welcome to JingleAllTheDay",
            "Let's get started.\n\nFirst, select the folder where your jingles and audio samples are stored.",
        )

        selected = QFileDialog.getExistingDirectory(
            self,
            "Choose Samples / Jingles Folder",
            str(Path.home()),
        )
        if not selected:
            QMessageBox.information(
                self,
                "Setup Cancelled",
                "No folder was selected. The application will now close.\n\n"
                "You will be prompted again on next launch.",
            )
            QApplication.quit()
            return

        path = Path(selected)
        if not path.exists() or not path.is_dir():
            QMessageBox.warning(self, "Invalid Folder", "The selected path is not a valid folder. The application will now close.")
            QApplication.quit()
            return

        self._samples_dir = path
        self._save_samples_dir()

        reply = QMessageBox.question(
            self,
            "Auto-tag from Folders?",
            "Would you like to automatically assign tags to jingles based on their folder names?\n\n"
            "For example, a file in Samples/Holiday/Christmas/ would receive the tags \"Holiday\" and \"Christmas\".\n\n"
            "You can change this later via Tools > Options.",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.Yes,
        )
        self._auto_folder_tags = reply == QMessageBox.StandardButton.Yes
        self._save_auto_folder_tags()

        self._rescan_library()

    def _scan_audio_files(self, root: Path) -> list[Path]:
        files: list[Path] = []
        if not root.exists() or not root.is_dir():
            return files
        for path in root.rglob("*"):
            if path.is_file() and path.suffix.lower() in AUDIO_EXTENSIONS:
                files.append(path)

        files.sort(key=self._file_sort_key)
        return files

    def _file_sort_key(self, path: Path) -> tuple[int, str, str, str, str]:
        root = self._samples_dir
        try:
            if root is not None:
                rel_parent = path.relative_to(root).parent
                parent_depth = len(rel_parent.parts)
                rel_parent_text = str(rel_parent).lower()
            else:
                raise ValueError
        except ValueError:
            parent_depth = 999
            rel_parent_text = str(path.parent).lower()

        folder_block = path.parent.name.lower()
        filename = path.stem.lower()
        full_path = str(path).lower()
        return (parent_depth, folder_block, rel_parent_text, filename, full_path)

    def _load_cached_records_for_selected_root(self) -> int:
        if self._samples_dir is None:
            return 0

        records: list[JingleRecord] = []
        for path_key, _info in self._store.iter_entries():
            path = Path(path_key)
            try:
                path.relative_to(self._samples_dir)
            except ValueError:
                continue

            if path.suffix.lower() not in AUDIO_EXTENSIONS:
                continue

            categories = self._store.get(path)
            media_cache = self._store.get_media_cache(path)
            if media_cache is None:
                size_bytes = 0
                duration_seconds = 0.0
            else:
                size_bytes, duration_seconds, _mtime_ns = media_cache

            records.append(
                JingleRecord(
                    path=path,
                    categories=categories,
                    size_bytes=size_bytes,
                    duration_seconds=duration_seconds,
                )
            )

        records.sort(key=lambda record: self._file_sort_key(record.path))
        self._records = records
        self._apply_filters()
        return len(records)

    def _refresh_library_watcher_paths(self, files: list[Path]) -> None:
        if self._samples_dir is None or not self._watch_library_changes:
            self._library_watcher.removePaths(self._library_watcher.directories())
            self._library_watcher.removePaths(self._library_watcher.files())
            return

        watch_dirs: set[str] = {str(self._samples_dir)}
        for path in files:
            watch_dirs.add(str(path.parent))

        current_dirs = set(self._library_watcher.directories())
        remove_dirs = sorted(current_dirs - watch_dirs)
        add_dirs = sorted(watch_dirs - current_dirs)

        if remove_dirs:
            self._library_watcher.removePaths(remove_dirs)
        if add_dirs:
            self._library_watcher.addPaths(add_dirs)

    def _on_library_watch_path_changed(self, _path: str) -> None:
        if self._samples_dir is None or not self._watch_library_changes:
            return
        # Coalesce bursty filesystem events into one incremental reconcile pass.
        self._watch_rescan_timer.start(700)

    def _on_library_watch_rescan_timeout(self) -> None:
        if self._samples_dir is None or self._is_rescanning:
            return
        self._status.showMessage("Library changes detected. Refreshing...")
        self._rescan_library()

    def _rescan_library(self) -> None:
        if self._samples_dir is None:
            self._refresh_library_watcher_paths([])
            self._records = []
            self._apply_filters()
            return
        if self._is_rescanning:
            return
        self._is_rescanning = True
        QApplication.setOverrideCursor(QCursor(Qt.CursorShape.WaitCursor))
        try:
            files = self._scan_audio_files(self._samples_dir)
            self._refresh_library_watcher_paths(files)
            self._store.sync_with_files(files)

            records: list[JingleRecord] = []
            changed_count = 0
            for path in files:
                categories = self._store.get(path)
                try:
                    stat = path.stat()
                    size_bytes = int(stat.st_size)
                    mtime_ns = int(stat.st_mtime_ns)
                except OSError:
                    size_bytes = 0
                    mtime_ns = 0

                cached = self._store.get_media_cache(path)
                if (
                    cached is not None
                    and cached[0] == size_bytes
                    and cached[2] == mtime_ns
                ):
                    duration_seconds = cached[1]
                else:
                    duration_seconds = _probe_duration_seconds(path)
                    changed_count += 1
                self._store.set_media_cache(path, size_bytes, duration_seconds, mtime_ns)

                records.append(
                    JingleRecord(
                        path=path,
                        categories=categories,
                        size_bytes=size_bytes,
                        duration_seconds=duration_seconds,
                    )
                )

            self._records = records
            self._store.save()

            if self._auto_folder_tags:
                self._apply_folder_titles_to_records(self._records, preserve_existing=True)

            self._apply_filters()
            self._status.showMessage(
                f"Rescan complete: {len(records)} jingles ({changed_count} changed/new files re-probed)."
            )
        finally:
            QApplication.restoreOverrideCursor()
            self._is_rescanning = False

    def _apply_filters(self) -> None:
        query = self._search_edit.text().strip().casefold()
        scope_data = self._search_scope_combo.currentData()
        search_scope = str(scope_data) if scope_data is not None else "all"
        selected_categories = _normalize_tags(self._category_filter_edit.text())
        mode_data = self._category_filter_mode.currentData()
        category_mode = str(mode_data) if mode_data is not None else "any"
        self._refresh_filter_chips(selected_categories)

        visible: list[int] = []
        for index, record in enumerate(self._records):
            if selected_categories:
                record_keys = {tag.casefold() for tag in record.categories}
                selected_keys = {tag.casefold() for tag in selected_categories}
                if category_mode == "all":
                    if not selected_keys.issubset(record_keys):
                        continue
                else:
                    if record_keys.isdisjoint(selected_keys):
                        continue

            if query:
                if search_scope == "name":
                    haystack = record.name.casefold()
                elif search_scope == "tag":
                    haystack = _tags_to_text(record.categories).casefold()
                elif search_scope == "path":
                    haystack = str(record.path).casefold()
                else:
                    haystack = " ".join(
                        [
                            record.name,
                            _tags_to_text(record.categories),
                            str(record.path),
                        ]
                    ).casefold()
                if query not in haystack:
                    continue

            visible.append(index)

        self._visible_indices = visible
        self._rebuild_table()
        self._refresh_status_summary()

    def _selected_record_indices(self) -> list[int]:
        selected_rows = sorted({idx.row() for idx in self._table.selectedIndexes()})
        selected_record_indices: list[int] = []
        seen: set[int] = set()
        for row in selected_rows:
            if row < 0 or row >= len(self._visible_indices):
                continue
            record_index = self._visible_indices[row]
            if record_index in seen:
                continue
            seen.add(record_index)
            selected_record_indices.append(record_index)
        return selected_record_indices

    def _refresh_status_summary(self) -> None:
        total_count = len(self._records)
        shown_count = len(self._visible_indices)
        shown_bytes = sum(self._records[i].size_bytes for i in self._visible_indices)
        shown_seconds = sum(self._records[i].duration_seconds for i in self._visible_indices)

        message = (
            f"Showing {shown_count} of {total_count} jingles - "
            f"({_format_size_label(shown_bytes)}, {_format_duration_hms(shown_seconds)})"
        )

        selected_indices = self._selected_record_indices()
        selected_count = len(selected_indices)
        if selected_count > 0:
            selected_bytes = sum(self._records[i].size_bytes for i in selected_indices)
            selected_seconds = sum(self._records[i].duration_seconds for i in selected_indices)
            message += (
                f" | Selected {selected_count} "
                f"- ({_format_size_label(selected_bytes)}, {_format_duration_hms(selected_seconds)})"
            )

        self._status.showMessage(message)

    def _refresh_filter_chips(self, selected_categories: list[str]) -> None:
        self._clear_filters_btn.setEnabled(bool(selected_categories))

        while self._chips_layout.count():
            item = self._chips_layout.takeAt(0)
            widget = item.widget()
            if widget is not None:
                widget.deleteLater()

        if not selected_categories:
            empty = QLabel("None")
            empty.setStyleSheet("color: #888;")
            self._chips_layout.addWidget(empty)
            self._chips_layout.addStretch()
            return

        for tag in selected_categories:
            btn = QPushButton(f"{tag}  x")
            btn.setCursor(Qt.CursorShape.PointingHandCursor)
            # Keep chip color stable while typing by seeding on the first character.
            seed = tag.strip()[:1] or "_"
            bg, border, hover = _chip_palette_for_tag_seed(seed)
            btn.setStyleSheet(
                "QPushButton {"
                f" border: 1px solid {border};"
                " border-radius: 10px;"
                " padding: 3px 8px;"
                f" background: {bg};"
                " color: #ffffff;"
                "}"
                f"QPushButton:hover {{ background: {hover}; }}"
            )
            btn.clicked.connect(lambda _checked=False, t=tag: self._remove_filter_tag(t))
            self._chips_layout.addWidget(btn)

        self._chips_layout.addStretch()

    def _remove_filter_tag(self, tag_to_remove: str) -> None:
        current = _normalize_tags(self._category_filter_edit.text())
        keep = [tag for tag in current if tag.casefold() != tag_to_remove.casefold()]
        self._category_filter_edit.setText(_tags_to_text(keep))

    def _clear_all_filter_tags(self) -> None:
        self._search_edit.clear()
        self._category_filter_edit.clear()

    def _on_search_scope_changed(self) -> None:
        scope_data = self._search_scope_combo.currentData()
        scope = str(scope_data) if scope_data is not None else "all"
        if scope == "name":
            self._search_edit.setPlaceholderText("Search by jingle name")
        elif scope == "tag":
            self._search_edit.setPlaceholderText("Search by tags")
        elif scope == "path":
            self._search_edit.setPlaceholderText("Search by path")
        else:
            self._search_edit.setPlaceholderText("Search by jingle name, tags, or path")
        self._apply_filters()

    def _rebuild_table(self) -> None:
        self._updating_table = True
        self._table.blockSignals(True)
        self._table.setRowCount(len(self._visible_indices))

        for row, record_index in enumerate(self._visible_indices):
            record = self._records[record_index]
            jingle_item = QTableWidgetItem(record.name)
            jingle_item.setFlags(jingle_item.flags() & ~Qt.ItemFlag.ItemIsEditable)
            jingle_item.setData(Qt.ItemDataRole.UserRole, record_index)

            category_item = QTableWidgetItem(_tags_to_text(record.categories))
            category_item.setToolTip("Comma or semicolon separated tags")
            category_item.setData(Qt.ItemDataRole.UserRole, record_index)

            folder_item = QTableWidgetItem(record.folder)
            folder_item.setFlags(folder_item.flags() & ~Qt.ItemFlag.ItemIsEditable)
            folder_item.setData(Qt.ItemDataRole.UserRole, record_index)
            folder_item.setToolTip(str(record.path))

            self._table.setItem(row, 0, jingle_item)
            self._table.setItem(row, 1, category_item)
            self._table.setItem(row, 2, folder_item)

        self._table.blockSignals(False)
        self._updating_table = False
        self._table.resizeColumnToContents(1)
        self._table.resizeColumnToContents(2)

    def _on_table_item_changed(self, item: QTableWidgetItem) -> None:
        if self._updating_table:
            return

        col = item.column()
        if col != 1:
            return

        record_index = item.data(Qt.ItemDataRole.UserRole)
        if not isinstance(record_index, int):
            return
        if record_index < 0 or record_index >= len(self._records):
            return

        record = self._records[record_index]
        new_categories = _normalize_tags(self._table.item(item.row(), 1).text())

        record.categories = new_categories
        self._store.set(record.path, new_categories)
        self._store.save()

        self._apply_filters()

    def _on_apply_bulk_to_selected(self) -> None:
        selected_rows = sorted({idx.row() for idx in self._table.selectedIndexes()})
        if not selected_rows:
            self._status.showMessage("Select one or more rows first.")
            return

        categories = _normalize_tags(self._bulk_category_edit.text())
        mode_data = self._bulk_mode_combo.currentData()
        mode = str(mode_data) if mode_data is not None else "replace"

        if mode in {"append", "remove"} and not categories:
            verb = "append" if mode == "append" else "remove"
            self._status.showMessage(f"Enter one or more tags to {verb}.")
            return

        updated = 0
        for row in selected_rows:
            if row < 0 or row >= len(self._visible_indices):
                continue
            record_index = self._visible_indices[row]
            record = self._records[record_index]

            if mode == "append":
                new_categories = _merge_tags(record.categories, categories)
            elif mode == "remove":
                new_categories = _remove_tags(record.categories, categories)
            else:
                new_categories = list(categories)

            record.categories = new_categories
            self._store.set(record.path, new_categories)
            updated += 1

        self._store.save()
        self._apply_filters()
        action_text = {
            "append": "Appended",
            "remove": "Removed",
            "replace": "Updated",
        }.get(mode, "Updated")
        self._status.showMessage(f"{action_text} tags for {updated} jingle(s).")

    def _selected_record(self) -> JingleRecord | None:
        selected = self._table.selectedRanges()
        if not selected:
            return None
        row = selected[0].topRow()
        if row < 0 or row >= len(self._visible_indices):
            return None
        return self._records[self._visible_indices[row]]

    def _on_table_context_menu_requested(self, pos: Any) -> None:
        row = self._table.rowAt(pos.y())
        if row >= 0:
            clicked_item = self._table.item(row, 0)
            if clicked_item is not None and not clicked_item.isSelected():
                self._table.selectRow(row)

        selected_count = len(self._selected_record_indices())

        menu = QMenu(self)
        rename_action = menu.addAction("Rename")
        rename_action.setEnabled(selected_count == 1)
        rename_action.triggered.connect(self._on_edit_rename)

        delete_action = menu.addAction("Delete")
        delete_action.setEnabled(selected_count > 0)
        delete_action.triggered.connect(self._on_edit_delete)

        menu.exec(self._table.viewport().mapToGlobal(pos))

    def _on_table_item_double_clicked(self, item: QTableWidgetItem) -> None:
        if item.column() != 0:
            return
        self._table.selectRow(item.row())
        if (
            self._player is not None
            and self._player.playbackState() == QMediaPlayer.PlaybackState.PlayingState
        ):
            self._player.stop()
            self._reset_continuous_queue()
            self._current_playing_name = ""
        self._on_play_clicked()

    def _selected_record_index(self) -> int | None:
        selected = self._table.selectedRanges()
        if not selected:
            return None
        row = selected[0].topRow()
        if row < 0 or row >= len(self._visible_indices):
            return None
        return self._visible_indices[row]

    def _visible_row_for_record_index(self, record_index: int) -> int:
        for row, visible_record_index in enumerate(self._visible_indices):
            if visible_record_index == record_index:
                return row
        return -1

    def _select_record_row(self, record_index: int) -> None:
        row = self._visible_row_for_record_index(record_index)
        if row < 0:
            return
        self._table.selectRow(row)
        self._table.setCurrentCell(row, 0)

    def _reset_continuous_queue(self) -> None:
        self._continuous_queue = []
        self._continuous_queue_position = -1

    def _play_record(self, record_index: int) -> bool:
        if self._player is None:
            return False
        if record_index < 0 or record_index >= len(self._records):
            return False

        record = self._records[record_index]
        if not record.path.exists():
            return False

        self._apply_output_device()
        self._apply_player_loop_mode()
        self._player.setSource(QUrl.fromLocalFile(str(record.path)))
        self._player.play()
        self._current_playing_name = record.path.name
        self._select_record_row(record_index)

        mode_text = ""
        if self._playback_mode == "loop":
            mode_text = " (loop)"
        elif self._continuous_queue:
            total = len(self._continuous_queue)
            current = self._continuous_queue_position + 1 if total > 0 else 1
            if self._playback_mode == "continuous":
                mode_text = f" (continuous {current}/{max(total, 1)})"
            else:
                mode_text = f" (queue {current}/{max(total, 1)})"
        self._status.showMessage(f"Playing: {record.path.name}{mode_text}")
        return True

    def _start_continuous_playback(self) -> bool:
        selected_record_index = self._selected_record_index()
        if selected_record_index is None:
            self._status.showMessage("Select a jingle first.")
            return False

        start_row = self._visible_row_for_record_index(selected_record_index)
        if start_row < 0:
            return False

        self._continuous_queue = list(self._visible_indices[start_row:])
        if not self._continuous_queue:
            return False

        self._continuous_queue_position = -1
        return self._play_next_continuous_record()

    def _start_selected_queue_playback(self) -> bool:
        selected_indices = self._selected_record_indices()
        if len(selected_indices) < 2:
            return False

        self._continuous_queue = list(selected_indices)
        self._continuous_queue_position = -1
        return self._play_next_continuous_record()

    def _play_next_continuous_record(self) -> bool:
        if self._player is None:
            return False

        next_position = self._continuous_queue_position + 1
        while next_position < len(self._continuous_queue):
            record_index = self._continuous_queue[next_position]
            self._continuous_queue_position = next_position
            if self._play_record(record_index):
                return True
            next_position += 1

        self._reset_continuous_queue()
        return False

    def _on_play_clicked(self) -> None:
        if self._player is None:
            self._status.showMessage("Playback unavailable: PyQt6 multimedia is not installed.")
            return

        if self._player.playbackState() == QMediaPlayer.PlaybackState.PlayingState:
            self._player.pause()
            self._status.showMessage("Playback paused.")
            return

        if self._player.playbackState() == QMediaPlayer.PlaybackState.PausedState:
            self._player.play()
            self._status.showMessage("Playback resumed.")
            return

        selected_record_index = self._selected_record_index()
        if selected_record_index is None:
            self._status.showMessage("Select a jingle first.")
            return

        if self._playback_mode == "continuous":
            if not self._start_continuous_playback():
                self._current_playing_name = ""
                self._status.showMessage("No playable jingles were found from the selected row onward.")
            return

        if self._playback_mode == "off" and self._start_selected_queue_playback():
            return

        self._reset_continuous_queue()
        record = self._records[selected_record_index]
        if not record.path.exists():
            self._status.showMessage("Selected file no longer exists.")
            return

        self._play_record(selected_record_index)

    def _on_stop_clicked(self) -> None:
        if self._player is None:
            self._status.showMessage("Playback unavailable: PyQt6 multimedia is not installed.")
            return

        if self._player.playbackState() in (
            QMediaPlayer.PlaybackState.PlayingState,
            QMediaPlayer.PlaybackState.PausedState,
        ):
            self._player.stop()
            self._player.setPosition(0)
            self._reset_continuous_queue()
            self._current_playing_name = ""
            self._status.showMessage("Playback stopped.")

    def _on_table_selection_changed(self) -> None:
        record = self._selected_record()
        if record is not None:
            self._bulk_category_edit.setText(_tags_to_text(record.categories))
        else:
            self._bulk_category_edit.clear()
        self._refresh_status_summary()

    def _on_duration_changed(self, duration_ms: int) -> None:
        self._position_slider.setRange(0, max(0, duration_ms))
        self._update_time_label(self._player.position() if self._player is not None else 0, duration_ms)

    def _on_position_changed(self, position_ms: int) -> None:
        if not self._slider_pressed:
            self._position_slider.setValue(position_ms)
        duration = self._player.duration() if self._player is not None else 0
        self._update_time_label(position_ms, duration)

    def _on_playback_state_changed(self, _state: Any) -> None:
        if self._player is None:
            return
        is_playing = self._player.playbackState() == QMediaPlayer.PlaybackState.PlayingState
        self._set_play_button_state(is_playing)
        # Only reset slider when stopped, not when paused
        is_stopped = self._player.playbackState() == QMediaPlayer.PlaybackState.StoppedState
        if is_stopped:
            self._position_slider.setValue(0)
        # Make stop button breathe when playing or paused
        is_active = self._player.playbackState() in (
            QMediaPlayer.PlaybackState.PlayingState,
            QMediaPlayer.PlaybackState.PausedState,
        )
        self._set_stop_button_breathing(is_active)

    def _on_media_status_changed(self, status: QMediaPlayer.MediaStatus) -> None:
        if self._player is None:
            return
        if status == QMediaPlayer.MediaStatus.EndOfMedia:
            ended_name = self._current_playing_name
            if self._play_next_continuous_record():
                return
            self._reset_continuous_queue()
            self._current_playing_name = ""
            if ended_name:
                self._status.showMessage(f"Playback finished: {ended_name}")
            else:
                self._status.showMessage("Playback finished.")

    def _refresh_playback_mode_button(self) -> None:
        if self._playback_mode == "loop":
            self._loop_btn.setText("Loop On")
            self._loop_btn.setStyleSheet(
                "QPushButton { background-color: #0d47a1; color: white; font-weight: bold; }"
                "QPushButton:hover { background-color: #1565c0; }"
            )
            self._set_loop_breathing(self._loop_btn.isEnabled())
        elif self._playback_mode == "continuous":
            self._loop_btn.setText("Continuous")
            self._loop_btn.setStyleSheet(
                "QPushButton { background-color: #ef6c00; color: white; font-weight: bold; }"
                "QPushButton:hover { background-color: #fb8c00; }"
            )
            self._set_loop_breathing(self._loop_btn.isEnabled())
        else:
            self._loop_btn.setText("Loop Off")
            self._loop_btn.setStyleSheet("")
            self._set_loop_breathing(False)

    def _on_loop_clicked(self) -> None:
        if self._playback_mode == "off":
            self._playback_mode = "loop"
        elif self._playback_mode == "loop":
            self._playback_mode = "continuous"
        else:
            self._playback_mode = "off"
        self._refresh_playback_mode_button()
        self._apply_player_loop_mode()

        # Switching to Loop Off: clear any active queue so the current track
        # finishes and playback stops naturally.
        if self._playback_mode == "off":
            self._reset_continuous_queue()
            return

        # If switching to continuous while playback is active, seed the queue
        # from the currently playing row so the next EndOfMedia can advance.
        if (
            self._playback_mode == "continuous"
            and self._player is not None
            and self._player.playbackState() == QMediaPlayer.PlaybackState.PlayingState
            and not self._continuous_queue
        ):
            current_index = self._selected_record_index()
            if current_index is not None:
                start_row = self._visible_row_for_record_index(current_index)
                if start_row >= 0:
                    self._continuous_queue = list(self._visible_indices[start_row:])
                    # Position 0 is the currently playing track; next advance starts at 1.
                    self._continuous_queue_position = 0

    def _apply_player_loop_mode(self) -> None:
        if self._player is None:
            return
        # Native backend looping avoids manual restart gaps between iterations.
        self._player.setLoops(-1 if self._playback_mode == "loop" else 1)

    def _set_loop_breathing(self, enabled: bool) -> None:
        if not enabled:
            if self._loop_breath_anim is not None:
                self._loop_breath_anim.stop()
            if self._loop_breath_effect is not None:
                try:
                    self._loop_breath_effect.setOpacity(1.0)
                except RuntimeError:
                    pass
            self._loop_btn.setGraphicsEffect(None)
            self._loop_breath_anim = None
            self._loop_breath_effect = None
            return

        if self._loop_breath_effect is None:
            self._loop_breath_effect = QGraphicsOpacityEffect(self._loop_btn)
        try:
            self._loop_btn.setGraphicsEffect(self._loop_breath_effect)
        except RuntimeError:
            self._loop_breath_effect = QGraphicsOpacityEffect(self._loop_btn)
            self._loop_btn.setGraphicsEffect(self._loop_breath_effect)
            self._loop_breath_anim = None

        if self._loop_breath_anim is None:
            self._loop_breath_anim = QPropertyAnimation(
                self._loop_breath_effect,
                b"opacity",
                self,
            )
            self._loop_breath_anim.setDuration(1100)
            self._loop_breath_anim.setStartValue(1.0)
            self._loop_breath_anim.setKeyValueAt(0.5, 0.72)
            self._loop_breath_anim.setEndValue(1.0)
            self._loop_breath_anim.setLoopCount(-1)
            self._loop_breath_anim.setEasingCurve(QEasingCurve.Type.InOutSine)

        self._loop_breath_anim.start()

    def _set_play_button_state(self, is_playing: bool) -> None:
        if is_playing:
            self._play_btn.setText("Pause")
            self._play_btn.setStyleSheet(
                "QPushButton { background-color: #f57c00; color: white; font-weight: bold; }"
                "QPushButton:hover { background-color: #e65100; }"
            )
            self._set_play_stop_breathing(self._play_btn.isEnabled())
        else:
            self._play_btn.setText("Play Selected")
            self._play_btn.setStyleSheet(
                "QPushButton { background-color: #2e7d32; color: white; font-weight: bold; }"
                "QPushButton:hover { background-color: #388e3c; }"
            )
            self._set_play_stop_breathing(False)

    def _set_play_stop_breathing(self, enabled: bool) -> None:
        if not enabled:
            if self._play_stop_breath_anim is not None:
                self._play_stop_breath_anim.stop()
            if self._play_stop_breath_effect is not None:
                try:
                    self._play_stop_breath_effect.setOpacity(1.0)
                except RuntimeError:
                    pass
            self._play_btn.setGraphicsEffect(None)
            self._play_stop_breath_anim = None
            self._play_stop_breath_effect = None
            return

        if self._play_stop_breath_effect is None:
            self._play_stop_breath_effect = QGraphicsOpacityEffect(self._play_btn)
        try:
            self._play_btn.setGraphicsEffect(self._play_stop_breath_effect)
        except RuntimeError:
            self._play_stop_breath_effect = QGraphicsOpacityEffect(self._play_btn)
            self._play_btn.setGraphicsEffect(self._play_stop_breath_effect)
            self._play_stop_breath_anim = None

        if self._play_stop_breath_anim is None:
            self._play_stop_breath_anim = QPropertyAnimation(
                self._play_stop_breath_effect,
                b"opacity",
                self,
            )
            self._play_stop_breath_anim.setDuration(1000)
            self._play_stop_breath_anim.setStartValue(1.0)
            self._play_stop_breath_anim.setKeyValueAt(0.5, 0.68)
            self._play_stop_breath_anim.setEndValue(1.0)
            self._play_stop_breath_anim.setLoopCount(-1)
            self._play_stop_breath_anim.setEasingCurve(QEasingCurve.Type.InOutSine)

        self._play_stop_breath_anim.start()

    def _set_stop_button_breathing(self, enabled: bool) -> None:
        if not enabled:
            if self._stop_btn_breath_anim is not None:
                self._stop_btn_breath_anim.stop()
            if self._stop_btn_breath_effect is not None:
                try:
                    self._stop_btn_breath_effect.setOpacity(1.0)
                except RuntimeError:
                    pass
            self._stop_btn.setGraphicsEffect(None)
            self._stop_btn_breath_anim = None
            self._stop_btn_breath_effect = None
            return

        if self._stop_btn_breath_effect is None:
            self._stop_btn_breath_effect = QGraphicsOpacityEffect(self._stop_btn)
        try:
            self._stop_btn.setGraphicsEffect(self._stop_btn_breath_effect)
        except RuntimeError:
            self._stop_btn_breath_effect = QGraphicsOpacityEffect(self._stop_btn)
            self._stop_btn.setGraphicsEffect(self._stop_btn_breath_effect)
            self._stop_btn_breath_anim = None

        if self._stop_btn_breath_anim is None:
            self._stop_btn_breath_anim = QPropertyAnimation(
                self._stop_btn_breath_effect,
                b"opacity",
                self,
            )
            self._stop_btn_breath_anim.setDuration(1000)
            self._stop_btn_breath_anim.setStartValue(1.0)
            self._stop_btn_breath_anim.setKeyValueAt(0.5, 0.68)
            self._stop_btn_breath_anim.setEndValue(1.0)
            self._stop_btn_breath_anim.setLoopCount(-1)
            self._stop_btn_breath_anim.setEasingCurve(QEasingCurve.Type.InOutSine)

        self._stop_btn_breath_anim.start()

    def _on_slider_pressed(self) -> None:
        self._slider_pressed = True

    def _on_slider_released(self) -> None:
        self._slider_pressed = False
        if self._player is not None:
            self._player.setPosition(int(self._position_slider.value()))

    def _update_time_label(self, position_ms: int, duration_ms: int) -> None:
        self._time_label.setText(
            f"{self._fmt_time(position_ms)} / {self._fmt_time(duration_ms)}"
        )

    @staticmethod
    def _fmt_time(ms: int) -> str:
        total = max(0, int(ms / 1000))
        minutes, seconds = divmod(total, 60)
        hours, minutes = divmod(minutes, 60)
        if hours > 0:
            return f"{hours:02d}:{minutes:02d}:{seconds:02d}"
        return f"{minutes:02d}:{seconds:02d}"

    def _on_open_options(self) -> None:
        dialog = OptionsDialog(
            self._output_device,
            self._preview_output_device,
            self._live_volume_percent,
            self._preview_volume_percent,
            self._samples_dir,
            self,
        )
        if dialog.exec() != int(QDialog.DialogCode.Accepted):
            return

        self._output_device, self._preview_output_device = dialog.selected_devices()
        self._live_volume_percent, self._preview_volume_percent = dialog.selected_volumes()
        self._settings.setValue("options/outputDevice", self._output_device)
        self._settings.setValue("options/previewOutputDevice", self._preview_output_device)
        self._save_volume_settings()
        self._refresh_mode_toggle_state(notify_if_disabled=True)
        self._refresh_volume_controls()
        self._apply_output_device()

        new_dir = dialog.selected_folder()
        if new_dir != self._samples_dir:
            self._samples_dir = new_dir
            self._save_samples_dir()
            self._rescan_library()

        if not self._can_use_preview_mode():
            QMessageBox.information(
                self,
                "Preview/Live Disabled",
                "Live and Preview devices are currently the same.\n\n"
                "Preview/Live switching is disabled until the Preview device is set to a different output.",
            )
        self._status.showMessage("Options saved.")

    def _apply_output_device(self) -> None:
        if self._audio_output is None or not _has_qt_multimedia:
            return

        selected = self._active_output_device().strip()
        target_device = QMediaDevices.defaultAudioOutput()

        if selected:
            matched = None
            for device in QMediaDevices.audioOutputs():
                if device.description().strip() == selected:
                    matched = device
                    break
            if matched is not None:
                target_device = matched
            else:
                self._status.showMessage(
                    f"Selected device unavailable. Using system default: {target_device.description()}"
                )

        self._audio_output.setDevice(target_device)
        self._audio_output.setMuted(self._is_muted)
        self._apply_active_volume()

    def _normalize_device_key(self, value: str) -> str:
        return value.strip().casefold()

    def _can_use_preview_mode(self) -> bool:
        return self._normalize_device_key(self._output_device) != self._normalize_device_key(
            self._preview_output_device
        )

    def _active_output_device(self) -> str:
        if self._is_preview_mode and self._can_use_preview_mode():
            return self._preview_output_device
        return self._output_device

    def _refresh_mode_toggle_state(self, notify_if_disabled: bool) -> None:
        can_use = self._can_use_preview_mode()
        if not can_use and self._is_preview_mode:
            self._is_preview_mode = False
            self._mode_btn.blockSignals(True)
            self._mode_btn.setChecked(False)
            self._mode_btn.blockSignals(False)

        self._mode_btn.setEnabled(can_use)
        if can_use:
            self._mode_btn.setToolTip("Toggle between Live and Preview output devices")
        else:
            self._mode_btn.setToolTip(
                "Preview/Live switch is disabled because Live and Preview devices are the same"
            )

        self._set_mode_button_visual()
        self._refresh_volume_controls()

        if notify_if_disabled and not can_use:
            self._status.showMessage(
                "Preview/Live switch disabled: set Preview device to a different output in Options."
            )

    def _set_mode_button_visual(self) -> None:
        if self._is_preview_mode:
            self._mode_btn.setText("Mode: Preview")
            self._mode_btn.setStyleSheet(
                "QPushButton { background-color: #1565c0; color: white; font-weight: bold; }"
                "QPushButton:hover { background-color: #1976d2; }"
            )
            self._set_mode_live_breathing(False)
        else:
            self._mode_btn.setText("Mode: Live")
            self._mode_btn.setStyleSheet(
                "QPushButton { background-color: #b71c1c; color: white; font-weight: bold; }"
                "QPushButton:hover { background-color: #c62828; }"
            )
            self._set_mode_live_breathing(self._mode_btn.isEnabled())

    def _set_mode_live_breathing(self, enabled: bool) -> None:
        if not enabled:
            if self._mode_live_breath_anim is not None:
                self._mode_live_breath_anim.stop()
            if self._mode_live_breath_effect is not None:
                try:
                    self._mode_live_breath_effect.setOpacity(1.0)
                except RuntimeError:
                    pass
            self._mode_btn.setGraphicsEffect(None)
            # Qt may delete the installed effect when detached; reset cached refs.
            self._mode_live_breath_anim = None
            self._mode_live_breath_effect = None
            return

        if self._mode_live_breath_effect is None:
            self._mode_live_breath_effect = QGraphicsOpacityEffect(self._mode_btn)
        try:
            self._mode_btn.setGraphicsEffect(self._mode_live_breath_effect)
        except RuntimeError:
            self._mode_live_breath_effect = QGraphicsOpacityEffect(self._mode_btn)
            self._mode_btn.setGraphicsEffect(self._mode_live_breath_effect)
            self._mode_live_breath_anim = None

        if self._mode_live_breath_anim is None:
            self._mode_live_breath_anim = QPropertyAnimation(
                self._mode_live_breath_effect,
                b"opacity",
                self,
            )
            self._mode_live_breath_anim.setDuration(1300)
            self._mode_live_breath_anim.setStartValue(1.0)
            self._mode_live_breath_anim.setKeyValueAt(0.5, 0.72)
            self._mode_live_breath_anim.setEndValue(1.0)
            self._mode_live_breath_anim.setLoopCount(-1)
            self._mode_live_breath_anim.setEasingCurve(QEasingCurve.Type.InOutSine)

        self._mode_live_breath_anim.start()

    def _on_mode_toggled(self, checked: bool) -> None:
        if checked and not self._can_use_preview_mode():
            self._mode_btn.blockSignals(True)
            self._mode_btn.setChecked(False)
            self._mode_btn.blockSignals(False)
            self._is_preview_mode = False
            self._set_mode_button_visual()
            self._refresh_volume_controls()
            return

        self._is_preview_mode = bool(checked)
        self._set_mode_button_visual()
        self._refresh_volume_controls()
        self._apply_output_device()

    def _on_browse_folder(self) -> None:
        selected = QFileDialog.getExistingDirectory(
            self,
            "Choose Samples Folder",
            str(self._samples_dir) if self._samples_dir else str(Path.home()),
        )
        if not selected:
            return

        path = Path(selected)
        if not path.exists() or not path.is_dir():
            QMessageBox.warning(self, "Invalid Folder", "Please choose a valid folder.")
            return

        self._samples_dir = path
        self._save_samples_dir()
        self._rescan_library()


def main() -> None:
    if sys.platform == "win32":
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID("JingleAllTheDay.App")
    app = QApplication(sys.argv)
    app.setApplicationName(APP_NAME)
    icon_path = _HERE / "icon.png"
    if icon_path.exists():
        app.setWindowIcon(QIcon(str(icon_path)))
    window = MainWindow()
    _apply_windows_taskbar_icon(window)
    window.show()
    try:
        exit_code = app.exec()
    except KeyboardInterrupt:
        # Gracefully handle Ctrl+C / debugger stop without noisy tracebacks.
        exit_code = 130
    sys.exit(exit_code)


if __name__ == "__main__":
    main()
