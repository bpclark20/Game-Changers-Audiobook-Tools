#!/usr/bin/env python3
"""GUI front-end for wav_markers_to_mp3.py — chapter list viewer."""

from __future__ import annotations

import io
import subprocess
import sys
import tempfile
from dataclasses import replace
from decimal import Decimal, InvalidOperation
from pathlib import Path
from typing import Any, Callable, cast

# Ensure wav_markers_to_mp3 is importable from the same directory
_HERE = Path(__file__).parent
if str(_HERE) not in sys.path:
    sys.path.insert(0, str(_HERE))

from wav_markers_to_mp3 import (  # noqa: E402
    Chapter,
    build_chapters,
    encode_output,
    ensure_executable,
    format_bytes,
    format_seconds_hms,
    probe_audio_info,
    read_riff_chunks,
    resolve_cover_images,
)

try:
    from PyQt6.QtCore import QEvent, QModelIndex, QObject, QSettings, QThread, QTimer, Qt, pyqtSignal  # noqa: E402
    from PyQt6.QtGui import QColor, QDropEvent, QMouseEvent, QPainter, QPixmap, QResizeEvent  # noqa: E402
    from PyQt6.QtWidgets import (  # noqa: E402
        QAbstractItemView,
        QApplication,
        QButtonGroup,
        QComboBox,
        QFileDialog,
        QHBoxLayout,
        QHeaderView,
        QLabel,
        QLineEdit,
        QMainWindow,
        QMessageBox,
        QPushButton,
        QRadioButton,
        QProgressBar,
        QSizePolicy,
        QSplitter,
        QSpinBox,
        QStatusBar,
        QStyledItemDelegate,
        QStyle,
        QStyleOptionViewItem,
        QTableWidget,
        QTableWidgetItem,
        QVBoxLayout,
        QWidget,
    )
except ModuleNotFoundError as e:
    print(
        "\n❌ Error: PyQt6 is not installed.\n"
        "PyQt6 is installed in the project's virtual environment (.venv).\n"
        "To run this GUI, use one of these methods:\n\n"
        "  Method 1 (Recommended - Activate venv first):\n"
        f"    & '.venv\\Scripts\\Activate.ps1'\n"
        "    python gui.py\n\n"
        "  Method 2 (Direct - Use venv Python):\n"
        f"    & '.venv\\Scripts\\python.exe' gui.py\n\n"
        "  Method 3 (Create a shortcut):\n"
        "    Create 'run_gui.ps1' with the Method 1 commands,\n"
        "    then run: .\\run_gui.ps1\n\n"
        "  Method 4 (NOT RECOMMENDED - Install to system Python):\n"
        "    pip install PyQt6\n"
        "    python gui.py\n"
        "    (This may cause version conflicts with other projects)\n",
        file=sys.stderr,
    )
    sys.exit(1)

THUMB_SIZE = 48  # cover thumbnail height/width in pixels
CHAPTER_PROGRESS_ROLE = int(Qt.ItemDataRole.UserRole) + 1
CHAPTER_KEY_ROLE = int(Qt.ItemDataRole.UserRole) + 2


def _format_duration_label(duration_sec: float) -> str:
    total_seconds = max(0, int(duration_sec))
    hours, remainder = divmod(total_seconds, 3600)
    minutes, seconds = divmod(remainder, 60)
    return f"{hours:02d}:{minutes:02d}:{seconds:02d}"


def _chapter_key(chapter: Chapter) -> str:
    return f"{chapter.start_ms}:{chapter.end_ms}"


class ChapterProgressDelegate(QStyledItemDelegate):
    def paint(
        self,
        painter: QPainter | None,
        option: QStyleOptionViewItem,
        index: QModelIndex,
    ) -> None:
        if painter is None:
            return
        progress_data = index.data(CHAPTER_PROGRESS_ROLE)
        if not isinstance(progress_data, (int, float)) or progress_data <= 0:
            super().paint(painter, option, index)
            return

        progress = max(0.0, min(1.0, float(progress_data)))
        paint_option = QStyleOptionViewItem(option)
        self.initStyleOption(paint_option, index)
        text = paint_option.text
        paint_option.text = ""

        style = paint_option.widget.style() if paint_option.widget else QApplication.style()
        if style is None:
            super().paint(painter, option, index)
            return
        style.drawControl(QStyle.ControlElement.CE_ItemViewItem, paint_option, painter, paint_option.widget)

        content_rect = paint_option.rect.adjusted(2, 2, -2, -2)
        if content_rect.width() > 0 and content_rect.height() > 0:
            fill_width = max(1, int(content_rect.width() * progress))
            fill_rect = content_rect.adjusted(0, 0, -(content_rect.width() - fill_width), 0)
            painter.save()
            painter.setPen(Qt.PenStyle.NoPen)
            painter.setBrush(QColor(70, 140, 90, 140))
            painter.drawRect(fill_rect)
            painter.restore()

        painter.save()
        text_rect = content_rect.adjusted(6, 0, -6, 0)
        if paint_option.state & QStyle.StateFlag.State_Selected:
            text_role = paint_option.palette.ColorRole.HighlightedText
        else:
            text_role = paint_option.palette.ColorRole.Text
        painter.setPen(paint_option.palette.color(text_role))
        painter.drawText(
            text_rect,
            Qt.AlignmentFlag.AlignVCenter | Qt.AlignmentFlag.AlignLeft,
            text,
        )
        painter.restore()


class ReorderableTableWidget(QTableWidget):
    rowMoveRequested = pyqtSignal(int, int)

    def __init__(self, *args: Any, **kwargs: Any) -> None:
        super().__init__(*args, **kwargs)
        self._drag_row = -1

    def mousePressEvent(self, e: QMouseEvent | None) -> None:
        if e is not None and e.button() == Qt.MouseButton.LeftButton:
            self._drag_row = self.rowAt(e.position().toPoint().y())
        super().mousePressEvent(e)

    def dropEvent(self, event: QDropEvent | None) -> None:
        if event is None:
            return
        source_row = self._drag_row
        self._drag_row = -1
        if source_row < 0:
            event.ignore()
            return

        drop_pos = event.position().toPoint()
        target_row = self.rowAt(drop_pos.y())
        if target_row < 0:
            target_row = self.rowCount()
        elif self.dropIndicatorPosition() == QAbstractItemView.DropIndicatorPosition.BelowItem:
            target_row += 1

        self.rowMoveRequested.emit(source_row, target_row)
        event.setDropAction(Qt.DropAction.CopyAction)
        event.accept()


def _sanitize_output_stem(title: str) -> str:
    sanitized = "".join(
        "_" if ch in '<>:"/\\|?*' else ch
        for ch in title.strip().rstrip(". ")
    )
    sanitized = sanitized.strip()
    return sanitized or "output"


def _book_title_from_chapter(chapter_title: str) -> str | None:
    parts = chapter_title.split(" - ", 1)
    if len(parts) != 2:
        return None
    title = parts[0].strip()
    return title or None


def _find_covers_by_title(input_dir: Path) -> dict[str, Path]:
    """Map case-insensitive image stem -> preferred image path."""
    image_files = [
        child
        for child in input_dir.iterdir()
        if child.is_file() and child.suffix.lower() in {".jpg", ".jpeg", ".png"}
    ]
    by_stem: dict[str, list[Path]] = {}
    for image in image_files:
        by_stem.setdefault(image.stem.casefold(), []).append(image)

    preferred_ext = {".jpg": 0, ".jpeg": 1, ".png": 2}
    resolved: dict[str, Path] = {}
    for stem, candidates in by_stem.items():
        candidates.sort(key=lambda p: preferred_ext.get(p.suffix.lower(), 99))
        resolved[stem] = candidates[0]
    return resolved


def _run_capturing_errors(
    fn: Callable[..., Any], *args: Any, **kwargs: Any
) -> tuple[Any, None] | tuple[None, str]:
    """Call fn(*args, **kwargs).

    Returns (result, None) on success.
    Returns (None, error_message) if the function raises SystemExit (via fail())
    or KeyboardInterrupt (can be fired spuriously by Python signal handling
    when subprocess.run is called on the Qt main thread on Windows).
    """
    buf = io.StringIO()
    old_stderr = sys.stderr
    sys.stderr = buf
    try:
        result = fn(*args, **kwargs)
        return result, None
    except (SystemExit, KeyboardInterrupt):
        msg = buf.getvalue().strip()
        if msg.startswith("ERROR: "):
            msg = msg[len("ERROR: "):]
        return None, msg or "Operation cancelled"
    finally:
        sys.stderr = old_stderr


class EncodeWorker(QObject):
    progress = pyqtSignal(int, str, float)
    finished = pyqtSignal(str, int)
    failed = pyqtSignal(str)

    def __init__(
        self,
        input_wav: Path,
        output_mp3: Path,
        chapters: list[Chapter],
        channels: int,
        vbr_quality: int,
        cbr_bitrate: str | None,
    ) -> None:
        super().__init__()
        self._input_wav = input_wav
        self._output_mp3 = output_mp3
        self._chapters = list(chapters)
        self._channels = channels
        self._vbr_quality = vbr_quality
        self._cbr_bitrate = cbr_bitrate
        self._cancel_requested = False

    def request_cancel(self) -> None:
        self._cancel_requested = True

    def _build_reordered_wav(
        self,
        ffmpeg_bin: str,
        input_wav: Path,
        chapters: list[Chapter],
    ) -> tuple[Path, list[Chapter], float, tempfile.TemporaryDirectory[str]]:
        temp_dir = tempfile.TemporaryDirectory(prefix="wav_reorder_")
        reordered_wav = Path(temp_dir.name) / "reordered.wav"

        filter_parts: list[str] = []
        concat_inputs: list[str] = []
        for idx, ch in enumerate(chapters):
            start_sec = f"{ch.start_ms / 1000.0:.6f}"
            end_sec = f"{ch.end_ms / 1000.0:.6f}"
            label = f"c{idx}"
            filter_parts.append(
                f"[0:a]atrim=start={start_sec}:end={end_sec},asetpts=PTS-STARTPTS[{label}]"
            )
            concat_inputs.append(f"[{label}]")

        filter_parts.append(
            f"{''.join(concat_inputs)}concat=n={len(chapters)}:v=0:a=1[outa]"
        )
        filter_complex = ";".join(filter_parts)

        cmd = [
            ffmpeg_bin,
            "-y",
            "-i",
            str(input_wav),
            "-filter_complex",
            filter_complex,
            "-map",
            "[outa]",
            "-c:a",
            "pcm_s16le",
            str(reordered_wav),
        ]
        result = subprocess.run(cmd, text=True, capture_output=True)
        if result.returncode != 0:
            raise SystemExit(
                "WAV chapter reordering failed. "
                f"ffmpeg stderr: {result.stderr.strip() or '(empty)'}"
            )

        remapped: list[Chapter] = []
        cursor_ms = 0
        for idx, ch in enumerate(chapters, start=1):
            seg_dur_ms = max(1, ch.end_ms - ch.start_ms)
            start_ms = cursor_ms
            end_ms = start_ms + seg_dur_ms
            remapped.append(
                Chapter(index=idx, title=ch.title, start_ms=start_ms, end_ms=end_ms)
            )
            cursor_ms = end_ms

        return reordered_wav, remapped, cursor_ms / 1000.0, temp_dir

    def run(self) -> None:
        stderr_buf = io.StringIO()
        old_stderr = sys.stderr
        sys.stderr = stderr_buf
        reorder_tempdir: tempfile.TemporaryDirectory[str] | None = None
        try:
            self.progress.emit(0, "Preparing encoder…", -1.0)
            ffmpeg_bin = ensure_executable("ffmpeg")
            ffprobe_bin = ensure_executable("ffprobe")

            self.progress.emit(2, "Reading markers…", -1.0)
            sample_rate, _markers = read_riff_chunks(self._input_wav)
            duration_sec, _ = probe_audio_info(ffprobe_bin, self._input_wav)
            chapters = list(self._chapters)
            cover_map = resolve_cover_images(self._input_wav.parent, chapters)

            input_for_encode = self._input_wav
            chapters_for_encode = chapters
            duration_for_encode = duration_sec
            needs_reorder = any(
                chapters[i].start_ms > chapters[i + 1].start_ms
                for i in range(len(chapters) - 1)
            )
            if needs_reorder:
                if self._cancel_requested:
                    raise SystemExit
                self.progress.emit(4, "Reordering chapters in WAV…", -1.0)
                (
                    input_for_encode,
                    chapters_for_encode,
                    duration_for_encode,
                    reorder_tempdir,
                ) = self._build_reordered_wav(ffmpeg_bin, self._input_wav, chapters)

            self.progress.emit(5, "Encoding audio…", -1.0)

            def _progress_callback(percent: float, status_line: str, out_time_sec: float) -> None:
                pct = int(max(0.0, min(100.0, percent)))
                self.progress.emit(pct, f"Encoding… {status_line.strip()}", out_time_sec)

            actual_size_bytes = encode_output(
                ffmpeg_bin=ffmpeg_bin,
                input_file=input_for_encode,
                chapters=chapters_for_encode,
                duration_sec=duration_for_encode,
                output_file=self._output_mp3,
                channels=self._channels,
                vbr_quality=self._vbr_quality,
                cbr_bitrate=self._cbr_bitrate,
                cover_map=cover_map,
                progress_callback=_progress_callback,
                cancel_requested=lambda: self._cancel_requested,
                sample_rate=sample_rate,
            )
            self.progress.emit(100, "Encoding complete", -1.0)
            self.finished.emit(
                f"Created: {self._output_mp3} ({format_bytes(actual_size_bytes)})",
                actual_size_bytes,
            )
        except SystemExit:
            if self._cancel_requested:
                self.failed.emit("Encode cancelled by user.")
            else:
                msg = stderr_buf.getvalue().strip()
                if msg.startswith("ERROR: "):
                    msg = msg[len("ERROR: "):]
                self.failed.emit(msg or "Encoding failed.")
        except Exception as exc:
            self.failed.emit(str(exc) or "Unexpected encoding error.")
        finally:
            if reorder_tempdir is not None:
                reorder_tempdir.cleanup()
            sys.stderr = old_stderr


class MainWindow(QMainWindow):
    def __init__(self) -> None:
        super().__init__()
        self.setWindowTitle("AudioBook Slicer")
        self.resize(900, 600)

        self._wav_path: Path | None = None
        self._output_dir: Path | None = None
        self._chapters: list[Chapter] = []
        self._original_chapters: list[Chapter] = []
        self._missing_cover_rows: set[int] = set()
        self._row_cover_paths: dict[int, Path] = {}
        self._preview_base_pixmap: QPixmap | None = None
        self._next_cover_action = "all"
        self._encode_mode_group: QButtonGroup | None = None
        self._channel_group: QButtonGroup | None = None
        self._size_unit_group: QButtonGroup | None = None
        self._encode_thread: QThread | None = None
        self._encode_worker: EncodeWorker | None = None
        self._initial_splitter_sizes: tuple[int, int] = (220, 420)
        self._initial_splitter_ratio: float | None = None
        self._initial_splitter_top_px: int | None = None
        self._splitter_reset_pending = False
        self._encode_progress_row: int | None = None
        self._encode_progress_timeline: list[Chapter] = []
        self._is_populating_table = False
        self._was_maximized = self.isMaximized()
        self._remembered_conflict_path: Path | None = None
        self._remembered_conflict_choice: str | None = None
        self._settings = QSettings("AudioBookSlicer", "AudioBookSlicer")

        # ── Central widget ────────────────────────────────────────────────────
        central = QWidget()
        self.setCentralWidget(central)
        root = QVBoxLayout(central)
        root.setContentsMargins(8, 8, 8, 8)
        root.setSpacing(6)

        # ── Top bar: file picker ──────────────────────────────────────────────
        top = QHBoxLayout()

        self._open_btn = QPushButton("Open WAV File…")
        self._open_btn.setFixedWidth(140)
        self._open_btn.clicked.connect(self._on_open_wav)
        top.addWidget(self._open_btn)

        self._reset_btn = QPushButton("Reset")
        self._reset_btn.setFixedWidth(70)
        self._reset_btn.clicked.connect(self._on_reset_clicked)
        top.addWidget(self._reset_btn)

        self._reset_markers_btn = QPushButton("Reset Markers")
        self._reset_markers_btn.setFixedWidth(100)
        self._reset_markers_btn.setEnabled(False)
        self._reset_markers_btn.clicked.connect(self._on_reset_markers_clicked)
        top.addWidget(self._reset_markers_btn)

        self._file_label = QLabel("No file selected")
        self._file_label.setSizePolicy(
            QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Preferred
        )
        top.addWidget(self._file_label)

        # ── Output settings + cover preview ─────────────────────────────────
        top_settings_row = QHBoxLayout()
        top_settings_row.setAlignment(Qt.AlignmentFlag.AlignTop)
        top_settings_row.setContentsMargins(0, 0, 0, 0)
        top_settings_row.setSpacing(6)

        settings_left = QVBoxLayout()
        settings_left.setSpacing(6)
        settings_left.setAlignment(Qt.AlignmentFlag.AlignTop)
        settings_left.addLayout(top)

        title_row = QHBoxLayout()

        title_label = QLabel("Book Title")
        title_label.setFixedWidth(90)
        title_row.addWidget(title_label)

        self._book_title_edit = QLineEdit()
        self._book_title_edit.setPlaceholderText("Book title for MP3 tags")
        self._book_title_edit.textChanged.connect(self._on_book_title_changed)
        title_row.addWidget(self._book_title_edit)

        settings_left.addLayout(title_row)

        output_row = QHBoxLayout()

        output_label = QLabel("Output File")
        output_label.setFixedWidth(90)
        output_row.addWidget(output_label)

        self._output_path_edit = QLineEdit()
        self._output_path_edit.setReadOnly(True)
        output_row.addWidget(self._output_path_edit)

        self._save_as_btn = QPushButton("Save As")
        self._save_as_btn.setEnabled(False)
        self._save_as_btn.clicked.connect(self._on_save_as)
        output_row.addWidget(self._save_as_btn)

        settings_left.addLayout(output_row)

        duration_row = QHBoxLayout()

        duration_label = QLabel("Duration")
        duration_label.setFixedWidth(90)
        duration_row.addWidget(duration_label)

        self._duration_edit = QLineEdit()
        self._duration_edit.setReadOnly(True)
        duration_row.addWidget(self._duration_edit)
        duration_row.addStretch()

        settings_left.addLayout(duration_row)

        encoding_row = QHBoxLayout()

        encoding_label = QLabel("Encoding")
        encoding_label.setFixedWidth(90)
        encoding_row.addWidget(encoding_label)

        self._vbr_radio = QRadioButton("VBR")
        self._cbr_radio = QRadioButton("CBR")
        self._encode_mode_group = QButtonGroup(self)
        self._encode_mode_group.addButton(self._vbr_radio)
        self._encode_mode_group.addButton(self._cbr_radio)
        self._vbr_radio.setChecked(True)
        self._vbr_radio.toggled.connect(self._on_encode_mode_changed)
        encoding_row.addWidget(self._vbr_radio)
        encoding_row.addWidget(self._cbr_radio)

        self._vbr_quality_label = QLabel("quality level")
        encoding_row.addWidget(self._vbr_quality_label)

        self._vbr_q_combo = QComboBox()
        self._vbr_q_combo.setFixedWidth(70)
        for value in range(10):
            self._vbr_q_combo.addItem(str(value), value)
        self._vbr_q_combo.setCurrentIndex(5)
        encoding_row.addWidget(self._vbr_q_combo)

        self._cbr_radio.toggled.connect(self._on_encode_mode_changed)
        self._cbr_bitrate_label = QLabel("kbps")
        encoding_row.addWidget(self._cbr_bitrate_label)

        self._cbr_bitrate_spin = QSpinBox()
        self._cbr_bitrate_spin.setRange(48, 384)
        self._cbr_bitrate_spin.setSingleStep(8)
        self._cbr_bitrate_spin.setValue(64)
        self._cbr_bitrate_spin.setFixedWidth(88)
        encoding_row.addWidget(self._cbr_bitrate_spin)
        encoding_row.addStretch()

        settings_left.addLayout(encoding_row)

        channels_row = QHBoxLayout()

        channels_label = QLabel("Channels")
        channels_label.setFixedWidth(90)
        channels_row.addWidget(channels_label)

        self._mono_radio = QRadioButton("Mono")
        self._stereo_radio = QRadioButton("Stereo")
        self._channel_group = QButtonGroup(self)
        self._channel_group.addButton(self._mono_radio)
        self._channel_group.addButton(self._stereo_radio)
        self._mono_radio.setChecked(True)
        channels_row.addWidget(self._mono_radio)
        channels_row.addWidget(self._stereo_radio)
        channels_row.addStretch()

        settings_left.addLayout(channels_row)

        size_limit_row = QHBoxLayout()

        size_limit_label = QLabel("Size Limit")
        size_limit_label.setFixedWidth(90)
        size_limit_row.addWidget(size_limit_label)

        self._size_limit_edit = QLineEdit()
        self._size_limit_edit.setPlaceholderText("")
        self._size_limit_edit.setToolTip("e.g. 1,500,000,000 (blank or 0 = no limit)")
        size_limit_label.setToolTip("e.g. 1,500,000,000 (blank or 0 = no limit)")
        self._size_limit_edit.setText("1000000000")
        self._size_limit_edit.setFixedWidth(185)
        size_limit_row.addWidget(self._size_limit_edit)

        self._size_unit_group = QButtonGroup(self)
        for _unit_label in ("b", "kb", "Mb", "Gb", "B", "KB", "MB", "GB"):
            _rb = QRadioButton(_unit_label)
            self._size_unit_group.addButton(_rb)
            size_limit_row.addWidget(_rb)
            if _unit_label == "B":
                _rb.setChecked(True)

        size_limit_row.addStretch()
        settings_left.addLayout(size_limit_row)

        encode_row = QHBoxLayout()

        encode_label = QLabel("")
        encode_label.setFixedWidth(90)
        encode_row.addWidget(encode_label)

        self._encode_btn = QPushButton("Encode MP3")
        self._encode_btn.setEnabled(False)
        self._encode_btn.clicked.connect(self._on_encode_clicked)
        self._encode_btn.setStyleSheet(
            "QPushButton { background-color: #2e7d32; color: white; font-weight: bold; }"
            "QPushButton:hover { background-color: #388e3c; }"
            "QPushButton:disabled { background-color: #9e9e9e; color: #e0e0e0; }"
        )
        encode_row.addWidget(self._encode_btn)

        self._encode_progress = QProgressBar()
        self._encode_progress.setRange(0, 100)
        self._encode_progress.setValue(0)
        self._encode_progress.setTextVisible(True)
        self._encode_progress.setFormat("%p%")
        self._encode_progress.setSizePolicy(
            QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed
        )
        encode_row.addWidget(self._encode_progress)

        settings_left.addLayout(encode_row)
        settings_left.addStretch(1)
        top_settings_row.addLayout(settings_left, 3)

        preview_right = QVBoxLayout()
        preview_right.setContentsMargins(0, 0, 0, 0)
        preview_right.setSpacing(0)
        preview_right.setAlignment(Qt.AlignmentFlag.AlignTop | Qt.AlignmentFlag.AlignHCenter)

        preview_label = QLabel("Cover Preview")
        preview_label.setAlignment(Qt.AlignmentFlag.AlignHCenter)
        preview_right.addWidget(preview_label)

        self._cover_preview = QLabel("No cover selected")
        self._cover_preview.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self._cover_preview.setMinimumSize(220, 120)
        self._cover_preview.setSizePolicy(
            QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding
        )
        preview_right.addWidget(self._cover_preview, 1)

        top_settings_row.addLayout(preview_right, 1)
        top_panel = QWidget()
        top_panel_layout = QVBoxLayout(top_panel)
        top_panel_layout.setContentsMargins(0, 0, 0, 0)
        top_panel_layout.setSpacing(0)
        top_panel_layout.addLayout(top_settings_row)

        # ── Chapter table ─────────────────────────────────────────────────────
        self._table = ReorderableTableWidget()
        self._table.setColumnCount(5)
        self._table.setHorizontalHeaderLabels(["#", "Title", "Start", "End", "Cover"])
        self._table.setEditTriggers(QAbstractItemView.EditTrigger.DoubleClicked)
        self._table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self._table.setDragEnabled(True)
        self._table.setAcceptDrops(True)
        self._table.setDropIndicatorShown(True)
        self._table.setDragDropMode(QAbstractItemView.DragDropMode.DragDrop)
        self._table.setDefaultDropAction(Qt.DropAction.CopyAction)
        self._table.setDragDropOverwriteMode(False)
        self._table.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)
        self._table.setAlternatingRowColors(True)
        self._table.itemSelectionChanged.connect(self._on_table_selection_changed)
        self._table.itemChanged.connect(self._on_table_item_changed)
        self._table.rowMoveRequested.connect(self._on_table_row_move_requested)
        self._table.setItemDelegateForColumn(1, ChapterProgressDelegate(self._table))
        vh = self._table.verticalHeader()
        if vh is not None:
            vh.setVisible(False)

        hdr = self._table.horizontalHeader()
        if hdr is None:
            raise RuntimeError("Table has no horizontal header")
        hdr.setSectionResizeMode(0, QHeaderView.ResizeMode.ResizeToContents)
        hdr.setSectionResizeMode(1, QHeaderView.ResizeMode.Stretch)
        hdr.setSectionResizeMode(2, QHeaderView.ResizeMode.ResizeToContents)
        hdr.setSectionResizeMode(3, QHeaderView.ResizeMode.ResizeToContents)
        hdr.setSectionResizeMode(4, QHeaderView.ResizeMode.Fixed)
        self._table.setColumnWidth(4, THUMB_SIZE + 8)

        self._covers_btn = QPushButton("Process Covers")
        self._covers_btn.setEnabled(False)
        self._covers_btn.clicked.connect(self._on_process_covers)

        bottom_panel = QWidget()
        bottom_panel_layout = QVBoxLayout(bottom_panel)
        bottom_panel_layout.setContentsMargins(0, 0, 0, 0)
        bottom_panel_layout.setSpacing(0)
        bottom_panel_layout.addWidget(self._table)

        self._main_splitter = QSplitter(Qt.Orientation.Vertical)
        self._main_splitter.setChildrenCollapsible(False)
        self._main_splitter.setHandleWidth(8)
        self._main_splitter.addWidget(top_panel)
        self._main_splitter.addWidget(bottom_panel)
        self._main_splitter.setStretchFactor(0, 0)
        self._main_splitter.setStretchFactor(1, 1)
        self._main_splitter.setSizes(list(self._initial_splitter_sizes))
        self._main_splitter.splitterMoved.connect(self._on_main_splitter_moved)

        root.addWidget(self._main_splitter)

        # Capture the true initial splitter ratio after Qt applies final layout.
        QTimer.singleShot(0, self._capture_initial_splitter_ratio)

        # ── Status bar ────────────────────────────────────────────────────────
        self._status = QStatusBar()
        self.setStatusBar(self._status)
        self._status.addPermanentWidget(self._covers_btn)
        self._status.showMessage("Open a WAV file to begin.")
        self._on_encode_mode_changed()

    # ── Slots ─────────────────────────────────────────────────────────────────

    def _default_open_wav_dir(self) -> str:
        remembered_dir = self._settings.value("lastOpenWavDir", "")
        remembered_text = str(remembered_dir).strip() if remembered_dir is not None else ""
        if remembered_text:
            remembered_path = Path(remembered_text)
            if remembered_path.is_dir():
                return str(remembered_path)

        if self._wav_path is not None and self._wav_path.parent.is_dir():
            return str(self._wav_path.parent)

        return ""

    def _on_open_wav(self) -> None:
        start_dir = self._default_open_wav_dir()
        path_str, _ = QFileDialog.getOpenFileName(
            self, "Open WAV File", start_dir, "WAV Files (*.wav)"
        )
        if path_str:
            selected_parent = Path(path_str).parent
            if selected_parent.is_dir():
                self._settings.setValue("lastOpenWavDir", str(selected_parent))
            self._load_wav(Path(path_str))

    def _on_reset_clicked(self) -> None:
        self._wav_path = None
        self._output_dir = None
        self._chapters.clear()
        self._original_chapters.clear()
        self._missing_cover_rows.clear()
        self._row_cover_paths.clear()
        self._next_cover_action = "all"

        self._file_label.setText("No file selected")
        self._book_title_edit.blockSignals(True)
        self._book_title_edit.clear()
        self._book_title_edit.blockSignals(False)
        self._set_output_path(None)
        self._duration_edit.clear()

        self._table.clearContents()
        self._table.setRowCount(0)
        self._clear_table_encode_progress(reset_timeline=True)
        self._update_cover_preview(None)

        self._set_controls_enabled(True)
        self._save_as_btn.setEnabled(False)
        self._covers_btn.setEnabled(False)
        self._reset_markers_btn.setEnabled(False)

        self._vbr_radio.setChecked(True)
        self._vbr_q_combo.setCurrentIndex(5)
        self._cbr_bitrate_spin.setValue(64)
        self._mono_radio.setChecked(True)
        self._size_limit_edit.setText("1000000000")
        if self._size_unit_group is not None:
            for _btn in self._size_unit_group.buttons():
                if _btn.text() == "B":
                    _btn.setChecked(True)
                    break
        self._on_encode_mode_changed()

        self._set_encode_btn_state("encode")
        self._encode_btn.setEnabled(False)
        self._encode_progress.setValue(0)

        self._reset_main_splitter_to_initial()
        self._status.showMessage("Open a WAV file to begin.")

    def _on_reset_markers_clicked(self) -> None:
        if not self._chapters or not self._original_chapters:
            return

        reply = QMessageBox.question(
            self,
            "Reset Chapter Markers",
            "Reset chapter titles/markers in the list back to the original WAV values?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No,
        )
        if reply != QMessageBox.StandardButton.Yes:
            self._status.showMessage("Marker reset cancelled.")
            return

        self._chapters = list(self._original_chapters)
        self._encode_progress_timeline.clear()
        self._missing_cover_rows.clear()
        self._row_cover_paths.clear()
        self._next_cover_action = "all"
        self._populate_table(self._chapters)
        self._update_cover_preview(None)
        self._status.showMessage("Chapter list reset to original WAV markers.")

    def _load_wav(self, wav_path: Path) -> None:
        self._status.showMessage("Reading markers…")
        QApplication.processEvents()

        # Locate ffprobe
        ffprobe, err = _run_capturing_errors(ensure_executable, "ffprobe")
        if err:
            self._status.showMessage(f"ffprobe not found — {err}")
            return

        # Get audio duration via ffprobe
        result, err = _run_capturing_errors(probe_audio_info, ffprobe, wav_path)
        if err:
            self._status.showMessage(f"Cannot probe audio — {err}")
            return
        assert result is not None
        duration_sec, _ = result

        # Parse RIFF/WAV markers
        result, err = _run_capturing_errors(read_riff_chunks, wav_path)
        if err:
            self._status.showMessage(f"Cannot read markers — {err}")
            return
        assert result is not None
        sample_rate, markers = result

        # Build chapter list from markers
        chapters, err = _run_capturing_errors(
            build_chapters, markers, sample_rate, duration_sec
        )
        if err:
            self._status.showMessage(f"Cannot build chapters — {err}")
            return
        assert chapters is not None

        self._wav_path = wav_path
        self._output_dir = wav_path.parent
        self._chapters = chapters
        self._original_chapters = list(chapters)
        self._encode_progress_timeline.clear()
        self._missing_cover_rows.clear()
        self._row_cover_paths.clear()
        self._next_cover_action = "all"
        self._file_label.setText(str(wav_path))
        self._duration_edit.setText(_format_duration_label(duration_sec))
        self._set_book_title(wav_path.stem, auto_unique_output=False)
        initial_output_path = self._output_path_for_current_title()
        if initial_output_path is not None:
            resolved_output_path = self._resolve_output_conflict(
                initial_output_path,
                cancel_returns_none=False,
            )
            if resolved_output_path is not None:
                self._output_path_edit.setText(str(resolved_output_path))
        self._populate_table(chapters)
        self._update_cover_preview(None)
        self._covers_btn.setEnabled(True)
        self._save_as_btn.setEnabled(True)
        self._reset_markers_btn.setEnabled(True)
        self._encode_btn.setEnabled(True)
        self._encode_progress.setValue(0)

        count = len(chapters)
        noun = "chapter" if count == 1 else "chapters"
        self._status.showMessage(f"Loaded {count} {noun} — {wav_path.name}")

    def _set_book_title(self, title: str, auto_unique_output: bool = False) -> None:
        self._book_title_edit.blockSignals(True)
        self._book_title_edit.setText(title)
        self._book_title_edit.blockSignals(False)
        self._sync_output_path(auto_unique=auto_unique_output)

    def _current_book_title(self) -> str:
        return self._book_title_edit.text().strip()

    def _output_path_for_current_title(self) -> Path | None:
        if self._output_dir is None:
            return None
        stem = _sanitize_output_stem(self._current_book_title())
        return self._output_dir / f"{stem}.mp3"

    def _clear_output_conflict_memory(self) -> None:
        self._remembered_conflict_path = None
        self._remembered_conflict_choice = None

    def _set_output_path(self, output_path: Path | None) -> None:
        old_text = self._output_path_edit.text().strip()
        new_text = str(output_path) if output_path is not None else ""
        self._output_path_edit.setText(new_text)
        if new_text != old_text:
            self._clear_output_conflict_memory()

    def _build_unique_output_path(self, output_path: Path) -> Path:
        stem = output_path.stem
        parent = output_path.parent
        counter = 1
        while True:
            candidate = parent / f"{stem}({counter}).mp3"
            if not candidate.exists():
                return candidate
            counter += 1

    def _resolve_output_conflict(
        self,
        output_path: Path,
        *,
        cancel_returns_none: bool,
        use_remembered_choice: bool = False,
    ) -> Path | None:
        if not output_path.exists():
            return output_path

        if (
            use_remembered_choice
            and self._remembered_conflict_path == output_path
            and self._remembered_conflict_choice is not None
        ):
            if self._remembered_conflict_choice == "overwrite":
                return output_path
            if self._remembered_conflict_choice == "unique":
                return self._build_unique_output_path(output_path)
            if self._remembered_conflict_choice == "cancel":
                return None if cancel_returns_none else output_path

        msg_box = QMessageBox(self)
        msg_box.setWindowTitle("Overwrite Output")
        msg_box.setText(f"The output file already exists:\n{output_path}\n\nWhat would you like to do?")
        msg_box.setIcon(QMessageBox.Icon.Question)
        overwrite_btn = msg_box.addButton("Overwrite", QMessageBox.ButtonRole.AcceptRole)
        unique_btn = msg_box.addButton("Create Unique", QMessageBox.ButtonRole.ActionRole)
        msg_box.addButton("Cancel", QMessageBox.ButtonRole.RejectRole)
        msg_box.setDefaultButton(unique_btn)
        msg_box.exec()
        clicked = msg_box.clickedButton()

        if clicked is overwrite_btn:
            self._remembered_conflict_path = output_path
            self._remembered_conflict_choice = "overwrite"
            return output_path
        if clicked is unique_btn:
            self._remembered_conflict_path = output_path
            self._remembered_conflict_choice = "unique"
            return self._build_unique_output_path(output_path)
        self._remembered_conflict_path = output_path
        self._remembered_conflict_choice = "cancel"
        if cancel_returns_none:
            return None
        return output_path

    def _sync_output_path(self, auto_unique: bool = False) -> None:
        output_path = self._output_path_for_current_title()
        if output_path is not None and auto_unique and output_path.exists():
            stem = output_path.stem
            parent = output_path.parent
            counter = 1
            while True:
                candidate = parent / f"{stem}({counter}).mp3"
                if not candidate.exists():
                    break
                counter += 1
            output_path = candidate
        self._set_output_path(output_path)

    def _on_book_title_changed(self, _text: str) -> None:
        self._sync_output_path()

    def _on_encode_mode_changed(self) -> None:
        use_vbr = self._vbr_radio.isChecked()
        self._vbr_quality_label.setVisible(use_vbr)
        self._vbr_q_combo.setVisible(use_vbr)
        self._cbr_bitrate_label.setVisible(not use_vbr)
        self._cbr_bitrate_spin.setVisible(not use_vbr)

    def _get_unpopulated_artwork_rows(self) -> list[int]:
        return [row for row in range(len(self._chapters)) if row not in self._row_cover_paths]

    def _update_cover_preview(self, row: int | None) -> None:
        if row is None:
            self._preview_base_pixmap = None
            self._cover_preview.clear()
            self._cover_preview.setText("No cover selected")
            return

        cover_path = self._row_cover_paths.get(row)
        if cover_path is None:
            self._preview_base_pixmap = None
            self._cover_preview.clear()
            self._cover_preview.setText("No cover available for selected chapter")
            return

        px = QPixmap(str(cover_path))
        if px.isNull():
            self._preview_base_pixmap = None
            self._cover_preview.clear()
            self._cover_preview.setText("Could not load cover preview")
            return

        self._preview_base_pixmap = px
        self._refresh_cover_preview_pixmap()

    def _build_encode_progress_timeline(self, chapters: list[Chapter]) -> list[Chapter]:
        timeline: list[Chapter] = []
        cursor_ms = 0
        for row, chapter in enumerate(chapters, start=1):
            duration_ms = max(1, chapter.end_ms - chapter.start_ms)
            start_ms = cursor_ms
            end_ms = start_ms + duration_ms
            timeline.append(
                Chapter(
                    index=row,
                    title=chapter.title,
                    start_ms=start_ms,
                    end_ms=end_ms,
                )
            )
            cursor_ms = end_ms
        return timeline

    def _clear_table_encode_progress(self, reset_timeline: bool = False) -> None:
        for row in range(self._table.rowCount()):
            title_item = self._table.item(row, 1)
            if title_item is not None:
                title_item.setData(CHAPTER_PROGRESS_ROLE, 0.0)
        self._encode_progress_row = None
        if reset_timeline:
            self._encode_progress_timeline.clear()
        viewport = self._table.viewport()
        if viewport is not None:
            viewport.update()

    def _update_table_encode_progress(self, out_time_sec: float) -> None:
        timeline = self._encode_progress_timeline or self._chapters
        if out_time_sec < 0 or not timeline:
            self._clear_table_encode_progress()
            return

        current_ms = int(out_time_sec * 1000.0)
        for row, chapter in enumerate(timeline):
            title_item = self._table.item(row, 1)
            if title_item is None:
                continue

            is_last = row == len(self._chapters) - 1
            if current_ms < chapter.start_ms:
                row_progress = 0.0
            elif current_ms >= chapter.end_ms and not is_last:
                row_progress = 1.0
            else:
                chapter_span = max(1, chapter.end_ms - chapter.start_ms)
                row_progress = max(
                    0.0,
                    min(1.0, (current_ms - chapter.start_ms) / chapter_span),
                )
            title_item.setData(CHAPTER_PROGRESS_ROLE, row_progress)

        self._encode_progress_row = None
        viewport = self._table.viewport()
        if viewport is not None:
            viewport.update()

    def _refresh_cover_preview_pixmap(self) -> None:
        if self._preview_base_pixmap is None:
            return
        target_w = max(80, self._cover_preview.width() - 8)
        target_h = max(80, self._cover_preview.height() - 8)
        scaled = self._preview_base_pixmap.scaled(
            target_w,
            target_h,
            Qt.AspectRatioMode.KeepAspectRatio,
            Qt.TransformationMode.SmoothTransformation,
        )
        self._cover_preview.setText("")
        self._cover_preview.setPixmap(scaled)

    def resizeEvent(self, a0: QResizeEvent | None) -> None:
        super().resizeEvent(a0)
        if not self._splitter_reset_pending:
            self._splitter_reset_pending = True
            QTimer.singleShot(0, self._apply_pending_splitter_reset)

    def _on_main_splitter_moved(self, _pos: int, _index: int) -> None:
        self._refresh_cover_preview_pixmap()

    def _capture_initial_splitter_ratio(self) -> None:
        sizes = self._main_splitter.sizes()
        if len(sizes) != 2:
            return
        total = sizes[0] + sizes[1]
        if total <= 0:
            return
        self._initial_splitter_top_px = sizes[0]
        self._initial_splitter_ratio = sizes[0] / total

    def _apply_pending_splitter_reset(self) -> None:
        self._splitter_reset_pending = False
        self._reset_main_splitter_to_initial()

    def _reset_main_splitter_to_initial(self) -> None:
        top_default, bottom_default = self._initial_splitter_sizes
        total_default = top_default + bottom_default
        current_height = self._main_splitter.height()
        if current_height <= 0 or total_default <= 0:
            self._main_splitter.setSizes([top_default, bottom_default])
            return

        top_size = self._initial_splitter_top_px
        if top_size is None:
            top_ratio = self._initial_splitter_ratio
            if top_ratio is None:
                top_ratio = top_default / total_default
            top_size = int(current_height * top_ratio)
        top_size = max(140, min(current_height - 120, top_size))
        bottom_size = max(120, current_height - top_size)
        self._main_splitter.setSizes([top_size, bottom_size])
        self._refresh_cover_preview_pixmap()

    def changeEvent(self, a0: QEvent | None) -> None:
        super().changeEvent(a0)
        if a0 is None or a0.type() != QEvent.Type.WindowStateChange:
            return
        is_maximized = self.isMaximized()
        if self._was_maximized and not is_maximized:
            QTimer.singleShot(0, self._reset_main_splitter_to_initial)
        self._was_maximized = is_maximized

    def _on_table_selection_changed(self) -> None:
        selected_ranges = self._table.selectedRanges()
        if not selected_ranges:
            self._update_cover_preview(None)
            return
        self._update_cover_preview(selected_ranges[0].topRow())

    def _on_save_as(self) -> None:
        if self._output_dir is None:
            return

        default_path = self._output_path_for_current_title()
        start_path = str(default_path) if default_path else str(self._output_dir)
        path_str, _ = QFileDialog.getSaveFileName(
            self,
            "Save Output MP3 As",
            start_path,
            "MP3 Files (*.mp3)",
        )
        if not path_str:
            return

        selected_path = Path(path_str)
        if selected_path.suffix.lower() != ".mp3":
            selected_path = selected_path.with_suffix(".mp3")
        self._output_dir = selected_path.parent
        self._set_book_title(selected_path.stem, auto_unique_output=False)
        self._set_output_path(selected_path)

    def _set_controls_enabled(self, enabled: bool) -> None:
        self._open_btn.setEnabled(enabled)
        self._reset_btn.setEnabled(enabled)
        self._reset_markers_btn.setEnabled(enabled and bool(self._chapters))
        self._save_as_btn.setEnabled(enabled and self._wav_path is not None)
        self._covers_btn.setEnabled(enabled and bool(self._chapters))
        self._book_title_edit.setEnabled(enabled)
        self._vbr_radio.setEnabled(enabled)
        self._cbr_radio.setEnabled(enabled)
        self._vbr_q_combo.setEnabled(enabled and self._vbr_radio.isChecked())
        self._cbr_bitrate_spin.setEnabled(enabled and self._cbr_radio.isChecked())
        self._mono_radio.setEnabled(enabled)
        self._stereo_radio.setEnabled(enabled)
        self._size_limit_edit.setEnabled(enabled)
        if self._size_unit_group is not None:
            for _btn in self._size_unit_group.buttons():
                _btn.setEnabled(enabled)

    def _get_size_limit_value(self) -> tuple[Decimal, str] | None:
        text = self._size_limit_edit.text().replace(",", "").replace(" ", "").strip()
        if not text:
            return None
        try:
            value = Decimal(text)
        except InvalidOperation:
            return None
        if value <= Decimal("0"):
            return None
        if self._size_unit_group is None:
            return None
        checked = self._size_unit_group.checkedButton()
        if checked is None:
            return None
        return value, checked.text()

    def _actual_size_in_unit(self, actual_bytes: int, unit: str) -> Decimal:
        as_decimal = Decimal(actual_bytes)
        if unit == "b":
            return as_decimal * Decimal(8)
        if unit == "kb":
            return (as_decimal * Decimal(8)) / Decimal(1000)
        if unit == "Mb":
            return (as_decimal * Decimal(8)) / Decimal(1_000_000)
        if unit == "Gb":
            return (as_decimal * Decimal(8)) / Decimal(1_000_000_000)
        if unit == "B":
            return as_decimal
        if unit == "KB":
            return as_decimal / Decimal(1000)
        if unit == "MB":
            return as_decimal / Decimal(1_000_000)
        if unit == "GB":
            return as_decimal / Decimal(1_000_000_000)
        return as_decimal

    def _set_encode_btn_state(self, state: str) -> None:
        """state: 'encode' | 'cancel' | 'cancelling'"""
        if state == "encode":
            self._encode_btn.setText("Encode MP3")
            self._encode_btn.setStyleSheet(
                "QPushButton { background-color: #2e7d32; color: white; font-weight: bold; }"
                "QPushButton:hover { background-color: #388e3c; }"
                "QPushButton:disabled { background-color: #9e9e9e; color: #e0e0e0; }"
            )
        elif state == "cancel":
            self._encode_btn.setText("Cancel Encode")
            self._encode_btn.setStyleSheet(
                "QPushButton { background-color: #c62828; color: white; font-weight: bold; }"
                "QPushButton:hover { background-color: #d32f2f; }"
                "QPushButton:disabled { background-color: #9e9e9e; color: #e0e0e0; }"
            )
        else:  # cancelling
            self._encode_btn.setText("Cancelling\u2026")
            self._encode_btn.setStyleSheet(
                "QPushButton { background-color: #9e9e9e; color: #e0e0e0; font-weight: bold; }"
            )

    def _on_encode_clicked(self) -> None:
        if self._encode_thread is not None and self._encode_thread.isRunning():
            if self._encode_worker is not None:
                self._encode_worker.request_cancel()
            self._encode_btn.setEnabled(False)
            self._encode_btn.setText("Cancelling…")
            self._status.showMessage("Cancellation requested…")
            return

        if self._wav_path is None:
            self._status.showMessage("Open a WAV file first.")
            return

        output_text = self._output_path_edit.text().strip()
        if not output_text:
            self._status.showMessage("Output file path is empty.")
            return

        output_path = Path(output_text)
        if output_path.suffix.lower() != ".mp3":
            output_path = output_path.with_suffix(".mp3")
            self._set_output_path(output_path)

        unpopulated_rows = self._get_unpopulated_artwork_rows()
        if unpopulated_rows:
            msg_box = QMessageBox(self)
            msg_box.setWindowTitle("Missing Artwork")
            msg_box.setIcon(QMessageBox.Icon.Warning)
            msg_box.setText(
                f"Artwork is not populated for {len(unpopulated_rows)} chapter(s).\n\n"
                "How would you like to proceed?"
            )
            parse_btn = msg_box.addButton(
                "Parse Artwork Then Encode",
                QMessageBox.ButtonRole.AcceptRole,
            )
            encode_anyway_btn = msg_box.addButton(
                "Encode Anyway (No Artwork Handling)",
                QMessageBox.ButtonRole.ActionRole,
            )
            msg_box.addButton("Cancel Encode", QMessageBox.ButtonRole.RejectRole)
            msg_box.setDefaultButton(parse_btn)
            msg_box.exec()

            clicked = msg_box.clickedButton()
            if clicked is parse_btn:
                self._on_process_covers()
                remaining = self._get_unpopulated_artwork_rows()
                if remaining:
                    self._status.showMessage(
                        f"Encode cancelled: artwork still missing for {len(remaining)} chapter(s)."
                    )
                    return
            elif clicked is encode_anyway_btn:
                pass
            else:
                self._status.showMessage("Encode cancelled.")
                return

        resolved_output_path = self._resolve_output_conflict(
            output_path,
            cancel_returns_none=True,
            use_remembered_choice=True,
        )
        if resolved_output_path is None:
            self._status.showMessage("Encode cancelled.")
            return
        output_path = resolved_output_path
        self._set_output_path(output_path)

        channels = 2 if self._stereo_radio.isChecked() else 1
        vbr_quality = int(self._vbr_q_combo.currentData())
        cbr_bitrate = None if self._vbr_radio.isChecked() else f"{self._cbr_bitrate_spin.value()}k"

        self._set_controls_enabled(False)
        self._encode_btn.setEnabled(True)
        self._set_encode_btn_state("cancel")
        self._encode_progress.setValue(0)
        self._encode_progress_timeline = self._build_encode_progress_timeline(self._chapters)
        self._status.showMessage("Starting encode…")

        self._encode_thread = QThread(self)
        self._encode_worker = EncodeWorker(
            input_wav=self._wav_path,
            output_mp3=output_path,
            chapters=list(self._chapters),
            channels=channels,
            vbr_quality=vbr_quality,
            cbr_bitrate=cbr_bitrate,
        )
        self._encode_worker.moveToThread(self._encode_thread)
        self._encode_thread.started.connect(self._encode_worker.run)
        self._encode_worker.progress.connect(self._on_encode_progress)
        self._encode_worker.finished.connect(self._on_encode_finished)
        self._encode_worker.failed.connect(self._on_encode_failed)
        self._encode_worker.finished.connect(self._encode_thread.quit)
        self._encode_worker.failed.connect(self._encode_thread.quit)
        self._encode_thread.finished.connect(self._on_encode_thread_finished)
        self._encode_thread.start()

    def _on_encode_progress(self, percent: int, message: str, out_time_sec: float) -> None:
        self._encode_progress.setValue(percent)
        self._status.showMessage(message)
        self._update_table_encode_progress(out_time_sec)

    def _on_encode_finished(self, message: str, actual_bytes: int) -> None:
        self._encode_progress.setValue(100)
        self._status.showMessage(message)
        self._clear_table_encode_progress(reset_timeline=True)
        limit_config = self._get_size_limit_value()
        if limit_config is not None:
            limit_value, unit = limit_config
            actual_value = self._actual_size_in_unit(actual_bytes, unit)
            if actual_value <= limit_value:
                return
            excess_value = actual_value - limit_value
            QMessageBox.warning(
                self,
                "File Size Limit Exceeded",
                f"The output file exceeds the configured size limit.\n\n"
                f"File size:   {format_bytes(actual_bytes)}\n"
                f"Size limit:  {limit_value} {unit}\n"
                f"Exceeds by:  {excess_value} {unit}",
            )

    def _on_encode_failed(self, message: str) -> None:
        self._encode_progress.setValue(0)
        self._clear_table_encode_progress(reset_timeline=True)
        if "cancelled" in message.casefold():
            self._status.showMessage(message)
            return
        self._status.showMessage(f"Encode failed: {message}")
        QMessageBox.critical(self, "Encode Failed", message)

    def _on_encode_thread_finished(self) -> None:
        self._set_controls_enabled(True)
        self._set_encode_btn_state("encode")
        self._encode_btn.setEnabled(self._wav_path is not None)
        if self._encode_thread is not None:
            self._encode_thread.deleteLater()
        if self._encode_worker is not None:
            self._encode_worker.deleteLater()
        self._encode_thread = None
        self._encode_worker = None

    def _populate_table(self, chapters: list[Chapter]) -> None:
        self._is_populating_table = True
        self._table.blockSignals(True)
        self._table.setRowCount(0)
        self._table.setRowCount(len(chapters))

        for row, ch in enumerate(chapters):
            self._table.setRowHeight(row, THUMB_SIZE + 4)

            num = QTableWidgetItem(str(ch.index))
            num.setFlags(num.flags() & ~Qt.ItemFlag.ItemIsEditable)
            num.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            self._table.setItem(row, 0, num)

            title_item = QTableWidgetItem(ch.title)
            title_item.setFlags(title_item.flags() | Qt.ItemFlag.ItemIsEditable)
            title_item.setData(CHAPTER_KEY_ROLE, _chapter_key(ch))
            self._table.setItem(row, 1, title_item)

            start = QTableWidgetItem(format_seconds_hms(ch.start_ms / 1000.0))
            start.setFlags(start.flags() & ~Qt.ItemFlag.ItemIsEditable)
            start.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            self._table.setItem(row, 2, start)

            end = QTableWidgetItem(format_seconds_hms(ch.end_ms / 1000.0))
            end.setFlags(end.flags() & ~Qt.ItemFlag.ItemIsEditable)
            end.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            self._table.setItem(row, 3, end)

            cover_lbl = QLabel()
            cover_lbl.setAlignment(Qt.AlignmentFlag.AlignCenter)
            self._table.setCellWidget(row, 4, cover_lbl)

        self._table.blockSignals(False)
        self._is_populating_table = False

    def _on_table_item_changed(self, item: QTableWidgetItem) -> None:
        if self._is_populating_table:
            return
        row = item.row()
        col = item.column()
        if col != 1:
            return
        if row < 0 or row >= len(self._chapters):
            return

        new_title = item.text().strip()
        old_chapter = self._chapters[row]
        if not new_title:
            self._table.blockSignals(True)
            item.setText(old_chapter.title)
            self._table.blockSignals(False)
            return
        if new_title == old_chapter.title:
            return

        self._chapters[row] = replace(old_chapter, title=new_title)

    def _on_table_row_move_requested(self, source_row: int, target_row: int) -> None:
        if self._is_populating_table or not self._chapters:
            return

        row_count = len(self._chapters)
        if not (0 <= source_row < row_count):
            return
        if target_row < 0:
            target_row = 0
        if target_row > row_count:
            target_row = row_count
        if target_row == source_row or target_row == source_row + 1:
            return

        old_chapters = list(self._chapters)
        cover_by_key: dict[str, Path] = {}
        missing_by_key: set[str] = set()
        for old_row, old_ch in enumerate(old_chapters):
            key = _chapter_key(old_ch)
            if old_row in self._row_cover_paths:
                cover_by_key[key] = self._row_cover_paths[old_row]
            if old_row in self._missing_cover_rows:
                missing_by_key.add(key)

        moved = old_chapters.pop(source_row)
        insert_row = target_row
        if insert_row > source_row:
            insert_row -= 1
        old_chapters.insert(insert_row, moved)

        new_chapters: list[Chapter] = []
        new_cover_paths: dict[int, Path] = {}
        new_missing_rows: set[int] = set()
        for new_row, chapter in enumerate(old_chapters):
            old_key = _chapter_key(chapter)
            chapter = replace(chapter, index=new_row + 1)
            new_chapters.append(chapter)
            if old_key in cover_by_key:
                new_cover_paths[new_row] = cover_by_key[old_key]
            if old_key in missing_by_key:
                new_missing_rows.add(new_row)

        self._chapters = new_chapters
        self._row_cover_paths = new_cover_paths
        self._missing_cover_rows = new_missing_rows
        self._populate_table(self._chapters)

        for row, cover_path in self._row_cover_paths.items():
            lbl = self._table.cellWidget(row, 4)
            if not isinstance(lbl, QLabel):
                continue
            px = QPixmap(str(cover_path))
            if px.isNull():
                continue
            lbl.setPixmap(
                px.scaled(
                    THUMB_SIZE,
                    THUMB_SIZE,
                    Qt.AspectRatioMode.KeepAspectRatio,
                    Qt.TransformationMode.SmoothTransformation,
                )
            )

        self._table.selectRow(insert_row)
        self._update_cover_preview(insert_row)

    def _on_process_covers(self) -> None:
        if not self._wav_path or not self._chapters:
            return

        target_rows: list[int]
        mode_used = "all"

        if self._next_cover_action == "missing" and self._missing_cover_rows:
            target_rows = sorted(self._missing_cover_rows)
            mode_used = "missing"
        elif self._next_cover_action == "prompt_force_all":
            reply = QMessageBox.question(
                self,
                "Force Update Covers",
                "All chapters currently have artwork previews.\n"
                "Force update all chapters?",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                QMessageBox.StandardButton.Yes,
            )
            if reply == QMessageBox.StandardButton.No:
                self._status.showMessage("No changes made.")
                return
            target_rows = list(range(len(self._chapters)))
            mode_used = "all"
        else:
            target_rows = list(range(len(self._chapters)))
            mode_used = "all"

        self._status.showMessage("Resolving cover images…")
        QApplication.processEvents()

        covers_by_title, err = _run_capturing_errors(
            _find_covers_by_title, self._wav_path.parent
        )
        if err:
            self._status.showMessage(f"Cover error — {err}")
            return
        assert covers_by_title is not None

        loaded = 0
        if mode_used == "all":
            next_missing_rows: set[int] = set()
        else:
            next_missing_rows = set(self._missing_cover_rows)

        for row in target_rows:
            ch = self._chapters[row]
            book_title = _book_title_from_chapter(ch.title)
            lbl = self._table.cellWidget(row, 4)
            if isinstance(lbl, QLabel):
                lbl.clear()
            self._row_cover_paths.pop(row, None)

            if not book_title:
                next_missing_rows.add(row)
                continue

            cover_path = covers_by_title.get(book_title.casefold())
            if not cover_path:
                next_missing_rows.add(row)
                continue

            px = QPixmap(str(cover_path))
            if px.isNull():
                next_missing_rows.add(row)
                continue
            px = px.scaled(
                THUMB_SIZE,
                THUMB_SIZE,
                Qt.AspectRatioMode.KeepAspectRatio,
                Qt.TransformationMode.SmoothTransformation,
            )
            if isinstance(lbl, QLabel):
                lbl.setPixmap(px)
            self._row_cover_paths[row] = cover_path
            next_missing_rows.discard(row)
            loaded += 1

        selected_ranges = self._table.selectedRanges()
        if selected_ranges:
            self._update_cover_preview(selected_ranges[0].topRow())
        else:
            self._update_cover_preview(None)

        self._missing_cover_rows = next_missing_rows

        if self._missing_cover_rows:
            if mode_used == "missing":
                self._next_cover_action = "prompt_force_all"
            else:
                self._next_cover_action = "missing"
        else:
            self._next_cover_action = "prompt_force_all"

        noun = "chapter" if loaded == 1 else "chapters"
        if self._missing_cover_rows:
            missing_titles: set[str] = set()
            for row in sorted(self._missing_cover_rows):
                chapter_title = self._chapters[row].title
                missing_titles.add(_book_title_from_chapter(chapter_title) or chapter_title)
            ordered = sorted(missing_titles, key=str.casefold)
            preview = ", ".join(ordered[:3])
            extra = len(ordered) - 3
            if extra > 0:
                preview = f"{preview} (+{extra} more)"
            self._status.showMessage(
                f"Processed {len(target_rows)} rows; loaded {loaded} {noun}. "
                f"One or more covers are missing: {preview}"
            )
        else:
            self._status.showMessage(
                f"Processed {len(target_rows)} rows; loaded {loaded} {noun}."
            )


def _make_no_window_run_command():
    """Return a patched run_command that suppresses the console window on Windows.

    On Windows, subprocess.run() called from a GUI app with no console will
    briefly flash a cmd window and can deliver spurious SIGINT to the parent
    process.  CREATE_NO_WINDOW prevents both problems.
    """
    import wav_markers_to_mp3 as _m

    def _patched(cmd: list[str]) -> subprocess.CompletedProcess[str]:
        kwargs: dict[str, Any] = dict(text=True, capture_output=True)
        if sys.platform == "win32":
            kwargs["creationflags"] = subprocess.CREATE_NO_WINDOW
        return cast(subprocess.CompletedProcess[str], subprocess.run(cmd, **kwargs))

    _m.run_command = _patched  # type: ignore[assignment]


def main() -> None:
    if sys.platform == "win32":
        _make_no_window_run_command()

    app = QApplication(sys.argv)
    app.setApplicationName("AudioBook Slicer")
    window = MainWindow()
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
