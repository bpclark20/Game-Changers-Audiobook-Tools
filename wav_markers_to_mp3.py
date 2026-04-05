#!/usr/bin/env python3
"""Convert a marker-authored WAV audiobook into chaptered MP3 output.

This script reads RIFF cue markers (including Adobe Audition marker labels)
from a WAV file, creates ffmpeg-compatible chapter metadata, encodes the
audio as VBR MP3, and writes ID3 chapter artwork for matching book covers.
"""

from __future__ import annotations

import argparse
import importlib
import html
import json
import os
import re
import shutil
import struct
import subprocess
import sys
import tempfile
import time
from dataclasses import dataclass
from pathlib import Path
from typing import Callable, NoReturn


BYTES_PER_DECIMAL_GB = 1000**3
DEFAULT_AUTO_HEADROOM_PERCENT = 1.5
DEFAULT_VBR_QUALITY = 6
DEFAULT_CBR_BITRATE = "96k"


@dataclass(frozen=True)
class Marker:
    cue_id: int
    sample_offset: int
    label: str


@dataclass(frozen=True)
class Chapter:
    index: int
    title: str
    start_ms: int
    end_ms: int


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description=(
            "Convert a WAV audiobook (with Adobe Audition/RIFF markers) into an "
            "MP3 with chapter markers and per-chapter cover art."
        )
    )
    parser.add_argument("input", type=Path, help="Input WAV file")
    parser.add_argument(
        "-o",
        "--output",
        type=Path,
        help="Output path (default: <input_stem>.mp3)",
    )
    parser.add_argument(
        "--stereo",
        action="store_true",
        help="Use stereo output (default: mono)",
    )
    parser.add_argument(
        "--max-size-gb",
        type=float,
        default=1.0,
        help=(
            "Warn if final output exceeds this size in decimal GB "
            "(1 GB = 1,000,000,000 bytes; default: 1.0)"
        ),
    )
    parser.add_argument(
        "--auto-headroom-percent",
        type=float,
        default=DEFAULT_AUTO_HEADROOM_PERCENT,
        help=(
            "Safety margin, expressed as percent below --max-size-gb, used when "
            "warning about likely oversized output (default: 1.5)."
        ),
    )
    parser.add_argument(
        "--vbr-quality",
        type=int,
        default=DEFAULT_VBR_QUALITY,
        help=(
            "VBR quality level for libmp3lame (0=best/largest, 9=worst/smallest). "
            "Default: 6."
        ),
    )
    parser.add_argument(
        "--cbr-bitrate",
        default=None,
        help=(
            "Use CBR mode at the specified bitrate instead of VBR, for example "
            "64k, 96k, or 128k."
        ),
    )
    parser.add_argument(
        "--estimate",
        action="store_true",
        help="Validate inputs and print VBR sizing notes before encoding",
    )
    parser.add_argument(
        "--estimate-only",
        action="store_true",
        help="Validate inputs, print VBR sizing notes, and exit without encoding",
    )
    parser.add_argument(
        "--debug-markers",
        action="store_true",
        help=(
            "Print marker/chapter timing diagnostics before conversion "
            "(first and last few entries)"
        ),
    )
    parser.add_argument(
        "--ffmpeg",
        default="ffmpeg",
        help="ffmpeg executable name or full path (default: ffmpeg)",
    )
    parser.add_argument(
        "--ffprobe",
        default="ffprobe",
        help="ffprobe executable name or full path (default: ffprobe)",
    )
    parser.add_argument(
        "--overwrite",
        action="store_true",
        help="Overwrite output if it already exists",
    )
    return parser.parse_args()


def fail(message: str, code: int = 1) -> NoReturn:
    print(f"ERROR: {message}", file=sys.stderr)
    raise SystemExit(code)


def os_name() -> str:
    return "windows" if sys.platform.startswith("win") else "other"


def find_winget_ffmpeg_binary(executable_name: str) -> str | None:
    """Best-effort lookup for WinGet FFmpeg installs on Windows."""
    if os_name() != "windows":
        return None

    home = Path.home()
    packages_root = home / "AppData" / "Local" / "Microsoft" / "WinGet" / "Packages"
    if not packages_root.exists():
        return None

    exe_name = f"{executable_name}.exe"
    candidates = list(packages_root.glob(f"*FFmpeg*/*/bin/{exe_name}"))
    if not candidates:
        return None

    newest = max(candidates, key=lambda path: path.stat().st_mtime)
    return str(newest)


def ensure_executable(name_or_path: str) -> str:
    if Path(name_or_path).exists():
        return str(Path(name_or_path))

    resolved = shutil.which(name_or_path)
    if resolved:
        return resolved

    if name_or_path in {"ffmpeg", "ffprobe"}:
        winget_path = find_winget_ffmpeg_binary(name_or_path)
        if winget_path:
            return winget_path

    fail(
        f"Could not find executable '{name_or_path}'. Install ffmpeg/ffprobe "
        "and ensure they are on PATH, or pass explicit --ffmpeg/--ffprobe paths."
    )
    return ""


def run_command(cmd: list[str]) -> subprocess.CompletedProcess[str]:
    return subprocess.run(cmd, text=True, capture_output=True)


def parse_ffmpeg_timestamp_to_seconds(timestamp: str) -> float | None:
    match = re.fullmatch(r"(\d+):(\d+):(\d+(?:\.\d+)?)", timestamp.strip())
    if not match:
        return None
    hours = int(match.group(1))
    minutes = int(match.group(2))
    seconds = float(match.group(3))
    return hours * 3600 + minutes * 60 + seconds


def print_ffmpeg_progress(
    out_time_sec: float,
    duration_sec: float,
    speed_text: str,
    started_at: float,
) -> tuple[float, str]:
    progress_ratio = max(0.0, min(1.0, out_time_sec / duration_sec))
    percent = progress_ratio * 100.0

    speed_match = re.fullmatch(r"\s*(\d+(?:\.\d+)?)x\s*", speed_text)
    eta_text = "--:--:--.---"
    if speed_match:
        speed = float(speed_match.group(1))
        if speed > 0:
            remaining = max(0.0, duration_sec - out_time_sec)
            eta_text = format_seconds_hms(remaining / speed)
    elif out_time_sec > 0:
        elapsed = max(0.0, time.monotonic() - started_at)
        remaining = max(0.0, duration_sec - out_time_sec)
        eta_estimate = elapsed * (remaining / out_time_sec)
        eta_text = format_seconds_hms(eta_estimate)

    line = (
        f"{percent:6.2f}% "
        f"({format_seconds_hms(out_time_sec)} / {format_seconds_hms(duration_sec)}) "
        f"ETA {eta_text} "
        f"Speed {speed_text or '?x'}"
    )
    return percent, line


def run_ffmpeg_with_progress(
    cmd: list[str],
    duration_sec: float,
    progress_label: str | None = None,
    progress_callback: Callable[[float, str, float], None] | None = None,
    cancel_requested: Callable[[], bool] | None = None,
) -> subprocess.CompletedProcess[str]:
    progress_cmd = [*cmd]
    output_path = progress_cmd.pop()
    progress_cmd.extend(["-progress", "pipe:1", "-nostats", output_path])

    process = subprocess.Popen(
        progress_cmd,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        text=True,
        bufsize=1,
    )

    progress_data: dict[str, str] = {}
    stdout_lines: list[str] = []
    started_at = time.monotonic()
    cancelled = False
    last_bucket = -1

    assert process.stdout is not None
    try:
        for raw_line in process.stdout:
            if cancel_requested is not None and cancel_requested():
                cancelled = True
                print("\nCancelling encode - stopping ffmpeg...", flush=True)
                process.terminate()
                try:
                    process.wait(timeout=10)
                except subprocess.TimeoutExpired:
                    process.kill()
                    process.wait()
                break

            line = raw_line.strip()
            if not line:
                continue
            stdout_lines.append(raw_line)

            if "=" not in line:
                continue
            key, value = line.split("=", 1)
            progress_data[key] = value

            if key != "progress":
                continue

            out_time_sec: float | None = None
            if "out_time_us" in progress_data:
                try:
                    out_time_sec = float(progress_data["out_time_us"]) / 1_000_000.0
                except ValueError:
                    out_time_sec = None
            elif "out_time_ms" in progress_data:
                try:
                    out_time_sec = float(progress_data["out_time_ms"]) / 1_000_000.0
                except ValueError:
                    out_time_sec = None
            elif "out_time" in progress_data:
                out_time_sec = parse_ffmpeg_timestamp_to_seconds(progress_data["out_time"])

            if out_time_sec is None:
                continue

            speed_text = progress_data.get("speed", "")
            percent, status_line = print_ffmpeg_progress(
                out_time_sec, duration_sec, speed_text, started_at
            )
            if progress_callback is not None:
                progress_callback(percent, status_line, out_time_sec)
            if progress_label:
                bucket = int(percent // 5)
                progress_state = progress_data.get("progress", "")
                if bucket > last_bucket or progress_state == "end":
                    last_bucket = bucket
                    line = f"[{progress_label}] {status_line}"
                    print(line, flush=True)
            else:
                print(
                    "\r"
                    f"Progress: {status_line}",
                    end="",
                    flush=True,
                )
    except KeyboardInterrupt:
        cancelled = True
        print("\nCancelling encode — stopping ffmpeg...", flush=True)
        process.terminate()
        try:
            process.wait(timeout=10)
        except subprocess.TimeoutExpired:
            process.kill()
            process.wait()

    return_code = process.wait() if not cancelled else (process.returncode or 1)
    stderr_text = process.stderr.read() if process.stderr is not None else ""
    stdout_text = "".join(stdout_lines)

    if cancelled:
        raise KeyboardInterrupt

    if return_code == 0:
        if progress_label:
            line = f"[{progress_label}] 100.00% (complete)"
            print(line)
            if progress_callback is not None:
                progress_callback(100.0, "100.00% (complete)", duration_sec)
        else:
            print("\rProgress: 100.00% (complete)" + " " * 40)
    else:
        if not progress_label:
            print()

    return subprocess.CompletedProcess(
        args=progress_cmd,
        returncode=return_code,
        stdout=stdout_text,
        stderr=stderr_text,
    )


def parse_bitrate_to_bps(bitrate: str) -> int | None:
    match = re.fullmatch(r"\s*(\d+(?:\.\d+)?)\s*([kKmM]?)\s*", bitrate)
    if not match:
        return None

    value = float(match.group(1))
    suffix = match.group(2).lower()
    if suffix == "m":
        value *= 1_000_000
    elif suffix == "k" or suffix == "":
        value *= 1_000
    return int(value)


def format_bytes(size_bytes: float) -> str:
    units = ["B", "KB", "MB", "GB", "TB"]
    value = float(size_bytes)
    for unit in units:
        if value < 1024 or unit == units[-1]:
            return f"{value:.2f} {unit}"
        value /= 1024
    return f"{size_bytes:.2f} B"


def format_seconds_hms(seconds: float) -> str:
    total_ms = max(0, int(round(seconds * 1000)))
    hours = total_ms // 3_600_000
    rem = total_ms % 3_600_000
    minutes = rem // 60_000
    rem = rem % 60_000
    secs = rem // 1000
    ms = rem % 1000
    return f"{hours:02d}:{minutes:02d}:{secs:02d}.{ms:03d}"


def probe_audio_info(ffprobe_bin: str, input_file: Path) -> tuple[float, int]:
    cmd = [
        ffprobe_bin,
        "-v",
        "error",
        "-select_streams",
        "a:0",
        "-show_entries",
        "format=duration:stream=channels",
        "-of",
        "json",
        str(input_file),
    ]
    result = run_command(cmd)
    if result.returncode != 0:
        fail(
            "ffprobe failed while reading audio info.\n"
            f"stderr: {result.stderr.strip() or '(empty)'}"
        )

    try:
        data = json.loads(result.stdout)
        duration_text = data["format"]["duration"]
        duration = float(duration_text)
        stream_info = data.get("streams", [])
        channels = int(stream_info[0]["channels"]) if stream_info else 1
    except (KeyError, TypeError, ValueError, json.JSONDecodeError):
        fail("Could not parse audio info from ffprobe output.")

    if duration <= 0:
        fail("Input audio duration is zero or negative.")
    if channels <= 0:
        channels = 1
    return duration, channels


def parse_pmx_markers(payload: bytes) -> list[Marker]:
    """Parse Adobe _PMX (XMP) markers when present in WAV files.

    RIFF cue markers are limited to 32-bit sample offsets, which truncates marker
    coverage on very long files. Adobe stores full marker data in _PMX/XMP.
    """
    text = payload.decode("utf-8", errors="ignore")
    pattern = re.compile(
        r"<xmpDM:startTime>(\d+)</xmpDM:startTime>.*?"
        r"<xmpDM:name>(.*?)</xmpDM:name>",
        re.DOTALL,
    )
    markers: list[Marker] = []
    for idx, match in enumerate(pattern.finditer(text), start=1):
        sample_offset = int(match.group(1))
        label = html.unescape(match.group(2)).strip()
        if not label:
            continue
        markers.append(Marker(cue_id=idx, sample_offset=sample_offset, label=label))
    markers.sort(key=lambda m: (m.sample_offset, m.cue_id))
    return markers


def read_riff_chunks(wav_path: Path) -> tuple[int, list[Marker]]:
    with wav_path.open("rb") as f:
        header = f.read(12)
        if len(header) < 12:
            fail("Input is too small to be a valid WAV file.")

        riff_id, _, wave_id = struct.unpack("<4sI4s", header)
        if riff_id not in {b"RIFF", b"RF64"} or wave_id != b"WAVE":
            fail("Input file is not a RIFF/RF64 WAVE file.")
        is_rf64 = riff_id == b"RF64"

        sample_rate: int | None = None
        cue_offsets: dict[int, int] = {}
        cue_labels: dict[int, str] = {}
        pmx_markers: list[Marker] = []
        ds64_data_size: int | None = None
        ds64_chunk_sizes: dict[bytes, list[int]] = {}

        file_size = wav_path.stat().st_size
        while f.tell() + 8 <= file_size:
            chunk_header = f.read(8)
            if len(chunk_header) < 8:
                break

            chunk_id, chunk_size = struct.unpack("<4sI", chunk_header)
            data_start = f.tell()

            if chunk_id == b"ds64":
                payload = f.read(chunk_size)
                if len(payload) >= 28:
                    _, ds64_data_size, _, table_len = struct.unpack("<QQQI", payload[:28])
                    pos = 28
                    for _ in range(table_len):
                        if pos + 12 > len(payload):
                            break
                        table_chunk_id = payload[pos : pos + 4]
                        table_chunk_size = struct.unpack("<Q", payload[pos + 4 : pos + 12])[0]
                        ds64_chunk_sizes.setdefault(table_chunk_id, []).append(
                            table_chunk_size
                        )
                        pos += 12
            elif chunk_id == b"fmt ":
                payload = f.read(min(chunk_size, 32))
                if len(payload) >= 8:
                    sample_rate = struct.unpack("<I", payload[4:8])[0]
            elif chunk_id == b"cue ":
                payload = f.read(chunk_size)
                if len(payload) >= 4:
                    count = struct.unpack("<I", payload[:4])[0]
                    base = 4
                    for _ in range(count):
                        if base + 24 > len(payload):
                            break
                        cue_id = struct.unpack("<I", payload[base : base + 4])[0]
                        sample_offset = struct.unpack(
                            "<I", payload[base + 20 : base + 24]
                        )[0]
                        cue_offsets[cue_id] = sample_offset
                        base += 24
            elif chunk_id == b"LIST" and chunk_size >= 4:
                payload = f.read(chunk_size)
                if len(payload) >= 4 and payload[:4] == b"adtl":
                    pos = 4
                    while pos + 8 <= len(payload):
                        sub_id = payload[pos : pos + 4]
                        sub_size = struct.unpack("<I", payload[pos + 4 : pos + 8])[0]
                        sub_start = pos + 8
                        sub_end = sub_start + sub_size
                        if sub_end > len(payload):
                            break

                        if sub_id in {b"labl", b"note"} and sub_size >= 4:
                            cue_id = struct.unpack(
                                "<I", payload[sub_start : sub_start + 4]
                            )[0]
                            text_bytes = payload[sub_start + 4 : sub_end]
                            text = text_bytes.split(b"\x00", 1)[0]
                            label = text.decode("utf-8", errors="replace").strip()
                            if label:
                                cue_labels[cue_id] = label

                        pos = sub_end + (sub_size % 2)
            elif chunk_id == b"_PMX":
                payload = f.read(chunk_size)
                parsed = parse_pmx_markers(payload)
                if parsed:
                    pmx_markers = parsed

            # For RF64, 0xFFFFFFFF means the true size is stored in ds64.
            resolved_chunk_size = chunk_size
            if is_rf64 and chunk_size == 0xFFFFFFFF:
                if chunk_id == b"data" and ds64_data_size is not None:
                    resolved_chunk_size = ds64_data_size
                elif chunk_id in ds64_chunk_sizes and ds64_chunk_sizes[chunk_id]:
                    resolved_chunk_size = ds64_chunk_sizes[chunk_id].pop(0)
                else:
                    fail(
                        "RF64 chunk size placeholder encountered but no ds64 size entry "
                        f"was found for chunk '{chunk_id.decode('latin1', errors='replace')}'."
                    )

            f.seek(data_start + resolved_chunk_size, 0)
            if resolved_chunk_size % 2 == 1:
                f.seek(1, 1)

    if sample_rate is None or sample_rate <= 0:
        fail("Could not read WAV sample rate from fmt chunk.")

    markers = [
        Marker(cue_id=cue_id, sample_offset=sample, label=cue_labels.get(cue_id, ""))
        for cue_id, sample in cue_offsets.items()
    ]
    markers.sort(key=lambda m: (m.sample_offset, m.cue_id))

    if pmx_markers:
        # Prefer Adobe's extended marker set when available.
        return sample_rate, pmx_markers

    return sample_rate, markers


def build_chapters(markers: list[Marker], sample_rate: int, duration_sec: float) -> list[Chapter]:
    if not markers:
        fail(
            "No WAV cue markers found. Ensure the file includes Adobe Audition/RIFF markers."
        )

    duration_ms = max(1, int(round(duration_sec * 1000)))

    chapter_starts: list[tuple[int, str]] = []
    seen_starts: set[int] = set()
    for i, marker in enumerate(markers, start=1):
        start_ms = int(round((marker.sample_offset / sample_rate) * 1000))
        if start_ms >= duration_ms:
            continue
        if start_ms in seen_starts:
            continue
        seen_starts.add(start_ms)
        raw_title = marker.label.strip()
        if raw_title.lower().endswith(".wav"):
            raw_title = raw_title[:-4].rstrip()
        title = raw_title or f"Chapter {i:03d}"
        chapter_starts.append((start_ms, title))

    if not chapter_starts:
        fail("All markers appear outside the audio duration.")

    chapters: list[Chapter] = []
    for i, (start_ms, title) in enumerate(chapter_starts, start=1):
        if i < len(chapter_starts):
            end_ms = chapter_starts[i][0] - 1
        else:
            end_ms = duration_ms
        if end_ms <= start_ms:
            continue
        chapters.append(Chapter(index=i, title=title, start_ms=start_ms, end_ms=end_ms))

    if not chapters:
        fail("Could not build valid chapter ranges from markers.")

    return chapters


def escape_ffmetadata_text(value: str) -> str:
    escaped = value.replace("\\", "\\\\")
    escaped = escaped.replace(";", r"\;")
    escaped = escaped.replace("=", r"\=")
    escaped = escaped.replace("#", r"\#")
    return escaped.replace("\n", " ").strip()


def write_ffmetadata(path: Path, chapters: list[Chapter]) -> None:
    lines = [";FFMETADATA1"]
    for chapter in chapters:
        lines.extend(
            [
                "[CHAPTER]",
                "TIMEBASE=1/1000",
                f"START={chapter.start_ms}",
                f"END={chapter.end_ms}",
                f"title={escape_ffmetadata_text(chapter.title)}",
            ]
        )
    path.write_text("\n".join(lines) + "\n", encoding="utf-8")


def print_marker_debug(markers: list[Marker], chapters: list[Chapter], sample_rate: int) -> None:
    print("Marker debug:")
    print(f"  Sample rate: {sample_rate} Hz")
    print(f"  Total markers: {len(markers)}")
    print(f"  Total chapters: {len(chapters)}")

    marker_limit = 5
    print("  First markers:")
    for marker in markers[:marker_limit]:
        secs = marker.sample_offset / sample_rate
        title = marker.label or "(no label)"
        print(
            "    "
            f"sample={marker.sample_offset} "
            f"time={format_seconds_hms(secs)} "
            f"label={title}"
        )

    if len(markers) > marker_limit:
        print("  Last markers:")
        for marker in markers[-marker_limit:]:
            secs = marker.sample_offset / sample_rate
            title = marker.label or "(no label)"
            print(
                "    "
                f"sample={marker.sample_offset} "
                f"time={format_seconds_hms(secs)} "
                f"label={title}"
            )

    chapter_limit = 5
    print("  First chapters:")
    for chapter in chapters[:chapter_limit]:
        print(
            "    "
            f"#{chapter.index:03d} "
            f"start={format_seconds_hms(chapter.start_ms / 1000.0)} "
            f"end={format_seconds_hms(chapter.end_ms / 1000.0)} "
            f"title={chapter.title}"
        )

    if len(chapters) > chapter_limit:
        print("  Last chapters:")
        for chapter in chapters[-chapter_limit:]:
            print(
                "    "
                f"#{chapter.index:03d} "
                f"start={format_seconds_hms(chapter.start_ms / 1000.0)} "
                f"end={format_seconds_hms(chapter.end_ms / 1000.0)} "
                f"title={chapter.title}"
            )


def make_unique_output_path(path: Path) -> Path:
    """Return a path that doesn't exist by appending (2), (3), ... to the stem."""
    if not path.exists():
        return path
    stem = path.stem
    suffix = path.suffix
    parent = path.parent
    counter = 2
    while True:
        candidate = parent / f"{stem} ({counter}){suffix}"
        if not candidate.exists():
            return candidate
        counter += 1


def resolve_output_path(path: Path, overwrite: bool) -> Path:
    """If path exists, prompt the user for overwrite/unique/bail unless --overwrite."""
    if not path.exists():
        return path
    if overwrite:
        print(f"Output already exists, overwriting (--overwrite): {path}")
        return path
    print(f"\nOutput file already exists: {path}")
    while True:
        print("  [O] Overwrite the existing file")
        print("  [U] Save with a new unique filename")
        print(">>> [B] Bail out (cancel, default)")
        try:
            choice = input("Choice [O/U/B]: ").strip().lower()
        except (EOFError, KeyboardInterrupt):
            print()
            raise SystemExit(0)
        if choice == "":
            print("Cancelled.")
            raise SystemExit(0)
        if choice == "o":
            return path
        if choice in {"u", "n"}:
            new_path = make_unique_output_path(path)
            print(f"New output path: {new_path}")
            return new_path
        if choice == "b":
            print("Cancelled.")
            raise SystemExit(0)
        print("Please enter O, U, or B.")


def build_ffmpeg_cmd(
    ffmpeg_bin: str,
    input_wav: Path,
    ffmetadata_file: Path,
    output_audio: Path,
    channels: int,
    vbr_quality: int,
    cbr_bitrate: str | None = None,
    sample_rate: int | None = None,
) -> list[str]:
    cmd = [
        ffmpeg_bin,
        "-hide_banner",
        "-loglevel",
        "error",
        "-nostdin",
        "-threads",
        "0",
        "-y",  # always overwrite; we encode to a temp file
        "-i",
        str(input_wav),
        "-i",
        str(ffmetadata_file),
    ]

    audio_flags = [
        "-map",
        "0:a:0",
        "-map_metadata",
        "1",
        "-map_chapters",
        "1",
        "-c:a",
        "libmp3lame",
        "-ac",
        str(channels),
    ]
    if cbr_bitrate is not None:
        audio_flags.extend(["-b:a", cbr_bitrate])
    else:
        audio_flags.extend(["-q:a", str(vbr_quality)])
    if sample_rate is not None:
        audio_flags.extend(["-ar", str(sample_rate)])
    audio_flags.append("-vn")
    cmd.extend(audio_flags)

    # -joint_stereo 0: disable joint stereo (use independent stereo or plain mono)
    # libmp3lame writes a Xing/LAME VBR info header automatically, so no MLLT needed
    cmd.extend(["-joint_stereo", "0", "-id3v2_version", "3", str(output_audio)])

    return cmd


def cover_mime_type(path: Path) -> str:
    suffix = path.suffix.lower()
    if suffix in {".jpg", ".jpeg"}:
        return "image/jpeg"
    if suffix == ".png":
        return "image/png"
    fail(f"Unsupported cover image format: {path}")
    return "application/octet-stream"


def write_mp3_chapters_with_artwork(
    output_file: Path,
    chapters: list[Chapter],
    cover_map: dict[str, Path],
) -> tuple[int, int, int]:
    try:
        id3_module = importlib.import_module("mutagen.id3")
    except ImportError:
        fail(
            "Chapter artwork mode requires the 'mutagen' package. "
            f"Active interpreter: {sys.executable}. "
            "Install it with this interpreter: python -m pip install mutagen"
        )

    APIC = getattr(id3_module, "APIC")
    CHAP = getattr(id3_module, "CHAP")
    CTOC = getattr(id3_module, "CTOC")
    CTOCFlags = getattr(id3_module, "CTOCFlags")
    ID3 = getattr(id3_module, "ID3")
    TIT2 = getattr(id3_module, "TIT2")
    PictureType = getattr(id3_module, "PictureType")

    tags = ID3(str(output_file))
    tags.delall("CHAP")
    tags.delall("CTOC")

    cover_data_cache: dict[Path, bytes] = {}
    chapter_ids: list[str] = []

    for chapter_index, chapter in enumerate(chapters, start=1):
        chapter_id = f"chp{chapter_index:04d}"
        chapter_ids.append(chapter_id)

        book_title = split_book_title(chapter.title)
        cover_path = cover_map[book_title]
        if cover_path not in cover_data_cache:
            cover_data_cache[cover_path] = cover_path.read_bytes()

        chap_subframes = [
            TIT2(encoding=3, text=[chapter.title]),
            APIC(
                encoding=3,
                mime=cover_mime_type(cover_path),
                type=PictureType.COVER_FRONT,
                desc=f"chapter-{chapter_index:04d}",
                data=cover_data_cache[cover_path],
            ),
        ]

        tags.add(
            CHAP(
                element_id=chapter_id,
                start_time=chapter.start_ms,
                end_time=chapter.end_ms,
                start_offset=0xFFFFFFFF,
                end_offset=0xFFFFFFFF,
                sub_frames=chap_subframes,
            )
        )

    tags.add(
        CTOC(
            element_id="toc",
            flags=CTOCFlags.TOP_LEVEL | CTOCFlags.ORDERED,
            child_element_ids=chapter_ids,
            sub_frames=[TIT2(encoding=3, text=["Table of Contents"])],
        )
    )
    tags.save(str(output_file), v2_version=3)
    return len(chapter_ids), 1, len(chapter_ids)


def validate_args(args: argparse.Namespace) -> None:
    if args.max_size_gb <= 0:
        fail("--max-size-gb must be greater than 0.")
    if args.auto_headroom_percent < 0 or args.auto_headroom_percent >= 100:
        fail("--auto-headroom-percent must be >= 0 and < 100.")

    if args.vbr_quality < 0 or args.vbr_quality > 9:
        fail("--vbr-quality must be between 0 and 9.")

    if args.cbr_bitrate is not None:
        bitrate_bps = parse_bitrate_to_bps(args.cbr_bitrate)
        if bitrate_bps is None or bitrate_bps <= 0:
            fail(
                "--cbr-bitrate must be a positive bitrate such as 64k, 96k, or 128k."
            )

    if args.estimate_only and args.overwrite:
        fail("Invalid flags: --overwrite cannot be combined with --estimate-only.")


def split_book_title(chapter_title: str) -> str:
    parts = chapter_title.split(" - ", 1)
    if len(parts) != 2 or not parts[0].strip() or not parts[1].strip():
        fail(
            "Chapter titles must use the format 'Book Title - *'. "
            f"Found chapter title without expected prefix: {chapter_title!r}"
        )
    return parts[0].strip()


def resolve_cover_images(input_dir: Path, chapters: list[Chapter]) -> dict[str, Path]:
    expected_titles: list[str] = []
    for chapter in chapters:
        book_title = split_book_title(chapter.title)
        if book_title not in expected_titles:
            expected_titles.append(book_title)

    image_files = [
        child
        for child in input_dir.iterdir()
        if child.is_file() and child.suffix.lower() in {".jpg", ".jpeg", ".png"}
    ]
    images_by_stem: dict[str, list[Path]] = {}
    for image in image_files:
        images_by_stem.setdefault(image.stem.casefold(), []).append(image)

    preferred_ext = {".jpg": 0, ".jpeg": 1, ".png": 2}
    cover_map: dict[str, Path] = {}
    missing: list[str] = []
    for title in expected_titles:
        candidates = images_by_stem.get(title.casefold(), [])
        if not candidates:
            missing.append(title)
            continue
        candidates.sort(key=lambda p: preferred_ext.get(p.suffix.lower(), 99))
        cover_map[title] = candidates[0]

    if missing:
        expected_names = ", ".join(f"{title}.jpg/.png" for title in missing)
        fail(
            "One or more cover files are missing in the input folder. "
            f"Missing: {expected_names}"
        )

    return cover_map


def encode_output(
    ffmpeg_bin: str,
    input_file: Path,
    chapters: list[Chapter],
    duration_sec: float,
    output_file: Path,
    channels: int,
    vbr_quality: int,
    cbr_bitrate: str | None,
    cover_map: dict[str, Path],
    progress_label: str | None = None,
    progress_callback: Callable[[float, str, float], None] | None = None,
    cancel_requested: Callable[[], bool] | None = None,
    sample_rate: int | None = None,
) -> int:
    output_file.parent.mkdir(parents=True, exist_ok=True)
    temp_fd, temp_path_str = tempfile.mkstemp(
        suffix=".mp3",
        prefix="wav2audio_tmp_",
        dir=output_file.parent,
    )
    os.close(temp_fd)
    temp_output = Path(temp_path_str)

    try:
        with tempfile.TemporaryDirectory(prefix="wav2mp3_meta_") as temp_dir:
            meta_file = Path(temp_dir) / "chapters.ffmeta"
            write_ffmetadata(meta_file, chapters)

            cmd = build_ffmpeg_cmd(
                ffmpeg_bin=ffmpeg_bin,
                input_wav=input_file,
                ffmetadata_file=meta_file,
                output_audio=temp_output,
                channels=channels,
                vbr_quality=vbr_quality,
                cbr_bitrate=cbr_bitrate,
                sample_rate=sample_rate,
            )

            run_line = "Running ffmpeg conversion... (press Ctrl+C to cancel)"
            if progress_label:
                prefixed = f"[{progress_label}] {run_line}"
                print(prefixed)
            else:
                print(run_line)

            result = run_ffmpeg_with_progress(
                cmd,
                duration_sec,
                progress_label=progress_label,
                progress_callback=progress_callback,
                cancel_requested=cancel_requested,
            )
            if result.returncode != 0:
                temp_output.unlink(missing_ok=True)
                fail(
                    "ffmpeg conversion failed.\n"
                    f"stderr: {result.stderr.strip() or '(empty)'}"
                )

        if not temp_output.exists() or temp_output.stat().st_size == 0:
            temp_output.unlink(missing_ok=True)
            fail("Conversion finished but output file was not created.")

        os.replace(temp_output, output_file)
    except KeyboardInterrupt:
        temp_output.unlink(missing_ok=True)
        print("Encode cancelled. Partial output file removed.")
        raise SystemExit(1)
    except SystemExit:
        temp_output.unlink(missing_ok=True)
        raise
    except Exception:
        temp_output.unlink(missing_ok=True)
        raise

    actual_size_bytes = output_file.stat().st_size

    chap_count, ctoc_count, apic_count = write_mp3_chapters_with_artwork(
        output_file,
        chapters,
        cover_map,
    )

    if progress_label:
        created_line = f"[{progress_label}] Created: {output_file}"
        size_line = f"[{progress_label}] Actual size: {format_bytes(actual_size_bytes)}"
        print(created_line)
        print(size_line)
        print(
            f"[{progress_label}] ID3 chapter tags written: "
            f"CHAP={chap_count}, CTOC={ctoc_count}, APIC(in CHAP)={apic_count}"
        )
    else:
        print(f"Created: {output_file}")
        print(f"Actual size: {format_bytes(actual_size_bytes)}")
        print(
            "ID3 chapter tags written: "
            f"CHAP={chap_count}, CTOC={ctoc_count}, APIC(in CHAP)={apic_count}"
        )
    return actual_size_bytes


def main() -> None:
    args = parse_args()
    validate_args(args)

    input_file = args.input.expanduser().resolve()
    if not input_file.exists():
        fail(f"Input file not found: {input_file}")
    if not input_file.is_file():
        fail(f"Input path is not a file: {input_file}")
    if input_file.suffix.lower() != ".wav":
        fail("Input must be a .wav file.")

    output_extension = ".mp3"

    output_file = (
        args.output.expanduser().resolve()
        if args.output
        else input_file.with_suffix(output_extension)
    )
    if output_file.suffix.lower() != output_extension:
        output_file = output_file.with_suffix(output_extension)

    ffmpeg_bin = ensure_executable(args.ffmpeg)
    ffprobe_bin = ensure_executable(args.ffprobe)

    print(f"Reading WAV markers from: {input_file}")
    sample_rate, markers = read_riff_chunks(input_file)
    duration_sec, _ = probe_audio_info(ffprobe_bin, input_file)
    chapters = build_chapters(markers, sample_rate, duration_sec)

    max_bytes = args.max_size_gb * BYTES_PER_DECIMAL_GB
    auto_headroom_percent = args.auto_headroom_percent
    effective_budget_bytes = max_bytes * (1.0 - (auto_headroom_percent / 100.0))
    cover_map = resolve_cover_images(input_file.parent, chapters)

    print(f"Found {len(markers)} marker(s), using {len(chapters)} chapter(s).")
    print(f"Duration: {duration_sec:.2f} sec")
    print("ffmpeg threads: auto (0)")

    if args.debug_markers:
        print_marker_debug(markers, chapters, sample_rate)

    channels = 2 if args.stereo else 1
    vbr_quality = args.vbr_quality
    cbr_bitrate = args.cbr_bitrate

    print(f"Output: {output_file}")
    channel_label = "stereo" if channels == 2 else "mono"
    print(f"Channels: {channel_label} ({channels})")
    if cbr_bitrate is None:
        print(
            f"Encoding mode: VBR -q:a {vbr_quality} "
            "(libmp3lame; Xing/LAME seek header written automatically)"
        )
    else:
        print(f"Encoding mode: CBR -b:a {cbr_bitrate} (libmp3lame)")
    print(f"Covers: enabled ({len(cover_map)} matched title image(s))")
    print(
        "Size warning threshold: "
        f"{format_bytes(effective_budget_bytes)} "
        f"({auto_headroom_percent:.1f}% below {args.max_size_gb:.2f} GB)"
    )

    if args.estimate or args.estimate_only:
        print("NOTE: Pre-encode size estimation is not available in always-VBR mode.")
    if args.estimate_only:
        print("Estimate-only mode enabled; skipping conversion.")
        return

    output_file = resolve_output_path(output_file, args.overwrite)
    actual_size_bytes = encode_output(
        ffmpeg_bin=ffmpeg_bin,
        input_file=input_file,
        chapters=chapters,
        duration_sec=duration_sec,
        output_file=output_file,
        channels=channels,
        cover_map=cover_map,
        vbr_quality=vbr_quality,
        cbr_bitrate=cbr_bitrate,
        sample_rate=sample_rate,
    )

    if actual_size_bytes > effective_budget_bytes:
        print(
            "WARNING: Output file exceeds the effective size budget "
            f"({format_bytes(effective_budget_bytes)} with {auto_headroom_percent:.1f}% headroom).",
            file=sys.stderr,
        )
    if actual_size_bytes > max_bytes:
        print(
            "WARNING: Output file exceeds configured max size "
            f"({args.max_size_gb:.2f} GB decimal).",
            file=sys.stderr,
        )

    print("Done.")


if __name__ == "__main__":
    main()
