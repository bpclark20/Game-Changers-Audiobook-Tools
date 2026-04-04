#!/usr/bin/env python3
"""Split a DRM-free .m4b audiobook into per-chapter WAV files.

This utility uses ffprobe to read chapter boundaries and audio stream info,
then decodes each chapter to a lossless PCM WAV at the source sample rate
and bit depth. No audio quality is lost beyond what is already present in
the lossy source codec (typically AAC).
"""

from __future__ import annotations

import argparse
import json
import re
import shutil
import subprocess
import sys
from dataclasses import dataclass
from pathlib import Path
from typing import Any, NoReturn, cast


@dataclass(frozen=True)
class Chapter:
    index: int
    title: str
    start: float
    end: float

    @property
    def duration(self) -> float:
        return max(0.0, self.end - self.start)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description=(
            "Split a DRM-free .m4b file into one WAV file per chapter using ffmpeg. "
            "Decodes each chapter to PCM at the source sample rate and bit depth."
        )
    )
    parser.add_argument("input", type=Path, help="Path to input .m4b/.m4a audiobook")
    parser.add_argument(
        "-o",
        "--output-dir",
        type=Path,
        help="Output directory (default: <input_stem>_chapters)",
    )
    parser.add_argument(
        "--audio-stream",
        type=int,
        default=0,
        help="Audio stream index to extract (default: 0)",
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
        help="Overwrite output files if they exist",
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="Print planned ffmpeg commands without executing",
    )
    return parser.parse_args()


def fail(message: str, code: int = 1) -> NoReturn:
    print(f"ERROR: {message}", file=sys.stderr)
    raise SystemExit(code)


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

    # Pick the most recently modified candidate in case multiple versions exist.
    newest = max(candidates, key=lambda path: path.stat().st_mtime)
    return str(newest)


def os_name() -> str:
    return "windows" if sys.platform.startswith("win") else "other"


def ensure_executable(name_or_path: str) -> str:
    if Path(name_or_path).exists():
        return str(Path(name_or_path))

    resolved = shutil.which(name_or_path)
    if resolved:
        return resolved

    # VS Code terminals sometimes keep an old PATH. Auto-discover common WinGet install.
    if name_or_path in {"ffmpeg", "ffprobe"}:
        winget_path = find_winget_ffmpeg_binary(name_or_path)
        if winget_path:
            return winget_path

    fail(
        f"Could not find executable '{name_or_path}'. Install ffmpeg/ffprobe "
        "and ensure they are available on PATH, or pass explicit --ffmpeg/--ffprobe paths."
    )
    return resolved


def run_command(cmd: list[str]) -> subprocess.CompletedProcess[str]:
    return subprocess.run(cmd, text=True, capture_output=True)


def probe_chapters(ffprobe_bin: str, input_file: Path) -> list[Chapter]:
    cmd = [
        ffprobe_bin,
        "-v",
        "error",
        "-print_format",
        "json",
        "-show_chapters",
        str(input_file),
    ]
    result = run_command(cmd)
    if result.returncode != 0:
        fail(
            "ffprobe failed while reading chapters.\n"
            f"Command: {' '.join(cmd)}\n"
            f"stderr: {result.stderr.strip() or '(empty)'}"
        )

    try:
        data: dict[str, Any] = json.loads(result.stdout)
    except json.JSONDecodeError as exc:
        fail(f"Failed to parse ffprobe JSON output: {exc}")

    raw_chapters_obj = data.get("chapters", [])
    raw_chapters: list[dict[str, Any]] = (
        cast(list[dict[str, Any]], raw_chapters_obj)
        if isinstance(raw_chapters_obj, list)
        else []
    )
    chapters: list[Chapter] = []

    for i, chapter in enumerate(raw_chapters, start=1):
        start_obj = chapter.get("start_time")
        end_obj = chapter.get("end_time")
        if start_obj is None or end_obj is None:
            continue

        try:
            start = float(start_obj)
            end = float(end_obj)
        except (KeyError, TypeError, ValueError):
            continue

        tags_obj = chapter.get("tags")
        tags: dict[str, Any] = (
            cast(dict[str, Any], tags_obj) if isinstance(tags_obj, dict) else {}
        )
        title = str(tags.get("title") or f"Chapter {i:03d}").strip()
        if not title:
            title = f"Chapter {i:03d}"

        ch = Chapter(index=i, title=title, start=start, end=end)
        if ch.duration > 0:
            chapters.append(ch)

    return chapters


def sanitize_filename(name: str) -> str:
    # Replace path separators and invalid Windows filename characters.
    cleaned = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", name)
    cleaned = re.sub(r"\s+", " ", cleaned).strip(" .")
    return cleaned or "untitled"


def unique_path(path: Path, used: set[Path], overwrite: bool = False) -> Path:
    candidate = path
    counter = 2
    while candidate in used or (not overwrite and candidate.exists()):
        candidate = path.with_name(f"{path.stem} ({counter}){path.suffix}")
        counter += 1
    used.add(candidate)
    return candidate


def format_seconds(value: float) -> str:
    # Millisecond precision is typically enough for chapter boundary extraction.
    return f"{value:.3f}"


@dataclass(frozen=True)
class AudioFormat:
    sample_rate: int
    pcm_codec: str  # e.g. "pcm_s16le"


def _bit_depth_to_pcm_codec(bits: int) -> str:
    """Map a bit depth to the appropriate signed little-endian PCM WAV codec.

    Lossy sources (AAC) report 0 bits; we default to 16-bit which is standard.
    """
    if bits == 24:
        return "pcm_s24le"
    if bits == 32:
        return "pcm_s32le"
    return "pcm_s16le"


def probe_audio_format(
    ffprobe_bin: str, input_file: Path, stream_index: int
) -> AudioFormat:
    cmd = [
        ffprobe_bin,
        "-v",
        "error",
        "-select_streams",
        f"a:{stream_index}",
        "-print_format",
        "json",
        "-show_streams",
        str(input_file),
    ]
    result = run_command(cmd)
    if result.returncode != 0:
        fail(
            "ffprobe failed while reading audio stream info.\n"
            f"stderr: {result.stderr.strip() or '(empty)'}"
        )

    try:
        data: dict[str, Any] = json.loads(result.stdout)
    except json.JSONDecodeError as exc:
        fail(f"Failed to parse ffprobe JSON output: {exc}")

    streams = cast(list[dict[str, Any]], data.get("streams", []))
    if not streams:
        fail(f"No audio stream at index {stream_index} found in '{input_file.name}'.")

    stream = streams[0]

    try:
        sample_rate = int(stream.get("sample_rate") or 44100)
    except (TypeError, ValueError):
        sample_rate = 44100

    bits = 0
    for key in ("bits_per_raw_sample", "bits_per_sample"):
        try:
            candidate = int(stream.get(key) or 0)
            if candidate > 0:
                bits = candidate
                break
        except (TypeError, ValueError):
            continue

    return AudioFormat(sample_rate=sample_rate, pcm_codec=_bit_depth_to_pcm_codec(bits))


def build_ffmpeg_wav_cmd(
    ffmpeg_bin: str,
    input_file: Path,
    output_file: Path,
    chapter: Chapter,
    audio_format: AudioFormat,
    audio_stream_index: int,
    overwrite: bool,
) -> list[str]:
    return [
        ffmpeg_bin,
        "-hide_banner",
        "-loglevel",
        "error",
        "-nostdin",
        "-y" if overwrite else "-n",
        "-ss",
        format_seconds(chapter.start),
        "-t",
        format_seconds(chapter.duration),
        "-i",
        str(input_file),
        "-map",
        f"0:a:{audio_stream_index}",
        "-c:a",
        audio_format.pcm_codec,
        "-ar",
        str(audio_format.sample_rate),
        str(output_file),
    ]


def main() -> None:
    args = parse_args()

    input_file = args.input.expanduser().resolve()
    if not input_file.exists():
        fail(f"Input file not found: {input_file}")
    if not input_file.is_file():
        fail(f"Input path is not a file: {input_file}")

    output_dir = (
        args.output_dir.expanduser().resolve()
        if args.output_dir
        else Path.cwd() / f"{input_file.stem}_chapters"
    )
    output_dir.mkdir(parents=True, exist_ok=True)

    ffmpeg_bin = ensure_executable(args.ffmpeg)
    ffprobe_bin = ensure_executable(args.ffprobe)

    audio_format = probe_audio_format(ffprobe_bin, input_file, args.audio_stream)
    chapters = probe_chapters(ffprobe_bin, input_file)
    if not chapters:
        fail(
            "No chapters found in input file. Ensure the .m4b actually contains chapter metadata."
        )

    print(f"Found {len(chapters)} chapters in: {input_file}")
    print(
        f"Audio: {audio_format.sample_rate} Hz / {audio_format.pcm_codec} "
        f"-> output as WAV"
    )
    print(f"Output directory: {output_dir}")

    used_paths: set[Path] = set()
    created = 0

    for chapter in chapters:
        safe_title = sanitize_filename(chapter.title)
        out_name = f"{chapter.index:03d} - {safe_title}.wav"
        output_file = unique_path(output_dir / out_name, used_paths, overwrite=args.overwrite)

        wav_cmd = build_ffmpeg_wav_cmd(
            ffmpeg_bin=ffmpeg_bin,
            input_file=input_file,
            output_file=output_file,
            chapter=chapter,
            audio_format=audio_format,
            audio_stream_index=args.audio_stream,
            overwrite=args.overwrite,
        )

        if args.dry_run:
            print("[DRY RUN]", " ".join(wav_cmd))
            continue

        print(
            f"[{chapter.index:03d}/{len(chapters):03d}] Converting: "
            f"{chapter.title} -> {output_file.name}"
        )

        result = run_command(wav_cmd)
        if result.returncode == 0:
            created += 1
        else:
            print(
                f"  Conversion failed for chapter {chapter.index}.\n"
                f"  ffmpeg stderr: {result.stderr.strip() or '(empty)'}",
                file=sys.stderr,
            )

    if args.dry_run:
        print("Dry run complete. No files were created.")
        return

    print(f"Done. Created {created}/{len(chapters)} WAV files.")


if __name__ == "__main__":
    main()
