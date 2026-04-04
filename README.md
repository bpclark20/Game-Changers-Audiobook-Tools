# AudioBookSlicer: Split `.m4b` by chapter

This utility splits a DRM-free `.m4b`/`.m4a` audiobook into one file per chapter using chapter metadata.

It uses:
- `ffprobe` to read chapter start/end times
- `ffmpeg` to extract each chapter with `-c copy` (no audio re-transcode) by default

## Prerequisites

1. Python 3.9+
2. FFmpeg (must include both `ffmpeg` and `ffprobe`)

### Windows FFmpeg install options

1. `winget install Gyan.FFmpeg`
2. Or download a build from https://www.gyan.dev/ffmpeg/builds/ and add its `bin` folder to `PATH`

Verify:

```powershell
ffmpeg -version
ffprobe -version
```

## Usage

```powershell
python .\split_m4b_chapters.py "C:\path\to\book.m4b"
```

Default behavior:
- Creates folder `<input_stem>_chapters` next to input file
- Writes files like `001 - Chapter Name.m4b`
- Uses stream copy (`-c copy`) for no-transcode extraction

### Common options

```powershell
# Custom output folder
python .\split_m4b_chapters.py "book.m4b" -o ".\out"

# Show commands only (no file writes)
python .\split_m4b_chapters.py "book.m4b" --dry-run

# Retry failed chapters with AAC encoding
python .\split_m4b_chapters.py "book.m4b" --fallback-reencode --aac-bitrate 96k

# If ffmpeg/ffprobe are not on PATH
python .\split_m4b_chapters.py "book.m4b" --ffmpeg "C:\ffmpeg\bin\ffmpeg.exe" --ffprobe "C:\ffmpeg\bin\ffprobe.exe"
```

## Convert Marker WAV To Chaptered M4B

If you have a single combined WAV with Adobe Audition markers (RIFF cue markers),
you can encode it to M4B and carry chapters over:

```powershell
python .\wav_markers_to_m4b.py "C:\path\to\combined.wav"
```

Default behavior:
- Reads WAV cue markers/labels and creates M4B chapters
- Downmixes to mono (spoken-word friendly)
- Encodes to AAC at `64k` bitrate
- Writes `<input_stem>.m4b` next to input
- Warns if final output is larger than `1.0 GB`

Size estimation options:

```powershell
# Print approximate output size, then continue converting
python .\wav_markers_to_m4b.py "combined.wav" --estimate

# Estimate only (no conversion)
python .\wav_markers_to_m4b.py "combined.wav" --estimate-only

# Change bitrate and max-size warning threshold
python .\wav_markers_to_m4b.py "combined.wav" --bitrate 80k --max-size-gb 1.0

# Print marker/chapter timing debug info before conversion
python .\wav_markers_to_m4b.py "combined.wav" --debug-markers

# Preserve stereo instead of default mono downmix
python .\wav_markers_to_m4b.py "combined.wav" --stereo
```

You can also pass explicit FFmpeg paths if needed:

```powershell
python .\wav_markers_to_m4b.py "combined.wav" --ffmpeg "C:\ffmpeg\bin\ffmpeg.exe" --ffprobe "C:\ffmpeg\bin\ffprobe.exe"
```

## Strip Chapter Number Prefix From WAV Files

If you have files named like `024 - Chapter 24.wav`, `024 - Chapter 24 - Part 1.wav`,
`000 - Prologue.wav`, or `040 - Epilogue.wav`, you can remove the leading numeric
prefix with:

```powershell
# Preview only
.\strip_chapter_number_prefix.ps1

# Rename files
.\strip_chapter_number_prefix.ps1 -Apply

# Target a specific folder
.\strip_chapter_number_prefix.ps1 "C:\path\to\chapter\folder"
.\strip_chapter_number_prefix.ps1 "C:\path\to\chapter\folder" -Apply

# Include subfolders
.\strip_chapter_number_prefix.ps1 "C:\path\to\library" -Recurse
.\strip_chapter_number_prefix.ps1 "C:\path\to\library" -Recurse -Apply
```

## Prefix Book Title To Files

If you want to prepend the book title to every filename in a folder, including
sidecar files like `.pkf`, run:

```powershell
# Preview only, then enter the title when prompted
.\prefix_book_title.ps1 "C:\path\to\chapter\folder"

# Rename files, then enter the title when prompted
.\prefix_book_title.ps1 "C:\path\to\chapter\folder" -Apply
```

Example:

```text
Chapter 1.wav -> Role Model - Chapter 1.wav
Chapter 1.pkf -> Role Model - Chapter 1.pkf
```

## Notes

- Input must be DRM-free.
- If your file has multiple audio streams, use `--audio-stream` to select one.
- Some files may fail stream copy at certain boundaries depending on container/stream quirks; use `--fallback-reencode` if needed.
