# AudioBookSlicer: Split `.m4b` by chapter

This utility splits a DRM-free `.m4b`/`.m4a` audiobook into one file per chapter using chapter metadata.

It uses:
- `ffprobe` to read chapter start/end times
- `ffmpeg` to extract each chapter with `-c copy` (no audio re-transcode) by default

## AudioBook Slicer GUI

A PyQt6 graphical interface for converting WAV files with Adobe Audition markers into chaptered MP3 files.

### GUI Features

- **Load WAV files** with chapter markers (Adobe Audition cue point format)
- **Edit chapter titles** in-place (double-click any title to edit)
- **Reorder chapters** via drag/drop in the chapter list
- **Preview chapter artwork** by selecting chapters in the list
- **Live encode progress** with per-chapter progress bars
- **Adjust encoding options**: VBR quality, CBR bitrate, mono/stereo, or dual mono
- **Size limit warnings**: Specify file size limit with unit selection (bytes, KB, MB, GB)
- **Persistent settings**: Last-opened directory remembered between sessions
- **Reset controls**: Reset all chapters to original markers or reset the entire app to defaults

### Running the GUI

The GUI requires PyQt6 to be installed in the project's virtual environment.

**Activate the virtual environment and run:**

```powershell
& '.venv\Scripts\Activate.ps1'
python gui.py
```

**Or use the venv Python directly:**

```powershell
& '.venv\Scripts\python.exe' gui.py
```

**Create a shortcut (optional):**

Create `run_gui.ps1`:
```powershell
& '.venv\Scripts\Activate.ps1'
python gui.py
```

Then double-click `run_gui.ps1` to launch.

### GUI Workflow

1. Click **Open WAV File** and select a WAV file with chapter markers
2. Enter your **Book Title** (used for matching cover artwork files)
3. Edit chapter titles by double-clicking them in the table
4. (Optional) **Drag chapters** to reorder them—the output will be encoded in your custom order
5. (Optional) Place cover images in the same folder as the WAV (e.g., `Book Title.jpg`, `Book Title.png`)
6. Click **Process Covers** to associate artwork with chapters
7. Choose encoding options (VBR/CBR, sample rate, mono/stereo)
8. Set a **Size Limit** if desired (file size warning threshold)
9. Click **Encode** and choose output location
10. Watch live progress bars update as the MP3 is encoded with chapter boundaries

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

## Convert Marker WAV To Chaptered MP3

If you have a single combined WAV with Adobe Audition markers (RIFF cue markers),
you can encode it to chaptered MP3 and carry chapters plus per-chapter artwork over.
The script expects chapter titles in the format `Book Title - *` and matching
cover images like `Book Title.jpg` or `Book Title.png` next to the WAV:

```powershell
python .\wav_markers_to_mp3.py "C:\path\to\combined.wav"
```

Default behavior:
- Reads WAV cue markers/labels and creates MP3 chapter tags
- Requires matching cover images before encoding starts
- Downmixes to mono (spoken-word friendly)
- Encodes as VBR MP3 with `-q:a 6`
- Preserves the input sample rate
- Writes `<input_stem>.mp3` next to input
- Warns if final output exceeds the effective 1.5% headroom budget, and again if it exceeds `1.0 GB`

Validation and inspection options:

```powershell
# Validate inputs and print VBR sizing notes, then continue converting
python .\wav_markers_to_mp3.py "combined.wav" --estimate

# Validate inputs only (no conversion)
python .\wav_markers_to_mp3.py "combined.wav" --estimate-only

# Change VBR quality and max-size warning threshold
python .\wav_markers_to_mp3.py "combined.wav" --vbr-quality 4 --max-size-gb 1.0

# Print marker/chapter timing debug info before conversion
python .\wav_markers_to_mp3.py "combined.wav" --debug-markers

# Preserve stereo instead of default mono downmix
python .\wav_markers_to_mp3.py "combined.wav" --stereo
```

You can also pass explicit FFmpeg paths if needed:

```powershell
python .\wav_markers_to_mp3.py "combined.wav" --ffmpeg "C:\ffmpeg\bin\ffmpeg.exe" --ffprobe "C:\ffmpeg\bin\ffprobe.exe"
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
