# JingleAllTheDay

Desktop jingle browser for quickly categorizing and finding samples in your GoXLR sample folder.

## Features

- Scans a sample library folder recursively for common audio files.
- Stores category tags per file.
- Supports multiple category tags per jingle.
- Fast search plus multi-category filters.
- Bulk apply category tags to selected rows.
- Play selected jingle directly from the app.
- Select audio output device and persist between launches.

## Default sample folder

The app starts with:

`C:\Users\brian\Documents\GoXLR\Samples`

You can change the folder from the GUI, and that choice is remembered.

## Install

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
```

## Run

```powershell
python gui.py
```

## Build EXE

Use PyInstaller from the project virtual environment:

```powershell
.\.venv\Scripts\Activate.ps1
pip install pyinstaller
pyinstaller .\JingleAllTheDay.spec
```

Build output is created in:

`dist\JingleAllTheDay\`

Distribute the entire `dist\JingleAllTheDay` folder, not just the `.exe`, because the default build is a one-folder package.

Or use the helper script in this repo, which flattens output into `dist\`:

```powershell
.\build_exe.ps1
```

That script places `JingleAllTheDay.exe` and `_internal\` directly under `dist\`.

## Build Installer (Windows)

This repo includes an Inno Setup script that builds a Windows installer with:

- Install location in the current user's AppData area:
	- `%LOCALAPPDATA%\Programs\JingleAllTheDay`
- A registered uninstaller in Windows `Installed Apps`
- Optional launch of the app at the end of setup
- Existing-install detection with choices to:
	- Re-install/upgrade
	- Uninstall current version and stop setup
	- Cancel
- Uninstall prompt asking whether to also remove user data:
	- `%APPDATA%\JingleAllTheDay\settings.ini`
	- `%APPDATA%\JingleAllTheDay\jingle-library.json`

### Prerequisite

Install Inno Setup 6 (provides `ISCC.exe`).

### Build steps

1. Build the bundled app output:

```powershell
.\build_exe.ps1
```

2. Build the installer:

```powershell
.\build_installer.ps1
```

The installer executable is written to:

`installer\JingleAllTheDay-Setup-<version>.exe`

You can override version/source when needed:

```powershell
.\build_installer.ps1 -AppVersion "1.0.0-custom" -SourceDir ".\dist"
```

## Runtime Dependencies

The packaged EXE does not require Python to be installed on the target system. PyInstaller bundles the Python runtime, PyQt6, and the Qt libraries used by the app.

Things that can still matter on the target machine:

- A working Windows audio device and drivers.
- Permission to read the user's jingle/sample folders.
- The `%APPDATA%` location, where the tag database is stored.
- `ffprobe` is optional. If it is not present, playback still works, but duration detection for some non-WAV formats may be less accurate.

The current build spec is stored in `JingleAllTheDay.spec`.

## Notes

- Metadata is saved to `%APPDATA%\JingleAllTheDay\jingle-library.json`.
- Settings are saved to `%APPDATA%\JingleAllTheDay\settings.ini`.
- In table edits and bulk fields, use comma or semicolon separated tags.
	- Example categories: `Holiday, Radio, Comedy`
- Category filtering supports multiple tags at once using comma-separated input.
  - Filter mode `Match Any` shows rows with at least one tag match.
  - Filter mode `Match All` shows rows containing every entered tag.
	- Default mode is `Match All`.
- Search scope dropdown supports:
	- `Name + Path + Tag`
	- `Name Only`
	- `Tag Only`
	- `Path Only`
- Active filter tags are shown as clickable chips under the filter box.
	- Click a chip to remove that tag from the active filter.
	- Use `Clear All` to remove every active category filter tag at once.
	- Chips are colorized per tag for quick visual scanning.
- Bulk update supports three modes:
	- Replace tags: overwrite tags on selected rows.
	- Append tags: add tags while preserving existing ones.
	- Remove tags: remove matching tags from selected rows.
	- Default mode is `Append tags`.

- Playback controls:
  - `Play Selected` toggles to `Stop Selected` while a sample is playing.
  - `Loop On/Off` controls continuous repeat until playback is stopped.
	- `Mode: Live/Preview` switches playback between your Live and Preview output devices.
	- When playback reaches the end naturally, status updates to `Playback finished`.

- Audio devices:
	- Configure both `Live Device` and `Preview Device` in `Tools > Options`.
	- If both devices are set to the same output, Preview/Live switching is disabled until Preview is changed.

- Tools menu:
	- `Update Categories from Folder Titles` derives category tags from folder names under your selected Samples folder.
	- Root-level jingles get no folder-derived tag.
	- The Samples root folder name is never added as a tag.
	- You can choose to preserve existing tags (merge, case-insensitive de-dup) or overwrite with derived folder tags.
	- `Clear All Categories` removes all category tags from every loaded jingle (with confirmation).

- Bulk section:
	- `From Folders (Selected)` applies the same folder-title update logic to selected rows only.
	- Uses the same preserve/overwrite prompt as the Tools command.

- File menu:
	- `Export Tag Database...` saves a JSON backup of your current tag database.
	- `Import Tag Database...` restores tags from a JSON backup (with confirmation).
