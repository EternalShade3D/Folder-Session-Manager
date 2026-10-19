# Folder Session Manager

<p align="center">
  <a href="https://github.com/EternalShade3D/Folder-Session-Manager/releases/latest"><img src="https://img.shields.io/github/v/release/EternalShade3D/Folder-Session-Manager?style=for-the-badge&labelColor=26282e&label=release&color=41CD52" alt="Latest release"></a>
  <a href="https://github.com/EternalShade3D/Folder-Session-Manager#requirements"><img src="https://img.shields.io/badge/Python-3.10%2B-3776AB?style=for-the-badge&labelColor=26282e&logo=python&logoColor=white" alt="Python 3.10+"></a>
  <img src="https://img.shields.io/badge/PySide6-6.x-41CD52?style=for-the-badge&labelColor=26282e&logo=qt&logoColor=white" alt="PySide6">
  <img src="https://img.shields.io/badge/Platform-Windows%2010%2F11-0078D4?style=for-the-badge&labelColor=26282e&logo=windows&logoColor=white" alt="Windows 10/11 only (pywin32 + UIA)">
  <a href="https://github.com/EternalShade3D/Folder-Session-Manager/blob/main/LICENSE"><img src="https://img.shields.io/badge/License-MIT-yellow?style=for-the-badge&labelColor=26282e" alt="License: MIT"></a>
</p>

Save and restore File Explorer window layouts — paths, positions, sizes, and open tabs.

## How It Works

1. **Save** — Open several Explorer windows with the folders you need, click **Save Session**. The app captures every window's paths, position, and size via COM, and falls back to UI Automation when COM cannot see the windows (Explorer often does not register them in the Running Object Table)
2. **Auto-save** — Click **Auto-save** in the status bar to save Explorer state every 5 minutes automatically
3. **Browse** — View saved sessions with path preview, search by folder name, expand/collapse tree
4. **Restore** — Select a session, click **Restore**. The app reopens Explorer windows at saved positions with all tabs restored via keyboard macro automation
5. **Manage** — Delete unwanted sessions to a .trash folder, choose custom session directory

## Features

- **Auto-save** — periodic save (QTimer, default 5 min), toggle from status bar
- **Save session** — capture all open Explorer windows (paths, tabs, position, size), with a UIA fallback when COM sees nothing
- **Restore session** — reopen windows at saved positions with all folder tabs via keystroke simulation
- **Preview** — tree view of saved sessions with window groups and paths
- **Search/filter** — type to filter paths across the preview tree
- **Trash management** — delete sessions to .trash with restore or permanent delete
- **Custom workspace** — choose which directory stores session files
- **Dark theme** — full dark UI with accent colors
- **Close existing windows** — optionally close all Explorers before restoring
- **Abort safe** — ESC or Cancel stops restoration mid-process

### Keyboard shortcuts
| Key | Action |
|-----|--------|
| ESC | Cancel restoration / abort auto-save |

## Requirements

- Windows 10/11
- Python 3.10+
- PySide6
- pywin32
- comtypes (UIA fallback for reading windows)
- pywinauto (UIA tab insertion)

```bash
pip install PySide6 pywin32 comtypes pywinauto
```

## Quick Start

Double-click `run_folder_session_manager.bat` (it picks the Python that has the
required modules), or run:

```bash
python FOLDERS_SESSION_MANAGER_v2.3.py
```

1. Open a few File Explorer windows to different folders
2. Click **Save Session** — a timestamped JSON file is created
3. Select a saved session in the list to preview its contents
4. Click **Restore Session** to reopen those windows

## Project History

The full evolution from v0.1 to v2.3 is in [`CHANGELOG.md`](CHANGELOG.md) and [`history/`](history/).

| Version | Highlights |
|---------|-----------|
| **v2.3** | Restore skips folders that no longer exist, finds windows through UIA |
| **v2.2** | UIA fallback when COM sees no windows, minimized-window notice |
| **v2.1** | Timing macros, keyboard macro tab insertion, abort-safe restoration |
| **v2.0** | Compact window sizing |
| **v1.9** | Graceful pywin32 import, streamlined code |
| **v1.8** | RestoreHelper class, subprocess explorer launch, context menus |
| **v1.7** | Qt Designer .ui file loading |
| **v1.6** | Progress dialog |
| **v1.4** | Trash/deletion system |
| **v0.8** | Emoji UI overhaul |
| **v0.7** | Dark/Light themes, platform check |
| **v0.1** | Initial session save/restore |

## License

MIT
