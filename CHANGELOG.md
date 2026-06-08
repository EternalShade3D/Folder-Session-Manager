# Changelog — Folder Session Manager

All notable changes to this project, in reverse chronological order.

---

## v2.1 (Latest)

### Added
- **Auto-save** — QTimer-based periodic save, toggle from status bar (default 5 min), saves immediately on enable
- **Fine-tuned restoration timing macros** — 12 configurable timing parameters for reliable Explorer window restoration
- **Robust restoration logic** — `check_cancelled()`, `safe_wait()`, `paths_match()` helper functions with progress dialog ESC abort
- **Keyboard macro tab insertion** — opens folder tabs via simulated keystrokes (Ctrl+T → Alt+D → Ctrl+V → Enter) with retry logic
- **HTML-formatted headers** — styled title, sessions, and preview labels
- **Emoji UI** — emoji prefixes restored on all labels and messages

### Changed
- `CONFIG` now includes `timings` block with per-step durations
- Escape key or Cancel button aborts restoration mid-process

---

## v2.0

### Changed
- Window sizing compacted — `fixed: (568, 721)`, `min: (200, 300)`
- `preview_min_height: 400` added

---

## v1.9

### Added
- Graceful `pywin32` import with try/except and user warning

### Changed
- Code streamlined — simplified TEXTS dict (only 8 keys), inline labels removed
- Clean section headers: `# CONFIGURATION`, `# CUSTOM WIDGETS`, `# MAIN APPLICATION`
- Window size changed to tall portrait: `(568, 1200)`

---

## v1.8

### Added
- **RestoreHelper class** — dedicated Explorer window management with:
  - `close_all_explorer_windows()` — sends WM_CLOSE, falls back to `taskkill /IM explorer.exe`
  - `open_explorer_window(path, x, y, width, height)` — open with position/size
  - `add_tabs_to_window(hwnd, paths)` — add folder tabs to existing windows
- Right-click context menu on tree items — Copy Path, Open Folder
- Trash dialog with restore/permanent delete options
- Search/filter preview in real time

### Changed
- Switched to `subprocess` for Explorer launch
- `import win32con`, `import pywintypes`

---

## v1.7

### Added
- **Qt Designer .ui file loading** — UI now loads from `session_manager.ui` at runtime via `QUiLoader`
- `CustomTreeWidget` emits `copyRequested` and `openRequested` signals
- Double-click opens folder in Explorer

### Changed
- Emoji removed from app title (encoding compatibility)
- Major import expansion: QPainter, QPalette, QPixmap, QUrl, QLocale, QDate, QTime
- Enhanced theming with depth comments

---

## v1.6

### Added
- **Progress dialog** — shows progress during restore with Cancel button
- `import win32api` (fixes NameError)

### Changed
- Header labels changed to ALL CAPS with emoji: `📁 FOLDER SESSION MANAGER`
- Font sizes bumped

---

## v1.5

### Added
- `QWheelEvent` import
- Wider layout: `fixed: (1200, 600)`, `min: (1200, 500)`
- `tree_col0_width` and `tree_col1_min_width` config

---

## v1.4

### Added
- **Trash/deletion system** — sessions moved to `.trash` folder with restore/permanent delete
- "Open Folder" action in context menu
- 24-hour trash auto-expiry via `QDateTime`

### Changed
- `import shutil` for trash operations
- `QListWidgetItem`, `QHeaderView`, `QMenu` imports

---

## v0.8

### Added
- **Emoji UI overhaul** — all labels and messages now use emoji prefixes
- Theme selector via `QComboBox`
- "Choose Session Folder" button
- `import unicodedata`

---

## v0.7

### Added
- **Dark/Light theme support** — configurable color schemes
- **Platform check** — exits with error if not Windows
- **Session preview search** — filter paths in real time
- Clear Preview button
- Configurable session directory
- QMenuBar, QStatusBar, QGridLayout, QLineEdit
- CustomTreeWidget with 2 columns and hover tooltips
- `apply_theme()` and `calculate_font_scale()` methods
- QSettings persists theme and session directory

---

## v0.5

No changes — duplicate of v0.4.

---

## v0.4

### Added
- **QSplitter** — preview tree + minimap in vertical splitter
- Compact view toggle
- Helper methods: `_create_spin_layout()`, `_init_spin()`

---

## v0.3

### Changed
- CONFIG restructured — colors moved to `styles` dict
- Sizes changed to tuple-based format
- Text dict compressed

---

## v0.2

### Added
- **Close confirmation dialog** — prompts before quitting
- "Close Existing Windows (Caution!)" checkbox

---

## v0.1

### Added
- Initial release: save/restore File Explorer window sessions
- **Dark theme** (`#121212` background with `#2979ff` accent)
- Explorer session capture via COM (`Shell.Application`)
- Save sessions as timestamped JSON
- Restore Explorer windows with paths and positions
- Folder preview tree by window
- MiniMapLabel showing scaled window positions
- Expand All / Collapse All buttons
- Copy path to clipboard
- Auto backup checkbox with interval spinner
- Timeout configuration
- QSettings persistence for geometry
