import os
import sys
import json
import re
import platform
import subprocess
import time
import locale
import ctypes
from ctypes import wintypes, POINTER
from datetime import datetime
from urllib.parse import unquote

# Windows-specific libraries
try:
    import win32com.client
    import win32gui
    import win32api
    import pywintypes
    import win32con
    import win32clipboard
except ImportError:
    print("Warning: Missing 'pywin32' library. Install it with: pip install pywin32")

# UI Automation is required for reliable tab insertion (see _uia module below).
try:
    from pywinauto import Desktop as _UiaDesktop
except ImportError:
    _UiaDesktop = None

# comtypes drives the UIA fallback that reads open windows when COM cannot.
try:
    import comtypes.client as _comtypes_client
except ImportError:
    _comtypes_client = None

from PySide6.QtWidgets import (QApplication, QMainWindow, QTreeWidget, QTreeWidgetItem,
                               QListWidget, QListWidgetItem, QMenu, QMessageBox, QDialog,
                               QVBoxLayout, QHBoxLayout, QPushButton, QFileDialog, QProgressDialog, QStyle)
from PySide6.QtCore import Qt, QObject, Signal, QFile, QIODevice, QTimer
from PySide6.QtGui import QAction, QColor
from PySide6.QtUiTools import QUiLoader

# ==========================================
# CONFIGURATION
# ==========================================

CONFIG = {
    "sizes": {
        "fixed": (568, 721),
        "min": (200, 300),
        "tree_col0_width": 300,
        "tree_col1_min_width": 900,
        "preview_min_height": 400
    },
    "max_visible_items": 10,
    "session_dir": "C:\\03_SOFTWARE_WORKSPACE\\FOLDERS_SESSION_MANAGER",
    "trash_dir": "C:\\03_SOFTWARE_WORKSPACE\\FOLDERS_SESSION_MANAGER\\.trash",
    
    # Fine-tuning parameters for the restoration macros (in seconds)
    "timings": {
        "close_windows_buffer": 1.5,      # Time to wait after closing all existing windows
        "primary_window_spawn": 0.8,      # Wait after initially spawning a new window group
        "poll_attempts": 40,              # Max times to poll for a new window ID
        "poll_interval": 0.25,            # Wait between each poll attempt
        "app_activate": 0.2,              # Wait after activating the target window
        "tab_animation_min": 0.8,         # Base wait for Ctrl+T animation to complete
        "address_bar_focus": 0.3,         # Wait after Alt+D to ensure address bar is highlighted
        "clipboard_set": 0.1,             # Wait after setting clipboard text
        "paste_buffer": 0.2,              # Wait after Ctrl+V before pressing ENTER
        "navigation_cooldown_min": 1.2,   # Base wait for folder to load after ENTER
        "retry_cooldown": 0.5,            # Wait if a tab insertion macro fails before retrying
        "window_group_cooldown": 1.5      # Rest period between entirely different window groups
    }
}

TEXTS = {
    "app_title": "📁 Folder Session Manager V22",
    "main_header": "<html><head/><body><p><span style='font-size:14pt; font-weight:700;'>📁 FOLDER SESSION MANAGER</span></p></body></html>",
    "sessions_header": "<html><head/><body><p><span style='font-size:12pt; font-weight:700;'>📜 SESSIONS</span></p></body></html>",
    "preview_header": "<html><head/><body><p><span style='font-size:12pt; font-weight:700;'>🔎 SESSION PREVIEW</span></p></body></html>",
    "platform_error": "❌ This application only supports Windows OS.",
    "error_msg": "🧨 Error: ",
    "error_restoring": "🧨 Failed to restore session: {error}",
    "no_paths": "⚠️ No open File Explorer windows found.",
    "min_only": "⚠️ Explorer windows found, but minimized ones cannot be read — UIA needs a visible window.",
    "invalid_path": "🔗💥 Path non-existent: {path}",
    "session_saved": "✨ Session saved: {file}",
    "restored_window": "🖼️ Restored window {hwnd}",
    "item_count": "📊 {count} items in session",
    "abort_msg": "🛑 Restore stopped by user",
    "auto_save_on": "⏺️ Auto-save ON",
    "auto_save_off": "⏹️ Auto-save OFF",
    "auto_save_tick": "💾 Auto-saved at {time}",
}

# ==========================================
# MODULE-LEVEL HELPERS
# ==========================================

def _is_valid_path(path):
    """Returns True for filesystem paths, CLSID shell paths, and paths with special chars."""
    if not path:
        return False
    if path.startswith('::'):
        return True
    try:
        return os.path.exists(path)
    except Exception:
        return False

def _set_clipboard_text(text):
    """Write the clipboard without Qt's OLE clipboard.

    pywinauto's UIA backend needs COM in multithreaded mode, so Qt's
    OleInitialize() fails on startup (0x80010106) and
    QApplication.clipboard() silently stops writing. The Win32 clipboard
    API does not care about the COM apartment, so pasting still works.
    Returns True on success, False so the caller can fall back to Qt."""
    try:
        win32clipboard.OpenClipboard()
        try:
            win32clipboard.EmptyClipboard()
            win32clipboard.SetClipboardText(text)
        finally:
            win32clipboard.CloseClipboard()
        return True
    except Exception:
        return False


def _open_path(path):
    """Open path in Explorer — safe for CLSID (This PC) and unicode paths."""
    try:
        if path.startswith('::'):
            # CLSID — use COM Shell; explorer.exe routes CLSIDs incorrectly
            shell = win32com.client.Dispatch("Shell.Application")
            shell.Explore(path)
        else:
            subprocess.Popen(['explorer', path])
    except Exception:
        try:
            shell = win32com.client.Dispatch("Shell.Application")
            shell.Explore(path)
        except Exception:
            pass


def _open_in_new_tab(path):
    """Alias — both double-click and right-click → Open use _open_path for reliability."""
    _open_path(path)


# ==========================================
# TAB INSERTION VIA UI AUTOMATION
# ==========================================
#
# Why UIA instead of the old keyboard macro (Ctrl+T / Alt+D / paste / Enter):
#
#   * Ctrl+T via SendKeys needs the target window to hold the foreground.
#     Windows' foreground lock makes that unreliable, and any other app
#     grabbing focus (chat apps, launchers) makes the keystroke land in the
#     WRONG window — the new tab then appears in someone else's window.
#   * COM Navigate2 CANNOT open a path in a new tab. Verified failures:
#       - Navigate2(path)             -> replaces the active tab
#       - Navigate2(path, 0x80000000) -> navOpenInNewTab is ignored
#       - `explorer.exe a b c`        -> one window, no tabs
#       - PostMessage(WM_KEYDOWN)     -> Explorer ignores synthetic keys
#     Those all produced a window full of duplicate tabs.
#   * Invoking the "Add New Tab" button (automation_id "AddButton") creates a
#     tab with no keyboard and no focus change, and typing into the tab's own
#     address bar (class TextBox) targets that tab specifically.
#
# Verification uses the UIA tab label. COM LocationURL returns a stale value
# for every tab in a window and must never be used to check tab contents.

_UIA_MISSING_MSG = ("pywinauto is required to restore tabs reliably. "
                    "Install it with: pip install pywinauto")


# UIA window reading (fallback for when COM sees nothing)
# Shell.Application.Windows() only lists Explorer windows that registered in
# the Running Object Table. Explorer often does not register a window, so COM
# can report zero windows while several are open. UIA resolves a window from its
# handle and reads the active tab's path from the breadcrumb bar.
# A minimized window exposes no UIA tree, so its path can only come from COM.

_UIA_CLSID = "{FF48DBA4-60EF-4201-AA87-54103EEF594E}"
_UIA_SCOPE_CHILDREN = 3
_UIA_SCOPE_DESCENDANTS = 4
_UIA_PROP_AUTOMATION_ID = 30011
_ENUMWINDOWS_PROC = ctypes.WINFUNCTYPE(ctypes.c_bool, ctypes.c_void_p, ctypes.c_void_p)


def _uia_automation():
    """IUIAutomation instance, or None when comtypes is not installed."""
    if _comtypes_client is None:
        return None
    try:
        module = _comtypes_client.GetModule("UIAutomationCore.dll")
        return _comtypes_client.CreateObject(_UIA_CLSID, interface=module.IUIAutomation)
    except Exception as e:
        print(f"UIA Error: {e}")
        return None


def _crumbs_to_path(crumbs):
    """Breadcrumb labels to a filesystem path ('Local Disk (C:)' becomes 'C:')."""
    crumbs = [c for c in crumbs if c and c not in ("More", "This PC")]
    if not crumbs:
        return ""
    drive = re.fullmatch(r".*\(([A-Z]):?\)", crumbs[0])
    if not drive:
        return ""
    return drive.group(1) + ":" + "".join("\\" + c for c in crumbs[1:])


def _uia_explorer_windows():
    """Every open Explorer window read straight from UIA, no COM or ROT needed."""
    auto = _uia_automation()
    if auto is None:
        return []
    bar_cond = auto.CreatePropertyCondition(_UIA_PROP_AUTOMATION_ID, "PART_BreadcrumbBar")
    true_cond = auto.CreateTrueCondition()
    user32 = ctypes.windll.user32
    found = []

    def visit(hwnd, _):
        if not user32.IsWindowVisible(hwnd):
            return True
        name = ctypes.create_unicode_buffer(256)
        user32.GetClassNameW(hwnd, name, 256)
        if name.value != "CabinetWClass":
            return True
        hwnd = int(hwnd)
        path, title = "", ""
        try:
            # Explorer redraws make single UIA calls fail now and then.
            element = _uia_retry(lambda: auto.ElementFromHandle(hwnd), tries=3, delay=0.2)
            if element:
                title = element.CurrentName or ""
                bar = _uia_retry(lambda: element.FindFirst(_UIA_SCOPE_DESCENDANTS, bar_cond), tries=3, delay=0.2)
                if bar:
                    items = _uia_retry(lambda: bar.FindAll(_UIA_SCOPE_CHILDREN, true_cond), tries=3, delay=0.2)
                    if items:
                        path = _crumbs_to_path([items.GetElement(i).CurrentName
                                                for i in range(items.Length)])
        except Exception:
            pass
        found.append({
            "hwnd": hwnd,
            "path": path,
            "title": title,
            "minimized": bool(user32.IsIconic(hwnd)),
        })
        return True

    user32.EnumWindows(_ENUMWINDOWS_PROC(visit), 0)
    return found


def _uia_retry(fn, tries=6, delay=0.6):
    """UIA/COM calls raise transient COMError while Explorer redraws."""
    last = None
    for _ in range(tries):
        try:
            return fn()
        except Exception as e:
            last = e
            time.sleep(delay)
    raise last


def _uia_win(hwnd):
    return _UiaDesktop(backend="uia").window(handle=hwnd)


def _tab_labels(hwnd):
    """Tab labels as Explorer itself shows them (may be truncated)."""
    if _UiaDesktop is None:
        return []
    try:
        return _uia_retry(lambda: [t.window_text() for t in
                                   _uia_win(hwnd).descendants(control_type="TabItem")])
    except Exception:
        return []


def _tab_count(hwnd):
    """Tabs from the title bar: 'X and 3 more tabs' -> 4, 'X' -> 1."""
    try:
        title = win32gui.GetWindowText(hwnd)
    except Exception:
        return -1
    m = re.search(r'and (\d+) more tab', title)
    return int(m.group(1)) + 1 if m else 1


def _add_tab_button(hwnd):
    for c in _uia_win(hwnd).descendants(control_type="Button"):
        try:
            if c.automation_id() == "AddButton":
                return c
        except Exception:
            continue
    return None


def _address_bar(hwnd):
    """The address bar is the Edit whose class is TextBox (not the
    System.ItemNameDisplay property edits belonging to the file list)."""
    for e in _uia_win(hwnd).descendants(control_type="Edit"):
        try:
            if e.element_info.class_name == "TextBox":
                return e
        except Exception:
            continue
    return None


def _expected_label(path):
    """Explorer's tab label for a path is its folder name."""
    name = os.path.basename(os.path.normpath(path).rstrip('\\/'))
    return name or path


def _tab_has_path(hwnd, index, path):
    """True when tab `index` shows `path`. Compares the UIA label; long folder
    names are truncated by Explorer, so also accept a label the path starts with."""
    labels = _tab_labels(hwnd)
    if index >= len(labels):
        return False
    label = labels[index]
    expected = _expected_label(path)
    return label == expected or expected.startswith(label)


def _insert_tab(hwnd, path, cancel_check=None, wait=0.35):
    """Create one tab in `hwnd` showing `path`.

    Returns True only when the tab count grew AND the tab really holds `path`.
    Retries a few times: a lost keystroke lands the tab on 'This PC' or the
    previous folder, which is why the result is verified rather than assumed.
    """
    if _UiaDesktop is None:
        return False
    # A folder that no longer exists can never be verified, so retrying
    # for ~30 seconds on it is what made the app look stuck.
    if not _is_valid_path(path):
        return False

    for _ in range(3):
        if cancel_check and cancel_check():
            return False
        base = _tab_count(hwnd)
        try:
            btn = _uia_retry(lambda: _add_tab_button(hwnd))
        except Exception:
            time.sleep(0.5)
            continue
        if btn is None:
            time.sleep(0.5)
            continue

        try:
            # invoke() targets the control directly: no mouse coordinates and
            # no focus change, so another app cannot misdirect it.
            btn.invoke()
        except Exception:
            time.sleep(0.4)
            continue

        created = False
        for _ in range(15):
            if cancel_check and cancel_check():
                return False
            time.sleep(0.2)
            if _tab_count(hwnd) > base:
                created = True
                break
        if not created:
            continue

        time.sleep(0.9)
        try:
            edit = _uia_retry(lambda: _address_bar(hwnd))
        except Exception:
            return False
        if edit is None:
            return False
        try:
            edit.set_focus()
        except Exception:
            pass
        time.sleep(0.4)
        try:
            edit.type_keys(path, with_spaces=True, with_tabs=False, with_newlines=False)
            time.sleep(0.4)
            edit.type_keys("{ENTER}")
        except Exception:
            return False
        time.sleep(1.7)

        index = _tab_count(hwnd) - 1
        if _tab_has_path(hwnd, index, path):
            return True
        # Tab exists but holds the wrong folder: try again in the same slot.
        for _ in range(3):
            if cancel_check and cancel_check():
                return False
            time.sleep(0.35)
            if _tab_has_path(hwnd, index, path):
                return True

    return False


def _refill_tab(hwnd, index, path, cancel_check=None):
    """Point an EXISTING tab at `path` (used to repair a wrong tab)."""
    if _UiaDesktop is None:
        return False
    if not _is_valid_path(path):
        return False
    try:
        items = _uia_retry(lambda: list(_uia_win(hwnd).descendants(control_type="TabItem")))
    except Exception:
        return False
    if index >= len(items):
        return False
    for attempt in range(3):
        if cancel_check and cancel_check():
            return False
        try:
            items[index].select()
        except Exception:
            pass
        time.sleep(1.0)
        try:
            edit = _uia_retry(lambda: _address_bar(hwnd))
        except Exception:
            return False
        if edit is None:
            return False
        try:
            edit.set_focus()
            time.sleep(0.5)
            edit.type_keys(path, with_spaces=True, with_tabs=False, with_newlines=False)
            time.sleep(0.4)
            edit.type_keys("{ENTER}")
        except Exception:
            pass
        time.sleep(2.0)
        for _ in range(3):
            if cancel_check and cancel_check():
                return False
            time.sleep(0.35)
            if _tab_has_path(hwnd, index, path):
                return True
    return False


# ==========================================
# CLSID → REAL PATH RESOLUTION
# ==========================================

def _resolve_clsid_to_path(clsid_str):
    """Map a Windows Shell CLSID (::{xxx}) to its real filesystem path.

    Uses SHGetFolderPathW for known CSIDL folders, then SHGetKnownFolderPath
    for modern folders, then expanduser fallback as last resort.
    """
    clean = clsid_str.lstrip(':').upper().strip('{}')

    # CLSID → CSIDL mapping for standard folders
    # CSIDL values: https://learn.microsoft.com/en-us/windows/win32/shell/csidl
    CLSID_TO_CSIDL = {
        # Desktop directory (C:\Users\NAME\Desktop)
        "5E6C858F-0E22-4760-9AFE-EA3317B67173": 0x0010,  # CSIDL_DESKTOPDIRECTORY
        # Documents (My Documents)
        "FDD39AD0-238F-46AF-ADB4-6C85480369C7": 0x0005,  # CSIDL_PERSONAL
        # Pictures
        "33E28130-4E1E-4256-9F93-63D3E1B2B1AE": 0x0027,  # CSIDL_MYPICTURES
        # Music
        "4BD8D571-6D19-48D3-BE97-422220080E43": 0x000D,  # CSIDL_MYMUSIC
        # Videos
        "18989B1D-99B5-455B-841C-AB7C74E4DDFC": 0x000E,  # CSIDL_MYVIDEO
    }

    # Try SHGetFolderPathW for known CSIDL folders
    csidl = CLSID_TO_CSIDL.get(clean)
    if csidl is not None:
        try:
            buf = ctypes.create_unicode_buffer(260)
            if ctypes.windll.shell32.SHGetFolderPathW(0, csidl, 0, 0, buf) == 0:
                return buf.value
        except Exception:
            pass

    # Fallback: SHGetKnownFolderPath for folders available via KNOWNFOLDERID
    # Downloads (no CSIDL — Vista+ only via KNOWNFOLDERID)
    if clean == "374DE290-123F-4565-9164-39C4925E467B":
        try:
            from uuid import UUID
            # KNOWNFOLDERID for Downloads = {374DE290-123F-4565-9164-39C4925E467B}
            guid_bytes = UUID("374DE290-123F-4565-9164-39C4925E467B").bytes_le
            rfid = (wintypes.BYTE * 16).from_buffer_copy(guid_bytes)
            fn = ctypes.windll.shell32.SHGetKnownFolderPath
            fn.restype = wintypes.HRESULT
            fn.argtypes = [ctypes.c_void_p, wintypes.DWORD, wintypes.HANDLE, POINTER(wintypes.PWSTR)]
            ptr = wintypes.PWSTR()
            if fn(ctypes.byref(rfid), 0, 0, ctypes.byref(ptr)) == 0:
                path = ptr.value
                ctypes.windll.ole32.CoTaskMemFree(ptr)
                if path:
                    return path
        except Exception:
            pass

    # Last resort — use expanduser for standard sub-directories
    home = os.path.expanduser('~')
    EXPANDUSER_MAP = {
        "5E6C858F-0E22-4760-9AFE-EA3317B67173": os.path.join(home, "Desktop"),
        "FDD39AD0-238F-46AF-ADB4-6C85480369C7": os.path.join(home, "Documents"),
        "374DE290-123F-4565-9164-39C4925E467B": os.path.join(home, "Downloads"),
        "33E28130-4E1E-4256-9F93-63D3E1B2B1AE": os.path.join(home, "Pictures"),
        "4BD8D571-6D19-48D3-BE97-422220080E43": os.path.join(home, "Music"),
        "18989B1D-99B5-455B-841C-AB7C74E4DDFC": os.path.join(home, "Videos"),
    }
    path = EXPANDUSER_MAP.get(clean)
    if path and os.path.exists(path):
        return path

    return None  # Not resolved — caller keeps the CLSID


def _path_display_name(path):
    """Return user-friendly display name for a filesystem or CLSID path."""
    if path.startswith('::'):
        # Map known CLSIDs to friendly names
        CLSID_TO_TITLE = {
            "20D04FE0-3AEA-1069-A2D8-08002B30309D": "This PC",
            "645FF040-5081-101B-9F08-00AA002F954E": "Recycle Bin",
            "208D2C60-3AEA-1069-A2D7-08002B30309D": "Network",
            "031E4825-7B94-4DC3-B131-E946B44C8DD5": "Libraries",
            "E88865EA-0E1C-4E20-9AA6-EDCD0212C87C": "Gallery",
        }
        clean = path.lstrip(':').upper().strip('{}')
        title = CLSID_TO_TITLE.get(clean)
        if title:
            return title
        # Unknown CLSID — show short form ::{ABCD...}
        short = clean[:8] + "…" if len(clean) > 8 else clean
        return f"::{short}"
    return os.path.basename(path) or path


# ==========================================
# CUSTOM WIDGETS
# ==========================================

class CustomTreeWidget(QTreeWidget):
    """Extends QTreeWidget to include right-click context menus for paths."""
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setContextMenuPolicy(Qt.CustomContextMenu)
        self.customContextMenuRequested.connect(self.show_context_menu)
        self.itemDoubleClicked.connect(self.on_item_double_clicked)

    def show_context_menu(self, position):
        item = self.itemAt(position)
        if not item or not item.text(1): return
        
        menu = QMenu(self)
        copy_action = QAction("📋 Copy Path", self)
        open_action = QAction("📂 Open Folder", self)
        
        copy_action.triggered.connect(lambda: _set_clipboard_text(item.text(1)) or QApplication.clipboard().setText(item.text(1)))
        open_action.triggered.connect(lambda: _open_path(item.text(1)))
        
        menu.addAction(copy_action)
        menu.addAction(open_action)
        menu.exec(self.viewport().mapToGlobal(position))

    def on_item_double_clicked(self, item, column):
        if item and item.text(1):
            _open_path(item.text(1))

# ==========================================
# MAIN APPLICATION
# ==========================================

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        if platform.system() != "Windows":
            QMessageBox.critical(self, TEXTS["error_msg"], TEXTS["platform_error"])
            sys.exit(1)
            
        # Save settings next to the script so workspace choice is remembered persistently
        self.settings_file = os.path.join(os.path.dirname(os.path.abspath(__file__)), "settings.json")
        self.load_settings()
        
        self.session_dir = self.settings.get("session_dir", CONFIG["session_dir"])
        self.trash_dir = os.path.join(self.session_dir, ".trash")
        
        self.ensure_dirs()
        self.setup_ui()
        self.update_session_list()

        # Auto-save timer
        self.auto_save_timer = QTimer(self)
        interval_min = self.settings.get("auto_save_interval", 5)
        self.auto_save_timer.setInterval(interval_min * 60 * 1000)
        self.auto_save_timer.timeout.connect(self.auto_save_tick)

    def load_settings(self):
        self.settings = {"timeout": 10, "session_dir": CONFIG["session_dir"]}
        try:
            if os.path.exists(self.settings_file):
                with open(self.settings_file, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                    self.settings.update(data)
        except Exception as e:
            print(f"Error loading settings: {e}")

    def save_settings(self):
        try:
            with open(self.settings_file, 'w', encoding='utf-8') as f:
                json.dump(self.settings, f, indent=4)
        except Exception as e:
            print(f"Error saving settings: {e}")

    def ensure_dirs(self):
        """Creates necessary workspace folders."""
        try:
            os.makedirs(self.session_dir, exist_ok=True)
            os.makedirs(self.trash_dir, exist_ok=True)
        except OSError as e:
            print(f"Error creating directories: {e}")

    def setup_ui(self):
        """Loads the V18 design from the UI file and wires logic safely."""
        loader = QUiLoader()
        loader.registerCustomWidget(CustomTreeWidget)
        
        ui_file_path = os.path.join(os.path.dirname(__file__), "session_manager.ui")
        if not os.path.exists(ui_file_path):
            QMessageBox.critical(self, "UI Error", f"Missing session_manager.ui at {ui_file_path}")
            sys.exit(1)

        file = QFile(ui_file_path)
        if not file.open(QIODevice.ReadOnly):
            QMessageBox.critical(self, "UI Error", f"Could not open UI file: {ui_file_path}")
            sys.exit(1)
        
        self.ui_container = loader.load(file)
        file.close()

        if not self.ui_container:
            QMessageBox.critical(self, "UI Error", "Failed to load UI object.")
            sys.exit(1)

        # Copy widgets from ui_container to self
        for child in self.ui_container.findChildren(QObject):
            name = child.objectName()
            if name:
                setattr(self, name, child)

        if hasattr(self.ui_container, 'centralwidget'):
            self.setCentralWidget(self.ui_container.centralwidget)
        
        if hasattr(self.ui_container, 'statusbar'):
            self.setStatusBar(self.ui_container.statusbar)
            
        self.setWindowTitle(TEXTS["app_title"])
        
        # Override UI File Labels with V21 Emojis
        if hasattr(self, 'tlabel'): self.tlabel.setText(TEXTS["main_header"])
        if hasattr(self, 'sessions_label'): self.sessions_label.setText(TEXTS["sessions_header"])
        if hasattr(self, 'preview_label'): self.preview_label.setText(TEXTS["preview_header"])
        
        if hasattr(self, 'preview_tree'):
            self.preview_tree.setHeaderLabels(["📂 Name", "📍 Full Path"])
            self.preview_tree.setColumnWidth(0, CONFIG["sizes"]["tree_col0_width"])

        if hasattr(self, 'timeout_spin'):
            self.timeout_spin.setValue(self.settings.get("timeout", 10))
            self.timeout_spin.valueChanged.connect(self.on_timeout_changed)

        # UI Element Labeling with Emojis
        if hasattr(self, 'choose_dir_button'): self.choose_dir_button.setText("🏢 Choose Root")
        if hasattr(self, 'view_trash_button'): self.view_trash_button.setText("♻️ View Trash")
        if hasattr(self, 'save_button'): self.save_button.setText("💾 Save Session")
        if hasattr(self, 'restore_button'): self.restore_button.setText("🔄 Restore Session")
        if hasattr(self, 'clear_button'): self.clear_button.setText("🗑️ Clear Preview")

        # Signal Mapping
        if hasattr(self, 'save_button'): self.save_button.clicked.connect(self.save_session)
        if hasattr(self, 'restore_button'): self.restore_button.clicked.connect(self.restore_session)
        if hasattr(self, 'delete_session_button'): self.delete_session_button.clicked.connect(self.delete_to_trash)
        if hasattr(self, 'sessions_list'): self.sessions_list.itemClicked.connect(self.update_preview)
        if hasattr(self, 'choose_dir_button'): self.choose_dir_button.clicked.connect(self.change_session_root)
        if hasattr(self, 'view_trash_button'): self.view_trash_button.clicked.connect(self.show_trash_dialog)
        if hasattr(self, 'search_input'): self.search_input.textChanged.connect(self.filter_preview)
        if hasattr(self, 'expand_button'): self.expand_button.clicked.connect(lambda: self.preview_tree.expandAll() if hasattr(self, 'preview_tree') else None)
        if hasattr(self, 'collapse_button'): self.collapse_button.clicked.connect(lambda: self.preview_tree.collapseAll() if hasattr(self, 'preview_tree') else None)
        if hasattr(self, 'clear_button'): self.clear_button.clicked.connect(self.clear_preview)

        # Auto-save toggle button in status bar
        self.auto_save_btn = QPushButton(TEXTS["auto_save_off"])
        self.auto_save_btn.setFixedHeight(24)
        self.auto_save_btn.setStyleSheet(
            "QPushButton { padding: 2px 10px; font-size: 11px; border: 1px solid #555; border-radius: 4px; }"
        )
        self.auto_save_btn.clicked.connect(self.toggle_auto_save)
        if hasattr(self, 'statusbar'):
            self.statusbar.addPermanentWidget(self.auto_save_btn)

    def on_timeout_changed(self):
        self.settings["timeout"] = self.timeout_spin.value()
        self.save_settings()

    def clear_preview(self):
        if hasattr(self, 'preview_tree'): self.preview_tree.clear()
        if hasattr(self, 'item_count_label'): self.item_count_label.setText(TEXTS["item_count"].format(count=0))

    # --- EXPLORER INTERACTION LOGIC ---

    @staticmethod
    def _resolve_explorer_path(window):
        """Get the real filesystem path from an Explorer window.

        Fixes CLSID paths (::{xxxx}) returned by Self.Path for library
        folders by reading LocationURL (file:/// format) instead, with
        encoding fallback for non-UTF-8 systems.
        """
        # Method 1: LocationURL (works for library folders)
        try:
            url = window.LocationURL
            if url and url.startswith('file:///'):
                path_part = url[8:]
                # Try UTF-8 first (standard), then system ANSI code page
                for enc in ('utf-8', locale.getpreferredencoding(), 'latin-1'):
                    try:
                        decoded = unquote(path_part, encoding=enc, errors='strict')
                        break
                    except (UnicodeDecodeError, LookupError):
                        continue
                else:
                    decoded = unquote(path_part, encoding='utf-8', errors='replace')
                return decoded.replace('/', '\\')
        except Exception:
            pass

        # Method 2: Self.Path (normal folders)
        try:
            path = window.Document.Folder.Self.Path
            if not path.startswith('::'):
                return path
            # CLSID — try to resolve to real path, fall back to CLSID itself
            resolved = _resolve_clsid_to_path(path)
            if resolved:
                return resolved
            # Virtual folder (Recycle Bin, Network, Libraries, Gallery, This PC) —
            # save the CLSID so we can reopen it via shell.Explore() later
            return path
        except Exception:
            pass

        return ""

    def get_active_explorer_data(self):
        """Captures paths and window positions for all open Explorers."""
        windows = []
        try:
            shell = win32com.client.Dispatch("Shell.Application")
            hwnd_map = {}
            for window in shell.Windows():
                try:
                    path = self._resolve_explorer_path(window)
                    if not path:
                        continue
                    hwnd = window.HWND
                    if hwnd not in hwnd_map:
                        hwnd_map[hwnd] = []
                    hwnd_map[hwnd].append(path)
                except: continue

            for hwnd, paths in hwnd_map.items():
                rect = win32gui.GetWindowRect(hwnd)
                windows.append({
                    "hwnd": hwnd,
                    "paths": paths,
                    "x": rect[0], "y": rect[1],
                    "width": rect[2] - rect[0],
                    "height": rect[3] - rect[1]
                })
        except Exception as e:
            print(f"Shell Error: {e}")

        # COM only sees windows registered in the ROT. Anything it missed
        # is picked up through UIA, which reads every window by handle.
        seen = {w["hwnd"] for w in windows}
        for win in _uia_explorer_windows():
            if win["hwnd"] in seen or not win["path"]:
                continue
            seen.add(win["hwnd"])
            try:
                rect = win32gui.GetWindowRect(win["hwnd"])
            except Exception:
                continue
            windows.append({
                "hwnd": win["hwnd"],
                "paths": [win["path"]],
                "x": rect[0], "y": rect[1],
                "width": rect[2] - rect[0],
                "height": rect[3] - rect[1],
                "source": "uia",
            })
        return windows

    def save_session(self):
        data = self.get_active_explorer_data()
        if not data:
            label = TEXTS["no_paths"]
            if any(not w["path"] for w in _uia_explorer_windows()):
                label = TEXTS["min_only"]
            if hasattr(self, 'statusbar'): self.statusbar.showMessage(label, 5000)
            return

        ts = datetime.now().strftime('%Y-%m-%d_%H-%M-%S')
        filename = f"explorer_session_{ts}.json"
        path = os.path.join(self.session_dir, filename)

        with open(path, 'w', encoding='utf-8') as f:
            json.dump(data, f, indent=4)
        
        self.update_session_list()
        if hasattr(self, 'statusbar'): self.statusbar.showMessage(TEXTS["session_saved"].format(file=filename), 5000)

    def restore_session(self):
        """Robust Restoration Logic with Isolated Macros and Reliable Polling."""
        if not hasattr(self, 'sessions_list'): return
        item = self.sessions_list.currentItem()
        if not item: return
        
        file_path = item.data(Qt.UserRole)
        with open(file_path, 'r', encoding='utf-8') as f:
            session_data = json.load(f)

        t_cfg = CONFIG["timings"]

        if hasattr(self, 'close_windows_cb') and self.close_windows_cb.isChecked():
            self.close_all_explorers()
            end_time = time.time() + t_cfg["close_windows_buffer"]
            while time.time() < end_time:
                QApplication.processEvents()
                time.sleep(0.05)

        shell = win32com.client.Dispatch("Shell.Application")
        wshell = win32com.client.Dispatch("WScript.Shell")
        clipboard = QApplication.clipboard()

        total = sum(len(w['paths']) for w in session_data)
        progress = QProgressDialog("🚀 Restoring Session... (Press ESC to cancel)", "Cancel", 0, total, self)
        progress.setWindowModality(Qt.WindowModal)
        progress.setMinimumDuration(0)
        progress.setValue(0)
        
        count = 0
        skipped_groups = 0
        failed_paths = []
        dead_paths = []
        timeout_base = max(1, self.settings.get("timeout", 10)) / 10.0

        if _UiaDesktop is None and hasattr(self, 'statusbar'):
            self.statusbar.showMessage(_UIA_MISSING_MSG, 10000)

        def check_cancelled():
            if progress.wasCanceled() or (win32api.GetAsyncKeyState(0x1B) & 0x8000):
                if hasattr(self, 'statusbar'): self.statusbar.showMessage(TEXTS["abort_msg"], 5000)
                progress.close()
                return True
            return False

        def safe_wait(duration_seconds):
            end_time = time.time() + duration_seconds
            while time.time() < end_time:
                if check_cancelled(): return True
                time.sleep(0.05)
                QApplication.processEvents()
            return False
            
        def paths_match(p1, p2):
            if not p1 or not p2: return False
            try: return os.path.normcase(os.path.normpath(p1)) == os.path.normcase(os.path.normpath(p2))
            except: return False

        restored_hwnds = set()
        
        for win_data in session_data:
            if check_cancelled(): return
            paths = win_data['paths']
            if not paths: continue

            first_path = paths[0]
            # CLSID paths (::{xxx}) are valid shell paths, not filesystem
            if not first_path.startswith('::') and not os.path.exists(first_path):
                dead_paths.append(first_path)
                # Folder is gone: open the nearest folder that still exists
                # instead of jumping to the home folder.
                parent = os.path.dirname(os.path.normpath(first_path))
                first_path = parent if parent and os.path.exists(parent) else os.path.expanduser("~")
            existing_hwnds = set()
            try:
                for w in shell.Windows():
                    try: existing_hwnds.add(w.HWND)
                    except: pass
            except: pass
            # COM only lists windows that registered in the ROT, so it can
            # see nothing while windows are open. UIA sees them all.
            try:
                for w in _uia_explorer_windows():
                    existing_hwnds.add(w["hwnd"])
            except Exception:
                pass

            try:
                if first_path.startswith('::'):
                    # CLSID path — open via COM Shell, not explorer.exe
                    shell.Explore(first_path)
                else:
                    # /separate forces a new process — prevents Explorer from
                    # reusing an existing window that already shows this path,
                    # which would override the saved window position/size
                    subprocess.Popen(['explorer.exe', '/separate', '/n,', os.path.normpath(first_path)])
            except Exception as e:
                try:
                    subprocess.Popen(['explorer.exe', '/separate', '/n,', os.path.normpath(first_path)])
                except Exception:
                    shell.Explore(first_path)
                
            if safe_wait(t_cfg["primary_window_spawn"]): return

            new_hwnd = None
            for attempt in range(t_cfg["poll_attempts"]):
                if check_cancelled(): return
                try:
                    fresh_shell = win32com.client.Dispatch("Shell.Application")
                    for w in fresh_shell.Windows():
                        try:
                            w_hwnd = w.HWND
                            if w_hwnd not in existing_hwnds and w_hwnd not in restored_hwnds:
                                if self._resolve_explorer_path(w): 
                                    new_hwnd = w_hwnd
                                    break
                        except: pass
                    if new_hwnd: break
                except: pass
                # UIA is heavier than COM, so only consult it every few polls.
                if attempt % 5 == 0:
                    try:
                        for w in _uia_explorer_windows():
                            if (w["hwnd"] not in existing_hwnds and
                                    w["hwnd"] not in restored_hwnds and w["path"]):
                                new_hwnd = w["hwnd"]
                                break
                    except Exception:
                        pass
                    if new_hwnd: break
                if safe_wait(t_cfg["poll_interval"]): return

            if not new_hwnd:
                try:
                    fresh_shell = win32com.client.Dispatch("Shell.Application")
                    for w in fresh_shell.Windows():
                        try:
                            w_hwnd = w.HWND
                            if w_hwnd not in restored_hwnds and paths_match(self._resolve_explorer_path(w), first_path):
                                new_hwnd = w_hwnd
                                break
                        except: pass
                except: pass
            if not new_hwnd:
                try:
                    for w in _uia_explorer_windows():
                        if w["hwnd"] not in restored_hwnds and paths_match(w["path"], first_path):
                            new_hwnd = w["hwnd"]
                            break
                except Exception:
                    pass

            if new_hwnd:
                restored_hwnds.add(new_hwnd)
                # A minimized window exposes no UIA tree, so tabs cannot be
                # created in it. Park it somewhere valid, do the work, and
                # re-minimize afterwards if the session had it minimized.
                was_minimized = (win_data['x'] <= -30000 and win_data['y'] <= -30000)
                try:
                    if win32gui.IsIconic(new_hwnd):
                        win32gui.ShowWindow(new_hwnd, win32con.SW_RESTORE)
                except Exception: pass
                try:
                    if was_minimized:
                        win32gui.MoveWindow(new_hwnd, 60, 60, 1294, 703, True)
                    else:
                        win32gui.MoveWindow(new_hwnd, win_data['x'], win_data['y'],
                                           win_data['width'], win_data['height'], True)
                except Exception: pass
                bad_tabs = []
                for idx_p, p in enumerate(paths[1:], start=1):
                    if check_cancelled(): return
                    # Dead folder: no typing, no retry, just record it.
                    if not _is_valid_path(p):
                        dead_paths.append(p)
                        count += 1
                        progress.setValue(count)
                        continue
                    # Insert via UI Automation: creates the tab without the
                    # keyboard (so focus theft can't misdirect it) and types
                    # the path into that tab's own address bar. Verified by
                    # reading the tab label back — a tab that landed on
                    # 'This PC' is detected and retried/repaired.
                    if _UiaDesktop is None:
                        # pywinauto unavailable: fall back to the old macro.
                        if not self._legacy_insert_tab(new_hwnd, p, wshell,
                                                       clipboard, check_cancelled,
                                                       safe_wait, t_cfg, timeout_base):
                            bad_tabs.append(p)
                    elif not _insert_tab(new_hwnd, p, cancel_check=check_cancelled):
                        # Tab exists but is wrong -> repair it in place.
                        if not _refill_tab(new_hwnd, idx_p, p, cancel_check=check_cancelled):
                            bad_tabs.append(p)

                    count += 1
                    progress.setValue(count)

                if bad_tabs:
                    failed_paths.extend(bad_tabs)

                if was_minimized:
                    try:
                        win32gui.ShowWindow(new_hwnd, win32con.SW_MINIMIZE)
                    except Exception: pass

            else:
                skipped_groups += 1
                count += len(paths) - 1

            count += 1
            progress.setValue(count)
            if safe_wait(t_cfg["window_group_cooldown"]): return

        progress.close()
        if hasattr(self, 'statusbar'):
            if dead_paths:
                self.statusbar.showMessage(
                    f"⚠️ {len(dead_paths)} folder(s) no longer exist and were skipped: " 
                    f"{', '.join(os.path.basename(x) for x in dead_paths[:4])}", 10000)
            elif failed_paths:
                n = len(failed_paths)
                self.statusbar.showMessage(
                    f"⚠️ Restoration finished with {n} tab(s) wrong: "
                    f"{', '.join(os.path.basename(x) for x in failed_paths[:4])}", 10000)
            elif skipped_groups > 0:
                self.statusbar.showMessage(f"⚠️ Restoration finished with {skipped_groups} skipped groups.", 8000)
            else:
                self.statusbar.showMessage("✨ Restoration Complete", 5000)

    def _legacy_insert_tab(self, hwnd, path, wshell, clipboard, check_cancelled,
                           safe_wait, t_cfg, timeout_base):
        """Keyboard fallback used only when pywinauto is not installed.
        Kept because it is the historical behaviour, but it is unreliable:
        Ctrl+T needs the foreground and the Windows foreground lock can send
        the keystroke to a different window.
        """
        for retry in range(3):
            if check_cancelled(): return False
            try:
                try: win32gui.SetForegroundWindow(hwnd)
                except Exception: pass
                try: wshell.AppActivate(hwnd)
                except Exception: pass
                if safe_wait(t_cfg["app_activate"]): return False

                wshell.SendKeys("^t")
                if safe_wait(max(t_cfg["tab_animation_min"], timeout_base)): return False

                nav_ok = False
                try:
                    for w in win32com.client.Dispatch("Shell.Application").Windows():
                        if w.HWND == hwnd:
                            w.Navigate2(path)
                            nav_ok = True
                            break
                except Exception:
                    pass

                if not nav_ok:
                    try: win32gui.SetForegroundWindow(hwnd)
                    except Exception: pass
                    try: wshell.AppActivate(hwnd)
                    except Exception: pass
                    time.sleep(0.1)
                    wshell.SendKeys("%d")
                    if safe_wait(t_cfg["address_bar_focus"]): return False
                    if not _set_clipboard_text(path): clipboard.setText(path)
                    if safe_wait(t_cfg["clipboard_set"]): return False
                    wshell.SendKeys("^v")
                    if safe_wait(t_cfg["paste_buffer"]): return False
                    wshell.SendKeys("{ENTER}")

                if safe_wait(max(t_cfg["navigation_cooldown_min"], timeout_base)): return False
                return True
            except Exception as e:
                print(f"Legacy tab retry {retry} failed for {path}: {e}")
                if safe_wait(t_cfg["retry_cooldown"]): return False
        return False

    # --- UI SUPPORT METHODS ---

    def update_session_list(self):
        if not hasattr(self, 'sessions_list'): return
        self.sessions_list.clear()
        if not os.path.exists(self.session_dir): return
        
        # Filter: only .json files that start with 'explorer_session_'
        files = [f for f in os.listdir(self.session_dir) 
                 if f.endswith('.json') and f.startswith('explorer_session_')]
        
        files.sort(key=lambda x: os.path.getmtime(os.path.join(self.session_dir, x)), reverse=True)
        
        for f in files:
            try:
                mtime = datetime.fromtimestamp(os.path.getmtime(os.path.join(self.session_dir, f)))
                item = QListWidgetItem(f"📜 {mtime.strftime('%Y-%m-%d %H:%M')} - {f}")
                item.setData(Qt.UserRole, os.path.join(self.session_dir, f))
                self.sessions_list.addItem(item)
            except Exception: continue

    def update_preview(self, item):
        if not hasattr(self, 'preview_tree'): return
        self.preview_tree.clear()
        path = item.data(Qt.UserRole)
        
        try:
            with open(path, 'r', encoding='utf-8') as f:
                data = json.load(f)
                
            folder_icon = QApplication.style().standardIcon(QStyle.SP_DirIcon)
            window_icon = QApplication.style().standardIcon(QStyle.SP_TitleBarMenuButton)
                
            total_items = 0
            for win in data:
                parent = QTreeWidgetItem(self.preview_tree)
                parent.setText(0, f"🖼️ Window {win.get('hwnd', 'Unknown')} ({len(win['paths'])} paths)")
                parent.setIcon(0, window_icon)
                parent.setExpanded(True)
                for p in win['paths']:
                    child = QTreeWidgetItem(parent)
                    child.setText(1, p)
                    name = _path_display_name(p)
                    if _is_valid_path(p):
                        child.setIcon(0, folder_icon)
                        child.setText(0, f"📁 {name}")
                    else:
                        child.setForeground(0, QColor("red"))
                        child.setForeground(1, QColor("red"))
                        child.setText(0, f"🔗💥 {name}")
                    total_items += 1
            
            if hasattr(self, 'item_count_label'):
                self.item_count_label.setText(TEXTS["item_count"].format(count=total_items))
                
        except Exception as e:
            if hasattr(self, 'statusbar'): self.statusbar.showMessage(f"Error reading session: {e}", 5000)

    def filter_preview(self, text):
        if not hasattr(self, 'preview_tree'): return
        text = text.lower()
        for i in range(self.preview_tree.topLevelItemCount()):
            parent = self.preview_tree.topLevelItem(i)
            win_match = False
            for j in range(parent.childCount()):
                child = parent.child(j)
                match = text in child.text(0).lower() or text in child.text(1).lower()
                child.setHidden(not match)
                if match: win_match = True
            parent.setHidden(not win_match)

    def delete_to_trash(self):
        if not hasattr(self, 'sessions_list'): return
        item = self.sessions_list.currentItem()
        if not item: return
        src = item.data(Qt.UserRole)
        dst = os.path.join(self.trash_dir, os.path.basename(src))
        try:
            os.replace(src, dst)
            self.update_session_list()
            self.clear_preview()
        except Exception as e: print(f"Error deleting session: {e}")

    def close_all_explorers(self):
        try:
            shell = win32com.client.Dispatch("Shell.Application")
            for window in shell.Windows():
                try: win32gui.PostMessage(window.HWND, win32con.WM_CLOSE, 0, 0)
                except: pass
        except Exception as e: print(f"Error closing explorers: {e}")

    def change_session_root(self):
        d = QFileDialog.getExistingDirectory(self, "Select Workspace", self.session_dir)
        if d:
            self.session_dir = d
            self.trash_dir = os.path.join(d, ".trash")
            self.ensure_dirs()
            self.settings["session_dir"] = self.session_dir
            self.save_settings()
            self.update_session_list()
            self.clear_preview()

    def show_trash_dialog(self):
        if os.path.exists(self.trash_dir): os.startfile(self.trash_dir)

    # --- AUTO-SAVE ---

    def toggle_auto_save(self):
        """Toggle periodic auto-save on/off."""
        if self.auto_save_timer.isActive():
            self.auto_save_timer.stop()
            self.auto_save_btn.setText(TEXTS["auto_save_off"])
            if hasattr(self, 'statusbar'):
                self.statusbar.showMessage("⏹️ Auto-save stopped", 3000)
        else:
            interval_min = self.settings.get("auto_save_interval", 5)
            self.auto_save_timer.setInterval(interval_min * 60 * 1000)
            self.auto_save_timer.start()
            self.auto_save_btn.setText(TEXTS["auto_save_on"])
            if hasattr(self, 'statusbar'):
                self.statusbar.showMessage(f"⏺️ Auto-save every {interval_min} min", 3000)
            # Save immediately on enable
            self.auto_save_tick()

    def auto_save_tick(self):
        """Called by QTimer — saves current Explorer state."""
        data = self.get_active_explorer_data()
        if not data:
            return  # silent — no windows open, nothing to save
        ts = datetime.now().strftime('%Y-%m-%d_%H-%M-%S')
        filename = f"explorer_session_{ts}.json"
        path = os.path.join(self.session_dir, filename)
        try:
            with open(path, 'w', encoding='utf-8') as f:
                json.dump(data, f, indent=4)
            self.update_session_list()
            label = TEXTS["auto_save_tick"].format(time=datetime.now().strftime('%H:%M:%S'))
            if hasattr(self, 'statusbar'):
                self.statusbar.showMessage(label, 3000)
        except Exception as e:
            print(f"Auto-save error: {e}")

if __name__ == '__main__':
    app = QApplication(sys.argv)
    app.setStyle('Fusion')
    window = MainWindow()
    window.show()
    sys.exit(app.exec())