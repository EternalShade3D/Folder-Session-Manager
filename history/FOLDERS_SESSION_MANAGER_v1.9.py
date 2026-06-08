import os
import sys
import json
import platform
import subprocess
import time
from datetime import datetime

# Windows-specific libraries
try:
    import win32com.client
    import win32gui
    import win32api
    import pywintypes
    import win32con
except ImportError:
    print("Warning: Missing 'pywin32' library. Install it with: pip install pywin32")

from PySide6.QtWidgets import (QApplication, QMainWindow, QTreeWidget, QTreeWidgetItem,
                               QListWidget, QListWidgetItem, QMenu, QMessageBox, QDialog,
                               QVBoxLayout, QHBoxLayout, QPushButton, QFileDialog, QProgressDialog, QStyle)
from PySide6.QtCore import Qt, QObject, Signal, QFile, QIODevice
from PySide6.QtGui import QAction, QColor
from PySide6.QtUiTools import QUiLoader

# ==========================================
# CONFIGURATION
# ==========================================

CONFIG = {
    "sizes": {
        "fixed": (568, 1200),
        "min": (200, 900),
        "tree_col0_width": 300,
        "tree_col1_min_width": 900,
        "preview_min_height": 400
    },
    "max_visible_items": 10,
    "session_dir": "C:\\03_SOFTWARE_WORKSPACE\\FOLDERS_SESSION_MANAGER",
    "trash_dir": "C:\\03_SOFTWARE_WORKSPACE\\FOLDERS_SESSION_MANAGER\\.trash"
}

TEXTS = {
    "app_title": "Folder Session Manager V19",
    "platform_error": "This application only supports Windows OS.",
    "error_msg": "❌ Error: ",
    "error_restoring": "Failed to restore session: {error}",
    "no_paths": "⚠️ No open File Explorer windows found.",
    "invalid_path": "Invalid path: {path}",
    "session_saved": "✅ Session saved: {file}",
    "restored_window": "🔄 Restored window {hwnd}",
    "item_count": "📊 {count} items in session",
}

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
        
        copy_action.triggered.connect(lambda: QApplication.clipboard().setText(item.text(1)))
        open_action.triggered.connect(lambda: subprocess.Popen(['explorer', item.text(1)]))
        
        menu.addAction(copy_action)
        menu.addAction(open_action)
        menu.exec(self.viewport().mapToGlobal(position))

    def on_item_double_clicked(self, item, column):
        if item and item.text(1):
            path = item.text(1)
            if os.path.exists(path):
                subprocess.Popen(['explorer', path])

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

        # Use QFile for more robust loading
        file = QFile(ui_file_path)
        if not file.open(QIODevice.ReadOnly):
            QMessageBox.critical(self, "UI Error", f"Could not open UI file: {ui_file_path}")
            sys.exit(1)
        
        # Load without parent first to avoid QMainWindow re-parenting issues.
        # IMPORTANT: Storing this as an instance variable prevents it and its children 
        # (like sessions_list) from being garbage collected.
        self.ui_container = loader.load(file)
        file.close()

        if not self.ui_container:
            QMessageBox.critical(self, "UI Error", "Failed to load UI object.")
            sys.exit(1)

        # Copy widgets from ui_container to self BEFORE moving the central widget
        # This ensures all children are indexed while the C++ object is definitely alive
        for child in self.ui_container.findChildren(QObject):
            name = child.objectName()
            if name:
                setattr(self, name, child)

        # Move the central widget and status bar to this instance
        if hasattr(self.ui_container, 'centralwidget'):
            self.setCentralWidget(self.ui_container.centralwidget)
        
        if hasattr(self.ui_container, 'statusbar'):
            self.setStatusBar(self.ui_container.statusbar)
            
        self.setWindowTitle(TEXTS["app_title"])
        
        # Setup Tree Headers
        if hasattr(self, 'preview_tree'):
            self.preview_tree.setHeaderLabels(["Name", "Full Path"])
            self.preview_tree.setColumnWidth(0, CONFIG["sizes"]["tree_col0_width"])

        # Set default timeout value if spinbox exists
        if hasattr(self, 'timeout_spin'):
            self.timeout_spin.setValue(self.settings.get("timeout", 10))
            self.timeout_spin.valueChanged.connect(self.on_timeout_changed)

        # Signal Mapping
        if hasattr(self, 'save_button'):
            self.save_button.clicked.connect(self.save_session)
        if hasattr(self, 'restore_button'):
            self.restore_button.clicked.connect(self.restore_session)
        if hasattr(self, 'delete_session_button'):
            self.delete_session_button.clicked.connect(self.delete_to_trash)
        if hasattr(self, 'sessions_list'):
            self.sessions_list.itemClicked.connect(self.update_preview)
        if hasattr(self, 'choose_dir_button'):
            self.choose_dir_button.clicked.connect(self.change_session_root)
        if hasattr(self, 'view_trash_button'):
            self.view_trash_button.clicked.connect(self.show_trash_dialog)
        if hasattr(self, 'search_input'):
            self.search_input.textChanged.connect(self.filter_preview)
        if hasattr(self, 'expand_button'):
            self.expand_button.clicked.connect(lambda: self.preview_tree.expandAll() if hasattr(self, 'preview_tree') else None)
        if hasattr(self, 'collapse_button'):
            self.collapse_button.clicked.connect(lambda: self.preview_tree.collapseAll() if hasattr(self, 'preview_tree') else None)
        if hasattr(self, 'clear_button'):
            self.clear_button.clicked.connect(self.clear_preview)

    def on_timeout_changed(self):
        self.settings["timeout"] = self.timeout_spin.value()
        self.save_settings()

    def clear_preview(self):
        if hasattr(self, 'preview_tree'):
            self.preview_tree.clear()
        if hasattr(self, 'item_count_label'):
            self.item_count_label.setText(TEXTS["item_count"].format(count=0))

    # --- EXPLORER INTERACTION LOGIC ---

    def get_active_explorer_data(self):
        """Captures paths and window positions for all open Explorers."""
        windows = []
        try:
            shell = win32com.client.Dispatch("Shell.Application")
            hwnd_map = {}
            for window in shell.Windows():
                try:
                    path = window.Document.Folder.Self.Path
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
        return windows

    def save_session(self):
        data = self.get_active_explorer_data()
        if not data:
            if hasattr(self, 'statusbar'):
                self.statusbar.showMessage(TEXTS["no_paths"], 5000)
            return

        ts = datetime.now().strftime('%Y-%m-%d_%H-%M-%S')
        filename = f"explorer_session_{ts}.json"
        path = os.path.join(self.session_dir, filename)

        with open(path, 'w', encoding='utf-8') as f:
            json.dump(data, f, indent=4)
        
        self.update_session_list()
        if hasattr(self, 'statusbar'):
            self.statusbar.showMessage(TEXTS["session_saved"].format(file=filename), 5000)

    def restore_session(self):
        """Ported logic from V16: Uses WScript.Shell for Tab Injection."""
        if not hasattr(self, 'sessions_list'): return
        item = self.sessions_list.currentItem()
        if not item: return
        
        file_path = item.data(Qt.UserRole)
        with open(file_path, 'r', encoding='utf-8') as f:
            session_data = json.load(f)

        if hasattr(self, 'close_windows_cb') and self.close_windows_cb.isChecked():
            self.close_all_explorers()

        shell = win32com.client.Dispatch("Shell.Application")
        wshell = win32com.client.Dispatch("WScript.Shell")

        total = sum(len(w['paths']) for w in session_data)
        progress = QProgressDialog("Restoring Session...", "Cancel", 0, total, self)
        progress.setWindowModality(Qt.WindowModal)
        
        count = 0
        timeout_seconds = max(1, self.settings.get("timeout", 10)) / 10.0
        
        for win_data in session_data:
            paths = win_data['paths']
            if not paths: continue

            first_path = paths[0] if os.path.exists(paths[0]) else os.path.expanduser("~")
            shell.Explore(first_path)
            time.sleep(0.8)

            new_hwnd = None
            for _ in range(5):
                for w in shell.Windows():
                    try:
                        if w.Document.Folder.Self.Path == first_path:
                            new_hwnd = w.HWND
                            break
                    except: pass
                if new_hwnd: break
                time.sleep(0.3)

            if new_hwnd:
                try:
                    win32gui.MoveWindow(new_hwnd, win_data['x'], win_data['y'], 
                                       win_data['width'], win_data['height'], True)
                except Exception:
                    pass
                
                for p in paths[1:]:
                    if progress.wasCanceled(): return
                    if os.path.exists(p):
                        try:
                            wshell.AppActivate(new_hwnd)
                            time.sleep(0.2)
                            wshell.SendKeys("^t")
                            time.sleep(0.5)
                            wshell.SendKeys("%d")
                            time.sleep(0.2)
                            wshell.SendKeys(p)
                            wshell.SendKeys("{ENTER}")
                            time.sleep(timeout_seconds)
                        except Exception as e:
                            print(f"Error injecting tab for {p}: {e}")
                    
                    count += 1
                    progress.setValue(count)
            
            count += 1
            progress.setValue(count)
        
        progress.close()
        if hasattr(self, 'statusbar'):
            self.statusbar.showMessage("✅ Restoration Complete", 5000)

    # --- UI SUPPORT METHODS ---

    def update_session_list(self):
        if not hasattr(self, 'sessions_list'): return
        self.sessions_list.clear()
        if not os.path.exists(self.session_dir): return
        
        files = [f for f in os.listdir(self.session_dir) if f.endswith('.json')]
        files.sort(key=lambda x: os.path.getmtime(os.path.join(self.session_dir, x)), reverse=True)
        
        for f in files:
            try:
                mtime = datetime.fromtimestamp(os.path.getmtime(os.path.join(self.session_dir, f)))
                item = QListWidgetItem(f"{mtime.strftime('%Y-%m-%d %H:%M')} - {f}")
                item.setData(Qt.UserRole, os.path.join(self.session_dir, f))
                self.sessions_list.addItem(item)
            except Exception:
                continue

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
                parent.setText(0, f"Window {win.get('hwnd', 'Unknown')} ({len(win['paths'])} paths)")
                parent.setIcon(0, window_icon)
                parent.setExpanded(True)
                for p in win['paths']:
                    child = QTreeWidgetItem(parent)
                    child.setText(0, os.path.basename(p) or p)
                    child.setText(1, p)
                    child.setIcon(0, folder_icon)
                    if not os.path.exists(p):
                        child.setForeground(0, QColor("red"))
                        child.setForeground(1, QColor("red"))
                    total_items += 1
            
            if hasattr(self, 'item_count_label'):
                self.item_count_label.setText(TEXTS["item_count"].format(count=total_items))
                
        except Exception as e:
            if hasattr(self, 'statusbar'):
                self.statusbar.showMessage(f"Error reading session: {e}", 5000)

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
        except Exception as e:
            print(f"Error deleting session: {e}")

    def close_all_explorers(self):
        try:
            shell = win32com.client.Dispatch("Shell.Application")
            for window in shell.Windows():
                try:
                    win32gui.PostMessage(window.HWND, win32con.WM_CLOSE, 0, 0)
                except: pass
        except Exception as e:
            print(f"Error closing explorers: {e}")

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
        if os.path.exists(self.trash_dir):
            os.startfile(self.trash_dir)

if __name__ == '__main__':
    app = QApplication(sys.argv)
    app.setStyle('Fusion')
    window = MainWindow()
    window.show()
    sys.exit(app.exec())