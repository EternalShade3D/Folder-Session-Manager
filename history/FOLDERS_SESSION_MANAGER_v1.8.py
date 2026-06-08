import os
import sys
import json
import platform
import subprocess
import time
from datetime import datetime
import win32com.client
import win32gui
import win32api
import pywintypes
import win32con
from PySide6.QtWidgets import (QApplication, QMainWindow, QTreeWidget, QTreeWidgetItem,
                               QListWidget, QListWidgetItem, QMenu, QMessageBox, QDialog,
                               QVBoxLayout, QHBoxLayout, QPushButton, QFileDialog, QProgressDialog)
from PySide6.QtCore import Qt, QObject, Signal
from PySide6.QtGui import QAction, QColor
from PySide6.QtUiTools import QUiLoader

CONFIG = {
    "sizes": {
        "fixed": (568, 721),
        "min": (200, 300),
        "font": (12, 16),
        "padding_prop": 0.02,
        "button_min_width": 100,
        "tree_col0_width": 300,
        "tree_col1_min_width": 900,
        "preview_min_height": 400
    },
    "max_visible_items": 10,
    "session_dir": "C:\\03_SOFTWARE_WORKSPACE",
    "trash_dir": "C:\\03_SOFTWARE_WORKSPACE\\.trash"
}

TEXTS = {
    "app_title": "Folder Session Manager",
    "platform_error": "This application only supports Windows.",
    "error_msg": "❌ Error: ",
    "error_restoring": "Failed to restore session: {error}",
    "no_paths": "⚠️ No open File Explorer windows found.",
    "invalid_path": "Invalid path: {path}",
    "session_saved": "✅ Session saved: {file}",
    "restored_window": "🔄 Restored window {hwnd}",
    "title_label": '<html><head/><body><p><span style=" font-size:14pt; font-weight:700;">FOLDER SESSION MANAGER</span></p></body></html>',
    "sessions_label": '<html><head/><body><p><span style=" font-size:12pt; font-weight:700;">📜 SESSIONS</span></p></body></html>',
    "choose_session_dir": "📂 Choose Session Folder",
    "view_trash": "♻️ View Trash",
    "close_windows_cb": "🛑 Close All Windows",
    "timeout_label": "⏱️ Timeout (s):",
    "preview_label": '<html><head/><body><p><span style=" font-size:12pt; font-weight:700;">🔎 SESSION PREVIEW</span></p></body></html>',
    "search_placeholder": "🔍 Search paths...",
    "item_count": "📊 {count} items in session",
    "expand_all": "⬆️ Expand All",
    "collapse_all": "⬇️ Collapse All",
    "clear_button": "🗑️ Clear Preview",
    "save_button": "💾 Save Session",
    "restore_button": "🔄 Restore Session",
    "delete_session_button": "🗑️ Delete Session"
}

class RestoreHelper:
    window_opened = Signal(int)  # Signal for window handle when opened

    def __init__(self, shell, statusbar=None):
        self.shell = shell
        self.statusbar = statusbar
        self.timeout = 10  # Max seconds to wait for conditions

    def close_all_explorer_windows(self):
        closed_hwnds = set()
        for window in self.shell.Windows():
            hwnd = window.HWND
            if hwnd not in closed_hwnds and win32gui.IsWindow(hwnd):
                try:
                    win32api.PostMessage(hwnd, win32con.WM_CLOSE, 0, 0)
                    closed_hwnds.add(hwnd)
                    print(f"Sent WM_CLOSE to window HWND: {hwnd}")
                    time.sleep(0.1)  # Brief delay to allow closure
                except Exception as e:
                    print(f"Error sending WM_CLOSE to HWND: {hwnd}, {e}")
                    if self.statusbar:
                        self.statusbar.showMessage(f"{TEXTS['error_msg']}Error closing window: {e}", 5000)
        if not closed_hwnds:
            try:
                subprocess.Popen(['taskkill', '/IM', 'explorer.exe', '/F'])
                print("Forced all Explorer instances to close via taskkill")
                self._wait_for_explorer_restart()
            except Exception as e:
                print(f"Error forcing Explorer close: {e}")
                if self.statusbar:
                    self.statusbar.showMessage(f"{TEXTS['error_msg']}Failed to force close Explorer: {e}", 5000)

    def _wait_for_explorer_restart(self):
        start_time = time.time()
        while time.time() - start_time < 5:  # Wait up to 5 seconds
            if any("explorer.exe" in win32process.GetModuleFileNameEx(h, 0) for h in win32process.EnumProcesses()):
                break
            time.sleep(0.5)
        print("Explorer restart detected or timeout")

    def open_explorer_window(self, path, x, y, width, height):
        self.shell.Explore(path)
        start_time = time.time()
        while time.time() - start_time < self.timeout:
            for window in self.shell.Windows():
                if window.Document.Folder.Self.Path == path and "File Explorer" in win32gui.GetWindowText(window.HWND):
                    hwnd = window.HWND
                    if win32gui.IsWindowVisible(hwnd) and win32gui.IsWindowEnabled(hwnd):
                        win32gui.MoveWindow(hwnd, x, y, width, height, True)
                        self.window_opened.emit(hwnd)
                        return hwnd
            time.sleep(0.1)
        print(f"Timeout waiting for window at {path}")
        return None

    def add_tabs_to_window(self, hwnd, paths):
        target_window = next((w for w in self.shell.Windows() if w.HWND == hwnd), None)
        if not target_window:
            return False
        for path in paths:
            if os.path.exists(path) and path not in [w.Document.Folder.Self.Path for w in self.shell.Windows() if w.HWND == hwnd]:
                start_time = time.time()
                while time.time() - start_time < self.timeout:
                    try:
                        target_window.Navigate(path)
                        time.sleep(0.1)  # Brief delay for navigation
                        if path == target_window.Document.Folder.Self.Path:
                            print(f"Added tab for path {path}")
                            return True
                    except Exception as e:
                        print(f"Navigation error for {path}: {e}")
                        time.sleep(0.1)
                print(f"Failed to navigate to {path} after timeout")
        return False
        
class CustomTreeWidget(QTreeWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setContextMenuPolicy(Qt.CustomContextMenu)
        self.customContextMenuRequested.connect(self.show_context_menu)

    def show_context_menu(self, position):
        item = self.itemAt(position)
        if not item:
            return
        menu = QMenu(self)
        copy_action = QAction("📋 Copy Path", self)
        open_action = QAction("📂 Open Folder", self)
        copy_action.triggered.connect(lambda: self.copy_to_clipboard(item))
        open_action.triggered.connect(lambda: self.open_folder(item))
        menu.addAction(copy_action)
        menu.addAction(open_action)
        menu.exec(self.viewport().mapToGlobal(position))

    def copy_to_clipboard(self, item):
        if item and item.text(1):
            QApplication.clipboard().setText(item.text(1))
            if hasattr(self, 'statusbar'):
                self.statusbar.showMessage(f"📋 Copied: {item.text(1)}", 5000)
            else:
                print(f"Copied: {item.text(1)}")

    def open_folder(self, item):
        if item and item.text(1) and os.path.exists(item.text(1)):
            subprocess.Popen(['explorer', item.text(1)])
        else:
            if hasattr(self, 'statusbar'):
                self.statusbar.showMessage(f"{TEXTS['error_msg']}Invalid path: {item.text(1)}", 5000)
            else:
                print(f"Invalid path: {item.text(1)}")

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        if platform.system() != "Windows":
            QMessageBox.critical(self, TEXTS["error_msg"], TEXTS["platform_error"])
            sys.exit(1)
        self.settings_file = os.path.join(CONFIG["session_dir"], "settings.json")
        self.load_settings()
        self.session_dir = self.settings.get("session_dir", CONFIG["session_dir"])
        self.trash_dir = os.path.join(self.session_dir, ".trash")
        try:
            os.makedirs(self.session_dir, exist_ok=True)
            os.makedirs(self.trash_dir, exist_ok=True)
            if not os.access(self.session_dir, os.W_OK):
                raise OSError(f"No write permission for {self.session_dir}")
            if not os.access(self.trash_dir, os.W_OK):
                raise OSError(f"No write permission for {self.trash_dir}")
            print(f"Session directory set to: {self.session_dir}")
            print(f"Trash directory set to: {self.trash_dir}")
        except (OSError, PermissionError) as e:
            print(f"Error creating directories: {e}")
            QMessageBox.critical(self, TEXTS["error_msg"], f"{TEXTS['error_msg']}Failed to create session directory: {e}")
            sys.exit(1)
        self.setWindowTitle(TEXTS["app_title"])
        self.setMinimumSize(*CONFIG["sizes"]["min"])
        self.resize(*CONFIG["sizes"]["fixed"])
        self.setToolTipDuration(-1)
        self.setup_ui()
        self.update_session_list()

    def load_settings(self):
        try:
            if os.path.exists(self.settings_file):
                with open(self.settings_file, 'r', encoding='utf-8') as f:
                    self.settings = json.load(f)
            else:
                self.settings = {"session_dir": CONFIG["session_dir"], "timeout": 1}
        except (OSError, json.JSONDecodeError) as e:
            print(f"Error loading settings: {e}")
            self.settings = {"session_dir": CONFIG["session_dir"], "timeout": 1}

    def save_settings(self):
        try:
            with open(self.settings_file, 'w', encoding='utf-8') as f:
                json.dump(self.settings, f, indent=4)
            print(f"Settings saved to: {self.settings_file}")
        except OSError as e:
            print(f"Error saving settings: {e}")
            QMessageBox.critical(self, TEXTS["error_msg"], f"{TEXTS['error_msg']}Failed to save settings: {e}")

    def setup_ui(self):
        loader = QUiLoader()
        loader.registerCustomWidget(CustomTreeWidget)
        path_to_ui = os.path.join(os.path.dirname(os.path.abspath(__file__)), "session_manager.ui")
        
        if not os.path.exists(path_to_ui):
            QMessageBox.critical(self, "UI Error", f"UI file not found: {path_to_ui}")
            sys.exit(1)

        try:
            loaded_ui = loader.load(path_to_ui, self)
            print(f"UI loaded from: {path_to_ui}")
            if hasattr(loaded_ui, 'centralwidget'):
                self.setCentralWidget(loaded_ui.centralwidget)
                print("Central widget set.")
            else:
                QMessageBox.critical(self, "UI Error", "Central widget not found. Check the .ui file.")
                sys.exit(1)

            if hasattr(loaded_ui, 'statusbar'):
                self.setStatusBar(loaded_ui.statusbar)
                print("Status bar set.")
            
            widget_names = []
            for child in self.findChildren(QObject):
                if child.objectName():
                    setattr(self, child.objectName(), child)
                    widget_names.append(child.objectName())
            print("Widgets loaded:", widget_names)

        except Exception as e:
            QMessageBox.critical(self, "UI Load Error", f"Failed to load UI file: {path_to_ui}\nError: {e}")
            sys.exit(1)

        self.setStyleSheet("""
            QTreeWidget::item:alternate { background: #333333; }
            QTreeWidget::item:selected { background: #1a1a1a; color: white; }
        """)
        if hasattr(self, 'preview_tree'):
            self.preview_tree.setColumnWidth(0, CONFIG["sizes"]["tree_col0_width"])
            self.preview_tree.setColumnWidth(1, CONFIG["sizes"]["tree_col1_min_width"])

        widget_text_mapping = {
            'tlabel': 'title_label',
            'sessions_label': 'sessions_label',
            'choose_dir_button': 'choose_session_dir',
            'view_trash_button': 'view_trash',
            'delete_session_button': 'delete_session_button',
            'close_windows_cb': 'close_windows_cb',
            'timeout_label': 'timeout_label',
            'preview_label': 'preview_label',
            'search_input': 'search_placeholder',
            'item_count_label': 'item_count',
            'expand_button': 'expand_all',
            'collapse_button': 'collapse_all',
            'clear_button': 'clear_button',
            'save_button': 'save_button',
            'restore_button': 'restore_button'
        }
        for widget_name, text_key in widget_text_mapping.items():
            if hasattr(self, widget_name):
                widget = getattr(self, widget_name)
                print(f"Setting text for {widget_name}: {TEXTS[text_key]}")
                if text_key == 'search_placeholder':
                    widget.setPlaceholderText(TEXTS[text_key])
                elif text_key == 'item_count':
                    widget.setText(TEXTS[text_key].format(count=0))
                else:
                    widget.setText(TEXTS[text_key])
            else:
                print(f"Widget {widget_name} not found in self!")

        if hasattr(self, 'theme_label'):
            self.theme_label.setText("")
        if hasattr(self, 'theme_combo'):
            self.theme_combo.hide()

        if hasattr(self, 'preview_tree'):
            print("Type of preview_tree:", type(self.preview_tree))
            self.preview_tree.setHeaderHidden(True)
            self.preview_tree.setColumnCount(2)
            header_item = QTreeWidgetItem()
            header_item.setText(0, "Name")
            header_item.setText(1, "Path")
            self.preview_tree.setHeaderItem(header_item)

        if hasattr(self, 'save_button'):
            self.save_button.setEnabled(True)
            self.save_button.clicked.connect(lambda: print("Save button clicked") or self.save_session())
        if hasattr(self, 'delete_session_button'):
            self.delete_session_button.setEnabled(True)
            self.delete_session_button.clicked.connect(lambda: print("Delete button clicked") or self.delete_session())
        if hasattr(self, 'sessions_list'):
            self.sessions_list.itemClicked.connect(self.update_preview)
            self.sessions_list.setContextMenuPolicy(Qt.CustomContextMenu)
            self.sessions_list.customContextMenuRequested.connect(self.show_session_context_menu)
        if hasattr(self, 'close_windows_cb'):
            self.close_windows_cb.stateChanged.connect(self.confirm_close_windows)
        if hasattr(self, 'timeout_spin'):
            self.timeout_spin.setMinimum(0)
            self.timeout_spin.setMaximum(100)
            self.timeout_spin.setValue(self.settings.get("timeout", 1))
            self.timeout_spin.valueChanged.connect(lambda: self.settings.update({"timeout": self.timeout_spin.value()}) or self.save_settings())
        if hasattr(self, 'choose_dir_button'):
            self.choose_dir_button.clicked.connect(self.choose_session_dir)
        if hasattr(self, 'view_trash_button'):
            self.view_trash_button.clicked.connect(self.view_trash)
        if hasattr(self, 'clear_button'):
            self.clear_button.clicked.connect(self.clear_preview)
        if hasattr(self, 'search_input'):
            self.search_input.textChanged.connect(self.filter_preview)
        if hasattr(self, 'restore_button'):
            self.restore_button.clicked.connect(lambda: print("Restore button clicked") or self.restore_session())

    def get_explorer_windows(self):
        windows_data = []
        try:
            shell = win32com.client.Dispatch("Shell.Application")
            path_by_hwnd = {}
            window_count = 0
            for window in shell.Windows():
                try:
                    if hasattr(window, 'Document') and hasattr(window.Document, 'Folder'):
                        hwnd = window.HWND
                        path = window.Document.Folder.Self.Path
                        path_by_hwnd.setdefault(hwnd, []).append(path)
                        window_count += 1
                except Exception as e:
                    if hasattr(self, 'statusbar'):
                        self.statusbar.showMessage(f"{TEXTS['error_msg']}Error getting window: {e}", 5000)
                    continue
            print(f"Detected {window_count} Explorer windows")
            for hwnd, paths in path_by_hwnd.items():
                try:
                    left, top, right, bottom = win32gui.GetWindowRect(hwnd)
                    windows_data.append({
                        'hwnd': hwnd,
                        'paths': paths,
                        'x': left,
                        'y': top,
                        'width': right - left,
                        'height': bottom - top
                    })
                    print(f"Window {hwnd}: {paths}, Position: ({left}, {top}, {right}, {bottom})")
                except Exception as e:
                    if hasattr(self, 'statusbar'):
                        self.statusbar.showMessage(f"{TEXTS['error_msg']}Error getting window position: {e}", 5000)
                    continue
            print(f"Saving {len(windows_data)} windows to session")
            if not windows_data:
                if hasattr(self, 'statusbar'):
                    self.statusbar.showMessage(TEXTS["no_paths"], 5000)
        except Exception as e:
            if hasattr(self, 'statusbar'):
                self.statusbar.showMessage(f"{TEXTS['error_msg']}Failed to access Shell: {e}", 5000)
        return windows_data

    def save_session(self):
        if not hasattr(self, 'save_button'):
            print("Save button not found")
            return
        self.save_button.setEnabled(False)
        timestamp = datetime.now().strftime('%Y-%m-%d_%H-%M-%S')
        file_name = os.path.join(self.session_dir, f"explorer_session_{timestamp}.json")
        backup_name = os.path.join(self.trash_dir, f"backup_explorer_session_{timestamp}.json")
        try:
            if not os.access(self.session_dir, os.W_OK):
                raise OSError(f"No write permission for {self.session_dir}")
            if not os.access(self.trash_dir, os.W_OK):
                raise OSError(f"No write permission for {self.trash_dir}")
            
            windows_data = self.get_explorer_windows()
            if not windows_data:
                if hasattr(self, 'statusbar'):
                    self.statusbar.showMessage(TEXTS["no_paths"], 5000)
                print("No windows to save")
                return

            print(f"Saving session to {file_name}")
            with open(file_name, 'w', encoding='utf-8') as f:
                json.dump(windows_data, f, indent=4)
            print(f"Session file created: {os.path.exists(file_name)}")
            
            print(f"Saving backup to {backup_name}")
            with open(backup_name, 'w', encoding='utf-8') as f_backup:
                json.dump(windows_data, f_backup, indent=4)
            print(f"Backup file created: {os.path.exists(backup_name)}")
            
            self.update_session_list()
            if hasattr(self, 'sessions_list') and self.sessions_list.count() > 0:
                self.sessions_list.setCurrentRow(0)
                self.sessions_list.scrollToItem(self.sessions_list.item(0))
            if hasattr(self, 'statusbar'):
                self.statusbar.showMessage(TEXTS["session_saved"].format(file=os.path.basename(file_name)), 5000)
        except (OSError, PermissionError) as e:
            print(f"Error saving session: {e}")
            QMessageBox.critical(self, TEXTS["error_msg"], f"{TEXTS['error_msg']}Failed to save session: {e}")
        finally:
            self.save_button.setEnabled(True)

    def update_session_list(self):
        if not hasattr(self, 'sessions_list'):
            print("sessions_list widget not found")
            return
        self.sessions_list.clear()
        try:
            all_sessions = []
            directory = self.session_dir
            print(f"Scanning directory: {directory}")
            if not os.path.exists(directory):
                print(f"Directory does not exist: {directory}")
                if hasattr(self, 'statusbar'):
                    self.statusbar.showMessage(f"{TEXTS['error_msg']}Session directory not found: {directory}", 5000)
                return
            for f in sorted(os.listdir(directory), key=lambda x: os.path.getmtime(os.path.join(directory, x)), reverse=True):
                if f.endswith('.json') and f.startswith('explorer_session'):
                    try:
                        file_path = os.path.join(directory, f)
                        with open(file_path, 'r', encoding='utf-8') as file:
                            json.load(file)
                        mtime = datetime.fromtimestamp(os.path.getmtime(file_path))
                        all_sessions.append((mtime, f, file_path))
                        print(f"Found session: {f}, mtime: {mtime}")
                    except (json.JSONDecodeError, OSError) as e:
                        print(f"Skipping invalid file {f}: {e}")
                        if hasattr(self, 'statusbar'):
                            self.statusbar.showMessage(f"{TEXTS['error_msg']}Skipping invalid JSON file: {f}", 5000)
            for mtime, f, full_path in sorted(all_sessions, key=lambda x: x[0], reverse=True):
                item = QListWidgetItem(f"{mtime.strftime('%Y-%m-%d %H:%M')} - {f}")
                item.setData(Qt.UserRole, full_path)
                self.sessions_list.addItem(item)
                print(f"Added to sessions_list: {mtime.strftime('%Y-%m-%d %H:%M')} - {f}")
            if not all_sessions:
                print("No valid session files found")
                if hasattr(self, 'statusbar'):
                    self.statusbar.showMessage(TEXTS["no_paths"], 5000)
        except OSError as e:
            print(f"Error updating session list: {e}")
            QMessageBox.critical(self, TEXTS["error_msg"], f"{TEXTS['error_msg']}{e}")

    def view_trash(self):
        if not os.path.exists(self.trash_dir):
            print(f"Trash directory does not exist: {self.trash_dir}")
            QMessageBox.information(self, TEXTS["app_title"], "No trashed sessions found.")
            return
        dialog = QDialog(self)
        dialog.setWindowTitle("🗑️ Trashed Sessions")
        dialog.setMinimumSize(400, 300)
        layout = QVBoxLayout(dialog)
        
        trash_list = QListWidget()
        layout.addWidget(trash_list)
        
        try:
            all_sessions = []
            for f in sorted(os.listdir(self.trash_dir), key=lambda x: os.path.getmtime(os.path.join(self.trash_dir, x)), reverse=True):
                if f.endswith('.json') and f.startswith('explorer_session'):
                    try:
                        file_path = os.path.join(self.trash_dir, f)
                        with open(file_path, 'r', encoding='utf-8') as file:
                            json.load(file)
                        mtime = datetime.fromtimestamp(os.path.getmtime(file_path))
                        all_sessions.append((mtime, f, file_path))
                        print(f"Found trashed session: {f}, mtime: {mtime}")
                    except (json.JSONDecodeError, OSError) as e:
                        print(f"Skipping invalid trashed file {f}: {e}")
                        continue
            for mtime, f, full_path in sorted(all_sessions, key=lambda x: x[0], reverse=True):
                item = QListWidgetItem(f"{mtime.strftime('%Y-%m-%d %H:%M')} - {f}")
                item.setData(Qt.UserRole, full_path)
                trash_list.addItem(item)
                print(f"Added to trash_list: {mtime.strftime('%Y-%m-%d %H:%M')} - {f}")
            
            button_layout = QHBoxLayout()
            restore_button = QPushButton("🔄 Restore Selected")
            delete_button = QPushButton("🗑️ Delete Permanently")
            button_layout.addWidget(restore_button)
            button_layout.addWidget(delete_button)
            layout.addLayout(button_layout)
            
            def restore_selected():
                selected_items = trash_list.selectedItems()
                if not selected_items:
                    QMessageBox.warning(dialog, TEXTS["app_title"], "No session selected.")
                    return
                for item in selected_items:
                    file_path = item.data(Qt.UserRole)
                    dest_path = os.path.join(self.session_dir, os.path.basename(file_path))
                    try:
                        os.rename(file_path, dest_path)
                        print(f"Restored session: {file_path} to {dest_path}")
                    except OSError as e:
                        print(f"Error restoring session {file_path}: {e}")
                        QMessageBox.critical(dialog, TEXTS["error_msg"], f"{TEXTS['error_msg']}Failed to restore: {e}")
                dialog.accept()
                self.update_session_list()
            
            def delete_permanently():
                selected_items = trash_list.selectedItems()
                if not selected_items:
                    QMessageBox.warning(dialog, TEXTS["app_title"], "No session selected.")
                    return
                reply = QMessageBox.question(dialog, TEXTS["app_title"], "Permanently delete selected sessions?",
                                             QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
                if reply == QMessageBox.Yes:
                    for item in selected_items:
                        file_path = item.data(Qt.UserRole)
                        try:
                            os.remove(file_path)
                            print(f"Permanently deleted session: {file_path}")
                        except OSError as e:
                            print(f"Error deleting session {file_path}: {e}")
                            QMessageBox.critical(dialog, TEXTS["error_msg"], f"{TEXTS['error_msg']}Failed to delete: {e}")
                    dialog.accept()
                    self.update_session_list()
            
            restore_button.clicked.connect(restore_selected)
            delete_button.clicked.connect(delete_permanently)
            
            if not all_sessions:
                print("No valid trashed sessions found")
                QMessageBox.information(dialog, TEXTS["app_title"], "No trashed sessions found.")
                dialog.reject()
                return
        except OSError as e:
            print(f"Error accessing trash directory: {e}")
            QMessageBox.critical(dialog, TEXTS["error_msg"], f"{TEXTS['error_msg']}{e}")
            dialog.reject()
            return
        
        dialog.exec()

    def delete_session(self):
        if not hasattr(self, 'sessions_list'):
            print("sessions_list widget not found")
            return
        selected_items = self.sessions_list.selectedItems()
        if not selected_items:
            QMessageBox.warning(self, TEXTS["app_title"], "No session selected.")
            return
        reply = QMessageBox.question(self, TEXTS["app_title"], "Move selected sessions to trash?",
                                     QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
        if reply == QMessageBox.Yes:
            try:
                for item in selected_items:
                    file_path = item.data(Qt.UserRole)
                    dest_path = os.path.join(self.trash_dir, os.path.basename(file_path))
                    os.rename(file_path, dest_path)
                    print(f"Moved session to trash: {file_path} -> {dest_path}")
                self.update_session_list()
                if hasattr(self, 'statusbar'):
                    self.statusbar.showMessage("🗑️ Selected sessions moved to trash.", 5000)
            except OSError as e:
                print(f"Error moving session to trash: {e}")
                QMessageBox.critical(self, TEXTS["error_msg"], f"{TEXTS['error_msg']}Failed to move session: {e}")

    def update_preview(self, item):
        if not hasattr(self, 'preview_tree'):
            return
        self.preview_tree.clear()
        if hasattr(self, 'search_input'):
            self.search_input.clear()
        if hasattr(self, 'item_count_label'):
            self.item_count_label.setText(TEXTS["item_count"].format(count=0))
        if not item:
            return
        file_path = item.data(Qt.UserRole) if item else None
        if not file_path or not os.path.exists(file_path):
            return
        try:
            with open(file_path, 'r', encoding='utf-8') as f:
                data = json.load(f)
                item_count = 0
                for entry in data:
                    hwnd, paths = entry['hwnd'], entry['paths']
                    window_item = QTreeWidgetItem(self.preview_tree)
                    window_item.setText(0, f"Window {hwnd} ({len(paths)} paths)")
                    window_item.setExpanded(True)  # Expand all window items by default
                    item_count += 1
                    for path in paths:
                        path_item = QTreeWidgetItem(window_item)
                        path_item.setText(0, f"{os.path.basename(path)}")
                        path_item.setText(1, path)
                        item_count += 1
                        if not os.path.exists(path):
                            path_item.setForeground(1, QColor("red"))
                            path_item.setToolTip(1, TEXTS["invalid_path"].format(path=path))
                        if window_item.childCount() % 2 == 0:
                            path_item.setBackground(0, QColor("#333333"))
                            path_item.setBackground(1, QColor("#333333"))
                    if hasattr(self, 'item_count_label'):
                        self.item_count_label.setText(TEXTS["item_count"].format(count=item_count))
        except (OSError, json.JSONDecodeError) as e:
            QMessageBox.critical(self, TEXTS["error_msg"], f"{TEXTS['error_restoring'].format(error=str(e))}")
            if hasattr(self, 'item_count_label'):
                self.item_count_label.setText(TEXTS["item_count"].format(count=0))

    def choose_session_dir(self):
        directory = QFileDialog.getExistingDirectory(self, "Choose Session Directory", self.session_dir)
        if directory:
            self.session_dir = directory
            self.trash_dir = os.path.join(self.session_dir, ".trash")
            try:
                os.makedirs(self.trash_dir, exist_ok=True)
                self.settings["session_dir"] = self.session_dir
                self.save_settings()
                print(f"Updated session directory to: {self.session_dir}")
                print(f"Updated trash directory to: {self.trash_dir}")
                self.update_session_list()
            except (OSError, PermissionError) as e:
                print(f"Error setting session directory: {e}")
                QMessageBox.critical(self, TEXTS["error_msg"], f"{TEXTS['error_msg']}Failed to set session directory: {e}")

    def show_session_context_menu(self, position):
        item = self.sessions_list.itemAt(position)
        if not item:
            return
        menu = QMenu(self)
        open_dir_action = QAction("📂 Open Session Directory", self)
        delete_action = QAction("🗑️ Delete Session", self)
        open_dir_action.triggered.connect(lambda: self.open_session_directory(item))
        delete_action.triggered.connect(lambda: self.delete_session())
        menu.addAction(open_dir_action)
        menu.addAction(delete_action)
        menu.exec(self.sessions_list.viewport().mapToGlobal(position))

    def open_session_directory(self, item):
        file_path = item.data(Qt.UserRole)
        if file_path and os.path.exists(file_path):
            directory = os.path.dirname(file_path)
            subprocess.Popen(['explorer', directory])
            if hasattr(self, 'statusbar'):
                self.statusbar.showMessage(f"📂 Opened directory: {directory}", 5000)
        else:
            if hasattr(self, 'statusbar'):
                self.statusbar.showMessage(f"{TEXTS['error_msg']}Invalid session file: {file_path}", 5000)

    def copy_to_clipboard(self, item):
        if item and item.text(1):
            QApplication.clipboard().setText(item.text(1))
            if hasattr(self, 'statusbar'):
                self.statusbar.showMessage(f"📋 Copied: {item.text(1)}", 5000)

    def open_folder(self, item):
        if item and item.text(1) and os.path.exists(item.text(1)):
            subprocess.Popen(['explorer', item.text(1)])
        else:
            if hasattr(self, 'statusbar'):
                self.statusbar.showMessage(f"{TEXTS['error_msg']}Invalid path: {item.text(1)}", 5000)

    def confirm_close_windows(self, state):
        if state == Qt.Checked:
            reply = QMessageBox.question(self, TEXTS["app_title"], "Close all File Explorer windows on restore?",
                                         QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
            if reply == QMessageBox.No:
                self.close_windows_cb.setChecked(False)

    def clear_preview(self):
        if hasattr(self, 'preview_tree'):
            self.preview_tree.clear()
        if hasattr(self, 'search_input'):
            self.search_input.clear()
        if hasattr(self, 'item_count_label'):
            self.item_count_label.setText(TEXTS["item_count"].format(count=0))

    def filter_preview(self, text):
        if not hasattr(self, 'preview_tree'):
            return
        text = text.lower()
        for i in range(self.preview_tree.topLevelItemCount()):
            window_item = self.preview_tree.topLevelItem(i)
            window_visible = False
            for j in range(window_item.childCount()):
                path_item = window_item.child(j)
                path = path_item.text(1).lower()
                name = path_item.text(0).lower()
                is_visible = text in path or text in name
                path_item.setHidden(not is_visible)
                if is_visible:
                    window_visible = True
            window_item.setHidden(not window_visible)
            window_item.setExpanded(window_visible)

    def restore_session(self):
            if not hasattr(self, 'sessions_list'):
                print("sessions_list widget not found")
                return
            if item := self.sessions_list.currentItem():
                file_path = item.data(Qt.UserRole)
                try:
                    with open(file_path, 'r', encoding='utf-8') as f:
                        session_data = json.load(f)
                    shell = win32com.client.Dispatch("Shell.Application")
                    helper = RestoreHelper(shell, self.statusbar if hasattr(self, 'statusbar') else None)

                    # Close existing windows if requested
                    if self.close_windows_cb.isChecked():
                        helper.close_all_explorer_windows()

                    # Setup progress overlay
                    total_paths = sum(len(entry['paths']) for entry in session_data)
                    progress = QProgressDialog("Restoring session...", None, 0, total_paths, self)
                    progress.setWindowModality(Qt.WindowModal)
                    progress.setWindowFlags(progress.windowFlags() | Qt.WindowStaysOnTopHint | Qt.WindowMinimizeButtonHint)
                    progress.setCancelButton(None)
                    progress.show()
                    QApplication.processEvents()
                    current_path_idx = 0

                    for entry in session_data:
                        hwnd, paths, x, y, width, height = entry['hwnd'], entry['paths'], entry['x'], entry['y'], entry['width'], entry['height']
                        if not paths:
                            continue
                        
                        initial_path = next((p for p in paths if os.path.exists(p)), os.path.expanduser("~"))
                        explorer_hwnd = helper.open_explorer_window(initial_path, x, y, width, height)
                        if explorer_hwnd:
                            for path in paths[1:]:
                                if os.path.exists(path):
                                    if helper.add_tabs_to_window(explorer_hwnd, [path]):
                                        if hasattr(self, 'statusbar'):
                                            self.statusbar.showMessage(
                                                f"{TEXTS['restored_window'].format(hwnd=hwnd)} in tab", 3000)
                                        print(f"Restored window {hwnd} for path {path}")
                                        current_path_idx += 1
                                        progress.setValue(current_path_idx)
                                        progress.setLabelText(f"Restoring: {os.path.basename(path)} ({current_path_idx}/{total_paths})")
                                        QApplication.processEvents()
                                else:
                                    if hasattr(self, 'statusbar'):
                                        self.statusbar.showMessage(TEXTS["invalid_path"].format(path=path), 5000)
                                    print(f"Invalid path: {path}")
                    progress.setValue(total_paths)
                    progress.close()
                    if hasattr(self, 'statusbar'):
                        self.statusbar.showMessage("🔄 Session restored successfully.", 5000)
                except (OSError, json.JSONDecodeError) as e:
                    QMessageBox.critical(self, TEXTS["error_msg"], f"{TEXTS['error_restoring'].format(error=str(e))}")
                    print(f"Error restoring session: {e}")

if __name__ == '__main__':
    app = QApplication(sys.argv)
    app.setStyle('Fusion')
    window = MainWindow()
    window.show()
    sys.exit(app.exec())
