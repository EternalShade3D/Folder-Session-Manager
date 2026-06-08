import sys
import os
import json
import time
from datetime import datetime
import win32gui
import win32api
import win32com.client
import platform
import unicodedata
import shutil
from PySide6.QtCore import (
    QCoreApplication, QDate, QDateTime, QLocale,
    QMetaObject, QObject, QPoint, QRect,
    QSize, QTime, QUrl, Qt, QFile, Signal, QSettings, QEvent
)
from PySide6.QtGui import (
    QBrush, QColor, QConicalGradient, QCursor,
    QFont, QFontDatabase, QGradient, QIcon,
    QImage, QKeySequence, QLinearGradient, QPainter,
    QPalette, QPixmap, QRadialGradient, QTransform
)
from PySide6.QtWidgets import (
    QApplication, QCheckBox, QComboBox, QFrame,
    QHBoxLayout, QHeaderView, QLabel, QLineEdit,
    QListWidget, QListWidgetItem, QMainWindow, QPushButton,
    QSizePolicy, QSpinBox, QStatusBar,
    QTreeWidget, QTreeWidgetItem, QVBoxLayout, QWidget,
    QProgressDialog, QMessageBox, QFileDialog, QToolTip,
    QSpacerItem, QMenu
)

# IMPORTANTE: Importar QUiLoader para carregar o .ui diretamente
from PySide6.QtUiTools import QUiLoader

# Defina a classe CustomTreeWidget antes de carregar o .ui
# Isso é crucial se você promoveu QTreeWidget para CustomTreeWidget no Designer
class CustomTreeWidget(QTreeWidget):
    copyRequested = Signal(str)
    openRequested = Signal(str)

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setHeaderHidden(True)
        self.setColumnCount(2)
        # Conexões para os sinais padrão do QTreeWidget
        self.itemClicked.connect(self.on_item_clicked)
        self.itemDoubleClicked.connect(self.on_item_double_clicked)
        self.viewport().installEventFilter(self)
        self.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOn)

    def eventFilter(self, obj, event):
        if event.type() == QEvent.HoverMove and obj is self.viewport():
            item = self.itemAt(event.position().toPoint())
            if item and item.text(1): # Check if the second column (path) has text
                QToolTip.showText(event.globalPosition().toPoint(), item.text(1), self)
        return super().eventFilter(obj, event)

    def on_item_clicked(self, item, column):
        # Emit custom signals based on the clicked column
        if column == 1 and item.text(1): # If second column (path) is clicked
            self.copyRequested.emit(item.text(1))
        elif column == 0 and item.text(1): # If first column (name) is clicked
            self.openRequested.emit(item.text(1))

    def on_item_double_clicked(self, item, column):
        # Double-click always opens the folder if path exists
        if item.text(1):
            self.openRequested.emit(item.text(1))

# Globalized Strings
TEXTS = {
    "app_title": "Folder Session Manager", # Removido emoji para evitar problemas de codificação
    "title_label": "<html><head/><body><p><span style=\" font-size:14pt; font-weight:700;\">FOLDER SESSION MANAGER</span></p></body></html>", # Removido emoji
    "sessions_label": "<html><head/><body><p><span style=\" font-size:12pt; font-weight:700;\">📜 SESSIONS</span></p></body></html>",
    "preview_label": "<html><head/><body><p><span style=\" font-size:12pt; font-weight:700;\">🔎 SESSION PREVIEW</span></p></body></html>",
    "search_placeholder": "🔍 Search paths...",
    "save_button": "💾 Save Session",
    "restore_button": "🔄 Restore Session",
    "clear_button": "🗑️ Clear Preview",
    "choose_session_dir": "📂 Choose Session Folder",
    "close_windows_cb": "🛑 Close All Windows",
    "timeout_label": "⏱️ Timeout (s):",
    "expand_all": "⬆️ Expand All",
    "collapse_all": "⬇️ Collapse All",
    "open_folder": "📂 Open Folder",
    "delete_session": "🗑️ Delete Session",
    "view_trash": "♻️ View Trash",
    "no_paths": "🚫 No paths found.",
    "error_msg": "❌ Error: ",
    "copy_icon": "📋",
    "confirm_close": "⚠️ Confirm Close",
    "confirm_msg": "Close the following windows?\n\n{window_list}",
    "yes": "✅ Yes",
    "no": "❌ No",
    "session_saved": "✅ Session saved: {file}",
    "session_deleted": "🗑️ Session deleted: {file}",
    "session_restored": "♻️ Session restored: {file}",
    "error_restoring": "❌ Error restoring session: {error}",
    "restored_window": "✅ Restored window {hwnd}",
    "theme_label": "🎨 Theme:",
    "dark_theme": "🌙 Dark",
    "light_theme": "☀️ Light",
    "item_count": "📊 {count} items in session",
    "invalid_path": "🚫 Path does not exist: {path}",
    "platform_error": "❌ This application is only supported on Windows.",
    "dir_changed": "📂 Session folder changed to: {dir}",
    "confirm_delete": "⚠️ Confirm Delete",
    "confirm_delete_msg": "Delete session {file}?\nIt will be moved to trash for 24 hours.",
}
# TEXTS = {
#     "app_title": "Folder Session Manager",
#     "title_label": "<html><head/><body><p><span style=\" font-size:14pt; font-weight:700;\">FOLDER SESSION MANAGER</span></p></body></html>",
#     "sessions_label": "<html><head/><body><p><span style=\" font-size:12pt; font-weight:700;\">SESSIONS</span></p></body></html>",
#     "preview_label": "<html><head/><body><p><span style=\" font-size:12pt; font-weight:700;\">SESSION PREVIEW</span></p></body></html>",
#     "search_placeholder": "Search paths...",
#     "save_button": "Save Session",
#     "restore_button": "Restore Session",
#     "clear_button": "Clear Preview",
#     "choose_session_dir": "Choose Session Folder",
#     "close_windows_cb": "Close All Windows",
#     "timeout_label": "Timeout (s):",
#     "expand_all": "Expand All",
#     "collapse_all": "Collapse All",
#     "open_folder": "Open Folder",
#     "delete_session": "Delete Session",
#     "view_trash": "View Trash",
#     "no_paths": "No paths found.",
#     "error_msg": "Error: ",
#     "copy_icon": "Copied: ",
#     "confirm_close": "Confirm Close",
#     "confirm_msg": "Close the following windows?\n\n{window_list}",
#     "yes": "Yes",
#     "no": "No",
#     "session_saved": "Session saved: {file}",
#     "session_deleted": "Session deleted: {file}",
#     "session_restored": "Session restored: {file}",
#     "error_restoring": "Error restoring session: {error}",
#     "restored_window": "Restored window {hwnd}",
#     "theme_label": "Theme:",
#     "dark_theme": "Dark",
#     "light_theme": "Light",
#     "item_count": "{count} items in session",
#     "invalid_path": "Path does not exist: {path}",
#     "platform_error": "This application is only supported on Windows.",
#     "dir_changed": "Session folder changed to: {dir}",
#     "confirm_delete": "Confirm Delete",
#     "confirm_delete_msg": "Delete session {file}?\nIt will be moved to trash for 24 hours.",
# }
# Configuration
CONFIG = {
    "styles": {
        "dark": {
            "bg": "#2b2b2b",  # Slightly lighter dark background for depth
            "text": "#f0f0f0",  # Brighter text for contrast
            "button": "#3c3f41",  # Softer button background
            "button_hover": "#4a4d4f",  # Lighter hover effect
            "accent": "#40c4ff",  # Vibrant blue accent (closer to Fusion)
            "input_bg": "#353535",  # Slightly lighter input background
            "row_alt_color": "#303030"  # Subtle alternating row color
        },
        "light": {
            "bg": "#fafafa",  # Clean light background
            "text": "#212121",  # Darker text for readability
            "button": "#e0e0e0",  # Light button background
            "button_hover": "#bdbdbd",  # Slightly darker hover
            "accent": "#0288d1",  # Fusion-like blue accent
            "input_bg": "#ffffff",  # Pure white input fields
            "row_alt_color": "#f5f5f5"  # Subtle alternating row color
        }
    },
    "sizes": {
        "fixed": (568, 721),
        "min": (200, 300),
        "font": (12, 16),
        "padding_prop": 0.03,  # Increased for more spacing
        "button_min_width": 120,  # Slightly wider buttons
        "tree_col0_width": 300,
        "tree_col1_min_width": 900,
        "preview_min_height": 400
    },
    "max_visible_items": 10,
    "session_dir": os.path.join(os.path.expanduser("~"), "FolderSessions"),
    "trash_dir": os.path.join(os.path.expanduser("~"), "FolderSessions", ".trash"),
    "font_scale_limits": (0.4, 1.5),
    "zoom_step": 0.05
}

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        if platform.system() != "Windows":
            QMessageBox.critical(self, TEXTS["error_msg"], TEXTS["platform_error"])
            sys.exit(1)
        self.settings = QSettings("Eternal 3D Solutions", "FolderSessionManager")
        self.theme = self.settings.value("theme", "dark", str)
        self.font_scale = self.settings.value("font_scale", 1.0, float)
        self.session_dir = self.settings.value("session_dir", CONFIG["session_dir"], str)
        self.trash_dir = os.path.join(self.session_dir, ".trash")
        try:
            os.makedirs(self.session_dir, exist_ok=True)
            os.makedirs(self.trash_dir, exist_ok=True)
        except OSError as e:
            QMessageBox.critical(self, TEXTS["error_msg"], f"{TEXTS['error_msg']}{e}")
            self.session_dir = os.getcwd()
            self.trash_dir = os.path.join(self.session_dir, ".trash")
        self.setWindowTitle(TEXTS["app_title"])
        self.setMinimumSize(*CONFIG["sizes"]["min"])
        self.resize(*CONFIG["sizes"]["fixed"])
        self.setToolTipDuration(-1)
        self.setup_ui()
        self.update_session_list()

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
            if hasattr(loaded_ui, 'centralwidget') and isinstance(loaded_ui.centralwidget, QWidget):
                self.setCentralWidget(loaded_ui.centralwidget)
                print("Central widget set.")
            else:
                QMessageBox.critical(self, "UI Error", "Central widget not found. Check the .ui file.")
                sys.exit(1)

            if hasattr(loaded_ui, 'statusbar') and isinstance(loaded_ui.statusbar, QStatusBar):
                self.setStatusBar(loaded_ui.statusbar)
                print("Status bar set.")
            
            # Assign all widgets to self
            widget_names = []
            for child in self.findChildren(QObject):
                if child.objectName():
                    setattr(self, child.objectName(), child)
                    widget_names.append(child.objectName())
            print("Widgets loaded:", widget_names)

        except Exception as e:
            QMessageBox.critical(self, "UI Load Error", f"Failed to load UI file: {path_to_ui}\nError: {e}")
            sys.exit(1)

        # Update widget texts
        widget_text_mapping = {
            'tlabel': 'title_label',
            'sessions_label': 'sessions_label',
            'choose_dir_button': 'choose_session_dir',
            'view_trash_button': 'view_trash',
            'theme_label': 'theme_label',
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

        # Configure preview_tree
        if hasattr(self, 'preview_tree'):
            print("Type of preview_tree:", type(self.preview_tree))
            self.preview_tree.setHeaderHidden(True)
            self.preview_tree.setColumnCount(2)
            header_item = QTreeWidgetItem()
            header_item.setText(0, "Name")
            header_item.setText(1, "Path")
            self.preview_tree.setHeaderItem(header_item)

        # Connect signals
        if hasattr(self, 'sessions_list'):
            self.sessions_list.itemClicked.connect(self.update_preview)
            self.sessions_list.setContextMenuPolicy(Qt.CustomContextMenu)
            self.sessions_list.customContextMenuRequested.connect(self.show_session_context_menu)
        if hasattr(self, 'preview_tree'):
            self.preview_tree.copyRequested.connect(self.copy_to_clipboard)
            self.preview_tree.openRequested.connect(self.open_folder)
        if hasattr(self, 'close_windows_cb'):
            self.close_windows_cb.stateChanged.connect(self.confirm_close_windows)
        if hasattr(self, 'timeout_spin'):
            self.timeout_spin.setMinimum(0)
            self.timeout_spin.setMaximum(100)
            self.timeout_spin.setValue(self.settings.value("timeout", 1, int))
            self.timeout_spin.valueChanged.connect(lambda: self.settings.setValue("timeout", self.timeout_spin.value()))
        if hasattr(self, 'theme_combo'):
            self.theme_combo.clear()
            self.theme_combo.addItems([TEXTS["dark_theme"], TEXTS["light_theme"]])
            self.theme_combo.setCurrentText(TEXTS["dark_theme"] if self.theme == "dark" else TEXTS["light_theme"])
            self.theme_combo.currentTextChanged.connect(self.change_theme)
        if hasattr(self, 'choose_dir_button'):
            self.choose_dir_button.clicked.connect(self.choose_session_dir)
        if hasattr(self, 'view_trash_button'):
            self.view_trash_button.clicked.connect(self.view_trash)
        if hasattr(self, 'clear_button'):
            self.clear_button.clicked.connect(self.clear_preview)
        if hasattr(self, 'search_input'):
            self.search_input.textChanged.connect(self.filter_preview)
        if hasattr(self, 'save_button'):
            self.save_button.clicked.connect(lambda: print("Save button clicked") or self.save_session())
        if hasattr(self, 'restore_button'):
            self.restore_button.clicked.connect(lambda: print("Restore button clicked") or self.restore_session())

        # Apply theme after setting up UI
        self.apply_theme()

    def apply_theme(self):
        padding = int(CONFIG["sizes"]["fixed"][0] * CONFIG["sizes"]["padding_prop"])
        button_padding = padding // 2
        font_m, font_l = [max(8, int(x * self.font_scale)) for x in CONFIG["sizes"]["font"]]
        styles = CONFIG["styles"][self.theme]
        self.setStyleSheet(f"""
            QWidget {{ 
                background: {styles['bg']}; 
                color: {styles['text']}; 
                font-family: 'Segoe UI'; 
                font-size: {font_m}pt; 
            }}
            QLabel#section_label {{ 
                font-size: {font_l}pt; 
                font-weight: bold; 
                text-transform: uppercase; 
                margin: 10px 0px; 
            }}
            QLabel#tlabel {{ 
                font-size: {font_l + 2}pt; 
                font-weight: bold; 
                text-transform: uppercase; 
                color: {styles['accent']}; 
                margin: 15px 0px 10px 0px; 
            }}
            QCheckBox {{ 
                font-size: {font_m}pt; 
                padding: {button_padding}px; 
                margin: 8px; 
            }}
            QSpinBox, QLineEdit, QComboBox {{ 
                background: {styles['input_bg']}; 
                color: {styles['text']}; 
                border: 1px solid {styles['accent']}; 
                border-radius: 8px; 
                font-size: {font_m}pt; 
                padding: {button_padding}px; 
                margin: 8px; 
            }}
            QTreeWidget {{ 
                background: {styles['input_bg']}; 
                color: {styles['text']}; 
                border: 1px solid {styles['accent']}; 
                border-radius: 8px; 
                font-size: {font_m}pt; 
                padding: {button_padding}px; 
                margin: 8px; 
            }}
            QTreeWidget::item {{ 
                padding: {button_padding//2}px; 
            }}
            QTreeWidget::item:alternate {{ 
                background: {styles['row_alt_color']}; 
            }}
            QTreeWidget::item:selected {{ 
                background: {styles['accent']}; 
                color: {styles['bg']}; 
            }}
            QPushButton {{ 
                background: {styles['button']}; 
                color: {styles['text']}; 
                border: 1px solid {styles['accent']}; 
                border-radius: 8px; 
                padding: {button_padding}px; 
                font-size: {font_m}pt; 
                min-width: {CONFIG['sizes']['button_min_width']}px; 
                margin: 8px; 
            }}
            QPushButton:hover {{ 
                background: {styles['button_hover']}; 
            }}
            QStatusBar {{ 
                background: {styles['input_bg']}; 
                color: {styles['text']}; 
                font-size: {font_m}pt; 
                padding: {button_padding}px; 
                margin: 5px; 
            }}
            QListWidget {{ 
                background: {styles['input_bg']}; 
                color: {styles['text']}; 
                border: 1px solid {styles['accent']}; 
                border-radius: 8px; 
                font-size: {font_m}pt; 
                padding: {button_padding}px; 
                margin: 8px; 
            }}
            QFrame#line, QFrame#line_2 {{ 
                background: {styles['accent']}; 
                margin: 8px 0px; 
                min-height: 2px; 
            }}
        """)
        print(f"Applying theme: {self.theme}")
        self.update()
        self.adjust_column_width()

    def wheelEvent(self, event):
        if event.modifiers() & Qt.ControlModifier:
            delta = event.angleDelta().y() / 120
            new_scale = self.font_scale + delta * CONFIG["zoom_step"]
            self.font_scale = max(CONFIG["font_scale_limits"][0], min(CONFIG["font_scale_limits"][1], new_scale))
            self.settings.setValue("font_scale", self.font_scale)
            self.apply_theme()
            if hasattr(self, 'statusbar'):
                self.statusbar.showMessage(f"Zoom adjusted to {int(self.font_scale * 100)}%", 2000)
            event.accept()
        else:
            super().wheelEvent(event)

    def adjust_column_width(self):
        if hasattr(self, 'preview_tree'):
            self.preview_tree.setColumnWidth(0, CONFIG["sizes"]["tree_col0_width"])
            self.preview_tree.setColumnWidth(1, CONFIG["sizes"]["tree_col1_min_width"])

    def change_theme(self, theme_text):
        self.theme = "dark" if theme_text == TEXTS["dark_theme"] else "light"
        self.settings.setValue("theme", self.theme)
        self.apply_theme()

    def choose_session_dir(self):
        new_dir = QFileDialog.getExistingDirectory(self, TEXTS["choose_session_dir"], self.session_dir)
        if new_dir:
            self.session_dir = new_dir
            self.trash_dir = os.path.join(new_dir, ".trash")
            try:
                os.makedirs(self.trash_dir, exist_ok=True)
                self.settings.setValue("session_dir", self.session_dir)
                self.settings.setValue("trash_dir", self.trash_dir)
                self.update_session_list()
                if hasattr(self, 'statusbar'):
                    self.statusbar.showMessage(TEXTS["dir_changed"].format(dir=new_dir), 5000)
            except OSError as e:
                QMessageBox.critical(self, TEXTS["error_msg"], f"{TEXTS['error_msg']}{e}")

    def update_session_list(self):
        if not hasattr(self, 'sessions_list'):
            return
        self.sessions_list.clear()
        try:
            all_sessions = []
            for directory in [self.session_dir, self.trash_dir]:
                for f in sorted(os.listdir(directory), key=lambda x: os.path.getmtime(os.path.join(directory, x)), reverse=True):
                    if f.endswith('.json') and f.startswith('explorer_session'):
                        try:
                            with open(os.path.join(directory, f), 'r', encoding='utf-8') as file:
                                json.load(file)
                            mtime = datetime.fromtimestamp(os.path.getmtime(os.path.join(directory, f)))
                            full_path = os.path.join(directory, f)
                            all_sessions.append((mtime, f, full_path, " (Trash)" if directory == self.trash_dir else ""))
                        except json.JSONDecodeError:
                            if hasattr(self, 'statusbar'):
                                self.statusbar.showMessage(f"{TEXTS['error_msg']}Skipping invalid JSON file: {f}", 5000)
            for mtime, f, full_path, trash_suffix in sorted(all_sessions, key=lambda x: x[0], reverse=True):
                item = QListWidgetItem(f"{mtime.strftime('%Y-%m-%d %H:%M')} - {f}{trash_suffix}")
                item.setData(Qt.UserRole, full_path)
                if trash_suffix:
                    item.setForeground(Qt.gray)
                self.sessions_list.addItem(item)
        except OSError as e:
            QMessageBox.critical(self, TEXTS["error_msg"], f"{TEXTS['error_msg']}{e}")

    def show_session_context_menu(self, pos):
        if not hasattr(self, 'sessions_list'): return
        item = self.sessions_list.itemAt(pos)
        if not item:
            return
        menu = QMenu()
        delete_action = menu.addAction(TEXTS["delete_session"])
        restore_action = menu.addAction(TEXTS["view_trash"]) if "(Trash)" in item.text() else None
        action = menu.exec(self.sessions_list.viewport().mapToGlobal(pos))
        if action == delete_action:
            self.delete_session(item)
        elif action == restore_action:
            self.restore_from_trash(item)

    def delete_session(self, item):
        file_path = item.data(Qt.UserRole)
        reply = QMessageBox.question(self, TEXTS["confirm_delete"],
                                     TEXTS["confirm_delete_msg"].format(file=os.path.basename(file_path)),
                                     QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
        if reply == QMessageBox.Yes:
            try:
                trash_path = os.path.join(self.trash_dir, os.path.basename(file_path))
                if os.path.exists(trash_path):
                    os.remove(trash_path)
                shutil.move(file_path, trash_path)
                self.update_session_list()
                if hasattr(self, 'statusbar'):
                    self.statusbar.showMessage(TEXTS["session_deleted"].format(file=os.path.basename(file_path)), 5000)
            except (OSError, shutil.Error) as e:
                QMessageBox.critical(self, TEXTS["error_msg"], f"{TEXTS['error_msg']}{e}")

    def view_trash(self):
        self.update_session_list()
        if hasattr(self, 'statusbar'):
            self.statusbar.showMessage("Trash updated", 3000)

    def restore_from_trash(self, item):
        file_path = item.data(Qt.UserRole)
        try:
            shutil.move(file_path, os.path.join(self.session_dir, os.path.basename(file_path)))
            self.update_session_list()
            if hasattr(self, 'statusbar'):
                self.statusbar.showMessage(TEXTS["session_restored"].format(file=os.path.basename(file_path)), 5000)
        except (OSError, shutil.Error) as e:
            QMessageBox.critical(self, TEXTS["error_msg"], f"{TEXTS['error_msg']}{e}")

    def get_explorer_windows(self):
        windows_data = []
        try:
            shell = win32com.client.Dispatch("Shell.Application")
            path_by_hwnd = {}
            for window in shell.Windows():
                try:
                    if hasattr(window, 'Document') and hasattr(window.Document, 'Folder'):
                        hwnd = window.HWND
                        path = window.Document.Folder.Self.Path
                        path_by_hwnd.setdefault(hwnd, []).append(path)
                except Exception as e:
                    if hasattr(self, 'statusbar'):
                        self.statusbar.showMessage(f"{TEXTS['error_msg']}Error getting window: {e}", 5000)
                    continue
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
                except Exception as e:
                    if hasattr(self, 'statusbar'):
                        self.statusbar.showMessage(f"{TEXTS['error_msg']}Error getting window position: {e}", 5000)
                    continue
        except Exception as e:
            if hasattr(self, 'statusbar'):
                self.statusbar.showMessage(f"{TEXTS['error_msg']}Failed to access Shell: {e}", 5000)
        return windows_data

    def copy_to_clipboard(self, text):
        QApplication.clipboard().setText(text)
        if hasattr(self, 'statusbar'):
            self.statusbar.showMessage(f"{TEXTS['copy_icon']} {text[:30]}...", 3000)

    def confirm_close_windows(self, state):
        if state == Qt.Checked:
            windows = self.get_explorer_windows()
            window_list = "\n".join([f"Window {w['hwnd']}: {', '.join(w['paths'])}" for w in windows]) or TEXTS["no_paths"]
            reply = QMessageBox.question(self, TEXTS["confirm_close"],
                                         TEXTS["confirm_msg"].format(window_list=window_list),
                                         QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
            if reply == QMessageBox.Yes:
                self.close_existing_windows()
            else:
                if hasattr(self, 'close_windows_cb'):
                    self.close_windows_cb.setChecked(False)

    def close_existing_windows(self):
        try:
            shell = win32com.client.Dispatch("Shell.Application")
            for window in shell.Windows():
                try:
                    if hasattr(window, 'HWND'):
                        win32gui.PostMessage(window.HWND, 0x0010, 0, 0)
                except Exception as e:
                    if hasattr(self, 'statusbar'):
                        self.statusbar.showMessage(f"{TEXTS['error_msg']}Error closing window: {e}", 5000)
                    continue
            if hasattr(self, 'close_windows_cb'):
                self.close_windows_cb.setChecked(False)
            if hasattr(self, 'statusbar'):
                self.statusbar.showMessage("All Explorer windows closed", 3000)
        except Exception as e:
            if hasattr(self, 'statusbar'):
                self.statusbar.showMessage(f"{TEXTS['error_msg']}Failed to close windows: {e}", 5000)

    def expand_all(self):
        if hasattr(self, 'preview_tree'):
            self.preview_tree.expandAll()
        if hasattr(self, 'statusbar'):
            self.statusbar.showMessage(TEXTS["expand_all"], 3000)

    def collapse_all(self):
        if hasattr(self, 'preview_tree'):
            self.preview_tree.collapseAll()
        if hasattr(self, 'statusbar'):
            self.statusbar.showMessage(TEXTS["collapse_all"], 3000)

    def clear_preview(self):
        if hasattr(self, 'preview_tree'):
            self.preview_tree.clear()
        if hasattr(self, 'search_input'):
            self.search_input.clear()
        if hasattr(self, 'item_count_label'):
            self.item_count_label.setText(TEXTS["item_count"].format(count=0))
        if hasattr(self, 'statusbar'):
            self.statusbar.showMessage(TEXTS["clear_button"], 3000)

    def filter_preview(self, text):
        if not hasattr(self, 'preview_tree'): return
        if not text:
            for i in range(self.preview_tree.topLevelItemCount()):
                item = self.preview_tree.topLevelItem(i)
                item.setHidden(False)
                for j in range(item.childCount()):
                    item.child(j).setHidden(False)
            return
        text = unicodedata.normalize('NFKD', text.lower()).encode('ASCII', 'ignore').decode('ASCII')
        for i in range(self.preview_tree.topLevelItemCount()):
            item = self.preview_tree.topLevelItem(i)
            item_hidden = True
            for j in range(item.childCount()):
                child = item.child(j)
                child_text = unicodedata.normalize('NFKD', child.text(0).lower()).encode('ASCII', 'ignore').decode('ASCII')
                full_path = unicodedata.normalize('NFKD', child.text(1).lower()).encode('ASCII', 'ignore').decode('ASCII')
                if text in child_text or text in full_path:
                    child.setHidden(False)
                    item_hidden = False
                else:
                    child.setHidden(True)
            item.setHidden(item_hidden)

    def update_preview(self, item):
        if not hasattr(self, 'preview_tree'): return
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
                    window_item.setExpanded(item_count < CONFIG["max_visible_items"])
                    item_count += 1
                    for path in paths:
                        path_item = QTreeWidgetItem(window_item)
                        path_item.setText(0, f"{os.path.basename(path)}")
                        path_item.setText(1, path)
                        item_count += 1
                        if not os.path.exists(path):
                            path_item.setForeground(1, QColor("red"))
                            path_item.setToolTip(1, TEXTS["invalid_path"].format(path=path))
                        current_theme_styles = CONFIG["styles"][self.theme]
                        if window_item.childCount() % 2 == 0:
                            path_item.setBackground(0, QColor(current_theme_styles["row_alt_color"]))
                            path_item.setBackground(1, QColor(current_theme_styles["row_alt_color"]))
                if hasattr(self, 'item_count_label'):
                    self.item_count_label.setText(TEXTS["item_count"].format(count=item_count))
        except (OSError, json.JSONDecodeError) as e:
            QMessageBox.critical(self, TEXTS["error_msg"], f"{TEXTS['error_restoring'].format(error=str(e))}")
            if hasattr(self, 'item_count_label'):
                self.item_count_label.setText(TEXTS["item_count"].format(count=0))

    def open_folder(self, path):
        if os.path.exists(path):
            try:
                os.startfile(path)
                if hasattr(self, 'statusbar'):
                    self.statusbar.showMessage(f"{TEXTS['open_folder']} {os.path.basename(path)}", 3000)
            except Exception as e:
                if hasattr(self, 'statusbar'):
                    self.statusbar.showMessage(f"{TEXTS['error_msg']}{e}", 5000)
        else:
            if hasattr(self, 'statusbar'):
                self.statusbar.showMessage(TEXTS["invalid_path"].format(path=path), 5000)

    def save_session(self):
        timestamp = datetime.now().strftime('%Y-%m-%d_%H-%M-%S')
        file_name = os.path.join(self.session_dir, f"explorer_session_{timestamp}.json")
        backup_name = os.path.join(self.trash_dir, f"backup_explorer_session_{timestamp}.json")
        try:
            with open(file_name, 'w', encoding='utf-8') as f:
                json.dump(self.get_explorer_windows(), f, indent=4)
            with open(backup_name, 'w', encoding='utf-8') as f_backup:
                json.dump(self.get_explorer_windows(), f_backup, indent=4)
            self.update_session_list()
            if hasattr(self, 'statusbar'):
                self.statusbar.showMessage(TEXTS["session_saved"].format(file=os.path.basename(file_name)), 5000)
        except OSError as e:
            QMessageBox.critical(self, TEXTS["error_msg"], f"{TEXTS['error_msg']}Failed to save session: {e}")

    def restore_session(self):
        if not hasattr(self, 'sessions_list'): return
        if item := self.sessions_list.currentItem():
            file_path = item.data(Qt.UserRole)
            try:
                with open(file_path, 'r', encoding='utf-8') as f:
                    session_data = json.load(f)
                    shell = win32com.client.Dispatch("Shell.Application")
                    wshell = win32com.client.Dispatch("WScript.Shell")
                    screen_width = win32api.GetSystemMetrics(0)
                    screen_height = win32api.GetSystemMetrics(1)

                    total_paths = sum(len(entry['paths']) for entry in session_data)
                    progress = QProgressDialog("Restoring session...", "Cancel", 0, total_paths, self)
                    progress.setWindowModality(Qt.WindowModal)
                    progress.setWindowFlags(progress.windowFlags() | Qt.WindowStaysOnTopHint)
                    progress.show()
                    QApplication.processEvents()
                    current_path_idx = 0

                    for entry in session_data:
                        hwnd, paths, x, y, width, height = entry['hwnd'], entry['paths'], entry['x'], entry['y'], entry['width'], entry['height']
                        if not paths:
                            continue

                        initial_path = paths[0] if os.path.exists(paths[0]) else os.path.expanduser("~")
                        explorer = shell.Explore(initial_path)
                        time.sleep(0.5)
                        explorer_hwnd = None
                        for _ in range(3):
                            for window in shell.Windows():
                                if window.Document.Folder.Self.Path == initial_path:
                                    explorer_hwnd = window.HWND
                                    break
                            if explorer_hwnd:
                                break
                            time.sleep(0.5)
                        if explorer_hwnd is None:
                            if hasattr(self, 'statusbar'):
                                self.statusbar.showMessage(f"{TEXTS['error_msg']}Failed to find window for hwnd {hwnd}", 5000)
                            continue

                        win32gui.MoveWindow(explorer_hwnd, x, y, width, height, True)
                        time.sleep(0.5)

                        for path in paths[1:]:
                            if os.path.exists(path):
                                max_retries = 2
                                for attempt in range(max_retries + 1):
                                    try:
                                        wshell.AppActivate(explorer_hwnd)
                                        wshell.SendKeys("^t")
                                        time.sleep(0.7)
                                        wshell.SendKeys("%d")
                                        time.sleep(0.5)
                                        wshell.SendKeys(path)
                                        time.sleep(0.5)
                                        timeout_val = self.timeout_spin.value() if hasattr(self, 'timeout_spin') else 1
                                        wshell.SendKeys("{ENTER}")
                                        time.sleep(max(1.0, timeout_val / 100.0))
                                        if hasattr(self, 'statusbar'):
                                            self.statusbar.showMessage(
                                                f"{TEXTS['restored_window'].format(hwnd=hwnd)} in tab", 3000)
                                        break
                                    except Exception as e:
                                        if attempt < max_retries:
                                            if hasattr(self, 'statusbar'):
                                                self.statusbar.showMessage(f"{TEXTS['error_msg']}Retry {attempt + 1}/{max_retries} for path {path}: {e}", 3000)
                                            time.sleep(1.0)
                                        else:
                                            reply = QMessageBox.question(self, "Manual Action Needed",
                                                                         f"Failed to open {path} after {max_retries} retries.\n"
                                                                         "1. Open manually?\n2. Skip?",
                                                                         QMessageBox.Yes | QMessageBox.No | QMessageBox.Cancel,
                                                                         QMessageBox.Yes)
                                            if reply == QMessageBox.Yes:
                                                shell.Explore(path)
                                                if hasattr(self, 'statusbar'):
                                                    self.statusbar.showMessage(f"Opened {path} manually", 3000)
                                            elif reply == QMessageBox.No:
                                                if hasattr(self, 'statusbar'):
                                                    self.statusbar.showMessage(f"Skipped {path}", 3000)
                                            else:
                                                if hasattr(self, 'statusbar'):
                                                    self.statusbar.showMessage("Restore cancelled by user", 3000)
                                                progress.close()
                                                return
                            else:
                                if hasattr(self, 'statusbar'):
                                    self.statusbar.showMessage(TEXTS["invalid_path"].format(path=path), 5000)
                            current_path_idx += 1
                            progress.setValue(current_path_idx)
                            progress.setLabelText(f"Restoring: {os.path.basename(path)} ({current_path_idx}/{total_paths})")
                            QApplication.processEvents()
                            if progress.wasCanceled():
                                if hasattr(self, 'statusbar'):
                                    self.statusbar.showMessage("Restore cancelled by user", 3000)
                                progress.close()
                                return
                            if win32api.GetAsyncKeyState(0x1B):  # Escape key
                                if hasattr(self, 'statusbar'):
                                    self.statusbar.showMessage("Restore stopped by user", 3000)
                                progress.close()
                                return
                        time.sleep(1.0)
                    progress.setValue(total_paths)
                    progress.close()
            except (OSError, json.JSONDecodeError) as e:
                QMessageBox.critical(self, TEXTS["error_msg"], f"{TEXTS['error_restoring'].format(error=str(e))}")

if __name__ == "__main__":
    app = QApplication(sys.argv)
    app.setStyle("Fusion")
    QLocale.setDefault(QLocale(QLocale.English, QLocale.UnitedStates))
    window = MainWindow()
    window.show()
    sys.exit(app.exec())
