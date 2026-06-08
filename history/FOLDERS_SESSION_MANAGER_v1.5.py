import sys
import os
import json
import time
from datetime import datetime
import win32gui
import win32com.client
import platform
import unicodedata
import shutil
from PySide6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QLabel,
                              QCheckBox, QSpinBox, QListWidget, QListWidgetItem, QFileDialog, QTreeWidget,
                              QTreeWidgetItem, QStatusBar, QComboBox, QLineEdit, QGridLayout, QToolTip, QHeaderView,
                              QMenu, QMessageBox)
from PySide6.QtCore import Qt, QEvent, Signal, QSettings, QPoint, QDateTime
from PySide6.QtGui import QColor, QClipboard, QFontMetrics, QIcon, QKeySequence, QWheelEvent
from PySide6.QtWidgets import QStyle

# Configurações
CONFIG = {
    "styles": {
        "dark": {
            "bg": "#1e1e1e", "text": "#e0e0e0", "button": "#333333", "button_hover": "#444444",
            "accent": "#2196f3", "input_bg": "#2a2a2a", "row_alt_color": "#252525"
        },
        "light": {
            "bg": "#f0f0f0", "text": "#333333", "button": "#e0e0e0", "button_hover": "#d0d0d0",
            "accent": "#0288d1", "input_bg": "#ffffff", "row_alt_color": "#e8e8e8"
        }
    },
    "sizes": {
        "fixed": (1200, 600), "min": (1200, 500), "font": (12, 14), "padding_prop": 0.02,
        "button_min_width": 100, "tree_col0_width": 300, "tree_col1_min_width": 900
    },
    "texts": {
        "app_title": "📁 Folder Session Manager",
        "title_label": "📁 Folder Session Manager",
        "sessions_label": "📜 Sessions",
        "preview_label": "🔎 Session Preview",
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
    },
    "max_visible_items": 10,
    "session_dir": os.path.join(os.path.expanduser("~"), "FolderSessions"),
    "trash_dir": os.path.join(os.path.expanduser("~"), "FolderSessions", ".trash"),
    "font_scale_limits": (0.4, 1.5),
    "zoom_step": 0.05
}

class CustomTreeWidget(QTreeWidget):
    """
    🌟 QTreeWidget para exibir caminhos de pastas com comportamento de clique personalizado.
    """
    copyRequested = Signal(str)
    openRequested = Signal(str)

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setHeaderHidden(True)
        self.setColumnCount(2)
        self.itemClicked.connect(self.on_item_clicked)
        self.itemDoubleClicked.connect(self.on_item_double_clicked)
        self.viewport().installEventFilter(self)
        self.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOn)

    def eventFilter(self, obj, event):
        if event.type() == QEvent.HoverMove and obj is self.viewport():
            item = self.itemAt(event.position().toPoint())
            if item and item.text(1):
                QToolTip.showText(event.globalPosition().toPoint(), item.text(1), self)
        return super().eventFilter(obj, event)

    def on_item_clicked(self, item, column):
        if column == 1 and item.text(1):
            self.copyRequested.emit(item.text(1))
        elif column == 0 and item.text(1):
            self.openRequested.emit(item.text(1))

    def on_item_double_clicked(self, item, column):
        if item.text(1):
            self.openRequested.emit(item.text(1))

class MainWindow(QMainWindow):
    """
    🌟 Janela principal para o Folder Session Manager.
    """
    def __init__(self):
        super().__init__()
        if platform.system() != "Windows":
            QMessageBox.critical(self, "Error", CONFIG["texts"]["platform_error"])
            sys.exit(1)
        self.settings = QSettings("xAI", "FolderSessionManager")
        self.theme = self.settings.value("theme", "dark", str)
        self.font_scale = self.settings.value("font_scale", 1.0, float)
        self.session_dir = self.settings.value("session_dir", CONFIG["session_dir"], str)
        self.trash_dir = self.settings.value("trash_dir", CONFIG["trash_dir"], str)
        try:
            os.makedirs(self.session_dir, exist_ok=True)
            os.makedirs(self.trash_dir, exist_ok=True)
        except OSError as e:
            QMessageBox.critical(self, "Error", f"❌ Failed to create directories: {e}")
            self.session_dir = os.getcwd()
            self.trash_dir = os.path.join(self.session_dir, ".trash")
        self.setWindowTitle(CONFIG["texts"]["app_title"])
        self.setMinimumSize(*CONFIG["sizes"]["min"])
        self.resize(*CONFIG["sizes"]["fixed"])
        self.setup_ui()
        self.apply_theme()
        self.update_session_list()

    def setup_ui(self):
        central = QWidget()
        self.setCentralWidget(central)
        main_layout = QVBoxLayout(central)
        padding = int(CONFIG["sizes"]["fixed"][0] * CONFIG["sizes"]["padding_prop"])
        main_layout.setContentsMargins(padding, padding, padding, padding)
        main_layout.setSpacing(padding)

        # Título
        title_label = QLabel(CONFIG["texts"]["title_label"])
        title_label.setAlignment(Qt.AlignLeft)
        title_label.setObjectName("section_label")
        main_layout.addWidget(title_label)

        # Layout Principal (Grade)
        content_layout = QGridLayout()
        content_layout.setSpacing(padding // 2)
        main_layout.addLayout(content_layout)

        # Painel Esquerdo (Sessões)
        left_panel = QWidget()
        left_layout = QVBoxLayout(left_panel)
        left_layout.addWidget(QLabel(CONFIG["texts"]["sessions_label"], objectName="section_label"))
        self.sessions_list = QListWidget()
        self.sessions_list.itemClicked.connect(self.update_preview)
        self.sessions_list.setContextMenuPolicy(Qt.CustomContextMenu)
        self.sessions_list.customContextMenuRequested.connect(self.show_session_context_menu)
        left_layout.addWidget(self.sessions_list)

        # Configurações
        settings_layout = QVBoxLayout()
        self.close_windows_cb = QCheckBox(CONFIG["texts"]["close_windows_cb"])
        self.close_windows_cb.setToolTip("Close all open Explorer windows 🛑")
        self.close_windows_cb.stateChanged.connect(self.confirm_close_windows)
        settings_layout.addWidget(self.close_windows_cb)
        timeout_layout = QHBoxLayout()
        timeout_layout.addWidget(QLabel(CONFIG["texts"]["timeout_label"]))
        self.timeout_spin = QSpinBox()
        self.timeout_spin.setRange(1, 60)
        self.timeout_spin.setValue(self.settings.value("timeout", 10, int))
        self.timeout_spin.setToolTip("Set delay for restoring windows ⏱️")
        timeout_layout.addWidget(self.timeout_spin)
        settings_layout.addLayout(timeout_layout)
        theme_layout = QHBoxLayout()
        theme_layout.addWidget(QLabel(CONFIG["texts"]["theme_label"]))
        self.theme_combo = QComboBox()
        self.theme_combo.addItems([CONFIG["texts"]["dark_theme"], CONFIG["texts"]["light_theme"]])
        self.theme_combo.setCurrentText(CONFIG["texts"][f"{self.theme}_theme"])
        self.theme_combo.setToolTip("Choose application theme 🎨")
        self.theme_combo.currentTextChanged.connect(self.change_theme)
        theme_layout.addWidget(self.theme_combo)
        settings_layout.addLayout(theme_layout)
        self.choose_dir_button = QPushButton(CONFIG["texts"]["choose_session_dir"])
        self.choose_dir_button.setToolTip("Choose where to save session files 📂")
        self.choose_dir_button.clicked.connect(self.choose_session_dir)
        settings_layout.addWidget(self.choose_dir_button)
        self.view_trash_button = QPushButton(CONFIG["texts"]["view_trash"])
        self.view_trash_button.setToolTip("View recently deleted sessions ♻️")
        self.view_trash_button.clicked.connect(self.view_trash)
        settings_layout.addWidget(self.view_trash_button)
        left_layout.addLayout(settings_layout)
        content_layout.addWidget(left_panel, 0, 0, 1, 1)

        # Painel Direito (Visualização)
        right_panel = QWidget()
        right_layout = QVBoxLayout(right_panel)
        preview_header = QHBoxLayout()
        preview_header.addWidget(QLabel(CONFIG["texts"]["preview_label"], objectName="section_label"))
        self.item_count_label = QLabel(CONFIG["texts"]["item_count"].format(count=0))
        preview_header.addWidget(self.item_count_label)
        preview_header.addStretch()
        self.search_input = QLineEdit()
        self.search_input.setPlaceholderText(CONFIG["texts"]["search_placeholder"])
        self.search_input.setToolTip("Search for folder paths 🔍")
        self.search_input.textChanged.connect(self.filter_preview)
        preview_header.addWidget(self.search_input)
        self.expand_button = QPushButton(self.style().standardIcon(QStyle.SP_ArrowUp), CONFIG["texts"]["expand_all"])
        self.expand_button.setToolTip("Expand all folders ⬆️")
        self.expand_button.setShortcut(QKeySequence("Ctrl+E"))
        preview_header.addWidget(self.expand_button)
        self.collapse_button = QPushButton(self.style().standardIcon(QStyle.SP_ArrowDown), CONFIG["texts"]["collapse_all"])
        self.collapse_button.setToolTip("Collapse all folders ⬇️")
        self.collapse_button.setShortcut(QKeySequence("Ctrl+C"))
        preview_header.addWidget(self.collapse_button)
        self.clear_button = QPushButton(self.style().standardIcon(QStyle.SP_TrashIcon), CONFIG["texts"]["clear_button"])
        self.clear_button.setToolTip("Clear preview 🗑️")
        self.clear_button.setShortcut(QKeySequence("Ctrl+L"))
        preview_header.addWidget(self.clear_button)
        right_layout.addLayout(preview_header)
        self.preview_tree = CustomTreeWidget()
        self.preview_tree.copyRequested.connect(self.copy_to_clipboard)
        self.preview_tree.openRequested.connect(self.open_folder)
        self.expand_button.clicked.connect(self.expand_all)
        self.collapse_button.clicked.connect(self.collapse_all)
        self.clear_button.clicked.connect(self.clear_preview)
        right_layout.addWidget(self.preview_tree)
        content_layout.addWidget(right_panel, 0, 1, 1, 2)

        # Ajustar proporções da grade
        content_layout.setColumnStretch(0, 1)
        content_layout.setColumnStretch(1, 2)
        content_layout.setColumnStretch(2, 2)

        # Botões de Ação
        action_layout = QHBoxLayout()
        self.save_button = QPushButton(CONFIG["texts"]["save_button"])
        self.save_button.setToolTip("Save current session 💾")
        self.save_button.setShortcut(QKeySequence("Ctrl+S"))
        self.restore_button = QPushButton(CONFIG["texts"]["restore_button"])
        self.restore_button.setToolTip("Restore selected session 🔄")
        self.restore_button.setShortcut(QKeySequence("Ctrl+R"))
        self.save_button.clicked.connect(self.save_session)
        self.restore_button.clicked.connect(self.restore_session)
        action_layout.addWidget(self.save_button)
        action_layout.addWidget(self.restore_button)
        action_layout.addStretch()
        main_layout.addLayout(action_layout)

        # Barra de Status
        self.status_bar = QStatusBar()
        main_layout.addWidget(self.status_bar)

        # Ajustar largura da coluna dinamicamente
        self.adjust_column_width()

    def wheelEvent(self, event: QWheelEvent):
        if event.modifiers() & Qt.ControlModifier:
            delta = event.angleDelta().y() / 120
            new_scale = self.font_scale + delta * CONFIG["zoom_step"]
            self.font_scale = max(CONFIG["font_scale_limits"][0], min(CONFIG["font_scale_limits"][1], new_scale))
            self.settings.setValue("font_scale", self.font_scale)
            self.apply_theme()
            self.status_bar.showMessage(f"🔍 Zoom ajustado para {int(self.font_scale * 100)}%", 2000)
            event.accept()
        else:
            super().wheelEvent(event)

    def adjust_column_width(self):
        window_width = self.width()
        col0_width = CONFIG["sizes"]["tree_col0_width"]
        self.preview_tree.setColumnWidth(0, col0_width)
        col1_min_width = CONFIG["sizes"]["tree_col1_min_width"]

        fm = QFontMetrics(self.font())
        max_width = 0
        for i in range(self.preview_tree.topLevelItemCount()):
            item = self.preview_tree.topLevelItem(i)
            for j in range(item.childCount()):
                child = item.child(j)
                text_width = fm.horizontalAdvance(child.text(1))
                max_width = max(max_width, text_width)
        self.preview_tree.setColumnWidth(1, max(col1_min_width, max_width + 20))

    def apply_theme(self):
        padding = int(CONFIG["sizes"]["fixed"][0] * CONFIG["sizes"]["padding_prop"])
        button_padding = padding // 2
        font_m, font_l = [max(8, int(x * self.font_scale)) for x in CONFIG["sizes"]["font"]]
        styles = CONFIG["styles"][self.theme]
        self.setStyleSheet(f"""
            QWidget {{ background: {styles['bg']}; color: {styles['text']}; font-family: 'Segoe UI'; font-size: {font_m}pt; }}
            QLabel#section_label {{ font-size: {font_l}pt; font-weight: bold; }}
            QCheckBox {{ font-size: {font_m}pt; }}
            QSpinBox, QLineEdit, QComboBox {{ background: {styles['input_bg']}; color: {styles['text']}; border: 1px solid {styles['accent']}; font-size: {font_m}pt; padding: {button_padding}px; }}
            QTreeWidget {{ background: {styles['input_bg']}; color: {styles['text']}; border: 1px solid {styles['accent']}; font-size: {font_m}pt; }}
            QTreeWidget::item {{ padding: {button_padding//2}px; }}
            QTreeWidget::item:alternate {{ background: {styles['row_alt_color']}; }}
            QTreeWidget::item:selected {{ background: {styles['accent']}; color: {styles['bg']}; }}
            QPushButton {{ background: {styles['button']}; color: {styles['text']}; border: 1px solid {styles['accent']};
                          padding: {button_padding}px; font-size: {font_m}pt; min-width: {CONFIG['sizes']['button_min_width']}px; }}
            QPushButton:hover {{ background: {styles['button_hover']}; }}
            QStatusBar {{ background: {styles['input_bg']}; color: {styles['text']}; font-size: {font_m}pt; }}
        """)
        self.update()
        self.adjust_column_width()

    def change_theme(self, theme_text):
        theme = "dark" if theme_text == CONFIG["texts"]["dark_theme"] else "light"
        self.theme = theme
        self.settings.setValue("theme", theme)
        self.apply_theme()

    def choose_session_dir(self):
        dir_path = QFileDialog.getExistingDirectory(self, "Choose Session Folder", self.session_dir)
        if dir_path:
            try:
                os.makedirs(dir_path, exist_ok=True)
                self.session_dir = dir_path
                self.trash_dir = os.path.join(dir_path, ".trash")
                os.makedirs(self.trash_dir, exist_ok=True)
                self.settings.setValue("session_dir", dir_path)
                self.settings.setValue("trash_dir", self.trash_dir)
                self.update_session_list()
                self.status_bar.showMessage(CONFIG["texts"]["dir_changed"].format(dir=dir_path), 5000)
            except OSError as e:
                QMessageBox.critical(self, "Error", f"❌ Failed to set session directory: {e}")

    def update_session_list(self):
        self.sessions_list.clear()
        try:
            for f in sorted(os.listdir(self.session_dir), key=lambda x: os.path.getmtime(os.path.join(self.session_dir, x)), reverse=True):
                if f.endswith('.json') and f.startswith('explorer_session'):
                    try:
                        with open(os.path.join(self.session_dir, f), 'r', encoding='utf-8') as file:
                            json.load(file)
                        mtime = datetime.fromtimestamp(os.path.getmtime(os.path.join(self.session_dir, f)))
                        item = QListWidgetItem(f"{f} - 📅 {mtime.strftime('%Y-%m-%d %H:%M')}")
                        item.setData(Qt.UserRole, os.path.join(self.session_dir, f))
                        self.sessions_list.addItem(item)
                    except json.JSONDecodeError:
                        self.status_bar.showMessage(f"🚫 Skipping invalid JSON file: {f}", 5000)
            for f in sorted(os.listdir(self.trash_dir), key=lambda x: os.path.getmtime(os.path.join(self.trash_dir, x)), reverse=True):
                if f.endswith('.json') and f.startswith('explorer_session'):
                    mtime = datetime.fromtimestamp(os.path.getmtime(os.path.join(self.trash_dir, f)))
                    item = QListWidgetItem(f"{f} - 📅 {mtime.strftime('%Y-%m-%d %H:%M')} (Trash)")
                    item.setData(Qt.UserRole, os.path.join(self.trash_dir, f))
                    item.setForeground(Qt.gray)
                    self.sessions_list.addItem(item)
        except OSError as e:
            QMessageBox.critical(self, "Error", f"❌ Failed to list sessions: {e}")

    def show_session_context_menu(self, pos):
        item = self.sessions_list.itemAt(pos)
        if not item:
            return
        menu = QMenu()
        delete_action = menu.addAction(CONFIG["texts"]["delete_session"])
        restore_action = menu.addAction(CONFIG["texts"]["view_trash"]) if "Trash" in item.text() else None
        action = menu.exec_(self.sessions_list.viewport().mapToGlobal(pos))
        if action == delete_action:
            self.delete_session(item)
        elif action == restore_action:
            self.restore_from_trash(item)

    def delete_session(self, item):
        file_path = item.data(Qt.UserRole)
        reply = QMessageBox.question(self, CONFIG["texts"]["confirm_delete"],
                                     CONFIG["texts"]["confirm_delete_msg"].format(file=os.path.basename(file_path)),
                                     QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
        if reply == QMessageBox.Yes:
            try:
                trash_path = os.path.join(self.trash_dir, os.path.basename(file_path))
                if os.path.exists(trash_path):
                    os.remove(trash_path)
                shutil.move(file_path, trash_path)
                self.update_session_list()
                self.status_bar.showMessage(CONFIG["texts"]["session_deleted"].format(file=os.path.basename(file_path)), 5000)
            except (OSError, shutil.Error) as e:
                QMessageBox.critical(self, "Error", f"❌ Failed to delete session: {e}")

    def view_trash(self):
        self.update_session_list()
        self.status_bar.showMessage("♻️ Trash updated", 3000)

    def restore_from_trash(self, item):
        file_path = item.data(Qt.UserRole)
        try:
            shutil.move(file_path, os.path.join(self.session_dir, os.path.basename(file_path)))
            self.update_session_list()
            self.status_bar.showMessage(CONFIG["texts"]["session_restored"].format(file=os.path.basename(file_path)), 5000)
        except (OSError, shutil.Error) as e:
            QMessageBox.critical(self, "Error", f"❌ Failed to restore session: {e}")

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
                    self.status_bar.showMessage(f"❌ Error getting window: {e}", 5000)
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
                    self.status_bar.showMessage(f"❌ Error getting window position: {e}", 5000)
                    continue
        except Exception as e:
            self.status_bar.showMessage(f"❌ Failed to access Shell: {e}", 5000)
        return windows_data

    def copy_to_clipboard(self, text):
        QApplication.clipboard().setText(text)
        self.status_bar.showMessage(f"📋 Copied: {text[:30]}...", 3000)

    def confirm_close_windows(self, state):
        if state == Qt.Checked:
            windows = self.get_explorer_windows()
            window_list = "\n".join([f"Window {w['hwnd']}: {', '.join(w['paths'])}" for w in windows]) or "🚫 No open Explorer windows."
            reply = QMessageBox.question(self, CONFIG["texts"]["confirm_close"],
                                         CONFIG["texts"]["confirm_msg"].format(window_list=window_list),
                                         QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
            if reply == QMessageBox.Yes:
                self.close_existing_windows()
            else:
                self.close_windows_cb.setChecked(False)

    def close_existing_windows(self):
        try:
            shell = win32com.client.Dispatch("Shell.Application")
            for window in shell.Windows():
                try:
                    if hasattr(window, 'HWND'):
                        win32gui.PostMessage(window.HWND, 0x0010, 0, 0)
                except Exception as e:
                    self.status_bar.showMessage(f"❌ Error closing window: {e}", 5000)
                    continue
            self.close_windows_cb.setChecked(False)
            self.status_bar.showMessage("🛑 All Explorer windows closed", 3000)
        except Exception as e:
            self.status_bar.showMessage(f"❌ Failed to close windows: {e}", 5000)

    def expand_all(self):
        self.preview_tree.expandAll()
        self.status_bar.showMessage("⬆️ All items expanded", 3000)

    def collapse_all(self):
        self.preview_tree.collapseAll()
        self.status_bar.showMessage("⬇️ All items collapsed", 3000)

    def clear_preview(self):
        self.preview_tree.clear()
        self.search_input.clear()
        self.item_count_label.setText(CONFIG["texts"]["item_count"].format(count=0))
        self.status_bar.showMessage("🗑️ Preview cleared", 3000)

    def filter_preview(self, text):
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
        self.preview_tree.clear()
        self.search_input.clear()
        if not item:
            self.item_count_label.setText(CONFIG["texts"]["item_count"].format(count=0))
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
                    window_item.setText(0, f"🖥️ Window {hwnd} ({len(paths)} paths)")
                    window_item.setIcon(0, self.style().standardIcon(QStyle.SP_DirIcon))
                    window_item.setExpanded(item_count < CONFIG["max_visible_items"])
                    item_count += 1
                    for path in paths:
                        path_item = QTreeWidgetItem(window_item)
                        path_item.setText(0, f"📁 {os.path.basename(path)}")
                        path_item.setText(1, path)
                        path_item.setIcon(0, self.style().standardIcon(QStyle.SP_FileIcon))
                        item_count += 1
                        if window_item.childCount() % 2 == 0:
                            path_item.setBackground(0, QColor(CONFIG["styles"][self.theme]["row_alt_color"]))
                            path_item.setBackground(1, QColor(CONFIG["styles"][self.theme]["row_alt_color"]))
                self.item_count_label.setText(CONFIG["texts"]["item_count"].format(count=item_count))
        except (OSError, json.JSONDecodeError) as e:
            QMessageBox.critical(self, "Error", f"❌ Failed to load session: {e}")
            self.item_count_label.setText(CONFIG["texts"]["item_count"].format(count=0))

    def open_folder(self, path):
        if os.path.exists(path):
            try:
                shell = win32com.client.Dispatch("Shell.Application")
                shell.Explore(path)
                self.status_bar.showMessage(f"📂 Opened folder: {os.path.basename(path)}", 3000)
            except Exception as e:
                self.status_bar.showMessage(f"❌ Error opening folder: {e}", 5000)
        else:
            self.status_bar.showMessage(CONFIG["texts"]["invalid_path"].format(path=path), 5000)

    def save_session(self):
        file_name, _ = QFileDialog.getSaveFileName(self, "Save Session",
                                                   os.path.join(self.session_dir, f"explorer_session_{datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}.json"),
                                                   "JSON Files (*.json)")
        if file_name:
            try:
                with open(file_name, 'w', encoding='utf-8') as f:
                    json.dump(self.get_explorer_windows(), f, indent=4)
                self.update_session_list()
                self.status_bar.showMessage(CONFIG["texts"]["session_saved"].format(file=os.path.basename(file_name)), 5000)
            except OSError as e:
                QMessageBox.critical(self, "Error", f"❌ Failed to save session: {e}")

    def restore_session(self):
        if item := self.sessions_list.currentItem():
            file_path = item.data(Qt.UserRole)
            try:
                with open(file_path, 'r', encoding='utf-8') as f:
                    session_data = json.load(f)
                    shell = win32com.client.Dispatch("Shell.Application")
                    for entry in session_data:
                        if not entry['paths']:
                            continue
                        for path in entry['paths']:
                            if os.path.exists(path):
                                window = shell.Explore(path)
                                time.sleep(max(0.2, self.timeout_spin.value() / 100.0))
                                win32gui.MoveWindow(window.HWND, entry['x'], entry['y'], entry['width'], entry['height'], True)
                                self.status_bar.showMessage(
                                    f"{CONFIG['texts']['restored_window'].format(hwnd=entry['hwnd'])}", 3000)
                            else:
                                self.status_bar.showMessage(CONFIG["texts"]["invalid_path"].format(path=path), 5000)
            except (OSError, json.JSONDecodeError) as e:
                QMessageBox.critical(self, "Error", CONFIG["texts"]["error_restoring"].format(error=str(e)))

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())