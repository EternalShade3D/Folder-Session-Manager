import sys
import os
import json
import time
from datetime import datetime
import win32gui
import win32com.client
from PySide6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QLabel,
                              QCheckBox, QSpinBox, QListWidget, QFileDialog, QTreeWidget, QTreeWidgetItem, QToolBar,
                              QToolTip, QMessageBox, QSplitter)
from PySide6.QtCore import Qt, QRect, QSize, QEvent, Signal, QSettings
from PySide6.QtGui import QPainter, QColor, QIcon, QClipboard

# Configurações Otimizadas
CONFIG = {
    "styles": {
        "bg": "#121212", "text": "#e0e0e0", "button": "#2a2a2a", "button_hover": "#333333",
        "accent": "#2979ff", "input_bg": "#1f1f1f", "folder_color": "#4CAF50", "subfolder_color": "#FF9800",
        "row_alt_color": "#1a1a1a", "checkbox_checked": "#4CAF50"
    },
    "sizes": {
        "min": (500, 400), "base": (600, 400), "font": (12, 14), "padding": 8, "button_min_width": 100,
        "checkbox_size": 16
    },
    "texts": {
        "app_title": "Folder Session Manager", "config_label": "Settings",
        "close_windows_cb": "Close Windows (Caution!)", "timeout_label": "Timeout (s):",
        "backup_cb": "Auto Backup", "interval_label": "Interval (min):",
        "sessions_label": "Saved Sessions", "preview_label": "Folder Preview",
        "save_button": "Save", "restore_button": "Restore", "no_paths": "No paths found.",
        "error_msg": "Error: ", "copy_icon": "📋", "expand_all": "Expand All",
        "collapse_all": "Collapse All", "toggle_compact": "Toggle Compact",
        "confirm_close": "Confirm Close", "confirm_msg": "Close all windows? Backup first?",
        "yes_backup": "Yes + Backup", "yes_no_backup": "Yes", "no": "No"
    },
    "truncate_length": 30
}

class MiniMapLabel(QLabel):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.windows_data = []

    def set_windows_data(self, data):
        self.windows_data = data
        self.update()

    def paintEvent(self, event):
        painter = QPainter(self)
        try:
            painter.fillRect(self.rect(), QColor(CONFIG["styles"]["bg"]))
            if not self.windows_data:
                return
            screen = QApplication.primaryScreen().geometry()
            scale_x, scale_y = self.width() / screen.width(), self.height() / screen.height()
            for w in self.windows_data:
                rect = QRect(int(w['x'] * scale_x), int(w['y'] * scale_y), int(w['width'] * scale_x), int(w['height'] * scale_y))
                painter.fillRect(rect, QColor(CONFIG["styles"]["accent"]))
                painter.setPen(QColor(CONFIG["styles"]["text"]))
                painter.drawText(rect, Qt.AlignCenter, str(w['hwnd'])[:4])
        finally:
            painter.end()

class CustomTreeWidget(QTreeWidget):
    copyRequested = Signal(str)

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setHeaderHidden(True)
        self.itemClicked.connect(self.on_item_clicked)
        self.viewport().installEventFilter(self)

    def eventFilter(self, obj, event):
        if event.type() == QEvent.HoverMove and obj is self.viewport():
            item = self.itemAt(event.position().toPoint())
            if item and item.text(1):
                QToolTip.showText(event.globalPosition().toPoint(), item.text(1), self)
        return super().eventFilter(obj, event)

    def on_item_clicked(self, item, column):
        if item.text(column).startswith(CONFIG["texts"]["copy_icon"]):
            self.copyRequested.emit(item.parent().text(1) if item.parent() else item.text(1))

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.settings = QSettings("xAI", "FolderSessionManager")
        self.restore_geometry()
        self.setWindowTitle(CONFIG["texts"]["app_title"])
        self.setMinimumSize(*CONFIG["sizes"]["min"])
        self.base_width, self.base_height = CONFIG["sizes"]["base"]
        self.font_scale = self.settings.value("font_scale", 1.0, float)
        self.compact_view = False

        central = QWidget()
        self.setCentralWidget(central)
        main_layout = QVBoxLayout(central)
        main_layout.setContentsMargins(CONFIG["sizes"]["padding"], CONFIG["sizes"]["padding"],
                                      CONFIG["sizes"]["padding"], CONFIG["sizes"]["padding"])
        main_layout.setSpacing(2)

        # Título
        main_layout.addWidget(QLabel(CONFIG["texts"]["app_title"], objectName="title_label",
                                    styleSheet=f"font-size: {int(CONFIG['sizes']['font'][1] * 1.5)}pt; font-weight: bold; color: {CONFIG['styles']['accent']}"))

        # Configurações
        config_layout = QVBoxLayout()
        config_layout.setSpacing(4)
        config_layout.addWidget(QLabel(CONFIG["texts"]["config_label"], objectName="config_label"))
        config_layout.addWidget(QCheckBox(CONFIG["texts"]["close_windows_cb"],
                                        checked=self.settings.value("close_windows", False, bool),
                                        stateChanged=self.confirm_close_windows))
        config_layout.addLayout(self._create_spin_layout(CONFIG["texts"]["timeout_label"],
                                                       self._init_spin(1, 60, 10, "timeout")))
        config_layout.addWidget(QCheckBox(CONFIG["texts"]["backup_cb"],
                                        checked=self.settings.value("backup", True, bool)))
        config_layout.addLayout(self._create_spin_layout(CONFIG["texts"]["interval_label"],
                                                       self._init_spin(1, 60, 15, "interval")))
        main_layout.addLayout(config_layout)

        # Sessões
        main_layout.addWidget(QLabel(CONFIG["texts"]["sessions_label"], objectName="sessions_label"))
        self.sessions_list = QListWidget(itemClicked=self.update_preview)
        main_layout.addWidget(self.sessions_list)

        # Preview e Minimap com Splitter
        splitter = QSplitter(Qt.Vertical)
        splitter.setHandleWidth(6)
        self.preview_tree = CustomTreeWidget()
        self.preview_tree.copyRequested.connect(self.copy_to_clipboard)
        splitter.addWidget(self.preview_tree)
        self.minimap = MiniMapLabel()
        splitter.addWidget(self.minimap)
        main_layout.addWidget(splitter)

        # Ações
        action_layout = QHBoxLayout()
        action_layout.setSpacing(2)
        self.save_button = QPushButton(CONFIG["texts"]["save_button"])
        self.restore_button = QPushButton(CONFIG["texts"]["restore_button"])
        view_toolbar = QToolBar("View")
        view_toolbar.setStyleSheet("QToolBar { background: transparent; border: none; spacing: 2px; }")
        expand_action = view_toolbar.addAction(CONFIG["texts"]["expand_all"])
        collapse_action = view_toolbar.addAction(CONFIG["texts"]["collapse_all"])
        toggle_action = view_toolbar.addAction(CONFIG["texts"]["toggle_compact"])

        action_layout.addWidget(self.save_button)
        action_layout.addWidget(self.restore_button)
        action_layout.addWidget(view_toolbar)
        main_layout.addLayout(action_layout)

        # Conexões
        self.save_button.clicked.connect(self.save_session)
        self.restore_button.clicked.connect(self.restore_session)
        expand_action.triggered.connect(self.expand_all)
        collapse_action.triggered.connect(self.collapse_all)
        toggle_action.triggered.connect(self.toggle_compact_view)

        self.update_session_list()
        self.apply_dark_theme()

    def _init_spin(self, min_val, max_val, default, key):
        spin = QSpinBox(value=self.settings.value(key, default, int))
        spin.setRange(min_val, max_val)
        return spin

    def _create_spin_layout(self, label_text, spin):
        layout = QHBoxLayout()
        layout.addWidget(QLabel(label_text), stretch=1)
        layout.addWidget(spin, stretch=1)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(4)
        return layout

    def closeEvent(self, event):
        self.save_geometry()
        super().closeEvent(event)

    def save_geometry(self):
        self.settings.setValue("geometry", self.saveGeometry())
        self.settings.setValue("font_scale", self.font_scale)
        self.settings.setValue("close_windows", self.close_windows_cb.isChecked())
        self.settings.setValue("timeout", self.timeout_spin.value())
        self.settings.setValue("backup", self.backup_cb.isChecked())
        self.settings.setValue("interval", self.interval_spin.value())

    def restore_geometry(self):
        if geom := self.settings.value("geometry"):
            self.restoreGeometry(geom)

    def apply_dark_theme(self):
        self.setStyleSheet(self.generate_stylesheet())

    def generate_stylesheet(self):
        font_m, font_l = [max(10, int(x * self.font_scale)) for x in CONFIG["sizes"]["font"]]
        padding = max(6, int(CONFIG["sizes"]["padding"] * self.font_scale))
        checkbox = max(12, int(CONFIG["sizes"]["checkbox_size"] * self.font_scale))

        return f"""
            QWidget {{ background: {CONFIG['styles']['bg']}; color: {CONFIG['styles']['text']}; font-family: 'Segoe UI'; }}
            QLabel#title_label {{ font-size: {int(font_l * 1.5)}pt; font-weight: bold; color: {CONFIG['styles']['accent']}; }}
            QLabel#config_label, QLabel#sessions_label, QLabel#preview_label {{ font-size: {font_l}pt; font-weight: bold; }}
            QCheckBox {{ color: {CONFIG['styles']['text']}; font-size: {font_m}pt; spacing: {int(4 * self.font_scale)}px; }}
            QCheckBox::indicator {{ width: {checkbox}px; height: {checkbox}px; border: 1px solid {CONFIG['styles']['accent']};
                                   background: {CONFIG['styles']['input_bg']}; border-radius: {int(2 * self.font_scale)}px; }}
            QCheckBox::indicator:checked {{ background: {CONFIG['styles']['checkbox_checked']}; border: 1px solid {CONFIG['styles']['checkbox_checked']}; }}
            QSpinBox {{ background: {CONFIG['styles']['input_bg']}; color: {CONFIG['styles']['text']}; border: 1px solid {CONFIG['styles']['accent']};
                        padding: {int(3 * self.font_scale)}px; font-size: {font_m}pt; }}
            QTreeWidget {{ background: {CONFIG['styles']['input_bg']}; color: {CONFIG['styles']['text']}; border: 1px solid {CONFIG['styles']['accent']};
                           font-size: {font_m}pt; }}
            QTreeWidget::item {{ padding: {int(2 * self.font_scale)}px; }}
            QTreeWidget::item:has-children {{ background: {CONFIG['styles']['input_bg']}; }}
            QTreeWidget::item:!has-children:alternate {{ background: {CONFIG['styles']['row_alt_color']}; }}
            QPushButton {{ background: {CONFIG['styles']['button']}; color: {CONFIG['styles']['text']}; border: 1px solid {CONFIG['styles']['accent']};
                           padding: {padding}px; font-size: {font_m}pt; min-width: {int(CONFIG['sizes']['button_min_width'] * self.font_scale)}px; }}
            QPushButton:hover {{ background: {CONFIG['styles']['button_hover']}; }}
            QToolBar {{ background: {CONFIG['styles']['input_bg']}; border: none; spacing: {int(3 * self.font_scale)}px; }}
            QSplitter::handle {{ background: {CONFIG['styles']['accent']}; width: 8px; height: 8px; }}
            QSplitter::handle:horizontal {{ width: 8px; }}
            QSplitter::handle:vertical {{ height: 8px; }}
        """

    def resizeEvent(self, event):
        super().resizeEvent(event)
        self.font_scale = min(max(0.8, min(self.width() / self.base_width, self.height() / self.base_height)), 1.5)
        self.apply_dark_theme()

    def update_session_list(self):
        self.sessions_list.clear()
        [self.sessions_list.addItem(f) for f in os.listdir(os.getcwd()) if f.endswith('.json') and f.startswith('explorer_session')]

    def get_explorer_windows(self):
        windows, shell = [], win32com.client.Dispatch("Shell.Application").Windows()
        path_count = {}
        for window in shell:
            try:
                if not (hasattr(window, 'Document') and hasattr(window.Document, 'Folder')):
                    continue
                hwnd = window.HWND
                path = window.Document.Folder.Self.Path
                path_count.setdefault(hwnd, []).append(path)
            except Exception:
                continue
        for hwnd, paths in path_count.items():
            left, top, right, bottom = win32gui.GetWindowRect(hwnd)
            windows.append({'hwnd': hwnd, 'paths': paths, 'x': left, 'y': top, 'width': right - left, 'height': bottom - top})
        return windows

    def wheelEvent(self, event):
        widget = self.childAt(event.position().toPoint())
        if event.modifiers() & Qt.ControlModifier:
            delta = event.angleDelta().y() / 120
            if widget == self.preview_tree:
                self.font_scale = max(0.5, min(2.0, self.font_scale + delta * 0.1))
            elif widget == self.minimap:
                self.minimap.setMinimumSize(max(80, self.minimap.minimumWidth() + delta * 8),
                                           max(80, self.minimap.minimumHeight() + delta * 8))
            self.apply_dark_theme()
            event.accept()
        else:
            super().wheelEvent(event)

    def expand_all(self): self.preview_tree.expandAll()
    def collapse_all(self): self.preview_tree.collapseAll()
    def toggle_compact_view(self):
        self.compact_view = not self.compact_view
        self.update_preview(self.sessions_list.currentItem())

    def copy_to_clipboard(self, text):
        QApplication.clipboard().setText(text)

    def confirm_close_windows(self, state):
        if state == Qt.Checked:
            reply = QMessageBox.question(self, CONFIG["texts"]["confirm_close"], CONFIG["texts"]["confirm_msg"],
                                         (QMessageBox.Yes | QMessageBox.YesToAll | QMessageBox.No), QMessageBox.No)
            if reply == QMessageBox.Yes:
                self.create_backup()
                self.close_existing_windows()
            elif reply == QMessageBox.YesToAll:
                self.close_existing_windows()
            else:
                self.close_windows_cb.setChecked(False)

    def create_backup(self):
        backup_file = f"backup_explorer_session_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json"
        with open(backup_file, 'w', encoding='utf-8') as f:
            json.dump(self.get_explorer_windows(), f, indent=4)
        print(f"Backup created: {backup_file}")

    def close_existing_windows(self):
        for window in win32com.client.Dispatch("Shell.Application").Windows():
            try:
                if hasattr(window, 'HWND'):
                    win32gui.CloseWindow(window.HWND)
            except Exception:
                continue
        self.close_windows_cb.setChecked(False)

    def update_preview(self, item):
        if not item:
            return
        self.preview_tree.clear()
        file_path = os.path.join(os.getcwd(), item.text())
        try:
            with open(file_path, 'r', encoding='utf-8') as f:
                data = json.load(f)
                if all('hwnd' in d and 'paths' in d for d in data):
                    windows_data = []
                    for entry in data:
                        hwnd, paths, x, y, w, h = entry['hwnd'], entry['paths'], entry['x'], entry['y'], entry['width'], entry['height']
                        window_item = QTreeWidgetItem(self.preview_tree)
                        window_item.setText(0, f"+ Window {hwnd} ({len(paths)} paths)")
                        window_item.setExpanded(False)
                        for path in paths:
                            path_item = QTreeWidgetItem(window_item)
                            max_len = int(CONFIG["truncate_length"] * (self.width() / self.base_width) * 2 if self.width() > self.base_width else CONFIG["truncate_length"])
                            short_path = "/".join(path.split("\\")[-2:]) if self.compact_view else path
                            if len(short_path) > max_len:
                                short_path = short_path[:max_len - 3] + "..."
                            path_item.setText(0, "📁 " + short_path)
                            path_item.setText(1, path)
                            path_item.setText(2, CONFIG["texts"]["copy_icon"])
                            if window_item.childCount() % 2 == 0:
                                path_item.setBackground(0, QColor(CONFIG["styles"]["row_alt_color"]))
                        windows_data.append({'hwnd': hwnd, 'paths': paths, 'x': x, 'y': y, 'width': w, 'height': h})
                    self.minimap.set_windows_data(windows_data)
                else:
                    for path in [entry.get('path', 'Unknown') for entry in data if 'path' in entry]:
                        path_item = QTreeWidgetItem(self.preview_tree)
                        max_len = int(CONFIG["truncate_length"] * (self.width() / self.base_width) * 2 if self.width() > self.base_width else CONFIG["truncate_length"])
                        short_path = "/".join(path.split("\\")[-2:]) if self.compact_view else path
                        if len(short_path) > max_len:
                            short_path = short_path[:max_len - 3] + "..."
                        path_item.setText(0, "📁 " + short_path)
                        path_item.setText(1, path)
                        path_item.setText(2, CONFIG["texts"]["copy_icon"])
                        if self.preview_tree.topLevelItemCount() % 2 == 0:
                            path_item.setBackground(0, QColor(CONFIG["styles"]["row_alt_color"]))
                    self.minimap.set_windows_data([])
        except Exception as e:
            QTreeWidgetItem(self.preview_tree).setText(0, f"{CONFIG['texts']['error_msg']}{str(e)}")
            self.minimap.set_windows_data([])

    def save_session(self):
        if file_path := QFileDialog.getSaveFileName(self, "Save Session", f"explorer_session_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json", "JSON Files (*.json)")[0]:
            with open(file_path, 'w', encoding='utf-8') as f:
                json.dump(self.get_explorer_windows(), f, indent=4)
            self.update_session_list()
            print(f"Session saved to: {file_path}")

    def restore_session(self):
        if item := self.sessions_list.currentItem():
            try:
                with open(os.path.join(os.getcwd(), item.text()), 'r', encoding='utf-8') as f:
                    for entry in json.load(f):
                        shell = win32com.client.Dispatch("Shell.Application")
                        window = shell.Explore(entry['paths'][0]) if entry['paths'] else shell.Explore(os.getcwd())
                        time.sleep(0.1)
                        win32gui.MoveWindow(entry['hwnd'], entry['x'], entry['y'], entry['width'], entry['height'], True)
                        print(f"Restored window {entry['hwnd']} at ({entry['x']}, {entry['y']}) with size ({entry['width']}, {entry['height']})")
            except Exception as e:
                print(f"Error restoring session: {str(e)}")

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())