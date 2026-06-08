import sys
import json
import os
import win32gui
import win32com.client
from datetime import datetime
from PySide6.QtWidgets import QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QLabel, QCheckBox, QSpinBox, QListWidget, QFileDialog, QTreeWidget, QTreeWidgetItem, QToolBar, QToolTip
from PySide6.QtCore import Qt, QRect, QSize, QEvent, Signal, QSettings
from PySide6.QtGui import QPainter, QColor, QIcon, QClipboard

# Configurações Globalizadas
CONFIG = {
    "colors": {
        "bg": "#121212",
        "text": "#e0e0e0",
        "button": "#2a2a2a",
        "button_hover": "#333333",
        "accent": "#2979ff",
        "input_bg": "#1f1f1f",
        "folder_color": "#4CAF50",
        "subfolder_color": "#FF9800",
        "row_alt_color": "#1a1a1a",
        "checkbox_checked": "#4CAF50",
    },
    "sizes": {
        "min_width": 600,
        "min_height": 400,
        "base_width": 600,
        "base_height": 400,
        "font_size_medium": 11,
        "font_size_large": 13,
        "padding": 12,
        "button_min_width": 120,
        "checkbox_size": 16,
    },
    "texts": {
        "app_title": "Folder Session Manager",
        "config_label": "Settings",
        "close_windows_cb": "Close Existing Windows",
        "timeout_label": "Timeout (seconds):",
        "backup_cb": "Enable Auto Backup",
        "interval_label": "Interval (minutes):",
        "sessions_label": "Saved Sessions",
        "preview_label": "Folder Preview by Window",
        "save_button": "Save Session",
        "restore_button": "Restore Session",
        "no_paths": "No folder paths found in JSON.",
        "error_msg": "Error loading file: ",
        "copy_icon": "📋",
        "expand_all": "Expand All",
        "collapse_all": "Collapse All",
    },
    "truncate_length": 30,
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
            painter.fillRect(self.rect(), QColor(CONFIG["colors"]["bg"]))
            if not self.windows_data:
                return

            screen_geometry = QApplication.primaryScreen().geometry()
            screen_width = screen_geometry.width()
            screen_height = screen_geometry.height()
            map_width = self.width()
            map_height = self.height()
            scale_x = map_width / screen_width
            scale_y = map_height / screen_height

            for window in self.windows_data:
                x, y, width, height = window['x'], window['y'], window['width'], window['height']
                rect = QRect(int(x * scale_x), int(y * scale_y), int(width * scale_x), int(height * scale_y))
                painter.fillRect(rect, QColor(CONFIG["colors"]["accent"]))
                painter.setPen(QColor(CONFIG["colors"]["text"]))
                painter.drawText(rect, Qt.AlignCenter, str(window['hwnd'])[:4])
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
            item = self.itemAt(event.pos())
            if item and item.text(1):  # Full path exists
                QToolTip.showText(event.globalPos(), item.text(1), self)
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
        self.setMinimumSize(CONFIG["sizes"]["min_width"], CONFIG["sizes"]["min_height"])
        self.base_width = CONFIG["sizes"]["base_width"]
        self.base_height = CONFIG["sizes"]["base_height"]
        self.font_scale = self.settings.value("font_scale", 1.0, float)
        self.compact_view = False

        # Widget central e layout principal
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)
        main_layout.setContentsMargins(CONFIG["sizes"]["padding"], CONFIG["sizes"]["padding"], CONFIG["sizes"]["padding"], CONFIG["sizes"]["padding"])

        # Título proeminente na interface
        title_label = QLabel(CONFIG["texts"]["app_title"])
        title_label.setObjectName("title_label")
        title_label.setStyleSheet(f"font-size: {int(CONFIG['sizes']['font_size_large'] * 1.5)}pt; font-weight: bold; color: {CONFIG['colors']['accent']};")
        main_layout.addWidget(title_label)

        # Seção de Configurações
        config_label = QLabel(CONFIG["texts"]["config_label"])
        config_label.setObjectName("config_label")
        main_layout.addWidget(config_label)

        self.close_windows_cb = QCheckBox(CONFIG["texts"]["close_windows_cb"])
        self.close_windows_cb.setChecked(self.settings.value("close_windows", False, bool))
        main_layout.addWidget(self.close_windows_cb)

        timeout_layout = QHBoxLayout()
        timeout_label = QLabel(CONFIG["texts"]["timeout_label"])
        self.timeout_spin = QSpinBox()
        self.timeout_spin.setRange(1, 60)
        self.timeout_spin.setValue(self.settings.value("timeout", 10, int))
        timeout_layout.addWidget(timeout_label)
        timeout_layout.addWidget(self.timeout_spin)
        main_layout.addLayout(timeout_layout)

        self.backup_cb = QCheckBox(CONFIG["texts"]["backup_cb"])
        self.backup_cb.setChecked(self.settings.value("backup", True, bool))
        main_layout.addWidget(self.backup_cb)

        interval_layout = QHBoxLayout()
        interval_label = QLabel(CONFIG["texts"]["interval_label"])
        self.interval_spin = QSpinBox()
        self.interval_spin.setRange(1, 60)
        self.interval_spin.setValue(self.settings.value("interval", 15, int))
        interval_layout.addWidget(interval_label)
        interval_layout.addWidget(self.interval_spin)
        main_layout.addLayout(interval_layout)

        # Seção de Sessões Salvas
        sessions_label = QLabel(CONFIG["texts"]["sessions_label"])
        sessions_label.setObjectName("sessions_label")
        main_layout.addWidget(sessions_label)

        self.sessions_list = QListWidget()
        self.sessions_list.itemClicked.connect(self.update_preview)
        main_layout.addWidget(self.sessions_list)

        # Área de pré-visualização
        preview_label = QLabel(CONFIG["texts"]["preview_label"])
        preview_label.setObjectName("preview_label")
        main_layout.addWidget(preview_label)

        self.preview_tree = CustomTreeWidget()
        self.preview_tree.copyRequested.connect(self.copy_to_clipboard)
        main_layout.addWidget(self.preview_tree, stretch=1)

        # Mini-mapa
        self.minimap = MiniMapLabel()
        main_layout.addWidget(self.minimap)

        # Barra de ferramentas
        toolbar = QToolBar()
        expand_action = toolbar.addAction(CONFIG["texts"]["expand_all"])
        expand_action.triggered.connect(self.expand_all)
        collapse_action = toolbar.addAction(CONFIG["texts"]["collapse_all"])
        collapse_action.triggered.connect(self.collapse_all)
        compact_action = toolbar.addAction("Toggle Compact View")
        compact_action.triggered.connect(self.toggle_compact_view)
        main_layout.addWidget(toolbar)

        # Botões
        button_layout = QHBoxLayout()
        self.save_button = QPushButton(CONFIG["texts"]["save_button"])
        self.restore_button = QPushButton(CONFIG["texts"]["restore_button"])
        button_layout.addWidget(self.save_button)
        button_layout.addWidget(self.restore_button)
        button_layout.addStretch()
        main_layout.addLayout(button_layout)

        self.update_session_list()
        self.apply_dark_theme()

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
        geometry = self.settings.value("geometry")
        if geometry:
            self.restoreGeometry(geometry)

    def apply_dark_theme(self):
        self.setStyleSheet(self.generate_stylesheet(1.0))

    def generate_stylesheet(self, scaling_factor):
        font_size_medium = max(8, int(CONFIG["sizes"]["font_size_medium"] * scaling_factor))
        font_size_large = max(10, int(CONFIG["sizes"]["font_size_large"] * scaling_factor))
        padding_md = max(6, int(CONFIG["sizes"]["padding"] * scaling_factor))
        checkbox_size = max(10, int(CONFIG["sizes"]["checkbox_size"] * scaling_factor))

        return f"""
            QWidget {{
                background-color: {CONFIG['colors']['bg']};
                color: {CONFIG['colors']['text']};
                font-family: 'Segoe UI', sans-serif;
            }}
            QLabel#title_label {{
                font-size: {int(CONFIG['sizes']['font_size_large'] * 1.5 * scaling_factor)}pt;
                font-weight: bold;
                color: {CONFIG['colors']['accent']};
            }}
            QLabel#config_label, QLabel#sessions_label, QLabel#preview_label {{
                font-size: {font_size_large}pt;
                font-weight: bold;
            }}
            QCheckBox {{
                color: {CONFIG['colors']['text']};
                font-size: {font_size_medium}pt;
                spacing: {int(5 * scaling_factor)}px;
            }}
            QCheckBox::indicator {{
                width: {checkbox_size}px;
                height: {checkbox_size}px;
                border: 2px solid {CONFIG['colors']['accent']};
                background-color: {CONFIG['colors']['input_bg']};
                border-radius: {int(2 * scaling_factor)}px;
            }}
            QCheckBox::indicator:checked {{
                background-color: {CONFIG['colors']['checkbox_checked']};
                border: 2px solid {CONFIG['colors']['checkbox_checked']};
            }}
            QSpinBox {{
                background-color: {CONFIG['colors']['input_bg']};
                color: {CONFIG['colors']['text']};
                border: 2px solid {CONFIG['colors']['accent']};
                padding: {int(4 * scaling_factor)}px;
                font-size: {font_size_medium}pt;
            }}
            QTreeWidget {{
                background-color: {CONFIG['colors']['input_bg']};
                color: {CONFIG['colors']['text']};
                border: 2px solid {CONFIG['colors']['accent']};
                font-size: {font_size_medium}pt;
            }}
            QTreeWidget::item {{
                padding: {int(2 * scaling_factor)}px;
            }}
            QTreeWidget::item:has-children {{
                background-color: {CONFIG['colors']['input_bg']};
            }}
            QTreeWidget::item:!has-children:alternate {{
                background-color: {CONFIG['colors']['row_alt_color']};
            }}
            QPushButton {{
                background-color: {CONFIG['colors']['button']};
                color: {CONFIG['colors']['text']};
                border: 2px solid {CONFIG['colors']['accent']};
                padding: {padding_md}px;
                font-size: {font_size_medium}pt;
                min-width: {int(CONFIG['sizes']['button_min_width'] * scaling_factor)}px;
            }}
            QPushButton:hover {{
                background-color: {CONFIG['colors']['button_hover']};
            }}
        """

    def resizeEvent(self, event):
        super().resizeEvent(event)
        scaling_factor = min(max(0.8, min(self.width() / self.base_width, self.height() / self.base_height)), 1.2)
        self.setStyleSheet(self.generate_stylesheet(scaling_factor))

    def update_session_list(self):
        self.sessions_list.clear()
        current_dir = os.getcwd()
        session_files = [f for f in os.listdir(current_dir) if f.endswith('.json') and f.startswith('explorer_session')]
        for session_file in session_files:
            self.sessions_list.addItem(session_file)

    def get_explorer_windows(self):
        windows = []
        shell = win32com.client.Dispatch("Shell.Application").Windows()
        path_count = {}
        for window in shell:
            try:
                if not hasattr(window, 'Document') or not hasattr(window.Document, 'Folder'):
                    continue
                hwnd = window.HWND
                path = window.Document.Folder.Self.Path
                if hwnd not in path_count:
                    path_count[hwnd] = []
                path_count[hwnd].append(path)
            except Exception:
                continue
        for hwnd, paths in path_count.items():
            left, top, right, bottom = win32gui.GetWindowRect(hwnd)
            width = right - left
            height = bottom - top
            windows.append({
                'hwnd': hwnd,
                'paths': paths,
                'x': left,
                'y': top,
                'width': width,
                'height': height
            })
        return windows

    def wheelEvent(self, event):
        widget_under_cursor = self.childAt(event.pos())
        if event.modifiers() & Qt.ControlModifier:
            delta = event.angleDelta().y() / 120
            if widget_under_cursor == self.preview_tree:
                self.font_scale = max(0.5, min(2.0, self.font_scale + delta * 0.1))
            elif widget_under_cursor == self.minimap:
                self.minimap.setMinimumSize(
                    max(100, self.minimap.minimumWidth() + delta * 10),
                    max(100, self.minimap.minimumHeight() + delta * 10)
                )
            self.setStyleSheet(self.generate_stylesheet(1.0))
            event.accept()
        else:
            super().wheelEvent(event)

    def expand_all(self):
        self.preview_tree.expandAll()

    def collapse_all(self):
        self.preview_tree.collapseAll()

    def toggle_compact_view(self):
        self.compact_view = not self.compact_view
        self.update_preview(self.sessions_list.currentItem())

    def copy_to_clipboard(self, text):
        clipboard = QApplication.clipboard()
        clipboard.setText(text)

    def update_preview(self, item):
        if not item:
            return
        self.preview_tree.clear()
        file_path = os.path.join(os.getcwd(), item.text())
        try:
            with open(file_path, 'r', encoding='utf-8') as f:
                data = json.load(f)
                if isinstance(data, list) and all('hwnd' in d and 'paths' in d for d in data):
                    windows_data = []
                    for entry in data:
                        hwnd = entry['hwnd']
                        paths = entry['paths']
                        x, y, width, height = entry['x'], entry['y'], entry['width'], entry['height']
                        window_item = QTreeWidgetItem(self.preview_tree)
                        window_item.setText(0, f"+ Window {hwnd} ({len(paths)} paths)")
                        window_item.setExpanded(False)
                        for path in paths:
                            path_item = QTreeWidgetItem(window_item)
                            max_length = int(CONFIG["truncate_length"] * (self.width() / CONFIG["sizes"]["base_width"]) * 2) if self.width() > CONFIG["sizes"]["base_width"] else CONFIG["truncate_length"]
                            short_path = "/".join(path.split("\\")[-2:]) if self.compact_view else path
                            if len(short_path) > max_length:
                                short_path = short_path[:max_length - 3] + "..."
                            path_item.setText(0, "📁 " + short_path)
                            path_item.setText(1, path)  # Full path for tooltip
                            path_item.setText(2, CONFIG["texts"]["copy_icon"])  # Copy icon
                            if window_item.childCount() % 2 == 0:
                                path_item.setBackground(0, QColor(CONFIG["colors"]["row_alt_color"]))
                        windows_data.append({'hwnd': hwnd, 'paths': paths, 'x': x, 'y': y, 'width': width, 'height': height})
                    self.minimap.set_windows_data(windows_data)
                else:
                    paths = [entry.get('path', 'Unknown') for entry in data if 'path' in entry]
                    for path in paths:
                        path_item = QTreeWidgetItem(self.preview_tree)
                        max_length = int(CONFIG["truncate_length"] * (self.width() / CONFIG["sizes"]["base_width"]) * 2) if self.width() > CONFIG["sizes"]["base_width"] else CONFIG["truncate_length"]
                        short_path = "/".join(path.split("\\")[-2:]) if self.compact_view else path
                        if len(short_path) > max_length:
                            short_path = short_path[:max_length - 3] + "..."
                        path_item.setText(0, "📁 " + short_path)
                        path_item.setText(1, path)
                        path_item.setText(2, CONFIG["texts"]["copy_icon"])
                        if self.preview_tree.topLevelItemCount() % 2 == 0:
                            path_item.setBackground(0, QColor(CONFIG["colors"]["row_alt_color"]))
                    self.minimap.set_windows_data([])
        except Exception as e:
            error_item = QTreeWidgetItem(self.preview_tree)
            error_item.setText(0, f"{CONFIG['texts']['error_msg']}{str(e)}")
            self.minimap.set_windows_data([])

    def save_session(self):
        default_file = f"explorer_session_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json"
        file_path, _ = QFileDialog.getSaveFileName(self, "Save Session", default_file, "JSON Files (*.json)")
        if file_path:
            session_data = self.get_explorer_windows()
            with open(file_path, 'w', encoding='utf-8') as f:
                json.dump(session_data, f, indent=4)
            self.update_session_list()

    def restore_session(self):
        selected_item = self.sessions_list.currentItem()
        if selected_item:
            file_path = os.path.join(os.getcwd(), selected_item.text())
            print(f"Restoring session from {file_path}")

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.save_button.clicked.connect(window.save_session)
    window.restore_button.clicked.connect(window.restore_session)
    window.show()
    sys.exit(app.exec())