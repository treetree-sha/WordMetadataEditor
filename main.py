import sys
import os

from PySide6.QtCore import Qt, QSize
from PySide6.QtGui import QIcon
from PySide6.QtWidgets import QApplication

from qfluentwidgets import (
    FluentWindow, NavigationItemPosition, FluentIcon, setTheme, Theme,
    setThemeColor, isDarkTheme
)

from ui.single_file_interface import SingleFileInterface
from ui.batch_interface import BatchInterface
from ui.settings_interface import SettingsInterface


class MainWindow(FluentWindow):
    """Main Application Window with Win11 Fluent Design Navigation Sidebar."""

    def __init__(self):
        # Set default theme to Dark mode for stunning aesthetics
        setTheme(Theme.DARK)
        setThemeColor('#0078D4') # Windows Accent Blue

        super().__init__()
        self.setWindowTitle("Word 文档属性高级编辑器 (Word Metadata Editor)")
        self.resize(1080, 750)

        # Center window on screen
        desktop = QApplication.primaryScreen().availableGeometry()
        w, h = desktop.width(), desktop.height()
        self.move(int((w - self.width()) / 2), int((h - self.height()) / 2))

        # Set navigation width so sidebar text is fully displayed
        self.navigationInterface.setExpandWidth(220)
        self.navigationInterface.setMinimumExpandWidth(200)

        self.update_theme_style()
        self._init_navigation()

    def update_theme_style(self):
        if isDarkTheme():
            bg_color = "#202020"
        else:
            bg_color = "#f9f9f9"

        self.setStyleSheet(f"""
            FluentWindow {{
                background-color: {bg_color};
            }}
            QStackedWidget {{
                background-color: {bg_color};
            }}
            QScrollArea {{
                border: none;
                background-color: transparent;
            }}
        """)

    def _init_navigation(self):
        # 1. Single File Interface
        self.single_interface = SingleFileInterface(self)
        self.single_interface.setObjectName("single_interface")
        self.addSubInterface(
            self.single_interface,
            FluentIcon.DOCUMENT,
            '单文件属性精修'
        )

        # 2. Batch Interface
        self.batch_interface = BatchInterface(self)
        self.batch_interface.setObjectName("batch_interface")
        self.addSubInterface(
            self.batch_interface,
            FluentIcon.FOLDER,
            '批量处理与脱敏'
        )

        # 3. Settings Interface
        self.settings_interface = SettingsInterface(self)
        self.settings_interface.setObjectName("settings_interface")
        self.addSubInterface(
            self.settings_interface,
            FluentIcon.SETTING,
            '设置与关于',
            NavigationItemPosition.BOTTOM
        )


if __name__ == '__main__':
    # Enable high DPI scaling
    QApplication.setHighDpiScaleFactorRoundingPolicy(
        Qt.HighDpiScaleFactorRoundingPolicy.PassThrough
    )

    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())

