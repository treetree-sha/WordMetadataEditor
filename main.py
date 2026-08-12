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
from ui.i18n import i18n


def get_asset_path(relative_path: str) -> str:
    """Get absolute path to resource, works for dev and for PyInstaller --onefile."""
    if hasattr(sys, '_MEIPASS'):
        return os.path.join(sys._MEIPASS, relative_path)
    return os.path.join(os.path.dirname(__file__), relative_path)


class MainWindow(FluentWindow):
    """Main Application Window with Win11 Fluent Design Navigation Sidebar."""

    def __init__(self):
        # Set default theme to Dark mode for stunning aesthetics
        setTheme(Theme.DARK)
        setThemeColor('#0078D4') # Windows Accent Blue

        super().__init__()
        self.resize(1080, 750)

        # Center window on screen
        desktop = QApplication.primaryScreen().availableGeometry()
        w, h = desktop.width(), desktop.height()
        self.move(int((w - self.width()) / 2), int((h - self.height()) / 2))

        # Set navigation width so sidebar text is fully displayed
        self.navigationInterface.setExpandWidth(220)
        self.navigationInterface.setMinimumExpandWidth(200)

        # Set Window Icon
        icon_path = get_asset_path(os.path.join('assets', 'app_icon.png'))
        if os.path.exists(icon_path):
            self.setWindowIcon(QIcon(icon_path))

        self.update_theme_style()
        self._init_navigation()
        self.retranslate_ui()

        i18n.languageChanged.connect(self.retranslate_ui)

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
        self.item_single = self.addSubInterface(
            self.single_interface,
            FluentIcon.DOCUMENT,
            i18n.t('单文件属性精修', 'Single File Editor')
        )

        # 2. Batch Interface
        self.batch_interface = BatchInterface(self)
        self.batch_interface.setObjectName("batch_interface")
        self.item_batch = self.addSubInterface(
            self.batch_interface,
            FluentIcon.FOLDER,
            i18n.t('批量处理与脱敏', 'Batch Processing')
        )

        # 3. Settings Interface
        self.settings_interface = SettingsInterface(self)
        self.settings_interface.setObjectName("settings_interface")
        self.item_settings = self.addSubInterface(
            self.settings_interface,
            FluentIcon.SETTING,
            i18n.t('设置与关于', 'Settings & About'),
            NavigationItemPosition.BOTTOM
        )

    def retranslate_ui(self):
        self.setWindowTitle(
            i18n.t("Word 文档属性高级编辑器 (Word Metadata Editor)", "Word Metadata Editor")
        )
        if hasattr(self, 'item_single') and self.item_single:
            self.item_single.setText(i18n.t('单文件属性精修', 'Single File Editor'))
        if hasattr(self, 'item_batch') and self.item_batch:
            self.item_batch.setText(i18n.t('批量处理与脱敏', 'Batch Processing'))
        if hasattr(self, 'item_settings') and self.item_settings:
            self.item_settings.setText(i18n.t('设置与关于', 'Settings & About'))


if __name__ == '__main__':
    # Enable high DPI scaling
    QApplication.setHighDpiScaleFactorRoundingPolicy(
        Qt.HighDpiScaleFactorRoundingPolicy.PassThrough
    )

    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())
