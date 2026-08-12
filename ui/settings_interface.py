from PySide6.QtCore import Qt
from PySide6.QtWidgets import QWidget, QVBoxLayout, QHBoxLayout

from qfluentwidgets import (
    ScrollArea, CardWidget, TitleLabel, SubtitleLabel, CaptionLabel,
    StrongBodyLabel, RadioButton, setTheme, Theme, FluentIcon, IconWidget,
    BodyLabel
)


class SettingsInterface(ScrollArea):
    """Application Settings and About Information Page."""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.scroll_widget = QWidget()
        self.scroll_widget.setObjectName("settingsScrollWidget")
        self.scroll_widget.setStyleSheet("background-color: transparent;")
        self.setWidget(self.scroll_widget)
        self.setWidgetResizable(True)
        self.setStyleSheet("QScrollArea { border: none; background-color: transparent; }")

        self.main_layout = QVBoxLayout(self.scroll_widget)
        self.main_layout.setSpacing(16)
        self.main_layout.setContentsMargins(36, 36, 36, 36)

        self._init_ui()

    def _init_ui(self):
        # Header Title
        title_label = TitleLabel("设置与软件信息", self)
        subtitle_label = CaptionLabel("自定义界面外观主题及查看软件相关信息", self)
        self.main_layout.addWidget(title_label)
        self.main_layout.addWidget(subtitle_label)

        # Card 1: Theme Settings
        theme_card = CardWidget(self)
        tc_layout = QVBoxLayout(theme_card)
        tc_layout.setContentsMargins(24, 20, 24, 20)
        tc_layout.setSpacing(16)

        tc_layout.addWidget(StrongBodyLabel("界面主题 (Theme)", theme_card))

        self.radio_dark = RadioButton("暗黑模式 (Dark Mode)", theme_card)
        self.radio_light = RadioButton("浅色模式 (Light Mode)", theme_card)

        self.radio_dark.setChecked(True) # Default dark mode for high aesthetics

        self.radio_dark.toggled.connect(self._on_dark_toggled)
        self.radio_light.toggled.connect(self._on_light_toggled)

        tc_layout.addWidget(self.radio_dark)
        tc_layout.addWidget(self.radio_light)

        self.main_layout.addWidget(theme_card)

        # Card 2: About Software
        about_card = CardWidget(self)
        ac_layout = QHBoxLayout(about_card)
        ac_layout.setContentsMargins(24, 20, 24, 20)
        ac_layout.setSpacing(16)

        icon_widget = IconWidget(FluentIcon.INFO, about_card)
        icon_widget.setFixedSize(48, 48)

        info_vbox = QVBoxLayout()
        info_vbox.addWidget(StrongBodyLabel("Word 文档属性高级修改器 v1.0.0", about_card))
        info_vbox.addWidget(CaptionLabel("基于 PySide6 与 Windows 11 Fluent Design 打造的高颜值 Word 元数据修改软件。", about_card))
        info_vbox.addWidget(BodyLabel("支持功能: 作者修改、总编辑时间自定义、系统时间同步、批量属性更新、一键隐私脱敏。", about_card))

        ac_layout.addWidget(icon_widget)
        ac_layout.addLayout(info_vbox, 1)

        self.main_layout.addWidget(about_card)
        self.main_layout.addStretch(1)

    def _on_dark_toggled(self, checked):
        if checked:
            setTheme(Theme.DARK)
            if self.window() and hasattr(self.window(), 'update_theme_style'):
                self.window().update_theme_style()

    def _on_light_toggled(self, checked):
        if checked:
            setTheme(Theme.LIGHT)
            if self.window() and hasattr(self.window(), 'update_theme_style'):
                self.window().update_theme_style()

