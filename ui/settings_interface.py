from PySide6.QtCore import Qt
from PySide6.QtWidgets import QWidget, QVBoxLayout, QHBoxLayout

from qfluentwidgets import (
    ScrollArea, CardWidget, TitleLabel, SubtitleLabel, CaptionLabel,
    StrongBodyLabel, RadioButton, setTheme, Theme, FluentIcon, IconWidget,
    BodyLabel, HyperlinkButton
)

from ui.i18n import i18n


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
        i18n.languageChanged.connect(self.retranslate_ui)

    def _init_ui(self):
        # Header Title
        self.title_label = TitleLabel(i18n.t("设置与软件信息", "Settings & About"), self)
        self.subtitle_label = CaptionLabel(i18n.t("自定义界面外观主题、语言及查看软件与作者信息", "Customize theme, language, and view software & author information"), self)
        self.main_layout.addWidget(self.title_label)
        self.main_layout.addWidget(self.subtitle_label)

        # Card 1: Theme Settings
        theme_card = CardWidget(self)
        tc_layout = QVBoxLayout(theme_card)
        tc_layout.setContentsMargins(24, 20, 24, 20)
        tc_layout.setSpacing(16)

        self.theme_card_title = StrongBodyLabel(i18n.t("界面主题 (Theme)", "Interface Theme"), theme_card)
        tc_layout.addWidget(self.theme_card_title)

        self.radio_dark = RadioButton(i18n.t("暗黑模式 (Dark Mode)", "Dark Mode"), theme_card)
        self.radio_light = RadioButton(i18n.t("浅色模式 (Light Mode)", "Light Mode"), theme_card)
        self.radio_dark.setChecked(True)

        self.radio_dark.toggled.connect(self._on_dark_toggled)
        self.radio_light.toggled.connect(self._on_light_toggled)

        tc_layout.addWidget(self.radio_dark)
        tc_layout.addWidget(self.radio_light)
        self.main_layout.addWidget(theme_card)

        # Card 2: Language Settings
        lang_card = CardWidget(self)
        lc_layout = QVBoxLayout(lang_card)
        lc_layout.setContentsMargins(24, 20, 24, 20)
        lc_layout.setSpacing(16)

        self.lang_card_title = StrongBodyLabel(i18n.t("界面语言 (Language)", "Language Settings"), lang_card)
        lc_layout.addWidget(self.lang_card_title)

        self.radio_zh = RadioButton("简体中文 (Simplified Chinese)", lang_card)
        self.radio_en = RadioButton("English", lang_card)

        if i18n.is_english():
            self.radio_en.setChecked(True)
        else:
            self.radio_zh.setChecked(True)

        self.radio_zh.toggled.connect(lambda checked: i18n.set_language('zh') if checked else None)
        self.radio_en.toggled.connect(lambda checked: i18n.set_language('en') if checked else None)

        lc_layout.addWidget(self.radio_zh)
        lc_layout.addWidget(self.radio_en)
        self.main_layout.addWidget(lang_card)

        # Card 3: About Software & Author Info
        about_card = CardWidget(self)
        ac_layout = QHBoxLayout(about_card)
        ac_layout.setContentsMargins(24, 20, 24, 20)
        ac_layout.setSpacing(16)

        icon_widget = IconWidget(FluentIcon.INFO, about_card)
        icon_widget.setFixedSize(48, 48)

        info_vbox = QVBoxLayout()
        self.app_name_label = StrongBodyLabel(i18n.t("Word 文档属性高级修改器 v1.0.0", "Word Metadata Editor v1.0.0"), about_card)
        self.app_desc_label = CaptionLabel(
            i18n.t(
                "基于 PySide6 与 Windows 11 Fluent Design 打造的高颜值 Word 元数据修改软件。",
                "A modern, elegant Word metadata editor built with PySide6 & Windows 11 Fluent Design."
            ),
            about_card
        )
        self.app_feature_label = BodyLabel(
            i18n.t(
                "支持功能: 作者修改、总编辑时间自定义、系统时间同步、批量属性更新、一键隐私脱敏。",
                "Features: Author editing, editing time customization, OS timestamp sync, batch processing, and anonymization."
            ),
            about_card
        )

        # Author section
        self.author_title_label = StrongBodyLabel(i18n.t("软件作者信息 (Author Information)", "Author Information"), about_card)
        self.author_name_label = BodyLabel(i18n.t("开发者: treetree-sha", "Developer: treetree-sha"), about_card)

        links_layout = QHBoxLayout()
        links_layout.setSpacing(12)

        self.github_profile_btn = HyperlinkButton(
            "https://github.com/treetree-sha",
            i18n.t("GitHub 主页 (@treetree-sha)", "GitHub Profile (@treetree-sha)"),
            about_card,
            FluentIcon.GITHUB
        )
        self.github_repo_btn = HyperlinkButton(
            "https://github.com/treetree-sha/WordMetadataEditor",
            i18n.t("GitHub 开源仓库 (WordMetadataEditor)", "GitHub Repository (WordMetadataEditor)"),
            about_card,
            FluentIcon.LINK
        )

        links_layout.addWidget(self.github_profile_btn)
        links_layout.addWidget(self.github_repo_btn)
        links_layout.addStretch(1)

        info_vbox.addWidget(self.app_name_label)
        info_vbox.addWidget(self.app_desc_label)
        info_vbox.addWidget(self.app_feature_label)
        info_vbox.addSpacing(10)
        info_vbox.addWidget(self.author_title_label)
        info_vbox.addWidget(self.author_name_label)
        info_vbox.addLayout(links_layout)

        ac_layout.addWidget(icon_widget)
        ac_layout.addLayout(info_vbox, 1)

        self.main_layout.addWidget(about_card)
        self.main_layout.addStretch(1)

    def retranslate_ui(self):
        self.title_label.setText(i18n.t("设置与软件信息", "Settings & About"))
        self.subtitle_label.setText(i18n.t("自定义界面外观主题、语言及查看软件与作者信息", "Customize theme, language, and view software & author information"))
        self.theme_card_title.setText(i18n.t("界面主题 (Theme)", "Interface Theme"))
        self.radio_dark.setText(i18n.t("暗黑模式 (Dark Mode)", "Dark Mode"))
        self.radio_light.setText(i18n.t("浅色模式 (Light Mode)", "Light Mode"))
        self.lang_card_title.setText(i18n.t("界面语言 (Language)", "Language Settings"))
        self.app_name_label.setText(i18n.t("Word 文档属性高级修改器 v1.0.0", "Word Metadata Editor v1.0.0"))
        self.app_desc_label.setText(
            i18n.t(
                "基于 PySide6 与 Windows 11 Fluent Design 打造的高颜值 Word 元数据修改软件。",
                "A modern, elegant Word metadata editor built with PySide6 & Windows 11 Fluent Design."
            )
        )
        self.app_feature_label.setText(
            i18n.t(
                "支持功能: 作者修改、总编辑时间自定义、系统时间同步、批量属性更新、一键隐私脱敏。",
                "Features: Author editing, editing time customization, OS timestamp sync, batch processing, and anonymization."
            )
        )
        self.author_title_label.setText(i18n.t("软件作者信息 (Author Information)", "Author Information"))
        self.author_name_label.setText(i18n.t("开发者: treetree-sha", "Developer: treetree-sha"))
        self.github_profile_btn.setText(i18n.t("GitHub 主页 (@treetree-sha)", "GitHub Profile (@treetree-sha)"))
        self.github_repo_btn.setText(i18n.t("GitHub 开源仓库 (WordMetadataEditor)", "GitHub Repository (WordMetadataEditor)"))

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
