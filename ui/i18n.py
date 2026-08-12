from PySide6.QtCore import QObject, Signal


class TranslationManager(QObject):
    """Global translation manager for switching between Simplified Chinese and English."""
    languageChanged = Signal(str)

    def __init__(self):
        super().__init__()
        self._language = 'zh' # Default language: 'zh' (Simplified Chinese) or 'en' (English)

    @property
    def language(self) -> str:
        return self._language

    def set_language(self, lang: str):
        if lang in ('zh', 'en') and lang != self._language:
            self._language = lang
            self.languageChanged.emit(lang)

    def is_english(self) -> bool:
        return self._language == 'en'

    def is_chinese(self) -> bool:
        return self._language == 'zh'

    def t(self, zh_text: str, en_text: str) -> str:
        """Return translated string based on current language."""
        return en_text if self._language == 'en' else zh_text


# Global i18n instance
i18n = TranslationManager()
