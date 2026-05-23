from __future__ import annotations

from PySide6.QtCore import QObject, Signal

_VALID_LANGS = {"zh", "en"}


class LanguageManager(QObject):
    """Process-wide singleton for runtime UI language switching."""

    language_changed = Signal(str)

    _instance: LanguageManager | None = None

    def __init__(self) -> None:
        super().__init__()
        self._current_language = "zh"
        assert LanguageManager._instance is None, "LanguageManager is a singleton"

    @classmethod
    def instance(cls) -> LanguageManager:
        if cls._instance is None:
            cls._instance = cls()
        return cls._instance

    @property
    def current_language(self) -> str:
        return self._current_language

    def set_language(self, code: str) -> None:
        if code not in _VALID_LANGS:
            raise ValueError(f"Unsupported language: {code!r}")
        if code == self._current_language:
            return
        self._current_language = code
        self.language_changed.emit(code)
