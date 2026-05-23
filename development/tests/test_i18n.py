from __future__ import annotations

import pytest

from app.core.settings import Settings
from app.ui.i18n import LanguageManager


def test_set_language_emits_signal_on_change() -> None:
    mgr = LanguageManager.instance()
    mgr.set_language("zh")
    received: list[str] = []
    mgr.language_changed.connect(received.append)

    mgr.set_language("en")

    assert received == ["en"]


def test_set_language_idempotent_when_unchanged() -> None:
    mgr = LanguageManager.instance()
    mgr.set_language("zh")
    received: list[str] = []
    mgr.language_changed.connect(received.append)

    mgr.set_language("zh")

    assert received == []


def test_set_language_rejects_invalid_code() -> None:
    with pytest.raises(ValueError, match="Unsupported language"):
        LanguageManager.instance().set_language("ja")


def test_settings_migrates_old_zh_cn_language() -> None:
    settings = Settings.model_validate({"ui": {"language": "zh_CN"}})

    assert settings.ui.language == "zh"


def test_settings_migrates_old_en_us_language() -> None:
    settings = Settings.model_validate({"ui": {"language": "en_US"}})

    assert settings.ui.language == "en"
