from __future__ import annotations

from types import SimpleNamespace
from unittest.mock import Mock

import pytest

from app.core import settings as settings_module
from app.core.settings import Settings, invalidate_dart_client_cache


def test_dart_api_key_is_canonical_over_opendart_alias() -> None:
    settings = Settings.model_validate(
        {
            "dart": {"api_key": "canonical"},
            "opendart": {"api_key": "alias"},
        }
    )

    assert settings.dart.api_key == "canonical"


def test_opendart_api_key_alias_is_accepted() -> None:
    settings = Settings.model_validate({"opendart": {"api_key": "alias"}})

    assert settings.dart.api_key == "alias"


def test_blank_dart_api_key_falls_back_to_opendart_alias() -> None:
    settings = Settings.model_validate(
        {
            "dart": {"api_key": "  "},
            "opendart": {"api_key": "alias"},
        }
    )

    assert settings.dart.api_key == "alias"


def test_opendart_api_key_env_fallback_when_config_has_no_key(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    monkeypatch.setenv("OPENDART_API_KEY", "env-key")
    settings = Settings.model_validate({})

    assert settings.dart.api_key == "env-key"


def test_saved_settings_use_dart_section_only() -> None:
    settings = Settings.model_validate({"opendart": {"api_key": "alias"}})

    dumped = settings.model_dump()
    text = settings_module.tomli_w.dumps(dumped)

    assert "dart" in dumped
    assert "opendart" not in dumped
    assert "[dart]" in text
    assert 'api_key = "alias"' in text
    assert "[opendart]" not in text


def test_load_settings_creates_and_returns_default_on_missing(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    cfg_file = SimpleNamespace(exists=lambda: False)
    save_mock = Mock()
    monkeypatch.setattr(settings_module, "config_path", lambda: cfg_file)
    monkeypatch.setattr(settings_module, "save_settings", save_mock)

    settings = settings_module.load_settings()

    assert isinstance(settings, Settings)
    assert settings == Settings()
    save_mock.assert_called_once_with(settings)


def test_dart_client_cache_invalidation_is_optional(monkeypatch: pytest.MonkeyPatch) -> None:
    def raise_import_error(_name: str):
        raise ImportError

    monkeypatch.setattr(settings_module.importlib, "import_module", raise_import_error)

    assert not invalidate_dart_client_cache()


def test_dart_client_cache_invalidation_calls_cache_clear(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    cached_factory = SimpleNamespace(cache_clear=Mock())
    module = SimpleNamespace(_dart=cached_factory)
    monkeypatch.setattr(settings_module.importlib, "import_module", lambda _name: module)

    assert invalidate_dart_client_cache()
    cached_factory.cache_clear.assert_called_once_with()


def test_dart_client_cache_invalidation_prefers_reset_hook(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    reset = Mock()
    cached_factory = SimpleNamespace(cache_clear=Mock())
    module = SimpleNamespace(reset_dart_client=reset, _dart_for_key=cached_factory)
    monkeypatch.setattr(settings_module.importlib, "import_module", lambda _name: module)

    assert invalidate_dart_client_cache()
    reset.assert_called_once_with()
    cached_factory.cache_clear.assert_not_called()


def test_kr_dart_client_cache_is_key_aware(monkeypatch: pytest.MonkeyPatch) -> None:
    from plugins.kr import name_resolver

    one = Settings.model_validate({"dart": {"api_key": "one"}})
    two = Settings.model_validate({"dart": {"api_key": "two"}})
    calls: list[str] = []

    def fake_factory(key: str) -> str:
        calls.append(key)
        return f"client:{key}"

    name_resolver.reset_dart_client()
    monkeypatch.setattr(name_resolver, "load_settings", lambda: one)
    monkeypatch.setattr(name_resolver, "_dart_for_key", fake_factory)
    assert name_resolver._dart() == "client:one"

    monkeypatch.setattr(name_resolver, "load_settings", lambda: two)
    assert name_resolver._dart() == "client:two"
    assert calls == ["one", "two"]


def test_save_invalidates_dart_cache_when_key_changes(monkeypatch: pytest.MonkeyPatch) -> None:
    try:
        from app.ui.settings_dialog import SettingsDialog
    except ImportError as exc:
        pytest.skip(f"PySide6 unavailable: {exc}")

    dialog = SimpleNamespace(
        settings=Settings.model_validate({"dart": {"api_key": "old-key"}}),
        dart_key=SimpleNamespace(text=lambda: "new-key"),
        workers=SimpleNamespace(value=lambda: 4),
        theme=SimpleNamespace(currentData=lambda: "light"),
        language=SimpleNamespace(currentData=lambda: "zh"),
        sec_ua=SimpleNamespace(text=lambda: "Filings Atlas test"),
        accept=Mock(),
    )
    save_mock = Mock()
    invalidate_mock = Mock()
    monkeypatch.setattr("app.ui.settings_dialog.save_settings", save_mock)
    monkeypatch.setattr("app.ui.settings_dialog.invalidate_dart_client_cache", invalidate_mock)

    SettingsDialog._save(dialog)

    save_mock.assert_called_once_with(dialog.settings)
    invalidate_mock.assert_called_once_with()
    dialog.accept.assert_called_once_with()
