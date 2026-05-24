from __future__ import annotations

from app.core.models import Exchange
from app.ui import main_view


def test_prewarm_name_resolver_calls_registered_loader(monkeypatch) -> None:
    calls: list[str] = []
    monkeypatch.setattr(main_view, "_PREWARM_LOADERS", {Exchange.US: lambda: calls.append("us")})

    main_view._prewarm_name_resolver(Exchange.US)

    assert calls == ["us"]


def test_prewarm_name_resolver_ignores_unregistered_exchange(monkeypatch) -> None:
    calls: list[str] = []
    monkeypatch.setattr(main_view, "_PREWARM_LOADERS", {Exchange.US: lambda: calls.append("us")})

    main_view._prewarm_name_resolver(Exchange.SG)

    assert calls == []


def test_prewarm_name_resolver_swallows_loader_errors(monkeypatch) -> None:
    def fail() -> None:
        raise RuntimeError("source unavailable")

    monkeypatch.setattr(main_view, "_PREWARM_LOADERS", {Exchange.US: fail})

    main_view._prewarm_name_resolver(Exchange.US)
