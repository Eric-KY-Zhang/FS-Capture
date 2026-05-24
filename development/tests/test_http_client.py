from __future__ import annotations

from app.core import http


def test_default_client_uses_tuned_connection_limits(monkeypatch) -> None:
    captured: dict[str, object] = {}

    class _Client:
        def __init__(self, **kwargs) -> None:
            captured.update(kwargs)

    monkeypatch.setattr(http.httpx, "Client", _Client)

    http.default_client()

    limits = captured["limits"]
    assert isinstance(limits, http.httpx.Limits)
    assert limits.max_connections == 100
    assert limits.max_keepalive_connections == 40
