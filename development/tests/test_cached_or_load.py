from __future__ import annotations

import threading
import time
from concurrent.futures import ThreadPoolExecutor

import pytest

from app.core import cache as cache_module


class _FakeCache:
    def __init__(self) -> None:
        self._lock = threading.Lock()
        self._store: dict[str, object] = {}
        self.expires: dict[str, float | None] = {}

    def get(self, key: str):
        with self._lock:
            return self._store.get(key)

    def set(self, key: str, value, expire=None) -> None:
        with self._lock:
            self._store[key] = value
            self.expires[key] = expire


def test_cached_or_load_single_flight_under_concurrency(monkeypatch) -> None:
    fake_cache = _FakeCache()
    monkeypatch.setattr(cache_module, "get_cache", lambda: fake_cache)
    calls = 0
    calls_lock = threading.Lock()

    def loader() -> str:
        nonlocal calls
        time.sleep(0.2)
        with calls_lock:
            calls += 1
        return "value"

    with ThreadPoolExecutor(max_workers=4) as executor:
        futures = [
            executor.submit(
                cache_module.cached_or_load, "test:single-flight", loader, expire=None
            )
            for _ in range(4)
        ]

    assert [future.result() for future in futures] == ["value"] * 4
    assert calls == 1


def test_cached_or_load_returns_cached_immediately_on_hit(monkeypatch) -> None:
    fake_cache = _FakeCache()
    fake_cache.set("test:hit", "X")
    monkeypatch.setattr(cache_module, "get_cache", lambda: fake_cache)

    value = cache_module.cached_or_load(
        "test:hit",
        loader=lambda: pytest.fail("loader should not run"),
        expire=None,
    )

    assert value == "X"


def test_cached_or_load_does_not_cache_none_result(monkeypatch) -> None:
    fake_cache = _FakeCache()
    monkeypatch.setattr(cache_module, "get_cache", lambda: fake_cache)

    value = cache_module.cached_or_load("test:none", loader=lambda: None, expire=None)

    assert value is None
    assert fake_cache.get("test:none") is None


def test_cached_or_load_passes_expire_to_cache(monkeypatch) -> None:
    fake_cache = _FakeCache()
    monkeypatch.setattr(cache_module, "get_cache", lambda: fake_cache)

    value = cache_module.cached_or_load("test:ttl", loader=lambda: "value", expire=60.0)

    assert value == "value"
    assert fake_cache.expires["test:ttl"] == 60.0
