from __future__ import annotations

import threading
from collections.abc import Callable
from functools import lru_cache
from typing import TypeVar

import diskcache

from .settings import load_settings

T = TypeVar("T")
_PER_KEY_LOCKS: dict[str, threading.Lock] = {}
_LOCKS_MUTEX = threading.Lock()


@lru_cache(maxsize=1)
def get_cache() -> diskcache.Cache:
    s = load_settings()
    return diskcache.Cache(str(s.cache_path()))


def close_cache() -> None:
    """Close the process-wide diskcache handle on application shutdown."""
    if get_cache.cache_info().currsize == 0:
        return
    get_cache().close()
    get_cache.cache_clear()


def _get_key_lock(key: str) -> threading.Lock:
    with _LOCKS_MUTEX:
        lock = _PER_KEY_LOCKS.get(key)
        if lock is None:
            lock = threading.Lock()
            _PER_KEY_LOCKS[key] = lock
        return lock


def cached_or_load(key: str, loader: Callable[[], T], *, expire: float | None) -> T:
    """Return a cached value, with single-flight loading for cache misses."""
    cache = get_cache()
    cached = cache.get(key)
    if cached is not None:
        return cached

    lock = _get_key_lock(key)
    with lock:
        cached = cache.get(key)
        if cached is not None:
            return cached
        value = loader()
        if value is not None:
            if expire is None:
                cache.set(key, value)
            else:
                cache.set(key, value, expire=expire)
        return value
