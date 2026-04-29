from __future__ import annotations

from functools import lru_cache

import diskcache

from .settings import load_settings


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
