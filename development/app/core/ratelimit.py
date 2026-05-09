from __future__ import annotations

import asyncio
import threading
import time


class TokenBucket:
    """Async + thread-safe token bucket for rate limiting outgoing requests."""

    def __init__(self, rate_per_sec: float, burst: int | None = None) -> None:
        self.rate = float(rate_per_sec)
        self.capacity = float(burst if burst is not None else max(1, int(rate_per_sec)))
        self._tokens = self.capacity
        self._last = time.monotonic()
        self._lock = threading.Lock()
        self._async_lock = asyncio.Lock()

    def _refill(self) -> None:
        now = time.monotonic()
        elapsed = now - self._last
        self._tokens = min(self.capacity, self._tokens + elapsed * self.rate)
        self._last = now

    def acquire_blocking(self) -> None:
        """Block (sleep) until a token is available. Use from sync code."""
        while True:
            with self._lock:
                self._refill()
                if self._tokens >= 1.0:
                    self._tokens -= 1.0
                    return
                deficit = 1.0 - self._tokens
                wait = deficit / self.rate if self.rate > 0 else 0.1
            time.sleep(wait)

    async def acquire(self) -> None:
        """Async acquire."""
        while True:
            async with self._async_lock:
                with self._lock:
                    self._refill()
                    if self._tokens >= 1.0:
                        self._tokens -= 1.0
                        return
                    deficit = 1.0 - self._tokens
                    wait = deficit / self.rate if self.rate > 0 else 0.1
            await asyncio.sleep(wait)


class RateLimiterRegistry:
    """One bucket per data source, lazily created."""

    def __init__(self) -> None:
        self._buckets: dict[str, TokenBucket] = {}
        self._lock = threading.Lock()

    def get(self, source: str, rate: float) -> TokenBucket:
        with self._lock:
            if source not in self._buckets:
                self._buckets[source] = TokenBucket(rate)
            return self._buckets[source]


_registry = RateLimiterRegistry()


def limiter(source: str, rate: float) -> TokenBucket:
    return _registry.get(source, rate)
