from __future__ import annotations

from app.core.ratelimit import RateLimiterRegistry


def test_token_bucket_rate_hot_update() -> None:
    registry = RateLimiterRegistry()
    bucket = registry.get("cninfo", 2.0)

    updated = registry.get("cninfo", 10.0)

    assert updated is bucket
    assert updated.rate == 10.0
    assert updated.capacity == 10.0


def test_token_bucket_rate_unchanged_returns_same_instance() -> None:
    registry = RateLimiterRegistry()

    first = registry.get("cninfo", 3.0)
    second = registry.get("cninfo", 3.0)

    assert second is first
