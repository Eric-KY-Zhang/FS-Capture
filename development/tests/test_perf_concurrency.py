from __future__ import annotations

from app.core.settings import Settings


def test_default_max_workers_is_six() -> None:
    assert Settings().concurrency.max_workers == 6
