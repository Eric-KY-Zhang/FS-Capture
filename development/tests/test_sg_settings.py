from __future__ import annotations

from app.core.settings import Settings


def test_sgxnet_rate_limit_default() -> None:
    assert Settings().rate_limits.sgxnet == 1.5


def test_sgxnet_rate_limit_can_be_configured() -> None:
    settings = Settings.model_validate({"rate_limits": {"sgxnet": 2.0}})

    assert settings.rate_limits.sgxnet == 2.0
