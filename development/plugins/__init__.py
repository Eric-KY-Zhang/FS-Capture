from __future__ import annotations

from typing import TYPE_CHECKING

from app.core.models import Exchange

if TYPE_CHECKING:
    from .base import ExchangePlugin


def get_plugin(exchange: Exchange) -> ExchangePlugin:
    """Lazy import to keep startup fast."""
    if exchange is Exchange.A_SHARE:
        from .ashare import AShare

        return AShare()
    if exchange is Exchange.HK:
        from .hk import HKShare

        return HKShare()
    if exchange is Exchange.US:
        from .us import USShare

        return USShare()
    if exchange is Exchange.KR:
        from .kr import KRShare

        return KRShare()
    if exchange is Exchange.TW:
        from .tw import TWShare

        return TWShare()
    raise ValueError(f"No plugin registered for exchange={exchange}")
