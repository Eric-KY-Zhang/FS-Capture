from __future__ import annotations

from dataclasses import dataclass

from app.core.models import Exchange


@dataclass(frozen=True)
class Palette:
    bg: str  # window background
    surface: str  # card background
    surface_alt: str  # alt card / hover
    border: str  # subtle separators
    border_strong: str
    text: str
    text_muted: str
    text_subtle: str

    primary: str
    primary_hover: str
    primary_pressed: str
    primary_text: str

    success: str
    warning: str
    danger: str

    shadow: str  # rgba string used in box shadow effects


light_palette = Palette(
    bg="#F4F5F8",
    surface="#FFFFFF",
    surface_alt="#F8F9FB",
    border="#E6E8EC",
    border_strong="#CDD1D9",
    text="#0F172A",
    text_muted="#475569",
    text_subtle="#94A3B8",
    primary="#6366F1",
    primary_hover="#5457E5",
    primary_pressed="#4338CA",
    primary_text="#FFFFFF",
    success="#10B981",
    warning="#F59E0B",
    danger="#EF4444",
    shadow="rgba(15, 23, 42, 0.08)",
)


dark_palette = Palette(
    bg="#0B1220",
    surface="#111827",
    surface_alt="#1F2937",
    border="#1F2A3C",
    border_strong="#334155",
    text="#F1F5F9",
    text_muted="#CBD5E1",
    text_subtle="#64748B",
    primary="#818CF8",
    primary_hover="#A5B4FC",
    primary_pressed="#6366F1",
    primary_text="#0F172A",
    success="#34D399",
    warning="#FBBF24",
    danger="#F87171",
    shadow="rgba(0, 0, 0, 0.4)",
)


# Per-exchange accent colors. Chinese convention: red = up = A-share home market.
_EXCHANGE_ACCENT = {
    Exchange.A_SHARE: "#E11D48",  # rose-600 — A股
    Exchange.HK: "#D97706",  # amber-600 — 港股
    Exchange.US: "#2563EB",  # blue-600 — 美股
    Exchange.KR: "#7C3AED",  # violet-600 — 韩股
    Exchange.TW: "#0EA5E9",  # sky-500 — 台股
    Exchange.JP: "#059669",  # emerald-600 — 日股
}


def exchange_accent(exchange: Exchange) -> str:
    return _EXCHANGE_ACCENT[exchange]
