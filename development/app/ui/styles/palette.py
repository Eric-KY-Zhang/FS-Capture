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

    map_line: str
    route_accent: str
    tile_bg: str
    tile_border: str
    pill_ok: str
    pill_error: str
    pill_pending: str

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
    map_line="#CBD5E1",
    route_accent="#38BDF8",
    tile_bg="#FFFFFF",
    tile_border="#E2E8F0",
    pill_ok="rgba(16, 185, 129, 0.12)",
    pill_error="rgba(239, 68, 68, 0.12)",
    pill_pending="rgba(99, 102, 241, 0.12)",
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
    map_line="#334155",
    route_accent="#38BDF8",
    tile_bg="#111827",
    tile_border="#273449",
    pill_ok="rgba(52, 211, 153, 0.16)",
    pill_error="rgba(248, 113, 113, 0.16)",
    pill_pending="rgba(129, 140, 248, 0.16)",
    shadow="rgba(0, 0, 0, 0.4)",
)


# Per-market accent colors from the UI Refresh design system.
_EXCHANGE_ACCENT = {
    Exchange.UK: "#1E3A8A",
    Exchange.US: "#0891B2",
    Exchange.A_SHARE: "#E11D48",
    Exchange.HK: "#F97316",
    Exchange.TW: "#F59E0B",
    Exchange.KR: "#C026D3",
    Exchange.JP: "#DC2626",
    Exchange.SG: "#A16207",
}


def exchange_accent(exchange: Exchange) -> str:
    return _EXCHANGE_ACCENT[exchange]
