from __future__ import annotations

from pathlib import Path
from string import Template

from .palette import Palette


def load_qss(palette: Palette) -> str:
    """Load app.qss and substitute color tokens."""
    qss_file = Path(__file__).with_name("app.qss")
    raw = qss_file.read_text(encoding="utf-8")
    return Template(raw).safe_substitute(
        bg=palette.bg,
        surface=palette.surface,
        surface_alt=palette.surface_alt,
        border=palette.border,
        border_strong=palette.border_strong,
        text=palette.text,
        text_muted=palette.text_muted,
        text_subtle=palette.text_subtle,
        primary=palette.primary,
        primary_hover=palette.primary_hover,
        primary_pressed=palette.primary_pressed,
        primary_text=palette.primary_text,
        success=palette.success,
        warning=palette.warning,
        danger=palette.danger,
        map_line=palette.map_line,
        route_accent=palette.route_accent,
        tile_bg=palette.tile_bg,
        tile_border=palette.tile_border,
        pill_ok=palette.pill_ok,
        pill_error=palette.pill_error,
        pill_pending=palette.pill_pending,
    )
