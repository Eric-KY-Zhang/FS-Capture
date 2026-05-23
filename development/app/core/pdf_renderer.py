from __future__ import annotations

import threading
from pathlib import Path
from typing import Any

from loguru import logger

from .settings import load_settings

_PW_LOCK = threading.Lock()
_PW_RUNTIME: Any | None = None
_PW_BROWSER: Any | None = None
_PW_CONTEXT_SEMAPHORE = threading.Semaphore(4)


def _user_agent(source: str) -> str:
    if source == "sec":
        return load_settings().sec.user_agent
    return (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/126.0.0.0 Safari/537.36"
    )


def _launch_browser(playwright):
    try:
        return playwright.chromium.launch(headless=True)
    except Exception as exc:
        logger.warning(f"bundled Playwright Chromium launch failed, trying Microsoft Edge: {exc}")
        return playwright.chromium.launch(channel="msedge", headless=True)


def _start_playwright():
    from playwright.sync_api import sync_playwright

    return sync_playwright().start()


def _ensure_browser():
    global _PW_BROWSER, _PW_RUNTIME
    with _PW_LOCK:
        if _PW_BROWSER is not None:
            return _PW_BROWSER
        _PW_RUNTIME = _start_playwright()
        _PW_BROWSER = _launch_browser(_PW_RUNTIME)
        return _PW_BROWSER


def shutdown_renderer() -> None:
    """Close the process-wide Playwright browser/runtime."""
    global _PW_BROWSER, _PW_RUNTIME
    with _PW_LOCK:
        if _PW_BROWSER is not None:
            try:
                _PW_BROWSER.close()
            except Exception as exc:
                logger.warning(f"playwright browser close failed: {exc}")
            _PW_BROWSER = None
        if _PW_RUNTIME is not None:
            try:
                _PW_RUNTIME.stop()
            except Exception as exc:
                logger.warning(f"playwright runtime stop failed: {exc}")
            _PW_RUNTIME = None


def _render_once(url: str, tmp_dest: Path, *, source: str, timeout_ms: int) -> None:
    browser = _ensure_browser()
    context = None
    with _PW_CONTEXT_SEMAPHORE:
        try:
            context = browser.new_context(
                user_agent=_user_agent(source),
                locale="en-US",
                extra_http_headers={"Accept-Language": "en-US,en;q=0.9"},
            )
            page = context.new_page()
            page.goto(url, wait_until="domcontentloaded", timeout=timeout_ms)
            try:
                page.wait_for_load_state("networkidle", timeout=15_000)
            except Exception:
                pass
            page.emulate_media(media="screen")
            page.pdf(
                path=str(tmp_dest),
                format="Letter",
                print_background=True,
                margin={
                    "top": "0.35in",
                    "right": "0.35in",
                    "bottom": "0.35in",
                    "left": "0.35in",
                },
                prefer_css_page_size=True,
            )
        finally:
            if context is not None:
                context.close()


def render_url_to_pdf(
    url: str,
    dest: Path,
    *,
    source: str = "generic",
    timeout_ms: int = 120_000,
) -> int:
    """Render a remote HTML document to PDF using Playwright.

    Returns bytes written. The Chromium browser is a process-wide singleton;
    each render gets an isolated browser context and then closes that context.
    """
    dest.parent.mkdir(parents=True, exist_ok=True)
    tmp_dest = dest.with_name(f"{dest.name}.part")
    tmp_dest.unlink(missing_ok=True)

    try:
        _render_once(url, tmp_dest, source=source, timeout_ms=timeout_ms)
    except Exception:
        logger.warning(f"playwright render failed for {url}; restarting browser and retrying")
        shutdown_renderer()
        tmp_dest.unlink(missing_ok=True)
        _render_once(url, tmp_dest, source=source, timeout_ms=timeout_ms)

    tmp_dest.replace(dest)
    n = dest.stat().st_size
    logger.debug(f"rendered {url} -> {dest} ({n} bytes)")
    return n
