from __future__ import annotations

import threading
from dataclasses import dataclass
from pathlib import Path
from typing import Any

from loguru import logger

from .settings import load_settings

_PW_LOCK = threading.Lock()
_THREAD_RENDERER = threading.local()
_PW_STATES: list[_RendererState] = []


@dataclass
class _RendererState:
    runtime: Any
    browser: Any
    thread_id: int
    context: Any | None = None
    context_user_agent: str | None = None
    closed: bool = False


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


def _register_state(state: _RendererState) -> None:
    with _PW_LOCK:
        _PW_STATES.append(state)


def _unregister_state(state: _RendererState) -> None:
    with _PW_LOCK:
        try:
            _PW_STATES.remove(state)
        except ValueError:
            pass


def _ensure_state() -> _RendererState:
    state: _RendererState | None = getattr(_THREAD_RENDERER, "state", None)
    if state is not None and not state.closed:
        return state

    runtime = _start_playwright()
    browser = _launch_browser(runtime)
    state = _RendererState(
        runtime=runtime,
        browser=browser,
        thread_id=threading.get_ident(),
    )
    _THREAD_RENDERER.state = state
    _register_state(state)
    return state


def _close_context(state: _RendererState) -> None:
    if state.context is None:
        return
    try:
        state.context.close()
    except Exception as exc:
        logger.warning(f"playwright context close failed: {exc}")
    state.context = None
    state.context_user_agent = None


def _close_state(state: _RendererState) -> None:
    if state.closed:
        return
    _close_context(state)
    try:
        state.browser.close()
    except Exception as exc:
        logger.warning(f"playwright browser close failed: {exc}")
    try:
        state.runtime.stop()
    except Exception as exc:
        logger.warning(f"playwright runtime stop failed: {exc}")
    state.closed = True
    _unregister_state(state)
    if getattr(_THREAD_RENDERER, "state", None) is state:
        delattr(_THREAD_RENDERER, "state")


def _shutdown_current_thread_renderer() -> None:
    state: _RendererState | None = getattr(_THREAD_RENDERER, "state", None)
    if state is not None:
        _close_state(state)


def shutdown_renderer() -> None:
    """Best-effort close all known Playwright browser/runtime states."""
    current_state: _RendererState | None = getattr(_THREAD_RENDERER, "state", None)
    if current_state is not None:
        _close_state(current_state)

    with _PW_LOCK:
        states = list(_PW_STATES)

    for state in states:
        _close_state(state)


def _context_for_source(source: str):
    state = _ensure_state()
    user_agent = _user_agent(source)
    if state.context is None or state.context_user_agent != user_agent:
        _close_context(state)
        state.context = state.browser.new_context(
            user_agent=user_agent,
            locale="en-US",
            extra_http_headers={"Accept-Language": "en-US,en;q=0.9"},
        )
        state.context_user_agent = user_agent
    return state.context


def _clear_context_state(context: Any) -> bool:
    ok = True
    for method_name in ("clear_cookies", "clear_permissions"):
        method = getattr(context, method_name, None)
        if method is None:
            continue
        try:
            method()
        except Exception as exc:
            logger.warning(f"playwright context {method_name} failed: {exc}")
            ok = False
    return ok


def _render_once(url: str, tmp_dest: Path, *, source: str, timeout_ms: int) -> None:
    context = _context_for_source(source)
    page = None
    try:
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
        if page is not None:
            try:
                page.close()
            except Exception as exc:
                logger.warning(f"playwright page close failed: {exc}")
        if not _clear_context_state(context):
            state = _ensure_state()
            _close_context(state)


def render_url_to_pdf(
    url: str,
    dest: Path,
    *,
    source: str = "generic",
    timeout_ms: int = 120_000,
) -> int:
    """Render a remote HTML document to PDF using Playwright.

    Returns bytes written. Playwright's sync API is thread-affine, so each
    worker thread owns and reuses its browser/context while clearing cookies
    and permissions after every render.
    """
    dest.parent.mkdir(parents=True, exist_ok=True)
    tmp_dest = dest.with_name(f"{dest.name}.part")
    tmp_dest.unlink(missing_ok=True)

    try:
        _render_once(url, tmp_dest, source=source, timeout_ms=timeout_ms)
    except Exception:
        logger.warning(f"playwright render failed for {url}; restarting renderer and retrying")
        _shutdown_current_thread_renderer()
        tmp_dest.unlink(missing_ok=True)
        _render_once(url, tmp_dest, source=source, timeout_ms=timeout_ms)

    tmp_dest.replace(dest)
    n = dest.stat().st_size
    logger.debug(f"rendered {url} -> {dest} ({n} bytes)")
    return n
