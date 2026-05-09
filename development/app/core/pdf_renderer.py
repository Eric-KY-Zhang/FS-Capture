from __future__ import annotations

import threading
from pathlib import Path

from loguru import logger

from .settings import load_settings

_RENDER_SEMAPHORE = threading.Semaphore(2)


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


def render_url_to_pdf(
    url: str,
    dest: Path,
    *,
    source: str = "generic",
    timeout_ms: int = 120_000,
) -> int:
    """Render a remote HTML document to PDF using Playwright.

    Returns bytes written. The function serializes browser creation enough to
    avoid unstable high-concurrency Chromium launches from worker threads.
    """
    from playwright.sync_api import sync_playwright

    dest.parent.mkdir(parents=True, exist_ok=True)
    tmp_dest = dest.with_name(f"{dest.name}.part")
    tmp_dest.unlink(missing_ok=True)

    with _RENDER_SEMAPHORE:
        with sync_playwright() as p:
            browser = _launch_browser(p)
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
                context.close()
            finally:
                browser.close()

    tmp_dest.replace(dest)
    n = dest.stat().st_size
    logger.debug(f"rendered {url} -> {dest} ({n} bytes)")
    return n
