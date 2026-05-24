from __future__ import annotations

import threading
import time
from concurrent.futures import ThreadPoolExecutor

from app.core import pdf_renderer


class _Runtime:
    def __init__(self) -> None:
        self.stopped = 0

    def stop(self) -> None:
        self.stopped += 1


class _Page:
    def __init__(self, context: _Context, sleep_seconds: float = 0.0) -> None:
        self.context = context
        self.sleep_seconds = sleep_seconds
        self.closed = 0

    def goto(self, *_args, **_kwargs) -> None:
        return None

    def wait_for_load_state(self, *_args, **_kwargs) -> None:
        return None

    def emulate_media(self, *_args, **_kwargs) -> None:
        return None

    def pdf(self, *, path: str, **_kwargs) -> None:
        if self.sleep_seconds:
            time.sleep(self.sleep_seconds)
        with open(path, "wb") as f:
            f.write(b"%PDF-1.4\nmock\n")

    def close(self) -> None:
        if self.closed:
            return
        self.closed += 1
        self.context.pages_closed += 1


class _Context:
    def __init__(self, browser: _Browser, user_agent: str) -> None:
        self.browser = browser
        self.user_agent = user_agent
        self.closed = False
        self.cookies_cleared = 0
        self.permissions_cleared = 0
        self.pages_closed = 0

    def new_page(self) -> _Page:
        return _Page(self, self.browser.page_sleep_seconds)

    def clear_cookies(self) -> None:
        self.cookies_cleared += 1

    def clear_permissions(self) -> None:
        self.permissions_cleared += 1

    def close(self) -> None:
        if self.closed:
            return
        self.closed = True
        self.browser.contexts_closed += 1


class _Browser:
    def __init__(self, *, page_sleep_seconds: float = 0.0) -> None:
        self.closed = 0
        self.contexts: list[_Context] = []
        self.contexts_closed = 0
        self.page_sleep_seconds = page_sleep_seconds
        self.thread_id = threading.get_ident()

    @property
    def contexts_created(self) -> int:
        return len(self.contexts)

    def new_context(self, **kwargs) -> _Context:
        context = _Context(self, kwargs["user_agent"])
        self.contexts.append(context)
        return context

    def close(self) -> None:
        self.closed += 1


def _reset_renderer() -> None:
    pdf_renderer.shutdown_renderer()


def _patch_playwright(monkeypatch, *, page_sleep_seconds: float = 0.0):
    runtimes: list[_Runtime] = []
    browsers: list[_Browser] = []

    def start() -> _Runtime:
        runtime = _Runtime()
        runtimes.append(runtime)
        return runtime

    def launch(_runtime):
        browser = _Browser(page_sleep_seconds=page_sleep_seconds)
        browsers.append(browser)
        return browser

    monkeypatch.setattr(pdf_renderer, "_start_playwright", start)
    monkeypatch.setattr(pdf_renderer, "_launch_browser", launch)
    return runtimes, browsers


def test_render_url_to_pdf_reuses_context_on_same_thread(tmp_path, monkeypatch) -> None:
    _reset_renderer()
    runtimes, browsers = _patch_playwright(monkeypatch)

    try:
        for i in range(3):
            dest = tmp_path / f"out-{i}.pdf"
            n_bytes = pdf_renderer.render_url_to_pdf("data:text/html,<p>x</p>", dest)
            assert n_bytes > 0
            assert dest.exists()

        assert len(runtimes) == 1
        assert len(browsers) == 1
        assert browsers[0].contexts_created == 1
        assert browsers[0].contexts_closed == 0
        context = browsers[0].contexts[0]
        assert context.cookies_cleared == 3
        assert context.permissions_cleared == 3
        assert context.pages_closed == 3
    finally:
        _reset_renderer()

    assert browsers[0].contexts_closed == 1
    assert browsers[0].closed == 1
    assert runtimes[0].stopped == 1


def test_shutdown_renderer_is_idempotent() -> None:
    _reset_renderer()
    runtime = _Runtime()
    browser = _Browser()
    state = pdf_renderer._RendererState(
        runtime=runtime,
        browser=browser,
        thread_id=threading.get_ident(),
        context=browser.new_context(user_agent="test"),
        context_user_agent="test",
    )
    pdf_renderer._THREAD_RENDERER.state = state
    pdf_renderer._register_state(state)

    pdf_renderer.shutdown_renderer()
    pdf_renderer.shutdown_renderer()

    assert browser.contexts_closed == 1
    assert browser.closed == 1
    assert runtime.stopped == 1
    assert state.closed is True
    assert pdf_renderer._PW_STATES == []


def test_render_url_to_pdf_keeps_playwright_state_thread_local(tmp_path, monkeypatch) -> None:
    _reset_renderer()
    runtimes, browsers = _patch_playwright(monkeypatch, page_sleep_seconds=0.1)

    try:
        with ThreadPoolExecutor(max_workers=2) as executor:
            futures = [
                executor.submit(
                    pdf_renderer.render_url_to_pdf,
                    "data:text/html,<p>x</p>",
                    tmp_path / f"out-{i}.pdf",
                )
                for i in range(2)
            ]

        assert [future.result() for future in futures] == [14, 14]
        assert len(runtimes) == 2
        assert len(browsers) == 2
        assert {browser.thread_id for browser in browsers} != {threading.get_ident()}
        assert all(browser.contexts_created == 1 for browser in browsers)
        assert all(browser.contexts_closed == 0 for browser in browsers)
    finally:
        _reset_renderer()

    assert all(browser.contexts_closed == 1 for browser in browsers)
    assert all(browser.closed == 1 for browser in browsers)
    assert all(runtime.stopped == 1 for runtime in runtimes)
