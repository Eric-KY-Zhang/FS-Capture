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
    def __init__(self, sleep_seconds: float = 0.0) -> None:
        self.sleep_seconds = sleep_seconds

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


class _Context:
    def __init__(self, browser: _Browser) -> None:
        self.browser = browser
        self.closed = False

    def new_page(self) -> _Page:
        return _Page(self.browser.page_sleep_seconds)

    def close(self) -> None:
        if self.closed:
            return
        self.closed = True
        self.browser.contexts_closed += 1
        with self.browser.active_lock:
            self.browser.active_contexts -= 1


class _Browser:
    def __init__(self, *, page_sleep_seconds: float = 0.0) -> None:
        self.closed = 0
        self.contexts_created = 0
        self.contexts_closed = 0
        self.active_contexts = 0
        self.max_active_contexts = 0
        self.active_lock = threading.Lock()
        self.page_sleep_seconds = page_sleep_seconds

    def new_context(self, **_kwargs) -> _Context:
        with self.active_lock:
            self.contexts_created += 1
            self.active_contexts += 1
            self.max_active_contexts = max(self.max_active_contexts, self.active_contexts)
        return _Context(self)

    def close(self) -> None:
        self.closed += 1


def _reset_renderer() -> None:
    pdf_renderer.shutdown_renderer()


def test_render_url_to_pdf_reuses_browser_across_calls(tmp_path, monkeypatch) -> None:
    _reset_renderer()
    runtime = _Runtime()
    browser = _Browser()
    launches = 0

    def launch(_runtime):
        nonlocal launches
        launches += 1
        return browser

    monkeypatch.setattr(pdf_renderer, "_start_playwright", lambda: runtime)
    monkeypatch.setattr(pdf_renderer, "_launch_browser", launch)

    try:
        for i in range(3):
            dest = tmp_path / f"out-{i}.pdf"
            n_bytes = pdf_renderer.render_url_to_pdf("data:text/html,<p>x</p>", dest)
            assert n_bytes > 0
            assert dest.exists()

        assert launches == 1
        assert browser.contexts_created == 3
        assert browser.contexts_closed == 3
    finally:
        _reset_renderer()


def test_shutdown_renderer_is_idempotent() -> None:
    _reset_renderer()
    runtime = _Runtime()
    browser = _Browser()
    pdf_renderer._PW_RUNTIME = runtime
    pdf_renderer._PW_BROWSER = browser

    pdf_renderer.shutdown_renderer()
    pdf_renderer.shutdown_renderer()

    assert browser.closed == 1
    assert runtime.stopped == 1
    assert pdf_renderer._PW_BROWSER is None
    assert pdf_renderer._PW_RUNTIME is None


def test_render_url_to_pdf_serializes_contexts_via_semaphore(tmp_path, monkeypatch) -> None:
    _reset_renderer()
    runtime = _Runtime()
    browser = _Browser(page_sleep_seconds=0.1)

    monkeypatch.setattr(pdf_renderer, "_start_playwright", lambda: runtime)
    monkeypatch.setattr(pdf_renderer, "_launch_browser", lambda _runtime: browser)
    monkeypatch.setattr(pdf_renderer, "_PW_CONTEXT_SEMAPHORE", threading.Semaphore(1))

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
        assert browser.contexts_created == 2
        assert browser.contexts_closed == 2
        assert browser.max_active_contexts == 1
    finally:
        _reset_renderer()
