from __future__ import annotations

import os
import time
from pathlib import Path

import httpx
import pytest

from app.core import http as http_module
from app.core.output_paths import cleanup_stale_parts


class _Limiter:
    def acquire_blocking(self) -> None:
        return None


class _StreamResponse:
    def __init__(self, status_code: int, chunks: list[bytes | Exception], url: str) -> None:
        self.status_code = status_code
        self._chunks = chunks
        self._request = httpx.Request("GET", url)
        self._response = httpx.Response(status_code, request=self._request)
        self.chunk_sizes: list[int] = []

    def __enter__(self):
        return self

    def __exit__(self, *_args):
        return False

    def raise_for_status(self) -> None:
        if self.status_code >= 400:
            raise httpx.HTTPStatusError(
                f"HTTP {self.status_code}",
                request=self._request,
                response=self._response,
            )

    def iter_bytes(self, chunk_size: int):
        self.chunk_sizes.append(chunk_size)
        for chunk in self._chunks:
            if isinstance(chunk, Exception):
                raise chunk
            yield chunk


class _Client:
    def __init__(self, responses: list[_StreamResponse]) -> None:
        self.responses = responses
        self.calls: list[dict] = []

    def stream(self, method: str, url: str, **kwargs):
        self.calls.append({"method": method, "url": url, "kwargs": kwargs})
        return self.responses.pop(0)


def _disable_rate_limit(monkeypatch) -> None:
    monkeypatch.setattr(http_module, "limiter", lambda *_args, **_kwargs: _Limiter())


def test_stream_to_file_resumes_from_existing_part(tmp_path, monkeypatch) -> None:
    _disable_rate_limit(monkeypatch)
    dest = tmp_path / "report.pdf"
    dest.with_name("report.pdf.part").write_bytes(b"AAA")
    client = _Client([_StreamResponse(206, [b"BBB"], "https://example.com/report.pdf")])

    n_bytes = http_module.stream_to_file.__wrapped__(
        client, "https://example.com/report.pdf", dest, source="test", rate=999.0
    )

    assert n_bytes == 6
    assert dest.read_bytes() == b"AAABBB"
    assert client.calls[0]["kwargs"]["headers"] == {"Range": "bytes=3-"}


def test_stream_to_file_restarts_when_server_ignores_range(tmp_path, monkeypatch) -> None:
    _disable_rate_limit(monkeypatch)
    dest = tmp_path / "report.pdf"
    dest.with_name("report.pdf.part").write_bytes(b"AAA")
    client = _Client([_StreamResponse(200, [b"FULL"], "https://example.com/report.pdf")])

    n_bytes = http_module.stream_to_file.__wrapped__(
        client, "https://example.com/report.pdf", dest, source="test", rate=999.0
    )

    assert n_bytes == 4
    assert dest.read_bytes() == b"FULL"


def test_stream_to_file_keeps_part_on_failure_for_resume(tmp_path, monkeypatch) -> None:
    _disable_rate_limit(monkeypatch)
    dest = tmp_path / "report.pdf"
    client = _Client(
        [
            _StreamResponse(
                200,
                [b"AA", httpx.ReadError("network reset")],
                "https://example.com/report.pdf",
            )
        ]
    )

    with pytest.raises(httpx.ReadError):
        http_module.stream_to_file.__wrapped__(
            client, "https://example.com/report.pdf", dest, source="test", rate=999.0
        )

    part = dest.with_name("report.pdf.part")
    assert part.exists()
    assert part.read_bytes() == b"AA"


def test_stream_to_file_drops_unsatisfiable_part_and_retries(tmp_path, monkeypatch) -> None:
    _disable_rate_limit(monkeypatch)
    dest = tmp_path / "report.pdf"
    dest.with_name("report.pdf.part").write_bytes(b"STALE")
    url = "https://example.com/report.pdf"
    client = _Client(
        [
            _StreamResponse(416, [], url),
            _StreamResponse(200, [b"FULL"], url),
        ]
    )

    n_bytes = http_module.stream_to_file(client, url, dest, source="test", rate=999.0)

    assert n_bytes == 4
    assert dest.read_bytes() == b"FULL"
    assert client.calls[0]["kwargs"]["headers"] == {"Range": "bytes=5-"}
    assert client.calls[1]["kwargs"]["headers"] == {}


def test_cleanup_stale_parts_drops_old_orphans(tmp_path: Path) -> None:
    fresh = tmp_path / "fresh.pdf.part"
    old = tmp_path / "nested" / "old.pdf.part"
    fresh.write_bytes(b"fresh")
    old.parent.mkdir()
    old.write_bytes(b"old")
    old_mtime = time.time() - 2 * 24 * 3600
    os.utime(old, (old_mtime, old_mtime))

    removed = cleanup_stale_parts(tmp_path, max_age_days=1)

    assert removed == 1
    assert fresh.exists()
    assert not old.exists()


def test_stream_to_file_default_chunk_size_is_256kb(tmp_path, monkeypatch) -> None:
    _disable_rate_limit(monkeypatch)
    dest = tmp_path / "report.pdf"
    response = _StreamResponse(200, [b"FULL"], "https://example.com/report.pdf")
    client = _Client([response])

    http_module.stream_to_file.__wrapped__(
        client, "https://example.com/report.pdf", dest, source="test", rate=999.0
    )

    assert response.chunk_sizes == [256 * 1024]
