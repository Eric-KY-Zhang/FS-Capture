from __future__ import annotations

import ssl
from functools import lru_cache
from pathlib import Path
from typing import Any

import certifi
import httpx
from loguru import logger
from tenacity import retry, stop_after_attempt, wait_exponential

from .ratelimit import limiter
from .settings import load_settings

_DEFAULT_HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/126.0.0.0 Safari/537.36"
    ),
    "Accept": "*/*",
    "Accept-Language": "zh-CN,zh;q=0.9,en;q=0.8",
}


def default_client(
    *,
    source: str = "generic",
    timeout: float | httpx.Timeout = 30.0,
) -> httpx.Client:
    """Sync httpx.Client with sensible defaults + cookie jar.

    Caller is responsible for `with client:` lifecycle.
    """
    headers = dict(_DEFAULT_HEADERS)
    if source == "sec":
        s = load_settings()
        headers["User-Agent"] = s.sec.user_agent
    # TWSE / MOPS serves certs missing "Subject Key Identifier" extension,
    # which Python 3.12+ OpenSSL rejects under strict verification. The user
    # has explicitly authorized verify=False for this source only; all other
    # markets keep strict verification with certifi's CA bundle.
    verify: ssl.SSLContext | bool = False if source == "twse" else _ssl_context()
    return httpx.Client(
        headers=headers,
        timeout=timeout,
        verify=verify,
        follow_redirects=True,
        http2=source != "dart",
    )


@lru_cache(maxsize=1)
def _ssl_context() -> ssl.SSLContext:
    """Cached SSL context using certifi's CA bundle.

    httpx 0.36+ deprecated ``verify=<str>``; pass an ``SSLContext`` instead.
    """
    return ssl.create_default_context(cafile=certifi.where())


@retry(
    stop=stop_after_attempt(3),
    wait=wait_exponential(multiplier=0.5, min=0.5, max=4),
    reraise=True,
)
def get_json(client: httpx.Client, url: str, *, source: str, rate: float, **kwargs: Any) -> Any:
    limiter(source, rate).acquire_blocking()
    r = client.get(url, **kwargs)
    r.raise_for_status()
    return r.json()


@retry(
    stop=stop_after_attempt(3),
    wait=wait_exponential(multiplier=0.5, min=0.5, max=4),
    reraise=True,
)
def post_json(
    client: httpx.Client,
    url: str,
    *,
    source: str,
    rate: float,
    data: dict | None = None,
    json_body: dict | None = None,
    **kwargs: Any,
) -> Any:
    limiter(source, rate).acquire_blocking()
    r = client.post(url, data=data, json=json_body, **kwargs)
    r.raise_for_status()
    return r.json()


@retry(
    stop=stop_after_attempt(3),
    wait=wait_exponential(multiplier=0.5, min=0.5, max=4),
    reraise=True,
)
def stream_to_file(
    client: httpx.Client,
    url: str,
    dest: Path,
    *,
    source: str,
    rate: float,
    chunk_size: int = 65536,
    read_timeout: float | None = None,
) -> int:
    """Stream a GET response to disk. Returns bytes written."""
    limiter(source, rate).acquire_blocking()
    dest.parent.mkdir(parents=True, exist_ok=True)
    tmp_dest = dest.with_name(f"{dest.name}.part")
    n = 0
    try:
        stream_timeout = httpx.Timeout(
            connect=30.0,
            read=read_timeout,
            write=30.0,
            pool=30.0,
        )
        with client.stream("GET", url, timeout=stream_timeout) as r:
            r.raise_for_status()
            with tmp_dest.open("wb") as f:
                for chunk in r.iter_bytes(chunk_size):
                    f.write(chunk)
                    n += len(chunk)
        tmp_dest.replace(dest)
    except Exception:
        tmp_dest.unlink(missing_ok=True)
        raise
    logger.debug(f"downloaded {url} -> {dest} ({n} bytes)")
    return n
