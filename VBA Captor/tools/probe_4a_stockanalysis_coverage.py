from __future__ import annotations

import time
from pathlib import Path

import requests

ROOT = Path(__file__).resolve().parent.parent
SAMPLES = ROOT / "samples"

UA = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
    "(KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
)

HEADERS = {
    "User-Agent": UA,
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
    "Accept-Language": "en-US,en;q=0.9,zh-CN;q=0.8",
    "Accept-Encoding": "identity",
}

TARGETS = [
    ("00700", "https://stockanalysis.com/quote/hkg/00700/financials/"),
    ("02519", "https://stockanalysis.com/quote/hkg/02519/financials/"),
    ("09988", "https://stockanalysis.com/quote/hkg/09988/financials/"),
    ("BABA", "https://stockanalysis.com/stocks/baba/financials/"),
    ("JD", "https://stockanalysis.com/stocks/jd/financials/"),
    ("PDD", "https://stockanalysis.com/stocks/pdd/financials/"),
]


def main() -> None:
    SAMPLES.mkdir(exist_ok=True)
    counts: dict[int, int] = {}

    for idx, (ticker, url) in enumerate(TARGETS):
        resp = requests.get(url, headers=HEADERS, timeout=30)
        counts[resp.status_code] = counts.get(resp.status_code, 0) + 1

        out_path = SAMPLES / f"stockanalysis_{ticker}.html"
        out_path.write_text(resp.text, encoding=resp.encoding or "utf-8", errors="replace")
        print(
            f"{ticker}: HTTP {resp.status_code}, {len(resp.content)} bytes, saved to {out_path}",
            flush=True,
        )

        if idx < len(TARGETS) - 1:
            time.sleep(2)

    ok = counts.get(200, 0)
    not_found = counts.get(404, 0)
    other = len(TARGETS) - ok - not_found
    print(f"SUMMARY: {ok}/6 HTTP 200, {not_found}/6 HTTP 404, {other}/6 other", flush=True)


if __name__ == "__main__":
    main()
