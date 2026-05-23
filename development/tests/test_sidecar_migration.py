from __future__ import annotations

import json

from app.core.sidecar import migrate_legacy_sidecars


def _write_legacy(path, exchange: str | None, ticker_code: str) -> None:
    payload = {
        "ticker_code": ticker_code,
        "kind": "annual_report",
        "source_url": "https://example.com/report.pdf",
    }
    if exchange is not None:
        payload["exchange"] = exchange
    path.write_text(json.dumps(payload, ensure_ascii=False), encoding="utf-8")


def test_migrate_legacy_sidecars_moves_exchange_payloads(tmp_path) -> None:
    output_root = tmp_path / "output"
    output_root.mkdir()
    cache_root = tmp_path / "cache"
    legacy_files = [
        output_root / "A_600519_2024_年报.pdf.meta.json",
        output_root / "HK_00700_2024_年报.pdf.meta.json",
        output_root / "US_AAPL_2024_年报.pdf.meta.json",
    ]
    _write_legacy(legacy_files[0], "A", "600519")
    _write_legacy(legacy_files[1], "HK", "00700")
    _write_legacy(legacy_files[2], "US", "AAPL")

    moved = migrate_legacy_sidecars(output_root, cache_root)

    assert moved == 3
    assert not any(path.exists() for path in legacy_files)
    assert (cache_root / "sidecars" / "A" / "A_600519_2024_年报.meta.json").exists()
    hk_payload = json.loads(
        (cache_root / "sidecars" / "HK" / "HK_00700_2024_年报.meta.json").read_text(
            encoding="utf-8"
        )
    )
    assert hk_payload["exchange"] == "HK"
    assert hk_payload["ticker_code"] == "00700"
    assert (cache_root / "sidecars" / "US" / "US_AAPL_2024_年报.meta.json").exists()


def test_migrate_legacy_sidecars_skips_orphan_without_exchange(tmp_path) -> None:
    output_root = tmp_path / "output"
    output_root.mkdir()
    cache_root = tmp_path / "cache"
    orphan = output_root / "orphan.meta.json"
    _write_legacy(orphan, None, "600519")

    moved = migrate_legacy_sidecars(output_root, cache_root)

    assert moved == 0
    assert orphan.exists()
    assert not (cache_root / "sidecars").exists()


def test_migrate_legacy_sidecars_keeps_cross_exchange_code_separate(tmp_path) -> None:
    output_root = tmp_path / "output"
    output_root.mkdir()
    cache_root = tmp_path / "cache"
    hk = output_root / "HK_00005_2024_年报.pdf.meta.json"
    us = output_root / "US_00005_2024_年报.pdf.meta.json"
    _write_legacy(hk, "HK", "00005")
    _write_legacy(us, "US", "00005")

    moved = migrate_legacy_sidecars(output_root, cache_root)

    assert moved == 2
    assert (cache_root / "sidecars" / "HK" / "HK_00005_2024_年报.meta.json").exists()
    assert (cache_root / "sidecars" / "US" / "US_00005_2024_年报.meta.json").exists()
