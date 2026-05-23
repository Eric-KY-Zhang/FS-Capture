from __future__ import annotations

import sys
from types import SimpleNamespace

from plugins.ashare import name_resolver


class _Column(list):
    def astype(self, _dtype):
        return self


class _Frame:
    def __init__(self, **columns) -> None:
        self._columns = {key: _Column(value) for key, value in columns.items()}

    def __getitem__(self, key: str):
        return self._columns[key]


def test_zip_strict_logs_error_on_mismatch(monkeypatch) -> None:
    errors: list[str] = []
    fake_akshare = SimpleNamespace(
        stock_info_a_code_name=lambda: _Frame(code=["600519", "000001"], name=["贵州茅台"])
    )
    monkeypatch.setitem(sys.modules, "akshare", fake_akshare)
    monkeypatch.setattr(
        name_resolver,
        "cached_or_load",
        lambda _key, loader, *, expire: loader(),
    )
    monkeypatch.setattr(
        name_resolver,
        "logger",
        SimpleNamespace(info=lambda *_args, **_kwargs: None, error=errors.append),
    )

    name_map = name_resolver._load_name_map()

    assert name_map == {"600519": "贵州茅台"}
    assert errors
    assert "code/name 列长度不一致" in errors[0]
