from __future__ import annotations

import ast
from pathlib import Path

_SPEC = Path(__file__).parents[1] / "filings_atlas.spec"


def _analysis_keyword_values(keyword_name: str) -> set[str]:
    tree = ast.parse(_SPEC.read_text(encoding="utf-8"))
    values: set[str] = set()

    for node in ast.walk(tree):
        if not isinstance(node, ast.Call):
            continue
        if not isinstance(node.func, ast.Name) or node.func.id != "Analysis":
            continue
        for keyword in node.keywords:
            if keyword.arg != keyword_name:
                continue
            for child in ast.walk(keyword.value):
                if isinstance(child, ast.Constant) and isinstance(child.value, str):
                    values.add(child.value)

    return values


def test_pandas_plotting_is_kept_for_frozen_akshare_runtime() -> None:
    hiddenimports = _analysis_keyword_values("hiddenimports")
    excludes = _analysis_keyword_values("excludes")

    assert "pandas.plotting" in hiddenimports
    assert "pandas.plotting" not in excludes
