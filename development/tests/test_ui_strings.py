from __future__ import annotations

import ast
import re
from pathlib import Path

from app.ui import strings as S

_CJK_RE = re.compile(r"[\u3400-\u9fff]")
_UI_ROOT = Path(__file__).parents[1] / "app" / "ui"


def _string_literals(path: Path) -> list[tuple[int, str]]:
    tree = ast.parse(path.read_text(encoding="utf-8"))
    found: list[tuple[int, str]] = []

    class _Visitor(ast.NodeVisitor):
        def visit_Constant(self, node: ast.Constant) -> None:  # noqa: N802
            if isinstance(node.value, str):
                found.append((node.lineno, node.value))

        def visit_JoinedStr(self, node: ast.JoinedStr) -> None:  # noqa: N802
            for part in node.values:
                if isinstance(part, ast.Constant) and isinstance(part.value, str):
                    found.append((part.lineno, part.value))
            self.generic_visit(node)

    _Visitor().visit(tree)
    return found


def test_no_cjk_string_literals_remain_in_ui_modules() -> None:
    offenders: list[str] = []
    for path in sorted(_UI_ROOT.glob("*.py")):
        if path.name == "strings.py":
            continue
        for line_no, value in _string_literals(path):
            if _CJK_RE.search(value):
                offenders.append(f"{path.name}:{line_no}:{value}")

    assert offenders == []


def test_strings_zh_and_en_have_identical_keys() -> None:
    assert set(S.STRINGS["zh"]) == set(S.STRINGS["en"])


def test_strings_en_values_contain_no_cjk() -> None:
    offenders = {
        name: value for name, value in S.STRINGS["en"].items() if _CJK_RE.search(value)
    }

    assert offenders == {}


def test_strings_zh_values_contain_cjk() -> None:
    offenders = {
        name: value for name, value in S.STRINGS["zh"].items() if not _CJK_RE.search(value)
    }

    assert offenders == {}


def test_ui_strings_avoid_qt_translate_wrappers() -> None:
    source = (_UI_ROOT / "strings.py").read_text(encoding="utf-8")
    assert "QCoreApplication.translate" not in source
