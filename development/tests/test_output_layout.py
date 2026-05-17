from __future__ import annotations

from pathlib import Path

from app.core.models import Exchange, Period, PeriodType, Ticker
from app.core.output_paths import report_output_path, report_output_path_for_filing


def _ticker(name: str | None = "贵州茅台") -> Ticker:
    return Ticker(exchange=Exchange.A_SHARE, code="600519", name=name)


def test_annual_report_filename_includes_company_name_and_chinese_kind() -> None:
    path = report_output_path(
        Path("output"), _ticker(), Period(year=2024, type=PeriodType.ANNUAL), "annual_report", ".pdf"
    )
    assert path.parent == Path("output")
    assert path.name == "A_600519_贵州茅台_2024_年报.pdf"


def test_q1_report_filename_uses_一季报() -> None:
    path = report_output_path(
        Path("output"), _ticker(), Period(year=2025, type=PeriodType.Q1), "q1_report", ".pdf"
    )
    assert path.name == "A_600519_贵州茅台_2025_一季报.pdf"


def test_q2_interim_report_uses_半年报() -> None:
    path = report_output_path(
        Path("output"), _ticker(), Period(year=2025, type=PeriodType.Q2), "interim_report", ".pdf"
    )
    assert path.name == "A_600519_贵州茅台_2025_半年报.pdf"


def test_q3_report_uses_三季报() -> None:
    path = report_output_path(
        Path("output"), _ticker(), Period(year=2025, type=PeriodType.Q3), "q3_report", ".pdf"
    )
    assert path.name == "A_600519_贵州茅台_2025_三季报.pdf"


def test_audit_report_uses_审计报告() -> None:
    path = report_output_path(
        Path("output"), _ticker(), Period(year=2024, type=PeriodType.ANNUAL), "audit_report", ".pdf"
    )
    assert path.name == "A_600519_贵州茅台_2024_审计报告.pdf"


def test_filename_drops_company_name_segment_when_unresolved() -> None:
    path = report_output_path(
        Path("output"),
        _ticker(name=None),
        Period(year=2024, type=PeriodType.ANNUAL),
        "annual_report",
        ".pdf",
    )
    # No empty `__` between code and year — company segment is simply omitted.
    assert path.name == "A_600519_2024_年报.pdf"


def test_ipo_prospectus_filing_uses_招股书_chinese_label() -> None:
    path = report_output_path_for_filing(
        Path("output"),
        _ticker(),
        "ipo",
        1,
        "ipo_prospectus",
        ".pdf",
        filing_date="2024-04-01",
    )
    assert path.name == "A_600519_贵州茅台_招股书_2024-04-01_001.pdf"


def test_ipo_filing_with_amendment_appends_补充() -> None:
    path = report_output_path_for_filing(
        Path("output"),
        _ticker(),
        "ipo",
        2,
        "ipo_prospectus",
        ".pdf",
        filing_date="2024-05-15",
        is_amendment=True,
    )
    assert path.name == "A_600519_贵州茅台_招股书_2024-05-15_002_补充.pdf"


def test_ipo_filing_with_unmapped_label_falls_back_to_招股书_for_ipo_family() -> None:
    path = report_output_path_for_filing(
        Path("output"),
        _ticker(),
        "ipo",
        3,
        "preliminary_prospectus",
        ".pdf",
        filing_date="2024-06-30",
    )
    # 'preliminary_prospectus' is not in _KIND_ZH but family contains 'ipo' → 招股书
    assert path.name == "A_600519_贵州茅台_招股书_2024-06-30_003.pdf"
