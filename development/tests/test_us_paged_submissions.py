from __future__ import annotations

from app.core.models import Period, PeriodType
from plugins.us import reports


def _table(rows: list[dict]) -> dict:
    columns = [
        "form",
        "accessionNumber",
        "primaryDocument",
        "primaryDocDescription",
        "filingDate",
        "reportDate",
        "acceptanceDateTime",
        "fileNumber",
        "size",
    ]
    return {column: [str(row.get(column, "")) for row in rows] for column in columns}


def test_filter_filings_falls_back_to_paged_files_when_recent_empty(monkeypatch) -> None:
    filings = {
        "recent": _table([]),
        "files": [{"name": "CIK0000789019-submissions-001.json"}],
    }
    paged = _table(
        [
            {
                "form": "10-K",
                "accessionNumber": "0000789019-24-000001",
                "primaryDocument": "msft-20240630.htm",
                "filingDate": "2024-07-30",
                "reportDate": "2024-06-30",
            }
        ]
    )
    monkeypatch.setattr(reports, "get_json", lambda *_args, **_kwargs: paged)

    rows = reports._filter_filings(object(), filings, Period(year=2024, type=PeriodType.ANNUAL))

    assert rows[0]["accessionNumber"] == "0000789019-24-000001"


def test_filter_filings_returns_empty_when_neither_recent_nor_files_match(monkeypatch) -> None:
    filings = {
        "recent": _table([]),
        "files": [{"name": "CIK0000789019-submissions-001.json"}],
    }
    paged = _table(
        [
            {
                "form": "10-K",
                "accessionNumber": "0000789019-23-000001",
                "primaryDocument": "msft-20230630.htm",
                "filingDate": "2023-07-30",
                "reportDate": "2023-06-30",
            }
        ]
    )
    monkeypatch.setattr(reports, "get_json", lambda *_args, **_kwargs: paged)

    assert reports._filter_filings(
        object(), filings, Period(year=2024, type=PeriodType.ANNUAL)
    ) == []


def test_all_filing_rows_merges_recent_and_files(monkeypatch) -> None:
    duplicate = {
        "form": "10-K",
        "accessionNumber": "0000789019-24-000001",
        "primaryDocument": "msft-20240630.htm",
        "filingDate": "2024-07-30",
        "reportDate": "2024-06-30",
    }
    filings = {
        "recent": _table([duplicate]),
        "files": [{"name": "CIK0000789019-submissions-001.json"}],
    }
    paged = _table(
        [
            duplicate,
            {
                "form": "S-1",
                "accessionNumber": "0000789019-86-000001",
                "primaryDocument": "msft-s1.htm",
                "filingDate": "1986-03-13",
                "reportDate": "",
            },
        ]
    )
    monkeypatch.setattr(reports, "get_json", lambda *_args, **_kwargs: paged)

    rows = reports._all_filing_rows(object(), filings)

    assert [row["form"] for row in rows] == ["10-K", "S-1"]


def test_paged_file_fetch_warns_on_error_and_continues(monkeypatch) -> None:
    warnings = []
    filings = {
        "recent": _table([]),
        "files": [{"name": "CIK0000789019-submissions-001.json"}],
    }

    def fail_fetch(*_args, **_kwargs):
        raise RuntimeError("network down")

    monkeypatch.setattr(reports, "get_json", fail_fetch)
    monkeypatch.setattr(reports.logger, "warning", lambda message: warnings.append(message))

    rows = reports._filter_filings(object(), filings, Period(year=2024, type=PeriodType.ANNUAL))

    assert rows == []
    assert warnings
    assert "CIK0000789019-submissions-001.json" in warnings[0]
