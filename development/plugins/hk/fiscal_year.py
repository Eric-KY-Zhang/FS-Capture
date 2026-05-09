from __future__ import annotations


# Hong Kong listed companies with a non-December fiscal year-end.
# Format: 5-digit HK ticker code -> fiscal year-end month.
NON_DEC_FISCAL_YEAR: dict[str, int] = {
    "09988": 3,  # Alibaba Group Holding Ltd.
}


def fiscal_year_end_month(hk_code: str) -> int:
    """Return fiscal year-end month for a Hong Kong ticker. Default is December."""
    return NON_DEC_FISCAL_YEAR.get(str(hk_code).zfill(5), 12)
