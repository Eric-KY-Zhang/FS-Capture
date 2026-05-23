from __future__ import annotations

# Hong Kong listed companies with a non-December fiscal year-end.
# Format: 5-digit HK ticker code -> fiscal year-end month.
# Sources verified against HKEX disclosures and company annual reports.
NON_DEC_FISCAL_YEAR: dict[str, int] = {
    # Fiscal year ending 31 March (Apr-Mar cycle, common for Greater China retail/tech)
    "09988": 3,  # Alibaba Group Holding Ltd.
    "00992": 3,  # Lenovo Group Ltd.
    "00151": 3,  # Want Want China Holdings Ltd.
    "01929": 3,  # Chow Tai Fook Jewellery Group Ltd.
    "00823": 3,  # Link Real Estate Investment Trust
    "03998": 3,  # Bosideng International Holdings Ltd.
    "00241": 3,  # Alibaba Health Information Technology Ltd.
    "01060": 3,  # Alibaba Pictures Group Ltd.
    # Fiscal year ending 30 June (HK property developers)
    "00017": 6,  # New World Development Co. Ltd.
    "00083": 6,  # Sino Land Co. Ltd.
    "00016": 6,  # Sun Hung Kai Properties Ltd.
    # Fiscal year ending 31 May (New Oriental group)
    "09901": 5,  # New Oriental Education & Technology Group Inc.
    "01797": 5,  # East Buy Holding Ltd.
}


def fiscal_year_end_month(hk_code: str) -> int:
    """Return fiscal year-end month for a Hong Kong ticker. Default is December."""
    return NON_DEC_FISCAL_YEAR.get(str(hk_code).zfill(5), 12)
