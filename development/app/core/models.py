from __future__ import annotations

from enum import Enum
from typing import Optional

from pydantic import BaseModel, Field, field_validator


class Exchange(str, Enum):
    A_SHARE = "A"
    HK = "HK"
    US = "US"
    KR = "KR"

    @property
    def display_name(self) -> str:
        return {
            "A": "A股",
            "HK": "港股",
            "US": "美股",
            "KR": "韩股",
        }[self.value]


class PeriodType(str, Enum):
    ANNUAL = "annual"
    Q1 = "q1"
    Q2 = "q2"   # interim / 半年报
    Q3 = "q3"
    IPO_PROSPECTUS = "ipo_prospectus"

    @property
    def display_name(self) -> str:
        return {
            "annual": "年报",
            "q1": "一季报",
            "q2": "半年报",
            "q3": "三季报",
            "ipo_prospectus": "IPO 招股书",
        }[self.value]


class Period(BaseModel):
    year: int = Field(ge=1990, le=2100)
    type: PeriodType

    def label(self) -> str:
        return f"{self.year}{self.type.display_name}"


class Ticker(BaseModel):
    exchange: Exchange
    code: str
    # Filled after name_resolver runs. None means unresolved.
    name: Optional[str] = None
    # Exchange-specific ID (cninfo orgId, SEC CIK, DART corp_code, HKEX stock_id).
    external_id: Optional[str] = None

    @field_validator("code")
    @classmethod
    def normalize_code(cls, v: str) -> str:
        return v.strip().upper()


class Company(BaseModel):
    ticker: Ticker
    listing_date: Optional[str] = None
    industry: Optional[str] = None
    currency: str = "CNY"
    # Free-form extra fields per market.
    extra: dict = Field(default_factory=dict)


class ReportFile(BaseModel):
    """A downloaded disclosure document."""
    ticker: Ticker
    period: Optional[Period] = None
    kind: str   # annual_report | q1_report | interim_report | q3_report | ipo_prospectus
    local_path: str
    source_url: str
    title: Optional[str] = None
    file_size_bytes: Optional[int] = None
    filing_date: Optional[str] = None
    report_date: Optional[str] = None
    form: Optional[str] = None
    source_id: Optional[str] = None
    accession_number: Optional[str] = None
    is_amendment: bool = False
    sequence: Optional[int] = None
    source_format: Optional[str] = None
    output_format: Optional[str] = None

    @field_validator("filing_date", "report_date", mode="before")
    @classmethod
    def normalize_date_text(cls, v):
        if v is None or v == "":
            return None
        if hasattr(v, "isoformat"):
            return v.isoformat()
        return str(v)
