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

    @property
    def display_name(self) -> str:
        return {
            "annual": "年报",
            "q1": "一季报",
            "q2": "半年报",
            "q3": "三季报",
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


class StatementType(str, Enum):
    BALANCE_SHEET = "balance_sheet"      # 资产负债表
    INCOME = "income"                    # 利润表
    CASH_FLOW = "cash_flow"              # 现金流量表


class FinancialStatement(BaseModel):
    """Canonical financial statement representation.

    `lines` maps Chinese metric name → value. Names match the existing 瑞华底稿
    schema so downstream Excel formulas keep working unchanged.
    """
    ticker: Ticker
    period: Period
    statement_type: StatementType
    currency: str = "CNY"
    unit: str = "元"   # 元 | 千元 | 万元 | 百万元
    lines: dict[str, Optional[float]] = Field(default_factory=dict)


class ReportFile(BaseModel):
    """A downloaded report PDF/HTML."""
    ticker: Ticker
    period: Period
    kind: str   # "annual_report" | "audit_report" | "q1_report" | "q3_report" | "interim_report"
    local_path: str
    source_url: str
    title: Optional[str] = None
    file_size_bytes: Optional[int] = None
