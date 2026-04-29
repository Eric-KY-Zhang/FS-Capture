"""US financials via SEC XBRL companyfacts API.

We map a curated set of US-GAAP / IFRS-US tags to our canonical Chinese metric
names so output rows align with A-share / HK / KR data.

Coverage is intentionally minimal in v1 — extend `_GAAP_MAPPING` to widen.
"""
from __future__ import annotations

from typing import Optional

from loguru import logger

from app.core.http import default_client, get_json
from app.core.models import (
    FinancialStatement,
    Period,
    PeriodType,
    StatementType,
    Ticker,
)
from plugins.us.reports import _filter_filings


_FACTS_URL = "https://data.sec.gov/api/xbrl/companyfacts/CIK{cik}.json"
_SUBMISSIONS_URL = "https://data.sec.gov/submissions/CIK{cik}.json"


# (statement, canonical_zh) -> list of GAAP tag candidates (first match wins)
_GAAP_MAPPING: dict[tuple[StatementType, str], list[str]] = {
    # 资产负债表
    (StatementType.BALANCE_SHEET, "货币资金"): [
        "CashAndCashEquivalentsAtCarryingValue",
        "Cash",
    ],
    (StatementType.BALANCE_SHEET, "应收账款"): ["AccountsReceivableNetCurrent"],
    (StatementType.BALANCE_SHEET, "存货"): ["InventoryNet"],
    (StatementType.BALANCE_SHEET, "流动资产合计"): ["AssetsCurrent"],
    (StatementType.BALANCE_SHEET, "长期股权投资"): ["LongTermInvestments"],
    (StatementType.BALANCE_SHEET, "固定资产"): ["PropertyPlantAndEquipmentNet"],
    (StatementType.BALANCE_SHEET, "无形资产"): ["IntangibleAssetsNetExcludingGoodwill"],
    (StatementType.BALANCE_SHEET, "商誉"): ["Goodwill"],
    (StatementType.BALANCE_SHEET, "非流动资产合计"): ["AssetsNoncurrent"],
    (StatementType.BALANCE_SHEET, "资产总计"): ["Assets"],
    (StatementType.BALANCE_SHEET, "短期借款"): ["ShortTermBorrowings", "LongTermDebtCurrent"],
    (StatementType.BALANCE_SHEET, "合同负债"): ["ContractWithCustomerLiabilityCurrent"],
    (StatementType.BALANCE_SHEET, "商业票据"): ["CommercialPaper"],
    (StatementType.BALANCE_SHEET, "应付账款"): ["AccountsPayableCurrent"],
    (StatementType.BALANCE_SHEET, "流动负债合计"): ["LiabilitiesCurrent"],
    (StatementType.BALANCE_SHEET, "长期借款"): ["LongTermDebtNoncurrent"],
    (StatementType.BALANCE_SHEET, "非流动负债合计"): ["LiabilitiesNoncurrent"],
    (StatementType.BALANCE_SHEET, "负债合计"): ["Liabilities"],
    (StatementType.BALANCE_SHEET, "实收资本"): ["CommonStockValue"],
    (StatementType.BALANCE_SHEET, "资本公积"): ["CommonStocksIncludingAdditionalPaidInCapital"],
    (StatementType.BALANCE_SHEET, "未分配利润"): ["RetainedEarningsAccumulatedDeficit"],
    (StatementType.BALANCE_SHEET, "其他综合收益"): ["AccumulatedOtherComprehensiveIncomeLossNetOfTax"],
    (StatementType.BALANCE_SHEET, "股东权益合计"): [
        "StockholdersEquity",
        "StockholdersEquityIncludingPortionAttributableToNoncontrollingInterest",
    ],
    # 利润表
    (StatementType.INCOME, "营业收入"): [
        "RevenueFromContractWithCustomerExcludingAssessedTax",
        "Revenues",
        "SalesRevenueNet",
        "SalesRevenueGoodsNet",
    ],
    (StatementType.INCOME, "营业成本"): ["CostOfGoodsAndServicesSold", "CostOfRevenue"],
    (StatementType.INCOME, "毛利"): ["GrossProfit"],
    (StatementType.INCOME, "研发费用"): ["ResearchAndDevelopmentExpense"],
    (StatementType.INCOME, "销售费用"): ["SellingAndMarketingExpense", "SellingGeneralAndAdministrativeExpense"],
    (StatementType.INCOME, "管理费用"): ["GeneralAndAdministrativeExpense"],
    (StatementType.INCOME, "营业利润"): ["OperatingIncomeLoss"],
    (StatementType.INCOME, "营业外收支"): ["NonoperatingIncomeExpense"],
    (StatementType.INCOME, "利润总额"): ["IncomeLossFromContinuingOperationsBeforeIncomeTaxesExtraordinaryItemsNoncontrollingInterest"],
    (StatementType.INCOME, "所得税"): ["IncomeTaxExpenseBenefit"],
    (StatementType.INCOME, "净利润"): ["NetIncomeLoss"],
    (StatementType.INCOME, "基本每股收益"): ["EarningsPerShareBasic"],
    (StatementType.INCOME, "稀释每股收益"): ["EarningsPerShareDiluted"],
    # 现金流量表
    (StatementType.CASH_FLOW, "经营活动现金流量净额"): ["NetCashProvidedByUsedInOperatingActivities"],
    (StatementType.CASH_FLOW, "购建固定资产、无形资产和其他长期资产支付的现金"): ["PaymentsToAcquirePropertyPlantAndEquipment"],
    (StatementType.CASH_FLOW, "其他投资活动现金流"): ["PaymentsForProceedsFromOtherInvestingActivities"],
    (StatementType.CASH_FLOW, "投资活动现金流量净额"): ["NetCashProvidedByUsedInInvestingActivities"],
    (StatementType.CASH_FLOW, "分配股利、利润或偿付利息支付的现金"): ["PaymentsOfDividends"],
    (StatementType.CASH_FLOW, "回购股份支付的现金"): ["PaymentsForRepurchaseOfCommonStock"],
    (StatementType.CASH_FLOW, "取得借款收到的现金"): ["ProceedsFromIssuanceOfDebt"],
    (StatementType.CASH_FLOW, "偿还债务支付的现金"): ["RepaymentsOfDebt"],
    (StatementType.CASH_FLOW, "筹资活动现金流量净额"): ["NetCashProvidedByUsedInFinancingActivities"],
    (StatementType.CASH_FLOW, "现金净增加额"): ["CashCashEquivalentsRestrictedCashAndRestrictedCashEquivalentsPeriodIncreaseDecreaseIncludingExchangeRateEffect"],
}


def _target_filing(client, cik: str, period: Period) -> Optional[dict]:
    payload = get_json(client, _SUBMISSIONS_URL.format(cik=cik), source="sec", rate=8.0)
    rows = _filter_filings(client, payload.get("filings") or {}, period)
    if not rows:
        return None
    rows.sort(key=lambda r: (("/A" in r["form"]), r["filingDate"]))
    return rows[0]


def _unit_priority(unit: str) -> int:
    if unit == "USD":
        return 0
    if unit == "USD/shares":
        return 1
    if unit == "shares":
        return 2
    return 3


def _select_period_value(units_dict: dict, period: Period, filing: Optional[dict]) -> Optional[float]:
    """A units_dict looks like {'USD': [{val, fp, fy, ...}, ...]}.
    Pick the value matching the requested period.
    """
    target_fy = period.year
    target_fp = {
        PeriodType.ANNUAL: "FY",
        PeriodType.Q1: "Q1",
        PeriodType.Q2: "Q2",
        PeriodType.Q3: "Q3",
    }[period.type]
    target_accn = filing.get("accessionNumber") if filing else None
    target_end = filing.get("reportDate") if filing else None

    candidates: list[tuple[str, dict]] = []
    for unit, rows in (units_dict or {}).items():
        for row in rows:
            if row.get("fp") == target_fp and row.get("fy") == target_fy:
                candidates.append((unit, row))

    if target_accn:
        accn_matches = [(u, r) for u, r in candidates if r.get("accn") == target_accn]
        if accn_matches:
            candidates = accn_matches

    if target_end:
        end_matches = [(u, r) for u, r in candidates if r.get("end") == target_end]
        if end_matches:
            candidates = end_matches

    candidates.sort(key=lambda item: (
        _unit_priority(item[0]),
        item[1].get("frame") is not None,
        item[1].get("filed") or "",
    ))
    for _, row in candidates:
        value = row.get("val")
        if isinstance(value, (int, float)):
            return float(value)
    return None


def fetch(ticker: Ticker, period: Period) -> list[FinancialStatement]:
    cik = ticker.external_id
    if not cik:
        return []

    url = _FACTS_URL.format(cik=cik)
    with default_client(source="sec") as client:
        try:
            filing = _target_filing(client, cik, period)
        except Exception as exc:
            logger.warning(f"SEC submissions lookup failed for {ticker.code}: {exc}")
            filing = None
        try:
            payload = get_json(client, url, source="sec", rate=8.0)
        except Exception as exc:
            logger.warning(f"SEC companyfacts failed for {ticker.code}: {exc}")
            return []

    facts = payload.get("facts") or {}
    gaap = facts.get("us-gaap") or facts.get("ifrs-full") or {}

    by_stmt: dict[StatementType, dict[str, Optional[float]]] = {
        StatementType.BALANCE_SHEET: {},
        StatementType.INCOME: {},
        StatementType.CASH_FLOW: {},
    }

    for (st_type, zh_name), candidates in _GAAP_MAPPING.items():
        for tag in candidates:
            entry = gaap.get(tag)
            if not entry:
                continue
            val = _select_period_value(entry.get("units") or {}, period, filing)
            if val is not None:
                by_stmt[st_type][zh_name] = val
                break

    out: list[FinancialStatement] = []
    for st_type, lines in by_stmt.items():
        if lines:
            out.append(FinancialStatement(
                ticker=ticker,
                period=period,
                statement_type=st_type,
                currency="USD",
                unit="美元",
                lines=lines,
            ))
    return out
