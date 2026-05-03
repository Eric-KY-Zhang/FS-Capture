# HK API Probe — Xueqiu 00700

Date: 2026-05-03  
Cookie source: `上市公司财务数据查询.xlsm` → `样本池!B5`  
Endpoint family: `https://stock.xueqiu.com/v5/stock/finance/hk/{kind}.json?symbol=00700&type=all&is_detail=true&count=8`

## Summary

- `00700` is the valid ticker format. All 4 HK endpoints returned HTTP `200`, `error_code=0`, and `data.list` length `8`.
- `700` is not usable for this API. It also returns HTTP `200` and `error_code=0`, but `data.list` is empty and `quote_name` is null.
- For Tencent (`00700`), `data.currency=CNY` and `data.currency_name=人民币` across all 4 endpoints. This conflicts with the Phase 4c assumption "HKD millions".
- HK field names are abbreviated snake_case-like keys (`ta`, `tlia`, `tto`, `nocf`, etc.), not the same as the US Xueqiu fallback fields (`total_assets`, `revenue`, `net_cash_provided_by_oa`, etc.). HK needs an independent Xueqiu field map.
- HK records include `sd`, `ed`, `report_name`, `report_date`, and `month_num`. They do not include US-style `report_annual` or `report_type_code` in the first 00700 annual record.

## Files Dumped

- `samples/xueqiu_HK_00700_balance.json`
- `samples/xueqiu_HK_00700_income.json`
- `samples/xueqiu_HK_00700_cash_flow.json`
- `samples/xueqiu_HK_00700_indicator.json`

## Endpoint Results

| Symbol | Kind | HTTP | error_code | list_len | currency | quote_name | last_report_name |
|---|---|---:|---:|---:|---|---|---|
| `00700` | balance | 200 | 0 | 8 | CNY / 人民币 | 腾讯控股 | 2025年报 |
| `00700` | income | 200 | 0 | 8 | CNY / 人民币 | 腾讯控股 | 2025年报 |
| `00700` | cash_flow | 200 | 0 | 8 | CNY / 人民币 | 腾讯控股 | 2025年报 |
| `00700` | indicator | 200 | 0 | 8 | CNY / 人民币 | 腾讯控股 | 2025年报 |
| `700` | balance | 200 | 0 | 0 | CNY / 人民币 | null | null |
| `700` | income | 200 | 0 | 0 | CNY / 人民币 | null | null |
| `700` | cash_flow | 200 | 0 | 0 | CNY / 人民币 | null | null |
| `700` | indicator | 200 | 0 | 0 | CNY / 人民币 | null | null |

## First Record Metadata

All 4 `00700` endpoints have the same first-record period metadata:

```json
{
  "report_name": "2025年报",
  "report_date": 1767110400000,
  "sd": "2025-01-01",
  "ed": "2025-12-31",
  "month_num": 12
}
```

## Field Names

### balance

HK first record key count: `41`  
US POM balance first record key count: `54`  
Common keys: `5`

HK keys:

```text
ca, caprx, cceq, clia, ctime, diftatclia, ed, fina, fxda, iga, inv, iv, ltdt, miint, month_num, nalia, ncalia, numtsh, otca, otltlia, otnc, otnca, otrx, otstdt, report_date, report_name, rpaculo, rptsourefc, sd, shhfd, shpm, stdt, ta, ta_tlia, teqy, tlia, tnca, tnclia, trpy, trrb, trx
```

Likely HK core mappings:

| Output | HK field candidate | US Xueqiu equivalent |
|---|---|---|
| Total assets | `ta` | `total_assets` |
| Total liabilities | `tlia` | `total_liab` |
| Total equity | `teqy` | `total_equity` |
| Cash & equivalents | `cceq` | `cce` / `total_cash` |
| Inventory | `inv` | `inventory` |
| Total current assets | `ca` | `total_current_assets` |
| Total current liabilities | `clia` | `total_current_liab` |
| Long-term debt | `ltdt` | `lt_debt` |
| Short-term debt | `stdt` | `st_debt` |

### income

HK first record key count: `38`  
US POM income first record key count: `35`  
Common keys: `5`

HK keys:

```text
admexp, amteqyhdcom, amtmiint, beps_aju, cmnshdiv, ctime, depaz, deps_aju, divdbups_ajupd, ed, fcgcost, gp, jtctletiascom, month_num, nosplitems, npdsubu, opeplo, opeploinclfincost, otcphio, otiog, otopeexp, ploashh, plobtx, plocyr, report_date, report_name, rptsourefc, rshdevexp, sd, slgcost, slgdstexp, sr_ta, tcphio, tipmcgpvs, topeexp, tto, tx, txexcliotx
```

Likely HK core mappings:

| Output | HK field candidate | US Xueqiu equivalent |
|---|---|---|
| Revenue | `tto` | `revenue` / `total_revenue` |
| Gross profit | `gp` | `gross_profit` |
| Operating income | `opeplo` or `opeploinclfincost` | `operating_income` |
| Net income | `ploashh` | `net_income` / `net_income_atcss` |
| R&D expense | `rshdevexp` | `rad_expenses` |
| Selling expense | `slgdstexp` | `marketing_selling_etc` |
| Administrative expense | `admexp` | no direct US key |

### cash_flow

HK first record key count: `31`  
US POM cash flow first record key count: `21`  
Common keys: `5`

HK keys:

```text
adtfxda, cceqbegyr, cceqeyr, ctime, dcinv, depaz, divp, divrc, dsfxda, ed, eqyfin, fxdiodtinstr, icdccceq, icinv, intp, intrc, lnrpa, month_num, ncfdchexrateot, ncfrldpty_finact, ncfrldpty_invact, nfcgcf, nicln, ninvcf, nocf, report_date, report_name, rpafxdiodtinstr, rptsourefc, sd, txprf
```

Likely HK core mappings:

| Output | HK field candidate | US Xueqiu equivalent |
|---|---|---|
| Cash from operations | `nocf` | `net_cash_provided_by_oa` |
| Cash from investing | `ninvcf` | `net_cash_used_in_ia` |
| Cash from financing | `nfcgcf` | `net_cash_used_in_fa` |
| Cash at beginning of period | `cceqbegyr` | `cce_at_boy` |
| Cash at end of period | `cceqeyr` | `cce_at_eoy` |
| FX effect | `ncfdchexrateot` | `effect_of_exchange_chg_on_cce` |
| Capex / PPE purchase | `fxdiodtinstr` or `rpafxdiodtinstr` | `payment_for_property_and_equip` |

### indicator

HK first record key count: `35`  
US POM indicator first record key count: `43`  
Common keys: `5`

HK keys:

```text
apycvspd, arbcvspd, beps, beps_aju, bps, cro, ctime, ctioro, divporo, dprerate, dps, ed, gpm, ivcvspd, lnrerate, month_num, ncfps, nfcgcfps, ninvcfps, nocfps, opemg, opps, ploashh, plobtxps, qro, report_date, report_name, roe, rota, rptsourefc, sd, tlia_ta, tsrps, tto, ttops
```

The Phase 4c plan says HK indicator should be generated from BS/IS using the standard 18 metrics, so these raw indicator keys are useful for spot-checking but should not drive the official `港股_指标表`.

## Decision Points Before Step 2

1. Unit/currency assumption needs review. The API reports Tencent HK financials as `CNY / 人民币`, not HKD. If we keep "不做汇率换算", the sheet unit should probably be "百万报告币种" or "百万人民币 for 00700", not hard-coded "百万港元".
2. Field mapping should be independent from US Xueqiu. Reusing US field names will miss most core fields.
3. Period matching should use `report_name` + `sd`/`ed`. HK records do not expose `report_annual` / `report_type_code` like US Xueqiu samples.
