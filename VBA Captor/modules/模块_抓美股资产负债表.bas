Attribute VB_Name = "模块_抓美股资产负债表"
Option Explicit

' =================================================================
'  抓美股 Balance Sheet — Phase 4b-1
'  数据源: SEC EDGAR companyfacts → us-gaap concepts
'  写入 Sheet: 美股_资产负债表
'  指标名 / 大类 用英文 (跟原 us-gaap 体系对齐, 用户便于跟 SEC 文件对照)
' =================================================================

Public Sub Main()
    RunUSStatement "BalanceSheet", "美股_资产负债表", GetBSConcepts(), 6
End Sub


Private Function GetBSConcepts() As Variant
    Dim a(0 To 26) As Variant
    a(0) = Array("Current Assets", "Cash & equivalents", _
                 "CashAndCashEquivalentsAtCarryingValue,CashCashEquivalentsRestrictedCashAndRestrictedCashEquivalents,CashAndCashEquivalentsAtCarryingValueIncludingDisposalGroupAndDiscontinuedOperations", _
                 "USD", 1000000#, _
                 "CashAndCashEquivalents,CashAndCashEquivalentsAtCarryingValue")
    a(1) = Array("Current Assets", "Marketable securities (current)", _
                 "MarketableSecuritiesCurrent,ShortTermInvestments,AvailableForSaleSecuritiesCurrent")
    a(2) = Array("Current Assets", "Accounts receivable, net", _
                 "AccountsReceivableNetCurrent,AccountsReceivableNet,ReceivablesNetCurrent,TradeAccountsReceivableNetCurrent", _
                 "USD", 1000000#, _
                 "TradeAndOtherCurrentReceivables,CurrentTradeReceivables")
    a(3) = Array("Current Assets", "Inventory", _
                 "InventoryNet,InventoriesNet,InventoryFinishedGoodsNetOfReserves", _
                 "USD", 1000000#, _
                 "Inventories,Inventory")
    a(4) = Array("Current Assets", "Other current assets", _
                 "OtherAssetsCurrent,OtherCurrentAssets,PrepaidExpenseAndOtherAssetsCurrent")
    a(5) = Array("Current Assets", "Total current assets", _
                 "AssetsCurrent", _
                 "USD", 1000000#, _
                 "CurrentAssets,AssetsCurrent")
    a(6) = Array("Non-Current Assets", "Marketable securities (non-current)", _
                 "MarketableSecuritiesNoncurrent,AvailableForSaleSecuritiesNoncurrent,LongTermInvestments,InvestmentsInDebtAndEquitySecurities")
    a(7) = Array("Non-Current Assets", "Property, plant & equipment, net", _
                 "PropertyPlantAndEquipmentNet,PropertyPlantAndEquipmentAndFinanceLeaseRightOfUseAssetAfterAccumulatedDepreciationAndAmortization", _
                 "USD", 1000000#, _
                 "PropertyPlantAndEquipment,PropertyPlantAndEquipmentNet")
    a(8) = Array("Non-Current Assets", "Goodwill", "Goodwill")
    a(9) = Array("Non-Current Assets", "Intangible assets", _
                 "IntangibleAssetsNetExcludingGoodwill,FiniteLivedIntangibleAssetsNet,IntangibleAssetsNet", _
                 "USD", 1000000#, _
                 "IntangibleAssetsOtherThanGoodwill,IntangibleAssets")
    a(10) = Array("Non-Current Assets", "Other non-current assets", _
                  "OtherAssetsNoncurrent,OtherNoncurrentAssets")
    a(11) = Array("Non-Current Assets", "Total non-current assets", _
                  "AssetsNoncurrent", _
                  "USD", 1000000#, _
                  "NoncurrentAssets,AssetsNoncurrent")
    a(12) = Array("", "Total assets", _
                  "Assets", _
                  "USD", 1000000#, _
                  "Assets")
    a(13) = Array("Current Liabilities", "Accounts payable", _
                  "AccountsPayableCurrent,AccountsPayableTradeCurrent", _
                  "USD", 1000000#, _
                  "TradeAndOtherCurrentPayables,CurrentTradePayables")
    a(14) = Array("Current Liabilities", "Short-term debt", _
                  "LongTermDebtCurrent,ShortTermBorrowings,ShortTermDebtCurrent,ShortTermBorrowingsAndCurrentPortionOfLongTermDebt")
    a(15) = Array("Current Liabilities", "Compensation & benefits", _
                  "EmployeeRelatedLiabilitiesCurrent,EmployeeBenefitsAndShareBasedCompensationCurrent")
    a(16) = Array("Current Liabilities", "Other current liabilities", _
                  "OtherLiabilitiesCurrent,OtherCurrentLiabilities")
    a(17) = Array("Current Liabilities", "Total current liabilities", _
                  "LiabilitiesCurrent", _
                  "USD", 1000000#, _
                  "CurrentLiabilities,LiabilitiesCurrent")
    a(18) = Array("Non-Current Liabilities", "Long-term debt", _
                  "LongTermDebtNoncurrent,LongTermDebtAndFinanceLeaseObligationsNoncurrent,LongTermDebt", _
                  "USD", 1000000#, _
                  "NoncurrentBorrowings,NoncurrentFinancialLiabilities")
    a(19) = Array("Non-Current Liabilities", "Other non-current liabilities", _
                  "OtherLiabilitiesNoncurrent,OtherNoncurrentLiabilities,DeferredTaxLiabilitiesNoncurrent")
    a(20) = Array("Non-Current Liabilities", "Total non-current liabilities", _
                  "LiabilitiesNoncurrent", _
                  "USD", 1000000#, _
                  "NoncurrentLiabilities,LiabilitiesNoncurrent")
    a(21) = Array("", "Total liabilities", _
                  "Liabilities", _
                  "USD", 1000000#, _
                  "Liabilities")
    a(22) = Array("Stockholders' Equity", "Common stock", _
                  "CommonStockValue,CommonStocksIncludingAdditionalPaidInCapital,CommonStockIncludingAdditionalPaidInCapital")
    a(23) = Array("Stockholders' Equity", "Retained earnings", _
                  "RetainedEarningsAccumulatedDeficit,RetainedEarnings", _
                  "USD", 1000000#, _
                  "RetainedEarnings")
    a(24) = Array("Stockholders' Equity", "Accumulated OCI", _
                  "AccumulatedOtherComprehensiveIncomeLossNetOfTax,AccumulatedOtherComprehensiveIncomeLoss")
    a(25) = Array("Stockholders' Equity", "Total stockholders' equity", _
                  "StockholdersEquity,StockholdersEquityIncludingPortionAttributableToNoncontrollingInterest", _
                  "USD", 1000000#, _
                  "Equity,EquityAttributableToOwnersOfParent")
    a(26) = Array("", "Total liabilities & equity", _
                  "LiabilitiesAndStockholdersEquity,LiabilitiesAndStockholdersEquityIncludingPortionAttributableToNoncontrollingInterest", _
                  "USD", 1000000#, _
                  "EquityAndLiabilities,LiabilitiesAndEquity")
    GetBSConcepts = a
End Function
