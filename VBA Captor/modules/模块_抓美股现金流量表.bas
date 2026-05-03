Attribute VB_Name = "模块_抓美股现金流量表"
Option Explicit

' =================================================================
'  抓美股 Cash Flow Statement — Phase 4b-2
'  CF 同 IS, duration 区间, YTD 选 min(start) 偏好
'  全部 USD, 默认 scale=1e6 (3-tuple)
' =================================================================

Public Sub Main()
    RunUSStatement "CashFlow", "美股_现金流量表", GetCFConcepts(), 6
End Sub


Private Function GetCFConcepts() As Variant
    Dim a(0 To 33) As Variant
    a(0) = Array("Operating", "Net income", _
                 "NetIncomeLoss,ProfitLoss", _
                 "USD", 1000000#, _
                 "ProfitLoss,ProfitLossAttributableToOwnersOfParent")
    a(1) = Array("Operating", "Depreciation & amortization", _
                 "DepreciationDepletionAndAmortization,DepreciationAmortizationAndAccretionNet,DepreciationAndAmortization,Depreciation")
    a(2) = Array("Operating", "Stock-based compensation", _
                 "ShareBasedCompensation,AllocatedShareBasedCompensationExpense")
    a(3) = Array("Operating", "Deferred income taxes", _
                 "DeferredIncomeTaxExpenseBenefit,DeferredTaxExpenseFromIncomeStatement")
    a(4) = Array("Operating", "Other non-cash items", _
                 "OtherNoncashIncomeExpense,OtherNoncashExpense")
    a(5) = Array("Operating", "Change in AR", _
                 "IncreaseDecreaseInAccountsReceivable,IncreaseDecreaseInReceivables,IncreaseDecreaseInOtherReceivables")
    a(6) = Array("Operating", "Change in inventory", _
                 "IncreaseDecreaseInInventories,IncreaseDecreaseInInventoriesNet,IncreaseDecreaseInInventory")
    a(7) = Array("Operating", "Change in AP", _
                 "IncreaseDecreaseInAccountsPayable,IncreaseDecreaseInAccountsPayableAndAccruedLiabilities")
    a(8) = Array("Operating", "Change in deferred revenue", _
                 "IncreaseDecreaseInDeferredRevenue,IncreaseDecreaseInContractWithCustomerLiability,IncreaseDecreaseInDeferredRevenueAndCustomerAdvancesAndDeposits")
    a(9) = Array("Operating", "Change in other operating assets", _
                 "IncreaseDecreaseInOtherOperatingAssets,IncreaseDecreaseInOtherCurrentAssets,IncreaseDecreaseInPrepaidDeferredExpenseAndOtherAssets")
    a(10) = Array("Operating", "Change in other operating liabilities", _
                  "IncreaseDecreaseInOtherOperatingLiabilities,IncreaseDecreaseInOtherCurrentLiabilities,IncreaseDecreaseInAccruedLiabilities")
    a(11) = Array("Operating", "Cash from operations", _
                  "NetCashProvidedByUsedInOperatingActivities,NetCashProvidedByUsedInOperatingActivitiesContinuingOperations", _
                  "USD", 1000000#, _
                  "CashFlowsFromUsedInOperatingActivities,NetCashFlowsFromUsedInOperatingActivities")
    a(12) = Array("Investing", "Purchases of marketable securities", _
                  "PaymentsToAcquireMarketableSecurities,PaymentsToAcquireAvailableForSaleSecurities,PaymentsToAcquireAvailableForSaleSecuritiesDebt,PaymentsToAcquireDebtSecurities")
    a(13) = Array("Investing", "Proceeds from maturities of marketable securities", _
                  "ProceedsFromMaturitiesPrepaymentsAndCallsOfMarketableSecurities,ProceedsFromMaturitiesPrepaymentsAndCallsOfAvailableForSaleSecurities,ProceedsFromCollectionOfMaturedAvailableForSaleSecurities")
    a(14) = Array("Investing", "Proceeds from sales of marketable securities", _
                  "ProceedsFromSaleOfMarketableSecurities,ProceedsFromSaleOfAvailableForSaleSecurities,ProceedsFromSaleOfAvailableForSaleSecuritiesDebt,ProceedsFromSaleOfDebtSecurities")
    a(15) = Array("Investing", "Purchases of investments", _
                  "PaymentsToAcquireInvestments,PaymentsToAcquireOtherInvestments,PaymentsToAcquireEquitySecurities")
    a(16) = Array("Investing", "Capex", _
                  "PaymentsToAcquirePropertyPlantAndEquipment,PaymentsToAcquireProductiveAssets,PaymentsToAcquirePropertyPlantAndEquipmentClassifiedAsInvestingActivities", _
                  "USD", 1000000#, _
                  "PurchaseOfPropertyPlantAndEquipmentClassifiedAsInvestingActivities,PaymentsToAcquirePropertyPlantAndEquipmentClassifiedAsInvestingActivities")
    a(17) = Array("Investing", "Business acquisitions", _
                  "PaymentsToAcquireBusinessesNetOfCashAcquired,PaymentsToAcquireBusinessesGross")
    a(18) = Array("Investing", "Proceeds from sale of PP&E", _
                  "ProceedsFromSaleOfPropertyPlantAndEquipment,ProceedsFromSaleOfProductiveAssets")
    a(19) = Array("Investing", "Other investing", _
                  "PaymentsForProceedsFromOtherInvestingActivities,PaymentsForProceedsFromOtherInvestingActivitiesNet")
    a(20) = Array("Investing", "Cash from investing", _
                  "NetCashProvidedByUsedInInvestingActivities,NetCashProvidedByUsedInInvestingActivitiesContinuingOperations", _
                  "USD", 1000000#, _
                  "CashFlowsFromUsedInInvestingActivities,NetCashFlowsFromUsedInInvestingActivities")
    a(21) = Array("Financing", "Dividends paid", _
                  "PaymentsOfDividends,PaymentsOfDividendsCommonStock,PaymentsOfOrdinaryDividends")
    a(22) = Array("Financing", "Stock repurchases", _
                  "PaymentsForRepurchaseOfCommonStock,PaymentsForRepurchaseOfEquity")
    a(23) = Array("Financing", "Stock issuance", _
                  "ProceedsFromIssuanceOfCommonStock,ProceedsFromStockOptionsExercised,ProceedsFromIssuanceOfShares")
    a(24) = Array("Financing", "Tax withholding on share awards", _
                  "PaymentsRelatedToTaxWithholdingForShareBasedCompensation,PaymentsForTaxWithholdingForShareBasedCompensation")
    a(25) = Array("Financing", "Long-term debt issued", _
                  "ProceedsFromIssuanceOfLongTermDebt")
    a(26) = Array("Financing", "Long-term debt repaid", _
                  "RepaymentsOfLongTermDebt")
    a(27) = Array("Financing", "Short-term borrowings, net", _
                  "ProceedsFromRepaymentsOfShortTermDebt,ProceedsFromRepaymentsOfCommercialPaper,ProceedsFromRepaymentsOfLinesOfCredit")
    a(28) = Array("Financing", "Other financing", _
                  "ProceedsFromPaymentsForOtherFinancingActivities,ProceedsFromPaymentsForOtherFinancingActivitiesNet")
    a(29) = Array("Financing", "Cash from financing", _
                  "NetCashProvidedByUsedInFinancingActivities,NetCashProvidedByUsedInFinancingActivitiesContinuingOperations", _
                  "USD", 1000000#, _
                  "CashFlowsFromUsedInFinancingActivities,NetCashFlowsFromUsedInFinancingActivities")
    a(30) = Array("", "FX effect on cash", _
                  "EffectOfExchangeRateOnCashCashEquivalentsRestrictedCashAndRestrictedCashEquivalents,EffectOfExchangeRateOnCashAndCashEquivalents", _
                  "USD", 1000000#, _
                  "EffectOfExchangeRateChangesOnCashAndCashEquivalents")
    a(31) = Array("", "Net change in cash (incl FX)", _
                  "CashCashEquivalentsRestrictedCashAndRestrictedCashEquivalentsPeriodIncreaseDecreaseIncludingExchangeRateEffect,CashAndCashEquivalentsPeriodIncreaseDecrease", _
                  "USD", 1000000#, _
                  "IncreaseDecreaseInCashAndCashEquivalents")
    a(32) = Array("", "Cash at beginning of period", _
                  "NoEdgarConceptCashAtBeginning", _
                  "USD", 1000000#, _
                  "CashAndCashEquivalentsAtBeginningOfPeriod,CashAndCashEquivalents")
    a(33) = Array("", "Cash at end of period", _
                  "CashCashEquivalentsRestrictedCashAndRestrictedCashEquivalents,CashAndCashEquivalentsAtCarryingValue", _
                  "USD", 1000000#, _
                  "CashAndCashEquivalents,CashAndCashEquivalentsAtEndOfPeriod")
    GetCFConcepts = a
End Function
