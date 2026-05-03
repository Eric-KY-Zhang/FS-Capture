Attribute VB_Name = "模块_抓美股利润表"
Option Explicit

' =================================================================
'  抓美股 Income Statement — Phase 4b-2
'  IS 是 duration entry: Q2/Q3 在 EDGAR 同时有 3-month 和 YTD 两种 start
'  RunUSStatement 的 canonical 选 max(end) + min(start), 自动偏好 YTD
'  EPS 单位 USD/share, 用 5-tuple (concept, unit, scale=1)
' =================================================================

Public Sub Main()
    RunUSStatement "Income", "美股_利润表", GetISConcepts(), 6
End Sub


Private Function GetISConcepts() As Variant
    Dim a(0 To 13) As Variant
    a(0) = Array("", "Revenue", _
                 "RevenueFromContractWithCustomerExcludingAssessedTax,Revenues,SalesRevenueNet,RevenueFromContractWithCustomerIncludingAssessedTax,SalesRevenueGoodsNet,SalesRevenueServicesNet", _
                 "USD", 1000000#, _
                 "Revenue,RevenueFromContractsWithCustomers")
    a(1) = Array("", "Cost of goods & services sold", _
                 "CostOfGoodsAndServicesSold,CostOfRevenue,CostOfGoodsSold", _
                 "USD", 1000000#, _
                 "CostOfSales")
    a(2) = Array("", "Gross profit", _
                 "GrossProfit", _
                 "USD", 1000000#, _
                 "GrossProfit")
    a(3) = Array("Operating Expenses", "R&D expense", _
                 "ResearchAndDevelopmentExpense,ResearchAndDevelopmentExpenseExcludingAcquiredInProcessCost")
    a(4) = Array("Operating Expenses", "SG&A expense", _
                 "SellingGeneralAndAdministrativeExpense,SellingGeneralAndAdministrativeExpenseExcludingResearchAndDevelopment")
    a(5) = Array("Operating Expenses", "Total operating expenses", _
                 "OperatingExpenses")
    a(6) = Array("", "Operating income", _
                 "OperatingIncomeLoss", _
                 "USD", 1000000#, _
                 "OperatingProfitLoss,ProfitLossFromOperatingActivities")
    a(7) = Array("", "Non-operating income / (expense)", _
                 "NonoperatingIncomeExpense,OtherNonoperatingIncomeExpense,OtherIncomeExpenseNet")
    a(8) = Array("", "Interest expense", _
                 "InterestExpense,InterestExpenseNonOperating,InterestAndDebtExpense", _
                 "USD", 1000000#, _
                 "FinanceCosts,InterestExpense")
    a(9) = Array("", "Pre-tax income", _
                 "IncomeLossFromContinuingOperationsBeforeIncomeTaxesExtraordinaryItemsNoncontrollingInterest,IncomeLossFromContinuingOperationsBeforeIncomeTaxesMinorityInterestAndIncomeLossFromEquityMethodInvestments,IncomeLossFromContinuingOperationsBeforeIncomeTaxes", _
                 "USD", 1000000#, _
                 "ProfitLossBeforeTax,ProfitLossBeforeTaxFromContinuingOperations")
    a(10) = Array("", "Income tax expense", _
                  "IncomeTaxExpenseBenefit,IncomeTaxExpenseBenefitContinuingOperations", _
                  "USD", 1000000#, _
                  "IncomeTaxExpenseContinuingOperations,TaxExpense")
    a(11) = Array("", "Net income", _
                  "NetIncomeLoss,ProfitLoss", _
                  "USD", 1000000#, _
                  "ProfitLoss,ProfitLossAttributableToOwnersOfParent")
    a(12) = Array("Per Share", "Basic EPS (USD/share)", _
                  "EarningsPerShareBasic,EarningsPerShareBasicAndDiluted", _
                  "USD/shares", 1#, _
                  "BasicEarningsLossPerShare")
    a(13) = Array("Per Share", "Diluted EPS (USD/share)", _
                  "EarningsPerShareDiluted,EarningsPerShareBasicAndDiluted", _
                  "USD/shares", 1#, _
                  "DilutedEarningsLossPerShare")
    GetISConcepts = a
End Function
