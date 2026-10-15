Attribute VB_Name = "modSetupLate"
Option Explicit

' RunSetup upgrade for Stages 5 to 8 (STAGES_5_TO_8.md part 0). It only ADDS: columns at the end of existing tables, missing tables, missing Config keys,
' missing GL accounts (and activates the ones the stages use), missing reference rows (products, rate tiers, fees, tax rates) and seeds for EMPTY reference tables.
' Existing data is never changed or cleared.

Private Const EFF As String = "2026-01-01"

Private Function IsoDate() As Date
    IsoDate = DateSerial(2026, 1, 1)
End Function

' ---- New tables ----

Private Function EnsureStaffTables() As Long
    Dim n As Long
    n = n + modSetup.EnsureTable(modConstants.TBL_POSITIONS, "PositionCode:T|Name:T|Category:T|MinHO:I|MinCity:I|MinTown:I|IsActive:T", "Category=Management,Operations,Lending,Support,HeadOffice;IsActive=Yes,No")
    n = n + modSetup.EnsureTable(modConstants.TBL_SALARYSCALES, "ScaleID:T|PositionCode:T|Grade:I|BaseSalary:I|EffectiveFrom:D", "")
    n = n + modSetup.EnsureTable(modConstants.TBL_ALLOWANCERULES, "RuleID:T|AllowanceCode:T|AppliesTo:T|AppliesToValue:T|CalcMethod:T|Amount:I|Rate:R|Taxable:T|EffectiveFrom:D", "AppliesTo=All,BranchType,Position;CalcMethod=Flat,PercentOfBase;Taxable=Yes,No")
    n = n + modSetup.EnsureTable(modConstants.TBL_TAXBRACKETS, "BracketID:T|FromAnnual:I|ToAnnual:I|Rate:R|EffectiveFrom:D", "")
    n = n + modSetup.EnsureTable(modConstants.TBL_EMPLOYEES, "EmployeeID:T|Name:T|PositionCode:T|Grade:I|BranchID:T|HireDate:D|ExitDate:D|Status:T|PayMode:T|BankAccountNo:T|Dependents:I|UserID:T|Phone:T", _
        "Status=Active,OnLeave,Resigned,Terminated;PayMode=Account,External")
    n = n + modSetup.EnsureTable(modConstants.TBL_EMPHISTORY, "HistoryID:T|EmployeeID:T|ChangeDate:D|ChangeType:T|FromBranch:T|ToBranch:T|FromPosition:T|ToPosition:T|FromGrade:I|ToGrade:I|Note:T", "")
    n = n + modSetup.EnsureTable(modConstants.TBL_ATTENDANCE, "AttendanceID:T|EmployeeID:T|Period:T|UnpaidLeaveDays:X|OvertimeHours:X|HolidayOvertimeHours:X|Adjustment:I|AdjustmentNote:T", "")
    n = n + modSetup.EnsureTable(modConstants.TBL_ADVANCES, "AdvanceID:T|EmployeeID:T|IssueDate:D|Amount:I|MonthlyRecovery:I|Outstanding:I|Status:T|JournalID:T", "Status=Active,Cleared")
    n = n + modSetup.EnsureTable(modConstants.TBL_PAYROLLRUNS, "PayrollRunID:T|Period:T|BranchID:T|PreparedDate:D|PreparedBy:T|ApprovedBy:T|EmployeesPaid:I|TotalGross:I|TotalEmployeeSSB:I|TotalTax:I|TotalAdvanceRecovered:I|TotalNet:I|TotalEmployerSSB:I|Status:T|PostJournalID:T|PayJournalID:T", _
        "Status=Draft,Posted,Paid,Reversed")
    n = n + modSetup.EnsureTable(modConstants.TBL_PAYSLIPS, "PayslipID:T|PayrollRunID:T|EmployeeID:T|BranchID:T|Period:T|BaseSalary:I|UnpaidDeduction:I|TaxableAllowances:I|NonTaxableAllowances:I|OvertimePay:I|Adjustment:I|Gross:I|EmployeeSSB:I|TaxableIncome:I|IncomeTax:I|AdvanceRecovered:I|NetPay:I|EmployerSSB:I|PayMode:T|Status:T", _
        "Status=Draft,Posted,Paid,Reversed")
    EnsureStaffTables = n
End Function

Private Function EnsureFxTables() As Long
    Dim n As Long
    n = n + modSetup.EnsureTable(modConstants.TBL_CURRENCIES, "Code:T|Name:T|Decimals:I|IsBase:T", "IsBase=Yes,No")
    n = n + modSetup.EnsureTable(modConstants.TBL_EXCHANGERATES, "RateDate:D|Currency:T|BuyRate:M|SellRate:M|MidRate:M", "")
    n = n + modSetup.EnsureTable(modConstants.TBL_FOREIGNBANKS, "BankCode:T|Name:T|Country:T|Currency:T|SettleDays:I|FailRate:R|ReturnRate:R", "")
    n = n + modSetup.EnsureTable(modConstants.TBL_BLOCKEDNAMES, "NameID:T|Name:T|Reason:T|AddedDate:D|Active:T", "Active=Yes,No")
    n = n + modSetup.EnsureTable(modConstants.TBL_WALLETLIMITS, "KYCLevel:I|MaxBalance:I|DailySpend:I|MonthlyVolume:I|SingleTxnMax:I", "")
    n = n + modSetup.EnsureTable(modConstants.TBL_FXDEALS, "DealID:T|DealDate:D|DealType:T|Currency:T|ForeignAmount:X|Rate:M|MidRate:M|MmkAmount:I|Spread:I|BranchID:T|AccountNo:T|CustomerID:T|Purpose:T|JournalID:T|Status:T|CreatedBy:T|ApprovedBy:T", _
        "DealType=CounterBuy,CounterSell,Convert,MarketBuy,MarketSell;Status=Posted,Reversed")
    n = n + modSetup.EnsureTable(modConstants.TBL_FXPOSITION, "PositionDate:D|Currency:T|Holdings:X|Liabilities:X|NetOpen:X|MidRate:M|NetOpenBase:I|LimitBase:I|Status:T", "Status=Within,Breach")
    n = n + modSetup.EnsureTable(modConstants.TBL_INTLPAYMENTS, "IntlPaymentID:T|Direction:T|AccountNo:T|Currency:T|Amount:X|MmkEquivalent:I|PartyName:T|ForeignBankCode:T|PartyAccount:T|Country:T|Purpose:T|Fee:X|MessageRef:T|SubmittedDate:D|ExpectedDate:D|Status:T|StatusDate:D|Reason:T|Screening:T|SubmitJournalID:T|SettleJournalID:T", _
        "Direction=Out,In;Status=Submitted,Delivered,Failed,Returned,Blocked;Screening=Clear,Hit")
    n = n + modSetup.EnsureTable(modConstants.TBL_MERCHANTS, "MerchantID:T|Name:T|Category:T|WalletAccountNo:T|QRCode:T|BranchID:T|FeeCode:T|Status:T", "Status=Active,Suspended")
    n = n + modSetup.EnsureTable(modConstants.TBL_AGENTS, "AgentID:T|Name:T|WalletAccountNo:T|BranchID:T|CashOnHand:I|Status:T", "Status=Active,Suspended")
    EnsureFxTables = n
End Function

Private Function EnsureInvestTables() As Long
    Dim n As Long
    n = n + modSetup.EnsureTable(modConstants.TBL_INVLIMITS, "LimitID:T|LimitType:T|LimitKey:T|MaxPctOfCapital:R", "LimitType=Total,Single,Sector,HighRisk")
    n = n + modSetup.EnsureTable(modConstants.TBL_INVESTEES, "InvesteeID:T|Name:T|Sector:T|RiskGrade:T|Status:T|Note:T", "RiskGrade=A,B,C,D,E;Status=Active,Distressed,Closed")
    n = n + modSetup.EnsureTable(modConstants.TBL_INVESTMENTS, "InvestmentID:T|InvesteeID:T|InvestmentType:T|Amount:I|StartDate:D|TermMonths:I|MaturityDate:D|RateOrShare:R|PayFrequency:T|Status:T|CarryingValue:I|AccruedReturn:M|LastAccrualDate:D|RiskGrade:T|Sector:T|ProposedBy:T|ApprovedBy:T|Note:T", _
        "InvestmentType=Equity,ProfitShare,FixedReturnNote,TradeFinance;PayFrequency=Monthly,Quarterly,AtMaturity;Status=Proposed,Approved,Rejected,Active,Matured,Disposed,WrittenOff")
    n = n + modSetup.EnsureTable(modConstants.TBL_INVRETURNS, "ReturnID:T|InvestmentID:T|ReturnDate:D|Amount:I|ReturnType:T|JournalID:T|Note:T", "ReturnType=Interest,Dividend,ProfitShare,Principal")
    n = n + modSetup.EnsureTable(modConstants.TBL_INVVALUATIONS, "ValuationID:T|InvestmentID:T|ValuationDate:D|Value:I|Method:T|Note:T|JournalID:T", "Method=Cost,Appraisal,Multiple")
    EnsureInvestTables = n
End Function

Private Function EnsureGovTables() As Long
    Dim n As Long
    n = n + modSetup.EnsureTable(modConstants.TBL_RATERULES, "RuleID:T|Target:T|ProductCode:T|FollowsPolicy:T|MinRate:R|MaxRate:R", "Target=Deposit,Loan;FollowsPolicy=Yes,No")
    n = n + modSetup.EnsureTable(modConstants.TBL_RISKWEIGHTS, "WeightID:T|AssetClass:T|GLCode:T|Weight:X", "")
    n = n + modSetup.EnsureTable(modConstants.TBL_RUNOFF, "DepositType:T|RunoffPct:R", "")
    n = n + modSetup.EnsureTable(modConstants.TBL_AMLRULES, "RuleID:T|RuleType:T|Threshold:I|WindowDays:I|MinCount:I|Active:T", "RuleType=LargeCash,Structuring,Velocity,BlockedName;Active=Yes,No")
    n = n + modSetup.EnsureTable(modConstants.TBL_TAXADJ, "AdjustmentID:T|Year:X|Description:T|Amount:I|Note:T", "")
    n = n + modSetup.EnsureTable(modConstants.TBL_BONDS, "BondID:T|BondType:T|Name:T|FaceValue:I|IssueDate:D|MaturityDate:D|Rate:R|CouponMonths:I|PurchasePrice:I|CarryingValue:I|AccruedCoupon:M|LastAccrualDate:D|Status:T", "BondType=TBill,GovBond;Status=Held,Matured,Sold")
    n = n + modSetup.EnsureTable(modConstants.TBL_AMLALERTS, "AlertID:T|AlertDate:D|RuleID:T|CustomerID:T|AccountNo:T|Amount:I|Description:T|Status:T|CaseID:T|ReviewedBy:T|ReviewedDate:D", "Status=New,Reviewing,ClosedNoAction,Escalated")
    n = n + modSetup.EnsureTable(modConstants.TBL_AMLCASES, "CaseID:T|OpenDate:D|CustomerID:T|Summary:T|Status:T|AssignedTo:T|Decision:T|ClosedDate:D|Notes:T", "Status=Open,Reported,Closed")
    n = n + modSetup.EnsureTable(modConstants.TBL_RATIOHISTORY, "RatioDate:D|RatioCode:T|Actual:M|Limit:M|Status:T", "Status=Within,Breach")
    n = n + modSetup.EnsureTable(modConstants.TBL_COMPLIANCE, "EventID:T|EventDate:D|RatioCode:T|Actual:M|Limit:M|Status:T|Note:T", "Status=Breach,Cleared")
    n = n + modSetup.EnsureTable(modConstants.TBL_FINES, "FineID:T|FineDate:D|Reason:T|Amount:I|IssuedBy:T|ApprovedBy:T|Status:T|JournalID:T", "Status=Proposed,Posted,Cancelled")
    n = n + modSetup.EnsureTable(modConstants.TBL_INSPECTIONS, "InspectionID:T|InspectionDate:D|RunBy:T|Result:T|FindingCount:I|Summary:T", "Result=Pass,Findings")
    EnsureGovTables = n
End Function

' ---- Config keys ----

Private Function EnsureKeys() As Long
    Dim defs As Variant, i As Long, p As Long, k As String
    defs = Array("WorkDaysPerMonth=26", "HoursPerDay=8", "OvertimeMultiplier=1.5", "HolidayOvertimeMultiplier=2.0", "EmployeeSSBRate=0.02", "EmployerSSBRate=0.03", "SSBMonthlyCap=300000", _
        "DependentReliefAnnual=1000000", "FxSpread=0.01", "MarketSpread=0.002", "AutoRates=Yes", "RateMoveMax=0.004", "FxLargeLimit=10000000", "FxPositionLimitPct=0.10", _
        "IntlDailyLimit=50000000", "IntlKyc3Amount=20000000", "AgentCommissionIn=0.001", "AgentCommissionOut=0.003", "InvestLiquidityBuffer=500000000", "InvestAutoRepay=Yes", _
        "MinCapitalRatio=0.10", "MinLiquidityRatio=1.00", "MaxNplRatio=0.05", "DepositInsuranceRate=0.001", "DepositInsuranceCap=10000000", "AmlReviewDays=7")
    For i = LBound(defs) To UBound(defs)
        p = InStr(defs(i), "=")
        k = Left$(defs(i), p - 1)
        If Not repoConfig.Has(k) Then
            repoConfig.SetValue k, Mid$(defs(i), p + 1)
            EnsureKeys = EnsureKeys + 1
        End If
    Next i
    If Not repoConfig.Has(modConstants.CFG_LAST_CLOSED) Then
        repoConfig.SetValue modConstants.CFG_LAST_CLOSED, ""
        EnsureKeys = EnsureKeys + 1
    End If
    ' The rate applied so far starts at today's policy rate, so the first RunApplyPolicyRate only follows later changes.
    If Not repoConfig.Has(modConstants.CFG_POLICY_APPLIED) Then
        repoConfig.SetValue modConstants.CFG_POLICY_APPLIED, repoConfig.GetText(modConstants.CFG_POLICYRATE, "0.10")
        EnsureKeys = EnsureKeys + 1
    End If
End Function

' ---- Chart of accounts ----

Private Function EnsureChart() As Long
    Dim data As Variant, i As Long, defs As Collection, d As Variant, found As Object, changed As Long, newRows As Collection
    If Not modSheetUtils.TableExists(modConstants.TBL_CHART) Then Exit Function
    Set defs = New Collection
    defs.Add Array("1600", "Investments", "Asset", "Debit", 7)
    defs.Add Array("1610", "Investment income receivable", "Asset", "Debit", 7)
    defs.Add Array("1620", "Government securities", "Asset", "Debit", 8)
    defs.Add Array("1700", "Foreign currency holdings", "Asset", "Debit", 6)
    defs.Add Array("1110", "E-money trust account", "Asset", "Debit", 6)
    defs.Add Array("1810", "Staff advances", "Asset", "Debit", 5)
    defs.Add Array("2030", "E-wallet balances", "Liability", "Credit", 6)
    defs.Add Array("2220", "Income tax payable", "Liability", "Credit", 8)
    defs.Add Array("2230", "Payroll tax and social security payable", "Liability", "Credit", 5)
    defs.Add Array("2400", "Salaries payable", "Liability", "Credit", 5)
    defs.Add Array("3100", "Retained earnings", "Equity", "Credit", 1)
    defs.Add Array("4200", "Foreign exchange income", "Income", "Credit", 6)
    defs.Add Array("4300", "Investment income", "Income", "Credit", 7)
    defs.Add Array("4310", "Interest income on government securities", "Income", "Credit", 8)
    defs.Add Array("4400", "Gain on disposal", "Income", "Credit", 2)
    defs.Add Array("5100", "Salaries", "Expense", "Debit", 5)
    defs.Add Array("5110", "Employer social security", "Expense", "Debit", 5)
    defs.Add Array("5290", "Agent commissions", "Expense", "Debit", 6)
    defs.Add Array("5310", "Loss on disposal", "Expense", "Debit", 2)
    defs.Add Array("5330", "Investment impairment and losses", "Expense", "Debit", 7)
    defs.Add Array("5340", "Deposit insurance premium", "Expense", "Debit", 8)
    defs.Add Array("5350", "Regulatory fines", "Expense", "Debit", 8)
    defs.Add Array("5400", "Income tax expense", "Expense", "Debit", 8)
    data = modSheetUtils.ReadTable(modConstants.TBL_CHART)
    Set found = CreateObject("Scripting.Dictionary")
    If Not IsEmpty(data) Then
        For i = 1 To UBound(data, 1)
            found(CStr(data(i, 1))) = i
        Next i
        For Each d In defs
            If found.Exists(CStr(d(0))) Then
                If CStr(data(found(CStr(d(0))), 6)) <> modConstants.FLAG_YES Then
                    data(found(CStr(d(0))), 6) = modConstants.FLAG_YES
                    changed = changed + 1
                End If
            End If
        Next d
        If changed > 0 Then modSheetUtils.WriteTable modConstants.TBL_CHART, data
    End If
    Set newRows = New Collection
    For Each d In defs
        If Not found.Exists(CStr(d(0))) Then newRows.Add Array(CStr(d(0)), CStr(d(1)), CStr(d(2)), CStr(d(3)), CLng(d(4)), modConstants.FLAG_YES)
    Next d
    If newRows.Count > 0 Then modSheetUtils.AppendRows modConstants.TBL_CHART, modSheetUtils.RowsTo2D(newRows)
    EnsureChart = changed + newRows.Count
End Function

' ---- Rows added to existing reference tables (only the codes that are missing) ----

Private Function HasCode(ByVal data As Variant, ByVal code As String) As Boolean
    Dim i As Long
    If IsEmpty(data) Then Exit Function
    For i = 1 To UBound(data, 1)
        If StrComp(CStr(data(i, 1)), code, vbTextCompare) = 0 Then HasCode = True: Exit Function
    Next i
End Function

Private Function EnsureTaxRates() As Long
    Dim data As Variant, newRows As Collection
    If Not modSheetUtils.TableExists(modConstants.TBL_TAXRATES) Then Exit Function
    data = modSheetUtils.ReadTable(modConstants.TBL_TAXRATES)
    Set newRows = New Collection
    If Not HasCode(data, "EMPLOYEE_SSB") Then newRows.Add Array("EMPLOYEE_SSB", "Employee social security", 0.02, "Payroll", "2230", IsoDate())
    If Not HasCode(data, modConstants.TAX_PAYROLL_SSB) And Not HasCode(data, "PAYROLL_TAX") Then newRows.Add Array("PAYROLL_TAX", "Payroll income tax (brackets)", Empty, "Payroll", "2230", IsoDate())
    If Not HasCode(data, modConstants.TAX_INCOME) Then newRows.Add Array(modConstants.TAX_INCOME, "Income tax on profit", 0.25, "Profit", "2220", IsoDate())
    If newRows.Count > 0 Then modSheetUtils.AppendRows modConstants.TBL_TAXRATES, modSheetUtils.RowsTo2D(newRows)
    EnsureTaxRates = newRows.Count
End Function

Private Function EnsureFees() As Long
    Dim data As Variant, newRows As Collection
    If Not modSheetUtils.TableExists(modConstants.TBL_FEES) Then Exit Function
    data = modSheetUtils.ReadTable(modConstants.TBL_FEES)
    Set newRows = New Collection
    ' FeeCode, Name, Method, FlatAmount, Rate, MinFee, MaxFee, Currency, IncomeGL, Taxable
    If Not HasCode(data, modConstants.FEE_INTL_OUT) Then newRows.Add Array(modConstants.FEE_INTL_OUT, "International payment out", "Percent", Empty, 0.005, 15000, 150000, "MMK", "4100", "Yes")
    If Not HasCode(data, modConstants.FEE_INTL_MSG) Then newRows.Add Array(modConstants.FEE_INTL_MSG, "International payment message", "Flat", 20000, Empty, Empty, Empty, "MMK", "4100", "Yes")
    If Not HasCode(data, modConstants.FEE_INTL_IN) Then newRows.Add Array(modConstants.FEE_INTL_IN, "International payment in", "Flat", 5000, Empty, Empty, Empty, "MMK", "4100", "Yes")
    If Not HasCode(data, modConstants.FEE_WLT_CASHOUT) Then newRows.Add Array(modConstants.FEE_WLT_CASHOUT, "Wallet cash-out", "Percent", Empty, 0.005, 200, 5000, "MMK", "4100", "Yes")
    If Not HasCode(data, modConstants.FEE_WLT_P2P) Then newRows.Add Array(modConstants.FEE_WLT_P2P, "Wallet person to person", "Flat", 0, Empty, Empty, Empty, "MMK", "4100", "No")
    If Not HasCode(data, modConstants.FEE_MERCHANT_MDR) Then newRows.Add Array(modConstants.FEE_MERCHANT_MDR, "Merchant discount rate", "Percent", Empty, 0.01, 0, 0, "MMK", "4100", "Yes")
    If newRows.Count > 0 Then modSheetUtils.AppendRows modConstants.TBL_FEES, modSheetUtils.RowsTo2D(newRows)
    EnsureFees = newRows.Count
End Function

Private Function NextRateNumber(ByVal tiers As Variant) As Long
    Dim i As Long, v As Long, s As String
    If IsEmpty(tiers) Then NextRateNumber = 1: Exit Function
    For i = 1 To UBound(tiers, 1)
        s = CStr(tiers(i, 1))
        If Len(s) > 1 And IsNumeric(Mid$(s, 2)) Then
            If CLng(Mid$(s, 2)) > v Then v = CLng(Mid$(s, 2))
        End If
    Next i
    NextRateNumber = v + 1
End Function

Private Function EnsureProducts() As Long
    Dim prods As Variant, tiers As Variant, newP As Collection, newT As Collection, d As Variant, defs As Collection, k As Long, i As Long, hasTier As Boolean
    If Not modSheetUtils.TableExists(modConstants.TBL_PRODUCTS) Or Not modSheetUtils.TableExists(modConstants.TBL_RATETIERS) Then Exit Function
    prods = modSheetUtils.ReadTable(modConstants.TBL_PRODUCTS)
    tiers = modSheetUtils.ReadTable(modConstants.TBL_RATETIERS)
    Set defs = New Collection
    ' ProductCode, Name, Type, Currency, MinBalance, MinOpening, tier rate
    defs.Add Array("USD01", "USD Savings", "Savings", "USD", 10, 50, 0.01)
    defs.Add Array("USD02", "USD Current", "Current", "USD", 100, 100, 0)
    defs.Add Array("JPY01", "JPY Savings", "Savings", "JPY", 1000, 5000, 0.001)
    defs.Add Array("THB01", "THB Savings", "Savings", "THB", 300, 1000, 0.015)
    defs.Add Array("SGD01", "SGD Savings", "Savings", "SGD", 20, 100, 0.01)
    defs.Add Array("WLT01", "E-Wallet", "Wallet", "MMK", 0, 0, 0)
    Set newP = New Collection
    Set newT = New Collection
    k = NextRateNumber(tiers)
    For Each d In defs
        If Not HasCode(prods, CStr(d(0))) Then newP.Add Array(CStr(d(0)), CStr(d(1)), CStr(d(2)), CStr(d(3)), d(4), d(5), Empty, modConstants.FLAG_YES)
        hasTier = False
        If Not IsEmpty(tiers) Then
            For i = 1 To UBound(tiers, 1)
                If StrComp(CStr(tiers(i, 2)), CStr(d(0)), vbTextCompare) = 0 Then hasTier = True: Exit For
            Next i
        End If
        If Not hasTier Then
            newT.Add Array("R" & Right$("0000" & CStr(k), 4), CStr(d(0)), 0, 0, Empty, d(6), IsoDate())
            k = k + 1
        End If
    Next d
    If newP.Count > 0 Then modSheetUtils.AppendRows modConstants.TBL_PRODUCTS, modSheetUtils.RowsTo2D(newP)
    If newT.Count > 0 Then modSheetUtils.AppendRows modConstants.TBL_RATETIERS, modSheetUtils.RowsTo2D(newT)
    EnsureProducts = newP.Count + newT.Count
End Function

' ---- Seeds for empty reference tables ----

Private Sub Seed(ByVal tableName As String, ByVal rowsList As Collection, ByRef counter As Long)
    If Not modSheetUtils.TableExists(tableName) Then Exit Sub
    If Not modSetupStages.TableIsEmpty(tableName) Then Exit Sub
    If rowsList.Count = 0 Then Exit Sub
    modSheetUtils.AppendRows tableName, modSheetUtils.RowsTo2D(rowsList)
    counter = counter + rowsList.Count
End Sub

Private Function SeedStaff() As Long
    Dim n As Long, c As Collection, p As Variant, scaleRows As Variant, i As Long, g As Long, k As Long
    Set c = New Collection
    ' PositionCode, Name, Category, MinHO, MinCity, MinTown, IsActive
    c.Add Array("GM", "General Manager", "Management", 1, 0, 0, "Yes")
    c.Add Array("HOD", "Head of Department", "Management", 5, 0, 0, "Yes")
    c.Add Array("ACC", "Accountant", "HeadOffice", 6, 0, 0, "Yes")
    c.Add Array("RSK", "Risk Officer", "HeadOffice", 4, 0, 0, "Yes")
    c.Add Array("HRO", "HR Officer", "HeadOffice", 3, 0, 0, "Yes")
    c.Add Array("AUD", "Internal Auditor", "HeadOffice", 3, 0, 0, "Yes")
    c.Add Array("TRS", "Treasury Officer", "HeadOffice", 3, 0, 0, "Yes")
    c.Add Array("OPS", "Operations Officer", "HeadOffice", 6, 0, 0, "Yes")
    c.Add Array("BM", "Branch Manager", "Management", 0, 1, 1, "Yes")
    c.Add Array("ABM", "Assistant Branch Manager", "Management", 0, 1, 0, "Yes")
    c.Add Array("TLR", "Teller", "Operations", 6, 3, 2, "Yes")
    c.Add Array("CSH", "Cashier", "Operations", 2, 1, 1, "Yes")
    c.Add Array("LNO", "Loan Officer", "Lending", 4, 2, 1, "Yes")
    c.Add Array("CSR", "Customer Service Officer", "Operations", 4, 2, 1, "Yes")
    c.Add Array("SEC", "Security Guard", "Support", 8, 3, 2, "Yes")
    c.Add Array("CLN", "Cleaner", "Support", 4, 1, 1, "Yes")
    c.Add Array("DRV", "Driver", "Support", 4, 1, 0, "Yes")
    c.Add Array("ITS", "IT Support", "Support", 4, 1, 0, "Yes")
    Seed modConstants.TBL_POSITIONS, c, n
    Set c = New Collection
    scaleRows = Array("GM|5000000|6000000|7000000", "HOD|2800000|3300000|3800000", "ACC|800000|950000|1100000", "RSK|900000|1050000|1200000", "HRO|700000|820000|950000", _
        "AUD|900000|1050000|1200000", "TRS|950000|1100000|1300000", "OPS|650000|760000|880000", "BM|1600000|2000000|2500000", "ABM|1100000|1350000|1600000", _
        "TLR|450000|520000|600000", "CSH|500000|580000|660000", "LNO|650000|780000|900000", "CSR|450000|520000|600000", "SEC|300000|340000|380000", _
        "CLN|250000|280000|310000", "DRV|350000|400000|450000", "ITS|700000|850000|1000000")
    k = 0
    For i = 0 To UBound(scaleRows)
        p = Split(CStr(scaleRows(i)), "|")
        For g = 1 To 3
            k = k + 1
            c.Add Array("SS" & Right$("0000" & CStr(k), 4), CStr(p(0)), g, CLng(p(g)), IsoDate())
        Next g
    Next i
    Seed modConstants.TBL_SALARYSCALES, c, n
    Set c = New Collection
    ' RuleID, AllowanceCode, AppliesTo, AppliesToValue, CalcMethod, Amount, Rate, Taxable, EffectiveFrom
    c.Add Array("AR0001", "CITY_ALLOW", "BranchType", "City", "Flat", 60000, Empty, "Yes", IsoDate())
    c.Add Array("AR0002", "TOWN_HARDSHIP", "BranchType", "Town", "Flat", 90000, Empty, "Yes", IsoDate())
    c.Add Array("AR0003", "HO_ALLOW", "BranchType", "HeadOffice", "Flat", 40000, Empty, "Yes", IsoDate())
    c.Add Array("AR0004", "MGR_ALLOW", "Position", "BM", "PercentOfBase", Empty, 0.1, "Yes", IsoDate())
    c.Add Array("AR0005", "MGR_ALLOW", "Position", "GM", "PercentOfBase", Empty, 0.1, "Yes", IsoDate())
    c.Add Array("AR0006", "MGR_ALLOW", "Position", "HOD", "PercentOfBase", Empty, 0.1, "Yes", IsoDate())
    c.Add Array("AR0007", "CASH_HANDLING", "Position", "TLR", "Flat", 25000, Empty, "Yes", IsoDate())
    c.Add Array("AR0008", "CASH_HANDLING", "Position", "CSH", "Flat", 25000, Empty, "Yes", IsoDate())
    c.Add Array("AR0009", "NIGHT_SHIFT", "Position", "SEC", "Flat", 15000, Empty, "Yes", IsoDate())
    c.Add Array("AR0010", "TRANSPORT", "All", Empty, "Flat", 30000, Empty, "No", IsoDate())
    c.Add Array("AR0011", "MEAL", "All", Empty, "Flat", 20000, Empty, "No", IsoDate())
    Seed modConstants.TBL_ALLOWANCERULES, c, n
    Set c = New Collection
    c.Add Array("TB0001", 0, 4800000, 0, IsoDate())
    c.Add Array("TB0002", 4800000, 10000000, 0.05, IsoDate())
    c.Add Array("TB0003", 10000000, 20000000, 0.1, IsoDate())
    c.Add Array("TB0004", 20000000, 30000000, 0.15, IsoDate())
    c.Add Array("TB0005", 30000000, Empty, 0.2, IsoDate())
    Seed modConstants.TBL_TAXBRACKETS, c, n
    SeedStaff = n
End Function

Private Function SeedFx() As Long
    Dim n As Long, c As Collection
    Set c = New Collection
    c.Add Array("FB01", "Pacific Union Bank (simulated)", "Singapore", "SGD", 2, 0.02, 0.01)
    c.Add Array("FB02", "Eastern Star Bank (simulated)", "Japan", "JPY", 2, 0.02, 0.01)
    c.Add Array("FB03", "Lakeside Bank (simulated)", "Thailand", "THB", 2, 0.03, 0.01)
    c.Add Array("FB04", "Liberty Trust Bank (simulated)", "United States", "USD", 1, 0.02, 0.01)
    c.Add Array("FB05", "Harbour Commerce Bank (simulated)", "Singapore", "USD", 2, 0.03, 0.02)
    Seed modConstants.TBL_FOREIGNBANKS, c, n
    Set c = New Collection
    c.Add Array("NM000001", "Black Orchid Trading", "Made-up sample entry", IsoDate(), "Yes")
    c.Add Array("NM000002", "Red Harbour Holdings", "Made-up sample entry", IsoDate(), "Yes")
    c.Add Array("NM000003", "Test Blocked Company", "Made-up sample entry", IsoDate(), "Yes")
    c.Add Array("NM000004", "Sample Sanctioned Person", "Made-up sample entry", IsoDate(), "Yes")
    c.Add Array("NM000005", "John Doe Blocked", "Made-up sample entry", IsoDate(), "Yes")
    Seed modConstants.TBL_BLOCKEDNAMES, c, n
    Set c = New Collection
    c.Add Array(1, 500000, 300000, 2000000, 200000)
    c.Add Array(2, 5000000, 2000000, 20000000, 1000000)
    c.Add Array(3, 30000000, 10000000, 100000000, 5000000)
    Seed modConstants.TBL_WALLETLIMITS, c, n
    SeedFx = n
End Function

Private Function SeedInvest() As Long
    Dim n As Long, c As Collection, s As Variant, k As Long
    Set c = New Collection
    c.Add Array("LM0001", "Total", "ALL", 0.3)
    c.Add Array("LM0002", "Single", "ALL", 0.05)
    k = 2
    For Each s In Array("Agriculture", "Manufacturing", "Trade", "Services", "RealEstate", "Energy", "Technology")
        k = k + 1
        c.Add Array("LM" & Right$("0000" & CStr(k), 4), "Sector", CStr(s), 0.1)
    Next s
    c.Add Array("LM" & Right$("0000" & CStr(k + 1), 4), "HighRisk", "D,E", 0.03)
    Seed modConstants.TBL_INVLIMITS, c, n
    SeedInvest = n
End Function

' Rate rules follow the products and loan products that exist when the setup runs.
Private Function SeedRateRules() As Long
    Dim n As Long, c As Collection, prods As Variant, lp As Variant, i As Long, k As Long
    If Not modSheetUtils.TableExists(modConstants.TBL_RATERULES) Then Exit Function
    If Not modSetupStages.TableIsEmpty(modConstants.TBL_RATERULES) Then Exit Function
    Set c = New Collection
    prods = modSheetUtils.ReadTable(modConstants.TBL_PRODUCTS)
    If Not IsEmpty(prods) Then
        For i = 1 To UBound(prods, 1)
            If Len(CStr(prods(i, 1))) > 0 And CStr(prods(i, 3)) <> modConstants.PT_WALLET Then
                k = k + 1
                c.Add Array("RR" & Right$("0000" & CStr(k), 4), "Deposit", CStr(prods(i, 1)), IIf(CStr(prods(i, 3)) = modConstants.PT_CURRENT, "No", "Yes"), 0.01, 0.15)
            End If
        Next i
    End If
    If modSheetUtils.TableExists(modConstants.TBL_LOANPRODUCTS) Then
        lp = modSheetUtils.ReadTable(modConstants.TBL_LOANPRODUCTS)
        If Not IsEmpty(lp) Then
            For i = 1 To UBound(lp, 1)
                If Len(CStr(lp(i, 1))) > 0 Then
                    k = k + 1
                    c.Add Array("RR" & Right$("0000" & CStr(k), 4), "Loan", CStr(lp(i, 1)), "Yes", 0.06, 0.3)
                End If
            Next i
        End If
    End If
    Seed modConstants.TBL_RATERULES, c, n
    SeedRateRules = n
End Function

Private Function SeedGov() As Long
    Dim n As Long, c As Collection
    Set c = New Collection
    c.Add Array("RW01", "Vault cash", "1000", 0)
    c.Add Array("RW02", "Treasury and settlement", "1100", 0.2)
    c.Add Array("RW03", "E-money trust", "1110", 0.2)
    c.Add Array("RW04", "Loan secured by home", "1300", 0.5)
    c.Add Array("RW05", "Loan secured, other", "1300", 0.75)
    c.Add Array("RW06", "Loan unsecured", "1300", 1)
    c.Add Array("RW07", "Loan past due (class Substandard or worse)", "1300", 1.5)
    c.Add Array("RW08", "Interest receivable", "1310", 1)
    c.Add Array("RW09", "Repossessed assets", "1400", 1.5)
    c.Add Array("RW10", "Premises and equipment (net)", "1500", 1)
    c.Add Array("RW11", "Investments", "1600", 1.5)
    c.Add Array("RW12", "Investment income receivable", "1610", 1)
    c.Add Array("RW13", "Government securities", "1620", 0)
    c.Add Array("RW14", "Foreign holdings", "1700", 1)
    c.Add Array("RW15", "Prepaid expenses", "1800", 1)
    c.Add Array("RW16", "Staff advances", "1810", 1)
    c.Add Array("RW17", "Inter-branch clearing", "1900", 0)
    c.Add Array("RW18", "Loan loss allowance (contra)", "1320", "counted with the loans it reduces (negative)")
    c.Add Array("RW19", "Accumulated depreciation (contra)", "1510", "counted with 1500 (negative)")
    Seed modConstants.TBL_RISKWEIGHTS, c, n
    Set c = New Collection
    c.Add Array("Savings", 0.1)
    c.Add Array("Current", 0.15)
    c.Add Array("Wallet", 0.15)
    c.Add Array("FixedMaturing30", 1)
    c.Add Array("FixedOther", 0)
    Seed modConstants.TBL_RUNOFF, c, n
    Set c = New Collection
    c.Add Array("LARGE_CASH", "LargeCash", 20000000, 0, 1, "Yes")
    c.Add Array("STRUCTURING", "Structuring", 20000000, 3, 3, "Yes")
    c.Add Array("VELOCITY", "Velocity", 100000000, 1, 1, "Yes")
    c.Add Array("BLOCKED_NAME", "BlockedName", 0, 0, 0, "Yes")
    Seed modConstants.TBL_AMLRULES, c, n
    SeedGov = n + SeedRateRules()
End Function

' Runs the whole Stage 5 to 8 upgrade. Returns the number of items added.
Public Function UpgradeLate() As Long
    Dim n As Long
    n = n + modSetupStages.AddColumns(modConstants.TBL_DAILYCLOSE, "FxRevaluation:I|IntlDecided:I|InvestmentAccrued:M|RatiosBreached:I|AmlAlerts:I")
    n = n + EnsureStaffTables()
    n = n + EnsureFxTables()
    n = n + EnsureInvestTables()
    n = n + EnsureGovTables()
    n = n + EnsureKeys()
    n = n + EnsureChart()
    n = n + EnsureTaxRates()
    n = n + EnsureFees()
    n = n + EnsureProducts()
    n = n + SeedStaff()
    n = n + SeedFx()
    n = n + SeedInvest()
    n = n + SeedGov()
    UpgradeLate = n
End Function
