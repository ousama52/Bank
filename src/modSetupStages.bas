Attribute VB_Name = "modSetupStages"
Option Explicit

' RunSetup upgrade for Stages 2 to 4 (STAGES_2_TO_4.md section 0.1). It only ADDS: columns at the end of existing tables, missing tables,
' missing Config keys, missing GL accounts, and seed rows for empty reference tables. Existing data is never changed or cleared
' (the only edit to existing rows is filling an EMPTY Taxable cell in tblFees).

' Builds a 1-based 2D array from rows given as Array(...).
Public Function Grid(ParamArray dataRows() As Variant) As Variant
    Dim out() As Variant, r As Long, c As Long, n As Long, k As Long
    n = UBound(dataRows) + 1
    k = UBound(dataRows(0)) + 1
    ReDim out(1 To n, 1 To k)
    For r = 1 To n
        For c = 1 To k
            out(r, c) = dataRows(r - 1)(c - 1)
        Next c
    Next r
    Grid = out
End Function

Public Function TableIsEmpty(ByVal tableName As String) As Boolean
    TableIsEmpty = IsEmpty(modSheetUtils.ReadTable(tableName))
End Function

' ---- Columns at the end of existing tables ----

Public Function AddColumns(ByVal tableName As String, ByVal cols As String) As Long
    Dim lo As ListObject, map As Object, parts() As String, pair() As String, i As Long, nm As String, newCol As ListColumn
    If Not modSheetUtils.TableExists(tableName) Then Exit Function
    Set lo = modSheetUtils.GetTable(tableName)
    Set map = modSheetUtils.HeaderMap(tableName)
    parts = Split(cols, "|")
    For i = 0 To UBound(parts)
        pair = Split(parts(i), ":")
        nm = pair(0)
        If Not map.Exists(nm) Then
            Set newCol = lo.ListColumns.Add
            newCol.Name = nm
            If Not newCol.DataBodyRange Is Nothing Then newCol.DataBodyRange.NumberFormat = modSetup.FormatFor(pair(1))
            AddColumns = AddColumns + 1
        End If
    Next i
End Function

' tblFees.Taxable: Yes for every fee except TRANSFER_SAME and INTERBANK_IN. Only empty cells are filled.
Private Sub SeedTaxable()
    Dim data As Variant, i As Long, changed As Boolean, c As Long
    If Not modSheetUtils.TableExists(modConstants.TBL_FEES) Then Exit Sub
    data = modSheetUtils.ReadTable(modConstants.TBL_FEES)
    If IsEmpty(data) Then Exit Sub
    c = modSheetUtils.HeaderMap(modConstants.TBL_FEES)(modConstants.FEE_TAXABLE_COL)
    For i = 1 To UBound(data, 1)
        If Len(CStr(data(i, 1))) > 0 And Len(CStr(data(i, c))) = 0 Then
            If CStr(data(i, 1)) = modConstants.FEE_TRANSFER_SAME Or CStr(data(i, 1)) = modConstants.FEE_IB_IN Then
                data(i, c) = modConstants.FLAG_NO
            Else
                data(i, c) = modConstants.FLAG_YES
            End If
            changed = True
        End If
    Next i
    If changed Then modSheetUtils.WriteTable modConstants.TBL_FEES, data
End Sub

' ---- New tables ----

Private Function EnsureNewTables() As Long
    Dim n As Long
    n = n + modSetup.EnsureTable(modConstants.TBL_COSTPOSTINGS, "PostingID:T|CostID:T|BranchID:T|Period:T|GLCode:T|Amount:I|JournalID:T|Status:T|PaidDate:D|PaidJournalID:T", "Status=Unpaid,Paid")
    n = n + modSetup.EnsureTable(modConstants.TBL_EXPENSES, "ExpenseID:T|BranchID:T|ExpenseDate:D|GLCode:T|Amount:I|Vendor:T|Description:T|PaidFrom:T|Status:T|RequestedBy:T|ApprovedBy:T|JournalID:T|PaidJournalID:T", _
        "PaidFrom=Vault,HeadOffice;Status=Pending,Posted,Paid,Rejected")
    n = n + modSetup.EnsureTable(modConstants.TBL_TAXPAYMENTS, "TaxPaymentID:T|TaxCode:T|PaidDate:D|Period:T|Amount:I|JournalID:T|Reference:T", "")
    n = n + modSetup.EnsureTable(modConstants.TBL_BUDGETS, "BudgetID:T|BranchID:T|Year:X|GLCode:T|AnnualAmount:I", "")
    n = n + modSetup.EnsureTable(modConstants.TBL_LOANPRODUCTS, "ProductCode:T|Name:T|LoanType:T|AnnualRate:R|Method:T|MinAmount:I|MaxAmount:I|MinTermMonths:I|MaxTermMonths:I|ProcessingFeePct:R|GraceDays:I|PenaltyRate:R|CollateralRequired:T|MaxLTV:R|EarlyPayoffFeePct:R|IsActive:T", _
        "LoanType=Personal,Vehicle,Business,Home,GoldItem,Agricultural;Method=Flat,Declining;CollateralRequired=Yes,No;IsActive=Yes,No")
    n = n + modSetup.EnsureTable(modConstants.TBL_LOANCLASSES, "ClassCode:T|Name:T|MinDaysPastDue:I|MaxDaysPastDue:I|ProvisionPct:R", "")
    n = n + modSetup.EnsureTable(modConstants.TBL_LOANS, "LoanID:T|CustomerID:T|BranchID:T|ProductCode:T|AccountNo:T|Principal:I|AnnualRate:R|Method:T|TermMonths:I|ApplicationDate:D|ApprovedDate:D|DisburseDate:D|FirstDueDate:D|MaturityDate:D|Status:T|PrincipalOutstanding:I|PenaltyAccrued:M|DaysPastDue:I|ClassCode:I|ProvisionAmount:I|AppliedBy:T|ApprovedBy:T|Note:T", _
        "Status=Applied,Approved,Rejected,Active,Closed,WrittenOff,Repossessed")
    n = n + modSetup.EnsureTable(modConstants.TBL_LOANSCHEDULE, "ScheduleID:T|LoanID:T|InstallmentNo:I|DueDate:D|PrincipalDue:I|InterestDue:I|TotalDue:I|PrincipalPaid:I|InterestPaid:I|PenaltyPaid:I|Status:T|PaidDate:D", "Status=Future,Due,Partial,Paid")
    n = n + modSetup.EnsureTable(modConstants.TBL_LOANPAYMENTS, "LoanPaymentID:T|LoanID:T|PaymentDate:D|Amount:I|PenaltyPart:I|InterestPart:I|PrincipalPart:I|Source:T|PaymentType:T|JournalID:T|GroupID:T|CreatedBy:T", _
        "Source=Account,Cash,AutoDebit;PaymentType=Installment,Extra,Payoff,Recovery")
    n = n + modSetup.EnsureTable(modConstants.TBL_COLLTYPES, "TypeCode:T|Name:T|MaxLTV:R|Haircut:R|AnnualValueChange:R|RevalueMonths:I|SellingCostPct:R|MonthlyHoldingCostPct:R", "")
    n = n + modSetup.EnsureTable(modConstants.TBL_COLLATERAL, "CollateralID:T|CustomerID:T|TypeCode:T|Description:T|Location:T|BranchID:T|Custody:T|LoanID:T|LienStatus:T|RegisteredDate:D|LatestValue:I|LatestValueDate:D", _
        "Custody=CustomerKeeps,BankHolds;LienStatus=Free,Pledged,Released,Repossessed,Sold,Redeemed")
    n = n + modSetup.EnsureTable(modConstants.TBL_APPRAISALS, "AppraisalID:T|CollateralID:T|AppraisalDate:D|AppraisedValue:I|Appraiser:T|Purpose:T|Note:T", "Purpose=Pledge,Periodic,Repossession,Sale")
    n = n + modSetup.EnsureTable(modConstants.TBL_REPOSSESSIONS, "RepoID:T|LoanID:T|CollateralID:T|HeldBranchID:T|NoticeDate:D|RepoDate:D|PrincipalAtRepo:I|InterestAtRepo:I|ProvisionUsed:I|ClaimTotal:I|FairValueLessCosts:I|CarryingValue:I|OriginalCarryingValue:I|Deficiency:I|Status:T|ClosedDate:D", _
        "Status=Notice,Held,Sold,Rented,InBusiness,Redeemed,Closed")
    n = n + modSetup.EnsureTable(modConstants.TBL_ASSETSALES, "SaleID:T|RepoID:T|SaleDate:D|Buyer:T|SalePrice:I|SellingCosts:I|NetProceeds:I|AppliedToClaim:I|SurplusToCustomer:I|CarryingValueAtSale:I|RecoveryOrLoss:I|JournalID:T", "")
    n = n + modSetup.EnsureTable(modConstants.TBL_ASSETRENTALS, "RentalID:T|RepoID:T|Tenant:T|MonthlyRent:I|StartDate:D|EndDate:D|Status:T", "Status=Active,Ended")
    n = n + modSetup.EnsureTable(modConstants.TBL_ASSETOPS, "OpID:T|RepoID:T|OpDate:D|OpType:T|Amount:I|Note:T|JournalID:T", "OpType=Rent,Revenue,Cost")
    EnsureNewTables = n
End Function

' ---- Config keys ----

Private Function EnsureKeys() As Long
    Dim defs As Variant, i As Long, p As Long, k As String
    defs = Array("ExpenseApprovalLimit=2000000", "PettyCashLimit=500000", "BudgetUplift=0.05", _
        "LoanBranchLimit=30000000", "LoanMinKYC=2", "LoanKyc3Amount=50000000", "MaxLoanToDeposit=0.80", "SingleBorrowerLimitPct=0.05", "LoanAutoDebit=Yes", "RequireCollateralCheck=No", _
        "RepoNoticeDays=90", "RepoNoticePeriodDays=30", "MaxHoldingMonths=24", "LtvAlertMargin=0.10", "AppraisalValidDays=30", "SaleDiscountLimit=0.10", "HeldAssetRevalueMonths=3", _
        "AppraisalSecondOpinionLimit=100000000")
    For i = LBound(defs) To UBound(defs)
        p = InStr(defs(i), "=")
        k = Left$(defs(i), p - 1)
        If Not repoConfig.Has(k) Then
            repoConfig.SetValue k, Mid$(defs(i), p + 1)
            EnsureKeys = EnsureKeys + 1
        End If
    Next i
End Function

' ---- Chart of accounts ----

' Activates the accounts Stages 2 to 4 use and adds the four new ones (1320, 4020, 4030, 5320).
Private Function EnsureChart() As Long
    Dim data As Variant, i As Long, act As Variant, g As Variant, found As Object, changed As Long, newRows As Collection, rows As Variant, c As Long, r As Long
    If Not modSheetUtils.TableExists(modConstants.TBL_CHART) Then Exit Function
    act = Array("1300", "1310", "1320", "1400", "1500", "1510", "2210", "2300", "4000", "4020", "4030", "4400", "4500", "5200", "5210", "5211", "5212", "5220", "5230", "5240", _
                "5250", "5260", "5270", "5280", "5300", "5310", "5320", "5500")
    data = modSheetUtils.ReadTable(modConstants.TBL_CHART)
    Set found = CreateObject("Scripting.Dictionary")
    If Not IsEmpty(data) Then
        For i = 1 To UBound(data, 1)
            found(CStr(data(i, 1))) = i
            For Each g In act
                If CStr(data(i, 1)) = CStr(g) And CStr(data(i, 6)) <> modConstants.FLAG_YES Then
                    data(i, 6) = modConstants.FLAG_YES
                    changed = changed + 1
                End If
            Next g
        Next i
        If changed > 0 Then modSheetUtils.WriteTable modConstants.TBL_CHART, data
    End If
    Set newRows = New Collection
    If Not found.Exists("1320") Then newRows.Add Array("1320", "Loan loss allowance", "Asset", "Credit", 3, "Yes")
    If Not found.Exists("4020") Then newRows.Add Array("4020", "Penalty income on loans", "Income", "Credit", 3, "Yes")
    If Not found.Exists("4030") Then newRows.Add Array("4030", "Recoveries on loans and repossessed assets", "Income", "Credit", 3, "Yes")
    If Not found.Exists("5320") Then newRows.Add Array("5320", "Impairment of repossessed assets", "Expense", "Debit", 4, "Yes")
    If newRows.Count > 0 Then modSheetUtils.AppendRows modConstants.TBL_CHART, modSheetUtils.RowsTo2D(newRows)
    EnsureChart = changed + newRows.Count
End Function

' ---- Seeds for empty reference tables ----

Private Function SeedTables() As Long
    If modSheetUtils.TableExists(modConstants.TBL_LOANPRODUCTS) Then
        If TableIsEmpty(modConstants.TBL_LOANPRODUCTS) Then
            modSheetUtils.AppendRows modConstants.TBL_LOANPRODUCTS, Grid( _
                Array("LN01", "Personal Loan", "Personal", 0.18, "Flat", 1000000, 20000000, 6, 36, 0.01, 5, 0.24, "No", 0, 0.01, "Yes"), _
                Array("LN02", "Vehicle Loan", "Vehicle", 0.14, "Declining", 10000000, 150000000, 12, 60, 0.01, 5, 0.2, "Yes", 0.7, 0.01, "Yes"), _
                Array("LN03", "Business Working Capital", "Business", 0.16, "Declining", 5000000, 200000000, 6, 36, 0.015, 5, 0.22, "Yes", 0.6, 0.01, "Yes"), _
                Array("LN04", "Home Loan", "Home", 0.12, "Declining", 20000000, 500000000, 60, 240, 0.01, 10, 0.18, "Yes", 0.7, 0.005, "Yes"), _
                Array("LN05", "Gold Loan", "GoldItem", 0.15, "Declining", 500000, 50000000, 3, 12, 0.005, 3, 0.24, "Yes", 0.75, 0, "Yes"), _
                Array("LN06", "Agricultural Loan", "Agricultural", 0.1, "Declining", 3000000, 100000000, 6, 24, 0.005, 15, 0.15, "Yes", 0.6, 0, "Yes"))
            SeedTables = SeedTables + 6
        End If
    End If
    If modSheetUtils.TableExists(modConstants.TBL_LOANCLASSES) Then
        If TableIsEmpty(modConstants.TBL_LOANCLASSES) Then
            modSheetUtils.AppendRows modConstants.TBL_LOANCLASSES, Grid( _
                Array("1", "Standard", 0, 29, 0.01), Array("2", "Watch", 30, 89, 0.05), Array("3", "Substandard", 90, 179, 0.25), _
                Array("4", "Doubtful", 180, 359, 0.5), Array("5", "Loss", 360, Empty, 1))
            SeedTables = SeedTables + 5
        End If
    End If
    If modSheetUtils.TableExists(modConstants.TBL_COLLTYPES) Then
        If TableIsEmpty(modConstants.TBL_COLLTYPES) Then
            modSheetUtils.AppendRows modConstants.TBL_COLLTYPES, Grid( _
                Array("LAND", "Land", 0.6, 0.2, 0.05, 12, 0.05, 0.001), _
                Array("HOUSE", "House or building", 0.7, 0.15, 0.03, 12, 0.03, 0.002), _
                Array("VEHICLE", "Vehicle", 0.7, 0.25, -0.12, 6, 0.04, 0.005), _
                Array("GOLD", "Gold", 0.75, 0.05, 0, 3, 0.01, 0.0005), _
                Array("MACHINERY", "Machinery", 0.5, 0.3, -0.1, 12, 0.06, 0.004), _
                Array("STOCK", "Stock and goods", 0.4, 0.35, 0, 3, 0.08, 0.006), _
                Array("EQUIPMENT", "Equipment", 0.5, 0.3, -0.15, 12, 0.06, 0.004))
            SeedTables = SeedTables + 7
        End If
    End If
End Function

' Runs the whole Stage 2 to 4 upgrade. Returns the number of items added.
Public Function UpgradeAll() As Long
    Dim n As Long
    n = n + AddColumns(modConstants.TBL_FIXEDASSETS, "AccumDepreciation:I|LastDepPeriod:T|LoadedOn:D|DisposedOn:D|DisposalProceeds:I")
    n = n + AddColumns(modConstants.TBL_FEES, modConstants.FEE_TAXABLE_COL & ":T")
    n = n + AddColumns(modConstants.TBL_DAILYCLOSE, "CostsPosted:I|DepreciationPosted:I|InstallmentsDue:I|AutoDebited:I|PenaltyAccrued:M|ProvisionPosted:I")
    SeedTaxable
    n = n + EnsureNewTables()
    n = n + EnsureKeys()
    n = n + EnsureChart()
    n = n + SeedTables()
    UpgradeAll = n
End Function
