Attribute VB_Name = "svcReportsCollateral"
Option Explicit

' Stage 4 reports: CollateralRegister, CollateralAlerts, RepossessedAssets, SaleResults, AssetPnL, Deficiencies.

Private Const SH_REGISTER As String = "CollateralRegister"
Private Const SH_ALERTS As String = "CollateralAlerts"
Private Const SH_REPOSSESSED As String = "RepossessedAssets"
Private Const SH_SALES As String = "SaleResults"
Private Const SH_PNL As String = "AssetPnL"
Private Const SH_DEFICIENCY As String = "Deficiencies"

Private Sub AddTo(ByVal d As Object, ByVal key As String, ByVal v As Variant)
    If d.Exists(key) Then d(key) = d(key) + v Else d(key) = v
End Sub

Private Function Num(ByVal d As Object, ByVal key As String) As Variant
    If d.Exists(key) Then Num = d(key) Else Num = CDec(0)
End Function

' Total latest value of the items pledged to each loan.
Private Function PledgedValue(ByVal items As Variant) As Object
    Dim d As Object, i As Long
    Set d = CreateObject("Scripting.Dictionary")
    If Not IsEmpty(items) Then
        For i = 1 To UBound(items, 1)
            If CStr(items(i, repoCollateral.COC_LIEN)) = repoCollateral.LIEN_PLEDGED And Len(CStr(items(i, repoCollateral.COC_LOAN))) > 0 Then
                AddTo d, CStr(items(i, repoCollateral.COC_LOAN)), modMoney.ToDec(items(i, repoCollateral.COC_VALUE))
            End If
        Next i
    End If
    Set PledgedValue = d
End Function

Private Sub BuildRegister(ByVal d As Date, ByVal items As Variant, ByVal loans As Variant)
    Dim i As Long, rowsA As Collection, trow As Variant, lr As Long, pv As Object, ltv As String, lend As Variant, ltvLimit As Variant, prow As Variant, loanId As String
    Set pv = PledgedValue(items)
    Set rowsA = New Collection
    If Not IsEmpty(items) Then
        For i = 1 To UBound(items, 1)
            trow = repoCollateral.TypeRow(CStr(items(i, repoCollateral.COC_TYPE)))
            ltvLimit = CDec(1)
            loanId = CStr(items(i, repoCollateral.COC_LOAN))
            ltv = ""
            If Len(loanId) > 0 Then
                lr = repoLoan.FindLoanRow(loans, loanId)
                If lr > 0 Then
                    prow = repoLoan.ProductRow(CStr(loans(lr, repoLoan.LOC_PRODUCT)))
                    If Not IsEmpty(prow) Then ltvLimit = modMoney.ToDec(prow(repoLoan.LPR_MAXLTV))
                    If Num(pv, loanId) > 0 Then ltv = Format$(CDbl(modMoney.ToDec(loans(lr, repoLoan.LOC_OUTSTANDING))) / CDbl(Num(pv, loanId)), "0.00")
                End If
            End If
            lend = CDec(0)
            If Not IsEmpty(trow) Then lend = repoCollateral.Lendable(modMoney.ToDec(items(i, repoCollateral.COC_VALUE)), trow, ltvLimit)
            rowsA.Add Array(CStr(items(i, repoCollateral.COC_CUSTOMER)), CStr(items(i, repoCollateral.COC_ID)), CStr(items(i, repoCollateral.COC_TYPE)), CStr(items(i, repoCollateral.COC_DESC)), _
                modMoney.ToDec(items(i, repoCollateral.COC_VALUE)), lend, CStr(items(i, repoCollateral.COC_LIEN)), loanId, ltv)
        Next i
    End If
    modSheetUtils.PrepareReport SH_REGISTER, "1,2"
    modSheetUtils.SetCell SH_REGISTER, 1, 1, "Collateral register " & modDateUtils.ToIso(d) & " (SIMULATION)", True
    modSheetUtils.WriteGrid SH_REGISTER, 3, 1, svcReports.ToGrid(Array("Customer", "Collateral", "Type", "Description", "Latest value", "Lendable value", "Lien", "Loan", "Loan to value"), rowsA)
    modSheetUtils.StyleReport SH_REGISTER, 3, 9, "5,6"
End Sub

Private Sub BuildAlerts(ByVal d As Date, ByVal items As Variant, ByVal loans As Variant)
    Dim i As Long, rowsA As Collection, pv As Object, prow As Variant, lim As Variant, v As Variant, outst As Variant, ltv As Variant, id As String, lend As Variant, trow As Variant, j As Long
    Set pv = PledgedValue(items)
    Set rowsA = New Collection
    If Not IsEmpty(loans) Then
        For i = 1 To UBound(loans, 1)
            If CStr(loans(i, repoLoan.LOC_STATUS)) = repoLoan.LS_ACTIVE Then
                id = CStr(loans(i, repoLoan.LOC_ID))
                prow = repoLoan.ProductRow(CStr(loans(i, repoLoan.LOC_PRODUCT)))
                If Not IsEmpty(prow) Then
                    If CStr(prow(repoLoan.LPR_COLLATERAL)) = modConstants.FLAG_YES Then
                        outst = modMoney.ToDec(loans(i, repoLoan.LOC_OUTSTANDING))
                        v = Num(pv, id)
                        lim = modMoney.ToDec(prow(repoLoan.LPR_MAXLTV)) + repoConfig.GetDec(modConstants.CFG_LTV_MARGIN)
                        If v <= 0 Then
                            rowsA.Add Array(id, CStr(loans(i, repoLoan.LOC_CUSTOMER)), outst, v, "", "No pledged collateral")
                        Else
                            ltv = outst / v
                            If ltv > lim Then rowsA.Add Array(id, CStr(loans(i, repoLoan.LOC_CUSTOMER)), outst, v, Format$(CDbl(ltv), "0.00"), "LTV above " & Format$(CDbl(lim), "0.00"))
                        End If
                    End If
                End If
            End If
        Next i
    End If
    modSheetUtils.PrepareReport SH_ALERTS, "1,2"
    modSheetUtils.SetCell SH_ALERTS, 1, 1, "Collateral alerts " & modDateUtils.ToIso(d) & " (SIMULATION)", True
    modSheetUtils.WriteGrid SH_ALERTS, 3, 1, svcReports.ToGrid(Array("Loan", "Customer", "Principal outstanding", "Pledged value", "Loan to value", "Alert"), rowsA)
    modSheetUtils.StyleReport SH_ALERTS, 3, 6, "3,4"
End Sub

Private Sub BuildRepossessed(ByVal d As Date, ByVal repos As Variant, ByVal items As Variant, ByVal ops As Variant)
    Dim i As Long, rowsA As Collection, income As Object, cost As Object, trow As Variant, ir As Long, rd As Date, months As Long, flag As String, fv As Variant, id As String
    Set income = CreateObject("Scripting.Dictionary")
    Set cost = CreateObject("Scripting.Dictionary")
    If Not IsEmpty(ops) Then
        For i = 1 To UBound(ops, 1)
            If CStr(ops(i, 4)) = "Cost" Then AddTo cost, CStr(ops(i, 2)), modMoney.ToDec(ops(i, 5)) Else AddTo income, CStr(ops(i, 2)), modMoney.ToDec(ops(i, 5))
        Next i
    End If
    Set rowsA = New Collection
    If Not IsEmpty(repos) Then
        For i = 1 To UBound(repos, 1)
            If CStr(repos(i, repoCollateral.RPC_STATUS)) <> repoCollateral.RS_NOTICE Then
                id = CStr(repos(i, repoCollateral.RPC_ID))
                ir = repoCollateral.FindCollateralRow(items, repoCollateral.COC_ID, CStr(repos(i, repoCollateral.RPC_COLLATERAL)))
                fv = CDec(0)
                If ir > 0 Then
                    trow = repoCollateral.TypeRow(CStr(items(ir, repoCollateral.COC_TYPE)))
                    If Not IsEmpty(trow) Then fv = repoCollateral.FairValueLessCosts(modMoney.ToDec(items(ir, repoCollateral.COC_VALUE)), trow)
                End If
                months = 0
                flag = ""
                If modDateUtils.TryToDate(repos(i, repoCollateral.RPC_REPODATE), rd) Then months = DateDiff("m", rd, d)
                If (CStr(repos(i, repoCollateral.RPC_STATUS)) = repoCollateral.RS_HELD Or CStr(repos(i, repoCollateral.RPC_STATUS)) = repoCollateral.RS_RENTED) And months >= repoConfig.GetLong(modConstants.CFG_MAX_HOLDING) Then flag = "Must sell"
                rowsA.Add Array(id, CStr(repos(i, repoCollateral.RPC_LOAN)), CStr(repos(i, repoCollateral.RPC_BRANCH)), modMoney.ToDec(repos(i, repoCollateral.RPC_CARRYING)), fv, months, _
                    CStr(repos(i, repoCollateral.RPC_STATUS)), Num(income, id), Num(cost, id), flag)
            End If
        Next i
    End If
    modSheetUtils.PrepareReport SH_REPOSSESSED, "1,2,3"
    modSheetUtils.SetCell SH_REPOSSESSED, 1, 1, "Repossessed assets " & modDateUtils.ToIso(d) & " (SIMULATION)", True
    modSheetUtils.WriteGrid SH_REPOSSESSED, 3, 1, svcReports.ToGrid(Array("Repo", "Loan", "Branch", "Carrying value", "Latest fair value less costs", "Months held", "Status", "Income so far", "Costs so far", "Flag"), rowsA)
    modSheetUtils.StyleReport SH_REPOSSESSED, 3, 10, "4,5,8,9"
End Sub

Private Sub BuildSales(ByVal d As Date, ByVal sales As Variant)
    Dim i As Long, rowsA As Collection
    Set rowsA = New Collection
    If Not IsEmpty(sales) Then
        For i = 1 To UBound(sales, 1)
            rowsA.Add Array(CStr(sales(i, 1)), CStr(sales(i, 2)), sales(i, 3), CStr(sales(i, 4)), modMoney.ToDec(sales(i, 5)), modMoney.ToDec(sales(i, 6)), modMoney.ToDec(sales(i, 7)), _
                modMoney.ToDec(sales(i, 8)), modMoney.ToDec(sales(i, 9)), modMoney.ToDec(sales(i, 10)), modMoney.ToDec(sales(i, 11)))
        Next i
    End If
    modSheetUtils.PrepareReport SH_SALES, "1,2"
    modSheetUtils.SetCell SH_SALES, 1, 1, "Sale results " & modDateUtils.ToIso(d) & " (SIMULATION)", True
    modSheetUtils.WriteGrid SH_SALES, 3, 1, svcReports.ToGrid(Array("Sale", "Repo", "Date", "Buyer", "Price", "Selling costs", "Net proceeds", "Applied to claim", "Surplus to customer", "Carrying value", "Recovery or loss"), rowsA)
    modSheetUtils.StyleReport SH_SALES, 3, 11, "5,6,7,8,9,10,11"
End Sub

Private Sub BuildPnL(ByVal d As Date, ByVal repos As Variant, ByVal ops As Variant)
    Dim i As Long, rowsA As Collection, income As Object, cost As Object, id As String
    Set income = CreateObject("Scripting.Dictionary")
    Set cost = CreateObject("Scripting.Dictionary")
    If Not IsEmpty(ops) Then
        For i = 1 To UBound(ops, 1)
            If CStr(ops(i, 4)) = "Cost" Then AddTo cost, CStr(ops(i, 2)), modMoney.ToDec(ops(i, 5)) Else AddTo income, CStr(ops(i, 2)), modMoney.ToDec(ops(i, 5))
        Next i
    End If
    Set rowsA = New Collection
    If Not IsEmpty(repos) Then
        For i = 1 To UBound(repos, 1)
            id = CStr(repos(i, repoCollateral.RPC_ID))
            If income.Exists(id) Or cost.Exists(id) Then rowsA.Add Array(id, CStr(repos(i, repoCollateral.RPC_STATUS)), Num(income, id), Num(cost, id), Num(income, id) - Num(cost, id))
        Next i
    End If
    modSheetUtils.PrepareReport SH_PNL, "1"
    modSheetUtils.SetCell SH_PNL, 1, 1, "Result per held asset (rent and revenue minus costs and impairment) " & modDateUtils.ToIso(d) & " (SIMULATION)", True
    modSheetUtils.WriteGrid SH_PNL, 3, 1, svcReports.ToGrid(Array("Repo", "Status", "Rent and revenue", "Costs and impairment", "Net"), rowsA)
    modSheetUtils.StyleReport SH_PNL, 3, 5, "3,4,5"
End Sub

Private Sub BuildDeficiencies(ByVal d As Date, ByVal repos As Variant, ByVal loans As Variant)
    Dim i As Long, rowsA As Collection, lr As Long, cust As String
    Set rowsA = New Collection
    If Not IsEmpty(repos) Then
        For i = 1 To UBound(repos, 1)
            If CStr(repos(i, repoCollateral.RPC_STATUS)) <> repoCollateral.RS_NOTICE And modMoney.ToDec(repos(i, repoCollateral.RPC_DEFICIENCY)) > 0 Then
                lr = repoLoan.FindLoanRow(loans, CStr(repos(i, repoCollateral.RPC_LOAN)))
                cust = ""
                If lr > 0 Then cust = CStr(loans(lr, repoLoan.LOC_CUSTOMER))
                rowsA.Add Array(CStr(repos(i, repoCollateral.RPC_ID)), CStr(repos(i, repoCollateral.RPC_LOAN)), cust, modMoney.ToDec(repos(i, repoCollateral.RPC_CLAIM)), modMoney.ToDec(repos(i, repoCollateral.RPC_DEFICIENCY)), CStr(repos(i, repoCollateral.RPC_STATUS)))
            End If
        Next i
    End If
    modSheetUtils.PrepareReport SH_DEFICIENCY, "1,2,3"
    modSheetUtils.SetCell SH_DEFICIENCY, 1, 1, "What customers still owe after repossession " & modDateUtils.ToIso(d) & " (SIMULATION)", True
    modSheetUtils.WriteGrid SH_DEFICIENCY, 3, 1, svcReports.ToGrid(Array("Repo", "Loan", "Customer", "Claim total", "Still owed", "Status"), rowsA)
    modSheetUtils.StyleReport SH_DEFICIENCY, 3, 6, "4,5"
End Sub

' Called from svcReportsExtra. Skipped when the Stage 4 tables are missing.
Public Sub RunCollateralReports(ByVal d As Date, ByVal inp As Object)
    Dim items As Variant, repos As Variant, loans As Variant, sales As Variant, ops As Variant
    If Not modSheetUtils.TableExists(modConstants.TBL_COLLATERAL) Or Not modSheetUtils.TableExists(modConstants.TBL_REPOSSESSIONS) Or Not modSheetUtils.TableExists(modConstants.TBL_LOANS) Then Exit Sub
    modSheetUtils.EnsureSheet SH_REGISTER
    modSheetUtils.EnsureSheet SH_ALERTS
    modSheetUtils.EnsureSheet SH_REPOSSESSED
    modSheetUtils.EnsureSheet SH_SALES
    modSheetUtils.EnsureSheet SH_PNL
    modSheetUtils.EnsureSheet SH_DEFICIENCY
    items = modSheetUtils.ReadTable(modConstants.TBL_COLLATERAL)
    repos = modSheetUtils.ReadTable(modConstants.TBL_REPOSSESSIONS)
    loans = modSheetUtils.ReadTable(modConstants.TBL_LOANS)
    sales = modSheetUtils.ReadTable(modConstants.TBL_ASSETSALES)
    ops = modSheetUtils.ReadTable(modConstants.TBL_ASSETOPS)
    BuildRegister d, items, loans
    BuildAlerts d, items, loans
    BuildRepossessed d, repos, items, ops
    BuildSales d, sales
    BuildPnL d, repos, ops
    BuildDeficiencies d, repos, loans
End Sub
