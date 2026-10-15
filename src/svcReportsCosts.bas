Attribute VB_Name = "svcReportsCosts"
Option Explicit

' Stage 2 reports: BranchPL, CostAnalysis, AssetRegister, BillsDue, TaxSummary, BudgetVsActual.
' Called from svcReportsExtra.RunExtraReports. Each one rewrites its own sheet.

Private Const SH_PL As String = "BranchPL"
Private Const SH_COST As String = "CostAnalysis"
Private Const SH_ASSET As String = "AssetRegister"
Private Const SH_BILLS As String = "BillsDue"
Private Const SH_TAX As String = "TaxSummary"
Private Const SH_BUDGET As String = "BudgetVsActual"

Private Sub AddTo(ByVal d As Object, ByVal key As String, ByVal v As Variant)
    If d.Exists(key) Then d(key) = d(key) + v Else d(key) = v
End Sub

Private Function Num(ByVal d As Object, ByVal key As String) As Variant
    If d.Exists(key) Then Num = d(key) Else Num = CDec(0)
End Function

' ---- Branch profit and loss ----

Private Sub BuildBranchPL(ByVal d As Date, ByVal inp As Object)
    Dim fromD As Date, toD As Date, lines As Variant, chart As Object, branches As Object, per As Object, i As Long, dt As Date
    Dim incGL As Collection, expGL As Collection, gl As Variant, b As Variant, rowsA As Collection, hdr() As Variant, n As Long, c As Long
    Dim one() As Variant, tot() As Variant, v As Variant, ti As Variant, te As Variant, info As Variant, active As Object
    If Not modDateUtils.TryToDate(inp(modConstants.INP_FROM), fromD) Then fromD = DateSerial(Year(d), Month(d), 1)
    If Not modDateUtils.TryToDate(inp(modConstants.INP_TO), toD) Then toD = d
    lines = repoJournal.LoadDatedLines()
    Set chart = repoChart.LoadChart()
    Set branches = repoBranch.LoadBranches()
    Set per = CreateObject("Scripting.Dictionary")
    Set active = CreateObject("Scripting.Dictionary")
    If Not IsEmpty(lines) Then
        For i = 1 To UBound(lines, 1)
            If Not IsEmpty(lines(i, 1)) Then
                dt = lines(i, 1)
                If dt >= fromD And dt <= toD And CStr(lines(i, 6)) <> modConstants.JT_YEAR_END Then
                    AddTo per, CStr(lines(i, 2)) & "|" & CStr(lines(i, 3)), lines(i, 4) - lines(i, 5)
                    active(CStr(lines(i, 3))) = True
                End If
            End If
        Next i
    End If
    Set incGL = New Collection
    Set expGL = New Collection
    For Each gl In chart.Keys
        info = chart(gl)
        If active.Exists(CStr(gl)) Then
            If info(1) = "Income" Then incGL.Add CStr(gl)
            If info(1) = "Expense" Then expGL.Add CStr(gl)
        End If
    Next gl
    n = 2 + incGL.Count + 1 + expGL.Count + 1 + 1
    ReDim hdr(0 To n - 1)
    ReDim tot(0 To n - 1)
    hdr(0) = "Branch"
    hdr(1) = "Name"
    c = 2
    For Each gl In incGL
        hdr(c) = gl & " " & chart(gl)(0)
        c = c + 1
    Next gl
    hdr(c) = "Total income"
    c = c + 1
    For Each gl In expGL
        hdr(c) = gl & " " & chart(gl)(0)
        c = c + 1
    Next gl
    hdr(c) = "Total expense"
    hdr(c + 1) = "Result"
    For c = 2 To n - 1
        tot(c) = CDec(0)
    Next c
    tot(0) = "TOTAL"
    tot(1) = ""
    Set rowsA = New Collection
    For Each b In branches.Keys
        ReDim one(0 To n - 1)
        one(0) = CStr(b)
        one(1) = branches(b).Name
        ti = CDec(0)
        te = CDec(0)
        c = 2
        For Each gl In incGL
            v = -Num(per, CStr(b) & "|" & CStr(gl))
            one(c) = v
            ti = ti + v
            c = c + 1
        Next gl
        one(c) = ti
        c = c + 1
        For Each gl In expGL
            v = Num(per, CStr(b) & "|" & CStr(gl))
            one(c) = v
            te = te + v
            c = c + 1
        Next gl
        one(c) = te
        one(c + 1) = ti - te
        For c = 2 To n - 1
            tot(c) = tot(c) + one(c)
        Next c
        rowsA.Add one
    Next b
    rowsA.Add tot
    modSheetUtils.PrepareReport SH_PL, "1"
    modSheetUtils.SetCell SH_PL, 1, 1, "Branch profit and loss " & modDateUtils.ToIso(fromD) & " to " & modDateUtils.ToIso(toD) & " (SIMULATION)", True
    modSheetUtils.WriteGrid SH_PL, 3, 1, svcReports.ToGrid(hdr, rowsA)
    modSheetUtils.StyleReport SH_PL, 3, n, ""
    FormatMoney SH_PL, 3, n
End Sub

Private Sub FormatMoney(ByVal sheetName As String, ByVal headerRow As Long, ByVal lastCol As Long)
    Dim cols As String, c As Long
    For c = 3 To lastCol
        cols = cols & IIf(Len(cols) > 0, ",", "") & c
    Next c
    modSheetUtils.StyleReport sheetName, headerRow, lastCol, cols
End Sub

' ---- Cost analysis ----

Private Sub BuildCostAnalysis(ByVal d As Date)
    Dim posted As Variant, branches As Object, custs As Object, accts As Object, net As Object, chart As Object, i As Long
    Dim sumCost As Object, months As Object, nCust As Object, nAcc As Object, k As Variant, b As Variant, rowsA As Collection
    Dim income As Variant, expense As Variant, gl As Variant, info As Variant, avg As Variant, c As clsCustomer, a As clsAccount
    Dim typeSum As Object, typeCnt As Object, bt As clsBranch, key As String, ratio As Variant, perCust As Variant, perAcc As Variant
    posted = modSheetUtils.ReadTable(modConstants.TBL_COSTPOSTINGS)
    Set branches = repoBranch.LoadBranches()
    Set custs = repoCustomer.LoadCustomers()
    Set accts = repoAccount.LoadAccounts()
    Set net = repoJournal.LoadGLNetNoClose()
    Set chart = repoChart.LoadChart()
    Set sumCost = CreateObject("Scripting.Dictionary")
    Set months = CreateObject("Scripting.Dictionary")
    Set nCust = CreateObject("Scripting.Dictionary")
    Set nAcc = CreateObject("Scripting.Dictionary")
    Set typeSum = CreateObject("Scripting.Dictionary")
    Set typeCnt = CreateObject("Scripting.Dictionary")
    If Not IsEmpty(posted) Then
        For i = 1 To UBound(posted, 1)
            AddTo sumCost, CStr(posted(i, repoCost.CPC_BRANCH)), modMoney.ToDec(posted(i, repoCost.CPC_AMOUNT))
            months(CStr(posted(i, repoCost.CPC_BRANCH)) & "|" & CStr(posted(i, repoCost.CPC_PERIOD))) = True
        Next i
    End If
    For Each k In custs.Keys
        Set c = custs(k)
        AddTo nCust, c.HomeBranchID, 1
    Next k
    For Each k In accts.Keys
        Set a = accts(k)
        If a.Status <> modConstants.STATUS_CLOSED Then AddTo nAcc, a.BranchID, 1
    Next k
    Set rowsA = New Collection
    For Each b In branches.Keys
        Set bt = branches(b)
        income = CDec(0)
        expense = CDec(0)
        For Each gl In chart.Keys
            info = chart(gl)
            If info(1) = "Income" Then income = income - Num(net, CStr(b) & "|" & CStr(gl))
            If info(1) = "Expense" Then expense = expense + Num(net, CStr(b) & "|" & CStr(gl))
        Next gl
        i = 0
        For Each k In months.Keys
            If Left$(CStr(k), Len(CStr(b)) + 1) = CStr(b) & "|" Then i = i + 1
        Next k
        avg = CDec(0)
        If i > 0 Then avg = Num(sumCost, CStr(b)) / i
        AddTo typeSum, bt.BranchType, avg
        AddTo typeCnt, bt.BranchType, 1
        ratio = ""
        If income > 0 Then ratio = Format$(CDbl(expense) / CDbl(income), "0.00")
        perCust = CDec(0)
        perAcc = CDec(0)
        If Num(nCust, CStr(b)) > 0 Then perCust = avg / Num(nCust, CStr(b))
        If Num(nAcc, CStr(b)) > 0 Then perAcc = avg / Num(nAcc, CStr(b))
        rowsA.Add Array(CStr(b), bt.Name, bt.BranchType, Num(sumCost, CStr(b)), i, avg, Num(nCust, CStr(b)), Num(nAcc, CStr(b)), perCust, perAcc, income, expense, ratio)
    Next b
    For Each k In typeSum.Keys
        rowsA.Add Array("AVERAGE", CStr(k) & " branch", CStr(k), "", "", typeSum(k) / typeCnt(k), "", "", "", "", "", "", "")
    Next k
    modSheetUtils.PrepareReport SH_COST, "1"
    modSheetUtils.SetCell SH_COST, 1, 1, "Cost analysis " & modDateUtils.ToIso(d) & " (SIMULATION)", True
    modSheetUtils.WriteGrid SH_COST, 3, 1, svcReports.ToGrid(Array("Branch", "Name", "Type", "Posted costs", "Months", "Average per month", "Customers", "Accounts", _
        "Cost per customer", "Cost per account", "Income", "Expense", "Cost to income"), rowsA)
    modSheetUtils.StyleReport SH_COST, 3, 13, "4,6,9,10,11,12"
End Sub

' ---- Asset register ----

Private Sub BuildAssetRegister(ByVal d As Date)
    Dim data As Variant, i As Long, rowsA As Collection, cost As Variant, accum As Variant, res As Variant, life As Long, monthly As Variant, remaining As Variant
    data = modSheetUtils.ReadTable(modConstants.TBL_FIXEDASSETS)
    Set rowsA = New Collection
    If Not IsEmpty(data) Then
        If repoAsset.IsUpgraded(data) Then
            For i = 1 To UBound(data, 1)
                cost = modMoney.ToDec(data(i, repoAsset.FAC_COST))
                res = modMoney.ToDec(data(i, repoAsset.FAC_RESIDUAL))
                accum = modMoney.ToDec(data(i, repoAsset.FAC_ACCUM))
                life = CLng(modMoney.ToDec(data(i, repoAsset.FAC_LIFE)))
                remaining = CDec(0)
                If life > 0 And CStr(data(i, repoAsset.FAC_STATUS)) = repoAsset.ASSET_INUSE Then
                    monthly = modMoney.RoundCcy((cost - res) / life, modConstants.CCY_MMK)
                    If monthly > 0 And cost - res - accum > 0 Then remaining = -Int(-(cost - res - accum) / monthly)
                End If
                rowsA.Add Array(CStr(data(i, repoAsset.FAC_BRANCH)), CStr(data(i, repoAsset.FAC_ID)), CStr(data(i, repoAsset.FAC_TYPE)), CStr(data(i, repoAsset.FAC_DESC)), _
                    cost, accum, cost - accum, remaining, CStr(data(i, repoAsset.FAC_STATUS)))
            Next i
        End If
    End If
    modSheetUtils.PrepareReport SH_ASSET, "1,2"
    modSheetUtils.SetCell SH_ASSET, 1, 1, "Fixed asset register " & modDateUtils.ToIso(d) & " (SIMULATION)", True
    modSheetUtils.WriteGrid SH_ASSET, 3, 1, svcReports.ToGrid(Array("Branch", "Asset", "Type", "Description", "Cost", "Accumulated depreciation", "Book value", "Months left", "Status"), rowsA)
    modSheetUtils.StyleReport SH_ASSET, 3, 9, "5,6,7,8"
End Sub

' ---- Bills due ----

Private Sub BuildBillsDue(ByVal d As Date)
    Dim posted As Variant, exps As Variant, per As Object, i As Long, k As Variant, parts() As String, rowsA As Collection, dt As Date
    Set per = CreateObject("Scripting.Dictionary")
    posted = modSheetUtils.ReadTable(modConstants.TBL_COSTPOSTINGS)
    If Not IsEmpty(posted) Then
        For i = 1 To UBound(posted, 1)
            If CStr(posted(i, repoCost.CPC_STATUS)) = "Unpaid" Then AddTo per, CStr(posted(i, repoCost.CPC_BRANCH)) & "|" & CStr(posted(i, repoCost.CPC_PERIOD)), modMoney.ToDec(posted(i, repoCost.CPC_AMOUNT))
        Next i
    End If
    If modSheetUtils.TableExists(modConstants.TBL_EXPENSES) Then
        exps = modSheetUtils.ReadTable(modConstants.TBL_EXPENSES)
        If Not IsEmpty(exps) Then
            For i = 1 To UBound(exps, 1)
                If CStr(exps(i, repoCost.EXC_STATUS)) = "Posted" And modDateUtils.TryToDate(exps(i, repoCost.EXC_DATE), dt) Then
                    AddTo per, CStr(exps(i, repoCost.EXC_BRANCH)) & "|" & modDateUtils.PeriodOf(dt), modMoney.ToDec(exps(i, repoCost.EXC_AMOUNT))
                End If
            Next i
        End If
    End If
    Set rowsA = New Collection
    For Each k In per.Keys
        parts = Split(CStr(k), "|")
        rowsA.Add Array(parts(0), parts(1), per(k))
    Next k
    modSheetUtils.PrepareReport SH_BILLS, "1,2"
    modSheetUtils.SetCell SH_BILLS, 1, 1, "Bills due (unpaid) " & modDateUtils.ToIso(d) & " (SIMULATION)", True
    modSheetUtils.WriteGrid SH_BILLS, 3, 1, svcReports.ToGrid(Array("Branch", "Period", "Unpaid amount"), rowsA)
    modSheetUtils.StyleReport SH_BILLS, 3, 3, "3"
End Sub

' ---- Tax summary ----

Private Sub BuildTaxSummary(ByVal d As Date)
    Dim lines As Variant, pays As Variant, i As Long, wht As Object, com As Object, paidW As Object, paidC As Object, periods As Object, p As String
    Dim dt As Date, rowsA As Collection, k As Variant, net As Object, payW As Variant, payC As Variant
    Set wht = CreateObject("Scripting.Dictionary")
    Set com = CreateObject("Scripting.Dictionary")
    Set paidW = CreateObject("Scripting.Dictionary")
    Set paidC = CreateObject("Scripting.Dictionary")
    Set periods = CreateObject("Scripting.Dictionary")
    lines = repoJournal.LoadDatedLines()
    If Not IsEmpty(lines) Then
        For i = 1 To UBound(lines, 1)
            If Not IsEmpty(lines(i, 1)) Then
                dt = lines(i, 1)
                p = modDateUtils.PeriodOf(dt)
                If CStr(lines(i, 3)) = modConstants.GL_WHT Then
                    AddTo wht, p, lines(i, 5)
                    periods(p) = True
                ElseIf CStr(lines(i, 3)) = modConstants.GL_TAX_FEES Then
                    AddTo com, p, lines(i, 5)
                    periods(p) = True
                End If
            End If
        Next i
    End If
    If modSheetUtils.TableExists(modConstants.TBL_TAXPAYMENTS) Then
        pays = modSheetUtils.ReadTable(modConstants.TBL_TAXPAYMENTS)
        If Not IsEmpty(pays) Then
            For i = 1 To UBound(pays, 1)
                p = CStr(pays(i, 4))
                If CStr(pays(i, 2)) = modConstants.TAX_WHT Then AddTo paidW, p, modMoney.ToDec(pays(i, 5))
                If CStr(pays(i, 2)) = modConstants.TAX_COMMERCIAL Then AddTo paidC, p, modMoney.ToDec(pays(i, 5))
                periods(p) = True
            Next i
        End If
    End If
    Set rowsA = New Collection
    For Each k In periods.Keys
        rowsA.Add Array(CStr(k), Num(wht, CStr(k)), Num(com, CStr(k)), Num(paidW, CStr(k)), Num(paidC, CStr(k)))
    Next k
    Set net = repoJournal.LoadGLNet()
    payW = CDec(0)
    payC = CDec(0)
    For Each k In net.Keys
        If Mid$(CStr(k), InStr(CStr(k), "|") + 1) = modConstants.GL_WHT Then payW = payW - net(k)
        If Mid$(CStr(k), InStr(CStr(k), "|") + 1) = modConstants.GL_TAX_FEES Then payC = payC - net(k)
    Next k
    rowsA.Add Array("STILL PAYABLE", payW, payC, "", "")
    modSheetUtils.PrepareReport SH_TAX, "1"
    modSheetUtils.SetCell SH_TAX, 1, 1, "Tax summary " & modDateUtils.ToIso(d) & " (SIMULATION)", True
    modSheetUtils.WriteGrid SH_TAX, 3, 1, svcReports.ToGrid(Array("Period", "Tax withheld on interest", "Commercial tax charged", "Paid: withholding", "Paid: commercial"), rowsA)
    modSheetUtils.StyleReport SH_TAX, 3, 5, "2,3,4,5"
End Sub

' ---- Budget versus actual (year in C25) ----

Private Sub BuildBudgetVsActual(ByVal d As Date, ByVal inp As Object)
    Dim yr As Long, budgets As Variant, lines As Variant, act As Object, i As Long, dt As Date, rowsA As Collection, key As String, a As Variant, bud As Variant, pct As Variant
    yr = Year(d)
    If Trim$(CStr(inp(modConstants.INP_PERIOD))) Like "####*" Then yr = CLng(Left$(Trim$(CStr(inp(modConstants.INP_PERIOD))), 4))
    Set act = CreateObject("Scripting.Dictionary")
    lines = repoJournal.LoadDatedLines()
    If Not IsEmpty(lines) Then
        For i = 1 To UBound(lines, 1)
            If Not IsEmpty(lines(i, 1)) Then
                dt = lines(i, 1)
                If Year(dt) = yr And CStr(lines(i, 6)) <> modConstants.JT_YEAR_END Then AddTo act, CStr(lines(i, 2)) & "|" & CStr(lines(i, 3)), lines(i, 4) - lines(i, 5)
            End If
        Next i
    End If
    Set rowsA = New Collection
    If modSheetUtils.TableExists(modConstants.TBL_BUDGETS) Then
        budgets = modSheetUtils.ReadTable(modConstants.TBL_BUDGETS)
        If Not IsEmpty(budgets) Then
            For i = 1 To UBound(budgets, 1)
                If CStr(budgets(i, 3)) = CStr(yr) Then
                    key = CStr(budgets(i, 2)) & "|" & CStr(budgets(i, 4))
                    bud = modMoney.ToDec(budgets(i, 5))
                    a = Num(act, key)
                    pct = ""
                    If bud <> 0 Then pct = Format$(CDbl(a) / CDbl(bud), "0.0%")
                    rowsA.Add Array(CStr(budgets(i, 2)), CStr(budgets(i, 4)), bud, a, a - bud, pct)
                End If
            Next i
        End If
    End If
    modSheetUtils.PrepareReport SH_BUDGET, "1,2"
    modSheetUtils.SetCell SH_BUDGET, 1, 1, "Budget versus actual " & yr & " (SIMULATION)", True
    modSheetUtils.WriteGrid SH_BUDGET, 3, 1, svcReports.ToGrid(Array("Branch", "GL", "Budget", "Actual to date", "Difference", "Used"), rowsA)
    modSheetUtils.StyleReport SH_BUDGET, 3, 6, "3,4,5"
End Sub

' Entry from svcReportsExtra. Sheets are created on demand. Skipped when Stage 2 tables are missing.
Public Sub RunCostReports(ByVal d As Date, ByVal inp As Object)
    If Not modSheetUtils.TableExists(modConstants.TBL_COSTPOSTINGS) Then Exit Sub
    modSheetUtils.EnsureSheet SH_PL
    modSheetUtils.EnsureSheet SH_COST
    modSheetUtils.EnsureSheet SH_ASSET
    modSheetUtils.EnsureSheet SH_BILLS
    modSheetUtils.EnsureSheet SH_TAX
    modSheetUtils.EnsureSheet SH_BUDGET
    BuildBranchPL d, inp
    BuildCostAnalysis d
    BuildAssetRegister d
    BuildBillsDue d
    BuildTaxSummary d
    BuildBudgetVsActual d, inp
End Sub
