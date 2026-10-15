Attribute VB_Name = "svcReportsGov"
Option Explicit

' Stage 8 reports: Dashboard, RatioReport, ComplianceEvents, DepositInsurance, LargeCashReport, SuspiciousReport, AmlAlerts,
' GovSecurities, IncomeStatement, BalanceSheet, CashFlowSummary, BranchRanking. (Inspection and TaxReturn are written by their own services.)

Private Sub AddTo(ByVal d As Object, ByVal key As String, ByVal v As Variant)
    If d.Exists(key) Then d(key) = d(key) + v Else d(key) = v
End Sub

Private Function Num(ByVal d As Object, ByVal key As String) As Variant
    If d.Exists(key) Then Num = d(key) Else Num = CDec(0)
End Function

Private Sub Emit(ByVal sheetName As String, ByVal title As String, ByVal headers As Variant, ByVal rowsA As Collection, ByVal textCols As String, ByVal moneyCols As String)
    modSheetUtils.EnsureSheet sheetName
    modSheetUtils.PrepareReport sheetName, textCols
    modSheetUtils.SetCell sheetName, 1, 1, title & " (SIMULATION)", True
    modSheetUtils.WriteGrid sheetName, 3, 1, svcReports.ToGrid(headers, rowsA)
    modSheetUtils.StyleReport sheetName, 3, UBound(headers) - LBound(headers) + 1, moneyCols
End Sub

Private Function YearFromInput(ByVal inp As Object, ByVal d As Date) As Long
    Dim t As String
    YearFromInput = Year(d)
    t = Trim$(CStr(inp(modConstants.INP_PERIOD)))
    If t Like "####*" Then YearFromInput = CLng(Left$(t, 4))
End Function

' Net (debit minus credit) per "BranchID|GLCode" for journals dated in calendar year yr up to d, without closing journals.
Private Function YearNet(ByVal lines As Variant, ByVal yr As Long, ByVal d As Date) As Object
    Dim res As Object, i As Long, dt As Date
    Set res = CreateObject("Scripting.Dictionary")
    If Not IsEmpty(lines) Then
        For i = 1 To UBound(lines, 1)
            If Not IsEmpty(lines(i, 1)) Then
                dt = lines(i, 1)
                If Year(dt) = yr And dt <= d And CStr(lines(i, 6)) <> modConstants.JT_YEAR_END Then AddTo res, CStr(lines(i, 2)) & "|" & CStr(lines(i, 3)), lines(i, 4) - lines(i, 5)
            End If
        Next i
    End If
    Set YearNet = res
End Function

Private Function GlOf(ByVal key As String) As String
    GlOf = Mid$(key, InStr(key, "|") + 1)
End Function

Private Function BrOf(ByVal key As String) As String
    BrOf = Left$(key, InStr(key, "|") - 1)
End Function

' ---- Ratios ----

Private Sub BuildRatios(ByVal d As Date, ByVal uow As clsUnitOfWork)
    Dim rowsA As Collection, r As Variant, hist As Variant, i As Long, dt As Date, cur As Collection
    Set cur = svcRatios.ComputeRatios(uow, d)
    Set rowsA = New Collection
    For Each r In cur
        rowsA.Add Array(CStr(r(0)), modDateUtils.ToIso(d), Format$(CDbl(r(1)), "0.0000"), Format$(CDbl(r(2)), IIf(CBool(r(3)), "0.0000", "0.0000")), IIf(CBool(r(3)), "Minimum", "Maximum"), CStr(r(4)))
    Next r
    hist = modSheetUtils.ReadTable(modConstants.TBL_RATIOHISTORY)
    If Not IsEmpty(hist) Then
        For i = UBound(hist, 1) To 1 Step -1
            If modDateUtils.TryToDate(hist(i, 1), dt) Then
                If dt >= d - 30 And dt < d Then rowsA.Add Array(CStr(hist(i, 2)), modDateUtils.ToIso(dt), Format$(CDbl(modMoney.ToDec(hist(i, 3))), "0.0000"), Format$(CDbl(modMoney.ToDec(hist(i, 4))), "0.0000"), "History", CStr(hist(i, 5)))
            End If
        Next i
    End If
    Emit "RatioReport", "Compliance ratios " & modDateUtils.ToIso(d) & " and the last 30 days", Array("Ratio", "Date", "Actual", "Limit", "Kind", "Status"), rowsA, "1,2,3,4,5,6", ""
End Sub

Private Sub BuildEvents(ByVal d As Date)
    Dim data As Variant, i As Long, rowsA As Collection
    Set rowsA = New Collection
    data = modSheetUtils.ReadTable(modConstants.TBL_COMPLIANCE)
    If Not IsEmpty(data) Then
        For i = 1 To UBound(data, 1)
            If Len(CStr(data(i, 1))) > 0 Then rowsA.Add Array(CStr(data(i, 1)), modDateUtils.ToIso(CDate(data(i, 2))), CStr(data(i, 3)), Format$(CDbl(modMoney.ToDec(data(i, 4))), "0.0000"), Format$(CDbl(modMoney.ToDec(data(i, 5))), "0.0000"), CStr(data(i, 6)), CStr(data(i, 7)))
        Next i
    End If
    Emit "ComplianceEventsReport", "Compliance events " & modDateUtils.ToIso(d), Array("Event", "Date", "Ratio", "Actual", "Limit", "Status", "Note"), rowsA, "1,2,3,4,5,6,7", ""
End Sub

' ---- Deposit insurance ----

Private Sub BuildInsurance(ByVal d As Date, ByVal uow As clsUnitOfWork)
    Dim tot As Object, custs As Object, k As Variant, c As clsCustomer, cap As Variant, ins As Object, unins As Object, over As Object, rowsA As Collection, b As Variant, brs As Object, ti As Variant, tu As Variant
    Set tot = svcInsurance.CustomerTotals(uow)
    Set custs = repoCustomer.LoadCustomers()
    cap = repoConfig.GetDec(modConstants.CFG_INS_CAP)
    Set ins = CreateObject("Scripting.Dictionary")
    Set unins = CreateObject("Scripting.Dictionary")
    Set over = CreateObject("Scripting.Dictionary")
    For Each k In tot.Keys
        If custs.Exists(CStr(k)) Then
            Set c = custs(CStr(k))
            If tot(k) > cap Then
                AddTo ins, c.HomeBranchID, cap
                AddTo unins, c.HomeBranchID, tot(k) - cap
                AddTo over, c.HomeBranchID, 1
            Else
                AddTo ins, c.HomeBranchID, tot(k)
                AddTo unins, c.HomeBranchID, CDec(0)
            End If
        End If
    Next k
    Set brs = repoBranch.LoadBranches()
    Set rowsA = New Collection
    ti = CDec(0)
    tu = CDec(0)
    For Each b In brs.Keys
        rowsA.Add Array(CStr(b), Num(ins, CStr(b)), Num(unins, CStr(b)), Num(over, CStr(b)))
        ti = ti + Num(ins, CStr(b))
        tu = tu + Num(unins, CStr(b))
    Next b
    rowsA.Add Array("TOTAL (cap " & modMoney.FormatAmount(cap) & " per customer)", ti, tu, "")
    Emit "DepositInsurance", "Deposit insurance " & modDateUtils.ToIso(d) & ": quarterly premium at the rate of " & Format$(CDbl(repoConfig.GetDec(modConstants.CFG_INS_RATE)), "0.0000") & " a year = " & modMoney.FormatAmount(modMoney.RoundCcy(ti * repoConfig.GetDec(modConstants.CFG_INS_RATE) / 4, modConstants.CCY_MMK)), _
        Array("Branch", "Insured", "Uninsured", "Customers above cap"), rowsA, "1", "2,3,4"
End Sub

' ---- AML reports ----

Private Sub BuildLargeCash(ByVal d As Date, ByVal inp As Object)
    Dim tx As Variant, i As Long, fromD As Date, toD As Date, th As Variant, rd As Variant, rowsA As Collection, dt As Date, t As String, amt As Variant
    If Not modDateUtils.TryToDate(inp(modConstants.INP_FROM), fromD) Then fromD = DateSerial(Year(d), Month(d), 1)
    If Not modDateUtils.TryToDate(inp(modConstants.INP_TO), toD) Then toD = d
    th = CDec(20000000)
    rd = modSheetUtils.ReadTable(modConstants.TBL_AMLRULES)
    If Not IsEmpty(rd) Then
        For i = 1 To UBound(rd, 1)
            If CStr(rd(i, 2)) = "LargeCash" Then th = modMoney.ToDec(rd(i, 3))
        Next i
    End If
    Set rowsA = New Collection
    tx = repoTransaction.LoadTransactions()
    If Not IsEmpty(tx) Then
        For i = 1 To UBound(tx, 1)
            If modDateUtils.TryToDate(tx(i, 3), dt) Then
                t = CStr(tx(i, 5))
                amt = modMoney.ToDec(tx(i, 7))
                If dt >= fromD And dt <= toD And amt >= th And (t = modConstants.TXN_DEPOSIT Or t = modConstants.TXN_WITHDRAWAL Or InStr(1, t, "Cash", vbTextCompare) > 0 Or InStr(1, t, "Agent", vbTextCompare) > 0) And CStr(tx(i, 14)) <> modConstants.STATUS_REVERSED Then
                    rowsA.Add Array(CStr(tx(i, 1)), modDateUtils.ToIso(dt), CStr(tx(i, 4)), t, amt, CStr(tx(i, 10)))
                End If
            End If
        Next i
    End If
    Emit "LargeCashReport", "Large cash transactions of " & modMoney.FormatAmount(th) & " MMK or more, " & modDateUtils.ToIso(fromD) & " to " & modDateUtils.ToIso(toD), Array("Transaction", "Date", "Account", "Type", "Amount", "Branch"), rowsA, "1,2,3,4,6", "5"
End Sub

Private Sub BuildAlerts(ByVal d As Date)
    Dim data As Variant, i As Long, rowsA As Collection
    Set rowsA = New Collection
    data = modSheetUtils.ReadTable(modConstants.TBL_AMLALERTS)
    If Not IsEmpty(data) Then
        For i = 1 To UBound(data, 1)
            If Len(CStr(data(i, 1))) > 0 Then rowsA.Add Array(CStr(data(i, 1)), modDateUtils.ToIso(CDate(data(i, 2))), CStr(data(i, 3)), CStr(data(i, 4)), CStr(data(i, 5)), modMoney.ToDec(data(i, 6)), CStr(data(i, 7)), CStr(data(i, 8)), CStr(data(i, 9)), CStr(data(i, 10)))
        Next i
    End If
    Emit "AmlAlertsReport", "AML alerts " & modDateUtils.ToIso(d), Array("Alert", "Date", "Rule", "Customer", "Account", "Amount", "Description", "Status", "Case", "Reviewed by"), rowsA, "1,2,3,4,5,7,8,9,10", "6"
End Sub

Private Sub BuildSuspicious(ByVal d As Date)
    Dim cases As Variant, alerts As Variant, i As Long, j As Long, rowsA As Collection
    Set rowsA = New Collection
    cases = modSheetUtils.ReadTable(modConstants.TBL_AMLCASES)
    alerts = modSheetUtils.ReadTable(modConstants.TBL_AMLALERTS)
    If Not IsEmpty(cases) Then
        For i = 1 To UBound(cases, 1)
            If CStr(cases(i, repoGov.CSC_STATUS)) = "Reported" Then
                rowsA.Add Array(CStr(cases(i, 1)), CStr(cases(i, repoGov.CSC_CUSTOMER)), CStr(cases(i, 4)), CStr(cases(i, repoGov.CSC_ASSIGNED)), "Case", "")
                If Not IsEmpty(alerts) Then
                    For j = 1 To UBound(alerts, 1)
                        If CStr(alerts(j, repoGov.ALC_CASE)) = CStr(cases(i, 1)) Then rowsA.Add Array(CStr(cases(i, 1)), CStr(alerts(j, repoGov.ALC_CUSTOMER)), CStr(alerts(j, 7)), CStr(alerts(j, 5)), "Alert " & CStr(alerts(j, 1)), modMoney.ToDec(alerts(j, 6)))
                    Next j
                End If
            End If
        Next i
    End If
    Emit "SuspiciousReport", "Suspicious activity report: reported cases " & modDateUtils.ToIso(d), Array("Case", "Customer", "Summary or description", "Assigned or account", "Kind", "Amount"), rowsA, "1,2,3,4,5", "6"
End Sub

Private Sub BuildSecurities(ByVal d As Date)
    Dim b As Variant, i As Long, rowsA As Collection, recv As Variant
    Set rowsA = New Collection
    b = modSheetUtils.ReadTable(modConstants.TBL_BONDS)
    If Not IsEmpty(b) Then
        For i = 1 To UBound(b, 1)
            If Len(CStr(b(i, 1))) > 0 Then rowsA.Add Array(CStr(b(i, 1)), CStr(b(i, repoGov.BDC_TYPE)), modMoney.ToDec(b(i, repoGov.BDC_FACE)), modDateUtils.ToIso(CDate(b(i, repoGov.BDC_ISSUE))), modDateUtils.ToIso(CDate(b(i, repoGov.BDC_MATURITY))), _
                Format$(CDbl(modMoney.ToDec(b(i, repoGov.BDC_RATE))), "0.0000"), modMoney.ToDec(b(i, repoGov.BDC_PRICE)), modMoney.ToDec(b(i, repoGov.BDC_CARRYING)), modMoney.ToDec(b(i, repoGov.BDC_ACCRUED)), CStr(b(i, repoGov.BDC_STATUS)))
        Next i
    End If
    recv = svcBonds.UnpaidCoupons(b, d)
    rowsA.Add Array("Coupon receivable (GL 1610 share)", "", recv, "", "", "", "", "", "", "")
    Emit "GovSecurities", "Government securities " & modDateUtils.ToIso(d), Array("Bond", "Type", "Face", "Issued", "Maturity", "Rate", "Price", "Carrying value", "Accrued coupon (memo)", "Status"), rowsA, "1,2,4,5,6,10", "3,7,8,9"
End Sub

' ---- Annual reports ----

Private Sub BuildIncomeStatement(ByVal d As Date, ByVal yr As Long, ByVal lines As Variant, ByVal chart As Object)
    Dim net As Object, per As Object, k As Variant, gl As String, rowsA As Collection, ti As Variant, te As Variant, info As Variant, gls As Collection, g As Variant
    Set net = YearNet(lines, yr, DateSerial(yr, 12, 31))
    Set per = CreateObject("Scripting.Dictionary")
    For Each k In net.Keys
        AddTo per, GlOf(CStr(k)), net(k)
    Next k
    Set rowsA = New Collection
    ti = CDec(0)
    te = CDec(0)
    For Each k In per.Keys
        gl = CStr(k)
        If chart.Exists(gl) And per(k) <> 0 Then
            If chart(gl)(1) = "Income" Then rowsA.Add Array("Income", gl, CStr(chart(gl)(0)), -per(k)): ti = ti - per(k)
        End If
    Next k
    rowsA.Add Array("Income", "", "TOTAL INCOME", ti)
    For Each k In per.Keys
        gl = CStr(k)
        If chart.Exists(gl) And per(k) <> 0 Then
            If chart(gl)(1) = "Expense" Then rowsA.Add Array("Expense", gl, CStr(chart(gl)(0)), per(k)): te = te + per(k)
        End If
    Next k
    rowsA.Add Array("Expense", "", "TOTAL EXPENSES", te)
    rowsA.Add Array("Result", "", "PROFIT FOR THE YEAR " & yr, ti - te)
    Emit "IncomeStatement", "Income statement " & yr, Array("Section", "GL", "Account", "MMK"), rowsA, "1,2,3", "4"
End Sub

Private Sub BuildBalanceSheet(ByVal d As Date, ByVal lines As Variant, ByVal chart As Object)
    Dim per As Object, i As Long, dt As Date, k As Variant, gl As String, rowsA As Collection, ta As Variant, tl As Variant, te As Variant, pl As Variant, typ As Variant
    Set per = CreateObject("Scripting.Dictionary")
    If Not IsEmpty(lines) Then
        For i = 1 To UBound(lines, 1)
            If Not IsEmpty(lines(i, 1)) Then
                dt = lines(i, 1)
                If dt <= d Then AddTo per, CStr(lines(i, 3)), lines(i, 4) - lines(i, 5)
            End If
        Next i
    End If
    Set rowsA = New Collection
    ta = CDec(0)
    tl = CDec(0)
    te = CDec(0)
    pl = CDec(0)
    For Each typ In Array("Asset", "Liability", "Equity")
        For Each k In per.Keys
            gl = CStr(k)
            If chart.Exists(gl) And per(k) <> 0 Then
                If chart(gl)(1) = typ Then
                    If typ = "Asset" Then
                        rowsA.Add Array(typ, gl, CStr(chart(gl)(0)), per(k))
                        ta = ta + per(k)
                    ElseIf typ = "Liability" Then
                        rowsA.Add Array(typ, gl, CStr(chart(gl)(0)), -per(k))
                        tl = tl - per(k)
                    Else
                        rowsA.Add Array(typ, gl, CStr(chart(gl)(0)), -per(k))
                        te = te - per(k)
                    End If
                End If
            End If
        Next k
    Next typ
    For Each k In per.Keys
        If repoGov.IsProfitGL(chart, CStr(k)) Then pl = pl - per(k)
    Next k
    rowsA.Add Array("Equity", "", "Profit of the current year (not yet closed)", pl)
    rowsA.Add Array("Total", "", "TOTAL ASSETS", ta)
    rowsA.Add Array("Total", "", "TOTAL LIABILITIES + EQUITY", tl + te + pl)
    Emit "BalanceSheet", "Balance sheet " & modDateUtils.ToIso(d), Array("Section", "GL", "Account", "MMK"), rowsA, "1,2,3", "4"
End Sub

Private Sub BuildCashFlow(ByVal d As Date, ByVal inp As Object, ByVal lines As Variant)
    Dim fromD As Date, toD As Date, i As Long, dt As Date, gl As String, opening As Object, causes As Object, k As Variant, rowsA As Collection, closing As Object, g As Variant
    If Not modDateUtils.TryToDate(inp(modConstants.INP_FROM), fromD) Then fromD = DateSerial(Year(d), 1, 1)
    If Not modDateUtils.TryToDate(inp(modConstants.INP_TO), toD) Then toD = d
    Set opening = CreateObject("Scripting.Dictionary")
    Set causes = CreateObject("Scripting.Dictionary")
    If Not IsEmpty(lines) Then
        For i = 1 To UBound(lines, 1)
            If Not IsEmpty(lines(i, 1)) Then
                dt = lines(i, 1)
                gl = CStr(lines(i, 3))
                If gl = modConstants.GL_VAULT Or gl = modConstants.GL_TREASURY Or gl = modConstants.GL_GOV_SEC Then
                    If dt < fromD Then
                        AddTo opening, gl, lines(i, 4) - lines(i, 5)
                    ElseIf dt <= toD Then
                        AddTo causes, gl & "|" & CStr(lines(i, 6)), lines(i, 4) - lines(i, 5)
                    End If
                End If
            End If
        Next i
    End If
    Set rowsA = New Collection
    For Each g In Array(modConstants.GL_VAULT, modConstants.GL_TREASURY, modConstants.GL_GOV_SEC)
        rowsA.Add Array(CStr(g), "Opening balance " & modDateUtils.ToIso(fromD), Num(opening, CStr(g)))
        For Each k In causes.Keys
            If GlOf(CStr(k)) <> "" And Left$(CStr(k), 4) = CStr(g) Then rowsA.Add Array(CStr(g), "Change by " & Mid$(CStr(k), 6), causes(k))
        Next k
        Dim tot As Variant
        tot = Num(opening, CStr(g))
        For Each k In causes.Keys
            If Left$(CStr(k), 4) = CStr(g) Then tot = tot + causes(k)
        Next k
        rowsA.Add Array(CStr(g), "Closing balance " & modDateUtils.ToIso(toD), tot)
    Next g
    Emit "CashFlowSummary", "Change in cash (1000), treasury (1100) and government securities (1620) " & modDateUtils.ToIso(fromD) & " to " & modDateUtils.ToIso(toD), Array("GL", "Line", "MMK"), rowsA, "1,2", "3"
End Sub

Private Sub BuildRanking(ByVal d As Date, ByVal yr As Long, ByVal lines As Variant, ByVal chart As Object, ByVal uow As clsUnitOfWork)
    Dim net As Object, brs As Object, b As Variant, k As Variant, inc As Object, ex As Object, staff As Object, rowsA As Collection, gl As String, dep As Variant, ln As Variant, bt As String
    Dim loans As Variant, i As Long, tot As Object, bad As Object, profit As Object, order As Variant, j As Long, tmp As Variant, n As Long, ranked() As Variant, pr As Variant, cti As Variant, npl As Variant
    Set net = YearNet(lines, yr, d)
    Set brs = repoBranch.LoadBranches()
    Set inc = CreateObject("Scripting.Dictionary")
    Set ex = CreateObject("Scripting.Dictionary")
    Set staff = CreateObject("Scripting.Dictionary")
    Set tot = CreateObject("Scripting.Dictionary")
    Set bad = CreateObject("Scripting.Dictionary")
    For Each k In net.Keys
        gl = GlOf(CStr(k))
        bt = BrOf(CStr(k))
        If chart.Exists(gl) Then
            If chart(gl)(1) = "Income" Then AddTo inc, bt, -net(k)
            If chart(gl)(1) = "Expense" Then AddTo ex, bt, net(k)
            If gl = modConstants.GL_SALARIES Or gl = modConstants.GL_EMPLOYER_SSB Then AddTo staff, bt, net(k)
        End If
    Next k
    If modSheetUtils.TableExists(modConstants.TBL_LOANS) Then
        loans = modSheetUtils.ReadTable(modConstants.TBL_LOANS)
        If Not IsEmpty(loans) Then
            For i = 1 To UBound(loans, 1)
                If CStr(loans(i, repoLoan.LOC_STATUS)) = repoLoan.LS_ACTIVE Then
                    AddTo tot, CStr(loans(i, repoLoan.LOC_BRANCH)), modMoney.ToDec(loans(i, repoLoan.LOC_OUTSTANDING))
                    If CLng(modMoney.ToDec(loans(i, repoLoan.LOC_CLASS))) >= 3 Then AddTo bad, CStr(loans(i, repoLoan.LOC_BRANCH)), modMoney.ToDec(loans(i, repoLoan.LOC_OUTSTANDING))
                End If
            Next i
        End If
    End If
    n = brs.Count
    ReDim ranked(1 To n, 1 To 8)
    i = 0
    For Each b In brs.Keys
        i = i + 1
        pr = Num(inc, CStr(b)) - Num(ex, CStr(b))
        If Num(inc, CStr(b)) > 0 Then cti = Num(ex, CStr(b)) / Num(inc, CStr(b)) Else cti = CDec(0)
        If Num(tot, CStr(b)) > 0 Then npl = Num(bad, CStr(b)) / Num(tot, CStr(b)) Else npl = CDec(0)
        dep = -(uow.GLNet(CStr(b), modConstants.GL_SAVINGS) + uow.GLNet(CStr(b), modConstants.GL_CURRENT) + uow.GLNet(CStr(b), modConstants.GL_FIXED) + uow.GLNet(CStr(b), modConstants.GL_WALLETS))
        ranked(i, 1) = CStr(b)
        ranked(i, 2) = pr
        ranked(i, 3) = Format$(CDbl(cti), "0.00")
        ranked(i, 4) = dep
        ranked(i, 5) = Num(tot, CStr(b))
        ranked(i, 6) = Format$(CDbl(npl), "0.0000")
        ranked(i, 7) = Num(staff, CStr(b))
    Next b
    ' Simple selection sort by profit, highest first, then the rank is the position.
    For i = 1 To n - 1
        For j = i + 1 To n
            If ranked(j, 2) > ranked(i, 2) Then
                For k = 1 To 7
                    tmp = ranked(i, k): ranked(i, k) = ranked(j, k): ranked(j, k) = tmp
                Next k
            End If
        Next j
    Next i
    Set rowsA = New Collection
    For i = 1 To n
        rowsA.Add Array(i, ranked(i, 1), ranked(i, 2), ranked(i, 3), ranked(i, 4), ranked(i, 5), ranked(i, 6), ranked(i, 7))
    Next i
    Emit "BranchRanking", "Branch ranking by profit " & yr & " (up to " & modDateUtils.ToIso(d) & ")", Array("Rank", "Branch", "Profit", "Cost to income", "Deposits", "Loans", "NPL ratio", "Staff cost"), rowsA, "2,4,7", "3,5,6,8"
End Sub

' ---- Dashboard ----

Private Function OpenBreaches() As Long
    Dim hist As Variant, i As Long, last As Object, dt As Date, k As Variant, code As String, best As Object
    hist = modSheetUtils.ReadTable(modConstants.TBL_RATIOHISTORY)
    If IsEmpty(hist) Then Exit Function
    Set last = CreateObject("Scripting.Dictionary")
    Set best = CreateObject("Scripting.Dictionary")
    For i = 1 To UBound(hist, 1)
        If modDateUtils.TryToDate(hist(i, 1), dt) Then
            code = CStr(hist(i, 2))
            If Not best.Exists(code) Then
                best(code) = dt
                last(code) = CStr(hist(i, 5))
            ElseIf dt >= best(code) Then
                best(code) = dt
                last(code) = CStr(hist(i, 5))
            End If
        End If
    Next i
    For Each k In last.Keys
        If last(k) = "Breach" Then OpenBreaches = OpenBreaches + 1
    Next k
End Function

Private Function CountWhere(ByVal tableName As String, ByVal col As Long, ByVal a As String, Optional ByVal b As String = "|") As Long
    Dim data As Variant, i As Long
    If Not modSheetUtils.TableExists(tableName) Then Exit Function
    data = modSheetUtils.ReadTable(tableName)
    If IsEmpty(data) Then Exit Function
    For i = 1 To UBound(data, 1)
        If Len(CStr(data(i, 1))) > 0 And (CStr(data(i, col)) = a Or CStr(data(i, col)) = b) Then CountWhere = CountWhere + 1
    Next i
End Function

Private Sub BuildDashboard(ByVal d As Date, ByVal uow As clsUnitOfWork, ByVal lines As Variant, ByVal chart As Object)
    Dim rowsA As Collection, savDep As Variant, curDep As Variant, fixDep As Variant, walDep As Variant, depTot As Variant, loans As Variant, ratios As Collection, r As Variant, net As Object, k As Variant, gl As String
    Dim incTot As Variant, expTot As Variant, runCost As Variant, intInc As Variant, intExp As Variant, baseBal As Variant, nim As Variant, cti As Variant, brProfit As Object, bt As String, items() As Variant, n As Long, i As Long, j As Long, tmp As Variant
    savDep = -uow.GLNetAll(modConstants.GL_SAVINGS)
    curDep = -uow.GLNetAll(modConstants.GL_CURRENT)
    fixDep = -uow.GLNetAll(modConstants.GL_FIXED)
    walDep = -uow.GLNetAll(modConstants.GL_WALLETS)
    depTot = savDep + curDep + fixDep + walDep
    loans = uow.GLNetAll(modConstants.GL_LOANS) + uow.GLNetAll(modConstants.GL_LOAN_ALLOWANCE)
    Set net = YearNet(lines, Year(d), d)
    Set brProfit = CreateObject("Scripting.Dictionary")
    incTot = CDec(0)
    expTot = CDec(0)
    runCost = CDec(0)
    intInc = CDec(0)
    intExp = CDec(0)
    For Each k In net.Keys
        gl = GlOf(CStr(k))
        bt = BrOf(CStr(k))
        If chart.Exists(gl) Then
            If chart(gl)(1) = "Income" Then
                incTot = incTot - net(k)
                AddTo brProfit, bt, -net(k)
                If gl = modConstants.GL_LOAN_INCOME Or gl = modConstants.GL_GOV_INCOME Then intInc = intInc - net(k)
            ElseIf chart(gl)(1) = "Expense" Then
                expTot = expTot + net(k)
                AddTo brProfit, bt, -net(k)
                If gl = modConstants.GL_INT_EXPENSE Then intExp = intExp + net(k)
                If gl <> modConstants.GL_INT_EXPENSE And gl <> modConstants.GL_INCOME_TAX_EXP Then runCost = runCost + net(k)
            End If
        End If
    Next k
    baseBal = loans + uow.GLNetAll(modConstants.GL_TREASURY) + uow.GLNetAll(modConstants.GL_GOV_SEC)
    nim = CDec(0)
    If baseBal > 0 Then nim = (intInc - intExp) / baseBal
    cti = CDec(0)
    If incTot > 0 Then cti = runCost / incTot
    Set rowsA = New Collection
    rowsA.Add Array("Deposits: savings", savDep)
    rowsA.Add Array("Deposits: current", curDep)
    rowsA.Add Array("Deposits: fixed", fixDep)
    rowsA.Add Array("Deposits: wallets (wallet float)", walDep)
    rowsA.Add Array("Deposits: total", depTot)
    rowsA.Add Array("Loans (net of allowance)", loans)
    If depTot > 0 Then rowsA.Add Array("Loan-to-deposit ratio", Format$(CDbl(loans / depTot), "0.0000")) Else rowsA.Add Array("Loan-to-deposit ratio", "n/a")
    rowsA.Add Array("Liquid assets (vault + 1100 + 1620)", uow.GLNetAll(modConstants.GL_VAULT) + uow.GLNetAll(modConstants.GL_TREASURY) + uow.GLNetAll(modConstants.GL_GOV_SEC))
    Set ratios = svcRatios.ComputeRatios(uow, d)
    For Each r In ratios
        rowsA.Add Array("Ratio " & CStr(r(0)) & " (limit " & Format$(CDbl(r(2)), "0.0000") & ")", Format$(CDbl(r(1)), "0.0000") & " " & CStr(r(4)))
    Next r
    rowsA.Add Array("Investments (1600)", uow.GLNetAll(modConstants.GL_INVESTMENTS))
    rowsA.Add Array("Net interest margin (year to date, on current balances)", Format$(CDbl(nim), "0.0000"))
    rowsA.Add Array("Cost-to-income (running costs, salaries, depreciation / income)", Format$(CDbl(cti), "0.0000"))
    rowsA.Add Array("Profit year to date", incTot - expTot)
    rowsA.Add Array("Staff count (Active)", CountWhere(modConstants.TBL_EMPLOYEES, repoStaff.EMC_STATUS, modConstants.STATUS_ACTIVE))
    rowsA.Add Array("Open ratio breaches", OpenBreaches())
    rowsA.Add Array("Open AML alerts (New or Reviewing)", CountWhere(modConstants.TBL_AMLALERTS, repoGov.ALC_STATUS, "New", "Reviewing"))
    rowsA.Add Array("Pending approvals", CountWhere(modConstants.TBL_APPROVALS, repoApproval.APC_STATUS, modConstants.STATUS_PENDING))
    n = brProfit.Count
    If n > 0 Then
        ReDim items(1 To n, 1 To 2)
        i = 0
        For Each k In brProfit.Keys
            i = i + 1
            items(i, 1) = CStr(k)
            items(i, 2) = brProfit(k)
        Next k
        For i = 1 To n - 1
            For j = i + 1 To n
                If items(j, 2) > items(i, 2) Then
                    tmp = items(i, 1): items(i, 1) = items(j, 1): items(j, 1) = tmp
                    tmp = items(i, 2): items(i, 2) = items(j, 2): items(j, 2) = tmp
                End If
            Next j
        Next i
        For i = 1 To IIf(n < 3, n, 3)
            rowsA.Add Array("Top branch " & i & ": " & items(i, 1), items(i, 2))
        Next i
        For i = 0 To IIf(n < 3, n, 3) - 1
            rowsA.Add Array("Bottom branch " & (i + 1) & ": " & items(n - i, 1), items(n - i, 2))
        Next i
    End If
    Emit "Dashboard", "Management dashboard " & modDateUtils.ToIso(d), Array("Item", "Value"), rowsA, "1", "2"
End Sub

' Entry for RunReports (called from svcReportsExtra).
Public Sub RunGovReports(ByVal d As Date, ByVal inp As Object)
    Dim uow As clsUnitOfWork, lines As Variant, chart As Object, yr As Long
    If Not modSheetUtils.TableExists(modConstants.TBL_RATIOHISTORY) Or Not modSheetUtils.TableExists(modConstants.TBL_AMLALERTS) Or Not modSheetUtils.TableExists(modConstants.TBL_BONDS) Then Exit Sub
    Set uow = New clsUnitOfWork
    uow.Init d, "REPORT"
    lines = repoJournal.LoadDatedLines()
    Set chart = repoChart.LoadChart()
    yr = YearFromInput(inp, d)
    BuildRatios d, uow
    BuildEvents d
    BuildInsurance d, uow
    BuildLargeCash d, inp
    BuildAlerts d
    BuildSuspicious d
    BuildSecurities d
    BuildIncomeStatement d, yr, lines, chart
    BuildBalanceSheet d, lines, chart
    BuildCashFlow d, inp, lines
    BuildRanking d, yr, lines, chart, uow
    BuildDashboard d, uow, lines, chart
End Sub
