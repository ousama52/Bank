Attribute VB_Name = "svcReports"
Option Explicit

' Generated sheets (cleared and rewritten each time): TrialBalance, BranchSummary, CashPosition, DepositSummary, Checks, Statement.
' Services never touch worksheets directly; they build arrays and hand them to modSheetUtils.

Private Function Num(ByVal d As Object, ByVal key As String) As Variant
    If d.Exists(key) Then Num = d(key) Else Num = CDec(0)
End Function

Private Function DebitSide(ByVal net As Variant) As Variant
    If net > 0 Then DebitSide = net Else DebitSide = CDec(0)
End Function

Private Function CreditSide(ByVal net As Variant) As Variant
    If net < 0 Then CreditSide = -net Else CreditSide = CDec(0)
End Function

' ---- Trial balance: by GL, then by branch and GL ----

Private Sub BuildTrial(ByVal d As Date)
    Dim net As Object, chart As Object, branches As Object, glTot As Object, k As Variant, p As Long, g As String
    Dim rowsA As Collection, rowsB As Collection, totD As Variant, totC As Variant, v As Variant, r As Long, b As Variant, gl As Variant, info As Variant
    Set net = repoJournal.LoadGLNet()
    Set chart = repoChart.LoadChart()
    Set branches = repoBranch.LoadBranches()
    Set glTot = CreateObject("Scripting.Dictionary")
    For Each k In net.Keys
        g = Mid$(CStr(k), InStr(CStr(k), "|") + 1)
        If glTot.Exists(g) Then glTot(g) = glTot(g) + net(k) Else glTot(g) = net(k)
    Next k
    Set rowsA = New Collection
    Set rowsB = New Collection
    totD = CDec(0)
    totC = CDec(0)
    For Each gl In chart.Keys
        info = chart(gl)
        v = Num(glTot, CStr(gl))
        If v <> 0 Or info(3) Then
            rowsA.Add Array(CStr(gl), info(0), DebitSide(v), CreditSide(v))
            totD = totD + DebitSide(v)
            totC = totC + CreditSide(v)
        End If
    Next gl
    rowsA.Add Array("", "TOTAL", totD, totC)
    For Each b In branches.Keys
        For Each gl In chart.Keys
            v = Num(net, CStr(b) & "|" & CStr(gl))
            If v <> 0 Then rowsB.Add Array(CStr(b), CStr(gl), chart(gl)(0), DebitSide(v), CreditSide(v))
        Next gl
    Next b
    modSheetUtils.PrepareReport modConstants.SHEET_TRIAL, "1,2"
    modSheetUtils.SetCell modConstants.SHEET_TRIAL, 1, 1, "Trial balance " & modDateUtils.ToIso(d) & " (SIMULATION)", True
    modSheetUtils.WriteGrid modConstants.SHEET_TRIAL, 3, 1, ToGrid(Array("GL", "Name", "Debit", "Credit"), rowsA)
    modSheetUtils.StyleReport modConstants.SHEET_TRIAL, 3, 4, "3,4"
    r = 3 + rowsA.Count + 3
    modSheetUtils.SetCell modConstants.SHEET_TRIAL, r - 1, 1, "By branch and GL", True
    modSheetUtils.WriteGrid modConstants.SHEET_TRIAL, r, 1, ToGrid(Array("Branch", "GL", "Name", "Debit", "Credit"), rowsB)
    modSheetUtils.StyleReport modConstants.SHEET_TRIAL, r, 5, "4,5"
    p = r
End Sub

' Header row + data rows as one 2D array.
Public Function ToGrid(ByVal headers As Variant, ByVal dataRows As Collection) As Variant
    Dim out() As Variant, r As Long, c As Long, n As Long, k As Long, one As Variant
    n = dataRows.Count + 1
    k = UBound(headers) - LBound(headers) + 1
    ReDim out(1 To n, 1 To k)
    For c = 1 To k
        out(1, c) = headers(LBound(headers) + c - 1)
    Next c
    For r = 1 To dataRows.Count
        one = dataRows(r)
        For c = 1 To k
            out(r + 1, c) = one(LBound(one) + c - 1)
        Next c
    Next r
    ToGrid = out
End Function

' ---- Branch summary and cash position ----

Private Sub BuildBranchSummary(ByVal d As Date)
    Dim net As Object, branches As Object, custs As Object, accts As Object, chart As Object, rowsA As Collection, rowsC As Collection
    Dim b As Variant, k As Variant, c As clsCustomer, a As clsAccount, bt As clsBranch, nCust As Object, nAcc As Object, dep As Object
    Dim income As Variant, expense As Variant, gl As Variant, info As Variant, vault As Variant, flag As String
    Set net = repoJournal.LoadGLNetNoClose()
    Set branches = repoBranch.LoadBranches()
    Set custs = repoCustomer.LoadCustomers()
    Set accts = repoAccount.LoadAccounts()
    Set chart = repoChart.LoadChart()
    Set nCust = CreateObject("Scripting.Dictionary")
    Set nAcc = CreateObject("Scripting.Dictionary")
    Set dep = CreateObject("Scripting.Dictionary")
    For Each k In custs.Keys
        Set c = custs(k)
        If nCust.Exists(c.HomeBranchID) Then nCust(c.HomeBranchID) = nCust(c.HomeBranchID) + 1 Else nCust(c.HomeBranchID) = 1
    Next k
    For Each k In accts.Keys
        Set a = accts(k)
        If a.Status <> modConstants.STATUS_CLOSED Then
            If nAcc.Exists(a.BranchID) Then nAcc(a.BranchID) = nAcc(a.BranchID) + 1 Else nAcc(a.BranchID) = 1
            If dep.Exists(a.BranchID) Then dep(a.BranchID) = dep(a.BranchID) + repoFxRate.ToBase(a.Balance, a.Ccy, d) Else dep(a.BranchID) = repoFxRate.ToBase(a.Balance, a.Ccy, d)
        End If
    Next k
    Set rowsA = New Collection
    Set rowsC = New Collection
    For Each b In branches.Keys
        Set bt = branches(b)
        income = CDec(0)
        expense = CDec(0)
        For Each gl In chart.Keys
            info = chart(gl)
            If info(1) = "Income" Then income = income - Num(net, CStr(b) & "|" & CStr(gl))
            If info(1) = "Expense" Then expense = expense + Num(net, CStr(b) & "|" & CStr(gl))
        Next gl
        vault = Num(net, CStr(b) & "|" & modConstants.GL_VAULT)
        rowsA.Add Array(CStr(b), bt.Name, Num(nCust, CStr(b)), Num(nAcc, CStr(b)), Num(dep, CStr(b)), vault, _
            Num(net, CStr(b) & "|" & modConstants.GL_INTERBRANCH), -Num(net, CStr(b) & "|" & modConstants.GL_FEE_INCOME), _
            Num(net, CStr(b) & "|" & modConstants.GL_INT_EXPENSE), income - expense)
        flag = "OK"
        If vault < bt.VaultMin Then flag = "BELOW VaultMin"
        If vault > bt.VaultMax Then flag = "ABOVE VaultMax"
        rowsC.Add Array(CStr(b), bt.Name, vault, bt.VaultMin, bt.VaultMax, flag)
    Next b
    modSheetUtils.PrepareReport modConstants.SHEET_BRANCHSUM, "1"
    modSheetUtils.SetCell modConstants.SHEET_BRANCHSUM, 1, 1, "Branch summary " & modDateUtils.ToIso(d) & " (SIMULATION)", True
    modSheetUtils.WriteGrid modConstants.SHEET_BRANCHSUM, 3, 1, ToGrid(Array("Branch", "Name", "Customers", "Accounts", "Deposits", "Vault cash", _
        "Inter-branch position", "Fee income", "Interest expense", "Result"), rowsA)
    modSheetUtils.StyleReport modConstants.SHEET_BRANCHSUM, 3, 10, "3,4,5,6,7,8,9,10"
    modSheetUtils.PrepareReport modConstants.SHEET_CASHPOS, "1"
    modSheetUtils.SetCell modConstants.SHEET_CASHPOS, 1, 1, "Cash position " & modDateUtils.ToIso(d) & " (SIMULATION)", True
    modSheetUtils.WriteGrid modConstants.SHEET_CASHPOS, 3, 1, ToGrid(Array("Branch", "Name", "Vault cash", "VaultMin", "VaultMax", "Flag"), rowsC)
    modSheetUtils.StyleReport modConstants.SHEET_CASHPOS, 3, 6, "3,4,5"
End Sub

' ---- Deposit summary: by product and branch ----

Private Sub BuildDepositSummary(ByVal d As Date)
    Dim accts As Object, k As Variant, a As clsAccount, cnt As Object, bal As Object, keys As Object, rowsA As Collection
    Dim key As String, parts() As String
    Set accts = repoAccount.LoadAccounts()
    Set cnt = CreateObject("Scripting.Dictionary")
    Set bal = CreateObject("Scripting.Dictionary")
    For Each k In accts.Keys
        Set a = accts(k)
        If a.Status <> modConstants.STATUS_CLOSED Then
            key = a.ProductCode & "|" & a.BranchID
            If cnt.Exists(key) Then
                cnt(key) = cnt(key) + 1
                bal(key) = bal(key) + repoFxRate.ToBase(a.Balance, a.Ccy, d)
            Else
                cnt(key) = 1
                bal(key) = repoFxRate.ToBase(a.Balance, a.Ccy, d)
            End If
        End If
    Next k
    Set rowsA = New Collection
    For Each k In cnt.Keys
        parts = Split(CStr(k), "|")
        rowsA.Add Array(parts(0), parts(1), cnt(k), bal(k))
    Next k
    modSheetUtils.PrepareReport modConstants.SHEET_DEPSUM, "1,2"
    modSheetUtils.SetCell modConstants.SHEET_DEPSUM, 1, 1, "Deposits by product and branch " & modDateUtils.ToIso(d) & " (SIMULATION)", True
    modSheetUtils.WriteGrid modConstants.SHEET_DEPSUM, 3, 1, ToGrid(Array("Product", "Branch", "Accounts", "Balance"), rowsA)
    modSheetUtils.StyleReport modConstants.SHEET_DEPSUM, 3, 4, "3,4"
End Sub

' ---- Checks sheet ----

Private Function BuildChecks(ByVal d As Date) As Variant
    Dim res As Variant, rowsA As Collection, i As Long
    res = svcChecks.RunChecks()
    Set rowsA = New Collection
    For i = 1 To UBound(res, 1)
        rowsA.Add Array(res(i, 1), res(i, 2), res(i, 3), res(i, 4))
    Next i
    modSheetUtils.PrepareReport modConstants.SHEET_CHECKS, "4"
    modSheetUtils.SetCell modConstants.SHEET_CHECKS, 1, 1, "Control checks " & modDateUtils.ToIso(d) & " (SIMULATION)", True
    modSheetUtils.WriteGrid modConstants.SHEET_CHECKS, 3, 1, ToGrid(Array("#", "Check", "Result", "Detail"), rowsA)
    modSheetUtils.StyleReport modConstants.SHEET_CHECKS, 3, 4, ""
    BuildChecks = res
End Function

' Control: nothing needed. Rewrites every report sheet.
Public Function RunAll(ByVal user As clsUser, ByVal inp As Object) As clsResult
    Dim msg As String, d As Date, res As Variant, warns As Long, i As Long
    msg = svcSecurity.Denied(user, svcSecurity.ACT_REPORTS)
    If Len(msg) > 0 Then Set RunAll = modMessages.FailResult(msg): Exit Function
    repoConfig.Reload
    d = repoConfig.BusinessDate()
    BuildTrial d
    BuildBranchSummary d
    BuildDepositSummary d
    res = BuildChecks(d)
    svcReportsExtra.RunExtraReports d, inp
    For i = 1 To UBound(res, 1)
        If res(i, 3) = "WARN" Then warns = warns + 1
    Next i
    Set RunAll = modMessages.OkResult("Reports written. Control checks: " & svcChecks.FailCount(res) & " FAIL, " & warns & " WARN, " & _
        (UBound(res, 1) - svcChecks.FailCount(res) - warns) & " PASS. " & svcChecks.FailText(res), 1)
End Function

' ---- Statement: Control C5 = account, C16 to C17 = dates (blank = all) ----

Public Function Statement(ByVal user As clsUser, ByVal inp As Object) As clsResult
    Dim msg As String, a As clsAccount, accts As Object, txn As Variant, i As Long, rowsA As Collection, dt As Date
    Dim fromD As Date, toD As Date, hasFrom As Boolean, hasTo As Boolean, amt As Variant, drc As Variant, crc As Variant, accNo As String
    msg = svcSecurity.Denied(user, svcSecurity.ACT_STATEMENT)
    If Len(msg) > 0 Then Set Statement = modMessages.FailResult(msg): Exit Function
    accNo = CStr(inp(modConstants.INP_ACCOUNT))
    Set accts = repoAccount.LoadAccounts()
    If Not accts.Exists(accNo) Then Set Statement = modMessages.FailResult(modMessages.UnknownAccount(accNo)): Exit Function
    Set a = accts(accNo)
    If user.Role = modConstants.ROLE_TELLER Or user.Role = modConstants.ROLE_MANAGER Then
        If a.BranchID <> user.BranchID Then Set Statement = modMessages.FailResult("This account belongs to branch " & a.BranchID & "."): Exit Function
    End If
    hasFrom = modDateUtils.TryToDate(inp(modConstants.INP_FROM), fromD)
    hasTo = modDateUtils.TryToDate(inp(modConstants.INP_TO), toD)
    Set rowsA = New Collection
    txn = repoTransaction.LoadTransactions()
    If Not IsEmpty(txn) Then
        For i = 1 To UBound(txn, 1)
            If CStr(txn(i, 4)) = accNo And modDateUtils.TryToDate(txn(i, 3), dt) Then
                If (Not hasFrom Or dt >= fromD) And (Not hasTo Or dt <= toD) Then
                    amt = modMoney.ToDec(txn(i, 7))
                    drc = CDec(0)
                    crc = CDec(0)
                    If CStr(txn(i, 6)) = modConstants.DIR_DEBIT Then drc = amt Else crc = amt
                    rowsA.Add Array(dt, CStr(txn(i, 1)), CStr(txn(i, 5)), CStr(txn(i, 11)), drc, crc, modMoney.ToDec(txn(i, 8)), CStr(txn(i, 9)), CStr(txn(i, 14)))
                End If
            End If
        Next i
    End If
    modSheetUtils.PrepareReport modConstants.SHEET_STATEMENT, "2,8"
    modSheetUtils.SetCell modConstants.SHEET_STATEMENT, 1, 1, "Account statement (SIMULATION)", True
    modSheetUtils.SetCell modConstants.SHEET_STATEMENT, 2, 1, "Account " & accNo & "  customer " & a.CustomerID & "  product " & a.ProductCode & _
        "  branch " & a.BranchID & "  status " & a.Status & "  balance " & modMoney.FormatAmount(a.Balance)
    modSheetUtils.WriteGrid modConstants.SHEET_STATEMENT, 4, 1, ToGrid(Array("Date", "TxnID", "Type", "Description", "Debit", "Credit", "Balance after", "Counter", "Status"), rowsA)
    modSheetUtils.StyleReport modConstants.SHEET_STATEMENT, 4, 9, "5,6,7"
    Set Statement = modMessages.OkResult("Statement of " & accNo & " written to the Statement sheet (" & rowsA.Count & " rows).", rowsA.Count)
End Function
