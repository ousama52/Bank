Attribute VB_Name = "svcYearEnd"
Option Explicit

' Stage 8: year-end close. Per branch one journal (type YearEndClose) that zeroes every income and expense account of the year
' and moves the result to 3100. Closing journals are left out of every profit and loss calculation (see repoJournal.LoadGLNetNoClose).

Private Function YearFromInput(ByVal inp As Object, ByVal d As Date) As Long
    Dim t As String
    YearFromInput = Year(d) - 1
    t = Trim$(CStr(inp(modConstants.INP_PERIOD)))
    If t Like "####*" Then YearFromInput = CLng(Left$(t, 4))
End Function

' Journals dated in a closed year cannot be reversed, and closing journals are never reversed.
Public Function ClosedYearProblem(ByVal postDate As Date, ByVal journalType As String) As String
    If journalType = modConstants.JT_YEAR_END Then ClosedYearProblem = "A year-end closing journal cannot be reversed.": Exit Function
    If Year(postDate) <= repoGov.LastClosedYear() Then ClosedYearProblem = "The journal belongs to the closed year " & Year(postDate) & " and cannot be reversed."
End Function

' Why the close of year yr is not allowed now ("" = allowed). Warnings go to the Log.
Public Function PreCheck(ByVal yr As Long, ByVal d As Date) As String
    Dim last As Long, res As Variant, pbt As Variant, adj As Variant, fines As Variant, taxable As Variant, tax As Variant, msg As String, m As Long, have As Object, data As Variant, i As Long
    last = repoGov.LastClosedYear()
    If d <= DateSerial(yr, 12, 31) Then PreCheck = "Year " & yr & " is not finished. The end of day of 31 December must be done first.": Exit Function
    If last <> 0 And last <> yr - 1 Then PreCheck = "The last closed year is " & last & ". Close the years in order.": Exit Function
    If last >= yr Then PreCheck = "Year " & yr & " is already closed.": Exit Function
    res = svcChecks.RunChecks()
    If svcChecks.FailCount(res) > 0 Then PreCheck = "Control checks fail: " & svcChecks.FailText(res): Exit Function
    If Not svcTaxReturn.ComputeTax(yr, d, pbt, adj, fines, taxable, tax, msg) Then PreCheck = msg: Exit Function
    If tax > 0 And Not svcTaxReturn.AlreadyBooked(yr) Then PreCheck = "Book the income tax of " & yr & " first (RunBookIncomeTax).": Exit Function
    If modSheetUtils.TableExists(modConstants.TBL_PAYROLLRUNS) Then
        Set have = CreateObject("Scripting.Dictionary")
        data = modSheetUtils.ReadTable(modConstants.TBL_PAYROLLRUNS)
        If Not IsEmpty(data) Then
            For i = 1 To UBound(data, 1)
                have(CStr(data(i, 2))) = True
            Next i
        End If
        For m = 1 To 12
            If Not have.Exists(CStr(yr) & Right$("0" & m, 2)) Then clsLogger.Warn "svcYearEnd.PreCheck", "No payroll run for " & yr & Right$("0" & m, 2) & "."
        Next m
    End If
End Function

' Control: C25 year (blank = last year). Admin. Creates the YearEndClose approval.
Public Function RequestYearEnd(ByVal user As clsUser, ByVal inp As Object) As clsResult
    Dim uow As clsUnitOfWork, msg As String, d As Date, yr As Long, args As Object, r As clsResult
    msg = svcSecurity.Denied(user, svcSecurity.ACT_YEAREND)
    If Len(msg) = 0 And Not repoConfig.Has(modConstants.CFG_LAST_CLOSED) Then msg = modMessages.MSG_SETUP_FIRST
    If Len(msg) > 0 Then Set RequestYearEnd = modMessages.FailResult(msg): Exit Function
    d = repoConfig.BusinessDate()
    yr = YearFromInput(inp, d)
    msg = PreCheck(yr, d)
    If Len(msg) > 0 Then Set RequestYearEnd = modMessages.FailResult(msg): Exit Function
    Set uow = New clsUnitOfWork
    uow.Init d, user.UserID
    Set args = modArgs.NewArgs()
    args("year") = CStr(yr)
    Set r = svcApproval.Request(uow, modConstants.REQ_YEAR_END, "", "", CDec(0), modConstants.HEAD_OFFICE_ID, args, user.UserID)
    uow.Commit
    Set RequestYearEnd = modMessages.OkResult("Year-end close of " & yr & " requested. " & r.Message, 1, r.RefID)
End Function

' Run on approval. One journal per branch.
Public Function ExecYearEnd(ByVal uow As clsUnitOfWork, ByVal args As Object, ByVal createdBy As String, ByVal approvedBy As String) As clsResult
    Dim yr As Long, msg As String, lines As Variant, chart As Object, net As Object, i As Long, dt As Date, gl As String, key As String, k As Variant, br As Object, parts() As String
    Dim jl As Collection, jr As clsResult, total As Variant, profit As Variant, n As Long, pbt As Variant, adj As Variant, fines As Variant, taxable As Variant, tax As Variant, income As Variant, expense As Variant, cr As Object
    Dim cfg As Variant, found As Boolean, v As Variant, taxBooked As Variant, b As Variant
    yr = CLng(modArgs.Dec(args, "year"))
    msg = PreCheck(yr, uow.BusinessDate)
    If Len(msg) > 0 Then Set ExecYearEnd = modMessages.FailResult(msg): Exit Function
    Set chart = repoChart.LoadChart()
    Set net = CreateObject("Scripting.Dictionary")
    lines = repoJournal.LoadDatedLines()
    If Not IsEmpty(lines) Then
        For i = 1 To UBound(lines, 1)
            If Not IsEmpty(lines(i, 1)) Then
                dt = lines(i, 1)
                gl = CStr(lines(i, 3))
                If Year(dt) = yr And CStr(lines(i, 6)) <> modConstants.JT_YEAR_END And repoGov.IsProfitGL(chart, gl) Then
                    key = CStr(lines(i, 2)) & "|" & gl
                    v = lines(i, 4) - lines(i, 5)
                    If net.Exists(key) Then net(key) = net(key) + v Else net(key) = v
                End If
            End If
        Next i
    End If
    Set br = CreateObject("Scripting.Dictionary")
    For Each k In net.Keys
        If net(k) <> 0 Then
            parts = Split(CStr(k), "|")
            br(parts(0)) = True
        End If
    Next k
    taxBooked = CDec(0)
    For Each b In br.Keys
        Set jl = New Collection
        total = CDec(0)
        For Each k In net.Keys
            parts = Split(CStr(k), "|")
            If parts(0) = CStr(b) And net(k) <> 0 Then
                If net(k) > 0 Then svcLedger.AddCr jl, CStr(b), parts(1), net(k) Else svcLedger.AddDr jl, CStr(b), parts(1), -net(k)
                total = total + net(k)
                If parts(1) = modConstants.GL_INCOME_TAX_EXP Then taxBooked = taxBooked + net(k)
            End If
        Next k
        profit = -total
        If profit > 0 Then svcLedger.AddCr jl, CStr(b), modConstants.GL_RETAINED, profit
        If profit < 0 Then svcLedger.AddDr jl, CStr(b), modConstants.GL_RETAINED, -profit
        Set jr = svcLedger.PostJournal(uow, modConstants.JT_YEAR_END, "Year-end close " & yr & " " & CStr(b), uow.NewGroupId(), jl)
        If Not jr.Success Then Set ExecYearEnd = jr: Exit Function
        n = n + 1
    Next b
    cfg = uow.WorkingTable(modConstants.TBL_CONFIG)
    For i = 1 To UBound(cfg, 1)
        If StrComp(CStr(cfg(i, 1)), modConstants.CFG_LAST_CLOSED, vbTextCompare) = 0 Then cfg(i, 2) = CStr(yr): found = True: Exit For
    Next i
    If found Then uow.ReplaceTable modConstants.TBL_CONFIG, cfg Else uow.AppendRow modConstants.TBL_CONFIG, Array(modConstants.CFG_LAST_CLOSED, CStr(yr))
    YearFigures4 yr, income, expense
    repoAudit.Add uow, "YearEndClose", CStr(yr), "Profit before tax " & modMoney.FormatAmount(income - expense) & ", income tax " & modMoney.FormatAmount(taxBooked) & ", profit after tax " & modMoney.FormatAmount(income - expense - taxBooked) & ", " & n & " branch journal(s), approved by " & approvedBy
    Set ExecYearEnd = modMessages.OkResult("Year " & yr & " closed with " & n & " branch journal(s).", n, "")
End Function

Private Sub YearFigures4(ByVal yr As Long, ByRef income As Variant, ByRef expense As Variant)
    Dim fines As Variant, cr As Object
    svcTaxReturn.YearFigures yr, income, expense, fines, cr
End Sub

' Check 28: for each closed year, income and expense accounts net to 0 per branch counting the closing journals, and 3100 rose by the profit after tax.
Public Sub CheckClosedYears(ByVal rowsList As Collection)
    Dim hdr As Variant, lines As Variant, closeYear As Object, i As Long, jid As String, yr As Long, chart As Object, net As Object, key As String, k As Variant, v As Variant
    Dim dates As Object, bad As Long, info As String, y As Variant, closing3100 As Object, p As Long, dt As Date, d As String, ret As Variant, profitAfter As Object, gl As String, typ As String
    hdr = repoJournal.LoadHeaders()
    lines = repoJournal.LoadLines()
    Set closeYear = CreateObject("Scripting.Dictionary")
    Set dates = CreateObject("Scripting.Dictionary")
    Set net = CreateObject("Scripting.Dictionary")
    Set closing3100 = CreateObject("Scripting.Dictionary")
    Set profitAfter = CreateObject("Scripting.Dictionary")
    If IsEmpty(hdr) Or IsEmpty(lines) Then
        svcChecks.AddRow rowsList, 28, "Closed years are consistent", "PASS", "No closed year"
        Exit Sub
    End If
    For i = 1 To UBound(hdr, 1)
        jid = CStr(hdr(i, repoJournal.JNC_ID))
        If modDateUtils.TryToDate(hdr(i, 2), dt) Then dates(jid) = dt
        If CStr(hdr(i, repoJournal.JNC_TYPE)) = modConstants.JT_YEAR_END Then
            d = CStr(hdr(i, 3))
            p = InStr(d, "Year-end close ")
            If p = 1 Then
                yr = CLng(Mid$(d, 16, 4))
                closeYear(jid) = yr
            End If
        End If
    Next i
    If closeYear.Count = 0 Then
        svcChecks.AddRow rowsList, 28, "Closed years are consistent", "PASS", "No closed year"
        Exit Sub
    End If
    Set chart = repoChart.LoadChart()
    For i = 1 To UBound(lines, 1)
        jid = CStr(lines(i, repoJournal.LNC_JOURNAL))
        gl = CStr(lines(i, repoJournal.LNC_GL))
        v = modMoney.ToDec(lines(i, repoJournal.LNC_DEBIT)) - modMoney.ToDec(lines(i, repoJournal.LNC_CREDIT))
        If closeYear.Exists(jid) Then
            yr = closeYear(jid)
            If repoGov.IsProfitGL(chart, gl) Then
                key = yr & "|" & CStr(lines(i, repoJournal.LNC_BRANCH)) & "|" & gl
                If net.Exists(key) Then net(key) = net(key) + v Else net(key) = v
            ElseIf gl = modConstants.GL_RETAINED Then
                If closing3100.Exists(CStr(yr)) Then closing3100(CStr(yr)) = closing3100(CStr(yr)) + v Else closing3100(CStr(yr)) = v
            End If
        ElseIf dates.Exists(jid) And repoGov.IsProfitGL(chart, gl) Then
            yr = Year(dates(jid))
            If closeYearExists(closeYear, yr) Then
                key = yr & "|" & CStr(lines(i, repoJournal.LNC_BRANCH)) & "|" & gl
                If net.Exists(key) Then net(key) = net(key) + v Else net(key) = v
                If closing3100.Exists("P" & yr) Then closing3100("P" & yr) = closing3100("P" & yr) + v Else closing3100("P" & yr) = v
            End If
        End If
    Next i
    For Each k In net.Keys
        If net(k) <> 0 Then bad = bad + 1: info = info & CStr(k) & " net " & modMoney.FormatAmount(net(k)) & "; "
    Next k
    For Each y In closeYear.Items
        If closing3100.Exists(CStr(y)) And closing3100.Exists("P" & CStr(y)) Then
            If closing3100(CStr(y)) <> closing3100("P" & CStr(y)) Then bad = bad + 1: info = info & "3100 for " & CStr(y) & " is not the year profit; "
        End If
    Next y
    If bad = 0 Then
        svcChecks.AddRow rowsList, 28, "Closed years are consistent", "PASS", "Closed year(s) net to 0 per branch"
    Else
        svcChecks.AddRow rowsList, 28, "Closed years are consistent", "FAIL", Left$(info, 400)
    End If
End Sub

Private Function closeYearExists(ByVal closeYear As Object, ByVal yr As Long) As Boolean
    Dim v As Variant
    For Each v In closeYear.Items
        If CLng(v) = yr Then closeYearExists = True: Exit Function
    Next v
End Function
