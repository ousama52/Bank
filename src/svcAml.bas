Attribute VB_Name = "svcAml"
Option Explicit

' Stage 8: simulated anti money laundering. Rules from tblAmlRules, alerts in tblAmlAlerts, cases in tblAmlCases.
' No money is ever frozen or moved here.

Private Const RULE_LARGE As String = "LargeCash"
Private Const RULE_STRUCT As String = "Structuring"
Private Const RULE_VELOCITY As String = "Velocity"
Private Const RULE_NAME As String = "BlockedName"

Private Function Ready() As Boolean
    Ready = modSheetUtils.TableExists(modConstants.TBL_AMLALERTS) And modSheetUtils.TableExists(modConstants.TBL_AMLRULES) And modSheetUtils.TableExists(modConstants.TBL_AMLCASES)
End Function

' Finds an active rule by type. Returns its row values through the ByRef arguments.
Private Function FindRule(ByVal ruleType As String, ByRef ruleId As String, ByRef threshold As Variant, ByRef windowDays As Long, ByRef minCount As Long) As Boolean
    Dim data As Variant, i As Long
    data = modSheetUtils.ReadTable(modConstants.TBL_AMLRULES)
    If IsEmpty(data) Then Exit Function
    For i = 1 To UBound(data, 1)
        If StrComp(CStr(data(i, 2)), ruleType, vbTextCompare) = 0 And CStr(data(i, 6)) <> modConstants.FLAG_NO Then
            ruleId = CStr(data(i, 1))
            If Not modMoney.TryToDec(data(i, 3), threshold) Then threshold = CDec(0)
            windowDays = CLng(modMoney.ToDec(data(i, 4)))
            minCount = CLng(modMoney.ToDec(data(i, 5)))
            FindRule = True
            Exit Function
        End If
    Next i
End Function

Private Function RuleIdOf(ByVal ruleType As String) As String
    Dim id As String, th As Variant, w As Long, m As Long
    If FindRule(ruleType, id, th, w, m) Then RuleIdOf = id Else RuleIdOf = UCase$(ruleType)
End Function

Private Sub AddAlert(ByVal uow As clsUnitOfWork, ByVal ruleId As String, ByVal customerId As String, ByVal accountNo As String, ByVal amount As Variant, ByVal description As String)
    uow.AppendRow modConstants.TBL_AMLALERTS, Array(uow.NewId(modConstants.TBL_AMLALERTS, "AlertID", "AL", 6), uow.BusinessDate, ruleId, customerId, IIf(Len(accountNo) = 0, Empty, accountNo), amount, description, "New", Empty, Empty, Empty)
End Sub

' Called by services when a screening hit blocks or flags something (international payments, investments).
Public Sub OnScreeningHit(ByVal uow As clsUnitOfWork, ByVal source As String, ByVal name As String, ByVal customerId As String, ByVal accountNo As String, ByVal amount As Variant)
    If Not Ready() Then Exit Sub
    AddAlert uow, RuleIdOf(RULE_NAME), customerId, accountNo, amount, "Name screening hit: " & name & " (" & source & ")"
End Sub

' Screens a customer name at account opening or loan application. Returns "" to go on, or the refusal text.
' A HeadOffice user goes on by starting the note (C15) with "OVERRIDE:"; the hit is still written as an alert.
Public Function ScreenName(ByVal uow As clsUnitOfWork, ByVal user As clsUser, ByVal name As String, ByVal customerId As String, ByVal source As String, ByVal note As String) As String
    If Not svcIntl.Screened(name) Then Exit Function
    If user.Role = modConstants.ROLE_HEADOFFICE And UCase$(Left$(Trim$(note), 9)) = "OVERRIDE:" And Len(Trim$(Mid$(Trim$(note), 10))) > 0 Then
        OnScreeningHit uow, source & " approved by " & user.UserID & ": " & Trim$(Mid$(Trim$(note), 10)), name, customerId, "", CDec(0)
        Exit Function
    End If
    clsLogger.Warn "svcAml.ScreenName", "Blocked name hit for " & customerId & " in " & source & "."
    ScreenName = "The customer name matches the blocked-name list. A HeadOffice user may go on by putting OVERRIDE: and a reason in Control C15."
End Function

' ---- Daily scan ----

Private Function HasAlert(ByVal alerts As Variant, ByVal ruleId As String, ByVal customerId As String, ByVal description As String) As Boolean
    Dim i As Long
    If IsEmpty(alerts) Then Exit Function
    For i = 1 To UBound(alerts, 1)
        If CStr(alerts(i, repoGov.ALC_RULE)) = ruleId And CStr(alerts(i, repoGov.ALC_CUSTOMER)) = customerId And CStr(alerts(i, 7)) = description Then HasAlert = True: Exit Function
    Next i
End Function

Private Function IsCashType(ByVal t As String) As Boolean
    IsCashType = (t = modConstants.TXN_DEPOSIT Or t = modConstants.TXN_WITHDRAWAL Or InStr(1, t, "Cash", vbTextCompare) > 0 Or InStr(1, t, "Agent", vbTextCompare) > 0)
End Function

' Scans the committed transactions of the business date d (and earlier days for structuring). Returns the number of new alerts.
' Called after the day's work is committed, with its own unit of work.
Public Function DailyScan(ByVal uow As clsUnitOfWork, ByVal d As Date) As Long
    Dim tx As Variant, alerts As Variant, i As Long, n As Long, dt As Date
    Dim lcId As String, lcTh As Variant, w As Long, m As Long, stId As String, stTh As Variant, stW As Long, stM As Long, vId As String, vTh As Variant
    Dim cust As Object, acc As clsAccount, cid As String, amt As Variant, typ As String, desc As String, creditDay As Object, k As Variant, c As clsCustomer
    Dim cnt As Object, ids As Object, lowBand As Variant
    If Not Ready() Then Exit Function
    tx = repoTransaction.LoadTransactions()
    If IsEmpty(tx) Then Exit Function
    alerts = uow.WorkingTable(modConstants.TBL_AMLALERTS)
    Set cust = repoCustomer.LoadCustomers()
    Set creditDay = CreateObject("Scripting.Dictionary")
    Set cnt = CreateObject("Scripting.Dictionary")
    Set ids = CreateObject("Scripting.Dictionary")
    Dim haveLc As Boolean, haveSt As Boolean, haveV As Boolean
    haveLc = FindRule(RULE_LARGE, lcId, lcTh, w, m)
    haveSt = FindRule(RULE_STRUCT, stId, stTh, stW, stM)
    haveV = FindRule(RULE_VELOCITY, vId, vTh, w, m)
    lowBand = stTh * CDec(0.8)
    For i = 1 To UBound(tx, 1)
        If CStr(tx(i, 14)) <> modConstants.STATUS_REVERSED And Len(CStr(tx(i, 1))) > 0 And modDateUtils.TryToDate(tx(i, 3), dt) Then
            Set acc = uow.GetAccount(CStr(tx(i, 4)))
            If Not acc Is Nothing Then
                cid = acc.CustomerID
                amt = modMoney.ToDec(tx(i, 7))
                typ = CStr(tx(i, 5))
                If dt = d Then
                    If haveLc And IsCashType(typ) And amt >= lcTh Then
                        desc = "Cash " & typ & " " & modMoney.FormatAmount(amt) & " on " & CStr(tx(i, 4)) & " (" & CStr(tx(i, 1)) & ")"
                        If Not HasAlert(alerts, lcId, cid, desc) Then AddAlert uow, lcId, cid, CStr(tx(i, 4)), amt, desc: n = n + 1
                    End If
                    If haveV And CStr(tx(i, 6)) = modConstants.DIR_CREDIT Then
                        If creditDay.Exists(cid) Then creditDay(cid) = creditDay(cid) + amt Else creditDay(cid) = amt
                    End If
                End If
                If haveSt And typ = modConstants.TXN_DEPOSIT And dt > d - stW - 0 And dt <= d And amt >= lowBand And amt < stTh Then
                    If cnt.Exists(cid) Then cnt(cid) = cnt(cid) + 1: ids(cid) = ids(cid) & "," & CStr(tx(i, 1)) Else cnt(cid) = 1: ids(cid) = CStr(tx(i, 1))
                End If
            End If
        End If
    Next i
    If haveSt Then
        For Each k In cnt.Keys
            If cnt(k) >= stM Then
                desc = "Possible structuring: " & cnt(k) & " cash deposits in " & stW & " days (" & ids(k) & ")"
                If Not HasAlert(alerts, stId, CStr(k), desc) And Not StructuringRecent(alerts, stId, CStr(k), d, stW) Then AddAlert uow, stId, CStr(k), "", CDec(0), desc: n = n + 1
            End If
        Next k
    End If
    If haveV Then
        For Each k In creditDay.Keys
            If cust.Exists(CStr(k)) Then
                Set c = cust(CStr(k))
                If c.KYCLevel <= 2 And creditDay(k) >= vTh Then
                    desc = "Credits of " & modMoney.FormatAmount(creditDay(k)) & " in one day for a KYC " & c.KYCLevel & " customer on " & modDateUtils.ToIso(d)
                    If Not HasAlert(alerts, vId, CStr(k), desc) Then AddAlert uow, vId, CStr(k), "", creditDay(k), desc: n = n + 1
                End If
            End If
        Next k
    End If
    DailyScan = n
End Function

' A structuring alert for the same customer within the window means the same set is still being seen.
Private Function StructuringRecent(ByVal alerts As Variant, ByVal ruleId As String, ByVal customerId As String, ByVal d As Date, ByVal windowDays As Long) As Boolean
    Dim i As Long, dt As Date
    If IsEmpty(alerts) Then Exit Function
    For i = 1 To UBound(alerts, 1)
        If CStr(alerts(i, repoGov.ALC_RULE)) = ruleId And CStr(alerts(i, repoGov.ALC_CUSTOMER)) = customerId And modDateUtils.TryToDate(alerts(i, repoGov.ALC_DATE), dt) Then
            If dt >= d - windowDays Then StructuringRecent = True: Exit Function
        End If
    Next i
End Function

' ---- Review, cases ----

Private Function OpenCaseRow(ByVal uow As clsUnitOfWork, ByVal customerId As String, ByVal summary As String, ByVal assignedTo As String) As String
    Dim id As String
    id = uow.NewId(modConstants.TBL_AMLCASES, "CaseID", "CS", 6)
    uow.AppendRow modConstants.TBL_AMLCASES, Array(id, uow.BusinessDate, customerId, summary, "Open", assignedTo, Empty, Empty, Empty)
    OpenCaseRow = id
End Function

' Control: C49 alert, C15 decision (ClosedNoAction or Escalated). Escalating opens a case assigned to the reviewer.
Public Function ReviewAlert(ByVal user As clsUser, ByVal inp As Object) As clsResult
    Dim uow As clsUnitOfWork, msg As String, alerts As Variant, ar As Long, decision As String, cs As String
    msg = svcSecurity.Denied(user, svcSecurity.ACT_AML)
    If Len(msg) = 0 And Not Ready() Then msg = modMessages.MSG_SETUP_FIRST
    If Len(msg) > 0 Then Set ReviewAlert = modMessages.FailResult(msg): Exit Function
    decision = StrConv(Trim$(CStr(inp(modConstants.INP_NOTE))), vbProperCase)
    If decision = "Closednoaction" Then decision = "ClosedNoAction"
    If decision <> "ClosedNoAction" And decision <> "Escalated" Then Set ReviewAlert = modMessages.FailResult("Put ClosedNoAction or Escalated in Control C15."): Exit Function
    Set uow = New clsUnitOfWork
    uow.Init repoConfig.BusinessDate(), user.UserID
    alerts = uow.WorkingTable(modConstants.TBL_AMLALERTS)
    ar = repoGov.FindGovById(alerts, CStr(inp(modConstants.INP_ALERT)))
    If ar = 0 Then Set ReviewAlert = modMessages.FailResult("Unknown alert (Control C49)."): Exit Function
    If CStr(alerts(ar, repoGov.ALC_STATUS)) = "ClosedNoAction" Or CStr(alerts(ar, repoGov.ALC_STATUS)) = "Escalated" Then Set ReviewAlert = modMessages.FailResult("Alert " & CStr(alerts(ar, 1)) & " is already " & CStr(alerts(ar, repoGov.ALC_STATUS)) & "."): Exit Function
    alerts(ar, repoGov.ALC_STATUS) = decision
    alerts(ar, repoGov.ALC_REVIEWEDBY) = user.UserID
    alerts(ar, repoGov.ALC_REVIEWEDDATE) = uow.BusinessDate
    If decision = "Escalated" Then
        cs = OpenCaseRow(uow, CStr(alerts(ar, repoGov.ALC_CUSTOMER)), CStr(alerts(ar, 7)), user.UserID)
        alerts(ar, repoGov.ALC_CASE) = cs
    End If
    uow.ReplaceTable modConstants.TBL_AMLALERTS, alerts
    repoAudit.Add uow, "AmlAlertReviewed", CStr(alerts(ar, 1)), decision & IIf(Len(cs) > 0, " case " & cs, "")
    uow.Commit
    Set ReviewAlert = modMessages.OkResult("Alert " & CStr(alerts(ar, 1)) & " is " & decision & IIf(Len(cs) > 0, ". Case " & cs & " opened.", "."), 1, cs)
End Function

' Control: C49 alert (any alert not yet in a case), C15 summary. Marks the alert Escalated.
Public Function OpenCase(ByVal user As clsUser, ByVal inp As Object) As clsResult
    Dim uow As clsUnitOfWork, msg As String, alerts As Variant, ar As Long, cs As String
    msg = svcSecurity.Denied(user, svcSecurity.ACT_AML)
    If Len(msg) = 0 And Not Ready() Then msg = modMessages.MSG_SETUP_FIRST
    If Len(msg) > 0 Then Set OpenCase = modMessages.FailResult(msg): Exit Function
    Set uow = New clsUnitOfWork
    uow.Init repoConfig.BusinessDate(), user.UserID
    alerts = uow.WorkingTable(modConstants.TBL_AMLALERTS)
    ar = repoGov.FindGovById(alerts, CStr(inp(modConstants.INP_ALERT)))
    If ar = 0 Then Set OpenCase = modMessages.FailResult("Unknown alert (Control C49)."): Exit Function
    If Len(CStr(alerts(ar, repoGov.ALC_CASE))) > 0 Then Set OpenCase = modMessages.FailResult("Alert " & CStr(alerts(ar, 1)) & " already has case " & CStr(alerts(ar, repoGov.ALC_CASE)) & "."): Exit Function
    If CStr(alerts(ar, repoGov.ALC_STATUS)) = "ClosedNoAction" Then Set OpenCase = modMessages.FailResult("The alert was closed without action."): Exit Function
    msg = Trim$(CStr(inp(modConstants.INP_NOTE)))
    If Len(msg) = 0 Then msg = CStr(alerts(ar, 7))
    cs = OpenCaseRow(uow, CStr(alerts(ar, repoGov.ALC_CUSTOMER)), msg, user.UserID)
    alerts(ar, repoGov.ALC_CASE) = cs
    alerts(ar, repoGov.ALC_STATUS) = "Escalated"
    alerts(ar, repoGov.ALC_REVIEWEDBY) = user.UserID
    alerts(ar, repoGov.ALC_REVIEWEDDATE) = uow.BusinessDate
    uow.ReplaceTable modConstants.TBL_AMLALERTS, alerts
    repoAudit.Add uow, "AmlCaseOpened", cs, "From alert " & CStr(alerts(ar, 1))
    uow.Commit
    Set OpenCase = modMessages.OkResult("Case " & cs & " opened for alert " & CStr(alerts(ar, 1)) & ".", 1, cs)
End Function

' Control: C12 case, C15 decision text. Only an Open or Reported case can be closed.
Public Function CloseCase(ByVal user As clsUser, ByVal inp As Object) As clsResult
    Dim uow As clsUnitOfWork, msg As String, cases As Variant, cr As Long
    msg = svcSecurity.Denied(user, svcSecurity.ACT_AML)
    If Len(msg) = 0 And Not Ready() Then msg = modMessages.MSG_SETUP_FIRST
    If Len(msg) > 0 Then Set CloseCase = modMessages.FailResult(msg): Exit Function
    Set uow = New clsUnitOfWork
    uow.Init repoConfig.BusinessDate(), user.UserID
    cases = uow.WorkingTable(modConstants.TBL_AMLCASES)
    cr = repoGov.FindGovById(cases, CStr(inp(modConstants.INP_REFID)))
    If cr = 0 Then Set CloseCase = modMessages.FailResult("Unknown case (put the case ID in Control C12)."): Exit Function
    If CStr(cases(cr, repoGov.CSC_STATUS)) = "Closed" Then Set CloseCase = modMessages.FailResult("The case is already Closed."): Exit Function
    If Len(Trim$(CStr(inp(modConstants.INP_NOTE)))) = 0 Then Set CloseCase = modMessages.FailResult("Write the decision in Control C15."): Exit Function
    cases(cr, repoGov.CSC_STATUS) = "Closed"
    cases(cr, repoGov.CSC_DECISION) = Trim$(CStr(inp(modConstants.INP_NOTE)))
    cases(cr, repoGov.CSC_CLOSED) = uow.BusinessDate
    uow.ReplaceTable modConstants.TBL_AMLCASES, cases
    repoAudit.Add uow, "AmlCaseClosed", CStr(cases(cr, 1)), CStr(cases(cr, repoGov.CSC_DECISION))
    uow.Commit
    Set CloseCase = modMessages.OkResult("Case " & CStr(cases(cr, 1)) & " closed.", 1)
End Function

' Control: C12 case. Needs approval (ReportCase) from a different HeadOffice user.
Public Function ReportCase(ByVal user As clsUser, ByVal inp As Object) As clsResult
    Dim uow As clsUnitOfWork, msg As String, cases As Variant, cr As Long, args As Object, r As clsResult
    msg = svcSecurity.Denied(user, svcSecurity.ACT_AML)
    If Len(msg) = 0 And Not Ready() Then msg = modMessages.MSG_SETUP_FIRST
    If Len(msg) > 0 Then Set ReportCase = modMessages.FailResult(msg): Exit Function
    Set uow = New clsUnitOfWork
    uow.Init repoConfig.BusinessDate(), user.UserID
    cases = uow.WorkingTable(modConstants.TBL_AMLCASES)
    cr = repoGov.FindGovById(cases, CStr(inp(modConstants.INP_REFID)))
    If cr = 0 Then Set ReportCase = modMessages.FailResult("Unknown case (put the case ID in Control C12)."): Exit Function
    If CStr(cases(cr, repoGov.CSC_STATUS)) <> "Open" Then Set ReportCase = modMessages.FailResult("Case " & CStr(cases(cr, 1)) & " is " & CStr(cases(cr, repoGov.CSC_STATUS)) & "."): Exit Function
    Set args = modArgs.NewArgs()
    args("case") = CStr(cases(cr, 1))
    Set r = svcApproval.Request(uow, modConstants.REQ_REPORT_CASE, "", "", CDec(0), modConstants.HEAD_OFFICE_ID, args, user.UserID)
    uow.Commit
    Set ReportCase = modMessages.OkResult("Report of case " & CStr(cases(cr, 1)) & " requested. " & r.Message, 1, r.RefID)
End Function

' Run on approval.
Public Function ExecReportCase(ByVal uow As clsUnitOfWork, ByVal args As Object, ByVal createdBy As String, ByVal approvedBy As String) As clsResult
    Dim cases As Variant, cr As Long
    cases = uow.WorkingTable(modConstants.TBL_AMLCASES)
    cr = repoGov.FindGovById(cases, modArgs.Text(args, "case"))
    If cr = 0 Then Set ExecReportCase = modMessages.FailResult("Unknown case."): Exit Function
    If CStr(cases(cr, repoGov.CSC_STATUS)) <> "Open" Then Set ExecReportCase = modMessages.FailResult("Case " & CStr(cases(cr, 1)) & " is " & CStr(cases(cr, repoGov.CSC_STATUS)) & "."): Exit Function
    cases(cr, repoGov.CSC_STATUS) = "Reported"
    uow.ReplaceTable modConstants.TBL_AMLCASES, cases
    repoAudit.Add uow, "AmlCaseReported", CStr(cases(cr, 1)), "Reported, approved by " & approvedBy
    Set ExecReportCase = modMessages.OkResult("Case " & CStr(cases(cr, 1)) & " is Reported.", 1, "")
End Function
