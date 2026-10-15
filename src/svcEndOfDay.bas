Attribute VB_Name = "svcEndOfDay"
Option Explicit

' End of day for the current business date D (Admin only). Everything is built in one unit of work and committed once.
' The control checks run on the committed data; if any FAILs, every table touched is restored from a snapshot and the date does not move.

Private Function TouchedTables() As Variant
    Dim base As Variant, extra As Variant, out() As Variant, i As Long, n As Long
    base = Array(modConstants.TBL_ACCOUNTS, modConstants.TBL_TRANSACTIONS, modConstants.TBL_JOURNAL, modConstants.TBL_JLINES, _
                 modConstants.TBL_INTERBANK, modConstants.TBL_AUDIT)
    extra = svcEndOfDayExtra.ExtraTables()
    n = (UBound(base) + 1) + (UBound(extra) + 1)
    ReDim out(0 To n - 1)
    For i = 0 To UBound(base)
        out(i) = base(i)
    Next i
    For i = 0 To UBound(extra)
        out(UBound(base) + 1 + i) = extra(i)
    Next i
    TouchedTables = out
End Function

Private Function Snapshot() As Object
    Dim d As Object, t As Variant
    Set d = CreateObject("Scripting.Dictionary")
    For Each t In TouchedTables()
        d(CStr(t)) = modSheetUtils.ReadTable(CStr(t))
    Next t
    Set Snapshot = d
End Function

Private Sub Restore(ByVal snap As Object)
    Dim k As Variant
    For Each k In snap.Keys
        modSheetUtils.WriteTable CStr(k), snap(k)
    Next k
End Sub

Private Function Fail(ByVal msg As String) As clsResult
    clsLogger.Fail "svcEndOfDay.CloseDay", msg
    Set Fail = modMessages.FailResult(msg & " Nothing was changed and the date was not advanced.")
End Function

' Steps 3 to 7 of the end of day (interbank, fixed deposits, accrual, dormancy, month end). Fills stats.
Private Function BuildDay(ByVal uow As clsUnitOfWork, ByVal d As Date, ByVal stats As Object) As clsResult
    Dim r As clsResult, decided As Long, matured As Long, accrued As Long, total As Variant, dormant As Long, posted As Long, charged As Long
    Set r = svcInterbank.DecideDue(uow, d, decided)
    If Not r.Success Then Set BuildDay = r: Exit Function
    Set r = svcFixedDeposit.Mature(uow, d, matured)
    If Not r.Success Then Set BuildDay = r: Exit Function
    Set r = svcEndOfDayExtra.DailySteps(uow, d, stats)
    If Not r.Success Then Set BuildDay = r: Exit Function
    Set r = svcInterest.Accrue(uow, d, accrued, total)
    If Not r.Success Then Set BuildDay = r: Exit Function
    dormant = svcAccounts.MarkDormant(uow, d)
    stats("InterbankSettled") = decided
    stats("FDMatured") = matured
    stats("AccountsAccrued") = accrued
    stats("InterestAccrued") = total
    stats("DormantMarked") = dormant
    If modDateUtils.IsLastDayOfMonth(d) Then
        Set r = svcInterest.PostMonthEnd(uow, d, posted)
        If Not r.Success Then Set BuildDay = r: Exit Function
        Set r = svcInterest.ChargeMaintenance(uow, d, charged)
        If Not r.Success Then Set BuildDay = r: Exit Function
        Set r = svcEndOfDayExtra.MonthEndSteps(uow, d, stats)
        If Not r.Success Then Set BuildDay = r: Exit Function
    End If
    Set r = svcEndOfDayExtra.FinalSteps(uow, d, stats)
    If Not r.Success Then Set BuildDay = r: Exit Function
    Set BuildDay = modMessages.OkResult("", 1)
End Function

Public Function CloseDay(ByVal user As clsUser) As clsResult
    Dim msg As String, d As Date, pending As Long, snap As Object, uow As clsUnitOfWork, uow2 As clsUnitOfWork, stats As Object
    Dim r As clsResult, checks As Variant, notes As String, errNum As Long, errDesc As String, committed As Boolean
    msg = svcSecurity.Denied(user, svcSecurity.ACT_ENDOFDAY)
    If Len(msg) > 0 Then Set CloseDay = modMessages.FailResult(msg): Exit Function
    repoConfig.Reload
    d = repoConfig.BusinessDate()
    If repoClose.Exists(d) Then Set CloseDay = modMessages.FailResult(modMessages.EndOfDayAlready(d)): Exit Function
    pending = repoApproval.CountPending()
    If pending > 0 Then
        notes = modMessages.PendingWarning(pending)
        clsLogger.Warn "svcEndOfDay.CloseDay", notes
    End If
    Set snap = Snapshot()
    On Error GoTo EH
    Randomize
    Set uow = New clsUnitOfWork
    uow.Init d, user.UserID
    Set stats = modArgs.NewArgs()
    Set r = BuildDay(uow, d, stats)
    If Not r.Success Then Set CloseDay = Fail("End of day stopped: " & r.Message): Exit Function
    uow.Commit
    committed = True
    Set r = svcEndOfDayExtra.PostCommitSteps(user, d, stats)
    If Not r.Success Then Restore snap: Set CloseDay = Fail(r.Message): Exit Function
    svcChecksFx.StrictEod = True
    checks = svcChecks.RunChecks()
    svcChecksFx.StrictEod = False
    If svcChecks.FailCount(checks) > 0 Then
        Restore snap
        Set CloseDay = Fail(modMessages.MSG_EOD_ROLLED_BACK & " " & svcChecks.FailText(checks))
        Exit Function
    End If
    Set uow2 = New clsUnitOfWork
    uow2.Init d, user.UserID
    uow2.AppendRow modConstants.TBL_DAILYCLOSE, repoClose.NewCloseRow(d, stats, True, notes)
    repoAudit.Add uow2, "EndOfDay", modDateUtils.ToIso(d), "Business day closed. " & notes
    uow2.Commit
    repoConfig.SetBusinessDate d + 1
    Set CloseDay = modMessages.OkResult("End of day " & modDateUtils.ToIso(d) & " finished. Business date is now " & modDateUtils.ToIso(d + 1) & _
        ". Accrued " & modMoney.FormatAmount(stats("InterestAccrued")) & " MMK on " & stats("AccountsAccrued") & " accounts; interbank decided " & stats("InterbankSettled") & _
        "; fixed deposits matured " & stats("FDMatured") & "; dormant " & stats("DormantMarked") & ". " & notes, 1)
    Exit Function
EH:
    errNum = Err.Number
    errDesc = Err.Description
    On Error Resume Next
    svcChecksFx.StrictEod = False
    If committed Then Restore snap
    clsLogger.Fail "svcEndOfDay.CloseDay", "Error " & errNum & ": " & errDesc & " (end of day rolled back)"
    On Error GoTo 0
    Err.Raise errNum, "svcEndOfDay.CloseDay", errDesc
End Function
