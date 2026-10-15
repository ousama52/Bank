Attribute VB_Name = "svcInspection"
Option Explicit

' Stage 8: regulator inspection and fines. The inspection runs every control check and every ratio and lists what it finds.
' A fine is proposed by the Regulator and posted (HO: Dr 5350, Cr 1100) after a different HeadOffice user approves.

Private Const SH_INSP As String = "Inspection"

Private Function Ready() As Boolean
    Ready = modSheetUtils.TableExists(modConstants.TBL_INSPECTIONS) And modSheetUtils.TableExists(modConstants.TBL_FINES)
End Function

Private Sub Find(ByVal findings As Collection, ByVal area As String, ByVal text As String)
    findings.Add Array(area, text)
End Sub

Public Function Inspect(ByVal user As clsUser, ByVal inp As Object) As clsResult
    Dim uow As clsUnitOfWork, msg As String, d As Date, res As Variant, i As Long, findings As Collection, ratios As Collection, r As Variant
    Dim apps As Variant, a As clsApproval, alerts As Variant, dt As Date, b As Object, k As Variant, bt As clsBranch, vault As Variant, id As String, result As String, rowsA As Collection, f As Variant
    msg = svcSecurity.Denied(user, svcSecurity.ACT_INSPECT)
    If Len(msg) = 0 And Not Ready() Then msg = modMessages.MSG_SETUP_FIRST
    If Len(msg) > 0 Then Set Inspect = modMessages.FailResult(msg): Exit Function
    d = repoConfig.BusinessDate()
    Set uow = New clsUnitOfWork
    uow.Init d, user.UserID
    Set findings = New Collection
    res = svcChecks.RunChecks()
    For i = 1 To UBound(res, 1)
        If res(i, 3) = "FAIL" Then Find findings, "Check " & res(i, 1), CStr(res(i, 2)) & ": " & CStr(res(i, 4))
        If res(i, 3) = "WARN" Then Find findings, "Check " & res(i, 1) & " (warning)", CStr(res(i, 2)) & ": " & CStr(res(i, 4))
    Next i
    Set ratios = svcRatios.ComputeRatios(uow, d)
    For Each r In ratios
        If CStr(r(4)) = "Breach" Then Find findings, "Ratio " & CStr(r(0)), "Actual " & Format$(CDbl(r(1)), "0.0000") & " against limit " & Format$(CDbl(r(2)), "0.0000")
    Next r
    apps = modSheetUtils.ReadTable(modConstants.TBL_APPROVALS)
    If Not IsEmpty(apps) Then
        For i = 1 To UBound(apps, 1)
            If CStr(apps(i, repoApproval.APC_STATUS)) = modConstants.STATUS_PENDING Then
                Set a = repoApproval.FromRow(apps, i)
                If d - Int(CDbl(a.RequestedAt)) > 3 Then Find findings, "Approval " & a.ApprovalID, a.RequestType & " pending for " & CLng(d - Int(CDbl(a.RequestedAt))) & " days"
            End If
        Next i
    End If
    If modSheetUtils.TableExists(modConstants.TBL_AMLALERTS) Then
        alerts = modSheetUtils.ReadTable(modConstants.TBL_AMLALERTS)
        If Not IsEmpty(alerts) Then
            For i = 1 To UBound(alerts, 1)
                If CStr(alerts(i, repoGov.ALC_STATUS)) = "New" And modDateUtils.TryToDate(alerts(i, repoGov.ALC_DATE), dt) Then
                    If d - dt > repoConfig.GetLong(modConstants.CFG_AML_DAYS) Then Find findings, "AML alert " & CStr(alerts(i, 1)), "Not reviewed for " & CLng(d - dt) & " days"
                End If
            Next i
        End If
    End If
    Set b = repoBranch.LoadBranches()
    For Each k In b.Keys
        Set bt = b(k)
        vault = uow.GLNet(CStr(k), modConstants.GL_VAULT)
        If vault < bt.VaultMin Then Find findings, "Vault " & CStr(k), "Cash " & modMoney.FormatAmount(vault) & " is below the minimum " & modMoney.FormatAmount(bt.VaultMin)
    Next k
    If findings.Count = 0 Then result = "Pass" Else result = "Findings"
    id = uow.NewId(modConstants.TBL_INSPECTIONS, "InspectionID", "IS", 6)
    uow.AppendRow modConstants.TBL_INSPECTIONS, Array(id, d, user.UserID, result, findings.Count, IIf(findings.Count = 0, "No findings", findings.Count & " finding(s)"))
    repoAudit.Add uow, "Inspection", id, result & ", " & findings.Count & " finding(s)"
    uow.Commit
    modSheetUtils.EnsureSheet SH_INSP
    modSheetUtils.PrepareReport SH_INSP, "1,2"
    modSheetUtils.SetCell SH_INSP, 1, 1, "Inspection " & id & " on " & modDateUtils.ToIso(d) & ": " & UCase$(result) & " (SIMULATION)", True
    Set rowsA = New Collection
    For Each f In findings
        rowsA.Add Array(f(0), f(1))
    Next f
    modSheetUtils.WriteGrid SH_INSP, 3, 1, svcReports.ToGrid(Array("Area", "Finding"), rowsA)
    modSheetUtils.StyleReport SH_INSP, 3, 2, ""
    Set Inspect = modMessages.OkResult("Inspection " & id & ": " & result & " (" & findings.Count & " finding(s)). See sheet " & SH_INSP & ".", findings.Count, id)
End Function

' ---- Fines ----

' Control: C7 amount, C54 reason. The Regulator proposes; a different HeadOffice user approves.
Public Function IssueFine(ByVal user As clsUser, ByVal inp As Object) As clsResult
    Dim uow As clsUnitOfWork, msg As String, amt As Variant, reason As String, id As String, args As Object, r As clsResult
    msg = svcSecurity.Denied(user, svcSecurity.ACT_FINE)
    If Len(msg) = 0 And Not Ready() Then msg = modMessages.MSG_SETUP_FIRST
    If Len(msg) > 0 Then Set IssueFine = modMessages.FailResult(msg): Exit Function
    If Not modArgs.ParseAmount(inp(modConstants.INP_AMOUNT), amt, msg) Then Set IssueFine = modMessages.FailResult(msg): Exit Function
    reason = Trim$(CStr(inp(modConstants.INP_REASON)))
    If Len(reason) = 0 Then Set IssueFine = modMessages.FailResult("Write the reason in Control C54."): Exit Function
    Set uow = New clsUnitOfWork
    uow.Init repoConfig.BusinessDate(), user.UserID
    id = uow.NewId(modConstants.TBL_FINES, "FineID", "FN", 6)
    uow.AppendRow modConstants.TBL_FINES, Array(id, uow.BusinessDate, reason, amt, user.UserID, Empty, "Proposed", Empty)
    Set args = modArgs.NewArgs()
    args("fine") = id
    Set r = svcApproval.Request(uow, modConstants.REQ_FINE, "", "", amt, modConstants.HEAD_OFFICE_ID, args, user.UserID)
    uow.Commit
    Set IssueFine = modMessages.OkResult("Fine " & id & " of " & modMoney.FormatAmount(amt) & " MMK proposed. " & r.Message, 1, id)
End Function

' Run on approval: HO: Dr 5350, Cr 1100.
Public Function ExecFine(ByVal uow As clsUnitOfWork, ByVal args As Object, ByVal createdBy As String, ByVal approvedBy As String) As clsResult
    Dim fines As Variant, fr As Long, amt As Variant, lines As Collection, jr As clsResult, grp As String, ho As String
    ho = modConstants.HEAD_OFFICE_ID
    fines = uow.WorkingTable(modConstants.TBL_FINES)
    fr = repoGov.FindGovById(fines, modArgs.Text(args, "fine"))
    If fr = 0 Then Set ExecFine = modMessages.FailResult("Unknown fine."): Exit Function
    If CStr(fines(fr, repoGov.FNC_STATUS)) <> "Proposed" Then Set ExecFine = modMessages.FailResult("Fine " & CStr(fines(fr, 1)) & " is " & CStr(fines(fr, repoGov.FNC_STATUS)) & "."): Exit Function
    amt = modMoney.ToDec(fines(fr, repoGov.FNC_AMOUNT))
    If uow.GLNet(ho, modConstants.GL_TREASURY) < amt Then Set ExecFine = modMessages.FailResult("The treasury balance (GL 1100) is too low to pay the fine."): Exit Function
    Set lines = New Collection
    svcLedger.AddDr lines, ho, modConstants.GL_FINES, amt
    svcLedger.AddCr lines, ho, modConstants.GL_TREASURY, amt
    grp = uow.NewGroupId()
    Set jr = svcLedger.PostJournal(uow, modConstants.JT_FINE, "Regulatory fine " & CStr(fines(fr, 1)), grp, lines)
    If Not jr.Success Then Set ExecFine = jr: Exit Function
    fines(fr, 6) = approvedBy
    fines(fr, repoGov.FNC_STATUS) = "Posted"
    fines(fr, repoGov.FNC_JOURNAL) = jr.RefID
    uow.ReplaceTable modConstants.TBL_FINES, fines
    repoAudit.Add uow, "FinePosted", CStr(fines(fr, 1)), modMoney.FormatAmount(amt) & ", approved by " & approvedBy
    Set ExecFine = modMessages.OkResult("Fine " & CStr(fines(fr, 1)) & " posted: " & modMoney.FormatAmount(amt) & " MMK. Group " & grp & ".", 1, grp)
End Function

' A rejected fine proposal is Cancelled.
Public Sub CancelFine(ByVal uow As clsUnitOfWork, ByVal fineId As String)
    Dim fines As Variant, fr As Long
    If Not modSheetUtils.TableExists(modConstants.TBL_FINES) Then Exit Sub
    fines = uow.WorkingTable(modConstants.TBL_FINES)
    fr = repoGov.FindGovById(fines, fineId)
    If fr = 0 Then Exit Sub
    If CStr(fines(fr, repoGov.FNC_STATUS)) <> "Proposed" Then Exit Sub
    fines(fr, repoGov.FNC_STATUS) = "Cancelled"
    uow.ReplaceTable modConstants.TBL_FINES, fines
End Sub
