Attribute VB_Name = "svcApproval"
Option Explicit

' Maker and checker. A request that needs approval creates ONE Pending row and posts NOTHING.
' On approval the action runs again with its stored arguments (all rules are re-checked); if the re-check fails the approval is Rejected.

' Adds a Pending approval to the unit of work. args are stored as the Note so the approved action runs with the same data.
Public Function Request(ByVal uow As clsUnitOfWork, ByVal requestType As String, ByVal accountNo As String, ByVal counter As String, _
                        ByVal amount As Variant, ByVal branchId As String, ByVal args As Object, ByVal requestedBy As String) As clsResult
    Dim a As clsApproval
    Set a = New clsApproval
    a.ApprovalID = uow.NewId(modConstants.TBL_APPROVALS, "ApprovalID", "AP", 6)
    a.RequestType = requestType
    a.AccountNo = accountNo
    a.CounterAccountNo = counter
    a.Amount = amount
    a.BranchID = branchId
    a.Note = modArgs.Encode(args)
    a.RequestedBy = requestedBy
    a.RequestedAt = modDateUtils.TimeStamp()
    a.Status = modConstants.STATUS_PENDING
    uow.AppendRow modConstants.TBL_APPROVALS, repoApproval.ApprovalToRow(a)
    repoAudit.Add uow, "ApprovalRequest", a.ApprovalID, requestType & " " & modMoney.FormatAmount(amount) & " " & accountNo
    Set Request = modMessages.OkResult("Approval " & a.ApprovalID & " was created and is Pending. Nothing was posted. A different authorised user must approve it (put " & a.ApprovalID & " in Control C12).", 1, a.ApprovalID)
End Function

Private Sub SetDecision(ByVal uow As clsUnitOfWork, ByVal approvalId As String, ByVal status As String, ByVal userId As String, _
                        ByVal note As String, ByVal groupId As String)
    Dim data As Variant, row As Long
    data = uow.WorkingTable(modConstants.TBL_APPROVALS)
    row = repoApproval.FindApproval(data, approvalId)
    data(row, repoApproval.APC_STATUS) = status
    data(row, repoApproval.APC_APPROVER) = userId
    data(row, repoApproval.APC_DECIDEDAT) = modDateUtils.TimeStamp()
    If Len(note) > 0 Then data(row, repoApproval.APC_DECISIONNOTE) = note
    If Len(groupId) > 0 Then data(row, repoApproval.APC_GROUP) = groupId
    uow.ReplaceTable modConstants.TBL_APPROVALS, data
End Sub

' Loads the approval named in Control C12 and checks that it is Pending and that the user may decide it.
Private Function LoadForDecision(ByVal uow As clsUnitOfWork, ByVal user As clsUser, ByVal approvalId As String, _
                                 ByRef req As clsApproval, ByRef msg As String) As Boolean
    Dim data As Variant, row As Long
    data = uow.WorkingTable(modConstants.TBL_APPROVALS)
    row = repoApproval.FindApproval(data, approvalId)
    If row = 0 Then msg = "Approval " & approvalId & " does not exist. Put the approval ID in Control C12.": Exit Function
    Set req = repoApproval.FromRow(data, row)
    If req.Status <> modConstants.STATUS_PENDING Then msg = "Approval " & approvalId & " is already " & req.Status & ".": Exit Function
    msg = svcSecurity.ApproverProblem(user, req)
    LoadForDecision = (Len(msg) = 0)
End Function

' Runs the requested action. Arguments come from the stored Note; the approval columns win for account, counter, amount and branch.
Private Function Dispatch(ByVal uow As clsUnitOfWork, ByVal req As clsApproval, ByVal approverId As String) As clsResult
    Dim args As Object, by As String
    Set args = modArgs.Decode(req.Note)
    If Len(req.AccountNo) > 0 Then args("account") = req.AccountNo
    If Len(req.CounterAccountNo) > 0 Then args("counter") = req.CounterAccountNo
    ' A foreign request stores its amount in its own currency in the note; the approval column holds the MMK value.
    If Not (args.Exists("ccy") And modArgs.Text(args, "ccy", modConstants.CCY_MMK) <> modConstants.CCY_MMK And args.Exists("amount")) Then args("amount") = req.Amount
    args("branch") = req.BranchID
    by = req.RequestedBy
    Select Case req.RequestType
        Case modConstants.REQ_LARGE_DEPOSIT
            If modArgs.Text(args, "open") = "FD" Then
                Set Dispatch = svcFixedDeposit.ExecOpenFD(uow, args, by, approverId)
            ElseIf modArgs.Text(args, "open") = "ACC" Then
                Set Dispatch = svcAccounts.ExecOpen(uow, args, by, approverId)
            Else
                Set Dispatch = svcCashOps.ExecDeposit(uow, args, by, approverId)
            End If
        Case modConstants.REQ_LARGE_WITHDRAWAL
            Set Dispatch = svcCashOps.ExecWithdraw(uow, args, by, approverId)
        Case modConstants.REQ_LARGE_TRANSFER
            Set Dispatch = svcTransfers.ExecTransfer(uow, args, by, approverId)
        Case modConstants.REQ_LARGE_INTERBANK
            Set Dispatch = svcInterbank.ExecOut(uow, args, by, approverId)
        Case modConstants.REQ_CASH_SHIPMENT
            Set Dispatch = svcCashOps.ExecShipment(uow, args, by, approverId)
        Case modConstants.REQ_REVERSAL
            Set Dispatch = svcLedger.ReverseGroup(uow, modArgs.Text(args, "group"), by, approverId)
        Case modConstants.REQ_BREAK_FD
            Set Dispatch = svcFixedDeposit.ExecBreak(uow, args, by, approverId)
        Case modConstants.REQ_CLOSE_ACCOUNT
            Set Dispatch = svcAccounts.ExecClose(uow, args, by, approverId)
        Case modConstants.REQ_UNFREEZE
            Set Dispatch = svcAccounts.ExecUnfreeze(uow, args, by, approverId)
        Case modConstants.REQ_REACTIVATE
            Set Dispatch = svcAccounts.ExecReactivate(uow, args, by, approverId)
        Case Else
            Set Dispatch = svcApprovalExtra.DispatchExtra(uow, req, args, by, approverId)
    End Select
End Function

' Control: C12 = the GroupID to reverse. The reversal is tried on a throw-away unit of work first, so an impossible one is refused
' now; the real reversal runs only after a different HeadOffice user approves.
Public Function RequestReversal(ByVal user As clsUser, ByVal inp As Object) As clsResult
    Dim uow As clsUnitOfWork, msg As String, grp As String, args As Object, trial As clsUnitOfWork, r As clsResult
    msg = svcSecurity.Denied(user, svcSecurity.ACT_REVERSE)
    If Len(msg) > 0 Then Set RequestReversal = modMessages.FailResult(msg): Exit Function
    grp = CStr(inp(modConstants.INP_REFID))
    If Len(grp) = 0 Then Set RequestReversal = modMessages.FailResult("Put the GroupID to reverse in Control C12."): Exit Function
    Set trial = New clsUnitOfWork
    trial.Init repoConfig.BusinessDate(), user.UserID
    Set r = svcLedger.ReverseGroup(trial, grp, user.UserID, "")
    If Not r.Success Then Set RequestReversal = r: Exit Function
    Set uow = New clsUnitOfWork
    uow.Init repoConfig.BusinessDate(), user.UserID
    Set args = modArgs.NewArgs()
    args("group") = grp
    args("note") = CStr(inp(modConstants.INP_NOTE))
    Set r = Request(uow, modConstants.REQ_REVERSAL, "", "", CDec(0), modConstants.HEAD_OFFICE_ID, args, user.UserID)
    uow.Commit
    Set RequestReversal = r
End Function

Public Function Approve(ByVal user As clsUser, ByVal inp As Object) As clsResult
    Dim uow As clsUnitOfWork, req As clsApproval, msg As String, r As clsResult, uow2 As clsUnitOfWork, id As String
    msg = svcSecurity.Denied(user, svcSecurity.ACT_APPROVE)
    If Len(msg) > 0 Then Set Approve = modMessages.FailResult(msg): Exit Function
    id = CStr(inp(modConstants.INP_REFID))
    Set uow = New clsUnitOfWork
    uow.Init repoConfig.BusinessDate(), user.UserID
    If Not LoadForDecision(uow, user, id, req, msg) Then Set Approve = modMessages.FailResult(msg): Exit Function
    Set r = Dispatch(uow, req, user.UserID)
    If r.Success Then
        SetDecision uow, id, modConstants.STATUS_APPROVED, user.UserID, "", r.RefID
        repoAudit.Add uow, "ApprovalApproved", id, req.RequestType & " approved by " & user.UserID
        uow.Commit
        Set Approve = modMessages.OkResult("Approval " & id & " approved and posted. " & r.Message, r.Count, r.RefID)
    Else
        ' The re-check failed (for example the balance changed): reject with the reason. The failed run is thrown away.
        Set uow2 = New clsUnitOfWork
        uow2.Init repoConfig.BusinessDate(), user.UserID
        SetDecision uow2, id, modConstants.STATUS_REJECTED, user.UserID, r.Message, ""
        svcApprovalExtra.OnReject uow2, req, r.Message
        repoAudit.Add uow2, "ApprovalRejected", id, "Re-check failed: " & r.Message
        uow2.Commit
        Set Approve = modMessages.FailResult("Approval " & id & " was Rejected because the re-check failed: " & r.Message)
    End If
End Function

Public Function Reject(ByVal user As clsUser, ByVal inp As Object) As clsResult
    Dim uow As clsUnitOfWork, req As clsApproval, msg As String, id As String
    msg = svcSecurity.Denied(user, svcSecurity.ACT_APPROVE)
    If Len(msg) > 0 Then Set Reject = modMessages.FailResult(msg): Exit Function
    id = CStr(inp(modConstants.INP_REFID))
    Set uow = New clsUnitOfWork
    uow.Init repoConfig.BusinessDate(), user.UserID
    If Not LoadForDecision(uow, user, id, req, msg) Then Set Reject = modMessages.FailResult(msg): Exit Function
    SetDecision uow, id, modConstants.STATUS_REJECTED, user.UserID, CStr(inp(modConstants.INP_NOTE)), ""
    svcApprovalExtra.OnReject uow, req, CStr(inp(modConstants.INP_NOTE))
    repoAudit.Add uow, "ApprovalRejected", id, req.RequestType & " rejected by " & user.UserID
    uow.Commit
    Set Reject = modMessages.OkResult("Approval " & id & " was rejected. Nothing was posted.", 1, id)
End Function
