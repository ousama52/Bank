Attribute VB_Name = "svcExpenses"
Option Explicit

' Stage 2: one-off expenses (repairs, purchases). Up to PettyCashLimit they are paid from the branch vault, above that by head office.
' Above ExpenseApprovalLimit a different authorised user must approve first.

' args: branch, gl, amount, vendor, desc, note.
Public Function ExecExpense(ByVal uow As clsUnitOfWork, ByVal args As Object, ByVal createdBy As String, ByVal approvedBy As String) As clsResult
    Dim b As String, gl As String, amt As Variant, lines As Collection, jr As clsResult, grp As String, id As String, paidFrom As String
    Dim ho As String, chartRow As Variant
    ho = modConstants.HEAD_OFFICE_ID
    b = modArgs.Text(args, "branch")
    gl = modArgs.Text(args, "gl")
    amt = modArgs.Dec(args, "amount")
    If Not uow.Branches().Exists(b) Then Set ExecExpense = modMessages.FailResult(modMessages.UnknownBranch(b)): Exit Function
    If Not uow.Chart().Exists(gl) Then Set ExecExpense = modMessages.FailResult("GL " & gl & " does not exist."): Exit Function
    chartRow = uow.Chart()(gl)
    If chartRow(1) <> "Expense" Or Not chartRow(3) Then Set ExecExpense = modMessages.FailResult("GL " & gl & " is not an active Expense account."): Exit Function
    Set lines = New Collection
    svcLedger.AddDr lines, b, gl, amt
    If amt <= repoConfig.GetDec(modConstants.CFG_PETTY_LIMIT) Then
        paidFrom = "Vault"
        If uow.GLNet(b, modConstants.GL_VAULT) - amt < 0 Then Set ExecExpense = modMessages.FailResult(modMessages.MSG_VAULT_SHORT): Exit Function
        svcLedger.AddCr lines, b, modConstants.GL_VAULT, amt
    Else
        paidFrom = "HeadOffice"
        If b = ho Then
            svcLedger.AddCr lines, ho, modConstants.GL_TREASURY, amt
        Else
            svcLedger.AddCr lines, b, modConstants.GL_INTERBRANCH, amt
            svcLedger.AddDr lines, ho, modConstants.GL_INTERBRANCH, amt
            svcLedger.AddCr lines, ho, modConstants.GL_TREASURY, amt
        End If
    End If
    grp = uow.NewGroupId()
    Set jr = svcLedger.PostJournal(uow, modConstants.JT_EXPENSE, "Expense " & modArgs.Text(args, "vendor") & " " & modArgs.Text(args, "desc"), grp, lines)
    If Not jr.Success Then Set ExecExpense = jr: Exit Function
    id = uow.NewId(modConstants.TBL_EXPENSES, "ExpenseID", "EX", 6)
    uow.AppendRow modConstants.TBL_EXPENSES, repoCost.ExpenseRow(id, b, uow.BusinessDate, gl, amt, modArgs.Text(args, "vendor"), modArgs.Text(args, "desc"), _
        paidFrom, "Paid", createdBy, approvedBy, jr.RefID, jr.RefID)
    repoAudit.Add uow, "Expense", id, modMoney.FormatAmount(amt) & " on " & gl & " at " & b & ", paid from " & paidFrom
    Set ExecExpense = modMessages.OkResult("Expense " & id & " of " & modMoney.FormatAmount(amt) & " MMK posted to GL " & gl & " at " & b & " (paid from " & paidFrom & "). Group " & grp & ".", 1, grp)
End Function

' Control: C8 branch, C19 GL, C7 amount, C20 vendor, C21 description.
Public Function PostExpense(ByVal user As clsUser, ByVal inp As Object) As clsResult
    Dim uow As clsUnitOfWork, msg As String, amt As Variant, args As Object, b As String, r As clsResult
    msg = svcSecurity.Denied(user, svcSecurity.ACT_EXPENSE)
    If Len(msg) > 0 Then Set PostExpense = modMessages.FailResult(msg): Exit Function
    If Not modArgs.ParseAmount(inp(modConstants.INP_AMOUNT), amt, msg) Then Set PostExpense = modMessages.FailResult(msg): Exit Function
    Set uow = New clsUnitOfWork
    uow.Init repoConfig.BusinessDate(), user.UserID
    If Not svcSecurity.ResolveBranch(uow, user, CStr(inp(modConstants.INP_BRANCH)), b, msg) Then Set PostExpense = modMessages.FailResult(msg): Exit Function
    Set args = modArgs.NewArgs()
    args("branch") = b
    args("gl") = CStr(inp(modConstants.INP_GL))
    args("amount") = amt
    args("vendor") = CStr(inp(modConstants.INP_VENDOR))
    args("desc") = CStr(inp(modConstants.INP_DESC))
    If amt > repoConfig.GetDec(modConstants.CFG_EXPENSE_LIMIT) Then
        Set r = svcApproval.Request(uow, modConstants.REQ_EXPENSE, "", "", amt, b, args, user.UserID)
    Else
        Set r = ExecExpense(uow, args, user.UserID, "")
    End If
    If r.Success Then uow.Commit
    Set PostExpense = r
End Function
