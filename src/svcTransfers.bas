Attribute VB_Name = "svcTransfers"
Option Explicit

' Transfers between accounts, inside a branch and between branches. All ledger lines stay in each account's home branch;
' a cross-branch movement runs through GL 1900 so every branch stays balanced.

' The core posting, also used by account opening (funded by transfer) and by fixed deposit maturity and break.
' isSystem = no cross-branch fee, fixed deposits may be debited, no minimum balance check.
Public Function PostTransfer(ByVal uow As clsUnitOfWork, ByVal src As clsAccount, ByVal dst As clsAccount, ByVal amt As Variant, _
                             ByVal p As String, ByVal grp As String, ByVal outType As String, ByVal inType As String, _
                             ByVal jType As String, ByVal desc As String, ByVal note As String, ByVal createdBy As String, _
                             ByVal approvedBy As String, ByVal isSystem As Boolean) As clsResult
    Dim fee As Variant, tax As Variant, total As Variant, msg As String, h1 As String, h2 As String
    Dim lines As Collection, jr As clsResult
    fee = CDec(0)
    tax = CDec(0)
    If src.AccountNo = dst.AccountNo Then Set PostTransfer = modMessages.FailResult(modMessages.MSG_SAME_ACCOUNT): Exit Function
    If src.Ccy <> dst.Ccy Then Set PostTransfer = modMessages.FailResult("A transfer between different currencies is not allowed. Use RunConvertCurrency."): Exit Function
    If src.Ccy <> modConstants.CCY_MMK Then
        Set PostTransfer = svcFx.FxTransfer(uow, src, dst, amt, p, grp, outType, inType, jType, desc, note, createdBy, approvedBy, isSystem)
        Exit Function
    End If
    h1 = src.BranchID
    h2 = dst.BranchID
    If h1 <> h2 And Not isSystem Then
        If Not svcFees.Compute(modConstants.FEE_TRANSFER_BRANCH, amt, fee, tax, msg) Then Set PostTransfer = modMessages.FailResult(msg): Exit Function
    End If
    total = amt + fee + tax
    If Not svcAccounts.CanDebit(src, total, isSystem, msg) Then Set PostTransfer = modMessages.FailResult(msg): Exit Function
    If Not svcAccounts.CanCredit(dst, msg) Then Set PostTransfer = modMessages.FailResult(msg): Exit Function
    Set lines = New Collection
    svcLedger.AddDr lines, h1, src.DepositGL, total, src.AccountNo
    svcLedger.AddCr lines, h1, modConstants.GL_FEE_INCOME, fee
    svcLedger.AddCr lines, h1, modConstants.GL_TAX_FEES, tax
    If h1 = h2 Then
        svcLedger.AddCr lines, h1, dst.DepositGL, amt, dst.AccountNo
    Else
        svcLedger.AddCr lines, h1, modConstants.GL_INTERBRANCH, amt
        svcLedger.AddDr lines, h2, modConstants.GL_INTERBRANCH, amt
        svcLedger.AddCr lines, h2, dst.DepositGL, amt, dst.AccountNo
    End If
    Set jr = svcLedger.PostJournal(uow, jType, desc, grp, lines)
    If Not jr.Success Then Set PostTransfer = jr: Exit Function
    svcAccounts.Effect uow, src, outType, modConstants.DIR_DEBIT, amt, dst.AccountNo, p, desc, note, jr.RefID, grp, createdBy, approvedBy, Not isSystem
    svcAccounts.FeeRows uow, src, fee, tax, modConstants.FEE_TRANSFER_BRANCH, dst.AccountNo, p, jr.RefID, grp, createdBy, approvedBy
    svcAccounts.Effect uow, dst, inType, modConstants.DIR_CREDIT, amt, src.AccountNo, p, desc, note, jr.RefID, grp, createdBy, approvedBy, Not isSystem
    ' Count carries the fee (with tax) back to the caller for its message.
    Set PostTransfer = modMessages.OkResult("", CLng(fee + tax), grp)
End Function

' args: account (from), counter (to), amount, branch (processing), note.
Public Function ExecTransfer(ByVal uow As clsUnitOfWork, ByVal args As Object, ByVal createdBy As String, ByVal approvedBy As String) As clsResult
    Dim src As clsAccount, dst As clsAccount, amt As Variant, grp As String, r As clsResult
    amt = modArgs.Dec(args, "amount")
    Set src = uow.GetAccount(modArgs.Text(args, "account"))
    Set dst = uow.GetAccount(modArgs.Text(args, "counter"))
    If src Is Nothing Then Set ExecTransfer = modMessages.FailResult(modMessages.UnknownAccount(modArgs.Text(args, "account"))): Exit Function
    If dst Is Nothing Then Set ExecTransfer = modMessages.FailResult(modMessages.UnknownAccount(modArgs.Text(args, "counter"))): Exit Function
    If Not uow.Branches().Exists(modArgs.Text(args, "branch")) Then Set ExecTransfer = modMessages.FailResult(modMessages.UnknownBranch(modArgs.Text(args, "branch"))): Exit Function
    grp = uow.NewGroupId()
    Set r = PostTransfer(uow, src, dst, amt, modArgs.Text(args, "branch"), grp, modConstants.TXN_TRANSFER_OUT, modConstants.TXN_TRANSFER_IN, _
        modConstants.JT_TRANSFER, "Transfer " & modMoney.FormatAmount(amt) & " from " & src.AccountNo & " to " & dst.AccountNo, modArgs.Text(args, "note"), createdBy, approvedBy, False)
    If Not r.Success Then Set ExecTransfer = r: Exit Function
    Set ExecTransfer = modMessages.OkResult("Transfer of " & modMoney.FormatCcy(amt, src.Ccy) & " from " & src.AccountNo & " to " & dst.AccountNo & _
        " posted. Fee " & modMoney.FormatAmount(r.Count) & ". Balance of " & src.AccountNo & ": " & modMoney.FormatAmount(src.Balance) & ". Group " & grp & ".", 1, grp)
End Function

' Control: C5 = from account, C6 = to account, C7 = amount.
Public Function Transfer(ByVal user As clsUser, ByVal inp As Object) As clsResult
    Dim uow As clsUnitOfWork, msg As String, amt As Variant, args As Object, p As String, r As clsResult, needApproval As Boolean
    msg = svcSecurity.Denied(user, svcSecurity.ACT_TRANSFER)
    If Len(msg) > 0 Then Set Transfer = modMessages.FailResult(msg): Exit Function
    If Not modArgs.ParseAmountAny(inp(modConstants.INP_AMOUNT), amt, msg) Then Set Transfer = modMessages.FailResult(msg): Exit Function
    Set uow = New clsUnitOfWork
    uow.Init repoConfig.BusinessDate(), user.UserID
    If Not svcSecurity.ResolveBranch(uow, user, CStr(inp(modConstants.INP_BRANCH)), p, msg) Then Set Transfer = modMessages.FailResult(msg): Exit Function
    Set args = modArgs.FromInputs(inp)
    args("amount") = amt
    args("branch") = p
    If uow.GetAccount(CStr(args("account"))) Is Nothing Then Set Transfer = modMessages.FailResult(modMessages.UnknownAccount(CStr(args("account")))): Exit Function
    If uow.GetAccount(CStr(args("counter"))) Is Nothing Then Set Transfer = modMessages.FailResult(modMessages.UnknownAccount(CStr(args("counter")))): Exit Function
    If CStr(args("account")) = CStr(args("counter")) Then Set Transfer = modMessages.FailResult(modMessages.MSG_SAME_ACCOUNT): Exit Function
    If Not svcFx.AmountOk(uow.GetAccount(CStr(args("account"))), amt, msg) Then Set Transfer = modMessages.FailResult(msg): Exit Function
    needApproval = svcSecurity.OverTellerLimit(svcFx.LimitBase(uow, CStr(args("account")), amt)) Or svcSecurity.OverDailyLimit(CStr(args("account")), uow.BusinessDate, amt, uow.GetAccount(CStr(args("account"))).Ccy)
    If needApproval Then
        Set r = svcApproval.Request(uow, modConstants.REQ_LARGE_TRANSFER, CStr(args("account")), CStr(args("counter")), amt, p, args, user.UserID)
    Else
        Set r = ExecTransfer(uow, args, user.UserID, "")
    End If
    If r.Success Then uow.Commit
    Set Transfer = r
End Function
