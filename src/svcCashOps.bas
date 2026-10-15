Attribute VB_Name = "svcCashOps"
Option Explicit

' Cash deposit, cash withdrawal and cash shipments between head office and branch vaults.
' Deposit and Withdraw are the Control-sheet entry points; ExecDeposit and ExecWithdraw do the work
' (and run again when an approval is approved).

' ---- Deposit ----

' args: account, amount, branch (processing), note. Optional: txntype, jtype, group (used by account opening).
Public Function ExecDeposit(ByVal uow As clsUnitOfWork, ByVal args As Object, ByVal createdBy As String, ByVal approvedBy As String) As clsResult
    Dim a As clsAccount, amt As Variant, p As String, h As String, msg As String, lines As Collection
    Dim grp As String, jr As clsResult, txnType As String, jType As String, desc As String
    amt = modArgs.Dec(args, "amount")
    p = modArgs.Text(args, "branch")
    Set a = uow.GetAccount(modArgs.Text(args, "account"))
    If a Is Nothing Then Set ExecDeposit = modMessages.FailResult(modMessages.UnknownAccount(modArgs.Text(args, "account"))): Exit Function
    If Not svcAccounts.CanCredit(a, msg) Then Set ExecDeposit = modMessages.FailResult(msg): Exit Function
    If Not uow.Branches().Exists(p) Then Set ExecDeposit = modMessages.FailResult(modMessages.UnknownBranch(p)): Exit Function
    txnType = modArgs.Text(args, "txntype", modConstants.TXN_DEPOSIT)
    jType = modArgs.Text(args, "jtype", modConstants.JT_DEPOSIT)
    grp = modArgs.Text(args, "group")
    If Len(grp) = 0 Then grp = uow.NewGroupId()
    h = a.BranchID
    desc = txnType & " " & modMoney.FormatCcy(amt, a.Ccy) & " to " & a.AccountNo
    If a.Ccy <> modConstants.CCY_MMK Then
        Set jr = svcFx.FxDeposit(uow, a, amt, p, grp, txnType, jType, desc, modArgs.Text(args, "note"), createdBy, approvedBy)
        If Not jr.Success Then Set ExecDeposit = jr: Exit Function
        Set ExecDeposit = modMessages.OkResult(txnType & " of " & modMoney.FormatCcy(amt, a.Ccy) & " posted to " & a.AccountNo & ". New balance " & modMoney.FormatCcy(a.Balance, a.Ccy) & ". Group " & grp & ".", 1, grp)
        Exit Function
    End If
    Set lines = New Collection
    svcLedger.AddDr lines, p, modConstants.GL_VAULT, amt
    If p <> h Then
        svcLedger.AddCr lines, p, modConstants.GL_INTERBRANCH, amt
        svcLedger.AddDr lines, h, modConstants.GL_INTERBRANCH, amt
    End If
    svcLedger.AddCr lines, h, a.DepositGL, amt, a.AccountNo
    Set jr = svcLedger.PostJournal(uow, jType, desc, grp, lines)
    If Not jr.Success Then Set ExecDeposit = jr: Exit Function
    svcAccounts.Effect uow, a, txnType, modConstants.DIR_CREDIT, amt, "", p, desc, modArgs.Text(args, "note"), jr.RefID, grp, createdBy, approvedBy, True
    Set ExecDeposit = modMessages.OkResult(txnType & " of " & modMoney.FormatAmount(amt) & " MMK posted to " & a.AccountNo & ". New balance " & _
        modMoney.FormatAmount(a.Balance) & ". Group " & grp & ".", 1, grp)
End Function

Public Function Deposit(ByVal user As clsUser, ByVal inp As Object) As clsResult
    Dim uow As clsUnitOfWork, msg As String, amt As Variant, args As Object, p As String, r As clsResult
    msg = svcSecurity.Denied(user, svcSecurity.ACT_DEPOSIT)
    If Len(msg) > 0 Then Set Deposit = modMessages.FailResult(msg): Exit Function
    If Not modArgs.ParseAmountAny(inp(modConstants.INP_AMOUNT), amt, msg) Then Set Deposit = modMessages.FailResult(msg): Exit Function
    Set uow = New clsUnitOfWork
    uow.Init repoConfig.BusinessDate(), user.UserID
    If Not svcSecurity.ResolveBranch(uow, user, CStr(inp(modConstants.INP_BRANCH)), p, msg) Then Set Deposit = modMessages.FailResult(msg): Exit Function
    Set args = modArgs.FromInputs(inp)
    args("amount") = amt
    args("branch") = p
    If uow.GetAccount(CStr(args("account"))) Is Nothing Then Set Deposit = modMessages.FailResult(modMessages.UnknownAccount(CStr(args("account")))): Exit Function
    If Not svcFx.AmountOk(uow.GetAccount(CStr(args("account"))), amt, msg) Then Set Deposit = modMessages.FailResult(msg): Exit Function
    If svcSecurity.OverTellerLimit(svcFx.LimitBase(uow, CStr(args("account")), amt)) Then
        Set r = svcApproval.Request(uow, modConstants.REQ_LARGE_DEPOSIT, CStr(args("account")), "", amt, p, args, user.UserID)
    Else
        Set r = ExecDeposit(uow, args, user.UserID, "")
    End If
    If r.Success Then uow.Commit
    Set Deposit = r
End Function

' ---- Withdrawal ----

' args: account, amount, branch (processing), note. Fee CROSS_BRANCH_WD only when the processing branch is not the home branch.
Public Function ExecWithdraw(ByVal uow As clsUnitOfWork, ByVal args As Object, ByVal createdBy As String, ByVal approvedBy As String) As clsResult
    Dim a As clsAccount, amt As Variant, p As String, h As String, msg As String, lines As Collection, total As Variant
    Dim fee As Variant, tax As Variant, grp As String, jr As clsResult, desc As String
    amt = modArgs.Dec(args, "amount")
    p = modArgs.Text(args, "branch")
    Set a = uow.GetAccount(modArgs.Text(args, "account"))
    If a Is Nothing Then Set ExecWithdraw = modMessages.FailResult(modMessages.UnknownAccount(modArgs.Text(args, "account"))): Exit Function
    If Not uow.Branches().Exists(p) Then Set ExecWithdraw = modMessages.FailResult(modMessages.UnknownBranch(p)): Exit Function
    h = a.BranchID
    If a.Ccy <> modConstants.CCY_MMK Then
        grp = uow.NewGroupId()
        Set jr = svcFx.FxWithdraw(uow, a, amt, p, grp, "Withdrawal " & modMoney.FormatCcy(amt, a.Ccy) & " from " & a.AccountNo, modArgs.Text(args, "note"), createdBy, approvedBy, False)
        If Not jr.Success Then Set ExecWithdraw = jr: Exit Function
        Set ExecWithdraw = modMessages.OkResult("Withdrawal of " & modMoney.FormatCcy(amt, a.Ccy) & " from " & a.AccountNo & " posted. New balance " & modMoney.FormatCcy(a.Balance, a.Ccy) & ". Group " & grp & ".", 1, grp)
        Exit Function
    End If
    fee = CDec(0)
    tax = CDec(0)
    If p <> h Then
        If Not svcFees.Compute(modConstants.FEE_CROSS_WD, amt, fee, tax, msg) Then Set ExecWithdraw = modMessages.FailResult(msg): Exit Function
    End If
    total = amt + fee + tax
    If Not svcAccounts.CanDebit(a, total, False, msg) Then Set ExecWithdraw = modMessages.FailResult(msg): Exit Function
    If uow.GLNet(p, modConstants.GL_VAULT) - amt < 0 Then Set ExecWithdraw = modMessages.FailResult(modMessages.MSG_VAULT_SHORT): Exit Function
    grp = uow.NewGroupId()
    desc = "Withdrawal " & modMoney.FormatAmount(amt) & " from " & a.AccountNo
    Set lines = New Collection
    svcLedger.AddDr lines, h, a.DepositGL, total, a.AccountNo
    If p <> h Then
        svcLedger.AddCr lines, h, modConstants.GL_INTERBRANCH, total
        svcLedger.AddDr lines, p, modConstants.GL_INTERBRANCH, total
    End If
    svcLedger.AddCr lines, p, modConstants.GL_FEE_INCOME, fee
    svcLedger.AddCr lines, p, modConstants.GL_TAX_FEES, tax
    svcLedger.AddCr lines, p, modConstants.GL_VAULT, amt
    Set jr = svcLedger.PostJournal(uow, modConstants.JT_WITHDRAWAL, desc, grp, lines)
    If Not jr.Success Then Set ExecWithdraw = jr: Exit Function
    svcAccounts.Effect uow, a, modConstants.TXN_WITHDRAWAL, modConstants.DIR_DEBIT, amt, "", p, desc, modArgs.Text(args, "note"), jr.RefID, grp, createdBy, approvedBy, True
    svcAccounts.FeeRows uow, a, fee, tax, modConstants.FEE_CROSS_WD, "", p, jr.RefID, grp, createdBy, approvedBy
    Set ExecWithdraw = modMessages.OkResult("Withdrawal of " & modMoney.FormatAmount(amt) & " MMK from " & a.AccountNo & " posted. Fee " & _
        modMoney.FormatAmount(fee + tax) & ". New balance " & modMoney.FormatAmount(a.Balance) & ". Group " & grp & ".", 1, grp)
End Function

Public Function Withdraw(ByVal user As clsUser, ByVal inp As Object) As clsResult
    Dim uow As clsUnitOfWork, msg As String, amt As Variant, args As Object, p As String, r As clsResult, needApproval As Boolean
    msg = svcSecurity.Denied(user, svcSecurity.ACT_WITHDRAW)
    If Len(msg) > 0 Then Set Withdraw = modMessages.FailResult(msg): Exit Function
    If Not modArgs.ParseAmountAny(inp(modConstants.INP_AMOUNT), amt, msg) Then Set Withdraw = modMessages.FailResult(msg): Exit Function
    Set uow = New clsUnitOfWork
    uow.Init repoConfig.BusinessDate(), user.UserID
    If Not svcSecurity.ResolveBranch(uow, user, CStr(inp(modConstants.INP_BRANCH)), p, msg) Then Set Withdraw = modMessages.FailResult(msg): Exit Function
    Set args = modArgs.FromInputs(inp)
    args("amount") = amt
    args("branch") = p
    If uow.GetAccount(CStr(args("account"))) Is Nothing Then Set Withdraw = modMessages.FailResult(modMessages.UnknownAccount(CStr(args("account")))): Exit Function
    If Not svcFx.AmountOk(uow.GetAccount(CStr(args("account"))), amt, msg) Then Set Withdraw = modMessages.FailResult(msg): Exit Function
    needApproval = svcSecurity.OverTellerLimit(svcFx.LimitBase(uow, CStr(args("account")), amt)) Or svcSecurity.OverDailyLimit(CStr(args("account")), uow.BusinessDate, amt, uow.GetAccount(CStr(args("account"))).Ccy)
    If needApproval Then
        Set r = svcApproval.Request(uow, modConstants.REQ_LARGE_WITHDRAWAL, CStr(args("account")), "", amt, p, args, user.UserID)
    Else
        Set r = ExecWithdraw(uow, args, user.UserID, "")
    End If
    If r.Success Then uow.Commit
    Set Withdraw = r
End Function

' ---- Cash shipments (always approved by a different HeadOffice user) ----

' args: branch (the branch B), amount, dir ("To" = head office to branch, "Return" = branch back to head office), note.
' When the head office vault is short, cash is first bought from treasury in blocks of CashShipmentRounding.
Public Function ExecShipment(ByVal uow As clsUnitOfWork, ByVal args As Object, ByVal createdBy As String, ByVal approvedBy As String) As clsResult
    Dim b As String, amt As Variant, grp As String, lines As Collection, jr As clsResult, need As Variant
    Dim blockSize As Variant, blocks As Variant, ho As String, toBranch As Boolean
    b = modArgs.Text(args, "branch")
    ho = modConstants.HEAD_OFFICE_ID
    amt = modArgs.Dec(args, "amount")
    toBranch = (modArgs.Text(args, "dir", "To") <> "Return")
    If b = ho Or Not uow.Branches().Exists(b) Then Set ExecShipment = modMessages.FailResult("Choose a branch other than head office in Control C8."): Exit Function
    If modArgs.Text(args, "ccy", modConstants.CCY_MMK) <> modConstants.CCY_MMK Then
        grp = uow.NewGroupId()
        Set jr = svcFx.FxShipment(uow, b, modArgs.Text(args, "ccy"), amt, toBranch, grp)
        If Not jr.Success Then Set ExecShipment = jr: Exit Function
        repoAudit.Add uow, "CashShipment", b, modArgs.Text(args, "ccy") & " " & amt & " " & IIf(toBranch, "to ", "return from ") & b & ", approved by " & approvedBy
        Set ExecShipment = modMessages.OkResult("Foreign cash shipment of " & modMoney.FormatCcy(amt, modArgs.Text(args, "ccy")) & IIf(toBranch, " to ", " from ") & b & " posted. Group " & grp & ".", 1, grp)
        Exit Function
    End If
    grp = uow.NewGroupId()
    If toBranch Then
        need = amt - uow.GLNet(ho, modConstants.GL_VAULT)
        If need > 0 Then
            blockSize = repoConfig.GetDec(modConstants.CFG_SHIPROUND)
            blocks = Int((need + blockSize - CDec(1)) / blockSize) * blockSize
            Set lines = New Collection
            svcLedger.AddDr lines, ho, modConstants.GL_VAULT, blocks
            svcLedger.AddCr lines, ho, modConstants.GL_TREASURY, blocks
            Set jr = svcLedger.PostJournal(uow, modConstants.JT_CASH_PURCHASE, "Cash order from treasury " & modMoney.FormatAmount(blocks), grp, lines)
            If Not jr.Success Then Set ExecShipment = jr: Exit Function
        End If
        Set lines = New Collection
        svcLedger.AddDr lines, b, modConstants.GL_VAULT, amt
        svcLedger.AddCr lines, b, modConstants.GL_INTERBRANCH, amt
        svcLedger.AddDr lines, ho, modConstants.GL_INTERBRANCH, amt
        svcLedger.AddCr lines, ho, modConstants.GL_VAULT, amt
    Else
        If uow.GLNet(b, modConstants.GL_VAULT) - amt < 0 Then Set ExecShipment = modMessages.FailResult("The vault of " & b & " does not hold that much cash."): Exit Function
        Set lines = New Collection
        svcLedger.AddDr lines, b, modConstants.GL_INTERBRANCH, amt
        svcLedger.AddCr lines, b, modConstants.GL_VAULT, amt
        svcLedger.AddDr lines, ho, modConstants.GL_VAULT, amt
        svcLedger.AddCr lines, ho, modConstants.GL_INTERBRANCH, amt
    End If
    Set jr = svcLedger.PostJournal(uow, modConstants.JT_CASH_SHIPMENT, "Cash shipment " & IIf(toBranch, "to ", "from ") & b & " " & modMoney.FormatAmount(amt), grp, lines)
    If Not jr.Success Then Set ExecShipment = jr: Exit Function
    repoAudit.Add uow, "CashShipment", b, IIf(toBranch, "To ", "Return from ") & b & " " & modMoney.FormatAmount(amt) & ", approved by " & approvedBy
    Set ExecShipment = modMessages.OkResult("Cash shipment of " & modMoney.FormatAmount(amt) & " MMK " & IIf(toBranch, "to ", "from ") & b & " posted. Group " & grp & ".", 1, grp)
End Function

' Control: C8 = branch, C7 = amount, C15 = note ("RETURN" sends the cash back to head office). Always creates a Pending approval.
Public Function Shipment(ByVal user As clsUser, ByVal inp As Object) As clsResult
    Dim ccy As String, uow As clsUnitOfWork, msg As String, amt As Variant, args As Object, b As String, r As clsResult
    msg = svcSecurity.Denied(user, svcSecurity.ACT_SHIPMENT)
    If Len(msg) > 0 Then Set Shipment = modMessages.FailResult(msg): Exit Function
    ccy = CStr(inp(modConstants.INP_CURRENCY))
    If Len(ccy) = 0 Then ccy = modConstants.CCY_MMK
    If Not modArgs.ParseAmountCcy(inp(modConstants.INP_AMOUNT), ccy, amt, msg) Then Set Shipment = modMessages.FailResult(msg): Exit Function
    b = CStr(inp(modConstants.INP_BRANCH))
    Set uow = New clsUnitOfWork
    uow.Init repoConfig.BusinessDate(), user.UserID
    If Not uow.Branches().Exists(b) Or b = modConstants.HEAD_OFFICE_ID Then Set Shipment = modMessages.FailResult("Choose the receiving branch (not head office) in Control C8."): Exit Function
    Set args = modArgs.NewArgs()
    args("branch") = b
    args("amount") = amt
    args("ccy") = ccy
    If UCase$(Trim$(CStr(inp(modConstants.INP_NOTE)))) = "RETURN" Then args("dir") = "Return" Else args("dir") = "To"
    Set r = svcApproval.Request(uow, modConstants.REQ_CASH_SHIPMENT, "", "", amt, b, args, user.UserID)
    uow.Commit
    Set Shipment = r
End Function
