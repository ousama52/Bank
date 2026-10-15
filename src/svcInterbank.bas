Attribute VB_Name = "svcInterbank"
Option Explicit

' FAKE other-bank transfers. Only our own books move: GL 2900 (simulated payable) and GL 1100 (simulated settlement balance).
' Nothing about the other bank is stored except its code and the 10 digit account text.

' ---- Outgoing ----

' args: account, bank (code), counter (other bank account, 10 digits), amount, branch, note.
Public Function ExecOut(ByVal uow As clsUnitOfWork, ByVal args As Object, ByVal createdBy As String, ByVal approvedBy As String) As clsResult
    Dim a As clsAccount, bank As clsOtherBank, amt As Variant, fee As Variant, tax As Variant, total As Variant, msg As String
    Dim lines As Collection, jr As clsResult, grp As String, id As String, other As String, p As String, expected As Date
    amt = modArgs.Dec(args, "amount")
    other = modArgs.Text(args, "counter")
    p = modArgs.Text(args, "branch")
    Set a = uow.GetAccount(modArgs.Text(args, "account"))
    If a Is Nothing Then Set ExecOut = modMessages.FailResult(modMessages.UnknownAccount(modArgs.Text(args, "account"))): Exit Function
    If a.Ccy <> modConstants.CCY_MMK Then Set ExecOut = modMessages.FailResult(modMessages.MSG_FOREIGN_CCY): Exit Function
    Set bank = repoOtherBank.FindBank(modArgs.Text(args, "bank"))
    If bank Is Nothing Then Set ExecOut = modMessages.FailResult("Unknown other bank code " & modArgs.Text(args, "bank") & "."): Exit Function
    If Not (other Like "##########") Then Set ExecOut = modMessages.FailResult("The other bank account must be 10 digits (Control C6)."): Exit Function
    If Not svcFees.Compute(modConstants.FEE_IB_OUT, amt, fee, tax, msg) Then Set ExecOut = modMessages.FailResult(msg): Exit Function
    total = amt + fee + tax
    If Not svcAccounts.CanDebit(a, total, False, msg) Then Set ExecOut = modMessages.FailResult(msg): Exit Function
    grp = uow.NewGroupId()
    Set lines = New Collection
    svcLedger.AddDr lines, a.BranchID, a.DepositGL, total, a.AccountNo
    svcLedger.AddCr lines, a.BranchID, modConstants.GL_IB_PAYABLE, amt
    svcLedger.AddCr lines, a.BranchID, modConstants.GL_FEE_INCOME, fee
    svcLedger.AddCr lines, a.BranchID, modConstants.GL_TAX_FEES, tax
    Set jr = svcLedger.PostJournal(uow, modConstants.JT_IB_OUT, "Interbank out " & modMoney.FormatAmount(amt) & " to " & bank.BankCode, grp, lines)
    If Not jr.Success Then Set ExecOut = jr: Exit Function
    svcAccounts.Effect uow, a, modConstants.TXN_IB_OUT, modConstants.DIR_DEBIT, amt, bank.BankCode, p, "Interbank transfer to " & bank.Name & " " & other, _
        modArgs.Text(args, "note"), jr.RefID, grp, createdBy, approvedBy, True
    svcAccounts.FeeRows uow, a, fee, tax, modConstants.FEE_IB_OUT, bank.BankCode, p, jr.RefID, grp, createdBy, approvedBy
    id = uow.NewId(modConstants.TBL_INTERBANK, "TransferID", "IB", 6)
    expected = uow.BusinessDate + bank.SettleDays
    uow.AppendRow modConstants.TBL_INTERBANK, repoInterbank.NewTransferRow(id, modConstants.DIR_OUT, a.AccountNo, bank.BankCode, other, amt, fee, _
        uow.BusinessDate, expected, modConstants.STATUS_SUBMITTED, jr.RefID, "")
    Set ExecOut = modMessages.OkResult("Interbank transfer " & id & " of " & modMoney.FormatAmount(amt) & " MMK submitted. Expected settlement " & _
        modDateUtils.ToIso(expected) & ". Balance " & modMoney.FormatAmount(a.Balance) & ". Group " & grp & ".", 1, grp)
End Function

' Control: C5 = our account, C6 = other bank account, C9 = bank code, C7 = amount.
Public Function SendOut(ByVal user As clsUser, ByVal inp As Object) As clsResult
    Dim uow As clsUnitOfWork, msg As String, amt As Variant, args As Object, p As String, r As clsResult, needApproval As Boolean
    msg = svcSecurity.Denied(user, svcSecurity.ACT_INTERBANK_OUT)
    If Len(msg) > 0 Then Set SendOut = modMessages.FailResult(msg): Exit Function
    If Not modArgs.ParseAmount(inp(modConstants.INP_AMOUNT), amt, msg) Then Set SendOut = modMessages.FailResult(msg): Exit Function
    Set uow = New clsUnitOfWork
    uow.Init repoConfig.BusinessDate(), user.UserID
    If Not svcSecurity.ResolveBranch(uow, user, CStr(inp(modConstants.INP_BRANCH)), p, msg) Then Set SendOut = modMessages.FailResult(msg): Exit Function
    Set args = modArgs.FromInputs(inp)
    args("amount") = amt
    args("branch") = p
    If uow.GetAccount(CStr(args("account"))) Is Nothing Then Set SendOut = modMessages.FailResult(modMessages.UnknownAccount(CStr(args("account")))): Exit Function
    needApproval = svcSecurity.OverTellerLimit(amt) Or svcSecurity.OverDailyLimit(CStr(args("account")), uow.BusinessDate, amt)
    If needApproval Then
        Set r = svcApproval.Request(uow, modConstants.REQ_LARGE_INTERBANK, CStr(args("account")), CStr(args("counter")), amt, p, args, user.UserID)
    Else
        Set r = ExecOut(uow, args, user.UserID, "")
    End If
    If r.Success Then uow.Commit
    Set SendOut = r
End Function

' ---- Incoming (HeadOffice types it) ----

Public Function ReceiveIn(ByVal user As clsUser, ByVal inp As Object) As clsResult
    Dim uow As clsUnitOfWork, msg As String, amt As Variant, a As clsAccount, bank As clsOtherBank, other As String
    Dim lines As Collection, jr As clsResult, grp As String, id As String, h As String, ho As String
    msg = svcSecurity.Denied(user, svcSecurity.ACT_INTERBANK_IN)
    If Len(msg) > 0 Then Set ReceiveIn = modMessages.FailResult(msg): Exit Function
    If Not modArgs.ParseAmount(inp(modConstants.INP_AMOUNT), amt, msg) Then Set ReceiveIn = modMessages.FailResult(msg): Exit Function
    Set uow = New clsUnitOfWork
    uow.Init repoConfig.BusinessDate(), user.UserID
    Set a = uow.GetAccount(CStr(inp(modConstants.INP_ACCOUNT)))
    If a Is Nothing Then Set ReceiveIn = modMessages.FailResult(modMessages.UnknownAccount(CStr(inp(modConstants.INP_ACCOUNT)))): Exit Function
    If a.Ccy <> modConstants.CCY_MMK Then Set ReceiveIn = modMessages.FailResult(modMessages.MSG_FOREIGN_CCY): Exit Function
    If Not svcAccounts.CanCredit(a, msg) Then Set ReceiveIn = modMessages.FailResult(msg): Exit Function
    Set bank = repoOtherBank.FindBank(CStr(inp(modConstants.INP_BANKCODE)))
    If bank Is Nothing Then Set ReceiveIn = modMessages.FailResult("Unknown other bank code " & CStr(inp(modConstants.INP_BANKCODE)) & "."): Exit Function
    other = CStr(inp(modConstants.INP_COUNTER))
    If Not (other Like "##########") Then Set ReceiveIn = modMessages.FailResult("The other bank account must be 10 digits (Control C6)."): Exit Function
    h = a.BranchID
    ho = modConstants.HEAD_OFFICE_ID
    grp = uow.NewGroupId()
    Set lines = New Collection
    svcLedger.AddDr lines, ho, modConstants.GL_TREASURY, amt
    If h <> ho Then
        svcLedger.AddCr lines, ho, modConstants.GL_INTERBRANCH, amt
        svcLedger.AddDr lines, h, modConstants.GL_INTERBRANCH, amt
    End If
    svcLedger.AddCr lines, h, a.DepositGL, amt, a.AccountNo
    Set jr = svcLedger.PostJournal(uow, modConstants.JT_IB_IN, "Interbank in " & modMoney.FormatAmount(amt) & " from " & bank.BankCode, grp, lines)
    If Not jr.Success Then Set ReceiveIn = jr: Exit Function
    svcAccounts.Effect uow, a, modConstants.TXN_IB_IN, modConstants.DIR_CREDIT, amt, bank.BankCode, ho, "Interbank transfer from " & bank.Name & " " & other, _
        CStr(inp(modConstants.INP_NOTE)), jr.RefID, grp, user.UserID, "", True
    id = uow.NewId(modConstants.TBL_INTERBANK, "TransferID", "IB", 6)
    uow.AppendRow modConstants.TBL_INTERBANK, repoInterbank.NewTransferRow(id, modConstants.DIR_IN, a.AccountNo, bank.BankCode, other, amt, CDec(0), _
        uow.BusinessDate, uow.BusinessDate, modConstants.STATUS_SETTLED, jr.RefID, jr.RefID)
    uow.Commit
    Set ReceiveIn = modMessages.OkResult("Interbank transfer " & id & " of " & modMoney.FormatAmount(amt) & " MMK credited to " & a.AccountNo & _
        ". New balance " & modMoney.FormatAmount(a.Balance) & ". Group " & grp & ".", 1, grp)
End Function

' ---- Decisions at end of day ----

Private Function SettleLines(ByVal h As String, ByVal amt As Variant) As Collection
    Dim lines As Collection, ho As String
    ho = modConstants.HEAD_OFFICE_ID
    Set lines = New Collection
    svcLedger.AddDr lines, h, modConstants.GL_IB_PAYABLE, amt
    If h <> ho Then
        svcLedger.AddCr lines, h, modConstants.GL_INTERBRANCH, amt
        svcLedger.AddDr lines, ho, modConstants.GL_INTERBRANCH, amt
    End If
    svcLedger.AddCr lines, ho, modConstants.GL_TREASURY, amt
    Set SettleLines = lines
End Function

Private Function FailLines(ByVal a As clsAccount, ByVal amt As Variant, ByVal fee As Variant, ByVal tax As Variant) As Collection
    Dim lines As Collection
    Set lines = New Collection
    svcLedger.AddDr lines, a.BranchID, modConstants.GL_IB_PAYABLE, amt
    svcLedger.AddDr lines, a.BranchID, modConstants.GL_FEE_INCOME, fee
    svcLedger.AddDr lines, a.BranchID, modConstants.GL_TAX_FEES, tax
    svcLedger.AddCr lines, a.BranchID, a.DepositGL, amt + fee + tax, a.AccountNo
    Set FailLines = lines
End Function

' Decides every Submitted outgoing transfer due on or before d: Rnd < FailRate = Failed (refund incl. fee), else Settled.
' Settlement does not touch the customer account, so it also works for closed or frozen accounts.
Public Function DecideDue(ByVal uow As clsUnitOfWork, ByVal d As Date, ByRef decided As Long) As clsResult
    Dim data As Variant, i As Long, due As Date, banks As Object, bank As clsOtherBank, code As String
    Dim a As clsAccount, amt As Variant, fee As Variant, jr As clsResult, grp As String, failIt As Boolean, tax As Variant
    decided = 0
    data = uow.WorkingTable(modConstants.TBL_INTERBANK)
    If IsEmpty(data) Then Set DecideDue = modMessages.OkResult("", 0): Exit Function
    Set banks = CreateObject("Scripting.Dictionary")
    For i = 1 To UBound(data, 1)
        If CStr(data(i, repoInterbank.IBC_DIR)) = modConstants.DIR_OUT And CStr(data(i, repoInterbank.IBC_STATUS)) = modConstants.STATUS_SUBMITTED Then
            If modDateUtils.TryToDate(data(i, repoInterbank.IBC_EXPECTED), due) Then
                If due <= d Then
                    code = CStr(data(i, repoInterbank.IBC_BANK))
                    If Not banks.Exists(code) Then Set banks(code) = repoOtherBank.FindBank(code)
                    Set bank = banks(code)
                    Set a = uow.GetAccount(CStr(data(i, repoInterbank.IBC_ACCOUNT)))
                    If a Is Nothing Then Set DecideDue = modMessages.FailResult("Interbank transfer " & CStr(data(i, 1)) & " has an unknown account."): Exit Function
                    amt = modMoney.ToDec(data(i, repoInterbank.IBC_AMOUNT))
                    fee = modMoney.ToDec(data(i, repoInterbank.IBC_FEE))
                    tax = CDec(0)
                    If fee > 0 Then tax = repoJournal.CreditOf(CStr(data(i, repoInterbank.IBC_SUBMITJ)), modConstants.GL_TAX_FEES)
                    failIt = False
                    If Not bank Is Nothing Then failIt = (Rnd < CDbl(bank.FailRate))
                    grp = uow.NewGroupId()
                    If failIt Then
                        Set jr = svcLedger.PostJournal(uow, modConstants.JT_IB_FAIL, "Interbank " & CStr(data(i, 1)) & " failed, refund", grp, FailLines(a, amt, fee, tax))
                        If Not jr.Success Then Set DecideDue = jr: Exit Function
                        svcAccounts.Effect uow, a, modConstants.TXN_IB_REFUND, modConstants.DIR_CREDIT, amt, code, a.BranchID, "Interbank " & CStr(data(i, 1)) & " failed, refund", "", jr.RefID, grp, "SYSTEM", "", False
                        If fee > 0 Then svcAccounts.Effect uow, a, modConstants.TXN_FEE_REFUND, modConstants.DIR_CREDIT, fee, code, a.BranchID, "Fee refund " & CStr(data(i, 1)), modConstants.FEE_IB_OUT, jr.RefID, grp, "SYSTEM", "", False
                        If tax > 0 Then svcAccounts.Effect uow, a, modConstants.TXN_FEE_REFUND, modConstants.DIR_CREDIT, tax, code, a.BranchID, "Tax refund " & CStr(data(i, 1)), "TAX", jr.RefID, grp, "SYSTEM", "", False
                        data(i, repoInterbank.IBC_STATUS) = modConstants.STATUS_FAILED
                        data(i, repoInterbank.IBC_REASON) = "Simulated failure, refunded"
                    Else
                        Set jr = svcLedger.PostJournal(uow, modConstants.JT_IB_SETTLE, "Interbank " & CStr(data(i, 1)) & " settled", grp, SettleLines(a.BranchID, amt))
                        If Not jr.Success Then Set DecideDue = jr: Exit Function
                        data(i, repoInterbank.IBC_STATUS) = modConstants.STATUS_SETTLED
                    End If
                    data(i, repoInterbank.IBC_STATUSDATE) = d
                    data(i, repoInterbank.IBC_SETTLEJ) = jr.RefID
                    decided = decided + 1
                End If
            End If
        End If
    Next i
    If decided > 0 Then uow.ReplaceTable modConstants.TBL_INTERBANK, data
    Set DecideDue = modMessages.OkResult("", decided)
End Function

' ---- Forced return of a Settled transfer (for testing) ----

Public Function ReturnTransfer(ByVal user As clsUser, ByVal inp As Object) As clsResult
    Dim uow As clsUnitOfWork, msg As String, data As Variant, row As Long, id As String, a As clsAccount, amt As Variant
    Dim lines As Collection, jr As clsResult, grp As String, h As String, ho As String
    msg = svcSecurity.Denied(user, svcSecurity.ACT_RETURN_IB)
    If Len(msg) > 0 Then Set ReturnTransfer = modMessages.FailResult(msg): Exit Function
    id = CStr(inp(modConstants.INP_REFID))
    Set uow = New clsUnitOfWork
    uow.Init repoConfig.BusinessDate(), user.UserID
    data = uow.WorkingTable(modConstants.TBL_INTERBANK)
    row = repoInterbank.FindTransfer(data, id)
    If row = 0 Then Set ReturnTransfer = modMessages.FailResult("Unknown transfer " & id & ". Put the TransferID in Control C12."): Exit Function
    If CStr(data(row, repoInterbank.IBC_DIR)) <> modConstants.DIR_OUT Or CStr(data(row, repoInterbank.IBC_STATUS)) <> modConstants.STATUS_SETTLED Then
        Set ReturnTransfer = modMessages.FailResult("Only a Settled outgoing transfer can be returned.")
        Exit Function
    End If
    Set a = uow.GetAccount(CStr(data(row, repoInterbank.IBC_ACCOUNT)))
    If a Is Nothing Then Set ReturnTransfer = modMessages.FailResult(modMessages.UnknownAccount(CStr(data(row, repoInterbank.IBC_ACCOUNT)))): Exit Function
    If Not svcAccounts.CanCredit(a, msg) Then Set ReturnTransfer = modMessages.FailResult("The return cannot be credited. " & msg): Exit Function
    amt = modMoney.ToDec(data(row, repoInterbank.IBC_AMOUNT))
    h = a.BranchID
    ho = modConstants.HEAD_OFFICE_ID
    grp = uow.NewGroupId()
    Set lines = New Collection
    svcLedger.AddDr lines, ho, modConstants.GL_TREASURY, amt
    If h <> ho Then
        svcLedger.AddCr lines, ho, modConstants.GL_INTERBRANCH, amt
        svcLedger.AddDr lines, h, modConstants.GL_INTERBRANCH, amt
    End If
    svcLedger.AddCr lines, h, a.DepositGL, amt, a.AccountNo
    Set jr = svcLedger.PostJournal(uow, modConstants.JT_IB_RETURN, "Interbank " & id & " returned", grp, lines)
    If Not jr.Success Then Set ReturnTransfer = jr: Exit Function
    svcAccounts.Effect uow, a, modConstants.TXN_IB_RETURN, modConstants.DIR_CREDIT, amt, CStr(data(row, repoInterbank.IBC_BANK)), ho, _
        "Interbank " & id & " returned (fee is kept)", "", jr.RefID, grp, user.UserID, "", False
    data(row, repoInterbank.IBC_STATUS) = modConstants.STATUS_RETURNED
    data(row, repoInterbank.IBC_STATUSDATE) = uow.BusinessDate
    data(row, repoInterbank.IBC_REASON) = "Returned (forced)"
    data(row, repoInterbank.IBC_RETURNJ) = jr.RefID
    uow.ReplaceTable modConstants.TBL_INTERBANK, data
    repoAudit.Add uow, "InterbankReturn", id, "Returned " & modMoney.FormatAmount(amt)
    uow.Commit
    Set ReturnTransfer = modMessages.OkResult("Transfer " & id & " was returned. " & modMoney.FormatAmount(amt) & " MMK credited to " & a.AccountNo & ". Group " & grp & ".", 1, grp)
End Function
