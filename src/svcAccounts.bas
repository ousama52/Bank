Attribute VB_Name = "svcAccounts"
Option Explicit

' Account rules: who may be credited or debited, statement rows, opening, status changes, closing, dormancy.

' ---- Rules ----

' Credits are allowed on Active and Frozen accounts.
' Set to True by the wallet services while they work, so wallets cannot be used by the ordinary cash macros.
Public WalletContext As Boolean

Public Function CanCredit(ByVal a As clsAccount, ByRef msg As String) As Boolean
    If a.ProductType = modConstants.PT_WALLET And Not WalletContext Then msg = "A wallet is used with the wallet macros only.": Exit Function
    If a.Status = modConstants.STATUS_ACTIVE Or a.Status = modConstants.STATUS_FROZEN Then
        CanCredit = True
    Else
        msg = modMessages.AccountStatusRefuses(a.AccountNo, a.Status)
    End If
End Function

' Debits need an Active account, enough balance and (for customers) the minimum balance left.
' isSystem = maturity, break and closing: fixed deposits may be debited and the minimum balance is not checked.
Public Function CanDebit(ByVal a As clsAccount, ByVal total As Variant, ByVal isSystem As Boolean, ByRef msg As String) As Boolean
    If a.Status <> modConstants.STATUS_ACTIVE Then
        msg = modMessages.AccountStatusRefuses(a.AccountNo, a.Status)
    ElseIf a.ProductType = modConstants.PT_WALLET And Not WalletContext Then
        msg = "A wallet is used with the wallet macros only."
    ElseIf Not isSystem And a.ProductType = modConstants.PT_FIXED Then
        msg = modMessages.MSG_FD_NO_DEBIT
    ElseIf total > a.Balance Then
        msg = modMessages.MSG_NO_BALANCE
    ElseIf Not isSystem And a.Balance - total < a.MinBalance Then
        msg = modMessages.MSG_MIN_BALANCE
    Else
        CanDebit = True
    End If
End Function

' Changes the balance and adds one statement row. customerActivity = True moves LastActivityDate
' (never for interest, fees, tax or maintenance).
Public Sub Effect(ByVal uow As clsUnitOfWork, ByVal a As clsAccount, ByVal txnType As String, ByVal direction As String, _
                  ByVal amount As Variant, ByVal counter As String, ByVal branchId As String, ByVal description As String, _
                  ByVal reference As String, ByVal journalId As String, ByVal groupId As String, ByVal createdBy As String, _
                  ByVal approvedBy As String, ByVal customerActivity As Boolean)
    Dim t As clsTxnRow
    If direction = modConstants.DIR_CREDIT Then
        a.Balance = a.Balance + amount
    Else
        a.Balance = a.Balance - amount
    End If
    a.Dirty = True
    If customerActivity Then a.LastActivityDate = uow.BusinessDate
    Set t = New clsTxnRow
    t.TxnID = uow.NewTxnId()
    t.GroupID = groupId
    t.TxnDate = uow.BusinessDate
    t.AccountNo = a.AccountNo
    t.TxnType = txnType
    t.Direction = direction
    t.Amount = amount
    t.BalanceAfter = a.Balance
    t.CounterAccountNo = counter
    t.BranchID = branchId
    t.Description = description
    t.Reference = reference
    t.JournalID = journalId
    t.Status = modConstants.STATUS_POSTED
    t.CreatedBy = createdBy
    t.ApprovedBy = approvedBy
    uow.AppendRow modConstants.TBL_TRANSACTIONS, repoTransaction.TxnToRow(t)
End Sub

' Statement rows for a fee and its tax (Debit). Tax rows only exist from Stage 2.
Public Sub FeeRows(ByVal uow As clsUnitOfWork, ByVal a As clsAccount, ByVal fee As Variant, ByVal tax As Variant, ByVal feeCode As String, _
                   ByVal counter As String, ByVal branchId As String, ByVal journalId As String, ByVal groupId As String, _
                   ByVal createdBy As String, ByVal approvedBy As String)
    If fee > 0 Then Effect uow, a, modConstants.TXN_FEE, modConstants.DIR_DEBIT, fee, counter, branchId, "Fee " & feeCode, feeCode, journalId, groupId, createdBy, approvedBy, False
    If tax > 0 Then Effect uow, a, modConstants.TXN_TAX, modConstants.DIR_DEBIT, tax, counter, branchId, "Tax on " & feeCode, feeCode, journalId, groupId, createdBy, approvedBy, False
End Sub

' ---- Opening ----

Public Function CreateAccount(ByVal uow As clsUnitOfWork, ByVal cust As clsCustomer, ByVal prod As clsProduct) As clsAccount
    Dim a As clsAccount
    Set a = New clsAccount
    a.AccountNo = modIds.NextAccountNo(cust.HomeBranchID)
    a.CustomerID = cust.CustomerID
    a.BranchID = cust.HomeBranchID
    a.ProductCode = prod.ProductCode
    a.Ccy = prod.Ccy
    a.Status = modConstants.STATUS_ACTIVE
    a.OpenDate = uow.BusinessDate
    a.Balance = CDec(0)
    a.AccruedInterest = CDec(0)
    a.LastActivityDate = uow.BusinessDate
    a.ProductType = prod.ProductType
    a.DepositGL = repoProduct.DepositGL(prod.ProductType)
    a.MinBalance = prod.MinBalance
    uow.AddAccount a
    Set CreateAccount = a
End Function

' Checks shared by savings, current and fixed opening. Returns "" or the refusal text.
Public Function OpeningProblem(ByVal cust As clsCustomer, ByVal prod As clsProduct, ByVal amount As Variant, ByVal customerId As String, ByVal productCode As String) As String
    If cust Is Nothing Then OpeningProblem = modMessages.UnknownCustomer(customerId): Exit Function
    If cust.Status <> modConstants.STATUS_ACTIVE Then OpeningProblem = "Customer " & cust.CustomerID & " is not Active.": Exit Function
    If prod Is Nothing Then OpeningProblem = "Unknown product " & productCode & ".": Exit Function
    If Not prod.IsActive Then OpeningProblem = "Product " & productCode & " is not active.": Exit Function
    If prod.ProductType = modConstants.PT_WALLET Then OpeningProblem = "Use RunOpenWallet for a wallet.": Exit Function
    If amount < prod.MinOpening Then OpeningProblem = "The opening deposit must be at least " & modMoney.FormatAmount(prod.MinOpening) & "."
End Function

' Puts the opening deposit into the new account: cash at processing branch P, or a transfer from another account of the same customer.
Public Function FundOpening(ByVal uow As clsUnitOfWork, ByVal a As clsAccount, ByVal amount As Variant, ByVal p As String, _
                            ByVal fundingNo As String, ByVal grp As String, ByVal note As String, ByVal createdBy As String, _
                            ByVal approvedBy As String) As clsResult
    Dim src As clsAccount, args As Object, r As clsResult
    If Len(fundingNo) = 0 Then
        Set args = modArgs.NewArgs()
        args("account") = a.AccountNo
        args("amount") = amount
        args("branch") = p
        args("note") = note
        args("txntype") = modConstants.TXN_OPENING
        args("jtype") = modConstants.JT_OPENING
        args("group") = grp
        Set FundOpening = svcCashOps.ExecDeposit(uow, args, createdBy, approvedBy)
    Else
        Set src = uow.GetAccount(fundingNo)
        If src Is Nothing Then
            Set FundOpening = modMessages.FailResult(modMessages.UnknownAccount(fundingNo))
        ElseIf src.CustomerID <> a.CustomerID Then
            Set FundOpening = modMessages.FailResult("The funding account must belong to the same customer.")
        Else
            Set FundOpening = svcTransfers.PostTransfer(uow, src, a, amount, p, grp, modConstants.TXN_TRANSFER_OUT, modConstants.TXN_OPENING, _
                modConstants.JT_OPENING, "Opening deposit", note, createdBy, approvedBy, False)
        End If
    End If
End Function

' Open a savings or current account. args: customer, product, amount, branch (processing), counter (funding account), note.
Public Function ExecOpen(ByVal uow As clsUnitOfWork, ByVal args As Object, ByVal createdBy As String, ByVal approvedBy As String) As clsResult
    Dim cust As clsCustomer, prod As clsProduct, a As clsAccount, amt As Variant, msg As String, grp As String, r As clsResult
    amt = modArgs.Dec(args, "amount")
    Set cust = repoCustomer.FindCustomer(modArgs.Text(args, "customer"))
    Set prod = repoProduct.FindProduct(modArgs.Text(args, "product"))
    msg = OpeningProblem(cust, prod, amt, modArgs.Text(args, "customer"), modArgs.Text(args, "product"))
    If Len(msg) = 0 Then
        If prod.ProductType = modConstants.PT_FIXED Then msg = "Use RunOpenFixedDeposit for a fixed deposit product."
    End If
    If Len(msg) = 0 Then
        If modMoney.RoundTo(amt, modMoney.Decimals(prod.Ccy)) <> amt Then msg = "A " & prod.Ccy & " amount can have at most " & modMoney.Decimals(prod.Ccy) & " decimals."
    End If
    If Len(msg) > 0 Then Set ExecOpen = modMessages.FailResult(msg): Exit Function
    grp = uow.NewGroupId()
    Set a = CreateAccount(uow, cust, prod)
    Set r = FundOpening(uow, a, amt, modArgs.Text(args, "branch"), modArgs.Text(args, "counter"), grp, modArgs.Text(args, "note"), createdBy, approvedBy)
    If Not r.Success Then Set ExecOpen = r: Exit Function
    repoAudit.Add uow, "OpenAccount", a.AccountNo, "Customer " & cust.CustomerID & ", product " & prod.ProductCode & ", opening " & modMoney.FormatAmount(amt)
    Set ExecOpen = modMessages.OkResult("Account " & a.AccountNo & " opened with " & modMoney.FormatAmount(amt) & " MMK. Group " & grp & ".", 1, grp)
End Function

Public Function OpenAccount(ByVal user As clsUser, ByVal inp As Object) As clsResult
    Dim uow As clsUnitOfWork, msg As String, amt As Variant, args As Object, p As String, r As clsResult
    msg = svcSecurity.Denied(user, svcSecurity.ACT_OPEN)
    If Len(msg) > 0 Then Set OpenAccount = modMessages.FailResult(msg): Exit Function
    If Not modArgs.ParseAmountAny(inp(modConstants.INP_AMOUNT), amt, msg) Then Set OpenAccount = modMessages.FailResult(msg): Exit Function
    Set uow = New clsUnitOfWork
    uow.Init repoConfig.BusinessDate(), user.UserID
    If Not svcSecurity.ResolveBranch(uow, user, CStr(inp(modConstants.INP_BRANCH)), p, msg) Then Set OpenAccount = modMessages.FailResult(msg): Exit Function
    Set args = modArgs.FromInputs(inp)
    args("amount") = amt
    args("branch") = p
    If svcSecurity.OverTellerLimit(amt) Then
        args("open") = "ACC"
        Set r = svcApproval.Request(uow, modConstants.REQ_LARGE_DEPOSIT, "", CStr(args("counter")), amt, p, args, user.UserID)
    Else
        Set r = ExecOpen(uow, args, user.UserID, "")
    End If
    If r.Success Then uow.Commit
    Set OpenAccount = r
End Function

' ---- Status changes ----

Private Function FindForStatus(ByVal uow As clsUnitOfWork, ByVal user As clsUser, ByVal accountNo As String, ByRef a As clsAccount, ByRef msg As String) As Boolean
    Set a = uow.GetAccount(accountNo)
    If a Is Nothing Then
        msg = modMessages.UnknownAccount(accountNo)
    ElseIf Not svcSecurity.CanManageAccount(user, a) Then
        msg = "This account belongs to branch " & a.BranchID & ". Your role can only manage accounts of your own branch."
    Else
        FindForStatus = True
    End If
End Function

Public Function Freeze(ByVal user As clsUser, ByVal inp As Object) As clsResult
    Dim uow As clsUnitOfWork, a As clsAccount, msg As String
    msg = svcSecurity.Denied(user, svcSecurity.ACT_FREEZE)
    If Len(msg) > 0 Then Set Freeze = modMessages.FailResult(msg): Exit Function
    Set uow = New clsUnitOfWork
    uow.Init repoConfig.BusinessDate(), user.UserID
    If Not FindForStatus(uow, user, CStr(inp(modConstants.INP_ACCOUNT)), a, msg) Then Set Freeze = modMessages.FailResult(msg): Exit Function
    If a.Status <> modConstants.STATUS_ACTIVE And a.Status <> modConstants.STATUS_DORMANT Then
        Set Freeze = modMessages.FailResult(modMessages.AccountStatusRefuses(a.AccountNo, a.Status))
        Exit Function
    End If
    repoAudit.Add uow, "Freeze", a.AccountNo, "Status " & a.Status & " -> Frozen. " & CStr(inp(modConstants.INP_NOTE))
    a.Status = modConstants.STATUS_FROZEN
    a.Dirty = True
    uow.Commit
    Set Freeze = modMessages.OkResult("Account " & a.AccountNo & " is now Frozen. Credits are allowed, debits are refused.", 1)
End Function

' Unfreeze, Reactivate and Close all work the same way at request time: create ONE Pending approval.
Private Function RequestStatusAction(ByVal user As clsUser, ByVal inp As Object, ByVal action As String, ByVal requestType As String, _
                                     ByVal needStatus As String) As clsResult
    Dim uow As clsUnitOfWork, a As clsAccount, msg As String, args As Object, r As clsResult
    msg = svcSecurity.Denied(user, action)
    If Len(msg) > 0 Then Set RequestStatusAction = modMessages.FailResult(msg): Exit Function
    Set uow = New clsUnitOfWork
    uow.Init repoConfig.BusinessDate(), user.UserID
    If Not FindForStatus(uow, user, CStr(inp(modConstants.INP_ACCOUNT)), a, msg) Then Set RequestStatusAction = modMessages.FailResult(msg): Exit Function
    If a.Status <> needStatus Then
        Set RequestStatusAction = modMessages.FailResult(modMessages.AccountStatusRefuses(a.AccountNo, a.Status) & " Required status: " & needStatus & ".")
        Exit Function
    End If
    If requestType = modConstants.REQ_CLOSE_ACCOUNT And a.ProductType = modConstants.PT_FIXED Then
        Set RequestStatusAction = modMessages.FailResult("A fixed deposit is closed by maturity or early break, not by closing.")
        Exit Function
    End If
    Set args = modArgs.NewArgs()
    args("note") = CStr(inp(modConstants.INP_NOTE))
    Set r = svcApproval.Request(uow, requestType, a.AccountNo, "", a.Balance, a.BranchID, args, user.UserID)
    uow.Commit
    Set RequestStatusAction = r
End Function

Public Function Unfreeze(ByVal user As clsUser, ByVal inp As Object) As clsResult
    Set Unfreeze = RequestStatusAction(user, inp, svcSecurity.ACT_UNFREEZE, modConstants.REQ_UNFREEZE, modConstants.STATUS_FROZEN)
End Function

Public Function Reactivate(ByVal user As clsUser, ByVal inp As Object) As clsResult
    Set Reactivate = RequestStatusAction(user, inp, svcSecurity.ACT_REACTIVATE, modConstants.REQ_REACTIVATE, modConstants.STATUS_DORMANT)
End Function

Public Function CloseAccount(ByVal user As clsUser, ByVal inp As Object) As clsResult
    Set CloseAccount = RequestStatusAction(user, inp, svcSecurity.ACT_CLOSE, modConstants.REQ_CLOSE_ACCOUNT, modConstants.STATUS_ACTIVE)
End Function

' Run when an approval is approved. The status is checked again because it may have changed.
Public Function ExecUnfreeze(ByVal uow As clsUnitOfWork, ByVal args As Object, ByVal createdBy As String, ByVal approvedBy As String) As clsResult
    Dim a As clsAccount
    Set a = uow.GetAccount(modArgs.Text(args, "account"))
    If a Is Nothing Then Set ExecUnfreeze = modMessages.FailResult(modMessages.UnknownAccount(modArgs.Text(args, "account"))): Exit Function
    If a.Status <> modConstants.STATUS_FROZEN Then Set ExecUnfreeze = modMessages.FailResult("Account " & a.AccountNo & " is " & a.Status & ", not Frozen."): Exit Function
    a.Status = modConstants.STATUS_ACTIVE
    a.Dirty = True
    repoAudit.Add uow, "Unfreeze", a.AccountNo, "Frozen -> Active, approved by " & approvedBy
    Set ExecUnfreeze = modMessages.OkResult("Account " & a.AccountNo & " is Active again.", 1, "")
End Function

' Reactivation also resets LastActivityDate to today, otherwise the next end of day would mark the account Dormant again.
Public Function ExecReactivate(ByVal uow As clsUnitOfWork, ByVal args As Object, ByVal createdBy As String, ByVal approvedBy As String) As clsResult
    Dim a As clsAccount
    Set a = uow.GetAccount(modArgs.Text(args, "account"))
    If a Is Nothing Then Set ExecReactivate = modMessages.FailResult(modMessages.UnknownAccount(modArgs.Text(args, "account"))): Exit Function
    If a.Status <> modConstants.STATUS_DORMANT Then Set ExecReactivate = modMessages.FailResult("Account " & a.AccountNo & " is " & a.Status & ", not Dormant."): Exit Function
    a.Status = modConstants.STATUS_ACTIVE
    a.LastActivityDate = uow.BusinessDate
    a.Dirty = True
    repoAudit.Add uow, "Reactivate", a.AccountNo, "Dormant -> Active, approved by " & approvedBy
    Set ExecReactivate = modMessages.OkResult("Account " & a.AccountNo & " is Active again.", 1, "")
End Function

' Closing: pay pending interest, pay the whole balance out as cash at the home branch, set Closed.
Public Function ExecClose(ByVal uow As clsUnitOfWork, ByVal args As Object, ByVal createdBy As String, ByVal approvedBy As String) As clsResult
    Dim a As clsAccount, grp As String, g As Variant, w As Variant, lines As Collection, jr As clsResult, ib As Variant, i As Long
    Set a = uow.GetAccount(modArgs.Text(args, "account"))
    If a Is Nothing Then Set ExecClose = modMessages.FailResult(modMessages.UnknownAccount(modArgs.Text(args, "account"))): Exit Function
    If a.Status <> modConstants.STATUS_ACTIVE Then Set ExecClose = modMessages.FailResult(modMessages.AccountStatusRefuses(a.AccountNo, a.Status)): Exit Function
    If a.ProductType = modConstants.PT_FIXED Then Set ExecClose = modMessages.FailResult(modMessages.MSG_FD_NO_DEBIT): Exit Function
    ib = uow.WorkingTable(modConstants.TBL_INTERBANK)
    If Not IsEmpty(ib) Then
        For i = 1 To UBound(ib, 1)
            If CStr(ib(i, repoInterbank.IBC_ACCOUNT)) = a.AccountNo And CStr(ib(i, repoInterbank.IBC_STATUS)) = modConstants.STATUS_SUBMITTED Then
                Set ExecClose = modMessages.FailResult("Account " & a.AccountNo & " has an interbank transfer that is still Submitted. Close it after it is decided.")
                Exit Function
            End If
        Next i
    End If
    grp = uow.NewGroupId()
    g = modMoney.RoundCcy(a.AccruedInterest, a.Ccy)
    If g >= svcInterest.MinUnit(a.Ccy) Then
        w = svcInterest.TaxFor(repoCustomer.FindCustomer(a.CustomerID), g, a.Ccy)
        Set jr = svcInterest.PostInterestPosting(uow, a, g, w, modConstants.JT_INTEREST, grp, createdBy)
        If Not jr.Success Then Set ExecClose = jr: Exit Function
    End If
    a.AccruedInterest = CDec(0)
    If a.Balance > 0 And a.Ccy <> modConstants.CCY_MMK Then
        Set jr = svcFx.FxWithdraw(uow, a, a.Balance, a.BranchID, grp, "Account closure " & a.AccountNo, "", createdBy, approvedBy, True)
        If Not jr.Success Then Set ExecClose = jr: Exit Function
    ElseIf a.Balance > 0 Then
        If uow.GLNet(a.BranchID, modConstants.GL_VAULT) - a.Balance < 0 Then Set ExecClose = modMessages.FailResult(modMessages.MSG_VAULT_SHORT): Exit Function
        Set lines = New Collection
        svcLedger.AddDr lines, a.BranchID, a.DepositGL, a.Balance, a.AccountNo
        svcLedger.AddCr lines, a.BranchID, modConstants.GL_VAULT, a.Balance
        Set jr = svcLedger.PostJournal(uow, modConstants.JT_WITHDRAWAL, "Account closure " & a.AccountNo, grp, lines)
        If Not jr.Success Then Set ExecClose = jr: Exit Function
        Effect uow, a, modConstants.TXN_WITHDRAWAL, modConstants.DIR_DEBIT, a.Balance, "", a.BranchID, "Account closure, cash paid out", "", jr.RefID, grp, createdBy, approvedBy, True
    End If
    a.Status = modConstants.STATUS_CLOSED
    a.ClosedDate = uow.BusinessDate
    a.Dirty = True
    repoAudit.Add uow, "CloseAccount", a.AccountNo, "Closed, approved by " & approvedBy
    Set ExecClose = modMessages.OkResult("Account " & a.AccountNo & " is Closed. Pending interest was paid and the balance was paid out in cash. Group " & grp & ".", 1, grp)
End Function

' ---- End of day: dormancy ----

' Active savings and current accounts without customer activity for more than DormantDays become Dormant.
Public Function MarkDormant(ByVal uow As clsUnitOfWork, ByVal d As Date) As Long
    Dim limitDays As Long, k As Variant, a As clsAccount, last As Variant
    limitDays = repoConfig.GetLong(modConstants.CFG_DORMANTDAYS)
    For Each k In uow.Accounts().Keys
        Set a = uow.Accounts()(k)
        If a.Status = modConstants.STATUS_ACTIVE And a.ProductType <> modConstants.PT_FIXED Then
            last = a.LastActivityDate
            If IsEmpty(last) Then last = a.OpenDate
            If Not IsEmpty(last) Then
                If modDateUtils.DaysBetween(CDate(last), d) > limitDays Then
                    a.Status = modConstants.STATUS_DORMANT
                    a.Dirty = True
                    repoAudit.Add uow, "Dormant", a.AccountNo, "No customer activity since " & modDateUtils.ToIso(CDate(last))
                    MarkDormant = MarkDormant + 1
                End If
            End If
        End If
    Next k
End Function
