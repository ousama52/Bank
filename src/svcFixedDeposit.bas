Attribute VB_Name = "svcFixedDeposit"
Option Explicit

' Fixed deposits: opening (rate locked), maturity at end of day, early break (needs approval).

' args: customer, product, term, amount, payout, counter (funding account, optional), branch (processing), note.
Public Function ExecOpenFD(ByVal uow As clsUnitOfWork, ByVal args As Object, ByVal createdBy As String, ByVal approvedBy As String) As clsResult
    Dim cust As clsCustomer, prod As clsProduct, a As clsAccount, pay As clsAccount, amt As Variant, msg As String
    Dim term As Long, rate As Variant, grp As String, r As clsResult, payNo As String
    amt = modArgs.Dec(args, "amount")
    Set cust = repoCustomer.FindCustomer(modArgs.Text(args, "customer"))
    Set prod = repoProduct.FindProduct(modArgs.Text(args, "product"))
    msg = svcAccounts.OpeningProblem(cust, prod, amt, modArgs.Text(args, "customer"), modArgs.Text(args, "product"))
    If Len(msg) = 0 Then
        If prod.ProductType <> modConstants.PT_FIXED Then msg = "Product " & prod.ProductCode & " is not a fixed deposit product."
    End If
    If Len(msg) > 0 Then Set ExecOpenFD = modMessages.FailResult(msg): Exit Function
    term = CLng(modArgs.Dec(args, "term"))
    If term <= 0 Or Not repoRate.TermExists(prod.ProductCode, term) Then
        Set ExecOpenFD = modMessages.FailResult("A term of " & term & " months is not offered for " & prod.ProductCode & " (Control C10).")
        Exit Function
    End If
    If Not repoRate.Lookup(prod.ProductCode, term, amt, uow.BusinessDate, rate) Then
        Set ExecOpenFD = modMessages.FailResult("No interest rate is set for this term and amount on " & modDateUtils.ToIso(uow.BusinessDate) & ".")
        Exit Function
    End If
    payNo = modArgs.Text(args, "payout")
    Set pay = uow.GetAccount(payNo)
    If pay Is Nothing Then Set ExecOpenFD = modMessages.FailResult("The payout account " & payNo & " (Control C11) does not exist."): Exit Function
    If pay.CustomerID <> cust.CustomerID Then Set ExecOpenFD = modMessages.FailResult("The payout account must belong to the same customer."): Exit Function
    If pay.Status <> modConstants.STATUS_ACTIVE Or pay.ProductType = modConstants.PT_FIXED Then
        Set ExecOpenFD = modMessages.FailResult("The payout account must be an Active savings or current account.")
        Exit Function
    End If
    grp = uow.NewGroupId()
    Set a = svcAccounts.CreateAccount(uow, cust, prod)
    a.TermMonths = term
    a.MaturityDate = modDateUtils.AddMonths(uow.BusinessDate, term)
    a.LockedRate = rate
    a.PayoutAccountNo = payNo
    Set r = svcAccounts.FundOpening(uow, a, amt, modArgs.Text(args, "branch"), modArgs.Text(args, "counter"), grp, modArgs.Text(args, "note"), createdBy, approvedBy)
    If Not r.Success Then Set ExecOpenFD = r: Exit Function
    repoAudit.Add uow, "OpenFixedDeposit", a.AccountNo, "Principal " & modMoney.FormatAmount(amt) & ", " & term & " months, matures " & modDateUtils.ToIso(CDate(a.MaturityDate))
    Set ExecOpenFD = modMessages.OkResult("Fixed deposit " & a.AccountNo & " opened: " & modMoney.FormatAmount(amt) & " MMK for " & term & " months, rate " & _
        Format$(CDbl(rate) * 100, "0.00") & "%, matures " & modDateUtils.ToIso(CDate(a.MaturityDate)) & ". Group " & grp & ".", 1, grp)
End Function

' Control: C13 customer, C14 product, C10 term, C7 principal, C11 payout account, C6 funding account (blank = cash at C8).
Public Function OpenFixedDeposit(ByVal user As clsUser, ByVal inp As Object) As clsResult
    Dim uow As clsUnitOfWork, msg As String, amt As Variant, args As Object, p As String, r As clsResult
    msg = svcSecurity.Denied(user, svcSecurity.ACT_OPEN)
    If Len(msg) > 0 Then Set OpenFixedDeposit = modMessages.FailResult(msg): Exit Function
    If Not modArgs.ParseAmount(inp(modConstants.INP_AMOUNT), amt, msg) Then Set OpenFixedDeposit = modMessages.FailResult(msg): Exit Function
    Set uow = New clsUnitOfWork
    uow.Init repoConfig.BusinessDate(), user.UserID
    If Not svcSecurity.ResolveBranch(uow, user, CStr(inp(modConstants.INP_BRANCH)), p, msg) Then Set OpenFixedDeposit = modMessages.FailResult(msg): Exit Function
    Set args = modArgs.FromInputs(inp)
    args("amount") = amt
    args("branch") = p
    If svcSecurity.OverTellerLimit(amt) Then
        args("open") = "FD"
        Set r = svcApproval.Request(uow, modConstants.REQ_LARGE_DEPOSIT, "", CStr(args("counter")), amt, p, args, user.UserID)
    Else
        Set r = ExecOpenFD(uow, args, user.UserID, "")
    End If
    If r.Success Then uow.Commit
    Set OpenFixedDeposit = r
End Function

' ---- Maturity (end of day) ----

' Matures every Active fixed deposit with MaturityDate <= d: pays the interest, then moves the whole balance to the payout account.
' If the payout account is not Active the deposit stays and a Log warning is written (retry next day).
Public Function Mature(ByVal uow As clsUnitOfWork, ByVal d As Date, ByRef matured As Long) As clsResult
    Dim k As Variant, a As clsAccount, pay As clsAccount, g As Variant, w As Variant, grp As String, r As clsResult
    matured = 0
    For Each k In uow.Accounts().Keys
        Set a = uow.Accounts()(k)
        If a.ProductType = modConstants.PT_FIXED And a.Status = modConstants.STATUS_ACTIVE And Not IsEmpty(a.MaturityDate) Then
            If CDate(a.MaturityDate) <= d Then
                Set pay = uow.GetAccount(a.PayoutAccountNo)
                If pay Is Nothing Then
                    clsLogger.Warn "svcFixedDeposit.Mature", "Fixed deposit " & a.AccountNo & " kept: payout account " & a.PayoutAccountNo & " does not exist."
                ElseIf pay.Status <> modConstants.STATUS_ACTIVE Then
                    clsLogger.Warn "svcFixedDeposit.Mature", "Fixed deposit " & a.AccountNo & " kept: payout account " & pay.AccountNo & " is " & pay.Status & ". Retry next day."
                Else
                    grp = uow.NewGroupId()
                    g = modMoney.RoundCcy(a.AccruedInterest, a.Ccy)
                    If g >= 1 Then
                        w = svcInterest.TaxFor(repoCustomer.FindCustomer(a.CustomerID), g)
                        Set r = svcInterest.PostInterestPosting(uow, a, g, w, modConstants.JT_FD_MATURITY, grp, "SYSTEM")
                        If Not r.Success Then Set Mature = r: Exit Function
                    End If
                    a.AccruedInterest = CDec(0)
                    Set r = svcTransfers.PostTransfer(uow, a, pay, a.Balance, a.BranchID, grp, modConstants.TXN_FD_MATURITY, modConstants.TXN_FD_MATURITY, _
                        modConstants.JT_FD_MATURITY, "Fixed deposit maturity " & a.AccountNo, "", "SYSTEM", "", True)
                    If Not r.Success Then Set Mature = r: Exit Function
                    a.Status = modConstants.STATUS_CLOSED
                    a.ClosedDate = d
                    a.Dirty = True
                    repoAudit.Add uow, "FDMaturity", a.AccountNo, "Matured, paid to " & pay.AccountNo
                    matured = matured + 1
                End If
            End If
        End If
    Next k
    Set Mature = modMessages.OkResult("", matured)
End Function

' ---- Early break ----

' args: account (the fixed deposit), note. Interest = principal x EarlyBreakRate x days held / DayCountBasis, then the FD_EARLY_BREAK fee on the principal.
Public Function ExecBreak(ByVal uow As clsUnitOfWork, ByVal args As Object, ByVal createdBy As String, ByVal approvedBy As String) As clsResult
    Dim a As clsAccount, pay As clsAccount, principal As Variant, days As Long, g As Variant, w As Variant, fee As Variant, tax As Variant
    Dim msg As String, grp As String, r As clsResult, lines As Collection, jr As clsResult, restAmt As Variant
    Set a = uow.GetAccount(modArgs.Text(args, "account"))
    If a Is Nothing Then Set ExecBreak = modMessages.FailResult(modMessages.UnknownAccount(modArgs.Text(args, "account"))): Exit Function
    If a.ProductType <> modConstants.PT_FIXED Then Set ExecBreak = modMessages.FailResult("Account " & a.AccountNo & " is not a fixed deposit."): Exit Function
    If a.Status <> modConstants.STATUS_ACTIVE Then Set ExecBreak = modMessages.FailResult(modMessages.AccountStatusRefuses(a.AccountNo, a.Status)): Exit Function
    Set pay = uow.GetAccount(a.PayoutAccountNo)
    If pay Is Nothing Then Set ExecBreak = modMessages.FailResult("The payout account " & a.PayoutAccountNo & " does not exist."): Exit Function
    If pay.Status <> modConstants.STATUS_ACTIVE Then Set ExecBreak = modMessages.FailResult("The payout account " & pay.AccountNo & " is " & pay.Status & ", not Active."): Exit Function
    principal = a.Balance
    days = modDateUtils.DaysBetween(CDate(a.OpenDate), uow.BusinessDate)
    g = modMoney.RoundCcy(principal * repoConfig.GetDec(modConstants.CFG_EARLYBREAK) * days / repoConfig.GetDec(modConstants.CFG_DAYCOUNT), a.Ccy)
    grp = uow.NewGroupId()
    If g >= 1 Then
        w = svcInterest.TaxFor(repoCustomer.FindCustomer(a.CustomerID), g)
        Set r = svcInterest.PostInterestPosting(uow, a, g, w, modConstants.JT_FD_BREAK, grp, createdBy)
        If Not r.Success Then Set ExecBreak = r: Exit Function
    End If
    a.AccruedInterest = CDec(0)
    If Not svcFees.Compute(modConstants.FEE_FD_BREAK, principal, fee, tax, msg) Then Set ExecBreak = modMessages.FailResult(msg): Exit Function
    If fee + tax > a.Balance Then Set ExecBreak = modMessages.FailResult(modMessages.MSG_NO_BALANCE): Exit Function
    If fee + tax > 0 Then
        Set lines = New Collection
        svcLedger.AddDr lines, a.BranchID, a.DepositGL, fee + tax, a.AccountNo
        svcLedger.AddCr lines, a.BranchID, modConstants.GL_FEE_INCOME, fee
        svcLedger.AddCr lines, a.BranchID, modConstants.GL_TAX_FEES, tax
        Set jr = svcLedger.PostJournal(uow, modConstants.JT_FD_BREAK, "Early break fee " & a.AccountNo, grp, lines)
        If Not jr.Success Then Set ExecBreak = jr: Exit Function
        svcAccounts.FeeRows uow, a, fee, tax, modConstants.FEE_FD_BREAK, "", a.BranchID, jr.RefID, grp, createdBy, approvedBy
    End If
    restAmt = a.Balance
    Set r = svcTransfers.PostTransfer(uow, a, pay, restAmt, a.BranchID, grp, modConstants.TXN_FD_BREAK, modConstants.TXN_FD_BREAK, _
        modConstants.JT_FD_BREAK, "Fixed deposit early break " & a.AccountNo, modArgs.Text(args, "note"), createdBy, approvedBy, True)
    If Not r.Success Then Set ExecBreak = r: Exit Function
    a.Status = modConstants.STATUS_CLOSED
    a.ClosedDate = uow.BusinessDate
    a.Dirty = True
    repoAudit.Add uow, "FDBreak", a.AccountNo, "Early break after " & days & " days, interest " & modMoney.FormatAmount(g) & ", fee " & modMoney.FormatAmount(fee) & ", approved by " & approvedBy
    Set ExecBreak = modMessages.OkResult("Fixed deposit " & a.AccountNo & " broken early after " & days & " days. Interest " & modMoney.FormatAmount(g) & _
        ", fee " & modMoney.FormatAmount(fee + tax) & ", " & modMoney.FormatAmount(restAmt) & " MMK paid to " & pay.AccountNo & ". Group " & grp & ".", 1, grp)
End Function

' Control: C5 = the fixed deposit. Creates a Pending approval (BranchManager or HeadOffice approves).
Public Function BreakFixedDeposit(ByVal user As clsUser, ByVal inp As Object) As clsResult
    Dim uow As clsUnitOfWork, msg As String, a As clsAccount, args As Object, r As clsResult
    msg = svcSecurity.Denied(user, svcSecurity.ACT_BREAKFD)
    If Len(msg) > 0 Then Set BreakFixedDeposit = modMessages.FailResult(msg): Exit Function
    Set uow = New clsUnitOfWork
    uow.Init repoConfig.BusinessDate(), user.UserID
    Set a = uow.GetAccount(CStr(inp(modConstants.INP_ACCOUNT)))
    If a Is Nothing Then Set BreakFixedDeposit = modMessages.FailResult(modMessages.UnknownAccount(CStr(inp(modConstants.INP_ACCOUNT)))): Exit Function
    If a.ProductType <> modConstants.PT_FIXED Or a.Status <> modConstants.STATUS_ACTIVE Then
        Set BreakFixedDeposit = modMessages.FailResult("Account " & a.AccountNo & " is not an Active fixed deposit.")
        Exit Function
    End If
    If Not svcSecurity.CanManageAccount(user, a) Then Set BreakFixedDeposit = modMessages.FailResult("This deposit belongs to branch " & a.BranchID & "."): Exit Function
    Set args = modArgs.NewArgs()
    args("note") = CStr(inp(modConstants.INP_NOTE))
    Set r = svcApproval.Request(uow, modConstants.REQ_BREAK_FD, a.AccountNo, "", a.Balance, a.BranchID, args, user.UserID)
    uow.Commit
    Set BreakFixedDeposit = r
End Function
