Attribute VB_Name = "svcLoanPayments"
Option Explicit

' Stage 3: repayments (penalty, then interest, then principal), payoff, write-off and recovery.
' The loans and schedule tables are handled as arrays inside the unit of work and handed back with ReplaceTable.

' ---- Allocation core ----

Private Function ScheduleRows(ByVal sched As Variant, ByVal loanId As String) As Collection
    Dim c As Collection, i As Long
    Set c = New Collection
    If Not IsEmpty(sched) Then
        For i = 1 To UBound(sched, 1)
            If CStr(sched(i, repoLoan.SCC_LOAN)) = loanId Then c.Add i
        Next i
    End If
    Set ScheduleRows = c
End Function

Private Function IsDue(ByVal status As String) As Boolean
    IsDue = (status = "Due" Or status = "Partial")
End Function

' Recognized interest not yet paid (rows that are Due or Partial).
Public Function UnpaidInterest(ByVal sched As Variant, ByVal loanId As String) As Variant
    Dim i As Long, total As Variant
    total = CDec(0)
    If Not IsEmpty(sched) Then
        For i = 1 To UBound(sched, 1)
            If CStr(sched(i, repoLoan.SCC_LOAN)) = loanId And IsDue(CStr(sched(i, repoLoan.SCC_STATUS))) Then
                total = total + modMoney.ToDec(sched(i, repoLoan.SCC_INT)) - modMoney.ToDec(sched(i, repoLoan.SCC_INTPAID))
            End If
        Next i
    End If
    UnpaidInterest = total
End Function

' Installment amount still unpaid (principal + recognized interest) of one schedule row.
Public Function RowUnpaid(ByVal sched As Variant, ByVal i As Long) As Variant
    RowUnpaid = modMoney.ToDec(sched(i, repoLoan.SCC_PRIN)) + modMoney.ToDec(sched(i, repoLoan.SCC_INT)) - modMoney.ToDec(sched(i, repoLoan.SCC_PRINPAID)) - modMoney.ToDec(sched(i, repoLoan.SCC_INTPAID))
End Function

Private Sub SetRowStatus(ByRef sched As Variant, ByVal i As Long, ByVal d As Date)
    If modMoney.ToDec(sched(i, repoLoan.SCC_PRINPAID)) >= modMoney.ToDec(sched(i, repoLoan.SCC_PRIN)) And _
       modMoney.ToDec(sched(i, repoLoan.SCC_INTPAID)) >= modMoney.ToDec(sched(i, repoLoan.SCC_INT)) And CStr(sched(i, repoLoan.SCC_STATUS)) <> "Future" Then
        sched(i, repoLoan.SCC_STATUS) = "Paid"
        sched(i, repoLoan.SCC_PAIDDATE) = d
    ElseIf modMoney.ToDec(sched(i, repoLoan.SCC_PRINPAID)) + modMoney.ToDec(sched(i, repoLoan.SCC_INTPAID)) > 0 And CStr(sched(i, repoLoan.SCC_STATUS)) <> "Future" Then
        sched(i, repoLoan.SCC_STATUS) = "Partial"
    End If
End Sub

' Spreads an extra principal payment over the remaining installments (same count): see svcLoanCalc.Reschedule.
Private Sub RescheduleFuture(ByRef sched As Variant, ByVal rowsOfLoan As Collection, ByVal balance As Variant, ByVal rate As Variant, ByVal method As String)
    Dim idx As Collection, i As Variant, k As Long, flatInt As Variant, s As Variant, j As Long
    Set idx = New Collection
    For Each i In rowsOfLoan
        If CStr(sched(i, repoLoan.SCC_STATUS)) = "Future" Then idx.Add i
    Next i
    k = idx.Count
    If k = 0 Then Exit Sub
    flatInt = modMoney.ToDec(sched(idx(1), repoLoan.SCC_INT))
    s = svcLoanCalc.Reschedule(balance, rate, method, k, flatInt)
    For j = 1 To k
        sched(idx(j), repoLoan.SCC_PRIN) = s(j, 1)
        sched(idx(j), repoLoan.SCC_INT) = s(j, 2)
        sched(idx(j), repoLoan.SCC_TOTAL) = s(j, 1) + s(j, 2)
    Next j
End Sub

' The extra principal is booked on the first remaining installment (due and paid both grow by it), so the schedule still adds up to the principal.
Private Sub BookExtra(ByRef sched As Variant, ByVal rowsOfLoan As Collection, ByVal extra As Variant)
    Dim i As Variant
    For Each i In rowsOfLoan
        If CStr(sched(i, repoLoan.SCC_STATUS)) = "Future" Then
            sched(i, repoLoan.SCC_PRIN) = modMoney.ToDec(sched(i, repoLoan.SCC_PRIN)) + extra
            sched(i, repoLoan.SCC_PRINPAID) = modMoney.ToDec(sched(i, repoLoan.SCC_PRINPAID)) + extra
            sched(i, repoLoan.SCC_TOTAL) = modMoney.ToDec(sched(i, repoLoan.SCC_TOTAL)) + extra
            Exit Sub
        End If
    Next i
End Sub

' Builds the repayment journal. fromAccount: H: Dr deposit GL. Cash at P: P: Dr 1000 (and 1900 to H when P is another branch).
Private Function RepayLines(ByVal h As String, ByVal p As String, ByVal a As clsAccount, ByVal cash As Boolean, ByVal total As Variant, _
                            ByVal pen As Variant, ByVal intr As Variant, ByVal prin As Variant, ByVal accrued As Variant, ByVal fee As Variant, ByVal tax As Variant) As Collection
    Dim lines As Collection
    Set lines = New Collection
    If cash Then
        svcLedger.AddDr lines, p, modConstants.GL_VAULT, total
        If p <> h Then
            svcLedger.AddCr lines, p, modConstants.GL_INTERBRANCH, total
            svcLedger.AddDr lines, h, modConstants.GL_INTERBRANCH, total
        End If
    Else
        svcLedger.AddDr lines, h, a.DepositGL, total, a.AccountNo
    End If
    svcLedger.AddCr lines, h, modConstants.GL_LOAN_PENALTY, pen
    svcLedger.AddCr lines, h, modConstants.GL_LOAN_INTEREST_REC, intr
    svcLedger.AddCr lines, h, modConstants.GL_LOANS, prin
    svcLedger.AddCr lines, h, modConstants.GL_LOAN_INCOME, accrued
    svcLedger.AddCr lines, h, modConstants.GL_FEE_INCOME, fee
    svcLedger.AddCr lines, h, modConstants.GL_TAX_FEES, tax
    Set RepayLines = lines
End Function

' Applies one payment to a loan. loans and sched are the working arrays (changed in place; the caller does ReplaceTable).
' source: Account, Cash or AutoDebit. For Account and AutoDebit `a` is the paying account; for Cash `p` is the processing branch.
Public Function ApplyPayment(ByVal uow As clsUnitOfWork, ByRef loans As Variant, ByVal row As Long, ByRef sched As Variant, ByVal amount As Variant, _
                             ByVal source As String, ByVal payType As String, ByVal a As clsAccount, ByVal p As String, ByVal grp As String, _
                             ByVal createdBy As String, ByVal approvedBy As String, ByRef pen As Variant, ByRef intr As Variant, ByRef prin As Variant) As clsResult
    Dim loanId As String, rowsOfLoan As Collection, rem2 As Variant, i As Variant, take As Variant, extra As Variant, outst As Variant, penAvail As Variant
    Dim h As String, lines As Collection, jr As clsResult, id As String, allPaid As Boolean, firstDue As Long, cash As Boolean, jType As String
    loanId = CStr(loans(row, repoLoan.LOC_ID))
    Set rowsOfLoan = ScheduleRows(sched, loanId)
    outst = modMoney.ToDec(loans(row, repoLoan.LOC_OUTSTANDING))
    rem2 = amount
    pen = CDec(0)
    intr = CDec(0)
    prin = CDec(0)
    penAvail = modMoney.RoundCcy(modMoney.ToDec(loans(row, repoLoan.LOC_PENALTY)), modConstants.CCY_MMK)
    pen = rem2
    If pen > penAvail Then pen = penAvail
    rem2 = rem2 - pen
    For Each i In rowsOfLoan
        If IsDue(CStr(sched(i, repoLoan.SCC_STATUS))) And rem2 > 0 Then
            take = modMoney.ToDec(sched(i, repoLoan.SCC_INT)) - modMoney.ToDec(sched(i, repoLoan.SCC_INTPAID))
            If take > rem2 Then take = rem2
            sched(i, repoLoan.SCC_INTPAID) = modMoney.ToDec(sched(i, repoLoan.SCC_INTPAID)) + take
            intr = intr + take
            rem2 = rem2 - take
        End If
    Next i
    For Each i In rowsOfLoan
        If IsDue(CStr(sched(i, repoLoan.SCC_STATUS))) And rem2 > 0 Then
            take = modMoney.ToDec(sched(i, repoLoan.SCC_PRIN)) - modMoney.ToDec(sched(i, repoLoan.SCC_PRINPAID))
            If take > rem2 Then take = rem2
            sched(i, repoLoan.SCC_PRINPAID) = modMoney.ToDec(sched(i, repoLoan.SCC_PRINPAID)) + take
            prin = prin + take
            rem2 = rem2 - take
        End If
    Next i
    If rem2 > 0 Then
        extra = outst - prin
        If rem2 > extra Then
            Set ApplyPayment = modMessages.FailResult("The payment is more than the loan needs. The most you can pay now is " & modMoney.FormatAmount(amount - rem2 + extra) & " MMK (use RunPayoffLoan to close the loan).")
            Exit Function
        End If
        prin = prin + rem2
        RescheduleFuture sched, rowsOfLoan, outst - prin, modMoney.ToDec(loans(row, repoLoan.LOC_RATE)), CStr(loans(row, repoLoan.LOC_METHOD))
        BookExtra sched, rowsOfLoan, rem2
    End If
    firstDue = 0
    For Each i In rowsOfLoan
        If firstDue = 0 And IsDue(CStr(sched(i, repoLoan.SCC_STATUS))) Then firstDue = i
        SetRowStatus sched, i, uow.BusinessDate
    Next i
    If firstDue > 0 And pen > 0 Then sched(firstDue, repoLoan.SCC_PENPAID) = modMoney.ToDec(sched(firstDue, repoLoan.SCC_PENPAID)) + pen
    h = CStr(loans(row, repoLoan.LOC_BRANCH))
    cash = (source = "Cash")
    Set lines = RepayLines(h, p, a, cash, amount, pen, intr, prin, CDec(0), CDec(0), CDec(0))
    Set jr = svcLedger.PostJournal(uow, modConstants.JT_LOAN_REPAY, "Loan repayment " & loanId, grp, lines)
    If Not jr.Success Then Set ApplyPayment = jr: Exit Function
    If Not cash Then
        svcAccounts.Effect uow, a, modConstants.TXN_LOAN_REPAY, modConstants.DIR_DEBIT, amount, "", p, "Loan repayment " & loanId & " (penalty " & modMoney.FormatAmount(pen) & _
            ", interest " & modMoney.FormatAmount(intr) & ", principal " & modMoney.FormatAmount(prin) & ")", loanId, jr.RefID, grp, createdBy, approvedBy, (source <> "AutoDebit")
    End If
    id = uow.NewId(modConstants.TBL_LOANPAYMENTS, "LoanPaymentID", "LP", 6)
    uow.AppendRow modConstants.TBL_LOANPAYMENTS, repoLoan.PaymentRow(id, loanId, uow.BusinessDate, amount, pen, intr, prin, source, payType, jr.RefID, grp, createdBy)
    loans(row, repoLoan.LOC_OUTSTANDING) = outst - prin
    loans(row, repoLoan.LOC_PENALTY) = modMoney.ToDec(loans(row, repoLoan.LOC_PENALTY)) - pen
    allPaid = True
    For Each i In rowsOfLoan
        If CStr(sched(i, repoLoan.SCC_STATUS)) <> "Paid" Then allPaid = False
    Next i
    If allPaid And outst - prin = 0 Then CloseLoan uow, loans, row
    Set ApplyPayment = modMessages.OkResult("", 1, jr.RefID)
End Function

' A fully paid loan becomes Closed and its collateral is released.
Public Sub CloseLoan(ByVal uow As clsUnitOfWork, ByRef loans As Variant, ByVal row As Long)
    loans(row, repoLoan.LOC_STATUS) = repoLoan.LS_CLOSED
    loans(row, repoLoan.LOC_DPD) = 0
    loans(row, repoLoan.LOC_CLASS) = 1
    loans(row, repoLoan.LOC_PROVISION) = CDec(0)
    loans(row, repoLoan.LOC_PENALTY) = CDec(0)
    svcCollateral.ReleaseForLoan uow, CStr(loans(row, repoLoan.LOC_ID))
    repoAudit.Add uow, "LoanClosed", CStr(loans(row, repoLoan.LOC_ID)), "Loan fully paid"
End Sub

' ---- Manual payment (RunLoanPayment) ----

' Control: C28 loan, C7 amount, C27 mode (Account or Cash), C5 paying account (blank = the loan account), C8 branch for cash.
Public Function Pay(ByVal user As clsUser, ByVal inp As Object) As clsResult
    Dim uow As clsUnitOfWork, msg As String, amt As Variant, loans As Variant, sched As Variant, row As Long, a As clsAccount, p As String, mode As String
    Dim pen As Variant, intr As Variant, prin As Variant, grp As String, r As clsResult, src As String, accNo As String
    msg = svcSecurity.Denied(user, svcSecurity.ACT_LOAN_PAY)
    If Len(msg) > 0 Then Set Pay = modMessages.FailResult(msg): Exit Function
    If Not modSheetUtils.TableExists(modConstants.TBL_LOANS) Then Set Pay = modMessages.FailResult(modMessages.MSG_SETUP_FIRST): Exit Function
    If Not modArgs.ParseAmount(inp(modConstants.INP_AMOUNT), amt, msg) Then Set Pay = modMessages.FailResult(msg): Exit Function
    Set uow = New clsUnitOfWork
    uow.Init repoConfig.BusinessDate(), user.UserID
    loans = uow.WorkingTable(modConstants.TBL_LOANS)
    sched = uow.WorkingTable(modConstants.TBL_LOANSCHEDULE)
    row = repoLoan.FindLoanRow(loans, CStr(inp(modConstants.INP_LOANID)))
    If row = 0 Then Set Pay = modMessages.FailResult("Unknown loan. Put the LoanID in Control C28."): Exit Function
    If CStr(loans(row, repoLoan.LOC_STATUS)) <> repoLoan.LS_ACTIVE Then Set Pay = modMessages.FailResult("Loan " & CStr(loans(row, repoLoan.LOC_ID)) & " is " & CStr(loans(row, repoLoan.LOC_STATUS)) & ". Only Active loans take repayments (use RunRecoverLoan after a write-off or repossession)."): Exit Function
    If Not svcSecurity.ResolveBranch(uow, user, CStr(inp(modConstants.INP_BRANCH)), p, msg) Then Set Pay = modMessages.FailResult(msg): Exit Function
    mode = UCase$(CStr(inp(modConstants.INP_MODE)))
    If mode = "CASH" Then
        src = "Cash"
    Else
        src = "Account"
        accNo = CStr(inp(modConstants.INP_ACCOUNT))
        If Len(accNo) = 0 Then accNo = CStr(loans(row, repoLoan.LOC_ACCOUNT))
        Set a = uow.GetAccount(accNo)
        If a Is Nothing Then Set Pay = modMessages.FailResult(modMessages.UnknownAccount(accNo)): Exit Function
        If Not svcAccounts.CanDebit(a, amt, False, msg) Then Set Pay = modMessages.FailResult(msg): Exit Function
    End If
    grp = uow.NewGroupId()
    Set r = ApplyPayment(uow, loans, row, sched, amt, src, "Installment", a, p, grp, user.UserID, "", pen, intr, prin)
    If Not r.Success Then Set Pay = r: Exit Function
    uow.ReplaceTable modConstants.TBL_LOANS, loans
    uow.ReplaceTable modConstants.TBL_LOANSCHEDULE, sched
    uow.Commit
    Set Pay = modMessages.OkResult("Payment of " & modMoney.FormatAmount(amt) & " MMK on loan " & CStr(loans(row, repoLoan.LOC_ID)) & " posted (penalty " & modMoney.FormatAmount(pen) & _
        ", interest " & modMoney.FormatAmount(intr) & ", principal " & modMoney.FormatAmount(prin) & "). Outstanding " & modMoney.FormatAmount(modMoney.ToDec(loans(row, repoLoan.LOC_OUTSTANDING))) & _
        ". Group " & grp & ".", 1, grp)
End Function

' ---- Payoff ----

' Pays everything: penalty + recognized unpaid interest + principal + interest accrued since the last due date (to 4000) + early payoff fee. Future interest is waived.
Public Function Payoff(ByVal user As clsUser, ByVal inp As Object) As clsResult
    Dim uow As clsUnitOfWork, msg As String, loans As Variant, sched As Variant, row As Long, a As clsAccount, p As String, mode As String, cash As Boolean
    Dim prow As Variant, pen As Variant, intr As Variant, prin As Variant, accrued As Variant, fee As Variant, tax As Variant, total As Variant
    Dim rowsOfLoan As Collection, i As Variant, lastDue As Date, dt As Date, days As Long, h As String, lines As Collection, jr As clsResult, grp As String, id As String
    Dim accNo As String, loanId As String, hasDue As Boolean
    msg = svcSecurity.Denied(user, svcSecurity.ACT_LOAN_PAY)
    If Len(msg) > 0 Then Set Payoff = modMessages.FailResult(msg): Exit Function
    If Not modSheetUtils.TableExists(modConstants.TBL_LOANS) Then Set Payoff = modMessages.FailResult(modMessages.MSG_SETUP_FIRST): Exit Function
    Set uow = New clsUnitOfWork
    uow.Init repoConfig.BusinessDate(), user.UserID
    loans = uow.WorkingTable(modConstants.TBL_LOANS)
    sched = uow.WorkingTable(modConstants.TBL_LOANSCHEDULE)
    row = repoLoan.FindLoanRow(loans, CStr(inp(modConstants.INP_LOANID)))
    If row = 0 Then Set Payoff = modMessages.FailResult("Unknown loan. Put the LoanID in Control C28."): Exit Function
    If CStr(loans(row, repoLoan.LOC_STATUS)) <> repoLoan.LS_ACTIVE Then Set Payoff = modMessages.FailResult("Only an Active loan can be paid off."): Exit Function
    If Not svcSecurity.ResolveBranch(uow, user, CStr(inp(modConstants.INP_BRANCH)), p, msg) Then Set Payoff = modMessages.FailResult(msg): Exit Function
    loanId = CStr(loans(row, repoLoan.LOC_ID))
    prow = repoLoan.ProductRow(CStr(loans(row, repoLoan.LOC_PRODUCT)))
    Set rowsOfLoan = ScheduleRows(sched, loanId)
    prin = modMoney.ToDec(loans(row, repoLoan.LOC_OUTSTANDING))
    pen = modMoney.RoundCcy(modMoney.ToDec(loans(row, repoLoan.LOC_PENALTY)), modConstants.CCY_MMK)
    intr = UnpaidInterest(sched, loanId)
    lastDue = CDate(loans(row, repoLoan.LOC_DISBDATE))
    For Each i In rowsOfLoan
        If modDateUtils.TryToDate(sched(i, repoLoan.SCC_DUE), dt) Then
            If dt <= uow.BusinessDate And dt > lastDue Then lastDue = dt
        End If
    Next i
    days = modDateUtils.DaysBetween(lastDue, uow.BusinessDate)
    accrued = modMoney.RoundCcy(prin * modMoney.ToDec(loans(row, repoLoan.LOC_RATE)) * days / repoConfig.GetDec(modConstants.CFG_DAYCOUNT), modConstants.CCY_MMK)
    fee = CDec(0)
    If Not IsEmpty(prow) Then fee = modMoney.RoundCcy(prin * modMoney.ToDec(prow(repoLoan.LPR_EARLYFEE)), modConstants.CCY_MMK)
    tax = svcFees.CommercialTax(fee)
    total = pen + intr + prin + accrued + fee + tax
    cash = (UCase$(CStr(inp(modConstants.INP_MODE))) = "CASH")
    If Not cash Then
        accNo = CStr(inp(modConstants.INP_ACCOUNT))
        If Len(accNo) = 0 Then accNo = CStr(loans(row, repoLoan.LOC_ACCOUNT))
        Set a = uow.GetAccount(accNo)
        If a Is Nothing Then Set Payoff = modMessages.FailResult(modMessages.UnknownAccount(accNo)): Exit Function
        If Not svcAccounts.CanDebit(a, total, True, msg) Then Set Payoff = modMessages.FailResult("Payoff of " & modMoney.FormatAmount(total) & " MMK: " & msg): Exit Function
    End If
    h = CStr(loans(row, repoLoan.LOC_BRANCH))
    grp = uow.NewGroupId()
    Set lines = RepayLines(h, p, a, cash, total, pen, intr, prin, accrued, fee, tax)
    Set jr = svcLedger.PostJournal(uow, modConstants.JT_LOAN_REPAY, "Loan payoff " & loanId, grp, lines)
    If Not jr.Success Then Set Payoff = jr: Exit Function
    If Not cash Then
        svcAccounts.Effect uow, a, modConstants.TXN_LOAN_REPAY, modConstants.DIR_DEBIT, total - fee - tax, "", p, "Loan payoff " & loanId, loanId, jr.RefID, grp, user.UserID, "", True
        svcAccounts.FeeRows uow, a, fee, tax, "EARLY_PAYOFF", "", p, jr.RefID, grp, user.UserID, ""
    End If
    For Each i In rowsOfLoan
        If CStr(sched(i, repoLoan.SCC_STATUS)) <> "Paid" Then
            If CStr(sched(i, repoLoan.SCC_STATUS)) = "Future" Then
                sched(i, repoLoan.SCC_INT) = CDec(0)
                sched(i, repoLoan.SCC_TOTAL) = sched(i, repoLoan.SCC_PRIN)
            Else
                sched(i, repoLoan.SCC_INTPAID) = sched(i, repoLoan.SCC_INT)
            End If
            sched(i, repoLoan.SCC_PRINPAID) = sched(i, repoLoan.SCC_PRIN)
            sched(i, repoLoan.SCC_STATUS) = "Paid"
            sched(i, repoLoan.SCC_PAIDDATE) = uow.BusinessDate
        End If
    Next i
    id = uow.NewId(modConstants.TBL_LOANPAYMENTS, "LoanPaymentID", "LP", 6)
    uow.AppendRow modConstants.TBL_LOANPAYMENTS, repoLoan.PaymentRow(id, loanId, uow.BusinessDate, total, pen, intr + accrued, prin, IIf(cash, "Cash", "Account"), "Payoff", jr.RefID, grp, user.UserID)
    loans(row, repoLoan.LOC_OUTSTANDING) = CDec(0)
    CloseLoan uow, loans, row
    uow.ReplaceTable modConstants.TBL_LOANS, loans
    uow.ReplaceTable modConstants.TBL_LOANSCHEDULE, sched
    uow.Commit
    Set Payoff = modMessages.OkResult("Loan " & loanId & " paid off and Closed: " & modMoney.FormatAmount(total) & " MMK (principal " & modMoney.FormatAmount(prin) & ", interest " & _
        modMoney.FormatAmount(intr + accrued) & ", penalty " & modMoney.FormatAmount(pen) & ", early fee " & modMoney.FormatAmount(fee + tax) & "). Group " & grp & ".", 1, grp)
End Function

' ---- Write-off and recovery ----

' Control: C28 loan, C15 note (needed for a Doubtful loan). HeadOffice requests, another HeadOffice approves.
Public Function WriteOff(ByVal user As clsUser, ByVal inp As Object) As clsResult
    Dim uow As clsUnitOfWork, msg As String, loans As Variant, row As Long, args As Object, r As clsResult, cls As Long
    msg = svcSecurity.Denied(user, svcSecurity.ACT_LOAN_WRITEOFF)
    If Len(msg) > 0 Then Set WriteOff = modMessages.FailResult(msg): Exit Function
    If Not modSheetUtils.TableExists(modConstants.TBL_LOANS) Then Set WriteOff = modMessages.FailResult(modMessages.MSG_SETUP_FIRST): Exit Function
    Set uow = New clsUnitOfWork
    uow.Init repoConfig.BusinessDate(), user.UserID
    loans = uow.WorkingTable(modConstants.TBL_LOANS)
    row = repoLoan.FindLoanRow(loans, CStr(inp(modConstants.INP_LOANID)))
    If row = 0 Then Set WriteOff = modMessages.FailResult("Unknown loan. Put the LoanID in Control C28."): Exit Function
    msg = WriteOffProblem(loans, row, CStr(inp(modConstants.INP_NOTE)))
    If Len(msg) > 0 Then Set WriteOff = modMessages.FailResult(msg): Exit Function
    Set args = modArgs.NewArgs()
    args("loan") = CStr(loans(row, repoLoan.LOC_ID))
    args("note") = CStr(inp(modConstants.INP_NOTE))
    Set r = svcApproval.Request(uow, modConstants.REQ_WRITEOFF, CStr(loans(row, repoLoan.LOC_ACCOUNT)), "", modMoney.ToDec(loans(row, repoLoan.LOC_OUTSTANDING)), _
        CStr(loans(row, repoLoan.LOC_BRANCH)), args, user.UserID)
    uow.Commit
    Set WriteOff = r
End Function

Private Function WriteOffProblem(ByVal loans As Variant, ByVal row As Long, ByVal note As String) As String
    Dim cls As Long
    If CStr(loans(row, repoLoan.LOC_STATUS)) <> repoLoan.LS_ACTIVE Then WriteOffProblem = "Only an Active loan can be written off.": Exit Function
    cls = CLng(modMoney.ToDec(loans(row, repoLoan.LOC_CLASS)))
    If cls < 4 Then WriteOffProblem = "Only a Doubtful (class 4) or Loss (class 5) loan can be written off. This loan is class " & cls & ".": Exit Function
    If cls = 4 And Len(note) = 0 Then WriteOffProblem = "A Doubtful loan needs an approval note in Control C15."
End Function

' B = principal outstanding + recognized unpaid interest. B: Dr 1320 (the loan's provision), Dr 5300 (B - provision), Cr 1300, Cr 1310.
Public Function ExecWriteOff(ByVal uow As clsUnitOfWork, ByVal args As Object, ByVal createdBy As String, ByVal approvedBy As String) As clsResult
    Dim loans As Variant, sched As Variant, row As Long, msg As String, prin As Variant, intr As Variant, prov As Variant, b As Variant, lines As Collection, jr As clsResult
    Dim grp As String, h As String, loanId As String
    loans = uow.WorkingTable(modConstants.TBL_LOANS)
    sched = uow.WorkingTable(modConstants.TBL_LOANSCHEDULE)
    row = repoLoan.FindLoanRow(loans, modArgs.Text(args, "loan"))
    If row = 0 Then Set ExecWriteOff = modMessages.FailResult("Unknown loan."): Exit Function
    msg = WriteOffProblem(loans, row, modArgs.Text(args, "note"))
    If Len(msg) > 0 Then Set ExecWriteOff = modMessages.FailResult(msg): Exit Function
    loanId = CStr(loans(row, repoLoan.LOC_ID))
    h = CStr(loans(row, repoLoan.LOC_BRANCH))
    prin = modMoney.ToDec(loans(row, repoLoan.LOC_OUTSTANDING))
    intr = UnpaidInterest(sched, loanId)
    prov = modMoney.ToDec(loans(row, repoLoan.LOC_PROVISION))
    b = prin + intr
    If prov > b Then prov = b
    Set lines = New Collection
    svcLedger.AddDr lines, h, modConstants.GL_LOAN_ALLOWANCE, prov
    svcLedger.AddDr lines, h, modConstants.GL_PROVISION_EXP, b - prov
    svcLedger.AddCr lines, h, modConstants.GL_LOANS, prin
    svcLedger.AddCr lines, h, modConstants.GL_LOAN_INTEREST_REC, intr
    grp = uow.NewGroupId()
    Set jr = svcLedger.PostJournal(uow, modConstants.JT_LOAN_WRITEOFF, "Loan write-off " & loanId, grp, lines)
    If Not jr.Success Then Set ExecWriteOff = jr: Exit Function
    loans(row, repoLoan.LOC_STATUS) = repoLoan.LS_WRITTENOFF
    loans(row, repoLoan.LOC_OUTSTANDING) = CDec(0)
    loans(row, repoLoan.LOC_PROVISION) = CDec(0)
    uow.ReplaceTable modConstants.TBL_LOANS, loans
    repoAudit.Add uow, "LoanWrittenOff", loanId, "Written off " & modMoney.FormatAmount(b) & ", approved by " & approvedBy
    Set ExecWriteOff = modMessages.OkResult("Loan " & loanId & " written off: " & modMoney.FormatAmount(b) & " MMK (provision used " & modMoney.FormatAmount(prov) & "). The customer still owes the money. Group " & grp & ".", 1, grp)
End Function

' Control: C28 loan (WrittenOff or Repossessed), C7 amount, C27 mode (Account or Cash), C5 account, C8 branch for cash. Cr 4030.
Public Function Recover(ByVal user As clsUser, ByVal inp As Object) As clsResult
    Dim uow As clsUnitOfWork, msg As String, amt As Variant, loans As Variant, row As Long, a As clsAccount, p As String, cash As Boolean, h As String
    Dim lines As Collection, jr As clsResult, grp As String, id As String, accNo As String, loanId As String
    msg = svcSecurity.Denied(user, svcSecurity.ACT_LOAN_RECOVER)
    If Len(msg) > 0 Then Set Recover = modMessages.FailResult(msg): Exit Function
    If Not modSheetUtils.TableExists(modConstants.TBL_LOANS) Then Set Recover = modMessages.FailResult(modMessages.MSG_SETUP_FIRST): Exit Function
    If Not modArgs.ParseAmount(inp(modConstants.INP_AMOUNT), amt, msg) Then Set Recover = modMessages.FailResult(msg): Exit Function
    Set uow = New clsUnitOfWork
    uow.Init repoConfig.BusinessDate(), user.UserID
    loans = uow.WorkingTable(modConstants.TBL_LOANS)
    row = repoLoan.FindLoanRow(loans, CStr(inp(modConstants.INP_LOANID)))
    If row = 0 Then Set Recover = modMessages.FailResult("Unknown loan. Put the LoanID in Control C28."): Exit Function
    loanId = CStr(loans(row, repoLoan.LOC_ID))
    If CStr(loans(row, repoLoan.LOC_STATUS)) <> repoLoan.LS_WRITTENOFF And CStr(loans(row, repoLoan.LOC_STATUS)) <> repoLoan.LS_REPOSSESSED Then
        Set Recover = modMessages.FailResult("Recoveries are only taken on WrittenOff or Repossessed loans."): Exit Function
    End If
    If Not svcSecurity.ResolveBranch(uow, user, CStr(inp(modConstants.INP_BRANCH)), p, msg) Then Set Recover = modMessages.FailResult(msg): Exit Function
    h = CStr(loans(row, repoLoan.LOC_BRANCH))
    cash = (UCase$(CStr(inp(modConstants.INP_MODE))) = "CASH")
    Set lines = New Collection
    If cash Then
        svcLedger.AddDr lines, p, modConstants.GL_VAULT, amt
        If p <> h Then
            svcLedger.AddCr lines, p, modConstants.GL_INTERBRANCH, amt
            svcLedger.AddDr lines, h, modConstants.GL_INTERBRANCH, amt
        End If
    Else
        accNo = CStr(inp(modConstants.INP_ACCOUNT))
        If Len(accNo) = 0 Then accNo = CStr(loans(row, repoLoan.LOC_ACCOUNT))
        Set a = uow.GetAccount(accNo)
        If a Is Nothing Then Set Recover = modMessages.FailResult(modMessages.UnknownAccount(accNo)): Exit Function
        If Not svcAccounts.CanDebit(a, amt, False, msg) Then Set Recover = modMessages.FailResult(msg): Exit Function
        svcLedger.AddDr lines, a.BranchID, a.DepositGL, amt, a.AccountNo
        If a.BranchID <> h Then
            svcLedger.AddCr lines, a.BranchID, modConstants.GL_INTERBRANCH, amt
            svcLedger.AddDr lines, h, modConstants.GL_INTERBRANCH, amt
        End If
    End If
    svcLedger.AddCr lines, h, modConstants.GL_RECOVERY, amt
    grp = uow.NewGroupId()
    Set jr = svcLedger.PostJournal(uow, modConstants.JT_LOAN_RECOVERY, "Recovery on loan " & loanId, grp, lines)
    If Not jr.Success Then Set Recover = jr: Exit Function
    If Not cash Then svcAccounts.Effect uow, a, modConstants.TXN_LOAN_REPAY, modConstants.DIR_DEBIT, amt, "", p, "Recovery on loan " & loanId, loanId, jr.RefID, grp, user.UserID, "", True
    id = uow.NewId(modConstants.TBL_LOANPAYMENTS, "LoanPaymentID", "LP", 6)
    uow.AppendRow modConstants.TBL_LOANPAYMENTS, repoLoan.PaymentRow(id, loanId, uow.BusinessDate, amt, CDec(0), CDec(0), CDec(0), IIf(cash, "Cash", "Account"), "Recovery", jr.RefID, grp, user.UserID)
    svcCollateral.OnRecovery uow, loanId, amt
    repoAudit.Add uow, "LoanRecovery", loanId, "Recovered " & modMoney.FormatAmount(amt)
    uow.Commit
    Set Recover = modMessages.OkResult("Recovery of " & modMoney.FormatAmount(amt) & " MMK on loan " & loanId & " posted to GL 4030. Group " & grp & ".", 1, grp)
End Function
