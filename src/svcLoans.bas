Attribute VB_Name = "svcLoans"
Option Explicit

' Stage 3: loan quote, application, approval, funding checks and disbursement.

Private Const SH_QUOTE As String = "LoanQuote"

Private Function NeedTables() As String
    If Not modSheetUtils.TableExists(modConstants.TBL_LOANS) Or Not modSheetUtils.TableExists(modConstants.TBL_LOANPRODUCTS) Then
        NeedTables = "The loan tables are missing. Run RunSetup first."
    End If
End Function

' ---- Shared checks ----

Private Function ProductProblem(ByVal prow As Variant, ByVal amount As Variant, ByVal term As Long, ByVal code As String) As String
    If IsEmpty(prow) Then ProductProblem = "Unknown loan product " & code & " (Control C26).": Exit Function
    If CStr(prow(repoLoan.LPR_ACTIVE)) <> modConstants.FLAG_YES Then ProductProblem = "Loan product " & code & " is not active.": Exit Function
    If amount < modMoney.ToDec(prow(repoLoan.LPR_MIN)) Or amount > modMoney.ToDec(prow(repoLoan.LPR_MAX)) Then
        ProductProblem = "The amount must be between " & modMoney.FormatAmount(modMoney.ToDec(prow(repoLoan.LPR_MIN))) & " and " & modMoney.FormatAmount(modMoney.ToDec(prow(repoLoan.LPR_MAX))) & " for " & code & "."
        Exit Function
    End If
    If term < CLng(modMoney.ToDec(prow(repoLoan.LPR_MINTERM))) Or term > CLng(modMoney.ToDec(prow(repoLoan.LPR_MAXTERM))) Then
        ProductProblem = "The term must be between " & CLng(modMoney.ToDec(prow(repoLoan.LPR_MINTERM))) & " and " & CLng(modMoney.ToDec(prow(repoLoan.LPR_MAXTERM))) & " months for " & code & " (Control C10)."
    End If
End Function

' Equity = credit balance of 3000 + 3100.
Private Function Capital(ByVal uow As clsUnitOfWork) As Variant
    Capital = -(uow.GLNetAll(modConstants.GL_CAPITAL) + uow.GLNetAll(modConstants.GL_RETAINED))
End Function

' Total of the customer's loans: outstanding of Active loans plus principal of Applied and Approved ones (the loan `skipId` is left out).
Private Function Exposure(ByVal uow As clsUnitOfWork, ByVal customerId As String, ByVal skipId As String) As Variant
    Dim data As Variant, i As Long, total As Variant, st As String
    total = CDec(0)
    data = uow.WorkingTable(modConstants.TBL_LOANS)
    If Not IsEmpty(data) Then
        For i = 1 To UBound(data, 1)
            If CStr(data(i, repoLoan.LOC_CUSTOMER)) = customerId And CStr(data(i, repoLoan.LOC_ID)) <> skipId Then
                st = CStr(data(i, repoLoan.LOC_STATUS))
                If st = repoLoan.LS_ACTIVE Then total = total + modMoney.ToDec(data(i, repoLoan.LOC_OUTSTANDING))
                If st = repoLoan.LS_APPLIED Or st = repoLoan.LS_APPROVED Then total = total + modMoney.ToDec(data(i, repoLoan.LOC_PRINCIPAL))
            End If
        Next i
    End If
    Exposure = total
End Function

Private Function SingleBorrowerProblem(ByVal uow As clsUnitOfWork, ByVal customerId As String, ByVal amount As Variant, ByVal skipId As String) As String
    Dim lim As Variant, ex As Variant
    lim = repoConfig.GetDec(modConstants.CFG_SINGLE_BORROWER) * Capital(uow)
    ex = Exposure(uow, customerId, skipId)
    If ex + amount > lim Then
        SingleBorrowerProblem = "Single borrower limit: existing loans " & modMoney.FormatAmount(ex) & " + new " & modMoney.FormatAmount(amount) & " is above the limit " & modMoney.FormatAmount(lim) & "."
    End If
End Function

Private Function OverdueProblem(ByVal uow As clsUnitOfWork, ByVal customerId As String) As String
    Dim data As Variant, i As Long
    data = uow.WorkingTable(modConstants.TBL_LOANS)
    If IsEmpty(data) Then Exit Function
    For i = 1 To UBound(data, 1)
        If CStr(data(i, repoLoan.LOC_CUSTOMER)) = customerId And CStr(data(i, repoLoan.LOC_STATUS)) = repoLoan.LS_ACTIVE Then
            If modMoney.ToDec(data(i, repoLoan.LOC_DPD)) >= 30 Then
                OverdueProblem = "The customer has loan " & CStr(data(i, repoLoan.LOC_ID)) & " that is " & CLng(modMoney.ToDec(data(i, repoLoan.LOC_DPD))) & " days past due."
                Exit Function
            End If
        End If
    Next i
End Function

' Customer, KYC, product and account rules that apply at application and again at approval.
Private Function EligibilityProblem(ByVal uow As clsUnitOfWork, ByVal customerId As String, ByVal prow As Variant, ByVal productCode As String, _
                                    ByVal amount As Variant, ByVal term As Long, ByVal accountNo As String, ByVal skipId As String) As String
    Dim cust As clsCustomer, a As clsAccount, msg As String
    Set cust = repoCustomer.FindCustomer(customerId)
    If cust Is Nothing Then EligibilityProblem = modMessages.UnknownCustomer(customerId): Exit Function
    If cust.Status <> modConstants.STATUS_ACTIVE Then EligibilityProblem = "Customer " & customerId & " is not Active.": Exit Function
    If cust.KYCLevel < repoConfig.GetLong(modConstants.CFG_LOAN_MIN_KYC) Then
        EligibilityProblem = "KYC level " & cust.KYCLevel & " is below the minimum " & repoConfig.GetLong(modConstants.CFG_LOAN_MIN_KYC) & " for a loan."
        Exit Function
    End If
    If amount >= repoConfig.GetDec(modConstants.CFG_LOAN_KYC3) And cust.KYCLevel < 3 Then EligibilityProblem = "A loan of this size needs KYC level 3.": Exit Function
    msg = ProductProblem(prow, amount, term, productCode)
    If Len(msg) > 0 Then EligibilityProblem = msg: Exit Function
    Set a = uow.GetAccount(accountNo)
    If a Is Nothing Then EligibilityProblem = modMessages.UnknownAccount(accountNo): Exit Function
    If a.CustomerID <> customerId Then EligibilityProblem = "The account must belong to the customer (Control C5).": Exit Function
    If a.Status <> modConstants.STATUS_ACTIVE Or a.Ccy <> modConstants.CCY_MMK Or a.ProductType = modConstants.PT_FIXED Then
        EligibilityProblem = "The loan account must be an Active MMK savings or current account."
        Exit Function
    End If
    msg = OverdueProblem(uow, customerId)
    If Len(msg) > 0 Then EligibilityProblem = msg: Exit Function
    EligibilityProblem = SingleBorrowerProblem(uow, customerId, amount, skipId)
End Function

' ---- Quote ----

' Control: C26 product, C7 amount, C10 term. Writes the LoanQuote sheet; changes no data.
Public Function Quote(ByVal user As clsUser, ByVal inp As Object) As clsResult
    Dim msg As String, prow As Variant, amt As Variant, term As Long, sched As Variant, i As Long, rowsA As Collection, bal As Variant
    Dim tv As Variant, totalInt As Variant, d As Date
    msg = svcSecurity.Denied(user, svcSecurity.ACT_LOAN_QUOTE)
    If Len(msg) > 0 Then Set Quote = modMessages.FailResult(msg): Exit Function
    msg = NeedTables()
    If Len(msg) > 0 Then Set Quote = modMessages.FailResult(msg): Exit Function
    If Not modArgs.ParseAmount(inp(modConstants.INP_AMOUNT), amt, msg) Then Set Quote = modMessages.FailResult(msg): Exit Function
    If Not modMoney.TryToDec(inp(modConstants.INP_TERM), tv) Then Set Quote = modMessages.FailResult("Enter the term in months in Control C10."): Exit Function
    term = CLng(tv)
    prow = repoLoan.ProductRow(CStr(inp(modConstants.INP_LOANPRODUCT)))
    msg = ProductProblem(prow, amt, term, CStr(inp(modConstants.INP_LOANPRODUCT)))
    If Len(msg) > 0 Then Set Quote = modMessages.FailResult(msg): Exit Function
    repoConfig.Reload
    d = repoConfig.BusinessDate()
    sched = svcLoanCalc.Build(amt, modMoney.ToDec(prow(repoLoan.LPR_RATE)), CStr(prow(repoLoan.LPR_METHOD)), term, d)
    Set rowsA = New Collection
    bal = amt
    totalInt = CDec(0)
    For i = 1 To UBound(sched, 1)
        bal = bal - sched(i, 2)
        totalInt = totalInt + sched(i, 3)
        rowsA.Add Array(i, sched(i, 1), sched(i, 2), sched(i, 3), sched(i, 2) + sched(i, 3), bal)
    Next i
    modSheetUtils.EnsureSheet SH_QUOTE
    modSheetUtils.PrepareReport SH_QUOTE, ""
    modSheetUtils.SetCell SH_QUOTE, 1, 1, "Loan quote (SIMULATION): " & CStr(prow(repoLoan.LPR_NAME)) & ", " & modMoney.FormatAmount(amt) & " MMK, " & term & " months, rate " & _
        Format$(CDbl(modMoney.ToDec(prow(repoLoan.LPR_RATE))) * 100, "0.00") & "% " & CStr(prow(repoLoan.LPR_METHOD)), True
    modSheetUtils.SetCell SH_QUOTE, 2, 1, "Total interest " & modMoney.FormatAmount(totalInt) & " MMK. Processing fee " & modMoney.FormatAmount(modMoney.RoundCcy(amt * modMoney.ToDec(prow(repoLoan.LPR_FEEPCT)), modConstants.CCY_MMK)) & " MMK."
    modSheetUtils.WriteGrid SH_QUOTE, 4, 1, svcReports.ToGrid(Array("No", "Due date", "Principal", "Interest", "Total", "Balance after"), rowsA)
    modSheetUtils.StyleReport SH_QUOTE, 4, 6, "3,4,5,6"
    Set Quote = modMessages.OkResult("Quote written to the LoanQuote sheet: installment 1 = " & modMoney.FormatAmount(sched(1, 2) + sched(1, 3)) & " MMK (principal " & _
        modMoney.FormatAmount(sched(1, 2)) & " + interest " & modMoney.FormatAmount(sched(1, 3)) & "), total interest " & modMoney.FormatAmount(totalInt) & " MMK.", UBound(sched, 1))
End Function

' ---- Application and approval ----

' Control: C13 customer, C26 product, C7 amount, C10 term, C5 loan account, C29 collateral (when the product needs it), C15 note.
Public Function Apply(ByVal user As clsUser, ByVal inp As Object) As clsResult
    Dim uow As clsUnitOfWork, msg As String, amt As Variant, tv As Variant, term As Long, prow As Variant, cust As clsCustomer, id As String
    Dim args As Object, r As clsResult, code As String, cid As String, collat As String
    msg = svcSecurity.Denied(user, svcSecurity.ACT_LOAN_APPLY)
    If Len(msg) > 0 Then Set Apply = modMessages.FailResult(msg): Exit Function
    msg = NeedTables()
    If Len(msg) > 0 Then Set Apply = modMessages.FailResult(msg): Exit Function
    If Not modArgs.ParseAmount(inp(modConstants.INP_AMOUNT), amt, msg) Then Set Apply = modMessages.FailResult(msg): Exit Function
    If Not modMoney.TryToDec(inp(modConstants.INP_TERM), tv) Then Set Apply = modMessages.FailResult("Enter the term in months in Control C10."): Exit Function
    term = CLng(tv)
    code = CStr(inp(modConstants.INP_LOANPRODUCT))
    cid = CStr(inp(modConstants.INP_CUSTOMER))
    collat = CStr(inp(modConstants.INP_COLLATERAL))
    prow = repoLoan.ProductRow(code)
    Set uow = New clsUnitOfWork
    uow.Init repoConfig.BusinessDate(), user.UserID
    msg = EligibilityProblem(uow, cid, prow, code, amt, term, CStr(inp(modConstants.INP_ACCOUNT)), "")
    If Len(msg) > 0 Then Set Apply = modMessages.FailResult(msg): Exit Function
    Set cust = repoCustomer.FindCustomer(cid)
    If user.Role <> modConstants.ROLE_HEADOFFICE And cust.HomeBranchID <> user.BranchID Then
        Set Apply = modMessages.FailResult("This customer belongs to branch " & cust.HomeBranchID & ". Apply at that branch."): Exit Function
    End If
    If CStr(prow(repoLoan.LPR_COLLATERAL)) = modConstants.FLAG_YES Then
        If repoConfig.GetBool(modConstants.CFG_REQUIRE_COLLATERAL) Then
            msg = svcCollateral.ApplyProblem(uow, cid, modMoney.ToDec(prow(repoLoan.LPR_MAXLTV)), collat, amt)
            If Len(msg) > 0 Then Set Apply = modMessages.FailResult(msg): Exit Function
        Else
            clsLogger.Info "svcLoans.Apply", "Collateral check skipped (RequireCollateralCheck = No) for a " & code & " application."
        End If
    End If
    id = uow.NewId(modConstants.TBL_LOANS, "LoanID", "LO", 6)
    uow.AppendRow modConstants.TBL_LOANS, repoLoan.LoanRow(id, cid, cust.HomeBranchID, code, CStr(inp(modConstants.INP_ACCOUNT)), amt, modMoney.ToDec(prow(repoLoan.LPR_RATE)), _
        CStr(prow(repoLoan.LPR_METHOD)), term, uow.BusinessDate, user.UserID, CStr(inp(modConstants.INP_NOTE)))
    Set args = modArgs.NewArgs()
    args("loan") = id
    Set r = svcApproval.Request(uow, modConstants.REQ_LOAN_APPROVAL, CStr(inp(modConstants.INP_ACCOUNT)), "", amt, cust.HomeBranchID, args, user.UserID)
    uow.Commit
    Set Apply = modMessages.OkResult("Loan " & id & " applied (" & code & ", " & modMoney.FormatAmount(amt) & " MMK, " & term & " months). " & r.Message, 1, id)
End Function

' Run when the loan approval is approved: all eligibility rules again, then lock the rate.
Public Function ExecApprove(ByVal uow As clsUnitOfWork, ByVal args As Object, ByVal createdBy As String, ByVal approvedBy As String) As clsResult
    Dim data As Variant, row As Long, prow As Variant, msg As String
    data = uow.WorkingTable(modConstants.TBL_LOANS)
    row = repoLoan.FindLoanRow(data, modArgs.Text(args, "loan"))
    If row = 0 Then Set ExecApprove = modMessages.FailResult("Unknown loan " & modArgs.Text(args, "loan") & "."): Exit Function
    If CStr(data(row, repoLoan.LOC_STATUS)) <> repoLoan.LS_APPLIED Then Set ExecApprove = modMessages.FailResult("Loan " & CStr(data(row, repoLoan.LOC_ID)) & " is " & CStr(data(row, repoLoan.LOC_STATUS)) & ", not Applied."): Exit Function
    prow = repoLoan.ProductRow(CStr(data(row, repoLoan.LOC_PRODUCT)))
    msg = EligibilityProblem(uow, CStr(data(row, repoLoan.LOC_CUSTOMER)), prow, CStr(data(row, repoLoan.LOC_PRODUCT)), modMoney.ToDec(data(row, repoLoan.LOC_PRINCIPAL)), _
        CLng(modMoney.ToDec(data(row, repoLoan.LOC_TERM))), CStr(data(row, repoLoan.LOC_ACCOUNT)), CStr(data(row, repoLoan.LOC_ID)))
    If Len(msg) > 0 Then Set ExecApprove = modMessages.FailResult(msg): Exit Function
    data(row, repoLoan.LOC_STATUS) = repoLoan.LS_APPROVED
    data(row, repoLoan.LOC_RATE) = modMoney.ToDec(prow(repoLoan.LPR_RATE))
    data(row, repoLoan.LOC_METHOD) = CStr(prow(repoLoan.LPR_METHOD))
    data(row, repoLoan.LOC_APPRDATE) = uow.BusinessDate
    data(row, repoLoan.LOC_APPROVEDBY) = approvedBy
    uow.ReplaceTable modConstants.TBL_LOANS, data
    repoAudit.Add uow, "LoanApproved", CStr(data(row, repoLoan.LOC_ID)), "Approved by " & approvedBy
    Set ExecApprove = modMessages.OkResult("Loan " & CStr(data(row, repoLoan.LOC_ID)) & " is Approved. Rate locked at " & Format$(CDbl(modMoney.ToDec(prow(repoLoan.LPR_RATE))) * 100, "0.00") & "%.", 1, "")
End Function

' When a loan approval is rejected the loan itself becomes Rejected.
Public Sub MarkRejected(ByVal uow As clsUnitOfWork, ByVal loanId As String, ByVal reason As String)
    Dim data As Variant, row As Long
    data = uow.WorkingTable(modConstants.TBL_LOANS)
    row = repoLoan.FindLoanRow(data, loanId)
    If row = 0 Then Exit Sub
    If CStr(data(row, repoLoan.LOC_STATUS)) <> repoLoan.LS_APPLIED Then Exit Sub
    data(row, repoLoan.LOC_STATUS) = repoLoan.LS_REJECTED
    If Len(reason) > 0 Then data(row, repoLoan.LOC_NOTE) = reason
    uow.ReplaceTable modConstants.TBL_LOANS, data
End Sub

' ---- Funding and disbursement ----

' The bank lends only what it has: reserve and loan-to-deposit limits, with the numbers in the message.
Public Function FundingProblem(ByVal uow As clsUnitOfWork, ByVal principal As Variant) As String
    Dim deposits As Variant, liquid As Variant, reserve As Variant, maxL2D As Variant, loans As Variant
    deposits = -(uow.GLNetAll(modConstants.GL_SAVINGS) + uow.GLNetAll(modConstants.GL_CURRENT) + uow.GLNetAll(modConstants.GL_FIXED))
    liquid = uow.GLNetAll(modConstants.GL_TREASURY) + uow.GLNetAll(modConstants.GL_VAULT)
    reserve = repoConfig.GetDec("ReserveRatio") * (deposits + principal)
    If liquid < reserve Then
        FundingProblem = "Liquidity: liquid funds " & modMoney.FormatAmount(liquid) & " would be below the reserve needed " & modMoney.FormatAmount(reserve) & " (see LoanLimits)."
        Exit Function
    End If
    maxL2D = repoConfig.GetDec(modConstants.CFG_MAX_LOAN_TO_DEPOSIT) * (deposits + principal)
    loans = uow.GLNetAll(modConstants.GL_LOANS) + principal
    If loans > maxL2D Then
        FundingProblem = "Loan-to-deposit: loans " & modMoney.FormatAmount(loans) & " would be above " & modMoney.FormatAmount(maxL2D) & " (see LoanLimits)."
    End If
End Function

' Control: C28 = the Approved loan. H: Dr 1300 principal, Cr deposit GL; fee + tax debited from the account (Cr 4100, Cr 2210).
Public Function Disburse(ByVal user As clsUser, ByVal inp As Object) As clsResult
    Dim uow As clsUnitOfWork, msg As String, data As Variant, row As Long, a As clsAccount, prow As Variant, p As Variant, fee As Variant, tax As Variant
    Dim lines As Collection, jr As clsResult, grp As String, sched As Variant, i As Long, n As Long, id As String, h As String, rate As Variant, term As Long
    msg = svcSecurity.Denied(user, svcSecurity.ACT_LOAN_DISBURSE)
    If Len(msg) > 0 Then Set Disburse = modMessages.FailResult(msg): Exit Function
    msg = NeedTables()
    If Len(msg) > 0 Then Set Disburse = modMessages.FailResult(msg): Exit Function
    Set uow = New clsUnitOfWork
    uow.Init repoConfig.BusinessDate(), user.UserID
    data = uow.WorkingTable(modConstants.TBL_LOANS)
    row = repoLoan.FindLoanRow(data, CStr(inp(modConstants.INP_LOANID)))
    If row = 0 Then Set Disburse = modMessages.FailResult("Unknown loan. Put the LoanID in Control C28."): Exit Function
    If CStr(data(row, repoLoan.LOC_STATUS)) <> repoLoan.LS_APPROVED Then Set Disburse = modMessages.FailResult("Loan " & CStr(data(row, repoLoan.LOC_ID)) & " is " & CStr(data(row, repoLoan.LOC_STATUS)) & ", not Approved."): Exit Function
    If user.Role <> modConstants.ROLE_HEADOFFICE And CStr(data(row, repoLoan.LOC_BRANCH)) <> user.BranchID Then
        Set Disburse = modMessages.FailResult("This loan belongs to branch " & CStr(data(row, repoLoan.LOC_BRANCH)) & "."): Exit Function
    End If
    Set a = uow.GetAccount(CStr(data(row, repoLoan.LOC_ACCOUNT)))
    If a Is Nothing Then Set Disburse = modMessages.FailResult(modMessages.UnknownAccount(CStr(data(row, repoLoan.LOC_ACCOUNT)))): Exit Function
    If a.Status <> modConstants.STATUS_ACTIVE Then Set Disburse = modMessages.FailResult(modMessages.AccountStatusRefuses(a.AccountNo, a.Status)): Exit Function
    prow = repoLoan.ProductRow(CStr(data(row, repoLoan.LOC_PRODUCT)))
    p = modMoney.ToDec(data(row, repoLoan.LOC_PRINCIPAL))
    msg = SingleBorrowerProblem(uow, CStr(data(row, repoLoan.LOC_CUSTOMER)), p, CStr(data(row, repoLoan.LOC_ID)))
    If Len(msg) = 0 Then msg = FundingProblem(uow, p)
    If Len(msg) = 0 And CStr(prow(repoLoan.LPR_COLLATERAL)) = modConstants.FLAG_YES And repoConfig.GetBool(modConstants.CFG_REQUIRE_COLLATERAL) Then
        msg = svcCollateral.DisbursementProblem(uow, CStr(data(row, repoLoan.LOC_ID)), modMoney.ToDec(prow(repoLoan.LPR_MAXLTV)), p)
    End If
    If Len(msg) > 0 Then Set Disburse = modMessages.FailResult("Disbursement refused. " & msg): Exit Function
    fee = modMoney.RoundCcy(p * modMoney.ToDec(prow(repoLoan.LPR_FEEPCT)), modConstants.CCY_MMK)
    tax = svcFees.CommercialTax(fee)
    h = a.BranchID
    Set lines = New Collection
    svcLedger.AddDr lines, h, modConstants.GL_LOANS, p
    svcLedger.AddCr lines, h, a.DepositGL, p, a.AccountNo
    svcLedger.AddDr lines, h, a.DepositGL, fee + tax, a.AccountNo
    svcLedger.AddCr lines, h, modConstants.GL_FEE_INCOME, fee
    svcLedger.AddCr lines, h, modConstants.GL_TAX_FEES, tax
    grp = uow.NewGroupId()
    Set jr = svcLedger.PostJournal(uow, modConstants.JT_LOAN_DISB, "Loan disbursement " & CStr(data(row, repoLoan.LOC_ID)), grp, lines)
    If Not jr.Success Then Set Disburse = jr: Exit Function
    svcAccounts.Effect uow, a, modConstants.TXN_LOAN_DISB, modConstants.DIR_CREDIT, p, "", h, "Loan " & CStr(data(row, repoLoan.LOC_ID)) & " disbursed", CStr(data(row, repoLoan.LOC_ID)), jr.RefID, grp, user.UserID, "", True
    If fee > 0 Then svcAccounts.Effect uow, a, modConstants.TXN_LOAN_FEE, modConstants.DIR_DEBIT, fee, "", h, "Loan processing fee", CStr(data(row, repoLoan.LOC_ID)), jr.RefID, grp, user.UserID, "", False
    If tax > 0 Then svcAccounts.Effect uow, a, modConstants.TXN_TAX, modConstants.DIR_DEBIT, tax, "", h, "Tax on loan fee", CStr(data(row, repoLoan.LOC_ID)), jr.RefID, grp, user.UserID, "", False
    rate = modMoney.ToDec(data(row, repoLoan.LOC_RATE))
    term = CLng(modMoney.ToDec(data(row, repoLoan.LOC_TERM)))
    sched = svcLoanCalc.Build(p, rate, CStr(data(row, repoLoan.LOC_METHOD)), term, uow.BusinessDate)
    n = UBound(sched, 1)
    For i = 1 To n
        id = uow.NewId(modConstants.TBL_LOANSCHEDULE, "ScheduleID", "SC", 7)
        uow.AppendRow modConstants.TBL_LOANSCHEDULE, repoLoan.ScheduleRow(id, CStr(data(row, repoLoan.LOC_ID)), i, CDate(sched(i, 1)), sched(i, 2), sched(i, 3))
    Next i
    data(row, repoLoan.LOC_DISBDATE) = uow.BusinessDate
    data(row, repoLoan.LOC_FIRSTDUE) = sched(1, 1)
    data(row, repoLoan.LOC_MATURITY) = sched(n, 1)
    data(row, repoLoan.LOC_STATUS) = repoLoan.LS_ACTIVE
    data(row, repoLoan.LOC_OUTSTANDING) = p
    data(row, repoLoan.LOC_PENALTY) = CDec(0)
    data(row, repoLoan.LOC_DPD) = 0
    data(row, repoLoan.LOC_CLASS) = 1
    data(row, repoLoan.LOC_PROVISION) = CDec(0)
    uow.ReplaceTable modConstants.TBL_LOANS, data
    repoAudit.Add uow, "LoanDisbursed", CStr(data(row, repoLoan.LOC_ID)), modMoney.FormatAmount(p) & " to " & a.AccountNo
    uow.Commit
    Set Disburse = modMessages.OkResult("Loan " & CStr(data(row, repoLoan.LOC_ID)) & " disbursed: " & modMoney.FormatAmount(p) & " MMK credited to " & a.AccountNo & ", fee " & _
        modMoney.FormatAmount(fee) & " + tax " & modMoney.FormatAmount(tax) & " debited. " & n & " installments scheduled, first due " & modDateUtils.ToIso(CDate(sched(1, 1))) & ". Group " & grp & ".", 1, grp)
End Function
