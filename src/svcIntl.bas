Attribute VB_Name = "svcIntl"
Option Explicit

' Stage 6: FAKE international payments. Only our own books move (foreign cash 1700, fees, spread). Foreign banks keep no books.

' Set to True only while the end of day refunds a failed payment (a manual reversal of an international payment is refused).
Public AllowSystemReversal As Boolean

Private Function NeedTables() As String
    If Not modSheetUtils.TableExists(modConstants.TBL_INTLPAYMENTS) Or Not modSheetUtils.TableExists(modConstants.TBL_FOREIGNBANKS) Then NeedTables = modMessages.MSG_SETUP_FIRST
End Function

Private Function FindBank(ByVal code As String, ByRef ccy As String, ByRef country As String, ByRef settleDays As Long, ByRef failRate As Variant) As Boolean
    Dim data As Variant, i As Long
    data = modSheetUtils.ReadTable(modConstants.TBL_FOREIGNBANKS)
    If IsEmpty(data) Then Exit Function
    For i = 1 To UBound(data, 1)
        If StrComp(CStr(data(i, 1)), code, vbTextCompare) = 0 Then
            country = CStr(data(i, 3))
            ccy = CStr(data(i, 4))
            settleDays = CLng(modMoney.ToDec(data(i, 5)))
            failRate = modMoney.ToDec(data(i, 6))
            FindBank = True
            Exit Function
        End If
    Next i
End Function

' True when the name contains any active name of tblBlockedNames (case-insensitive).
Public Function Screened(ByVal partyName As String) As Boolean
    Dim data As Variant, i As Long
    If Not modSheetUtils.TableExists(modConstants.TBL_BLOCKEDNAMES) Then Exit Function
    data = modSheetUtils.ReadTable(modConstants.TBL_BLOCKEDNAMES)
    If IsEmpty(data) Or Len(Trim$(partyName)) = 0 Then Exit Function
    For i = 1 To UBound(data, 1)
        If Len(CStr(data(i, 2))) > 0 And CStr(data(i, 5)) <> modConstants.FLAG_NO Then
            If InStr(1, partyName, CStr(data(i, 2)), vbTextCompare) > 0 Then Screened = True: Exit Function
        End If
    Next i
End Function

Private Function PurposeOk(ByVal p As String) As Boolean
    PurposeOk = (p = "Trade" Or p = "Family" Or p = "Education" Or p = "Services" Or p = "Investment" Or p = "Other")
End Function

Private Function MessageRef(ByVal idNo As Long, ByVal d As Date) As String
    MessageRef = "SW" & Right$("000000000000" & CStr(CDbl(idNo) * 7919# + CDbl(CLng(d))), 12)
End Function

' Today's total (MMK equivalent) of international payments sent from the accounts of one customer.
Private Function TodayTotal(ByVal uow As clsUnitOfWork, ByVal customerId As String) As Variant
    Dim data As Variant, i As Long, a As clsAccount, d As Date, total As Variant, st As String
    total = CDec(0)
    data = uow.WorkingTable(modConstants.TBL_INTLPAYMENTS)
    If Not IsEmpty(data) Then
        For i = 1 To UBound(data, 1)
            st = CStr(data(i, repoFx.IPC_STATUS))
            If CStr(data(i, repoFx.IPC_DIR)) = modConstants.DIR_OUT And (st = "Submitted" Or st = "Delivered") Then
                If modDateUtils.TryToDate(data(i, 14), d) Then
                    If d = uow.BusinessDate Then
                        Set a = uow.GetAccount(CStr(data(i, repoFx.IPC_ACCOUNT)))
                        If Not a Is Nothing Then
                            If a.CustomerID = customerId Then total = total + modMoney.ToDec(data(i, repoFx.IPC_MMK))
                        End If
                    End If
                End If
            End If
        Next i
    End If
    TodayTotal = total
End Function

' KYC rule: level 2 for any payment, level 3 at or above IntlKyc3Amount (MMK value). "" when fine.
Private Function KycProblem(ByVal customerId As String, ByVal mmk As Variant) As String
    Dim c As clsCustomer
    Set c = repoCustomer.FindCustomer(customerId)
    If c Is Nothing Then KycProblem = modMessages.UnknownCustomer(customerId): Exit Function
    If c.KYCLevel < 2 Then KycProblem = "International payments need KYC level 2 or 3 (this customer has level " & c.KYCLevel & ").": Exit Function
    If mmk >= repoConfig.GetDec(modConstants.CFG_INTL_KYC3) And c.KYCLevel < 3 Then KycProblem = "A payment of " & modMoney.FormatAmount(mmk) & " MMK or more needs KYC level 3."
End Function

' Fees of an outgoing payment in MMK: INTL_OUT on the MMK equivalent + INTL_MSG, with commercial tax.
Private Function OutFees(ByVal mmk As Variant, ByRef feeM As Variant, ByRef taxM As Variant, ByRef msg As String) As Boolean
    Dim f1 As Variant, t1 As Variant, f2 As Variant, t2 As Variant
    If Not svcFees.Compute(modConstants.FEE_INTL_OUT, mmk, f1, t1, msg) Then Exit Function
    If Not svcFees.Compute(modConstants.FEE_INTL_MSG, mmk, f2, t2, msg) Then Exit Function
    feeM = f1 + f2
    taxM = t1 + t2
    OutFees = True
End Function

' ---- Outgoing ----
' args: account, ccy (the payment currency), amount (foreign), name, bank, partyacct, purpose, branch, note.
Public Function ExecSend(ByVal uow As clsUnitOfWork, ByVal args As Object, ByVal createdBy As String, ByVal approvedBy As String) As clsResult
    Dim a As clsAccount, ccy As String, amt As Variant, mid As Variant, buy As Variant, sell As Variant, msg As String, bank As String, bCcy As String, country As String, settle As Long, fail As Variant
    Dim bA As Variant, mmkEq As Variant, feeM As Variant, taxM As Variant, feeA As Variant, taxA As Variant, total As Variant, lines As Collection, jr As clsResult, grp As String, id As String
    Dim h As String, ho As String, x As Variant, bF As Variant, bX As Variant, bTotal As Variant, screeningHit As Boolean, expected As Date, idNo As Long, p As String, fee As Variant, acctAmt As Variant
    ho = modConstants.HEAD_OFFICE_ID
    p = modArgs.Text(args, "branch")
    amt = modArgs.Dec(args, "amount")
    bank = modArgs.Text(args, "bank")
    Set a = uow.GetAccount(modArgs.Text(args, "account"))
    If a Is Nothing Then Set ExecSend = modMessages.FailResult(modMessages.UnknownAccount(modArgs.Text(args, "account"))): Exit Function
    If Not FindBank(bank, bCcy, country, settle, fail) Then Set ExecSend = modMessages.FailResult("Unknown foreign bank " & bank & " (Control C40)."): Exit Function
    If a.Ccy <> modConstants.CCY_MMK Then ccy = a.Ccy Else ccy = modArgs.Text(args, "ccy")
    If ccy = modConstants.CCY_MMK Or Len(ccy) = 0 Then Set ExecSend = modMessages.FailResult("Enter the payment currency in Control C35."): Exit Function
    If a.Ccy <> modConstants.CCY_MMK And modArgs.Text(args, "ccy") <> "" And modArgs.Text(args, "ccy") <> a.Ccy Then Set ExecSend = modMessages.FailResult("The payment currency must be the account currency."): Exit Function
    If bCcy <> ccy Then Set ExecSend = modMessages.FailResult("Foreign bank " & bank & " works in " & bCcy & ", not " & ccy & "."): Exit Function
    If Not PurposeOk(modArgs.Text(args, "purpose")) Then Set ExecSend = modMessages.FailResult("Purpose must be Trade, Family, Education, Services, Investment or Other (Control C42)."): Exit Function
    If a.Status <> modConstants.STATUS_ACTIVE Then Set ExecSend = modMessages.FailResult(modMessages.AccountStatusRefuses(a.AccountNo, a.Status)): Exit Function
    If Not svcFx.MidFor(uow, ccy, mid, buy, sell, msg) Then Set ExecSend = modMessages.FailResult(msg): Exit Function
    bA = svcFx.BaseOf(amt, mid)
    If a.Ccy = modConstants.CCY_MMK Then x = modMoney.RoundCcy(amt * sell, modConstants.CCY_MMK) Else x = bA
    mmkEq = x
    msg = KycProblem(a.CustomerID, mmkEq)
    If Len(msg) > 0 Then Set ExecSend = modMessages.FailResult(msg): Exit Function
    idNo = modIds.MaxNumber(modConstants.TBL_INTLPAYMENTS, "IntlPaymentID", "IP")
    id = uow.NewId(modConstants.TBL_INTLPAYMENTS, "IntlPaymentID", "IP", 6)
    expected = uow.BusinessDate + settle
    ' Screening first: a hit posts nothing.
    If Screened(modArgs.Text(args, "name")) Then
        uow.AppendRow modConstants.TBL_INTLPAYMENTS, repoFx.IntlRow(id, modConstants.DIR_OUT, a.AccountNo, ccy, amt, mmkEq, modArgs.Text(args, "name"), bank, modArgs.Text(args, "partyacct"), country, modArgs.Text(args, "purpose"), CDec(0), _
            MessageRef(idNo + 1, uow.BusinessDate), uow.BusinessDate, expected, "Blocked", "Name screening hit", "Hit", "")
        svcAml.OnScreeningHit uow, "International payment " & id, modArgs.Text(args, "name"), a.CustomerID, a.AccountNo, mmkEq
        Set ExecSend = modMessages.OkResult("Payment " & id & " is BLOCKED: the name matches the blocked names list. Nothing was posted.", 0, id)
        Exit Function
    End If
    If Not OutFees(mmkEq, feeM, taxM, msg) Then Set ExecSend = modMessages.FailResult(msg): Exit Function
    h = a.BranchID
    If uow.FxBalance(ho, modConstants.GL_FX_CASH, ccy) - amt < 0 Then Set ExecSend = modMessages.FailResult("Not enough foreign funds at head office. Buy currency first (RunMarketBuy)."): Exit Function
    Set lines = New Collection
    If a.Ccy <> modConstants.CCY_MMK Then
        feeA = svcFx.ToCcy(feeM, mid, ccy)
        taxA = svcFx.ToCcy(taxM, mid, ccy)
        bF = svcFx.BaseOf(feeA, mid)
        bX = svcFx.BaseOf(taxA, mid)
        total = amt + feeA + taxA
        bTotal = bA + bF + bX
        If Not svcAccounts.CanDebit(a, total, False, msg) Then Set ExecSend = modMessages.FailResult(msg): Exit Function
        svcLedger.AddDrFx lines, h, a.DepositGL, ccy, total, mid, bTotal, a.AccountNo
        svcLedger.AddCr lines, h, modConstants.GL_FEE_INCOME, bF
        svcLedger.AddCr lines, h, modConstants.GL_TAX_FEES, bX
        acctAmt = amt
        fee = feeA
    Else
        feeA = feeM
        taxA = taxM
        total = x + feeM + taxM
        If Not svcAccounts.CanDebit(a, total, False, msg) Then Set ExecSend = modMessages.FailResult(msg): Exit Function
        svcLedger.AddDr lines, h, a.DepositGL, total, a.AccountNo
        svcLedger.AddCr lines, h, modConstants.GL_FEE_INCOME, feeM
        svcLedger.AddCr lines, h, modConstants.GL_TAX_FEES, taxM
        svcLedger.AddCr lines, h, modConstants.GL_FX_INCOME, x - bA
        acctAmt = x
        fee = feeM
    End If
    If h <> ho Then
        svcLedger.AddCr lines, h, modConstants.GL_INTERBRANCH, bA
        svcLedger.AddDr lines, ho, modConstants.GL_INTERBRANCH, bA
    End If
    svcLedger.AddCrFx lines, ho, modConstants.GL_FX_CASH, ccy, amt, mid, bA
    grp = uow.NewGroupId()
    Set jr = svcLedger.PostJournal(uow, modConstants.JT_INTL_OUT, "International payment " & id & " " & modMoney.FormatCcy(amt, ccy), grp, lines)
    If Not jr.Success Then Set ExecSend = jr: Exit Function
    svcAccounts.Effect uow, a, modConstants.TXN_INTL_OUT, modConstants.DIR_DEBIT, acctAmt, bank, p, "International payment " & id & " to " & modArgs.Text(args, "name"), modArgs.Text(args, "note"), jr.RefID, grp, createdBy, approvedBy, True
    svcAccounts.FeeRows uow, a, feeA, taxA, modConstants.FEE_INTL_OUT, bank, p, jr.RefID, grp, createdBy, approvedBy
    uow.AppendRow modConstants.TBL_INTLPAYMENTS, repoFx.IntlRow(id, modConstants.DIR_OUT, a.AccountNo, ccy, amt, mmkEq, modArgs.Text(args, "name"), bank, modArgs.Text(args, "partyacct"), country, modArgs.Text(args, "purpose"), fee, _
        MessageRef(idNo + 1, uow.BusinessDate), uow.BusinessDate, expected, "Submitted", "", "Clear", jr.RefID)
    repoAudit.Add uow, "IntlSend", id, modMoney.FormatCcy(amt, ccy) & " to " & modArgs.Text(args, "name")
    Set ExecSend = modMessages.OkResult("International payment " & id & " of " & modMoney.FormatCcy(amt, ccy) & " submitted. Fees and tax " & Format$(CDbl(feeA + taxA), "#,##0.00") & IIf(a.Ccy = modConstants.CCY_MMK, " MMK", " " & ccy) & ". Account debited " & _
        modMoney.FormatCcy(total, a.Ccy) & ". Expected " & modDateUtils.ToIso(expected) & ". Group " & grp & ".", 1, grp)
End Function

' Control: C5 account, C7 foreign amount, C35 currency (MMK accounts), C39 beneficiary, C40 foreign bank, C41 beneficiary account, C42 purpose.
Public Function SendInternational(ByVal user As clsUser, ByVal inp As Object) As clsResult
    Dim uow As clsUnitOfWork, msg As String, a As clsAccount, amt As Variant, ccy As String, args As Object, p As String, mid As Variant, buy As Variant, sell As Variant, mmk As Variant, r As clsResult
    msg = svcSecurity.Denied(user, svcSecurity.ACT_INTL)
    If Len(msg) = 0 Then msg = NeedTables()
    If Len(msg) > 0 Then Set SendInternational = modMessages.FailResult(msg): Exit Function
    Set uow = New clsUnitOfWork
    uow.Init repoConfig.BusinessDate(), user.UserID
    Set a = uow.GetAccount(CStr(inp(modConstants.INP_ACCOUNT)))
    If a Is Nothing Then Set SendInternational = modMessages.FailResult(modMessages.UnknownAccount(CStr(inp(modConstants.INP_ACCOUNT)))): Exit Function
    If a.Ccy <> modConstants.CCY_MMK Then ccy = a.Ccy Else ccy = UCase$(CStr(inp(modConstants.INP_CURRENCY)))
    If ccy = modConstants.CCY_MMK Or Len(ccy) = 0 Then Set SendInternational = modMessages.FailResult("Enter the payment currency in Control C35."): Exit Function
    If Not modArgs.ParseAmountCcy(inp(modConstants.INP_AMOUNT), ccy, amt, msg) Then Set SendInternational = modMessages.FailResult(msg): Exit Function
    If Not svcSecurity.ResolveBranch(uow, user, CStr(inp(modConstants.INP_BRANCH)), p, msg) Then Set SendInternational = modMessages.FailResult(msg): Exit Function
    If Not svcSecurity.CanManageAccount(user, a) And user.Role <> modConstants.ROLE_TELLER Then Set SendInternational = modMessages.FailResult("This account belongs to branch " & a.BranchID & "."): Exit Function
    If Not svcFx.MidFor(uow, ccy, mid, buy, sell, msg) Then Set SendInternational = modMessages.FailResult(msg): Exit Function
    mmk = svcFx.BaseOf(amt, mid)
    Set args = modArgs.NewArgs()
    args("account") = a.AccountNo
    args("ccy") = ccy
    args("amount") = amt
    args("name") = CStr(inp(modConstants.INP_PARTYNAME))
    args("bank") = CStr(inp(modConstants.INP_FOREIGNBANK))
    args("partyacct") = CStr(inp(modConstants.INP_PARTYACCT))
    args("purpose") = CStr(inp(modConstants.INP_PURPOSE))
    args("branch") = p
    args("note") = CStr(inp(modConstants.INP_NOTE))
    If TodayTotal(uow, a.CustomerID) + mmk > repoConfig.GetDec(modConstants.CFG_INTL_DAILY) Then
        Set r = svcApproval.Request(uow, modConstants.REQ_INTL_PAYMENT, a.AccountNo, "", mmk, p, args, user.UserID)
    Else
        Set r = ExecSend(uow, args, user.UserID, "")
    End If
    If r.Success Then uow.Commit
    Set SendInternational = r
End Function

' ---- Incoming (HeadOffice types it) ----
' Foreign account of the same currency: HO: Dr 1700; H: Cr deposit GL. MMK account: customer gets amount x Buy, spread to 4200. Fee INTL_IN (+tax) charged from the account.
Public Function ReceiveInternational(ByVal user As clsUser, ByVal inp As Object) As clsResult
    Dim uow As clsUnitOfWork, msg As String, a As clsAccount, ccy As String, amt As Variant, mid As Variant, buy As Variant, sell As Variant, bank As String, bCcy As String, country As String
    Dim settle As Long, fail As Variant, bA As Variant, feeM As Variant, taxM As Variant, lines As Collection, jr As clsResult, grp As String, id As String, ho As String, h As String
    Dim y As Variant, feeA As Variant, taxA As Variant, bF As Variant, bX As Variant, idNo As Long, party As String
    ho = modConstants.HEAD_OFFICE_ID
    msg = svcSecurity.Denied(user, svcSecurity.ACT_INTL_IN)
    If Len(msg) = 0 Then msg = NeedTables()
    If Len(msg) > 0 Then Set ReceiveInternational = modMessages.FailResult(msg): Exit Function
    Set uow = New clsUnitOfWork
    uow.Init repoConfig.BusinessDate(), user.UserID
    Set a = uow.GetAccount(CStr(inp(modConstants.INP_ACCOUNT)))
    If a Is Nothing Then Set ReceiveInternational = modMessages.FailResult(modMessages.UnknownAccount(CStr(inp(modConstants.INP_ACCOUNT)))): Exit Function
    bank = CStr(inp(modConstants.INP_FOREIGNBANK))
    If Not FindBank(bank, bCcy, country, settle, fail) Then Set ReceiveInternational = modMessages.FailResult("Unknown foreign bank " & bank & " (Control C40)."): Exit Function
    ccy = bCcy
    If a.Ccy <> modConstants.CCY_MMK And a.Ccy <> ccy Then Set ReceiveInternational = modMessages.FailResult("The account is in " & a.Ccy & " but bank " & bank & " pays " & ccy & "."): Exit Function
    If Not modArgs.ParseAmountCcy(inp(modConstants.INP_AMOUNT), ccy, amt, msg) Then Set ReceiveInternational = modMessages.FailResult(msg): Exit Function
    If Not PurposeOk(CStr(inp(modConstants.INP_PURPOSE))) Then Set ReceiveInternational = modMessages.FailResult("Purpose must be Trade, Family, Education, Services, Investment or Other (Control C42)."): Exit Function
    If Not svcAccounts.CanCredit(a, msg) Then Set ReceiveInternational = modMessages.FailResult(msg): Exit Function
    If Not svcFx.MidFor(uow, ccy, mid, buy, sell, msg) Then Set ReceiveInternational = modMessages.FailResult(msg): Exit Function
    party = CStr(inp(modConstants.INP_PARTYNAME))
    bA = svcFx.BaseOf(amt, mid)
    idNo = modIds.MaxNumber(modConstants.TBL_INTLPAYMENTS, "IntlPaymentID", "IP")
    id = uow.NewId(modConstants.TBL_INTLPAYMENTS, "IntlPaymentID", "IP", 6)
    If Screened(party) Then
        uow.AppendRow modConstants.TBL_INTLPAYMENTS, repoFx.IntlRow(id, modConstants.DIR_IN, a.AccountNo, ccy, amt, bA, party, bank, CStr(inp(modConstants.INP_PARTYACCT)), country, CStr(inp(modConstants.INP_PURPOSE)), CDec(0), _
            MessageRef(idNo + 1, uow.BusinessDate), uow.BusinessDate, uow.BusinessDate, "Blocked", "Name screening hit", "Hit", "")
        svcAml.OnScreeningHit uow, "Incoming international payment " & id, party, a.CustomerID, a.AccountNo, bA
        uow.Commit
        Set ReceiveInternational = modMessages.OkResult("Incoming payment " & id & " is BLOCKED: the sender name matches the blocked names list. Nothing was posted.", 0, id)
        Exit Function
    End If
    If Not svcFees.Compute(modConstants.FEE_INTL_IN, bA, feeM, taxM, msg) Then Set ReceiveInternational = modMessages.FailResult(msg): Exit Function
    h = a.BranchID
    Set lines = New Collection
    svcLedger.AddDrFx lines, ho, modConstants.GL_FX_CASH, ccy, amt, mid, bA
    If h <> ho Then
        svcLedger.AddCr lines, ho, modConstants.GL_INTERBRANCH, bA
        svcLedger.AddDr lines, h, modConstants.GL_INTERBRANCH, bA
    End If
    If a.Ccy <> modConstants.CCY_MMK Then
        feeA = svcFx.ToCcy(feeM, mid, ccy)
        taxA = svcFx.ToCcy(taxM, mid, ccy)
        bF = svcFx.BaseOf(feeA, mid)
        bX = svcFx.BaseOf(taxA, mid)
        svcLedger.AddCrFx lines, h, a.DepositGL, ccy, amt, mid, bA, a.AccountNo
        svcLedger.AddDrFx lines, h, a.DepositGL, ccy, feeA + taxA, mid, bF + bX, a.AccountNo
        svcLedger.AddCr lines, h, modConstants.GL_FEE_INCOME, bF
        svcLedger.AddCr lines, h, modConstants.GL_TAX_FEES, bX
        y = amt
    Else
        feeA = feeM
        taxA = taxM
        y = modMoney.RoundCcy(amt * buy, modConstants.CCY_MMK)
        svcLedger.AddCr lines, h, a.DepositGL, y, a.AccountNo
        svcLedger.AddCr lines, h, modConstants.GL_FX_INCOME, bA - y
        svcLedger.AddDr lines, h, a.DepositGL, feeM + taxM, a.AccountNo
        svcLedger.AddCr lines, h, modConstants.GL_FEE_INCOME, feeM
        svcLedger.AddCr lines, h, modConstants.GL_TAX_FEES, taxM
    End If
    grp = uow.NewGroupId()
    Set jr = svcLedger.PostJournal(uow, modConstants.JT_INTL_IN, "Incoming international payment " & id, grp, lines)
    If Not jr.Success Then Set ReceiveInternational = jr: Exit Function
    svcAccounts.Effect uow, a, modConstants.TXN_INTL_IN, modConstants.DIR_CREDIT, y, bank, ho, "Incoming international payment " & id & " from " & party, "", jr.RefID, grp, user.UserID, "", True
    svcAccounts.FeeRows uow, a, feeA, taxA, modConstants.FEE_INTL_IN, bank, ho, jr.RefID, grp, user.UserID, ""
    uow.AppendRow modConstants.TBL_INTLPAYMENTS, repoFx.IntlRow(id, modConstants.DIR_IN, a.AccountNo, ccy, amt, bA, party, bank, CStr(inp(modConstants.INP_PARTYACCT)), country, CStr(inp(modConstants.INP_PURPOSE)), feeA, _
        MessageRef(idNo + 1, uow.BusinessDate), uow.BusinessDate, uow.BusinessDate, "Delivered", "", "Clear", jr.RefID)
    repoAudit.Add uow, "IntlReceive", id, modMoney.FormatCcy(amt, ccy) & " from " & party
    uow.Commit
    Set ReceiveInternational = modMessages.OkResult("Incoming payment " & id & ": " & modMoney.FormatCcy(amt, ccy) & " credited to " & a.AccountNo & " as " & modMoney.FormatCcy(y, a.Ccy) & ". Fee and tax " & Format$(CDbl(feeA + taxA), "#,##0.00") & ". Group " & grp & ".", 1, grp)
End Function

' ---- End of day: decisions ----

Private Function GroupOfJournal(ByVal uow As clsUnitOfWork, ByVal journalId As String) As String
    Dim h As Variant, i As Long
    h = uow.WorkingTable(modConstants.TBL_JOURNAL)
    If IsEmpty(h) Then Exit Function
    For i = UBound(h, 1) To 1 Step -1
        If CStr(h(i, repoJournal.JNC_ID)) = journalId Then GroupOfJournal = CStr(h(i, repoJournal.JNC_GROUP)): Exit Function
    Next i
End Function

' Submitted outgoing payments due on or before d: with probability FailRate they fail and are refunded in full (the journal is reversed), otherwise they are Delivered.
Public Function DecideIntlDue(ByVal uow As clsUnitOfWork, ByVal d As Date, ByRef decided As Long) As clsResult
    Dim data As Variant, i As Long, due As Date, ccy As String, country As String, settle As Long, fail As Variant, failIt As Boolean, grp As String, r As clsResult
    decided = 0
    If Not modSheetUtils.TableExists(modConstants.TBL_INTLPAYMENTS) Then Set DecideIntlDue = modMessages.OkResult("", 0): Exit Function
    data = uow.WorkingTable(modConstants.TBL_INTLPAYMENTS)
    If IsEmpty(data) Then Set DecideIntlDue = modMessages.OkResult("", 0): Exit Function
    For i = 1 To UBound(data, 1)
        If CStr(data(i, repoFx.IPC_DIR)) = modConstants.DIR_OUT And CStr(data(i, repoFx.IPC_STATUS)) = "Submitted" Then
            If modDateUtils.TryToDate(data(i, repoFx.IPC_EXPECTED), due) Then
                If due <= d Then
                    failIt = False
                    If FindBank(CStr(data(i, repoFx.IPC_BANK)), ccy, country, settle, fail) Then failIt = (Rnd < CDbl(fail))
                    If failIt Then
                        grp = GroupOfJournal(uow, CStr(data(i, repoFx.IPC_SUBMITJ)))
                        AllowSystemReversal = True
                        Set r = svcLedger.ReverseGroup(uow, grp, "SYSTEM", "")
                        AllowSystemReversal = False
                        If Not r.Success Then Set DecideIntlDue = r: Exit Function
                        data(i, repoFx.IPC_STATUS) = "Failed"
                        data(i, repoFx.IPC_REASON) = "Simulated failure, refunded in full"
                        data(i, repoFx.IPC_SETTLEJ) = r.RefID
                    Else
                        data(i, repoFx.IPC_STATUS) = "Delivered"
                    End If
                    data(i, repoFx.IPC_STATUSDATE) = d
                    decided = decided + 1
                End If
            End If
        End If
    Next i
    If decided > 0 Then uow.ReplaceTable modConstants.TBL_INTLPAYMENTS, data
    Set DecideIntlDue = modMessages.OkResult("", decided)
End Function

' Manual international reversals are refused (a failed payment is refunded by the end of day only). Called by svcLedger.ReverseGroup.
Public Function IntlReversalProblem(ByVal uow As clsUnitOfWork, ByVal groupId As String) As String
    Dim h As Variant, i As Long, t As String
    If AllowSystemReversal Then Exit Function
    h = uow.WorkingTable(modConstants.TBL_JOURNAL)
    If IsEmpty(h) Then Exit Function
    For i = 1 To UBound(h, 1)
        If CStr(h(i, repoJournal.JNC_GROUP)) = groupId Then
            t = CStr(h(i, repoJournal.JNC_TYPE))
            If Left$(t, 4) = "Intl" Then IntlReversalProblem = "An international payment cannot be reversed by hand.": Exit Function
        End If
    Next i
End Function

' ---- Returned payment (a Delivered payment comes back; fees are kept) ----

' Control: C12 = IntlPaymentID. HO: Dr 1700, Cr 1900; H: Dr 1900, Cr deposit GL (an MMK account receives amount x Buy, spread to 4200).
Public Function ReturnInternational(ByVal user As clsUser, ByVal inp As Object) As clsResult
    Dim uow As clsUnitOfWork, msg As String, data As Variant, row As Long, a As clsAccount, ccy As String, amt As Variant, mid As Variant, buy As Variant, sell As Variant, bA As Variant, y As Variant
    Dim lines As Collection, jr As clsResult, grp As String, ho As String, h As String
    ho = modConstants.HEAD_OFFICE_ID
    msg = svcSecurity.Denied(user, svcSecurity.ACT_INTL_RETURN)
    If Len(msg) = 0 Then msg = NeedTables()
    If Len(msg) > 0 Then Set ReturnInternational = modMessages.FailResult(msg): Exit Function
    Set uow = New clsUnitOfWork
    uow.Init repoConfig.BusinessDate(), user.UserID
    data = uow.WorkingTable(modConstants.TBL_INTLPAYMENTS)
    row = repoFx.FindFxById(data, CStr(inp(modConstants.INP_REFID)))
    If row = 0 Then Set ReturnInternational = modMessages.FailResult("Unknown payment. Put the IntlPaymentID in Control C12."): Exit Function
    If CStr(data(row, repoFx.IPC_DIR)) <> modConstants.DIR_OUT Or CStr(data(row, repoFx.IPC_STATUS)) <> "Delivered" Then Set ReturnInternational = modMessages.FailResult("Only a Delivered outgoing payment can be returned."): Exit Function
    Set a = uow.GetAccount(CStr(data(row, repoFx.IPC_ACCOUNT)))
    If a Is Nothing Then Set ReturnInternational = modMessages.FailResult("Unknown account."): Exit Function
    If Not svcAccounts.CanCredit(a, msg) Then Set ReturnInternational = modMessages.FailResult("The return cannot be credited. " & msg): Exit Function
    ccy = CStr(data(row, repoFx.IPC_CCY))
    amt = modMoney.ToDec(data(row, repoFx.IPC_AMOUNT))
    If Not svcFx.MidFor(uow, ccy, mid, buy, sell, msg) Then Set ReturnInternational = modMessages.FailResult(msg): Exit Function
    bA = svcFx.BaseOf(amt, mid)
    h = a.BranchID
    Set lines = New Collection
    svcLedger.AddDrFx lines, ho, modConstants.GL_FX_CASH, ccy, amt, mid, bA
    If h <> ho Then
        svcLedger.AddCr lines, ho, modConstants.GL_INTERBRANCH, bA
        svcLedger.AddDr lines, h, modConstants.GL_INTERBRANCH, bA
    End If
    If a.Ccy <> modConstants.CCY_MMK Then
        svcLedger.AddCrFx lines, h, a.DepositGL, ccy, amt, mid, bA, a.AccountNo
        y = amt
    Else
        y = modMoney.RoundCcy(amt * buy, modConstants.CCY_MMK)
        svcLedger.AddCr lines, h, a.DepositGL, y, a.AccountNo
        svcLedger.AddCr lines, h, modConstants.GL_FX_INCOME, bA - y
    End If
    grp = uow.NewGroupId()
    Set jr = svcLedger.PostJournal(uow, modConstants.JT_INTL_RETURN, "International payment returned " & CStr(data(row, 1)), grp, lines)
    If Not jr.Success Then Set ReturnInternational = jr: Exit Function
    svcAccounts.Effect uow, a, modConstants.TXN_INTL_RETURN, modConstants.DIR_CREDIT, y, "", ho, "International payment " & CStr(data(row, 1)) & " returned (fees are kept)", "", jr.RefID, grp, user.UserID, "", False
    data(row, repoFx.IPC_STATUS) = "Returned"
    data(row, repoFx.IPC_STATUSDATE) = uow.BusinessDate
    data(row, repoFx.IPC_REASON) = "Returned (forced)"
    data(row, repoFx.IPC_SETTLEJ) = jr.RefID
    uow.ReplaceTable modConstants.TBL_INTLPAYMENTS, data
    repoAudit.Add uow, "IntlReturn", CStr(data(row, 1)), "Returned"
    uow.Commit
    Set ReturnInternational = modMessages.OkResult("Payment " & CStr(data(row, 1)) & " returned: " & modMoney.FormatCcy(y, a.Ccy) & " credited to " & a.AccountNo & ". Group " & grp & ".", 1, grp)
End Function
