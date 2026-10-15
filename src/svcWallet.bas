Attribute VB_Name = "svcWallet"
Option Explicit

' Stage 6: e-wallets (GL 2030), merchants (QR payments) and agents. Every change of the total of all wallet balances is matched in the same
' journal by head office moving the same amount in the trust account (Dr 1110, Cr 1100 for an increase; Dr 1100, Cr 1110 for a decrease).

Private Function NeedTables() As String
    If Not modSheetUtils.TableExists(modConstants.TBL_MERCHANTS) Or Not modSheetUtils.TableExists(modConstants.TBL_AGENTS) Or Not modSheetUtils.TableExists(modConstants.TBL_WALLETLIMITS) Then NeedTables = modMessages.MSG_SETUP_FIRST
End Function

' ---- Helpers ----

Private Function WalletAccount(ByVal uow As clsUnitOfWork, ByVal accountNo As String, ByRef msg As String) As clsAccount
    Dim a As clsAccount
    Set a = uow.GetAccount(accountNo)
    If a Is Nothing Then msg = modMessages.UnknownAccount(accountNo): Exit Function
    If a.ProductType <> modConstants.PT_WALLET Then msg = "Account " & accountNo & " is not a wallet.": Exit Function
    If a.Status <> modConstants.STATUS_ACTIVE Then msg = modMessages.AccountStatusRefuses(accountNo, a.Status): Exit Function
    Set WalletAccount = a
End Function

Private Function BankAccount(ByVal uow As clsUnitOfWork, ByVal accountNo As String, ByRef msg As String) As clsAccount
    Dim a As clsAccount
    Set a = uow.GetAccount(accountNo)
    If a Is Nothing Then msg = modMessages.UnknownAccount(accountNo): Exit Function
    If a.ProductType = modConstants.PT_WALLET Or a.ProductType = modConstants.PT_FIXED Or a.Ccy <> modConstants.CCY_MMK Then msg = "The bank account must be an MMK savings or current account.": Exit Function
    Set BankAccount = a
End Function

Private Function IsAgentWallet(ByVal uow As clsUnitOfWork, ByVal accountNo As String) As Boolean
    Dim data As Variant, i As Long
    data = uow.WorkingTable(modConstants.TBL_AGENTS)
    If IsEmpty(data) Then Exit Function
    For i = 1 To UBound(data, 1)
        If CStr(data(i, repoFx.AGC_WALLET)) = accountNo Then IsAgentWallet = True: Exit Function
    Next i
End Function

Private Function LimitsFor(ByVal kyc As Long, ByRef maxBal As Variant, ByRef daily As Variant, ByRef monthly As Variant, ByRef singleMax As Variant) As Boolean
    Dim data As Variant, i As Long
    data = modSheetUtils.ReadTable(modConstants.TBL_WALLETLIMITS)
    If IsEmpty(data) Then Exit Function
    For i = 1 To UBound(data, 1)
        If CLng(modMoney.ToDec(data(i, 1))) = kyc Then
            maxBal = modMoney.ToDec(data(i, 2))
            daily = modMoney.ToDec(data(i, 3))
            monthly = modMoney.ToDec(data(i, 4))
            singleMax = modMoney.ToDec(data(i, 5))
            LimitsFor = True
            Exit Function
        End If
    Next i
End Function

' Today's spending (payments, transfers out, cash-out) and this month's volume (everything but fee and tax rows) of one wallet.
Private Sub Usage(ByVal w As clsAccount, ByVal d As Date, ByRef spent As Variant, ByRef volume As Variant)
    Dim data As Variant, i As Long, dt As Date, t As String, amt As Variant
    spent = CDec(0)
    volume = CDec(0)
    data = repoTransaction.LoadTransactions()
    If IsEmpty(data) Then Exit Sub
    For i = 1 To UBound(data, 1)
        If CStr(data(i, 4)) = w.AccountNo And CStr(data(i, 14)) = modConstants.STATUS_POSTED Then
            t = CStr(data(i, 5))
            If t <> modConstants.TXN_FEE And t <> modConstants.TXN_TAX And modDateUtils.TryToDate(data(i, 3), dt) Then
                If Year(dt) = Year(d) And Month(dt) = Month(d) Then
                    amt = modMoney.ToDec(data(i, 7))
                    volume = volume + amt
                    If dt = d And (t = modConstants.TXN_WALLET_PAY Or t = modConstants.TXN_WALLET_OUT Or t = modConstants.TXN_WALLET_CASHOUT) Then spent = spent + amt
                End If
            End If
        End If
    Next i
End Sub

' "" when the wallet may take `credit` in and send `debit` out (customer wallets only; agent wallets are exempt).
Private Function LimitProblem(ByVal uow As clsUnitOfWork, ByVal w As clsAccount, ByVal credit As Variant, ByVal debit As Variant, ByVal isSpend As Boolean) As String
    Dim c As clsCustomer, maxBal As Variant, daily As Variant, monthly As Variant, singleMax As Variant, spent As Variant, vol As Variant, amt As Variant
    If IsAgentWallet(uow, w.AccountNo) Then Exit Function
    Set c = repoCustomer.FindCustomer(w.CustomerID)
    If c Is Nothing Then LimitProblem = modMessages.UnknownCustomer(w.CustomerID): Exit Function
    If Not LimitsFor(c.KYCLevel, maxBal, daily, monthly, singleMax) Then LimitProblem = "No wallet limits for KYC level " & c.KYCLevel & ".": Exit Function
    amt = credit
    If debit > amt Then amt = debit
    If amt > singleMax Then LimitProblem = "One transaction is limited to " & modMoney.FormatAmount(singleMax) & " MMK for KYC level " & c.KYCLevel & ".": Exit Function
    If credit > 0 And w.Balance + credit > maxBal Then LimitProblem = "The wallet balance would be above the maximum " & modMoney.FormatAmount(maxBal) & " MMK for KYC level " & c.KYCLevel & ".": Exit Function
    Usage w, uow.BusinessDate, spent, vol
    If isSpend And spent + debit > daily Then LimitProblem = "The daily spending limit is " & modMoney.FormatAmount(daily) & " MMK (already spent " & modMoney.FormatAmount(spent) & ").": Exit Function
    If vol + credit + debit > monthly Then LimitProblem = "The monthly volume limit is " & modMoney.FormatAmount(monthly) & " MMK (used " & modMoney.FormatAmount(vol) & ")."
End Function

' Adds the lines that move money between two branches through 1900.
Private Sub Move(ByVal lines As Collection, ByVal fromBranch As String, ByVal toBranch As String, ByVal amount As Variant)
    If fromBranch <> toBranch Then
        svcLedger.AddCr lines, fromBranch, modConstants.GL_INTERBRANCH, amount
        svcLedger.AddDr lines, toBranch, modConstants.GL_INTERBRANCH, amount
    End If
End Sub

' Trust rule: the net change of GL 2030 in these lines is matched at head office with 1110 and 1100, then the journal is posted.
Private Function PostWithTrust(ByVal uow As clsUnitOfWork, ByVal lines As Collection, ByVal jType As String, ByVal desc As String, ByVal grp As String) As clsResult
    Dim ln As clsJournalLine, delta As Variant, ho As String
    ho = modConstants.HEAD_OFFICE_ID
    delta = CDec(0)
    For Each ln In lines
        If ln.GLCode = modConstants.GL_WALLETS Then delta = delta + ln.Credit - ln.Debit
    Next ln
    If delta > 0 Then
        svcLedger.AddDr lines, ho, modConstants.GL_EMONEY_TRUST, delta
        svcLedger.AddCr lines, ho, modConstants.GL_TREASURY, delta
    ElseIf delta < 0 Then
        svcLedger.AddDr lines, ho, modConstants.GL_TREASURY, -delta
        svcLedger.AddCr lines, ho, modConstants.GL_EMONEY_TRUST, -delta
    End If
    Set PostWithTrust = svcLedger.PostJournal(uow, jType, desc, grp, lines)
End Function

' ---- Open wallet, merchant, agent ----

' Control: C13 customer. One wallet per customer, at the customer's home branch.
Public Function OpenWallet(ByVal user As clsUser, ByVal inp As Object) As clsResult
    Dim uow As clsUnitOfWork, msg As String, cust As clsCustomer, prod As clsProduct, a As clsAccount, k As Variant, w As clsAccount
    msg = svcSecurity.Denied(user, svcSecurity.ACT_WALLET)
    If Len(msg) = 0 Then msg = NeedTables()
    If Len(msg) > 0 Then Set OpenWallet = modMessages.FailResult(msg): Exit Function
    Set uow = New clsUnitOfWork
    uow.Init repoConfig.BusinessDate(), user.UserID
    Set cust = repoCustomer.FindCustomer(CStr(inp(modConstants.INP_CUSTOMER)))
    If cust Is Nothing Then Set OpenWallet = modMessages.FailResult(modMessages.UnknownCustomer(CStr(inp(modConstants.INP_CUSTOMER)))): Exit Function
    If cust.Status <> modConstants.STATUS_ACTIVE Then Set OpenWallet = modMessages.FailResult("Customer " & cust.CustomerID & " is not Active."): Exit Function
    If user.Role <> modConstants.ROLE_HEADOFFICE And cust.HomeBranchID <> user.BranchID Then Set OpenWallet = modMessages.FailResult("This customer belongs to branch " & cust.HomeBranchID & "."): Exit Function
    Set prod = repoProduct.FindProduct("WLT01")
    If prod Is Nothing Then Set OpenWallet = modMessages.FailResult("Wallet product WLT01 is missing. Run RunSetup."): Exit Function
    For Each k In uow.Accounts().Keys
        Set w = uow.Accounts()(k)
        If w.CustomerID = cust.CustomerID And w.ProductType = modConstants.PT_WALLET And w.Status <> modConstants.STATUS_CLOSED Then Set OpenWallet = modMessages.FailResult("The customer already has wallet " & w.AccountNo & "."): Exit Function
    Next k
    Set a = svcAccounts.CreateAccount(uow, cust, prod)
    repoAudit.Add uow, "OpenWallet", a.AccountNo, "Customer " & cust.CustomerID
    uow.Commit
    Set OpenWallet = modMessages.OkResult("Wallet " & a.AccountNo & " opened for " & cust.CustomerID & " (KYC level " & cust.KYCLevel & ").", 1, a.AccountNo)
End Function

' Control: C13 customer (becomes the merchant), C20 name, C21 category, C8 branch (blank = the customer's branch). Creates the merchant wallet.
Public Function AddMerchant(ByVal user As clsUser, ByVal inp As Object) As clsResult
    Dim uow As clsUnitOfWork, msg As String, cust As clsCustomer, prod As clsProduct, a As clsAccount, id As String
    msg = svcSecurity.Denied(user, svcSecurity.ACT_WALLET_ADMIN)
    If Len(msg) = 0 Then msg = NeedTables()
    If Len(msg) > 0 Then Set AddMerchant = modMessages.FailResult(msg): Exit Function
    Set uow = New clsUnitOfWork
    uow.Init repoConfig.BusinessDate(), user.UserID
    Set cust = repoCustomer.FindCustomer(CStr(inp(modConstants.INP_CUSTOMER)))
    If cust Is Nothing Then Set AddMerchant = modMessages.FailResult(modMessages.UnknownCustomer(CStr(inp(modConstants.INP_CUSTOMER)))): Exit Function
    Set prod = repoProduct.FindProduct("WLT01")
    If prod Is Nothing Then Set AddMerchant = modMessages.FailResult("Wallet product WLT01 is missing. Run RunSetup."): Exit Function
    Set a = svcAccounts.CreateAccount(uow, cust, prod)
    id = uow.NewId(modConstants.TBL_MERCHANTS, "MerchantID", "MC", 6)
    uow.AppendRow modConstants.TBL_MERCHANTS, Array(id, CStr(inp(modConstants.INP_VENDOR)), CStr(inp(modConstants.INP_DESC)), a.AccountNo, "QR-" & id, cust.HomeBranchID, modConstants.FEE_MERCHANT_MDR, modConstants.STATUS_ACTIVE)
    repoAudit.Add uow, "AddMerchant", id, "Wallet " & a.AccountNo
    uow.Commit
    Set AddMerchant = modMessages.OkResult("Merchant " & id & " added with wallet " & a.AccountNo & " and QR code QR-" & id & ".", 1, id)
End Function

' Control: C13 customer (becomes the agent), C20 name. Creates the agent float wallet.
Public Function AddAgent(ByVal user As clsUser, ByVal inp As Object) As clsResult
    Dim uow As clsUnitOfWork, msg As String, cust As clsCustomer, prod As clsProduct, a As clsAccount, id As String
    msg = svcSecurity.Denied(user, svcSecurity.ACT_WALLET_ADMIN)
    If Len(msg) = 0 Then msg = NeedTables()
    If Len(msg) > 0 Then Set AddAgent = modMessages.FailResult(msg): Exit Function
    Set uow = New clsUnitOfWork
    uow.Init repoConfig.BusinessDate(), user.UserID
    Set cust = repoCustomer.FindCustomer(CStr(inp(modConstants.INP_CUSTOMER)))
    If cust Is Nothing Then Set AddAgent = modMessages.FailResult(modMessages.UnknownCustomer(CStr(inp(modConstants.INP_CUSTOMER)))): Exit Function
    Set prod = repoProduct.FindProduct("WLT01")
    If prod Is Nothing Then Set AddAgent = modMessages.FailResult("Wallet product WLT01 is missing. Run RunSetup."): Exit Function
    Set a = svcAccounts.CreateAccount(uow, cust, prod)
    id = uow.NewId(modConstants.TBL_AGENTS, "AgentID", "AG", 6)
    uow.AppendRow modConstants.TBL_AGENTS, Array(id, CStr(inp(modConstants.INP_VENDOR)), a.AccountNo, cust.HomeBranchID, CDec(0), modConstants.STATUS_ACTIVE)
    repoAudit.Add uow, "AddAgent", id, "Float wallet " & a.AccountNo
    uow.Commit
    Set AddAgent = modMessages.OkResult("Agent " & id & " added with float wallet " & a.AccountNo & ".", 1, id)
End Function

' ---- Money moves ----

' Common start: user, tables, unit of work, amount (MMK, whole), processing branch.
Private Function Start(ByVal user As clsUser, ByVal inp As Object, ByVal action As String, ByRef uow As clsUnitOfWork, ByRef amt As Variant, ByRef p As String) As String
    Dim msg As String
    msg = svcSecurity.Denied(user, action)
    If Len(msg) = 0 Then msg = NeedTables()
    If Len(msg) = 0 Then
        If Not modArgs.ParseAmount(inp(modConstants.INP_AMOUNT), amt, msg) Then msg = msg
    End If
    If Len(msg) > 0 Then Start = msg: Exit Function
    Set uow = New clsUnitOfWork
    uow.Init repoConfig.BusinessDate(), user.UserID
    If Not svcSecurity.ResolveBranch(uow, user, CStr(inp(modConstants.INP_BRANCH)), p, msg) Then Start = msg
End Function

' Control: C5 wallet, C6 bank account, C7 amount.
Public Function TopUp(ByVal user As clsUser, ByVal inp As Object) As clsResult
    Dim uow As clsUnitOfWork, msg As String, amt As Variant, p As String, w As clsAccount, b As clsAccount, lines As Collection, jr As clsResult, grp As String
    msg = Start(user, inp, svcSecurity.ACT_WALLET, uow, amt, p)
    If Len(msg) > 0 Then Set TopUp = modMessages.FailResult(msg): Exit Function
    svcAccounts.WalletContext = True
    Set w = WalletAccount(uow, CStr(inp(modConstants.INP_ACCOUNT)), msg)
    If w Is Nothing Then svcAccounts.WalletContext = False: Set TopUp = modMessages.FailResult(msg): Exit Function
    Set b = BankAccount(uow, CStr(inp(modConstants.INP_COUNTER)), msg)
    If b Is Nothing Then svcAccounts.WalletContext = False: Set TopUp = modMessages.FailResult(msg): Exit Function
    msg = LimitProblem(uow, w, amt, CDec(0), False)
    If Len(msg) = 0 Then
        If Not svcAccounts.CanDebit(b, amt, False, msg) Then msg = msg
    End If
    If Len(msg) > 0 Then svcAccounts.WalletContext = False: Set TopUp = modMessages.FailResult(msg): Exit Function
    Set lines = New Collection
    svcLedger.AddDr lines, b.BranchID, b.DepositGL, amt, b.AccountNo
    Move lines, b.BranchID, w.BranchID, amt
    svcLedger.AddCr lines, w.BranchID, modConstants.GL_WALLETS, amt, w.AccountNo
    grp = uow.NewGroupId()
    Set jr = PostWithTrust(uow, lines, modConstants.JT_WALLET_OP, "Wallet top-up " & w.AccountNo, grp)
    svcAccounts.WalletContext = False
    If Not jr.Success Then Set TopUp = jr: Exit Function
    svcAccounts.Effect uow, b, modConstants.TXN_WALLET_TOPUP, modConstants.DIR_DEBIT, amt, w.AccountNo, p, "Top-up of wallet " & w.AccountNo, "", jr.RefID, grp, user.UserID, "", True
    svcAccounts.Effect uow, w, modConstants.TXN_WALLET_TOPUP, modConstants.DIR_CREDIT, amt, b.AccountNo, p, "Top-up from " & b.AccountNo, "", jr.RefID, grp, user.UserID, "", True
    uow.Commit
    Set TopUp = modMessages.OkResult("Wallet " & w.AccountNo & " topped up with " & modMoney.FormatAmount(amt) & " MMK from " & b.AccountNo & ". Wallet balance " & modMoney.FormatAmount(w.Balance) & ". Group " & grp & ".", 1, grp)
End Function

' Control: C5 wallet, C6 bank account, C7 amount. Fee WLT_CASHOUT (+ tax) is charged from the wallet.
Public Function CashOut(ByVal user As clsUser, ByVal inp As Object) As clsResult
    Dim uow As clsUnitOfWork, msg As String, amt As Variant, p As String, w As clsAccount, b As clsAccount, lines As Collection, jr As clsResult, grp As String, fee As Variant, tax As Variant, total As Variant
    msg = Start(user, inp, svcSecurity.ACT_WALLET, uow, amt, p)
    If Len(msg) > 0 Then Set CashOut = modMessages.FailResult(msg): Exit Function
    svcAccounts.WalletContext = True
    Set w = WalletAccount(uow, CStr(inp(modConstants.INP_ACCOUNT)), msg)
    If w Is Nothing Then svcAccounts.WalletContext = False: Set CashOut = modMessages.FailResult(msg): Exit Function
    Set b = BankAccount(uow, CStr(inp(modConstants.INP_COUNTER)), msg)
    If b Is Nothing Then svcAccounts.WalletContext = False: Set CashOut = modMessages.FailResult(msg): Exit Function
    If Not svcFees.Compute(modConstants.FEE_WLT_CASHOUT, amt, fee, tax, msg) Then svcAccounts.WalletContext = False: Set CashOut = modMessages.FailResult(msg): Exit Function
    total = amt + fee + tax
    msg = LimitProblem(uow, w, CDec(0), total, True)
    If Len(msg) = 0 Then
        If Not svcAccounts.CanDebit(w, total, False, msg) Then msg = msg
    End If
    If Len(msg) = 0 Then
        If Not svcAccounts.CanCredit(b, msg) Then msg = msg
    End If
    If Len(msg) > 0 Then svcAccounts.WalletContext = False: Set CashOut = modMessages.FailResult(msg): Exit Function
    Set lines = New Collection
    svcLedger.AddDr lines, w.BranchID, modConstants.GL_WALLETS, total, w.AccountNo
    svcLedger.AddCr lines, w.BranchID, modConstants.GL_FEE_INCOME, fee
    svcLedger.AddCr lines, w.BranchID, modConstants.GL_TAX_FEES, tax
    Move lines, w.BranchID, b.BranchID, amt
    svcLedger.AddCr lines, b.BranchID, b.DepositGL, amt, b.AccountNo
    grp = uow.NewGroupId()
    Set jr = PostWithTrust(uow, lines, modConstants.JT_WALLET_OP, "Wallet cash-out " & w.AccountNo, grp)
    svcAccounts.WalletContext = False
    If Not jr.Success Then Set CashOut = jr: Exit Function
    svcAccounts.Effect uow, w, modConstants.TXN_WALLET_CASHOUT, modConstants.DIR_DEBIT, amt, b.AccountNo, p, "Cash-out to " & b.AccountNo, "", jr.RefID, grp, user.UserID, "", True
    svcAccounts.FeeRows uow, w, fee, tax, modConstants.FEE_WLT_CASHOUT, b.AccountNo, p, jr.RefID, grp, user.UserID, ""
    svcAccounts.Effect uow, b, modConstants.TXN_WALLET_CASHOUT, modConstants.DIR_CREDIT, amt, w.AccountNo, p, "Cash-out from wallet " & w.AccountNo, "", jr.RefID, grp, user.UserID, "", True
    uow.Commit
    Set CashOut = modMessages.OkResult("Cash-out of " & modMoney.FormatAmount(amt) & " MMK to " & b.AccountNo & " posted, fee " & modMoney.FormatAmount(fee + tax) & ". Wallet balance " & modMoney.FormatAmount(w.Balance) & ". Group " & grp & ".", 1, grp)
End Function

' Control: C5 from wallet, C6 to wallet, C7 amount. The trust account does not change.
Public Function WalletTransfer(ByVal user As clsUser, ByVal inp As Object) As clsResult
    Dim uow As clsUnitOfWork, msg As String, amt As Variant, p As String, w1 As clsAccount, w2 As clsAccount, lines As Collection, jr As clsResult, grp As String
    msg = Start(user, inp, svcSecurity.ACT_WALLET, uow, amt, p)
    If Len(msg) > 0 Then Set WalletTransfer = modMessages.FailResult(msg): Exit Function
    svcAccounts.WalletContext = True
    Set w1 = WalletAccount(uow, CStr(inp(modConstants.INP_ACCOUNT)), msg)
    If w1 Is Nothing Then svcAccounts.WalletContext = False: Set WalletTransfer = modMessages.FailResult(msg): Exit Function
    Set w2 = WalletAccount(uow, CStr(inp(modConstants.INP_COUNTER)), msg)
    If w2 Is Nothing Then svcAccounts.WalletContext = False: Set WalletTransfer = modMessages.FailResult(msg): Exit Function
    If w1.AccountNo = w2.AccountNo Then svcAccounts.WalletContext = False: Set WalletTransfer = modMessages.FailResult(modMessages.MSG_SAME_ACCOUNT): Exit Function
    msg = LimitProblem(uow, w1, CDec(0), amt, True)
    If Len(msg) = 0 Then msg = LimitProblem(uow, w2, amt, CDec(0), False)
    If Len(msg) = 0 Then
        If Not svcAccounts.CanDebit(w1, amt, False, msg) Then msg = msg
    End If
    If Len(msg) > 0 Then svcAccounts.WalletContext = False: Set WalletTransfer = modMessages.FailResult(msg): Exit Function
    Set lines = New Collection
    svcLedger.AddDr lines, w1.BranchID, modConstants.GL_WALLETS, amt, w1.AccountNo
    Move lines, w1.BranchID, w2.BranchID, amt
    svcLedger.AddCr lines, w2.BranchID, modConstants.GL_WALLETS, amt, w2.AccountNo
    grp = uow.NewGroupId()
    Set jr = PostWithTrust(uow, lines, modConstants.JT_WALLET_OP, "Wallet transfer " & w1.AccountNo & " to " & w2.AccountNo, grp)
    svcAccounts.WalletContext = False
    If Not jr.Success Then Set WalletTransfer = jr: Exit Function
    svcAccounts.Effect uow, w1, modConstants.TXN_WALLET_OUT, modConstants.DIR_DEBIT, amt, w2.AccountNo, p, "Wallet transfer to " & w2.AccountNo, "", jr.RefID, grp, user.UserID, "", True
    svcAccounts.Effect uow, w2, modConstants.TXN_WALLET_IN, modConstants.DIR_CREDIT, amt, w1.AccountNo, p, "Wallet transfer from " & w1.AccountNo, "", jr.RefID, grp, user.UserID, "", True
    uow.Commit
    Set WalletTransfer = modMessages.OkResult("Wallet transfer of " & modMoney.FormatAmount(amt) & " MMK posted. Group " & grp & ".", 1, grp)
End Function

' Control: C5 wallet, C37 merchant, C7 amount. Merchant receives A - fee - tax; fee to 4100, tax to 2210.
Public Function WalletPay(ByVal user As clsUser, ByVal inp As Object) As clsResult
    Dim uow As clsUnitOfWork, msg As String, amt As Variant, p As String, w As clsAccount, m As clsAccount, lines As Collection, jr As clsResult, grp As String, fee As Variant, tax As Variant
    Dim merchants As Variant, mr As Long, feeCode As String
    msg = Start(user, inp, svcSecurity.ACT_WALLET, uow, amt, p)
    If Len(msg) > 0 Then Set WalletPay = modMessages.FailResult(msg): Exit Function
    svcAccounts.WalletContext = True
    Set w = WalletAccount(uow, CStr(inp(modConstants.INP_ACCOUNT)), msg)
    If w Is Nothing Then svcAccounts.WalletContext = False: Set WalletPay = modMessages.FailResult(msg): Exit Function
    merchants = uow.WorkingTable(modConstants.TBL_MERCHANTS)
    mr = repoFx.FindFxById(merchants, CStr(inp(modConstants.INP_MERCHANT)))
    If mr = 0 Then svcAccounts.WalletContext = False: Set WalletPay = modMessages.FailResult("Unknown merchant (Control C37)."): Exit Function
    If CStr(merchants(mr, repoFx.MEC_STATUS)) <> modConstants.STATUS_ACTIVE Then svcAccounts.WalletContext = False: Set WalletPay = modMessages.FailResult("The merchant is not Active."): Exit Function
    Set m = WalletAccount(uow, CStr(merchants(mr, repoFx.MEC_WALLET)), msg)
    If m Is Nothing Then svcAccounts.WalletContext = False: Set WalletPay = modMessages.FailResult("Merchant wallet: " & msg): Exit Function
    feeCode = CStr(merchants(mr, repoFx.MEC_FEECODE))
    If Len(feeCode) = 0 Then feeCode = modConstants.FEE_MERCHANT_MDR
    If Not svcFees.Compute(feeCode, amt, fee, tax, msg) Then svcAccounts.WalletContext = False: Set WalletPay = modMessages.FailResult(msg): Exit Function
    msg = LimitProblem(uow, w, CDec(0), amt, True)
    If Len(msg) = 0 Then
        If Not svcAccounts.CanDebit(w, amt, False, msg) Then msg = msg
    End If
    If Len(msg) > 0 Then svcAccounts.WalletContext = False: Set WalletPay = modMessages.FailResult(msg): Exit Function
    Set lines = New Collection
    svcLedger.AddDr lines, w.BranchID, modConstants.GL_WALLETS, amt, w.AccountNo
    Move lines, w.BranchID, m.BranchID, amt
    svcLedger.AddCr lines, m.BranchID, modConstants.GL_WALLETS, amt - fee - tax, m.AccountNo
    svcLedger.AddCr lines, m.BranchID, modConstants.GL_FEE_INCOME, fee
    svcLedger.AddCr lines, m.BranchID, modConstants.GL_TAX_FEES, tax
    grp = uow.NewGroupId()
    Set jr = PostWithTrust(uow, lines, modConstants.JT_WALLET_OP, "QR payment " & w.AccountNo & " to " & CStr(merchants(mr, 1)), grp)
    svcAccounts.WalletContext = False
    If Not jr.Success Then Set WalletPay = jr: Exit Function
    svcAccounts.Effect uow, w, modConstants.TXN_WALLET_PAY, modConstants.DIR_DEBIT, amt, m.AccountNo, p, "QR payment to " & CStr(merchants(mr, 2)), CStr(merchants(mr, 1)), jr.RefID, grp, user.UserID, "", True
    svcAccounts.Effect uow, m, modConstants.TXN_MERCHANT_RECEIPT, modConstants.DIR_CREDIT, amt - fee - tax, w.AccountNo, p, "QR receipt (fee " & modMoney.FormatAmount(fee) & ", tax " & modMoney.FormatAmount(tax) & ")", CStr(merchants(mr, 1)), jr.RefID, grp, user.UserID, "", False
    uow.Commit
    Set WalletPay = modMessages.OkResult("Paid " & modMoney.FormatAmount(amt) & " MMK to " & CStr(merchants(mr, 2)) & ". The merchant receives " & modMoney.FormatAmount(amt - fee - tax) & " (fee " & modMoney.FormatAmount(fee) & ", tax " & modMoney.FormatAmount(tax) & "). Group " & grp & ".", 1, grp)
End Function

' ---- Agents ----

Private Function AgentRow(ByVal uow As clsUnitOfWork, ByVal id As String, ByRef data As Variant, ByRef msg As String) As Long
    data = uow.WorkingTable(modConstants.TBL_AGENTS)
    AgentRow = repoFx.FindFxById(data, id)
    If AgentRow = 0 Then msg = "Unknown agent (Control C38)."
End Function

Private Function FloatMove(ByVal user As clsUser, ByVal inp As Object, ByVal intoWallet As Boolean) As clsResult
    Dim uow As clsUnitOfWork, msg As String, amt As Variant, p As String, ag As Variant, ar As Long, w As clsAccount, lines As Collection, jr As clsResult, grp As String
    msg = Start(user, inp, svcSecurity.ACT_WALLET, uow, amt, p)
    If Len(msg) > 0 Then Set FloatMove = modMessages.FailResult(msg): Exit Function
    ar = AgentRow(uow, CStr(inp(modConstants.INP_AGENT)), ag, msg)
    If ar = 0 Then Set FloatMove = modMessages.FailResult(msg): Exit Function
    svcAccounts.WalletContext = True
    Set w = WalletAccount(uow, CStr(ag(ar, repoFx.AGC_WALLET)), msg)
    If w Is Nothing Then svcAccounts.WalletContext = False: Set FloatMove = modMessages.FailResult(msg): Exit Function
    Set lines = New Collection
    If intoWallet Then
        svcLedger.AddDr lines, p, modConstants.GL_VAULT, amt
        Move lines, p, w.BranchID, amt
        svcLedger.AddCr lines, w.BranchID, modConstants.GL_WALLETS, amt, w.AccountNo
    Else
        If w.Balance < amt Then svcAccounts.WalletContext = False: Set FloatMove = modMessages.FailResult("The agent wallet does not hold that much."): Exit Function
        If uow.GLNet(p, modConstants.GL_VAULT) < 0 Then svcAccounts.WalletContext = False: Set FloatMove = modMessages.FailResult(modMessages.MSG_VAULT_SHORT): Exit Function
        svcLedger.AddDr lines, w.BranchID, modConstants.GL_WALLETS, amt, w.AccountNo
        Move lines, w.BranchID, p, amt
        svcLedger.AddCr lines, p, modConstants.GL_VAULT, amt
    End If
    grp = uow.NewGroupId()
    Set jr = PostWithTrust(uow, lines, modConstants.JT_AGENT_OP, "Agent float " & IIf(intoWallet, "in ", "out ") & CStr(ag(ar, 1)), grp)
    svcAccounts.WalletContext = False
    If Not jr.Success Then Set FloatMove = jr: Exit Function
    If intoWallet Then
        svcAccounts.Effect uow, w, modConstants.TXN_AGENT_FLOAT_IN, modConstants.DIR_CREDIT, amt, "", p, "Agent float in", "", jr.RefID, grp, user.UserID, "", False
    Else
        svcAccounts.Effect uow, w, modConstants.TXN_AGENT_FLOAT_OUT, modConstants.DIR_DEBIT, amt, "", p, "Agent float out", "", jr.RefID, grp, user.UserID, "", False
    End If
    uow.Commit
    Set FloatMove = modMessages.OkResult("Agent float " & IIf(intoWallet, "in ", "out ") & modMoney.FormatAmount(amt) & " MMK posted. Agent wallet balance " & modMoney.FormatAmount(w.Balance) & ". Group " & grp & ".", 1, grp)
End Function

' Control: C38 agent, C7 amount, C8 branch where the cash is handed over.
Public Function AgentFloatIn(ByVal user As clsUser, ByVal inp As Object) As clsResult
    Set AgentFloatIn = FloatMove(user, inp, True)
End Function

Public Function AgentFloatOut(ByVal user As clsUser, ByVal inp As Object) As clsResult
    Set AgentFloatOut = FloatMove(user, inp, False)
End Function

' Cash-in: customer gives cash to the agent (agent wallet to customer wallet, commission to the agent). Cash-out: the agent pays cash to the customer.
Private Function AgentCash(ByVal user As clsUser, ByVal inp As Object, ByVal isCashIn As Boolean) As clsResult
    Dim uow As clsUnitOfWork, msg As String, amt As Variant, p As String, ag As Variant, ar As Long, aw As clsAccount, cw As clsAccount, lines As Collection, jr As clsResult, grp As String
    Dim comm As Variant, fee As Variant, tax As Variant, cash As Variant
    msg = Start(user, inp, svcSecurity.ACT_WALLET, uow, amt, p)
    If Len(msg) > 0 Then Set AgentCash = modMessages.FailResult(msg): Exit Function
    ar = AgentRow(uow, CStr(inp(modConstants.INP_AGENT)), ag, msg)
    If ar = 0 Then Set AgentCash = modMessages.FailResult(msg): Exit Function
    If CStr(ag(ar, repoFx.AGC_STATUS)) <> modConstants.STATUS_ACTIVE Then Set AgentCash = modMessages.FailResult("The agent is not Active."): Exit Function
    svcAccounts.WalletContext = True
    Set aw = WalletAccount(uow, CStr(ag(ar, repoFx.AGC_WALLET)), msg)
    If aw Is Nothing Then svcAccounts.WalletContext = False: Set AgentCash = modMessages.FailResult("Agent wallet: " & msg): Exit Function
    Set cw = WalletAccount(uow, CStr(inp(modConstants.INP_ACCOUNT)), msg)
    If cw Is Nothing Then svcAccounts.WalletContext = False: Set AgentCash = modMessages.FailResult(msg): Exit Function
    cash = modMoney.ToDec(ag(ar, repoFx.AGC_CASH))
    Set lines = New Collection
    If isCashIn Then
        comm = modMoney.RoundCcy(amt * repoConfig.GetDec(modConstants.CFG_AGENT_COMM_IN), modConstants.CCY_MMK)
        msg = LimitProblem(uow, cw, amt, CDec(0), False)
        If Len(msg) = 0 Then
            If Not svcAccounts.CanDebit(aw, amt, False, msg) Then msg = msg
        End If
        If Len(msg) > 0 Then svcAccounts.WalletContext = False: Set AgentCash = modMessages.FailResult(msg): Exit Function
        svcLedger.AddDr lines, aw.BranchID, modConstants.GL_WALLETS, amt, aw.AccountNo
        Move lines, aw.BranchID, cw.BranchID, amt
        svcLedger.AddCr lines, cw.BranchID, modConstants.GL_WALLETS, amt, cw.AccountNo
    Else
        comm = modMoney.RoundCcy(amt * repoConfig.GetDec(modConstants.CFG_AGENT_COMM_OUT), modConstants.CCY_MMK)
        If Not svcFees.Compute(modConstants.FEE_WLT_CASHOUT, amt, fee, tax, msg) Then svcAccounts.WalletContext = False: Set AgentCash = modMessages.FailResult(msg): Exit Function
        If cash < amt Then svcAccounts.WalletContext = False: Set AgentCash = modMessages.FailResult("The agent has only " & modMoney.FormatAmount(cash) & " MMK cash on hand."): Exit Function
        msg = LimitProblem(uow, cw, CDec(0), amt + fee + tax, True)
        If Len(msg) = 0 Then
            If Not svcAccounts.CanDebit(cw, amt + fee + tax, False, msg) Then msg = msg
        End If
        If Len(msg) > 0 Then svcAccounts.WalletContext = False: Set AgentCash = modMessages.FailResult(msg): Exit Function
        svcLedger.AddDr lines, cw.BranchID, modConstants.GL_WALLETS, amt + fee + tax, cw.AccountNo
        svcLedger.AddCr lines, cw.BranchID, modConstants.GL_FEE_INCOME, fee
        svcLedger.AddCr lines, cw.BranchID, modConstants.GL_TAX_FEES, tax
        Move lines, cw.BranchID, aw.BranchID, amt
        svcLedger.AddCr lines, aw.BranchID, modConstants.GL_WALLETS, amt, aw.AccountNo
    End If
    svcLedger.AddDr lines, aw.BranchID, modConstants.GL_AGENT_COMM, comm
    svcLedger.AddCr lines, aw.BranchID, modConstants.GL_WALLETS, comm, aw.AccountNo
    grp = uow.NewGroupId()
    Set jr = PostWithTrust(uow, lines, modConstants.JT_AGENT_OP, "Agent cash-" & IIf(isCashIn, "in ", "out ") & CStr(ag(ar, 1)), grp)
    svcAccounts.WalletContext = False
    If Not jr.Success Then Set AgentCash = jr: Exit Function
    If isCashIn Then
        svcAccounts.Effect uow, aw, modConstants.TXN_AGENT_CASHIN, modConstants.DIR_DEBIT, amt, cw.AccountNo, p, "Agent cash-in for " & cw.AccountNo, "", jr.RefID, grp, user.UserID, "", False
        svcAccounts.Effect uow, cw, modConstants.TXN_AGENT_CASHIN, modConstants.DIR_CREDIT, amt, aw.AccountNo, p, "Cash-in at agent " & CStr(ag(ar, 1)), "", jr.RefID, grp, user.UserID, "", True
        ag(ar, repoFx.AGC_CASH) = cash + amt
    Else
        svcAccounts.Effect uow, cw, modConstants.TXN_AGENT_CASHOUT, modConstants.DIR_DEBIT, amt, aw.AccountNo, p, "Cash-out at agent " & CStr(ag(ar, 1)), "", jr.RefID, grp, user.UserID, "", True
        svcAccounts.FeeRows uow, cw, fee, tax, modConstants.FEE_WLT_CASHOUT, aw.AccountNo, p, jr.RefID, grp, user.UserID, ""
        svcAccounts.Effect uow, aw, modConstants.TXN_AGENT_CASHOUT, modConstants.DIR_CREDIT, amt, cw.AccountNo, p, "Agent paid cash to " & cw.AccountNo, "", jr.RefID, grp, user.UserID, "", False
        ag(ar, repoFx.AGC_CASH) = cash - amt
    End If
    If comm > 0 Then svcAccounts.Effect uow, aw, IIf(isCashIn, modConstants.TXN_AGENT_CASHIN, modConstants.TXN_AGENT_CASHOUT), modConstants.DIR_CREDIT, comm, "", p, "Agent commission", "", jr.RefID, grp, user.UserID, "", False
    uow.ReplaceTable modConstants.TBL_AGENTS, ag
    uow.Commit
    Set AgentCash = modMessages.OkResult("Agent cash-" & IIf(isCashIn, "in", "out") & " of " & modMoney.FormatAmount(amt) & " MMK posted, commission " & modMoney.FormatAmount(comm) & ". Group " & grp & ".", 1, grp)
End Function

' Control: C38 agent, C5 customer wallet, C7 amount.
Public Function AgentCashIn(ByVal user As clsUser, ByVal inp As Object) As clsResult
    Set AgentCashIn = AgentCash(user, inp, True)
End Function

Public Function AgentCashOut(ByVal user As clsUser, ByVal inp As Object) As clsResult
    Set AgentCashOut = AgentCash(user, inp, False)
End Function
