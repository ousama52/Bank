Attribute VB_Name = "svcFxDeals"
Option Explicit

' Stage 6: exchange rates, the currency counter, conversion between a customer's own accounts, and fake market deals.

Private Function NeedTables() As String
    If Not modSheetUtils.TableExists(modConstants.TBL_FXDEALS) Or Not modSheetUtils.TableExists(modConstants.TBL_EXCHANGERATES) Then NeedTables = modMessages.MSG_SETUP_FIRST
End Function

Private Function IsForeignCurrency(ByVal ccy As String) As Boolean
    Dim data As Variant, i As Long
    If ccy = modConstants.CCY_MMK Or Len(ccy) = 0 Then Exit Function
    data = modSheetUtils.ReadTable(modConstants.TBL_CURRENCIES)
    If IsEmpty(data) Then Exit Function
    For i = 1 To UBound(data, 1)
        If StrComp(CStr(data(i, 1)), ccy, vbTextCompare) = 0 Then IsForeignCurrency = True: Exit Function
    Next i
End Function

' ---- Rates ----

' Control: C35 currency, C36 mid rate. Writes (or replaces) the row of the business date: Buy = Mid x (1 - FxSpread), Sell = Mid x (1 + FxSpread), 2 decimals.
Public Function SetRates(ByVal user As clsUser, ByVal inp As Object) As clsResult
    Dim uow As clsUnitOfWork, msg As String, mid As Variant, ccy As String, data As Variant, i As Long, rd As Date, found As Long, buy As Variant, sell As Variant, sp As Variant
    msg = svcSecurity.Denied(user, svcSecurity.ACT_FX_RATES)
    If Len(msg) = 0 Then msg = NeedTables()
    If Len(msg) > 0 Then Set SetRates = modMessages.FailResult(msg): Exit Function
    ccy = UCase$(CStr(inp(modConstants.INP_CURRENCY)))
    If Not IsForeignCurrency(ccy) Then Set SetRates = modMessages.FailResult("Unknown foreign currency " & ccy & " (Control C35)."): Exit Function
    If Not modMoney.TryToDec(inp(modConstants.INP_MIDRATE), mid) Then Set SetRates = modMessages.FailResult("Enter the mid rate in Control C36."): Exit Function
    If mid <= 0 Then Set SetRates = modMessages.FailResult("The mid rate must be above zero."): Exit Function
    Set uow = New clsUnitOfWork
    uow.Init repoConfig.BusinessDate(), user.UserID
    sp = repoConfig.GetDec(modConstants.CFG_FX_SPREAD)
    buy = modMoney.RoundTo(mid * (1 - sp), 2)
    sell = modMoney.RoundTo(mid * (1 + sp), 2)
    data = uow.WorkingTable(modConstants.TBL_EXCHANGERATES)
    If Not IsEmpty(data) Then
        For i = 1 To UBound(data, 1)
            If StrComp(CStr(data(i, repoFxRate.FXR_CCY)), ccy, vbTextCompare) = 0 And modDateUtils.TryToDate(data(i, repoFxRate.FXR_DATE), rd) Then
                If rd = uow.BusinessDate Then found = i
            End If
        Next i
    End If
    If found > 0 Then
        data(found, repoFxRate.FXR_BUY) = buy
        data(found, repoFxRate.FXR_SELL) = sell
        data(found, repoFxRate.FXR_MID) = mid
        uow.ReplaceTable modConstants.TBL_EXCHANGERATES, data
    Else
        uow.AppendRow modConstants.TBL_EXCHANGERATES, Array(uow.BusinessDate, ccy, buy, sell, mid)
    End If
    repoAudit.Add uow, "SetRates", ccy, "Mid " & mid & ", buy " & buy & ", sell " & sell
    uow.Commit
    Set SetRates = modMessages.OkResult("Rates for " & ccy & " on " & modDateUtils.ToIso(uow.BusinessDate) & ": mid " & mid & ", buy " & buy & ", sell " & sell & ".", 1)
End Function

' ---- Position limit ----

' True when the net open position of the currency, after changing by `delta` (foreign units), is beyond FxPositionLimitPct x capital (MMK).
Public Function BreachAfter(ByVal uow As clsUnitOfWork, ByVal ccy As String, ByVal delta As Variant, ByVal mid As Variant) As Boolean
    Dim holdings As Variant, liab As Variant, cap As Variant, openBase As Variant
    svcFx.FxPosition uow, ccy, holdings, liab
    cap = -(uow.GLNetAll(modConstants.GL_CAPITAL) + uow.GLNetAll(modConstants.GL_RETAINED))
    openBase = modMoney.RoundCcy((holdings - liab + delta) * mid, modConstants.CCY_MMK)
    If openBase < 0 Then openBase = -openBase
    BreachAfter = (openBase > repoConfig.GetDec(modConstants.CFG_FX_POSITION_PCT) * cap)
End Function

' ---- Deals ----
' args: kind = BUY (customer buys foreign cash with MMK cash), SELL (customer sells foreign cash), CONVERT (own accounts: account = C5, counter = C6),
' ccy, amount (foreign), branch (processing), purpose, note.
Public Function ExecDeal(ByVal uow As clsUnitOfWork, ByVal args As Object, ByVal createdBy As String, ByVal approvedBy As String) As clsResult
    Dim kind As String, ccy As String, amt As Variant, p As String, mid As Variant, buy As Variant, sell As Variant, msg As String, bA As Variant, x As Variant, lines As Collection
    Dim jr As clsResult, grp As String, id As String, spread As Variant, rate As Variant, dealType As String, cust As String
    kind = modArgs.Text(args, "kind")
    ccy = modArgs.Text(args, "ccy")
    amt = modArgs.Dec(args, "amount")
    p = modArgs.Text(args, "branch")
    If Not uow.Branches().Exists(p) Then Set ExecDeal = modMessages.FailResult(modMessages.UnknownBranch(p)): Exit Function
    If kind = "CONVERT" Then Set ExecDeal = ExecConvert(uow, args, createdBy, approvedBy): Exit Function
    If Not svcFx.MidFor(uow, ccy, mid, buy, sell, msg) Then Set ExecDeal = modMessages.FailResult(msg): Exit Function
    If modMoney.RoundTo(amt, modMoney.Decimals(ccy)) <> amt Then Set ExecDeal = modMessages.FailResult("A " & ccy & " amount can have at most " & modMoney.Decimals(ccy) & " decimals."): Exit Function
    bA = svcFx.BaseOf(amt, mid)
    Set lines = New Collection
    If kind = "BUY" Then
        rate = sell
        x = modMoney.RoundCcy(amt * sell, modConstants.CCY_MMK)
        If uow.FxBalance(p, modConstants.GL_FX_CASH, ccy) - amt < 0 Then Set ExecDeal = modMessages.FailResult(svcFx.MSG_NO_FX_CASH): Exit Function
        svcLedger.AddDr lines, p, modConstants.GL_VAULT, x
        svcLedger.AddCrFx lines, p, modConstants.GL_FX_CASH, ccy, amt, mid, bA
        spread = x - bA
        svcLedger.AddCr lines, p, modConstants.GL_FX_INCOME, spread
        dealType = "CounterSell"
    ElseIf kind = "SELL" Then
        rate = buy
        x = modMoney.RoundCcy(amt * buy, modConstants.CCY_MMK)
        If uow.GLNet(p, modConstants.GL_VAULT) - x < 0 Then Set ExecDeal = modMessages.FailResult(modMessages.MSG_VAULT_SHORT): Exit Function
        svcLedger.AddDrFx lines, p, modConstants.GL_FX_CASH, ccy, amt, mid, bA
        svcLedger.AddCr lines, p, modConstants.GL_VAULT, x
        spread = bA - x
        svcLedger.AddCr lines, p, modConstants.GL_FX_INCOME, spread
        dealType = "CounterBuy"
    Else
        Set ExecDeal = modMessages.FailResult("Unknown deal kind."): Exit Function
    End If
    grp = uow.NewGroupId()
    Set jr = svcLedger.PostJournal(uow, modConstants.JT_FX_DEAL, dealType & " " & ccy & " " & amt, grp, lines)
    If Not jr.Success Then Set ExecDeal = jr: Exit Function
    id = uow.NewId(modConstants.TBL_FXDEALS, "DealID", "XD", 6)
    uow.AppendRow modConstants.TBL_FXDEALS, repoFx.DealRow(id, uow.BusinessDate, dealType, ccy, amt, rate, mid, x, spread, p, "", "", modArgs.Text(args, "purpose"), jr.RefID, createdBy, approvedBy)
    repoAudit.Add uow, "FxDeal", id, dealType & " " & modMoney.FormatCcy(amt, ccy) & " at " & rate
    Set ExecDeal = modMessages.OkResult("Deal " & id & ": " & IIf(kind = "BUY", "customer buys ", "customer sells ") & modMoney.FormatCcy(amt, ccy) & " at " & rate & " = " & modMoney.FormatAmount(x) & " MMK (spread " & modMoney.FormatAmount(spread) & "). Group " & grp & ".", 1, grp)
End Function

' Conversion between the customer's own MMK account and foreign account (account order decides the direction).
Private Function ExecConvert(ByVal uow As clsUnitOfWork, ByVal args As Object, ByVal createdBy As String, ByVal approvedBy As String) As clsResult
    Dim a1 As clsAccount, a2 As clsAccount, mm As clsAccount, fa As clsAccount, amt As Variant, mid As Variant, buy As Variant, sell As Variant, msg As String, bA As Variant, x As Variant
    Dim lines As Collection, jr As clsResult, grp As String, id As String, spread As Variant, toForeign As Boolean, p As String, desc As String
    p = modArgs.Text(args, "branch")
    Set a1 = uow.GetAccount(modArgs.Text(args, "account"))
    Set a2 = uow.GetAccount(modArgs.Text(args, "counter"))
    If a1 Is Nothing Then Set ExecConvert = modMessages.FailResult(modMessages.UnknownAccount(modArgs.Text(args, "account"))): Exit Function
    If a2 Is Nothing Then Set ExecConvert = modMessages.FailResult(modMessages.UnknownAccount(modArgs.Text(args, "counter"))): Exit Function
    If a1.CustomerID <> a2.CustomerID Then Set ExecConvert = modMessages.FailResult("Both accounts must belong to the same customer."): Exit Function
    If (a1.Ccy = modConstants.CCY_MMK) = (a2.Ccy = modConstants.CCY_MMK) Then Set ExecConvert = modMessages.FailResult("Convert needs one MMK account (C5 or C6) and one foreign account."): Exit Function
    If a1.Ccy = modConstants.CCY_MMK Then
        Set mm = a1
        Set fa = a2
        toForeign = True
    Else
        Set mm = a2
        Set fa = a1
        toForeign = False
    End If
    If mm.BranchID <> fa.BranchID Then Set ExecConvert = modMessages.FailResult("Both accounts must be in the same branch."): Exit Function
    amt = modArgs.Dec(args, "amount")
    If Not svcFx.MidFor(uow, fa.Ccy, mid, buy, sell, msg) Then Set ExecConvert = modMessages.FailResult(msg): Exit Function
    If modMoney.RoundTo(amt, modMoney.Decimals(fa.Ccy)) <> amt Then Set ExecConvert = modMessages.FailResult("A " & fa.Ccy & " amount can have at most " & modMoney.Decimals(fa.Ccy) & " decimals."): Exit Function
    bA = svcFx.BaseOf(amt, mid)
    Set lines = New Collection
    desc = "Convert " & modMoney.FormatCcy(amt, fa.Ccy)
    If toForeign Then
        x = modMoney.RoundCcy(amt * sell, modConstants.CCY_MMK)
        If Not svcAccounts.CanDebit(mm, x, False, msg) Then Set ExecConvert = modMessages.FailResult(msg): Exit Function
        If Not svcAccounts.CanCredit(fa, msg) Then Set ExecConvert = modMessages.FailResult(msg): Exit Function
        spread = x - bA
        svcLedger.AddDr lines, mm.BranchID, mm.DepositGL, x, mm.AccountNo
        svcLedger.AddCrFx lines, fa.BranchID, fa.DepositGL, fa.Ccy, amt, mid, bA, fa.AccountNo
        svcLedger.AddCr lines, mm.BranchID, modConstants.GL_FX_INCOME, spread
    Else
        x = modMoney.RoundCcy(amt * buy, modConstants.CCY_MMK)
        If Not svcAccounts.CanDebit(fa, amt, False, msg) Then Set ExecConvert = modMessages.FailResult(msg): Exit Function
        If Not svcAccounts.CanCredit(mm, msg) Then Set ExecConvert = modMessages.FailResult(msg): Exit Function
        spread = bA - x
        svcLedger.AddDrFx lines, fa.BranchID, fa.DepositGL, fa.Ccy, amt, mid, bA, fa.AccountNo
        svcLedger.AddCr lines, mm.BranchID, mm.DepositGL, x, mm.AccountNo
        svcLedger.AddCr lines, mm.BranchID, modConstants.GL_FX_INCOME, spread
    End If
    grp = uow.NewGroupId()
    Set jr = svcLedger.PostJournal(uow, modConstants.JT_FX_DEAL, desc, grp, lines)
    If Not jr.Success Then Set ExecConvert = jr: Exit Function
    If toForeign Then
        svcAccounts.Effect uow, mm, modConstants.TXN_FX_CONVERT, modConstants.DIR_DEBIT, x, fa.AccountNo, p, desc, "", jr.RefID, grp, createdBy, approvedBy, True
        svcAccounts.Effect uow, fa, modConstants.TXN_FX_CONVERT, modConstants.DIR_CREDIT, amt, mm.AccountNo, p, desc, "", jr.RefID, grp, createdBy, approvedBy, True
    Else
        svcAccounts.Effect uow, fa, modConstants.TXN_FX_CONVERT, modConstants.DIR_DEBIT, amt, mm.AccountNo, p, desc, "", jr.RefID, grp, createdBy, approvedBy, True
        svcAccounts.Effect uow, mm, modConstants.TXN_FX_CONVERT, modConstants.DIR_CREDIT, x, fa.AccountNo, p, desc, "", jr.RefID, grp, createdBy, approvedBy, True
    End If
    id = uow.NewId(modConstants.TBL_FXDEALS, "DealID", "XD", 6)
    uow.AppendRow modConstants.TBL_FXDEALS, repoFx.DealRow(id, uow.BusinessDate, "Convert", fa.Ccy, amt, IIf(toForeign, sell, buy), mid, x, spread, p, mm.AccountNo, mm.CustomerID, modArgs.Text(args, "purpose"), jr.RefID, createdBy, approvedBy)
    repoAudit.Add uow, "FxDeal", id, "Convert " & modMoney.FormatCcy(amt, fa.Ccy) & " for " & mm.CustomerID
    Set ExecConvert = modMessages.OkResult("Deal " & id & ": " & IIf(toForeign, "MMK to ", "") & modMoney.FormatCcy(amt, fa.Ccy) & IIf(toForeign, "", " to MMK") & " = " & modMoney.FormatAmount(x) & " MMK (spread " & modMoney.FormatAmount(spread) & "). Group " & grp & ".", 1, grp)
End Function

' ---- Entry points ----

' kind = BUY or SELL (cash at the counter); Control: C35 currency, C7 foreign amount, C8 branch, C42 purpose.
Private Function Counter(ByVal user As clsUser, ByVal inp As Object, ByVal kind As String) As clsResult
    Dim uow As clsUnitOfWork, msg As String, ccy As String, amt As Variant, args As Object, p As String, mid As Variant, buy As Variant, sell As Variant, mmk As Variant, delta As Variant, r As clsResult
    msg = svcSecurity.Denied(user, svcSecurity.ACT_FX_DEAL)
    If Len(msg) = 0 Then msg = NeedTables()
    If Len(msg) > 0 Then Set Counter = modMessages.FailResult(msg): Exit Function
    ccy = UCase$(CStr(inp(modConstants.INP_CURRENCY)))
    If Not IsForeignCurrency(ccy) Then Set Counter = modMessages.FailResult("Unknown foreign currency (Control C35)."): Exit Function
    If Not modArgs.ParseAmountCcy(inp(modConstants.INP_AMOUNT), ccy, amt, msg) Then Set Counter = modMessages.FailResult(msg): Exit Function
    Set uow = New clsUnitOfWork
    uow.Init repoConfig.BusinessDate(), user.UserID
    If Not svcSecurity.ResolveBranch(uow, user, CStr(inp(modConstants.INP_BRANCH)), p, msg) Then Set Counter = modMessages.FailResult(msg): Exit Function
    If Not svcFx.MidFor(uow, ccy, mid, buy, sell, msg) Then Set Counter = modMessages.FailResult(msg): Exit Function
    Set args = modArgs.NewArgs()
    args("kind") = kind
    args("ccy") = ccy
    args("amount") = amt
    args("branch") = p
    args("purpose") = CStr(inp(modConstants.INP_PURPOSE))
    mmk = svcFx.BaseOf(amt, mid)
    If kind = "BUY" Then delta = -amt Else delta = amt
    If BreachAfter(uow, ccy, delta, mid) Then
        Set r = svcApproval.Request(uow, modConstants.REQ_FX_LIMIT, "", "", mmk, p, args, user.UserID)
    ElseIf mmk >= repoConfig.GetDec(modConstants.CFG_FX_LARGE) Then
        If Len(CStr(inp(modConstants.INP_PURPOSE))) = 0 Then Set Counter = modMessages.FailResult("A deal of " & modMoney.FormatAmount(mmk) & " MMK or more needs a purpose (Control C42)."): Exit Function
        Set r = svcApproval.Request(uow, modConstants.REQ_LARGE_FX, "", "", mmk, p, args, user.UserID)
    Else
        Set r = ExecDeal(uow, args, user.UserID, "")
    End If
    If r.Success Then uow.Commit
    Set Counter = r
End Function

Public Function BuyCurrency(ByVal user As clsUser, ByVal inp As Object) As clsResult
    Set BuyCurrency = Counter(user, inp, "BUY")
End Function

Public Function SellCurrency(ByVal user As clsUser, ByVal inp As Object) As clsResult
    Set SellCurrency = Counter(user, inp, "SELL")
End Function

' Control: C5 and C6 = the customer's MMK and foreign accounts (the order sets the direction), C7 foreign amount, C42 purpose.
Public Function ConvertCurrency(ByVal user As clsUser, ByVal inp As Object) As clsResult
    Dim uow As clsUnitOfWork, msg As String, a1 As clsAccount, a2 As clsAccount, fa As clsAccount, amt As Variant, args As Object, p As String, mid As Variant, buy As Variant, sell As Variant
    Dim mmk As Variant, delta As Variant, r As clsResult
    msg = svcSecurity.Denied(user, svcSecurity.ACT_FX_DEAL)
    If Len(msg) = 0 Then msg = NeedTables()
    If Len(msg) > 0 Then Set ConvertCurrency = modMessages.FailResult(msg): Exit Function
    Set uow = New clsUnitOfWork
    uow.Init repoConfig.BusinessDate(), user.UserID
    Set a1 = uow.GetAccount(CStr(inp(modConstants.INP_ACCOUNT)))
    Set a2 = uow.GetAccount(CStr(inp(modConstants.INP_COUNTER)))
    If a1 Is Nothing Or a2 Is Nothing Then Set ConvertCurrency = modMessages.FailResult("Both accounts (Control C5 and C6) must exist."): Exit Function
    If a1.Ccy = modConstants.CCY_MMK Then Set fa = a2 Else Set fa = a1
    If Not modArgs.ParseAmountCcy(inp(modConstants.INP_AMOUNT), fa.Ccy, amt, msg) Then Set ConvertCurrency = modMessages.FailResult(msg): Exit Function
    If Not svcSecurity.ResolveBranch(uow, user, CStr(inp(modConstants.INP_BRANCH)), p, msg) Then Set ConvertCurrency = modMessages.FailResult(msg): Exit Function
    If Not svcFx.MidFor(uow, fa.Ccy, mid, buy, sell, msg) Then Set ConvertCurrency = modMessages.FailResult(msg): Exit Function
    Set args = modArgs.NewArgs()
    args("kind") = "CONVERT"
    args("account") = a1.AccountNo
    args("counter") = a2.AccountNo
    args("ccy") = fa.Ccy
    args("amount") = amt
    args("branch") = p
    args("purpose") = CStr(inp(modConstants.INP_PURPOSE))
    mmk = svcFx.BaseOf(amt, mid)
    If a1.Ccy = modConstants.CCY_MMK Then delta = -amt Else delta = amt
    If BreachAfter(uow, fa.Ccy, delta, mid) Then
        Set r = svcApproval.Request(uow, modConstants.REQ_FX_LIMIT, a1.AccountNo, a2.AccountNo, mmk, p, args, user.UserID)
    ElseIf mmk >= repoConfig.GetDec(modConstants.CFG_FX_LARGE) Then
        If Len(CStr(inp(modConstants.INP_PURPOSE))) = 0 Then Set ConvertCurrency = modMessages.FailResult("A deal of " & modMoney.FormatAmount(mmk) & " MMK or more needs a purpose (Control C42)."): Exit Function
        Set r = svcApproval.Request(uow, modConstants.REQ_LARGE_FX, a1.AccountNo, a2.AccountNo, mmk, p, args, user.UserID)
    Else
        Set r = ExecDeal(uow, args, user.UserID, "")
    End If
    If r.Success Then uow.Commit
    Set ConvertCurrency = r
End Function

' Fake market (HeadOffice): BUY = Dr 1700, Dr 4200 (spread), Cr 1100 cost. SELL = Dr 1100 proceeds, Dr 4200, Cr 1700.
Private Function Market(ByVal user As clsUser, ByVal inp As Object, ByVal kind As String) As clsResult
    Dim uow As clsUnitOfWork, msg As String, ccy As String, amt As Variant, mid As Variant, buy As Variant, sell As Variant, bA As Variant, cost As Variant, lines As Collection, jr As clsResult
    Dim grp As String, id As String, ho As String
    ho = modConstants.HEAD_OFFICE_ID
    msg = svcSecurity.Denied(user, svcSecurity.ACT_FX_MARKET)
    If Len(msg) = 0 Then msg = NeedTables()
    If Len(msg) > 0 Then Set Market = modMessages.FailResult(msg): Exit Function
    ccy = UCase$(CStr(inp(modConstants.INP_CURRENCY)))
    If Not IsForeignCurrency(ccy) Then Set Market = modMessages.FailResult("Unknown foreign currency (Control C35)."): Exit Function
    If Not modArgs.ParseAmountCcy(inp(modConstants.INP_AMOUNT), ccy, amt, msg) Then Set Market = modMessages.FailResult(msg): Exit Function
    Set uow = New clsUnitOfWork
    uow.Init repoConfig.BusinessDate(), user.UserID
    If Not svcFx.MidFor(uow, ccy, mid, buy, sell, msg) Then Set Market = modMessages.FailResult(msg): Exit Function
    bA = svcFx.BaseOf(amt, mid)
    Set lines = New Collection
    If kind = "BUY" Then
        cost = modMoney.RoundCcy(amt * mid * (1 + repoConfig.GetDec(modConstants.CFG_MARKET_SPREAD)), modConstants.CCY_MMK)
        svcLedger.AddDrFx lines, ho, modConstants.GL_FX_CASH, ccy, amt, mid, bA
        svcLedger.AddDr lines, ho, modConstants.GL_FX_INCOME, cost - bA
        svcLedger.AddCr lines, ho, modConstants.GL_TREASURY, cost
    Else
        cost = modMoney.RoundCcy(amt * mid * (1 - repoConfig.GetDec(modConstants.CFG_MARKET_SPREAD)), modConstants.CCY_MMK)
        If uow.FxBalance(ho, modConstants.GL_FX_CASH, ccy) - amt < 0 Then Set Market = modMessages.FailResult("Head office does not hold that much " & ccy & "."): Exit Function
        svcLedger.AddDr lines, ho, modConstants.GL_TREASURY, cost
        svcLedger.AddDr lines, ho, modConstants.GL_FX_INCOME, bA - cost
        svcLedger.AddCrFx lines, ho, modConstants.GL_FX_CASH, ccy, amt, mid, bA
    End If
    grp = uow.NewGroupId()
    Set jr = svcLedger.PostJournal(uow, modConstants.JT_FX_DEAL, "Market " & kind & " " & ccy & " " & amt, grp, lines)
    If Not jr.Success Then Set Market = jr: Exit Function
    id = uow.NewId(modConstants.TBL_FXDEALS, "DealID", "XD", 6)
    uow.AppendRow modConstants.TBL_FXDEALS, repoFx.DealRow(id, uow.BusinessDate, IIf(kind = "BUY", "MarketBuy", "MarketSell"), ccy, amt, cost / amt, mid, cost, IIf(kind = "BUY", cost - bA, bA - cost), ho, "", "", "", jr.RefID, user.UserID, "")
    repoAudit.Add uow, "FxMarket", id, kind & " " & modMoney.FormatCcy(amt, ccy) & " for " & modMoney.FormatAmount(cost) & " MMK"
    uow.Commit
    Set Market = modMessages.OkResult("Market " & IIf(kind = "BUY", "purchase", "sale") & " " & id & ": " & modMoney.FormatCcy(amt, ccy) & " for " & modMoney.FormatAmount(cost) & " MMK (cost of dealing " & modMoney.FormatAmount(Abs(cost - bA)) & "). Group " & grp & ".", 1, grp)
End Function

Public Function MarketBuy(ByVal user As clsUser, ByVal inp As Object) As clsResult
    Set MarketBuy = Market(user, inp, "BUY")
End Function

Public Function MarketSell(ByVal user As clsUser, ByVal inp As Object) As clsResult
    Set MarketSell = Market(user, inp, "SELL")
End Function
