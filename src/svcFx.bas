Attribute VB_Name = "svcFx"
Option Explicit

' Stage 6: postings on foreign currency accounts and foreign cash. Every foreign leg is booked at the MID rate of the business date;
' the MMK side of a leg is RoundCcy(foreign x mid). Only GLs 1700, 2000, 2010, 2030 and 2200 carry a foreign tag.

Public Const MSG_NO_RATE As String = "No rate for today. Ask head office to set rates (RunSetRates)."
Public Const MSG_NO_FX_CASH As String = "Branch has no foreign cash. Ask head office for a foreign cash shipment."

' Rate row of the business date (from the unit of work, so rates set earlier in the same unit are seen).
Public Function MidFor(ByVal uow As clsUnitOfWork, ByVal ccy As String, ByRef mid As Variant, ByRef buy As Variant, ByRef sell As Variant, ByRef msg As String) As Boolean
    If ccy = modConstants.CCY_MMK Then
        mid = CDec(1)
        buy = CDec(1)
        sell = CDec(1)
        MidFor = True
        Exit Function
    End If
    If modSheetUtils.TableExists(modConstants.TBL_EXCHANGERATES) Then
        MidFor = repoFxRate.OnDate(uow.WorkingTable(modConstants.TBL_EXCHANGERATES), ccy, uow.BusinessDate, mid, buy, sell)
    End If
    If Not MidFor Then msg = MSG_NO_RATE
End Function

Public Function BaseOf(ByVal amount As Variant, ByVal mid As Variant) As Variant
    BaseOf = modMoney.RoundCcy(amount * mid, modConstants.CCY_MMK)
End Function

' An MMK amount (a fee) expressed in the account currency at mid, rounded to the currency decimals.
Public Function ToCcy(ByVal mmk As Variant, ByVal mid As Variant, ByVal ccy As String) As Variant
    ToCcy = modMoney.RoundTo(mmk / mid, modMoney.Decimals(ccy))
End Function

' Net open position data for one currency: holdings (1700) minus foreign liabilities (2000, 2010, 2030, 2200), all in the currency.
Public Sub FxPosition(ByVal uow As clsUnitOfWork, ByVal ccy As String, ByRef holdings As Variant, ByRef liabilities As Variant)
    holdings = uow.FxBalanceAll(modConstants.GL_FX_CASH, ccy)
    liabilities = -(uow.FxBalanceAll(modConstants.GL_SAVINGS, ccy) + uow.FxBalanceAll(modConstants.GL_CURRENT, ccy) + uow.FxBalanceAll(modConstants.GL_WALLETS, ccy) + uow.FxBalanceAll(modConstants.GL_WHT, ccy))
End Sub

' ---- Cash deposit ----
' P: Dr 1700 (foreign A). If P <> H: P: Cr 1900 base, H: Dr 1900 base. H: Cr deposit GL (foreign A).
Public Function FxDeposit(ByVal uow As clsUnitOfWork, ByVal a As clsAccount, ByVal amt As Variant, ByVal p As String, ByVal grp As String, ByVal txnType As String, _
                        ByVal jType As String, ByVal desc As String, ByVal note As String, ByVal createdBy As String, ByVal approvedBy As String) As clsResult
    Dim mid As Variant, buy As Variant, sell As Variant, msg As String, bA As Variant, lines As Collection, jr As clsResult, h As String
    If Not MidFor(uow, a.Ccy, mid, buy, sell, msg) Then Set FxDeposit = modMessages.FailResult(msg): Exit Function
    bA = BaseOf(amt, mid)
    h = a.BranchID
    Set lines = New Collection
    svcLedger.AddDrFx lines, p, modConstants.GL_FX_CASH, a.Ccy, amt, mid, bA
    If p <> h Then
        svcLedger.AddCr lines, p, modConstants.GL_INTERBRANCH, bA
        svcLedger.AddDr lines, h, modConstants.GL_INTERBRANCH, bA
    End If
    svcLedger.AddCrFx lines, h, a.DepositGL, a.Ccy, amt, mid, bA, a.AccountNo
    Set jr = svcLedger.PostJournal(uow, jType, desc, grp, lines)
    If Not jr.Success Then Set FxDeposit = jr: Exit Function
    svcAccounts.Effect uow, a, txnType, modConstants.DIR_CREDIT, amt, "", p, desc, note, jr.RefID, grp, createdBy, approvedBy, True
    Set FxDeposit = modMessages.OkResult("", 1, grp)
End Function

' ---- Cash withdrawal (and the cash payout when an account is closed: isClose = True, no fee, no minimum balance) ----
' H: Dr deposit GL (A + fee + tax). If P <> H: H: Cr 1900, P: Dr 1900, P: Cr 4100 fee, Cr 2210 tax. P: Cr 1700 (foreign A).
Public Function FxWithdraw(ByVal uow As clsUnitOfWork, ByVal a As clsAccount, ByVal amt As Variant, ByVal p As String, ByVal grp As String, ByVal desc As String, _
                         ByVal note As String, ByVal createdBy As String, ByVal approvedBy As String, ByVal isClose As Boolean) As clsResult
    Dim mid As Variant, buy As Variant, sell As Variant, msg As String, bA As Variant, lines As Collection, jr As clsResult, h As String
    Dim feeM As Variant, taxM As Variant, feeF As Variant, taxF As Variant, bF As Variant, bT As Variant, total As Variant, bTotal As Variant
    If Not MidFor(uow, a.Ccy, mid, buy, sell, msg) Then Set FxWithdraw = modMessages.FailResult(msg): Exit Function
    bA = BaseOf(amt, mid)
    h = a.BranchID
    feeF = CDec(0)
    taxF = CDec(0)
    If p <> h And Not isClose Then
        If Not svcFees.Compute(modConstants.FEE_CROSS_WD, bA, feeM, taxM, msg) Then Set FxWithdraw = modMessages.FailResult(msg): Exit Function
        feeF = ToCcy(feeM, mid, a.Ccy)
        taxF = ToCcy(taxM, mid, a.Ccy)
    End If
    bF = BaseOf(feeF, mid)
    bT = BaseOf(taxF, mid)
    total = amt + feeF + taxF
    bTotal = bA + bF + bT
    If Not svcAccounts.CanDebit(a, total, isClose, msg) Then Set FxWithdraw = modMessages.FailResult(msg): Exit Function
    If uow.FxBalance(p, modConstants.GL_FX_CASH, a.Ccy) - amt < 0 Then Set FxWithdraw = modMessages.FailResult(MSG_NO_FX_CASH): Exit Function
    Set lines = New Collection
    svcLedger.AddDrFx lines, h, a.DepositGL, a.Ccy, total, mid, bTotal, a.AccountNo
    If p <> h Then
        svcLedger.AddCr lines, h, modConstants.GL_INTERBRANCH, bTotal
        svcLedger.AddDr lines, p, modConstants.GL_INTERBRANCH, bTotal
    End If
    svcLedger.AddCr lines, p, modConstants.GL_FEE_INCOME, bF
    svcLedger.AddCr lines, p, modConstants.GL_TAX_FEES, bT
    svcLedger.AddCrFx lines, p, modConstants.GL_FX_CASH, a.Ccy, amt, mid, bA
    Set jr = svcLedger.PostJournal(uow, modConstants.JT_WITHDRAWAL, desc, grp, lines)
    If Not jr.Success Then Set FxWithdraw = jr: Exit Function
    svcAccounts.Effect uow, a, modConstants.TXN_WITHDRAWAL, modConstants.DIR_DEBIT, amt, "", p, desc, note, jr.RefID, grp, createdBy, approvedBy, True
    svcAccounts.FeeRows uow, a, feeF, taxF, modConstants.FEE_CROSS_WD, "", p, jr.RefID, grp, createdBy, approvedBy
    Set FxWithdraw = modMessages.OkResult("", CLng(feeF + taxF), grp)
End Function

' ---- Transfer between two accounts of the same currency ----
Public Function FxTransfer(ByVal uow As clsUnitOfWork, ByVal src As clsAccount, ByVal dst As clsAccount, ByVal amt As Variant, ByVal p As String, ByVal grp As String, _
                         ByVal outType As String, ByVal inType As String, ByVal jType As String, ByVal desc As String, ByVal note As String, ByVal createdBy As String, _
                         ByVal approvedBy As String, ByVal isSystem As Boolean) As clsResult
    Dim mid As Variant, buy As Variant, sell As Variant, msg As String, bA As Variant, lines As Collection, jr As clsResult, h1 As String, h2 As String
    Dim feeM As Variant, taxM As Variant, feeF As Variant, taxF As Variant, bF As Variant, bT As Variant, total As Variant, bTotal As Variant
    If Not MidFor(uow, src.Ccy, mid, buy, sell, msg) Then Set FxTransfer = modMessages.FailResult(msg): Exit Function
    bA = BaseOf(amt, mid)
    h1 = src.BranchID
    h2 = dst.BranchID
    feeF = CDec(0)
    taxF = CDec(0)
    If h1 <> h2 And Not isSystem Then
        If Not svcFees.Compute(modConstants.FEE_TRANSFER_BRANCH, bA, feeM, taxM, msg) Then Set FxTransfer = modMessages.FailResult(msg): Exit Function
        feeF = ToCcy(feeM, mid, src.Ccy)
        taxF = ToCcy(taxM, mid, src.Ccy)
    End If
    bF = BaseOf(feeF, mid)
    bT = BaseOf(taxF, mid)
    total = amt + feeF + taxF
    bTotal = bA + bF + bT
    If Not svcAccounts.CanDebit(src, total, isSystem, msg) Then Set FxTransfer = modMessages.FailResult(msg): Exit Function
    If Not svcAccounts.CanCredit(dst, msg) Then Set FxTransfer = modMessages.FailResult(msg): Exit Function
    Set lines = New Collection
    svcLedger.AddDrFx lines, h1, src.DepositGL, src.Ccy, total, mid, bTotal, src.AccountNo
    svcLedger.AddCr lines, h1, modConstants.GL_FEE_INCOME, bF
    svcLedger.AddCr lines, h1, modConstants.GL_TAX_FEES, bT
    If h1 = h2 Then
        svcLedger.AddCrFx lines, h1, dst.DepositGL, dst.Ccy, amt, mid, bA, dst.AccountNo
    Else
        svcLedger.AddCr lines, h1, modConstants.GL_INTERBRANCH, bA
        svcLedger.AddDr lines, h2, modConstants.GL_INTERBRANCH, bA
        svcLedger.AddCrFx lines, h2, dst.DepositGL, dst.Ccy, amt, mid, bA, dst.AccountNo
    End If
    Set jr = svcLedger.PostJournal(uow, jType, desc, grp, lines)
    If Not jr.Success Then Set FxTransfer = jr: Exit Function
    svcAccounts.Effect uow, src, outType, modConstants.DIR_DEBIT, amt, dst.AccountNo, p, desc, note, jr.RefID, grp, createdBy, approvedBy, Not isSystem
    svcAccounts.FeeRows uow, src, feeF, taxF, modConstants.FEE_TRANSFER_BRANCH, dst.AccountNo, p, jr.RefID, grp, createdBy, approvedBy
    svcAccounts.Effect uow, dst, inType, modConstants.DIR_CREDIT, amt, src.AccountNo, p, desc, note, jr.RefID, grp, createdBy, approvedBy, Not isSystem
    Set FxTransfer = modMessages.OkResult("", CLng(feeF + taxF), grp)
End Function

' ---- Interest posting on a foreign account: H: Dr 5000 (base), Cr deposit GL (net, foreign), Cr 2200 (tax, foreign) ----
Public Function FxInterest(ByVal uow As clsUnitOfWork, ByVal a As clsAccount, ByVal gross As Variant, ByVal tax As Variant, ByVal jType As String, ByVal grp As String, _
                         ByVal createdBy As String) As clsResult
    Dim mid As Variant, buy As Variant, sell As Variant, msg As String, lines As Collection, jr As clsResult, net As Variant, bNet As Variant, bW As Variant
    If Not MidFor(uow, a.Ccy, mid, buy, sell, msg) Then Set FxInterest = modMessages.FailResult(msg): Exit Function
    net = gross - tax
    bNet = BaseOf(net, mid)
    bW = BaseOf(tax, mid)
    Set lines = New Collection
    svcLedger.AddDr lines, a.BranchID, modConstants.GL_INT_EXPENSE, bNet + bW
    svcLedger.AddCrFx lines, a.BranchID, a.DepositGL, a.Ccy, net, mid, bNet, a.AccountNo
    svcLedger.AddCrFx lines, a.BranchID, modConstants.GL_WHT, a.Ccy, tax, mid, bW
    Set jr = svcLedger.PostJournal(uow, jType, "Interest " & a.AccountNo & " gross " & gross & " " & a.Ccy, grp, lines)
    If Not jr.Success Then Set FxInterest = jr: Exit Function
    svcAccounts.Effect uow, a, modConstants.TXN_INTEREST, modConstants.DIR_CREDIT, net, "", a.BranchID, "Interest gross " & gross & ", tax " & tax, "", jr.RefID, grp, createdBy, "", False
    Set FxInterest = modMessages.OkResult("", 1, jr.RefID)
End Function

' ---- Foreign cash shipment between head office and a branch: B: Dr 1700, Cr 1900 base; HO: Dr 1900 base, Cr 1700 ----
Public Function FxShipment(ByVal uow As clsUnitOfWork, ByVal b As String, ByVal ccy As String, ByVal amt As Variant, ByVal toBranch As Boolean, ByVal grp As String) As clsResult
    Dim mid As Variant, buy As Variant, sell As Variant, msg As String, bA As Variant, lines As Collection, jr As clsResult, ho As String, sender As String
    ho = modConstants.HEAD_OFFICE_ID
    If Not MidFor(uow, ccy, mid, buy, sell, msg) Then Set FxShipment = modMessages.FailResult(msg): Exit Function
    bA = BaseOf(amt, mid)
    If toBranch Then sender = ho Else sender = b
    If uow.FxBalance(sender, modConstants.GL_FX_CASH, ccy) - amt < 0 Then
        If toBranch Then
            Set FxShipment = modMessages.FailResult("Not enough " & ccy & " cash at head office. Buy currency first (RunMarketBuy).")
        Else
            Set FxShipment = modMessages.FailResult("The vault of " & b & " does not hold that much " & ccy & ".")
        End If
        Exit Function
    End If
    Set lines = New Collection
    If toBranch Then
        svcLedger.AddDrFx lines, b, modConstants.GL_FX_CASH, ccy, amt, mid, bA
        svcLedger.AddCr lines, b, modConstants.GL_INTERBRANCH, bA
        svcLedger.AddDr lines, ho, modConstants.GL_INTERBRANCH, bA
        svcLedger.AddCrFx lines, ho, modConstants.GL_FX_CASH, ccy, amt, mid, bA
    Else
        svcLedger.AddDr lines, b, modConstants.GL_INTERBRANCH, bA
        svcLedger.AddCrFx lines, b, modConstants.GL_FX_CASH, ccy, amt, mid, bA
        svcLedger.AddDrFx lines, ho, modConstants.GL_FX_CASH, ccy, amt, mid, bA
        svcLedger.AddCr lines, ho, modConstants.GL_INTERBRANCH, bA
    End If
    Set jr = svcLedger.PostJournal(uow, modConstants.JT_CASH_SHIPMENT, "Foreign cash shipment " & ccy & " " & amt & " " & IIf(toBranch, "to ", "from ") & b, grp, lines)
    Set FxShipment = jr
End Function

' True when the amount has no more decimals than the account currency allows (MMK: whole).
Public Function AmountOk(ByVal a As clsAccount, ByVal amt As Variant, ByRef msg As String) As Boolean
    If modMoney.RoundTo(amt, modMoney.Decimals(a.Ccy)) <> amt Then
        msg = "A " & a.Ccy & " amount can have at most " & modMoney.Decimals(a.Ccy) & " decimals."
        Exit Function
    End If
    AmountOk = True
End Function

' MMK value of an amount on an account (latest known mid), for the teller and approval limits.
Public Function LimitBase(ByVal uow As clsUnitOfWork, ByVal accountNo As String, ByVal amt As Variant) As Variant
    Dim a As clsAccount
    Set a = uow.GetAccount(accountNo)
    If a Is Nothing Then LimitBase = amt: Exit Function
    If a.Ccy = modConstants.CCY_MMK Then LimitBase = amt Else LimitBase = modMoney.RoundCcy(amt * repoFxRate.LatestMid(a.Ccy, uow.BusinessDate), modConstants.CCY_MMK)
End Function
