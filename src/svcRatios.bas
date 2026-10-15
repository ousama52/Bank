Attribute VB_Name = "svcRatios"
Option Explicit

' Stage 8: compliance ratios (reserve, capital adequacy, liquidity coverage, bad loans, FX position, single borrower, investments),
' their history and the Breach / Cleared events. Breaches do not stop the day.

Private Function Num(ByVal d As Object, ByVal key As String) As Variant
    If d.Exists(key) Then Num = d(key) Else Num = CDec(0)
End Function

' Equity for the ratios: 3000 + 3100 + profit so far this year. After a year-end close the income and expense accounts start from zero again,
' so their current balances are exactly this year's profit.
Public Function Capital(ByVal uow As clsUnitOfWork) As Variant
    Dim chart As Object, k As Variant, profit As Variant
    Set chart = uow.Chart()
    profit = CDec(0)
    For Each k In chart.Keys
        If repoGov.IsProfitGL(chart, CStr(k)) Then profit = profit - uow.GLNetAll(CStr(k))
    Next k
    Capital = -(uow.GLNetAll(modConstants.GL_CAPITAL) + uow.GLNetAll(modConstants.GL_RETAINED)) + profit
End Function

Private Function Deposits(ByVal uow As clsUnitOfWork) As Variant
    Deposits = -(uow.GLNetAll(modConstants.GL_SAVINGS) + uow.GLNetAll(modConstants.GL_CURRENT) + uow.GLNetAll(modConstants.GL_FIXED) + uow.GLNetAll(modConstants.GL_WALLETS))
End Function

Private Function LiquidAssets(ByVal uow As clsUnitOfWork) As Variant
    LiquidAssets = uow.GLNetAll(modConstants.GL_VAULT) + uow.GLNetAll(modConstants.GL_TREASURY) + uow.GLNetAll(modConstants.GL_GOV_SEC)
End Function

Private Function HasPledge(ByVal items As Variant, ByVal loanId As String) As Boolean
    Dim i As Long
    If IsEmpty(items) Then Exit Function
    For i = 1 To UBound(items, 1)
        If CStr(items(i, repoCollateral.COC_LOAN)) = loanId And CStr(items(i, repoCollateral.COC_LIEN)) = repoCollateral.LIEN_PLEDGED Then HasPledge = True: Exit Function
    Next i
End Function

' Risk-weighted assets from tblRiskWeights. Loans (1300) are split by security and class; allowance 1320 and accumulated depreciation 1510 are counted against their assets.
Public Function RiskWeightedAssets(ByVal uow As clsUnitOfWork) As Variant
    Dim w As Variant, i As Long, gl As String, total As Variant, bal As Variant, loans As Variant, items As Variant, prow As Variant, cls As Long, wt As Variant, lw As Variant, out As Variant
    total = CDec(0)
    w = modSheetUtils.ReadTable(modConstants.TBL_RISKWEIGHTS)
    If Not IsEmpty(w) Then
        For i = 1 To UBound(w, 1)
            gl = CStr(w(i, 3))
            If gl <> modConstants.GL_LOANS And gl <> modConstants.GL_LOAN_ALLOWANCE And gl <> modConstants.GL_ACCUM_DEP And IsNumeric(w(i, 4)) Then
                bal = uow.GLNetAll(gl)
                If gl = modConstants.GL_ASSETS Then bal = bal + uow.GLNetAll(modConstants.GL_ACCUM_DEP)
                If bal > 0 Then total = total + bal * modMoney.ToDec(w(i, 4))
            End If
        Next i
    End If
    If modSheetUtils.TableExists(modConstants.TBL_LOANS) Then
        loans = uow.WorkingTable(modConstants.TBL_LOANS)
        If modSheetUtils.TableExists(modConstants.TBL_COLLATERAL) Then items = uow.WorkingTable(modConstants.TBL_COLLATERAL)
        lw = CDec(0)
        If Not IsEmpty(loans) Then
            For i = 1 To UBound(loans, 1)
                If CStr(loans(i, repoLoan.LOC_STATUS)) = repoLoan.LS_ACTIVE Then
                    cls = CLng(modMoney.ToDec(loans(i, repoLoan.LOC_CLASS)))
                    If cls >= 3 Then
                        wt = CDec(1.5)
                    Else
                        prow = repoLoan.ProductRow(CStr(loans(i, repoLoan.LOC_PRODUCT)))
                        If Not IsEmpty(prow) And Not IsEmpty(prow) Then
                            If CStr(prow(3)) = "Home" Then
                                wt = CDec(0.5)
                            ElseIf HasPledge(items, CStr(loans(i, 1))) Then
                                wt = CDec(0.75)
                            Else
                                wt = CDec(1)
                            End If
                        Else
                            wt = CDec(1)
                        End If
                    End If
                    lw = lw + modMoney.ToDec(loans(i, repoLoan.LOC_OUTSTANDING)) * wt
                End If
            Next i
        End If
        lw = lw + uow.GLNetAll(modConstants.GL_LOAN_ALLOWANCE)
        If lw > 0 Then total = total + lw
    End If
    RiskWeightedAssets = total
End Function

' 30 day outflows less capped inflows.
Private Function NetOutflows(ByVal uow As clsUnitOfWork, ByVal d As Date) As Variant
    Dim rr As Variant, i As Long, out As Variant, inflow As Variant, a As clsAccount, k As Variant, run30 As Variant, runOther As Variant, sav As Variant, cur As Variant, wal As Variant
    Dim loans As Variant, sched As Variant, due As Date, st As String
    sav = CDec(0.1)
    cur = CDec(0.15)
    wal = CDec(0.15)
    run30 = CDec(1)
    runOther = CDec(0)
    rr = modSheetUtils.ReadTable(modConstants.TBL_RUNOFF)
    If Not IsEmpty(rr) Then
        For i = 1 To UBound(rr, 1)
            Select Case CStr(rr(i, 1))
                Case "Savings": sav = modMoney.ToDec(rr(i, 2))
                Case "Current": cur = modMoney.ToDec(rr(i, 2))
                Case "Wallet": wal = modMoney.ToDec(rr(i, 2))
                Case "FixedMaturing30": run30 = modMoney.ToDec(rr(i, 2))
                Case "FixedOther": runOther = modMoney.ToDec(rr(i, 2))
            End Select
        Next i
    End If
    out = -uow.GLNetAll(modConstants.GL_SAVINGS) * sav - uow.GLNetAll(modConstants.GL_CURRENT) * cur - uow.GLNetAll(modConstants.GL_WALLETS) * wal
    For Each k In uow.Accounts().Keys
        Set a = uow.Accounts()(k)
        If a.ProductType = modConstants.PT_FIXED And a.Status <> modConstants.STATUS_CLOSED And Not IsEmpty(a.MaturityDate) Then
            If CDate(a.MaturityDate) <= d + 30 Then out = out + a.Balance * run30 Else out = out + a.Balance * runOther
        End If
    Next k
    inflow = CDec(0)
    If modSheetUtils.TableExists(modConstants.TBL_LOANS) Then
        loans = uow.WorkingTable(modConstants.TBL_LOANS)
        If Not IsEmpty(loans) Then
            For i = 1 To UBound(loans, 1)
                If CStr(loans(i, repoLoan.LOC_STATUS)) = repoLoan.LS_APPROVED Then out = out + modMoney.ToDec(loans(i, repoLoan.LOC_PRINCIPAL))
            Next i
        End If
        sched = uow.WorkingTable(modConstants.TBL_LOANSCHEDULE)
        If Not IsEmpty(sched) Then
            For i = 1 To UBound(sched, 1)
                st = CStr(sched(i, repoLoan.SCC_STATUS))
                If st <> "Paid" And modDateUtils.TryToDate(sched(i, repoLoan.SCC_DUE), due) Then
                    If due <= d + 30 Then inflow = inflow + svcLoanPayments.RowUnpaid(sched, i) * CDec(0.5)
                End If
            Next i
        End If
    End If
    If inflow > out * CDec(0.75) Then inflow = out * CDec(0.75)
    NetOutflows = out - inflow
End Function

' Bad loan ratio and the largest single borrower share, from tblLoans.
Private Sub LoanQuality(ByVal uow As clsUnitOfWork, ByRef npl As Variant, ByRef largest As Variant)
    Dim loans As Variant, i As Long, total As Variant, bad As Variant, per As Object, k As Variant
    npl = CDec(0)
    largest = CDec(0)
    If Not modSheetUtils.TableExists(modConstants.TBL_LOANS) Then Exit Sub
    loans = uow.WorkingTable(modConstants.TBL_LOANS)
    If IsEmpty(loans) Then Exit Sub
    Set per = CreateObject("Scripting.Dictionary")
    total = CDec(0)
    bad = CDec(0)
    For i = 1 To UBound(loans, 1)
        If CStr(loans(i, repoLoan.LOC_STATUS)) = repoLoan.LS_ACTIVE Then
            total = total + modMoney.ToDec(loans(i, repoLoan.LOC_OUTSTANDING))
            If CLng(modMoney.ToDec(loans(i, repoLoan.LOC_CLASS))) >= 3 Then bad = bad + modMoney.ToDec(loans(i, repoLoan.LOC_OUTSTANDING))
            If per.Exists(CStr(loans(i, repoLoan.LOC_CUSTOMER))) Then per(CStr(loans(i, repoLoan.LOC_CUSTOMER))) = per(CStr(loans(i, repoLoan.LOC_CUSTOMER))) + modMoney.ToDec(loans(i, repoLoan.LOC_OUTSTANDING)) Else per(CStr(loans(i, repoLoan.LOC_CUSTOMER))) = modMoney.ToDec(loans(i, repoLoan.LOC_OUTSTANDING))
        End If
    Next i
    If total > 0 Then npl = bad / total
    For Each k In per.Keys
        If per(k) > largest Then largest = per(k)
    Next k
End Sub

' Worst foreign net open position (MMK) over all currencies that have a rate today.
Private Function WorstFxPosition(ByVal uow As clsUnitOfWork, ByVal d As Date) As Variant
    Dim cur As Variant, i As Long, ccy As String, h As Variant, l As Variant, mid As Variant, buy As Variant, sell As Variant, msg As String, v As Variant
    WorstFxPosition = CDec(0)
    If Not modSheetUtils.TableExists(modConstants.TBL_CURRENCIES) Then Exit Function
    cur = modSheetUtils.ReadTable(modConstants.TBL_CURRENCIES)
    If IsEmpty(cur) Then Exit Function
    For i = 1 To UBound(cur, 1)
        ccy = CStr(cur(i, 1))
        If ccy <> modConstants.CCY_MMK And Len(ccy) > 0 Then
            If svcFx.MidFor(uow, ccy, mid, buy, sell, msg) Then
                svcFx.FxPosition uow, ccy, h, l
                v = Abs(modMoney.RoundCcy((h - l) * mid, modConstants.CCY_MMK))
                If v > WorstFxPosition Then WorstFxPosition = v
            End If
        End If
    Next i
End Function

Private Function InvestLimitPct() As Variant
    Dim data As Variant, i As Long
    InvestLimitPct = CDec(0.3)
    If Not modSheetUtils.TableExists(modConstants.TBL_INVLIMITS) Then Exit Function
    data = modSheetUtils.ReadTable(modConstants.TBL_INVLIMITS)
    If IsEmpty(data) Then Exit Function
    For i = 1 To UBound(data, 1)
        If CStr(data(i, 2)) = "Total" Then InvestLimitPct = modMoney.ToDec(data(i, 4)): Exit Function
    Next i
End Function

Private Function Item(ByVal code As String, ByVal actual As Variant, ByVal limit As Variant, ByVal breachWhenBelow As Boolean) As Variant
    Dim st As String
    st = "Within"
    If breachWhenBelow Then
        If actual < limit Then st = "Breach"
    Else
        If actual > limit Then st = "Breach"
    End If
    Item = Array(code, actual, limit, breachWhenBelow, st)
End Function

' Every ratio of the business date d as a Collection of Array(code, actual, limit, breachWhenBelow, status). Amounts in MMK.
Public Function ComputeRatios(ByVal uow As clsUnitOfWork, ByVal d As Date) As Collection
    Dim c As Collection, dep As Variant, req As Variant, held As Variant, cap As Variant, rwa As Variant, net As Variant, npl As Variant, largest As Variant
    Set c = New Collection
    dep = Deposits(uow)
    req = repoConfig.GetDec("ReserveRatio") * dep
    held = LiquidAssets(uow)
    cap = Capital(uow)
    If req > 0 Then c.Add Item("RESERVE", held / req, CDec(1), True) Else c.Add Item("RESERVE", CDec(999), CDec(1), True)
    rwa = RiskWeightedAssets(uow)
    If rwa > 0 Then c.Add Item("CAR", cap / rwa, repoConfig.GetDec(modConstants.CFG_MIN_CAR), True) Else c.Add Item("CAR", CDec(999), repoConfig.GetDec(modConstants.CFG_MIN_CAR), True)
    net = NetOutflows(uow, d)
    If net > 0 Then c.Add Item("LCR", held / net, repoConfig.GetDec(modConstants.CFG_MIN_LCR), True) Else c.Add Item("LCR", CDec(999), repoConfig.GetDec(modConstants.CFG_MIN_LCR), True)
    LoanQuality uow, npl, largest
    c.Add Item("NPL", npl, repoConfig.GetDec(modConstants.CFG_MAX_NPL), False)
    If cap > 0 Then
        c.Add Item("FX_POSITION", WorstFxPosition(uow, d) / cap, repoConfig.GetDec(modConstants.CFG_FX_POSITION_PCT), False)
        c.Add Item("SINGLE_BORROWER", largest / cap, repoConfig.GetDec(modConstants.CFG_SINGLE_BORROWER), False)
        c.Add Item("INVEST_TOTAL", uow.GLNetAll(modConstants.GL_INVESTMENTS) / cap, InvestLimitPct(), False)
    Else
        c.Add Item("FX_POSITION", CDec(0), repoConfig.GetDec(modConstants.CFG_FX_POSITION_PCT), False)
        c.Add Item("SINGLE_BORROWER", CDec(0), repoConfig.GetDec(modConstants.CFG_SINGLE_BORROWER), False)
        c.Add Item("INVEST_TOTAL", CDec(0), InvestLimitPct(), False)
    End If
    Set ComputeRatios = c
End Function

' Writes today's rows of tblRatioHistory (replacing rows of the same date) and a Breach or Cleared event when a ratio changes status.
Public Function WriteHistory(ByVal uow As clsUnitOfWork, ByVal d As Date, ByVal ratios As Collection) As Long
    Dim hist As Variant, i As Long, r As Variant, prev As String, rd As Date, best As Date, row As Long, ev As String, breaches As Long
    If Not modSheetUtils.TableExists(modConstants.TBL_RATIOHISTORY) Then Exit Function
    hist = uow.WorkingTable(modConstants.TBL_RATIOHISTORY)
    For Each r In ratios
        prev = "Within"
        row = 0
        best = 0
        If Not IsEmpty(hist) Then
            For i = 1 To UBound(hist, 1)
                If CStr(hist(i, 2)) = CStr(r(0)) And modDateUtils.TryToDate(hist(i, 1), rd) Then
                    If rd = d Then
                        row = i
                    ElseIf rd < d And rd >= best Then
                        best = rd
                        prev = CStr(hist(i, 5))
                    End If
                End If
            Next i
        End If
        If row > 0 Then
            hist(row, 3) = r(1)
            hist(row, 4) = r(2)
            hist(row, 5) = r(4)
            uow.ReplaceTable modConstants.TBL_RATIOHISTORY, hist
        Else
            uow.AppendRow modConstants.TBL_RATIOHISTORY, Array(d, CStr(r(0)), r(1), r(2), CStr(r(4)))
        End If
        If CStr(r(4)) = "Breach" Then breaches = breaches + 1
        ev = ""
        If prev = "Within" And CStr(r(4)) = "Breach" Then ev = "Breach"
        If prev = "Breach" And CStr(r(4)) = "Within" Then ev = "Cleared"
        If Len(ev) > 0 And modSheetUtils.TableExists(modConstants.TBL_COMPLIANCE) Then
            uow.AppendRow modConstants.TBL_COMPLIANCE, Array(uow.NewId(modConstants.TBL_COMPLIANCE, "EventID", "CE", 6), d, CStr(r(0)), r(1), r(2), ev, "")
            clsLogger.Warn "svcRatios.WriteHistory", "Ratio " & CStr(r(0)) & " " & ev & " (actual " & Format$(CDbl(r(1)), "0.0000") & ", limit " & Format$(CDbl(r(2)), "0.0000") & ")."
        End If
    Next r
    WriteHistory = breaches
End Function

' RunComplianceCheck: the same as end-of-day step 8, on demand.
Public Function CheckCompliance(ByVal user As clsUser, ByVal inp As Object) As clsResult
    Dim uow As clsUnitOfWork, msg As String, ratios As Collection, n As Long, d As Date, r As Variant, txt As String
    msg = svcSecurity.Denied(user, svcSecurity.ACT_COMPLIANCE)
    If Len(msg) = 0 Then
        If Not modSheetUtils.TableExists(modConstants.TBL_RATIOHISTORY) Then msg = modMessages.MSG_SETUP_FIRST
    End If
    If Len(msg) > 0 Then Set CheckCompliance = modMessages.FailResult(msg): Exit Function
    Set uow = New clsUnitOfWork
    uow.Init repoConfig.BusinessDate(), user.UserID
    d = uow.BusinessDate
    Set ratios = ComputeRatios(uow, d)
    n = WriteHistory(uow, d, ratios)
    uow.Commit
    For Each r In ratios
        txt = txt & CStr(r(0)) & " " & Format$(CDbl(r(1)), "0.000") & " (" & CStr(r(4)) & "); "
    Next r
    Set CheckCompliance = modMessages.OkResult("Ratios written to tblRatioHistory. " & n & " breach(es). " & txt, n)
End Function
