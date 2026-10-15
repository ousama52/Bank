Attribute VB_Name = "svcInvest"
Option Explicit

' Stage 7: investing in businesses. Everything is held and funded at head office (GL 1600, 1610, 1100).

Private Function NeedTables() As String
    If Not modSheetUtils.TableExists(modConstants.TBL_INVESTMENTS) Or Not modSheetUtils.TableExists(modConstants.TBL_INVESTEES) Then NeedTables = modMessages.MSG_SETUP_FIRST
End Function

Private Function Capital(ByVal uow As clsUnitOfWork) As Variant
    Capital = -(uow.GLNetAll(modConstants.GL_CAPITAL) + uow.GLNetAll(modConstants.GL_RETAINED))
End Function

' MaxPctOfCapital of a limit row, or -1 when there is no such limit.
Private Function LimitPct(ByVal limitType As String, ByVal key As String) As Variant
    Dim data As Variant, i As Long
    LimitPct = CDec(-1)
    data = modSheetUtils.ReadTable(modConstants.TBL_INVLIMITS)
    If IsEmpty(data) Then Exit Function
    For i = 1 To UBound(data, 1)
        If CStr(data(i, 2)) = limitType And StrComp(CStr(data(i, 3)), key, vbTextCompare) = 0 Then LimitPct = modMoney.ToDec(data(i, 4)): Exit Function
    Next i
End Function

' Cost still at work for a filter: carrying value of Active investments plus the amount of Approved ones not funded yet.
Private Function Exposure(ByVal invs As Variant, ByVal filterCol As Long, ByVal filterValue As String, ByVal skipId As String) As Variant
    Dim i As Long, t As Variant, st As String, match As Boolean
    t = CDec(0)
    If Not IsEmpty(invs) Then
        For i = 1 To UBound(invs, 1)
            If CStr(invs(i, repoInvest.INC_ID)) <> skipId Then
                st = CStr(invs(i, repoInvest.INC_STATUS))
                If filterCol = 0 Then
                    match = True
                ElseIf filterCol = -1 Then
                    match = (CStr(invs(i, repoInvest.INC_GRADE)) = "D" Or CStr(invs(i, repoInvest.INC_GRADE)) = "E")
                Else
                    match = (CStr(invs(i, filterCol)) = filterValue)
                End If
                If match Then
                    If st = "Active" Then t = t + modMoney.ToDec(invs(i, repoInvest.INC_CARRYING))
                    If st = "Approved" Then t = t + modMoney.ToDec(invs(i, repoInvest.INC_AMOUNT))
                End If
            End If
        Next i
    End If
    Exposure = t
End Function

' Limits and liquidity after investing `amount` with this investee. "" when fine, else the reason with the numbers.
Private Function LimitProblem(ByVal uow As clsUnitOfWork, ByVal invs As Variant, ByVal investeeId As String, ByVal sector As String, ByVal grade As String, ByVal amount As Variant, ByVal skipId As String) As String
    Dim cap As Variant, pct As Variant, used As Variant, deposits As Variant, liquid As Variant, need As Variant
    cap = Capital(uow)
    pct = LimitPct("Total", "ALL")
    used = Exposure(invs, 0, "", skipId) + amount
    If pct >= 0 And used > pct * cap Then LimitProblem = "Total limit: investments " & modMoney.FormatAmount(used) & " would be above " & modMoney.FormatAmount(pct * cap) & " (" & Format$(CDbl(pct) * 100, "0") & "% of capital " & modMoney.FormatAmount(cap) & ").": Exit Function
    pct = LimitPct("Single", "ALL")
    used = Exposure(invs, repoInvest.INC_INVESTEE, investeeId, skipId) + amount
    If pct >= 0 And used > pct * cap Then LimitProblem = "Single investee limit: " & modMoney.FormatAmount(used) & " would be above " & modMoney.FormatAmount(pct * cap) & " (" & Format$(CDbl(pct) * 100, "0") & "% of capital).": Exit Function
    pct = LimitPct("Sector", sector)
    used = Exposure(invs, repoInvest.INC_SECTOR, sector, skipId) + amount
    If pct >= 0 And used > pct * cap Then LimitProblem = "Sector limit (" & sector & "): " & modMoney.FormatAmount(used) & " would be above " & modMoney.FormatAmount(pct * cap) & ".": Exit Function
    If grade = "D" Or grade = "E" Then
        pct = LimitPct("HighRisk", "D,E")
        used = Exposure(invs, -1, "", skipId) + amount
        If pct >= 0 And used > pct * cap Then LimitProblem = "High risk limit (grade D and E): " & modMoney.FormatAmount(used) & " would be above " & modMoney.FormatAmount(pct * cap) & ".": Exit Function
    End If
    deposits = -(uow.GLNetAll(modConstants.GL_SAVINGS) + uow.GLNetAll(modConstants.GL_CURRENT) + uow.GLNetAll(modConstants.GL_FIXED) + uow.GLNetAll(modConstants.GL_WALLETS))
    liquid = uow.GLNetAll(modConstants.GL_TREASURY) + uow.GLNetAll(modConstants.GL_VAULT) - amount
    need = repoConfig.GetDec("ReserveRatio") * deposits + repoConfig.GetDec(modConstants.CFG_INVEST_BUFFER)
    If liquid < need Then LimitProblem = "Liquidity: liquid funds after the investment " & modMoney.FormatAmount(liquid) & " would be below " & modMoney.FormatAmount(need) & " (reserve + buffer)."
End Function

' ---- Investees ----

' Control: C20 name, C21 sector, C32 risk grade (A to E).
Public Function AddInvestee(ByVal user As clsUser, ByVal inp As Object) As clsResult
    Dim uow As clsUnitOfWork, msg As String, id As String, g As String
    msg = svcSecurity.Denied(user, svcSecurity.ACT_INVEST)
    If Len(msg) = 0 Then msg = NeedTables()
    If Len(msg) > 0 Then Set AddInvestee = modMessages.FailResult(msg): Exit Function
    g = UCase$(Trim$(CStr(inp(modConstants.INP_GRADE))))
    If Not (g Like "[A-E]") Then Set AddInvestee = modMessages.FailResult("Risk grade must be A, B, C, D or E (Control C32)."): Exit Function
    If Len(CStr(inp(modConstants.INP_VENDOR))) = 0 Or Len(CStr(inp(modConstants.INP_DESC))) = 0 Then Set AddInvestee = modMessages.FailResult("Enter the investee name (C20) and sector (C21)."): Exit Function
    Set uow = New clsUnitOfWork
    uow.Init repoConfig.BusinessDate(), user.UserID
    id = uow.NewId(modConstants.TBL_INVESTEES, "InvesteeID", "IV", 6)
    uow.AppendRow modConstants.TBL_INVESTEES, Array(id, CStr(inp(modConstants.INP_VENDOR)), CStr(inp(modConstants.INP_DESC)), g, modConstants.STATUS_ACTIVE, CStr(inp(modConstants.INP_NOTE)))
    repoAudit.Add uow, "AddInvestee", id, CStr(inp(modConstants.INP_VENDOR))
    uow.Commit
    Set AddInvestee = modMessages.OkResult("Investee " & id & " added (" & CStr(inp(modConstants.INP_VENDOR)) & ", grade " & g & ").", 1, id)
End Function

' ---- Proposal and approval ----

' Control: C44 investee, C46 type, C7 amount, C10 term (months), C47 rate or share (0.14 or 14), C15 pay frequency (Monthly, Quarterly, AtMaturity).
Public Function Propose(ByVal user As clsUser, ByVal inp As Object) As clsResult
    Dim uow As clsUnitOfWork, msg As String, amt As Variant, invs As Variant, vees As Variant, vr As Long, typ As String, term As Variant, rate As Variant, freq As String, id As String
    Dim args As Object, r As clsResult
    msg = svcSecurity.Denied(user, svcSecurity.ACT_INVEST)
    If Len(msg) = 0 Then msg = NeedTables()
    If Len(msg) > 0 Then Set Propose = modMessages.FailResult(msg): Exit Function
    If Not modArgs.ParseAmount(inp(modConstants.INP_AMOUNT), amt, msg) Then Set Propose = modMessages.FailResult(msg): Exit Function
    typ = CStr(inp(modConstants.INP_INVTYPE))
    If typ <> "Equity" And typ <> "ProfitShare" And typ <> "FixedReturnNote" And typ <> "TradeFinance" Then Set Propose = modMessages.FailResult("Type must be Equity, ProfitShare, FixedReturnNote or TradeFinance (Control C46)."): Exit Function
    If Not modMoney.TryToDec(inp(modConstants.INP_RATESHARE), rate) Then Set Propose = modMessages.FailResult("Enter the rate or share in Control C47."): Exit Function
    If rate > 1 Then rate = rate / 100
    If rate <= 0 Or rate > 1 Then Set Propose = modMessages.FailResult("The rate or share must be above 0 and at most 100%."): Exit Function
    If Not modMoney.TryToDec(inp(modConstants.INP_TERM), term) Then term = CDec(0)
    If (typ = "FixedReturnNote" Or typ = "TradeFinance") And term <= 0 Then Set Propose = modMessages.FailResult("A note or trade finance needs a term in months (Control C10)."): Exit Function
    freq = StrConv(Trim$(CStr(inp(modConstants.INP_NOTE))), vbProperCase)
    If freq <> "Monthly" And freq <> "Quarterly" And freq <> "Atmaturity" Then freq = "Monthly"
    If freq = "Atmaturity" Then freq = "AtMaturity"
    Set uow = New clsUnitOfWork
    uow.Init repoConfig.BusinessDate(), user.UserID
    vees = uow.WorkingTable(modConstants.TBL_INVESTEES)
    vr = repoInvest.FindInvestById(vees, CStr(inp(modConstants.INP_INVESTEE)))
    If vr = 0 Then Set Propose = modMessages.FailResult("Unknown investee (Control C44)."): Exit Function
    If CStr(vees(vr, repoInvest.IVC_STATUS)) <> modConstants.STATUS_ACTIVE Then Set Propose = modMessages.FailResult("The investee is " & CStr(vees(vr, repoInvest.IVC_STATUS)) & "."): Exit Function
    invs = uow.WorkingTable(modConstants.TBL_INVESTMENTS)
    msg = LimitProblem(uow, invs, CStr(vees(vr, 1)), CStr(vees(vr, repoInvest.IVC_SECTOR)), CStr(vees(vr, repoInvest.IVC_GRADE)), amt, "")
    If Len(msg) > 0 Then Set Propose = modMessages.FailResult(msg): Exit Function
    id = uow.NewId(modConstants.TBL_INVESTMENTS, "InvestmentID", "IN", 6)
    uow.AppendRow modConstants.TBL_INVESTMENTS, Array(id, CStr(vees(vr, 1)), typ, amt, Empty, term, Empty, rate, freq, "Proposed", CDec(0), CDec(0), Empty, CStr(vees(vr, repoInvest.IVC_GRADE)), _
        CStr(vees(vr, repoInvest.IVC_SECTOR)), user.UserID, Empty, Empty)
    Set args = modArgs.NewArgs()
    args("investment") = id
    Set r = svcApproval.Request(uow, modConstants.REQ_INVESTMENT, "", "", amt, modConstants.HEAD_OFFICE_ID, args, user.UserID)
    uow.Commit
    Set Propose = modMessages.OkResult("Investment " & id & " proposed (" & typ & ", " & modMoney.FormatAmount(amt) & " MMK). " & r.Message, 1, id)
End Function

' Run on approval: the limits are checked again, then the investment is Approved (funding is a separate step).
Public Function ExecApproveInvestment(ByVal uow As clsUnitOfWork, ByVal args As Object, ByVal createdBy As String, ByVal approvedBy As String) As clsResult
    Dim invs As Variant, ir As Long, vees As Variant, msg As String
    invs = uow.WorkingTable(modConstants.TBL_INVESTMENTS)
    ir = repoInvest.FindInvestById(invs, modArgs.Text(args, "investment"))
    If ir = 0 Then Set ExecApproveInvestment = modMessages.FailResult("Unknown investment."): Exit Function
    If CStr(invs(ir, repoInvest.INC_STATUS)) <> "Proposed" Then Set ExecApproveInvestment = modMessages.FailResult("Investment " & CStr(invs(ir, 1)) & " is " & CStr(invs(ir, repoInvest.INC_STATUS)) & "."): Exit Function
    msg = LimitProblem(uow, invs, CStr(invs(ir, repoInvest.INC_INVESTEE)), CStr(invs(ir, repoInvest.INC_SECTOR)), CStr(invs(ir, repoInvest.INC_GRADE)), modMoney.ToDec(invs(ir, repoInvest.INC_AMOUNT)), CStr(invs(ir, 1)))
    If Len(msg) > 0 Then Set ExecApproveInvestment = modMessages.FailResult(msg): Exit Function
    invs(ir, repoInvest.INC_STATUS) = "Approved"
    invs(ir, repoInvest.INC_APPROVEDBY) = approvedBy
    uow.ReplaceTable modConstants.TBL_INVESTMENTS, invs
    repoAudit.Add uow, "InvestmentApproved", CStr(invs(ir, 1)), "Approved by " & approvedBy
    Set ExecApproveInvestment = modMessages.OkResult("Investment " & CStr(invs(ir, 1)) & " is Approved. Fund it with RunFundInvestment.", 1, "")
End Function

' A rejected proposal keeps its reason.
Public Sub MarkInvestRejected(ByVal uow As clsUnitOfWork, ByVal investmentId As String, ByVal reason As String)
    Dim invs As Variant, ir As Long
    If Not modSheetUtils.TableExists(modConstants.TBL_INVESTMENTS) Then Exit Sub
    invs = uow.WorkingTable(modConstants.TBL_INVESTMENTS)
    ir = repoInvest.FindInvestById(invs, investmentId)
    If ir = 0 Then Exit Sub
    If CStr(invs(ir, repoInvest.INC_STATUS)) <> "Proposed" Then Exit Sub
    invs(ir, repoInvest.INC_STATUS) = "Rejected"
    If Len(reason) > 0 Then invs(ir, repoInvest.INC_NOTE) = reason
    uow.ReplaceTable modConstants.TBL_INVESTMENTS, invs
End Sub

' ---- Funding ----

' Control: C45 investment. HO: Dr 1600, Cr 1100.
Public Function Fund(ByVal user As clsUser, ByVal inp As Object) As clsResult
    Dim uow As clsUnitOfWork, msg As String, invs As Variant, ir As Long, lines As Collection, jr As clsResult, grp As String, amt As Variant, ho As String, mat As Date
    ho = modConstants.HEAD_OFFICE_ID
    msg = svcSecurity.Denied(user, svcSecurity.ACT_INVEST)
    If Len(msg) = 0 Then msg = NeedTables()
    If Len(msg) > 0 Then Set Fund = modMessages.FailResult(msg): Exit Function
    Set uow = New clsUnitOfWork
    uow.Init repoConfig.BusinessDate(), user.UserID
    invs = uow.WorkingTable(modConstants.TBL_INVESTMENTS)
    ir = repoInvest.FindInvestById(invs, CStr(inp(modConstants.INP_INVESTMENT)))
    If ir = 0 Then Set Fund = modMessages.FailResult("Unknown investment. Put the InvestmentID in Control C45."): Exit Function
    If CStr(invs(ir, repoInvest.INC_STATUS)) <> "Approved" Then Set Fund = modMessages.FailResult("Investment " & CStr(invs(ir, 1)) & " is " & CStr(invs(ir, repoInvest.INC_STATUS)) & ", not Approved."): Exit Function
    amt = modMoney.ToDec(invs(ir, repoInvest.INC_AMOUNT))
    msg = LimitProblem(uow, invs, CStr(invs(ir, repoInvest.INC_INVESTEE)), CStr(invs(ir, repoInvest.INC_SECTOR)), CStr(invs(ir, repoInvest.INC_GRADE)), amt, CStr(invs(ir, 1)))
    If Len(msg) > 0 Then Set Fund = modMessages.FailResult(msg): Exit Function
    Set lines = New Collection
    svcLedger.AddDr lines, ho, modConstants.GL_INVESTMENTS, amt
    svcLedger.AddCr lines, ho, modConstants.GL_TREASURY, amt
    grp = uow.NewGroupId()
    Set jr = svcLedger.PostJournal(uow, modConstants.JT_INVESTMENT, "Investment " & CStr(invs(ir, 1)), grp, lines)
    If Not jr.Success Then Set Fund = jr: Exit Function
    invs(ir, repoInvest.INC_STATUS) = "Active"
    invs(ir, repoInvest.INC_START) = uow.BusinessDate
    invs(ir, repoInvest.INC_CARRYING) = amt
    invs(ir, repoInvest.INC_ACCRUED) = CDec(0)
    invs(ir, repoInvest.INC_LASTACCRUAL) = uow.BusinessDate
    If modMoney.ToDec(invs(ir, repoInvest.INC_TERM)) > 0 Then
        mat = modDateUtils.AddMonths(uow.BusinessDate, CLng(modMoney.ToDec(invs(ir, repoInvest.INC_TERM))))
        invs(ir, repoInvest.INC_MATURITY) = mat
    End If
    uow.ReplaceTable modConstants.TBL_INVESTMENTS, invs
    repoAudit.Add uow, "InvestmentFunded", CStr(invs(ir, 1)), modMoney.FormatAmount(amt)
    uow.Commit
    Set Fund = modMessages.OkResult("Investment " & CStr(invs(ir, 1)) & " funded: " & modMoney.FormatAmount(amt) & " MMK moved from treasury. Group " & grp & ".", 1, grp)
End Function

' ---- Returns ----

' Control: C45 investment, C7 amount received, C15 type (Interest, Dividend, ProfitShare). Cr 1610 up to the unpaid recognized return, the rest Cr 4300. Dividend: all Cr 4300.
Public Function ReceiveReturn(ByVal user As clsUser, ByVal inp As Object) As clsResult
    Dim uow As clsUnitOfWork, msg As String, amt As Variant, invs As Variant, ir As Long, rets As Variant, unpaid As Variant, c1610 As Variant, lines As Collection, jr As clsResult, grp As String, typ As String, ho As String
    ho = modConstants.HEAD_OFFICE_ID
    msg = svcSecurity.Denied(user, svcSecurity.ACT_INVEST)
    If Len(msg) = 0 Then msg = NeedTables()
    If Len(msg) > 0 Then Set ReceiveReturn = modMessages.FailResult(msg): Exit Function
    If Not modArgs.ParseAmount(inp(modConstants.INP_AMOUNT), amt, msg) Then Set ReceiveReturn = modMessages.FailResult(msg): Exit Function
    Set uow = New clsUnitOfWork
    uow.Init repoConfig.BusinessDate(), user.UserID
    invs = uow.WorkingTable(modConstants.TBL_INVESTMENTS)
    ir = repoInvest.FindInvestById(invs, CStr(inp(modConstants.INP_INVESTMENT)))
    If ir = 0 Then Set ReceiveReturn = modMessages.FailResult("Unknown investment (Control C45)."): Exit Function
    typ = StrConv(Trim$(CStr(inp(modConstants.INP_NOTE))), vbProperCase)
    If typ <> "Dividend" And typ <> "Profitshare" Then typ = "Interest"
    If typ = "Profitshare" Then typ = "ProfitShare"
    rets = uow.WorkingTable(modConstants.TBL_INVRETURNS)
    unpaid = repoInvest.UnpaidRecognized(rets, CStr(invs(ir, 1)))
    If typ = "Dividend" Then c1610 = CDec(0) Else c1610 = unpaid
    If c1610 > amt Then c1610 = amt
    If c1610 < 0 Then c1610 = CDec(0)
    Set lines = New Collection
    svcLedger.AddDr lines, ho, modConstants.GL_TREASURY, amt
    svcLedger.AddCr lines, ho, modConstants.GL_INVEST_RECEIVABLE, c1610
    svcLedger.AddCr lines, ho, modConstants.GL_INVEST_INCOME, amt - c1610
    grp = uow.NewGroupId()
    Set jr = svcLedger.PostJournal(uow, modConstants.JT_INV_RETURN, "Investment return " & CStr(invs(ir, 1)), grp, lines)
    If Not jr.Success Then Set ReceiveReturn = jr: Exit Function
    uow.AppendRow modConstants.TBL_INVRETURNS, repoInvest.ReturnRow(uow.NewId(modConstants.TBL_INVRETURNS, "ReturnID", "IR", 6), CStr(invs(ir, 1)), uow.BusinessDate, amt, typ, jr.RefID, "Received;1610=" & CStr(c1610))
    repoAudit.Add uow, "InvestmentReturn", CStr(invs(ir, 1)), modMoney.FormatAmount(amt)
    uow.Commit
    Set ReceiveReturn = modMessages.OkResult("Return of " & modMoney.FormatAmount(amt) & " MMK received on " & CStr(invs(ir, 1)) & ": " & modMoney.FormatAmount(c1610) & " cleared the receivable, " & modMoney.FormatAmount(amt - c1610) & " is income. Group " & grp & ".", 1, grp)
End Function

' Control: C45 investment, C48 reported profit (or loss, negative). Income = RoundCcy(profit x share). Profit: Dr 1610, Cr 4300. Loss share: Dr 5330, Cr 1600.
Public Function RecordProfitShare(ByVal user As clsUser, ByVal inp As Object) As clsResult
    Dim uow As clsUnitOfWork, msg As String, profit As Variant, invs As Variant, ir As Long, share As Variant, income As Variant, lines As Collection, jr As clsResult, grp As String, ho As String, carry As Variant
    ho = modConstants.HEAD_OFFICE_ID
    msg = svcSecurity.Denied(user, svcSecurity.ACT_INVEST)
    If Len(msg) = 0 Then msg = NeedTables()
    If Len(msg) > 0 Then Set RecordProfitShare = modMessages.FailResult(msg): Exit Function
    If Not modMoney.TryToDec(inp(modConstants.INP_PROFIT), profit) Then Set RecordProfitShare = modMessages.FailResult("Enter the reported profit or loss in Control C48."): Exit Function
    Set uow = New clsUnitOfWork
    uow.Init repoConfig.BusinessDate(), user.UserID
    invs = uow.WorkingTable(modConstants.TBL_INVESTMENTS)
    ir = repoInvest.FindInvestById(invs, CStr(inp(modConstants.INP_INVESTMENT)))
    If ir = 0 Then Set RecordProfitShare = modMessages.FailResult("Unknown investment (Control C45)."): Exit Function
    If CStr(invs(ir, repoInvest.INC_STATUS)) <> "Active" Then Set RecordProfitShare = modMessages.FailResult("The investment is not Active."): Exit Function
    If CStr(invs(ir, repoInvest.INC_TYPE)) <> "ProfitShare" And CStr(invs(ir, repoInvest.INC_TYPE)) <> "Equity" Then Set RecordProfitShare = modMessages.FailResult("Profit is recorded for ProfitShare and Equity investments only."): Exit Function
    share = modMoney.ToDec(invs(ir, repoInvest.INC_RATE))
    income = modMoney.RoundCcy(profit * share, modConstants.CCY_MMK)
    carry = modMoney.ToDec(invs(ir, repoInvest.INC_CARRYING))
    If income = 0 Then Set RecordProfitShare = modMessages.FailResult("The share of this result is zero."): Exit Function
    Set lines = New Collection
    grp = uow.NewGroupId()
    If income > 0 Then
        svcLedger.AddDr lines, ho, modConstants.GL_INVEST_RECEIVABLE, income
        svcLedger.AddCr lines, ho, modConstants.GL_INVEST_INCOME, income
        Set jr = svcLedger.PostJournal(uow, modConstants.JT_INV_RETURN, "Profit share " & CStr(invs(ir, 1)), grp, lines)
        If Not jr.Success Then Set RecordProfitShare = jr: Exit Function
        uow.AppendRow modConstants.TBL_INVRETURNS, repoInvest.ReturnRow(uow.NewId(modConstants.TBL_INVRETURNS, "ReturnID", "IR", 6), CStr(invs(ir, 1)), uow.BusinessDate, income, "ProfitShare", jr.RefID, "Recognized")
    Else
        income = -income
        If income > carry Then income = carry
        svcLedger.AddDr lines, ho, modConstants.GL_INVEST_LOSS, income
        svcLedger.AddCr lines, ho, modConstants.GL_INVESTMENTS, income
        Set jr = svcLedger.PostJournal(uow, modConstants.JT_INV_VALUATION, "Loss share " & CStr(invs(ir, 1)), grp, lines)
        If Not jr.Success Then Set RecordProfitShare = jr: Exit Function
        invs(ir, repoInvest.INC_CARRYING) = carry - income
        uow.ReplaceTable modConstants.TBL_INVESTMENTS, invs
    End If
    repoAudit.Add uow, "ProfitShare", CStr(invs(ir, 1)), "Reported " & modMoney.FormatAmount(profit)
    uow.Commit
    Set RecordProfitShare = modMessages.OkResult("Profit share on " & CStr(invs(ir, 1)) & " booked: " & modMoney.FormatAmount(income) & " MMK " & IIf(profit >= 0, "income", "loss") & ". Group " & grp & ".", 1, grp)
End Function

' ---- Valuation, disposal, write-off ----

' Control: C45 investment, C48 new value, C15 method (Cost, Appraisal, Multiple). Below carrying value: Dr 5330, Cr 1600. Above: write-up never above cost.
Public Function ValueInvestment(ByVal user As clsUser, ByVal inp As Object) As clsResult
    Dim uow As clsUnitOfWork, msg As String, nv As Variant, invs As Variant, ir As Long, carry As Variant, cost As Variant, diff As Variant, lines As Collection, jr As clsResult, grp As String, ho As String, method As String
    ho = modConstants.HEAD_OFFICE_ID
    msg = svcSecurity.Denied(user, svcSecurity.ACT_INVEST)
    If Len(msg) = 0 Then msg = NeedTables()
    If Len(msg) > 0 Then Set ValueInvestment = modMessages.FailResult(msg): Exit Function
    If Not modMoney.TryToDec(inp(modConstants.INP_PROFIT), nv) Then Set ValueInvestment = modMessages.FailResult("Enter the new value in Control C48."): Exit Function
    If nv < 0 Then Set ValueInvestment = modMessages.FailResult("A value cannot be negative."): Exit Function
    Set uow = New clsUnitOfWork
    uow.Init repoConfig.BusinessDate(), user.UserID
    invs = uow.WorkingTable(modConstants.TBL_INVESTMENTS)
    ir = repoInvest.FindInvestById(invs, CStr(inp(modConstants.INP_INVESTMENT)))
    If ir = 0 Then Set ValueInvestment = modMessages.FailResult("Unknown investment (Control C45)."): Exit Function
    If CStr(invs(ir, repoInvest.INC_STATUS)) <> "Active" Then Set ValueInvestment = modMessages.FailResult("The investment is not Active."): Exit Function
    carry = modMoney.ToDec(invs(ir, repoInvest.INC_CARRYING))
    cost = modMoney.ToDec(invs(ir, repoInvest.INC_AMOUNT))
    If nv > cost Then nv = cost
    Set lines = New Collection
    If nv < carry Then
        diff = carry - nv
        svcLedger.AddDr lines, ho, modConstants.GL_INVEST_LOSS, diff
        svcLedger.AddCr lines, ho, modConstants.GL_INVESTMENTS, diff
    ElseIf nv > carry Then
        diff = nv - carry
        svcLedger.AddDr lines, ho, modConstants.GL_INVESTMENTS, diff
        svcLedger.AddCr lines, ho, modConstants.GL_INVEST_LOSS, diff
    Else
        Set ValueInvestment = modMessages.FailResult("The carrying value is already " & modMoney.FormatAmount(carry) & "."): Exit Function
    End If
    method = StrConv(Trim$(CStr(inp(modConstants.INP_NOTE))), vbProperCase)
    If method <> "Cost" And method <> "Appraisal" And method <> "Multiple" Then method = "Appraisal"
    grp = uow.NewGroupId()
    Set jr = svcLedger.PostJournal(uow, modConstants.JT_INV_VALUATION, "Valuation " & CStr(invs(ir, 1)), grp, lines)
    If Not jr.Success Then Set ValueInvestment = jr: Exit Function
    invs(ir, repoInvest.INC_CARRYING) = nv
    uow.ReplaceTable modConstants.TBL_INVESTMENTS, invs
    uow.AppendRow modConstants.TBL_INVVALUATIONS, Array(uow.NewId(modConstants.TBL_INVVALUATIONS, "ValuationID", "IW", 6), CStr(invs(ir, 1)), uow.BusinessDate, nv, method, "", jr.RefID)
    repoAudit.Add uow, "InvestmentValued", CStr(invs(ir, 1)), "Carrying " & modMoney.FormatAmount(carry) & " to " & modMoney.FormatAmount(nv)
    uow.Commit
    Set ValueInvestment = modMessages.OkResult("Investment " & CStr(invs(ir, 1)) & " valued at " & modMoney.FormatAmount(nv) & " MMK (" & IIf(nv < carry, "impairment ", "write-up ") & modMoney.FormatAmount(diff) & "). Group " & grp & ".", 1, grp)
End Function

' Control: C45 investment, C7 proceeds. HeadOffice requests, another HeadOffice approves.
Public Function DisposeInvestment(ByVal user As clsUser, ByVal inp As Object) As clsResult
    Set DisposeInvestment = RequestChange(user, inp, modConstants.REQ_DISPOSE_INVESTMENT, True)
End Function

Public Function WriteOffInvestment(ByVal user As clsUser, ByVal inp As Object) As clsResult
    Set WriteOffInvestment = RequestChange(user, inp, modConstants.REQ_WRITEOFF_INVESTMENT, False)
End Function

Private Function RequestChange(ByVal user As clsUser, ByVal inp As Object, ByVal requestType As String, ByVal needsProceeds As Boolean) As clsResult
    Dim uow As clsUnitOfWork, msg As String, amt As Variant, invs As Variant, ir As Long, args As Object
    msg = svcSecurity.Denied(user, svcSecurity.ACT_INVEST)
    If Len(msg) = 0 Then msg = NeedTables()
    If Len(msg) > 0 Then Set RequestChange = modMessages.FailResult(msg): Exit Function
    amt = CDec(0)
    If needsProceeds Then
        If Not modArgs.ParseAmount(inp(modConstants.INP_AMOUNT), amt, msg, True) Then Set RequestChange = modMessages.FailResult(msg): Exit Function
    End If
    Set uow = New clsUnitOfWork
    uow.Init repoConfig.BusinessDate(), user.UserID
    invs = uow.WorkingTable(modConstants.TBL_INVESTMENTS)
    ir = repoInvest.FindInvestById(invs, CStr(inp(modConstants.INP_INVESTMENT)))
    If ir = 0 Then Set RequestChange = modMessages.FailResult("Unknown investment (Control C45)."): Exit Function
    If CStr(invs(ir, repoInvest.INC_STATUS)) <> "Active" And CStr(invs(ir, repoInvest.INC_STATUS)) <> "Matured" Then Set RequestChange = modMessages.FailResult("The investment is " & CStr(invs(ir, repoInvest.INC_STATUS)) & "."): Exit Function
    Set args = modArgs.NewArgs()
    args("investment") = CStr(invs(ir, 1))
    args("amount") = amt
    Set RequestChange = svcApproval.Request(uow, requestType, "", "", IIf(needsProceeds, amt, modMoney.ToDec(invs(ir, repoInvest.INC_CARRYING))), modConstants.HEAD_OFFICE_ID, args, user.UserID)
    uow.Commit
End Function

' HO: Dr 1100 proceeds; Cr 1600 carrying value; gain Cr 4400, loss Dr 5310. Unpaid recognized returns stay in 1610.
Public Function ExecDisposeInvestment(ByVal uow As clsUnitOfWork, ByVal args As Object, ByVal createdBy As String, ByVal approvedBy As String) As clsResult
    Dim invs As Variant, ir As Long, carry As Variant, proceeds As Variant, lines As Collection, jr As clsResult, grp As String, ho As String, st As String
    ho = modConstants.HEAD_OFFICE_ID
    invs = uow.WorkingTable(modConstants.TBL_INVESTMENTS)
    ir = repoInvest.FindInvestById(invs, modArgs.Text(args, "investment"))
    If ir = 0 Then Set ExecDisposeInvestment = modMessages.FailResult("Unknown investment."): Exit Function
    st = CStr(invs(ir, repoInvest.INC_STATUS))
    If st <> "Active" And st <> "Matured" Then Set ExecDisposeInvestment = modMessages.FailResult("The investment is " & st & "."): Exit Function
    carry = modMoney.ToDec(invs(ir, repoInvest.INC_CARRYING))
    proceeds = modArgs.Dec(args, "amount")
    Set lines = New Collection
    svcLedger.AddDr lines, ho, modConstants.GL_TREASURY, proceeds
    svcLedger.AddCr lines, ho, modConstants.GL_INVESTMENTS, carry
    If proceeds > carry Then svcLedger.AddCr lines, ho, modConstants.GL_GAIN_DISPOSAL, proceeds - carry
    If proceeds < carry Then svcLedger.AddDr lines, ho, modConstants.GL_LOSS_DISPOSAL, carry - proceeds
    grp = uow.NewGroupId()
    Set jr = svcLedger.PostJournal(uow, modConstants.JT_INV_DISPOSAL, "Disposal " & CStr(invs(ir, 1)), grp, lines)
    If Not jr.Success Then Set ExecDisposeInvestment = jr: Exit Function
    invs(ir, repoInvest.INC_STATUS) = "Disposed"
    invs(ir, repoInvest.INC_CARRYING) = CDec(0)
    uow.ReplaceTable modConstants.TBL_INVESTMENTS, invs
    repoAudit.Add uow, "InvestmentDisposed", CStr(invs(ir, 1)), "Proceeds " & modMoney.FormatAmount(proceeds) & ", approved by " & approvedBy
    Set ExecDisposeInvestment = modMessages.OkResult("Investment " & CStr(invs(ir, 1)) & " disposed for " & modMoney.FormatAmount(proceeds) & " MMK (carrying " & modMoney.FormatAmount(carry) & "). Group " & grp & ".", 1, grp)
End Function

' HO: Dr 5330, Cr 1600 (carrying value) and Dr 5330, Cr 1610 (unpaid recognized returns).
Public Function ExecWriteOffInvestment(ByVal uow As clsUnitOfWork, ByVal args As Object, ByVal createdBy As String, ByVal approvedBy As String) As clsResult
    Dim invs As Variant, ir As Long, carry As Variant, unpaid As Variant, lines As Collection, jr As clsResult, grp As String, ho As String, st As String
    ho = modConstants.HEAD_OFFICE_ID
    invs = uow.WorkingTable(modConstants.TBL_INVESTMENTS)
    ir = repoInvest.FindInvestById(invs, modArgs.Text(args, "investment"))
    If ir = 0 Then Set ExecWriteOffInvestment = modMessages.FailResult("Unknown investment."): Exit Function
    st = CStr(invs(ir, repoInvest.INC_STATUS))
    If st <> "Active" And st <> "Matured" Then Set ExecWriteOffInvestment = modMessages.FailResult("The investment is " & st & "."): Exit Function
    carry = modMoney.ToDec(invs(ir, repoInvest.INC_CARRYING))
    unpaid = repoInvest.UnpaidRecognized(uow.WorkingTable(modConstants.TBL_INVRETURNS), CStr(invs(ir, 1)))
    If unpaid < 0 Then unpaid = CDec(0)
    Set lines = New Collection
    svcLedger.AddDr lines, ho, modConstants.GL_INVEST_LOSS, carry + unpaid
    svcLedger.AddCr lines, ho, modConstants.GL_INVESTMENTS, carry
    svcLedger.AddCr lines, ho, modConstants.GL_INVEST_RECEIVABLE, unpaid
    grp = uow.NewGroupId()
    Set jr = svcLedger.PostJournal(uow, modConstants.JT_INV_WRITEOFF, "Write-off " & CStr(invs(ir, 1)), grp, lines)
    If Not jr.Success Then Set ExecWriteOffInvestment = jr: Exit Function
    invs(ir, repoInvest.INC_STATUS) = "WrittenOff"
    invs(ir, repoInvest.INC_CARRYING) = CDec(0)
    invs(ir, repoInvest.INC_ACCRUED) = CDec(0)
    uow.ReplaceTable modConstants.TBL_INVESTMENTS, invs
    If unpaid > 0 Then uow.AppendRow modConstants.TBL_INVRETURNS, repoInvest.ReturnRow(uow.NewId(modConstants.TBL_INVRETURNS, "ReturnID", "IR", 6), CStr(invs(ir, 1)), uow.BusinessDate, CDec(0), "Interest", jr.RefID, "WriteOff;1610=" & CStr(unpaid))
    repoAudit.Add uow, "InvestmentWrittenOff", CStr(invs(ir, 1)), "Loss " & modMoney.FormatAmount(carry + unpaid) & ", approved by " & approvedBy
    Set ExecWriteOffInvestment = modMessages.OkResult("Investment " & CStr(invs(ir, 1)) & " written off: " & modMoney.FormatAmount(carry + unpaid) & " MMK to expense 5330. Group " & grp & ".", 1, grp)
End Function

' ---- End of day ----

' Maturity and simulated repayment (step 4), daily memo accrual of fixed returns, and overdue-return flags.
Public Function InvestDaily(ByVal uow As clsUnitOfWork, ByVal d As Date, ByVal stats As Object) As clsResult
    Dim invs As Variant, i As Long, mat As Date, rets As Variant, unpaid As Variant, carry As Variant, accrued As Variant, lines As Collection, jr As clsResult, grp As String, ho As String, typ As String
    Dim last As Date, days As Long, upto As Date, total As Variant, nAcc As Long, rest As Variant
    ho = modConstants.HEAD_OFFICE_ID
    If Not modSheetUtils.TableExists(modConstants.TBL_INVESTMENTS) Then Set InvestDaily = modMessages.OkResult("", 0): Exit Function
    invs = uow.WorkingTable(modConstants.TBL_INVESTMENTS)
    If IsEmpty(invs) Then Set InvestDaily = modMessages.OkResult("", 0): Exit Function
    rets = uow.WorkingTable(modConstants.TBL_INVRETURNS)
    total = CDec(0)
    For i = 1 To UBound(invs, 1)
        If CStr(invs(i, repoInvest.INC_STATUS)) = "Active" Then
            typ = CStr(invs(i, repoInvest.INC_TYPE))
            If (typ = "FixedReturnNote" Or typ = "TradeFinance") And modDateUtils.TryToDate(invs(i, repoInvest.INC_MATURITY), mat) Then
                If mat <= d Then
                    ' Maturity: accrual of the days before maturity is in AccruedReturn already.
                    invs(i, repoInvest.INC_STATUS) = "Matured"
                    If repoConfig.GetBool(modConstants.CFG_INVEST_AUTOREPAY) Then
                        carry = modMoney.ToDec(invs(i, repoInvest.INC_CARRYING))
                        unpaid = repoInvest.UnpaidRecognized(rets, CStr(invs(i, 1)))
                        If unpaid < 0 Then unpaid = CDec(0)
                        rest = modMoney.RoundCcy(modMoney.ToDec(invs(i, repoInvest.INC_ACCRUED)), modConstants.CCY_MMK)
                        Set lines = New Collection
                        svcLedger.AddDr lines, ho, modConstants.GL_TREASURY, carry + unpaid + rest
                        svcLedger.AddCr lines, ho, modConstants.GL_INVESTMENTS, carry
                        svcLedger.AddCr lines, ho, modConstants.GL_INVEST_RECEIVABLE, unpaid
                        svcLedger.AddCr lines, ho, modConstants.GL_INVEST_INCOME, rest
                        grp = uow.NewGroupId()
                        Set jr = svcLedger.PostJournal(uow, modConstants.JT_INV_REPAY, "Investment repaid " & CStr(invs(i, 1)), grp, lines)
                        If Not jr.Success Then Set InvestDaily = jr: Exit Function
                        uow.AppendRow modConstants.TBL_INVRETURNS, repoInvest.ReturnRow(uow.NewId(modConstants.TBL_INVRETURNS, "ReturnID", "IR", 6), CStr(invs(i, 1)), d, unpaid + rest, "Interest", jr.RefID, "Received;1610=" & CStr(unpaid))
                        uow.AppendRow modConstants.TBL_INVRETURNS, repoInvest.ReturnRow(uow.NewId(modConstants.TBL_INVRETURNS, "ReturnID", "IR", 6), CStr(invs(i, 1)), d, carry, "Principal", jr.RefID, "Principal repaid")
                        invs(i, repoInvest.INC_CARRYING) = CDec(0)
                        invs(i, repoInvest.INC_ACCRUED) = CDec(0)
                    Else
                        clsLogger.Warn "svcInvest.InvestDaily", "Investment " & CStr(invs(i, 1)) & " matured and is waiting for repayment."
                    End If
                Else
                    last = d - 1
                    If modDateUtils.TryToDate(invs(i, repoInvest.INC_LASTACCRUAL), last) Then
                        upto = d
                        If upto > mat - 1 Then upto = mat - 1
                        days = modDateUtils.DaysBetween(last, upto)
                        If days > 0 Then
                            accrued = modMoney.RoundAccrual(modMoney.ToDec(invs(i, repoInvest.INC_AMOUNT)) * modMoney.ToDec(invs(i, repoInvest.INC_RATE)) / repoConfig.GetDec(modConstants.CFG_DAYCOUNT) * days)
                            invs(i, repoInvest.INC_ACCRUED) = modMoney.ToDec(invs(i, repoInvest.INC_ACCRUED)) + accrued
                            invs(i, repoInvest.INC_LASTACCRUAL) = upto
                            total = total + accrued
                            nAcc = nAcc + 1
                        End If
                    End If
                End If
            End If
        End If
    Next i
    uow.ReplaceTable modConstants.TBL_INVESTMENTS, invs
    stats("InvestmentAccrued") = total
    Set InvestDaily = modMessages.OkResult("", nAcc)
End Function

' Month end: post the accrued return (HO: Dr 1610, Cr 4300) for every Active note, once.
Public Function InvestMonthEnd(ByVal uow As clsUnitOfWork, ByVal d As Date, ByVal stats As Object) As clsResult
    Dim invs As Variant, i As Long, post As Variant, lines As Collection, jr As clsResult, grp As String, ho As String
    ho = modConstants.HEAD_OFFICE_ID
    If Not modSheetUtils.TableExists(modConstants.TBL_INVESTMENTS) Then Set InvestMonthEnd = modMessages.OkResult("", 0): Exit Function
    invs = uow.WorkingTable(modConstants.TBL_INVESTMENTS)
    If IsEmpty(invs) Then Set InvestMonthEnd = modMessages.OkResult("", 0): Exit Function
    For i = 1 To UBound(invs, 1)
        If CStr(invs(i, repoInvest.INC_STATUS)) = "Active" Then
            post = modMoney.RoundCcy(modMoney.ToDec(invs(i, repoInvest.INC_ACCRUED)), modConstants.CCY_MMK)
            If post >= 1 Then
                Set lines = New Collection
                svcLedger.AddDr lines, ho, modConstants.GL_INVEST_RECEIVABLE, post
                svcLedger.AddCr lines, ho, modConstants.GL_INVEST_INCOME, post
                grp = uow.NewGroupId()
                Set jr = svcLedger.PostJournal(uow, modConstants.JT_INV_ACCRUAL, "Return accrual " & CStr(invs(i, 1)) & " " & modDateUtils.PeriodOf(d), grp, lines)
                If Not jr.Success Then Set InvestMonthEnd = jr: Exit Function
                uow.AppendRow modConstants.TBL_INVRETURNS, repoInvest.ReturnRow(uow.NewId(modConstants.TBL_INVRETURNS, "ReturnID", "IR", 6), CStr(invs(i, 1)), d, post, "Interest", jr.RefID, "Recognized")
                invs(i, repoInvest.INC_ACCRUED) = modMoney.ToDec(invs(i, repoInvest.INC_ACCRUED)) - post
            End If
        End If
    Next i
    uow.ReplaceTable modConstants.TBL_INVESTMENTS, invs
    Set InvestMonthEnd = modMessages.OkResult("", 1)
End Function

' ---- Test helper ----

' Admin: draws a random monthly outcome for every Active investment by risk grade (yearly default chance A 0.5%, B 1%, C 3%, D 8%, E 15%, applied monthly as chance / 12).
' A default makes the investee Distressed. Profit share investees get a random reported profit of -20% to +40% of the amount x 5%.
Public Function SimulateInvesteeMonth(ByVal user As clsUser, ByVal inp As Object) As clsResult
    Dim uow As clsUnitOfWork, msg As String, invs As Variant, vees As Variant, i As Long, vr As Long, chance As Double, n As Long, defaults As Long, profit As Variant, note As String
    msg = svcSecurity.Denied(user, svcSecurity.ACT_INVEST_SIM)
    If Len(msg) = 0 Then msg = NeedTables()
    If Len(msg) > 0 Then Set SimulateInvesteeMonth = modMessages.FailResult(msg): Exit Function
    Set uow = New clsUnitOfWork
    uow.Init repoConfig.BusinessDate(), user.UserID
    invs = uow.WorkingTable(modConstants.TBL_INVESTMENTS)
    vees = uow.WorkingTable(modConstants.TBL_INVESTEES)
    If IsEmpty(invs) Then Set SimulateInvesteeMonth = modMessages.OkResult("No investments.", 0): Exit Function
    Randomize
    For i = 1 To UBound(invs, 1)
        If CStr(invs(i, repoInvest.INC_STATUS)) = "Active" Then
            Select Case CStr(invs(i, repoInvest.INC_GRADE))
                Case "A": chance = 0.005
                Case "B": chance = 0.01
                Case "C": chance = 0.03
                Case "D": chance = 0.08
                Case Else: chance = 0.15
            End Select
            n = n + 1
            If Rnd < chance / 12 Then
                vr = repoInvest.FindInvestById(vees, CStr(invs(i, repoInvest.INC_INVESTEE)))
                If vr > 0 Then
                    vees(vr, repoInvest.IVC_STATUS) = "Distressed"
                    vees(vr, 6) = "Simulated default on " & modDateUtils.ToIso(uow.BusinessDate)
                    defaults = defaults + 1
                End If
            ElseIf CStr(invs(i, repoInvest.INC_TYPE)) = "ProfitShare" Or CStr(invs(i, repoInvest.INC_TYPE)) = "Equity" Then
                profit = modMoney.RoundCcy(modMoney.ToDec(invs(i, repoInvest.INC_AMOUNT)) * 0.05 * (Rnd * 0.6 - 0.2), modConstants.CCY_MMK)
                note = note & CStr(invs(i, 1)) & ": reported " & modMoney.FormatAmount(profit) & "; "
            End If
        End If
    Next i
    uow.ReplaceTable modConstants.TBL_INVESTEES, vees
    uow.Commit
    Set SimulateInvesteeMonth = modMessages.OkResult("Simulated one month for " & n & " investments: " & defaults & " investee(s) became Distressed. " & note & "Record profit shares with RunRecordProfitShare.", n)
End Function
