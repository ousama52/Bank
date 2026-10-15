Attribute VB_Name = "svcFxReval"
Option Explicit

' Stage 6 end-of-day steps: daily revaluation at the mid rate of D, the FX position snapshot, and simulated rates for the next day.

' Revalues every (branch, GL, currency) group so that its base balance = RoundCcy(foreign balance x mid of D).
' Asset or liability, the same rule: base needs to go up = Dr the GL, Cr 4200; down = Dr 4200, Cr the GL (always inside one branch).
Public Function RevalueFx(ByVal uow As clsUnitOfWork, ByVal d As Date, ByVal stats As Object) As clsResult
    Dim keys As Variant, k As Variant, parts() As String, v As Variant, mid As Variant, buy As Variant, sell As Variant, msg As String, target As Variant, diff As Variant
    Dim perBranch As Object, b As Variant, lines As Collection, jr As clsResult, grp As String, gain As Variant, ln As clsJournalLine, n As Long
    If Not modSheetUtils.TableExists(modConstants.TBL_EXCHANGERATES) Then Set RevalueFx = modMessages.OkResult("", 0): Exit Function
    Set perBranch = CreateObject("Scripting.Dictionary")
    gain = CDec(0)
    keys = uow.FxGroups().Keys
    For Each k In keys
        parts = Split(CStr(k), "|")
        v = uow.FxGroups()(k)
        If v(0) <> 0 Or v(1) <> 0 Then
            If svcFx.MidFor(uow, parts(2), mid, buy, sell, msg) Then
                target = modMoney.RoundCcy(v(0) * mid, modConstants.CCY_MMK)
                diff = target - v(1)
                If diff <> 0 Then
                    If Not perBranch.Exists(parts(0)) Then Set perBranch(parts(0)) = New Collection
                    Set lines = perBranch(parts(0))
                    If diff > 0 Then
                        lines.Add svcLedger.FxLine(True, parts(0), parts(1), parts(2), 0, mid, diff)
                        lines.Add svcLedger.Cr(parts(0), modConstants.GL_FX_INCOME, diff)
                    Else
                        lines.Add svcLedger.FxLine(False, parts(0), parts(1), parts(2), 0, mid, -diff)
                        lines.Add svcLedger.Dr(parts(0), modConstants.GL_FX_INCOME, -diff)
                    End If
                    gain = gain + diff
                    n = n + 1
                End If
            Else
                clsLogger.Warn "svcFxReval.RevalueFx", "No rate for " & parts(2) & " on " & modDateUtils.ToIso(d) & ": group " & CStr(k) & " not revalued."
            End If
        End If
    Next k
    For Each b In perBranch.Keys
        grp = uow.NewGroupId()
        Set jr = svcLedger.PostJournal(uow, modConstants.JT_FX_REVAL, "FX revaluation " & modDateUtils.ToIso(d) & " " & CStr(b), grp, perBranch(b))
        If Not jr.Success Then Set RevalueFx = jr: Exit Function
    Next b
    stats("FxRevaluation") = gain
    WritePositions uow, d
    Set RevalueFx = modMessages.OkResult("", n)
End Function

' One tblFxPosition row per currency for the business day: holdings, liabilities, net open position and its limit.
Private Sub WritePositions(ByVal uow As clsUnitOfWork, ByVal d As Date)
    Dim cur As Variant, i As Long, ccy As String, holdings As Variant, liab As Variant, mid As Variant, buy As Variant, sell As Variant, msg As String, cap As Variant, lim As Variant, openB As Variant, st As String
    If Not modSheetUtils.TableExists(modConstants.TBL_FXPOSITION) Or Not modSheetUtils.TableExists(modConstants.TBL_CURRENCIES) Then Exit Sub
    cur = modSheetUtils.ReadTable(modConstants.TBL_CURRENCIES)
    If IsEmpty(cur) Then Exit Sub
    cap = -(uow.GLNetAll(modConstants.GL_CAPITAL) + uow.GLNetAll(modConstants.GL_RETAINED))
    lim = modMoney.RoundCcy(repoConfig.GetDec(modConstants.CFG_FX_POSITION_PCT) * cap, modConstants.CCY_MMK)
    For i = 1 To UBound(cur, 1)
        ccy = CStr(cur(i, 1))
        If ccy <> modConstants.CCY_MMK And Len(ccy) > 0 Then
            If svcFx.MidFor(uow, ccy, mid, buy, sell, msg) Then
                svcFx.FxPosition uow, ccy, holdings, liab
                openB = modMoney.RoundCcy((holdings - liab) * mid, modConstants.CCY_MMK)
                If Abs(openB) > lim Then st = "Breach" Else st = "Within"
                uow.AppendRow modConstants.TBL_FXPOSITION, Array(d, ccy, holdings, liab, holdings - liab, mid, openB, lim, st)
                If st = "Breach" Then clsLogger.Warn "svcFxReval.WritePositions", "FX position breach: " & ccy & " net open " & modMoney.FormatAmount(openB) & " MMK, limit " & modMoney.FormatAmount(lim) & "."
            End If
        End If
    Next i
End Sub

' If AutoRates = Yes and there is no rate for D + 1: Mid(D+1) = Mid(D) x (1 + random between -RateMoveMax and +RateMoveMax). Uses Rnd (Randomize was called once by the end of day).
Public Function AutoRates(ByVal uow As clsUnitOfWork, ByVal d As Date) As clsResult
    Dim cur As Variant, i As Long, ccy As String, mid As Variant, buy As Variant, sell As Variant, m2 As Variant, b2 As Variant, s2 As Variant, nm As Variant, move As Double, sp As Variant, n As Long, msg As String
    Dim saveDate As Date
    If Not repoConfig.GetBool(modConstants.CFG_AUTO_RATES) Then Set AutoRates = modMessages.OkResult("", 0): Exit Function
    If Not modSheetUtils.TableExists(modConstants.TBL_EXCHANGERATES) Or Not modSheetUtils.TableExists(modConstants.TBL_CURRENCIES) Then Set AutoRates = modMessages.OkResult("", 0): Exit Function
    cur = modSheetUtils.ReadTable(modConstants.TBL_CURRENCIES)
    If IsEmpty(cur) Then Set AutoRates = modMessages.OkResult("", 0): Exit Function
    sp = repoConfig.GetDec(modConstants.CFG_FX_SPREAD)
    For i = 1 To UBound(cur, 1)
        ccy = CStr(cur(i, 1))
        If ccy <> modConstants.CCY_MMK And Len(ccy) > 0 Then
            If repoFxRate.OnDate(uow.WorkingTable(modConstants.TBL_EXCHANGERATES), ccy, d, mid, buy, sell) Then
                If Not repoFxRate.OnDate(uow.WorkingTable(modConstants.TBL_EXCHANGERATES), ccy, d + 1, m2, b2, s2) Then
                    move = (Rnd * 2 - 1) * CDbl(repoConfig.GetDec(modConstants.CFG_RATE_MOVE))
                    nm = modMoney.RoundTo(mid * (1 + CDec(move)), 2)
                    uow.AppendRow modConstants.TBL_EXCHANGERATES, Array(d + 1, ccy, modMoney.RoundTo(nm * (1 - sp), 2), modMoney.RoundTo(nm * (1 + sp), 2), nm)
                    n = n + 1
                End If
            End If
        End If
    Next i
    Set AutoRates = modMessages.OkResult("", n)
End Function
