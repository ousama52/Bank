Attribute VB_Name = "svcChecksFx"
Option Explicit

' Control checks 19 to 23 (Stage 6). Called by svcChecksExtra. Skipped when the Stage 6 tables are missing.

' Set to True by the end of day while it runs the checks: check 20 (revaluation) is strict only then, because rates change between end-of-day runs.
Public StrictEod As Boolean

Private Function Verdict(ByVal ok As Boolean) As String
    If ok Then Verdict = "PASS" Else Verdict = "FAIL"
End Function

Private Function Cmp(ByVal a As Variant, ByVal b As Variant) As String
    Cmp = modMoney.FormatAmount(a) & " / " & modMoney.FormatAmount(b)
End Function

Public Sub RunFxChecks(ByVal rowsList As Collection, ByVal lns As Variant, ByVal gl As Object)
    Dim fx As Object, accts As Object, k As Variant, parts() As String, a As clsAccount, bal As Object, ccy As String, v As Variant
    Dim bad19 As Long, bad20 As Long, bad21 As Long, bad23 As Long, d As Date, mid As Variant, buy As Variant, sell As Variant, target As Variant, i As Long, hdr As Variant, types As Object
    Dim info19 As String, sample20 As String, cnt As Long
    If Not modSheetUtils.TableExists(modConstants.TBL_EXCHANGERATES) Or Not modSheetUtils.TableExists(modConstants.TBL_FXPOSITION) Then Exit Sub
    d = repoConfig.BusinessDate()
    Set fx = repoJournal.LoadFxNet()
    Set accts = repoAccount.LoadAccounts()
    Set bal = CreateObject("Scripting.Dictionary")
    ' --- 19: per currency, balances of foreign deposit accounts = foreign deposit lines ---
    For Each k In accts.Keys
        Set a = accts(k)
        If a.Ccy <> modConstants.CCY_MMK And a.ProductType <> modConstants.PT_WALLET Then
            If bal.Exists(a.Ccy) Then bal(a.Ccy) = bal(a.Ccy) + a.Balance Else bal(a.Ccy) = a.Balance
        End If
    Next k
    Dim lines19 As Object
    Set lines19 = CreateObject("Scripting.Dictionary")
    For Each k In fx.Keys
        parts = Split(CStr(k), "|")
        If parts(1) = modConstants.GL_SAVINGS Or parts(1) = modConstants.GL_CURRENT Then
            v = fx(k)
            If lines19.Exists(parts(2)) Then lines19(parts(2)) = lines19(parts(2)) - v(0) Else lines19(parts(2)) = -v(0)
        End If
    Next k
    For Each k In lines19.Keys
        If Not bal.Exists(CStr(k)) Then bal(CStr(k)) = CDec(0)
    Next k
    For Each k In bal.Keys
        v = CDec(0)
        If lines19.Exists(CStr(k)) Then v = lines19(CStr(k))
        If v <> bal(k) Then bad19 = bad19 + 1
        info19 = info19 & CStr(k) & ": accounts " & bal(k) & " / lines " & v & "; "
    Next k
    svcChecks.AddRow rowsList, 19, "Foreign accounts: balances = foreign deposit lines per currency", Verdict(bad19 = 0), info19
    ' --- 20: revaluation (base = RoundCcy(foreign x mid of D)) ---
    For Each k In fx.Keys
        parts = Split(CStr(k), "|")
        v = fx(k)
        If v(0) <> 0 Or v(1) <> 0 Then
            If repoFxRate.OnDate(modSheetUtils.ReadTable(modConstants.TBL_EXCHANGERATES), parts(2), d, mid, buy, sell) Then
                target = modMoney.RoundCcy(v(0) * mid, modConstants.CCY_MMK)
                cnt = cnt + 1
                If target <> v(1) Then
                    bad20 = bad20 + 1
                    If Len(sample20) < 60 Then sample20 = sample20 & CStr(k) & " "
                End If
            End If
        End If
    Next k
    If bad20 = 0 Then
        svcChecks.AddRow rowsList, 20, "After revaluation every foreign group has base = foreign x mid", "PASS", cnt & " groups checked"
    ElseIf StrictEod Then
        svcChecks.AddRow rowsList, 20, "After revaluation every foreign group has base = foreign x mid", "FAIL", bad20 & " groups differ: " & sample20
    Else
        svcChecks.AddRow rowsList, 20, "After revaluation every foreign group has base = foreign x mid", "WARN", bad20 & " groups differ until the next end of day revaluation: " & sample20
    End If
    ' --- 21: foreign tags only on allowed GLs, with ForeignAmount and FxRate ---
    Set types = CreateObject("Scripting.Dictionary")
    hdr = repoJournal.LoadHeaders()
    If Not IsEmpty(hdr) Then
        For i = 1 To UBound(hdr, 1)
            types(CStr(hdr(i, repoJournal.JNC_ID))) = CStr(hdr(i, repoJournal.JNC_TYPE))
        Next i
    End If
    If Not IsEmpty(lns) Then
        For i = 1 To UBound(lns, 1)
            ccy = CStr(lns(i, 6))
            If Len(ccy) > 0 And ccy <> modConstants.CCY_MMK Then
                Select Case CStr(lns(i, repoJournal.LNC_GL))
                    Case modConstants.GL_FX_CASH, modConstants.GL_SAVINGS, modConstants.GL_CURRENT, modConstants.GL_WALLETS, modConstants.GL_WHT
                        If Len(CStr(lns(i, 9))) = 0 Then bad21 = bad21 + 1
                        If Len(CStr(lns(i, 10))) = 0 Then
                            bad21 = bad21 + 1
                        ElseIf modMoney.ToDec(lns(i, 10)) = 0 And types(CStr(lns(i, repoJournal.LNC_JOURNAL))) <> modConstants.JT_FX_REVAL Then
                            bad21 = bad21 + 1
                        End If
                    Case Else
                        bad21 = bad21 + 1
                End Select
            End If
        Next i
    End If
    svcChecks.AddRow rowsList, 21, "Foreign tags only on GL 1700, 2000, 2010, 2030, 2200 with ForeignAmount and FxRate", Verdict(bad21 = 0), bad21 & " problems"
    ' --- 23: no negative foreign cash per branch and currency ---
    For Each k In fx.Keys
        parts = Split(CStr(k), "|")
        If parts(1) = modConstants.GL_FX_CASH Then
            v = fx(k)
            If v(0) < 0 Then bad23 = bad23 + 1
        End If
    Next k
    For Each k In accts.Keys
        Set a = accts(k)
        If a.ProductType = modConstants.PT_WALLET And a.Balance < 0 Then bad23 = bad23 + 1
    Next k
    svcChecks.AddRow rowsList, 23, "No branch has negative foreign cash; no wallet is negative", Verdict(bad23 = 0), bad23 & " problems"
    CheckWallets rowsList, accts, gl
End Sub

' --- 22: GL 2030 credit = GL 1110 debit = sum of wallet balances; wallets inside their KYC maximum ---
Private Sub CheckWallets(ByVal rowsList As Collection, ByVal accts As Object, ByVal gl As Object)
    Dim k As Variant, a As clsAccount, total As Variant, bad As Long, limits As Variant, c As clsCustomer, i As Long, maxBal As Variant, agents As Variant, isAgent As Boolean, j As Long
    Dim g2030 As Variant, g1110 As Variant
    If Not modSheetUtils.TableExists(modConstants.TBL_WALLETLIMITS) Or Not modSheetUtils.TableExists(modConstants.TBL_AGENTS) Then Exit Sub
    limits = modSheetUtils.ReadTable(modConstants.TBL_WALLETLIMITS)
    agents = modSheetUtils.ReadTable(modConstants.TBL_AGENTS)
    total = CDec(0)
    For Each k In accts.Keys
        Set a = accts(k)
        If a.ProductType = modConstants.PT_WALLET Then
            total = total + a.Balance
            isAgent = False
            If Not IsEmpty(agents) Then
                For j = 1 To UBound(agents, 1)
                    If CStr(agents(j, repoFx.AGC_WALLET)) = a.AccountNo Then isAgent = True
                Next j
            End If
            If Not isAgent And Not IsEmpty(limits) Then
                Set c = repoCustomer.FindCustomer(a.CustomerID)
                If Not c Is Nothing Then
                    For i = 1 To UBound(limits, 1)
                        If CLng(modMoney.ToDec(limits(i, 1))) = c.KYCLevel Then
                            If a.Balance > modMoney.ToDec(limits(i, 2)) Then bad = bad + 1
                        End If
                    Next i
                End If
            End If
        End If
    Next k
    If gl.Exists(modConstants.GL_WALLETS) Then g2030 = -gl(modConstants.GL_WALLETS) Else g2030 = CDec(0)
    If gl.Exists(modConstants.GL_EMONEY_TRUST) Then g1110 = gl(modConstants.GL_EMONEY_TRUST) Else g1110 = CDec(0)
    svcChecks.AddRow rowsList, 22, "GL 2030 = GL 1110 = sum of wallet balances; no wallet above its KYC maximum", Verdict(g2030 = g1110 And g2030 = total And bad = 0), _
        "2030 " & Cmp(g2030, total) & "; 1110 " & modMoney.FormatAmount(g1110) & "; " & bad & " wallets above maximum"
End Sub
