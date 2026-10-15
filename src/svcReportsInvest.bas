Attribute VB_Name = "svcReportsInvest"
Option Explicit

' Stage 7 reports: InvestmentBook, Concentration, InvestmentIncome, InvestmentAlerts, InvestmentPnL.

Private Const SH_BOOK As String = "InvestmentBook"
Private Const SH_CONC As String = "Concentration"
Private Const SH_INCOME As String = "InvestmentIncome"
Private Const SH_ALERTS As String = "InvestmentAlerts"
Private Const SH_PNL As String = "InvestmentPnL"

Private Sub AddTo(ByVal d As Object, ByVal key As String, ByVal v As Variant)
    If d.Exists(key) Then d(key) = d(key) + v Else d(key) = v
End Sub

Private Sub BuildBook(ByVal d As Date, ByVal invs As Variant, ByVal rets As Variant)
    Dim i As Long, rowsA As Collection
    Set rowsA = New Collection
    If Not IsEmpty(invs) Then
        For i = 1 To UBound(invs, 1)
            rowsA.Add Array(CStr(invs(i, 1)), CStr(invs(i, repoInvest.INC_INVESTEE)), CStr(invs(i, repoInvest.INC_TYPE)), CStr(invs(i, repoInvest.INC_SECTOR)), CStr(invs(i, repoInvest.INC_GRADE)), _
                CStr(invs(i, repoInvest.INC_STATUS)), modMoney.ToDec(invs(i, repoInvest.INC_AMOUNT)), modMoney.ToDec(invs(i, repoInvest.INC_CARRYING)), modMoney.ToDec(invs(i, repoInvest.INC_ACCRUED)), _
                repoInvest.UnpaidRecognized(rets, CStr(invs(i, 1))))
        Next i
    End If
    modSheetUtils.PrepareReport SH_BOOK, "1,2"
    modSheetUtils.SetCell SH_BOOK, 1, 1, "Investment book " & modDateUtils.ToIso(d) & " (SIMULATION)", True
    modSheetUtils.WriteGrid SH_BOOK, 3, 1, svcReports.ToGrid(Array("Investment", "Investee", "Type", "Sector", "Risk", "Status", "Cost", "Carrying value", "Accrued (memo)", "Unpaid recognized return"), rowsA)
    modSheetUtils.StyleReport SH_BOOK, 3, 10, "7,8,9,10"
End Sub

Private Sub BuildConcentration(ByVal d As Date, ByVal invs As Variant)
    Dim lim As Variant, i As Long, rowsA As Collection, cap As Variant, used As Variant, j As Long, key As String, typ As String, pct As Variant, st As String, g As String
    Dim net As Object, k As Variant
    Set net = repoJournal.LoadGLNet()
    cap = CDec(0)
    For Each k In net.Keys
        If Mid$(CStr(k), InStr(CStr(k), "|") + 1) = modConstants.GL_CAPITAL Or Mid$(CStr(k), InStr(CStr(k), "|") + 1) = modConstants.GL_RETAINED Then cap = cap - net(k)
    Next k
    lim = modSheetUtils.ReadTable(modConstants.TBL_INVLIMITS)
    Set rowsA = New Collection
    If Not IsEmpty(lim) Then
        For i = 1 To UBound(lim, 1)
            typ = CStr(lim(i, 2))
            key = CStr(lim(i, 3))
            pct = modMoney.ToDec(lim(i, 4))
            used = CDec(0)
            If Not IsEmpty(invs) Then
                For j = 1 To UBound(invs, 1)
                    st = CStr(invs(j, repoInvest.INC_STATUS))
                    g = CStr(invs(j, repoInvest.INC_GRADE))
                    If st = "Active" Or st = "Matured" Then
                        If typ = "Total" Or (typ = "Sector" And CStr(invs(j, repoInvest.INC_SECTOR)) = key) Or (typ = "HighRisk" And (g = "D" Or g = "E")) Then used = used + modMoney.ToDec(invs(j, repoInvest.INC_CARRYING))
                    End If
                Next j
            End If
            rowsA.Add Array(typ, key, pct, pct * cap, used, pct * cap - used, IIf(used > pct * cap, "BREACH", "Within"))
        Next i
    End If
    modSheetUtils.PrepareReport SH_CONC, "1,2"
    modSheetUtils.SetCell SH_CONC, 1, 1, "Concentration limits (capital " & modMoney.FormatAmount(cap) & " MMK) " & modDateUtils.ToIso(d) & " (SIMULATION)", True
    modSheetUtils.WriteGrid SH_CONC, 3, 1, svcReports.ToGrid(Array("Limit", "Key", "Max % of capital", "Max amount", "Used", "Available", "Status"), rowsA)
    modSheetUtils.StyleReport SH_CONC, 3, 7, "4,5,6"
    modSheetUtils.GetSheet(SH_CONC).Range("C4:C100").NumberFormat = "0.0%"
End Sub

Private Sub BuildIncome(ByVal d As Date, ByVal rets As Variant, ByVal invs As Variant)
    Dim i As Long, per As Object, key As String, k As Variant, parts() As String, rowsA As Collection, dt As Date, ir As Long, typ As String
    Set per = CreateObject("Scripting.Dictionary")
    If Not IsEmpty(rets) Then
        For i = 1 To UBound(rets, 1)
            If Left$(CStr(rets(i, repoInvest.IRC_NOTE)), 10) = "Recognized" And modDateUtils.TryToDate(rets(i, repoInvest.IRC_DATE), dt) Then
                ir = repoInvest.FindInvestById(invs, CStr(rets(i, repoInvest.IRC_INVESTMENT)))
                typ = "?"
                If ir > 0 Then typ = CStr(invs(ir, repoInvest.INC_TYPE))
                key = modDateUtils.PeriodOf(dt) & "|" & typ
                AddTo per, key, modMoney.ToDec(rets(i, repoInvest.IRC_AMOUNT))
            End If
        Next i
    End If
    Set rowsA = New Collection
    For Each k In per.Keys
        parts = Split(CStr(k), "|")
        rowsA.Add Array(parts(0), parts(1), per(k))
    Next k
    modSheetUtils.PrepareReport SH_INCOME, "1,2"
    modSheetUtils.SetCell SH_INCOME, 1, 1, "Investment income recognized by month and type " & modDateUtils.ToIso(d) & " (SIMULATION)", True
    modSheetUtils.WriteGrid SH_INCOME, 3, 1, svcReports.ToGrid(Array("Period", "Type", "Income"), rowsA)
    modSheetUtils.StyleReport SH_INCOME, 3, 3, "3"
End Sub

Private Sub BuildAlerts(ByVal d As Date, ByVal invs As Variant, ByVal rets As Variant, ByVal vees As Variant)
    Dim i As Long, rowsA As Collection, mat As Date, unpaid As Variant, lastRecv As Date, j As Long, rd As Date
    Set rowsA = New Collection
    If Not IsEmpty(invs) Then
        For i = 1 To UBound(invs, 1)
            If CStr(invs(i, repoInvest.INC_STATUS)) = "Matured" Then
                If modDateUtils.TryToDate(invs(i, repoInvest.INC_MATURITY), mat) Then
                    If d - mat >= 30 Then rowsA.Add Array(CStr(invs(i, 1)), "Matured and unpaid for " & (d - mat) & " days")
                End If
            End If
            If CStr(invs(i, repoInvest.INC_STATUS)) = "Active" Then
                unpaid = repoInvest.UnpaidRecognized(rets, CStr(invs(i, 1)))
                If unpaid > 0 Then
                    lastRecv = 0
                    If Not IsEmpty(rets) Then
                        For j = 1 To UBound(rets, 1)
                            If CStr(rets(j, repoInvest.IRC_INVESTMENT)) = CStr(invs(i, 1)) And Left$(CStr(rets(j, repoInvest.IRC_NOTE)), 8) = "Received" And modDateUtils.TryToDate(rets(j, repoInvest.IRC_DATE), rd) Then
                                If rd > lastRecv Then lastRecv = rd
                            End If
                        Next j
                    End If
                    If lastRecv = 0 And modDateUtils.TryToDate(invs(i, repoInvest.INC_START), rd) Then lastRecv = rd
                    If d - lastRecv >= 90 Then rowsA.Add Array(CStr(invs(i, 1)), "No return received for " & (d - lastRecv) & " days (unpaid " & modMoney.FormatAmount(unpaid) & ")")
                End If
            End If
        Next i
    End If
    If Not IsEmpty(vees) Then
        For i = 1 To UBound(vees, 1)
            If CStr(vees(i, repoInvest.IVC_STATUS)) = "Distressed" Then rowsA.Add Array(CStr(vees(i, 1)), "Investee " & CStr(vees(i, repoInvest.IVC_NAME)) & " is Distressed")
        Next i
    End If
    modSheetUtils.PrepareReport SH_ALERTS, "1"
    modSheetUtils.SetCell SH_ALERTS, 1, 1, "Investment alerts " & modDateUtils.ToIso(d) & " (SIMULATION)", True
    modSheetUtils.WriteGrid SH_ALERTS, 3, 1, svcReports.ToGrid(Array("Item", "Alert"), rowsA)
    modSheetUtils.StyleReport SH_ALERTS, 3, 2, ""
End Sub

Private Sub BuildPnL(ByVal d As Date, ByVal invs As Variant, ByVal rets As Variant, ByVal vals As Variant)
    Dim i As Long, rowsA As Collection, income As Variant, j As Long, loss As Variant, id As String
    Set rowsA = New Collection
    If Not IsEmpty(invs) Then
        For i = 1 To UBound(invs, 1)
            id = CStr(invs(i, 1))
            income = CDec(0)
            If Not IsEmpty(rets) Then
                For j = 1 To UBound(rets, 1)
                    If CStr(rets(j, repoInvest.IRC_INVESTMENT)) = id And Left$(CStr(rets(j, repoInvest.IRC_NOTE)), 10) = "Recognized" Then income = income + modMoney.ToDec(rets(j, repoInvest.IRC_AMOUNT))
                    If CStr(rets(j, repoInvest.IRC_INVESTMENT)) = id And Left$(CStr(rets(j, repoInvest.IRC_NOTE)), 8) = "Received" And InStr(CStr(rets(j, repoInvest.IRC_NOTE)), "1610=") > 0 Then
                        income = income + modMoney.ToDec(rets(j, repoInvest.IRC_AMOUNT))
                    End If
                Next j
            End If
            loss = modMoney.ToDec(invs(i, repoInvest.INC_AMOUNT)) - modMoney.ToDec(invs(i, repoInvest.INC_CARRYING))
            rowsA.Add Array(id, CStr(invs(i, repoInvest.INC_STATUS)), income, loss, income - loss)
        Next i
    End If
    modSheetUtils.PrepareReport SH_PNL, "1"
    modSheetUtils.SetCell SH_PNL, 1, 1, "Investment result (income minus cost not carried) " & modDateUtils.ToIso(d) & " (SIMULATION)", True
    modSheetUtils.WriteGrid SH_PNL, 3, 1, svcReports.ToGrid(Array("Investment", "Status", "Income (recognized + cash)", "Cost not carried", "Net"), rowsA)
    modSheetUtils.StyleReport SH_PNL, 3, 5, "3,4,5"
End Sub

Public Sub RunInvestReports(ByVal d As Date, ByVal inp As Object)
    Dim invs As Variant, rets As Variant, vees As Variant, vals As Variant, t As Variant
    If Not modSheetUtils.TableExists(modConstants.TBL_INVESTMENTS) Or Not modSheetUtils.TableExists(modConstants.TBL_INVLIMITS) Then Exit Sub
    For Each t In Array(SH_BOOK, SH_CONC, SH_INCOME, SH_ALERTS, SH_PNL)
        modSheetUtils.EnsureSheet CStr(t)
    Next t
    invs = modSheetUtils.ReadTable(modConstants.TBL_INVESTMENTS)
    rets = modSheetUtils.ReadTable(modConstants.TBL_INVRETURNS)
    vees = modSheetUtils.ReadTable(modConstants.TBL_INVESTEES)
    vals = modSheetUtils.ReadTable(modConstants.TBL_INVVALUATIONS)
    BuildBook d, invs, rets
    BuildConcentration d, invs
    BuildIncome d, rets, invs
    BuildAlerts d, invs, rets, vees
    BuildPnL d, invs, rets, vals
End Sub
