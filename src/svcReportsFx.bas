Attribute VB_Name = "svcReportsFx"
Option Explicit

' Stage 6 reports: FxRates, FxPosition, FxIncome, CurrencyCash, FxDeals, IntlPayments, WalletSummary, MerchantSummary, AgentSummary.

Private Sub AddTo(ByVal d As Object, ByVal key As String, ByVal v As Variant)
    If d.Exists(key) Then d(key) = d(key) + v Else d(key) = v
End Sub

Private Sub Emit(ByVal sheetName As String, ByVal title As String, ByVal headers As Variant, ByVal rowsA As Collection, ByVal textCols As String, ByVal moneyCols As String)
    modSheetUtils.EnsureSheet sheetName
    modSheetUtils.PrepareReport sheetName, textCols
    modSheetUtils.SetCell sheetName, 1, 1, title & " (SIMULATION)", True
    modSheetUtils.WriteGrid sheetName, 3, 1, svcReports.ToGrid(headers, rowsA)
    modSheetUtils.StyleReport sheetName, 3, UBound(headers) - LBound(headers) + 1, moneyCols
End Sub

Private Function IsoOf(ByVal v As Variant) As String
    Dim d As Date
    If modDateUtils.TryToDate(v, d) Then IsoOf = modDateUtils.ToIso(d)
End Function

Private Sub BuildRates(ByVal d As Date)
    Dim data As Variant, i As Long, rowsA As Collection
    Set rowsA = New Collection
    data = modSheetUtils.ReadTable(modConstants.TBL_EXCHANGERATES)
    If Not IsEmpty(data) Then
        For i = UBound(data, 1) To 1 Step -1
            If Len(CStr(data(i, 1))) > 0 Then rowsA.Add Array(IsoOf(data(i, 1)), CStr(data(i, 2)), CStr(data(i, 3)), CStr(data(i, 4)), CStr(data(i, 5)))
        Next i
    End If
    Emit "FxRates", "Exchange rates, latest first " & modDateUtils.ToIso(d), Array("Date", "Currency", "Mid", "Buy", "Sell"), rowsA, "1,2,3,4,5", ""
End Sub

Private Sub BuildPosition(ByVal d As Date)
    Dim data As Variant, i As Long, rowsA As Collection, dt As Date
    Set rowsA = New Collection
    data = modSheetUtils.ReadTable(modConstants.TBL_FXPOSITION)
    If Not IsEmpty(data) Then
        For i = UBound(data, 1) To 1 Step -1
            If modDateUtils.TryToDate(data(i, 1), dt) Then
                If dt >= d - 30 Then rowsA.Add Array(modDateUtils.ToIso(dt), CStr(data(i, 2)), modMoney.ToDec(data(i, 3)), modMoney.ToDec(data(i, 4)), modMoney.ToDec(data(i, 5)), CStr(data(i, 6)), modMoney.ToDec(data(i, 7)), modMoney.ToDec(data(i, 8)), CStr(data(i, 9)))
            End If
        Next i
    End If
    Emit "FxPositionReport", "Foreign currency position, last 30 days " & modDateUtils.ToIso(d), Array("Date", "Currency", "Holdings", "Liabilities", "Net open", "Mid", "Net open (MMK)", "Limit (MMK)", "Status"), rowsA, "1,2,6,9", "3,4,5,7,8"
End Sub

Private Sub BuildIncome(ByVal d As Date, ByVal lines As Variant)
    Dim per As Object, i As Long, dt As Date, k As Variant, rowsA As Collection, gl As String, parts() As String
    Set per = CreateObject("Scripting.Dictionary")
    If Not IsEmpty(lines) Then
        For i = 1 To UBound(lines, 1)
            If Not IsEmpty(lines(i, 1)) Then
                dt = lines(i, 1)
                gl = CStr(lines(i, 3))
                If gl = modConstants.GL_FX_INCOME Then
                    If CStr(lines(i, 6)) = modConstants.JT_FX_REVAL Then
                        AddTo per, modDateUtils.PeriodOf(dt) & "|" & CStr(lines(i, 2)) & "|Revaluation", lines(i, 5) - lines(i, 4)
                    Else
                        AddTo per, modDateUtils.PeriodOf(dt) & "|" & CStr(lines(i, 2)) & "|Spread and deals", lines(i, 5) - lines(i, 4)
                    End If
                End If
            End If
        Next i
    End If
    Set rowsA = New Collection
    For Each k In per.Keys
        parts = Split(CStr(k), "|")
        rowsA.Add Array(parts(0), parts(1), parts(2), per(k))
    Next k
    Emit "FxIncome", "Foreign exchange income (GL 4200) by month and branch " & modDateUtils.ToIso(d), Array("Period", "Branch", "Source", "Income"), rowsA, "1,2,3", "4"
End Sub

Private Sub BuildCash(ByVal d As Date)
    Dim fx As Object, k As Variant, parts() As String, rowsA As Collection
    Set fx = repoJournal.LoadFxNet()
    Set rowsA = New Collection
    For Each k In fx.Keys
        parts = Split(CStr(k), "|")
        If parts(1) = modConstants.GL_FX_CASH And (fx(k)(0) <> 0 Or fx(k)(1) <> 0) Then rowsA.Add Array(parts(0), parts(2), fx(k)(0), fx(k)(1))
    Next k
    Emit "CurrencyCash", "Foreign cash (GL 1700) by branch and currency " & modDateUtils.ToIso(d), Array("Branch", "Currency", "Foreign amount", "Book value (MMK)"), rowsA, "1,2", "3,4"
End Sub

Private Sub BuildDeals(ByVal d As Date)
    Dim data As Variant, i As Long, rowsA As Collection
    Set rowsA = New Collection
    data = modSheetUtils.ReadTable(modConstants.TBL_FXDEALS)
    If Not IsEmpty(data) Then
        For i = UBound(data, 1) To 1 Step -1
            If Len(CStr(data(i, 1))) > 0 Then rowsA.Add Array(CStr(data(i, 1)), IsoOf(data(i, 2)), CStr(data(i, 3)), CStr(data(i, 4)), modMoney.ToDec(data(i, 5)), CStr(data(i, 6)), modMoney.ToDec(data(i, 8)), modMoney.ToDec(data(i, 9)), CStr(data(i, 10)), CStr(data(i, 11)), CStr(data(i, 15)))
        Next i
    End If
    Emit "FxDealsReport", "Currency deals " & modDateUtils.ToIso(d), Array("Deal", "Date", "Type", "Currency", "Foreign", "Rate", "MMK", "Spread", "Branch", "Account", "Status"), rowsA, "1,2,3,4,6,9,10,11", "5,7,8"
End Sub

Private Sub BuildIntl(ByVal d As Date)
    Dim data As Variant, i As Long, rowsA As Collection
    Set rowsA = New Collection
    data = modSheetUtils.ReadTable(modConstants.TBL_INTLPAYMENTS)
    If Not IsEmpty(data) Then
        For i = UBound(data, 1) To 1 Step -1
            If Len(CStr(data(i, repoFx.IPC_ID))) > 0 Then rowsA.Add Array(CStr(data(i, repoFx.IPC_ID)), CStr(data(i, repoFx.IPC_DIR)), CStr(data(i, repoFx.IPC_ACCOUNT)), CStr(data(i, repoFx.IPC_CCY)), modMoney.ToDec(data(i, repoFx.IPC_AMOUNT)), modMoney.ToDec(data(i, repoFx.IPC_MMK)), _
                CStr(data(i, repoFx.IPC_PARTY)), CStr(data(i, repoFx.IPC_BANK)), CStr(data(i, repoFx.IPC_STATUS)), IsoOf(data(i, repoFx.IPC_STATUSDATE)), CStr(data(i, 19)), CStr(data(i, repoFx.IPC_REASON)))
        Next i
    End If
    Emit "IntlPaymentsReport", "International payments " & modDateUtils.ToIso(d), Array("Payment", "Direction", "Account", "Currency", "Amount", "MMK", "Party", "Bank", "Status", "Status date", "Screening", "Reason"), rowsA, "1,2,3,4,7,8,9,10,11,12", "5,6"
End Sub

Private Sub BuildWallets(ByVal d As Date, ByVal uow As clsUnitOfWork)
    Dim k As Variant, a As clsAccount, custs As Object, kyc As Object, cnt As Object, sum As Object, lvl As String, rowsA As Collection, total As Variant, trust As Variant, held As Variant
    Set custs = repoCustomer.LoadCustomers()
    Set cnt = CreateObject("Scripting.Dictionary")
    Set sum = CreateObject("Scripting.Dictionary")
    total = CDec(0)
    For Each k In uow.Accounts().Keys
        Set a = uow.Accounts()(k)
        If a.ProductType = modConstants.PT_WALLET And a.Status <> modConstants.STATUS_CLOSED Then
            lvl = "KYC ?"
            If custs.Exists(a.CustomerID) Then lvl = "KYC " & custs(a.CustomerID).KYCLevel
            AddTo cnt, lvl, 1
            AddTo sum, lvl, a.Balance
            total = total + a.Balance
        End If
    Next k
    Set rowsA = New Collection
    For Each k In cnt.Keys
        rowsA.Add Array(CStr(k), cnt(k), sum(k))
    Next k
    rowsA.Add Array("TOTAL wallets", "", total)
    held = -uow.GLNetAll(modConstants.GL_WALLETS)
    trust = uow.GLNetAll(modConstants.GL_EMONEY_TRUST)
    rowsA.Add Array("GL 2030 wallet liability", "", held)
    rowsA.Add Array("GL 1110 trust account", "", trust)
    rowsA.Add Array("Reconciled (2030 = 1110 = wallets)", "", IIf(held = trust And held = total, "YES", "NO"))
    Emit "WalletSummary", "Wallets by KYC level and trust reconciliation " & modDateUtils.ToIso(d), Array("Level", "Wallets", "Balance"), rowsA, "1,2,3", "3"
End Sub

Private Sub BuildMerchants(ByVal d As Date, ByVal uow As clsUnitOfWork)
    Dim data As Variant, i As Long, tx As Variant, vol As Object, rowsA As Collection, a As clsAccount, wal As String, bal As Variant
    Set vol = CreateObject("Scripting.Dictionary")
    tx = repoTransaction.LoadTransactions()
    If Not IsEmpty(tx) Then
        For i = 1 To UBound(tx, 1)
            If CStr(tx(i, 5)) = modConstants.TXN_WALLET_PAY And CStr(tx(i, 6)) = modConstants.DIR_CREDIT And CStr(tx(i, 14)) <> modConstants.STATUS_REVERSED Then AddTo vol, CStr(tx(i, 4)), modMoney.ToDec(tx(i, 7))
        Next i
    End If
    Set rowsA = New Collection
    data = modSheetUtils.ReadTable(modConstants.TBL_MERCHANTS)
    If Not IsEmpty(data) Then
        For i = 1 To UBound(data, 1)
            If Len(CStr(data(i, 1))) > 0 Then
                wal = CStr(data(i, repoFx.MEC_WALLET))
                bal = CDec(0)
                Set a = uow.GetAccount(wal)
                If Not a Is Nothing Then bal = a.Balance
                rowsA.Add Array(CStr(data(i, 1)), CStr(data(i, 2)), wal, CStr(data(i, repoFx.MEC_STATUS)), IIf(vol.Exists(wal), vol(wal), CDec(0)), bal)
            End If
        Next i
    End If
    Emit "MerchantSummary", "Merchants " & modDateUtils.ToIso(d), Array("Merchant", "Name", "Wallet", "Status", "Payments received", "Wallet balance"), rowsA, "1,2,3,4", "5,6"
End Sub

Private Sub BuildAgents(ByVal d As Date, ByVal uow As clsUnitOfWork)
    Dim data As Variant, i As Long, rowsA As Collection, a As clsAccount, bal As Variant, tx As Variant, comm As Variant
    Set rowsA = New Collection
    data = modSheetUtils.ReadTable(modConstants.TBL_AGENTS)
    If Not IsEmpty(data) Then
        For i = 1 To UBound(data, 1)
            If Len(CStr(data(i, 1))) > 0 Then
                bal = CDec(0)
                Set a = uow.GetAccount(CStr(data(i, repoFx.AGC_WALLET)))
                If Not a Is Nothing Then bal = a.Balance
                rowsA.Add Array(CStr(data(i, 1)), CStr(data(i, 2)), CStr(data(i, repoFx.AGC_BRANCH)), CStr(data(i, repoFx.AGC_WALLET)), bal, modMoney.ToDec(data(i, repoFx.AGC_CASH)), CStr(data(i, repoFx.AGC_STATUS)))
            End If
        Next i
    End If
    Emit "AgentSummary", "Agents " & modDateUtils.ToIso(d), Array("Agent", "Name", "Branch", "Float wallet", "Float balance", "Cash on hand", "Status"), rowsA, "1,2,3,4,7", "5,6"
End Sub

Public Sub RunFxReports(ByVal d As Date, ByVal inp As Object)
    Dim uow As clsUnitOfWork, lines As Variant
    If Not modSheetUtils.TableExists(modConstants.TBL_EXCHANGERATES) Or Not modSheetUtils.TableExists(modConstants.TBL_INTLPAYMENTS) Or Not modSheetUtils.TableExists(modConstants.TBL_AGENTS) Then Exit Sub
    Set uow = New clsUnitOfWork
    uow.Init d, "REPORT"
    lines = repoJournal.LoadDatedLines()
    BuildRates d
    BuildPosition d
    BuildIncome d, lines
    BuildCash d
    BuildDeals d
    BuildIntl d
    BuildWallets d, uow
    BuildMerchants d, uow
    BuildAgents d, uow
End Sub
