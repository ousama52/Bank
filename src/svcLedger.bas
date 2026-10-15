Attribute VB_Name = "svcLedger"
Option Explicit

' The ONLY code that creates Journal and JournalLines rows. Everything goes through PostJournal,
' which refuses unless every ledger rule holds. Rows are added to the unit of work, not to the sheet.

Public Function Dr(ByVal branchId As String, ByVal glCode As String, ByVal amount As Variant, Optional ByVal accountNo As String = "") As clsJournalLine
    Dim ln As clsJournalLine
    Set ln = New clsJournalLine
    ln.BranchID = branchId
    ln.GLCode = glCode
    ln.AccountNo = accountNo
    ln.Ccy = modConstants.CCY_MMK
    ln.Debit = CDec(amount)
    ln.Credit = CDec(0)
    Set Dr = ln
End Function

Public Function Cr(ByVal branchId As String, ByVal glCode As String, ByVal amount As Variant, Optional ByVal accountNo As String = "") As clsJournalLine
    Dim ln As clsJournalLine
    Set ln = New clsJournalLine
    ln.BranchID = branchId
    ln.GLCode = glCode
    ln.AccountNo = accountNo
    ln.Ccy = modConstants.CCY_MMK
    ln.Debit = CDec(0)
    ln.Credit = CDec(amount)
    Set Cr = ln
End Function

' A foreign leg: ForeignAmount in the currency, base = Debit or Credit in MMK (RoundCcy(foreign x rate) unless a base is given).
Public Function FxLine(ByVal isDebit As Boolean, ByVal branchId As String, ByVal glCode As String, ByVal ccy As String, ByVal foreignAmount As Variant, _
                       ByVal rate As Variant, ByVal baseAmount As Variant, Optional ByVal accountNo As String = "") As clsJournalLine
    Dim ln As clsJournalLine
    Set ln = New clsJournalLine
    ln.BranchID = branchId
    ln.GLCode = glCode
    ln.AccountNo = accountNo
    ln.Ccy = ccy
    ln.FxRate = rate
    ln.ForeignAmount = CDec(foreignAmount)
    If isDebit Then
        ln.Debit = CDec(baseAmount)
        ln.Credit = CDec(0)
    Else
        ln.Debit = CDec(0)
        ln.Credit = CDec(baseAmount)
    End If
    Set FxLine = ln
End Function

Public Sub AddDrFx(ByVal lines As Collection, ByVal branchId As String, ByVal glCode As String, ByVal ccy As String, ByVal foreignAmount As Variant, ByVal rate As Variant, _
                   ByVal baseAmount As Variant, Optional ByVal accountNo As String = "")
    If CDec(foreignAmount) <> 0 Or CDec(baseAmount) <> 0 Then lines.Add FxLine(True, branchId, glCode, ccy, foreignAmount, rate, baseAmount, accountNo)
End Sub

Public Sub AddCrFx(ByVal lines As Collection, ByVal branchId As String, ByVal glCode As String, ByVal ccy As String, ByVal foreignAmount As Variant, ByVal rate As Variant, _
                   ByVal baseAmount As Variant, Optional ByVal accountNo As String = "")
    If CDec(foreignAmount) <> 0 Or CDec(baseAmount) <> 0 Then lines.Add FxLine(False, branchId, glCode, ccy, foreignAmount, rate, baseAmount, accountNo)
End Sub

' Adds a line only when the amount is not zero (fees and taxes are often zero).
Public Sub AddDr(ByVal lines As Collection, ByVal branchId As String, ByVal glCode As String, ByVal amount As Variant, Optional ByVal accountNo As String = "")
    If CDec(amount) <> 0 Then lines.Add Dr(branchId, glCode, amount, accountNo)
End Sub

Public Sub AddCr(ByVal lines As Collection, ByVal branchId As String, ByVal glCode As String, ByVal amount As Variant, Optional ByVal accountNo As String = "")
    If CDec(amount) <> 0 Then lines.Add Cr(branchId, glCode, amount, accountNo)
End Sub

' Returns "" when the lines may be posted, else the reason.
Private Function ValidateLines(ByVal uow As clsUnitOfWork, ByVal lines As Collection, ByVal journalType As String) As String
    Dim ln As clsJournalLine, byBranch As Object, totDr As Variant, totCr As Variant, k As Variant
    Dim isDr As Boolean, isCr As Boolean, a As clsAccount, chartRow As Variant
    If lines Is Nothing Then ValidateLines = "no lines": Exit Function
    If lines.Count < 2 Then ValidateLines = "a journal needs at least 2 lines": Exit Function
    If uow.BusinessDate <> repoConfig.BusinessDate() Then ValidateLines = "PostDate must be the business date": Exit Function
    Set byBranch = CreateObject("Scripting.Dictionary")
    totDr = CDec(0)
    totCr = CDec(0)
    For Each ln In lines
        If ln.Ccy <> modConstants.CCY_MMK Then
            Select Case ln.GLCode
                Case modConstants.GL_FX_CASH, modConstants.GL_SAVINGS, modConstants.GL_CURRENT, modConstants.GL_WALLETS, modConstants.GL_WHT
                Case Else
                    ValidateLines = "GL " & ln.GLCode & " cannot carry a foreign currency."
                    Exit Function
            End Select
            If IsEmpty(ln.FxRate) Or IsEmpty(ln.ForeignAmount) Then ValidateLines = "a foreign line needs ForeignAmount and FxRate": Exit Function
            If journalType <> modConstants.JT_FX_REVAL And ln.ForeignAmount <= 0 Then ValidateLines = "a foreign line needs a ForeignAmount above zero": Exit Function
        End If
        If Not uow.Chart().Exists(ln.GLCode) Then ValidateLines = "GL " & ln.GLCode & " does not exist": Exit Function
        chartRow = uow.Chart()(ln.GLCode)
        If Not chartRow(3) Then ValidateLines = "GL " & ln.GLCode & " is not active": Exit Function
        If Not uow.Branches().Exists(ln.BranchID) Then ValidateLines = "branch " & ln.BranchID & " does not exist": Exit Function
        isDr = (ln.Debit > 0)
        isCr = (ln.Credit > 0)
        If ln.Debit < 0 Or ln.Credit < 0 Or isDr = isCr Then ValidateLines = "each line needs a debit or a credit, not both and not zero": Exit Function
        If Len(ln.AccountNo) > 0 Then
            Set a = uow.GetAccount(ln.AccountNo)
            If a Is Nothing Then ValidateLines = "account " & ln.AccountNo & " does not exist": Exit Function
            If a.BranchID <> ln.BranchID Then ValidateLines = "account " & ln.AccountNo & " lines must sit in its home branch " & a.BranchID: Exit Function
        End If
        If Not byBranch.Exists(ln.BranchID) Then byBranch(ln.BranchID) = CDec(0)
        byBranch(ln.BranchID) = byBranch(ln.BranchID) + ln.Debit - ln.Credit
        totDr = totDr + ln.Debit
        totCr = totCr + ln.Credit
    Next ln
    If totDr <> totCr Then ValidateLines = "total debit " & modMoney.FormatAmount(totDr) & " <> total credit " & modMoney.FormatAmount(totCr): Exit Function
    For Each k In byBranch.Keys
        If byBranch(k) <> 0 Then ValidateLines = "branch " & CStr(k) & " does not balance": Exit Function
    Next k
End Function

' Adds the journal and its lines to the unit of work. RefID of the result = JournalID.
Public Function PostJournal(ByVal uow As clsUnitOfWork, ByVal journalType As String, ByVal description As String, _
                            ByVal groupId As String, ByVal lines As Collection, Optional ByVal reversalOf As String = "") As clsResult
    Dim ln As clsJournalLine, msg As String, jid As String
    msg = ValidateLines(uow, lines, journalType)
    If Len(msg) > 0 Then
        Set PostJournal = modMessages.FailResult("Journal refused: " & msg)
        Exit Function
    End If
    jid = uow.NewJournalId()
    uow.AppendRow modConstants.TBL_JOURNAL, repoJournal.HeaderRow(jid, uow.BusinessDate, description, journalType, groupId, reversalOf, uow.UserID)
    For Each ln In lines
        uow.AppendRow modConstants.TBL_JLINES, repoJournal.LineRow(uow.NewLineId(), jid, ln)
        uow.AddGLNet ln.BranchID, ln.GLCode, ln.Debit - ln.Credit
        If ln.Ccy <> modConstants.CCY_MMK Then
            If ln.Debit > 0 Then uow.AddFx ln.BranchID, ln.GLCode, ln.Ccy, ln.ForeignAmount, ln.Debit Else uow.AddFx ln.BranchID, ln.GLCode, ln.Ccy, -ln.ForeignAmount, -ln.Credit
        End If
    Next ln
    uow.LastJournalID = jid
    Set PostJournal = modMessages.OkResult("", lines.Count, jid)
End Function

' ---- Reversal ----

Private Function IsReversible(ByVal journalType As String) As Boolean
    Select Case journalType
        Case modConstants.JT_REVERSAL, modConstants.JT_INTEREST, modConstants.JT_FD_MATURITY, modConstants.JT_FD_BREAK, _
             modConstants.JT_IB_OUT, modConstants.JT_IB_SETTLE, modConstants.JT_IB_FAIL, modConstants.JT_IB_RETURN, modConstants.JT_IB_IN, _
             modConstants.JT_FX_REVAL, modConstants.JT_YEAR_END, modConstants.JT_BOND_ACCRETION, modConstants.JT_BOND_MATURITY, modConstants.JT_INV_ACCRUAL, modConstants.JT_INV_VALUATION
            IsReversible = False
        Case Else
            IsReversible = True
    End Select
End Function

' Reverses every journal of one action (GroupID): same lines, debit and credit swapped, dated today.
Public Function ReverseGroup(ByVal uow As clsUnitOfWork, ByVal groupId As String, ByVal createdBy As String, ByVal approvedBy As String) As clsResult
    Dim hdr As Variant, lns As Variant, txns As Variant, i As Long, ids As Collection, id As Variant
    Dim newGrp As String, jr As clsResult, msg As String, jd As Date
    hdr = uow.WorkingTable(modConstants.TBL_JOURNAL)
    msg = svcPayroll.PayrollReversalProblem(uow, groupId)
    If Len(msg) = 0 Then msg = svcIntl.IntlReversalProblem(uow, groupId)
    If Len(msg) > 0 Then Set ReverseGroup = modMessages.FailResult(msg): Exit Function
    Set ids = New Collection
    If Not IsEmpty(hdr) Then
        For i = 1 To UBound(hdr, 1)
            If CStr(hdr(i, repoJournal.JNC_GROUP)) = groupId Then
                If CStr(hdr(i, repoJournal.JNC_STATUS)) <> modConstants.STATUS_POSTED Then
                    Set ReverseGroup = modMessages.FailResult("Group " & groupId & " is already reversed.")
                    Exit Function
                End If
                If Not IsReversible(CStr(hdr(i, repoJournal.JNC_TYPE))) Then
                    Set ReverseGroup = modMessages.FailResult("A " & CStr(hdr(i, repoJournal.JNC_TYPE)) & " journal cannot be reversed.")
                    Exit Function
                End If
                If modDateUtils.TryToDate(hdr(i, 2), jd) Then
                    msg = svcYearEnd.ClosedYearProblem(jd, CStr(hdr(i, repoJournal.JNC_TYPE)))
                    If Len(msg) > 0 Then Set ReverseGroup = modMessages.FailResult(msg): Exit Function
                End If
                ids.Add CStr(hdr(i, repoJournal.JNC_ID))
            End If
        Next i
    End If
    If ids.Count = 0 Then
        Set ReverseGroup = modMessages.FailResult("No journal found for group " & groupId & ".")
        Exit Function
    End If
    lns = uow.WorkingTable(modConstants.TBL_JLINES)
    txns = uow.WorkingTable(modConstants.TBL_TRANSACTIONS)
    newGrp = uow.NewGroupId()
    For Each id In ids
        Set jr = ReverseOne(uow, CStr(id), newGrp, createdBy, approvedBy, lns, txns)
        If Not jr.Success Then
            Set ReverseGroup = jr
            Exit Function
        End If
    Next id
    For i = 1 To UBound(hdr, 1)
        If CStr(hdr(i, repoJournal.JNC_GROUP)) = groupId Then hdr(i, repoJournal.JNC_STATUS) = modConstants.STATUS_REVERSED
    Next i
    uow.ReplaceTable modConstants.TBL_JOURNAL, hdr
    If Not IsEmpty(txns) Then uow.ReplaceTable modConstants.TBL_TRANSACTIONS, txns
    svcPayroll.AfterReversal uow, groupId
    repoAudit.Add uow, "Reversal", groupId, "Reversed " & ids.Count & " journal(s) as group " & newGrp
    Set ReverseGroup = modMessages.OkResult("Group " & groupId & " reversed (new group " & newGrp & ").", ids.Count, newGrp)
End Function

Private Function ReverseOne(ByVal uow As clsUnitOfWork, ByVal origId As String, ByVal newGrp As String, ByVal createdBy As String, _
                            ByVal approvedBy As String, ByVal lns As Variant, ByRef txns As Variant) As clsResult
    Dim lines As Collection, i As Long, ln As clsJournalLine, jr As clsResult, a As clsAccount
    Dim dirNew As String, amt As Variant
    Set lines = New Collection
    For i = 1 To UBound(lns, 1)
        If CStr(lns(i, repoJournal.LNC_JOURNAL)) = origId Then
            Set ln = New clsJournalLine
            ln.BranchID = CStr(lns(i, repoJournal.LNC_BRANCH))
            ln.GLCode = CStr(lns(i, repoJournal.LNC_GL))
            ln.AccountNo = CStr(lns(i, repoJournal.LNC_ACCOUNT))
            ln.Ccy = CStr(lns(i, 6))
            If Len(ln.Ccy) = 0 Then ln.Ccy = modConstants.CCY_MMK
            If ln.Ccy <> modConstants.CCY_MMK Then
                ln.FxRate = modMoney.ToDec(lns(i, 9))
                ln.ForeignAmount = modMoney.ToDec(lns(i, 10))
            End If
            ln.Debit = modMoney.ToDec(lns(i, repoJournal.LNC_CREDIT))
            ln.Credit = modMoney.ToDec(lns(i, repoJournal.LNC_DEBIT))
            lines.Add ln
        End If
    Next i
    Set jr = PostJournal(uow, modConstants.JT_REVERSAL, "Reversal of " & origId, newGrp, lines, origId)
    If Not jr.Success Then
        Set ReverseOne = jr
        Exit Function
    End If
    For Each ln In lines
        If ln.GLCode = modConstants.GL_VAULT And ln.Credit > 0 Then
            If uow.GLNet(ln.BranchID, modConstants.GL_VAULT) < 0 Then
                Set ReverseOne = modMessages.FailResult("Reversal refused: the vault of " & ln.BranchID & " would go below 0.")
                Exit Function
            End If
        End If
    Next ln
    If Not IsEmpty(txns) Then
        For i = 1 To UBound(txns, 1)
            If CStr(txns(i, repoTransaction.TXC_JOURNAL)) = origId And CStr(txns(i, repoTransaction.TXC_STATUS)) = modConstants.STATUS_POSTED Then
                Set a = uow.GetAccount(CStr(txns(i, repoTransaction.TXC_ACCOUNT)))
                amt = modMoney.ToDec(txns(i, 7))
                If a Is Nothing Then
                    Set ReverseOne = modMessages.FailResult("Reversal refused: unknown account " & CStr(txns(i, repoTransaction.TXC_ACCOUNT)) & ".")
                    Exit Function
                End If
                If a.Status = modConstants.STATUS_CLOSED Then
                    Set ReverseOne = modMessages.FailResult("Reversal refused: account " & a.AccountNo & " is Closed.")
                    Exit Function
                End If
                If CStr(txns(i, 6)) = modConstants.DIR_CREDIT Then dirNew = modConstants.DIR_DEBIT Else dirNew = modConstants.DIR_CREDIT
                If dirNew = modConstants.DIR_DEBIT And a.Balance < amt Then
                    Set ReverseOne = modMessages.FailResult("Reversal refused: the balance of " & a.AccountNo & " would go negative.")
                    Exit Function
                End If
                svcAccounts.Effect uow, a, modConstants.TXN_REVERSAL, dirNew, amt, CStr(txns(i, 9)), CStr(txns(i, 10)), _
                    "Reversal of " & CStr(txns(i, 1)), "", jr.RefID, newGrp, createdBy, approvedBy, False
                txns(i, repoTransaction.TXC_STATUS) = modConstants.STATUS_REVERSED
            End If
        Next i
    End If
    Set ReverseOne = modMessages.OkResult("", 1, jr.RefID)
End Function
