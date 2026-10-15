Attribute VB_Name = "svcLoanEod"
Option Explicit

' Stage 3 end-of-day steps: interest becomes due, auto-debit, penalties, days past due and class, monthly provisions.

Private Function ProductOf(ByVal cache As Object, ByVal code As String) As Variant
    If Not cache.Exists(code) Then cache(code) = repoLoan.ProductRow(code)
    ProductOf = cache(code)
End Function

' Interest of each installment that falls due on or before d is recognized: H: Dr 1310, Cr 4000.
Private Function RecognizeInterest(ByVal uow As clsUnitOfWork, ByVal d As Date, ByVal loans As Variant, ByVal row As Long, ByRef sched As Variant, ByRef count As Long) As clsResult
    Dim i As Long, due As Date, lines As Collection, jr As clsResult, loanId As String, h As String, intr As Variant, grp As String
    loanId = CStr(loans(row, repoLoan.LOC_ID))
    h = CStr(loans(row, repoLoan.LOC_BRANCH))
    If IsEmpty(sched) Then Set RecognizeInterest = modMessages.OkResult("", 0): Exit Function
    For i = 1 To UBound(sched, 1)
        If CStr(sched(i, repoLoan.SCC_LOAN)) = loanId And CStr(sched(i, repoLoan.SCC_STATUS)) = "Future" Then
            If modDateUtils.TryToDate(sched(i, repoLoan.SCC_DUE), due) Then
                If due <= d Then
                    sched(i, repoLoan.SCC_STATUS) = "Due"
                    count = count + 1
                    intr = modMoney.ToDec(sched(i, repoLoan.SCC_INT))
                    If intr > 0 Then
                        Set lines = New Collection
                        svcLedger.AddDr lines, h, modConstants.GL_LOAN_INTEREST_REC, intr
                        svcLedger.AddCr lines, h, modConstants.GL_LOAN_INCOME, intr
                        grp = uow.NewGroupId()
                        Set jr = svcLedger.PostJournal(uow, modConstants.JT_LOAN_INT_DUE, "Interest due " & loanId & " installment " & CStr(sched(i, repoLoan.SCC_NO)), grp, lines)
                        If Not jr.Success Then Set RecognizeInterest = jr: Exit Function
                    End If
                End If
            End If
        End If
    Next i
    Set RecognizeInterest = modMessages.OkResult("", count)
End Function

' Takes unpaid installments (oldest first) from the loan account when it can pay the whole installment plus penalty and keep MinBalance.
Private Function AutoDebit(ByVal uow As clsUnitOfWork, ByVal d As Date, ByRef loans As Variant, ByVal row As Long, ByRef sched As Variant, ByRef count As Long) As clsResult
    Dim i As Long, due As Date, a As clsAccount, need As Variant, pen As Variant, intr As Variant, prin As Variant, r As clsResult, penDue As Variant, loanId As String, grp As String
    loanId = CStr(loans(row, repoLoan.LOC_ID))
    Set a = uow.GetAccount(CStr(loans(row, repoLoan.LOC_ACCOUNT)))
    If a Is Nothing Then Set AutoDebit = modMessages.OkResult("", 0): Exit Function
    If IsEmpty(sched) Then Set AutoDebit = modMessages.OkResult("", 0): Exit Function
    For i = 1 To UBound(sched, 1)
        If CStr(sched(i, repoLoan.SCC_LOAN)) = loanId And (CStr(sched(i, repoLoan.SCC_STATUS)) = "Due" Or CStr(sched(i, repoLoan.SCC_STATUS)) = "Partial") Then
            If modDateUtils.TryToDate(sched(i, repoLoan.SCC_DUE), due) Then
                If due <= d Then
                    penDue = modMoney.RoundCcy(modMoney.ToDec(loans(row, repoLoan.LOC_PENALTY)), modConstants.CCY_MMK)
                    If penDue < 0 Then penDue = CDec(0)
                    need = svcLoanPayments.RowUnpaid(sched, i) + penDue
                    If need > 0 And a.Status = modConstants.STATUS_ACTIVE And a.Balance - need >= a.MinBalance Then
                        grp = uow.NewGroupId()
                        Set r = svcLoanPayments.ApplyPayment(uow, loans, row, sched, need, "AutoDebit", "Installment", a, a.BranchID, grp, "SYSTEM", "", pen, intr, prin)
                        If Not r.Success Then
                            clsLogger.Warn "svcLoanEod.AutoDebit", "Auto-debit of loan " & loanId & " failed: " & r.Message
                            Set AutoDebit = modMessages.OkResult("", count)
                            Exit Function
                        End If
                        count = count + 1
                        If CStr(loans(row, repoLoan.LOC_STATUS)) <> repoLoan.LS_ACTIVE Then Exit For
                    Else
                        Exit For
                    End If
                End If
            End If
        End If
    Next i
    Set AutoDebit = modMessages.OkResult("", count)
End Function

' Penalty per day on every unpaid installment past due date + grace days: unpaid amount x penalty rate / 365 (memo only, 4 decimals).
Private Function AccruePenalty(ByVal d As Date, ByVal loans As Variant, ByVal row As Long, ByVal sched As Variant, ByVal prow As Variant) As Variant
    Dim i As Long, due As Date, loanId As String, add As Variant, rate As Variant, grace As Long
    add = CDec(0)
    If IsEmpty(prow) Or IsEmpty(sched) Then AccruePenalty = add: Exit Function
    loanId = CStr(loans(row, repoLoan.LOC_ID))
    rate = modMoney.ToDec(prow(repoLoan.LPR_PENALTY))
    grace = CLng(modMoney.ToDec(prow(repoLoan.LPR_GRACE)))
    For i = 1 To UBound(sched, 1)
        If CStr(sched(i, repoLoan.SCC_LOAN)) = loanId And (CStr(sched(i, repoLoan.SCC_STATUS)) = "Due" Or CStr(sched(i, repoLoan.SCC_STATUS)) = "Partial") Then
            If modDateUtils.TryToDate(sched(i, repoLoan.SCC_DUE), due) Then
                If d > due + grace Then add = add + modMoney.RoundAccrual(svcLoanPayments.RowUnpaid(sched, i) * rate / repoConfig.GetDec(modConstants.CFG_DAYCOUNT))
            End If
        End If
    Next i
    AccruePenalty = add
End Function

' Days past due = d minus the oldest unpaid due date; class and required provision follow from it.
Private Sub UpdateClass(ByVal d As Date, ByRef loans As Variant, ByVal row As Long, ByVal sched As Variant, ByVal classes As Variant)
    Dim i As Long, due As Date, oldest As Date, found As Boolean, dpd As Long, cls As Long, pct As Variant, loanId As String
    loanId = CStr(loans(row, repoLoan.LOC_ID))
    If Not IsEmpty(sched) Then
        For i = 1 To UBound(sched, 1)
            If CStr(sched(i, repoLoan.SCC_LOAN)) = loanId And (CStr(sched(i, repoLoan.SCC_STATUS)) = "Due" Or CStr(sched(i, repoLoan.SCC_STATUS)) = "Partial") Then
                If modDateUtils.TryToDate(sched(i, repoLoan.SCC_DUE), due) Then
                    If due <= d And ((Not found) Or due < oldest) Then
                        oldest = due
                        found = True
                    End If
                End If
            End If
        Next i
    End If
    If found Then dpd = modDateUtils.DaysBetween(oldest, d)
    If repoLoan.ClassFor(classes, dpd, cls, pct) Then
        loans(row, repoLoan.LOC_CLASS) = cls
        loans(row, repoLoan.LOC_PROVISION) = modMoney.RoundCcy(modMoney.ToDec(loans(row, repoLoan.LOC_OUTSTANDING)) * pct, modConstants.CCY_MMK)
    End If
    loans(row, repoLoan.LOC_DPD) = dpd
End Sub

' Daily loan steps (step 3 of the end of day) for every Active loan. Fills InstallmentsDue, AutoDebited, PenaltyAccrued in stats.
Public Function Daily(ByVal uow As clsUnitOfWork, ByVal d As Date, ByVal stats As Object) As clsResult
    Dim loans As Variant, sched As Variant, classes As Variant, products As Object, row As Long, r As clsResult, nDue As Long, nAuto As Long, pen As Variant, add As Variant
    If Not modSheetUtils.TableExists(modConstants.TBL_LOANS) Or Not modSheetUtils.TableExists(modConstants.TBL_LOANSCHEDULE) Then Set Daily = modMessages.OkResult("", 0): Exit Function
    loans = uow.WorkingTable(modConstants.TBL_LOANS)
    If IsEmpty(loans) Then Set Daily = modMessages.OkResult("", 0): Exit Function
    sched = uow.WorkingTable(modConstants.TBL_LOANSCHEDULE)
    classes = modSheetUtils.ReadTable(modConstants.TBL_LOANCLASSES)
    Set products = CreateObject("Scripting.Dictionary")
    pen = CDec(0)
    For row = 1 To UBound(loans, 1)
        If CStr(loans(row, repoLoan.LOC_STATUS)) = repoLoan.LS_ACTIVE Then
            Set r = RecognizeInterest(uow, d, loans, row, sched, nDue)
            If Not r.Success Then Set Daily = r: Exit Function
            If repoConfig.GetBool(modConstants.CFG_LOAN_AUTODEBIT) Then
                Set r = AutoDebit(uow, d, loans, row, sched, nAuto)
                If Not r.Success Then Set Daily = r: Exit Function
            End If
            If CStr(loans(row, repoLoan.LOC_STATUS)) = repoLoan.LS_ACTIVE Then
                add = AccruePenalty(d, loans, row, sched, ProductOf(products, CStr(loans(row, repoLoan.LOC_PRODUCT))))
                loans(row, repoLoan.LOC_PENALTY) = modMoney.ToDec(loans(row, repoLoan.LOC_PENALTY)) + add
                pen = pen + add
                UpdateClass d, loans, row, sched, classes
            End If
        End If
    Next row
    uow.ReplaceTable modConstants.TBL_LOANS, loans
    If Not IsEmpty(sched) Then uow.ReplaceTable modConstants.TBL_LOANSCHEDULE, sched
    stats("InstallmentsDue") = nDue
    stats("AutoDebited") = nAuto
    stats("PenaltyAccrued") = pen
    Set Daily = modMessages.OkResult("", nDue)
End Function

' Month end: bring the loan loss allowance (GL 1320) of every branch to the sum of the required provisions of its Active loans.
Public Function Provisions(ByVal uow As clsUnitOfWork, ByVal d As Date, ByVal stats As Object) As clsResult
    Dim loans As Variant, required As Object, i As Long, b As Variant, cur As Variant, diff As Variant, lines As Collection, jr As clsResult, total As Variant, grp As String
    If Not modSheetUtils.TableExists(modConstants.TBL_LOANS) Then Set Provisions = modMessages.OkResult("", 0): Exit Function
    loans = uow.WorkingTable(modConstants.TBL_LOANS)
    Set required = CreateObject("Scripting.Dictionary")
    If Not IsEmpty(loans) Then
        For i = 1 To UBound(loans, 1)
            If CStr(loans(i, repoLoan.LOC_STATUS)) = repoLoan.LS_ACTIVE Then
                b = CStr(loans(i, repoLoan.LOC_BRANCH))
                If required.Exists(b) Then required(b) = required(b) + modMoney.ToDec(loans(i, repoLoan.LOC_PROVISION)) Else required(b) = modMoney.ToDec(loans(i, repoLoan.LOC_PROVISION))
            End If
        Next i
    End If
    total = CDec(0)
    For Each b In uow.Branches().Keys
        cur = -uow.GLNet(CStr(b), modConstants.GL_LOAN_ALLOWANCE)
        diff = CDec(0)
        If required.Exists(CStr(b)) Then diff = required(CStr(b)) - cur Else diff = -cur
        If diff <> 0 Then
            Set lines = New Collection
            If diff > 0 Then
                svcLedger.AddDr lines, CStr(b), modConstants.GL_PROVISION_EXP, diff
                svcLedger.AddCr lines, CStr(b), modConstants.GL_LOAN_ALLOWANCE, diff
            Else
                svcLedger.AddDr lines, CStr(b), modConstants.GL_LOAN_ALLOWANCE, -diff
                svcLedger.AddCr lines, CStr(b), modConstants.GL_PROVISION_EXP, -diff
            End If
            grp = uow.NewGroupId()
            Set jr = svcLedger.PostJournal(uow, modConstants.JT_LOAN_PROVISION, "Loan loss provision " & modDateUtils.PeriodOf(d) & " " & CStr(b), grp, lines)
            If Not jr.Success Then Set Provisions = jr: Exit Function
            total = total + diff
        End If
    Next b
    stats("ProvisionPosted") = total
    Set Provisions = modMessages.OkResult("", 1)
End Function
