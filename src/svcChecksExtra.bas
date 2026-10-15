Attribute VB_Name = "svcChecksExtra"
Option Explicit

' Control checks 7 to 16 (Stages 2 to 4); checks 17 to 29 are in svcChecksLate and svcChecksFx. A check is skipped when its tables do not exist yet.

Private Function GLTotals(ByVal lns As Variant) As Object
    Dim d As Object, i As Long, g As String
    Set d = CreateObject("Scripting.Dictionary")
    If Not IsEmpty(lns) Then
        For i = 1 To UBound(lns, 1)
            g = CStr(lns(i, repoJournal.LNC_GL))
            If d.Exists(g) Then
                d(g) = d(g) + modMoney.ToDec(lns(i, repoJournal.LNC_DEBIT)) - modMoney.ToDec(lns(i, repoJournal.LNC_CREDIT))
            Else
                d(g) = modMoney.ToDec(lns(i, repoJournal.LNC_DEBIT)) - modMoney.ToDec(lns(i, repoJournal.LNC_CREDIT))
            End If
        Next i
    End If
    Set GLTotals = d
End Function

Private Function Net(ByVal d As Object, ByVal gl As String) As Variant
    If d.Exists(gl) Then Net = d(gl) Else Net = CDec(0)
End Function

Private Function Verdict(ByVal ok As Boolean) As String
    If ok Then Verdict = "PASS" Else Verdict = "FAIL"
End Function

Private Function Cmp(ByVal a As Variant, ByVal b As Variant) As String
    Cmp = "GL " & modMoney.FormatAmount(a) & " / records " & modMoney.FormatAmount(b)
End Function

' ---- Stage 2: checks 7 to 9 ----

Private Sub CheckStage2(ByVal rowsList As Collection, ByVal gl As Object)
    Dim assets As Variant, i As Long, cost As Variant, accum As Variant, bad As Long, u As clsUnitOfWork, unpaid As Variant
    Dim posted As Variant, seen As Object, key As String, dup As Long
    If Not modSheetUtils.TableExists(modConstants.TBL_FIXEDASSETS) Then Exit Sub
    assets = modSheetUtils.ReadTable(modConstants.TBL_FIXEDASSETS)
    If Not repoAsset.IsUpgraded(assets) Then Exit Sub
    cost = CDec(0)
    accum = CDec(0)
    If Not IsEmpty(assets) Then
        For i = 1 To UBound(assets, 1)
            If CStr(assets(i, repoAsset.FAC_STATUS)) = repoAsset.ASSET_INUSE And Len(CStr(assets(i, repoAsset.FAC_LOADED))) > 0 Then
                cost = cost + modMoney.ToDec(assets(i, repoAsset.FAC_COST))
                accum = accum + modMoney.ToDec(assets(i, repoAsset.FAC_ACCUM))
                If modMoney.ToDec(assets(i, repoAsset.FAC_ACCUM)) > modMoney.ToDec(assets(i, repoAsset.FAC_COST)) - modMoney.ToDec(assets(i, repoAsset.FAC_RESIDUAL)) Then bad = bad + 1
            End If
        Next i
    End If
    svcChecks.AddRow rowsList, 7, "GL 1500 = cost of assets in use; GL 1510 = accumulated depreciation", Verdict(Net(gl, modConstants.GL_ASSETS) = cost And -Net(gl, modConstants.GL_ACCUM_DEP) = accum), _
        "1500: " & Cmp(Net(gl, modConstants.GL_ASSETS), cost) & "; 1510: " & Cmp(-Net(gl, modConstants.GL_ACCUM_DEP), accum)
    If modSheetUtils.TableExists(modConstants.TBL_COSTPOSTINGS) Then
        Set u = New clsUnitOfWork
        u.Init repoConfig.BusinessDate(), "SYSTEM"
        unpaid = svcCosts.UnpaidTotal(u)
        svcChecks.AddRow rowsList, 8, "GL 2300 = unpaid cost postings + unpaid expenses", Verdict(-Net(gl, modConstants.GL_ACCRUED) = unpaid), Cmp(-Net(gl, modConstants.GL_ACCRUED), unpaid)
        posted = modSheetUtils.ReadTable(modConstants.TBL_COSTPOSTINGS)
        Set seen = CreateObject("Scripting.Dictionary")
        If Not IsEmpty(posted) Then
            For i = 1 To UBound(posted, 1)
                key = CStr(posted(i, repoCost.CPC_COST)) & "|" & CStr(posted(i, repoCost.CPC_PERIOD))
                If seen.Exists(key) Then dup = dup + 1 Else seen(key) = True
            Next i
        End If
        svcChecks.AddRow rowsList, 9, "No cost posted twice; depreciation never above cost less residual", Verdict(dup = 0 And bad = 0), dup & " duplicate postings, " & bad & " over-depreciated assets"
    End If
End Sub

' ---- Stage 3: checks 10 to 13 ----

Private Sub CheckStage3(ByVal rowsList As Collection, ByVal gl As Object)
    Dim loans As Variant, sched As Variant, i As Long, outst As Variant, intr As Variant, prov As Variant, st As String, loanId As String
    Dim sumDue As Object, sumPaid As Object, bad As Long, k As Variant, ok12 As Boolean, d As Date, strict As Boolean, sample As String
    If Not modSheetUtils.TableExists(modConstants.TBL_LOANS) Or Not modSheetUtils.TableExists(modConstants.TBL_LOANSCHEDULE) Then Exit Sub
    loans = modSheetUtils.ReadTable(modConstants.TBL_LOANS)
    sched = modSheetUtils.ReadTable(modConstants.TBL_LOANSCHEDULE)
    outst = CDec(0)
    prov = CDec(0)
    If Not IsEmpty(loans) Then
        For i = 1 To UBound(loans, 1)
            If CStr(loans(i, repoLoan.LOC_STATUS)) = repoLoan.LS_ACTIVE Then
                outst = outst + modMoney.ToDec(loans(i, repoLoan.LOC_OUTSTANDING))
                prov = prov + modMoney.ToDec(loans(i, repoLoan.LOC_PROVISION))
            End If
        Next i
    End If
    intr = CDec(0)
    Set sumDue = CreateObject("Scripting.Dictionary")
    Set sumPaid = CreateObject("Scripting.Dictionary")
    If Not IsEmpty(sched) Then
        For i = 1 To UBound(sched, 1)
            loanId = CStr(sched(i, repoLoan.SCC_LOAN))
            st = CStr(sched(i, repoLoan.SCC_STATUS))
            If st = "Due" Or st = "Partial" Then intr = intr + modMoney.ToDec(sched(i, repoLoan.SCC_INT)) - modMoney.ToDec(sched(i, repoLoan.SCC_INTPAID))
            If sumDue.Exists(loanId) Then
                sumDue(loanId) = sumDue(loanId) + modMoney.ToDec(sched(i, repoLoan.SCC_PRIN))
                sumPaid(loanId) = sumPaid(loanId) + modMoney.ToDec(sched(i, repoLoan.SCC_PRINPAID))
            Else
                sumDue(loanId) = modMoney.ToDec(sched(i, repoLoan.SCC_PRIN))
                sumPaid(loanId) = modMoney.ToDec(sched(i, repoLoan.SCC_PRINPAID))
            End If
        Next i
    End If
    ' Interest of Active loans only: add back the interest of loans that are no longer Active.
    If Not IsEmpty(sched) And Not IsEmpty(loans) Then
        For i = 1 To UBound(sched, 1)
            st = CStr(sched(i, repoLoan.SCC_STATUS))
            If st = "Due" Or st = "Partial" Then
                If LoanStatusOf(loans, CStr(sched(i, repoLoan.SCC_LOAN))) <> repoLoan.LS_ACTIVE Then intr = intr - (modMoney.ToDec(sched(i, repoLoan.SCC_INT)) - modMoney.ToDec(sched(i, repoLoan.SCC_INTPAID)))
            End If
        Next i
    End If
    svcChecks.AddRow rowsList, 10, "GL 1300 = principal outstanding of Active loans", Verdict(Net(gl, modConstants.GL_LOANS) = outst), Cmp(Net(gl, modConstants.GL_LOANS), outst)
    svcChecks.AddRow rowsList, 11, "GL 1310 = recognized unpaid interest of Active loans", Verdict(Net(gl, modConstants.GL_LOAN_INTEREST_REC) = intr), Cmp(Net(gl, modConstants.GL_LOAN_INTEREST_REC), intr)
    d = repoConfig.BusinessDate()
    strict = modDateUtils.IsLastDayOfMonth(d)
    ok12 = (-Net(gl, modConstants.GL_LOAN_ALLOWANCE) = prov)
    If ok12 Then
        svcChecks.AddRow rowsList, 12, "GL 1320 = provisions of Active loans", "PASS", Cmp(-Net(gl, modConstants.GL_LOAN_ALLOWANCE), prov)
    ElseIf strict Then
        svcChecks.AddRow rowsList, 12, "GL 1320 = provisions of Active loans", "FAIL", Cmp(-Net(gl, modConstants.GL_LOAN_ALLOWANCE), prov)
    Else
        svcChecks.AddRow rowsList, 12, "GL 1320 = provisions of Active loans", "WARN", "Differs until the month-end posting: " & Cmp(-Net(gl, modConstants.GL_LOAN_ALLOWANCE), prov)
    End If
    If Not IsEmpty(loans) Then
        For i = 1 To UBound(loans, 1)
            st = CStr(loans(i, repoLoan.LOC_STATUS))
            If st = repoLoan.LS_ACTIVE Or st = repoLoan.LS_CLOSED Then
                loanId = CStr(loans(i, repoLoan.LOC_ID))
                If Not sumDue.Exists(loanId) Then
                    bad = bad + 1
                ElseIf sumDue(loanId) <> modMoney.ToDec(loans(i, repoLoan.LOC_PRINCIPAL)) Or sumPaid(loanId) + modMoney.ToDec(loans(i, repoLoan.LOC_OUTSTANDING)) <> modMoney.ToDec(loans(i, repoLoan.LOC_PRINCIPAL)) Then
                    bad = bad + 1
                    If Len(sample) < 40 Then sample = sample & loanId & " "
                End If
            End If
        Next i
    End If
    svcChecks.AddRow rowsList, 13, "Loan schedules add up to the principal", Verdict(bad = 0), bad & " loans differ " & sample
End Sub

Private Function LoanStatusOf(ByVal loans As Variant, ByVal loanId As String) As String
    Dim r As Long
    r = repoLoan.FindLoanRow(loans, loanId)
    If r > 0 Then LoanStatusOf = CStr(loans(r, repoLoan.LOC_STATUS))
End Function

' ---- Stage 4: checks 14 to 16 ----

Private Sub CheckStage4(ByVal rowsList As Collection, ByVal gl As Object)
    Dim repos As Variant, items As Variant, loans As Variant, sales As Variant, i As Long, held As Variant, bad15 As Long, bad16 As Long, seen As Object, st As String, j As Long, found As Boolean
    If Not modSheetUtils.TableExists(modConstants.TBL_REPOSSESSIONS) Or Not modSheetUtils.TableExists(modConstants.TBL_COLLATERAL) Then Exit Sub
    repos = modSheetUtils.ReadTable(modConstants.TBL_REPOSSESSIONS)
    items = modSheetUtils.ReadTable(modConstants.TBL_COLLATERAL)
    loans = modSheetUtils.ReadTable(modConstants.TBL_LOANS)
    sales = modSheetUtils.ReadTable(modConstants.TBL_ASSETSALES)
    held = CDec(0)
    If Not IsEmpty(repos) Then
        For i = 1 To UBound(repos, 1)
            st = CStr(repos(i, repoCollateral.RPC_STATUS))
            If st = repoCollateral.RS_HELD Or st = repoCollateral.RS_RENTED Or st = repoCollateral.RS_INBUSINESS Then held = held + modMoney.ToDec(repos(i, repoCollateral.RPC_CARRYING))
        Next i
    End If
    svcChecks.AddRow rowsList, 14, "GL 1400 = carrying value of repossessed assets held", Verdict(Net(gl, modConstants.GL_REPO_ASSETS) = held), Cmp(Net(gl, modConstants.GL_REPO_ASSETS), held)
    Set seen = CreateObject("Scripting.Dictionary")
    If Not IsEmpty(items) Then
        For i = 1 To UBound(items, 1)
            If seen.Exists(CStr(items(i, repoCollateral.COC_ID))) Then bad15 = bad15 + 1 Else seen(CStr(items(i, repoCollateral.COC_ID))) = True
            If CStr(items(i, repoCollateral.COC_LIEN)) = repoCollateral.LIEN_PLEDGED Then
                st = LoanStatusOf(loans, CStr(items(i, repoCollateral.COC_LOAN)))
                If st <> repoLoan.LS_APPLIED And st <> repoLoan.LS_APPROVED And st <> repoLoan.LS_ACTIVE Then bad15 = bad15 + 1
            End If
        Next i
    End If
    svcChecks.AddRow rowsList, 15, "No item has two liens; every Pledged item backs an open loan", Verdict(bad15 = 0), bad15 & " problems"
    If Not IsEmpty(items) Then
        For i = 1 To UBound(items, 1)
            If CStr(items(i, repoCollateral.COC_LIEN)) = repoCollateral.LIEN_REPOSSESSED Or CStr(items(i, repoCollateral.COC_LIEN)) = repoCollateral.LIEN_SOLD Then
                If repoCollateral.FindCollateralRow(repos, repoCollateral.RPC_COLLATERAL, CStr(items(i, repoCollateral.COC_ID))) = 0 Then bad16 = bad16 + 1
            End If
        Next i
    End If
    If Not IsEmpty(repos) Then
        For i = 1 To UBound(repos, 1)
            If CStr(repos(i, repoCollateral.RPC_STATUS)) = repoCollateral.RS_CLOSED Then
                found = False
                If Not IsEmpty(sales) Then
                    For j = 1 To UBound(sales, 1)
                        If CStr(sales(j, 2)) = CStr(repos(i, repoCollateral.RPC_ID)) Then found = True
                    Next j
                End If
                If Not found Then bad16 = bad16 + 1
            End If
        Next i
    End If
    svcChecks.AddRow rowsList, 16, "Every repossessed or sold item has a record; every closed record has a sale", Verdict(bad16 = 0), bad16 & " problems"
End Sub

' Called by svcChecks.RunChecks with the journal lines already loaded.
Public Sub RunExtraChecks(ByVal rowsList As Collection, ByVal lns As Variant)
    Dim gl As Object
    Set gl = GLTotals(lns)
    CheckStage2 rowsList, gl
    CheckStage3 rowsList, gl
    CheckStage4 rowsList, gl
    svcChecksFx.RunFxChecks rowsList, lns, gl
    svcChecksLate.RunLateChecks rowsList, lns, gl
End Sub
