Attribute VB_Name = "svcEndOfDayExtra"
Option Explicit

' The end-of-day steps of Stages 2 to 4 (STAGES_2_TO_4.md section 0.6). Everything here is skipped when its tables do not exist.
' stats maps tblDailyClose column names to values.

' Extra tables the end of day may change (saved before the run and restored if a check fails).
Public Function ExtraTables() As Variant
    Dim c As Collection, out() As Variant, i As Long, t As Variant
    Set c = New Collection
    For Each t In Array(modConstants.TBL_COSTPOSTINGS, modConstants.TBL_EXPENSES, modConstants.TBL_FIXEDASSETS, modConstants.TBL_LOANS, modConstants.TBL_LOANSCHEDULE, _
                        modConstants.TBL_LOANPAYMENTS, modConstants.TBL_COLLATERAL, modConstants.TBL_REPOSSESSIONS, modConstants.TBL_ASSETOPS, _
                        modConstants.TBL_EXCHANGERATES, modConstants.TBL_FXDEALS, modConstants.TBL_FXPOSITION, modConstants.TBL_INTLPAYMENTS, modConstants.TBL_AGENTS, modConstants.TBL_MERCHANTS, _
                        modConstants.TBL_PAYROLLRUNS, modConstants.TBL_PAYSLIPS, modConstants.TBL_ADVANCES, _
                        modConstants.TBL_INVESTEES, modConstants.TBL_INVESTMENTS, modConstants.TBL_INVRETURNS, modConstants.TBL_INVVALUATIONS, modConstants.TBL_BONDS, _
                        modConstants.TBL_AMLALERTS, modConstants.TBL_AMLCASES, modConstants.TBL_RATIOHISTORY, modConstants.TBL_COMPLIANCE)
        If modSheetUtils.TableExists(CStr(t)) Then c.Add CStr(t)
    Next t
    If c.Count = 0 Then
        ExtraTables = Array()
        Exit Function
    End If
    ReDim out(0 To c.Count - 1)
    For i = 1 To c.Count
        out(i - 1) = c(i)
    Next i
    ExtraTables = out
End Function

' Stage 3: interest due, auto-debit, penalties, classes. Runs after interbank and fixed deposit maturity, before deposit interest accrual.
Public Function DailySteps(ByVal uow As clsUnitOfWork, ByVal d As Date, ByVal stats As Object) As clsResult
    Dim r As clsResult, decided As Long
    Set r = svcLoanEod.Daily(uow, d, stats)
    If Not r.Success Then Set DailySteps = r: Exit Function
    Set r = svcIntl.DecideIntlDue(uow, d, decided)
    If Not r.Success Then Set DailySteps = r: Exit Function
    stats("IntlDecided") = decided
    Set r = svcInvest.InvestDaily(uow, d, stats)
    If Not r.Success Then Set DailySteps = r: Exit Function
    Set DailySteps = svcBonds.BondDaily(uow, d, stats)
End Function

' Last steps of the day, after every other posting: revaluation of foreign balances and the rates of the next day.
Public Function FinalSteps(ByVal uow As clsUnitOfWork, ByVal d As Date, ByVal stats As Object) As clsResult
    Dim r As clsResult
    Set r = svcFxReval.RevalueFx(uow, d, stats)
    If Not r.Success Then Set FinalSteps = r: Exit Function
    Set FinalSteps = svcFxReval.AutoRates(uow, d)
End Function

' Step 8 (after the day is committed): compliance ratios with their history and events, then the AML scan of the day.
Public Function PostCommitSteps(ByVal user As clsUser, ByVal d As Date, ByVal stats As Object) As clsResult
    Dim uow As clsUnitOfWork, ratios As Collection, n As Long, a As Long
    If Not modSheetUtils.TableExists(modConstants.TBL_RATIOHISTORY) Then Set PostCommitSteps = modMessages.OkResult("", 0): Exit Function
    Set uow = New clsUnitOfWork
    uow.Init d, user.UserID
    Set ratios = svcRatios.ComputeRatios(uow, d)
    n = svcRatios.WriteHistory(uow, d, ratios)
    a = svcAml.DailyScan(uow, d)
    uow.Commit
    stats("RatiosBreached") = n
    stats("AmlAlerts") = a
    Set PostCommitSteps = modMessages.OkResult("", 1)
End Function

' Month end: Stage 2 costs and depreciation, Stage 3 provisions, Stage 4 holding costs.
Public Function MonthEndSteps(ByVal uow As clsUnitOfWork, ByVal d As Date, ByVal stats As Object) As clsResult
    Dim r As clsResult, n As Long, tot As Variant, nd As Long, td As Variant, period As String
    period = modDateUtils.PeriodOf(d)
    If modSheetUtils.TableExists(modConstants.TBL_COSTPOSTINGS) Then
        Set r = svcCosts.PostPeriod(uow, period, n, tot)
        If Not r.Success Then Set MonthEndSteps = r: Exit Function
        Set r = svcFixedAssets.Depreciate(uow, period, nd, td)
        If Not r.Success Then Set MonthEndSteps = r: Exit Function
        stats("CostsPosted") = tot
        stats("DepreciationPosted") = td
    End If
    Set r = svcLoanEod.Provisions(uow, d, stats)
    If Not r.Success Then Set MonthEndSteps = r: Exit Function
    Set r = svcAssetOps.MonthEnd(uow, d, stats)
    If Not r.Success Then Set MonthEndSteps = r: Exit Function
    Set r = svcInvest.InvestMonthEnd(uow, d, stats)
    If Not r.Success Then Set MonthEndSteps = r: Exit Function
    Set r = svcBonds.BondMonthEnd(uow, d, stats)
    If Not r.Success Then Set MonthEndSteps = r: Exit Function
    Set r = svcInsurance.QuarterEnd(uow, d, stats)
    If Not r.Success Then Set MonthEndSteps = r: Exit Function
    MonthEndWarnings d
    Set MonthEndSteps = modMessages.OkResult("", 1)
End Function

' Warnings only (never stop the day): no payroll run for the month, branches without a manager.
Private Sub MonthEndWarnings(ByVal d As Date)
    Dim runs As Variant, i As Long, found As Boolean, emps As Variant, s As String
    If modSheetUtils.TableExists(modConstants.TBL_PAYROLLRUNS) Then
        runs = modSheetUtils.ReadTable(modConstants.TBL_PAYROLLRUNS)
        If Not IsEmpty(runs) Then
            For i = 1 To UBound(runs, 1)
                If CStr(runs(i, 2)) = modDateUtils.PeriodOf(d) Then found = True
            Next i
        End If
        If Not found Then clsLogger.Warn "svcEndOfDayExtra.MonthEndWarnings", "No payroll run exists for " & modDateUtils.PeriodOf(d) & "."
    End If
    If modSheetUtils.TableExists(modConstants.TBL_EMPLOYEES) Then
        emps = modSheetUtils.ReadTable(modConstants.TBL_EMPLOYEES)
        If Not IsEmpty(emps) Then
            s = svcReportsStaff.BranchesWithoutManager(emps, repoBranch.LoadBranches())
            If Len(Trim$(s)) > 0 Then clsLogger.Warn "svcEndOfDayExtra.MonthEndWarnings", "Branches without a manager: " & Trim$(s)
        End If
    End If
End Sub
