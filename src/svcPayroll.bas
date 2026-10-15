Attribute VB_Name = "svcPayroll"
Option Explicit

' Stage 5: payslip calculation, payroll runs (prepare, post, pay, reverse) and staff advances.
' Payslips are calculated in MMK with modMoney.RoundCcy (see STAGES_5_TO_8.md part 1.3 C).

Private Function NeedTables() As String
    If Not modSheetUtils.TableExists(modConstants.TBL_PAYROLLRUNS) Or Not modSheetUtils.TableExists(modConstants.TBL_EMPLOYEES) Then NeedTables = modMessages.MSG_SETUP_FIRST
End Function

Private Function Cfg(ByVal key As String) As Variant
    Cfg = repoConfig.GetDec(key)
End Function

' ---- Calculation pieces ----

' Monthly base salary from the scale: the latest EffectiveFrom on or before lastDay.
Private Function BaseFor(ByVal scales As Variant, ByVal pos As String, ByVal grade As Long, ByVal lastDay As Date, ByRef base As Variant) As Boolean
    Dim i As Long, eff As Date, best As Date, found As Boolean
    If IsEmpty(scales) Then Exit Function
    For i = 1 To UBound(scales, 1)
        If CStr(scales(i, 2)) = pos And CLng(modMoney.ToDec(scales(i, 3))) = grade Then
            If modDateUtils.TryToDate(scales(i, 5), eff) Then
                If eff <= lastDay And ((Not found) Or eff >= best) Then
                    best = eff
                    base = modMoney.ToDec(scales(i, 4))
                    found = True
                End If
            End If
        End If
    Next i
    BaseFor = found
End Function

' Annual income tax from the progressive brackets (each slice at its own rate). Uses the rows of the latest EffectiveFrom on or before lastDay.
Private Function AnnualTax(ByVal brackets As Variant, ByVal annual As Variant, ByVal lastDay As Date) As Variant
    Dim i As Long, eff As Date, best As Date, found As Boolean, lo As Variant, hi As Variant, slice As Variant, tax As Variant
    tax = CDec(0)
    If IsEmpty(brackets) Or annual <= 0 Then AnnualTax = tax: Exit Function
    For i = 1 To UBound(brackets, 1)
        If modDateUtils.TryToDate(brackets(i, 5), eff) Then
            If eff <= lastDay And ((Not found) Or eff > best) Then
                best = eff
                found = True
            End If
        End If
    Next i
    For i = 1 To UBound(brackets, 1)
        If modDateUtils.TryToDate(brackets(i, 5), eff) Then
            If eff = best Then
                lo = modMoney.ToDec(brackets(i, 2))
                If annual > lo Then
                    If Len(CStr(brackets(i, 3))) = 0 Then
                        slice = annual - lo
                    Else
                        hi = modMoney.ToDec(brackets(i, 3))
                        If annual < hi Then slice = annual - lo Else slice = hi - lo
                    End If
                    tax = tax + slice * modMoney.ToDec(brackets(i, 4))
                End If
            End If
        End If
    Next i
    AnnualTax = tax
End Function

' Allowances of one employee: taxable and non-taxable totals (flat amounts times the month factor, percent of the factored base).
Private Sub Allowances(ByVal rules As Variant, ByVal branchType As String, ByVal pos As String, ByVal lastDay As Date, ByVal factor As Variant, ByVal factoredBase As Variant, _
                       ByRef taxable As Variant, ByRef nonTaxable As Variant)
    Dim i As Long, eff As Date, ok As Boolean, v As Variant, ap As String
    taxable = CDec(0)
    nonTaxable = CDec(0)
    If IsEmpty(rules) Then Exit Sub
    For i = 1 To UBound(rules, 1)
        If Len(CStr(rules(i, 2))) > 0 And modDateUtils.TryToDate(rules(i, 9), eff) Then
            If eff <= lastDay Then
                ap = CStr(rules(i, 3))
                ok = (ap = "All") Or (ap = "BranchType" And CStr(rules(i, 4)) = branchType) Or (ap = "Position" And CStr(rules(i, 4)) = pos)
                If ok Then
                    If CStr(rules(i, 5)) = "PercentOfBase" Then
                        v = modMoney.RoundCcy(factoredBase * modMoney.ToDec(rules(i, 7)), modConstants.CCY_MMK)
                    Else
                        v = modMoney.RoundCcy(modMoney.ToDec(rules(i, 6)) * factor, modConstants.CCY_MMK)
                    End If
                    If CStr(rules(i, 8)) = modConstants.FLAG_NO Then nonTaxable = nonTaxable + v Else taxable = taxable + v
                End If
            End If
        End If
    Next i
End Sub

' Builds one payslip row (20 columns, Draft). Returns False when the employee has no salary scale row.
Private Function BuildSlip(ByVal slipId As String, ByVal runId As String, ByVal period As String, ByVal firstDay As Date, ByVal lastDay As Date, ByVal emps As Variant, ByVal er As Long, _
                           ByVal branchId As String, ByVal pos As String, ByVal grade As Long, ByVal branchType As String, ByVal scales As Variant, ByVal rules As Variant, _
                           ByVal brackets As Variant, ByVal att As Variant, ByVal adv As Variant, ByRef slip As Variant, ByRef msg As String) As Boolean
    Dim base As Variant, factor As Variant, fBase As Variant, hire As Date, ex As Date, s As Date, e As Date, daysIn As Long, daysEmp As Long
    Dim unpaidDays As Variant, otH As Variant, hotH As Variant, adj As Variant, unpaid As Variant, ta As Variant, nta As Variant, ot As Variant, hourly As Variant
    Dim gross As Variant, ssbBase As Variant, eeSsb As Variant, erSsb As Variant, taxInc As Variant, annual As Variant, tax As Variant, advRec As Variant, i As Long, r(1 To 20) As Variant
    If Not BaseFor(scales, pos, grade, lastDay, base) Then msg = "No salary scale for " & pos & " grade " & grade & ".": Exit Function
    hire = CDate(emps(er, repoStaff.EMC_HIRE))
    s = hire
    If s < firstDay Then s = firstDay
    e = lastDay
    If modDateUtils.TryToDate(emps(er, repoStaff.EMC_EXIT), ex) Then
        If ex < e Then e = ex
    End If
    daysIn = Day(lastDay)
    daysEmp = modDateUtils.DaysBetween(s, e) + 1
    If daysEmp >= daysIn Then factor = CDec(1) Else factor = CDec(daysEmp) / CDec(daysIn)
    fBase = base
    If factor < 1 Then fBase = modMoney.RoundCcy(base * factor, modConstants.CCY_MMK)
    unpaidDays = CDec(0)
    otH = CDec(0)
    hotH = CDec(0)
    adj = CDec(0)
    If Not IsEmpty(att) Then
        For i = 1 To UBound(att, 1)
            If CStr(att(i, 2)) = CStr(emps(er, repoStaff.EMC_ID)) And CStr(att(i, 3)) = period Then
                unpaidDays = modMoney.ToDec(att(i, 4))
                otH = modMoney.ToDec(att(i, 5))
                hotH = modMoney.ToDec(att(i, 6))
                adj = modMoney.ToDec(att(i, 7))
            End If
        Next i
    End If
    unpaid = modMoney.RoundCcy(base / Cfg(modConstants.CFG_WORKDAYS) * unpaidDays, modConstants.CCY_MMK)
    Allowances rules, branchType, pos, lastDay, factor, fBase, ta, nta
    hourly = base / (Cfg(modConstants.CFG_WORKDAYS) * Cfg(modConstants.CFG_HOURSPERDAY))
    ot = modMoney.RoundCcy(otH * hourly * Cfg(modConstants.CFG_OT_MULT), modConstants.CCY_MMK) + modMoney.RoundCcy(hotH * hourly * Cfg(modConstants.CFG_HOLIDAY_OT_MULT), modConstants.CCY_MMK)
    gross = fBase - unpaid + ta + nta + ot + adj
    ssbBase = gross
    If ssbBase > Cfg(modConstants.CFG_SSB_CAP) Then ssbBase = Cfg(modConstants.CFG_SSB_CAP)
    If ssbBase < 0 Then ssbBase = CDec(0)
    eeSsb = modMoney.RoundCcy(ssbBase * Cfg(modConstants.CFG_EE_SSB), modConstants.CCY_MMK)
    erSsb = modMoney.RoundCcy(ssbBase * Cfg(modConstants.CFG_ER_SSB), modConstants.CCY_MMK)
    taxInc = fBase - unpaid + ta + ot + adj - eeSsb
    annual = taxInc * 12 - modMoney.ToDec(emps(er, repoStaff.EMC_DEPENDENTS)) * Cfg(modConstants.CFG_DEPENDENT_RELIEF)
    If annual < 0 Then annual = CDec(0)
    tax = modMoney.RoundCcy(AnnualTax(brackets, annual, lastDay) / 12, modConstants.CCY_MMK)
    advRec = CDec(0)
    If Not IsEmpty(adv) Then
        For i = 1 To UBound(adv, 1)
            If CStr(adv(i, repoStaff.ADC_EMPLOYEE)) = CStr(emps(er, repoStaff.EMC_ID)) And CStr(adv(i, repoStaff.ADC_STATUS)) = "Active" Then
                advRec = modMoney.ToDec(adv(i, repoStaff.ADC_MONTHLY))
                If advRec > modMoney.ToDec(adv(i, repoStaff.ADC_OUTSTANDING)) Then advRec = modMoney.ToDec(adv(i, repoStaff.ADC_OUTSTANDING))
            End If
        Next i
    End If
    r(repoStaff.PSC_ID) = slipId
    r(repoStaff.PSC_RUN) = runId
    r(repoStaff.PSC_EMPLOYEE) = CStr(emps(er, repoStaff.EMC_ID))
    r(repoStaff.PSC_BRANCH) = branchId
    r(repoStaff.PSC_PERIOD) = period
    r(repoStaff.PSC_BASE) = fBase
    r(repoStaff.PSC_UNPAID) = unpaid
    r(repoStaff.PSC_TAXABLE_ALLOW) = ta
    r(repoStaff.PSC_NONTAX_ALLOW) = nta
    r(repoStaff.PSC_OVERTIME) = ot
    r(repoStaff.PSC_ADJUST) = adj
    r(repoStaff.PSC_GROSS) = gross
    r(repoStaff.PSC_EESSB) = eeSsb
    r(repoStaff.PSC_TAXABLE) = taxInc
    r(repoStaff.PSC_TAX) = tax
    r(repoStaff.PSC_ADVANCE) = advRec
    r(repoStaff.PSC_NET) = gross - eeSsb - tax - advRec
    r(repoStaff.PSC_ERSSB) = erSsb
    r(repoStaff.PSC_PAYMODE) = CStr(emps(er, repoStaff.EMC_PAYMODE))
    r(repoStaff.PSC_STATUS) = "Draft"
    slip = r
    BuildSlip = True
End Function

' ---- Prepare ----

' True when an active run (Draft, Posted or Paid) already covers this branch for the period.
Private Function Overlaps(ByVal runs As Variant, ByVal period As String, ByVal scope As String) As String
    Dim i As Long, st As String, b As String
    If IsEmpty(runs) Then Exit Function
    For i = 1 To UBound(runs, 1)
        st = CStr(runs(i, repoStaff.PRC_STATUS))
        If CStr(runs(i, repoStaff.PRC_PERIOD)) = period And (st = "Draft" Or st = "Posted" Or st = "Paid") Then
            b = CStr(runs(i, repoStaff.PRC_BRANCH))
            If b = "ALL" Or scope = "ALL" Or b = scope Then
                Overlaps = "Payroll run " & CStr(runs(i, repoStaff.PRC_ID)) & " (" & b & ", " & st & ") already covers " & period & "."
                Exit Function
            End If
        End If
    Next i
End Function

' Control: C25 period (YYYYMM), C8 one branch (blank = ALL). Creates a Draft run, Draft payslips and a Payroll approval. No ledger yet.
Public Function PreparePayroll(ByVal user As clsUser, ByVal inp As Object) As clsResult
    Dim uow As clsUnitOfWork, msg As String, period As String, first As Date, last As Date, scope As String, emps As Variant, hist As Variant, runs As Variant
    Dim scales As Variant, rules As Variant, brackets As Variant, att As Variant, adv As Variant, i As Long, hd As Date, ex As Date, b As String, pos As String, g As Long
    Dim slip As Variant, runId As String, rows As Collection, n As Long, tg As Variant, tee As Variant, ttax As Variant, tadv As Variant, tnet As Variant, ter As Variant, args As Object, r As clsResult
    Dim missing As String
    msg = svcSecurity.Denied(user, svcSecurity.ACT_PAYROLL_PREP)
    If Len(msg) = 0 Then msg = NeedTables()
    If Len(msg) > 0 Then Set PreparePayroll = modMessages.FailResult(msg): Exit Function
    period = CStr(inp(modConstants.INP_PERIOD))
    If Not modDateUtils.TryParsePeriod(period, first) Then Set PreparePayroll = modMessages.FailResult("Period must be YYYYMM (Control C25)."): Exit Function
    last = modDateUtils.EndOfMonth(first)
    Set uow = New clsUnitOfWork
    uow.Init repoConfig.BusinessDate(), user.UserID
    If period > modDateUtils.PeriodOf(uow.BusinessDate) Then Set PreparePayroll = modMessages.FailResult("Period " & period & " is later than the month of the business date."): Exit Function
    scope = Trim$(CStr(inp(modConstants.INP_BRANCH)))
    If Len(scope) = 0 Then scope = "ALL"
    If scope <> "ALL" And Not uow.Branches().Exists(scope) Then Set PreparePayroll = modMessages.FailResult(modMessages.UnknownBranch(scope)): Exit Function
    runs = uow.WorkingTable(modConstants.TBL_PAYROLLRUNS)
    msg = Overlaps(runs, period, scope)
    If Len(msg) > 0 Then Set PreparePayroll = modMessages.FailResult(msg): Exit Function
    emps = uow.WorkingTable(modConstants.TBL_EMPLOYEES)
    If IsEmpty(emps) Then Set PreparePayroll = modMessages.FailResult("There are no employees."): Exit Function
    hist = uow.WorkingTable(modConstants.TBL_EMPHISTORY)
    scales = modSheetUtils.ReadTable(modConstants.TBL_SALARYSCALES)
    rules = modSheetUtils.ReadTable(modConstants.TBL_ALLOWANCERULES)
    brackets = modSheetUtils.ReadTable(modConstants.TBL_TAXBRACKETS)
    att = modSheetUtils.ReadTable(modConstants.TBL_ATTENDANCE)
    adv = uow.WorkingTable(modConstants.TBL_ADVANCES)
    runId = uow.NewId(modConstants.TBL_PAYROLLRUNS, "PayrollRunID", "PR", 6)
    Set rows = New Collection
    tg = CDec(0)
    tee = CDec(0)
    ttax = CDec(0)
    tadv = CDec(0)
    tnet = CDec(0)
    ter = CDec(0)
    For i = 1 To UBound(emps, 1)
        If CStr(emps(i, repoStaff.EMC_STATUS)) <> "Terminated" Or modDateUtils.TryToDate(emps(i, repoStaff.EMC_EXIT), ex) Then
            If modDateUtils.TryToDate(emps(i, repoStaff.EMC_HIRE), hd) Then
                If hd <= last Then
                    If Not modDateUtils.TryToDate(emps(i, repoStaff.EMC_EXIT), ex) Then ex = last
                    If ex >= first Then
                        repoStaff.StateAt hist, emps, i, last, b, pos, g
                        If scope = "ALL" Or scope = b Then
                            If Not uow.Branches().Exists(b) Then Set PreparePayroll = modMessages.FailResult("Employee " & CStr(emps(i, 1)) & " works at an unknown branch."): Exit Function
                            If Not BuildSlip(uow.NewId(modConstants.TBL_PAYSLIPS, "PayslipID", "PS", 7), runId, period, first, last, emps, i, b, pos, g, uow.Branches()(b).BranchType, scales, rules, brackets, att, adv, slip, msg) Then
                                missing = missing & CStr(emps(i, 1)) & " (" & msg & ") "
                            Else
                                rows.Add slip
                                n = n + 1
                                tg = tg + slip(repoStaff.PSC_GROSS)
                                tee = tee + slip(repoStaff.PSC_EESSB)
                                ttax = ttax + slip(repoStaff.PSC_TAX)
                                tadv = tadv + slip(repoStaff.PSC_ADVANCE)
                                tnet = tnet + slip(repoStaff.PSC_NET)
                                ter = ter + slip(repoStaff.PSC_ERSSB)
                            End If
                        End If
                    End If
                End If
            End If
        End If
    Next i
    If Len(missing) > 0 Then Set PreparePayroll = modMessages.FailResult("Cannot prepare payroll: " & missing): Exit Function
    If n = 0 Then Set PreparePayroll = modMessages.FailResult("No employee is employed in " & period & " for " & scope & "."): Exit Function
    For i = 1 To rows.Count
        uow.AppendRow modConstants.TBL_PAYSLIPS, rows(i)
    Next i
    uow.AppendRow modConstants.TBL_PAYROLLRUNS, Array(runId, period, scope, uow.BusinessDate, user.UserID, Empty, n, tg, tee, ttax, tadv, tnet, ter, "Draft", Empty, Empty)
    Set args = modArgs.NewArgs()
    args("run") = runId
    Set r = svcApproval.Request(uow, modConstants.REQ_PAYROLL, "", "", tnet, modConstants.HEAD_OFFICE_ID, args, user.UserID)
    uow.Commit
    Set PreparePayroll = modMessages.OkResult("Payroll run " & runId & " prepared for " & period & " (" & scope & "): " & n & " payslips, gross " & modMoney.FormatAmount(tg) & ", income tax " & modMoney.FormatAmount(ttax) & _
        ", net pay " & modMoney.FormatAmount(tnet) & ". " & r.Message, n, runId)
End Function

' ---- Posting (after the Payroll approval) ----

' Per branch: Dr 5100 gross, Dr 5110 employer SSB; Cr 2400 net, Cr 2230 (tax + employee SSB + employer SSB), Cr 1810 advance recovered.
Public Function ExecPostPayroll(ByVal uow As clsUnitOfWork, ByVal args As Object, ByVal createdBy As String, ByVal approvedBy As String) As clsResult
    Dim runs As Variant, slips As Variant, adv As Variant, rr As Long, i As Long, j As Long, per As Object, b As Variant, v As Variant, lines As Collection, jr As clsResult, grp As String
    Dim runId As String, k As Long, a As Variant
    runs = uow.WorkingTable(modConstants.TBL_PAYROLLRUNS)
    runId = modArgs.Text(args, "run")
    rr = repoStaff.FindRowById(runs, repoStaff.PRC_ID, runId)
    If rr = 0 Then Set ExecPostPayroll = modMessages.FailResult("Unknown payroll run " & runId & "."): Exit Function
    If CStr(runs(rr, repoStaff.PRC_STATUS)) <> "Draft" Then Set ExecPostPayroll = modMessages.FailResult("Run " & runId & " is " & CStr(runs(rr, repoStaff.PRC_STATUS)) & ", not Draft."): Exit Function
    slips = uow.WorkingTable(modConstants.TBL_PAYSLIPS)
    adv = uow.WorkingTable(modConstants.TBL_ADVANCES)
    Set per = CreateObject("Scripting.Dictionary")
    For i = 1 To UBound(slips, 1)
        If CStr(slips(i, repoStaff.PSC_RUN)) = runId And CStr(slips(i, repoStaff.PSC_STATUS)) = "Draft" Then
            b = CStr(slips(i, repoStaff.PSC_BRANCH))
            If Not per.Exists(b) Then per(b) = Array(CDec(0), CDec(0), CDec(0), CDec(0), CDec(0))
            v = per(b)
            v(0) = v(0) + modMoney.ToDec(slips(i, repoStaff.PSC_GROSS))
            v(1) = v(1) + modMoney.ToDec(slips(i, repoStaff.PSC_ERSSB))
            v(2) = v(2) + modMoney.ToDec(slips(i, repoStaff.PSC_NET))
            v(3) = v(3) + modMoney.ToDec(slips(i, repoStaff.PSC_TAX)) + modMoney.ToDec(slips(i, repoStaff.PSC_EESSB)) + modMoney.ToDec(slips(i, repoStaff.PSC_ERSSB))
            v(4) = v(4) + modMoney.ToDec(slips(i, repoStaff.PSC_ADVANCE))
            per(b) = v
            If modMoney.ToDec(slips(i, repoStaff.PSC_ADVANCE)) > 0 Then
                k = 0
                If Not IsEmpty(adv) Then
                    For j = 1 To UBound(adv, 1)
                        If CStr(adv(j, repoStaff.ADC_EMPLOYEE)) = CStr(slips(i, repoStaff.PSC_EMPLOYEE)) And CStr(adv(j, repoStaff.ADC_STATUS)) = "Active" Then k = j
                    Next j
                End If
                If k = 0 Then Set ExecPostPayroll = modMessages.FailResult("The staff advance of " & CStr(slips(i, repoStaff.PSC_EMPLOYEE)) & " changed. Prepare the payroll again."): Exit Function
                If modMoney.ToDec(adv(k, repoStaff.ADC_OUTSTANDING)) < modMoney.ToDec(slips(i, repoStaff.PSC_ADVANCE)) Then Set ExecPostPayroll = modMessages.FailResult("The staff advance of " & CStr(slips(i, repoStaff.PSC_EMPLOYEE)) & " is too small. Prepare the payroll again."): Exit Function
                adv(k, repoStaff.ADC_OUTSTANDING) = modMoney.ToDec(adv(k, repoStaff.ADC_OUTSTANDING)) - modMoney.ToDec(slips(i, repoStaff.PSC_ADVANCE))
                If modMoney.ToDec(adv(k, repoStaff.ADC_OUTSTANDING)) = 0 Then adv(k, repoStaff.ADC_STATUS) = "Cleared"
            End If
            slips(i, repoStaff.PSC_STATUS) = "Posted"
        End If
    Next i
    grp = uow.NewGroupId()
    For Each b In per.Keys
        v = per(b)
        Set lines = New Collection
        svcLedger.AddDr lines, CStr(b), modConstants.GL_SALARIES, v(0)
        svcLedger.AddDr lines, CStr(b), modConstants.GL_EMPLOYER_SSB, v(1)
        svcLedger.AddCr lines, CStr(b), modConstants.GL_STAFF_PAYABLE, v(2)
        svcLedger.AddCr lines, CStr(b), modConstants.GL_PAYROLL_TAX, v(3)
        svcLedger.AddCr lines, CStr(b), modConstants.GL_STAFF_ADVANCES, v(4)
        Set jr = svcLedger.PostJournal(uow, modConstants.JT_PAYROLL, "Payroll " & runId & " " & CStr(b), grp, lines)
        If Not jr.Success Then Set ExecPostPayroll = jr: Exit Function
    Next b
    uow.ReplaceTable modConstants.TBL_PAYSLIPS, slips
    If Not IsEmpty(adv) Then uow.ReplaceTable modConstants.TBL_ADVANCES, adv
    runs(rr, repoStaff.PRC_STATUS) = "Posted"
    runs(rr, repoStaff.PRC_APPROVEDBY) = approvedBy
    runs(rr, repoStaff.PRC_POSTGROUP) = grp
    uow.ReplaceTable modConstants.TBL_PAYROLLRUNS, runs
    repoAudit.Add uow, "PayrollPosted", runId, "Posted as group " & grp & ", approved by " & approvedBy
    Set ExecPostPayroll = modMessages.OkResult("Payroll " & runId & " posted: net pay " & modMoney.FormatAmount(modMoney.ToDec(runs(rr, repoStaff.PRC_NET))) & " MMK is owed to staff (GL 2400). Group " & grp & ".", 1, grp)
End Function

' ---- Payment ----

' Control: C34 = a Posted run. HeadOffice requests, another HeadOffice approves.
Public Function PayPayroll(ByVal user As clsUser, ByVal inp As Object) As clsResult
    Dim uow As clsUnitOfWork, msg As String, runs As Variant, rr As Long, args As Object, r As clsResult
    msg = svcSecurity.Denied(user, svcSecurity.ACT_PAYROLL_PAY)
    If Len(msg) = 0 Then msg = NeedTables()
    If Len(msg) > 0 Then Set PayPayroll = modMessages.FailResult(msg): Exit Function
    Set uow = New clsUnitOfWork
    uow.Init repoConfig.BusinessDate(), user.UserID
    runs = uow.WorkingTable(modConstants.TBL_PAYROLLRUNS)
    rr = repoStaff.FindRowById(runs, repoStaff.PRC_ID, CStr(inp(modConstants.INP_PAYRUN)))
    If rr = 0 Then Set PayPayroll = modMessages.FailResult("Unknown payroll run. Put the run ID in Control C34."): Exit Function
    If CStr(runs(rr, repoStaff.PRC_STATUS)) <> "Posted" Then Set PayPayroll = modMessages.FailResult("Run " & CStr(runs(rr, repoStaff.PRC_ID)) & " is " & CStr(runs(rr, repoStaff.PRC_STATUS)) & ". Only a Posted run can be paid."): Exit Function
    Set args = modArgs.NewArgs()
    args("run") = CStr(runs(rr, repoStaff.PRC_ID))
    Set r = svcApproval.Request(uow, modConstants.REQ_PAY_PAYROLL, "", "", modMoney.ToDec(runs(rr, repoStaff.PRC_NET)), modConstants.HEAD_OFFICE_ID, args, user.UserID)
    uow.Commit
    Set PayPayroll = r
End Function

' Account mode: B: Dr 2400, (Cr 1900 / A: Dr 1900 when the account is in another branch), A: Cr deposit GL. External mode: B: Dr 2400, Cr 1900; HO: Dr 1900, Cr 1100.
Public Function ExecPayPayroll(ByVal uow As clsUnitOfWork, ByVal args As Object, ByVal createdBy As String, ByVal approvedBy As String) As clsResult
    Dim runs As Variant, slips As Variant, emps As Variant, rr As Long, i As Long, er As Long, a As clsAccount, lines As Collection, jr As clsResult, grp As String, runId As String
    Dim ext As Object, b As Variant, net As Variant, total As Variant, ho As String, acctSlips As Collection, s As Variant, mode As String
    ho = modConstants.HEAD_OFFICE_ID
    runs = uow.WorkingTable(modConstants.TBL_PAYROLLRUNS)
    runId = modArgs.Text(args, "run")
    rr = repoStaff.FindRowById(runs, repoStaff.PRC_ID, runId)
    If rr = 0 Then Set ExecPayPayroll = modMessages.FailResult("Unknown payroll run."): Exit Function
    If CStr(runs(rr, repoStaff.PRC_STATUS)) <> "Posted" Then Set ExecPayPayroll = modMessages.FailResult("Run " & runId & " is not Posted."): Exit Function
    slips = uow.WorkingTable(modConstants.TBL_PAYSLIPS)
    emps = uow.WorkingTable(modConstants.TBL_EMPLOYEES)
    Set ext = CreateObject("Scripting.Dictionary")
    Set lines = New Collection
    Set acctSlips = New Collection
    total = CDec(0)
    For i = 1 To UBound(slips, 1)
        If CStr(slips(i, repoStaff.PSC_RUN)) = runId And CStr(slips(i, repoStaff.PSC_STATUS)) = "Posted" Then
            b = CStr(slips(i, repoStaff.PSC_BRANCH))
            net = modMoney.ToDec(slips(i, repoStaff.PSC_NET))
            mode = "External"
            Set a = Nothing
            If CStr(slips(i, repoStaff.PSC_PAYMODE)) = "Account" Then
                er = repoStaff.FindRowById(emps, repoStaff.EMC_ID, CStr(slips(i, repoStaff.PSC_EMPLOYEE)))
                If er > 0 Then Set a = uow.GetAccount(CStr(emps(er, repoStaff.EMC_ACCOUNT)))
                If Not a Is Nothing Then
                    If a.Status = modConstants.STATUS_ACTIVE And a.Ccy = modConstants.CCY_MMK Then mode = "Account"
                End If
                If mode = "External" Then clsLogger.Warn "svcPayroll.ExecPayPayroll", "Salary account of " & CStr(slips(i, repoStaff.PSC_EMPLOYEE)) & " is not usable. Paid externally."
            End If
            If mode = "Account" Then
                svcLedger.AddDr lines, CStr(b), modConstants.GL_STAFF_PAYABLE, net
                If a.BranchID <> CStr(b) Then
                    svcLedger.AddCr lines, CStr(b), modConstants.GL_INTERBRANCH, net
                    svcLedger.AddDr lines, a.BranchID, modConstants.GL_INTERBRANCH, net
                End If
                svcLedger.AddCr lines, a.BranchID, a.DepositGL, net, a.AccountNo
                acctSlips.Add Array(a.AccountNo, net, CStr(b), CStr(slips(i, repoStaff.PSC_EMPLOYEE)))
            Else
                If ext.Exists(b) Then ext(b) = ext(b) + net Else ext(b) = net
                slips(i, repoStaff.PSC_PAYMODE) = "External"
            End If
            slips(i, repoStaff.PSC_STATUS) = "Paid"
        End If
    Next i
    For Each b In ext.Keys
        If CStr(b) = ho Then
            svcLedger.AddDr lines, ho, modConstants.GL_STAFF_PAYABLE, ext(b)
            svcLedger.AddCr lines, ho, modConstants.GL_TREASURY, ext(b)
        Else
            svcLedger.AddDr lines, CStr(b), modConstants.GL_STAFF_PAYABLE, ext(b)
            svcLedger.AddCr lines, CStr(b), modConstants.GL_INTERBRANCH, ext(b)
            total = total + ext(b)
        End If
    Next b
    svcLedger.AddDr lines, ho, modConstants.GL_INTERBRANCH, total
    svcLedger.AddCr lines, ho, modConstants.GL_TREASURY, total
    grp = uow.NewGroupId()
    Set jr = svcLedger.PostJournal(uow, modConstants.JT_PAYROLL_PAYMENT, "Salary payment " & runId, grp, lines)
    If Not jr.Success Then Set ExecPayPayroll = jr: Exit Function
    For Each s In acctSlips
        svcAccounts.Effect uow, uow.GetAccount(CStr(s(0))), modConstants.TXN_SALARY, modConstants.DIR_CREDIT, s(1), "", CStr(s(2)), "Salary " & runId, runId, jr.RefID, grp, createdBy, approvedBy, False
    Next s
    uow.ReplaceTable modConstants.TBL_PAYSLIPS, slips
    runs(rr, repoStaff.PRC_STATUS) = "Paid"
    runs(rr, repoStaff.PRC_PAYGROUP) = grp
    uow.ReplaceTable modConstants.TBL_PAYROLLRUNS, runs
    repoAudit.Add uow, "PayrollPaid", runId, "Paid as group " & grp & ", approved by " & approvedBy
    Set ExecPayPayroll = modMessages.OkResult("Payroll " & runId & " paid: " & acctSlips.Count & " salaries credited to accounts, the rest paid through head office. Group " & grp & ".", 1, grp)
End Function

' ---- Reversal hooks (called by svcLedger.ReverseGroup through svcApprovalExtra) ----

' A Paid run is never reversed (correct it next month with an attendance Adjustment).
Public Function PayrollReversalProblem(ByVal uow As clsUnitOfWork, ByVal groupId As String) As String
    Dim runs As Variant, i As Long
    If Not modSheetUtils.TableExists(modConstants.TBL_PAYROLLRUNS) Then Exit Function
    runs = uow.WorkingTable(modConstants.TBL_PAYROLLRUNS)
    If IsEmpty(runs) Then Exit Function
    For i = 1 To UBound(runs, 1)
        If CStr(runs(i, repoStaff.PRC_PAYGROUP)) = groupId Then PayrollReversalProblem = "A payroll payment cannot be reversed. Correct it next month with an attendance Adjustment.": Exit Function
        If CStr(runs(i, repoStaff.PRC_POSTGROUP)) = groupId And CStr(runs(i, repoStaff.PRC_STATUS)) = "Paid" Then PayrollReversalProblem = "A Paid payroll run cannot be reversed.": Exit Function
    Next i
End Function

' After the posting journals of a run were reversed: payslips and run become Reversed and the advance balances come back.
Public Sub AfterReversal(ByVal uow As clsUnitOfWork, ByVal groupId As String)
    Dim runs As Variant, slips As Variant, adv As Variant, i As Long, j As Long, rr As Long, runId As String
    If Not modSheetUtils.TableExists(modConstants.TBL_PAYROLLRUNS) Then Exit Sub
    runs = uow.WorkingTable(modConstants.TBL_PAYROLLRUNS)
    If IsEmpty(runs) Then Exit Sub
    For rr = 1 To UBound(runs, 1)
        If CStr(runs(rr, repoStaff.PRC_POSTGROUP)) = groupId And CStr(runs(rr, repoStaff.PRC_STATUS)) = "Posted" Then
            runId = CStr(runs(rr, repoStaff.PRC_ID))
            slips = uow.WorkingTable(modConstants.TBL_PAYSLIPS)
            adv = uow.WorkingTable(modConstants.TBL_ADVANCES)
            For i = 1 To UBound(slips, 1)
                If CStr(slips(i, repoStaff.PSC_RUN)) = runId And CStr(slips(i, repoStaff.PSC_STATUS)) = "Posted" Then
                    slips(i, repoStaff.PSC_STATUS) = "Reversed"
                    If modMoney.ToDec(slips(i, repoStaff.PSC_ADVANCE)) > 0 And Not IsEmpty(adv) Then
                        For j = UBound(adv, 1) To 1 Step -1
                            If CStr(adv(j, repoStaff.ADC_EMPLOYEE)) = CStr(slips(i, repoStaff.PSC_EMPLOYEE)) Then
                                adv(j, repoStaff.ADC_OUTSTANDING) = modMoney.ToDec(adv(j, repoStaff.ADC_OUTSTANDING)) + modMoney.ToDec(slips(i, repoStaff.PSC_ADVANCE))
                                adv(j, repoStaff.ADC_STATUS) = "Active"
                                Exit For
                            End If
                        Next j
                    End If
                End If
            Next i
            uow.ReplaceTable modConstants.TBL_PAYSLIPS, slips
            If Not IsEmpty(adv) Then uow.ReplaceTable modConstants.TBL_ADVANCES, adv
            runs(rr, repoStaff.PRC_STATUS) = "Reversed"
            uow.ReplaceTable modConstants.TBL_PAYROLLRUNS, runs
            Exit Sub
        End If
    Next rr
End Sub

' ---- Staff advances ----

' args: employee, amount, monthly. Up to PettyCashLimit from the vault (B: Dr 1810, Cr 1000), above it through head office (Cr 1900; HO: Dr 1900, Cr 1100).
Public Function ExecAdvance(ByVal uow As clsUnitOfWork, ByVal args As Object, ByVal createdBy As String, ByVal approvedBy As String) As clsResult
    Dim emps As Variant, adv As Variant, er As Long, i As Long, b As String, amt As Variant, monthly As Variant, lines As Collection, jr As clsResult, grp As String, id As String, ho As String
    ho = modConstants.HEAD_OFFICE_ID
    amt = modArgs.Dec(args, "amount")
    monthly = modArgs.Dec(args, "monthly")
    emps = uow.WorkingTable(modConstants.TBL_EMPLOYEES)
    er = repoStaff.FindRowById(emps, repoStaff.EMC_ID, modArgs.Text(args, "employee"))
    If er = 0 Then Set ExecAdvance = modMessages.FailResult("Unknown employee."): Exit Function
    If CStr(emps(er, repoStaff.EMC_STATUS)) <> "Active" Then Set ExecAdvance = modMessages.FailResult("Only an Active employee can get an advance."): Exit Function
    adv = uow.WorkingTable(modConstants.TBL_ADVANCES)
    If Not IsEmpty(adv) Then
        For i = 1 To UBound(adv, 1)
            If CStr(adv(i, repoStaff.ADC_EMPLOYEE)) = CStr(emps(er, 1)) And CStr(adv(i, repoStaff.ADC_STATUS)) = "Active" Then Set ExecAdvance = modMessages.FailResult("The employee already has an Active advance."): Exit Function
        Next i
    End If
    If monthly <= 0 Or monthly > amt Then Set ExecAdvance = modMessages.FailResult("The monthly recovery (Control C24) must be between 1 and the advance amount."): Exit Function
    b = CStr(emps(er, repoStaff.EMC_BRANCH))
    Set lines = New Collection
    svcLedger.AddDr lines, b, modConstants.GL_STAFF_ADVANCES, amt
    If amt <= repoConfig.GetDec(modConstants.CFG_PETTY_LIMIT) Then
        If uow.GLNet(b, modConstants.GL_VAULT) - amt < 0 Then Set ExecAdvance = modMessages.FailResult(modMessages.MSG_VAULT_SHORT): Exit Function
        svcLedger.AddCr lines, b, modConstants.GL_VAULT, amt
    ElseIf b = ho Then
        svcLedger.AddCr lines, ho, modConstants.GL_TREASURY, amt
    Else
        svcLedger.AddCr lines, b, modConstants.GL_INTERBRANCH, amt
        svcLedger.AddDr lines, ho, modConstants.GL_INTERBRANCH, amt
        svcLedger.AddCr lines, ho, modConstants.GL_TREASURY, amt
    End If
    grp = uow.NewGroupId()
    Set jr = svcLedger.PostJournal(uow, modConstants.JT_STAFF_ADVANCE, "Staff advance " & CStr(emps(er, 1)), grp, lines)
    If Not jr.Success Then Set ExecAdvance = jr: Exit Function
    id = uow.NewId(modConstants.TBL_ADVANCES, "AdvanceID", "AD", 6)
    uow.AppendRow modConstants.TBL_ADVANCES, Array(id, CStr(emps(er, 1)), uow.BusinessDate, amt, monthly, amt, "Active", jr.RefID)
    repoAudit.Add uow, "StaffAdvance", id, CStr(emps(er, 1)) & " " & modMoney.FormatAmount(amt) & ", approved by " & approvedBy
    Set ExecAdvance = modMessages.OkResult("Advance " & id & " of " & modMoney.FormatAmount(amt) & " MMK issued to " & CStr(emps(er, 1)) & ". Recovery " & modMoney.FormatAmount(monthly) & " a month. Group " & grp & ".", 1, grp)
End Function

' Control: C30 employee, C7 amount, C24 monthly recovery. HeadOffice requests, another HeadOffice approves.
Public Function IssueAdvance(ByVal user As clsUser, ByVal inp As Object) As clsResult
    Dim uow As clsUnitOfWork, msg As String, amt As Variant, monthly As Variant, emps As Variant, er As Long, args As Object, r As clsResult
    msg = svcSecurity.Denied(user, svcSecurity.ACT_ADVANCE)
    If Len(msg) = 0 Then msg = NeedTables()
    If Len(msg) > 0 Then Set IssueAdvance = modMessages.FailResult(msg): Exit Function
    If Not modArgs.ParseAmount(inp(modConstants.INP_AMOUNT), amt, msg) Then Set IssueAdvance = modMessages.FailResult(msg): Exit Function
    If Not modMoney.TryToDec(inp(modConstants.INP_RESIDUAL), monthly) Then Set IssueAdvance = modMessages.FailResult("Enter the monthly recovery in Control C24."): Exit Function
    Set uow = New clsUnitOfWork
    uow.Init repoConfig.BusinessDate(), user.UserID
    emps = uow.WorkingTable(modConstants.TBL_EMPLOYEES)
    er = repoStaff.FindRowById(emps, repoStaff.EMC_ID, CStr(inp(modConstants.INP_EMPLOYEE)))
    If er = 0 Then Set IssueAdvance = modMessages.FailResult("Unknown employee. Put the EmployeeID in Control C30."): Exit Function
    Set args = modArgs.NewArgs()
    args("employee") = CStr(emps(er, 1))
    args("amount") = amt
    args("monthly") = monthly
    Set r = svcApproval.Request(uow, modConstants.REQ_STAFF_ADVANCE, "", "", amt, CStr(emps(er, repoStaff.EMC_BRANCH)), args, user.UserID)
    uow.Commit
    Set IssueAdvance = r
End Function
