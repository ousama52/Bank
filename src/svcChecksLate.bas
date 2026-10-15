Attribute VB_Name = "svcChecksLate"
Option Explicit

' Control checks of Stages 5, 7 and 8: 17, 18 (payroll), 24, 25 (investments), 26, 27, 28, 29 (government rules).
' Called by svcChecksExtra. A check is skipped when its tables do not exist yet.

Private Function Net(ByVal d As Object, ByVal gl As String) As Variant
    If d.Exists(gl) Then Net = d(gl) Else Net = CDec(0)
End Function

Private Function Verdict(ByVal ok As Boolean) As String
    If ok Then Verdict = "PASS" Else Verdict = "FAIL"
End Function

Private Function Cmp(ByVal a As Variant, ByVal b As Variant) As String
    Cmp = "GL " & modMoney.FormatAmount(a) & " / records " & modMoney.FormatAmount(b)
End Function

' Sum of (credit - debit) of one GL over the journals of one type that are still Posted.
Private Function CreditNetOfType(ByVal lns As Variant, ByVal hdr As Variant, ByVal journalType As String, ByVal glCode As String) As Variant
    Dim ok As Object, i As Long, total As Variant
    Set ok = CreateObject("Scripting.Dictionary")
    If Not IsEmpty(hdr) Then
        For i = 1 To UBound(hdr, 1)
            If CStr(hdr(i, repoJournal.JNC_TYPE)) = journalType And CStr(hdr(i, repoJournal.JNC_STATUS)) = modConstants.STATUS_POSTED Then ok(CStr(hdr(i, repoJournal.JNC_ID))) = True
        Next i
    End If
    total = CDec(0)
    If Not IsEmpty(lns) Then
        For i = 1 To UBound(lns, 1)
            If CStr(lns(i, repoJournal.LNC_GL)) = glCode Then
                If ok.Exists(CStr(lns(i, repoJournal.LNC_JOURNAL))) Then total = total + modMoney.ToDec(lns(i, repoJournal.LNC_CREDIT)) - modMoney.ToDec(lns(i, repoJournal.LNC_DEBIT))
            End If
        Next i
    End If
    CreditNetOfType = total
End Function

' ---- Stage 5: checks 17 and 18 ----

Private Sub CheckPayroll(ByVal rowsList As Collection, ByVal gl As Object, ByVal lns As Variant)
    Dim runs As Variant, slips As Variant, adv As Variant, i As Long, j As Long, seen As Object, key As String, dup As Long, unpaid As Variant, bad As Long, st As String
    Dim sg As Variant, sn As Variant, outst As Variant, runCredit As Variant, hdr As Variant, rid As String
    If Not modSheetUtils.TableExists(modConstants.TBL_PAYROLLRUNS) Or Not modSheetUtils.TableExists(modConstants.TBL_PAYSLIPS) Then Exit Sub
    runs = modSheetUtils.ReadTable(modConstants.TBL_PAYROLLRUNS)
    slips = modSheetUtils.ReadTable(modConstants.TBL_PAYSLIPS)
    Set seen = CreateObject("Scripting.Dictionary")
    unpaid = CDec(0)
    runCredit = CDec(0)
    If Not IsEmpty(slips) Then
        For i = 1 To UBound(slips, 1)
            st = CStr(slips(i, repoStaff.PSC_STATUS))
            If st = "Posted" Or st = "Paid" Then
                key = CStr(slips(i, repoStaff.PSC_EMPLOYEE)) & "|" & CStr(slips(i, repoStaff.PSC_PERIOD))
                If seen.Exists(key) Then dup = dup + 1 Else seen(key) = True
                If st = "Posted" Then unpaid = unpaid + modMoney.ToDec(slips(i, repoStaff.PSC_NET))
            End If
        Next i
    End If
    If Not IsEmpty(runs) Then
        For i = 1 To UBound(runs, 1)
            st = CStr(runs(i, repoStaff.PRC_STATUS))
            If st = "Posted" Or st = "Paid" Then
                rid = CStr(runs(i, repoStaff.PRC_ID))
                sg = CDec(0)
                sn = CDec(0)
                If Not IsEmpty(slips) Then
                    For j = 1 To UBound(slips, 1)
                        If CStr(slips(j, repoStaff.PSC_RUN)) = rid And (CStr(slips(j, repoStaff.PSC_STATUS)) = "Posted" Or CStr(slips(j, repoStaff.PSC_STATUS)) = "Paid") Then
                            sg = sg + modMoney.ToDec(slips(j, repoStaff.PSC_GROSS))
                            sn = sn + modMoney.ToDec(slips(j, repoStaff.PSC_NET))
                        End If
                    Next j
                End If
                If sg <> modMoney.ToDec(runs(i, repoStaff.PRC_GROSS)) Or sn <> modMoney.ToDec(runs(i, repoStaff.PRC_NET)) Then bad = bad + 1
                runCredit = runCredit + modMoney.ToDec(runs(i, repoStaff.PRC_TAX)) + modMoney.ToDec(runs(i, repoStaff.PRC_EESSB)) + modMoney.ToDec(runs(i, repoStaff.PRC_ERSSB))
            End If
        Next i
    End If
    svcChecks.AddRow rowsList, 17, "Payslips equal their runs; one payslip per employee and period; GL 2400 = unpaid net pay", Verdict(bad = 0 And dup = 0 And -Net(gl, modConstants.GL_STAFF_PAYABLE) = unpaid), _
        bad & " run(s) differ, " & dup & " duplicate(s); 2400: " & Cmp(-Net(gl, modConstants.GL_STAFF_PAYABLE), unpaid)
    outst = CDec(0)
    If modSheetUtils.TableExists(modConstants.TBL_ADVANCES) Then
        adv = modSheetUtils.ReadTable(modConstants.TBL_ADVANCES)
        If Not IsEmpty(adv) Then
            For i = 1 To UBound(adv, 1)
                If CStr(adv(i, repoStaff.ADC_STATUS)) = "Active" Then outst = outst + modMoney.ToDec(adv(i, repoStaff.ADC_OUTSTANDING))
            Next i
        End If
    End If
    hdr = repoJournal.LoadHeaders()
    svcChecks.AddRow rowsList, 18, "GL 1810 = staff advances outstanding; GL 2230 credits = payroll tax and social security of posted runs", _
        Verdict(Net(gl, modConstants.GL_STAFF_ADVANCES) = outst And CreditNetOfType(lns, hdr, modConstants.JT_PAYROLL, modConstants.GL_PAYROLL_TAX) = runCredit), _
        "1810: " & Cmp(Net(gl, modConstants.GL_STAFF_ADVANCES), outst) & "; 2230 payroll credits: " & Cmp(CreditNetOfType(lns, hdr, modConstants.JT_PAYROLL, modConstants.GL_PAYROLL_TAX), runCredit)
End Sub

' ---- Stage 7: checks 24 and 25 ----

Private Function CouponAsOf() As Date
    If svcChecksFx.StrictEod Then CouponAsOf = repoConfig.BusinessDate() Else CouponAsOf = repoConfig.BusinessDate() - 1
End Function

Private Sub CheckInvestments(ByVal rowsList As Collection, ByVal gl As Object)
    Dim invs As Variant, rets As Variant, i As Long, carry As Variant, unpaid As Variant, st As String, bonds As Variant
    If Not modSheetUtils.TableExists(modConstants.TBL_INVESTMENTS) Or Not modSheetUtils.TableExists(modConstants.TBL_INVRETURNS) Then Exit Sub
    invs = modSheetUtils.ReadTable(modConstants.TBL_INVESTMENTS)
    rets = modSheetUtils.ReadTable(modConstants.TBL_INVRETURNS)
    carry = CDec(0)
    unpaid = CDec(0)
    If Not IsEmpty(invs) Then
        For i = 1 To UBound(invs, 1)
            st = CStr(invs(i, repoInvest.INC_STATUS))
            If st = "Active" Or st = "Matured" Then
                carry = carry + modMoney.ToDec(invs(i, repoInvest.INC_CARRYING))
                unpaid = unpaid + repoInvest.UnpaidRecognized(rets, CStr(invs(i, 1)))
            End If
        Next i
    End If
    svcChecks.AddRow rowsList, 24, "GL 1600 = carrying value of investments not yet repaid", Verdict(Net(gl, modConstants.GL_INVESTMENTS) = carry), Cmp(Net(gl, modConstants.GL_INVESTMENTS), carry)
    If modSheetUtils.TableExists(modConstants.TBL_BONDS) Then
        bonds = modSheetUtils.ReadTable(modConstants.TBL_BONDS)
        unpaid = unpaid + svcBonds.UnpaidCoupons(bonds, CouponAsOf())
    End If
    svcChecks.AddRow rowsList, 25, "GL 1610 = recognized unpaid investment returns and bond coupons", Verdict(Net(gl, modConstants.GL_INVEST_RECEIVABLE) = unpaid), Cmp(Net(gl, modConstants.GL_INVEST_RECEIVABLE), unpaid)
End Sub

' ---- Stage 8: checks 26 to 29 ----

Private Sub CheckGovernment(ByVal rowsList As Collection, ByVal gl As Object)
    Dim bonds As Variant, chart As Object, k As Variant, rw As Variant, i As Long, mapped As Object, miss As String, alerts As Variant, rules As Variant, cases As Variant
    Dim ruleIds As Object, bad As Long, old As Long, dt As Date, d As Date, info As String
    If modSheetUtils.TableExists(modConstants.TBL_BONDS) Then
        bonds = modSheetUtils.ReadTable(modConstants.TBL_BONDS)
        svcChecks.AddRow rowsList, 26, "GL 1620 = carrying value of held government securities", Verdict(Net(gl, modConstants.GL_GOV_SEC) = svcBonds.HeldCarrying(bonds)), Cmp(Net(gl, modConstants.GL_GOV_SEC), svcBonds.HeldCarrying(bonds))
    End If
    If modSheetUtils.TableExists(modConstants.TBL_RISKWEIGHTS) Then
        rw = modSheetUtils.ReadTable(modConstants.TBL_RISKWEIGHTS)
        Set mapped = CreateObject("Scripting.Dictionary")
        If Not IsEmpty(rw) Then
            For i = 1 To UBound(rw, 1)
                mapped(CStr(rw(i, 3))) = True
            Next i
        End If
        Set chart = repoChart.LoadChart()
        For Each k In gl.Keys
            If chart.Exists(CStr(k)) Then
                If chart(CStr(k))(1) = "Asset" And gl(k) <> 0 And Not mapped.Exists(CStr(k)) Then miss = miss & CStr(k) & " "
            End If
        Next k
        svcChecks.AddRow rowsList, 27, "Every asset GL with a balance has a risk weight", Verdict(Len(miss) = 0), IIf(Len(miss) = 0, "All mapped", "Not mapped: " & Trim$(miss))
    End If
    svcYearEnd.CheckClosedYears rowsList
    If modSheetUtils.TableExists(modConstants.TBL_AMLALERTS) And modSheetUtils.TableExists(modConstants.TBL_AMLRULES) And modSheetUtils.TableExists(modConstants.TBL_AMLCASES) Then
        alerts = modSheetUtils.ReadTable(modConstants.TBL_AMLALERTS)
        rules = modSheetUtils.ReadTable(modConstants.TBL_AMLRULES)
        cases = modSheetUtils.ReadTable(modConstants.TBL_AMLCASES)
        Set ruleIds = CreateObject("Scripting.Dictionary")
        If Not IsEmpty(rules) Then
            For i = 1 To UBound(rules, 1)
                ruleIds(CStr(rules(i, 1))) = True
            Next i
        End If
        d = repoConfig.BusinessDate()
        If Not IsEmpty(alerts) Then
            For i = 1 To UBound(alerts, 1)
                If Len(CStr(alerts(i, 1))) > 0 Then
                    If Not ruleIds.Exists(CStr(alerts(i, repoGov.ALC_RULE))) Then bad = bad + 1: info = info & CStr(alerts(i, 1)) & " has no rule; "
                    If CStr(alerts(i, repoGov.ALC_STATUS)) = "New" And modDateUtils.TryToDate(alerts(i, repoGov.ALC_DATE), dt) Then
                        If d - dt > repoConfig.GetLong(modConstants.CFG_AML_DAYS) Then old = old + 1
                    End If
                End If
            Next i
        End If
        If Not IsEmpty(cases) Then
            For i = 1 To UBound(cases, 1)
                If CStr(cases(i, repoGov.CSC_STATUS)) = "Open" And Len(CStr(cases(i, repoGov.CSC_ASSIGNED))) = 0 Then bad = bad + 1: info = info & CStr(cases(i, 1)) & " has no assignee; "
            Next i
        End If
        If bad > 0 Then
            svcChecks.AddRow rowsList, 29, "AML alerts have rules, open cases have an assignee", "FAIL", Left$(info, 300)
        ElseIf old > 0 Then
            svcChecks.AddRow rowsList, 29, "AML alerts have rules, open cases have an assignee", "WARN", old & " alert(s) still New after " & repoConfig.GetLong(modConstants.CFG_AML_DAYS) & " days"
        Else
            svcChecks.AddRow rowsList, 29, "AML alerts have rules, open cases have an assignee", "PASS", "No problem"
        End If
    End If
End Sub

Public Sub RunLateChecks(ByVal rowsList As Collection, ByVal lns As Variant, ByVal gl As Object)
    CheckPayroll rowsList, gl, lns
    CheckInvestments rowsList, gl
    CheckGovernment rowsList, gl
End Sub
