Attribute VB_Name = "svcReportsStaff"
Option Explicit

' Stage 5 reports: Payslip, PayrollSummary, StaffCostByBranch, StaffingGaps, Headcount, AdvanceRegister, SalaryScaleReport.

Private Const SH_PAYSLIP As String = "Payslip"
Private Const SH_SUMMARY As String = "PayrollSummary"
Private Const SH_COST As String = "StaffCostByBranch"
Private Const SH_GAPS As String = "StaffingGaps"
Private Const SH_HEAD As String = "Headcount"
Private Const SH_ADV As String = "AdvanceRegister"
Private Const SH_SCALE As String = "SalaryScaleReport"

Private Sub AddTo(ByVal d As Object, ByVal key As String, ByVal v As Variant)
    If d.Exists(key) Then d(key) = d(key) + v Else d(key) = v
End Sub

Private Function Num(ByVal d As Object, ByVal key As String) As Variant
    If d.Exists(key) Then Num = d(key) Else Num = CDec(0)
End Function

' Active employees per "branch|position".
Private Function Counts(ByVal emps As Variant) As Object
    Dim d As Object, i As Long
    Set d = CreateObject("Scripting.Dictionary")
    If Not IsEmpty(emps) Then
        For i = 1 To UBound(emps, 1)
            If CStr(emps(i, repoStaff.EMC_STATUS)) = "Active" Or CStr(emps(i, repoStaff.EMC_STATUS)) = "OnLeave" Then AddTo d, CStr(emps(i, repoStaff.EMC_BRANCH)) & "|" & CStr(emps(i, repoStaff.EMC_POSITION)), 1
        Next i
    End If
    Set Counts = d
End Function

' Branches with no Active branch manager (used by the end-of-day warning and by StaffingGaps).
Public Function BranchesWithoutManager(ByVal emps As Variant, ByVal branches As Object) As String
    Dim c As Object, b As Variant, bt As clsBranch
    Set c = Counts(emps)
    For Each b In branches.Keys
        Set bt = branches(b)
        If bt.BranchType <> modConstants.BT_HEADOFFICE And Num(c, CStr(b) & "|BM") = 0 Then BranchesWithoutManager = BranchesWithoutManager & CStr(b) & " "
    Next b
End Function

Private Sub BuildGaps(ByVal d As Date, ByVal emps As Variant)
    Dim pos As Variant, branches As Variant, c As Object, i As Long, j As Long, minCol As Long, need As Long, have As Long, rowsA As Collection
    Set c = Counts(emps)
    Set rowsA = New Collection
    pos = modSheetUtils.ReadTable(modConstants.TBL_POSITIONS)
    branches = modSheetUtils.ReadTable(modConstants.TBL_BRANCHES)
    If Not IsEmpty(pos) And Not IsEmpty(branches) Then
        For i = 1 To UBound(branches, 1)
            Select Case CStr(branches(i, 3))
                Case modConstants.BT_HEADOFFICE: minCol = 4
                Case modConstants.BT_CITY: minCol = 5
                Case Else: minCol = 6
            End Select
            For j = 1 To UBound(pos, 1)
                need = CLng(modMoney.ToDec(pos(j, minCol)))
                have = CLng(Num(c, CStr(branches(i, 1)) & "|" & CStr(pos(j, 1))))
                If need <> have Then rowsA.Add Array(CStr(branches(i, 1)), CStr(pos(j, 1)), CStr(pos(j, 2)), need, have, have - need, IIf(have < need, "Missing", "Excess"))
            Next j
        Next i
    End If
    modSheetUtils.PrepareReport SH_GAPS, "1,2"
    modSheetUtils.SetCell SH_GAPS, 1, 1, "Staffing gaps " & modDateUtils.ToIso(d) & " (SIMULATION)", True
    modSheetUtils.WriteGrid SH_GAPS, 3, 1, svcReports.ToGrid(Array("Branch", "Position", "Name", "Guide minimum", "Active staff", "Difference", "Result"), rowsA)
    modSheetUtils.StyleReport SH_GAPS, 3, 7, ""
End Sub

Private Sub BuildHeadcount(ByVal d As Date, ByVal emps As Variant)
    Dim pos As Variant, cat As Object, i As Long, bt As Object, branches As Object, key As String, rowsA As Collection, k As Variant, parts() As String, pr As Long
    Dim cnt As Object
    Set cnt = CreateObject("Scripting.Dictionary")
    Set branches = repoBranch.LoadBranches()
    Set cat = CreateObject("Scripting.Dictionary")
    pos = modSheetUtils.ReadTable(modConstants.TBL_POSITIONS)
    If Not IsEmpty(pos) Then
        For i = 1 To UBound(pos, 1)
            cat(CStr(pos(i, 1))) = CStr(pos(i, 3))
        Next i
    End If
    If Not IsEmpty(emps) Then
        For i = 1 To UBound(emps, 1)
            If CStr(emps(i, repoStaff.EMC_STATUS)) = "Active" Or CStr(emps(i, repoStaff.EMC_STATUS)) = "OnLeave" Then
                key = ""
                If branches.Exists(CStr(emps(i, repoStaff.EMC_BRANCH))) Then key = branches(CStr(emps(i, repoStaff.EMC_BRANCH))).BranchType
                key = key & "|" & IIf(cat.Exists(CStr(emps(i, repoStaff.EMC_POSITION))), cat(CStr(emps(i, repoStaff.EMC_POSITION))), "?")
                AddTo cnt, key, 1
            End If
        Next i
    End If
    Set rowsA = New Collection
    For Each k In cnt.Keys
        parts = Split(CStr(k), "|")
        rowsA.Add Array(parts(0), parts(1), cnt(k))
    Next k
    modSheetUtils.PrepareReport SH_HEAD, "1,2"
    modSheetUtils.SetCell SH_HEAD, 1, 1, "Headcount " & modDateUtils.ToIso(d) & " (SIMULATION)", True
    modSheetUtils.WriteGrid SH_HEAD, 3, 1, svcReports.ToGrid(Array("Branch type", "Category", "Employees"), rowsA)
    modSheetUtils.StyleReport SH_HEAD, 3, 3, ""
End Sub

' Latest non-reversed payslip of one employee and period (Control C30 and C25).
Private Sub BuildPayslip(ByVal d As Date, ByVal emps As Variant, ByVal slips As Variant, ByVal inp As Object)
    Dim id As String, period As String, i As Long, found As Long, er As Long, rowsA As Collection
    id = CStr(inp(modConstants.INP_EMPLOYEE))
    period = CStr(inp(modConstants.INP_PERIOD))
    modSheetUtils.PrepareReport SH_PAYSLIP, "1"
    If Not IsEmpty(slips) Then
        For i = 1 To UBound(slips, 1)
            If CStr(slips(i, repoStaff.PSC_EMPLOYEE)) = id And CStr(slips(i, repoStaff.PSC_PERIOD)) = period And CStr(slips(i, repoStaff.PSC_STATUS)) <> "Reversed" Then found = i
        Next i
    End If
    If found = 0 Then
        modSheetUtils.SetCell SH_PAYSLIP, 1, 1, "Payslip: put the EmployeeID in Control C30 and the period in C25, prepare payroll, then run the reports again.", True
        Exit Sub
    End If
    er = repoStaff.FindRowById(emps, repoStaff.EMC_ID, id)
    Set rowsA = New Collection
    rowsA.Add Array("Employee", id & " " & CStr(emps(er, repoStaff.EMC_NAME)))
    rowsA.Add Array("Position and grade", CStr(emps(er, repoStaff.EMC_POSITION)) & " grade " & CStr(emps(er, repoStaff.EMC_GRADE)))
    rowsA.Add Array("Branch", CStr(slips(found, repoStaff.PSC_BRANCH)))
    rowsA.Add Array("Period", period & " (run " & CStr(slips(found, repoStaff.PSC_RUN)) & ", " & CStr(slips(found, repoStaff.PSC_STATUS)) & ")")
    rowsA.Add Array("Base salary", modMoney.ToDec(slips(found, repoStaff.PSC_BASE)))
    rowsA.Add Array("Unpaid leave deduction", -modMoney.ToDec(slips(found, repoStaff.PSC_UNPAID)))
    rowsA.Add Array("Taxable allowances", modMoney.ToDec(slips(found, repoStaff.PSC_TAXABLE_ALLOW)))
    rowsA.Add Array("Non-taxable allowances", modMoney.ToDec(slips(found, repoStaff.PSC_NONTAX_ALLOW)))
    rowsA.Add Array("Overtime", modMoney.ToDec(slips(found, repoStaff.PSC_OVERTIME)))
    rowsA.Add Array("Adjustment", modMoney.ToDec(slips(found, repoStaff.PSC_ADJUST)))
    rowsA.Add Array("GROSS", modMoney.ToDec(slips(found, repoStaff.PSC_GROSS)))
    rowsA.Add Array("Employee social security", -modMoney.ToDec(slips(found, repoStaff.PSC_EESSB)))
    rowsA.Add Array("Taxable income", modMoney.ToDec(slips(found, repoStaff.PSC_TAXABLE)))
    rowsA.Add Array("Income tax", -modMoney.ToDec(slips(found, repoStaff.PSC_TAX)))
    rowsA.Add Array("Advance recovered", -modMoney.ToDec(slips(found, repoStaff.PSC_ADVANCE)))
    rowsA.Add Array("NET PAY", modMoney.ToDec(slips(found, repoStaff.PSC_NET)))
    rowsA.Add Array("Employer social security (memo)", modMoney.ToDec(slips(found, repoStaff.PSC_ERSSB)))
    modSheetUtils.SetCell SH_PAYSLIP, 1, 1, "PAYSLIP (SIMULATION) " & modConstants.CCY_MMK, True
    modSheetUtils.WriteGrid SH_PAYSLIP, 3, 1, svcReports.ToGrid(Array("Item", "Amount"), rowsA)
    modSheetUtils.StyleReport SH_PAYSLIP, 3, 2, "2"
End Sub

Private Sub BuildSummary(ByVal d As Date, ByVal slips As Variant, ByVal emps As Variant)
    Dim cnt As Object, gross As Object, net As Object, tax As Object, i As Long, er As Long, key As String, k As Variant, parts() As String, rowsA As Collection
    Set cnt = CreateObject("Scripting.Dictionary")
    Set gross = CreateObject("Scripting.Dictionary")
    Set net = CreateObject("Scripting.Dictionary")
    Set tax = CreateObject("Scripting.Dictionary")
    If Not IsEmpty(slips) Then
        For i = 1 To UBound(slips, 1)
            If CStr(slips(i, repoStaff.PSC_STATUS)) <> "Reversed" Then
                er = repoStaff.FindRowById(emps, repoStaff.EMC_ID, CStr(slips(i, repoStaff.PSC_EMPLOYEE)))
                key = CStr(slips(i, repoStaff.PSC_PERIOD)) & "|" & CStr(slips(i, repoStaff.PSC_BRANCH)) & "|"
                If er > 0 Then key = key & CStr(emps(er, repoStaff.EMC_POSITION))
                AddTo cnt, key, 1
                AddTo gross, key, modMoney.ToDec(slips(i, repoStaff.PSC_GROSS))
                AddTo net, key, modMoney.ToDec(slips(i, repoStaff.PSC_NET))
                AddTo tax, key, modMoney.ToDec(slips(i, repoStaff.PSC_TAX))
            End If
        Next i
    End If
    Set rowsA = New Collection
    For Each k In cnt.Keys
        parts = Split(CStr(k), "|")
        rowsA.Add Array(parts(0), parts(1), parts(2), cnt(k), gross(k), tax(k), net(k))
    Next k
    modSheetUtils.PrepareReport SH_SUMMARY, "1,2,3"
    modSheetUtils.SetCell SH_SUMMARY, 1, 1, "Payroll summary by period, branch and position " & modDateUtils.ToIso(d) & " (SIMULATION)", True
    modSheetUtils.WriteGrid SH_SUMMARY, 3, 1, svcReports.ToGrid(Array("Period", "Branch", "Position", "Payslips", "Gross", "Income tax", "Net pay"), rowsA)
    modSheetUtils.StyleReport SH_SUMMARY, 3, 7, "5,6,7"
End Sub

' Staff cost (gross + employer SSB) per branch for a period against the Stage 2 running costs of the same period.
Private Sub BuildCost(ByVal d As Date, ByVal slips As Variant, ByVal inp As Object)
    Dim period As String, staff As Object, running As Object, i As Long, posted As Variant, branches As Object, b As Variant, rowsA As Collection, sc As Variant, rc As Variant, ratio As String
    period = CStr(inp(modConstants.INP_PERIOD))
    If Len(period) <> 6 Then period = modDateUtils.PeriodOf(d)
    Set staff = CreateObject("Scripting.Dictionary")
    Set running = CreateObject("Scripting.Dictionary")
    If Not IsEmpty(slips) Then
        For i = 1 To UBound(slips, 1)
            If CStr(slips(i, repoStaff.PSC_PERIOD)) = period And CStr(slips(i, repoStaff.PSC_STATUS)) <> "Reversed" Then AddTo staff, CStr(slips(i, repoStaff.PSC_BRANCH)), modMoney.ToDec(slips(i, repoStaff.PSC_GROSS)) + modMoney.ToDec(slips(i, repoStaff.PSC_ERSSB))
        Next i
    End If
    If modSheetUtils.TableExists(modConstants.TBL_COSTPOSTINGS) Then
        posted = modSheetUtils.ReadTable(modConstants.TBL_COSTPOSTINGS)
        If Not IsEmpty(posted) Then
            For i = 1 To UBound(posted, 1)
                If CStr(posted(i, repoCost.CPC_PERIOD)) = period Then AddTo running, CStr(posted(i, repoCost.CPC_BRANCH)), modMoney.ToDec(posted(i, repoCost.CPC_AMOUNT))
            Next i
        End If
    End If
    Set branches = repoBranch.LoadBranches()
    Set rowsA = New Collection
    For Each b In branches.Keys
        sc = Num(staff, CStr(b))
        rc = Num(running, CStr(b))
        ratio = ""
        If rc > 0 Then ratio = Format$(CDbl(sc) / CDbl(rc), "0.00")
        rowsA.Add Array(CStr(b), branches(b).Name, sc, rc, sc + rc, ratio)
    Next b
    modSheetUtils.PrepareReport SH_COST, "1"
    modSheetUtils.SetCell SH_COST, 1, 1, "Staff cost against running costs, period " & period & " (SIMULATION)", True
    modSheetUtils.WriteGrid SH_COST, 3, 1, svcReports.ToGrid(Array("Branch", "Name", "Staff cost (gross + employer SSB)", "Running costs (Stage 2)", "Total", "Staff / running"), rowsA)
    modSheetUtils.StyleReport SH_COST, 3, 6, "3,4,5"
End Sub

Private Sub BuildAdvances(ByVal d As Date, ByVal adv As Variant)
    Dim i As Long, rowsA As Collection
    Set rowsA = New Collection
    If Not IsEmpty(adv) Then
        For i = 1 To UBound(adv, 1)
            rowsA.Add Array(CStr(adv(i, 1)), CStr(adv(i, 2)), adv(i, 3), modMoney.ToDec(adv(i, 4)), modMoney.ToDec(adv(i, 5)), modMoney.ToDec(adv(i, 6)), CStr(adv(i, 7)))
        Next i
    End If
    modSheetUtils.PrepareReport SH_ADV, "1,2"
    modSheetUtils.SetCell SH_ADV, 1, 1, "Staff advances " & modDateUtils.ToIso(d) & " (SIMULATION)", True
    modSheetUtils.WriteGrid SH_ADV, 3, 1, svcReports.ToGrid(Array("Advance", "Employee", "Issued", "Amount", "Monthly recovery", "Outstanding", "Status"), rowsA)
    modSheetUtils.StyleReport SH_ADV, 3, 7, "4,5,6"
End Sub

Private Sub BuildScale(ByVal d As Date)
    Dim data As Variant, i As Long, rowsA As Collection
    Set rowsA = New Collection
    data = modSheetUtils.ReadTable(modConstants.TBL_SALARYSCALES)
    If Not IsEmpty(data) Then
        For i = 1 To UBound(data, 1)
            rowsA.Add Array(CStr(data(i, 2)), CLng(modMoney.ToDec(data(i, 3))), modMoney.ToDec(data(i, 4)), data(i, 5))
        Next i
    End If
    modSheetUtils.PrepareReport SH_SCALE, "1"
    modSheetUtils.SetCell SH_SCALE, 1, 1, "Salary scales (monthly, MMK) (SIMULATION)", True
    modSheetUtils.WriteGrid SH_SCALE, 3, 1, svcReports.ToGrid(Array("Position", "Grade", "Base salary", "Effective from"), rowsA)
    modSheetUtils.StyleReport SH_SCALE, 3, 4, "3"
End Sub

' Called from svcReportsExtra. Skipped when the Stage 5 tables are missing.
Public Sub RunStaffReports(ByVal d As Date, ByVal inp As Object)
    Dim emps As Variant, slips As Variant, adv As Variant, t As Variant
    If Not modSheetUtils.TableExists(modConstants.TBL_EMPLOYEES) Or Not modSheetUtils.TableExists(modConstants.TBL_PAYSLIPS) Then Exit Sub
    For Each t In Array(SH_PAYSLIP, SH_SUMMARY, SH_COST, SH_GAPS, SH_HEAD, SH_ADV, SH_SCALE)
        modSheetUtils.EnsureSheet CStr(t)
    Next t
    emps = modSheetUtils.ReadTable(modConstants.TBL_EMPLOYEES)
    slips = modSheetUtils.ReadTable(modConstants.TBL_PAYSLIPS)
    adv = modSheetUtils.ReadTable(modConstants.TBL_ADVANCES)
    BuildGaps d, emps
    BuildHeadcount d, emps
    BuildPayslip d, emps, slips, inp
    BuildSummary d, slips, emps
    BuildCost d, slips, inp
    BuildAdvances d, adv
    BuildScale d
End Sub
