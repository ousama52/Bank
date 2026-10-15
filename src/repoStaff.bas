Attribute VB_Name = "repoStaff"
Option Explicit

' Column positions and row builders for the Stage 5 tables (read and changed through clsUnitOfWork.WorkingTable).

' tblEmployees (13): EmployeeID, Name, PositionCode, Grade, BranchID, HireDate, ExitDate, Status, PayMode, BankAccountNo, Dependents, UserID, Phone
Public Const EMC_ID As Long = 1
Public Const EMC_NAME As Long = 2
Public Const EMC_POSITION As Long = 3
Public Const EMC_GRADE As Long = 4
Public Const EMC_BRANCH As Long = 5
Public Const EMC_HIRE As Long = 6
Public Const EMC_EXIT As Long = 7
Public Const EMC_STATUS As Long = 8
Public Const EMC_PAYMODE As Long = 9
Public Const EMC_ACCOUNT As Long = 10
Public Const EMC_DEPENDENTS As Long = 11
Public Const EMC_USER As Long = 12
Public Const EMC_PHONE As Long = 13

' tblEmployeeHistory (11): HistoryID, EmployeeID, ChangeDate, ChangeType, FromBranch, ToBranch, FromPosition, ToPosition, FromGrade, ToGrade, Note
Public Const EHC_EMPLOYEE As Long = 2
Public Const EHC_DATE As Long = 3
Public Const EHC_TYPE As Long = 4
Public Const EHC_FROMBRANCH As Long = 5
Public Const EHC_FROMPOS As Long = 7
Public Const EHC_FROMGRADE As Long = 9

' tblAttendance (8): AttendanceID, EmployeeID, Period, UnpaidLeaveDays, OvertimeHours, HolidayOvertimeHours, Adjustment, AdjustmentNote
' tblStaffAdvances (8): AdvanceID, EmployeeID, IssueDate, Amount, MonthlyRecovery, Outstanding, Status, JournalID
Public Const ADC_EMPLOYEE As Long = 2
Public Const ADC_MONTHLY As Long = 5
Public Const ADC_OUTSTANDING As Long = 6
Public Const ADC_STATUS As Long = 7

' tblPayrollRuns (16): PayrollRunID, Period, BranchID, PreparedDate, PreparedBy, ApprovedBy, EmployeesPaid, TotalGross, TotalEmployeeSSB, TotalTax,
' TotalAdvanceRecovered, TotalNet, TotalEmployerSSB, Status, PostJournalID (holds the GroupID of the posting), PayJournalID (GroupID of the payment)
Public Const PRC_ID As Long = 1
Public Const PRC_PERIOD As Long = 2
Public Const PRC_BRANCH As Long = 3
Public Const PRC_PREPAREDBY As Long = 5
Public Const PRC_APPROVEDBY As Long = 6
Public Const PRC_COUNT As Long = 7
Public Const PRC_GROSS As Long = 8
Public Const PRC_EESSB As Long = 9
Public Const PRC_TAX As Long = 10
Public Const PRC_ADVANCE As Long = 11
Public Const PRC_NET As Long = 12
Public Const PRC_ERSSB As Long = 13
Public Const PRC_STATUS As Long = 14
Public Const PRC_POSTGROUP As Long = 15
Public Const PRC_PAYGROUP As Long = 16

' tblPayslips (20): PayslipID, PayrollRunID, EmployeeID, BranchID, Period, BaseSalary, UnpaidDeduction, TaxableAllowances, NonTaxableAllowances, OvertimePay,
' Adjustment, Gross, EmployeeSSB, TaxableIncome, IncomeTax, AdvanceRecovered, NetPay, EmployerSSB, PayMode, Status
Public Const PSC_ID As Long = 1
Public Const PSC_RUN As Long = 2
Public Const PSC_EMPLOYEE As Long = 3
Public Const PSC_BRANCH As Long = 4
Public Const PSC_PERIOD As Long = 5
Public Const PSC_BASE As Long = 6
Public Const PSC_UNPAID As Long = 7
Public Const PSC_TAXABLE_ALLOW As Long = 8
Public Const PSC_NONTAX_ALLOW As Long = 9
Public Const PSC_OVERTIME As Long = 10
Public Const PSC_ADJUST As Long = 11
Public Const PSC_GROSS As Long = 12
Public Const PSC_EESSB As Long = 13
Public Const PSC_TAXABLE As Long = 14
Public Const PSC_TAX As Long = 15
Public Const PSC_ADVANCE As Long = 16
Public Const PSC_NET As Long = 17
Public Const PSC_ERSSB As Long = 18
Public Const PSC_PAYMODE As Long = 19
Public Const PSC_STATUS As Long = 20
Public Const PSC_COLS As Long = 20

Public Function FindRowById(ByVal data As Variant, ByVal idCol As Long, ByVal id As String) As Long
    Dim i As Long
    If IsEmpty(data) Then Exit Function
    For i = 1 To UBound(data, 1)
        If StrComp(CStr(data(i, idCol)), id, vbTextCompare) = 0 Then
            FindRowById = i
            Exit Function
        End If
    Next i
End Function

Public Function EmployeeRow(ByVal id As String, ByVal name As String, ByVal pos As String, ByVal grade As Long, ByVal branchId As String, ByVal hire As Date, _
                            ByVal payMode As String, ByVal account As String, ByVal dependents As Long, ByVal phone As String) As Variant
    Dim r(1 To 13) As Variant
    r(EMC_ID) = id
    r(EMC_NAME) = name
    r(EMC_POSITION) = pos
    r(EMC_GRADE) = grade
    r(EMC_BRANCH) = branchId
    r(EMC_HIRE) = hire
    r(EMC_STATUS) = "Active"
    r(EMC_PAYMODE) = payMode
    If Len(account) > 0 Then r(EMC_ACCOUNT) = account
    r(EMC_DEPENDENTS) = dependents
    If Len(phone) > 0 Then r(EMC_PHONE) = phone
    EmployeeRow = r
End Function

Public Function HistoryRow(ByVal id As String, ByVal employeeId As String, ByVal d As Date, ByVal changeType As String, ByVal fromBranch As String, ByVal toBranch As String, _
                           ByVal fromPos As String, ByVal toPos As String, ByVal fromGrade As Variant, ByVal toGrade As Variant, ByVal note As String) As Variant
    Dim r(1 To 11) As Variant
    r(1) = id
    r(2) = employeeId
    r(3) = d
    r(4) = changeType
    If Len(fromBranch) > 0 Then r(5) = fromBranch
    If Len(toBranch) > 0 Then r(6) = toBranch
    If Len(fromPos) > 0 Then r(7) = fromPos
    If Len(toPos) > 0 Then r(8) = toPos
    r(9) = fromGrade
    r(10) = toGrade
    If Len(note) > 0 Then r(11) = note
    HistoryRow = r
End Function

' Branch, position and grade the employee had on date d (walks the history backwards from today).
Public Sub StateAt(ByVal hist As Variant, ByVal emps As Variant, ByVal row As Long, ByVal d As Date, ByRef branch As String, ByRef pos As String, ByRef grade As Long)
    Dim i As Long, hd As Date, id As String, t As String
    id = CStr(emps(row, EMC_ID))
    branch = CStr(emps(row, EMC_BRANCH))
    pos = CStr(emps(row, EMC_POSITION))
    grade = CLng(modMoney.ToDec(emps(row, EMC_GRADE)))
    If IsEmpty(hist) Then Exit Sub
    For i = UBound(hist, 1) To 1 Step -1
        If CStr(hist(i, EHC_EMPLOYEE)) = id Then
            If modDateUtils.TryToDate(hist(i, EHC_DATE), hd) Then
                If hd > d Then
                    t = CStr(hist(i, EHC_TYPE))
                    If t = "Transfer" And Len(CStr(hist(i, EHC_FROMBRANCH))) > 0 Then branch = CStr(hist(i, EHC_FROMBRANCH))
                    If t = "Promotion" Then
                        If Len(CStr(hist(i, EHC_FROMPOS))) > 0 Then pos = CStr(hist(i, EHC_FROMPOS))
                        If Len(CStr(hist(i, EHC_FROMGRADE))) > 0 Then grade = CLng(modMoney.ToDec(hist(i, EHC_FROMGRADE)))
                    End If
                    If t = "GradeChange" And Len(CStr(hist(i, EHC_FROMGRADE))) > 0 Then grade = CLng(modMoney.ToDec(hist(i, EHC_FROMGRADE)))
                End If
            End If
        End If
    Next i
End Sub
