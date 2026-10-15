Attribute VB_Name = "repoLoan"
Option Explicit

' Column positions and row builders for the Stage 3 tables. Tables are read and changed through clsUnitOfWork.WorkingTable.

' tblLoans (23): LoanID, CustomerID, BranchID, ProductCode, AccountNo, Principal, AnnualRate, Method, TermMonths, ApplicationDate, ApprovedDate,
' DisburseDate, FirstDueDate, MaturityDate, Status, PrincipalOutstanding, PenaltyAccrued, DaysPastDue, ClassCode, ProvisionAmount, AppliedBy, ApprovedBy, Note
Public Const LOC_ID As Long = 1
Public Const LOC_CUSTOMER As Long = 2
Public Const LOC_BRANCH As Long = 3
Public Const LOC_PRODUCT As Long = 4
Public Const LOC_ACCOUNT As Long = 5
Public Const LOC_PRINCIPAL As Long = 6
Public Const LOC_RATE As Long = 7
Public Const LOC_METHOD As Long = 8
Public Const LOC_TERM As Long = 9
Public Const LOC_APPDATE As Long = 10
Public Const LOC_APPRDATE As Long = 11
Public Const LOC_DISBDATE As Long = 12
Public Const LOC_FIRSTDUE As Long = 13
Public Const LOC_MATURITY As Long = 14
Public Const LOC_STATUS As Long = 15
Public Const LOC_OUTSTANDING As Long = 16
Public Const LOC_PENALTY As Long = 17
Public Const LOC_DPD As Long = 18
Public Const LOC_CLASS As Long = 19
Public Const LOC_PROVISION As Long = 20
Public Const LOC_APPLIEDBY As Long = 21
Public Const LOC_APPROVEDBY As Long = 22
Public Const LOC_NOTE As Long = 23
Public Const LOC_COLS As Long = 23

Public Const LS_APPLIED As String = "Applied"
Public Const LS_APPROVED As String = "Approved"
Public Const LS_REJECTED As String = "Rejected"
Public Const LS_ACTIVE As String = "Active"
Public Const LS_CLOSED As String = "Closed"
Public Const LS_WRITTENOFF As String = "WrittenOff"
Public Const LS_REPOSSESSED As String = "Repossessed"

' tblLoanSchedule (12): ScheduleID, LoanID, InstallmentNo, DueDate, PrincipalDue, InterestDue, TotalDue, PrincipalPaid, InterestPaid, PenaltyPaid, Status, PaidDate
Public Const SCC_ID As Long = 1
Public Const SCC_LOAN As Long = 2
Public Const SCC_NO As Long = 3
Public Const SCC_DUE As Long = 4
Public Const SCC_PRIN As Long = 5
Public Const SCC_INT As Long = 6
Public Const SCC_TOTAL As Long = 7
Public Const SCC_PRINPAID As Long = 8
Public Const SCC_INTPAID As Long = 9
Public Const SCC_PENPAID As Long = 10
Public Const SCC_STATUS As Long = 11
Public Const SCC_PAIDDATE As Long = 12

' tblLoanProducts (16): ProductCode, Name, LoanType, AnnualRate, Method, MinAmount, MaxAmount, MinTermMonths, MaxTermMonths, ProcessingFeePct,
' GraceDays, PenaltyRate, CollateralRequired, MaxLTV, EarlyPayoffFeePct, IsActive
Public Const LPR_CODE As Long = 1
Public Const LPR_NAME As Long = 2
Public Const LPR_RATE As Long = 4
Public Const LPR_METHOD As Long = 5
Public Const LPR_MIN As Long = 6
Public Const LPR_MAX As Long = 7
Public Const LPR_MINTERM As Long = 8
Public Const LPR_MAXTERM As Long = 9
Public Const LPR_FEEPCT As Long = 10
Public Const LPR_GRACE As Long = 11
Public Const LPR_PENALTY As Long = 12
Public Const LPR_COLLATERAL As Long = 13
Public Const LPR_MAXLTV As Long = 14
Public Const LPR_EARLYFEE As Long = 15
Public Const LPR_ACTIVE As Long = 16

' tblLoanClasses (5): ClassCode, Name, MinDaysPastDue, MaxDaysPastDue, ProvisionPct
' tblLoanPayments (12): LoanPaymentID, LoanID, PaymentDate, Amount, PenaltyPart, InterestPart, PrincipalPart, Source, PaymentType, JournalID, GroupID, CreatedBy

Public Function FindLoanRow(ByVal data As Variant, ByVal loanId As String) As Long
    Dim i As Long
    If IsEmpty(data) Then Exit Function
    For i = 1 To UBound(data, 1)
        If StrComp(CStr(data(i, LOC_ID)), loanId, vbTextCompare) = 0 Then
            FindLoanRow = i
            Exit Function
        End If
    Next i
End Function

' One loan product as a 1D array (copy of the table row), or Empty when unknown.
Public Function ProductRow(ByVal code As String) As Variant
    Dim data As Variant, i As Long, c As Long, r(1 To 16) As Variant
    If Not modSheetUtils.TableExists(modConstants.TBL_LOANPRODUCTS) Then Exit Function
    data = modSheetUtils.ReadTable(modConstants.TBL_LOANPRODUCTS)
    If IsEmpty(data) Then Exit Function
    For i = 1 To UBound(data, 1)
        If StrComp(CStr(data(i, LPR_CODE)), code, vbTextCompare) = 0 Then
            For c = 1 To 16
                r(c) = data(i, c)
            Next c
            ProductRow = r
            Exit Function
        End If
    Next i
End Function

Public Function LoanRow(ByVal id As String, ByVal customerId As String, ByVal branchId As String, ByVal productCode As String, ByVal accountNo As String, _
                        ByVal principal As Variant, ByVal rate As Variant, ByVal method As String, ByVal term As Long, ByVal appDate As Date, _
                        ByVal appliedBy As String, ByVal note As String) As Variant
    Dim r(1 To 23) As Variant
    r(LOC_ID) = id
    r(LOC_CUSTOMER) = customerId
    r(LOC_BRANCH) = branchId
    r(LOC_PRODUCT) = productCode
    r(LOC_ACCOUNT) = accountNo
    r(LOC_PRINCIPAL) = principal
    r(LOC_RATE) = rate
    r(LOC_METHOD) = method
    r(LOC_TERM) = term
    r(LOC_APPDATE) = appDate
    r(LOC_STATUS) = LS_APPLIED
    r(LOC_OUTSTANDING) = CDec(0)
    r(LOC_PENALTY) = CDec(0)
    r(LOC_DPD) = 0
    r(LOC_PROVISION) = CDec(0)
    r(LOC_APPLIEDBY) = appliedBy
    If Len(note) > 0 Then r(LOC_NOTE) = note
    LoanRow = r
End Function

Public Function ScheduleRow(ByVal id As String, ByVal loanId As String, ByVal no As Long, ByVal due As Date, ByVal prin As Variant, ByVal intr As Variant) As Variant
    Dim r(1 To 12) As Variant
    r(SCC_ID) = id
    r(SCC_LOAN) = loanId
    r(SCC_NO) = no
    r(SCC_DUE) = due
    r(SCC_PRIN) = prin
    r(SCC_INT) = intr
    r(SCC_TOTAL) = prin + intr
    r(SCC_PRINPAID) = CDec(0)
    r(SCC_INTPAID) = CDec(0)
    r(SCC_PENPAID) = CDec(0)
    r(SCC_STATUS) = "Future"
    ScheduleRow = r
End Function

Public Function PaymentRow(ByVal id As String, ByVal loanId As String, ByVal d As Date, ByVal amount As Variant, ByVal pen As Variant, ByVal intr As Variant, _
                           ByVal prin As Variant, ByVal source As String, ByVal payType As String, ByVal journalId As String, ByVal groupId As String, _
                           ByVal createdBy As String) As Variant
    Dim r(1 To 12) As Variant
    r(1) = id
    r(2) = loanId
    r(3) = d
    r(4) = amount
    r(5) = pen
    r(6) = intr
    r(7) = prin
    r(8) = source
    r(9) = payType
    r(10) = journalId
    r(11) = groupId
    r(12) = createdBy
    PaymentRow = r
End Function

' Class and provision percentage for a number of days past due (from tblLoanClasses). False when no class matches.
Public Function ClassFor(ByVal classes As Variant, ByVal dpd As Long, ByRef classCode As Long, ByRef pct As Variant) As Boolean
    Dim i As Long, lo As Long, hi As Long
    If IsEmpty(classes) Then Exit Function
    For i = 1 To UBound(classes, 1)
        lo = CLng(modMoney.ToDec(classes(i, 3)))
        If dpd >= lo Then
            If Len(CStr(classes(i, 4))) = 0 Then
                hi = 2147483647
            Else
                hi = CLng(modMoney.ToDec(classes(i, 4)))
            End If
            If dpd <= hi Then
                classCode = CLng(modMoney.ToDec(classes(i, 1)))
                pct = modMoney.ToDec(classes(i, 5))
                ClassFor = True
                Exit Function
            End If
        End If
    Next i
End Function
