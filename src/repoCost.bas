Attribute VB_Name = "repoCost"
Option Explicit

' Column positions and row builders for the Stage 2 tables. The tables themselves are read through clsUnitOfWork.WorkingTable.

' tblBranchCosts: CostID, BranchID, CostType, Description, Frequency, Amount, GLCode, StartDate, EndDate, Status
Public Const BCC_ID As Long = 1
Public Const BCC_BRANCH As Long = 2
Public Const BCC_FREQ As Long = 5
Public Const BCC_AMOUNT As Long = 6
Public Const BCC_GL As Long = 7
Public Const BCC_START As Long = 8
Public Const BCC_END As Long = 9
Public Const BCC_STATUS As Long = 10

' tblCostPostings: PostingID, CostID, BranchID, Period, GLCode, Amount, JournalID, Status, PaidDate, PaidJournalID
Public Const CPC_ID As Long = 1
Public Const CPC_COST As Long = 2
Public Const CPC_BRANCH As Long = 3
Public Const CPC_PERIOD As Long = 4
Public Const CPC_GL As Long = 5
Public Const CPC_AMOUNT As Long = 6
Public Const CPC_STATUS As Long = 8
Public Const CPC_PAIDDATE As Long = 9
Public Const CPC_PAIDJ As Long = 10

' tblExpenses: ExpenseID, BranchID, ExpenseDate, GLCode, Amount, Vendor, Description, PaidFrom, Status, RequestedBy, ApprovedBy, JournalID, PaidJournalID
Public Const EXC_ID As Long = 1
Public Const EXC_BRANCH As Long = 2
Public Const EXC_DATE As Long = 3
Public Const EXC_GL As Long = 4
Public Const EXC_AMOUNT As Long = 5
Public Const EXC_STATUS As Long = 9
Public Const EXC_PAIDJ As Long = 13

Public Function PostingRow(ByVal id As String, ByVal costId As String, ByVal branchId As String, ByVal period As String, _
                           ByVal glCode As String, ByVal amount As Variant, ByVal journalId As String) As Variant
    Dim r(1 To 10) As Variant
    r(1) = id
    r(2) = costId
    r(3) = branchId
    r(4) = period
    r(5) = glCode
    r(6) = amount
    r(7) = journalId
    r(8) = "Unpaid"
    PostingRow = r
End Function

Public Function ExpenseRow(ByVal id As String, ByVal branchId As String, ByVal d As Date, ByVal glCode As String, ByVal amount As Variant, _
                           ByVal vendor As String, ByVal description As String, ByVal paidFrom As String, ByVal status As String, _
                           ByVal requestedBy As String, ByVal approvedBy As String, ByVal journalId As String, ByVal paidJournal As String) As Variant
    Dim r(1 To 13) As Variant
    r(1) = id
    r(2) = branchId
    r(3) = d
    r(4) = glCode
    r(5) = amount
    r(6) = vendor
    r(7) = description
    r(8) = paidFrom
    r(9) = status
    r(10) = requestedBy
    If Len(approvedBy) > 0 Then r(11) = approvedBy
    r(12) = journalId
    If Len(paidJournal) > 0 Then r(13) = paidJournal
    ExpenseRow = r
End Function

' Amount of one cost row for one month: Monthly = Amount, Quarterly = Amount / 3, Yearly = Amount / 12 (rounded).
Public Function MonthlyAmount(ByVal frequency As String, ByVal amount As Variant) As Variant
    Select Case frequency
        Case "Quarterly": MonthlyAmount = modMoney.RoundCcy(amount / 3, modConstants.CCY_MMK)
        Case "Yearly": MonthlyAmount = modMoney.RoundCcy(amount / 12, modConstants.CCY_MMK)
        Case Else: MonthlyAmount = modMoney.ToDec(amount)
    End Select
End Function
