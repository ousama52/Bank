Attribute VB_Name = "repoGov"
Option Explicit

' Column positions and small helpers for the Stage 8 tables (read and changed through clsUnitOfWork.WorkingTable).

' tblBonds (13): BondID, BondType, Name, FaceValue, IssueDate, MaturityDate, Rate, CouponMonths, PurchasePrice, CarryingValue, AccruedCoupon, LastAccrualDate, Status
Public Const BDC_ID As Long = 1
Public Const BDC_TYPE As Long = 2
Public Const BDC_NAME As Long = 3
Public Const BDC_FACE As Long = 4
Public Const BDC_ISSUE As Long = 5
Public Const BDC_MATURITY As Long = 6
Public Const BDC_RATE As Long = 7
Public Const BDC_COUPONMONTHS As Long = 8
Public Const BDC_PRICE As Long = 9
Public Const BDC_CARRYING As Long = 10
Public Const BDC_ACCRUED As Long = 11
Public Const BDC_LASTACCRUAL As Long = 12
Public Const BDC_STATUS As Long = 13

' tblAmlRules (6): RuleID, RuleType, Threshold, WindowDays, MinCount, Active
' tblAmlAlerts (11): AlertID, AlertDate, RuleID, CustomerID, AccountNo, Amount, Description, Status, CaseID, ReviewedBy, ReviewedDate
Public Const ALC_ID As Long = 1
Public Const ALC_DATE As Long = 2
Public Const ALC_RULE As Long = 3
Public Const ALC_CUSTOMER As Long = 4
Public Const ALC_STATUS As Long = 8
Public Const ALC_CASE As Long = 9
Public Const ALC_REVIEWEDBY As Long = 10
Public Const ALC_REVIEWEDDATE As Long = 11

' tblAmlCases (9): CaseID, OpenDate, CustomerID, Summary, Status, AssignedTo, Decision, ClosedDate, Notes
Public Const CSC_ID As Long = 1
Public Const CSC_CUSTOMER As Long = 3
Public Const CSC_STATUS As Long = 5
Public Const CSC_ASSIGNED As Long = 6
Public Const CSC_DECISION As Long = 7
Public Const CSC_CLOSED As Long = 8

' tblRatioHistory (5): RatioDate, RatioCode, Actual, Limit, Status
' tblComplianceEvents (7): EventID, EventDate, RatioCode, Actual, Limit, Status, Note
' tblFines (8): FineID, FineDate, Reason, Amount, IssuedBy, ApprovedBy, Status, JournalID
Public Const FNC_ID As Long = 1
Public Const FNC_AMOUNT As Long = 4
Public Const FNC_STATUS As Long = 7
Public Const FNC_JOURNAL As Long = 8
' tblInspections (6): InspectionID, InspectionDate, RunBy, Result, FindingCount, Summary

Public Function FindGovById(ByVal data As Variant, ByVal id As String) As Long
    Dim i As Long
    If IsEmpty(data) Then Exit Function
    For i = 1 To UBound(data, 1)
        If StrComp(CStr(data(i, 1)), id, vbTextCompare) = 0 Then
            FindGovById = i
            Exit Function
        End If
    Next i
End Function

' True when GL code is an income or expense account in the chart.
Public Function IsProfitGL(ByVal chart As Object, ByVal gl As String) As Boolean
    If chart.Exists(gl) Then IsProfitGL = (chart(gl)(1) = "Income" Or chart(gl)(1) = "Expense")
End Function

' Last closed fiscal year from Config (0 when none).
Public Function LastClosedYear() As Long
    Dim t As String
    t = repoConfig.GetText(modConstants.CFG_LAST_CLOSED)
    If t Like "####" Then LastClosedYear = CLng(t)
End Function
