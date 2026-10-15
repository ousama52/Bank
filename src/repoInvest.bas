Attribute VB_Name = "repoInvest"
Option Explicit

' Column positions for the Stage 7 tables (read and changed through clsUnitOfWork.WorkingTable).

' tblInvestees (6): InvesteeID, Name, Sector, RiskGrade, Status, Note
Public Const IVC_ID As Long = 1
Public Const IVC_NAME As Long = 2
Public Const IVC_SECTOR As Long = 3
Public Const IVC_GRADE As Long = 4
Public Const IVC_STATUS As Long = 5

' tblInvestments (18): InvestmentID, InvesteeID, InvestmentType, Amount, StartDate, TermMonths, MaturityDate, RateOrShare, PayFrequency, Status, CarryingValue,
' AccruedReturn, LastAccrualDate, RiskGrade, Sector, ProposedBy, ApprovedBy, Note
Public Const INC_ID As Long = 1
Public Const INC_INVESTEE As Long = 2
Public Const INC_TYPE As Long = 3
Public Const INC_AMOUNT As Long = 4
Public Const INC_START As Long = 5
Public Const INC_TERM As Long = 6
Public Const INC_MATURITY As Long = 7
Public Const INC_RATE As Long = 8
Public Const INC_FREQ As Long = 9
Public Const INC_STATUS As Long = 10
Public Const INC_CARRYING As Long = 11
Public Const INC_ACCRUED As Long = 12
Public Const INC_LASTACCRUAL As Long = 13
Public Const INC_GRADE As Long = 14
Public Const INC_SECTOR As Long = 15
Public Const INC_PROPOSEDBY As Long = 16
Public Const INC_APPROVEDBY As Long = 17
Public Const INC_NOTE As Long = 18
Public Const INC_COLS As Long = 18

' tblInvestmentReturns (7): ReturnID, InvestmentID, ReturnDate, Amount, ReturnType, JournalID, Note
' Note tells recognition from cash: "Recognized" (income booked to 1610), "Received;1610=n" (cash, n of it cleared 1610), "WriteOff;1610=n".
Public Const IRC_INVESTMENT As Long = 2
Public Const IRC_DATE As Long = 3
Public Const IRC_AMOUNT As Long = 4
Public Const IRC_NOTE As Long = 7

' tblInvestmentValuations (7): ValuationID, InvestmentID, ValuationDate, Value, Method, Note, JournalID

Public Function FindInvestById(ByVal data As Variant, ByVal id As String) As Long
    Dim i As Long
    If IsEmpty(data) Then Exit Function
    For i = 1 To UBound(data, 1)
        If StrComp(CStr(data(i, 1)), id, vbTextCompare) = 0 Then
            FindInvestById = i
            Exit Function
        End If
    Next i
End Function

' The number after "1610=" in a return note (0 when there is none).
Private Function Cleared1610(ByVal note As String) As Variant
    Dim p As Long, s As String
    p = InStr(note, "1610=")
    Cleared1610 = CDec(0)
    If p > 0 Then
        s = Mid$(note, p + 5)
        If s Like "*[!0-9]*" Then Exit Function
        If Len(s) > 0 Then Cleared1610 = CDec(s)
    End If
End Function

' Recognized but unpaid return of one investment: all "Recognized" rows minus what cash and write-offs cleared from 1610.
Public Function UnpaidRecognized(ByVal returns As Variant, ByVal investmentId As String) As Variant
    Dim i As Long, total As Variant, note As String
    total = CDec(0)
    If Not IsEmpty(returns) Then
        For i = 1 To UBound(returns, 1)
            If CStr(returns(i, IRC_INVESTMENT)) = investmentId Then
                note = CStr(returns(i, IRC_NOTE))
                If Left$(note, 10) = "Recognized" Then
                    total = total + modMoney.ToDec(returns(i, IRC_AMOUNT))
                Else
                    total = total - Cleared1610(note)
                End If
            End If
        Next i
    End If
    UnpaidRecognized = total
End Function

Public Function ReturnRow(ByVal id As String, ByVal investmentId As String, ByVal d As Date, ByVal amount As Variant, ByVal returnType As String, ByVal journalId As String, ByVal note As String) As Variant
    Dim r(1 To 7) As Variant
    r(1) = id
    r(2) = investmentId
    r(3) = d
    r(4) = amount
    r(5) = returnType
    r(6) = journalId
    r(7) = note
    ReturnRow = r
End Function
