Attribute VB_Name = "repoTransaction"
Option Explicit

' Columns of tblTransactions (16): TxnID, GroupID, TxnDate, AccountNo, TxnType, Direction, Amount, BalanceAfter,
' CounterAccountNo, BranchID, Description, Reference, JournalID, Status, CreatedBy, ApprovedBy
Public Const TXC_ACCOUNT As Long = 4
Public Const TXC_JOURNAL As Long = 13
Public Const TXC_STATUS As Long = 14
Public Const TXC_COUNT As Long = 16

Public Function LoadTransactions() As Variant
    LoadTransactions = modSheetUtils.ReadTable(modConstants.TBL_TRANSACTIONS)
End Function

' A statement row as a 1D array, ready for clsUnitOfWork.AppendRow.
Public Function TxnToRow(ByVal t As clsTxnRow) As Variant
    Dim r(1 To 16) As Variant
    r(1) = t.TxnID
    r(2) = t.GroupID
    r(3) = t.TxnDate
    r(4) = t.AccountNo
    r(5) = t.TxnType
    r(6) = t.Direction
    r(7) = t.Amount
    r(8) = t.BalanceAfter
    If Len(t.CounterAccountNo) > 0 Then r(9) = t.CounterAccountNo
    r(10) = t.BranchID
    r(11) = t.Description
    If Len(t.Reference) > 0 Then r(12) = t.Reference
    r(13) = t.JournalID
    r(14) = t.Status
    r(15) = t.CreatedBy
    If Len(t.ApprovedBy) > 0 Then r(16) = t.ApprovedBy
    TxnToRow = r
End Function

' Withdrawals, outgoing transfers and interbank-out of one account on one business date (not reversed).
Public Function DebitTotalForDay(ByVal accountNo As String, ByVal d As Date) As Variant
    Dim data As Variant, i As Long, total As Variant, t As String, dt As Date
    total = CDec(0)
    data = LoadTransactions()
    If Not IsEmpty(data) Then
        For i = 1 To UBound(data, 1)
            If CStr(data(i, 4)) = accountNo And CStr(data(i, 14)) = modConstants.STATUS_POSTED Then
                t = CStr(data(i, 5))
                If t = modConstants.TXN_WITHDRAWAL Or t = modConstants.TXN_TRANSFER_OUT Or t = modConstants.TXN_IB_OUT Then
                    If modDateUtils.TryToDate(data(i, 3), dt) Then
                        If dt = d Then total = total + modMoney.ToDec(data(i, 7))
                    End If
                End If
            End If
        Next i
    End If
    DebitTotalForDay = total
End Function
