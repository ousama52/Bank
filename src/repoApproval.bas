Attribute VB_Name = "repoApproval"
Option Explicit

' tblApprovals columns (14): ApprovalID, RequestType, AccountNo, CounterAccountNo, Amount, BranchID, Note, RequestedBy,
' RequestedAt, Status, ApprovedBy, DecidedAt, DecisionNote, TxnGroupID
Public Const APC_ID As Long = 1
Public Const APC_STATUS As Long = 10
Public Const APC_APPROVER As Long = 11
Public Const APC_DECIDEDAT As Long = 12
Public Const APC_DECISIONNOTE As Long = 13
Public Const APC_GROUP As Long = 14

Public Function LoadApprovals() As Variant
    LoadApprovals = modSheetUtils.ReadTable(modConstants.TBL_APPROVALS)
End Function

Public Function ApprovalToRow(ByVal a As clsApproval) As Variant
    Dim r(1 To 14) As Variant
    r(1) = a.ApprovalID
    r(2) = a.RequestType
    If Len(a.AccountNo) > 0 Then r(3) = a.AccountNo
    If Len(a.CounterAccountNo) > 0 Then r(4) = a.CounterAccountNo
    r(5) = a.Amount
    r(6) = a.BranchID
    If Len(a.Note) > 0 Then r(7) = a.Note
    r(8) = a.RequestedBy
    r(9) = a.RequestedAt
    r(10) = a.Status
    r(11) = Empty
    r(12) = Empty
    r(13) = Empty
    r(14) = Empty
    ApprovalToRow = r
End Function

Public Function FromRow(ByVal data As Variant, ByVal i As Long) As clsApproval
    Dim a As clsApproval, d As Date
    Set a = New clsApproval
    a.ApprovalID = CStr(data(i, 1))
    a.RequestType = CStr(data(i, 2))
    a.AccountNo = CStr(data(i, 3))
    a.CounterAccountNo = CStr(data(i, 4))
    a.Amount = modMoney.ToDec(data(i, 5))
    a.BranchID = CStr(data(i, 6))
    a.Note = CStr(data(i, 7))
    a.RequestedBy = CStr(data(i, 8))
    If modDateUtils.TryToDate(data(i, 9), d) Then a.RequestedAt = d
    a.Status = CStr(data(i, 10))
    a.ApprovedBy = CStr(data(i, 11))
    a.DecisionNote = CStr(data(i, 13))
    a.TxnGroupID = CStr(data(i, 14))
    Set FromRow = a
End Function

Public Function FindApproval(ByVal data As Variant, ByVal approvalId As String) As Long
    Dim i As Long
    If IsEmpty(data) Then Exit Function
    For i = 1 To UBound(data, 1)
        If StrComp(CStr(data(i, 1)), approvalId, vbTextCompare) = 0 Then
            FindApproval = i
            Exit Function
        End If
    Next i
End Function

Public Function CountPending() As Long
    Dim data As Variant, i As Long
    data = LoadApprovals()
    If IsEmpty(data) Then Exit Function
    For i = 1 To UBound(data, 1)
        If CStr(data(i, APC_STATUS)) = modConstants.STATUS_PENDING Then CountPending = CountPending + 1
    Next i
End Function
