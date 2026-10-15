Attribute VB_Name = "repoInterbank"
Option Explicit

' tblInterbankTransfers columns (15): TransferID, Direction, AccountNo, OtherBankCode, OtherBankAccount, Amount, Fee,
' SubmittedDate, ExpectedSettleDate, Status, StatusDate, Reason, SubmitJournalID, SettleJournalID, ReturnJournalID
Public Const IBC_ID As Long = 1
Public Const IBC_DIR As Long = 2
Public Const IBC_ACCOUNT As Long = 3
Public Const IBC_BANK As Long = 4
Public Const IBC_BANKACCT As Long = 5
Public Const IBC_AMOUNT As Long = 6
Public Const IBC_FEE As Long = 7
Public Const IBC_SUBMITTED As Long = 8
Public Const IBC_EXPECTED As Long = 9
Public Const IBC_STATUS As Long = 10
Public Const IBC_STATUSDATE As Long = 11
Public Const IBC_REASON As Long = 12
Public Const IBC_SUBMITJ As Long = 13
Public Const IBC_SETTLEJ As Long = 14
Public Const IBC_RETURNJ As Long = 15

Public Function LoadTransfers() As Variant
    LoadTransfers = modSheetUtils.ReadTable(modConstants.TBL_INTERBANK)
End Function

Public Function NewTransferRow(ByVal transferId As String, ByVal direction As String, ByVal accountNo As String, ByVal bankCode As String, _
                       ByVal bankAccount As String, ByVal amount As Variant, ByVal fee As Variant, ByVal submitted As Date, _
                       ByVal expected As Date, ByVal status As String, ByVal submitJournal As String, ByVal settleJournal As String) As Variant
    Dim r(1 To 15) As Variant
    r(1) = transferId
    r(2) = direction
    r(3) = accountNo
    r(4) = bankCode
    r(5) = bankAccount
    r(6) = amount
    r(7) = fee
    r(8) = submitted
    r(9) = expected
    r(10) = status
    r(11) = submitted
    r(13) = submitJournal
    If Len(settleJournal) > 0 Then r(14) = settleJournal
    NewTransferRow = r
End Function

' Row number of a transfer in a ReadTable result, 0 when not found.
Public Function FindTransfer(ByVal data As Variant, ByVal transferId As String) As Long
    Dim i As Long
    If IsEmpty(data) Then Exit Function
    For i = 1 To UBound(data, 1)
        If CStr(data(i, IBC_ID)) = transferId Then
            FindTransfer = i
            Exit Function
        End If
    Next i
End Function
