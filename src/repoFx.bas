Attribute VB_Name = "repoFx"
Option Explicit

Public Const IPC_ID As Long = 1
Public Const IPC_DIR As Long = 2
Public Const IPC_ACCOUNT As Long = 3
Public Const IPC_CCY As Long = 4
Public Const IPC_AMOUNT As Long = 5
Public Const IPC_MMK As Long = 6
Public Const IPC_PARTY As Long = 7
Public Const IPC_BANK As Long = 8
Public Const IPC_FEE As Long = 12
Public Const IPC_EXPECTED As Long = 15
Public Const IPC_STATUS As Long = 16
Public Const IPC_STATUSDATE As Long = 17
Public Const IPC_REASON As Long = 18
Public Const IPC_SUBMITJ As Long = 20
Public Const IPC_SETTLEJ As Long = 21
Public Const MEC_ID As Long = 1
Public Const MEC_WALLET As Long = 4
Public Const MEC_FEECODE As Long = 7
Public Const MEC_STATUS As Long = 8
Public Const AGC_ID As Long = 1
Public Const AGC_WALLET As Long = 3
Public Const AGC_BRANCH As Long = 4
Public Const AGC_CASH As Long = 5
Public Const AGC_STATUS As Long = 6

' Column positions and row builders for the Stage 6 tables (read and changed through clsUnitOfWork.WorkingTable).

' tblFxDeals (17): DealID, DealDate, DealType, Currency, ForeignAmount, Rate, MidRate, MmkAmount, Spread, BranchID, AccountNo, CustomerID, Purpose, JournalID, Status, CreatedBy, ApprovedBy
Public Function DealRow(ByVal id As String, ByVal d As Date, ByVal dealType As String, ByVal ccy As String, ByVal foreign As Variant, ByVal rate As Variant, ByVal mid As Variant, _
                        ByVal mmk As Variant, ByVal spread As Variant, ByVal branchId As String, ByVal accountNo As String, ByVal customerId As String, ByVal purpose As String, _
                        ByVal journalId As String, ByVal createdBy As String, ByVal approvedBy As String) As Variant
    Dim r(1 To 17) As Variant
    r(1) = id
    r(2) = d
    r(3) = dealType
    r(4) = ccy
    r(5) = foreign
    r(6) = rate
    r(7) = mid
    r(8) = mmk
    r(9) = spread
    r(10) = branchId
    If Len(accountNo) > 0 Then r(11) = accountNo
    If Len(customerId) > 0 Then r(12) = customerId
    If Len(purpose) > 0 Then r(13) = purpose
    r(14) = journalId
    r(15) = modConstants.STATUS_POSTED
    r(16) = createdBy
    If Len(approvedBy) > 0 Then r(17) = approvedBy
    DealRow = r
End Function

' tblFxPosition (9): PositionDate, Currency, Holdings, Liabilities, NetOpen, MidRate, NetOpenBase, LimitBase, Status

' tblIntlPayments (21): IntlPaymentID, Direction, AccountNo, Currency, Amount, MmkEquivalent, PartyName, ForeignBankCode, PartyAccount, Country, Purpose, Fee, MessageRef,
' SubmittedDate, ExpectedDate, Status, StatusDate, Reason, Screening, SubmitJournalID, SettleJournalID

Public Function IntlRow(ByVal id As String, ByVal direction As String, ByVal accountNo As String, ByVal ccy As String, ByVal amount As Variant, ByVal mmk As Variant, _
                        ByVal party As String, ByVal bankCode As String, ByVal partyAccount As String, ByVal country As String, ByVal purpose As String, ByVal fee As Variant, _
                        ByVal msgRef As String, ByVal submitted As Date, ByVal expected As Date, ByVal status As String, ByVal reason As String, ByVal screening As String, _
                        ByVal submitJournal As String) As Variant
    Dim r(1 To 21) As Variant
    r(1) = id
    r(2) = direction
    r(3) = accountNo
    r(4) = ccy
    r(5) = amount
    r(6) = mmk
    r(7) = party
    r(8) = bankCode
    r(9) = partyAccount
    r(10) = country
    r(11) = purpose
    r(12) = fee
    r(13) = msgRef
    r(14) = submitted
    r(15) = expected
    r(16) = status
    r(17) = submitted
    If Len(reason) > 0 Then r(18) = reason
    r(19) = screening
    If Len(submitJournal) > 0 Then r(20) = submitJournal
    IntlRow = r
End Function

' tblMerchants (8): MerchantID, Name, Category, WalletAccountNo, QRCode, BranchID, FeeCode, Status

' tblAgents (6): AgentID, Name, WalletAccountNo, BranchID, CashOnHand, Status

Public Function FindFxById(ByVal data As Variant, ByVal id As String) As Long
    Dim i As Long
    If IsEmpty(data) Then Exit Function
    For i = 1 To UBound(data, 1)
        If StrComp(CStr(data(i, 1)), id, vbTextCompare) = 0 Then
            FindFxById = i
            Exit Function
        End If
    Next i
End Function
