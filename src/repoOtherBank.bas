Attribute VB_Name = "repoOtherBank"
Option Explicit

' Columns of tblOtherBanks: BankCode, Name, SettleDays, FailRate, ReturnRate

Public Function FindBank(ByVal bankCode As String) As clsOtherBank
    Dim data As Variant, i As Long, b As clsOtherBank
    data = modSheetUtils.ReadTable(modConstants.TBL_OTHERBANKS)
    If IsEmpty(data) Then Exit Function
    For i = 1 To UBound(data, 1)
        If StrComp(CStr(data(i, 1)), bankCode, vbTextCompare) = 0 Then
            Set b = New clsOtherBank
            b.BankCode = CStr(data(i, 1))
            b.Name = CStr(data(i, 2))
            b.SettleDays = CLng(modMoney.ToDec(data(i, 3)))
            b.FailRate = modMoney.ToDec(data(i, 4))
            Set FindBank = b
            Exit Function
        End If
    Next i
End Function
