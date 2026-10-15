Attribute VB_Name = "repoFee"
Option Explicit

' Columns of tblFees: FeeCode, Name, Method, FlatAmount, Rate, MinFee, MaxFee, Currency, IncomeGL (Stage 2 adds Taxable at the end)

Public Function FindFee(ByVal feeCode As String) As clsFee
    Dim data As Variant, i As Long, f As clsFee
    data = modSheetUtils.ReadTable(modConstants.TBL_FEES)
    If IsEmpty(data) Then Exit Function
    For i = 1 To UBound(data, 1)
        If StrComp(CStr(data(i, 1)), feeCode, vbTextCompare) = 0 Then
            Set f = New clsFee
            f.FeeCode = CStr(data(i, 1))
            f.Method = CStr(data(i, 3))
            f.FlatAmount = modMoney.ToDec(data(i, 4))
            f.Rate = modMoney.ToDec(data(i, 5))
            f.MinFee = modMoney.ToDec(data(i, 6))
            f.MaxFee = modMoney.ToDec(data(i, 7))
            f.Ccy = CStr(data(i, 8))
            f.IncomeGL = CStr(data(i, 9))
            If UBound(data, 2) >= 10 Then f.Taxable = (CStr(data(i, 10)) = modConstants.FLAG_YES)
            Set FindFee = f
            Exit Function
        End If
    Next i
End Function
