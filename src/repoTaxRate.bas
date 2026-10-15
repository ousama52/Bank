Attribute VB_Name = "repoTaxRate"
Option Explicit

' tblTaxRates columns: TaxCode, Name, Rate, AppliesTo, GLCode, EffectiveFrom (the latest row on or before the date wins)

Public Function RateOn(ByVal taxCode As String, ByVal asOf As Date, ByRef rate As Variant) As Boolean
    Dim data As Variant, i As Long, eff As Date, best As Date, found As Boolean
    If Not modSheetUtils.TableExists(modConstants.TBL_TAXRATES) Then Exit Function
    data = modSheetUtils.ReadTable(modConstants.TBL_TAXRATES)
    If IsEmpty(data) Then Exit Function
    For i = 1 To UBound(data, 1)
        If StrComp(CStr(data(i, 1)), taxCode, vbTextCompare) = 0 Then
            If modDateUtils.TryToDate(data(i, 6), eff) Then
                If eff <= asOf And ((Not found) Or eff >= best) Then
                    best = eff
                    rate = modMoney.ToDec(data(i, 3))
                    found = True
                End If
            End If
        End If
    Next i
    RateOn = found
End Function
