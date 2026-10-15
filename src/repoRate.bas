Attribute VB_Name = "repoRate"
Option Explicit

' Columns of tblRateTiers: RateID, ProductCode, TermMonths, FromAmount, ToAmount, AnnualRate, EffectiveFrom

Public Function Load() As Variant
    Load = modSheetUtils.ReadTable(modConstants.TBL_RATETIERS)
End Function

Private Function InTier(ByVal data As Variant, ByVal i As Long, ByVal amount As Variant) As Boolean
    If amount < modMoney.ToDec(data(i, 4)) Then Exit Function
    If Len(CStr(data(i, 5))) = 0 Then
        InTier = True
    Else
        InTier = (amount < modMoney.ToDec(data(i, 5)))
    End If
End Function

' Rate for the product and term whose tier holds the WHOLE amount, using the latest EffectiveFrom on or before asOf.
' `data` comes from Load, so loops (end of day) read the table once.
Public Function RateFrom(ByVal data As Variant, ByVal productCode As String, ByVal termMonths As Long, _
                         ByVal amount As Variant, ByVal asOf As Date, ByRef rate As Variant) As Boolean
    Dim i As Long, eff As Date, bestEff As Date, found As Boolean
    If IsEmpty(data) Then Exit Function
    For i = 1 To UBound(data, 1)
        If CStr(data(i, 2)) = productCode And CLng(modMoney.ToDec(data(i, 3))) = termMonths Then
            If modDateUtils.TryToDate(data(i, 7), eff) Then
                If eff <= asOf And InTier(data, i, amount) Then
                    If (Not found) Or eff >= bestEff Then
                        bestEff = eff
                        rate = modMoney.ToDec(data(i, 6))
                        found = True
                    End If
                End If
            End If
        End If
    Next i
    RateFrom = found
End Function

Public Function Lookup(ByVal productCode As String, ByVal termMonths As Long, ByVal amount As Variant, _
                       ByVal asOf As Date, ByRef rate As Variant) As Boolean
    Lookup = RateFrom(Load(), productCode, termMonths, amount, asOf, rate)
End Function

' True when the product has any rate row for this term (a fixed deposit term must exist).
Public Function TermExists(ByVal productCode As String, ByVal termMonths As Long) As Boolean
    Dim data As Variant, i As Long
    data = Load()
    If IsEmpty(data) Then Exit Function
    For i = 1 To UBound(data, 1)
        If CStr(data(i, 2)) = productCode And CLng(modMoney.ToDec(data(i, 3))) = termMonths Then
            TermExists = True
            Exit Function
        End If
    Next i
End Function
