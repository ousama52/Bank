Attribute VB_Name = "repoFxRate"
Option Explicit

' tblExchangeRates columns: RateDate, Currency, BuyRate, SellRate, MidRate (MMK per 1 unit of the currency).

Public Const FXR_DATE As Long = 1
Public Const FXR_CCY As Long = 2
Public Const FXR_BUY As Long = 3
Public Const FXR_SELL As Long = 4
Public Const FXR_MID As Long = 5

' The rate row of one currency on one exact date. False when there is none.
Public Function OnDate(ByVal data As Variant, ByVal ccy As String, ByVal d As Date, ByRef mid As Variant, ByRef buy As Variant, ByRef sell As Variant) As Boolean
    Dim i As Long, rd As Date
    If IsEmpty(data) Then Exit Function
    For i = UBound(data, 1) To 1 Step -1
        If StrComp(CStr(data(i, FXR_CCY)), ccy, vbTextCompare) = 0 Then
            If modDateUtils.TryToDate(data(i, FXR_DATE), rd) Then
                If rd = d Then
                    mid = modMoney.ToDec(data(i, FXR_MID))
                    buy = modMoney.ToDec(data(i, FXR_BUY))
                    sell = modMoney.ToDec(data(i, FXR_SELL))
                    OnDate = True
                    Exit Function
                End If
            End If
        End If
    Next i
End Function

' Mid rate for a currency on a date from the sheet (convenience for reports and checks). 0 when missing.
Public Function MidOn(ByVal ccy As String, ByVal d As Date) As Variant
    Dim mid As Variant, buy As Variant, sell As Variant
    MidOn = CDec(0)
    If ccy = modConstants.CCY_MMK Or Len(ccy) = 0 Then MidOn = CDec(1): Exit Function
    If Not modSheetUtils.TableExists(modConstants.TBL_EXCHANGERATES) Then Exit Function
    If OnDate(modSheetUtils.ReadTable(modConstants.TBL_EXCHANGERATES), ccy, d, mid, buy, sell) Then MidOn = mid
End Function

' Base (MMK) value of an amount in any currency at the mid rate of date d.
Public Function ToBase(ByVal amount As Variant, ByVal ccy As String, ByVal d As Date) As Variant
    If ccy = modConstants.CCY_MMK Or Len(ccy) = 0 Then
        ToBase = amount
    Else
        ToBase = modMoney.RoundCcy(amount * MidOn(ccy, d), modConstants.CCY_MMK)
    End If
End Function

' Latest mid rate of a currency on or before d (used by reports when today has no rate yet). 0 when none.
Public Function LatestMid(ByVal ccy As String, ByVal d As Date) As Variant
    Dim data As Variant, i As Long, rd As Date, best As Date, found As Boolean
    LatestMid = CDec(0)
    If ccy = modConstants.CCY_MMK Or Len(ccy) = 0 Then LatestMid = CDec(1): Exit Function
    If Not modSheetUtils.TableExists(modConstants.TBL_EXCHANGERATES) Then Exit Function
    data = modSheetUtils.ReadTable(modConstants.TBL_EXCHANGERATES)
    If IsEmpty(data) Then Exit Function
    For i = 1 To UBound(data, 1)
        If StrComp(CStr(data(i, FXR_CCY)), ccy, vbTextCompare) = 0 Then
            If modDateUtils.TryToDate(data(i, FXR_DATE), rd) Then
                If rd <= d And ((Not found) Or rd >= best) Then
                    best = rd
                    LatestMid = modMoney.ToDec(data(i, FXR_MID))
                    found = True
                End If
            End If
        End If
    Next i
End Function
