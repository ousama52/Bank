Attribute VB_Name = "modMoney"
Option Explicit

' Money is always Decimal (a Variant holding CDec). Never Double or Single.

Private m_dec As Object

' Decimals of a currency from tblCurrencies (MMK = 0). Falls back to 0 for MMK and 2 for others when the table is missing.
Public Function Decimals(ByVal ccy As String) As Long
    Dim data As Variant, i As Long
    If UCase$(ccy) = modConstants.CCY_MMK Or Len(ccy) = 0 Then
        Decimals = 0
        Exit Function
    End If
    If m_dec Is Nothing Then
        Set m_dec = CreateObject("Scripting.Dictionary")
        If modSheetUtils.TableExists("tblCurrencies") Then
            data = modSheetUtils.ReadTable("tblCurrencies")
            If Not IsEmpty(data) Then
                For i = 1 To UBound(data, 1)
                    If IsNumeric(data(i, 3)) Then m_dec(UCase$(CStr(data(i, 1)))) = CLng(data(i, 3))
                Next i
            End If
        End If
    End If
    If m_dec.Exists(UCase$(ccy)) Then Decimals = m_dec(UCase$(ccy)) Else Decimals = 2
End Function

' The one rounding engine: round half up to a number of decimal places.
' Negative amounts mirror the positive result so a reversal rounds the same way.
Public Function RoundTo(ByVal amount As Variant, ByVal places As Long) As Variant
    Dim v As Variant, factor As Variant, negFlag As Boolean, i As Long
    v = CDec(amount)
    If v < 0 Then
        negFlag = True
        v = -v
    End If
    factor = CDec(1)
    For i = 1 To places
        factor = factor * 10
    Next i
    v = Int(v * factor + CDec(0.5)) / factor
    If negFlag Then v = -v
    RoundTo = v
End Function

' THE rounding helper for money (hard rule 7).
Public Function RoundCcy(ByVal amount As Variant, ByVal ccy As String) As Variant
    RoundCcy = RoundTo(amount, Decimals(ccy))
End Function

' Interest accruals are kept to 4 decimals until they are posted.
Public Function RoundAccrual(ByVal amount As Variant) As Variant
    RoundAccrual = RoundTo(amount, 4)
End Function

' Parses text such as "0.10" or "1,250,000" without using the computer's decimal separator.
Public Function TryParseDec(ByVal s As String, ByRef result As Variant) As Boolean
    Dim negFlag As Boolean, p As Long, intPart As String, fracPart As String
    Dim v As Variant, factor As Variant, i As Long
    s = Replace(Replace(Trim$(s), ",", ""), " ", "")
    If Len(s) = 0 Then Exit Function
    If Left$(s, 1) = "-" Then
        negFlag = True
        s = Mid$(s, 2)
    ElseIf Left$(s, 1) = "+" Then
        s = Mid$(s, 2)
    End If
    p = InStr(s, ".")
    If p > 0 Then
        intPart = Left$(s, p - 1)
        fracPart = Mid$(s, p + 1)
    Else
        intPart = s
    End If
    If Len(intPart) = 0 And Len(fracPart) = 0 Then Exit Function
    If intPart Like "*[!0-9]*" Or fracPart Like "*[!0-9]*" Then Exit Function
    If Len(intPart) > 20 Or Len(fracPart) > 8 Then Exit Function
    v = CDec(0)
    If Len(intPart) > 0 Then v = CDec(intPart)
    If Len(fracPart) > 0 Then
        factor = CDec(1)
        For i = 1 To Len(fracPart)
            factor = factor * 10
        Next i
        v = v + CDec(fracPart) / factor
    End If
    If negFlag Then v = -v
    result = v
    TryParseDec = True
End Function

' Converts a cell value (number or text) to Decimal. False when it is blank or not a number.
Public Function TryToDec(ByVal v As Variant, ByRef result As Variant) As Boolean
    If IsError(v) Then Exit Function
    If IsNull(v) Or IsEmpty(v) Then Exit Function
    Select Case VarType(v)
        Case vbString
            TryToDec = TryParseDec(CStr(v), result)
        Case vbByte, vbInteger, vbLong, vbSingle, vbDouble, vbCurrency, vbDecimal
            On Error Resume Next
            result = CDec(v)
            TryToDec = (Err.Number = 0)
            On Error GoTo 0
    End Select
End Function

' Same as TryToDec but blank becomes 0 and bad text raises error 13.
Public Function ToDec(ByVal v As Variant) As Variant
    Dim d As Variant
    If IsError(v) Then Err.Raise 13, "modMoney.ToDec", "Not a number"
    If IsNull(v) Or IsEmpty(v) Then
        ToDec = CDec(0)
    ElseIf VarType(v) = vbString And Len(Trim$(CStr(v))) = 0 Then
        ToDec = CDec(0)
    ElseIf TryToDec(v, d) Then
        ToDec = d
    Else
        Err.Raise 13, "modMoney.ToDec", "Not a number: " & CStr(v)
    End If
End Function

Public Function IsWhole(ByVal amount As Variant) As Boolean
    IsWhole = (CDec(amount) = Int(CDec(amount)))
End Function

Public Function FormatAmount(ByVal amount As Variant) As String
    FormatAmount = Format$(CDec(amount), "#,##0")
End Function

' Amount with the decimals of its currency (MMK shows none).
Public Function FormatCcy(ByVal amount As Variant, ByVal ccy As String) As String
    Dim n As Long
    n = Decimals(ccy)
    If n > 0 Then
        FormatCcy = Format$(CDec(amount), "#,##0." & String$(n, "0")) & " " & ccy
    Else
        FormatCcy = Format$(CDec(amount), "#,##0") & " " & ccy
    End If
End Function
