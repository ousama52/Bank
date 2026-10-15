Attribute VB_Name = "modArgs"
Option Explicit

' Action arguments are a small Dictionary (text keys). The same dictionary is stored as the Note of an approval
' ("key=value|key=value") so an approved action runs with exactly the arguments that were requested.

Public Function NewArgs() As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = 1
    Set NewArgs = d
End Function

Private Function Clean(ByVal s As String) As String
    Clean = Replace(Replace(s, "|", "/"), "=", "-")
End Function

Public Function Encode(ByVal args As Object) As String
    Dim k As Variant, s As String
    For Each k In args.Keys
        If Len(CStr(args(k))) > 0 Then
            If Len(s) > 0 Then s = s & "|"
            s = s & CStr(k) & "=" & Clean(CStr(args(k)))
        End If
    Next k
    Encode = s
End Function

Public Function Decode(ByVal note As String) As Object
    Dim d As Object, parts() As String, i As Long, p As Long
    Set d = NewArgs()
    If Len(note) > 0 Then
        parts = Split(note, "|")
        For i = 0 To UBound(parts)
            p = InStr(parts(i), "=")
            If p > 1 Then d(Left$(parts(i), p - 1)) = Mid$(parts(i), p + 1)
        Next i
    End If
    Set Decode = d
End Function

Public Function Text(ByVal args As Object, ByVal key As String, Optional ByVal defaultValue As String = "") As String
    If args.Exists(key) Then Text = CStr(args(key)) Else Text = defaultValue
End Function

Public Function Dec(ByVal args As Object, ByVal key As String) As Variant
    Dim v As Variant
    If args.Exists(key) Then
        If modMoney.TryToDec(args(key), v) Then
            Dec = v
            Exit Function
        End If
    End If
    Dec = CDec(0)
End Function

' Amount rule shared by every action: a whole number of MMK greater than zero.
Public Function ParseAmount(ByVal raw As Variant, ByRef amount As Variant, ByRef msg As String, Optional ByVal allowZero As Boolean = False) As Boolean
    Dim v As Variant
    If Not modMoney.TryToDec(raw, v) Then
        msg = modMessages.MSG_BAD_AMOUNT
        Exit Function
    End If
    If v < 0 Or (v = 0 And Not allowZero) Or Not modMoney.IsWhole(v) Then
        msg = modMessages.MSG_BAD_AMOUNT
        Exit Function
    End If
    amount = v
    ParseAmount = True
End Function

' Amount in a currency: greater than zero with at most the currency's decimals (MMK = whole).
Public Function ParseAmountCcy(ByVal raw As Variant, ByVal ccy As String, ByRef amount As Variant, ByRef msg As String) As Boolean
    Dim v As Variant
    If Not modMoney.TryToDec(raw, v) Then msg = "Enter a valid amount in Control C7.": Exit Function
    If v <= 0 Then msg = "The amount must be greater than zero.": Exit Function
    If modMoney.RoundTo(v, modMoney.Decimals(ccy)) <> v Then
        msg = "A " & ccy & " amount can have at most " & modMoney.Decimals(ccy) & " decimals."
        Exit Function
    End If
    amount = v
    ParseAmountCcy = True
End Function

' Copies the text inputs of the Control sheet into args under short keys.
Public Function FromInputs(ByVal inp As Object) As Object
    Dim a As Object
    Set a = NewArgs()
    a("account") = CStr(inp(modConstants.INP_ACCOUNT))
    a("counter") = CStr(inp(modConstants.INP_COUNTER))
    a("branch") = CStr(inp(modConstants.INP_BRANCH))
    a("bank") = CStr(inp(modConstants.INP_BANKCODE))
    a("payout") = CStr(inp(modConstants.INP_PAYOUT))
    a("ref") = CStr(inp(modConstants.INP_REFID))
    a("customer") = CStr(inp(modConstants.INP_CUSTOMER))
    a("product") = CStr(inp(modConstants.INP_PRODUCT))
    a("note") = CStr(inp(modConstants.INP_NOTE))
    If Not IsEmpty(inp(modConstants.INP_TERM)) Then a("term") = CStr(inp(modConstants.INP_TERM))
    Set FromInputs = a
End Function

' Positive amount with any decimals. The decimals allowed for the account currency are checked later (svcFx.AmountOk).
Public Function ParseAmountAny(ByVal raw As Variant, ByRef amount As Variant, ByRef msg As String) As Boolean
    Dim v As Variant
    If Not modMoney.TryToDec(raw, v) Then
        msg = modMessages.MSG_BAD_AMOUNT
        Exit Function
    End If
    If v <= 0 Then
        msg = modMessages.MSG_BAD_AMOUNT
        Exit Function
    End If
    amount = v
    ParseAmountAny = True
End Function
