Attribute VB_Name = "modDateUtils"
Option Explicit

' Locale-proof date helpers. Dates are built from Year/Month/Day, never from Format strings.

Private Function Pad2(ByVal n As Long) As String
    Pad2 = Right$("0" & CStr(n), 2)
End Function

' 2026-10-13
Public Function ToIso(ByVal d As Date) As String
    ToIso = Right$("0000" & CStr(Year(d)), 4) & "-" & Pad2(Month(d)) & "-" & Pad2(Day(d))
End Function

' 20261013 (used inside TxnID and GroupID)
Public Function DateKey(ByVal d As Date) As String
    DateKey = Right$("0000" & CStr(Year(d)), 4) & Pad2(Month(d)) & Pad2(Day(d))
End Function

' Accepts a Date, a date serial number, "yyyy-mm-dd" text or any text VBA can read.
Public Function TryToDate(ByVal v As Variant, ByRef result As Date) As Boolean
    Dim s As String, y As Long, m As Long, dd As Long, t As Date
    If IsError(v) Then Exit Function
    If IsNull(v) Or IsEmpty(v) Then Exit Function
    Select Case VarType(v)
        Case vbDate, vbDouble, vbSingle, vbLong, vbInteger, vbCurrency, vbDecimal
            If CDbl(v) < 1 Or CDbl(v) > 2958465 Then Exit Function
            result = CDate(Int(CDbl(v)))
            TryToDate = True
        Case vbString
            s = Trim$(CStr(v))
            If s Like "####-##-##" Then
                y = CLng(Left$(s, 4))
                m = CLng(Mid$(s, 6, 2))
                dd = CLng(Right$(s, 2))
                If m < 1 Or m > 12 Or dd < 1 Or dd > 31 Then Exit Function
                t = DateSerial(y, m, dd)
                If Month(t) <> m Then Exit Function
                result = t
                TryToDate = True
            ElseIf IsDate(s) Then
                result = CDate(Int(CDbl(CDate(s))))
                TryToDate = True
            End If
    End Select
End Function

Public Function IsLastDayOfMonth(ByVal d As Date) As Boolean
    IsLastDayOfMonth = (Month(d + 1) <> Month(d))
End Function

Public Function EndOfMonth(ByVal d As Date) As Date
    EndOfMonth = DateSerial(Year(d), Month(d) + 1, 0)
End Function

Public Function AddMonths(ByVal d As Date, ByVal n As Long) As Date
    AddMonths = DateAdd("m", n, d)
End Function

' Whole days from a to b (b minus a).
Public Function DaysBetween(ByVal a As Date, ByVal b As Date) As Long
    DaysBetween = CLng(b) - CLng(a)
End Function

' The ONLY place the computer clock is read. Use it for CreatedAt and Log time stamps, never for business logic.
Public Function TimeStamp() As Date
    TimeStamp = Now
End Function

' Cell value to Date, or Empty when blank or not a date. Used when loading optional date columns.
Public Function ToDateOrEmpty(ByVal v As Variant) As Variant
    Dim d As Date
    If TryToDate(v, d) Then ToDateOrEmpty = d
End Function

' ---- Periods (YYYYMM text) ----

Public Function PeriodOf(ByVal d As Date) As String
    PeriodOf = Right$("0000" & CStr(Year(d)), 4) & Pad2(Month(d))
End Function

' "202610" -> first day of that month. False when the text is not a valid period.
Public Function TryParsePeriod(ByVal s As String, ByRef firstDay As Date) As Boolean
    Dim y As Long, m As Long
    s = Trim$(s)
    If Not (s Like "######") Then Exit Function
    y = CLng(Left$(s, 4))
    m = CLng(Right$(s, 2))
    If m < 1 Or m > 12 Or y < 1900 Then Exit Function
    firstDay = DateSerial(y, m, 1)
    TryParsePeriod = True
End Function
