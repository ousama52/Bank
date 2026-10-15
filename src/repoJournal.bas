Attribute VB_Name = "repoJournal"
Option Explicit

' tblJournal columns (9): JournalID, PostDate, Description, JournalType, GroupID, Status, ReversalOf, CreatedBy, CreatedAt
' tblJournalLines columns (10): LineID, JournalID, BranchID, GLCode, AccountNo, Currency, Debit, Credit, FxRate, ForeignAmount
Public Const JNC_ID As Long = 1
Public Const JNC_TYPE As Long = 4
Public Const JNC_GROUP As Long = 5
Public Const JNC_STATUS As Long = 6
Public Const LNC_JOURNAL As Long = 2
Public Const LNC_BRANCH As Long = 3
Public Const LNC_GL As Long = 4
Public Const LNC_ACCOUNT As Long = 5
Public Const LNC_DEBIT As Long = 7
Public Const LNC_CREDIT As Long = 8

Public Function LoadHeaders() As Variant
    LoadHeaders = modSheetUtils.ReadTable(modConstants.TBL_JOURNAL)
End Function

Public Function LoadLines() As Variant
    LoadLines = modSheetUtils.ReadTable(modConstants.TBL_JLINES)
End Function

Public Function HeaderRow(ByVal journalId As String, ByVal postDate As Date, ByVal description As String, ByVal journalType As String, _
                          ByVal groupId As String, ByVal reversalOf As String, ByVal createdBy As String) As Variant
    Dim r(1 To 9) As Variant
    r(1) = journalId
    r(2) = postDate
    r(3) = description
    r(4) = journalType
    r(5) = groupId
    r(6) = modConstants.STATUS_POSTED
    If Len(reversalOf) > 0 Then r(7) = reversalOf
    r(8) = createdBy
    r(9) = modDateUtils.TimeStamp()
    HeaderRow = r
End Function

Public Function LineRow(ByVal lineId As String, ByVal journalId As String, ByVal ln As clsJournalLine) As Variant
    Dim r(1 To 10) As Variant
    r(1) = lineId
    r(2) = journalId
    r(3) = ln.BranchID
    r(4) = ln.GLCode
    If Len(ln.AccountNo) > 0 Then r(5) = ln.AccountNo
    r(6) = ln.Ccy
    r(7) = ln.Debit
    r(8) = ln.Credit
    If ln.Ccy <> modConstants.CCY_MMK Then
        r(9) = ln.FxRate
        r(10) = ln.ForeignAmount
    End If
    LineRow = r
End Function

' Net (debit minus credit) per "BranchID|GLCode" from all posted lines.
Public Function LoadGLNet() As Object
    Dim lines As Variant, d As Object, i As Long, key As String
    Set d = CreateObject("Scripting.Dictionary")
    lines = LoadLines()
    If Not IsEmpty(lines) Then
        For i = 1 To UBound(lines, 1)
            key = CStr(lines(i, LNC_BRANCH)) & "|" & CStr(lines(i, LNC_GL))
            If d.Exists(key) Then
                d(key) = d(key) + modMoney.ToDec(lines(i, LNC_DEBIT)) - modMoney.ToDec(lines(i, LNC_CREDIT))
            Else
                d(key) = modMoney.ToDec(lines(i, LNC_DEBIT)) - modMoney.ToDec(lines(i, LNC_CREDIT))
            End If
        Next i
    End If
    Set LoadGLNet = d
End Function

' Total credit of one GL inside one journal (used to find the tax that was charged with a fee).
Public Function CreditOf(ByVal journalId As String, ByVal glCode As String) As Variant
    Dim lines As Variant, i As Long, total As Variant
    total = CDec(0)
    lines = LoadLines()
    If Not IsEmpty(lines) Then
        For i = 1 To UBound(lines, 1)
            If CStr(lines(i, LNC_JOURNAL)) = journalId And CStr(lines(i, LNC_GL)) = glCode Then total = total + modMoney.ToDec(lines(i, LNC_CREDIT))
        Next i
    End If
    CreditOf = total
End Function

' Net (debit minus credit) per "BranchID|GLCode" from all posted lines except those of year-end closing journals.
' Profit and loss reports use this so a closed year still shows its income and expenses.
Public Function LoadGLNetNoClose() As Object
    Dim lines As Variant, hdr As Variant, closing As Object, d As Object, i As Long, key As String, v As Variant
    Set d = CreateObject("Scripting.Dictionary")
    lines = LoadLines()
    If IsEmpty(lines) Then Set LoadGLNetNoClose = d: Exit Function
    hdr = LoadHeaders()
    Set closing = CreateObject("Scripting.Dictionary")
    If Not IsEmpty(hdr) Then
        For i = 1 To UBound(hdr, 1)
            If CStr(hdr(i, JNC_TYPE)) = modConstants.JT_YEAR_END Then closing(CStr(hdr(i, JNC_ID))) = True
        Next i
    End If
    For i = 1 To UBound(lines, 1)
        If Not closing.Exists(CStr(lines(i, LNC_JOURNAL))) Then
            key = CStr(lines(i, LNC_BRANCH)) & "|" & CStr(lines(i, LNC_GL))
            v = modMoney.ToDec(lines(i, LNC_DEBIT)) - modMoney.ToDec(lines(i, LNC_CREDIT))
            If d.Exists(key) Then d(key) = d(key) + v Else d(key) = v
        End If
    Next i
    Set LoadGLNetNoClose = d
End Function

' Lines joined with their journal: columns PostDate, BranchID, GLCode, Debit, Credit, JournalType (for period reports).
Public Function LoadDatedLines() As Variant
    Dim hdr As Variant, lines As Variant, dates As Object, types As Object, i As Long, out() As Variant, d As Date, j As String
    hdr = LoadHeaders()
    lines = LoadLines()
    If IsEmpty(hdr) Or IsEmpty(lines) Then Exit Function
    Set dates = CreateObject("Scripting.Dictionary")
    Set types = CreateObject("Scripting.Dictionary")
    For i = 1 To UBound(hdr, 1)
        If modDateUtils.TryToDate(hdr(i, 2), d) Then dates(CStr(hdr(i, JNC_ID))) = d
        types(CStr(hdr(i, JNC_ID))) = CStr(hdr(i, JNC_TYPE))
    Next i
    ReDim out(1 To UBound(lines, 1), 1 To 6)
    For i = 1 To UBound(lines, 1)
        j = CStr(lines(i, LNC_JOURNAL))
        If dates.Exists(j) Then out(i, 1) = dates(j)
        out(i, 2) = CStr(lines(i, LNC_BRANCH))
        out(i, 3) = CStr(lines(i, LNC_GL))
        out(i, 4) = modMoney.ToDec(lines(i, LNC_DEBIT))
        out(i, 5) = modMoney.ToDec(lines(i, LNC_CREDIT))
        If types.Exists(j) Then out(i, 6) = types(j)
    Next i
    LoadDatedLines = out
End Function

' Foreign groups: "BranchID|GLCode|Currency" -> Array(foreign signed, base signed). Debit counts plus, credit counts minus.
Public Function LoadFxNet() As Object
    Dim lines As Variant, d As Object, i As Long, key As String, f As Variant, b As Variant, cur As Variant, ccy As String
    Set d = CreateObject("Scripting.Dictionary")
    lines = LoadLines()
    If Not IsEmpty(lines) Then
        For i = 1 To UBound(lines, 1)
            ccy = CStr(lines(i, 6))
            If Len(ccy) > 0 And ccy <> modConstants.CCY_MMK Then
                key = CStr(lines(i, LNC_BRANCH)) & "|" & CStr(lines(i, LNC_GL)) & "|" & ccy
                b = modMoney.ToDec(lines(i, LNC_DEBIT)) - modMoney.ToDec(lines(i, LNC_CREDIT))
                f = modMoney.ToDec(lines(i, 10))
                If modMoney.ToDec(lines(i, LNC_CREDIT)) > 0 Then f = -f
                If d.Exists(key) Then
                    cur = d(key)
                    d(key) = Array(cur(0) + f, cur(1) + b)
                Else
                    d(key) = Array(f, b)
                End If
            End If
        Next i
    End If
    Set LoadFxNet = d
End Function
