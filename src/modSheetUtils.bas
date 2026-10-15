Attribute VB_Name = "modSheetUtils"
Option Explicit

' The only module (with modSetup) that touches worksheets. Repositories read and write
' whole tables through it, always as 2D arrays (1-based, rows x columns).
' A table whose only row has an empty first column holds "no data": ReadTable returns Empty.

Private Const ERR_BASE As Long = vbObjectError + 1000

' ---- Sheets and tables ----

Public Function SheetNameOf(ByVal tableName As String) As String
    If Left$(tableName, 3) = "tbl" Then
        SheetNameOf = Mid$(tableName, 4)
    Else
        SheetNameOf = tableName
    End If
End Function

Public Function SheetExists(ByVal sheetName As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    SheetExists = Not ws Is Nothing
End Function

Public Function GetSheet(ByVal sheetName As String) As Worksheet
    If Not SheetExists(sheetName) Then
        Err.Raise ERR_BASE + 1, "modSheetUtils.GetSheet", "Sheet " & sheetName & " not found. Run RunSetup first."
    End If
    Set GetSheet = ThisWorkbook.Worksheets(sheetName)
End Function

' Returns the sheet, adding it after the last sheet when it is missing.
Public Function EnsureSheet(ByVal sheetName As String) As Worksheet
    Dim ws As Worksheet
    If SheetExists(sheetName) Then
        Set EnsureSheet = ThisWorkbook.Worksheets(sheetName)
    Else
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = sheetName
        Set EnsureSheet = ws
    End If
End Function

Public Function TableExists(ByVal tableName As String) As Boolean
    Dim lo As ListObject
    If Not SheetExists(SheetNameOf(tableName)) Then Exit Function
    On Error Resume Next
    Set lo = ThisWorkbook.Worksheets(SheetNameOf(tableName)).ListObjects(tableName)
    On Error GoTo 0
    TableExists = Not lo Is Nothing
End Function

Public Function GetTable(ByVal tableName As String) As ListObject
    If Not TableExists(tableName) Then
        Err.Raise ERR_BASE + 2, "modSheetUtils.GetTable", "Table " & tableName & " not found. Run RunSetup first."
    End If
    Set GetTable = ThisWorkbook.Worksheets(SheetNameOf(tableName)).ListObjects(tableName)
End Function

' Header text -> 1-based column number (not case sensitive).
Public Function HeaderMap(ByVal tableName As String) As Object
    Dim lo As ListObject, d As Object, i As Long
    Set lo = GetTable(tableName)
    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = 1
    For i = 1 To lo.ListColumns.Count
        d(lo.ListColumns(i).Name) = i
    Next i
    Set HeaderMap = d
End Function

' ---- Reading ----

Private Function IsBlankTable(ByVal lo As ListObject) As Boolean
    Dim v As Variant
    If lo.DataBodyRange Is Nothing Then
        IsBlankTable = True
        Exit Function
    End If
    If lo.ListRows.Count <> 1 Then Exit Function
    v = lo.DataBodyRange.Cells(1, 1).Value
    If IsError(v) Then Exit Function
    IsBlankTable = (Len(CStr(v)) = 0)
End Function

' A 1 x 1 range returns a scalar from .Value; always hand back a 2D array.
Private Function AsArray2D(ByVal v As Variant) As Variant
    Dim a(1 To 1, 1 To 1) As Variant
    If IsArray(v) Then
        AsArray2D = v
    Else
        a(1, 1) = v
        AsArray2D = a
    End If
End Function

' The whole table body as a 2D array, or Empty when the table holds no data.
Public Function ReadTable(ByVal tableName As String) As Variant
    Dim lo As ListObject
    Set lo = GetTable(tableName)
    If IsBlankTable(lo) Then Exit Function
    ReadTable = AsArray2D(lo.DataBodyRange.Value)
End Function

' One column as an n x 1 array, or Empty when the table holds no data.
Public Function ReadColumn(ByVal tableName As String, ByVal columnName As String) As Variant
    Dim lo As ListObject
    Set lo = GetTable(tableName)
    If IsBlankTable(lo) Then Exit Function
    ReadColumn = AsArray2D(lo.ListColumns(columnName).DataBodyRange.Value)
End Function

' Number of rows of a ReadTable or ReadColumn result (0 for Empty).
Public Function RowCount(ByVal data As Variant) As Long
    If IsEmpty(data) Then Exit Function
    If Not IsArray(data) Then Exit Function
    RowCount = UBound(data, 1) - LBound(data, 1) + 1
End Function

' ---- Writing ----

Private Sub CheckShape(ByVal lo As ListObject, ByVal data As Variant)
    If Not IsArray(data) Then
        Err.Raise ERR_BASE + 3, "modSheetUtils", "Data for " & lo.Name & " must be a 2D array."
    End If
    If UBound(data, 2) - LBound(data, 2) + 1 <> lo.ListColumns.Count Then
        Err.Raise ERR_BASE + 4, "modSheetUtils", "Data for " & lo.Name & " has " & (UBound(data, 2) - LBound(data, 2) + 1) & _
            " columns, the table has " & lo.ListColumns.Count & "."
    End If
End Sub

' A table with all rows deleted gets one blank row back (an Excel Table needs one).
Private Sub EnsureBodyRow(ByVal lo As ListObject)
    If lo.DataBodyRange Is Nothing Then lo.ListRows.Add
End Sub

' Text that looks like a number with leading zeros (account numbers) must stay text: the column is set to text format before the write.
Private Sub KeepLeadingZeros(ByVal rng As Range, ByVal data As Variant)
    Dim r As Long, c As Long, v As Variant, s As String
    For c = LBound(data, 2) To UBound(data, 2)
        For r = LBound(data, 1) To UBound(data, 1)
            v = data(r, c)
            If VarType(v) = vbString Then
                s = CStr(v)
                If Len(s) > 1 And Left$(s, 1) = "0" And Not (s Like "*[!0-9]*") Then
                    rng.Columns(c - LBound(data, 2) + 1).NumberFormat = "@"
                    Exit For
                End If
            End If
        Next r
    Next c
End Sub

' Replaces the whole body in one assignment. Empty data leaves the single blank row.
Public Sub WriteTable(ByVal tableName As String, ByVal data As Variant)
    Dim lo As ListObject, n As Long, cur As Long
    Set lo = GetTable(tableName)
    EnsureBodyRow lo
    cur = lo.ListRows.Count
    If IsEmpty(data) Then
        lo.DataBodyRange.ClearContents
        If cur > 1 Then lo.Resize lo.Range.Resize(2)
        Exit Sub
    End If
    CheckShape lo, data
    n = UBound(data, 1) - LBound(data, 1) + 1
    If n < cur Then
        lo.DataBodyRange.Offset(n).Resize(cur - n).ClearContents
        lo.Resize lo.Range.Resize(n + 1)
    ElseIf n > cur Then
        lo.Resize lo.Range.Resize(n + 1)
    End If
    KeepLeadingZeros lo.DataBodyRange, data
    lo.DataBodyRange.Value = data
End Sub

' Adds rows at the end. A table that holds only the blank row is overwritten from its first row.
Public Sub AppendRows(ByVal tableName As String, ByVal data As Variant)
    Dim lo As ListObject, n As Long, k As Long, startRow As Long, total As Long
    Dim tgt As Range, c As Long
    If IsEmpty(data) Then Exit Sub
    Set lo = GetTable(tableName)
    CheckShape lo, data
    EnsureBodyRow lo
    n = UBound(data, 1) - LBound(data, 1) + 1
    k = lo.ListColumns.Count
    If IsBlankTable(lo) Then
        startRow = 1
    Else
        startRow = lo.ListRows.Count + 1
    End If
    total = startRow - 1 + n
    If total > lo.ListRows.Count Then lo.Resize lo.Range.Resize(total + 1)
    Set tgt = lo.DataBodyRange.Resize(n).Offset(startRow - 1)
    ' New rows copy the column format of the first body row (keeps IDs as text, dates as dates).
    If startRow > 1 Then
        For c = 1 To k
            tgt.Columns(c).NumberFormat = lo.DataBodyRange.Cells(1, c).NumberFormat
        Next c
    End If
    KeepLeadingZeros tgt, data
    tgt.Value = data
End Sub

' ---- Control sheet ----

Private Function CellText(ByVal ws As Worksheet, ByVal addr As String) As String
    Dim v As Variant
    v = ws.Range(addr).Value
    If IsError(v) Then Exit Function
    CellText = Trim$(CStr(v))
End Function

' Excel turns 0060000398 into a number; put the leading zeros back.
Private Function PadAccount(ByVal s As String) As String
    If Len(s) > 0 And Len(s) < 10 And Not (s Like "*[!0-9]*") Then
        PadAccount = Right$("0000000000" & s, 10)
    Else
        PadAccount = s
    End If
End Function

' All Control inputs in one dictionary. Amount and dates stay raw (Variant) for the service to check.
Public Function ControlInputs() As Object
    Dim ws As Worksheet, d As Object
    Set ws = GetSheet(modConstants.SHEET_CONTROL)
    Set d = CreateObject("Scripting.Dictionary")
    d(modConstants.INP_USER) = CellText(ws, modConstants.CTL_USER)
    d(modConstants.INP_ACCOUNT) = PadAccount(CellText(ws, modConstants.CTL_ACCOUNT))
    d(modConstants.INP_COUNTER) = PadAccount(CellText(ws, modConstants.CTL_COUNTER))
    d(modConstants.INP_AMOUNT) = ws.Range(modConstants.CTL_AMOUNT).Value
    d(modConstants.INP_BRANCH) = CellText(ws, modConstants.CTL_BRANCH)
    d(modConstants.INP_BANKCODE) = CellText(ws, modConstants.CTL_BANKCODE)
    d(modConstants.INP_TERM) = ws.Range(modConstants.CTL_TERM).Value
    d(modConstants.INP_PAYOUT) = PadAccount(CellText(ws, modConstants.CTL_PAYOUT))
    d(modConstants.INP_REFID) = CellText(ws, modConstants.CTL_REFID)
    d(modConstants.INP_CUSTOMER) = CellText(ws, modConstants.CTL_CUSTOMER)
    d(modConstants.INP_PRODUCT) = CellText(ws, modConstants.CTL_PRODUCT)
    d(modConstants.INP_NOTE) = CellText(ws, modConstants.CTL_NOTE)
    d(modConstants.INP_FROM) = ws.Range(modConstants.CTL_FROM).Value
    d(modConstants.INP_TO) = ws.Range(modConstants.CTL_TO).Value
    d(modConstants.INP_GL) = CellText(ws, modConstants.CTL_GL)
    d(modConstants.INP_VENDOR) = CellText(ws, modConstants.CTL_VENDOR)
    d(modConstants.INP_DESC) = CellText(ws, modConstants.CTL_DESC)
    d(modConstants.INP_ASSETTYPE) = CellText(ws, modConstants.CTL_ASSETTYPE)
    d(modConstants.INP_LIFE) = ws.Range(modConstants.CTL_LIFE).Value
    d(modConstants.INP_RESIDUAL) = ws.Range(modConstants.CTL_RESIDUAL).Value
    d(modConstants.INP_PERIOD) = CellText(ws, modConstants.CTL_PERIOD)
    d(modConstants.INP_LOANPRODUCT) = CellText(ws, modConstants.CTL_LOANPRODUCT)
    d(modConstants.INP_MODE) = CellText(ws, modConstants.CTL_MODE)
    d(modConstants.INP_LOANID) = CellText(ws, modConstants.CTL_LOANID)
    d(modConstants.INP_COLLATERAL) = CellText(ws, modConstants.CTL_COLLATERAL)
    d(modConstants.INP_EMPLOYEE) = CellText(ws, modConstants.CTL_EMPLOYEE)
    d(modConstants.INP_POSITION) = CellText(ws, modConstants.CTL_POSITION)
    d(modConstants.INP_GRADE) = CellText(ws, modConstants.CTL_GRADE)
    d(modConstants.INP_EFFDATE) = ws.Range(modConstants.CTL_EFFDATE).Value
    d(modConstants.INP_PAYRUN) = CellText(ws, modConstants.CTL_PAYRUN)
    d(modConstants.INP_CURRENCY) = CellText(ws, modConstants.CTL_CURRENCY)
    d(modConstants.INP_MIDRATE) = ws.Range(modConstants.CTL_MIDRATE).Value
    d(modConstants.INP_MERCHANT) = CellText(ws, modConstants.CTL_MERCHANT)
    d(modConstants.INP_AGENT) = CellText(ws, modConstants.CTL_AGENT)
    d(modConstants.INP_PARTYNAME) = CellText(ws, modConstants.CTL_PARTYNAME)
    d(modConstants.INP_FOREIGNBANK) = CellText(ws, modConstants.CTL_FOREIGNBANK)
    d(modConstants.INP_PARTYACCT) = CellText(ws, modConstants.CTL_PARTYACCT)
    d(modConstants.INP_PURPOSE) = CellText(ws, modConstants.CTL_PURPOSE)
    d(modConstants.INP_INVESTEE) = CellText(ws, modConstants.CTL_INVESTEE)
    d(modConstants.INP_INVESTMENT) = CellText(ws, modConstants.CTL_INVESTMENT)
    d(modConstants.INP_INVTYPE) = CellText(ws, modConstants.CTL_INVTYPE)
    d(modConstants.INP_RATESHARE) = ws.Range(modConstants.CTL_RATESHARE).Value
    d(modConstants.INP_PROFIT) = ws.Range(modConstants.CTL_PROFIT).Value
    d(modConstants.INP_ALERT) = CellText(ws, modConstants.CTL_ALERT)
    d(modConstants.INP_BONDTYPE) = CellText(ws, modConstants.CTL_BONDTYPE)
    d(modConstants.INP_FACE) = ws.Range(modConstants.CTL_FACE).Value
    d(modConstants.INP_TERMDAYS) = ws.Range(modConstants.CTL_TERMDAYS).Value
    d(modConstants.INP_BONDRATE) = ws.Range(modConstants.CTL_BONDRATE).Value
    d(modConstants.INP_REASON) = CellText(ws, modConstants.CTL_REASON)
    Set ControlInputs = d
End Function

Public Sub SetStatus(ByVal message As String)
    GetSheet(modConstants.SHEET_CONTROL).Range(modConstants.CTL_STATUS).Value = message
End Sub

Public Sub ShowBusinessDate(ByVal d As Date)
    GetSheet(modConstants.SHEET_CONTROL).Range(modConstants.CTL_BUSDATE).Value = d
End Sub

' One Control cell as trimmed text (for a single input such as the current user).
Public Function ControlText(ByVal addr As String) As String
    ControlText = CellText(GetSheet(modConstants.SHEET_CONTROL), addr)
End Function

' Turns a Collection of 1D row arrays (1-based, same length) into one 2D array for AppendRows.
Public Function RowsTo2D(ByVal rowsList As Collection) As Variant
    Dim out() As Variant, r As Long, c As Long, n As Long, k As Long, one As Variant
    n = rowsList.Count
    If n = 0 Then Exit Function
    one = rowsList(1)
    k = UBound(one) - LBound(one) + 1
    ReDim out(1 To n, 1 To k)
    For r = 1 To n
        one = rowsList(r)
        For c = 1 To k
            out(r, c) = one(LBound(one) + c - 1)
        Next c
    Next r
    RowsTo2D = out
End Function

' ---- Generated report sheets ----

' Clears a generated sheet. Refuses sheets that hold a table, so data can never be wiped by mistake.
Public Sub ClearReportSheet(ByVal sheetName As String)
    Dim ws As Worksheet
    Set ws = EnsureSheet(sheetName)
    If ws.ListObjects.Count > 0 Then
        Err.Raise ERR_BASE + 5, "modSheetUtils.ClearReportSheet", "Sheet " & sheetName & " holds a table and is not a report sheet."
    End If
    ws.Cells.Clear
End Sub

' Clears a report sheet and makes the listed columns (e.g. "1,2") text so GL codes and account numbers keep their zeros.
Public Sub PrepareReport(ByVal sheetName As String, ByVal textCols As String)
    Dim ws As Worksheet, p As Variant
    ClearReportSheet sheetName
    Set ws = GetSheet(sheetName)
    If Len(textCols) > 0 Then
        For Each p In Split(textCols, ",")
            ws.Columns(CLng(p)).NumberFormat = "@"
        Next p
    End If
End Sub

Public Sub SetCell(ByVal sheetName As String, ByVal row As Long, ByVal col As Long, ByVal v As Variant, Optional ByVal bold As Boolean = False)
    With GetSheet(sheetName).Cells(row, col)
        .Value = v
        .Font.Bold = bold
    End With
End Sub

' Styles the header row of a block, formats money columns ("3,4") and fits the column widths.
Public Sub StyleReport(ByVal sheetName As String, ByVal headerRow As Long, ByVal lastCol As Long, ByVal moneyCols As String)
    Dim ws As Worksheet, p As Variant
    Set ws = GetSheet(sheetName)
    With ws.Range(ws.Cells(headerRow, 1), ws.Cells(headerRow, lastCol))
        .Font.Bold = True
        .Interior.Color = RGB(31, 78, 121)
        .Font.Color = RGB(255, 255, 255)
    End With
    If Len(moneyCols) > 0 Then
        For Each p In Split(moneyCols, ",")
            ws.Range(ws.Cells(headerRow + 1, CLng(p)), ws.Cells(ws.Rows.Count, CLng(p))).NumberFormat = "#,##0"
        Next p
    End If
    ws.Range(ws.Columns(1), ws.Columns(lastCol)).AutoFit
End Sub

' Writes a 2D array into a sheet in one assignment.
Public Sub WriteGrid(ByVal sheetName As String, ByVal topRow As Long, ByVal leftCol As Long, ByVal data As Variant)
    Dim ws As Worksheet
    If IsEmpty(data) Then Exit Sub
    Set ws = GetSheet(sheetName)
    ws.Cells(topRow, leftCol).Resize(UBound(data, 1) - LBound(data, 1) + 1, UBound(data, 2) - LBound(data, 2) + 1).Value = data
End Sub
