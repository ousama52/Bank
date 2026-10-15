Attribute VB_Name = "repoClose"
Option Explicit

' tblDailyClose columns: BusinessDate, RunAt, AccountsAccrued, InterestAccrued, InterbankSettled, FDMatured,
' DormantMarked, ChecksPassed, Notes   (later stages add columns at the end; they are filled by header name)

Public Function Exists(ByVal d As Date) As Boolean
    Dim col As Variant, i As Long, x As Date
    col = modSheetUtils.ReadColumn(modConstants.TBL_DAILYCLOSE, "BusinessDate")
    If IsEmpty(col) Then Exit Function
    For i = 1 To UBound(col, 1)
        If modDateUtils.TryToDate(col(i, 1), x) Then
            If x = d Then
                Exists = True
                Exit Function
            End If
        End If
    Next i
End Function

' One row sized to the table. `stats` maps header names (AccountsAccrued, InterestAccrued, ...) to values.
Public Function NewCloseRow(ByVal d As Date, ByVal stats As Object, ByVal checksOk As Boolean, ByVal notes As String) As Variant
    Dim map As Object, n As Long, r() As Variant, k As Variant
    Set map = modSheetUtils.HeaderMap(modConstants.TBL_DAILYCLOSE)
    n = map.Count
    ReDim r(1 To n)
    r(map("BusinessDate")) = d
    r(map("RunAt")) = modDateUtils.TimeStamp()
    For Each k In stats.Keys
        If map.Exists(CStr(k)) Then r(map(CStr(k))) = stats(k)
    Next k
    If checksOk Then r(map("ChecksPassed")) = modConstants.FLAG_YES Else r(map("ChecksPassed")) = modConstants.FLAG_NO
    If Len(notes) > 0 Then r(map("Notes")) = notes
    NewCloseRow = r
End Function
