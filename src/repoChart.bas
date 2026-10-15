Attribute VB_Name = "repoChart"
Option Explicit

' Columns of tblChart: GLCode, Name, Type, NormalSide, Stage, IsActive

' Dictionary GLCode -> Array(Name, Type, NormalSide, IsActive as Boolean).
Public Function LoadChart() As Object
    Dim data As Variant, i As Long, d As Object
    Set d = CreateObject("Scripting.Dictionary")
    data = modSheetUtils.ReadTable(modConstants.TBL_CHART)
    If Not IsEmpty(data) Then
        For i = 1 To UBound(data, 1)
            If Len(CStr(data(i, 1))) > 0 Then
                d(CStr(data(i, 1))) = Array(CStr(data(i, 2)), CStr(data(i, 3)), CStr(data(i, 4)), (CStr(data(i, 6)) = modConstants.FLAG_YES))
            End If
        Next i
    End If
    Set LoadChart = d
End Function
