Attribute VB_Name = "repoBranch"
Option Explicit

' Columns of tblBranches: BranchID, Name, Type, City, Region, PremisesType, OpenDate, Status, VaultMin, VaultMax

' Dictionary BranchID -> clsBranch (in table order).
Public Function LoadBranches() As Object
    Dim data As Variant, i As Long, b As clsBranch, d As Object
    Set d = CreateObject("Scripting.Dictionary")
    data = modSheetUtils.ReadTable(modConstants.TBL_BRANCHES)
    If Not IsEmpty(data) Then
        For i = 1 To UBound(data, 1)
            If Len(CStr(data(i, 1))) > 0 Then
                Set b = New clsBranch
                b.BranchID = CStr(data(i, 1))
                b.Name = CStr(data(i, 2))
                b.BranchType = CStr(data(i, 3))
                b.Status = CStr(data(i, 8))
                b.VaultMin = modMoney.ToDec(data(i, 9))
                b.VaultMax = modMoney.ToDec(data(i, 10))
                d.Add b.BranchID, b
            End If
        Next i
    End If
    Set LoadBranches = d
End Function

Public Function FindBranch(ByVal branchId As String) As clsBranch
    Dim d As Object
    Set d = LoadBranches()
    If d.Exists(branchId) Then Set FindBranch = d(branchId)
End Function
