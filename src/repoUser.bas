Attribute VB_Name = "repoUser"
Option Explicit

' Columns of tblUsers: UserID, Name, Role, BranchID, IsActive

Public Function FindUser(ByVal userId As String) As clsUser
    Dim data As Variant, i As Long, u As clsUser
    data = modSheetUtils.ReadTable(modConstants.TBL_USERS)
    If IsEmpty(data) Then Exit Function
    For i = 1 To UBound(data, 1)
        If StrComp(CStr(data(i, 1)), userId, vbTextCompare) = 0 Then
            Set u = New clsUser
            u.UserID = CStr(data(i, 1))
            u.Name = CStr(data(i, 2))
            u.Role = CStr(data(i, 3))
            u.BranchID = CStr(data(i, 4))
            u.IsActive = (CStr(data(i, 5)) = modConstants.FLAG_YES)
            Set FindUser = u
            Exit Function
        End If
    Next i
End Function

' The user chosen in Control C4, or Nothing when blank, unknown or inactive.
' Also refreshes the Config cache, because every entry macro starts here.
Public Function Current() As clsUser
    Dim u As clsUser
    repoConfig.Reload
    Set u = FindUser(modSheetUtils.ControlText(modConstants.CTL_USER))
    If u Is Nothing Then Exit Function
    If u.IsActive Then Set Current = u
End Function
