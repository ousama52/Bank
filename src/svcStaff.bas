Attribute VB_Name = "svcStaff"
Option Explicit

' Stage 5: staff register. Every change writes tblEmployeeHistory and an audit row. HeadOffice runs the HR macros.

Private Function NeedTables() As String
    If Not modSheetUtils.TableExists(modConstants.TBL_EMPLOYEES) Or Not modSheetUtils.TableExists(modConstants.TBL_POSITIONS) Then NeedTables = modMessages.MSG_SETUP_FIRST
End Function

Private Function PositionExists(ByVal code As String) As Boolean
    Dim data As Variant, i As Long
    data = modSheetUtils.ReadTable(modConstants.TBL_POSITIONS)
    If IsEmpty(data) Then Exit Function
    For i = 1 To UBound(data, 1)
        If StrComp(CStr(data(i, 1)), code, vbTextCompare) = 0 And CStr(data(i, 7)) <> modConstants.FLAG_NO Then PositionExists = True: Exit Function
    Next i
End Function

Private Sub AddHistory(ByVal uow As clsUnitOfWork, ByVal employeeId As String, ByVal d As Date, ByVal changeType As String, ByVal fromBranch As String, ByVal toBranch As String, _
                       ByVal fromPos As String, ByVal toPos As String, ByVal fromGrade As Variant, ByVal toGrade As Variant, ByVal note As String)
    uow.AppendRow modConstants.TBL_EMPHISTORY, repoStaff.HistoryRow(uow.NewId(modConstants.TBL_EMPHISTORY, "HistoryID", "HI", 6), employeeId, d, changeType, fromBranch, toBranch, fromPos, toPos, fromGrade, toGrade, note)
    repoAudit.Add uow, "Staff" & changeType, employeeId, note
End Sub

' True when another Active branch manager stays in the branch (used for the warning).
Private Function OtherManagerExists(ByVal emps As Variant, ByVal row As Long) As Boolean
    Dim i As Long
    For i = 1 To UBound(emps, 1)
        If i <> row And CStr(emps(i, repoStaff.EMC_BRANCH)) = CStr(emps(row, repoStaff.EMC_BRANCH)) And CStr(emps(i, repoStaff.EMC_POSITION)) = "BM" And CStr(emps(i, repoStaff.EMC_STATUS)) = "Active" Then
            OtherManagerExists = True
            Exit Function
        End If
    Next i
End Function

Private Function ManagerWarning(ByVal emps As Variant, ByVal row As Long) As String
    If CStr(emps(row, repoStaff.EMC_POSITION)) = "BM" Then
        If Not OtherManagerExists(emps, row) Then ManagerWarning = " Warning: branch " & CStr(emps(row, repoStaff.EMC_BRANCH)) & " now has no branch manager."
    End If
End Function

Private Function Prepare(ByVal user As clsUser, ByVal inp As Object, ByRef uow As clsUnitOfWork, ByRef emps As Variant, ByRef row As Long, ByRef d As Date, ByRef msg As String) As Boolean
    msg = svcSecurity.Denied(user, svcSecurity.ACT_HR)
    If Len(msg) = 0 Then msg = NeedTables()
    If Len(msg) > 0 Then Exit Function
    Set uow = New clsUnitOfWork
    uow.Init repoConfig.BusinessDate(), user.UserID
    emps = uow.WorkingTable(modConstants.TBL_EMPLOYEES)
    row = repoStaff.FindRowById(emps, repoStaff.EMC_ID, CStr(inp(modConstants.INP_EMPLOYEE)))
    If row = 0 Then msg = "Unknown employee. Put the EmployeeID in Control C30.": Exit Function
    If Not modDateUtils.TryToDate(inp(modConstants.INP_EFFDATE), d) Then d = uow.BusinessDate
    Prepare = True
End Function

Private Sub DeactivateUser(ByVal uow As clsUnitOfWork, ByVal userId As String)
    Dim users As Variant, i As Long
    If Len(userId) = 0 Then Exit Sub
    users = uow.WorkingTable(modConstants.TBL_USERS)
    If IsEmpty(users) Then Exit Sub
    For i = 1 To UBound(users, 1)
        If StrComp(CStr(users(i, 1)), userId, vbTextCompare) = 0 Then
            users(i, 5) = modConstants.FLAG_NO
            uow.ReplaceTable modConstants.TBL_USERS, users
            repoAudit.Add uow, "UserDeactivated", userId, "Employee left"
            Exit Sub
        End If
    Next i
End Sub

' ---- Hire ----

' Control: C20 name, C31 position, C32 grade, C8 branch, C33 hire date, C10 dependents, C5 bank account (blank = paid externally), C15 phone.
Public Function Hire(ByVal user As clsUser, ByVal inp As Object) As clsResult
    Dim uow As clsUnitOfWork, msg As String, emps As Variant, id As String, g As Variant, d As Date, b As String, a As clsAccount, mode As String, dep As Variant, acct As String
    msg = svcSecurity.Denied(user, svcSecurity.ACT_HR)
    If Len(msg) = 0 Then msg = NeedTables()
    If Len(msg) > 0 Then Set Hire = modMessages.FailResult(msg): Exit Function
    If Len(CStr(inp(modConstants.INP_VENDOR))) = 0 Then Set Hire = modMessages.FailResult("Enter the employee name in Control C20."): Exit Function
    If Not PositionExists(CStr(inp(modConstants.INP_POSITION))) Then Set Hire = modMessages.FailResult("Unknown position code (Control C31)."): Exit Function
    If Not modMoney.TryToDec(inp(modConstants.INP_GRADE), g) Then g = CDec(1)
    If g < 1 Or g > 3 Then Set Hire = modMessages.FailResult("Grade must be 1, 2 or 3 (Control C32)."): Exit Function
    If Not modMoney.TryToDec(inp(modConstants.INP_TERM), dep) Then dep = CDec(0)
    Set uow = New clsUnitOfWork
    uow.Init repoConfig.BusinessDate(), user.UserID
    b = Trim$(CStr(inp(modConstants.INP_BRANCH)))
    If Not uow.Branches().Exists(b) Then Set Hire = modMessages.FailResult("Choose the branch in Control C8."): Exit Function
    If Not modDateUtils.TryToDate(inp(modConstants.INP_EFFDATE), d) Then d = uow.BusinessDate
    acct = CStr(inp(modConstants.INP_ACCOUNT))
    mode = "External"
    If Len(acct) > 0 Then
        Set a = uow.GetAccount(acct)
        If a Is Nothing Then Set Hire = modMessages.FailResult(modMessages.UnknownAccount(acct)): Exit Function
        If a.Status <> modConstants.STATUS_ACTIVE Or a.Ccy <> modConstants.CCY_MMK Then Set Hire = modMessages.FailResult("The salary account must be an Active MMK account."): Exit Function
        mode = "Account"
    End If
    emps = uow.WorkingTable(modConstants.TBL_EMPLOYEES)
    id = uow.NewId(modConstants.TBL_EMPLOYEES, "EmployeeID", "E", 5)
    uow.AppendRow modConstants.TBL_EMPLOYEES, repoStaff.EmployeeRow(id, CStr(inp(modConstants.INP_VENDOR)), CStr(inp(modConstants.INP_POSITION)), CLng(g), b, d, mode, acct, CLng(dep), CStr(inp(modConstants.INP_NOTE)))
    AddHistory uow, id, d, "Hire", "", b, "", CStr(inp(modConstants.INP_POSITION)), Empty, g, "Hired " & CStr(inp(modConstants.INP_VENDOR))
    uow.Commit
    Set Hire = modMessages.OkResult("Employee " & id & " hired at " & b & " as " & CStr(inp(modConstants.INP_POSITION)) & " grade " & CLng(g) & ".", 1, id)
End Function

' ---- Changes ----

Public Function TransferEmployee(ByVal user As clsUser, ByVal inp As Object) As clsResult
    Dim uow As clsUnitOfWork, msg As String, emps As Variant, row As Long, d As Date, b As String, fromB As String, warn As String
    If Not Prepare(user, inp, uow, emps, row, d, msg) Then Set TransferEmployee = modMessages.FailResult(msg): Exit Function
    b = Trim$(CStr(inp(modConstants.INP_BRANCH)))
    If Not uow.Branches().Exists(b) Then Set TransferEmployee = modMessages.FailResult("Choose the new branch in Control C8."): Exit Function
    If CStr(emps(row, repoStaff.EMC_STATUS)) <> "Active" And CStr(emps(row, repoStaff.EMC_STATUS)) <> "OnLeave" Then Set TransferEmployee = modMessages.FailResult("Only a current employee can be transferred."): Exit Function
    fromB = CStr(emps(row, repoStaff.EMC_BRANCH))
    If fromB = b Then Set TransferEmployee = modMessages.FailResult("The employee already works at " & b & "."): Exit Function
    warn = ManagerWarning(emps, row)
    emps(row, repoStaff.EMC_BRANCH) = b
    uow.ReplaceTable modConstants.TBL_EMPLOYEES, emps
    AddHistory uow, CStr(emps(row, repoStaff.EMC_ID)), d, "Transfer", fromB, b, "", "", Empty, Empty, "Transfer " & fromB & " to " & b
    uow.Commit
    Set TransferEmployee = modMessages.OkResult("Employee " & CStr(emps(row, repoStaff.EMC_ID)) & " transferred from " & fromB & " to " & b & "." & warn, 1)
End Function

Public Function ChangePosition(ByVal user As clsUser, ByVal inp As Object) As clsResult
    Dim uow As clsUnitOfWork, msg As String, emps As Variant, row As Long, d As Date, newPos As String, g As Variant, oldG As Variant, oldPos As String
    If Not Prepare(user, inp, uow, emps, row, d, msg) Then Set ChangePosition = modMessages.FailResult(msg): Exit Function
    newPos = CStr(inp(modConstants.INP_POSITION))
    If Not PositionExists(newPos) Then Set ChangePosition = modMessages.FailResult("Unknown position code (Control C31)."): Exit Function
    oldPos = CStr(emps(row, repoStaff.EMC_POSITION))
    oldG = modMoney.ToDec(emps(row, repoStaff.EMC_GRADE))
    If Not modMoney.TryToDec(inp(modConstants.INP_GRADE), g) Then g = oldG
    If g < 1 Or g > 3 Then Set ChangePosition = modMessages.FailResult("Grade must be 1, 2 or 3 (Control C32)."): Exit Function
    msg = ManagerWarning(emps, row)
    If newPos = "BM" Then msg = ""
    emps(row, repoStaff.EMC_POSITION) = newPos
    emps(row, repoStaff.EMC_GRADE) = g
    uow.ReplaceTable modConstants.TBL_EMPLOYEES, emps
    AddHistory uow, CStr(emps(row, repoStaff.EMC_ID)), d, "Promotion", "", "", oldPos, newPos, oldG, g, "Position " & oldPos & " to " & newPos
    uow.Commit
    Set ChangePosition = modMessages.OkResult("Employee " & CStr(emps(row, repoStaff.EMC_ID)) & " is now " & newPos & " grade " & CLng(g) & "." & IIf(oldPos = "BM" And newPos <> "BM", msg, ""), 1)
End Function

Public Function ChangeGrade(ByVal user As clsUser, ByVal inp As Object) As clsResult
    Dim uow As clsUnitOfWork, msg As String, emps As Variant, row As Long, d As Date, g As Variant, oldG As Variant
    If Not Prepare(user, inp, uow, emps, row, d, msg) Then Set ChangeGrade = modMessages.FailResult(msg): Exit Function
    If Not modMoney.TryToDec(inp(modConstants.INP_GRADE), g) Then Set ChangeGrade = modMessages.FailResult("Enter the new grade in Control C32."): Exit Function
    If g < 1 Or g > 3 Then Set ChangeGrade = modMessages.FailResult("Grade must be 1, 2 or 3."): Exit Function
    oldG = modMoney.ToDec(emps(row, repoStaff.EMC_GRADE))
    emps(row, repoStaff.EMC_GRADE) = g
    uow.ReplaceTable modConstants.TBL_EMPLOYEES, emps
    AddHistory uow, CStr(emps(row, repoStaff.EMC_ID)), d, "GradeChange", "", "", "", "", oldG, g, "Grade " & CLng(oldG) & " to " & CLng(g)
    uow.Commit
    Set ChangeGrade = modMessages.OkResult("Employee " & CStr(emps(row, repoStaff.EMC_ID)) & " is now grade " & CLng(g) & ".", 1)
End Function

Private Function Leave(ByVal user As clsUser, ByVal inp As Object, ByVal newStatus As String, ByVal changeType As String) As clsResult
    Dim uow As clsUnitOfWork, msg As String, emps As Variant, row As Long, d As Date, warn As String
    If Not Prepare(user, inp, uow, emps, row, d, msg) Then Set Leave = modMessages.FailResult(msg): Exit Function
    If CStr(emps(row, repoStaff.EMC_STATUS)) = "Resigned" Or CStr(emps(row, repoStaff.EMC_STATUS)) = "Terminated" Then Set Leave = modMessages.FailResult("The employee has already left."): Exit Function
    warn = ManagerWarning(emps, row)
    emps(row, repoStaff.EMC_STATUS) = newStatus
    emps(row, repoStaff.EMC_EXIT) = d
    uow.ReplaceTable modConstants.TBL_EMPLOYEES, emps
    AddHistory uow, CStr(emps(row, repoStaff.EMC_ID)), d, changeType, CStr(emps(row, repoStaff.EMC_BRANCH)), "", CStr(emps(row, repoStaff.EMC_POSITION)), "", Empty, Empty, changeType & " " & modDateUtils.ToIso(d)
    DeactivateUser uow, CStr(emps(row, repoStaff.EMC_USER))
    uow.Commit
    Set Leave = modMessages.OkResult("Employee " & CStr(emps(row, repoStaff.EMC_ID)) & " is now " & newStatus & " from " & modDateUtils.ToIso(d) & "." & warn, 1)
End Function

Public Function Resign(ByVal user As clsUser, ByVal inp As Object) As clsResult
    Set Resign = Leave(user, inp, "Resigned", "Resign")
End Function

Public Function Terminate(ByVal user As clsUser, ByVal inp As Object) As clsResult
    Set Terminate = Leave(user, inp, "Terminated", "Terminate")
End Function

' ---- Sample staff ----

Private Function SampleName(ByVal i As Long) As String
    Dim first As Variant, last As Variant
    first = Array("Aung", "Kyaw", "Zaw", "Htet", "Min", "Thura", "Naing", "Soe", "Myo", "Hla", "Tun", "Phyo", "Ye", "Wai", "Lin", "Su", "Thida", "Mya", "Khin", "Nilar", "Aye", "Moe", "Hnin", "Thandar", "Yin", "Ei", "Cho", "Shwe", "Nandar", "Pyae")
    last = Array("Htut", "Oo", "Lwin", "Naing", "Win", "Tun", "Aye", "Kyaw", "Maung", "Thein", "Zin", "Han", "Linn", "Sett", "Paing", "Hein", "Latt", "Than", "Khant", "Mon", "Yadanar", "Swe", "Nyein", "Aung", "Phyu", "San", "Kyi", "Mar", "Sandar", "Thu")
    SampleName = first(i Mod 30) & " " & last((i \ 30) Mod 30) & " " & Chr$(65 + (i \ 900) Mod 26)
End Function

' Admin only, only when tblEmployees is empty: exactly the staffing guide minimum for every branch (242 employees).
Public Function GenerateSample(ByVal user As clsUser, ByVal inp As Object) As clsResult
    Dim uow As clsUnitOfWork, msg As String, branches As Variant, pos As Variant, i As Long, j As Long, n As Long, k As Long, minCol As Long, id As String, g As Long, hire As Date, seq As Long
    msg = svcSecurity.Denied(user, svcSecurity.ACT_SAMPLESTAFF)
    If Len(msg) = 0 Then msg = NeedTables()
    If Len(msg) > 0 Then Set GenerateSample = modMessages.FailResult(msg): Exit Function
    Set uow = New clsUnitOfWork
    uow.Init repoConfig.BusinessDate(), user.UserID
    If Not IsEmpty(uow.WorkingTable(modConstants.TBL_EMPLOYEES)) Then Set GenerateSample = modMessages.FailResult("tblEmployees already holds staff. Sample staff is only generated into an empty table."): Exit Function
    branches = modSheetUtils.ReadTable(modConstants.TBL_BRANCHES)
    pos = modSheetUtils.ReadTable(modConstants.TBL_POSITIONS)
    If IsEmpty(branches) Or IsEmpty(pos) Then Set GenerateSample = modMessages.FailResult("Branches or positions are missing."): Exit Function
    For i = 1 To UBound(branches, 1)
        Select Case CStr(branches(i, 3))
            Case modConstants.BT_HEADOFFICE: minCol = 4
            Case modConstants.BT_CITY: minCol = 5
            Case Else: minCol = 6
        End Select
        If Not modDateUtils.TryToDate(branches(i, 7), hire) Then hire = uow.BusinessDate
        For j = 1 To UBound(pos, 1)
            If CStr(pos(j, 7)) <> modConstants.FLAG_NO Then
                n = CLng(modMoney.ToDec(pos(j, minCol)))
                For k = 1 To n
                    seq = seq + 1
                    id = modIds.FormatId("E", seq, 5)
                    Select Case CStr(pos(j, 1))
                        Case "GM": g = 3
                        Case "BM", "ABM", "HOD": g = 2
                        Case Else: g = 1
                    End Select
                    uow.AppendRow modConstants.TBL_EMPLOYEES, repoStaff.EmployeeRow(id, SampleName(seq), CStr(pos(j, 1)), g, CStr(branches(i, 1)), hire, "External", "", 0, "")
                    uow.AppendRow modConstants.TBL_EMPHISTORY, repoStaff.HistoryRow(modIds.FormatId("HI", seq, 6), id, hire, "Hire", "", CStr(branches(i, 1)), "", CStr(pos(j, 1)), Empty, g, "Sample staff")
                Next k
            End If
        Next j
    Next i
    repoAudit.Add uow, "GenerateSampleStaff", modConstants.TBL_EMPLOYEES, seq & " employees"
    uow.Commit
    Set GenerateSample = modMessages.OkResult(seq & " sample employees created (the staffing guide minimum for every branch).", seq)
End Function
