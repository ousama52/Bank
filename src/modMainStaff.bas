Attribute VB_Name = "modMainStaff"
Option Explicit

' Stage 5 entry macros: staff and payroll.
' Same pattern as modMain: current user, error handling, one service call, one message.

Public Sub RunGenerateSampleStaff()
    Dim result As clsResult
    On Error GoTo ErrHandler
    modPerformance.TurnOff
    Set result = svcStaff.GenerateSample(repoUser.Current(), modSheetUtils.ControlInputs())
    modMain.Finish result
Cleanup:
    On Error Resume Next
    modPerformance.TurnOn
    Exit Sub
ErrHandler:
    clsErrorHandler.LogError "RunGenerateSampleStaff", Err
    MsgBox "Error: " & Err.Description, vbCritical
    Resume Cleanup
End Sub

Public Sub RunHireEmployee()
    Dim result As clsResult
    On Error GoTo ErrHandler
    modPerformance.TurnOff
    Set result = svcStaff.Hire(repoUser.Current(), modSheetUtils.ControlInputs())
    modMain.Finish result
Cleanup:
    On Error Resume Next
    modPerformance.TurnOn
    Exit Sub
ErrHandler:
    clsErrorHandler.LogError "RunHireEmployee", Err
    MsgBox "Error: " & Err.Description, vbCritical
    Resume Cleanup
End Sub

Public Sub RunTransferEmployee()
    Dim result As clsResult
    On Error GoTo ErrHandler
    modPerformance.TurnOff
    Set result = svcStaff.TransferEmployee(repoUser.Current(), modSheetUtils.ControlInputs())
    modMain.Finish result
Cleanup:
    On Error Resume Next
    modPerformance.TurnOn
    Exit Sub
ErrHandler:
    clsErrorHandler.LogError "RunTransferEmployee", Err
    MsgBox "Error: " & Err.Description, vbCritical
    Resume Cleanup
End Sub

Public Sub RunChangePosition()
    Dim result As clsResult
    On Error GoTo ErrHandler
    modPerformance.TurnOff
    Set result = svcStaff.ChangePosition(repoUser.Current(), modSheetUtils.ControlInputs())
    modMain.Finish result
Cleanup:
    On Error Resume Next
    modPerformance.TurnOn
    Exit Sub
ErrHandler:
    clsErrorHandler.LogError "RunChangePosition", Err
    MsgBox "Error: " & Err.Description, vbCritical
    Resume Cleanup
End Sub

Public Sub RunChangeGrade()
    Dim result As clsResult
    On Error GoTo ErrHandler
    modPerformance.TurnOff
    Set result = svcStaff.ChangeGrade(repoUser.Current(), modSheetUtils.ControlInputs())
    modMain.Finish result
Cleanup:
    On Error Resume Next
    modPerformance.TurnOn
    Exit Sub
ErrHandler:
    clsErrorHandler.LogError "RunChangeGrade", Err
    MsgBox "Error: " & Err.Description, vbCritical
    Resume Cleanup
End Sub

Public Sub RunResignEmployee()
    Dim result As clsResult
    On Error GoTo ErrHandler
    modPerformance.TurnOff
    Set result = svcStaff.Resign(repoUser.Current(), modSheetUtils.ControlInputs())
    modMain.Finish result
Cleanup:
    On Error Resume Next
    modPerformance.TurnOn
    Exit Sub
ErrHandler:
    clsErrorHandler.LogError "RunResignEmployee", Err
    MsgBox "Error: " & Err.Description, vbCritical
    Resume Cleanup
End Sub

Public Sub RunTerminateEmployee()
    Dim result As clsResult
    On Error GoTo ErrHandler
    modPerformance.TurnOff
    Set result = svcStaff.Terminate(repoUser.Current(), modSheetUtils.ControlInputs())
    modMain.Finish result
Cleanup:
    On Error Resume Next
    modPerformance.TurnOn
    Exit Sub
ErrHandler:
    clsErrorHandler.LogError "RunTerminateEmployee", Err
    MsgBox "Error: " & Err.Description, vbCritical
    Resume Cleanup
End Sub

Public Sub RunIssueAdvance()
    Dim result As clsResult
    On Error GoTo ErrHandler
    modPerformance.TurnOff
    Set result = svcPayroll.IssueAdvance(repoUser.Current(), modSheetUtils.ControlInputs())
    modMain.Finish result
Cleanup:
    On Error Resume Next
    modPerformance.TurnOn
    Exit Sub
ErrHandler:
    clsErrorHandler.LogError "RunIssueAdvance", Err
    MsgBox "Error: " & Err.Description, vbCritical
    Resume Cleanup
End Sub

Public Sub RunPreparePayroll()
    Dim result As clsResult
    On Error GoTo ErrHandler
    modPerformance.TurnOff
    Set result = svcPayroll.PreparePayroll(repoUser.Current(), modSheetUtils.ControlInputs())
    modMain.Finish result
Cleanup:
    On Error Resume Next
    modPerformance.TurnOn
    Exit Sub
ErrHandler:
    clsErrorHandler.LogError "RunPreparePayroll", Err
    MsgBox "Error: " & Err.Description, vbCritical
    Resume Cleanup
End Sub

Public Sub RunPayPayroll()
    Dim result As clsResult
    On Error GoTo ErrHandler
    modPerformance.TurnOff
    Set result = svcPayroll.PayPayroll(repoUser.Current(), modSheetUtils.ControlInputs())
    modMain.Finish result
Cleanup:
    On Error Resume Next
    modPerformance.TurnOn
    Exit Sub
ErrHandler:
    clsErrorHandler.LogError "RunPayPayroll", Err
    MsgBox "Error: " & Err.Description, vbCritical
    Resume Cleanup
End Sub
