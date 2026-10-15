Attribute VB_Name = "modMainGov"
Option Explicit

' Stage 8 entry macros: policy rate, securities, compliance, AML, inspection, fines, tax and year end.
' Same pattern as modMain: current user, error handling, one service call, one message.

Public Sub RunApplyPolicyRate()
    Dim result As clsResult
    On Error GoTo ErrHandler
    modPerformance.TurnOff
    Set result = svcPolicy.ApplyPolicyRate(repoUser.Current(), modSheetUtils.ControlInputs())
    modMain.Finish result
Cleanup:
    On Error Resume Next
    modPerformance.TurnOn
    Exit Sub
ErrHandler:
    clsErrorHandler.LogError "RunApplyPolicyRate", Err
    MsgBox "Error: " & Err.Description, vbCritical
    Resume Cleanup
End Sub

Public Sub RunBuyGovSecurity()
    Dim result As clsResult
    On Error GoTo ErrHandler
    modPerformance.TurnOff
    Set result = svcBonds.BuyGovSecurity(repoUser.Current(), modSheetUtils.ControlInputs())
    modMain.Finish result
Cleanup:
    On Error Resume Next
    modPerformance.TurnOn
    Exit Sub
ErrHandler:
    clsErrorHandler.LogError "RunBuyGovSecurity", Err
    MsgBox "Error: " & Err.Description, vbCritical
    Resume Cleanup
End Sub

Public Sub RunComplianceCheck()
    Dim result As clsResult
    On Error GoTo ErrHandler
    modPerformance.TurnOff
    Set result = svcRatios.CheckCompliance(repoUser.Current(), modSheetUtils.ControlInputs())
    modMain.Finish result
Cleanup:
    On Error Resume Next
    modPerformance.TurnOn
    Exit Sub
ErrHandler:
    clsErrorHandler.LogError "RunComplianceCheck", Err
    MsgBox "Error: " & Err.Description, vbCritical
    Resume Cleanup
End Sub

Public Sub RunReviewAlert()
    Dim result As clsResult
    On Error GoTo ErrHandler
    modPerformance.TurnOff
    Set result = svcAml.ReviewAlert(repoUser.Current(), modSheetUtils.ControlInputs())
    modMain.Finish result
Cleanup:
    On Error Resume Next
    modPerformance.TurnOn
    Exit Sub
ErrHandler:
    clsErrorHandler.LogError "RunReviewAlert", Err
    MsgBox "Error: " & Err.Description, vbCritical
    Resume Cleanup
End Sub

Public Sub RunOpenCase()
    Dim result As clsResult
    On Error GoTo ErrHandler
    modPerformance.TurnOff
    Set result = svcAml.OpenCase(repoUser.Current(), modSheetUtils.ControlInputs())
    modMain.Finish result
Cleanup:
    On Error Resume Next
    modPerformance.TurnOn
    Exit Sub
ErrHandler:
    clsErrorHandler.LogError "RunOpenCase", Err
    MsgBox "Error: " & Err.Description, vbCritical
    Resume Cleanup
End Sub

Public Sub RunCloseCase()
    Dim result As clsResult
    On Error GoTo ErrHandler
    modPerformance.TurnOff
    Set result = svcAml.CloseCase(repoUser.Current(), modSheetUtils.ControlInputs())
    modMain.Finish result
Cleanup:
    On Error Resume Next
    modPerformance.TurnOn
    Exit Sub
ErrHandler:
    clsErrorHandler.LogError "RunCloseCase", Err
    MsgBox "Error: " & Err.Description, vbCritical
    Resume Cleanup
End Sub

Public Sub RunReportCase()
    Dim result As clsResult
    On Error GoTo ErrHandler
    modPerformance.TurnOff
    Set result = svcAml.ReportCase(repoUser.Current(), modSheetUtils.ControlInputs())
    modMain.Finish result
Cleanup:
    On Error Resume Next
    modPerformance.TurnOn
    Exit Sub
ErrHandler:
    clsErrorHandler.LogError "RunReportCase", Err
    MsgBox "Error: " & Err.Description, vbCritical
    Resume Cleanup
End Sub

Public Sub RunInspection()
    Dim result As clsResult
    On Error GoTo ErrHandler
    modPerformance.TurnOff
    Set result = svcInspection.Inspect(repoUser.Current(), modSheetUtils.ControlInputs())
    modMain.Finish result
Cleanup:
    On Error Resume Next
    modPerformance.TurnOn
    Exit Sub
ErrHandler:
    clsErrorHandler.LogError "RunInspection", Err
    MsgBox "Error: " & Err.Description, vbCritical
    Resume Cleanup
End Sub

Public Sub RunIssueFine()
    Dim result As clsResult
    On Error GoTo ErrHandler
    modPerformance.TurnOff
    Set result = svcInspection.IssueFine(repoUser.Current(), modSheetUtils.ControlInputs())
    modMain.Finish result
Cleanup:
    On Error Resume Next
    modPerformance.TurnOn
    Exit Sub
ErrHandler:
    clsErrorHandler.LogError "RunIssueFine", Err
    MsgBox "Error: " & Err.Description, vbCritical
    Resume Cleanup
End Sub

Public Sub RunTaxReturn()
    Dim result As clsResult
    On Error GoTo ErrHandler
    modPerformance.TurnOff
    Set result = svcTaxReturn.PrepareTaxReturn(repoUser.Current(), modSheetUtils.ControlInputs())
    modMain.Finish result
Cleanup:
    On Error Resume Next
    modPerformance.TurnOn
    Exit Sub
ErrHandler:
    clsErrorHandler.LogError "RunTaxReturn", Err
    MsgBox "Error: " & Err.Description, vbCritical
    Resume Cleanup
End Sub

Public Sub RunBookIncomeTax()
    Dim result As clsResult
    On Error GoTo ErrHandler
    modPerformance.TurnOff
    Set result = svcTaxReturn.RequestIncomeTax(repoUser.Current(), modSheetUtils.ControlInputs())
    modMain.Finish result
Cleanup:
    On Error Resume Next
    modPerformance.TurnOn
    Exit Sub
ErrHandler:
    clsErrorHandler.LogError "RunBookIncomeTax", Err
    MsgBox "Error: " & Err.Description, vbCritical
    Resume Cleanup
End Sub

Public Sub RunYearEndClose()
    Dim result As clsResult
    On Error GoTo ErrHandler
    modPerformance.TurnOff
    Set result = svcYearEnd.RequestYearEnd(repoUser.Current(), modSheetUtils.ControlInputs())
    modMain.Finish result
Cleanup:
    On Error Resume Next
    modPerformance.TurnOn
    Exit Sub
ErrHandler:
    clsErrorHandler.LogError "RunYearEndClose", Err
    MsgBox "Error: " & Err.Description, vbCritical
    Resume Cleanup
End Sub
