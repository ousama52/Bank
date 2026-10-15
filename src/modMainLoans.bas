Attribute VB_Name = "modMainLoans"
Option Explicit

' Stage 3 entry macros: loans.
' Same pattern as modMain: current user, error handling, one service call, one message.

Public Sub RunQuoteLoan()
    Dim result As clsResult
    On Error GoTo ErrHandler
    modPerformance.TurnOff
    Set result = svcLoans.Quote(repoUser.Current(), modSheetUtils.ControlInputs())
    modMain.Finish result
Cleanup:
    On Error Resume Next
    modPerformance.TurnOn
    Exit Sub
ErrHandler:
    clsErrorHandler.LogError "RunQuoteLoan", Err
    MsgBox "Error: " & Err.Description, vbCritical
    Resume Cleanup
End Sub

Public Sub RunApplyLoan()
    Dim result As clsResult
    On Error GoTo ErrHandler
    modPerformance.TurnOff
    Set result = svcLoans.Apply(repoUser.Current(), modSheetUtils.ControlInputs())
    modMain.Finish result
Cleanup:
    On Error Resume Next
    modPerformance.TurnOn
    Exit Sub
ErrHandler:
    clsErrorHandler.LogError "RunApplyLoan", Err
    MsgBox "Error: " & Err.Description, vbCritical
    Resume Cleanup
End Sub

Public Sub RunDisburseLoan()
    Dim result As clsResult
    On Error GoTo ErrHandler
    modPerformance.TurnOff
    Set result = svcLoans.Disburse(repoUser.Current(), modSheetUtils.ControlInputs())
    modMain.Finish result
Cleanup:
    On Error Resume Next
    modPerformance.TurnOn
    Exit Sub
ErrHandler:
    clsErrorHandler.LogError "RunDisburseLoan", Err
    MsgBox "Error: " & Err.Description, vbCritical
    Resume Cleanup
End Sub

Public Sub RunLoanPayment()
    Dim result As clsResult
    On Error GoTo ErrHandler
    modPerformance.TurnOff
    Set result = svcLoanPayments.Pay(repoUser.Current(), modSheetUtils.ControlInputs())
    modMain.Finish result
Cleanup:
    On Error Resume Next
    modPerformance.TurnOn
    Exit Sub
ErrHandler:
    clsErrorHandler.LogError "RunLoanPayment", Err
    MsgBox "Error: " & Err.Description, vbCritical
    Resume Cleanup
End Sub

Public Sub RunPayoffLoan()
    Dim result As clsResult
    On Error GoTo ErrHandler
    modPerformance.TurnOff
    Set result = svcLoanPayments.Payoff(repoUser.Current(), modSheetUtils.ControlInputs())
    modMain.Finish result
Cleanup:
    On Error Resume Next
    modPerformance.TurnOn
    Exit Sub
ErrHandler:
    clsErrorHandler.LogError "RunPayoffLoan", Err
    MsgBox "Error: " & Err.Description, vbCritical
    Resume Cleanup
End Sub

Public Sub RunWriteOffLoan()
    Dim result As clsResult
    On Error GoTo ErrHandler
    modPerformance.TurnOff
    Set result = svcLoanPayments.WriteOff(repoUser.Current(), modSheetUtils.ControlInputs())
    modMain.Finish result
Cleanup:
    On Error Resume Next
    modPerformance.TurnOn
    Exit Sub
ErrHandler:
    clsErrorHandler.LogError "RunWriteOffLoan", Err
    MsgBox "Error: " & Err.Description, vbCritical
    Resume Cleanup
End Sub

Public Sub RunRecoverLoan()
    Dim result As clsResult
    On Error GoTo ErrHandler
    modPerformance.TurnOff
    Set result = svcLoanPayments.Recover(repoUser.Current(), modSheetUtils.ControlInputs())
    modMain.Finish result
Cleanup:
    On Error Resume Next
    modPerformance.TurnOn
    Exit Sub
ErrHandler:
    clsErrorHandler.LogError "RunRecoverLoan", Err
    MsgBox "Error: " & Err.Description, vbCritical
    Resume Cleanup
End Sub

Public Sub RunLoanReports()
    Dim result As clsResult
    On Error GoTo ErrHandler
    modPerformance.TurnOff
    Set result = svcReportsLoans.WriteLoanReports(repoUser.Current(), modSheetUtils.ControlInputs())
    modMain.Finish result
Cleanup:
    On Error Resume Next
    modPerformance.TurnOn
    Exit Sub
ErrHandler:
    clsErrorHandler.LogError "RunLoanReports", Err
    MsgBox "Error: " & Err.Description, vbCritical
    Resume Cleanup
End Sub
