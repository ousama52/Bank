Attribute VB_Name = "modMain"
Option Explicit

' Entry macros (Control sheet buttons). They only orchestrate: current user, error handling, one service call, one message.
' Services check the role themselves; nothing below this module shows a MsgBox.

' Shows the result in Control C18 and a message box, and refreshes the business date in C2.
Public Sub Finish(ByVal result As clsResult)
    modSheetUtils.SetStatus result.Message
    modSheetUtils.ShowBusinessDate repoConfig.BusinessDate()
    If result.Success Then
        MsgBox result.Message, vbInformation, "Golden Delta Bank (SIMULATION)"
    Else
        MsgBox result.Message, vbExclamation, "Golden Delta Bank (SIMULATION)"
    End If
End Sub

' Setup works on a fresh workbook (no Control sheet yet). Once a user is chosen, only an Admin may run it.
Public Sub RunSetup()
    Dim result As clsResult, u As clsUser
    On Error GoTo ErrHandler
    modPerformance.TurnOff
    If modSheetUtils.SheetExists(modConstants.SHEET_CONTROL) And modSheetUtils.TableExists(modConstants.TBL_USERS) Then
        Set u = repoUser.Current()
        If Not u Is Nothing Then
            If u.Role <> modConstants.ROLE_ADMIN Then
                MsgBox modMessages.RoleDenied(u.Role, svcSecurity.ACT_SETUP), vbExclamation
                GoTo Cleanup
            End If
        End If
    End If
    Set result = modSetup.SetupAll()
    Finish result
Cleanup:
    On Error Resume Next
    modPerformance.TurnOn
    Exit Sub
ErrHandler:
    clsErrorHandler.LogError "RunSetup", Err
    MsgBox "Error: " & Err.Description, vbCritical
    Resume Cleanup
End Sub

Public Sub RunOpenAccount()
    Dim result As clsResult
    On Error GoTo ErrHandler
    modPerformance.TurnOff
    Set result = svcAccounts.OpenAccount(repoUser.Current(), modSheetUtils.ControlInputs())
    Finish result
Cleanup:
    On Error Resume Next
    modPerformance.TurnOn
    Exit Sub
ErrHandler:
    clsErrorHandler.LogError "RunOpenAccount", Err
    MsgBox "Error: " & Err.Description, vbCritical
    Resume Cleanup
End Sub

Public Sub RunOpenFixedDeposit()
    Dim result As clsResult
    On Error GoTo ErrHandler
    modPerformance.TurnOff
    Set result = svcFixedDeposit.OpenFixedDeposit(repoUser.Current(), modSheetUtils.ControlInputs())
    Finish result
Cleanup:
    On Error Resume Next
    modPerformance.TurnOn
    Exit Sub
ErrHandler:
    clsErrorHandler.LogError "RunOpenFixedDeposit", Err
    MsgBox "Error: " & Err.Description, vbCritical
    Resume Cleanup
End Sub

Public Sub RunDeposit()
    Dim result As clsResult
    On Error GoTo ErrHandler
    modPerformance.TurnOff
    Set result = svcCashOps.Deposit(repoUser.Current(), modSheetUtils.ControlInputs())
    Finish result
Cleanup:
    On Error Resume Next
    modPerformance.TurnOn
    Exit Sub
ErrHandler:
    clsErrorHandler.LogError "RunDeposit", Err
    MsgBox "Error: " & Err.Description, vbCritical
    Resume Cleanup
End Sub

Public Sub RunWithdraw()
    Dim result As clsResult
    On Error GoTo ErrHandler
    modPerformance.TurnOff
    Set result = svcCashOps.Withdraw(repoUser.Current(), modSheetUtils.ControlInputs())
    Finish result
Cleanup:
    On Error Resume Next
    modPerformance.TurnOn
    Exit Sub
ErrHandler:
    clsErrorHandler.LogError "RunWithdraw", Err
    MsgBox "Error: " & Err.Description, vbCritical
    Resume Cleanup
End Sub

Public Sub RunTransfer()
    Dim result As clsResult
    On Error GoTo ErrHandler
    modPerformance.TurnOff
    Set result = svcTransfers.Transfer(repoUser.Current(), modSheetUtils.ControlInputs())
    Finish result
Cleanup:
    On Error Resume Next
    modPerformance.TurnOn
    Exit Sub
ErrHandler:
    clsErrorHandler.LogError "RunTransfer", Err
    MsgBox "Error: " & Err.Description, vbCritical
    Resume Cleanup
End Sub

Public Sub RunInterbankOut()
    Dim result As clsResult
    On Error GoTo ErrHandler
    modPerformance.TurnOff
    Set result = svcInterbank.SendOut(repoUser.Current(), modSheetUtils.ControlInputs())
    Finish result
Cleanup:
    On Error Resume Next
    modPerformance.TurnOn
    Exit Sub
ErrHandler:
    clsErrorHandler.LogError "RunInterbankOut", Err
    MsgBox "Error: " & Err.Description, vbCritical
    Resume Cleanup
End Sub

Public Sub RunInterbankIn()
    Dim result As clsResult
    On Error GoTo ErrHandler
    modPerformance.TurnOff
    Set result = svcInterbank.ReceiveIn(repoUser.Current(), modSheetUtils.ControlInputs())
    Finish result
Cleanup:
    On Error Resume Next
    modPerformance.TurnOn
    Exit Sub
ErrHandler:
    clsErrorHandler.LogError "RunInterbankIn", Err
    MsgBox "Error: " & Err.Description, vbCritical
    Resume Cleanup
End Sub

Public Sub RunReturnInterbank()
    Dim result As clsResult
    On Error GoTo ErrHandler
    modPerformance.TurnOff
    Set result = svcInterbank.ReturnTransfer(repoUser.Current(), modSheetUtils.ControlInputs())
    Finish result
Cleanup:
    On Error Resume Next
    modPerformance.TurnOn
    Exit Sub
ErrHandler:
    clsErrorHandler.LogError "RunReturnInterbank", Err
    MsgBox "Error: " & Err.Description, vbCritical
    Resume Cleanup
End Sub

Public Sub RunCashShipment()
    Dim result As clsResult
    On Error GoTo ErrHandler
    modPerformance.TurnOff
    Set result = svcCashOps.Shipment(repoUser.Current(), modSheetUtils.ControlInputs())
    Finish result
Cleanup:
    On Error Resume Next
    modPerformance.TurnOn
    Exit Sub
ErrHandler:
    clsErrorHandler.LogError "RunCashShipment", Err
    MsgBox "Error: " & Err.Description, vbCritical
    Resume Cleanup
End Sub

Public Sub RunApprove()
    Dim result As clsResult
    On Error GoTo ErrHandler
    modPerformance.TurnOff
    Set result = svcApproval.Approve(repoUser.Current(), modSheetUtils.ControlInputs())
    Finish result
Cleanup:
    On Error Resume Next
    modPerformance.TurnOn
    Exit Sub
ErrHandler:
    clsErrorHandler.LogError "RunApprove", Err
    MsgBox "Error: " & Err.Description, vbCritical
    Resume Cleanup
End Sub

Public Sub RunReject()
    Dim result As clsResult
    On Error GoTo ErrHandler
    modPerformance.TurnOff
    Set result = svcApproval.Reject(repoUser.Current(), modSheetUtils.ControlInputs())
    Finish result
Cleanup:
    On Error Resume Next
    modPerformance.TurnOn
    Exit Sub
ErrHandler:
    clsErrorHandler.LogError "RunReject", Err
    MsgBox "Error: " & Err.Description, vbCritical
    Resume Cleanup
End Sub

Public Sub RunReverse()
    Dim result As clsResult
    On Error GoTo ErrHandler
    modPerformance.TurnOff
    Set result = svcApproval.RequestReversal(repoUser.Current(), modSheetUtils.ControlInputs())
    Finish result
Cleanup:
    On Error Resume Next
    modPerformance.TurnOn
    Exit Sub
ErrHandler:
    clsErrorHandler.LogError "RunReverse", Err
    MsgBox "Error: " & Err.Description, vbCritical
    Resume Cleanup
End Sub

Public Sub RunFreeze()
    Dim result As clsResult
    On Error GoTo ErrHandler
    modPerformance.TurnOff
    Set result = svcAccounts.Freeze(repoUser.Current(), modSheetUtils.ControlInputs())
    Finish result
Cleanup:
    On Error Resume Next
    modPerformance.TurnOn
    Exit Sub
ErrHandler:
    clsErrorHandler.LogError "RunFreeze", Err
    MsgBox "Error: " & Err.Description, vbCritical
    Resume Cleanup
End Sub

Public Sub RunUnfreeze()
    Dim result As clsResult
    On Error GoTo ErrHandler
    modPerformance.TurnOff
    Set result = svcAccounts.Unfreeze(repoUser.Current(), modSheetUtils.ControlInputs())
    Finish result
Cleanup:
    On Error Resume Next
    modPerformance.TurnOn
    Exit Sub
ErrHandler:
    clsErrorHandler.LogError "RunUnfreeze", Err
    MsgBox "Error: " & Err.Description, vbCritical
    Resume Cleanup
End Sub

Public Sub RunReactivate()
    Dim result As clsResult
    On Error GoTo ErrHandler
    modPerformance.TurnOff
    Set result = svcAccounts.Reactivate(repoUser.Current(), modSheetUtils.ControlInputs())
    Finish result
Cleanup:
    On Error Resume Next
    modPerformance.TurnOn
    Exit Sub
ErrHandler:
    clsErrorHandler.LogError "RunReactivate", Err
    MsgBox "Error: " & Err.Description, vbCritical
    Resume Cleanup
End Sub

Public Sub RunCloseAccount()
    Dim result As clsResult
    On Error GoTo ErrHandler
    modPerformance.TurnOff
    Set result = svcAccounts.CloseAccount(repoUser.Current(), modSheetUtils.ControlInputs())
    Finish result
Cleanup:
    On Error Resume Next
    modPerformance.TurnOn
    Exit Sub
ErrHandler:
    clsErrorHandler.LogError "RunCloseAccount", Err
    MsgBox "Error: " & Err.Description, vbCritical
    Resume Cleanup
End Sub

Public Sub RunBreakFixedDeposit()
    Dim result As clsResult
    On Error GoTo ErrHandler
    modPerformance.TurnOff
    Set result = svcFixedDeposit.BreakFixedDeposit(repoUser.Current(), modSheetUtils.ControlInputs())
    Finish result
Cleanup:
    On Error Resume Next
    modPerformance.TurnOn
    Exit Sub
ErrHandler:
    clsErrorHandler.LogError "RunBreakFixedDeposit", Err
    MsgBox "Error: " & Err.Description, vbCritical
    Resume Cleanup
End Sub

Public Sub RunEndOfDay()
    Dim result As clsResult
    On Error GoTo ErrHandler
    modPerformance.TurnOff
    Set result = svcEndOfDay.CloseDay(repoUser.Current())
    Finish result
Cleanup:
    On Error Resume Next
    modPerformance.TurnOn
    Exit Sub
ErrHandler:
    clsErrorHandler.LogError "RunEndOfDay", Err
    MsgBox "Error: " & Err.Description, vbCritical
    Resume Cleanup
End Sub

Public Sub RunReports()
    Dim result As clsResult
    On Error GoTo ErrHandler
    modPerformance.TurnOff
    Set result = svcReports.RunAll(repoUser.Current(), modSheetUtils.ControlInputs())
    Finish result
Cleanup:
    On Error Resume Next
    modPerformance.TurnOn
    Exit Sub
ErrHandler:
    clsErrorHandler.LogError "RunReports", Err
    MsgBox "Error: " & Err.Description, vbCritical
    Resume Cleanup
End Sub

Public Sub RunStatement()
    Dim result As clsResult
    On Error GoTo ErrHandler
    modPerformance.TurnOff
    Set result = svcReports.Statement(repoUser.Current(), modSheetUtils.ControlInputs())
    Finish result
Cleanup:
    On Error Resume Next
    modPerformance.TurnOn
    Exit Sub
ErrHandler:
    clsErrorHandler.LogError "RunStatement", Err
    MsgBox "Error: " & Err.Description, vbCritical
    Resume Cleanup
End Sub
