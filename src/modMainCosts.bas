Attribute VB_Name = "modMainCosts"
Option Explicit

' Stage 2 entry macros: branch costs, bills, expenses, fixed assets, taxes, budgets.
' Same pattern as modMain: current user, error handling, one service call, one message.

Public Sub RunPostCosts()
    Dim result As clsResult
    On Error GoTo ErrHandler
    modPerformance.TurnOff
    Set result = svcCosts.PostCosts(repoUser.Current(), modSheetUtils.ControlInputs())
    modMain.Finish result
Cleanup:
    On Error Resume Next
    modPerformance.TurnOn
    Exit Sub
ErrHandler:
    clsErrorHandler.LogError "RunPostCosts", Err
    MsgBox "Error: " & Err.Description, vbCritical
    Resume Cleanup
End Sub

Public Sub RunPostCostsForMonth()
    Dim result As clsResult
    On Error GoTo ErrHandler
    modPerformance.TurnOff
    Set result = svcCosts.PostCostsForMonth(repoUser.Current(), modSheetUtils.ControlInputs())
    modMain.Finish result
Cleanup:
    On Error Resume Next
    modPerformance.TurnOn
    Exit Sub
ErrHandler:
    clsErrorHandler.LogError "RunPostCostsForMonth", Err
    MsgBox "Error: " & Err.Description, vbCritical
    Resume Cleanup
End Sub

Public Sub RunPayBills()
    Dim result As clsResult
    On Error GoTo ErrHandler
    modPerformance.TurnOff
    Set result = svcCosts.PayBills(repoUser.Current(), modSheetUtils.ControlInputs())
    modMain.Finish result
Cleanup:
    On Error Resume Next
    modPerformance.TurnOn
    Exit Sub
ErrHandler:
    clsErrorHandler.LogError "RunPayBills", Err
    MsgBox "Error: " & Err.Description, vbCritical
    Resume Cleanup
End Sub

Public Sub RunPostExpense()
    Dim result As clsResult
    On Error GoTo ErrHandler
    modPerformance.TurnOff
    Set result = svcExpenses.PostExpense(repoUser.Current(), modSheetUtils.ControlInputs())
    modMain.Finish result
Cleanup:
    On Error Resume Next
    modPerformance.TurnOn
    Exit Sub
ErrHandler:
    clsErrorHandler.LogError "RunPostExpense", Err
    MsgBox "Error: " & Err.Description, vbCritical
    Resume Cleanup
End Sub

Public Sub RunLoadFixedAssets()
    Dim result As clsResult
    On Error GoTo ErrHandler
    modPerformance.TurnOff
    Set result = svcFixedAssets.LoadAssets(repoUser.Current(), modSheetUtils.ControlInputs())
    modMain.Finish result
Cleanup:
    On Error Resume Next
    modPerformance.TurnOn
    Exit Sub
ErrHandler:
    clsErrorHandler.LogError "RunLoadFixedAssets", Err
    MsgBox "Error: " & Err.Description, vbCritical
    Resume Cleanup
End Sub

Public Sub RunBuyAsset()
    Dim result As clsResult
    On Error GoTo ErrHandler
    modPerformance.TurnOff
    Set result = svcFixedAssets.BuyAsset(repoUser.Current(), modSheetUtils.ControlInputs())
    modMain.Finish result
Cleanup:
    On Error Resume Next
    modPerformance.TurnOn
    Exit Sub
ErrHandler:
    clsErrorHandler.LogError "RunBuyAsset", Err
    MsgBox "Error: " & Err.Description, vbCritical
    Resume Cleanup
End Sub

Public Sub RunDisposeAsset()
    Dim result As clsResult
    On Error GoTo ErrHandler
    modPerformance.TurnOff
    Set result = svcFixedAssets.DisposeAsset(repoUser.Current(), modSheetUtils.ControlInputs())
    modMain.Finish result
Cleanup:
    On Error Resume Next
    modPerformance.TurnOn
    Exit Sub
ErrHandler:
    clsErrorHandler.LogError "RunDisposeAsset", Err
    MsgBox "Error: " & Err.Description, vbCritical
    Resume Cleanup
End Sub

Public Sub RunPayTaxes()
    Dim result As clsResult
    On Error GoTo ErrHandler
    modPerformance.TurnOff
    Set result = svcTaxes.PayTaxes(repoUser.Current(), modSheetUtils.ControlInputs())
    modMain.Finish result
Cleanup:
    On Error Resume Next
    modPerformance.TurnOn
    Exit Sub
ErrHandler:
    clsErrorHandler.LogError "RunPayTaxes", Err
    MsgBox "Error: " & Err.Description, vbCritical
    Resume Cleanup
End Sub

Public Sub RunCreateBudget()
    Dim result As clsResult
    On Error GoTo ErrHandler
    modPerformance.TurnOff
    Set result = svcBudgets.CreateBudget(repoUser.Current(), modSheetUtils.ControlInputs())
    modMain.Finish result
Cleanup:
    On Error Resume Next
    modPerformance.TurnOn
    Exit Sub
ErrHandler:
    clsErrorHandler.LogError "RunCreateBudget", Err
    MsgBox "Error: " & Err.Description, vbCritical
    Resume Cleanup
End Sub
