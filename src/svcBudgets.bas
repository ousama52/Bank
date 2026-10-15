Attribute VB_Name = "svcBudgets"
Option Explicit

' Stage 2: a default yearly budget per branch and GL from the active cost rows (typed by hand afterwards).

Private Function Annual(ByVal frequency As String, ByVal amount As Variant) As Variant
    Select Case frequency
        Case "Monthly": Annual = amount * 12
        Case "Quarterly": Annual = amount * 4
        Case Else: Annual = amount
    End Select
End Function

' Control: C25 = year (YYYY). Refuses when the year already has budgets. HeadOffice only.
Public Function CreateBudget(ByVal user As clsUser, ByVal inp As Object) As clsResult
    Dim uow As clsUnitOfWork, msg As String, yr As String, existing As Variant, costs As Variant, assets As Variant, i As Long
    Dim per As Object, k As Variant, uplift As Variant, n As Long, id As String, parts() As String, b As Variant
    msg = svcSecurity.Denied(user, svcSecurity.ACT_BUDGET)
    If Len(msg) > 0 Then Set CreateBudget = modMessages.FailResult(msg): Exit Function
    yr = Left$(Trim$(CStr(inp(modConstants.INP_PERIOD))), 4)
    If Not (yr Like "####") Then Set CreateBudget = modMessages.FailResult("Enter the year (YYYY) in Control C25."): Exit Function
    Set uow = New clsUnitOfWork
    uow.Init repoConfig.BusinessDate(), user.UserID
    existing = uow.WorkingTable(modConstants.TBL_BUDGETS)
    If Not IsEmpty(existing) Then
        For i = 1 To UBound(existing, 1)
            If CStr(existing(i, 3)) = yr Then Set CreateBudget = modMessages.FailResult("Budgets for " & yr & " already exist."): Exit Function
        Next i
    End If
    costs = uow.WorkingTable(modConstants.TBL_BRANCHCOSTS)
    If IsEmpty(costs) Then Set CreateBudget = modMessages.FailResult("tblBranchCosts holds no cost rows."): Exit Function
    uplift = repoConfig.GetDec(modConstants.CFG_BUDGET_UPLIFT)
    Set per = CreateObject("Scripting.Dictionary")
    For i = 1 To UBound(costs, 1)
        If CStr(costs(i, repoCost.BCC_STATUS)) = "Active" And Len(CStr(costs(i, repoCost.BCC_ID))) > 0 Then
            k = CStr(costs(i, repoCost.BCC_BRANCH)) & "|" & CStr(costs(i, repoCost.BCC_GL))
            If per.Exists(k) Then
                per(k) = per(k) + Annual(CStr(costs(i, repoCost.BCC_FREQ)), modMoney.ToDec(costs(i, repoCost.BCC_AMOUNT)))
            Else
                per(k) = Annual(CStr(costs(i, repoCost.BCC_FREQ)), modMoney.ToDec(costs(i, repoCost.BCC_AMOUNT)))
            End If
        End If
    Next i
    For Each k In per.Keys
        per(k) = modMoney.RoundCcy(per(k) * (1 + uplift), modConstants.CCY_MMK)
    Next k
    assets = uow.WorkingTable(modConstants.TBL_FIXEDASSETS)
    For Each b In uow.Branches().Keys
        k = CStr(b) & "|" & modConstants.GL_DEPRECIATION
        If svcFixedAssets.MonthlyChargeOf(assets, CStr(b)) > 0 Then per(k) = svcFixedAssets.MonthlyChargeOf(assets, CStr(b)) * 12
    Next b
    For Each k In per.Keys
        parts = Split(CStr(k), "|")
        id = uow.NewId(modConstants.TBL_BUDGETS, "BudgetID", "BG", 6)
        uow.AppendRow modConstants.TBL_BUDGETS, Array(id, parts(0), CLng(yr), parts(1), per(k))
        n = n + 1
    Next k
    repoAudit.Add uow, "CreateBudget", yr, n & " budget rows"
    uow.Commit
    Set CreateBudget = modMessages.OkResult("Budget " & yr & " created: " & n & " rows (cost + " & Format$(CDbl(uplift) * 100, "0") & "%). You can edit the amounts in tblBudgets.", n)
End Function
