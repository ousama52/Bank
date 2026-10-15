Attribute VB_Name = "svcFixedAssets"
Option Explicit

' Stage 2: fixed asset register in the ledger (load once, monthly depreciation, buy, dispose).

Private Function NeedUpgrade(ByVal data As Variant) As clsResult
    If Not repoAsset.IsUpgraded(data) Then Set NeedUpgrade = modMessages.FailResult("tblFixedAssets lacks the Stage 2 columns. Run RunSetup first.")
End Function

Private Sub AddTo(ByVal d As Object, ByVal key As String, ByVal v As Variant)
    If d.Exists(key) Then d(key) = d(key) + v Else d(key) = v
End Sub

' ---- Opening load (once, idempotent through LoadedOn) ----

' Per branch one journal: Dr 1500 cost, Cr 1510 accumulated depreciation, Cr 3000 net book value.
Public Function LoadAssets(ByVal user As clsUser, ByVal inp As Object) As clsResult
    Dim uow As clsUnitOfWork, msg As String, data As Variant, i As Long, cost As Object, accum As Object, n As Long, b As Variant
    Dim lines As Collection, jr As clsResult, grp As String, r As clsResult, totC As Variant
    msg = svcSecurity.Denied(user, svcSecurity.ACT_LOADASSETS)
    If Len(msg) > 0 Then Set LoadAssets = modMessages.FailResult(msg): Exit Function
    Set uow = New clsUnitOfWork
    uow.Init repoConfig.BusinessDate(), user.UserID
    data = uow.WorkingTable(modConstants.TBL_FIXEDASSETS)
    If IsEmpty(data) Then Set LoadAssets = modMessages.FailResult("tblFixedAssets holds no assets."): Exit Function
    Set r = NeedUpgrade(data)
    If Not r Is Nothing Then Set LoadAssets = r: Exit Function
    Set cost = CreateObject("Scripting.Dictionary")
    Set accum = CreateObject("Scripting.Dictionary")
    For i = 1 To UBound(data, 1)
        If CStr(data(i, repoAsset.FAC_STATUS)) = repoAsset.ASSET_INUSE And Len(CStr(data(i, repoAsset.FAC_LOADED))) = 0 Then
            AddTo cost, CStr(data(i, repoAsset.FAC_BRANCH)), modMoney.ToDec(data(i, repoAsset.FAC_COST))
            AddTo accum, CStr(data(i, repoAsset.FAC_BRANCH)), modMoney.ToDec(data(i, repoAsset.FAC_ACCUM0930))
            data(i, repoAsset.FAC_ACCUM) = modMoney.ToDec(data(i, repoAsset.FAC_ACCUM0930))
            data(i, repoAsset.FAC_LASTPERIOD) = "202609"
            data(i, repoAsset.FAC_LOADED) = uow.BusinessDate
            n = n + 1
        End If
    Next i
    If n = 0 Then Set LoadAssets = modMessages.OkResult("Nothing to load: every asset in use is already loaded.", 0): Exit Function
    totC = CDec(0)
    For Each b In cost.Keys
        Set lines = New Collection
        svcLedger.AddDr lines, CStr(b), modConstants.GL_ASSETS, cost(b)
        svcLedger.AddCr lines, CStr(b), modConstants.GL_ACCUM_DEP, accum(b)
        svcLedger.AddCr lines, CStr(b), modConstants.GL_CAPITAL, cost(b) - accum(b)
        grp = uow.NewGroupId()
        Set jr = svcLedger.PostJournal(uow, modConstants.JT_ASSET_LOAD, "Fixed asset opening load " & CStr(b), grp, lines)
        If Not jr.Success Then Set LoadAssets = jr: Exit Function
        totC = totC + cost(b)
    Next b
    uow.ReplaceTable modConstants.TBL_FIXEDASSETS, data
    repoAudit.Add uow, "LoadFixedAssets", modConstants.TBL_FIXEDASSETS, n & " assets, cost " & modMoney.FormatAmount(totC)
    uow.Commit
    Set LoadAssets = modMessages.OkResult(n & " assets loaded. Cost " & modMoney.FormatAmount(totC) & " MMK.", n)
End Function

' ---- Monthly depreciation ----

' Straight line: RoundCcy((Cost - Residual) / Life) per month, never more than what is left. Per branch: Dr 5250, Cr 1510.
Public Function Depreciate(ByVal uow As clsUnitOfWork, ByVal period As String, ByRef count As Long, ByRef total As Variant) As clsResult
    Dim data As Variant, i As Long, byBranch As Object, depreciable As Variant, monthly As Variant, charge As Variant, accum As Variant
    Dim acq As Date, b As Variant, lines As Collection, jr As clsResult, grp As String, life As Long, r As clsResult
    count = 0
    total = CDec(0)
    data = uow.WorkingTable(modConstants.TBL_FIXEDASSETS)
    If IsEmpty(data) Then Set Depreciate = modMessages.OkResult("", 0): Exit Function
    Set r = NeedUpgrade(data)
    If Not r Is Nothing Then Set Depreciate = r: Exit Function
    Set byBranch = CreateObject("Scripting.Dictionary")
    For i = 1 To UBound(data, 1)
        If CStr(data(i, repoAsset.FAC_STATUS)) = repoAsset.ASSET_INUSE And Len(CStr(data(i, repoAsset.FAC_LOADED))) > 0 Then
            If CStr(data(i, repoAsset.FAC_LASTPERIOD)) < period And modDateUtils.TryToDate(data(i, repoAsset.FAC_ACQUIRED), acq) Then
                If modDateUtils.PeriodOf(acq) < period Then
                    depreciable = modMoney.ToDec(data(i, repoAsset.FAC_COST)) - modMoney.ToDec(data(i, repoAsset.FAC_RESIDUAL))
                    accum = modMoney.ToDec(data(i, repoAsset.FAC_ACCUM))
                    life = CLng(modMoney.ToDec(data(i, repoAsset.FAC_LIFE)))
                    If accum < depreciable And life > 0 Then
                        monthly = modMoney.RoundCcy(depreciable / life, modConstants.CCY_MMK)
                        charge = monthly
                        If charge > depreciable - accum Then charge = depreciable - accum
                        If charge > 0 Then
                            AddTo byBranch, CStr(data(i, repoAsset.FAC_BRANCH)), charge
                            data(i, repoAsset.FAC_ACCUM) = accum + charge
                            data(i, repoAsset.FAC_LASTPERIOD) = period
                            count = count + 1
                            total = total + charge
                        End If
                    End If
                End If
            End If
        End If
    Next i
    If count = 0 Then Set Depreciate = modMessages.OkResult("", 0): Exit Function
    For Each b In byBranch.Keys
        Set lines = New Collection
        svcLedger.AddDr lines, CStr(b), modConstants.GL_DEPRECIATION, byBranch(b)
        svcLedger.AddCr lines, CStr(b), modConstants.GL_ACCUM_DEP, byBranch(b)
        grp = uow.NewGroupId()
        Set jr = svcLedger.PostJournal(uow, modConstants.JT_DEPRECIATION, "Depreciation " & period & " " & CStr(b), grp, lines)
        If Not jr.Success Then Set Depreciate = jr: Exit Function
    Next b
    uow.ReplaceTable modConstants.TBL_FIXEDASSETS, data
    Set Depreciate = modMessages.OkResult("", count)
End Function

' Monthly charge of all assets in use at a branch (budget uses it).
Public Function MonthlyChargeOf(ByVal data As Variant, ByVal branchId As String) As Variant
    Dim i As Long, total As Variant, life As Long
    total = CDec(0)
    If Not IsEmpty(data) Then
        If repoAsset.IsUpgraded(data) Then
            For i = 1 To UBound(data, 1)
                If CStr(data(i, repoAsset.FAC_STATUS)) = repoAsset.ASSET_INUSE And CStr(data(i, repoAsset.FAC_BRANCH)) = branchId Then
                    life = CLng(modMoney.ToDec(data(i, repoAsset.FAC_LIFE)))
                    If life > 0 And modMoney.ToDec(data(i, repoAsset.FAC_ACCUM)) < modMoney.ToDec(data(i, repoAsset.FAC_COST)) - modMoney.ToDec(data(i, repoAsset.FAC_RESIDUAL)) Then
                        total = total + modMoney.RoundCcy((modMoney.ToDec(data(i, repoAsset.FAC_COST)) - modMoney.ToDec(data(i, repoAsset.FAC_RESIDUAL))) / life, modConstants.CCY_MMK)
                    End If
                End If
            Next i
        End If
    End If
    MonthlyChargeOf = total
End Function

' ---- Buying ----

' args: branch, type, desc, amount (cost), life, residual. B: Dr 1500, Cr 1900; HO: Dr 1900, Cr 1100.
Public Function ExecBuy(ByVal uow As clsUnitOfWork, ByVal args As Object, ByVal createdBy As String, ByVal approvedBy As String) As clsResult
    Dim b As String, cost As Variant, life As Long, res As Variant, data As Variant, lines As Collection, jr As clsResult, grp As String
    Dim r As clsResult, row(1 To 16) As Variant, id As String, ho As String, method As String
    ho = modConstants.HEAD_OFFICE_ID
    b = modArgs.Text(args, "branch")
    cost = modArgs.Dec(args, "amount")
    life = CLng(modArgs.Dec(args, "life"))
    res = modArgs.Dec(args, "residual")
    If Not uow.Branches().Exists(b) Then Set ExecBuy = modMessages.FailResult(modMessages.UnknownBranch(b)): Exit Function
    If life <= 0 Then Set ExecBuy = modMessages.FailResult("Useful life in months must be greater than 0 (Control C23)."): Exit Function
    If res < 0 Or res >= cost Then Set ExecBuy = modMessages.FailResult("Residual value must be between 0 and the cost (Control C24)."): Exit Function
    data = uow.WorkingTable(modConstants.TBL_FIXEDASSETS)
    Set r = NeedUpgrade(data)
    If Not r Is Nothing Then Set ExecBuy = r: Exit Function
    Set lines = New Collection
    svcLedger.AddDr lines, b, modConstants.GL_ASSETS, cost
    If b = ho Then
        svcLedger.AddCr lines, ho, modConstants.GL_TREASURY, cost
    Else
        svcLedger.AddCr lines, b, modConstants.GL_INTERBRANCH, cost
        svcLedger.AddDr lines, ho, modConstants.GL_INTERBRANCH, cost
        svcLedger.AddCr lines, ho, modConstants.GL_TREASURY, cost
    End If
    grp = uow.NewGroupId()
    Set jr = svcLedger.PostJournal(uow, modConstants.JT_ASSET_PURCHASE, "Asset purchase " & modArgs.Text(args, "desc"), grp, lines)
    If Not jr.Success Then Set ExecBuy = jr: Exit Function
    id = repoAsset.NextAssetId(data)
    method = "StraightLine"
    If Not IsEmpty(data) Then
        If Len(CStr(data(1, repoAsset.FAC_METHOD))) > 0 Then method = CStr(data(1, repoAsset.FAC_METHOD))
    End If
    row(repoAsset.FAC_ID) = id
    row(repoAsset.FAC_BRANCH) = b
    row(repoAsset.FAC_TYPE) = modArgs.Text(args, "type")
    row(repoAsset.FAC_DESC) = modArgs.Text(args, "desc")
    row(repoAsset.FAC_ACQUIRED) = uow.BusinessDate
    row(repoAsset.FAC_COST) = cost
    row(repoAsset.FAC_LIFE) = life
    row(repoAsset.FAC_RESIDUAL) = res
    row(repoAsset.FAC_METHOD) = method
    row(repoAsset.FAC_ACCUM0930) = CDec(0)
    row(repoAsset.FAC_STATUS) = repoAsset.ASSET_INUSE
    row(repoAsset.FAC_ACCUM) = CDec(0)
    row(repoAsset.FAC_LOADED) = uow.BusinessDate
    uow.AppendRow modConstants.TBL_FIXEDASSETS, row
    repoAudit.Add uow, "BuyAsset", id, modMoney.FormatAmount(cost) & " at " & b & ", approved by " & approvedBy
    Set ExecBuy = modMessages.OkResult("Asset " & id & " bought for " & modMoney.FormatAmount(cost) & " MMK at " & b & ". Group " & grp & ".", 1, grp)
End Function

' Control: C8 branch, C22 type, C21 description, C7 cost, C23 life, C24 residual. HeadOffice requests, another HeadOffice approves.
Public Function BuyAsset(ByVal user As clsUser, ByVal inp As Object) As clsResult
    Dim uow As clsUnitOfWork, msg As String, cost As Variant, args As Object, b As String, r As clsResult, life As Variant, res As Variant
    msg = svcSecurity.Denied(user, svcSecurity.ACT_BUYASSET)
    If Len(msg) > 0 Then Set BuyAsset = modMessages.FailResult(msg): Exit Function
    If Not modArgs.ParseAmount(inp(modConstants.INP_AMOUNT), cost, msg) Then Set BuyAsset = modMessages.FailResult(msg): Exit Function
    Set uow = New clsUnitOfWork
    uow.Init repoConfig.BusinessDate(), user.UserID
    If Not svcSecurity.ResolveBranch(uow, user, CStr(inp(modConstants.INP_BRANCH)), b, msg) Then Set BuyAsset = modMessages.FailResult(msg): Exit Function
    If Not modMoney.TryToDec(inp(modConstants.INP_LIFE), life) Then Set BuyAsset = modMessages.FailResult("Enter the useful life in months in Control C23."): Exit Function
    If Not modMoney.TryToDec(inp(modConstants.INP_RESIDUAL), res) Then res = CDec(0)
    Set args = modArgs.NewArgs()
    args("branch") = b
    args("type") = CStr(inp(modConstants.INP_ASSETTYPE))
    args("desc") = CStr(inp(modConstants.INP_DESC))
    args("amount") = cost
    args("life") = life
    args("residual") = res
    Set r = svcApproval.Request(uow, modConstants.REQ_BUY_ASSET, "", "", cost, b, args, user.UserID)
    uow.Commit
    Set BuyAsset = r
End Function

' ---- Disposal ----

' args: asset (ID), amount (proceeds, may be 0). B: Dr 1510 accumulated, Dr 1900 proceeds, Cr 1500 cost, gain Cr 4400 or loss Dr 5310.
' HO: Dr 1100 proceeds, Cr 1900 proceeds.
Public Function ExecDispose(ByVal uow As clsUnitOfWork, ByVal args As Object, ByVal createdBy As String, ByVal approvedBy As String) As clsResult
    Dim data As Variant, i As Long, row As Long, b As String, cost As Variant, accum As Variant, book As Variant, proceeds As Variant, gain As Variant
    Dim lines As Collection, jr As clsResult, grp As String, ho As String, r As clsResult
    ho = modConstants.HEAD_OFFICE_ID
    data = uow.WorkingTable(modConstants.TBL_FIXEDASSETS)
    Set r = NeedUpgrade(data)
    If Not r Is Nothing Then Set ExecDispose = r: Exit Function
    If Not IsEmpty(data) Then
        For i = 1 To UBound(data, 1)
            If StrComp(CStr(data(i, repoAsset.FAC_ID)), modArgs.Text(args, "asset"), vbTextCompare) = 0 Then row = i
        Next i
    End If
    If row = 0 Then Set ExecDispose = modMessages.FailResult("Unknown asset " & modArgs.Text(args, "asset") & " (Control C12)."): Exit Function
    If CStr(data(row, repoAsset.FAC_STATUS)) <> repoAsset.ASSET_INUSE Or Len(CStr(data(row, repoAsset.FAC_LOADED))) = 0 Then
        Set ExecDispose = modMessages.FailResult("Asset " & CStr(data(row, repoAsset.FAC_ID)) & " is not an asset in use (or not loaded yet).")
        Exit Function
    End If
    b = CStr(data(row, repoAsset.FAC_BRANCH))
    cost = modMoney.ToDec(data(row, repoAsset.FAC_COST))
    accum = modMoney.ToDec(data(row, repoAsset.FAC_ACCUM))
    proceeds = modArgs.Dec(args, "amount")
    book = cost - accum
    gain = proceeds - book
    Set lines = New Collection
    svcLedger.AddDr lines, b, modConstants.GL_ACCUM_DEP, accum
    svcLedger.AddCr lines, b, modConstants.GL_ASSETS, cost
    If gain > 0 Then svcLedger.AddCr lines, b, modConstants.GL_GAIN_DISPOSAL, gain
    If gain < 0 Then svcLedger.AddDr lines, b, modConstants.GL_LOSS_DISPOSAL, -gain
    If b = ho Then
        svcLedger.AddDr lines, ho, modConstants.GL_TREASURY, proceeds
    Else
        svcLedger.AddDr lines, b, modConstants.GL_INTERBRANCH, proceeds
        svcLedger.AddDr lines, ho, modConstants.GL_TREASURY, proceeds
        svcLedger.AddCr lines, ho, modConstants.GL_INTERBRANCH, proceeds
    End If
    grp = uow.NewGroupId()
    Set jr = svcLedger.PostJournal(uow, modConstants.JT_ASSET_DISPOSAL, "Asset disposal " & CStr(data(row, repoAsset.FAC_ID)), grp, lines)
    If Not jr.Success Then Set ExecDispose = jr: Exit Function
    data(row, repoAsset.FAC_STATUS) = repoAsset.ASSET_SOLD
    data(row, repoAsset.FAC_DISPOSED) = uow.BusinessDate
    data(row, repoAsset.FAC_PROCEEDS) = proceeds
    uow.ReplaceTable modConstants.TBL_FIXEDASSETS, data
    repoAudit.Add uow, "DisposeAsset", CStr(data(row, repoAsset.FAC_ID)), "Proceeds " & modMoney.FormatAmount(proceeds) & ", book value " & modMoney.FormatAmount(book) & ", approved by " & approvedBy
    Set ExecDispose = modMessages.OkResult("Asset " & CStr(data(row, repoAsset.FAC_ID)) & " sold for " & modMoney.FormatAmount(proceeds) & " MMK (book value " & _
        modMoney.FormatAmount(book) & ", " & IIf(gain >= 0, "gain ", "loss ") & modMoney.FormatAmount(Abs(gain)) & "). Group " & grp & ".", 1, grp)
End Function

' Control: C12 asset ID, C7 proceeds (0 allowed). HeadOffice requests, another HeadOffice approves.
Public Function DisposeAsset(ByVal user As clsUser, ByVal inp As Object) As clsResult
    Dim uow As clsUnitOfWork, msg As String, proceeds As Variant, args As Object, r As clsResult
    msg = svcSecurity.Denied(user, svcSecurity.ACT_DISPOSE)
    If Len(msg) > 0 Then Set DisposeAsset = modMessages.FailResult(msg): Exit Function
    If Not modArgs.ParseAmount(inp(modConstants.INP_AMOUNT), proceeds, msg, True) Then Set DisposeAsset = modMessages.FailResult(msg): Exit Function
    If Len(CStr(inp(modConstants.INP_REFID))) = 0 Then Set DisposeAsset = modMessages.FailResult("Put the asset ID in Control C12."): Exit Function
    Set uow = New clsUnitOfWork
    uow.Init repoConfig.BusinessDate(), user.UserID
    Set args = modArgs.NewArgs()
    args("asset") = CStr(inp(modConstants.INP_REFID))
    args("amount") = proceeds
    Set r = svcApproval.Request(uow, modConstants.REQ_DISPOSE_ASSET, "", "", proceeds, modConstants.HEAD_OFFICE_ID, args, user.UserID)
    uow.Commit
    Set DisposeAsset = r
End Function
