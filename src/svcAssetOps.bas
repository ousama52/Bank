Attribute VB_Name = "svcAssetOps"
Option Explicit

' Stage 4: running a repossessed asset (rent it out, use it in a business, record income and costs) and the month-end holding costs.

Private Function NeedTables() As String
    If Not modSheetUtils.TableExists(modConstants.TBL_REPOSSESSIONS) Or Not modSheetUtils.TableExists(modConstants.TBL_ASSETOPS) Or Not modSheetUtils.TableExists(modConstants.TBL_ASSETRENTALS) Then
        NeedTables = "The held-asset tables are missing. Run RunSetup first."
    End If
End Function

' Finds the repo row for Control C28 and checks that this user may handle its branch.
Private Function FindHeld(ByVal uow As clsUnitOfWork, ByVal user As clsUser, ByVal repoId As String, ByRef repos As Variant, ByRef rr As Long, ByRef msg As String) As Boolean
    Dim st As String
    repos = uow.WorkingTable(modConstants.TBL_REPOSSESSIONS)
    rr = repoCollateral.FindCollateralRow(repos, repoCollateral.RPC_ID, repoId)
    If rr = 0 Then msg = "Unknown repossession record. Put the RepoID in Control C28.": Exit Function
    st = CStr(repos(rr, repoCollateral.RPC_STATUS))
    If st <> repoCollateral.RS_HELD And st <> repoCollateral.RS_RENTED And st <> repoCollateral.RS_INBUSINESS Then msg = "Repossession " & repoId & " is " & st & ", not a held asset.": Exit Function
    If user.Role <> modConstants.ROLE_HEADOFFICE And CStr(repos(rr, repoCollateral.RPC_BRANCH)) <> user.BranchID Then
        msg = "This asset is held at branch " & CStr(repos(rr, repoCollateral.RPC_BRANCH)) & "."
        Exit Function
    End If
    FindHeld = True
End Function

' Control: C28 repo, C20 tenant, C7 monthly rent. One active rental per asset.
Public Function StartRental(ByVal user As clsUser, ByVal inp As Object) As clsResult
    Dim uow As clsUnitOfWork, msg As String, repos As Variant, rr As Long, rent As Variant, rentals As Variant, i As Long, id As String
    msg = svcSecurity.Denied(user, svcSecurity.ACT_ASSETOPS)
    If Len(msg) > 0 Then Set StartRental = modMessages.FailResult(msg): Exit Function
    msg = NeedTables()
    If Len(msg) > 0 Then Set StartRental = modMessages.FailResult(msg): Exit Function
    If Not modArgs.ParseAmount(inp(modConstants.INP_AMOUNT), rent, msg) Then Set StartRental = modMessages.FailResult(msg): Exit Function
    Set uow = New clsUnitOfWork
    uow.Init repoConfig.BusinessDate(), user.UserID
    If Not FindHeld(uow, user, CStr(inp(modConstants.INP_LOANID)), repos, rr, msg) Then Set StartRental = modMessages.FailResult(msg): Exit Function
    rentals = uow.WorkingTable(modConstants.TBL_ASSETRENTALS)
    If Not IsEmpty(rentals) Then
        For i = 1 To UBound(rentals, 1)
            If CStr(rentals(i, 2)) = CStr(repos(rr, repoCollateral.RPC_ID)) And CStr(rentals(i, 7)) = "Active" Then Set StartRental = modMessages.FailResult("This asset already has an active rental."): Exit Function
        Next i
    End If
    id = uow.NewId(modConstants.TBL_ASSETRENTALS, "RentalID", "RN", 6)
    uow.AppendRow modConstants.TBL_ASSETRENTALS, Array(id, CStr(repos(rr, repoCollateral.RPC_ID)), CStr(inp(modConstants.INP_VENDOR)), rent, uow.BusinessDate, Empty, "Active")
    repos(rr, repoCollateral.RPC_STATUS) = repoCollateral.RS_RENTED
    uow.ReplaceTable modConstants.TBL_REPOSSESSIONS, repos
    repoAudit.Add uow, "StartRental", id, "Repo " & CStr(repos(rr, repoCollateral.RPC_ID)) & " rent " & modMoney.FormatAmount(rent)
    uow.Commit
    Set StartRental = modMessages.OkResult("Rental " & id & " started at " & modMoney.FormatAmount(rent) & " MMK a month. Record each rent with RunRecordAssetOp (note Rent).", 1, id)
End Function

Public Function EndRental(ByVal user As clsUser, ByVal inp As Object) As clsResult
    Dim uow As clsUnitOfWork, msg As String, repos As Variant, rr As Long, rentals As Variant, i As Long, found As Boolean
    msg = svcSecurity.Denied(user, svcSecurity.ACT_ASSETOPS)
    If Len(msg) > 0 Then Set EndRental = modMessages.FailResult(msg): Exit Function
    msg = NeedTables()
    If Len(msg) > 0 Then Set EndRental = modMessages.FailResult(msg): Exit Function
    Set uow = New clsUnitOfWork
    uow.Init repoConfig.BusinessDate(), user.UserID
    If Not FindHeld(uow, user, CStr(inp(modConstants.INP_LOANID)), repos, rr, msg) Then Set EndRental = modMessages.FailResult(msg): Exit Function
    rentals = uow.WorkingTable(modConstants.TBL_ASSETRENTALS)
    If Not IsEmpty(rentals) Then
        For i = 1 To UBound(rentals, 1)
            If CStr(rentals(i, 2)) = CStr(repos(rr, repoCollateral.RPC_ID)) And CStr(rentals(i, 7)) = "Active" Then
                rentals(i, 6) = uow.BusinessDate
                rentals(i, 7) = "Ended"
                found = True
            End If
        Next i
    End If
    If Not found Then Set EndRental = modMessages.FailResult("This asset has no active rental."): Exit Function
    uow.ReplaceTable modConstants.TBL_ASSETRENTALS, rentals
    repos(rr, repoCollateral.RPC_STATUS) = repoCollateral.RS_HELD
    uow.ReplaceTable modConstants.TBL_REPOSSESSIONS, repos
    uow.Commit
    Set EndRental = modMessages.OkResult("Rental ended. The asset is Held again.", 1)
End Function

Public Function UseInBusiness(ByVal user As clsUser, ByVal inp As Object) As clsResult
    Dim uow As clsUnitOfWork, msg As String, repos As Variant, rr As Long
    msg = svcSecurity.Denied(user, svcSecurity.ACT_ASSETOPS)
    If Len(msg) > 0 Then Set UseInBusiness = modMessages.FailResult(msg): Exit Function
    msg = NeedTables()
    If Len(msg) > 0 Then Set UseInBusiness = modMessages.FailResult(msg): Exit Function
    Set uow = New clsUnitOfWork
    uow.Init repoConfig.BusinessDate(), user.UserID
    If Not FindHeld(uow, user, CStr(inp(modConstants.INP_LOANID)), repos, rr, msg) Then Set UseInBusiness = modMessages.FailResult(msg): Exit Function
    If CStr(repos(rr, repoCollateral.RPC_STATUS)) <> repoCollateral.RS_HELD Then Set UseInBusiness = modMessages.FailResult("Only a Held asset can be put to business use."): Exit Function
    repos(rr, repoCollateral.RPC_STATUS) = repoCollateral.RS_INBUSINESS
    uow.ReplaceTable modConstants.TBL_REPOSSESSIONS, repos
    repoAudit.Add uow, "UseInBusiness", CStr(repos(rr, repoCollateral.RPC_ID)), "Asset put into business use"
    uow.Commit
    Set UseInBusiness = modMessages.OkResult("The asset is now InBusiness. Record its revenue and costs with RunRecordAssetOp.", 1)
End Function

' Control: C28 repo, C15 type (Rent, Revenue or Cost), C7 amount, C27 mode (Cash or Account), C5 tenant account (Account mode), C8 branch (Cash mode).
' Rent and Revenue: Dr cash or the account, Cr 4500. Cost: Dr 5220, Cr 1000 (Cash) or through head office.
Public Function RecordOp(ByVal user As clsUser, ByVal inp As Object) As clsResult
    Dim uow As clsUnitOfWork, msg As String, repos As Variant, rr As Long, amt As Variant, opType As String, mode As String, p As String, h As String, lines As Collection
    Dim jr As clsResult, grp As String, a As clsAccount, id As String, ho As String, cash As Boolean
    msg = svcSecurity.Denied(user, svcSecurity.ACT_ASSETOPS)
    If Len(msg) > 0 Then Set RecordOp = modMessages.FailResult(msg): Exit Function
    msg = NeedTables()
    If Len(msg) > 0 Then Set RecordOp = modMessages.FailResult(msg): Exit Function
    If Not modArgs.ParseAmount(inp(modConstants.INP_AMOUNT), amt, msg) Then Set RecordOp = modMessages.FailResult(msg): Exit Function
    opType = StrConv(Trim$(CStr(inp(modConstants.INP_NOTE))), vbProperCase)
    If opType <> "Rent" And opType <> "Revenue" And opType <> "Cost" Then Set RecordOp = modMessages.FailResult("Type Rent, Revenue or Cost in Control C15."): Exit Function
    Set uow = New clsUnitOfWork
    uow.Init repoConfig.BusinessDate(), user.UserID
    If Not FindHeld(uow, user, CStr(inp(modConstants.INP_LOANID)), repos, rr, msg) Then Set RecordOp = modMessages.FailResult(msg): Exit Function
    If Not svcSecurity.ResolveBranch(uow, user, CStr(inp(modConstants.INP_BRANCH)), p, msg) Then Set RecordOp = modMessages.FailResult(msg): Exit Function
    ho = modConstants.HEAD_OFFICE_ID
    h = CStr(repos(rr, repoCollateral.RPC_BRANCH))
    mode = UCase$(CStr(inp(modConstants.INP_MODE)))
    cash = (mode = "CASH")
    Set lines = New Collection
    If opType = "Cost" Then
        svcLedger.AddDr lines, h, modConstants.GL_MAINTENANCE, amt
        If cash Then
            If uow.GLNet(h, modConstants.GL_VAULT) - amt < 0 Then Set RecordOp = modMessages.FailResult(modMessages.MSG_VAULT_SHORT): Exit Function
            svcLedger.AddCr lines, h, modConstants.GL_VAULT, amt
        ElseIf h = ho Then
            svcLedger.AddCr lines, ho, modConstants.GL_TREASURY, amt
        Else
            svcLedger.AddCr lines, h, modConstants.GL_INTERBRANCH, amt
            svcLedger.AddDr lines, ho, modConstants.GL_INTERBRANCH, amt
            svcLedger.AddCr lines, ho, modConstants.GL_TREASURY, amt
        End If
    Else
        If cash Then
            svcLedger.AddDr lines, p, modConstants.GL_VAULT, amt
            If p <> h Then
                svcLedger.AddCr lines, p, modConstants.GL_INTERBRANCH, amt
                svcLedger.AddDr lines, h, modConstants.GL_INTERBRANCH, amt
            End If
        Else
            Set a = uow.GetAccount(CStr(inp(modConstants.INP_ACCOUNT)))
            If a Is Nothing Then Set RecordOp = modMessages.FailResult("Account mode needs the tenant account in Control C5."): Exit Function
            If Not svcAccounts.CanDebit(a, amt, False, msg) Then Set RecordOp = modMessages.FailResult(msg): Exit Function
            svcLedger.AddDr lines, a.BranchID, a.DepositGL, amt, a.AccountNo
            If a.BranchID <> h Then
                svcLedger.AddCr lines, a.BranchID, modConstants.GL_INTERBRANCH, amt
                svcLedger.AddDr lines, h, modConstants.GL_INTERBRANCH, amt
            End If
        End If
        svcLedger.AddCr lines, h, modConstants.GL_RENT_INCOME, amt
    End If
    grp = uow.NewGroupId()
    Set jr = svcLedger.PostJournal(uow, modConstants.JT_ASSET_OP, opType & " " & CStr(repos(rr, repoCollateral.RPC_ID)), grp, lines)
    If Not jr.Success Then Set RecordOp = jr: Exit Function
    If Not a Is Nothing Then svcAccounts.Effect uow, a, modConstants.TXN_WITHDRAWAL, modConstants.DIR_DEBIT, amt, "", p, opType & " for asset " & CStr(repos(rr, repoCollateral.RPC_ID)), CStr(repos(rr, repoCollateral.RPC_ID)), jr.RefID, grp, user.UserID, "", True
    id = uow.NewId(modConstants.TBL_ASSETOPS, "OpID", "OP", 6)
    uow.AppendRow modConstants.TBL_ASSETOPS, repoCollateral.OpRow(id, CStr(repos(rr, repoCollateral.RPC_ID)), uow.BusinessDate, IIf(opType = "Cost", "Cost", opType), amt, "", jr.RefID)
    uow.Commit
    Set RecordOp = modMessages.OkResult(opType & " of " & modMoney.FormatAmount(amt) & " MMK recorded for asset " & CStr(repos(rr, repoCollateral.RPC_ID)) & ". Group " & grp & ".", 1, grp)
End Function

' ---- Month end: holding costs, "must sell" and re-appraisal lists ----

' Holding cost = CarryingValue x MonthlyHoldingCostPct, posted like a Stage 2 bill: B: Dr 5220, Cr 2300, and a Posted expense that
' head office pays with RunPayBills. Held assets older than MaxHoldingMonths and assets due for re-appraisal are written to the Log.
Public Function MonthEnd(ByVal uow As clsUnitOfWork, ByVal d As Date, ByVal stats As Object) As clsResult
    Dim repos As Variant, items As Variant, i As Long, ir As Long, trow As Variant, cost As Variant, lines As Collection, jr As clsResult, grp As String, id As String
    Dim h As String, rd As Date, ad As Date
    If Not modSheetUtils.TableExists(modConstants.TBL_REPOSSESSIONS) Or Not modSheetUtils.TableExists(modConstants.TBL_EXPENSES) Then Set MonthEnd = modMessages.OkResult("", 0): Exit Function
    repos = uow.WorkingTable(modConstants.TBL_REPOSSESSIONS)
    If IsEmpty(repos) Then Set MonthEnd = modMessages.OkResult("", 0): Exit Function
    items = uow.WorkingTable(modConstants.TBL_COLLATERAL)
    For i = 1 To UBound(repos, 1)
        If CStr(repos(i, repoCollateral.RPC_STATUS)) = repoCollateral.RS_HELD Or CStr(repos(i, repoCollateral.RPC_STATUS)) = repoCollateral.RS_RENTED Or CStr(repos(i, repoCollateral.RPC_STATUS)) = repoCollateral.RS_INBUSINESS Then
            ir = repoCollateral.FindCollateralRow(items, repoCollateral.COC_ID, CStr(repos(i, repoCollateral.RPC_COLLATERAL)))
            trow = repoCollateral.TypeRow(CStr(items(ir, repoCollateral.COC_TYPE)))
            h = CStr(repos(i, repoCollateral.RPC_BRANCH))
            If Not IsEmpty(trow) Then
                cost = modMoney.RoundCcy(modMoney.ToDec(repos(i, repoCollateral.RPC_CARRYING)) * modMoney.ToDec(trow(repoCollateral.CTC_HOLDCOST)), modConstants.CCY_MMK)
                If cost > 0 Then
                    Set lines = New Collection
                    svcLedger.AddDr lines, h, modConstants.GL_MAINTENANCE, cost
                    svcLedger.AddCr lines, h, modConstants.GL_ACCRUED, cost
                    grp = uow.NewGroupId()
                    Set jr = svcLedger.PostJournal(uow, modConstants.JT_EXPENSE, "Holding cost " & CStr(repos(i, repoCollateral.RPC_ID)), grp, lines)
                    If Not jr.Success Then Set MonthEnd = jr: Exit Function
                    id = uow.NewId(modConstants.TBL_EXPENSES, "ExpenseID", "EX", 6)
                    uow.AppendRow modConstants.TBL_EXPENSES, repoCost.ExpenseRow(id, h, d, modConstants.GL_MAINTENANCE, cost, "Holding cost", "Repossessed asset " & CStr(repos(i, repoCollateral.RPC_ID)), _
                        "HeadOffice", "Posted", "SYSTEM", "", jr.RefID, "")
                    uow.AppendRow modConstants.TBL_ASSETOPS, repoCollateral.OpRow(uow.NewId(modConstants.TBL_ASSETOPS, "OpID", "OP", 6), CStr(repos(i, repoCollateral.RPC_ID)), d, "Cost", cost, "Holding cost", jr.RefID)
                End If
            End If
            If modDateUtils.TryToDate(repos(i, repoCollateral.RPC_REPODATE), rd) Then
                If DateDiff("m", rd, d) >= repoConfig.GetLong(modConstants.CFG_MAX_HOLDING) Then clsLogger.Warn "svcAssetOps.MonthEnd", "Must sell: " & CStr(repos(i, repoCollateral.RPC_ID)) & " held since " & modDateUtils.ToIso(rd)
            End If
            If modDateUtils.TryToDate(items(ir, repoCollateral.COC_VALUEDATE), ad) Then
                If DateDiff("m", ad, d) >= repoConfig.GetLong(modConstants.CFG_HELD_REVALUE) Then clsLogger.Info "svcAssetOps.MonthEnd", "Re-appraisal due for held asset " & CStr(repos(i, repoCollateral.RPC_ID)) & " (run RunRevalueCollateral)."
            End If
        End If
    Next i
    Set MonthEnd = modMessages.OkResult("", 1)
End Function
