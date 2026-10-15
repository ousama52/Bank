Attribute VB_Name = "svcRepossession"
Option Explicit

' Stage 4: repossession, carrying value, impairment, sale, redemption.
' Repo rows track: CarryingValue (book value in GL 1400), OriginalCarryingValue, Deficiency (what the customer still owes).
' What is still claimed = OriginalCarryingValue + Deficiency (equals ClaimTotal right after repossession).

Private Function NeedTables() As String
    If Not modSheetUtils.TableExists(modConstants.TBL_REPOSSESSIONS) Or Not modSheetUtils.TableExists(modConstants.TBL_ASSETSALES) Or Not modSheetUtils.TableExists(modConstants.TBL_LOANS) Then
        NeedTables = "The repossession tables are missing. Run RunSetup first."
    End If
End Function

Private Function IsHeld(ByVal status As String) As Boolean
    IsHeld = (status = repoCollateral.RS_HELD Or status = repoCollateral.RS_RENTED Or status = repoCollateral.RS_INBUSINESS)
End Function

Private Function RemainingClaim(ByVal repos As Variant, ByVal i As Long) As Variant
    RemainingClaim = modMoney.ToDec(repos(i, repoCollateral.RPC_ORIGINAL)) + modMoney.ToDec(repos(i, repoCollateral.RPC_DEFICIENCY))
End Function

' ---- Repossession ----

' Returns "" when the repossession may go ahead (checked at request time and again at approval), else the reason.
' On success apprValue holds the value of the fresh Repossession appraisal.
Private Function RepossessProblem(ByVal uow As clsUnitOfWork, ByVal repos As Variant, ByVal rr As Long, ByRef lr As Long, ByRef ir As Long, ByRef apprValue As Variant) As String
    Dim loans As Variant, items As Variant, appr As Variant, i As Long, d As Date, nd As Date, best As Date, found As Boolean
    If CStr(repos(rr, repoCollateral.RPC_STATUS)) <> repoCollateral.RS_NOTICE Then RepossessProblem = "Repossession " & CStr(repos(rr, repoCollateral.RPC_ID)) & " is " & CStr(repos(rr, repoCollateral.RPC_STATUS)) & ", not Notice.": Exit Function
    If Not modDateUtils.TryToDate(repos(rr, repoCollateral.RPC_NOTICE), nd) Then RepossessProblem = "The notice date is missing.": Exit Function
    If uow.BusinessDate < nd + repoConfig.GetLong(modConstants.CFG_REPO_NOTICE_PERIOD) Then
        RepossessProblem = "The notice period is not over. Repossession is allowed from " & modDateUtils.ToIso(nd + repoConfig.GetLong(modConstants.CFG_REPO_NOTICE_PERIOD)) & "."
        Exit Function
    End If
    loans = uow.WorkingTable(modConstants.TBL_LOANS)
    items = uow.WorkingTable(modConstants.TBL_COLLATERAL)
    lr = repoLoan.FindLoanRow(loans, CStr(repos(rr, repoCollateral.RPC_LOAN)))
    ir = repoCollateral.FindCollateralRow(items, repoCollateral.COC_ID, CStr(repos(rr, repoCollateral.RPC_COLLATERAL)))
    If lr = 0 Or ir = 0 Then RepossessProblem = "The loan or the collateral of this record does not exist.": Exit Function
    If CStr(loans(lr, repoLoan.LOC_STATUS)) <> repoLoan.LS_ACTIVE Or modMoney.ToDec(loans(lr, repoLoan.LOC_DPD)) <= 0 Then
        RepossessProblem = "The loan is no longer overdue (or not Active).": Exit Function
    End If
    If CStr(items(ir, repoCollateral.COC_LIEN)) <> repoCollateral.LIEN_PLEDGED Or CStr(items(ir, repoCollateral.COC_LOAN)) <> CStr(loans(lr, repoLoan.LOC_ID)) Then
        RepossessProblem = "The item is not pledged to this loan.": Exit Function
    End If
    appr = uow.WorkingTable(modConstants.TBL_APPRAISALS)
    If Not IsEmpty(appr) Then
        For i = 1 To UBound(appr, 1)
            If CStr(appr(i, repoCollateral.APPR_COLLATERAL)) = CStr(items(ir, repoCollateral.COC_ID)) And CStr(appr(i, repoCollateral.APPR_PURPOSE)) = "Repossession" Then
                If modDateUtils.TryToDate(appr(i, repoCollateral.APPR_DATE), d) Then
                    If uow.BusinessDate - d <= repoConfig.GetLong(modConstants.CFG_APPRAISAL_VALID) And ((Not found) Or d >= best) Then
                        best = d
                        apprValue = modMoney.ToDec(appr(i, repoCollateral.APPR_VALUE))
                        found = True
                    End If
                End If
            End If
        Next i
    End If
    If Not found Then RepossessProblem = "A Repossession appraisal made within the last " & repoConfig.GetLong(modConstants.CFG_APPRAISAL_VALID) & " days is needed (RunAppraise with the note Repossession)."
End Function

' Control: C28 = the RepoID (from RunIssueDefaultNotice). HeadOffice requests, another HeadOffice approves.
Public Function Repossess(ByVal user As clsUser, ByVal inp As Object) As clsResult
    Dim uow As clsUnitOfWork, msg As String, repos As Variant, rr As Long, lr As Long, ir As Long, av As Variant, loans As Variant, args As Object, r As clsResult
    msg = svcSecurity.Denied(user, svcSecurity.ACT_REPOSSESS)
    If Len(msg) > 0 Then Set Repossess = modMessages.FailResult(msg): Exit Function
    msg = NeedTables()
    If Len(msg) > 0 Then Set Repossess = modMessages.FailResult(msg): Exit Function
    Set uow = New clsUnitOfWork
    uow.Init repoConfig.BusinessDate(), user.UserID
    repos = uow.WorkingTable(modConstants.TBL_REPOSSESSIONS)
    rr = repoCollateral.FindCollateralRow(repos, repoCollateral.RPC_ID, CStr(inp(modConstants.INP_LOANID)))
    If rr = 0 Then Set Repossess = modMessages.FailResult("Unknown repossession record. Put the RepoID in Control C28."): Exit Function
    msg = RepossessProblem(uow, repos, rr, lr, ir, av)
    If Len(msg) > 0 Then Set Repossess = modMessages.FailResult(msg): Exit Function
    loans = uow.WorkingTable(modConstants.TBL_LOANS)
    Set args = modArgs.NewArgs()
    args("repo") = CStr(repos(rr, repoCollateral.RPC_ID))
    Set r = svcApproval.Request(uow, modConstants.REQ_REPOSSESSION, "", "", modMoney.ToDec(loans(lr, repoLoan.LOC_OUTSTANDING)), CStr(repos(rr, repoCollateral.RPC_BRANCH)), args, user.UserID)
    uow.Commit
    Set Repossess = r
End Function

' B = principal + recognized unpaid interest, CV = min(fair value less costs, B).
' Dr 1400 CV, Dr 1320 provision, Dr/Cr 5300 (B - CV - provision), Cr 1300 principal, Cr 1310 interest.
Public Function ExecRepossess(ByVal uow As clsUnitOfWork, ByVal args As Object, ByVal createdBy As String, ByVal approvedBy As String) As clsResult
    Dim repos As Variant, loans As Variant, items As Variant, sched As Variant, rr As Long, lr As Long, ir As Long, av As Variant, msg As String, trow As Variant
    Dim prin As Variant, intr As Variant, b As Variant, fvlc As Variant, cv As Variant, prov As Variant, x As Variant, claim As Variant, lines As Collection, jr As clsResult
    Dim grp As String, h As String
    msg = NeedTables()
    If Len(msg) > 0 Then Set ExecRepossess = modMessages.FailResult(msg): Exit Function
    repos = uow.WorkingTable(modConstants.TBL_REPOSSESSIONS)
    rr = repoCollateral.FindCollateralRow(repos, repoCollateral.RPC_ID, modArgs.Text(args, "repo"))
    If rr = 0 Then Set ExecRepossess = modMessages.FailResult("Unknown repossession record."): Exit Function
    msg = RepossessProblem(uow, repos, rr, lr, ir, av)
    If Len(msg) > 0 Then Set ExecRepossess = modMessages.FailResult(msg): Exit Function
    loans = uow.WorkingTable(modConstants.TBL_LOANS)
    items = uow.WorkingTable(modConstants.TBL_COLLATERAL)
    sched = uow.WorkingTable(modConstants.TBL_LOANSCHEDULE)
    trow = repoCollateral.TypeRow(CStr(items(ir, repoCollateral.COC_TYPE)))
    If IsEmpty(trow) Then Set ExecRepossess = modMessages.FailResult("Unknown collateral type."): Exit Function
    h = CStr(loans(lr, repoLoan.LOC_BRANCH))
    prin = modMoney.ToDec(loans(lr, repoLoan.LOC_OUTSTANDING))
    intr = svcLoanPayments.UnpaidInterest(sched, CStr(loans(lr, repoLoan.LOC_ID)))
    b = prin + intr
    fvlc = repoCollateral.FairValueLessCosts(av, trow)
    cv = fvlc
    If b < cv Then cv = b
    prov = modMoney.ToDec(loans(lr, repoLoan.LOC_PROVISION))
    claim = b + modMoney.RoundCcy(modMoney.ToDec(loans(lr, repoLoan.LOC_PENALTY)), modConstants.CCY_MMK)
    x = b - cv - prov
    Set lines = New Collection
    svcLedger.AddDr lines, h, modConstants.GL_REPO_ASSETS, cv
    svcLedger.AddDr lines, h, modConstants.GL_LOAN_ALLOWANCE, prov
    If x > 0 Then svcLedger.AddDr lines, h, modConstants.GL_PROVISION_EXP, x
    If x < 0 Then svcLedger.AddCr lines, h, modConstants.GL_PROVISION_EXP, -x
    svcLedger.AddCr lines, h, modConstants.GL_LOANS, prin
    svcLedger.AddCr lines, h, modConstants.GL_LOAN_INTEREST_REC, intr
    grp = uow.NewGroupId()
    Set jr = svcLedger.PostJournal(uow, modConstants.JT_REPOSSESSION, "Repossession " & modArgs.Text(args, "repo"), grp, lines)
    If Not jr.Success Then Set ExecRepossess = jr: Exit Function
    loans(lr, repoLoan.LOC_STATUS) = repoLoan.LS_REPOSSESSED
    loans(lr, repoLoan.LOC_OUTSTANDING) = CDec(0)
    loans(lr, repoLoan.LOC_PROVISION) = CDec(0)
    items(ir, repoCollateral.COC_LIEN) = repoCollateral.LIEN_REPOSSESSED
    repos(rr, repoCollateral.RPC_REPODATE) = uow.BusinessDate
    repos(rr, repoCollateral.RPC_PRINCIPAL) = prin
    repos(rr, repoCollateral.RPC_INTEREST) = intr
    repos(rr, repoCollateral.RPC_PROVISION) = prov
    repos(rr, repoCollateral.RPC_CLAIM) = claim
    repos(rr, repoCollateral.RPC_FVLC) = fvlc
    repos(rr, repoCollateral.RPC_CARRYING) = cv
    repos(rr, repoCollateral.RPC_ORIGINAL) = cv
    repos(rr, repoCollateral.RPC_DEFICIENCY) = claim - cv
    repos(rr, repoCollateral.RPC_STATUS) = repoCollateral.RS_HELD
    uow.ReplaceTable modConstants.TBL_LOANS, loans
    uow.ReplaceTable modConstants.TBL_COLLATERAL, items
    uow.ReplaceTable modConstants.TBL_REPOSSESSIONS, repos
    repoAudit.Add uow, "Repossession", modArgs.Text(args, "repo"), "Loan " & CStr(loans(lr, repoLoan.LOC_ID)) & ", carrying value " & modMoney.FormatAmount(cv) & ", approved by " & approvedBy
    Set ExecRepossess = modMessages.OkResult("Repossession " & modArgs.Text(args, "repo") & " posted: asset booked at " & modMoney.FormatAmount(cv) & " MMK (fair value less costs " & modMoney.FormatAmount(fvlc) & _
        "). The customer still owes " & modMoney.FormatAmount(claim - cv) & " MMK. Group " & grp & ".", 1, grp)
End Function

' ---- Re-appraisal of a held asset: impairment or write-up ----

' Called after a new appraisal of a repossessed item. New fair value less costs below the carrying value: Dr 5320, Cr 1400.
' Above: write-up Dr 1400, Cr 5320, never above OriginalCarryingValue.
Public Function ReappraiseHeld(ByVal uow As clsUnitOfWork, ByVal collateralId As String, ByVal newValue As Variant) As clsResult
    Dim repos As Variant, items As Variant, i As Long, ir As Long, trow As Variant, nf As Variant, cv As Variant, orig As Variant, delta As Variant, lines As Collection
    Dim jr As clsResult, grp As String, h As String, id As String, note As String, opType As String
    repos = uow.WorkingTable(modConstants.TBL_REPOSSESSIONS)
    If IsEmpty(repos) Then Set ReappraiseHeld = modMessages.OkResult("", 0): Exit Function
    For i = 1 To UBound(repos, 1)
        If CStr(repos(i, repoCollateral.RPC_COLLATERAL)) = collateralId And IsHeld(CStr(repos(i, repoCollateral.RPC_STATUS))) Then
            items = uow.WorkingTable(modConstants.TBL_COLLATERAL)
            ir = repoCollateral.FindCollateralRow(items, repoCollateral.COC_ID, collateralId)
            trow = repoCollateral.TypeRow(CStr(items(ir, repoCollateral.COC_TYPE)))
            nf = repoCollateral.FairValueLessCosts(newValue, trow)
            cv = modMoney.ToDec(repos(i, repoCollateral.RPC_CARRYING))
            orig = modMoney.ToDec(repos(i, repoCollateral.RPC_ORIGINAL))
            h = CStr(repos(i, repoCollateral.RPC_BRANCH))
            Set lines = New Collection
            delta = CDec(0)
            If nf < cv Then
                delta = cv - nf
                svcLedger.AddDr lines, h, modConstants.GL_IMPAIRMENT, delta
                svcLedger.AddCr lines, h, modConstants.GL_REPO_ASSETS, delta
                note = "Impairment"
                opType = "Cost"
                cv = nf
            ElseIf nf > cv Then
                If nf > orig Then nf = orig
                delta = nf - cv
                If delta > 0 Then
                    svcLedger.AddDr lines, h, modConstants.GL_REPO_ASSETS, delta
                    svcLedger.AddCr lines, h, modConstants.GL_IMPAIRMENT, delta
                    note = "Write-up"
                    opType = "Revenue"
                    cv = cv + delta
                End If
            End If
            repos(i, repoCollateral.RPC_FVLC) = repoCollateral.FairValueLessCosts(newValue, trow)
            If delta > 0 Then
                grp = uow.NewGroupId()
                Set jr = svcLedger.PostJournal(uow, modConstants.JT_ASSET_IMPAIRMENT, note & " " & CStr(repos(i, repoCollateral.RPC_ID)), grp, lines)
                If Not jr.Success Then Set ReappraiseHeld = jr: Exit Function
                repos(i, repoCollateral.RPC_CARRYING) = cv
                id = uow.NewId(modConstants.TBL_ASSETOPS, "OpID", "OP", 6)
                uow.AppendRow modConstants.TBL_ASSETOPS, repoCollateral.OpRow(id, CStr(repos(i, repoCollateral.RPC_ID)), uow.BusinessDate, opType, delta, note, jr.RefID)
            End If
            uow.ReplaceTable modConstants.TBL_REPOSSESSIONS, repos
            Set ReappraiseHeld = modMessages.OkResult("Carrying value is now " & modMoney.FormatAmount(cv) & " MMK.", 1, "")
            Exit Function
        End If
    Next i
    Set ReappraiseHeld = modMessages.OkResult("", 0)
End Function

' ---- Settlement shared by sale and redemption ----

' The customer's account that receives a sale surplus: the loan account if usable, else any Active savings or current account of the customer.
Private Function SurplusAccount(ByVal uow As clsUnitOfWork, ByVal customerId As String, ByVal preferred As String) As clsAccount
    Dim a As clsAccount, k As Variant
    Set a = uow.GetAccount(preferred)
    If Not a Is Nothing Then
        If a.Status = modConstants.STATUS_ACTIVE And a.ProductType <> modConstants.PT_FIXED Then Set SurplusAccount = a: Exit Function
    End If
    For Each k In uow.Accounts().Keys
        Set a = uow.Accounts()(k)
        If a.CustomerID = customerId And a.Status = modConstants.STATUS_ACTIVE And a.ProductType <> modConstants.PT_FIXED Then Set SurplusAccount = a: Exit Function
    Next k
End Function

' Dr cash (1000 at P, or 1100 at head office for a transfer) Net; Cr 1400 carrying value; Cr 4030 gain or Dr 5310 loss against the claim; Cr customer account surplus.
Private Function PostSettlement(ByVal uow As clsUnitOfWork, ByVal h As String, ByVal p As String, ByVal byTransfer As Boolean, ByVal net As Variant, _
                                ByVal cv As Variant, ByVal applied As Variant, ByVal surplus As Variant, ByVal acct As clsAccount, ByVal jType As String, _
                                ByVal desc As String, ByVal grp As String, ByVal createdBy As String, ByVal approvedBy As String, ByVal reference As String) As clsResult
    Dim lines As Collection, r As String, cashGL As String, jr As clsResult
    If byTransfer Then
        r = modConstants.HEAD_OFFICE_ID
        cashGL = modConstants.GL_TREASURY
    Else
        r = p
        cashGL = modConstants.GL_VAULT
    End If
    Set lines = New Collection
    svcLedger.AddDr lines, r, cashGL, net
    If r <> h Then
        svcLedger.AddCr lines, r, modConstants.GL_INTERBRANCH, net
        svcLedger.AddDr lines, h, modConstants.GL_INTERBRANCH, net
    End If
    svcLedger.AddCr lines, h, modConstants.GL_REPO_ASSETS, cv
    If applied > cv Then svcLedger.AddCr lines, h, modConstants.GL_RECOVERY, applied - cv
    If applied < cv Then svcLedger.AddDr lines, h, modConstants.GL_LOSS_DISPOSAL, cv - applied
    If surplus > 0 Then
        If acct.BranchID = h Then
            svcLedger.AddCr lines, h, acct.DepositGL, surplus, acct.AccountNo
        Else
            svcLedger.AddCr lines, h, modConstants.GL_INTERBRANCH, surplus
            svcLedger.AddDr lines, acct.BranchID, modConstants.GL_INTERBRANCH, surplus
            svcLedger.AddCr lines, acct.BranchID, acct.DepositGL, surplus, acct.AccountNo
        End If
    End If
    Set jr = svcLedger.PostJournal(uow, jType, desc, grp, lines)
    If jr.Success And surplus > 0 Then
        svcAccounts.Effect uow, acct, modConstants.TXN_SURPLUS, modConstants.DIR_CREDIT, surplus, "", p, "Surplus from sale of repossessed asset " & reference, reference, jr.RefID, grp, createdBy, approvedBy, False
    End If
    Set PostSettlement = jr
End Function

Private Sub EndRental(ByVal uow As clsUnitOfWork, ByVal repoId As String)
    Dim data As Variant, i As Long
    If Not modSheetUtils.TableExists(modConstants.TBL_ASSETRENTALS) Then Exit Sub
    data = uow.WorkingTable(modConstants.TBL_ASSETRENTALS)
    If IsEmpty(data) Then Exit Sub
    For i = 1 To UBound(data, 1)
        If CStr(data(i, 2)) = repoId And CStr(data(i, 7)) = "Active" Then
            data(i, 6) = uow.BusinessDate
            data(i, 7) = "Ended"
            uow.ReplaceTable modConstants.TBL_ASSETRENTALS, data
        End If
    Next i
End Sub

' ---- Sale ----

' args: repo, amount (price), costs, buyer, mode (Cash or Account), branch (processing).
Public Function ExecSell(ByVal uow As clsUnitOfWork, ByVal args As Object, ByVal createdBy As String, ByVal approvedBy As String) As clsResult
    Dim repos As Variant, loans As Variant, items As Variant, rr As Long, lr As Long, ir As Long, msg As String, price As Variant, costs As Variant, net As Variant
    Dim remaining As Variant, applied As Variant, surplus As Variant, cv As Variant, acct As clsAccount, jr As clsResult, grp As String, id As String, h As String
    Dim p As String, byTransfer As Boolean, newDef As Variant, status As String
    msg = NeedTables()
    If Len(msg) > 0 Then Set ExecSell = modMessages.FailResult(msg): Exit Function
    repos = uow.WorkingTable(modConstants.TBL_REPOSSESSIONS)
    rr = repoCollateral.FindCollateralRow(repos, repoCollateral.RPC_ID, modArgs.Text(args, "repo"))
    If rr = 0 Then Set ExecSell = modMessages.FailResult("Unknown repossession record (Control C28)."): Exit Function
    If Not IsHeld(CStr(repos(rr, repoCollateral.RPC_STATUS))) Then Set ExecSell = modMessages.FailResult("Only a held asset can be sold. This one is " & CStr(repos(rr, repoCollateral.RPC_STATUS)) & "."): Exit Function
    loans = uow.WorkingTable(modConstants.TBL_LOANS)
    items = uow.WorkingTable(modConstants.TBL_COLLATERAL)
    lr = repoLoan.FindLoanRow(loans, CStr(repos(rr, repoCollateral.RPC_LOAN)))
    ir = repoCollateral.FindCollateralRow(items, repoCollateral.COC_ID, CStr(repos(rr, repoCollateral.RPC_COLLATERAL)))
    price = modArgs.Dec(args, "amount")
    costs = modArgs.Dec(args, "costs")
    net = price - costs
    If net < 0 Then Set ExecSell = modMessages.FailResult("Selling costs are higher than the price."): Exit Function
    remaining = RemainingClaim(repos, rr)
    applied = net
    If applied > remaining Then applied = remaining
    surplus = net - applied
    cv = modMoney.ToDec(repos(rr, repoCollateral.RPC_CARRYING))
    h = CStr(repos(rr, repoCollateral.RPC_BRANCH))
    p = modArgs.Text(args, "proc", modConstants.HEAD_OFFICE_ID)
    byTransfer = (UCase$(modArgs.Text(args, "mode")) = "ACCOUNT")
    If surplus > 0 Then
        Set acct = SurplusAccount(uow, CStr(loans(lr, repoLoan.LOC_CUSTOMER)), CStr(loans(lr, repoLoan.LOC_ACCOUNT)))
        If acct Is Nothing Then Set ExecSell = modMessages.FailResult("The sale leaves a surplus of " & modMoney.FormatAmount(surplus) & " MMK for the customer, but the customer has no Active account. Open one first."): Exit Function
    End If
    grp = uow.NewGroupId()
    Set jr = PostSettlement(uow, h, p, byTransfer, net, cv, applied, surplus, acct, modConstants.JT_ASSET_SALE, "Sale of repossessed asset " & modArgs.Text(args, "repo"), grp, createdBy, approvedBy, modArgs.Text(args, "repo"))
    If Not jr.Success Then Set ExecSell = jr: Exit Function
    newDef = remaining - applied
    id = uow.NewId(modConstants.TBL_ASSETSALES, "SaleID", "SL", 6)
    uow.AppendRow modConstants.TBL_ASSETSALES, Array(id, modArgs.Text(args, "repo"), uow.BusinessDate, modArgs.Text(args, "buyer"), price, costs, net, applied, surplus, cv, applied - cv, jr.RefID)
    EndRental uow, modArgs.Text(args, "repo")
    items(ir, repoCollateral.COC_LIEN) = repoCollateral.LIEN_SOLD
    repos(rr, repoCollateral.RPC_CARRYING) = CDec(0)
    repos(rr, repoCollateral.RPC_DEFICIENCY) = newDef
    If newDef = 0 Then
        repos(rr, repoCollateral.RPC_STATUS) = repoCollateral.RS_CLOSED
        repos(rr, repoCollateral.RPC_CLOSED) = uow.BusinessDate
    Else
        repos(rr, repoCollateral.RPC_STATUS) = repoCollateral.RS_SOLD
    End If
    uow.ReplaceTable modConstants.TBL_COLLATERAL, items
    uow.ReplaceTable modConstants.TBL_REPOSSESSIONS, repos
    repoAudit.Add uow, "AssetSale", id, "Repo " & modArgs.Text(args, "repo") & " sold for " & modMoney.FormatAmount(price) & ", approved by " & approvedBy
    Set ExecSell = modMessages.OkResult("Asset sold: " & modMoney.FormatAmount(net) & " MMK received, " & modMoney.FormatAmount(applied) & " applied to the claim, surplus to the customer " & modMoney.FormatAmount(surplus) & _
        ". The customer still owes " & modMoney.FormatAmount(newDef) & " MMK. Group " & grp & ".", 1, grp)
End Function

' Control: C28 repo, C7 price, C24 selling costs, C20 buyer, C27 mode (Cash or Account), C8 branch. Below SaleDiscountLimit under fair value needs HeadOffice approval.
Public Function Sell(ByVal user As clsUser, ByVal inp As Object) As clsResult
    Dim uow As clsUnitOfWork, msg As String, price As Variant, costs As Variant, args As Object, p As String, repos As Variant, items As Variant, rr As Long, ir As Long
    Dim trow As Variant, fvlc As Variant, r As clsResult
    msg = svcSecurity.Denied(user, svcSecurity.ACT_SELL)
    If Len(msg) > 0 Then Set Sell = modMessages.FailResult(msg): Exit Function
    msg = NeedTables()
    If Len(msg) > 0 Then Set Sell = modMessages.FailResult(msg): Exit Function
    If Not modArgs.ParseAmount(inp(modConstants.INP_AMOUNT), price, msg) Then Set Sell = modMessages.FailResult(msg): Exit Function
    If Not modMoney.TryToDec(inp(modConstants.INP_RESIDUAL), costs) Then costs = CDec(0)
    Set uow = New clsUnitOfWork
    uow.Init repoConfig.BusinessDate(), user.UserID
    If Not svcSecurity.ResolveBranch(uow, user, CStr(inp(modConstants.INP_BRANCH)), p, msg) Then Set Sell = modMessages.FailResult(msg): Exit Function
    repos = uow.WorkingTable(modConstants.TBL_REPOSSESSIONS)
    items = uow.WorkingTable(modConstants.TBL_COLLATERAL)
    rr = repoCollateral.FindCollateralRow(repos, repoCollateral.RPC_ID, CStr(inp(modConstants.INP_LOANID)))
    If rr = 0 Then Set Sell = modMessages.FailResult("Unknown repossession record. Put the RepoID in Control C28."): Exit Function
    ir = repoCollateral.FindCollateralRow(items, repoCollateral.COC_ID, CStr(repos(rr, repoCollateral.RPC_COLLATERAL)))
    trow = repoCollateral.TypeRow(CStr(items(ir, repoCollateral.COC_TYPE)))
    fvlc = repoCollateral.FairValueLessCosts(modMoney.ToDec(items(ir, repoCollateral.COC_VALUE)), trow)
    Set args = modArgs.NewArgs()
    args("repo") = CStr(repos(rr, repoCollateral.RPC_ID))
    args("amount") = price
    args("costs") = costs
    args("buyer") = CStr(inp(modConstants.INP_VENDOR))
    args("mode") = CStr(inp(modConstants.INP_MODE))
    args("proc") = p
    If price < fvlc * (1 - repoConfig.GetDec(modConstants.CFG_SALE_DISCOUNT)) Then
        Set r = svcApproval.Request(uow, modConstants.REQ_ASSET_SALE, "", "", price, CStr(repos(rr, repoCollateral.RPC_BRANCH)), args, user.UserID)
    Else
        Set r = ExecSell(uow, args, user.UserID, "")
    End If
    If r.Success Then uow.Commit
    Set Sell = r
End Function

' ---- Redemption: the customer pays the full remaining claim and gets the item back ----

Public Function ExecRedeem(ByVal uow As clsUnitOfWork, ByVal args As Object, ByVal createdBy As String, ByVal approvedBy As String) As clsResult
    Dim repos As Variant, items As Variant, rr As Long, ir As Long, msg As String, amt As Variant, remaining As Variant, cv As Variant, jr As clsResult, grp As String
    msg = NeedTables()
    If Len(msg) > 0 Then Set ExecRedeem = modMessages.FailResult(msg): Exit Function
    repos = uow.WorkingTable(modConstants.TBL_REPOSSESSIONS)
    rr = repoCollateral.FindCollateralRow(repos, repoCollateral.RPC_ID, modArgs.Text(args, "repo"))
    If rr = 0 Then Set ExecRedeem = modMessages.FailResult("Unknown repossession record (Control C28)."): Exit Function
    If Not IsHeld(CStr(repos(rr, repoCollateral.RPC_STATUS))) Then Set ExecRedeem = modMessages.FailResult("Only a held asset can be redeemed."): Exit Function
    amt = modArgs.Dec(args, "amount")
    remaining = RemainingClaim(repos, rr)
    If amt <> remaining Then Set ExecRedeem = modMessages.FailResult("Only the full remaining claim is accepted: " & modMoney.FormatAmount(remaining) & " MMK."): Exit Function
    items = uow.WorkingTable(modConstants.TBL_COLLATERAL)
    ir = repoCollateral.FindCollateralRow(items, repoCollateral.COC_ID, CStr(repos(rr, repoCollateral.RPC_COLLATERAL)))
    cv = modMoney.ToDec(repos(rr, repoCollateral.RPC_CARRYING))
    grp = uow.NewGroupId()
    Set jr = PostSettlement(uow, CStr(repos(rr, repoCollateral.RPC_BRANCH)), modArgs.Text(args, "proc", modConstants.HEAD_OFFICE_ID), (UCase$(modArgs.Text(args, "mode")) = "ACCOUNT"), _
        amt, cv, amt, CDec(0), Nothing, modConstants.JT_REDEMPTION, "Redemption " & modArgs.Text(args, "repo"), grp, createdBy, approvedBy, modArgs.Text(args, "repo"))
    If Not jr.Success Then Set ExecRedeem = jr: Exit Function
    EndRental uow, modArgs.Text(args, "repo")
    items(ir, repoCollateral.COC_LIEN) = repoCollateral.LIEN_REDEEMED
    items(ir, repoCollateral.COC_LOAN) = Empty
    repos(rr, repoCollateral.RPC_CARRYING) = CDec(0)
    repos(rr, repoCollateral.RPC_DEFICIENCY) = CDec(0)
    repos(rr, repoCollateral.RPC_STATUS) = repoCollateral.RS_REDEEMED
    repos(rr, repoCollateral.RPC_CLOSED) = uow.BusinessDate
    uow.ReplaceTable modConstants.TBL_COLLATERAL, items
    uow.ReplaceTable modConstants.TBL_REPOSSESSIONS, repos
    repoAudit.Add uow, "Redeem", modArgs.Text(args, "repo"), "Redeemed for " & modMoney.FormatAmount(amt) & ", approved by " & approvedBy
    Set ExecRedeem = modMessages.OkResult("Asset redeemed for " & modMoney.FormatAmount(amt) & " MMK. The item returns to the customer and the claim is cleared. Group " & grp & ".", 1, grp)
End Function

' Control: C28 repo, C7 = the full remaining claim, C27 mode, C8 branch. HeadOffice requests, another HeadOffice approves.
Public Function Redeem(ByVal user As clsUser, ByVal inp As Object) As clsResult
    Dim uow As clsUnitOfWork, msg As String, amt As Variant, args As Object, p As String, repos As Variant, rr As Long, r As clsResult
    msg = svcSecurity.Denied(user, svcSecurity.ACT_REDEEM)
    If Len(msg) > 0 Then Set Redeem = modMessages.FailResult(msg): Exit Function
    msg = NeedTables()
    If Len(msg) > 0 Then Set Redeem = modMessages.FailResult(msg): Exit Function
    If Not modArgs.ParseAmount(inp(modConstants.INP_AMOUNT), amt, msg) Then Set Redeem = modMessages.FailResult(msg): Exit Function
    Set uow = New clsUnitOfWork
    uow.Init repoConfig.BusinessDate(), user.UserID
    If Not svcSecurity.ResolveBranch(uow, user, CStr(inp(modConstants.INP_BRANCH)), p, msg) Then Set Redeem = modMessages.FailResult(msg): Exit Function
    repos = uow.WorkingTable(modConstants.TBL_REPOSSESSIONS)
    rr = repoCollateral.FindCollateralRow(repos, repoCollateral.RPC_ID, CStr(inp(modConstants.INP_LOANID)))
    If rr = 0 Then Set Redeem = modMessages.FailResult("Unknown repossession record. Put the RepoID in Control C28."): Exit Function
    Set args = modArgs.NewArgs()
    args("repo") = CStr(repos(rr, repoCollateral.RPC_ID))
    args("amount") = amt
    args("mode") = CStr(inp(modConstants.INP_MODE))
    args("proc") = p
    Set r = svcApproval.Request(uow, modConstants.REQ_REDEEM, "", "", amt, CStr(repos(rr, repoCollateral.RPC_BRANCH)), args, user.UserID)
    uow.Commit
    Set Redeem = r
End Function
