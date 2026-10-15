Attribute VB_Name = "svcCollateral"
Option Explicit

' Stage 4: registering and valuing collateral, pledging it to loans, releasing it, default notices and the coverage rules the loan steps use.

Private Function NeedTables() As String
    If Not modSheetUtils.TableExists(modConstants.TBL_COLLATERAL) Or Not modSheetUtils.TableExists(modConstants.TBL_COLLTYPES) Then
        NeedTables = "The collateral tables are missing. Run RunSetup first."
    End If
End Function

' ---- Coverage rules (called by svcLoans) ----

' At application: the item in Control C29 must be free, the customer's, and lend enough for the amount. "" when fine.
Public Function ApplyProblem(ByVal uow As clsUnitOfWork, ByVal customerId As String, ByVal productLtv As Variant, ByVal collateralId As String, ByVal amount As Variant) As String
    Dim data As Variant, row As Long, trow As Variant, lend As Variant, msg As String
    msg = NeedTables()
    If Len(msg) > 0 Then ApplyProblem = msg: Exit Function
    If Len(collateralId) = 0 Then ApplyProblem = "This product needs collateral. Register an item and put its CollateralID in Control C29.": Exit Function
    data = uow.WorkingTable(modConstants.TBL_COLLATERAL)
    row = repoCollateral.FindCollateralRow(data, repoCollateral.COC_ID, collateralId)
    If row = 0 Then ApplyProblem = "Unknown collateral " & collateralId & ".": Exit Function
    If CStr(data(row, repoCollateral.COC_CUSTOMER)) <> customerId Then ApplyProblem = "The collateral belongs to another customer.": Exit Function
    If CStr(data(row, repoCollateral.COC_LIEN)) <> repoCollateral.LIEN_FREE And CStr(data(row, repoCollateral.COC_LIEN)) <> repoCollateral.LIEN_RELEASED Then
        ApplyProblem = "Collateral " & collateralId & " is " & CStr(data(row, repoCollateral.COC_LIEN)) & " and cannot be offered."
        Exit Function
    End If
    trow = repoCollateral.TypeRow(CStr(data(row, repoCollateral.COC_TYPE)))
    If IsEmpty(trow) Then ApplyProblem = "Unknown collateral type " & CStr(data(row, repoCollateral.COC_TYPE)) & ".": Exit Function
    lend = repoCollateral.Lendable(modMoney.ToDec(data(row, repoCollateral.COC_VALUE)), trow, productLtv)
    If lend < amount Then ApplyProblem = "Collateral is not enough: lendable value " & modMoney.FormatAmount(lend) & " is below the loan " & modMoney.FormatAmount(amount) & "."
End Function

' At disbursement: the items pledged to the loan must together lend at least the principal.
Public Function DisbursementProblem(ByVal uow As clsUnitOfWork, ByVal loanId As String, ByVal productLtv As Variant, ByVal principal As Variant) As String
    Dim data As Variant, i As Long, trow As Variant, total As Variant, msg As String
    msg = NeedTables()
    If Len(msg) > 0 Then DisbursementProblem = msg: Exit Function
    total = CDec(0)
    data = uow.WorkingTable(modConstants.TBL_COLLATERAL)
    If Not IsEmpty(data) Then
        For i = 1 To UBound(data, 1)
            If CStr(data(i, repoCollateral.COC_LOAN)) = loanId And CStr(data(i, repoCollateral.COC_LIEN)) = repoCollateral.LIEN_PLEDGED Then
                trow = repoCollateral.TypeRow(CStr(data(i, repoCollateral.COC_TYPE)))
                If Not IsEmpty(trow) Then total = total + repoCollateral.Lendable(modMoney.ToDec(data(i, repoCollateral.COC_VALUE)), trow, productLtv)
            End If
        Next i
    End If
    If total < principal Then DisbursementProblem = "Pledged collateral lends " & modMoney.FormatAmount(total) & ", which is below the loan " & modMoney.FormatAmount(principal) & ". Pledge more collateral (RunPledgeCollateral)."
End Function

' A closed loan frees its collateral. Does nothing when Stage 4 is not set up.
Public Sub ReleaseForLoan(ByVal uow As clsUnitOfWork, ByVal loanId As String)
    Dim data As Variant, i As Long, n As Long
    If Not modSheetUtils.TableExists(modConstants.TBL_COLLATERAL) Then Exit Sub
    data = uow.WorkingTable(modConstants.TBL_COLLATERAL)
    If IsEmpty(data) Then Exit Sub
    For i = 1 To UBound(data, 1)
        If CStr(data(i, repoCollateral.COC_LOAN)) = loanId And CStr(data(i, repoCollateral.COC_LIEN)) = repoCollateral.LIEN_PLEDGED Then
            data(i, repoCollateral.COC_LIEN) = repoCollateral.LIEN_RELEASED
            data(i, repoCollateral.COC_LOAN) = Empty
            data(i, repoCollateral.COC_CUSTODY) = "CustomerKeeps"
            repoAudit.Add uow, "CollateralReleased", CStr(data(i, repoCollateral.COC_ID)), "Released from loan " & loanId
            n = n + 1
        End If
    Next i
    If n > 0 Then uow.ReplaceTable modConstants.TBL_COLLATERAL, data
End Sub

' Money recovered later on a repossessed loan lowers what the customer still owes (Deficiency).
Public Sub OnRecovery(ByVal uow As clsUnitOfWork, ByVal loanId As String, ByVal amount As Variant)
    Dim data As Variant, i As Long, d As Variant
    If Not modSheetUtils.TableExists(modConstants.TBL_REPOSSESSIONS) Then Exit Sub
    data = uow.WorkingTable(modConstants.TBL_REPOSSESSIONS)
    If IsEmpty(data) Then Exit Sub
    For i = 1 To UBound(data, 1)
        If CStr(data(i, repoCollateral.RPC_LOAN)) = loanId And CStr(data(i, repoCollateral.RPC_STATUS)) <> repoCollateral.RS_NOTICE Then
            d = modMoney.ToDec(data(i, repoCollateral.RPC_DEFICIENCY)) - amount
            If d < 0 Then d = CDec(0)
            data(i, repoCollateral.RPC_DEFICIENCY) = d
            uow.ReplaceTable modConstants.TBL_REPOSSESSIONS, data
            Exit Sub
        End If
    Next i
End Sub

' ---- Registration and appraisal ----

' args: customer, type, desc, loc, amount (value), appraiser.
Public Function ExecRegister(ByVal uow As clsUnitOfWork, ByVal args As Object, ByVal createdBy As String, ByVal approvedBy As String) As clsResult
    Dim cust As clsCustomer, trow As Variant, id As String, aid As String, amt As Variant, msg As String
    msg = NeedTables()
    If Len(msg) > 0 Then Set ExecRegister = modMessages.FailResult(msg): Exit Function
    amt = modArgs.Dec(args, "amount")
    Set cust = repoCustomer.FindCustomer(modArgs.Text(args, "customer"))
    If cust Is Nothing Then Set ExecRegister = modMessages.FailResult(modMessages.UnknownCustomer(modArgs.Text(args, "customer"))): Exit Function
    trow = repoCollateral.TypeRow(modArgs.Text(args, "type"))
    If IsEmpty(trow) Then Set ExecRegister = modMessages.FailResult("Unknown collateral type " & modArgs.Text(args, "type") & " (Control C22)."): Exit Function
    id = uow.NewId(modConstants.TBL_COLLATERAL, "CollateralID", "CL", 6)
    aid = uow.NewId(modConstants.TBL_APPRAISALS, "AppraisalID", "VA", 6)
    uow.AppendRow modConstants.TBL_COLLATERAL, repoCollateral.CollateralRow(id, cust.CustomerID, modArgs.Text(args, "type"), modArgs.Text(args, "desc"), modArgs.Text(args, "loc"), _
        cust.HomeBranchID, "CustomerKeeps", uow.BusinessDate, amt)
    uow.AppendRow modConstants.TBL_APPRAISALS, repoCollateral.AppraisalRow(aid, id, uow.BusinessDate, amt, modArgs.Text(args, "appraiser", createdBy), "Pledge", "First valuation")
    repoAudit.Add uow, "RegisterCollateral", id, modArgs.Text(args, "type") & " " & modMoney.FormatAmount(amt) & " for " & cust.CustomerID
    Set ExecRegister = modMessages.OkResult("Collateral " & id & " registered for " & cust.CustomerID & ": " & modMoney.FormatAmount(amt) & " MMK (" & modArgs.Text(args, "type") & ").", 1, "")
End Function

' Control: C13 customer, C22 type, C21 description, C7 value, C15 location. BranchManager or HeadOffice is the appraiser.
Public Function Register(ByVal user As clsUser, ByVal inp As Object) As clsResult
    Dim uow As clsUnitOfWork, msg As String, amt As Variant, args As Object, cust As clsCustomer, r As clsResult
    msg = svcSecurity.Denied(user, svcSecurity.ACT_COLLATERAL)
    If Len(msg) > 0 Then Set Register = modMessages.FailResult(msg): Exit Function
    msg = NeedTables()
    If Len(msg) > 0 Then Set Register = modMessages.FailResult(msg): Exit Function
    If Not modArgs.ParseAmount(inp(modConstants.INP_AMOUNT), amt, msg) Then Set Register = modMessages.FailResult(msg): Exit Function
    Set cust = repoCustomer.FindCustomer(CStr(inp(modConstants.INP_CUSTOMER)))
    If cust Is Nothing Then Set Register = modMessages.FailResult(modMessages.UnknownCustomer(CStr(inp(modConstants.INP_CUSTOMER)))): Exit Function
    If user.Role <> modConstants.ROLE_HEADOFFICE And cust.HomeBranchID <> user.BranchID Then Set Register = modMessages.FailResult("This customer belongs to branch " & cust.HomeBranchID & "."): Exit Function
    Set uow = New clsUnitOfWork
    uow.Init repoConfig.BusinessDate(), user.UserID
    Set args = modArgs.NewArgs()
    args("flow") = "register"
    args("customer") = cust.CustomerID
    args("type") = CStr(inp(modConstants.INP_ASSETTYPE))
    args("desc") = CStr(inp(modConstants.INP_DESC))
    args("loc") = CStr(inp(modConstants.INP_NOTE))
    args("amount") = amt
    args("appraiser") = user.UserID
    If amt > repoConfig.GetDec(modConstants.CFG_SECOND_OPINION) Then
        Set r = svcApproval.Request(uow, modConstants.REQ_APPRAISAL, "", "", amt, cust.HomeBranchID, args, user.UserID)
    Else
        Set r = ExecRegister(uow, args, user.UserID, "")
    End If
    If r.Success Then uow.Commit
    Set Register = r
End Function

' args: collateral, amount (new value), purpose, note, appraiser. A repossessed item that is re-appraised is impaired or written up.
Public Function ExecAppraise(ByVal uow As clsUnitOfWork, ByVal args As Object, ByVal createdBy As String, ByVal approvedBy As String) As clsResult
    Dim data As Variant, row As Long, aid As String, amt As Variant, r As clsResult
    amt = modArgs.Dec(args, "amount")
    data = uow.WorkingTable(modConstants.TBL_COLLATERAL)
    row = repoCollateral.FindCollateralRow(data, repoCollateral.COC_ID, modArgs.Text(args, "collateral"))
    If row = 0 Then Set ExecAppraise = modMessages.FailResult("Unknown collateral " & modArgs.Text(args, "collateral") & " (Control C29)."): Exit Function
    aid = uow.NewId(modConstants.TBL_APPRAISALS, "AppraisalID", "VA", 6)
    uow.AppendRow modConstants.TBL_APPRAISALS, repoCollateral.AppraisalRow(aid, CStr(data(row, repoCollateral.COC_ID)), uow.BusinessDate, amt, modArgs.Text(args, "appraiser", createdBy), _
        modArgs.Text(args, "purpose", "Periodic"), modArgs.Text(args, "note"))
    data(row, repoCollateral.COC_VALUE) = amt
    data(row, repoCollateral.COC_VALUEDATE) = uow.BusinessDate
    uow.ReplaceTable modConstants.TBL_COLLATERAL, data
    If CStr(data(row, repoCollateral.COC_LIEN)) = repoCollateral.LIEN_REPOSSESSED Then
        Set r = svcRepossession.ReappraiseHeld(uow, CStr(data(row, repoCollateral.COC_ID)), amt)
        If Not r.Success Then Set ExecAppraise = r: Exit Function
    End If
    repoAudit.Add uow, "Appraise", CStr(data(row, repoCollateral.COC_ID)), modArgs.Text(args, "purpose", "Periodic") & " " & modMoney.FormatAmount(amt)
    Set ExecAppraise = modMessages.OkResult("Collateral " & CStr(data(row, repoCollateral.COC_ID)) & " re-appraised at " & modMoney.FormatAmount(amt) & " MMK (" & modArgs.Text(args, "purpose", "Periodic") & ").", 1, "")
End Function

' Control: C29 collateral, C7 new value, C15 purpose (Pledge, Periodic, Repossession or Sale; default Periodic).
Public Function Appraise(ByVal user As clsUser, ByVal inp As Object) As clsResult
    Dim uow As clsUnitOfWork, msg As String, amt As Variant, args As Object, purpose As String, r As clsResult, w As String
    msg = svcSecurity.Denied(user, svcSecurity.ACT_COLLATERAL)
    If Len(msg) > 0 Then Set Appraise = modMessages.FailResult(msg): Exit Function
    msg = NeedTables()
    If Len(msg) > 0 Then Set Appraise = modMessages.FailResult(msg): Exit Function
    If Not modArgs.ParseAmount(inp(modConstants.INP_AMOUNT), amt, msg) Then Set Appraise = modMessages.FailResult(msg): Exit Function
    w = Trim$(CStr(inp(modConstants.INP_NOTE)))
    purpose = "Periodic"
    Select Case StrConv(w, vbProperCase)
        Case "Pledge", "Periodic", "Repossession", "Sale": purpose = StrConv(w, vbProperCase)
    End Select
    Set uow = New clsUnitOfWork
    uow.Init repoConfig.BusinessDate(), user.UserID
    Set args = modArgs.NewArgs()
    args("flow") = "appraise"
    args("collateral") = CStr(inp(modConstants.INP_COLLATERAL))
    args("amount") = amt
    args("purpose") = purpose
    args("note") = w
    args("appraiser") = user.UserID
    If amt > repoConfig.GetDec(modConstants.CFG_SECOND_OPINION) Then
        Set r = svcApproval.Request(uow, modConstants.REQ_APPRAISAL, "", "", amt, user.BranchID, args, user.UserID)
    Else
        Set r = ExecAppraise(uow, args, user.UserID, "")
    End If
    If r.Success Then uow.Commit
    Set Appraise = r
End Function

' ---- Pledge and release ----

' Control: C28 loan, C29 collateral. The item must be free and the customer's; the loan Applied, Approved or Active.
Public Function Pledge(ByVal user As clsUser, ByVal inp As Object) As clsResult
    Dim uow As clsUnitOfWork, msg As String, items As Variant, loans As Variant, ir As Long, lr As Long, typ As String
    msg = svcSecurity.Denied(user, svcSecurity.ACT_PLEDGE)
    If Len(msg) > 0 Then Set Pledge = modMessages.FailResult(msg): Exit Function
    msg = NeedTables()
    If Len(msg) > 0 Then Set Pledge = modMessages.FailResult(msg): Exit Function
    Set uow = New clsUnitOfWork
    uow.Init repoConfig.BusinessDate(), user.UserID
    items = uow.WorkingTable(modConstants.TBL_COLLATERAL)
    loans = uow.WorkingTable(modConstants.TBL_LOANS)
    ir = repoCollateral.FindCollateralRow(items, repoCollateral.COC_ID, CStr(inp(modConstants.INP_COLLATERAL)))
    lr = repoLoan.FindLoanRow(loans, CStr(inp(modConstants.INP_LOANID)))
    If ir = 0 Then Set Pledge = modMessages.FailResult("Unknown collateral (Control C29)."): Exit Function
    If lr = 0 Then Set Pledge = modMessages.FailResult("Unknown loan (Control C28)."): Exit Function
    If CStr(items(ir, repoCollateral.COC_LIEN)) <> repoCollateral.LIEN_FREE And CStr(items(ir, repoCollateral.COC_LIEN)) <> repoCollateral.LIEN_RELEASED Then
        Set Pledge = modMessages.FailResult("Collateral " & CStr(items(ir, repoCollateral.COC_ID)) & " is " & CStr(items(ir, repoCollateral.COC_LIEN)) & " and cannot be pledged."): Exit Function
    End If
    If CStr(items(ir, repoCollateral.COC_CUSTOMER)) <> CStr(loans(lr, repoLoan.LOC_CUSTOMER)) Then
        Set Pledge = modMessages.FailResult("The collateral belongs to another customer than the loan."): Exit Function
    End If
    Select Case CStr(loans(lr, repoLoan.LOC_STATUS))
        Case repoLoan.LS_APPLIED, repoLoan.LS_APPROVED, repoLoan.LS_ACTIVE
        Case Else
            Set Pledge = modMessages.FailResult("Loan " & CStr(loans(lr, repoLoan.LOC_ID)) & " is " & CStr(loans(lr, repoLoan.LOC_STATUS)) & ". Pledge only to an Applied, Approved or Active loan."): Exit Function
    End Select
    If user.Role <> modConstants.ROLE_HEADOFFICE And CStr(loans(lr, repoLoan.LOC_BRANCH)) <> user.BranchID Then Set Pledge = modMessages.FailResult("This loan belongs to branch " & CStr(loans(lr, repoLoan.LOC_BRANCH)) & "."): Exit Function
    typ = CStr(items(ir, repoCollateral.COC_TYPE))
    items(ir, repoCollateral.COC_LOAN) = CStr(loans(lr, repoLoan.LOC_ID))
    items(ir, repoCollateral.COC_LIEN) = repoCollateral.LIEN_PLEDGED
    If typ = "GOLD" Or typ = "VEHICLE" Then items(ir, repoCollateral.COC_CUSTODY) = "BankHolds"
    uow.ReplaceTable modConstants.TBL_COLLATERAL, items
    repoAudit.Add uow, "Pledge", CStr(items(ir, repoCollateral.COC_ID)), "Pledged to loan " & CStr(loans(lr, repoLoan.LOC_ID))
    uow.Commit
    Set Pledge = modMessages.OkResult("Collateral " & CStr(items(ir, repoCollateral.COC_ID)) & " is pledged to loan " & CStr(loans(lr, repoLoan.LOC_ID)) & ".", 1)
End Function

' Control: C28 = a Closed loan. Normally the release happens by itself when the loan closes; this is the manual way.
Public Function ReleaseManual(ByVal user As clsUser, ByVal inp As Object) As clsResult
    Dim uow As clsUnitOfWork, msg As String, loans As Variant, lr As Long
    msg = svcSecurity.Denied(user, svcSecurity.ACT_PLEDGE)
    If Len(msg) > 0 Then Set ReleaseManual = modMessages.FailResult(msg): Exit Function
    msg = NeedTables()
    If Len(msg) > 0 Then Set ReleaseManual = modMessages.FailResult(msg): Exit Function
    Set uow = New clsUnitOfWork
    uow.Init repoConfig.BusinessDate(), user.UserID
    loans = uow.WorkingTable(modConstants.TBL_LOANS)
    lr = repoLoan.FindLoanRow(loans, CStr(inp(modConstants.INP_LOANID)))
    If lr = 0 Then Set ReleaseManual = modMessages.FailResult("Unknown loan (Control C28)."): Exit Function
    If CStr(loans(lr, repoLoan.LOC_STATUS)) <> repoLoan.LS_CLOSED Then Set ReleaseManual = modMessages.FailResult("Collateral can be released by hand only when the loan is Closed."): Exit Function
    ReleaseForLoan uow, CStr(loans(lr, repoLoan.LOC_ID))
    uow.Commit
    Set ReleaseManual = modMessages.OkResult("Collateral of loan " & CStr(loans(lr, repoLoan.LOC_ID)) & " released.", 1)
End Function

' ---- Re-valuation (simulated market movement) ----

' Items whose latest appraisal is older than the type's RevalueMonths get a Periodic appraisal: value x (1 + yearly change x months / 12).
' Repossessed items use HeldAssetRevalueMonths and are impaired or written up.
Public Function Revalue(ByVal user As clsUser, ByVal inp As Object) As clsResult
    Dim uow As clsUnitOfWork, msg As String, items As Variant, i As Long, trow As Variant, d As Date, vd As Date, months As Long, lim As Long, nv As Variant, n As Long
    Dim aid As String, r As clsResult, lien As String
    msg = svcSecurity.Denied(user, svcSecurity.ACT_REVALUE)
    If Len(msg) > 0 Then Set Revalue = modMessages.FailResult(msg): Exit Function
    msg = NeedTables()
    If Len(msg) > 0 Then Set Revalue = modMessages.FailResult(msg): Exit Function
    Set uow = New clsUnitOfWork
    uow.Init repoConfig.BusinessDate(), user.UserID
    d = uow.BusinessDate
    items = uow.WorkingTable(modConstants.TBL_COLLATERAL)
    If IsEmpty(items) Then Set Revalue = modMessages.OkResult("No collateral to revalue.", 0): Exit Function
    For i = 1 To UBound(items, 1)
        lien = CStr(items(i, repoCollateral.COC_LIEN))
        If lien = repoCollateral.LIEN_PLEDGED Or lien = repoCollateral.LIEN_REPOSSESSED Then
            trow = repoCollateral.TypeRow(CStr(items(i, repoCollateral.COC_TYPE)))
            If Not IsEmpty(trow) And modDateUtils.TryToDate(items(i, repoCollateral.COC_VALUEDATE), vd) Then
                months = DateDiff("m", vd, d)
                If lien = repoCollateral.LIEN_REPOSSESSED Then lim = repoConfig.GetLong(modConstants.CFG_HELD_REVALUE) Else lim = CLng(modMoney.ToDec(trow(repoCollateral.CTC_REVALUE)))
                If months >= lim And months > 0 Then
                    nv = modMoney.RoundCcy(modMoney.ToDec(items(i, repoCollateral.COC_VALUE)) * (1 + modMoney.ToDec(trow(repoCollateral.CTC_VALUECHANGE)) * months / 12), modConstants.CCY_MMK)
                    If nv < 0 Then nv = CDec(0)
                    aid = uow.NewId(modConstants.TBL_APPRAISALS, "AppraisalID", "VA", 6)
                    uow.AppendRow modConstants.TBL_APPRAISALS, repoCollateral.AppraisalRow(aid, CStr(items(i, repoCollateral.COC_ID)), d, nv, "SYSTEM", "Periodic", "Simulated market movement")
                    items(i, repoCollateral.COC_VALUE) = nv
                    items(i, repoCollateral.COC_VALUEDATE) = d
                    If lien = repoCollateral.LIEN_REPOSSESSED Then
                        uow.ReplaceTable modConstants.TBL_COLLATERAL, items
                        Set r = svcRepossession.ReappraiseHeld(uow, CStr(items(i, repoCollateral.COC_ID)), nv)
                        If Not r.Success Then Set Revalue = r: Exit Function
                        items = uow.WorkingTable(modConstants.TBL_COLLATERAL)
                    End If
                    n = n + 1
                End If
            End If
        End If
    Next i
    uow.ReplaceTable modConstants.TBL_COLLATERAL, items
    repoAudit.Add uow, "RevalueCollateral", modDateUtils.ToIso(d), n & " items revalued"
    uow.Commit
    Set Revalue = modMessages.OkResult(n & " item(s) revalued.", n)
End Function

' ---- Default notice ----

' Control: C28 loan, C29 the item (blank when the loan has exactly one pledged item). The loan must be RepoNoticeDays past due.
Public Function IssueNotice(ByVal user As clsUser, ByVal inp As Object) As clsResult
    Dim uow As clsUnitOfWork, msg As String, loans As Variant, items As Variant, lr As Long, i As Long, ir As Long, n As Long, repos As Variant, id As String
    Dim row(1 To 16) As Variant, collat As String
    msg = svcSecurity.Denied(user, svcSecurity.ACT_NOTICE)
    If Len(msg) > 0 Then Set IssueNotice = modMessages.FailResult(msg): Exit Function
    msg = NeedTables()
    If Len(msg) > 0 Then Set IssueNotice = modMessages.FailResult(msg): Exit Function
    Set uow = New clsUnitOfWork
    uow.Init repoConfig.BusinessDate(), user.UserID
    loans = uow.WorkingTable(modConstants.TBL_LOANS)
    items = uow.WorkingTable(modConstants.TBL_COLLATERAL)
    repos = uow.WorkingTable(modConstants.TBL_REPOSSESSIONS)
    lr = repoLoan.FindLoanRow(loans, CStr(inp(modConstants.INP_LOANID)))
    If lr = 0 Then Set IssueNotice = modMessages.FailResult("Unknown loan (Control C28)."): Exit Function
    If CStr(loans(lr, repoLoan.LOC_STATUS)) <> repoLoan.LS_ACTIVE Then Set IssueNotice = modMessages.FailResult("Only an Active loan can get a default notice."): Exit Function
    If user.Role <> modConstants.ROLE_HEADOFFICE And CStr(loans(lr, repoLoan.LOC_BRANCH)) <> user.BranchID Then Set IssueNotice = modMessages.FailResult("This loan belongs to branch " & CStr(loans(lr, repoLoan.LOC_BRANCH)) & "."): Exit Function
    If modMoney.ToDec(loans(lr, repoLoan.LOC_DPD)) < repoConfig.GetDec(modConstants.CFG_REPO_NOTICE_DAYS) Then
        Set IssueNotice = modMessages.FailResult("The loan is " & CLng(modMoney.ToDec(loans(lr, repoLoan.LOC_DPD))) & " days past due. A notice needs " & repoConfig.GetLong(modConstants.CFG_REPO_NOTICE_DAYS) & " days.")
        Exit Function
    End If
    collat = CStr(inp(modConstants.INP_COLLATERAL))
    If IsEmpty(items) Then Set IssueNotice = modMessages.FailResult("tblCollateral holds no items."): Exit Function
    If Len(collat) = 0 Then
        For i = 1 To UBound(items, 1)
            If CStr(items(i, repoCollateral.COC_LOAN)) = CStr(loans(lr, repoLoan.LOC_ID)) And CStr(items(i, repoCollateral.COC_LIEN)) = repoCollateral.LIEN_PLEDGED Then
                n = n + 1
                collat = CStr(items(i, repoCollateral.COC_ID))
            End If
        Next i
        If n <> 1 Then Set IssueNotice = modMessages.FailResult("The loan has " & n & " pledged items. Put the CollateralID in Control C29."): Exit Function
    End If
    ir = repoCollateral.FindCollateralRow(items, repoCollateral.COC_ID, collat)
    If ir = 0 Then Set IssueNotice = modMessages.FailResult("Unknown collateral " & collat & "."): Exit Function
    If CStr(items(ir, repoCollateral.COC_LOAN)) <> CStr(loans(lr, repoLoan.LOC_ID)) Or CStr(items(ir, repoCollateral.COC_LIEN)) <> repoCollateral.LIEN_PLEDGED Then
        Set IssueNotice = modMessages.FailResult("Collateral " & collat & " is not pledged to this loan."): Exit Function
    End If
    If Not IsEmpty(repos) Then
        For i = 1 To UBound(repos, 1)
            If CStr(repos(i, repoCollateral.RPC_COLLATERAL)) = collat And CStr(repos(i, repoCollateral.RPC_STATUS)) <> repoCollateral.RS_REDEEMED And CStr(repos(i, repoCollateral.RPC_STATUS)) <> repoCollateral.RS_CLOSED Then
                Set IssueNotice = modMessages.FailResult("There is already a repossession record (" & CStr(repos(i, repoCollateral.RPC_ID)) & ") for this item."): Exit Function
            End If
        Next i
    End If
    id = uow.NewId(modConstants.TBL_REPOSSESSIONS, "RepoID", "RP", 6)
    row(repoCollateral.RPC_ID) = id
    row(repoCollateral.RPC_LOAN) = CStr(loans(lr, repoLoan.LOC_ID))
    row(repoCollateral.RPC_COLLATERAL) = collat
    row(repoCollateral.RPC_BRANCH) = CStr(loans(lr, repoLoan.LOC_BRANCH))
    row(repoCollateral.RPC_NOTICE) = uow.BusinessDate
    row(repoCollateral.RPC_STATUS) = repoCollateral.RS_NOTICE
    uow.AppendRow modConstants.TBL_REPOSSESSIONS, row
    repoAudit.Add uow, "DefaultNotice", id, "Loan " & CStr(loans(lr, repoLoan.LOC_ID)) & ", item " & collat
    uow.Commit
    Set IssueNotice = modMessages.OkResult("Default notice issued: repossession record " & id & ". Repossession is allowed after " & repoConfig.GetLong(modConstants.CFG_REPO_NOTICE_PERIOD) & " days, with a fresh Repossession appraisal.", 1, id)
End Function
