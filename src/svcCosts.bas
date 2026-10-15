Attribute VB_Name = "svcCosts"
Option Explicit

' Stage 2: monthly branch cost postings (accrued to GL 2300) and paying the bills through head office.

' ---- Posting a month of costs ----

' Posts every active cost row of period (YYYYMM) that is not posted yet: one journal per branch,
' branch: Dr expense GL (each cost row), Cr 2300. Safe to run twice (a CostID + Period is never posted twice).
Public Function PostPeriod(ByVal uow As clsUnitOfWork, ByVal period As String, ByRef count As Long, ByRef total As Variant) As clsResult
    Dim first As Date, last As Date, costs As Variant, posted As Variant, done As Object, i As Long, amt As Variant, st As Date, en As Date
    Dim byBranch As Object, pend As Object, b As Variant, lines As Collection, jr As clsResult, grp As String, item As Variant
    Dim branchTotal As Variant, id As String
    count = 0
    total = CDec(0)
    If Not modDateUtils.TryParsePeriod(period, first) Then Set PostPeriod = modMessages.FailResult("Period must be YYYYMM, for example 202610 (Control C25)."): Exit Function
    last = modDateUtils.EndOfMonth(first)
    costs = uow.WorkingTable(modConstants.TBL_BRANCHCOSTS)
    If IsEmpty(costs) Then Set PostPeriod = modMessages.OkResult("No cost rows.", 0): Exit Function
    posted = uow.WorkingTable(modConstants.TBL_COSTPOSTINGS)
    Set done = CreateObject("Scripting.Dictionary")
    If Not IsEmpty(posted) Then
        For i = 1 To UBound(posted, 1)
            done(CStr(posted(i, repoCost.CPC_COST)) & "|" & CStr(posted(i, repoCost.CPC_PERIOD))) = True
        Next i
    End If
    Set byBranch = CreateObject("Scripting.Dictionary")
    Set pend = CreateObject("Scripting.Dictionary")
    For i = 1 To UBound(costs, 1)
        If CStr(costs(i, repoCost.BCC_STATUS)) = "Active" And Len(CStr(costs(i, repoCost.BCC_ID))) > 0 Then
            If modDateUtils.TryToDate(costs(i, repoCost.BCC_START), st) Then
                If st <= last And (Not modDateUtils.TryToDate(costs(i, repoCost.BCC_END), en) Or en >= first) Then
                    If Not done.Exists(CStr(costs(i, repoCost.BCC_ID)) & "|" & period) Then
                        amt = repoCost.MonthlyAmount(CStr(costs(i, repoCost.BCC_FREQ)), costs(i, repoCost.BCC_AMOUNT))
                        If amt > 0 Then
                            b = CStr(costs(i, repoCost.BCC_BRANCH))
                            If Not byBranch.Exists(b) Then
                                Set byBranch(b) = New Collection
                                Set pend(b) = New Collection
                            End If
                            byBranch(b).Add svcLedger.Dr(CStr(b), CStr(costs(i, repoCost.BCC_GL)), amt)
                            pend(b).Add Array(CStr(costs(i, repoCost.BCC_ID)), CStr(costs(i, repoCost.BCC_GL)), amt)
                        End If
                    End If
                End If
            End If
        End If
    Next i
    For Each b In byBranch.Keys
        Set lines = byBranch(b)
        branchTotal = CDec(0)
        For Each item In pend(b)
            branchTotal = branchTotal + item(2)
        Next item
        lines.Add svcLedger.Cr(CStr(b), modConstants.GL_ACCRUED, branchTotal)
        grp = uow.NewGroupId()
        Set jr = svcLedger.PostJournal(uow, modConstants.JT_COST_POSTING, "Branch costs " & period & " " & CStr(b), grp, lines)
        If Not jr.Success Then Set PostPeriod = jr: Exit Function
        For Each item In pend(b)
            id = uow.NewId(modConstants.TBL_COSTPOSTINGS, "PostingID", "CP", 7)
            uow.AppendRow modConstants.TBL_COSTPOSTINGS, repoCost.PostingRow(id, CStr(item(0)), CStr(b), period, CStr(item(1)), item(2), jr.RefID)
            count = count + 1
        Next item
        total = total + branchTotal
    Next b
    Set PostPeriod = modMessages.OkResult("", count)
End Function

' RunPostCosts (current month) and RunPostCostsForMonth (period in Control C25): costs and depreciation, Admin only.
Private Function PostMonth(ByVal user As clsUser, ByVal period As String) As clsResult
    Dim uow As clsUnitOfWork, msg As String, first As Date, n As Long, tot As Variant, nd As Long, td As Variant, r As clsResult
    msg = svcSecurity.Denied(user, svcSecurity.ACT_POSTCOSTS)
    If Len(msg) > 0 Then Set PostMonth = modMessages.FailResult(msg): Exit Function
    If Not modDateUtils.TryParsePeriod(period, first) Then Set PostMonth = modMessages.FailResult("Period must be YYYYMM (Control C25)."): Exit Function
    If period > modDateUtils.PeriodOf(repoConfig.BusinessDate()) Then
        Set PostMonth = modMessages.FailResult("Period " & period & " is later than the month of the business date.")
        Exit Function
    End If
    Set uow = New clsUnitOfWork
    uow.Init repoConfig.BusinessDate(), user.UserID
    Set r = PostPeriod(uow, period, n, tot)
    If Not r.Success Then Set PostMonth = r: Exit Function
    Set r = svcFixedAssets.Depreciate(uow, period, nd, td)
    If Not r.Success Then Set PostMonth = r: Exit Function
    repoAudit.Add uow, "PostCosts", period, n & " cost rows " & modMoney.FormatAmount(tot) & ", " & nd & " assets depreciated " & modMoney.FormatAmount(td)
    uow.Commit
    Set PostMonth = modMessages.OkResult("Period " & period & ": " & n & " cost rows posted (" & modMoney.FormatAmount(tot) & " MMK), " & nd & _
        " assets depreciated (" & modMoney.FormatAmount(td) & " MMK).", n)
End Function

Public Function PostCosts(ByVal user As clsUser, ByVal inp As Object) As clsResult
    Set PostCosts = PostMonth(user, modDateUtils.PeriodOf(repoConfig.BusinessDate()))
End Function

Public Function PostCostsForMonth(ByVal user As clsUser, ByVal inp As Object) As clsResult
    Set PostCostsForMonth = PostMonth(user, CStr(inp(modConstants.INP_PERIOD)))
End Function

' ---- Paying bills ----

' Collects the unpaid cost postings (and Posted expenses) up to `period` for one branch or all (scope "ALL").
' Fills per-branch totals and the row numbers to mark Paid. posted / exp are the working tables.
Private Function Collect(ByVal posted As Variant, ByVal exps As Variant, ByVal period As String, ByVal scope As String, _
                         ByVal byBranch As Object, ByVal postRows As Collection, ByVal expRows As Collection) As Variant
    Dim i As Long, b As String, total As Variant, d As Date
    total = CDec(0)
    If Not IsEmpty(posted) Then
        For i = 1 To UBound(posted, 1)
            If CStr(posted(i, repoCost.CPC_STATUS)) = "Unpaid" And CStr(posted(i, repoCost.CPC_PERIOD)) <= period Then
                b = CStr(posted(i, repoCost.CPC_BRANCH))
                If scope = "ALL" Or scope = b Then
                    If byBranch.Exists(b) Then byBranch(b) = byBranch(b) + modMoney.ToDec(posted(i, repoCost.CPC_AMOUNT)) Else byBranch(b) = modMoney.ToDec(posted(i, repoCost.CPC_AMOUNT))
                    total = total + modMoney.ToDec(posted(i, repoCost.CPC_AMOUNT))
                    postRows.Add i
                End If
            End If
        Next i
    End If
    If Not IsEmpty(exps) Then
        For i = 1 To UBound(exps, 1)
            If CStr(exps(i, repoCost.EXC_STATUS)) = "Posted" And modDateUtils.TryToDate(exps(i, repoCost.EXC_DATE), d) Then
                If modDateUtils.PeriodOf(d) <= period Then
                    b = CStr(exps(i, repoCost.EXC_BRANCH))
                    If scope = "ALL" Or scope = b Then
                        If byBranch.Exists(b) Then byBranch(b) = byBranch(b) + modMoney.ToDec(exps(i, repoCost.EXC_AMOUNT)) Else byBranch(b) = modMoney.ToDec(exps(i, repoCost.EXC_AMOUNT))
                        total = total + modMoney.ToDec(exps(i, repoCost.EXC_AMOUNT))
                        expRows.Add i
                    End If
                End If
            End If
        Next i
    End If
    Collect = total
End Function

' Control: C25 = period up to which bills are paid, C8 = one branch (blank = all branches). HeadOffice requests, another HeadOffice approves.
Public Function PayBills(ByVal user As clsUser, ByVal inp As Object) As clsResult
    Dim uow As clsUnitOfWork, msg As String, period As String, scope As String, args As Object, total As Variant, first As Date
    Dim byBranch As Object, r As clsResult
    msg = svcSecurity.Denied(user, svcSecurity.ACT_PAYBILLS)
    If Len(msg) > 0 Then Set PayBills = modMessages.FailResult(msg): Exit Function
    period = CStr(inp(modConstants.INP_PERIOD))
    If Not modDateUtils.TryParsePeriod(period, first) Then Set PayBills = modMessages.FailResult("Period must be YYYYMM (Control C25)."): Exit Function
    scope = Trim$(CStr(inp(modConstants.INP_BRANCH)))
    If Len(scope) = 0 Then scope = "ALL"
    Set uow = New clsUnitOfWork
    uow.Init repoConfig.BusinessDate(), user.UserID
    If scope <> "ALL" And Not uow.Branches().Exists(scope) Then Set PayBills = modMessages.FailResult(modMessages.UnknownBranch(scope)): Exit Function
    Set byBranch = CreateObject("Scripting.Dictionary")
    total = Collect(uow.WorkingTable(modConstants.TBL_COSTPOSTINGS), uow.WorkingTable(modConstants.TBL_EXPENSES), period, scope, byBranch, New Collection, New Collection)
    If total <= 0 Then Set PayBills = modMessages.FailResult("There are no unpaid bills up to " & period & "."): Exit Function
    Set args = modArgs.NewArgs()
    args("period") = period
    args("scope") = scope
    Set r = svcApproval.Request(uow, modConstants.REQ_PAY_BILLS, "", "", total, modConstants.HEAD_OFFICE_ID, args, user.UserID)
    uow.Commit
    Set PayBills = r
End Function

' Per branch: Dr 2300, Cr 1900. Head office: Dr 1900 (total of the other branches), Cr 1100. One journal.
Public Function ExecPayBills(ByVal uow As clsUnitOfWork, ByVal args As Object, ByVal createdBy As String, ByVal approvedBy As String) As clsResult
    Dim posted As Variant, exps As Variant, byBranch As Object, postRows As Collection, expRows As Collection, total As Variant
    Dim b As Variant, lines As Collection, jr As clsResult, grp As String, other As Variant, ho As String, v As Variant
    ho = modConstants.HEAD_OFFICE_ID
    posted = uow.WorkingTable(modConstants.TBL_COSTPOSTINGS)
    exps = uow.WorkingTable(modConstants.TBL_EXPENSES)
    Set byBranch = CreateObject("Scripting.Dictionary")
    Set postRows = New Collection
    Set expRows = New Collection
    total = Collect(posted, exps, modArgs.Text(args, "period"), modArgs.Text(args, "scope", "ALL"), byBranch, postRows, expRows)
    If total <= 0 Then Set ExecPayBills = modMessages.FailResult("There are no unpaid bills any more."): Exit Function
    Set lines = New Collection
    other = CDec(0)
    For Each b In byBranch.Keys
        If CStr(b) = ho Then
            svcLedger.AddDr lines, ho, modConstants.GL_ACCRUED, byBranch(b)
            svcLedger.AddCr lines, ho, modConstants.GL_TREASURY, byBranch(b)
        Else
            svcLedger.AddDr lines, CStr(b), modConstants.GL_ACCRUED, byBranch(b)
            svcLedger.AddCr lines, CStr(b), modConstants.GL_INTERBRANCH, byBranch(b)
            other = other + byBranch(b)
        End If
    Next b
    svcLedger.AddDr lines, ho, modConstants.GL_INTERBRANCH, other
    svcLedger.AddCr lines, ho, modConstants.GL_TREASURY, other
    grp = uow.NewGroupId()
    Set jr = svcLedger.PostJournal(uow, modConstants.JT_BILL_PAYMENT, "Bills paid up to " & modArgs.Text(args, "period"), grp, lines)
    If Not jr.Success Then Set ExecPayBills = jr: Exit Function
    For Each v In postRows
        posted(v, repoCost.CPC_STATUS) = "Paid"
        posted(v, repoCost.CPC_PAIDDATE) = uow.BusinessDate
        posted(v, repoCost.CPC_PAIDJ) = jr.RefID
    Next v
    If postRows.Count > 0 Then uow.ReplaceTable modConstants.TBL_COSTPOSTINGS, posted
    For Each v In expRows
        exps(v, repoCost.EXC_STATUS) = "Paid"
        exps(v, repoCost.EXC_PAIDJ) = jr.RefID
    Next v
    If expRows.Count > 0 Then uow.ReplaceTable modConstants.TBL_EXPENSES, exps
    repoAudit.Add uow, "PayBills", modArgs.Text(args, "period"), "Paid " & modMoney.FormatAmount(total) & ", approved by " & approvedBy
    Set ExecPayBills = modMessages.OkResult("Bills of " & modMoney.FormatAmount(total) & " MMK paid up to " & modArgs.Text(args, "period") & ". Group " & grp & ".", postRows.Count + expRows.Count, grp)
End Function

' Unpaid cost postings plus Posted expenses = what GL 2300 must hold (used by control check 8).
Public Function UnpaidTotal(ByVal uow As clsUnitOfWork) As Variant
    Dim byBranch As Object
    Set byBranch = CreateObject("Scripting.Dictionary")
    UnpaidTotal = Collect(uow.WorkingTable(modConstants.TBL_COSTPOSTINGS), uow.WorkingTable(modConstants.TBL_EXPENSES), "999912", "ALL", byBranch, New Collection, New Collection)
End Function
