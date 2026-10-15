Attribute VB_Name = "svcReportsLoans"
Option Explicit

' Stage 3 reports: LoanBook, LoanAgeing, LoanDue, LoanStatement, LoanLimits, LoanIncome.

Private Const SH_BOOK As String = "LoanBook"
Private Const SH_AGEING As String = "LoanAgeing"
Private Const SH_DUE As String = "LoanDue"
Private Const SH_STATEMENT As String = "LoanStatement"
Private Const SH_LIMITS As String = "LoanLimits"
Private Const SH_INCOME As String = "LoanIncome"

Private Sub AddTo(ByVal d As Object, ByVal key As String, ByVal v As Variant)
    If d.Exists(key) Then d(key) = d(key) + v Else d(key) = v
End Sub

Private Function Num(ByVal d As Object, ByVal key As String) As Variant
    If d.Exists(key) Then Num = d(key) Else Num = CDec(0)
End Function

Private Sub BuildBook(ByVal d As Date, ByVal loans As Variant)
    Dim cnt As Object, outst As Object, rateSum As Object, i As Long, key As String, k As Variant, parts() As String, rowsA As Collection
    Set cnt = CreateObject("Scripting.Dictionary")
    Set outst = CreateObject("Scripting.Dictionary")
    Set rateSum = CreateObject("Scripting.Dictionary")
    If Not IsEmpty(loans) Then
        For i = 1 To UBound(loans, 1)
            If CStr(loans(i, repoLoan.LOC_STATUS)) = repoLoan.LS_ACTIVE Then
                key = CStr(loans(i, repoLoan.LOC_PRODUCT)) & "|" & CStr(loans(i, repoLoan.LOC_BRANCH))
                AddTo cnt, key, 1
                AddTo outst, key, modMoney.ToDec(loans(i, repoLoan.LOC_OUTSTANDING))
                AddTo rateSum, key, modMoney.ToDec(loans(i, repoLoan.LOC_RATE))
            End If
        Next i
    End If
    Set rowsA = New Collection
    For Each k In cnt.Keys
        parts = Split(CStr(k), "|")
        rowsA.Add Array(parts(0), parts(1), cnt(k), outst(k), rateSum(k) / cnt(k))
    Next k
    modSheetUtils.PrepareReport SH_BOOK, "1,2"
    modSheetUtils.SetCell SH_BOOK, 1, 1, "Loan book (Active loans) " & modDateUtils.ToIso(d) & " (SIMULATION)", True
    modSheetUtils.WriteGrid SH_BOOK, 3, 1, svcReports.ToGrid(Array("Product", "Branch", "Loans", "Principal outstanding", "Average rate"), rowsA)
    modSheetUtils.StyleReport SH_BOOK, 3, 5, "4"
    modSheetUtils.GetSheet(SH_BOOK).Range("E4:E500").NumberFormat = "0.00%"
End Sub

Private Sub BuildAgeing(ByVal d As Date, ByVal loans As Variant)
    Dim classes As Variant, i As Long, c As Long, cnt As Object, prin As Object, prov As Object, rowsA As Collection, code As String
    classes = modSheetUtils.ReadTable(modConstants.TBL_LOANCLASSES)
    Set cnt = CreateObject("Scripting.Dictionary")
    Set prin = CreateObject("Scripting.Dictionary")
    Set prov = CreateObject("Scripting.Dictionary")
    If Not IsEmpty(loans) Then
        For i = 1 To UBound(loans, 1)
            If CStr(loans(i, repoLoan.LOC_STATUS)) = repoLoan.LS_ACTIVE Then
                code = CStr(CLng(modMoney.ToDec(loans(i, repoLoan.LOC_CLASS))))
                AddTo cnt, code, 1
                AddTo prin, code, modMoney.ToDec(loans(i, repoLoan.LOC_OUTSTANDING))
                AddTo prov, code, modMoney.ToDec(loans(i, repoLoan.LOC_PROVISION))
            End If
        Next i
    End If
    Set rowsA = New Collection
    If Not IsEmpty(classes) Then
        For c = 1 To UBound(classes, 1)
            code = CStr(CLng(modMoney.ToDec(classes(c, 1))))
            rowsA.Add Array(code, CStr(classes(c, 2)), Num(cnt, code), Num(prin, code), Num(prov, code))
        Next c
    End If
    modSheetUtils.PrepareReport SH_AGEING, "1"
    modSheetUtils.SetCell SH_AGEING, 1, 1, "Loan ageing by class " & modDateUtils.ToIso(d) & " (SIMULATION)", True
    modSheetUtils.WriteGrid SH_AGEING, 3, 1, svcReports.ToGrid(Array("Class", "Name", "Loans", "Principal", "Provision"), rowsA)
    modSheetUtils.StyleReport SH_AGEING, 3, 5, "4,5"
End Sub

Private Sub BuildDue(ByVal d As Date, ByVal loans As Variant, ByVal sched As Variant)
    Dim i As Long, due As Date, rowsA As Collection, r As Long, st As String
    Set rowsA = New Collection
    If Not IsEmpty(sched) Then
        For i = 1 To UBound(sched, 1)
            st = CStr(sched(i, repoLoan.SCC_STATUS))
            If st <> "Paid" And modDateUtils.TryToDate(sched(i, repoLoan.SCC_DUE), due) Then
                If due <= d + 30 Then
                    r = repoLoan.FindLoanRow(loans, CStr(sched(i, repoLoan.SCC_LOAN)))
                    If r > 0 Then
                        If CStr(loans(r, repoLoan.LOC_STATUS)) = repoLoan.LS_ACTIVE Then
                            rowsA.Add Array(CStr(sched(i, repoLoan.SCC_LOAN)), CStr(loans(r, repoLoan.LOC_CUSTOMER)), CLng(modMoney.ToDec(sched(i, repoLoan.SCC_NO))), due, _
                                svcLoanPayments.RowUnpaid(sched, i), st)
                        End If
                    End If
                End If
            End If
        Next i
    End If
    modSheetUtils.PrepareReport SH_DUE, "1,2"
    modSheetUtils.SetCell SH_DUE, 1, 1, "Installments due in the next 30 days (and overdue) " & modDateUtils.ToIso(d) & " (SIMULATION)", True
    modSheetUtils.WriteGrid SH_DUE, 3, 1, svcReports.ToGrid(Array("Loan", "Customer", "Installment", "Due date", "Unpaid", "Status"), rowsA)
    modSheetUtils.StyleReport SH_DUE, 3, 6, "5"
End Sub

' One loan: C28.
Private Sub BuildStatement(ByVal d As Date, ByVal loans As Variant, ByVal sched As Variant, ByVal inp As Object)
    Dim r As Long, i As Long, rowsA As Collection, pays As Variant, rowsP As Collection, id As String, top As Long
    id = CStr(inp(modConstants.INP_LOANID))
    r = repoLoan.FindLoanRow(loans, id)
    modSheetUtils.PrepareReport SH_STATEMENT, "1"
    If r = 0 Then
        modSheetUtils.SetCell SH_STATEMENT, 1, 1, "Loan statement: put a LoanID in Control C28 and run the loan reports again.", True
        Exit Sub
    End If
    modSheetUtils.SetCell SH_STATEMENT, 1, 1, "Loan statement " & id & " (SIMULATION)", True
    modSheetUtils.SetCell SH_STATEMENT, 2, 1, "Customer " & CStr(loans(r, repoLoan.LOC_CUSTOMER)) & ", product " & CStr(loans(r, repoLoan.LOC_PRODUCT)) & ", status " & CStr(loans(r, repoLoan.LOC_STATUS)) & _
        ", principal " & modMoney.FormatAmount(modMoney.ToDec(loans(r, repoLoan.LOC_PRINCIPAL))) & ", outstanding " & modMoney.FormatAmount(modMoney.ToDec(loans(r, repoLoan.LOC_OUTSTANDING))) & _
        ", days past due " & CLng(modMoney.ToDec(loans(r, repoLoan.LOC_DPD))) & ", class " & CStr(loans(r, repoLoan.LOC_CLASS)) & ", penalty accrued " & _
        modMoney.FormatAmount(modMoney.RoundCcy(modMoney.ToDec(loans(r, repoLoan.LOC_PENALTY)), modConstants.CCY_MMK))
    Set rowsA = New Collection
    If Not IsEmpty(sched) Then
        For i = 1 To UBound(sched, 1)
            If CStr(sched(i, repoLoan.SCC_LOAN)) = id Then
                rowsA.Add Array(CLng(modMoney.ToDec(sched(i, repoLoan.SCC_NO))), sched(i, repoLoan.SCC_DUE), modMoney.ToDec(sched(i, repoLoan.SCC_PRIN)), modMoney.ToDec(sched(i, repoLoan.SCC_INT)), _
                    modMoney.ToDec(sched(i, repoLoan.SCC_PRINPAID)), modMoney.ToDec(sched(i, repoLoan.SCC_INTPAID)), modMoney.ToDec(sched(i, repoLoan.SCC_PENPAID)), CStr(sched(i, repoLoan.SCC_STATUS)))
            End If
        Next i
    End If
    modSheetUtils.WriteGrid SH_STATEMENT, 4, 1, svcReports.ToGrid(Array("No", "Due date", "Principal", "Interest", "Principal paid", "Interest paid", "Penalty paid", "Status"), rowsA)
    modSheetUtils.StyleReport SH_STATEMENT, 4, 8, "3,4,5,6,7"
    top = 4 + rowsA.Count + 3
    Set rowsP = New Collection
    pays = modSheetUtils.ReadTable(modConstants.TBL_LOANPAYMENTS)
    If Not IsEmpty(pays) Then
        For i = 1 To UBound(pays, 1)
            If CStr(pays(i, 2)) = id Then rowsP.Add Array(CStr(pays(i, 1)), pays(i, 3), modMoney.ToDec(pays(i, 4)), modMoney.ToDec(pays(i, 5)), modMoney.ToDec(pays(i, 6)), modMoney.ToDec(pays(i, 7)), CStr(pays(i, 8)), CStr(pays(i, 9)))
        Next i
    End If
    modSheetUtils.SetCell SH_STATEMENT, top - 1, 1, "Payments", True
    modSheetUtils.WriteGrid SH_STATEMENT, top, 1, svcReports.ToGrid(Array("Payment", "Date", "Amount", "Penalty", "Interest", "Principal", "Source", "Type"), rowsP)
    modSheetUtils.StyleReport SH_STATEMENT, top, 8, "3,4,5,6"
End Sub

Private Sub BuildLimits(ByVal d As Date, ByVal loans As Variant)
    Dim u As clsUnitOfWork, deposits As Variant, liquid As Variant, reserveRatio As Variant, maxL2D As Variant, loanBal As Variant, cap As Variant, reserveNeed As Variant
    Dim ratioText As String, room1 As Variant, room2 As Variant, room As Variant, rowsA As Collection, ex As Object, i As Long, k As Variant, lim As Variant, st As String
    Set u = New clsUnitOfWork
    u.Init d, "SYSTEM"
    deposits = -(u.GLNetAll(modConstants.GL_SAVINGS) + u.GLNetAll(modConstants.GL_CURRENT) + u.GLNetAll(modConstants.GL_FIXED))
    liquid = u.GLNetAll(modConstants.GL_TREASURY) + u.GLNetAll(modConstants.GL_VAULT)
    reserveRatio = repoConfig.GetDec("ReserveRatio")
    maxL2D = repoConfig.GetDec(modConstants.CFG_MAX_LOAN_TO_DEPOSIT)
    loanBal = u.GLNetAll(modConstants.GL_LOANS)
    cap = -(u.GLNetAll(modConstants.GL_CAPITAL) + u.GLNetAll(modConstants.GL_RETAINED))
    reserveNeed = reserveRatio * deposits
    room1 = CDec(0)
    If reserveRatio > 0 Then room1 = liquid / reserveRatio - deposits
    room2 = CDec(0)
    If maxL2D < 1 Then room2 = (maxL2D * deposits - loanBal) / (1 - maxL2D)
    room = room1
    If room2 < room Then room = room2
    If room < 0 Then room = CDec(0)
    Set rowsA = New Collection
    rowsA.Add Array("Total deposits (2000 + 2010 + 2020)", deposits)
    rowsA.Add Array("Liquid funds (1100 + all vaults)", liquid)
    rowsA.Add Array("Reserve needed (ReserveRatio x deposits)", reserveNeed)
    rowsA.Add Array("Loans outstanding (GL 1300)", loanBal)
    ratioText = "n/a"
    If deposits > 0 Then ratioText = Format$(CDbl(loanBal) / CDbl(deposits), "0.0000")
    rowsA.Add Array("Loan-to-deposit ratio", ratioText)
    rowsA.Add Array("Room to lend (smaller of the liquidity and loan-to-deposit rooms)", room)
    rowsA.Add Array("Capital (3000 + 3100)", cap)
    rowsA.Add Array("Single borrower limit per customer", repoConfig.GetDec(modConstants.CFG_SINGLE_BORROWER) * cap)
    Set ex = CreateObject("Scripting.Dictionary")
    If Not IsEmpty(loans) Then
        For i = 1 To UBound(loans, 1)
            st = CStr(loans(i, repoLoan.LOC_STATUS))
            If st = repoLoan.LS_ACTIVE Then AddTo ex, CStr(loans(i, repoLoan.LOC_CUSTOMER)), modMoney.ToDec(loans(i, repoLoan.LOC_OUTSTANDING))
            If st = repoLoan.LS_APPLIED Or st = repoLoan.LS_APPROVED Then AddTo ex, CStr(loans(i, repoLoan.LOC_CUSTOMER)), modMoney.ToDec(loans(i, repoLoan.LOC_PRINCIPAL))
        Next i
    End If
    For Each k In ex.Keys
        rowsA.Add Array("Exposure of customer " & CStr(k), ex(k))
    Next k
    modSheetUtils.PrepareReport SH_LIMITS, "1"
    modSheetUtils.SetCell SH_LIMITS, 1, 1, "Lending limits " & modDateUtils.ToIso(d) & " (SIMULATION)", True
    modSheetUtils.WriteGrid SH_LIMITS, 3, 1, svcReports.ToGrid(Array("Item", "Amount"), rowsA)
    modSheetUtils.StyleReport SH_LIMITS, 3, 2, "2"
End Sub

Private Sub BuildIncome(ByVal d As Date)
    Dim net As Object, branches As Object, b As Variant, rowsA As Collection
    Set net = repoJournal.LoadGLNet()
    Set branches = repoBranch.LoadBranches()
    Set rowsA = New Collection
    For Each b In branches.Keys
        rowsA.Add Array(CStr(b), branches(b).Name, -Num(net, CStr(b) & "|" & modConstants.GL_LOAN_INCOME), -Num(net, CStr(b) & "|" & modConstants.GL_LOAN_PENALTY), _
            -Num(net, CStr(b) & "|" & modConstants.GL_RECOVERY), -Num(net, CStr(b) & "|" & modConstants.GL_FEE_INCOME), Num(net, CStr(b) & "|" & modConstants.GL_PROVISION_EXP))
    Next b
    modSheetUtils.PrepareReport SH_INCOME, "1"
    modSheetUtils.SetCell SH_INCOME, 1, 1, "Income from loans " & modDateUtils.ToIso(d) & " (SIMULATION)", True
    modSheetUtils.WriteGrid SH_INCOME, 3, 1, svcReports.ToGrid(Array("Branch", "Name", "Interest income", "Penalty income", "Recoveries", "Fee income (all fees)", "Provision expense"), rowsA)
    modSheetUtils.StyleReport SH_INCOME, 3, 7, "3,4,5,6,7"
End Sub

' Writes all loan report sheets (the LoanStatement shows the loan in Control C28).
Public Sub BuildAll(ByVal d As Date, ByVal inp As Object)
    Dim loans As Variant, sched As Variant
    If Not modSheetUtils.TableExists(modConstants.TBL_LOANS) Then Exit Sub
    modSheetUtils.EnsureSheet SH_BOOK
    modSheetUtils.EnsureSheet SH_AGEING
    modSheetUtils.EnsureSheet SH_DUE
    modSheetUtils.EnsureSheet SH_STATEMENT
    modSheetUtils.EnsureSheet SH_LIMITS
    modSheetUtils.EnsureSheet SH_INCOME
    loans = modSheetUtils.ReadTable(modConstants.TBL_LOANS)
    sched = modSheetUtils.ReadTable(modConstants.TBL_LOANSCHEDULE)
    BuildBook d, loans
    BuildAgeing d, loans
    BuildDue d, loans, sched
    BuildStatement d, loans, sched, inp
    BuildLimits d, loans
    BuildIncome d
End Sub

' Control: C28 = the loan for the statement. Entry from WriteLoanReports.
Public Function WriteLoanReports(ByVal user As clsUser, ByVal inp As Object) As clsResult
    Dim msg As String
    msg = svcSecurity.Denied(user, svcSecurity.ACT_LOAN_REPORTS)
    If Len(msg) > 0 Then Set WriteLoanReports = modMessages.FailResult(msg): Exit Function
    If Not modSheetUtils.TableExists(modConstants.TBL_LOANS) Then Set WriteLoanReports = modMessages.FailResult(modMessages.MSG_SETUP_FIRST): Exit Function
    repoConfig.Reload
    BuildAll repoConfig.BusinessDate(), inp
    Set WriteLoanReports = modMessages.OkResult("Loan reports written: LoanBook, LoanAgeing, LoanDue, LoanStatement, LoanLimits, LoanIncome.", 6)
End Function
