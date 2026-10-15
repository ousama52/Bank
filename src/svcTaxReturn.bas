Attribute VB_Name = "svcTaxReturn"
Option Explicit

' Stage 8: income tax return (sheet TaxReturn) and the booking of income tax (HO: Dr 5400, Cr 2220), once per year.
' Year closing journals and the tax expense itself (5400) never count as profit before tax.

Private Const SH_TAX As String = "TaxReturn"

Private Function YearFromInput(ByVal inp As Object, ByVal d As Date) As Long
    Dim t As String
    YearFromInput = Year(d)
    t = Trim$(CStr(inp(modConstants.INP_PERIOD)))
    If t Like "####*" Then YearFromInput = CLng(Left$(t, 4))
End Function

' Income, expenses (without 5400) and fines (5350, part of the expenses) of the calendar year yr from the journals.
Public Sub YearFigures(ByVal yr As Long, ByRef income As Variant, ByRef expense As Variant, ByRef fines As Variant, ByRef creditsByGl As Object)
    Dim lines As Variant, chart As Object, i As Long, gl As String, dt As Date, dr As Variant, cr As Variant
    income = CDec(0)
    expense = CDec(0)
    fines = CDec(0)
    Set creditsByGl = CreateObject("Scripting.Dictionary")
    lines = repoJournal.LoadDatedLines()
    If IsEmpty(lines) Then Exit Sub
    Set chart = repoChart.LoadChart()
    For i = 1 To UBound(lines, 1)
        If Not IsEmpty(lines(i, 1)) Then
            dt = lines(i, 1)
            If Year(dt) = yr And CStr(lines(i, 6)) <> modConstants.JT_YEAR_END Then
                gl = CStr(lines(i, 3))
                dr = lines(i, 4)
                cr = lines(i, 5)
                If chart.Exists(gl) Then
                    If chart(gl)(1) = "Income" Then
                        income = income + cr - dr
                    ElseIf chart(gl)(1) = "Expense" And gl <> modConstants.GL_INCOME_TAX_EXP Then
                        expense = expense + dr - cr
                        If gl = modConstants.GL_FINES Then fines = fines + dr - cr
                    End If
                End If
                If gl = modConstants.GL_WHT Or gl = modConstants.GL_TAX_FEES Or gl = modConstants.GL_PAYROLL_TAX Then
                    If creditsByGl.Exists(gl) Then creditsByGl(gl) = creditsByGl(gl) + cr Else creditsByGl(gl) = cr
                End If
            End If
        End If
    Next i
End Sub

Private Function AdjustmentsFor(ByVal yr As Long) As Variant
    Dim data As Variant, i As Long, s As Variant
    s = CDec(0)
    If modSheetUtils.TableExists(modConstants.TBL_TAXADJ) Then
        data = modSheetUtils.ReadTable(modConstants.TBL_TAXADJ)
        If Not IsEmpty(data) Then
            For i = 1 To UBound(data, 1)
                If CStr(data(i, 2)) = CStr(yr) Then s = s + modMoney.ToDec(data(i, 4))
            Next i
        End If
    End If
    AdjustmentsFor = s
End Function

' Income tax of year yr. Returns False and a message when the rate is missing.
Public Function ComputeTax(ByVal yr As Long, ByVal d As Date, ByRef pbt As Variant, ByRef adj As Variant, ByRef fines As Variant, ByRef taxable As Variant, ByRef tax As Variant, ByRef msg As String) As Boolean
    Dim income As Variant, expense As Variant, cr As Object, rate As Variant
    YearFigures yr, income, expense, fines, cr
    pbt = income - expense
    adj = AdjustmentsFor(yr)
    taxable = pbt + adj + fines
    If Not repoTaxRate.RateOn(modConstants.TAX_INCOME, d, rate) Then msg = "The tax rate " & modConstants.TAX_INCOME & " is missing in tblTaxRates.": Exit Function
    If taxable > 0 Then tax = modMoney.RoundCcy(taxable * rate, modConstants.CCY_MMK) Else tax = CDec(0)
    ComputeTax = True
End Function

' Control: C25 year (blank = current year).
Public Function PrepareTaxReturn(ByVal user As clsUser, ByVal inp As Object) As clsResult
    Dim msg As String, d As Date, yr As Long, pbt As Variant, adj As Variant, fines As Variant, taxable As Variant, tax As Variant, income As Variant, expense As Variant, cr As Object
    Dim rowsA As Collection, k As Variant, nm As String
    msg = svcSecurity.Denied(user, svcSecurity.ACT_TAXRETURN)
    If Len(msg) = 0 And Not modSheetUtils.TableExists(modConstants.TBL_TAXADJ) Then msg = modMessages.MSG_SETUP_FIRST
    If Len(msg) > 0 Then Set PrepareTaxReturn = modMessages.FailResult(msg): Exit Function
    d = repoConfig.BusinessDate()
    yr = YearFromInput(inp, d)
    If Not ComputeTax(yr, d, pbt, adj, fines, taxable, tax, msg) Then Set PrepareTaxReturn = modMessages.FailResult(msg): Exit Function
    YearFigures yr, income, expense, fines, cr
    Set rowsA = New Collection
    rowsA.Add Array("Income", income)
    rowsA.Add Array("Expenses (without income tax)", expense)
    rowsA.Add Array("Accounting profit before tax", pbt)
    rowsA.Add Array("Tax adjustments (table)", adj)
    rowsA.Add Array("Add back: regulatory fines (5350)", fines)
    rowsA.Add Array("Taxable profit", taxable)
    rowsA.Add Array("Income tax", tax)
    For Each k In cr.Keys
        Select Case CStr(k)
            Case modConstants.GL_WHT: nm = "Withholding tax on interest (2200) credited in the year"
            Case modConstants.GL_TAX_FEES: nm = "Commercial tax (2210) credited in the year"
            Case Else: nm = "Payroll tax and social security (2230) credited in the year"
        End Select
        rowsA.Add Array(nm, cr(k))
    Next k
    modSheetUtils.EnsureSheet SH_TAX
    modSheetUtils.PrepareReport SH_TAX, "1"
    modSheetUtils.SetCell SH_TAX, 1, 1, "Income tax return " & yr & " (SIMULATION)", True
    modSheetUtils.WriteGrid SH_TAX, 3, 1, svcReports.ToGrid(Array("Line", "MMK"), rowsA)
    modSheetUtils.StyleReport SH_TAX, 3, 2, "2"
    Set PrepareTaxReturn = modMessages.OkResult("Tax return " & yr & ": taxable profit " & modMoney.FormatAmount(taxable) & ", income tax " & modMoney.FormatAmount(tax) & " MMK. See sheet " & SH_TAX & ".", 1)
End Function

Public Function AlreadyBooked(ByVal yr As Long) As Boolean
    Dim hdr As Variant, i As Long, label As String
    label = "Income tax " & yr
    hdr = repoJournal.LoadHeaders()
    If IsEmpty(hdr) Then Exit Function
    For i = 1 To UBound(hdr, 1)
        If CStr(hdr(i, 3)) = label And CStr(hdr(i, repoJournal.JNC_TYPE)) = modConstants.JT_INCOME_TAX And CStr(hdr(i, repoJournal.JNC_STATUS)) = modConstants.STATUS_POSTED Then AlreadyBooked = True: Exit Function
    Next i
End Function

' Control: C25 year. Creates the IncomeTax approval.
Public Function RequestIncomeTax(ByVal user As clsUser, ByVal inp As Object) As clsResult
    Dim uow As clsUnitOfWork, msg As String, d As Date, yr As Long, pbt As Variant, adj As Variant, fines As Variant, taxable As Variant, tax As Variant, args As Object, r As clsResult
    msg = svcSecurity.Denied(user, svcSecurity.ACT_TAXRETURN)
    If Len(msg) = 0 And Not modSheetUtils.TableExists(modConstants.TBL_TAXADJ) Then msg = modMessages.MSG_SETUP_FIRST
    If Len(msg) > 0 Then Set RequestIncomeTax = modMessages.FailResult(msg): Exit Function
    d = repoConfig.BusinessDate()
    yr = YearFromInput(inp, d)
    If AlreadyBooked(yr) Then Set RequestIncomeTax = modMessages.FailResult("The income tax of " & yr & " is already booked."): Exit Function
    If Not ComputeTax(yr, d, pbt, adj, fines, taxable, tax, msg) Then Set RequestIncomeTax = modMessages.FailResult(msg): Exit Function
    If tax <= 0 Then Set RequestIncomeTax = modMessages.FailResult("There is no taxable profit for " & yr & ", so no tax to book."): Exit Function
    Set uow = New clsUnitOfWork
    uow.Init d, user.UserID
    Set args = modArgs.NewArgs()
    args("year") = CStr(yr)
    Set r = svcApproval.Request(uow, modConstants.REQ_INCOME_TAX, "", "", tax, modConstants.HEAD_OFFICE_ID, args, user.UserID)
    uow.Commit
    Set RequestIncomeTax = modMessages.OkResult("Income tax " & yr & " of " & modMoney.FormatAmount(tax) & " MMK requested. " & r.Message, 1, r.RefID)
End Function

' Run on approval. The tax is computed again; HO: Dr 5400, Cr 2220.
Public Function ExecBookIncomeTax(ByVal uow As clsUnitOfWork, ByVal args As Object, ByVal createdBy As String, ByVal approvedBy As String) As clsResult
    Dim yr As Long, pbt As Variant, adj As Variant, fines As Variant, taxable As Variant, tax As Variant, msg As String, lines As Collection, jr As clsResult, grp As String, ho As String
    ho = modConstants.HEAD_OFFICE_ID
    yr = CLng(modArgs.Dec(args, "year"))
    If AlreadyBooked(yr) Then Set ExecBookIncomeTax = modMessages.FailResult("The income tax of " & yr & " is already booked."): Exit Function
    If Not ComputeTax(yr, uow.BusinessDate, pbt, adj, fines, taxable, tax, msg) Then Set ExecBookIncomeTax = modMessages.FailResult(msg): Exit Function
    If tax <= 0 Then Set ExecBookIncomeTax = modMessages.FailResult("There is no taxable profit for " & yr & "."): Exit Function
    Set lines = New Collection
    svcLedger.AddDr lines, ho, modConstants.GL_INCOME_TAX_EXP, tax
    svcLedger.AddCr lines, ho, modConstants.GL_INCOME_TAX_PAYABLE, tax
    grp = uow.NewGroupId()
    Set jr = svcLedger.PostJournal(uow, modConstants.JT_INCOME_TAX, "Income tax " & yr, grp, lines)
    If Not jr.Success Then Set ExecBookIncomeTax = jr: Exit Function
    repoAudit.Add uow, "IncomeTaxBooked", CStr(yr), "Profit before tax " & modMoney.FormatAmount(pbt) & ", taxable " & modMoney.FormatAmount(taxable) & ", tax " & modMoney.FormatAmount(tax) & ", approved by " & approvedBy
    Set ExecBookIncomeTax = modMessages.OkResult("Income tax " & yr & " booked: " & modMoney.FormatAmount(tax) & " MMK. Group " & grp & ".", 1, grp)
End Function
