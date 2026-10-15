Attribute VB_Name = "svcTaxes"
Option Explicit

' Paying the taxes held in 2200 (withholding on interest), 2210 (commercial tax), 2230 (payroll tax and social security, Stage 5)
' and 2220 (income tax, Stage 8). The government is not a party in the books: paying tax is money leaving the bank (treasury 1100).
' Withholding tax held in a foreign currency is paid in kind from head office foreign cash (GL 1700).

Private Function TaxGls() As Variant
    TaxGls = Array(modConstants.GL_WHT, modConstants.GL_TAX_FEES, modConstants.GL_PAYROLL_TAX, modConstants.GL_INCOME_TAX_PAYABLE)
End Function

Private Function CodeOf(ByVal gl As String) As String
    Select Case gl
        Case modConstants.GL_WHT: CodeOf = modConstants.TAX_WHT
        Case modConstants.GL_TAX_FEES: CodeOf = modConstants.TAX_COMMERCIAL
        Case modConstants.GL_PAYROLL_TAX: CodeOf = modConstants.TAX_PAYROLL_SSB
        Case Else: CodeOf = modConstants.TAX_INCOME
    End Select
End Function

Private Function IsTaxGl(ByVal gl As String) As Boolean
    Dim g As Variant
    For Each g In TaxGls()
        If CStr(g) = gl Then IsTaxGl = True: Exit Function
    Next g
End Function

' mm: "branch|gl" -> MMK amount payable (credit balance without the foreign groups). fx: "branch|gl|ccy" -> Array(foreign, base) payable.
' Returns the total in MMK equivalent. Foreign groups are only taken while head office holds enough of that currency.
Private Function Collect(ByVal uow As clsUnitOfWork, ByVal mm As Object, ByVal fx As Object, ByVal byCode As Object) As Variant
    Dim b As Variant, g As Variant, v As Variant, k As Variant, parts() As String, fxBase As Object, key As String, need As Object, f As Variant, bs As Variant, total As Variant, ho As String
    ho = modConstants.HEAD_OFFICE_ID
    Set fxBase = CreateObject("Scripting.Dictionary")
    Set need = CreateObject("Scripting.Dictionary")
    total = CDec(0)
    For Each k In uow.FxGroups().Keys
        parts = Split(CStr(k), "|")
        If IsTaxGl(parts(1)) And uow.FxGroups()(k)(0) < 0 Then
            f = -uow.FxGroups()(k)(0)
            bs = -uow.FxGroups()(k)(1)
            If bs < 0 Then bs = CDec(0)
            If Not need.Exists(parts(2)) Then need(parts(2)) = CDec(0)
            If uow.FxBalance(ho, modConstants.GL_FX_CASH, parts(2)) >= need(parts(2)) + f Then
                need(parts(2)) = need(parts(2)) + f
                fx(CStr(k)) = Array(f, bs)
                total = total + bs
                If byCode.Exists(CodeOf(parts(1))) Then byCode(CodeOf(parts(1))) = byCode(CodeOf(parts(1))) + bs Else byCode(CodeOf(parts(1))) = bs
            End If
        End If
        If IsTaxGl(parts(1)) Then
            key = parts(0) & "|" & parts(1)
            If fxBase.Exists(key) Then fxBase(key) = fxBase(key) + uow.FxGroups()(k)(1) Else fxBase(key) = uow.FxGroups()(k)(1)
        End If
    Next k
    For Each b In uow.Branches().Keys
        For Each g In TaxGls()
            v = -uow.GLNet(CStr(b), CStr(g))
            key = CStr(b) & "|" & CStr(g)
            If fxBase.Exists(key) Then v = v + fxBase(key)
            If v > 0 Then
                mm(key) = v
                total = total + v
                If byCode.Exists(CodeOf(CStr(g))) Then byCode(CodeOf(CStr(g))) = byCode(CodeOf(CStr(g))) + v Else byCode(CodeOf(CStr(g))) = v
            End If
        Next g
    Next b
    Collect = total
End Function

' Control: nothing. HeadOffice requests, another HeadOffice user approves.
Public Function PayTaxes(ByVal user As clsUser, ByVal inp As Object) As clsResult
    Dim uow As clsUnitOfWork, msg As String, total As Variant, r As clsResult
    msg = svcSecurity.Denied(user, svcSecurity.ACT_PAYTAXES)
    If Len(msg) > 0 Then Set PayTaxes = modMessages.FailResult(msg): Exit Function
    Set uow = New clsUnitOfWork
    uow.Init repoConfig.BusinessDate(), user.UserID
    total = Collect(uow, CreateObject("Scripting.Dictionary"), CreateObject("Scripting.Dictionary"), CreateObject("Scripting.Dictionary"))
    If total <= 0 Then Set PayTaxes = modMessages.FailResult("No tax is payable: GL 2200, 2210, 2220 and 2230 hold no credit balance."): Exit Function
    Set r = svcApproval.Request(uow, modConstants.REQ_PAY_TAXES, "", "", total, modConstants.HEAD_OFFICE_ID, modArgs.NewArgs(), user.UserID)
    uow.Commit
    Set PayTaxes = r
End Function

' Per branch: Dr each tax GL, Cr 1900 (sum). Head office: Dr 1900 (total of the other branches), Cr 1100, and Cr 1700 for foreign tax paid in kind.
' One journal, one TaxPayments row per tax code.
Public Function ExecPayTaxes(ByVal uow As clsUnitOfWork, ByVal args As Object, ByVal createdBy As String, ByVal approvedBy As String) As clsResult
    Dim mm As Object, fx As Object, byCode As Object, total As Variant, k As Variant, parts() As String, lines As Collection, jr As clsResult, grp As String, ho As String
    Dim branchTot As Object, otherMmk As Variant, otherFx As Object, b As String, rate As Variant, period As String, id As String, mmkTotal As Variant, c As Variant, fxKey As String, f As Variant, bs As Variant, a As Variant
    ho = modConstants.HEAD_OFFICE_ID
    Set mm = CreateObject("Scripting.Dictionary")
    Set fx = CreateObject("Scripting.Dictionary")
    Set byCode = CreateObject("Scripting.Dictionary")
    total = Collect(uow, mm, fx, byCode)
    If total <= 0 Then Set ExecPayTaxes = modMessages.FailResult("No tax is payable any more."): Exit Function
    Set lines = New Collection
    Set branchTot = CreateObject("Scripting.Dictionary")
    Set otherFx = CreateObject("Scripting.Dictionary")
    mmkTotal = CDec(0)
    For Each k In mm.Keys
        parts = Split(CStr(k), "|")
        svcLedger.AddDr lines, parts(0), parts(1), mm(k)
        If branchTot.Exists(parts(0)) Then branchTot(parts(0)) = branchTot(parts(0)) + mm(k) Else branchTot(parts(0)) = mm(k)
        mmkTotal = mmkTotal + mm(k)
    Next k
    For Each k In fx.Keys
        parts = Split(CStr(k), "|")
        f = fx(k)(0)
        bs = fx(k)(1)
        rate = repoFxRate.LatestMid(parts(2), uow.BusinessDate)
        svcLedger.AddDrFx lines, parts(0), parts(1), parts(2), f, rate, bs
        If branchTot.Exists(parts(0)) Then branchTot(parts(0)) = branchTot(parts(0)) + bs Else branchTot(parts(0)) = bs
        If otherFx.Exists(parts(2)) Then
            a = otherFx(parts(2))
            otherFx(parts(2)) = Array(a(0) + f, a(1) + bs)
        Else
            otherFx(parts(2)) = Array(f, bs)
        End If
    Next k
    ' Credit side: other branches clear through 1900; head office pays the cash and the foreign cash.
    otherMmk = CDec(0)
    For Each k In branchTot.Keys
        b = CStr(k)
        If b <> ho Then svcLedger.AddCr lines, b, modConstants.GL_INTERBRANCH, branchTot(k)
    Next k
    For Each k In branchTot.Keys
        If CStr(k) <> ho Then otherMmk = otherMmk + branchTot(k)
    Next k
    svcLedger.AddDr lines, ho, modConstants.GL_INTERBRANCH, otherMmk
    ' MMK paid from treasury = total minus all foreign bases
    Dim fxBaseTot As Variant
    fxBaseTot = CDec(0)
    For Each k In otherFx.Keys
        fxBaseTot = fxBaseTot + otherFx(k)(1)
        rate = repoFxRate.LatestMid(CStr(k), uow.BusinessDate)
        svcLedger.AddCrFx lines, ho, modConstants.GL_FX_CASH, CStr(k), otherFx(k)(0), rate, otherFx(k)(1)
    Next k
    If uow.GLNet(ho, modConstants.GL_TREASURY) < total - fxBaseTot Then Set ExecPayTaxes = modMessages.FailResult("The treasury balance (GL 1100) is too low to pay the taxes."): Exit Function
    svcLedger.AddCr lines, ho, modConstants.GL_TREASURY, total - fxBaseTot
    grp = uow.NewGroupId()
    Set jr = svcLedger.PostJournal(uow, modConstants.JT_TAX_PAYMENT, "Taxes paid to the government", grp, lines)
    If Not jr.Success Then Set ExecPayTaxes = jr: Exit Function
    period = modDateUtils.PeriodOf(uow.BusinessDate)
    For Each c In byCode.Keys
        id = uow.NewId(modConstants.TBL_TAXPAYMENTS, "TaxPaymentID", "TP", 6)
        uow.AppendRow modConstants.TBL_TAXPAYMENTS, Array(id, CStr(c), uow.BusinessDate, period, byCode(c), jr.RefID, "Approved by " & approvedBy)
    Next c
    repoAudit.Add uow, "PayTaxes", period, "Paid " & modMoney.FormatAmount(total) & " MMK equivalent (" & byCode.Count & " tax code(s))"
    Set ExecPayTaxes = modMessages.OkResult("Taxes of " & modMoney.FormatAmount(total) & " MMK paid. Group " & grp & ".", 1, grp)
End Function
