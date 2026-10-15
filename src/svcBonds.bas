Attribute VB_Name = "svcBonds"
Option Explicit

' Stage 8: government securities held by head office (GL 1620). T-bills are bought at a discount and accreted at month end;
' bonds are bought at par and pay coupons. Coupon receivable (GL 1610) is a function of the dates, so it needs no extra column:
' it equals RoundCcy(Face x Rate x (last posting date - start of the coupon period) / 365).

Private Function NeedTables() As String
    If Not modSheetUtils.TableExists(modConstants.TBL_BONDS) Then NeedTables = modMessages.MSG_SETUP_FIRST
End Function

Private Function DayBasis() As Variant
    DayBasis = repoConfig.GetDec(modConstants.CFG_DAYCOUNT)
End Function

' ---- Date helpers for one bond row ----

Private Function CouponDate(ByVal issue As Date, ByVal maturity As Date, ByVal months As Long, ByVal k As Long) As Date
    CouponDate = modDateUtils.AddMonths(issue, k * months)
    If CouponDate > maturity Then CouponDate = maturity
End Function

' Latest coupon date on or before x (maturity counts as the last coupon date). Returns issue when none.
Private Function LastCouponOnOrBefore(ByVal issue As Date, ByVal maturity As Date, ByVal months As Long, ByVal x As Date) As Date
    Dim k As Long, c As Date
    LastCouponOnOrBefore = issue
    If months <= 0 Then
        If x >= maturity Then LastCouponOnOrBefore = maturity
        Exit Function
    End If
    For k = 1 To 1000
        c = CouponDate(issue, maturity, months, k)
        If c > x Then Exit Function
        LastCouponOnOrBefore = c
        If c >= maturity Then Exit Function
    Next k
End Function

Private Function PrevMonthEndOnOrBefore(ByVal x As Date) As Date
    Dim me1 As Date
    me1 = modDateUtils.EndOfMonth(x)
    If me1 > x Then me1 = DateSerial(Year(x), Month(x), 1) - 1
    PrevMonthEndOnOrBefore = me1
End Function

' Coupon receivable of a government bond as booked at the end of day x.
Private Function BondReceivable(ByVal face As Variant, ByVal rate As Variant, ByVal issue As Date, ByVal maturity As Date, ByVal months As Long, ByVal x As Date) As Variant
    Dim pt As Date, ps As Date, cd As Date, me1 As Date
    BondReceivable = CDec(0)
    If x < issue Then Exit Function
    If x > maturity Then x = maturity
    cd = LastCouponOnOrBefore(issue, maturity, months, x)
    me1 = PrevMonthEndOnOrBefore(x)
    pt = cd
    If me1 > pt Then pt = me1
    ps = LastCouponOnOrBefore(issue, maturity, months, pt)
    If pt <= ps Then Exit Function
    BondReceivable = modMoney.RoundCcy(face * rate * (pt - ps) / DayBasis(), modConstants.CCY_MMK)
End Function

' For check 25: total coupon receivable of the Held bonds in the table as of day x.
Public Function UnpaidCoupons(ByVal bonds As Variant, ByVal x As Date) As Variant
    Dim i As Long, iss As Date, mat As Date, total As Variant
    total = CDec(0)
    If Not IsEmpty(bonds) Then
        For i = 1 To UBound(bonds, 1)
            If CStr(bonds(i, repoGov.BDC_STATUS)) = "Held" And CStr(bonds(i, repoGov.BDC_TYPE)) = "GovBond" Then
                If modDateUtils.TryToDate(bonds(i, repoGov.BDC_ISSUE), iss) And modDateUtils.TryToDate(bonds(i, repoGov.BDC_MATURITY), mat) Then
                    total = total + BondReceivable(modMoney.ToDec(bonds(i, repoGov.BDC_FACE)), modMoney.ToDec(bonds(i, repoGov.BDC_RATE)), iss, mat, CLng(modMoney.ToDec(bonds(i, repoGov.BDC_COUPONMONTHS))), x)
                End If
            End If
        Next i
    End If
    UnpaidCoupons = total
End Function

' For check 26: carrying value of the Held bonds.
Public Function HeldCarrying(ByVal bonds As Variant) As Variant
    Dim i As Long, total As Variant
    total = CDec(0)
    If Not IsEmpty(bonds) Then
        For i = 1 To UBound(bonds, 1)
            If CStr(bonds(i, repoGov.BDC_STATUS)) = "Held" Then total = total + modMoney.ToDec(bonds(i, repoGov.BDC_CARRYING))
        Next i
    End If
    HeldCarrying = total
End Function

' ---- Buying ----

' Control: C50 type (TBill or GovBond), C51 face value, C52 term in days, C53 rate (0.08 or 8).
Public Function BuyGovSecurity(ByVal user As clsUser, ByVal inp As Object) As clsResult
    Dim uow As clsUnitOfWork, msg As String, typ As String, face As Variant, days As Variant, rate As Variant, args As Object, r As clsResult
    msg = svcSecurity.Denied(user, svcSecurity.ACT_BONDS)
    If Len(msg) = 0 Then msg = NeedTables()
    If Len(msg) > 0 Then Set BuyGovSecurity = modMessages.FailResult(msg): Exit Function
    typ = Trim$(CStr(inp(modConstants.INP_BONDTYPE)))
    If StrComp(typ, "TBill", vbTextCompare) = 0 Then typ = "TBill" Else If StrComp(typ, "GovBond", vbTextCompare) = 0 Then typ = "GovBond"
    If typ <> "TBill" And typ <> "GovBond" Then Set BuyGovSecurity = modMessages.FailResult("Type must be TBill or GovBond (Control C50)."): Exit Function
    If Not modMoney.TryToDec(inp(modConstants.INP_FACE), face) Then Set BuyGovSecurity = modMessages.FailResult("Enter the face value in Control C51."): Exit Function
    If face <= 0 Or Not modMoney.IsWhole(face) Then Set BuyGovSecurity = modMessages.FailResult("The face value must be a whole amount above 0."): Exit Function
    If Not modMoney.TryToDec(inp(modConstants.INP_TERMDAYS), days) Then Set BuyGovSecurity = modMessages.FailResult("Enter the term in days in Control C52."): Exit Function
    If days < 1 Or Not modMoney.IsWhole(days) Then Set BuyGovSecurity = modMessages.FailResult("The term must be a whole number of days."): Exit Function
    If Not modMoney.TryToDec(inp(modConstants.INP_BONDRATE), rate) Then Set BuyGovSecurity = modMessages.FailResult("Enter the rate in Control C53."): Exit Function
    If rate > 1 Then rate = rate / 100
    If rate <= 0 Or rate >= 1 Then Set BuyGovSecurity = modMessages.FailResult("The rate must be above 0 and below 100%."): Exit Function
    Set uow = New clsUnitOfWork
    uow.Init repoConfig.BusinessDate(), user.UserID
    Set args = modArgs.NewArgs()
    args("type") = typ
    args("face") = CStr(face)
    args("days") = CStr(days)
    args("rate") = CStr(rate)
    Set r = svcApproval.Request(uow, modConstants.REQ_BUY_BOND, "", "", face, modConstants.HEAD_OFFICE_ID, args, user.UserID)
    uow.Commit
    Set BuyGovSecurity = modMessages.OkResult(typ & " of " & modMoney.FormatAmount(face) & " MMK requested. " & r.Message, 1, r.RefID)
End Function

' Run on approval: all rules again, then HO: Dr 1620 price, Cr 1100 price.
Public Function ExecBuyBond(ByVal uow As clsUnitOfWork, ByVal args As Object, ByVal createdBy As String, ByVal approvedBy As String) As clsResult
    Dim typ As String, face As Variant, days As Long, rate As Variant, price As Variant, lines As Collection, jr As clsResult, grp As String, ho As String, id As String, d As Date, cm As Long, carrying As Variant
    ho = modConstants.HEAD_OFFICE_ID
    d = uow.BusinessDate
    typ = modArgs.Text(args, "type")
    face = modArgs.Dec(args, "face")
    days = CLng(modArgs.Dec(args, "days"))
    rate = modArgs.Dec(args, "rate")
    If face <= 0 Or days < 1 Or rate <= 0 Then Set ExecBuyBond = modMessages.FailResult("The purchase data is not valid."): Exit Function
    If typ = "TBill" Then
        price = modMoney.RoundCcy(face * (1 - rate * days / DayBasis()), modConstants.CCY_MMK)
        carrying = price
        cm = 0
    Else
        price = face
        carrying = face
        cm = 6
    End If
    If uow.GLNet(ho, modConstants.GL_TREASURY) < price Then Set ExecBuyBond = modMessages.FailResult("The treasury balance (GL 1100) is too low for this purchase."): Exit Function
    Set lines = New Collection
    svcLedger.AddDr lines, ho, modConstants.GL_GOV_SEC, price
    svcLedger.AddCr lines, ho, modConstants.GL_TREASURY, price
    id = uow.NewId(modConstants.TBL_BONDS, "BondID", "BD", 6)
    grp = uow.NewGroupId()
    Set jr = svcLedger.PostJournal(uow, modConstants.JT_BOND_PURCHASE, typ & " " & id & " purchase", grp, lines)
    If Not jr.Success Then Set ExecBuyBond = jr: Exit Function
    uow.AppendRow modConstants.TBL_BONDS, Array(id, typ, IIf(typ = "TBill", "T-Bill ", "Government bond ") & id, face, d, d + days, rate, IIf(cm > 0, cm, Empty), price, carrying, CDec(0), d, "Held")
    repoAudit.Add uow, "BondBought", id, typ & " face " & modMoney.FormatAmount(face) & " price " & modMoney.FormatAmount(price) & ", approved by " & approvedBy
    Set ExecBuyBond = modMessages.OkResult(typ & " " & id & " bought for " & modMoney.FormatAmount(price) & " MMK (face " & modMoney.FormatAmount(face) & "). Group " & grp & ".", 1, grp)
End Function

' ---- End of day ----

Private Function PostPair(ByVal uow As clsUnitOfWork, ByVal jType As String, ByVal text As String, ByVal drGl As String, ByVal crGl As String, ByVal amount As Variant, ByVal ho As String) As clsResult
    Dim lines As Collection
    Set lines = New Collection
    svcLedger.AddDr lines, ho, drGl, amount
    svcLedger.AddCr lines, ho, crGl, amount
    Set PostPair = svcLedger.PostJournal(uow, jType, text, uow.NewGroupId(), lines)
End Function

' Pays one coupon or the last one: Dr 1100 X, Cr 1610 receivable, Cr 4310 the rest. Returns Nothing on success, a failed result otherwise.
Private Function PayCoupon(ByVal uow As clsUnitOfWork, ByVal b As Variant, ByVal i As Long, ByVal c As Date, ByVal ps As Date, ByVal iss As Date, ByVal mat As Date, ByVal months As Long, ByVal principal As Variant, ByVal ho As String) As clsResult
    Dim x As Variant, recv As Variant, lines As Collection, jr As clsResult
    x = modMoney.RoundCcy(modMoney.ToDec(b(i, repoGov.BDC_FACE)) * modMoney.ToDec(b(i, repoGov.BDC_RATE)) * (c - ps) / DayBasis(), modConstants.CCY_MMK)
    recv = BondReceivable(modMoney.ToDec(b(i, repoGov.BDC_FACE)), modMoney.ToDec(b(i, repoGov.BDC_RATE)), iss, mat, months, c - 1)
    If recv > x Then recv = x
    Set lines = New Collection
    svcLedger.AddDr lines, ho, modConstants.GL_TREASURY, x + principal
    svcLedger.AddCr lines, ho, modConstants.GL_INVEST_RECEIVABLE, recv
    svcLedger.AddCr lines, ho, modConstants.GL_GOV_INCOME, x - recv
    svcLedger.AddCr lines, ho, modConstants.GL_GOV_SEC, principal
    Set jr = svcLedger.PostJournal(uow, IIf(principal > 0, modConstants.JT_BOND_MATURITY, modConstants.JT_BOND_ACCRETION), "Bond " & CStr(b(i, 1)) & IIf(principal > 0, " maturity", " coupon"), uow.NewGroupId(), lines)
    If Not jr.Success Then Set PayCoupon = jr
End Function

' Maturities and coupon dates of day d; memo accrual of bond coupons.
Public Function BondDaily(ByVal uow As clsUnitOfWork, ByVal d As Date, ByVal stats As Object) As clsResult
    Dim b As Variant, i As Long, iss As Date, mat As Date, cm As Long, typ As String, ho As String, face As Variant, carry As Variant, lines As Collection, jr As clsResult, ps As Date, cd As Date, n As Long, top As Variant, res As clsResult
    ho = modConstants.HEAD_OFFICE_ID
    If Not modSheetUtils.TableExists(modConstants.TBL_BONDS) Then Set BondDaily = modMessages.OkResult("", 0): Exit Function
    b = uow.WorkingTable(modConstants.TBL_BONDS)
    If IsEmpty(b) Then Set BondDaily = modMessages.OkResult("", 0): Exit Function
    For i = 1 To UBound(b, 1)
        If CStr(b(i, repoGov.BDC_STATUS)) = "Held" And modDateUtils.TryToDate(b(i, repoGov.BDC_ISSUE), iss) And modDateUtils.TryToDate(b(i, repoGov.BDC_MATURITY), mat) Then
            typ = CStr(b(i, repoGov.BDC_TYPE))
            face = modMoney.ToDec(b(i, repoGov.BDC_FACE))
            cm = CLng(modMoney.ToDec(b(i, repoGov.BDC_COUPONMONTHS)))
            If mat <= d Then
                If typ = "TBill" Then
                    carry = modMoney.ToDec(b(i, repoGov.BDC_CARRYING))
                    If face > carry Then
                        Set jr = PostPair(uow, modConstants.JT_BOND_ACCRETION, "T-bill " & CStr(b(i, 1)) & " final accretion", modConstants.GL_GOV_SEC, modConstants.GL_GOV_INCOME, face - carry, ho)
                        If Not jr.Success Then Set BondDaily = jr: Exit Function
                    End If
                    Set lines = New Collection
                    svcLedger.AddDr lines, ho, modConstants.GL_TREASURY, face
                    svcLedger.AddCr lines, ho, modConstants.GL_GOV_SEC, face
                    Set jr = svcLedger.PostJournal(uow, modConstants.JT_BOND_MATURITY, "T-bill " & CStr(b(i, 1)) & " maturity", uow.NewGroupId(), lines)
                    If Not jr.Success Then Set BondDaily = jr: Exit Function
                Else
                    ps = LastCouponOnOrBefore(iss, mat, cm, mat - 1)
                    Set res = PayCoupon(uow, b, i, mat, ps, iss, mat, cm, face, ho)
                    If Not res Is Nothing Then Set BondDaily = res: Exit Function
                End If
                b(i, repoGov.BDC_STATUS) = "Matured"
                b(i, repoGov.BDC_CARRYING) = CDec(0)
                b(i, repoGov.BDC_ACCRUED) = CDec(0)
                n = n + 1
            ElseIf typ = "GovBond" Then
                cd = LastCouponOnOrBefore(iss, mat, cm, d)
                If cd = d And d > iss Then
                    ps = LastCouponOnOrBefore(iss, mat, cm, d - 1)
                    Set res = PayCoupon(uow, b, i, d, ps, iss, mat, cm, CDec(0), ho)
                    If Not res Is Nothing Then Set BondDaily = res: Exit Function
                    b(i, repoGov.BDC_ACCRUED) = CDec(0)
                    n = n + 1
                Else
                    b(i, repoGov.BDC_ACCRUED) = modMoney.ToDec(b(i, repoGov.BDC_ACCRUED)) + modMoney.RoundAccrual(face * modMoney.ToDec(b(i, repoGov.BDC_RATE)) / DayBasis())
                End If
                b(i, repoGov.BDC_LASTACCRUAL) = d
            End If
        End If
    Next i
    uow.ReplaceTable modConstants.TBL_BONDS, b
    Set BondDaily = modMessages.OkResult("", n)
End Function

' Month end: T-bill accretion to the straight-line target; bond coupon accrual posted (Dr 1610, Cr 4310).
Public Function BondMonthEnd(ByVal uow As clsUnitOfWork, ByVal d As Date, ByVal stats As Object) As clsResult
    Dim b As Variant, i As Long, iss As Date, mat As Date, cm As Long, ho As String, price As Variant, face As Variant, target As Variant, carry As Variant, post As Variant, jr As clsResult, term As Long, n As Long
    ho = modConstants.HEAD_OFFICE_ID
    If Not modSheetUtils.TableExists(modConstants.TBL_BONDS) Then Set BondMonthEnd = modMessages.OkResult("", 0): Exit Function
    b = uow.WorkingTable(modConstants.TBL_BONDS)
    If IsEmpty(b) Then Set BondMonthEnd = modMessages.OkResult("", 0): Exit Function
    For i = 1 To UBound(b, 1)
        If CStr(b(i, repoGov.BDC_STATUS)) = "Held" And modDateUtils.TryToDate(b(i, repoGov.BDC_ISSUE), iss) And modDateUtils.TryToDate(b(i, repoGov.BDC_MATURITY), mat) Then
            face = modMoney.ToDec(b(i, repoGov.BDC_FACE))
            If CStr(b(i, repoGov.BDC_TYPE)) = "TBill" Then
                price = modMoney.ToDec(b(i, repoGov.BDC_PRICE))
                term = mat - iss
                target = price + modMoney.RoundCcy((face - price) * (d - iss) / term, modConstants.CCY_MMK)
                carry = modMoney.ToDec(b(i, repoGov.BDC_CARRYING))
                post = target - carry
                If post > 0 Then
                    Set jr = PostPair(uow, modConstants.JT_BOND_ACCRETION, "T-bill " & CStr(b(i, 1)) & " accretion " & modDateUtils.PeriodOf(d), modConstants.GL_GOV_SEC, modConstants.GL_GOV_INCOME, post, ho)
                    If Not jr.Success Then Set BondMonthEnd = jr: Exit Function
                    b(i, repoGov.BDC_CARRYING) = carry + post
                    n = n + 1
                End If
            Else
                cm = CLng(modMoney.ToDec(b(i, repoGov.BDC_COUPONMONTHS)))
                post = BondReceivable(face, modMoney.ToDec(b(i, repoGov.BDC_RATE)), iss, mat, cm, d) - BondReceivable(face, modMoney.ToDec(b(i, repoGov.BDC_RATE)), iss, mat, cm, d - 1)
                If post > 0 Then
                    Set jr = PostPair(uow, modConstants.JT_BOND_ACCRETION, "Bond " & CStr(b(i, 1)) & " coupon accrual " & modDateUtils.PeriodOf(d), modConstants.GL_INVEST_RECEIVABLE, modConstants.GL_GOV_INCOME, post, ho)
                    If Not jr.Success Then Set BondMonthEnd = jr: Exit Function
                    n = n + 1
                End If
            End If
        End If
    Next i
    uow.ReplaceTable modConstants.TBL_BONDS, b
    Set BondMonthEnd = modMessages.OkResult("", n)
End Function
