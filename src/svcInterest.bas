Attribute VB_Name = "svcInterest"
Option Explicit

' Deposit interest: daily accrual (4 decimals, memo only), posting at month end / maturity / break / closing,
' withholding tax, and the monthly maintenance fee.

' Withholding tax on gross interest unless the customer is tax exempt.
' Smallest amount worth posting: 1 MMK, or one cent of a foreign currency.
Public Function MinUnit(ByVal ccy As String) As Variant
    MinUnit = CDec(1) / CDec(10 ^ modMoney.Decimals(ccy))
End Function

Public Function TaxFor(ByVal cust As clsCustomer, ByVal gross As Variant, Optional ByVal ccy As String = "MMK") As Variant
    If Not cust Is Nothing Then
        If cust.TaxExempt Then TaxFor = CDec(0): Exit Function
    End If
    TaxFor = modMoney.RoundCcy(gross * repoConfig.GetDec(modConstants.CFG_WHTRATE), ccy)
End Function

' Interest posting rule: H: Dr 5000 gross, Cr deposit GL (gross - tax), Cr 2200 tax. One statement row for the net amount.
Public Function PostInterestPosting(ByVal uow As clsUnitOfWork, ByVal a As clsAccount, ByVal gross As Variant, ByVal tax As Variant, _
                                    ByVal jType As String, ByVal grp As String, ByVal createdBy As String) As clsResult
    Dim lines As Collection, jr As clsResult
    If a.Ccy <> modConstants.CCY_MMK Then Set PostInterestPosting = svcFx.FxInterest(uow, a, gross, tax, jType, grp, createdBy): Exit Function
    Set lines = New Collection
    svcLedger.AddDr lines, a.BranchID, modConstants.GL_INT_EXPENSE, gross
    svcLedger.AddCr lines, a.BranchID, a.DepositGL, gross - tax, a.AccountNo
    svcLedger.AddCr lines, a.BranchID, modConstants.GL_WHT, tax
    Set jr = svcLedger.PostJournal(uow, jType, "Interest " & a.AccountNo & " gross " & modMoney.FormatAmount(gross) & " tax " & modMoney.FormatAmount(tax), grp, lines)
    If Not jr.Success Then Set PostInterestPosting = jr: Exit Function
    svcAccounts.Effect uow, a, modConstants.TXN_INTEREST, modConstants.DIR_CREDIT, gross - tax, "", a.BranchID, _
        "Interest gross " & modMoney.FormatAmount(gross) & ", tax " & modMoney.FormatAmount(tax), "", jr.RefID, grp, createdBy, "", False
    Set PostInterestPosting = modMessages.OkResult("", 1, jr.RefID)
End Function

' ---- Daily accrual ----

' Savings and current: balance x rate / DayCountBasis, rate from the tier that holds the whole balance.
' Fixed deposits: principal x locked rate / DayCountBasis, only for dates BEFORE the maturity date.
Public Function Accrue(ByVal uow As clsUnitOfWork, ByVal d As Date, ByRef count As Long, ByRef total As Variant) As clsResult
    Dim rates As Variant, basis As Variant, k As Variant, a As clsAccount, rate As Variant, acc As Variant
    rates = repoRate.Load()
    basis = repoConfig.GetDec(modConstants.CFG_DAYCOUNT)
    count = 0
    total = CDec(0)
    For Each k In uow.Accounts().Keys
        Set a = uow.Accounts()(k)
        acc = CDec(0)
        If a.Status <> modConstants.STATUS_CLOSED And a.Balance > 0 Then
            If a.ProductType = modConstants.PT_FIXED Then
                If Not IsEmpty(a.MaturityDate) And Not IsEmpty(a.LockedRate) Then
                    If d < CDate(a.MaturityDate) Then acc = modMoney.RoundAccrual(a.Balance * a.LockedRate / basis)
                End If
            ElseIf Len(a.ProductType) > 0 Then
                If repoRate.RateFrom(rates, a.ProductCode, 0, a.Balance, d, rate) Then acc = modMoney.RoundAccrual(a.Balance * rate / basis)
            End If
        End If
        If acc > 0 Then
            a.AccruedInterest = a.AccruedInterest + acc
            a.LastAccrualDate = d
            a.Dirty = True
            count = count + 1
            total = total + acc
        End If
    Next k
    Set Accrue = modMessages.OkResult("", count)
End Function

' ---- Month end ----

' Pays accrued interest of savings and current accounts (gross rounded to whole MMK, at least 1) and keeps the fraction.
Public Function PostMonthEnd(ByVal uow As clsUnitOfWork, ByVal d As Date, ByRef posted As Long) As clsResult
    Dim custs As Object, k As Variant, a As clsAccount, g As Variant, w As Variant, grp As String, r As clsResult, c As clsCustomer
    Set custs = repoCustomer.LoadCustomers()
    posted = 0
    For Each k In uow.Accounts().Keys
        Set a = uow.Accounts()(k)
        If a.Status <> modConstants.STATUS_CLOSED And (a.ProductType = modConstants.PT_SAVINGS Or a.ProductType = modConstants.PT_CURRENT) Then
            g = modMoney.RoundCcy(a.AccruedInterest, a.Ccy)
            If g >= MinUnit(a.Ccy) Then
                Set c = Nothing
                If custs.Exists(a.CustomerID) Then Set c = custs(a.CustomerID)
                w = TaxFor(c, g, a.Ccy)
                grp = uow.NewGroupId()
                Set r = PostInterestPosting(uow, a, g, w, modConstants.JT_INTEREST, grp, "SYSTEM")
                If Not r.Success Then Set PostMonthEnd = r: Exit Function
                a.AccruedInterest = a.AccruedInterest - g
                a.Dirty = True
                posted = posted + 1
            End If
        End If
    Next k
    Set PostMonthEnd = modMessages.OkResult("", posted)
End Function

' Monthly maintenance fee: Active current accounts below the product minimum balance, if the balance can cover the fee.
Public Function ChargeMaintenance(ByVal uow As clsUnitOfWork, ByVal d As Date, ByRef charged As Long) As clsResult
    Dim prods As Object, k As Variant, a As clsAccount, p As clsProduct, fee As Variant, tax As Variant, msg As String
    Dim lines As Collection, jr As clsResult, grp As String
    Set prods = repoProduct.LoadProducts()
    charged = 0
    For Each k In uow.Accounts().Keys
        Set a = uow.Accounts()(k)
        If a.Status = modConstants.STATUS_ACTIVE And a.ProductType = modConstants.PT_CURRENT And a.Ccy = modConstants.CCY_MMK And prods.Exists(a.ProductCode) Then
            Set p = prods(a.ProductCode)
            If Len(p.MaintenanceFeeCode) > 0 And a.Balance < p.MinBalance Then
                If Not svcFees.Compute(p.MaintenanceFeeCode, a.Balance, fee, tax, msg) Then
                    Set ChargeMaintenance = modMessages.FailResult(msg)
                    Exit Function
                End If
                If fee + tax > 0 Then
                    If a.Balance >= fee + tax Then
                        grp = uow.NewGroupId()
                        Set lines = New Collection
                        svcLedger.AddDr lines, a.BranchID, a.DepositGL, fee + tax, a.AccountNo
                        svcLedger.AddCr lines, a.BranchID, modConstants.GL_FEE_INCOME, fee
                        svcLedger.AddCr lines, a.BranchID, modConstants.GL_TAX_FEES, tax
                        Set jr = svcLedger.PostJournal(uow, modConstants.JT_MAINT_FEE, "Maintenance fee " & a.AccountNo, grp, lines)
                        If Not jr.Success Then Set ChargeMaintenance = jr: Exit Function
                        svcAccounts.FeeRows uow, a, fee, tax, p.MaintenanceFeeCode, "", a.BranchID, jr.RefID, grp, "SYSTEM", ""
                        charged = charged + 1
                    Else
                        clsLogger.Warn "svcInterest.ChargeMaintenance", "Maintenance fee skipped for " & a.AccountNo & ": balance " & modMoney.FormatAmount(a.Balance) & " cannot cover it."
                    End If
                End If
            End If
        End If
    Next k
    Set ChargeMaintenance = modMessages.OkResult("", charged)
End Function
