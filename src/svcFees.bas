Attribute VB_Name = "svcFees"
Option Explicit

' Fee = FlatAmount (Method Flat), else RoundCcy(Amount x Rate) with MinFee as floor and MaxFee as cap (0 = no cap).
' Tax on a fee is commercial tax (Stage 2); it stays 0 until the fee has Taxable = Yes.

Public Function FeeAmount(ByVal f As clsFee, ByVal amount As Variant) As Variant
    Dim v As Variant
    If f.Method = "Flat" Then
        v = f.FlatAmount
    Else
        v = modMoney.RoundCcy(amount * f.Rate, f.Ccy)
        If v < f.MinFee Then v = f.MinFee
        If f.MaxFee > 0 Then
            If v > f.MaxFee Then v = f.MaxFee
        End If
    End If
    FeeAmount = v
End Function

' Commercial tax (tblTaxRates COMMERCIAL_FEES) on a taxable fee. 0 when the fee is not taxable or no rate is set.
Public Function TaxOn(ByVal f As clsFee, ByVal fee As Variant) As Variant
    Dim r As Variant
    TaxOn = CDec(0)
    If Not f.Taxable Or fee <= 0 Then Exit Function
    If repoTaxRate.RateOn(modConstants.TAX_COMMERCIAL, repoConfig.BusinessDate(), r) Then TaxOn = modMoney.RoundCcy(fee * r, f.Ccy)
End Function

' Looks the fee code up and returns fee and tax. False (with msg) when the fee code is missing.
Public Function Compute(ByVal feeCode As String, ByVal amount As Variant, ByRef fee As Variant, ByRef tax As Variant, ByRef msg As String) As Boolean
    Dim f As clsFee
    fee = CDec(0)
    tax = CDec(0)
    Set f = repoFee.FindFee(feeCode)
    If f Is Nothing Then
        msg = "Fee code " & feeCode & " is missing in tblFees."
        Exit Function
    End If
    fee = FeeAmount(f, amount)
    tax = TaxOn(f, fee)
    Compute = True
End Function

' Commercial tax on any taxable fee amount (loan processing fee, early payoff fee). 0 when Stage 2 tax rates are not set up.
Public Function CommercialTax(ByVal fee As Variant) As Variant
    Dim r As Variant
    CommercialTax = CDec(0)
    If fee <= 0 Then Exit Function
    If repoTaxRate.RateOn(modConstants.TAX_COMMERCIAL, repoConfig.BusinessDate(), r) Then CommercialTax = modMoney.RoundCcy(fee * r, modConstants.CCY_MMK)
End Function
