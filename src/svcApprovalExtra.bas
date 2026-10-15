Attribute VB_Name = "svcApprovalExtra"
Option Explicit

' Approval request types of Stages 2 to 8: which action runs when the approval is approved, and what happens when it is rejected.

Public Function DispatchExtra(ByVal uow As clsUnitOfWork, ByVal req As clsApproval, ByVal args As Object, _
                              ByVal createdBy As String, ByVal approvedBy As String) As clsResult
    Select Case req.RequestType
        Case modConstants.REQ_PAY_BILLS
            Set DispatchExtra = svcCosts.ExecPayBills(uow, args, createdBy, approvedBy)
        Case modConstants.REQ_EXPENSE
            Set DispatchExtra = svcExpenses.ExecExpense(uow, args, createdBy, approvedBy)
        Case modConstants.REQ_BUY_ASSET
            Set DispatchExtra = svcFixedAssets.ExecBuy(uow, args, createdBy, approvedBy)
        Case modConstants.REQ_DISPOSE_ASSET
            Set DispatchExtra = svcFixedAssets.ExecDispose(uow, args, createdBy, approvedBy)
        Case modConstants.REQ_PAY_TAXES
            Set DispatchExtra = svcTaxes.ExecPayTaxes(uow, args, createdBy, approvedBy)
        Case modConstants.REQ_LOAN_APPROVAL
            Set DispatchExtra = svcLoans.ExecApprove(uow, args, createdBy, approvedBy)
        Case modConstants.REQ_WRITEOFF
            Set DispatchExtra = svcLoanPayments.ExecWriteOff(uow, args, createdBy, approvedBy)
        Case modConstants.REQ_APPRAISAL
            If modArgs.Text(args, "flow") = "register" Then
                Set DispatchExtra = svcCollateral.ExecRegister(uow, args, createdBy, approvedBy)
            Else
                Set DispatchExtra = svcCollateral.ExecAppraise(uow, args, createdBy, approvedBy)
            End If
        Case modConstants.REQ_REPOSSESSION
            Set DispatchExtra = svcRepossession.ExecRepossess(uow, args, createdBy, approvedBy)
        Case modConstants.REQ_ASSET_SALE
            Set DispatchExtra = svcRepossession.ExecSell(uow, args, createdBy, approvedBy)
        Case modConstants.REQ_REDEEM
            Set DispatchExtra = svcRepossession.ExecRedeem(uow, args, createdBy, approvedBy)
        Case modConstants.REQ_PAYROLL
            Set DispatchExtra = svcPayroll.ExecPostPayroll(uow, args, createdBy, approvedBy)
        Case modConstants.REQ_PAY_PAYROLL
            Set DispatchExtra = svcPayroll.ExecPayPayroll(uow, args, createdBy, approvedBy)
        Case modConstants.REQ_STAFF_ADVANCE
            Set DispatchExtra = svcPayroll.ExecAdvance(uow, args, createdBy, approvedBy)
        Case modConstants.REQ_LARGE_FX, modConstants.REQ_FX_LIMIT
            Set DispatchExtra = svcFxDeals.ExecDeal(uow, args, createdBy, approvedBy)
        Case modConstants.REQ_INTL_PAYMENT
            Set DispatchExtra = svcIntl.ExecSend(uow, args, createdBy, approvedBy)
        Case modConstants.REQ_INVESTMENT
            Set DispatchExtra = svcInvest.ExecApproveInvestment(uow, args, createdBy, approvedBy)
        Case modConstants.REQ_DISPOSE_INVESTMENT
            Set DispatchExtra = svcInvest.ExecDisposeInvestment(uow, args, createdBy, approvedBy)
        Case modConstants.REQ_WRITEOFF_INVESTMENT
            Set DispatchExtra = svcInvest.ExecWriteOffInvestment(uow, args, createdBy, approvedBy)
        Case modConstants.REQ_BUY_BOND
            Set DispatchExtra = svcBonds.ExecBuyBond(uow, args, createdBy, approvedBy)
        Case modConstants.REQ_POLICY_RATE
            Set DispatchExtra = svcPolicy.ExecPolicy(uow, args, createdBy, approvedBy)
        Case modConstants.REQ_FINE
            Set DispatchExtra = svcInspection.ExecFine(uow, args, createdBy, approvedBy)
        Case modConstants.REQ_YEAR_END
            Set DispatchExtra = svcYearEnd.ExecYearEnd(uow, args, createdBy, approvedBy)
        Case modConstants.REQ_REPORT_CASE
            Set DispatchExtra = svcAml.ExecReportCase(uow, args, createdBy, approvedBy)
        Case modConstants.REQ_INCOME_TAX
            Set DispatchExtra = svcTaxReturn.ExecBookIncomeTax(uow, args, createdBy, approvedBy)
        Case Else
            Set DispatchExtra = modMessages.FailResult("Unknown request type " & req.RequestType & ".")
    End Select
End Function

' Called when an approval is rejected (by the approver, or because the re-check failed).
Public Sub OnReject(ByVal uow As clsUnitOfWork, ByVal req As clsApproval, ByVal reason As String)
    If req.RequestType = modConstants.REQ_LOAN_APPROVAL Then svcLoans.MarkRejected uow, modArgs.Text(modArgs.Decode(req.Note), "loan"), reason
    If req.RequestType = modConstants.REQ_INVESTMENT Then svcInvest.MarkInvestRejected uow, modArgs.Text(modArgs.Decode(req.Note), "investment"), reason
    If req.RequestType = modConstants.REQ_FINE Then svcInspection.CancelFine uow, modArgs.Text(modArgs.Decode(req.Note), "fine")
End Sub
