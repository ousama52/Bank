Attribute VB_Name = "svcSecurity"
Option Explicit

' Roles, processing branch rules and amount limits.

Public Const ACT_DEPOSIT As String = "deposit"
Public Const ACT_WITHDRAW As String = "withdraw"
Public Const ACT_TRANSFER As String = "transfer"
Public Const ACT_OPEN As String = "open an account"
Public Const ACT_INTERBANK_OUT As String = "send an interbank transfer"
Public Const ACT_INTERBANK_IN As String = "receive an interbank transfer"
Public Const ACT_RETURN_IB As String = "return an interbank transfer"
Public Const ACT_SHIPMENT As String = "ship cash"
Public Const ACT_APPROVE As String = "approve or reject"
Public Const ACT_REVERSE As String = "reverse"
Public Const ACT_FREEZE As String = "freeze"
Public Const ACT_UNFREEZE As String = "unfreeze"
Public Const ACT_REACTIVATE As String = "reactivate"
Public Const ACT_CLOSE As String = "close an account"
Public Const ACT_BREAKFD As String = "break a fixed deposit"
Public Const ACT_ENDOFDAY As String = "run end of day"
Public Const ACT_SETUP As String = "run setup"
Public Const ACT_REPORTS As String = "run reports"
Public Const ACT_STATEMENT As String = "print a statement"

' Stage 2
Public Const ACT_POSTCOSTS As String = "post branch costs"
Public Const ACT_LOADASSETS As String = "load fixed assets"
Public Const ACT_PAYBILLS As String = "pay bills"
Public Const ACT_EXPENSE As String = "post an expense"
Public Const ACT_BUYASSET As String = "buy an asset"
Public Const ACT_DISPOSE As String = "dispose of an asset"
Public Const ACT_PAYTAXES As String = "pay taxes"
Public Const ACT_BUDGET As String = "create a budget"
' Stage 3
Public Const ACT_LOAN_QUOTE As String = "quote a loan"
Public Const ACT_LOAN_APPLY As String = "apply for a loan"
Public Const ACT_LOAN_DISBURSE As String = "disburse a loan"
Public Const ACT_LOAN_PAY As String = "record a loan payment"
Public Const ACT_LOAN_WRITEOFF As String = "write off a loan"
Public Const ACT_LOAN_RECOVER As String = "record a loan recovery"
Public Const ACT_LOAN_REPORTS As String = "run loan reports"
' Stage 4
Public Const ACT_COLLATERAL As String = "register or appraise collateral"
Public Const ACT_PLEDGE As String = "pledge or release collateral"
Public Const ACT_REVALUE As String = "revalue collateral"
Public Const ACT_NOTICE As String = "issue a default notice"
Public Const ACT_REPOSSESS As String = "repossess an asset"
Public Const ACT_SELL As String = "sell a repossessed asset"
Public Const ACT_ASSETOPS As String = "run a held asset"
Public Const ACT_REDEEM As String = "redeem a repossessed asset"

' Stages 5 to 8
Public Const ACT_HR As String = "change staff records"
Public Const ACT_SAMPLESTAFF As String = "generate sample staff"
Public Const ACT_PAYROLL_PREP As String = "prepare payroll"
Public Const ACT_PAYROLL_PAY As String = "pay payroll"
Public Const ACT_ADVANCE As String = "issue a staff advance"
Public Const ACT_FX_RATES As String = "set exchange rates"
Public Const ACT_FX_DEAL As String = "deal in foreign currency"
Public Const ACT_FX_MARKET As String = "deal with the foreign market"
Public Const ACT_INTL As String = "send an international payment"
Public Const ACT_INTL_IN As String = "receive an international payment"
Public Const ACT_INTL_RETURN As String = "return an international payment"
Public Const ACT_WALLET As String = "use a wallet"
Public Const ACT_WALLET_ADMIN As String = "manage wallets, merchants and agents"
Public Const ACT_INVEST As String = "manage investments"
Public Const ACT_INVEST_SIM As String = "simulate investees"
Public Const ACT_POLICY As String = "apply the policy rate"
Public Const ACT_BONDS As String = "buy government securities"
Public Const ACT_COMPLIANCE As String = "run the compliance check"
Public Const ACT_AML As String = "review AML alerts"
Public Const ACT_INSPECT As String = "run an inspection"
Public Const ACT_FINE As String = "propose a fine"
Public Const ACT_TAXRETURN As String = "prepare or book income tax"
Public Const ACT_YEAREND As String = "close a year"

Private Function IsOneOf(ByVal role As String, ParamArray roles() As Variant) As Boolean
    Dim i As Long
    For i = LBound(roles) To UBound(roles)
        If role = CStr(roles(i)) Then
            IsOneOf = True
            Exit Function
        End If
    Next i
End Function

Public Function Allowed(ByVal role As String, ByVal action As String) As Boolean
    Select Case action
        Case ACT_DEPOSIT, ACT_WITHDRAW, ACT_TRANSFER, ACT_OPEN, ACT_INTERBANK_OUT
            Allowed = IsOneOf(role, modConstants.ROLE_TELLER, modConstants.ROLE_MANAGER, modConstants.ROLE_HEADOFFICE)
        Case ACT_APPROVE, ACT_FREEZE, ACT_UNFREEZE, ACT_REACTIVATE, ACT_CLOSE, ACT_BREAKFD
            Allowed = IsOneOf(role, modConstants.ROLE_MANAGER, modConstants.ROLE_HEADOFFICE)
        Case ACT_INTERBANK_IN, ACT_RETURN_IB, ACT_SHIPMENT, ACT_REVERSE
            Allowed = (role = modConstants.ROLE_HEADOFFICE)
        Case ACT_REPORTS
            Allowed = IsOneOf(role, modConstants.ROLE_MANAGER, modConstants.ROLE_HEADOFFICE, modConstants.ROLE_AUDITOR, modConstants.ROLE_ADMIN, modConstants.ROLE_REGULATOR)
        Case ACT_STATEMENT
            Allowed = IsOneOf(role, modConstants.ROLE_TELLER, modConstants.ROLE_MANAGER, modConstants.ROLE_HEADOFFICE, modConstants.ROLE_AUDITOR, modConstants.ROLE_ADMIN)
        Case ACT_ENDOFDAY, ACT_SETUP
            Allowed = (role = modConstants.ROLE_ADMIN)
        Case Else
            Allowed = ExtraAllowed(role, action)
    End Select
End Function

' Later stages add their actions here.
Private Function ExtraAllowed(ByVal role As String, ByVal action As String) As Boolean
    Select Case action
        Case ACT_POSTCOSTS, ACT_LOADASSETS
            ExtraAllowed = (role = modConstants.ROLE_ADMIN)
        Case ACT_PAYBILLS, ACT_BUYASSET, ACT_DISPOSE, ACT_PAYTAXES, ACT_BUDGET, ACT_LOAN_WRITEOFF, ACT_REPOSSESS, ACT_SELL, ACT_REDEEM
            ExtraAllowed = (role = modConstants.ROLE_HEADOFFICE)
        Case ACT_REVALUE
            ExtraAllowed = IsOneOf(role, modConstants.ROLE_HEADOFFICE, modConstants.ROLE_ADMIN)
        Case ACT_EXPENSE, ACT_COLLATERAL, ACT_NOTICE, ACT_ASSETOPS
            ExtraAllowed = IsOneOf(role, modConstants.ROLE_MANAGER, modConstants.ROLE_HEADOFFICE)
        Case ACT_LOAN_QUOTE, ACT_LOAN_APPLY, ACT_LOAN_DISBURSE, ACT_LOAN_RECOVER, ACT_PLEDGE
            ExtraAllowed = IsOneOf(role, modConstants.ROLE_LOANOFFICER, modConstants.ROLE_MANAGER, modConstants.ROLE_HEADOFFICE)
        Case ACT_LOAN_PAY
            ExtraAllowed = IsOneOf(role, modConstants.ROLE_TELLER, modConstants.ROLE_LOANOFFICER, modConstants.ROLE_MANAGER, modConstants.ROLE_HEADOFFICE)
        Case ACT_LOAN_REPORTS
            ExtraAllowed = IsOneOf(role, modConstants.ROLE_LOANOFFICER, modConstants.ROLE_MANAGER, modConstants.ROLE_HEADOFFICE, modConstants.ROLE_AUDITOR, modConstants.ROLE_ADMIN)
        Case ACT_HR, ACT_PAYROLL_PREP, ACT_FX_RATES, ACT_WALLET_ADMIN
            ExtraAllowed = IsOneOf(role, modConstants.ROLE_HEADOFFICE, modConstants.ROLE_ADMIN)
        Case ACT_SAMPLESTAFF, ACT_INVEST_SIM, ACT_YEAREND
            ExtraAllowed = (role = modConstants.ROLE_ADMIN)
        Case ACT_PAYROLL_PAY, ACT_FX_MARKET, ACT_INTL_IN, ACT_INTL_RETURN, ACT_INVEST, ACT_POLICY, ACT_BONDS, ACT_AML, ACT_TAXRETURN
            ExtraAllowed = (role = modConstants.ROLE_HEADOFFICE)
        Case ACT_ADVANCE
            ExtraAllowed = IsOneOf(role, modConstants.ROLE_MANAGER, modConstants.ROLE_HEADOFFICE)
        Case ACT_FX_DEAL, ACT_INTL, ACT_WALLET
            ExtraAllowed = IsOneOf(role, modConstants.ROLE_TELLER, modConstants.ROLE_MANAGER, modConstants.ROLE_HEADOFFICE)
        Case ACT_COMPLIANCE
            ExtraAllowed = IsOneOf(role, modConstants.ROLE_HEADOFFICE, modConstants.ROLE_ADMIN, modConstants.ROLE_AUDITOR, modConstants.ROLE_REGULATOR)
        Case ACT_INSPECT
            ExtraAllowed = IsOneOf(role, modConstants.ROLE_HEADOFFICE, modConstants.ROLE_REGULATOR)
        Case ACT_FINE
            ExtraAllowed = (role = modConstants.ROLE_REGULATOR)
        Case Else
            ExtraAllowed = False
    End Select
End Function

' "" when the user may do the action, else the refusal text.
Public Function Denied(ByVal user As clsUser, ByVal action As String) As String
    If user Is Nothing Then
        Denied = modMessages.MSG_NO_USER
    ElseIf Not Allowed(user.Role, action) Then
        Denied = modMessages.RoleDenied(user.Role, action)
    End If
End Function

' Teller and BranchManager work only for their own branch. HeadOffice may pick any branch (Control C8).
Public Function ResolveBranch(ByVal uow As clsUnitOfWork, ByVal user As clsUser, ByVal requested As String, _
                              ByRef branch As String, ByRef msg As String) As Boolean
    requested = Trim$(requested)
    If Len(requested) = 0 Then
        branch = user.BranchID
    ElseIf requested = user.BranchID Then
        branch = requested
    ElseIf user.Role = modConstants.ROLE_HEADOFFICE Or user.Role = modConstants.ROLE_ADMIN Then
        branch = requested
    Else
        msg = "Your role can only work for your own branch (" & user.BranchID & ")."
        Exit Function
    End If
    If Not uow.Branches().Exists(branch) Then
        msg = modMessages.UnknownBranch(branch)
        Exit Function
    End If
    If uow.Branches()(branch).Status <> modConstants.STATUS_OPEN Then
        msg = "Branch " & branch & " is not Open."
        Exit Function
    End If
    ResolveBranch = True
End Function

' Teller and BranchManager may only handle accounts of their own branch (HeadOffice: any).
Public Function CanManageAccount(ByVal user As clsUser, ByVal a As clsAccount) As Boolean
    CanManageAccount = (user.Role = modConstants.ROLE_HEADOFFICE) Or (a.BranchID = user.BranchID)
End Function

' Amount above Config.TellerLimit needs approval.
Public Function OverTellerLimit(ByVal amount As Variant) As Boolean
    OverTellerLimit = (amount > repoConfig.GetDec(modConstants.CFG_TELLERLIMIT))
End Function

' Withdrawals, outgoing transfers and interbank-out of one account in one business date may not pass the daily limit.
Public Function OverDailyLimit(ByVal accountNo As String, ByVal businessDate As Date, ByVal amount As Variant, Optional ByVal ccy As String = "MMK") As Boolean
    Dim total As Variant
    total = repoTransaction.DebitTotalForDay(accountNo, businessDate) + amount
    If ccy <> modConstants.CCY_MMK Then total = modMoney.RoundCcy(total * repoFxRate.LatestMid(ccy, businessDate), modConstants.CCY_MMK)
    OverDailyLimit = (total > repoConfig.GetDec(modConstants.CFG_DAILYLIMIT))
End Function

' Request types that only a HeadOffice user may decide.
Private Function IsHeadOfficeOnly(ByVal requestType As String) As Boolean
    Select Case requestType
        Case modConstants.REQ_CASH_SHIPMENT, modConstants.REQ_REVERSAL, modConstants.REQ_PAY_BILLS, modConstants.REQ_BUY_ASSET, modConstants.REQ_DISPOSE_ASSET, _
             modConstants.REQ_PAY_TAXES, modConstants.REQ_WRITEOFF, modConstants.REQ_REPOSSESSION, modConstants.REQ_ASSET_SALE, modConstants.REQ_REDEEM, _
             modConstants.REQ_PAYROLL, modConstants.REQ_PAY_PAYROLL, modConstants.REQ_STAFF_ADVANCE, modConstants.REQ_FX_LIMIT, modConstants.REQ_INVESTMENT, _
             modConstants.REQ_DISPOSE_INVESTMENT, modConstants.REQ_WRITEOFF_INVESTMENT, modConstants.REQ_BUY_BOND, modConstants.REQ_POLICY_RATE, modConstants.REQ_FINE, _
             modConstants.REQ_YEAR_END, modConstants.REQ_REPORT_CASE, modConstants.REQ_INCOME_TAX
            IsHeadOfficeOnly = True
    End Select
End Function

' Maker and checker rules for deciding an approval. "" when allowed, else the reason.
Public Function ApproverProblem(ByVal approver As clsUser, ByVal req As clsApproval) As String
    If approver Is Nothing Then ApproverProblem = modMessages.MSG_NO_USER: Exit Function
    If Not Allowed(approver.Role, ACT_APPROVE) Then ApproverProblem = modMessages.RoleDenied(approver.Role, ACT_APPROVE): Exit Function
    If approver.UserID = req.RequestedBy Then ApproverProblem = modMessages.MSG_MAKER_CHECKER: Exit Function
    If approver.Role = modConstants.ROLE_HEADOFFICE Then Exit Function
    If IsHeadOfficeOnly(req.RequestType) Then
        ApproverProblem = "Only a HeadOffice user can decide a " & req.RequestType & " request."
    ElseIf req.BranchID <> approver.BranchID Then
        ApproverProblem = "A BranchManager can only decide requests of their own branch (" & approver.BranchID & ")."
    ElseIf req.RequestType = modConstants.REQ_LOAN_APPROVAL Then
        If req.Amount > repoConfig.GetDec(modConstants.CFG_LOAN_BRANCH_LIMIT) Then ApproverProblem = "The loan is above the branch limit. Only HeadOffice can approve it."
    ElseIf req.Amount > repoConfig.GetDec(modConstants.CFG_BMLIMIT) Then
        ApproverProblem = "The amount is above the BranchManager limit. Only HeadOffice can decide it."
    End If
End Function
