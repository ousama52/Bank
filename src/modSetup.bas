Attribute VB_Name = "modSetup"
Option Explicit

' RunSetup logic: makes anything that is missing (sheets, tables, dropdowns, Config keys,
' the Control sheet and its buttons). It NEVER clears or changes existing data.
' Column types: T text, I integer, D date, S date+time, R rate, M decimal(4), X general.

Private Const TITLE_TEXT As String = "Golden Delta Bank - SIMULATION - Control"
Private Const BTN_PREFIX As String = "btn_"
Private Const ERR_BASE As Long = vbObjectError + 1100

' Called by modMain.RunSetup. Returns a result; the entry macro shows it.
Public Function SetupAll() As clsResult
    Dim specs As Collection, spec As Variant, made As Long
    Set specs = TableSpecs()
    For Each spec In specs
        made = made + EnsureTable(CStr(spec(0)), CStr(spec(1)), CStr(spec(2)))
    Next spec
    made = made + EnsureReportSheets()
    made = made + EnsureConfigKeys()
    made = made + modSetupStages.UpgradeAll()
    made = made + modSetupLate.UpgradeLate()
    BuildControlSheet
    clsLogger.Info "modSetup.SetupAll", "Setup finished, " & made & " item(s) created."
    Set SetupAll = modMessages.OkResult("Setup finished. " & made & " item(s) created. Existing data was not changed.", made)
End Function

' ---- Table definitions (schema = DATA_TABLES.md, same column order) ----

Private Sub AddSpec(ByVal specs As Collection, ByVal tbl As String, ByVal cols As String, ByVal enums As String)
    specs.Add Array(tbl, cols, enums)
End Sub

Private Function TableSpecs() As Collection
    Dim s As New Collection
    AddSpec s, modConstants.TBL_CONFIG, "Key:T|Value:T", ""
    AddSpec s, modConstants.TBL_USERS, "UserID:T|Name:T|Role:T|BranchID:T|IsActive:T", _
        "Role=Teller,BranchManager,HeadOffice,Auditor,Admin,LoanOfficer,Regulator;IsActive=Yes,No"
    AddSpec s, modConstants.TBL_BRANCHES, "BranchID:T|Name:T|Type:T|City:T|Region:T|PremisesType:T|OpenDate:D|Status:T|VaultMin:I|VaultMax:I", _
        "Type=HeadOffice,City,Town;PremisesType=Owned,Rented;Status=Open,Closed"
    AddSpec s, modConstants.TBL_CUSTOMERS, "CustomerID:T|Name:T|CustomerType:T|Phone:T|Address:T|HomeBranchID:T|KYCLevel:I|TaxExempt:T|JoinDate:D|Status:T", _
        "CustomerType=Individual,Business;TaxExempt=Yes,No;Status=Active,Inactive"
    AddSpec s, modConstants.TBL_PRODUCTS, "ProductCode:T|Name:T|Type:T|Currency:T|MinBalance:I|MinOpening:I|MaintenanceFeeCode:T|IsActive:T", _
        "Type=Savings,Current,Fixed,Wallet;IsActive=Yes,No"
    AddSpec s, modConstants.TBL_RATETIERS, "RateID:T|ProductCode:T|TermMonths:I|FromAmount:I|ToAmount:I|AnnualRate:R|EffectiveFrom:D", ""
    AddSpec s, modConstants.TBL_FEES, "FeeCode:T|Name:T|Method:T|FlatAmount:I|Rate:R|MinFee:I|MaxFee:I|Currency:T|IncomeGL:T", "Method=Flat,Percent"
    AddSpec s, modConstants.TBL_OTHERBANKS, "BankCode:T|Name:T|SettleDays:I|FailRate:R|ReturnRate:R", ""
    AddSpec s, modConstants.TBL_CHART, "GLCode:T|Name:T|Type:T|NormalSide:T|Stage:I|IsActive:T", _
        "Type=Asset,Liability,Equity,Income,Expense;NormalSide=Debit,Credit;IsActive=Yes,No"
    AddSpec s, modConstants.TBL_ACCOUNTS, "AccountNo:T|CustomerID:T|BranchID:T|ProductCode:T|Currency:T|Status:T|OpenDate:D|TermMonths:I|MaturityDate:D|LockedRate:R|PayoutAccountNo:T|Balance:I|AccruedInterest:M|LastAccrualDate:D|LastActivityDate:D|ClosedDate:D", _
        "Status=Active,Dormant,Frozen,Closed"
    AddMoreSpecs s
    Set TableSpecs = s
End Function

Private Sub AddMoreSpecs(ByVal s As Collection)
    AddSpec s, modConstants.TBL_TRANSACTIONS, "TxnID:T|GroupID:T|TxnDate:D|AccountNo:T|TxnType:T|Direction:T|Amount:I|BalanceAfter:I|CounterAccountNo:T|BranchID:T|Description:T|Reference:T|JournalID:T|Status:T|CreatedBy:T|ApprovedBy:T", _
        "TxnType=Opening,Deposit,Withdrawal,TransferOut,TransferIn,InterbankOut,InterbankIn,InterbankRefund,InterbankReturn,Fee,FeeRefund,InterestCredit,Reversal,FDMaturity,FDBreak,Tax,LoanDisbursement,LoanFee,LoanRepayment,SurplusPaid,SalaryCredit,FxBuy,FxSell,FxConvert,IntlOut,IntlIn,IntlRefund,IntlReturn,WalletTopUp,WalletCashOut,WalletPay,WalletTransferOut,WalletTransferIn,AgentCashIn,AgentCashOut,AgentFloatIn,AgentFloatOut,MerchantReceipt;Direction=Debit,Credit;Status=Posted,Reversed"
    AddSpec s, modConstants.TBL_INTERBANK, "TransferID:T|Direction:T|AccountNo:T|OtherBankCode:T|OtherBankAccount:T|Amount:I|Fee:I|SubmittedDate:D|ExpectedSettleDate:D|Status:T|StatusDate:D|Reason:T|SubmitJournalID:T|SettleJournalID:T|ReturnJournalID:T", _
        "Direction=Out,In;Status=Submitted,Settled,Failed,Returned"
    AddSpec s, modConstants.TBL_APPROVALS, "ApprovalID:T|RequestType:T|AccountNo:T|CounterAccountNo:T|Amount:I|BranchID:T|Note:T|RequestedBy:T|RequestedAt:S|Status:T|ApprovedBy:T|DecidedAt:S|DecisionNote:T|TxnGroupID:T", _
        "RequestType=LargeDeposit,LargeWithdrawal,LargeTransfer,LargeInterbank,CashShipment,Reversal,BreakFixedDeposit,CloseAccount,Unfreeze,Reactivate,PayBills,Expense,BuyAsset,DisposeAsset,PayTaxes,LoanApproval,WriteOffLoan,Appraisal,Repossession,AssetSale,Redeem,Payroll,PayPayroll,StaffAdvance,LargeFx,FxLimit,IntlPayment,Investment,DisposeInvestment,WriteOffInvestment,BuyBond,PolicyRate,Fine,YearEndClose,ReportCase,IncomeTax;Status=Pending,Approved,Rejected"
    AddSpec s, modConstants.TBL_JOURNAL, "JournalID:T|PostDate:D|Description:T|JournalType:T|GroupID:T|Status:T|ReversalOf:T|CreatedBy:T|CreatedAt:S", "Status=Posted,Reversed"
    AddSpec s, modConstants.TBL_JLINES, "LineID:T|JournalID:T|BranchID:T|GLCode:T|AccountNo:T|Currency:T|Debit:I|Credit:I|FxRate:X|ForeignAmount:X", ""
    AddSpec s, modConstants.TBL_DAILYCLOSE, "BusinessDate:D|RunAt:S|AccountsAccrued:I|InterestAccrued:M|InterbankSettled:I|FDMatured:I|DormantMarked:I|ChecksPassed:T|Notes:T", "ChecksPassed=Yes,No"
    AddSpec s, modConstants.TBL_AUDIT, "AuditID:T|Timestamp:S|UserID:T|Action:T|Target:T|Detail:T", ""
    AddSpec s, modConstants.TBL_LOG, "Timestamp:S|Level:T|Procedure:T|Message:T", "Level=Info,Warning,Error"
End Sub

Public Function FormatFor(ByVal code As String) As String
    Select Case code
        Case "T": FormatFor = "@"
        Case "I": FormatFor = "#,##0"
        Case "D": FormatFor = "yyyy-mm-dd"
        Case "S": FormatFor = "yyyy-mm-dd hh:mm:ss"
        Case "R": FormatFor = "0.00%"
        Case "M": FormatFor = "#,##0.0000"
        Case Else: FormatFor = "General"
    End Select
End Function

' ---- Tables ----

' Returns 1 when the table had to be created, else 0. Dropdowns are checked either way.
Public Function EnsureTable(ByVal tbl As String, ByVal cols As String, ByVal enums As String) As Long
    If modSheetUtils.TableExists(tbl) Then
        CheckHeaders tbl, cols
    Else
        CreateTable tbl, cols
        EnsureTable = 1
    End If
    ApplyEnums tbl, enums
End Function

Private Sub CreateTable(ByVal tbl As String, ByVal cols As String)
    Dim ws As Worksheet, lo As ListObject, parts() As String, pair() As String
    Dim hdr() As Variant, n As Long, i As Long
    Set ws = modSheetUtils.EnsureSheet(modSheetUtils.SheetNameOf(tbl))
    ' Never write headers over a sheet that already holds something else.
    If Application.WorksheetFunction.CountA(ws.Cells) > 0 Then
        Err.Raise ERR_BASE + 1, "modSetup.CreateTable", "Sheet " & ws.Name & " has content but no table " & tbl & ". Fix it by hand first."
    End If
    parts = Split(cols, "|")
    n = UBound(parts) + 1
    ReDim hdr(1 To 1, 1 To n)
    For i = 1 To n
        pair = Split(parts(i - 1), ":")
        hdr(1, i) = pair(0)
        ws.Columns(i).NumberFormat = FormatFor(pair(1))
    Next i
    ws.Range(ws.Cells(1, 1), ws.Cells(1, n)).NumberFormat = "@"
    ws.Range(ws.Cells(1, 1), ws.Cells(1, n)).Value = hdr
    Set lo = ws.ListObjects.Add(xlSrcRange, ws.Range(ws.Cells(1, 1), ws.Cells(2, n)), , xlYes)
    lo.Name = tbl
    lo.TableStyle = "TableStyleMedium2"
    ws.Range(ws.Columns(1), ws.Columns(n)).AutoFit
End Sub

' An existing table is never changed, but a missing column is reported in the Log.
Private Sub CheckHeaders(ByVal tbl As String, ByVal cols As String)
    Dim map As Object, parts() As String, i As Long, nm As String
    Set map = modSheetUtils.HeaderMap(tbl)
    parts = Split(cols, "|")
    For i = 0 To UBound(parts)
        nm = Split(parts(i), ":")(0)
        If Not map.Exists(nm) Then
            clsLogger.Warn "modSetup.CheckHeaders", "Table " & tbl & " has no column " & nm & "."
        End If
    Next i
End Sub

Private Function HasValidation(ByVal rng As Range) As Boolean
    Dim t As Long
    On Error Resume Next
    t = rng.Cells(1, 1).Validation.Type
    HasValidation = (Err.Number = 0)
    Err.Clear
    On Error GoTo 0
End Function

Private Sub SetList(ByVal rng As Range, ByVal formula As String)
    With rng.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:=formula
        .IgnoreBlank = True
        .InCellDropdown = True
    End With
End Sub

' A dropdown list longer than 255 characters cannot sit in the validation formula, so it is written to the hidden sheet Lists.
Private Sub SetListFromSheet(ByVal rng As Range, ByVal listKey As String, ByVal items As String)
    Dim ws As Worksheet, parts() As String, c As Long, n As Long, i As Long, arr() As Variant, found As Range
    Set ws = modSheetUtils.EnsureSheet("Lists")
    Set found = ws.Rows(1).Find(What:=listKey, LookIn:=xlValues, LookAt:=xlWhole)
    If found Is Nothing Then
        If Len(CStr(ws.Cells(1, 1).Value)) = 0 Then
            c = 1
        Else
            c = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column + 1
        End If
        ws.Cells(1, c).Value = listKey
    Else
        c = found.Column
    End If
    parts = Split(items, ",")
    n = UBound(parts) + 1
    ReDim arr(1 To n, 1 To 1)
    For i = 1 To n
        arr(i, 1) = parts(i - 1)
    Next i
    ws.Range(ws.Cells(2, c), ws.Cells(1 + n, c)).Value = arr
    ws.Visible = xlSheetHidden
    SetList rng, "=Lists!" & ws.Range(ws.Cells(2, c), ws.Cells(1 + n, c)).Address
End Sub

' enums = "Column=a,b,c;Column2=x,y". The dropdown is (re)written every time so new values from later stages appear.
Private Sub ApplyEnums(ByVal tbl As String, ByVal enums As String)
    Dim lo As ListObject, item As Variant, pair() As String, rng As Range
    If Len(enums) = 0 Then Exit Sub
    Set lo = modSheetUtils.GetTable(tbl)
    If lo.DataBodyRange Is Nothing Then Exit Sub
    For Each item In Split(enums, ";")
        pair = Split(CStr(item), "=")
        Set rng = Nothing
        On Error Resume Next
        Set rng = lo.ListColumns(pair(0)).DataBodyRange
        On Error GoTo 0
        If Not rng Is Nothing Then
            If Len(pair(1)) > 250 Then SetListFromSheet rng, tbl & "." & pair(0), pair(1) Else SetList rng, pair(1)
        End If
    Next item
End Sub

Private Function EnsureReportSheets() As Long
    Dim nm As Variant
    For Each nm In Array(modConstants.SHEET_TRIAL, modConstants.SHEET_BRANCHSUM, modConstants.SHEET_CASHPOS, _
                         modConstants.SHEET_DEPSUM, modConstants.SHEET_CHECKS, modConstants.SHEET_STATEMENT)
        If Not modSheetUtils.SheetExists(CStr(nm)) Then
            modSheetUtils.EnsureSheet CStr(nm)
            EnsureReportSheets = EnsureReportSheets + 1
        End If
    Next nm
End Function

' ---- Config keys ----

Private Function ConfigDefaults() As Variant
    ConfigDefaults = Array( _
        modConstants.CFG_BANKNAME & "=Golden Delta Bank (SIMULATION)", _
        modConstants.CFG_BASECCY & "=MMK", _
        modConstants.CFG_BUSDATE & "=2026-10-13", _
        modConstants.CFG_POLICYRATE & "=0.10", _
        modConstants.CFG_WHTRATE & "=0.10", _
        modConstants.CFG_TELLERLIMIT & "=5000000", _
        modConstants.CFG_BMLIMIT & "=50000000", _
        modConstants.CFG_DAILYLIMIT & "=20000000", _
        modConstants.CFG_DORMANTDAYS & "=365", _
        modConstants.CFG_EARLYBREAK & "=0.03", _
        modConstants.CFG_DAYCOUNT & "=365", _
        modConstants.CFG_SHIPROUND & "=10000000", _
        modConstants.CFG_RESERVE & "=0.05")
End Function

' Adds only the keys that are missing. Existing values are never touched.
Private Function EnsureConfigKeys() As Long
    Dim have As Object, keys As Variant, defs As Variant, i As Long, p As Long
    Dim missing As New Collection, outRows() As Variant, k As Long
    Set have = CreateObject("Scripting.Dictionary")
    have.CompareMode = 1
    keys = modSheetUtils.ReadColumn(modConstants.TBL_CONFIG, "Key")
    If Not IsEmpty(keys) Then
        For i = 1 To UBound(keys, 1)
            If Len(CStr(keys(i, 1))) > 0 Then have(CStr(keys(i, 1))) = True
        Next i
    End If
    defs = ConfigDefaults()
    For i = LBound(defs) To UBound(defs)
        p = InStr(defs(i), "=")
        If Not have.Exists(Left$(defs(i), p - 1)) Then missing.Add defs(i)
    Next i
    If missing.Count = 0 Then Exit Function
    ReDim outRows(1 To missing.Count, 1 To 2)
    For k = 1 To missing.Count
        p = InStr(missing(k), "=")
        outRows(k, 1) = Left$(missing(k), p - 1)
        outRows(k, 2) = Mid$(missing(k), p + 1)
    Next k
    modSheetUtils.AppendRows modConstants.TBL_CONFIG, outRows
    EnsureConfigKeys = missing.Count
End Function

' Setup only needs the business date to show it; the full config reader is repoConfig.
Private Function ConfigDate(ByVal key As String, ByRef result As Date) As Boolean
    Dim data As Variant, i As Long
    data = modSheetUtils.ReadTable(modConstants.TBL_CONFIG)
    If IsEmpty(data) Then Exit Function
    For i = 1 To UBound(data, 1)
        If StrComp(CStr(data(i, 1)), key, vbTextCompare) = 0 Then
            ConfigDate = modDateUtils.TryToDate(data(i, 2), result)
            Exit Function
        End If
    Next i
End Function

' ---- Control sheet ----

Private Sub BuildControlSheet()
    Dim ws As Worksheet, d As Date
    If modSheetUtils.SheetExists(modConstants.SHEET_CONTROL) Then
        Set ws = modSheetUtils.GetSheet(modConstants.SHEET_CONTROL)
    Else
        Set ws = ThisWorkbook.Worksheets.Add(Before:=ThisWorkbook.Worksheets(1))
        ws.Name = modConstants.SHEET_CONTROL
    End If
    WriteLabels ws
    FormatInputs ws
    AddInputLists ws
    If ConfigDate(modConstants.CFG_BUSDATE, d) Then ws.Range(modConstants.CTL_BUSDATE).Value = d
    LayoutButtons ws
End Sub

Private Sub WriteLabels(ByVal ws As Worksheet)
    Dim labels As Variant, i As Long
    labels = Array("Business date (set by the macros)", "Inputs", "Current user ID", "Account number", _
        "Counter account / other bank account", "Amount (MMK)", "Processing branch (blank = my branch)", _
        "Other bank code", "Term months (fixed deposit)", "Payout account (fixed deposit)", _
        "Group / approval / transfer ID", "Customer ID (open account)", "Product code (open account)", _
        "Reference / note", "Statement from", "Statement to", "Last status message", "GL code (expense)", "Vendor, tenant or buyer", "Description", _
        "Asset or collateral type code", "Useful life (months)", "Residual value or selling costs", "Period YYYYMM (or year YYYY)", "Loan product code", "Mode: Account or Cash", _
        "Loan ID or repossession ID", "Collateral ID", _
        "Employee ID", "Position code", "Grade (1 to 3)", "Effective date", "Payroll run ID", "Currency code", "Mid rate (set rates)", "Merchant ID", "Agent ID", _
        "Beneficiary or sender name", "Foreign bank code", "Beneficiary account (text)", "Purpose", "(not used)", "Investee ID", "Investment ID", "Investment type", _
        "Rate or share percent", "Reported profit or new valuation", "Alert ID", "Bond type (TBill or GovBond)", "Face value", "Term in days", "Rate (0.08 or 8)", "Reason")
    ws.Range("B1").Value = TITLE_TEXT
    ws.Range("B1").Font.Bold = True
    ws.Range("B1").Font.Size = 14
    For i = 0 To UBound(labels)
        ws.Cells(2 + i, 2).Value = labels(i)
    Next i
    ws.Range("B2:B54").Font.Bold = True
    ws.Columns(1).ColumnWidth = 2
    ws.Columns(2).ColumnWidth = 38
    ws.Columns(3).ColumnWidth = 34
    ws.Columns(4).ColumnWidth = 3
    ws.Range("B2:B54").EntireRow.RowHeight = 21
End Sub

Private Sub FormatInputs(ByVal ws As Worksheet)
    Dim addr As Variant
    ' IDs and account numbers must stay text so leading zeros survive.
    For Each addr In Array(modConstants.CTL_USER, modConstants.CTL_ACCOUNT, modConstants.CTL_COUNTER, modConstants.CTL_BRANCH, _
            modConstants.CTL_BANKCODE, modConstants.CTL_PAYOUT, modConstants.CTL_REFID, modConstants.CTL_CUSTOMER, _
            modConstants.CTL_PRODUCT, modConstants.CTL_NOTE, modConstants.CTL_STATUS, modConstants.CTL_GL, modConstants.CTL_VENDOR, modConstants.CTL_DESC, _
            modConstants.CTL_ASSETTYPE, modConstants.CTL_PERIOD, modConstants.CTL_LOANPRODUCT, modConstants.CTL_MODE, modConstants.CTL_LOANID, modConstants.CTL_COLLATERAL, _
            modConstants.CTL_EMPLOYEE, modConstants.CTL_POSITION, modConstants.CTL_GRADE, modConstants.CTL_PAYRUN, modConstants.CTL_CURRENCY, modConstants.CTL_MERCHANT, modConstants.CTL_AGENT, _
            modConstants.CTL_PARTYNAME, modConstants.CTL_FOREIGNBANK, modConstants.CTL_PARTYACCT, modConstants.CTL_PURPOSE, modConstants.CTL_INVESTEE, modConstants.CTL_INVESTMENT, _
            modConstants.CTL_INVTYPE, modConstants.CTL_ALERT, modConstants.CTL_BONDTYPE, modConstants.CTL_REASON)
        ws.Range(CStr(addr)).NumberFormat = "@"
    Next addr
    ws.Range(modConstants.CTL_EFFDATE).NumberFormat = "yyyy-mm-dd"
    ws.Range(modConstants.CTL_MIDRATE).NumberFormat = "#,##0.0000"
    ws.Range(modConstants.CTL_RATESHARE).NumberFormat = "0.00"
    ws.Range(modConstants.CTL_PROFIT).NumberFormat = "#,##0"
    ws.Range(modConstants.CTL_FACE).NumberFormat = "#,##0"
    ws.Range(modConstants.CTL_TERMDAYS).NumberFormat = "0"
    ws.Range(modConstants.CTL_BONDRATE).NumberFormat = "0.00"
    ws.Range(modConstants.CTL_AMOUNT).NumberFormat = "#,##0"
    ws.Range(modConstants.CTL_TERM).NumberFormat = "0"
    ws.Range(modConstants.CTL_BUSDATE).NumberFormat = "yyyy-mm-dd"
    ws.Range(modConstants.CTL_FROM).NumberFormat = "yyyy-mm-dd"
    ws.Range(modConstants.CTL_TO).NumberFormat = "yyyy-mm-dd"
    ws.Range("C4:C17,C19:C42,C44:C54").Interior.Color = RGB(255, 249, 214)
    ws.Range("C2").Interior.Color = RGB(226, 232, 240)
    ws.Range("C18").Interior.Color = RGB(226, 232, 240)
    ws.Range("C2,C4:C54").Borders.LineStyle = xlContinuous
    ws.Range(modConstants.CTL_LIFE).NumberFormat = "0"
    ws.Range(modConstants.CTL_RESIDUAL).NumberFormat = "#,##0"
    ws.Range("C2").Font.Bold = True
    ws.Range("C2:C54").HorizontalAlignment = xlLeft
    ws.Range(modConstants.CTL_STATUS).WrapText = True
End Sub

Private Sub AddInputLists(ByVal ws As Worksheet)
    SetList ws.Range(modConstants.CTL_USER), "=INDIRECT(""" & modConstants.TBL_USERS & "[UserID]"")"
    SetList ws.Range(modConstants.CTL_BRANCH), "=INDIRECT(""" & modConstants.TBL_BRANCHES & "[BranchID]"")"
    SetList ws.Range(modConstants.CTL_BANKCODE), "=INDIRECT(""" & modConstants.TBL_OTHERBANKS & "[BankCode]"")"
    SetList ws.Range(modConstants.CTL_PRODUCT), "=INDIRECT(""" & modConstants.TBL_PRODUCTS & "[ProductCode]"")"
    SetList ws.Range(modConstants.CTL_TERM), "3,6,12,24"
    SetList ws.Range(modConstants.CTL_MODE), "Account,Cash"
    SetList ws.Range(modConstants.CTL_GRADE), "1,2,3"
    SetList ws.Range(modConstants.CTL_PURPOSE), "Trade,Family,Education,Services,Investment,Other"
    SetList ws.Range(modConstants.CTL_INVTYPE), "Equity,ProfitShare,FixedReturnNote,TradeFinance"
    SetList ws.Range(modConstants.CTL_BONDTYPE), "TBill,GovBond"
    If modSheetUtils.TableExists(modConstants.TBL_POSITIONS) Then SetList ws.Range(modConstants.CTL_POSITION), "=INDIRECT(""" & modConstants.TBL_POSITIONS & "[PositionCode]"")"
    If modSheetUtils.TableExists(modConstants.TBL_FOREIGNBANKS) Then SetList ws.Range(modConstants.CTL_FOREIGNBANK), "=INDIRECT(""" & modConstants.TBL_FOREIGNBANKS & "[BankCode]"")"
    If modSheetUtils.TableExists(modConstants.TBL_LOANPRODUCTS) Then SetList ws.Range(modConstants.CTL_LOANPRODUCT), "=INDIRECT(""" & modConstants.TBL_LOANPRODUCTS & "[ProductCode]"")"
End Sub

' ---- Buttons (form controls that call the entry macros in modMain) ----

Private Sub LayoutButtons(ByVal ws As Worksheet)
    Dim items As Variant, i As Long, parts() As String, anchor As Range
    items = Array("RunSetup|Setup", "RunOpenAccount|Open account", "RunOpenFixedDeposit|Open fixed deposit", _
        "RunDeposit|Deposit", "RunWithdraw|Withdraw", "RunTransfer|Transfer", _
        "RunInterbankOut|Interbank out", "RunInterbankIn|Interbank in", "RunReturnInterbank|Return interbank", _
        "RunCashShipment|Cash shipment", "RunApprove|Approve", "RunReject|Reject", _
        "RunReverse|Reverse", "RunFreeze|Freeze", "RunUnfreeze|Unfreeze", _
        "RunReactivate|Reactivate", "RunCloseAccount|Close account", "RunBreakFixedDeposit|Break fixed deposit", _
        "RunEndOfDay|End of day", "RunReports|Reports", "RunStatement|Statement", _
        "RunLoadFixedAssets|S2 Load fixed assets", "RunPostCosts|S2 Post costs (this month)", "RunPostCostsForMonth|S2 Post costs for period", "RunPayBills|S2 Pay bills", _
        "RunPostExpense|S2 Post expense", "RunBuyAsset|S2 Buy asset", "RunDisposeAsset|S2 Dispose asset", "RunPayTaxes|S2 Pay taxes", "RunCreateBudget|S2 Create budget", _
        "RunQuoteLoan|S3 Quote loan", "RunApplyLoan|S3 Apply for loan", "RunDisburseLoan|S3 Disburse loan", "RunLoanPayment|S3 Loan payment", "RunPayoffLoan|S3 Pay off loan", _
        "RunWriteOffLoan|S3 Write off loan", "RunRecoverLoan|S3 Recover loan", "RunLoanReports|S3 Loan reports", _
        "RunRegisterCollateral|S4 Register collateral", "RunAppraise|S4 Appraise", "RunPledgeCollateral|S4 Pledge", "RunReleaseCollateral|S4 Release", "RunRevalueCollateral|S4 Revalue", _
        "RunIssueDefaultNotice|S4 Default notice", "RunRepossess|S4 Repossess", "RunSellAsset|S4 Sell asset", "RunStartRental|S4 Start rental", "RunEndRental|S4 End rental", _
        "RunRecordAssetOp|S4 Record asset income/cost", "RunUseInBusiness|S4 Use in business", "RunRedeemAsset|S4 Redeem asset", _
        "RunGenerateSampleStaff|S5 Generate sample staff", "RunHireEmployee|S5 Hire", "RunTransferEmployee|S5 Transfer employee", "RunChangePosition|S5 Change position", "RunChangeGrade|S5 Change grade", _
        "RunResignEmployee|S5 Resign", "RunTerminateEmployee|S5 Terminate", "RunIssueAdvance|S5 Staff advance", "RunPreparePayroll|S5 Prepare payroll", "RunPayPayroll|S5 Pay payroll", _
        "RunSetRates|S6 Set rates", "RunBuyCurrency|S6 Buy currency", "RunSellCurrency|S6 Sell currency", "RunConvertCurrency|S6 Convert currency", "RunMarketBuy|S6 Market buy", "RunMarketSell|S6 Market sell", _
        "RunSendInternational|S6 Send international", "RunReceiveInternational|S6 Receive international", "RunReturnInternational|S6 Return international", "RunOpenWallet|S6 Open wallet", _
        "RunWalletTopUp|S6 Wallet top-up", "RunWalletCashOut|S6 Wallet cash-out", "RunWalletTransfer|S6 Wallet transfer", "RunWalletPay|S6 Wallet pay", "RunAgentFloatIn|S6 Agent float in", _
        "RunAgentFloatOut|S6 Agent float out", "RunAgentCashIn|S6 Agent cash-in", "RunAgentCashOut|S6 Agent cash-out", "RunAddMerchant|S6 Add merchant", "RunAddAgent|S6 Add agent", _
        "RunAddInvestee|S7 Add investee", "RunProposeInvestment|S7 Propose investment", "RunFundInvestment|S7 Fund investment", "RunReceiveInvestmentReturn|S7 Receive return", "RunRecordProfitShare|S7 Profit share", _
        "RunValueInvestment|S7 Value investment", "RunDisposeInvestment|S7 Dispose investment", "RunWriteOffInvestment|S7 Write off investment", "RunSimulateInvesteeMonth|S7 Simulate investees", _
        "RunApplyPolicyRate|S8 Apply policy rate", "RunBuyGovSecurity|S8 Buy gov security", "RunComplianceCheck|S8 Compliance check", "RunReviewAlert|S8 Review AML alert", "RunOpenCase|S8 Open case", _
        "RunCloseCase|S8 Close case", "RunReportCase|S8 Report case", "RunInspection|S8 Inspection", "RunIssueFine|S8 Issue fine", "RunTaxReturn|S8 Tax return", "RunBookIncomeTax|S8 Book income tax", _
        "RunYearEndClose|S8 Year-end close")
    ws.Columns(5).ColumnWidth = 26
    ws.Columns(6).ColumnWidth = 26
    ws.Columns(7).ColumnWidth = 26
    ws.Range("E2").Value = "Actions"
    ws.Range("E2").Font.Bold = True
    For i = 0 To UBound(items)
        parts = Split(CStr(items(i)), "|")
        Set anchor = ws.Cells(4 + i \ 3, 5 + (i Mod 3))
        PlaceButton ws, parts(0), parts(1), anchor
    Next i
End Sub

' Re-running setup moves and relabels the same buttons instead of adding duplicates.
Private Sub PlaceButton(ByVal ws As Worksheet, ByVal macroName As String, ByVal caption As String, ByVal anchor As Range)
    Dim btn As Button
    On Error Resume Next
    Set btn = ws.Buttons(BTN_PREFIX & macroName)
    On Error GoTo 0
    If btn Is Nothing Then
        Set btn = ws.Buttons.Add(anchor.Left + 2, anchor.Top + 1, anchor.Width - 4, anchor.Height - 2)
        btn.Name = BTN_PREFIX & macroName
    Else
        btn.Left = anchor.Left + 2
        btn.Top = anchor.Top + 1
        btn.Width = anchor.Width - 4
        btn.Height = anchor.Height - 2
    End If
    btn.Caption = caption
    btn.OnAction = macroName
End Sub
