# DATA_TABLES.md - Golden Delta Bank (simulation)

Each table is an Excel Table (ListObject) on its own worksheet, starting at A1 (headers in row 1).
Sheet name = table name without the `tbl` prefix. Example: sheet `Accounts` holds table `tblAccounts`.
Column names are the exact header text. Keep this order. `BankStarter.xlsx` already follows this file.

Conventions
- Money is whole MMK (Decimal in code). Rates are stored as fractions (0.035 = 3.5%).
- IDs, account numbers, phone numbers, GL codes and codes are Text (format `@`). Dates are real Excel dates. `RunAt`, `CreatedAt`, `RequestedAt`, `DecidedAt`, `Timestamp` are date and time.
- A table that has no rows holds ONE blank row (an Excel Table needs one). A row whose key column is empty is no data.
- Enum columns have dropdown validation created by `RunSetup`.
- "Typed" tables are master data I type by hand. "Macro" tables are written only by macros.

## Relationships

```
Branches 1--* Customers (HomeBranchID)      Customers 1--* Accounts      Products 1--* Accounts
Accounts 1--* Transactions                  Journal 1--* JournalLines    Journal 1--* Transactions
Accounts 1--* InterbankTransfers            OtherBanks 1--* InterbankTransfers
Chart 1--* JournalLines                     Users --> Branches
```

---

## Stage 1 tables

### tblConfig (Typed) - Key/Value, Key unique. Values are stored as text.

| Key | Sample | Meaning |
|---|---|---|
| BankName | `Golden Delta Bank (SIMULATION)` | Title |
| BaseCurrency | `MMK` | Base currency |
| BusinessDate | `2026-10-13` | The current business date (written by end of day) |
| PolicyRate | `0.10` | Reserved for Stage 8 |
| WithholdingTaxRate | `0.10` | Tax on interest paid to customers |
| TellerLimit | `5000000` | Above this a transaction needs approval |
| BranchManagerLimit | `50000000` | Above this only HeadOffice can approve |
| DailyWithdrawalLimit | `20000000` | Per account per business date |
| DormantDays | `365` | Days without customer activity before Dormant |
| EarlyBreakRate | `0.03` | Annual interest rate on early break of a fixed deposit |
| DayCountBasis | `365` | Interest day count |
| CashShipmentRounding | `10000000` | Block size when head office buys cash from treasury |
| ReserveRatio | `0.05` | Reserved for Stage 3 and 8 |

### tblUsers (Typed)

| Column | Type | Rules |
|---|---|---|
| UserID | Text | Key. `U001` |
| Name | Text | Required |
| Role | Enum | `Teller`, `BranchManager`, `HeadOffice`, `Auditor`, `Admin` |
| BranchID | Text | Must exist in tblBranches |
| IsActive | Enum | `Yes`, `No` |

### tblBranches (Typed)

| Column | Type | Rules |
|---|---|---|
| BranchID | Text | Key. `BR001`. `BR001` is the head office |
| Name | Text | Required |
| Type | Enum | `HeadOffice`, `City`, `Town` |
| City | Text | Required |
| Region | Text | Required |
| PremisesType | Enum | `Owned`, `Rented` |
| OpenDate | Date | Required |
| Status | Enum | `Open`, `Closed` |
| VaultMin | Int | Minimum cash wanted in the vault |
| VaultMax | Int | Maximum cash wanted in the vault |

### tblCustomers (Typed)

| Column | Type | Rules |
|---|---|---|
| CustomerID | Text | Key. `C000001` |
| Name | Text | Required |
| CustomerType | Enum | `Individual`, `Business` |
| Phone | Text | Required |
| Address | Text | Required |
| HomeBranchID | Text | Must exist in tblBranches. All accounts of the customer sit in this branch |
| KYCLevel | Int | 1 to 3 (used for wallet limits in Stage 6) |
| TaxExempt | Enum | `Yes`, `No`. `Yes` = no withholding tax on interest |
| JoinDate | Date | Required |
| Status | Enum | `Active`, `Inactive` |

### tblProducts (Typed)

| Column | Type | Rules |
|---|---|---|
| ProductCode | Text | Key. `SAV01`, `SAV02`, `CUR01`, `CUR02`, `FD01` |
| Name | Text | Required |
| Type | Enum | `Savings`, `Current`, `Fixed` |
| Currency | Text | `MMK` in Stage 1 |
| MinBalance | Int | Balance after a debit must be at least this. 0 for fixed |
| MinOpening | Int | Minimum opening deposit (principal for fixed) |
| MaintenanceFeeCode | Text | Fee code charged monthly when balance is below MinBalance. Empty = none |
| IsActive | Enum | `Yes`, `No` |

### tblRateTiers (Typed) - interest rate history

| Column | Type | Rules |
|---|---|---|
| RateID | Text | Key. `R0001` |
| ProductCode | Text | Must exist in tblProducts |
| TermMonths | Int | 0 for savings and current. 3, 6, 12, 24 for fixed deposits |
| FromAmount | Int | Tier starts here (inclusive) |
| ToAmount | Int | Tier ends here (exclusive). Empty = no upper limit |
| AnnualRate | Rate | Annual fraction |
| EffectiveFrom | Date | The latest row on or before the calculation date wins |

The tier that contains the WHOLE balance applies to the whole balance (not a marginal rate).

### tblFees (Typed)

| Column | Type | Rules |
|---|---|---|
| FeeCode | Text | Key |
| Name | Text | Required |
| Method | Enum | `Flat`, `Percent` |
| FlatAmount | Int | Used when Method = Flat |
| Rate | Rate | Used when Method = Percent |
| MinFee | Int | Lowest fee (Percent) |
| MaxFee | Int | Highest fee (Percent). 0 = no cap |
| Currency | Text | `MMK` |
| IncomeGL | Text | GL code that receives the fee, `4100` |

Codes used by Stage 1: `CROSS_BRANCH_WD`, `TRANSFER_SAME`, `TRANSFER_BRANCH`, `INTERBANK_OUT`, `INTERBANK_IN`, `MAINT_CUR_LOW`, `FD_EARLY_BREAK`.

### tblOtherBanks (Typed) - SIMULATED banks

| Column | Type | Rules |
|---|---|---|
| BankCode | Text | Key. `OB01` |
| Name | Text | Required |
| SettleDays | Int | Days from submission to settlement |
| FailRate | Rate | Chance that an outgoing transfer fails (decided at end of day) |
| ReturnRate | Rate | Reserved for random returns |

### tblChart (Typed) - chart of accounts

| Column | Type | Rules |
|---|---|---|
| GLCode | Text | Key |
| Name | Text | Required |
| Type | Enum | `Asset`, `Liability`, `Equity`, `Income`, `Expense` |
| NormalSide | Enum | `Debit`, `Credit` |
| Stage | Int | Stage that starts to use it |
| IsActive | Enum | `Yes` only for accounts the code uses now (Stage 1) |

Stage 1 accounts: `1000` Cash in vault, `1100` Treasury and settlement balances, `1900` Inter-branch clearing, `2000` Savings deposits, `2010` Current account deposits, `2020` Fixed deposits, `2200` Withholding tax payable, `2900` Interbank payable (simulated), `3000` Share capital, `3100` Retained earnings, `4100` Fee income, `5000` Interest expense on deposits.
The file also lists inactive accounts for Stages 2 to 8 (loans 13xx, repossessed assets 1400, fixed assets 1500/1510, investments 1600, foreign currency 1700, e-wallets 2030, taxes 221x/223x, accrued expenses 2300, staff 2400, income 4xxx, running costs 52xx, salaries 51xx, provisions 5300). Stages 3 to 8 add accounts as they need them (for example `1320` Loan loss allowance in Stage 3).

### tblAccounts (Macro)

| Column | Type | Rules |
|---|---|---|
| AccountNo | Text | Key. 3 digit branch number + 7 digit serial, 10 characters |
| CustomerID | Text | Must exist |
| BranchID | Text | The customer's home branch |
| ProductCode | Text | Must exist |
| Currency | Text | `MMK` |
| Status | Enum | `Active`, `Dormant`, `Frozen`, `Closed` |
| OpenDate | Date | |
| TermMonths | Int | Fixed deposits only |
| MaturityDate | Date | Fixed deposits only |
| LockedRate | Rate | Fixed deposits only. Annual rate fixed at opening |
| PayoutAccountNo | Text | Fixed deposits only. Receives principal and interest at maturity or break |
| Balance | Int | Must equal the sum of its Transactions rows and its journal lines |
| AccruedInterest | Decimal | Interest earned but not posted yet (4 decimals) |
| LastAccrualDate | Date | Last business date interest was accrued |
| LastActivityDate | Date | Last customer-initiated transaction |
| ClosedDate | Date | |

### tblTransactions (Macro) - one row per account per effect (the account statement)

| Column | Type | Rules |
|---|---|---|
| TxnID | Text | Key. `TX20261013-00001` (sequence per business date) |
| GroupID | Text | One business action. `G20261013-00001`. Shared by all rows of that action |
| TxnDate | Date | Business date |
| AccountNo | Text | Must exist |
| TxnType | Enum | `Opening`, `Deposit`, `Withdrawal`, `TransferOut`, `TransferIn`, `InterbankOut`, `InterbankIn`, `InterbankRefund`, `InterbankReturn`, `Fee`, `FeeRefund`, `InterestCredit`, `Reversal`, `FDMaturity`, `FDBreak` |
| Direction | Enum | `Debit`, `Credit` (from the customer's point of view: Debit reduces the balance) |
| Amount | Int | > 0 |
| BalanceAfter | Int | Account balance after this row |
| CounterAccountNo | Text | Other account, or other bank code |
| BranchID | Text | Processing branch |
| Description | Text | |
| Reference | Text | Fee code for Fee rows, otherwise free text |
| JournalID | Text | Must exist in tblJournal |
| Status | Enum | `Posted`, `Reversed` |
| CreatedBy | Text | UserID |
| ApprovedBy | Text | UserID, empty if no approval was needed |

### tblInterbankTransfers (Macro) - FAKE other-bank transfers

| Column | Type | Rules |
|---|---|---|
| TransferID | Text | Key. `IB000001` |
| Direction | Enum | `Out`, `In` |
| AccountNo | Text | Our account |
| OtherBankCode | Text | Must exist in tblOtherBanks |
| OtherBankAccount | Text | 10 digit text. Nothing is stored about it |
| Amount | Int | > 0 |
| Fee | Int | Fee charged (Out only) |
| SubmittedDate | Date | |
| ExpectedSettleDate | Date | |
| Status | Enum | `Submitted`, `Settled`, `Failed`, `Returned` |
| StatusDate | Date | |
| Reason | Text | For Failed and Returned |
| SubmitJournalID | Text | |
| SettleJournalID | Text | Settlement journal, or the failure journal |
| ReturnJournalID | Text | Only for Returned |

### tblApprovals (Macro)

| Column | Type | Rules |
|---|---|---|
| ApprovalID | Text | Key. `AP000001` |
| RequestType | Enum | `LargeDeposit`, `LargeWithdrawal`, `LargeTransfer`, `LargeInterbank`, `CashShipment`, `Reversal`, `BreakFixedDeposit`, `CloseAccount`, `Unfreeze`, `Reactivate` |
| AccountNo | Text | |
| CounterAccountNo | Text | |
| Amount | Int | |
| BranchID | Text | Processing branch |
| Note | Text | |
| RequestedBy | Text | UserID |
| RequestedAt | DateTime | |
| Status | Enum | `Pending`, `Approved`, `Rejected` |
| ApprovedBy | Text | Must differ from RequestedBy |
| DecidedAt | DateTime | |
| DecisionNote | Text | Reason for a rejection |
| TxnGroupID | Text | Set when the approved action was posted |

### tblJournal (Macro)

| Column | Type | Rules |
|---|---|---|
| JournalID | Text | Key. `J00000001` |
| PostDate | Date | Business date |
| Description | Text | |
| JournalType | Text | `Opening`, `CashPurchase`, `CashShipment`, `Deposit`, `Withdrawal`, `Transfer`, `InterbankOut`, `InterbankSettle`, `InterbankFail`, `InterbankReturn`, `InterbankIn`, `InterestPosting`, `MaintenanceFee`, `FDMaturity`, `FDBreak`, `Reversal` |
| GroupID | Text | Same as the Transactions rows |
| Status | Enum | `Posted`, `Reversed` |
| ReversalOf | Text | JournalID this reverses |
| CreatedBy | Text | |
| CreatedAt | DateTime | |

### tblJournalLines (Macro)

| Column | Type | Rules |
|---|---|---|
| LineID | Text | Key. `L0000001` |
| JournalID | Text | Must exist |
| BranchID | Text | Must exist |
| GLCode | Text | Must exist and be active |
| AccountNo | Text | Customer deposit lines only |
| Currency | Text | `MMK` |
| Debit | Int | One of Debit and Credit is > 0 |
| Credit | Int | |
| FxRate | Decimal | Empty in Stage 1 |
| ForeignAmount | Decimal | Empty in Stage 1 |

Rules: per journal total Debit = total Credit, and the same inside every BranchID.

### tblDailyClose (Macro) - one row per finished business day

| Column | Type | Rules |
|---|---|---|
| BusinessDate | Date | Key |
| RunAt | DateTime | |
| AccountsAccrued | Int | |
| InterestAccrued | Decimal | |
| InterbankSettled | Int | Transfers decided that day |
| FDMatured | Int | |
| DormantMarked | Int | |
| ChecksPassed | Enum | `Yes`, `No` |
| Notes | Text | |

### tblAudit (Macro) and tblLog (Macro)

`tblAudit`: `AuditID`, `Timestamp`, `UserID`, `Action`, `Target`, `Detail`.
`tblLog`: `Timestamp`, `Level` (`Info`, `Warning`, `Error`), `Procedure`, `Message`.

---

## Generated report sheets (cleared and rewritten, never typed)

`Control`, `TrialBalance`, `BranchSummary`, `CashPosition`, `DepositSummary`, `Checks`, `Statement`.

---

## Reference tables for later stages (in the starter workbook, not used by Stage 1 code)

| Table | Columns | Stage |
|---|---|---|
| tblCurrencies | `Code`, `Name`, `Decimals`, `IsBase` | 6 |
| tblExchangeRates | `RateDate`, `Currency`, `BuyRate`, `SellRate`, `MidRate` (MMK per unit, 30 days) | 6 |
| tblTaxRates | `TaxCode`, `Name`, `Rate`, `AppliesTo`, `GLCode`, `EffectiveFrom` (WHT on interest 10%, commercial tax on fees 5%, income tax 25%, employer social security 3%. Sample values, not real law) | 2, 5, 8 |
| tblBranchCosts | `CostID`, `BranchID`, `CostType`, `Description`, `Frequency` (`Monthly`, `Quarterly`, `Yearly`), `Amount`, `GLCode`, `StartDate`, `EndDate`, `Status` | 2 |
| tblFixedAssets | `AssetID`, `BranchID`, `AssetType`, `Description`, `AcquireDate`, `Cost`, `UsefulLifeMonths`, `ResidualValue`, `DepMethod`, `AccumDepAt20260930`, `Status` | 2 |

Branch cost rows: rent (rented premises) or property tax and licences (owned), electricity, water, internet and phone, security, cleaning, maintenance contract, stationery, generator fuel, insurance. Amounts are higher for city branches and head office than for towns. Straight line depreciation: monthly charge = (Cost - ResidualValue) / UsefulLifeMonths until AccumulatedDepreciation reaches Cost - ResidualValue.

---

## What the starter workbook holds (as of 2026-10-13)

- 15 branches (head office, 7 city, 7 town), 49 users (4 head office staff, 3 per branch)
- 320 customers, 447 accounts (301 savings, 104 current, 42 fixed), status 437 Active, 4 Dormant, 3 Frozen, 3 Closed
- Opening balances dated 2026-09-30, then sample activity 2026-10-01 to 2026-10-12: 703 journals, 2,389 journal lines, 1,337 transaction rows, 36 interbank transfers, 9 approvals (4 pending, 4 approved, 1 rejected), 12 daily close rows
- One deposit reversed, a few cross-branch deposits and withdrawals, a few large deposits that needed approval
- The ledger passes all control checks. GL 1900 nets to 0.

---

## Added by Stages 2 to 4 (matches the code; full detail in `STAGES_2_TO_4.md`)

RunSetup adds, without touching existing data:
- Columns at the end of `tblFixedAssets` (`AccumDepreciation`, `LastDepPeriod`, `LoadedOn`, `DisposedOn`, `DisposalProceeds`), `tblFees` (`Taxable`) and `tblDailyClose` (`CostsPosted`, `DepreciationPosted`, `InstallmentsDue`, `AutoDebited`, `PenaltyAccrued`, `ProvisionPosted`).
- Tables: `tblCostPostings`, `tblExpenses`, `tblTaxPayments`, `tblBudgets` (Stage 2); `tblLoanProducts`, `tblLoanClasses`, `tblLoans`, `tblLoanSchedule`, `tblLoanPayments` (Stage 3); `tblCollateralTypes`, `tblCollateral`, `tblAppraisals`, `tblRepossessions`, `tblAssetSales`, `tblAssetRentals`, `tblAssetOps` (Stage 4).
- Config keys listed in `STAGES_2_TO_4.md` section 0.2, GL accounts `1320`, `4020`, `4030`, `5320`, and activation of the accounts the stages use.
- Dropdown values: role `LoanOfficer`; the new transaction types and approval request types.

Implementation notes: a one-off expense is written to `tblExpenses` with status `Paid` when it is posted (an expense above the approval limit is written when approved); holding costs of repossessed assets are written with status `Posted` and are paid by `RunPayBills`.
`tblRepossessions.Deficiency` is what the customer still owes; the claim still open is `OriginalCarryingValue + Deficiency`.
