# Golden Delta Bank (SIMULATION) - Excel VBA, Stages 1 to 8

A simulation of ONE bank with a head office and branches in cities and towns, built in Excel VBA. Everything is made up: names, banks, rates, taxes, salaries and limits. Nothing here is legal, tax or financial advice.

This file is the full specification and the user manual. It is the only document you need.

Contents: 1 What it is - 2 Install and first run - 3 The Control sheet - 4 Roles and approvals - 5 How the bank works (rules) - 6 Posting rules - 7 Feature manual (step by step) - 8 End of day - 9 Control checks - 10 Reports - 11 Data tables - 12 Config keys - 13 Design and rules for the code - 14 Test checklist - 15 Troubleshooting

---

## 1. What it is

| Stage | What it adds |
|---|---|
| 1 | Customers, accounts (savings, current, fixed deposit), cash deposit and withdrawal, transfers inside the bank and between branches, fake transfers to and from other banks, a double-entry ledger, interest, fees, approvals, end of day, month end, reports, control checks |
| 2 | Branch running costs, fixed assets and depreciation, one-off expenses, taxes (withholding tax on interest, commercial tax on fees), budgets, branch profit and loss |
| 3 | Loans (flat and declining), schedules, auto-debit, penalties, classes, provisions, write-off, recovery |
| 4 | Collateral, repossession, sale, rent, redemption |
| 5 | Staff, salary scales, allowances, payroll, income tax and social security, staff advances |
| 6 | Foreign currency accounts, currency counter, fake market, international payments (fake), daily revaluation, e-wallets, QR payments, agents |
| 7 | Investments in businesses (equity, profit share, notes, trade finance) |
| 8 | Policy rate, reserve and capital rules, liquidity, loan quality, deposit insurance, government securities, income tax, anti money laundering, inspection and fines, year-end close, annual reports, dashboard |

Scope decisions that never change:
- ONE bank only. There is no central bank and no second bank in the books.
- Transfers between branches are real internal transfers.
- Transfers to and from OTHER banks are FAKE. They use one simulated settlement balance (GL 1100) and one simulated payable (GL 2900). Foreign banks (Stage 6) are fake in the same way. Nothing about another bank is stored except its name and the account text.
- Branch types: `HeadOffice`, `City`, `Town`. Head office is branch `BR001`.
- The business date is a number you control (Config `BusinessDate`), not the computer clock. Every day is a business day.
- Money is whole MMK (no decimals). Foreign currencies use their own decimals (table `tblCurrencies`).

---

## 2. Install and first run

1. Save `BankStarter.xlsx` as `BankStarter.xlsm` (macro-enabled).
2. Open the VBA editor (Alt+F11). Import ALL files from the `src` folder with File > Import File. **Remove old versions first** when you re-import: an old module gives "Method or data member not found" or "User-defined type not defined".
   Import order (the order only matters between groups):
   1. All class files `cls*.cls` (14): clsAccount, clsApproval, clsBranch, clsCustomer, clsErrorHandler, clsFee, clsJournalLine, clsLogger, clsOtherBank, clsProduct, clsResult, clsTxnRow, clsUnitOfWork, clsUser.
   2. `modConstants`, `modMessages`, `modSheetUtils`, `modPerformance`, `modMoney`, `modDateUtils`, `modIds`, `modArgs`.
   3. All repositories `repo*.bas`, then all services `svc*.bas`.
   4. Last: `modSetupStages`, `modSetupLate`, `modSetup`, then the entry modules `modMain`, `modMainCosts`, `modMainLoans`, `modMainCollateral`, `modMainStaff`, `modMainFx`, `modMainInvest`, `modMainGov`.
3. Debug > Compile VBAProject. There must be no error.
4. Run **RunSetup** (Alt+F8). It creates the `Control` sheet with its buttons, every missing sheet and table, dropdowns, Config keys, GL accounts and seed rows. **It never clears or changes existing data.** On the starter workbook it keeps all rows (Customers 320, Accounts 447, Transactions 1,337, Journal 703, JournalLines 2,389).
5. Open the `Control` sheet, choose a user in C4, and use the buttons (section 7).
6. To check the install press **Reports**, then open the `Checks` sheet: every check must say PASS.

The Control sheet title and this file say SIMULATION on purpose.

---

## 3. The Control sheet

You never type into `Accounts.Balance`, `Transactions`, `Journal`, `JournalLines`, `InterbankTransfers`, `DailyClose` or `Approvals`. Only the macros (buttons) change money. Master data (`Customers`, `Branches`, `Users`, `Products`, `RateTiers`, `Fees`, `OtherBanks`, `Config` and the Stage 2 to 8 reference tables) is typed by hand.

How to act: pick a user in **C4**, fill the cells the action needs, press the button. The result text appears in **C18** and in a message box.

| Cell | Meaning | Cell | Meaning |
|---|---|---|---|
| C2 | Business date (shown, written by the macros) | C30 | Employee ID |
| C4 | Current user ID (dropdown) | C31 | Position code |
| C5 | Account number | C32 | Grade (1 to 3) |
| C6 | Counter account (transfer), other bank account (interbank), bank account (wallet) | C33 | Effective date (staff changes, hire date) |
| C7 | Amount | C34 | Payroll run ID |
| C8 | Processing branch (blank = the user's branch) | C35 | Currency code |
| C9 | Other bank code | C36 | Mid rate (when setting rates) |
| C10 | Term in months (fixed deposit, loan, investment); dependents (hire) | C37 | Merchant ID |
| C11 | Payout account (fixed deposit) | C38 | Agent ID |
| C12 | Transaction group, approval, transfer, asset, case or payment ID | C39 | Beneficiary or sender name |
| C13 | Customer ID | C40 | Foreign bank code |
| C14 | Product code (deposit products) | C41 | Beneficiary account (text) |
| C15 | Reference or note (see "C15 is used for") | C42 | Purpose |
| C16, C17 | Statement or report from and to dates | C44 | Investee ID |
| C18 | Last status message | C45 | Investment ID |
| C19 | GL code (expense) | C46 | Investment type |
| C20 | Vendor, tenant, buyer, name, investee name | C47 | Rate or share percent |
| C21 | Description, sector | C48 | Reported profit, or new valuation |
| C22 | Asset or collateral type code | C49 | AML alert ID |
| C23 | Useful life in months | C50 | Bond type (`TBill` or `GovBond`) |
| C24 | Residual value, selling costs, monthly recovery | C51 | Face value |
| C25 | Period `YYYYMM`, or year `YYYY` | C52 | Term in days |
| C26 | Loan product code | C53 | Rate (0.08 or 8) |
| C27 | Mode: `Account` or `Cash` | C54 | Reason |
| C28 | Loan ID or repossession ID | | |
| C29 | Collateral ID | | |

C43 is not used. Account numbers keep their leading zeros (the cells are text).

**C15 is used for:** `RETURN` on a cash shipment (send cash back to head office); the decision `ClosedNoAction` or `Escalated` on `RunReviewAlert`; the type `Rent`, `Revenue`, `Cost` on `RunRecordAssetOp`; the purpose `Pledge`, `Periodic`, `Repossession`, `Sale` on `RunAppraise`; the collateral location on `RunRegisterCollateral`; the type `Interest`, `Dividend`, `ProfitShare` on `RunReceiveInvestmentReturn`; the method `Cost`, `Appraisal`, `Multiple` on `RunValueInvestment`; the pay frequency `Monthly`, `Quarterly`, `AtMaturity` on `RunProposeInvestment`; a phone number on `RunHireEmployee`; the case summary or decision on `RunOpenCase` and `RunCloseCase`; and `OVERRIDE: reason` (HeadOffice) to go on after a blocked-name hit.

**IDs** (always the highest existing number plus 1, never reused): CustomerID `C000001`; AccountNo = 3-digit branch number + 7-digit serial (`0060000398` is branch BR006); TxnID `TX20261013-00001`; GroupID `G20261013-00001`; JournalID `J00000001`; TransferID `IB000001`; ApprovalID `AP000001`; AuditID `AU000001`; CostPostingID `CP0000001`; ExpenseID `EX000001`; TaxPaymentID `TP000001`; BudgetID `BG000001`; LoanID `LO000001`; ScheduleID `SC0000001`; LoanPaymentID `LP000001`; CollateralID `CL000001`; AppraisalID `VA000001`; RepoID `RP000001`; SaleID `SL000001`; RentalID `RN000001`; OpID `OP000001`; EmployeeID `E00001`; HistoryID `HI000001`; AttendanceID `AT000001`; AdvanceID `AD000001`; PayrollRunID `PR000001`; PayslipID `PS0000001`; DealID `XD000001`; IntlPaymentID `IP000001`; MerchantID `MC000001`; AgentID `AG000001`; NameID `NM000001`; InvesteeID `IV000001`; InvestmentID `IN000001`; ReturnID `IR000001`; ValuationID `IW000001`; BondID `BD000001`; AlertID `AL000001`; CaseID `CS000001`; EventID `CE000001`; FineID `FN000001`.

---

## 4. Roles and approvals

| Role | May do |
|---|---|
| Teller | deposit, withdraw, transfer, open account, open fixed deposit, small interbank out, statement, loan payment, currency deals, international payments, wallet operations |
| LoanOfficer | quote, apply, disburse, pledge, record loan payments (cannot approve) |
| BranchManager | everything a Teller does, approve for their own branch up to `BranchManagerLimit`, freeze, unfreeze, reactivate, close account, break fixed deposit, expenses, collateral, default notices, staff advances |
| HeadOffice | approve anything, cash shipments, reversals, interbank in and return, bills, taxes, assets, write-offs, repossession, payroll, rates, market deals, investments, bonds, policy rate, AML, income tax, reports for all branches |
| Auditor | reports only |
| Admin | setup, end of day, config, catch-up cost posting, load fixed assets, sample staff, simulate investees, year-end close. Cannot post customer transactions |
| Regulator | reports, compliance check, inspection, propose fines (read-only otherwise) |

A Teller and a BranchManager work only for their own branch as processing branch. HeadOffice and Admin may pick any branch in C8.

**Maker and checker.** A request that needs approval creates ONE `Approvals` row with status `Pending` and posts NOTHING. A different user decides it: put the ApprovalID in C12 and press **Approve** (or **Reject**, with a reason in C15). On approval the macro re-checks every rule (the balance may have changed) and then posts. If a re-check fails the approval becomes `Rejected` with the reason. The same user can never approve their own request.

Who may approve: a BranchManager for their own branch up to `BranchManagerLimit` (loans: up to `LoanBranchLimit`); HeadOffice anything. **HeadOffice only:** cash shipments, reversals, bills, taxes, assets, write-offs, repossession, sales, redemption, payroll, staff advances, foreign limit breaches, investments, bonds, policy rate, fines, year-end close, case reports, income tax.

What needs approval: any deposit, withdrawal, transfer or interbank out above `TellerLimit`; any amount above the daily limit; every cash shipment; every reversal; early break of a fixed deposit; closing an account; unfreezing or reactivating; plus the items listed in the feature manual (marked "approval").

Every approval decision, status change, reversal, end of day and config change writes a row in `tblAudit`.

---

## 5. How the bank works (rules)

### 5.1 Ledger
- Every posting is a journal with lines. Total debit = total credit in the journal AND inside every branch separately. A movement across branches uses GL `1900` Inter-branch clearing in each branch. GL 1900 over all branches is always 0.
- `svcLedger.PostJournal` is the ONLY code that writes `Journal` and `JournalLines`. It refuses unless: at least 2 lines; every GL exists and is active; every branch exists; each line has a debit or a credit (not both, not zero); debits = credits; every branch balances; PostDate = business date.
- Customer deposit lines carry `AccountNo`; all other lines leave it empty. Lines of a customer account always sit in the account's HOME branch.
- A reversal posts the same lines with debit and credit swapped, dated today, `ReversalOf` = the original JournalID. The original journal and its statement rows become `Reversed`. Refused: reversing a reversal, interest and tax postings, interbank settlement journals, payroll payments, international payments by hand, revaluations, closing journals, anything in a closed year.
- Nothing posted is ever deleted. A mistake is fixed with a reversing journal.
- Each business action builds ALL its changes in memory (journal, lines, statement rows, balances, other tables), checks them, then writes them once (unit of work). An error halfway never leaves a half-posted action.

### 5.2 Fees
`Fee = FlatAmount` for `Flat`. Otherwise `Fee = RoundCcy(Amount x Rate)`, at least `MinFee`, at most `MaxFee` when `MaxFee > 0`. A fee is refused if the account cannot cover amount + fee (+ tax).
**Commercial tax (from Stage 2):** a fee with `Taxable = Yes` also carries `Tax = RoundCcy(Fee x COMMERCIAL_FEES rate)`. The customer pays fee + tax. Wherever the fee is credited to 4100 in a branch, the tax is credited to 2210 in the same branch. Statement rows: `Fee` and `Tax`. Fee codes: `CROSS_BRANCH_WD`, `TRANSFER_SAME`, `TRANSFER_BRANCH`, `INTERBANK_OUT`, `INTERBANK_IN`, `MAINT_CUR_LOW`, `FD_EARLY_BREAK`, `INTL_OUT`, `INTL_MSG`, `INTL_IN`, `WLT_CASHOUT`, `WLT_P2P`, `MERCHANT_MDR`. Not taxable: `TRANSFER_SAME`, `INTERBANK_IN`, `WLT_P2P`.
Examples: cash withdrawal of 1,000,000 at another branch: fee 1,000, tax 50, customer pays 1,001,050. Interbank out 1,000,000: fee 1,000, tax 50.

### 5.3 Accounts
- Opening: customer exists and is Active; product Active; opening deposit >= product `MinOpening`; the currency is the product's. A fixed deposit needs a `TermMonths` that exists in the product's rate tiers, a payout account of the same customer, and is funded by cash or by a transfer from another account of the same customer.
- Withdrawal or outgoing transfer: balance after >= product `MinBalance`. Fixed deposits cannot be withdrawn or transferred out (only matured or broken). The balance never goes negative. No overdrafts.
- Status: `Frozen` allows credits, refuses debits. `Dormant` is set at end of day when `BusinessDate - LastActivityDate > DormantDays` for an Active savings or current account; customer transactions are refused until a BranchManager runs Reactivate; interest still accrues; fixed deposits are never Dormant. `Closed` is refused for everything. Closing pays pending interest, pays the whole balance out as cash, sets Closed (approval, at the account's home branch, refused while an interbank transfer is Submitted).
- `LastActivityDate` changes only for customer transactions, not for interest, fees, tax, maintenance, auto-debit.
- Daily limit: withdrawals + outgoing transfers + interbank out of one account on one business date may not exceed `DailyWithdrawalLimit` without approval.

### 5.4 Interest
- Day count `DayCountBasis` (365).
- **Savings and current:** at every end of day, for every account that is not Closed with Balance > 0: `AccruedInterest += Balance x Rate / 365`. Rate = the tier that contains the WHOLE balance (`FromAmount <= Balance < ToAmount`, blank `ToAmount` = no upper limit), with `TermMonths = 0`, using the latest `EffectiveFrom` on or before the business date. Accounts opened today accrue today. Accruals are kept to 4 decimals until posted.
- **Fixed deposit:** the rate is locked at opening (`LockedRate`). It accrues daily on the principal for every date BEFORE `MaturityDate`. Interest is posted only at maturity or break.
- **Month end** (end of day on the last calendar day, after accrual): for every savings and current account whose `AccruedInterest` rounds to at least 1: gross `G = RoundCcy(AccruedInterest)`, tax `W = RoundCcy(G x WithholdingTaxRate)` unless `Customers.TaxExempt = Yes`; post; `AccruedInterest -= G`. A statement row `InterestCredit` shows gross and tax.
- **Maturity:** at end of day of `MaturityDate`: gross, tax, post, then transfer the whole balance to `PayoutAccountNo` (must be Active; if not, keep the deposit, log a warning, retry next day). Status becomes Closed.
- **Early break** (approval): interest = `RoundCcy(Principal x EarlyBreakRate x DaysHeld / 365)`, tax as above, fee `FD_EARLY_BREAK` on the principal, then the rest goes to the payout account.
- **Maintenance fee:** at month end, for each Active current account whose product has a `MaintenanceFeeCode` and whose balance is below `MinBalance`: charge it if the balance can cover it, else skip and log.
- Foreign accounts: interest, tax and fees are in the account currency; interest expense 5000 is booked in MMK at mid.

### 5.5 Cash and branch vaults
- Every branch has a vault (GL 1000 of that branch). A withdrawal is refused if the processing branch vault would go below 0: "Branch vault is short. Ask head office for a cash shipment." Shipments are never created automatically.
- A cash shipment moves cash from head office to a branch or back. It ALWAYS needs approval from a different HeadOffice user. If the head office vault is short, the same action first posts a "cash order from treasury" in blocks of `CashShipmentRounding`.
- The cash position report flags branches below `VaultMin` or above `VaultMax`.

### 5.6 Other-bank transfers (fake)
- Outgoing: posted at once, with a `tblInterbankTransfers` row `Submitted`, `ExpectedSettleDate = business date + SettleDays`.
- At end of day every Submitted outgoing transfer due on or before the date is decided with `Rnd`: with probability `FailRate` it is `Failed` (refund, fee refunded), otherwise `Settled`. Randomize is called once per end of day. `RunReturnInterbank` (C12 = TransferID, Settled only) forces a `Returned` result for testing (fee is kept).
- Incoming (`RunInterbankIn`): posts at once with a `Settled` row.
- Settlement of a transfer after the account was closed or frozen still posts (it touches only 2900, 1900, 1100).

---

## 6. Posting rules

H = account's home branch, P = processing branch, HO = BR001. Deposit GL by product type: Savings `2000`, Current `2010`, Fixed `2020`, Wallet `2030`. Lines are shown as "branch: Dr or Cr GL amount".

### Stage 1
| Event | Lines |
|---|---|
| Cash deposit A to account X | P: Dr 1000 A. If P<>H: P: Cr 1900 A, H: Dr 1900 A. H: Cr deposit GL A (AccountNo X) |
| Cash withdrawal A, fee F (only if P<>H, `CROSS_BRANCH_WD`) | H: Dr deposit GL (A+F). If P<>H: H: Cr 1900 (A+F), P: Dr 1900 (A+F), P: Cr 4100 F. P: Cr 1000 A. (With tax X: customer pays A+F+X, Cr 2210 X in P) |
| Transfer A from X (H1) to Y (H2), fee F (only if H1<>H2, `TRANSFER_BRANCH`) | H1: Dr deposit GL X (A+F), Cr 4100 F. Same branch: H1: Cr deposit GL Y A. Different: H1: Cr 1900 A, H2: Dr 1900 A, H2: Cr deposit GL Y A |
| Cash shipment HO to branch B, A | B: Dr 1000 A, Cr 1900 A. HO: Dr 1900 A, Cr 1000 A. Return shipment is the reverse |
| Cash order from treasury | HO: Dr 1000, Cr 1100 |
| Interbank out A, fee F | H: Dr deposit GL (A+F), Cr 2900 A, Cr 4100 F |
| ...settled | H: Dr 2900 A. If H<>HO: H: Cr 1900 A, HO: Dr 1900 A. HO: Cr 1100 A. If H=HO: H: Cr 1100 A |
| ...failed (before settlement) | H: Dr 2900 A, Dr 4100 F, Cr deposit GL (A+F). Fee refunded |
| ...returned (after settlement) | HO: Dr 1100 A. If H<>HO: HO: Cr 1900 A, H: Dr 1900 A. H: Cr deposit GL A. Fee kept |
| Interbank in A | HO: Dr 1100 A. If H<>HO: HO: Cr 1900 A, H: Dr 1900 A. H: Cr deposit GL A |
| Interest posting (gross G, tax W) | H: Dr 5000 G, Cr deposit GL (G-W) (AccountNo X), Cr 2200 W |
| Monthly maintenance fee F | H: Dr deposit GL F, Cr 4100 F |
| Fixed deposit maturity | interest posting, then transfer of the whole balance to the payout account (no fee) |
| Fixed deposit early break | interest posting at the early rate, fee `FD_EARLY_BREAK` on the principal (H: Dr 2020, Cr 4100), transfer of the rest to the payout account |
| Opening deposit | like a cash deposit (or the transfer rule if funded from another account) |

### Stage 2
| Event | Lines |
|---|---|
| Cost posting (month, one journal per branch) | B: Dr expense GL, Cr 2300 |
| Pay bills | B: Dr 2300, Cr 1900. HO: Dr 1900 (total), Cr 1100 (total) |
| Expense up to `PettyCashLimit` | B: Dr expense GL, Cr 1000 |
| Larger expense | B: Dr expense GL, Cr 1900. HO: Dr 1900, Cr 1100 |
| Load fixed assets (once) | B: Dr 1500 (cost), Cr 1510 (accumulated), Cr 3000 (net book value) |
| Depreciation | B: Dr 5250, Cr 1510 |
| Buy asset | B: Dr 1500, Cr 1900. HO: Dr 1900, Cr 1100 |
| Dispose asset | B: Dr 1510 (accumulated), Dr 1900 (proceeds), Cr 1500 (cost); gain: Cr 4400, loss: Dr 5310. HO: Dr 1100 (proceeds), Cr 1900 |
| Pay taxes | B: Dr 2200, Dr 2210 (and 2230, 2220), Cr 1900. HO: Dr 1900 (total), Cr 1100 |

### Stage 3
| Event | Lines |
|---|---|
| Disbursement | H: Dr 1300 principal, Cr deposit GL (loan account) |
| Processing fee + tax | H: Dr deposit GL (fee+tax), Cr 4100 fee, Cr 2210 tax |
| Installment becomes due | H: Dr 1310 interest, Cr 4000 interest |
| Repayment from account | H: Dr deposit GL (total), Cr 4020 (penalty), Cr 1310 (interest), Cr 1300 (principal) |
| Repayment in cash at P | P: Dr 1000. If P<>H: P: Cr 1900, H: Dr 1900. H: Cr 4020, Cr 1310, Cr 1300 |
| Payoff | as repayment, plus interest since the last due date credited to 4000, plus the early payoff fee (taxable) |
| Provisions (month end, per branch) | required > current: B: Dr 5300, Cr 1320 for the difference; lower: B: Dr 1320, Cr 5300 |
| Write-off | B: Dr 1320 (provision), Dr 5300 (rest), Cr 1300 (principal), Cr 1310 (interest) |
| Recovery | Dr deposit GL or cash, Cr 4030 |

### Stage 4
| Event | Lines |
|---|---|
| Repossession (loan branch) | Dr 1400 CV; Dr 1320 Prov; Dr or Cr 5300 (B - CV - Prov); Cr 1300 principal; Cr 1310 recognized interest |
| Holding cost (month end) | an expense row (GL 5220, vendor "Holding cost"), unpaid until Pay bills |
| Impairment | Dr 5320, Cr 1400 (write-up the reverse, never above the original carrying value) |
| Sale | Dr 1000 (or 1100) Net; Cr 1400 CarryingValue; Cr 4030 (Applied - CV) or Dr 5310 (CV - Applied); Cr customer deposit Surplus (statement `SurplusPaid`) |
| Rent or revenue | H: Dr 1000 (or tenant account), Cr 4500 |
| Asset cost | H: Dr 5220, Cr 1000 |
| Redemption | same postings as a sale with Net = the full remaining claim |

### Stage 5
| Event | Lines |
|---|---|
| Payroll posting (per branch) | Dr 5100 (gross), Dr 5110 (employer SSB); Cr 2400 (net), Cr 2230 (tax + employee SSB + employer SSB), Cr 1810 (advance recovered) |
| Pay, Account mode | B: Dr 2400 Net. Same branch: Cr deposit GL. Other branch: B: Cr 1900, A: Dr 1900, A: Cr deposit GL (statement `SalaryCredit`) |
| Pay, External mode | B: Dr 2400 Net, Cr 1900. HO: Dr 1900 (total), Cr 1100 (total) |
| Staff advance up to `PettyCashLimit` | B: Dr 1810, Cr 1000 |
| Larger advance | B: Dr 1810, Cr 1900. HO: Dr 1900, Cr 1100 |

### Stage 6 (foreign currency)
Base currency is MMK; a rate is MMK per 1 unit. Only GLs `1700`, `2000`, `2010`, `2030`, `2200` may carry a foreign currency; everything else is MMK. A foreign line stores `Currency`, `ForeignAmount`, `FxRate` (the mid rate of the business date) and `Debit` or `Credit` in MMK = `RoundCcy(ForeignAmount x FxRate)`. Every foreign leg is booked at MID; the customer spread goes to 4200. A deal is refused with no rate for today ("No rate for today. Ask head office to set rates").
| Event | Lines |
|---|---|
| Foreign cash deposit A | P: Dr 1700 (A, base A x Mid). If P<>H: P: Cr 1900, H: Dr 1900. H: Cr deposit GL (A, base A x Mid) |
| Foreign cash withdrawal | the reverse; the foreign cash at P must cover it |
| Foreign cash shipment | B: Dr 1700, Cr 1900. HO: Dr 1900, Cr 1700 |
| Customer buys foreign cash with MMK | `X = RoundCcy(A x Sell)`. P: Dr 1000 X; Cr 1700 (A, base A x Mid); Cr 4200 (X - A x Mid) |
| Customer sells foreign cash | `Y = RoundCcy(A x Buy)`. P: Dr 1700 (A, base A x Mid); Cr 1000 Y; Cr 4200 (A x Mid - Y) |
| Convert MMK to foreign | H: Dr deposit MMK X; Cr deposit foreign (A, base A x Mid); Cr 4200 (X - A x Mid) |
| Convert foreign to MMK | H: Dr deposit foreign (A, base A x Mid); Cr deposit MMK Y; Cr 4200 (A x Mid - Y) |
| Market buy (fake market) | `Cost = RoundCcy(A x Mid x (1 + MarketSpread))`. HO: Dr 1700 (A, base A x Mid); Dr 4200 (Cost - A x Mid); Cr 1100 Cost |
| Market sell | `Proceeds = RoundCcy(A x Mid x (1 - MarketSpread))`. HO: Dr 1100 Proceeds; Dr 4200 (A x Mid - Proceeds); Cr 1700 |
| Revaluation (end of day) | per (branch, GL, currency): `Target = RoundCcy(Foreign x Mid)`. Assets (1700): Diff>0: Dr GL, Cr 4200; Diff<0: Dr 4200, Cr GL. Liabilities (2000, 2010, 2030, 2200): Diff>0: Dr 4200, Cr GL; Diff<0: Dr GL, Cr 4200. Same branch. Journal type `FxRevaluation` |
| Foreign interest | H: Dr 5000 (base G x Mid), Cr deposit GL (net, foreign), Cr 2200 (W, foreign) |
| International out, foreign account | H: Dr deposit GL (amount + fees + tax, foreign). Cr 4100 (fees), Cr 2210 (tax). If H<>HO: H: Cr 1900 (amount base), HO: Dr 1900. HO: Cr 1700 (amount foreign) |
| International out, MMK account | `X = RoundCcy(A x Sell)`. H: Dr deposit MMK (X + fees + tax); Cr 4100, Cr 2210; Cr 1900 and HO: Dr 1900, Cr 1700 (A, base A x Mid); Cr 4200 (X - A x Mid) |
| International failed | the lines are reversed (amount, fees, tax refunded) |
| International returned (after Delivered) | HO: Dr 1700 (amount), Cr 1900; H: Dr 1900, Cr deposit GL. Fees kept |
| International in, foreign account | HO: Dr 1700 (A). If H<>HO: HO: Cr 1900, H: Dr 1900. H: Cr deposit GL (A) |
| International in, MMK account | `Y = RoundCcy(A x Buy)`. HO: Dr 1700 (A, base A x Mid); Cr 1900 / H: Dr 1900, Cr deposit MMK Y; Cr 4200 (A x Mid - Y). Fee `INTL_IN` charged from the account |
| Wallet top-up | H: Dr deposit GL A; if W<>H: H: Cr 1900, W: Dr 1900. W: Cr 2030 A. HO: Dr 1110 A, Cr 1100 A |
| Wallet cash-out | the reverse of the top-up; fee `WLT_CASHOUT` (+ tax) from the wallet |
| Wallet to wallet | W1: Dr 2030 A; same branch: Cr 2030 A; else W1: Cr 1900, W2: Dr 1900, Cr 2030. Trust account unchanged |
| QR payment | customer wallet debited A; merchant wallet credited A - M - TX; Cr 4100 M; Cr 2210 TX (merchant branch). `M = RoundCcy(A x MDR)` |
| Agent float in | P: Dr 1000 A; if P<>W: P: Cr 1900, W: Dr 1900. W: Cr 2030 (agent wallet). HO: Dr 1110, Cr 1100 |
| Agent cash-in | Dr 2030 agent wallet A, Cr 2030 customer wallet A. Commission: Dr 5290, Cr 2030 agent wallet |
| Agent cash-out | Dr 2030 customer wallet (A+fee+tax); Cr 2030 agent wallet A; Cr 4100 fee; Cr 2210 tax. Commission as above |
Trust rule: whenever the total of all wallets (GL 2030) changes, head office moves the same amount into or out of the trust account 1110 in the same journal.

### Stage 7
| Event | Lines |
|---|---|
| Fund investment | HO: Dr 1600 Amount, Cr 1100 |
| Month-end return accrual | HO: Dr 1610, Cr 4300 |
| Cash received | HO: Dr 1100; Cr 1610 (up to the unpaid recognized return); the rest Cr 4300. Dividend: Dr 1100, Cr 4300 |
| Profit share | profit: HO: Dr 1610, Cr 4300. Loss share: HO: Dr 5330, Cr 1600 |
| Maturity repayment | HO: Dr 1100 (principal + unpaid + accrued), Cr 1600 (carrying), Cr 1610 (recognized unpaid), Cr 4300 (the rest) |
| Impairment | HO: Dr 5330, Cr 1600 (write-up: Dr 1600, Cr 5330, never above cost) |
| Disposal | HO: Dr 1100 Proceeds; Cr 1600 CarryingValue; gain Cr 4400; loss Dr 5310 |
| Write-off | HO: Dr 5330, Cr 1600 (carrying) and Dr 5330, Cr 1610 (unpaid recognized) |

### Stage 8
| Event | Lines |
|---|---|
| Buy T-bill | HO: Dr 1620 Price, Cr 1100. `Price = RoundCcy(Face x (1 - Rate x Days / 365))` |
| T-bill accretion (month end) | HO: Dr 1620, Cr 4310 (straight line to the face value) |
| T-bill maturity | HO: Dr 1100 Face, Cr 1620 (any 1-unit rounding to 4310) |
| Buy government bond | HO: Dr 1620 Face, Cr 1100 |
| Bond coupon accrual (month end) | HO: Dr 1610, Cr 4310 |
| Bond coupon date | HO: Dr 1100, Cr 1610 (the rest, if any, Cr 4310) |
| Bond maturity | HO: Dr 1100 Face, Cr 1620 Face |
| Deposit insurance (quarter end) | an expense row (GL 5340, vendor "Deposit Insurance Fund"): Dr 5340, Cr 2300; unpaid until Pay bills |
| Income tax booking | HO: Dr 5400, Cr 2220 |
| Pay taxes | also pays 2230 and 2220: B: Dr tax GL, Cr 1900. HO: Dr 1900, Cr 1100 |
| Fine | HO: Dr 5350, Cr 1100 |
| Year-end close (per branch) | Dr each income GL balance, Cr each expense GL balance, the difference to 3100 (Cr if profit, Dr if loss). Journal type `YearEndClose` |

---

## 7. Feature manual (step by step)

Pattern for every feature: set C4, fill the cells, press the button, read C18. "Approval" means a Pending approval is created and a different user must approve it (section 4).

### 7.1 Customers and accounts (Stage 1)
- **Open account** (Teller, BranchManager): C13 customer, C14 product (SAV01, SAV02, CUR01, CUR02 and the foreign products), C7 opening deposit, C8 branch (blank = yours). Refused for an unknown or inactive customer or product, an opening deposit below `MinOpening`, a foreign currency that has no rate. Names screened against the blocked-name list (section 7.14).
- **Open fixed deposit**: C13 customer, C14 product `FD01`, C10 term (3, 6, 12, 24), C7 principal, C11 payout account (same customer), C6 funding account (blank = cash at C8). Rate is locked at opening.
- **Deposit**: C5 account, C7 amount, C8 branch (blank = yours). Above `TellerLimit`: approval. A frozen account accepts deposits.
- **Withdraw**: C5, C7, C8. Cross-branch withdrawal charges `CROSS_BRANCH_WD`. Refused if the balance would fall below the minimum, the account is frozen or dormant, or the vault is short.
- **Transfer**: C5 from, C6 to, C7 amount. Between branches `TRANSFER_BRANCH` is charged. Same account on both sides is refused. Different currencies are refused (use Convert).
- **Interbank out**: C5 our account, C6 the other bank's 10-digit account text, C9 bank code (dropdown), C7 amount. Status Submitted until end of day decides it.
- **Interbank in** (HeadOffice): C5 account, C6 other account text, C9 bank code, C7 amount. Posts at once.
- **Return interbank** (HeadOffice): C12 = TransferID (a Settled one).
- **Cash shipment** (HeadOffice): C8 branch, C7 amount, C15 `RETURN` to send cash back; C35 for a foreign currency. Always creates an approval. Approve as a different HeadOffice user.
- **Approve / Reject**: C12 = ApprovalID, C15 = reason (for Reject).
- **Reverse** (HeadOffice): C12 = the GroupID (see the Transactions sheet). Creates a reversal approval; a different HeadOffice user approves.
- **Freeze / Unfreeze / Reactivate / Close account / Break fixed deposit** (BranchManager of the account's branch, or HeadOffice): C5 account. Unfreeze, Reactivate, Close and Break create an approval.
- **Statement**: C5 account, C16 from, C17 to. Writes the `Statement` sheet.

### 7.2 Costs, assets, taxes, budgets (Stage 2)
- **Load fixed assets** (Admin, once): moves the sample assets onto the books (idempotent).
- **Post costs (this month)** / **Post costs for period** (Admin, C25 = YYYYMM, not later than the current month): posts one journal per branch for the active cost rows. Idempotent: (CostID, Period) is never posted twice. Whole month, no proration; quarterly rows = amount / 3, yearly rows = amount / 12. Depreciation is posted with it.
- **Pay bills** (HeadOffice, approval): C25 = pay everything up to this period, C8 = one branch (blank = all).
- **Post expense** (BranchManager, HeadOffice): C8 branch, C19 GL (an active expense account, 52xx or 5500), C7 amount, C20 vendor, C21 description. Up to `PettyCashLimit` paid from the vault; above `ExpenseApprovalLimit` approval.
- **Buy asset** (HeadOffice, approval): C8 branch, C22 type, C21 description, C7 cost, C23 life, C24 residual.
- **Dispose asset** (HeadOffice, approval): C12 asset ID, C7 proceeds (0 allowed). Gain to 4400, loss to 5310. No depreciation for the month of disposal.
- **Pay taxes** (HeadOffice, approval): pays every tax GL balance (2200, 2210, 2230, 2220) from all branches through head office. Foreign withholding tax is paid in kind from head office foreign cash. A `TaxPayments` row per tax code. A zero or negative balance is never paid.
- **Create budget** (HeadOffice): C25 = year. Annualised costs x (1 + `BudgetUplift`) per branch and GL, plus depreciation. Refused if the year already has budgets. Edit the rows by hand afterwards.
- Depreciation: monthly = (Cost - ResidualValue) / UsefulLifeMonths, never beyond cost less residual. A new asset starts depreciating the month after acquisition.

### 7.3 Loans (Stage 3)
- **Quote** (no data change): C26 product, C7 amount, C10 term. Writes the `LoanQuote` sheet.
- **Apply** (LoanOfficer, BranchManager, HeadOffice): C13 customer, C26 product, C7 amount, C10 term, C5 loan account (customer's Active MMK account), C29 collateral (when needed), C15 note. Checks: customer Active; KYC >= `LoanMinKYC` (3 at or above `LoanKyc3Amount`); amount and term inside the product limits; no loan of the same customer 30 or more days past due; single-borrower limit; collateral covers the principal when the product requires it and `RequireCollateralCheck = Yes`. Creates a loan `Applied` and a `LoanApproval` approval.
- **Approval**: the BranchManager of the loan's branch up to `LoanBranchLimit`, HeadOffice above it. The approver differs from the applicant. The rate is locked at approval.
- **Disburse**: C28 loan (Approved). Refused unless: liquid funds (1100 + all vaults) >= `ReserveRatio x (deposits + principal)`; loans (1300 + principal) <= `MaxLoanToDeposit x (deposits + principal)`; single-borrower limit; the account is Active. Posts the principal and the processing fee + tax, and creates the whole schedule.
- Schedule: first due date = disbursement date + 1 month (clamped to the end of a shorter month). **Flat**: total interest = Principal x Rate x Term / 12, split evenly, the last installment takes the remainder. **Declining**: annuity with monthly rate Rate / 12.
- **Loan payment**: C28 loan, C7 amount, C27 `Account` or `Cash`, C5 paying account (blank = the loan account), C8 branch for cash. Allocation: penalty, interest (oldest first), principal (oldest first). An extra amount shortens the schedule from the end.
- **Pay off**: C28 loan. Penalty + recognized interest + principal + interest since the last due date + early payoff fee. Closes the loan and releases collateral.
- **Write off** (HeadOffice, approval): C28 loan, only class Loss (or Doubtful with a note in C15). The customer still owes the money (off the books).
- **Recover**: C28 loan (WrittenOff or Repossessed), C7 amount, C27 mode, C5 or C8. Credits 4030.
- **Loan reports**: C28 for the statement of one loan.
- Daily (end of day): installments due are recognized as interest income; with `LoanAutoDebit = Yes` the oldest unpaid installments are taken from the account when it can cover the whole installment and keep its minimum balance (no partial auto-debit, `LastActivityDate` unchanged); penalty accrues per day on the unpaid amount after the grace days (4 decimals, memo until paid); days past due and class are updated; a loan whose schedule is fully paid is Closed and its collateral released.
- Classes: 1 Standard 0-29 days (provision 1%), 2 Watch 30-89 (5%), 3 Substandard 90-179 (25%), 4 Doubtful 180-359 (50%), 5 Loss 360+ (100%). Provisions are brought in line with GL 1320 at month end.

### 7.4 Collateral and repossession (Stage 4)
- **Register collateral** (BranchManager, HeadOffice): C13 customer, C22 type (LAND, HOUSE, VEHICLE, GOLD, MACHINERY, STOCK, EQUIPMENT), C21 description, C7 value, C15 location. A value above `AppraisalSecondOpinionLimit` needs a second user's approval.
- **Appraise**: C29 collateral, C7 value, C15 purpose.
- **Pledge**: C28 loan, C29 collateral (must be Free and the customer's; the loan Applied, Approved or Active). A pledged item cannot be pledged again.
- **Release**: C28 (a Closed loan). Normally automatic when the loan closes.
- **Revalue collateral**: re-appraises items whose appraisal is older than their type's `RevalueMonths` with the simulated market movement; lists LTV alerts (loan principal / pledged value above `MaxLTV + LtvAlertMargin`).
- Lendable value = value x min(type MaxLTV, product MaxLTV). Fair value less costs = value x (1 - Haircut) x (1 - SellingCostPct).
- **Default notice** (BranchManager): C28 loan, allowed when the loan is `RepoNoticeDays` past due.
- **Repossess** (HeadOffice, approval): C28 = RepoID. Allowed when the notice is `RepoNoticePeriodDays` old, the loan is still overdue and a Repossession appraisal is `AppraisalValidDays` old or less. Carrying value CV = min(fair value less costs, B) where B = principal + recognized interest.
- **Sell asset**: C28 repo, C7 price, C24 selling costs, C20 buyer, C27 mode, C8 branch. A price more than `SaleDiscountLimit` below fair value needs approval. Surplus over the claim is paid to the customer's account (refused if the customer has no active account).
- **Start rental / End rental**: C28 repo, C20 tenant, C7 monthly rent. **Record asset op**: C28, C15 type `Rent`, `Revenue` or `Cost`, C7, C27. **Use in business**: C28.
- **Redeem asset** (HeadOffice, approval): C28 repo, C7 = the full remaining claim, C27, C8.
- Month end: holding cost = carrying value x `MonthlyHoldingCostPct` as an expense; re-appraisal list; items older than `MaxHoldingMonths` show "Must sell".

### 7.5 Staff and payroll (Stage 5)
- **Generate sample staff** (Admin, only when `tblEmployees` is empty): creates EXACTLY the staffing-guide minimum for every branch: 242 employees (head office 67, each city branch 16, each town branch 9). Grade 3 for GM, grade 2 for BM, ABM and HOD, grade 1 for the rest. Paid externally, 0 dependents.
- **Hire** (HeadOffice, Admin): C20 name, C31 position, C32 grade, C8 branch, C33 hire date, C10 dependents, C5 bank account (blank = paid externally), C15 phone.
- **Transfer employee**: C30, C8 new branch, C33 date. **Change position**: C30, C31, C32. **Change grade**: C30, C32. **Resign** / **Terminate**: C30, C33. Every change writes `tblEmployeeHistory` and an audit row; a leaving employee's linked user is set inactive; a warning appears when a branch is left without a manager.
- **Attendance** (typed by hand, `tblAttendance`): unpaid leave days, overtime hours, holiday overtime hours, an adjustment (taxable, positive or negative). No row = full pay, no overtime.
- **Staff advance** (BranchManager, HeadOffice; approval): C30 employee, C7 amount, C24 monthly recovery. One Active advance per employee, recovered through payroll.
- **Prepare payroll** (HeadOffice, Admin): C25 period YYYYMM, C8 one branch (blank = ALL). Creates a Draft run and Draft payslips (no ledger) and a Payroll approval. Refused for a period later than the current month or one already covered by an active run.
- **Approve** the run (another HeadOffice user): the run is Posted (ledger: salaries, employer SSB, net pay payable, payroll tax payable, advance recovery).
- **Pay payroll** (HeadOffice; approval): C34 = the Posted run. Account-mode employees are credited in their accounts (statement `SalaryCredit`); the rest are paid through head office.
- **Reverse** a Posted run before payment with Reverse (HeadOffice approval). A Paid run is never reversed: correct it next month with an attendance Adjustment.
- Payslip calculation (period P): Base from the salary scale (latest `EffectiveFrom` on or before the last day of P); hire or exit month is pro-rated; `UnpaidDeduction = Base / WorkDaysPerMonth x days`; allowances by rule (All, BranchType, Position; Flat or percent of base; taxable or not); overtime `hours x Base / (WorkDaysPerMonth x HoursPerDay) x multiplier`; `Gross = Base - Unpaid + allowances + overtime + Adjustment`; `EmployeeSSB = min(Gross, SSBMonthlyCap) x EmployeeSSBRate`; `EmployerSSB` likewise; taxable income = Gross minus non-taxable allowances minus SSB; annual tax from brackets after `DependentReliefAnnual` per dependent; net = Gross - SSB - tax - advance recovered.
- Worked example (city teller, grade 1): base 450,000; taxable allowances 85,000; non-taxable 50,000; gross 585,000; SSB 6,000; tax 6,450; net 572,550; employer SSB 9,000.
- Reports: `Payslip` (C30, C25), `PayrollSummary`, `StaffCostByBranch`, `StaffingGaps`, `Headcount`, `AdvanceRegister`, `SalaryScaleReport`.

### 7.6 Foreign currency (Stage 6)
- **Set rates** (HeadOffice, Admin): C35 currency, C36 mid. Writes the rate row of the business date: Buy = Mid x (1 - `FxSpread`), Sell = Mid x (1 + `FxSpread`), 2 decimals. With `AutoRates = Yes`, end of day creates tomorrow's rates by a random move of up to `RateMoveMax`.
- **Foreign accounts**: open with the products USD01, USD02, JPY01, THB01, SGD01. Deposits, withdrawals, transfers (same currency only), interest and fees work in the account currency. A foreign cash withdrawal needs foreign cash in the branch vault (GL 1700); get it with a cash shipment (C35) or a market buy.
- **Buy currency / Sell currency** (counter): C35 currency, C7 foreign amount, C27 `Cash` or `Account`, C8 branch. The bank sells at Sell and buys at Buy.
- **Convert currency**: C5 MMK account, C6 foreign account (the order sets the direction), C7 foreign amount, C42 purpose.
- A deal at or above `FxLargeLimit` (MMK value) needs a purpose and approval (`LargeFx`). A deal that pushes the net open position of a currency beyond `FxPositionLimitPct` of capital needs HeadOffice approval (`FxLimit`).
- **Market buy / Market sell** (HeadOffice): C35, C7. Used to supply foreign cash and close the open position.
- **Send international**: C5 account, C7 foreign amount, C35 currency (MMK accounts), C39 beneficiary, C40 foreign bank, C41 beneficiary account, C42 purpose. Checks: KYC >= 2 (3 at or above `IntlKyc3Amount`); the day's total under `IntlDailyLimit` (else approval); purpose given; the foreign bank exists and its currency matches; the beneficiary name is screened. A hit sets the payment `Blocked`, posts nothing and opens an AML alert. Head office foreign holdings must cover it ("Not enough foreign funds at head office. Buy currency first."). At end of day a due payment becomes `Failed` (fully refunded) by chance, else `Delivered`.
- **Receive international** (HeadOffice): C5 account, C7 amount, C39 sender, C40 foreign bank, C42 purpose. Screened. **Return international** (HeadOffice): C12 = IntlPaymentID (a Delivered one).
- **Wallets**: **Open wallet** C13 (one per customer, product WLT01). **Top-up** C5 wallet, C6 bank account, C7. **Cash-out** C5 wallet, C6 bank account, C7. **Wallet transfer** C5 from, C6 to, C7. **Wallet pay** C5 wallet, C37 merchant, C7. **Add merchant** (HeadOffice, Admin): C13 customer, C20 name, C21 category, C8. **Add agent**: C13 customer, C20 name. **Agent float in / out**: C38 agent, C7, C8 branch. **Agent cash-in**: C38 agent, C5 customer wallet, C7. **Agent cash-out**: C38 agent, C5 customer wallet, C7.
- Wallet limits by KYC level (`tblWalletLimits`): level 1: max balance 500,000, daily spend 300,000, monthly volume 2,000,000, one transaction 200,000; level 2: 5,000,000 / 2,000,000 / 20,000,000 / 1,000,000; level 3: 30,000,000 / 10,000,000 / 100,000,000 / 5,000,000. Over a limit: refused with the limit shown. Agent wallets cannot go below 0 and are exempt from the limits.
- Daily (end of day): revaluation at the mid rate of the day, then the FX position per currency (`tblFxPosition`).

### 7.7 Investments (Stage 7)
- **Add investee** (HeadOffice): C20 name, C21 sector (Agriculture, Manufacturing, Trade, Services, RealEstate, Energy, Technology), C32 risk grade A to E.
- **Propose investment** (HeadOffice): C44 investee, C46 type (`Equity`, `ProfitShare`, `FixedReturnNote`, `TradeFinance`), C7 amount, C10 term (months; notes and trade finance), C47 rate or share (0.14 or 14), C15 pay frequency. Limits after the investment: total <= `Total` limit x capital (3000 + 3100); one investee <= `Single`; sector <= its limit; grades D and E together <= `HighRisk`; liquid funds >= `ReserveRatio x deposits + InvestLiquidityBuffer`. Creates an approval by another HeadOffice user.
- **Fund investment**: C45. HO: Dr 1600, Cr 1100. Status Active.
- Daily: notes and trade finance accrue `Amount x Rate / 365` (memo). Month end: the accrual is posted (Dr 1610, Cr 4300). Maturity: with `InvestAutoRepay = Yes` the simulated investee repays principal and returns; otherwise the note shows on `InvestmentAlerts` after 30 days.
- **Receive return**: C45, C7, C15 type. **Record profit share**: C45, C48 reported profit (negative = loss). **Value investment**: C45, C48 new value, C15 method. **Dispose** / **Write off** (HeadOffice, approval): C45 (and C7 proceeds for dispose).
- **Simulate investee month** (Admin): random monthly outcomes by risk grade, a default sets the investee Distressed.
- Reports: `InvestmentBook`, `Concentration`, `InvestmentIncome`, `InvestmentAlerts`, `InvestmentPnL`.

### 7.8 Policy rate (Stage 8)
Set `PolicyRate` in Config (a number you type, for example 0.12). Press **Apply policy rate** (HeadOffice, approval by another HeadOffice user). `Delta = PolicyRate - PolicyRateApplied`. For each deposit rate tier whose product `FollowsPolicy = Yes`, a NEW `tblRateTiers` row is added with today's date and `AnnualRate = clamp(old + Delta, MinRate, MaxRate)`; loan product rates change the same way (existing loans keep their locked rate, fixed deposits keep theirs). An audit row lists every change.

### 7.9 Government securities (Stage 8)
**Buy gov security** (HeadOffice, approval): C50 `TBill` or `GovBond`, C51 face value, C52 term in days, C53 rate. T-bill price = Face x (1 - Rate x Days / 365). Example: face 100,000,000, rate 0.08, 91 days: price 98,005,479. Bonds are bought at par and pay a coupon every 6 months. Securities count as liquid assets at carrying value and carry a 0% risk weight. Month end: accretion or coupon accrual (4310). Coupon dates and maturity are handled at end of day. Report: `GovSecurities`.

### 7.10 Compliance ratios (Stage 8)
**Compliance check** (HeadOffice, Admin, Auditor, Regulator) computes every ratio and writes `tblRatioHistory`; end of day does the same automatically. A change from Within to Breach writes an event `Breach`; back to Within writes `Cleared`. Breaches never stop the day; they show on `RatioReport`, `ComplianceEventsReport` and the Dashboard.
- `RESERVE` = (vault cash + 1100 + 1620) / (`ReserveRatio` x deposits); breach below 1.
- `CAR` = capital / risk-weighted assets; breach below `MinCapitalRatio`. Capital = 3000 + 3100 + profit so far this year. Risk weights come from `tblRiskWeights` (loans by security and class).
- `LCR` = liquid assets / (30-day outflows - capped inflows); breach below `MinLiquidityRatio`. Outflows use `tblRunoffRates`.
- `NPL` = principal in class 3, 4, 5 / total loan principal; breach above `MaxNplRatio`.
- `FX_POSITION` (worst currency net open position / capital), `SINGLE_BORROWER` (largest customer exposure / capital), `INVEST_TOTAL` (investments / capital) against their limits.

### 7.11 Deposit insurance (Stage 8)
At every quarter-end month end the bank books a premium: for each customer `Insured = min(deposit and wallet balances in MMK, DepositInsuranceCap)`; premium = total insured x `DepositInsuranceRate` / 4. It is an expense row (GL 5340) unpaid until Pay bills. Idempotent per quarter (description `Deposit insurance <year>Q<n>`). Report: `DepositInsurance`.

### 7.12 Income tax (Stage 8)
- **Tax return** (HeadOffice): C25 year. Writes the `TaxReturn` sheet: accounting profit before tax (income minus expenses, excluding 5400 and closing journals) + `tblTaxAdjustments` of that year + the regulatory fines added back = taxable profit; tax = taxable profit x the `INCOME_TAX` rate (0 if negative); plus the other taxes withheld, charged and paid.
- **Book income tax** (HeadOffice, approval): C25 year. HO: Dr 5400, Cr 2220, once per year. Pay it with Pay taxes.

### 7.13 Inspection and fines (Stage 8)
- **Inspection** (Regulator, HeadOffice): runs every control check and every ratio; lists breaches, unreviewed AML alerts older than `AmlReviewDays`, Pending approvals older than 3 days, and branches below `VaultMin`. Writes `tblInspections` and the `Inspection` sheet (`Pass` when there are no findings).
- **Issue fine** (Regulator proposes; a different HeadOffice user approves): C7 amount, C54 reason. HO: Dr 5350, Cr 1100. The tax return adds fines back.

### 7.14 Anti money laundering (Stage 8)
- End of day scans the day's transactions: `LARGE_CASH` (any single cash transaction at or above the threshold, default 20,000,000), `STRUCTURING` (3 or more cash deposits within 3 days, each from 80% of the threshold up to just below it, one alert per set), `VELOCITY` (credits in one day for a KYC 1 or 2 customer reaching 100,000,000), `BLOCKED_NAME` (every screening hit).
- **Name screening**: a customer name that contains an active entry of `tblBlockedNames` (case-insensitive) is refused at account opening and at loan application. A HeadOffice user goes on by writing `OVERRIDE: reason` in C15 (the hit is still written as an alert).
- **Review alert** (HeadOffice): C49 alert, C15 `ClosedNoAction` or `Escalated`. Escalation opens a case assigned to the reviewer. **Open case**: C49, C15 summary. **Close case**: C12 case, C15 decision. **Report case** (approval by another HeadOffice user): C12 case; a Reported case appears on `SuspiciousReport`. No money is ever frozen automatically; a BranchManager may freeze an account by hand.
- Reports: `LargeCashReport` (C16 to C17), `SuspiciousReport`, `AmlAlertsReport`.

### 7.15 Year-end close (Stage 8)
**Year-end close** (Admin requests, a different HeadOffice user approves): C25 year (blank = last year). Allowed only when: the business date is after 31 December of that year; `LastClosedYear` is the previous year or empty; all control checks pass; the income tax of the year is booked (when there is tax). Warnings are written for months without a payroll run. One journal per branch (type `YearEndClose`) zeroes every income and expense account of the year and moves the result to 3100. `LastClosedYear` is set and the audit row shows profit before and after tax. Closing journals are excluded from every profit and loss, branch P&L, tax return and ratio calculation, so a closed year still shows its income and expenses. A journal inside a closed year cannot be reversed.
Annual reports: `IncomeStatement` (year C25), `BalanceSheet` (current date), `CashFlowSummary` (C16 to C17), `BranchRanking` (profit, cost-to-income, deposits, loans, NPL, staff cost, ranked), `Dashboard`.

---

## 8. End of day (Admin only)

Press **End of day** with an Admin user in C4. It closes the current business date D and moves the date to D + 1. Order:
1. Stop with a message if D already exists in `DailyClose`. Warn (do not stop) about Pending approvals.
2. Decide due interbank transfers and international payments. Mature fixed deposits.
3. Loan steps: installments due, auto-debit, penalties, classes, closing.
4. Investment maturities and accruals; government bond coupons and maturities.
5. Accrue deposit interest (all currencies). Mark Dormant accounts.
6. If D is the last day of the month: deposit interest posting and maintenance fees; branch costs and depreciation; loan provisions; holding costs of repossessed assets; investment return accruals; T-bill accretion and bond coupon accrual; on a quarter end the deposit insurance premium. Warnings for a missing payroll run or a branch without a manager.
7. FX revaluation at the mid rate of D and the FX position snapshot; rates for D + 1 when `AutoRates = Yes`.
8. After the day is committed: compliance ratios and the AML scan of the day.
9. Run ALL control checks. **If any check fails, everything from this run is rolled back (every touched table is restored), the failure is written to the Log and the date does not move.**
10. Write a `DailyClose` row and an audit row; set `BusinessDate = D + 1`.

Running it twice for the same date is refused. Run it repeatedly to move through the month; month end happens on the last day of each month.

---

## 9. Control checks (sheet `Checks`, also run by end of day)

| # | Check |
|---|---|
| 1 | Every journal: total debit = total credit |
| 2 | Every journal, every branch: debit = credit |
| 3 | GL 1900 nets to 0 over all branches |
| 4 | Every account: Balance = sum of its Transactions rows = sum of its journal lines; each deposit GL = total of balances |
| 5 | GL 2900 = total of Submitted outgoing transfers |
| 6 | No branch vault below 0 (warning if below `VaultMin`) |
| 7 | GL 1500 = cost of assets in use; GL 1510 = accumulated depreciation |
| 8 | GL 2300 = unpaid cost postings + unpaid expenses |
| 9 | No cost posted twice; depreciation never above cost less residual |
| 10 | GL 1300 = principal outstanding of Active loans |
| 11 | GL 1310 = recognized unpaid interest of Active loans |
| 12 | GL 1320 = provisions of Active loans (a WARN between month ends, strict at month end) |
| 13 | Loan schedules add up to the principal |
| 14 | GL 1400 = carrying value of repossessed assets held |
| 15 | No item has two liens; every Pledged item backs an open loan |
| 16 | Every repossessed or sold item has a record; every closed record has a sale |
| 17 | Payslips equal their runs; one payslip per employee and period; GL 2400 = unpaid net pay |
| 18 | GL 1810 = staff advances outstanding; GL 2230 payroll credits = payroll tax and SSB of posted runs |
| 19 | Foreign accounts: balance = foreign amounts of lines = statement rows; per currency the balances = the foreign deposit lines |
| 20 | After revaluation every (branch, GL, currency) group has base = RoundCcy(foreign x mid) |
| 21 | Foreign tags only on GL 1700, 2000, 2010, 2030, 2200, each with ForeignAmount and FxRate |
| 22 | GL 2030 = GL 1110 = sum of wallet balances; no wallet above its KYC maximum |
| 23 | No branch has negative foreign cash; no agent wallet is negative |
| 24 | GL 1600 = carrying value of investments not yet repaid |
| 25 | GL 1610 = recognized unpaid investment returns + bond coupons |
| 26 | GL 1620 = carrying value of held government securities |
| 27 | Every asset GL with a balance has a risk weight |
| 28 | Closed years: income and expense accounts net to 0 per branch counting the closing journals; 3100 rose by the year's profit after tax |
| 29 | Every AML alert has a valid rule; every Open case has an assignee; alerts still New after `AmlReviewDays` are a warning |

---

## 10. Reports (generated sheets, cleared and rewritten by **Reports**)

- **Core:** `TrialBalance` (by GL, then by branch and GL), `BranchSummary`, `CashPosition`, `DepositSummary`, `Checks`, `Statement` (C5, C16, C17).
- **Stage 2:** `BranchPL`, `CostAnalysis`, `AssetRegister`, `BillsDue`, `TaxSummary`, `BudgetVsActual`.
- **Stage 3:** `LoanBook`, `LoanAgeing`, `LoanDue`, `LoanStatement`, `LoanLimits`, `LoanIncome`, `LoanQuote`.
- **Stage 4:** `CollateralRegister`, `CollateralAlerts`, `RepossessedAssets`, `SaleResults`, `AssetPnL`, `Deficiencies`.
- **Stage 5:** `Payslip`, `PayrollSummary`, `StaffCostByBranch`, `StaffingGaps`, `Headcount`, `AdvanceRegister`, `SalaryScaleReport`.
- **Stage 6:** `FxRates`, `FxPositionReport`, `FxIncome`, `CurrencyCash`, `FxDealsReport`, `IntlPaymentsReport`, `WalletSummary` (with the trust reconciliation), `MerchantSummary`, `AgentSummary`.
- **Stage 7:** `InvestmentBook`, `Concentration`, `InvestmentIncome`, `InvestmentAlerts`, `InvestmentPnL`.
- **Stage 8:** `Dashboard`, `RatioReport`, `ComplianceEventsReport`, `Inspection`, `TaxReturn`, `DepositInsurance`, `LargeCashReport`, `SuspiciousReport`, `AmlAlertsReport`, `GovSecurities`, `IncomeStatement`, `BalanceSheet`, `CashFlowSummary`, `BranchRanking`.
Report sheets whose natural name equals a table sheet carry a `Report` suffix.

The Dashboard shows deposits by type, loans, loan-to-deposit, liquid assets, the ratios, wallet float, investments, net interest margin (on current balances), cost-to-income, profit year to date, staff count, open breaches, open AML alerts, pending approvals and the top and bottom 3 branches by profit.

---

## 11. Data tables (Excel Tables; the sheet name is the table name without `tbl`)

Typed = you type it; Macro = only macros change it.

**Stage 1**
- `tblConfig` (Typed): Key, Value.
- `tblUsers` (Typed): UserID, Name, Role, BranchID, IsActive.
- `tblBranches` (Typed): BranchID, Name, Type, City, Region, PremisesType, OpenDate, Status, VaultMin, VaultMax.
- `tblCustomers` (Typed): CustomerID, Name, CustomerType, Phone, Address, HomeBranchID, KYCLevel, TaxExempt, JoinDate, Status.
- `tblProducts` (Typed): ProductCode, Name, Type (Savings, Current, Fixed, Wallet), Currency, MinBalance, MinOpening, MaintenanceFeeCode, IsActive.
- `tblRateTiers` (Typed): RateID, ProductCode, TermMonths, FromAmount, ToAmount, AnnualRate, EffectiveFrom.
- `tblFees` (Typed): FeeCode, Name, Method, FlatAmount, Rate, MinFee, MaxFee, Currency, IncomeGL, Taxable.
- `tblOtherBanks` (Typed, simulated): BankCode, Name, SettleDays, FailRate, ReturnRate.
- `tblChart` (Typed): GLCode, Name, Type, NormalSide, Stage, IsActive.
- `tblAccounts` (Macro): AccountNo, CustomerID, BranchID, ProductCode, Currency, Status, OpenDate, TermMonths, MaturityDate, LockedRate, PayoutAccountNo, Balance, AccruedInterest, LastAccrualDate, LastActivityDate, ClosedDate.
- `tblTransactions` (Macro, the account statement): TxnID, GroupID, TxnDate, AccountNo, TxnType, Direction, Amount, BalanceAfter, CounterAccountNo, BranchID, Description, Reference, JournalID, Status, CreatedBy, ApprovedBy.
- `tblInterbankTransfers` (Macro): TransferID, Direction, AccountNo, OtherBankCode, OtherBankAccount, Amount, Fee, SubmittedDate, ExpectedSettleDate, Status, StatusDate, Reason, SubmitJournalID, SettleJournalID, ReturnJournalID.
- `tblApprovals` (Macro): ApprovalID, RequestType, AccountNo, CounterAccountNo, Amount, BranchID, Note, RequestedBy, RequestedAt, Status, ApprovedBy, DecidedAt, DecisionNote, TxnGroupID.
- `tblJournal` (Macro): JournalID, PostDate, Description, JournalType, GroupID, Status, ReversalOf, CreatedBy, CreatedAt.
- `tblJournalLines` (Macro): LineID, JournalID, BranchID, GLCode, AccountNo, Currency, Debit, Credit, FxRate, ForeignAmount.
- `tblDailyClose` (Macro): BusinessDate, RunAt, AccountsAccrued, InterestAccrued, InterbankSettled, FDMatured, DormantMarked, ChecksPassed, Notes, CostsPosted, DepreciationPosted, InstallmentsDue, AutoDebited, PenaltyAccrued, ProvisionPosted, FxRevaluation, IntlDecided, InvestmentAccrued, RatiosBreached, AmlAlerts.
- `tblAudit` (Macro): AuditID, Timestamp, UserID, Action, Target, Detail. `tblLog` (Macro): Timestamp, Level, Procedure, Message.
- `tblCurrencies` (Typed): Code, Name, Decimals, IsBase. `tblExchangeRates` (Typed/Macro): RateDate, Currency, BuyRate, SellRate, MidRate. `tblTaxRates` (Typed): TaxCode, Name, Rate, AppliesTo, GLCode, EffectiveFrom (WHT_INTEREST, COMMERCIAL_FEES, INCOME_TAX, EMPLOYEE_SSB, PAYROLL_TAX; sample values).

**Stage 2:** `tblBranchCosts` (Typed): CostID, BranchID, CostType, Description, Frequency, Amount, GLCode, StartDate, EndDate, Status. `tblFixedAssets`: AssetID, BranchID, AssetType, Description, AcquireDate, Cost, UsefulLifeMonths, ResidualValue, DepMethod, AccumDepAt20260930, Status, AccumDepreciation, LastDepPeriod, LoadedOn, DisposedOn, DisposalProceeds. `tblCostPostings`, `tblExpenses`, `tblTaxPayments`, `tblBudgets` (columns in the posting rules above and section 7.2).

**Stage 3:** `tblLoanProducts` (LN01 Personal, LN02 Vehicle, LN03 Business, LN04 Home, LN05 Gold, LN06 Agricultural: ProductCode, Name, LoanType, AnnualRate, Method, MinAmount, MaxAmount, MinTermMonths, MaxTermMonths, ProcessingFeePct, GraceDays, PenaltyRate, CollateralRequired, MaxLTV, EarlyPayoffFeePct, IsActive), `tblLoanClasses` (ClassCode, Name, MinDaysPastDue, MaxDaysPastDue, ProvisionPct), `tblLoans`, `tblLoanSchedule`, `tblLoanPayments`.

**Stage 4:** `tblCollateralTypes` (TypeCode, Name, MaxLTV, Haircut, AnnualValueChange, RevalueMonths, SellingCostPct, MonthlyHoldingCostPct), `tblCollateral`, `tblAppraisals`, `tblRepossessions`, `tblAssetSales`, `tblAssetRentals`, `tblAssetOps`.

**Stage 5:** `tblPositions` (18 positions with the staffing guide), `tblSalaryScales` (grades 1 to 3), `tblAllowanceRules`, `tblTaxBrackets` (0 to 4.8M 0%; to 10M 5%; to 20M 10%; to 30M 15%; above 20%), `tblEmployees`, `tblEmployeeHistory`, `tblAttendance` (typed), `tblStaffAdvances`, `tblPayrollRuns`, `tblPayslips`.

**Stage 6:** `tblForeignBanks` (5 simulated banks FB01 to FB05), `tblBlockedNames` (5 made-up names), `tblWalletLimits`, `tblFxDeals`, `tblFxPosition`, `tblIntlPayments`, `tblMerchants`, `tblAgents`.

**Stage 7:** `tblInvestmentLimits`, `tblInvestees`, `tblInvestments`, `tblInvestmentReturns`, `tblInvestmentValuations`.

**Stage 8:** `tblRateRules`, `tblRiskWeights`, `tblRunoffRates`, `tblAmlRules`, `tblTaxAdjustments` (typed), `tblBonds`, `tblAmlAlerts`, `tblAmlCases`, `tblRatioHistory`, `tblComplianceEvents`, `tblFines`, `tblInspections`.

Reference tables are seeded by `RunSetup` only when empty. Hidden sheet `Lists` holds long dropdown lists.

---

## 12. Config keys (sheet `Config`, typed by hand)

Stage 1: BankName, BaseCurrency, BusinessDate (the date macros move), PolicyRate, WithholdingTaxRate (0.10), TellerLimit (5,000,000), BranchManagerLimit (50,000,000), DailyWithdrawalLimit (20,000,000), DormantDays (365), EarlyBreakRate (0.03), DayCountBasis (365), CashShipmentRounding (10,000,000), ReserveRatio (0.05).
Stage 2: ExpenseApprovalLimit 2,000,000; PettyCashLimit 500,000; BudgetUplift 0.05.
Stage 3: LoanBranchLimit 30,000,000; LoanMinKYC 2; LoanKyc3Amount 50,000,000; MaxLoanToDeposit 0.80; SingleBorrowerLimitPct 0.05; LoanAutoDebit Yes; RequireCollateralCheck No (set Yes when you use collateral).
Stage 4: RepoNoticeDays 90; RepoNoticePeriodDays 30; MaxHoldingMonths 24; LtvAlertMargin 0.10; AppraisalValidDays 30; SaleDiscountLimit 0.10; HeldAssetRevalueMonths 3; AppraisalSecondOpinionLimit 100,000,000.
Stage 5: WorkDaysPerMonth 26; HoursPerDay 8; OvertimeMultiplier 1.5; HolidayOvertimeMultiplier 2.0; EmployeeSSBRate 0.02; EmployerSSBRate 0.03; SSBMonthlyCap 300,000; DependentReliefAnnual 1,000,000.
Stage 6: FxSpread 0.01; MarketSpread 0.002; AutoRates Yes; RateMoveMax 0.004; FxLargeLimit 10,000,000; FxPositionLimitPct 0.10; IntlDailyLimit 50,000,000; IntlKyc3Amount 20,000,000; AgentCommissionIn 0.001; AgentCommissionOut 0.003.
Stage 7: InvestLiquidityBuffer 500,000,000; InvestAutoRepay Yes.
Stage 8: MinCapitalRatio 0.10; MinLiquidityRatio 1.00; MaxNplRatio 0.05; DepositInsuranceRate 0.001; DepositInsuranceCap 10,000,000; AmlReviewDays 7; LastClosedYear (set by the year-end close); PolicyRateApplied (set by Apply policy rate).
A change of any Config key writes an audit row when done through a macro; typing in the sheet is allowed for master data.

---

## 13. Design and rules for the code

- Output only `.bas` and `.cls` files (plus this README): no UserForms, no binaries. The UI is the `Control` sheet with buttons that call entry macros. Code is ASCII only, CRLF line endings. Every `.bas` starts with `Attribute VB_Name` then `Option Explicit`; every `.cls` starts with the full class header then `Option Explicit` (`clsLogger` and `clsErrorHandler` use `VB_PredeclaredId = True`).
- All public names are unique across the project and prefixed by layer (`svc`, `repo`, `cls`, `mod`), always called with the module name (`svcLedger.PostJournal`).
- Money is `Decimal` (`CDec`), never Double or Single, rounded with ONE helper `modMoney.RoundCcy(amount, currency)` (round half up to the currency decimals, MMK 0). Random draws use `Rnd`; no money does.
- No `Select`, `Activate`, `ActiveSheet`, `ActiveCell`, `Selection`. Sheets and tables are referenced by name through `modConstants`. No cell-by-cell loops over table data: read a table into a 2D array, work in memory (`Scripting.Dictionary` for lookups), write back in one assignment.
- The business date comes from `repoConfig.BusinessDate`. The computer clock is used only for `CreatedAt` and Log time stamps.
- No automated tests and no test modules. No deleting posted records.
- Layers call downward only:
  ```
  Entry (modMain*)  ->  Services (svc*)  ->  Repositories (repo*)  ->  Infrastructure (clsLogger, clsErrorHandler, modSheetUtils, modPerformance, modMoney, modDateUtils, modIds, modArgs)
  Models (cls*) are plain data holders.
  ```
  Entry macros only orchestrate (current user, error handling, one service call, one message). Services never touch worksheets and never show a message box; they return a `clsResult` (Success, Message, Count, RefID).
- Every entry macro follows the same pattern: `On Error GoTo ErrHandler`, `modPerformance.TurnOff`, call the service, show the message, cleanup `modPerformance.TurnOn`.
- Modules: Entry `modMain`, `modMainCosts`, `modMainLoans`, `modMainCollateral`, `modMainStaff`, `modMainFx`, `modMainInvest`, `modMainGov`. Config `modConstants`, `modMessages`, `modSetup`, `modSetupStages`, `modSetupLate`. Core services `svcLedger`, `svcSecurity`, `svcFees`, `svcAccounts`, `svcApproval`, `svcCashOps`, `svcTransfers`, `svcInterbank`, `svcFixedDeposit`, `svcInterest`, `svcEndOfDay`, `svcChecks`, `svcReports`, with extension modules `svcApprovalExtra`, `svcEndOfDayExtra`, `svcChecksExtra`, `svcReportsExtra`. Stage 2: `svcCosts`, `svcExpenses`, `svcFixedAssets`, `svcTaxes`, `svcBudgets`, `svcReportsCosts`. Stage 3: `svcLoanCalc`, `svcLoans`, `svcLoanPayments`, `svcLoanEod`, `svcReportsLoans`. Stage 4: `svcCollateral`, `svcRepossession`, `svcAssetOps`, `svcReportsCollateral`. Stage 5: `svcStaff`, `svcPayroll`, `svcReportsStaff`. Stage 6: `svcFx`, `svcFxDeals`, `svcFxReval`, `svcIntl`, `svcWallet`, `svcChecksFx`, `svcReportsFx`. Stage 7: `svcInvest`, `svcReportsInvest`. Stage 8: `svcPolicy`, `svcRatios`, `svcAml`, `svcBonds`, `svcInsurance`, `svcTaxReturn`, `svcInspection`, `svcYearEnd`, `svcChecksLate`, `svcReportsGov`. Repositories: `repoConfig`, `repoUser`, `repoBranch`, `repoCustomer`, `repoProduct`, `repoRate`, `repoFee`, `repoChart`, `repoAccount`, `repoTransaction`, `repoJournal`, `repoInterbank`, `repoApproval`, `repoOtherBank`, `repoClose`, `repoAudit`, `repoTaxRate`, `repoCost`, `repoAsset`, `repoLoan`, `repoCollateral`, `repoStaff`, `repoFxRate`, `repoFx`, `repoInvest`, `repoGov`. Models and infrastructure: `clsAccount`, `clsJournalLine`, `clsTxnRow`, `clsApproval`, `clsResult`, `clsUser`, `clsBranch`, `clsProduct`, `clsCustomer`, `clsFee`, `clsOtherBank`, `clsUnitOfWork`, `clsLogger`, `clsErrorHandler`, `modSheetUtils`, `modPerformance`, `modMoney`, `modDateUtils`, `modIds`, `modArgs`.
- `clsUnitOfWork` holds all in-memory changes of one action (accounts, journals, statement rows, other tables, foreign groups) and writes them once with `Commit`. End of day builds the whole day in one unit, commits, runs the checks, and restores every touched table from a snapshot if a check fails.

### Decisions and simplifications
- Reactivating a Dormant account also sets its `LastActivityDate` to the business date, otherwise the next end of day would mark it Dormant again.
- Interest on fixed deposits is posted only at maturity or break (not expensed month by month).
- Loan interest income is recognized on each due date. Loan provisions are brought in line with GL 1320 at month end. Loan due dates are counted from the disbursement date.
- A one-off expense above `ExpenseApprovalLimit` creates only an approval; the `tblExpenses` row is written when it is approved.
- Name screening at account opening and loan application: a HeadOffice user goes on with `OVERRIDE: reason` in C15 (no separate approval row).
- Bond coupon receivable (GL 1610) is derived from dates, not stored; bond coupons are paid as accrued; T-bill accretion is straight line.
- The Dashboard net interest margin uses current balances, not averages.
- The quick-reference note: `RunSetup` can run without a chosen user on a fresh workbook; once a user is chosen it needs an Admin.

---

## 14. Test checklist (numbers are for the starter workbook)

1. Compile has no error. `RunSetup` keeps all rows and creates the Control sheet.
2. `Reports`: the `Checks` sheet passes. Trial balance: 1000 = 4,714,385,127; 1100 = 18,772,303,480; 2000 = 1,476,938,571 credit; 2010 = 1,315,667,248 credit; 2020 = 693,575,910 credit; 2900 = 427,367 credit; 3000 = 20,000,000,000 credit; 4100 = 79,511 credit. GL 1900 = 0.
3. Deposit 100,000 into an account at its own branch, then at another branch: both post, GL 1900 stays 0.
4. Withdraw at another branch: the cross-branch fee is charged and the income appears at the processing branch.
5. Withdraw more than the balance, or below the minimum: refused. Withdraw from a frozen account: refused. Deposit to a frozen account: allowed.
6. Transfer inside a branch (no fee) and between branches (fee `TRANSFER_BRANCH`).
7. A deposit above 5,000,000 as a Teller creates a Pending approval and posts nothing. Approve as the same user: refused. Approve as the branch manager: posts.
8. Interbank out: Submitted, GL 2900 rises. Interbank in: credited at once.
9. A cash shipment needs a HeadOffice approver different from the requester.
10. Reverse a deposit (HeadOffice approval): the original becomes Reversed, the balance is back, the checks pass.
11. First end of day (13 Oct 2026): the fixed deposits `0060000398` (gross interest 214,635, tax 21,464, net 193,171) and `0130000399` (gross 143,269, tax 14,327, net 128,942) mature and pay their payout accounts; `0120000122` and `0020000330` become Dormant; interbank `IB000035` is decided; the date becomes 14 Oct. Running it again for 13 Oct is refused.
12. Run end of day to 31 Oct: month end posts interest with 10% tax (2200 rises) and maintenance fees for current accounts below the minimum balance. After the first run the `Checks` sheet still shows no FAIL.
13. Break a fixed deposit early; close an account (interest paid, cash out, Closed).
14. Stage 2: withdraw 1,000,000 at another branch: the customer pays 1,001,050. Load fixed assets (89 assets, 1500 = 17,590,000,000). Post October costs: 153,890,249 in 150 rows. Pay bills: 2300 returns to 0.
15. Stage 3: quote LN01 10,000,000 for 12 months: installments 983,333 (x11) and 983,337, total interest 1,800,000. A declining loan of 10,000,000 at 12% for 12 months: payment 888,488, total interest 661,853.
16. Stage 4: a house valued 80,000,000: lendable value for LN04 56,000,000, fair value less costs 65,960,000. The repossession example (principal 50,000,000, interest 1,000,000, provision 12,500,000, re-appraised 40,000,000) gives Dr 1400 32,980,000, Dr 1320 12,500,000, Dr 5300 5,520,000, Cr 1300 50,000,000, Cr 1310 1,000,000.
17. Stage 5: generate sample staff (242). Prepare payroll 202610 for ALL: gross 198,190,000; employee SSB 1,452,000; employer SSB 2,178,000; income tax 8,621,109; net 188,116,891.
18. Stage 6: send USD 2,000 from a USD account: fees USD 19.30 + tax USD 0.97, account debited USD 2,020.27. A beneficiary name containing `Red Harbour` is blocked. A KYC 1 wallet cannot be topped up with 600,000 (maximum 500,000).
19. Stage 7: a 500,000,000 note at 14%: daily accrual 191,780.82, month-end posting 5,945,205 for a 31-day month.
20. Stage 8: a T-bill of 100,000,000 at 8% for 91 days costs 98,005,479. A cash deposit of 25,000,000 creates a `LARGE_CASH` alert. After 31 December, book the income tax and run the year-end close: income and expense accounts are 0 per branch, 3100 rose by the profit after tax, the `IncomeStatement` still shows the year.
21. All 29 checks pass. No leftover `EXCEL.EXE` after closing Excel.

---

## 15. Troubleshooting

- **"Method or data member not found" / "User-defined type not defined" / "Sub or Function not defined"** on compile: a module is missing, duplicated (`clsUnitOfWork1`) or an old version. Delete all imported modules and import the whole `src` folder again.
- **Syntax error on a line using a variable called `single`, `base`, `scale`, `date`**: those are VBA keywords; the shipped code avoids them. Re-import the shipped file.
- **"Run RunSetup" message**: a table or Config key is missing. Run `RunSetup` (Admin) again; it never clears data.
- **"Role X is not allowed to ..."**: choose a user with the right role in C4 (section 4).
- **A message starting with a business rule (balance, vault short, limit)** is normal: the action was refused and nothing was posted.
- **End of day says a check failed**: the whole day was rolled back. Read the failing check in the message (and the `Checks` sheet after `Reports`), fix the cause and run it again.
- **A foreign deal says "No rate for today"**: run `Set rates` for the currency on the business date.
- **Account numbers lose their leading zeros**: the Control cells are text; if you paste a number, retype it as text.
- Check the `Log` sheet for warnings and errors written by the macros.
