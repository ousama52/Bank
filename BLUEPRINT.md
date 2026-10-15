# BLUEPRINT.md - Golden Delta Bank (simulation): the whole plan

One bank, a head office and branches in cities and towns, built in Excel VBA with an agent, one stage at a time.
Stage 1 is specified in `CLAUDE.md` and `DATA_TABLES.md`. This file says what Stages 2 to 8 add, so Stage 1 is built with the right foundations. Each later stage gets its own `CLAUDE.md` section when you reach it.

Everything is a simulation. Names, banks, rates and tax rates are made up. Do not use it for real money, real customers or legal advice.

## Scope decisions

| Decision | Meaning |
|---|---|
| One bank | No central bank, no second bank, no government institution in the books |
| Branches | Head office plus city and town branches. Money moves between branches through inter-branch clearing |
| Other banks | FAKE. One simulated settlement balance and one simulated payable. No other bank keeps books |
| Government | Not a separate party. It appears as tax rates, reserve and capital rules, reports and tax payments out of the bank |
| Central bank | Not a party. Its policy rate is a number in Config |
| Foreign countries | FAKE too: foreign payments use a simulated foreign settlement balance |
| Time | A business date that you control. Every day is a business day |
| Users | One person at a time, switching user and role on the Control sheet |

## The idea that holds it together: one ledger

Every money movement is a journal of debit and credit lines. Each line has a branch. A journal balances overall and inside each branch. Branch positions net out through `1900 Inter-branch clearing`. Because of this:

- a wrong posting shows up at once as a failed control check
- every branch has its own profit and loss and balance sheet
- every stage plugs in by adding new posting rules, not by changing old tables

## The people and parts

| Who or what | In the simulation |
|---|---|
| Customers | People and businesses with accounts, loans, wallets, investments |
| Bank staff | Users with roles: Teller, BranchManager, HeadOffice, Auditor, Admin. Employees and payroll arrive in Stage 5 |
| Branches | Head office, city and town branches, each with a vault, staff, costs and a profit report |
| Treasury (head office) | Holds the money not lent out: `1100`. Funds loans, investments and cash orders |
| Other banks (fake) | Names, settle days, failure rates. Nothing else |
| Government and regulator | Rules and reports (Stage 8) |

## Stages

| Stage | Name | You can do this when it is finished |
|---|---|---|
| 1 | Core bank and ledger | Open accounts, deposit, withdraw, transfer inside and between branches, fake other-bank transfers, fixed deposits, interest, fees, approvals, end of day, reports |
| 2 | Branch running costs, assets and taxes | Every branch pays rent or property tax, utilities, security and more. Buildings and equipment lose value. Tax on fee income. Branch profit and loss |
| 3 | Loans | Several loan types, repayment tables, repayments, late penalties, overdue classification and provisions, lending limits from the bank's own money |
| 4 | Collateral and property | Pledged items, valuation, loan limits, default, taking the item, pricing it again, selling it or renting it out |
| 5 | Staff and payroll | Positions, salary scales, allowances by branch type, overtime, tax and social security, payslips, staff costs per branch |
| 6 | Foreign currency, international payments and digital money | Buy and sell currency, foreign accounts, fake international payments, e-wallets, QR payments, agents |
| 7 | Business investments | The bank invests in businesses, records returns and losses, keeps investment limits |
| 8 | Government and regulator | Reserve and capital rules, bonds, deposit insurance, tax returns, suspicious activity cases, inspections, year-end close |

Each stage works on its own and is checked with its own manual checklist before the next one starts.

---

## Stage 2: branch running costs, assets and taxes

**Goal:** show the real cost of running a branch network in cities and towns.

**Tables (sample data is already in the starter workbook):** `tblBranchCosts`, `tblFixedAssets`, `tblTaxRates`. New: `tblExpenses` (one-off repairs and purchases), `tblBudgets`.

**Rules**
- Month end posts every active cost row of each branch. Monthly rows: the amount. Yearly rows: amount / 12. Debit the expense GL (5200 rent, 5210 electricity, 5211 water, 5212 internet and phone, 5220 maintenance, 5230 security and cleaning, 5240 insurance, 5260 property tax, 5270 supplies, 5280 generator fuel) in the branch, credit `2300 Accrued expenses` in the branch.
- Paying bills: head office pays all bills from treasury. Branch: Dr 2300, Cr 1900. Head office: Dr 1900, Cr 1100.
- A change of price (new electricity tariff) = end the old cost row, add a new one with a start date.
- Depreciation each month for each asset in use: Dr 5250, Cr 1510, at the asset's branch.
- Opening the fixed asset register: Dr 1500 (cost), Cr 1510 (accumulated depreciation), Cr 3000 (net book value) at the branch.
- Owned premises: building asset, depreciation, property tax and maintenance. Rented premises: rent.
- One-off repairs: `RunPostExpense` (branch, GL, amount, supplier, description) needs approval above a limit.
- Tax on fees: commercial tax (5% sample) is charged on top of every fee: Dr customer account (fee + tax), Cr 4100 fee, Cr 2210 tax. This changes the Stage 1 fee posting, so the agent updates the posting rules table and the checklist.
- Paying taxes to the government: `RunPayTaxes` clears 2200, 2210 (and later 2220, 2230): Dr tax payable, Cr 1100 at head office through 1900.
- Budget vs actual per branch and per cost type.

**Reports:** Branch profit and loss, cost per branch type (city vs town), cost per customer, assets register, tax payable.

**Checks:** depreciation never exceeds cost less residual value; every cost row posts once per month; 2300 equals unpaid bills.

---

## Stage 3: loans

**Loan types:** personal, vehicle, business working capital, home (mortgage), gold or item-backed, agricultural. Each has: rate, method, maximum term, maximum amount, processing fee, grace period, late penalty rate, collateral required (Yes/No) and maximum loan-to-value.

**Tables:** `tblLoanProducts`, `tblLoans`, `tblLoanSchedule`, `tblLoanPayments`, `tblLoanClasses`.

**Rules**
- Interest methods: flat rate and declining balance. Monthly payment from term, rate and amount, with the last payment cleaning up rounding.
- Disbursement: Dr 1300 Loans receivable, Cr customer account (or cash). Processing fee to 4100 (with tax from Stage 2).
- Repayment: split into penalty, interest, principal in that order. Early payoff possible, with an optional fee.
- Interest income: accrue daily, post monthly. Dr 1310 Interest receivable, Cr 4000. When paid: Dr customer, Cr 1310 (interest part) and Cr 1300 (principal part).
- Late penalty after the grace days.
- Overdue classes by days past due: Standard (0 to 29), Watch (30 to 89), Substandard (90 to 179), Doubtful (180 to 359), Loss (360 and more) with provision percentages in `tblLoanClasses`. Monthly provision: Dr 5300, Cr 1320 Loan loss allowance (new GL). Write-off when Loss.
- **The bank only lends what it has:** before disbursement check that `1100` after the loan is at least `ReserveRatio x total deposits`, that total loans stay under a share of total deposits, and that one borrower's total stays under a share of capital. Otherwise refuse, or ask head office for approval.
- Approval levels by amount: branch manager, head office. Loan officer and approver must differ.

**Reports:** loan book by type and branch, overdue ageing, provisions, loan-to-deposit ratio, income from loans.

**Checks:** schedule totals equal the loan amount and interest; 1300 equals the sum of loan balances; allowance equals the sum of provisions.

---

## Stage 4: collateral, taking property, selling or renting it

**Tables:** `tblCollateral`, `tblAppraisals`, `tblRepossessions`, `tblAssetSales`, `tblAssetRentals`.

**Rules**
- Collateral types: land, house, vehicle, gold, machinery, stock, equipment. Each has a maximum loan-to-value percentage and a yearly value change.
- Valuation: appraised value, appraiser, date. Loan limit = appraised value x the type's loan-to-value, after a haircut. Re-appraise at set intervals. If value falls below a safe level, flag the loan.
- Lien: an item can back only one loan at a time. The item cannot be sold or moved while the lien is active. Released when the loan is repaid.
- Default process: a notice period with dates, then repossession approved by head office.
- Repossession: the bank records the item as its own property, `1400 Repossessed assets held for sale`, at the LOWER of fair value less selling costs and the loan balance. The rest of the loan is written off or kept as a claim. Interest stops being accrued.
- Re-pricing: new appraisal at repossession and every quarter. Value changes go to profit and loss (loss or gain).
- Options for each held asset: **sell** (auction or private sale), **rent out** (rental income 4500, maintenance cost, tenant), or **use in business** (put into a business unit and record its income and costs).
- Selling: proceeds first cover selling costs, then the unpaid loan, then interest and penalties. Surplus goes back to the customer. A shortfall stays a debt or is written off. Difference between sale price and book value = gain 4400 or loss 5310.
- Time limit: assets held too long are flagged for sale.

**Reports:** repossessed asset register, ageing, sale results, rental income.

**Checks:** 1400 equals the book value of held assets; no item with two liens; every sale has a proceeds journal that balances.

---

## Stage 5: staff, salaries and payroll

**Tables:** `tblPositions`, `tblSalaryScales`, `tblEmployees`, `tblAllowances`, `tblAttendance`, `tblTaxBrackets`, `tblPayrollRuns`, `tblPayslips`, `tblStaffAdvances`.

**Positions:** branch manager, assistant manager, teller, cashier, loan officer, customer service, security guard, cleaner, driver, IT support, head office departments (accounting, risk, HR, audit, treasury, operations).

**Rules**
- Salary scale per position and grade with an effective date.
- Allowances by branch type: city allowance, town hardship allowance, position allowance, overtime.
- Staffing guide per branch type (minimum positions) with a report of gaps.
- Monthly payroll run for the whole bank or for one branch: gross = base + allowances + overtime; income tax from progressive brackets; employee social security (2% sample); advances recovered; net pay. Employer social security (3% sample).
- Posting at the employee's branch: Dr 5100 salaries, Dr 5110 employer social security; Cr 2400 staff payable (net), Cr 2230 tax and social security payable. Paying the staff: Dr 2400, Cr 1900 (branch) and Dr 1900, Cr 1100 (head office).
- Transfers of staff between branches keep their history. Leave and unpaid days cut pay.
- Payslips as a printable sheet. Payroll cost per branch and per position.
- Users from Stage 1 link to employees.

**Checks:** payslips total equals the journal; each employee paid once per month.

---

## Stage 6: foreign currency, international payments, digital money

**Foreign currency**
- Foreign currency accounts. Every ledger line stores currency, rate and foreign amount.
- Counter exchange: the bank buys at the buy rate and sells at the sell rate (spread = profit, `4200`). Rates come from `tblExchangeRates` by date.
- The bank holds a cash position per currency per branch (`1700`). Daily revaluation of holdings at the mid rate: gain or loss to `4200`. Position limits and a replenishment order from treasury.
- Large or suspicious exchanges need approval and a purpose.

**International payments (fake)**
- Sending: customer pays in local or foreign currency, bank converts, takes a fee, books to a simulated foreign settlement balance. A message record has a status (sent, received, returned, failed). Incoming works in reverse. Screening against a blocked-names list (Stage 8).

**Digital money (e-wallets)**
- Wallet accounts (`2030`) with limits by KYC level: maximum balance, daily spending, monthly volume.
- Top-up from a bank account or cash at an agent; pay merchants by QR code (a merchant is a wallet); cash out at an agent; send to another wallet; fees and rewards.
- Agents: small cash points with float and commission.
- **Backing rule:** total wallet balances must be held in a trust asset of the same amount. A check proves it every day.
- Optional later: a simple central digital currency compared with ordinary deposits.

**Checks:** each currency position equals the ledger; trust account equals wallet total; FX revaluation journal balances.

---

## Stage 7: investing in businesses

**Tables:** `tblInvestees`, `tblInvestments`, `tblInvestmentReturns`, `tblInvestmentValuations`.

**Rules**
- Types: equity share, profit sharing, fixed return note, short term trade finance.
- Decision record: amount, term, expected return, risk grade, sector. Limits per business, per sector and in total as a percent of capital. Approval by head office only.
- Funding comes from treasury `1100`. Dr 1600 Business investments, Cr 1100.
- Returns: fixed return accrues daily; profit share is entered at each reporting date. Income to 4300.
- Valuation at each period end. Falls below cost go to expense (impairment). Disposal at gain or loss.
- Liquidity guard: no investment if it would break the reserve and liquidity rules (Stage 8 ratios).

**Checks:** 1600 equals the carrying value of investments.

---

## Stage 8: government rules, regulator, year end

**Rules and reports**
- **Policy rate and rate rules:** deposit and loan rates can be set as policy rate plus or minus a margin, with limits.
- **Reserve requirement:** `ReserveRatio` of deposits must stay in cash and treasury.
- **Capital adequacy:** capital divided by risk-weighted assets (loans, investments, bonds with different weights) must be above a minimum.
- **Liquidity:** liquid assets divided by expected outflows for 30 days (based on notice and maturities).
- **Bad loan ratio** and **single borrower limit** reports.
- **Government bonds:** the bank buys treasury bills and bonds with `1600`-type accounts, earns interest, counts as low risk.
- **Deposit insurance:** yearly premium on deposits, and a cover limit per customer shown in reports.
- **Taxes:** withholding tax and commercial tax paid monthly, income tax from annual profit, payroll tax from Stage 5; a tax return summary.
- **Anti money laundering (simulated):** reports for large cash deposits, a flag for repeated deposits just under the limit, a blocked names list, and a case list with status and notes.
- **Inspection:** a report that runs all checks, lists breaches, and can add a fine posted to `5500`.
- **Year end:** profit and loss closes to retained earnings (`3100`), accruals reviewed, a balance sheet and income statement for the year, branch ranking.
- **Management dashboard:** deposits, loans, loan-to-deposit, net interest margin, cost to income, bad loan ratio, liquidity, capital ratio, profit per branch.

**Checks:** after year end all income and expense accounts are zero and retained earnings equals the profit moved.

---

## Cross-cutting rules for every stage

1. Only macros post money. Every action is a unit of work: build in memory, check, write once.
2. Every change to a posted record is a reversal, never an edit or delete.
3. Maker and checker for risky actions. Every approval, status change and reversal is audited.
4. The control checks (journal balance, branch balance, 1900 = 0, account tie-out, 2900 and other control accounts) run at every end of day and stop the day when they fail.
5. No automated tests. Each stage has a manual checklist with numbers from the starter workbook.
6. `.bas` and `.cls` only, ASCII, CRLF, `Option Explicit`, unique prefixed names, Control sheet with buttons.

## Honest limits

- Excel runs one person at a time. Branches are data and roles are a switch on the Control sheet.
- A few thousand customers is fine. More than that needs a real database. The design carries over to VB.NET with SQL Server.
- It cannot be a real bank: no cards, ATMs, real payment networks, real law or real security.
- The plan is large. Expect weeks of stage by stage work with the agent.

## Suggested way of working

1. Give the agent `CLAUDE.md`, `DATA_TABLES.md` and `BankStarter.xlsx`. Say: "Read CLAUDE.md and DATA_TABLES.md. Start Phase 1 only."
2. After each phase, check what you can by hand before the next one.
3. When Stage 1 passes its checklist, ask for the Stage 2 `CLAUDE.md` section (it needs a short rewrite of the fee posting for tax on fees).
4. Commit to git after each finished stage.
