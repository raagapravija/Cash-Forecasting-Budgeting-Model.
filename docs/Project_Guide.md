# Project Guide — Cash Flow Forecasting & Budgeting Model

## How to use this guide

This is a beginner-friendly execution guide for the workbook `financial-model/cash_flow_forecasting_budgeting_model.xlsx`. It is designed as a portfolio project, but it follows the workflow used in a Finance/FP&A team:

1. Understand the business and modelling objective.
2. Load and clean source transactions.
3. Confirm categories and cash-flow direction.
4. Enter controlled assumptions.
5. Build Budget, Actual, Forecast and Cash Flow layers.
6. Reconcile the model.
7. Investigate variances and working capital.
8. Test scenarios and alerts.
9. Publish a management dashboard and recommendations.

All currency values in the example are **INR millions**. In a live model, the raw source may be INR units; choose one unit and apply it consistently.

---

## 1. Financial concepts in plain language

### Cash flow forecasting

Cash flow forecasting projects the timing and amount of cash receipts and payments. It is a liquidity view, not a profit view. A forecast typically combines scheduled invoices, collection behavior, payroll, vendor commitments, debt service, CapEx and funding actions.

### Budgeting

A budget is an approved plan. It translates business objectives into revenue targets, headcount, marketing spend, operations, technology, administration, financing and CapEx limits. The budget is used as the baseline for accountability and resource allocation.

### Budget vs Forecast vs Actual

- **Budget:** what management approved before the period.
- **Forecast:** the latest estimate after incorporating actual performance and new information.
- **Actual:** what was recorded in the accounting system, bank statement or operational source.

The difference between Budget and Actual explains performance. The difference between Forecast and Actual explains forecast quality. A strong FP&A analyst investigates both.

### Cash flow vs profit

Profit may recognize a SaaS invoice before cash is collected. Profit may also include depreciation or other non-cash expenses. Cash flow captures the actual receipt and payment date. A company can show accounting profit while cash is falling because customers are slow to pay, inventory/CapEx is paid upfront or debt must be repaid.

### Why this is an analyst project

The project demonstrates data preparation, Excel formulas, financial statement classification, budgeting, driver-based forecasting, working capital analysis, scenario planning, dashboard design, business insight and executive communication—all common Financial Analyst and FP&A responsibilities.

---

## 2. Company scenario

**ClarityCloud SaaS Pvt. Ltd.** sells workflow and analytics software to Indian logistics and retail companies.

| Item | Assumption |
|---|---|
| Industry | B2B SaaS / enterprise software |
| Business model | Subscription contracts plus implementation and support |
| Revenue sources | Subscription revenue, implementation fees, customer advances and other income |
| Cost structure | Payroll, cloud/software, marketing, vendors, facilities, administration, travel and CapEx |
| Customer terms | 60-day B2B terms; some annual prepayments |
| Supplier terms | Normally 30 days; strategic vendors managed within agreed terms |
| FY2025 opening cash | INR 50.0m illustrative starting point |
| FY2025 closing/current cash | INR 16.2m at 31-Dec-2025 |
| FY2026 minimum reserve | INR 12.0m |
| Forecast objective | Grow while protecting liquidity and avoiding emergency borrowing |

The FY2025 starting cash is included to make the historical cash bridge internally consistent. FY2026 opens at the FY2025 closing cash of INR 16.2m.

---

## 3. Dataset and data dictionary

The dataset contains 228 rows: 19 standardized cash-flow lines across 12 historical months.

### Required fields

| Field | Purpose |
|---|---|
| Transaction ID | Unique key for duplicate detection |
| Date | Transaction date |
| Month | Aggregation key, for example Jan-2025 |
| Quarter | Q1 to Q4 |
| Department | Owner / cost center group |
| Category | Operating Activities, Investing Activities or Financing Activities |
| Subcategory | Detailed cash-flow line |
| Transaction Type | Inflow or Outflow |
| Cash Inflow | Positive amount received |
| Cash Outflow | Positive amount paid |
| Budget Amount | Planned amount for the line |
| Actual Amount | Cash Inflow + Cash Outflow |
| Forecast Amount | Illustrative run-rate reference |
| Revenue | Recognized revenue reference; not necessarily collected cash |
| Operating Expense | Operating cash cost |
| Capital Expenditure | Investing cash cost |
| Accounts Receivable | Month-end AR snapshot on collection rows |
| Accounts Payable | Month-end AP snapshot on vendor rows |
| Opening Cash Balance | Month opening cash |
| Closing Cash Balance | Month closing cash |
| Variance Amount | Actual - Budget |
| Variance % | Variance / Budget |
| Payment Status | Collected, Received, Paid or Not Applicable |
| Cost Center | Controlled owner code |
| Remarks | Business explanation |

### Important design choice: positive cash magnitudes

Cash Inflow and Cash Outflow are stored as positive magnitudes. The Transaction Type and Category identify the direction. This avoids double negatives in formulas:

```excel
Net Cash Flow = SUM(Cash Inflow) - SUM(Cash Outflow)
```

If your organization stores outflows as negative values, use one convention consistently and document it in the data dictionary.

---

## 4. Data cleaning: step by step

Create a copy of the source before cleaning. In Excel, import the CSV with **Data → From Text/CSV** and select **Transform Data** if Power Query is available.

### 4.1 Remove duplicate transactions

1. Sort by Transaction ID, Date, Amount and Source Document.
2. Identify exact duplicate IDs.
3. Remove only when the source confirms it is a duplicate; two payments on the same date can be legitimate.
4. Reconcile the cleaned total to the bank/GL.

**Before:** `TXN-EX01` appears twice with the same date, department and amount.

**After:** retain one record and document the removed ID in an audit log.

Excel check:

```excel
=COUNTIF($A:$A,A2)>1
```

Power Query: **Remove Rows → Remove Duplicates** after choosing the business key.

### 4.2 Handle missing values

- Missing Department: map from Cost Center, vendor or source document; otherwise send to `Review`.
- Missing Category: do not guess silently; require a reviewer.
- Missing amount: check invoice/bank source. A blank amount is not automatically zero.
- Missing Payment Status: derive from transaction type only if the cash has cleared.

**Before:** `TXN-EX04` has no Department and no Subcategory.

**After:** either map it from the invoice or retain `Review - Missing Classification` in a data-quality queue.

### 4.3 Standardize dates

Convert `15/01/25`, `2025/01/15` and `31-03-2025` to real date values, not text. In Excel:

```excel
=DATEVALUE(A2)
```

If the source is ambiguous, confirm the locale before transforming. Then derive:

```excel
Month = TEXT([@Date],"mmm-yyyy")
Quarter = "Q"&ROUNDUP(MONTH([@Date])/3,0)
```

The workbook uses `Jan-2025` style Month keys for readable `SUMIFS` formulas.

### 4.4 Correct currency formats

Convert text such as `₹15,200,000`, `INR 3,500,000`, `(2.20)` and `1,200,000` to numeric values. A robust Power Query approach is to remove currency symbols and commas, then divide by 1,000,000 when converting INR units to INR millions.

**Before:** `₹15,200,000` stored as text.

**After:** `15.20` stored as a number with cell format `#,##0.00`.

### 4.5 Standardize department names

Create a mapping table:

| Source value | Standard value |
|---|---|
| Human Resources | HR |
| Marketing  | Marketing |
| Mktg | Marketing |
| Ops | Operations |
| Technology | IT |
| G&A | Administration |

Use `XLOOKUP` or Power Query Merge rather than repeated manual edits.

### 4.6 Validate signs and direction

Use these checks:

```excel
=IF(AND([@[Transaction Type]]="Inflow",[@[Cash Outflow]]<>0),"Review","OK")
=IF(AND([@[Transaction Type]]="Outflow",[@[Cash Inflow]]<>0),"Review","OK")
=IF([@[Actual Amount]]=([@[Cash Inflow]]+[@[Cash Outflow]]),"OK","Review")
```

**Before:** Vendor Payments stored as `(2.20)` in the outflow column.

**After:** Cash Outflow = `2.20`, Transaction Type = `Outflow`.

### 4.7 Detect unusual transactions

Start with a business threshold, not a blind statistical rule. Examples:

- Software or travel > 3× category median.
- Any individual transaction > INR 10m.
- Any zero-budget transaction with a material actual.
- Any cash balance movement inconsistent with bank reconciliation.

**Before:** IT software transaction is `999.999999` INR m.

**After:** flag it for invoice/source review. Do not delete an outlier just because it is large.

### 4.8 Validate budget and actuals

- Reconcile Budget to approved plan.
- Reconcile Actual to GL/bank extract.
- Check that totals by month and department tie to source reports.
- Check that budget is not blank for an approved line.
- Confirm that a zero budget is handled safely in Variance %.

```excel
=IFERROR(([@Actual]-[@Budget]) / [@Budget],0)
```

### 4.9 Check missing categories and data-entry errors

Use Data Validation lists for Category, Subcategory, Department, Transaction Type and Payment Status. Conditional-format blank cells in required fields. Check:

- Date outside the fiscal year.
- Quarter not matching Date.
- Department not matching Cost Center.
- Inflow/outflow direction not matching type.
- Budget or Actual below zero when the policy does not allow reversals.
- Closing cash not equal to Opening Cash + Net Cash Flow.

The repository includes `data/data_quality_examples.csv` so a student can demonstrate these checks before using the clean file.

---

## 5. Cash flow structure

### Operating Activities

| Line | Cash impact |
|---|---|
| Customer Collections | Main operating inflow; converts AR into cash |
| Customer Advances | Positive cash before the service is fully delivered; creates a delivery obligation |
| Salaries | Predictable payroll outflow; protect it for business continuity |
| Rent | Contractual facilities outflow |
| Marketing | Demand-generation outflow; assess ROI and timing |
| Utilities | Usage and cloud/data-transfer outflow |
| Software | Recurring licenses and platform subscriptions |
| Administration | Insurance, legal, audit and office support |
| Vendor Payments | Implementation/support supplier cash payments |
| Travel | Customer and sales activity outflow |
| Other Income | Minor operating receipts or recoveries |

### Investing Activities

| Line | Cash impact |
|---|---|
| Equipment Purchases | Cash paid for equipment; usually discretionary and deferrable |
| Technology Investment | Platform, security and architecture investment; may protect future growth |
| Property / Assets | Large asset purchase; create an approval gate |
| Other Capital Expenditure | Non-recurring investment not included in routine operating expense |

### Financing Activities

| Line | Cash impact |
|---|---|
| Loans / Loan Proceeds | Positive cash today, but creates principal and interest obligations |
| Loan Repayments | Reduces cash and debt outstanding |
| Equity Funding | Positive cash without scheduled repayment; may dilute ownership |
| Interest Payments | Financing outflow; does not reduce principal |

A management cash forecast should show Operating, Investing and Financing totals separately even if the dashboard also shows one total Outflow KPI.

---

## 6. Step-by-step Excel build

### Step 1 — Raw data

Load `data/cash_flow_transactions.csv` into `Raw_Data` and convert it to an Excel Table named `tblRawData`. Tables expand automatically when new rows are added.

Recommended controls:

- Freeze the top row.
- Turn on filters.
- Format dates as `dd-mmm-yyyy`.
- Format amounts as `#,##0.00;[Red](#,##0.00);-`.
- Add conditional formatting for Review flags and negative closing cash.

### Step 2 — Assumptions

Enter assumptions in blue cells. The workbook includes:

- Current / starting cash.
- Minimum cash reserve.
- Base, Best and Worst revenue growth.
- Base, Best and Worst collection rate.
- Operating expense factors.
- Pipeline conversion uplift.
- Monthly customer growth.
- Advance and other-income growth.
- Loan repayment and interest.
- Seasonality and staged CapEx.

Use named ranges in a more advanced version, for example `Min_Cash_Reserve`, `Base_Growth` and `Base_Collection_Rate`.

### Step 3 — Actuals

The `Actuals` sheet aggregates Raw_Data with `SUMIFS`:

```excel
=SUMIFS(Raw_Data!$N:$N,Raw_Data!$C:$C,$A5)
=SUMIFS(Raw_Data!$J:$J,Raw_Data!$C:$C,$A5,Raw_Data!$E:$E,"Marketing")
=SUMIFS(Raw_Data!$I:$I,Raw_Data!$C:$C,$A5)
=SUMIFS(Raw_Data!$J:$J,Raw_Data!$C:$C,$A5)
=J5-K5
```

Use the Actuals sheet to reconcile historical totals before forecasting.

### Step 4 — Budget

Create an annual budget by department and a monthly budget profile. The workbook includes Sales, Marketing, HR, Operations, IT, Administration, Finance and CapEx.

Core formulas:

```excel
Monthly Budget = Annual Budget / 12
Variance = Actual - Budget
Variance % = IFERROR(Variance / Budget,0)
Remaining Budget = Budget - Actual
```

For expenses, positive variance is usually unfavorable. For revenue, positive variance is usually favorable. Add a separate Interpretation column so the dashboard does not mislabel variances.

### Step 5 — Revenue Forecast

For each FY2026 month, link historical revenue to assumptions:

```excel
Base Revenue = Historical Revenue
             * (1 + Base Growth)
             * Seasonality
             * Customer Growth Factor
             + Historical Revenue * Pipeline Conversion Uplift

Cash Collections = Forecast Revenue * Collection Rate
```

The workbook has columns for Historical Revenue, Seasonality, Customer Growth Factor, Pipeline Conversion, Base/Best/Worst Revenue and Base/Best/Worst Cash Collections.

### Step 6 — Expense Forecast

Forecast fixed, variable and semi-variable expenses separately. The workbook applies scenario cost factors to monthly FY2025 run-rate expenses and links CapEx to the staged schedule.

```excel
Base Expense = Historical Same-Month Expense * Base Operating Expense Factor
```

This is a beginner-friendly driver. A more advanced model would use headcount, salary bands, campaign volume, customer count, vendor contracts and usage drivers.

### Step 7 — Cash Flow Forecast

For each scenario and month:

```excel
Total Cash Inflows = Customer Collections + Customer Advances
                   + Other Income + Financing Inflows

Total Cash Outflows = Operating Outflows + Investing Outflows
                    + Financing Outflows

Net Cash Flow = Total Cash Inflows - Total Cash Outflows
Closing Cash = Opening Cash + Net Cash Flow
```

Alert formula:

```excel
=IF(ClosingCash<MinimumReserve,"Critical",
 IF(ClosingCash<MinimumReserve*1.25,"Watch","Healthy"))
```

Opening cash for January is the current cash position. Opening cash for each later month equals the prior month closing cash.

### Step 8 — AR Aging

Create aging buckets from invoice due date or days past due. The project uses an illustrative snapshot:

| Bucket | AR |
|---|---:|
| 0–30 | 18.5 |
| 31–60 | 10.2 |
| 61–90 | 4.7 |
| 90+ | 2.6 |

Formulas:

```excel
Total AR = SUM(B5:B8)
Overdue AR = SUM(B6:B8)
Collection Rate = Total Cash Collections / Revenue
DSO = Total AR / Annual Revenue * 365
```

A delayed collection increases AR and reduces cash. The analyst should identify customer, owner, invoice, reason for delay and expected payment date—not simply report the total.

### Step 9 — AP Analysis

Use invoice due date and payment status to prepare current and overdue AP. Formulas:

```excel
Total AP = SUM(B5:B8)
Overdue AP = SUM(B6:B8)
DPO = Total AP / Selected Operating Cost Base * 365
```

The selected denominator must be documented. DPO should not be inflated by dividing AP by an unrelated expense line. Manage suppliers within agreed terms and communicate early before requesting a change.

### Step 10 — Working Capital

```excel
Working Capital = AR + Inventory / Prepaid Infrastructure - AP
DIO = Inventory / Relevant Cost Base * 365
Cash Conversion Cycle = DSO + DIO - DPO
```

This case uses INR 2.4m as a small inventory / prepaid-infrastructure proxy because the company is SaaS rather than manufacturing. A manufacturing model should use raw material, WIP and finished-goods inventory.

### Step 11 — Scenario Manager option

To build this yourself in Excel:

1. Put growth, collection rate, operating expense factor and CapEx controls in Assumptions.
2. Select **Data → What-If Analysis → Scenario Manager**.
3. Add `Base Case`, `Best Case` and `Worst Case`.
4. Use the assumption cells as changing cells.
5. Set the result cell to FY2026 Closing Cash.
6. Create a Summary report.
7. Compare the summary with the formula-driven `Scenario_Analysis` sheet.

The workbook already contains formula-driven scenarios so a reviewer can trace the logic without opening Scenario Manager.

### Step 12 — Dashboard

Create PivotTables or formulas from the Cash_Flow, Budget, Revenue_Forecast and AR/AP sheets. Add slicers for Month, Quarter, Department, Cost Category, Scenario and Transaction Type. The included Dashboard is formula/chart driven; Power BI can recreate it with measures.

---

## 7. Dashboard visualization purpose

| Visual | Purpose |
|---|---|
| KPI cards | Give management a one-minute view of revenue, cash, net movement and variance |
| Monthly Inflow vs Outflow | Shows cash generation and timing pressure |
| Closing Cash vs Reserve | Identifies Watch/Critical months before the gap occurs |
| Revenue Forecast by Scenario | Quantifies upside/downside revenue risk |
| Budget vs Actual | Highlights departments with material over/under-run |
| AR Aging | Shows trapped cash and collection priorities |
| AP / supplier view | Shows upcoming payment pressure and relationship risk |
| Scenario closing cash | Compares resilience under Best/Base/Worst assumptions |
| Financial alerts | Converts analysis into an action queue |
| Recommendations panel | Links insights to concrete management actions |

### Recommended filters

- **Month:** focus on one month or trend.
- **Quarter:** review Q1–Q4 performance.
- **Department:** assign variance ownership.
- **Cost Category:** isolate payroll, marketing, technology or CapEx.
- **Scenario:** switch Base, Best and Worst views.
- **Transaction Type:** separate receipts from payments.

---

## 8. Budget variance heatmap

Create a matrix with Department in rows and Month in columns. Use expense variance % or absolute variance, but define the sign convention clearly.

For expense lines:

- Green: actual at or below budget, or within ±5% with no issue.
- Yellow: moderate variance such as 5%–10% requiring commentary.
- Red: significant unfavorable expense overrun, for example >10% or above a materiality threshold.

For revenue lines, reverse the business interpretation: actual above budget is green; a revenue shortfall is red. A single color rule for all line items is dangerous because revenue and expense variance directions have different meanings.

Management use:

1. Scan for red cells.
2. Drill from department to category to transaction.
3. Ask whether the variance is volume, price, timing, one-off, FX, headcount or coding.
4. Assign an owner and corrective action.
5. Update the forecast if the variance is structural.

---

## 9. Monthly management report template

### Executive summary

Revenue was INR 232.8m in FY2025 versus a budget of INR 225.0m. However, expenses and CapEx exceeded plan and closing cash reduced to INR 16.2m. AR of INR 36.0m, including INR 17.5m overdue, is the main near-term cash conversion opportunity.

### Revenue performance

Revenue beat budget by INR 7.8m. The forecast assumes strong FY2026 growth supported by customer expansion, pipeline conversion and seasonality. Finance should reconcile recognized revenue to invoice and collection timing.

### Expense performance

Marketing, HR, Operations, IT and CapEx were above budget. Review whether the drivers are permanent or timing-related. Finance was below budget, but financing underspend is not a replacement for operating cash generation.

### Cash position and forecast

FY2026 Base Case opens at INR 16.2m. January–March are Watch months; April improves above the reserve as collections and forecast revenue build. The Worst Case is Critical from January and requires pre-emptive funding and cost actions.

### Working capital

AR and DSO are high relative to the cash position. AP should be managed within terms. The working capital bridge must be monitored monthly because growth can consume cash even when revenue is ahead of budget.

### Major variances

- Marketing: +INR 3.5m unfavorable expense variance.
- HR: +INR 3.9m unfavorable expense variance.
- Operations: +INR 2.3m unfavorable expense variance.
- CapEx: +INR 1.5m unfavorable variance.
- Sales: +INR 7.8m favorable revenue variance.

### Risks

1. Collection slippage or customer concentration.
2. Payroll and marketing cost creep.
3. Non-essential CapEx executed during a cash Watch period.
4. Supplier arrears affecting cloud or implementation continuity.
5. Worst-case revenue shortfall requiring external funding.

### Management recommendations

- Run a weekly AR collection sprint for 31+ day invoices.
- Maintain at least INR 12.0m cash reserve.
- Pre-approve or defer discretionary CapEx.
- Review marketing and HR variance with owners.
- Maintain a committed contingency facility.
- Refresh the model monthly and maintain a 13-week cash view.

---

## 10. Power BI / Tableau / SQL / Python extension

### Power BI

Load `cash_flow_transactions.csv`, create a Date table and relationships, then build measures:

```DAX
Cash Inflow = SUM('Raw_Data'[Cash Inflow])
Cash Outflow = SUM('Raw_Data'[Cash Outflow])
Net Cash Flow = [Cash Inflow] - [Cash Outflow]
Budget Variance = SUM('Raw_Data'[Actual Amount]) - SUM('Raw_Data'[Budget Amount])
```

Use slicers, conditional formatting, drill-through to transactions and a scenario table.

### Tableau

Connect the CSV, create calculated fields for Net Cash Flow and Variance %, then build dashboards using Month, Department, Category and Scenario filters. Use reference lines for Minimum Cash and budget.

### SQL

A SQL implementation would create a cleaned fact table and aggregate by month and department:

```sql
SELECT
  month,
  department,
  SUM(cash_inflow) AS cash_inflow,
  SUM(cash_outflow) AS cash_outflow,
  SUM(cash_inflow) - SUM(cash_outflow) AS net_cash_flow
FROM cash_flow_transactions
GROUP BY month, department;
```

### Python

Python can automate CSV quality checks, outlier detection, statistical forecast methods, Monte Carlo scenarios and scheduled report exports. It is optional; the core project remains non-coding Excel/Power BI.

---

## 11. Quality-control checklist

Before publishing:

- [ ] No duplicate Transaction IDs without documented explanation.
- [ ] All dates are valid dates within FY2025.
- [ ] Month and Quarter agree with Date.
- [ ] Departments and Cost Centers map to the approved list.
- [ ] Cash Inflow / Cash Outflow signs follow the documented convention.
- [ ] Actual Amount equals Inflow + Outflow.
- [ ] Budget and Actual totals reconcile to source.
- [ ] Opening Cash + Net Cash Flow = Closing Cash each month.
- [ ] Forecast formulas are linked to assumptions, not hard-coded outputs.
- [ ] Minimum-cash alerts behave correctly in a test scenario.
- [ ] AR/AP totals and aging buckets reconcile.
- [ ] Dashboard charts have titles, units and readable labels.
- [ ] Recommendations are linked to quantified insights.

---
