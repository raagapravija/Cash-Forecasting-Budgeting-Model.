from datetime import date
from pathlib import Path
import csv, math, shutil, textwrap
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from matplotlib.patches import FancyBboxPatch
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.formatting.rule import CellIsRule, FormulaRule, ColorScaleRule
from openpyxl.chart import LineChart, BarChart, DoughnutChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.chart.marker import DataPoint
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.comments import Comment

ROOT = Path(__file__).resolve().parents[1]
for p in [ROOT / 'data', ROOT / 'financial-model', ROOT / 'forecast', ROOT / 'budget-analysis', ROOT / 'dashboard', ROOT / 'screenshots', ROOT / 'docs', ROOT / 'scripts']:
    p.mkdir(parents=True, exist_ok=True)

# -----------------------------
# 1) Source data and calculations
# -----------------------------
company = 'ClarityCloud SaaS Pvt. Ltd.'
currency = 'INR (INR millions unless stated)'
months = [date(2025, m, 1) for m in range(1, 13)]
month_keys = [d.strftime('%b-%Y') for d in months]
month_short = [d.strftime('%b') for d in months]
quarters = [f'Q{((d.month-1)//3)+1}' for d in months]

revenue_billed = [14.5, 15.2, 16.0, 17.0, 17.8, 18.6, 19.5, 20.5, 21.2, 22.5, 24.0, 26.0]
collections = [12.5, 13.0, 13.8, 14.5, 15.2, 16.0, 16.8, 17.5, 18.2, 19.0, 20.5, 22.0]
advances = [0.6, 0.5, 0.7, 0.8, 0.8, 0.9, 0.9, 1.0, 1.0, 1.1, 1.2, 1.3]
other_income = [0.1] * 12

expense_series = {
    'Salaries': [8.0, 8.0, 8.2, 8.2, 8.5, 8.5, 8.8, 8.8, 9.0, 9.2, 9.2, 9.5],
    'Rent': [1.5] * 12,
    'Marketing': [3.5, 3.2, 3.5, 3.8, 4.0, 4.2, 4.0, 4.3, 4.5, 5.0, 5.5, 6.0],
    'Utilities': [0.60, 0.58, 0.62, 0.65, 0.68, 0.70, 0.72, 0.74, 0.76, 0.80, 0.85, 0.90],
    'Software': [1.40, 1.40, 1.45, 1.45, 1.50, 1.50, 1.55, 1.55, 1.60, 1.60, 1.65, 1.70],
    'Administration': [1.00, 0.95, 1.05, 1.10, 1.10, 1.15, 1.10, 1.20, 1.25, 1.30, 1.35, 1.50],
    'Vendor Payments': [2.0, 2.1, 2.2, 2.3, 2.4, 2.5, 2.6, 2.7, 2.8, 3.0, 3.2, 3.4],
    'Travel': [0.50, 0.45, 0.55, 0.60, 0.70, 0.75, 0.70, 0.80, 0.85, 1.10, 1.20, 1.30],
}
capex_equipment = [0, 0, 0, 3.0, 0, 0, 0, 2.0, 0, 0, 4.0, 0]
capex_technology = [0, 0, 1.5, 0, 0, 0, 1.0, 0, 0, 0, 2.0, 0]
loan_proceeds = [0, 0, 0, 10.0, 0, 0, 0, 0, 0, 0, 0, 0]
equity_funding = [0, 0, 0, 0, 0, 0, 23.0, 0, 0, 0, 0, 0]
loan_repayment = [0, 0, 0, 0, 0.7, 0.7, 0.7, 0.7, 0.7, 0.7, 0.7, 0.7]
interest_payments = [0.3] * 12

line_defs = [
    ('Sales', 'Operating Activities', 'Customer Collections', 'Inflow', collections, 1.05, 1.10, 'Collected', 'CC-SALES', '60-day customer collection cycle'),
    ('Sales', 'Operating Activities', 'Customer Advances', 'Inflow', advances, 0.95, 1.05, 'Collected', 'CC-SALES', 'Annual prepayment and implementation deposits'),
    ('Finance', 'Operating Activities', 'Other Income', 'Inflow', other_income, 1.00, 1.03, 'Received', 'CC-FIN', 'Interest income and small recoveries'),
    ('HR', 'Operating Activities', 'Salaries', 'Outflow', expense_series['Salaries'], 0.96, 1.06, 'Paid', 'CC-HR', 'Monthly payroll including benefits'),
    ('Operations', 'Operating Activities', 'Rent', 'Outflow', expense_series['Rent'], 1.02, 1.04, 'Paid', 'CC-OPS', 'Office and co-working facilities'),
    ('Marketing', 'Operating Activities', 'Marketing', 'Outflow', expense_series['Marketing'], 0.93, 1.08, 'Paid', 'CC-MKT', 'Demand generation and events'),
    ('Operations', 'Operating Activities', 'Utilities', 'Outflow', expense_series['Utilities'], 0.95, 1.05, 'Paid', 'CC-OPS', 'Cloud data transfer and office utilities'),
    ('IT', 'Operating Activities', 'Software', 'Outflow', expense_series['Software'], 1.02, 1.05, 'Paid', 'CC-IT', 'SaaS tools and cloud subscriptions'),
    ('Administration', 'Operating Activities', 'Administration', 'Outflow', expense_series['Administration'], 1.01, 1.04, 'Paid', 'CC-ADMIN', 'Insurance, legal, audit and office admin'),
    ('Operations', 'Operating Activities', 'Vendor Payments', 'Outflow', expense_series['Vendor Payments'], 0.95, 1.10, 'Paid', 'CC-OPS', 'Implementation partners and support vendors'),
    ('Sales', 'Operating Activities', 'Travel', 'Outflow', expense_series['Travel'], 0.90, 1.12, 'Paid', 'CC-SALES', 'Customer visits and sales travel'),
    ('Capital Expenditure', 'Investing Activities', 'Equipment Purchases', 'Outflow', capex_equipment, 0.90, 1.10, 'Paid', 'CC-CAPEX', 'Laptop, networking and office equipment'),
    ('Capital Expenditure', 'Investing Activities', 'Technology Investment', 'Outflow', capex_technology, 0.90, 1.12, 'Paid', 'CC-CAPEX', 'Platform architecture and security investment'),
    ('Capital Expenditure', 'Investing Activities', 'Property/Assets', 'Outflow', [0.0] * 12, 1.00, 1.00, 'Not Applicable', 'CC-CAPEX', 'No property purchase planned in FY2025'),
    ('Capital Expenditure', 'Investing Activities', 'Other Capital Expenditure', 'Outflow', [0.0] * 12, 1.00, 1.00, 'Not Applicable', 'CC-CAPEX', 'Placeholder for future approved asset projects'),
    ('Finance', 'Financing Activities', 'Loan Proceeds', 'Inflow', loan_proceeds, 1.00, 1.00, 'Received', 'CC-FIN', 'Working capital term loan drawdown'),
    ('Finance', 'Financing Activities', 'Loan Repayment', 'Outflow', loan_repayment, 1.00, 1.00, 'Paid', 'CC-FIN', 'Scheduled principal repayment'),
    ('Finance', 'Financing Activities', 'Equity Funding', 'Inflow', equity_funding, 1.00, 1.00, 'Received', 'CC-FIN', 'Series A capital contribution'),
    ('Finance', 'Financing Activities', 'Interest Payments', 'Outflow', interest_payments, 0.95, 1.05, 'Paid', 'CC-FIN', 'Monthly interest on working capital loan'),
]

# Historical AR/AP snapshots used for analytics.
ar_aging = {'0-30 days': 18.5, '31-60 days': 10.2, '61-90 days': 4.7, '90+ days': 2.6}
ap_aging = {'0-30 days': 13.8, '31-60 days': 5.9, '61-90 days': 2.4, '90+ days': 1.1}
closing_ar = sum(ar_aging.values())
closing_ap = sum(ap_aging.values())

# Calculate monthly cash balances for FY2025.
opening_cash = 50.0
monthly_balances = []
for i in range(12):
    inflow = collections[i] + advances[i] + other_income[i] + loan_proceeds[i] + equity_funding[i]
    outflow = sum(v[i] for _, _, sub, typ, v, *_ in line_defs if typ == 'Outflow')
    closing_cash = opening_cash + inflow - outflow
    monthly_balances.append({'opening': opening_cash, 'inflow': inflow, 'outflow': outflow, 'net': inflow-outflow, 'closing': closing_cash})
    opening_cash = closing_cash
current_cash = monthly_balances[-1]['closing']  # 16.2

# -----------------------------
# 2) CSV dataset
# -----------------------------
headers = ['Transaction ID', 'Date', 'Month', 'Quarter', 'Department', 'Category', 'Subcategory', 'Transaction Type',
           'Cash Inflow', 'Cash Outflow', 'Budget Amount', 'Actual Amount', 'Forecast Amount', 'Revenue',
           'Operating Expense', 'Capital Expenditure', 'Accounts Receivable', 'Accounts Payable', 'Opening Cash Balance',
           'Closing Cash Balance', 'Variance Amount', 'Variance %', 'Payment Status', 'Cost Center', 'Remarks']
rows = []
transaction_no = 1
for i, d in enumerate(months):
    for idx, (dept, cat, sub, typ, values, budget_factor, forecast_factor, status, cc, remarks) in enumerate(line_defs):
        actual = float(values[i])
        inflow = actual if typ == 'Inflow' else 0.0
        outflow = actual if typ == 'Outflow' else 0.0
        budget = actual * budget_factor
        forecast = actual * forecast_factor
        revenue = revenue_billed[i] if sub == 'Customer Collections' else 0.0
        opex = actual if (typ == 'Outflow' and cat == 'Operating Activities') else 0.0
        capex = actual if cat == 'Investing Activities' else 0.0
        ar = closing_ar if sub == 'Customer Collections' else 0.0
        ap = closing_ap if sub == 'Vendor Payments' else 0.0
        bal = monthly_balances[i]
        rows.append([
            f'TXN-{transaction_no:04d}', date(d.year, d.month, min(28, 8 + idx * 2)), month_keys[i], quarters[i], dept, cat, sub, typ,
            round(inflow, 2), round(outflow, 2), round(budget, 2), round(actual, 2), round(forecast, 2), round(revenue, 2),
            round(opex, 2), round(capex, 2), round(ar, 2), round(ap, 2), round(bal['opening'], 2), round(bal['closing'], 2),
            round(actual - budget, 2), round((actual - budget) / budget if budget else 0.0, 4), status if actual else 'Not Applicable', cc, remarks
        ])
        transaction_no += 1

csv_path = ROOT / 'data' / 'cash_flow_transactions.csv'
with csv_path.open('w', newline='', encoding='utf-8') as f:
    writer = csv.writer(f)
    writer.writerow(headers)
    writer.writerows(rows)

# Data-quality examples for cleaning practice.
quality_examples = [
    ['TXN-EX01', '15/01/25', 'Jan 2025', 'Marketing ', 'Marketing', 'Marketing', 'Outflow', 'INR 3,500,000', '3.5', 'Duplicate candidate: same ID/date/amount appears twice'],
    ['TXN-EX02', '2025/02/20', 'Feb-2025', 'Human Resources', 'Operating Activities', 'Salaries', 'Outflow', '8000000', '8.0', 'Department name must map to HR'],
    ['TXN-EX03', '31-03-2025', 'Mar-2025', 'Operations', 'Operating Activities', 'Vendor Payments', 'Outflow', '(2.20)', '-2.2', 'Outflow sign should be stored as positive magnitude in Cash Outflow'],
    ['TXN-EX04', '2025-04-15', 'Apr-2025', '', 'Operating Activities', '', 'Outflow', '1,200,000', '1.2', 'Missing department/category requires review'],
    ['TXN-EX05', '2025-05-10', 'May-2025', 'Sales', 'Operating Activities', 'Customer Collections', 'Inflow', '₹15,200,000', '15.2', 'Currency text must be converted to numeric'],
    ['TXN-EX06', '2025-06-12', 'Jun-2025', 'IT', 'Operating Activities', 'Software', 'Outflow', '999999999', '999.999999', 'Unusual transaction: test against threshold and source invoice'],
]
with (ROOT / 'data' / 'data_quality_examples.csv').open('w', newline='', encoding='utf-8') as f:
    writer = csv.writer(f)
    writer.writerow(['Example ID', 'Date Before', 'Month Before', 'Department Before', 'Category', 'Subcategory', 'Type', 'Amount Before', 'Clean Amount (INR m)', 'Cleaning Note'])
    writer.writerows(quality_examples)

# Forecast calculations used for documentation and static image.
seasonality = [0.95, 0.93, 0.98, 0.97, 1.00, 1.02, 1.04, 1.03, 1.05, 1.12, 1.23, 1.31]
customer_growth = [1 + 0.01 * (i + 1) for i in range(12)]
base_growth, best_growth, worst_growth = 0.28, 0.40, -0.10
base_pipeline, best_pipeline, worst_pipeline = 0.05, 0.08, 0.01
base_collection, best_collection, worst_collection = 0.86, 0.92, 0.72
base_opex, best_opex, worst_opex = 0.92, 0.85, 1.12
base_capex = [1.0, 0, 1.2, 0, 0, 0, 1.0, 0, 0, 0, 2.0, 0]
best_capex = [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 1.0, 0]
worst_capex = [3.0, 1.0, 2.0, 3.0, 1.0, 0, 2.0, 3.0, 1.0, 0, 4.0, 0]

rev_forecast = []
for i in range(12):
    base_rev = revenue_billed[i] * (1 + base_growth) * seasonality[i] * customer_growth[i] + revenue_billed[i] * base_pipeline
    best_rev = revenue_billed[i] * (1 + best_growth) * seasonality[i] * (1 + 0.015*(i+1)) + revenue_billed[i] * best_pipeline
    worst_rev = revenue_billed[i] * (1 + worst_growth) * seasonality[i] * (1 - 0.01*(i+1)) + revenue_billed[i] * worst_pipeline
    rev_forecast.append({'base_revenue': base_rev, 'best_revenue': best_rev, 'worst_revenue': worst_rev,
                         'base_collections': base_rev*base_collection, 'best_collections': best_rev*best_collection, 'worst_collections': worst_rev*worst_collection})

# Base operating expense by month excludes capex and financing.
base_operating = []
for i in range(12):
    total = sum(v[i] for _, _, sub, typ, v, *_ in line_defs if typ == 'Outflow' and _ is not None and sub not in ['Equipment Purchases', 'Technology Investment', 'Loan Repayment', 'Interest Payments'])
    # The comprehension above intentionally filters by subcategory; use explicit list for clarity.
    total = sum(expense_series[k][i] for k in expense_series)
    base_operating.append(total)

scenario_rows = {}
for name, cf_key, cf_rate, opex_factor, capex_schedule in [
    ('Base Case', 'base_collections', base_collection, base_opex, base_capex),
    ('Best Case', 'best_collections', best_collection, best_opex, best_capex),
    ('Worst Case', 'worst_collections', worst_collection, worst_opex, worst_capex),
]:
    cash = current_cash
    result = []
    for i in range(12):
        collections_fc = rev_forecast[i][cf_key]
        advances_fc = advances[i] * 1.05
        other_fc = other_income[i] * 1.03
        inflow = collections_fc + advances_fc + other_fc
        operating = base_operating[i] * opex_factor
        financing_out = loan_repayment[i] + interest_payments[i]
        outflow = operating + capex_schedule[i] + financing_out
        net = inflow - outflow
        cash += net
        alert = 'Critical' if cash < 12.0 else ('Watch' if cash < 15.0 else 'Healthy')
        result.append({'opening': cash-net, 'collections': collections_fc, 'advances': advances_fc, 'other': other_fc, 'inflow': inflow,
                       'operating': operating, 'capex': capex_schedule[i], 'financing_out': financing_out, 'outflow': outflow, 'net': net,
                       'closing': cash, 'alert': alert})
    scenario_rows[name] = result

# -----------------------------
# 3) Excel workbook generation
# -----------------------------
wb = Workbook()
# remove default
ws = wb.active
ws.title = 'Raw_Data'

# Palette
NAVY = '17324D'
BLUE = '1F5A85'
TEAL = '0E9F9A'
LIGHT_TEAL = 'DFF4F2'
LIGHT_BLUE = 'EAF2F8'
GOLD = 'F4B942'
RED = 'D64545'
LIGHT_RED = 'FCE8E6'
GREEN = '2E8B57'
LIGHT_GREEN = 'E8F5E9'
YELLOW = 'FFF4CC'
GREY = '667085'
LIGHT_GREY = 'F4F6F8'
WHITE = 'FFFFFF'
BLACK = '17202A'
thin_grey = Side(style='thin', color='D9E2EC')
medium_navy = Side(style='medium', color=NAVY)

currency_fmt = '#,##0.00;[Red](#,##0.00);-'
pct_fmt = '0.0%;[Red](0.0%);-'
int_fmt = '#,##0;[Red](#,##0);-'


def title(ws, text, subtitle=None, end_col=12):
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=end_col)
    c = ws.cell(1, 1, text)
    c.font = Font(name='Aptos Display', size=18, bold=True, color=WHITE)
    c.fill = PatternFill('solid', fgColor=NAVY)
    c.alignment = Alignment(horizontal='left', vertical='center')
    ws.row_dimensions[1].height = 30
    if subtitle:
        ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=end_col)
        c2 = ws.cell(2, 1, subtitle)
        c2.font = Font(name='Aptos', size=10, italic=True, color=GREY)
        c2.alignment = Alignment(wrap_text=True)
        ws.row_dimensions[2].height = 28


def header_row(ws, row, start_col, end_col, fill=NAVY):
    for col in range(start_col, end_col+1):
        cell = ws.cell(row, col)
        cell.font = Font(name='Aptos', bold=True, color=WHITE)
        cell.fill = PatternFill('solid', fgColor=fill)
        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        cell.border = Border(bottom=medium_navy)
    ws.row_dimensions[row].height = 30


def style_body(ws, min_row, max_row, min_col, max_col):
    for row in ws.iter_rows(min_row=min_row, max_row=max_row, min_col=min_col, max_col=max_col):
        for c in row:
            c.font = Font(name='Aptos', size=10, color=BLACK)
            c.alignment = Alignment(vertical='center')
            c.border = Border(bottom=thin_grey)


def set_widths(ws, widths):
    for col, width in widths.items():
        ws.column_dimensions[col].width = width

# Raw data
for c, h in enumerate(headers, 1):
    ws.cell(1, c, h)
header_row(ws, 1, 1, len(headers))
for r_idx, row in enumerate(rows, 2):
    for c_idx, value in enumerate(row, 1):
        cell = ws.cell(r_idx, c_idx, value)
        if c_idx == 2:
            cell.number_format = 'dd-mmm-yyyy'
        elif c_idx in [9,10,11,12,13,14,15,16,17,18,19,20,21]:
            cell.number_format = currency_fmt
        elif c_idx == 22:
            cell.number_format = pct_fmt
style_body(ws, 2, len(rows)+1, 1, len(headers))
ws.freeze_panes = 'A2'
ws.auto_filter.ref = f'A1:Y{len(rows)+1}'
raw_tab = Table(displayName='tblRawData', ref=f'A1:Y{len(rows)+1}')
raw_tab.tableStyleInfo = TableStyleInfo(name='TableStyleMedium2', showRowStripes=True, showColumnStripes=False)
ws.add_table(raw_tab)
set_widths(ws, {'A':14,'B':14,'C':12,'D':10,'E':18,'F':23,'G':24,'H':16,'I':14,'J':14,'K':14,'L':14,'M':15,'N':12,'O':16,'P':18,'Q':18,'R':16,'S':18,'T':18,'U':15,'V':12,'W':18,'X':13,'Y':42})
ws.sheet_view.showGridLines = False
ws['A1'].comment = Comment('Cleaned, analysis-ready transaction table. Positive values are stored as magnitudes in Cash Inflow/Cash Outflow; Transaction Type determines direction.', 'Arena')

# Assumptions
ws = wb.create_sheet('Assumptions')
title(ws, 'Assumptions & Model Drivers', 'Editable blue cells drive the FY2026 revenue, expense, liquidity and scenario calculations. Amounts are INR millions.', 14)
assumptions = [
    ('Company Name', company), ('Currency / Unit', currency), ('Historical Financial Year', 'FY2025'), ('Forecast Year', 'FY2026'),
    ('Current Cash Position at 31-Dec-2025', current_cash), ('Minimum Required Cash Reserve', 12.0), ('Starting Cash for FY2026', current_cash),
    ('Base Revenue Growth', base_growth), ('Best Revenue Growth', best_growth), ('Worst Revenue Growth', worst_growth),
    ('Base Cash Collection Rate', base_collection), ('Best Cash Collection Rate', best_collection), ('Worst Cash Collection Rate', worst_collection),
    ('Base Operating Expense Factor', base_opex), ('Best Operating Expense Factor', best_opex), ('Worst Operating Expense Factor', worst_opex),
    ('Base Pipeline Conversion Uplift', base_pipeline), ('Best Pipeline Conversion Uplift', best_pipeline), ('Worst Pipeline Conversion Uplift', worst_pipeline),
    ('Base Monthly Customer Growth', 0.01), ('Best Monthly Customer Growth', 0.015), ('Worst Monthly Customer Growth', -0.01),
    ('Advance Collections Growth', 0.05), ('Other Income Growth', 0.03), ('Monthly Loan Repayment', 0.70), ('Monthly Interest Payment', 0.30),
    ('Tax Rate (reference only)', 0.25), ('As-of Date', date(2025, 12, 31)),
]
ws['A4'] = 'Driver'; ws['B4'] = 'Value'; ws['C4'] = 'Interpretation'
header_row(ws, 4, 1, 3)
for i, (label, val) in enumerate(assumptions, 5):
    ws.cell(i, 1, label); ws.cell(i, 2, val)
    interpretations = {
        'Current Cash Position at 31-Dec-2025': 'Opening liquidity for the forecast year',
        'Minimum Required Cash Reserve': 'Alert threshold for liquidity monitoring',
        'Base Revenue Growth': 'Expected annual growth before seasonality and pipeline uplift',
        'Base Cash Collection Rate': 'Revenue converted to cash during the month',
        'Base Operating Expense Factor': 'Applied to FY2025 run-rate to reflect productivity and inflation',
        'Base Pipeline Conversion Uplift': 'Incremental revenue from qualified pipeline',
        'As-of Date': 'Cut-off date for the illustrative historical dataset',
    }
    ws.cell(i, 3, interpretations.get(label, 'Scenario or model control')).alignment = Alignment(wrap_text=True)
    ws.cell(i, 2).fill = PatternFill('solid', fgColor='DCEBFA')
    if isinstance(val, (int, float)) and label not in ['Monthly Loan Repayment', 'Monthly Interest Payment']:
        if 'Rate' in label or 'Growth' in label or 'Factor' in label or 'Uplift' in label or 'Tax' in label:
            ws.cell(i, 2).number_format = pct_fmt
        else:
            ws.cell(i, 2).number_format = currency_fmt
    elif label in ['Monthly Loan Repayment', 'Monthly Interest Payment']:
        ws.cell(i, 2).number_format = currency_fmt
    elif isinstance(val, date):
        ws.cell(i, 2).number_format = 'dd-mmm-yyyy'
style_body(ws, 5, 4+len(assumptions), 1, 3)
# Historical opening cash is kept separate from the forecast starting cash so the FY2025 actual bridge is auditable.
ws.cell(34, 1, 'Historical Opening Cash FY2025'); ws.cell(34, 2, 50.0); ws.cell(34, 3, 'Opening balance used only for FY2025 historical reconciliation')
ws.cell(34, 2).number_format = currency_fmt; ws.cell(34, 2).fill = PatternFill('solid', fgColor='DCEBFA')
style_body(ws, 34, 34, 1, 3)

# Seasonality
ws['E4'] = 'Month'; ws['F4'] = 'Seasonality Factor'; ws['G4'] = 'Advance Collections'; ws['H4'] = 'Other Income'; ws['I4'] = 'Base CapEx'; ws['J4'] = 'Best CapEx'; ws['K4'] = 'Worst CapEx'
header_row(ws, 4, 5, 11)
for i, m in enumerate(month_short, 5):
    j = i - 5
    ws.cell(i, 5, f'{m}-2026')
    ws.cell(i, 6, seasonality[j])
    ws.cell(i, 7, advances[j] * 1.05)
    ws.cell(i, 8, other_income[j] * 1.03)
    ws.cell(i, 9, base_capex[j]); ws.cell(i, 10, best_capex[j]); ws.cell(i, 11, worst_capex[j])
    ws.cell(i, 6).number_format = '0.00x'
    for c in range(7, 12): ws.cell(i, c).number_format = currency_fmt
style_body(ws, 5, 16, 5, 11)
ws['M4'] = 'Model notes'; ws['M4'].font = Font(bold=True, color=WHITE); ws['M4'].fill = PatternFill('solid', fgColor=NAVY)
notes = ['Base case assumes strong SaaS growth with improving collections.', 'Worst case models a 10% revenue contraction, 72% cash conversion and 12% cost inflation.', 'CapEx is intentionally staged so management can defer non-essential purchases when cash is tight.', 'All amounts are illustrative and should be replaced by ERP/bank data in a live model.']
for r, note in enumerate(notes, 5):
    ws.cell(r, 13, '• ' + note); ws.cell(r, 13).alignment = Alignment(wrap_text=True, vertical='top')
    ws.merge_cells(start_row=r, start_column=13, end_row=r, end_column=14)
ws.freeze_panes = 'A5'; ws.sheet_view.showGridLines = False
set_widths(ws, {'A':39,'B':18,'C':48,'E':14,'F':18,'G':18,'H':15,'I':14,'J':14,'K':14,'M':55,'N':18})

# Actuals
ws = wb.create_sheet('Actuals')
title(ws, 'FY2025 Actual Cash Flow by Department', 'Formula-driven monthly actuals from Raw_Data. Use this sheet to validate the source data before building the forecast.', 13)
actual_headers = ['Month', 'Sales Revenue', 'Marketing', 'HR', 'Operations', 'IT', 'Administration', 'Finance', 'CapEx', 'Total Cash Inflows', 'Total Cash Outflows', 'Net Cash Flow', 'Closing Cash Balance']
for c,h in enumerate(actual_headers,1): ws.cell(4,c,h)
header_row(ws, 4, 1, len(actual_headers))
for r, mk in enumerate(month_keys, 5):
    ws.cell(r,1,mk)
    ws.cell(r,2, f'=SUMIFS(Raw_Data!$N:$N,Raw_Data!$C:$C,$A{r})')
    for c, dept in zip(range(3,9), ['Marketing','HR','Operations','IT','Administration','Finance']):
        ws.cell(r,c, f'=SUMIFS(Raw_Data!$J:$J,Raw_Data!$C:$C,$A{r},Raw_Data!$E:$E,{chr(34)}{dept}{chr(34)})')
    ws.cell(r,9, f'=SUMIFS(Raw_Data!$P:$P,Raw_Data!$C:$C,$A{r})')
    ws.cell(r,10, f'=SUMIFS(Raw_Data!$I:$I,Raw_Data!$C:$C,$A{r})')
    ws.cell(r,11, f'=SUMIFS(Raw_Data!$J:$J,Raw_Data!$C:$C,$A{r})')
    ws.cell(r,12, f'=J{r}-K{r}')
    ws.cell(r,13, f'=IF(ROW()=5,Assumptions!$B$34+L{r},M{r-1}+L{r})')
for r in range(5,17):
    for c in range(2,14): ws.cell(r,c).number_format = currency_fmt
style_body(ws, 5, 16, 1, 13)
ws.cell(17,1,'FY2025 Total / Ending')
for c in range(2,13):
    letter = get_column_letter(c)
    ws.cell(17,c, f'=SUM({letter}5:{letter}16)')
    ws.cell(17,c).number_format = currency_fmt
ws.cell(17,13,'=M16'); ws.cell(17,13).number_format = currency_fmt
for c in range(1,14):
    ws.cell(17,c).font = Font(bold=True, color=NAVY); ws.cell(17,c).fill = PatternFill('solid', fgColor=LIGHT_BLUE)
ws.freeze_panes='A5'; ws.sheet_view.showGridLines=False
set_widths(ws, {get_column_letter(i): 17 for i in range(1,14)})

# Revenue Forecast
ws = wb.create_sheet('Revenue_Forecast')
title(ws, 'FY2026 Revenue Forecast & Cash Conversion', 'Base, best and worst cases combine historical revenue, growth, seasonality, customer growth, pipeline conversion and collection rates.', 14)
rf_headers = ['Month', 'Historical Revenue', 'Seasonality', 'Customer Growth Factor', 'Pipeline Conversion', 'Base Revenue', 'Best Revenue', 'Worst Revenue', 'Base Collection Rate', 'Best Collection Rate', 'Worst Collection Rate', 'Base Cash Collections', 'Best Cash Collections', 'Worst Cash Collections']
for c,h in enumerate(rf_headers,1): ws.cell(4,c,h)
header_row(ws, 4, 1, 14)
for r, mk in enumerate([f'{m}-2026' for m in month_short], 5):
    i = r-5
    ws.cell(r,1,mk); ws.cell(r,2,revenue_billed[i]); ws.cell(r,2).number_format=currency_fmt
    ws.cell(r,3, f'=INDEX(Assumptions!$F$5:$F$16,MATCH($A{r},Assumptions!$E$5:$E$16,0))')
    ws.cell(r,4, f'=1+Assumptions!$B$24*(ROW()-4)')
    ws.cell(r,5, '=Assumptions!$B$21')
    ws.cell(r,6, f'=B{r}*(1+Assumptions!$B$12)*C{r}*D{r}+B{r}*E{r}')
    ws.cell(r,7, f'=B{r}*(1+Assumptions!$B$13)*C{r}*(1+Assumptions!$B$25*(ROW()-4))+B{r}*Assumptions!$B$22')
    ws.cell(r,8, f'=B{r}*(1+Assumptions!$B$14)*C{r}*(1+Assumptions!$B$26*(ROW()-4))+B{r}*Assumptions!$B$23')
    ws.cell(r,9, '=Assumptions!$B$15'); ws.cell(r,10, '=Assumptions!$B$16'); ws.cell(r,11, '=Assumptions!$B$17')
    ws.cell(r,12, f'=F{r}*I{r}'); ws.cell(r,13, f'=G{r}*J{r}'); ws.cell(r,14, f'=H{r}*K{r}')
    for c in [3,4,5,9,10,11]: ws.cell(r,c).number_format='0.0%'
    for c in [6,7,8,12,13,14]: ws.cell(r,c).number_format=currency_fmt
style_body(ws,5,16,1,14)
ws.cell(17,1,'FY2026 Total')
for c in [2,6,7,8,12,13,14]:
    letter=get_column_letter(c); ws.cell(17,c,f'=SUM({letter}5:{letter}16)'); ws.cell(17,c).number_format=currency_fmt
for c in range(1,15): ws.cell(17,c).font=Font(bold=True,color=NAVY); ws.cell(17,c).fill=PatternFill('solid',fgColor=LIGHT_BLUE)
# Revenue chart
chart = LineChart(); chart.title='Revenue Forecast by Scenario'; chart.y_axis.title='INR m'; chart.x_axis.title='Month'; chart.height=7; chart.width=14
chart.add_data(Reference(ws,min_col=6,max_col=8,min_row=4,max_row=16), titles_from_data=True)
chart.set_categories(Reference(ws,min_col=1,min_row=5,max_row=16)); chart.style=13
ws.add_chart(chart,'A20')
ws.freeze_panes='A5'; ws.sheet_view.showGridLines=False
set_widths(ws, {'A':13,'B':18,'C':13,'D':20,'E':18,'F':16,'G':16,'H':16,'I':18,'J':18,'K':18,'L':20,'M':20,'N':20})

# Expense Forecast
ws = wb.create_sheet('Expense_Forecast')
title(ws, 'FY2026 Expense Forecast by Cost Behavior', 'Fixed costs use run-rate assumptions; variable costs respond to demand; semi-variable costs combine a base run-rate with activity sensitivity.', 16)
expense_categories = [
    ('Salaries','Fixed','Salaries','Headcount and compensation plan'), ('Marketing','Variable','Marketing','Pipeline and campaign intensity'), ('Rent','Fixed','Rent','Lease contract'),
    ('Utilities','Semi-variable','Utilities','Usage and cloud consumption'), ('Technology','Semi-variable','Software','Licenses and platform subscriptions'), ('Administration','Semi-variable','Administration','Business support overhead'),
    ('Travel','Variable','Travel','Customer meetings and sales activity'), ('Vendor Costs','Variable','Vendor Payments','Implementation and support volume'), ('CapEx','Variable','CapEx','Approved investment schedule'),
]
scenario_specs = [('Base Case', 5, '$B$18', 'base'), ('Best Case', 20, '$B$19', 'best'), ('Worst Case', 35, '$B$20', 'worst')]
for scenario, start, factor_cell, sc_key in scenario_specs:
    ws.cell(start-1,1,scenario); ws.cell(start-1,1).font=Font(size=12,bold=True,color=TEAL)
    exp_headers=['Category','Cost Behavior','Raw Subcategory','Forecast Driver'] + [f'{m}-2026' for m in month_short] + ['Total']
    for c,h in enumerate(exp_headers,1): ws.cell(start,c,h)
    header_row(ws,start,1,len(exp_headers),fill=BLUE)
    for r, (category, behavior, raw_sub, driver) in enumerate(expense_categories, start+1):
        ws.cell(r,1,category); ws.cell(r,2,behavior); ws.cell(r,3,raw_sub); ws.cell(r,4,driver)
        for c in range(5,17):
            col_letter=get_column_letter(c)
            month_index = c-5
            if category == 'CapEx':
                cap_col = {'base':'I','best':'J','worst':'K'}[sc_key]
                ws.cell(r,c, f'=INDEX(Assumptions!${cap_col}$5:${cap_col}$16,MATCH({col_letter}${start},Assumptions!$E$5:$E$16,0))')
            else:
                ws.cell(r,c, f'=SUMIFS(Raw_Data!$J:$J,Raw_Data!$G:$G,$C{r},Raw_Data!$C:$C,LEFT({col_letter}${start},3)&"-2025")*Assumptions!{factor_cell}')
            ws.cell(r,c).number_format=currency_fmt
        ws.cell(r,17, f'=SUM(E{r}:P{r})'); ws.cell(r,17).number_format=currency_fmt
    # totals
    total_op = start+10
    total_capex = start+11
    ws.cell(total_op,1,'Total Operating Expenses'); ws.cell(total_capex,1,'Total CapEx')
    for c in range(5,17):
        L=get_column_letter(c)
        ws.cell(total_op,c,f'=SUM({L}{start+1}:{L}{start+8})')
        ws.cell(total_capex,c,f'={L}{start+9}')
        ws.cell(total_op,c).number_format=currency_fmt; ws.cell(total_capex,c).number_format=currency_fmt
    ws.cell(total_op,17,f'=SUM(E{total_op}:P{total_op})'); ws.cell(total_capex,17,f'=SUM(E{total_capex}:P{total_capex})')
    for r in [total_op,total_capex]:
        for c in range(1,18):
            ws.cell(r,c).font=Font(bold=True,color=NAVY); ws.cell(r,c).fill=PatternFill('solid',fgColor=LIGHT_BLUE)
            if c>=5: ws.cell(r,c).number_format=currency_fmt
style_body(ws,6,46,1,17)
ws.freeze_panes='E6'; ws.sheet_view.showGridLines=False
set_widths(ws, {'A':22,'B':16,'C':20,'D':36, **{get_column_letter(i):14 for i in range(5,18)}})

# Budget
ws = wb.create_sheet('Budget')
title(ws, 'FY2025 Budget vs Actual & FY2026 Forecast', 'Revenue is treated as favorable when actual exceeds budget; expenses are favorable when actual is below budget.', 12)
budget_headers=['Department / Line Item','Cost Center','Annual Budget FY2025','Monthly Budget','Actual FY2025','Forecast FY2026 Base','Remaining Budget / (Gap)','Variance Actual - Budget','Variance %','Interpretation']
for c,h in enumerate(budget_headers,1): ws.cell(4,c,h)
header_row(ws,4,1,len(budget_headers))
budget_data=[
    ('Sales','CC-SALES',225.0,'Revenue'), ('Marketing','CC-MKT',48.0,'Expense'), ('HR','CC-HR',100.0,'Expense'), ('Operations','CC-OPS',65.0,'Expense'),
    ('IT','CC-IT',18.0,'Expense'), ('Administration','CC-ADMIN',14.0,'Expense'), ('Finance','CC-FIN',10.0,'Expense'), ('Capital Expenditure','CC-CAPEX',12.0,'Expense')]
forecast_map={'Sales':'=SUM(Revenue_Forecast!$F$5:$F$16)','Marketing':'=SUM(Expense_Forecast!$E$22:$P$22)','HR':'=SUM(Expense_Forecast!$E$21:$P$21)','Operations':'=SUM(Expense_Forecast!$E$23:$P$23)+SUM(Expense_Forecast!$E$24:$P$24)+SUM(Expense_Forecast!$E$28:$P$28)+SUM(Expense_Forecast!$E$27:$P$27)','IT':'=SUM(Expense_Forecast!$E$25:$P$25)','Administration':'=SUM(Expense_Forecast!$E$26:$P$26)','Finance':'=SUM(Assumptions!$B$29:$B$30)*12','Capital Expenditure':'=SUM(Expense_Forecast!$E$29:$P$29)'}
# Explicit actual formulas by row
for r,(dept,cc,bud,kind) in enumerate(budget_data,5):
    ws.cell(r,1,dept); ws.cell(r,2,cc); ws.cell(r,3,bud); ws.cell(r,4,f'=C{r}/12')
    if dept=='Sales': actual_formula='=SUM(Raw_Data!$N:$N)'
    elif dept=='Capital Expenditure': actual_formula='=SUM(Raw_Data!$P:$P)'
    else: actual_formula=f'=SUMIFS(Raw_Data!$J:$J,Raw_Data!$E:$E,$A{r})'
    ws.cell(r,5,actual_formula); ws.cell(r,6,forecast_map[dept]); ws.cell(r,7,f'=C{r}-E{r}'); ws.cell(r,8,f'=E{r}-C{r}'); ws.cell(r,9,f'=IFERROR(H{r}/C{r},0)')
    if dept=='Sales': interp='Revenue: positive variance is favorable'
    else: interp='Expense: positive variance is unfavorable'
    ws.cell(r,10,interp)
    for c in range(3,9): ws.cell(r,c).number_format=currency_fmt
    ws.cell(r,9).number_format=pct_fmt
style_body(ws,5,12,1,10)
# conditional formatting variance percent
ws.conditional_formatting.add('I5:I12', ColorScaleRule(start_type='min', start_color=LIGHT_GREEN, mid_type='percentile', mid_value=50, mid_color=YELLOW, end_type='max', end_color=LIGHT_RED))
for r in range(5,13): ws.cell(r,10).alignment=Alignment(wrap_text=True)
# monthly budget table
ws['A15']='Monthly Budget Profile FY2025'; ws['A15'].font=Font(size=12,bold=True,color=TEAL)
monthly_budget_headers=['Department'] + month_keys + ['Annual Total']
for c,h in enumerate(monthly_budget_headers,1): ws.cell(16,c,h)
header_row(ws,16,1,len(monthly_budget_headers),fill=BLUE)
for r,(dept,cc,bud,kind) in enumerate(budget_data,17):
    ws.cell(r,1,dept)
    for c in range(2,14):
        factor = f'=Assumptions!$F${c+3}' if dept=='Sales' else '=1'
        ws.cell(r,c, f'=$C${r-12}/12*{factor}')
        ws.cell(r,c).number_format=currency_fmt
    ws.cell(r,14,f'=SUM(B{r}:M{r})'); ws.cell(r,14).number_format=currency_fmt
style_body(ws,17,24,1,14)
ws.freeze_panes='A5'; ws.sheet_view.showGridLines=False
set_widths(ws, {'A':25,'B':15,'C':19,'D':16,'E':16,'F':19,'G':21,'H':20,'I':13,'J':36, **{get_column_letter(i):14 for i in range(11,15)}})

# Cash Flow - three scenario tables
ws = wb.create_sheet('Cash_Flow')
title(ws, 'FY2026 Monthly Cash Flow Forecast', 'Net Cash Flow = Total Cash Inflows – Total Cash Outflows. Closing Cash = Opening Cash + Net Cash Flow. All three cases use the same liquidity logic.', 40)
cf_headers=['Month','Opening Cash','Customer Collections','Customer Advances','Other Income','Financing Inflows','Total Cash Inflows','Operating Outflows','Investing Outflows','Financing Outflows','Total Cash Outflows','Net Cash Flow','Closing Cash','Minimum Cash','Alert']
cf_specs=[('Base Case',1,1,'base'),('Best Case',17,17,'best'),('Worst Case',33,33,'worst')]
for scenario, title_col, start_col, key in cf_specs:
    ws.cell(3,start_col,scenario); ws.cell(3,start_col).font=Font(size=12,bold=True,color=TEAL)
    for j,h in enumerate(cf_headers,start_col): ws.cell(4,j,h)
    header_row(ws,4,start_col,start_col+14,fill=BLUE)
    for r,mk in enumerate([f'{m}-2026' for m in month_short],5):
        i=r-5
        c0=start_col
        ws.cell(r,c0,mk)
        # opening cash
        ws.cell(r,c0+1, f'=Assumptions!$B$11' if r==5 else f'={get_column_letter(c0+12)}{r-1}')
        rev_col={'base':'L','best':'M','worst':'N'}[key]
        ws.cell(r,c0+2, f'=Revenue_Forecast!${rev_col}{r}')
        ws.cell(r,c0+3, f'=INDEX(Assumptions!$G$5:$G$16,MATCH(${get_column_letter(c0)}{r},Assumptions!$E$5:$E$16,0))')
        ws.cell(r,c0+4, f'=INDEX(Assumptions!$H$5:$H$16,MATCH(${get_column_letter(c0)}{r},Assumptions!$E$5:$E$16,0))')
        ws.cell(r,c0+5, '=0')
        ws.cell(r,c0+6, f'=SUM({get_column_letter(c0+2)}{r}:{get_column_letter(c0+5)}{r})')
        # Expense forecast rows: base total op row 15, capex row 16; best 30/31; worst 45/46
        op_row={'base':15,'best':30,'worst':45}[key]; cap_row={'base':16,'best':31,'worst':46}[key]
        exp_col=get_column_letter(5+i)
        ws.cell(r,c0+7, f'=Expense_Forecast!${exp_col}${op_row}')
        ws.cell(r,c0+8, f'=Expense_Forecast!${exp_col}${cap_row}')
        fin_factor = {'base':'1','best':'0.9','worst':'1.1'}[key]
        ws.cell(r,c0+9, f'=(Assumptions!$B$29+Assumptions!$B$30)*{fin_factor}')
        ws.cell(r,c0+10, f'=SUM({get_column_letter(c0+7)}{r}:{get_column_letter(c0+9)}{r})')
        ws.cell(r,c0+11, f'={get_column_letter(c0+6)}{r}-{get_column_letter(c0+10)}{r}')
        ws.cell(r,c0+12, f'={get_column_letter(c0+1)}{r}+{get_column_letter(c0+11)}{r}')
        ws.cell(r,c0+13, '=Assumptions!$B$10')
        ws.cell(r,c0+14, f'=IF({get_column_letter(c0+12)}{r}<{get_column_letter(c0+13)}{r},"Critical",IF({get_column_letter(c0+12)}{r}<{get_column_letter(c0+13)}{r}*1.25,"Watch","Healthy"))')
        for c in range(c0+1,c0+14): ws.cell(r,c).number_format=currency_fmt
    # Total row
    tr=17
    ws.cell(tr,start_col,'FY2026 Total / Ending')
    for c in range(start_col+2,start_col+12):
        L=get_column_letter(c); ws.cell(tr,c,f'=SUM({L}5:{L}16)'); ws.cell(tr,c).number_format=currency_fmt
    ws.cell(tr,start_col+1,'=0'); ws.cell(tr,start_col+12,f'={get_column_letter(start_col+12)}16'); ws.cell(tr,start_col+12).number_format=currency_fmt
    for c in range(start_col,start_col+15):
        ws.cell(tr,c).font=Font(bold=True,color=NAVY); ws.cell(tr,c).fill=PatternFill('solid',fgColor=LIGHT_BLUE)
style_body(ws,5,16,1,47)
# Note the blank spaces between tables.
for c in range(1,48): ws.column_dimensions[get_column_letter(c)].width = 14
for c in [1,17,33]: ws.column_dimensions[get_column_letter(c)].width=13
for c in [15,31,47]: ws.column_dimensions[get_column_letter(c)].width=14
ws.freeze_panes='A5'; ws.sheet_view.showGridLines=False

# AR aging
ws = wb.create_sheet('AR_Aging')
title(ws, 'Accounts Receivable Aging', 'Aging highlights how much revenue is not yet cash. Overdue receivables are the 31+ day buckets.', 8)
for c,h in enumerate(['Aging Bucket','Receivable Balance','% of Total AR','Action / Interpretation'],1): ws.cell(4,c,h)
header_row(ws,4,1,4)
for r,(bucket,val) in enumerate(ar_aging.items(),5):
    ws.cell(r,1,bucket); ws.cell(r,2,val); ws.cell(r,3,f'=B{r}/$B$9')
    action={'0-30 days':'Current; maintain invoice discipline','31-60 days':'Send collection reminder and confirm disputes','61-90 days':'Escalate to account owner / credit control','90+ days':'High-risk; consider provision or payment plan'}[bucket]
    ws.cell(r,4,action); ws.cell(r,2).number_format=currency_fmt; ws.cell(r,3).number_format=pct_fmt
ws.cell(9,1,'Total Receivables'); ws.cell(9,2,'=SUM(B5:B8)'); ws.cell(9,2).number_format=currency_fmt; ws.cell(9,3,'=SUM(C5:C8)'); ws.cell(9,3).number_format=pct_fmt
ws.cell(10,1,'Overdue Receivables'); ws.cell(10,2,'=SUM(B6:B8)'); ws.cell(10,2).number_format=currency_fmt
ws.cell(11,1,'Collection Rate'); ws.cell(11,2,'=SUM(Raw_Data!$I:$I)/SUM(Raw_Data!$N:$N)'); ws.cell(11,2).number_format=pct_fmt
ws.cell(12,1,'DSO'); ws.cell(12,2,'=B9/SUM(Raw_Data!$N:$N)*365'); ws.cell(12,2).number_format='0.0'
for r in range(9,13):
    for c in range(1,5): ws.cell(r,c).font=Font(bold=True,color=NAVY); ws.cell(r,c).fill=PatternFill('solid',fgColor=LIGHT_BLUE)
style_body(ws,5,8,1,4); ws.freeze_panes='A5'; ws.sheet_view.showGridLines=False
set_widths(ws, {'A':24,'B':20,'C':16,'D':58})

# AP analysis
ws = wb.create_sheet('AP_Analysis')
title(ws, 'Accounts Payable Analysis', 'AP management balances supplier relationships with liquidity preservation. Do not delay undisputed invoices beyond agreed terms.', 8)
for c,h in enumerate(['Aging Bucket','Payable Balance','% of Total AP','Action / Interpretation'],1): ws.cell(4,c,h)
header_row(ws,4,1,4)
for r,(bucket,val) in enumerate(ap_aging.items(),5):
    ws.cell(r,1,bucket); ws.cell(r,2,val); ws.cell(r,3,f'=B{r}/$B$9')
    action={'0-30 days':'Pay within agreed terms','31-60 days':'Prioritize strategic suppliers','61-90 days':'Agree payment plan before escalation','90+ days':'Resolve overdue balance immediately'}[bucket]
    ws.cell(r,4,action); ws.cell(r,2).number_format=currency_fmt; ws.cell(r,3).number_format=pct_fmt
ws.cell(9,1,'Total Payables'); ws.cell(9,2,'=SUM(B5:B8)'); ws.cell(9,2).number_format=currency_fmt; ws.cell(9,3,'=SUM(C5:C8)'); ws.cell(9,3).number_format=pct_fmt
ws.cell(10,1,'Overdue Payables'); ws.cell(10,2,'=SUM(B6:B8)'); ws.cell(10,2).number_format=currency_fmt
ws.cell(11,1,'DPO'); ws.cell(11,2,'=B9/SUMIFS(Raw_Data!$O:$O,Raw_Data!$G:$G,"<>Salaries")*365'); ws.cell(11,2).number_format='0.0'
ws.cell(12,1,'Supplier Payment Discipline'); ws.cell(12,2,'Maintain approved payment terms'); ws.merge_cells('B12:D12')
ws['A15']='Upcoming Supplier Payment Plan'; ws['A15'].font=Font(size=12,bold=True,color=TEAL)
ap_headers=['Supplier / Vendor','Due Date','Amount','Payment Status','Cash Planning Action']
for c,h in enumerate(ap_headers,1): ws.cell(16,c,h)
header_row(ws,16,1,5,fill=BLUE)
ap_due=[('Nimbus Cloud Services',date(2026,1,15),4.8,'Scheduled','Protect cloud continuity'),('Orbit Implementation Partners',date(2026,1,25),3.2,'Scheduled','Confirm milestone acceptance'),('Urban Office Lease',date(2026,2,1),1.5,'Scheduled','Contractual payment'),('Growth Media Network',date(2026,2,10),2.1,'Review','Validate campaign invoices'),('SecureStack',date(2026,2,20),1.7,'Scheduled','Protect security services')]
for r,(vendor,due,amt,status,action) in enumerate(ap_due,17):
    ws.cell(r,1,vendor); ws.cell(r,2,due); ws.cell(r,2).number_format='dd-mmm-yyyy'; ws.cell(r,3,amt); ws.cell(r,3).number_format=currency_fmt; ws.cell(r,4,status); ws.cell(r,5,action)
style_body(ws,17,21,1,5)
for r in range(9,13):
    for c in range(1,5): ws.cell(r,c).font=Font(bold=True,color=NAVY); ws.cell(r,c).fill=PatternFill('solid',fgColor=LIGHT_BLUE)
style_body(ws,5,8,1,4); ws.freeze_panes='A5'; ws.sheet_view.showGridLines=False
set_widths(ws, {'A':30,'B':20,'C':16,'D':58,'E':35})

# Working capital and scenario analysis
ws = wb.create_sheet('Scenario_Analysis')
title(ws, 'Scenario Analysis & Working Capital', 'Scenarios test the effect of revenue, collection speed, operating cost control and CapEx timing on liquidity.', 12)
for c,h in enumerate(['Scenario','Revenue','Cash Inflow','Cash Outflow','Net Cash Flow','Closing Cash','Minimum Cash','Cash Runway (months)','First Critical Month','Management Interpretation'],1): ws.cell(4,c,h)
header_row(ws,4,1,10)
scenario_names=['Best Case','Base Case','Worst Case']
for r,sc in enumerate(scenario_names,5):
    ws.cell(r,1,sc)
    rev_col={'Best Case':'G','Base Case':'F','Worst Case':'H'}[sc]
    start_col={'Best Case':'Q','Base Case':'A','Worst Case':'AG'}[sc]
    ws.cell(r,2,f'=SUM(Revenue_Forecast!${rev_col}$5:${rev_col}$16)')
    in_col={'Best Case':'W','Base Case':'G','Worst Case':'AM'}[sc]
    ws.cell(r,3,f'=SUM(Cash_Flow!${in_col}$5:${in_col}$16)')
    # explicit scenario cash flow columns: Base G/K/L/M/O, Best W/AA/AB/AC/AE, Worst AM/AQ/AR/AS/AU.
    out_col={'Best Case':'AA','Base Case':'K','Worst Case':'AQ'}[sc]
    net_col={'Best Case':'AB','Base Case':'L','Worst Case':'AR'}[sc]
    close_col={'Best Case':'AC','Base Case':'M','Worst Case':'AS'}[sc]
    alert_col={'Best Case':'AE','Base Case':'O','Worst Case':'AU'}[sc]
    ws.cell(r,4,f'=SUM(Cash_Flow!${out_col}$5:${out_col}$16)')
    ws.cell(r,5,f'=SUM(Cash_Flow!${net_col}$5:${net_col}$16)')
    ws.cell(r,6,f'=Cash_Flow!${close_col}$16')
    ws.cell(r,7,f'=MIN(Cash_Flow!${close_col}$5:${close_col}$16)')
    ws.cell(r,8,f'=IFERROR(F{r}/ABS(AVERAGEIF(Cash_Flow!${net_col}$5:${net_col}$16,"<0",Cash_Flow!${net_col}$5:${net_col}$16)),"Self-funding")')
    ws.cell(r,9,f'=IFERROR(INDEX(Cash_Flow!${start_col}$5:${start_col}$16,MATCH("Critical",Cash_Flow!${alert_col}$5:${alert_col}$16,0)),"None")')
    interp={'Best Case':'Cash remains above reserve; fund growth selectively.','Base Case':'Liquidity is adequate but April is close to the reserve line.','Worst Case':'Immediate funding and cost actions are required.'}[sc]
    ws.cell(r,10,interp)
    for c in range(2,8): ws.cell(r,c).number_format=currency_fmt
    ws.cell(r,8).number_format='0.0'
style_body(ws,5,7,1,10)
for c in range(2,8): ws.cell(8,c,'')
ws['A11']='Working Capital Bridge'; ws['A11'].font=Font(size=12,bold=True,color=TEAL)
wc_headers=['Metric','Value','Formula / Interpretation']
for c,h in enumerate(wc_headers,1): ws.cell(12,c,h)
header_row(ws,12,1,3,fill=BLUE)
wc_rows=[
    ('Accounts Receivable','=AR_Aging!$B$9','Closing AR balance'), ('Inventory / Prepaid Infrastructure',2.4,'Small operating buffer for hardware / prepaid services'),
    ('Accounts Payable','=AP_Analysis!$B$9','Closing AP balance'), ('Working Capital','=B13+B14-B15','AR + inventory - AP'), ('DSO','=AR_Aging!$B$12','Receivable days'),
    ('DIO','=B14/SUMIFS(Raw_Data!$O:$O,Raw_Data!$G:$G,"Vendor Payments")*365','Inventory days proxy'), ('DPO','=AP_Analysis!$B$11','Payable days'), ('Cash Conversion Cycle','=B17+B18-B19','DSO + DIO - DPO'),
]
for r,(metric,val,note) in enumerate(wc_rows,13):
    ws.cell(r,1,metric); ws.cell(r,2,val); ws.cell(r,3,note)
    if r in [13,14,15,16]: ws.cell(r,2).number_format=currency_fmt
    else: ws.cell(r,2).number_format='0.0'
style_body(ws,13,20,1,3); ws.freeze_panes='A5'; ws.sheet_view.showGridLines=False
set_widths(ws, {'A':28,'B':19,'C':19,'D':19,'E':19,'F':19,'G':19,'H':22,'I':22,'J':52})

# KPI summary
ws = wb.create_sheet('KPI_Summary')
title(ws, 'KPI Summary & Management Metrics', 'Use the KPI layer to turn transaction-level data into concise FP&A reporting for senior management.', 6)
for c,h in enumerate(['KPI','Value','Unit','Definition','Management Use','Direction'],1): ws.cell(4,c,h)
header_row(ws,4,1,6)
kpis=[
    ('Revenue','=SUM(Raw_Data!$N:$N)','INR m','Recognized revenue in FY2025','Top-line performance','Higher is generally better'),
    ('Revenue Growth %','=SUM(Revenue_Forecast!$F$5:$F$16)/SUM(Raw_Data!$N:$N)-1','%','Base FY2026 revenue vs FY2025','Test growth plan','Higher is generally better'),
    ('Total Cash Inflow','=SUM(Raw_Data!$I:$I)','INR m','Cash received, including financing','Liquidity sources','Higher is not always better if debt-funded'),
    ('Total Cash Outflow','=SUM(Raw_Data!$J:$J)','INR m','Cash paid across operating, investing and financing','Liquidity uses','Lower is generally better'),
    ('Net Cash Flow','=SUM(Raw_Data!$I:$I)-SUM(Raw_Data!$J:$J)','INR m','Cash inflow less cash outflow','Funding requirement','Positive is better'),
    ('Closing Cash Balance','=Actuals!$M$16','INR m','FY2025 ending cash','Liquidity buffer','Higher is better above reserve'),
    ('Operating Expenses','=SUM(Raw_Data!$O:$O)','INR m','Operating cash costs excluding CapEx and financing','Cost base','Lower with service quality maintained'),
    ('Budget Utilization %','=SUM(Budget!$E$6:$E$12)/SUM(Budget!$C$6:$C$12)','%','Actual expense / budget expense','Budget control','Near or below 100% for expenses'),
    ('Budget Variance','=SUM(Budget!$H$6:$H$12)','INR m','Actual minus budget for expense lines','Root-cause analysis','Negative is favorable for expenses'),
    ('Gross Margin %','=(SUM(Raw_Data!$N:$N)-SUMIFS(Raw_Data!$O:$O,Raw_Data!$G:$G,"Vendor Payments"))/SUM(Raw_Data!$N:$N)','%','Revenue less direct vendor cost / revenue','Unit economics','Higher is better'),
    ('Operating Margin %','=(SUM(Raw_Data!$N:$N)-SUM(Raw_Data!$O:$O))/SUM(Raw_Data!$N:$N)','%','Revenue less operating expenses / revenue','Profitability proxy','Higher is better'),
    ('Accounts Receivable','=AR_Aging!$B$9','INR m','Closing receivable balance','Collection focus','Lower at same growth is better'),
    ('Accounts Payable','=AP_Analysis!$B$9','INR m','Closing payable balance','Payment planning','Manage within terms'),
    ('DSO','=AR_Aging!$B$12','Days','Receivables / revenue x 365','Collection efficiency','Lower is better'),
    ('DPO','=AP_Analysis!$B$11','Days','Payables / operating cost base x 365','Supplier funding','Within agreed terms'),
    ('Working Capital','=Scenario_Analysis!$B$16','INR m','AR + inventory - AP','Liquidity tied up in operations','Lower need is better'),
    ('Burn Rate','=ABS(AVERAGEIF(Cash_Flow!$L$5:$L$16,"<0",Cash_Flow!$L$5:$L$16))','INR m / month','Average negative base-case net cash flow','Runway planning','Lower is better'),
    ('Cash Runway','=IFERROR(Cash_Flow!$M$16/Burn_Rate,"Self-funding")','Months','Available cash / average monthly net burn','Funding urgency','Higher is better'),
]
for r,(k,v,u,d,use,dirn) in enumerate(kpis,5):
    ws.cell(r,1,k); ws.cell(r,2,v); ws.cell(r,3,u); ws.cell(r,4,d); ws.cell(r,5,use); ws.cell(r,6,dirn)
    if u=='%': ws.cell(r,2).number_format=pct_fmt
    elif u.startswith('INR'): ws.cell(r,2).number_format=currency_fmt
    else: ws.cell(r,2).number_format='0.0'
    ws.cell(r,4).alignment=Alignment(wrap_text=True); ws.cell(r,5).alignment=Alignment(wrap_text=True)
style_body(ws,5,22,1,6); ws.freeze_panes='A5'; ws.sheet_view.showGridLines=False
set_widths(ws, {'A':25,'B':18,'C':18,'D':42,'E':38,'F':27})

# Dashboard
ws = wb.create_sheet('Dashboard')
ws.sheet_view.showGridLines = False
title(ws, 'Cash Flow Forecasting & Budgeting Dashboard', 'ClarityCloud SaaS Pvt. Ltd. | FY2026 Base Forecast | INR millions | Refresh source data, then review alerts and scenarios.', 16)
# KPI cards
cards=[('Total Revenue','=KPI_Summary!$B$5','INR m'),('Cash Inflow','=KPI_Summary!$B$7','INR m'),('Cash Outflow','=KPI_Summary!$B$8','INR m'),('Net Cash Flow','=KPI_Summary!$B$9','INR m'),('Closing Cash','=Cash_Flow!$M$16','INR m'),('Budget Variance %','=KPI_Summary!$B$12','%')]
for idx,(label,formula,unit) in enumerate(cards):
    start=1+idx*2
    ws.merge_cells(start_row=4,start_column=start,end_row=4,end_column=start+1)
    ws.cell(4,start,label); ws.cell(4,start).font=Font(bold=True,color=WHITE); ws.cell(4,start).fill=PatternFill('solid',fgColor=BLUE); ws.cell(4,start).alignment=Alignment(horizontal='center')
    ws.merge_cells(start_row=5,start_column=start,end_row=6,end_column=start+1)
    ws.cell(5,start,formula); ws.cell(5,start).font=Font(size=18,bold=True,color=NAVY); ws.cell(5,start).fill=PatternFill('solid',fgColor=LIGHT_BLUE); ws.cell(5,start).alignment=Alignment(horizontal='center',vertical='center')
    ws.cell(5,start).number_format=pct_fmt if unit=='%' else currency_fmt
    ws.merge_cells(start_row=7,start_column=start,end_row=7,end_column=start+1)
    ws.cell(7,start,unit); ws.cell(7,start).font=Font(size=9,italic=True,color=GREY); ws.cell(7,start).alignment=Alignment(horizontal='center')
# alert panel
ws.merge_cells('A9:F9'); ws['A9']='Financial Alerts'; ws['A9'].font=Font(size=12,bold=True,color=WHITE); ws['A9'].fill=PatternFill('solid',fgColor=NAVY)
alerts=[('Base-case minimum cash','=MIN(Cash_Flow!$M$5:$M$16)','Reserve test'),('Base-case watch / critical months','=COUNTIF(Cash_Flow!$O$5:$O$16,"<>Healthy")','Action queue'),('AR overdue','=AR_Aging!$B$10','Collections focus'),('AP overdue','=AP_Analysis!$B$10','Supplier risk')]
for r,(label,formula,note) in enumerate(alerts,10):
    ws.cell(r,1,label); ws.cell(r,2,formula); ws.cell(r,3,note); ws.cell(r,2).number_format=currency_fmt if 'cash' in label.lower() or 'AR' in label or 'AP' in label else int_fmt
    ws.cell(r,1).font=Font(bold=True,color=NAVY)
    ws.cell(r,2).fill=PatternFill('solid',fgColor=LIGHT_TEAL)
# Add charts
# Cash inflow/outflow chart based on Cash_Flow base
chart1=LineChart(); chart1.title='Monthly Cash Inflow vs Outflow'; chart1.y_axis.title='INR m'; chart1.height=7; chart1.width=12
chart1.add_data(Reference(wb['Cash_Flow'],min_col=7,max_col=11,min_row=4,max_row=16), from_rows=False, titles_from_data=True)
# only use G and K; remove extra data series if any by constructing separately
chart1=LineChart(); chart1.title='Monthly Cash Inflow vs Outflow'; chart1.y_axis.title='INR m'; chart1.height=7; chart1.width=12
chart1.add_data(Reference(wb['Cash_Flow'],min_col=7,max_col=7,min_row=4,max_row=16), titles_from_data=True)
chart1.add_data(Reference(wb['Cash_Flow'],min_col=11,max_col=11,min_row=4,max_row=16), titles_from_data=True)
chart1.set_categories(Reference(wb['Cash_Flow'],min_col=1,min_row=5,max_row=16)); chart1.style=13
ws.add_chart(chart1,'A14')
chart2=LineChart(); chart2.title='Closing Cash Forecast vs Minimum Reserve'; chart2.y_axis.title='INR m'; chart2.height=7; chart2.width=12
chart2.add_data(Reference(wb['Cash_Flow'],min_col=13,max_col=14,min_row=4,max_row=16), titles_from_data=True)
chart2.set_categories(Reference(wb['Cash_Flow'],min_col=1,min_row=5,max_row=16)); chart2.style=10
ws.add_chart(chart2,'I14')
chart3=LineChart(); chart3.title='Revenue Forecast by Scenario'; chart3.y_axis.title='INR m'; chart3.height=7; chart3.width=12
chart3.add_data(Reference(wb['Revenue_Forecast'],min_col=6,max_col=8,min_row=4,max_row=16), titles_from_data=True)
chart3.set_categories(Reference(wb['Revenue_Forecast'],min_col=1,min_row=5,max_row=16)); chart3.style=12
ws.add_chart(chart3,'A30')
chart4=BarChart(); chart4.type='col'; chart4.title='Department Budget vs Actual'; chart4.y_axis.title='INR m'; chart4.height=7; chart4.width=12
chart4.add_data(Reference(wb['Budget'],min_col=3,max_col=5,min_row=4,max_row=12), titles_from_data=True)
chart4.set_categories(Reference(wb['Budget'],min_col=1,min_row=5,max_row=12)); chart4.style=11
ws.add_chart(chart4,'I30')
# Scenario summary table for visual
ws['A46']='Scenario Closing Cash'; ws['A46'].font=Font(size=12,bold=True,color=TEAL)
for c,h in enumerate(['Scenario','Closing Cash','Minimum Cash','Alert'],1): ws.cell(47,c,h)
header_row(ws,47,1,4,fill=BLUE)
for r,sc in enumerate(['Best Case','Base Case','Worst Case'],48):
    row_src={'Best Case':5,'Base Case':6,'Worst Case':7}[sc]
    ws.cell(r,1,sc); ws.cell(r,2,f'=Scenario_Analysis!$F${row_src}'); ws.cell(r,3,f'=Scenario_Analysis!$G${row_src}'); ws.cell(r,4,f'=IF(C{r}<Assumptions!$B$10,"Critical",IF(C{r}<Assumptions!$B$10*1.25,"Watch","Healthy"))')
    ws.cell(r,2).number_format=currency_fmt; ws.cell(r,3).number_format=currency_fmt
style_body(ws,48,50,1,4)
# Recommendations
ws['I46']='Management Recommendations'; ws['I46'].font=Font(size=12,bold=True,color=TEAL)
recommendations=[
    'Accelerate 31+ day AR collections before approving discretionary spend.',
    'Keep minimum cash reserve at INR 12m and pre-approve a contingency facility.',
    'Defer worst-case CapEx and protect payroll, cloud and customer delivery costs.',
    'Review Marketing and HR variances monthly with accountable cost-center owners.',
]
for r,txt in enumerate(recommendations,47):
    ws.merge_cells(start_row=r,start_column=9,end_row=r,end_column=16); ws.cell(r,9,'• '+txt); ws.cell(r,9).alignment=Alignment(wrap_text=True,vertical='top'); ws.cell(r,9).fill=PatternFill('solid',fgColor=LIGHT_GREY)
for col in range(1,17): ws.column_dimensions[get_column_letter(col)].width=15
for col in ['A','C','E','G','I','K','M','O']: ws.column_dimensions[col].width=14
ws.freeze_panes='A4'

# Add a named range for the KPI formula reference used in KPI_Summary.
from openpyxl.workbook.defined_name import DefinedName
wb.defined_names.add(DefinedName('Burn_Rate', attr_text="'KPI_Summary'!$B$21"))
# Data validation for scenario fields / input cells.
for sheet_name in ['Assumptions']:
    ws = wb[sheet_name]
    dv = DataValidation(type='decimal', operator='between', formula1='0', formula2='1000', allow_blank=True)
    dv.promptTitle='Enter a numeric assumption'; dv.prompt='Use a decimal or percentage as shown in the cell format.'; dv.error='Please enter a valid number.'
    ws.add_data_validation(dv); dv.add('B5:B32')
# Conditional formatting on Cash Flow alerts and closing cash.
ws=wb['Cash_Flow']
for rng in ['O5:O16','AE5:AE16','AU5:AU16']:
    ws.conditional_formatting.add(rng, FormulaRule(formula=[f'{rng.split(":")[0]}="Critical"'], fill=PatternFill('solid',fgColor=LIGHT_RED)))
    ws.conditional_formatting.add(rng, FormulaRule(formula=[f'{rng.split(":")[0]}="Watch"'], fill=PatternFill('solid',fgColor=YELLOW)))
# Freeze, tab colors
for s in wb.worksheets:
    s.sheet_properties.pageSetUpPr.fitToPage = True
    s.sheet_view.zoomScale = 90
    s.sheet_properties.tabColor = {'Dashboard':TEAL,'Raw_Data':BLUE,'Assumptions':GOLD,'Cash_Flow':TEAL}.get(s.title, BLUE)

xlsx_path = ROOT / 'financial-model' / 'cash_flow_forecasting_budgeting_model.xlsx'
wb.calculation.fullCalcOnLoad = True
wb.calculation.forceFullCalc = True
wb.calculation.calcMode = 'auto'
wb.save(xlsx_path)

# -----------------------------
# 4) Static dashboard image (portfolio screenshot)
# -----------------------------
# Use computed values to make a readable portfolio-ready image.
base_res = scenario_rows['Base Case']; best_res = scenario_rows['Best Case']; worst_res = scenario_rows['Worst Case']
fig = plt.figure(figsize=(18, 11), facecolor='#F6F8FB')
fig.suptitle('CASH FLOW FORECASTING & BUDGETING DASHBOARD', x=0.04, y=0.985, ha='left', fontsize=21, fontweight='bold', color='#17324D')
fig.text(0.04, 0.955, f'{company}  |  FY2026 outlook  |  INR millions  |  Base case focus', fontsize=10.5, color='#667085')
# KPI cards
kpi_values=[('Revenue',sum(x['base_revenue'] for x in rev_forecast)),('Cash Inflow',sum(x['inflow'] for x in base_res)),('Cash Outflow',sum(x['outflow'] for x in base_res)),('Net Cash Flow',sum(x['net'] for x in base_res)),('Closing Cash',base_res[-1]['closing']),('Budget Var %', (sum([51.5,103.9,67.3,18.35,14.05,9.2,13.5])-267)/267*100)]
lefts=np.linspace(0.04,0.82,6)
for (label,val),x in zip(kpi_values,lefts):
    ax=fig.add_axes([x,0.855,0.135,0.085]); ax.axis('off')
    patch=FancyBboxPatch((0,0),1,1,boxstyle='round,pad=0.02,rounding_size=0.03',facecolor='white',edgecolor='#D9E2EC',linewidth=1.0)
    ax.add_patch(patch); ax.text(0.06,0.75,label,fontsize=9,color='#667085',fontweight='bold'); ax.text(0.06,0.35,f'{val:,.1f}' + ('%' if label=='Budget Var %' else ''),fontsize=18,color='#17324D',fontweight='bold'); ax.text(0.06,0.12,'INR m' if label!='Budget Var %' else 'FY2025',fontsize=8,color='#667085')
# Main charts
ax1=fig.add_axes([0.04,0.51,0.43,0.29]); ax1.set_facecolor('white')
ax1.plot(month_short,[x['inflow'] for x in base_res],color='#0E9F9A',marker='o',linewidth=2.5,label='Cash inflow')
ax1.plot(month_short,[x['outflow'] for x in base_res],color='#D64545',marker='o',linewidth=2.5,label='Cash outflow')
ax1.set_title('Monthly Cash Inflow vs Outflow',loc='left',fontsize=12,fontweight='bold',color='#17324D'); ax1.grid(axis='y',alpha=.2); ax1.legend(frameon=False,ncol=2,loc='upper left'); ax1.tick_params(labelsize=8); ax1.spines[['top','right']].set_visible(False)
ax2=fig.add_axes([0.53,0.51,0.43,0.29]); ax2.set_facecolor('white')
ax2.plot(month_short,[x['closing'] for x in base_res],color='#1F5A85',marker='o',linewidth=2.5,label='Base closing cash')
ax2.plot(month_short,[x['closing'] for x in best_res],color='#2E8B57',linewidth=1.8,label='Best case')
ax2.plot(month_short,[x['closing'] for x in worst_res],color='#D64545',linewidth=1.8,label='Worst case')
ax2.axhline(12,color='#F4B942',linestyle='--',linewidth=1.6,label='Minimum reserve')
ax2.set_title('Closing Cash & Scenario Range',loc='left',fontsize=12,fontweight='bold',color='#17324D'); ax2.grid(axis='y',alpha=.2); ax2.legend(frameon=False,ncol=2,fontsize=8,loc='upper left'); ax2.tick_params(labelsize=8); ax2.spines[['top','right']].set_visible(False)
# Bottom panels
ax3=fig.add_axes([0.04,0.12,0.28,0.29]); ax3.set_facecolor('white')
labels=['Marketing','HR','Operations','IT','Admin','Finance','CapEx']; budget_vals=[48,100,65,18,14,10,12]; actual_vals=[51.5,103.9,67.3,18.35,14.05,9.2,13.5]
y=np.arange(len(labels)); ax3.barh(y+0.18,budget_vals,height=.32,color='#B8D9E9',label='Budget'); ax3.barh(y-0.18,actual_vals,height=.32,color='#1F5A85',label='Actual'); ax3.set_yticks(y,labels,fontsize=8); ax3.invert_yaxis(); ax3.set_title('Budget vs Actual',loc='left',fontsize=12,fontweight='bold',color='#17324D'); ax3.grid(axis='x',alpha=.2); ax3.legend(frameon=False,fontsize=8); ax3.tick_params(axis='x',labelsize=8); ax3.spines[['top','right']].set_visible(False)
ax4=fig.add_axes([0.36,0.12,0.25,0.29]); ax4.set_facecolor('white')
aging_labels=list(ar_aging.keys()); aging_vals=list(ar_aging.values()); colors=['#0E9F9A','#5DB7B4','#F4B942','#D64545']; ax4.bar(aging_labels,aging_vals,color=colors); ax4.set_title('AR Aging',loc='left',fontsize=12,fontweight='bold',color='#17324D'); ax4.set_ylabel('INR m',fontsize=8); ax4.tick_params(axis='x',rotation=30,labelsize=8); ax4.tick_params(axis='y',labelsize=8); ax4.grid(axis='y',alpha=.2); ax4.spines[['top','right']].set_visible(False)
ax5=fig.add_axes([0.65,0.12,0.31,0.29]); ax5.axis('off'); ax5.set_facecolor('white')
ax5.text(0.02,0.92,'FINANCIAL RISK ALERTS',fontsize=12,fontweight='bold',color='#17324D')
alert_lines=[('BASE','April closes at INR %.1fm, only slightly above reserve.'%min(x['closing'] for x in base_res),'#F4B942'),('WORST','Cash falls below reserve in January; funding gap escalates.','#D64545'),('AR','INR %.1fm overdue receivables require collection action.'%sum(list(ar_aging.values())[1:]),'#D64545'),('ACTION','Defer discretionary CapEx and protect payroll/cloud spend.','#0E9F9A')]
for j,(tag,txt,color) in enumerate(alert_lines):
    y0=.74-j*.18; ax5.add_patch(FancyBboxPatch((.02,y0-.07),.17,.1,boxstyle='round,pad=.02',facecolor=color,edgecolor=color)); ax5.text(.105,y0-.02,tag,ha='center',va='center',fontsize=8,color='white',fontweight='bold'); ax5.text(.23,y0-.02,txt,va='center',fontsize=8.5,color='#344054',wrap=True)
fig.savefig(ROOT/'screenshots'/'cash_flow_dashboard.png',dpi=180,bbox_inches='tight')
plt.close(fig)

# Forecast snapshot CSVs for easy Power BI import.
forecast_rows=[]
for i,m in enumerate([f'{x}-2026' for x in month_short]):
    forecast_rows.append([m,rev_forecast[i]['base_revenue'],rev_forecast[i]['best_revenue'],rev_forecast[i]['worst_revenue'],base_res[i]['inflow'],base_res[i]['outflow'],base_res[i]['net'],base_res[i]['closing'],best_res[i]['closing'],worst_res[i]['closing'],base_res[i]['alert']])
with (ROOT/'forecast'/'monthly_forecast_snapshot.csv').open('w',newline='',encoding='utf-8') as f:
    writer=csv.writer(f); writer.writerow(['Month','Base Revenue','Best Revenue','Worst Revenue','Base Cash Inflow','Base Cash Outflow','Base Net Cash Flow','Base Closing Cash','Best Closing Cash','Worst Closing Cash','Base Alert']); writer.writerows([[m]+[round(x,2) if isinstance(x,(int,float)) else x for x in vals] for m,*vals in forecast_rows])
with (ROOT/'budget-analysis'/'budget_vs_actual_snapshot.csv').open('w',newline='',encoding='utf-8') as f:
    writer=csv.writer(f); writer.writerow(['Department','Budget FY2025','Actual FY2025','Variance Actual - Budget','Variance %'])
    budget_actuals={'Sales':232.8,'Marketing':51.5,'HR':103.9,'Operations':67.3,'IT':18.35,'Administration':14.05,'Finance':9.2,'Capital Expenditure':13.5}
    budget_values={'Sales':225,'Marketing':48,'HR':100,'Operations':65,'IT':18,'Administration':14,'Finance':10,'Capital Expenditure':12}
    for dept in budget_values: writer.writerow([dept,budget_values[dept],budget_actuals[dept],round(budget_actuals[dept]-budget_values[dept],2),round((budget_actuals[dept]-budget_values[dept])/budget_values[dept],4)])

# Save a compact model data dictionary.
with (ROOT/'data'/'data_dictionary.csv').open('w',newline='',encoding='utf-8') as f:
    writer=csv.writer(f); writer.writerow(['Field','Definition','Validation Rule'])
    definitions={
        'Transaction ID':('Unique transaction key','Unique and non-blank'),
        'Date':('Transaction date','Valid date between 01-Jan-2025 and 31-Dec-2025'),
        'Month':('Month key used for aggregation','Must match mmm-yyyy'),
        'Quarter':('Fiscal quarter','Must be Q1 to Q4'),
        'Department':('Responsible cost center owner','Use approved mapping list'),
        'Category':('Operating, investing or financing activity','Use approved list'),
        'Subcategory':('Detailed cash flow line','Use approved list'),
        'Transaction Type':('Inflow or Outflow','Direction must match cash column'),
        'Cash Inflow':('Positive cash received','>=0; populated only for inflows'),
        'Cash Outflow':('Positive cash paid','>=0; populated only for outflows'),
        'Budget Amount':('Approved plan for the transaction','>=0 unless reversal policy applies'),
        'Actual Amount':('Recorded cash amount','Equals Cash Inflow + Cash Outflow'),
        'Forecast Amount':('Forward-looking run-rate estimate','>=0'),
        'Revenue':('Recognized revenue reference','Populated for revenue transactions'),
        'Operating Expense':('Operating cash cost','Populated for operating outflows'),
        'Capital Expenditure':('Investing cash cost','Populated for investing outflows'),
        'Accounts Receivable':('Month-end receivable snapshot','Non-negative; reconcile to AR aging'),
        'Accounts Payable':('Month-end payable snapshot','Non-negative; reconcile to AP aging'),
        'Opening Cash Balance':('Month opening cash','Reconciles to prior closing cash'),
        'Closing Cash Balance':('Month closing cash','Opening cash + inflows - outflows'),
        'Variance Amount':('Actual minus budget','Reconcile to Actual Amount - Budget Amount'),
        'Variance %':('Actual minus budget divided by budget','Blank/0 if budget is zero'),
        'Payment Status':('Collected, Received, Paid or N/A','Controlled list'),
        'Cost Center':('Owner code','Must map to department'),
        'Remarks':('Business explanation or source note','Required for material exceptions')
    }
    for k,v in definitions.items(): writer.writerow([k,v[0],v[1]])

print(f'Created project at {ROOT}')
print(f'Rows: {len(rows)} | FY2025 ending cash: {current_cash:.2f} | Workbook: {xlsx_path}')
