#!/usr/bin/env python3
"""
Create Excel workbook for Finance homework:
- Monthly returns for AAPL, VFIAX, S&P 500
- Risk-free rate from DGS1MO and Fama-French RF
- Excess returns
- Descriptive statistics
- Market Model (CAPM) regression
- Fama-French 3-Factor + Momentum regression
"""

import csv
import re
from datetime import datetime
from collections import OrderedDict

import openpyxl
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill, numbers
from openpyxl.utils import get_column_letter

# ============================================================
# 1. PARSE ALL DATA
# ============================================================

def parse_yahoo_csv(filename):
    """Parse Yahoo Finance CSV, return dict of (year, month) -> adj_close, skipping dividend rows."""
    data = {}
    with open(filename, 'r') as f:
        reader = csv.reader(f)
        header = next(reader)
        for row in reader:
            date_str = row[0]
            if 'Dividend' in date_str or 'Distribution' in date_str:
                continue
            # Parse "Mar 1 2026" format - but also handle partial month like "Mar 11 2026"
            parts = date_str.split()
            if len(parts) < 3:
                continue
            month_str = parts[0]
            year = int(parts[-1])
            months = {'Jan':1,'Feb':2,'Mar':3,'Apr':4,'May':5,'Jun':6,
                      'Jul':7,'Aug':8,'Sep':9,'Oct':10,'Nov':11,'Dec':12}
            month = months.get(month_str, 0)
            if month == 0:
                continue
            adj_close = row[5]
            if adj_close == '' or adj_close == '-':
                adj_close = row[4]  # use Close if no Adj Close
            try:
                val = float(adj_close)
            except:
                continue
            key = (year, month)
            # Keep the "1st of month" entry (which is the end-of-month close for that month)
            # If multiple entries for same month, prefer the "1" day entry
            day = int(parts[1]) if len(parts) >= 3 else 1
            if key not in data or day <= 1:
                data[key] = val
    return data


def parse_ff_factors(filename):
    """Parse Fama-French factors file, return dict of YYYYMM -> {factor: value}."""
    data = {}
    with open(filename, 'r') as f:
        lines = f.readlines()

    in_monthly = False
    for line in lines:
        line = line.rstrip()
        # Detect header line for monthly data
        if 'Mkt-RF' in line and 'SMB' in line and 'HML' in line:
            in_monthly = True
            continue
        if 'Annual' in line:
            break
        if not in_monthly:
            continue
        parts = line.split()
        if len(parts) >= 5:
            try:
                yyyymm = int(parts[0])
                if yyyymm < 100000:  # annual data (4 digits)
                    continue
                data[yyyymm] = {
                    'Mkt-RF': float(parts[1]),
                    'SMB': float(parts[2]),
                    'HML': float(parts[3]),
                    'RF': float(parts[4])
                }
            except:
                continue
    return data


def parse_momentum(filename):
    """Parse momentum factor file, return dict of YYYYMM -> Mom value."""
    data = {}
    with open(filename, 'r') as f:
        lines = f.readlines()

    in_monthly = False
    for line in lines:
        line = line.rstrip()
        if 'Mom' in line and len(line.strip().split()) <= 2:
            in_monthly = True
            continue
        if 'Annual' in line:
            break
        if not in_monthly:
            continue
        parts = line.split()
        if len(parts) >= 2:
            try:
                yyyymm = int(parts[0])
                if yyyymm < 100000:
                    continue
                data[yyyymm] = float(parts[1])
            except:
                continue
    return data


def parse_dgs1mo(filename):
    """Parse DGS1MO Excel file, return dict of (year, month) -> avg monthly rate in %."""
    wb = openpyxl.load_workbook(filename, data_only=True)
    ws = wb['Daily']

    monthly_rates = {}
    monthly_counts = {}

    for row in ws.iter_rows(min_row=2, values_only=True):
        date_val, rate = row[0], row[1]
        if date_val is None or rate is None:
            continue
        if isinstance(rate, str):
            if rate == '.' or rate == '#N/A':
                continue
            try:
                rate = float(rate)
            except:
                continue

        if isinstance(date_val, datetime):
            y, m = date_val.year, date_val.month
        else:
            continue

        key = (y, m)
        if key not in monthly_rates:
            monthly_rates[key] = 0.0
            monthly_counts[key] = 0
        monthly_rates[key] += rate
        monthly_counts[key] += 1

    # Average daily rate per month
    result = {}
    for key in monthly_rates:
        result[key] = monthly_rates[key] / monthly_counts[key]

    return result


# Parse all data
print("Parsing data files...")
aapl = parse_yahoo_csv('AAPL_Historical_Prices.csv')
vfiax = parse_yahoo_csv('VFIAX_Historical_Prices.csv')
gspc = parse_yahoo_csv('GSPC_Historical_Prices.csv')
ff_factors = parse_ff_factors('F-F_Research_Data_Factors.txt')
momentum = parse_momentum('F-F_Momentum_Factor.txt')
dgs1mo = parse_dgs1mo('DGS1MO.xlsx')

# Determine common date range (months with data for all three assets)
# We need at least 2 consecutive months to compute returns
all_months = sorted(set(aapl.keys()) & set(vfiax.keys()) & set(gspc.keys()))
print(f"Common months with price data: {len(all_months)}")
print(f"Range: {all_months[0]} to {all_months[-1]}")

# Return months = all months except the first (need prior month for return calc)
return_months = all_months[1:]
print(f"Return months: {len(return_months)}")

# ============================================================
# 2. CREATE EXCEL WORKBOOK
# ============================================================

wb = Workbook()

# Styles
header_font = Font(bold=True, size=11)
title_font = Font(bold=True, size=14)
subtitle_font = Font(bold=True, size=12)
pct_format = '0.00%'
pct4_format = '0.0000%'
num2_format = '0.00'
num4_format = '0.0000'
thin_border = Border(
    left=Side(style='thin'), right=Side(style='thin'),
    top=Side(style='thin'), bottom=Side(style='thin')
)
header_fill = PatternFill(start_color='4472C4', end_color='4472C4', fill_type='solid')
header_font_white = Font(bold=True, size=11, color='FFFFFF')
light_fill = PatternFill(start_color='D6E4F0', end_color='D6E4F0', fill_type='solid')

def style_header(ws, row, cols, fill=None, font=None):
    for c in range(1, cols + 1):
        cell = ws.cell(row=row, column=c)
        cell.font = font or header_font_white
        cell.fill = fill or header_fill
        cell.alignment = Alignment(horizontal='center', wrap_text=True)
        cell.border = thin_border

def auto_width(ws, max_col=None):
    for col in ws.columns:
        max_length = 0
        col_letter = get_column_letter(col[0].column)
        for cell in col:
            if cell.value:
                max_length = max(max_length, len(str(cell.value)))
        ws.column_dimensions[col_letter].width = min(max_length + 3, 20)


# ============================================================
# SHEET 1: Monthly Data & Returns
# ============================================================
ws1 = wb.active
ws1.title = "Monthly Returns"

# Headers
headers1 = ['Date', 'AAPL Price', 'VFIAX NAV', 'S&P 500',
            'AAPL Return', 'VFIAX Return', 'S&P 500 Return',
            'RF (FF, %/mo)', 'RF (DGS1MO, %/yr)',
            'AAPL Excess', 'VFIAX Excess', 'S&P 500 Excess',
            'Mkt-RF (FF)', 'SMB (FF)', 'HML (FF)', 'Mom']

ws1.cell(row=1, column=1, value="Monthly Price Data, Returns & Fama-French Factors").font = title_font
ws1.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(headers1))

row = 3
for i, h in enumerate(headers1, 1):
    ws1.cell(row=row, column=i, value=h)
style_header(ws1, row, len(headers1))

# Data rows
data_start_row = 4
prev_month = all_months[0]

for idx, month_key in enumerate(all_months):
    r = data_start_row + idx
    y, m = month_key

    # Date
    ws1.cell(row=r, column=1, value=datetime(y, m, 1))
    ws1.cell(row=r, column=1).number_format = 'MMM YYYY'
    ws1.cell(row=r, column=1).border = thin_border

    # Prices
    ws1.cell(row=r, column=2, value=aapl[month_key])
    ws1.cell(row=r, column=2).number_format = '#,##0.00'
    ws1.cell(row=r, column=2).border = thin_border

    ws1.cell(row=r, column=3, value=vfiax[month_key])
    ws1.cell(row=r, column=3).number_format = '#,##0.00'
    ws1.cell(row=r, column=3).border = thin_border

    ws1.cell(row=r, column=4, value=gspc[month_key])
    ws1.cell(row=r, column=4).number_format = '#,##0.00'
    ws1.cell(row=r, column=4).border = thin_border

    # Returns (from 2nd month onward) using Excel formulas
    if idx > 0:
        prev_r = r - 1
        # AAPL Return = (Price_t / Price_t-1) - 1
        ws1.cell(row=r, column=5).value = f'=B{r}/B{prev_r}-1'
        ws1.cell(row=r, column=5).number_format = pct_format
        ws1.cell(row=r, column=5).border = thin_border

        # VFIAX Return
        ws1.cell(row=r, column=6).value = f'=C{r}/C{prev_r}-1'
        ws1.cell(row=r, column=6).number_format = pct_format
        ws1.cell(row=r, column=6).border = thin_border

        # S&P 500 Return
        ws1.cell(row=r, column=7).value = f'=D{r}/D{prev_r}-1'
        ws1.cell(row=r, column=7).number_format = pct_format
        ws1.cell(row=r, column=7).border = thin_border

    # Fama-French RF (monthly, in %)
    yyyymm = y * 100 + m
    if yyyymm in ff_factors:
        rf_ff = ff_factors[yyyymm]['RF']
        ws1.cell(row=r, column=8, value=rf_ff / 100)  # Convert to decimal
        ws1.cell(row=r, column=8).number_format = pct4_format
        ws1.cell(row=r, column=8).border = thin_border

    # DGS1MO (annualized rate)
    if month_key in dgs1mo:
        ws1.cell(row=r, column=9, value=dgs1mo[month_key] / 100)  # Convert to decimal
        ws1.cell(row=r, column=9).number_format = pct4_format
        ws1.cell(row=r, column=9).border = thin_border

    # Excess returns (Return - RF_monthly) - only from 2nd month
    if idx > 0:
        # AAPL Excess = AAPL Return - RF(FF)
        ws1.cell(row=r, column=10).value = f'=E{r}-H{r}'
        ws1.cell(row=r, column=10).number_format = pct_format
        ws1.cell(row=r, column=10).border = thin_border

        # VFIAX Excess
        ws1.cell(row=r, column=11).value = f'=F{r}-H{r}'
        ws1.cell(row=r, column=11).number_format = pct_format
        ws1.cell(row=r, column=11).border = thin_border

        # S&P 500 Excess
        ws1.cell(row=r, column=12).value = f'=G{r}-H{r}'
        ws1.cell(row=r, column=12).number_format = pct_format
        ws1.cell(row=r, column=12).border = thin_border

    # FF Factors (Mkt-RF, SMB, HML)
    if yyyymm in ff_factors:
        ws1.cell(row=r, column=13, value=ff_factors[yyyymm]['Mkt-RF'] / 100)
        ws1.cell(row=r, column=13).number_format = pct_format
        ws1.cell(row=r, column=13).border = thin_border

        ws1.cell(row=r, column=14, value=ff_factors[yyyymm]['SMB'] / 100)
        ws1.cell(row=r, column=14).number_format = pct_format
        ws1.cell(row=r, column=14).border = thin_border

        ws1.cell(row=r, column=15, value=ff_factors[yyyymm]['HML'] / 100)
        ws1.cell(row=r, column=15).number_format = pct_format
        ws1.cell(row=r, column=15).border = thin_border

    # Momentum
    if yyyymm in momentum:
        ws1.cell(row=r, column=16, value=momentum[yyyymm] / 100)
        ws1.cell(row=r, column=16).number_format = pct_format
        ws1.cell(row=r, column=16).border = thin_border

last_data_row = data_start_row + len(all_months) - 1
first_return_row = data_start_row + 1  # First row with return data

auto_width(ws1)

# ============================================================
# SHEET 2: Descriptive Statistics
# ============================================================
ws2 = wb.create_sheet("Descriptive Statistics")
ws2.cell(row=1, column=1, value="Descriptive Statistics of Monthly Returns").font = title_font
ws2.merge_cells(start_row=1, start_column=1, end_row=1, end_column=7)

stats_headers = ['Statistic', 'AAPL', 'VFIAX', 'S&P 500', 'AAPL Excess', 'VFIAX Excess', 'S&P 500 Excess']
for i, h in enumerate(stats_headers, 1):
    ws2.cell(row=3, column=i, value=h)
style_header(ws2, 3, len(stats_headers))

# Return columns in Sheet1: E=AAPL, F=VFIAX, G=S&P500
# Excess columns: J=AAPL, K=VFIAX, L=S&P500
ret_cols = ['E', 'F', 'G', 'J', 'K', 'L']
ret_range = [f"'Monthly Returns'!{c}{first_return_row}:{c}{last_data_row}" for c in ret_cols]

stats = [
    ('Mean', 'AVERAGE'),
    ('Median', 'MEDIAN'),
    ('Std Dev', 'STDEV.S'),
    ('Variance', 'VAR.S'),
    ('Min', 'MIN'),
    ('Max', 'MAX'),
    ('Skewness', 'SKEW'),
    ('Kurtosis', 'KURT'),
    ('Count', 'COUNT'),
]

for i, (label, func) in enumerate(stats):
    r = 4 + i
    ws2.cell(row=r, column=1, value=label)
    ws2.cell(row=r, column=1).font = header_font
    ws2.cell(row=r, column=1).border = thin_border
    for j, rng in enumerate(ret_range):
        ws2.cell(row=r, column=j+2).value = f'={func}({rng})'
        if label == 'Count':
            ws2.cell(row=r, column=j+2).number_format = '0'
        else:
            ws2.cell(row=r, column=j+2).number_format = pct4_format if label != 'Skewness' and label != 'Kurtosis' else num4_format
        ws2.cell(row=r, column=j+2).border = thin_border

# Add annualized stats
ann_row = 4 + len(stats) + 1
ws2.cell(row=ann_row, column=1, value="Annualized Statistics").font = subtitle_font
ws2.merge_cells(start_row=ann_row, start_column=1, end_row=ann_row, end_column=7)

ann_headers = ['Statistic', 'AAPL', 'VFIAX', 'S&P 500']
for i, h in enumerate(ann_headers, 1):
    ws2.cell(row=ann_row+1, column=i, value=h)
style_header(ws2, ann_row+1, len(ann_headers))

# Annualized Return = (1+mean)^12 - 1
# Annualized Std Dev = monthly_std * sqrt(12)
mean_row = 4  # row of Mean
std_row = 6   # row of Std Dev

ws2.cell(row=ann_row+2, column=1, value='Annualized Return')
ws2.cell(row=ann_row+2, column=1).font = header_font
ws2.cell(row=ann_row+2, column=1).border = thin_border
for j in range(3):
    col = j + 2
    ws2.cell(row=ann_row+2, column=col).value = f'=(1+{get_column_letter(col)}{mean_row})^12-1'
    ws2.cell(row=ann_row+2, column=col).number_format = pct_format
    ws2.cell(row=ann_row+2, column=col).border = thin_border

ws2.cell(row=ann_row+3, column=1, value='Annualized Std Dev')
ws2.cell(row=ann_row+3, column=1).font = header_font
ws2.cell(row=ann_row+3, column=1).border = thin_border
for j in range(3):
    col = j + 2
    ws2.cell(row=ann_row+3, column=col).value = f'={get_column_letter(col)}{std_row}*SQRT(12)'
    ws2.cell(row=ann_row+3, column=col).number_format = pct_format
    ws2.cell(row=ann_row+3, column=col).border = thin_border

ws2.cell(row=ann_row+4, column=1, value='Sharpe Ratio')
ws2.cell(row=ann_row+4, column=1).font = header_font
ws2.cell(row=ann_row+4, column=1).border = thin_border
# Sharpe = Annualized Excess / Annualized Std
# Use excess mean from columns E, F, G (cols 5,6,7 in stats sheet)
for j in range(3):
    col = j + 2
    excess_col = j + 5  # columns E, F, G for excess stats
    ws2.cell(row=ann_row+4, column=col).value = f'=((1+{get_column_letter(excess_col)}{mean_row})^12-1)/({get_column_letter(col)}{std_row}*SQRT(12))'
    ws2.cell(row=ann_row+4, column=col).number_format = num4_format
    ws2.cell(row=ann_row+4, column=col).border = thin_border

# Correlation matrix
corr_row = ann_row + 7
ws2.cell(row=corr_row, column=1, value="Correlation Matrix").font = subtitle_font
ws2.merge_cells(start_row=corr_row, start_column=1, end_row=corr_row, end_column=4)

corr_headers = ['', 'AAPL', 'VFIAX', 'S&P 500']
for i, h in enumerate(corr_headers, 1):
    ws2.cell(row=corr_row+1, column=i, value=h)
style_header(ws2, corr_row+1, len(corr_headers))

assets = ['AAPL', 'VFIAX', 'S&P 500']
ret_cols_corr = ['E', 'F', 'G']

for i in range(3):
    r = corr_row + 2 + i
    ws2.cell(row=r, column=1, value=assets[i])
    ws2.cell(row=r, column=1).font = header_font
    ws2.cell(row=r, column=1).border = thin_border
    for j in range(3):
        rng1 = f"'Monthly Returns'!{ret_cols_corr[i]}{first_return_row}:{ret_cols_corr[i]}{last_data_row}"
        rng2 = f"'Monthly Returns'!{ret_cols_corr[j]}{first_return_row}:{ret_cols_corr[j]}{last_data_row}"
        ws2.cell(row=r, column=j+2).value = f'=CORREL({rng1},{rng2})'
        ws2.cell(row=r, column=j+2).number_format = num4_format
        ws2.cell(row=r, column=j+2).border = thin_border

auto_width(ws2)

# ============================================================
# SHEET 3: Market Model (CAPM) Regression
# ============================================================
ws3 = wb.create_sheet("CAPM Regression")
ws3.cell(row=1, column=1, value="Market Model (CAPM) Regression").font = title_font
ws3.merge_cells(start_row=1, start_column=1, end_row=1, end_column=6)

ws3.cell(row=2, column=1, value="R_i - R_f = α + β(R_m - R_f) + ε").font = Font(italic=True, size=11)

# For AAPL
ws3.cell(row=4, column=1, value="AAPL Market Model").font = subtitle_font
capm_headers = ['Statistic', 'Value']
for i, h in enumerate(capm_headers, 1):
    ws3.cell(row=5, column=i, value=h)
style_header(ws3, 5, 2)

# Using LINEST: =LINEST(y_range, x_range, TRUE, TRUE)
# y = AAPL excess return (col J), x = Mkt-RF (col M)
y_aapl = f"'Monthly Returns'!J{first_return_row}:J{last_data_row}"
y_vfiax = f"'Monthly Returns'!K{first_return_row}:K{last_data_row}"
x_mkt = f"'Monthly Returns'!M{first_return_row}:M{last_data_row}"

# AAPL CAPM
capm_stats_aapl = [
    ('Alpha (α)', f'=INDEX(LINEST({y_aapl},{x_mkt},TRUE,TRUE),1,2)'),
    ('Beta (β)', f'=INDEX(LINEST({y_aapl},{x_mkt},TRUE,TRUE),1,1)'),
    ('R-squared', f'=RSQ({y_aapl},{x_mkt})'),
    ('Std Error (α)', f'=INDEX(LINEST({y_aapl},{x_mkt},TRUE,TRUE),2,2)'),
    ('Std Error (β)', f'=INDEX(LINEST({y_aapl},{x_mkt},TRUE,TRUE),2,1)'),
    ('t-stat (α)', f'=B6/B9'),
    ('t-stat (β)', f'=B7/B10'),
    ('F-statistic', f'=INDEX(LINEST({y_aapl},{x_mkt},TRUE,TRUE),4,1)'),
    ('Observations', f'=COUNT({y_aapl})'),
]

for i, (label, formula) in enumerate(capm_stats_aapl):
    r = 6 + i
    ws3.cell(row=r, column=1, value=label)
    ws3.cell(row=r, column=1).font = header_font
    ws3.cell(row=r, column=1).border = thin_border
    ws3.cell(row=r, column=2).value = formula
    ws3.cell(row=r, column=2).number_format = num4_format
    ws3.cell(row=r, column=2).border = thin_border

# For VFIAX
vfiax_start = 17
ws3.cell(row=vfiax_start, column=1, value="VFIAX Market Model").font = subtitle_font

for i, h in enumerate(capm_headers, 1):
    ws3.cell(row=vfiax_start+1, column=i, value=h)
style_header(ws3, vfiax_start+1, 2)

capm_stats_vfiax = [
    ('Alpha (α)', f'=INDEX(LINEST({y_vfiax},{x_mkt},TRUE,TRUE),1,2)'),
    ('Beta (β)', f'=INDEX(LINEST({y_vfiax},{x_mkt},TRUE,TRUE),1,1)'),
    ('R-squared', f'=RSQ({y_vfiax},{x_mkt})'),
    ('Std Error (α)', f'=INDEX(LINEST({y_vfiax},{x_mkt},TRUE,TRUE),2,2)'),
    ('Std Error (β)', f'=INDEX(LINEST({y_vfiax},{x_mkt},TRUE,TRUE),2,1)'),
    ('t-stat (α)', f'=B{vfiax_start+2}/B{vfiax_start+5}'),
    ('t-stat (β)', f'=B{vfiax_start+3}/B{vfiax_start+6}'),
    ('F-statistic', f'=INDEX(LINEST({y_vfiax},{x_mkt},TRUE,TRUE),4,1)'),
    ('Observations', f'=COUNT({y_vfiax})'),
]

for i, (label, formula) in enumerate(capm_stats_vfiax):
    r = vfiax_start + 2 + i
    ws3.cell(row=r, column=1, value=label)
    ws3.cell(row=r, column=1).font = header_font
    ws3.cell(row=r, column=1).border = thin_border
    ws3.cell(row=r, column=2).value = formula
    ws3.cell(row=r, column=2).number_format = num4_format
    ws3.cell(row=r, column=2).border = thin_border

auto_width(ws3)

# ============================================================
# SHEET 4: Fama-French 3-Factor + Momentum
# ============================================================
ws4 = wb.create_sheet("FF4 Regression")
ws4.cell(row=1, column=1, value="Fama-French 3-Factor + Momentum Regression").font = title_font
ws4.merge_cells(start_row=1, start_column=1, end_row=1, end_column=6)

ws4.cell(row=2, column=1, value="R_i - R_f = α + β₁(Mkt-RF) + β₂(SMB) + β₃(HML) + β₄(Mom) + ε").font = Font(italic=True, size=11)

# x ranges for 4-factor model
x_smb = f"'Monthly Returns'!N{first_return_row}:N{last_data_row}"
x_hml = f"'Monthly Returns'!O{first_return_row}:O{last_data_row}"
x_mom = f"'Monthly Returns'!P{first_return_row}:P{last_data_row}"

# LINEST with multiple X: =LINEST(y, x1:x4, TRUE, TRUE) - but X must be contiguous cols M:P
x_all = f"'Monthly Returns'!M{first_return_row}:P{last_data_row}"

# AAPL FF4
ws4.cell(row=4, column=1, value="AAPL Fama-French 4-Factor").font = subtitle_font

ff4_headers = ['Statistic', 'Value']
for i, h in enumerate(ff4_headers, 1):
    ws4.cell(row=5, column=i, value=h)
style_header(ws4, 5, 2)

# LINEST returns: row1=[b4,b3,b2,b1,a], row2=[se4,se3,se2,se1,se_a], row3=[R2,se_y,...], row4=[F,df,...]
ff4_aapl = [
    ('Alpha (α)',     f'=INDEX(LINEST({y_aapl},{x_all},TRUE,TRUE),1,5)'),
    ('β (Mkt-RF)',    f'=INDEX(LINEST({y_aapl},{x_all},TRUE,TRUE),1,4)'),
    ('β (SMB)',       f'=INDEX(LINEST({y_aapl},{x_all},TRUE,TRUE),1,3)'),
    ('β (HML)',       f'=INDEX(LINEST({y_aapl},{x_all},TRUE,TRUE),1,2)'),
    ('β (Mom)',       f'=INDEX(LINEST({y_aapl},{x_all},TRUE,TRUE),1,1)'),
    ('R-squared',     f'=INDEX(LINEST({y_aapl},{x_all},TRUE,TRUE),3,1)'),
    ('SE (α)',        f'=INDEX(LINEST({y_aapl},{x_all},TRUE,TRUE),2,5)'),
    ('SE (Mkt-RF)',   f'=INDEX(LINEST({y_aapl},{x_all},TRUE,TRUE),2,4)'),
    ('SE (SMB)',      f'=INDEX(LINEST({y_aapl},{x_all},TRUE,TRUE),2,3)'),
    ('SE (HML)',      f'=INDEX(LINEST({y_aapl},{x_all},TRUE,TRUE),2,2)'),
    ('SE (Mom)',      f'=INDEX(LINEST({y_aapl},{x_all},TRUE,TRUE),2,1)'),
    ('t-stat (α)',    f'=B6/B12'),
    ('t-stat (Mkt-RF)', f'=B7/B13'),
    ('t-stat (SMB)',  f'=B8/B14'),
    ('t-stat (HML)',  f'=B9/B15'),
    ('t-stat (Mom)',  f'=B10/B16'),
    ('F-statistic',   f'=INDEX(LINEST({y_aapl},{x_all},TRUE,TRUE),4,1)'),
    ('Observations',  f'=COUNT({y_aapl})'),
]

for i, (label, formula) in enumerate(ff4_aapl):
    r = 6 + i
    ws4.cell(row=r, column=1, value=label)
    ws4.cell(row=r, column=1).font = header_font
    ws4.cell(row=r, column=1).border = thin_border
    ws4.cell(row=r, column=2).value = formula
    ws4.cell(row=r, column=2).number_format = num4_format
    ws4.cell(row=r, column=2).border = thin_border

# VFIAX FF4
vf_start = 26
ws4.cell(row=vf_start, column=1, value="VFIAX Fama-French 4-Factor").font = subtitle_font

for i, h in enumerate(ff4_headers, 1):
    ws4.cell(row=vf_start+1, column=i, value=h)
style_header(ws4, vf_start+1, 2)

ff4_vfiax = [
    ('Alpha (α)',     f'=INDEX(LINEST({y_vfiax},{x_all},TRUE,TRUE),1,5)'),
    ('β (Mkt-RF)',    f'=INDEX(LINEST({y_vfiax},{x_all},TRUE,TRUE),1,4)'),
    ('β (SMB)',       f'=INDEX(LINEST({y_vfiax},{x_all},TRUE,TRUE),1,3)'),
    ('β (HML)',       f'=INDEX(LINEST({y_vfiax},{x_all},TRUE,TRUE),1,2)'),
    ('β (Mom)',       f'=INDEX(LINEST({y_vfiax},{x_all},TRUE,TRUE),1,1)'),
    ('R-squared',     f'=INDEX(LINEST({y_vfiax},{x_all},TRUE,TRUE),3,1)'),
    ('SE (α)',        f'=INDEX(LINEST({y_vfiax},{x_all},TRUE,TRUE),2,5)'),
    ('SE (Mkt-RF)',   f'=INDEX(LINEST({y_vfiax},{x_all},TRUE,TRUE),2,4)'),
    ('SE (SMB)',      f'=INDEX(LINEST({y_vfiax},{x_all},TRUE,TRUE),2,3)'),
    ('SE (HML)',      f'=INDEX(LINEST({y_vfiax},{x_all},TRUE,TRUE),2,2)'),
    ('SE (Mom)',      f'=INDEX(LINEST({y_vfiax},{x_all},TRUE,TRUE),2,1)'),
    ('t-stat (α)',    f'=B{vf_start+2}/B{vf_start+8}'),
    ('t-stat (Mkt-RF)', f'=B{vf_start+3}/B{vf_start+9}'),
    ('t-stat (SMB)',  f'=B{vf_start+4}/B{vf_start+10}'),
    ('t-stat (HML)',  f'=B{vf_start+5}/B{vf_start+11}'),
    ('t-stat (Mom)',  f'=B{vf_start+6}/B{vf_start+12}'),
    ('F-statistic',   f'=INDEX(LINEST({y_vfiax},{x_all},TRUE,TRUE),4,1)'),
    ('Observations',  f'=COUNT({y_vfiax})'),
]

for i, (label, formula) in enumerate(ff4_vfiax):
    r = vf_start + 2 + i
    ws4.cell(row=r, column=1, value=label)
    ws4.cell(row=r, column=1).font = header_font
    ws4.cell(row=r, column=1).border = thin_border
    ws4.cell(row=r, column=2).value = formula
    ws4.cell(row=r, column=2).number_format = num4_format
    ws4.cell(row=r, column=2).border = thin_border

auto_width(ws4)

# ============================================================
# SAVE
# ============================================================
output_file = 'Finance_Homework.xlsx'
wb.save(output_file)
print(f"\nExcel file saved: {output_file}")
print(f"Sheets: {wb.sheetnames}")
print(f"Data range: {all_months[0]} to {all_months[-1]} ({len(return_months)} return months)")
