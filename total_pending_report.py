"""
total_pending_report.py — TOTAL PENDING Report Generator
==========================================================
Generates CEO executive summary and audit dataset for all pending shipments
across all branches, grouped by Branch, Facility Type (P, S, A), and Age (Days).
Designed with the official CEO Executive Teal Table aesthetic and intuitive color tiers.

ALL ACTIVE STATUSES INCLUDED (exclude only: 99, 100, 201=Cancelled | 410=Delivered | 520=Returned to Hub)
Test orders excluded via keyword/ID filtering.

Approved Statuses (25 — all active, non-completed):
- Pickup chain    (5): 110, 120, 200, 210, 230
- Transit chain   (5): 300, 302, 310, 311, 306
- At Branch       (2): 309, 400
- Delivery        (5): 401, 402, 420, 430, 402
- Delivery issue  (3): 460, 470, 471
- Resolving       (3): 472, 480, 500
- Return chain    (4): 510, 511, 512, 540

Age Buckets (ACTUAL HOURS - NO GRACE ADJUSTMENTS):
- 0 Days:   0:00 -> 23:59 (< 24.0 hours)
- 1 Day:   24:00 -> 47:59 (>= 24.0 and < 48.0 hours)
- 2 Days:  48:00 -> 71:59 (>= 48.0 and < 72.0 hours)
- 3 Days:  72:00 -> 95:59 (>= 72.0 and < 96.0 hours)
- 4 Days:  96:00 -> 119:59 (>= 96.0 and < 120.0 hours)
- 5 Days: 120:00 -> 167:59 (>= 120.0 and < 168.0 hours, covers 5 to 7 days)
- > 7 Days: >= 168.0 and < 720.0 hours (> 7 Days and < 30 Days, limited to this month)
- Over 30 Days: Excluded (>= 720.0 hours)
"""

import os
import io
import re
import json
import tempfile
from datetime import datetime
from concurrent.futures import ThreadPoolExecutor
import requests
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import shutil


import excel_to_image

# EXCLUDED STATUSES:
# 1. Done / Cancelled orders (410=Delivered, 520=Returned to Hub, 99/100/201=Cancelled)
# 2. Outgoing long-haul truck transit to Mega (210, 230, 300, 302)
# 3. Pickup chain — not yet collected from sender (110, 120, 200)
# NOTE: 306 and 311 at branch are ACTIVE branch pending inventory (Not Assigned to courier)!
# Mega Hub facilities (MEGA1, DVCMEGA1, etc.) are excluded via CURRENT POST OFFICE check below.
EXCLUDED_STATUSES = {
    # Done / Cancelled
    '99',   # Cancelled / Test
    '100',  # Cancelled confirmed
    '201',  # Pickup cancelled
    '410',  # Delivered Successfully ✅ Done
    '520',  # Returned to Hub ✅ Done
    # Pickup chain — uncollected orders
    '110',  # New order / Pickup pending
    '120',  # Assigned pickup rider
    '200',  # Picking up
    # Outgoing truck transit to Mega
    '210',  # Picked up / Handover to driver
    '230',  # In transit to Mega
    '300',  # Dispatched / Transit
    '302',  # Completed loading to Hub
    # Packaging / At Agent (excluded per boss directive, like test bill)
    '310',  # Đóng kiện / Packaging / At Agent
}

FACILITY_COLS = ['Servicepoint', 'Showroom', 'Agent']
DAY_COLS = ['0 Days', '1 Day', '2 Days', '3 Days', '4 Days', '5 Days', '> 7 Days']
DELAY_COLS = ['3 Days', '4 Days', '5 Days', '> 7 Days']
DELAY_COL_NAME = 'Total >= 3 Days'

TS_COLS_PRIORITY = [
    'CURRENT TIME',
    'STATUS 306 AT STORE / AGENT (LAST TIME)',
    'STATUS 306 AT STORE / AGENT FROM HUB (FIRST TIME)',
    'STATUS 302/310 AT RECEIVING STORE / RECEIVING AGENT (FIRST TIME)',
    'STATUS 306  AT ORIGIN HUB (FIRST TIME)',
    'STATUS 210 TIME',
    'CREATED DATE'
]


_po_lookup_cache = None

def _get_po_lookup_map():
    global _po_lookup_cache
    if _po_lookup_cache is None:
        _po_lookup_cache = {}
        csv_candidates = [
            "post_office_lookup.csv",
            os.path.join(os.path.dirname(os.path.abspath(__file__)), "post_office_lookup.csv"),
            os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "post_office_lookup.csv")
        ]
        lookup_file = next((p for p in csv_candidates if os.path.exists(p)), None)
        if lookup_file:
            try:
                df_po = pd.read_csv(lookup_file, encoding="utf-8-sig", dtype=str).fillna("")
                df_po.columns = [str(c).strip().lower() for c in df_po.columns]
                if "current_post_office" in df_po.columns and "post_office_handle" in df_po.columns:
                    for _, r in df_po.iterrows():
                        cur = str(r["current_post_office"]).strip().upper()
                        hnd = str(r["post_office_handle"]).strip().upper()
                        if cur and hnd and cur != "NAN" and hnd != "NAN":
                            _po_lookup_cache[cur] = hnd
            except Exception:
                pass
    return _po_lookup_cache


def resolve_post_office_handle(raw_po: str) -> str:
    c = str(raw_po or '').strip().upper()
    if not c or c == 'NAN':
        return 'UNKNOWN'
    po_map = _get_po_lookup_map()
    return po_map.get(c, c)


def get_branch_code(po_code: str) -> str:
    """
    Map post office / showroom / agent code to parent Branch code.
    First resolves responsible post office handle (e.g. KANS005 -> PNPP012 -> PNP,
    KANA019 -> PNPP004 -> PNP, KANA001 -> KANP001 -> KAN).
    Only 'PNP' retains 'P'; all other provincial branches have 3 letters with no 'P'
    (e.g., BAN, BAT, CHA, KAN, KAM, KOH, KRA, MON, ODD, PRE, PUR, SIE, SIH, SPE, TAK, etc.).
    """
    c = str(po_code or '').strip().upper()
    if not c or c == 'NAN':
        return 'UNKNOWN'
    resolved = resolve_post_office_handle(c)
    if resolved.startswith('PNP'):
        return 'PNP'
    return resolved[:3]


def get_facility_type(po_code: str) -> str:
    """Classify facility as Servicepoint, Showroom, or Agent."""
    c = str(po_code or '').strip().upper()
    if len(c) >= 4:
        ch = c[3]
        if ch == 'S':
            return 'Showroom'
        elif ch == 'A':
            return 'Agent'
        elif ch == 'P':
            return 'Servicepoint'
    return 'Servicepoint'

# REMOVED - No longer needed (no grace period logic)
# def fetch_pending_history_allowances(order_ids: list[str], bearer_token: str = None) -> dict[str, dict[str, bool]]:

# REMOVED - No longer needed (no grace period logic)
# def fetch_pending_history_allowances(order_ids: list[str], bearer_token: str = None) -> dict[str, dict[str, bool]]:
#     """
#     Check real-time tracking trips for candidate pending bills (>= 24h old).
#     Returns a mapping of order_id -> {'has_420': bool, 'has_472': bool}
#     indicating whether status 420 or 472 appeared anywhere in the tracking history log.
#     """
#     [FUNCTION REMOVED - No longer checking history for 420/472]


def process_pending_data(src_path_or_df):
    """
    Load and process TMS data to build the TOTAL PENDING report dataset.
    Returns:
        (summary_df, grand_total_dict, df_pending_detail)
    """
    if isinstance(src_path_or_df, pd.DataFrame):
        df = src_path_or_df.copy()
    else:
        try:
            xl = pd.ExcelFile(src_path_or_df)
            sheet = 'Pending Details' if 'Pending Details' in xl.sheet_names else xl.sheet_names[0]
            df = xl.parse(sheet)
        except Exception:
            try:
                df = pd.read_excel(src_path_or_df, engine='calamine')
            except Exception:
                df = pd.read_excel(src_path_or_df)

    # Standardize column names
    df.columns = [str(c).strip() for c in df.columns]

    # Extract 3-digit status code — same exclusion logic as generate_report.py (push report)
    # Exclude: 99/100/201=Cancelled, 410=Delivered, 520=Returned to Hub, 310=Packaging/Agent
    # Everything else is still pending/active — matches what push report counts
    sc = df['CURRENT STATUS'].astype(str).str.extract(r'^(\d{3})')[0]
    active_mask = sc.notna() & ~sc.isin(EXCLUDED_STATUSES)
    df_p = df[active_mask].copy()
    df_p['STATUS_CODE'] = sc[active_mask]
    df_p = df_p[df_p['STATUS_CODE'] != '310'].copy()

    # Exclude test orders & test post offices (e.g. BANCHI_TEST, TEST PH NOM PENH, test orders)
    test_col = next((c for c in df_p.columns if str(c).strip().lower() in ('is test', 'đơn test', 'don test')), None)
    if test_col:
        df_p = df_p[df_p[test_col].isna() | df_p[test_col].astype(str).str.strip().isin(['', 'nan', 'NaN', '#N/A'])].copy()

    # Exclude any row where 'test' appears in key text fields
    check_cols = [c for c in df_p.columns if any(k in str(c).upper() for k in ('ORDER', 'STATUS', 'POST', 'OFFICE', 'SENDER', 'RECEIVER', 'NOTE', 'REMARK', 'GOODS', 'DESC', 'PRODUCT'))]
    if check_cols:
        test_mask = df_p[check_cols].astype(str).apply(lambda s: s.str.lower().str.contains('test', na=False)).any(axis=1)
        df_p = df_p[~test_mask].copy()

    # Exclude test bills (test_bills.txt) & confirmed shipped bills (shipped_bills.txt) UPFRONT
    try:
        from penalty_report import load_test_bills
        test_bill_ids = load_test_bills()
    except Exception:
        test_bill_ids = set()

    try:
        from shipped_filter import load_confirmed_shipped_ids
        shipped_bill_ids = load_confirmed_shipped_ids()
    except Exception:
        shipped_bill_ids = set()

    all_excluded_ids = test_bill_ids | shipped_bill_ids
    if all_excluded_ids:
        order_col = next(
            (c for c in df_p.columns if str(c).strip().upper() in ('ORDER ID', 'ORDER_ID', 'BILL_ID', 'BILL ID', 'WAYBILL', 'ORDER ID/ WAYBILL')),
            'ORDER ID' if 'ORDER ID' in df_p.columns else None
        )
        if order_col:
            clean_ids = df_p[order_col].astype(str).str.strip().str.upper().str.replace(r'\.0$', '', regex=True)
            before_cnt = len(df_p)
            df_p = df_p[~clean_ids.isin(all_excluded_ids)].copy()
            purged_cnt = before_cnt - len(df_p)
            if purged_cnt > 0:
                print(f"[TOTAL_PENDING] Purged {purged_cnt} test & confirmed shipped bills upfront (like test bills)")

    # Deduplicate by ORDER ID (keep latest record)
    if 'ORDER ID' in df_p.columns:
        df_p = df_p.drop_duplicates(subset=['ORDER ID'], keep='last').copy()

    # Exclude central MEGA/HUB/DVC sorting facilities (not branch delivery points)
    is_hub = df_p['CURRENT POST OFFICE'].astype(str).str.contains('MEGA|HUB|DVC', case=False, na=False)
    df_branch = df_p[~is_hub].copy()

    # ── PURGE DELIVERED / SHIPPED BILLS (LIVE TRACKING CROSS-CHECK ON BRANCH PENDING) ───
    try:
        from shipped_filter import filter_shipped_bills_from_df
        df_branch, removed_shipped = filter_shipped_bills_from_df(df_branch, verify_live=True)
    except Exception as e_shipped:
        print(f"[TOTAL_PENDING] Warning: Live shipped verification error: {e_shipped}")

    df_branch['Branch'] = df_branch['CURRENT POST OFFICE'].apply(get_branch_code)
    df_branch['Facility_Type'] = df_branch['CURRENT POST OFFICE'].apply(get_facility_type)

    # ── LOGISTICS ACTION SPLIT: Current Post vs Delivery Post (per boss @Tian_xin_yue) ───
    # 1/ Current post == deliver post => Need deliver
    # 2/ Current post != deliver post => Need transport
    cur_clean = df_branch['CURRENT POST OFFICE'].astype(str).str.strip().str.upper()
    deliv_col_name = next((c for c in df_branch.columns if 'DELIVERY POST' in str(c).upper()), 'DELIVERY POST OFFICE')
    deliv_clean = df_branch[deliv_col_name].astype(str).str.strip().str.upper() if deliv_col_name in df_branch.columns else cur_clean

    df_branch['Need_Deliver'] = (cur_clean == deliv_clean)
    df_branch['Need_Transport'] = ~df_branch['Need_Deliver']
    df_branch['Action_Type'] = df_branch['Need_Deliver'].map({True: 'Need deliver', False: 'Need transport'})

    # Calculate timestamps and actual age
    now = datetime.now()

    def _calc_actual_ts(row):
        ts = None
        for c in TS_COLS_PRIORITY:
            if c in row and pd.notna(row[c]):
                val = str(row[c]).strip()
                if val and val.lower() != 'nan':
                    dt = pd.to_datetime(val, dayfirst=True, format='mixed', errors='coerce')
                    if pd.notna(dt):
                        ts = dt
                        break
        if ts is None:
            actual_hours = 0.0
            ts_str = ''
        else:
            actual_hours = max(0.0, (now - ts).total_seconds() / 3600.0)
            ts_str = ts.strftime('%d/%m/%Y %H:%M')
        return pd.Series([ts_str, round(actual_hours, 1)])

    ts_df = df_branch.apply(_calc_actual_ts, axis=1)
    df_branch['History_Timestamp'] = ts_df[0]
    df_branch['Actual_Hours'] = ts_df[1]

    # NO MORE GRACE PERIOD OR HISTORY CHECKING - SIMPLIFIED LIKE /PENALTY
    def _calc_aging_bucket(row):
        actual_hours = row['Actual_Hours']
        
        # NO GRACE ADJUSTMENTS - USE ACTUAL HOURS DIRECTLY
        adjusted_hours = actual_hours

        # Bucket classification based on ACTUAL hours (no grace deductions)
        if adjusted_hours >= 720.0:
            bucket = 'EXCLUDED_OVER_30'
        elif adjusted_hours < 24.0:
            bucket = '0 Days'
        elif adjusted_hours < 48.0:
            bucket = '1 Day'
        elif adjusted_hours < 72.0:
            bucket = '2 Days'
        elif adjusted_hours < 96.0:
            bucket = '3 Days'
        elif adjusted_hours < 120.0:
            bucket = '4 Days'
        elif adjusted_hours < 168.0:
            bucket = '5 Days'
        else:
            bucket = '> 7 Days'

        return pd.Series([round(adjusted_hours, 1), bucket])  # Only return hours and bucket

    aging_df = df_branch.apply(_calc_aging_bucket, axis=1)
    df_branch['Adjusted_Hours'] = aging_df[0]
    df_branch['Age_Bucket'] = aging_df[1]
    # REMOVED: Grace_Note column completely - no longer needed

    # Exclude orders >= 30 days (older than 30 days / not in this month)
    df_branch = df_branch[df_branch['Age_Bucket'] != 'EXCLUDED_OVER_30'].copy()

    # Parse COD amount per order
    cod_col = next((c for c in df_branch.columns if 'COD' in str(c).upper()), None)
    if cod_col:
        df_branch['COD_NUM'] = pd.to_numeric(df_branch[cod_col], errors='coerce').fillna(0.0)
    else:
        df_branch['COD_NUM'] = 0.0

    # Aggregate by branch
    unique_branches = sorted(df_branch['Branch'].unique(), key=lambda x: (0 if x == 'PNP' else 1, x))
    rows = []

    for b in unique_branches:
        sub = df_branch[df_branch['Branch'] == b]
        sp_cnt = int((sub['Facility_Type'] == 'Servicepoint').sum())
        sr_cnt = int((sub['Facility_Type'] == 'Showroom').sum())
        ag_cnt = int((sub['Facility_Type'] == 'Agent').sum())
        tot = len(sub)
        day_counts = {d: int((sub['Age_Bucket'] == d).sum()) for d in DAY_COLS}

        # Reconciliation check
        assert sp_cnt + sr_cnt + ag_cnt == tot, f"Facility mismatch for {b}"
        assert sum(day_counts.values()) == tot, f"Age sum mismatch for {b}"

        delay_cnt = sum(day_counts[d] for d in DELAY_COLS)
        cod_amt = round(float(sub['COD_NUM'].sum()), 2)

        nd_cnt = int(sub['Need_Deliver'].sum())
        nt_cnt = int(sub['Need_Transport'].sum())
        assert nd_cnt + nt_cnt == tot, f"Action mismatch for {b}"

        row_data = {
            'Branch': b,
            'Servicepoint': sp_cnt,
            'Showroom': sr_cnt,
            'Agent': ag_cnt
        }
        row_data.update(day_counts)
        row_data[DELAY_COL_NAME] = delay_cnt
        row_data['COD ($)'] = cod_amt
        row_data['Need Deliver'] = nd_cnt
        row_data['Need Transport'] = nt_cnt
        row_data['Total'] = tot
        rows.append(row_data)

    summary_df = pd.DataFrame(rows)

    # Grand total calculation
    grand_total = {
        'Branch': 'TOTAL',
        'Servicepoint': int(summary_df['Servicepoint'].sum()) if not summary_df.empty else 0,
        'Showroom': int(summary_df['Showroom'].sum()) if not summary_df.empty else 0,
        'Agent': int(summary_df['Agent'].sum()) if not summary_df.empty else 0,
    }
    for d in DAY_COLS:
        grand_total[d] = int(summary_df[d].sum()) if not summary_df.empty else 0
    grand_total[DELAY_COL_NAME] = int(summary_df[DELAY_COL_NAME].sum()) if not summary_df.empty else 0
    grand_total['COD ($)'] = round(float(summary_df['COD ($)'].sum()), 2) if not summary_df.empty else 0.0
    grand_total['Need Deliver'] = int(summary_df['Need Deliver'].sum()) if not summary_df.empty else 0
    grand_total['Need Transport'] = int(summary_df['Need Transport'].sum()) if not summary_df.empty else 0
    grand_total['Total'] = int(summary_df['Total'].sum()) if not summary_df.empty else 0

    # Backwards compatibility aliases
    grand_total['P'] = grand_total['Servicepoint']
    grand_total['S'] = grand_total['Showroom']
    grand_total['A'] = grand_total['Agent']

    return summary_df, grand_total, df_branch


def export_total_pending_excel(summary_df, grand_total, df_detail, out_xlsx_path):
    """
    Exports clean two-sheet Excel file matching the CEO Executive Teal Table style
    with color-coded numbers:
    1. 'Total Pending Summary' — executive pivot table with teal header, color-coded numbers, and light-teal total
    2. 'Pending Details' — audit drill-down records
    """
    wb = openpyxl.Workbook()
    ws_sum = wb.active
    ws_sum.title = "Total Pending Summary"
    ws_sum.views.sheetView[0].showGridLines = True

    # Official CEO Colors & Fonts
    header_fill = PatternFill(start_color="2E8B8B", end_color="2E8B8B", fill_type="solid")  # Teal
    total_fill  = PatternFill(start_color="B8E6E6", end_color="B8E6E6", fill_type="solid")  # Light Teal
    white_fill  = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")  # Pure White

    title_font  = Font(name="Arial", size=12, bold=True, color="FFFFFF")
    header_font = Font(name="Arial", size=10, bold=True, color="FFFFFF")
    branch_font = Font(name="Arial", size=10, bold=True, color="000000")
    data_font   = Font(name="Arial", size=10, color="000000")

    # Dynamic Number Colors for Data Rows
    fn_zero      = Font(name="Arial", size=10, color="94A3B8")              # Dim Gray for zeros
    fn_facility  = Font(name="Arial", size=10, color="002060")              # Dark Blue (Servicepoint, Showroom, Agent)
    fn_normal    = Font(name="Arial", size=10, color="0F172A")              # Normal Dark (0 Days, 1 Day, 2 Days)
    fn_alert     = Font(name="Arial", size=10, bold=True, color="FF0000")   # Red (3 Days, 4 Days, 5 Days, > 7 Days)
    fn_deliver   = Font(name="Arial", size=10, bold=True, color="047857")   # Emerald Green (Need Deliver)
    fn_transport = Font(name="Arial", size=10, bold=True, color="D97706")   # Amber (Need Transport)
    fn_col_total = Font(name="Arial", size=10.5, bold=True, color="002060") # Bold Dark Blue (Total column)

    # Grand Total Row Fonts
    fn_tot_label    = Font(name="Arial", size=11, bold=True, color="000000")
    fn_tot_facility = Font(name="Arial", size=11, bold=True, color="002060") # Dark Blue
    fn_tot_normal   = Font(name="Arial", size=11, bold=True, color="0F172A") # Dark
    fn_tot_alert    = Font(name="Arial", size=11, bold=True, color="FF0000") # Red
    fn_tot_deliver  = Font(name="Arial", size=11, bold=True, color="047857")
    fn_tot_transport= Font(name="Arial", size=11, bold=True, color="D97706")
    fn_tot_total    = Font(name="Arial", size=11.5, bold=True, color="002060") # Bold Dark Blue

    border_style = Border(
        left=Side(style="thin", color="CCCCCC"),
        right=Side(style="thin", color="CCCCCC"),
        top=Side(style="thin", color="CCCCCC"),
        bottom=Side(style="thin", color="CCCCCC")
    )

    columns = ['Branch'] + FACILITY_COLS + DAY_COLS + [DELAY_COL_NAME, 'COD ($)', 'Need Deliver', 'Need Transport', 'Total']

    # 1. Title Banner
    ws_sum.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(columns))
    now_str = datetime.now().strftime("%d/%m/%Y %H:%M")
    title_cell = ws_sum.cell(row=1, column=1, value=f"TOTAL PENDING REPORT (ALL BRANCHES)  —  {now_str}")
    title_cell.font = title_font
    title_cell.fill = header_fill
    title_cell.alignment = Alignment(horizontal="center", vertical="center")
    title_cell.border = border_style
    ws_sum.row_dimensions[1].height = 28

    # 2. Table Headers
    ws_sum.row_dimensions[2].height = 28
    for c_idx, col_name in enumerate(columns, 1):
        cell = ws_sum.cell(row=2, column=c_idx, value=col_name)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = border_style

    # 3. Data Rows with Colored Numbers
    start_data_row = 3
    for r_idx, row in summary_df.iterrows():
        cur_r = start_data_row + r_idx
        ws_sum.row_dimensions[cur_r].height = 22
        for c_idx, col_name in enumerate(columns, 1):
            val = row[col_name]
            cell = ws_sum.cell(row=cur_r, column=c_idx, value=val)
            cell.fill = white_fill
            cell.border = border_style
            if col_name == 'Branch':
                cell.font = branch_font
                cell.alignment = Alignment(horizontal="center", vertical="center")
            elif col_name == 'COD ($)':
                cell.alignment = Alignment(horizontal="right", vertical="center")
                cell.number_format = '$#,##0.00'
                if val == 0:
                    cell.font = fn_zero
                else:
                    cell.font = fn_normal
            elif col_name == 'Need Deliver':
                cell.alignment = Alignment(horizontal="right", vertical="center")
                cell.number_format = '#,##0'
                cell.font = fn_deliver if val > 0 else fn_zero
            elif col_name == 'Need Transport':
                cell.alignment = Alignment(horizontal="right", vertical="center")
                cell.number_format = '#,##0'
                cell.font = fn_transport if val > 0 else fn_zero
            else:
                cell.alignment = Alignment(horizontal="right", vertical="center")
                cell.number_format = '#,##0'
                if val == 0:
                    cell.font = fn_zero
                elif col_name in ('Servicepoint', 'Showroom', 'Agent'):
                    cell.font = fn_facility
                elif col_name in ('0 Days', '1 Day', '2 Days'):
                    cell.font = fn_normal
                elif col_name in ('3 Days', '4 Days', '5 Days', '> 7 Days', DELAY_COL_NAME):
                    cell.font = fn_alert
                elif col_name == 'Total':
                    cell.font = fn_col_total

    # 4. Total Row with Excel SUM formulas and Colored Numbers
    end_data_row = start_data_row + len(summary_df) - 1
    tot_r = end_data_row + 1
    ws_sum.row_dimensions[tot_r].height = 26

    tot_label = ws_sum.cell(row=tot_r, column=1, value="TOTAL")
    tot_label.font = fn_tot_label
    tot_label.fill = total_fill
    tot_label.border = border_style
    tot_label.alignment = Alignment(horizontal="left", vertical="center")

    for c_idx in range(2, len(columns) + 1):
        col_letter = get_column_letter(c_idx)
        col_name = columns[c_idx - 1]
        tot_val = grand_total.get(col_name, 0)
        cell = ws_sum.cell(row=tot_r, column=c_idx, value=tot_val)
        cell.fill = total_fill
        cell.border = border_style
        cell.alignment = Alignment(horizontal="right", vertical="center")

        if col_name == 'COD ($)':
            cell.font = fn_tot_normal
            cell.number_format = '$#,##0.00'
        elif col_name == 'Need Deliver':
            cell.font = fn_tot_deliver
            cell.number_format = '#,##0'
        elif col_name == 'Need Transport':
            cell.font = fn_tot_transport
            cell.number_format = '#,##0'
        elif col_name in ('Servicepoint', 'Showroom', 'Agent'):
            cell.font = fn_tot_facility
            cell.number_format = '#,##0'
        elif col_name in ('0 Days', '1 Day', '2 Days'):
            cell.font = fn_tot_normal
            cell.number_format = '#,##0'
        elif col_name in ('3 Days', '4 Days', '5 Days', '> 7 Days', DELAY_COL_NAME):
            cell.font = fn_tot_alert
            cell.number_format = '#,##0'
        elif col_name == 'Total':
            cell.font = fn_tot_total
            cell.number_format = '#,##0'

    # Set Column Widths for Summary
    for c_idx, col_name in enumerate(columns, 1):
        col_let = get_column_letter(c_idx)
        if col_name in ('Branch', 'Total'):
            ws_sum.column_dimensions[col_let].width = 13.0
        elif col_name == 'COD ($)':
            ws_sum.column_dimensions[col_let].width = 14.5
        elif col_name in ('Servicepoint', 'Showroom', 'Agent'):
            ws_sum.column_dimensions[col_let].width = 14.0
        elif col_name in ('Need Deliver', 'Need Transport'):
            ws_sum.column_dimensions[col_let].width = 15.0
        elif col_name == DELAY_COL_NAME:
            ws_sum.column_dimensions[col_let].width = 18.0
        elif col_name == '> 7 Days':
            ws_sum.column_dimensions[col_let].width = 13.0
        else:
            ws_sum.column_dimensions[col_let].width = 10.0

    # =========================================================================
    # 10 Individual Sheets for the Top 10 Pending Branches (ONLY >= 3D Bills)
    # =========================================================================
    coral_fill = PatternFill(start_color="C00000", end_color="C00000", fill_type="solid")
    coral_banner_fill = PatternFill(start_color="B22222", end_color="B22222", fill_type="solid")
    gray_alt   = PatternFill(start_color="F8FAFC", end_color="F8FAFC", fill_type="solid")
    kpi_hdr_fill = PatternFill(start_color="334155", end_color="334155", fill_type="solid")
    kpi_val_fill = PatternFill(start_color="F1F5F9", end_color="F1F5F9", fill_type="solid")

    detail_cols_meta = [
        ('NO', 8),
        ('Branch', 12),
        ('Facility Type', 14),
        ('Current Post Office', 20),
        ('Delivery Post Office', 20),
        ('Action Type', 16),
        ('Order ID', 18),
        ('Current Status', 25),
        ('Status Code', 12),
        ('History Timestamp', 20),
        ('Actual Hours', 14),
        ('Adjusted Hours', 15),
        ('Age Bucket', 14),
        ('COD (USD)', 14),
        ('Sender', 25),
        ('Receiver', 25),
        ('Phone', 16),
    ]

    top10_df = pd.DataFrame()
    if summary_df is not None and not summary_df.empty and 'Branch' in summary_df.columns and 'Total' in summary_df.columns:
        valid_sum = summary_df[summary_df['Branch'] != 'TOTAL'].copy()
        top10_df = valid_sum.sort_values(by='Total', ascending=False).head(10).copy()

    col_age = 'Age_Bucket' if (df_detail is not None and 'Age_Bucket' in df_detail.columns) else 'Age Bucket'
    col_hrs = 'Actual_Hours' if (df_detail is not None and 'Actual_Hours' in df_detail.columns) else 'Actual Hours'

    for idx, (_, r) in enumerate(top10_df.iterrows(), 1):
        b_code = str(r['Branch']).strip()
        tot = int(r.get('Total', 0))
        b_nd = int(r.get('Need Deliver', 0))
        b_nt = int(r.get('Need Transport', 0))
        delay_col_key = DELAY_COL_NAME if DELAY_COL_NAME in r else 'Total >= 3 Days'
        delay = int(r.get(delay_col_key, 0))
        pct = (delay / tot * 100.0) if tot > 0 else 0.0
        d3 = int(r.get('3 Days', 0))
        d4 = int(r.get('4 Days', 0))
        d5 = int(r.get('5 Days', 0))
        d7 = int(r.get('> 7 Days', 0))

        # Filter bills strictly to this branch and delay >= 3 Days
        branch_delayed = pd.DataFrame()
        if df_detail is not None and not df_detail.empty and 'Branch' in df_detail.columns:
            branch_delayed = df_detail[
                (df_detail['Branch'] == b_code) &
                (df_detail[col_age].isin(DELAY_COLS))
            ].copy()
            if col_hrs in branch_delayed.columns:
                branch_delayed = branch_delayed.sort_values(by=col_hrs, ascending=False)

        clean_b = re.sub(r'[\\/*?:\[\]]', '', b_code).strip()
        sheet_title = f"{idx}. {clean_b} ({len(branch_delayed)})"[:31]

        ws_b = wb.create_sheet(title=sheet_title)
        ws_b.views.sheetView[0].showGridLines = True

        # Row 1: Title Banner
        ws_b.merge_cells("A1:Q1")
        b_banner = ws_b.cell(1, 1, value=f"🚨 TOP {idx}: {b_code} — DELAYED BILLS AUDIT (>= 3 DAYS) — {now_str}")
        b_banner.font = Font(name="Arial", size=11, bold=True, color="FFFFFF")
        b_banner.fill = coral_banner_fill
        b_banner.alignment = Alignment(horizontal="center", vertical="center")
        ws_b.row_dimensions[1].height = 28

        # Row 2 & 3: Branch KPI Summary Table
        kpi_cols = ["Rank", "Branch", "Need Deliver", "Need Transport", "Total Pending", "Delayed (>= 3D)", "% Delayed", "3 Days", "4 Days", "5 Days", "> 7 Days", "COD ($)"]
        ws_b.row_dimensions[2].height = 22
        for c_i, h in enumerate(kpi_cols, 1):
            c = ws_b.cell(2, c_i, value=h)
            c.font = Font(name="Arial", size=9.5, bold=True, color="FFFFFF")
            c.fill = kpi_hdr_fill
            c.alignment = Alignment(horizontal="center", vertical="center")
            c.border = border_style

        b_delayed_cod = round(float(branch_delayed['COD_NUM'].sum()), 2) if ('COD_NUM' in branch_delayed.columns and not branch_delayed.empty) else 0.0
        kpi_vals = [idx, b_code, b_nd, b_nt, tot, len(branch_delayed), f"{pct:.1f}%", d3, d4, d5, d7, b_delayed_cod]
        ws_b.row_dimensions[3].height = 22
        for c_i, val in enumerate(kpi_vals, 1):
            c = ws_b.cell(3, c_i, value=val)
            c.fill = kpi_val_fill
            c.border = border_style
            if c_i in (1, 2):
                c.font = branch_font
                c.alignment = Alignment(horizontal="center", vertical="center")
            elif c_i == 3:
                c.font = fn_deliver if (isinstance(val, int) and val > 0) else fn_zero
                c.alignment = Alignment(horizontal="right", vertical="center")
                if isinstance(val, int): c.number_format = '#,##0'
            elif c_i == 4:
                c.font = fn_transport if (isinstance(val, int) and val > 0) else fn_zero
                c.alignment = Alignment(horizontal="right", vertical="center")
                if isinstance(val, int): c.number_format = '#,##0'
            elif c_i == 12:
                c.font = fn_normal
                c.alignment = Alignment(horizontal="right", vertical="center")
                c.number_format = '$#,##0.00'
            elif c_i in (6, 8, 9, 10, 11):
                c.font = fn_alert if (isinstance(val, int) and val > 0) else fn_normal
                c.alignment = Alignment(horizontal="right", vertical="center")
                if isinstance(val, int):
                    c.number_format = '#,##0'
            else:
                c.font = fn_normal
                c.alignment = Alignment(horizontal="right", vertical="center")
                if isinstance(val, int):
                    c.number_format = '#,##0'

        # Row 4: Blank separator row
        ws_b.row_dimensions[4].height = 12

        # Row 5: Detail Table Headers
        ws_b.row_dimensions[5].height = 25
        for c_idx, (col_name, width) in enumerate(detail_cols_meta, 1):
            c = ws_b.cell(row=5, column=c_idx, value=col_name)
            c.font = header_font
            c.fill = header_fill
            c.border = border_style
            c.alignment = Alignment(horizontal="center", vertical="center")
            ws_b.column_dimensions[get_column_letter(c_idx)].width = width

        # Row 6+: Detailed Delayed Orders
        cur_r = 6
        if branch_delayed.empty:
            ws_b.row_dimensions[cur_r].height = 25
            c_empty = ws_b.cell(row=cur_r, column=1, value="🎉 No delayed bills (>= 3 days) for this branch!")
            c_empty.font = Font(name="Arial", size=10, italic=True, bold=True, color="008000")
            c_empty.alignment = Alignment(horizontal="left", vertical="center")
            ws_b.merge_cells(start_row=cur_r, start_column=1, end_row=cur_r, end_column=len(detail_cols_meta))
            cur_r += 1
        else:
            for d_idx, (_, item) in enumerate(branch_delayed.iterrows(), 1):
                ws_b.row_dimensions[cur_r].height = 20
                is_alt = (d_idx % 2 == 0)
                bg = gray_alt if is_alt else white_fill

                cur_po = str(item.get('CURRENT POST OFFICE', item.get('Current Post Office', ''))).strip()
                deliv_po = str(item.get('DELIVERY POST OFFICE', item.get('Delivery Post Office', ''))).strip()
                action_val = str(item.get('Action_Type', item.get('Action Type', 'Need deliver' if cur_po.upper() == deliv_po.upper() else 'Need transport')))

                ws_b.cell(row=cur_r, column=1, value=d_idx).alignment = Alignment(horizontal='center')
                ws_b.cell(row=cur_r, column=2, value=str(item.get('Branch', ''))).alignment = Alignment(horizontal='center')
                ws_b.cell(row=cur_r, column=3, value=str(item.get('Facility_Type', item.get('Facility Type', '')))).alignment = Alignment(horizontal='center')
                ws_b.cell(row=cur_r, column=4, value=cur_po).alignment = Alignment(horizontal='center')
                ws_b.cell(row=cur_r, column=5, value=deliv_po).alignment = Alignment(horizontal='center')
                c_act = ws_b.cell(row=cur_r, column=6, value=action_val)
                c_act.alignment = Alignment(horizontal='center')
                c_act.font = Font(name="Arial", size=10, bold=True, color="047857" if action_val == 'Need deliver' else "D97706")
                ws_b.cell(row=cur_r, column=7, value=str(item.get('ORDER ID', item.get('Order ID', '')))).alignment = Alignment(horizontal='center')
                ws_b.cell(row=cur_r, column=8, value=str(item.get('CURRENT STATUS', item.get('Current Status', '')))).alignment = Alignment(horizontal='left')
                ws_b.cell(row=cur_r, column=9, value=str(item.get('STATUS_CODE', item.get('Status Code', '')))).alignment = Alignment(horizontal='center')
                ws_b.cell(row=cur_r, column=10, value=str(item.get('History_Timestamp', item.get('History Timestamp', '')))).alignment = Alignment(horizontal='center')
                ws_b.cell(row=cur_r, column=11, value=item.get('Actual_Hours', item.get('Actual Hours', 0.0))).alignment = Alignment(horizontal='right')
                ws_b.cell(row=cur_r, column=12, value=item.get('Adjusted_Hours', item.get('Adjusted Hours', 0.0))).alignment = Alignment(horizontal='right')
                ws_b.cell(row=cur_r, column=13, value=str(item.get('Age_Bucket', item.get('Age Bucket', '')))).alignment = Alignment(horizontal='center')
                c_cod = ws_b.cell(row=cur_r, column=14, value=float(item.get('COD_NUM', item.get('COD (USD)', 0.0))))
                c_cod.alignment = Alignment(horizontal='right')
                c_cod.number_format = '$#,##0.00'
                ws_b.cell(row=cur_r, column=15, value=str(item.get('SENDER', item.get('Sender', '')))).alignment = Alignment(horizontal='left')
                ws_b.cell(row=cur_r, column=16, value=str(item.get('RECEIVER', item.get('Receiver', '')))).alignment = Alignment(horizontal='left')
                ws_b.cell(row=cur_r, column=17, value=str(item.get('Phone', item.get('PHONE', '')))).alignment = Alignment(horizontal='center')

                for col_i in range(1, len(detail_cols_meta) + 1):
                    cell = ws_b.cell(row=cur_r, column=col_i)
                    cell.border = border_style
                    cell.fill = bg
                    if col_i in (2, 7):
                        cell.font = branch_font
                    elif col_i == 6:
                        cell.font = Font(name="Arial", size=10, bold=True, color="047857" if action_val == 'Need deliver' else "D97706")
                    elif col_i == 13:
                        cell.font = fn_alert
                    else:
                        cell.font = data_font

                    if col_i in (11, 12) and isinstance(cell.value, (int, float)):
                        cell.number_format = '0.0'
                    elif col_i == 14:
                        cell.number_format = '$#,##0.00'
                    elif col_i == 17:
                        cell.number_format = '@'

                cur_r += 1

            ws_b.auto_filter.ref = f"A5:Q{cur_r - 1}"
            ws_b.freeze_panes = "A6"

    # =========================================================================
    # Sheet 3: Pending Details (All Branches & All Ages)
    # =========================================================================
    ws_det = wb.create_sheet(title="Pending Details")
    ws_det.views.sheetView[0].showGridLines = True

    detail_cols = [
        ('NO', 8),
        ('Branch', 12),
        ('Facility Type', 14),
        ('Current Post Office', 20),
        ('Delivery Post Office', 20),
        ('Action Type', 16),
        ('Order ID', 18),
        ('Current Status', 25),
        ('Status Code', 12),
        ('History Timestamp', 20),
        ('Actual Hours', 14),
        ('Adjusted Hours', 15),
        ('Age Bucket', 14),
        ('COD (USD)', 14),
        ('Sender', 25),
        ('Receiver', 25),
        ('Phone', 16),
    ]

    ws_det.row_dimensions[1].height = 25
    for c_idx, (col_name, width) in enumerate(detail_cols, 1):
        cell = ws_det.cell(row=1, column=c_idx, value=col_name)
        cell.font = header_font
        cell.fill = header_fill
        cell.border = border_style
        cell.alignment = Alignment(horizontal="center", vertical="center")
        ws_det.column_dimensions[get_column_letter(c_idx)].width = width

    # Populate details
    det_row = 2
    for _, item in df_detail.iterrows():
        ws_det.row_dimensions[det_row].height = 20
        cur_po = str(item.get('CURRENT POST OFFICE', item.get('Current Post Office', ''))).strip()
        deliv_po = str(item.get('DELIVERY POST OFFICE', item.get('Delivery Post Office', ''))).strip()
        action_val = str(item.get('Action_Type', item.get('Action Type', 'Need deliver' if cur_po.upper() == deliv_po.upper() else 'Need transport')))

        ws_det.cell(row=det_row, column=1, value=det_row - 1).alignment = Alignment(horizontal='center')
        ws_det.cell(row=det_row, column=2, value=str(item.get('Branch', ''))).alignment = Alignment(horizontal='center')
        ws_det.cell(row=det_row, column=3, value=str(item.get('Facility_Type', item.get('Facility Type', '')))).alignment = Alignment(horizontal='center')
        ws_det.cell(row=det_row, column=4, value=cur_po).alignment = Alignment(horizontal='center')
        ws_det.cell(row=det_row, column=5, value=deliv_po).alignment = Alignment(horizontal='center')
        c_act = ws_det.cell(row=det_row, column=6, value=action_val)
        c_act.alignment = Alignment(horizontal='center')
        c_act.font = Font(name="Arial", size=10, bold=True, color="047857" if action_val == 'Need deliver' else "D97706")
        ws_det.cell(row=det_row, column=7, value=str(item.get('ORDER ID', item.get('Order ID', '')))).alignment = Alignment(horizontal='center')
        ws_det.cell(row=det_row, column=8, value=str(item.get('CURRENT STATUS', item.get('Current Status', '')))).alignment = Alignment(horizontal='left')
        ws_det.cell(row=det_row, column=9, value=str(item.get('STATUS_CODE', item.get('Status Code', '')))).alignment = Alignment(horizontal='center')
        ws_det.cell(row=det_row, column=10, value=str(item.get('History_Timestamp', item.get('History Timestamp', '')))).alignment = Alignment(horizontal='center')
        ws_det.cell(row=det_row, column=11, value=item.get('Actual_Hours', item.get('Actual Hours', 0.0))).alignment = Alignment(horizontal='right')
        ws_det.cell(row=det_row, column=12, value=item.get('Adjusted_Hours', item.get('Adjusted Hours', 0.0))).alignment = Alignment(horizontal='right')
        ws_det.cell(row=det_row, column=13, value=str(item.get('Age_Bucket', item.get('Age Bucket', '')))).alignment = Alignment(horizontal='center')
        c_cod_det = ws_det.cell(row=det_row, column=14, value=float(item.get('COD_NUM', item.get('COD (USD)', 0.0))))
        c_cod_det.alignment = Alignment(horizontal='right')
        c_cod_det.number_format = '$#,##0.00'
        ws_det.cell(row=det_row, column=15, value=str(item.get('SENDER', item.get('Sender', '')))).alignment = Alignment(horizontal='left')
        ws_det.cell(row=det_row, column=16, value=str(item.get('RECEIVER', item.get('Receiver', '')))).alignment = Alignment(horizontal='left')
        ws_det.cell(row=det_row, column=17, value=str(item.get('Phone', item.get('PHONE', '')))).alignment = Alignment(horizontal='center')

        data_font = Font(name="Arial", size=10, color="000000")
        for col_i in range(1, len(detail_cols) + 1):
            ws_det.cell(row=det_row, column=col_i).font = data_font
            ws_det.cell(row=det_row, column=col_i).border = border_style
            if col_i == 6:
                ws_det.cell(row=det_row, column=col_i).font = Font(name="Arial", size=10, bold=True, color="047857" if action_val == 'Need deliver' else "D97706")
            elif col_i in (11, 12) and isinstance(ws_det.cell(row=det_row, column=col_i).value, (int, float)):
                ws_det.cell(row=det_row, column=col_i).number_format = '0.0'
            elif col_i == 14:
                ws_det.cell(row=det_row, column=col_i).number_format = '$#,##0.00'
            elif col_i == 17:
                ws_det.cell(row=det_row, column=col_i).number_format = '@'
        det_row += 1

    ws_det.auto_filter.ref = f"A1:Q{det_row - 1}"

    # Helper to add dedicated detail tabs for Need Deliver and Need Transport
    def _create_action_detail_sheet(target_sheet, sub_df, banner_color):
        target_sheet.views.sheetView[0].showGridLines = True
        h_fill = PatternFill(start_color=banner_color, end_color=banner_color, fill_type="solid")
        target_sheet.row_dimensions[1].height = 25
        for c_idx, (col_name, width) in enumerate(detail_cols, 1):
            cell = target_sheet.cell(row=1, column=c_idx, value=col_name)
            cell.font = header_font
            cell.fill = h_fill
            cell.border = border_style
            cell.alignment = Alignment(horizontal="center", vertical="center")
            target_sheet.column_dimensions[get_column_letter(c_idx)].width = width

        cur_row = 2
        for _, item in sub_df.iterrows():
            target_sheet.row_dimensions[cur_row].height = 20
            cur_po = str(item.get('CURRENT POST OFFICE', item.get('Current Post Office', ''))).strip()
            deliv_po = str(item.get('DELIVERY POST OFFICE', item.get('Delivery Post Office', ''))).strip()
            action_val = str(item.get('Action_Type', item.get('Action Type', 'Need deliver'))).strip()

            target_sheet.cell(row=cur_row, column=1, value=cur_row - 1).alignment = Alignment(horizontal='center')
            target_sheet.cell(row=cur_row, column=2, value=str(item.get('Branch', item.get('BRANCH', '')))).alignment = Alignment(horizontal='center')
            target_sheet.cell(row=cur_row, column=3, value=str(item.get('Facility_Type', item.get('Facility Type', '')))).alignment = Alignment(horizontal='center')
            target_sheet.cell(row=cur_row, column=4, value=cur_po).alignment = Alignment(horizontal='left')
            target_sheet.cell(row=cur_row, column=5, value=deliv_po).alignment = Alignment(horizontal='left')
            c_act = target_sheet.cell(row=cur_row, column=6, value=action_val)
            c_act.alignment = Alignment(horizontal='center')
            c_act.font = Font(name="Arial", size=10, bold=True, color="047857" if action_val == 'Need deliver' else "D97706")
            target_sheet.cell(row=cur_row, column=7, value=str(item.get('ORDER ID', item.get('Order ID', '')))).alignment = Alignment(horizontal='center')
            target_sheet.cell(row=cur_row, column=8, value=str(item.get('CURRENT STATUS', item.get('Current Status', '')))).alignment = Alignment(horizontal='left')
            target_sheet.cell(row=cur_row, column=9, value=str(item.get('STATUS_CODE', item.get('Status Code', '')))).alignment = Alignment(horizontal='center')
            target_sheet.cell(row=cur_row, column=10, value=str(item.get('History_Timestamp', item.get('History Timestamp', '')))).alignment = Alignment(horizontal='center')
            target_sheet.cell(row=cur_row, column=11, value=item.get('Actual_Hours', item.get('Actual Hours', 0.0))).alignment = Alignment(horizontal='right')
            target_sheet.cell(row=cur_row, column=12, value=item.get('Adjusted_Hours', item.get('Adjusted Hours', 0.0))).alignment = Alignment(horizontal='right')
            target_sheet.cell(row=cur_row, column=13, value=str(item.get('Age_Bucket', item.get('Age Bucket', '')))).alignment = Alignment(horizontal='center')
            c_cod_sub = target_sheet.cell(row=cur_row, column=14, value=float(item.get('COD_NUM', item.get('COD (USD)', 0.0))))
            c_cod_sub.alignment = Alignment(horizontal='right')
            c_cod_sub.number_format = '$#,##0.00'
            target_sheet.cell(row=cur_row, column=15, value=str(item.get('SENDER', item.get('Sender', '')))).alignment = Alignment(horizontal='left')
            target_sheet.cell(row=cur_row, column=16, value=str(item.get('RECEIVER', item.get('Receiver', '')))).alignment = Alignment(horizontal='left')
            target_sheet.cell(row=cur_row, column=17, value=str(item.get('Phone', item.get('PHONE', '')))).alignment = Alignment(horizontal='center')

            data_fnt = Font(name="Arial", size=10, color="000000")
            for col_i in range(1, len(detail_cols) + 1):
                target_sheet.cell(row=cur_row, column=col_i).font = data_fnt
                target_sheet.cell(row=cur_row, column=col_i).border = border_style
                if col_i == 6:
                    target_sheet.cell(row=cur_row, column=col_i).font = Font(name="Arial", size=10, bold=True, color="047857" if action_val == 'Need deliver' else "D97706")
                elif col_i in (11, 12) and isinstance(target_sheet.cell(row=cur_row, column=col_i).value, (int, float)):
                    target_sheet.cell(row=cur_row, column=col_i).number_format = '0.0'
                elif col_i == 14:
                    target_sheet.cell(row=cur_row, column=col_i).number_format = '$#,##0.00'
                elif col_i == 17:
                    target_sheet.cell(row=cur_row, column=col_i).number_format = '@'
            cur_row += 1

        if cur_row > 2:
            target_sheet.auto_filter.ref = f"A1:Q{cur_row - 1}"

    # Add dedicated Need Deliver & Need Transport detail sheets
    if df_detail is not None and not df_detail.empty:
        if 'Need_Deliver' in df_detail.columns:
            m_nd = df_detail['Need_Deliver'] == True
            m_nt = df_detail['Need_Transport'] == True
        elif 'Action Type' in df_detail.columns:
            m_nd = df_detail['Action Type'].astype(str).str.lower().str.contains('deliver')
            m_nt = df_detail['Action Type'].astype(str).str.lower().str.contains('transport')
        elif 'Action_Type' in df_detail.columns:
            m_nd = df_detail['Action_Type'].astype(str).str.lower().str.contains('deliver')
            m_nt = df_detail['Action_Type'].astype(str).str.lower().str.contains('transport')
        else:
            cur = df_detail.get('Current Post Office', df_detail.get('CURRENT POST OFFICE', pd.Series())).astype(str).str.strip().str.upper()
            deliv = df_detail.get('Delivery Post Office', df_detail.get('DELIVERY POST OFFICE', cur)).astype(str).str.strip().str.upper()
            m_nd = (cur == deliv)
            m_nt = (cur != deliv)

        df_nd = df_detail[m_nd].copy()
        if not df_nd.empty:
            ws_nd = wb.create_sheet(title=f"Need Deliver ({len(df_nd)})"[:31])
            _create_action_detail_sheet(ws_nd, df_nd, "047857")

        df_nt = df_detail[m_nt].copy()
        if not df_nt.empty:
            ws_nt = wb.create_sheet(title=f"Need Transport ({len(df_nt)})"[:31])
            _create_action_detail_sheet(ws_nt, df_nt, "D97706")


    wb.active = ws_sum
    wb.save(out_xlsx_path)
    return out_xlsx_path


def render_total_pending_image(summary_df, grand_total, out_png_path=None, df_detail=None, xlsx_path=None):
    """
    Renders the exact CEO Executive Table design directly to image using excel_to_image.
    Produces a crisp, high-definition spreadsheet representation.
    """
    temp_created = False
    if not xlsx_path or not os.path.exists(xlsx_path):
        temp_dir = tempfile.mkdtemp(prefix="pending_render_")
        xlsx_path = os.path.join(temp_dir, "temp_pending.xlsx")
        export_total_pending_excel(summary_df, grand_total, df_detail if df_detail is not None else pd.DataFrame(), xlsx_path)
        temp_created = True

    try:
        img_buf = excel_to_image.excel_to_image(xlsx_path)
        if out_png_path:
            with open(out_png_path, 'wb') as f:
                f.write(img_buf.getvalue())
            img_buf.seek(0)
        return img_buf
    finally:
        if temp_created and os.path.exists(xlsx_path):
            try:
                os.remove(xlsx_path)
                os.rmdir(os.path.dirname(xlsx_path))
            except Exception:
                pass


def render_split_pending_images(df_detail, target_date=None, out_dir=None):
    """
    Renders two separate dedicated executive dashboards:
    1. Need Deliver Dashboard (CURRENT POST == DELIVER POST) — Emerald Theme (#047857)
    2. Need Transport Dashboard (CURRENT POST != DELIVER POST) — Amber Theme (#D97706)
    Returns: (img_buf_deliver, img_buf_transport, gt_deliver, gt_transport)
    """
    if target_date is None:
        target_date = datetime.now()
    date_str = target_date.strftime("%d/%m/%Y %H:%M")

    if 'Need_Deliver' not in df_detail.columns:
        cur = df_detail['Current Post Office'].astype(str).str.strip().str.upper() if 'Current Post Office' in df_detail.columns else df_detail['CURRENT POST OFFICE'].astype(str).str.strip().str.upper()
        deliv_col = next((c for c in df_detail.columns if 'DELIVERY' in str(c).upper()), None)
        deliv = df_detail[deliv_col].astype(str).str.strip().str.upper() if deliv_col else cur
        df_detail = df_detail.copy()
        df_detail['Need_Deliver'] = (cur == deliv)
        df_detail['Need_Transport'] = ~df_detail['Need_Deliver']

    def _build_single_split_sheet(df_sub, title_text, out_xlsx_path, banner_color):
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = 'Summary'
        ws.views.sheetView[0].showGridLines = True

        header_fill = PatternFill(start_color=banner_color, end_color=banner_color, fill_type='solid')
        total_fill = PatternFill(start_color='B8E6E6', end_color='B8E6E6', fill_type='solid')
        white_fill = PatternFill(start_color='FFFFFF', end_color='FFFFFF', fill_type='solid')

        title_font = Font(name='Arial', size=12, bold=True, color='FFFFFF')
        header_font = Font(name='Arial', size=10, bold=True, color='FFFFFF')
        branch_font = Font(name='Arial', size=10, bold=True, color='000000')
        fn_zero = Font(name='Arial', size=10, color='94A3B8')
        fn_facility = Font(name='Arial', size=10, color='002060')
        fn_normal = Font(name='Arial', size=10, color='0F172A')
        fn_alert = Font(name='Arial', size=10, bold=True, color='FF0000')
        fn_col_total = Font(name='Arial', size=10.5, bold=True, color='002060')
        fn_tot_label = Font(name='Arial', size=11, bold=True, color='000000')
        fn_tot_total = Font(name='Arial', size=11.5, bold=True, color='002060')

        border_style = Border(
            left=Side(style='thin', color='CCCCCC'),
            right=Side(style='thin', color='CCCCCC'),
            top=Side(style='thin', color='CCCCCC'),
            bottom=Side(style='thin', color='CCCCCC')
        )

        b_col = 'Branch' if 'Branch' in df_sub.columns else 'BRANCH'
        fac_col = 'Facility Type' if 'Facility Type' in df_sub.columns else ('Facility_Type' if 'Facility_Type' in df_sub.columns else 'FACILITY TYPE')
        age_col = 'Age Bucket' if 'Age Bucket' in df_sub.columns else ('Age_Bucket' if 'Age_Bucket' in df_sub.columns else 'AGE BUCKET')
        cod_col = 'COD (USD)' if 'COD (USD)' in df_sub.columns else ('COD_NUM' if 'COD_NUM' in df_sub.columns else 'COD')

        all_branches = sorted(df_detail[b_col].dropna().unique(), key=lambda x: (0 if str(x).upper() == 'PNP' else 1, str(x)))
        rows = []
        for b in all_branches:
            sub = df_sub[df_sub[b_col] == b]
            sp_cnt = int((sub[fac_col] == 'Servicepoint').sum()) if fac_col in sub else 0
            sr_cnt = int((sub[fac_col] == 'Showroom').sum()) if fac_col in sub else 0
            ag_cnt = int((sub[fac_col] == 'Agent').sum()) if fac_col in sub else 0
            tot = len(sub)
            day_counts = {d: int((sub[age_col] == d).sum()) if age_col in sub else 0 for d in DAY_COLS}
            delay_cnt = sum(day_counts[d] for d in DELAY_COLS)
            if cod_col in sub:
                cod_amt = round(float(pd.to_numeric(sub[cod_col], errors='coerce').fillna(0.0).sum()), 2)
            else:
                cod_amt = 0.0
            r = {'Branch': b, 'Servicepoint': sp_cnt, 'Showroom': sr_cnt, 'Agent': ag_cnt}
            r.update(day_counts)
            r[DELAY_COL_NAME] = delay_cnt
            r['COD ($)'] = cod_amt
            r['Total'] = tot
            rows.append(r)

        s_df = pd.DataFrame(rows)
        gt = {
            'Branch': 'TOTAL',
            'Servicepoint': int(s_df['Servicepoint'].sum()) if 'Servicepoint' in s_df else 0,
            'Showroom': int(s_df['Showroom'].sum()) if 'Showroom' in s_df else 0,
            'Agent': int(s_df['Agent'].sum()) if 'Agent' in s_df else 0,
        }
        for d in DAY_COLS:
            gt[d] = int(s_df[d].sum()) if d in s_df else 0
        gt[DELAY_COL_NAME] = int(s_df[DELAY_COL_NAME].sum()) if DELAY_COL_NAME in s_df else 0
        gt['COD ($)'] = round(float(s_df['COD ($)'].sum()), 2) if 'COD ($)' in s_df else 0.0
        gt['Total'] = int(s_df['Total'].sum()) if 'Total' in s_df else 0

        cols = ['Branch'] + FACILITY_COLS + DAY_COLS + [DELAY_COL_NAME, 'COD ($)', 'Total']

        ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(cols))
        t_cell = ws.cell(row=1, column=1, value=title_text)
        t_cell.font = title_font
        t_cell.fill = header_fill
        t_cell.alignment = Alignment(horizontal='center', vertical='center')
        ws.row_dimensions[1].height = 28

        ws.row_dimensions[2].height = 28
        for ci, cn in enumerate(cols, 1):
            c = ws.cell(row=2, column=ci, value=cn)
            c.font = header_font
            c.fill = header_fill
            c.alignment = Alignment(horizontal='center', vertical='center')
            c.border = border_style

        start_r = 3
        for ri, row in s_df.iterrows():
            cr = start_r + ri
            ws.row_dimensions[cr].height = 22
            for ci, cn in enumerate(cols, 1):
                val = row[cn]
                cell = ws.cell(row=cr, column=ci, value=val)
                cell.fill = white_fill
                cell.border = border_style
                if cn == 'Branch':
                    cell.font = branch_font
                    cell.alignment = Alignment(horizontal='center', vertical='center')
                elif cn == 'COD ($)':
                    cell.font = fn_zero if val == 0 else fn_normal
                    cell.alignment = Alignment(horizontal='right', vertical='center')
                    cell.number_format = '$#,##0.00'
                elif cn == 'Total':
                    cell.font = fn_col_total
                    cell.alignment = Alignment(horizontal='right', vertical='center')
                    cell.number_format = '#,##0'
                elif cn in ('Servicepoint', 'Showroom', 'Agent'):
                    cell.font = fn_facility if val > 0 else fn_zero
                    cell.alignment = Alignment(horizontal='right', vertical='center')
                    cell.number_format = '#,##0'
                elif cn in ('0 Days', '1 Day', '2 Days'):
                    cell.font = fn_normal if val > 0 else fn_zero
                    cell.alignment = Alignment(horizontal='right', vertical='center')
                    cell.number_format = '#,##0'
                elif cn in ('3 Days', '4 Days', '5 Days', '> 7 Days', DELAY_COL_NAME):
                    cell.font = fn_alert if val > 0 else fn_zero
                    cell.alignment = Alignment(horizontal='right', vertical='center')
                    cell.number_format = '#,##0'

        tot_r = start_r + len(s_df)
        ws.row_dimensions[tot_r].height = 26
        lbl = ws.cell(row=tot_r, column=1, value='TOTAL')
        lbl.font = fn_tot_label
        lbl.fill = total_fill
        lbl.border = border_style
        lbl.alignment = Alignment(horizontal='left', vertical='center')
        for ci in range(2, len(cols) + 1):
            cn = cols[ci - 1]
            val = gt.get(cn, 0)
            c = ws.cell(row=tot_r, column=ci, value=val)
            c.fill = total_fill
            c.border = border_style
            c.alignment = Alignment(horizontal='right', vertical='center')
            if cn == 'COD ($)':
                c.font = fn_normal
                c.number_format = '$#,##0.00'
            elif cn == 'Total':
                c.font = fn_tot_total
                c.number_format = '#,##0'
            elif cn in ('3 Days', '4 Days', '5 Days', '> 7 Days', DELAY_COL_NAME):
                c.font = fn_alert
                c.number_format = '#,##0'
            else:
                c.font = fn_normal
                c.number_format = '#,##0'

        for ci, cn in enumerate(cols, 1):
            cl = get_column_letter(ci)
            if cn in ('Branch', 'Total'):
                ws.column_dimensions[cl].width = 13.0
            elif cn == 'COD ($)':
                ws.column_dimensions[cl].width = 14.5
            elif cn in ('Servicepoint', 'Showroom', 'Agent'):
                ws.column_dimensions[cl].width = 14.0
            elif cn == DELAY_COL_NAME:
                ws.column_dimensions[cl].width = 18.0
            elif cn == '> 7 Days':
                ws.column_dimensions[cl].width = 13.0
            else:
                ws.column_dimensions[cl].width = 10.0

        wb.save(out_xlsx_path)
        buf = excel_to_image.excel_to_image(out_xlsx_path)
        return buf, gt

    temp_dir = out_dir if (out_dir and os.path.exists(out_dir)) else tempfile.mkdtemp(prefix="split_pending_")
    clean_temp = (temp_dir != out_dir)
    try:
        deliver_xlsx = os.path.join(temp_dir, "need_deliver.xlsx")
        transport_xlsx = os.path.join(temp_dir, "need_transport.xlsx")

        buf_deliv, gt_deliv = _build_single_split_sheet(
            df_detail[df_detail['Need_Deliver']],
            f"PENDING — 1. NEED DELIVER (CURRENT POST == DELIVER POST)  —  {date_str}",
            deliver_xlsx,
            banner_color="047857"
        )

        buf_trans, gt_trans = _build_single_split_sheet(
            df_detail[df_detail['Need_Transport']],
            f"PENDING — 2. NEED TRANSPORT (CURRENT POST != DELIVER POST)  —  {date_str}",
            transport_xlsx,
            banner_color="D97706"
        )
        return buf_deliv, buf_trans, gt_deliv, gt_trans
    finally:
        if clean_temp:
            try:
                shutil.rmtree(temp_dir, ignore_errors=True)
            except Exception:
                pass



def format_pending_text_summary(summary_df, grand_total, target_date=None):
    """Formats markdown caption for scheduled or interactive Total Pending reports matching official detail view."""
    if target_date is None:
        target_date = datetime.now()
    stamp_str = target_date.strftime("%d/%m/%Y %H:%M")

    total_val = grand_total.get("Total", 0) if isinstance(grand_total, dict) else (grand_total or 0)
    nd_val = grand_total.get("Need Deliver", 0) if isinstance(grand_total, dict) else 0
    nt_val = grand_total.get("Need Transport", 0) if isinstance(grand_total, dict) else 0
    sp_val = grand_total.get("Servicepoint", 0) if isinstance(grand_total, dict) else 0
    sr_val = grand_total.get("Showroom", 0) if isinstance(grand_total, dict) else 0
    ag_val = grand_total.get("Agent", 0) if isinstance(grand_total, dict) else 0
    delay_val = grand_total.get(DELAY_COL_NAME, grand_total.get("Total >= 3 Days", 0)) if isinstance(grand_total, dict) else 0

    if delay_val == 0 and summary_df is not None and not summary_df.empty:
        if DELAY_COL_NAME in summary_df.columns:
            delay_val = int(summary_df[DELAY_COL_NAME].sum())

    cod_val = grand_total.get("COD ($)", 0.0) if isinstance(grand_total, dict) else 0.0

    lines = [
        f"TOTAL PENDING REPORT — {stamp_str}",
        f"",
        f"Total Pending: {total_val:,} | Total COD: ${cod_val:,.2f}",
        f"🚚 Need Deliver: {nd_val:,} | 🚛 Need Transport: {nt_val:,}",
        f"Servicepoint: {sp_val:,} | Showroom: {sr_val:,} | Agent: {ag_val:,}",
        f"Total Delay (>= 3 Days): {delay_val:,}",
    ]

    if summary_df is not None and not summary_df.empty and 'Branch' in summary_df.columns and 'Total' in summary_df.columns:
        lines.append("")
        lines.append("Top 10 Pending Branches:")
        top10 = summary_df.sort_values(by='Total', ascending=False).head(10)
        for idx, (_, row) in enumerate(top10.iterrows(), 1):
            b_code = str(row['Branch'])
            b_tot = int(row['Total'])
            b_delay = int(row.get(DELAY_COL_NAME, row.get('Total >= 3 Days', 0)))
            b_nd = int(row.get('Need Deliver', 0))
            b_nt = int(row.get('Need Transport', 0))
            lines.append(f"{idx}. {b_code}: {b_tot:,} (Deliver: {b_nd:,}, Transport: {b_nt:,} | >=3D: {b_delay:,})")

    return "\n".join(lines)

