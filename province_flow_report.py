# -*- coding: utf-8 -*-
"""
province_flow_report.py
Analyzes and compares RECEIVE PROVINCE (Origin) vs DELIVERY PROVINCE (Destination)
for Metfone TMS data.
Supports generating reports for October 2026 (current month), September 2026, and full period.
"""

import os
import sys
import json
import pandas as pd
from datetime import datetime
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.utils import get_column_letter

HERE = os.path.dirname(os.path.abspath(__file__))
INPUT_FILE = os.path.join(HERE, "cache", "detail_master_until_today.xlsx")

PROVINCE_NAMES = {
    "PNP": "Phnom Penh",
    "KAN": "Kandal",
    "PRE": "Prey Veng",
    "SVA": "Svay Rieng",
    "SIH": "Preah Sihanouk",
    "SPE": "Kampong Speu",
    "KOH": "Koh Kong",
    "TAK": "Takeo",
    "KAM": "Kampot",
    "BAT": "Battambang",
    "BAN": "Banteay Meanchey",
    "CHH": "Kampong Chhnang",
    "PUR": "Pursat",
    "SIE": "Siem Reap",
    "THO": "Kampong Thom",
    "ODD": "Oddar Meanchey",
    "PRH": "Preah Vihear",
    "KRA": "Kratie",
    "CHA": "Kampong Cham",
    "TBK": "Tboung Khmum",
    "MON": "Mondulkiri",
    "ROT": "Ratanakiri",
    "STU": "Stung Treng",
    "KEP": "Kep",
    "PAI": "Pailin"
}

def generate_province_report(input_xlsx, output_xlsx, target_month=None, title_label=None):
    print(f"[INFO] Reading detail file: {input_xlsx}...")
    try:
        df = pd.read_excel(input_xlsx, engine="calamine")
    except Exception:
        df = pd.read_excel(input_xlsx)
    
    df.columns = [str(c).strip().upper() for c in df.columns]

    rec_col = next((c for c in df.columns if "RECEIVE PROVINCE" in c), "RECEIVE PROVINCE")
    del_col = next((c for c in df.columns if "DELIVERY PROVINCE" in c), "DELIVERY PROVINCE")
    cur_col = next((c for c in df.columns if "CURRENT PROVINCE" in c), "CURRENT PROVINCE")
    st_col  = next((c for c in df.columns if "CURRENT STATUS" in c), "CURRENT STATUS")
    oid_col = next((c for c in df.columns if "ORDER ID" in c), "ORDER ID")
    time_col = next((c for c in df.columns if "CURRENT TIME" in c), None)
    cdate_col = next((c for c in df.columns if "CREATED DATE" in c), None)

    # Filter invalid/test rows
    df_clean = df.dropna(subset=[oid_col]).copy()
    
    # Optional month filter
    if target_month is not None:
        df_clean["_dt_created"] = pd.to_datetime(df_clean[cdate_col], dayfirst=True, errors='coerce') if cdate_col else None
        df_clean["_dt_cur"] = pd.to_datetime(df_clean[time_col], dayfirst=True, errors='coerce') if time_col else None
        
        m_created = df_clean["_dt_created"].dt.month == target_month if cdate_col else False
        m_cur = df_clean["_dt_cur"].dt.month == target_month if time_col else False
        df_clean = df_clean[m_created | m_cur].copy()

    # Extract province codes
    df_clean["REC_PROV"] = df_clean[rec_col].astype(str).str.strip().str.upper()
    df_clean["DEL_PROV"] = df_clean[del_col].astype(str).str.strip().str.upper()
    df_clean["CUR_PROV"] = df_clean[cur_col].astype(str).str.strip().str.upper()
    df_clean["STATUS_CODE"] = df_clean[st_col].astype(str).str.split("-").str[0].str.strip()

    # Collect all unique provinces
    rec_provs = [str(p).strip().upper() for p in df_clean["REC_PROV"].dropna().unique() if pd.notna(p)]
    del_provs = [str(p).strip().upper() for p in df_clean["DEL_PROV"].dropna().unique() if pd.notna(p)]
    all_provs = sorted(list(set(rec_provs) | set(del_provs)))
    all_provs = [p for p in all_provs if p and p not in ("NAN", "NONE", "NULL")]

    # Calculate metrics per province
    summary_rows = []
    for p in all_provs:
        p_name = PROVINCE_NAMES.get(p, p)
        
        # Total received from this province (Origin)
        rec_count = len(df_clean[df_clean["REC_PROV"] == p])
        
        # Total destined to this province (Destination)
        del_count = len(df_clean[df_clean["DEL_PROV"] == p])
        
        # Delivered success (410) destined for this province
        delivered_count = len(df_clean[(df_clean["DEL_PROV"] == p) & (df_clean["STATUS_CODE"] == "410")])
        
        # Pending / In-Transit destined for this province
        pending_count = del_count - delivered_count
        
        # Net balance (Received - Destination)
        net_balance = rec_count - del_count
        
        summary_rows.append({
            "PROVINCE_CODE": p,
            "PROVINCE_NAME": p_name,
            "RECEIVED_COUNT": rec_count,
            "DELIVERY_COUNT": del_count,
            "DELIVERED_SUCCESS_410": delivered_count,
            "PENDING_IN_TRANSIT": pending_count,
            "NET_BALANCE": net_balance
        })

    df_summary = pd.DataFrame(summary_rows)
    df_summary.sort_values(by="DELIVERY_COUNT", ascending=False, inplace=True)

    # Route Matrix (Origin x Destination)
    route_matrix = pd.crosstab(df_clean["REC_PROV"], df_clean["DEL_PROV"], margins=True, margins_name="TOTAL")

    # Build Excel Workbook
    wb = Workbook()
    
    # Styles
    f_title = Font(name="Calibri", size=14, bold=True, color="FFFFFF")
    f_header = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
    f_data = Font(name="Calibri", size=10)
    f_bold = Font(name="Calibri", size=10, bold=True)
    f_green = Font(name="Calibri", size=10, bold=True, color="047857")
    f_blue = Font(name="Calibri", size=10, bold=True, color="1E3A8A")

    fill_title = PatternFill("solid", fgColor="0F172A")
    fill_hdr1  = PatternFill("solid", fgColor="1E293B")
    fill_hdr2  = PatternFill("solid", fgColor="334155")
    fill_alt   = PatternFill("solid", fgColor="F8FAFC")
    fill_tot   = PatternFill("solid", fgColor="E0E7FF")

    thin = Side(border_style="thin", color="CBD5E1")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    align_c = Alignment(horizontal="center", vertical="center")
    align_r = Alignment(horizontal="right", vertical="center")
    align_l = Alignment(horizontal="left", vertical="center")

    # 1. Summary Sheet
    ws_sum = wb.active
    ws_sum.title = "Province Summary"
    
    headers_sum = [
        "PROVINCE CODE", "PROVINCE NAME", 
        "RECEIVED (ORIGIN / POSTED)", "DELIVERY (DESTINATION / TARGET)", 
        "DELIVERED SUCCESS (410)", "PENDING / IN TRANSIT", "NET VOLUME (RECEIVED - DELIVERY)"
    ]

    report_title = title_label or f"PROVINCE RECEIVE VS DELIVERY REPORT ({len(df_clean):,} Total Bills)"
    ws_sum.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(headers_sum))
    t_cell = ws_sum.cell(1, 1, report_title)
    t_cell.font = f_title
    t_cell.fill = fill_title
    t_cell.alignment = align_c
    ws_sum.row_dimensions[1].height = 30

    ws_sum.row_dimensions[3].height = 25
    for c_idx, h_text in enumerate(headers_sum, 1):
        cell = ws_sum.cell(3, c_idx, h_text)
        cell.font = f_header
        cell.fill = fill_hdr1
        cell.alignment = align_c
        cell.border = border

    r_idx = 4
    tot_rec = tot_del = tot_succ = tot_pend = 0

    for row in df_summary.to_dict("records"):
        row_fill = fill_alt if r_idx % 2 == 0 else None
        ws_sum.row_dimensions[r_idx].height = 20

        c_code = ws_sum.cell(r_idx, 1, row["PROVINCE_CODE"])
        c_code.font = f_bold; c_code.border = border; c_code.alignment = align_c
        if row_fill: c_code.fill = row_fill

        c_name = ws_sum.cell(r_idx, 2, row["PROVINCE_NAME"])
        c_name.font = f_data; c_name.border = border; c_name.alignment = align_l
        if row_fill: c_name.fill = row_fill

        c_rec = ws_sum.cell(r_idx, 3, row["RECEIVED_COUNT"])
        c_rec.font = f_blue; c_rec.border = border; c_rec.alignment = align_r
        if row_fill: c_rec.fill = row_fill

        c_del = ws_sum.cell(r_idx, 4, row["DELIVERY_COUNT"])
        c_del.font = f_bold; c_del.border = border; c_del.alignment = align_r
        if row_fill: c_del.fill = row_fill

        c_succ = ws_sum.cell(r_idx, 5, row["DELIVERED_SUCCESS_410"])
        c_succ.font = f_green; c_succ.border = border; c_succ.alignment = align_r
        if row_fill: c_succ.fill = row_fill

        c_pend = ws_sum.cell(r_idx, 6, row["PENDING_IN_TRANSIT"])
        c_pend.font = f_data; c_pend.border = border; c_pend.alignment = align_r
        if row_fill: c_pend.fill = row_fill

        c_net = ws_sum.cell(r_idx, 7, row["NET_BALANCE"])
        c_net.font = f_data; c_net.border = border; c_net.alignment = align_r
        if row_fill: c_net.fill = row_fill

        tot_rec += row["RECEIVED_COUNT"]
        tot_del += row["DELIVERY_COUNT"]
        tot_succ += row["DELIVERED_SUCCESS_410"]
        tot_pend += row["PENDING_IN_TRANSIT"]

        r_idx += 1

    # Total Row
    ws_sum.row_dimensions[r_idx].height = 24
    c_t1 = ws_sum.cell(r_idx, 1, "GRAND TOTAL")
    c_t1.font = f_bold; c_t1.fill = fill_tot; c_t1.border = border; c_t1.alignment = align_c
    ws_sum.merge_cells(start_row=r_idx, start_column=1, end_row=r_idx, end_column=2)

    for c_i, val in [(3, tot_rec), (4, tot_del), (5, tot_succ), (6, tot_pend), (7, tot_rec - tot_del)]:
        cell = ws_sum.cell(r_idx, c_i, val)
        cell.font = f_bold; cell.fill = fill_tot; cell.border = border; cell.alignment = align_r

    ws_sum.freeze_panes = "A4"
    for col in ws_sum.columns:
        max_len = max(len(str(cell.value or "")) for cell in col[:50])
        col_letter = get_column_letter(col[0].column)
        ws_sum.column_dimensions[col_letter].width = max(max_len + 4, 15)

    # 2. Route Flow Matrix Sheet
    ws_mat = wb.create_sheet(title="Route Flow Grid")
    ws_mat.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(route_matrix.columns) + 1)
    t_mat = ws_mat.cell(1, 1, "ORIGIN (RECEIVE PROVINCE) TO DESTINATION (DELIVERY PROVINCE) VOLUME MATRIX")
    t_mat.font = f_title; t_mat.fill = fill_title; t_mat.alignment = align_c
    ws_mat.row_dimensions[1].height = 30

    # Headers
    ws_mat.cell(3, 1, "ORIGIN \\ DEST").font = f_header
    ws_mat.cell(3, 1).fill = fill_hdr1
    ws_mat.cell(3, 1).border = border
    ws_mat.cell(3, 1).alignment = align_c

    for c_idx, dest_code in enumerate(route_matrix.columns, 2):
        cell = ws_mat.cell(3, c_idx, dest_code)
        cell.font = f_header; cell.fill = fill_hdr2; cell.alignment = align_c; cell.border = border

    m_r_idx = 4
    for orig_code, row_series in route_matrix.iterrows():
        is_tot = (orig_code == "TOTAL")
        ws_mat.row_dimensions[m_r_idx].height = 20
        c_orig = ws_mat.cell(m_r_idx, 1, orig_code)
        c_orig.font = f_bold
        c_orig.fill = fill_tot if is_tot else fill_hdr1
        if not is_tot: c_orig.font = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
        c_orig.alignment = align_c; c_orig.border = border

        for c_idx, dest_code in enumerate(route_matrix.columns, 2):
            val = row_series[dest_code]
            cell = ws_mat.cell(m_r_idx, c_idx, val if val > 0 else "")
            cell.font = f_bold if is_tot or dest_code == "TOTAL" else f_data
            if is_tot or dest_code == "TOTAL": cell.fill = fill_tot
            cell.alignment = align_c; cell.border = border
        m_r_idx += 1

    ws_mat.freeze_panes = "B4"

    wb.save(output_xlsx)
    print(f"[SUCCESS] Saved: {output_xlsx}")
    return df_summary

if __name__ == "__main__":
    src_file = os.path.join(HERE, "cache", "detail_master_until_today.xlsx")
    
    # 1. October 2026 (Current Month)
    out_oct = r"C:\Users\DELL\Desktop\Province_Receive_vs_Delivery_Oct2026.xlsx"
    print("\n--- GENERATING OCTOBER 2026 (THIS MONTH) REPORT ---")
    df_oct = generate_province_report(src_file, out_oct, target_month=10, title_label="PROVINCE RECEIVE VS DELIVERY — OCTOBER 2026 (THIS MONTH)")
    print("\nOCTOBER 2026 (THIS MONTH) TOP PROVINCES:")
    print(df_oct.head(15).to_string(index=False))

    # 2. Total Period (September + October 2026 until today)
    out_total = r"C:\Users\DELL\Desktop\Province_Receive_vs_Delivery_Total_Sep_Oct2026.xlsx"
    print("\n--- GENERATING TOTAL PERIOD (SEP + OCT UNTIL TODAY) REPORT ---")
    df_total = generate_province_report(src_file, out_total, target_month=None, title_label="PROVINCE RECEIVE VS DELIVERY — TOTAL UNTIL TODAY (SEP + OCT 2026)")
