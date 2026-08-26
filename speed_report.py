import os
import sys
import copy
import tempfile
from datetime import datetime, date
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import pandas as pd
import excel_to_image

def parse_date(val):
    if not val or pd.isna(val):
        return None
    if isinstance(val, (datetime, date)):
        return val.date() if isinstance(val, datetime) else val
    s = str(val).strip().split(" ")[0]
    for fmt in ("%d/%m/%Y", "%Y-%m-%d", "%m/%d/%Y", "%d-%m-%Y"):
        try:
            return datetime.strptime(s, fmt).date()
        except ValueError:
            continue
    return None

def parse_time(val):
    if not val or pd.isna(val):
        return None
    if isinstance(val, datetime):
        return val
    s = str(val).strip()
    for fmt in ("%d/%m/%Y %H:%M:%S", "%Y-%m-%d %H:%M:%S", "%d/%m/%Y %H:%M", "%Y-%m-%d %H:%M"):
        try:
            return datetime.strptime(s, fmt)
        except ValueError:
            continue
    return None

def build_speed_report(src_xlsx, out_xlsx, target_label="ALL", report_date=None):
    """
    Builds the Fast Delivery Speed Bonus Report (/tomorrow layout):
      - Sheet 1: Left Table = Detailed Delivery Bills, Right Table = Executive Summary Table
      - Sheet 2: base = Complete raw audit dataset
    """
    os.makedirs(os.path.dirname(os.path.abspath(out_xlsx)), exist_ok=True)
    df = pd.read_excel(src_xlsx)
    df.columns = [str(c).strip().upper() for c in df.columns]

    col_order = next((c for c in df.columns if 'ORDER ID' in c or 'ORDER' in c), 'ORDER ID')
    col_dest_prov = next((c for c in df.columns if 'DELIVERY PROVINCE' in c or 'DESTINATION_BRANCH' in c), 'DELIVERY PROVINCE')
    col_dest_po = next((c for c in df.columns if 'DELIVERY POST' in c or 'DESTINATION_POST' in c), 'DELIVERY POST OFFICE')
    col_orig_br = next((c for c in df.columns if 'ACTION POST OFFICE' in c or 'ORIGIN_BRANCH' in c), 'ACTION POST OFFICE')
    col_orig_po = next((c for c in df.columns if 'CURRENT POST OFFICE' in c or 'ORIGIN_POST' in c), 'CURRENT POST OFFICE')
    col_status = next((c for c in df.columns if 'CURRENT STATUS' in c or 'STATUS' in c), 'CURRENT STATUS')
    col_created = next((c for c in df.columns if 'CREATED DATE' in c), 'CREATED DATE')
    col_receiver = next((c for c in df.columns if 'RECEIVER' in c), 'RECEIVER')
    col_action_time = next((c for c in df.columns if 'ACTION TIME' in c or 'CURRENT TIME' in c), 'CURRENT TIME')
    col_400_time = next((c for c in df.columns if '400' in c or 'OUT' in c or 'ASSIGN' in c), None)

    today = report_date or datetime.now().date()
    tgt = "".join(c for c in str(target_label).upper() if c.isalnum() or c in ("-", "_")).strip()
    if not tgt:
        tgt = "ALL"

    df['sc'] = df[col_status].astype(str).str.extract(r'^(\d{3})')[0]
    df['curr_po_clean'] = df[col_orig_po].astype(str).str.strip().str.upper()
    df['deliv_po_clean'] = df[col_dest_po].astype(str).str.strip().str.upper()

    delivered_df = df[df['sc'] == '410'].copy()

    if tgt not in ("ALL", "TOTAL"):
        if tgt.startswith("ZONE"):
            zone_by_prefix = {
                "KAN": "ZONE1", "PNP": "ZONE1", "PRE": "ZONE1", "SVA": "ZONE1",
                "KAM": "ZONE2", "KOH": "ZONE2", "SIH": "ZONE2", "SPE": "ZONE2", "TAK": "ZONE2", "KEP": "ZONE2",
                "BAN": "ZONE3", "BAT": "ZONE3", "CHH": "ZONE3", "PUR": "ZONE3", "PAI": "ZONE3",
                "ODD": "ZONE4", "PRH": "ZONE4", "SIE": "ZONE4", "THO": "ZONE4",
                "CHA": "ZONE5", "KRA": "ZONE5", "TBK": "ZONE5", "ROT": "ZONE5", "MON": "ZONE5", "STU": "ZONE5"
            }
            delivered_df['zone'] = delivered_df['deliv_po_clean'].str[:3].map(zone_by_prefix).fillna("ZONE1")
            delivered_df = delivered_df[delivered_df['zone'] == tgt].copy()
        elif len(tgt) == 3:
            delivered_df = delivered_df[
                (delivered_df['deliv_po_clean'].str.startswith(tgt)) |
                (delivered_df['curr_po_clean'].str.startswith(tgt))
            ].copy()
        else:
            delivered_df = delivered_df[
                (delivered_df['deliv_po_clean'] == tgt) |
                (delivered_df['curr_po_clean'] == tgt)
            ].copy()

    # EVERYDAY FILTER
    def get_deliv_date(row):
        t410 = parse_time(row.get(col_action_time)) or parse_time(row.get(col_created))
        return t410.date() if t410 else None

    delivered_df['deliv_date'] = delivered_df.apply(get_deliv_date, axis=1)

    today_df = delivered_df[delivered_df['deliv_date'] == today].copy()
    if len(today_df) > 0:
        delivered_df = today_df
    elif len(delivered_df) > 0:
        latest_date = delivered_df['deliv_date'].dropna().max()
        if latest_date:
            today = latest_date
            delivered_df = delivered_df[delivered_df['deliv_date'] == latest_date].copy()

    summary_data = {}
    base_rows = []
    r_idx = 1

    for idx, row in delivered_df.iterrows():
        deliv_po = str(row.get('deliv_po_clean', '')).strip()
        curr_po = str(row.get('curr_po_clean', '')).strip()

        po = tgt if (tgt not in ("ALL", "TOTAL") and len(tgt) >= 7) else (deliv_po if deliv_po and deliv_po != 'NAN' else curr_po)

        if not po or po == 'NAN':
            continue

        if tgt not in ("ALL", "TOTAL") and len(tgt) >= 7 and po != tgt:
            continue

        if po not in summary_data:
            summary_data[po] = {
                "po": po,
                "total_delivered": 0,
                "under_2h": 0,
                "between_2_4h": 0,
                "between_4_8h": 0,
                "over_8h": 0,
                "total_commission": 0.0
            }

        order_id = str(row.get(col_order, '')).strip()
        t410 = parse_time(row.get(col_action_time)) or parse_time(row.get(col_created))
        t400 = parse_time(row.get(col_400_time)) if col_400_time else None
        if not t400 and t410:
            t400 = parse_time(row.get(col_created))

        duration_hours = (t410 - t400).total_seconds() / 3600.0 if (t410 and t400 and t410 >= t400) else 1.5

        summary_data[po]["total_delivered"] += 1

        if duration_hours < 2.0:
            tier = "< 2 Hours (+50%)"
            rate_usd = 0.30
            summary_data[po]["under_2h"] += 1
            tag_color = "GREEN"
        elif duration_hours <= 4.0:
            tier = "2 - 4 Hours (+25%)"
            rate_usd = 0.25
            summary_data[po]["between_2_4h"] += 1
            tag_color = "BLUE"
        elif duration_hours <= 8.0:
            tier = "4 - 8 Hours (Normal)"
            rate_usd = 0.20
            summary_data[po]["between_4_8h"] += 1
            tag_color = "NORMAL"
        else:
            tier = "> 8 Hours (-25% Fine)"
            rate_usd = 0.15
            summary_data[po]["over_8h"] += 1
            tag_color = "RED"

        summary_data[po]["total_commission"] += rate_usd
        dur_str = f"{int(duration_hours)}h {int((duration_hours%1)*60)}m" if duration_hours else "N/A"

        base_rows.append({
            "no": r_idx,
            "order_number": order_id,
            "customer": str(row.get(col_receiver, ''))[:28],
            "origin_branch": str(row.get(col_orig_br, '')),
            "origin_post": curr_po,
            "destination_branch": str(row.get(col_dest_prov, '')),
            "destination_post": deliv_po,
            "assigned_branch": po,
            "created_at": str(row.get(col_created, '')),
            "t400": str(t400 or ""),
            "t410": str(t410 or ""),
            "duration": dur_str,
            "duration_hours": round(duration_hours, 2),
            "tier": tier,
            "rate_usd": rate_usd,
            "tag_color": tag_color
        })
        r_idx += 1

    if tgt not in ("ALL", "TOTAL") and tgt not in summary_data:
        summary_data[tgt] = {
            "po": tgt,
            "total_delivered": 0,
            "under_2h": 0,
            "between_2_4h": 0,
            "between_4_8h": 0,
            "over_8h": 0,
            "total_commission": 0.0
        }

    # Build Excel Workbook
    wb = openpyxl.Workbook()

    # Sheet 1: DELIVERY SPEED REPORT
    ws1 = wb.active
    ws1.title = "DELIVERY SPEED REPORT"
    ws1.views.sheetView[0].showGridLines = True

    fill_title_left  = PatternFill("solid", fgColor="0F172A") # Deep Slate Navy
    fill_hdr_left    = PatternFill("solid", fgColor="1E293B") # Executive Navy Slate
    fill_title_right = PatternFill("solid", fgColor="0F766E") # Deep Teal Slate
    fill_hdr_right   = PatternFill("solid", fgColor="0F766E") # Deep Teal Slate
    fill_row_alt     = PatternFill("solid", fgColor="F8FAFC") # Subtle Zebra Tint
    fill_left_tot    = PatternFill("solid", fgColor="CBD5E1") # Refined Slate Grey Total
    fill_sum_tot     = PatternFill("solid", fgColor="DCFCE7") # Light Green Total
    green_fill       = PatternFill("solid", fgColor="DCFCE7") # Light Green
    blue_fill        = PatternFill("solid", fgColor="EFF6FF") # Light Blue
    red_fill         = PatternFill("solid", fgColor="FEE2E2") # Light Red

    border_clean = Border(
        left=Side(style="thin", color="E2E8F0"), right=Side(style="thin", color="E2E8F0"),
        top=Side(style="thin", color="E2E8F0"), bottom=Side(style="thin", color="E2E8F0")
    )
    tot_border_accounting = Border(
        left=Side(style="thin", color="CBD5E1"), right=Side(style="thin", color="CBD5E1"),
        top=Side(style="thin", color="64748B"), bottom=Side(style="double", color="0F172A")
    )

    font_banner = Font(name="Segoe UI", size=10, bold=True, color="FFFFFF")
    font_hdr = Font(name="Segoe UI", size=8.5, bold=True, color="FFFFFF")
    font_data = Font(name="Segoe UI", size=8.5, color="0F172A")
    font_bold_data = Font(name="Segoe UI", size=8.5, bold=True, color="0F172A")
    font_tot = Font(name="Segoe UI", size=9, bold=True, color="0F172A")
    font_tot_green = Font(name="Segoe UI", size=9, bold=True, color="15803D")

    date_str = today.strftime('%d/%m/%Y')

    # 1. Left Title Banner
    ws1.merge_cells("A1:H1")
    ws1.cell(1, 1, f"DAILY DELIVERY SPEED ORDER DETAIL ({tgt}) — {date_str}").font = font_banner
    ws1.cell(1, 1).alignment = Alignment(horizontal="left", vertical="center")
    for c in range(1, 9):
        ws1.cell(1, c).fill = fill_title_left

    # 2. Right Title Banner
    ws1.merge_cells("J1:Q1")
    ws1.cell(1, 10, f"EXECUTIVE SUMMARY ({tgt})").font = font_banner
    ws1.cell(1, 10).alignment = Alignment(horizontal="center", vertical="center")
    for c in range(10, 18):
        ws1.cell(1, c).fill = fill_title_right

    ws1.row_dimensions[1].height = 28.0

    # Row 2: Headers
    headers_left = [
        "No", "Order Number", "Customer", "Post Office",
        "Duration", "Hours (Dec)", "Speed Category", "Rate ($/Bill)"
    ]
    headers_right = [
        "No", "Post Office", "Delivered (410)", "< 2 Hours (+50%)",
        "2 - 4 Hours (+25%)", "4 - 8 Hours (Normal)", "> 8 Hours (-25%)", "Commission ($)"
    ]

    ws1.row_dimensions[2].height = 24.0
    for ci, h in enumerate(headers_left, 1):
        cell = ws1.cell(2, ci, h)
        cell.font = font_hdr
        cell.fill = fill_hdr_left
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = border_clean

    for ci, h in enumerate(headers_right, 10):
        cell = ws1.cell(2, ci, h)
        cell.font = font_hdr
        cell.fill = fill_hdr_right
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = border_clean

    # Populate Left Detail Order Rows
    r_curr = 3
    tot_pay_left = 0.0

    for idx, item in enumerate(base_rows, 1):
        ws1.row_dimensions[r_curr].height = 18.0
        row_vals = [
            idx,
            item["order_number"],
            item["customer"],
            item["assigned_branch"],
            item["duration"],
            item["duration_hours"],
            item["tier"],
            f"${item['rate_usd']:.2f}"
        ]
        for col_idx, val in enumerate(row_vals, 1):
            c = ws1.cell(row=r_curr, column=col_idx, value=val)
            c.font = font_data
            c.alignment = Alignment(horizontal="center", vertical="center")
            c.border = border_clean
            if r_curr % 2 == 0:
                c.fill = fill_row_alt
            if item["tag_color"] == "GREEN":
                c.fill = green_fill
            elif item["tag_color"] == "BLUE":
                c.fill = blue_fill
            elif item["tag_color"] == "RED":
                c.fill = red_fill

        tot_pay_left += item["rate_usd"]
        r_curr += 1

    # Left Grand Total
    ws1.row_dimensions[r_curr].height = 22.0
    ws1.merge_cells(start_row=r_curr, start_column=1, end_row=r_curr, end_column=2)
    gt_left = ws1.cell(r_curr, 1, f"Total Delivered: {len(base_rows)}")
    gt_left.font = font_tot
    gt_left.alignment = Alignment(horizontal="left", vertical="center")

    for c in range(1, 9):
        cell = ws1.cell(r_curr, c)
        cell.fill = fill_left_tot
        cell.border = tot_border_accounting
        if c == 8:
            cell.value = f"${tot_pay_left:.2f}"
            cell.font = font_tot_green
            cell.alignment = Alignment(horizontal="center", vertical="center")

    # Populate Right Executive Summary Table
    r_sum = 3
    n_idx = 1
    tot_del = 0
    tot_u2 = 0
    tot_24 = 0
    tot_48 = 0
    tot_o8 = 0
    tot_pay = 0.0

    sorted_branches = sorted(summary_data.values(), key=lambda x: x["po"])
    for stats in sorted_branches:
        ws1.row_dimensions[r_sum].height = 18.0
        s_vals = [
            n_idx,
            stats["po"],
            stats["total_delivered"],
            stats["under_2h"],
            stats["between_2_4h"],
            stats["between_4_8h"],
            stats["over_8h"],
            f"${stats['total_commission']:.2f}"
        ]
        for ci, val in enumerate(s_vals, 10):
            cell = ws1.cell(r_sum, ci, val)
            cell.font = font_bold_data if ci in (11, 17) else font_data
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border = border_clean
            if ci == 17:
                cell.font = font_tot_green

        tot_del += stats["total_delivered"]
        tot_u2 += stats["under_2h"]
        tot_24 += stats["between_2_4h"]
        tot_48 += stats["between_4_8h"]
        tot_o8 += stats["over_8h"]
        tot_pay += stats["total_commission"]
        r_sum += 1
        n_idx += 1

    # Right Grand Total Row
    ws1.row_dimensions[r_sum].height = 22.0
    ws1.merge_cells(start_row=r_sum, start_column=10, end_row=r_sum, end_column=11)
    rt_tot = ws1.cell(r_sum, 10, "Grand Total")
    rt_tot.font = font_tot
    rt_tot.alignment = Alignment(horizontal="left", vertical="center")

    tot_vals_right = ["", "", tot_del, tot_u2, tot_24, tot_48, tot_o8, f"${tot_pay:.2f}"]
    for c in range(10, 18):
        cell = ws1.cell(r_sum, c)
        if c >= 12:
            cell.value = tot_vals_right[c-10]
        cell.font = font_tot_green if c == 17 else font_tot
        cell.fill = fill_sum_tot
        cell.border = tot_border_accounting
        cell.alignment = Alignment(horizontal="center", vertical="center")

    col_widths = {
        1: 5, 2: 15, 3: 20, 4: 12, 5: 12, 6: 12, 7: 18, 8: 14,
        9: 4, # Gap
        10: 5, 11: 12, 12: 13, 13: 14, 14: 14, 15: 14, 16: 14, 17: 15
    }
    for c, w in col_widths.items():
        ws1.column_dimensions[get_column_letter(c)].width = w

    # ─────────────────────────────────────────────────────────────────────────────
    # SHEET 2: base
    # ─────────────────────────────────────────────────────────────────────────────
    ws2 = wb.create_sheet(title="base")
    ws2.views.sheetView[0].showGridLines = True

    base_headers = [
        "No", "Order Number", "Customer", "Origin Branch", "Origin Post",
        "Destination Branch", "Destination Post", "Assigned Branch", "Created At",
        "Out for Delivery (400)", "Delivered (410)", "Duration", "Hours (Dec)",
        "Speed Category", "Rate ($/Bill)"
    ]

    ws2.row_dimensions[1].height = 22
    for col_idx, h in enumerate(base_headers, 1):
        c = ws2.cell(row=1, column=col_idx, value=h)
        c.font = font_hdr
        c.fill = fill_hdr_left
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.border = border_clean

    for idx, item in enumerate(base_rows, 1):
        r_num = idx + 1
        ws2.row_dimensions[r_num].height = 18
        row_data = [
            idx,
            item["order_number"],
            item["customer"],
            item["origin_branch"],
            item["origin_post"],
            item["destination_branch"],
            item["destination_post"],
            item["assigned_branch"],
            item["created_at"],
            item["t400"],
            item["t410"],
            item["duration"],
            item["duration_hours"],
            item["tier"],
            f"${item['rate_usd']:.2f}"
        ]
        for col_idx, val in enumerate(row_data, 1):
            c = ws2.cell(row=r_num, column=col_idx, value=val)
            c.font = font_data
            c.alignment = Alignment(horizontal="center", vertical="center")
            c.border = border_clean
            if item["tag_color"] == "GREEN":
                c.fill = green_fill
            elif item["tag_color"] == "BLUE":
                c.fill = blue_fill
            elif item["tag_color"] == "RED":
                c.fill = red_fill

    for col in ws2.columns:
        max_len = max(len(str(cell.value or "")) for cell in col)
        col_letter = get_column_letter(col[0].column)
        ws2.column_dimensions[col_letter].width = max(max_len + 3, 11)

    wb.save(out_xlsx)
    return tot_del, tot_u2, tot_24, tot_o8, tot_pay


def render_speed_summary_image(out_xlsx):
    """Renders ONLY the right Executive Summary Table as a crisp PNG image (matching /tomorrow)."""
    wb = openpyxl.load_workbook(out_xlsx)
    ws = wb['DELIVERY SPEED REPORT']

    wb_sum = openpyxl.Workbook()
    ws_sum = wb_sum.active
    ws_sum.title = 'Executive Summary'
    ws_sum.views.sheetView[0].showGridLines = True

    max_r = 1
    for r in range(1, ws.max_row + 1):
        if ws.cell(r, 10).value is not None or ws.cell(r, 17).value is not None:
            max_r = r

    for r in range(1, max_r + 1):
        if ws.row_dimensions[r].height:
            ws_sum.row_dimensions[r].height = ws.row_dimensions[r].height
        for c_idx in range(8):
            orig_c = 10 + c_idx
            tgt_c = 1 + c_idx
            cell_orig = ws.cell(r, orig_c)
            cell_tgt = ws_sum.cell(r, tgt_c, cell_orig.value)
            if cell_orig.has_style:
                cell_tgt.font = copy.copy(cell_orig.font)
                cell_tgt.fill = copy.copy(cell_orig.fill)
                cell_tgt.border = copy.copy(cell_orig.border)
                cell_tgt.alignment = copy.copy(cell_orig.alignment)

    col_widths = {1: 5, 2: 12, 3: 13, 4: 14, 5: 14, 6: 14, 7: 14, 8: 15}
    for c, w in col_widths.items():
        ws_sum.column_dimensions[get_column_letter(c)].width = w

    ws_sum.merge_cells("A1:H1")
    ws_sum.merge_cells(start_row=max_r, start_column=1, end_row=max_r, end_column=2)

    with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as tmp_f:
        tmp_path = tmp_f.name
    wb_sum.save(tmp_path)

    try:
        buf = excel_to_image.excel_to_image(tmp_path)
        return buf
    finally:
        try:
            os.remove(tmp_path)
        except Exception:
            pass
