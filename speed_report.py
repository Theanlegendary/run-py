import os
import sys
import io
from datetime import datetime, date
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import pandas as pd

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
    Builds the Executive Daily Fast Delivery Speed Bonus Report (/tomorrow style):
      - Sheet 1: DELIVERY SPEED REPORT (Summary Breakdown + Executive Subtotals)
      - Sheet 2: base (Full Raw / Audit Dataset with Duration & Tier)
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

    df['branch'] = df.apply(
        lambda r: r['deliv_po_clean'] if r['deliv_po_clean'] != 'NAN' and r['deliv_po_clean'] else r['curr_po_clean'],
        axis=1
    )

    # Filter Delivered bills (410)
    delivered_df = df[df['sc'] == '410'].copy()

    zone_by_prefix = {
        "KAN": "Zone 1", "PNP": "Zone 1", "PRE": "Zone 1", "SVA": "Zone 1",
        "KAM": "Zone 2", "KOH": "Zone 2", "SIH": "Zone 2", "SPE": "Zone 2", "TAK": "Zone 2", "KEP": "Zone 2",
        "BAN": "Zone 3", "BAT": "Zone 3", "CHH": "Zone 3", "PUR": "Zone 3", "PAI": "Zone 3",
        "ODD": "Zone 4", "PRH": "Zone 4", "SIE": "Zone 4", "THO": "Zone 4",
        "CHA": "Zone 5", "KRA": "Zone 5", "TBK": "Zone 5", "ROT": "Zone 5", "MON": "Zone 5", "STU": "Zone 5"
    }

    # Filter target branch / zone / all first
    if tgt not in ("ALL", "TOTAL"):
        if tgt.startswith("ZONE"):
            delivered_df['zone'] = delivered_df['branch'].str[:3].map(zone_by_prefix).fillna("Zone 1")
            delivered_df = delivered_df[delivered_df['zone'].str.upper().str.replace(" ", "") == tgt].copy()
        elif len(tgt) == 3:
            delivered_df = delivered_df[delivered_df['branch'].str.startswith(tgt)].copy()
        else:
            delivered_df = delivered_df[
                (delivered_df['branch'] == tgt) |
                (delivered_df['curr_po_clean'] == tgt) |
                (delivered_df['deliv_po_clean'] == tgt)
            ].copy()

    # EVERYDAY FILTER: Filter by delivery date
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
        po = str(row.get('branch', 'OTHER')).strip().upper()
        if not po or po == 'NAN':
            continue

        prov = po[:3] if len(po) >= 3 else 'OTH'
        zone_str = zone_by_prefix.get(prov, "Zone 1")

        key = (zone_str, prov, po)
        if key not in summary_data:
            summary_data[key] = {
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

        summary_data[key]["total_delivered"] += 1

        if duration_hours < 2.0:
            tier = "< 2 Hours (+50%)"
            rate_usd = 0.30
            summary_data[key]["under_2h"] += 1
            tag_color = "GREEN"
        elif duration_hours <= 4.0:
            tier = "2 - 4 Hours (+25%)"
            rate_usd = 0.25
            summary_data[key]["between_2_4h"] += 1
            tag_color = "BLUE"
        elif duration_hours <= 8.0:
            tier = "4 - 8 Hours (Normal)"
            rate_usd = 0.20
            summary_data[key]["between_4_8h"] += 1
            tag_color = "NORMAL"
        else:
            tier = "> 8 Hours (-25% Fine)"
            rate_usd = 0.15
            summary_data[key]["over_8h"] += 1
            tag_color = "RED"

        summary_data[key]["total_commission"] += rate_usd
        dur_str = f"{int(duration_hours)}h {int((duration_hours%1)*60)}m" if duration_hours else "N/A"

        base_rows.append({
            "no": r_idx,
            "order_number": order_id,
            "customer": str(row.get(col_receiver, ''))[:30],
            "origin_branch": str(row.get(col_orig_br, '')),
            "origin_post": str(row.get(col_orig_po, '')),
            "destination_branch": str(row.get(col_dest_prov, '')),
            "destination_post": str(row.get(col_dest_po, '')),
            "assigned_branch": po,
            "created_at": str(row.get(col_created, '')),
            "t400": str(t400 or ""),
            "t410": str(t410 or ""),
            "duration": dur_str,
            "duration_hours": round(duration_hours, 2),
            "tier": tier,
            "rate_usd": rate_usd,
            "tag_color": tag_color,
            "zone": zone_str
        })
        r_idx += 1

    # Build Excel Workbook (/tomorrow style)
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
    fill_left_tot    = PatternFill("solid", fgColor="DCFCE7") # Light Green Total
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

    font_banner = Font(name="Segoe UI", size=11, bold=True, color="FFFFFF")
    font_hdr = Font(name="Segoe UI", size=9, bold=True, color="FFFFFF")
    font_data = Font(name="Segoe UI", size=9, color="0F172A")
    font_bold_data = Font(name="Segoe UI", size=9, bold=True, color="0F172A")
    font_tot = Font(name="Segoe UI", size=10, bold=True, color="0F172A")
    font_tot_green = Font(name="Segoe UI", size=10, bold=True, color="15803D")

    date_str = today.strftime('%d/%m/%Y')

    # Left Banner
    ws1.merge_cells("A1:H1")
    ws1["A1"].value = f"METFONE EXPRESS — DAILY DELIVERY SPEED BONUS REPORT ({tgt}) — {date_str}"
    ws1["A1"].font = font_banner
    ws1["A1"].fill = fill_title_left
    ws1["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws1.row_dimensions[1].height = 28

    ws1.merge_cells("A2:H2")
    ws1["A2"].value = f"Date: {date_str} | <2h (+50% / $0.30) | 2-4h (+25% / $0.25) | 4-8h ($0.20) | >8h (-25% / $0.15)"
    ws1["A2"].font = Font(name="Segoe UI", size=8, italic=True, color="FFFFFF")
    ws1["A2"].fill = PatternFill("solid", fgColor="334155")
    ws1["A2"].alignment = Alignment(horizontal="center", vertical="center")
    ws1.row_dimensions[2].height = 20

    left_headers = [
        "No", "Post Office", "Delivered (410)", "< 2 Hours (+50%)",
        "2 - 4 Hours (+25%)", "4 - 8 Hours (Normal)", "> 8 Hours (-25%)", "Commission ($)"
    ]

    ws1.row_dimensions[3].height = 24
    for col_idx, h in enumerate(left_headers, 1):
        c = ws1.cell(row=3, column=col_idx, value=h)
        c.font = font_hdr
        c.fill = fill_hdr_left
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = border_clean

    r_curr = 4
    n_idx = 1
    tot_del = 0
    tot_u2 = 0
    tot_24 = 0
    tot_48 = 0
    tot_o8 = 0
    tot_pay = 0.0

    sorted_keys = sorted(summary_data.keys())
    for key in sorted_keys:
        stats = summary_data[key]
        ws1.row_dimensions[r_curr].height = 20
        row_vals = [
            n_idx,
            key[2],
            stats["total_delivered"],
            stats["under_2h"],
            stats["between_2_4h"],
            stats["between_4_8h"],
            stats["over_8h"],
            f"${stats['total_commission']:.2f}"
        ]
        for col_idx, val in enumerate(row_vals, 1):
            c = ws1.cell(row=r_curr, column=col_idx, value=val)
            c.font = font_bold_data if col_idx in (2, 8) else font_data
            c.alignment = Alignment(horizontal="center", vertical="center")
            c.border = border_clean
            if r_curr % 2 == 0:
                c.fill = fill_row_alt
            if col_idx == 8:
                c.font = font_tot_green

        tot_del += stats["total_delivered"]
        tot_u2 += stats["under_2h"]
        tot_24 += stats["between_2_4h"]
        tot_48 += stats["between_4_8h"]
        tot_o8 += stats["over_8h"]
        tot_pay += stats["total_commission"]
        r_curr += 1
        n_idx += 1

    # Left Grand Total Row
    ws1.row_dimensions[r_curr].height = 25
    ws1.merge_cells(start_row=r_curr, start_column=1, end_row=r_curr, end_column=2)
    gt_left = ws1.cell(r_curr, 1, "Grand Total / សរុប")
    gt_left.font = font_tot
    gt_left.alignment = Alignment(horizontal="left", vertical="center")

    tot_vals_left = ["", "", tot_del, tot_u2, tot_24, tot_48, tot_o8, f"${tot_pay:.2f}"]
    for c in range(1, 9):
        cell = ws1.cell(r_curr, c)
        if c >= 3:
            cell.value = tot_vals_left[c-1]
        cell.font = font_tot_green if c == 8 else font_tot
        cell.fill = fill_left_tot
        cell.border = tot_border_accounting
        cell.alignment = Alignment(horizontal="center", vertical="center")

    # Right Executive Summary Table
    ws1.merge_cells("J1:M1")
    ws1["J1"].value = f"EXECUTIVE SUMMARY — {tgt}"
    ws1["J1"].font = font_banner
    ws1["J1"].fill = fill_title_right
    ws1["J1"].alignment = Alignment(horizontal="center", vertical="center")

    ws1.merge_cells("J2:M2")
    ws1["J2"].value = "Province KPI Subtotals"
    ws1["J2"].font = Font(name="Segoe UI", size=8, italic=True, color="FFFFFF")
    ws1["J2"].fill = PatternFill("solid", fgColor="115E59")
    ws1["J2"].alignment = Alignment(horizontal="center", vertical="center")

    right_headers = ["Zone", "Branch", "Fast Delivery (<4h)", "Commission ($)"]
    for col_idx, h in enumerate(right_headers, 10):
        c = ws1.cell(row=3, column=col_idx, value=h)
        c.font = font_hdr
        c.fill = fill_hdr_right
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.border = border_clean

    prov_groups = {}
    for key, stats in summary_data.items():
        prov = key[1]
        if prov not in prov_groups:
            prov_groups[prov] = {"zone": key[0], "fast": 0, "pay": 0.0}
        prov_groups[prov]["fast"] += (stats["under_2h"] + stats["between_2_4h"])
        prov_groups[prov]["pay"] += stats["total_commission"]

    r_sum = 4
    for prov in sorted(prov_groups.keys()):
        pdata = prov_groups[prov]
        ws1.row_dimensions[r_sum].height = 20
        s_vals = [pdata["zone"], prov, pdata["fast"], f"${pdata['pay']:.2f}"]
        for ci, val in enumerate(s_vals, 10):
            cell = ws1.cell(r_sum, ci, val)
            cell.font = font_data
            cell.border = border_clean
            cell.alignment = Alignment(horizontal="center", vertical="center")
            if ci == 13:
                cell.font = font_tot_green
        r_sum += 1

    # Right Grand Total Row
    ws1.row_dimensions[r_sum].height = 25
    ws1.merge_cells(start_row=r_sum, start_column=10, end_row=r_sum, end_column=11)
    rt_tot = ws1.cell(r_sum, 10, "Total / សរុប")
    rt_tot.font = font_tot
    rt_tot.alignment = Alignment(horizontal="left", vertical="center")

    cell_tot_fast = ws1.cell(r_sum, 12, (tot_u2 + tot_24))
    cell_tot_fast.font = font_tot
    cell_tot_fast.alignment = Alignment(horizontal="center", vertical="center")

    cell_tot_p = ws1.cell(r_sum, 13, f"${tot_pay:.2f}")
    cell_tot_p.font = font_tot_green
    cell_tot_p.alignment = Alignment(horizontal="center", vertical="center")

    for c in range(10, 14):
        cell = ws1.cell(r_sum, c)
        cell.fill = fill_sum_tot
        cell.border = tot_border_accounting

    for col in ws1.columns:
        max_len = max(len(str(cell.value or "")) for cell in col)
        col_letter = get_column_letter(col[0].column)
        ws1.column_dimensions[col_letter].width = max(max_len + 3, 12)

    # ─────────────────────────────────────────────────────────────────────────────
    # SHEET 2: base (Full Raw / Audit Dataset matching /tomorrow style)
    # ─────────────────────────────────────────────────────────────────────────────
    ws2 = wb.create_sheet(title="base")
    ws2.views.sheetView[0].showGridLines = True

    base_headers = [
        "No", "Order Number", "Customer", "Origin Branch", "Origin Post",
        "Destination Branch", "Destination Post", "Assigned Branch", "Created At",
        "Out for Delivery (400)", "Delivered (410)", "Duration", "Hours (Dec)",
        "Speed Category", "Rate ($/Bill)", "Zone"
    ]

    ws2.row_dimensions[1].height = 24
    for col_idx, h in enumerate(base_headers, 1):
        c = ws2.cell(row=1, column=col_idx, value=h)
        c.font = font_hdr
        c.fill = fill_hdr_left
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.border = border_clean

    for idx, item in enumerate(base_rows, 1):
        r_num = idx + 1
        ws2.row_dimensions[r_num].height = 20
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
            f"${item['rate_usd']:.2f}",
            item["zone"]
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
        ws2.column_dimensions[col_letter].width = max(max_len + 3, 12)

    wb.save(out_xlsx)
    return tot_del, tot_u2, tot_24, tot_o8, tot_pay


def render_speed_summary_image(xlsx_path):
    """Renders high quality summary image preview of Sheet 1 for Telegram."""
    from PIL import Image, ImageDraw, ImageFont
    wb = openpyxl.load_workbook(xlsx_path, data_only=True)
    ws = wb["DELIVERY SPEED REPORT"]

    cell_w, cell_h = 130, 32
    cols_to_render = 8
    rows = list(ws.iter_rows(values_only=True))
    total_r = min(len(rows), 45)

    img_w = cols_to_render * cell_w + 40
    img_h = total_r * cell_h + 80

    im = Image.new("RGB", (img_w, img_h), color="#FFFFFF")
    draw = ImageDraw.Draw(im)

    try:
        font_main = ImageFont.truetype("seguiemj.ttf", 14)
        font_bold = ImageFont.truetype("seguisb.ttf", 15)
        font_title = ImageFont.truetype("seguisb.ttf", 18)
    except:
        font_main = ImageFont.load_default()
        font_bold = font_main
        font_title = font_main

    draw.rectangle([0, 0, img_w, 50], fill="#0F172A")
    title_text = str(ws["A1"].value or "DELIVERY SPEED REPORT")
    draw.text((20, 15), title_text, fill="#FFFFFF", font=font_title)

    y_pos = 60
    for r_idx in range(2, total_r):
        row = rows[r_idx][:cols_to_render]
        bg_color = "#1E293B" if r_idx == 2 else ("#DCFCE7" if r_idx == total_r - 1 else ("#F8FAFC" if r_idx % 2 == 0 else "#FFFFFF"))
        draw.rectangle([20, y_pos, img_w - 20, y_pos + cell_h], fill=bg_color)

        x_pos = 20
        for val in row:
            text = str(val or "")
            t_color = "#FFFFFF" if r_idx == 2 else ("#15803D" if "$" in text else "#0F172A")
            use_font = font_bold if (r_idx == 2 or r_idx == total_r - 1) else font_main
            draw.text((x_pos + 8, y_pos + 6), text[:18], fill=t_color, font=use_font)
            x_pos += cell_w

        y_pos += cell_h

    buf = io.BytesIO()
    im.save(buf, format="PNG")
    buf.seek(0)
    return buf
