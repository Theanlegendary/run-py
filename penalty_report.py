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

def build_penalty_report(src_xlsx, out_xlsx, target_label="ALL"):
    """
    Builds the Executive Stagnant Inventory & Handover Penalty Report (/tomorrow style):
      - Sheet 1: INVENTORY PENALTY REPORT (Summary Table + Executive Subtotals)
      - Sheet 2: base (Full Raw / Audit Dataset)
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
    col_fee = next((c for c in df.columns if 'TOTAL FEE' in c or 'TOTAL_AMOUNT' in c), 'TOTAL FEE')
    col_cod = next((c for c in df.columns if 'COD' in c), 'COD')
    col_action_time = next((c for c in df.columns if 'ACTION TIME' in c or 'CURRENT TIME' in c), 'CURRENT TIME')

    today = datetime.now().date()
    tgt = "".join(c for c in str(target_label).upper() if c.isalnum() or c in ("-", "_")).strip()
    if not tgt:
        tgt = "ALL"

    df['sc'] = df[col_status].astype(str).str.extract(r'^(\d{3})')[0]
    df['curr_po_clean'] = df[col_orig_po].astype(str).str.strip().str.upper()
    df['deliv_po_clean'] = df[col_dest_po].astype(str).str.strip().str.upper()

    df['branch'] = df.apply(
        lambda r: r['deliv_po_clean'] if str(r['sc']).startswith('4') and r['deliv_po_clean'] != 'NAN' and r['deliv_po_clean'] else r['curr_po_clean'],
        axis=1
    )

    # Exclude completed 410, 520, cancelled 201, 99, 100, -99
    active_df = df[~df['sc'].isin(['410', '520', '201', '99', '100', '-99'])].copy()

    # Zone mapping
    zone_by_prefix = {
        "KAN": "Zone 1", "PNP": "Zone 1", "PRE": "Zone 1", "SVA": "Zone 1",
        "KAM": "Zone 2", "KOH": "Zone 2", "SIH": "Zone 2", "SPE": "Zone 2", "TAK": "Zone 2", "KEP": "Zone 2",
        "BAN": "Zone 3", "BAT": "Zone 3", "CHH": "Zone 3", "PUR": "Zone 3", "PAI": "Zone 3",
        "ODD": "Zone 4", "PRH": "Zone 4", "SIE": "Zone 4", "THO": "Zone 4",
        "CHA": "Zone 5", "KRA": "Zone 5", "TBK": "Zone 5", "ROT": "Zone 5", "MON": "Zone 5", "STU": "Zone 5"
    }

    # Filter target branch / zone / all
    if tgt not in ("ALL", "TOTAL"):
        if tgt.startswith("ZONE"):
            target_zone_name = tgt.replace("ZONE", "Zone ")
            active_df['zone'] = active_df['branch'].str[:3].map(zone_by_prefix).fillna("Zone 1")
            active_df = active_df[active_df['zone'].str.upper().str.replace(" ", "") == tgt].copy()
        elif len(tgt) == 3:
            active_df = active_df[active_df['branch'].str.startswith(tgt)].copy()
        else:
            active_df = active_df[
                (active_df['branch'] == tgt) |
                (active_df['curr_po_clean'] == tgt) |
                (active_df['deliv_po_clean'] == tgt)
            ].copy()

    excused_statuses = {"420", "472"}
    summary_data = {}
    base_rows = []
    r_idx = 1

    for _, row in active_df.iterrows():
        po = str(row.get('branch', 'OTHER')).strip().upper()
        if not po or po == 'NAN':
            continue

        prov = po[:3] if len(po) >= 3 else 'OTH'
        zone_str = zone_by_prefix.get(prov, "Zone 1")

        key = (zone_str, prov, po)
        if key not in summary_data:
            summary_data[key] = {
                "total_handover": 0,
                "total_delivery": 0,
                "penalty_handover": 0,
                "penalty_delivery": 0,
                "excused_count": 0,
                "total_fine": 0.0
            }

        order_id = str(row.get(col_order, '')).strip()
        sc = str(row.get('sc', '')).strip()
        status_raw = str(row.get(col_status, '')).strip()

        act_val = row.get(col_action_time) if pd.notna(row.get(col_action_time)) else row.get(col_created)
        act_date = parse_date(act_val) or parse_date(row.get(col_created))
        age_days = (today - act_date).days if act_date else 0

        is_handover = sc in {"306", "309", "310", "302", "311", "210"}
        is_delivery = sc.startswith("4")

        fine = 0.0
        risk_level = "Normal"
        is_excused = False

        if is_handover:
            summary_data[key]["total_handover"] += 1
            if age_days >= 3:
                fine = 0.40
                risk_level = "Urgent (> 3 days)"
                summary_data[key]["penalty_handover"] += 1
            elif age_days >= 1:
                fine = 0.10
                risk_level = "Backlog (1-2 days)"
                summary_data[key]["penalty_handover"] += 1
            else:
                risk_level = "Safe (< 1 day)"

        elif is_delivery:
            summary_data[key]["total_delivery"] += 1
            if sc in excused_statuses:
                is_excused = True
                risk_level = f"Excused ({sc})"
                summary_data[key]["excused_count"] += 1
                fine = 0.0
            else:
                if age_days >= 3:
                    fine = 0.40
                    risk_level = "Critical (> 3 days)"
                    summary_data[key]["penalty_delivery"] += 1
                elif age_days >= 1:
                    fine = 0.10
                    risk_level = "Stagnant (1-2 days)"
                    summary_data[key]["penalty_delivery"] += 1
                else:
                    risk_level = "Safe (< 1 day)"

        summary_data[key]["total_fine"] += fine

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
            "status": status_raw,
            "status_code": sc,
            "type": "Handover" if is_handover else "Delivery",
            "age_days": age_days,
            "excused": "YES (Green $0)" if is_excused else "NO",
            "risk_level": risk_level,
            "penalty_fine": fine,
            "zone": zone_str
        })
        r_idx += 1

    # Build Excel Workbook (/tomorrow style)
    wb = openpyxl.Workbook()

    # Sheet 1: INVENTORY PENALTY REPORT
    ws1 = wb.active
    ws1.title = "INVENTORY PENALTY REPORT"
    ws1.views.sheetView[0].showGridLines = True

    # Color Palette matching /tomorrow executive styling
    fill_title_left  = PatternFill("solid", fgColor="0F172A") # Deep Slate Navy
    fill_hdr_left    = PatternFill("solid", fgColor="1E293B") # Executive Navy Slate
    fill_title_right = PatternFill("solid", fgColor="0F766E") # Deep Teal Slate
    fill_hdr_right   = PatternFill("solid", fgColor="0F766E") # Deep Teal Slate
    fill_row_alt     = PatternFill("solid", fgColor="F8FAFC") # Subtle Zebra Tint
    fill_left_tot    = PatternFill("solid", fgColor="FEE2E2") # Light Red Total
    fill_sum_tot     = PatternFill("solid", fgColor="FEE2E2") # Light Red Total
    sub_fill         = PatternFill("solid", fgColor="E0F2FE") # Light Blue Subtotal
    green_fill       = PatternFill("solid", fgColor="DCFCE7") # Light Green

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
    font_tot_red = Font(name="Segoe UI", size=10, bold=True, color="991B1B")

    # Left Banner
    ws1.merge_cells("A1:G1")
    ws1["A1"].value = f"METFONE EXPRESS — INVENTORY PENALTY & STAGNANT GOODS REPORT ({tgt})"
    ws1["A1"].font = font_banner
    ws1["A1"].fill = fill_title_left
    ws1["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws1.row_dimensions[1].height = 28

    ws1.merge_cells("A2:G2")
    ws1["A2"].value = f"Generated: {datetime.now().strftime('%d/%m/%Y %H:%M')} | Excused Green: 420, 472 ($0 fine) | Penalty: 1-2d (-$0.10), >3d (-$0.40)"
    ws1["A2"].font = Font(name="Segoe UI", size=8, italic=True, color="FFFFFF")
    ws1["A2"].fill = PatternFill("solid", fgColor="334155")
    ws1["A2"].alignment = Alignment(horizontal="center", vertical="center")
    ws1.row_dimensions[2].height = 20

    left_headers = [
        "No", "Post Office", "Total Handover", "Total Delivery",
        "Penalty Handover", "Penalty Delivery", "Total Penalty ($)"
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
    tot_ho = 0
    tot_del = 0
    tot_pen_ho = 0
    tot_pen_del = 0
    tot_exc = 0
    tot_fine = 0.0

    sorted_keys = sorted(summary_data.keys())
    for key in sorted_keys:
        stats = summary_data[key]
        ws1.row_dimensions[r_curr].height = 20
        row_vals = [
            n_idx,
            key[2], # PO Code
            stats["total_handover"],
            stats["total_delivery"],
            stats["penalty_handover"],
            stats["penalty_delivery"],
            f"-${stats['total_fine']:.2f}" if stats["total_fine"] > 0 else "$0.00"
        ]
        for col_idx, val in enumerate(row_vals, 1):
            c = ws1.cell(row=r_curr, column=col_idx, value=val)
            c.font = font_bold_data if col_idx in (2, 7) else font_data
            c.alignment = Alignment(horizontal="center", vertical="center")
            c.border = border_clean
            if r_curr % 2 == 0:
                c.fill = fill_row_alt
            if col_idx == 7 and stats["total_fine"] > 0:
                c.font = font_tot_red
                c.fill = PatternFill("solid", fgColor="FEE2E2")

        tot_ho += stats["total_handover"]
        tot_del += stats["total_delivery"]
        tot_pen_ho += stats["penalty_handover"]
        tot_pen_del += stats["penalty_delivery"]
        tot_exc += stats["excused_count"]
        tot_fine += stats["total_fine"]
        r_curr += 1
        n_idx += 1

    # Left Grand Total Row
    ws1.row_dimensions[r_curr].height = 25
    ws1.merge_cells(start_row=r_curr, start_column=1, end_row=r_curr, end_column=2)
    gt_left = ws1.cell(r_curr, 1, "Grand Total / សរុប")
    gt_left.font = font_tot
    gt_left.alignment = Alignment(horizontal="left", vertical="center")

    tot_vals_left = [
        "", "",
        tot_ho,
        tot_del,
        tot_pen_ho,
        tot_pen_del,
        f"-${tot_fine:.2f}" if tot_fine > 0 else "$0.00"
    ]
    for c in range(1, 8):
        cell = ws1.cell(r_curr, c)
        if c >= 3:
            cell.value = tot_vals_left[c-1]
        cell.font = font_tot_red if c == 7 else font_tot
        cell.fill = fill_left_tot
        cell.border = tot_border_accounting
        cell.alignment = Alignment(horizontal="center", vertical="center")

    # Right Executive Summary Table (Executive KPI Block)
    ws1.merge_cells("I1:L1")
    ws1["I1"].value = f"EXECUTIVE SUMMARY — {tgt}"
    ws1["I1"].font = font_banner
    ws1["I1"].fill = fill_title_right
    ws1["I1"].alignment = Alignment(horizontal="center", vertical="center")

    ws1.merge_cells("I2:L2")
    ws1["I2"].value = "Province KPI Subtotals"
    ws1["I2"].font = Font(name="Segoe UI", size=8, italic=True, color="FFFFFF")
    ws1["I2"].fill = PatternFill("solid", fgColor="115E59")
    ws1["I2"].alignment = Alignment(horizontal="center", vertical="center")

    right_headers = ["Zone", "Branch", "Penalized Bills", "Total Fine ($)"]
    for col_idx, h in enumerate(right_headers, 9):
        c = ws1.cell(row=3, column=col_idx, value=h)
        c.font = font_hdr
        c.fill = fill_hdr_right
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.border = border_clean

    # Group by Province
    prov_groups = {}
    for key, stats in summary_data.items():
        prov = key[1]
        if prov not in prov_groups:
            prov_groups[prov] = {"zone": key[0], "penalized": 0, "fine": 0.0}
        prov_groups[prov]["penalized"] += (stats["penalty_handover"] + stats["penalty_delivery"])
        prov_groups[prov]["fine"] += stats["total_fine"]

    r_sum = 4
    for prov in sorted(prov_groups.keys()):
        pdata = prov_groups[prov]
        ws1.row_dimensions[r_sum].height = 20
        s_vals = [pdata["zone"], prov, pdata["penalized"], f"-${pdata['fine']:.2f}" if pdata["fine"] > 0 else "$0.00"]
        for ci, val in enumerate(s_vals, 9):
            cell = ws1.cell(r_sum, ci, val)
            cell.font = font_data
            cell.border = border_clean
            cell.alignment = Alignment(horizontal="center", vertical="center")
            if ci == 12 and pdata["fine"] > 0:
                cell.font = font_tot_red
        r_sum += 1

    # Right Grand Total Row
    ws1.row_dimensions[r_sum].height = 25
    ws1.merge_cells(start_row=r_sum, start_column=9, end_row=r_sum, end_column=10)
    rt_tot = ws1.cell(r_sum, 9, "Total / សរុប")
    rt_tot.font = font_tot
    rt_tot.alignment = Alignment(horizontal="left", vertical="center")

    cell_tot_pen = ws1.cell(r_sum, 11, (tot_pen_ho + tot_pen_del))
    cell_tot_pen.font = font_tot
    cell_tot_pen.alignment = Alignment(horizontal="center", vertical="center")

    cell_tot_fine = ws1.cell(r_sum, 12, f"-${tot_fine:.2f}" if tot_fine > 0 else "$0.00")
    cell_tot_fine.font = font_tot_red
    cell_tot_fine.alignment = Alignment(horizontal="center", vertical="center")

    for c in range(9, 13):
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
        "Status", "Status Code", "Type", "Age (Days)", "Excused?",
        "Risk / SLA Level", "Penalty Fine ($)", "Zone"
    ]

    ws2.row_dimensions[1].height = 24
    for col_idx, h in enumerate(base_headers, 1):
        c = ws2.cell(row=1, column=col_idx, value=h)
        c.font = font_hdr
        c.fill = fill_hdr_left
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.border = border_clean

    base_rows.sort(key=lambda x: (x["penalty_fine"] == 0, -x["age_days"]))

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
            item["status"],
            item["status_code"],
            item["type"],
            item["age_days"],
            item["excused"],
            item["risk_level"],
            f"-${item['penalty_fine']:.2f}" if item["penalty_fine"] > 0 else "$0.00",
            item["zone"]
        ]
        for col_idx, val in enumerate(row_data, 1):
            c = ws2.cell(row=r_num, column=col_idx, value=val)
            c.font = font_data
            c.alignment = Alignment(horizontal="center", vertical="center")
            c.border = border_clean
            if item["excused"].startswith("YES"):
                c.fill = green_fill
            elif item["penalty_fine"] > 0:
                c.fill = PatternFill("solid", fgColor="FEE2E2") if item["penalty_fine"] >= 0.40 else PatternFill("solid", fgColor="FEF3C7")

    for col in ws2.columns:
        max_len = max(len(str(cell.value or "")) for cell in col)
        col_letter = get_column_letter(col[0].column)
        ws2.column_dimensions[col_letter].width = max(max_len + 3, 12)

    wb.save(out_xlsx)
    return tot_ho, tot_del, (tot_pen_ho + tot_pen_del), tot_fine


def render_penalty_summary_image(xlsx_path):
    """Renders high quality summary image preview of Sheet 1 for Telegram."""
    from PIL import Image, ImageDraw, ImageFont
    wb = openpyxl.load_workbook(xlsx_path, data_only=True)
    ws = wb["INVENTORY PENALTY REPORT"]

    # Simple clean PIL table renderer
    cell_w, cell_h = 130, 32
    cols_to_render = 7
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

    # Draw header banner
    draw.rectangle([0, 0, img_w, 50], fill="#0F172A")
    title_text = str(ws["A1"].value or "INVENTORY PENALTY REPORT")
    draw.text((20, 15), title_text, fill="#FFFFFF", font=font_title)

    y_pos = 60
    for r_idx in range(2, total_r):
        row = rows[r_idx][:cols_to_render]
        bg_color = "#1E293B" if r_idx == 2 else ("#FEE2E2" if r_idx == total_r - 1 else ("#F8FAFC" if r_idx % 2 == 0 else "#FFFFFF"))
        draw.rectangle([20, y_pos, img_w - 20, y_pos + cell_h], fill=bg_color)

        x_pos = 20
        for val in row:
            text = str(val or "")
            t_color = "#FFFFFF" if r_idx == 2 else ("#B91C1C" if "$" in text and "-" in text else "#0F172A")
            use_font = font_bold if (r_idx == 2 or r_idx == total_r - 1) else font_main
            draw.text((x_pos + 8, y_pos + 6), text[:18], fill=t_color, font=use_font)
            x_pos += cell_w

        y_pos += cell_h

    buf = io.BytesIO()
    im.save(buf, format="PNG")
    buf.seek(0)
    return buf
