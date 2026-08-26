import os
import sys
import copy
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

def build_penalty_report(src_xlsx, out_xlsx, target_label="ALL"):
    """
    Builds the Clean Executive Stagnant Inventory & Handover Penalty Report:
      - If branch code (e.g. SVAP001) is given: Strictly reports ONLY that branch.
      - Sheet 1: INVENTORY PENALTY REPORT (Compact Summary Table)
      - Sheet 2: base (Exact Bill List for that Branch)
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

    today = datetime.now().date()
    tgt = "".join(c for c in str(target_label).upper() if c.isalnum() or c in ("-", "_")).strip()
    if not tgt:
        tgt = "ALL"

    df['sc'] = df[col_status].astype(str).str.extract(r'^(\d{3})')[0]
    df['curr_po_clean'] = df[col_orig_po].astype(str).str.strip().str.upper()
    df['deliv_po_clean'] = df[col_dest_po].astype(str).str.strip().str.upper()

    # Exclude completed 410, 520, cancelled 201, 99, 100, -99
    active_df = df[~df['sc'].isin(['410', '520', '201', '99', '100', '-99'])].copy()

    excused_statuses = {"420", "472"}
    summary_data = {}
    base_rows = []
    r_idx = 1

    # Handover condition
    is_ho_mask = active_df['sc'].isin(['306', '309', '310', '302', '311', '210'])
    # Delivery condition
    is_del_mask = active_df['sc'].str.startswith('4')

    # Strict branch filtering
    if tgt not in ("ALL", "TOTAL"):
        if tgt.startswith("ZONE"):
            zone_by_prefix = {
                "KAN": "ZONE1", "PNP": "ZONE1", "PRE": "ZONE1", "SVA": "ZONE1",
                "KAM": "ZONE2", "KOH": "ZONE2", "SIH": "ZONE2", "SPE": "ZONE2", "TAK": "ZONE2", "KEP": "ZONE2",
                "BAN": "ZONE3", "BAT": "ZONE3", "CHH": "ZONE3", "PUR": "ZONE3", "PAI": "ZONE3",
                "ODD": "ZONE4", "PRH": "ZONE4", "SIE": "ZONE4", "THO": "ZONE4",
                "CHA": "ZONE5", "KRA": "ZONE5", "TBK": "ZONE5", "ROT": "ZONE5", "MON": "ZONE5", "STU": "ZONE5"
            }
            active_df['zone'] = active_df['curr_po_clean'].str[:3].map(zone_by_prefix).fillna("ZONE1")
            active_df = active_df[active_df['zone'] == tgt].copy()
        elif len(tgt) == 3: # Province prefix e.g. SVA, BAT
            active_df = active_df[
                (active_df['curr_po_clean'].str.startswith(tgt)) |
                (active_df['deliv_po_clean'].str.startswith(tgt))
            ].copy()
        else: # Exact Branch e.g. SVAP001
            # For exact branch: Handover at this branch OR Delivery at this branch
            active_df = active_df[
                ((active_df['curr_po_clean'] == tgt) & is_ho_mask) |
                ((active_df['deliv_po_clean'] == tgt) & is_del_mask)
            ].copy()

    for _, row in active_df.iterrows():
        sc = str(row.get('sc', '')).strip()
        curr_po = str(row.get('curr_po_clean', '')).strip()
        deliv_po = str(row.get('deliv_po_clean', '')).strip()

        is_handover = sc in {"306", "309", "310", "302", "311", "210"}
        is_delivery = sc.startswith("4")

        # Assign exact branch
        if tgt not in ("ALL", "TOTAL") and len(tgt) >= 7:
            po = tgt
        else:
            po = deliv_po if is_delivery and deliv_po and deliv_po != 'NAN' else curr_po

        if not po or po == 'NAN':
            continue

        if tgt not in ("ALL", "TOTAL") and len(tgt) >= 7 and po != tgt:
            continue

        if po not in summary_data:
            summary_data[po] = {
                "po": po,
                "total_handover": 0,
                "total_delivery": 0,
                "penalty_handover": 0,
                "penalty_delivery": 0,
                "excused_count": 0,
                "total_fine": 0.0
            }

        order_id = str(row.get(col_order, '')).strip()
        status_raw = str(row.get(col_status, '')).strip()

        act_val = row.get(col_action_time) if pd.notna(row.get(col_action_time)) else row.get(col_created)
        act_date = parse_date(act_val) or parse_date(row.get(col_created))
        age_days = (today - act_date).days if act_date else 0

        fine = 0.0
        risk_level = "Normal"
        is_excused = False

        if is_handover:
            summary_data[po]["total_handover"] += 1
            if age_days >= 3:
                fine = 0.40
                risk_level = "Urgent (> 3 days)"
                summary_data[po]["penalty_handover"] += 1
            elif age_days >= 1:
                fine = 0.10
                risk_level = "Backlog (1-2 days)"
                summary_data[po]["penalty_handover"] += 1
            else:
                risk_level = "Safe (< 1 day)"

        elif is_delivery:
            summary_data[po]["total_delivery"] += 1
            if sc in excused_statuses:
                is_excused = True
                risk_level = f"Excused ({sc})"
                summary_data[po]["excused_count"] += 1
                fine = 0.0
            else:
                if age_days >= 3:
                    fine = 0.40
                    risk_level = "Critical (> 3 days)"
                    summary_data[po]["penalty_delivery"] += 1
                elif age_days >= 1:
                    fine = 0.10
                    risk_level = "Stagnant (1-2 days)"
                    summary_data[po]["penalty_delivery"] += 1
                else:
                    risk_level = "Safe (< 1 day)"

        summary_data[po]["total_fine"] += fine

        base_rows.append({
            "no": r_idx,
            "order_number": order_id,
            "customer": str(row.get(col_receiver, ''))[:30],
            "origin_branch": str(row.get(col_orig_br, '')),
            "origin_post": curr_po,
            "destination_branch": str(row.get(col_dest_prov, '')),
            "destination_post": deliv_po,
            "assigned_branch": po,
            "created_at": str(row.get(col_created, '')),
            "status": status_raw,
            "status_code": sc,
            "type": "Handover" if is_handover else "Delivery",
            "age_days": age_days,
            "excused": "YES (Green $0)" if is_excused else "NO",
            "risk_level": risk_level,
            "penalty_fine": fine
        })
        r_idx += 1

    # Ensure target branch exists in summary even if 0
    if tgt not in ("ALL", "TOTAL") and tgt not in summary_data:
        summary_data[tgt] = {
            "po": tgt,
            "total_handover": 0,
            "total_delivery": 0,
            "penalty_handover": 0,
            "penalty_delivery": 0,
            "excused_count": 0,
            "total_fine": 0.0
        }

    # Build Excel Workbook
    wb = openpyxl.Workbook()

    # Sheet 1: INVENTORY PENALTY REPORT
    ws1 = wb.active
    ws1.title = "INVENTORY PENALTY REPORT"
    ws1.views.sheetView[0].showGridLines = True

    fill_title_left  = PatternFill("solid", fgColor="0F172A") # Deep Slate Navy
    fill_hdr_left    = PatternFill("solid", fgColor="1E293B") # Executive Navy Slate
    fill_row_alt     = PatternFill("solid", fgColor="F8FAFC") # Subtle Zebra Tint
    fill_left_tot    = PatternFill("solid", fgColor="FEE2E2") # Light Red Total
    green_fill       = PatternFill("solid", fgColor="DCFCE7") # Light Green

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
    font_tot_red = Font(name="Segoe UI", size=9, bold=True, color="991B1B")

    # Left Banner
    ws1.merge_cells("A1:G1")
    ws1["A1"].value = f"METFONE EXPRESS — INVENTORY PENALTY & STAGNANT GOODS ({tgt})"
    ws1["A1"].font = font_banner
    ws1["A1"].fill = fill_title_left
    ws1["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws1.row_dimensions[1].height = 24

    ws1.merge_cells("A2:G2")
    ws1["A2"].value = f"Generated: {datetime.now().strftime('%d/%m/%Y %H:%M')} | Excused Green: 420, 472 ($0 fine) | Penalty: 1-2d (-$0.10), >3d (-$0.40)"
    ws1["A2"].font = Font(name="Segoe UI", size=7.5, italic=True, color="FFFFFF")
    ws1["A2"].fill = PatternFill("solid", fgColor="334155")
    ws1["A2"].alignment = Alignment(horizontal="center", vertical="center")
    ws1.row_dimensions[2].height = 18

    left_headers = [
        "No", "Post Office", "Total Handover", "Total Delivery",
        "Penalty Handover", "Penalty Delivery", "Total Penalty ($)"
    ]

    ws1.row_dimensions[3].height = 22
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

    sorted_branches = sorted(summary_data.values(), key=lambda x: x["po"])
    for stats in sorted_branches:
        ws1.row_dimensions[r_curr].height = 18
        row_vals = [
            n_idx,
            stats["po"],
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

    # Grand Total Row
    ws1.row_dimensions[r_curr].height = 22
    ws1.merge_cells(start_row=r_curr, start_column=1, end_row=r_curr, end_column=2)
    gt_left = ws1.cell(r_curr, 1, "Grand Total")
    gt_left.font = font_tot
    gt_left.alignment = Alignment(horizontal="left", vertical="center")

    tot_vals_left = ["", "", tot_ho, tot_del, tot_pen_ho, tot_pen_del, f"-${tot_fine:.2f}" if tot_fine > 0 else "$0.00"]
    for c in range(1, 8):
        cell = ws1.cell(r_curr, c)
        if c >= 3:
            cell.value = tot_vals_left[c-1]
        cell.font = font_tot_red if c == 7 else font_tot
        cell.fill = fill_left_tot
        cell.border = tot_border_accounting
        cell.alignment = Alignment(horizontal="center", vertical="center")

    col_widths_left = {1: 5, 2: 12, 3: 13, 4: 13, 5: 14, 6: 14, 7: 15}
    for c, w in col_widths_left.items():
        ws1.column_dimensions[get_column_letter(c)].width = w

    # ─────────────────────────────────────────────────────────────────────────────
    # SHEET 2: base
    # ─────────────────────────────────────────────────────────────────────────────
    ws2 = wb.create_sheet(title="base")
    ws2.views.sheetView[0].showGridLines = True

    base_headers = [
        "No", "Order Number", "Customer", "Origin Branch", "Origin Post",
        "Destination Branch", "Destination Post", "Assigned Branch", "Created At",
        "Status", "Status Code", "Type", "Age (Days)", "Excused?",
        "Risk / SLA Level", "Penalty Fine ($)"
    ]

    ws2.row_dimensions[1].height = 22
    for col_idx, h in enumerate(base_headers, 1):
        c = ws2.cell(row=1, column=col_idx, value=h)
        c.font = font_hdr
        c.fill = fill_hdr_left
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.border = border_clean

    base_rows.sort(key=lambda x: (x["penalty_fine"] == 0, -x["age_days"]))

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
            item["status"],
            item["status_code"],
            item["type"],
            item["age_days"],
            item["excused"],
            item["risk_level"],
            f"-${item['penalty_fine']:.2f}" if item["penalty_fine"] > 0 else "$0.00"
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
        ws2.column_dimensions[col_letter].width = max(max_len + 3, 11)

    wb.save(out_xlsx)
    return tot_ho, tot_del, (tot_pen_ho + tot_pen_del), tot_fine


def render_penalty_summary_image(out_xlsx):
    """Renders compact HD image preview matching /tomorrow style."""
    import tempfile, copy
    wb = openpyxl.load_workbook(out_xlsx)
    ws = wb['INVENTORY PENALTY REPORT']

    wb_sum = openpyxl.Workbook()
    ws_sum = wb_sum.active
    ws_sum.title = 'Penalty Summary'
    ws_sum.views.sheetView[0].showGridLines = True

    max_r = ws.max_row
    for r in range(1, max_r + 1):
        if ws.row_dimensions[r].height:
            ws_sum.row_dimensions[r].height = ws.row_dimensions[r].height
        for c in range(1, 8):
            cell_orig = ws.cell(r, c)
            cell_tgt = ws_sum.cell(r, c, cell_orig.value)
            if cell_orig.has_style:
                cell_tgt.font = copy.copy(cell_orig.font)
                cell_tgt.fill = copy.copy(cell_orig.fill)
                cell_tgt.border = copy.copy(cell_orig.border)
                cell_tgt.alignment = copy.copy(cell_orig.alignment)

    for c in range(1, 8):
        col_l = get_column_letter(c)
        ws_sum.column_dimensions[col_l].width = ws.column_dimensions[col_l].width or 12

    ws_sum.merge_cells("A1:G1")
    ws_sum.merge_cells("A2:G2")
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
