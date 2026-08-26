import os
import sys
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
    os.makedirs(os.path.dirname(os.path.abspath(out_xlsx)), exist_ok=True)
    """
    Builds the Stagnant Inventory & Handover Penalty Report.
    Rules:
      - Excused Green Statuses (NO penalty): 420 (Reschedule/Waiting), 472 (Problem resolving)
      - Age < 1 day: Safe ($0)
      - Age 1-2 days: Fine -$0.10/bill
      - Age > 3 days: Max Fine -$0.40/bill
    """
    df = pd.read_excel(src_xlsx)
    df.columns = [str(c).strip().upper() for c in df.columns]

    col_order = next((c for c in df.columns if 'ORDER ID' in c or 'ORDER' in c), 'ORDER ID')
    col_curr_po = next((c for c in df.columns if 'CURRENT POST OFFICE' in c), 'CURRENT POST OFFICE')
    col_deliv_po = next((c for c in df.columns if 'DELIVERY POST OFFICE' in c), 'DELIVERY POST OFFICE')
    col_status = next((c for c in df.columns if 'CURRENT STATUS' in c or 'STATUS' in c), 'CURRENT STATUS')
    col_created = next((c for c in df.columns if 'CREATED DATE' in c), 'CREATED DATE')
    col_action_time = next((c for c in df.columns if 'ACTION TIME' in c or 'CURRENT TIME' in c), 'CURRENT TIME')

    today = datetime.now().date()
    tgt = target_label.upper().replace(" ", "")

    df['sc'] = df[col_status].astype(str).str.extract(r'^(\d{3})')[0]
    df['curr_po_clean'] = df[col_curr_po].astype(str).str.strip().str.upper()
    df['deliv_po_clean'] = df[col_deliv_po].astype(str).str.strip().str.upper()

    # Determine assigned branch for each row (delivery uses delivery post office, handover uses current post office)
    df['branch'] = df.apply(
        lambda r: r['deliv_po_clean'] if str(r['sc']).startswith('4') and r['deliv_po_clean'] != 'NAN' and r['deliv_po_clean'] else r['curr_po_clean'],
        axis=1
    )

    # Exclude completed (410), returned (520), cancelled (201, 99, 100, -99)
    active_df = df[~df['sc'].isin(['410', '520', '201', '99', '100', '-99'])].copy()

    # Filter target branch / zone / all
    if tgt not in ("ALL", "TOTAL"):
        if tgt.startswith("ZONE"):
            zone_by_prefix = {
                "KAN": "ZONE1", "PNP": "ZONE1", "PRE": "ZONE1", "SVA": "ZONE1",
                "KAM": "ZONE2", "KOH": "ZONE2", "SIH": "ZONE2", "SPE": "ZONE2", "TAK": "ZONE2", "KEP": "ZONE2",
                "BAN": "ZONE3", "BAT": "ZONE3", "CHH": "ZONE3", "PUR": "ZONE3", "PAI": "ZONE3",
                "ODD": "ZONE4", "PRH": "ZONE4", "SIE": "ZONE4", "THO": "ZONE4",
                "CHA": "ZONE5", "KRA": "ZONE5", "TBK": "ZONE5", "ROT": "ZONE5", "MON": "ZONE5", "STU": "ZONE5"
            }
            active_df['zone'] = active_df['branch'].str[:3].map(zone_by_prefix).fillna("ZONE1")
            active_df = active_df[active_df['zone'] == tgt].copy()
        elif len(tgt) == 3:
            active_df = active_df[active_df['branch'].str.startswith(tgt)].copy()
        else:
            active_df = active_df[active_df['branch'] == tgt].copy()

    excused_statuses = {"420", "472"}

    # Aggregate stats per branch
    branch_stats = {}
    detail_rows = []

    for _, row in active_df.iterrows():
        po = str(row.get('branch', 'OTHER')).strip()
        if not po or po == 'NAN':
            continue

        if po not in branch_stats:
            branch_stats[po] = {
                "po": po,
                "total_handover": 0,
                "total_delivery": 0,
                "penalty_handover_count": 0,
                "penalty_delivery_count": 0,
                "excused_delivery_count": 0,
                "total_penalty_usd": 0.0
            }

        order_id = str(row.get(col_order, '')).strip()
        sc = str(row.get('sc', '')).strip()
        status_raw = str(row.get(col_status, '')).strip()

        act_val = row.get(col_action_time) if pd.notna(row.get(col_action_time)) else row.get(col_created)
        act_date = parse_date(act_val) or parse_date(row.get(col_created))
        age_days = (today - act_date).days if act_date else 0

        is_handover = sc in {"306", "309", "310", "302", "311", "210"}
        is_delivery = sc.startswith("4")

        if is_handover:
            branch_stats[po]["total_handover"] += 1
            fine = 0.0
            risk = "Safe (< 1 day)"
            if age_days >= 3:
                fine = 0.40
                risk = "Urgent (> 3 days)"
                branch_stats[po]["penalty_handover_count"] += 1
            elif age_days >= 1:
                fine = 0.10
                risk = "Backlog (1-2 days)"
                branch_stats[po]["penalty_handover_count"] += 1

            branch_stats[po]["total_penalty_usd"] += fine
            detail_rows.append({
                "order_id": order_id,
                "po": po,
                "type": "Handover / Transit",
                "status": status_raw,
                "status_code": sc,
                "act_date": str(act_date or ""),
                "age_days": age_days,
                "is_excused": False,
                "risk": risk,
                "fine_usd": fine
            })

        elif is_delivery:
            branch_stats[po]["total_delivery"] += 1
            fine = 0.0
            is_excused = sc in excused_statuses
            risk = "Normal"

            if is_excused:
                branch_stats[po]["excused_delivery_count"] += 1
                risk = f"Excused ({sc})"
                fine = 0.0
            else:
                if age_days >= 3:
                    fine = 0.40
                    risk = "Critical (> 3 days)"
                    branch_stats[po]["penalty_delivery_count"] += 1
                elif age_days >= 1:
                    fine = 0.10
                    risk = "Stagnant (1-2 days)"
                    branch_stats[po]["penalty_delivery_count"] += 1
                else:
                    risk = "Safe (< 1 day)"

            branch_stats[po]["total_penalty_usd"] += fine
            detail_rows.append({
                "order_id": order_id,
                "po": po,
                "type": "Delivery",
                "status": status_raw,
                "status_code": sc,
                "act_date": str(act_date or ""),
                "age_days": age_days,
                "is_excused": is_excused,
                "risk": risk,
                "fine_usd": fine
            })

    # Sort branches alphabetically
    sorted_branches = sorted(branch_stats.values(), key=lambda x: x["po"])

    # Build Excel
    wb = openpyxl.Workbook()
    fn = "Segoe UI"
    banner_font = Font(name=fn, size=11, bold=True, color="FFFFFF")
    hdr_font = Font(name=fn, size=10, bold=True, color="FFFFFF")
    data_font = Font(name=fn, size=10, color="0F172A")
    bold_data_font = Font(name=fn, size=10, bold=True, color="0F172A")
    tot_font = Font(name=fn, size=10, bold=True, color="B91C1C")

    hdr_fill = PatternFill("solid", fgColor="0B132B")      # Dark Navy
    sub_hdr_fill = PatternFill("solid", fgColor="1E3A8A")  # Royal Blue
    tot_fill = PatternFill("solid", fgColor="FEE2E2")      # Light Red
    green_fill = PatternFill("solid", fgColor="DCFCE7")    # Light Green
    amber_fill = PatternFill("solid", fgColor="FEF3C7")    # Light Amber
    red_fill = PatternFill("solid", fgColor="FEE2E2")      # Light Red
    zebra_fill = PatternFill("solid", fgColor="F8FAFC")

    border_thin = Border(
        left=Side(style="thin", color="CBD5E1"), right=Side(style="thin", color="CBD5E1"),
        top=Side(style="thin", color="CBD5E1"), bottom=Side(style="thin", color="CBD5E1")
    )

    # 1. SUMMARY SHEET
    ws_sum = wb.active
    ws_sum.title = "SUMMARY"
    ws_sum.views.sheetView[0].showGridLines = True

    ws_sum.merge_cells("A1:G1")
    ws_sum["A1"].value = f"METFONE EXPRESS — INVENTORY PENALTY & STAGNANT GOODS REPORT ({target_label.upper()})"
    ws_sum["A1"].font = banner_font
    ws_sum["A1"].fill = hdr_fill
    ws_sum["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws_sum.row_dimensions[1].height = 28

    ws_sum.merge_cells("A2:G2")
    ws_sum["A2"].value = f"Generated at: {datetime.now().strftime('%d/%m/%Y %H:%M')}  |  Excused Green: 420, 472 ($0 Fine)  |  SLA Penalty: 1-2d (-$0.10), >3d (-$0.40)"
    ws_sum["A2"].font = Font(name=fn, size=9, italic=True, color="FFFFFF")
    ws_sum["A2"].fill = sub_hdr_fill
    ws_sum["A2"].alignment = Alignment(horizontal="center", vertical="center")
    ws_sum.row_dimensions[2].height = 20

    sum_headers = [
        "Post Office",
        "Total Handover",
        "Total Delivery",
        "Penalty Handover (Count)",
        "Penalty Delivery (Count)",
        "Excused Green (Count)",
        "Total Penalty ($)"
    ]

    ws_sum.row_dimensions[4].height = 26
    for col_idx, h in enumerate(sum_headers, 1):
        c = ws_sum.cell(row=4, column=col_idx, value=h)
        c.font = hdr_font
        c.fill = sub_hdr_fill
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = border_thin

    tot_ho = 0
    tot_del = 0
    tot_pen_ho = 0
    tot_pen_del = 0
    tot_exc = 0
    tot_usd = 0.0

    r_cur = 5
    for b in sorted_branches:
        ws_sum.row_dimensions[r_cur].height = 22
        vals = [
            b["po"],
            b["total_handover"],
            b["total_delivery"],
            b["penalty_handover_count"],
            b["penalty_delivery_count"],
            b["excused_delivery_count"],
            f"-${b['total_penalty_usd']:.2f}" if b["total_penalty_usd"] > 0 else "$0.00"
        ]
        for col_idx, val in enumerate(vals, 1):
            c = ws_sum.cell(row=r_cur, column=col_idx, value=val)
            c.font = bold_data_font if col_idx in (1, 7) else data_font
            c.alignment = Alignment(horizontal="center", vertical="center")
            c.border = border_thin
            if r_cur % 2 == 0:
                c.fill = zebra_fill
            if col_idx == 7 and b["total_penalty_usd"] > 0:
                c.font = tot_font
                c.fill = red_fill

        tot_ho += b["total_handover"]
        tot_del += b["total_delivery"]
        tot_pen_ho += b["penalty_handover_count"]
        tot_pen_del += b["penalty_delivery_count"]
        tot_exc += b["excused_delivery_count"]
        tot_usd += b["total_penalty_usd"]
        r_cur += 1

    # TOTAL ROW
    ws_sum.row_dimensions[r_cur].height = 26
    tot_vals = [
        "TOTAL",
        tot_ho,
        tot_del,
        tot_pen_ho,
        tot_pen_del,
        tot_exc,
        f"-${tot_usd:.2f}" if tot_usd > 0 else "$0.00"
    ]
    for col_idx, val in enumerate(tot_vals, 1):
        c = ws_sum.cell(row=r_cur, column=col_idx, value=val)
        c.font = tot_font
        c.fill = tot_fill
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.border = border_thin

    for col in ws_sum.columns:
        max_len = max(len(str(cell.value or "")) for cell in col)
        col_letter = get_column_letter(col[0].column)
        ws_sum.column_dimensions[col_letter].width = max(max_len + 4, 16)

    # 2. DETAIL SHEET
    ws_det = wb.create_sheet(title="DETAIL_PENALTY")
    ws_det.views.sheetView[0].showGridLines = True

    det_headers = [
        "No", "Order ID", "Post Office", "Type", "Status", "Action Date",
        "Age (Days)", "Excused?", "Risk / SLA Level", "Penalty Fine ($)"
    ]

    ws_det.row_dimensions[1].height = 24
    for col_idx, h in enumerate(det_headers, 1):
        c = ws_det.cell(row=1, column=col_idx, value=h)
        c.font = hdr_font
        c.fill = hdr_fill
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.border = border_thin

    detail_rows.sort(key=lambda x: (x["fine_usd"] == 0, -x["age_days"]))

    for idx, item in enumerate(detail_rows, 1):
        r_num = idx + 1
        ws_det.row_dimensions[r_num].height = 20
        row_data = [
            idx,
            item["order_id"],
            item["po"],
            item["type"],
            item["status"],
            item["act_date"],
            item["age_days"],
            "YES (Green - $0 Fine)" if item["is_excused"] else "NO",
            item["risk"],
            f"-${item['fine_usd']:.2f}" if item["fine_usd"] > 0 else "$0.00"
        ]
        for col_idx, val in enumerate(row_data, 1):
            c = ws_det.cell(row=r_num, column=col_idx, value=val)
            c.font = data_font
            c.alignment = Alignment(horizontal="center", vertical="center")
            c.border = border_thin
            if item["is_excused"]:
                c.fill = green_fill
            elif item["fine_usd"] > 0:
                c.fill = red_fill if item["fine_usd"] >= 0.40 else amber_fill

    for col in ws_det.columns:
        max_len = max(len(str(cell.value or "")) for cell in col)
        col_letter = get_column_letter(col[0].column)
        ws_det.column_dimensions[col_letter].width = max(max_len + 4, 13)

    wb.save(out_xlsx)
    return tot_ho, tot_del, (tot_pen_ho + tot_pen_del), tot_usd
