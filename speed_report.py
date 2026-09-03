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

MAIN_36_BRANCHES = [
    'BANP001', 'BATP001', 'CHAP001', 'CHHP001', 'KAMP001', 'KANP001', 'KOHP001', 'KRAP001',
    'MONP001', 'ODDP001', 'PNPP001', 'PNPP002', 'PNPP003', 'PNPP004', 'PNPP005', 'PNPP006',
    'PNPP007', 'PNPP008', 'PNPP009', 'PNPP010', 'PNPP011', 'PNPP012', 'PNPP013', 'PNPP014',
    'PREP001', 'PRHP001', 'PURP001', 'ROTP001', 'SIEP001', 'SIHP001', 'SPEP001', 'STUP001',
    'SVAP001', 'TAKP001', 'TBKP001', 'THOP001'
]

ZONE_BRANCHES_MAP = {
    "ZONE1": ['PNPP001', 'PNPP002', 'PNPP003', 'PNPP004', 'PNPP005', 'PNPP006', 'PNPP007', 'PNPP008', 'PNPP009', 'PNPP010', 'PNPP011', 'PNPP012', 'PNPP013', 'PNPP014', 'KANP001', 'PREP001', 'SVAP001'],
    "ZONE2": ['KAMP001', 'KOHP001', 'SIHP001', 'SPEP001', 'TAKP001'],
    "ZONE3": ['BANP001', 'BATP001', 'CHHP001', 'PURP001'],
    "ZONE4": ['ODDP001', 'PRHP001', 'SIEP001', 'THOP001'],
    "ZONE5": ['CHAP001', 'KRAP001', 'MONP001', 'ROTP001', 'STUP001', 'TBKP001'],
}

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
    CEO-Grade Fast Delivery Speed & Commission Report:
      - Distinct Theme: Deep Forest Emerald (#064E3B) & Electric Teal/Cyan (#0F766E, #0284C7)
      - Executive Summary with Commission Rates Subtitle & Accounting Double-Borders
      - Sheet 1: Left Table = Detailed Delivered Bills, Right Table = 36-Branch Executive Summary
      - Sheet 2: base = Raw Audit Dataset
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
    
    col_306_last = next((c for c in df.columns if '306' in c and ('STORE' in c or 'AGENT' in c) and 'LAST' in c), None)
    col_306_first = next((c for c in df.columns if '306' in c and ('STORE' in c or 'AGENT' in c or 'HUB' in c)), None)
    col_400_time = next((c for c in df.columns if '400' in c or 'OUT' in c or 'ASSIGN' in c), None)
    col_210_time = next((c for c in df.columns if '210' in c), None)

    today = report_date or datetime.now().date()
    tgt = "".join(c for c in str(target_label).upper() if c.isalnum() or c in ("-", "_")).strip()
    if not tgt:
        tgt = "ALL"

    # Load Revenue to get accurate VAS_SERVICE VTT mapping
    revenue_cache = os.path.join(os.path.dirname(os.path.abspath(__file__)), "cache", "latest_revenue.xlsx")
    vtt_order_ids = set()
    if os.path.exists(revenue_cache):
        try:
            rev_df = pd.read_excel(revenue_cache, skiprows=4)
            rev_df['ORDER_NUMBER'] = rev_df['ORDER_NUMBER'].dropna().astype(str).str.split('.').str[0].str.strip()
            vtt_rev = rev_df[rev_df['VAS_SERVICE'].astype(str).str.contains('VTT', case=False, na=False)]
            vtt_order_ids = set(vtt_rev['ORDER_NUMBER'])
        except Exception as e:
            pass

    df['order_clean'] = df[col_order].astype(str).str.split('.').str[0].str.strip()
    
    # STRICT VTT FILTER: Order must be in VTT revenue list OR have SERVICE/NOTE as VTT
    if vtt_order_ids:
        df = df[
            (df['order_clean'].isin(vtt_order_ids)) |
            (df['SERVICE'].astype(str).str.contains('VTT', case=False, na=False))
        ].copy()

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
    if tgt in ("ALL", "TOTAL"):
        for b in MAIN_36_BRANCHES:
            summary_data[b] = {
                "po": b,
                "total_delivered": 0,
                "under_2h": 0,
                "between_2_4h": 0,
                "between_4_8h": 0,
                "over_8h": 0,
                "total_commission": 0.0
            }
    elif tgt.startswith("ZONE") and tgt in ZONE_BRANCHES_MAP:
        for b in ZONE_BRANCHES_MAP[tgt]:
            summary_data[b] = {
                "po": b,
                "total_delivered": 0,
                "under_2h": 0,
                "between_2_4h": 0,
                "between_4_8h": 0,
                "over_8h": 0,
                "total_commission": 0.0
            }

    base_rows = []
    r_idx = 1

    for idx, row in delivered_df.iterrows():
        deliv_po = str(row.get('deliv_po_clean', '')).strip()
        curr_po = str(row.get('curr_po_clean', '')).strip()
        raw_po = deliv_po if deliv_po and deliv_po != 'NAN' else curr_po

        if not raw_po or raw_po == 'NAN':
            continue

        # STRICT FILTER: 36 Main Post Offices only (Exclude agents & showrooms)
        if tgt in ("ALL", "TOTAL"):
            if raw_po not in MAIN_36_BRANCHES:
                continue
            po = raw_po
        elif tgt.startswith("ZONE") and tgt in ZONE_BRANCHES_MAP:
            if raw_po not in ZONE_BRANCHES_MAP[tgt]:
                continue
            po = raw_po
        elif len(tgt) >= 7:
            po = tgt
            if raw_po != tgt:
                continue
        else:
            if raw_po not in MAIN_36_BRANCHES:
                continue
            po = raw_po

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
        
        t_start = None
        if col_210_time and pd.notna(row.get(col_210_time)):
            t_start = parse_time(row.get(col_210_time))
        elif col_306_last and pd.notna(row.get(col_306_last)):
            t_start = parse_time(row.get(col_306_last))
        elif col_306_first and pd.notna(row.get(col_306_first)):
            t_start = parse_time(row.get(col_306_first))
        elif col_400_time and pd.notna(row.get(col_400_time)):
            t_start = parse_time(row.get(col_400_time))
        else:
            t_start = parse_time(row.get(col_created))

        if t410 and t_start and t410 >= t_start:
            duration_hours = (t410 - t_start).total_seconds() / 3600.0
        else:
            duration_hours = 1.5

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
            "t_start": str(t_start or ""),
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

    # ── CEO MIDNIGHT & ROYAL SAPPHIRE BLUE PALETTE (Exact Match to Executive Dashboard) ──
    fill_title_left   = PatternFill("solid", fgColor="0F172A") # Midnight Navy
    fill_hdr_left     = PatternFill("solid", fgColor="1E293B") # Midnight Slate
    fill_title_right  = PatternFill("solid", fgColor="0B132B") # Deepest Midnight Navy
    fill_sub_right    = PatternFill("solid", fgColor="1C2541") # Subtitle Midnight Slate
    fill_hdr_right    = PatternFill("solid", fgColor="1C3D82") # Royal Sapphire Blue Header
    fill_row_white    = PatternFill("solid", fgColor="FFFFFF") # 100% Pure White (No 2 alternating colors)
    fill_left_tot     = PatternFill("solid", fgColor="E2E8F0") # Soft Slate Total
    fill_sum_tot      = PatternFill("solid", fgColor="E2E8F0") # Executive Accounting Total

    border_clean = Border(
        left=Side(style="thin", color="CBD5E1"), right=Side(style="thin", color="CBD5E1"),
        top=Side(style="thin", color="CBD5E1"), bottom=Side(style="thin", color="CBD5E1")
    )
    tot_border_accounting = Border(
        left=Side(style="thin", color="CBD5E1"), right=Side(style="thin", color="CBD5E1"),
        top=Side(style="thin", color="94A3B8"), bottom=Side(style="double", color="0B132B")
    )

    font_banner = Font(name="Segoe UI", size=10.5, bold=True, color="FFFFFF")
    font_sub = Font(name="Segoe UI", size=8.0, italic=True, color="93C5FD")
    font_hdr = Font(name="Segoe UI", size=8.5, bold=True, color="FFFFFF")
    font_data = Font(name="Segoe UI", size=8.5, color="0F172A")
    font_bold_data = Font(name="Segoe UI", size=8.5, bold=True, color="0F172A")
    font_tot = Font(name="Segoe UI", size=9.5, bold=True, color="0F172A")

    # High-contrast bold font colors for % on-time metrics (matching Penalty Dashboard)
    font_pct_green = Font(name="Segoe UI", size=8.5, bold=True, color="16A34A") # >= 90%
    font_pct_amber = Font(name="Segoe UI", size=8.5, bold=True, color="D97706") # 75% - 89.9%
    font_pct_red   = Font(name="Segoe UI", size=8.5, bold=True, color="DC2626") # < 75%

    font_tot_pct_green = Font(name="Segoe UI", size=9.5, bold=True, color="16A34A")
    font_tot_pct_amber = Font(name="Segoe UI", size=9.5, bold=True, color="D97706")
    font_tot_pct_red   = Font(name="Segoe UI", size=9.5, bold=True, color="DC2626")

    font_comm_green = Font(name="Segoe UI", size=8.5, bold=True, color="047857") # Clean Emerald Green
    font_tot_comm_green = Font(name="Segoe UI", size=9.5, bold=True, color="047857")

    def get_pct_font(pct_val, is_tot=False):
        if pct_val >= 90.0:
            return font_tot_pct_green if is_tot else font_pct_green
        elif pct_val >= 75.0:
            return font_tot_pct_amber if is_tot else font_pct_amber
        else:
            return font_tot_pct_red if is_tot else font_pct_red

    date_str = today.strftime('%d/%m/%Y')

    # 1. Left Title Banner
    ws1.merge_cells("A1:H1")
    ws1.cell(1, 1, f"METFONE EXPRESS — DAILY DELIVERY SPEED DETAIL ({tgt}) — {date_str}").font = font_banner
    ws1.cell(1, 1).alignment = Alignment(horizontal="left", vertical="center")
    for c in range(1, 9):
        ws1.cell(1, c).fill = fill_title_left
    ws1.row_dimensions[1].height = 26.0

    # 2. Right Title Banner
    ws1.merge_cells("J1:S1")
    ws1.cell(1, 10, f"EXECUTIVE DELIVERY SPEED DASHBOARD ({tgt})").font = font_banner
    ws1.cell(1, 10).alignment = Alignment(horizontal="center", vertical="center")
    for c in range(10, 20):
        ws1.cell(1, c).fill = fill_title_right

    ws1.merge_cells("J2:S2")
    ws1.cell(2, 10, "Speed SLA Bonus: < 2h (+50% / $0.30) • 2-4h (+25% / $0.25) • 4-8h ($0.20) • > 8h (-25% / $0.15)").font = font_sub
    ws1.cell(2, 10).alignment = Alignment(horizontal="center", vertical="center")
    for c in range(10, 20):
        ws1.cell(2, c).fill = fill_sub_right
    ws1.row_dimensions[2].height = 18.0

    # Row 3: Headers
    headers_left = [
        "No", "Order Number", "Customer", "Post Office",
        "Duration", "Hours (Dec)", "Speed Category", "Rate ($/Bill)"
    ]
    headers_right = [
        "No", "Post Office", "Delivered (410)", "< 2 Hours (+50%)",
        "2 - 4 Hours (+25%)", "4 - 8 Hours (Normal)", "> 8 Hours (-25%)",
        "% < 8 HOUR", "% > 8 HOUR", "Commission ($)"
    ]

    ws1.row_dimensions[3].height = 26.0
    for ci, h in enumerate(headers_left, 1):
        cell = ws1.cell(3, ci, h)
        cell.font = font_hdr
        cell.fill = fill_hdr_left
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = border_clean

    for ci, h in enumerate(headers_right, 10):
        cell = ws1.cell(3, ci, h)
        cell.font = font_hdr
        cell.fill = fill_hdr_right
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = border_clean

    # Populate Left Detail Order Rows
    r_curr = 4
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
            c.fill = fill_row_white

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
            cell.font = font_tot_comm_green
            cell.alignment = Alignment(horizontal="center", vertical="center")

    # Populate Right Executive Summary Table
    # Sort: Active branches first, then % < 8h ascending (worst SLA on top for CEO review), then deliveries descending
    def calc_speed_sort_key(stats):
        tot_b = stats["total_delivered"]
        on_time_b = stats["under_2h"] + stats["between_2_4h"] + stats["between_4_8h"]
        pct_u8 = (on_time_b / tot_b * 100.0) if tot_b > 0 else 100.0
        return (0 if tot_b > 0 else 1, pct_u8, -tot_b)

    if tgt in ("ALL", "TOTAL"):
        all_branches = [summary_data[b] for b in MAIN_36_BRANCHES if b in summary_data]
    elif tgt.startswith("ZONE") and tgt in ZONE_BRANCHES_MAP:
        all_branches = [summary_data[b] for b in ZONE_BRANCHES_MAP[tgt] if b in summary_data]
    else:
        all_branches = list(summary_data.values())

    sorted_branches = sorted(all_branches, key=calc_speed_sort_key)

    r_sum = 4
    n_idx = 1
    tot_del = 0
    tot_u2 = 0
    tot_24 = 0
    tot_48 = 0
    tot_o8 = 0
    tot_pay = 0.0

    for stats in sorted_branches:
        ws1.row_dimensions[r_sum].height = 18.0
        
        tot_b = stats["total_delivered"]
        on_time_b = stats["under_2h"] + stats["between_2_4h"] + stats["between_4_8h"]
        pct_under_8h = (on_time_b / tot_b * 100.0) if tot_b > 0 else 100.0
        pct_over_8h = (stats["over_8h"] / tot_b * 100.0) if tot_b > 0 else 0.0
        
        s_vals = [
            n_idx,
            stats["po"],
            stats["total_delivered"],
            stats["under_2h"],
            stats["between_2_4h"],
            stats["between_4_8h"],
            stats["over_8h"],
            f"{pct_under_8h:.1f}%",
            f"{pct_over_8h:.1f}%",
            f"${stats['total_commission']:.2f}"
        ]
        for ci, val in enumerate(s_vals, 10):
            cell = ws1.cell(r_sum, ci, val)
            cell.font = font_bold_data if ci in (11, 19) else font_data
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border = border_clean
            cell.fill = fill_row_white
                
            # Clean styling - NO red fills on counts, clean normal cells
            if ci == 17:
                cell.font = get_pct_font(pct_under_8h)
            elif ci == 18:
                cell.font = font_pct_red if stats["over_8h"] > 0 else font_data
            elif ci == 19:
                cell.font = font_comm_green

        tot_del += stats["total_delivered"]
        tot_u2 += stats["under_2h"]
        tot_24 += stats["between_2_4h"]
        tot_48 += stats["between_4_8h"]
        tot_o8 += stats["over_8h"]
        tot_pay += stats["total_commission"]
        r_sum += 1
        n_idx += 1

    # Right Grand Total Row
    ws1.row_dimensions[r_sum].height = 24.0
    ws1.merge_cells(start_row=r_sum, start_column=10, end_row=r_sum, end_column=11)
    rt_tot = ws1.cell(r_sum, 10, "Grand Total")
    rt_tot.font = font_tot
    rt_tot.alignment = Alignment(horizontal="left", vertical="center")

    tot_on_time = tot_u2 + tot_24 + tot_48
    tot_pct_u8 = (tot_on_time / tot_del * 100.0) if tot_del > 0 else 100.0
    tot_pct_o8 = (tot_o8 / tot_del * 100.0) if tot_del > 0 else 0.0

    tot_vals_right = [
        "", "", tot_del, tot_u2, tot_24, tot_48, tot_o8,
        f"{tot_pct_u8:.1f}%", f"{tot_pct_o8:.1f}%", f"${tot_pay:.2f}"
    ]
    for c in range(10, 20):
        cell = ws1.cell(r_sum, c)
        if c >= 12:
            cell.value = tot_vals_right[c-10]
        if c == 17:
            cell.font = get_pct_font(tot_pct_u8, is_tot=True)
        elif c == 18:
            cell.font = font_tot_pct_red if tot_o8 > 0 else font_tot
        elif c == 19:
            cell.font = font_tot_comm_green
        else:
            cell.font = font_tot
        cell.fill = fill_sum_tot
        cell.border = tot_border_accounting
        cell.alignment = Alignment(horizontal="center", vertical="center")

    col_widths = {
        1: 5, 2: 15, 3: 20, 4: 12, 5: 12, 6: 12, 7: 18, 8: 14,
        9: 4,
        10: 5, 11: 12, 12: 13, 13: 14, 14: 14, 15: 14, 16: 14, 17: 14, 18: 14, 19: 15
    }
    for c, w in col_widths.items():
        ws1.column_dimensions[get_column_letter(c)].width = w

    # Sheet 2: base
    ws2 = wb.create_sheet(title="base")
    ws2.views.sheetView[0].showGridLines = True

    base_headers = [
        "No", "Order Number", "Customer", "Origin Branch", "Origin Post",
        "Destination Branch", "Destination Post", "Assigned Branch", "Created At",
        "Branch Arrival (306/400)", "Delivered (410)", "Duration", "Hours (Dec)",
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
            item["t_start"],
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

    for col in ws2.columns:
        max_len = max(len(str(cell.value or "")) for cell in col)
        col_letter = get_column_letter(col[0].column)
        ws2.column_dimensions[col_letter].width = max(max_len + 3, 11)

    wb.save(out_xlsx)
    return tot_del, tot_u2, tot_24, tot_o8, tot_pay


def render_speed_summary_image(out_xlsx):
    wb = openpyxl.load_workbook(out_xlsx)
    ws = wb['DELIVERY SPEED REPORT']

    wb_sum = openpyxl.Workbook()
    ws_sum = wb_sum.active
    ws_sum.title = 'Executive Summary'
    ws_sum.views.sheetView[0].showGridLines = True

    max_r = 1
    for r in range(1, ws.max_row + 1):
        if ws.cell(r, 10).value is not None or ws.cell(r, 19).value is not None:
            max_r = r

    for r in range(1, max_r + 1):
        if ws.row_dimensions[r].height:
            ws_sum.row_dimensions[r].height = ws.row_dimensions[r].height
        for c_idx in range(10):
            orig_c = 10 + c_idx
            tgt_c = 1 + c_idx
            cell_orig = ws.cell(r, orig_c)
            cell_tgt = ws_sum.cell(r, tgt_c, cell_orig.value)
            if cell_orig.has_style:
                cell_tgt.font = copy.copy(cell_orig.font)
                cell_tgt.fill = copy.copy(cell_orig.fill)
                cell_tgt.border = copy.copy(cell_orig.border)
                cell_tgt.alignment = copy.copy(cell_orig.alignment)

    col_widths = {1: 6, 2: 14, 3: 15, 4: 16, 5: 16, 6: 16, 7: 16, 8: 14, 9: 14, 10: 18}
    for c, w in col_widths.items():
        ws_sum.column_dimensions[get_column_letter(c)].width = w

    ws_sum.merge_cells("A1:J1")
    ws_sum.merge_cells("A2:J2")
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
