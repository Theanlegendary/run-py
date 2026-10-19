"""
truck_kpi_report.py — Truck & Route On-Time Connection Quality KPI Report
=========================================================================
Parses Metfone TMS 'Báo cáo tỷ lệ kết nối đúng đủ giờ' (BCtyleketnoidungdugio_*.xlsx)
and generates executive visual dashboards & Telegram summaries matching the exact
management format:
  - Hành trình (Trip Type: Nội tỉnh, Liên vùng, Nội vùng)
  - Tên tuyến (Route name)
  - Biển số xe (Truck plate)
  - Chiều (Direction)
  - TỔNG CHUYẾN (Total trips)
  - ĐÚNG GIỜ (%) (On-time rate)
  - TRỄ GIỜ (%) (Delay rate)
  - TRỄ < 30p (Delayed < 30 mins)
  - TRỄ > 30p (Delayed > 30 mins)
"""

import os
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from collections import defaultdict
from datetime import datetime
import pandas as pd
import excel_to_image


def classify_trip_type(row) -> str:
    """
    Classify route into Trip Type (Hành trình):
      1. Nội tỉnh (Intra-province / Local route): starts with NT-
      2. Liên vùng (Inter-regional / Linehaul): starts with LV
      3. Nội vùng (Intra-regional): starts with NV-
    """
    raw_ht = str(row.get('Hành trình', '')).strip()
    route = str(row.get('Tên tuyến', '')).strip().upper()

    if route.startswith('LV'):
        return 'Liên vùng'
    elif route.startswith('NV'):
        return 'Nội vùng'
    elif route.startswith('NT'):
        return 'Nội tỉnh'

    # Check raw string
    raw_lower = raw_ht.lower()
    if 'liên vùng' in raw_lower or 'lien vung' in raw_lower:
        return 'Liên vùng'
    elif 'nội vùng' in raw_lower or 'noi vung' in raw_lower:
        return 'Nội vùng'
    elif 'nội tỉnh' in raw_lower or 'noi tinh' in raw_lower:
        return 'Nội tỉnh'

    return 'Nội tỉnh' if not raw_ht or raw_ht.lower() == 'nan' else raw_ht


def parse_truck_kpi(source_path: str, exclude_generated: bool = True) -> dict:
    """
    Parse BCtyleketnoidungdugio Excel or CSV file.
    Returns structured stats dict grouped by Hành trình, Tên tuyến, Biển số xe, Chiều.
    If exclude_generated is True, excludes 'Chuyến phát sinh' (Generated route) from Column G.
    """
    if str(source_path).lower().endswith(".csv"):
        df = pd.read_csv(source_path)
    else:
        df = pd.read_excel(source_path, sheet_name=0)
        xl = pd.ExcelFile(source_path)
        for s in xl.sheet_names:
            if "đủ đúng giờ" in s.lower() or "dung" in s.lower():
                df = pd.read_excel(source_path, sheet_name=s)
                break

    # Determine date
    date_str = ""
    for col in df.columns:
        if "TG xuất phát" in str(col) or "TG đến" in str(col):
            valid_vals = df[col].dropna()
            if not valid_vals.empty:
                val_str = str(valid_vals.iloc[0])
                if " " in val_str:
                    date_str = val_str.split(" ")[0]
                break

    # ── Exclude "generated route" (Chuyến phát sinh) if requested ──
    # Column G: 'Chuyến phát sinh' -> Exclude any row where it is 'Có' / 'yes' / 'true'
    gen_col = None
    for col in df.columns:
        c_clean = str(col).strip().lower()
        if "phát sinh" in c_clean or "phat sinh" in c_clean or "generated" in c_clean:
            gen_col = col
            break
    if gen_col is None and len(df.columns) > 6:
        # Check column index 6 (Column G)
        col_g = df.columns[6]
        val_sample = df[col_g].dropna().astype(str).str.lower().tolist()
        if any(v in ('có', 'co', 'không', 'khong', 'yes', 'no') for v in val_sample[:20]):
            gen_col = col_g

    excluded_gen_count = 0
    if exclude_generated and gen_col and gen_col in df.columns:
        is_gen = df[gen_col].astype(str).str.strip().str.lower().isin(['có', 'co', 'yes', 'true', '1'])
        excluded_gen_count = int(is_gen.sum())
        df = df[~is_gen].copy()

    # Assign trip type (Hành trình)
    df['Hành trình'] = df.apply(classify_trip_type, axis=1)

    HT_PRIORITY = {'Nội tỉnh': 1, 'Liên vùng': 2, 'Nội vùng': 3}

    groups = []
    trip_types = defaultdict(lambda: {"total": 0, "dung": 0, "sai": 0, "sai_le30": 0, "sai_gt30": 0})
    trucks = defaultdict(lambda: {"total": 0, "dung": 0, "sai": 0, "sai_le30": 0, "sai_gt30": 0})
    routes = defaultdict(lambda: {"total": 0, "dung": 0, "sai": 0, "sai_le30": 0, "sai_gt30": 0})

    # Group by Hành trình, Tên tuyến, Biển số xe, Chiều
    grouped = df.groupby(['Hành trình', 'Tên tuyến', 'Biển số xe', 'Chiều'], sort=False)
    
    # Sort keys by HT_PRIORITY, then route, truck, chieu
    sorted_keys = sorted(
        grouped.groups.keys(),
        key=lambda k: (HT_PRIORITY.get(k[0], 99), str(k[1]), str(k[2]), str(k[3]))
    )

    for key in sorted_keys:
        g = grouped.get_group(key)
        ht, route, truck, chieu = key
        tot = int(len(g))
        dung = int((g['KPI đúng giờ'] == 'Đúng').sum())
        sai = int((g['KPI đúng giờ'] == 'Sai').sum())
        pct_dung = (dung / tot) if tot > 0 else 0.0
        pct_sai = (sai / tot) if tot > 0 else 0.0
        tre_lt_30 = int(((g['KPI đúng giờ'] == 'Sai') & (g['Chênh lệch TG đến (phút)'] <= 30)).sum())
        tre_gt_30 = int(((g['KPI đúng giờ'] == 'Sai') & (g['Chênh lệch TG đến (phút)'] > 30)).sum())

        groups.append({
            'ht': str(ht).strip(),
            'route': str(route).strip(),
            'truck': str(truck).strip(),
            'chieu': str(chieu).strip(),
            'tot': tot,
            'dung': dung,
            'sai': sai,
            'pct_dung': pct_dung,
            'pct_sai': pct_sai,
            'tre_lt_30': tre_lt_30,
            'tre_gt_30': tre_gt_30,
        })

        # Trip type aggregates
        tt = trip_types[str(ht).strip()]
        tt["total"] += tot
        tt["dung"] += dung
        tt["sai"] += sai
        tt["sai_le30"] += tre_lt_30
        tt["sai_gt30"] += tre_gt_30

        # Truck aggregates
        td = trucks[str(truck).strip()]
        td["total"] += tot
        td["dung"] += dung
        td["sai"] += sai
        td["sai_le30"] += tre_lt_30
        td["sai_gt30"] += tre_gt_30

        # Route aggregates
        rd = routes[str(route).strip()]
        rd["total"] += tot
        rd["dung"] += dung
        rd["sai"] += sai
        rd["sai_le30"] += tre_lt_30
        rd["sai_gt30"] += tre_gt_30

    # Overall Totals
    total_trips = int(len(df))
    total_dung = int((df['KPI đúng giờ'] == 'Đúng').sum())
    total_sai = int((df['KPI đúng giờ'] == 'Sai').sum())
    total_le30 = int(((df['KPI đúng giờ'] == 'Sai') & (df['Chênh lệch TG đến (phút)'] <= 30)).sum())
    total_gt30 = int(((df['KPI đúng giờ'] == 'Sai') & (df['Chênh lệch TG đến (phút)'] > 30)).sum())

    pct_dung = (total_dung / total_trips * 100) if total_trips > 0 else 0.0
    pct_sai = (total_sai / total_trips * 100) if total_trips > 0 else 0.0

    return {
        "date_str": date_str,
        "total_trips": total_trips,
        "total_dung": total_dung,
        "total_sai": total_sai,
        "pct_dung": pct_dung,
        "pct_sai": pct_sai,
        "total_le30": total_le30,
        "total_gt30": total_gt30,
        "groups": groups,
        "trip_types": dict(trip_types),
        "trucks": dict(trucks),
        "routes": dict(routes),
        "excluded_gen_count": excluded_gen_count,
        "df_clean": df,
    }


def export_truck_kpi_workbook(stats: dict, out_xlsx: str, source_path: str = None) -> None:
    """
    Renders an Executive Truck KPI Dashboard in Excel matching management template:
      - Sheet 1: BC Ty le Ket noi (Exact Pivot layout matching @Tian_xin_yue with Hành trình column)
      - Sheet 2: Raw Data (Original source trip data)
    """
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "BC Ty le Ket noi"
    ws.views.sheetView[0].showGridLines = True

    # Palette
    HDR_LEFT_BG = "BDD7EE"     # Soft light blue for Hành trình, Tên tuyến, Biển số xe, Chiều
    HDR_RIGHT_BG = "1F3864"    # Dark navy blue for KPI columns
    GT_BG = "BDD7EE"           # Soft light blue for Grand Total row
    BORDER_CLR = "B8CCE4"      # Subtle grid border

    thin_side = Side(style="thin", color=BORDER_CLR)
    cell_border = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)

    fn = "Segoe UI"
    hdr_left_font = Font(name=fn, size=10, bold=True, color="000000")
    hdr_right_font = Font(name=fn, size=10, bold=True, color="FFFFFF")
    data_font = Font(name=fn, size=9.5, color="000000")
    gt_font = Font(name=fn, size=10, bold=True, color="000000")

    headers = [
        ("Hành trình", HDR_LEFT_BG, hdr_left_font, "left"),
        ("Tên tuyến", HDR_LEFT_BG, hdr_left_font, "left"),
        ("Biển số xe", HDR_LEFT_BG, hdr_left_font, "left"),
        ("Chiều", HDR_LEFT_BG, hdr_left_font, "center"),
        ("TỔNG CHUYẾN", HDR_RIGHT_BG, hdr_right_font, "center"),
        ("ĐÚNG GIỜ (%)", HDR_RIGHT_BG, hdr_right_font, "center"),
        ("TRỄ GIỜ (%)", HDR_RIGHT_BG, hdr_right_font, "center"),
        ("TRỄ < 30p", HDR_RIGHT_BG, hdr_right_font, "center"),
        ("TRỄ > 30p", HDR_RIGHT_BG, hdr_right_font, "center"),
    ]

    # Row 1: Title Banner with Current Date & Time Remark
    date_str = stats.get("date_str") or datetime.now().strftime("%d/%m/%Y")
    now_str = datetime.now().strftime("%d/%m/%Y %H:%M")
    title_text = f"🚚 BÁO CÁO TỶ LỆ KẾT NỐI XE TẢI (DVC & TRUCK KPI) — NGÀY: {date_str} (Cập nhật: {now_str})"

    ws.merge_cells("A1:I1")
    t_cell = ws.cell(1, 1, title_text)
    t_cell.fill = PatternFill("solid", fgColor=HDR_RIGHT_BG)
    t_cell.font = Font(name=fn, size=11, bold=True, color="FFFFFF")
    t_cell.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 28

    # Row 2: Table Column Headers
    for col_idx, (h_text, bg, fnt, algn) in enumerate(headers, 1):
        cell = ws.cell(2, col_idx, h_text)
        cell.fill = PatternFill("solid", fgColor=bg)
        cell.font = fnt
        cell.alignment = Alignment(horizontal=algn, vertical="center")
        cell.border = cell_border
    ws.row_dimensions[2].height = 24

    row_idx = 3
    last_ht = None
    last_route = None

    for g in stats["groups"]:
        ht_name = g['ht']
        route_name = g['route']

        ht_display = f"⊟{ht_name}" if ht_name != last_ht else ""
        route_display = f"⊟{route_name}" if (route_name != last_route or ht_name != last_ht) else ""
        truck_display = f"⊟{g['truck']}"

        last_ht = ht_name
        last_route = route_name

        # Col A: Hành trình
        cA = ws.cell(row_idx, 1, ht_display)
        cA.font = data_font; cA.alignment = Alignment(horizontal="left", vertical="center"); cA.border = cell_border

        # Col B: Tên tuyến
        cB = ws.cell(row_idx, 2, route_display)
        cB.font = data_font; cB.alignment = Alignment(horizontal="left", vertical="center"); cB.border = cell_border

        # Col C: Biển số xe
        cC = ws.cell(row_idx, 3, truck_display)
        cC.font = data_font; cC.alignment = Alignment(horizontal="left", vertical="center"); cC.border = cell_border

        # Col D: Chiều
        cD = ws.cell(row_idx, 4, g['chieu'])
        cD.font = data_font; cD.alignment = Alignment(horizontal="center", vertical="center"); cD.border = cell_border

        # Col E: TỔNG CHUYẾN
        cE = ws.cell(row_idx, 5, g['tot'])
        cE.font = data_font; cE.alignment = Alignment(horizontal="center", vertical="center"); cE.border = cell_border
        cE.number_format = "#,##0"

        # Col F: ĐÚNG GIỜ (%)
        cF = ws.cell(row_idx, 6, g['pct_dung'])
        cF.font = data_font; cF.alignment = Alignment(horizontal="center", vertical="center"); cF.border = cell_border
        cF.number_format = "0.0%"

        # Col G: TRỄ GIỜ (%)
        cG = ws.cell(row_idx, 7, g['pct_sai'])
        cG.font = data_font; cG.alignment = Alignment(horizontal="center", vertical="center"); cG.border = cell_border
        cG.number_format = "0.0%"

        # Col H: TRỄ < 30p
        cH = ws.cell(row_idx, 8, g['tre_lt_30'] if g['tre_lt_30'] > 0 else "")
        cH.font = data_font; cH.alignment = Alignment(horizontal="center", vertical="center"); cH.border = cell_border
        if g['tre_lt_30'] > 0:
            cH.number_format = "#,##0"

        # Col I: TRỄ > 30p
        cI = ws.cell(row_idx, 9, g['tre_gt_30'] if g['tre_gt_30'] > 0 else "")
        cI.font = data_font; cI.alignment = Alignment(horizontal="center", vertical="center"); cI.border = cell_border
        if g['tre_gt_30'] > 0:
            cI.number_format = "#,##0"

        ws.row_dimensions[row_idx].height = 20
        row_idx += 1

    # Grand Total row
    tot_all = stats["total_trips"]
    pct_dung_all = (stats["total_dung"] / tot_all) if tot_all > 0 else 0.0
    pct_sai_all = (stats["total_sai"] / tot_all) if tot_all > 0 else 0.0
    tre_lt_all = stats["total_le30"]
    tre_gt_all = stats["total_gt30"]

    gt_cells = [
        (1, "Grand Total", "left", "@"),
        (2, "", "left", None),
        (3, "", "left", None),
        (4, "", "center", None),
        (5, tot_all, "center", "#,##0"),
        (6, pct_dung_all, "center", "0.0%"),
        (7, pct_sai_all, "center", "0.0%"),
        (8, tre_lt_all, "center", "#,##0"),
        (9, tre_gt_all, "center", "#,##0"),
    ]

    for col_idx, val, algn, num_fmt in gt_cells:
        cell = ws.cell(row_idx, col_idx, val)
        cell.fill = PatternFill("solid", fgColor=GT_BG)
        cell.font = gt_font
        cell.alignment = Alignment(horizontal=algn, vertical="center")
        cell.border = cell_border
        if num_fmt:
            cell.number_format = num_fmt
    ws.row_dimensions[row_idx].height = 24

    # Set filter arrows on columns A:D (starts from Row 2 headers)
    ws.auto_filter.ref = f"A2:D{row_idx-1}"

    # Column widths
    ws.column_dimensions["A"].width = 16
    ws.column_dimensions["B"].width = 65
    ws.column_dimensions["C"].width = 16
    ws.column_dimensions["D"].width = 12
    ws.column_dimensions["E"].width = 16
    ws.column_dimensions["F"].width = 16
    ws.column_dimensions["G"].width = 16
    ws.column_dimensions["H"].width = 14
    ws.column_dimensions["I"].width = 14

    # Copy clean filtered raw data sheet
    df_clean = stats.get("df_clean")
    if df_clean is not None and not df_clean.empty:
        try:
            ws_raw = wb.create_sheet(title="Raw Data")
            for c_idx, c_name in enumerate(df_clean.columns, 1):
                cell = ws_raw.cell(1, c_idx, str(c_name))
                cell.font = Font(name="Segoe UI", size=10, bold=True, color="FFFFFF")
                cell.fill = PatternFill("solid", fgColor="1F3864")
                cell.alignment = Alignment(horizontal="center", vertical="center")
            for r_idx, row_vals in enumerate(df_clean.itertuples(index=False), 2):
                for c_idx, val in enumerate(row_vals, 1):
                    ws_raw.cell(r_idx, c_idx, val if pd.notna(val) else "")
        except Exception:
            pass
    elif source_path and os.path.exists(source_path):
        try:
            src_wb = openpyxl.load_workbook(source_path, data_only=True)
            src_sheet = src_wb.active
            for s in src_wb.sheetnames:
                if "đủ đúng giờ" in s.lower() or "dung" in s.lower():
                    src_sheet = src_wb[s]
                    break
            ws_raw = wb.create_sheet(title="Raw Data")
            for r in range(1, src_sheet.max_row + 1):
                for c in range(1, src_sheet.max_column + 1):
                    ws_raw.cell(r, c, src_sheet.cell(r, c).value)
        except Exception:
            pass

    wb.save(out_xlsx)


def build_telegram_caption(stats: dict) -> str:
    """Format clean Telegram markdown caption with trip type breakdown."""
    date_str = stats.get("date_str") or datetime.now().strftime("%d/%m/%Y")
    now_str = datetime.now().strftime("%d/%m/%Y %H:%M")
    tot_trips = stats["total_trips"]
    tot_dung = stats["total_dung"]
    tot_sai = stats["total_sai"]
    pct_dung = stats["pct_dung"]
    pct_sai = stats["pct_sai"]
    tot_le30 = stats["total_le30"]
    tot_gt30 = stats["total_gt30"]
    ex_gen = stats.get("excluded_gen_count", 0)

    lines = [
        f"🚚 *BÁO CÁO TỶ LỆ KẾT NỐI XE TẢI (DVC & TRUCK KPI)*",
        f"📅 *Ngày*: `{date_str}` | ⏰ *Cập nhật*: `{now_str}`",
        f"📊 *Chỉ số Đúng/Sai Giờ theo Tuyến & Loại Hành Trình*",
    ]
    if ex_gen > 0:
        lines.append(f"ℹ️ _(Đã loại trừ `{ex_gen}` chuyến phát sinh / generated route từ Cột G)_\n")
    else:
        lines.append("")

    lines.extend([
        f"• *Tổng số chuyến*: `{tot_trips:,}` chuyến ({len(stats['routes'])} tuyến | {len(stats['trucks'])} xe)",
        f"• 🟢 *Đúng giờ (On-time)*: `{tot_dung:,}` chuyến (`{pct_dung:.1f}%`)",
        f"• 🔴 *Trễ giờ (Delayed)*: `{tot_sai:,}` chuyến (`{pct_sai:.1f}%`)",
        f"    🟡 *Trễ < 30p*: `{tot_le30:,}` | 🔴 *Trễ > 30p*: `{tot_gt30:,}` (vi phạm SLA)",
    ])

    # Trip type breakdown
    if stats.get("trip_types"):
        lines.append("\n🛣️ *Theo Loại Hành Trình (Trip Type):*")
        for ht, hd in stats["trip_types"].items():
            tot_h = hd["total"]
            pct_d = (hd["dung"] / tot_h * 100) if tot_h > 0 else 0.0
            lines.append(f"  • *{ht}*: `{tot_h:,}` chuyến | Đúng: `{pct_d:.1f}%` (>30p: `{hd['sai_gt30']}`)")

    # Best trucks
    lines.append("\n🏆 *Top Xe Đúng Giờ Nhất:*")
    sorted_trucks = sorted(
        stats["trucks"].items(),
        key=lambda x: (x[1]["dung"] / x[1]["total"] if x[1]["total"] > 0 else 0, x[1]["total"]),
        reverse=True
    )
    for t_plate, td in sorted_trucks[:3]:
        tot = td["total"]
        pct = (td["dung"] / tot * 100) if tot > 0 else 0.0
        lines.append(f"  ✅ `{t_plate}`: `{pct:.1f}%` ({td['dung']}/{tot} | >30p: {td['sai_gt30']})")

    lines.append("\n⚠️ *Top Xe Trễ Nhiều (Cần Cải Thiện):*")
    worst_trucks = sorted(
        stats["trucks"].items(),
        key=lambda x: (x[1]["sai_gt30"], x[1]["sai"]),
        reverse=True
    )
    for t_plate, td in worst_trucks[:3]:
        tot = td["total"]
        pct = (td["sai"] / tot * 100) if tot > 0 else 0.0
        lines.append(f"  ❌ `{t_plate}`: `{pct:.1f}%` trễ ({td['sai']}/{tot} | >30p: {td['sai_gt30']})")

    res = "\n".join(lines)
    if len(res) > 1020:
        res = res[:1015] + "..."
    return res


def find_latest_bctyle_file(extra_dirs=None) -> str:
    """
    Find the most recently modified BCtyleketnoidungdugio file across all common locations:
    project root, Downloads/Telegram Desktop, Downloads, Desktop, cache, etc.
    Runs instantaneously by checking specific directories.
    """
    import glob
    home = os.path.expanduser("~")
    here = os.path.dirname(os.path.abspath(__file__))
    search_dirs = [
        here,
        os.path.join(home, "Downloads", "Telegram Desktop"),
        os.path.join(home, "Downloads"),
        os.path.join(home, "Desktop"),
        os.path.join(home, "Desktop", "daily_push"),
        os.path.join(here, "cache"),
    ]
    if extra_dirs:
        for ed in extra_dirs:
            if ed and os.path.isdir(ed) and ed not in search_dirs:
                search_dirs.append(ed)

    found = []
    for d in search_dirs:
        if os.path.isdir(d):
            try:
                for f in glob.glob(os.path.join(d, "BCtyle*.*")):
                    if os.path.isfile(f) and f not in found and not os.path.basename(f).startswith("~$"):
                        found.append(f)
            except Exception:
                pass

    if not found:
        return None
    return sorted(found, key=os.path.getmtime, reverse=True)[0]


def generate_truck_kpi_report(source_path: str = None, out_dir: str = None, exclude_generated: bool = True):
    """
    Main entry point: parses Excel, generates dashboard workbook,
    renders image, and returns paths and Telegram caption.
    If exclude_generated is True, excludes 'Chuyến phát sinh' (Column G = Có).
    """
    if not source_path:
        source_path = find_latest_bctyle_file()
    if not source_path or not os.path.exists(source_path):
        raise FileNotFoundError(f"BCtyle source file not found: {source_path}")

    if not out_dir:
        out_dir = "."
    stats = parse_truck_kpi(source_path, exclude_generated=exclude_generated)
    stamp = datetime.now().strftime("%d.%m_%HH%M")
    out_xlsx = os.path.join(out_dir, f"Truck_KPI_Report_{stamp}.xlsx")
    export_truck_kpi_workbook(stats, out_xlsx, source_path=source_path)

    img_buf = excel_to_image.excel_to_image(out_xlsx)
    caption = build_telegram_caption(stats)

    return out_xlsx, img_buf, caption, stats


if __name__ == "__main__":
    test_src = "BCtyleketnoidungdugio_27092026.xlsx"
    if os.path.exists(test_src):
        out_xlsx, img_buf, caption, stats = generate_truck_kpi_report(test_src, ".")
        with open("truck_kpi_dashboard.png", "wb") as f:
            f.write(img_buf.getbuffer())
        print(f"Generated {out_xlsx} and truck_kpi_dashboard.png successfully!")
        print(caption)
