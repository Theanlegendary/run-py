# -*- coding: utf-8 -*-
"""
generate_tracking_log.py
Generates full Tracking Status Logs Report for All Bills dynamically.
Downloads wide-range data from the previous month to today, filters by today's month,
and builds the 229-column Excel report.
"""

import os
import sys
import json
import sqlite3
import shutil
import requests
import pandas as pd
from datetime import datetime, timedelta
from concurrent.futures import ThreadPoolExecutor, as_completed
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.utils import get_column_letter

HERE = os.path.dirname(os.path.abspath(__file__))
CONFIG_PATH = os.path.join(HERE, "config.json")
CACHE_DB_PATH = os.path.join(HERE, "cache", "tracking_cache_bot.db")
TERMINAL_STATUSES = {"410", "520", "201", "600", "540"}

FULL_STATUS_CODES_TEMPLATE = [
    "110", "120", "130", "140", "150",
    "200", "201", "202", "205", "210", "220", "230", "240", "250",
    "300", "302", "304", "306", "308", "309", "310", "311", "312", "315", "320", "330",
    "400", "401", "402", "403", "404", "405", "406", "408", "415", "417", "420", "421", "422", "425", "428",
    "430", "431", "432", "440", "450", "460", "470", "471", "472", "475", "480", "485", "490", "495",
    "500", "501", "502", "505", "510", "512", "515", "520", "525", "530", "540", "550", "560", "570", "580", "590",
    "600", "410", "99"
]

def init_cache_db():
    os.makedirs(os.path.dirname(CACHE_DB_PATH), exist_ok=True)
    conn = sqlite3.connect(CACHE_DB_PATH)
    conn.execute("""
        CREATE TABLE IF NOT EXISTS tracking_cache (
            order_id TEXT PRIMARY KEY,
            data_json TEXT
        )
    """)
    conn.commit()
    conn.close()

def load_cached_orders():
    init_cache_db()
    conn = sqlite3.connect(CACHE_DB_PATH)
    cached = {}
    for r in conn.execute("SELECT order_id, data_json FROM tracking_cache").fetchall():
        try:
            cached[r[0]] = json.loads(r[1])
        except Exception:
            pass
    conn.close()
    return cached

def save_cache_batch(items):
    if not items:
        return
    conn = sqlite3.connect(CACHE_DB_PATH)
    conn.executemany("INSERT OR REPLACE INTO tracking_cache (order_id, data_json) VALUES (?, ?)", items)
    conn.commit()
    conn.close()

def _has_khmer(text):
    if not isinstance(text, str):
        return False
    return any(0x1780 <= ord(c) <= 0x17FF for c in text)

def _font(text="", is_bold=False, size=11, color="000000"):
    f_name = "Khmer UI" if _has_khmer(text) else "Calibri"
    return Font(name=f_name, size=size, bold=is_bold, color=color)

def generate_tracking_log_report(progress_callback=None):
    if progress_callback:
        progress_callback("Reading config...")
        
    with open(CONFIG_PATH, "r", encoding="utf-8") as f:
        cfg = json.load(f)

    # Determine date range dynamically
    now = datetime.now()
    yesterday = now - timedelta(days=1)
    first_of_this_month = yesterday.replace(day=1)
    last_day_of_prev_month = first_of_this_month - timedelta(days=1)
    first_of_prev_month = last_day_of_prev_month.replace(day=1)

    str_from_date = first_of_prev_month.strftime("%Y%m%d")
    str_to_date = yesterday.strftime("%Y%m%d")

    # API request
    if progress_callback:
        progress_callback(f"Downloading base data from API ({str_from_date} - {str_to_date})...\nThis may take 1-2 minutes.")
        
    payload = {
        'from_date': str_from_date,
        'to_date': str_to_date,
        'branch_code': cfg['api']['branch_code']
    }
    r_all = requests.post(cfg['api']['url'], headers={
        'Authorization': 'Bearer ' + cfg['api']['bearer_token'],
        'Referer': 'https://opsexpress.metfone.com.kh/',
        'Accept-Language': 'vi-VN',
        'Accept': 'application/json, text/plain, */*',
        'Content-Type': 'application/json',
        'x-client-id': 'TMS_ANDROID',
        'User-Agent': 'Mozilla/5.0'
    }, json=payload, timeout=300)

    if r_all.status_code != 200:
        raise Exception(f"Failed to download wide range data: HTTP {r_all.status_code}")

    detail_file = os.path.join(HERE, "tmp_wide_range_for_tracking.xlsx")
    with open(detail_file, 'wb') as f:
        f.write(r_all.content)

    if progress_callback:
        progress_callback("Processing downloaded data...")

    df = pd.read_excel(detail_file)
    df.columns = [str(c).strip().upper() for c in df.columns]

    time_col = next((c for c in df.columns if "CURRENT TIME" in c), None)
    created_col = next((c for c in df.columns if "CREATED DATE" in c), None)
    
    if not time_col:
        raise Exception("CURRENT TIME column not found!")

    df["_parsed_cur_time"] = pd.to_datetime(df[time_col], dayfirst=True, errors='coerce')
    df["_parsed_created"]  = pd.to_datetime(df[created_col], dayfirst=True, errors='coerce') if created_col else None

    sep_start = pd.Timestamp(first_of_this_month.strftime("%Y-%m-%d 00:00:00"))
    sep_end   = pd.Timestamp(yesterday.strftime("%Y-%m-%d 23:59:59"))

    mask = (df["_parsed_cur_time"] >= sep_start) & (df["_parsed_cur_time"] <= sep_end)
    if created_col:
        mask = mask | ((df["_parsed_created"] >= sep_start) & (df["_parsed_created"] <= sep_end))

    df_filtered = df[mask].copy()
    
    if progress_callback:
        progress_callback(f"Found {len(df_filtered)} bills for current month. Preparing tracking data...")

    df_filtered['ORDER ID'] = df_filtered['ORDER ID'].astype(str).str.strip()

    service_map = {}
    svc_col = next((c for c in df_filtered.columns if c in ["SERVICE", "SERVICE TYPE", "SERVICE_TYPE", "SERVICE NAME", "SERVICETYPE"]), None)
    if svc_col:
        for _, r in df_filtered.iterrows():
            oid = str(r.get("ORDER ID", "")).strip()
            val = str(r.get(svc_col, "") or "").strip()
            if oid and val and val.lower() != "nan":
                service_map[oid] = val

    action_user_map = {}
    au_col = next((c for c in df_filtered.columns if c in ["ACTION USER", "ACTION_USER", "LAST ACTION USER", "LAST USER", "USER"]), None)
    if au_col:
        for _, r in df_filtered.iterrows():
            oid = str(r.get("ORDER ID", "")).strip()
            val = str(r.get(au_col, "") or "").strip()
            if oid and val and val.lower() != "nan":
                action_user_map[oid] = val

    vas_fee_map = {}
    vf_col = next((c for c in df_filtered.columns if "VAS FEE" in c), None)
    if vf_col:
        for _, r in df_filtered.iterrows():
            oid = str(r.get("ORDER ID", "")).strip()
            val = r.get(vf_col, 0)
            try:
                if float(val) > 0:
                    vas_fee_map[oid] = float(val)
            except Exception:
                pass

    current_status_map = {}
    st_col = next((c for c in df_filtered.columns if "CURRENT STATUS" in c), None)
    if st_col:
        for _, r in df_filtered.iterrows():
            oid = str(r.get("ORDER ID", "")).strip()
            val = str(r.get(st_col, "") or "").strip()
            if oid and val and val.lower() != "nan":
                current_status_map[oid] = val

    current_po_map = {}
    po_col = next((c for c in df_filtered.columns if "CURRENT POST OFFICE" in c), None)
    if po_col:
        for _, r in df_filtered.iterrows():
            oid = str(r.get("ORDER ID", "")).strip()
            val = str(r.get(po_col, "") or "").strip()
            if oid and val and val.lower() != "nan":
                current_po_map[oid] = val

    current_time_map = {}
    if time_col:
        for _, r in df_filtered.iterrows():
            oid = str(r.get("ORDER ID", "")).strip()
            val = r.get(time_col)
            if oid and pd.notna(val):
                try:
                    current_time_map[oid] = val.strftime("%d/%m/%Y %H:%M:%S")
                except Exception:
                    current_time_map[oid] = str(val)

    cached_results = load_cached_orders()
    unique_oids = df_filtered["ORDER ID"].dropna().astype(str).unique()
    to_fetch = [oid for oid in unique_oids if oid not in cached_results]

    if progress_callback and to_fetch:
        progress_callback(f"Need to fetch API for {len(to_fetch)} new bills...")

    headers_api = {
        "Authorization": "Bearer " + cfg["api"]["bearer_token"],
        "Accept": "application/json, text/plain, */*",
        "User-Agent": "Mozilla/5.0"
    }
    session = requests.Session()
    adapter = HTTPAdapter(pool_connections=100, pool_maxsize=100, max_retries=Retry(total=2, backoff_factor=0.1))
    session.mount("https://", adapter)
    session.headers.update(headers_api)

    def fetch_bill(oid_str):
        if not oid_str or oid_str.lower() == "nan":
            return None

        svc_type  = service_map.get(oid_str, "")
        raw_st    = current_status_map.get(oid_str, "")
        def_st    = raw_st.split("-")[0].strip() if raw_st else ""
        def_unit  = current_po_map.get(oid_str, "")
        def_time  = current_time_map.get(oid_str, "")
        def_user  = action_user_map.get(oid_str, "")

        row_dict = {
            "BILL ID": oid_str, "SERVICE TYPE": svc_type, "VAS": "",
            "_trips_map": {},
            "LATEST_STATUS": def_st, "LATEST_UNIT": def_unit,
            "LATEST_TIME": def_time, "LATEST_USER": def_user
        }

        trips = []
        data_tr = None
        try:
            r_tr = session.get(
                "https://gw-express.metfone.com.kh/tms-tracking/api/v1/order-tracking",
                params={"order_id": oid_str}, timeout=12
            )
            if r_tr.status_code == 200:
                data_tr = r_tr.json()
                trips = data_tr.get("trackingTrips", [])
        except Exception:
            pass

        vas_val = ""
        try:
            if data_tr and "orderDetail" in data_tr:
                od = data_tr["orderDetail"]
                vas_obj = od.get("orderVas", {})
                if vas_obj:
                    vases = []
                    for k, v in vas_obj.items():
                        if isinstance(v, (int, float)) and v > 0:
                            vases.append(f"{k}:{v}")
                    if vases:
                        vas_val = ", ".join(vases)
        except Exception:
            pass
        if not vas_val and oid_str in vas_fee_map:
            vas_val = f"VAS_FEE:{vas_fee_map[oid_str]}"
        row_dict["VAS"] = vas_val

        for trip in sorted(trips, key=lambda x: str(x.get("time", ""))):
            sc = str(trip.get("status", "")).strip()
            unit = str(trip.get("postOfficeCode", "")).strip()
            time_str = str(trip.get("time", "")).strip()
            if sc and sc in set(FULL_STATUS_CODES_TEMPLATE):
                row_dict["_trips_map"][sc] = (unit, time_str)

        return (oid_str, data_tr, row_dict)

    all_results = []
    
    for oid_str in unique_oids:
        if oid_str in cached_results:
            data_tr = cached_results[oid_str]
            svc_type  = service_map.get(oid_str, "")
            raw_st    = current_status_map.get(oid_str, "")
            def_st    = raw_st.split("-")[0].strip() if raw_st else ""
            def_unit  = current_po_map.get(oid_str, "")
            def_time  = current_time_map.get(oid_str, "")
            def_user  = action_user_map.get(oid_str, "")

            row_dict = {
                "BILL ID": oid_str, "SERVICE TYPE": svc_type, "VAS": "",
                "_trips_map": {},
                "LATEST_STATUS": def_st, "LATEST_UNIT": def_unit,
                "LATEST_TIME": def_time, "LATEST_USER": def_user
            }

            vas_val = ""
            try:
                if data_tr and "orderDetail" in data_tr:
                    od = data_tr["orderDetail"]
                    vas_obj = od.get("orderVas", {})
                    if vas_obj:
                        vases = []
                        for k, v in vas_obj.items():
                            if isinstance(v, (int, float)) and v > 0:
                                vases.append(f"{k}:{v}")
                        if vases:
                            vas_val = ", ".join(vases)
            except Exception:
                pass
            if not vas_val and oid_str in vas_fee_map:
                vas_val = f"VAS_FEE:{vas_fee_map[oid_str]}"
            row_dict["VAS"] = vas_val

            trips = data_tr.get("trackingTrips", []) if data_tr else []
            for trip in sorted(trips, key=lambda x: str(x.get("time", ""))):
                sc = str(trip.get("status", "")).strip()
                unit = str(trip.get("postOfficeCode", "")).strip()
                time_str = str(trip.get("time", "")).strip()
                if sc and sc in set(FULL_STATUS_CODES_TEMPLATE):
                    row_dict["_trips_map"][sc] = (unit, time_str)
                    
            all_results.append(row_dict)

    if to_fetch:
        batch_to_save = []
        fetched_count = 0
        total_fetch = len(to_fetch)
        
        with ThreadPoolExecutor(max_workers=80) as executor:
            future_to_oid = {executor.submit(fetch_bill, o): o for o in to_fetch}
            for future in as_completed(future_to_oid):
                res = future.result()
                if res:
                    oid_str, data_tr, r_dict = res
                    if data_tr is not None:
                        batch_to_save.append((oid_str, json.dumps(data_tr)))
                    all_results.append(r_dict)
                fetched_count += 1
                if fetched_count % 500 == 0:
                    save_cache_batch(batch_to_save)
                    batch_to_save = []
                    if progress_callback:
                        progress_callback(f"Fetched {fetched_count}/{total_fetch} new bills from tracking API...")

        if batch_to_save:
            save_cache_batch(batch_to_save)

    if progress_callback:
        progress_callback("Formatting large Excel file...")

    wb = Workbook()
    ws = wb.active
    ws.title = "Tracking Logs"

    border = Border(
        left=Side(style="thin"), right=Side(style="thin"),
        top=Side(style="thin"), bottom=Side(style="thin")
    )
    fill_hdr1 = PatternFill(start_color="FFE699", end_color="FFE699", fill_type="solid")
    fill_hdr2 = PatternFill(start_color="BDD7EE", end_color="BDD7EE", fill_type="solid")
    fill_alt = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")

    f_title = _font(is_bold=True, size=16, color="C00000")
    
    col_headers = ["BILL ID", "SERVICE TYPE", "VAS"]
    for sc in FULL_STATUS_CODES_TEMPLATE:
        col_headers.extend([f"[{sc}] LOG", f"[{sc}] UNIT", f"[{sc}] TIME"])
    col_headers.extend(["LATEST_STATUS", "LATEST_UNIT", "LATEST_TIME", "LATEST_USER"])

    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(col_headers))
    t_cell = ws.cell(1, 1, f"TRACKING STATUS LOGS REPORT - CURRENT MONTH ({len(all_results)} BILLS)")
    t_cell.font = f_title
    t_cell.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 32

    ws.row_dimensions[3].height = 26
    for c_idx, h_text in enumerate(col_headers, 1):
        cell = ws.cell(3, c_idx, h_text)
        cell.font = _font(h_text, is_bold=True)
        cell.fill = fill_hdr1 if c_idx in (1, 2, 3) or ((c_idx - 4) // 3) % 2 == 0 else fill_hdr2
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = border

    all_results.sort(key=lambda x: str(x.get("BILL ID", "")))
    r_idx = 4
    for r_data in all_results:
        row_fill = fill_alt if r_idx % 2 == 0 else None
        ws.row_dimensions[r_idx].height = 20

        c_bill = ws.cell(r_idx, 1, str(r_data.get("BILL ID", "")))
        c_bill.font = _font(c_bill.value); c_bill.border = border
        c_bill.alignment = Alignment(horizontal="left", vertical="center")

        svc_val = str(r_data.get("SERVICE TYPE", ""))
        c_svc = ws.cell(r_idx, 2, svc_val)
        c_svc.font = _font(svc_val); c_svc.border = border
        if row_fill: c_svc.fill = row_fill
        c_svc.alignment = Alignment(horizontal="center", vertical="center")

        vas_val = str(r_data.get("VAS", ""))
        c_vas = ws.cell(r_idx, 3, vas_val)
        c_vas.font = _font(vas_val, is_bold=True, color="C00000") if vas_val else _font(vas_val)
        c_vas.border = border
        if row_fill: c_vas.fill = row_fill
        c_vas.alignment = Alignment(horizontal="center", vertical="center")

        trips_map = r_data.get("_trips_map", {})
        col_pos = 4
        for sc in FULL_STATUS_CODES_TEMPLATE:
            unit_val, time_val = trips_map.get(sc, ("", ""))
            log_val = sc if unit_val or time_val else ""

            c_log = ws.cell(r_idx, col_pos, log_val)
            c_log.font = _font(log_val, is_bold=True); c_log.border = border
            if row_fill: c_log.fill = row_fill
            c_log.alignment = Alignment(horizontal="center", vertical="center")
            col_pos += 1

            c_unit = ws.cell(r_idx, col_pos, unit_val)
            c_unit.font = _font(unit_val); c_unit.border = border
            if row_fill: c_unit.fill = row_fill
            c_unit.alignment = Alignment(horizontal="center", vertical="center")
            col_pos += 1

            c_time = ws.cell(r_idx, col_pos, time_val)
            c_time.font = _font(time_val); c_time.border = border
            if row_fill: c_time.fill = row_fill
            c_time.alignment = Alignment(horizontal="center", vertical="center")
            col_pos += 1

        for val in [r_data.get("LATEST_STATUS",""), r_data.get("LATEST_UNIT",""), r_data.get("LATEST_TIME",""), r_data.get("LATEST_USER","")]:
            c_last = ws.cell(r_idx, col_pos, val)
            c_last.font = _font(val); c_last.border = border
            if row_fill: c_last.fill = row_fill
            c_last.alignment = Alignment(horizontal="center", vertical="center")
            col_pos += 1

        r_idx += 1

    ws.freeze_panes = "D4"
    for col in ws.columns:
        max_len = max(len(str(cell.value or "")) for cell in col[:80])
        col_letter = get_column_letter(col[0].column)
        ws.column_dimensions[col_letter].width = max(max_len + 3, 11)

    out_file_name = f"Bill_Tracking_Status_Logs_{first_of_this_month.strftime('%d%b')}_{yesterday.strftime('%d%b')}_{now.strftime('%Y%m%d')}_FIXED.xlsx"
    out_path = os.path.join(HERE, out_file_name)
    wb.save(out_path)
    
    desktop_path = os.path.join(os.path.expanduser("~"), "Desktop", out_file_name)
    try:
        shutil.copy2(out_path, desktop_path)
    except Exception:
        pass

    if progress_callback:
        progress_callback("Done!")
        
    return out_path

if __name__ == "__main__":
    def pcb(msg):
        print("[STATUS]", msg)
    res = generate_tracking_log_report(progress_callback=pcb)
    print("Report generated at:", res)
