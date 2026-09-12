# -*- coding: utf-8 -*-
"""
export_sep_fixed_wide_range.py
Generates full Tracking Status Logs Report for All Bills:
  - Source data: Aug 1 - Sep 8 wide range (tmp_wide_range.xlsx, 71k rows)
  - Filter: CURRENT TIME within 01/09/2026 - 08/09/2026 (last action in Sep 1-8)
  - This captures bills CREATED in August but DELIVERED/UPDATED in September
  - Identical format to Bill_Tracking_Status_Logs_Returned_520_01Aug_31Aug_20260904.xlsx
"""

import os
import sys
import json
import sqlite3
import shutil
import requests
import pandas as pd
from datetime import datetime
from concurrent.futures import ThreadPoolExecutor, as_completed
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.utils import get_column_letter

HERE = os.path.dirname(os.path.abspath(__file__))
CONFIG_PATH = os.path.join(HERE, "config.json")
CACHE_DB_PATH = os.path.join(HERE, "cache", "tracking_cache_wide_sep_01_11.db")
OLD_CACHE_DB = os.path.join(HERE, "cache", "tracking_cache_wide_sep_01_09.db")
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

def seed_cache_from_old():
    init_cache_db()
    if not os.path.exists(OLD_CACHE_DB):
        return
    conn_old = sqlite3.connect(OLD_CACHE_DB)
    old_items = conn_old.execute("SELECT order_id, data_json FROM tracking_cache").fetchall()
    conn_old.close()

    to_seed = []
    for oid, d_str in old_items:
        try:
            d = json.loads(d_str)
            st = str(d.get("LATEST_STATUS", "")).strip()
            if st in TERMINAL_STATUSES:
                to_seed.append((oid, d_str))
        except Exception:
            pass

    conn_new = sqlite3.connect(CACHE_DB_PATH)
    conn_new.executemany("INSERT OR IGNORE INTO tracking_cache (order_id, data_json) VALUES (?, ?)", to_seed)
    conn_new.commit()
    conn_new.close()
    print(f"[CACHE] Seeded {len(to_seed)} completed/terminal bills from old cache.")

def main():
    sys.stdout.reconfigure(encoding="utf-8")
    with open(CONFIG_PATH, "r", encoding="utf-8") as f:
        cfg = json.load(f)

    seed_cache_from_old()

    # Load wide range data (Aug 1 - Sep 11)
    detail_file = os.path.join(HERE, "tmp_wide_range_01_11_sep.xlsx")
    print(f"[INFO] Loading wide-range data: {detail_file}...")
    df = pd.read_excel(detail_file)
    df.columns = [str(c).strip().upper() for c in df.columns]
    print(f"[INFO] Total rows in wide range: {len(df)}")

    # Filter: CURRENT TIME or CREATED DATE must be within Sep 1 - Sep 11 2026 (midnight to midnight)
    time_col = next((c for c in df.columns if "CURRENT TIME" in c), None)
    created_col = next((c for c in df.columns if "CREATED DATE" in c), None)
    if not time_col:
        print("[ERROR] CURRENT TIME column not found!")
        return

    df["_parsed_cur_time"] = pd.to_datetime(df[time_col], dayfirst=True, errors='coerce')
    df["_parsed_created"]  = pd.to_datetime(df[created_col], dayfirst=True, errors='coerce') if created_col else None

    sep_start = pd.Timestamp("2026-09-01 00:00:00")
    sep_end   = pd.Timestamp("2026-09-11 23:59:59")

    mask = (df["_parsed_cur_time"] >= sep_start) & (df["_parsed_cur_time"] <= sep_end)
    if created_col:
        mask = mask | ((df["_parsed_created"] >= sep_start) & (df["_parsed_created"] <= sep_end))

    df_filtered = df[mask].copy()
    print(f"[INFO] Bills in Sep 1-11 (00:00:00 to 23:59:59): {len(df_filtered)}")

    # Verify key test bills
    test_bills = ['3204136498', '3204137254', '3204219910', '3204256687', '3304645944']
    df_filtered['ORDER ID'] = df_filtered['ORDER ID'].astype(str).str.strip()
    for b in test_bills:
        found = b in df_filtered['ORDER ID'].values
        print(f"  {'✓ FOUND' if found else '✗ MISSING'}: {b}")

    # Build metadata maps
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

    # Cache management
    cached_results = load_cached_orders()
    unique_oids = df_filtered["ORDER ID"].dropna().astype(str).unique()
    print(f"\n[INFO] Unique bills to process: {len(unique_oids)}")
    print(f"[INFO] Already cached: {len(cached_results)}")

    to_fetch = [oid for oid in unique_oids if oid not in cached_results]
    print(f"[INFO] Need to fetch from API: {len(to_fetch)}")

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
            r_sc = session.get(
                "https://gw-express.metfone.com.kh/tms-receiving/api/v1/orders/search",
                params={"order_code": oid_str}, timeout=8
            )
            if r_sc.status_code == 200:
                vas_val = str(r_sc.json().get("added_service_code") or "").strip()
        except Exception:
            pass

        if not vas_val and oid_str in vas_fee_map:
            vas_val = "VTT"
        row_dict["VAS"] = vas_val

        if not trips:
            return row_dict

        trips_sorted = list(reversed(trips))
        if not svc_type and isinstance(data_tr, dict):
            svc_type = str(data_tr.get("serviceType") or data_tr.get("serviceName") or data_tr.get("service") or "").strip()
            row_dict["SERVICE TYPE"] = svc_type

        latest_st = latest_unit = latest_time = latest_user = ""
        for t in trips_sorted:
            st = str(t.get("status", "") or "").lstrip("S").strip()
            po = t.get("postOffice") or {}
            unit = t.get("postcode") or (po.get("code") if isinstance(po, dict) else "") or ""
            if not unit and "handoverInfo" in t:
                unit = t.get("handoverInfo", {}).get("departmentCode", "")

            dt_raw = t.get("updatedAt", "")
            dt_str = ""
            dt_obj = None
            if dt_raw:
                try:
                    dt_obj = pd.to_datetime(dt_raw)
                    dt_str = dt_obj.strftime("%d/%m/%Y %H:%M:%S")
                except Exception:
                    dt_str = str(dt_raw)

            # ── STRICT DATE FILTER: skip any trip outside Sep 1-9 ──────────
            if dt_obj is not None:
                dt_naive = dt_obj.replace(tzinfo=None) if dt_obj.tzinfo else dt_obj
                if not (sep_start <= dt_naive <= sep_end):
                    continue  # skip trips outside range (e.g. Sep 10 updates)

            upd_obj = t.get("updatedBy")
            upd_user = ""
            if isinstance(upd_obj, dict):
                upd_user = str(upd_obj.get("name") or "").strip()
            elif isinstance(upd_obj, str):
                upd_user = upd_obj.strip()

            if st:
                row_dict["_trips_map"][st] = (unit, dt_str)
            latest_st = st
            latest_unit = unit
            latest_time = dt_str
            if upd_user:
                latest_user = upd_user

        if not latest_user:
            latest_user = action_user_map.get(oid_str, "")

        row_dict["LATEST_STATUS"] = latest_st or def_st
        row_dict["LATEST_UNIT"]   = latest_unit or def_unit
        row_dict["LATEST_TIME"]   = latest_time or def_time
        row_dict["LATEST_USER"]   = latest_user or def_user
        return row_dict

    results_dict = dict(cached_results)
    if to_fetch:
        print(f"\n[INFO] Fetching {len(to_fetch)} bills with 80 threads...")
        batch_to_save = []
        count = 0
        total = len(to_fetch)
        with ThreadPoolExecutor(max_workers=80) as executor:
            futures = {executor.submit(fetch_bill, oid): oid for oid in to_fetch}
            for future in as_completed(futures):
                oid = futures[future]
                try:
                    res = future.result()
                    if res:
                        results_dict[oid] = res
                        batch_to_save.append((oid, json.dumps(res, ensure_ascii=False)))
                except Exception:
                    pass
                count += 1
                if count % 1000 == 0 or count == total:
                    save_cache_batch(batch_to_save)
                    batch_to_save = []
                    print(f"[PROGRESS] {count}/{total} fetched ({len(results_dict)} total)...")
        if batch_to_save:
            save_cache_batch(batch_to_save)

    all_results = [results_dict[oid] for oid in unique_oids if oid in results_dict]
    print(f"\n[INFO] Total bills for Excel: {len(all_results)}")

    # Build Excel
    report_title = "ALL BILL TRACKING STATUS LOGS REPORT (PENDING & SHIPPED) — 01 TO 11 SEPTEMBER 2026"
    sheet_name = "All Tracking Logs"
    print(f"[INFO] Building Excel ({len(FULL_STATUS_CODES_TEMPLATE)} status codes, 229 cols)...")

    wb = Workbook()
    ws = wb.active
    ws.title = sheet_name

    font_family = "Calibri"
    f_title  = Font(name=font_family, size=14, bold=True, color="FFFFFF")
    f_header = Font(name=font_family, size=10, bold=True, color="FFFFFF")
    f_data   = Font(name=font_family, size=9)
    f_log    = Font(name=font_family, size=9, bold=True, color="1E3A8A")
    f_vas    = Font(name=font_family, size=9, bold=True, color="047857")

    fill_title = PatternFill(start_color="0F172A", end_color="0F172A", fill_type="solid")
    fill_hdr1  = PatternFill(start_color="1E293B", end_color="1E293B", fill_type="solid")
    fill_hdr2  = PatternFill(start_color="334155", end_color="334155", fill_type="solid")
    fill_alt   = PatternFill(start_color="F8FAFC", end_color="F8FAFC", fill_type="solid")
    thin       = Side(border_style="thin", color="CBD5E1")
    border     = Border(left=thin, right=thin, top=thin, bottom=thin)

    col_headers = ["BILL ID", "SERVICE TYPE", "VAS"]
    for sc in FULL_STATUS_CODES_TEMPLATE:
        col_headers.extend([f"LOG {sc}", f"UNIT {sc}", f"TIME {sc}"])
    col_headers.extend(["LATEST STATUS", "LATEST UNIT", "LATEST TIME", "LAST USER"])

    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(col_headers))
    t_cell = ws.cell(1, 1, f"{report_title} ({len(all_results)} Bills)")
    t_cell.font = f_title
    t_cell.fill = fill_title
    t_cell.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 32

    ws.row_dimensions[3].height = 26
    for c_idx, h_text in enumerate(col_headers, 1):
        cell = ws.cell(3, c_idx, h_text)
        cell.font = f_header
        cell.fill = fill_hdr1 if c_idx in (1, 2, 3) or ((c_idx - 4) // 3) % 2 == 0 else fill_hdr2
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = border

    all_results.sort(key=lambda x: str(x.get("BILL ID", "")))
    print("[INFO] Writing rows...")
    r_idx = 4
    for r_data in all_results:
        row_fill = fill_alt if r_idx % 2 == 0 else None
        ws.row_dimensions[r_idx].height = 20

        c_bill = ws.cell(r_idx, 1, str(r_data.get("BILL ID", "")))
        c_bill.font = f_data; c_bill.border = border
        c_bill.alignment = Alignment(horizontal="left", vertical="center")

        c_svc = ws.cell(r_idx, 2, str(r_data.get("SERVICE TYPE", "")))
        c_svc.font = f_data; c_svc.border = border
        if row_fill: c_svc.fill = row_fill
        c_svc.alignment = Alignment(horizontal="center", vertical="center")

        c_vas = ws.cell(r_idx, 3, str(r_data.get("VAS", "")))
        c_vas.font = f_vas if r_data.get("VAS") else f_data
        c_vas.border = border
        if row_fill: c_vas.fill = row_fill
        c_vas.alignment = Alignment(horizontal="center", vertical="center")

        trips_map = r_data.get("_trips_map", {})
        col_pos = 4
        for sc in FULL_STATUS_CODES_TEMPLATE:
            unit_val, time_val = trips_map.get(sc, ("", ""))
            log_val = sc if unit_val or time_val else ""

            c_log = ws.cell(r_idx, col_pos, log_val)
            c_log.font = f_log; c_log.border = border
            if row_fill: c_log.fill = row_fill
            c_log.alignment = Alignment(horizontal="center", vertical="center")
            col_pos += 1

            c_unit = ws.cell(r_idx, col_pos, unit_val)
            c_unit.font = f_data; c_unit.border = border
            if row_fill: c_unit.fill = row_fill
            c_unit.alignment = Alignment(horizontal="center", vertical="center")
            col_pos += 1

            c_time = ws.cell(r_idx, col_pos, time_val)
            c_time.font = f_data; c_time.border = border
            if row_fill: c_time.fill = row_fill
            c_time.alignment = Alignment(horizontal="center", vertical="center")
            col_pos += 1

        for val in [r_data.get("LATEST_STATUS",""), r_data.get("LATEST_UNIT",""), r_data.get("LATEST_TIME",""), r_data.get("LATEST_USER","")]:
            c_last = ws.cell(r_idx, col_pos, val)
            c_last.font = f_data; c_last.border = border
            if row_fill: c_last.fill = row_fill
            c_last.alignment = Alignment(horizontal="center", vertical="center")
            col_pos += 1

        r_idx += 1
        if (r_idx - 4) % 3000 == 0:
            print(f"[PROGRESS] Written {r_idx - 4}/{len(all_results)} rows...")

    ws.freeze_panes = "D4"
    print("[INFO] Setting column widths...")
    for col in ws.columns:
        max_len = max(len(str(cell.value or "")) for cell in col[:80])
        col_letter = get_column_letter(col[0].column)
        ws.column_dimensions[col_letter].width = max(max_len + 3, 11)

    primary_new = r"C:\Users\DELL\Desktop\Bill_Tracking_Status_Logs_01Sep_11Sep_20260911_FIXED.xlsx"
    local_new   = os.path.join(HERE, "Bill_Tracking_Status_Logs_01Sep_11Sep_20260911_FIXED.xlsx")
    desk_09     = r"C:\Users\DELL\Desktop\Bill_Tracking_Status_Logs_01Sep_09Sep_20260910_FIXED.xlsx"
    local_09    = os.path.join(HERE, "Bill_Tracking_Status_Logs_01Sep_09Sep_20260910_FIXED.xlsx")
    desk_08     = r"C:\Users\DELL\Desktop\Bill_Tracking_Status_Logs_01Sep_08Sep_20260909_FIXED.xlsx"
    local_08    = os.path.join(HERE, "Bill_Tracking_Status_Logs_01Sep_08Sep_20260909_FIXED.xlsx")

    print(f"[INFO] Saving to primary: {primary_new}...")
    wb.save(primary_new)
    print(f"[SUCCESS] Saved: {primary_new}")

    for p in [local_new, desk_09, local_09, desk_08, local_08]:
        try:
            shutil.copy2(primary_new, p)
            print(f"[SUCCESS] Synced report to: {p}")
        except Exception as e:
            print(f"[WARNING] Could not sync to {p}: {e}")

    print(f"\n[COMPLETE] Report ready — {len(all_results)} bills (Day 1-11, Midnight to Midnight)!")

if __name__ == "__main__":
    main()
