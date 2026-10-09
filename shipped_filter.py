"""
shipped_filter.py — Live Tracking Cross-Check & Shipped Bill Exclusion Filter
==============================================================================
Validates candidate pending orders against the real-time TMS live tracking API.
If an order is already delivered (S410/SHIPPED), returned (S520), or cancelled,
it is automatically purged from pending reports (/total pending, /push, /penalty).

Confirmed shipped bill IDs are saved to cache/confirmed_shipped_bills.json
so subsequent report runs filter them instantaneously with zero network overhead.
"""

import os
import re
import json
import time
import threading
from concurrent.futures import ThreadPoolExecutor, as_completed
import pandas as pd
import requests
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry

_LOCK = threading.Lock()
_CACHE_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "cache", "confirmed_shipped_bills.json")
_TRACKING_URL = "https://gw-express.metfone.com.kh/tms-tracking/api/v1/order-tracking"

_REALTIME_DONE_STATUSES = {
    'S410', '410',   # Delivered / Shipped / Giao thành công
    'S520', '520',   # Returned / Đã trả hàng
    'S99',  '99',    # Cancelled
    'S100', '100',   # Cancelled
    'S201', '201',   # Cancelled pickup
}

_TXT_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "shipped_bills.txt")

_memory_cache: set[str] = set()
_cache_loaded = False


def load_confirmed_shipped_ids() -> set[str]:
    """Load confirmed shipped bill IDs from persistent disk storage (both JSON cache & shipped_bills.txt)."""
    global _memory_cache, _cache_loaded
    with _LOCK:
        if _cache_loaded and _memory_cache:
            return set(_memory_cache)

        _memory_cache = set()

        # 1. Load from JSON cache
        if os.path.exists(_CACHE_FILE):
            try:
                with open(_CACHE_FILE, "r", encoding="utf-8") as f:
                    data = json.load(f)
                if isinstance(data, list):
                    _memory_cache.update(str(x).strip() for x in data if str(x).strip())
                elif isinstance(data, dict):
                    _memory_cache.update(str(k).strip() for k in data.keys() if str(k).strip())
            except Exception as e:
                print(f"[SHIPPED_FILTER] Warning: Failed reading cache file: {e}")

        # 2. Also load from shipped_bills.txt (plain-text, one per line like test_bills.txt)
        base_dirs = [
            os.path.dirname(os.path.abspath(__file__)),
            os.getcwd(),
        ]
        for d in base_dirs:
            t_path = os.path.join(d, "shipped_bills.txt")
            if os.path.exists(t_path):
                try:
                    with open(t_path, "r", encoding="utf-8") as f:
                        for line in f:
                            val = line.strip()
                            if val and not val.startswith("#"):
                                clean_val = re.sub(r'\.0$', '', val).strip()
                                if clean_val:
                                    _memory_cache.add(clean_val)
                except Exception as e:
                    print(f"[SHIPPED_FILTER] Warning: Failed reading shipped_bills.txt: {e}")

        _cache_loaded = True
        return set(_memory_cache)


def load_shipped_bills() -> set[str]:
    """Convenience alias matching penalty_report.load_test_bills()."""
    return load_confirmed_shipped_ids()


def save_confirmed_shipped_ids(new_ids: set[str] | list[str]) -> None:
    """Save newly confirmed shipped bill IDs to persistent disk storage (both JSON cache & shipped_bills.txt)."""
    global _memory_cache
    if not new_ids:
        return
    clean_new = {str(x).strip() for x in new_ids if str(x).strip()}
    if not clean_new:
        return

    with _LOCK:
        _memory_cache.update(clean_new)
        sorted_ids = sorted(list(_memory_cache))

        # 1. Save JSON cache
        os.makedirs(os.path.dirname(_CACHE_FILE), exist_ok=True)
        try:
            with open(_CACHE_FILE, "w", encoding="utf-8") as f:
                json.dump(sorted_ids, f, indent=2)
        except Exception as e:
            print(f"[SHIPPED_FILTER] Warning: Failed saving JSON cache file: {e}")

        # 2. Save shipped_bills.txt (like test_bills.txt)
        try:
            with open(_TXT_FILE, "w", encoding="utf-8") as f:
                f.write("# Confirmed Shipped / Delivered Bills (Excluded from Pending & Push Reports)\n")
                for oid in sorted_ids:
                    f.write(f"{oid}\n")
        except Exception as e:
            print(f"[SHIPPED_FILTER] Warning: Failed saving shipped_bills.txt: {e}")


def _load_api_cfg() -> dict:
    base_dirs = [
        os.path.dirname(os.path.abspath(__file__)),
        os.getcwd(),
        os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    ]
    for d in base_dirs:
        cfg_path = os.path.join(d, "config.json")
        if os.path.exists(cfg_path):
            try:
                with open(cfg_path, "r", encoding="utf-8") as f:
                    cfg = json.load(f)
                return cfg.get("api", {})
            except Exception:
                pass
    return {}


def batch_verify_shipped_live(order_ids: list[str], api_cfg: dict = None, max_workers: int = 80, timeout: int = 5) -> set[str]:
    """
    Check real-time live tracking for a list of candidate order IDs in parallel.
    Returns a set of order_id strings that are confirmed shipped/delivered on web.
    """
    if not order_ids:
        return set()

    cfg = api_cfg or _load_api_cfg()
    token = cfg.get("bearer_token", "")
    if not token:
        print("[SHIPPED_FILTER] Warning: No bearer token available for tracking check")
        return set()

    headers = {
        "Authorization": f"Bearer {token}",
        "Referer": cfg.get("referer", "https://opsexpress.metfone.com.kh/"),
        "Accept-Language": "vi-VN",
        "Accept": "application/json, text/plain, */*",
        "Content-Type": "application/json",
        "x-client-id": cfg.get("x_client_id", "TMS_ANDROID"),
        "User-Agent": (
            "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
            "AppleWebKit/537.36 (KHTML, like Gecko) "
            "Chrome/148.0.0.0 Safari/537.36"
        ),
    }

    session = requests.Session()
    adapter = HTTPAdapter(pool_connections=max_workers, pool_maxsize=max_workers, max_retries=Retry(total=1, backoff_factor=0.1))
    session.mount("https://", adapter)
    session.headers.update(headers)

    def _check_bill(oid: str) -> tuple[str, bool]:
        try:
            r = session.get(_TRACKING_URL, params={"order_id": oid}, timeout=timeout)
            if r.status_code != 200:
                return oid, False
            data = r.json()
            trips = data.get("trackingTrips", [])
            if not trips:
                return oid, False

            # Check latest trip (index 0)
            latest_status = str(trips[0].get("status", "")).upper().strip()
            status_name = str(trips[0].get("statusName", "")).upper().strip()
            item_type = str(trips[0].get("itemType", "")).upper().strip()
            desc = str(trips[0].get("desc", "")).upper().strip()

            is_done = (
                latest_status in _REALTIME_DONE_STATUSES
                or status_name in ('SHIPPED', 'DELIVERED', 'RETURNED')
                or item_type in ('SHIPPED', 'DELIVERED', 'RETURN_SUCCESS', 'CUSTOMER_RETURN')
                or (latest_status.startswith('S4') and 'giao thành công' in desc.lower() and 'bàn giao' not in desc.lower())
            )

            # Also scan all trips: if status S410 or SHIPPED is anywhere in history
            if not is_done:
                for trip in trips:
                    t_st = str(trip.get("status", "")).upper().strip()
                    t_name = str(trip.get("statusName", "")).upper().strip()
                    t_desc = str(trip.get("desc", "")).strip().lower()
                    if t_st in ('S410', '410') or t_name in ('SHIPPED', 'DELIVERED') or 'delivered.' in t_desc:
                        is_done = True
                        break

            return oid, is_done
        except Exception:
            return oid, False

    confirmed = set()
    with ThreadPoolExecutor(max_workers=max_workers) as executor:
        futures = {executor.submit(_check_bill, oid): oid for oid in order_ids}
        for f in as_completed(futures):
            try:
                oid, is_done = f.result()
                if is_done:
                    confirmed.add(oid)
            except Exception:
                pass

    return confirmed


def filter_shipped_bills_from_df(df: pd.DataFrame, api_cfg: dict = None, verify_live: bool = True, max_workers: int = 80) -> tuple[pd.DataFrame, set[str]]:
    """
    Remove all delivered / shipped bills from a DataFrame so they never appear in pending reports.
    1. Removes all bills already known to be shipped from disk cache.
    2. Removes all bills with status 410, 520, 201, 99, 100, or text 'SHIPPED'/'DELIVERED'.
    3. If verify_live=True, verifies candidate pending orders against live tracking API in parallel.
    4. Updates disk cache with newly found shipped bills.
    Returns (cleaned_df, set_of_removed_order_ids).
    """
    if df.empty:
        return df, set()

    order_col = next(
        (c for c in df.columns if str(c).strip().upper() in ('ORDER ID', 'ORDER_ID', 'BILL_ID', 'BILL ID', 'WAYBILL', 'ORDER ID/ WAYBILL')),
        None
    )
    if not order_col:
        for c in df.columns:
            if 'ORDER' in str(c).upper():
                order_col = c
                break

    if not order_col:
        return df, set()

    removed_ids: set[str] = set()

    # Step 1: Remove bills already in confirmed shipped cache
    cached_shipped = load_confirmed_shipped_ids()
    clean_ids = df[order_col].astype(str).str.strip().str.replace(r'\.0$', '', regex=True)
    mask_cached = clean_ids.isin(cached_shipped)
    if mask_cached.any():
        c_count = mask_cached.sum()
        removed_ids.update(clean_ids[mask_cached].tolist())
        df = df[~mask_cached].copy()
        clean_ids = df[order_col].astype(str).str.strip().str.replace(r'\.0$', '', regex=True)
        print(f"  [SHIPPED_FILTER] Removed {c_count} bills from persistent shipped cache")

    # Step 2: Remove explicit shipped statuses in DataFrame
    status_col = next((c for c in df.columns if str(c).strip().upper() in ('CURRENT STATUS', 'STATUS', 'STATUS_CODE')), None)
    if status_col:
        sc = df[status_col].astype(str).str.extract(r'^(\d{3})')[0]
        st_text = df[status_col].astype(str).str.upper()

        explicit_done = (
            sc.isin(['410', '520', '201', '99', '100'])
            | st_text.str.contains('410', na=False)
            | st_text.str.contains('520', na=False)
            | st_text.str.contains('DELIVERED', na=False)
            | st_text.str.contains('SHIPPED', na=False)
            | st_text.str.contains('GIAO THÀNH CÔNG', na=False)
            | st_text.str.contains('ĐÃ GIAO', na=False)
            | st_text.str.contains('RETURN COMPLETED', na=False)
        )
        if explicit_done.any():
            e_count = explicit_done.sum()
            removed_ids.update(clean_ids[explicit_done].tolist())
            save_confirmed_shipped_ids(clean_ids[explicit_done].tolist())
            df = df[~explicit_done].copy()
            clean_ids = df[order_col].astype(str).str.strip().str.replace(r'\.0$', '', regex=True)
            print(f"  [SHIPPED_FILTER] Removed {e_count} bills by explicit shipped/delivered status")

    # Step 3: Live real-time check for remaining candidate pending bills
    if verify_live and not df.empty:
        cand_ids = clean_ids.unique().tolist()
        cand_ids = [x for x in cand_ids if x and x != 'nan' and x not in cached_shipped]

        if cand_ids:
            print(f"  [SHIPPED_FILTER] Live verifying {len(cand_ids)} candidate pending bills against tracking API...")
            t0 = time.time()
            live_delivered = batch_verify_shipped_live(cand_ids, api_cfg=api_cfg, max_workers=max_workers)
            t1 = time.time()
            if live_delivered:
                print(f"  [SHIPPED_FILTER] Live check confirmed {len(live_delivered)} actually shipped bills in {t1-t0:.2f}s! Purging them now...")
                save_confirmed_shipped_ids(live_delivered)
                removed_ids.update(live_delivered)
                mask_live = clean_ids.isin(live_delivered)
                df = df[~mask_live].copy()
            else:
                print(f"  [SHIPPED_FILTER] Live check verified: all {len(cand_ids)} bills are genuinely pending ({t1-t0:.2f}s)")

    return df, removed_ids
