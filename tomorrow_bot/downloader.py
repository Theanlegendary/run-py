"""
downloader.py
Goi API export-detail (giong lenh curl) de tai file Excel chi tiet don.
"""

import json
from datetime import datetime, timedelta

import requests


def day14_to_today_range():
    """Return (from_date, to_date) starting from 30 days ago to today to prevent cut-off on the 14th."""
    today = datetime.now()
    start_date = today - timedelta(days=30)
    return start_date.strftime("%Y%m%d"), today.strftime("%Y%m%d")


def download_detail(api_cfg, out_path, from_date=None, to_date=None, branch_code=None, force_refresh=False):
    """
    Goi API, luu file Excel ve out_path.
    Mac dinh lay tu ngay 14 den hom nay.
    branch_code: override branch_code from config.
    """
    import os
    import time
    import shutil

    # Cache logic: Only cache when downloading the default date range
    is_default_range = (from_date is None and to_date is None)
    cache_minutes = api_cfg.get("cache_minutes", 0)

    # Cache file stored locally in a 'cache' folder under this script's directory
    cache_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "cache")
    cache_file = os.path.join(cache_dir, "latest_detail.xlsx")

    if is_default_range and cache_minutes > 0 and not force_refresh:
        if os.path.exists(cache_file):
            mtime = os.path.getmtime(cache_file)
            age_seconds = time.time() - mtime
            if age_seconds < (cache_minutes * 60):
                print(f"[CACHE] Using cached Excel data ({int(age_seconds)}s old)")
                shutil.copy2(cache_file, out_path)
                return out_path

    if from_date is None or to_date is None:
        from_date, to_date = day14_to_today_range()


    headers = {
        "Authorization": f"Bearer {api_cfg['bearer_token']}",
        "Referer": api_cfg.get("referer", "https://opsexpress.metfone.com.kh/"),
        "Accept-Language": "vi-VN",
        "Accept": "application/json, text/plain, */*",
        "Content-Type": "application/json",
        "x-client-id": api_cfg.get("x_client_id", "TMS_ANDROID"),
        "User-Agent": ("Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
                       "AppleWebKit/537.36 (KHTML, like Gecko) "
                       "Chrome/148.0.0.0 Safari/537.36"),
    }
    
    # If caching is active, we always download the full branch list to ensure the cached
    # file contains all data and can satisfy future partial or full requests.
    api_branch_code = api_cfg.get("branch_code", "PRE,PNP,SVA,KAN")
    fetch_branch_code = api_branch_code if (is_default_range and cache_minutes > 0) else (branch_code or api_branch_code)

    payload = {
        "from_date": from_date,
        "to_date": to_date,
        "branch_code": fetch_branch_code,
    }

    resp = requests.post(api_cfg["url"], headers=headers,
                         data=json.dumps(payload), timeout=120)
    resp.raise_for_status()

    ctype = resp.headers.get("Content-Type", "")
    
    def _save_to_cache():
        if is_default_range and cache_minutes > 0:
            try:
                os.makedirs(cache_dir, exist_ok=True)
                shutil.copy2(out_path, cache_file)
                print(f"[CACHE] Saved downloaded Excel data to cache")
            except Exception as e:
                print(f"[CACHE] Warning: Failed to save to cache: {e}")

    # API tra ve truc tiep file excel
    if ("spreadsheet" in ctype or "octet-stream" in ctype
            or "excel" in ctype or out_path.endswith(".xlsx")):
        with open(out_path, "wb") as f:
            f.write(resp.content)
        _save_to_cache()
        return out_path

    # Truong hop tra ve JSON chua link tai
    try:
        data = resp.json()
    except ValueError:
        with open(out_path, "wb") as f:
            f.write(resp.content)
        _save_to_cache()
        return out_path

    file_url = (data.get("data", {}) or {}).get("url") or data.get("url")
    if file_url:
        r2 = requests.get(file_url, headers={"Authorization": headers["Authorization"]},
                          timeout=120)
        r2.raise_for_status()
        with open(out_path, "wb") as f:
            f.write(r2.content)
        _save_to_cache()
        return out_path

    raise RuntimeError(f"Khong tim thay file Excel trong phan hoi API: {data}")

def download_revenue_detail(api_cfg, out_path, from_date=None, to_date=None, force_refresh=False):
    """
    Fetch the Revenue Pickup Detail Export to get VAS_SERVICE per bill.
    """
    import os
    import time
    import shutil

    # Cache logic
    is_default_range = (from_date is None and to_date is None)
    cache_minutes = api_cfg.get("cache_minutes", 0)
    cache_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "cache")
    cache_file = os.path.join(cache_dir, "latest_revenue.xlsx")

    if is_default_range and cache_minutes > 0 and not force_refresh:
        if os.path.exists(cache_file):
            mtime = os.path.getmtime(cache_file)
            age_seconds = time.time() - mtime
            if age_seconds < (cache_minutes * 60):
                print(f"[CACHE] Using cached Revenue data ({int(age_seconds)}s old)")
                shutil.copy2(cache_file, out_path)
                return out_path

    if from_date is None or to_date is None:
        from_date, to_date = day14_to_today_range()

    base = api_cfg["url"].split("/tms-report/")[0]
    url = f"{base}/tms-report/api/v1/revenue/pickup-revenue/export-detail?from_date={from_date}&to_date={to_date}&type=PICKUP_REVENUE"

    headers = {
        "Authorization": f"Bearer {api_cfg['bearer_token']}",
        "Referer": api_cfg.get("referer", "https://opsexpress.metfone.com.kh/"),
        "Accept-Language": "vi-VN",
        "Accept": "application/json, text/plain, */*",
        "Content-Type": "application/json",
        "x-client-id": api_cfg.get("x_client_id", "TMS_ANDROID"),
        "User-Agent": ("Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
                       "AppleWebKit/537.36 (KHTML, like Gecko) "
                       "Chrome/148.0.0.0 Safari/537.36"),
    }

    resp = requests.get(url, headers=headers, timeout=120)
    resp.raise_for_status()

    ctype = resp.headers.get("Content-Type", "")
    
    def _save_to_cache():
        if is_default_range and cache_minutes > 0:
            try:
                os.makedirs(cache_dir, exist_ok=True)
                shutil.copy2(out_path, cache_file)
            except Exception as e:
                pass

    if ("spreadsheet" in ctype or "octet-stream" in ctype or "excel" in ctype or out_path.endswith(".xlsx")):
        with open(out_path, "wb") as f:
            f.write(resp.content)
        _save_to_cache()
        return out_path

    try:
        data = resp.json()
    except ValueError:
        with open(out_path, "wb") as f:
            f.write(resp.content)
        _save_to_cache()
        return out_path

    file_url = (data.get("data", {}) or {}).get("url") or data.get("url")
    if file_url:
        r2 = requests.get(file_url, headers={"Authorization": headers["Authorization"]}, timeout=120)
        r2.raise_for_status()
        with open(out_path, "wb") as f:
            f.write(r2.content)
        _save_to_cache()
        return out_path

    raise RuntimeError(f"Failed to fetch Revenue Excel: {data}")


# Alias for backward compatibility
download_pickup_revenue = download_revenue_detail


def download_post_offices(api_cfg, branch_code, limit=200):
    """
    Fetch post office list for a branch from the OPS management API.
    E.g. branch_code="KAMA" returns KAMA001, KAMA002, ...
         branch_code="KAM"  returns KAMA001, KAMP001, ...
    Returns list of post office dicts.
    """
    import time
    base = api_cfg["url"].split("/tms-report/")[0]  # https://gw-express.metfone.com.kh
    url = f"{base}/vtp-user/api/v1/departments/posts/search"

    headers = {
        "Authorization": f"Bearer {api_cfg['bearer_token']}",
        "Referer": api_cfg.get("referer", "https://opsexpress.metfone.com.kh/"),
        "Accept": "application/json, text/plain, */*",
        "Content-Type": "application/json",
        "x-client-id": api_cfg.get("x_client_id", "TMS_ANDROID"),
        "x-need-count": "true",
        "User-Agent": ("Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
                       "AppleWebKit/537.36 (KHTML, like Gecko) "
                       "Chrome/148.0.0.0 Safari/537.36"),
    }

    keyword = branch_code.upper()
    all_items = []
    offset = 0

    while True:
        params = {"offset": offset, "limit": limit}
        body = {"keyword": keyword}

        # Try with retries
        success = False
        last_err = None
        for attempt in range(3):
            try:
                resp = requests.post(url, headers=headers, params=params,
                                     data=json.dumps(body), timeout=180)
                resp.raise_for_status()
                success = True
                break
            except Exception as e:
                last_err = e
                # If we are on page > 1 (offset > 0) and get a 500 error, 
                # it's highly likely to be out of range, so we break the pagination loop
                if offset > 0 and hasattr(e, 'response') and e.response is not None:
                    if e.response.status_code == 500:
                        print(f"[DEBUG] Out of bounds offset {offset} for {keyword} (HTTP 500). Stopping pagination.")
                        break
                time.sleep(1.0)
        
        if not success:
            # If we broke because of offset out-of-bounds, success is False but last_err is status code 500
            if offset > 0 and last_err and hasattr(last_err, 'response') and last_err.response is not None and last_err.response.status_code == 500:
                break
            # Otherwise, it's a real failure
            raise last_err

        data = resp.json()
        items = data.get("items", data.get("data", []))
        if not items:
            break

        all_items.extend(items)
        if len(items) < limit:
            break
        offset += limit

    # Filter to only codes starting with the keyword (e.g. "KAMA" or "KAM")
    filtered = [
        item for item in all_items
        if str(item.get("code", "")).upper().startswith(keyword)
    ]

    return filtered


def download_all_post_offices(api_cfg, limit=500):
    """Fetch ALL post offices without keyword filter. Uses limit=500 to minimize round-trips."""
    import time
    base = api_cfg["url"].split("/tms-report/")[0]
    url = f"{base}/vtp-user/api/v1/departments/posts/search"

    headers = {
        "Authorization": f"Bearer {api_cfg['bearer_token']}",
        "Referer": api_cfg.get("referer", "https://opsexpress.metfone.com.kh/"),
        "Accept": "application/json, text/plain, */*",
        "Content-Type": "application/json",
        "x-client-id": api_cfg.get("x_client_id", "TMS_ANDROID"),
        "x-need-count": "true",
        "User-Agent": ("Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
                       "AppleWebKit/537.36 (KHTML, like Gecko) "
                       "Chrome/148.0.0.0 Safari/537.36"),
    }

    all_items = []
    offset = 0

    while True:
        params = {"offset": offset, "limit": limit}
        body = {"keyword": ""}

        success = False
        last_err = None
        for attempt in range(3):
            try:
                resp = requests.post(url, headers=headers, params=params,
                                     data=json.dumps(body), timeout=180)
                resp.raise_for_status()
                success = True
                break
            except Exception as e:
                last_err = e
                if offset > 0 and hasattr(e, 'response') and e.response is not None:
                    if e.response.status_code == 500:
                        break
                time.sleep(1.0)

        if not success:
            if offset > 0 and last_err and hasattr(last_err, 'response') and last_err.response is not None and last_err.response.status_code == 500:
                break
            raise last_err

        data = resp.json()
        items = data.get("items", data.get("data", []))
        if not items:
            break

        all_items.extend(items)
        if len(items) < limit:
            break
        offset += limit

    return all_items


if __name__ == "__main__":
    import sys
    cfg = json.load(open("config.json"))
    out = download_detail(cfg["api"], sys.argv[1] if len(sys.argv) > 1
                          else "detail.xlsx")
    print("Da tai:", out)
