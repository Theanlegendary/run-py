"""
vip_tracker.py
Module for tracking all orders/bills belonging to a VIP customer by phone number or name.
Generates grouped daily summaries and detailed status breakdowns with real-time tracking checks.
"""

import os
import re
import glob
import json
import logging
from datetime import datetime
from concurrent.futures import ThreadPoolExecutor
import requests
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

log = logging.getLogger("vip_tracker")
logging.getLogger("urllib3.connectionpool").setLevel(logging.ERROR)
HERE = os.path.dirname(os.path.abspath(__file__))

_TRACKING_URL = "https://gw-express.metfone.com.kh/tms-tracking/api/v1/order-tracking"


KNOWN_VIP_PROFILES = {
    "060203333": {
        "name": "Shop Leng Kim San (lengkimsan)",
        "phone": "+85560203333",
        "address": "ភូមិពាណិជ្ជកម្ម (Peanicheakkam Village) - Rokar Thum - Chbar Mon - Kampong Speu - Cambodia",
        "group_id": -5205920813,
        "group_title": "VIP_Lengkimsan"
    },
    "0962752464": {
        "name": "Kang Rithy",
        "phone": "+855962752464",
        "address": "SPEP001 (SPE)",
        "group_id": -5295202050,
        "group_title": "VIP_Kang Rithy"
    },
    "067620062": {
        "name": "Kang Rithy",
        "phone": "+855962752464",
        "address": "SPEP001 (SPE)",
        "group_id": -5295202050,
        "group_title": "VIP_Kang Rithy"
    }
}


def _clean_phone(phone_str: str) -> str:
    """Normalize phone number to local 9-digit format (e.g. 060203333)."""
    digits = re.sub(r"\D", "", str(phone_str or ""))
    if digits.startswith("855") and len(digits) > 8:
        digits = "0" + digits[3:]
    elif not digits.startswith("0") and 8 <= len(digits) <= 9:
        digits = "0" + digits
    return digits


def get_latest_export_file() -> str:
    """Locate the most recently modified export_*.xlsx or latest_detail.xlsx."""
    candidates = []
    # 1. Project cache directory
    cache_dir = os.path.join(HERE, "cache")
    if os.path.exists(cache_dir):
        for root, _, files in os.walk(cache_dir):
            for f in files:
                if f.endswith(".xlsx") and ("export_" in f or f == "latest_detail.xlsx"):
                    candidates.append(os.path.join(root, f))
    
    # 2. Root directory fallbacks
    for f in glob.glob(os.path.join(HERE, "export_*.xlsx")):
        candidates.append(f)
    
    if not candidates:
        return ""
    
    # Sort by modification time descending
    candidates.sort(key=os.path.getmtime, reverse=True)
    return candidates[0]


def _check_bill_tracking(order_id: str, bearer_token: str, default_branch: str = "", timeout: int = 4, session: requests.Session = None) -> tuple:
    """
    Query real-time tracking API for a single order.
    Returns: (order_id, branch, status_text, is_delivered)
    """
    headers = {
        "Authorization": f"Bearer {bearer_token}",
        "Accept": "application/json",
        "User-Agent": "VIPTracker/1.0"
    }
    client = session or requests
    try:
        r = client.get(_TRACKING_URL, headers=headers, params={"order_id": order_id}, timeout=timeout)
        if r.status_code == 200:
            data = r.json()
            trips = data.get("trackingTrips", [])
            if trips:
                t0 = trips[0]
                status_code = str(t0.get("status", "")).upper().strip()
                status_name = str(t0.get("statusName", "")).upper().strip()
                desc = str(t0.get("desc", "")).strip()
                postcode = (
                    t0.get("deliveryPostcode")
                    or t0.get("postcode")
                    or default_branch
                    or ""
                ).strip()

                # 1. Delivered / Successful
                if status_code in ("S410", "410") or status_name in ("SHIPPED", "DELIVERED"):
                    return order_id, postcode, "Successful", True

                # 2. Returned Successfully (S520) - Final Flow / Closed
                if (
                    status_code in ("S520", "520")
                    or status_name in ("RETURNED", "RETURN_SUCCESSFUL", "DA_TRA_HANG")
                    or "TRẢ HÀNG THÀNH CÔNG" in desc.upper()
                    or "ĐÃ TRẢ HÀNG" in desc.upper()
                    or "RETURN SUCCESSFULLY" in desc.upper()
                ):
                    return order_id, postcode, "Return Successfully", True

                # 3. Postponed by Customer (S420)
                if status_code in ("S420", "420"):
                    m_date = re.search(r"(\d{1,2}/\d{1,2}/\d{4})", desc)
                    if m_date:
                        return order_id, postcode, f"Customer will on {m_date.group(1)}", False
                    if desc:
                        return order_id, postcode, desc, False
                    return order_id, postcode, "Customer postponed", False

                # 4. Storage / Hold (S472)
                if status_code in ("S472", "472"):
                    m_store = re.search(r"storage for (\d+) days", desc, re.I)
                    if m_store:
                        return order_id, postcode, f"Warehouse storage ({m_store.group(1)} days)", False
                    return order_id, postcode, "Warehouse storage", False

                # 5. Out for delivery (S402)
                if status_code in ("S402", "402"):
                    return order_id, postcode, "Customer not yet take", False

                # 6. Return in transit (S500, S510, S512)
                if status_code in ("S500", "S510", "S512", "500", "510", "512"):
                    return order_id, postcode, "Returning to sender", False

                # 7. Other in transit
                clean_desc = desc if len(desc) <= 35 else desc[:32] + "..."
                return order_id, postcode, clean_desc or "In Transit", False
    except Exception as e:
        log.debug("Tracking API check failed for %s: %s", order_id, e)

    return order_id, default_branch, "Pending", False


def track_vip_customer(query: str, date_filter: str = None, days_back: int = 14, max_workers: int = 15, export_file: str = None) -> dict:
    """
    Search and track all orders for a VIP customer by phone or name.
    
    Args:
        query: Phone number (e.g. '060203333') or customer name (e.g. 'lengkimsan')
        date_filter: Specific date 'DD/MM/YYYY' or 'all'
        days_back: Number of recent active days to display if date_filter is None (default 14)
        max_workers: Parallel tracking threads
        export_file: Optional path to pre-downloaded export Excel file
        
    Returns:
        dict with customer info, daily breakdown, message_text, excel_path
    """
    if not export_file or not os.path.exists(export_file):
        export_file = get_latest_export_file()
    if not export_file or not os.path.exists(export_file):
        return {
            "found": False,
            "error": "No export data file found in cache. Please run /push or /schedule run first to download data."
        }

    # Load configuration for API bearer token
    cfg_path = os.path.join(HERE, "config.json")
    bearer_token = ""
    try:
        with open(cfg_path, "r", encoding="utf-8") as f:
            cfg = json.load(f)
            bearer_token = cfg.get("api", {}).get("bearer_token", "")
    except Exception:
        pass

    clean_digits = _clean_phone(query)
    is_phone_query = clean_digits.isdigit() and len(clean_digits) >= 6

    # Load DataFrame
    try:
        df = pd.read_excel(export_file, dtype=str).fillna("")
    except Exception as e:
        return {"found": False, "error": f"Failed reading export file: {e}"}

    if df.empty or "ORDER ID" not in df.columns:
        return {"found": False, "error": "Export data is empty or invalid format."}

    # Filter rows by phone or name
    sender_col = df["SENDER"].astype(str) if "SENDER" in df.columns else pd.Series("", index=df.index)
    receiver_col = df["RECEIVER"].astype(str) if "RECEIVER" in df.columns else pd.Series("", index=df.index)

    if is_phone_query:
        # Match by last 8 digits to handle prefix differences seamlessly (06020333 vs 060203333)
        pattern = clean_digits[-8:] if len(clean_digits) >= 8 else clean_digits
        mask = sender_col.str.contains(pattern, na=False) | receiver_col.str.contains(pattern, na=False)
    else:
        # Match by name (case-insensitive)
        clean_q = query.strip()
        mask = sender_col.str.contains(clean_q, case=False, na=False) | receiver_col.str.contains(clean_q, case=False, na=False)

    matched_df = df[mask].copy()
    if matched_df.empty:
        return {
            "found": False,
            "error": f"No bills found matching '{query}'. Verify phone number or customer name."
        }

    # Extract customer info
    customer_phone = clean_digits if is_phone_query else ""
    customer_name = ""
    customer_address = ""

    # Check known VIP directory first
    vip_group_id = None
    vip_group_title = ""
    for k_phone, prof in KNOWN_VIP_PROFILES.items():
        k_clean = re.sub(r"\D", "", str(k_phone))
        c_clean = clean_digits
        is_phone_match = bool(c_clean and (c_clean in k_clean or k_clean in c_clean or c_clean.lstrip("0") in k_clean or k_clean.lstrip("0") in c_clean))
        is_name_match = bool(query.strip() and (query.lower() in prof["name"].lower() or prof["name"].lower() in query.lower()))
        if is_phone_match or is_name_match:
            customer_name = prof["name"]
            customer_phone = prof["phone"]
            customer_address = prof["address"]
            vip_group_id = prof.get("group_id")
            vip_group_title = prof.get("group_title", "")
            break

    first_row = matched_df.iloc[0]
    if not customer_name:
        sender_val = str(first_row.get("SENDER", "")).strip()
        if " - " in sender_val:
            p_part, n_part = sender_val.split(" - ", 1)
            if not customer_phone:
                customer_phone = _clean_phone(p_part)
            customer_name = n_part.strip()
        elif sender_val:
            customer_name = sender_val

    if not customer_address:
        rec_po = str(first_row.get("RECEIVE POST OFFICE", "")).strip()
        rec_prov = str(first_row.get("RECEIVE PROVINCE", "")).strip()
        if rec_po:
            customer_address = f"{rec_po} ({rec_prov})" if rec_prov else rec_po

    # Group by created date
    date_col = matched_df["CREATED DATE"].astype(str)
    matched_df["date_only"] = date_col.apply(lambda d: str(d).split()[0] if str(d).strip() else "")

    def _parse_d(d_str):
        try:
            return datetime.strptime(d_str, "%d/%m/%Y")
        except Exception:
            return datetime.min

    all_dates = sorted([d for d in matched_df["date_only"].unique() if d], key=_parse_d)

    # Apply date filtering
    if date_filter and date_filter.lower() != "all":
        df_target = date_filter.strip()
        # Normalise single digit day/month e.g. 8/10/2026 -> 08/10/2026
        parts = df_target.split("/")
        if len(parts) == 3:
            df_target = f"{int(parts[0]):02d}/{int(parts[1]):02d}/{parts[2]}"
        target_dates = [d for d in all_dates if d == df_target]
        if not target_dates:
            target_dates = [d for d in all_dates if df_target in d]
    elif date_filter and date_filter.lower() == "all":
        target_dates = all_dates
    else:
        # Default: last N active dates
        target_dates = all_dates[-days_back:] if len(all_dates) > days_back else all_dates

    if not target_dates:
        return {
            "found": False,
            "error": f"No bills found for date '{date_filter}'. Available dates: {', '.join(all_dates[-5:])}"
        }

    # Configure high-capacity session for thread pool
    session = requests.Session()
    adapter = HTTPAdapter(
        pool_connections=max_workers + 5,
        pool_maxsize=max_workers + 5,
        max_retries=Retry(total=1, backoff_factor=0.1)
    )
    session.mount("https://", adapter)
    session.mount("http://", adapter)

    # Parallel status verification
    def _evaluate_row(row):
        oid = str(row["ORDER ID"]).strip()
        exp_status = str(row.get("CURRENT STATUS", "")).strip()
        branch = str(
            row.get("DELIVERY POST OFFICE", "")
            or row.get("POST OFFICE HANDLE", "")
            or row.get("CURRENT POST OFFICE", "")
        ).strip()
        if branch in ("None", "nan", ""):
            branch = ""

        # Extract receiver phone and name
        rec_str = str(row.get("RECEIVER", "")).strip()
        r_phone, r_name = "", ""
        if " - " in rec_str:
            r_phone, r_name = rec_str.split(" - ", 1)
        else:
            r_phone = rec_str

        # If already delivered or returned in export file (terminal state), instant return
        exp_lower = exp_status.lower()
        if "410" in exp_status or "giao thành công" in exp_lower:
            return {
                "order_id": oid,
                "branch": branch,
                "status": "Successful",
                "is_complete": True,
                "receiver_phone": r_phone.strip(),
                "receiver_name": r_name.strip(),
                "cod": str(row.get("COD (USD)", "0.00")),
                "fee": str(row.get("TOTAL FEE (USD) (4)=(1)+(2)-(3)", "0.00")),
                "note": str(row.get("NOTE", "")).strip(),
            }

        if "520" in exp_status or "trả hàng thành công" in exp_lower or "đã trả hàng" in exp_lower:
            return {
                "order_id": oid,
                "branch": branch,
                "status": "Return Successfully",
                "is_complete": True,
                "receiver_phone": r_phone.strip(),
                "receiver_name": r_name.strip(),
                "cod": str(row.get("COD (USD)", "0.00")),
                "fee": str(row.get("TOTAL FEE (USD) (4)=(1)+(2)-(3)", "0.00")),
                "note": str(row.get("NOTE", "")).strip(),
            }

        # Otherwise verify with real-time tracking API
        if bearer_token:
            _, final_branch, final_status, is_succ = _check_bill_tracking(
                oid, bearer_token, default_branch=branch, session=session
            )
        else:
            is_succ = ("410" in exp_status or "520" in exp_status)
            if "520" in exp_status:
                final_status = "Return Successfully"
            elif is_succ:
                final_status = "Successful"
            else:
                final_status = exp_status[:25]
            final_branch = branch

        return {
            "order_id": oid,
            "branch": final_branch,
            "status": final_status,
            "is_complete": is_succ,
            "receiver_phone": r_phone.strip(),
            "receiver_name": r_name.strip(),
            "cod": str(row.get("COD (USD)", "0.00")),
            "fee": str(row.get("TOTAL FEE (USD) (4)=(1)+(2)-(3)", "0.00")),
            "note": str(row.get("NOTE", "")).strip(),
        }

    daily_groups = []
    total_complete = 0
    total_remain = 0

    try:
        for dt in target_dates:
            sub_df = matched_df[matched_df["date_only"] == dt]
            rows = [r for _, r in sub_df.iterrows()]

            with ThreadPoolExecutor(max_workers=max_workers) as executor:
                bill_results = list(executor.map(_evaluate_row, rows))

            c_count = sum(1 for b in bill_results if b["is_complete"])
            r_count = len(bill_results) - c_count

            total_complete += c_count
            total_remain += r_count

            daily_groups.append({
                "date_str": dt,
                "bills": bill_results,
                "complete_count": c_count,
                "remain_count": r_count,
            })
    finally:
        try:
            session.close()
        except Exception:
            pass

    # Build Telegram Formatted Message
    lines = []
    lines.append(f"⭐ *VIP CUSTOMER TRACKING*")
    if customer_name:
        lines.append(f"👤 *Customer:* `{customer_name}`")
    if customer_phone:
        lines.append(f"📞 *Phone:* `{customer_phone}`")
    if customer_address:
        lines.append(f"📍 *Location:* {customer_address}")
    lines.append("━━━━━━━━━━━━━━━━━━━━━━")

    for dg in daily_groups:
        dt = dg["date_str"]
        lines.append(f"📅 *Bills on {dt}*")
        for b in dg["bills"]:
            oid = b["order_id"]
            br = b["branch"]
            st = b["status"]
            br_str = f"`{br}` " if br else ""
            lines.append(f"`{oid}`  {br_str}{st}")
        
        rem_str = "1 Bill" if dg["remain_count"] == 1 else f"{dg['remain_count']} Bills"
        lines.append(f"👉 *Complete {dg['complete_count']}  Remain {rem_str}*\n")

    lines.append("━━━━━━━━━━━━━━━━━━━━━━")
    total_all = total_complete + total_remain
    rem_all_str = "1 Bill" if total_remain == 1 else f"{total_remain} Bills"
    lines.append(f"📊 *TOTAL:* `{total_all}` Bills | ✅ Complete: `{total_complete}` | ⏳ Remain: `{rem_all_str}`")

    message_text = "\n".join(lines)

    # Build Excel Export File
    out_dir = os.path.join(HERE, "cache", "vip_reports")
    os.makedirs(out_dir, exist_ok=True)
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    safe_phone = re.sub(r"\D", "", customer_phone) or "customer"
    excel_path = os.path.join(out_dir, f"VIP_Report_{safe_phone}_{stamp}.xlsx")
    _generate_vip_excel(customer_name, customer_phone, daily_groups, excel_path)

    return {
        "found": True,
        "customer_name": customer_name,
        "customer_phone": customer_phone,
        "customer_address": customer_address,
        "group_id": vip_group_id,
        "group_title": vip_group_title,
        "total_bills": total_all,
        "total_complete": total_complete,
        "total_remain": total_remain,
        "daily_groups": daily_groups,
        "message_text": message_text,
        "excel_path": excel_path,
    }


def _generate_vip_excel(cust_name: str, cust_phone: str, daily_groups: list, out_path: str):
    """Create a styled, professional Excel sheet for the customer."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "VIP Order Report"
    ws.views.sheetView[0].showGridLines = True

    header_font = Font(name="Arial", size=14, bold=True, color="1F497D")
    sub_font = Font(name="Arial", size=10, bold=False, color="595959")
    col_hdr_font = Font(name="Arial", size=10, bold=True, color="FFFFFF")
    data_font = Font(name="Arial", size=10)
    bold_font = Font(name="Arial", size=10, bold=True)

    col_fill = PatternFill(start_color="1F497D", end_color="1F497D", fill_type="solid")
    date_fill = PatternFill(start_color="DCE6F1", end_color="DCE6F1", fill_type="solid")
    succ_fill = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")
    pend_fill = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")

    thin_border = Border(
        left=Side(style='thin', color='D9D9D9'),
        right=Side(style='thin', color='D9D9D9'),
        top=Side(style='thin', color='D9D9D9'),
        bottom=Side(style='thin', color='D9D9D9')
    )

    # Title Banner
    ws["A1"] = "VIP CUSTOMER SHIPMENT REPORT"
    ws["A1"].font = header_font
    ws["A2"] = f"Customer: {cust_name or 'N/A'}  |  Phone: {cust_phone or 'N/A'}  |  Generated: {datetime.now().strftime('%d/%m/%Y %H:%M')}"
    ws["A2"].font = sub_font

    headers = ["Date", "Order ID", "Delivery Branch", "Status", "Receiver Phone", "Receiver Name", "COD ($)", "Fee ($)", "Note"]
    start_row = 4
    for c_idx, h in enumerate(headers, 1):
        cell = ws.cell(row=start_row, column=c_idx, value=h)
        cell.font = col_hdr_font
        cell.fill = col_fill
        cell.alignment = Alignment(horizontal="center", vertical="center")

    curr_row = start_row + 1

    for dg in daily_groups:
        dt = dg["date_str"]
        c_count = dg["complete_count"]
        r_count = dg["remain_count"]

        # Date subheader row
        d_cell = ws.cell(row=curr_row, column=1, value=f"Bills on {dt} (Complete: {c_count} | Remain: {r_count})")
        d_cell.font = bold_font
        d_cell.fill = date_fill
        ws.merge_cells(start_row=curr_row, start_column=1, end_row=curr_row, end_column=len(headers))
        curr_row += 1

        for b in dg["bills"]:
            clean_oid = re.sub(r"\.0$", "", str(b["order_id"])).strip()
            clean_rphone = re.sub(r"\.0$", "", str(b["receiver_phone"])).strip()
            row_data = [
                dt,
                clean_oid,
                b["branch"],
                b["status"],
                clean_rphone,
                b["receiver_name"],
                float(b["cod"]) if b["cod"].replace('.', '', 1).isdigit() else b["cod"],
                float(b["fee"]) if b["fee"].replace('.', '', 1).isdigit() else b["fee"],
                b["note"]
            ]

            for c_idx, val in enumerate(row_data, 1):
                c = ws.cell(row=curr_row, column=c_idx, value=val)
                c.font = data_font
                if c_idx in (1, 2, 3, 5):
                    c.number_format = '@'
                    c.alignment = Alignment(horizontal="center")
                elif c_idx in (7, 8):
                    c.number_format = '#,##0.00'
                c.border = thin_border
                c.alignment = Alignment(vertical="center")
                if c_idx in (1, 2, 3):
                    c.alignment = Alignment(horizontal="center", vertical="center")
                if c_idx == 4:
                    c.alignment = Alignment(horizontal="center", vertical="center")
                    if b["is_complete"]:
                        c.fill = succ_fill
                    else:
                        c.fill = pend_fill
            curr_row += 1

    # Auto-fit columns
    for col in ws.columns:
        max_len = max(len(str(cell.value or '')) for cell in col)
        col_letter = get_column_letter(col[0].column)
        ws.column_dimensions[col_letter].width = max(max_len + 3, 12)

    wb.save(out_path)
