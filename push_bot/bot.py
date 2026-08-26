"""
bot.py — Push Bot with pause/resume/register/status commands.

Commands:
  push      — fetch data and send reports
  /test     — fetch data and send reports only to the requester
  /pause    — pause forwarding to groups (bot still works for you)
  /resume   — resume forwarding to groups
  /status   — show current state
  /mode     — toggle wide/long image mode
  /register — register current group to receive forwards
  /groups   — list registered groups
"""

import json
import logging
import os
import re
import sys
import tempfile
import asyncio
import io
from datetime import datetime, timedelta
import threading
import pandas as pd
from http.server import HTTPServer, BaseHTTPRequestHandler

from telegram import Update
from telegram.ext import (Application, MessageHandler, CommandHandler,
                          ContextTypes, filters)
from telegram.error import RetryAfter, NetworkError, Forbidden

class WebAppHandler(BaseHTTPRequestHandler):
    def do_HEAD(self):
        self.send_response(200)
        self.send_header('Content-type', 'text/html; charset=utf-8')
        self.end_headers()

    def do_GET(self):
        if self.path == '/' or self.path.startswith('/?') or self.path == '/healthz':
            self.send_response(200)
            self.send_header('Content-type', 'text/html; charset=utf-8')
            self.end_headers()
            try:
                # Read index.html from same directory
                with open(os.path.join(HERE, "index.html"), "rb") as f:
                    self.wfile.write(f.read())
            except Exception as e:
                self.wfile.write(f"Error loading index.html: {e}".encode('utf-8'))
        elif self.path == '/data':
            self.send_response(200)
            self.send_header('Content-type', 'application/json; charset=utf-8')
            self.send_header('Access-Control-Allow-Origin', '*')
            self.end_headers()
            data_path = os.path.join(HERE, "cache", "latest_data.json")
            if os.path.exists(data_path):
                with open(data_path, "rb") as f:
                    self.wfile.write(f.read())
            else:
                self.wfile.write(b"[]")
        elif self.path == '/dashboard-data':
            # Require valid token in Authorization header
            auth = self.headers.get('Authorization', '')
            token = auth.replace('Bearer ', '').strip()
            if not self._is_token_valid(token):
                self.send_response(401)
                self.end_headers()
                self.wfile.write(b'{"error":"Unauthorized"}')
                return
            self.send_response(200)
            self.send_header('Content-type', 'application/json; charset=utf-8')
            self.send_header('Access-Control-Allow-Origin', '*')
            self.end_headers()
            data_path = os.path.join(HERE, "cache", "dashboard_data.json")
            if os.path.exists(data_path):
                with open(data_path, "rb") as f:
                    self.wfile.write(f.read())
            else:
                self.wfile.write(b'{"branches":{},"summary":{},"updated":null}')
        elif self.path.startswith('/images/'):
            filename = os.path.basename(self.path.split('?')[0])
            file_path = os.path.join(HERE, "images", filename)
            if os.path.exists(file_path) and filename.lower().endswith(('.png', '.jpg', '.jpeg', '.svg')):
                self.send_response(200)
                self.send_header('Content-type', 'image/png')
                self.send_header('Cache-Control', 'max-age=86400')
                self.end_headers()
                with open(file_path, "rb") as f:
                    self.wfile.write(f.read())
            else:
                self.send_response(404)
                self.end_headers()
                self.wfile.write(b"Image Not Found")
        elif self.path == '/dashboard-tokens':
            self.send_response(404)
            self.end_headers()
            self.wfile.write(b"Not Found")
        else:
            self.send_response(404)
            self.end_headers()
            self.wfile.write(b"Not Found")

    def log_message(self, format, *args):
        # Suppress logging every single asset request to avoid cluttering the bot terminal logs
        pass

    def _is_token_valid(self, token):
        """Check if token exists and is not expired."""
        if not token:
            return False
        token_path = os.path.join(HERE, "dashboard_tokens.json")
        try:
            if not os.path.exists(token_path):
                return False
            with open(token_path, encoding="utf-8") as f:
                data = json.load(f)
            now = datetime.now()
            for t in data.get("tokens", []):
                if t["token"] == token:
                    expires = datetime.fromisoformat(t["expires"])
                    if now < expires:
                        return True
        except Exception:
            pass
        return False

def start_webapp_server():
    try:
        port = int(os.environ.get('PORT', 8080))
        server = HTTPServer(('0.0.0.0', port), WebAppHandler)
        log.info(f"Telegram WebApp Server listening on port {port}...")
        server.serve_forever()
    except Exception as e:
        log.error(f"Failed to start Telegram WebApp Server: {e}")

def update_webapp_cache(result):
    items = []
    try:
        type_data = result.get("type_data", {})
        for sheet_name, df in type_data.items():
            if df is None or df.empty:
                continue
            
            for _, row in df.iterrows():
                order_id = str(row.get("ORDER ID") or row.get("Order ID") or "").strip()
                if not order_id or order_id.lower() == "nan":
                    continue
                
                province = str(row.get("PROVINCE") or row.get("Province") or "").strip()
                district = str(row.get("DISTRICT") or row.get("District") or "").strip()
                store = str(row.get("POST OFFICE HANDLE") or row.get("Post Office") or row.get("CURRENT STATUS") or "").strip()
                address = str(row.get("DELIVERY ADDRESS") or row.get("Address") or row.get("RECEIVER ADDRESS") or "").strip()
                
                if province.lower() == "nan": province = ""
                if district.lower() == "nan": district = ""
                if store.lower() == "nan": store = ""
                if address.lower() == "nan": address = ""
                
                items.append({
                    "id": order_id,
                    "province": province,
                    "district": district,
                    "store": store,
                    "address": address
                })
        
        cache_dir = os.path.join(HERE, "cache")
        os.makedirs(cache_dir, exist_ok=True)
        cache_path = os.path.join(cache_dir, "latest_data.json")
        with open(cache_path, "w", encoding="utf-8") as f:
            json.dump(items, f, ensure_ascii=False, indent=2)
        log.info(f"Updated WebApp data cache with {len(items)} items.")
    except Exception as e:
        log.error(f"Failed to update WebApp data cache: {e}")

def update_dashboard_cache(result):
    """Build enriched JSON for the web dashboard with manager-friendly views.

    Views per branch:
      - all_pending: ALL bills at this PO that are NOT completed (Pickup+Delivery+Transit+Branch combined)
      - completed: Bills completed today (410 + 520 + 201) at this PO
      - incoming: Bills heading TO this PO from elsewhere
      - total_today: ALL bills this PO handled today (pending + completed = full workload)
    """
    try:
        type_data = result.get("type_data", {})
        dm_all = result.get("dm_all")  # Full data including completed
        today = datetime.now().date()
        branches = {}

        def _extract_order(row):
            order_id = str(row.get("ORDER ID") or "").strip()
            if not order_id or order_id.lower() == "nan":
                return None
            receiver = str(row.get("RECEIVER") or row.get("Cus name") or "").strip()
            if receiver.lower() == "nan": receiver = ""
            status_code = str(row.get("STATUS_CODE") or "").strip()
            if status_code.lower() == "nan": status_code = ""
            phone = str(row.get("Phone") or row.get("RECEIVER PHONE") or row.get("PHONE") or "").strip()
            if phone.lower() == "nan": phone = ""
            fee = 0.0
            try: fee = float(row.get("TOTAL FEE (USD)") or 0)
            except (ValueError, TypeError): pass
            cod = 0.0
            try: cod = float(row.get("COD (USD)") or 0)
            except (ValueError, TypeError): pass
            created_date = ""
            cd_val = row.get("CREATED DATE")
            if cd_val and str(cd_val).lower() != "nan":
                try:
                    cd = pd.to_datetime(cd_val, dayfirst=True, format="mixed", errors="coerce")
                    if not pd.isna(cd):
                        created_date = cd.strftime("%d/%m/%Y")
                except Exception: pass
            return {
                "order_id": order_id,
                "receiver": receiver,
                "phone": phone,
                "status_code": status_code,
                "category": _get_category(status_code),
                "stage": _get_stage(status_code),
                "fee": fee,
                "cod": cod,
                "created_date": created_date,
            }

        def _get_category(sc):
            """Categorize for filter checkboxes (shipped/cancelled/return/active)."""
            if sc in ('410',):
                return 'shipped'
            if sc in ('201',):
                return 'cancelled'
            if sc in ('500', '510', '511', '512', '520', '540'):
                return 'return'
            return 'active'

        def _get_stage(sc):
            """Categorize for tab views (pickup/delivery/completed/transit)."""
            if sc in ('110', '120', '200'):
                return 'pickup'
            if sc in ('410', '520', '201'):
                return 'completed'
            if sc in ('401', '402', '420', '430', '460', '472', '480', '400', '306', '309'):
                return 'delivery'
            # Transit/pending: 210, 300, 302, 310, 311, 500, 510, 511, 512, 540
            return 'transit'

        def _ensure_branch(b):
            if b not in branches:
                branches[b] = {
                    "all_pending": [],    # All pending at this PO (all status types)
                    "completed": [],      # Completed today (410/520/201)
                    "incoming": [],       # Heading to this PO from elsewhere
                    "total_today": [],    # Everything (pending + completed)
                    "counts": {
                        "all_pending": 0,
                        "completed": 0,
                        "incoming": 0,
                        "total_today": 0,
                    },
                    "total_fee": 0.0,
                    "total_cod": 0.0,
                }

        # ── 1. All Pending at Current PO (all report types: Pickup+Delivery+Transit+Branch) ──
        seen_pending_ids = {}  # track per PO to avoid duplicates
        for report_type in ['Pickup', 'Delivery', 'Transit', 'Branch']:
            df = type_data.get(report_type)
            if df is None or df.empty:
                continue
            if 'CURRENT POST OFFICE' not in df.columns:
                continue
            for po, df_po in df.groupby('CURRENT POST OFFICE'):
                po_str = str(po).strip().upper()
                if not po_str:
                    continue
                _ensure_branch(po_str)
                if po_str not in seen_pending_ids:
                    seen_pending_ids[po_str] = set()
                for _, row in df_po.iterrows():
                    order = _extract_order(row)
                    if not order:
                        continue
                    if order["order_id"] in seen_pending_ids[po_str]:
                        continue
                    seen_pending_ids[po_str].add(order["order_id"])
                    order["report_type"] = report_type
                    branches[po_str]["all_pending"].append(order)
                    branches[po_str]["total_today"].append(order)
                    branches[po_str]["total_fee"] += order["fee"]
                    branches[po_str]["total_cod"] += order["cod"]

        # ── 2. Incoming to PO (RECEIVE POST OFFICE = this PO, but CURRENT != this PO) ──
        for report_type in ['Pickup', 'Delivery', 'Transit', 'Branch']:
            df = type_data.get(report_type)
            if df is None or df.empty:
                continue
            if 'RECEIVE POST OFFICE' not in df.columns or 'CURRENT POST OFFICE' not in df.columns:
                continue
            for _, row in df.iterrows():
                recv_po = str(row.get("RECEIVE POST OFFICE") or "").strip().upper()
                curr_po = str(row.get("CURRENT POST OFFICE") or "").strip().upper()
                if not recv_po or recv_po == curr_po:
                    continue
                _ensure_branch(recv_po)
                order = _extract_order(row)
                if not order:
                    continue
                order["report_type"] = report_type
                order["current_po"] = curr_po
                branches[recv_po]["incoming"].append(order)

        # ── 3. Completed Today (410 + 520 + 201 from dm_all) ──
        if dm_all is not None and not dm_all.empty:
            completed_mask = dm_all['STATUS_CODE'].isin(['410', '520', '201'])
            df_completed = dm_all[completed_mask].copy()
            if 'CURRENT POST OFFICE' in df_completed.columns:
                for po, df_po in df_completed.groupby('CURRENT POST OFFICE'):
                    po_str = str(po).strip().upper()
                    if not po_str:
                        continue
                    _ensure_branch(po_str)
                    for _, row in df_po.iterrows():
                        order = _extract_order(row)
                        if not order:
                            continue
                        order["report_type"] = "Completed"
                        branches[po_str]["completed"].append(order)
                        branches[po_str]["total_today"].append(order)

        # ── Build counts ──
        for b, b_data in branches.items():
            b_data["counts"]["all_pending"] = len(b_data["all_pending"])
            b_data["counts"]["completed"] = len(b_data["completed"])
            b_data["counts"]["incoming"] = len(b_data["incoming"])
            b_data["counts"]["total_today"] = len(b_data["total_today"])

        # ── Overall summary ──
        summary = {"all_pending": 0, "completed": 0, "incoming": 0, "total_today": 0}
        for b_data in branches.values():
            for k in summary:
                summary[k] += b_data["counts"].get(k, 0)

        dashboard_data = {
            "branches": branches,
            "summary": summary,
            "updated": datetime.now().strftime("%d/%m/%Y %H:%M"),
        }

        cache_dir = os.path.join(HERE, "cache")
        os.makedirs(cache_dir, exist_ok=True)
        cache_path = os.path.join(cache_dir, "dashboard_data.json")
        with open(cache_path, "w", encoding="utf-8") as f:
            json.dump(dashboard_data, f, ensure_ascii=False)
        log.info(f"Updated dashboard cache: {len(branches)} branches, "
                 f"summary={summary}")
    except Exception as e:
        log.error(f"Failed to update dashboard cache: {e}")


def save_highlight_history(result):
    try:
        today_date = datetime.now().date()
        today_str = today_date.strftime("%Y-%m-%d")
        timestamp_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        run_data = {"Pickup": {}, "Delivery": {}, "Transit": {}, "Branch": {}}
        type_data = result.get("type_data", {})
        filter_cols = {
            'Pickup': 'POST OFFICE HANDLE',
            'Delivery': 'POST OFFICE HANDLE',
            'Transit': 'POST OFFICE HANDLE',
            'Branch': 'POST OFFICE HANDLE'
        }

        for rn in ['Pickup', 'Delivery', 'Transit', 'Branch']:
            df = type_data.get(rn)
            if df is None or df.empty:
                continue

            fcol = filter_cols[rn]
            if fcol not in df.columns:
                continue

            for h, df_h in df.groupby(fcol):
                h_str = str(h).strip().upper()
                if not h_str:
                    continue
                
                oids = []
                for _, row in df_h.iterrows():
                    oid = str(row.get("ORDER ID") or "").strip()
                    if not oid or oid.lower() == "nan":
                        continue
                    
                    sc = str(row.get("STATUS_CODE") or "").strip()
                    cd_val = row.get("CREATED DATE")
                    
                    is_highlight = False
                    if sc in ('420', '472'):
                        if cd_val:
                            cd = pd.to_datetime(cd_val, dayfirst=True, format="mixed", errors="coerce")
                            if not pd.isna(cd):
                                if (today_date - cd.date()).days > 7:
                                    is_highlight = True
                    else:
                        if cd_val:
                            cd = pd.to_datetime(cd_val, dayfirst=True, format="mixed", errors="coerce")
                            if not pd.isna(cd):
                                if (today_date - cd.date()).days > 1:
                                    is_highlight = True
                    
                    if is_highlight:
                        oids.append(oid)
                
                if oids:
                    run_data[rn][h_str] = oids

        history_path = os.path.join(HERE, "highlight_history.json")
        history = {}
        if os.path.exists(history_path):
            try:
                with open(history_path, encoding="utf-8") as f:
                    history = json.load(f)
            except Exception:
                pass

        # Keep last 7 days of history
        history = {k: v for k, v in history.items() if (datetime.now() - datetime.strptime(k, "%Y-%m-%d")).days < 7}

        if today_str not in history:
            history[today_str] = []

        history[today_str].append({
            "timestamp": timestamp_str,
            "data": run_data
        })

        with open(history_path, "w", encoding="utf-8") as f:
            json.dump(history, f, ensure_ascii=False, indent=2)

        log.info(f"Saved highlighted orders history to highlight_history.json at {timestamp_str}")
    except Exception as e:
        log.warning(f"Failed to save highlight history: {e}")

import warnings
warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')

import downloader
import generate_report
import generate_summary
import excel_to_image
import report_cmd

HERE = os.path.dirname(os.path.abspath(__file__))
CONFIG_PATH        = os.path.join(HERE, "config.json")
REF_PATH           = os.path.join(HERE, "post_office_lookup.csv")
PICKUP_BRANCH_LOOKUP_PATH = os.path.join(HERE, "pickup_branch_lookup.csv")
REGISTERED_GROUPS_PATH = os.path.join(HERE, "registered_groups.json")
REPORTS_LOG_PATH   = os.path.join(HERE, "reports_today.json")

from logging.handlers import RotatingFileHandler
log_file_path = os.path.join(HERE, "bot.log")
file_handler = RotatingFileHandler(
    log_file_path,
    maxBytes=5 * 1024 * 1024,
    backupCount=3,
    encoding="utf-8"
)

logging.basicConfig(
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
    level=logging.INFO,
    handlers=[
        logging.StreamHandler(sys.stdout),
        file_handler
    ]
)
logging.getLogger("httpx").setLevel(logging.WARNING)
logging.getLogger("httpcore").setLevel(logging.WARNING)
log = logging.getLogger("push_bot")



# ── Today's report tracker (for /clean) ───────────────────────────────────────

def track_report_dir(tmpdir: str):
    """Record a tempdir so /clean can delete it later."""
    today_str = datetime.now().strftime("%Y-%m-%d")
    try:
        data = {}
        if os.path.exists(REPORTS_LOG_PATH) and os.path.getsize(REPORTS_LOG_PATH) > 0:
            with open(REPORTS_LOG_PATH, encoding="utf-8") as f:
                try:
                    data = json.load(f)
                except Exception:
                    data = {}
        if not isinstance(data, dict):
            data = {}
        # Prune old days (keep only today)
        data = {k: v for k, v in data.items() if k == today_str}
        if today_str not in data:
            data[today_str] = []
        if tmpdir not in data[today_str]:
            data[today_str].append(tmpdir)
        with open(REPORTS_LOG_PATH, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
    except Exception as e:
        log.warning(f"track_report_dir failed: {e}")


from functools import wraps

class PrivateMessageRequired(Exception):
    """Raised when a command cannot reply privately because the user hasn't started the bot in PM."""
    pass

def pm_required_handler(func):
    @wraps(func)
    async def wrapper(update: Update, context: ContextTypes.DEFAULT_TYPE, *args, **kwargs):
        # ── STRICT GROUP LOCK: Block ALL commands in group chats ────────────────
        if is_group_chat(update):
            await delete_group_command(update, context)
            log.info("Blocked command '%s' in group chat %s", func.__name__, update.effective_chat.id)
            return

        # ── User allowlist check ───────────────────────────────────────────────
        try:
            cfg = load_config()
            allowed_users = cfg["telegram"].get("allowed_user_ids") or []
            if allowed_users:
                user = update.effective_user
                if not user or user.id not in allowed_users:
                    log.info("Ignoring command from unauthorized user %s", user and user.id)
                    return
        except Exception:
            pass  # If config fails, allow through
        # ── Run handler ───────────────────────────────────────────────────────
        try:
            return await func(update, context, *args, **kwargs)
        except PrivateMessageRequired:
            log.info("Command aborted because user has not started PM with the bot.")
            return
    return wrapper


def user_guard(func):
    """Standalone decorator for commands NOT decorated with pm_required_handler.
    Blocks group execution and enforces allowlist."""
    @wraps(func)
    async def wrapper(update: Update, context: ContextTypes.DEFAULT_TYPE, *args, **kwargs):
        if is_group_chat(update) and func.__name__ not in ("cmd_delete_report",):
            await delete_group_command(update, context)
            log.info("Blocked user_guard command '%s' in group chat %s", func.__name__, update.effective_chat.id)
            return

        try:
            cfg = load_config()
            allowed_users = cfg["telegram"].get("allowed_user_ids") or []
            if allowed_users:
                user = update.effective_user
                if not user or user.id not in allowed_users:
                    log.info("Ignoring command from unauthorized user %s", user and user.id)
                    return
        except Exception:
            pass
        return await func(update, context, *args, **kwargs)
    return wrapper



# ── Config helpers ─────────────────────────────────────────────────────────────

def load_config():
    with open(CONFIG_PATH, encoding="utf-8") as f:
        return json.load(f)


def save_config(cfg):
    with open(CONFIG_PATH, "w", encoding="utf-8") as f:
        json.dump(cfg, f, ensure_ascii=False, indent=2)


def load_registered_groups():
    if not os.path.exists(REGISTERED_GROUPS_PATH):
        return []
    with open(REGISTERED_GROUPS_PATH, encoding="utf-8") as f:
        return json.load(f)


def save_registered_groups(groups):
    with open(REGISTERED_GROUPS_PATH, "w", encoding="utf-8") as f:
        json.dump(groups, f, ensure_ascii=False, indent=2)


_MIGRATION_PATTERN = re.compile(r"new chat id:\s*(-?\d+)", re.IGNORECASE)


def _extract_migrated_chat_id(error_msg: str):
    """Return new chat id from a 'Group migrated to supergroup' error, or None."""
    match = _MIGRATION_PATTERN.search(error_msg or "")
    if not match:
        return None
    try:
        return int(match.group(1))
    except (TypeError, ValueError):
        return None


def _replace_chat_id_in_args(args, kwargs, new_chat_id):
    """Return (args, kwargs) with chat_id replaced by new_chat_id."""
    new_args = tuple(new_chat_id if isinstance(a, int) else a for a in args)
    new_kwargs = dict(kwargs)
    if "chat_id" in new_kwargs:
        new_kwargs["chat_id"] = new_chat_id
    return new_args, new_kwargs


def _reset_io_buffers(args, kwargs):
    """Seek any file-like arguments back to position 0 before retrying.

    Telegram reads uploaded files during the request, which moves the
    BytesIO cursor to the end. Without resetting, retries would send
    empty data and trigger "File must be non-empty" errors.
    """
    candidates = list(args) + list(kwargs.values())

    def _maybe_reset(obj):
        if isinstance(obj, io.IOBase):
            try:
                obj.seek(0)
            except Exception:
                pass

    for obj in candidates:
        _maybe_reset(obj)


def _update_migrated_chat_id(old_chat_id, new_chat_id):
    """Persist chat_id migration to registered_groups.json and config.json."""
    old_str = str(old_chat_id)
    new_str = str(new_chat_id)

    # 1. registered_groups.json — replace old entry with new entry (dedup)
    try:
        groups = load_registered_groups()
        old_entry = None
        for g in groups:
            if str(g.get("chat_id")) == old_str:
                old_entry = dict(g)
                break
        if old_entry is not None:
            groups = [g for g in groups if str(g.get("chat_id")) != old_str]
            already_has_new = any(str(g.get("chat_id")) == new_str for g in groups)
            if not already_has_new:
                old_entry["chat_id"] = new_chat_id
                groups.append(old_entry)
            save_registered_groups(groups)
            log.info("Updated registered_groups.json: %s -> %s", old_str, new_str)
    except Exception as e:
        log.warning("Could not update registered_groups.json: %s", e)

    # 2. config.json — forward_mapping + zone_forward_mapping
    try:
        cfg = load_config()
        changed = False
        fwd = cfg.get("telegram", {}).get("forward_mapping", {})
        if old_str in fwd:
            if new_str not in fwd:
                fwd[new_str] = fwd[old_str]
            del fwd[old_str]
            cfg["telegram"]["forward_mapping"] = fwd
            changed = True
        zfwd = cfg.get("zone_forward_mapping", {})
        if old_str in zfwd:
            if new_str not in zfwd:
                zfwd[new_str] = zfwd[old_str]
            del zfwd[old_str]
            cfg["zone_forward_mapping"] = zfwd
            changed = True
        if changed:
            save_config(cfg)
            log.info("Updated config.json: %s -> %s", old_str, new_str)
    except Exception as e:
        log.warning("Could not update config.json: %s", e)


async def safe_api_call(func, *args, **kwargs):
    max_retries = 5
    for attempt in range(max_retries):
        try:
            return await func(*args, **kwargs)
        except RetryAfter as e:
            wait_time = e.retry_after
            log.warning(f"Flood control exceeded. Waiting for {wait_time} seconds before retrying (attempt {attempt + 1})...")
            await asyncio.sleep(wait_time + 1)
        except NetworkError as e:
            log.warning(f"Network error: {e}. Retrying in 3 seconds (attempt {attempt + 1})...")
            await asyncio.sleep(3)
            _reset_io_buffers(args, kwargs)
        except Forbidden as e:
            raise e
        except Exception as e:
            error_msg = str(e)
            new_chat_id = _extract_migrated_chat_id(error_msg)
            if new_chat_id is not None:
                old_chat_id = kwargs.get("chat_id")
                if old_chat_id is None:
                    for a in args:
                        if isinstance(a, int):
                            old_chat_id = a
                            break
                if old_chat_id is not None and new_chat_id != old_chat_id:
                    log.warning(
                        "Group %s migrated to supergroup %s. Updating storage and retrying...",
                        old_chat_id, new_chat_id,
                    )
                    _update_migrated_chat_id(old_chat_id, new_chat_id)
                    args, kwargs = _replace_chat_id_in_args(args, kwargs, new_chat_id)
                    _reset_io_buffers(args, kwargs)
                    # Don't count migration handling as a retry attempt
                    continue
            if attempt == max_retries - 1:
                raise e
            log.warning(f"Error calling Telegram API: {e}. Retrying in 2 seconds...")
            await asyncio.sleep(2)
            _reset_io_buffers(args, kwargs)


import re

def get_forward_mapping(cfg):
    # Groups configured manually in config.json
    mapping = dict(cfg["telegram"].get("forward_mapping", {}))
    
    # Auto-detect from registered group titles!
    for g in load_registered_groups():
        chat_id_str = str(g["chat_id"])
        title = g.get("title", "")
        # Look for handles like PNPP014, SVAP001, KANP001 in the title
        found_handles = re.findall(r'\b[A-Z]{3}P\d{3}\b', title.upper())
        
        # If no full handles found, try matching 3-letter branch codes (like SIH, KOH)
        if not found_handles:
            words = re.findall(r'\b[A-Z]{3}\b', title.upper())
            for w in words:
                known_prefixes = []
                for zb in cfg.get("zone_branches", {}).values():
                    known_prefixes.extend([p.strip().upper() for p in zb.split(",") if p.strip()])
                
                if w in known_prefixes:
                    found_handles.append(f"{w}P001")
        
        if found_handles:
            # If the group is already in mapping, don't overwrite manual settings
            if chat_id_str not in mapping:
                mapping[chat_id_str] = found_handles
                
    return mapping


def get_all_forward_groups(cfg):
    """Merge config groups + registered groups, excluding zone groups."""
    mapping = get_forward_mapping(cfg)
    configured  = list(mapping.keys())
    registered  = [str(g["chat_id"]) for g in load_registered_groups()]
    all_groups = list(dict.fromkeys(configured + registered))
    
    # Exclude zone groups from regular forwarding (they receive zone reports instead)
    zone_groups = set(str(k) for k in cfg.get("zone_forward_mapping", {}).keys())
    return [g for g in all_groups if g not in zone_groups]


def is_paused(cfg):
    return bool(cfg["telegram"].get("paused", False))


def get_mode(cfg):
    return cfg["telegram"].get("image_mode", "long")


def is_group_chat(update: Update):
    chat = update.effective_chat
    return bool(chat and chat.type in ("group", "supergroup"))


def is_user_allowed(update: Update, cfg: dict) -> bool:
    """Return True if the user is allowed to use the bot.
    If allowed_user_ids is empty, all users are allowed.
    """
    allowed = cfg["telegram"].get("allowed_user_ids") or []
    if not allowed:
        return True
    user = update.effective_user
    return bool(user and user.id in allowed)


def requester_chat_id(update: Update):
    if is_group_chat(update):
        user = update.effective_user
        return user.id if user else None

    chat = update.effective_chat
    return chat.id if chat else None


async def delete_group_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Best-effort removal of control commands typed in groups."""
    if not is_group_chat(update) or not update.message:
        return
    try:
        await context.bot.delete_message(
            chat_id=update.effective_chat.id,
            message_id=update.message.message_id,
        )
    except Exception as e:
        log.info("Could not delete group command message: %s", e)


async def private_or_current_reply(
    update: Update,
    context: ContextTypes.DEFAULT_TYPE,
    text: str,
    parse_mode: str = None,
):
    """Send command feedback privately when invoked from a group."""
    msg = await send_requester_text(update, context, text, parse_mode=parse_mode)
    return bool(msg)


async def send_requester_text(
    update: Update,
    context: ContextTypes.DEFAULT_TYPE,
    text: str,
    parse_mode: str = None,
):
    chat_id = requester_chat_id(update)
    if chat_id is None:
        log.warning("Cannot send requester text without a chat id.")
        return None

    try:
        # Try to send privately to user
        return await safe_api_call(
            context.bot.send_message, chat_id=chat_id, text=text, parse_mode=parse_mode
        )
    except Exception as e:
        log.warning("Could not send requester text privately to %s: %s", chat_id, e)
        # If in a group chat, fallback to sending directly in the group
        if is_group_chat(update) and update.effective_chat:
            try:
                return await safe_api_call(
                    context.bot.send_message,
                    chat_id=update.effective_chat.id,
                    text=text,
                    parse_mode=parse_mode,
                )
            except Exception as e2:
                log.warning("Failed to send fallback requester text to group: %s", e2)
        return None



async def edit_or_send_requester_text(
    message,
    update: Update,
    context: ContextTypes.DEFAULT_TYPE,
    text: str,
    parse_mode: str = None,
):
    if message:
        try:
            await safe_api_call(message.edit_text, text, parse_mode=parse_mode)
            return message
        except Exception as e:
            log.warning("Could not edit requester status message: %s", e)

    return await send_requester_text(update, context, text, parse_mode=parse_mode)


async def send_requester_photo(update: Update, context: ContextTypes.DEFAULT_TYPE, photo):
    chat_id = requester_chat_id(update)
    if chat_id is None:
        log.warning("Cannot send requester photo without a chat id.")
        return False

    # Prepare buffer copy for potential document fallback if send_photo fails
    photo_copy = None
    if hasattr(photo, "getvalue"):
        import io
        photo_copy = io.BytesIO(photo.getvalue())

    try:
        await safe_api_call(context.bot.send_photo, chat_id=chat_id, photo=photo)
        return True
    except Exception as e:
        log.warning("Could not send requester photo to %s: %s (trying send_document fallback)", chat_id, e)
        try:
            doc_buf = photo_copy if photo_copy else photo
            doc_name = getattr(photo, "name", "report_image.png") or "report_image.png"
            await safe_api_call(context.bot.send_document, chat_id=chat_id, document=doc_buf, filename=doc_name)
            return True
        except Exception as e2:
            log.warning("Fallback send document failed: %s", e2)
            return False


async def send_requester_document(
    update: Update,
    context: ContextTypes.DEFAULT_TYPE,
    document,
    filename: str,
    caption: str = None,
):
    chat_id = requester_chat_id(update)
    if chat_id is None:
        log.warning("Cannot send requester document without a chat id.")
        return False

    try:
        await safe_api_call(
            context.bot.send_document,
            chat_id=chat_id,
            document=document,
            filename=filename,
            caption=caption,
        )
        return True
    except Exception as e:
        log.warning("Could not send requester document to %s: %s", chat_id, e)
        if is_group_chat(update) and update.effective_chat:
            try:
                document.seek(0)
                await safe_api_call(
                    context.bot.send_document,
                    chat_id=update.effective_chat.id,
                    document=document,
                    filename=filename,
                    caption=caption,
                )
                return True
            except Exception as e2:
                log.warning("Fallback send document to group failed: %s", e2)
        return False



async def forward_result_to_groups(context: ContextTypes.DEFAULT_TYPE, payload):
    result = payload["result"]
    forward_groups = payload["forward_groups"]
    forward_mapping = payload["forward_mapping"]
    sent_groups = 0

    for group_id_str in forward_groups:
        try:
            group_id = int(group_id_str)
        except ValueError:
            group_id = group_id_str  # fallback if it's a string like "@channel"

        allowed_handles = forward_mapping.get(group_id_str, ["*"])
        wants_all = "*" in allowed_handles
        sent_any = False

        for hr in result["handle_results"]:
            handle = hr["handle"]
            if not wants_all and handle not in allowed_handles:
                continue

            for hf in hr["handle_files"]:
                try:
                    img_buf = excel_to_image.excel_to_image(hf["path"])
                    img_buf.name = f"{handle}.png"
                    await safe_api_call(context.bot.send_photo, chat_id=group_id, photo=img_buf)
                    sent_any = True
                    await asyncio.sleep(0.5)
                except Exception as e:
                    log.error(f"Image to group {group_id}: {e}")

            # Send the single consolidated Excel file for this branch containing all tabs (Pickup, Delivery, Transit, Branch)
            excel_path = hr.get("handle_excel_path")
            if excel_path and os.path.exists(excel_path):
                try:
                    with open(excel_path, "rb") as ef:
                        await safe_api_call(
                            context.bot.send_document,
                            chat_id=group_id,
                            document=ef,
                            filename=os.path.basename(excel_path),
                        )
                        sent_any = True
                        await asyncio.sleep(0.5)
                except Exception as e:
                    log.error(f"Excel file to group {group_id} for {handle}: {e}")

            try:
                await safe_api_call(context.bot.send_message, chat_id=group_id, text=hr["remark"])
                sent_any = True
                await asyncio.sleep(0.5)
            except Exception as e:
                log.error(f"Remark to group {group_id}: {e}")

        if wants_all:
            try:
                with open(result["final_xlsx"], "rb") as f:
                    await safe_api_call(
                        context.bot.send_document,
                        chat_id=group_id,
                        document=f,
                        filename=os.path.basename(result["final_xlsx"]),
                    )
                    sent_any = True
                    await asyncio.sleep(0.5)
            except Exception as e:
                log.error(f"Excel to group {group_id}: {e}")
            try:
                await safe_api_call(
                    context.bot.send_message,
                    chat_id=group_id,
                    text=result["summary_caption"],
                )
                sent_any = True
                await asyncio.sleep(0.5)
            except Exception as e:
                log.error(f"Summary to group {group_id}: {e}")

        if sent_any:
            sent_groups += 1
            await asyncio.sleep(1.0)

    return sent_groups


# ── Commands ───────────────────────────────────────────────────────────────────

@user_guard
async def cmd_pause(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Stop forwarding to groups. Bot will NEVER send auto reports to groups."""
    await delete_group_command(update, context)
    args = [a.strip().lower() for a in (context.args or []) if a.strip()]
    cfg = load_config()
    cfg["telegram"]["paused"] = True
    save_config(cfg)
    if "thean" in args:
        await private_or_current_reply(
            update,
            context,
            "🛑 Bot PAUSED (Forwarding disabled)."
        )


@user_guard
async def cmd_resume(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Re-enable forwarding to groups."""
    await delete_group_command(update, context)
    args = [a.strip().lower() for a in (context.args or []) if a.strip()]
    cfg = load_config()
    cfg["telegram"]["paused"] = False
    save_config(cfg)
    if "thean" in args:
        groups = get_all_forward_groups(cfg)
        await private_or_current_reply(
            update,
            context,
            f"▶️ Bot RESUMED (Active for {len(groups)} groups)."
        )


@user_guard
async def cmd_status(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Show current bot state."""
    await delete_group_command(update, context)
    cfg    = load_config()
    paused = is_paused(cfg)
    mode   = get_mode(cfg)
    groups = get_all_forward_groups(cfg)
    reg    = load_registered_groups()

    lines = [
        f"{'⏸ PAUSED' if paused else '▶️ ACTIVE'}",
        f"Image mode: {mode.upper()}",
        f"Forward groups: {len(groups)} total",
    ]
    if reg:
        lines.append("Registered groups:")
        for g in reg:
            lines.append(f"  • {g.get('title', '')} ({g['chat_id']})")
    else:
        lines.append("No registered groups (configure forward_groups in config.json)")

    await private_or_current_reply(update, context, "\n".join(lines))


@user_guard
async def cmd_statues(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """/statues [code] - show status meanings, next flow, and app action mapping."""
    await delete_group_command(update, context)
    cfg = load_config()
    flows = cfg.get("ask_status_flow", {})
    actions = cfg.get("ask_app_actions", {})
    if not isinstance(flows, dict):
        flows = {}
    if not isinstance(actions, dict):
        actions = {}

    requested = [a.strip() for a in (context.args or []) if a.strip()]
    if requested:
        codes = [requested[0]]
    else:
        codes = sorted(
            {k for k in set(flows) | set(actions) if k != "default"},
            key=lambda v: int(v) if str(v).lstrip("-").isdigit() else 99999,
        )

    def one_status(code):
        flow = flows.get(code, {})
        action = actions.get(code, {})
        if not isinstance(flow, dict):
            flow = {}
        if not isinstance(action, dict):
            action = {}

        lines = [f"{code}"]
        meaning = flow.get("meaning")
        next_statuses = flow.get("next_statuses", [])
        if not isinstance(next_statuses, list):
            next_statuses = [next_statuses] if next_statuses else []
        next_statuses = [str(s) for s in next_statuses if str(s).strip()]

        if meaning:
            lines.append(f"Meaning: {meaning}")
        if next_statuses:
            lines.append(f"Next: {', '.join(next_statuses)}")
        elif flow:
            lines.append("Next: none")
        if flow.get("next_flow"):
            lines.append(f"Flow: {flow['next_flow']}")

        actor = action.get("actor")
        app = action.get("app")
        tab = action.get("tab")
        function = action.get("function")
        action_text = action.get("action")
        if actor:
            lines.append(f"Actor: {actor}")
        if app:
            lines.append(f"App: {app}")
        if tab:
            lines.append(f"Tab: {tab}")
        if function:
            lines.append(f"Function: {function}")
        if action_text:
            lines.append(f"Do: {action_text}")

        if len(lines) == 1:
            lines.append("No mapping found.")
        return "\n".join(lines)

    if not codes:
        await private_or_current_reply(update, context, "No status mappings configured.")
        return

    header = "Status details"
    if requested:
        header = f"Status details for {codes[0]}"
    chunks = []
    cur = [header]
    for code in codes:
        block = one_status(code)
        candidate = "\n\n".join(cur + [block])
        if len(candidate) > 3600 and len(cur) > 1:
            chunks.append("\n\n".join(cur))
            cur = [header + " (continued)", block]
        else:
            cur.append(block)
    if cur:
        chunks.append("\n\n".join(cur))

    for chunk in chunks:
        await private_or_current_reply(update, context, chunk)


@user_guard
async def cmd_test_mode(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await run_push(update, context, force_test=True)


@user_guard
async def cmd_clean(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Delete all report files generated today."""
    await delete_group_command(update, context)
    today_str = datetime.now().strftime("%Y-%m-%d")
    deleted_dirs = 0
    deleted_files = 0
    errors = []

    try:
        if not os.path.exists(REPORTS_LOG_PATH):
            await private_or_current_reply(update, context,
                "🗑 No reports found for today.")
            return

        with open(REPORTS_LOG_PATH, encoding="utf-8") as f:
            data = json.load(f)

        dirs_today = data.get(today_str, [])
        if not dirs_today:
            await private_or_current_reply(update, context,
                "🗑 No reports found for today.")
            return

        import shutil
        for d in dirs_today:
            if os.path.isdir(d):
                try:
                    # Count files before deleting
                    for root, _, files in os.walk(d):
                        deleted_files += len(files)
                    shutil.rmtree(d, ignore_errors=True)
                    deleted_dirs += 1
                except Exception as e:
                    errors.append(str(e))

        # Clear today's log
        data[today_str] = []
        with open(REPORTS_LOG_PATH, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)

        msg_lines = [
            f"🗑 Cleaned today's reports ({today_str})",
            f"• Removed {deleted_dirs} report folder(s)",
            f"• Deleted {deleted_files} file(s)",
        ]
        if errors:
            msg_lines.append(f"⚠️ {len(errors)} error(s): {'; '.join(errors[:3])}")
        await private_or_current_reply(update, context, "\n".join(msg_lines))

    except Exception as e:
        log.exception("Error in cmd_clean")
        await private_or_current_reply(update, context, f"❌ Clean failed: {e}")



@user_guard
async def cmd_delete_report(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Delete a report message by replying to it with /deletereport."""
    user_id = update.effective_user.id if update.effective_user else "Unknown"
    chat_id = update.effective_chat.id if update.effective_chat else "Unknown"
    log.info("=== cmd_delete_report triggered by user %s in chat %s ===", user_id, chat_id)
    
    await delete_group_command(update, context)
    
    message = update.effective_message
    if not message:
        log.warning("cmd_delete_report: No effective message found in update.")
        return
        
    if not message.reply_to_message:
        log.info("cmd_delete_report: Message is not a reply to another message.")
        await private_or_current_reply(
            update, context, 
            "⚠️ Please reply to the report message you want to delete with `/deletereport`."
        )
        return

    target_msg = message.reply_to_message
    log.info("cmd_delete_report: Replying to message ID %s sent by user %s (is_bot: %s)", 
             target_msg.message_id, 
             target_msg.from_user.id if target_msg.from_user else "Unknown",
             target_msg.from_user.is_bot if target_msg.from_user else "Unknown")
    
    try:
        bot_user = await context.bot.get_me()
        log.info("cmd_delete_report: Bot user ID is %s", bot_user.id)
        if target_msg.from_user.id != bot_user.id:
            log.info("cmd_delete_report: Message was not sent by the bot (sender ID %s != bot ID %s). Ignoring delete request.",
                     target_msg.from_user.id, bot_user.id)
            await private_or_current_reply(
                update, context, 
                "⚠️ You can only delete messages sent by the bot."
            )
            return
    except Exception as e:
        log.warning("Could not verify bot identity: %s", e)

    try:
        log.info("cmd_delete_report: Attempting to delete message %s in chat %s", target_msg.message_id, chat_id)
        await context.bot.delete_message(
            chat_id=chat_id,
            message_id=target_msg.message_id
        )
        log.info("cmd_delete_report: Message %s deleted successfully.", target_msg.message_id)
    except Exception as e:
        log.error("Failed to delete report message: %s", e)
        await private_or_current_reply(
            update, context, 
            f"❌ Failed to delete message: {e}"
        )



@user_guard
async def cmd_mode(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Toggle wide/long image mode."""
    await delete_group_command(update, context)
    cfg = load_config()
    cur = get_mode(cfg)
    new = "long" if cur == "wide" else "wide"
    cfg["telegram"]["image_mode"] = new
    save_config(cfg)
    desc = "stacked vertically (tall)" if new == "long" else "side by side (wide)"
    await private_or_current_reply(update, context, f"Image mode → {new.upper()} ({desc})")


@user_guard
async def cmd_register(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Register the current group to receive report forwards."""
    await delete_group_command(update, context)
    chat = update.effective_chat
    if chat.type not in ("group", "supergroup"):
        await private_or_current_reply(
            update,
            context,
            "Use /register inside the Telegram group you want reports sent to."
        )
        return
    groups = load_registered_groups()
    if any(g["chat_id"] == chat.id for g in groups):
        await private_or_current_reply(
            update,
            context,
            f"This group is already registered.\nChat ID: {chat.id}"
        )
        return
    groups.append({"chat_id": chat.id, "title": chat.title or str(chat.id)})
    save_registered_groups(groups)
    await private_or_current_reply(
        update,
        context,
        f"✅ Group registered for report forwards.\n"
        f"Title: {chat.title}\n"
        f"Chat ID: {chat.id}\n\n"
        f"Use /pause to pause forwarding and /resume to re-enable."
    )


@user_guard
async def cmd_unregister(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Remove the current group from forwards."""
    await delete_group_command(update, context)
    chat = update.effective_chat
    groups = load_registered_groups()
    new_groups = [g for g in groups if g["chat_id"] != chat.id]
    if len(new_groups) == len(groups):
        await private_or_current_reply(update, context, "This group is not registered.")
        return
    save_registered_groups(new_groups)
    await private_or_current_reply(update, context, f"Group removed from forwards: {chat.title}")


@user_guard
async def cmd_help(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Show all available commands only when secret token 'thean' is provided."""
    args = [a.strip().lower() for a in (context.args or []) if a.strip()]
    secret_token = args[0] if args else ""

    # Secret token protection: Stay completely silent if not 'thean'
    if secret_token != "thean":
        return

    # Authorized view when using: /help thean
    text = (
        "🔑 *Security Token Accepted: `THEAN`*\n"
        "━━━━━━━━━━━━━━━━━━━━━━\n"
        "📋 *Authorized Bot Commands*\n"
        "\n"
        "📊 *Reports & Pushes*\n"
        "`push` — Fetch data & send to 24 branch groups\n"
        "`push zone` — Fetch & send to 5 zone groups\n"
        "`push all` — Fetch & send to all groups + zones\n"
        "`/total` — Summary image + Excel (all branches)\n"
        "`/total zone` — Summary for Zone 1-5\n"
        "`/morning` — Morning report (-12h hold filter)\n"
        "`/tomorrow [zone/all]` — Shipments Tomorrow report\n"
        "`/delayed` or `/ge3` — Delayed backlog (>= 3 days / 72+ hours)\n"
        "`/totalkpi` — KPI performance report\n"
        "`/tpg` — TPG Operational Summary\n"
        "\n"
        "📥 *Export & Search*\n"
        "`/export <branch>` — Export branch post office list\n"
        "`/exportall2` or `/export all2` — Export post offices with details, address, coordinates\n"
        "`/find <phone/order>` — Search order tracking\n"
        "`/ask <order_id>` — Show status & next steps\n"
        "`/statues [code]` — Show status flow & details\n"
        "`/qr <order_id>` — Generate scannable QR code\n"
        "\n"
        "🛠 *Bill Management*\n"
        "`/done <id> [note]` — Mark delayed bill as called/done (removes from /delayed)\n"
        "`/undone <id>` — Remove bill from done list\n"
        "`/donelist` — View all bills marked as done\n"
        "`/add <id>` — Ignore test bill\n"
        "`/remove <id>` — Stop ignoring test bill\n"
        "`/list` — List ignored test bills\n"
        "`/delay <id> <days>` — Delay a bill temporarily\n"
        "`/delaylist` — List delayed bills\n"
        "\n"
        "⚙️ *Bot Controls*\n"
        "`/pause` or `/stop` — Pause auto-forwarding\n"
        "`/resume` or `/start` — Resume forwarding\n"
        "`/status` — Show bot operational status\n"
        "`/mode` — Toggle image layout (LONG / WIDE)\n"
        "`/groups` — List registered groups\n"
        "━━━━━━━━━━━━━━━━━━━━━━\n"
        "🔒 *Keep this token confidential.*"
    )
    await update.message.reply_text(text, parse_mode="Markdown")


@pm_required_handler
async def cmd_tomorrow(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Generates exact SHIPMENTS TOMORROW REPORT (Báo cáo hàng đến) matching official template."""
    await delete_group_command(update, context)
    cfg = load_config()

    args = [a.strip() for a in (context.args or []) if a.strip()]
    target_label = " ".join(args) if args else "Zone 1"


    msg = await send_requester_text(update, context, f"Generating SHIPMENTS TOMORROW REPORT ({target_label})...")
    tmpdir = tempfile.mkdtemp(prefix="tomorrow_")
    track_report_dir(tmpdir)
    stamp = datetime.now().strftime("%d.%m_%HH%M")
    src   = os.path.join(tmpdir, f"export_{stamp}.xlsx")

    try:
        downloader.download_detail(cfg["api"], src, force_refresh=True)
        import shipments_tomorrow
        out_xlsx = os.path.join(tmpdir, f"SHIPMENTS_TOMORROW_REPORT_{stamp}_{target_label.replace(' ', '_')}.xlsx")
        bills, weight = shipments_tomorrow.build_shipments_tomorrow_report(src, out_xlsx, target_label=target_label)

        # Generate clean Executive Summary cropped image matching user template
        try:
            img_buf = shipments_tomorrow.render_executive_summary_image(out_xlsx)
            img_buf.name = f"EXECUTIVE_SUMMARY_{target_label.replace(' ', '_')}.png"
            await send_requester_photo(update, context, img_buf)
        except Exception as e_img:
            log.warning("Could not render executive summary image: %s", e_img)


        caption = f"🚚 *SHIPMENTS TOMORROW REPORT ({target_label})*\n📦 Total Bills: `{bills}`\n⚖️ Total Weight: `{weight/1000:,.2f} kg`"
        await send_requester_document(update, context, out_xlsx, filename=os.path.basename(out_xlsx), caption=caption)


        # Forward to target Telegram group (both branch groups in forward_mapping and zone groups in zone_forward_mapping)
        tgt_upper = target_label.upper().replace(" ", "")
        chats_to_send = {update.effective_chat.id} if update.effective_chat else set()


        if tgt_upper in ("BRANCH", "BRANCHES", "PROVINCE", "PROVINCES"):
            # Forward ONLY to provincial branch groups in forward_mapping (excluding PNP and KAN)
            fwd_map = get_forward_mapping(cfg)
            total_sent_branches = 0
            for gid, handles in fwd_map.items():
                if not handles or "*" in handles:
                    continue
                br_code = handles[0].upper()
                # Exclude Phnom Penh (PNPP001..PNPP014) and Kandal (KANP001)
                if br_code.startswith("PNP") or br_code.startswith("KAN"):
                    continue
                br_xlsx = os.path.join(tmpdir, f"SHIPMENTS_TOMORROW_REPORT_{stamp}_{br_code}.xlsx")
                try:
                    b_bills, b_weight = shipments_tomorrow.build_shipments_tomorrow_report(src, br_xlsx, target_label=br_code)
                    if b_bills > 0:
                        b_caption = f"🚚 *SHIPMENTS TOMORROW REPORT ({br_code})*\n📦 Total Bills: `{b_bills}`\n⚖️ Total Weight: `{b_weight/1000:,.2f} kg`"
                        try:
                            b_img = shipments_tomorrow.render_executive_summary_image(br_xlsx)
                            b_img.name = f"EXECUTIVE_SUMMARY_{br_code}.png"
                            await safe_api_call(context.bot.send_photo, chat_id=int(gid), photo=b_img)
                        except Exception as e_bp:
                            log.warning("Failed sending branch photo to group %s: %s", gid, e_bp)

                        with open(br_xlsx, "rb") as f_doc:
                            await safe_api_call(
                                context.bot.send_document,
                                chat_id=int(gid),
                                document=f_doc,
                                filename=os.path.basename(br_xlsx),
                                caption=b_caption
                            )
                            total_sent_branches += 1
                except Exception as e_br:
                    log.warning("Failed building/forwarding tomorrow report for branch %s: %s", br_code, e_br)

            await edit_or_send_requester_text(msg, update, context, f"✅ Done! Forwarded SHIPMENTS TOMORROW REPORTS to {total_sent_branches} Provincial Branch Groups (excluding PNP/KAN).")
            return

        if tgt_upper in ("ALL", "MEGA"):
            # 1. Forward for all 5 Zones to their respective Zone groups
            zone_fwd_map = cfg.get("zone_forward_mapping", {})
            total_sent_zones = 0
            for z_idx in range(1, 6):
                z_name = f"Zone {z_idx}"
                z_clean = f"zone{z_idx}"
                z_xlsx = os.path.join(tmpdir, f"SHIPMENTS_TOMORROW_REPORT_{stamp}_{z_name.replace(' ', '_')}.xlsx")
                z_bills, z_weight = shipments_tomorrow.build_shipments_tomorrow_report(src, z_xlsx, target_label=z_name)
                z_caption = f"🚚 *SHIPMENTS TOMORROW REPORT ({z_name})*\n📦 Total Bills: `{z_bills}`\n⚖️ Total Weight: `{z_weight/1000:,.2f} kg`"

                for gid, zkey in zone_fwd_map.items():
                    if zkey.lower() == z_clean:
                        try:
                            try:
                                z_img = shipments_tomorrow.render_executive_summary_image(z_xlsx)
                                z_img.name = f"EXECUTIVE_SUMMARY_{z_name.replace(' ', '_')}.png"
                                await safe_api_call(context.bot.send_photo, chat_id=int(gid), photo=z_img)
                            except Exception as e_zp:
                                log.warning("Failed sending zone photo to group %s: %s", gid, e_zp)

                            with open(z_xlsx, "rb") as f_doc:
                                await safe_api_call(
                                    context.bot.send_document,
                                    chat_id=int(gid),
                                    document=f_doc,
                                    filename=os.path.basename(z_xlsx),
                                    caption=z_caption
                                )
                                total_sent_zones += 1
                        except Exception as e_fwd:
                            log.warning("Failed forwarding /tomorrow document to zone group %s: %s", gid, e_fwd)

            # 2. Forward for all registered Provincial Branch groups in forward_mapping (excluding PNP and KAN)
            fwd_map = get_forward_mapping(cfg)
            total_sent_branches = 0
            for gid, handles in fwd_map.items():
                if not handles or "*" in handles:
                    continue
                br_code = handles[0].upper()
                if br_code.startswith("PNP") or br_code.startswith("KAN"):
                    continue
                br_xlsx = os.path.join(tmpdir, f"SHIPMENTS_TOMORROW_REPORT_{stamp}_{br_code}.xlsx")
                try:
                    b_bills, b_weight = shipments_tomorrow.build_shipments_tomorrow_report(src, br_xlsx, target_label=br_code)
                    if b_bills > 0:
                        b_caption = f"🚚 *SHIPMENTS TOMORROW REPORT ({br_code})*\n📦 Total Bills: `{b_bills}`\n⚖️ Total Weight: `{b_weight/1000:,.2f} kg`"
                        try:
                            b_img = shipments_tomorrow.render_executive_summary_image(br_xlsx)
                            b_img.name = f"EXECUTIVE_SUMMARY_{br_code}.png"
                            await safe_api_call(context.bot.send_photo, chat_id=int(gid), photo=b_img)
                        except Exception as e_bp:
                            log.warning("Failed sending branch photo to group %s: %s", gid, e_bp)

                        with open(br_xlsx, "rb") as f_doc:
                            await safe_api_call(
                                context.bot.send_document,
                                chat_id=int(gid),
                                document=f_doc,
                                filename=os.path.basename(br_xlsx),
                                caption=b_caption
                            )
                            total_sent_branches += 1
                except Exception as e_br:
                    log.warning("Failed building/forwarding tomorrow report for branch %s: %s", br_code, e_br)

            await edit_or_send_requester_text(msg, update, context, f"✅ Done! Forwarded SHIPMENTS TOMORROW REPORTS to {total_sent_zones} Zone Groups and {total_sent_branches} Provincial Branch Groups (excluding PNP/KAN).")
            return

        # Single target forwarding (Zone or Branch)
        # 1. Check zone_forward_mapping (e.g. zone1 -> "-1004390650725")
        zone_fwd_map = cfg.get("zone_forward_mapping", {})
        if tgt_upper.startswith("ZONE"):
            zone_key_clean = "zone" + tgt_upper.replace("ZONE", "").strip()
            for gid, zkey in zone_fwd_map.items():
                if zkey.lower() == zone_key_clean.lower():
                    chats_to_send.add(int(gid))

        # 2. Check forward_mapping (e.g. SVAP001 -> branch chat)
        fwd_map = get_forward_mapping(cfg)
        for gid, handles in fwd_map.items():
            if any(h.upper() in (tgt_upper, tgt_upper[:3]) for h in handles):
                chats_to_send.add(int(gid))

        for cid in chats_to_send:
            try:
                try:
                    g_img = shipments_tomorrow.render_executive_summary_image(out_xlsx)
                    g_img.name = f"EXECUTIVE_SUMMARY_{target_label.replace(' ', '_')}.png"
                    await safe_api_call(context.bot.send_photo, chat_id=cid, photo=g_img)
                except Exception as e_gp:
                    log.warning("Failed sending photo to group %s: %s", cid, e_gp)

                with open(out_xlsx, "rb") as f_doc:
                    await safe_api_call(
                        context.bot.send_document,
                        chat_id=cid,
                        document=f_doc,
                        filename=os.path.basename(out_xlsx),
                        caption=caption
                    )
            except Exception as e_fwd:
                log.warning("Failed forwarding /tomorrow document to group %s: %s", cid, e_fwd)

        await edit_or_send_requester_text(msg, update, context, f"✅ Done! Sent & forwarded SHIPMENTS TOMORROW REPORT ({target_label}) with {bills} bills ({weight/1000:,.2f} kg).")








    except Exception as e:
        log.exception("Error in /tomorrow command: %s", e)
        await edit_or_send_requester_text(msg, update, context, f"❌ Error generating tomorrow report: {e}")



@pm_required_handler
async def cmd_penalty(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Generates Stagnant Inventory & Handover Penalty Report."""
    await delete_group_command(update, context)
    cfg = load_config()

    args = [a.strip() for a in (context.args or []) if a.strip()]
    target_label = " ".join(args) if args else "ALL"

    msg = await send_requester_text(update, context, f"⏳ Generating INVENTORY PENALTY REPORT ({target_label})...")
    tmpdir = tempfile.mkdtemp(prefix="penalty_")
    track_report_dir(tmpdir)
    stamp = datetime.now().strftime("%d.%m_%HH%M")
    src = os.path.join(tmpdir, f"export_{stamp}.xlsx")

    try:
        downloader.download_detail(cfg["api"], src, force_refresh=True)
        import penalty_report
        out_xlsx = os.path.join(tmpdir, f"INVENTORY_PENALTY_REPORT_{stamp}_{target_label.replace(' ', '_')}.xlsx")
        tot_ho, tot_del, tot_pen_cnt, tot_fine = penalty_report.build_penalty_report(src, out_xlsx, target_label=target_label)

        caption = (
            f"📊 *INVENTORY PENALTY REPORT ({target_label.upper()})*\n"
            f"📦 Handover Bills: `{tot_ho}`\n"
            f"🚚 Delivery Bills: `{tot_del}`\n"
            f"⚠️ Penalized Bills: `{tot_pen_cnt}`\n"
            f"💰 Total Fine: `-${tot_fine:.2f}`\n\n"
            f"_• Excused Green: 420, 472 ($0 fine)_\n"
            f"_• SLA Penalty: 1-2d (-$0.10), >3d (-$0.40)_"
        )
        await send_requester_document(update, context, out_xlsx, filename=os.path.basename(out_xlsx), caption=caption)
        await edit_or_send_requester_text(msg, update, context, f"✅ Done! Sent INVENTORY PENALTY REPORT ({target_label}).")
    except Exception as e:
        log.exception("Error in /penalty command: %s", e)
        await edit_or_send_requester_text(msg, update, context, f"❌ Error generating penalty report: {e}")


@pm_required_handler
async def cmd_speed(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Generates Fast Delivery Speed Bonus Report (Status 410)."""
    await delete_group_command(update, context)
    cfg = load_config()

    args = [a.strip() for a in (context.args or []) if a.strip()]
    target_label = " ".join(args) if args else "ALL"

    msg = await send_requester_text(update, context, f"⏳ Generating FAST DELIVERY SPEED REPORT ({target_label})...")
    tmpdir = tempfile.mkdtemp(prefix="speed_")
    track_report_dir(tmpdir)
    stamp = datetime.now().strftime("%d.%m_%HH%M")
    src = os.path.join(tmpdir, f"export_{stamp}.xlsx")

    try:
        downloader.download_detail(cfg["api"], src, force_refresh=True)
        import speed_report
        out_xlsx = os.path.join(tmpdir, f"DELIVERY_SPEED_REPORT_{stamp}_{target_label.replace(' ', '_')}.xlsx")
        tot_del, tot_u2, tot_24, tot_o8, tot_pay = speed_report.build_speed_report(src, out_xlsx, target_label=target_label)

        fast_pct = ((tot_u2 + tot_24) / tot_del * 100) if tot_del > 0 else 0
        caption = (
            f"⏱️ *FAST DELIVERY SPEED REPORT ({target_label.upper()})*\n"
            f"📦 Total Delivered: `{tot_del}`\n"
            f"🟢 < 2 Hours (+50%): `{tot_u2}`\n"
            f"🔵 2 - 4 Hours (+25%): `{tot_24}`\n"
            f"🔴 > 8 Hours (-25%): `{tot_o8}`\n"
            f"⚡ Fast Rate (<4h): `{fast_pct:.1f}%`\n"
            f"💵 Total Commission: `${tot_pay:.2f}`"
        )
        await send_requester_document(update, context, out_xlsx, filename=os.path.basename(out_xlsx), caption=caption)
        await edit_or_send_requester_text(msg, update, context, f"✅ Done! Sent FAST DELIVERY SPEED REPORT ({target_label}).")
    except Exception as e:
        log.exception("Error in /speed command: %s", e)
        await edit_or_send_requester_text(msg, update, context, f"❌ Error generating speed report: {e}")

@pm_required_handler
async def cmd_delayed(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Handler for /delayed or /ge3 or /backlog command to export NOT ASSIGN & DELIVERY >= 3 DAYS."""
    await delete_group_command(update, context)
    cfg = load_config()

    args = [a.strip() for a in (context.args or []) if a.strip()]
    min_days = 3
    if args and args[0].isdigit():
        min_days = int(args[0])

    msg = await send_requester_text(update, context, f"⏳ Fetching live TMS data for Delayed Backlog (>= {min_days} Days / 72+ Hours)...")
    tmpdir = tempfile.mkdtemp(prefix="delayed_")
    track_report_dir(tmpdir)
    stamp = datetime.now().strftime("%d.%m_%HH%M")
    src = os.path.join(tmpdir, f"export_{stamp}.xlsx")

    try:
        downloader.download_detail(cfg["api"], src, force_refresh=True)
        import delayed_report
        out_xlsx = os.path.join(tmpdir, f"NOT_ASSIGN_AND_DELIVERY_GE_{min_days}DAYS_{stamp}.xlsx")
        desktop_xlsx = rf"c:\Users\DELL\Desktop\NOT_ASSIGN_AND_DELIVERY_GE_{min_days}DAYS.xlsx"
        
        tot_bills, na_bills, del_bills, weight_kg, cod_usd = delayed_report.build_delayed_ge3_report(
            src, out_xlsx, min_days=min_days
        )
        
        try:
            import shutil
            shutil.copyfile(out_xlsx, desktop_xlsx)
        except Exception:
            pass

        with open(out_xlsx, "rb") as f_doc:
            await send_requester_document(
                update,
                context,
                f_doc,
                filename=f"NOT_ASSIGN_AND_DELIVERY_GE_{min_days}DAYS.xlsx",
                caption=None
            )
        
        # Remove temporary status message so only the file appears
        if msg:
            try:
                await msg.delete()
            except Exception:
                pass
    except Exception as e:
        log.exception("Error in /delayed command: %s", e)
        await edit_or_send_requester_text(msg, update, context, f"❌ Error generating delayed report: {e}")


async def cmd_lag(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Handler for /lag, /mismatch, /audit to export bills where Web Tracking is S410/S520 but Database is lagging."""
    await delete_group_command(update, context)
    cfg = load_config()

    msg = await send_requester_text(update, context, "⏳ Scanning live Web Tracking database for lagging API bills (S410 Delivered on Web vs Pending in DB)...")
    tmpdir = tempfile.mkdtemp(prefix="lag_")
    track_report_dir(tmpdir)
    stamp = datetime.now().strftime("%d.%m_%HH%M")
    src = os.path.join(tmpdir, f"export_{stamp}.xlsx")

    try:
        downloader.download_detail(cfg["api"], src, force_refresh=True)
        import lag_report
        out_xlsx = os.path.join(tmpdir, f"API_LAG_REPORT_{stamp}.xlsx")
        desktop_xlsx = rf"c:\Users\DELL\Desktop\API_LAG_REPORT_{stamp}.xlsx"

        tot_lag, deliv, ret = lag_report.build_lag_report(src, out_xlsx, cfg["api"])

        try:
            import shutil
            shutil.copyfile(out_xlsx, desktop_xlsx)
        except Exception:
            pass

        caption = (
            f"📊 API Lag Audit Report — {datetime.now().strftime('%d/%m/%Y %H:%M')}\n"
            f"• Total Lagging Bills: {tot_lag:,}\n"
            f"• Delivered on Web (S410): {deliv:,}\n"
            f"• Returned on Web (S520): {ret:,}\n\n"
            f"💡 These bills are already completed on Web/Mobile App, but the database API was lagging behind."
        )

        with open(out_xlsx, "rb") as f_doc:
            await send_requester_document(
                update,
                context,
                f_doc,
                filename=f"API_LAG_REPORT_{stamp}.xlsx",
                caption=caption
            )

        if msg:
            try:
                await msg.delete()
            except Exception:
                pass
    except Exception as e:
        log.exception("Error in /lag command: %s", e)
        await edit_or_send_requester_text(msg, update, context, f"❌ Error generating lag report: {e}")


def _extract_bill_ids_and_note(raw_text: str) -> tuple[list[str], str]:
    """Extracts bill IDs and optional note from flexible formats: commas, parentheses, spaces."""
    import re
    text = re.sub(r"^/\w+\s*", "", raw_text.strip())
    text = text.replace("(", " ").replace(")", " ").replace("[", " ").replace("]", " ")
    
    if "," in text:
        parts = [p.strip() for p in text.split(",") if p.strip()]
        bill_ids = []
        note = ""
        for p in parts:
            words = p.split()
            if words:
                bid = re.sub(r"[^\w]", "", words[0]).strip()
                if bid:
                    bill_ids.append(bid)
                if len(words) > 1 and not note:
                    note = " ".join(words[1:])
        return bill_ids, note
    
    tokens = text.split()
    if not tokens:
        return [], ""
    
    if len(tokens) > 1 and re.match(r"^\d{8,12}$", tokens[0]) and not re.match(r"^\d{8,12}$", tokens[1]):
        return [tokens[0]], " ".join(tokens[1:])
    
    bill_ids = []
    for t in tokens:
        bid = re.sub(r"[^\w]", "", t).strip()
        if bid:
            bill_ids.append(bid)
    return bill_ids, ""


@pm_required_handler
async def cmd_done(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Mark delayed bill(s) as DONE so they no longer appear in /delayed (>= 3 days) reports."""
    await delete_group_command(update, context)
    raw_text = update.message.text if (update.message and update.message.text) else ""
    bill_ids, remark = _extract_bill_ids_and_note(raw_text)

    if not bill_ids and context.args:
        bill_ids = [re.sub(r"[^\w]", "", a).strip() for a in context.args if re.sub(r"[^\w]", "", a).strip()]

    if not bill_ids:
        await private_or_current_reply(
            update, context,
            "ℹ️ *Usage:*\n"
            "• `/done 3204298321`\n"
            "• `/done 3204298321, 3204280987, 3204306743`\n"
            "• `/done (3204298321, 3204280987)`\n"
            "• `/done 3204298321 called customer confirmed`",
            parse_mode="Markdown"
        )
        return

    import delayed_report
    user = update.effective_user
    username = user.full_name if user else "Unknown"

    for bid in bill_ids:
        delayed_report.add_done_bill(bid, user_name=username, remark=remark)

    tot_done = len(delayed_report.get_done_bill_ids())
    note_str = f"\n📝 *Note:* {remark}" if remark else ""
    await private_or_current_reply(
        update, context,
        f"✅ *Marked as DONE:* `{', '.join(bill_ids)}`{note_str}\n\n"
        f"📊 *Total Done Bills:* `{tot_done}`\n"
        f"*(These bills are now hidden from `/delayed` reports. Use `/undelayed <id>` to unhide)*",
        parse_mode="Markdown"
    )


@pm_required_handler
async def cmd_undelayed(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Remove bill(s) from DONE/hidden list so they appear back in /delayed reports."""
    await delete_group_command(update, context)
    raw_text = update.message.text if (update.message and update.message.text) else ""
    bill_ids, _ = _extract_bill_ids_and_note(raw_text)

    if not bill_ids and context.args:
        bill_ids = [re.sub(r"[^\w]", "", a).strip() for a in context.args if re.sub(r"[^\w]", "", a).strip()]

    if not bill_ids:
        await private_or_current_reply(
            update, context,
            "ℹ️ *Usage:*\n• `/undelayed 3204298321`\n• `/undelayed 3204298321, 3204280987`\n• `/undelayed (3204298321, 3204280987)`",
            parse_mode="Markdown"
        )
        return

    import delayed_report
    removed = []
    not_found = []
    for bid in bill_ids:
        if delayed_report.remove_done_bill(bid):
            removed.append(bid)
        else:
            not_found.append(bid)

    msg_lines = []
    if removed:
        msg_lines.append(f"↩️ *Unhidden from DONE list:* `{', '.join(removed)}`\n*(They will appear back in `/delayed` reports)*")
    if not_found:
        msg_lines.append(f"⚠️ *Not found in DONE list:* `{', '.join(not_found)}`")
    
    tot_done = len(delayed_report.get_done_bill_ids())
    msg_lines.append(f"\n📊 *Remaining Done Bills:* `{tot_done}`")
    await private_or_current_reply(update, context, "\n".join(msg_lines), parse_mode="Markdown")


@pm_required_handler
async def cmd_donelist(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """List all bills marked as DONE."""
    await delete_group_command(update, context)
    import delayed_report
    done_dict = delayed_report.load_done_bills()

    if not done_dict:
        await private_or_current_reply(update, context, "📋 *DONE List is currently empty.*", parse_mode="Markdown")
        return

    lines = [f"📋 *BILLS MARKED AS DONE ({len(done_dict)} Total)*", "━━━━━━━━━━━━━━━━━━━━━━"]
    for idx, (oid, info) in enumerate(sorted(done_dict.items(), key=lambda x: x[1].get('marked_at', ''), reverse=True)[:50], 1):
        dt = info.get('marked_at', '')
        user = info.get('marked_by', '')
        rmk = info.get('remark', '')
        rmk_str = f" — *Note:* {rmk}" if rmk else ""
        lines.append(f"{idx}. `{oid}` ({dt} by {user}){rmk_str}")

    if len(done_dict) > 50:
        lines.append(f"\n_...and {len(done_dict)-50} more._")

    await private_or_current_reply(update, context, "\n".join(lines), parse_mode="Markdown")


@pm_required_handler
async def cmd_today(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Handler for /today [branch/zone/post_office] command."""
    await delete_group_command(update, context)
    cfg = load_config()

    args = [a.strip() for a in (context.args or []) if a.strip()]
    target_label = " ".join(args) if args else "Zone 1"

    msg = await send_requester_text(update, context, f"Generating TODAY PERFORMANCE REPORT ({target_label})...")
    tmpdir = tempfile.mkdtemp(prefix="today_")
    track_report_dir(tmpdir)
    stamp = datetime.now().strftime("%d.%m_%HH%M")
    src   = os.path.join(tmpdir, f"export_{stamp}.xlsx")

    try:
        downloader.download_detail(cfg["api"], src, force_refresh=True)
        import branch_today
        out_xlsx = os.path.join(tmpdir, f"BRANCH_TODAY_PERFORMANCE_{stamp}_{target_label.replace(' ', '_')}.xlsx")
        from_mega, pending, success = branch_today.build_branch_today_report(src, out_xlsx, target_label=target_label)


        # Generate summary image for instant Telegram mobile viewing
        try:
            img_buf = branch_today.render_today_summary_image(out_xlsx)
            img_buf.name = f"TODAY_PERFORMANCE_{target_label.replace(' ', '_')}.png"
            await send_requester_photo(update, context, img_buf)
        except Exception as e_img:
            log.warning("Could not render today summary image: %s", e_img)

        caption = os.path.basename(out_xlsx)
        await send_requester_document(update, context, out_xlsx, filename=os.path.basename(out_xlsx), caption=None)

        # Forward to target Telegram group (both branch groups in forward_mapping and zone groups in zone_forward_mapping)
        tgt_upper = target_label.upper().replace(" ", "")
        chats_to_send = {update.effective_chat.id} if update.effective_chat else set()

        if tgt_upper in ("ALL", "MEGA"):
            # Forward for all 5 Zones to their respective Zone groups
            zone_fwd_map = cfg.get("zone_forward_mapping", {})
            total_sent_zones = 0
            for z_idx in range(1, 6):
                z_name = f"Zone {z_idx}"
                z_clean = f"zone{z_idx}"
                z_xlsx = os.path.join(tmpdir, f"BRANCH_TODAY_PERFORMANCE_{stamp}_{z_name.replace(' ', '_')}.xlsx")
                z_fm, z_pend, z_succ = branch_today.build_branch_today_report(src, z_xlsx, target_label=z_name)
                z_caption = None

                for gid, zkey in zone_fwd_map.items():
                    if zkey.lower() == z_clean:
                        try:
                            try:
                                z_img = branch_today.render_today_summary_image(z_xlsx)
                                z_img.name = f"TODAY_PERFORMANCE_{z_name.replace(' ', '_')}.png"
                                await safe_api_call(context.bot.send_photo, chat_id=int(gid), photo=z_img)
                            except Exception as e_zp:
                                log.warning("Failed sending today photo to zone group %s: %s", gid, e_zp)

                            with open(z_xlsx, "rb") as f_doc:
                                await safe_api_call(
                                    context.bot.send_document,
                                    chat_id=int(gid),
                                    document=f_doc,
                                    filename=os.path.basename(z_xlsx),
                                    caption=z_caption
                                )
                                total_sent_zones += 1
                        except Exception as e_fwd:
                            log.warning("Failed forwarding today document to zone group %s: %s", gid, e_fwd)

            # Forward for all registered Branch groups in forward_mapping
            fwd_map = get_forward_mapping(cfg)
            total_sent_branches = 0
            for gid, handles in fwd_map.items():
                if not handles or "*" in handles:
                    continue
                br_code = handles[0].upper()
                br_xlsx = os.path.join(tmpdir, f"BRANCH_TODAY_PERFORMANCE_{stamp}_{br_code}.xlsx")
                try:
                    b_fm, b_pend, b_succ = branch_today.build_branch_today_report(src, br_xlsx, target_label=br_code)
                    if b_fm > 0 or b_pend > 0 or b_succ > 0:
                        b_caption = None
                        try:
                            b_img = branch_today.render_today_summary_image(br_xlsx)
                            b_img.name = f"TODAY_PERFORMANCE_{br_code}.png"
                            await safe_api_call(context.bot.send_photo, chat_id=int(gid), photo=b_img)
                        except Exception as e_bp:
                            log.warning("Failed sending today photo to branch group %s: %s", gid, e_bp)

                        with open(br_xlsx, "rb") as f_doc:
                            await safe_api_call(
                                context.bot.send_document,
                                chat_id=int(gid),
                                document=f_doc,
                                filename=os.path.basename(br_xlsx),
                                caption=b_caption
                            )
                            total_sent_branches += 1
                except Exception as e_br:
                    log.warning("Failed building/forwarding today report for branch %s: %s", br_code, e_br)

            await edit_or_send_requester_text(msg, update, context, f"✅ Done! Forwarded TODAY PERFORMANCE REPORTS to {total_sent_zones} Zone Groups and {total_sent_branches} Branch Groups.")
            return

        # Single target forwarding (Zone or Branch)

        zone_fwd_map = cfg.get("zone_forward_mapping", {})
        if tgt_upper.startswith("ZONE"):
            zone_key_clean = "zone" + tgt_upper.replace("ZONE", "").strip()
            for gid, zkey in zone_fwd_map.items():
                if zkey.lower() == zone_key_clean.lower():
                    chats_to_send.add(int(gid))

        fwd_map = get_forward_mapping(cfg)
        for gid, handles in fwd_map.items():
            if any(h.upper() in (tgt_upper, tgt_upper[:3]) for h in handles):
                chats_to_send.add(int(gid))

        for cid in chats_to_send:
            try:
                try:
                    g_img = branch_today.render_today_summary_image(out_xlsx)
                    g_img.name = f"TODAY_PERFORMANCE_{target_label.replace(' ', '_')}.png"
                    await safe_api_call(context.bot.send_photo, chat_id=cid, photo=g_img)
                except Exception as e_gp:
                    log.warning("Failed sending today photo to group %s: %s", cid, e_gp)

                with open(out_xlsx, "rb") as f_doc:
                    await safe_api_call(
                        context.bot.send_document,
                        chat_id=cid,
                        document=f_doc,
                        filename=os.path.basename(out_xlsx),
                        caption=caption
                    )
            except Exception as e_fwd:
                log.warning("Failed forwarding today document to group %s: %s", cid, e_fwd)

        await edit_or_send_requester_text(msg, update, context, f"✅ Done! Sent & forwarded TODAY PERFORMANCE REPORT ({target_label}) with {from_mega:,.0f} From Mega, {pending:,.0f} Pending & {success:,.0f} Success.")


    except Exception as e:
        log.exception("Error in /today command: %s", e)
        await edit_or_send_requester_text(msg, update, context, f"❌ Error generating today branch report: {e}")


@pm_required_handler
async def cmd_allthetime(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Handler for /allthetime [branch/zone] — sends text summary (From Mega / Pending / Success) to all registered groups."""
    await delete_group_command(update, context)
    cfg = load_config()

    args = [a.strip() for a in (context.args or []) if a.strip()]
    target_label = " ".join(args) if args else "all"

    msg = await send_requester_text(update, context, f"Sending TODAY TEXT SUMMARY ({target_label}) to all groups...")
    tmpdir = tempfile.mkdtemp(prefix="allthetime_")
    track_report_dir(tmpdir)
    stamp = datetime.now().strftime("%d.%m_%HH%M")
    src   = os.path.join(tmpdir, f"export_{stamp}.xlsx")

    try:
        downloader.download_detail(cfg["api"], src, force_refresh=True)
        import branch_today

        tgt_upper = target_label.upper().replace(" ", "")
        total_sent = 0

        # Determine list of targets to broadcast
        if tgt_upper in ("ALL", "MEGA", ""):
            # Send to all 5 Zones
            zone_fwd_map = cfg.get("zone_forward_mapping", {})
            for z_idx in range(1, 6):
                z_name = f"Zone {z_idx}"
                z_clean = f"zone{z_idx}"
                z_xlsx = os.path.join(tmpdir, f"ATT_{stamp}_{z_name.replace(' ', '_')}.xlsx")
                z_fm, z_pend, z_succ = branch_today.build_branch_today_report(src, z_xlsx, target_label=z_name)
                z_text = (
                    f"📊 *TODAY PERFORMANCE SUMMARY ({z_name})*\n"
                    f"📥 *From Mega*: `{z_fm:,.0f} bills`\n"
                    f"⏳ *Pending*: `{z_pend:,.0f} bills`\n"
                    f"✅ *Success*: `{z_succ:,.0f} bills`"
                )
                for gid, zkey in zone_fwd_map.items():
                    if zkey.lower() == z_clean:
                        try:
                            await safe_api_call(context.bot.send_message, chat_id=int(gid), text=z_text, parse_mode="Markdown")
                            total_sent += 1
                        except Exception as ez:
                            log.warning("Failed sending allthetime to zone %s: %s", gid, ez)

            # Send to all Branch groups
            fwd_map = get_forward_mapping(cfg)
            for gid, handles in fwd_map.items():
                if not handles or "*" in handles:
                    continue
                br_code = handles[0].upper()
                br_xlsx = os.path.join(tmpdir, f"ATT_{stamp}_{br_code}.xlsx")
                try:
                    b_fm, b_pend, b_succ = branch_today.build_branch_today_report(src, br_xlsx, target_label=br_code)
                    if b_fm > 0 or b_pend > 0 or b_succ > 0:
                        b_text = (
                            f"📊 *TODAY PERFORMANCE SUMMARY ({br_code})*\n"
                            f"📥 *From Mega*: `{b_fm:,.0f} bills`\n"
                            f"⏳ *Pending*: `{b_pend:,.0f} bills`\n"
                            f"✅ *Success*: `{b_succ:,.0f} bills`"
                        )
                        await safe_api_call(context.bot.send_message, chat_id=int(gid), text=b_text, parse_mode="Markdown")
                        total_sent += 1
                except Exception as eb:
                    log.warning("Failed sending allthetime to branch %s: %s", br_code, eb)

            await edit_or_send_requester_text(msg, update, context, f"✅ Done! Sent TODAY TEXT SUMMARY to {total_sent} groups.")

        else:
            # Single target
            single_xlsx = os.path.join(tmpdir, f"ATT_{stamp}_{tgt_upper}.xlsx")
            fm, pend, succ = branch_today.build_branch_today_report(src, single_xlsx, target_label=target_label)
            text = (
                f"📊 *TODAY PERFORMANCE SUMMARY ({target_label})*\n"
                f"📥 *From Mega*: `{fm:,.0f} bills`\n"
                f"⏳ *Pending*: `{pend:,.0f} bills`\n"
                f"✅ *Success*: `{succ:,.0f} bills`"
            )
            # Send to requester chat + matching group chats
            chats_to_send = {update.effective_chat.id} if update.effective_chat else set()
            zone_fwd_map = cfg.get("zone_forward_mapping", {})
            if tgt_upper.startswith("ZONE"):
                z_clean = "zone" + tgt_upper.replace("ZONE", "").strip()
                for gid, zkey in zone_fwd_map.items():
                    if zkey.lower() == z_clean.lower():
                        chats_to_send.add(int(gid))
            fwd_map = get_forward_mapping(cfg)
            for gid, handles in fwd_map.items():
                if any(h.upper() in (tgt_upper, tgt_upper[:3]) for h in handles):
                    chats_to_send.add(int(gid))

            for cid in chats_to_send:
                try:
                    await safe_api_call(context.bot.send_message, chat_id=cid, text=text, parse_mode="Markdown")
                    total_sent += 1
                except Exception as ec:
                    log.warning("Failed sending allthetime to chat %s: %s", cid, ec)

            await edit_or_send_requester_text(msg, update, context, f"✅ Done! Sent TODAY TEXT SUMMARY ({target_label}) to {total_sent} chats.")

    except Exception as e:
        log.exception("Error in /allthetime command: %s", e)
        await edit_or_send_requester_text(msg, update, context, f"❌ Error: {e}")


@pm_required_handler

async def cmd_total(update: Update, context: ContextTypes.DEFAULT_TYPE):

    """/total [zone] — summary image + Excel sorted by report type.
    Examples: /total  (all data)  |  /total zone5  |  /total zone1 | /total mega
    """
    await delete_group_command(update, context)
    cfg = load_config()

    # Parse optional zone and force arguments
    args = [a.strip().lower() for a in (context.args or []) if a.strip()]
    force_refresh = "force" in args
    args = [a for a in args if a != "force"]

    zone_filter = None
    zone_label = "ALL"
    if args:
        zone_key = args[0]
        if zone_key == "mega":
            msg = await send_requester_text(update, context, "Fetching data for TỒN MEGA CHECK...")
            tmpdir = tempfile.mkdtemp(prefix="mega_")
            track_report_dir(tmpdir)
            stamp  = datetime.now().strftime("%d.%m_%HH%M")
            src    = os.path.join(tmpdir, f"export_{stamp}.xlsx")
            try:
                downloader.download_detail(cfg["api"], src, branch_code="MEGA", force_refresh=True)
                msg = await edit_or_send_requester_text(msg, update, context, "Building TỒN MEGA CHECK report...")

                import pivot

                # Generate 2 separate pivot tables/images: MEGA1 and DVCMEGA1
                rows = pivot.read_source(src)
                tree, day_keys, extra_data = pivot.build_mega_pivot(rows, cfg.get("pivot", {}), cfg.get("zone_mapping", {}))

                for hub in ["MEGA1", "DVCMEGA1"]:
                    if hub in tree and tree[hub]:
                        sub_tree = {hub: tree[hub]}
                        hub_xlsx = os.path.join(tmpdir, f"Report_{hub}_{stamp}.xlsx")
                        pivot.export_mega_pivot(sub_tree, day_keys, hub_xlsx, extra_data=extra_data)
                        img_buf = excel_to_image.excel_to_image(hub_xlsx)
                        img_buf.name = f"{hub}_check.png"
                        await send_requester_photo(update, context, img_buf)

                # Detail Excel with actual order data split into Urgent / Normal
                try:
                    import mega_detail
                    detail_xlsx = os.path.join(tmpdir, f"MEGA_Detail_{stamp}.xlsx")
                    result_detail = mega_detail.build_mega_detail(src, detail_xlsx, cfg)
                    total_orders  = result_detail[0] if result_detail else 0
                    urgent_orders = result_detail[1] if result_detail else 0
                    with open(detail_xlsx, "rb") as f:
                        await send_requester_document(
                            update, context, f,
                            os.path.basename(detail_xlsx),
                            caption=f"📋 ទិន្នន័យលម្អិត MEGA {datetime.now().strftime('%d/%m/%Y %H:%M')}\nTotal: {total_orders} | Urgent: {urgent_orders}",
                        )
                except Exception as e:
                    log.warning("Failed to build mega detail Excel: %s", e)

                await edit_or_send_requester_text(msg, update, context, f"Done. TỒN MEGA CHECK {datetime.now().strftime('%d.%m.%Y %H:%M')}")

            except Exception as e:
                log.exception("Error in /total mega")
                await edit_or_send_requester_text(msg, update, context, f"Error: {e}")
            return

        total_zones = cfg.get("total_zones", {})
        if zone_key in ["zone", "zones", "allzone", "allzones", "all_zones", "zone1-5", "zone1-zone5"]:
            zone_filter = []
            for z_handles in total_zones.values():
                zone_filter.extend([h.upper() for h in z_handles])
            zone_label = "TOTAL ZONE 1-5"
        elif zone_key in total_zones:
            zone_filter = [h.upper() for h in total_zones[zone_key]]
            zone_label = zone_key.upper()
        else:
            available = ", ".join(sorted(total_zones.keys())) if total_zones else "none configured"
            await private_or_current_reply(
                update,
                context,
                f"Unknown zone '{zone_key}'.\n"
                f"Available zones: {available}, zone (all zones 1-5)\n"
                f"Usage: /total zone  or  /total zone1"
            )
            return

    msg = await send_requester_text(update, context, f"Fetching data for {zone_label} summary...")

    tmpdir = tempfile.mkdtemp(prefix="total_")
    track_report_dir(tmpdir)
    stamp  = datetime.now().strftime("%d.%m_%HH%M")
    src    = os.path.join(tmpdir, f"export_{stamp}.xlsx")

    try:
        downloader.download_detail(cfg["api"], src, force_refresh=force_refresh)
        msg = await edit_or_send_requester_text(msg, update, context, "Building summary...")

        mode = get_mode(cfg)
        result = generate_report.generate_reports_from_data(
            src, REF_PATH, tmpdir, return_metadata=True, mode=mode,
        )
        update_webapp_cache(result)
        update_dashboard_cache(result)
        save_highlight_history(result)

        # ── Zone filtering ────────────────────────────────────────────────
        if zone_filter:
            # Filter handle_results to only matching handles
            result["handle_results"] = [
                hr for hr in result["handle_results"]
                if hr["handle"] in zone_filter
            ]
            # Recalculate overall_counts from filtered handles
            overall = {"Pickup": 0, "Delivery": 0, "Transit": 0, "Branch": 0}
            for hr in result["handle_results"]:
                for k in overall:
                    overall[k] += hr["handle_counts"].get(k, 0)
            result["overall_counts"] = overall

            # Filter type_data DataFrames
            import pandas as pd
            for rn in ["Pickup", "Delivery", "Transit", "Branch"]:
                df = result.get("type_data", {}).get(rn)
                if df is not None and not df.empty:
                    filter_col = "CURRENT POST OFFICE" if rn == "Transit" else "POST OFFICE HANDLE"
                    if filter_col in df.columns:
                        result["type_data"][rn] = df[df[filter_col].isin(zone_filter)].copy()

            # Re-fetch overall if zone_filter was processed
            overall = result["overall_counts"]

        # ── Exclude showroom (A) and agent (S) from type_data for image + Excel ──
        # Their counts are already included in the main post office figures.
        import pandas as pd
        valid_handles = set(hr["handle"] for hr in result.get("handle_results", []))
        for rn in ["Pickup", "Delivery", "Transit", "Branch"]:
            df_t = result.get("type_data", {}).get(rn)
            if df_t is not None and not df_t.empty:
                po_col = "CURRENT POST OFFICE" if rn == "Transit" else "POST OFFICE HANDLE"
                if po_col in df_t.columns:
                    mask = df_t[po_col].apply(
                        lambda h: str(h).strip().upper() in valid_handles if valid_handles else not (len(str(h).strip()) >= 4 and str(h).strip()[3] in ('A', 'S'))
                    )
                    result["type_data"][rn] = df_t[mask].copy()

        # Use pre-calculated exact per-handle metrics from generate_report
        total_day_date_counts = result.get("day_date_counts", {})
        total_urgent_counts   = result.get("urgent_counts", {})
        total_vip_counts      = result.get("vip_counts", {})
        total_fee_counts      = result.get("fee_counts", {})
        total_cod_counts      = result.get("cod_counts", {})

        if zone_filter:
            zf_set = set(zone_filter)
            total_day_date_counts = {h: total_day_date_counts.get(h, {}) for h in zf_set if h in total_day_date_counts}
            total_urgent_counts   = {h: total_urgent_counts.get(h, {}) for h in zf_set if h in total_urgent_counts}
            total_vip_counts      = {h: total_vip_counts.get(h, 0) for h in zf_set if h in total_vip_counts}
            total_fee_counts      = {h: total_fee_counts.get(h, 0.0) for h in zf_set if h in total_fee_counts}
            total_cod_counts      = {h: total_cod_counts.get(h, 0.0) for h in zf_set if h in total_cod_counts}

        urgent_by_type = {"Pickup": 0, "Delivery": 0, "Transit": 0, "Branch": 0}
        today_date = datetime.now().date()
        valid_handles = set(hr["handle"] for hr in result.get("handle_results", []))

        for rn in ["Pickup", "Delivery", "Transit", "Branch"]:
            df_z = result.get("type_data", {}).get(rn)
            if df_z is None or df_z.empty:
                continue
            handle_col = "CURRENT POST OFFICE" if rn == "Transit" else "POST OFFICE HANDLE"
            if handle_col not in df_z.columns:
                continue
            for _, row_z in df_z.iterrows():
                h = str(row_z.get(handle_col, "")).strip().upper()
                if not h or (valid_handles and h not in valid_handles):
                    continue
                c_val = row_z.get("CREATED DATE") or row_z.get("CURRENT TIME")
                if pd.notna(c_val) and str(c_val).strip() and str(c_val).strip().lower() != 'nan':
                    cd = pd.to_datetime(c_val, dayfirst=True, format="mixed", errors="coerce")
                    if not pd.isna(cd):
                        if (today_date - cd.date()).days > 1:
                            urgent_by_type[rn] += 1

        overall = result["overall_counts"]
        # Ensure urgent count cannot exceed category total
        for rn in ["Pickup", "Delivery", "Transit", "Branch"]:
            urgent_by_type[rn] = min(urgent_by_type[rn], overall.get(rn, 0))

        grand_total = sum(overall.values())
        total_urgent_sum = sum(urgent_by_type.values())

        # Build final formatted caption
        result["summary_caption"] = "\n".join([
            f"📋 {zone_label} Report  {datetime.now().strftime('%d/%m/%Y %H:%M')}",
            f"Delivery: {overall.get('Delivery',0)} (U:{urgent_by_type.get('Delivery',0)})  |  "
            f"Assign Deliver: {overall.get('Branch',0)} (U:{urgent_by_type.get('Branch',0)})  |  "
            f"Pickup: {overall.get('Pickup',0)} (U:{urgent_by_type.get('Pickup',0)})  |  "
            f"Handover to Mega: {overall.get('Transit',0)} (U:{urgent_by_type.get('Transit',0)})",
            f"Grand Total: {grand_total}  |  Total Urgent: {total_urgent_sum}",
        ])

        # 1. Summary image — totals per handle
        img_buf = generate_summary.build_summary_image(
            result["handle_results"],
            result["overall_counts"],
            zone_label=zone_label,
            day_date_counts=total_day_date_counts if total_day_date_counts else None,
            urgent_counts=total_urgent_counts if total_urgent_counts else None,
            fee_counts=total_fee_counts if any(total_fee_counts.values()) else None,
            cod_counts=total_cod_counts if any(total_cod_counts.values()) else None,
            vip_counts=total_vip_counts if any(total_vip_counts.values()) else None,
        )
        img_buf.name = "summary.png"
        await send_requester_photo(update, context, img_buf)

        # 2. Total Excel — 4 tables on one sheet (Pickup / Delivery / Transit / Branch)
        label = f"Total_{zone_label}_" if zone_filter else "Total_"
        total_xlsx = os.path.join(tmpdir, f"{label}{stamp}.xlsx")
        generate_summary.build_total_excel(result, total_xlsx)

        with open(total_xlsx, "rb") as f:
            await send_requester_document(
                update,
                context,
                f,
                os.path.basename(total_xlsx),
                caption=result["summary_caption"],
            )

        await edit_or_send_requester_text(msg, update, context, f"Done. {zone_label} {datetime.now().strftime('%d.%m.%Y %H:%M')}")

    except Exception as e:
        log.exception("Error in /total")
        await edit_or_send_requester_text(msg, update, context, f"Error: {e}")

def get_highlighted_order_ids(df_t, today_date):
    """Return a set of order IDs that are highlighted in this DataFrame."""
    highlighted = set()
    if df_t.empty:
        return highlighted

    for _, row in df_t.iterrows():
        oid = str(row.get("ORDER ID") or "").strip()
        if not oid or oid.lower() == "nan":
            continue
        
        sc = str(row.get("STATUS_CODE") or "").strip()
        cd_val = row.get("CREATED DATE")
        
        is_highlight = False
        if sc in ('420', '472'):
            if cd_val:
                cd = pd.to_datetime(cd_val, dayfirst=True, format="mixed", errors="coerce")
                if not pd.isna(cd):
                    if (today_date - cd.date()).days > 7:
                        is_highlight = True
        else:
            if cd_val:
                cd = pd.to_datetime(cd_val, dayfirst=True, format="mixed", errors="coerce")
                if not pd.isna(cd):
                    if (today_date - cd.date()).days > 1:
                        is_highlight = True
        
        if is_highlight:
            highlighted.add(oid)
            
    return highlighted


@pm_required_handler
async def cmd_morning(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """/morning [zone] — Morning report with 12h age adjustment (excludes overnight hold time).
    Examples: /morning  (all data)  |  /morning zone5  |  /morning zone1
    """
    await delete_group_command(update, context)
    cfg = load_config()

    # Parse optional zone argument
    args = [a.strip().lower() for a in (context.args or []) if a.strip()]
    zone_filter = None
    zone_label = "ALL"
    
    if args:
        zone_key = args[0]
        total_zones = cfg.get("total_zones", {})
        if zone_key in total_zones:
            zone_filter = [h.upper() for h in total_zones[zone_key]]
            zone_label = zone_key.upper()
        else:
            available = ", ".join(sorted(total_zones.keys())) if total_zones else "none configured"
            await private_or_current_reply(
                update,
                context,
                f"Unknown zone '{zone_key}'.\n"
                f"Available zones: {available}\n"
                f"Usage: /morning zone5"
            )
            return

    msg = await send_requester_text(update, context, f"☀️ Fetching data for {zone_label} MORNING report (age -12h)...")

    tmpdir = tempfile.mkdtemp(prefix="morning_")
    track_report_dir(tmpdir)
    stamp = datetime.now().strftime("%d.%m_%HH%M")
    src = os.path.join(tmpdir, f"export_{stamp}.xlsx")

    try:
        downloader.download_detail(cfg["api"], src, force_refresh=True)
        msg = await edit_or_send_requester_text(msg, update, context, "Building morning summary (age adjusted -12h)...")

        mode = get_mode(cfg)
        result = generate_report.generate_reports_from_data(
            src, REF_PATH, tmpdir, return_metadata=True, mode=mode,
        )

        # ── Zone filtering ────────────────────────────────────────────────
        if zone_filter:
            result["handle_results"] = [
                hr for hr in result["handle_results"]
                if hr["handle"] in zone_filter
            ]
            overall = {"Pickup": 0, "Delivery": 0, "Transit": 0, "Branch": 0}
            for hr in result["handle_results"]:
                for k in overall:
                    overall[k] += hr["handle_counts"].get(k, 0)
            result["overall_counts"] = overall

            import pandas as pd
            for rn in ["Pickup", "Delivery", "Transit", "Branch"]:
                df = result.get("type_data", {}).get(rn)
                if df is not None and not df.empty:
                    filter_col = "CURRENT POST OFFICE" if rn == "Transit" else "POST OFFICE HANDLE"
                    if filter_col in df.columns:
                        result["type_data"][rn] = df[df[filter_col].isin(zone_filter)].copy()

        # Calculate counts
        total_day_date_counts = {}
        total_urgent_counts = {}
        urgent_by_type = {"Pickup": 0, "Delivery": 0, "Transit": 0, "Branch": 0}
        today_date = datetime.now().date()
        import pandas as pd
        valid_handles = set(hr["handle"] for hr in result.get("handle_results", []))

        for rn in ["Pickup", "Delivery", "Transit", "Branch"]:
            df_z = result.get("type_data", {}).get(rn)
            if df_z is None or df_z.empty:
                continue
            date_col_z = result.get("date_col") or (
                "CREATED DATE" if "CREATED DATE" in df_z.columns else
                "CURRENT TIME" if "CURRENT TIME" in df_z.columns else None
            )
            if date_col_z and date_col_z in df_z.columns:
                parsed_z = pd.to_datetime(df_z[date_col_z], dayfirst=True,
                                          format="mixed", errors="coerce")
                df_z = df_z.copy()
                df_z["_zdate"] = parsed_z.dt.date

            handle_col = "CURRENT POST OFFICE" if rn == "Transit" else "POST OFFICE HANDLE"
            if handle_col not in df_z.columns:
                continue

            for _, row_z in df_z.iterrows():
                h = str(row_z.get(handle_col, "")).strip().upper()
                if not h or (valid_handles and h not in valid_handles):
                    continue
                d_val = row_z.get("_zdate") if "_zdate" in df_z.columns else None
                if d_val and not pd.isna(d_val):
                    total_day_date_counts.setdefault(h, {})
                    total_day_date_counts[h][d_val] = total_day_date_counts[h].get(d_val, 0) + 1
                created_d = None
                if "CREATED DATE" in df_z.columns:
                    cd = pd.to_datetime(row_z.get("CREATED DATE"), dayfirst=True,
                                        format="mixed", errors="coerce")
                    if not pd.isna(cd):
                        created_d = cd.date()
                if created_d and (today_date - created_d).days > 1:
                    if h not in total_urgent_counts:
                        total_urgent_counts[h] = {"Pickup": 0, "Delivery": 0, "Transit": 0, "Branch": 0}
                    total_urgent_counts[h][rn] = total_urgent_counts[h].get(rn, 0) + 1
                    urgent_by_type[rn] += 1

        overall = result["overall_counts"]
        # Ensure urgent count cannot exceed category total
        for rn in ["Pickup", "Delivery", "Transit", "Branch"]:
            urgent_by_type[rn] = min(urgent_by_type[rn], overall.get(rn, 0))

        grand_total = sum(overall.values())
        total_urgent_sum = sum(urgent_by_type.values())

        result["summary_caption"] = "\n".join([
            f"☀️ {zone_label} MORNING Report (Age -12h)  {datetime.now().strftime('%d/%m/%Y %H:%M')}",
            f"Delivery: {overall.get('Delivery',0)} (U:{urgent_by_type.get('Delivery',0)})  |  "
            f"Assign Deliver: {overall.get('Branch',0)} (U:{urgent_by_type.get('Branch',0)})  |  "
            f"Pickup: {overall.get('Pickup',0)} (U:{urgent_by_type.get('Pickup',0)})  |  "
            f"Handover to Mega: {overall.get('Transit',0)} (U:{urgent_by_type.get('Transit',0)})",
            f"Grand Total: {grand_total}  |  Total Urgent: {total_urgent_sum}",
            f"📝 Age adjusted: -12 hours (excludes overnight hold)",
        ])

        # 1. Summary image
        img_buf = generate_summary.build_summary_image(
            result["handle_results"],
            result["overall_counts"],
            zone_label=f"{zone_label} MORNING",
            day_date_counts=total_day_date_counts if total_day_date_counts else None,
            urgent_counts=total_urgent_counts if total_urgent_counts else None,
        )
        img_buf.name = "morning_summary.png"
        await send_requester_photo(update, context, img_buf)

        # 2. Total Excel with 12h age adjustment
        label = f"Morning_{zone_label}_" if zone_filter else "Morning_"
        total_xlsx = os.path.join(tmpdir, f"{label}{stamp}.xlsx")
        generate_summary.build_total_excel(result, total_xlsx, age_adjust_hours=12)

        with open(total_xlsx, "rb") as f:
            await send_requester_document(
                update,
                context,
                f,
                os.path.basename(total_xlsx),
                caption=result["summary_caption"],
            )

        await edit_or_send_requester_text(msg, update, context, 
            f"☀️ Done. {zone_label} MORNING report {datetime.now().strftime('%d.%m.%Y %H:%M')}\n"
            f"Age adjusted: -12 hours"
        )

    except Exception as e:
        log.exception("Error in /morning")
        await edit_or_send_requester_text(msg, update, context, f"Error: {e}")


@pm_required_handler
async def cmd_total_kpi(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """/kpi or /totalkpi — Generates overall KPI performance summary report."""
    await delete_group_command(update, context)
    cfg = load_config()
    msg = await send_requester_text(update, context, "⏳ Fetching data for Total KPI Performance Report...")

    tmpdir = tempfile.mkdtemp(prefix="kpi_")
    track_report_dir(tmpdir)
    stamp = datetime.now().strftime("%d.%m_%HH%M")
    src = os.path.join(tmpdir, f"export_{stamp}.xlsx")

    try:
        downloader.download_detail(cfg["api"], src, force_refresh=True)
        import generate_report
        res = generate_report.generate_reports_from_data(
            src,
            "post_office_lookup.csv",
            tmpdir,
            return_metadata=True,
            mode="wide"
        )
        
        # Combine type data
        type_data = res.get("type_data", {})
        df_del = type_data.get("Delivery", pd.DataFrame())
        df_br  = type_data.get("Branch", pd.DataFrame())
        dfs = [d for d in [df_del, df_br] if isinstance(d, pd.DataFrame) and not d.empty]
        df_all = pd.concat(dfs, ignore_index=True) if dfs else pd.DataFrame()

        if df_all.empty:
            await edit_or_send_requester_text(msg, update, context, "⚠️ No active orders found for KPI report.")
            return

        total_orders = len(df_all)
        hits_10h = 0
        special_green = 0
        warning_10_24h = 0
        overdue_24h = 0

        now = datetime.now()

        for _, row in df_all.iterrows():
            sc = str(row.get("STATUS_CODE", "") or "").strip()
            if sc in ("420", "472"):
                # 7-day rule for 420 / 472 special condition
                cd_val = row.get("CREATED DATE")
                is_over_7d = False
                if pd.notna(cd_val):
                    cd = pd.to_datetime(cd_val, dayfirst=True, format="mixed", errors="coerce")
                    if pd.notna(cd) and (now.date() - cd.date()).days > 7:
                        is_over_7d = True
                
                if is_over_7d:
                    overdue_24h += 1
                else:
                    special_green += 1
            else:
                age_str = str(row.get("Age", "") or "")
                match = re.search(r'(\d+)\s*h(?:\s*(\d+)\s*m)?', age_str, re.IGNORECASE)
                if match:
                    h_val = int(match.group(1))
                    m_val = int(match.group(2)) if match.group(2) else 0
                    t_mins = h_val * 60 + m_val
                    if t_mins <= 600:
                        hits_10h += 1
                    else:
                        overdue_24h += 1
                else:
                    hits_10h += 1

        deliv_pct   = (hits_10h / total_orders) * 100
        special_pct = (special_green / total_orders) * 100
        overdue_pct = (overdue_24h / total_orders) * 100
        overall_pct = ((hits_10h + special_green) / total_orders) * 100

        caption = (
            f"📊 *TOTAL KPI PERFORMANCE REPORT — {now.strftime('%d/%m/%Y %H:%M')}*\n"
            f"━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n"
            f"📦 *Total Active Orders* : `{total_orders:,}`\n\n"
            f"🟢 *Delivery KPI (<=10h)*  : `{hits_10h:,}` ({deliv_pct:.1f}%)\n"
            f"🟢 *Special Hold (420/472)*: `{special_green:,}` ({special_pct:.1f}%)\n"
            f"🔴 *Overdue (>10h / >7d)*  : `{overdue_24h:,}` ({overdue_pct:.1f}%)\n"
            f"━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n"
            f"🎯 *OVERALL KPI HIT RATE*  : `{overall_pct:.1f}%` 🟢\n"
            f"  ↳ Courier Delivery KPI: `{deliv_pct:.1f}%`\n"
            f"  ↳ Store Pickup Rate   : `{special_pct:.1f}%`\n"
            f"━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━"
        )

        out_xlsx = res.get("final_xlsx")
        if out_xlsx and os.path.exists(out_xlsx):
            await edit_or_send_requester_text(msg, update, context, caption, parse_mode="Markdown")
            await send_requester_document(update, context, out_xlsx, filename=os.path.basename(out_xlsx), caption="📄 Detailed KPI Report Excel")
        else:
            await edit_or_send_requester_text(msg, update, context, caption, parse_mode="Markdown")

    except Exception as e:
        log.exception("Error generating KPI report: %s", e)
        await edit_or_send_requester_text(msg, update, context, f"❌ Failed to generate KPI report: {e}")


def build_branch_kpi_excel(df_in, out_file, cfg=None):
    """Builds a painted 9-column 10H KPI summary Excel report for registered main post offices."""
    import openpyxl
    from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
    from datetime import datetime, timedelta
    
    df = df_in.copy()
    po_col = 'POST OFFICE HANDLE' if 'POST OFFICE HANDLE' in df.columns else 'CURRENT POST OFFICE'
    date_col = 'CREATED DATE' if 'CREATED DATE' in df.columns else 'CURRENT TIME'
    
    if 'Parsed_Date' not in df.columns:
        df['Parsed_Date'] = pd.to_datetime(df[date_col], dayfirst=True, format='mixed', errors='coerce')
    now = datetime.now()
    if 'Age_Hours' not in df.columns:
        df['Age_Hours'] = ((now - df['Parsed_Date']).dt.total_seconds() / 3600.0).fillna(0)
        
    sc = df['CURRENT STATUS'].astype(str).str.extract(r'^(\d{3})')[0] if 'CURRENT STATUS' in df.columns else df.get('STATUS_CODE', pd.Series())
    df['STATUS_CODE'] = sc
    
    completed_mask = df['STATUS_CODE'].isin(['410', '520', '201'])
    green_pending_mask = (~completed_mask) & ((df['Age_Hours'] - 12.0) <= 10.0)
    red_pending_mask = (~completed_mask) & ((df['Age_Hours'] - 12.0) > 10.0)
    
    df['Is_Green'] = completed_mask | green_pending_mask
    df['Is_Red'] = red_pending_mask
    
    today_start = now.replace(hour=0, minute=0, second=0, microsecond=0)
    yesterday_start = today_start - timedelta(days=1)
    week_start = today_start - timedelta(days=today_start.weekday())
    month_start = today_start.replace(day=1)
    
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "10H KPI 9-Column Report"
    ws.views.sheetView[0].showGridLines = True
    
    headers = [
        "Branch Code",
        "Completed Deliveries (410)",
        "Pending Green KPI",
        "Pending Red Over",
        "🎯 Overall Hit Rate %",
        "Today",
        "Yesterday",
        "This Week",
        "This Month"
    ]
    ws.append(headers)
    
    registered = set()
    if cfg:
        fm = cfg.get('telegram', {}).get('forward_mapping', {})
        for b_list in fm.values():
            for b in b_list:
                if b and str(b).strip():
                    registered.add(str(b).strip().upper())
        zm = cfg.get('zone_mapping', {}).get('by_post_office', {})
        for b in zm.keys():
            if b and str(b).strip():
                registered.add(str(b).strip().upper())
                
    if po_col in df.columns:
        for b in df[po_col].dropna().unique():
            b_str = str(b).strip().upper()
            if b_str and len(b_str) >= 4 and not (len(b_str) >= 4 and b_str[3] in ('A', 'S')):
                registered.add(b_str)

    if registered:
        branches = sorted(list(registered))
    else:
        branches = sorted([str(b) for b in df[po_col].dropna().unique() if str(b).strip()])
    
    def format_period_str(sub):
        g = int(sub['Is_Green'].sum())
        r = int(sub['Is_Red'].sum())
        tot = len(sub)
        rate = (g / tot * 100.0) if tot > 0 else 100.0
        return f"🟢 {g:,} | 🔴 {r:,} ({rate:.1f}% hit)", rate

    green_fill = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
    green_font = Font(name="Calibri", size=10, bold=True, color="006100")

    red_fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
    red_font = Font(name="Calibri", size=10, bold=True, color="9C0006")

    header_fill = PatternFill(start_color="1B365D", end_color="1B365D", fill_type="solid")
    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")

    thin_border = Border(
        left=Side(style='thin', color='D9D9D9'),
        right=Side(style='thin', color='D9D9D9'),
        top=Side(style='thin', color='D9D9D9'),
        bottom=Side(style='thin', color='D9D9D9')
    )

    for col_idx, h in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col_idx)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")

    for b in branches:
        b_df = df[df[po_col].astype(str).str.upper().str.contains(b, na=False)]
        
        b_comp = b_df[b_df['STATUS_CODE'].isin(['410', '520', '201'])]
        b_pend = b_df[~b_df['STATUS_CODE'].isin(['410', '520', '201'])]
        
        comp_g = int(b_comp['Is_Green'].sum())
        comp_r = int(b_comp['Is_Red'].sum())
        pend_g = int(b_pend['Is_Green'].sum())
        pend_r = int(b_pend['Is_Red'].sum())
        
        tot_g = comp_g + pend_g
        tot_r = comp_r + pend_r
        tot_all = len(b_df)
        overall_rate = (tot_g / tot_all * 100.0) if tot_all > 0 else 100.0
        
        td_str, td_rate = format_period_str(b_df[b_df['Parsed_Date'] >= today_start])
        yd_str, yd_rate = format_period_str(b_df[(b_df['Parsed_Date'] >= yesterday_start) & (b_df['Parsed_Date'] < today_start)])
        wk_str, wk_rate = format_period_str(b_df[b_df['Parsed_Date'] >= week_start])
        mo_str, mo_rate = format_period_str(b_df[b_df['Parsed_Date'] >= month_start])
        
        row_idx = ws.max_row + 1
        ws.append([
            b,
            f"🟢 {comp_g:,} | 🔴 {comp_r:,}",
            f"🟢 {pend_g:,}",
            f"🔴 {pend_r:,}",
            f"{overall_rate:.1f}%",
            td_str,
            yd_str,
            wk_str,
            mo_str
        ])
        
        # Paint background fills and font colors
        for c in range(1, 10):
            ws.cell(row=row_idx, column=c).border = thin_border
            ws.cell(row=row_idx, column=c).alignment = Alignment(horizontal="left" if c == 1 else "center", vertical="center")
            
        ws.cell(row=row_idx, column=1).font = Font(name="Calibri", size=10, bold=True)
        
        # Col 2: Completed 410 (Always Green)
        ws.cell(row=row_idx, column=2).fill = green_fill
        ws.cell(row=row_idx, column=2).font = green_font
        
        # Col 3: Pending Green
        ws.cell(row=row_idx, column=3).fill = green_fill
        ws.cell(row=row_idx, column=3).font = green_font
        
        # Col 4: Pending Red
        ws.cell(row=row_idx, column=4).fill = red_fill
        ws.cell(row=row_idx, column=4).font = red_font
        
        # Col 5: Overall Rate
        ws.cell(row=row_idx, column=5).fill = green_fill if overall_rate >= 80.0 else red_fill
        ws.cell(row=row_idx, column=5).font = green_font if overall_rate >= 80.0 else red_font
        
        # Col 6: Today
        ws.cell(row=row_idx, column=6).fill = green_fill if td_rate >= 80.0 else red_fill
        ws.cell(row=row_idx, column=6).font = green_font if td_rate >= 80.0 else red_font
        
        # Col 7: Yesterday
        ws.cell(row=row_idx, column=7).fill = green_fill if yd_rate >= 80.0 else red_fill
        ws.cell(row=row_idx, column=7).font = green_font if yd_rate >= 80.0 else red_font
        
        # Col 8: This Week
        ws.cell(row=row_idx, column=8).fill = green_fill if wk_rate >= 80.0 else red_fill
        ws.cell(row=row_idx, column=8).font = green_font if wk_rate >= 80.0 else red_font
        
        # Col 9: This Month
        ws.cell(row=row_idx, column=9).fill = green_fill if mo_rate >= 80.0 else red_fill
        ws.cell(row=row_idx, column=9).font = green_font if mo_rate >= 80.0 else red_font

    for col in ws.columns:
        max_len = max(len(str(cell.value or '')) for cell in col)
        col_letter = openpyxl.utils.get_column_letter(col[0].column)
        ws.column_dimensions[col_letter].width = max(max_len + 4, 18)

    wb.save(out_file)
    return out_file


async def cmd_kpi10h(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    """Shows concise 10H KPI breakdown (/kpi) with Today, Yesterday, This Week, and This Month stats."""
    import time
    from datetime import datetime, timedelta
    args = context.args or []
    branch_target = args[0].strip().upper() if args else None
    
    msg = await send_requester_text(update, context, "📊 Calculating 10H KPI performance report...")
    
    try:
        cfg = load_config()
        latest_file = os.path.join(os.path.dirname(os.path.abspath(__file__)), "cache", "latest_detail.xlsx")
        
        if not os.path.exists(latest_file) or (time.time() - os.path.getmtime(latest_file) > 300):
            import downloader
            downloader.download_detail(cfg["api"], latest_file, force_refresh=True)
            
        df = pd.read_excel(latest_file)
        
        if 'CURRENT STATUS' in df.columns:
            sc = df['CURRENT STATUS'].astype(str).str.extract(r'^(\d{3})')[0]
            df['STATUS_CODE'] = sc
        else:
            df['STATUS_CODE'] = ''
            
        po_col = 'POST OFFICE HANDLE' if 'POST OFFICE HANDLE' in df.columns else 'CURRENT POST OFFICE'
        
        # Calculate age in hours and parse creation date
        date_col = 'CREATED DATE' if 'CREATED DATE' in df.columns else 'CURRENT TIME'
        parsed_dates = pd.to_datetime(df[date_col], dayfirst=True, format='mixed', errors='coerce')
        df['Parsed_Date'] = parsed_dates
        now = datetime.now()
        ages_hours = (now - parsed_dates).dt.total_seconds() / 3600.0
        df['Age_Hours'] = ages_hours.fillna(0)
        
        # Generate registered main branches Excel report file
        out_excel = os.path.join(os.path.dirname(os.path.abspath(__file__)), "cache", "Branch_KPI_Breakdown_Report.xlsx")
        build_branch_kpi_excel(df, out_excel, cfg)
        
        if branch_target and branch_target != "TOTAL":
            df = df[df[po_col].astype(str).str.upper().str.contains(branch_target, na=False)].copy()
            
        n_total = len(df)
        if n_total == 0:
            await edit_or_send_requester_text(msg, update, context, f"ℹ️ No orders found for '{branch_target or 'All Branches'}'.")
            return
            
        completed_mask = df['STATUS_CODE'].isin(['410', '520', '201'])
        green_pending_mask = (~completed_mask) & ((df['Age_Hours'] - 12.0) <= 10.0)
        red_pending_mask = (~completed_mask) & ((df['Age_Hours'] - 12.0) > 10.0)
        
        # Classify overall Green vs Red
        df['Is_Green'] = completed_mask | green_pending_mask
        df['Is_Red'] = red_pending_mask
        
        comp_g = int(completed_mask.sum())
        comp_r = 0  # Completed 410/520 delivered count as 10h hit
        pend_g = int(green_pending_mask.sum())
        pend_r = int(red_pending_mask.sum())
        
        tot_green = comp_g + pend_g
        tot_red = pend_r
        overall_rate = (tot_green / n_total * 100.0) if n_total > 0 else 100.0
        
        # Time breakdowns
        today_start = now.replace(hour=0, minute=0, second=0, microsecond=0)
        yesterday_start = today_start - timedelta(days=1)
        week_start = today_start - timedelta(days=today_start.weekday())
        month_start = today_start.replace(day=1)
        
        def get_period_stats(start_dt, end_dt=None):
            if end_dt:
                sub = df[(df['Parsed_Date'] >= start_dt) & (df['Parsed_Date'] < end_dt)]
            else:
                sub = df[df['Parsed_Date'] >= start_dt]
            g = int(sub['Is_Green'].sum())
            r = int(sub['Is_Red'].sum())
            tot = len(sub)
            rate = (g / tot * 100.0) if tot > 0 else 100.0
            return g, r, tot, rate
            
        td_g, td_r, td_tot, td_rate = get_period_stats(today_start)
        yd_g, yd_r, yd_tot, yd_rate = get_period_stats(yesterday_start, today_start)
        wk_g, wk_r, wk_tot, wk_rate = get_period_stats(week_start)
        mo_g, mo_r, mo_tot, mo_rate = get_period_stats(month_start)
        
        title = f"📊 *10H KPI SUMMARY REPORT*" + (f" ({branch_target})" if branch_target else "")
        resp = (
            f"{title}\n"
            f"━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n"
            f"✅ *Completed Deliveries (410)*: 🟢 <=10h: `{comp_g:,}` | 🔴 >10h: `{comp_r:,}`\n"
            f"📦 *Pending Green KPI*        : 🟢 <=10h: `{pend_g:,}`\n"
            f"📦 *Pending Red Over*         : 🔴 >10h: `{pend_r:,}`\n"
            f"━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n"
            f"🎯 *OVERALL 10H KPI HIT RATE*: `{overall_rate:.1f}%` " + ("🟢" if overall_rate >= 80 else "🔴") + "\n\n"
            f"📅 *TIME PERIOD BREAKDOWN*:\n"
            f"• *Today*      : 🟢 <=10h: `{td_g:,}` | 🔴 >10h: `{td_r:,}` (`{td_rate:.1f}%` hit)\n"
            f"• *Yesterday*  : 🟢 <=10h: `{yd_g:,}` | 🔴 >10h: `{yd_r:,}` (`{yd_rate:.1f}%` hit)\n"
            f"• *This Week*  : 🟢 <=10h: `{wk_g:,}` | 🔴 >10h: `{wk_r:,}` (`{wk_rate:.1f}%` hit)\n"
            f"• *This Month* : 🟢 <=10h: `{mo_g:,}` | 🔴 >10h: `{mo_r:,}` (`{mo_rate:.1f}%` hit)\n"
            f"━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n"
            f"_Note: Counts completed 410 + pending with 12h overnight hold adjustment._"
        )
        
        await edit_or_send_requester_text(msg, update, context, resp, parse_mode="Markdown")
        
        # Send Excel file strictly to requester only (zero group forwarding)
        if os.path.exists(out_excel):
            with open(out_excel, "rb") as doc_file:
                await send_requester_document(update, context, document=doc_file, filename="Branch_KPI_Breakdown_Report.xlsx", caption="📊 36-Branch KPI Breakdown Report (Excel)")
            
    except Exception as e:
        log.exception("Error in cmd_kpi10h: %s", e)
        await edit_or_send_requester_text(msg, update, context, f"❌ Error computing 10H KPI: {e}")


def build_tpg_excel(branch_counts, type_data, out_path):
    """Builds a dedicated, beautifully formatted Excel workbook for /tpg command."""
    import pandas as pd
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
    from openpyxl.utils import get_column_letter

    wb = Workbook()
    ws_sum = wb.active
    ws_sum.title = "TPG Summary"

    font_family = "Calibri"
    f_title = Font(name=font_family, size=14, bold=True, color="FFFFFF")
    f_header = Font(name=font_family, size=11, bold=True, color="FFFFFF")
    f_data = Font(name=font_family, size=10)
    f_tot = Font(name=font_family, size=11, bold=True, color="991B1B")

    fill_title = PatternFill(start_color="1E293B", end_color="1E293B", fill_type="solid")
    fill_hdr = PatternFill(start_color="334155", end_color="334155", fill_type="solid")
    fill_tot = PatternFill(start_color="FEE2E2", end_color="FEE2E2", fill_type="solid")
    fill_alt = PatternFill(start_color="F8FAFC", end_color="F8FAFC", fill_type="solid")

    thin = Side(border_style="thin", color="CBD5E1")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    # Title row
    ws_sum.merge_cells("A1:F1")
    t_cell = ws_sum.cell(1, 1, "TPG BRANCH DAILY vs WEEKLY COMPARISON REPORT")
    t_cell.font = f_title
    t_cell.fill = fill_title
    t_cell.alignment = Alignment(horizontal="center", vertical="center")
    ws_sum.row_dimensions[1].height = 30

    # Headers
    headers = ["POST OFFICE HANDLE", "DAILY VOLUME", "WEEKLY VOLUME", "SPECIAL HOLD (420/472)", "OVERDUE (>24h)", "DAILY SHARE (%)"]
    ws_sum.row_dimensions[3].height = 24

    for c_idx, h_text in enumerate(headers, 1):
        cell = ws_sum.cell(3, c_idx, h_text)
        cell.font = f_header
        cell.fill = fill_hdr
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = border

    r_idx = 4
    tot_d = tot_w = tot_h = tot_o = 0

    for h, counts in sorted(branch_counts.items(), key=lambda x: x[1]["Weekly"], reverse=True):
        d = counts["Daily"]
        w = counts["Weekly"]
        sh = counts["Hold"]
        o = counts["Overdue"]
        pct = (d / w * 100) if w > 0 else 0.0

        tot_d += d; tot_w += w; tot_h += sh; tot_o += o

        row_fill = fill_alt if r_idx % 2 == 0 else None
        row_vals = [h, d, w, sh, o, f"{pct:.1f}%"]
        ws_sum.row_dimensions[r_idx].height = 20

        for c_idx, val in enumerate(row_vals, 1):
            cell = ws_sum.cell(r_idx, c_idx, val)
            cell.font = f_data
            cell.border = border
            if row_fill:
                cell.fill = row_fill
            cell.alignment = Alignment(horizontal="left" if c_idx == 1 else "center", vertical="center")

        r_idx += 1

    # Totals Row
    tot_pct = (tot_d / tot_w * 100) if tot_w > 0 else 0.0
    tot_vals = ["GRAND TOTAL", tot_d, tot_w, tot_h, tot_o, f"{tot_pct:.1f}%"]
    ws_sum.row_dimensions[r_idx].height = 22

    for c_idx, val in enumerate(tot_vals, 1):
        cell = ws_sum.cell(r_idx, c_idx, val)
        cell.font = f_tot
        cell.fill = fill_tot
        cell.border = border
        cell.alignment = Alignment(horizontal="left" if c_idx == 1 else "center", vertical="center")

    # Sheet 2: Itemized Data
    ws_item = wb.create_sheet(title="Itemized Orders")
    item_headers = ["TYPE", "POST OFFICE HANDLE", "ORDER ID", "RECEIVER", "STATUS CODE", "AGE", "TOTAL FEE (USD)", "COD (USD)"]
    ws_item.row_dimensions[1].height = 24

    for c_idx, h_text in enumerate(item_headers, 1):
        cell = ws_item.cell(1, c_idx, h_text)
        cell.font = f_header
        cell.fill = fill_hdr
        cell.alignment = Alignment(horizontal="center", vertical="center")

    item_r = 2
    for t_name, df_type in type_data.items():
        if not isinstance(df_type, pd.DataFrame) or df_type.empty:
            continue
        handle_col = "POST OFFICE HANDLE" if "POST OFFICE HANDLE" in df_type.columns else ("CURRENT POST OFFICE" if "CURRENT POST OFFICE" in df_type.columns else None)
        if not handle_col:
            continue
        
        for _, r in df_type.iterrows():
            h_val = str(r.get(handle_col, "") or "").strip().upper()
            if not h_val or h_val in ("NAN", "GRAND TOTAL"):
                continue
            
            oid = str(r.get("ORDER ID", "") or "").strip()
            rec = str(r.get("RECEIVER", "") or r.get("Cus name", "") or "").strip()
            sc  = str(r.get("STATUS_CODE", "") or "").strip()
            age = str(r.get("Age", "") or "").strip()
            fee = r.get("TOTAL FEE (USD)", 0)
            cod = r.get("COD (USD)", 0)

            ws_item.append([t_name, h_val, oid, rec, sc, age, fee, cod])
            item_r += 1

    # Auto-adjust column widths
    for ws in [ws_sum, ws_item]:
        for col in ws.columns:
            max_len = max(len(str(cell.value or "")) for cell in col)
            col_letter = get_column_letter(col[0].column)
            ws.column_dimensions[col_letter].width = max(max_len + 4, 12)

    wb.save(out_path)


@pm_required_handler
async def cmd_tpg(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """/tpg — TPG Branch Operational Summary comparing Daily & Weekly volume per branch with Excel attachment."""
    await delete_group_command(update, context)
    cfg = load_config()
    msg = await send_requester_text(update, context, "⏳ Fetching data for TPG Daily & Weekly Branch Comparison...")

    tmpdir = tempfile.mkdtemp(prefix="tpg_")
    track_report_dir(tmpdir)
    stamp = datetime.now().strftime("%d.%m_%HH%M")
    src = os.path.join(tmpdir, f"export_{stamp}.xlsx")

    try:
        downloader.download_detail(cfg["api"], src, force_refresh=True)
        import generate_report
        res = generate_report.generate_reports_from_data(
            src,
            "post_office_lookup.csv",
            tmpdir,
            return_metadata=True,
            mode="wide"
        )

        type_data = res.get("type_data", {})
        branch_counts = {}
        now = datetime.now()

        for t_name, df_type in type_data.items():
            if not isinstance(df_type, pd.DataFrame) or df_type.empty:
                continue
            
            handle_col = "POST OFFICE HANDLE" if "POST OFFICE HANDLE" in df_type.columns else ("CURRENT POST OFFICE" if "CURRENT POST OFFICE" in df_type.columns else None)
            if not handle_col:
                continue
            
            for _, r in df_type.iterrows():
                h_val = str(r.get(handle_col, "") or "").strip().upper()
                if not h_val or h_val in ("NAN", "GRAND TOTAL"):
                    continue
                
                if h_val not in branch_counts:
                    branch_counts[h_val] = {
                        "Daily": 0, "Weekly": 0, "Hold": 0, "Overdue": 0
                    }
                
                # All active items count towards Weekly total
                branch_counts[h_val]["Weekly"] += 1

                # Check age for Daily (<=24h)
                age_str = str(r.get("Age", "") or "")
                match = re.search(r'(\d+)\s*h(?:\s*(\d+)\s*m)?', age_str, re.IGNORECASE)
                h_val_num = int(match.group(1)) if match else 0
                
                if h_val_num <= 24:
                    branch_counts[h_val]["Daily"] += 1

                sc = str(r.get("STATUS_CODE", "") or "").strip()
                if sc in ("420", "472"):
                    branch_counts[h_val]["Hold"] += 1
                elif h_val_num > 24:
                    branch_counts[h_val]["Overdue"] += 1

        if not branch_counts:
            await edit_or_send_requester_text(msg, update, context, "⚠️ No active TPG branch data found.")
            return

        lines = [
            f"🏬 *TPG BRANCH DAILY vs WEEKLY COMPARISON*",
            f"📅 `{now.strftime('%d/%m/%Y %H:%M')}`",
            f"━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━",
            f"`BRANCH   │ DAILY │ WEEKLY│ HOLD │ OVERDUE│ DAILY %`",
            f"─────────┼───────┼───────┼──────┼────────┼────────"
        ]

        tot_d = tot_w = tot_h = tot_o = 0

        # Sort by Weekly volume descending
        for h, counts in sorted(branch_counts.items(), key=lambda x: x[1]["Weekly"], reverse=True):
            d = counts["Daily"]
            w = counts["Weekly"]
            sh = counts["Hold"]
            o = counts["Overdue"]
            pct = (d / w * 100) if w > 0 else 0.0

            tot_d += d
            tot_w += w
            tot_h += sh
            tot_o += o

            lines.append(f"`{h:<9}│ {d:<6}│ {w:<6}│ {sh:<5}│ {o:<7}│ {pct:>5.1f}%`")

        tot_pct = (tot_d / tot_w * 100) if tot_w > 0 else 0.0
        lines.append(f"─────────┼───────┼───────┼──────┼────────┼────────")
        lines.append(f"`TOTAL    │ {tot_d:<6}│ {tot_w:<6}│ {tot_h:<5}│ {tot_o:<7}│ {tot_pct:>5.1f}%`")
        lines.append(f"━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━")

        # Chunk text if lines are very long (> 3800 chars) to prevent Telegram length limit errors
        full_text = "\n".join(lines)
        if len(full_text) > 3800:
            chunks = []
            cur_chunk = []
            cur_len = 0
            for line in lines:
                if cur_len + len(line) + 1 > 3800:
                    chunks.append("\n".join(cur_chunk))
                    cur_chunk = [line]
                    cur_len = len(line)
                else:
                    cur_chunk.append(line)
                    cur_len += len(line) + 1
            if cur_chunk:
                chunks.append("\n".join(cur_chunk))
            
            first = True
            for ch in chunks:
                if first:
                    msg = await edit_or_send_requester_text(msg, update, context, ch, parse_mode="Markdown")
                    first = False
                else:
                    await send_requester_text(update, context, ch, parse_mode="Markdown")
        else:
            await edit_or_send_requester_text(msg, update, context, full_text, parse_mode="Markdown")

        # Build dedicated TPG Excel file matching the summary table + itemized orders
        tpg_excel_path = os.path.join(tmpdir, f"TPG_Daily_vs_Weekly_Comparison_{stamp}.xlsx")
        build_tpg_excel(branch_counts, type_data, tpg_excel_path)

        # Render high-resolution PNG summary image for instant Telegram mobile viewing
        try:
            import excel_to_image
            img_buf = excel_to_image.excel_to_image(tpg_excel_path)
            img_buf.name = f"TPG_Summary_{stamp}.png"
            await send_requester_photo(update, context, img_buf)
        except Exception as e_img:
            log.warning("Could not render TPG summary image: %s", e_img)

        if os.path.exists(tpg_excel_path):
            await send_requester_document(update, context, tpg_excel_path, filename=os.path.basename(tpg_excel_path), caption="📄 TPG Daily vs Weekly Comparison Excel Report")

    except Exception as e:
        log.exception("Error in /tpg command: %s", e)
        await edit_or_send_requester_text(msg, update, context, f"❌ Failed to generate TPG summary: {e}")


@pm_required_handler
async def cmd_export_logs(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """/exportlogs or /logs — Exports complete status-divided tracking logs report (01/07/2026 to 03/08/2026), excluding test/trainer/global."""
    await delete_group_command(update, context)
    cfg = load_config()
    msg = await send_requester_text(update, context, "⏳ Extracting tracking logs divided by status code (01/07/2026 - 03/08/2026)...")

    tmpdir = tempfile.mkdtemp(prefix="logs_")
    track_report_dir(tmpdir)
    stamp = datetime.now().strftime("%Y%m%d_%H%M")
    src = os.path.join(tmpdir, f"export_01Jul_03Aug_{stamp}.xlsx")
    out_xlsx = os.path.join(tmpdir, f"Bill_Tracking_Status_Logs_01Jul_03Aug_{stamp}.xlsx")

    try:
        today_str = datetime.now().strftime("%Y%m%d")
        downloader.download_detail(cfg["api"], src, from_date="20260701", to_date=today_str, force_refresh=True)
        import generate_tracking_logs_report
        generate_tracking_logs_report.generate_tracking_logs(src, out_xlsx)

        if os.path.exists(out_xlsx):
            await edit_or_send_requester_text(msg, update, context, "✅ Status-divided tracking logs generated successfully! Attaching Excel file...")
            await send_requester_document(
                update,
                context,
                out_xlsx,
                filename=os.path.basename(out_xlsx),
                caption="📄 Bill Tracking Status Logs Excel Report (01/07 - 03/08)"
            )
        else:
            await edit_or_send_requester_text(msg, update, context, "❌ Could not generate tracking logs file.")
    except Exception as e:
        log.exception("Error in /exportlogs command: %s", e)
        await edit_or_send_requester_text(msg, update, context, f"❌ Failed to generate tracking logs: {e}")


@pm_required_handler
async def cmd_compare(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """/compare — Urgent Bill Status Clearance & Shift Comparison Report (9 AM -> 2 PM -> 5 PM)."""
    await delete_group_command(update, context)
    cfg = load_config()
    
    shift_target = None
    if context.args:
        arg_val = str(context.args[0]).strip()
        if arg_val == "1":
            shift_target = "9AM"
        elif arg_val == "2":
            shift_target = "2PM"
        elif arg_val == "3":
            shift_target = "5PM"

    msg = await send_requester_text(update, context, "⏳ Fetching data and processing Urgent Bill Shift Comparison...")

    tmpdir = tempfile.mkdtemp(prefix="compare_")
    track_report_dir(tmpdir)
    stamp = datetime.now().strftime("%Y%m%d_%H%M")
    src = os.path.join(tmpdir, f"export_compare_{stamp}.xlsx")
    today_str = datetime.now().strftime("%d/%m/%Y")

    try:
        downloader.download_detail(cfg["api"], src, force_refresh=True)
        import generate_report, compare_manager

        df_detail = pd.read_excel(src)
        df_detail.columns = [str(c).strip().upper() for c in df_detail.columns]

        rows, totals, itemized = compare_manager.build_comparison_summary(today_str, df_detail, target_shift=shift_target)

        if not rows:
            await edit_or_send_requester_text(msg, update, context, "⚠️ No urgent bill comparison data found for today.")
            return

        current_shift = shift_target or compare_manager.determine_shift()
        label = "Full Day Total (5 PM)" if current_shift == "5PM" else ("Afternoon (2 PM)" if current_shift == "2PM" else "Morning (9 AM)")
        
        out_excel = os.path.join("compare", f"Urgent_Clearance_Compare_{stamp}.xlsx")
        compare_manager.build_compare_excel(today_str, rows, totals, itemized, out_excel)

        if os.path.exists(out_excel):
            await edit_or_send_requester_text(msg, update, context, f"✅ Urgent Bill Shift Comparison Report (`{today_str}` | Mode: `{label}`) generated successfully! Attaching Excel file...")
            await send_requester_document(
                update,
                context,
                out_excel,
                filename=os.path.basename(out_excel),
                caption=f"📄 Urgent Bill Shift Comparison Excel Report ({today_str} | Shift {current_shift})"
            )
        else:
            await edit_or_send_requester_text(msg, update, context, "❌ Could not generate comparison report file.")

    except Exception as e:
        log.exception("Error in /compare command: %s", e)
        await edit_or_send_requester_text(msg, update, context, f"❌ Failed to generate comparison report: {e}")


async def run_time_vs(update: Update, context: ContextTypes.DEFAULT_TYPE, start_hour: int, end_hour: int, command_label: str):
    await delete_group_command(update, context)

    cfg = load_config()
    allowed = cfg["telegram"].get("allowed_chat_ids") or []
    chat_id = update.effective_chat.id
    if allowed and chat_id not in allowed:
        await private_or_current_reply(update, context, "This chat is not allowed to use the bot.")
        return

    # Determine which handles to show based on arguments or chat context
    args = [a.strip().upper() for a in (context.args or []) if a.strip()]
    target_handles = None
    title_label = ""

    if args:
        arg = args[0]
        total_zones = cfg.get("total_zones", {})
        if arg.lower() in total_zones:
            target_handles = [h.upper() for h in total_zones[arg.lower()]]
            title_label = f"{arg} "
        else:
            target_handles = [arg]
            title_label = f"{arg} "
    else:
        # Check if in a registered group
        forward_mapping = get_forward_mapping(cfg)
        group_id_str = str(chat_id)
        if group_id_str in forward_mapping:
            group_handles = forward_mapping[group_id_str]
            if "*" not in group_handles:
                target_handles = [h.upper() for h in group_handles if h]
                title_label = f"{', '.join(target_handles)} "

        # Check if in a zone group
        zone_fwd_map = cfg.get("zone_forward_mapping", {})
        if group_id_str in zone_fwd_map:
            zone_key = zone_fwd_map[group_id_str]
            total_zones = cfg.get("total_zones", {})
            if zone_key in total_zones:
                target_handles = [h.upper() for h in total_zones[zone_key]]
                title_label = f"{zone_key.upper()} "

    # ── Try reading from JSON history first ─────────────────────────────────────
    history_path = os.path.join(HERE, "highlight_history.json")
    if os.path.exists(history_path):
        try:
            with open(history_path, encoding="utf-8") as f:
                history = json.load(f)

            today_str = datetime.now().strftime("%Y-%m-%d")
            runs = history.get(today_str, [])
            if runs:
                run_start = None
                run_end = None
                diff_start = None
                diff_end = None

                for r in runs:
                    dt = datetime.strptime(r["timestamp"], "%Y-%m-%d %H:%M:%S")
                    target_start = datetime.combine(dt.date(), datetime.min.time().replace(hour=start_hour))
                    target_end = datetime.combine(dt.date(), datetime.min.time().replace(hour=end_hour))

                    # start matching
                    is_valid_start = False
                    if start_hour == 8:
                        is_valid_start = (dt.hour < 12)
                    elif start_hour == 14:
                        is_valid_start = (dt.hour >= 12 and dt.hour < 16)

                    if is_valid_start:
                        d_start = abs((dt - target_start).total_seconds())
                        if diff_start is None or d_start < diff_start:
                            diff_start = d_start
                            run_start = r

                    # end matching
                    is_valid_end = False
                    if end_hour == 14:
                        is_valid_end = (dt.hour >= 12)
                    elif end_hour == 17:
                        is_valid_end = (dt.hour >= 16)

                    if is_valid_end:
                        d_end = abs((dt - target_end).total_seconds())
                        if diff_end is None or d_end < diff_end:
                            diff_end = d_end
                            run_end = r

                # Fallbacks
                if not run_start:
                    run_start = runs[0]
                if not run_end:
                    latest = runs[-1]
                    if latest["timestamp"] != run_start["timestamp"]:
                        run_end = latest

                if run_start and run_end and run_start["timestamp"] != run_end["timestamp"]:
                    dt_start = datetime.strptime(run_start["timestamp"], "%Y-%m-%d %H:%M:%S")
                    dt_end = datetime.strptime(run_end["timestamp"], "%Y-%m-%d %H:%M:%S")

                    t_start = dt_start.strftime('%I:%M %p')
                    t_end = dt_end.strftime('%I:%M %p')

                    all_h = set()
                    map_1 = run_start["data"]
                    map_2 = run_end["data"]

                    for rn in ['Pickup', 'Delivery', 'Transit', 'Branch']:
                        all_h.update(map_1.get(rn, {}).keys())
                        all_h.update(map_2.get(rn, {}).keys())

                    all_h = sorted(list(h for h in all_h if str(h).strip()))
                    if target_handles:
                        all_h = [h for h in all_h if h in target_handles]

                    if all_h:
                        text_lines = [
                            f"📊 {title_label}VS REPORT ({command_label}) — {t_start} vs {t_end}",
                            "=============================="
                        ]

                        grand_sets_1 = {"Pickup": set(), "Delivery": set(), "Transit": set(), "Branch": set()}
                        grand_sets_2 = {"Pickup": set(), "Delivery": set(), "Transit": set(), "Branch": set()}

                        for h in all_h:
                            h_lines = []
                            has_any_data = False
                            
                            for rn in ['Pickup', 'Delivery', 'Transit', 'Branch']:
                                set_1 = set(map_1.get(rn, {}).get(h, []))
                                set_2 = set(map_2.get(rn, {}).get(h, []))
                                
                                grand_sets_1[rn].update(set_1)
                                grand_sets_2[rn].update(set_2)
                                
                                n1 = len(set_1)
                                n2 = len(set_2)
                                cleared = len(set_1 - set_2)
                                
                                if n1 > 0 or n2 > 0:
                                    has_any_data = True
                                    h_lines.append(f"  • {rn}: {n1} vs {n2} (Cleared {cleared})")

                            if has_any_data and len(all_h) <= 10:
                                text_lines.append(f"\n🏢 {h}:")
                                text_lines.extend(h_lines)

                        # Grand Summary
                        text_lines.append("\n==============================")
                        text_lines.append("📈 GRAND SUMMARY:")
                        g_1 = 0
                        g_2 = 0
                        g_clear = 0
                        for rn in ['Pickup', 'Delivery', 'Transit', 'Branch']:
                            s1 = grand_sets_1[rn]
                            s2 = grand_sets_2[rn]
                            n1 = len(s1)
                            n2 = len(s2)
                            cleared = len(s1 - s2)
                            
                            g_1 += n1
                            g_2 += n2
                            g_clear += cleared
                            
                            text_lines.append(f"  • {rn}: {n1} vs {n2} (Cleared {cleared})")

                        text_lines.append(f"  • Total: {g_1} vs {g_2} (Cleared {g_clear})")

                        await send_requester_text(update, context, "\n".join(text_lines))
                        return
        except Exception as e:
            log.warning(f"Error checking highlight history JSON: {e}")

    if not os.path.exists(REPORTS_LOG_PATH):
        await private_or_current_reply(update, context, "No reports run today yet.")
        return

    try:
        with open(REPORTS_LOG_PATH, encoding="utf-8") as f:
            log_data = json.load(f)
    except Exception as e:
        await private_or_current_reply(update, context, f"Error reading reports log: {e}")
        return

    today_str = datetime.now().strftime("%Y-%m-%d")
    dirs = log_data.get(today_str, [])
    if not dirs:
        await private_or_current_reply(update, context, "No reports run today yet.")
        return

    # Find all export files today
    exports = []
    for d in dirs:
        if os.path.exists(d):
            for f in os.listdir(d):
                if f.startswith("export_") and f.endswith(".xlsx"):
                    fpath = os.path.join(d, f)
                    mtime = os.path.getmtime(fpath)
                    mtime_dt = datetime.fromtimestamp(mtime)
                    exports.append((mtime_dt, fpath))

    if not exports:
        await private_or_current_reply(update, context, "No raw exports found for today.")
        return

    exports = sorted(exports, key=lambda x: x[0])
    
    file_start = None
    file_end = None
    diff_start = None
    diff_end = None

    today_dt = exports[0][0].date()
    target_start = datetime.combine(today_dt, datetime.min.time().replace(hour=start_hour))
    target_end = datetime.combine(today_dt, datetime.min.time().replace(hour=end_hour))

    for dt, path in exports:
        # start file matching
        is_valid_start = False
        if start_hour == 8:
            is_valid_start = (dt.hour < 12)
        elif start_hour == 14:
            is_valid_start = (dt.hour >= 12 and dt.hour < 16)

        if is_valid_start:
            d_start = abs((dt - target_start).total_seconds())
            if diff_start is None or d_start < diff_start:
                diff_start = d_start
                file_start = (dt, path)

        # end file matching
        is_valid_end = False
        if end_hour == 14:
            is_valid_end = (dt.hour >= 12)
        elif end_hour == 17:
            is_valid_end = (dt.hour >= 16)

        if is_valid_end:
            d_end = abs((dt - target_end).total_seconds())
            if diff_end is None or d_end < diff_end:
                diff_end = d_end
                file_end = (dt, path)

    # Fallbacks
    if not file_start:
        file_start = exports[0]

    if not file_end:
        latest = exports[-1]
        if latest[1] != file_start[1]:
            file_end = latest

    if not file_end:
        await private_or_current_reply(
            update,
            context,
            f"Only one report has been run today (at {file_start[0].strftime('%H:%M')}). "
            "Please run another push first to compare."
        )
        return

    msg = await send_requester_text(update, context, "Calculating differences...")

    try:
        tmpdir = tempfile.mkdtemp(prefix="vs_")
        res_1 = generate_report.generate_reports_from_data(
            file_start[1], REF_PATH, tmpdir, return_metadata=True, mode="wide"
        )
        res_2 = generate_report.generate_reports_from_data(
            file_end[1], REF_PATH, tmpdir, return_metadata=True, mode="wide"
        )
        # Clean up
        for f in os.listdir(tmpdir):
            try:
                os.remove(os.path.join(tmpdir, f))
            except Exception:
                pass
        try:
            os.rmdir(tmpdir)
        except Exception:
            pass
    except Exception as e:
        await edit_or_send_requester_text(msg, update, context, f"Error processing reports: {e}")
        return

    today_date = datetime.now().date()
    filter_cols = {
        'Pickup': 'POST OFFICE HANDLE',
        'Delivery': 'POST OFFICE HANDLE',
        'Transit': 'POST OFFICE HANDLE',
            'Branch': 'POST OFFICE HANDLE'
    }

    # Aggregate by handle
    all_h = set()
    for rn in ['Pickup', 'Delivery', 'Transit', 'Branch']:
        df1 = res_1['type_data'].get(rn, pd.DataFrame())
        df2 = res_2['type_data'].get(rn, pd.DataFrame())
        fcol = filter_cols[rn]
        if fcol in df1.columns:
            all_h.update(df1[fcol].dropna().unique())
        if fcol in df2.columns:
            all_h.update(df2[fcol].dropna().unique())

    all_h = sorted(list(h for h in all_h if str(h).strip()))
    if target_handles:
        all_h = [h for h in all_h if h in target_handles]

    if not all_h:
        await edit_or_send_requester_text(msg, update, context, "No matching handles found in today's data.")
        return

    t_start = file_start[0].strftime('%I:%M %p')
    t_end = file_end[0].strftime('%I:%M %p')

    text_lines = [
        f"📊 {title_label}VS REPORT ({command_label}) — {t_start} vs {t_end}",
        "=============================="
    ]

    grand_sets_1 = {"Pickup": set(), "Delivery": set(), "Transit": set(), "Branch": set()}
    grand_sets_2 = {"Pickup": set(), "Delivery": set(), "Transit": set(), "Branch": set()}

    for h in all_h:
        h_lines = []
        has_any_data = False
        
        for rn in ['Pickup', 'Delivery', 'Transit', 'Branch']:
            df1 = res_1['type_data'].get(rn, pd.DataFrame())
            df2 = res_2['type_data'].get(rn, pd.DataFrame())
            fcol = filter_cols[rn]
            
            df1_h = df1[df1[fcol] == h] if fcol in df1.columns else pd.DataFrame()
            df2_h = df2[df2[fcol] == h] if fcol in df2.columns else pd.DataFrame()
            
            set_1 = get_highlighted_order_ids(df1_h, today_date)
            set_2 = get_highlighted_order_ids(df2_h, today_date)
            
            grand_sets_1[rn].update(set_1)
            grand_sets_2[rn].update(set_2)
            
            n1 = len(set_1)
            n2 = len(set_2)
            cleared = len(set_1 - set_2)
            
            if n1 > 0 or n2 > 0:
                has_any_data = True
                h_lines.append(f"  • {rn}: {n1} vs {n2} (Cleared {cleared})")

        if has_any_data and len(all_h) <= 10:
            text_lines.append(f"\n🏢 {h}:")
            text_lines.extend(h_lines)

    # Grand Summary
    text_lines.append("\n==============================")
    text_lines.append("📈 GRAND SUMMARY:")
    g_1 = 0
    g_2 = 0
    g_clear = 0
    for rn in ['Pickup', 'Delivery', 'Transit', 'Branch']:
        s1 = grand_sets_1[rn]
        s2 = grand_sets_2[rn]
        n1 = len(s1)
        n2 = len(s2)
        cleared = len(s1 - s2)
        
        g_1 += n1
        g_2 += n2
        g_clear += cleared
        
        text_lines.append(f"  • {rn}: {n1} vs {n2} (Cleared {cleared})")

    text_lines.append(f"  • Total: {g_1} vs {g_2} (Cleared {g_clear})")

    await edit_or_send_requester_text(msg, update, context, "\n".join(text_lines))


@pm_required_handler
async def cmd_vs(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """/vs [handle/zone] — compare morning (8 AM) and afternoon (2 PM/current) reports."""
    await run_time_vs(update, context, start_hour=8, end_hour=14, command_label="8AM vs 2PM")


@pm_required_handler
async def cmd_vs2(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """/vs2 [handle/zone] — compare afternoon (2 PM) and evening (5 PM/current) reports."""
    await run_time_vs(update, context, start_hour=14, end_hour=17, command_label="2PM vs 5PM")



@user_guard
async def cmd_groups(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """List all registered groups."""
    await delete_group_command(update, context)
    groups = load_registered_groups()
    if not groups:
        await private_or_current_reply(
            update,
            context,
            "No registered groups.\n"
            "Use /register inside a group to add it, or set forward_groups in config.json."
        )
        return
    lines = ["Registered groups:"]
    for g in groups:
        lines.append(f"  • {g.get('title', '')} — {g['chat_id']}")
    await private_or_current_reply(update, context, "\n".join(lines))


def _configured_export_branches(cfg):
    branches = []

    def add_branch(value):
        value = str(value or "").strip().upper()
        if value and value not in branches:
            branches.append(value)

    for raw in cfg.get("api", {}).get("branch_code", "").split(","):
        add_branch(raw)

    for raw_list in cfg.get("zone_branches", {}).values():
        for raw in str(raw_list or "").split(","):
            add_branch(raw)

    return branches


def _parse_export_branches(raw_args, cfg):
    branches = []
    for arg in raw_args:
        for piece in str(arg).replace(";", ",").split(","):
            piece = piece.strip().upper()
            if piece and piece not in branches:
                branches.append(piece)
    return branches or _configured_export_branches(cfg)


def _classify_facility(code, type_label):
    code = str(code or "").strip().upper()
    type_label = str(type_label or "").strip()
    
    m = re.search(r"^[A-Z]{3}([PSA])\d+$", code)
    if m:
        letter = m.group(1)
        if letter == "P": return "Post Office"
        if letter == "S": return "Showroom"
        if letter == "A": return "Agent"

    if re.search(r"[A-Z]{3}P\d+", code): return "Post Office"
    if re.search(r"[A-Z]{3}S\d+", code): return "Showroom"
    if re.search(r"[A-Z]{3}A\d+", code): return "Agent"

    if "Warehouse" in type_label or "Hub" in type_label or "Operations" in type_label:
        return "Warehouse / Hub"

    if "Authorized" in type_label or "Agent" in type_label or "Dealer" in type_label:
        return "Agent"

    if "Post office" in type_label:
        return "Post Office"

    return "Post Office"


def _strip_department_code(value, code=""):
    text = str(value or "").strip()
    if not text or text.lower() == "nan":
        return ""
    if code:
        text = re.sub(rf"^{re.escape(str(code).strip())}\s*-\s*", "", text, flags=re.IGNORECASE)
    return re.sub(r"^[A-Z0-9]{3,10}\s*-\s*", "", text).strip()


def _clean_export_phone(value):
    phone = str(value or "").strip()
    if not phone or phone.lower() == "nan":
        return ""
    phone = re.sub(r"[^\d+]", "", phone)
    if phone.startswith("+855"):
        phone = "0" + phone[4:]
    if phone.isdigit() and not phone.startswith("0"):
        phone = "0" + phone
    return phone


async def fetch_lat_long(code: str, token: str, sem: asyncio.Semaphore):
    async with sem:
        url = f"https://gw-express.metfone.com.kh/vtp-user/api/v1/departments/{code}"
        headers = {
            "Authorization": f"Bearer {token}",
            "Referer": "https://opsexpress.metfone.com.kh/",
            "Accept": "application/json, text/plain, */*",
            "User-Agent": "Mozilla/5.0",
        }
        for attempt in range(3):
            try:
                import requests
                r = await asyncio.to_thread(
                    requests.get, url, headers=headers, timeout=10
                )
                if r.status_code == 200:
                    data = r.json()
                    addr = data.get("departmentAddress") or {}
                    return code, addr.get("latitude"), addr.get("longitude"), addr.get("address")
                elif r.status_code == 404:
                    return code, None, None, None
            except Exception as e:
                log.warning(f"Error fetching coordinates for {code} (attempt {attempt+1}): {e}")
                await asyncio.sleep(0.5)
        return code, None, None, None


def _post_office_export_row(item, fallback_branch=""):
    branch = item.get("branch") if isinstance(item.get("branch"), dict) else {}
    code = str(item.get("code", "")).strip().upper()
    branch_code = str(
        item.get("parentDepartmentCode")
        or branch.get("code")
        or fallback_branch
        or ""
    ).strip().upper()

    commune_en = _strip_department_code(
        item.get("enUsName") or item.get("name") or item.get("viVnName"),
        code,
    )
    commune_khmer = _strip_department_code(
        item.get("kmKhmName") or item.get("enUsName") or item.get("name"),
        code,
    )
    branch_en = _strip_department_code(
        branch.get("enUsName") or branch.get("name") or item.get("branch_name"),
        branch_code,
    )
    branch_khmer = _strip_department_code(
        branch.get("kmKhmName") or branch.get("enUsName") or branch.get("name"),
        branch_code,
    )
    phone = _clean_export_phone(item.get("phone"))

    category = _classify_facility(code, item.get("typeLabel") or item.get("type"))

    search_parts = [
        code,
        commune_en,
        commune_khmer,
        phone,
        branch_code,
        branch_en,
        branch_khmer,
        category,
    ]

    status = str(item.get("statusLabel") or item.get("status") or "In effect").strip()
    branch_display = f"{branch_code} - {branch_en}" if branch_en and branch_en != branch_code else (branch_code or "")

    return {
        "Post code": code,
        "Post office name": commune_en,
        "Branch": branch_display,
        "Post office level": category,
        "Status": status,
        "Pickup Branch": code,
        "Commune EN": commune_en,
        "Branch Code": branch_code,
        "Category": category,
        "Search Text": " | ".join(part for part in search_parts if part),
        "Branch Detail Address": item.get("address", ""),
        "Latitude": item.get("latitude"),
        "Longitude": item.get("longitude"),
    }


def _safe_excel_label(value, fallback="Export", max_len=80):
    label = re.sub(r"[^A-Za-z0-9_-]+", "_", str(value or "").strip())
    label = re.sub(r"_+", "_", label).strip("_")
    return (label or fallback)[:max_len]


BRANCH_TO_PROVINCE_EN = {
    "PRE": "Prey Veng",
    "PNP": "Phnom Penh",
    "SVA": "Svay Rieng",
    "KAN": "Kandal",
    "KAM": "Kampot",
    "KOH": "Koh Kong",
    "SIH": "Preah Sihanouk",
    "SPE": "Kampong Speu",
    "TAK": "Takeo",
    "BAN": "Banteay Meanchey",
    "BAT": "Battambang",
    "CHH": "Kampong Chhnang",
    "PUR": "Pursat",
    "SIE": "Siem Reap",
    "PRH": "Preah Vihear",
    "ODD": "Oddar Meanchey",
    "THO": "Kampong Thom",
    "CHA": "Kampong Cham",
    "KRA": "Kratie",
    "TBK": "Tboung Khmum",
    "ROT": "Ratanak Kiri",
    "MON": "Mondul Kiri",
    "STU": "Stung Treng",
    "KEP": "Kep",
    "PAI": "Pailin"
}

PROVINCE_MAP_KH = {
    "PRE": "ព្រៃវែង",
    "PNP": "ភ្នំពេញ",
    "SVA": "ស្វាយរៀង",
    "KAN": "កណ្តាល",
    "KAM": "កំពត",
    "KOH": "កោះកុង",
    "SIH": "ព្រះសីហនុ",
    "SPE": "កំពង់ស្ពឺ",
    "TAK": "តាកែវ",
    "BAN": "បន្ទាយមានជ័យ",
    "BAT": "បាត់ដំបង",
    "CHH": "កំពង់ឆ្នាំង",
    "PUR": "ពោធិ៍សាត់",
    "SIE": "សៀមរាប",
    "PRH": "ព្រះវិហារ",
    "ODD": "ឧត្តរមានជ័យ",
    "THO": "កំពង់ធំ",
    "CHA": "កំពង់ចាម",
    "KRA": "ក្រចេះ",
    "TBK": "ត្បូងឃ្មុំ",
    "ROT": "រតនគិរី",
    "MON": "មណ្ឌលគិរី",
    "STU": "ស្ទឹងត្រែង",
    "KEP": "កែប",
    "PAI": "ប៉ៃលិន"
}

DISTRICT_FALLBACK_KH = {
    "PRE": "ព្រៃវែង",
    "PNP": "ចំការមន",
    "SVA": "ស្វាយរៀង",
    "KAN": "តាខ្មៅ",
    "KAM": "កំពត",
    "KOH": "ខេមរភូមិន្ទ",
    "SIH": "ព្រះសីហនុ",
    "SPE": "ច្បារមន",
    "TAK": "ដូនកែវ",
    "BAN": "សិរីសោភ័ណ",
    "BAT": "បាត់ដំបង",
    "CHH": "កំពង់ឆ្នាំង",
    "PUR": "ពោធិ៍សាត់",
    "SIE": "សៀមរាប",
    "PRH": "ព្រះវិហារ",
    "ODD": "សំរោង",
    "THO": "ស្ទឹងសែន",
    "CHA": "កំពង់ចាម",
    "KRA": "ក្រចេះ",
    "TBK": "សួង",
    "ROT": "បានលុង",
    "MON": "សែនមនោរម្យ",
    "STU": "ស្ទឹងត្រែង",
}

DISTRICT_FALLBACK_EN = {
    "PRE": "Prey Veng",
    "PNP": "Chamkar Mon",
    "SVA": "Svay Rieng",
    "KAN": "Ta Khmau",
    "KAM": "Kampot",
    "KOH": "Khemarak Phoumin",
    "SIH": "Preah Sihanouk",
    "SPE": "Chbar Mon",
    "TAK": "Doun Kaev",
    "BAN": "Serei Saophoan",
    "BAT": "Battambang",
    "CHH": "Kampong Chhnang",
    "PUR": "Pursat",
    "SIE": "Siem Reap",
    "PRH": "Preah Vihear",
    "ODD": "Samraong",
    "THO": "Steung Saen",
    "CHA": "Kampong Cham",
    "KRA": "Kratie",
    "TBK": "Suong",
    "ROT": "Banlung",
    "MON": "Senmonorom",
    "STU": "Stung Treng",
}

_gazetteer_data = None

def _get_gazetteer():
    global _gazetteer_data
    if _gazetteer_data is not None:
        return _gazetteer_data
    path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "cambodia_gazetteer.json")
    if os.path.exists(path):
        try:
            with open(path, "r", encoding="utf-8") as f:
                _gazetteer_data = json.load(f)
        except Exception:
            _gazetteer_data = []
    else:
        _gazetteer_data = []
    return _gazetteer_data

def _map_to_administrative_division(branch_code, commune_en_raw, commune_kh_raw):
    gazetteer = _get_gazetteer()
    prov_target = BRANCH_TO_PROVINCE_EN.get(branch_code, "")
    
    target_en = str(commune_en_raw).lower().replace(" ", "").replace("-", "")
    target_kh = str(commune_kh_raw).lower().replace(" ", "").replace("-", "")
    
    for item in gazetteer:
        if prov_target and item["prov_en"].lower().replace(" ", "") != prov_target.lower().replace(" ", ""):
            continue
        
        c_en = item["comm_en"].lower().replace(" ", "").replace("-", "")
        c_kh = item["comm_kh"].lower().replace(" ", "").replace("-", "")
        
        if c_en == target_en or c_kh == target_kh or c_en in target_en or target_en in c_en:
            return item["prov_en"], item["prov_kh"], item["dist_en"], item["dist_kh"], item["comm_kh"]
            
    prov_en = BRANCH_TO_PROVINCE_EN.get(branch_code, "Battambang")
    prov_kh = PROVINCE_MAP_KH.get(branch_code, "បាត់ដំបង")
    dist_en = DISTRICT_FALLBACK_EN.get(branch_code, "Battambang")
    dist_kh = DISTRICT_FALLBACK_KH.get(branch_code, "បាត់ដំបង")
    comm_kh = commune_kh_raw or commune_en_raw
    return prov_en, prov_kh, dist_en, dist_kh, comm_kh


def _write_post_office_export_excel(df, out_path, sheet_label, title):
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
    import pandas as pd

    PRIMARY = "00A651"      # Metfone Green
    SECONDARY = "EAF7EF"    # Light Green
    WHITE = "FFFFFF"
    DARK_TEXT = "333333"
    BORDER_CLR = "B2D8B2"

    thin_border = Border(
        left=Side(style='thin', color=BORDER_CLR),
        right=Side(style='thin', color=BORDER_CLR),
        top=Side(style='thin', color=BORDER_CLR),
        bottom=Side(style='thin', color=BORDER_CLR),
    )
    header_font = Font(name='Calibri', bold=True, color=WHITE, size=11)
    header_fill = PatternFill(start_color=PRIMARY, end_color=PRIMARY, fill_type='solid')
    header_align = Alignment(horizontal='center', vertical='center', wrap_text=True)
    alt_fill = PatternFill(start_color=SECONDARY, end_color=SECONDARY, fill_type='solid')
    data_font = Font(name='Calibri', color=DARK_TEXT, size=10)
    title_font = Font(name='Calibri', bold=True, color=PRIMARY, size=14)

    wb = Workbook()
    ws = wb.active
    ws.title = "Post Offices"
    ws.views.sheetView[0].showGridLines = True

    export_headers = ["Post code", "Post office name", "Branch", "Post office level", "Status"]

    # Title row (Row 1)
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(export_headers))
    title_cell = ws.cell(row=1, column=1, value=title)
    title_cell.font = title_font
    title_cell.alignment = Alignment(horizontal='left', vertical='center')
    ws.row_dimensions[1].height = 30

    # Header row (Row 2)
    for col_idx, col_name in enumerate(export_headers, 1):
        cell = ws.cell(row=2, column=col_idx, value=col_name)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = header_align
        cell.border = thin_border
    ws.row_dimensions[2].height = 28

    # Data rows (Row 3 onwards)
    for idx, row in enumerate(df.itertuples(index=False), 3):
        is_alt = (idx % 2 == 1)
        r_dict = dict(zip(df.columns, row))

        post_code = str(r_dict.get('Post code') or r_dict.get('Pickup Branch') or r_dict.get('Department Code') or r_dict.get('code') or '').strip().upper()
        po_name = str(r_dict.get('Post office name') or r_dict.get('Commune EN') or r_dict.get('Department Name') or r_dict.get('name') or '').strip()
        b_code = str(r_dict.get('Branch Code') or r_dict.get('parentDepartmentCode') or '').strip().upper()
        b_name = str(r_dict.get('Branch EN') or r_dict.get('branch_name') or '').strip()
        branch_str = str(r_dict.get('Branch') or (f"{b_code} - {b_name}" if b_name and b_name != b_code else b_code)).strip()
        po_level = str(r_dict.get('Post office level') or r_dict.get('Category') or r_dict.get('Type') or '').strip()
        status = str(r_dict.get('Status') or 'In effect').strip()

        row_vals = [post_code, po_name, branch_str, po_level, status]

        for col_idx, value in enumerate(row_vals, 1):
            cell = ws.cell(row=idx, column=col_idx, value=value)
            cell.font = data_font
            cell.alignment = Alignment(horizontal='center' if col_idx in (1, 4, 5) else 'left', vertical='center')
            cell.border = thin_border
            if is_alt:
                cell.fill = alt_fill

    ws.freeze_panes = "A3"
    ws.auto_filter.ref = f"A2:E{len(df)+2}"

    # Auto-fit column widths
    for col_idx, col_name in enumerate(export_headers, 1):
        max_len = len(str(col_name))
        col_letter = get_column_letter(col_idx)
        for r_i in range(3, min(len(df) + 3, 200)):
            v = ws.cell(row=r_i, column=col_idx).value
            if v:
                max_len = max(max_len, len(str(v)))
        ws.column_dimensions[col_letter].width = max(max_len + 4, 18)

    wb.save(out_path)


def _write_post_office_export_excel_v2(df, out_path, sheet_label, title):
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter

    PRIMARY = "00A651"      # Metfone Green
    SECONDARY = "EAF7EF"    # Light Green
    WHITE = "FFFFFF"
    DARK_TEXT = "333333"
    BORDER_CLR = "B2D8B2"

    thin_border = Border(
        left=Side(style='thin', color=BORDER_CLR),
        right=Side(style='thin', color=BORDER_CLR),
        top=Side(style='thin', color=BORDER_CLR),
        bottom=Side(style='thin', color=BORDER_CLR),
    )
    header_font = Font(name='Calibri', bold=True, color=WHITE, size=11)
    header_fill = PatternFill(start_color=PRIMARY, end_color=PRIMARY, fill_type='solid')
    header_align = Alignment(horizontal='center', vertical='center', wrap_text=True)
    alt_fill = PatternFill(start_color=SECONDARY, end_color=SECONDARY, fill_type='solid')
    data_font = Font(name='Calibri', color=DARK_TEXT, size=10)
    title_font = Font(name='Calibri', bold=True, color=PRIMARY, size=14)

    wb = Workbook()
    ws = wb.active
    ws.title = "Post Offices"
    ws.views.sheetView[0].showGridLines = True

    export_headers = ["Post code", "Post office name", "Branch Detail Address", "Post office level", "Status", "Latitude", "Longitude"]

    # Title row (Row 1)
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(export_headers))
    title_cell = ws.cell(row=1, column=1, value=title)
    title_cell.font = title_font
    title_cell.alignment = Alignment(horizontal='left', vertical='center')
    ws.row_dimensions[1].height = 30

    # Header row (Row 2)
    for col_idx, col_name in enumerate(export_headers, 1):
        cell = ws.cell(row=2, column=col_idx, value=col_name)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = header_align
        cell.border = thin_border
    ws.row_dimensions[2].height = 28

    # Data rows (Row 3 onwards)
    for idx, row in enumerate(df.itertuples(index=False), 3):
        is_alt = (idx % 2 == 1)
        r_dict = dict(zip(df.columns, row))

        post_code = str(r_dict.get('Post code') or r_dict.get('Pickup Branch') or r_dict.get('Department Code') or r_dict.get('code') or '').strip().upper()
        po_name = str(r_dict.get('Post office name') or r_dict.get('Commune EN') or r_dict.get('Department Name') or r_dict.get('name') or '').strip()
        address = str(r_dict.get('Branch Detail Address') or r_dict.get('address') or '').strip()
        po_level = str(r_dict.get('Post office level') or r_dict.get('Category') or r_dict.get('Type') or '').strip()
        status = str(r_dict.get('Status') or 'In effect').strip()
        lat = r_dict.get('Latitude')
        lon = r_dict.get('Longitude')

        row_vals = [post_code, po_name, address, po_level, status, lat, lon]

        for col_idx, value in enumerate(row_vals, 1):
            cell = ws.cell(row=idx, column=col_idx, value=value)
            cell.font = data_font
            if col_idx in (1, 4, 5, 6, 7):
                cell.alignment = Alignment(horizontal='center', vertical='center')
            else:
                cell.alignment = Alignment(horizontal='left', vertical='center')
            cell.border = thin_border
            if is_alt:
                cell.fill = alt_fill

    ws.freeze_panes = "A3"
    ws.auto_filter.ref = f"A2:{get_column_letter(len(export_headers))}{len(df)+2}"

    # Auto-fit column widths
    for col_idx, col_name in enumerate(export_headers, 1):
        max_len = len(str(col_name))
        col_letter = get_column_letter(col_idx)
        for r_i in range(3, min(len(df) + 3, 200)):
            v = ws.cell(row=r_i, column=col_idx).value
            if v:
                max_len = max(max_len, len(str(v)))
        ws.column_dimensions[col_letter].width = max(max_len + 4, 15)

    wb.save(out_path)


def _write_post_office_export_excel_by_category(df, out_path, sheet_label, title):
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
    import pandas as pd

    DARK_BLUE = "172033"
    PRIMARY = "00A651"
    WHITE = "FFFFFF"
    BORDER_CLR = "CCCCCC"
    
    thin_border = Border(
        left=Side(style="thin", color=BORDER_CLR),
        right=Side(style="thin", color=BORDER_CLR),
        top=Side(style="thin", color=BORDER_CLR),
        bottom=Side(style="thin", color=BORDER_CLR),
    )
    
    lookup_map = {}
    try:
        ref_df = pd.read_csv("pickup_branch_lookup.csv", dtype=str)
        for _, r in ref_df.iterrows():
            code_val = str(r.get("Pickup Branch", "")).strip().upper()
            comm_val = str(r.get("Commune EN", "")).strip()
            if code_val and comm_val:
                lookup_map[code_val] = comm_val
    except Exception as e:
        log.warning("Could not load pickup_branch_lookup.csv: %s", e)

    gazetteer = _get_gazetteer()
    valid_locations = set()
    for item in gazetteer:
        c_en = str(item.get("comm_en", "")).lower().replace(" ", "").replace("-", "")
        d_en = str(item.get("dist_en", "")).lower().replace(" ", "").replace("-", "")
        p_en = str(item.get("prov_en", "")).lower().replace(" ", "").replace("-", "")
        if c_en: valid_locations.add(c_en)
        if d_en: valid_locations.add(d_en)
        if p_en: valid_locations.add(p_en)

    wb = Workbook()
    # Remove default sheet
    wb.remove(wb.active)

    category_tabs = [
        ("Post Offices", "Post Office", "137333", "E6F4EA"),
        ("Showrooms", "Showroom", "B06000", "FEF7E0"),
        ("Agents", "Agent", "1A73E8", "E8F0FE"),
    ]

    for tab_title, cat_name, fg_color, bg_color in category_tabs:
        sub_df = df[df["Category"] == cat_name] if "Category" in df.columns else df
        ws = wb.create_sheet(title=tab_title)
        ws.views.sheetView[0].showGridLines = True

        data_headers = ["Province *", "District *", "District KH", "Delivery Store *", "Category *", "Phone Number", "Latitude", "Longitude", "Suggest Edit"]
        
        for col_idx, col_name in enumerate(data_headers, 1):
            cell = ws.cell(row=1, column=col_idx, value=col_name)
            cell.font = Font(name="Calibri", bold=True, color=WHITE, size=11)
            cell.fill = PatternFill(start_color=DARK_BLUE, end_color=DARK_BLUE, fill_type="solid")
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border = thin_border
        ws.row_dimensions[1].height = 28

        for idx, row in sub_df.reset_index(drop=True).iterrows():
            branch_code = str(row.get("Branch Code", "")).strip().upper()
            commune_en = str(row.get("Commune EN", ""))
            commune_kh = str(row.get("Commune Khmer", ""))
            code = str(row.get("Pickup Branch", ""))
            phone = str(row.get("Phone", ""))
            lat = row.get("Latitude")
            lon = row.get("Longitude")
            category = str(row.get("Category") or _classify_facility(code, row.get("Type"))).strip()
            
            prov_en, prov_kh, dist_en, dist_kh, comm_kh = _map_to_administrative_division(branch_code, commune_en, commune_kh)
            store_name = f"{code} - {commune_en}"
            
            suggest_val = ""
            if commune_en:
                comm_norm = str(commune_en).lower().replace(" ", "").replace("-", "")
                correct_name = lookup_map.get(code)
                if correct_name:
                    correct_norm = str(correct_name).lower().replace(" ", "").replace("-", "")
                    if comm_norm != correct_norm:
                        suggest_val = f"Change to \"{correct_name}\""
                else:
                    if comm_norm not in valid_locations:
                        suggest_val = "Verify Location Name"

            row_idx = idx + 2
            ws.cell(row=row_idx, column=1, value=prov_en).font = Font(name="Calibri", size=10)
            ws.cell(row=row_idx, column=2, value=dist_en).font = Font(name="Calibri", size=10)
            ws.cell(row=row_idx, column=3, value=dist_kh).font = Font(name="Calibri", size=10)
            ws.cell(row=row_idx, column=4, value=store_name).font = Font(name="Calibri", size=10)
            
            cat_cell = ws.cell(row=row_idx, column=5, value=category)
            cat_cell.alignment = Alignment(horizontal="center", vertical="center")
            cat_cell.font = Font(name="Calibri", size=10, bold=True, color=fg_color)
            cat_cell.fill = PatternFill(start_color=bg_color, end_color=bg_color, fill_type="solid")

            ws.cell(row=row_idx, column=6, value=phone).font = Font(name="Calibri", size=10)
            ws.cell(row=row_idx, column=7, value=lat).font = Font(name="Calibri", size=10)
            ws.cell(row=row_idx, column=8, value=lon).font = Font(name="Calibri", size=10)
            
            cell_suggest = ws.cell(row=row_idx, column=9, value=suggest_val)
            if suggest_val:
                cell_suggest.font = Font(name="Calibri", size=10, bold=True, color="C00000")
                cell_suggest.fill = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
            else:
                cell_suggest.font = Font(name="Calibri", size=10)
                
            for col_idx in range(1, 10):
                ws.cell(row=row_idx, column=col_idx).border = thin_border

        ws.column_dimensions["A"].width = 25
        ws.column_dimensions["B"].width = 25
        ws.column_dimensions["C"].width = 25
        ws.column_dimensions["D"].width = 45
        ws.column_dimensions["E"].width = 22
        ws.column_dimensions["F"].width = 20
        ws.column_dimensions["G"].width = 15
        ws.column_dimensions["H"].width = 15
        ws.column_dimensions["I"].width = 30
        
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = f"A1:I{len(sub_df)+1}"

    # ── Sheet 4: All Details ──
    ws2 = wb.create_sheet(title="All Details")
    ws2.views.sheetView[0].showGridLines = True
    
    detail_headers = [
        "Department Code", "Department Name", "Commune Khmer",
        "Branch Code", "Branch EN", "Branch Khmer",
        "Type", "Category *", "Status", "Phone Number", "Latitude", "Longitude"
    ]
    
    for col_idx, col_name in enumerate(detail_headers, 1):
        cell = ws2.cell(row=1, column=col_idx, value=col_name)
        cell.font = Font(name="Calibri", bold=True, color=WHITE, size=11)
        cell.fill = PatternFill(start_color=PRIMARY, end_color=PRIMARY, fill_type="solid")
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = thin_border
    ws2.row_dimensions[1].height = 28
    
    detail_cols = [
        "Pickup Branch", "Commune EN", "Commune Khmer",
        "Branch Code", "Branch EN", "Branch Khmer",
        "Type", "Category", "Status", "Phone", "Latitude", "Longitude"
    ]
    
    for idx, row in df.iterrows():
        row_idx = idx + 2
        for col_idx, col_key in enumerate(detail_cols, 1):
            val = row.get(col_key, "")
            cell = ws2.cell(row=row_idx, column=col_idx, value=val)
            cell.font = Font(name="Calibri", size=10)
            cell.border = thin_border
            
    ws2.freeze_panes = "A2"
    ws2.auto_filter.ref = f"A1:{get_column_letter(len(detail_headers))}{len(df)+1}"
    
    for col_idx, col_name in enumerate(detail_headers, 1):
        max_len = len(str(col_name))
        col_letter = get_column_letter(col_idx)
        for row_idx in range(2, min(len(df) + 2, 50)):
            v = ws2.cell(row=row_idx, column=col_idx).value
            if v:
                max_len = max(max_len, len(str(v)))
        ws2.column_dimensions[col_letter].width = min(max_len + 4, 45)

    wb.save(out_path)


async def send_pickup_branch_export(update, context, cfg, raw_args, all2_mode=False):
    first_arg = raw_args[0].lower() if raw_args else ""
    cat_mode = first_arg in ("po", "postoffice", "postoffices", "post_office", "office", "showroom", "showrooms", "agent", "agents", "dealer", "type", "types", "category", "categories", "divide", "split")
    all_mode = cat_mode or (first_arg in ("all", "pickup", "pickups", "search", "branches", "all2", "pickup2", "pickups2", "branches2"))
    branch_args = raw_args[1:] if (all_mode or cat_mode) else raw_args
    branch_codes = _parse_export_branches(branch_args, cfg)
    if not branch_codes:
        await private_or_current_reply(update, context, "No branch codes configured for export.")
        return

    label = "ALL_PICKUP" if all_mode and not branch_args else ",".join(branch_codes)
    description = "all pickup branches" if all_mode and not branch_args else ", ".join(branch_codes)
    msg = await send_requester_text(update, context, f"Fetching post offices for {description} in parallel...")

    try:
        import pandas as pd

        post_offices = []
        branch_errors = []

        # If "all" mode with no specific branches, fetch ALL post offices at once
        if all_mode and not branch_args:
            try:
                post_offices = await asyncio.to_thread(
                    downloader.download_all_post_offices, cfg["api"]
                )
                for item in post_offices:
                    if isinstance(item, dict):
                        item["_export_branch_query"] = str(item.get("branchCode", ""))
            except Exception as e:
                log.warning("download_all_post_offices failed, falling back: %s", e)
                post_offices = []

        # Fallback or specific branch mode: fetch per branch
        if not post_offices:
            sem = asyncio.Semaphore(4)
        
            async def sem_download(code):
                async with sem:
                    await asyncio.sleep(0.3)
                    return await asyncio.to_thread(
                        downloader.download_post_offices,
                        cfg["api"],
                        code,
                    )
        
            tasks = [sem_download(bc) for bc in branch_codes]
            results = await asyncio.gather(*tasks, return_exceptions=True)
        
            for branch_code, res in zip(branch_codes, results):
                if isinstance(res, Exception):
                    branch_errors.append(f"{branch_code}: {res}")
                    log.warning("Export failed for branch %s: %s", branch_code, res)
                else:
                    for item in res:
                        if isinstance(item, dict):
                            item = dict(item)
                            item["_export_branch_query"] = branch_code
                        post_offices.append(item)

        if not post_offices:
            extra = "\n".join(branch_errors[:5])
            await edit_or_send_requester_text(
                msg,
                update,
                context,
                f"No post offices found for {description}." + (f"\n{extra}" if extra else "")
            )
            return

        # Fetch coordinates in parallel
        unique_codes = list(set(
            str(item.get("code", "")).strip().upper()
            for item in post_offices
            if isinstance(item, dict) and item.get("code")
        ))
        
        if unique_codes:
            await edit_or_send_requester_text(
                msg,
                update,
                context,
                f"Fetched {len(post_offices)} offices. Retrieving coordinates..."
            )
            detail_sem = asyncio.Semaphore(15)
            detail_tasks = [fetch_lat_long(code, cfg["api"]["bearer_token"], detail_sem) for code in unique_codes]
            detail_results = await asyncio.gather(*detail_tasks, return_exceptions=True)
            
            coords_map = {}
            for res in detail_results:
                if isinstance(res, tuple):
                    if len(res) == 4:
                        code, lat, lon, addr_str = res
                        coords_map[code] = (lat, lon, addr_str)
                    elif len(res) == 3:
                        code, lat, lon = res
                        coords_map[code] = (lat, lon, "")
            
            for item in post_offices:
                if isinstance(item, dict):
                    code = str(item.get("code", "")).strip().upper()
                    lat, lon, addr_str = coords_map.get(code, (None, None, ""))
                    item["latitude"] = lat
                    item["longitude"] = lon
                    item["address"] = addr_str

        rows = [
            _post_office_export_row(item, item.get("_export_branch_query", ""))
            for item in post_offices
            if isinstance(item, dict)
        ]
        df = pd.DataFrame(rows)
        if "Pickup Branch" in df.columns:
            df = df[df["Pickup Branch"].astype(str).str.strip() != ""].copy()
            df = df.drop_duplicates(subset=["Pickup Branch"], keep="first")
        sort_cols = [c for c in ("Branch Code", "Pickup Branch") if c in df.columns]
        if sort_cols:
            df = df.sort_values(sort_cols).reset_index(drop=True)

        if df.empty:
            await edit_or_send_requester_text(
                msg,
                update,
                context,
                f"No valid post office rows found for {description}."
            )
            return

        if all_mode and not cat_mode:
            df.to_csv(PICKUP_BRANCH_LOOKUP_PATH, index=False, encoding="utf-8-sig")

        # Category filtering & divide options
        cat_filter = None
        divide_mode = False
        if first_arg in ("po", "postoffice", "postoffices", "post_office", "office"):
            cat_filter = "Post Office"
        elif first_arg in ("showroom", "showrooms", "sub", "suboffice"):
            cat_filter = "Showroom"
        elif first_arg in ("agent", "agents", "dealer", "dealers"):
            cat_filter = "Agent"
        elif first_arg in ("type", "types", "category", "categories", "divide", "split"):
            divide_mode = True

        if cat_filter and "Category" in df.columns:
            df = df[df["Category"] == cat_filter].copy()
            description += f" ({cat_filter}s only)"

        await edit_or_send_requester_text(
            msg,
            update,
            context,
            f"Found {len(df)} locations. Building modern Excel..."
        )

        tmpdir = tempfile.mkdtemp(prefix="export_po_")
        stamp = datetime.now().strftime("%d.%m_%HH%M")
        safe_label = _safe_excel_label(label)
        filename = f"PickupBranches2_{safe_label}_{stamp}.xlsx" if all2_mode else f"PickupBranches_{safe_label}_{stamp}.xlsx"
        out_path = os.path.join(tmpdir, filename)
        title = f"Pickup Branch Search Export v2 - {description} ({len(df)} locations)" if all2_mode else f"Pickup Branch Search Export - {description} ({len(df)} locations)"

        if divide_mode:
            _write_post_office_export_excel_by_category(df, out_path, label, title)
        elif all2_mode:
            _write_post_office_export_excel_v2(df, out_path, label, title)
        else:
            _write_post_office_export_excel(df, out_path, label, title)

        with open(out_path, "rb") as f:
            await send_requester_document(update, context, f, filename)

        done_lines = [
            f"Exported {len(df)} locations for {description}.",
            f"Time: {datetime.now().strftime('%d.%m.%Y %H:%M')}",
        ]
        if all_mode:
            done_lines.append(f"Updated local lookup: {os.path.basename(PICKUP_BRANCH_LOOKUP_PATH)}")
        if branch_errors:
            done_lines.append("Some branches failed: " + "; ".join(branch_errors[:3]))
            if len(branch_errors) > 3:
                done_lines.append(f"...and {len(branch_errors) - 3} more.")

        await edit_or_send_requester_text(msg, update, context, "\n".join(done_lines))

    except Exception as e:
        log.exception("Error in pickup branch export")
        await edit_or_send_requester_text(msg, update, context, f"Export failed: {e}")


@pm_required_handler
async def cmd_export(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """/export <branch_code> — export post office list for a branch.
    Examples: /export KAM  |  /export PNP  |  /export SVA  |  /export all
    """
    await delete_group_command(update, context)
    cfg = load_config()

    raw_args = [a.strip() for a in (context.args or []) if a.strip()]
    if not raw_args:
        await private_or_current_reply(
            update,
            context,
            "Usage: /export <branch_code>\n"
            "Downloads the post office list for a branch.\n\n"
            "Examples:\n"
            "  /export KAM — all Kampot post offices\n"
            "  /export PNP — all Phnom Penh post offices\n"
            "  /export SVA — all Svay Rieng post offices\n"
            "  /export all — export all configured branches"
        )
        return

    first_arg = raw_args[0].lower()
    if first_arg in ("all2", "pickup2", "pickups2", "branches2"):
        await send_pickup_branch_export(update, context, cfg, raw_args, all2_mode=True)
        return
    if first_arg in ("all", "pickup", "pickups", "search", "branches"):
        await send_pickup_branch_export(update, context, cfg, raw_args, all2_mode=False)
        return

    args = [a.strip().upper() for a in raw_args]
    branch_code = args[0]

    msg = await send_requester_text(
        update, context,
        f"📥 Fetching post offices for branch: {branch_code}..."
    )

    try:
        import pandas as pd

        post_offices = downloader.download_post_offices(cfg["api"], branch_code)

        if not post_offices:
            await edit_or_send_requester_text(
                msg, update, context,
                f"No post offices found for branch {branch_code}."
            )
            return

        rows = [
            _post_office_export_row(item, branch_code)
            for item in (post_offices if isinstance(post_offices, list) else [])
            if isinstance(item, dict)
        ]
        if not rows and isinstance(post_offices, list) and post_offices:
            df_raw = pd.DataFrame(post_offices)
            df_raw['Post code'] = df_raw.get('code', '')
            df_raw['Post office name'] = df_raw.get('name', '')
            df_raw['Branch'] = branch_code
            df_raw['Post office level'] = df_raw.get('typeLabel', df_raw.get('type', 'Post Office'))
            df_raw['Status'] = df_raw.get('statusLabel', df_raw.get('status', 'In effect'))
            df = df_raw[['Post code', 'Post office name', 'Branch', 'Post office level', 'Status']].copy()
        else:
            df = pd.DataFrame(rows)

        if df.empty:
            await edit_or_send_requester_text(
                msg, update, context,
                f"No post office rows found for branch {branch_code}."
            )
            return

        await edit_or_send_requester_text(
            msg, update, context,
            f"Found {len(df)} post offices. Building Excel..."
        )

        tmpdir = tempfile.mkdtemp(prefix="export_po_")
        stamp = datetime.now().strftime("%d.%m_%HH%M")
        filename = f"PostOffices_{branch_code}_{stamp}.xlsx"
        out_path = os.path.join(tmpdir, filename)
        title = f"📋 Post Office List — {branch_code} ({len(df)} offices)"

        _write_post_office_export_excel(df, out_path, branch_code, title)

        with open(out_path, "rb") as f:
            await send_requester_document(
                update, context, f, filename,
            )

        await edit_or_send_requester_text(
            msg, update, context,
            f"✅ Exported {len(df)} post offices for {branch_code} — "
            f"{datetime.now().strftime('%d.%m.%Y %H:%M')}"
        )

    except Exception as e:
        log.exception("Error in /export")
        await edit_or_send_requester_text(
            msg, update, context,
            f"❌ Export failed: {e}"
        )


@pm_required_handler
async def cmd_export_all2(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """/exportall2 — export post office list with Detail Address, Lat, and Lon."""
    await delete_group_command(update, context)
    cfg = load_config()
    raw_args = ["all2"] + [a.strip() for a in (context.args or []) if a.strip()]
    await send_pickup_branch_export(update, context, cfg, raw_args, all2_mode=True)


@pm_required_handler
async def cmd_find(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """/find <query> — search for orders by phone number or order ID."""
    await delete_group_command(update, context)
    args = [a.strip() for a in (context.args or []) if a.strip()]
    if not args:
        await private_or_current_reply(
            update,
            context,
            "Usage: `/find <phone_number>` or `/find <order_id>`\n"
            "Example: `/find 0712345678`",
        )
        return

    force_refresh = False
    if "force" in [a.lower() for a in args]:
        force_refresh = True
        args = [a for a in args if a.lower() != "force"]

    if not args:
        await private_or_current_reply(
            update,
            context,
            "Please specify a phone number or order ID to find.",
        )
        return

    query = args[0].lower()
    msg = await send_requester_text(
        update, context,
        f"🔍 Searching for '{query}' in recent orders (up to 7 days)..."
    )

    cfg = load_config()
    tmpdir = tempfile.mkdtemp(prefix="find_")
    stamp = datetime.now().strftime("%d.%m_%HH%M")
    src = os.path.join(tmpdir, f"export_{stamp}.xlsx")

    try:
        import downloader
        downloader.download_detail(cfg["api"], src, force_refresh=force_refresh)
        
        import pandas as pd
        xl = pd.ExcelFile(src)
        sheet = xl.sheet_names[0]
        if len(xl.sheet_names) > 1:
            for sname in xl.sheet_names:
                try:
                    s = xl.parse(sname, nrows=3)
                    if 'ORDER ID' in s.columns:
                        sheet = sname
                        break
                except Exception:
                    continue
        df = xl.parse(sheet)

        search_cols = ['ORDER ID', 'SENDER', 'RECEIVER', 'Cus name', 'Phone']
        match_idx = pd.Series(False, index=df.index)

        for col in search_cols:
            if col in df.columns:
                match_idx = match_idx | df[col].astype(str).str.lower().str.contains(query, na=False)

        results = df[match_idx]

        if results.empty:
            await edit_or_send_requester_text(
                msg, update, context,
                f"❌ No orders found matching '{query}'."
            )
            return

        out_lines = [f"✅ Found {len(results)} order(s) for '{query}':\n"]
        # Limit to first 20 to avoid blowing up Telegram message length too much
        limit = 20
        for _, row in results.head(limit).iterrows():
            order_id = row.get('ORDER ID', 'Unknown')
            status = row.get('CURRENT STATUS', 'Unknown')
            po = row.get('CURRENT POST OFFICE', 'Unknown')
            date = row.get('CREATED DATE', '')
            if pd.isna(date): date = ''
            sender = row.get('SENDER', '')
            if pd.isna(sender): sender = ''

            out_lines.append(f"📦 *Order ID:* `{order_id}`")
            out_lines.append(f"📍 *Status:* {status}")
            out_lines.append(f"🏢 *Post Office:* {po}")
            if date: out_lines.append(f"📅 *Created:* {date}")
            if sender: out_lines.append(f"👤 *Sender:* {sender}")
            out_lines.append("━━━━━━━━━━━━━━━━━━━━━━")

        if len(results) > limit:
            out_lines.append(f"... and {len(results) - limit} more. Showing first {limit}.")

        final_text = "\n".join(out_lines)
        # Final safety for telegram limits
        if len(final_text) > 4000:
            final_text = final_text[:4000] + "\n... (Message truncated)"

        await edit_or_send_requester_text(msg, update, context, final_text)

    except Exception as e:
        log.exception("Error in /find")
        await edit_or_send_requester_text(
            msg, update, context,
            f"❌ Search failed: {e}"
        )

@pm_required_handler
async def cmd_ask(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """/ask <order_id> [kh] - find an order, identify the responsible post office to scan it, and describe the next step."""
    await delete_group_command(update, context)
    args = [a.strip() for a in (context.args or []) if a.strip()]
    if not args:
        await private_or_current_reply(
            update,
            context,
            "Usage: `/ask <order_id> [kh]`\n"
            "Example: `/ask 2900492262 kh`"
        )
        return

    use_khmer = False
    query = ""
    for arg in args:
        if arg.lower() in ("kh", "-kh"):
            use_khmer = True
        else:
            query = arg.lower()

    if not query:
        await private_or_current_reply(
            update,
            context,
            "Usage: `/find <phone_number | order_id | name | branch>`\n"
            "Example: `/find 0715834688` or `/find KANP001`"
        )
        return

    order_id = query
    loading_msg = f"🔍 កំពុងស្វែងរកព័ត៌មានអំពី '{query}'..." if use_khmer else f"🔍 Searching details for '{query}'..."
    msg = await send_requester_text(
        update, context,
        loading_msg
    )

    # 1. Search SQLite database index for matching sender/receiver phone numbers, names, or branch
    try:
        from search_engine import search_orders
        db_results = search_orders(query, limit=15)
        # If searching by phone number or branch or name (returns multiple results or non-single-order-id)
        if len(db_results) > 1 or (not query.isdigit() and len(query) < 10) or (len(query) in (9, 10, 11, 12) and not query.startswith("31") and not query.startswith("32") and not query.startswith("29")):
            if db_results:
                resp = f"🔍 *RESULTS FOR '{query}'* ({len(db_results)} matches):\n" + "━"*30 + "\n"
                for res in db_results[:10]:
                    vip_badge = " 🌟VIP" if res.get("vip") == "VIP" else ""
                    resp += (
                        f"📦 *Order*: `{res['order_id']}`{vip_badge}\n"
                        f"👤 *Sender*: {res['sender_name']} (`{res['sender_phone'] or 'N/A'}`)\n"
                        f"📥 *Receiver*: {res['receiver_name']} (`{res['receiver_phone'] or 'N/A'}`)\n"
                        f"📍 *Branches*: Rec `{res['receive_po']}` → Cur `{res['current_po']}` → Del `{res['delivery_po']}`\n"
                        f"📊 *Status*: `{res['current_status']}` (COD: `${res['cod']:.2f}`)\n"
                        + "─"*30 + "\n"
                    )
                if len(db_results) > 10:
                    resp += f"\n_Showing top 10 of {len(db_results)} matches._"
                await edit_or_send_requester_text(msg, update, context, resp, parse_mode="Markdown")
                return
    except Exception as e_db:
        log.warning("SQLite search engine error: %s", e_db)

    cfg = load_config()
    token = cfg["api"]["bearer_token"]

    try:
        import requests
        import pandas as pd
        
        headers = {
            "Authorization": f"Bearer {token}",
            "Accept": "application/json, text/plain, */*",
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
        }
        
        # 1. Always attempt to query order search for rich details first
        order_data = None
        search_url = "https://gw-express.metfone.com.kh/tms-receiving/api/v1/orders/search"
        try:
            r_search = requests.get(search_url, params={"order_code": query}, headers=headers, timeout=15)
            if r_search.status_code == 200:
                order_data = r_search.json()
        except Exception as e:
            log.warning("Failed to query orders/search: %s", e)

        # 2. Query tracking trips API
        url = "https://gw-express.metfone.com.kh/tms-tracking/api/v1/order-tracking"
        params = {"order_id": query}
        
        trips = []
        is_synthetic = False
        
        try:
            r = requests.get(url, params=params, headers=headers, timeout=15)
            if r.status_code == 401:
                await edit_or_send_requester_text(
                    msg, update, context,
                    "❌ API Token expired. Please update config.json with a valid Bearer token."
                )
                return
            elif r.status_code == 200:
                data = r.json()
                trips = data.get("trackingTrips", [])
        except Exception as e:
            log.warning("Failed to query order-tracking: %s", e)

        # 3. Fallback to synthetic trip if tracking trips empty (e.g. no scans yet or branch restriction)
        if not trips:
            if order_data:
                status_val = str(order_data.get("order_status", ""))
                is_synthetic = True
                
                # Map status code to status name (-99 = Metfone undefined/unscanned)
                status_name_map = {
                    "-99": "REGISTERED",
                    "0": "REGISTERED",
                    "100": "REGISTERED",
                    "110": "CONFIRMED",
                    "120": "PICKING",
                    "200": "RECEIVED",
                    "201": "CANCELLED",
                    "210": "TRANSITING",
                    "302": "TRANSITING",
                    "306": "TRANSITING",
                    "310": "TRANSITING",
                    "309": "TRANSITING",
                    "311": "TRANSITING",
                    "500": "TRANSITING",
                    "300": "TRANSITING",
                    "400": "Assigned to deliver - Not completed",
                    "401": "DELIVERING",
                    "402": "DELIVERING",
                    "420": "PTC_STORE",
                    "430": "PTC_RE_DELIVERY",
                    "460": "PTC_REDELIVERING",
                    "472": "STORED",
                    "480": "CHANGE_ADDRESS",
                    "405": "DELIVERED",
                    "410": "DELIVERED_COMPLETED",
                    "512": "RETURN_EXCEPTION",
                    "520": "RETURNED",
                    "540": "RETURN_FAILED"
                }
                status_name_str = status_name_map.get(status_val, "CREATED")
                
                # Format time nicely from epoch millis
                created_date_epoch = order_data.get("order_created_date") or order_data.get("created_at") or 0
                created_time_str = ""
                if created_date_epoch:
                    try:
                        from datetime import timezone
                        dt_created = datetime.fromtimestamp(created_date_epoch / 1000.0, tz=timezone.utc)
                        dt_created_local = dt_created.astimezone(timezone(timedelta(hours=7)))
                        created_time_str = dt_created_local.strftime("%Y-%m-%dT%H:%M:%S+07:00")
                    except Exception:
                        pass
                
                synthetic_trip = {
                    "status": f"S{status_val}" if not status_val.startswith("S") else status_val,
                    "statusName": status_name_str,
                    "postcode": order_data.get("post_code", ""),
                    "deliveryPostcode": order_data.get("delivery_post_code", ""),
                    "updatedAt": created_time_str,
                    "updatedBy": {
                        "name": order_data.get("shipper", {}).get("name") or "System",
                        "phone": order_data.get("shipper", {}).get("phone") or ""
                    },
                    "desc": order_data.get("order_name") or "Order registered",
                    "shipperName": order_data.get("shipper", {}).get("name") or "Express",
                    "is_synthetic": True,
                    # Force milestone 1 to resolve from creation date for -99 / undefined orders
                    "_created_epoch": created_date_epoch,
                }
                trips = [synthetic_trip]
            else:
                # Neither tracking nor search found the order
                await edit_or_send_requester_text(
                    msg, update, context,
                    f"❌ មិនរកឃើញការបញ្ជាទិញដែលត្រូវនឹងលេខសម្គាល់ '{query}' ទេ។"
                    if use_khmer else
                    f"❌ No order found matching ID '{query}'."
                )
                return

        # Get latest event (first in list)
        latest = trips[0]
        status = latest.get("status", "").lstrip("S")
        status_name = latest.get("statusName", "")
        po = latest.get("postcode", "")
        delivery_po = latest.get("deliveryPostcode", "")

        # /ask flow buckets are intentionally fixed and separate from report generation.
        pickup_codes = {"110", "120", "200"}
        delivery_codes = {"401", "400", "430", "420", "402", "460"}
        pending_codes = {"210", "302", "306", "310", "309", "311", "500", "300"}
        done_codes = {"410", "201", "520"}

        def get_app_action(status_code):
            actions = cfg.get("ask_app_actions", {})
            if not isinstance(actions, dict):
                return {}
            default = actions.get("default", {})
            specific = actions.get(status_code, {})
            if not isinstance(default, dict):
                default = {}
            if not isinstance(specific, dict):
                specific = {}
            merged = {**default, **specific}
            return {
                key: str(value).strip()
                for key, value in merged.items()
                if value is not None and str(value).strip()
            }

        def get_status_flow(status_code):
            flows = cfg.get("ask_status_flow", {})
            if not isinstance(flows, dict):
                return {}
            default = flows.get("default", {})
            specific = flows.get(status_code, {})
            if not isinstance(default, dict):
                default = {}
            if not isinstance(specific, dict):
                specific = {}
            merged = {**default, **specific}
            next_statuses = merged.get("next_statuses", [])
            if not isinstance(next_statuses, list):
                next_statuses = [next_statuses] if next_statuses else []
            return {
                "meaning": str(merged.get("meaning", "")).strip(),
                "next_flow": str(merged.get("next_flow", "")).strip(),
                "next_statuses": [str(s).strip() for s in next_statuses if str(s).strip()],
            }

        app_action = get_app_action(status)
        status_flow = get_status_flow(status)
        responsible_handle = po.upper() if po else (delivery_po.upper() if delivery_po else "Unknown")

        # Load reference post office mapping
        ref_df = pd.DataFrame()
        if os.path.exists(REF_PATH):
            try:
                ref_df = pd.read_csv(REF_PATH, dtype=str)
                ref_df['current_post_office'] = ref_df['current_post_office'].astype(str).str.strip().str.upper()
                ref_df['post_office_handle'] = ref_df['post_office_handle'].astype(str).str.strip().str.upper()
            except Exception as e:
                log.warning("Could not load post office lookup ref: %s", e)

        def clean_phone_helper(ph):
            if not ph or str(ph).lower() == 'nan': return "Unknown Phone"
            ph_str = str(ph).strip()
            if ph_str.startswith("+855"):
                ph_str = "0" + ph_str[4:]
            if ph_str.startswith("+"):
                ph_str = "+" + "".join(c for c in ph_str[1:] if c.isdigit())
                if ph_str.startswith("+855"):
                    ph_str = "0" + ph_str[4:]
            if ph_str.isdigit() and not ph_str.startswith("0"):
                ph_str = "0" + ph_str
            return ph_str

        handle_name = "Unknown Name"
        handle_phone = "Unknown Phone"
        
        # 1. Lookup in reference CSV
        if not ref_df.empty and responsible_handle != "Unknown":
            match_ref = ref_df[ref_df['current_post_office'] == responsible_handle]
            if not match_ref.empty:
                name_val = match_ref.iloc[0].get('customer_name', '')
                phone_val = match_ref.iloc[0].get('phone', '')
                if name_val and str(name_val).lower() != 'nan':
                    handle_name = name_val
                if phone_val:
                    handle_phone = clean_phone_helper(phone_val)
            else:
                match_handle = ref_df[ref_df['post_office_handle'] == responsible_handle]
                if not match_handle.empty:
                    name_val = match_handle.iloc[0].get('customer_name', '')
                    phone_val = match_handle.iloc[0].get('phone', '')
                    if name_val and str(name_val).lower() != 'nan':
                        handle_name = name_val
                    if phone_val:
                        handle_phone = clean_phone_helper(phone_val)

        # 2. If still Unknown, search in tracking trips for postOffice info
        if handle_name == "Unknown Name" or handle_phone == "Unknown Phone":
            for trip in trips:
                po_info = trip.get("postOffice")
                if po_info and po_info.get("postcode", "").upper() == responsible_handle:
                    if handle_name == "Unknown Name" and po_info.get("name"):
                        handle_name = po_info.get("name")
                    if handle_phone == "Unknown Phone" and po_info.get("phone"):
                        handle_phone = clean_phone_helper(po_info.get("phone"))
                    break
                    
        # 3. If still Unknown, fall back to last scan operator's info
        if handle_name == "Unknown Name" or handle_phone == "Unknown Phone":
            upd = latest.get("updatedBy")
            if upd:
                if handle_name == "Unknown Name" and upd.get("name"):
                    handle_name = upd.get("name")
                if handle_phone == "Unknown Phone" and upd.get("phone"):
                    handle_phone = clean_phone_helper(upd.get("phone"))

        next_step = ""
        stage = "Unknown"
        # Undefined / unregistered (-99, 0, 100) are treated same as pickup stage
        registered_codes = ["-99", "0", "100"]
        if status in pickup_codes or status in registered_codes:
            stage = "📥 វគ្គទទួលអីវ៉ាន់ (Pickup Stage)" if use_khmer else "📥 Pickup Stage"
            stage = "📥 វគ្គទទួលអីវ៉ាន់ (Pickup Stage)" if use_khmer else "📥 Pickup Stage"
            next_step = (
                f"ការបញ្ជាទិញកំពុងរង់ចាំការទទួលអីវ៉ាន់។\n👉 **សកម្មភាពដែលត្រូវធ្វើ:** បួគុកដែលទទួលខុសត្រូវ (**{responsible_handle}**) ត្រូវតែស្កេន និងទទួលយកកញ្ចប់អីវ៉ាន់។"
                if use_khmer else
                f"The order is waiting to be picked up.\n👉 **Action needed:** The responsible post office (**{responsible_handle}**) must scan and pick up the parcel."
            )
        elif status == "400":
            stage = "🚚 វគ្គដឹកជញ្ជូន (Delivery Stage)" if use_khmer else "🚚 Delivery Stage"
            next_step = (
                f"ការបញ្ជាទិញត្រូវបានចាត់ឲ្យបុគ្គលិកដឹកជញ្ជូន ប៉ុន្តែមិនទាន់បញ្ចប់។\n👉 **សកម្មភាពដែលត្រូវធ្វើ:** បុគ្គលិកដឹកជញ្ជូនដែលបានចាត់តាំងត្រូវបើក Delivery tab/function ហើយស្កេន ឬធ្វើបច្ចុប្បន្នភាពការដឹកជញ្ជូនឲ្យបានបញ្ចប់។"
                if use_khmer else
                "The order has been assigned to a delivery staff but is not completed.\n👉 **Action needed:** The assigned delivery staff must open the Delivery tab/function, then scan or update the order after delivery is completed."
            )
        elif status in delivery_codes:
            stage = "🚚 វគ្គដឹកជញ្ជូន (Delivery Stage)" if use_khmer else "🚚 Delivery Stage"
            next_step = (
                f"ការបញ្ជាទិញកំពុងស្ថិតក្នុងការដឹកជញ្ជូន ឬនៅហាង។\n👉 **សកម្មភាពដែលត្រូវធ្វើ:** អ្នកទទួល/អតិថិជនត្រូវមកទទួលកញ្ចប់អីវ៉ាន់ និង **ស្កេនកូដ QR** លើកម្មវិធី mExpress App ដើម្បីបង់ថ្លៃដឹកជញ្ជូន/បញ្ចប់ការដឹកជញ្ជូន។"
                if use_khmer else
                f"The order is out for delivery or at the store.\n👉 **Action needed:** The receiver/customer needs to pick up the parcel and **scan the QR code** on the mExpress App to pay the fee/complete the delivery."
            )
        elif status in pending_codes:
            stage = "⏳ វគ្គបញ្ជូន/រង់ចាំ (Pending/Transit Stage)" if use_khmer else "⏳ Pending/Transit Stage"
            next_step = (
                f"ការបញ្ជាទិញកំពុងស្ថិតក្នុងការរង់ចាំ កំពុងបញ្ជូន ឬពន្យារពេល។\n👉 **សកម្មភាពដែលត្រូវធ្វើ:** បួគុកដែលទទួលខុសត្រូវ (**{responsible_handle}**) ត្រូវតែស្កេនកញ្ចប់អីវ៉ាន់ ដើម្បីធ្វើបច្ចុប្បន្នភាពស្ថានភាពបញ្ជូន ឬដោះស្រាយការផ្អាកដឹកជញ្ជូន។"
                if use_khmer else
                f"The order is currently pending, in transit, or delayed.\n👉 **Action needed:** The responsible post office (**{responsible_handle}**) must scan the package to update its transit status or resolve the delivery hold."
            )
        elif status == "410":
            stage = "✅ បានប្រគល់ជោគជ័យ (Successfully Delivered)" if use_khmer else "✅ Successfully Delivered"
            next_step = (
                "🎉 ការបញ្ជាទិញត្រូវបានប្រគល់ជូនដោយជោគជ័យ និងបានបញ្ចប់! មិនត្រូវការស្កេនបន្ថែមទៀតទេ។"
                if use_khmer else
                "🎉 The order has been successfully delivered and completed! No further scanning is needed."
            )
        elif status in {"201", "520"}:
            stage = "✅ បានបិទ/បញ្ចប់ (Closed/Done)" if use_khmer else "✅ Closed / Done"
            next_step = (
                "ការបញ្ជាទិញត្រូវបានបិទ/បញ្ចប់។ មិនត្រូវការស្កេនដឹកជញ្ជូនបន្ថែមទៀតទេ។"
                if use_khmer else
                "The order is closed/done. No further delivery scan is needed."
            )
        else:
            stage = f"ស្ថានភាព {status}" if use_khmer else f"Status {status}"
            next_step = (
                f"ការបញ្ជាទិញបច្ចុប្បន្នស្ថិតក្នុងស្ថានភាព: *{status_name}*។\n👉 **សកម្មភាពដែលត្រូវធ្វើ:** ពិនិត្យជាមួយបួគុកដែលទទួលខុសត្រូវ (**{responsible_handle}**) ដើម្បីបញ្ជាក់ថាតើត្រូវស្កេនដែរឬទេ។"
                if use_khmer else
                f"The order is currently in status: *{status_name}*.\n👉 **Action needed:** Check with the responsible post office (**{responsible_handle}**) to verify if scanning is required."
            )

        # Translations
        T = {
            "order_tracking": "ការតាមដានការបញ្ជាទិញ" if use_khmer else "Order tracking",
            "current_stage": "វគ្គបច្ចុប្បន្ន" if use_khmer else "Current Stage",
            "current_status": "ស្ថានភាពបច្ចុប្បន្ន" if use_khmer else "Current Status",
            "current_po": "បច្ចុប្បន្ននៅបួគុក" if use_khmer else "Current Post Office",
            "latest_tracking": "Tracking ចុងក្រោយ" if use_khmer else "Latest Tracking",
            "tracking_time": "ពេលវេលា" if use_khmer else "Time",
            "tracking_by": "អ្នកធ្វើ" if use_khmer else "By",
            "tracking_note": "កំណត់ចំណាំ" if use_khmer else "Note",
            "order_flow": "វឌ្ឍនភាពលំហូរការងារ" if use_khmer else "Order Flow Progress",
            "next_flow": "លំហូរបន្ទាប់" if use_khmer else "Next Flow",
            "meaning": "ន័យស្ថានភាព" if use_khmer else "Meaning",
            "expected_next": "ស្ថានភាពបន្ទាប់ដែលរំពឹង" if use_khmer else "Expected next status",
            "who_scan": "តើនរណាត្រូវស្កេន/ចាត់ចែងការងារនេះ?" if use_khmer else "Who needs to scan/handle this?",
            "resp_po": "បួគុកទទួលខុសត្រូវ" if use_khmer else "Responsible Post Office",
            "phone": "លេខទូរស័ព្ទ" if use_khmer else "Phone",
            "app_action": "ជំហានក្នុង App" if use_khmer else "App action",
            "actor": "អ្នកធ្វើ" if use_khmer else "Actor",
            "app": "App" if use_khmer else "App",
            "tab": "Tab" if use_khmer else "Tab",
            "function": "Function" if use_khmer else "Function",
            "receiver": "អ្នកទទួល" if use_khmer else "Receiver",
            "assigned_delivery_staff": "បុគ្គលិកដឹកជញ្ជូនដែលបានចាត់តាំង" if use_khmer else "Assigned delivery staff",
            "what_to_do": "ត្រូវធ្វើ" if use_khmer else "What to do",
            "next_step_title": "ជំហានបន្ទាប់ដើម្បីជោគជ័យ" if use_khmer else "Next Step to Success",
            "pending": "រង់ចាំ" if use_khmer else "Pending",
            "created": "បានបង្កើត" if use_khmer else "Created",
            "picked_up": "បានទទួល & កំពុងបញ្ជូន" if use_khmer else "Picked Up & In Transit",
            "rec_store": "ដល់បណ្តាញទទួល/ភ្នាក់ងារ" if use_khmer else "At Receiving Store/Agent",
            "orig_hub": "ដល់មជ្ឈមណ្ឌលដើមទី" if use_khmer else "At Origin Hub",
            "dest_store": "ដល់បណ្តាញចែកចាយ/ភ្នាក់ងារ" if use_khmer else "At Destination Store/Agent",
            "delivered": "បានប្រគល់ & ជោគជ័យ" if use_khmer else "Delivered & Successful",
        }

        # Helpers
        def format_time(ts_str):
            if not ts_str: return ""
            try:
                clean_ts = ts_str.split("+")[0]
                dt = datetime.strptime(clean_ts.split(".")[0], "%Y-%m-%dT%H:%M:%S")
                return dt.strftime("%d/%m/%Y %H:%M:%S")
            except Exception:
                return ts_str

        def format_epoch_ms(epoch_ms):
            if not epoch_ms: return ""
            try:
                from datetime import timezone
                dt = datetime.fromtimestamp(epoch_ms / 1000.0, tz=timezone.utc)
                dt_local = dt.astimezone(timezone(timedelta(hours=7)))
                return dt_local.strftime("%d/%m/%Y %H:%M:%S")
            except Exception:
                return ""

        def is_hub_postcode(pc):
            if not pc: return False
            pc = str(pc).upper()
            return "MEGA" in pc or "HUB" in pc or "DVC" in pc

        def get_trip_user(t):
            upd = t.get("updatedBy", {})
            name = upd.get("name") or ""
            ho = t.get("handoverInfo", {})
            staff_code = ""
            if ho:
                staff = ho.get("handoverStaff", {})
                staff_cre = ho.get("handoverStaffCreation", {})
                if staff.get("code"):
                    staff_code = staff.get("code")
                elif staff_cre.get("code"):
                    staff_code = staff_cre.get("code")
            if not name:
                name = t.get("shipperName") or ""
            if staff_code and name:
                return f"{staff_code} - {name}"
            return name

        def get_trip_phone(t):
            upd = t.get("updatedBy", {})
            phone = upd.get("phone") or upd.get("mobile") or upd.get("telephone") or ""
            return clean_phone_helper(phone) if phone else ""

        def get_trip_note(t):
            for key in ("desc", "description", "remark", "message", "statusDesc", "content", "reason", "note"):
                value = t.get(key)
                if value and str(value).strip() and str(value).strip() != status_name:
                    return str(value).strip()

            ignored = {
                str(t.get("status", "")).strip(),
                str(t.get("statusName", "")).strip(),
                str(t.get("postcode", "")).strip(),
                str(t.get("deliveryPostcode", "")).strip(),
                str(t.get("updatedAt", "")).strip(),
            }

            def iter_strings(value):
                if isinstance(value, str):
                    yield value
                elif isinstance(value, dict):
                    for child in value.values():
                        yield from iter_strings(child)
                elif isinstance(value, list):
                    for child in value:
                        yield from iter_strings(child)

            for text in iter_strings(t):
                clean = text.strip()
                if not clean or clean in ignored:
                    continue
                lowered = clean.lower()
                if any(word in lowered for word in ("assign", "deliver", "complete", "pickup", "receive", "transit")):
                    return clean
            return ""

        chrono_trips = trips[::-1]
        
        m1 = None  # Created
        m2 = None  # Picked up & In Transit
        m3 = None  # At Receiving Store/Agent
        m4 = None  # At Origin Hub
        m5 = None  # At Destination Store/Agent
        m6 = None  # Delivered & Successful
        
        # Milestone 1: Created (S110/S100 trip or creation date from order data)
        for t in chrono_trips:
            st = t.get("status", "").lstrip("S")
            if st in ("110", "100") or t.get("is_synthetic"):
                m1 = {
                    "updatedAt": format_epoch_ms(t.get("_created_epoch", 0)) or t.get("updatedAt", ""),
                    "is_fallback": True
                } if t.get("is_synthetic") else t
                break
        if not m1 and order_data:
            created_date_epoch = order_data.get("order_created_date") or order_data.get("created_at") or 0
            if created_date_epoch:
                m1 = {
                    "updatedAt": format_epoch_ms(created_date_epoch),
                    "is_fallback": True
                }
        if not m1 and chrono_trips:
            m1 = chrono_trips[0]

        # Milestone 2: Picked up & In Transit (status 200, 210, 120)
        for t in chrono_trips:
            st = t.get("status", "").lstrip("S")
            if st in ("200", "210", "120"):
                m2 = t
                break

        # Milestone 3: At Receiving Store/Agent (status 302, 310)
        for t in chrono_trips:
            st = t.get("status", "").lstrip("S")
            if st in ("302", "310"):
                m3 = t
                break

        # Milestone 4: At Origin Hub (status 306 at hub postcode)
        for t in chrono_trips:
            st = t.get("status", "").lstrip("S")
            pc = t.get("postcode", "")
            if st == "306" and is_hub_postcode(pc):
                m4 = t
                break

        # Milestone 5: At Destination Store/Agent (status 306 or 309 at non-hub postcode)
        dest_pc = order_data.get("delivery_post_code") if order_data else None
        if not dest_pc and chrono_trips:
            dest_pc = chrono_trips[0].get("deliveryPostcode")
            
        for t in chrono_trips:
            st = t.get("status", "").lstrip("S")
            pc = t.get("postcode", "")
            if st in ("306", "309", "311", "300", "302", "420", "430", "400"):
                if dest_pc and str(pc).upper() == str(dest_pc).upper():
                    if not m5:
                        m5 = t
                        break
                elif not is_hub_postcode(pc):
                    if m4:
                        t_time = t.get("updatedAt")
                        m4_time = m4.get("updatedAt")
                        if t_time and m4_time and t_time > m4_time:
                            if not m5:
                                m5 = t
                                break
                    else:
                        if not m5:
                            m5 = t
                            break

        # Milestone 6: Delivered & Successful
        for t in chrono_trips:
            st = t.get("status", "").lstrip("S")
            if st in done_codes:
                m6 = t
                break

        def format_milestone(m, label, num):
            if not m:
                return f"⏳ {num}. {label}: {T['pending']}"
            if m.get("is_fallback"):
                time_str = m.get("updatedAt")
                return f"✅ {num}. {label}: {time_str}"
            time_str = format_time(m.get("updatedAt"))
            po = m.get("postcode")
            user = get_trip_user(m)
            po_str = f" at {po}" if po else ""
            user_str = f" by {user}" if user else ""
            return f"✅ {num}. {label}: {time_str}{po_str}{user_str}"

        timeline = [
            format_milestone(m1, T["created"], 1),
            format_milestone(m2, T["picked_up"], 2),
            format_milestone(m3, T["rec_store"], 3),
            format_milestone(m4, T["orig_hub"], 4),
            format_milestone(m5, T["dest_store"], 5),
            format_milestone(m6, T["delivered"], 6),
        ]

        latest_time = format_time(latest.get("updatedAt", ""))
        latest_user = get_trip_user(latest)
        latest_phone = get_trip_phone(latest)
        latest_note = get_trip_note(latest)

        # Extract shipper, consignee and payment details if available
        shipper_info = ""
        consignee_info = ""
        payment_info = ""
        receiver_name = "Unknown Receiver"
        receiver_phone = "Unknown Phone"
        assigned_delivery_name = "Unknown delivery staff"
        assigned_delivery_phone = "Unknown Phone"

        def extract_assigned_delivery_staff(tracking_trips):
            pattern = re.compile(
                r"Assign\s+(.+?)\s*[-–—]\s*([+0-9][+0-9\s().-]*)\s+to\s+deliver",
                re.IGNORECASE,
            )

            def iter_strings(value):
                if isinstance(value, str):
                    yield value
                elif isinstance(value, dict):
                    for child in value.values():
                        yield from iter_strings(child)
                elif isinstance(value, list):
                    for child in value:
                        yield from iter_strings(child)

            for trip in tracking_trips:
                for text in iter_strings(trip):
                    match = pattern.search(text)
                    if match:
                        return match.group(1).strip(), clean_phone_helper(match.group(2))
            return assigned_delivery_name, assigned_delivery_phone

        if status == "400":
            assigned_delivery_name, assigned_delivery_phone = extract_assigned_delivery_staff(trips)

        # 1. Extract Store Sender from Status 310 / 210 / 110 trip first
        s310_name, s310_phone = None, None
        for t in reversed(trips):
            st = str(t.get("status", "")).lstrip("S")
            if st in ("310", "210", "110", "200"):
                upd = t.get("updatedBy", {})
                un = upd.get("name") or t.get("shipperName")
                up = upd.get("phone") or upd.get("mobile")
                if un and un.upper() != "SYSTEM":
                    s310_name = un
                    s310_phone = clean_phone_helper(up) if up else ""
                    break

        if order_data:
            s_name = s310_name or order_data.get("shipper", {}).get("name") or order_data.get("seller", {}).get("name") or "N/A"
            s_phone = s310_phone or order_data.get("shipper", {}).get("phone") or order_data.get("seller", {}).get("phone") or "N/A"
            c_name = order_data.get("consignee", {}).get("name") or order_data.get("buyer", {}).get("name") or "N/A"
            c_phone = order_data.get("consignee", {}).get("phone") or order_data.get("buyer", {}).get("phone") or "N/A"
            
            def clean_phone(ph):
                if not ph: return "N/A"
                ph_str = str(ph).strip()
                if ph_str.startswith("+855"):
                    ph_str = "0" + ph_str[4:]
                return ph_str
                
            s_phone = clean_phone(s_phone)
            c_phone = clean_phone(c_phone)
            if c_name and c_name != "N/A":
                receiver_name = c_name
            if c_phone and c_phone != "N/A":
                receiver_phone = c_phone
            
            serv = order_data.get("service_name") or order_data.get("service_code") or "N/A"
            w = order_data.get("calculated_weight") or 0.0
            cod_val = order_data.get("cod_money") or 0.0
            fee_val = order_data.get("total_fees") or 0.0
            payer_val = order_data.get("payer_type") or "N/A"
            
            # Format weight: if >= 1000g, format as kg, else g
            w_str = f"{w/1000.0:.2f} kg" if w >= 1000 else f"{w:.0f} g"
            
            if use_khmer:
                shipper_info = f"👤 **អ្នកផ្ញើ (ហាង):** {s_name} ({s_phone})"
                consignee_info = f"📥 **អ្នកទទួល (អតិថិជន):** {c_name} ({c_phone})"
                payment_info = f"💰 **សេវាកម្ម:** {serv} | **ទម្ងន់:** {w_str}\n💵 **COD:** {cod_val} USD | **ថ្លៃសេវា:** {fee_val} USD ({payer_val})"
            else:
                shipper_info = f"👤 **Sender (Store):** {s_name} ({s_phone})"
                consignee_info = f"📥 **Receiver (Customer):** {c_name} ({c_phone})"
                payment_info = f"💰 **Service:** {serv} | **Weight:** {w_str}\n💵 **COD:** {cod_val} USD | **Fee:** {fee_val} USD ({payer_val})"

        # Fallback lookup in SQLite database if sender or receiver info missing
        if not shipper_info or not consignee_info or "N/A" in shipper_info or "N/A" in consignee_info:
            try:
                from search_engine import search_orders
                db_matches = search_orders(order_id, limit=1)
                if db_matches:
                    r_db = db_matches[0]
                    s_n = r_db.get("sender_name") or "N/A"
                    s_p = r_db.get("sender_phone") or "N/A"
                    r_n = r_db.get("receiver_name") or "N/A"
                    r_p = r_db.get("receiver_phone") or "N/A"
                    if (not shipper_info or "N/A" in shipper_info) and (s_n != "N/A" or s_p != "N/A"):
                        shipper_info = f"👤 **Sender (Store):** {s_n} ({s_p})" if not use_khmer else f"👤 **អ្នកផ្ញើ (ហាង):** {s_n} ({s_p})"
                    if (not consignee_info or "N/A" in consignee_info) and (r_n != "N/A" or r_p != "N/A"):
                        consignee_info = f"📥 **Receiver (Customer):** {r_n} ({r_p})" if not use_khmer else f"📥 **អ្នកទទួល (អតិថិជន):** {r_n} ({r_p})"
            except Exception:
                pass

        # Extract Origin Receiving Staff (staff who accepted package at origin branch)
        origin_staff_info = ""
        for t in chrono_trips:
            if not t.get("is_synthetic"):
                upd = t.get("updatedBy", {})
                uname = upd.get("name")
                uphone = clean_phone_helper(upd.get("phone"))
                upo = t.get("postcode") or ""
                if uname and uname.upper() != "SYSTEM" and uphone and uphone != "Unknown Phone":
                    origin_staff_info = f"🏢 **Origin Receiving Staff ({upo}):** {uname} ({uphone})" if not use_khmer else f"🏢 **បុគ្គលិកទទួលដំបូង ({upo}):** {uname} ({uphone})"
                    break

        def md_escape(text):
            return str(text).replace("\\", "\\\\").replace("_", "\\_").replace("*", "\\*").replace("`", "\\`")

        qr_url = f"https://api.qrserver.com/v1/create-qr-code/?size=350x350&data={order_id}"
        out_lines = [
            f"📦 **{T['order_tracking']}: `{md_escape(order_id)}`**  ([📱 View QR to Scan]({qr_url}))",
        ]
        if shipper_info:
            out_lines.append(shipper_info)
        if origin_staff_info:
            out_lines.append(origin_staff_info)
        if consignee_info:
            out_lines.append(consignee_info)
        if payment_info:
            out_lines.append(payment_info)

        out_lines.extend([
            f"📊 **{T['current_stage']}:** {stage}",
            f"📍 **{T['current_status']}:** {md_escape(status_name)} ({md_escape(status)})",
            f"🏢 **{T['current_po']}:** `{md_escape(po)}`\n",
        ])

        out_lines.append(f"📈 **{T['order_flow']}:**\n" + "\n".join(timeline) + "\n")

        out_lines.append(f"🏢 **{T['resp_po']}:** `{md_escape(responsible_handle)}` ({md_escape(handle_name)}) - 📞 {md_escape(handle_phone)}")

        if app_action:
            actor = app_action.get("actor", "")
            actor_lower = actor.lower()
            if "delivery staff" in actor_lower or "courier" in actor_lower or "driver" in actor_lower:
                out_lines.append(f"🚚 **{T['assigned_delivery_staff']}:** {md_escape(assigned_delivery_name)} - 📞 {md_escape(assigned_delivery_phone)}")

        out_lines.append("")
        out_lines.append(f"📋 **{T['next_step_title']}:**\n{md_escape(next_step)}")

        await edit_or_send_requester_text(msg, update, context, "\n".join(out_lines), parse_mode="Markdown")

    except Exception as e:
        log.exception("Error in /ask")
        await edit_or_send_requester_text(
            msg, update, context,
            f"❌ បរាជ័យក្នុងការវិភាគការតាមដាន: {e}" if use_khmer else f"❌ Tracking analysis failed: {e}"
        )


@pm_required_handler
async def cmd_check(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """/check <order_id1>, <order_id2>, ... — check scans and status details of multiple order IDs."""
    await delete_group_command(update, context)
    
    # Parse arguments
    raw_args = context.args or []
    # If the user pasted a single long string or multiple lines, let's extract all numbers
    full_text = " ".join(raw_args)
    
    # Check if the message is a reply to another message containing IDs
    if not full_text and update.message and update.message.reply_to_message:
        reply = update.message.reply_to_message
        full_text = reply.text or reply.caption or ""
        
    # Extract digit sequences of length >= 8
    order_ids = []
    for piece in re.split(r'[\s,;\n]+', full_text):
        piece = piece.strip()
        if piece.isdigit() and len(piece) >= 8:
            if piece not in order_ids:
                order_ids.append(piece)
                
    if not order_ids:
        await private_or_current_reply(
            update,
            context,
            "Usage: `/check <order_id1> <order_id2> ...`\n"
            "Example: `/check 3003568063 3003568385`"
        )
        return
        
    # Limit to 50 IDs to avoid rate limiting
    if len(order_ids) > 50:
        await private_or_current_reply(
            update,
            context,
            f"⚠️ Too many order IDs! Checked first 50 out of {len(order_ids)}."
        )
        order_ids = order_ids[:50]
        
    msg = await send_requester_text(
        update, context,
        f"🔍 Querying {len(order_ids)} orders from Metfone API in parallel..."
    )
    
    cfg = load_config()
    token = cfg["api"]["bearer_token"]
    headers = {
        "Authorization": f"Bearer {token}",
        "Accept": "application/json, text/plain, */*",
        "User-Agent": "Mozilla/5.0",
    }
    
    import requests
    # Query in parallel with concurrency semaphore
    sem = asyncio.Semaphore(5)
    
    async def fetch_order_info(oid):
        async with sem:
            await asyncio.sleep(0.1) # Stagger slightly
            
            # Fetch tracking
            url_track = "https://gw-express.metfone.com.kh/tms-tracking/api/v1/order-tracking"
            url_search = "https://gw-express.metfone.com.kh/tms-receiving/api/v1/orders/search"
            
            status_name = "Unknown"
            cod = 0.0
            fee = 0.0
            s402_scans = []
            s410_scans = []
            branch_code = ""
            commune_en = ""
            commune_kh = ""
            
            try:
                # 1. Search detail info
                r_search = await asyncio.to_thread(
                    requests.get, url_search, params={"order_code": oid}, headers=headers, timeout=15
                )
                if r_search.status_code == 200:
                    sdata = r_search.json()
                    cod = sdata.get("cod_money") or 0.0
                    fee = sdata.get("total_fees") or 0.0
                    branch_code = str(sdata.get("delivery_post_code") or sdata.get("post_code") or "").strip()
                    # Mapped commune
                    consignee_addr = sdata.get("consignee", {}).get("address", {})
                    for comp in consignee_addr.get("components", []):
                        if comp.get("type") in ("VILLAGE", "WARD"):
                            commune_en = comp.get("name")
                            commune_kh = comp.get("name")
                
                # 2. Tracking details
                r_track = await asyncio.to_thread(
                    requests.get, url_track, params={"order_id": oid}, headers=headers, timeout=15
                )
                if r_track.status_code == 200:
                    tdata = r_track.json()
                    trips = tdata.get("trackingTrips", [])
                    if trips:
                        latest = trips[0]
                        status_name = latest.get("statusName", "Unknown")
                        
                        # Find S402 and S410 details
                        # trips list is descending in time
                        for t in trips:
                            status = str(t.get("status", "")).lstrip("S")
                            user_name = t.get("updatedBy", {}).get("name") or t.get("shipperName") or ""
                            user_phone = t.get("updatedBy", {}).get("phone") or ""
                            user_info = f"{user_name} ({user_phone})" if user_phone else user_name
                            
                            ts = t.get("updatedAt", "")
                            formatted_ts = ts
                            if ts:
                                try:
                                    clean_ts = ts.split("+")[0]
                                    dt = datetime.strptime(clean_ts.split(".")[0], "%Y-%m-%dT%H:%M:%S")
                                    formatted_ts = dt.strftime("%d/%m/%Y %H:%M")
                                except Exception:
                                    pass
                            
                            if status == "402":
                                s402_scans.append(f"{formatted_ts} by {user_info}")
                            elif status == "410":
                                s410_scans.append(f"{formatted_ts} by {user_info}")
                                
            except Exception as e:
                status_name = f"Error: {e}"
                
            prov_kh = ""
            dist_kh = ""
            if branch_code:
                # Strip department code to get province prefix
                b_code = re.sub(r'^[A-Z]{3,5}P?\d*', '', branch_code).strip()
                if len(branch_code) >= 3:
                    prefix = branch_code[:3].upper()
                    prov_kh, _, dist_kh, _ = _map_to_administrative_division(prefix, commune_en, commune_kh)
                    
            # Join multiple scans if they exist (newest first)
            s402_str = "; ".join(s402_scans) if s402_scans else ""
            s410_str = "; ".join(s410_scans) if s410_scans else ""
            
            return {
                "Order ID": oid,
                "Status": status_name,
                "Province": prov_kh,
                "District": dist_kh,
                "COD": cod,
                "Fee": fee,
                "S402 Scans": s402_str,
                "S410 Scans": s410_str
            }

    tasks = [fetch_order_info(oid) for oid in order_ids]
    results = await asyncio.gather(*tasks)
    
    # Send text summary of the orders
    summary_lines = [f"📦 **Scan & Commission Summary ({len(order_ids)} orders):**\n"]
    for r in results:
        oid = r["Order ID"]
        status = r["Status"]
        s402 = r["S402 Scans"]
        s410 = r["S410 Scans"]
        
        status_emoji = "✅" if s410 else "🚚" if s402 else "⏳"
        line = f"{status_emoji} `{oid}`: {status}"
        if r["Province"] or r["District"]:
            line += f" ({r['Province']} - {r['District']})"
        if s402:
            # Split multiple scans for clean bullet points
            for scan in r["S402 Scans"].split("; "):
                line += f"\n   • 🚚 S402: {scan}"
        if s410:
            for scan in r["S410 Scans"].split("; "):
                line += f"\n   • ✅ S410: {scan}"
        summary_lines.append(line)
        summary_lines.append("")
        
    final_text = "\n".join(summary_lines)
    if len(final_text) > 4000:
        final_text = final_text[:4000] + "\n... (Summary truncated)"
        
    await edit_or_send_requester_text(msg, update, context, final_text)
    
    # Build Excel document for download
    try:
        import pandas as pd
        df_out = pd.DataFrame(results)
        
        tmpdir = tempfile.mkdtemp(prefix="check_")
        stamp = datetime.now().strftime("%d.%m_%HH%M")
        filename = f"CheckOrders_Details_{stamp}.xlsx"
        out_path = os.path.join(tmpdir, filename)
        
        # Write to Excel beautifully
        from openpyxl import Workbook
        from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
        from openpyxl.utils import get_column_letter

        DARK_BLUE = "172033"
        WHITE = "FFFFFF"
        BORDER_CLR = "CCCCCC"
        
        thin_border = Border(
            left=Side(style="thin", color=BORDER_CLR),
            right=Side(style="thin", color=BORDER_CLR),
            top=Side(style="thin", color=BORDER_CLR),
            bottom=Side(style="thin", color=BORDER_CLR),
        )
        
        wb = Workbook()
        ws = wb.active
        ws.title = "Order Scans"
        ws.views.sheetView[0].showGridLines = True
        
        headers_excel = list(results[0].keys())
        for col_idx, col_name in enumerate(headers_excel, 1):
            cell = ws.cell(row=1, column=col_idx, value=col_name)
            cell.font = Font(name="Calibri", bold=True, color=WHITE, size=11)
            cell.fill = PatternFill(start_color=DARK_BLUE, end_color=DARK_BLUE, fill_type="solid")
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border = thin_border
        ws.row_dimensions[1].height = 28
        
        for row_idx, row_data in enumerate(results, 2):
            for col_idx, key in enumerate(headers_excel, 1):
                val = row_data[key]
                cell = ws.cell(row=row_idx, column=col_idx, value=val)
                cell.font = Font(name="Calibri", size=10)
                cell.border = thin_border
                
        for col in ws.columns:
            max_len = max(len(str(cell.value or '')) for cell in col)
            col_letter = get_column_letter(col[0].column)
            ws.column_dimensions[col_letter].width = min(max(max_len + 3, 12), 50)
            
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = f"A1:{get_column_letter(len(headers_excel))}{len(results)+1}"
        
        wb.save(out_path)
        
        # Send Document
        with open(out_path, "rb") as f:
            await send_requester_document(update, context, f, filename)
            
    except Exception as e:
        log.exception("Failed to build Excel for /check")
        await send_requester_text(update, context, f"❌ Failed to build Excel file: {e}")


@pm_required_handler
async def cmd_trace(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """/trace <bill_id> — diagnose and track errors/status for a specific bill ID."""
    await delete_group_command(update, context)
    
    args = [a.strip() for a in (context.args or []) if a.strip()]
    bill_id = " ".join(args)
    
    if not bill_id and update.message and update.message.reply_to_message:
        reply = update.message.reply_to_message
        bill_id = (reply.text or reply.caption or "").strip()
        
    # Extract the first digit sequence of length >= 8 if it's a long text
    if bill_id:
        match = re.search(r'\d{8,}', bill_id)
        if match:
            bill_id = match.group(0)

    # Clean up non-digits just in case
    bill_id = "".join(c for c in bill_id if c.isdigit())
    
    if not bill_id:
        await private_or_current_reply(
            update,
            context,
            "Usage: `/trace <bill_id>`\n"
            "Example: `/trace 3003568063`"
        )
        return
        
    msg = await send_requester_text(
        update, context,
        f"🔍 Initiating trace for Bill ID `{bill_id}`..."
    )
    
    trace_results = []
    trace_results.append(f"🔍 **Trace Report for Bill ID:** `{bill_id}`\n")
    
    # ── 1. Check ignore / test lists ──
    trace_results.append("📋 **1. Settings & Config Check:**")
    is_ignored = False
    if os.path.exists("test_bills.txt"):
        try:
            with open("test_bills.txt", "r", encoding="utf-8") as f:
                for line in f:
                    if line.strip() == bill_id:
                        is_ignored = True
                        break
        except Exception as e:
            trace_results.append(f"   • ❌ Failed to read `test_bills.txt`: {e}")
            
    if is_ignored:
        trace_results.append(f"   • ⚠️ Listed in `test_bills.txt` (This bill is marked as TEST/IGNORED).")
    else:
        trace_results.append(f"   • ✅ Not listed in `test_bills.txt` (Active/not ignored).")
        
    # Check delayed_bills.json
    is_delayed = False
    if os.path.exists("delayed_bills.json"):
        try:
            with open("delayed_bills.json", "r", encoding="utf-8") as f:
                delayed = json.load(f)
                if bill_id in delayed:
                    is_delayed = True
                    trace_results.append(f"   • ⚠️ Listed in `delayed_bills.json` (Delayed until: `{delayed[bill_id]}`).")
        except Exception as e:
            trace_results.append(f"   • ❌ Failed to read `delayed_bills.json`: {e}")
            
    if not is_delayed:
        trace_results.append("   • ✅ Not listed in `delayed_bills.json` (No delay filter applied).")
    
    trace_results.append("")
    
    # ── 2. Search local Excel cache ──
    trace_results.append("📂 **2. Local Excel Cache Search:**")
    found_in_cache = False
    cache_file = os.path.join(HERE, "cache", "latest_detail.xlsx")
    if os.path.exists(cache_file):
        try:
            import pandas as pd
            xl = pd.ExcelFile(cache_file)
            for sheet in xl.sheet_names:
                df = xl.parse(sheet)
                order_col = None
                for col in df.columns:
                    if "order" in str(col).lower() or "id" in str(col).lower():
                        order_col = col
                        break
                if not order_col and len(df.columns) > 3:
                    for col in df.columns:
                        if df[col].astype(str).str.contains(bill_id, na=False).any():
                            order_col = col
                            break
                if order_col is not None:
                    matches = df[df[order_col].astype(str).str.strip() == bill_id]
                    if not matches.empty:
                        found_in_cache = True
                        row = matches.iloc[0]
                        trace_results.append(f"   • ✅ Found in cached Excel (`{sheet}` sheet):")
                        status = row.get("CURRENT STATUS") or row.get("Current Status") or row.get("Trạng thái hiện tại") or "N/A"
                        po = row.get("CURRENT POST OFFICE") or row.get("Current Post Office") or row.get("Bưu cục hiện tại") or "N/A"
                        sender = row.get("SENDER") or row.get("Sender") or "N/A"
                        receiver = row.get("RECEIVER") or row.get("Receiver") or "N/A"
                        trace_results.append(f"     - Status: `{status}`")
                        trace_results.append(f"     - Post Office: `{po}`")
                        trace_results.append(f"     - Sender: `{sender}` | Receiver: `{receiver}`")
                        break
            if not found_in_cache:
                trace_results.append("   • ℹ️ Bill ID not found in the latest cached Excel data.")
        except Exception as e:
            trace_results.append(f"   • ❌ Error reading cached Excel: {e}")
    else:
        trace_results.append("   • ℹ️ No cached Excel data found.")
        
    trace_results.append("")
    
    # ── 3. Live API Diagnostics ──
    trace_results.append("⚡ **3. Live TMS API Diagnostics:**")
    cfg = load_config()
    token = cfg.get("api", {}).get("bearer_token")
    if not token:
        trace_results.append("   • ❌ API Token is missing in config.json")
    else:
        import requests
        headers = {
            "Authorization": f"Bearer {token}",
            "Accept": "application/json, text/plain, */*",
            "User-Agent": "Mozilla/5.0",
        }
        
        # Live Search call
        url_search = "https://gw-express.metfone.com.kh/tms-receiving/api/v1/orders/search"
        trace_results.append("   *Querying Search API...*")
        try:
            r_search = await asyncio.to_thread(
                requests.get, url_search, params={"order_code": bill_id}, headers=headers, timeout=15
            )
            trace_results.append(f"   • Status: HTTP `{r_search.status_code}`")
            if r_search.status_code == 200:
                sdata = r_search.json()
                cod = sdata.get("cod_money") or 0.0
                fee = sdata.get("total_fees") or 0.0
                bp = sdata.get("delivery_post_code") or sdata.get("post_code") or "N/A"
                trace_results.append(f"     - COD: `{cod}` | Fee: `{fee}` | Post Code: `{bp}`")
            else:
                trace_results.append(f"     - Error Response: `{r_search.text[:250]}`")
        except Exception as e:
            trace_results.append(f"     - Search API Connection Failed: `{e}`")
            
        # Live Tracking call
        url_track = "https://gw-express.metfone.com.kh/tms-tracking/api/v1/order-tracking"
        trace_results.append("   *Querying Tracking API...*")
        try:
            r_track = await asyncio.to_thread(
                requests.get, url_track, params={"order_id": bill_id}, headers=headers, timeout=15
            )
            trace_results.append(f"   • Status: HTTP `{r_track.status_code}`")
            if r_track.status_code == 200:
                tdata = r_track.json()
                trips = tdata.get("trackingTrips", [])
                if trips:
                    latest = trips[0]
                    status_name = latest.get("statusName", "Unknown")
                    trace_results.append(f"     - Current Status Name: `{status_name}`")
                    trace_results.append(f"     - Recent Trip Scans (max 3):")
                    for idx, t in enumerate(trips[:3]):
                        status = t.get("status", "")
                        desc = t.get("description", "")
                        user = t.get("updatedBy", {}).get("name") or t.get("shipperName") or "System"
                        ts = t.get("updatedAt") or ""
                        trace_results.append(f"       {idx+1}. [{status}] {desc} (by {user} at {ts})")
                else:
                    trace_results.append("     - No tracking trips found.")
            else:
                trace_results.append(f"     - Error Response: `{r_track.text[:250]}`")
        except Exception as e:
            trace_results.append(f"     - Tracking API Connection Failed: `{e}`")
            
    trace_results.append("")
    
    # ── 4. Log Scan ──
    trace_results.append("📝 **4. Execution Logs (`bot.log`):**")
    log_file = os.path.join(HERE, "bot.log")
    if os.path.exists(log_file):
        try:
            matching_lines = []
            with open(log_file, "r", encoding="utf-8", errors="ignore") as lf:
                for line in lf:
                    if bill_id in line:
                        matching_lines.append(line.strip())
            if matching_lines:
                trace_results.append(f"   • Found {len(matching_lines)} matching log entry/entries:")
                for ml in matching_lines[-15:]:
                    trace_results.append(f"     - `{ml[:150]}`")
            else:
                trace_results.append("   • No logs found matching this Bill ID.")
        except Exception as e:
            trace_results.append(f"   • ❌ Failed to read log file: {e}")
    else:
        trace_results.append("   • ℹ/ No log file `bot.log` exists yet.")
        
    final_text = "\n".join(trace_results)
    if len(final_text) > 4000:
        final_text = final_text[:4000] + "\n... (Trace truncated due to size)"
        
    await edit_or_send_requester_text(msg, update, context, final_text)


@user_guard
async def cmd_qr(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """/qr <order_id> - generate a QR code for the order ID to scan on the phone screen."""
    await delete_group_command(update, context)
    args = [a.strip() for a in (context.args or []) if a.strip()]
    if not args:
        await private_or_current_reply(
            update,
            context,
            "Usage: `/qr <order_id>`\n"
            "Example: `/qr 3003516618`"
        )
        return

    order_id = args[0]
    chat_id = requester_chat_id(update)
    if not chat_id:
        return

    qr_url = f"https://api.qrserver.com/v1/create-qr-code/?size=350x350&data={order_id}"
    caption = (
        f"📦 **QR Code for Order:** `{order_id}`\n"
        f"👉 *Scan this directly off your screen using your mExpress App.*"
    )

    try:
        await safe_api_call(
            context.bot.send_photo,
            chat_id=chat_id,
            photo=qr_url,
            caption=caption,
            parse_mode="Markdown"
        )
    except Exception as e:
        log.warning("Could not send QR photo to private chat %s: %s", chat_id, e)
        if is_group_chat(update) and update.effective_chat:
            try:
                await safe_api_call(
                    context.bot.send_photo,
                    chat_id=update.effective_chat.id,
                    photo=qr_url,
                    caption=caption,
                    parse_mode="Markdown",
                )
            except Exception as e2:
                log.warning("Fallback send QR to group failed: %s", e2)
        else:
            await private_or_current_reply(
                update,
                context,
                f"❌ Failed to send QR code image: {e}\n"
                f"🔗 Click here to view QR code: {qr_url}"
            )


# ── Push handler ───────────────────────────────────────────────────────────────

@pm_required_handler
async def run_push(
    update: Update,
    context: ContextTypes.DEFAULT_TYPE,
    force_test: bool = False,
):
    await delete_group_command(update, context)

    cfg = load_config()
    allowed = cfg["telegram"].get("allowed_chat_ids") or []
    chat_id = update.effective_chat.id
    if allowed and chat_id not in allowed:
        await private_or_current_reply(update, context, "This chat is not allowed to use the bot.")
        return

    paused = is_paused(cfg)
    test_mode = force_test or paused
    if force_test:
        await private_or_current_reply(
            update,
            context,
            "🧪 TEST MODE.\n"
            "Reports will only be sent to you. Nothing will be forwarded to groups."
        )
    elif paused:
        await private_or_current_reply(
            update,
            context,
            "⏸ Bot is paused — running in TEST mode.\n"
            "Reports will only be sent to you (no group forwarding)."
        )

    msg = await send_requester_text(update, context, "Downloading data from OPS...")
    if is_group_chat(update) and msg is None:
        log.warning(
            "Stopping push because private requester messages are unavailable."
        )
        return

    tmpdir = tempfile.mkdtemp(prefix="push_")
    track_report_dir(tmpdir)
    stamp  = datetime.now().strftime("%d.%m_%HH%M")
    src    = os.path.join(tmpdir, f"export_{stamp}.xlsx")

    try:
        # ── Determine zone/mode before download so we fetch only needed branches ──
        raw_args = [arg.strip() for arg in getattr(context, "args", []) if arg.strip()]

        # ── Extract inline remark: `push remark: some text` ──────────────────
        # Supports: `push remark: text`, `push zone remark: text`, etc.
        inline_remark = None
        clean_args = []
        remark_started = False
        remark_parts = []
        for tok in raw_args:
            if remark_started:
                remark_parts.append(tok)
            elif tok.lower().startswith("remark:"):
                # e.g. "remark:Please" or "remark: Please" (next tokens)
                remark_started = True
                suffix = tok[len("remark:"):].strip()
                if suffix:
                    remark_parts.append(suffix)
            elif tok.lower() == "remark":
                # next token should be ":" or text
                remark_started = True
            else:
                clean_args.append(tok)
        if remark_parts:
            inline_remark = " ".join(remark_parts).strip()
        raw_args = clean_args  # remove remark tokens before zone parsing

        # Check for force refresh in arguments
        force_refresh = False
        if "force" in [a.lower() for a in raw_args]:
            force_refresh = True
            raw_args = [a for a in raw_args if a.lower() != "force"]

        zone_key_arg = raw_args[0].lower() if raw_args else None
        total_zones = cfg.get("total_zones", {})
        zone_branches_map = cfg.get("zone_branches", {})

        zone_override_branch = None
        zone_mode = None
        target_handles = []

        if zone_key_arg == "all":
            # "push all" = ALL 23 groups + 5 zones → fetch ALL branches
            all_branches = []
            for zb in zone_branches_map.values():
                for b in zb.split(","):
                    if b.strip() and b.strip() not in all_branches:
                        all_branches.append(b.strip())
            # Also include default branch_code branches
            for b in cfg["api"].get("branch_code", "").split(","):
                if b.strip() and b.strip() not in all_branches:
                    all_branches.append(b.strip())
            zone_override_branch = ",".join(all_branches)
            zone_mode = "ALL"

        elif zone_key_arg == "zone":
            # "push zone" = 5 zone groups only → fetch ALL zone branches
            all_branches = []
            for zb in zone_branches_map.values():
                for b in zb.split(","):
                    if b.strip() and b.strip() not in all_branches:
                        all_branches.append(b.strip())
            zone_override_branch = ",".join(all_branches)
            zone_mode = "ZONE"

        elif zone_key_arg and zone_key_arg in total_zones:
            # "push zone5" = specific zone only
            if zone_key_arg in zone_branches_map:
                zone_override_branch = zone_branches_map[zone_key_arg]
            target_handles = [h.upper() for h in total_zones[zone_key_arg]]
            zone_mode = zone_key_arg.upper()

        else:
            # plain "push" or "push <handle>" (e.g. "push SVAP001")
            target_handles = [arg.upper() for arg in raw_args if arg]
            if target_handles and not zone_override_branch:
                # Fast branch detection for specific handles
                detected_branches = set()
                for h in target_handles:
                    b_found = None
                    for zk, hlist in total_zones.items():
                        if h in [x.upper() for x in hlist]:
                            zb = zone_branches_map.get(zk, "")
                            for b in zb.split(","):
                                if b.strip() and h.startswith(b.strip()):
                                    b_found = b.strip()
                                    break
                            if not b_found and zb:
                                b_found = zb.split(",")[0].strip()
                            break
                    if not b_found:
                        b_found = h[:3]
                    if b_found:
                        detected_branches.add(b_found)
                if detected_branches:
                    for mega in ["MEGA", "MEGA1", "DVCMEGA1"]:
                        detected_branches.add(mega)
                    zone_override_branch = ",".join(sorted(detected_branches))

        downloader.download_detail(cfg["api"], src, branch_code=zone_override_branch, force_refresh=force_refresh)
        
        rev_src = os.path.join(tmpdir, "latest_revenue.xlsx")
        try:
            downloader.download_revenue_detail(cfg["api"], rev_src, force_refresh=force_refresh)
        except Exception as e:
            log.error(f"Failed to download revenue detail: {e}")
            rev_src = None

        msg = await edit_or_send_requester_text(
            msg, update, context, "Download done. Generating reports..."
        )

        if not os.path.exists(REF_PATH):
            await edit_or_send_requester_text(
                msg, update, context, f"Error: Missing reference file {REF_PATH}"
            )
            return

        if force_test and not target_handles and not zone_mode and is_group_chat(update):
            group_handles = get_forward_mapping(cfg).get(str(update.effective_chat.id), [])
            group_handles = [h.upper() for h in group_handles if h and h != "*"]
            if group_handles:
                target_handles = group_handles

        mode = get_mode(cfg)
        result = generate_report.generate_reports_from_data(
            src, REF_PATH, tmpdir, return_metadata=True, mode=mode, target_handles=target_handles, revenue_path=rev_src
        )
        update_webapp_cache(result)
        update_dashboard_cache(result)
        save_highlight_history(result)

        new_ignored_count = result.get('new_ignored_count', 0)

        # ── Apply handle filters ──────────────────────────────────────────
        if target_handles:
            result["handle_results"] = [
                hr for hr in result["handle_results"]
                if hr["handle"] in target_handles
            ]
        if target_handles or zone_mode:
            overall = {"Pickup": 0, "Delivery": 0, "Transit": 0, "Branch": 0}
            for hr in result["handle_results"]:
                for k in overall:
                    overall[k] += hr["handle_counts"].get(k, 0)
            result["overall_counts"] = overall
            grand_total = sum(overall.values())
            label = f"{zone_mode} " if zone_mode else ""
            result["summary_caption"] = "\n".join([
                f"📋 {label}Daily Report  {datetime.now().strftime('%d/%m/%Y %H:%M')}",
                f"Delivery: {overall.get('Delivery',0)}  |  Assign Deliver: {overall.get('Branch',0)}  |  Pickup: {overall.get('Pickup',0)}  |  Handover to Mega: {overall.get('Transit',0)}",
                f"Grand Total: {grand_total}",
            ])

        if inline_remark:
            for hr in result["handle_results"]:
                hr["remark"] += f"  |  Remark: {inline_remark}"
            if "summary_caption" in result:
                result["summary_caption"] += f"\n📝 Remark: {inline_remark}"

        if not result["handle_results"]:
            await edit_or_send_requester_text(
                msg, update, context, "No data found.",
            )
            return

        msg = await edit_or_send_requester_text(msg, update, context, "Sending images...")

        # ── Pre-calculate covered handles for hiding ─────────────────────────
        forward_groups = get_all_forward_groups(cfg)
        forward_mapping = get_forward_mapping(cfg)
        covered_handles = set()
        actual_forward_groups = []

        for group_id_str in forward_groups:
            allowed = forward_mapping.get(str(group_id_str), ["*"])
            wants_all = "*" in allowed
            will_receive = False
            for hr in result["handle_results"]:
                if wants_all or hr["handle"] in allowed:
                    will_receive = True
                    break
            if will_receive:
                actual_forward_groups.append(group_id_str)

        forward_groups = actual_forward_groups
        
        if not test_mode and forward_groups:
            for group_id_str in forward_groups:
                allowed = forward_mapping.get(str(group_id_str), ["*"])
                if "*" in allowed:
                    covered_handles.add("*")
                else:
                    covered_handles.update(allowed)
            # Also cover all handles mapped in any zone so they aren't sent to the requester
            for zone_key, zone_hlist in total_zones.items():
                covered_handles.update([h.upper() for h in zone_hlist if h])

        # ── Send to the requester ─────────────────────────────────────────────
        for hr in result["handle_results"]:
            handle_str = hr["handle"]
            
            # Hide from requester in production mode (only show in test mode)
            if not test_mode and handle_str in covered_handles:
                continue
                
            for hf in hr["handle_files"]:
                try:
                    img_buf = excel_to_image.excel_to_image(hf["path"])
                    img_buf.name = f"{hr['handle']}.png"
                    await send_requester_photo(update, context, img_buf)
                except Exception as e:
                    log.warning(f"Image render failed {hr['handle']}: {e}")
            await send_requester_text(update, context, hr["remark"])

        with open(result["final_xlsx"], "rb") as f:
            await send_requester_document(
                update,
                context,
                f,
                os.path.basename(result["final_xlsx"]),
            )
        await send_requester_text(update, context, result["summary_caption"])

        # ── Forward to groups only outside test mode ──────────────────────────
        if test_mode:
            if force_test:
                test_text = (
                    f"🧪 TEST MODE — not forwarded to groups.\n"
                    f"Data: {datetime.now().strftime('%d.%m.%Y %H:%M')}"
                )
            else:
                test_text = (
                    f"⏸ TEST MODE — not forwarded to groups.\n"
                    f"Use /resume to enable group forwarding.\n"
                    f"Data: {datetime.now().strftime('%d.%m.%Y %H:%M')}"
                )
            await edit_or_send_requester_text(msg, update, context, test_text)
            return

        # ── Zone group forwarding (push zone / push all) ────────────────────
        zone_fwd_map = cfg.get("zone_forward_mapping", {})
        send_to_regular = zone_mode is None or zone_mode == "ALL"
        send_to_zones = zone_mode in ("ZONE", "ALL") or (
            zone_mode and zone_mode.startswith("ZONE") and zone_mode != "ALL"
        )

        # Forward to regular groups (push / push all)
        if send_to_regular and forward_groups:
            context.user_data["pending_forward"] = {
                "result": result,
                "forward_groups": forward_groups,
                "forward_mapping": forward_mapping,
                "created_at": datetime.now().isoformat(timespec="seconds"),
            }
            await edit_or_send_requester_text(
                msg, update, context,
                f"Forwarding to {len(forward_groups)} group(s)...",
            )
            payload = context.user_data.pop("pending_forward", None)
            sent_groups = await forward_result_to_groups(context, payload)
            await send_requester_text(
                update, context,
                f"✅ Forwarded to {sent_groups} group(s).",
            )

        # Forward to zone groups (push zone / push all / push zone5)
        # Each zone gets a summary image + total Excel (same as /total)
        if send_to_zones and zone_fwd_map:
            import pandas as pd
            zone_sent = 0
            for group_id_str, zone_key in zone_fwd_map.items():
                # For "push zone5", only forward to zone5 group
                if zone_mode and zone_mode not in ("ZONE", "ALL"):
                    if zone_key != zone_mode.lower():
                        continue

                zone_handles = [h.upper() for h in total_zones.get(zone_key, [])]
                if not zone_handles:
                    continue

                # Filter handle_results for this zone
                zone_results = [
                    hr for hr in result["handle_results"]
                    if hr["handle"] in zone_handles
                ]
                if not zone_results:
                    continue

                try:
                    group_id = int(group_id_str)
                except ValueError:
                    group_id = group_id_str

                # Calculate zone totals
                zone_overall = {"Pickup": 0, "Delivery": 0, "Transit": 0, "Branch": 0}
                for hr in zone_results:
                    for k in zone_overall:
                        zone_overall[k] += hr["handle_counts"].get(k, 0)
                zone_grand = sum(zone_overall.values())

                zone_label = zone_key.upper()

                # ── Calculate Fee + COD + VIP totals per handle and zone total ──
                zone_fee_total = 0.0
                zone_cod_total = 0.0
                zone_vip_counts = {}
                fee_cod_lines = []

                # Count VIPs across all report types for this zone
                for hr in zone_results:
                    h = hr["handle"]
                    h_v = 0
                    for rn in ["Pickup", "Delivery", "Transit", "Branch"]:
                        df_tab = result.get("type_data", {}).get(rn)
                        if df_tab is not None and not df_tab.empty and "VIP" in df_tab.columns:
                            flt_col = "CURRENT POST OFFICE" if rn == "Transit" else "POST OFFICE HANDLE"
                            if flt_col in df_tab.columns:
                                df_h_v = df_tab[df_tab[flt_col] == h]
                                h_v += (df_h_v["VIP"] == "VIP").sum()
                    zone_vip_counts[h] = h_v
                zone_vip_total = sum(zone_vip_counts.values())

                for rn in ["Delivery", "Branch"]:
                    df_fc = result.get("type_data", {}).get(rn)
                    if df_fc is None or df_fc.empty:
                        continue
                    if "POST OFFICE HANDLE" not in df_fc.columns:
                        continue
                    df_z_fc = df_fc[df_fc["POST OFFICE HANDLE"].isin(zone_handles)].copy()
                    if df_z_fc.empty:
                        continue
                    fee_col = next((c for c in df_z_fc.columns if "TOTAL FEE" in c.upper()), None)
                    cod_col = next((c for c in df_z_fc.columns if c.upper().startswith("COD")), None)
                    if fee_col:
                        zone_fee_total += pd.to_numeric(df_z_fc[fee_col], errors="coerce").fillna(0).sum()
                    if cod_col:
                        zone_cod_total += pd.to_numeric(df_z_fc[cod_col], errors="coerce").fillna(0).sum()

                # Per-branch breakdown
                for hr in zone_results:
                    h = hr["handle"]
                    h_fee = 0.0
                    h_cod = 0.0
                    h_vip = zone_vip_counts.get(h, 0)
                    for rn in ["Delivery", "Branch"]:
                        df_fc = result.get("type_data", {}).get(rn)
                        if df_fc is None or df_fc.empty:
                            continue
                        if "POST OFFICE HANDLE" not in df_fc.columns:
                            continue
                        df_h = df_fc[df_fc["POST OFFICE HANDLE"] == h]
                        fee_col = next((c for c in df_fc.columns if "TOTAL FEE" in c.upper()), None)
                        cod_col = next((c for c in df_fc.columns if c.upper().startswith("COD")), None)
                        if fee_col:
                            h_fee += pd.to_numeric(df_h[fee_col], errors="coerce").fillna(0).sum()
                        if cod_col:
                            h_cod += pd.to_numeric(df_h[cod_col], errors="coerce").fillna(0).sum()
                    total_orders = sum(hr["handle_counts"].get(k, 0) for k in ["Pickup","Delivery","Transit","Branch"])
                    fee_cod_lines.append(
                        f"  {h}: {total_orders} orders (VIP: {h_vip}) | Fee: ${h_fee:.2f} | COD: ${h_cod:.2f}"
                    )

                zone_caption = "\n".join([
                    f"📋 {zone_label} Report  {datetime.now().strftime('%d/%m/%Y %H:%M')}",
                    f"Delivery: {zone_overall.get('Delivery',0)}  |  Assign Deliver: {zone_overall.get('Branch',0)}  |  Pickup: {zone_overall.get('Pickup',0)}  |  Handover to Mega: {zone_overall.get('Transit',0)}",
                    f"Grand Total: {zone_grand}  |  VIP: {zone_vip_total}  |  Fee: ${zone_fee_total:.2f}  |  COD: ${zone_cod_total:.2f}",
                ])
                if inline_remark:
                    zone_caption += f"\n📝 Remark: {inline_remark}"


                # Build zone-filtered result for total Excel
                zone_result = {**result, "handle_results": zone_results, "overall_counts": zone_overall}
                # Filter type_data DataFrames
                zone_result["type_data"] = {}
                for rn in ["Pickup", "Delivery", "Transit", "Branch"]:
                    df = result.get("type_data", {}).get(rn)
                    if df is not None and not df.empty:
                        filter_col = "POST OFFICE HANDLE"
                        if filter_col in df.columns:
                            zone_result["type_data"][rn] = df[df[filter_col].isin(zone_handles)].copy()
                        else:
                            zone_result["type_data"][rn] = df.copy()
                    else:
                        zone_result["type_data"][rn] = pd.DataFrame()

                try:
                    # ── Build exact day_date_counts, urgent_counts, fee, cod, vip for zone image ──
                    zh_set = set(zone_handles)
                    zone_day_date_counts = {h: result.get("day_date_counts", {}).get(h, {}) for h in zh_set if h in result.get("day_date_counts", {})}
                    zone_urgent_counts   = {h: result.get("urgent_counts", {}).get(h, {}) for h in zh_set if h in result.get("urgent_counts", {})}
                    zone_vip_counts      = {h: result.get("vip_counts", {}).get(h, 0) for h in zh_set if h in result.get("vip_counts", {})}
                    zone_fee_counts      = {h: result.get("fee_counts", {}).get(h, 0.0) for h in zh_set if h in result.get("fee_counts", {})}
                    zone_cod_counts      = {h: result.get("cod_counts", {}).get(h, 0.0) for h in zh_set if h in result.get("cod_counts", {})}

                    # 1. Summary image
                    img_buf = generate_summary.build_summary_image(
                        zone_results,
                        zone_overall,
                        zone_label=zone_label,
                        day_date_counts=zone_day_date_counts if zone_day_date_counts else None,
                        urgent_counts=zone_urgent_counts if zone_urgent_counts else None,
                        fee_counts=zone_fee_counts if any(zone_fee_counts.values()) else None,
                        cod_counts=zone_cod_counts if any(zone_cod_counts.values()) else None,
                        vip_counts=zone_vip_counts if any(zone_vip_counts.values()) else None,
                    )
                    img_buf.name = f"{zone_key}_summary.png"
                    await safe_api_call(context.bot.send_photo, chat_id=group_id, photo=img_buf)
                    await asyncio.sleep(0.5)

                    # 2. Total Excel
                    zone_xlsx = os.path.join(tmpdir, f"Total_{zone_label}_{stamp}.xlsx")
                    generate_summary.build_total_excel(zone_result, zone_xlsx, lang="en")
                    with open(zone_xlsx, "rb") as f:
                        await safe_api_call(
                            context.bot.send_document,
                            chat_id=group_id,
                            document=f,
                            filename=os.path.basename(zone_xlsx),
                            caption=zone_caption,
                        )
                    zone_sent += 1
                    await asyncio.sleep(1.0)
                except Exception as e:
                    log.error(f"Zone forward to {group_id} ({zone_key}): {e}")

            if zone_sent:
                await send_requester_text(
                    update, context,
                    f"✅ Forwarded to {zone_sent} zone group(s).",
                )

        if not send_to_regular and not send_to_zones:
            # Specific zone push (push zone5) — already handled above via send_to_zones
            pass

        final_msg_text = f"Done. {datetime.now().strftime('%d.%m.%Y %H:%M')}"
        if new_ignored_count > 0:
            final_msg_text += f"\n🧹 Old Pickups auto-ignored today: {new_ignored_count} bills"

        await edit_or_send_requester_text(
            msg, update, context,
            final_msg_text
        )

        # (Remark prompt removed as requested)

    except Exception as e:
        log.exception("Error in run_push")
        # 1. Try to send the detailed error privately to the user who requested it.
        user_id = requester_chat_id(update)
        if user_id:
            try:
                await safe_api_call(
                    context.bot.send_message,
                    chat_id=user_id,
                    text=f"❌ Error in run_push: {e}"
                )
            except Exception as pm_err:
                log.warning(f"Failed to send private error message to user {user_id}: {pm_err}")
                
        # 2. Clean up or update the status message in the group so there is no lingering status or error message there.
        if msg and is_group_chat(update):
            try:
                await context.bot.delete_message(
                    chat_id=update.effective_chat.id,
                    message_id=msg.message_id
                )
            except Exception as del_err:
                log.warning(f"Failed to delete status message in group: {del_err}")
                # Fallback: Edit it to a simple status so it doesn't look stuck, without leaking the detailed error
                try:
                    await safe_api_call(
                        msg.edit_text,
                        "❌ Push failed."
                    )
                except Exception:
                    pass


@user_guard
async def cmd_test(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Run a private test push without forwarding anything to groups."""
    await run_push(update, context, force_test=True)


@user_guard
async def cmd_add(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Add test bill(s) directly to the ignore list and the configured Excel/CSV file."""
    await delete_group_command(update, context)
    if not context.args:
        # Check if they used it as `/testbill add <id>` fallback
        if update.message and update.message.text and update.message.text.startswith("/testbill"):
            args = context.args[1:] if len(context.args) > 1 else []
            if not args:
                await private_or_current_reply(update, context, "Usage: `/add <bill_id>`", parse_mode="Markdown")
                return
        else:
            await private_or_current_reply(update, context, "Usage: `/add <bill_id1> <bill_id2> ...`", parse_mode="Markdown")
            return
    else:
        if context.args[0].lower() == "add":
            args = context.args[1:]
        else:
            args = context.args

    txt_path = "test_bills.txt"
    current_ids = set()
    if os.path.exists(txt_path):
        with open(txt_path, "r", encoding="utf-8") as f:
            for line in f:
                val = line.strip()
                if val:
                    current_ids.add(val)

    added = []
    already_exist = []
    for bill_id in args:
        bill_id = bill_id.strip()
        if bill_id:
            if bill_id in current_ids:
                already_exist.append(bill_id)
            else:
                current_ids.add(bill_id)
                added.append(bill_id)
    
    if added:
        # 1. Update text file
        with open(txt_path, "w", encoding="utf-8") as f:
            for bid in sorted(current_ids):
                f.write(bid + "\n")
        
        # 2. Update configured Excel/CSV file
        cfg = load_config()
        tr_cfg = cfg.get("test_receipts", {})
        if tr_cfg.get("enabled"):
            path = tr_cfg.get("path")
            if path:
                try:
                    import pandas as pd
                    df_test = None
                    order_col = "ORDER ID"
                    
                    if os.path.exists(path):
                        if path.lower().endswith((".xlsx", ".xls")):
                            df_test = pd.read_excel(path, dtype=str)
                        else:
                            df_test = pd.read_csv(path, dtype=str, keep_default_na=False)
                        
                        df_test = df_test.fillna("")
                        if not df_test.empty:
                            found_col = next(
                                (
                                    c for c in df_test.columns
                                    if "order" in str(c).lower()
                                    or "phi" in str(c).lower()
                                    or "shipment" in str(c).lower()
                                ),
                                df_test.columns[0]
                            )
                            if found_col:
                                order_col = found_col
                    
                    if df_test is None or df_test.empty:
                        df_test = pd.DataFrame(columns=[order_col])
                    
                    # Read current items in the file to avoid duplicates
                    file_ids = set(df_test[order_col].astype(str).str.strip().tolist())
                    new_rows = []
                    for bid in added:
                        if bid not in file_ids:
                            new_rows.append({order_col: bid})
                        else:
                            if bid not in already_exist:
                                already_exist.append(bid)
                    
                    if new_rows:
                        df_new = pd.DataFrame(new_rows)
                        df_test = pd.concat([df_test, df_new], ignore_index=True)
                        
                        if path.lower().endswith((".xlsx", ".xls")):
                            df_test.to_excel(path, index=False)
                        else:
                            df_test.to_csv(path, index=False, encoding="utf-8-sig")
                except Exception as e:
                    log.error(f"Could not update test file at {path}: {e}")

        cfg = load_config()
        all_ids = generate_report.load_test_order_ids(cfg)
        msg_text = f"✅ Added test bills: {', '.join(added)}\n"
        if already_exist:
            msg_text += f"⚠️ Already in the list: {', '.join(already_exist)}\n"
        msg_text += f"Total ignored test bills: {len(all_ids)} (synchronised with your test file)"
        await private_or_current_reply(update, context, msg_text)
    else:
        if already_exist:
            await private_or_current_reply(
                update, 
                context, 
                f"⚠️ All specified bills ({', '.join(already_exist)}) are **already** in your ignore list."
            )
        else:
            await private_or_current_reply(update, context, "No bills were specified.")


@user_guard
async def cmd_remove(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Remove test bill(s) from the ignore list and the configured Excel/CSV file."""
    await delete_group_command(update, context)
    
    if context.args and context.args[0].lower() in ("remove", "delete", "rm", "del"):
        args = context.args[1:]
    else:
        args = context.args

    if not args:
        await private_or_current_reply(update, context, "Usage: `/remove <bill_id>`", parse_mode="Markdown")
        return

    txt_path = "test_bills.txt"
    current_ids = set()
    if os.path.exists(txt_path):
        with open(txt_path, "r", encoding="utf-8") as f:
            for line in f:
                val = line.strip()
                if val:
                    current_ids.add(val)

    removed = []
    for bill_id in args:
        bill_id = bill_id.strip()
        if bill_id in current_ids:
            current_ids.remove(bill_id)
        # We track all arguments as potentially removed from the excel/csv file
        removed.append(bill_id)
            
    # 1. Update text file
    with open(txt_path, "w", encoding="utf-8") as f:
        for bid in sorted(current_ids):
            f.write(bid + "\n")
            
    # 2. Update Excel/CSV file
    cfg = load_config()
    tr_cfg = cfg.get("test_receipts", {})
    file_removed_count = 0
    if tr_cfg.get("enabled"):
        path = tr_cfg.get("path")
        if path and os.path.exists(path):
            try:
                import pandas as pd
                if path.lower().endswith((".xlsx", ".xls")):
                    df_test = pd.read_excel(path, dtype=str)
                else:
                    df_test = pd.read_csv(path, dtype=str, keep_default_na=False)
                
                df_test = df_test.fillna("")
                if not df_test.empty:
                    order_col = next(
                        (
                            c for c in df_test.columns
                            if "order" in str(c).lower()
                            or "phi" in str(c).lower()
                            or "shipment" in str(c).lower()
                        ),
                        df_test.columns[0]
                    )
                    if order_col:
                        initial_len = len(df_test)
                        # Remove the matched IDs
                        df_test = df_test[~df_test[order_col].astype(str).str.strip().isin(removed)].copy()
                        file_removed_count = initial_len - len(df_test)
                        
                        if file_removed_count > 0:
                            if path.lower().endswith((".xlsx", ".xls")):
                                df_test.to_excel(path, index=False)
                            else:
                                df_test.to_csv(path, index=False, encoding="utf-8-sig")
            except Exception as e:
                log.error(f"Could not update test file at {path}: {e}")

    cfg = load_config()
    all_ids = generate_report.load_test_order_ids(cfg)
    await private_or_current_reply(
        update, 
        context, 
        f"❌ Removed test bills from ignore list: {', '.join(removed)}\n"
        f"Total ignored test bills remaining: {len(all_ids)} (removed {file_removed_count} directly from your test file)"
    )


@user_guard
async def cmd_list(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """List all test bills currently ignored by exporting them to an Excel file."""
    await delete_group_command(update, context)
    cfg = load_config()
    current_ids = generate_report.load_test_order_ids(cfg)
                    
    if not current_ids:
        await private_or_current_reply(update, context, "No test bills in the ignore list.")
        return

    import pandas as pd
    import io
    
    # Create Excel in memory
    df = pd.DataFrame(sorted(current_ids), columns=["Ignored Bill ID"])
    bio = io.BytesIO()
    with pd.ExcelWriter(bio, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Test Bills")
    bio.seek(0)
    
    msg = await send_requester_text(
        update, 
        context, 
        f"Generating Excel list of ignored test bills ({len(current_ids)})..."
    )
    
    try:
        await send_requester_document(
            update,
            context,
            bio,
            filename=f"ignored_test_bills_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        )
        await edit_or_send_requester_text(
            msg, 
            update, 
            context, 
            f"✅ Exported {len(current_ids)} ignored test bills to Excel."
        )
    except Exception as e:
        log.error(f"Error sending test bills Excel: {e}")
        await edit_or_send_requester_text(
            msg, 
            update, 
            context, 
            f"Error generating Excel: {e}"
        )


@user_guard
async def cmd_delay(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Delay a bill (ignore it) until a certain date or for a number of days.
    Usage: /delay <bill_id> <date_or_days>
    Examples: /delay 2900492262 25.06.2026
              /delay 2900492262 3
    """
    await delete_group_command(update, context)
    if not context.args or len(context.args) < 2:
        await private_or_current_reply(
            update, 
            context, 
            "Usage:\n`/delay <bill_id> <date_or_days>`\n\n"
            "Examples:\n"
            "`/delay 2900492262 25.06.2026` (ignore until June 25)\n"
            "`/delay 2900492262 3` (ignore for 3 days)",
            parse_mode="Markdown"
        )
        return

    bill_id = context.args[0].strip()
    val = context.args[1].strip()
    delay_path = "delayed_bills.json"

    # Try parsing value
    exp_date_str = None
    
    # 1. Parse as number of days
    try:
        days = int(val)
        exp_date_str = (datetime.now() + timedelta(days=days)).date().isoformat()
    except (ValueError, OverflowError):
        pass

    # 2. Parse as DD.MM.YYYY or DD/MM/YYYY or YYYY-MM-DD
    if not exp_date_str:
        for fmt in ("%d.%m.%Y", "%d/%m/%Y", "%Y-%m-%d"):
            try:
                dt = datetime.strptime(val, fmt)
                exp_date_str = dt.date().isoformat()
                break
            except ValueError:
                pass

    if not exp_date_str:
        await private_or_current_reply(
            update,
            context,
            "❌ Invalid date or days format.\n"
            "Please use a number of days (e.g., `3`) or a date format like `DD.MM.YYYY` (e.g., `25.06.2026`)."
        )
        return

    # Load current delayed bills
    delayed = {}
    if os.path.exists(delay_path):
        try:
            with open(delay_path, "r", encoding="utf-8") as f:
                delayed = json.load(f)
        except Exception:
            pass

    delayed[bill_id] = exp_date_str

    with open(delay_path, "w", encoding="utf-8") as f:
        json.dump(delayed, f, ensure_ascii=False, indent=2)

    # Format expiration date nicely for user
    exp_date_nice = datetime.strptime(exp_date_str, "%Y-%m-%d").strftime("%d.%m.%Y")
    await private_or_current_reply(
        update,
        context,
        f"✅ Bill `{bill_id}` is now delayed (ignored) until **{exp_date_nice}**.",
        parse_mode="Markdown"
    )


@user_guard
async def cmd_undelay(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Remove the delay on a bill so it stops being ignored.
    Usage: /undelay <bill_id>
    """
    await delete_group_command(update, context)
    if not context.args:
        await private_or_current_reply(update, context, "Usage: `/undelay <bill_id>`", parse_mode="Markdown")
        return

    bill_id = context.args[0].strip()
    delay_path = "delayed_bills.json"

    if not os.path.exists(delay_path):
        await private_or_current_reply(update, context, "No delayed bills exist.")
        return

    try:
        with open(delay_path, "r", encoding="utf-8") as f:
            delayed = json.load(f)
    except Exception:
        delayed = {}

    if bill_id in delayed:
        del delayed[bill_id]
        with open(delay_path, "w", encoding="utf-8") as f:
            json.dump(delayed, f, ensure_ascii=False, indent=2)
        await private_or_current_reply(update, context, f"✅ Removed delay on bill `{bill_id}`. It will now be processed in reports.", parse_mode="Markdown")
    else:
        await private_or_current_reply(update, context, f"Bill `{bill_id}` was not found in the delayed bills list.", parse_mode="Markdown")


@pm_required_handler
async def cmd_delaylist(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """List all currently delayed bills and their expiration dates."""
    await delete_group_command(update, context)
    delay_path = "delayed_bills.json"

    if not os.path.exists(delay_path):
        await private_or_current_reply(update, context, "No delayed bills in the list.")
        return

    try:
        with open(delay_path, "r", encoding="utf-8") as f:
            delayed = json.load(f)
    except Exception:
        delayed = {}

    if not delayed:
        await private_or_current_reply(update, context, "No delayed bills in the list.")
        return

    lines = []
    today = datetime.now().date()
    
    # Sort by expiration date
    sorted_delayed = sorted(delayed.items(), key=lambda x: x[1])
    
    for bill_id, exp_date_str in sorted_delayed:
        try:
            exp_date = datetime.strptime(exp_date_str, "%Y-%m-%d").date()
            days_left = (exp_date - today).days
            exp_date_nice = exp_date.strftime("%d.%m.%Y")
            lines.append(f"• `{bill_id}`: Resumes on {exp_date_nice} ({days_left} days left)")
        except Exception:
            lines.append(f"• `{bill_id}`: exp {exp_date_str}")

    text = f"📋 **Currently Delayed Bills ({len(delayed)}):**\n" + "\n".join(lines)
    await private_or_current_reply(update, context, text, parse_mode="Markdown")


@user_guard
async def cmd_app(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """/app — Open the interactive Mini App explorer."""
    await delete_group_command(update, context)
    cfg = load_config()
    webapp_url = cfg["telegram"].get("webapp_url", "http://localhost:8080")
    
    from telegram import InlineKeyboardMarkup, InlineKeyboardButton, WebAppInfo
    keyboard = [
        [InlineKeyboardButton("📱 Open Data Explorer", web_app=WebAppInfo(url=webapp_url))]
    ]
    reply_markup = InlineKeyboardMarkup(keyboard)
    
    await private_or_current_reply(
        update,
        context,
        "🔍 **Interactive Data Explorer**\n\n"
        "Click the button below to search, filter, and view delivery details directly inside Telegram.",
        reply_markup=reply_markup,
        parse_mode="Markdown"
    )


@pm_required_handler
async def cmd_report(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """/report — generate Excel with all active bills (excludes done statuses)."""
    await delete_group_command(update, context)
    cfg = load_config()
    msg = await send_requester_text(update, context, "Building active bills report...")
    tmpdir = tempfile.mkdtemp(prefix="report_")
    track_report_dir(tmpdir)
    stamp = datetime.now().strftime("%d.%m_%HH%M")
    src = os.path.join(tmpdir, f"export_{stamp}.xlsx")
    try:
        downloader.download_detail(cfg["api"], src)
        out_xlsx = os.path.join(tmpdir, f"Active_Bills_{stamp}.xlsx")
        count = report_cmd.build_report_excel(src, out_xlsx, cfg)
        caption = f"Active Bills Report {datetime.now().strftime('%d/%m/%Y %H:%M')}\nTotal: {count} bills"
        with open(out_xlsx, "rb") as f:
            await send_requester_document(update, context, f, os.path.basename(out_xlsx), caption=caption)
        await edit_or_send_requester_text(msg, update, context, f"Done. {count} active bills exported.")
    except Exception as e:
        log.exception("Error in /report")
        await edit_or_send_requester_text(msg, update, context, f"Error: {e}")


async def on_text(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not update.message:
        return
    cfg = load_config()
    if not is_user_allowed(update, cfg):
        log.info("Ignoring message from unauthorized user %s", update.effective_user and update.effective_user.id)
        return
    keyword = cfg["telegram"].get("keyword", "push").lower().strip()
    text    = (update.message.text or "").strip()
    parts   = text.split()
    lower_parts = [p.lower() for p in parts]

    pending = context.user_data.get("pending_forward")
    answer = lower_parts[0].strip(".,!?") if lower_parts else ""
    yes_answers = {"yes", "y", "ok", "send", "confirm", "forward"}
    no_answers = {"no", "n", "cancel", "stop"}

    if pending and answer in yes_answers | no_answers:
        await delete_group_command(update, context)
        if answer in no_answers:
            context.user_data.pop("pending_forward", None)
            await private_or_current_reply(
                update,
                context,
                "Cancelled. Report was not forwarded to any group.",
            )
            return

        payload = context.user_data.pop("pending_forward", None)
        if not payload:
            await private_or_current_reply(
                update,
                context,
                "No pending report to forward.",
            )
            return

        await private_or_current_reply(
            update,
            context,
            f"Confirmed. Forwarding to {len(payload['forward_groups'])} group(s)...",
        )
        sent_groups = await forward_result_to_groups(context, payload)
        await private_or_current_reply(
            update,
            context,
            f"Done. Forwarded to {sent_groups} group(s). "
            f"{datetime.now().strftime('%d.%m.%Y %H:%M')}",
        )
        return
    
    # Only trigger push keyword in Private Chat (PM) AND when the message explicitly begins with 'push'
    # In group chats, plain text conversations (e.g. 'Push driver to assign') will NEVER trigger reports
    if not is_group_chat(update) and lower_parts and lower_parts[0] == keyword:
        context.args = parts[1:]
        await run_push(update, context)




# ── /forward command ───────────────────────────────────────────────────────────

@pm_required_handler
async def cmd_forward(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """/forward - Send all zone reports to forward group using cached data."""
    await delete_group_command(update, context)
    cfg = load_config()

    FORWARD_GROUP_ID = None  # disabled

    cache_file = os.path.join(HERE, 'cache', 'latest_detail.xlsx')
    if not os.path.exists(cache_file):
        await send_requester_text(update, context, "No cached data. Run push first.")
        return

    msg = await send_requester_text(update, context, "Building zone reports from cached data...")

    tmpdir = tempfile.mkdtemp(prefix="forward_")
    track_report_dir(tmpdir)
    stamp = datetime.now().strftime("%d.%m_%HH%M")

    try:
        mode = get_mode(cfg)
        result = generate_report.generate_reports_from_data(
            cache_file, REF_PATH, tmpdir, return_metadata=True, mode=mode,
        )

        total_zones = cfg.get("total_zones", {})
        import pandas as pd

        zone_keys = sorted(total_zones.keys())
        sent_count = 0

        for zone_key in zone_keys:
            zone_handles = [h.upper() for h in total_zones.get(zone_key, [])]
            if not zone_handles:
                continue

            zone_results = [
                hr for hr in result["handle_results"]
                if hr["handle"] in zone_handles
            ]

            zone_overall = {"Pickup": 0, "Delivery": 0, "Transit": 0, "Branch": 0}
            for hr in zone_results:
                for k in zone_overall:
                    zone_overall[k] += hr["handle_counts"].get(k, 0)
            zone_grand = sum(zone_overall.values())

            zone_label = zone_key.upper()
            zone_caption = "\n".join([
                f"\U0001f4cb {zone_label} Report  {datetime.now().strftime('%d/%m/%Y %H:%M')}",
                f"Pickup: {zone_overall['Pickup']}  |  Delivery: {zone_overall['Delivery']}  |  Transit: {zone_overall.get('Transit',0)}  |  Branch: {zone_overall.get('Branch',0)}",
                f"Grand Total: {zone_grand}",
            ])

            zone_result = {**result, "handle_results": zone_results, "overall_counts": zone_overall}
            zone_result["type_data"] = {}
            for rn in ["Pickup", "Delivery", "Transit", "Branch"]:
                df = result.get("type_data", {}).get(rn)
                if df is not None and not df.empty:
                    filter_col = "POST OFFICE HANDLE"
                    if filter_col in df.columns:
                        zone_result["type_data"][rn] = df[df[filter_col].isin(zone_handles)].copy()
                    else:
                        zone_result["type_data"][rn] = df.copy()
                else:
                    zone_result["type_data"][rn] = pd.DataFrame()

            try:
                zh_set = set(zone_handles)
                zone_day_date_counts = {h: result.get("day_date_counts", {}).get(h, {}) for h in zh_set if h in result.get("day_date_counts", {})}
                zone_urgent_counts   = {h: result.get("urgent_counts", {}).get(h, {}) for h in zh_set if h in result.get("urgent_counts", {})}
                zone_vip_counts      = {h: result.get("vip_counts", {}).get(h, 0) for h in zh_set if h in result.get("vip_counts", {})}
                zone_fee_counts      = {h: result.get("fee_counts", {}).get(h, 0.0) for h in zh_set if h in result.get("fee_counts", {})}
                zone_cod_counts      = {h: result.get("cod_counts", {}).get(h, 0.0) for h in zh_set if h in result.get("cod_counts", {})}

                img_buf = generate_summary.build_summary_image(
                    zone_results, zone_overall,
                    zone_label=zone_label,
                    day_date_counts=zone_day_date_counts if zone_day_date_counts else None,
                    urgent_counts=zone_urgent_counts if zone_urgent_counts else None,
                    fee_counts=zone_fee_counts if any(zone_fee_counts.values()) else None,
                    cod_counts=zone_cod_counts if any(zone_cod_counts.values()) else None,
                    vip_counts=zone_vip_counts if any(zone_vip_counts.values()) else None,
                )
                img_buf.name = f"{zone_key}_summary.png"
                await safe_api_call(context.bot.send_photo, chat_id=FORWARD_GROUP_ID, photo=img_buf)
                await asyncio.sleep(0.5)

                zone_xlsx = os.path.join(tmpdir, f"Total_{zone_label}_{stamp}.xlsx")
                generate_summary.build_total_excel(zone_result, zone_xlsx)
                with open(zone_xlsx, "rb") as f:
                    await safe_api_call(
                        context.bot.send_document,
                        chat_id=FORWARD_GROUP_ID,
                        document=f,
                        filename=os.path.basename(zone_xlsx),
                        caption=zone_caption,
                    )
                sent_count += 1
                await asyncio.sleep(1.0)
            except Exception as e:
                log.error(f"Forward {zone_key}: {e}")

        await edit_or_send_requester_text(
            msg, update, context,
            f"\u2705 Forwarded {sent_count} zone reports to forward group."
        )
    except Exception as e:
        log.exception("Error in /forward")
MEDIA_GROUP_FILES = {}

@pm_required_handler
async def cmd_notice(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """
    /notice [target] <message> — Send a custom text remark, notice, or attached documents directly to groups without running a report.
    Targets:
      /notice <message>              -> Send remark to all regular branch groups
      /notice zone <message>         -> Send remark to all zone groups
      /notice all <message>          -> Send remark to ALL groups (branches + zones)
      /notice KANP001 <message>      -> Send remark to KANP001 group
    Supports document/photo attachments if sent as caption!
    """
    await delete_group_command(update, context)
    cfg = load_config()

    message_obj = update.effective_message
    if not message_obj:
        return

    text_input = message_obj.caption or message_obj.text or ""
    
    # Collect media group files if album upload
    mg_id = message_obj.media_group_id
    current_file = message_obj.document or (message_obj.photo[-1] if message_obj.photo else None)
    
    if mg_id and current_file:
        MEDIA_GROUP_FILES.setdefault(mg_id, [])
        if current_file not in MEDIA_GROUP_FILES[mg_id]:
            MEDIA_GROUP_FILES[mg_id].append(current_file)
            
    # Check if this update has the command prefix (e.g. /notice)
    has_command = bool(re.search(r"^/(notice|announce|remark|sendmsg|msg)\b", text_input, re.IGNORECASE))
    
    if mg_id and not has_command:
        # File part of media group without command - already stored in buffer
        return

    if not text_input and not current_file:
        await private_or_current_reply(
            update,
            context,
            "📢 Usage: /notice [target] <your message>\n\n"
            "Sends a custom remark or document directly to groups (no report downloaded).\n\n"
            "Examples:\n"
            "  /notice Please clear all pending orders today!\n"
            "  /notice zone Urgent: check inventory before 5 PM\n"
            "  /notice all Important announcement for all branches\n"
            "  /notice KANP001 Please call receivers first\n\n"
            "💡 You can also attach a document/file/photo with /notice in the caption!"
        )
        return

    # Parse target and text content
    cleaned = re.sub(r"^/(notice|announce|remark|sendmsg|msg)(?:@\w+)?\s*", "", text_input, flags=re.IGNORECASE).strip()
    words = cleaned.split()
    first_word = words[0].lower() if words else ""

    target_type = "branches"
    text_content = cleaned

    if first_word in ("zone", "zones"):
        target_type = "zones"
        text_content = re.sub(r"^\bzones?\b\s*", "", cleaned, flags=re.IGNORECASE).strip()
    elif first_word in ("all", "everyone"):
        target_type = "all"
        text_content = re.sub(r"^\b(all|everyone)\b\s*", "", cleaned, flags=re.IGNORECASE).strip()
    elif first_word and (
        first_word.upper() in [h.upper() for zb in cfg.get("zone_branches", {}).values() for h in str(zb).split(",") if h.strip()]
        or re.match(r"^[A-Z]{3,4}P?\d*$", first_word, re.I)
    ):
        target_type = first_word.upper()
        text_content = re.sub(rf"^{re.escape(first_word)}\s*", "", cleaned, flags=re.IGNORECASE).strip()

    # Determine files to send
    attached_files = []
    if mg_id:
        await asyncio.sleep(1.0)
        attached_files = MEDIA_GROUP_FILES.pop(mg_id, [current_file] if current_file else [])
    elif current_file:
        attached_files = [current_file]

    if not text_content and not attached_files:
        await private_or_current_reply(update, context, "Please specify a message text or attach a file to send.")
        return

    forward_mapping = get_forward_mapping(cfg)
    all_branch_groups = get_all_forward_groups(cfg)
    zone_groups = list(cfg.get("zone_forward_mapping", {}).keys())

    target_group_ids = []
    if target_type == "branches":
        target_group_ids = all_branch_groups
    elif target_type == "zones":
        target_group_ids = zone_groups
    elif target_type == "all":
        target_group_ids = list(dict.fromkeys(all_branch_groups + zone_groups))
    else:
        for gid, handles in forward_mapping.items():
            if target_type in [h.upper() for h in handles]:
                target_group_ids.append(gid)

    if not target_group_ids:
        await private_or_current_reply(update, context, f"No group found for target: {target_type}")
        return

    msg = await send_requester_text(update, context, f"📤 Sending notice to {len(target_group_ids)} group(s)...")

    formatted_msg = f"📢 REMARK / NOTICE\n{text_content}\n\n⏰ {datetime.now().strftime('%d/%m/%Y %H:%M')}" if text_content else f"📢 REMARK / NOTICE — {datetime.now().strftime('%d/%m/%Y %H:%M')}"

    sent_count = 0
    for gid in target_group_ids:
        try:
            group_id = int(gid) if str(gid).lstrip("-").isdigit() else gid
            if attached_files:
                for idx_f, f_item in enumerate(attached_files):
                    cap = formatted_msg if idx_f == 0 else ""
                    if isinstance(f_item, telegram.Document):
                        await safe_api_call(
                            context.bot.send_document,
                            chat_id=group_id,
                            document=f_item.file_id,
                            caption=cap,
                        )
                    else:
                        await safe_api_call(
                            context.bot.send_photo,
                            chat_id=group_id,
                            photo=f_item.file_id,
                            caption=cap,
                        )
                    await asyncio.sleep(0.3)
            else:
                await safe_api_call(
                    context.bot.send_message,
                    chat_id=group_id,
                    text=formatted_msg,
                )
            sent_count += 1
            await asyncio.sleep(0.3)
        except Exception as e:
            log.error(f"Notice send failed to {gid}: {e}")

    await edit_or_send_requester_text(
        msg, update, context, f"✅ Notice sent successfully to {sent_count} group(s)."
    )




# ── Main ───────────────────────────────────────────────────────────────────────

DASHBOARD_TOKENS_PATH = os.path.join(HERE, "dashboard_tokens.json")

def generate_dashboard_token():
    """Generate a random 8-char alphanumeric token and save it with 24h expiry."""
    import secrets
    import string
    alphabet = string.ascii_uppercase + string.digits
    token = ''.join(secrets.choice(alphabet) for _ in range(8))

    now = datetime.now()
    expires = now + timedelta(hours=24)

    # Load existing tokens
    data = {"tokens": []}
    if os.path.exists(DASHBOARD_TOKENS_PATH):
        try:
            with open(DASHBOARD_TOKENS_PATH, encoding="utf-8") as f:
                data = json.load(f)
        except Exception:
            pass

    # Prune expired tokens
    data["tokens"] = [
        t for t in data.get("tokens", [])
        if datetime.fromisoformat(t["expires"]) > now
    ]

    # Add new token
    data["tokens"].append({
        "token": token,
        "created": now.isoformat(),
        "expires": expires.isoformat(),
    })

    with open(DASHBOARD_TOKENS_PATH, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

    return token, expires


@user_guard
async def cmd_adminthean(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Generate a daily rotating access key for the web dashboard."""
    await delete_group_command(update, context)

    token, expires = generate_dashboard_token()
    expires_str = expires.strftime("%d/%m/%Y %H:%M")

    await private_or_current_reply(
        update, context,
        f"🔑 Dashboard Access Key Generated\n\n"
        f"Key: `{token}`\n"
        f"Expires: {expires_str}\n\n"
        f"Use this key to login at your dashboard URL.",
        parse_mode="Markdown"
    )


def main():
    # Start the WebApp HTTP Server in a background daemon thread
    server_thread = threading.Thread(target=start_webapp_server, daemon=True)
    server_thread.start()

    cfg   = load_config()
    tg    = cfg["telegram"]
    token = tg["bot_token"]
    if "DIEN_" in token:
        raise SystemExit("Set bot_token in config.json first.")

    builder = (
        Application.builder()
        .token(token)
        .connect_timeout(30.0)
        .read_timeout(30.0)
        .write_timeout(30.0)
        .pool_timeout(30.0)
        .get_updates_connect_timeout(30.0)
        .get_updates_read_timeout(30.0)
    )

    proxy_url = tg.get("proxy_url")
    if proxy_url:
        builder = builder.proxy(proxy_url).get_updates_proxy(proxy_url)
        log.info("Using proxy: %s", proxy_url)

    app = builder.build()
    app.add_handler(CommandHandler("app",        cmd_app))
    app.add_handler(CommandHandler("push",       run_push))
    app.add_handler(CommandHandler("total",      cmd_total))
    app.add_handler(CommandHandler("morning",    cmd_morning))
    app.add_handler(CommandHandler("tpg",        cmd_tpg))
    app.add_handler(CommandHandler("compare",    cmd_compare))
    app.add_handler(CommandHandler("exportlogs", cmd_export_logs))
    app.add_handler(CommandHandler("logs",       cmd_export_logs))
    app.add_handler(CommandHandler("totalkpi",   cmd_total_kpi))
    app.add_handler(CommandHandler("kpi",        cmd_kpi10h))
    app.add_handler(CommandHandler("kpi10h",     cmd_kpi10h))
    app.add_handler(CommandHandler("under10h",   cmd_kpi10h))
    app.add_handler(CommandHandler("kpi10",      cmd_kpi10h))
    app.add_handler(CommandHandler("tomorrow",   cmd_tomorrow))
    app.add_handler(CommandHandler("penalty",     cmd_penalty))
    app.add_handler(CommandHandler("speed",       cmd_speed))
    app.add_handler(CommandHandler("today",       cmd_today))
    app.add_handler(CommandHandler("daily",       cmd_today))
    app.add_handler(CommandHandler("delayed",     cmd_delayed))
    app.add_handler(CommandHandler("ge3",         cmd_delayed))
    app.add_handler(CommandHandler("backlog",     cmd_delayed))
    app.add_handler(CommandHandler("overdue",     cmd_delayed))
    app.add_handler(CommandHandler("done",        cmd_done))
    app.add_handler(CommandHandler("undelayed",   cmd_undelayed))
    app.add_handler(CommandHandler("undone",      cmd_undelayed))
    app.add_handler(CommandHandler("donelist",    cmd_donelist))
    app.add_handler(CommandHandler("lag",         cmd_lag))
    app.add_handler(CommandHandler("mismatch",    cmd_lag))
    app.add_handler(CommandHandler("audit",       cmd_lag))
    app.add_handler(CommandHandler("synclag",     cmd_lag))
    app.add_handler(CommandHandler("allthetime",  cmd_allthetime))


    app.add_handler(CommandHandler("vs",         cmd_vs))
    app.add_handler(CommandHandler("vs2",        cmd_vs2))
    app.add_handler(CommandHandler("help",       cmd_help))
    app.add_handler(CommandHandler("pause",      cmd_pause))
    app.add_handler(CommandHandler("stop",       cmd_pause))
    app.add_handler(CommandHandler("resume",     cmd_resume))
    app.add_handler(CommandHandler("start",      cmd_resume))
    app.add_handler(CommandHandler("status",     cmd_status))
    app.add_handler(CommandHandler("statues",    cmd_statues))
    app.add_handler(CommandHandler("statuses",   cmd_statues))
    app.add_handler(CommandHandler("mode",       cmd_mode))
    app.add_handler(CommandHandler("register",   cmd_register))
    app.add_handler(CommandHandler("unregister", cmd_unregister))
    app.add_handler(CommandHandler("groups",     cmd_groups))
    app.add_handler(CommandHandler("export",     cmd_export))
    app.add_handler(CommandHandler("exportall2", cmd_export_all2))
    app.add_handler(CommandHandler("export_all2", cmd_export_all2))
    app.add_handler(CommandHandler("find",       cmd_find))
    app.add_handler(CommandHandler("ask",        cmd_ask))
    app.add_handler(CommandHandler("check",      cmd_check))
    app.add_handler(CommandHandler("qr",         cmd_qr))
    app.add_handler(CommandHandler("trace",      cmd_trace))
    app.add_handler(CommandHandler("add",        cmd_add))
    app.add_handler(CommandHandler("remove",     cmd_remove))
    app.add_handler(CommandHandler("del",        cmd_remove))
    app.add_handler(CommandHandler("list",       cmd_list))
    app.add_handler(CommandHandler("testbill",   cmd_add)) # fallback
    app.add_handler(CommandHandler("delay",      cmd_delay))
    app.add_handler(CommandHandler("undelay",    cmd_undelay))
    app.add_handler(CommandHandler("delaylist",  cmd_delaylist))
    app.add_handler(CommandHandler("clean",      cmd_clean))
    app.add_handler(CommandHandler("report",     cmd_report))
    app.add_handler(CommandHandler("deletereport", cmd_delete_report))
    app.add_handler(CommandHandler("delreport",    cmd_delete_report))
    app.add_handler(CommandHandler("forward",  cmd_forward))
    app.add_handler(CommandHandler("notice",   cmd_notice))
    app.add_handler(CommandHandler("announce", cmd_notice))
    app.add_handler(CommandHandler("remark",   cmd_notice))
    app.add_handler(CommandHandler("sendmsg",  cmd_notice))
    app.add_handler(CommandHandler("adminthean", cmd_adminthean))
    app.add_handler(MessageHandler(filters.CaptionRegex(r"^/(notice|announce|remark|sendmsg|msg)\b"), cmd_notice))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, on_text))

    log.info("Bot running. Commands: push, /total, /vs, /vs2, /export, /find, /ask, /check, /trace, /statues, /help, /pause, /resume, /status, /mode, /register, /groups, /add, /remove, /list, /delay, /undelay, /delaylist, /clean, /qr, /deletereport")
    try:
        app.run_polling(allowed_updates=Update.ALL_TYPES)
    except Exception as e:
        if "Conflict" in str(e):
            log.error("Another bot instance is already running. Stop it first, then restart.")
        else:
            raise


if __name__ == "__main__":
    main()
