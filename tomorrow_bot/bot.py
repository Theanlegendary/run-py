#!/usr/bin/env python3
"""
Tomorrow Shipments Standalone Bot
Specialized bot for generating Shipments Tomorrow Reports (00:00 - 06:00 Midnight/Morning).
"""

import os
import sys
import json
import logging
import tempfile
import asyncio
from datetime import datetime

from telegram import Update
from telegram.ext import (
    Application,
    CommandHandler,
    ContextTypes,
)

# ── Logging Configuration ────────────────────────────────────────────────────
logging.basicConfig(
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
    level=logging.INFO
)
log = logging.getLogger("tomorrow_bot")

HERE = os.path.dirname(os.path.abspath(__file__))
CONFIG_PATH = os.path.join(HERE, "config.json")

def load_config():
    with open(CONFIG_PATH, "r", encoding="utf-8") as f:
        return json.load(f)

def save_config(cfg):
    with open(CONFIG_PATH, "w", encoding="utf-8") as f:
        json.dump(cfg, f, indent=2, ensure_ascii=False)

# ── Handlers ──────────────────────────────────────────────────────────────────

async def cmd_start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Start command / unpause bot."""
    cfg = load_config()
    cfg["telegram"]["paused"] = False
    save_config(cfg)
    
    welcome_text = (
        "🚚 *Shipments Tomorrow Bot Online!*\n\n"
        "This bot generates *Shipments Tomorrow Reports (00:00 - 06:00 Morning)*.\n\n"
        "📌 *Available Commands:*\n"
        "• `/tomorrow` — Generate for Zone 1\n"
        "• `/tomorrow zone 2` — Generate for Zone 2 (or zone 3, 4, 5)\n"
        "• `/tomorrow all` — Generate & forward for ALL 5 zones\n"
        "• `/tomorrow <post_office>` — Generate for specific post office (e.g. `SVAP001`, `PNPP014`)\n"
        "• `/stop` — Pause bot (prevent group forwarding)\n"
        "• `/status` — View current status\n"
        "• `/help` — Show help message"
    )
    await update.message.reply_text(welcome_text, parse_mode="Markdown")


async def cmd_stop(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Pause bot."""
    cfg = load_config()
    cfg["telegram"]["paused"] = True
    save_config(cfg)
    await update.message.reply_text(
        "🛑 *Bot PAUSED.*\nGroup forwarding is disabled. Reports will only be sent directly to you.",
        parse_mode="Markdown"
    )


async def cmd_status(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Check status."""
    cfg = load_config()
    paused = cfg.get("telegram", {}).get("paused", False)
    status_str = "⏸ PAUSED (Private Only)" if paused else "▶️ ACTIVE (Broadcasting to groups)"
    await update.message.reply_text(
        f"📊 *Bot Status:* {status_str}\n"
        f"🕒 *Server Time:* `{datetime.now().strftime('%d/%m/%Y %H:%M:%S')}`\n"
        f"🕒 *Filter Window:* `00:00:00 - 06:00:00 (Morning)`",
        parse_mode="Markdown"
    )


async def cmd_help(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Show help."""
    await cmd_start(update, context)


async def cmd_tomorrow(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Generates exact SHIPMENTS TOMORROW REPORT (00:00 - 06:00) with Excel & Executive Summary Image."""
    cfg = load_config()
    paused = cfg.get("telegram", {}).get("paused", False)

    args = [a.strip() for a in (context.args or []) if a.strip()]
    target_label = " ".join(args) if args else "Zone 1"

    status_msg = await update.message.reply_text(f"⏳ Generating SHIPMENTS TOMORROW REPORT for *{target_label}* (00:00 - 06:00)...", parse_mode="Markdown")
    
    tmpdir = tempfile.mkdtemp(prefix="tomorrow_")
    stamp = datetime.now().strftime("%d.%m_%HH%M")
    src = os.path.join(tmpdir, f"export_{stamp}.xlsx")

    try:
        import downloader
        import shipments_tomorrow

        # Download raw detail export from Metfone API
        downloader.download_detail(cfg["api"], src, force_refresh=True)

        tgt_upper = target_label.upper().replace(" ", "")

        if tgt_upper in ("ALL", "MEGA"):
            # Process for all 5 Zones
            zone_fwd_map = cfg.get("zone_forward_mapping", {})
            total_sent = 0
            
            for z_idx in range(1, 6):
                z_name = f"Zone {z_idx}"
                z_clean = f"zone{z_idx}"
                z_xlsx = os.path.join(tmpdir, f"SHIPMENTS_TOMORROW_REPORT_{stamp}_{z_name.replace(' ', '_')}.xlsx")
                z_bills, z_weight = shipments_tomorrow.build_shipments_tomorrow_report(src, z_xlsx, target_label=z_name, start_time="00:00", end_time="06:00")
                z_caption = f"🚚 *SHIPMENTS TOMORROW REPORT ({z_name} | 00:00 - 06:00)*\n📦 Total Bills: `{z_bills}`\n⚖️ Total Weight: `{z_weight/1000:,.2f} kg`"

                # Send directly to requester
                try:
                    z_img = shipments_tomorrow.render_executive_summary_image(z_xlsx)
                    z_img.name = f"EXECUTIVE_SUMMARY_{z_name.replace(' ', '_')}.png"
                    await update.message.reply_photo(photo=z_img)
                except Exception as e_img:
                    log.warning("Could not render image: %s", e_img)

                with open(z_xlsx, "rb") as f_doc:
                    await update.message.reply_document(document=f_doc, filename=os.path.basename(z_xlsx), caption=z_caption, parse_mode="Markdown")

                # Forward to zone groups if not paused
                if not paused:
                    for gid, zkey in zone_fwd_map.items():
                        if zkey.lower() == z_clean:
                            try:
                                try:
                                    z_img_fwd = shipments_tomorrow.render_executive_summary_image(z_xlsx)
                                    z_img_fwd.name = f"EXECUTIVE_SUMMARY_{z_name.replace(' ', '_')}.png"
                                    await context.bot.send_photo(chat_id=int(gid), photo=z_img_fwd)
                                except Exception:
                                    pass
                                with open(z_xlsx, "rb") as f_doc_fwd:
                                    await context.bot.send_document(chat_id=int(gid), document=f_doc_fwd, filename=os.path.basename(z_xlsx), caption=z_caption, parse_mode="Markdown")
                                total_sent += 1
                            except Exception as e_fwd:
                                log.warning("Failed forwarding to %s: %s", gid, e_fwd)

            await status_msg.edit_text(f"✅ Generated and sent Shipments Tomorrow Reports for all 5 Zones!")
        else:
            # Single Target (e.g. Zone 1, SVAP001)
            out_xlsx = os.path.join(tmpdir, f"SHIPMENTS_TOMORROW_REPORT_{stamp}_{target_label.replace(' ', '_')}.xlsx")
            bills, weight = shipments_tomorrow.build_shipments_tomorrow_report(src, out_xlsx, target_label=target_label, start_time="00:00", end_time="06:00")

            # Send Executive Summary Image
            try:
                img_buf = shipments_tomorrow.render_executive_summary_image(out_xlsx)
                img_buf.name = f"EXECUTIVE_SUMMARY_{target_label.replace(' ', '_')}.png"
                await update.message.reply_photo(photo=img_buf)
            except Exception as e_img:
                log.warning("Could not render executive summary image: %s", e_img)

            caption = f"🚚 *SHIPMENTS TOMORROW REPORT ({target_label} | 00:00 - 06:00)*\n📦 Total Bills: `{bills}`\n⚖️ Total Weight: `{weight/1000:,.2f} kg`"
            with open(out_xlsx, "rb") as f_doc:
                await update.message.reply_document(document=f_doc, filename=os.path.basename(out_xlsx), caption=caption, parse_mode="Markdown")

            # Forward to group if registered and not paused
            if not paused:
                fwd_map = cfg.get("forward_mapping", {})
                for gid, handles in fwd_map.items():
                    if target_label.upper() in [h.upper() for h in handles]:
                        try:
                            with open(out_xlsx, "rb") as f_doc_fwd:
                                await context.bot.send_document(chat_id=int(gid), document=f_doc_fwd, filename=os.path.basename(out_xlsx), caption=caption, parse_mode="Markdown")
                        except Exception as e_fwd:
                            log.warning("Failed forwarding to group %s: %s", gid, e_fwd)

            await status_msg.edit_text(f"✅ Done! Report for *{target_label}* (00:00 - 06:00) generated successfully.", parse_mode="Markdown")

    except Exception as e:
        log.exception("Error in /tomorrow")
        await status_msg.edit_text(f"❌ Error generating report: `{e}`", parse_mode="Markdown")


# ── Main Entrypoint ───────────────────────────────────────────────────────────

def main():
    cfg = load_config()
    token = cfg["telegram"]["bot_token"]

    app = (
        Application.builder()
        .token(token)
        .connect_timeout(30.0)
        .read_timeout(30.0)
        .build()
    )

    app.add_handler(CommandHandler("start", cmd_start))
    app.add_handler(CommandHandler("stop", cmd_stop))
    app.add_handler(CommandHandler("pause", cmd_stop))
    app.add_handler(CommandHandler("resume", cmd_start))
    app.add_handler(CommandHandler("status", cmd_status))
    app.add_handler(CommandHandler("help", cmd_help))
    app.add_handler(CommandHandler("tomorrow", cmd_tomorrow))

    log.info("Shipments Tomorrow Bot started successfully. Listening for commands...")
    app.run_polling()


if __name__ == "__main__":
    main()
