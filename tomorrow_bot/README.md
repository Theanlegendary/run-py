# 🚚 Shipments Tomorrow Standalone Bot (00:00 - 06:00)

Specialized lightweight Telegram Bot for generating **Executive Shipments Tomorrow Reports** filtered between **00:00 (Midnight) and 06:00 (Morning)**.

---

## 🚀 How to Run

### On Windows:
Double-click `start.bat` or run:
```cmd
pip install -r requirements.txt
python bot.py
```

### On Linux / Cloud:
```bash
pip install -r requirements.txt
python bot.py
```

---

## 📱 Bot Commands

| Command | Description |
| :--- | :--- |
| **`/tomorrow`** | Generates report for **Zone 1 (00:00 - 06:00)** |
| **`/tomorrow zone 2`** | Generates report for **Zone 2** (or 3, 4, 5) |
| **`/tomorrow all`** | Generates and sends reports for **ALL 5 Zones** |
| **`/tomorrow <post_office>`** | Generates for specific post office (e.g. `/tomorrow SVAP001`) |
| **`/status`** | Checks server time, filter window, and bot status |
| **`/stop`** | Pauses group forwarding (safe testing mode) |
| **`/start`** | Resumes normal operations |
| **`/help`** | Shows available commands |

---

## 📁 Included Files
* `bot.py` — Dedicated standalone Telegram bot runner.
* `shipments_tomorrow.py` — Report generator with 00:00 - 06:00 time filtering & executive summary image rendering.
* `downloader.py` — Metfone API export downloader.
* `config.json` — Pre-configured with your Telegram bot token and Metfone API bearer token.
* `fonts/` — Bundled TrueType fonts (Arial, Calibri) for pixel-perfect executive summary rendering.
* `requirements.txt` — Dependencies list.
* `start.bat` — One-click launcher for Windows.
