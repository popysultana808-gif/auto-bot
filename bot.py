# ==============================================================================
# CyberNest Digital Mart — Full-Featured Telegram Auto-Shop Bot + Mini App
# Single-File Architecture: Python (python-telegram-bot v21+), SQLite / PostgreSQL,
# OpenPyXL, ReportLab, PyOTP, and Embedded Telegram WebApp (Mini App)
# ==============================================================================

import os
import sys
import io
import base64
import urllib.request
import urllib.parse
import asyncio
import time
import json
import logging
import sqlite3
import datetime
import secrets
import random
import html
import threading
from http.server import ThreadingHTTPServer, BaseHTTPRequestHandler
import hmac
import hashlib
import imaplib
import email as email_lib
from typing import Optional, List, Dict, Any, Tuple

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
import pyotp

# ReportLab for Official PDF Receipt Generation
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image as RLImage
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.graphics.shapes import Drawing
from reportlab.graphics.barcode import qr

from telegram import (
    Update,
    InlineKeyboardButton,
    InlineKeyboardMarkup,
    ReplyKeyboardMarkup,
    KeyboardButton,
    InputFile,
)
from telegram.constants import ParseMode
from telegram.ext import (
    ApplicationBuilder,
    CommandHandler,
    MessageHandler,
    CallbackQueryHandler,
    ChatMemberHandler,
    ContextTypes,
    filters,
)

CYBERNEST_LOGO_BASE64 = "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"

# ------------------------------------------------------------------------------
# 1. LOGGING & CORE CONFIGURATION
# ------------------------------------------------------------------------------
logging.basicConfig(
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
    level=logging.INFO,
)
logger = logging.getLogger("CyberNestBot")
logging.getLogger("httpx").setLevel(logging.WARNING)

BOT_TOKEN = os.getenv("BOT_TOKEN", "9198296559:AAFwotwWzDzNjIuqrl6iG0JY6NykvQOPl8Y")
SUPER_ADMIN_ID = int(os.getenv("SUPER_ADMIN_ID", "8919985167"))

WELCOME_GROUP_ID = os.getenv("WELCOME_GROUP_ID", "-5340748538")
WELCOME_GROUP_LINK = os.getenv("WELCOME_GROUP_LINK", "https://t.me/+jzGqcGVgiYRmYjJl")

TRANSACTION_CHANNEL_ID = os.getenv("TRANSACTION_CHANNEL_ID", "@Yourcybernestmart")
TRANSACTION_CHANNEL_LINK = os.getenv("TRANSACTION_CHANNEL_LINK", "https://t.me/Yourcybernestmart")

_AUTO_SALES_ACTIVE = True
_RECENTLY_WELCOMED = set()

CUSTOMER_SAMPLE_NAMES = [
    "7x_TALHA", "Arif Hasan", "Tanvir Ahmed", "Sakib Al Hasan", "Rahim Uddin",
    "Mehedi Hasan", "Nusrat Jahan", "Sadia Islam", "Farhana Akter",
    "Jannatul Ferdous", "Ayesha Siddiqua", "Rifat Hossain", "Sumaiya Kabir",
    "Shakil Khan", "Fahim Shahriar", "Mim Akter", "Tasnim Sultana",
    "Mahmudul Hasan", "Naimur Rahman", "Sabiha Sultana", "Habibur Rahman"
]

def get_welcome_group_id() -> str:
    return str(get_setting("welcome_group_id", WELCOME_GROUP_ID)).strip()

def get_welcome_group_link() -> str:
    return str(get_setting("welcome_group_link", WELCOME_GROUP_LINK)).strip()

def get_transaction_channel_id() -> str:
    return str(get_setting("transaction_channel_id", TRANSACTION_CHANNEL_ID)).strip()

def get_transaction_channel_link() -> str:
    return str(get_setting("transaction_channel_link", TRANSACTION_CHANNEL_LINK)).strip()

DB_FILE = "store_database.db"
_GLOBAL_BOT_APP = None
_GLOBAL_EVENT_LOOP = None

def get_bot(query=None):
    """Safely retrieves the active Bot instance."""
    if query is not None and hasattr(query, "get_bot"):
        try:
            b = query.get_bot()
            if b:
                return b
        except Exception:
            pass
    if _GLOBAL_BOT_APP:
        return _GLOBAL_BOT_APP.bot
    return None

def _http_post_json(url: str, data: dict, timeout: int = 15) -> dict:
    encoded = urllib.parse.urlencode(data).encode("utf-8")
    req = urllib.request.Request(url, data=encoded, headers={"User-Agent": "CyberNestBot/1.0"})
    with urllib.request.urlopen(req, timeout=timeout) as resp:
        return json.loads(resp.read().decode("utf-8"))

def _http_get_bytes(url: str, timeout: int = 10):
    req = urllib.request.Request(url, headers={"User-Agent": "CyberNestBot/1.0"})
    with urllib.request.urlopen(req, timeout=timeout) as resp:
        return resp.read(), resp.headers.get("Content-Type", "image/jpeg")

def add_notification(user_id: int, title: str, message: str):
    try:
        conn = get_db()
        c = conn.cursor()
        now_str = datetime.datetime.now().strftime("%Y-%m-%d %I:%M %p")
        c.execute(
            "INSERT INTO notifications (user_id, title, message, is_read, created_at) VALUES (?, ?, ?, 0, ?)",
            (user_id, title, message, now_str),
        )
        conn.commit()
        conn.close()
    except Exception as e:
        logger.warning(f"add_notification error: {e}")

def verify_tg_init_data(init_data: str) -> Optional[dict]:
    if not init_data:
        return None
    try:
        parsed = dict(urllib.parse.parse_qsl(init_data, keep_blank_values=True))
        hash_val = parsed.pop("hash", None)
        if not hash_val:
            return None
        data_check_string = "\n".join(f"{k}={v}" for k, v in sorted(parsed.items()))
        secret_key = hmac.new(b"WebAppData", BOT_TOKEN.encode("utf-8"), hashlib.sha256).digest()
        calc_hash = hmac.new(secret_key, data_check_string.encode("utf-8"), hashlib.sha256).hexdigest()
        if not hmac.compare_digest(calc_hash, hash_val):
            return None
        if "user" in parsed:
            return json.loads(parsed["user"])
    except Exception as e:
        logger.warning(f"initData verify failed: {e}")
    return None

def fetch_outlook_otp(account_str: str) -> dict:
    parts = [p.strip() for p in account_str.split("|") if p.strip()]
    if len(parts) < 4:
        return {"ok": False, "error": "Format must be: email|password|refresh_token|client_id"}
    mail_addr, _, refresh_token, client_id = parts[0], parts[1], parts[2], parts[3]
    try:
        token_url = "https://login.microsoftonline.com/common/oauth2/v2.0/token"
        tdata = _http_post_json(token_url, {
            "client_id": client_id,
            "refresh_token": refresh_token,
            "grant_type": "refresh_token",
            "scope": "https://outlook.office.com/IMAP.AccessAsUser.All offline_access"
        }, timeout=15)
        access_token = tdata.get("access_token")
        if not access_token:
            return {"ok": False, "error": tdata.get("error_description", "Token refresh failed")}
        auth_string = f"user={mail_addr}\x01auth=Bearer {access_token}\x01\x01"
        imap = imaplib.IMAP4_SSL("outlook.office365.com", 993)
        imap.authenticate("XOAUTH2", lambda x: auth_string.encode("utf-8"))
        imap.select("INBOX")
        typ, data = imap.search(None, "ALL")
        ids = data[0].split()
        if not ids:
            imap.logout()
            return {"ok": False, "error": "Inbox is empty (no emails found yet)."}
        latest_id = ids[-1]
        typ, msg_data = imap.fetch(latest_id, "(RFC822)")
        raw_email = msg_data[0][1]
        msg = email_lib.message_from_bytes(raw_email)
        subject = str(msg.get("Subject", "Inbox Message"))
        body = ""
        if msg.is_multipart():
            for part in msg.walk():
                ctype = part.get_content_type()
                if ctype in ("text/plain", "text/html"):
                    try:
                        body += part.get_payload(decode=True).decode(errors="ignore")
                    except Exception:
                        pass
        else:
            body = msg.get_payload(decode=True).decode(errors="ignore")
        imap.logout()
        import re as _re
        clean_txt = _re.sub(r"<[^>]+>", " ", body)
        m = _re.search(r"\b(\d{4,8})\b", clean_txt)
        code_found = m.group(1) if m else clean_txt.strip()[:80]
        return {"ok": True, "subject": subject, "code": code_found}
    except Exception as e:
        return {"ok": False, "error": str(e)}

# ------------------------------------------------------------------------------
# 2. EMBEDDED TELEGRAM MINI APP HTML (NO SEPARATE webapp.html NEEDED)
# ------------------------------------------------------------------------------
EMBEDDED_WEBAPP_HTML = r"""<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no" />
  <title>CyberNest Market</title>
  <script src="https://telegram.org/js/telegram-web-app.js"></script>
  <style>
    :root {
      --bg: #f6f7fb;
      --card: #ffffff;
      --text: #111827;
      --muted: #6b7280;
      --border: #e5e7eb;
      --accent: #ff6b00;
      --accent-soft: #fff3e8;
      --green: #10b981;
      --green-soft: #ecfdf5;
      --red: #ef4444;
      --red-soft: #fef2f2;
      --shadow: 0 6px 20px rgba(17, 24, 39, 0.06);
    }
    body.dark {
      --bg: #0b0f19;
      --card: #131a2a;
      --text: #f3f4f6;
      --muted: #9ca3af;
      --border: #1f2937;
      --accent: #ff7a1a;
      --accent-soft: rgba(255, 107, 0, 0.14);
      --green-soft: rgba(16, 185, 129, 0.14);
      --red-soft: rgba(239, 68, 68, 0.14);
      --shadow: 0 8px 24px rgba(0, 0, 0, 0.35);
    }
    * { box-sizing: border-box; -webkit-tap-highlight-color: transparent; }
    body {
      margin: 0; font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif;
      background: var(--bg); color: var(--text); padding-bottom: 92px;
      transition: background 0.25s, color 0.25s;
    }
    #splash {
      position: fixed; inset: 0; background: linear-gradient(135deg, #0b0f19 0%, #1a103d 100%);
      color: #fff; display: flex; flex-direction: column; align-items: center; justify-content: center;
      z-index: 9999; transition: opacity 0.45s ease;
    }
    #splash.hide { opacity: 0; pointer-events: none; }
    .loader-ring {
      width: 56px; height: 56px; border-radius: 50%;
      border: 4px solid rgba(255,255,255,0.15); border-top-color: #ff6b00;
      animation: spin 0.9s linear infinite; margin-bottom: 14px;
    }
    @keyframes spin { to { transform: rotate(360deg); } }

    .topbar {
      position: sticky; top: 0; z-index: 50; background: var(--card);
      border-bottom: 1px solid var(--border); padding: 10px 14px;
      display: flex; align-items: center; justify-content: space-between; gap: 8px;
    }
    .user-mini { display: flex; align-items: center; gap: 10px; min-width: 0; }
    .avatar {
      width: 38px; height: 38px; border-radius: 50%; background: linear-gradient(135deg, #ff6b00, #7c3aed);
      color: #fff; font-weight: 700; display: flex; align-items: center; justify-content: center; flex-shrink: 0;
    }
    .user-info { min-width: 0; }
    .user-info small { display: block; font-size: 11px; color: var(--muted); }
    .user-info b { display: block; font-size: 14px; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }
    .top-actions { display: flex; align-items: center; gap: 8px; }
    .bal-pill {
      background: linear-gradient(135deg, #ff7a1a, #ff5500); color: #fff;
      padding: 6px 12px; border-radius: 999px; font-weight: 700; font-size: 13px;
      display: flex; align-items: center; gap: 6px; cursor: pointer; border: none;
    }
    .icon-btn {
      width: 36px; height: 36px; border-radius: 50%; border: 1px solid var(--border);
      background: var(--bg); color: var(--text); display: flex; align-items: center; justify-content: center;
      cursor: pointer; position: relative; font-size: 16px;
    }
    .badge-dot {
      position: absolute; top: 5px; right: 6px; width: 8px; height: 8px;
      border-radius: 50%; background: var(--red); display: none;
    }

    .container { padding: 14px; max-width: 600px; margin: 0 auto; }

    .search-wrap {
      background: var(--card); border: 1.5px solid var(--accent); border-radius: 14px;
      padding: 8px 12px; display: flex; align-items: center; gap: 8px; box-shadow: var(--shadow);
      margin-bottom: 14px;
    }
    .search-wrap input {
      border: none; outline: none; background: transparent; color: var(--text);
      font-size: 14px; width: 100%;
    }

    .hero-banner {
      background: linear-gradient(135deg, #2e0249 0%, #570a57 50%, #ff6b00 100%);
      color: #fff; border-radius: 16px; padding: 16px; margin-bottom: 16px;
      box-shadow: var(--shadow); position: relative; overflow: hidden;
    }
    .hero-banner h3 { margin: 4px 0 6px; font-size: 18px; }
    .hero-banner p { margin: 0; font-size: 12px; opacity: 0.9; line-height: 1.4; }
    .hero-tag {
      display: inline-block; background: rgba(255,255,255,0.2); padding: 3px 9px;
      border-radius: 999px; font-size: 11px; font-weight: 600;
    }

    .cat-row {
      display: flex; gap: 10px; overflow-x: auto; padding-bottom: 6px; margin-bottom: 14px;
      scrollbar-width: none;
    }
    .cat-row::-webkit-scrollbar { display: none; }
    .cat-pill {
      flex-shrink: 0; background: var(--card); border: 1px solid var(--border);
      border-radius: 14px; padding: 10px 14px; font-size: 12px; font-weight: 600;
      cursor: pointer; display: flex; align-items: center; gap: 6px; color: var(--text);
    }
    .cat-pill.active {
      background: var(--accent); color: #fff; border-color: var(--accent);
    }

    .sec-title {
      display: flex; align-items: center; justify-content: space-between;
      margin: 12px 0 10px; font-weight: 700; font-size: 14px;
    }

    .top-card {
      background: var(--card); border: 1px solid var(--border); border-radius: 14px;
      padding: 10px 12px; display: flex; align-items: center; justify-content: space-between;
      gap: 10px; margin-bottom: 14px; box-shadow: var(--shadow); cursor: pointer;
    }
    .top-left { display: flex; align-items: center; gap: 10px; min-width: 0; }
    .prod-thumb {
      width: 48px; height: 48px; border-radius: 12px; object-fit: cover;
      background: var(--accent-soft); display: flex; align-items: center; justify-content: center;
      font-size: 22px; flex-shrink: 0;
    }

    .grid { display: grid; grid-template-columns: repeat(2, 1fr); gap: 12px; }
    .p-card {
      background: var(--card); border: 1px solid var(--border); border-radius: 16px;
      padding: 10px; box-shadow: var(--shadow); display: flex; flex-direction: column;
      justify-content: space-between; position: relative;
    }
    .p-img-box {
      height: 92px; border-radius: 12px; background: var(--bg);
      display: flex; align-items: center; justify-content: center; overflow: hidden;
      margin-bottom: 8px; position: relative;
    }
    .p-img-box img { width: 100%; height: 100%; object-fit: cover; }
    .fav-btn {
      position: absolute; top: 6px; right: 6px; width: 28px; height: 28px;
      border-radius: 50%; border: none; background: rgba(255,255,255,0.9);
      cursor: pointer; font-size: 14px; display: flex; align-items: center; justify-content: center;
    }
    .p-brand { font-size: 11px; color: var(--muted); font-weight: 600; }
    .p-name { font-size: 13px; font-weight: 700; margin: 2px 0 6px; }
    .p-price { font-size: 15px; font-weight: 800; color: var(--text); margin-bottom: 4px; }
    .p-meta { display: flex; align-items: center; justify-content: space-between; font-size: 11px; margin-bottom: 8px; }
    .st-in { color: var(--green); font-weight: 700; }
    .st-out { color: var(--red); font-weight: 700; }
    .btn-primary {
      width: 100%; border: none; border-radius: 10px; padding: 9px 12px;
      background: linear-gradient(135deg, #ff7a1a, #ff5500); color: #fff;
      font-weight: 700; font-size: 13px; cursor: pointer;
    }
    .btn-primary:disabled { opacity: 0.45; cursor: not-allowed; }
    .btn-outline {
      border: 1px solid var(--border); background: var(--card); color: var(--text);
      border-radius: 10px; padding: 8px 12px; font-weight: 600; font-size: 12px; cursor: pointer;
    }

    .bottom-nav {
      position: fixed; bottom: 0; left: 0; right: 0; height: 68px;
      background: var(--card); border-top: 1px solid var(--border);
      display: flex; align-items: center; justify-content: space-around; z-index: 60;
      padding-bottom: env(safe-area-inset-bottom);
    }
    .nav-item {
      display: flex; flex-direction: column; align-items: center; gap: 3px;
      font-size: 11px; color: var(--muted); cursor: pointer; border: none; background: transparent;
    }
    .nav-item.active { color: var(--accent); font-weight: 700; }
    .nav-center {
      width: 52px; height: 52px; border-radius: 16px; margin-top: -22px;
      background: linear-gradient(135deg, #ff7a1a, #ff5500); color: #fff;
      display: flex; align-items: center; justify-content: center; font-size: 22px;
      box-shadow: 0 8px 20px rgba(255, 107, 0, 0.35); border: 3px solid var(--bg);
    }

    .card {
      background: var(--card); border: 1px solid var(--border); border-radius: 16px;
      padding: 14px; margin-bottom: 12px; box-shadow: var(--shadow);
    }
    .input, textarea, select {
      width: 100%; padding: 10px 12px; border-radius: 10px;
      border: 1px solid var(--border); background: var(--bg); color: var(--text);
      font-size: 13px; outline: none; margin-top: 6px; margin-bottom: 10px;
    }
    .qty-box {
      display: flex; align-items: center; justify-content: space-between;
      background: var(--bg); border-radius: 12px; padding: 6px 12px; margin: 8px 0;
    }
    .qty-btn {
      width: 34px; height: 34px; border-radius: 10px; border: 1px solid var(--border);
      background: var(--card); color: var(--text); font-size: 18px; font-weight: 700; cursor: pointer;
    }
    .quick-qtys { display: flex; gap: 6px; flex-wrap: wrap; margin-bottom: 10px; }
    .q-pill {
      padding: 6px 12px; border-radius: 8px; border: 1px solid var(--border);
      background: var(--card); color: var(--text); font-size: 12px; font-weight: 600; cursor: pointer;
    }
    .q-pill.active { background: var(--accent); color: #fff; border-color: var(--accent); }

    .modal {
      position: fixed; inset: 0; background: rgba(0,0,0,0.6); z-index: 100;
      display: none; align-items: flex-end; justify-content: center;
    }
    .modal.open { display: flex; }
    .modal.center { align-items: center; padding: 16px; }
    .sheet {
      background: var(--card); width: 100%; max-width: 600px; max-height: 90vh;
      overflow-y: auto; border-radius: 22px 22px 0 0; padding: 16px;
      animation: slideUp 0.25s ease;
    }
    .modal.center .sheet { border-radius: 20px; max-width: 400px; }
    @keyframes slideUp { from { transform: translateY(30px); opacity: 0; } to { transform: translateY(0); opacity: 1; } }

    .row-between { display: flex; align-items: center; justify-content: space-between; gap: 8px; }
    .pill-tabs { display: flex; gap: 8px; margin-bottom: 12px; overflow-x: auto; }
    .tab-btn {
      flex: 1; white-space: nowrap; padding: 9px 12px; border-radius: 10px;
      border: 1px solid var(--border); background: var(--card); color: var(--text);
      font-size: 12px; font-weight: 700; cursor: pointer;
    }
    .tab-btn.active { background: var(--accent); color: #fff; border-color: var(--accent); }
    .copy-box {
      background: var(--bg); border: 1px dashed var(--accent); border-radius: 10px;
      padding: 10px; font-family: monospace; font-size: 13px; word-break: break-all;
      display: flex; align-items: center; justify-content: space-between; gap: 8px; margin: 6px 0;
    }
    .hidden { display: none !important; }
  </style>
</head>
<body>
  <div id="splash">
    <div class="loader-ring"></div>
    <h2 style="margin:0">CyberNest Market</h2>
    <small style="opacity:.75;margin-top:4px">Loading your digital store...</small>
  </div>

  <div class="topbar">
    <div class="user-mini">
      <div class="avatar" id="uAvatar">C</div>
      <div class="user-info">
        <small>Welcome back,</small>
        <b id="uName">Loading...</b>
      </div>
    </div>
    <div class="top-actions">
      <button class="bal-pill" onclick="switchPage('deposit')">
        <span>💳</span> <span id="uBal">৳0.00</span> <span>+</span>
      </button>
      <button class="icon-btn" onclick="openNotifs()">
        🔔 <span class="badge-dot" id="notifDot"></span>
      </button>
    </div>
  </div>

  <!-- PAGE: HOME -->
  <div id="page-home" class="container">
    <div class="search-wrap">
      <span>🔍</span>
      <input id="searchInput" type="text" placeholder="Search VPN, Proxy, Mail..." oninput="renderProducts()" />
    </div>

    <div class="hero-banner">
      <span class="hero-tag">⚡ 24/7 Automated Store</span>
      <h3 id="bannerTitle">CyberNest Digital Mart</h3>
      <p id="bannerSub">Fastest VPN, Proxy & Mail delivery with bKash, Nagad & Binance.</p>
    </div>

    <div class="cat-row" id="catRow"></div>
    <div id="topSellingWrap"></div>

    <div class="sec-title">
      <span id="prodSecTitle">🛍️ All Products</span>
      <small id="prodCount" style="color:var(--muted)"></small>
    </div>
    <div class="grid" id="prodGrid"></div>
  </div>

  <!-- PAGE: DEPOSIT -->
  <div id="page-deposit" class="container hidden">
    <div class="card">
      <h3 style="margin-top:0">💳 Manual Deposit</h3>
      <p style="font-size:12px;color:var(--muted)">Send Money to any number below, then submit Amount & TrxID.</p>
      <div id="depMethodsList"></div>
      <label style="font-size:12px;font-weight:700">Payment Method</label>
      <select id="depMethod" class="input">
        <option value="bKash">bKash</option>
        <option value="Nagad">Nagad</option>
        <option value="Binance">Binance (USDT)</option>
      </select>
      <label style="font-size:12px;font-weight:700">Amount (BDT)</label>
      <input id="depAmount" type="number" class="input" placeholder="e.g. 100" />
      <label style="font-size:12px;font-weight:700">Transaction ID (TrxID)</label>
      <input id="depTrx" type="text" class="input" placeholder="Paste TrxID here" />
      <button class="btn-primary" onclick="submitDeposit()">✅ Submit Deposit Request</button>
    </div>
  </div>

  <!-- PAGE: TOOLS -->
  <div id="page-tools" class="container hidden">
    <div class="pill-tabs">
      <button class="tab-btn active" id="tBtnOutlook" onclick="switchTool('outlook')">✉️ Outlook / Hotmail Code</button>
      <button class="tab-btn" id="tBtn2fa" onclick="switchTool('2fa')">🔑 2FA Live Code</button>
    </div>

    <div id="tool-outlook" class="card">
      <h3 style="margin-top:0">✉️ Outlook / Hotmail Code</h3>
      <p style="font-size:12px;color:var(--muted)">Paste <code>email|password|refresh_token|client_id</code> below:</p>
      <textarea id="outlookInput" rows="4" placeholder="email|password|refresh_token|client_id"></textarea>
      <div class="row-between" style="margin-bottom:10px">
        <button class="btn-outline" style="flex:1" onclick="pasteTo('outlookInput')">📋 Paste</button>
        <button class="btn-outline" style="flex:1" onclick="clearTool('outlook')">❌ Clear</button>
      </div>
      <button class="btn-primary" onclick="getOutlookCode()">🔄 Get Code</button>
      <div id="outlookResult" class="hidden" style="margin-top:12px"></div>
    </div>

    <div id="tool-2fa" class="card hidden">
      <h3 style="margin-top:0">🔑 2FA Authenticator</h3>
      <p style="font-size:12px;color:var(--muted)">Paste your 2FA Secret Key to generate live 6-digit OTP:</p>
      <input id="tfaInput" type="text" class="input" placeholder="e.g. JBSWY3DPEHPK3PXP" />
      <div class="row-between" style="margin-bottom:10px">
        <button class="btn-outline" style="flex:1" onclick="pasteTo('tfaInput')">📋 Paste</button>
        <button class="btn-outline" style="flex:1" onclick="clearTool('2fa')">❌ Clear</button>
      </div>
      <button class="btn-primary" onclick="get2faCode()">🔑 Generate 2FA Code</button>
      <div id="tfaResult" class="hidden" style="margin-top:12px"></div>
    </div>
  </div>

  <!-- PAGE: PROFILE -->
  <div id="page-profile" class="container hidden">
    <div class="card">
      <div class="row-between">
        <div>
          <b id="profName" style="font-size:16px">User</b>
          <div id="profMeta" style="font-size:12px;color:var(--muted)">ID: ---</div>
        </div>
        <div style="display:flex;gap:6px">
          <button class="btn-outline" id="themeBtn" onclick="toggleTheme()">🌙 Dark</button>
          <button class="btn-outline" id="currBtn" onclick="toggleCurrency()">🇧🇩 BDT</button>
        </div>
      </div>
      <div class="grid" style="grid-template-columns:repeat(3,1fr);margin-top:12px;text-align:center">
        <div class="card" style="margin:0;padding:10px"><b id="stOrders">0</b><div style="font-size:11px;color:var(--muted)">Orders</div></div>
        <div class="card" style="margin:0;padding:10px"><b id="stSpent">৳0</b><div style="font-size:11px;color:var(--muted)">Spent</div></div>
        <div class="card" style="margin:0;padding:10px"><b id="stRef">0</b><div style="font-size:11px;color:var(--muted)">Referrals</div></div>
      </div>
    </div>

    <div class="card">
      <div class="row-between">
        <b>🛒 Purchase History</b>
        <button class="btn-outline" onclick="loadHistory()">🔄 Refresh</button>
      </div>
      <div id="historyList" style="margin-top:10px;font-size:13px"></div>
    </div>

    <div class="card">
      <b>👥 Referral Link</b>
      <div class="copy-box" style="margin-top:8px">
        <span id="refLink" style="font-size:12px">Loading...</span>
        <button class="btn-outline" onclick="copyText(document.getElementById('refLink').innerText)">Copy</button>
      </div>
    </div>

    <div class="card row-between">
      <div><b>💬 Live Support</b><div style="font-size:12px;color:var(--muted)">Need help with an order?</div></div>
      <button class="btn-primary" style="width:auto" onclick="openSupport()">Message Admin</button>
    </div>
  </div>

  <!-- PAGE: ADMIN PANEL -->
  <div id="page-admin" class="container hidden">
    <div class="card">
      <h3 style="margin:0">👑 Web Admin Dashboard</h3>
      <small style="color:var(--muted)">Full live sync with Telegram Bot & Database</small>
      <div class="grid" style="grid-template-columns:repeat(3,1fr);margin-top:10px;text-align:center">
        <div class="card" style="margin:0;padding:8px"><b id="admUsers">0</b><div style="font-size:11px;color:var(--muted)">Users</div></div>
        <div class="card" style="margin:0;padding:8px"><b id="admPendDep">0</b><div style="font-size:11px;color:var(--muted)">Pend. Dep</div></div>
        <div class="card" style="margin:0;padding:8px"><b id="admPendOrd">0</b><div style="font-size:11px;color:var(--muted)">Pend. Ord</div></div>
      </div>
    </div>

    <div class="pill-tabs">
      <button class="tab-btn active" onclick="switchAdmTab('orders', this)">📦 Orders</button>
      <button class="tab-btn" onclick="switchAdmTab('deposits', this)">💳 Deposits</button>
      <button class="tab-btn" onclick="switchAdmTab('products', this)">🏷️ Products/Stock</button>
      <button class="tab-btn" onclick="switchAdmTab('popup', this)">🎁 Popup & Coupon</button>
    </div>

    <div id="adm-tab-orders"></div>
    <div id="adm-tab-deposits" class="hidden"></div>
    <div id="adm-tab-products" class="hidden"></div>
    <div id="adm-tab-popup" class="hidden"></div>
  </div>

  <!-- Bottom Navigation -->
  <div class="bottom-nav">
    <button class="nav-item active" id="nav-home" onclick="switchPage('home')"><span>🏠</span>Home</button>
    <button class="nav-item" id="nav-tools" onclick="switchPage('tools')"><span>🔐</span>Tools</button>
    <button class="nav-item" id="nav-deposit" onclick="switchPage('deposit')">
      <div class="nav-center">💳</div>Deposit
    </button>
    <button class="nav-item" id="nav-profile" onclick="switchPage('profile')"><span>👤</span>Profile</button>
    <button class="nav-item hidden" id="nav-admin" onclick="switchPage('admin')"><span>👑</span>Admin</button>
  </div>

  <!-- Product Purchase Modal -->
  <div class="modal" id="buyModal">
    <div class="sheet">
      <div class="row-between" style="margin-bottom:10px">
        <b id="mBrand" style="color:var(--accent);font-size:12px">CATEGORY</b>
        <button class="btn-outline" onclick="closeModal('buyModal')">✕</button>
      </div>
      <div id="mImgWrap" style="height:130px;border-radius:14px;overflow:hidden;background:var(--bg);display:flex;align-items:center;justify-content:center;margin-bottom:10px"></div>
      <h3 id="mName" style="margin:0 0 4px">Product Name</h3>
      <p id="mDesc" style="font-size:12px;color:var(--muted);margin:0 0 10px;line-height:1.45"></p>
      <div class="row-between" style="margin-bottom:8px">
        <span id="mPrice" style="font-size:20px;font-weight:800;color:var(--accent)">৳0</span>
        <span id="mStockBadge" style="font-size:12px;font-weight:700"></span>
      </div>

      <label style="font-size:12px;font-weight:700">QUANTITY</label>
      <div class="qty-box">
        <button class="qty-btn" onclick="changeQty(-1)">−</button>
        <input id="mQtyInput" type="number" value="1" min="1" style="width:70px;text-align:center;font-weight:800;font-size:16px;border:none;background:transparent;color:var(--text)" oninput="recalcOrder()" />
        <button class="qty-btn" onclick="changeQty(1)">+</button>
      </div>
      <div class="quick-qtys">
        <button class="q-pill" onclick="setQty(1)">1</button>
        <button class="q-pill" onclick="setQty(5)">5</button>
        <button class="q-pill" onclick="setQty(10)">10</button>
        <button class="q-pill" onclick="setQty(25)">25</button>
        <button class="q-pill" onclick="setQty(50)">50</button>
        <button class="q-pill" onclick="setQty(100)">100</button>
      </div>

      <div class="row-between">
        <input id="couponInput" class="input" style="margin:0;flex:1" placeholder="Promo / Coupon code (optional)" />
        <button class="btn-outline" onclick="applyCoupon()">Apply</button>
      </div>
      <div id="couponMsg" style="font-size:11px;margin:4px 0 8px"></div>

      <div class="card" style="background:var(--bg);box-shadow:none;margin-bottom:12px">
        <div class="row-between" style="font-size:12px;margin-bottom:4px"><span>Total Payable</span><b id="mTotal" style="font-size:16px;color:var(--accent)">৳0</b></div>
        <div class="row-between" style="font-size:12px"><span>Your Balance</span><b id="mYourBal">৳0</b></div>
      </div>
      <button class="btn-primary" id="mConfirmBtn" onclick="confirmOrder()">🛒 Confirm & Buy Now</button>
    </div>
  </div>

  <!-- Welcome Offer Popup -->
  <div class="modal center" id="offerModal">
    <div class="sheet" style="text-align:center">
      <div id="offerImgWrap" style="margin-bottom:10px"></div>
      <h3 id="offerTitle" style="margin:4px 0 8px">Special Offer!</h3>
      <p id="offerText" style="font-size:13px;color:var(--muted);line-height:1.5;margin-bottom:14px"></p>
      <button class="btn-primary" onclick="closeModal('offerModal')">🛍️ Start Shopping</button>
    </div>
  </div>

  <!-- Notifications Modal -->
  <div class="modal" id="notifModal">
    <div class="sheet">
      <div class="row-between" style="margin-bottom:10px">
        <h3 style="margin:0">🔔 Notifications</h3>
        <button class="btn-outline" onclick="closeModal('notifModal')">✕</button>
      </div>
      <div id="notifList"></div>
    </div>
  </div>

  <script>
    const tg = window.Telegram ? window.Telegram.WebApp : null;
    if (tg) { tg.ready(); tg.expand(); }
    const initData = tg ? tg.initData : "";

    let APP = {
      user: {}, categories: [], products: [], settings: {},
      currency: localStorage.getItem("cn_curr") || "BDT",
      theme: localStorage.getItem("cn_theme") || "light",
      selectedCat: "all",
      favorites: JSON.parse(localStorage.getItem("cn_favs") || "[]"),
      activeProduct: null,
      discountPct: 0,
      appliedCoupon: ""
    };

    if (APP.theme === "dark") document.body.classList.add("dark");

    async function api(path, body = null) {
      const opts = {
        method: body ? "POST" : "GET",
        headers: { "Content-Type": "application/json", "X-Telegram-Init-Data": initData }
      };
      if (body) opts.body = JSON.stringify(body);
      const res = await fetch(path, opts);
      return await res.json();
    }

    function fmtPrice(bdt) {
      if (APP.currency === "USDT") return "$" + (Number(bdt) / (APP.settings.usdt_rate || 125)).toFixed(3);
      return "৳" + (Number(bdt) < 1 ? Number(bdt).toFixed(2) : Number(bdt).toFixed(2).replace(/\.00$/, ""));
    }

    function prodImgUrl(p) {
      if (!p.image_file_id) return "";
      if (p.image_file_id.startsWith("http") || p.image_file_id.startsWith("data:")) return p.image_file_id;
      return "/api/tg_photo?file_id=" + encodeURIComponent(p.image_file_id);
    }

    function catIcon(catName) {
      const n = (catName || "").toLowerCase();
      if (n.includes("vpn")) return "🛡️";
      if (n.includes("proxy")) return "🌐";
      if (n.includes("recovery")) return "🔐";
      if (n.includes("mail")) return "✉️";
      return "📦";
    }

    async function boot() {
      try {
        const data = await api("/api/init");
        if (data.error) {
          document.getElementById("uName").innerText = "Open inside Telegram";
        } else {
          APP.user = data.user;
          APP.categories = data.categories;
          APP.products = data.products;
          APP.settings = data.settings;

          document.getElementById("uName").innerText = APP.user.full_name || "Customer";
          document.getElementById("uAvatar").innerText = (APP.user.full_name || "C")[0].toUpperCase();
          document.getElementById("uBal").innerText = fmtPrice(APP.user.balance_bdt || 0);
          document.getElementById("profName").innerText = APP.user.full_name || "Customer";
          document.getElementById("profMeta").innerText = `ID: ${APP.user.user_id} • @${APP.user.username || "user"}`;
          document.getElementById("stOrders").innerText = APP.user.total_orders || 0;
          document.getElementById("stSpent").innerText = fmtPrice(APP.user.total_spent || 0);
          document.getElementById("stRef").innerText = APP.user.referrals || 0;
          document.getElementById("refLink").innerText = `https://t.me/${APP.settings.bot_username || "bot"}?start=${APP.user.user_id}`;

          document.getElementById("depMethodsList").innerHTML = `
            <div class="copy-box"><span>📱 bKash: <b>${APP.settings.bkash_number || ""}</b></span><button class="btn-outline" onclick="copyText('${APP.settings.bkash_number || ""}')">Copy</button></div>
            <div class="copy-box"><span>📱 Nagad: <b>${APP.settings.nagad_number || ""}</b></span><button class="btn-outline" onclick="copyText('${APP.settings.nagad_number || ""}')">Copy</button></div>
            <div class="copy-box"><span>💎 Binance: <b>${APP.settings.binance_id || ""}</b></span><button class="btn-outline" onclick="copyText('${APP.settings.binance_id || ""}')">Copy</button></div>
          `;

          if (APP.user.is_admin) {
            document.getElementById("nav-admin").classList.remove("hidden");
          }
          if (data.unread_notifs > 0) document.getElementById("notifDot").style.display = "block";

          renderCategories();
          renderTopSelling();
          renderProducts();

          if (APP.settings.popup_enabled === "1" && !sessionStorage.getItem("popup_shown")) {
            sessionStorage.setItem("popup_shown", "1");
            document.getElementById("offerTitle").innerText = APP.settings.popup_title || "Special Offer!";
            document.getElementById("offerText").innerText = APP.settings.popup_text || "";
            const iw = document.getElementById("offerImgWrap");
            if (APP.settings.popup_image) {
              const src = APP.settings.popup_image.startsWith("http") || APP.settings.popup_image.startsWith("data:")
                ? APP.settings.popup_image
                : "/api/tg_photo?file_id=" + encodeURIComponent(APP.settings.popup_image);
              iw.innerHTML = `<img src="${src}" style="width:100%;max-height:170px;object-fit:cover;border-radius:12px" />`;
            } else iw.innerHTML = "";
            document.getElementById("offerModal").classList.add("open");
          }
        }
      } catch (e) { console.error(e); }
      setTimeout(() => document.getElementById("splash").classList.add("hide"), 350);
    }

    function renderCategories() {
      const row = document.getElementById("catRow");
      let html = `<button class="cat-pill ${APP.selectedCat === 'all' ? 'active' : ''}" onclick="selectCat('all')">📦 All Products</button>`;
      APP.categories.forEach(c => {
        html += `<button class="cat-pill ${APP.selectedCat == c.id ? 'active' : ''}" onclick="selectCat(${c.id})">${c.name}</button>`;
      });
      row.innerHTML = html;
    }

    function selectCat(cid) {
      APP.selectedCat = cid;
      renderCategories();
      renderProducts();
    }

    function renderTopSelling() {
      const wrap = document.getElementById("topSellingWrap");
      if (!APP.products.length) return;
      const top = [...APP.products].sort((a, b) => (b.sold_count || 0) - (a.sold_count || 0))[0];
      if (!top) return;
      const img = prodImgUrl(top);
      wrap.innerHTML = `
        <div class="sec-title"><span>🔥 TOP SELLING</span></div>
        <div class="top-card" onclick="openBuy(${top.id})">
          <div class="top-left">
            <div class="prod-thumb">${img ? `<img src="${img}" style="width:100%;height:100%;object-fit:cover;border-radius:12px">` : catIcon(top.cat_name)}</div>
            <div>
              <b style="font-size:13px">${top.sub_name} — ${top.name}</b>
              <div style="font-size:12px;color:var(--accent);font-weight:700">${fmtPrice(top.price_bdt)} <span style="color:var(--muted);font-weight:400">• ${top.sold_count || 0} sold</span></div>
            </div>
          </div>
          <button class="btn-primary" style="width:auto;padding:8px 14px">🛒 Buy</button>
        </div>
      `;
    }

    function renderProducts() {
      const q = (document.getElementById("searchInput").value || "").toLowerCase().trim();
      const filtered = APP.products.filter(p => {
        if (APP.selectedCat !== "all" && p.category_id != APP.selectedCat) return false;
        if (q && !(p.name + " " + p.sub_name + " " + p.cat_name).toLowerCase().includes(q)) return false;
        return true;
      });
      document.getElementById("prodCount").innerText = `${filtered.length} items`;
      const grid = document.getElementById("prodGrid");
      if (!filtered.length) {
        grid.innerHTML = `<div style="grid-column:span 2;text-align:center;color:var(--muted);padding:30px">No matching products found.</div>`;
        return;
      }
      grid.innerHTML = filtered.map(p => {
        const img = prodImgUrl(p);
        const isVpnOrProxy = p.cat_name.toLowerCase().includes("vpn") || p.cat_name.toLowerCase().includes("proxy");
        const available = p.is_active && (p.in_stock > 0 || isVpnOrProxy);
        const stockLabel = !p.is_active
          ? `<span class="st-out">● Unavailable</span>`
          : (p.in_stock > 0 ? `<span class="st-in">● In stock (${p.in_stock})</span>` : (isVpnOrProxy ? `<span class="st-in">● Available</span>` : `<span class="st-out">● Out of stock</span>`));
        const isFav = APP.favorites.includes(p.id);
        return `
          <div class="p-card">
            <div>
              <div class="p-img-box">
                ${img ? `<img src="${img}" />` : `<span style="font-size:34px">${catIcon(p.cat_name)}</span>`}
                <button class="fav-btn" onclick="toggleFav(event, ${p.id})">${isFav ? "❤️" : "🤍"}</button>
              </div>
              <div class="p-brand">${p.sub_name}</div>
              <div class="p-name">${p.name}</div>
              <div class="p-price">${fmtPrice(p.price_bdt)}</div>
              <div class="p-meta">${stockLabel} <span>🔥 ${p.sold_count || 0}</span></div>
            </div>
            <button class="btn-primary" ${available ? "" : "disabled"} onclick="openBuy(${p.id})">${available ? "🛒 Buy now" : "Out of stock"}</button>
          </div>
        `;
      }).join("");
    }

    function toggleFav(e, pid) {
      e.stopPropagation();
      if (APP.favorites.includes(pid)) APP.favorites = APP.favorites.filter(x => x !== pid);
      else APP.favorites.push(pid);
      localStorage.setItem("cn_favs", JSON.stringify(APP.favorites));
      renderProducts();
    }

    function openBuy(pid) {
      const p = APP.products.find(x => x.id === pid);
      if (!p) return;
      APP.activeProduct = p;
      APP.discountPct = 0;
      APP.appliedCoupon = "";
      document.getElementById("couponInput").value = "";
      document.getElementById("couponMsg").innerText = "";
      document.getElementById("mBrand").innerText = `${p.cat_name} • ${p.sub_name}`;
      document.getElementById("mName").innerText = `${p.sub_name} — ${p.name}`;
      document.getElementById("mDesc").innerText = p.description || "Fast & verified digital service with instant or rapid delivery.";
      const img = prodImgUrl(p);
      document.getElementById("mImgWrap").innerHTML = img ? `<img src="${img}" style="width:100%;height:100%;object-fit:cover">` : `<span style="font-size:48px">${catIcon(p.cat_name)}</span>`;
      document.getElementById("mPrice").innerText = fmtPrice(p.price_bdt);
      const isVpnOrProxy = p.cat_name.toLowerCase().includes("vpn") || p.cat_name.toLowerCase().includes("proxy");
      document.getElementById("mStockBadge").innerHTML = p.in_stock > 0
        ? `<span class="st-in">● ${p.in_stock} in stock</span>`
        : (isVpnOrProxy ? `<span class="st-in">● Manual / On-Demand</span>` : `<span class="st-out">● Out of stock</span>`);
      document.getElementById("mQtyInput").value = 1;
      recalcOrder();
      document.getElementById("buyModal").classList.add("open");
    }

    function changeQty(d) {
      const inp = document.getElementById("mQtyInput");
      inp.value = Math.max(1, (parseInt(inp.value) || 1) + d);
      recalcOrder();
    }
    function setQty(q) {
      document.getElementById("mQtyInput").value = q;
      recalcOrder();
    }
    function recalcOrder() {
      if (!APP.activeProduct) return;
      const qty = Math.max(1, parseInt(document.getElementById("mQtyInput").value) || 1);
      let total = APP.activeProduct.price_bdt * qty;
      if (APP.discountPct > 0) total = total * (1 - APP.discountPct / 100);
      document.getElementById("mTotal").innerText = fmtPrice(total);
      const bal = APP.user.balance_bdt || 0;
      document.getElementById("mYourBal").innerHTML = fmtPrice(bal) + (bal < total ? ` <span style="color:var(--red)">(Insufficient)</span>` : "");
    }

    async function applyCoupon() {
      const code = document.getElementById("couponInput").value.trim();
      if (!code) return;
      const res = await api("/api/coupon/check", { code });
      const msg = document.getElementById("couponMsg");
      if (res.ok) {
        APP.discountPct = res.discount_pct;
        APP.appliedCoupon = code;
        msg.innerHTML = `<span style="color:var(--green)">✅ ${res.discount_pct}% discount applied!</span>`;
        recalcOrder();
      } else {
        APP.discountPct = 0; APP.appliedCoupon = "";
        msg.innerHTML = `<span style="color:var(--red)">❌ ${res.error || "Invalid coupon"}</span>`;
        recalcOrder();
      }
    }

    async function confirmOrder() {
      if (!APP.activeProduct) return;
      const qty = Math.max(1, parseInt(document.getElementById("mQtyInput").value) || 1);
      const btn = document.getElementById("mConfirmBtn");
      btn.disabled = true; btn.innerText = "⏳ Processing...";
      const res = await api("/api/order", { product_id: APP.activeProduct.id, quantity: qty, coupon: APP.appliedCoupon });
      btn.disabled = false; btn.innerText = "🛒 Confirm & Buy Now";
      if (res.ok) {
        closeModal("buyModal");
        if (tg && tg.showAlert) tg.showAlert(res.message);
        boot();
      } else {
        if (tg && tg.showAlert) tg.showAlert("❌ " + (res.error || "Order failed"));
      }
    }

    async function submitDeposit() {
      const method = document.getElementById("depMethod").value;
      const amount = parseFloat(document.getElementById("depAmount").value || 0);
      const trx_id = document.getElementById("depTrx").value.trim();
      if (!amount || !trx_id) return;
      const res = await api("/api/deposit", { method, amount, trx_id });
      if (res.ok) {
        document.getElementById("depAmount").value = "";
        document.getElementById("depTrx").value = "";
        if (tg && tg.showAlert) tg.showAlert("✅ Deposit request submitted! Admin will approve shortly.");
      }
    }

    function switchTool(t) {
      document.getElementById("tool-outlook").classList.toggle("hidden", t !== "outlook");
      document.getElementById("tool-2fa").classList.toggle("hidden", t !== "2fa");
      document.getElementById("tBtnOutlook").classList.toggle("active", t === "outlook");
      document.getElementById("tBtn2fa").classList.toggle("active", t === "2fa");
    }

    async function pasteTo(id) {
      try { document.getElementById(id).value = await navigator.clipboard.readText(); }
      catch (e) {}
    }
    function clearTool(t) {
      if (t === "outlook") {
        document.getElementById("outlookInput").value = "";
        document.getElementById("outlookResult").classList.add("hidden");
      } else {
        document.getElementById("tfaInput").value = "";
        document.getElementById("tfaResult").classList.add("hidden");
      }
    }
    async function getOutlookCode() {
      const raw = document.getElementById("outlookInput").value.trim();
      if (!raw) return;
      const box = document.getElementById("outlookResult");
      box.classList.remove("hidden");
      box.innerHTML = "⏳ Checking inbox...";
      const res = await api("/api/tools/outlook", { account: raw });
      if (res.ok) {
        box.innerHTML = `
          <div style="font-size:12px;color:var(--muted)">Subject: ${res.subject || "Inbox"}</div>
          <div class="copy-box"><b>${res.code}</b><button class="btn-outline" onclick="copyText('${res.code}')">📋 Copy Code</button></div>
        `;
      } else {
        box.innerHTML = `<span style="color:var(--red)">❌ ${res.error}</span>`;
      }
    }
    async function get2faCode() {
      const secret = document.getElementById("tfaInput").value.trim();
      if (!secret) return;
      const box = document.getElementById("tfaResult");
      box.classList.remove("hidden");
      const res = await api("/api/tools/2fa", { secret });
      if (res.ok) {
        box.innerHTML = `
          <div class="copy-box"><b style="font-size:20px;letter-spacing:2px">${res.code}</b><button class="btn-outline" onclick="copyText('${res.code}')">📋 Copy Code</button></div>
          <small style="color:var(--muted)">⏳ Valid for ${res.remaining}s</small>
        `;
      } else {
        box.innerHTML = `<span style="color:var(--red)">❌ ${res.error}</span>`;
      }
    }

    async function loadHistory() {
      const res = await api("/api/orders/history");
      const list = document.getElementById("historyList");
      if (!res.orders || !res.orders.length) {
        list.innerHTML = "<span style='color:var(--muted)'>No orders yet.</span>"; return;
      }
      list.innerHTML = res.orders.map(o => `
        <div style="border-bottom:1px solid var(--border);padding:8px 0">
          <div class="row-between"><b>${o.product_name} (x${o.quantity})</b><span>${fmtPrice(o.total_bdt)}</span></div>
          <div class="row-between" style="font-size:11px;color:var(--muted);margin-top:2px">
            <span>#${o.id} • ${o.created_at}</span>
            <span style="color:${o.status === 'completed' ? 'var(--green)' : 'var(--accent)'};font-weight:700">${o.status.toUpperCase()}</span>
          </div>
          ${o.delivered_data ? `<div class="copy-box" style="margin-top:6px"><span style="font-size:11px">${o.delivered_data}</span><button class="btn-outline" onclick="copyText(\`${o.delivered_data.replace(/`/g, '')}\`)">Copy</button></div>` : ""}
        </div>
      `).join("");
    }

    async function openNotifs() {
      document.getElementById("notifModal").classList.add("open");
      document.getElementById("notifDot").style.display = "none";
      const res = await api("/api/notifications");
      const box = document.getElementById("notifList");
      if (!res.items || !res.items.length) {
        box.innerHTML = "<p style='color:var(--muted);text-align:center'>No notifications yet.</p>"; return;
      }
      box.innerHTML = res.items.map(n => `
        <div class="card" style="padding:10px;margin-bottom:8px;box-shadow:none">
          <b style="font-size:13px">${n.title}</b>
          <div style="font-size:12px;color:var(--muted);margin:3px 0">${n.message}</div>
          <small style="font-size:10px;color:var(--muted)">${n.created_at}</small>
        </div>
      `).join("");
    }

    async function loadAdmin() {
      const res = await api("/api/admin/overview");
      if (!res.ok) return;
      document.getElementById("admUsers").innerText = res.stats.users;
      document.getElementById("admPendDep").innerText = res.stats.pending_deposits;
      document.getElementById("admPendOrd").innerText = res.stats.pending_orders;

      document.getElementById("adm-tab-orders").innerHTML = res.pending_orders.length
        ? res.pending_orders.map(o => `
          <div class="card">
            <b>Order #${o.id} — ${o.product_name} (x${o.quantity})</b>
            <div style="font-size:12px;color:var(--muted)">User: ${o.user_id} • Paid: ${fmtPrice(o.total_bdt)}</div>
            <input id="deliv_in_${o.id}" class="input" placeholder="Enter email:pass OR host:port:user:pass" />
            <div class="row-between">
              <button class="btn-primary" style="flex:1" onclick="admDeliver(${o.id})">🚀 Deliver</button>
              <button class="btn-outline" style="color:var(--red)" onclick="admRefund(${o.id})">❌ Refund</button>
            </div>
          </div>
        `).join("")
        : `<div class="card">No pending manual orders.</div>`;

      document.getElementById("adm-tab-deposits").innerHTML = res.pending_deposits.length
        ? res.pending_deposits.map(d => `
          <div class="card">
            <b>Deposit #${d.id} — ${d.method} (${fmtPrice(d.amount)})</b>
            <div style="font-size:12px;color:var(--muted)">User: ${d.user_id} • TrxID: <code>${d.trx_id}</code></div>
            <div class="row-between" style="margin-top:8px">
              <button class="btn-primary" style="flex:1" onclick="admDepAction(${d.id}, 'approve')">✅ Approve</button>
              <button class="btn-outline" style="color:var(--red)" onclick="admDepAction(${d.id}, 'reject')">❌ Reject</button>
            </div>
          </div>
        `).join("")
        : `<div class="card">No pending deposits.</div>`;

      document.getElementById("adm-tab-products").innerHTML = APP.products.map(p => `
        <div class="card">
          <div class="row-between">
            <b>${p.sub_name} — ${p.name}</b>
            <button class="btn-outline" onclick="admToggleProd(${p.id})">${p.is_active ? "🟢 Active" : "🔴 Inactive"}</button>
          </div>
          <div class="row-between" style="margin-top:6px">
            <input id="adm_p_${p.id}" type="number" class="input" style="margin:0;width:110px" value="${p.price_bdt}" />
            <button class="btn-outline" onclick="admSavePrice(${p.id})">💰 Save Price</button>
          </div>
          <input id="adm_img_${p.id}" class="input" placeholder="Paste Photo URL or Telegram file_id" value="${p.image_file_id || ""}" />
          <textarea id="adm_stk_${p.id}" rows="2" placeholder="Paste stock lines (1 per line) to add"></textarea>
          <div class="row-between">
            <button class="btn-outline" onclick="admSavePhoto(${p.id})">🖼️ Save Photo</button>
            <button class="btn-primary" style="width:auto" onclick="admAddStock(${p.id})">📥 Add Stock</button>
          </div>
        </div>
      `).join("");

      document.getElementById("adm-tab-popup").innerHTML = `
        <div class="card">
          <h4 style="margin:0 0 8px">🎁 Welcome Offer Popup</h4>
          <select id="popEn" class="input">
            <option value="1" ${APP.settings.popup_enabled === "1" ? "selected" : ""}>🟢 Enabled</option>
            <option value="0" ${APP.settings.popup_enabled !== "1" ? "selected" : ""}>🔴 Disabled</option>
          </select>
          <input id="popTitle" class="input" placeholder="Popup Title" value="${APP.settings.popup_title || ""}" />
          <textarea id="popText" rows="3" placeholder="Popup Message">${APP.settings.popup_text || ""}</textarea>
          <input id="popImg" class="input" placeholder="Popup Image URL or file_id" value="${APP.settings.popup_image || ""}" />
          <button class="btn-primary" onclick="admSavePopup()">💾 Save Popup</button>
        </div>
        <div class="card">
          <h4 style="margin:0 0 8px">🎟️ Create Coupon Code</h4>
          <input id="newCoupCode" class="input" placeholder="Code (e.g. CYBER10)" />
          <input id="newCoupPct" type="number" class="input" placeholder="Discount % (e.g. 10)" />
          <button class="btn-primary" onclick="admCreateCoupon()">➕ Add Coupon</button>
        </div>
      `;
    }

    function switchAdmTab(t, btn) {
      ["orders", "deposits", "products", "popup"].forEach(x => {
        document.getElementById("adm-tab-" + x).classList.toggle("hidden", x !== t);
      });
      document.querySelectorAll("#page-admin .tab-btn").forEach(b => b.classList.remove("active"));
      btn.classList.add("active");
    }
    async function admDeliver(oid) {
      const val = document.getElementById("deliv_in_" + oid).value.trim();
      if (!val) return;
      const r = await api("/api/admin/deliver_order", { order_id: oid, data: val });
      if (r.ok) { loadAdmin(); }
    }
    async function admRefund(oid) {
      const r = await api("/api/admin/refund_order", { order_id: oid });
      if (r.ok) { loadAdmin(); }
    }
    async function admDepAction(did, action) {
      const r = await api("/api/admin/deposit_action", { deposit_id: did, action });
      if (r.ok) { loadAdmin(); }
    }
    async function admToggleProd(pid) {
      await api("/api/admin/product_update", { product_id: pid, toggle_active: true });
      await boot(); loadAdmin();
    }
    async function admSavePrice(pid) {
      const price = parseFloat(document.getElementById("adm_p_" + pid).value);
      await api("/api/admin/product_update", { product_id: pid, price_bdt: price });
      await boot(); loadAdmin();
    }
    async function admSavePhoto(pid) {
      const image_file_id = document.getElementById("adm_img_" + pid).value.trim();
      await api("/api/admin/product_update", { product_id: pid, image_file_id });
      await boot(); loadAdmin();
    }
    async function admAddStock(pid) {
      const lines = document.getElementById("adm_stk_" + pid).value.trim();
      if (!lines) return;
      await api("/api/admin/add_stock", { product_id: pid, lines });
      document.getElementById("adm_stk_" + pid).value = "";
      await boot(); loadAdmin();
    }
    async function admSavePopup() {
      await api("/api/admin/popup", {
        enabled: document.getElementById("popEn").value,
        title: document.getElementById("popTitle").value,
        text: document.getElementById("popText").value,
        image: document.getElementById("popImg").value
      });
      await boot();
    }
    async function admCreateCoupon() {
      const code = document.getElementById("newCoupCode").value.trim().toUpperCase();
      const pct = parseFloat(document.getElementById("newCoupPct").value || 0);
      if (!code || !pct) return;
      await api("/api/admin/coupon", { code, discount_pct: pct });
    }

    function switchPage(p) {
      ["home", "deposit", "tools", "profile", "admin"].forEach(x => {
        document.getElementById("page-" + x).classList.toggle("hidden", x !== p);
        const nav = document.getElementById("nav-" + x);
        if (nav) nav.classList.toggle("active", x === p);
      });
      if (p === "profile") loadHistory();
      if (p === "admin") loadAdmin();
    }
    function closeModal(id) { document.getElementById(id).classList.remove("open"); }
    function toggleTheme() {
      document.body.classList.toggle("dark");
      APP.theme = document.body.classList.contains("dark") ? "dark" : "light";
      localStorage.setItem("cn_theme", APP.theme);
      document.getElementById("themeBtn").innerText = APP.theme === "dark" ? "☀️ Light" : "🌙 Dark";
    }
    function toggleCurrency() {
      APP.currency = APP.currency === "BDT" ? "USDT" : "BDT";
      localStorage.setItem("cn_curr", APP.currency);
      document.getElementById("currBtn").innerText = APP.currency === "BDT" ? "🇧🇩 BDT" : "💲 USDT";
      document.getElementById("uBal").innerText = fmtPrice(APP.user.balance_bdt || 0);
      renderProducts(); renderTopSelling();
    }
    function copyText(txt) {
      navigator.clipboard.writeText(txt).catch(() => {});
    }
    function openSupport() {
      const u = (APP.settings.support_username || "@Talha_juba098").replace("@", "");
      if (tg) tg.openTelegramLink("https://t.me/" + u);
      else window.open("https://t.me/" + u, "_blank");
    }

    boot();
  </script>
</body>
</html>"""

# ------------------------------------------------------------------------------
# 3. RENDER WEB SERVICE & MINI APP API SERVER
# ------------------------------------------------------------------------------
class HealthCheckHandler(BaseHTTPRequestHandler):
    def _send_json(self, obj, code_num=200):
        body = json.dumps(obj, ensure_ascii=False, default=str).encode("utf-8")
        self.send_response(code_num)
        self.send_header("Content-type", "application/json; charset=utf-8")
        self.send_header("Access-Control-Allow-Origin", "*")
        self.end_headers()
        self.wfile.write(body)

    def _get_auth_user(self):
        init_data = self.headers.get("X-Telegram-Init-Data", "")
        u = verify_tg_init_data(init_data)
        if u and "id" in u:
            return int(u["id"]), u
        return None, None

    def do_HEAD(self):
        self.send_response(200)
        self.send_header("Content-type", "text/plain; charset=utf-8")
        self.end_headers()

    def do_GET(self):
        parsed = urllib.parse.urlparse(self.path)
        path = parsed.path

        if path == "/app" or path == "/app/":
            content = EMBEDDED_WEBAPP_HTML.encode("utf-8")
            self.send_response(200)
            self.send_header("Content-type", "text/html; charset=utf-8")
            self.end_headers()
            self.wfile.write(content)
            return

        if path == "/api/tg_photo":
            qs = urllib.parse.parse_qs(parsed.query)
            fid = qs.get("file_id", [""])[0]
            if not fid:
                self.send_response(404)
                self.end_headers()
                return
            try:
                raw_info, _ = _http_get_bytes(f"https://api.telegram.org/bot{BOT_TOKEN}/getFile?file_id={fid}", timeout=8)
                r1 = json.loads(raw_info.decode("utf-8"))
                fpath = r1["result"]["file_path"]
                img_bytes, ctype = _http_get_bytes(f"https://api.telegram.org/file/bot{BOT_TOKEN}/{fpath}", timeout=10)
                self.send_response(200)
                self.send_header("Content-type", ctype)
                self.send_header("Cache-Control", "public, max-age=86400")
                self.end_headers()
                self.wfile.write(img_bytes)
            except Exception:
                self.send_response(404)
                self.end_headers()
            return

        if path == "/api/init":
            uid, tg_u = self._get_auth_user()
            if not uid:
                return self._send_json({"error": "Unauthorized"}, 401)
            db_u = get_or_create_user(uid, tg_u.get("first_name", "Customer"), tg_u.get("username", ""))
            conn = get_db()
            c = conn.cursor()
            c.execute("SELECT id, name FROM categories ORDER BY position ASC, id ASC")
            cats = [dict(r) for r in c.fetchall()]
            c.execute("""
                SELECT p.id, p.name, p.price_bdt, p.price_usdt, p.image_file_id, p.description,
                       COALESCE(p.is_active, 1) as is_active, s.id as sub_id, s.name as sub_name,
                       s.category_id, cat.name as cat_name,
                       (SELECT COUNT(*) FROM stock st WHERE st.product_id = p.id AND st.is_sold = 0) as in_stock,
                       (SELECT COALESCE(SUM(o.quantity), 0) FROM orders o WHERE o.product_id = p.id AND o.status != 'refunded') as sold_count
                FROM products p
                JOIN subcategories s ON p.subcategory_id = s.id
                JOIN categories cat ON s.category_id = cat.id
                ORDER BY cat.position ASC, s.position ASC, p.id ASC
            """)
            prods = [dict(r) for r in c.fetchall()]
            c.execute("SELECT COUNT(*) as cnt, COALESCE(SUM(total_bdt), 0) as spent FROM orders WHERE user_id = ? AND status != 'refunded'", (uid,))
            ostat = dict(c.fetchone())
            c.execute("SELECT COUNT(*) as cnt FROM notifications WHERE user_id = ? AND is_read = 0", (uid,))
            unread_n = c.fetchone()["cnt"]
            c.execute("SELECT key, value FROM settings")
            sdict = {r["key"]: r["value"] for r in c.fetchall()}
            conn.close()
            bot_uname = ""
            try:
                if _GLOBAL_BOT_APP and _GLOBAL_BOT_APP.bot:
                    bot_uname = _GLOBAL_BOT_APP.bot.username or ""
            except Exception:
                pass
            sdict["bot_username"] = bot_uname
            return self._send_json({
                "user": {
                    "user_id": uid,
                    "full_name": db_u.get("name", "Customer"),
                    "username": db_u.get("username", ""),
                    "balance_bdt": db_u.get("balance_bdt", 0.0),
                    "is_admin": is_admin(uid),
                    "total_orders": ostat["cnt"],
                    "total_spent": ostat["spent"],
                    "referrals": 0,
                },
                "categories": cats,
                "products": prods,
                "settings": sdict,
                "unread_notifs": unread_n,
            })

        if path == "/api/orders/history":
            uid, _ = self._get_auth_user()
            if not uid:
                return self._send_json({"error": "Unauthorized"}, 401)
            conn = get_db()
            c = conn.cursor()
            c.execute("SELECT id, receipt_no, product_name, quantity, total_bdt, status, delivered_data, order_date as created_at FROM orders WHERE user_id = ? ORDER BY id DESC LIMIT 30", (uid,))
            rows = [dict(r) for r in c.fetchall()]
            conn.close()
            return self._send_json({"orders": rows})

        if path == "/api/notifications":
            uid, _ = self._get_auth_user()
            if not uid:
                return self._send_json({"error": "Unauthorized"}, 401)
            conn = get_db()
            c = conn.cursor()
            c.execute("SELECT id, title, message, created_at FROM notifications WHERE user_id = ? ORDER BY id DESC LIMIT 25", (uid,))
            items = [dict(r) for r in c.fetchall()]
            c.execute("UPDATE notifications SET is_read = 1 WHERE user_id = ?", (uid,))
            conn.commit()
            conn.close()
            return self._send_json({"items": items})

        if path == "/api/admin/overview":
            uid, _ = self._get_auth_user()
            if not uid or not is_admin(uid):
                return self._send_json({"error": "Forbidden"}, 403)
            conn = get_db()
            c = conn.cursor()
            c.execute("SELECT COUNT(*) as cnt FROM users")
            u_cnt = c.fetchone()["cnt"]
            c.execute("SELECT * FROM deposits WHERE status = 'pending' ORDER BY id DESC")
            p_deps = [dict(r) for r in c.fetchall()]
            c.execute("SELECT * FROM orders WHERE status = 'pending_delivery' ORDER BY id DESC")
            p_ords = [dict(r) for r in c.fetchall()]
            conn.close()
            return self._send_json({
                "ok": True,
                "stats": {"users": u_cnt, "pending_deposits": len(p_deps), "pending_orders": len(p_ords)},
                "pending_deposits": p_deps,
                "pending_orders": p_ords,
            })

        self.send_response(200)
        self.send_header("Content-type", "text/plain; charset=utf-8")
        self.end_headers()
        self.wfile.write(b"CyberNest Bot is running and healthy!")

    def do_POST(self):
        path = urllib.parse.urlparse(self.path).path
        if path.startswith("/api/"):
            content_length = int(self.headers.get("Content-Length", 0))
            raw_body = self.rfile.read(content_length) if content_length > 0 else b"{}"
            try:
                body = json.loads(raw_body.decode("utf-8"))
            except Exception:
                body = {}
            uid, tg_u = self._get_auth_user()
            if not uid:
                return self._send_json({"ok": False, "error": "Unauthorized"}, 401)

            if path == "/api/coupon/check":
                code_str = str(body.get("code", "")).strip().upper()
                conn = get_db()
                c = conn.cursor()
                c.execute("SELECT discount_pct FROM coupons WHERE code = ? AND is_active = 1", (code_str,))
                row = c.fetchone()
                conn.close()
                if row:
                    return self._send_json({"ok": True, "discount_pct": row["discount_pct"]})
                return self._send_json({"ok": False, "error": "Invalid or expired coupon"})

            if path == "/api/tools/2fa":
                sec = str(body.get("secret", "")).replace(" ", "").upper()
                try:
                    totp = pyotp.TOTP(sec)
                    rem = 30 - (int(datetime.datetime.now().timestamp()) % 30)
                    return self._send_json({"ok": True, "code": totp.now(), "remaining": rem})
                except Exception:
                    return self._send_json({"ok": False, "error": "Invalid 2FA Secret Key"})

            if path == "/api/tools/outlook":
                res = fetch_outlook_otp(str(body.get("account", "")))
                return self._send_json(res)

            if path == "/api/deposit":
                method = str(body.get("method", "bKash"))
                amount = float(body.get("amount", 0))
                trx_id = str(body.get("trx_id", "")).strip()
                if amount <= 0 or not trx_id:
                    return self._send_json({"ok": False, "error": "Invalid amount or TrxID"})
                conn = get_db()
                c = conn.cursor()
                now_str = datetime.datetime.now().strftime("%Y-%m-%d %I:%M %p")
                c.execute(
                    "INSERT INTO deposits (user_id, method, amount, trx_id, status, created_at) VALUES (?, ?, ?, ?, 'pending', ?)",
                    (uid, method, amount, trx_id, now_str),
                )
                dep_id = c.lastrowid
                conn.commit()
                conn.close()
                add_notification(uid, "💳 Deposit Submitted", f"{method} {amount:.2f} BDT (TrxID: {trx_id}) is pending approval.")
                if _GLOBAL_EVENT_LOOP and _GLOBAL_BOT_APP:
                    kb = InlineKeyboardMarkup([[
                        InlineKeyboardButton(f"✅ Approve #{dep_id}", callback_data=f"adm_dep_approve_{dep_id}"),
                        InlineKeyboardButton("❌ Reject", callback_data=f"adm_dep_reject_{dep_id}")
                    ]])
                    asyncio.run_coroutine_threadsafe(
                        _GLOBAL_BOT_APP.bot.send_message(
                            chat_id=SUPER_ADMIN_ID,
                            text=f"🔔 <b>Mini App Deposit Request #{dep_id}</b>\n👤 User: <code>{uid}</code>\n💳 Method: {method}\n💰 Amount: <b>{amount:.2f} BDT</b>\n🔑 TrxID: <code>{trx_id}</code>",
                            parse_mode=ParseMode.HTML,
                            reply_markup=kb,
                        ),
                        _GLOBAL_EVENT_LOOP
                    )
                return self._send_json({"ok": True})

            if path == "/api/order":
                pid = int(body.get("product_id", 0))
                qty = max(1, int(body.get("quantity", 1)))
                coupon = str(body.get("coupon", "")).strip().upper()
                conn = get_db()
                c = conn.cursor()
                c.execute("""
                    SELECT p.*, s.name as sub_name, c.name as cat_name
                    FROM products p
                    JOIN subcategories s ON p.subcategory_id = s.id
                    JOIN categories c ON s.category_id = c.id
                    WHERE p.id = ?
                """, (pid,))
                prod = c.fetchone()
                if not prod or not (prod["is_active"] if "is_active" in prod.keys() else 1):
                    conn.close()
                    return self._send_json({"ok": False, "error": "Product unavailable"})
                disc = 0.0
                if coupon:
                    c.execute("SELECT discount_pct FROM coupons WHERE code = ? AND is_active = 1", (coupon,))
                    crow = c.fetchone()
                    if crow:
                        disc = float(crow["discount_pct"])
                total_bdt = prod["price_bdt"] * qty * (1 - disc / 100.0)
                c.execute("SELECT balance_bdt FROM users WHERE user_id = ?", (uid,))
                bal = c.fetchone()["balance_bdt"]
                if bal < total_bdt:
                    conn.close()
                    return self._send_json({"ok": False, "error": f"Insufficient balance! Need {total_bdt:.2f} BDT"})

                c.execute("SELECT id, content FROM stock WHERE product_id = ? AND is_sold = 0 LIMIT ?", (pid, qty))
                srows = c.fetchall()
                is_vpn_or_proxy = ("vpn" in prod["cat_name"].lower()) or ("proxy" in prod["cat_name"].lower())
                if len(srows) < qty and not is_vpn_or_proxy:
                    conn.close()
                    return self._send_json({"ok": False, "error": "Out of stock! Please check back later."})

                now_str = datetime.datetime.now().strftime("%Y-%m-%d %I:%M:%S %p")
                day_name = datetime.datetime.now().strftime("%A")
                receipt_no = f"RCP-{secrets.token_hex(5).upper()}"
                full_name = f"{prod['sub_name']} ({prod['name']})"
                c.execute("UPDATE users SET balance_bdt = balance_bdt - ? WHERE user_id = ?", (total_bdt, uid))

                if len(srows) >= qty:
                    sids = [r["id"] for r in srows]
                    items = [r["content"] for r in srows]
                    c.execute(
                        f"UPDATE stock SET is_sold = 1, sold_to = ?, sold_at = ? WHERE id IN ({','.join(['?']*len(sids))})",
                        [uid, now_str] + sids,
                    )
                    c.execute("""
                        INSERT INTO orders (receipt_no, user_id, product_id, product_name, category_name, quantity, unit_price, total_bdt, total_usdt, status, delivered_data, order_date, day_of_week)
                        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, 'completed', ?, ?, ?)
                    """, (receipt_no, uid, pid, full_name, prod["cat_name"], qty, prod["price_bdt"], total_bdt, total_bdt / 125.0, "\n".join(items), now_str, day_name))
                    conn.commit()
                    conn.close()
                    add_notification(uid, "✅ Order Delivered!", f"{full_name} (x{qty}) delivered to your Telegram inbox.")
                    if _GLOBAL_EVENT_LOOP and _GLOBAL_BOT_APP:
                        msg_card = format_credentials_card(items, full_name, prod["cat_name"])
                        asyncio.run_coroutine_threadsafe(_GLOBAL_BOT_APP.bot.send_message(chat_id=uid, text=msg_card, parse_mode=ParseMode.HTML), _GLOBAL_EVENT_LOOP)
                    return self._send_json({"ok": True, "message": "✅ Order Completed! Credentials sent to your Telegram inbox & Profile History."})
                else:
                    c.execute("""
                        INSERT INTO orders (receipt_no, user_id, product_id, product_name, category_name, quantity, unit_price, total_bdt, total_usdt, status, order_date, day_of_week)
                        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, 'pending_delivery', ?, ?)
                    """, (receipt_no, uid, pid, full_name, prod["cat_name"], qty, prod["price_bdt"], total_bdt, total_bdt / 125.0, now_str, day_name))
                    oid = c.lastrowid
                    conn.commit()
                    conn.close()
                    add_notification(uid, "⏳ Manual Order Pending", f"Order #{oid} ({full_name}) is waiting for admin delivery.")
                    if _GLOBAL_EVENT_LOOP and _GLOBAL_BOT_APP:
                        adm_kb = InlineKeyboardMarkup([
                            [InlineKeyboardButton(f"🚀 Deliver Order #{oid}", callback_data=f"adm_fulfill_{oid}")],
                            [InlineKeyboardButton("❌ Cancel & Refund", callback_data=f"adm_refund_{oid}")]
                        ])
                        asyncio.run_coroutine_threadsafe(
                            _GLOBAL_BOT_APP.bot.send_message(
                                chat_id=SUPER_ADMIN_ID,
                                text=f"🔔 <b>Mini App Manual Order #{oid}</b>\n🛡️ {full_name} (x{qty})\n💰 Paid: {total_bdt:.2f} BDT\n👤 User: <code>{uid}</code>",
                                parse_mode=ParseMode.HTML,
                                reply_markup=adm_kb,
                            ),
                            _GLOBAL_EVENT_LOOP
                        )
                    return self._send_json({"ok": True, "message": f"🎉 Order #{oid} placed! Admin will deliver it in 15-30 mins."})

            if path.startswith("/api/admin/"):
                if not is_admin(uid):
                    return self._send_json({"ok": False, "error": "Forbidden"}, 403)
                conn = get_db()
                c = conn.cursor()

                if path == "/api/admin/deliver_order":
                    oid = int(body.get("order_id", 0))
                    ddata = str(body.get("data", "")).strip()
                    c.execute("SELECT * FROM orders WHERE id = ?", (oid,))
                    o = c.fetchone()
                    if not o:
                        conn.close()
                        return self._send_json({"ok": False, "error": "Order not found"})
                    c.execute("UPDATE orders SET status = 'completed', delivered_data = ? WHERE id = ?", (ddata, oid))
                    conn.commit()
                    conn.close()
                    add_notification(o["user_id"], "✅ Order Delivered!", f"Order #{oid} ({o['product_name']}) has been delivered!")
                    if _GLOBAL_EVENT_LOOP and _GLOBAL_BOT_APP:
                        card = format_credentials_card([ddata], o["product_name"], o["category_name"])
                        asyncio.run_coroutine_threadsafe(_GLOBAL_BOT_APP.bot.send_message(chat_id=o["user_id"], text=card, parse_mode=ParseMode.HTML), _GLOBAL_EVENT_LOOP)
                    return self._send_json({"ok": True})

                if path == "/api/admin/refund_order":
                    oid = int(body.get("order_id", 0))
                    c.execute("SELECT * FROM orders WHERE id = ?", (oid,))
                    o = c.fetchone()
                    if o and o["status"] == "pending_delivery":
                        c.execute("UPDATE orders SET status = 'refunded' WHERE id = ?", (oid,))
                        c.execute("UPDATE users SET balance_bdt = balance_bdt + ? WHERE user_id = ?", (o["total_bdt"], o["user_id"]))
                        conn.commit()
                        add_notification(o["user_id"], "❌ Order Refunded", f"Order #{oid} refunded: +{o['total_bdt']:.2f} BDT")
                    conn.close()
                    return self._send_json({"ok": True})

                if path == "/api/admin/deposit_action":
                    did = int(body.get("deposit_id", 0))
                    act = body.get("action")
                    c.execute("SELECT * FROM deposits WHERE id = ?", (did,))
                    d = c.fetchone()
                    if d and d["status"] == "pending":
                        if act == "approve":
                            c.execute("UPDATE deposits SET status = 'approved' WHERE id = ?", (did,))
                            c.execute("UPDATE users SET balance_bdt = balance_bdt + ? WHERE user_id = ?", (d["amount"], d["user_id"]))
                            add_notification(d["user_id"], "✅ Deposit Approved", f"+{d['amount']:.2f} BDT added to your wallet!")
                        else:
                            c.execute("UPDATE deposits SET status = 'rejected' WHERE id = ?", (did,))
                            add_notification(d["user_id"], "❌ Deposit Rejected", f"Deposit #{did} ({d['amount']:.2f} BDT) was rejected.")
                        conn.commit()
                    conn.close()
                    return self._send_json({"ok": True})

                if path == "/api/admin/product_update":
                    pid = int(body.get("product_id", 0))
                    if body.get("toggle_active"):
                        c.execute("UPDATE products SET is_active = CASE WHEN COALESCE(is_active,1)=1 THEN 0 ELSE 1 END WHERE id = ?", (pid,))
                    if "price_bdt" in body:
                        pr = float(body["price_bdt"])
                        c.execute("UPDATE products SET price_bdt = ?, price_usdt = ? WHERE id = ?", (pr, pr / 125.0, pid))
                    if "image_file_id" in body:
                        c.execute("UPDATE products SET image_file_id = ? WHERE id = ?", (str(body["image_file_id"]).strip() or None, pid))
                    conn.commit()
                    conn.close()
                    return self._send_json({"ok": True})

                if path == "/api/admin/add_stock":
                    pid = int(body.get("product_id", 0))
                    lines = [ln.strip() for ln in str(body.get("lines", "")).splitlines() if ln.strip()]
                    for ln in lines:
                        c.execute("INSERT INTO stock (product_id, content, is_sold) VALUES (?, ?, 0)", (pid, ln))
                    conn.commit()
                    conn.close()
                    return self._send_json({"ok": True, "added": len(lines)})

                if path == "/api/admin/popup":
                    set_setting("popup_enabled", str(body.get("enabled", "0")))
                    set_setting("popup_title", str(body.get("title", "")))
                    set_setting("popup_text", str(body.get("text", "")))
                    set_setting("popup_image", str(body.get("image", "")))
                    conn.close()
                    return self._send_json({"ok": True})

                if path == "/api/admin/coupon":
                    code_s = str(body.get("code", "")).strip().upper()
                    pct = float(body.get("discount_pct", 0))
                    if conn.is_pg:
                        c.execute("INSERT INTO coupons (code, discount_pct, is_active) VALUES (%s, %s, 1) ON CONFLICT (code) DO UPDATE SET discount_pct = EXCLUDED.discount_pct, is_active = 1", (code_s, pct))
                    else:
                        c.execute("INSERT OR REPLACE INTO coupons (code, discount_pct, is_active) VALUES (?, ?, 1)", (code_s, pct))
                    conn.commit()
                    conn.close()
                    return self._send_json({"ok": True})

            return self._send_json({"ok": False, "error": "Unknown route"}, 404)

        self.send_response(404)
        self.end_headers()

    def log_message(self, format, *args):
        pass

def start_health_server():
    port = int(os.environ.get("PORT", 8080))
    server = ThreadingHTTPServer(("0.0.0.0", port), HealthCheckHandler)
    server.serve_forever()

# ------------------------------------------------------------------------------
# 4. HIGH-SPEED PERSISTENT DATABASE LAYER (Neon PostgreSQL + SQLite Fallback)
# ------------------------------------------------------------------------------
DATABASE_URL = os.getenv("DATABASE_URL", "").strip()
IS_POSTGRES = bool(DATABASE_URL and ("postgres" in DATABASE_URL))
_GLOBAL_CONN = None

def get_raw_connection():
    global _GLOBAL_CONN
    if IS_POSTGRES:
        import psycopg2
        import psycopg2.extras
        if _GLOBAL_CONN is not None and not _GLOBAL_CONN.closed:
            try:
                with _GLOBAL_CONN.cursor() as ping_cur:
                    ping_cur.execute("SELECT 1")
            except Exception:
                try:
                    _GLOBAL_CONN.close()
                except Exception:
                    pass
                _GLOBAL_CONN = None
        if _GLOBAL_CONN is None or _GLOBAL_CONN.closed:
            url = DATABASE_URL
            if url.startswith("postgres://"):
                url = url.replace("postgres://", "postgresql://", 1)
            _GLOBAL_CONN = psycopg2.connect(url, cursor_factory=psycopg2.extras.RealDictCursor)
            _GLOBAL_CONN.autocommit = True
        return _GLOBAL_CONN
    else:
        if _GLOBAL_CONN is None:
            _GLOBAL_CONN = sqlite3.connect(DB_FILE, check_same_thread=False)
            _GLOBAL_CONN.row_factory = sqlite3.Row
        return _GLOBAL_CONN

class SmartCursor:
    def __init__(self, raw_cursor, is_pg):
        self.raw = raw_cursor
        self.is_pg = is_pg
        self._lastrowid = None

    def execute(self, sql, params=None):
        if self.is_pg:
            sql = sql.replace("?", "%s")
            sql = sql.replace("INTEGER PRIMARY KEY AUTOINCREMENT", "SERIAL PRIMARY KEY")
            sql = sql.replace("AUTOINCREMENT", "")
            sql = sql.replace("INTEGER PRIMARY KEY", "BIGINT PRIMARY KEY")
            sql = sql.replace("INSERT OR IGNORE INTO", "INSERT INTO")
            if "INSERT INTO settings" in sql and "ON CONFLICT" not in sql:
                sql += " ON CONFLICT (key) DO NOTHING"
            if "INSERT INTO users" in sql and "ON CONFLICT" not in sql:
                sql += " ON CONFLICT (user_id) DO NOTHING"

            if sql.strip().upper().startswith("INSERT INTO") and ("RETURNING" not in sql.upper()) and ("SETTINGS" not in sql.upper()) and ("USERS" not in sql.upper()) and ("COUPONS" not in sql.upper()):
                sql += " RETURNING id"
                try:
                    res = self.raw.execute(sql, params) if params is not None else self.raw.execute(sql)
                    row = self.raw.fetchone()
                    if row:
                        self._lastrowid = row["id"] if isinstance(row, dict) else row[0]
                    return res
                except Exception as e:
                    logger.warning(f"Insert returning error: {e}")
                    pass
            try:
                if params is not None:
                    return self.raw.execute(sql, params)
                return self.raw.execute(sql)
            except Exception as e:
                try:
                    self.raw.connection.rollback()
                except Exception:
                    pass
                logger.error(f"SQL execution error: {e} in {sql}")
                raise
        else:
            if params is not None:
                res = self.raw.execute(sql, params)
            else:
                res = self.raw.execute(sql)
            self._lastrowid = getattr(self.raw, "lastrowid", None)
            return res

    @property
    def lastrowid(self):
        return self._lastrowid

    def fetchone(self):
        return self.raw.fetchone()

    def fetchall(self):
        return self.raw.fetchall()

class DBWrapper:
    def __init__(self):
        self.is_pg = IS_POSTGRES
        self.conn = get_raw_connection()

    def cursor(self):
        try:
            return SmartCursor(self.conn.cursor(), self.is_pg)
        except Exception:
            global _GLOBAL_CONN
            _GLOBAL_CONN = None
            self.conn = get_raw_connection()
            return SmartCursor(self.conn.cursor(), self.is_pg)

    def commit(self):
        if not self.is_pg:
            self.conn.commit()

    def close(self):
        pass

def get_db():
    return DBWrapper()

def sync_vpn_catalog_v3(conn, cursor):
    try:
        cursor.execute("SELECT value FROM settings WHERE key = ?", ("vpn_catalog_v3_synced",))
        if cursor.fetchone():
            return
        cursor.execute("SELECT id FROM categories WHERE name LIKE ?", ("%VPN%",))
        cat_row = cursor.fetchone()
        if cat_row:
            vpn_cat_id = cat_row["id"] if isinstance(cat_row, dict) else cat_row[0]
            cursor.execute("SELECT id FROM subcategories WHERE category_id = ?", (vpn_cat_id,))
            old_subs = [r["id"] if isinstance(r, dict) else r[0] for r in cursor.fetchall()]
            for sid in old_subs:
                cursor.execute("DELETE FROM stock WHERE product_id IN (SELECT id FROM products WHERE subcategory_id = ?)", (sid,))
                cursor.execute("DELETE FROM products WHERE subcategory_id = ?", (sid,))
            cursor.execute("DELETE FROM subcategories WHERE category_id = ?", (vpn_cat_id,))
            vpn_catalog = {
                "Nord VPN": [("7 Days", 20.0), ("3 Days", 15.0), ("15 Days", 35.0), ("30 Days", 60.0)],
                "Express VPN": [("3 Days", 15.0), ("7 Days", 20.0), ("15 Days", 35.0), ("30 Days", 65.0)],
                "IP Vanish": [("7 Days", 26.0), ("3 Days", 18.0), ("30 Days", 60.0)],
                "Turbo VPN": [("7 Days", 38.0), ("30 Days", 75.0)],
                "Surfshark VPN": [("7 Days", 26.0), ("15 Days", 38.0), ("30 Days", 65.0)],
                "HMA VPN": [("7 Days", 26.0), ("30 Days", 65.0)],
                "Panda VPN": [("3 Days", 20.0), ("7 Days", 28.0)],
                "PROTON VPN": [("14 Days", 51.0), ("7 Days", 30.0), ("30 Days", 80.0)],
            }
            for pos_idx, (brand, variants) in enumerate(vpn_catalog.items(), start=1):
                cursor.execute("INSERT INTO subcategories (category_id, name, position) VALUES (?, ?, ?)", (vpn_cat_id, brand, pos_idx))
                sub_id = cursor.lastrowid
                for dur, bdt in variants:
                    cursor.execute(
                        "INSERT INTO products (subcategory_id, name, price_bdt, price_usdt, is_on_demand, is_active) VALUES (?, ?, ?, ?, 1, 1)",
                        (sub_id, dur, bdt, bdt / 125.0),
                    )
        if conn.is_pg:
            cursor.execute("INSERT INTO settings (key, value) VALUES (%s, %s) ON CONFLICT (key) DO NOTHING", ("vpn_catalog_v3_synced", "1"))
        else:
            cursor.execute("INSERT OR IGNORE INTO settings (key, value) VALUES (?, ?)", ("vpn_catalog_v3_synced", "1"))
        conn.commit()
    except Exception as e:
        logger.error(f"sync_vpn_catalog_v3 error: {e}", exc_info=True)

def init_db():
    conn = get_db()
    cursor = conn.cursor()

    cursor.execute("""
    CREATE TABLE IF NOT EXISTS users (
        user_id BIGINT PRIMARY KEY,
        name TEXT,
        username TEXT,
        balance_bdt REAL DEFAULT 0.0,
        joined_date TEXT,
        is_admin INTEGER DEFAULT 0,
        current_state TEXT,
        state_data TEXT
    )
    """)

    cursor.execute("""
    CREATE TABLE IF NOT EXISTS categories (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        name TEXT NOT NULL,
        icon TEXT DEFAULT '🛍️',
        position INTEGER DEFAULT 0
    )
    """)

    cursor.execute("""
    CREATE TABLE IF NOT EXISTS subcategories (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        category_id INTEGER,
        name TEXT NOT NULL,
        position INTEGER DEFAULT 0,
        FOREIGN KEY (category_id) REFERENCES categories (id) ON DELETE CASCADE
    )
    """)

    cursor.execute("""
    CREATE TABLE IF NOT EXISTS products (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        subcategory_id INTEGER,
        name TEXT NOT NULL,
        price_bdt REAL NOT NULL,
        price_usdt REAL NOT NULL,
        image_file_id TEXT DEFAULT NULL,
        is_on_demand INTEGER DEFAULT 0,
        description TEXT,
        is_active INTEGER DEFAULT 1,
        FOREIGN KEY (subcategory_id) REFERENCES subcategories (id) ON DELETE CASCADE
    )
    """)

    cursor.execute("""
    CREATE TABLE IF NOT EXISTS stock (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        product_id INTEGER,
        content TEXT NOT NULL,
        stock_type TEXT DEFAULT 'account',
        is_sold INTEGER DEFAULT 0,
        sold_to BIGINT,
        sold_at TEXT,
        FOREIGN KEY (product_id) REFERENCES products (id) ON DELETE CASCADE
    )
    """)

    cursor.execute("""
    CREATE TABLE IF NOT EXISTS orders (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        receipt_no TEXT UNIQUE,
        user_id BIGINT,
        product_id INTEGER,
        product_name TEXT,
        category_name TEXT,
        quantity INTEGER,
        unit_price REAL,
        total_bdt REAL,
        total_usdt REAL,
        status TEXT DEFAULT 'completed',
        delivered_data TEXT,
        order_date TEXT,
        day_of_week TEXT,
        FOREIGN KEY (user_id) REFERENCES users (user_id)
    )
    """)

    cursor.execute("""
    CREATE TABLE IF NOT EXISTS deposits (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        user_id BIGINT,
        method TEXT,
        amount REAL,
        trx_id TEXT,
        status TEXT DEFAULT 'pending',
        created_at TEXT,
        FOREIGN KEY (user_id) REFERENCES users (user_id)
    )
    """)

    cursor.execute("""
    CREATE TABLE IF NOT EXISTS settings (
        key TEXT PRIMARY KEY,
        value TEXT
    )
    """)

    cursor.execute("""
    CREATE TABLE IF NOT EXISTS notifications (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        user_id BIGINT,
        title TEXT,
        message TEXT,
        is_read INTEGER DEFAULT 0,
        created_at TEXT
    )
    """)

    cursor.execute("""
    CREATE TABLE IF NOT EXISTS coupons (
        code TEXT PRIMARY KEY,
        discount_pct REAL DEFAULT 0.0,
        is_active INTEGER DEFAULT 1
    )
    """)

    if conn.is_pg:
        try:
            cursor.execute("ALTER TABLE categories ADD COLUMN IF NOT EXISTS position INTEGER DEFAULT 0")
            cursor.execute("ALTER TABLE subcategories ADD COLUMN IF NOT EXISTS position INTEGER DEFAULT 0")
            cursor.execute("ALTER TABLE stock ADD COLUMN IF NOT EXISTS stock_type TEXT DEFAULT 'account'")
            cursor.execute("ALTER TABLE products ADD COLUMN IF NOT EXISTS is_active INTEGER DEFAULT 1")
        except Exception as e:
            logger.warning(f"Postgres column check notice: {e}")
    else:
        for tbl in ("categories", "subcategories"):
            try:
                cursor.execute(f"ALTER TABLE {tbl} ADD COLUMN position INTEGER DEFAULT 0")
            except Exception:
                pass
        try:
            cursor.execute("ALTER TABLE stock ADD COLUMN stock_type TEXT DEFAULT 'account'")
        except Exception:
            pass
        try:
            cursor.execute("ALTER TABLE products ADD COLUMN is_active INTEGER DEFAULT 1")
        except Exception:
            pass

    default_settings = {
        "shop_name": "CyberNest Market",
        "shop_tagline": "Digital Products & Services",
        "usdt_rate": "125.0",
        "min_deposit_manual": "20.0",
        "support_username": "@Talha_juba098",
        "bkash_number": "01XXXXXXXXX (Personal - Send Money)",
        "nagad_number": "01XXXXXXXXX (Personal - Send Money)",
        "binance_id": "Pay ID: 12345678 / USDT (TRC20): TQxxxxxxxxxxxxxxxxxx",
        "mail_code_url_1": "https://code.yamin.bd/",
    }
    for k, v in default_settings.items():
        if conn.is_pg:
            cursor.execute("INSERT INTO settings (key, value) VALUES (%s, %s) ON CONFLICT (key) DO NOTHING", (k, v))
        else:
            cursor.execute("INSERT OR IGNORE INTO settings (key, value) VALUES (?, ?)", (k, v))
    conn.commit()

    cursor.execute("SELECT COUNT(*) as cnt FROM categories")
    row = cursor.fetchone()
    cnt = row["cnt"] if isinstance(row, dict) else row[0]
    cursor.execute("SELECT COUNT(*) as cnt FROM subcategories")
    s_row = cursor.fetchone()
    s_cnt = s_row["cnt"] if isinstance(s_row, dict) else s_row[0]
    if cnt == 0 or s_cnt == 0:
        try:
            cursor.execute("DELETE FROM products")
            cursor.execute("DELETE FROM subcategories")
            cursor.execute("DELETE FROM categories")
            conn.commit()
        except Exception:
            pass
        seed_clean_store(conn)

    sync_vpn_catalog_v3(conn, cursor)
    conn.close()

def seed_clean_store(conn):
    cursor = conn.cursor()
    cats = [
        ("🛡️ VPN", "🛡️"),
        ("🌐 Proxy", "🌐"),
        ("✉️ Mail", "✉️"),
        ("✉️ Recovery Added Mail", "✉️"),
    ]
    for cname, cicon in cats:
        cursor.execute("INSERT INTO categories (name, icon) VALUES (?, ?)", (cname, cicon))
    conn.commit()

    cursor.execute("SELECT id, name FROM categories")
    cat_map = {row["name"]: row["id"] for row in cursor.fetchall()}

    vpn_id = cat_map.get("🛡️ VPN")
    vpn_catalog = {
        "Nord VPN": [("7 Days", 20.0), ("3 Days", 15.0), ("15 Days", 35.0), ("30 Days", 60.0)],
        "Express VPN": [("3 Days", 15.0), ("7 Days", 20.0), ("15 Days", 35.0), ("30 Days", 65.0)],
        "IP Vanish": [("7 Days", 26.0), ("3 Days", 18.0), ("30 Days", 60.0)],
        "Turbo VPN": [("7 Days", 38.0), ("30 Days", 75.0)],
        "Surfshark VPN": [("7 Days", 26.0), ("15 Days", 38.0), ("30 Days", 65.0)],
        "HMA VPN": [("7 Days", 26.0), ("30 Days", 65.0)],
        "Panda VPN": [("3 Days", 20.0), ("7 Days", 28.0)],
        "PROTON VPN": [("14 Days", 51.0), ("7 Days", 30.0), ("30 Days", 80.0)],
    }
    for brand, variants in vpn_catalog.items():
        cursor.execute("INSERT INTO subcategories (category_id, name) VALUES (?, ?)", (vpn_id, brand))
        sub_id = cursor.lastrowid
        for dur_name, bdt in variants:
            cursor.execute(
                "INSERT INTO products (subcategory_id, name, price_bdt, price_usdt, is_on_demand) VALUES (?, ?, ?, ?, 1)",
                (sub_id, dur_name, bdt, bdt / 125.0),
            )

    proxy_id = cat_map.get("🌐 Proxy")
    proxy_catalog = {
        "9 Proxy": [("100 MB", 12.0), ("200 MB", 23.0), ("500 MB", 60.0), ("1 GB", 110.0)],
        "Rapid Proxy": [("100 MB", 12.0), ("200 MB", 25.0), ("500 MB", 60.0), ("1 GB", 115.0)],
        "711 Proxy": [("100 MB", 10.0), ("200 MB", 18.0), ("500 MB", 45.0)],
    }
    for pbrand, pvariants in proxy_catalog.items():
        cursor.execute("INSERT INTO subcategories (category_id, name) VALUES (?, ?)", (proxy_id, pbrand))
        psub_id = cursor.lastrowid
        for pkg_name, bdt in pvariants:
            cursor.execute(
                "INSERT INTO products (subcategory_id, name, price_bdt, price_usdt, is_on_demand) VALUES (?, ?, ?, ?, 0)",
                (psub_id, pkg_name, bdt, bdt / 125.0),
            )

    mail_id = cat_map.get("✉️ Mail")
    mail_catalog = {
        "Hotmail": [("Hotmail", 0.65)],
        "Outlook": [("Outlook", 0.68)],
        "FR Outlook": [("FR Outlook", 0.87)],
    }
    for mpos, (mbrand, mvariants) in enumerate(mail_catalog.items(), start=1):
        cursor.execute("INSERT INTO subcategories (category_id, name, position) VALUES (?, ?, ?)", (mail_id, mbrand, mpos))
        msub_id = cursor.lastrowid
        for vname, bdt in mvariants:
            cursor.execute(
                "INSERT INTO products (subcategory_id, name, price_bdt, price_usdt, is_on_demand) VALUES (?, ?, ?, ?, 0)",
                (msub_id, vname, bdt, bdt / 125.0),
            )

    rec_mail_id = cat_map.get("✉️ Recovery Added Mail")
    rec_catalog = {
        "Recovery Hotmail": [("Recovery Added Hotmail", 0.95)],
        "Recovery Outlook": [("Recovery Added Outlook", 0.95)],
        "Recovery FR Outlook": [("Recovery Added FR Outlook", 1.20)],
    }
    for rpos, (rbrand, rvariants) in enumerate(rec_catalog.items(), start=1):
        cursor.execute("INSERT INTO subcategories (category_id, name, position) VALUES (?, ?, ?)", (rec_mail_id, rbrand, rpos))
        rsub_id = cursor.lastrowid
        for vname, bdt in rvariants:
            cursor.execute(
                "INSERT INTO products (subcategory_id, name, price_bdt, price_usdt, is_on_demand) VALUES (?, ?, ?, ?, 0)",
                (rsub_id, vname, bdt, bdt / 125.0),
            )
    conn.commit()

def format_price(val: float) -> str:
    if val is None:
        return "0"
    if val < 1:
        return f"{val:.2f}"
    if val == int(val):
        return f"{int(val)}"
    return f"{val:.2f}".rstrip('0').rstrip('.')

def get_setting(key: str, default: str = "") -> str:
    conn = get_db()
    c = conn.cursor()
    c.execute("SELECT value FROM settings WHERE key = ?", (key,))
    row = c.fetchone()
    conn.close()
    return row["value"] if row else default

def set_setting(key: str, value: str):
    conn = get_db()
    c = conn.cursor()
    if conn.is_pg:
        c.execute("INSERT INTO settings (key, value) VALUES (%s, %s) ON CONFLICT (key) DO UPDATE SET value = EXCLUDED.value", (key, value))
    else:
        c.execute("INSERT OR REPLACE INTO settings (key, value) VALUES (?, ?)", (key, value))
    conn.commit()
    conn.close()

def get_or_create_user(user_id: int, name: str, username: str):
    conn = get_db()
    c = conn.cursor()
    c.execute("SELECT * FROM users WHERE user_id = ?", (user_id,))
    row = c.fetchone()
    if not row:
        joined = datetime.datetime.now().strftime("%Y-%m-%d")
        is_adm = 1 if user_id == SUPER_ADMIN_ID else 0
        c.execute(
            "INSERT INTO users (user_id, name, username, balance_bdt, joined_date, is_admin) VALUES (?, ?, ?, ?, ?, ?)",
            (user_id, name, username or "", 0.0, joined, is_adm),
        )
        conn.commit()
        c.execute("SELECT * FROM users WHERE user_id = ?", (user_id,))
        row = c.fetchone()
    else:
        if user_id == SUPER_ADMIN_ID and row["is_admin"] != 1:
            c.execute("UPDATE users SET is_admin = 1 WHERE user_id = ?", (user_id,))
            conn.commit()
            c.execute("SELECT * FROM users WHERE user_id = ?", (user_id,))
            row = c.fetchone()
    conn.close()
    return dict(row)

def set_user_state(user_id: int, state: Optional[str], data: Optional[Dict] = None):
    conn = get_db()
    c = conn.cursor()
    data_str = json.dumps(data) if data else "{}"
    c.execute("UPDATE users SET current_state = ?, state_data = ? WHERE user_id = ?", (state, data_str, user_id))
    conn.commit()
    conn.close()

def get_user_state(user_id: int):
    conn = get_db()
    c = conn.cursor()
    c.execute("SELECT current_state, state_data FROM users WHERE user_id = ?", (user_id,))
    row = c.fetchone()
    conn.close()
    if row:
        data = json.loads(row["state_data"]) if row["state_data"] else {}
        return row["current_state"], data
    return None, {}

def is_admin(user_id: int) -> bool:
    return int(user_id) == SUPER_ADMIN_ID

# ------------------------------------------------------------------------------
# 5. KEYBOARD LAYOUTS
# ------------------------------------------------------------------------------
DEFAULT_MENU_LAYOUT = [
    ["🛍️ Buy Products"],
    ["👤 My Profile", "💳 Deposit"],
    ["🔐 Get Code", "💬 Support"],
]

def get_menu_layout() -> List[List[str]]:
    raw = get_setting("main_menu_layout", "")
    if raw:
        try:
            parsed = json.loads(raw)
            if isinstance(parsed, list) and parsed:
                return parsed
        except Exception:
            pass
    return [list(r) for r in DEFAULT_MENU_LAYOUT]

def save_menu_layout(layout: List[List[str]]):
    clean = [r for r in layout if r]
    set_setting("main_menu_layout", json.dumps(clean, ensure_ascii=False))

def get_main_reply_keyboard(user_id: int) -> ReplyKeyboardMarkup:
    layout = get_menu_layout()
    buttons = [[KeyboardButton(label) for label in row] for row in layout]
    if is_admin(user_id):
        buttons.append([KeyboardButton("👑 Admin Panel")])
    return ReplyKeyboardMarkup(buttons, resize_keyboard=True)

# ------------------------------------------------------------------------------
# 6. EXCEL & PDF RECEIPT GENERATORS
# ------------------------------------------------------------------------------
def generate_excel_delivery(items: List[str], product_name: str, receipt_no: str) -> io.BytesIO:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Credentials"

    header_fill = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    data_font = Font(name="Calibri", size=10)
    thin_border = Border(
        left=Side(style='thin', color='D9D9D9'),
        right=Side(style='thin', color='D9D9D9'),
        top=Side(style='thin', color='D9D9D9'),
        bottom=Side(style='thin', color='D9D9D9')
    )

    headers = ["SL", "Product", "Email / Username", "Password", "Token / Session / Key", "Full Raw String", "Delivery Date"]
    ws.append(headers)
    for col_num in range(1, len(headers) + 1):
        cell = ws.cell(row=1, column=col_num)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")

    now_str = datetime.datetime.now().strftime("%Y-%m-%d %I:%M %p")
    for idx, item in enumerate(items, start=1):
        delim = "|" if "|" in item else ":"
        parts = [p.strip() for p in item.split(delim) if p.strip()]
        email = parts[0] if len(parts) >= 1 else item
        pwd = parts[1] if len(parts) >= 2 else ""
        token = parts[2] if len(parts) >= 3 else ""
        row_data = [idx, product_name, email, pwd, token, item, now_str]
        ws.append(row_data)
        row_idx = idx + 1
        for c in range(1, len(row_data) + 1):
            cell = ws.cell(row=row_idx, column=c)
            cell.font = data_font
            cell.border = thin_border
            if c in (1, 7):
                cell.alignment = Alignment(horizontal="center", vertical="center")

    ws.column_dimensions['A'].width = 6
    ws.column_dimensions['B'].width = 22
    ws.column_dimensions['C'].width = 32
    ws.column_dimensions['D'].width = 20
    ws.column_dimensions['E'].width = 45
    ws.column_dimensions['F'].width = 50
    ws.column_dimensions['G'].width = 22

    bio = io.BytesIO()
    wb.save(bio)
    bio.seek(0)
    return bio

def read_items_from_excel(file_bytes: bytes) -> List[str]:
    items = []
    wb = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=True)
    ws = wb.active
    for row in ws.iter_rows(values_only=True):
        for cell in row:
            if cell is not None:
                val = str(cell).strip()
                if val.lower() in ("sl", "no", "mail", "email", "password", "account", "credential", "product / variant", "delivery date"):
                    continue
                if len(val) > 2:
                    items.append(val)
    return items

def generate_pdf_receipt(
    receipt_no: str,
    date_str: str,
    day_name: str,
    customer_name: str,
    telegram_id: int,
    username: str,
    item_desc: str,
    category_name: str,
    qty: int,
    unit_price: float,
    total_bdt: float,
    bot_username: str,
) -> io.BytesIO:
    bio = io.BytesIO()
    doc = SimpleDocTemplate(
        bio,
        pagesize=A4,
        rightMargin=36,
        leftMargin=36,
        topMargin=36,
        bottomMargin=36,
    )

    styles = getSampleStyleSheet()
    shop_title_style = ParagraphStyle(
        'ShopTitle',
        parent=styles['Heading1'],
        fontName='Helvetica-Bold',
        fontSize=20,
        textColor=colors.HexColor('#FFFFFF'),
        alignment=0,
        leading=22,
    )

    story = []
    shop_name = get_setting("shop_name", "CyberNest Market")
    shop_tagline = get_setting("shop_tagline", "Digital Products & Services")

    try:
        logo_stream = io.BytesIO(base64.b64decode(CYBERNEST_LOGO_BASE64))
        logo_img = RLImage(logo_stream, width=54, height=54)
    except Exception:
        logo_img = ""

    header_html = f"""
    <b>{shop_name}</b><br/>
    <font size=9 color='#E0E0E0'>{shop_tagline}</font><br/>
    <font size=11 color='#FFD700'><b>★ OFFICIAL PAYMENT RECEIPT ★</b></font>
    """
    p_header = Paragraph(header_html, shop_title_style)
    badge_html = "<br/><font color='#000000' size=11><b>PAID IN FULL</b></font>"
    p_badge = Paragraph(badge_html, ParagraphStyle('Badge', alignment=1, textColor=colors.black))

    if logo_img:
        header_table = Table([[logo_img, p_header, p_badge]], colWidths=[65, 335, 120])
        header_table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (1, 0), colors.HexColor('#4A154B')),
            ('BACKGROUND', (2, 0), (2, 0), colors.HexColor('#FFCC00')),
            ('ALIGN', (0, 0), (0, 0), 'CENTER'),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 12),
            ('TOPPADDING', (0, 0), (-1, -1), 12),
            ('LEFTPADDING', (0, 0), (0, 0), 10),
            ('RIGHTPADDING', (0, 0), (0, 0), 4),
        ]))
    else:
        header_table = Table([[p_header, p_badge]], colWidths=[400, 120])
        header_table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (0, 0), colors.HexColor('#4A154B')),
            ('BACKGROUND', (1, 0), (1, 0), colors.HexColor('#FFCC00')),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 14),
            ('TOPPADDING', (0, 0), (-1, -1), 14),
        ]))
    story.append(header_table)
    story.append(Spacer(1, 20))

    meta_data = [
        [
            Paragraph(f"<b>RECEIPT NO.</b><br/><font color='#C2185B' size=12><b>{receipt_no}</b></font>", styles['Normal']),
            Paragraph(f"<b>DATE ISSUED</b><br/>{date_str} ({day_name})", styles['Normal']),
        ],
        [
            Paragraph(f"<b>CUSTOMER</b><br/>{customer_name}", styles['Normal']),
            Paragraph(f"<b>TELEGRAM ID</b><br/><code>{telegram_id}</code>", styles['Normal']),
        ],
        [
            Paragraph(f"<b>USERNAME</b><br/>@{username if username else 'N/A'}", styles['Normal']),
            Paragraph("<b>PAYMENT METHOD</b><br/><font color='#2E7D32'><b>Wallet Balance</b></font>", styles['Normal']),
        ]
    ]
    meta_table = Table(meta_data, colWidths=[260, 260])
    meta_table.setStyle(TableStyle([
        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 10),
        ('LINEBELOW', (0, -1), (-1, -1), 1, colors.HexColor('#E0E0E0')),
    ]))
    story.append(meta_table)
    story.append(Spacer(1, 15))

    badge_data = [[Paragraph(f"<font color='#4A154B'><b>CATEGORIES: {category_name.upper()}</b></font>", styles['Normal'])]]
    t_badge = Table(badge_data, colWidths=[520])
    t_badge.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, -1), colors.HexColor('#F3E5F5')),
        ('TOPPADDING', (0, 0), (-1, -1), 6),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 6),
    ]))
    story.append(t_badge)
    story.append(Spacer(1, 8))

    item_table_data = [
        ["ITEM DESCRIPTION", "QTY", "PRICE", "AMOUNT"],
        [Paragraph(f"<b>{item_desc}</b><br/><font size=8 color='#666'>Instant Account / Service Delivery</font>", styles['Normal']),
         str(qty), f"{unit_price:.2f}", f"{total_bdt:.2f}"],
    ]
    item_table = Table(item_table_data, colWidths=[280, 60, 90, 90])
    item_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#1A237E')),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, 0), 10),
        ('ALIGN', (1, 0), (-1, -1), 'CENTER'),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('TOPPADDING', (0, 0), (-1, -1), 10),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 10),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor('#E0E0E0')),
    ]))
    story.append(item_table)

    total_data = [["TOTAL PAID (BDT)", f"{total_bdt:.2f} BDT"]]
    total_table = Table(total_data, colWidths=[370, 150])
    total_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, -1), colors.HexColor('#7B1FA2')),
        ('TEXTCOLOR', (0, 0), (-1, -1), colors.white),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, 0), 13),
        ('ALIGN', (1, 0), (1, 0), 'CENTER'),
        ('TOPPADDING', (0, 0), (-1, -1), 10),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 10),
    ]))
    story.append(total_table)
    story.append(Spacer(1, 25))

    clean_bot = bot_username.lstrip('@')
    verify_url = f"https://t.me/{clean_bot}?start=order_{receipt_no}"

    qr_widget = qr.QrCodeWidget(verify_url)
    qr_widget.barWidth = 85
    qr_widget.barHeight = 85
    qr_widget.qrVersion = 2
    d_qr = Drawing(85, 85)
    d_qr.add(qr_widget)

    ref_text = (
        f"<b>Order Reference:</b><br/>"
        f"<b>{receipt_no}</b><br/>"
        f"<font color='#2E7D32'>✔ Payment Verified</font><br/><br/>"
        f"<font size=8 color='#555'>Scan the QR code to verify this order directly in Telegram.</font>"
    )
    p_ref = Paragraph(ref_text, styles['Normal'])

    footer_grid = Table([[p_ref, d_qr]], colWidths=[400, 120])
    footer_grid.setStyle(TableStyle([
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('ALIGN', (1, 0), (1, 0), 'CENTER'),
    ]))
    story.append(footer_grid)
    story.append(Spacer(1, 20))

    disc = (
        "<i>This is a system-generated receipt and does not require a signature.</i><br/>"
        f"<b>Thank you for shopping with {shop_name}!</b>"
    )
    story.append(Paragraph(disc, ParagraphStyle('Disc', alignment=1, fontSize=9, textColor=colors.HexColor('#555555'))))

    doc.build(story)
    bio.seek(0)
    return bio

def format_credentials_card(items: List[str], prod_name: str, cat_name: str) -> str:
    """Formats credentials with one-tap copy code blocks."""
    is_proxy = "PROXY" in cat_name.upper() or "PROXY" in prod_name.upper()
    lines = []

    if is_proxy:
        lines.append("✅ <b>Proxy Delivered!</b>")
        lines.append(f"🌐 <b>{prod_name}</b>")
        lines.append("────────────────────────")
        for idx, item in enumerate(items, start=1):
            parts = [p.strip() for p in item.replace("@", ":").split(":") if p.strip()]
            if len(parts) >= 4:
                host, port, u_name, pwd = parts[0], parts[1], parts[2], parts[3]
            elif len(parts) == 2:
                host, port = parts[0], parts[1]
                u_name, pwd = "N/A", "N/A"
            else:
                host, port, u_name, pwd = item, "N/A", "N/A", "N/A"

            if len(items) > 1:
                lines.append(f"🌐 <b>Proxy #{idx}</b>")
            lines.append(f"🖥️ <b>Host</b>     : <code>{host}</code>")
            lines.append(f"🔌 <b>Port</b>     : <code>{port}</code>")
            lines.append(f"👤 <b>Username</b> : <code>{u_name}</code>")
            lines.append(f"🔐 <b>Password</b> : <code>{pwd}</code>")
            lines.append(f"📋 <b>Full</b>     : <code>{item}</code>")
            if idx < len(items):
                lines.append("")
    else:
        lines.append("✅ <b>Service Delivered!</b>")
        lines.append(f"🛡️ <b>{prod_name}</b>")
        lines.append("────────────────────────")
        for idx, item in enumerate(items, start=1):
            delim = "|" if "|" in item else ":"
            parts = [p.strip() for p in item.split(delim) if p.strip()]
            if len(parts) >= 2:
                email, pwd = parts[0], parts[1]
                if len(items) > 1:
                    lines.append(f"📦 <b>Item #{idx}</b>")
                lines.append(f"📧 <b>Email / User</b> : <code>{email}</code>")
                lines.append(f"🔐 <b>Password</b>     : <code>{pwd}</code>")
                if len(parts) >= 3:
                    tok = parts[2]
                    tok_disp = tok[:100] + "..." if len(tok) > 100 else tok
                    lines.append(f"🎟️ <b>Token / Session</b> : <code>{tok_disp}</code>")
            else:
                if len(items) > 1:
                    lines.append(f"🔑 <b>Key / Code #{idx}</b>")
                item_disp = item[:100] + "..." if len(item) > 100 else item
                lines.append(f"🎟️ <b>Key / Code</b> : <code>{item_disp}</code>")
            if idx < len(items):
                lines.append("")

    lines.append("────────────────────────")
    lines.append("💡 <i>Tap any value above to copy in one click</i>")
    return "\n".join(lines)

# ------------------------------------------------------------------------------
# 7. USER HANDLERS (Shop, Profile, Deposit, 2FA, Support)
# ------------------------------------------------------------------------------
async def start_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user = update.effective_user
    get_or_create_user(user.id, user.full_name, user.username)
    shop_name = get_setting("shop_name", "CyberNest Market")

    if context.args and len(context.args) > 0:
        arg = context.args[0]
        if arg.startswith("order_"):
            receipt_code = arg.replace("order_", "")
            await show_order_verification(update, receipt_code)
            return

    in_group, in_channel = await check_user_membership(context.bot, user.id)
    if not (in_group and in_channel):
        force_txt = (
            f"👋 <b>স্বাগতম, {user.full_name}!</b>\n\n"
            f"বট ব্যবহার করতে নিচের বাটনে চাপ দিয়ে গ্রুপ ও চ্যানেলে যুক্ত হয়ে ভেরিফাই করুন:"
        )
        kb = get_force_join_keyboard(in_group, in_channel)
        await update.message.reply_text(force_txt, parse_mode=ParseMode.HTML, reply_markup=kb)
        return

    welcome_text = (
        f"🎉 <b>স্বাগতম, {user.full_name}!</b> ✨\n\n"
        f"💎 <b>{shop_name}</b>\n"
        f"─────────────────────────────\n"
        f"🛍️ প্রিমিয়াম VPN, Proxy আর Mail সার্ভিস — সবচেয়ে দ্রুত ডেলিভারিতে, ২৪/৭।\n\n"
        f"📌 নিচের মেনু থেকে বা বাম কোণার <b>Open Shop</b> বাটনে চাপ দিয়ে শপ ওপেন করুন 👇"
    )
    reply_kb = get_main_reply_keyboard(user.id)
    await update.message.reply_text(welcome_text, parse_mode=ParseMode.HTML, reply_markup=reply_kb)

async def check_user_membership(bot, user_id: int) -> Tuple[bool, bool]:
    if user_id == SUPER_ADMIN_ID:
        return True, True

    in_group = True
    in_channel = True

    try:
        m_group = await bot.get_chat_member(chat_id=get_welcome_group_id(), user_id=user_id)
        if m_group.status in ("left", "kicked"):
            in_group = False
    except Exception:
        in_group = False

    try:
        m_chan = await bot.get_chat_member(chat_id=get_transaction_channel_id(), user_id=user_id)
        if m_chan.status in ("left", "kicked"):
            in_channel = False
    except Exception:
        in_channel = False

    return in_group, in_channel

def get_force_join_keyboard(in_group: bool, in_channel: bool) -> InlineKeyboardMarkup:
    buttons = []
    if not in_group:
        buttons.append([InlineKeyboardButton("📢 Join Update Group", url=get_welcome_group_link())])
    if not in_channel:
        buttons.append([InlineKeyboardButton("💳 Join Proof Channel", url=get_transaction_channel_link())])
    buttons.append([InlineKeyboardButton("✅ ভেরিফাই করুন", callback_data="btn_check_verification")])
    return InlineKeyboardMarkup(buttons)

async def execute_smart_broadcast(bot, reply_target, sdata: dict, btn_text: Optional[str] = None, btn_url: Optional[str] = None, buttons_list: Optional[list] = None):
    msg_text = sdata.get("msg_text", "")
    photo_id = sdata.get("photo_id")
    reply_markup = None
    if buttons_list:
        reply_markup = InlineKeyboardMarkup([[InlineKeyboardButton(b["text"], url=b["url"])] for b in buttons_list])
    elif btn_text and btn_url:
        reply_markup = InlineKeyboardMarkup([[InlineKeyboardButton(btn_text, url=btn_url)]])

    conn = get_db()
    c = conn.cursor()
    c.execute("SELECT user_id FROM users")
    users = c.fetchall()
    conn.close()

    sent_cnt = 0
    for u in users:
        uid = u["user_id"]
        try:
            if photo_id:
                await bot.send_photo(chat_id=uid, photo=photo_id, caption=msg_text, parse_mode=ParseMode.HTML, reply_markup=reply_markup)
            else:
                await bot.send_message(chat_id=uid, text=msg_text, parse_mode=ParseMode.HTML, reply_markup=reply_markup)
            sent_cnt += 1
        except Exception:
            pass

    grp_status = "✅ আপডেট গ্রুপে পাঠানো হয়েছে"
    try:
        if photo_id:
            await bot.send_photo(chat_id=get_welcome_group_id(), photo=photo_id, caption=msg_text, parse_mode=ParseMode.HTML, reply_markup=reply_markup)
        else:
            await bot.send_message(chat_id=get_welcome_group_id(), text=msg_text, parse_mode=ParseMode.HTML, reply_markup=reply_markup)
    except Exception as e:
        grp_status = f"⚠️ আপডেট গ্রুপে পাঠানো যায়নি ({e})"

    await reply_target.reply_text(
        f"✅ <b>ব্রডকাস্ট সম্পন্ন!</b>\n"
        f"👥 মোট ইউজার: <b>{sent_cnt}</b> জনের ইনবক্সে পৌঁছেছে।\n"
        f"📢 {grp_status}।\n"
        f"🛡️ <i>ট্রানজ্যাকশন চ্যানেল সম্পূর্ণ পরিষ্কার রাখা হয়েছে।</i>",
        parse_mode=ParseMode.HTML,
    )

async def show_order_verification(update: Update, receipt_no: str):
    conn = get_db()
    c = conn.cursor()
    c.execute("SELECT * FROM orders WHERE receipt_no = ?", (receipt_no,))
    order = c.fetchone()
    conn.close()

    if not order:
        await update.message.reply_text("❌ দুঃখিত! এই রিসিট নম্বরের কোনো অর্ডার পাওয়া যায়নি।")
        return

    txt = (
        f"🧾 <b>অর্ডার ভেরিফিকেশন কার্ড (VERIFIED)</b>\n"
        f"─────────────────────────────\n"
        f"🔹 <b>Receipt No:</b> <code>{order['receipt_no']}</code>\n"
        f"📅 <b>তারিখ ও বার:</b> {order['order_date']} ({order['day_of_week']})\n"
        f"👤 <b>কাস্টমার আইডি:</b> <code>{order['user_id']}</code>\n"
        f"📦 <b>প্রোডাক্ট:</b> {order['product_name']}\n"
        f"🔢 <b>পরিমাণ:</b> {order['quantity']} টি\n"
        f"💰 <b>পরিশোধিত মূল্য:</b> {order['total_bdt']:.2f} BDT\n"
        f"─────────────────────────────\n"
        f"✔ <i>এই ট্রানজ্যাকশনটি সফল ও শতভাগ অথেনটিক।</i>"
    )
    await update.message.reply_text(txt, parse_mode=ParseMode.HTML)

async def profile_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user = update.effective_user
    db_user = get_or_create_user(user.id, user.full_name, user.username)
    rate = float(get_setting("usdt_rate", "125.0"))
    balance_bdt = db_user["balance_bdt"]
    balance_usdt = balance_bdt / rate if rate > 0 else 0.0

    profile_text = (
        f"👤 <b>My Profile</b>\n\n"
        f"👤 <b>Name</b> : {db_user['name']}\n"
        f"💰 <b>Balance</b> : {balance_bdt:.2f} BDT / {balance_usdt:.4f} USDT\n"
        f"🎉 <b>Joined</b> : {db_user['joined_date']}\n"
        f"─────────────────────────────"
    )
    inline_kb = InlineKeyboardMarkup([
        [
            InlineKeyboardButton("💰 Balance Details", callback_data="prof_balance_details"),
            InlineKeyboardButton("🔄 Deposit History", callback_data="prof_deposit_history"),
        ],
        [
            InlineKeyboardButton("👥 Referral", callback_data="prof_referral"),
            InlineKeyboardButton("🛒 Purchases", callback_data="prof_purchases"),
        ]
    ])
    await update.message.reply_text(profile_text, parse_mode=ParseMode.HTML, reply_markup=inline_kb)

async def buy_products_menu_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    conn = get_db()
    c = conn.cursor()
    c.execute("SELECT id, name FROM categories ORDER BY position ASC, id ASC")
    cats = c.fetchall()
    conn.close()

    text = "🛍️ <b>Buy Products</b>\n\nSelect a category:"
    keyboard = []
    for cat in cats:
        keyboard.append([InlineKeyboardButton(cat["name"], callback_data=f"usr_cat_{cat['id']}")])
    keyboard.append([InlineKeyboardButton("❌ Cancel", callback_data="usr_cancel_home")])

    await update.message.reply_text(text, parse_mode=ParseMode.HTML, reply_markup=InlineKeyboardMarkup(keyboard))

async def get_code_menu_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    text = (
        "🔐 <b>Get Code & Authentication Tools</b>\n\n"
        "হটমেইল/আউটলুক ওটিপি কোড এবং লাইভ 2FA কোড পাওয়ার জন্য নিচের অপশন থেকে নির্বাচন করুন:"
    )
    kb = InlineKeyboardMarkup([
        [InlineKeyboardButton("🔑 2FA Authenticator (Live Code)", callback_data="tool_2fa_prompt")],
        [InlineKeyboardButton("🌐 Hotmail & Outlook Code Server", url="https://code.yamin.bd/")],
    ])
    await update.message.reply_text(text, parse_mode=ParseMode.HTML, reply_markup=kb)

async def support_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    text = (
        "💬 <b>কাস্টমার সাপোর্ট</b>\n\n"
        "যেকোনো সমস্যায় নিচে এডমিনকে মেসেজ দিন:"
    )
    kb = InlineKeyboardMarkup([
        [InlineKeyboardButton("📨 Message Admin", url=f"tg://user?id={SUPER_ADMIN_ID}")]
    ])
    await update.message.reply_text(text, parse_mode=ParseMode.HTML, reply_markup=kb)

async def deposit_menu_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user = update.effective_user
    db_user = get_or_create_user(user.id, user.full_name, user.username)
    min_man = float(get_setting("min_deposit_manual", "20.0"))
    set_user_state(user.id, "WAITING_DEP_AMOUNT_MANUAL", {})

    text = (
        f"💳 <b>ডিপোজিট ব্যালেন্স (Send Money / Binance)</b>\n"
        f"─────────────────────────────\n"
        f"💰 বর্তমান ব্যালেন্স: <b>{db_user['balance_bdt']:.2f} BDT</b>\n"
        f"📌 সর্বনিম্ন ডিপোজিট: <b>{min_man:.0f} BDT</b>\n\n"
        f"আপনি কত টাকা ডিপোজিট করতে চান? টাকার পরিমাণ লিখে পাঠান (যেমন: 50, 100, 500):"
    )
    await update.message.reply_text(text, parse_mode=ParseMode.HTML)

# ------------------------------------------------------------------------------
# 8. INLINE CALLBACK QUERY DISPATCHER (Users & Admin Nested Flow)
# ------------------------------------------------------------------------------
async def callback_query_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    _alert_callbacks = (
        "btn_check_verification",
        "adm_toggle_autosales",
        "adm_toggle_autosales_mgr",
    )
    if query.data not in _alert_callbacks:
        try:
            await query.answer()
        except Exception:
            pass
    data = query.data
    user_id = query.from_user.id
    db_user = get_or_create_user(user_id, query.from_user.full_name, query.from_user.username)

    if data == "btn_check_verification":
        try:
            in_group, in_channel = await check_user_membership(context.bot, user_id)
            if in_group and in_channel:
                try:
                    await query.answer("🎉 ভেরিফিকেশন সফল হয়েছে!", show_alert=True)
                except Exception:
                    pass
                try:
                    await query.message.delete()
                except Exception:
                    pass
                shop_name = get_setting("shop_name", "CyberNest Market")
                safe_name = html.escape(query.from_user.full_name or "বন্ধু")
                welcome_text = (
                    f"🎉 <b>স্বাগতম, {safe_name}!</b> ✨\n\n"
                    f"💎 <b>{shop_name}</b>\n"
                    f"─────────────────────────────\n"
                    f"🛍️ প্রিমিয়াম VPN, Proxy আর Mail সার্ভিস — সবচেয়ে দ্রুত ডেলিভারিতে, ২৪/৭।\n\n"
                    f"📌 নিচের মেনু থেকে বেছে নিন 👇"
                )
                reply_kb = get_main_reply_keyboard(user_id)
                await query.message.reply_text(welcome_text, parse_mode=ParseMode.HTML, reply_markup=reply_kb)
            else:
                try:
                    await query.answer("❌ আপনি এখনো গ্রুপ বা চ্যানেলে জয়েন করেননি! দয়া করে জয়েন করে আবার চাপুন।", show_alert=True)
                except Exception:
                    pass
                kb = get_force_join_keyboard(in_group, in_channel)
                try:
                    await query.edit_message_reply_markup(reply_markup=kb)
                except Exception:
                    pass
        except Exception as e:
            logger.error(f"Verification button error: {e}", exc_info=True)
        return

    if data.startswith("usr_cat_"):
        cat_id = int(data.split("_")[2])
        conn = get_db()
        c = conn.cursor()
        c.execute("SELECT name FROM categories WHERE id = ?", (cat_id,))
        cat_row = c.fetchone()
        cat_name = cat_row["name"] if cat_row else "Category"

        c.execute("SELECT id, name FROM subcategories WHERE category_id = ? ORDER BY position ASC, id ASC", (cat_id,))
        subs = c.fetchall()
        conn.close()

        text = f"<b>{cat_name} — Select Brand / Section:</b>"
        kb = []
        row = []
        for s in subs:
            row.append(InlineKeyboardButton(s["name"], callback_data=f"usr_sub_{s['id']}"))
            if len(row) == 2:
                kb.append(row)
                row = []
        if row:
            kb.append(row)
        kb.append([InlineKeyboardButton("< Back", callback_data="usr_root_cats")])
        await query.edit_message_text(text, parse_mode=ParseMode.HTML, reply_markup=InlineKeyboardMarkup(kb))
        return

    if data == "usr_cancel_home":
        await query.message.reply_text("❌ কেনাকাটা বাতিল করা হয়েছে। নতুন অর্ডারের জন্য মেনু ব্যবহার করুন।")
        try:
            await query.message.delete()
        except Exception:
            pass
        return

    if data == "usr_root_cats":
        conn = get_db()
        c = conn.cursor()
        c.execute("SELECT id, name FROM categories ORDER BY position ASC, id ASC")
        cats = c.fetchall()
        conn.close()
        text = "🛍️ <b>Buy Products</b>\n\nSelect a category:"
        keyboard = []
        for cat in cats:
            keyboard.append([InlineKeyboardButton(cat["name"], callback_data=f"usr_cat_{cat['id']}")])
        keyboard.append([InlineKeyboardButton("❌ Cancel", callback_data="usr_cancel_home")])
        await query.edit_message_text(text, parse_mode=ParseMode.HTML, reply_markup=InlineKeyboardMarkup(keyboard))
        return

    if data.startswith("usr_sub_"):
        sub_id = int(data.split("_")[2])
        conn = get_db()
        c = conn.cursor()
        c.execute("""
        SELECT s.name as sub_name, s.category_id, c.name as cat_name
        FROM subcategories s
        JOIN categories c ON s.category_id = c.id
        WHERE s.id = ?
        """, (sub_id,))
        meta = c.fetchone()

        c.execute("""
        SELECT p.id, p.name, p.price_bdt, p.price_usdt, p.is_on_demand, COALESCE(p.is_active, 1) as is_active, COUNT(st.id) as in_stock
        FROM products p
        LEFT JOIN stock st ON p.id = st.product_id AND st.is_sold = 0
        WHERE p.subcategory_id = ?
        GROUP BY p.id, p.name, p.price_bdt, p.price_usdt, p.is_on_demand, p.is_active
        ORDER BY p.id ASC
        """, (sub_id,))
        prods = c.fetchall()
        conn.close()

        if not meta:
            await query.message.reply_text("⚠️ মেনুটি আপডেট করা হয়েছে! অনুগ্রহ করে নিচের 🛍️ Buy Products বাটনে আবার চাপুন।")
            return
        title = f"{meta['cat_name']} ➔ {meta['sub_name']}"
        text = f"<b>{title}</b>\n\nনিচের প্ল্যান বা মেয়াদ নির্বাচন করুন:"
        kb = []
        is_mail_cat = bool(meta and ("mail" in meta["cat_name"].lower() or "মেইল" in meta["cat_name"]))
        for p in prods:
            is_act_val = p["is_active"] if "is_active" in p.keys() else 1
            status_tag = "" if is_act_val else " 🔴 (Unavailable)"
            if is_mail_cat:
                btn_text = f"{p['name']} | {format_price(p['price_bdt'])} BDT | Stock: {p['in_stock']}{status_tag}"
            else:
                btn_text = f"{p['name']} | {format_price(p['price_bdt'])} BDT{status_tag}"
            kb.append([InlineKeyboardButton(btn_text, callback_data=f"usr_prod_{p['id']}")])
        back_cat_id = meta["category_id"] if meta else 1
        kb.append([InlineKeyboardButton("< Back", callback_data=f"usr_cat_{back_cat_id}")])
        await query.edit_message_text(text, parse_mode=ParseMode.HTML, reply_markup=InlineKeyboardMarkup(kb))
        return

    if data.startswith("usr_prod_"):
        prod_id = int(data.split("_")[2])
        conn = get_db()
        c = conn.cursor()
        c.execute("SELECT name, COALESCE(is_active, 1) as is_active FROM products WHERE id = ?", (prod_id,))
        prow = c.fetchone()
        conn.close()
        if prow and not prow["is_active"]:
            await query.message.reply_text(
                f"❌ <b>দুঃখিত!</b> <b>{prow['name']}</b> প্ল্যানটি বর্তমানে এভেলেবেল নেই (স্টকে নেই)।\n"
                f"অনুগ্রহ করে অন্য একটি প্ল্যান নির্বাচন করুন।",
                parse_mode=ParseMode.HTML,
            )
            return
        await render_purchase_summary(query, prod_id, quantity=1, edit=True)
        return

    if data.startswith("calc_add_"):
        parts = data.split("_")
        prod_id = int(parts[2])
        qty = int(parts[3]) + 1
        await render_purchase_summary(query, prod_id, quantity=qty, edit=True)
        return

    if data.startswith("calc_sub_"):
        parts = data.split("_")
        prod_id = int(parts[2])
        qty = max(1, int(parts[3]) - 1)
        await render_purchase_summary(query, prod_id, quantity=qty, edit=True)
        return

    if data.startswith("calc_custom_"):
        prod_id = int(data.split("_")[2])
        set_user_state(user_id, "WAITING_CALC_CUSTOM_QTY", {"prod_id": prod_id})
        await query.message.reply_text("✏️ <b>Enter Quantity:</b>\nআপনি কতগুলো কিনতে চান তা সংখ্যায় লিখে পাঠান (যেমন: 5, 20, 50):", parse_mode=ParseMode.HTML)
        return

    if data.startswith("calc_confirm_"):
        in_group, in_channel = await check_user_membership(context.bot, user_id)
        if not (in_group and in_channel):
            kb = get_force_join_keyboard(in_group, in_channel)
            await query.message.reply_text(
                "⚠️ কিনতে হলে আগে নিচের বাটনে চাপ দিয়ে গ্রুপ ও চ্যানেলে জয়েন করে ভেরিফাই করুন:",
                parse_mode=ParseMode.HTML,
                reply_markup=kb,
            )
            return

        parts = data.split("_")
        prod_id = int(parts[2])
        qty = int(parts[3])
        await process_order(query, prod_id, qty)
        return

    if data.startswith("qty_"):
        parts = data.split("_")
        prod_id = int(parts[1])
        qty = int(parts[2])
        await process_order(query, prod_id, qty)
        return

    if data.startswith("custqty_"):
        prod_id = int(data.split("_")[1])
        set_user_state(user_id, "WAITING_CUSTOM_QTY", {"prod_id": prod_id})
        await query.message.reply_text(
            "✏️ <b>Custom Quantity:</b>\nআপনি কতগুলো কিনতে চান তা সংখ্যায় লিখে পাঠান (যেমন: 5, 20, 50):",
            parse_mode=ParseMode.HTML,
        )
        return

    if data == "usr_dep_manual_menu":
        kb = InlineKeyboardMarkup([
            [InlineKeyboardButton("📱 bKash", callback_data="usr_dep_method_bKash"), InlineKeyboardButton("📱 Nagad", callback_data="usr_dep_method_Nagad")],
            [InlineKeyboardButton("💎 Binance", callback_data="usr_dep_method_Binance")],
        ])
        await query.edit_message_text("নিচের কোন মাধ্যমে সেন্ড মানি করবেন সিলেক্ট করুন:", reply_markup=kb)
        return

    if data.startswith("usr_dep_method_"):
        method = data.split("_")[3]
        state, sdata = get_user_state(user_id)
        amount = sdata.get("amount", 20.0)

        num = ""
        if method == "bKash":
            num = get_setting("bkash_number")
        elif method == "Nagad":
            num = get_setting("nagad_number")
        elif method == "Binance":
            num = get_setting("binance_id")

        sdata["method"] = method
        set_user_state(user_id, "WAITING_DEP_TRX", sdata)

        text = (
            f"📱 <b>{method} পেমেন্ট নির্দেশিকা:</b>\n"
            f"─────────────────────────────\n"
            f"💰 ডিপোজিট পরিমাণ: <b>{amount:.2f} BDT</b>\n"
            f"📌 পেমেন্ট নাম্বার / আইডি (ক্লিক করে কপি করুন):\n"
            f"<code>{num}</code>\n\n"
            f"⚠️ <i>উপরের নাম্বারে <b>{amount:.2f} BDT</b> Send Money করার পর নিচে আপনার Transaction ID (TrxID) টি লিখে পাঠান:</i>"
        )
        await query.edit_message_text(text, parse_mode=ParseMode.HTML)
        return

    if data == "tool_2fa_prompt":
        set_user_state(user_id, "WAITING_2FA_KEY", {})
        await query.message.reply_text(
            "🔑 <b>2FA Authenticator:</b>\nআপনার 2FA Secret Key টি এখানে পেস্ট করে পাঠান:",
            parse_mode=ParseMode.HTML,
        )
        return

    if data == "prof_purchases":
        conn = get_db()
        c = conn.cursor()
        c.execute("SELECT id, receipt_no, product_name, total_bdt, order_date FROM orders WHERE user_id = ? ORDER BY id DESC LIMIT 10", (user_id,))
        orders = c.fetchall()
        conn.close()
        if not orders:
            await query.message.reply_text("🛒 আপনার পূর্বের কোনো কেনাকাটার হিস্টোরি পাওয়া যায়নি।")
            return

        txt = "🛒 <b>আপনার কেনাকাটার তালিকা:</b>\nবিস্তারিত দেখতে ও রিসিট পেতে অর্ডারের ওপর ক্লিক করুন:"
        kb = []
        for o in orders:
            btn_title = f"🧾 {o['receipt_no']} — {o['product_name']} ({o['total_bdt']:.0f} Tk)"
            kb.append([InlineKeyboardButton(btn_title, callback_data=f"prof_view_order_{o['id']}")])
        await query.message.reply_text(txt, parse_mode=ParseMode.HTML, reply_markup=InlineKeyboardMarkup(kb))
        return

    if data.startswith("prof_view_order_"):
        order_id = int(data.split("_")[3])
        conn = get_db()
        c = conn.cursor()
        c.execute("SELECT * FROM orders WHERE id = ? AND user_id = ?", (order_id, user_id))
        o = c.fetchone()
        conn.close()

        if not o:
            await query.message.reply_text("❌ অর্ডারটি পাওয়া যায়নি।")
            return

        creds_display = ""
        if o["delivered_data"]:
            try:
                items = json.loads(o["delivered_data"]) if o["delivered_data"].startswith("[") else [o["delivered_data"]]
            except Exception:
                items = [o["delivered_data"]]

            for itm in items:
                parts = [p.strip() for p in itm.split(":") if p.strip()]
                if len(parts) >= 2:
                    creds_display += f"👤 <code>{parts[0]}</code>\n🔐 <code>{parts[1]}</code>\n"
                else:
                    creds_display += f"🔑 <code>{itm}</code>\n"
        else:
            creds_display = "<i>⏳ অর্ডারটি প্রসেসিং হচ্ছে, খুব শীঘ্রই ডেলিভারি দেওয়া হবে।</i>"

        date_only = o["order_date"].split()[0] if o["order_date"] else "N/A"
        txt = (
            f"⚡ <b>Purchase Details</b>\n\n"
            f"Type  : <b>{o['category_name']}</b>\n"
            f"Date  : <code>{date_only}</code>\n"
            f"Total : <b>{o['total_bdt']:.2f} BDT</b>\n"
            f"🧾 Receipt : <code>{o['receipt_no']}</code>\n\n"
            f"<b>Credentials:</b>\n"
            f"{creds_display}"
        )
        kb = [
            [InlineKeyboardButton("🧾 Resend Receipt PDF", callback_data=f"prof_resend_pdf_{o['id']}")],
            [InlineKeyboardButton("< Back", callback_data="prof_purchases")],
        ]
        await query.message.reply_text(txt, parse_mode=ParseMode.HTML, reply_markup=InlineKeyboardMarkup(kb))
        return

    if data.startswith("prof_resend_pdf_"):
        order_id = int(data.split("_")[3])
        conn = get_db()
        c = conn.cursor()
        c.execute("SELECT * FROM orders WHERE id = ? AND user_id = ?", (order_id, user_id))
        o = c.fetchone()
        conn.close()

        if not o:
            await query.message.reply_text("অর্ডার পাওয়া যায়নি।")
            return

        bot_info = await get_bot(query).get_me()
        try:
            pdf_bio = generate_pdf_receipt(
                receipt_no=o["receipt_no"],
                date_str=o["order_date"],
                day_name=o["day_of_week"] or "N/A",
                customer_name=query.from_user.full_name,
                telegram_id=user_id,
                username=query.from_user.username or "",
                item_desc=o["product_name"],
                category_name=o["category_name"] or "GENERAL",
                qty=o["quantity"],
                unit_price=o["unit_price"] or o["total_bdt"],
                total_bdt=o["total_bdt"],
                bot_username=bot_info.username,
            )
            await query.message.reply_document(
                document=InputFile(pdf_bio, filename=f"Receipt_{o['receipt_no']}.pdf"),
                caption=f"🧾 <b>Official Payment Receipt:</b> <code>{o['receipt_no']}</code>",
                parse_mode=ParseMode.HTML,
            )
        except Exception as e:
            logger.error(f"Error resending receipt: {e}")
            await query.message.reply_text("❌ রিসিট তৈরি করতে সমস্যা হয়েছে।")
        return

    if data == "prof_deposit_history":
        conn = get_db()
        c = conn.cursor()
        c.execute("SELECT * FROM deposits WHERE user_id = ? ORDER BY id DESC LIMIT 10", (user_id,))
        deps = c.fetchall()
        conn.close()
        if not deps:
            await query.message.reply_text("🔄 আপনার কোনো ডিপোজিট হিস্টোরি পাওয়া যায়নি।")
            return
        msg = "🔄 <b>আপনার ডিপোজিট হিস্টোরি:</b>\n─────────────────────────────\n"
        for d in deps:
            status_icon = "⏳ Pending" if d['status'] == "pending" else ("✅ Approved" if d['status'] == "approved" else "❌ Rejected")
            msg += (
                f"💳 মেথড: {d['method']} | TrxID: <code>{d['trx_id']}</code>\n"
                f"💰 পরিমাণ: {d['amount']:.2f} BDT | অবস্থা: {status_icon}\n"
                f"📅 তারিখ: {d['created_at']}\n"
                f"─────────────────────────────\n"
            )
        await query.message.reply_text(msg, parse_mode=ParseMode.HTML)
        return

    if data == "prof_balance_details":
        rate = float(get_setting("usdt_rate", "125.0"))
        bdt = db_user['balance_bdt']
        usdt = bdt / rate if rate > 0 else 0.0
        msg = (
            f"💰 <b>ব্যালেন্স বিবরণী:</b>\n\n"
            f"💵 মূল ব্যালেন্স: <b>{bdt:.2f} BDT</b>\n"
            f"💎 ডলার সমমূল্য: <b>{usdt:.4f} USDT</b>\n"
            f"📈 এক্সচেঞ্জ রেট: 1 USDT = {rate} BDT"
        )
        await query.message.reply_text(msg, parse_mode=ParseMode.HTML)
        return

    if data == "prof_referral":
        bot_info = await context.bot.get_me()
        ref_link = f"https://t.me/{bot_info.username}?start=ref_{user_id}"
        msg = (
            f"👥 <b>রেফারাল প্রোগ্রাম</b>\n\n"
            f"আপনার বন্ধুদের আমন্ত্রণ জানান এবং স্পেশাল অফার উপভোগ করুন!\n\n"
            f"🔗 আপনার রেফার লিংক:\n<code>{ref_link}</code>"
        )
        await query.message.reply_text(msg, parse_mode=ParseMode.HTML)
        return

    if data.startswith("adm_"):
        if not is_admin(user_id):
            return
        try:
            await handle_admin_callbacks(query, data, context)
        except Exception as e:
            logger.error(f"Admin callback error ({data}): {e}", exc_info=True)
            try:
                await query.message.reply_text(f"❌ এডমিন প্যানেল এরর: <code>{e}</code>", parse_mode=ParseMode.HTML)
            except Exception:
                pass
        return

# ------------------------------------------------------------------------------
# 9. PRODUCT CARD & PURCHASE ENGINE
# ------------------------------------------------------------------------------
async def post_transaction_to_channel(bot, receipt_no: str, customer_name: str, telegram_id: int, product_name: str, quantity: int, total_bdt: float):
    bot_info = await bot.get_me()
    text = (
        f"🔔 <b>নতুন প্রোডাক্ট সেল হয়েছে!</b>\n"
        f"🧾 Receipt: {receipt_no}\n"
        f"👤 কাস্টমার: {customer_name} (ID: <code>{telegram_id}</code>)\n"
        f"📦 প্রোডাক্ট: {product_name}\n"
        f"🔢 পরিমাণ: {quantity} টি | মূল্য: {total_bdt:.2f} BDT"
    )
    kb = InlineKeyboardMarkup([
        [InlineKeyboardButton("🛒 Buy Products", url=f"https://t.me/{bot_info.username}")]
    ])
    ch_id = get_transaction_channel_id()
    try:
        await bot.send_message(
            chat_id=ch_id,
            text=text,
            parse_mode=ParseMode.HTML,
            reply_markup=kb,
        )
        return True, None
    except Exception as e:
        logger.error(f"Channel transaction post error (chat_id={ch_id}): {e}")
        return False, str(e)

async def render_purchase_summary(query_or_message, prod_id: int, quantity: int = 1, edit: bool = True):
    conn = get_db()
    c = conn.cursor()
    c.execute("""
    SELECT p.*, s.name as sub_name, s.category_id, c.name as cat_name
    FROM products p
    JOIN subcategories s ON p.subcategory_id = s.id
    JOIN categories c ON s.category_id = c.id
    WHERE p.id = ?
    """, (prod_id,))
    p = c.fetchone()
    conn.close()

    if not p:
        return

    rate = p["price_bdt"]
    total = rate * quantity
    full_name = f"{p['sub_name']} — {p['name']}"

    cat_low = p["cat_name"].lower()
    default_desc = (
        "⚡ হাই-স্পিড প্রিমিয়াম VPN অ্যাকাউন্ট — ১৫–৩০ মিনিটের মধ্যে ইনবক্সে Email ও Password পাবেন।" if "vpn" in cat_low
        else "🚀 আল্ট্রা-ফাস্ট ও ক্লিন রেসিডেনশিয়াল Proxy — হাই অ্যানোনিমিটি ও ১০০% ওয়ার্কিং গ্যারান্টি।" if "proxy" in cat_low
        else "🔐 রিকভারি অ্যাডেড হাই-সিকিউরিটি মেইল — সম্পূর্ণ নিরাপদ ও ইনস্ট্যান্ট ডেলিভারি।" if "recovery" in cat_low
        else "📧 ফ্রেশ ও ভেরিফাইড মেইল অ্যাকাউন্ট — ১০০% চেক করা স্টক ও ইনস্ট্যান্ট ডেলিভারি।"
    )
    desc_show = p["description"].strip() if p["description"] and str(p["description"]).strip() else default_desc

    text = (
        f"🧾 <b>PURCHASE SUMMARY</b>\n"
        f"─────────────────────────────\n"
        f"📁 <b>Category:</b> {p['cat_name']}\n"
        f"✉️ <b>Package:</b> {full_name}\n"
        f"📝 <b>Description:</b> <i>{desc_show}</i>\n"
        f"─────────────────────────────\n"
        f"✅ <b>Quantity:</b> {quantity}\n"
        f"💲 <b>Rate:</b> {format_price(rate)} BDT\n"
        f"🛒 <b>Total Price:</b> {format_price(total)} BDT"
    )

    kb = [
        [
            InlineKeyboardButton("➖", callback_data=f"calc_sub_{p['id']}_{quantity}"),
            InlineKeyboardButton(f"{quantity}", callback_data=f"calc_noop_{p['id']}_{quantity}"),
            InlineKeyboardButton("➕", callback_data=f"calc_add_{p['id']}_{quantity}"),
        ],
        [
            InlineKeyboardButton("✏️ Enter quantity", callback_data=f"calc_custom_{p['id']}")
        ],
        [
            InlineKeyboardButton("Back", callback_data=f"usr_sub_{p['subcategory_id']}"),
            InlineKeyboardButton("Cancel", callback_data="usr_root_cats"),
        ],
        [
            InlineKeyboardButton("Confirm", callback_data=f"calc_confirm_{p['id']}_{quantity}")
        ]
    ]
    reply_markup = InlineKeyboardMarkup(kb)

    if edit and hasattr(query_or_message, "edit_message_text"):
        try:
            await query_or_message.edit_message_text(text, parse_mode=ParseMode.HTML, reply_markup=reply_markup)
            return
        except Exception:
            pass

    if p["image_file_id"]:
        try:
            await query_or_message.reply_photo(photo=p["image_file_id"], caption=text, parse_mode=ParseMode.HTML, reply_markup=reply_markup)
            return
        except Exception:
            pass

    await query_or_message.reply_text(text, parse_mode=ParseMode.HTML, reply_markup=reply_markup)

async def process_order(query, prod_id: int, quantity: int):
    try:
        user_id = query.from_user.id
        user_name = query.from_user.full_name
        username = query.from_user.username or ""
        conn = get_db()
        c = conn.cursor()
        c.execute("""
        SELECT p.*, s.name as sub_name, c.name as cat_name
        FROM products p
        JOIN subcategories s ON p.subcategory_id = s.id
        JOIN categories c ON s.category_id = c.id
        WHERE p.id = ?
        """, (prod_id,))
        prod = c.fetchone()
        c.execute("SELECT balance_bdt FROM users WHERE user_id = ?", (user_id,))
        user_balance = c.fetchone()["balance_bdt"]

        if not prod:
            conn.close()
            await query.message.reply_text("প্রোডাক্ট পাওয়া যায়নি।")
            return

        total_bdt = prod["price_bdt"] * quantity
        total_usdt = prod["price_usdt"] * quantity

        if user_balance < total_bdt:
            conn.close()
            shortage = total_bdt - user_balance
            await query.message.reply_text(
                f"❌ <b>অপর্যাপ্ত ব্যালেন্স!</b>\n\n"
                f"প্রয়োজনীয় টাকা: {total_bdt:.2f} BDT\n"
                f"আপনার ব্যালেন্স: {user_balance:.2f} BDT\n"
                f"ঘাটতি: {shortage:.2f} BDT\n\n"
                f"দয়া করে নিচের 💳 Deposit মেনু থেকে ব্যালেন্স রিচার্জ করুন।",
                parse_mode=ParseMode.HTML,
            )
            return

        now = datetime.datetime.now()
        now_str = now.strftime("%Y-%m-%d %I:%M:%S %p")
        days_bn = {"Monday": "সোমবার", "Tuesday": "মঙ্গলবার", "Wednesday": "বুধবার", "Thursday": "বৃহস্পতিবার", "Friday": "শুক্রবার", "Saturday": "শনিবার", "Sunday": "রবিবার"}
        day_name = days_bn.get(now.strftime("%A"), now.strftime("%A"))
        receipt_no = f"RCP-{secrets.token_hex(5).upper()}"
        full_product_name = f"{prod['sub_name']} ({prod['name']})"

        c.execute("SELECT id, content FROM stock WHERE product_id = ? AND is_sold = 0 LIMIT ?", (prod_id, quantity))
        stock_rows = c.fetchall()

        is_vpn_or_proxy = ("vpn" in prod["cat_name"].lower()) or ("proxy" in prod["cat_name"].lower())
        if len(stock_rows) < quantity and not is_vpn_or_proxy:
            conn.close()
            await query.message.reply_text(
                f"❌ <b>দুঃখিত! পর্যাপ্ত স্টক নেই।</b>\n"
                f"বর্তমানে স্টকে আছে: <b>{len(stock_rows)}</b> টি। অনুগ্রহ করে পরিমাণ কমিয়ে দিন বা পরে চেষ্টা করুন।",
                parse_mode=ParseMode.HTML,
            )
            return

        if len(stock_rows) >= quantity:
            c.execute("UPDATE users SET balance_bdt = balance_bdt - ? WHERE user_id = ?", (total_bdt, user_id))
            stock_ids = [r["id"] for r in stock_rows]
            items_content = [r["content"] for r in stock_rows]
            c.execute(
                f"UPDATE stock SET is_sold = 1, sold_to = ?, sold_at = ? WHERE id IN ({','.join(['?']*len(stock_ids))})",
                [user_id, now_str] + stock_ids,
            )
            delivered_json = json.dumps(items_content)
            c.execute("""
            INSERT INTO orders (receipt_no, user_id, product_id, product_name, category_name, quantity, unit_price, total_bdt, total_usdt, status, delivered_data, order_date, day_of_week)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, 'completed', ?, ?, ?)
            """, (receipt_no, user_id, prod_id, full_product_name, prod["cat_name"], quantity, prod["price_bdt"], total_bdt, total_usdt, delivered_json, now_str, day_name))
            conn.commit()
            conn.close()

            bot_info = await get_bot(query).get_me()

            if quantity == 1:
                card_text = format_credentials_card(items_content, full_product_name, prod["cat_name"])
                await query.message.reply_text(card_text, parse_mode=ParseMode.HTML)
                if "MAIL" in prod["cat_name"].upper():
                    try:
                        excel_bio = generate_excel_delivery(items_content, full_product_name, receipt_no)
                        await query.message.reply_document(
                            document=InputFile(excel_bio, filename=f"{receipt_no}_{prod['name'].replace(' ', '_')}.xlsx"),
                            caption=f"📂 <b>Excel Data Backup:</b> <code>{receipt_no}</code>",
                            parse_mode=ParseMode.HTML,
                        )
                    except Exception:
                        pass
            else:
                bulk_msg = (
                    f"🎉 <b>আপনার অর্ডার সফলভাবে সম্পন্ন হয়েছে!</b>\n"
                    f"─────────────────────────────\n"
                    f"🧾 <b>Receipt:</b> <code>{receipt_no}</code>\n"
                    f"📦 <b>প্রোডাক্ট:</b> {full_product_name}\n"
                    f"🔢 <b>পরিমাণ:</b> {quantity} টি | <b>মোট পরিশোধ:</b> {format_price(total_bdt)} BDT\n"
                    f"─────────────────────────────\n"
                    f"📂 <b>আপনার সবকটি {quantity} টি অ্যাকাউন্ট/মেইল নিচের এক্সেল ফাইলে সুন্দরভাবে সাজিয়ে দেওয়া হলো।</b>"
                )
                await query.message.reply_text(bulk_msg, parse_mode=ParseMode.HTML)
                try:
                    excel_bio = generate_excel_delivery(items_content, full_product_name, receipt_no)
                    await query.message.reply_document(
                        document=InputFile(excel_bio, filename=f"{receipt_no}_{prod['name'].replace(' ', '_')}.xlsx"),
                        caption=f"📂 <b>Accounts Excel File:</b> <code>{receipt_no}</code> ({quantity} pcs)",
                        parse_mode=ParseMode.HTML,
                    )
                except Exception:
                    pass

            try:
                pdf_bio = generate_pdf_receipt(
                    receipt_no=receipt_no,
                    date_str=now_str,
                    day_name=day_name,
                    customer_name=user_name,
                    telegram_id=user_id,
                    username=username,
                    item_desc=full_product_name,
                    category_name=prod["cat_name"],
                    qty=quantity,
                    unit_price=prod["price_bdt"],
                    total_bdt=total_bdt,
                    bot_username=bot_info.username,
                )
                await query.message.reply_document(
                    document=InputFile(pdf_bio, filename=f"Receipt_{receipt_no}.pdf"),
                    caption=f"🧾 <b>Official Payment Receipt:</b> <code>{receipt_no}</code>",
                    parse_mode=ParseMode.HTML,
                )
            except Exception:
                pass

            try:
                await get_bot(query).send_message(
                    chat_id=SUPER_ADMIN_ID,
                    text=(
                        f"🔔 <b>নতুন প্রোডাক্ট সেল হয়েছে!</b>\n"
                        f"🧾 Receipt: <code>{receipt_no}</code>\n"
                        f"👤 কাস্টমার: {user_name} (ID: <code>{user_id}</code>)\n"
                        f"📦 প্রোডাক্ট: {full_product_name}\n"
                        f"🔢 পরিমাণ: {quantity} টি | মূল্য: {total_bdt:.2f} BDT"
                    ),
                    parse_mode=ParseMode.HTML,
                )
            except Exception:
                pass

            asyncio.create_task(
                post_transaction_to_channel(get_bot(query), receipt_no, user_name, user_id, full_product_name, quantity, total_bdt)
            )
            return

        c.execute("UPDATE users SET balance_bdt = balance_bdt - ? WHERE user_id = ?", (total_bdt, user_id))
        c.execute("""
        INSERT INTO orders (receipt_no, user_id, product_id, product_name, category_name, quantity, unit_price, total_bdt, total_usdt, status, order_date, day_of_week)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, 'pending_delivery', ?, ?)
        """, (receipt_no, user_id, prod_id, full_product_name, prod["cat_name"], quantity, prod["price_bdt"], total_bdt, total_usdt, now_str, day_name))
        order_id = c.lastrowid
        conn.commit()
        conn.close()

        bot_info = await get_bot(query).get_me()

        confirm_msg = (
            f"🎉 <b>আপনার অর্ডারটি সফলভাবে গ্রহণ করা হয়েছে!</b>\n"
            f"─────────────────────────────\n"
            f"🧾 <b>অর্ডার নম্বর:</b> #{order_id} (Receipt: <code>{receipt_no}</code>)\n"
            f"🛡️ <b>প্রোডাক্ট:</b> {full_product_name}\n"
            f"🔢 <b>পরিমাণ:</b> {quantity} টি | <b>মোট পরিশোধ:</b> {format_price(total_bdt)} BDT\n"
            f"📅 <b>তারিখ:</b> {now_str} ({day_name})\n"
            f"─────────────────────────────\n"
            f"⏳ <b>ডেলিভারি সময়:</b> আগামী ১৫ থেকে ৩০ মিনিটের মধ্যে এডমিন আপনার ইনবক্সে সার্ভিসটি ডেলিভারি পাঠাবে।"
        )
        await query.message.reply_text(confirm_msg, parse_mode=ParseMode.HTML)

        try:
            pdf_bio = generate_pdf_receipt(
                receipt_no=receipt_no,
                date_str=now_str,
                day_name=day_name,
                customer_name=user_name,
                telegram_id=user_id,
                username=username,
                item_desc=full_product_name,
                category_name=prod["cat_name"],
                qty=quantity,
                unit_price=prod["price_bdt"],
                total_bdt=total_bdt,
                bot_username=bot_info.username,
            )
            await query.message.reply_document(
                document=InputFile(pdf_bio, filename=f"Receipt_{receipt_no}.pdf"),
                caption=f"🧾 <b>Official Payment Receipt:</b> <code>{receipt_no}</code>",
                parse_mode=ParseMode.HTML,
            )
        except Exception:
            pass

        adm_kb = InlineKeyboardMarkup([
            [InlineKeyboardButton(f"🚀 Deliver Order #{order_id}", callback_data=f"adm_fulfill_{order_id}")],
            [InlineKeyboardButton(f"❌ Cancel & Refund #{order_id}", callback_data=f"adm_refund_{order_id}")]
        ])
        try:
            await get_bot(query).send_message(
                chat_id=SUPER_ADMIN_ID,
                text=(
                    f"🔔 <b>নতুন ম্যানুয়াল অর্ডার প্রয়োজন!</b>\n"
                    f"─────────────────────────────\n"
                    f"🧾 <b>অর্ডার নম্বর:</b> #{order_id} (Receipt: <code>{receipt_no}</code>)\n"
                    f"👤 <b>কাস্টমার:</b> {user_name} (ID: <code>{user_id}</code>)\n"
                    f"📦 <b>প্রোডাক্ট:</b> {full_product_name}\n"
                    f"🔢 <b>পরিমাণ:</b> {quantity} টি | <b>মূল্য:</b> {format_price(total_bdt)} BDT\n"
                    f"📅 <b>সময়:</b> {now_str}\n"
                    f"─────────────────────────────\n"
                    f"👇 ডেলিভারি দিতে নিচের বাটনে চাপুন:"
                ),
                parse_mode=ParseMode.HTML,
                reply_markup=adm_kb,
            )
        except Exception:
            pass
    except Exception as e:
        logger.error(f"Critical error in process_order: {e}", exc_info=True)
        try:
            await query.message.reply_text(f"❌ অর্ডার প্রসেস করতে সাময়িক সমস্যা হয়েছে: {e}")
        except Exception:
            pass

# ------------------------------------------------------------------------------
# 10. ADMIN PANEL (4-TIER NESTED MENUS FOR STOCK, BUTTONS, PHOTOS, USERS)
# ------------------------------------------------------------------------------
def build_admin_dashboard_markup(pend_dep: int, pend_ord: int) -> InlineKeyboardMarkup:
    status_tx = "🟢 Auto Channel Sales: ON" if _AUTO_SALES_ACTIVE else "🔴 Auto Channel Sales: OFF"
    return InlineKeyboardMarkup([
        [
            InlineKeyboardButton("📥 Add Stock", callback_data="adm_stk_step1_cats"),
            InlineKeyboardButton("🏷️ Manage Buttons & Prices", callback_data="adm_btn_step1_cats"),
        ],
        [
            InlineKeyboardButton("🖼️ Manage Product Photos", callback_data="adm_pho_step1_cats"),
            InlineKeyboardButton("💳 Payment Settings", callback_data="adm_pay_settings"),
        ],
        [
            InlineKeyboardButton(f"💳 Deposits ({pend_dep})", callback_data="adm_pending_deposits"),
            InlineKeyboardButton(f"🛡️ Orders ({pend_ord})", callback_data="adm_pending_orders"),
        ],
        [
            InlineKeyboardButton("👥 User Management", callback_data="adm_usr_list_0"),
            InlineKeyboardButton("📊 Shop Statistics", callback_data="adm_statistics"),
        ],
        [
            InlineKeyboardButton("🧾 Check / Download Receipt", callback_data="adm_check_receipt"),
            InlineKeyboardButton("📢 Broadcast (with Button)", callback_data="adm_broadcast_start"),
        ],
        [
            InlineKeyboardButton(status_tx, callback_data="adm_toggle_autosales"),
            InlineKeyboardButton("⚡ Channel TX Manager", callback_data="adm_tx_manager"),
        ],
        [
            InlineKeyboardButton("🎛️ Menu Layout Manager", callback_data="adm_menu_layout"),
            InlineKeyboardButton("🔄 Factory Reset / Clear Data", callback_data="adm_factory_reset_warn"),
        ]
    ])

async def admin_panel_menu_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    if not is_admin(user_id):
        return

    conn = get_db()
    c = conn.cursor()
    c.execute("SELECT COUNT(*) as cnt FROM deposits WHERE status = 'pending'")
    pend_dep = c.fetchone()["cnt"]
    c.execute("SELECT COUNT(*) as cnt FROM orders WHERE status = 'pending_delivery'")
    pend_ord = c.fetchone()["cnt"]
    conn.close()

    text = (
        f"👑 <b>সুপার এডমিন কন্ট্রোল ড্যাশবোর্ড</b>\n"
        f"─────────────────────────────\n"
        f"সব কাজ নেস্টেড বাটন দিয়ে ৩-৪ ধাপে পরিচালনা করুন:\n\n"
        f"💳 পেন্ডিং ডিপোজিট: <b>{pend_dep}</b> টি\n"
        f"🛡️ পেন্ডিং ম্যানুয়াল অর্ডার: <b>{pend_ord}</b> টি"
    )
    await update.message.reply_text(text, parse_mode=ParseMode.HTML, reply_markup=build_admin_dashboard_markup(pend_dep, pend_ord))

async def handle_admin_callbacks(query, data: str, context: ContextTypes.DEFAULT_TYPE):
    global _AUTO_SALES_ACTIVE
    user_id = query.from_user.id

    if data == "adm_tx_manager":
        st_label = "🟢 অটো ট্রানজ্যাকশন: চালু" if _AUTO_SALES_ACTIVE else "🔴 অটো ট্রানজ্যাকশন: বন্ধ"
        tx_txt = (
            "⚡ <b>চ্যানেল ট্রানজ্যাকশন ম্যানেজার</b>\n"
            "─────────────────────────────\n"
            f"📢 চ্যানেল আইডি: <code>{get_transaction_channel_id()}</code>\n"
            f"👥 গ্রুপ আইডি: <code>{get_welcome_group_id()}</code>\n"
            f"⚙️ অটো লুপ স্ট্যাটাস: <b>{'চালু আছে 🟢' if _AUTO_SALES_ACTIVE else 'বন্ধ আছে 🔴'}</b>\n\n"
            "এখান থেকে আপনি কাস্টম ট্রানজ্যাকশন তৈরি করতে বা আইডি/লিংক সেট করতে পারবেন:"
        )
        tx_kb = InlineKeyboardMarkup([
            [InlineKeyboardButton(st_label, callback_data="adm_toggle_autosales_mgr")],
            [InlineKeyboardButton("✍️ Custom Transaction (নিজে বানান)", callback_data="adm_custom_tx_start")],
            [InlineKeyboardButton("⚙️ Group & Channel ID/Link Settings", callback_data="adm_gc_settings")],
            [InlineKeyboardButton("< Back to Dashboard", callback_data="adm_root")],
        ])
        await query.edit_message_text(tx_txt, parse_mode=ParseMode.HTML, reply_markup=tx_kb)
        return

    if data == "adm_toggle_autosales_mgr":
        _AUTO_SALES_ACTIVE = not _AUTO_SALES_ACTIVE
        state_str = "চালু করা হয়েছে 🟢" if _AUTO_SALES_ACTIVE else "বন্ধ করা হয়েছে 🔴"
        await query.answer(f"চ্যানেল অটো ট্রানজ্যাকশন {state_str}!", show_alert=True)
        await handle_admin_callbacks(query, "adm_tx_manager", context)
        return

    if data == "adm_gc_settings":
        gc_txt = (
            "⚙️ <b>গ্রুপ ও চ্যানেল আইডি/লিংক সেটিংস</b>\n"
            "─────────────────────────────\n"
            f"👥 <b>Update Group ID:</b> <code>{get_welcome_group_id()}</code>\n"
            f"🔗 <b>Update Group Link:</b> {get_welcome_group_link()}\n\n"
            f"📢 <b>TX Channel ID:</b> <code>{get_transaction_channel_id()}</code>\n"
            f"🔗 <b>TX Channel Link:</b> {get_transaction_channel_link()}\n\n"
            "💡 <i>গ্রুপ বা চ্যানেলে গিয়ে <code>/id</code> লিখলেই সঠিক আইডি পেয়ে যাবেন।</i>"
        )
        gc_kb = InlineKeyboardMarkup([
            [
                InlineKeyboardButton("✏️ Edit Group ID", callback_data="adm_set_wgrp_id"),
                InlineKeyboardButton("🔗 Edit Group Link", callback_data="adm_set_wgrp_link"),
            ],
            [
                InlineKeyboardButton("✏️ Edit Channel ID", callback_data="adm_set_txch_id"),
                InlineKeyboardButton("🔗 Edit Channel Link", callback_data="adm_set_txch_link"),
            ],
            [InlineKeyboardButton("< Back to TX Manager", callback_data="adm_tx_manager")],
        ])
        await query.edit_message_text(gc_txt, parse_mode=ParseMode.HTML, reply_markup=gc_kb, disable_web_page_preview=True)
        return

    if data == "adm_set_wgrp_id":
        set_user_state(user_id, "ADMIN_WAITING_WGRP_ID", {})
        await query.message.reply_text("👥 আপনার <b>Update Group ID</b> লিখে পাঠান:", parse_mode=ParseMode.HTML)
        return

    if data == "adm_set_wgrp_link":
        set_user_state(user_id, "ADMIN_WAITING_WGRP_LINK", {})
        await query.message.reply_text("🔗 আপনার <b>Update Group ইনভাইট লিংক</b> লিখে পাঠান:", parse_mode=ParseMode.HTML)
        return

    if data == "adm_set_txch_id":
        set_user_state(user_id, "ADMIN_WAITING_TXCH_ID", {})
        await query.message.reply_text("📢 আপনার <b>Transaction Channel ID</b> লিখে পাঠান:", parse_mode=ParseMode.HTML)
        return

    if data == "adm_set_txch_link":
        set_user_state(user_id, "ADMIN_WAITING_TXCH_LINK", {})
        await query.message.reply_text("🔗 আপনার <b>Transaction Channel লিংক</b> লিখে পাঠান:", parse_mode=ParseMode.HTML)
        return

    if data == "adm_custom_tx_start":
        set_user_state(user_id, "ADMIN_WAITING_CUSTOM_TX_NAME", {})
        await query.message.reply_text("✍️ <b>কাস্টম ট্রানজ্যাকশন (ধাপ ১/৪):</b>\nকাস্টমারের নাম লিখে পাঠান:", parse_mode=ParseMode.HTML)
        return

    # Step-by-step Multi-Button Broadcast
    if data == "adm_broadcast_start":
        set_user_state(user_id, "ADMIN_WAITING_BCAST_MSG", {})
        await query.message.reply_text(
            "📢 <b>স্মার্ট ব্রডকাস্ট (ধাপ ১):</b>\n"
            "আপনার নোটিশ বা আপডেট মেসেজটি লিখে পাঠান (ছবি বা টেক্সট):",
            parse_mode=ParseMode.HTML,
        )
        return

    if data == "adm_bcast_skip_btn":
        sdata = get_user_state(user_id)[1]
        await execute_smart_broadcast(context.bot, query.message, sdata, btn_text=None, btn_url=None)
        set_user_state(user_id, None)
        return

    if data == "adm_bcast_add_btn":
        sdata = get_user_state(user_id)[1]
        set_user_state(user_id, "ADMIN_WAITING_BCAST_BTN_COUNT", sdata)
        await query.message.reply_text(
            "🔢 <b>আপনি কয়টি বাটন অ্যাড করতে চান?</b>\n"
            "সংখ্যা লিখুন (যেমন: <code>1</code>, <code>2</code>, <code>3</code>... সর্বোচ্চ ১০টি):",
            parse_mode=ParseMode.HTML,
        )
        return

    if data == "adm_check_receipt":
        set_user_state(user_id, "ADMIN_WAITING_RECEIPT_CODE", {})
        await query.message.reply_text("🧾 <b>রিসিপ্ট কোডটি লিখুন:</b>", parse_mode=ParseMode.HTML)
        return

    if data == "adm_toggle_autosales":
        _AUTO_SALES_ACTIVE = not _AUTO_SALES_ACTIVE
        state_str = "চালু করা হয়েছে 🟢" if _AUTO_SALES_ACTIVE else "বন্ধ করা হয়েছে 🔴"
        await query.answer(f"চ্যানেল অটো ট্রানজ্যাকশন {state_str}!", show_alert=True)
        data = "adm_root"

    if data == "adm_factory_reset_warn":
        warn_txt = (
            "⚠️ <b>সতর্কবার্তা: আপনি কি নিশ্চিত যে ফ্যাক্টরি রিসেট করবেন?</b>\n\n"
            "• সকল টেস্ট অর্ডার, স্টক ও ডিপোজিট মুছে যাবে এবং ব্যালেন্স ০ হবে।"
        )
        r_kb = InlineKeyboardMarkup([
            [InlineKeyboardButton("🔴 Yes, Wipe All Data", callback_data="adm_factory_reset_confirm")],
            [InlineKeyboardButton("❌ Cancel", callback_data="adm_root")]
        ])
        await query.edit_message_text(warn_txt, parse_mode=ParseMode.HTML, reply_markup=r_kb)
        return

    if data == "adm_factory_reset_confirm":
        conn = get_db()
        c = conn.cursor()
        c.execute("DELETE FROM orders")
        c.execute("DELETE FROM stock")
        c.execute("DELETE FROM deposits")
        c.execute("UPDATE users SET balance_bdt = 0")
        conn.commit()
        conn.close()
        await query.edit_message_text("✅ <b>ফ্যাক্টরি রিসেট সম্পন্ন!</b>", parse_mode=ParseMode.HTML)
        return

    if data == "adm_root":
        conn = get_db()
        c = conn.cursor()
        c.execute("SELECT COUNT(*) as cnt FROM deposits WHERE status = 'pending'")
        pend_dep = c.fetchone()["cnt"]
        c.execute("SELECT COUNT(*) as cnt FROM orders WHERE status = 'pending_delivery'")
        pend_ord = c.fetchone()["cnt"]
        conn.close()

        text = (
            f"👑 <b>সুপার এডমিন কন্ট্রোল ড্যাশবোর্ড</b>\n"
            f"─────────────────────────────\n"
            f"💳 পেন্ডিং ডিপোজিট: <b>{pend_dep}</b> টি\n"
            f"🛡️ পেন্ডিং ম্যানুয়াল অর্ডার: <b>{pend_ord}</b> টি"
        )
        await query.edit_message_text(text, parse_mode=ParseMode.HTML, reply_markup=build_admin_dashboard_markup(pend_dep, pend_ord))
        return

    if data == "adm_stk_step1_cats":
        conn = get_db()
        c = conn.cursor()
        c.execute("SELECT id, name FROM categories ORDER BY position ASC, id ASC")
        cats = c.fetchall()
        conn.close()
        kb = [[InlineKeyboardButton(cat["name"], callback_data=f"adm_stk_step2_sub_{cat['id']}")] for cat in cats]
        kb.append([InlineKeyboardButton("< Back to Dashboard", callback_data="adm_root")])
        await query.edit_message_text("📥 <b>[Add Stock — ধাপ ১]</b>\nক্যাটাগরি সিলেক্ট করুন:", parse_mode=ParseMode.HTML, reply_markup=InlineKeyboardMarkup(kb))
        return

    if data.startswith("adm_stk_step2_sub_"):
        cat_id = int(data.split("_")[4])
        conn = get_db()
        c = conn.cursor()
        c.execute("SELECT name FROM categories WHERE id = ?", (cat_id,))
        cat_row = c.fetchone()
        c.execute("SELECT id, name FROM subcategories WHERE category_id = ? ORDER BY position ASC, id ASC", (cat_id,))
        subs = c.fetchall()
        conn.close()

        cat_name = cat_row["name"] if cat_row else "Category"
        kb = []
        row = []
        for s in subs:
            row.append(InlineKeyboardButton(s["name"], callback_data=f"adm_stk_step3_prod_{s['id']}"))
            if len(row) == 2:
                kb.append(row)
                row = []
        if row:
            kb.append(row)
        kb.append([InlineKeyboardButton("< Back", callback_data="adm_stk_step1_cats")])
        await query.edit_message_text(f"📥 <b>[Add Stock — ধাপ ২]</b>\n<b>{cat_name}</b> ➔ ব্র্যান্ড সিলেক্ট করুন:", parse_mode=ParseMode.HTML, reply_markup=InlineKeyboardMarkup(kb))
        return

    if data.startswith("adm_stk_step3_prod_"):
        sub_id = int(data.split("_")[4])
        conn = get_db()
        c = conn.cursor()
        c.execute("""
        SELECT s.name as sub_name, s.category_id, c.name as cat_name
        FROM subcategories s
        JOIN categories c ON s.category_id = c.id
        WHERE s.id = ?
        """, (sub_id,))
        meta = c.fetchone()
        c.execute("SELECT id, name, price_bdt FROM products WHERE subcategory_id = ? ORDER BY id ASC", (sub_id,))
        prods = c.fetchall()
        conn.close()

        title = f"{meta['cat_name']} ➔ {meta['sub_name']}" if meta else "Select Variant"
        kb = [[InlineKeyboardButton(f"📦 {p['name']} ({format_price(p['price_bdt'])} Tk)", callback_data=f"adm_stk_step4_upload_{p['id']}")] for p in prods]
        back_cat = meta["category_id"] if meta else 1
        kb.append([InlineKeyboardButton("< Back", callback_data=f"adm_stk_step2_sub_{back_cat}")])
        await query.edit_message_text(f"📥 <b>[Add Stock — ধাপ ৩]</b>\n<b>{title}</b> ➔ প্যাকেজ সিলেক্ট করুন:", parse_mode=ParseMode.HTML, reply_markup=InlineKeyboardMarkup(kb))
        return

    if data.startswith("adm_stk_step4_upload_"):
        prod_id = int(data.split("_")[4])
        conn = get_db()
        c = conn.cursor()
        c.execute("SELECT p.name as prod_name, s.name as sub_name FROM products p JOIN subcategories s ON p.subcategory_id = s.id WHERE p.id = ?", (prod_id,))
        p = c.fetchone()
        conn.close()

        set_user_state(user_id, "ADMIN_WAITING_STOCK_INPUT", {"prod_id": prod_id})
        await query.message.reply_text(
            f"📥 <b>[Add Stock — ধাপ ৪: ডেটা আপলোড]</b>\n"
            f"📦 প্রোডাক্ট: <b>{p['sub_name']} ({p['prod_name']})</b>\n"
            f"─────────────────────────────\n"
            f"• <code>.xlsx</code> এক্সেল ফাইল পাঠান অথবা প্রতি লাইনে একটি করে অ্যাকাউন্ট/প্রক্সি পেস্ট করুন।",
            parse_mode=ParseMode.HTML,
        )
        return

    if data == "adm_btn_step1_cats":
        conn = get_db()
        c = conn.cursor()
        c.execute("SELECT id, name, position FROM categories ORDER BY position ASC, id ASC")
        cats = c.fetchall()
        conn.close()
        kb = []
        for idx, cat in enumerate(cats):
            up_btn = InlineKeyboardButton("⬆️", callback_data=f"adm_cat_up_{cat['id']}") if idx > 0 else InlineKeyboardButton("·", callback_data="adm_noop")
            down_btn = InlineKeyboardButton("⬇️", callback_data=f"adm_cat_down_{cat['id']}") if idx < len(cats) - 1 else InlineKeyboardButton("·", callback_data="adm_noop")
            kb.append([InlineKeyboardButton(f"📁 {cat['name']}", callback_data=f"adm_btn_step2_sub_{cat['id']}"), up_btn, down_btn])
        kb.append([InlineKeyboardButton("➕ Add New Category", callback_data="adm_btn_add_cat")])
        kb.append([InlineKeyboardButton("< Back to Dashboard", callback_data="adm_root")])
        await query.edit_message_text("🏷️ <b>[বাটন ও প্রোডাক্ট পরিচালনা — ধাপ ১]</b>", parse_mode=ParseMode.HTML, reply_markup=InlineKeyboardMarkup(kb))
        return

    if data.startswith("adm_btn_step2_sub_"):
        cat_id = int(data.split("_")[4])
        conn = get_db()
        c = conn.cursor()
        c.execute("SELECT name FROM categories WHERE id = ?", (cat_id,))
        cat_row = c.fetchone()
        c.execute("SELECT id, name, position FROM subcategories WHERE category_id = ? ORDER BY position ASC, id ASC", (cat_id,))
        subs = c.fetchall()
        conn.close()

        cat_name = cat_row["name"] if cat_row else "Category"
        kb = []
        for idx, s in enumerate(subs):
            up_btn = InlineKeyboardButton("⬆️", callback_data=f"adm_sub_up_{s['id']}") if idx > 0 else InlineKeyboardButton("·", callback_data="adm_noop")
            down_btn = InlineKeyboardButton("⬇️", callback_data=f"adm_sub_down_{s['id']}") if idx < len(subs) - 1 else InlineKeyboardButton("·", callback_data="adm_noop")
            kb.append([InlineKeyboardButton(f"📁 {s['name']}", callback_data=f"adm_btn_step3_prod_{s['id']}"), up_btn, down_btn])
        kb.append([InlineKeyboardButton(f"➕ Add New Brand in {cat_name}", callback_data=f"adm_btn_add_sub_{cat_id}")])
        kb.append([
            InlineKeyboardButton("✏️ Rename Category", callback_data=f"adm_ren_cat_{cat_id}"),
            InlineKeyboardButton("🗑️ Delete Category", callback_data=f"adm_del_cat_{cat_id}")
        ])
        kb.append([InlineKeyboardButton("< Back to Categories", callback_data="adm_btn_step1_cats")])
        await query.edit_message_text(f"🏷️ <b>{cat_name}</b> ➔ ব্র্যান্ড বেছে নিন:", parse_mode=ParseMode.HTML, reply_markup=InlineKeyboardMarkup(kb))
        return

    if data.startswith("adm_btn_step3_prod_"):
        sub_id = int(data.split("_")[4])
        conn = get_db()
        c = conn.cursor()
        c.execute("SELECT s.name as sub_name, s.category_id, c.name as cat_name FROM subcategories s JOIN categories c ON s.category_id = c.id WHERE s.id = ?", (sub_id,))
        meta = c.fetchone()
        c.execute("SELECT id, name, price_bdt FROM products WHERE subcategory_id = ? ORDER BY id ASC", (sub_id,))
        prods = c.fetchall()
        conn.close()

        title = f"{meta['cat_name']} ➔ {meta['sub_name']}" if meta else "Section"
        brand_name = meta["sub_name"] if meta else "Brand"
        kb = [[InlineKeyboardButton(f"📦 {p['name']} ({format_price(p['price_bdt'])} Tk)", callback_data=f"adm_btn_step4_edit_{p['id']}")] for p in prods]
        kb.append([InlineKeyboardButton(f"➕ Add New Duration / Package in {brand_name}", callback_data=f"adm_btn_add_prod_{sub_id}")])
        kb.append([
            InlineKeyboardButton(f"✏️ Rename Brand ({brand_name})", callback_data=f"adm_ren_sub_{sub_id}"),
            InlineKeyboardButton(f"🗑️ Delete Brand ({brand_name})", callback_data=f"adm_del_sub_{sub_id}")
        ])
        back_cat = meta["category_id"] if meta else 1
        kb.append([InlineKeyboardButton("< Back to Brands", callback_data=f"adm_btn_step2_sub_{back_cat}")])
        await query.edit_message_text(f"🏷️ <b>{title}</b>", parse_mode=ParseMode.HTML, reply_markup=InlineKeyboardMarkup(kb))
        return

    if data.startswith("adm_btn_step4_edit_"):
        prod_id = int(data.split("_")[4])
        conn = get_db()
        c = conn.cursor()
        c.execute("SELECT p.*, s.name as sub_name, s.category_id FROM products p JOIN subcategories s ON p.subcategory_id = s.id WHERE p.id = ?", (prod_id,))
        p = c.fetchone()
        conn.close()

        is_act = bool(p["is_active"] if "is_active" in p.keys() else 1)
        act_label = "🟢 Status: Active (Available)" if is_act else "🔴 Status: Inactive (Out of Stock)"
        photo_status = "🖼️ নিজস্ব ছবি সেট করা আছে" if p["image_file_id"] else "⚪ কোনো ছবি নেই"
        desc_txt = p["description"] if p["description"] else "ডিফল্ট বিবরণ চালু আছে"

        txt = (
            f"📦 <b>{p['sub_name']} — {p['name']}</b>\n"
            f"─────────────────────────────\n"
            f"⚡ <b>অবস্থা:</b> <b>{'Active 🟢' if is_act else 'Inactive 🔴'}</b>\n"
            f"💰 <b>বর্তমান দাম:</b> <b>{p['price_bdt']:.2f} BDT</b>\n"
            f"📸 <b>ফটো স্ট্যাটাস:</b> {photo_status}\n"
            f"📝 <b>বিবরণ:</b> <i>{desc_txt}</i>"
        )
        kb = [
            [InlineKeyboardButton(act_label, callback_data=f"adm_toggact_{p['id']}")],
            [InlineKeyboardButton("💰 Change Price (BDT)", callback_data=f"adm_chgprice_{p['id']}")],
            [InlineKeyboardButton("✏️ Rename Variant Name", callback_data=f"adm_rename_{p['id']}")],
            [InlineKeyboardButton("📝 Edit Custom Description", callback_data=f"adm_editdesc_{p['id']}")],
            [InlineKeyboardButton("📸 Change / Set Photo", callback_data=f"adm_setphoto_{p['id']}")],
            [InlineKeyboardButton("🗑️ Delete This Variant", callback_data=f"adm_delprod_{p['id']}")],
            [InlineKeyboardButton("< Back", callback_data=f"adm_btn_step3_prod_{p['subcategory_id']}")],
        ]
        await query.edit_message_text(txt, parse_mode=ParseMode.HTML, reply_markup=InlineKeyboardMarkup(kb))
        return

    if data.startswith("adm_toggact_"):
        prod_id = int(data.split("_")[2])
        conn = get_db()
        c = conn.cursor()
        c.execute("SELECT COALESCE(is_active, 1) as is_active FROM products WHERE id = ?", (prod_id,))
        pr = c.fetchone()
        if pr:
            new_st = 0 if pr["is_active"] else 1
            c.execute("UPDATE products SET is_active = ? WHERE id = ?", (new_st, prod_id))
            conn.commit()
        conn.close()
        await handle_admin_callbacks(query, f"adm_btn_step4_edit_{prod_id}", context)
        return

    if data == "adm_menu_layout":
        layout = get_menu_layout()
        txt = "🎛️ <b>মেইন মেনু বাটন লেআউট ম্যানেজার</b>\n\n"
        kb = []
        for r_idx, row in enumerate(layout):
            txt += f"<b>Layer {r_idx+1}:</b> " + " | ".join(row) + "\n"
            kb.append([InlineKeyboardButton(f"✏️ {lbl}", callback_data=f"adm_ml_sel_{r_idx}_{c_idx}") for c_idx, lbl in enumerate(row)])
        kb.append([InlineKeyboardButton("🔄 Reset Default Layout", callback_data="adm_ml_reset")])
        kb.append([InlineKeyboardButton("< Back to Dashboard", callback_data="adm_root")])
        await query.edit_message_text(txt, parse_mode=ParseMode.HTML, reply_markup=InlineKeyboardMarkup(kb))
        return

    if data == "adm_ml_reset":
        save_menu_layout([list(r) for r in DEFAULT_MENU_LAYOUT])
        await query.message.reply_text("✅ ডিফল্ট লেআউট রিসেট করা হয়েছে!", reply_markup=get_main_reply_keyboard(user_id))
        await handle_admin_callbacks(query, "adm_menu_layout", context)
        return

    if data.startswith("adm_ml_sel_"):
        parts = data.split("_")
        r_idx, c_idx = int(parts[3]), int(parts[4])
        layout = get_menu_layout()
        if r_idx >= len(layout) or c_idx >= len(layout[r_idx]):
            await handle_admin_callbacks(query, "adm_menu_layout", context)
            return
        lbl = layout[r_idx][c_idx]
        kb = [
            [InlineKeyboardButton("⬆️ Move Up", callback_data=f"adm_ml_mv_{r_idx}_{c_idx}_up"), InlineKeyboardButton("⬇️ Move Down", callback_data=f"adm_ml_mv_{r_idx}_{c_idx}_down")],
            [InlineKeyboardButton("⬅️ Move Left", callback_data=f"adm_ml_mv_{r_idx}_{c_idx}_left"), InlineKeyboardButton("➡️ Move Right", callback_data=f"adm_ml_mv_{r_idx}_{c_idx}_right")],
            [InlineKeyboardButton("↔️ Separate Row", callback_data=f"adm_ml_mv_{r_idx}_{c_idx}_solo")],
            [InlineKeyboardButton("< Back to Layout", callback_data="adm_menu_layout")],
        ]
        await query.edit_message_text(f"🎛️ <b>বাটন এডিটর:</b> <code>{lbl}</code>", parse_mode=ParseMode.HTML, reply_markup=InlineKeyboardMarkup(kb))
        return

    if data.startswith("adm_ml_mv_"):
        parts = data.split("_")
        r_idx, c_idx, direction = int(parts[3]), int(parts[4]), parts[5]
        layout = get_menu_layout()
        if r_idx < len(layout) and c_idx < len(layout[r_idx]):
            lbl = layout[r_idx][c_idx]
            if direction == "left" and c_idx > 0:
                layout[r_idx][c_idx], layout[r_idx][c_idx - 1] = layout[r_idx][c_idx - 1], layout[r_idx][c_idx]
            elif direction == "right" and c_idx < len(layout[r_idx]) - 1:
                layout[r_idx][c_idx], layout[r_idx][c_idx + 1] = layout[r_idx][c_idx + 1], layout[r_idx][c_idx]
            elif direction == "up":
                layout[r_idx].pop(c_idx)
                if r_idx > 0:
                    layout[r_idx - 1].append(lbl)
                else:
                    layout.insert(0, [lbl])
            elif direction == "down":
                layout[r_idx].pop(c_idx)
                if r_idx + 1 < len(layout):
                    layout[r_idx + 1].append(lbl)
                else:
                    layout.append([lbl])
            elif direction == "solo" and len(layout[r_idx]) > 1:
                layout[r_idx].pop(c_idx)
                layout.insert(r_idx + 1, [lbl])
            save_menu_layout(layout)
            await query.message.reply_text(f"✅ <b>{lbl}</b> আপডেট করা হয়েছে!", parse_mode=ParseMode.HTML, reply_markup=get_main_reply_keyboard(user_id))
        await handle_admin_callbacks(query, "adm_menu_layout", context)
        return

    if data.startswith("adm_cat_up_") or data.startswith("adm_cat_down_"):
        is_up = "up" in data
        target_id = int(data.split("_")[3])
        conn = get_db()
        c = conn.cursor()
        c.execute("SELECT id, position FROM categories ORDER BY position ASC, id ASC")
        all_cats = c.fetchall()
        for idx, ct in enumerate(all_cats):
            if ct["id"] == target_id:
                swap_idx = idx - 1 if is_up else idx + 1
                if 0 <= swap_idx < len(all_cats):
                    other_id = all_cats[swap_idx]["id"]
                    p1, p2 = idx, swap_idx
                    c.execute("UPDATE categories SET position = ? WHERE id = ?", (p2, target_id))
                    c.execute("UPDATE categories SET position = ? WHERE id = ?", (p1, other_id))
                    conn.commit()
                break
        conn.close()
        await handle_admin_callbacks(query, "adm_btn_step1_cats", context)
        return

    if data.startswith("adm_sub_up_") or data.startswith("adm_sub_down_"):
        is_up = "up" in data
        target_id = int(data.split("_")[3])
        conn = get_db()
        c = conn.cursor()
        c.execute("SELECT category_id FROM subcategories WHERE id = ?", (target_id,))
        p_row = c.fetchone()
        parent_cat_id = p_row["category_id"] if p_row else 1
        c.execute("SELECT id, position FROM subcategories WHERE category_id = ? ORDER BY position ASC, id ASC", (parent_cat_id,))
        all_subs = c.fetchall()
        for idx, sb in enumerate(all_subs):
            if sb["id"] == target_id:
                swap_idx = idx - 1 if is_up else idx + 1
                if 0 <= swap_idx < len(all_subs):
                    other_id = all_subs[swap_idx]["id"]
                    c.execute("UPDATE subcategories SET position = ? WHERE id = ?", (swap_idx, target_id))
                    c.execute("UPDATE subcategories SET position = ? WHERE id = ?", (idx, other_id))
                    conn.commit()
                break
        conn.close()
        await handle_admin_callbacks(query, f"adm_btn_step2_sub_{parent_cat_id}", context)
        return

    if data == "adm_noop":
        return

    if data.startswith("adm_ren_sub_"):
        sub_id = int(data.split("_")[3])
        set_user_state(user_id, "ADMIN_WAITING_REN_SUB", {"sub_id": sub_id})
        await query.message.reply_text("✏️ <b>ব্র্যান্ডের নতুন নাম লিখে পাঠান:</b>", parse_mode=ParseMode.HTML)
        return

    if data.startswith("adm_del_sub_"):
        sub_id = int(data.split("_")[3])
        conn = get_db()
        c = conn.cursor()
        c.execute("SELECT name FROM subcategories WHERE id = ?", (sub_id,))
        sub_row = c.fetchone()
        if sub_row:
            c.execute("DELETE FROM stock WHERE product_id IN (SELECT id FROM products WHERE subcategory_id = ?)", (sub_id,))
            c.execute("DELETE FROM products WHERE subcategory_id = ?", (sub_id,))
            c.execute("DELETE FROM subcategories WHERE id = ?", (sub_id,))
            conn.commit()
            await query.message.reply_text(f"🗑️ ব্র্যান্ড '{sub_row['name']}' মুছে ফেলা হয়েছে!")
        conn.close()
        return

    if data.startswith("adm_ren_cat_"):
        cat_id = int(data.split("_")[3])
        set_user_state(user_id, "ADMIN_WAITING_REN_CAT", {"cat_id": cat_id})
        await query.message.reply_text("✏️ <b>ক্যাটাগরির নতুন নাম লিখে পাঠান:</b>", parse_mode=ParseMode.HTML)
        return

    if data.startswith("adm_del_cat_"):
        cat_id = int(data.split("_")[3])
        conn = get_db()
        c = conn.cursor()
        c.execute("SELECT name FROM categories WHERE id = ?", (cat_id,))
        cat_row = c.fetchone()
        if cat_row:
            c.execute("DELETE FROM stock WHERE product_id IN (SELECT p.id FROM products p JOIN subcategories s ON p.subcategory_id = s.id WHERE s.category_id = ?)", (cat_id,))
            c.execute("DELETE FROM products WHERE subcategory_id IN (SELECT id FROM subcategories WHERE category_id = ?)", (cat_id,))
            c.execute("DELETE FROM subcategories WHERE category_id = ?", (cat_id,))
            c.execute("DELETE FROM categories WHERE id = ?", (cat_id,))
            conn.commit()
            await query.message.reply_text(f"🗑️ ক্যাটাগরি '{cat_row['name']}' মুছে ফেলা হয়েছে!")
        conn.close()
        return

    if data.startswith("adm_chgprice_"):
        prod_id = int(data.split("_")[2])
        set_user_state(user_id, "ADMIN_WAITING_PRICE", {"prod_id": prod_id})
        await query.message.reply_text("নতুন BDT দাম লিখুন (যেমন: 0.87 বা 15):")
        return

    if data.startswith("adm_rename_"):
        prod_id = int(data.split("_")[2])
        set_user_state(user_id, "ADMIN_WAITING_NAME", {"prod_id": prod_id})
        await query.message.reply_text("নতুন নাম বা মেয়াদ লিখুন:")
        return

    if data.startswith("adm_editdesc_"):
        prod_id = int(data.split("_")[2])
        set_user_state(user_id, "ADMIN_WAITING_DESC", {"prod_id": prod_id})
        await query.message.reply_text("📝 কাস্টম বিবরণ লিখে পাঠান:")
        return

    if data.startswith("adm_delprod_"):
        prod_id = int(data.split("_")[2])
        conn = get_db()
        c = conn.cursor()
        c.execute("DELETE FROM products WHERE id = ?", (prod_id,))
        conn.commit()
        conn.close()
        await query.edit_message_text("✅ প্রোডাক্টটি মুছে ফেলা হয়েছে।")
        return

    if data == "adm_btn_add_cat":
        set_user_state(user_id, "ADMIN_WAITING_NEW_CAT_NAME", {})
        await query.message.reply_text("➕ নতুন ক্যাটাগরির নাম লিখুন:")
        return

    if data.startswith("adm_btn_add_sub_"):
        cat_id = int(data.split("_")[4])
        set_user_state(user_id, "ADMIN_WAITING_NEW_SUB_NAME", {"cat_id": cat_id})
        await query.message.reply_text("➕ নতুন ব্র্যান্ডের নাম লিখুন:")
        return

    if data.startswith("adm_btn_add_prod_"):
        sub_id = int(data.split("_")[4])
        set_user_state(user_id, "ADMIN_WAITING_ADDPROD_NAME", {"sub_id": sub_id})
        await query.message.reply_text("➕ নতুন মেয়াদ বা প্যাকেজের নাম লিখুন:")
        return

    if data == "adm_pho_step1_cats":
        conn = get_db()
        c = conn.cursor()
        c.execute("SELECT id, name FROM categories ORDER BY position ASC, id ASC")
        cats = c.fetchall()
        conn.close()
        kb = [[InlineKeyboardButton(cat["name"], callback_data=f"adm_pho_step2_sub_{cat['id']}")] for cat in cats]
        kb.append([InlineKeyboardButton("< Back to Dashboard", callback_data="adm_root")])
        await query.edit_message_text("🖼️ <b>[প্রোডাক্ট ফটো ম্যানেজার — ধাপ ১]</b>", parse_mode=ParseMode.HTML, reply_markup=InlineKeyboardMarkup(kb))
        return

    if data.startswith("adm_pho_step2_sub_"):
        cat_id = int(data.split("_")[4])
        conn = get_db()
        c = conn.cursor()
        c.execute("SELECT id, name FROM subcategories WHERE category_id = ? ORDER BY position ASC, id ASC", (cat_id,))
        subs = c.fetchall()
        conn.close()
        kb = [[InlineKeyboardButton(f"📁 {s['name']}", callback_data=f"adm_pho_step3_prod_{s['id']}")] for s in subs]
        kb.append([InlineKeyboardButton("< Back", callback_data="adm_pho_step1_cats")])
        await query.edit_message_text("🖼️ <b>[প্রোডাক্ট ফটো ম্যানেজার — ধাপ ২]</b>", parse_mode=ParseMode.HTML, reply_markup=InlineKeyboardMarkup(kb))
        return

    if data.startswith("adm_pho_step3_prod_"):
        sub_id = int(data.split("_")[4])
        conn = get_db()
        c = conn.cursor()
        c.execute("SELECT id, name, image_file_id FROM products WHERE subcategory_id = ? ORDER BY id ASC", (sub_id,))
        prods = c.fetchall()
        conn.close()
        kb = [[InlineKeyboardButton(f"{'🖼️' if p['image_file_id'] else '⚪'} {p['name']}", callback_data=f"adm_pho_step4_view_{p['id']}")] for p in prods]
        kb.append([InlineKeyboardButton("< Back", callback_data="adm_pho_step1_cats")])
        await query.edit_message_text("🖼️ <b>[প্রোডাক্ট ফটো ম্যানেজার — ধাপ ৩]</b>", parse_mode=ParseMode.HTML, reply_markup=InlineKeyboardMarkup(kb))
        return

    if data.startswith("adm_pho_step4_view_"):
        prod_id = int(data.split("_")[4])
        conn = get_db()
        c = conn.cursor()
        c.execute("SELECT p.*, s.name as sub_name FROM products p JOIN subcategories s ON p.subcategory_id = s.id WHERE p.id = ?", (prod_id,))
        p = c.fetchone()
        conn.close()
        status = "✅ নিজস্ব ছবি আপলোড করা আছে" if p["image_file_id"] else "❌ কোনো ছবি নেই"
        kb = [[InlineKeyboardButton("📸 Upload / Change Photo", callback_data=f"adm_setphoto_{p['id']}")]]
        if p["image_file_id"]:
            kb.append([InlineKeyboardButton("🗑️ Remove Photo", callback_data=f"adm_delphoto_{p['id']}")])
        kb.append([InlineKeyboardButton("< Back", callback_data=f"adm_pho_step3_prod_{p['subcategory_id']}")])
        await query.edit_message_text(f"🖼️ <b>{p['sub_name']} — {p['name']}</b>\n\nঅবস্থা: {status}", parse_mode=ParseMode.HTML, reply_markup=InlineKeyboardMarkup(kb))
        return

    if data.startswith("adm_setphoto_"):
        prod_id = int(data.split("_")[2])
        set_user_state(user_id, "ADMIN_WAITING_PHOTO_UPLOAD", {"prod_id": prod_id})
        await query.message.reply_text("📸 এই প্রোডাক্টের জন্য ছবিটি টেলিগ্রামের চ্যাটে ফটো আকারে পাঠান:")
        return

    if data.startswith("adm_delphoto_"):
        prod_id = int(data.split("_")[2])
        conn = get_db()
        c = conn.cursor()
        c.execute("UPDATE products SET image_file_id = NULL WHERE id = ?", (prod_id,))
        conn.commit()
        conn.close()
        await query.edit_message_text("✅ ছবি মুছে ফেলা হয়েছে।")
        return

    if data == "adm_pay_settings":
        bkash = get_setting("bkash_number")
        nagad = get_setting("nagad_number")
        binance = get_setting("binance_id")
        rate = get_setting("usdt_rate", "125.0")
        min_man = get_setting("min_deposit_manual", "20.0")

        text = (
            f"💳 <b>পেমেন্ট সেটিংস:</b>\n"
            f"─────────────────────────────\n"
            f"📱 <b>bKash:</b> <code>{bkash}</code>\n"
            f"📱 <b>Nagad:</b> <code>{nagad}</code>\n"
            f"💎 <b>Binance:</b> <code>{binance}</code>\n"
            f"💰 <b>সর্বনিম্ন ডিপোজিট:</b> <b>{float(min_man):.0f} BDT</b>\n"
            f"💵 <b>USDT রেট:</b> 1 USDT = {rate} BDT"
        )
        kb = InlineKeyboardMarkup([
            [InlineKeyboardButton("📱 Edit bKash Number", callback_data="adm_set_bkash")],
            [InlineKeyboardButton("📱 Edit Nagad Number", callback_data="adm_set_nagad")],
            [InlineKeyboardButton("💎 Edit Binance Info", callback_data="adm_set_binance")],
            [InlineKeyboardButton("💰 Edit Min Deposit", callback_data="adm_set_mindep_manual")],
            [InlineKeyboardButton("💵 Edit USDT Rate", callback_data="adm_set_rate")],
            [InlineKeyboardButton("< Back to Dashboard", callback_data="adm_root")],
        ])
        await query.edit_message_text(text, parse_mode=ParseMode.HTML, reply_markup=kb)
        return

    if data == "adm_set_mindep_manual":
        set_user_state(user_id, "ADMIN_WAITING_MINDEP_MANUAL", {})
        await query.message.reply_text("📱 সর্বনিম্ন ডিপোজিট কত টাকা হবে লিখুন:")
        return

    if data == "adm_set_bkash":
        set_user_state(user_id, "ADMIN_WAITING_BKASH", {})
        await query.message.reply_text("নতুন বিকাশ নাম্বার লিখুন:")
        return

    if data == "adm_set_nagad":
        set_user_state(user_id, "ADMIN_WAITING_NAGAD", {})
        await query.message.reply_text("নতুন নগদ নাম্বার লিখুন:")
        return

    if data == "adm_set_binance":
        set_user_state(user_id, "ADMIN_WAITING_BINANCE", {})
        await query.message.reply_text("নতুন বাইনান্স তথ্য লিখুন:")
        return

    if data == "adm_set_rate":
        set_user_state(user_id, "ADMIN_WAITING_RATE", {})
        await query.message.reply_text("নতুন USDT রেট লিখুন:")
        return

    if data.startswith("adm_usr_list_"):
        page = int(data.split("_")[3])
        limit = 8
        offset = page * limit
        conn = get_db()
        c = conn.cursor()
        c.execute("SELECT COUNT(*) as total FROM users")
        total_u = c.fetchone()["total"]
        c.execute("SELECT user_id, name, balance_bdt FROM users ORDER BY user_id DESC LIMIT ? OFFSET ?", (limit, offset))
        u_list = c.fetchall()
        conn.close()

        kb = [[InlineKeyboardButton(f"👤 {(u['name'] or 'User')[:16]} ({u['balance_bdt']:.0f} Tk)", callback_data=f"adm_usr_view_{u['user_id']}")] for u in u_list]
        nav_row = []
        if page > 0:
            nav_row.append(InlineKeyboardButton("◀️ Previous", callback_data=f"adm_usr_list_{page - 1}"))
        if (offset + limit) < total_u:
            nav_row.append(InlineKeyboardButton("Next ▶️", callback_data=f"adm_usr_list_{page + 1}"))
        if nav_row:
            kb.append(nav_row)
        kb.append([InlineKeyboardButton("< Back to Dashboard", callback_data="adm_root")])
        await query.edit_message_text(f"👥 <b>মোট ইউজার: {total_u} জন</b>", parse_mode=ParseMode.HTML, reply_markup=InlineKeyboardMarkup(kb))
        return

    if data.startswith("adm_usr_view_"):
        target_uid = int(data.split("_")[3])
        conn = get_db()
        c = conn.cursor()
        c.execute("SELECT * FROM users WHERE user_id = ?", (target_uid,))
        u = c.fetchone()
        conn.close()
        if not u:
            return
        info = f"👤 <b>{u['name']}</b>\n🆔 ID: <code>{u['user_id']}</code>\n💰 ব্যালেন্স: <b>{u['balance_bdt']:.2f} BDT</b>"
        ukb = InlineKeyboardMarkup([
            [InlineKeyboardButton("➕ Add Balance", callback_data=f"adm_usr_addbal_{u['user_id']}"), InlineKeyboardButton("➖ Deduct Balance", callback_data=f"adm_usr_dedbal_{u['user_id']}")],
            [InlineKeyboardButton("💬 Direct Message", url=f"tg://user?id={u['user_id']}")],
            [InlineKeyboardButton("< Back", callback_data="adm_usr_list_0")],
        ])
        await query.edit_message_text(info, parse_mode=ParseMode.HTML, reply_markup=ukb)
        return

    if data.startswith("adm_usr_addbal_"):
        target_uid = int(data.split("_")[3])
        set_user_state(user_id, "ADMIN_WAITING_ADDBAL_AMT", {"target_uid": target_uid})
        await query.message.reply_text(f"💰 ইউজার <code>{target_uid}</code>-কে কত টাকা যোগ করবেন লিখুন:", parse_mode=ParseMode.HTML)
        return

    if data.startswith("adm_usr_dedbal_"):
        target_uid = int(data.split("_")[3])
        set_user_state(user_id, "ADMIN_WAITING_DEDBAL_AMT", {"target_uid": target_uid})
        await query.message.reply_text(f"💸 ইউজার <code>{target_uid}</code>-এর কত টাকা কাটবেন লিখুন:", parse_mode=ParseMode.HTML)
        return

    if data == "adm_pending_deposits":
        conn = get_db()
        c = conn.cursor()
        c.execute("SELECT d.*, u.name as uname FROM deposits d JOIN users u ON d.user_id = u.user_id WHERE d.status = 'pending'")
        deps = c.fetchall()
        conn.close()
        if not deps:
            await query.message.reply_text("✅ বর্তমানে কোনো পেন্ডিং ডিপোজিট নেই।")
            return
        for d in deps:
            kb = InlineKeyboardMarkup([[InlineKeyboardButton("✅ Approve", callback_data=f"adm_dep_approve_{d['id']}"), InlineKeyboardButton("❌ Reject", callback_data=f"adm_dep_reject_{d['id']}")]])
            await query.message.reply_text(f"💳 <b>Deposit #{d['id']}</b>\n👤 {d['uname']} (<code>{d['user_id']}</code>)\n💰 {d['amount']:.2f} BDT ({d['method']})\n🔢 TrxID: <code>{d['trx_id']}</code>", parse_mode=ParseMode.HTML, reply_markup=kb)
        return

    if data.startswith("adm_dep_"):
        parts = data.split("_")
        action = parts[2]
        dep_id = int(parts[3])
        conn = get_db()
        c = conn.cursor()
        c.execute("SELECT * FROM deposits WHERE id = ?", (dep_id,))
        dep = c.fetchone()
        if not dep:
            conn.close()
            return
        if action == "approve":
            c.execute("UPDATE deposits SET status = 'approved' WHERE id = ?", (dep_id,))
            c.execute("UPDATE users SET balance_bdt = balance_bdt + ? WHERE user_id = ?", (dep["amount"], dep["user_id"]))
            conn.commit()
            conn.close()
            add_notification(dep["user_id"], "✅ Deposit Approved", f"+{dep['amount']:.2f} BDT added to your wallet!")
            await query.edit_message_text(f"✅ ডিপোজিট #{dep_id} অনুমোদিত হয়েছে।")
            try:
                await context.bot.send_message(chat_id=dep["user_id"], text=f"🎉 <b>ডিপোজিট সফল!</b>\nআপনার অ্যাকাউন্টে {dep['amount']:.2f} BDT যুক্ত হয়েছে।", parse_mode=ParseMode.HTML)
            except Exception:
                pass
        elif action == "reject":
            c.execute("UPDATE deposits SET status = 'rejected' WHERE id = ?", (dep_id,))
            conn.commit()
            conn.close()
            await query.edit_message_text(f"❌ ডিপোজিট #{dep_id} বাতিল করা হয়েছে।")
        return

    if data == "adm_pending_orders":
        conn = get_db()
        c = conn.cursor()
        c.execute("SELECT o.*, u.name as uname FROM orders o JOIN users u ON o.user_id = u.user_id WHERE o.status = 'pending_delivery'")
        ords = c.fetchall()
        conn.close()
        if not ords:
            await query.message.reply_text("✅ বর্তমানে কোনো পেন্ডিং অর্ডার নেই।")
            return
        for o in ords:
            kb = InlineKeyboardMarkup([
                [InlineKeyboardButton(f"🚀 Deliver Order #{o['id']}", callback_data=f"adm_fulfill_{o['id']}")],
                [InlineKeyboardButton(f"❌ Cancel & Refund #{o['id']}", callback_data=f"adm_refund_{o['id']}")]
            ])
            await query.message.reply_text(f"🛡️ <b>Pending Order #{o['id']}</b>\n👤 {o['uname']} (<code>{o['user_id']}</code>)\n📦 {o['product_name']} (x{o['quantity']})\n💰 {o['total_bdt']:.2f} BDT", parse_mode=ParseMode.HTML, reply_markup=kb)
        return

    if data.startswith("adm_fulfill_"):
        order_id = int(data.split("_")[2])
        conn = get_db()
        c = conn.cursor()
        c.execute("SELECT category_name FROM orders WHERE id = ?", (order_id,))
        o_row = c.fetchone()
        conn.close()
        cat_n = (o_row["category_name"] if o_row else "").lower()
        if "vpn" in cat_n:
            set_user_state(user_id, "ADMIN_WAITING_VPN_EMAIL", {"order_id": order_id})
            await query.message.reply_text(
                f"🚀 <b>VPN Order #{order_id} ডেলিভারি (ধাপ ১/২):</b>\n"
                f"📧 কাস্টমারকে পাঠানোর জন্য VPN-এর <b>Email / Username</b> লিখে পাঠান:",
                parse_mode=ParseMode.HTML,
            )
        else:
            set_user_state(user_id, "ADMIN_WAITING_MANUAL_DELIVERY", {"order_id": order_id})
            await query.message.reply_text(
                f"🚀 <b>Order #{order_id} ডেলিভারি:</b>\n"
                f"🌐 কাস্টমারকে পাঠানোর জন্য <b>Proxy String</b> (যেমন: <code>host:port:user:pass</code>) পেস্ট করে পাঠান:",
                parse_mode=ParseMode.HTML,
            )
        return

    if data.startswith("adm_refund_"):
        order_id = int(data.split("_")[2])
        conn = get_db()
        c = conn.cursor()
        c.execute("SELECT * FROM orders WHERE id = ?", (order_id,))
        o = c.fetchone()
        if o and o["status"] == "pending_delivery":
            c.execute("UPDATE orders SET status = 'refunded' WHERE id = ?", (order_id,))
            c.execute("UPDATE users SET balance_bdt = balance_bdt + ? WHERE user_id = ?", (o["total_bdt"], o["user_id"]))
            conn.commit()
            await query.edit_message_text(f"❌ Order #{order_id} বাতিল ও রিফান্ড করা হয়েছে।")
        conn.close()
        return

    if data == "adm_statistics":
        conn = get_db()
        c = conn.cursor()
        c.execute("SELECT COUNT(*) as total_users, COALESCE(SUM(balance_bdt), 0) as total_bal FROM users")
        u_stat = c.fetchone()
        c.execute("SELECT COUNT(*) as total_orders, COALESCE(SUM(total_bdt), 0) as total_rev FROM orders WHERE status = 'completed'")
        o_stat = c.fetchone()
        conn.close()

        stat_text = (
            f"📊 <b>শপ সার্বিক পরিসংখ্যান</b>\n"
            f"─────────────────────────────\n"
            f"👥 মোট ইউজার: <b>{u_stat['total_users']}</b> জন\n"
            f"🛍️ মোট সফল অর্ডার: <b>{o_stat['total_orders']}</b> টি\n"
            f"💰 মোট বিক্রয় আয়: <b>{o_stat['total_rev']:.2f} BDT</b>\n"
            f"💳 ইউজার ব্যালেন্স: <b>{u_stat['total_bal']:.2f} BDT</b>"
        )
        await query.message.reply_text(stat_text, parse_mode=ParseMode.HTML)
        return

# ------------------------------------------------------------------------------
# 11. TEXT, PHOTO & DOCUMENT ROUTER (STATE MACHINE)
# ------------------------------------------------------------------------------
async def text_message_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user = update.effective_user
    text = update.message.text.strip()
    state, sdata = get_user_state(user.id)

    if text == "🛍️ Buy Products":
        await buy_products_menu_handler(update, context)
        return
    if text == "👤 My Profile":
        await profile_handler(update, context)
        return
    if text == "💳 Deposit":
        await deposit_menu_handler(update, context)
        return
    if text == "🔐 Get Code":
        await get_code_menu_handler(update, context)
        return
    if text == "💬 Support":
        await support_handler(update, context)
        return
    if text == "👑 Admin Panel":
        if is_admin(user.id):
            await admin_panel_menu_handler(update, context)
        return

    if state == "WAITING_DEP_AMOUNT_MANUAL":
        min_man = float(get_setting("min_deposit_manual", "20.0"))
        try:
            amt = float(text)
            if amt < min_man:
                await update.message.reply_text(f"❌ সর্বনিম্ন ডিপোজিট হলো <b>{min_man:.0f} টাকা</b>। আবার লিখুন:", parse_mode=ParseMode.HTML)
                return
        except ValueError:
            await update.message.reply_text("দয়া করে সঠিক সংখ্যায় টাকার পরিমাণ লিখুন:")
            return

        set_user_state(user.id, "WAITING_DEP_TRX", {"amount": amt})
        kb = InlineKeyboardMarkup([
            [InlineKeyboardButton("📱 bKash", callback_data="usr_dep_method_bKash"), InlineKeyboardButton("📱 Nagad", callback_data="usr_dep_method_Nagad")],
            [InlineKeyboardButton("💎 Binance", callback_data="usr_dep_method_Binance")],
        ])
        await update.message.reply_text(
            f"✅ আপনি <b>{amt:.2f} BDT</b> ডিপোজিট করতে চেয়েছেন।\n\nনিচের কোন মাধ্যমে সেন্ড মানি করবেন সিলেক্ট করুন:",
            parse_mode=ParseMode.HTML,
            reply_markup=kb,
        )
        return

    if state == "WAITING_DEP_TRX":
        trx = text
        method = sdata.get("method", "Manual")
        amount = sdata.get("amount", 20.0)
        now_str = datetime.datetime.now().strftime("%Y-%m-%d %I:%M %p")

        conn = get_db()
        c = conn.cursor()
        c.execute(
            "INSERT INTO deposits (user_id, method, amount, trx_id, status, created_at) VALUES (?, ?, ?, ?, 'pending', ?)",
            (user.id, method, amount, trx, now_str),
        )
        dep_id = c.lastrowid
        conn.commit()
        conn.close()

        set_user_state(user.id, None)
        add_notification(user.id, "💳 Deposit Submitted", f"{method} {amount:.2f} BDT (TrxID: {trx}) is pending approval.")
        await update.message.reply_text(
            f"✅ <b>ডিপোজিট রিকোয়েস্ট সাবমিটেড!</b>\n\nআইডি: #{dep_id} | মেথড: {method}\nপরিমাণ: {amount:.2f} BDT | TrxID: <code>{trx}</code>",
            parse_mode=ParseMode.HTML,
        )

        adm_kb = InlineKeyboardMarkup([[InlineKeyboardButton("✅ Approve", callback_data=f"adm_dep_approve_{dep_id}"), InlineKeyboardButton("❌ Reject", callback_data=f"adm_dep_reject_{dep_id}")]])
        try:
            await context.bot.send_message(
                chat_id=SUPER_ADMIN_ID,
                text=f"🔔 <b>নতুন ডিপোজিট রিকোয়েস্ট #{dep_id}</b>\n👤 {user.full_name} (<code>{user.id}</code>)\n📱 {method} | 💰 {amount:.2f} BDT\n🔢 TrxID: <code>{trx}</code>",
                parse_mode=ParseMode.HTML,
                reply_markup=adm_kb,
            )
        except Exception:
            pass
        return

    if state == "WAITING_2FA_KEY":
        secret = text.replace(" ", "").strip()
        try:
            totp = pyotp.TOTP(secret)
            code_2fa = totp.now()
            time_remaining = 30 - (int(time.time()) % 30)
            set_user_state(user.id, None)
            await update.message.reply_text(
                f"🔑 <b>2FA Live Code</b>\n\n🔢 কোড: <code>{code_2fa}</code>\n⏳ মেয়াদ বাকি: <b>{time_remaining} সেকেন্ড</b>",
                parse_mode=ParseMode.HTML,
            )
        except Exception:
            await update.message.reply_text("❌ অবৈধ 2FA Secret Key! দয়া করে সঠিক কি দিন।")
        return

    if state == "WAITING_CALC_CUSTOM_QTY":
        try:
            qty = int(text)
            if qty <= 0:
                raise ValueError
        except ValueError:
            await update.message.reply_text("দয়া করে সঠিক সংখ্যা লিখুন:")
            return
        prod_id = sdata.get("prod_id")
        set_user_state(user.id, None)
        await render_purchase_summary(update.message, prod_id, quantity=qty, edit=False)
        return

    if not is_admin(user.id):
        return

    if state == "ADMIN_WAITING_STOCK_INPUT":
        lines = [line.strip() for line in text.splitlines() if line.strip()]
        prod_id = sdata.get("prod_id")
        if lines and prod_id:
            conn = get_db()
            c = conn.cursor()
            for line in lines:
                c.execute("INSERT INTO stock (product_id, content) VALUES (?, ?)", (prod_id, line))
            conn.commit()
            c.execute("SELECT COUNT(*) as cnt FROM stock WHERE product_id = ? AND is_sold = 0", (prod_id,))
            total_now = c.fetchone()["cnt"]
            conn.close()
            set_user_state(user.id, None)
            await update.message.reply_text(f"✅ সফলভাবে <b>{len(lines)}</b> টি স্টক যুক্ত হয়েছে! মোট স্টক: <b>{total_now}</b> টি।", parse_mode=ParseMode.HTML)
        return

    if state == "ADMIN_WAITING_VPN_EMAIL":
        order_id = sdata.get("order_id")
        set_user_state(user.id, "ADMIN_WAITING_VPN_PASS", {"order_id": order_id, "email": text.strip()})
        await update.message.reply_text(
            f"✅ Email গ্রহণ করা হয়েছে: <code>{text.strip()}</code>\n\n"
            f"🔐 <b>(ধাপ ২/২):</b> এবার VPN-এর <b>Password</b> লিখে পাঠান:",
            parse_mode=ParseMode.HTML,
        )
        return

    if state == "ADMIN_WAITING_VPN_PASS":
        order_id = sdata.get("order_id")
        v_email = sdata.get("email", "")
        v_pass = text.strip()
        combined_cred = f"{v_email}:{v_pass}"
        conn = get_db()
        c = conn.cursor()
        c.execute("SELECT o.*, p.image_file_id FROM orders o JOIN products p ON o.product_id = p.id WHERE o.id = ?", (order_id,))
        order = c.fetchone()
        if order:
            c.execute("UPDATE orders SET status = 'completed', delivered_data = ? WHERE id = ?", (combined_cred, order_id))
            conn.commit()
            conn.close()
            set_user_state(user.id, None)
            add_notification(order["user_id"], "✅ VPN Order Delivered!", f"Order #{order_id} ({order['product_name']}) has been delivered.")
            await update.message.reply_text(f"✅ VPN Order #{order_id} সফলভাবে কাস্টমারের ইনবক্সে পাঠানো হয়েছে!")
            user_msg = format_credentials_card([combined_cred], order["product_name"], order["category_name"])
            if order["image_file_id"]:
                try:
                    await context.bot.send_photo(chat_id=order["user_id"], photo=order["image_file_id"], caption=user_msg, parse_mode=ParseMode.HTML)
                    return
                except Exception:
                    pass
            await context.bot.send_message(chat_id=order["user_id"], text=user_msg, parse_mode=ParseMode.HTML)
        else:
            conn.close()
            await update.message.reply_text("অর্ডার পাওয়া যায়নি।")
        return

    if state == "ADMIN_WAITING_MANUAL_DELIVERY":
        order_id = sdata.get("order_id")
        conn = get_db()
        c = conn.cursor()
        c.execute("SELECT o.*, p.image_file_id FROM orders o JOIN products p ON o.product_id = p.id WHERE o.id = ?", (order_id,))
        order = c.fetchone()
        if order:
            c.execute("UPDATE orders SET status = 'completed', delivered_data = ? WHERE id = ?", (text, order_id))
            conn.commit()
            conn.close()
            set_user_state(user.id, None)
            add_notification(order["user_id"], "✅ Order Delivered!", f"Order #{order_id} ({order['product_name']}) has been delivered.")
            await update.message.reply_text(f"✅ Order #{order_id} সফলভাবে কাস্টমারকে ডেলিভারি পাঠানো হয়েছে।")
            user_msg = format_credentials_card([text], order['product_name'], order['category_name'])
            if order["image_file_id"]:
                try:
                    await context.bot.send_photo(chat_id=order["user_id"], photo=order["image_file_id"], caption=user_msg, parse_mode=ParseMode.HTML)
                    return
                except Exception:
                    pass
            await context.bot.send_message(chat_id=order["user_id"], text=user_msg, parse_mode=ParseMode.HTML)
        else:
            conn.close()
        return

    if state == "ADMIN_WAITING_WGRP_ID":
        set_setting("welcome_group_id", text.strip())
        set_user_state(user.id, None)
        await update.message.reply_text(f"✅ Update Group ID সেভ হয়েছে: <code>{text.strip()}</code>", parse_mode=ParseMode.HTML)
        return

    if state == "ADMIN_WAITING_WGRP_LINK":
        set_setting("welcome_group_link", text.strip())
        set_user_state(user.id, None)
        await update.message.reply_text(f"✅ Update Group Link সেভ হয়েছে: {text.strip()}")
        return

    if state == "ADMIN_WAITING_TXCH_ID":
        set_setting("transaction_channel_id", text.strip())
        set_user_state(user.id, None)
        await update.message.reply_text(f"✅ Transaction Channel ID সেভ হয়েছে: <code>{text.strip()}</code>", parse_mode=ParseMode.HTML)
        return

    if state == "ADMIN_WAITING_TXCH_LINK":
        set_setting("transaction_channel_link", text.strip())
        set_user_state(user.id, None)
        await update.message.reply_text(f"✅ Transaction Channel Link সেভ হয়েছে: {text.strip()}")
        return

    if state == "ADMIN_WAITING_CUSTOM_TX_NAME":
        set_user_state(user.id, "ADMIN_WAITING_CUSTOM_TX_PROD", {"cname": text})
        await update.message.reply_text("📦 <b>(ধাপ ২/৪):</b> প্রোডাক্টের নাম লিখুন:", parse_mode=ParseMode.HTML)
        return

    if state == "ADMIN_WAITING_CUSTOM_TX_PROD":
        sdata["pname"] = text
        set_user_state(user.id, "ADMIN_WAITING_CUSTOM_TX_QTY", sdata)
        await update.message.reply_text("🔢 <b>(ধাপ ৩/৪):</b> পরিমাণ (Quantity) লিখুন:", parse_mode=ParseMode.HTML)
        return

    if state == "ADMIN_WAITING_CUSTOM_TX_QTY":
        try:
            qty = int(text.strip())
        except ValueError:
            await update.message.reply_text("সঠিক সংখ্যা লিখুন:")
            return
        sdata["qty"] = qty
        set_user_state(user.id, "ADMIN_WAITING_CUSTOM_TX_PRICE", sdata)
        await update.message.reply_text("💰 <b>(ধাপ ৪/৪):</b> মোট মূল্য (BDT) লিখুন:", parse_mode=ParseMode.HTML)
        return

    if state == "ADMIN_WAITING_CUSTOM_TX_PRICE":
        try:
            price = float(text.strip())
        except ValueError:
            await update.message.reply_text("সঠিক সংখ্যা লিখুন:")
            return
        cname = sdata.get("cname", "Customer")
        pname = sdata.get("pname", "VPN")
        qty = sdata.get("qty", 1)
        fake_uid = random.randint(7100000000, 7999999999)
        fake_receipt = f"RCP-{secrets.token_hex(5).upper()}"
        set_user_state(user.id, None)
        ok, err = await post_transaction_to_channel(context.bot, fake_receipt, cname, fake_uid, pname, qty, price)
        if ok:
            await update.message.reply_text(f"✅ কাস্টম ট্রানজ্যাকশন (<code>{fake_receipt}</code>) চ্যানেলে পাঠানো হয়েছে!", parse_mode=ParseMode.HTML)
        else:
            await update.message.reply_text(f"❌ পাঠানো যায়নি: <code>{err}</code>", parse_mode=ParseMode.HTML)
        return

    if state == "ADMIN_WAITING_BCAST_MSG":
        set_user_state(user.id, "ADMIN_WAITING_BCAST_BTN_CHOICE", {"msg_text": text, "photo_id": None})
        kb = InlineKeyboardMarkup([
            [InlineKeyboardButton("✅ হ্যাঁ (বাটন যোগ করব)", callback_data="adm_bcast_add_btn")],
            [InlineKeyboardButton("❌ না (বাটন ছাড়াই পাঠাব)", callback_data="adm_bcast_skip_btn")],
        ])
        await update.message.reply_text("👇 আপনি কি মেসেজের নিচে লিংক-বাটন অ্যাড করতে চান?", reply_markup=kb)
        return

    if state == "ADMIN_WAITING_BCAST_BTN_COUNT":
        try:
            cnt = int(text.strip())
            if cnt < 1 or cnt > 10:
                raise ValueError()
        except Exception:
            await update.message.reply_text("❌ ১ থেকে ১০-এর মধ্যে একটি সংখ্যা লিখুন:")
            return
        sdata["btn_count"] = cnt
        sdata["buttons"] = []
        set_user_state(user.id, "ADMIN_WAITING_BCAST_BTN_NAME", sdata)
        await update.message.reply_text(f"✏️ <b>[বাটন ১/{cnt}]</b> প্রথম বাটনের <b>নাম</b> লিখে পাঠান:", parse_mode=ParseMode.HTML)
        return

    if state == "ADMIN_WAITING_BCAST_BTN_NAME":
        idx = len(sdata.get("buttons", [])) + 1
        cnt = sdata.get("btn_count", 1)
        sdata["cur_btn_name"] = text.strip()
        set_user_state(user.id, "ADMIN_WAITING_BCAST_BTN_URL", sdata)
        await update.message.reply_text(
            f"✅ <b>[{idx}/{cnt}] বাটন গ্রহণ করা হয়েছে!</b>\n"
            f"🔗 এবার এই বাটনের নিচে কোন <b>লিংক (https://...)</b> যাবে তা দিন:",
            parse_mode=ParseMode.HTML,
        )
        return

    if state == "ADMIN_WAITING_BCAST_BTN_URL":
        b_url = text.strip()
        if not (b_url.startswith("http://") or b_url.startswith("https://") or b_url.startswith("tg://")):
            await update.message.reply_text("❌ লিংকটি অবশ্যই <code>https://</code> দিয়ে শুরু হতে হবে। আবার দিন:", parse_mode=ParseMode.HTML)
            return
        btns = sdata.get("buttons", [])
        btns.append({"text": sdata.get("cur_btn_name", "Open"), "url": b_url})
        sdata["buttons"] = btns
        cnt = sdata.get("btn_count", 1)
        if len(btns) < cnt:
            nxt = len(btns) + 1
            set_user_state(user.id, "ADMIN_WAITING_BCAST_BTN_NAME", sdata)
            await update.message.reply_text(
                f"✅ <b>[{len(btns)}/{cnt}]</b> বাটন সম্পন্ন হয়েছে!\n"
                f"✏️ এবার <b>[{nxt}/{cnt}] বাটনের নাম</b> লিখে পাঠান:",
                parse_mode=ParseMode.HTML,
            )
            return
        await execute_smart_broadcast(context.bot, update.message, sdata, buttons_list=btns)
        set_user_state(user.id, None)
        return

    if state == "ADMIN_WAITING_RECEIPT_CODE":
        set_user_state(user.id, None)
        rcode = text.strip().upper()
        if not rcode.startswith("RCP-") and len(rcode) == 10:
            rcode = f"RCP-{rcode}"
        conn = get_db()
        c = conn.cursor()
        c.execute("SELECT o.*, u.name as uname, u.username as u_username FROM orders o JOIN users u ON o.user_id = u.user_id WHERE o.receipt_no LIKE ? LIMIT 1", (f"%{rcode}%",))
        o = c.fetchone()
        conn.close()
        if not o:
            await update.message.reply_text("❌ রিসিট পাওয়া যায়নি।")
            return
        bot_info = await context.bot.get_me()
        pdf_bio = generate_pdf_receipt(
            receipt_no=o["receipt_no"], date_str=o["order_date"], day_name=o["day_of_week"] or "N/A",
            customer_name=o["uname"], telegram_id=o["user_id"], username=o["u_username"] or "",
            item_desc=o["product_name"], category_name=o["category_name"] or "DIGITAL",
            qty=o["quantity"], unit_price=o["unit_price"] or o["total_bdt"], total_bdt=o["total_bdt"],
            bot_username=bot_info.username,
        )
        await update.message.reply_document(document=InputFile(pdf_bio, filename=f"Receipt_{o['receipt_no']}.pdf"), caption=f"🧾 <code>{o['receipt_no']}</code>", parse_mode=ParseMode.HTML)
        return

    if state == "ADMIN_WAITING_REN_SUB":
        sub_id = sdata.get("sub_id")
        conn = get_db()
        c = conn.cursor()
        c.execute("UPDATE subcategories SET name = ? WHERE id = ?", (text, sub_id))
        conn.commit()
        conn.close()
        set_user_state(user.id, None)
        await update.message.reply_text(f"✅ ব্র্যান্ডের নাম পরিবর্তন হয়েছে: <b>{text}</b>", parse_mode=ParseMode.HTML)
        return

    if state == "ADMIN_WAITING_REN_CAT":
        cat_id = sdata.get("cat_id")
        conn = get_db()
        c = conn.cursor()
        c.execute("UPDATE categories SET name = ? WHERE id = ?", (text, cat_id))
        conn.commit()
        conn.close()
        set_user_state(user.id, None)
        await update.message.reply_text(f"✅ ক্যাটাগরির নাম পরিবর্তন হয়েছে: <b>{text}</b>", parse_mode=ParseMode.HTML)
        return

    if state == "ADMIN_WAITING_NEW_CAT_NAME":
        conn = get_db()
        c = conn.cursor()
        c.execute("INSERT INTO categories (name, icon) VALUES (?, '🛍️')", (text,))
        conn.commit()
        conn.close()
        set_user_state(user.id, None)
        await update.message.reply_text(f"✅ নতুন ক্যাটাগরি <b>{text}</b> তৈরি হয়েছে!", parse_mode=ParseMode.HTML)
        return

    if state == "ADMIN_WAITING_NEW_SUB_NAME":
        cat_id = sdata.get("cat_id")
        conn = get_db()
        c = conn.cursor()
        c.execute("INSERT INTO subcategories (category_id, name) VALUES (?, ?)", (cat_id, text))
        conn.commit()
        conn.close()
        set_user_state(user.id, None)
        await update.message.reply_text(f"✅ নতুন ব্র্যান্ড <b>{text}</b> তৈরি হয়েছে!", parse_mode=ParseMode.HTML)
        return

    if state == "ADMIN_WAITING_ADDPROD_NAME":
        sdata["name"] = text
        set_user_state(user.id, "ADMIN_WAITING_ADDPROD_PRICE", sdata)
        await update.message.reply_text(f"এবার <b>{text}</b>-এর BDT দাম লিখুন:", parse_mode=ParseMode.HTML)
        return

    if state == "ADMIN_WAITING_ADDPROD_PRICE":
        try:
            price_bdt = float(text)
        except ValueError:
            await update.message.reply_text("সঠিক সংখ্যা লিখুন:")
            return
        sub_id = sdata.get("sub_id")
        p_name = sdata.get("name")
        rate = float(get_setting("usdt_rate", "125.0"))
        price_usdt = price_bdt / rate if rate > 0 else 0.0
        conn = get_db()
        c = conn.cursor()
        c.execute("INSERT INTO products (subcategory_id, name, price_bdt, price_usdt, is_on_demand) VALUES (?, ?, ?, ?, 0)", (sub_id, p_name, price_bdt, price_usdt))
        conn.commit()
        conn.close()
        set_user_state(user.id, None)
        await update.message.reply_text(f"✅ নতুন ভ্যারিয়েন্ট <b>{p_name}</b> ({price_bdt:.2f} BDT) যুক্ত হয়েছে!", parse_mode=ParseMode.HTML)
        return

    if state == "ADMIN_WAITING_DESC":
        prod_id = sdata.get("prod_id")
        conn = get_db()
        c = conn.cursor()
        c.execute("UPDATE products SET description = ? WHERE id = ?", (text, prod_id))
        conn.commit()
        conn.close()
        set_user_state(user.id, None)
        await update.message.reply_text("✅ কাস্টম বিবরণ সেভ হয়েছে!")
        return

    if state == "ADMIN_WAITING_PRICE":
        try:
            p_bdt = float(text)
            rate = float(get_setting("usdt_rate", "125.0"))
            p_usdt = p_bdt / rate
            prod_id = sdata.get("prod_id")
            conn = get_db()
            c = conn.cursor()
            c.execute("UPDATE products SET price_bdt = ?, price_usdt = ? WHERE id = ?", (p_bdt, p_usdt, prod_id))
            conn.commit()
            conn.close()
            set_user_state(user.id, None)
            await update.message.reply_text(f"✅ নতুন দাম সেট হয়েছে: {p_bdt:.2f} BDT")
        except ValueError:
            await update.message.reply_text("সঠিক সংখ্যা লিখুন:")
        return

    if state == "ADMIN_WAITING_NAME":
        prod_id = sdata.get("prod_id")
        conn = get_db()
        c = conn.cursor()
        c.execute("UPDATE products SET name = ? WHERE id = ?", (text, prod_id))
        conn.commit()
        conn.close()
        set_user_state(user.id, None)
        await update.message.reply_text(f"✅ নাম পরিবর্তন হয়েছে: {text}")
        return

    if state == "ADMIN_WAITING_BKASH":
        set_setting("bkash_number", text)
        set_user_state(user.id, None)
        await update.message.reply_text("✅ বিকাশ নাম্বার সেট হয়েছে।")
        return

    if state == "ADMIN_WAITING_NAGAD":
        set_setting("nagad_number", text)
        set_user_state(user.id, None)
        await update.message.reply_text("✅ নগদ নাম্বার সেট হয়েছে।")
        return

    if state == "ADMIN_WAITING_BINANCE":
        set_setting("binance_id", text)
        set_user_state(user.id, None)
        await update.message.reply_text("✅ বাইনান্স তথ্য সেট হয়েছে।")
        return

    if state == "ADMIN_WAITING_MINDEP_MANUAL":
        try:
            new_min = float(text)
            set_setting("min_deposit_manual", str(new_min))
            set_user_state(user.id, None)
            await update.message.reply_text(f"✅ সর্বনিম্ন ডিপোজিট সেট হয়েছে: {new_min:.0f} BDT")
        except ValueError:
            await update.message.reply_text("সঠিক সংখ্যা লিখুন:")
        return

    if state == "ADMIN_WAITING_RATE":
        try:
            new_rate = float(text)
            set_setting("usdt_rate", str(new_rate))
            set_user_state(user.id, None)
            await update.message.reply_text(f"✅ নতুন রেট: {new_rate}")
        except ValueError:
            await update.message.reply_text("সঠিক সংখ্যা লিখুন:")
        return

    if state == "ADMIN_WAITING_ADDBAL_AMT":
        try:
            amt = float(text)
        except ValueError:
            return
        t_id = sdata.get("target_uid")
        conn = get_db()
        c = conn.cursor()
        c.execute("UPDATE users SET balance_bdt = balance_bdt + ? WHERE user_id = ?", (amt, t_id))
        conn.commit()
        conn.close()
        set_user_state(user.id, None)
        await update.message.reply_text(f"✅ <code>{t_id}</code>-কে {amt:.2f} BDT যোগ করা হয়েছে!", parse_mode=ParseMode.HTML)
        return

    if state == "ADMIN_WAITING_DEDBAL_AMT":
        try:
            amt = float(text)
        except ValueError:
            return
        t_id = sdata.get("target_uid")
        conn = get_db()
        c = conn.cursor()
        c.execute("UPDATE users SET balance_bdt = MAX(0, balance_bdt - ?) WHERE user_id = ?", (amt, t_id))
        conn.commit()
        conn.close()
        set_user_state(user.id, None)
        await update.message.reply_text(f"✅ <code>{t_id}</code> থেকে {amt:.2f} BDT কাটা হয়েছে!", parse_mode=ParseMode.HTML)
        return

async def photo_message_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user = update.effective_user
    if not is_admin(user.id):
        return

    state, sdata = get_user_state(user.id)
    if state == "ADMIN_WAITING_BCAST_MSG":
        photo_file_id = update.message.photo[-1].file_id
        caption = update.message.caption or ""
        set_user_state(user.id, "ADMIN_WAITING_BCAST_BTN_CHOICE", {"msg_text": caption, "photo_id": photo_file_id})
        kb = InlineKeyboardMarkup([
            [InlineKeyboardButton("✅ হ্যাঁ (বাটন যোগ করব)", callback_data="adm_bcast_add_btn")],
            [InlineKeyboardButton("❌ না (বাটন ছাড়াই পাঠাব)", callback_data="adm_bcast_skip_btn")],
        ])
        await update.message.reply_text("👇 আপনি কি ছবির নিচে লিংক-বাটন অ্যাড করতে চান?", reply_markup=kb)
        return

    if state == "ADMIN_WAITING_PHOTO_UPLOAD":
        photo_file_id = update.message.photo[-1].file_id
        prod_id = sdata.get("prod_id")
        conn = get_db()
        c = conn.cursor()
        c.execute("UPDATE products SET image_file_id = ? WHERE id = ?", (photo_file_id, prod_id))
        conn.commit()
        conn.close()
        set_user_state(user.id, None)
        await update.message.reply_text("✅ নতুন ফটো সফলভাবে সেভ হয়েছে!")

async def document_message_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user = update.effective_user
    if not is_admin(user.id):
        return

    state, sdata = get_user_state(user.id)
    if state == "ADMIN_WAITING_STOCK_INPUT":
        doc = update.message.document
        if not doc.file_name.endswith(".xlsx"):
            await update.message.reply_text("❌ দয়া করে <code>.xlsx</code> এক্সেল ফাইল আপলোড করুন।", parse_mode=ParseMode.HTML)
            return

        tg_file = await context.bot.get_file(doc.file_id)
        file_bytes = await tg_file.download_as_bytearray()
        try:
            items = read_items_from_excel(file_bytes)
        except Exception as e:
            await update.message.reply_text(f"❌ এক্সেল ফাইল পড়তে সমস্যা হয়েছে: {e}")
            return

        prod_id = sdata.get("prod_id")
        if not items:
            await update.message.reply_text("❌ এক্সেল ফাইলে কোনো ডেটা পাওয়া যায়নি।")
            return

        conn = get_db()
        c = conn.cursor()
        for item in items:
            c.execute("INSERT INTO stock (product_id, content) VALUES (?, ?)", (prod_id, item))
        conn.commit()
        c.execute("SELECT COUNT(*) as cnt FROM stock WHERE product_id = ? AND is_sold = 0", (prod_id,))
        total_now = c.fetchone()["cnt"]
        conn.close()

        set_user_state(user.id, None)
        await update.message.reply_text(
            f"🎉 <b>এক্সেল স্টক আপলোড সম্পন্ন!</b>\n📊 নতুন যুক্ত: <b>{len(items)}</b> টি | মোট স্টক: <b>{total_now}</b> টি।",
            parse_mode=ParseMode.HTML,
        )

# ------------------------------------------------------------------------------
# 12. APPLICATION INITIALIZATION & RUNNER
# ------------------------------------------------------------------------------
_PERIODIC_TX_TASK = None

async def post_init(application):
    global _GLOBAL_BOT_APP, _GLOBAL_EVENT_LOOP, _PERIODIC_TX_TASK
    _GLOBAL_BOT_APP = application
    _GLOBAL_EVENT_LOOP = asyncio.get_running_loop()
    logger.info("CyberNest Bot post_init completed successfully.")
    _PERIODIC_TX_TASK = asyncio.create_task(send_periodic_channel_transactions(application))

async def post_stop(application):
    global _PERIODIC_TX_TASK
    if _PERIODIC_TX_TASK and not _PERIODIC_TX_TASK.done():
        _PERIODIC_TX_TASK.cancel()
        try:
            await _PERIODIC_TX_TASK
        except asyncio.CancelledError:
            pass

async def send_periodic_channel_transactions(application):
    try:
        await asyncio.sleep(10)
    except asyncio.CancelledError:
        return
    while True:
        if _AUTO_SALES_ACTIVE:
            try:
                conn = get_db()
                c = conn.cursor()
                c.execute("""
                SELECT p.name as prod_name, s.name as sub_name, p.price_bdt 
                FROM products p 
                JOIN subcategories s ON p.subcategory_id = s.id 
                ORDER BY RANDOM() LIMIT 1
                """)
                row = c.fetchone()
                conn.close()

                if row:
                    c_name = random.choice(CUSTOMER_SAMPLE_NAMES)
                    fake_uid = random.randint(7100000000, 7999999999)
                    qty = random.choice([1, 1, 2, 4, 5, 10])
                    p_name = f"{row['sub_name']} ({row['prod_name']})"
                    total_p = row['price_bdt'] * qty
                    fake_receipt = f"RCP-{secrets.token_hex(5).upper()}"
                    await post_transaction_to_channel(application.bot, fake_receipt, c_name, fake_uid, p_name, qty, total_p)
            except Exception as e:
                logger.error(f"Auto periodic sales loop error: {e}")

        try:
            await asyncio.sleep(random.randint(300, 420))
        except asyncio.CancelledError:
            break

async def send_auto_delete_welcome_group(bot, chat_id, user):
    user_id = user.id
    if user_id in _RECENTLY_WELCOMED:
        return
    _RECENTLY_WELCOMED.add(user_id)

    user_link = f'<a href="tg://user?id={user.id}">{user.full_name or "মেম্বার"}</a>'
    welcome_text = (
        f"🌸 <b>আসসালামু আলাইকুম</b>, {user_link}! ✨\n\n"
        f"আমাদের কমিউনিটি গ্রুপে আপনাকে স্বাগতম। 🎉"
    )
    try:
        sent_msg = await bot.send_message(chat_id=chat_id, text=welcome_text, parse_mode=ParseMode.HTML)
        await asyncio.sleep(30)
        await bot.delete_message(chat_id=chat_id, message_id=sent_msg.message_id)
    except Exception:
        pass
    finally:
        await asyncio.sleep(60)
        _RECENTLY_WELCOMED.discard(user_id)

async def handle_new_group_members(update: Update, context: ContextTypes.DEFAULT_TYPE):
    chat_id = str(update.effective_chat.id)
    if chat_id != get_welcome_group_id():
        return
    if update.message and update.message.new_chat_members:
        for member in update.message.new_chat_members:
            if not member.is_bot:
                asyncio.create_task(send_auto_delete_welcome_group(context.bot, chat_id, member))

async def handle_group_member_updated(update: Update, context: ContextTypes.DEFAULT_TYPE):
    chat_id = str(update.effective_chat.id)
    if chat_id != get_welcome_group_id():
        return
    result = update.chat_member
    if result:
        new_m = result.new_chat_member
        if new_m.status in ["member", "administrator"] and result.old_chat_member.status in ["left", "kicked"]:
            if not new_m.user.is_bot:
                asyncio.create_task(send_auto_delete_welcome_group(context.bot, chat_id, new_m.user))

async def id_command_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    chat = update.effective_chat
    msg = update.effective_message
    if not chat or not msg:
        return
    title = chat.title or chat.full_name or "Chat"
    await msg.reply_text(
        f"🆔 <b>Chat Info:</b>\n📌 <b>Name:</b> {title}\n🔹 <b>Chat ID:</b> <code>{chat.id}</code>\n📂 <b>Type:</b> {chat.type}",
        parse_mode=ParseMode.HTML,
    )

def main():
    print("=" * 60)
    print("Starting Render Web Service + Embedded Mini App Server...")
    threading.Thread(target=start_health_server, daemon=True).start()
    print("Initializing Database...")
    init_db()
    print("Database Initialized Successfully.")
    print(f"Configured Super Admin ID: {SUPER_ADMIN_ID}")
    print("=" * 60)

    global _GLOBAL_BOT_APP, _GLOBAL_EVENT_LOOP
    loop = asyncio.new_event_loop()
    asyncio.set_event_loop(loop)
    _GLOBAL_EVENT_LOOP = loop

    async def global_error_handler(update: object, context: ContextTypes.DEFAULT_TYPE) -> None:
        logger.error(f"Update exception caught: {context.error}", exc_info=context.error)

    app = ApplicationBuilder().token(BOT_TOKEN).post_init(post_init).post_stop(post_stop).build()
    app.add_error_handler(global_error_handler)
    _GLOBAL_BOT_APP = app

    app.add_handler(MessageHandler(filters.StatusUpdate.NEW_CHAT_MEMBERS, handle_new_group_members))
    app.add_handler(ChatMemberHandler(handle_group_member_updated, ChatMemberHandler.CHAT_MEMBER))

    app.add_handler(CommandHandler("start", start_handler))
    app.add_handler(CommandHandler("id", id_command_handler))
    app.add_handler(MessageHandler(filters.ChatType.CHANNEL & filters.Regex(r"^/id"), id_command_handler))
    app.add_handler(CallbackQueryHandler(callback_query_handler))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, text_message_handler))
    app.add_handler(MessageHandler(filters.PHOTO, photo_message_handler))
    app.add_handler(MessageHandler(filters.Document.ALL, document_message_handler))

    print("Bot is polling for updates...")
    app.run_polling(allowed_updates=Update.ALL_TYPES)

if __name__ == "__main__":
    main()
