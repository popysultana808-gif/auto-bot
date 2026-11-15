# ==============================================================================
# CyberNest Digital Mart — Full-Featured Telegram Auto-Shop Bot + Mini App API
# Two-File Architecture (Part 1: bot.py)
# Companion File: webapp.html (Served dynamically at /app)
# Tech Stack: Python (python-telegram-bot v21+), SQLite / Neon PostgreSQL,
# OpenPyXL (.xlsx Stock & Delivery), ReportLab (QR + Logo PDF Receipts),
# PyOTP (2FA Live Code), IMAP4_SSL (Outlook/Hotmail OTP), Smart AI Assistant,
# and Complete REST API Backend for webapp.html
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
    WebAppInfo,
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

WEBAPP_FILE_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "webapp.html")

_AUTO_SALES_ACTIVE = True
_RECENTLY_WELCOMED = set()
_USER_PHOTO_CACHE: Dict[int, Tuple[bytes, str, float]] = {}

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
    req = urllib.request.Request(url, data=encoded, headers={"User-Agent": "CyberNestBot/2.0"})
    with urllib.request.urlopen(req, timeout=timeout) as resp:
        return json.loads(resp.read().decode("utf-8"))

def _http_get_bytes(url: str, timeout: int = 10):
    req = urllib.request.Request(url, headers={"User-Agent": "CyberNestBot/2.0"})
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

def get_vip_badge_info(total_spent: float) -> Dict[str, Any]:
    """Calculates VIP Membership Badge based on user's lifetime spending."""
    spent = float(total_spent or 0.0)
    if spent >= 2000:
        return {"tier": "Gold VIP", "icon": "🥇", "badge": "🥇 Gold VIP", "color": "#f59e0b", "next_tier": "MAX", "progress": 100}
    elif spent >= 500:
        prog = min(100, int(((spent - 500) / 1500) * 100))
        return {"tier": "Silver VIP", "icon": "🥈", "badge": "🥈 Silver VIP", "color": "#94a3b8", "next_tier": "Gold VIP (৳2000)", "progress": prog}
    elif spent >= 100:
        prog = min(100, int(((spent - 100) / 400) * 100))
        return {"tier": "Bronze VIP", "icon": "🥉", "badge": "🥉 Bronze VIP", "color": "#d97706", "next_tier": "Silver VIP (৳500)", "progress": prog}
    else:
        prog = min(100, int((spent / 100) * 100))
        return {"tier": "Member", "icon": "⭐", "badge": "⭐ Member", "color": "#6366f1", "next_tier": "Bronze VIP (৳100)", "progress": prog}

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
# 2. SMART LIVE DATABASE-BACKED AI SUPPORT ENGINE (Bangla / Banglish / English)
# ------------------------------------------------------------------------------
def _build_live_store_context(user_id: Optional[int] = None) -> Dict[str, Any]:
    conn = get_db()
    c = conn.cursor()
    c.execute("""
        SELECT p.id, p.name, p.price_bdt, COALESCE(p.is_active, 1) as is_active,
               s.name as sub_name, cat.name as cat_name,
               (SELECT COUNT(*) FROM stock st WHERE st.product_id = p.id AND st.is_sold = 0) as in_stock
        FROM products p
        JOIN subcategories s ON p.subcategory_id = s.id
        JOIN categories cat ON s.category_id = cat.id
        ORDER BY cat.position ASC, s.position ASC, p.id ASC
    """)
    prods = [dict(r) for r in c.fetchall()]
    user_bal = 0.0
    user_name = "Customer"
    if user_id:
        c.execute("SELECT balance_bdt, name FROM users WHERE user_id = ?", (user_id,))
        urow = c.fetchone()
        if urow:
            user_bal = float(urow["balance_bdt"] or 0.0)
            user_name = urow["name"] or "Customer"
    conn.close()
    return {
        "prods": prods,
        "bkash": get_setting("bkash_number", "N/A"),
        "nagad": get_setting("nagad_number", "N/A"),
        "binance": get_setting("binance_id", "N/A"),
        "min_dep": get_setting("min_deposit_manual", "20"),
        "usdt_rate": get_setting("usdt_rate", "125"),
        "support": get_setting("support_username", "@Talha_juba098"),
        "shop_name": get_setting("shop_name", "CyberNest Market"),
        "user_bal": user_bal,
        "user_name": user_name,
    }

_AI_CHAT_HISTORY: Dict[int, List[Dict[str, str]]] = {}

def _call_gemini_if_configured(question: str, ctx: Dict[str, Any], user_id: Optional[int] = None) -> Optional[str]:
    db_ai_key = get_setting("ai_api_key", "").strip()
    raw_gemini = os.getenv("GEMINI_API_KEY", "").strip()
    groq_key = os.getenv("GROQ_API_KEY", "").strip()

    if db_ai_key:
        if db_ai_key.startswith("gsk_"):
            groq_key = db_ai_key
        elif db_ai_key.startswith("AIza"):
            raw_gemini = db_ai_key

    # Smart auto-detection: if user put gsk_ key inside GEMINI_API_KEY, use it for Groq
    if not groq_key and raw_gemini.startswith("gsk_"):
        groq_key = raw_gemini
        raw_gemini = ""

    catalog_lines = []
    for p in ctx["prods"]:
        if not p["is_active"]:
            continue
        cat_low = p["cat_name"].lower()
        st_str = f"Stock: {p['in_stock']}" if "mail" in cat_low else ("Ready to order" if "proxy" in cat_low else "Verified/Warranty")
        catalog_lines.append(f"- {p['cat_name']} | {p['sub_name']} | {p['name']} = {format_price(p['price_bdt'])} BDT ({st_str})")

    sys_prompt = (
        f"You are the highly intelligent, helpful, and friendly AI Assistant of '{ctx['shop_name']}' — a 24/7 Automated Telegram Digital Store & Mini App.\n"
        f"Think carefully and give thoughtful, accurate, well-structured answers to ANY question the user asks:\n"
        f"1. Store & Business Questions: Explain our A-to-Z features, live product prices, stock, how to buy, deposit methods, 2FA/Outlook OTP tools, referral system, or VIP badges using the LIVE DATA below.\n"
        f"2. General & Outside Knowledge: Answer casual greetings, general knowledge, tech, study, or any everyday question naturally like an intelligent human collaborator.\n\n"
        f"LANGUAGE RULE:\n"
        f"- Reply in the exact language/style the user uses. If they write in Bangla or Banglish (Bengali written with English letters), reply in clear, natural Bangla. If they write in English, reply in English.\n"
        f"- Use Telegram HTML formatting ONLY (<b>, <i>, <code>) when formatting lists or highlights. Never use Markdown ** or ###.\n\n"
        f"COMPLETE A-TO-Z SHOP FEATURES:\n"
        f"- Automated Bot & Premium Mini App (Open Shop): RGB animated search, Filter & Sort, Auto-sliding Promo Banners, Top Selling slider, Brand-grouped Plan Selector, Light/Dark Theme, EN/BN Language Switch, BDT/USDT Currency Switch.\n"
        f"- Products: Premium VPNs (with Verified/Warranty), Residential Proxies (Ready to order), Fresh & Recovery-Added Mails (live stock count + instant .xlsx Excel backup).\n"
        f"- Instant Delivery & Official Receipts: One-Tap Copy credential cards, Excel (.xlsx) bulk files, and Official PDF Receipts with QR code verification.\n"
        f"- Free Tools: Live 2FA Authenticator (6-digit TOTP) & Outlook/Hotmail IMAP Live OTP Reader.\n"
        f"- Wallet & Deposit: Calculator-style deposit with bKash, Nagad & Binance (USDT).\n"
        f"- VIP Membership Badges: Automatic Bronze (৳100+), Silver (৳500+), Gold VIP (৳2000+) badges + Referral & Earnings tracking.\n\n"
        f"LIVE DATABASE SNAPSHOT:\n"
        f"- Customer: {ctx['user_name']} | Wallet Balance: {ctx['user_bal']:.2f} BDT\n"
        f"- Payment Numbers: bKash={ctx['bkash']} | Nagad={ctx['nagad']} | Binance={ctx['binance']}\n"
        f"- Min Deposit: {ctx['min_dep']} BDT | USDT Rate: 1 USDT = {ctx['usdt_rate']} BDT | Support: {ctx['support']}\n"
        f"- Live Catalog:\n" + "\n".join(catalog_lines[:75])
    )

    uid_key = int(user_id or 0)
    history = _AI_CHAT_HISTORY.get(uid_key, [])[-6:]

    # 1. GROQCLOUD FIRST (Ultra-fast ~0.5-1s response)
    if groq_key:
        messages = [{"role": "system", "content": sys_prompt}]
        for h in history:
            messages.append(h)
        messages.append({"role": "user", "content": question})

        for model_name in ("llama-3.3-70b-versatile", "llama-3.1-8b-instant"):
            try:
                g_payload = json.dumps({
                    "model": model_name,
                    "messages": messages,
                    "temperature": 0.65,
                    "max_tokens": 700,
                }).encode("utf-8")
                g_req = urllib.request.Request(
                    "https://api.groq.com/openai/v1/chat/completions",
                    data=g_payload,
                    headers={
                        "Content-Type": "application/json",
                        "Authorization": f"Bearer {groq_key}",
                        "User-Agent": "CyberNestBot/2.0",
                    },
                )
                with urllib.request.urlopen(g_req, timeout=10) as resp:
                    g_data = json.loads(resp.read().decode("utf-8"))
                ans = g_data["choices"][0]["message"]["content"].strip()
                if ans:
                    history.append({"role": "user", "content": question})
                    history.append({"role": "assistant", "content": ans})
                    _AI_CHAT_HISTORY[uid_key] = history[-6:]
                    return ans
            except Exception as e:
                logger.warning(f"Groq model {model_name} failed: {e}")

    # 2. GEMINI API (Only if valid AIza... key is set; skips slow invalid loops)
    if raw_gemini.startswith("AIza"):
        contents = []
        for h in history:
            role = "model" if h["role"] == "assistant" else "user"
            contents.append({"role": role, "parts": [{"text": h["content"]}]})
        contents.append({"role": "user", "parts": [{"text": question}]})

        payload = json.dumps({
            "system_instruction": {"parts": [{"text": sys_prompt}]},
            "contents": contents,
            "generationConfig": {"temperature": 0.65, "maxOutputTokens": 700},
        }).encode("utf-8")

        for g_model in ("gemini-2.5-flash", "gemini-2.0-flash"):
            try:
                url = f"https://generativelanguage.googleapis.com/v1beta/models/{g_model}:generateContent?key={raw_gemini}"
                req = urllib.request.Request(url, data=payload, headers={"Content-Type": "application/json", "User-Agent": "CyberNestBot/2.0"})
                with urllib.request.urlopen(req, timeout=10) as resp:
                    data = json.loads(resp.read().decode("utf-8"))
                cands = data.get("candidates", [])
                if cands:
                    parts = cands[0].get("content", {}).get("parts", [])
                    if parts and parts[0].get("text"):
                        ans = parts[0]["text"].strip()
                        history.append({"role": "user", "content": question})
                        history.append({"role": "assistant", "content": ans})
                        _AI_CHAT_HISTORY[uid_key] = history[-6:]
                        return ans
            except Exception as e:
                logger.warning(f"Gemini {g_model} failed: {e}")

    return None

def generate_smart_ai_reply(question: str, user_id: Optional[int] = None) -> str:
    """Conversational AI Assistant: understands Bangla, Banglish & English and answers the exact question."""
    raw_q = (question or "").strip()
    q = raw_q.lower()
    ctx = _build_live_store_context(user_id=user_id)

    # Try live Gemini API first if GEMINI_API_KEY is configured in Render
    gemini_ans = _call_gemini_if_configured(raw_q, ctx, user_id=user_id)
    if gemini_ans:
        return gemini_ans

    prods = ctx["prods"]
    has_bn_chars = any("\u0980" <= ch <= "\u09ff" for ch in raw_q)
    banglish_words = [
        "kivabe", "koto", "dam", "kinbo", "kinte", "chai", "lagbe", "ache", "ase",
        "amar", "taka", "bhai", "vai", "assalamu", "salam", "walaikum", "kemon",
        "achen", "ashen", "valo", "bhalo", "korbo", "korte", "parbo", "holo", "problem",
        "somossa", "oshubidha", "bkash", "nagad", "kothay", "debo", "dibo", "nibo", "please"
    ]
    is_bangla_ctx = has_bn_chars or any(w in q for w in banglish_words)

    # 1. Greetings (Salam, Hello, Hi, How are you)
    if any(g in q for g in ["assalamu", "salam", "সালাম", "আসসালামু"]):
        return (
            f"ওয়ালাইকুম আসসালাম, {html.escape(ctx['user_name'])}! 😊\n"
            f"কেমন আছেন? আপনার কোনো প্রশ্ন বা অসুবিধা থাকলে আমাকে নির্দ্বিধায় জিজ্ঞেস করুন—আমি সাথে সাথে সাহায্য করব।"
        )
    if any(g in q for g in ["হ্যালো", "হাই", "কেমন আছ", "kemon ach", "kemon as", "ki khobor", "কি খবর", "hello", "hlw", "hey", "how are you"]) or q == "hi":
        if is_bangla_ctx or has_bn_chars:
            return f"হ্যালো {html.escape(ctx['user_name'])}! 😊 আলহামদুলিল্লাহ ভালো আছি। আপনি কেমন আছেন? বলুন, আজ আপনাকে কীভাবে সাহায্য করতে পারি?"
        return f"Hello {html.escape(ctx['user_name'])}! 😊 I'm doing great! How can I help you today?"

    # 2. How to Buy / Purchase Guide (Specific step-by-step instructions)
    is_how_to_buy = any(k in q for k in ["kivabe", "how to buy", "how to purchase", "purchase", "কিনব", "কিভাবে", "অর্ডার করব", "order korbo", "kinbo", "kinte"])
    if is_how_to_buy:
        if any(k in q for k in ["vpn", "ভিপিএন", "nord", "express", "surfshark", "proton"]):
            if is_bangla_ctx:
                return (
                    "🛡️ <b>VPN কেনার সম্পূর্ণ নিয়ম (Step-by-Step):</b>\n"
                    "────────────────────────\n"
                    "১. প্রথমে ওয়ালেটে ব্যালেন্স না থাকলে <b>💳 Deposit</b> থেকে টাকা অ্যাড করে নিন।\n"
                    "২. এরপর বটের <b>🛍️ Buy Products</b> অথবা <b>Open Shop</b> (মিনি অ্যাপ)-এ যান।\n"
                    "৩. <b>🛡️ VPN</b> ক্যাটাগরি থেকে আপনার পছন্দের ব্র্যান্ড (যেমন: Nord VPN, Express VPN) ও মেয়াদ (৩/৭/১৫/৩০ দিন) সিলেক্ট করুন।\n"
                    "৪. <b>Confirm</b> বাটনে চাপলেই আপনার অর্ডার সম্পন্ন হবে এবং ১৫–৩০ মিনিটের মধ্যে আপনার ইনবক্সে Email, Password ও অফিশিয়াল PDF রিসিট পৌঁছে যাবে।"
                )
            return (
                "🛡️ <b>How to Purchase a VPN (Step-by-Step):</b>\n"
                "────────────────────────\n"
                "1. Make sure you have sufficient wallet balance (use <b>💳 Deposit</b> if needed).\n"
                "2. Open <b>🛍️ Buy Products</b> or the <b>Open Shop</b> Mini App.\n"
                "3. Select <b>🛡️ VPN</b>, pick your preferred brand & duration, and tap <b>Confirm</b>.\n"
                "4. You will receive your verified Email & Password plus an official PDF receipt in your inbox within 15–30 minutes."
            )
        if any(k in q for k in ["proxy", "প্রক্সি", "rapid", "9 proxy", "711"]):
            if is_bangla_ctx:
                return (
                    "🌐 <b>Proxy কেনার সম্পূর্ণ নিয়ম:</b>\n"
                    "────────────────────────\n"
                    "১. <b>💳 Deposit</b> থেকে ওয়ালেটে পর্যাপ্ত ব্যালেন্স যুক্ত করুন।\n"
                    "২. <b>Open Shop</b> বা <b>🛍️ Buy Products</b> থেকে <b>🌐 Proxy</b> ক্যাটাগরিতে যান।\n"
                    "৩. আপনার পছন্দের ব্র্যান্ড ও ডাটা প্যাক (100 MB, 200 MB, 500 MB, 1 GB) সিলেক্ট করে <b>Confirm</b> করুন।\n"
                    "৪. অর্ডার কনফার্ম হলেই আপনার ইনবক্সে One-Tap Copy ফরম্যাটে প্রক্সি ও PDF রিসিট ডেলিভারি পাবেন।"
                )
            return (
                "🌐 <b>How to Purchase a Proxy:</b>\n"
                "────────────────────────\n"
                "1. Top up your wallet via <b>💳 Deposit</b>.\n"
                "2. Go to <b>🌐 Proxy</b> in the bot menu or Mini App.\n"
                "3. Choose your brand and data plan (100 MB – 1 GB) and tap <b>Confirm</b>.\n"
                "4. Your proxy credentials (One-Tap Copy format) and PDF receipt will be delivered to your inbox."
            )
        if is_bangla_ctx:
            return (
                "🛍️ <b>কেনাকাটা করার সম্পূর্ণ গাইডলাইন:</b>\n"
                "────────────────────────\n"
                "১. প্রথমে <b>💳 Deposit</b> বাটনে ক্লিক করে bKash, Nagad বা Binance-এর মাধ্যমে ব্যালেন্স অ্যাড করুন।\n"
                "২. এরপর <b>🛍️ Buy Products</b> অথবা বাম পাশের <b>Open Shop</b> মিনি অ্যাপ ওপেন করুন।\n"
                "৩. পছন্দের প্রোডাক্ট ও প্ল্যান সিলেক্ট করে <b>Confirm</b> চাপুন।\n"
                "৪. Mail সাথে সাথেই অটো-ডেলিভারি হয় (এক্সেল ফাইল ও PDF রিসিটসহ), আর VPN/Proxy ১৫–৩০ মিনিটের মধ্যে ইনবক্সে ডেলিভারি দেওয়া হয়।"
            )
        return (
            "🛍️ <b>How to Purchase (Complete Guide):</b>\n"
            "────────────────────────\n"
            "1. Add balance to your wallet using <b>💳 Deposit</b> (bKash, Nagad, or Binance).\n"
            "2. Open <b>🛍️ Buy Products</b> or the <b>Open Shop</b> Mini App.\n"
            "3. Choose your product, select quantity/plan, and tap <b>Confirm</b>.\n"
            "4. Mail accounts deliver instantly with Excel & PDF receipts; VPN & Proxy orders deliver within 15–30 mins."
        )

    # 3. Deposit / Payment queries
    if any(k in q for k in ["deposit", "payment", "bkash", "nagad", "binance", "ডিপোজিট", "বিকাশ", "নগদ", "বাইনান্স", "পেমেন্ট", "টাকা", "রিচার্জ", "recharge"]):
        if is_bangla_ctx:
            return (
                f"💳 <b>ডিপোজিট করার নিয়ম ও লাইভ নাম্বার:</b>\n"
                f"────────────────────────\n"
                f"📱 <b>bKash:</b> <code>{ctx['bkash']}</code>\n"
                f"📱 <b>Nagad:</b> <code>{ctx['nagad']}</code>\n"
                f"💎 <b>Binance:</b> <code>{ctx['binance']}</code>\n\n"
                f"📌 সর্বনিম্ন ডিপোজিট: <b>৳{format_price(float(ctx['min_dep']))} BDT</b> (1 USDT = ৳{ctx['usdt_rate']})\n"
                f"💰 আপনার বর্তমান ব্যালেন্স: <b>৳{ctx['user_bal']:.2f} BDT</b>\n\n"
                f"👉 উপরের নাম্বারে Send Money করে <b>💳 Deposit</b> মেনুতে গিয়ে টাকার পরিমাণ ও TrxID সাবমিট করলেই ব্যালেন্স যুক্ত হয়ে যাবে।"
            )
        return (
            f"💳 <b>Deposit Instructions & Live Numbers:</b>\n"
            f"────────────────────────\n"
            f"📱 <b>bKash:</b> <code>{ctx['bkash']}</code>\n"
            f"📱 <b>Nagad:</b> <code>{ctx['nagad']}</code>\n"
            f"💎 <b>Binance:</b> <code>{ctx['binance']}</code>\n\n"
            f"📌 Min Deposit: <b>৳{format_price(float(ctx['min_dep']))} BDT</b> (1 USDT = ৳{ctx['usdt_rate']})\n"
            f"💰 Your Balance: <b>৳{ctx['user_bal']:.2f} BDT</b>\n\n"
            f"Send Money to any number above, then submit the amount & TrxID in the <b>💳 Deposit</b> tab."
        )

    # 4. Check if user asked about a SPECIFIC Brand (e.g. "Rapid Proxy", "Nord VPN", "Hotmail")
    brand_matches = []
    seen_subs = set()
    for p in prods:
        if not p["is_active"]:
            continue
        sub_low = p["sub_name"].lower()
        tokens = [t for t in sub_low.split() if len(t) > 2]
        if sub_low in q or any(t in q for t in tokens):
            if p["sub_id"] if "sub_id" in p else p["sub_name"] not in seen_subs:
                brand_matches.append(p)
    if brand_matches:
        grouped_b: Dict[str, List[dict]] = {}
        for p in brand_matches:
            grouped_b.setdefault(p["sub_name"], []).append(p)
        lines = []
        for bname, plist in grouped_b.items():
            lines.append(f"🏷️ <b>{bname}</b> ({plist[0]['cat_name']}):")
            for pl in plist:
                cat_l = pl["cat_name"].lower()
                extra = f" — Stock: {pl['in_stock']}" if "mail" in cat_l else ""
                lines.append(f"   • {pl['name']}: <b>৳{format_price(pl['price_bdt'])} BDT</b>{extra}")
        return "\n".join(lines)

    # 5. Category-specific Price / Stock Queries
    if any(k in q for k in ["vpn", "ভিপিএন"]):
        vpn_items = [p for p in prods if "vpn" in p["cat_name"].lower() and p["is_active"]]
        grouped: Dict[str, List[str]] = {}
        for p in vpn_items:
            grouped.setdefault(p["sub_name"], []).append(f"{p['name']}: ৳{format_price(p['price_bdt'])}")
        lines = ["🛡️ <b>লাইভ VPN প্রাইস লিস্ট:</b>" if is_bangla_ctx else "🛡️ <b>Live VPN Price List:</b>", "────────────────────────"]
        for brand, plans in grouped.items():
            lines.append(f"🔹 <b>{brand}:</b> " + " | ".join(plans))
        return "\n".join(lines)

    if any(k in q for k in ["proxy", "প্রক্সি"]):
        proxy_items = [p for p in prods if "proxy" in p["cat_name"].lower() and p["is_active"]]
        grouped_p: Dict[str, List[str]] = {}
        for p in proxy_items:
            grouped_p.setdefault(p["sub_name"], []).append(f"{p['name']}: ৳{format_price(p['price_bdt'])}")
        lines = ["🌐 <b>লাইভ Proxy প্যাকেজ ও দাম:</b>" if is_bangla_ctx else "🌐 <b>Live Proxy Packages & Prices:</b>", "────────────────────────"]
        for brand, plans in grouped_p.items():
            lines.append(f"🔸 <b>{brand}:</b> " + " | ".join(plans))
        return "\n".join(lines)

    if any(k in q for k in ["mail", "hotmail", "outlook", "recovery", "মেইল", "হটমেইল", "আউটলুক", "স্টক", "stock"]):
        mail_items = [p for p in prods if "mail" in p["cat_name"].lower() and p["is_active"]]
        lines = ["✉️ <b>লাইভ Mail দাম ও স্টক:</b>" if is_bangla_ctx else "✉️ <b>Live Mail Prices & Stock:</b>", "────────────────────────"]
        for p in mail_items:
            st_txt = f"🟢 Stock: {p['in_stock']} pcs" if p["in_stock"] > 0 else "🔴 Out of Stock"
            lines.append(f"📧 <b>{p['sub_name']} ({p['name']}):</b> ৳{format_price(p['price_bdt'])} — {st_txt}")
        return "\n".join(lines)

    # 6. Balance, 2FA / OTP, Referral, Support
    if any(k in q for k in ["balance", "ব্যালেন্স", "tk", "taka"]):
        return (
            f"💰 আপনার বর্তমান ওয়ালেট ব্যালেন্স: <b>৳{ctx['user_bal']:.2f} BDT</b>।"
            if is_bangla_ctx
            else f"💰 Your current wallet balance is <b>৳{ctx['user_bal']:.2f} BDT</b>."
        )

    if any(k in q for k in ["2fa", "otp", "code", "কোড", "ওটিপি"]):
        return (
            "🔐 ফ্রি 2FA ও Mail OTP পেতে বটের <b>🔐 Get Code</b> মেনু অথবা ওয়েব অ্যাপের <b>Tools</b> ট্যাবে যান। সেখানে Secret Key বা Outlook টোকেন দিলেই সাথে সাথে কোড পাবেন।"
            if is_bangla_ctx
            else "🔐 For free 2FA or Mail OTP codes, open <b>🔐 Get Code</b> in the bot or the <b>Tools</b> tab in the Mini App."
        )

    if any(k in q for k in ["refer", "referral", "রেফার", "রেফারেল"]):
        return (
            "👥 প্রোফাইলের <b>Referral & Earnings</b> পেজ থেকে আপনার রেফারেল লিংক কপি করে বন্ধুদের শেয়ার করুন—তারা জয়েন করলেই আপনার লিস্টে লাইভ দেখতে পাবেন!"
            if is_bangla_ctx
            else "👥 Copy your invite link from the <b>Referral & Earnings</b> page in your Profile to track every friend who joins."
        )

    # 7. Polite conversational fallback
    if is_bangla_ctx:
        return (
            f"আমি আপনার প্রশ্নটি পুরোপুরি বুঝতে পারিনি। আপনি নির্দিষ্ট করে জিজ্ঞেস করতে পারেন—যেমন:\n"
            f"• <i>&quot;VPN কীভাবে কিনব?&quot;</i>\n"
            f"• <i>&quot;Rapid Proxy-এর দাম কত?&quot;</i>\n"
            f"• <i>&quot;ডিপোজিট করব কীভাবে?&quot;</i>\n\n"
            f"অথবা সরাসরি এডমিনের সাথে কথা বলুন: {ctx['support']}"
        )
    return (
        f"I didn't quite catch that. Could you ask more specifically—for example, about a specific VPN/Proxy price, how to purchase, or how to deposit?\n"
        f"For direct help, you can also message our admin: {ctx['support']}"
    )

# ------------------------------------------------------------------------------
# 3. RENDER WEB SERVICE & MINI APP API SERVER (Serves external webapp.html)
# ------------------------------------------------------------------------------
def load_webapp_html_bytes() -> bytes:
    """Loads webapp.html from disk so edits take effect immediately without embedding."""
    candidates = [
        WEBAPP_FILE_PATH,
        os.path.join(os.getcwd(), "webapp.html"),
    ]
    for path in candidates:
        if os.path.exists(path):
            try:
                with open(path, "rb") as f:
                    return f.read()
            except Exception as e:
                logger.error(f"Failed reading webapp.html from {path}: {e}")
    fallback = """<!DOCTYPE html><html><head><meta charset="utf-8"><title>CyberNest Market</title></head>
    <body style="background:#0b0f19;color:#fff;font-family:sans-serif;text-align:center;padding:40px">
    <h2>⚠️ webapp.html not found!</h2>
    <p>Please keep <code>webapp.html</code> in the same folder as <code>bot.py</code> on GitHub/Render.</p>
    </body></html>"""
    return fallback.encode("utf-8")

class HealthCheckHandler(BaseHTTPRequestHandler):
    def _send_json(self, obj, code_num=200):
        body = json.dumps(obj, ensure_ascii=False, default=str).encode("utf-8")
        self.send_response(code_num)
        self.send_header("Content-type", "application/json; charset=utf-8")
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Access-Control-Allow-Headers", "Content-Type, X-Telegram-Init-Data")
        self.send_header("Access-Control-Allow-Methods", "GET, POST, OPTIONS")
        self.end_headers()
        self.wfile.write(body)

    def do_OPTIONS(self):
        self.send_response(200)
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Access-Control-Allow-Headers", "Content-Type, X-Telegram-Init-Data")
        self.send_header("Access-Control-Allow-Methods", "GET, POST, OPTIONS")
        self.end_headers()

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

        # Serve webapp.html on /app or /webapp.html
        if path in ("/app", "/app/", "/webapp.html"):
            content = load_webapp_html_bytes()
            self.send_response(200)
            self.send_header("Content-type", "text/html; charset=utf-8")
            self.send_header("Cache-Control", "no-cache")
            self.end_headers()
            self.wfile.write(content)
            return

        # Serve Product / Banner Photo from Telegram file_id
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

        # Serve Bot's Official Profile Picture / Logo for Splash Screen (/api/bot_photo)
        if path == "/api/bot_photo":
            now_ts = time.time()
            cached_bot = _USER_PHOTO_CACHE.get(0)
            if cached_bot and (now_ts - cached_bot[2] < 3600):
                self.send_response(200)
                self.send_header("Content-type", cached_bot[1])
                self.send_header("Cache-Control", "public, max-age=3600")
                self.end_headers()
                self.wfile.write(cached_bot[0])
                return
            try:
                bot_id = None
                if _GLOBAL_BOT_APP and _GLOBAL_BOT_APP.bot:
                    bot_id = _GLOBAL_BOT_APP.bot.id
                if not bot_id:
                    me_raw, _ = _http_get_bytes(f"https://api.telegram.org/bot{BOT_TOKEN}/getMe", timeout=6)
                    bot_id = json.loads(me_raw.decode("utf-8"))["result"]["id"]
                raw_photos, _ = _http_get_bytes(
                    f"https://api.telegram.org/bot{BOT_TOKEN}/getUserProfilePhotos?user_id={bot_id}&limit=1",
                    timeout=8,
                )
                pdata = json.loads(raw_photos.decode("utf-8"))
                photos_list = pdata.get("result", {}).get("photos", [])
                if photos_list:
                    best_fid = photos_list[0][-1]["file_id"]
                    raw_finfo, _ = _http_get_bytes(f"https://api.telegram.org/bot{BOT_TOKEN}/getFile?file_id={best_fid}", timeout=8)
                    fpath = json.loads(raw_finfo.decode("utf-8"))["result"]["file_path"]
                    img_bytes, ctype = _http_get_bytes(f"https://api.telegram.org/file/bot{BOT_TOKEN}/{fpath}", timeout=10)
                    _USER_PHOTO_CACHE[0] = (img_bytes, ctype, now_ts)
                    self.send_response(200)
                    self.send_header("Content-type", ctype)
                    self.send_header("Cache-Control", "public, max-age=3600")
                    self.end_headers()
                    self.wfile.write(img_bytes)
                    return
            except Exception:
                pass
            # Fallback to embedded official CyberNest logo JPEG
            try:
                logo_bytes = base64.b64decode(CYBERNEST_LOGO_BASE64)
                self.send_response(200)
                self.send_header("Content-type", "image/jpeg")
                self.send_header("Cache-Control", "public, max-age=86400")
                self.end_headers()
                self.wfile.write(logo_bytes)
            except Exception:
                self.send_response(404)
                self.end_headers()
            return

        # NEW API: Load User's Real Telegram Profile Picture (/api/user_photo)
        if path == "/api/user_photo":
            qs = urllib.parse.parse_qs(parsed.query)
            uid_str = qs.get("user_id", [""])[0]
            if not uid_str.isdigit():
                self.send_response(404)
                self.end_headers()
                return
            uid_int = int(uid_str)
            now_ts = time.time()
            cached = _USER_PHOTO_CACHE.get(uid_int)
            if cached and (now_ts - cached[2] < 3600):
                self.send_response(200)
                self.send_header("Content-type", cached[1])
                self.send_header("Cache-Control", "public, max-age=3600")
                self.end_headers()
                self.wfile.write(cached[0])
                return
            try:
                raw_photos, _ = _http_get_bytes(
                    f"https://api.telegram.org/bot{BOT_TOKEN}/getUserProfilePhotos?user_id={uid_int}&limit=1",
                    timeout=8,
                )
                pdata = json.loads(raw_photos.decode("utf-8"))
                photos_list = pdata.get("result", {}).get("photos", [])
                if not photos_list:
                    self.send_response(404)
                    self.end_headers()
                    return
                best_fid = photos_list[0][-1]["file_id"]
                raw_finfo, _ = _http_get_bytes(f"https://api.telegram.org/bot{BOT_TOKEN}/getFile?file_id={best_fid}", timeout=8)
                fpath = json.loads(raw_finfo.decode("utf-8"))["result"]["file_path"]
                img_bytes, ctype = _http_get_bytes(f"https://api.telegram.org/file/bot{BOT_TOKEN}/{fpath}", timeout=10)
                _USER_PHOTO_CACHE[uid_int] = (img_bytes, ctype, now_ts)
                self.send_response(200)
                self.send_header("Content-type", ctype)
                self.send_header("Cache-Control", "public, max-age=3600")
                self.end_headers()
                self.wfile.write(img_bytes)
            except Exception:
                self.send_response(404)
                self.end_headers()
            return

        # Full Store & User Init API (/api/init)
        if path == "/api/init":
            uid, tg_u = self._get_auth_user()
            if not uid:
                return self._send_json({"error": "Unauthorized"}, 401)
            full_n = f"{tg_u.get('first_name', '')} {tg_u.get('last_name', '')}".strip() or "Customer"
            db_u = get_or_create_user(uid, full_n, tg_u.get("username", ""))
            conn = get_db()
            c = conn.cursor()

            c.execute("SELECT id, name, icon, position FROM categories ORDER BY position ASC, id ASC")
            cats = [dict(r) for r in c.fetchall()]

            c.execute("SELECT id, category_id, name, position, brand_image FROM subcategories ORDER BY position ASC, id ASC")
            subs = [dict(r) for r in c.fetchall()]

            c.execute("""
                SELECT p.id, p.name, p.price_bdt, p.price_usdt, p.image_file_id, p.description,
                       COALESCE(p.is_active, 1) as is_active, s.id as sub_id, s.name as sub_name,
                       s.brand_image as brand_image, s.category_id, cat.name as cat_name,
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

            c.execute("SELECT COUNT(*) as cnt FROM orders WHERE user_id = ? AND status = 'refunded'", (uid,))
            cancelled_cnt = c.fetchone()["cnt"]

            c.execute("SELECT user_id, name, username, joined_date FROM users WHERE referred_by = ? ORDER BY user_id DESC", (uid,))
            ref_users = [dict(r) for r in c.fetchall()]

            c.execute("SELECT COUNT(*) as cnt FROM notifications WHERE user_id = ? AND is_read = 0", (uid,))
            unread_n = c.fetchone()["cnt"]

            c.execute("SELECT id, title, subtitle, badge_text, image_url, target_sub_id, target_cat_id, bg_gradient FROM ad_banners WHERE is_active = 1 ORDER BY position ASC, id DESC")
            banners = [dict(r) for r in c.fetchall()]

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
            sdict["auto_sales_active"] = "1" if _AUTO_SALES_ACTIVE else "0"
            sdict["welcome_group_id"] = get_welcome_group_id()
            sdict["welcome_group_link"] = get_welcome_group_link()
            sdict["transaction_channel_id"] = get_transaction_channel_id()
            sdict["transaction_channel_link"] = get_transaction_channel_link()

            vip_info = get_vip_badge_info(ostat["spent"])
            tg_photo_url = tg_u.get("photo_url") or f"/api/user_photo?user_id={uid}"

            return self._send_json({
                "user": {
                    "user_id": uid,
                    "full_name": db_u.get("name", full_n),
                    "username": db_u.get("username", ""),
                    "photo_url": tg_photo_url,
                    "balance_bdt": float(db_u.get("balance_bdt", 0.0)),
                    "joined_date": db_u.get("joined_date", "N/A"),
                    "referred_by": db_u.get("referred_by"),
                    "referral_earnings": float(db_u.get("referral_earnings", 0.0) or 0.0),
                    "pending_referral_tk": float(db_u.get("pending_referral_tk", 0.0) or 0.0),
                    "language": db_u.get("language", "EN") or "EN",
                    "is_admin": is_admin(uid),
                    "total_orders": ostat["cnt"],
                    "total_spent": float(ostat["spent"] or 0.0),
                    "cancelled_orders": cancelled_cnt,
                    "referrals": len(ref_users),
                    "referral_list": ref_users,
                    "vip": vip_info,
                },
                "categories": cats,
                "subcategories": subs,
                "products": prods,
                "banners": banners,
                "settings": sdict,
                "menu_layout": get_menu_layout(),
                "unread_notifs": unread_n,
            })

        if path == "/api/orders/history":
            uid, _ = self._get_auth_user()
            if not uid:
                return self._send_json({"error": "Unauthorized"}, 401)
            conn = get_db()
            c = conn.cursor()
            c.execute("""
                SELECT o.id, o.receipt_no, o.product_id, o.product_name, o.category_name,
                       o.quantity, o.unit_price, o.total_bdt, o.status, o.delivered_data,
                       o.order_date as created_at, o.day_of_week,
                       p.subcategory_id as sub_id
                FROM orders o
                LEFT JOIN products p ON o.product_id = p.id
                WHERE o.user_id = ?
                ORDER BY o.id DESC LIMIT 40
            """, (uid,))
            rows = [dict(r) for r in c.fetchall()]
            conn.close()
            return self._send_json({"orders": rows})

        if path == "/api/notifications":
            uid, _ = self._get_auth_user()
            if not uid:
                return self._send_json({"error": "Unauthorized"}, 401)
            conn = get_db()
            c = conn.cursor()
            c.execute("SELECT id, title, message, is_read, created_at FROM notifications WHERE user_id = ? ORDER BY id DESC LIMIT 30", (uid,))
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
            c.execute("SELECT COUNT(*) as cnt, COALESCE(SUM(balance_bdt), 0) as total_bal FROM users")
            u_row = c.fetchone()
            u_cnt = u_row["cnt"]
            total_bal = float(u_row["total_bal"] or 0.0)

            c.execute("SELECT COUNT(*) as cnt, COALESCE(SUM(total_bdt), 0) as rev FROM orders WHERE status = 'completed'")
            o_row = c.fetchone()
            comp_orders = o_row["cnt"]
            total_rev = float(o_row["rev"] or 0.0)

            c.execute("SELECT d.*, u.name as user_name, u.username as user_username FROM deposits d LEFT JOIN users u ON d.user_id = u.user_id WHERE d.status = 'pending' ORDER BY d.id DESC")
            p_deps = [dict(r) for r in c.fetchall()]

            c.execute("SELECT o.*, u.name as user_name, u.username as user_username FROM orders o LEFT JOIN users u ON o.user_id = u.user_id WHERE o.status = 'pending_delivery' ORDER BY o.id DESC")
            p_ords = [dict(r) for r in c.fetchall()]

            c.execute("SELECT user_id, name, username, balance_bdt, joined_date, referred_by FROM users ORDER BY user_id DESC LIMIT 60")
            users_list = [dict(r) for r in c.fetchall()]

            c.execute("SELECT code, discount_pct, is_active FROM coupons ORDER BY code ASC")
            coupons_list = [dict(r) for r in c.fetchall()]

            c.execute("SELECT * FROM ad_banners ORDER BY position ASC, id DESC")
            banners_list = [dict(r) for r in c.fetchall()]
            conn.close()

            return self._send_json({
                "ok": True,
                "stats": {
                    "users": u_cnt,
                    "total_balance": total_bal,
                    "completed_orders": comp_orders,
                    "total_revenue": total_rev,
                    "pending_deposits": len(p_deps),
                    "pending_orders": len(p_ords),
                    "auto_sales_active": _AUTO_SALES_ACTIVE,
                },
                "pending_deposits": p_deps,
                "pending_orders": p_ords,
                "users": users_list,
                "coupons": coupons_list,
                "banners": banners_list,
                "menu_layout": get_menu_layout(),
            })

        self.send_response(200)
        self.send_header("Content-type", "text/plain; charset=utf-8")
        self.end_headers()
        self.wfile.write(b"CyberNest Bot & Mini App API Server is running and healthy!")

    def do_POST(self):
        global _AUTO_SALES_ACTIVE
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

            # Notification add (e.g. "Deposit started" when calculator opens)
            if path == "/api/notifications/add":
                title = str(body.get("title", "🔔 Notification")).strip()
                msg = str(body.get("message", "")).strip()
                if msg:
                    add_notification(uid, title, msg)
                return self._send_json({"ok": True})

            # Save User Language preference
            if path == "/api/user/preferences":
                lang = str(body.get("language", "EN")).strip().upper()
                conn = get_db()
                c = conn.cursor()
                c.execute("UPDATE users SET language = ? WHERE user_id = ?", (lang, uid))
                conn.commit()
                conn.close()
                return self._send_json({"ok": True, "language": lang})

            # Live AI Support Chatbot Endpoint
            if path == "/api/ai_support":
                question = str(body.get("question", "")).strip()
                ans = generate_smart_ai_reply(question, user_id=uid)
                return self._send_json({"ok": True, "reply": ans})

            # Coupon Verification
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

            # 2FA Live OTP Generator
            if path == "/api/tools/2fa":
                sec = str(body.get("secret", "")).replace(" ", "").upper()
                try:
                    totp = pyotp.TOTP(sec)
                    rem = 30 - (int(datetime.datetime.now().timestamp()) % 30)
                    return self._send_json({"ok": True, "code": totp.now(), "remaining": rem})
                except Exception:
                    return self._send_json({"ok": False, "error": "Invalid 2FA Secret Key"})

            # Outlook / Hotmail IMAP OTP Reader
            if path == "/api/tools/outlook":
                res = fetch_outlook_otp(str(body.get("account", "")))
                return self._send_json(res)

            # NEW API: Resend Purchased Product to Telegram Inbox (/api/order/resend_product)
            if path == "/api/order/resend_product":
                oid = int(body.get("order_id", 0))
                conn = get_db()
                c = conn.cursor()
                c.execute("SELECT * FROM orders WHERE id = ? AND user_id = ?", (oid, uid))
                o = c.fetchone()
                conn.close()
                if not o:
                    return self._send_json({"ok": False, "error": "Order not found"})
                if not o["delivered_data"]:
                    return self._send_json({"ok": False, "error": "This order is still pending delivery."})

                raw_deliv = o["delivered_data"]
                try:
                    items = json.loads(raw_deliv) if raw_deliv.startswith("[") else [ln for ln in raw_deliv.splitlines() if ln.strip()]
                except Exception:
                    items = [ln for ln in raw_deliv.splitlines() if ln.strip()]

                if _GLOBAL_EVENT_LOOP and _GLOBAL_BOT_APP:
                    card_msg = format_credentials_card(items, o["product_name"], o["category_name"] or "Service")
                    asyncio.run_coroutine_threadsafe(
                        _GLOBAL_BOT_APP.bot.send_message(
                            chat_id=uid,
                            text=f"📦 <b>Re-sent Order #{o['id']} ({o['receipt_no']})</b>\n\n{card_msg}",
                            parse_mode=ParseMode.HTML,
                        ),
                        _GLOBAL_EVENT_LOOP,
                    )
                    if len(items) > 1 or "MAIL" in (o["category_name"] or "").upper():
                        try:
                            excel_bio = generate_excel_delivery(items, o["product_name"], o["receipt_no"])
                            asyncio.run_coroutine_threadsafe(
                                _GLOBAL_BOT_APP.bot.send_document(
                                    chat_id=uid,
                                    document=InputFile(excel_bio, filename=f"{o['receipt_no']}_Credentials.xlsx"),
                                    caption=f"📂 <b>Excel Backup:</b> <code>{o['receipt_no']}</code>",
                                    parse_mode=ParseMode.HTML,
                                ),
                                _GLOBAL_EVENT_LOOP,
                            )
                        except Exception:
                            pass
                return self._send_json({"ok": True, "message": "📦 প্রোডাক্টটি আপনার টেলিগ্রাম ইনবক্সে আবার পাঠানো হয়েছে!"})

            # NEW API: Resend Official PDF Receipt to Telegram Inbox (/api/order/resend_pdf)
            if path == "/api/order/resend_pdf":
                oid = int(body.get("order_id", 0))
                conn = get_db()
                c = conn.cursor()
                c.execute("SELECT o.*, u.name as uname, u.username as u_username FROM orders o LEFT JOIN users u ON o.user_id = u.user_id WHERE o.id = ? AND o.user_id = ?", (oid, uid))
                o = c.fetchone()
                conn.close()
                if not o:
                    return self._send_json({"ok": False, "error": "Order not found"})

                if _GLOBAL_EVENT_LOOP and _GLOBAL_BOT_APP:
                    bot_uname = _GLOBAL_BOT_APP.bot.username or "bot"
                    pdf_bio = generate_pdf_receipt(
                        receipt_no=o["receipt_no"],
                        date_str=o["order_date"],
                        day_name=o["day_of_week"] or "N/A",
                        customer_name=o["uname"] or tg_u.get("first_name", "Customer"),
                        telegram_id=uid,
                        username=o["u_username"] or tg_u.get("username", ""),
                        item_desc=o["product_name"],
                        category_name=o["category_name"] or "DIGITAL",
                        qty=o["quantity"],
                        unit_price=o["unit_price"] or o["total_bdt"],
                        total_bdt=o["total_bdt"],
                        bot_username=bot_uname,
                    )
                    asyncio.run_coroutine_threadsafe(
                        _GLOBAL_BOT_APP.bot.send_document(
                            chat_id=uid,
                            document=InputFile(pdf_bio, filename=f"Receipt_{o['receipt_no']}.pdf"),
                            caption=f"🧾 <b>Official Payment Receipt:</b> <code>{o['receipt_no']}</code>",
                            parse_mode=ParseMode.HTML,
                        ),
                        _GLOBAL_EVENT_LOOP,
                    )
                return self._send_json({"ok": True, "message": "🧾 অফিশিয়াল PDF রিসিট আপনার টেলিগ্রাম ইনবক্সে পাঠানো হয়েছে!"})

            # Deposit Request Submission
            if path == "/api/deposit":
                method = str(body.get("method", "bKash"))
                amount = float(body.get("amount", 0))
                trx_id = str(body.get("trx_id", "")).strip()
                min_dep = float(get_setting("min_deposit_manual", "20.0"))
                if amount < min_dep or not trx_id:
                    return self._send_json({"ok": False, "error": f"Minimum deposit is ৳{min_dep:.0f} BDT and valid TrxID is required."})
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
                add_notification(uid, "💳 Deposit Submitted", f"{method} ৳{amount:.2f} BDT (TrxID: {trx_id}) is pending admin verification.")
                if _GLOBAL_EVENT_LOOP and _GLOBAL_BOT_APP:
                    kb = InlineKeyboardMarkup([[
                        InlineKeyboardButton(f"✅ Approve #{dep_id}", callback_data=f"adm_dep_approve_{dep_id}"),
                        InlineKeyboardButton("❌ Reject", callback_data=f"adm_dep_reject_{dep_id}")
                    ]])
                    asyncio.run_coroutine_threadsafe(
                        _GLOBAL_BOT_APP.bot.send_message(
                            chat_id=SUPER_ADMIN_ID,
                            text=(
                                f"🔔 <b>Mini App Deposit Request #{dep_id}</b>\n"
                                f"👤 User: {tg_u.get('first_name', 'User')} (<code>{uid}</code>)\n"
                                f"💳 Method: <b>{method}</b>\n"
                                f"💰 Amount: <b>{amount:.2f} BDT</b>\n"
                                f"🔑 TrxID: <code>{trx_id}</code>"
                            ),
                            parse_mode=ParseMode.HTML,
                            reply_markup=kb,
                        ),
                        _GLOBAL_EVENT_LOOP
                    )
                return self._send_json({"ok": True})

            # Order Placement from Mini App (with Full Bot Inbox Delivery + PDF Receipt + Excel + Channel Post)
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
                c.execute("SELECT balance_bdt, name, username FROM users WHERE user_id = ?", (uid,))
                urow = c.fetchone()
                bal = float(urow["balance_bdt"] or 0.0)
                u_fullname = urow["name"] or tg_u.get("first_name", "Customer")
                u_username = urow["username"] or tg_u.get("username", "")

                if bal < total_bdt:
                    conn.close()
                    return self._send_json({"ok": False, "error": f"Insufficient balance! Need ৳{total_bdt:.2f} BDT"})

                c.execute("SELECT id, content FROM stock WHERE product_id = ? AND is_sold = 0 LIMIT ?", (pid, qty))
                srows = c.fetchall()
                is_vpn_or_proxy = ("vpn" in prod["cat_name"].lower()) or ("proxy" in prod["cat_name"].lower())
                if len(srows) < qty and not is_vpn_or_proxy:
                    conn.close()
                    return self._send_json({"ok": False, "error": "Out of stock! Please check back later."})

                now = datetime.datetime.now()
                now_str = now.strftime("%Y-%m-%d %I:%M:%S %p")
                days_bn = {"Monday": "সোমবার", "Tuesday": "মঙ্গলবার", "Wednesday": "বুধবার", "Thursday": "বৃহস্পতিবার", "Friday": "শুক্রবার", "Saturday": "শনিবার", "Sunday": "রবিবার"}
                day_name = days_bn.get(now.strftime("%A"), now.strftime("%A"))
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
                    delivered_json = json.dumps(items)
                    c.execute("""
                        INSERT INTO orders (receipt_no, user_id, product_id, product_name, category_name, quantity, unit_price, total_bdt, total_usdt, status, delivered_data, order_date, day_of_week)
                        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, 'completed', ?, ?, ?)
                    """, (receipt_no, uid, pid, full_name, prod["cat_name"], qty, prod["price_bdt"], total_bdt, total_bdt / 125.0, delivered_json, now_str, day_name))
                    conn.commit()
                    conn.close()
                    add_notification(uid, "✅ Order Delivered!", f"{full_name} (x{qty}) — Receipt {receipt_no} delivered to your Telegram inbox.")

                    if _GLOBAL_EVENT_LOOP and _GLOBAL_BOT_APP:
                        async def _deliver_web_auto():
                            bot_inst = _GLOBAL_BOT_APP.bot
                            bot_info = await bot_inst.get_me()
                            if qty == 1:
                                msg_card = format_credentials_card(items, full_name, prod["cat_name"])
                                await bot_inst.send_message(chat_id=uid, text=msg_card, parse_mode=ParseMode.HTML)
                                if "MAIL" in prod["cat_name"].upper():
                                    try:
                                        ex_bio = generate_excel_delivery(items, full_name, receipt_no)
                                        await bot_inst.send_document(
                                            chat_id=uid,
                                            document=InputFile(ex_bio, filename=f"{receipt_no}_{prod['name'].replace(' ', '_')}.xlsx"),
                                            caption=f"📂 <b>Excel Data Backup:</b> <code>{receipt_no}</code>",
                                            parse_mode=ParseMode.HTML,
                                        )
                                    except Exception:
                                        pass
                            else:
                                bulk_msg = (
                                    f"🎉 <b>আপনার অর্ডার সফলভাবে সম্পন্ন হয়েছে!</b>\n"
                                    f"─────────────────────────────\n"
                                    f"🧾 <b>Receipt:</b> <code>{receipt_no}</code>\n"
                                    f"📦 <b>প্রোডাক্ট:</b> {full_name}\n"
                                    f"🔢 <b>পরিমাণ:</b> {qty} টি | <b>মোট পরিশোধ:</b> {format_price(total_bdt)} BDT\n"
                                    f"─────────────────────────────\n"
                                    f"📂 <b>আপনার সবকটি {qty} টি অ্যাকাউন্ট/মেইল নিচের এক্সেল ফাইলে সুন্দরভাবে সাজিয়ে দেওয়া হলো।</b>"
                                )
                                await bot_inst.send_message(chat_id=uid, text=bulk_msg, parse_mode=ParseMode.HTML)
                                try:
                                    ex_bio = generate_excel_delivery(items, full_name, receipt_no)
                                    await bot_inst.send_document(
                                        chat_id=uid,
                                        document=InputFile(ex_bio, filename=f"{receipt_no}_{prod['name'].replace(' ', '_')}.xlsx"),
                                        caption=f"📂 <b>Accounts Excel File:</b> <code>{receipt_no}</code> ({qty} pcs)",
                                        parse_mode=ParseMode.HTML,
                                    )
                                except Exception:
                                    pass
                            try:
                                pdf_bio = generate_pdf_receipt(
                                    receipt_no=receipt_no, date_str=now_str, day_name=day_name,
                                    customer_name=u_fullname, telegram_id=uid, username=u_username,
                                    item_desc=full_name, category_name=prod["cat_name"],
                                    qty=qty, unit_price=prod["price_bdt"], total_bdt=total_bdt,
                                    bot_username=bot_info.username,
                                )
                                await bot_inst.send_document(
                                    chat_id=uid,
                                    document=InputFile(pdf_bio, filename=f"Receipt_{receipt_no}.pdf"),
                                    caption=f"🧾 <b>Official Payment Receipt:</b> <code>{receipt_no}</code>",
                                    parse_mode=ParseMode.HTML,
                                )
                            except Exception:
                                pass
                            try:
                                await bot_inst.send_message(
                                    chat_id=SUPER_ADMIN_ID,
                                    text=(
                                        f"🔔 <b>Mini App প্রোডাক্ট সেল হয়েছে!</b>\n"
                                        f"🧾 Receipt: <code>{receipt_no}</code>\n"
                                        f"👤 কাস্টমার: {u_fullname} (ID: <code>{uid}</code>)\n"
                                        f"📦 প্রোডাক্ট: {full_name}\n"
                                        f"🔢 পরিমাণ: {qty} টি | মূল্য: {total_bdt:.2f} BDT"
                                    ),
                                    parse_mode=ParseMode.HTML,
                                )
                            except Exception:
                                pass
                            await post_transaction_to_channel(bot_inst, receipt_no, u_fullname, uid, full_name, qty, total_bdt)

                        asyncio.run_coroutine_threadsafe(_deliver_web_auto(), _GLOBAL_EVENT_LOOP)
                    return self._send_json({"ok": True, "message": "✅ অর্ডার সফল! আপনার টেলিগ্রাম ইনবক্সে One-Tap Copy নোট ও অফিশিয়াল PDF রিসিট পাঠানো হয়েছে।"})
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
                        async def _notify_web_manual():
                            bot_inst = _GLOBAL_BOT_APP.bot
                            bot_info = await bot_inst.get_me()
                            confirm_msg = (
                                f"🎉 <b>আপনার অর্ডারটি সফলভাবে গ্রহণ করা হয়েছে!</b>\n"
                                f"─────────────────────────────\n"
                                f"🧾 <b>অর্ডার নম্বর:</b> #{oid} (Receipt: <code>{receipt_no}</code>)\n"
                                f"🛡️ <b>প্রোডাক্ট:</b> {full_name}\n"
                                f"🔢 <b>পরিমাণ:</b> {qty} টি | <b>মোট পরিশোধ:</b> {format_price(total_bdt)} BDT\n"
                                f"📅 <b>তারিখ:</b> {now_str} ({day_name})\n"
                                f"─────────────────────────────\n"
                                f"⏳ <b>ডেলিভারি সময়:</b> আগামী ১৫ থেকে ৩০ মিনিটের মধ্যে এডমিন আপনার ইনবক্সে সার্ভিসটি ডেলিভারি পাঠাবে।"
                            )
                            await bot_inst.send_message(chat_id=uid, text=confirm_msg, parse_mode=ParseMode.HTML)
                            try:
                                pdf_bio = generate_pdf_receipt(
                                    receipt_no=receipt_no, date_str=now_str, day_name=day_name,
                                    customer_name=u_fullname, telegram_id=uid, username=u_username,
                                    item_desc=full_name, category_name=prod["cat_name"],
                                    qty=qty, unit_price=prod["price_bdt"], total_bdt=total_bdt,
                                    bot_username=bot_info.username,
                                )
                                await bot_inst.send_document(
                                    chat_id=uid,
                                    document=InputFile(pdf_bio, filename=f"Receipt_{receipt_no}.pdf"),
                                    caption=f"🧾 <b>Official Payment Receipt:</b> <code>{receipt_no}</code>",
                                    parse_mode=ParseMode.HTML,
                                )
                            except Exception:
                                pass
                            adm_kb = InlineKeyboardMarkup([
                                [InlineKeyboardButton(f"🚀 Deliver Order #{oid}", callback_data=f"adm_fulfill_{oid}")],
                                [InlineKeyboardButton(f"❌ Cancel & Refund #{oid}", callback_data=f"adm_refund_{oid}")]
                            ])
                            await bot_inst.send_message(
                                chat_id=SUPER_ADMIN_ID,
                                text=(
                                    f"🔔 <b>Mini App ম্যানুয়াল অর্ডার #{oid}</b>\n"
                                    f"🧾 Receipt: <code>{receipt_no}</code>\n"
                                    f"🛡️ {full_name} (x{qty})\n"
                                    f"💰 Paid: {total_bdt:.2f} BDT\n"
                                    f"👤 User: {u_fullname} (<code>{uid}</code>)"
                                ),
                                parse_mode=ParseMode.HTML,
                                reply_markup=adm_kb,
                            )

                        asyncio.run_coroutine_threadsafe(_notify_web_manual(), _GLOBAL_EVENT_LOOP)
                    return self._send_json({"ok": True, "message": f"🎉 অর্ডার #{oid} গ্রহণ করা হয়েছে! অফিশিয়াল PDF রিসিট আপনার ইনবক্সে পাঠানো হয়েছে এবং ১৫-৩০ মিনিটের মধ্যে ডেলিভারি পাবেন।"})

            # ==================================================================
            # SUPER ADMIN API ROUTES (/api/admin/*) — Restricted to 8919985167
            # ==================================================================
            if path.startswith("/api/admin/"):
                if not is_admin(uid):
                    return self._send_json({"ok": False, "error": "Forbidden"}, 403)
                conn = get_db()
                c = conn.cursor()

                # 1. Manual Order Delivery (Supports separate VPN email/pass or Proxy string + PDF receipt)
                if path == "/api/admin/deliver_order":
                    oid = int(body.get("order_id", 0))
                    ddata = str(body.get("data", "")).strip()
                    vpn_email = str(body.get("vpn_email", "")).strip()
                    vpn_pass = str(body.get("vpn_pass", "")).strip()
                    if vpn_email and vpn_pass:
                        ddata = f"{vpn_email}:{vpn_pass}"
                    if not ddata:
                        conn.close()
                        return self._send_json({"ok": False, "error": "Credentials required"})

                    c.execute("SELECT o.*, u.name as uname, u.username as u_username, p.image_file_id FROM orders o LEFT JOIN users u ON o.user_id = u.user_id LEFT JOIN products p ON o.product_id = p.id WHERE o.id = ?", (oid,))
                    o = c.fetchone()
                    if not o:
                        conn.close()
                        return self._send_json({"ok": False, "error": "Order not found"})
                    c.execute("UPDATE orders SET status = 'completed', delivered_data = ? WHERE id = ?", (ddata, oid))
                    conn.commit()
                    conn.close()
                    add_notification(o["user_id"], "✅ Order Delivered!", f"Order #{oid} ({o['product_name']}) has been delivered!")

                    if _GLOBAL_EVENT_LOOP and _GLOBAL_BOT_APP:
                        async def _send_manual_deliv():
                            bot_inst = _GLOBAL_BOT_APP.bot
                            bot_info = await bot_inst.get_me()
                            card = format_credentials_card([ddata], o["product_name"], o["category_name"] or "Service")
                            if o["image_file_id"]:
                                try:
                                    await bot_inst.send_photo(chat_id=o["user_id"], photo=o["image_file_id"], caption=card, parse_mode=ParseMode.HTML)
                                except Exception:
                                    await bot_inst.send_message(chat_id=o["user_id"], text=card, parse_mode=ParseMode.HTML)
                            else:
                                await bot_inst.send_message(chat_id=o["user_id"], text=card, parse_mode=ParseMode.HTML)
                            try:
                                pdf_bio = generate_pdf_receipt(
                                    receipt_no=o["receipt_no"], date_str=o["order_date"], day_name=o["day_of_week"] or "N/A",
                                    customer_name=o["uname"] or "Customer", telegram_id=o["user_id"], username=o["u_username"] or "",
                                    item_desc=o["product_name"], category_name=o["category_name"] or "DIGITAL",
                                    qty=o["quantity"], unit_price=o["unit_price"] or o["total_bdt"], total_bdt=o["total_bdt"],
                                    bot_username=bot_info.username,
                                )
                                await bot_inst.send_document(
                                    chat_id=o["user_id"],
                                    document=InputFile(pdf_bio, filename=f"Receipt_{o['receipt_no']}.pdf"),
                                    caption=f"🧾 <b>Verified Payment Receipt:</b> <code>{o['receipt_no']}</code>",
                                    parse_mode=ParseMode.HTML,
                                )
                            except Exception:
                                pass
                            await post_transaction_to_channel(bot_inst, o["receipt_no"], o["uname"] or "Customer", o["user_id"], o["product_name"], o["quantity"], o["total_bdt"])

                        asyncio.run_coroutine_threadsafe(_send_manual_deliv(), _GLOBAL_EVENT_LOOP)
                    return self._send_json({"ok": True})

                # 2. Refund Pending Order
                if path == "/api/admin/refund_order":
                    oid = int(body.get("order_id", 0))
                    c.execute("SELECT * FROM orders WHERE id = ?", (oid,))
                    o = c.fetchone()
                    if o and o["status"] == "pending_delivery":
                        c.execute("UPDATE orders SET status = 'refunded' WHERE id = ?", (oid,))
                        c.execute("UPDATE users SET balance_bdt = balance_bdt + ? WHERE user_id = ?", (o["total_bdt"], o["user_id"]))
                        conn.commit()
                        add_notification(o["user_id"], "❌ Order Refunded", f"Order #{oid} refunded: +{o['total_bdt']:.2f} BDT")
                        if _GLOBAL_EVENT_LOOP and _GLOBAL_BOT_APP:
                            asyncio.run_coroutine_threadsafe(
                                _GLOBAL_BOT_APP.bot.send_message(
                                    chat_id=o["user_id"],
                                    text=f"❌ আপনার অর্ডার #{oid} বাতিল করা হয়েছে এবং <b>{o['total_bdt']:.2f} BDT</b> আপনার ওয়ালেটে রিফান্ড করা হয়েছে।",
                                    parse_mode=ParseMode.HTML,
                                ),
                                _GLOBAL_EVENT_LOOP,
                            )
                    conn.close()
                    return self._send_json({"ok": True})

                # 3. Approve / Reject Deposit
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
                            if _GLOBAL_EVENT_LOOP and _GLOBAL_BOT_APP:
                                asyncio.run_coroutine_threadsafe(
                                    _GLOBAL_BOT_APP.bot.send_message(
                                        chat_id=d["user_id"],
                                        text=f"🎉 <b>ডিপোজিট সফল!</b>\nআপনার অ্যাকাউন্টে <b>{d['amount']:.2f} BDT</b> যুক্ত হয়েছে।",
                                        parse_mode=ParseMode.HTML,
                                    ),
                                    _GLOBAL_EVENT_LOOP,
                                )
                        else:
                            c.execute("UPDATE deposits SET status = 'rejected' WHERE id = ?", (did,))
                            add_notification(d["user_id"], "❌ Deposit Rejected", f"Deposit #{did} ({d['amount']:.2f} BDT) was rejected.")
                        conn.commit()
                    conn.close()
                    return self._send_json({"ok": True})

                # 4. NEW API: Receipt Checker & One-Click PDF Download (/api/admin/check_receipt)
                if path == "/api/admin/check_receipt":
                    rcode = str(body.get("receipt_no", "")).strip().upper()
                    send_pdf = bool(body.get("send_pdf", False))
                    if not rcode:
                        conn.close()
                        return self._send_json({"ok": False, "error": "Enter a Receipt Code (e.g. RCP-XXXXX)"})
                    if not rcode.startswith("RCP-") and len(rcode) == 10:
                        rcode = f"RCP-{rcode}"
                    c.execute("""
                        SELECT o.*, u.name as customer_name, u.username as customer_username
                        FROM orders o
                        LEFT JOIN users u ON o.user_id = u.user_id
                        WHERE UPPER(o.receipt_no) LIKE ?
                        ORDER BY o.id DESC LIMIT 1
                    """, (f"%{rcode}%",))
                    o = c.fetchone()
                    conn.close()
                    if not o:
                        return self._send_json({"ok": False, "authentic": False, "error": "❌ নকল বা অবৈধ রিসিট! এই কোডের কোনো অর্ডার ডাটাবেসে নেই।"})

                    order_dict = dict(o)
                    if send_pdf and _GLOBAL_EVENT_LOOP and _GLOBAL_BOT_APP:
                        bot_uname = _GLOBAL_BOT_APP.bot.username or "bot"
                        pdf_bio = generate_pdf_receipt(
                            receipt_no=o["receipt_no"], date_str=o["order_date"], day_name=o["day_of_week"] or "N/A",
                            customer_name=o["customer_name"] or "Customer", telegram_id=o["user_id"], username=o["customer_username"] or "",
                            item_desc=o["product_name"], category_name=o["category_name"] or "DIGITAL",
                            qty=o["quantity"], unit_price=o["unit_price"] or o["total_bdt"], total_bdt=o["total_bdt"],
                            bot_username=bot_uname,
                        )
                        asyncio.run_coroutine_threadsafe(
                            _GLOBAL_BOT_APP.bot.send_document(
                                chat_id=uid,
                                document=InputFile(pdf_bio, filename=f"Receipt_{o['receipt_no']}.pdf"),
                                caption=f"🧾 <b>Verified Receipt PDF:</b> <code>{o['receipt_no']}</code>",
                                parse_mode=ParseMode.HTML,
                            ),
                            _GLOBAL_EVENT_LOOP,
                        )
                    return self._send_json({"ok": True, "authentic": True, "order": order_dict})

                # 5. NEW API: Rotating Ad Banners Management (/api/admin/banners)
                if path == "/api/admin/banners":
                    action = str(body.get("action", "add"))
                    if action == "add":
                        title = str(body.get("title", "")).strip() or "Special Promo"
                        subtitle = str(body.get("subtitle", "")).strip()
                        badge_text = str(body.get("badge_text", "🔥 HOT OFFER")).strip()
                        image_url = str(body.get("image_url", "")).strip()
                        target_sub_id = int(body.get("target_sub_id", 0) or 0)
                        target_cat_id = int(body.get("target_cat_id", 0) or 0)
                        bg_gradient = str(body.get("bg_gradient", "linear-gradient(135deg, #2e0249 0%, #570a57 50%, #ff6b00 100%)")).strip()
                        c.execute("""
                            INSERT INTO ad_banners (title, subtitle, badge_text, image_url, target_sub_id, target_cat_id, bg_gradient, is_active, position)
                            VALUES (?, ?, ?, ?, ?, ?, ?, 1, 0)
                        """, (title, subtitle, badge_text, image_url, target_sub_id, target_cat_id, bg_gradient))
                        conn.commit()
                    elif action == "delete":
                        bid = int(body.get("banner_id", 0))
                        c.execute("DELETE FROM ad_banners WHERE id = ?", (bid,))
                        conn.commit()
                    elif action == "toggle":
                        bid = int(body.get("banner_id", 0))
                        c.execute("UPDATE ad_banners SET is_active = CASE WHEN COALESCE(is_active,1)=1 THEN 0 ELSE 1 END WHERE id = ?", (bid,))
                        conn.commit()
                    c.execute("SELECT * FROM ad_banners ORDER BY position ASC, id DESC")
                    blist = [dict(r) for r in c.fetchall()]
                    conn.close()
                    return self._send_json({"ok": True, "banners": blist})

                # 6. NEW API: Step-by-Step Multi-Button Broadcast from WebApp (/api/admin/broadcast)
                if path == "/api/admin/broadcast":
                    msg_text = str(body.get("message", "")).strip()
                    photo_id = str(body.get("photo", "")).strip() or None
                    buttons = body.get("buttons", [])
                    if not msg_text and not photo_id:
                        conn.close()
                        return self._send_json({"ok": False, "error": "Message or photo required"})

                    c.execute("SELECT user_id FROM users")
                    all_u = [r["user_id"] for r in c.fetchall()]
                    conn.close()

                    valid_btns = []
                    if isinstance(buttons, list):
                        for b in buttons[:10]:
                            bt = str(b.get("text", "")).strip()
                            bu = str(b.get("url", "")).strip()
                            if bt and (bu.startswith("http://") or bu.startswith("https://") or bu.startswith("tg://")):
                                valid_btns.append({"text": bt, "url": bu})

                    if _GLOBAL_EVENT_LOOP and _GLOBAL_BOT_APP:
                        async def _run_web_bcast():
                            bot_inst = _GLOBAL_BOT_APP.bot
                            rm = InlineKeyboardMarkup([[InlineKeyboardButton(b["text"], url=b["url"])] for b in valid_btns]) if valid_btns else None
                            sent_c = 0
                            for target_u in all_u:
                                try:
                                    if photo_id:
                                        await bot_inst.send_photo(chat_id=target_u, photo=photo_id, caption=msg_text, parse_mode=ParseMode.HTML, reply_markup=rm)
                                    else:
                                        await bot_inst.send_message(chat_id=target_u, text=msg_text, parse_mode=ParseMode.HTML, reply_markup=rm)
                                    sent_c += 1
                                except Exception:
                                    pass
                            try:
                                if photo_id:
                                    await bot_inst.send_photo(chat_id=get_welcome_group_id(), photo=photo_id, caption=msg_text, parse_mode=ParseMode.HTML, reply_markup=rm)
                                else:
                                    await bot_inst.send_message(chat_id=get_welcome_group_id(), text=msg_text, parse_mode=ParseMode.HTML, reply_markup=rm)
                            except Exception:
                                pass
                            try:
                                await bot_inst.send_message(
                                    chat_id=SUPER_ADMIN_ID,
                                    text=f"✅ <b>Web Admin Broadcast Completed!</b>\n👥 Delivered to: <b>{sent_c}</b> users + Update Group.",
                                    parse_mode=ParseMode.HTML,
                                )
                            except Exception:
                                pass

                        asyncio.run_coroutine_threadsafe(_run_web_bcast(), _GLOBAL_EVENT_LOOP)
                    return self._send_json({"ok": True, "total_targets": len(all_u)})

                # 6.5 Brand Photo Update (1 photo per brand for Website, keeps Bot's per-plan photos separate)
                if path == "/api/admin/brand_photo":
                    sid = int(body.get("sub_id", 0))
                    img_val = str(body.get("brand_image", "")).strip() or None
                    if sid:
                        c.execute("UPDATE subcategories SET brand_image = ? WHERE id = ?", (img_val, sid))
                        conn.commit()
                    conn.close()
                    return self._send_json({"ok": True})

                # 6.6 AI API Key Settings & Live Connection Test (Syncs Bot + Website automatically)
                if path == "/api/admin/ai_settings":
                    act = str(body.get("action", "save"))
                    new_key = str(body.get("ai_api_key", "")).strip()
                    if act == "save" and new_key:
                        set_setting("ai_api_key", new_key)
                    elif act == "clear":
                        set_setting("ai_api_key", "")
                        conn.close()
                        return self._send_json({"ok": True, "message": "AI Key cleared."})
                    conn.close()
                    test_key = new_key or get_setting("ai_api_key", "").strip() or os.getenv("GROQ_API_KEY", "").strip() or os.getenv("GEMINI_API_KEY", "").strip()
                    if not test_key:
                        return self._send_json({"ok": False, "error": "No API Key found. Paste a Groq (gsk_...) or Gemini (AIza...) key."})
                    if test_key.startswith("gsk_"):
                        try:
                            tp = json.dumps({
                                "model": "llama-3.3-70b-versatile",
                                "messages": [{"role": "user", "content": "Say 'AI Connected Successfully!' in Bangla in 1 short sentence."}],
                                "max_tokens": 60,
                            }).encode("utf-8")
                            tr = urllib.request.Request(
                                "https://api.groq.com/openai/v1/chat/completions",
                                data=tp,
                                headers={"Content-Type": "application/json", "Authorization": f"Bearer {test_key}", "User-Agent": "CyberNestBot/2.0"},
                            )
                            with urllib.request.urlopen(tr, timeout=10) as resp:
                                td = json.loads(resp.read().decode("utf-8"))
                            return self._send_json({"ok": True, "provider": "GroqCloud", "reply": td["choices"][0]["message"]["content"].strip()})
                        except Exception as e:
                            return self._send_json({"ok": False, "error": f"Groq Error: {e}"})
                    elif test_key.startswith("AIza"):
                        try:
                            tp = json.dumps({"contents": [{"role": "user", "parts": [{"text": "Say 'AI Connected Successfully!' in Bangla in 1 short sentence."}]}]}).encode("utf-8")
                            tr = urllib.request.Request(
                                f"https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key={test_key}",
                                data=tp,
                                headers={"Content-Type": "application/json", "User-Agent": "CyberNestBot/2.0"},
                            )
                            with urllib.request.urlopen(tr, timeout=10) as resp:
                                td = json.loads(resp.read().decode("utf-8"))
                            return self._send_json({"ok": True, "provider": "Google Gemini", "reply": td["candidates"][0]["content"]["parts"][0]["text"].strip()})
                        except Exception as e:
                            return self._send_json({"ok": False, "error": f"Gemini Error: {e}"})
                    else:
                        return self._send_json({"ok": False, "error": "Invalid Key format. Groq keys start with gsk_ and Gemini keys start with AIza."})

                # 7. Product & Catalog Management (Edit / Create / Delete Category, Brand, Plan)
                if path == "/api/admin/product_update":
                    pid = int(body.get("product_id", 0))
                    if body.get("toggle_active"):
                        c.execute("UPDATE products SET is_active = CASE WHEN COALESCE(is_active,1)=1 THEN 0 ELSE 1 END WHERE id = ?", (pid,))
                    if "price_bdt" in body and body["price_bdt"] is not None:
                        pr = float(body["price_bdt"])
                        rate = float(get_setting("usdt_rate", "125.0") or 125.0)
                        c.execute("UPDATE products SET price_bdt = ?, price_usdt = ? WHERE id = ?", (pr, pr / rate, pid))
                    if "name" in body and str(body["name"]).strip():
                        c.execute("UPDATE products SET name = ? WHERE id = ?", (str(body["name"]).strip(), pid))
                    if "description" in body:
                        c.execute("UPDATE products SET description = ? WHERE id = ?", (str(body["description"]).strip(), pid))
                    if "image_file_id" in body:
                        img_val = str(body["image_file_id"]).strip() or None
                        c.execute("UPDATE products SET image_file_id = ? WHERE id = ?", (img_val, pid))
                        if body.get("apply_to_brand"):
                            c.execute("SELECT subcategory_id FROM products WHERE id = ?", (pid,))
                            sr = c.fetchone()
                            if sr:
                                c.execute("UPDATE products SET image_file_id = ? WHERE subcategory_id = ?", (img_val, sr["subcategory_id"]))
                    conn.commit()
                    conn.close()
                    return self._send_json({"ok": True})

                if path == "/api/admin/catalog_manage":
                    act = str(body.get("action", ""))
                    if act == "add_category":
                        name = str(body.get("name", "")).strip()
                        icon = str(body.get("icon", "🛍️")).strip() or "🛍️"
                        if name:
                            c.execute("INSERT INTO categories (name, icon) VALUES (?, ?)", (name, icon))
                    elif act == "rename_category":
                        cid = int(body.get("category_id", 0))
                        name = str(body.get("name", "")).strip()
                        if cid and name:
                            c.execute("UPDATE categories SET name = ? WHERE id = ?", (name, cid))
                    elif act == "delete_category":
                        cid = int(body.get("category_id", 0))
                        c.execute("DELETE FROM stock WHERE product_id IN (SELECT p.id FROM products p JOIN subcategories s ON p.subcategory_id = s.id WHERE s.category_id = ?)", (cid,))
                        c.execute("DELETE FROM products WHERE subcategory_id IN (SELECT id FROM subcategories WHERE category_id = ?)", (cid,))
                        c.execute("DELETE FROM subcategories WHERE category_id = ?", (cid,))
                        c.execute("DELETE FROM categories WHERE id = ?", (cid,))
                    elif act == "add_brand":
                        cid = int(body.get("category_id", 0))
                        name = str(body.get("name", "")).strip()
                        if cid and name:
                            c.execute("INSERT INTO subcategories (category_id, name) VALUES (?, ?)", (cid, name))
                    elif act == "rename_brand":
                        sid = int(body.get("sub_id", 0))
                        name = str(body.get("name", "")).strip()
                        if sid and name:
                            c.execute("UPDATE subcategories SET name = ? WHERE id = ?", (name, sid))
                    elif act == "delete_brand":
                        sid = int(body.get("sub_id", 0))
                        c.execute("DELETE FROM stock WHERE product_id IN (SELECT id FROM products WHERE subcategory_id = ?)", (sid,))
                        c.execute("DELETE FROM products WHERE subcategory_id = ?", (sid,))
                        c.execute("DELETE FROM subcategories WHERE id = ?", (sid,))
                    elif act == "add_product":
                        sid = int(body.get("sub_id", 0))
                        name = str(body.get("name", "")).strip()
                        price_bdt = float(body.get("price_bdt", 0))
                        desc = str(body.get("description", "")).strip()
                        rate = float(get_setting("usdt_rate", "125.0") or 125.0)
                        if sid and name:
                            c.execute(
                                "INSERT INTO products (subcategory_id, name, price_bdt, price_usdt, description, is_on_demand, is_active) VALUES (?, ?, ?, ?, ?, 0, 1)",
                                (sid, name, price_bdt, price_bdt / rate, desc),
                            )
                    elif act == "delete_product":
                        pid = int(body.get("product_id", 0))
                        c.execute("DELETE FROM stock WHERE product_id = ?", (pid,))
                        c.execute("DELETE FROM products WHERE id = ?", (pid,))
                    conn.commit()
                    conn.close()
                    return self._send_json({"ok": True})

                # 8. Add Stock
                if path == "/api/admin/add_stock":
                    pid = int(body.get("product_id", 0))
                    lines = [ln.strip() for ln in str(body.get("lines", "")).splitlines() if ln.strip()]
                    for ln in lines:
                        c.execute("INSERT INTO stock (product_id, content, is_sold) VALUES (?, ?, 0)", (pid, ln))
                    conn.commit()
                    c.execute("SELECT COUNT(*) as cnt FROM stock WHERE product_id = ? AND is_sold = 0", (pid,))
                    total_now = c.fetchone()["cnt"]
                    conn.close()
                    return self._send_json({"ok": True, "added": len(lines), "total_stock": total_now})

                # 9. Payment Settings Update
                if path == "/api/admin/payment_settings":
                    for k in ("bkash_number", "nagad_number", "binance_id", "min_deposit_manual", "usdt_rate", "support_username"):
                        if k in body and body[k] is not None:
                            set_setting(k, str(body[k]).strip())
                    conn.close()
                    return self._send_json({"ok": True})

                # 10. User Management (Search + Add/Deduct Balance)
                if path == "/api/admin/user_balance":
                    target_uid = int(body.get("user_id", 0))
                    action = str(body.get("action", "add"))
                    amount = float(body.get("amount", 0))
                    if not target_uid or amount <= 0:
                        conn.close()
                        return self._send_json({"ok": False, "error": "Valid user_id and amount required"})
                    if action == "add":
                        c.execute("UPDATE users SET balance_bdt = balance_bdt + ? WHERE user_id = ?", (amount, target_uid))
                        add_notification(target_uid, "💰 Balance Added", f"Admin added +৳{amount:.2f} BDT to your wallet.")
                    else:
                        if conn.is_pg:
                            c.execute("UPDATE users SET balance_bdt = GREATEST(0, balance_bdt - %s) WHERE user_id = %s", (amount, target_uid))
                        else:
                            c.execute("UPDATE users SET balance_bdt = MAX(0, balance_bdt - ?) WHERE user_id = ?", (amount, target_uid))
                        add_notification(target_uid, "💸 Balance Deducted", f"Admin deducted -৳{amount:.2f} BDT from your wallet.")
                    conn.commit()
                    c.execute("SELECT user_id, name, username, balance_bdt FROM users WHERE user_id = ?", (target_uid,))
                    updated_u = c.fetchone()
                    conn.close()
                    return self._send_json({"ok": True, "user": dict(updated_u) if updated_u else None})

                # 11. Channel TX & Group Settings + 4-Step Custom Channel TX Post
                if path == "/api/admin/channel_tx":
                    act = str(body.get("action", ""))
                    if act == "toggle_autosales":
                        _AUTO_SALES_ACTIVE = not _AUTO_SALES_ACTIVE
                        conn.close()
                        return self._send_json({"ok": True, "auto_sales_active": _AUTO_SALES_ACTIVE})
                    elif act == "save_gc_settings":
                        for k in ("welcome_group_id", "welcome_group_link", "transaction_channel_id", "transaction_channel_link"):
                            if k in body and body[k] is not None:
                                set_setting(k, str(body[k]).strip())
                        conn.close()
                        return self._send_json({"ok": True})
                    elif act == "custom_tx_post":
                        cname = str(body.get("customer_name", "Customer")).strip()
                        pname = str(body.get("product_name", "Premium VPN")).strip()
                        qty = max(1, int(body.get("quantity", 1)))
                        price = float(body.get("total_bdt", 50.0))
                        fake_uid = random.randint(7100000000, 7999999999)
                        fake_receipt = f"RCP-{secrets.token_hex(5).upper()}"
                        conn.close()
                        if _GLOBAL_EVENT_LOOP and _GLOBAL_BOT_APP:
                            asyncio.run_coroutine_threadsafe(
                                post_transaction_to_channel(_GLOBAL_BOT_APP.bot, fake_receipt, cname, fake_uid, pname, qty, price),
                                _GLOBAL_EVENT_LOOP,
                            )
                        return self._send_json({"ok": True, "receipt_no": fake_receipt})
                    conn.close()
                    return self._send_json({"ok": False, "error": "Invalid action"})

                # 12. Menu Layout & Factory Reset
                if path == "/api/admin/menu_layout":
                    if body.get("reset"):
                        save_menu_layout([list(r) for r in DEFAULT_MENU_LAYOUT])
                    elif "layout" in body and isinstance(body["layout"], list):
                        save_menu_layout(body["layout"])
                    conn.close()
                    return self._send_json({"ok": True, "menu_layout": get_menu_layout()})

                if path == "/api/admin/factory_reset":
                    confirm_code = str(body.get("confirm", "")).strip()
                    if confirm_code != "RESET":
                        conn.close()
                        return self._send_json({"ok": False, "error": "Type RESET to confirm"})
                    c.execute("DELETE FROM orders")
                    c.execute("DELETE FROM stock")
                    c.execute("DELETE FROM deposits")
                    c.execute("UPDATE users SET balance_bdt = 0")
                    conn.commit()
                    conn.close()
                    return self._send_json({"ok": True})

                # 13. Popup & Coupon Controls
                if path == "/api/admin/popup":
                    set_setting("popup_enabled", str(body.get("enabled", "0")))
                    set_setting("popup_title", str(body.get("title", "")))
                    set_setting("popup_text", str(body.get("text", "")))
                    set_setting("popup_image", str(body.get("image", "")))
                    conn.close()
                    return self._send_json({"ok": True})

                if path == "/api/admin/coupon":
                    act = str(body.get("action", "add"))
                    code_s = str(body.get("code", "")).strip().upper()
                    if act == "delete" and code_s:
                        c.execute("DELETE FROM coupons WHERE code = ?", (code_s,))
                        conn.commit()
                        conn.close()
                        return self._send_json({"ok": True})
                    pct = float(body.get("discount_pct", 0))
                    if conn.is_pg:
                        c.execute("INSERT INTO coupons (code, discount_pct, is_active) VALUES (%s, %s, 1) ON CONFLICT (code) DO UPDATE SET discount_pct = EXCLUDED.discount_pct, is_active = 1", (code_s, pct))
                    else:
                        c.execute("INSERT OR REPLACE INTO coupons (code, discount_pct, is_active) VALUES (?, ?, 1)", (code_s, pct))
                    conn.commit()
                    conn.close()
                    return self._send_json({"ok": True})

                conn.close()

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
        state_data TEXT,
        referred_by BIGINT DEFAULT NULL,
        referral_earnings REAL DEFAULT 0.0,
        pending_referral_tk REAL DEFAULT 0.0,
        language TEXT DEFAULT 'EN'
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
        brand_image TEXT DEFAULT NULL,
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

    cursor.execute("""
    CREATE TABLE IF NOT EXISTS ad_banners (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        title TEXT NOT NULL,
        subtitle TEXT,
        badge_text TEXT DEFAULT '⚡ SPECIAL OFFER',
        image_url TEXT DEFAULT '',
        target_sub_id INTEGER DEFAULT 0,
        target_cat_id INTEGER DEFAULT 0,
        bg_gradient TEXT DEFAULT 'linear-gradient(135deg, #2e0249 0%, #570a57 50%, #ff6b00 100%)',
        is_active INTEGER DEFAULT 1,
        position INTEGER DEFAULT 0
    )
    """)

    # Safe Schema Migrations for Existing Databases
    if conn.is_pg:
        try:
            cursor.execute("ALTER TABLE categories ADD COLUMN IF NOT EXISTS position INTEGER DEFAULT 0")
            cursor.execute("ALTER TABLE subcategories ADD COLUMN IF NOT EXISTS position INTEGER DEFAULT 0")
            cursor.execute("ALTER TABLE subcategories ADD COLUMN IF NOT EXISTS brand_image TEXT DEFAULT NULL")
            cursor.execute("ALTER TABLE stock ADD COLUMN IF NOT EXISTS stock_type TEXT DEFAULT 'account'")
            cursor.execute("ALTER TABLE products ADD COLUMN IF NOT EXISTS is_active INTEGER DEFAULT 1")
            cursor.execute("ALTER TABLE users ADD COLUMN IF NOT EXISTS referred_by BIGINT DEFAULT NULL")
            cursor.execute("ALTER TABLE users ADD COLUMN IF NOT EXISTS referral_earnings REAL DEFAULT 0.0")
            cursor.execute("ALTER TABLE users ADD COLUMN IF NOT EXISTS pending_referral_tk REAL DEFAULT 0.0")
            cursor.execute("ALTER TABLE users ADD COLUMN IF NOT EXISTS language TEXT DEFAULT 'EN'")
        except Exception as e:
            logger.warning(f"Postgres column check notice: {e}")
    else:
        for tbl in ("categories", "subcategories"):
            try:
                cursor.execute(f"ALTER TABLE {tbl} ADD COLUMN position INTEGER DEFAULT 0")
            except Exception:
                pass
        for col_sql in (
            "ALTER TABLE stock ADD COLUMN stock_type TEXT DEFAULT 'account'",
            "ALTER TABLE products ADD COLUMN is_active INTEGER DEFAULT 1",
            "ALTER TABLE subcategories ADD COLUMN brand_image TEXT DEFAULT NULL",
            "ALTER TABLE users ADD COLUMN referred_by BIGINT DEFAULT NULL",
            "ALTER TABLE users ADD COLUMN referral_earnings REAL DEFAULT 0.0",
            "ALTER TABLE users ADD COLUMN pending_referral_tk REAL DEFAULT 0.0",
            "ALTER TABLE users ADD COLUMN language TEXT DEFAULT 'EN'",
        ):
            try:
                cursor.execute(col_sql)
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

    # Seed Default Rotating Ad Banners if empty
    cursor.execute("SELECT COUNT(*) as cnt FROM ad_banners")
    b_row = cursor.fetchone()
    b_cnt = b_row["cnt"] if isinstance(b_row, dict) else b_row[0]
    if b_cnt == 0:
        default_banners = [
            ("CyberNest Digital Mart", "Fastest VPN, Residential Proxy & Verified Mail Store — 24/7 Instant Delivery", "⚡ 24/7 AUTOMATED STORE", "", 0, 0, "linear-gradient(135deg, #2e0249 0%, #570a57 50%, #ff6b00 100%)"),
            ("Premium High-Speed VPNs", "Nord, Express, Surfshark, Proton & More — 100% Verified with Warranty", "🛡️ PREMIUM VPN DEALS", "", 0, 0, "linear-gradient(135deg, #0f172a 0%, #1e3a8a 55%, #0ea5e9 100%)"),
            ("Clean Residential Proxies", "Rapid Proxy, 9 Proxy & 711 Proxy — Select 100MB to 1GB Packages", "🌐 RESIDENTIAL PROXY", "", 0, 0, "linear-gradient(135deg, #064e3b 0%, #047857 55%, #10b981 100%)"),
        ]
        for idx, (bt, bsub, btag, bimg, bsubid, bcatid, bgrad) in enumerate(default_banners):
            cursor.execute("""
                INSERT INTO ad_banners (title, subtitle, badge_text, image_url, target_sub_id, target_cat_id, bg_gradient, is_active, position)
                VALUES (?, ?, ?, ?, ?, ?, ?, 1, ?)
            """, (bt, bsub, btag, bimg, bsubid, bcatid, bgrad, idx))
        conn.commit()

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
        "Rapid Proxy": [("100 MB", 15.0), ("200 MB", 25.0), ("500 MB", 60.0), ("1 GB", 115.0)],
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

def get_or_create_user(user_id: int, name: str, username: str, referrer_id: Optional[int] = None) -> Dict[str, Any]:
    conn = get_db()
    c = conn.cursor()
    c.execute("SELECT * FROM users WHERE user_id = ?", (user_id,))
    row = c.fetchone()
    is_new_user = False
    if not row:
        is_new_user = True
        joined = datetime.datetime.now().strftime("%Y-%m-%d")
        is_adm = 1 if user_id == SUPER_ADMIN_ID else 0
        valid_ref = None
        if referrer_id and int(referrer_id) != int(user_id):
            c.execute("SELECT user_id FROM users WHERE user_id = ?", (int(referrer_id),))
            if c.fetchone():
                valid_ref = int(referrer_id)
        c.execute(
            "INSERT INTO users (user_id, name, username, balance_bdt, joined_date, is_admin, referred_by) VALUES (?, ?, ?, ?, ?, ?, ?)",
            (user_id, name, username or "", 0.0, joined, is_adm, valid_ref),
        )
        conn.commit()
        if valid_ref:
            add_notification(valid_ref, "👥 New Referral Joined!", f"{name} (@{username or 'user'}) joined using your referral link!")
        c.execute("SELECT * FROM users WHERE user_id = ?", (user_id,))
        row = c.fetchone()
    else:
        updates = []
        params = []
        if user_id == SUPER_ADMIN_ID and row["is_admin"] != 1:
            updates.append("is_admin = 1")
        if name and row["name"] != name:
            updates.append("name = ?")
            params.append(name)
        if username and row["username"] != username:
            updates.append("username = ?")
            params.append(username)
        if updates:
            params.append(user_id)
            c.execute(f"UPDATE users SET {', '.join(updates)} WHERE user_id = ?", tuple(params))
            conn.commit()
            c.execute("SELECT * FROM users WHERE user_id = ?", (user_id,))
            row = c.fetchone()
    conn.close()
    res = dict(row)
    res["_is_new_user"] = is_new_user
    return res

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
# 5. KEYBOARD LAYOUTS (Includes 🤖 AI Assistant Button)
# ------------------------------------------------------------------------------
DEFAULT_MENU_LAYOUT = [
    ["🛍️ Buy Products", "🤖 AI Assistant"],
    ["👤 My Profile", "💳 Deposit"],
    ["🔐 Get Code", "💬 Support"],
]

def get_menu_layout() -> List[List[str]]:
    raw = get_setting("main_menu_layout", "")
    if raw:
        try:
            parsed = json.loads(raw)
            if isinstance(parsed, list) and parsed:
                # Ensure 🤖 AI Assistant is present at least once unless admin explicitly customized after v2
                flat = [btn for r in parsed for btn in r]
                if "🤖 AI Assistant" not in flat and not get_setting("ai_btn_migrated", ""):
                    parsed[0].append("🤖 AI Assistant")
                    set_setting("ai_btn_migrated", "1")
                    save_menu_layout(parsed)
                return parsed
        except Exception:
            pass
    return [list(r) for r in DEFAULT_MENU_LAYOUT]

def save_menu_layout(layout: List[List[str]]):
    clean = [r for r in layout if r]
    set_setting("main_menu_layout", json.dumps(clean, ensure_ascii=False))
    set_setting("ai_btn_migrated", "1")

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

    clean_bot = (bot_username or "bot").lstrip('@')
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
# 7. USER HANDLERS (New vs Returning Welcome, AI Assistant, Shop, Profile, Deposit)
# ------------------------------------------------------------------------------
def build_welcome_message(full_name: str, shop_name: str, is_new: bool) -> str:
    safe_name = html.escape(full_name or "গ্রাহক")
    if is_new:
        return (
            f"🌟 <b>স্বাগতম, {safe_name}!</b> ✨\n"
            f"💎 <b>{shop_name} — ২৪/৭ অটোমেটেড ডিজিটাল শপ</b>\n"
            f"─────────────────────────────\n"
            f"আমাদের প্ল্যাটফর্মে আপনি পাচ্ছেন শতভাগ ভেরিফায়েড ও প্রিমিয়াম ডিজিটাল সার্ভিস:\n\n"
            f"🛡️ <b>প্রিমিয়াম VPN সার্ভিস:</b> ফুল ওয়ারেন্টি ও হাই-স্পিড সিকিউরিটি।\n"
            f"🌐 <b>প্রিমিয়াম Proxy সার্ভিস:</b> ক্লিন ও সুপার-ফাস্ট রেসিডেনশিয়াল প্রক্সি।\n"
            f"✉️ <b>প্রিমিয়াম Mail সার্ভিস:</b> ইনস্ট্যান্ট অটো-ডেলিভারি ও এক্সেল ব্যাকআপ।\n"
            f"🛍️ <b>স্মার্ট ওয়েব স্টোর (Mini App):</b> এক ক্লিকে লাইভ শপিং ও অর্ডার ট্র্যাকিং।\n"
            f"🔐 <b>ফ্রি 2FA ও মেইল টুলস:</b> লাইভ 2FA কোড এবং ইনস্ট্যান্ট মেইল OTP রিডার।\n"
            f"🤖 <b>২৪/৭ AI সাপোর্ট ও রেফারেল:</b> যেকোনো প্রশ্নের তাৎক্ষণিক উত্তর ও বন্ধুদের আমন্ত্রণ জানিয়ে আয়ের সুযোগ।\n"
            f"─────────────────────────────\n"
            f"👇 নিচের মেনু বাটন অথবা বাম কোণার <b>Open Shop</b> বাটনে চাপ দিয়ে আপনার কেনাকাটা শুরু করুন!"
        )
    else:
        return (
            f"👋 <b>স্বাগতম, {safe_name}!</b> ✨\n"
            f"💎 <b>{shop_name}</b> — ২৪/৭ অটোমেটেড ডিজিটাল স্টোর।\n"
            f"─────────────────────────────\n"
            f"🛍️ প্রিমিয়াম VPN, প্রিমিয়াম Proxy ও প্রিমিয়াম Mail কিনতে নিচের মেনু অথবা <b>Open Shop</b> বাটন ব্যবহার করুন 👇"
        )

async def start_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user = update.effective_user
    ref_id = None
    if context.args and len(context.args) > 0:
        arg = context.args[0].strip()
        if arg.startswith("order_"):
            receipt_code = arg.replace("order_", "")
            db_u = get_or_create_user(user.id, user.full_name, user.username)
            await show_order_verification(update, receipt_code)
            return
        elif arg.startswith("ref_"):
            r_str = arg.replace("ref_", "").strip()
            if r_str.isdigit():
                ref_id = int(r_str)
        elif arg.isdigit():
            ref_id = int(arg)

    db_user = get_or_create_user(user.id, user.full_name, user.username, referrer_id=ref_id)
    is_new = bool(db_user.get("_is_new_user", False))
    shop_name = get_setting("shop_name", "CyberNest Market")

    in_group, in_channel = await check_user_membership(context.bot, user.id)
    if not (in_group and in_channel):
        force_txt = (
            f"👋 <b>স্বাগতম, {html.escape(user.full_name or 'বন্ধু')}!</b>\n\n"
            f"বট ব্যবহার করতে নিচের বাটনে চাপ দিয়ে গ্রুপ ও চ্যানেলে যুক্ত হয়ে ভেরিফাই করুন:"
        )
        kb = get_force_join_keyboard(in_group, in_channel)
        await update.message.reply_text(force_txt, parse_mode=ParseMode.HTML, reply_markup=kb)
        return

    welcome_text = build_welcome_message(user.full_name, shop_name, is_new=is_new)
    reply_kb = get_main_reply_keyboard(user.id)
    await update.message.reply_text(welcome_text, parse_mode=ParseMode.HTML, reply_markup=reply_kb)

async def ai_assistant_menu_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user = update.effective_user
    set_user_state(user.id, "WAITING_AI_QUESTION", {})
    txt = (
        "🤖 <b>CyberNest AI Assistant</b>\n"
        "─────────────────────────────\n"
        "আসসালামু আলাইকুম! আপনার কোনো প্রশ্ন বা অসুবিধা থাকলে নিচে লিখে জিজ্ঞেস করুন—আমি সাথে সাথে সাহায্য করব।"
    )
    await update.message.reply_text(txt, parse_mode=ParseMode.HTML)

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

    grp_status = "✅ আপডেট গ্রুপে পাঠানো হয়েছে"
    try:
        if photo_id:
            await bot.send_photo(chat_id=get_welcome_group_id(), photo=photo_id, caption=msg_text, parse_mode=ParseMode.HTML, reply_markup=reply_markup)
        else:
            await bot.send_message(chat_id=get_welcome_group_id(), text=msg_text, parse_mode=ParseMode.HTML, reply_markup=reply_markup)
    except Exception as e:
        grp_status = f"⚠️ আপডেট গ্রুপে পাঠানো যায়নি ({e})"

    await reply_target.reply_text(
        f"✅ <b>ব্রডকাস্ট সম্পন্ন!</b>\n"
        f"👥 মোট ইউজার: <b>{sent_cnt}</b> জনের ইনবক্সে পৌঁছেছে।\n"
        f"📢 {grp_status}।\n"
        f"🛡️ <i>ট্রানজ্যাকশন চ্যানেল সম্পূর্ণ পরিষ্কার রাখা হয়েছে।</i>",
        parse_mode=ParseMode.HTML,
    )

async def show_order_verification(update: Update, receipt_no: str):
    conn = get_db()
    c = conn.cursor()
    c.execute("SELECT * FROM orders WHERE receipt_no = ?", (receipt_no,))
    order = c.fetchone()
    conn.close()

    if not order:
        await update.message.reply_text("❌ দুঃখিত! এই রিসিট নম্বরের কোনো অর্ডার পাওয়া যায়নি।")
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

    conn = get_db()
    c = conn.cursor()
    c.execute("SELECT COALESCE(SUM(total_bdt), 0) as spent FROM orders WHERE user_id = ? AND status != 'refunded'", (user.id,))
    spent = float(c.fetchone()["spent"] or 0.0)
    conn.close()
    vip = get_vip_badge_info(spent)

    profile_text = (
        f"👤 <b>My Profile</b> ({vip['badge']})\n\n"
        f"👤 <b>Name</b> : {db_user['name']}\n"
        f"🆔 <b>User ID</b> : <code>{user.id}</code>\n"
        f"💰 <b>Balance</b> : {balance_bdt:.2f} BDT / {balance_usdt:.4f} USDT\n"
        f"🛍️ <b>Total Spent</b> : {spent:.2f} BDT\n"
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
        "হটমেইল/আউটলুক ওটিপি কোড এবং লাইভ 2FA কোড পাওয়ার জন্য নিচের অপশন থেকে নির্বাচন করুন:"
    )
    kb = InlineKeyboardMarkup([
        [InlineKeyboardButton("🔑 2FA Authenticator (Live Code)", callback_data="tool_2fa_prompt")],
        [InlineKeyboardButton("✉️ Outlook / Hotmail OTP Reader", callback_data="tool_outlook_prompt")],
        [InlineKeyboardButton("🌐 Hotmail & Outlook Code Server", url="https://code.yamin.bd/")],
    ])
    await update.message.reply_text(text, parse_mode=ParseMode.HTML, reply_markup=kb)

async def support_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    text = (
        "💬 <b>কাস্টমার সাপোর্ট</b>\n\n"
        "যেকোনো সমস্যায় নিচে এডমিনকে মেসেজ দিন অথবা 🤖 AI Assistant ব্যবহার করুন:"
    )
    kb = InlineKeyboardMarkup([
        [InlineKeyboardButton("📨 Message Admin", url=f"tg://user?id={SUPER_ADMIN_ID}")],
        [InlineKeyboardButton("🤖 Ask AI Assistant", callback_data="ai_q_open")],
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
# 8. INLINE CALLBACK QUERY DISPATCHER (Users, AI Assistant & Admin Nested Flow)
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
                    await query.answer("🎉 ভেরিফিকেশন সফল হয়েছে!", show_alert=True)
                except Exception:
                    pass
                try:
                    await query.message.delete()
                except Exception:
                    pass
                shop_name = get_setting("shop_name", "CyberNest Market")
                welcome_text = build_welcome_message(query.from_user.full_name, shop_name, is_new=True)
                reply_kb = get_main_reply_keyboard(user_id)
                await query.message.reply_text(welcome_text, parse_mode=ParseMode.HTML, reply_markup=reply_kb)
            else:
                try:
                    await query.answer("❌ আপনি এখনো গ্রুপ বা চ্যানেলে জয়েন করেননি! দয়া করে জয়েন করে আবার চাপুন।", show_alert=True)
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

    # AI Quick Question Callbacks
    if data.startswith("ai_q_"):
        q_type = data.replace("ai_q_", "")
        if q_type == "exit":
            set_user_state(user_id, None)
            await query.message.reply_text("✅ AI Assistant মোড বন্ধ করা হয়েছে।")
            return
        if q_type == "open":
            set_user_state(user_id, "WAITING_AI_QUESTION", {})
            await query.message.reply_text("🤖 আসসালামু আলাইকুম! আপনার কোনো প্রশ্ন বা অসুবিধা থাকলে নিচে লিখে জিজ্ঞেস করুন।")
            return
        map_q = {
            "vpn": "vpn price list",
            "proxy": "proxy packages",
            "mail": "mail stock",
            "deposit": "how to deposit",
        }
        ans = generate_smart_ai_reply(map_q.get(q_type, "help"), user_id=user_id)
        await query.message.reply_text(ans, parse_mode=ParseMode.HTML)
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
        await query.message.reply_text("❌ কেনাকাটা বাতিল করা হয়েছে। নতুন অর্ডারের জন্য মেনু ব্যবহার করুন।")
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
            await query.message.reply_text("⚠️ মেনুটি আপডেট করা হয়েছে! অনুগ্রহ করে নিচের 🛍️ Buy Products বাটনে আবার চাপুন।")
            return
        title = f"{meta['cat_name']} ➔ {meta['sub_name']}"
        text = f"<b>{title}</b>\n\nনিচের প্ল্যান বা মেয়াদ নির্বাচন করুন:"
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
        await query.message.reply_text("✏️ <b>Enter Quantity:</b>\nআপনি কতগুলো কিনতে চান তা সংখ্যায় লিখে পাঠান (যেমন: 5, 20, 50):", parse_mode=ParseMode.HTML)
        return

    if data.startswith("calc_confirm_"):
        in_group, in_channel = await check_user_membership(context.bot, user_id)
        if not (in_group and in_channel):
            kb = get_force_join_keyboard(in_group, in_channel)
            await query.message.reply_text(
                "⚠️ কিনতে হলে আগে নিচের বাটনে চাপ দিয়ে গ্রুপ ও চ্যানেলে জয়েন করে ভেরিফাই করুন:",
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
            "✏️ <b>Custom Quantity:</b>\nআপনি কতগুলো কিনতে চান তা সংখ্যায় লিখে পাঠান (যেমন: 5, 20, 50):",
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

    if data == "tool_outlook_prompt":
        set_user_state(user_id, "WAITING_OUTLOOK_STR", {})
        await query.message.reply_text(
            "✉️ <b>Outlook / Hotmail Live OTP Reader:</b>\n"
            "আপনার মেইল স্ট্রিংটি <code>email|password|refresh_token|client_id</code> ফরম্যাটে পেস্ট করে পাঠান:",
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
            await query.message.reply_text("🛒 আপনার পূর্বের কোনো কেনাকাটার হিস্টোরি পাওয়া যায়নি।")
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
            await query.message.reply_text("❌ অর্ডারটি পাওয়া যায়নি।")
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
            creds_display = "<i>⏳ অর্ডারটি প্রসেসিং হচ্ছে, খুব শীঘ্রই ডেলিভারি দেওয়া হবে।</i>"

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
            [InlineKeyboardButton("📦 Resend Product", callback_data=f"prof_resend_prod_{o['id']}"), InlineKeyboardButton("🧾 Resend Receipt PDF", callback_data=f"prof_resend_pdf_{o['id']}")],
            [InlineKeyboardButton("🔄 Buy Again", callback_data=f"usr_prod_{o['product_id']}")],
            [InlineKeyboardButton("< Back", callback_data="prof_purchases")],
        ]
        await query.message.reply_text(txt, parse_mode=ParseMode.HTML, reply_markup=InlineKeyboardMarkup(kb))
        return

    if data.startswith("prof_resend_prod_"):
        order_id = int(data.split("_")[3])
        conn = get_db()
        c = conn.cursor()
        c.execute("SELECT * FROM orders WHERE id = ? AND user_id = ?", (order_id, user_id))
        o = c.fetchone()
        conn.close()
        if not o or not o["delivered_data"]:
            await query.message.reply_text("⏳ প্রোডাক্ট এখনো ডেলিভারি হয়নি বা অর্ডার পাওয়া যায়নি।")
            return
        try:
            items = json.loads(o["delivered_data"]) if o["delivered_data"].startswith("[") else [ln for ln in o["delivered_data"].splitlines() if ln.strip()]
        except Exception:
            items = [o["delivered_data"]]
        card = format_credentials_card(items, o["product_name"], o["category_name"] or "Service")
        await query.message.reply_text(card, parse_mode=ParseMode.HTML)
        return

    if data.startswith("prof_resend_pdf_"):
        order_id = int(data.split("_")[3])
        conn = get_db()
        c = conn.cursor()
        c.execute("SELECT * FROM orders WHERE id = ? AND user_id = ?", (order_id, user_id))
        o = c.fetchone()
        conn.close()

        if not o:
            await query.message.reply_text("অর্ডার পাওয়া যায়নি।")
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
            await query.message.reply_text("❌ রিসিট তৈরি করতে সমস্যা হয়েছে।")
        return

    if data == "prof_deposit_history":
        conn = get_db()
        c = conn.cursor()
        c.execute("SELECT * FROM deposits WHERE user_id = ? ORDER BY id DESC LIMIT 10", (user_id,))
        deps = c.fetchall()
        conn.close()
        if not deps:
            await query.message.reply_text("🔄 আপনার কোনো ডিপোজিট হিস্টোরি পাওয়া যায়নি।")
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
        conn = get_db()
        c = conn.cursor()
        c.execute("SELECT COUNT(*) as cnt FROM users WHERE referred_by = ?", (user_id,))
        r_cnt = c.fetchone()["cnt"]
        conn.close()
        msg = (
            f"👥 <b>রেফারেল ও আর্নিং প্রোগ্রাম</b>\n"
            f"─────────────────────────────\n"
            f"👤 মোট রেফারেল: <b>{r_cnt} জন</b>\n"
            f"💰 মোট রেফারেল আয়: <b>{float(db_user.get('referral_earnings', 0) or 0):.2f} BDT</b>\n\n"
            f"🔗 আপনার রেফার লিংক (ট্যাপ করে কপি করুন):\n<code>{ref_link}</code>"
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
        f"🔔 <b>নতুন প্রোডাক্ট সেল হয়েছে!</b>\n"
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
        "⚡ হাই-স্পিড প্রিমিয়াম VPN অ্যাকাউন্ট — ১৫–৩০ মিনিটের মধ্যে ইনবক্সে Email ও Password পাবেন।" if "vpn" in cat_low
        else "🚀 আল্ট্রা-ফাস্ট ও ক্লিন রেসিডেনশিয়াল Proxy — হাই অ্যানোনিমিটি ও ১০০% ওয়ার্কিং গ্যারান্টি।" if "proxy" in cat_low
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
            await query.message.reply_text("প্রোডাক্ট পাওয়া যায়নি।")
            return

        total_bdt = prod["price_bdt"] * quantity
        total_usdt = prod["price_usdt"] * quantity

        if user_balance < total_bdt:
            conn.close()
            shortage = total_bdt - user_balance
            await query.message.reply_text(
                f"❌ <b>অপর্যাপ্ত ব্যালেন্স!</b>\n\n"
                f"প্রয়োজনীয় টাকা: {total_bdt:.2f} BDT\n"
                f"আপনার ব্যালেন্স: {user_balance:.2f} BDT\n"
                f"ঘাটতি: {shortage:.2f} BDT\n\n"
                f"দয়া করে নিচের 💳 Deposit মেনু থেকে ব্যালেন্স রিচার্জ করুন।",
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
                f"বর্তমানে স্টকে আছে: <b>{len(stock_rows)}</b> টি। অনুগ্রহ করে পরিমাণ কমিয়ে দিন বা পরে চেষ্টা করুন।",
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
                    f"🎉 <b>আপনার অর্ডার সফলভাবে সম্পন্ন হয়েছে!</b>\n"
                    f"─────────────────────────────\n"
                    f"🧾 <b>Receipt:</b> <code>{receipt_no}</code>\n"
                    f"📦 <b>প্রোডাক্ট:</b> {full_product_name}\n"
                    f"🔢 <b>পরিমাণ:</b> {quantity} টি | <b>মোট পরিশোধ:</b> {format_price(total_bdt)} BDT\n"
                    f"─────────────────────────────\n"
                    f"📂 <b>আপনার সবকটি {quantity} টি অ্যাকাউন্ট/মেইল নিচের এক্সেল ফাইলে সুন্দরভাবে সাজিয়ে দেওয়া হলো।</b>"
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
                        f"🔔 <b>নতুন প্রোডাক্ট সেল হয়েছে!</b>\n"
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
            f"🎉 <b>আপনার অর্ডারটি সফলভাবে গ্রহণ করা হয়েছে!</b>\n"
            f"─────────────────────────────\n"
            f"🧾 <b>অর্ডার নম্বর:</b> #{order_id} (Receipt: <code>{receipt_no}</code>)\n"
            f"🛡️ <b>প্রোডাক্ট:</b> {full_product_name}\n"
            f"🔢 <b>পরিমাণ:</b> {quantity} টি | <b>মোট পরিশোধ:</b> {format_price(total_bdt)} BDT\n"
            f"📅 <b>তারিখ:</b> {now_str} ({day_name})\n"
            f"─────────────────────────────\n"
            f"⏳ <b>ডেলিভারি সময়:</b> আগামী ১৫ থেকে ৩০ মিনিটের মধ্যে এডমিন আপনার ইনবক্সে সার্ভিসটি ডেলিভারি পাঠাবে।"
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
                    f"🔔 <b>নতুন ম্যানুয়াল অর্ডার প্রয়োজন!</b>\n"
                    f"─────────────────────────────\n"
                    f"🧾 <b>অর্ডার নম্বর:</b> #{order_id} (Receipt: <code>{receipt_no}</code>)\n"
                    f"👤 <b>কাস্টমার:</b> {user_name} (ID: <code>{user_id}</code>)\n"
                    f"📦 <b>প্রোডাক্ট:</b> {full_product_name}\n"
                    f"🔢 <b>পরিমাণ:</b> {quantity} টি | <b>মূল্য:</b> {format_price(total_bdt)} BDT\n"
                    f"📅 <b>সময়:</b> {now_str}\n"
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
            await query.message.reply_text(f"❌ অর্ডার প্রসেস করতে সাময়িক সমস্যা হয়েছে: {e}")
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
        f"সব কাজ নেস্টেড বাটন দিয়ে ৩-৪ ধাপে পরিচালনা করুন:\n\n"
        f"💳 পেন্ডিং ডিপোজিট: <b>{pend_dep}</b> টি\n"
        f"🛡️ পেন্ডিং ম্যানুয়াল অর্ডার: <b>{pend_ord}</b> টি"
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
        state_str = "চালু করা হয়েছে 🟢" if _AUTO_SALES_ACTIVE else "বন্ধ করা হয়েছে 🔴"
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
            "💡 <i>গ্রুপ বা চ্যানেলে গিয়ে <code>/id</code> লিখলেই সঠিক আইডি পেয়ে যাবেন।</i>"
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
            "🔢 <b>আপনি কয়টি বাটন অ্যাড করতে চান?</b>\n"
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
        state_str = "চালু করা হয়েছে 🟢" if _AUTO_SALES_ACTIVE else "বন্ধ করা হয়েছে 🔴"
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
            f"🛡️ পেন্ডিং ম্যানুয়াল অর্ডার: <b>{pend_ord}</b> টি"
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
        await query.message.reply_text("✅ ডিফল্ট লেআউট রিসেট করা হয়েছে!", reply_markup=get_main_reply_keyboard(user_id))
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
            [InlineKeyboardButton("🗑️ Hide / Remove Button", callback_data=f"adm_ml_mv_{r_idx}_{c_idx}_remove")],
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
            elif direction == "remove":
                layout[r_idx].pop(c_idx)
            save_menu_layout(layout)
            await query.message.reply_text(f"✅ <b>{lbl}</b> আপডেট করা হয়েছে!", parse_mode=ParseMode.HTML, reply_markup=get_main_reply_keyboard(user_id))
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
            await query.message.reply_text(f"🗑️ ব্র্যান্ড '{sub_row['name']}' মুছে ফেলা হয়েছে!")
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
            await query.message.reply_text(f"🗑️ ক্যাটাগরি '{cat_row['name']}' মুছে ফেলা হয়েছে!")
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
        await query.message.reply_text("নতুন নাম বা মেয়াদ লিখুন:")
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
        await query.edit_message_text("✅ প্রোডাক্টটি মুছে ফেলা হয়েছে।")
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
        await query.message.reply_text("➕ নতুন মেয়াদ বা প্যাকেজের নাম লিখুন:")
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
        await query.edit_message_text("✅ ছবি মুছে ফেলা হয়েছে।")
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
            await query.edit_message_text(f"✅ ডিপোজিট #{dep_id} অনুমোদিত হয়েছে।")
            try:
                await context.bot.send_message(chat_id=dep["user_id"], text=f"🎉 <b>ডিপোজিট সফল!</b>\nআপনার অ্যাকাউন্টে {dep['amount']:.2f} BDT যুক্ত হয়েছে।", parse_mode=ParseMode.HTML)
            except Exception:
                pass
        elif action == "reject":
            c.execute("UPDATE deposits SET status = 'rejected' WHERE id = ?", (dep_id,))
            conn.commit()
            conn.close()
            await query.edit_message_text(f"❌ ডিপোজিট #{dep_id} বাতিল করা হয়েছে।")
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
            await query.edit_message_text(f"❌ Order #{order_id} বাতিল ও রিফান্ড করা হয়েছে।")
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
            f"💰 মোট বিক্রয় আয়: <b>{o_stat['total_rev']:.2f} BDT</b>\n"
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
        set_user_state(user.id, None)
        await buy_products_menu_handler(update, context)
        return
    if text == "🤖 AI Assistant":
        await ai_assistant_menu_handler(update, context)
        return
    if text == "👤 My Profile":
        set_user_state(user.id, None)
        await profile_handler(update, context)
        return
    if text == "💳 Deposit":
        await deposit_menu_handler(update, context)
        return
    if text == "🔐 Get Code":
        set_user_state(user.id, None)
        await get_code_menu_handler(update, context)
        return
    if text == "💬 Support":
        set_user_state(user.id, None)
        await support_handler(update, context)
        return
    if text == "👑 Admin Panel":
        if is_admin(user.id):
            set_user_state(user.id, None)
            await admin_panel_menu_handler(update, context)
        return

    if state == "WAITING_AI_QUESTION":
        reply_txt = generate_smart_ai_reply(text, user_id=user.id)
        await update.message.reply_text(reply_txt, parse_mode=ParseMode.HTML)
        return

    if state == "WAITING_OUTLOOK_STR":
        res = fetch_outlook_otp(text)
        set_user_state(user.id, None)
        if res.get("ok"):
            await update.message.reply_text(
                f"✉️ <b>Outlook / Hotmail OTP Found!</b>\n"
                f"📌 <b>Subject:</b> {html.escape(str(res.get('subject', 'Inbox')))}\n"
                f"🔢 <b>Code / Message:</b> <code>{html.escape(str(res.get('code', '')))}</code>",
                parse_mode=ParseMode.HTML,
            )
        else:
            await update.message.reply_text(f"❌ <b>Error:</b> {html.escape(str(res.get('error', 'Failed')))}", parse_mode=ParseMode.HTML)
        return

    if state == "WAITING_DEP_AMOUNT_MANUAL":
        min_man = float(get_setting("min_deposit_manual", "20.0"))
        try:
            amt = float(text)
            if amt < min_man:
                await update.message.reply_text(f"❌ সর্বনিম্ন ডিপোজিট হলো <b>{min_man:.0f} টাকা</b>। আবার লিখুন:", parse_mode=ParseMode.HTML)
                return
        except ValueError:
            await update.message.reply_text("দয়া করে সঠিক সংখ্যায় টাকার পরিমাণ লিখুন:")
            return

        set_user_state(user.id, "WAITING_DEP_TRX", {"amount": amt})
        kb = InlineKeyboardMarkup([
            [InlineKeyboardButton("📱 bKash", callback_data="usr_dep_method_bKash"), InlineKeyboardButton("📱 Nagad", callback_data="usr_dep_method_Nagad")],
            [InlineKeyboardButton("💎 Binance", callback_data="usr_dep_method_Binance")],
        ])
        await update.message.reply_text(
            f"✅ আপনি <b>{amt:.2f} BDT</b> ডিপোজিট করতে চেয়েছেন।\n\nনিচের কোন মাধ্যমে সেন্ড মানি করবেন সিলেক্ট করুন:",
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
            f"✅ <b>ডিপোজিট রিকোয়েস্ট সাবমিটেড!</b>\n\nআইডি: #{dep_id} | মেথড: {method}\nপরিমাণ: {amount:.2f} BDT | TrxID: <code>{trx}</code>",
            parse_mode=ParseMode.HTML,
        )

        adm_kb = InlineKeyboardMarkup([[InlineKeyboardButton("✅ Approve", callback_data=f"adm_dep_approve_{dep_id}"), InlineKeyboardButton("❌ Reject", callback_data=f"adm_dep_reject_{dep_id}")]])
        try:
            await context.bot.send_message(
                chat_id=SUPER_ADMIN_ID,
                text=f"🔔 <b>নতুন ডিপোজিট রিকোয়েস্ট #{dep_id}</b>\n👤 {user.full_name} (<code>{user.id}</code>)\n📱 {method} | 💰 {amount:.2f} BDT\n🔢 TrxID: <code>{trx}</code>",
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
                f"🔑 <b>2FA Live Code</b>\n\n🔢 কোড: <code>{code_2fa}</code>\n⏳ মেয়াদ বাকি: <b>{time_remaining} সেকেন্ড</b>",
                parse_mode=ParseMode.HTML,
            )
        except Exception:
            await update.message.reply_text("❌ অবৈধ 2FA Secret Key! দয়া করে সঠিক কি দিন।")
        return

    if state in ("WAITING_CALC_CUSTOM_QTY", "WAITING_CUSTOM_QTY"):
        try:
            qty = int(text)
            if qty <= 0:
                raise ValueError
        except ValueError:
            await update.message.reply_text("দয়া করে সঠিক সংখ্যা লিখুন:")
            return
        prod_id = sdata.get("prod_id")
        set_user_state(user.id, None)
        await render_purchase_summary(update.message, prod_id, quantity=qty, edit=False)
        return

    if not is_admin(user.id):
        # Fallback for normal users typing questions directly in chat -> Smart AI Assistant
        ans = generate_smart_ai_reply(text, user_id=user.id)
        await update.message.reply_text(ans, parse_mode=ParseMode.HTML)
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
            await update.message.reply_text(f"✅ সফলভাবে <b>{len(lines)}</b> টি স্টক যুক্ত হয়েছে! মোট স্টক: <b>{total_now}</b> টি।", parse_mode=ParseMode.HTML)
        return

    if state == "ADMIN_WAITING_VPN_EMAIL":
        order_id = sdata.get("order_id")
        set_user_state(user.id, "ADMIN_WAITING_VPN_PASS", {"order_id": order_id, "email": text.strip()})
        await update.message.reply_text(
            f"✅ Email গ্রহণ করা হয়েছে: <code>{text.strip()}</code>\n\n"
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
            await update.message.reply_text(f"✅ VPN Order #{order_id} সফলভাবে কাস্টমারের ইনবক্সে পাঠানো হয়েছে!")
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
            await update.message.reply_text("অর্ডার পাওয়া যায়নি।")
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
            await update.message.reply_text(f"✅ Order #{order_id} সফলভাবে কাস্টমারকে ডেলিভারি পাঠানো হয়েছে।")
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
        await update.message.reply_text(f"✅ Update Group ID সেভ হয়েছে: <code>{text.strip()}</code>", parse_mode=ParseMode.HTML)
        return

    if state == "ADMIN_WAITING_WGRP_LINK":
        set_setting("welcome_group_link", text.strip())
        set_user_state(user.id, None)
        await update.message.reply_text(f"✅ Update Group Link সেভ হয়েছে: {text.strip()}")
        return

    if state == "ADMIN_WAITING_TXCH_ID":
        set_setting("transaction_channel_id", text.strip())
        set_user_state(user.id, None)
        await update.message.reply_text(f"✅ Transaction Channel ID সেভ হয়েছে: <code>{text.strip()}</code>", parse_mode=ParseMode.HTML)
        return

    if state == "ADMIN_WAITING_TXCH_LINK":
        set_setting("transaction_channel_link", text.strip())
        set_user_state(user.id, None)
        await update.message.reply_text(f"✅ Transaction Channel Link সেভ হয়েছে: {text.strip()}")
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
            await update.message.reply_text(f"✅ কাস্টম ট্রানজ্যাকশন (<code>{fake_receipt}</code>) চ্যানেলে পাঠানো হয়েছে!", parse_mode=ParseMode.HTML)
        else:
            await update.message.reply_text(f"❌ পাঠানো যায়নি: <code>{err}</code>", parse_mode=ParseMode.HTML)
        return

    if state == "ADMIN_WAITING_BCAST_MSG":
        set_user_state(user.id, "ADMIN_WAITING_BCAST_BTN_CHOICE", {"msg_text": text, "photo_id": None})
        kb = InlineKeyboardMarkup([
            [InlineKeyboardButton("✅ হ্যাঁ (বাটন যোগ করব)", callback_data="adm_bcast_add_btn")],
            [InlineKeyboardButton("❌ না (বাটন ছাড়াই পাঠাব)", callback_data="adm_bcast_skip_btn")],
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
            f"✅ <b>[{idx}/{cnt}] বাটন গ্রহণ করা হয়েছে!</b>\n"
            f"🔗 এবার এই বাটনের নিচে কোন <b>লিংক (https://...)</b> যাবে তা দিন:",
            parse_mode=ParseMode.HTML,
        )
        return

    if state == "ADMIN_WAITING_BCAST_BTN_URL":
        b_url = text.strip()
        if not (b_url.startswith("http://") or b_url.startswith("https://") or b_url.startswith("tg://")):
            await update.message.reply_text("❌ লিংকটি অবশ্যই <code>https://</code> দিয়ে শুরু হতে হবে। আবার দিন:", parse_mode=ParseMode.HTML)
            return
        btns = sdata.get("buttons", [])
        btns.append({"text": sdata.get("cur_btn_name", "Open"), "url": b_url})
        sdata["buttons"] = btns
        cnt = sdata.get("btn_count", 1)
        if len(btns) < cnt:
            nxt = len(btns) + 1
            set_user_state(user.id, "ADMIN_WAITING_BCAST_BTN_NAME", sdata)
            await update.message.reply_text(
                f"✅ <b>[{len(btns)}/{cnt}]</b> বাটন সম্পন্ন হয়েছে!\n"
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
            await update.message.reply_text("❌ রিসিট পাওয়া যায়নি।")
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
        await update.message.reply_text(f"✅ ব্র্যান্ডের নাম পরিবর্তন হয়েছে: <b>{text}</b>", parse_mode=ParseMode.HTML)
        return

    if state == "ADMIN_WAITING_REN_CAT":
        cat_id = sdata.get("cat_id")
        conn = get_db()
        c = conn.cursor()
        c.execute("UPDATE categories SET name = ? WHERE id = ?", (text, cat_id))
        conn.commit()
        conn.close()
        set_user_state(user.id, None)
        await update.message.reply_text(f"✅ ক্যাটাগরির নাম পরিবর্তন হয়েছে: <b>{text}</b>", parse_mode=ParseMode.HTML)
        return

    if state == "ADMIN_WAITING_NEW_CAT_NAME":
        conn = get_db()
        c = conn.cursor()
        c.execute("INSERT INTO categories (name, icon) VALUES (?, '🛍️')", (text,))
        conn.commit()
        conn.close()
        set_user_state(user.id, None)
        await update.message.reply_text(f"✅ নতুন ক্যাটাগরি <b>{text}</b> তৈরি হয়েছে!", parse_mode=ParseMode.HTML)
        return

    if state == "ADMIN_WAITING_NEW_SUB_NAME":
        cat_id = sdata.get("cat_id")
        conn = get_db()
        c = conn.cursor()
        c.execute("INSERT INTO subcategories (category_id, name) VALUES (?, ?)", (cat_id, text))
        conn.commit()
        conn.close()
        set_user_state(user.id, None)
        await update.message.reply_text(f"✅ নতুন ব্র্যান্ড <b>{text}</b> তৈরি হয়েছে!", parse_mode=ParseMode.HTML)
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
        await update.message.reply_text(f"✅ নতুন ভ্যারিয়েন্ট <b>{p_name}</b> ({price_bdt:.2f} BDT) যুক্ত হয়েছে!", parse_mode=ParseMode.HTML)
        return

    if state == "ADMIN_WAITING_DESC":
        prod_id = sdata.get("prod_id")
        conn = get_db()
        c = conn.cursor()
        c.execute("UPDATE products SET description = ? WHERE id = ?", (text, prod_id))
        conn.commit()
        conn.close()
        set_user_state(user.id, None)
        await update.message.reply_text("✅ কাস্টম বিবরণ সেভ হয়েছে!")
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
            await update.message.reply_text(f"✅ নতুন দাম সেট হয়েছে: {p_bdt:.2f} BDT")
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
        await update.message.reply_text(f"✅ নাম পরিবর্তন হয়েছে: {text}")
        return

    if state == "ADMIN_WAITING_BKASH":
        set_setting("bkash_number", text)
        set_user_state(user.id, None)
        await update.message.reply_text("✅ বিকাশ নাম্বার সেট হয়েছে।")
        return

    if state == "ADMIN_WAITING_NAGAD":
        set_setting("nagad_number", text)
        set_user_state(user.id, None)
        await update.message.reply_text("✅ নগদ নাম্বার সেট হয়েছে।")
        return

    if state == "ADMIN_WAITING_BINANCE":
        set_setting("binance_id", text)
        set_user_state(user.id, None)
        await update.message.reply_text("✅ বাইনান্স তথ্য সেট হয়েছে।")
        return

    if state == "ADMIN_WAITING_MINDEP_MANUAL":
        try:
            new_min = float(text)
            set_setting("min_deposit_manual", str(new_min))
            set_user_state(user.id, None)
            await update.message.reply_text(f"✅ সর্বনিম্ন ডিপোজিট সেট হয়েছে: {new_min:.0f} BDT")
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
        await update.message.reply_text(f"✅ <code>{t_id}</code>-কে {amt:.2f} BDT যোগ করা হয়েছে!", parse_mode=ParseMode.HTML)
        return

    if state == "ADMIN_WAITING_DEDBAL_AMT":
        try:
            amt = float(text)
        except ValueError:
            return
        t_id = sdata.get("target_uid")
        conn = get_db()
        c = conn.cursor()
        if conn.is_pg:
            c.execute("UPDATE users SET balance_bdt = GREATEST(0, balance_bdt - %s) WHERE user_id = %s", (amt, t_id))
        else:
            c.execute("UPDATE users SET balance_bdt = MAX(0, balance_bdt - ?) WHERE user_id = ?", (amt, t_id))
        conn.commit()
        conn.close()
        set_user_state(user.id, None)
        await update.message.reply_text(f"✅ <code>{t_id}</code> থেকে {amt:.2f} BDT কাটা হয়েছে!", parse_mode=ParseMode.HTML)
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
            [InlineKeyboardButton("❌ না (বাটন ছাড়াই পাঠাব)", callback_data="adm_bcast_skip_btn")],
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
        await update.message.reply_text(f"✅ নতুন ফটো সফলভাবে সেভ হয়েছে!\n📋 File ID: <code>{photo_file_id}</code>", parse_mode=ParseMode.HTML)
        return

    # Convenience: Admin sends any photo outside state -> return its Telegram file_id for Web Admin use
    photo_file_id = update.message.photo[-1].file_id
    await update.message.reply_text(
        f"🖼️ <b>Telegram Photo File ID:</b>\n<code>{photo_file_id}</code>\n\n"
        f"💡 <i>এই কোডটি কপি করে ওয়েব এডমিন প্যানেলের যেকোনো প্রোডাক্ট বা অ্যাড ব্যানারে পেস্ট করতে পারবেন।</i>",
        parse_mode=ParseMode.HTML,
    )

async def document_message_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user = update.effective_user
    if not is_admin(user.id):
        return

    state, sdata = get_user_state(user.id)
    if state == "ADMIN_WAITING_STOCK_INPUT":
        doc = update.message.document
        if not doc.file_name.endswith(".xlsx"):
            await update.message.reply_text("❌ দয়া করে <code>.xlsx</code> এক্সেল ফাইল আপলোড করুন।", parse_mode=ParseMode.HTML)
            return

        tg_file = await context.bot.get_file(doc.file_id)
        file_bytes = await tg_file.download_as_bytearray()
        try:
            items = read_items_from_excel(file_bytes)
        except Exception as e:
            await update.message.reply_text(f"❌ এক্সেল ফাইল পড়তে সমস্যা হয়েছে: {e}")
            return

        prod_id = sdata.get("prod_id")
        if not items:
            await update.message.reply_text("❌ এক্সেল ফাইলে কোনো ডেটা পাওয়া যায়নি।")
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
    print("Starting Render Web Service + External webapp.html Server...")
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
