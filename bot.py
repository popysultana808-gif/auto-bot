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

# ------------------------------------------------------------------------------
# TELEGRAM BOT API 9.4 COLORED BUTTON HELPERS (PTB v21+ & v22.7+ SAFE)
# ------------------------------------------------------------------------------
import inspect as _inspect

_IKB_SUPPORTS_STYLE = "style" in _inspect.signature(InlineKeyboardButton.__init__).parameters
_KB_SUPPORTS_STYLE = "style" in _inspect.signature(KeyboardButton.__init__).parameters

_DEFAULT_BTN_COLORS: Dict[str, Any] = {
    "main_menu": "primary",
    "admin_panel": "primary",
    "inline_default": "primary",
    "inline_success": "success",
    "inline_danger": "danger",
    "overrides": {},
}
_BTN_COLOR_CACHE: Optional[Dict[str, Any]] = None

# All Admin Panel Buttons available for individual color customization
ADMIN_PANEL_CUSTOM_BUTTONS: List[str] = [
    "📥 Add Stock",
    "🏷️ Manage Buttons & Prices",
    "🖼️ Manage Product Photos",
    "🖼️ Message Photo Manager",
    "📝 Message Text Manager",
    "💳 Payment Settings",
    "💳 Deposits",
    "🛡️ Orders",
    "👥 User Management",
    "📊 Shop Statistics",
    "🧾 Check / Download Receipt",
    "📢 Broadcast (with Button)",
    "🟢 Auto Channel Sales: ON",
    "🔴 Auto Channel Sales: OFF",
    "⚡ Channel TX Manager",
    "🎟️ Promo / Coupon Manager",
    "📢 Force Join Manager",
    "🎛️ Menu Layout Manager",
    "🎨 Button Color Manager",
    "🔄 Factory Reset / Clear Data",
    "❌ Close Admin Panel",
]

# All Bot Message Slots that support custom attached Photos
BOT_MESSAGE_SLOTS: List[Tuple[str, str]] = [
    ("welcome_new", "🌟 Welcome Message (/start — New User)"),
    ("welcome_back", "👋 Welcome Message (/start — Returning User)"),
    ("force_join", "📢 Force Join Verification Message"),
    ("buy_products", "🛍️ Buy Products Menu Message"),
    ("ai_assistant", "🤖 AI Assistant Message"),
    ("my_profile", "👤 My Profile Message"),
    ("deposit_menu", "💳 Deposit Amount Prompt Message"),
    ("deposit_method", "📱 Deposit Method Selection Message"),
    ("deposit_instr", "🧾 Deposit Payment Instructions Message"),
    ("get_code", "🔐 Get Code & Auth Tools Message"),
    ("support", "💬 Customer Support Message"),
    ("purchase_summary", "🛒 Purchase Summary Card Message"),
    ("order_cancel", "❌ Order Cancelled Message"),
    ("group_welcome", "🌸 Group New Member Welcome Message"),
]
BOT_MESSAGE_SLOT_MAP: Dict[str, str] = dict(BOT_MESSAGE_SLOTS)

BOT_MESSAGE_PLACEHOLDERS: Dict[str, str] = {
    "welcome_new": "{name}, {shop_name}",
    "welcome_back": "{name}, {shop_name}",
    "force_join": "{name}, {shop_name}",
    "buy_products": "{name}, {shop_name}",
    "ai_assistant": "{name}, {shop_name}",
    "my_profile": "{name}, {user_id}, {balance}, {balance_usdt}, {total_spent}, {joined_date}, {vip_badge}, {shop_name}",
    "deposit_menu": "{name}, {balance}, {min_deposit}, {shop_name}",
    "deposit_method": "{amount}, {name}, {shop_name}",
    "deposit_instr": "{method}, {amount}, {number}, {name}, {shop_name}",
    "get_code": "{name}, {shop_name}",
    "support": "{name}, {support}, {shop_name}",
    "purchase_summary": "{category}, {package}, {description}, {quantity}, {rate}, {total}",
    "order_cancel": "{name}, {shop_name}",
    "group_welcome": "{user_link}, {name}, {shop_name}",
}

def get_btn_color_config() -> Dict[str, Any]:
    global _BTN_COLOR_CACHE
    if _BTN_COLOR_CACHE is not None:
        return _BTN_COLOR_CACHE
    cfg = {
        "main_menu": "primary",
        "admin_panel": "primary",
        "inline_default": "primary",
        "inline_success": "success",
        "inline_danger": "danger",
        "overrides": {},
    }
    try:
        raw = get_setting("btn_color_rules", "")
        if raw:
            parsed = json.loads(raw)
            if isinstance(parsed, dict):
                for k in ("main_menu", "admin_panel", "inline_default", "inline_success", "inline_danger"):
                    if parsed.get(k) in ("primary", "success", "danger", "none"):
                        cfg[k] = parsed[k]
                if isinstance(parsed.get("overrides"), dict):
                    cfg["overrides"] = parsed["overrides"]
        _BTN_COLOR_CACHE = cfg
    except Exception:
        pass
    return cfg

def save_btn_color_config(cfg: Dict[str, Any]):
    global _BTN_COLOR_CACHE
    _BTN_COLOR_CACHE = cfg
    set_setting("btn_color_rules", json.dumps(cfg, ensure_ascii=False))

def _infer_btn_style(text: str, explicit_style: Optional[str] = None, group: Optional[str] = None) -> Optional[str]:
    cfg = get_btn_color_config()
    overrides = cfg.get("overrides") or {}
    raw_t = (text or "").strip()
    if raw_t in overrides:
        ov = overrides[raw_t]
        return None if ov == "none" else ov
    t = raw_t.lower()
    for ov_k, ov_v in overrides.items():
        if ov_k and ov_k.lower() in t:
            return None if ov_v == "none" else ov_v

    if explicit_style == "none":
        return None
    if explicit_style in ("primary", "success", "danger"):
        return explicit_style

    danger_kw = (
        "cancel", "reject", "delete", "refund", "close", "admin panel",
        "wipe", "reset", "remove", "hide", "off", "বাতিল", "মুছে", "বন্ধ", "কাটবেন", "🗑️", "❌", "🔴", "👑", "⏹️"
    )
    success_kw = (
        "confirm", "approve", "verify", "deliver", "buy", "deposit", "refresh",
        "add", "save", "enable", "হ্যাঁ", "ভেরিফাই", "কিনুন", "ডিপোজিট", "✅", "🟢", "💳", "🛍️", "🔄", "➕", "🚀"
    )

    if group == "main_menu":
        res = cfg.get("main_menu", "primary")
        return None if res == "none" else res
    if group == "admin_panel":
        if any(k in t for k in ("close admin", "factory reset", "❌")):
            res = cfg.get("inline_danger", "danger")
        else:
            res = cfg.get("admin_panel", "primary")
        return None if res == "none" else res

    if any(k in t for k in danger_kw):
        res = cfg.get("inline_danger", "danger")
        return None if res == "none" else res
    if any(k in t for k in success_kw):
        res = cfg.get("inline_success", "success")
        return None if res == "none" else res
    res = cfg.get("inline_default", "primary")
    return None if res == "none" else res

def make_inline_btn(text: str, style: Optional[str] = None, group: Optional[str] = None, **kwargs) -> InlineKeyboardButton:
    st = _infer_btn_style(text, style, group=group)
    if not st:
        return InlineKeyboardButton(text=text, **kwargs)
    if _IKB_SUPPORTS_STYLE:
        return InlineKeyboardButton(text=text, style=st, **kwargs)
    ak = dict(kwargs.pop("api_kwargs", {}) or {})
    ak["style"] = st
    return InlineKeyboardButton(text=text, api_kwargs=ak, **kwargs)

def make_reply_btn(text: str, style: Optional[str] = None, group: Optional[str] = None, **kwargs) -> KeyboardButton:
    st = _infer_btn_style(text, style, group=group)
    if not st:
        return KeyboardButton(text=text, **kwargs)
    if _KB_SUPPORTS_STYLE:
        return KeyboardButton(text=text, style=st, **kwargs)
    ak = dict(kwargs.pop("api_kwargs", {}) or {})
    ak["style"] = st
    return KeyboardButton(text=text, api_kwargs=ak, **kwargs)

def colorize_markup(markup, group: Optional[str] = None):
    """Automatically applies Bot API 9.4 colored styles to InlineKeyboardMarkup or ReplyKeyboardMarkup."""
    if markup is None:
        return None
    if isinstance(markup, InlineKeyboardMarkup):
        new_rows = []
        for row in markup.inline_keyboard:
            new_row = []
            for b in row:
                st = _infer_btn_style(b.text, group=group)
                kw = {}
                if b.url:
                    kw["url"] = b.url
                if b.callback_data:
                    kw["callback_data"] = b.callback_data
                if b.web_app:
                    kw["web_app"] = b.web_app
                if b.switch_inline_query is not None:
                    kw["switch_inline_query"] = b.switch_inline_query
                if b.switch_inline_query_current_chat is not None:
                    kw["switch_inline_query_current_chat"] = b.switch_inline_query_current_chat
                new_row.append(make_inline_btn(b.text, style=(st or "none"), group=group, **kw))
            new_rows.append(new_row)
        return InlineKeyboardMarkup(new_rows)
    if isinstance(markup, ReplyKeyboardMarkup):
        new_rows = []
        for row in markup.keyboard:
            new_row = []
            for b in row:
                st = _infer_btn_style(b.text, group=group)
                kw = {}
                if getattr(b, "web_app", None):
                    kw["web_app"] = b.web_app
                new_row.append(make_reply_btn(b.text, style=(st or "none"), group=group, **kw))
            new_rows.append(new_row)
        return ReplyKeyboardMarkup(
            new_rows,
            resize_keyboard=markup.resize_keyboard,
            one_time_keyboard=markup.one_time_keyboard,
            selective=markup.selective,
            input_field_placeholder=markup.input_field_placeholder,
        )
    return markup

def render_custom_bot_message(slot_key: str, default_text: str, **kwargs) -> str:
    """Returns admin-customized text for a bot message slot if configured, with safe placeholder substitution."""
    custom = get_setting(f"msg_text_{slot_key}", "").strip()
    if not custom:
        return default_text
    out = custom
    for k, v in kwargs.items():
        out = out.replace(f"{{{k}}}", str(v))
    return out

async def send_smart_bot_message(target, slot_key: str, text: str, reply_markup=None, parse_mode=ParseMode.HTML):
    """Sends a bot message with its admin-configured Photo if attached, else sends clean text."""
    photo_id = get_setting(f"msg_photo_{slot_key}", "").strip()
    if photo_id:
        photo_src = photo_id
        if isinstance(photo_src, str) and photo_src.startswith("data:image") and "," in photo_src:
            try:
                b64_part = photo_src.split(",", 1)[1]
                photo_src = InputFile(io.BytesIO(base64.b64decode(b64_part)), filename=f"{slot_key}.jpg")
            except Exception:
                photo_src = None
        if photo_src:
            try:
                return await target.reply_photo(
                    photo=photo_src,
                    caption=text[:1024],
                    parse_mode=parse_mode,
                    reply_markup=reply_markup,
                )
            except Exception as e:
                logger.warning(f"send_smart_bot_message photo fallback for {slot_key}: {e}")
    return await target.reply_text(
        text=text,
        parse_mode=parse_mode,
        reply_markup=reply_markup,
    )


_TOOL_CACHE: Dict[str, str] = {}
_LIVE_TOOL_STOP: set = set()

CYBERNEST_LOGO_BASE64 = "/9j/4AAQSkZJRgABAQAAAQABAAD/2wBDAAUDBAQEAwUEBAQFBQUGBwwIBwcHBw8LCwkMEQ8SEhEPERETFhwXExQaFRERGCEYGh0dHx8fExciJCIeJBweHx7/2wBDAQUFBQcGBw4ICA4eFBEUHh4eHh4eHh4eHh4eHh4eHh4eHh4eHh4eHh4eHh4eHh4eHh4eHh4eHh4eHh4eHh4eHh7/wAARCACMAIwDASIAAhEBAxEB/8QAHwAAAQUBAQEBAQEAAAAAAAAAAAECAwQFBgcICQoL/8QAtRAAAgEDAwIEAwUFBAQAAAF9AQIDAAQRBRIhMUEGE1FhByJxFDKBkaEII0KxwRVS0fAkM2JyggkKFhcYGRolJicoKSo0NTY3ODk6Q0RFRkdISUpTVFVWV1hZWmNkZWZnaGlqc3R1dnd4eXqDhIWGh4iJipKTlJWWl5iZmqKjpKWmp6ipqrKztLW2t7i5usLDxMXGx8jKytLT1NXW19jZ2uHi4+Tl5ufo6erx8vP09fb3+Pn6/8QAHwEAAwEBAQEBAQEBAQAAAAAAAAECAwQFBgcICQoL/8QAtREAAgECBAQDBAcFBAQAAQJ3AAECAxEEBSExBhJBUQdhcRMiMoEIFEKRobHBCSMzUvAVYnLRChYkNOEl8RcYGRomJygpKjU2Nzg5OkNERUZHSElKU1RVVldYWVpjZGVmZ2hpanN0dXZ3eHl6goOEhYaHiImKkpOUlZaXmJmaoqOkpaanqKmqsrO0tba3uLm6wsPExcbHyMnK0tPU1dbX2Nna4uPk5ebn6Onq8vP09fb3+Pn6/9oADAMBAAIRAxEAPwD5DAzSgUop6ivRORsFWpFWlValCnuKpGbGhaeFzUsURJB25Gfzra0XRNQ1SdbSytJZ5GO4RxpuP6c4ptpbkmIsRPQVKts5HCmvdvA/7PfirWo1nuY0soDyWb5iPxyFH/fVegWXwQ+Hmjxb/EHi+1LJ99VuV+X6hFOPzriq4+jSdpSSNadGdT4It+i/pHyYLKQj7ppGtJB/Ca+2bP4T/CJrdJIzqVyjgFZI7a5ZWB7ggYIrK1H4X/Bma5NpD4j+w3Y48ueSSJge3Dqaw/tah1a/r5G7wNf/AJ9y/r5nxs8DL1FRMmOor6p139m5byJ5/C+vW2ooBnCOkv8A6AQf/Ha8Z8a/DLxN4blkF/psmyP70kYLKPr3X8QK7KOKp1vgdzlnHkdpaPz0POStMZa07i0Kx9EUr1+Y5bPt7VTaI966rCTKjLUbLjmrTJ1ycY9ahYc1LRSZARSGpGGDTCKg0Q9BzU8a1HGKuW0TMRhSfwoFa4sUJPar8dk5IyCW4wCO3vXQ+DvDtxrFwLSGzlkmkKiNgDheecjHOf0r6H8KfBG28Og+IPGjLbWVtGshQgMenAA/icnoO3U9hXFiMwo4e/O7W/r/AIY6Y4SpON0tzzH4T/B3U/Ew/tDUMWGlxDfLNK2wKvXJY8KPzJ7DHNfQfh/S/DXhLR4joVrpdlauwjXWdaBiglkJwPJhJ3zHJHzN+Brj/FWuDxl8LtY1CyF1pOm6fqdvZ2NlC20FWZczOeruc4yeBz1PNJdeL38R+HdO11LXR4/F2mL9gub7UIzItsE5MyJnbvIKkE46nniuGVDH4+Kmk4wu4tLdaJ+8+zT6aLsz1MBl9F1lSl707X8kvLvbv+h1Hh+412b4pap4b8e6rZ6232Q6enzrbRyoAwzGCBgq3I56Vzmu66NctToLeJrHUlE6+fpei+H2ishtOW3zEZ4x7c1zy+IvFfjK+sE8IW8/iO70ZZHu9UkVYYnYrtKmThRnJJx6Co5JNVeYQeIfiNpOktn5rTRbdr2ZfUZHGa9ChhMHhK0G2udJXjFc2q3ezav/TPQnQgqNWjVqO+qTSWia6PbS9tOx3uveJ9Lh8P+C9GbxzrnhWR7c28X9nKPLlPy43lvugZGPrVjVL3TPEvxW0qPU9OZo9HS3RNQusF7vZl2lyvBGQBz3JrlbPwn4M1SK4+3S/EXWICCISdMiVYlPXaD07flWdPpuhWI8uD4ieLdHud21W1jRt8ax4wVynI+vtXbHDUKUnU5Z6qX2Xpe+vw39L7HnOthq6VJu1nF6NXdrOXpr66Hb3lnovxG8SaF4w8OXs+nK1+1jqUNo/l8xSN8xxgjco6j1FXPDet+J/EHinX7CG20/W/Cuns620mp3GJvlbZtE2MqWIbBORivOtJ0PxP4UkuLnwFNp3i/T7xVZ10m7zJDIowHaJvmGRjI56U/wn4siXwg3hrSrKx1N9+NY03UC0EsvHzAd1YNggnjjtXnyyvD1qclh5KUo2StpJa3u9U20tEnY7KOGliVHl63coy1tbSy9d9O3mavjz4OaH4oluJPDsE+k65EN82l3KBJh7qB8sq/7S8/U8V4H4l8B6zpN3NFdwmNiSN8h2iTnsTx1/LvXt2l+J/EmseJtP8AB3hPVdRsoElR7dtaRTc2cwBJgWX7zRbec854xmvU9dmv/E+hR6RrFxpEupb3jstQtQktreSqMNGwYHy5uOh64rpo145eoQx0782vml3l3/PfVrU8PF5fec/quvLuul+y9O3/AAT4Mu7F4ZHjeJkK8Nx0Pv6VnzeZ5axFjsQkqvoT1P6V734lstP/ALWfSPEEGnRXgJXMYfTpQc8gE7oWIPZgOe4615/4/wDBF9ocRv4JZbrTy4jZ3j2SwMRlUlXJAyASrAlGAO0nBA9mthLLmi7o4FCaSlJaff8A0/I86ZcVERzVqZSDzVZhzXmyVmVsSx9BV6xdRKvmFtg5O04PTscVShKhhv3Fe+3rWxpEE739ulssMkqyAKFAfceOTngjJx+dCjcly5T2X9m3w9dap4njvJpTFbWf7x5CfkXbyWPbCjB9yVHevXvjb418aWN9pWs6adPgthIY9Ps70ebNKNpPnsn8O4DgnkZ+tc7oN9pfw98O+E9Euyqt4hvI5b5zxstFOefQO3zH/Z47Vh+KtDvJvixqfirWtRingE5ktpDJkRxAcLj+ED8sdK8bBYP+1MycpJeyhffaTWkn9+ifZX63X0mCoVvq96abnKy2vy32b7K2rv6Fa5+LMmsWl7ovjLRrTTFlVXWXT42HnSIcqMD+PONvY9Krah4b0PTo7bUPFcd5cajfqstt4bt5NrTf3XuSPujnp745PA6XwxpX2m41DxXNZ2k+oyt5miW91KkcagqAJnGeOmVXHv1IxyHxA0Hxr4U02fxLewSXT3sm2XU4ZVl2uexZT8g7DpgcCvolgakpPC0r0qGjb+1J22j1t3lv0VrHj4nOMNTkrzjUrR5orl0Vr7ytp8vP7tLXdcsbKySLxlqflQRAeV4a0VhDBEOwlboD7Hc3sKxrb4t6lbkW/g3w5pmjQjhXhtleT8ZJAT+QFYvwy8BQeMIbnUdR1W4R4rmOCOGK2815HcFsKNwHb3617v4D+HnhzQZ0uG04yzryk2pOrlT6rEnAP+9+dd9ChSw8OShBKP5/5/iz5nMMx5m3iZtvtsvkjzl/ip8VrEbb3Vr2FnQMqyMV4PIIGBwfWs0/GP4jxuftN6L+HvHcIkykfR1NeyeONW8O3+o3Oj+LrU7YWxbzXMRiYrgfNHIMcfp9axNF+GngO+WeX+2ryIeU7xxuiSA4Xdwynnj2zXW+blTlFJf4U/yR5sMXS35Xf5r9Tzmy8feE9WvEk1bRZPDmrZ+TUtGZoXU+piJwR/usPpWl4hNnfrBc+KbhdQtWOLTxZpQ23EJ9JlABbHcEA+x607wB8M/D/jXVtQe71C5/su2GwSRxBWaVvuhck5wOfyrlPGuly/DPxXNoUWtvqFlcQo52wkB0bO0MCeSMdRnH6Vx4rL6GJlaorTW0luvn+m3kezl+dVKFRRpSbX8r/Tqjp7dNS0e5bSJ9ZikvZrhNT0PW0UFLgqBgZHPYgpnoxxnFa3xD8UT6HbSSabb2UcvieANqtuZ2WK0u4yuLuIgZUknGcdRn1rk/Dj6YdMi03xErS+FtUlKJKuN2nT5H7xf7nJBK++RxmnWXhDRfDet6vpvirWFFraqpuJHc5nhbOx0JzwRxgZOcjFeUsPCpX9hjvjirpracfTuna66PbRn6A8bUrYWCw1O8n1btbq3K/ls/z0v6J8TvC134v8Js2p/Z/wDhK9Ls47iee35jvrdlG24Q98fdb6ey1w/wVK67p+q+EPEDMLeOIQoz8mJJJBGVP+ysjRyD0Kkjqc9Vo3xf8Gz2Ki0tm0JfDRjXS1uCS1/bP8ssJHJ+bqAc4IBNReJNIsfCmp6hrtg6ix1iS1WzI6eSGFw7H6LGF/KufhypXpVJYKvFxtdxv20uuu101q7JtfZOTE4aneVaDvBp3+Sdn87fefNHirS7rRtavNLvY/LubSZ4ZV9GU4P6isNhz1r0L44yeb8Qb+Vx+9eO3ef/AK6GCMvn33Zrz5uvSu3ExtM+ehLmimyWBAzqGPGMnDDP057+1eg/BPRX1vxtYWLRqAkuGOORuODn6DcfwrzuIjdyMivd/wBl2BItUvtVKj/RbaWTOO4UD/2c1w4qu8NhataO8Ytr1tp+Nh06fta0Kb6tL5dfwNjxYG8b/F/UbW2AMEbppFmv8KAYDH6AfoTXNWHhm48TePo/Dlpqt8dANyVETTsVW3j5ZsE/3V+mWrd+EUjJb654jY/vbbTbu8Df9NZW8pD/AOPj8qufD+KTS/A/jPxKoImisEsLdvR52+Yj8AK9mphllmSwo0l70YxS/wActL/e0zLL61XF5tJ8zUW3f/Ctl+BxuteMtW0yLVbvQ7mSHRra+MFkrlWyh3YUZBPAXPsMV6D8NvFt34x+Ht7Z3cgf7VFJbzxkDa/GVyOmcjqK8h+KluumaDoGiRDbutvtco9WlORn6Isf61d+BGsnTp721YnYrLIAO3PX8wfzr3cvqzlWhg6r5vdtr3Wt/wADg4oy7DQg6mHgoyjZu2m+r+69vkdr+z1eR2l9LZStgR6xHnPUAwyAfj8pr2XwK99qs88+nSQ7xG/kM7AbZCGCMM8ZBGfTivMD4P8A+EbsLvx9pF5NcxyayI5bPygBDImJoCrA8iRGZeQOWAq94c8VxeGdXnD6lbxpcSRXdq9xKI0kt23kbTjGRkAjrnPpUUqKlGcIvVNrX1PncxVRtVaau2k0df4z8JeMf+EW8Sx61q8OqwC5S70gyXvny22wN5yg84yByucZHrXnPw+uLyLTZriWV3me9MWShXOLeQk+/GK108V+H9E0G80fTNc0qa2lnvbpxbzAPvn3FUCZYnGVXPtmucm1K50PRbW2T5tQjBk2qMn7RcYCJwBlliVmPoWFOlTlGPvNXfb+vn8xxdSbceXS+l92kt38ztvDE0Hhbw1a6TC8lyFcyXcki5jWVlXdH8oyNoBOeSPTmmvrkc8lmZ9GsLy9gfFjNPbiYJnnCOD1GQcHOPY1gt4hjTwzLp2kgy393ujuIxGUeGNTvk3Kehd1yR2SJB3NUrDU08LeCdT8VTXkiSSwPa2io/GMrjAKk9GXrntxT5YpOTRz+xbf95v8y58RNe8PS+L4/BwtLWC5vrcDUpohtRLk/wCrOOmRnDH0Y1j32mL4q+Hrm/iJ1rwhN5Exfl5LRjgZ9ShH/jp9a8Q+3Xc2qLqdzPJNctLvlkZss5PUk19NfDyBr34gWszws9j4n0R47wKuQXKlWP8A32jH/gVeFjk8RSVemvfpSUl5r7S+cbn6dkdFfU/qsnfS1/XX8Ht6nOeHvCVh4m8IxSQWsEV5YXD215IflBgkUssrH/ZKsCf92vQbDT9I8QfDNrwrMtjocz2szTNiSW3jxvYA/wATsoQAdPMI7VR8DaZazyav4Qs4ZJIXs5UuL3aFeW5T5ljQNjA+RgAeuSx4FSeH9Ys7yx8XeELNY2WLTVuXmjOVd0+6kf8A0zTGAerMWY9Rj0M4xEcPRVZu3JJf+Ay0d/v28jzaDVJTwFN3lJO7846pL06vqfLfjLUbjV9dv9XvQVnu7hpmXbjG/wCYY9sHA9sVzr43HHI7ZrY8RRGPVbm2WMfupXC7RyRk4z9BWKeteHiL87TJpO8E0SwYzzzX0N+znCv/AAiXiV1kQN/Z83f2P+Ar55tVZ32ou4+mO3c19A/suq19PqOjQgu13byRIDxuyBj6ferix1N1MFVilfRfhKL/ACOjDPlxVNt9X+TN74TaQLr4YeJVjngEj2dnGxL/AHV8xSSfbgV6FD4Ss7b4C3kbXlvEl1qUMhkkfG5QoA+meSB6VyvhIW/h7wH4h06K3ju4raxV3dm2G4mhdGPHUoARxx8pBPLEDO0HxJea98HPFMNxcGS4t76C9bnkrjBwOwG08dq+zx1pQTeylB/hFfmceRYKoqrk3aUk7Lyu3f8Ay/q/A/tH6VDZ+No4kureRI7SFE2v0AiQD+VcH8O547bxhDDLIFiuo3hJzwDjK/qK6T45zyXeoaRrGSUu9PgJP+0Iwjf+PRtXm0Vw1teQ3KnDRSK4/A1xVK3sMVCp1jb9LnRmFCU5zhJ73/HY+y/2eb/TfEVrr3g/WMy2Wo2K4AbDBoH271PZgrxMD2KCvPPiH4aOheKZfBviCYXzSKbqyuogBM0bE/vIlP3gSDvhzkNkrkGud+GHiLVNF8VW13okMVzfNMEt4JDhJfO/d7SdwwDuXnI6CvTPiv8ADz4n/E19Mk17wZLp8+m+YsUthPBuZXIJVt8x6EZGPU16WZU5YfHSqRceWavZtL82eFRpRlh1CSd49Un69DzKy0rT/DV/Pqq3iO7Oqx3EsbxW8XAxlWw00ncRqMA/eOKW9ttbuvE1hcy6ZdahpVjJ5kCm/RZLnLbmdXPV2PPAOOABxXR3n7OPjfVZUk1SHxNdyIgRWnurWQqo7DMhwKe/7NXi7bAptfErrAcxA3Np8nfj5+K43iI6q8f/AAKP+Yex6u7e3ws5f+29R1nUbtkuPO1S/l8pvPQrcRksCQwA2lSeCewUA9a9j8HeA9P1XwLdeONavZbTTNKt5V0qFbeGRLgICHnZZUYHe+VU4zge4rCu/g542Or32qR+F9V+1XkAhZvtNr8uQFdgd4ILAEE9txqz8XvFXxA0bwXb+HfE3hax8P6POFhtxbuCfLgAbYMSthQAuTjn8acuarKNKlUjq9fejt6X7Co0IRvOUHotFZ7nj2u+GYNX8W31/iy0bSUaNbq4WPZAJtgLrEi/eYn+Be/oK+lPhto6JB4IbT70Q6cLhuJHXzZwrZAOPdidudoPcmvkRNauri4hn1eZroA7Ybd3O1ATkAgdBz0HJ719D+C9YupvGHgLSxKzJBCbuVQAoVSzsoAHAXCoQPesq1Sk6NWcFZa/dZn1WWqaSo0OiSbe71X/AATvbhLK0+Ma6NYXFqqSz3KzAHBw6sWRR2yT8zdWwBwABXlnwltIoviJrkAuYWX+xpVYhuCA5H9ag8K+I21D4xajrjPmO2j1C9LdgojkIP6rWT8HLplv/E+sOeINJ2Z/2mLN/Ja5eJOWGTVYbtxgvnzM8PLKM45te+icn+B4r42jx4hugvzZftzn5Qa56tjxRN5ut3TBukhGfpx/SsU1wYtp1p+rOrDJ+zivJElqN8yLvVAWHzN0X3PtX0b+zKzaebnX5Tvvb25+y2g5GenmkY6ZLxJ6AM3pXzpp0vlXCSbEfH8LrkH8K+hvh/ONP8GeHdRicottp2q3RZeokjYqD+HmofwFdWXUlUbi9mrHTSpe1xFOHmvuWr/It+I9cjs/jKmjXJFvpRD2D85Vo7pDumJ7ku+7PsPSqXwUtLiHxPr3hu9GN9rNbXMR6lkbH6c/nWZ8TIU1zwppniu1AD2iixv8dQCSYn+gbcmfTb611/gzwV4l8UzW3jiyvBompGz5hlG2W8uY1I3Ivo6LuOe+eOc10PMqLwvNipKFvdd+jX9X/E9OFOrTzX2tNXW/lZ7fK2h574u0m4u/DN3oky5vdDnbZnq0LkEEewbB/wC2leQyq3IYEEcEV9U6xf8AhzXrRPE+p6fe6fqVsDa6nDYqrKxAw4aNscEHPB6Hpxx5t4g8PfDtA2oW954i1ASL5jQW9vCjJz/ESxIPrx7966K9D6xBST95bnoZtly5lNO3r22X/wAi/Na7nKeFbtzFY3CyeWyjyzJydjKeG454IB/Cvriw+NWl+REtz/aC3OxTMIDFLHvxztbzORnpwPcCvkO18u61lNJ8M6R5SSks4v7sOFAHLswChAPYegrq9W8H3FtpTXmmax4a1aaOMyS2trcTCVQOuA2M4wenPtXqqnhsxoQVa94aXS0/I+Dxs/qldxjNe9/XkfU1t8XbCYZjGqt/2xh/+OU+f4uWMIzKuqL/ANsYf/jlfCyeJ0XhtJQkel04qSDxDNeXMVra6SpllcIi/aZGyTXmLCZO5WU5fc/8in9eW9vw/wAz7Wn+NehopMj6tgdcRR/0kr54/aa+IkPi2/t/sUsv2aO28m3SQjeSzZkcgE44AHU9B9BnnwdLFYLJL4g8LRap1Fi1xMzK3ozgFAfxI968murO/wD7ZuLS5iKXiysJlb+A55yfSqxmEoYGzoRk5SurvZX001dgwld4ltSmml2/rY2vAOlyax4htoC22FGy7nog6lj7AAt+Fe2eG9W+yW3iXxwqGONYP7P0pD1GQEQD3ChfyNefeCrFntv+Ef0qOQX92xS5mZcCOEYLY+pxn2AXua9Y8I22iah4misbi9t7bwz4Pia81FnkH7yRRycdTjp9SfauDFOnhsK1U2Wsv0Xm3sfoOU4aNPDe2vvs/Ozt/4Cm2/Npbo878M3T6V4I8Taw2VlvgmkW2euXYPL+SIAf9+trw240P4PajqUp2PrF0Qh7mJBtz9CBJ+Va3xm0PU/F2swav4Z1LTNV0a+uS8JtMR+U0hVdzL/ALKhQT1+XJFcj8btWs7W3s/C+mzgWFjAtvG2M5AX72B3PX/trXnYzFrMI4aEVo2qkvLl6O6W8rLY+bjgKuA9rUqpq6tHz5tLr5XZ5DfSiWZ5cMd3LZ/vHk/rmqhPNT3mPMB3xsdik7OgOOnQc+v8zVU8msJyu7nNBWQ+FsMK94+Cmpadf+DrvSNQl2fYZZZXPf7JcxCGcj/dYRP/AD4yR4Ep5rotC8Qy6JqNlf6MXtbq3GWkYht55BBBGCpBIKnIIJB6114HEKlPU1jKVOSnDdO6PVtBl1fwjqt14d1OMkOB5eeY7hMgo6noVbC4PqF75roobK91Hx7oHia38VG0tNLmjmeGZyPI2nJK44+YcHd6nqKwvD/jnQtVsUs7g2VgoJP9n6jG8tkrHqbeZcy2+f7p3KPXtXVyeEtF8U2kU/8AYE90sMgcSWd3Hext6gmM7iueSp/qcejXw9OrGTjb3t/utfz03R9LgcywsqTw8nZNppNtfK6TTS6fcVNYk1W51nxT8U9MS2g8IyXYiSO4kEX25c4LxA+hztOOfcZrmNV0K0uo08SeGz9v06Q/v7VXKuueoHXa3tz6jcK6+5LGDVdd8d6LLeWfh1VTT/D8wMKj5R+9ZcAOTkAEZUAHAo8T32iQ+HJNatNETwJ4pgEW3SXkDwapE/3RGMYc54K4yp9K87C4qrg3GjXi7aKMlqraJX1u/N2sr9NzumqFenKjV96m7tNbp7XXa9tnv1s9TzG38NQk3l/4bna7gljAntgAl1AA2SGj7j/aXK8DO3pXoXxb+IGn+JPD+gQabpNtph0ePLTRRlAg2BfLX5VwuRnbySce5OZ4p0uAaobfxJp9z4W12EK7TpuaIFhlTvXlc/7X51c0rS/ElyI3uLbw94utU4jkuEVpAPaVCG/Nq+koVYNQqJX5buLT01/P8D5TG8N4uN/q7VWL7aSt5xevzPAktTMzSYxkk4x61Y0kPY6xa3K7gY5M7gpO3tn8K+qdD0/RlQi++FUQfH/LDUzt/wDH2Y/rVDWtLModdG+Gei257SXlw9xj3279v5qa544WMJKabunfp/8AJfoeesDmU5On9Xl9z/ysc/pnjfT4/gjfeC5tCge9lkZvtIj3FmLhxJuC/fGMBtwAGPcHitJtYb6cnSYf7R1q4VFnumXEFoAAo293fA+8eB/CCeRpeKNL2yp/wmXiext4Yh+7sLJF+UeixRgBfxA+tTX39vL8MLjxL4Q0oWPh6CcW0txvzcSknBPH3FzgHHPzda0r4mlSTu1HmlfV7yemnqetg8ghTXPmErcq+GLTlZdHbRLvfXyNLw7bWFlrVp4I0XVLWLWdRk8u41CWQL5Pqqk/8tOuB/D7t07/AF3wf4TWw8Q+Ek8Dx2MWnWghg1+4Pl3F5dsMgp3MY45Oc+lZXijwh8NtX8MaLFbaFdabMdIjvV1m0lARXPLRyhjyQeQfTuMVymv313d+Hk0bxrc2+vW1oUlsr5Lgq0ijoZAOQwHGRkMGHU4r46vXq5lKM6blCEXtqm2no7p63s1Z6K92j62VNyjerBRpxTUVsrdnu1621e/YwvDWiTeBLHUL/UdRBumwkVpb3BMRGMlm6fMTwOOAGPavLtd1STU9TnuJJMiTOSfrnOPc9voO1a/jjxE1/MbaD5IFJG0DAHtjt0Ax2AA9c8jJK5RUJyqkkDHTPX+VdtRez5r/ABS38u0V6dfM+LxleFaShS/hx2v+f+XkNkYseSSfc1GaUmmE1ytnOkIDT1OKjFOBpJlNFmKd1IwTWvput3Wnyi4hnkjnDZAX5eOucg8HNYKtinxvtYNhW9mHFdNLETpu8WZyimrM9i8O/Em/uwiahdPeqE8tkuH3EIc8AnIA5PH3Tk9M4rq9N1NbzWdIvLmOGLw3oQa9SNpVYtOv3QVPKqg+bB6n1HT53jmkEiushRs53ZxgnvxXQaN4puLOZGLOuw/K6Hkfhx+mPxru9tTxMeWpp+X37r5aHs4LOalCKpVVzRX397Puj3T4KfEvU9c8ceI77Xmgm8PXRkvr2C5to5BkLtiVXI3LhF6cj5T61X8P3vhv4keLbTU7TQ4fBuiaVBLd65cWl3nzVAyiocAqSAePY1xNj4vs9QiuLeSC0VL2Pyrs2aJbzSjBGfu43YJGQvfvWno2sW2iaSul+HZV0/ErSNLfwGUyg9mCY5xgZ9BwOTWEco9lzVaF4NxS927iraaJaN22bWll2O3DYnD1VGMqt93Lu3f3Uk7W/pHaava6np3xl0jw9p/iTWH8K6t9nkhu1uGYpFKuOWz1DD9a4r4hWXjJfFWsaZBqN9NptreSQwz3V6qK6KcAkswFal5q8eqHRLu+1GzTUdLvFn8yFJFiaMfeVQRkEgLweMjrzTdaufBepG7kvLGB727Uh74xs0isRjeATjI4OPavRwlPEwcXVm5JRs1Z3b77b20a7l43DqpTkqdZLVNX7cuq+/5Gj8Ih4NvfAniLwv4usLHUbq3cXKXdmwedIiD91x1ZWBPGcg456Vr6frmk+H/B2l/DuyvU1HSdVivC12MDerYZWz0HBxnpuFec2jeHPC0zXmhXerXNw0Zjea8WKGFlJB+4N3cAjLVzmoeLLW1tns9Nihghf78VqgVW/wB5u/47hXBVyuM6rq1L8rkpWk9E0rXS1fy0McLmmHw9CKqJSmlZ9dLrTTTVaa7bnX6pqFnF4W03QtTlkuRYKQ6Rz/ubkKSE345ZQMErkc8EivO/EniR7tzDakJH22ABR6BQOPy4Hb1ORqGo3N9HI7zIqKR+738t6cdTj9KzS3Az19fWnUrQpu9Ld9f8u35s8vG5jVxjs9I6K3VpKyu/0HyOCi4JJ5zxURNBNNJrzpSucKQE0wmlppxmoLQClpKXNK42KDTgabRVCsSA0A0wU6mmTYfvO0LgcHPTmrVvqN7AMQ3U8Y9FkOPyqlRWsK04O8XYlwUtGbCeINVUf8fsn4qp/pUcmuam/W+mH+623+VZeaM10PHYhq3O/vZHsKf8qJ5riWZt0sjO3qzEn9aYkro25XIOCMj0IwajormlUk9WzTlSFzSE0lNNZ3KsKTSHmikNTcYZpDzQetFSykf/2Q=="

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

def _fetch_2fa_live_fallback(cleaned_secret: str) -> Optional[str]:
    """Fallback to 2fa.live public TOTP API if local Base32 fails on unusual keys."""
    try:
        url = f"https://2fa.live/tok/{urllib.parse.quote(cleaned_secret)}"
        req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0 CyberNestBot/3.0"})
        with urllib.request.urlopen(req, timeout=8) as resp:
            data = json.loads(resp.read().decode("utf-8", errors="ignore"))
            tok = str(data.get("token", "")).strip()
            if tok and tok.isdigit():
                return tok
    except Exception:
        pass
    return None

def generate_universal_2fa_codes(raw_input: str) -> List[Dict[str, Any]]:
    """Universal 2FA TOTP parser supporting Facebook, Instagram, Twitter/X, Google, otpauth:// URIs & multi-line keys."""
    import re as _re
    results = []
    rem = 30 - (int(time.time()) % 30)
    for line in (raw_input or "").splitlines():
        ln = line.strip()
        if not ln:
            continue
        label = ""
        secret_candidate = ln

        if "secret=" in ln.lower():
            m_sec = _re.search(r"secret=([A-Za-z0-9\s\-_]+)", ln, flags=_re.IGNORECASE)
            if m_sec:
                secret_candidate = m_sec.group(1)
        elif "|" in ln:
            parts = [p.strip() for p in ln.split("|") if p.strip()]
            label = parts[0] if len(parts) >= 2 else ""
            secret_candidate = parts[-1]
        elif ":" in ln and not ln.lower().startswith("http"):
            parts = [p.strip() for p in ln.split(":") if p.strip()]
            label = parts[0] if len(parts) >= 2 else ""
            secret_candidate = parts[-1]

        # Clean spaces/hyphens; keep A-Z and 0-9 first to test valid Base32 directly
        raw_clean = _re.sub(r"[\s\-_]+", "", secret_candidate).upper().rstrip("=")
        if len(raw_clean) < 6:
            continue

        # Normalize ambiguous characters: 0->O, 1->L (or I), 8->B, 9->G
        norm_variants = [
            raw_clean,
            raw_clean.replace("0", "O").replace("1", "L").replace("8", "B").replace("9", "G"),
            raw_clean.replace("0", "O").replace("1", "I").replace("8", "B").replace("9", "G"),
        ]

        code_now = None
        used_secret = raw_clean
        for variant in norm_variants:
            b32_only = _re.sub(r"[^A-Z2-7]", "", variant)
            if len(b32_only) < 6:
                continue
            pad_len = (8 - (len(b32_only) % 8)) % 8
            padded_secret = b32_only + ("=" * pad_len)
            try:
                totp = pyotp.TOTP(padded_secret)
                code_now = totp.now()
                used_secret = padded_secret
                break
            except Exception:
                continue

        if not code_now:
            code_now = _fetch_2fa_live_fallback(raw_clean)

        if code_now:
            results.append({
                "label": label,
                "secret": used_secret,
                "code": code_now,
                "remaining": rem,
            })
    return results

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

def _decode_mime_hdr(raw_val: str) -> str:
    if not raw_val:
        return ""
    try:
        from email.header import decode_header
        parts = decode_header(raw_val)
        out = []
        for chunk, enc in parts:
            if isinstance(chunk, bytes):
                out.append(chunk.decode(enc or "utf-8", errors="ignore"))
            else:
                out.append(str(chunk))
        return "".join(out).strip()
    except Exception:
        return str(raw_val).strip()

def _extract_otp_from_text(subject: str, clean_body: str) -> str:
    import re as _re
    combined = f"{subject} {clean_body}"
    ctx_m = _re.search(r"(?:code|otp|pin|verify|verification|confirm|login|use|security)[^0-9]{0,25}\b(\d{6,8})\b", combined, flags=_re.IGNORECASE)
    if ctx_m:
        return ctx_m.group(1)
    m68 = _re.search(r"\b(\d{6,8})\b", combined)
    if m68:
        return m68.group(1)
    for m45 in _re.finditer(r"\b(\d{4,5})\b", combined):
        val = m45.group(1)
        if not (2020 <= int(val) <= 2035):
            return val
    return ""

def _fetch_via_yamin_fallback(account_str: str) -> Optional[dict]:
    """Optional fast fallback via code.yamin.bd public endpoint if direct token permissions vary."""
    try:
        payload = json.dumps({"credentials": account_str}).encode("utf-8")
        req = urllib.request.Request(
            "https://code.yamin.bd/api/get-code",
            data=payload,
            headers={
                "Content-Type": "application/json",
                "Accept": "application/json",
                "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) CyberNestBot/3.0",
                "Origin": "https://code.yamin.bd",
                "Referer": "https://code.yamin.bd/",
            },
        )
        with urllib.request.urlopen(req, timeout=10) as resp:
            raw = resp.read().decode("utf-8", errors="ignore")
            data = json.loads(raw)
        rows = data if isinstance(data, list) else (data.get("results") or data.get("data") or [data])
        for r in rows:
            if not isinstance(r, dict):
                continue
            code_val = str(r.get("code") or r.get("otp") or "").strip()
            msgs = r.get("messages") or r.get("emails") or []
            if code_val or msgs:
                norm_msgs = []
                for m in msgs[:8]:
                    subj = str(m.get("subject") or "Inbox Message")
                    body_txt = str(m.get("body") or m.get("text") or m.get("preview") or "")
                    c_found = str(m.get("code") or "") or _extract_otp_from_text(subj, body_txt)
                    norm_msgs.append({
                        "sender_name": str(m.get("from_name") or m.get("sender") or "Sender"),
                        "sender_email": str(m.get("from") or m.get("sender_email") or ""),
                        "subject": subj,
                        "preview": body_txt[:140] + ("..." if len(body_txt) > 140 else ""),
                        "body": body_txt[:2000],
                        "date": str(m.get("date") or "")[:25],
                        "code": c_found,
                    })
                if not code_val and norm_msgs:
                    code_val = next((x["code"] for x in norm_msgs if x["code"]), "") or norm_msgs[0]["preview"][:40]
                if code_val:
                    return {
                        "ok": True,
                        "email": str(r.get("email") or account_str.split("|")[0].strip()),
                        "subject": norm_msgs[0]["subject"] if norm_msgs else "Verification Code",
                        "code": code_val,
                        "messages": norm_msgs,
                    }
    except Exception:
        pass
    return None

def fetch_outlook_otp(account_str: str) -> dict:
    """Super-fast 3-Layer OTP Reader (Microsoft Graph API -> All-Folder IMAP -> code.yamin.bd fallback)."""
    import re as _re
    from email.utils import parseaddr

    first_line = next((ln.strip() for ln in (account_str or "").splitlines() if ln.strip()), "")
    parts = [p.strip() for p in first_line.split("|") if p.strip()]
    if len(parts) < 4:
        return {"ok": False, "error": "Format must be: email|password|refresh_token|client_id"}
    mail_addr, _, refresh_token, client_id = parts[0], parts[1], parts[2], parts[3]

    token_url = "https://login.microsoftonline.com/common/oauth2/v2.0/token"
    access_token = None
    last_token_err = "Token refresh failed"

    # Layer 1: Try Microsoft Graph API first (0.5s ultra-fast, reads Inbox + Junk + Other simultaneously)
    for scope_try in (
        "https://graph.microsoft.com/Mail.Read offline_access",
        None,
        "https://outlook.office.com/IMAP.AccessAsUser.All offline_access",
    ):
        try:
            post_body = {
                "client_id": client_id,
                "refresh_token": refresh_token,
                "grant_type": "refresh_token",
            }
            if scope_try:
                post_body["scope"] = scope_try
            tdata = _http_post_json(token_url, post_body, timeout=12)
            if tdata.get("access_token"):
                access_token = tdata["access_token"]
                granted_scope = str(tdata.get("scope", ""))
                # If Graph token granted, fetch directly via Graph API
                if "graph.microsoft.com" in granted_scope or scope_try == "https://graph.microsoft.com/Mail.Read offline_access":
                    try:
                        g_url = "https://graph.microsoft.com/v1.0/me/messages?$top=8&$orderby=receivedDateTime%20desc&$select=subject,bodyPreview,body,from,receivedDateTime"
                        g_req = urllib.request.Request(
                            g_url,
                            headers={
                                "Authorization": f"Bearer {access_token}",
                                "Accept": "application/json",
                                "User-Agent": "CyberNestBot/3.0",
                            },
                        )
                        with urllib.request.urlopen(g_req, timeout=10) as g_resp:
                            g_json = json.loads(g_resp.read().decode("utf-8"))
                        g_items = g_json.get("value", [])
                        if g_items:
                            messages_out = []
                            for itm in g_items:
                                subj = str(itm.get("subject") or "Inbox Message").strip()
                                sender_obj = (itm.get("from") or {}).get("emailAddress") or {}
                                s_name = str(sender_obj.get("name") or "Sender").strip()
                                s_email = str(sender_obj.get("address") or "").strip()
                                raw_body = ((itm.get("body") or {}).get("content") or "") or str(itm.get("bodyPreview") or "")
                                clean_txt = _re.sub(r"\s+", " ", _re.sub(r"<[^>]+>", " ", raw_body)).strip()
                                date_str = str(itm.get("receivedDateTime") or "").replace("T", " ").replace("Z", "")[:19]
                                code_found = _extract_otp_from_text(subj, clean_txt)
                                messages_out.append({
                                    "sender_name": s_name or (s_email.split("@")[0] if "@" in s_email else "Sender"),
                                    "sender_email": s_email,
                                    "subject": subj,
                                    "preview": clean_txt[:140] + ("..." if len(clean_txt) > 140 else ""),
                                    "body": clean_txt[:2000],
                                    "date": date_str,
                                    "code": code_found,
                                })
                            latest_code = next((m["code"] for m in messages_out if m["code"]), "") or messages_out[0]["preview"][:40]
                            return {
                                "ok": True,
                                "email": mail_addr,
                                "subject": messages_out[0]["subject"],
                                "code": latest_code,
                                "messages": messages_out,
                            }
                    except Exception:
                        pass
                break
            else:
                last_token_err = tdata.get("error_description", last_token_err)
        except Exception as e:
            last_token_err = str(e)

    # Layer 2: IMAP XOAUTH2 checking BOTH 'INBOX' and 'Junk' (Spam) folders
    if access_token:
        try:
            auth_string = f"user={mail_addr}\x01auth=Bearer {access_token}\x01\x01"
            imap = imaplib.IMAP4_SSL("outlook.office365.com", 993)
            imap.authenticate("XOAUTH2", lambda x: auth_string.encode("utf-8"))

            messages_out = []
            for folder_name in ("INBOX", "Junk"):
                try:
                    st_sel, _ = imap.select(folder_name, readonly=True)
                    if st_sel != "OK":
                        continue
                    _, data = imap.search(None, "ALL")
                    ids = data[0].split() if data and data[0] else []
                    if not ids:
                        continue
                    for mid in reversed(ids[-5:]):
                        _, msg_data = imap.fetch(mid, "(RFC822)")
                        if not msg_data or not msg_data[0]:
                            continue
                        raw_email = msg_data[0][1]
                        msg = email_lib.message_from_bytes(raw_email)
                        subject = _decode_mime_hdr(msg.get("Subject", "Inbox Message")) or "Inbox Message"
                        if folder_name == "Junk":
                            subject = f"[Junk] {subject}"
                        from_raw = _decode_mime_hdr(msg.get("From", ""))
                        s_name, s_email = parseaddr(from_raw)
                        s_name = s_name or (s_email.split("@")[0] if "@" in s_email else "Sender")
                        date_str = str(msg.get("Date", "")).strip()[:25]

                        body_plain = ""
                        body_html = ""
                        if msg.is_multipart():
                            for part in msg.walk():
                                ctype = part.get_content_type()
                                if ctype == "text/plain" and not body_plain:
                                    try:
                                        body_plain = part.get_payload(decode=True).decode(errors="ignore")
                                    except Exception:
                                        pass
                                elif ctype == "text/html" and not body_html:
                                    try:
                                        body_html = part.get_payload(decode=True).decode(errors="ignore")
                                    except Exception:
                                        pass
                        else:
                            try:
                                raw_b = msg.get_payload(decode=True).decode(errors="ignore")
                                if "<html" in raw_b.lower():
                                    body_html = raw_b
                                else:
                                    body_plain = raw_b
                            except Exception:
                                pass

                        raw_text = body_plain or _re.sub(r"<[^>]+>", " ", body_html)
                        clean_txt = _re.sub(r"\s+", " ", raw_text).strip()
                        code_found = _extract_otp_from_text(subject, clean_txt)
                        messages_out.append({
                            "sender_name": s_name,
                            "sender_email": s_email or from_raw,
                            "subject": subject,
                            "preview": clean_txt[:140] + ("..." if len(clean_txt) > 140 else ""),
                            "body": clean_txt[:2000],
                            "date": date_str,
                            "code": code_found,
                        })
                except Exception:
                    continue

            try:
                imap.logout()
            except Exception:
                pass

            if messages_out:
                latest_code = next((m["code"] for m in messages_out if m["code"]), "") or messages_out[0]["preview"][:40]
                return {
                    "ok": True,
                    "email": mail_addr,
                    "subject": messages_out[0]["subject"],
                    "code": latest_code,
                    "messages": messages_out,
                }
        except Exception as imap_err:
            last_token_err = str(imap_err)

    # Layer 3: Fallback via code.yamin.bd if direct IMAP/Graph returned empty
    yamin_res = _fetch_via_yamin_fallback(first_line)
    if yamin_res and yamin_res.get("ok"):
        return yamin_res

    if access_token:
        return {"ok": False, "error": "Inbox & Junk folders are currently empty. Tap 🔄 Refresh Code in 3 seconds."}
    return {"ok": False, "error": last_token_err}

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
        "custom_rules": get_setting("ai_custom_rules", "").strip(),
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
        + (f"\n\nMANDATORY ADMIN INSTRUCTIONS (Strictly obey these rules without contradicting them, and craft a helpful, natural explanation around them):\n{ctx['custom_rules']}" if ctx.get("custom_rules") else "")
    )

    uid_key = int(user_id or 0)
    history = _AI_CHAT_HISTORY.get(uid_key, [])[-6:]

    if groq_key:
        messages = [{"role": "system", "content": sys_prompt}]
        for h in history:
            messages.append(h)
        messages.append({"role": "user", "content": question})

        for model_name in ("openai/gpt-oss-120b", "openai/gpt-oss-20b", "llama-3.1-8b-instant"):
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

def _sanitize_ai_for_telegram(text: str) -> str:
    import re as _re
    s = (text or "").strip()
    s = _re.sub(r"<br\s*/?>", "\n", s, flags=_re.IGNORECASE)
    s = _re.sub(r"</?(?:ul|ol|p|div)>", "\n", s, flags=_re.IGNORECASE)
    s = _re.sub(r"<li[^>]*>", "• ", s, flags=_re.IGNORECASE)
    s = _re.sub(r"</li>", "\n", s, flags=_re.IGNORECASE)
    s = _re.sub(r"\*\*(.+?)\*\*", r"<b>\1</b>", s)
    s = _re.sub(r"\n{3,}", "\n\n", s)
    return s.strip()

def generate_smart_ai_reply(question: str, user_id: Optional[int] = None) -> str:
    """Conversational AI Assistant: understands Bangla, Banglish & English and answers the exact question."""
    raw_q = (question or "").strip()
    q = raw_q.lower()
    ctx = _build_live_store_context(user_id=user_id)

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

    if any(g in q for g in ["assalamu", "salam", "সালাম", "আসসালামু"]):
        return (
            f"ওয়ালাইকুম আসসালাম, {html.escape(ctx['user_name'])}! 😊\n"
            f"কেমন আছেন? আপনার কোনো প্রশ্ন বা অসুবিধা থাকলে আমাকে নির্দ্বিধায় জিজ্ঞেস করুন—আমি সাথে সাথে সাহায্য করব।"
        )
    if any(g in q for g in ["হ্যালো", "হাই", "কেমন আছ", "kemon ach", "kemon as", "ki khobor", "কি খবর", "hello", "hlw", "hey", "how are you"]) or q == "hi":
        if is_bangla_ctx or has_bn_chars:
            return f"হ্যালো {html.escape(ctx['user_name'])}! 😊 আলহামদুলিল্লাহ ভালো আছি। আপনি কেমন আছেন? বলুন, আজ আপনাকে কীভাবে সাহায্য করতে পারি?"
        return f"Hello {html.escape(ctx['user_name'])}! 😊 I'm doing great! How can I help you today?"

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

    if is_bangla_ctx:
        return (
            f"আমি আপনার প্রশ্নটি পুরোপুরি বুঝতে পারিনি। আপনি নির্দিষ্ট করে জিজ্ঞেস করতে পারেন—যেমন:\n"
            f"• <i>\"VPN কীভাবে কিনব?\"</i>\n"
            f"• <i>\"Rapid Proxy-এর দাম কত?\"</i>\n"
            f"• <i>\"ডিপোজিট করব কীভাবে?\"</i>\n\n"
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

        if path in ("/app", "/app/", "/webapp.html"):
            content = load_webapp_html_bytes()
            self.send_response(200)
            self.send_header("Content-type", "text/html; charset=utf-8")
            self.send_header("Cache-Control", "no-cache")
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
                       COALESCE(p.promo_tagline, '') as promo_tagline,
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

            c.execute("SELECT id, title, chat_id, invite_link, is_active FROM force_joins WHERE is_active = 1 ORDER BY id ASC")
            fj_list = [dict(r) for r in c.fetchall()]

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

            force_join_required = False
            missing_channels = []
            if not is_admin(uid) and _GLOBAL_EVENT_LOOP and _GLOBAL_BOT_APP:
                try:
                    fut = asyncio.run_coroutine_threadsafe(check_all_force_joins(_GLOBAL_BOT_APP.bot, uid), _GLOBAL_EVENT_LOOP)
                    all_ok, missing_fj = fut.result(timeout=8)
                    if not all_ok:
                        force_join_required = True
                        missing_channels = missing_fj
                except Exception as fj_err:
                    logger.warning(f"/api/init force join check error: {fj_err}")

            vip_info = get_vip_badge_info(ostat["spent"])
            tg_photo_url = tg_u.get("photo_url") or f"/api/user_photo?user_id={uid}"

            return self._send_json({
                "force_join_required": force_join_required,
                "missing_channels": missing_channels,
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
                "force_joins": fj_list,
                "settings": sdict,
                "menu_layout": get_menu_layout(),
                "ai_custom_rules": get_setting("ai_custom_rules", ""),
                "unread_notifs": unread_n,
            })

        if path == "/api/user/balance":
            uid, _ = self._get_auth_user()
            if not uid:
                return self._send_json({"ok": False, "error": "Unauthorized"}, 401)
            conn = get_db()
            c = conn.cursor()
            c.execute("SELECT balance_bdt FROM users WHERE user_id = ?", (uid,))
            urow = c.fetchone()
            bal = float(urow["balance_bdt"] or 0.0) if urow else 0.0
            c.execute("SELECT COUNT(*) as cnt, COALESCE(SUM(total_bdt), 0) as spent FROM orders WHERE user_id = ? AND status != 'refunded'", (uid,))
            ostat = dict(c.fetchone())
            c.execute("SELECT COUNT(*) as cnt FROM orders WHERE user_id = ? AND status = 'refunded'", (uid,))
            cancelled_cnt = c.fetchone()["cnt"]
            c.execute("SELECT COUNT(*) as cnt FROM notifications WHERE user_id = ? AND is_read = 0", (uid,))
            unread_n = c.fetchone()["cnt"]
            conn.close()
            return self._send_json({
                "ok": True,
                "balance_bdt": bal,
                "total_orders": ostat["cnt"],
                "total_spent": float(ostat["spent"] or 0.0),
                "cancelled_orders": cancelled_cnt,
                "unread_notifs": unread_n,
                "vip": get_vip_badge_info(ostat["spent"]),
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

            c.execute("SELECT code, discount_pct, is_active, COALESCE(product_id, 0) as product_id, COALESCE(product_name, 'All Products') as product_name FROM coupons ORDER BY code ASC")
            coupons_list = [dict(r) for r in c.fetchall()]

            c.execute("SELECT id, title, chat_id, invite_link, is_active FROM force_joins ORDER BY id ASC")
            force_joins_list = [dict(r) for r in c.fetchall()]

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
                "force_joins": force_joins_list,
                "banners": banners_list,
                "menu_layout": get_menu_layout(),
                "ai_custom_rules": get_setting("ai_custom_rules", ""),
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

            if path == "/api/notifications/add":
                title = str(body.get("title", "🔔 Notification")).strip()
                msg = str(body.get("message", "")).strip()
                if msg:
                    add_notification(uid, title, msg)
                return self._send_json({"ok": True})

            if path == "/api/user/preferences":
                lang = str(body.get("language", "EN")).strip().upper()
                conn = get_db()
                c = conn.cursor()
                c.execute("UPDATE users SET language = ? WHERE user_id = ?", (lang, uid))
                conn.commit()
                conn.close()
                return self._send_json({"ok": True, "language": lang})

            if path == "/api/ai_support":
                question = str(body.get("question", "")).strip()
                ans = generate_smart_ai_reply(question, user_id=uid)
                return self._send_json({"ok": True, "reply": ans})

            if path == "/api/verify_force_join":
                if is_admin(uid):
                    return self._send_json({"ok": True, "joined": True, "missing_channels": []})
                if _GLOBAL_EVENT_LOOP and _GLOBAL_BOT_APP:
                    try:
                        fut = asyncio.run_coroutine_threadsafe(check_all_force_joins(_GLOBAL_BOT_APP.bot, uid), _GLOBAL_EVENT_LOOP)
                        all_ok, missing_fj = fut.result(timeout=8)
                        return self._send_json({
                            "ok": True,
                            "joined": bool(all_ok),
                            "missing_channels": missing_fj if not all_ok else [],
                        })
                    except Exception as fj_err:
                        return self._send_json({"ok": False, "joined": False, "error": str(fj_err)})
                return self._send_json({"ok": True, "joined": True, "missing_channels": []})

            if path == "/api/coupon/check":
                code_str = str(body.get("code", "")).strip().upper()
                req_pid = int(body.get("product_id", 0) or 0)
                conn = get_db()
                c = conn.cursor()
                c.execute("SELECT discount_pct, COALESCE(product_id, 0) as product_id, COALESCE(product_name, 'All Products') as product_name FROM coupons WHERE code = ? AND is_active = 1", (code_str,))
                row = c.fetchone()
                conn.close()
                if row:
                    c_pid = int(row["product_id"] or 0)
                    if c_pid > 0 and req_pid > 0 and c_pid != req_pid:
                        return self._send_json({"ok": False, "error": f"This coupon is only valid for {row['product_name']}"})
                    return self._send_json({"ok": True, "discount_pct": row["discount_pct"], "product_id": c_pid, "product_name": row["product_name"]})
                return self._send_json({"ok": False, "error": "Invalid or expired coupon"})

            if path == "/api/tools/2fa":
                raw_sec = str(body.get("secret", "")).strip()
                parsed_codes = generate_universal_2fa_codes(raw_sec)
                if parsed_codes:
                    first = parsed_codes[0]
                    return self._send_json({"ok": True, "code": first["code"], "remaining": first["remaining"], "all": parsed_codes})
                return self._send_json({"ok": False, "error": "Invalid 2FA Secret Key"})

            if path == "/api/tools/outlook":
                res = fetch_outlook_otp(str(body.get("account", "")))
                return self._send_json(res)

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
                    async def _send_web_dep_alert():
                        kb = colorize_markup(InlineKeyboardMarkup([[
                            InlineKeyboardButton("✅ Approve", callback_data=f"adm_dep_approve_{dep_id}"),
                            InlineKeyboardButton("❌ Reject", callback_data=f"adm_dep_reject_{dep_id}")
                        ]]))
                        try:
                            sent = await _GLOBAL_BOT_APP.bot.send_message(
                                chat_id=SUPER_ADMIN_ID,
                                text=(
                                    f"🔔 <b>Mini App Deposit Request</b>\n"
                                    f"👤 User: {tg_u.get('first_name', 'User')} (<code>{uid}</code>)\n"
                                    f"💳 Method: <b>{method}</b>\n"
                                    f"💰 Amount: <b>{amount:.2f} BDT</b>\n"
                                    f"🔑 TrxID: <code>{trx_id}</code>"
                                ),
                                parse_mode=ParseMode.HTML,
                                reply_markup=kb,
                            )
                            _c2 = get_db()
                            _c2.cursor().execute("UPDATE deposits SET admin_msg_id = ? WHERE id = ?", (sent.message_id, dep_id))
                            _c2.commit()
                            _c2.close()
                        except Exception:
                            pass
                    asyncio.run_coroutine_threadsafe(_send_web_dep_alert(), _GLOBAL_EVENT_LOOP)
                return self._send_json({"ok": True})

            if path == "/api/order":
                pid = int(body.get("product_id", 0))
                qty = max(1, int(body.get("quantity", 1)))
                coupon = str(body.get("coupon", "")).strip().upper()
                if not is_admin(uid) and _GLOBAL_EVENT_LOOP and _GLOBAL_BOT_APP:
                    try:
                        fut = asyncio.run_coroutine_threadsafe(check_all_force_joins(_GLOBAL_BOT_APP.bot, uid), _GLOBAL_EVENT_LOOP)
                        all_ok, missing_fj = fut.result(timeout=8)
                        if not all_ok:
                            return self._send_json({
                                "ok": False,
                                "error": "কেনার আগে সবগুলো আবশ্যিক গ্রুপ ও চ্যানেলে জয়েন করুন!",
                                "force_join_required": True,
                                "missing_channels": missing_fj,
                                "force_join_missing": missing_fj,
                            })
                    except Exception:
                        pass
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
                    c.execute("SELECT discount_pct, COALESCE(product_id, 0) as product_id, COALESCE(product_name, 'All Products') as product_name FROM coupons WHERE code = ? AND is_active = 1", (coupon,))
                    crow = c.fetchone()
                    if crow:
                        c_pid = int(crow["product_id"] or 0)
                        if c_pid == 0 or c_pid == pid:
                            disc = float(crow["discount_pct"])
                        else:
                            conn.close()
                            return self._send_json({"ok": False, "error": f"Coupon {coupon} is only valid for {crow['product_name']}"})
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
                            adm_kb = colorize_markup(InlineKeyboardMarkup([
                                [InlineKeyboardButton(f"🚀 Deliver Order #{oid}", callback_data=f"adm_fulfill_{oid}")],
                                [InlineKeyboardButton(f"❌ Cancel & Refund #{oid}", callback_data=f"adm_refund_{oid}")]
                            ]))
                            try:
                                sent_ord = await bot_inst.send_message(
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
                                _c3 = get_db()
                                _c3.cursor().execute("UPDATE orders SET admin_msg_id = ? WHERE id = ?", (sent_ord.message_id, oid))
                                _c3.commit()
                                _c3.close()
                            except Exception:
                                pass

                        asyncio.run_coroutine_threadsafe(_notify_web_manual(), _GLOBAL_EVENT_LOOP)
                    return self._send_json({"ok": True, "message": f"🎉 অর্ডার #{oid} গ্রহণ করা হয়েছে! অফিশিয়াল PDF রিসিট আপনার ইনবক্সে পাঠানো হয়েছে এবং ১৫-৩০ মিনিটের মধ্যে ডেলিভারি পাবেন।"})

            if path.startswith("/api/admin/"):
                if not is_admin(uid):
                    return self._send_json({"ok": False, "error": "Forbidden"}, 403)
                conn = get_db()
                c = conn.cursor()

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
                        adm_mid = o["admin_msg_id"] if "admin_msg_id" in o.keys() else None
                        async def _send_manual_deliv():
                            bot_inst = _GLOBAL_BOT_APP.bot
                            if adm_mid:
                                await _delete_admin_msg_safe(bot_inst, adm_mid)
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
                            adm_mid = o["admin_msg_id"] if "admin_msg_id" in o.keys() else None
                            async def _notify_refund():
                                if adm_mid:
                                    await _delete_admin_msg_safe(_GLOBAL_BOT_APP.bot, adm_mid)
                                try:
                                    await _GLOBAL_BOT_APP.bot.send_message(
                                        chat_id=o["user_id"],
                                        text=f"❌ আপনার অর্ডার #{oid} বাতিল করা হয়েছে এবং <b>{o['total_bdt']:.2f} BDT</b> আপনার ওয়ালেটে রিফান্ড করা হয়েছে।",
                                        parse_mode=ParseMode.HTML,
                                    )
                                except Exception:
                                    pass
                            asyncio.run_coroutine_threadsafe(_notify_refund(), _GLOBAL_EVENT_LOOP)
                    conn.close()
                    return self._send_json({"ok": True})

                if path == "/api/admin/deposit_action":
                    did = int(body.get("deposit_id", 0))
                    act = body.get("action")
                    c.execute("SELECT * FROM deposits WHERE id = ?", (did,))
                    d = c.fetchone()
                    if d and d["status"] == "pending":
                        adm_mid = d["admin_msg_id"] if "admin_msg_id" in d.keys() else None
                        if act == "approve":
                            c.execute("UPDATE deposits SET status = 'approved' WHERE id = ?", (did,))
                            c.execute("UPDATE users SET balance_bdt = balance_bdt + ? WHERE user_id = ?", (d["amount"], d["user_id"]))
                            add_notification(d["user_id"], "✅ Deposit Approved", f"+{d['amount']:.2f} BDT added to your wallet!")
                            if _GLOBAL_EVENT_LOOP and _GLOBAL_BOT_APP:
                                async def _dep_approve_sync():
                                    if adm_mid:
                                        await _delete_admin_msg_safe(_GLOBAL_BOT_APP.bot, adm_mid)
                                    try:
                                        await _GLOBAL_BOT_APP.bot.send_message(
                                            chat_id=d["user_id"],
                                            text=f"🎉 <b>ডিপোজিট সফল!</b>\nআপনার অ্যাকাউন্টে <b>{d['amount']:.2f} BDT</b> যুক্ত হয়েছে।",
                                            parse_mode=ParseMode.HTML,
                                        )
                                    except Exception:
                                        pass
                                asyncio.run_coroutine_threadsafe(_dep_approve_sync(), _GLOBAL_EVENT_LOOP)
                        else:
                            c.execute("UPDATE deposits SET status = 'rejected' WHERE id = ?", (did,))
                            add_notification(d["user_id"], "❌ Deposit Rejected", f"Your deposit ({d['amount']:.2f} BDT) was rejected.")
                            if _GLOBAL_EVENT_LOOP and _GLOBAL_BOT_APP:
                                async def _dep_reject_sync():
                                    if adm_mid:
                                        await _delete_admin_msg_safe(_GLOBAL_BOT_APP.bot, adm_mid)
                                    try:
                                        await _GLOBAL_BOT_APP.bot.send_message(
                                            chat_id=d["user_id"],
                                            text=f"❌ আপনার <b>{d['amount']:.2f} BDT</b> ডিপোজিট রিকোয়েস্টটি বাতিল করা হয়েছে।",
                                            parse_mode=ParseMode.HTML,
                                        )
                                    except Exception:
                                        pass
                                asyncio.run_coroutine_threadsafe(_dep_reject_sync(), _GLOBAL_EVENT_LOOP)
                        conn.commit()
                    conn.close()
                    return self._send_json({"ok": True})

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
                            rm = colorize_markup(InlineKeyboardMarkup([[InlineKeyboardButton(b["text"], url=b["url"])] for b in valid_btns])) if valid_btns else None
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

                if path == "/api/admin/brand_photo":
                    sid = int(body.get("sub_id", 0))
                    img_val = str(body.get("brand_image", "")).strip() or None
                    if sid:
                        c.execute("UPDATE subcategories SET brand_image = ? WHERE id = ?", (img_val, sid))
                        conn.commit()
                    conn.close()
                    return self._send_json({"ok": True})

                if path == "/api/admin/ai_settings":
                    act = str(body.get("action", "save"))
                    new_key = str(body.get("ai_api_key", "")).strip()
                    if act == "save_rules":
                        rules_txt = str(body.get("ai_custom_rules", "")).strip()
                        set_setting("ai_custom_rules", rules_txt)
                        conn.close()
                        return self._send_json({"ok": True, "message": "Custom AI Instructions saved!"})
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
                        last_err = "Unknown error"
                        for tm in ("openai/gpt-oss-120b", "openai/gpt-oss-20b", "llama-3.1-8b-instant"):
                            try:
                                tp = json.dumps({
                                    "model": tm,
                                    "messages": [{"role": "user", "content": "Say 'AI Connected Successfully!' in Bangla in 1 short sentence."}],
                                    "max_tokens": 80,
                                }).encode("utf-8")
                                tr = urllib.request.Request(
                                    "https://api.groq.com/openai/v1/chat/completions",
                                    data=tp,
                                    headers={"Content-Type": "application/json", "Authorization": f"Bearer {test_key}", "User-Agent": "CyberNestBot/2.0"},
                                )
                                with urllib.request.urlopen(tr, timeout=10) as resp:
                                    td = json.loads(resp.read().decode("utf-8"))
                                return self._send_json({"ok": True, "provider": f"GroqCloud ({tm})", "reply": td["choices"][0]["message"]["content"].strip()})
                            except Exception as e:
                                last_err = str(e)
                        return self._send_json({"ok": False, "error": f"Groq Error: {last_err}"})
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
                    if "promo_tagline" in body:
                        c.execute("UPDATE products SET promo_tagline = ? WHERE id = ?", (str(body["promo_tagline"] or "").strip(), pid))
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

                if path == "/api/admin/payment_settings":
                    for k in ("bkash_number", "nagad_number", "binance_id", "min_deposit_manual", "usdt_rate", "support_username"):
                        if k in body and body[k] is not None:
                            set_setting(k, str(body[k]).strip())
                    conn.close()
                    return self._send_json({"ok": True})

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
                        c.execute("SELECT code, discount_pct, is_active, COALESCE(product_id, 0) as product_id, COALESCE(product_name, 'All Products') as product_name FROM coupons ORDER BY code ASC")
                        clist = [dict(r) for r in c.fetchall()]
                        conn.close()
                        return self._send_json({"ok": True, "coupons": clist})
                    pct = float(body.get("discount_pct", 0))
                    prod_id = int(body.get("product_id", 0) or 0)
                    prod_name = str(body.get("product_name", "All Products") or "All Products").strip()
                    if prod_id > 0 and (not prod_name or prod_name == "All Products"):
                        c.execute("SELECT p.name, s.name as sub_name FROM products p JOIN subcategories s ON p.subcategory_id = s.id WHERE p.id = ?", (prod_id,))
                        pr_r = c.fetchone()
                        if pr_r:
                            prod_name = f"{pr_r['sub_name']} — {pr_r['name']}"
                    if prod_id == 0:
                        prod_name = "All Products"
                    if conn.is_pg:
                        c.execute(
                            "INSERT INTO coupons (code, discount_pct, is_active, product_id, product_name) VALUES (%s, %s, 1, %s, %s) ON CONFLICT (code) DO UPDATE SET discount_pct = EXCLUDED.discount_pct, is_active = 1, product_id = EXCLUDED.product_id, product_name = EXCLUDED.product_name",
                            (code_s, pct, prod_id, prod_name),
                        )
                    else:
                        c.execute(
                            "INSERT OR REPLACE INTO coupons (code, discount_pct, is_active, product_id, product_name) VALUES (?, ?, 1, ?, ?)",
                            (code_s, pct, prod_id, prod_name),
                        )
                    conn.commit()
                    c.execute("SELECT code, discount_pct, is_active, COALESCE(product_id, 0) as product_id, COALESCE(product_name, 'All Products') as product_name FROM coupons ORDER BY code ASC")
                    clist = [dict(r) for r in c.fetchall()]
                    conn.close()
                    return self._send_json({"ok": True, "coupons": clist})

                if path == "/api/admin/force_joins":
                    act = str(body.get("action", "list"))
                    if act == "toggle_enabled":
                        cur_fj_en = get_setting("force_join_enabled", "1")
                        set_setting("force_join_enabled", "0" if cur_fj_en == "1" else "1")
                    elif act == "add":
                        title = str(body.get("title", "")).strip() or "Required Channel"
                        chat_id = str(body.get("chat_id", "")).strip()
                        invite_link = str(body.get("invite_link", "")).strip()
                        if not chat_id or not invite_link:
                            conn.close()
                            return self._send_json({"ok": False, "error": "Chat ID and Invite Link are required"})
                        c.execute(
                            "INSERT INTO force_joins (title, chat_id, invite_link, is_active) VALUES (?, ?, ?, 1)",
                            (title, chat_id, invite_link),
                        )
                        conn.commit()
                    elif act == "edit":
                        fid = int(body.get("id", 0))
                        title = str(body.get("title", "")).strip()
                        chat_id = str(body.get("chat_id", "")).strip()
                        invite_link = str(body.get("invite_link", "")).strip()
                        if fid and title and chat_id and invite_link:
                            c.execute(
                                "UPDATE force_joins SET title = ?, chat_id = ?, invite_link = ? WHERE id = ?",
                                (title, chat_id, invite_link, fid),
                            )
                            conn.commit()
                    elif act == "toggle":
                        fid = int(body.get("id", 0))
                        if fid:
                            c.execute("UPDATE force_joins SET is_active = CASE WHEN COALESCE(is_active,1)=1 THEN 0 ELSE 1 END WHERE id = ?", (fid,))
                            conn.commit()
                    elif act == "delete":
                        fid = int(body.get("id", 0))
                        if fid:
                            c.execute("DELETE FROM force_joins WHERE id = ?", (fid,))
                            conn.commit()
                    c.execute("SELECT id, title, chat_id, invite_link, is_active FROM force_joins ORDER BY id ASC")
                    fj_all = [dict(r) for r in c.fetchall()]
                    conn.close()
                    fj_en = (get_setting("force_join_enabled", "1") == "1")
                    return self._send_json({"ok": True, "enabled": fj_en, "channels": fj_all, "force_joins": fj_all})

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
        promo_tagline TEXT DEFAULT '',
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
        admin_msg_id BIGINT DEFAULT NULL,
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
        admin_msg_id BIGINT DEFAULT NULL,
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
        is_active INTEGER DEFAULT 1,
        product_id INTEGER DEFAULT 0,
        product_name TEXT DEFAULT 'All Products'
    )
    """)

    cursor.execute("""
    CREATE TABLE IF NOT EXISTS force_joins (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        title TEXT NOT NULL,
        chat_id TEXT NOT NULL,
        invite_link TEXT NOT NULL,
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
            cursor.execute("ALTER TABLE products ADD COLUMN IF NOT EXISTS promo_tagline TEXT DEFAULT ''")
            cursor.execute("ALTER TABLE orders ADD COLUMN IF NOT EXISTS admin_msg_id BIGINT DEFAULT NULL")
            cursor.execute("ALTER TABLE deposits ADD COLUMN IF NOT EXISTS admin_msg_id BIGINT DEFAULT NULL")
            cursor.execute("ALTER TABLE coupons ADD COLUMN IF NOT EXISTS product_id INTEGER DEFAULT 0")
            cursor.execute("ALTER TABLE coupons ADD COLUMN IF NOT EXISTS product_name TEXT DEFAULT 'All Products'")
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
            "ALTER TABLE products ADD COLUMN promo_tagline TEXT DEFAULT ''",
            "ALTER TABLE subcategories ADD COLUMN brand_image TEXT DEFAULT NULL",
            "ALTER TABLE users ADD COLUMN referred_by BIGINT DEFAULT NULL",
            "ALTER TABLE users ADD COLUMN referral_earnings REAL DEFAULT 0.0",
            "ALTER TABLE users ADD COLUMN pending_referral_tk REAL DEFAULT 0.0",
            "ALTER TABLE users ADD COLUMN language TEXT DEFAULT 'EN'",
            "ALTER TABLE orders ADD COLUMN admin_msg_id BIGINT DEFAULT NULL",
            "ALTER TABLE deposits ADD COLUMN admin_msg_id BIGINT DEFAULT NULL",
            "ALTER TABLE coupons ADD COLUMN product_id INTEGER DEFAULT 0",
            "ALTER TABLE coupons ADD COLUMN product_name TEXT DEFAULT 'All Products'",
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

    cursor.execute("SELECT COUNT(*) as cnt FROM force_joins")
    fj_row = cursor.fetchone()
    fj_cnt = fj_row["cnt"] if isinstance(fj_row, dict) else fj_row[0]
    if fj_cnt == 0 and not get_setting("fj_seeded", ""):
        try:
            if WELCOME_GROUP_ID and WELCOME_GROUP_LINK:
                cursor.execute(
                    "INSERT INTO force_joins (title, chat_id, invite_link, is_active) VALUES (?, ?, ?, 1)",
                    ("📢 Join Update Group", str(WELCOME_GROUP_ID), str(WELCOME_GROUP_LINK)),
                )
            if TRANSACTION_CHANNEL_ID and TRANSACTION_CHANNEL_LINK:
                cursor.execute(
                    "INSERT INTO force_joins (title, chat_id, invite_link, is_active) VALUES (?, ?, ?, 1)",
                    ("💳 Join Proof Channel", str(TRANSACTION_CHANNEL_ID), str(TRANSACTION_CHANNEL_LINK)),
                )
            if conn.is_pg:
                cursor.execute("INSERT INTO settings (key, value) VALUES (%s, %s) ON CONFLICT (key) DO NOTHING", ("fj_seeded", "1"))
            else:
                cursor.execute("INSERT OR IGNORE INTO settings (key, value) VALUES (?, ?)", ("fj_seeded", "1"))
            conn.commit()
        except Exception as e:
            logger.warning(f"force_joins seed notice: {e}")

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
    buttons = [[make_reply_btn(label, group="main_menu") for label in row] for row in layout]
    if is_admin(user_id):
        buttons.append([make_reply_btn("👑 Admin Panel", group="main_menu")])
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

def format_credentials_card(items: List[str], prod_name: str, cat_name: str) -> str:
    """Formats credentials with one-tap copy code blocks."""
    is_proxy = "PROXY" in cat_name.upper() or "PROXY" in prod_name.upper()
    is_mail = "MAIL" in cat_name.upper() or "HOTMAIL" in prod_name.upper() or "OUTLOOK" in prod_name.upper()
    lines = []

    if is_proxy:
        lines.append("✅ <b>Proxy Delivered!</b>")
        lines.append(f"🌐 <b>{html.escape(prod_name)}</b>")
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
            lines.append(f"🖥️ <b>Host</b>     : <code>{html.escape(host)}</code>")
            lines.append(f"🔌 <b>Port</b>     : <code>{html.escape(port)}</code>")
            lines.append(f"👤 <b>Username</b> : <code>{html.escape(u_name)}</code>")
            lines.append(f"🔐 <b>Password</b> : <code>{html.escape(pwd)}</code>")
            lines.append(f"📋 <b>Full</b>     : <code>{html.escape(item)}</code>")
            if idx < len(items):
                lines.append("")
    elif is_mail:
        lines.append("✅ <b>Mail Delivered!</b>")
        lines.append(f"📧 <b>{len(items)}× {html.escape(prod_name)}</b>")
        lines.append("────────────────────────")
        for idx, item in enumerate(items, start=1):
            full_str = item.strip()
            delim = "|" if "|" in full_str else ":"
            email_only = full_str.split(delim)[0].strip()

            lines.append(f"📧 <b>#{idx}</b>")
            lines.append(f"✉️ <b>Mail</b> : <code>{html.escape(email_only)}</code>")
            lines.append(f"<code>{html.escape(full_str)}</code>")
            if idx < len(items):
                lines.append("")
    else:
        lines.append("✅ <b>Service Delivered!</b>")
        lines.append(f"🛡️ <b>{html.escape(prod_name)}</b>")
        lines.append("────────────────────────")
        for idx, item in enumerate(items, start=1):
            delim = "|" if "|" in item else ":"
            parts = [p.strip() for p in item.split(delim) if p.strip()]
            if len(parts) >= 2:
                email, pwd = parts[0], parts[1]
                if len(items) > 1:
                    lines.append(f"📦 <b>Item #{idx}</b>")
                lines.append(f"📧 <b>Email / User</b> : <code>{html.escape(email)}</code>")
                lines.append(f"🔐 <b>Password</b>     : <code>{html.escape(pwd)}</code>")
                if len(parts) >= 3:
                    tok = parts[2]
                    lines.append(f"🎟️ <b>Token / Session</b> : <code>{html.escape(tok)}</code>")
            else:
                if len(items) > 1:
                    lines.append(f"🔑 <b>Key / Code #{idx}</b>")
                lines.append(f"🎟️ <b>Key / Code</b> : <code>{html.escape(item)}</code>")
            if idx < len(items):
                lines.append("")

    lines.append("────────────────────────")
    lines.append("💡 <i>Tap any value above to copy</i>")
    return "\n".join(lines)



# ------------------------------------------------------------------------------
# 7. USER HANDLERS (New vs Returning Welcome, AI Assistant, Shop, Profile, Deposit)
# ------------------------------------------------------------------------------
def build_welcome_message(full_name: str, shop_name: str, is_new: bool) -> str:
    safe_name = html.escape(full_name or "গ্রাহক")
    if is_new:
        default_txt = (
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
        return render_custom_bot_message("welcome_new", default_txt, name=safe_name, shop_name=shop_name)
    else:
        default_txt = (
            f"👋 <b>স্বাগতম, {safe_name}!</b> ✨\n"
            f"💎 <b>{shop_name}</b> — ২৪/৭ অটোমেটেড ডিজিটাল স্টোর।\n"
            f"─────────────────────────────\n"
            f"🛍️ প্রিমিয়াম VPN, প্রিমিয়াম Proxy ও প্রিমিয়াম Mail কিনতে নিচের মেনু অথবা <b>Open Shop</b> বাটন ব্যবহার করুন 👇"
        )
        return render_custom_bot_message("welcome_back", default_txt, name=safe_name, shop_name=shop_name)

def _format_2fa_live_text(parsed_codes: List[Dict[str, Any]], live_active: bool = True, note: str = "") -> str:
    rem = parsed_codes[0]["remaining"]
    status_line = f"⏳ লাইভ টাইম বাকি: <b>{rem} সেকেন্ড</b> (প্রতি ৫ সেকেন্ডে অটো আপডেট)" if live_active else f"⏳ মেয়াদ বাকি: <b>{rem} সেকেন্ড</b>"
    if note:
        status_line += f"\n{note}"
    if len(parsed_codes) == 1:
        item = parsed_codes[0]
        lbl_str = f" — <b>{html.escape(item['label'])}</b>" if item.get("label") else ""
        return (
            f"🔑 <b>2FA Live Authenticator{lbl_str}</b>\n"
            f"─────────────────────────────\n"
            f"🔢 কোড: <code>{item['code']}</code>\n"
            f"{status_line}"
        )
    lines_out = [f"🔑 <b>2FA Live Codes ({len(parsed_codes)} Keys)</b>", "─────────────────────────────"]
    for idx, it in enumerate(parsed_codes, start=1):
        lbl = f" ({html.escape(it['label'])})" if it["label"] else ""
        lines_out.append(f"{idx}.{lbl} <code>{it['code']}</code>")
    lines_out.append(f"\n{status_line}")
    return "\n".join(lines_out)

def _build_2fa_live_kb(tok_key: str, live: bool = True) -> InlineKeyboardMarkup:
    if live:
        return InlineKeyboardMarkup([[
            make_inline_btn("🔄 Refresh 2FA Code", callback_data=f"tool_ref_2fa_{tok_key}"),
            make_inline_btn("⏹️ Off", callback_data=f"tool_stop_2fa_{tok_key}"),
        ]])
    return InlineKeyboardMarkup([[
        make_inline_btn("🔄 Refresh 2FA Code", callback_data=f"tool_ref_2fa_{tok_key}")
    ]])

def _build_mail_live_kb(tok_key: str, live: bool = True) -> InlineKeyboardMarkup:
    if live:
        return InlineKeyboardMarkup([[
            make_inline_btn("🔄 Refresh Code", callback_data=f"tool_ref_mail_{tok_key}"),
            make_inline_btn("⏹️ Off", callback_data=f"tool_stop_mail_{tok_key}"),
        ]])
    return InlineKeyboardMarkup([[
        make_inline_btn("🔄 Refresh Code", callback_data=f"tool_ref_mail_{tok_key}")
    ]])

async def _run_live_2fa_countdown(bot, chat_id: int, message_id: int, raw_secret: str, tok_key: str):
    """Edits the 2FA message every 5 seconds for 60 seconds, auto-rotating codes unless stopped via Off button."""
    for step in range(12):  # 12 * 5s = 60 seconds
        if tok_key in _LIVE_TOOL_STOP:
            break
        await asyncio.sleep(5)
        if tok_key in _LIVE_TOOL_STOP:
            break
        parsed_codes = generate_universal_2fa_codes(raw_secret)
        if not parsed_codes:
            break
        is_last = (step == 11)
        note_txt = "ℹ️ <i>৬০ সেকেন্ড লাইভ সেশন শেষ — নতুন কোডের জন্য নিচের Refresh বাটনে চাপুন।</i>" if is_last else ""
        new_txt = _format_2fa_live_text(parsed_codes, live_active=(not is_last), note=note_txt)
        try:
            await bot.edit_message_text(
                chat_id=chat_id,
                message_id=message_id,
                text=new_txt,
                parse_mode=ParseMode.HTML,
                reply_markup=_build_2fa_live_kb(tok_key, live=(not is_last)),
            )
        except Exception as e:
            if "message is not modified" not in str(e).lower():
                break

async def _run_live_outlook_countdown(bot, chat_id: int, message_id: int, raw_str: str, tok_key: str):
    """Auto-refreshes Outlook/Hotmail OTP every 5 seconds for 60 seconds unless stopped via Off button."""
    for step in range(12):  # 12 * 5s = 60 seconds
        if tok_key in _LIVE_TOOL_STOP:
            break
        await asyncio.sleep(5)
        if tok_key in _LIVE_TOOL_STOP:
            break
        res = await asyncio.to_thread(fetch_outlook_otp, raw_str)
        if tok_key in _LIVE_TOOL_STOP:
            break
        now_t = datetime.datetime.now().strftime("%I:%M:%S %p")
        is_last = (step == 11)
        status_note = (
            "ℹ️ <i>৬০ সেকেন্ড লাইভ সেশন শেষ — নতুন কোডের জন্য Refresh বাটনে চাপুন।</i>"
            if is_last
            else f"⏳ <i>প্রতি ৫ সেকেন্ডে অটো-আপডেট হচ্ছে ({now_t})</i>"
        )
        if res.get("ok"):
            new_txt = (
                f"✉️ <b>Outlook / Hotmail OTP Found!</b>\n"
                f"📧 <b>Email:</b> <code>{html.escape(str(res.get('email', '')))}</code>\n"
                f"📌 <b>Subject:</b> {html.escape(str(res.get('subject', 'Inbox')))}\n"
                f"🔢 <b>Code / Message:</b> <code>{html.escape(str(res.get('code', '')))}</code>\n"
                f"{status_note}"
            )
        else:
            new_txt = (
                f"❌ <b>Status:</b> {html.escape(str(res.get('error', 'Checking...')))}\n"
                f"{status_note}"
            )
        try:
            await bot.edit_message_text(
                chat_id=chat_id,
                message_id=message_id,
                text=new_txt,
                parse_mode=ParseMode.HTML,
                reply_markup=_build_mail_live_kb(tok_key, live=(not is_last)),
            )
        except Exception as e:
            if "message is not modified" not in str(e).lower():
                break

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

    all_joined, missing_fj = await check_all_force_joins(context.bot, user.id)
    if not all_joined:
        default_fj = (
            f"👋 <b>স্বাগতম, {html.escape(user.full_name or 'বন্ধু')}!</b>\n\n"
            f"বট ব্যবহার করতে নিচের সবগুলো গ্রুপ ও চ্যানেলে যুক্ত হয়ে ভেরিফাই করুন:"
        )
        force_txt = render_custom_bot_message("force_join", default_fj, name=html.escape(user.full_name or "বন্ধু"), shop_name=shop_name)
        kb = build_force_join_markup(missing_fj)
        await send_smart_bot_message(update.message, "force_join", force_txt, reply_markup=kb)
        return

    welcome_text = build_welcome_message(user.full_name, shop_name, is_new=is_new)
    reply_kb = get_main_reply_keyboard(user.id)
    slot_k = "welcome_new" if is_new else "welcome_back"
    await send_smart_bot_message(update.message, slot_k, welcome_text, reply_markup=reply_kb)

async def ai_assistant_menu_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user = update.effective_user
    set_user_state(user.id, "WAITING_AI_QUESTION", {})
    default_ai = (
        "🤖 <b>CyberNest AI Assistant</b>\n"
        "─────────────────────────────\n"
        "আসসালামু আলাইকুম! আপনার কোনো প্রশ্ন বা অসুবিধা থাকলে নিচে লিখে জিজ্ঞেস করুন—আমি সাথে সাথে সাহায্য করব।"
    )
    txt = render_custom_bot_message("ai_assistant", default_ai, name=html.escape(user.full_name or ""), shop_name=get_setting("shop_name", "CyberNest Market"))
    await send_smart_bot_message(update.message, "ai_assistant", txt)

async def _delete_admin_msg_safe(bot, msg_id):
    if not bot or not msg_id:
        return
    try:
        await bot.delete_message(chat_id=SUPER_ADMIN_ID, message_id=int(msg_id))
    except Exception:
        pass

def get_active_force_joins() -> List[Dict[str, Any]]:
    conn = get_db()
    c = conn.cursor()
    c.execute("SELECT id, title, chat_id, invite_link, is_active FROM force_joins WHERE is_active = 1 ORDER BY id ASC")
    rows = [dict(r) for r in c.fetchall()]
    conn.close()
    return rows

async def check_all_force_joins(bot, user_id: int) -> Tuple[bool, List[Dict[str, Any]]]:
    if user_id == SUPER_ADMIN_ID:
        return True, []
    if get_setting("force_join_enabled", "1") != "1":
        return True, []
    active_fjs = get_active_force_joins()
    if not active_fjs:
        return True, []
    missing = []
    for fj in active_fjs:
        cid = str(fj["chat_id"]).strip()
        try:
            member = await bot.get_chat_member(chat_id=cid, user_id=user_id)
            if member.status in ("left", "kicked"):
                missing.append(fj)
        except Exception:
            missing.append(fj)
    return (len(missing) == 0), missing

def build_force_join_markup(missing: List[Dict[str, Any]]) -> InlineKeyboardMarkup:
    buttons = []
    for fj in missing:
        lbl = fj.get("title") or "📢 Join Channel"
        url = fj.get("invite_link") or "https://t.me/"
        buttons.append([InlineKeyboardButton(lbl, url=url)])
    buttons.append([InlineKeyboardButton("✅ ভেরিফাই করুন", callback_data="btn_check_verification")])
    return colorize_markup(InlineKeyboardMarkup(buttons))

async def execute_smart_broadcast(bot, reply_target, sdata: dict, btn_text: Optional[str] = None, btn_url: Optional[str] = None, buttons_list: Optional[list] = None):
    msg_text = sdata.get("msg_text", "")
    photo_id = sdata.get("photo_id")
    reply_markup = None
    if buttons_list:
        reply_markup = colorize_markup(InlineKeyboardMarkup([[InlineKeyboardButton(b["text"], url=b["url"])] for b in buttons_list]))
    elif btn_text and btn_url:
        reply_markup = colorize_markup(InlineKeyboardMarkup([[InlineKeyboardButton(btn_text, url=btn_url)]]))

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

    default_prof = (
        f"👤 <b>My Profile</b> ({vip['badge']})\n\n"
        f"👤 <b>Name</b> : {db_user['name']}\n"
        f"🆔 <b>User ID</b> : <code>{user.id}</code>\n"
        f"💰 <b>Balance</b> : {balance_bdt:.2f} BDT / {balance_usdt:.4f} USDT\n"
        f"🛍️ <b>Total Spent</b> : {spent:.2f} BDT\n"
        f"🎉 <b>Joined</b> : {db_user['joined_date']}\n"
        f"─────────────────────────────"
    )
    profile_text = render_custom_bot_message(
        "my_profile",
        default_prof,
        name=html.escape(str(db_user["name"] or "")),
        user_id=user.id,
        balance=f"{balance_bdt:.2f}",
        balance_usdt=f"{balance_usdt:.4f}",
        total_spent=f"{spent:.2f}",
        joined_date=db_user["joined_date"],
        vip_badge=vip["badge"],
        shop_name=get_setting("shop_name", "CyberNest Market"),
    )
    inline_kb = colorize_markup(InlineKeyboardMarkup([
        [
            InlineKeyboardButton("💰 Balance Details", callback_data="prof_balance_details"),
            InlineKeyboardButton("🔄 Deposit History", callback_data="prof_deposit_history"),
        ],
        [
            InlineKeyboardButton("👥 Referral", callback_data="prof_referral"),
            InlineKeyboardButton("🛒 Purchases", callback_data="prof_purchases"),
        ]
    ]))
    await send_smart_bot_message(update.message, "my_profile", profile_text, reply_markup=inline_kb)

async def buy_products_menu_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    conn = get_db()
    c = conn.cursor()
    c.execute("SELECT id, name FROM categories ORDER BY position ASC, id ASC")
    cats = c.fetchall()
    conn.close()

    text = render_custom_bot_message(
        "buy_products",
        "🛍️ <b>Buy Products</b>\n\nSelect a category:",
        name=html.escape(update.effective_user.full_name or ""),
        shop_name=get_setting("shop_name", "CyberNest Market"),
    )
    keyboard = []
    for cat in cats:
        keyboard.append([InlineKeyboardButton(cat["name"], callback_data=f"usr_cat_{cat['id']}")])
    keyboard.append([InlineKeyboardButton("❌ Cancel", callback_data="usr_cancel_home")])

    await send_smart_bot_message(update.message, "buy_products", text, reply_markup=colorize_markup(InlineKeyboardMarkup(keyboard)))

async def get_code_menu_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    default_gc = (
        "🔐 <b>Get Code & Authentication Tools</b>\n\n"
        "হটমেইল/আউটলুক ওটিপি কোড এবং লাইভ 2FA কোড পাওয়ার জন্য নিচের অপশন থেকে নির্বাচন করুন:"
    )
    text = render_custom_bot_message(
        "get_code",
        default_gc,
        name=html.escape(update.effective_user.full_name or ""),
        shop_name=get_setting("shop_name", "CyberNest Market"),
    )
    kb = colorize_markup(InlineKeyboardMarkup([
        [InlineKeyboardButton("🔑 2FA Authenticator (Live Code)", callback_data="tool_2fa_prompt")],
        [InlineKeyboardButton("✉️ Outlook / Hotmail OTP Reader", callback_data="tool_outlook_prompt")],
        [InlineKeyboardButton("🌐 Hotmail & Outlook Code Server", url="https://code.yamin.bd/")],
    ]))
    await send_smart_bot_message(update.message, "get_code", text, reply_markup=kb)

async def support_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    default_sup = (
        "💬 <b>কাস্টমার সাপোর্ট</b>\n\n"
        "যেকোনো সমস্যায় নিচে এডমিনকে মেসেজ দিন অথবা 🤖 AI Assistant ব্যবহার করুন:"
    )
    text = render_custom_bot_message(
        "support",
        default_sup,
        name=html.escape(update.effective_user.full_name or ""),
        support=get_setting("support_username", "@Talha_juba098"),
        shop_name=get_setting("shop_name", "CyberNest Market"),
    )
    kb = colorize_markup(InlineKeyboardMarkup([
        [InlineKeyboardButton("📨 Message Admin", url=f"tg://user?id={SUPER_ADMIN_ID}")],
        [InlineKeyboardButton("🤖 Ask AI Assistant", callback_data="ai_q_open")],
    ]))
    await send_smart_bot_message(update.message, "support", text, reply_markup=kb)

async def deposit_menu_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user = update.effective_user
    db_user = get_or_create_user(user.id, user.full_name, user.username)
    min_man = float(get_setting("min_deposit_manual", "20.0"))
    set_user_state(user.id, "WAITING_DEP_AMOUNT_MANUAL", {})

    default_depm = (
        f"💳 <b>ডিপোজিট ব্যালেন্স (Send Money / Binance)</b>\n"
        f"─────────────────────────────\n"
        f"💰 বর্তমান ব্যালেন্স: <b>{db_user['balance_bdt']:.2f} BDT</b>\n"
        f"📌 সর্বনিম্ন ডিপোজিট: <b>{min_man:.0f} BDT</b>\n\n"
        f"আপনি কত টাকা ডিপোজিট করতে চান? টাকার পরিমাণ লিখে পাঠান (যেমন: 50, 100, 500):"
    )
    text = render_custom_bot_message(
        "deposit_menu",
        default_depm,
        name=html.escape(user.full_name or ""),
        balance=f"{db_user['balance_bdt']:.2f}",
        min_deposit=f"{min_man:.0f}",
        shop_name=get_setting("shop_name", "CyberNest Market"),
    )
    await send_smart_bot_message(update.message, "deposit_menu", text)

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
            all_joined, missing_fj = await check_all_force_joins(context.bot, user_id)
            if all_joined:
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
                await send_smart_bot_message(query.message, "welcome_new", welcome_text, reply_markup=reply_kb)
            else:
                try:
                    await query.answer("❌ আপনি এখনো সবগুলো গ্রুপ বা চ্যানেলে জয়েন করেননি! দয়া করে জয়েন করে আবার চাপুন।", show_alert=True)
                except Exception:
                    pass
                kb = build_force_join_markup(missing_fj)
                try:
                    await query.edit_message_reply_markup(reply_markup=kb)
                except Exception:
                    pass
        except Exception as e:
            logger.error(f"Verification button error: {e}", exc_info=True)
        return

    if data.startswith("ai_q_"):
        q_type = data.replace("ai_q_", "")
        if q_type == "exit":
            set_user_state(user_id, None)
            await query.message.reply_text("✅ AI Assistant মোড বন্ধ করা হয়েছে।")
            return
        if q_type == "open":
            set_user_state(user_id, "WAITING_AI_QUESTION", {})
            ai_open_txt = render_custom_bot_message(
                "ai_assistant",
                "🤖 আসসালামু আলাইকুম! আপনার কোনো প্রশ্ন বা অসুবিধা থাকলে নিচে লিখে জিজ্ঞেস করুন।",
                name=html.escape(query.from_user.full_name or ""),
                shop_name=get_setting("shop_name", "CyberNest Market"),
            )
            await send_smart_bot_message(query.message, "ai_assistant", ai_open_txt)
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
        markup = colorize_markup(InlineKeyboardMarkup(kb))
        if getattr(query.message, "photo", None):
            try:
                await query.message.delete()
            except Exception:
                pass
            await query.message.reply_text(text, parse_mode=ParseMode.HTML, reply_markup=markup)
        else:
            await query.edit_message_text(text, parse_mode=ParseMode.HTML, reply_markup=markup)
        return

    if data == "usr_cancel_home":
        cancel_txt = render_custom_bot_message(
            "order_cancel",
            "❌ কেনাকাটা বাতিল করা হয়েছে। নতুন অর্ডারের জন্য মেনু ব্যবহার করুন।",
            name=html.escape(query.from_user.full_name or ""),
            shop_name=get_setting("shop_name", "CyberNest Market"),
        )
        await query.message.reply_text(cancel_txt, parse_mode=ParseMode.HTML)
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
        text = render_custom_bot_message(
            "buy_products",
            "🛍️ <b>Buy Products</b>\n\nSelect a category:",
            name=html.escape(query.from_user.full_name or ""),
            shop_name=get_setting("shop_name", "CyberNest Market"),
        )
        keyboard = []
        for cat in cats:
            keyboard.append([InlineKeyboardButton(cat["name"], callback_data=f"usr_cat_{cat['id']}")])
        keyboard.append([InlineKeyboardButton("❌ Cancel", callback_data="usr_cancel_home")])
        markup = colorize_markup(InlineKeyboardMarkup(keyboard))
        if getattr(query.message, "photo", None):
            try:
                await query.message.delete()
            except Exception:
                pass
            await send_smart_bot_message(query.message, "buy_products", text, reply_markup=markup)
        else:
            await query.edit_message_text(text, parse_mode=ParseMode.HTML, reply_markup=markup)
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
        if getattr(query.message, "photo", None):
            try:
                await query.message.delete()
            except Exception:
                pass
            await query.message.reply_text(text, parse_mode=ParseMode.HTML, reply_markup=colorize_markup(InlineKeyboardMarkup(kb)))
        else:
            await query.edit_message_text(text, parse_mode=ParseMode.HTML, reply_markup=colorize_markup(InlineKeyboardMarkup(kb)))
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

    if data.startswith("calc_noop_"):
        return

    if data.startswith("calc_confirm_"):
        all_joined, missing_fj = await check_all_force_joins(context.bot, user_id)
        if not all_joined:
            kb = build_force_join_markup(missing_fj)
            fj_calc_txt = render_custom_bot_message(
                "force_join",
                "⚠️ কিনতে হলে আগে নিচের সবগুলো গ্রুপ ও চ্যানেলে জয়েন করে ভেরিফাই করুন:",
                name=html.escape(query.from_user.full_name or ""),
                shop_name=get_setting("shop_name", "CyberNest Market"),
            )
            await send_smart_bot_message(
                query.message,
                "force_join",
                fj_calc_txt,
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
        kb = colorize_markup(InlineKeyboardMarkup([
            [InlineKeyboardButton("📱 bKash", callback_data="usr_dep_method_bKash"), InlineKeyboardButton("📱 Nagad", callback_data="usr_dep_method_Nagad")],
            [InlineKeyboardButton("💎 Binance", callback_data="usr_dep_method_Binance")],
        ]))
        _, sdata_dm = get_user_state(user_id)
        dm_amt = float(sdata_dm.get("amount", 0) or 0)
        dm_txt = render_custom_bot_message(
            "deposit_method",
            "নিচের কোন মাধ্যমে সেন্ড মানি করবেন সিলেক্ট করুন:",
            amount=f"{dm_amt:.2f}",
            name=html.escape(query.from_user.full_name or ""),
            shop_name=get_setting("shop_name", "CyberNest Market"),
        )
        if getattr(query.message, "photo", None):
            try:
                await query.message.delete()
            except Exception:
                pass
            await send_smart_bot_message(query.message, "deposit_method", dm_txt, reply_markup=kb)
        else:
            await query.edit_message_text(dm_txt, parse_mode=ParseMode.HTML, reply_markup=kb)
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

        default_instr = (
            f"📱 <b>{method} পেমেন্ট নির্দেশিকা:</b>\n"
            f"─────────────────────────────\n"
            f"💰 ডিপোজিট পরিমাণ: <b>{amount:.2f} BDT</b>\n"
            f"📌 পেমেন্ট নাম্বার / আইডি (ক্লিক করে কপি করুন):\n"
            f"<code>{num}</code>\n\n"
            f"⚠️ <i>উপরের নাম্বারে <b>{amount:.2f} BDT</b> Send Money করার পর নিচে আপনার Transaction ID (TrxID) টি লিখে পাঠান:</i>"
        )
        text = render_custom_bot_message(
            "deposit_instr",
            default_instr,
            method=method,
            amount=f"{amount:.2f}",
            number=num,
            name=html.escape(query.from_user.full_name or ""),
            shop_name=get_setting("shop_name", "CyberNest Market"),
        )
        if get_setting("msg_photo_deposit_instr", "").strip() or getattr(query.message, "photo", None):
            try:
                await query.message.delete()
            except Exception:
                pass
            await send_smart_bot_message(query.message, "deposit_instr", text)
        else:
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

    if data.startswith("tool_stop_mail_"):
        tok_key = data.replace("tool_stop_mail_", "", 1)
        _LIVE_TOOL_STOP.add(tok_key)
        cur_txt = query.message.text_html or query.message.text or "✉️ Outlook / Hotmail OTP"
        if "অটো-আপডেট বন্ধ করা হয়েছে" not in cur_txt:
            cur_txt += "\n⏹️ <i>অটো-আপডেট বন্ধ করা হয়েছে।</i>"
        try:
            await query.edit_message_text(cur_txt, parse_mode=ParseMode.HTML, reply_markup=_build_mail_live_kb(tok_key, live=False))
        except Exception:
            try:
                await query.edit_message_reply_markup(reply_markup=_build_mail_live_kb(tok_key, live=False))
            except Exception:
                pass
        return

    if data.startswith("tool_ref_mail_"):
        tok_key = data.replace("tool_ref_mail_", "", 1)
        _LIVE_TOOL_STOP.discard(tok_key)
        raw_str = _TOOL_CACHE.get(tok_key)
        if not raw_str:
            await query.message.reply_text("⚠️ সেশন শেষ হয়ে গেছে। অনুগ্রহ করে মেইল স্ট্রিংটি আবার পেস্ট করুন।")
            return
        res = await asyncio.to_thread(fetch_outlook_otp, raw_str)
        now_t = datetime.datetime.now().strftime("%I:%M:%S %p")
        kb = _build_mail_live_kb(tok_key, live=True)
        if res.get("ok"):
            new_txt = (
                f"✉️ <b>Outlook / Hotmail OTP Found!</b>\n"
                f"📧 <b>Email:</b> <code>{html.escape(str(res.get('email', '')))}</code>\n"
                f"📌 <b>Subject:</b> {html.escape(str(res.get('subject', 'Inbox')))}\n"
                f"🔢 <b>Code / Message:</b> <code>{html.escape(str(res.get('code', '')))}</code>\n"
                f"⏳ <i>প্রতি ৫ সেকেন্ডে অটো-আপডেট হচ্ছে ({now_t})</i>"
            )
        else:
            new_txt = (
                f"❌ <b>Refresh Error:</b> {html.escape(str(res.get('error', 'Failed')))}\n"
                f"⏳ <i>প্রতি ৫ সেকেন্ডে অটো-আপডেট হচ্ছে ({now_t})</i>"
            )
        try:
            await query.edit_message_text(new_txt, parse_mode=ParseMode.HTML, reply_markup=kb)
            asyncio.create_task(_run_live_outlook_countdown(context.bot, query.message.chat_id, query.message.message_id, raw_str, tok_key))
        except Exception:
            pass
        return

    if data.startswith("tool_stop_2fa_"):
        tok_key = data.replace("tool_stop_2fa_", "", 1)
        _LIVE_TOOL_STOP.add(tok_key)
        raw_sec = _TOOL_CACHE.get(tok_key)
        parsed_codes = generate_universal_2fa_codes(raw_sec) if raw_sec else None
        if parsed_codes:
            new_txt = _format_2fa_live_text(parsed_codes, live_active=False, note="⏹️ <i>অটো-আপডেট বন্ধ করা হয়েছে।</i>")
        else:
            new_txt = (query.message.text_html or "🔑 2FA Authenticator") + "\n⏹️ <i>অটো-আপডেট বন্ধ করা হয়েছে।</i>"
        try:
            await query.edit_message_text(new_txt, parse_mode=ParseMode.HTML, reply_markup=_build_2fa_live_kb(tok_key, live=False))
        except Exception:
            try:
                await query.edit_message_reply_markup(reply_markup=_build_2fa_live_kb(tok_key, live=False))
            except Exception:
                pass
        return

    if data.startswith("tool_ref_2fa_"):
        tok_key = data.replace("tool_ref_2fa_", "", 1)
        _LIVE_TOOL_STOP.discard(tok_key)
        raw_sec = _TOOL_CACHE.get(tok_key)
        if not raw_sec:
            await query.message.reply_text("⚠️ সেশন শেষ হয়ে গেছে। অনুগ্রহ করে 2FA Secret Key আবার দিন।")
            return
        parsed_codes = generate_universal_2fa_codes(raw_sec)
        if parsed_codes:
            now_t = datetime.datetime.now().strftime("%I:%M:%S %p")
            kb = _build_2fa_live_kb(tok_key, live=True)
            new_txt = _format_2fa_live_text(parsed_codes, live_active=True, note=f"🕒 <i>Refreshed at: {now_t}</i>")
            try:
                await query.edit_message_text(new_txt, parse_mode=ParseMode.HTML, reply_markup=kb)
                asyncio.create_task(_run_live_2fa_countdown(context.bot, query.message.chat_id, query.message.message_id, raw_sec, tok_key))
            except Exception:
                pass
        return

    if data == "prof_balance_details":
        rate = float(get_setting("usdt_rate", "125.0") or 125.0)
        bal_bdt = float(db_user.get("balance_bdt", 0.0) or 0.0)
        bal_usdt = bal_bdt / rate if rate > 0 else 0.0
        conn = get_db()
        c = conn.cursor()
        c.execute("SELECT COALESCE(SUM(amount), 0) as tot_dep FROM deposits WHERE user_id = ? AND status = 'approved'", (user_id,))
        tot_dep = float(c.fetchone()["tot_dep"] or 0.0)
        c.execute("SELECT COALESCE(SUM(total_bdt), 0) as tot_spent FROM orders WHERE user_id = ? AND status != 'refunded'", (user_id,))
        tot_spent = float(c.fetchone()["tot_spent"] or 0.0)
        conn.close()
        txt = (
            f"💰 <b>Balance Details</b>\n"
            f"─────────────────────────────\n"
            f"💵 <b>Current Balance:</b> {bal_bdt:.2f} BDT ({bal_usdt:.4f} USDT)\n"
            f"📥 <b>Total Deposited:</b> {tot_dep:.2f} BDT\n"
            f"🛍️ <b>Total Spent:</b> {tot_spent:.2f} BDT\n"
            f"💱 <b>Exchange Rate:</b> 1 USDT = {rate:.2f} BDT"
        )
        await query.message.reply_text(txt, parse_mode=ParseMode.HTML)
        return

    if data == "prof_deposit_history":
        conn = get_db()
        c = conn.cursor()
        c.execute("SELECT method, amount, trx_id, status, created_at FROM deposits WHERE user_id = ? ORDER BY id DESC LIMIT 10", (user_id,))
        deps = c.fetchall()
        conn.close()
        if not deps:
            await query.message.reply_text("🔄 আপনার এখনো কোনো ডিপোজিট হিস্টোরি নেই।")
            return
        lines = ["🔄 <b>সর্বশেষ ডিপোজিট হিস্টোরি:</b>", "─────────────────────────────"]
        for d in deps:
            st_icon = "✅" if d["status"] == "approved" else ("❌" if d["status"] == "rejected" else "⏳")
            lines.append(f"{st_icon} <b>{d['amount']:.2f} BDT</b> ({d['method']}) — <code>{d['trx_id']}</code>\n    <i>{d['created_at']} • {d['status'].upper()}</i>")
        await query.message.reply_text("\n".join(lines), parse_mode=ParseMode.HTML)
        return

    if data == "prof_referral":
        bot_info = await context.bot.get_me()
        ref_link = f"https://t.me/{bot_info.username}?start=ref_{user_id}"
        conn = get_db()
        c = conn.cursor()
        c.execute("SELECT COUNT(*) as cnt FROM users WHERE referred_by = ?", (user_id,))
        rcnt = c.fetchone()["cnt"]
        conn.close()
        earn = float(db_user.get("referral_earnings", 0.0) or 0.0)
        txt = (
            f"👥 <b>Referral Program</b>\n"
            f"─────────────────────────────\n"
            f"👤 <b>Total Referrals:</b> {rcnt} জন\n"
            f"💰 <b>Total Earnings:</b> {earn:.2f} BDT\n\n"
            f"🔗 <b>Your Invite Link (ট্যাপ করে কপি করুন):</b>\n"
            f"<code>{ref_link}</code>"
        )
        await query.message.reply_text(txt, parse_mode=ParseMode.HTML)
        return

    if data == "prof_purchases":
        conn = get_db()
        c = conn.cursor()
        c.execute("SELECT id, receipt_no, product_name, quantity, total_bdt, status, order_date FROM orders WHERE user_id = ? ORDER BY id DESC LIMIT 10", (user_id,))
        ords = c.fetchall()
        conn.close()
        if not ords:
            await query.message.reply_text("🛒 আপনার এখনো কোনো অর্ডার হিস্টোরি নেই।")
            return
        kb = []
        for o in ords:
            st_ic = "✅" if o["status"] == "completed" else ("⏳" if o["status"] == "pending_delivery" else "❌")
            kb.append([InlineKeyboardButton(f"{st_ic} {o['product_name']} (x{o['quantity']}) — {o['total_bdt']:.0f} Tk", callback_data=f"prof_view_order_{o['id']}")])
        await query.message.reply_text("🛒 <b>আপনার সর্বশেষ অর্ডারসমূহ:</b>\nযেকোনো অর্ডারে ক্লিক করে বিস্তারিত ও প্রোডাক্ট দেখুন:", parse_mode=ParseMode.HTML, reply_markup=colorize_markup(InlineKeyboardMarkup(kb)))
        return

    if data.startswith("prof_view_order_"):
        oid = int(data.split("_")[3])
        conn = get_db()
        c = conn.cursor()
        c.execute("SELECT * FROM orders WHERE id = ? AND user_id = ?", (oid, user_id))
        o = c.fetchone()
        conn.close()
        if not o:
            await query.message.reply_text("❌ অর্ডার পাওয়া যায়নি।")
            return
        if o["delivered_data"]:
            raw_deliv = o["delivered_data"]
            try:
                items = json.loads(raw_deliv) if raw_deliv.startswith("[") else [ln for ln in raw_deliv.splitlines() if ln.strip()]
            except Exception:
                items = [ln for ln in raw_deliv.splitlines() if ln.strip()]
            card = format_credentials_card(items, o["product_name"], o["category_name"] or "Service")
            await query.message.reply_text(f"🧾 <b>Receipt:</b> <code>{o['receipt_no']}</code>\n\n{card}", parse_mode=ParseMode.HTML)
        else:
            await query.message.reply_text(
                f"🧾 <b>Receipt:</b> <code>{o['receipt_no']}</code>\n"
                f"📦 <b>Product:</b> {o['product_name']} (x{o['quantity']})\n"
                f"💰 <b>Total:</b> {o['total_bdt']:.2f} BDT\n"
                f"📌 <b>Status:</b> {o['status'].upper()}",
                parse_mode=ParseMode.HTML,
            )
        return

    if data.startswith("adm_"):
        if is_admin(user_id):
            await handle_admin_callbacks(query, data, context)
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
    kb = colorize_markup(InlineKeyboardMarkup([
        [InlineKeyboardButton("🛒 Buy Products", url=f"https://t.me/{bot_info.username}")]
    ]))
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

    default_ps = (
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
    text = render_custom_bot_message(
        "purchase_summary",
        default_ps,
        category=p["cat_name"],
        package=full_name,
        description=desc_show,
        quantity=quantity,
        rate=format_price(rate),
        total=format_price(total),
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
    reply_markup = colorize_markup(InlineKeyboardMarkup(kb))

    if p["image_file_id"]:
        photo_src = p["image_file_id"]
        if isinstance(photo_src, str) and photo_src.startswith("data:image") and "," in photo_src:
            try:
                b64_part = photo_src.split(",", 1)[1]
                photo_src = InputFile(io.BytesIO(base64.b64decode(b64_part)), filename="product.jpg")
            except Exception:
                photo_src = None
        if photo_src and edit and hasattr(query_or_message, "message") and query_or_message.message:
            if getattr(query_or_message.message, "photo", None):
                try:
                    await query_or_message.edit_message_caption(caption=text, parse_mode=ParseMode.HTML, reply_markup=reply_markup)
                    return
                except Exception:
                    pass
            else:
                try:
                    await query_or_message.message.delete()
                except Exception:
                    pass
                try:
                    await query_or_message.message.reply_photo(photo=photo_src, caption=text, parse_mode=ParseMode.HTML, reply_markup=reply_markup)
                    return
                except Exception:
                    pass
        try:
            target = query_or_message.message if hasattr(query_or_message, "message") else query_or_message
            await target.reply_photo(photo=photo_src, caption=text, parse_mode=ParseMode.HTML, reply_markup=reply_markup)
            return
        except Exception:
            pass

    if edit and hasattr(query_or_message, "edit_message_text"):
        try:
            await query_or_message.edit_message_text(text, parse_mode=ParseMode.HTML, reply_markup=reply_markup)
            return
        except Exception:
            pass

    target = query_or_message.message if hasattr(query_or_message, "message") else query_or_message
    await target.reply_text(text, parse_mode=ParseMode.HTML, reply_markup=reply_markup)

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

        adm_kb = colorize_markup(InlineKeyboardMarkup([
            [InlineKeyboardButton(f"🚀 Deliver Order #{order_id}", callback_data=f"adm_fulfill_{order_id}")],
            [InlineKeyboardButton(f"❌ Cancel & Refund #{order_id}", callback_data=f"adm_refund_{order_id}")]
        ]))
        try:
            sent_ord = await get_bot(query).send_message(
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
            _c4 = get_db()
            _c4.cursor().execute("UPDATE orders SET admin_msg_id = ? WHERE id = ?", (sent_ord.message_id, order_id))
            _c4.commit()
            _c4.close()
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
    return colorize_markup(InlineKeyboardMarkup([
        [
            InlineKeyboardButton("📥 Add Stock", callback_data="adm_stk_step1_cats"),
            InlineKeyboardButton("🏷️ Manage Buttons & Prices", callback_data="adm_btn_step1_cats"),
        ],
        [
            InlineKeyboardButton("🖼️ Manage Product Photos", callback_data="adm_pho_step1_cats"),
            InlineKeyboardButton("🖼️ Message Photo Manager", callback_data="adm_msg_photos"),
        ],
        [
            InlineKeyboardButton("📝 Message Text Manager", callback_data="adm_msg_texts"),
        ],
        [
            InlineKeyboardButton("💳 Payment Settings", callback_data="adm_pay_settings"),
            InlineKeyboardButton(f"💳 Deposits ({pend_dep})", callback_data="adm_pending_deposits"),
        ],
        [
            InlineKeyboardButton(f"🛡️ Orders ({pend_ord})", callback_data="adm_pending_orders"),
            InlineKeyboardButton("👥 User Management", callback_data="adm_usr_list_0"),
        ],
        [
            InlineKeyboardButton("📊 Shop Statistics", callback_data="adm_statistics"),
            InlineKeyboardButton("🧾 Check / Download Receipt", callback_data="adm_check_receipt"),
        ],
        [
            InlineKeyboardButton("📢 Broadcast (with Button)", callback_data="adm_broadcast_start"),
            InlineKeyboardButton(status_tx, callback_data="adm_toggle_autosales"),
        ],
        [
            InlineKeyboardButton("⚡ Channel TX Manager", callback_data="adm_tx_manager"),
            InlineKeyboardButton("🎟️ Promo / Coupon Manager", callback_data="adm_promo_mgr"),
        ],
        [
            InlineKeyboardButton("📢 Force Join Manager", callback_data="adm_fj_mgr"),
            InlineKeyboardButton("🎛️ Menu Layout Manager", callback_data="adm_menu_layout"),
        ],
        [
            InlineKeyboardButton("🎨 Button Color Manager", callback_data="adm_btn_colors"),
            InlineKeyboardButton("🔄 Factory Reset / Clear Data", callback_data="adm_factory_reset_warn"),
        ],
        [
            make_inline_btn("❌ Close Admin Panel", style="danger", callback_data="adm_close_panel"),
        ]
    ]), group="admin_panel")

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

    if data == "adm_close_panel":
        try:
            await query.message.delete()
        except Exception:
            try:
                await query.edit_message_text("✅ Admin Panel বন্ধ করা হয়েছে।")
            except Exception:
                pass
        return

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
        tx_kb = colorize_markup(InlineKeyboardMarkup([
            [InlineKeyboardButton(st_label, callback_data="adm_toggle_autosales_mgr")],
            [InlineKeyboardButton("✍️ Custom Transaction (নিজে বানান)", callback_data="adm_custom_tx_start")],
            [InlineKeyboardButton("⚙️ Group & Channel ID/Link Settings", callback_data="adm_gc_settings")],
            [InlineKeyboardButton("< Back to Dashboard", callback_data="adm_root")],
        ]))
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
        gc_kb = colorize_markup(InlineKeyboardMarkup([
            [
                InlineKeyboardButton("✏️ Edit Group ID", callback_data="adm_set_wgrp_id"),
                InlineKeyboardButton("🔗 Edit Group Link", callback_data="adm_set_wgrp_link"),
            ],
            [
                InlineKeyboardButton("✏️ Edit Channel ID", callback_data="adm_set_txch_id"),
                InlineKeyboardButton("🔗 Edit Channel Link", callback_data="adm_set_txch_link"),
            ],
            [InlineKeyboardButton("< Back to TX Manager", callback_data="adm_tx_manager")],
        ]))
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
        r_kb = colorize_markup(InlineKeyboardMarkup([
            [InlineKeyboardButton("🔴 Yes, Wipe All Data", callback_data="adm_factory_reset_confirm")],
            [InlineKeyboardButton("❌ Cancel", callback_data="adm_root")]
        ]))
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

    # --- 🖼️ MESSAGE PHOTO MANAGER (ATTACH PHOTOS TO ANY BOT MESSAGE) ---
    if data == "adm_msg_photos":
        set_cnt = sum(1 for sk, _ in BOT_MESSAGE_SLOTS if get_setting(f"msg_photo_{sk}", "").strip())
        txt = (
            "🖼️ <b>Message Photo Manager (মেসেজ ফটো ম্যানেজার)</b>\n"
            "─────────────────────────────\n"
            f"📸 মোট ছবি যুক্ত করা মেসেজ: <b>{set_cnt} / {len(BOT_MESSAGE_SLOTS)}</b> টি\n\n"
            "বটের যেকোনো মেসেজের সাথে ছবি (Photo) যুক্ত, পরিবর্তন বা রিমুভ করতে নিচ থেকে মেসেজটি সিলেক্ট করুন:"
        )
        kb = []
        for sk, slabel in BOT_MESSAGE_SLOTS:
            has_img = bool(get_setting(f"msg_photo_{sk}", "").strip())
            icon = "🖼️" if has_img else "⚪"
            kb.append([InlineKeyboardButton(f"{icon} {slabel}", callback_data=f"adm_mphoto_view_{sk}")])
        kb.append([InlineKeyboardButton("< Back to Dashboard", callback_data="adm_root")])
        await query.edit_message_text(txt, parse_mode=ParseMode.HTML, reply_markup=colorize_markup(InlineKeyboardMarkup(kb)))
        return

    if data.startswith("adm_mphoto_view_"):
        sk = data.replace("adm_mphoto_view_", "", 1)
        slabel = BOT_MESSAGE_SLOT_MAP.get(sk, sk)
        cur_photo = get_setting(f"msg_photo_{sk}", "").strip()
        status_txt = "✅ <b>ছবি যুক্ত করা আছে (Active)</b>" if cur_photo else "⚪ <b>কোনো ছবি যুক্ত নেই (শুধু টেক্সট যাবে)</b>"
        kb = [[InlineKeyboardButton("📸 Upload / Change Photo", callback_data=f"adm_mphoto_set_{sk}")]]
        if cur_photo:
            kb.append([
                InlineKeyboardButton("👁️ Preview Photo", callback_data=f"adm_mphoto_prev_{sk}"),
                InlineKeyboardButton("🗑️ Remove Photo", callback_data=f"adm_mphoto_del_{sk}"),
            ])
        kb.append([InlineKeyboardButton("< Back to Message Photos", callback_data="adm_msg_photos")])
        await query.edit_message_text(
            f"🖼️ <b>মেসেজ সেকশন:</b> {html.escape(slabel)}\n"
            f"─────────────────────────────\n"
            f"📌 বর্তমান অবস্থা: {status_txt}\n\n"
            f"ছবি সেট করলে ইউজার যখনই এই মেসেজে আসবে, টেক্সট মেসেজটি ছবির ক্যাপশন হিসেবে যাবে।",
            parse_mode=ParseMode.HTML,
            reply_markup=colorize_markup(InlineKeyboardMarkup(kb)),
        )
        return

    if data.startswith("adm_mphoto_set_"):
        sk = data.replace("adm_mphoto_set_", "", 1)
        slabel = BOT_MESSAGE_SLOT_MAP.get(sk, sk)
        set_user_state(user_id, "ADMIN_WAITING_MSG_PHOTO", {"slot_key": sk})
        await query.message.reply_text(
            f"📸 <b>{html.escape(slabel)}</b>-এর জন্য ছবিটি টেলিগ্রাম চ্যাটে <b>Photo</b> আকারে পাঠান (অথবা Photo File ID / Direct Image URL পেস্ট করুন):",
            parse_mode=ParseMode.HTML,
        )
        return

    if data.startswith("adm_mphoto_prev_"):
        sk = data.replace("adm_mphoto_prev_", "", 1)
        slabel = BOT_MESSAGE_SLOT_MAP.get(sk, sk)
        cur_photo = get_setting(f"msg_photo_{sk}", "").strip()
        if cur_photo:
            try:
                await query.message.reply_photo(
                    photo=cur_photo,
                    caption=f"👁️ <b>Preview:</b> {html.escape(slabel)}",
                    parse_mode=ParseMode.HTML,
                )
            except Exception as e:
                await query.message.reply_text(f"❌ ছবিটি লোড করা যায়নি: {e}")
        return

    if data.startswith("adm_mphoto_del_"):
        sk = data.replace("adm_mphoto_del_", "", 1)
        set_setting(f"msg_photo_{sk}", "")
        await handle_admin_callbacks(query, f"adm_mphoto_view_{sk}", context)
        return

    # --- 📝 MESSAGE TEXT MANAGER (EDIT ANY BOT MESSAGE TEXT) ---
    if data == "adm_msg_texts":
        custom_cnt = sum(1 for sk, _ in BOT_MESSAGE_SLOTS if get_setting(f"msg_text_{sk}", "").strip())
        txt = (
            "📝 <b>Message Text Manager (মেসেজ টেক্সট এডিটর)</b>\n"
            "─────────────────────────────\n"
            f"✏️ কাস্টমাইজ করা মেসেজ: <b>{custom_cnt} / {len(BOT_MESSAGE_SLOTS)}</b> টি\n\n"
            "বটের যেকোনো বাটন বা মেনুর মেসেজ টেক্সট এডিট করতে বা ডিফল্টে রিসেট করতে নিচ থেকে মেসেজটি সিলেক্ট করুন:"
        )
        kb = []
        for sk, slabel in BOT_MESSAGE_SLOTS:
            has_custom = bool(get_setting(f"msg_text_{sk}", "").strip())
            icon = "✏️" if has_custom else "⚪"
            kb.append([InlineKeyboardButton(f"{icon} {slabel}", callback_data=f"adm_mtext_view_{sk}")])
        kb.append([InlineKeyboardButton("< Back to Dashboard", callback_data="adm_root")])
        await query.edit_message_text(txt, parse_mode=ParseMode.HTML, reply_markup=colorize_markup(InlineKeyboardMarkup(kb)))
        return

    if data.startswith("adm_mtext_view_"):
        sk = data.replace("adm_mtext_view_", "", 1)
        slabel = BOT_MESSAGE_SLOT_MAP.get(sk, sk)
        cur_txt = get_setting(f"msg_text_{sk}", "").strip()
        ph_hint = BOT_MESSAGE_PLACEHOLDERS.get(sk, "{name}, {shop_name}")
        status_txt = "✅ <b>কাস্টম টেক্সট সেট করা আছে</b>" if cur_txt else "⚪ <b>ডিফল্ট মেসেজ চালু আছে</b>"
        preview_block = (
            f"\n\n📄 <b>বর্তমান কাস্টম টেক্সট:</b>\n<pre>{html.escape(cur_txt[:800])}</pre>"
            if cur_txt
            else ""
        )
        kb = [[InlineKeyboardButton("✏️ Edit Message Text", callback_data=f"adm_mtext_set_{sk}")]]
        if cur_txt:
            kb.append([InlineKeyboardButton("🔄 Reset to Default Text", callback_data=f"adm_mtext_reset_{sk}")])
        kb.append([InlineKeyboardButton("< Back to Message Texts", callback_data="adm_msg_texts")])
        await query.edit_message_text(
            f"📝 <b>মেসেজ সেকশন:</b> {html.escape(slabel)}\n"
            f"─────────────────────────────\n"
            f"📌 বর্তমান অবস্থা: {status_txt}\n"
            f"🧩 <b>ব্যবহারযোগ্য ভ্যারিয়েবল (Optional):</b> <code>{html.escape(ph_hint)}</code>"
            f"{preview_block}\n\n"
            f"💡 <i>আপনি চাইলে &lt;b&gt;Bold&lt;/b&gt;, &lt;i&gt;Italic&lt;/i&gt;, &lt;code&gt;Code&lt;/code&gt; HTML ট্যাগও ব্যবহার করতে পারবেন।</i>",
            parse_mode=ParseMode.HTML,
            reply_markup=colorize_markup(InlineKeyboardMarkup(kb)),
        )
        return

    if data.startswith("adm_mtext_set_"):
        sk = data.replace("adm_mtext_set_", "", 1)
        slabel = BOT_MESSAGE_SLOT_MAP.get(sk, sk)
        ph_hint = BOT_MESSAGE_PLACEHOLDERS.get(sk, "{name}, {shop_name}")
        set_user_state(user_id, "ADMIN_WAITING_MSG_TEXT", {"slot_key": sk})
        await query.message.reply_text(
            f"✏️ <b>{html.escape(slabel)}</b>-এর জন্য নতুন মেসেজটি লিখে পাঠান:\n\n"
            f"🧩 <b>ভ্যারিয়েবল (প্রয়োজন হলে ব্যবহার করুন):</b> <code>{html.escape(ph_hint)}</code>\n"
            f"💡 <i>HTML ফরম্যাট (&lt;b&gt;, &lt;i&gt;, &lt;code&gt;) সাপোর্ট করে।</i>",
            parse_mode=ParseMode.HTML,
        )
        return

    if data.startswith("adm_mtext_reset_"):
        sk = data.replace("adm_mtext_reset_", "", 1)
        set_setting(f"msg_text_{sk}", "")
        await handle_admin_callbacks(query, f"adm_mtext_view_{sk}", context)
        return

    if data == "adm_stk_step1_cats":
        conn = get_db()
        c = conn.cursor()
        c.execute("SELECT id, name FROM categories ORDER BY position ASC, id ASC")
        cats = c.fetchall()
        conn.close()
        kb = [[InlineKeyboardButton(cat["name"], callback_data=f"adm_stk_step2_sub_{cat['id']}")] for cat in cats]
        kb.append([InlineKeyboardButton("< Back to Dashboard", callback_data="adm_root")])
        await query.edit_message_text("📥 <b>[Add Stock — ধাপ ১]</b>\nক্যাটাগরি সিলেক্ট করুন:", parse_mode=ParseMode.HTML, reply_markup=colorize_markup(InlineKeyboardMarkup(kb)))
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
        await query.edit_message_text(f"📥 <b>[Add Stock — ধাপ ২]</b>\n<b>{cat_name}</b> ➔ ব্র্যান্ড সিলেক্ট করুন:", parse_mode=ParseMode.HTML, reply_markup=colorize_markup(InlineKeyboardMarkup(kb)))
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
        await query.edit_message_text(f"📥 <b>[Add Stock — ধাপ ৩]</b>\n<b>{title}</b> ➔ প্যাকেজ সিলেক্ট করুন:", parse_mode=ParseMode.HTML, reply_markup=colorize_markup(InlineKeyboardMarkup(kb)))
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
        await query.edit_message_text("🏷️ <b>[বাটন ও প্রোডাক্ট পরিচালনা — ধাপ ১]</b>", parse_mode=ParseMode.HTML, reply_markup=colorize_markup(InlineKeyboardMarkup(kb)))
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
        await query.edit_message_text(f"🏷️ <b>{cat_name}</b> ➔ ব্র্যান্ড বেছে নিন:", parse_mode=ParseMode.HTML, reply_markup=colorize_markup(InlineKeyboardMarkup(kb)))
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
        await query.edit_message_text(f"🏷️ <b>{title}</b>", parse_mode=ParseMode.HTML, reply_markup=colorize_markup(InlineKeyboardMarkup(kb)))
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
        await query.edit_message_text(txt, parse_mode=ParseMode.HTML, reply_markup=colorize_markup(InlineKeyboardMarkup(kb)))
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

    if data == "adm_btn_colors":
        cfg = get_btn_color_config()
        c_map = {"primary": "🔵 Blue", "success": "🟢 Green", "danger": "🔴 Red", "none": "⚪ Default"}
        ov_cnt = len(cfg.get("overrides") or {})
        txt = (
            "🎨 <b>Button Color Manager (বটের সব বাটনের কালার)</b>\n"
            "─────────────────────────────\n"
            f"📱 <b>মেইন মেনু বাটন:</b> {c_map.get(cfg.get('main_menu', 'primary'))}\n"
            f"👑 <b>এডমিন প্যানেল বাটন:</b> {c_map.get(cfg.get('admin_panel', 'primary'))}\n"
            f"🔹 <b>সাধারণ ইনলাইন বাটন:</b> {c_map.get(cfg.get('inline_default', 'primary'))}\n"
            f"✅ <b>কনফার্ম/অ্যাকশন বাটন:</b> {c_map.get(cfg.get('inline_success', 'success'))}\n"
            f"❌ <b>ক্যানসেল/ডিলিট/অফ বাটন:</b> {c_map.get(cfg.get('inline_danger', 'danger'))}\n"
            f"🎯 <b>নির্দিষ্ট বাটন ওভাররাইড:</b> {ov_cnt} টি\n\n"
            "যেকোনো সেকশন বা নির্দিষ্ট বাটনের কালার পরিবর্তন করতে নিচে চাপুন:"
        )
        kb = colorize_markup(InlineKeyboardMarkup([
            [InlineKeyboardButton("📱 মেইন মেনু বাটন কালার", callback_data="adm_bcol_grp_main_menu")],
            [InlineKeyboardButton("👑 এডমিন প্যানেল বাটন কালার", callback_data="adm_bcol_grp_admin_panel")],
            [InlineKeyboardButton("🔹 সাধারণ ইনলাইন বাটন", callback_data="adm_bcol_grp_inline_default")],
            [
                InlineKeyboardButton("✅ অ্যাকশন বাটন", callback_data="adm_bcol_grp_inline_success"),
                InlineKeyboardButton("❌ ডিলিট/অফ বাটন", callback_data="adm_bcol_grp_inline_danger"),
            ],
            [InlineKeyboardButton("🎯 নির্দিষ্ট বাটনের কালার সেট করুন", callback_data="adm_bcol_custom")],
            [InlineKeyboardButton("🔄 Reset Default Colors", callback_data="adm_bcol_reset")],
            [InlineKeyboardButton("< Back to Dashboard", callback_data="adm_root")],
        ]), group="admin_panel")
        await query.edit_message_text(txt, parse_mode=ParseMode.HTML, reply_markup=kb)
        return

    if data == "adm_bcol_reset":
        save_btn_color_config({
            "main_menu": "primary",
            "admin_panel": "primary",
            "inline_default": "primary",
            "inline_success": "success",
            "inline_danger": "danger",
            "overrides": {},
        })
        await query.message.reply_text("✅ সব বাটনের ডিফল্ট কালার রিসেট করা হয়েছে!", reply_markup=get_main_reply_keyboard(user_id))
        await handle_admin_callbacks(query, "adm_btn_colors", context)
        return

    if data.startswith("adm_bcol_grp_"):
        grp = data.replace("adm_bcol_grp_", "", 1)
        grp_names = {
            "main_menu": "📱 মেইন মেনু বাটন",
            "admin_panel": "👑 এডমিন প্যানেল বাটন",
            "inline_default": "🔹 সাধারণ ইনলাইন বাটন",
            "inline_success": "✅ কনফার্ম / অ্যাকশন বাটন",
            "inline_danger": "❌ ক্যানসেল / ডিলিট / অফ বাটন",
        }
        kb = InlineKeyboardMarkup([
            [
                make_inline_btn("🔵 Blue (Primary)", style="primary", callback_data=f"adm_bcol_set_{grp}_primary"),
                make_inline_btn("🟢 Green (Success)", style="success", callback_data=f"adm_bcol_set_{grp}_success"),
            ],
            [
                make_inline_btn("🔴 Red (Danger)", style="danger", callback_data=f"adm_bcol_set_{grp}_danger"),
                make_inline_btn("⚪ Default (No Color)", style="none", callback_data=f"adm_bcol_set_{grp}_none"),
            ],
            [make_inline_btn("< Back", callback_data="adm_btn_colors")],
        ])
        await query.edit_message_text(
            f"🎨 <b>{grp_names.get(grp, grp)}</b>-এর জন্য পছন্দের কালার সিলেক্ট করুন:",
            parse_mode=ParseMode.HTML,
            reply_markup=kb,
        )
        return

    if data.startswith("adm_bcol_set_"):
        rest = data.replace("adm_bcol_set_", "", 1)
        grp, col = rest.rsplit("_", 1)
        cfg = get_btn_color_config()
        cfg[grp] = col
        save_btn_color_config(cfg)
        if grp == "main_menu":
            await query.message.reply_text(
                f"✅ মেইন মেনু বাটনের কালার আপডেট হয়েছে!",
                reply_markup=get_main_reply_keyboard(user_id),
            )
        await handle_admin_callbacks(query, "adm_btn_colors", context)
        return

    if data == "adm_bcol_custom":
        cfg = get_btn_color_config()
        overrides = cfg.get("overrides") or {}
        menu_btns = [lbl for r in get_menu_layout() for lbl in r] + ["👑 Admin Panel"] + ADMIN_PANEL_CUSTOM_BUTTONS
        kb = []
        row = []
        for idx, lbl in enumerate(menu_btns):
            cur_c = overrides.get(lbl, "")
            tag = {"primary": "🔵", "success": "🟢", "danger": "🔴", "none": "⚪"}.get(cur_c, "")
            row.append(InlineKeyboardButton(f"{tag} {lbl}", callback_data=f"adm_bcol_pickmenu_{idx}"))
            if len(row) == 2:
                kb.append(row)
                row = []
        if row:
            kb.append(row)
        kb.append([InlineKeyboardButton("✏️ যেকোনো বাটনের নাম লিখে কালার দিন", callback_data="adm_bcol_type_btn")])
        if overrides:
            kb.append([InlineKeyboardButton("🗑️ Clear All Custom Overrides", callback_data="adm_bcol_clear_ov")])
        kb.append([InlineKeyboardButton("< Back", callback_data="adm_btn_colors")])
        await query.edit_message_text(
            "🎯 <b>নির্দিষ্ট বাটনের কালার সেট করুন (মেইন মেনু + অ্যাডমিন প্যানেল বাটন):</b>\nযেকোনো বাটনে ক্লিক করুন অথবা নিজের মতো বাটনের টেক্সট লিখে কালার দিন:",
            parse_mode=ParseMode.HTML,
            reply_markup=colorize_markup(InlineKeyboardMarkup(kb)),
        )
        return

    if data == "adm_bcol_clear_ov":
        cfg = get_btn_color_config()
        cfg["overrides"] = {}
        save_btn_color_config(cfg)
        await query.message.reply_text("✅ সব কাস্টম বাটন কালার মুছে ফেলা হয়েছে!", reply_markup=get_main_reply_keyboard(user_id))
        await handle_admin_callbacks(query, "adm_bcol_custom", context)
        return

    if data.startswith("adm_bcol_pickmenu_"):
        idx = int(data.split("_")[3])
        menu_btns = [lbl for r in get_menu_layout() for lbl in r] + ["👑 Admin Panel"] + ADMIN_PANEL_CUSTOM_BUTTONS
        if idx >= len(menu_btns):
            await handle_admin_callbacks(query, "adm_bcol_custom", context)
            return
        lbl = menu_btns[idx]
        set_user_state(user_id, "ADMIN_WAITING_CUSTOM_BTN_COLOR_CHOICE", {"btn_label": lbl})
        kb = InlineKeyboardMarkup([
            [
                make_inline_btn("🔵 Blue", style="primary", callback_data="adm_bcol_ov_primary"),
                make_inline_btn("🟢 Green", style="success", callback_data="adm_bcol_ov_success"),
            ],
            [
                make_inline_btn("🔴 Red", style="danger", callback_data="adm_bcol_ov_danger"),
                make_inline_btn("⚪ Default", style="none", callback_data="adm_bcol_ov_none"),
            ],
            [make_inline_btn("< Back", callback_data="adm_bcol_custom")],
        ])
        await query.edit_message_text(
            f"🎯 বাটন: <b>{html.escape(lbl)}</b>\nএই বাটনটির জন্য কালার সিলেক্ট করুন:",
            parse_mode=ParseMode.HTML,
            reply_markup=kb,
        )
        return

    if data == "adm_bcol_type_btn":
        set_user_state(user_id, "ADMIN_WAITING_BTN_COLOR_TEXT", {})
        await query.message.reply_text(
            "✏️ আপনি যে বাটনের কালার পরিবর্তন করতে চান, সেই বাটনের লেখাটি হুবহু লিখে পাঠান (যেমন: <code>Buy Now</code> বা <code>🛒 Purchases</code>):",
            parse_mode=ParseMode.HTML,
        )
        return

    if data.startswith("adm_bcol_ov_"):
        col = data.replace("adm_bcol_ov_", "", 1)
        _, sdata_b = get_user_state(user_id)
        lbl = sdata_b.get("btn_label", "")
        if lbl:
            cfg = get_btn_color_config()
            if not isinstance(cfg.get("overrides"), dict):
                cfg["overrides"] = {}
            cfg["overrides"][lbl] = col
            save_btn_color_config(cfg)
            set_user_state(user_id, None)
            await query.message.reply_text(
                f"✅ <b>{html.escape(lbl)}</b> বাটনের কালার আপডেট হয়েছে!",
                parse_mode=ParseMode.HTML,
                reply_markup=get_main_reply_keyboard(user_id),
            )
        await handle_admin_callbacks(query, "adm_bcol_custom", context)
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
        await query.edit_message_text(txt, parse_mode=ParseMode.HTML, reply_markup=colorize_markup(InlineKeyboardMarkup(kb)))
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
        await query.edit_message_text(f"🎛️ <b>বাটন এডিটর:</b> <code>{lbl}</code>", parse_mode=ParseMode.HTML, reply_markup=colorize_markup(InlineKeyboardMarkup(kb)))
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
        await query.edit_message_text("🖼️ <b>[প্রোডাক্ট ফটো ম্যানেজার — ধাপ ১]</b>", parse_mode=ParseMode.HTML, reply_markup=colorize_markup(InlineKeyboardMarkup(kb)))
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
        await query.edit_message_text("🖼️ <b>[প্রোডাক্ট ফটো ম্যানেজার — ধাপ ২]</b>", parse_mode=ParseMode.HTML, reply_markup=colorize_markup(InlineKeyboardMarkup(kb)))
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
        await query.edit_message_text("🖼️ <b>[প্রোডাক্ট ফটো ম্যানেজার — ধাপ ৩]</b>", parse_mode=ParseMode.HTML, reply_markup=colorize_markup(InlineKeyboardMarkup(kb)))
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
        await query.edit_message_text(f"🖼️ <b>{p['sub_name']} — {p['name']}</b>\n\nঅবস্থা: {status}", parse_mode=ParseMode.HTML, reply_markup=colorize_markup(InlineKeyboardMarkup(kb)))
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
        kb = colorize_markup(InlineKeyboardMarkup([
            [InlineKeyboardButton("📱 Edit bKash Number", callback_data="adm_set_bkash")],
            [InlineKeyboardButton("📱 Edit Nagad Number", callback_data="adm_set_nagad")],
            [InlineKeyboardButton("💎 Edit Binance Info", callback_data="adm_set_binance")],
            [InlineKeyboardButton("💰 Edit Min Deposit", callback_data="adm_set_mindep_manual")],
            [InlineKeyboardButton("💵 Edit USDT Rate", callback_data="adm_set_rate")],
            [InlineKeyboardButton("< Back to Dashboard", callback_data="adm_root")],
        ]))
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
        await query.edit_message_text(f"👥 <b>মোট ইউজার: {total_u} জন</b>", parse_mode=ParseMode.HTML, reply_markup=colorize_markup(InlineKeyboardMarkup(kb)))
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
        ukb = colorize_markup(InlineKeyboardMarkup([
            [InlineKeyboardButton("➕ Add Balance", callback_data=f"adm_usr_addbal_{u['user_id']}"), InlineKeyboardButton("➖ Deduct Balance", callback_data=f"adm_usr_dedbal_{u['user_id']}")],
            [InlineKeyboardButton("💬 Direct Message", url=f"tg://user?id={u['user_id']}")],
            [InlineKeyboardButton("< Back", callback_data="adm_usr_list_0")],
        ]))
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
            kb = colorize_markup(InlineKeyboardMarkup([[InlineKeyboardButton("✅ Approve", callback_data=f"adm_dep_approve_{d['id']}"), InlineKeyboardButton("❌ Reject", callback_data=f"adm_dep_reject_{d['id']}")]]))
            await query.message.reply_text(f"💳 <b>Pending Deposit</b>\n👤 {d['uname']} (<code>{d['user_id']}</code>)\n💰 {d['amount']:.2f} BDT ({d['method']})\n🔢 TrxID: <code>{d['trx_id']}</code>", parse_mode=ParseMode.HTML, reply_markup=kb)
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
            await query.edit_message_text(f"✅ ডিপোজিট ({dep['amount']:.2f} BDT) অনুমোদিত হয়েছে।")
            try:
                await context.bot.send_message(chat_id=dep["user_id"], text=f"🎉 <b>ডিপোজিট সফল!</b>\nআপনার অ্যাকাউন্টে {dep['amount']:.2f} BDT যুক্ত হয়েছে।", parse_mode=ParseMode.HTML)
            except Exception:
                pass
        elif action == "reject":
            c.execute("UPDATE deposits SET status = 'rejected' WHERE id = ?", (dep_id,))
            conn.commit()
            conn.close()
            add_notification(dep["user_id"], "❌ Deposit Rejected", f"Your deposit ({dep['amount']:.2f} BDT) was rejected.")
            await query.edit_message_text(f"❌ ডিপোজিট ({dep['amount']:.2f} BDT) বাতিল করা হয়েছে।")
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
            kb = colorize_markup(InlineKeyboardMarkup([
                [InlineKeyboardButton(f"🚀 Deliver Order #{o['id']}", callback_data=f"adm_fulfill_{o['id']}")],
                [InlineKeyboardButton(f"❌ Cancel & Refund #{o['id']}", callback_data=f"adm_refund_{o['id']}")]
            ]))
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

    # --- 🎟️ PROMO / COUPON MANAGER (BOT ADMIN) ---
    if data == "adm_promo_mgr":
        conn = get_db()
        c = conn.cursor()
        c.execute("SELECT code, discount_pct, COALESCE(product_name, 'All Products') as product_name FROM coupons WHERE is_active = 1 ORDER BY code ASC")
        coupons = c.fetchall()
        conn.close()
        txt = "🎟️ <b>Promo / Coupon Manager</b>\n─────────────────────────────\n"
        kb = []
        if coupons:
            for cp in coupons:
                txt += f"• <code>{cp['code']}</code> — <b>{cp['discount_pct']:.0f}% OFF</b> ({cp['product_name']})\n"
                kb.append([InlineKeyboardButton(f"🗑️ Delete {cp['code']} ({cp['discount_pct']:.0f}%)", callback_data=f"adm_promo_del_{cp['code']}")])
        else:
            txt += "<i>বর্তমানে কোনো অ্যাক্টিভ কুপন নেই।</i>\n"
        kb.append([InlineKeyboardButton("➕ Add New Coupon", callback_data="adm_promo_add_select")])
        kb.append([InlineKeyboardButton("< Back to Dashboard", callback_data="adm_root")])
        await query.edit_message_text(txt, parse_mode=ParseMode.HTML, reply_markup=colorize_markup(InlineKeyboardMarkup(kb)))
        return

    if data.startswith("adm_promo_del_"):
        ccode = data.replace("adm_promo_del_", "", 1).strip()
        conn = get_db()
        c = conn.cursor()
        c.execute("DELETE FROM coupons WHERE code = ?", (ccode,))
        conn.commit()
        conn.close()
        await handle_admin_callbacks(query, "adm_promo_mgr", context)
        return

    if data == "adm_promo_add_select":
        conn = get_db()
        c = conn.cursor()
        c.execute("SELECT p.id, p.name, s.name as sub_name FROM products p JOIN subcategories s ON p.subcategory_id = s.id ORDER BY s.name ASC, p.id ASC LIMIT 35")
        prods = c.fetchall()
        conn.close()
        kb = [[InlineKeyboardButton("🌐 All Products (সকল প্রোডাক্ট)", callback_data="adm_promo_target_0")]]
        for p in prods:
            kb.append([InlineKeyboardButton(f"📦 {p['sub_name']} — {p['name']}", callback_data=f"adm_promo_target_{p['id']}")])
        kb.append([InlineKeyboardButton("< Back", callback_data="adm_promo_mgr")])
        await query.edit_message_text("🎟️ <b>নতুন কুপন তৈরি (ধাপ ১/৩):</b>\nকুপনটি কোন প্রোডাক্টে কাজ করবে সিলেক্ট করুন:", parse_mode=ParseMode.HTML, reply_markup=colorize_markup(InlineKeyboardMarkup(kb)))
        return

    if data.startswith("adm_promo_target_"):
        t_pid = int(data.split("_")[3])
        t_pname = "All Products"
        if t_pid > 0:
            conn = get_db()
            c = conn.cursor()
            c.execute("SELECT p.name, s.name as sub_name FROM products p JOIN subcategories s ON p.subcategory_id = s.id WHERE p.id = ?", (t_pid,))
            pr = c.fetchone()
            conn.close()
            if pr:
                t_pname = f"{pr['sub_name']} — {pr['name']}"
        set_user_state(user_id, "ADMIN_WAITING_COUPON_CODE", {"product_id": t_pid, "product_name": t_pname})
        await query.message.reply_text(
            f"🎟️ <b>টার্গেট প্রোডাক্ট:</b> {t_pname}\n\n"
            f"<b>(ধাপ ২/৩):</b> কুপন কোডটি লিখে পাঠান (যেমন: <code>VIP20</code>):",
            parse_mode=ParseMode.HTML,
        )
        return

    # --- 📢 UNLIMITED FORCE JOIN MANAGER (BOT ADMIN) ---
    if data == "adm_fj_mgr":
        conn = get_db()
        c = conn.cursor()
        c.execute("SELECT id, title, chat_id, invite_link, is_active FROM force_joins ORDER BY id ASC")
        fjs = c.fetchall()
        conn.close()
        txt = (
            "📢 <b>Unlimited Force Join Manager</b>\n"
            "─────────────────────────────\n"
            f"মোট যুক্ত গ্রুপ/চ্যানেল: <b>{len(fjs)}</b> টি\n\n"
        )
        kb = []
        for fj in fjs:
            st_ic = "🟢" if fj["is_active"] else "🔴"
            txt += f"{st_ic} <b>{html.escape(fj['title'])}</b> — <code>{html.escape(fj['chat_id'])}</code>\n"
            kb.append([
                InlineKeyboardButton(f"{st_ic} {fj['title'][:18]}", callback_data=f"adm_fj_toggle_{fj['id']}"),
                InlineKeyboardButton("✏️", callback_data=f"adm_fj_edit_{fj['id']}"),
                InlineKeyboardButton("🗑️", callback_data=f"adm_fj_del_{fj['id']}"),
            ])
        kb.append([InlineKeyboardButton("➕ Add New Group / Channel", callback_data="adm_fj_add")])
        kb.append([InlineKeyboardButton("< Back to Dashboard", callback_data="adm_root")])
        await query.edit_message_text(txt, parse_mode=ParseMode.HTML, reply_markup=colorize_markup(InlineKeyboardMarkup(kb)))
        return

    if data == "adm_fj_add":
        set_user_state(user_id, "ADMIN_WAITING_FJ_TITLE", {})
        await query.message.reply_text("📢 <b>নতুন Force Join যুক্ত করুন (ধাপ ১/৩):</b>\nবাটনের নাম লিখুন (যেমন: <code>📢 Join Main Channel</code>):", parse_mode=ParseMode.HTML)
        return

    if data.startswith("adm_fj_toggle_"):
        fid = int(data.split("_")[3])
        conn = get_db()
        c = conn.cursor()
        c.execute("UPDATE force_joins SET is_active = CASE WHEN COALESCE(is_active,1)=1 THEN 0 ELSE 1 END WHERE id = ?", (fid,))
        conn.commit()
        conn.close()
        await handle_admin_callbacks(query, "adm_fj_mgr", context)
        return

    if data.startswith("adm_fj_del_"):
        fid = int(data.split("_")[3])
        conn = get_db()
        c = conn.cursor()
        c.execute("DELETE FROM force_joins WHERE id = ?", (fid,))
        conn.commit()
        conn.close()
        await handle_admin_callbacks(query, "adm_fj_mgr", context)
        return

    if data.startswith("adm_fj_edit_"):
        fid = int(data.split("_")[3])
        set_user_state(user_id, "ADMIN_WAITING_FJ_EDIT_TITLE", {"fj_id": fid})
        await query.message.reply_text("✏️ <b>Force Join এডিট (ধাপ ১/৩):</b>\nনতুন বাটন নাম লিখে পাঠান:", parse_mode=ParseMode.HTML)
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
        try:
            await context.bot.send_chat_action(chat_id=update.effective_chat.id, action="typing")
        except Exception:
            pass
        reply_txt = await asyncio.to_thread(generate_smart_ai_reply, text, user.id)
        safe_txt = _sanitize_ai_for_telegram(reply_txt)
        try:
            await update.message.reply_text(safe_txt[:4000], parse_mode=ParseMode.HTML)
        except Exception:
            import re as _re
            plain_txt = _re.sub(r"<[^>]+>", "", safe_txt)
            await update.message.reply_text(plain_txt[:4000])
        return

    if state == "WAITING_OUTLOOK_STR":
        try:
            await context.bot.send_chat_action(chat_id=update.effective_chat.id, action="typing")
        except Exception:
            pass
        res = await asyncio.to_thread(fetch_outlook_otp, text)
        set_user_state(user.id, None)
        tok_key = secrets.token_hex(4)
        _LIVE_TOOL_STOP.discard(tok_key)
        _TOOL_CACHE[tok_key] = text
        if len(_TOOL_CACHE) > 300:
            _TOOL_CACHE.pop(next(iter(_TOOL_CACHE)))
        kb = _build_mail_live_kb(tok_key, live=True)
        now_t = datetime.datetime.now().strftime("%I:%M:%S %p")
        if res.get("ok"):
            sent_mail = await update.message.reply_text(
                f"✉️ <b>Outlook / Hotmail OTP Found!</b>\n"
                f"📧 <b>Email:</b> <code>{html.escape(str(res.get('email', '')))}</code>\n"
                f"📌 <b>Subject:</b> {html.escape(str(res.get('subject', 'Inbox')))}\n"
                f"🔢 <b>Code / Message:</b> <code>{html.escape(str(res.get('code', '')))}</code>\n"
                f"⏳ <i>প্রতি ৫ সেকেন্ডে অটো-আপডেট হচ্ছে ({now_t})</i>",
                parse_mode=ParseMode.HTML,
                reply_markup=kb,
            )
        else:
            sent_mail = await update.message.reply_text(
                f"❌ <b>Error:</b> {html.escape(str(res.get('error', 'Failed')))}\n"
                f"⏳ <i>প্রতি ৫ সেকেন্ডে অটো-আপডেট হচ্ছে ({now_t})</i>",
                parse_mode=ParseMode.HTML,
                reply_markup=kb,
            )
        asyncio.create_task(_run_live_outlook_countdown(context.bot, sent_mail.chat_id, sent_mail.message_id, text, tok_key))
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
        kb = colorize_markup(InlineKeyboardMarkup([
            [InlineKeyboardButton("📱 bKash", callback_data="usr_dep_method_bKash"), InlineKeyboardButton("📱 Nagad", callback_data="usr_dep_method_Nagad")],
            [InlineKeyboardButton("💎 Binance", callback_data="usr_dep_method_Binance")],
        ]))
        dm_wait_txt = render_custom_bot_message(
            "deposit_method",
            f"✅ আপনি <b>{amt:.2f} BDT</b> ডিপোজিট করতে চেয়েছেন।\n\nনিচের কোন মাধ্যমে সেন্ড মানি করবেন সিলেক্ট করুন:",
            amount=f"{amt:.2f}",
            name=html.escape(user.full_name or ""),
            shop_name=get_setting("shop_name", "CyberNest Market"),
        )
        await send_smart_bot_message(
            update.message,
            "deposit_method",
            dm_wait_txt,
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
            f"✅ <b>ডিপোজিট রিকোয়েস্ট সাবমিটেড!</b>\n\n💳 মেথড: <b>{method}</b>\n💰 পরিমাণ: <b>{amount:.2f} BDT</b>\n🔢 TrxID: <code>{trx}</code>",
            parse_mode=ParseMode.HTML,
        )

        adm_kb = colorize_markup(InlineKeyboardMarkup([[InlineKeyboardButton("✅ Approve", callback_data=f"adm_dep_approve_{dep_id}"), InlineKeyboardButton("❌ Reject", callback_data=f"adm_dep_reject_{dep_id}")]]))
        try:
            sent_dep = await context.bot.send_message(
                chat_id=SUPER_ADMIN_ID,
                text=f"🔔 <b>নতুন ডিপোজিট রিকোয়েস্ট</b>\n👤 {user.full_name} (<code>{user.id}</code>)\n📱 {method} | 💰 {amount:.2f} BDT\n🔢 TrxID: <code>{trx}</code>",
                parse_mode=ParseMode.HTML,
                reply_markup=adm_kb,
            )
            _cd = get_db()
            _cd.cursor().execute("UPDATE deposits SET admin_msg_id = ? WHERE id = ?", (sent_dep.message_id, dep_id))
            _cd.commit()
            _cd.close()
        except Exception:
            pass
        return

    if state == "WAITING_2FA_KEY":
        parsed_codes = generate_universal_2fa_codes(text)
        if parsed_codes:
            set_user_state(user.id, None)
            tok_key = secrets.token_hex(4)
            _LIVE_TOOL_STOP.discard(tok_key)
            _TOOL_CACHE[tok_key] = text
            if len(_TOOL_CACHE) > 300:
                _TOOL_CACHE.pop(next(iter(_TOOL_CACHE)))
            kb = _build_2fa_live_kb(tok_key, live=True)
            msg_out = _format_2fa_live_text(parsed_codes, live_active=True)
            sent_2fa = await update.message.reply_text(
                msg_out,
                parse_mode=ParseMode.HTML,
                reply_markup=kb,
            )
            asyncio.create_task(_run_live_2fa_countdown(context.bot, sent_2fa.chat_id, sent_2fa.message_id, text, tok_key))
        else:
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
        try:
            await context.bot.send_chat_action(chat_id=update.effective_chat.id, action="typing")
        except Exception:
            pass
        ans = await asyncio.to_thread(generate_smart_ai_reply, text, user.id)
        safe_ans = _sanitize_ai_for_telegram(ans)
        try:
            await update.message.reply_text(safe_ans[:4000], parse_mode=ParseMode.HTML)
        except Exception:
            import re as _re
            await update.message.reply_text(_re.sub(r"<[^>]+>", "", safe_ans)[:4000])
        return

    if state == "ADMIN_WAITING_MSG_TEXT":
        sk = sdata.get("slot_key", "")
        raw_val = update.message.text.strip()
        if sk and raw_val:
            set_setting(f"msg_text_{sk}", raw_val)
            set_user_state(user_id, None)
            slabel = BOT_MESSAGE_SLOT_MAP.get(sk, sk)
            try:
                await update.message.reply_text(
                    f"✅ <b>{html.escape(slabel)}</b>-এর মেসেজ সফলভাবে আপডেট হয়েছে!\n\n"
                    f"👁️ <b>লাইভ প্রিভিউ:</b>\n─────────────────────\n{raw_val}",
                    parse_mode=ParseMode.HTML,
                )
            except Exception as html_err:
                await update.message.reply_text(
                    f"✅ {slabel}-এর মেসেজ সেভ হয়েছে!\n⚠️ তবে HTML ট্যাগে ছোট ভুল থাকতে পারে ({html_err}) — প্রয়োজনে আবার এডিট করে ঠিক করে নিন।"
                )
        return

    if state == "ADMIN_WAITING_MSG_PHOTO":
        sk = sdata.get("slot_key", "")
        val = text.strip()
        if sk and val:
            set_setting(f"msg_photo_{sk}", val)
            set_user_state(user.id, None)
            slabel = BOT_MESSAGE_SLOT_MAP.get(sk, sk)
            await update.message.reply_text(
                f"✅ <b>{html.escape(slabel)}</b>-এর ছবি সফলভাবে সেট হয়েছে!\n📋 Saved ID/URL: <code>{html.escape(val)}</code>",
                parse_mode=ParseMode.HTML,
            )
        return

    if state == "ADMIN_WAITING_BTN_COLOR_TEXT":
        lbl = text.strip()
        set_user_state(user.id, "ADMIN_WAITING_CUSTOM_BTN_COLOR_CHOICE", {"btn_label": lbl})
        kb = InlineKeyboardMarkup([
            [
                make_inline_btn("🔵 Blue", style="primary", callback_data="adm_bcol_ov_primary"),
                make_inline_btn("🟢 Green", style="success", callback_data="adm_bcol_ov_success"),
            ],
            [
                make_inline_btn("🔴 Red", style="danger", callback_data="adm_bcol_ov_danger"),
                make_inline_btn("⚪ Default", style="none", callback_data="adm_bcol_ov_none"),
            ],
        ])
        await update.message.reply_text(
            f"🎯 বাটন: <b>{html.escape(lbl)}</b>\nএবার এই বাটনটির জন্য কালার সিলেক্ট করুন:",
            parse_mode=ParseMode.HTML,
            reply_markup=kb,
        )
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
        kb = colorize_markup(InlineKeyboardMarkup([
            [InlineKeyboardButton("✅ হ্যাঁ (বাটন যোগ করব)", callback_data="adm_bcast_add_btn")],
            [InlineKeyboardButton("❌ না (বাটন ছাড়াই পাঠাব)", callback_data="adm_bcast_skip_btn")],
        ]))
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

    if state == "ADMIN_WAITING_BCAST_BTN_CHOICE":
        await update.message.reply_text("👆 অনুগ্রহ করে উপরের 'হ্যাঁ' অথবা 'না' বাটনে চাপুন।")
        return

    if state == "ADMIN_WAITING_COUPON_CODE":
        sdata["code"] = text.strip().upper()
        set_user_state(user.id, "ADMIN_WAITING_COUPON_PCT", sdata)
        await update.message.reply_text(f"✅ কোড: <code>{sdata['code']}</code>\n<b>(ধাপ ৩/৩):</b> কত পার্সেন্ট (%) ডিসকাউন্ট দেবেন লিখুন (যেমন: 10, 25):", parse_mode=ParseMode.HTML)
        return

    if state == "ADMIN_WAITING_COUPON_PCT":
        try:
            pct = float(text.strip())
            if pct <= 0 or pct > 100:
                raise ValueError()
        except ValueError:
            await update.message.reply_text("❌ ১ থেকে ১০০-এর মধ্যে সঠিক সংখ্যা লিখুন:")
            return
        ccode = sdata.get("code", "")
        pid = int(sdata.get("product_id", 0) or 0)
        pname = sdata.get("product_name", "All Products")
        conn = get_db()
        c = conn.cursor()
        if conn.is_pg:
            c.execute(
                "INSERT INTO coupons (code, discount_pct, is_active, product_id, product_name) VALUES (%s, %s, 1, %s, %s) ON CONFLICT (code) DO UPDATE SET discount_pct = EXCLUDED.discount_pct, is_active = 1, product_id = EXCLUDED.product_id, product_name = EXCLUDED.product_name",
                (ccode, pct, pid, pname),
            )
        else:
            c.execute(
                "INSERT OR REPLACE INTO coupons (code, discount_pct, is_active, product_id, product_name) VALUES (?, ?, 1, ?, ?)",
                (ccode, pct, pid, pname),
            )
        conn.commit()
        conn.close()
        set_user_state(user.id, None)
        await update.message.reply_text(f"🎉 কুপন <code>{ccode}</code> ({pct:.0f}% OFF — {pname}) সেভ হয়েছে!", parse_mode=ParseMode.HTML)
        return

    if state == "ADMIN_WAITING_FJ_TITLE":
        set_user_state(user.id, "ADMIN_WAITING_FJ_CHATID", {"title": text.strip()})
        await update.message.reply_text("📢 <b>(ধাপ ২/৩):</b> গ্রুপ বা চ্যানেলের <b>Chat ID বা @username</b> লিখুন (যেমন: <code>-1001234567890</code> বা <code>@channel</code>):", parse_mode=ParseMode.HTML)
        return

    if state == "ADMIN_WAITING_FJ_CHATID":
        sdata["chat_id"] = text.strip()
        set_user_state(user.id, "ADMIN_WAITING_FJ_LINK", sdata)
        await update.message.reply_text("🔗 <b>(ধাপ ৩/৩):</b> গ্রুপ বা চ্যানেলের <b>Invite Link (https://t.me/...)</b> লিখে পাঠান:", parse_mode=ParseMode.HTML)
        return

    if state == "ADMIN_WAITING_FJ_LINK":
        link = text.strip()
        conn = get_db()
        c = conn.cursor()
        c.execute(
            "INSERT INTO force_joins (title, chat_id, invite_link, is_active) VALUES (?, ?, ?, 1)",
            (sdata.get("title", "Join Channel"), sdata.get("chat_id", ""), link),
        )
        conn.commit()
        conn.close()
        set_user_state(user.id, None)
        await update.message.reply_text(f"✅ নতুন Force Join <b>{html.escape(sdata.get('title', ''))}</b> সফলভাবে যুক্ত হয়েছে!", parse_mode=ParseMode.HTML)
        return

    if state == "ADMIN_WAITING_FJ_EDIT_TITLE":
        sdata["title"] = text.strip()
        set_user_state(user.id, "ADMIN_WAITING_FJ_EDIT_CHATID", sdata)
        await update.message.reply_text("✏️ <b>(ধাপ ২/৩):</b> নতুন Chat ID বা @username লিখুন:", parse_mode=ParseMode.HTML)
        return

    if state == "ADMIN_WAITING_FJ_EDIT_CHATID":
        sdata["chat_id"] = text.strip()
        set_user_state(user.id, "ADMIN_WAITING_FJ_EDIT_LINK", sdata)
        await update.message.reply_text("🔗 <b>(ধাপ ৩/৩):</b> নতুন Invite Link লিখুন:", parse_mode=ParseMode.HTML)
        return

    if state == "ADMIN_WAITING_FJ_EDIT_LINK":
        fid = int(sdata.get("fj_id", 0))
        conn = get_db()
        c = conn.cursor()
        c.execute(
            "UPDATE force_joins SET title = ?, chat_id = ?, invite_link = ? WHERE id = ?",
            (sdata.get("title", "Join"), sdata.get("chat_id", ""), text.strip(), fid),
        )
        conn.commit()
        conn.close()
        set_user_state(user.id, None)
        await update.message.reply_text("✅ Force Join তথ্য আপডেট হয়েছে!")
        return

async def photo_message_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user = update.effective_user
    if not is_admin(user.id):
        return

    state, sdata = get_user_state(user.id)
    if state == "ADMIN_WAITING_MSG_PHOTO":
        photo_file_id = update.message.photo[-1].file_id
        sk = sdata.get("slot_key", "")
        if sk:
            set_setting(f"msg_photo_{sk}", photo_file_id)
            set_user_state(user.id, None)
            slabel = BOT_MESSAGE_SLOT_MAP.get(sk, sk)
            await update.message.reply_text(
                f"✅ <b>{html.escape(slabel)}</b>-এর জন্য ছবি সফলভাবে সেভ হয়েছে!\n📋 File ID: <code>{photo_file_id}</code>",
                parse_mode=ParseMode.HTML,
            )
        return

    if state == "ADMIN_WAITING_BCAST_MSG":
        photo_file_id = update.message.photo[-1].file_id
        caption = update.message.caption or ""
        set_user_state(user.id, "ADMIN_WAITING_BCAST_BTN_CHOICE", {"msg_text": caption, "photo_id": photo_file_id})
        kb = colorize_markup(InlineKeyboardMarkup([
            [InlineKeyboardButton("✅ হ্যাঁ (বাটন যোগ করব)", callback_data="adm_bcast_add_btn")],
            [InlineKeyboardButton("❌ না (বাটন ছাড়াই পাঠাব)", callback_data="adm_bcast_skip_btn")],
        ]))
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
    default_gw = (
        f"🌸 <b>আসসালামু আলাইকুম</b>, {user_link}! ✨\n\n"
        f"আমাদের কমিউনিটি গ্রুপে আপনাকে স্বাগতম। 🎉"
    )
    welcome_text = render_custom_bot_message(
        "group_welcome",
        default_gw,
        user_link=user_link,
        name=html.escape(user.full_name or "মেম্বার"),
        shop_name=get_setting("shop_name", "CyberNest Market"),
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
    app.run_polling(allowed_updates=Update.ALL_TYPES, drop_pending_updates=True)

if __name__ == "__main__":
    main()
