import asyncio
import json
import os
import random
import time
from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup
from telegram.ext import (
    ApplicationBuilder, 
    ChatMemberHandler, 
    MessageHandler,
    CallbackQueryHandler, 
    CommandHandler,
    filters,
    ContextTypes
)

# --- কনফিগারেশন ফাইল ও ডিফল্ট সেটিংস ---
CONFIG_FILE = "config.json"

DEFAULT_SETTINGS = {
    "is_tx_active": True,
    "transaction_interval": 300,  # ৫ মিনিট (সেকেন্ডে)
    "min_amount": 50.0,
    "max_amount": 2000.0,
    "welcome_delete_delay": 30,
    "welcome_text": (
        "🌸 <b>আসসালামু আলাইকুম</b>, {user_link}!\n\n"
        "আমাদের কমিউনিটিতে আপনাকে স্বাগতম। 🎉\n"
        "📌 গ্রুপের নিয়ম-কানুন মেনে চলুন এবং নিয়মিত আপডেট উপভোগ করুন।"
    ),
    "buttons": [
        {"name": "FAST GMAIL SELL", "url": "https://t.me/NEW_FRESH_GMAILACCOUNTSELL50_bot"}
    ]
}

def load_settings():
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                return {**DEFAULT_SETTINGS, **json.load(f)}
        except Exception:
            return DEFAULT_SETTINGS.copy()
    return DEFAULT_SETTINGS.copy()

def save_settings(data):
    try:
        with open(CONFIG_FILE, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
    except Exception as e:
        print(f"Error saving settings: {e}")

SETTINGS = load_settings()

# টেলিগ্রাম ও গ্রুপ তথ্য
BOT_TOKEN = "8826168593:AAEfgcSn0Ckte8Ayv3OKfzAPic43FIFdT_g"
WELCOME_CHAT_ID = "-1004471047712"     # ১ম গ্রুপ
TRANSACTION_CHAT_ID = "-1003991468184"  # ২য় গ্রুপ

SECRET_ADMIN_KEY = "talha#secret99"
AUTHORIZED_ADMIN_ID = None

NAMES = [
    "Arif Hasan", "Tanvir Ahmed", "Sakib Al Hasan", "Rahim Uddin", 
    "Mehedi Hasan", "Nusrat Jahan", "Sadia Islam", "Farhana Akter", 
    "Jannatul Ferdous", "Ayesha Siddiqua", "Rifat Hossain", "Sumaiya Kabir", 
    "Shakil Khan", "Fahim Shahriar", "Mim Akter", "Tasnim Sultana",
    "Mahmudul Hasan", "Naimur Rahman", "Sabiha Sultana", "Habibur Rahman"
]

PAYMENT_METHODS = ["Bkash", "Nagad", "Binance"]
recently_welcomed_users = set()

# --- Render স্লিপ প্রতিরোধক সার্ভার ---
async def start_dummy_web_server():
    port = int(os.environ.get("PORT", 10000))
    async def handle_ping(reader, writer):
        await reader.read(1024)
        response = (
            "HTTP/1.1 200 OK\r\n"
            "Content-Type: text/plain; charset=utf-8\r\n"
            "Content-Length: 17\r\n"
            "Connection: close\r\n\r\n"
            "Bot is Active 24/7"
        )
        writer.write(response.encode("utf-8"))
        await writer.drain()
        writer.close()
        await writer.wait_closed()

    server = await asyncio.start_server(handle_ping, "0.0.0.0", port)
    print(f"Server running on port {port}")
    async with server:
        await server.serve_forever()

# --- ডাইনামিক বাটন কিবোর্ড তৈরি ---
def get_transaction_keyboard():
    keyboard = []
    buttons = SETTINGS.get("buttons", [])
    for btn in buttons:
        keyboard.append([InlineKeyboardButton(btn["name"], url=btn["url"])])
    return InlineKeyboardMarkup(keyboard) if keyboard else None

# --- ৩ ধাপের ট্রানজ্যাকশন মেসেজ ---
async def execute_3_step_transaction(bot, chat_id):
    customer_name = random.choice(NAMES)
    min_a = SETTINGS.get("min_amount", 50.0)
    max_a = SETTINGS.get("max_amount", 2000.0)
    amount = f"{random.uniform(min_a, max_a):.2f}"
    method = random.choice(PAYMENT_METHODS)
    tx_id = f"TX{random.randint(10000000, 99999999)}"
    reply_markup = get_transaction_keyboard()

    # ধাপ ১
    step_1_text = (
        f"📥 <b>পেমেন্ট উইথড্র রিকোয়েস্ট!</b>\n"
        f"━━━━━━━━━━━━━━━━━━━━━━\n"
        f"👤 <b>কাস্টমার:</b> <code>{customer_name}</code>\n"
        f"💰 <b>পরিমাণ:</b> <code>৳{amount}</code>\n"
        f"🏦 <b>মাধ্যম:</b> <b>{method}</b>\n"
        f"⏳ <b>স্ট্যাটাস:</b> পেন্ডিং (Processing...)"
    )

    try:
        await bot.send_message(
            chat_id=chat_id,
            text=step_1_text,
            reply_markup=reply_markup,
            parse_mode="HTML"
        )
    except Exception as e:
        print(f"Error Step 1: {e}")
        return

    await asyncio.sleep(30)

    # ধাপ ২
    step_2_text = (
        f"🛡️ <b>সিকিউরিটি ভেরিফিকেশন চেক</b>\n"
        f"━━━━━━━━━━━━━━━━━━━━━━\n"
        f"🔍 <b>কাস্টমার:</b> <code>{customer_name}</code>\n"
        f"🛡️ <b>সিস্টেম:</b> অ্যান্টি-হ্যাক এবং ডাবল পেমেন্ট চেক করা হচ্ছে...\n"
        f"⚖️ <b>ফলাফল:</b> Verified ✅ (Safe)\n"
        f"📲 <b>একশন:</b> পেমেন্ট গেটওয়েতে পাঠানো হয়েছে।"
    )

    try:
        await bot.send_message(
            chat_id=chat_id,
            text=step_2_text,
            reply_markup=reply_markup,
            parse_mode="HTML"
        )
    except Exception as e:
        print(f"Error Step 2: {e}")
        return

    await asyncio.sleep(30)

    # ধাপ ৩
    step_3_text = (
        f"🏦 <b>পেমেন্ট ডিপার্টমেন্ট (Finance)</b>\n"
        f"━━━━━━━━━━━━━━━━━━━━━━\n"
        f"👤 <b>কাস্টমার:</b> <code>{customer_name}</code>\n"
        f"🏦 <b>ওয়ালেট:</b> <b>{method}</b>\n"
        f"💰 <b>পরিমাণ:</b> <code>৳{amount}</code>\n"
        f"🆔 <b>TrxID:</b> <code>{tx_id}</code>\n"
        f"💵 <b>স্ট্যাটাস:</b> ফান্ড রিলিজ সম্পন্ন হয়েছে ✅\n"
        f"⏰ <b>সময়:</b> {time.strftime('%I:%M:%S %p')}"
    )

    try:
        await bot.send_message(
            chat_id=chat_id,
            text=step_3_text,
            reply_markup=reply_markup,
            parse_mode="HTML"
        )
        print(f"Transaction finished for {customer_name}. Interval sleep...")
    except Exception as e:
        print(f"Error Step 3: {e}")

# --- ট্রানজ্যাকশন লুপ ---
async def send_periodic_transactions(application):
    await asyncio.sleep(5)
    while True:
        if SETTINGS.get("is_tx_active", True):
            try:
                await execute_3_step_transaction(application.bot, TRANSACTION_CHAT_ID)
            except Exception as e:
                print(f"Loop error: {e}")
        
        interval = SETTINGS.get("transaction_interval", 300)
        await asyncio.sleep(interval)

# --- ওয়েলকাম মেসেজ ---
async def send_and_auto_delete_welcome(bot, chat_id, user):
    user_id = user.id
    if user_id in recently_welcomed_users:
        return
    recently_welcomed_users.add(user_id)

    user_link = f'<a href="tg://user?id={user.id}">{user.full_name or "মেম্বার"}</a>'
    template = SETTINGS.get("welcome_text", DEFAULT_SETTINGS["welcome_text"])
    welcome_text = template.replace("{user_link}", user_link)
    delay = SETTINGS.get("welcome_delete_delay", 30)

    try:
        sent_msg = await bot.send_message(chat_id=chat_id, text=welcome_text, parse_mode="HTML")
        await asyncio.sleep(delay)
        await bot.delete_message(chat_id=chat_id, message_id=sent_msg.message_id)
    except Exception as e:
        print(f"Welcome Error: {e}")
    finally:
        await asyncio.sleep(60)
        recently_welcomed_users.discard(user_id)

async def handle_new_chat_members(update: Update, context: ContextTypes.DEFAULT_TYPE):
    chat_id = str(update.effective_chat.id)
    if chat_id != WELCOME_CHAT_ID:
        return
    for member in update.message.new_chat_members:
        if not member.is_bot:
            asyncio.create_task(send_and_auto_delete_welcome(context.bot, chat_id, member))

async def handle_chat_member_updated(update: Update, context: ContextTypes.DEFAULT_TYPE):
    chat_id = str(update.effective_chat.id)
    if chat_id != WELCOME_CHAT_ID:
        return
    result = update.chat_member
    new_member = result.new_chat_member
    if new_member.status in ["member", "administrator"] and result.old_chat_member.status in ["left", "kicked"]:
        if not new_member.user.is_bot:
            asyncio.create_task(send_and_auto_delete_welcome(context.bot, chat_id, new_member.user))

# --- মাস্টার কন্ট্রোল অ্যাডমিন প্যানেল ---
def get_admin_dashboard_markup():
    is_active = SETTINGS.get("is_tx_active", True)
    status_label = "🔴 ট্রানজ্যাকশন বন্ধ করুন" if is_active else "🟢 ট্রানজ্যাকশন চালু করুন"
    
    keyboard = [
        [InlineKeyboardButton(status_label, callback_data="toggle_tx")],
        [InlineKeyboardButton("⚡ টেস্ট ট্রানজ্যাকশন পাঠান", callback_data="instant_tx")],
        [InlineKeyboardButton("🔘 বাটন তালিকা ও ভিউ", callback_data="view_buttons"),
         InlineKeyboardButton("🗑 বাটন রিসেট", callback_data="reset_buttons")],
        [InlineKeyboardButton("📋 সেটিংস ও কমান্ড হেল্প", callback_data="view_help")]
    ]
    return InlineKeyboardMarkup(keyboard)

def render_dashboard_text():
    is_active = SETTINGS.get("is_tx_active", True)
    status_str = "চালু আছে 🟢" if is_active else "বন্ধ আছে 🔴"
    interval_min = SETTINGS.get("transaction_interval", 300) // 60
    min_a = SETTINGS.get("min_amount", 50.0)
    max_a = SETTINGS.get("max_amount", 2000.0)
    buttons_count = len(SETTINGS.get("buttons", []))

    return (
        f"👑 <b>মাস্টার অ্যাডমিন ড্যাশবোর্ড</b>\n"
        f"━━━━━━━━━━━━━━━━━━━━━━\n"
        f"⚙️ <b>ট্রানজ্যাকশন:</b> {status_str}\n"
        f"⏱️ <b>বিরতি:</b> প্রতি {interval_min} মিনিট পর পর\n"
        f"💰 <b>ব্যালেন্স লিমিট:</b> ৳{min_a} - ৳{max_a}\n"
        f"🔘 <b>সক্রিয় বাটন:</b> {buttons_count} টি\n\n"
        f"<i>নিচের মেনু বা কমান্ডের মাধ্যমে যেকোনো কিছু পরিবর্তন করুন।</i>"
    )

# প্রাইভেট মেসেজ ও কমান্ড হ্যান্ডলার
async def handle_private_admin(update: Update, context: ContextTypes.DEFAULT_TYPE):
    global AUTHORIZED_ADMIN_ID
    if update.effective_chat.type != "private":
        return

    sender_id = update.effective_user.id
    text = (update.message.text or "").strip()

    # ১. সিক্রেট কি অথেনটিকেশন
    if text == SECRET_ADMIN_KEY:
        AUTHORIZED_ADMIN_ID = sender_id
        await update.message.reply_text(
            f"✅ <b>অথেনটিকেশন সফল!</b> আইডি: <code>{sender_id}</code> মাস্টার অ্যাডমিন হিসেবে সেভ হয়েছে।\n\n"
            + render_dashboard_text(),
            reply_markup=get_admin_dashboard_markup(),
            parse_mode="HTML"
        )
        return

    # শুধুমাত্র মাস্টার অ্যাডমিনের জন্য
    if not AUTHORIZED_ADMIN_ID or sender_id != AUTHORIZED_ADMIN_ID:
        await update.message.reply_text("🔒 <i>অ্যাক্সেস সংরক্ষিত।</i>", parse_mode="HTML")
        return

    # ২. ডাইনামিক বাটন অ্যাড করার কমান্ড: /addbtn Name | URL
    if text.startswith("/addbtn"):
        parts = text.replace("/addbtn", "").strip().split("|")
        if len(parts) == 2:
            name = parts[0].strip()
            url = parts[1].strip()
            if not url.startswith("http"):
                await update.message.reply_text("❌ লিংকে অবশ্যই https:// থাকতে হবে!")
                return
            SETTINGS.setdefault("buttons", []).append({"name": name, "url": url})
            save_settings(SETTINGS)
            await update.message.reply_text(f"✅ নতুন বাটন যোগ করা হয়েছে:\n<b>{name}</b> -> {url}", parse_mode="HTML")
        else:
            await update.message.reply_text("❌ ফরম্যাট সঠিক নয়!\nব্যবহার করুন: <code>/addbtn বাটনের নাম | https://link...</code>", parse_mode="HTML")
        return

    # ৩. বিরতি সেট করার কমান্ড: /setinterval <মিনিট>
    if text.startswith("/setinterval"):
        parts = text.split()
        if len(parts) == 2 and parts[1].isdigit():
            mins = int(parts[1])
            SETTINGS["transaction_interval"] = mins * 60
            save_settings(SETTINGS)
            await update.message.reply_text(f"✅ ট্রানজ্যাকশন বিরতি <b>{mins} মিনিট</b> সেট করা হয়েছে!", parse_mode="HTML")
        else:
            await update.message.reply_text("❌ ব্যবহার করুন: <code>/setinterval 5</code> (মিনিটে)", parse_mode="HTML")
        return

    # ৪. ব্যালেন্স রেঞ্জ সেট: /setamount <min> <max>
    if text.startswith("/setamount"):
        parts = text.split()
        if len(parts) == 3:
            try:
                min_v = float(parts[1])
                max_v = float(parts[2])
                SETTINGS["min_amount"] = min_v
                SETTINGS["max_amount"] = max_v
                save_settings(SETTINGS)
                await update.message.reply_text(f"✅ ব্যালেন্স রেঞ্জ <b>৳{min_v} - ৳{max_v}</b> সেট করা হয়েছে!", parse_mode="HTML")
            except ValueError:
                await update.message.reply_text("❌ সংখ্যা লিখুন! উদাহরণ: <code>/setamount 100 3000</code>", parse_mode="HTML")
        else:
            await update.message.reply_text("❌ ব্যবহার করুন: <code>/setamount 100 2000</code>", parse_mode="HTML")
        return

    # ৫. ওয়েলকাম টেক্সট সেট: /setwelcome <মেসেজ>
    if text.startswith("/setwelcome"):
        new_w = text.replace("/setwelcome", "").strip()
        if new_w:
            SETTINGS["welcome_text"] = new_w
            save_settings(SETTINGS)
            await update.message.reply_text("✅ নতুন ওয়েলকাম মেসেজ সেট করা হয়েছে!", parse_mode="HTML")
        else:
            await update.message.reply_text("❌ ব্যবহার করুন: <code>/setwelcome আপনার মেসেজ (নামের জন্য {user_link} ব্যবহার করুন)</code>", parse_mode="HTML")
        return

    # সাধারণ চ্যাট করলে ড্যাশবোর্ড ওপেন হবে
    await update.message.reply_text(
        render_dashboard_text(),
        reply_markup=get_admin_dashboard_markup(),
        parse_mode="HTML"
    )

# বাটন কলব্যাক হ্যান্ডলার
async def admin_callback(update: Update, context: ContextTypes.DEFAULT_TYPE):
    global AUTHORIZED_ADMIN_ID
    query = update.callback_query
    await query.answer()

    if AUTHORIZED_ADMIN_ID and query.from_user.id != AUTHORIZED_ADMIN_ID:
        await query.message.reply_text("⛔ আপনি অনুমোদিত অ্যাডমিন নন!")
        return

    data = query.data

    if data == "toggle_tx":
        SETTINGS["is_tx_active"] = not SETTINGS.get("is_tx_active", True)
        save_settings(SETTINGS)
        await query.edit_message_text(
            render_dashboard_text(),
            reply_markup=get_admin_dashboard_markup(),
            parse_mode="HTML"
        )
    elif data == "instant_tx":
        await query.message.reply_text("⚡ ২য় গ্রুপে ৩ ধাপের ট্রানজ্যাকশন শুরু করা হয়েছে...")
        asyncio.create_task(execute_3_step_transaction(context.bot, TRANSACTION_CHAT_ID))
    elif data == "view_buttons":
        btn_list = SETTINGS.get("buttons", [])
        if not btn_list:
            text = "বর্তমানে কোনো বাটন সক্রিয় নেই।"
        else:
            text = "🔘 <b>বর্তমান সক্রিয় বাটনসমূহ:</b>\n\n"
            for idx, b in enumerate(btn_list, 1):
                text += f"{idx}. <b>{b['name']}</b> -> {b['url']}\n"
        await query.message.reply_text(text, parse_mode="HTML")
    elif data == "reset_buttons":
        SETTINGS["buttons"] = [
            {"name": "FAST GMAIL SELL", "url": "https://t.me/NEW_FRESH_GMAILACCOUNTSELL50_bot"}
        ]
        save_settings(SETTINGS)
        await query.message.reply_text("✅ বাটন রিসেট করে শুধুমাত্র ডিফল্ট <b>FAST GMAIL SELL</b> বাটন রাখা হয়েছে।", parse_mode="HTML")
    elif data == "view_help":
        help_msg = (
            "🛠 <b>অ্যাডমিন কমান্ড তালিকা:</b>\n"
            "━━━━━━━━━━━━━━━━━━━━━━\n"
            "➕ <b>বাটন যোগ করতে:</b>\n"
            "<code>/addbtn বাটনের নাম | https://লিংক</code>\n\n"
            "⏱️ <b>টাইমার পরিবর্তন করতে:</b>\n"
            "<code>/setinterval 5</code> (মিনিটে)\n\n"
            "💰 <b>ব্যালেন্স রেঞ্জ পরিবর্তন করতে:</b>\n"
            "<code>/setamount 100 3000</code>\n\n"
            "🌸 <b>ওয়েলকাম মেসেজ পরিবর্তন করতে:</b>\n"
            "<code>/setwelcome স্বাগতম {user_link} আমাদের গ্রুপে!</code>"
        )
        await query.message.reply_text(help_msg, parse_mode="HTML")

# --- ইনিশিয়ালাইজেশন ---
async def post_init(application):
    asyncio.create_task(send_periodic_transactions(application))
    asyncio.create_task(start_dummy_web_server())

def main():
    print("বট চালু হচ্ছে...")
    app = ApplicationBuilder().token(BOT_TOKEN).post_init(post_init).build()

    # প্রাইভেট মেসেজ ও অ্যাডমিন হ্যান্ডলার
    app.add_handler(MessageHandler(filters.ChatType.PRIVATE & filters.TEXT, handle_private_admin))
    app.add_handler(CallbackQueryHandler(admin_callback))

    # জয়েনিং ওয়েলকাম হ্যান্ডলার
    app.add_handler(MessageHandler(filters.StatusUpdate.NEW_CHAT_MEMBERS, handle_new_chat_members))
    app.add_handler(ChatMemberHandler(handle_chat_member_updated, ChatMemberHandler.CHAT_MEMBER))

    app.run_polling(allowed_updates=Update.ALL_TYPES)

if __name__ == "__main__":
    main()
