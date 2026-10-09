import asyncio
import os
import random
import time
from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup
from telegram.ext import (
    ApplicationBuilder, 
    ChatMemberHandler, 
    MessageHandler,
    CallbackQueryHandler, 
    filters,
    ContextTypes
)

# --- কনফিগারেশন ---
BOT_TOKEN = "8826168593:AAEfgcSn0Ckte8Ayv3OKfzAPic43FIFdT_g"

# আপনার জিমেইল সেল বটের লিংক
GMAIL_BOT_URL = "https://t.me/NEW_FRESH_GMAILACCOUNTSELL50_bot"

# অ্যাডমিন অ্যাক্সেসের সিক্রেট পাসওয়ার্ড
SECRET_ADMIN_KEY = "talha#secret99"
AUTHORIZED_ADMIN_ID = None

# গ্রুপ আইডিসমূহ
WELCOME_CHAT_ID = "-1004471047712"     # ১ম গ্রুপ: ওয়েলকাম মেসেজ
TRANSACTION_CHAT_ID = "-1003991468184"  # ২য় গ্রুপ: ট্রানজ্যাকশন

# প্রতি ট্রানজ্যাকশন সাইকেলের বিরতি: ৫ মিনিট (৩০০ সেকেন্ড)
TRANSACTION_INTERVAL = 300

is_tx_active = True

# বাংলাদেশি কাস্টমার নামসমূহ
NAMES = [
    "Arif Hasan", "Tanvir Ahmed", "Sakib Al Hasan", "Rahim Uddin", 
    "Mehedi Hasan", "Nusrat Jahan", "Sadia Islam", "Farhana Akter", 
    "Jannatul Ferdous", "Ayesha Siddiqua", "Rifat Hossain", "Sumaiya Kabir", 
    "Shakil Khan", "Fahim Shahriar", "Mim Akter", "Tasnim Sultana",
    "Mahmudul Hasan", "Naimur Rahman", "Sabiha Sultana", "Habibur Rahman"
]

PAYMENT_METHODS = ["Bkash", "Nagad", "Binance"]

# ডুপ্লিকেট ওয়েলকাম মেসেজ রোধ করতে রিসেন্টলি ওয়েলকাম দেওয়া ইউজারদের সেট
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

# --- জিমেইল সেল বট বাটন ---
def get_action_keyboard():
    keyboard = [
        [InlineKeyboardButton("🤖 জিমেইল সেল বট", url=GMAIL_BOT_URL)]
    ]
    return InlineKeyboardMarkup(keyboard)

# --- ৩০ সেকেন্ড পর পর ৩টি আলাদা ট্রানজ্যাকশন মেসেজ পাঠানোর প্রসেস ---
async def execute_3_step_transaction(bot, chat_id):
    customer_name = random.choice(NAMES)
    amount = f"{random.uniform(50.00, 2000.00):.2f}"
    method = random.choice(PAYMENT_METHODS)
    tx_id = f"TX{random.randint(10000000, 99999999)}"
    reply_markup = get_action_keyboard()

    # ধাপ ১: ১ম মেসেজ - পেমেন্ট উইথড্র রিকোয়েস্ট
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

    # ৩০ সেকেন্ড অপেক্ষা
    await asyncio.sleep(30)

    # ধাপ ২: ২য় মেসেজ - সিকিউরিটি ভেরিফিকেশন চেক
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

    # আবার ৩০ সেকেন্ড অপেক্ষা
    await asyncio.sleep(30)

    # ধাপ ৩: ৩য় মেসেজ - পেমেন্ট ডিপার্টমেন্ট (Finance)
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
        print(f"Transaction cycle completed for {customer_name}")
    except Exception as e:
        print(f"Error Step 3: {e}")

# --- প্রতি ৫ মিনিট পর পর ট্রানজ্যাকশন লুপ ---
async def send_periodic_transactions(application):
    await asyncio.sleep(5)
    while True:
        if is_tx_active:
            try:
                # ব্যাকগ্রাউন্ডে ৩টি মেসেজের প্রসেস চলবে
                asyncio.create_task(execute_3_step_transaction(application.bot, TRANSACTION_CHAT_ID))
            except Exception as e:
                print(f"Transaction loop error: {e}")
        
        # প্রতি ৫ মিনিট পর পর পরবর্তী ট্রানজ্যাকশন রিকোয়েস্ট শুরু হবে
        await asyncio.sleep(TRANSACTION_INTERVAL)

# --- একক ও নিখুঁত ওয়েলকাম মেসেজ (৩০ সেকেন্ডে ডিলিট) ---
async def send_and_auto_delete_welcome(bot, chat_id, user):
    user_id = user.id
    # ডুপ্লিকেট মেসেজ রোধ করার ফিল্টার
    if user_id in recently_welcomed_users:
        return
    recently_welcomed_users.add(user_id)

    user_name = user.full_name or "মেম্বার"
    welcome_text = (
        f"🌸 <b>আসসালামু আলাইকুম</b>, <a href=\"tg://user?id={user.id}\">{user_name}</a>!\n\n"
        f"আমাদের কমিউনিটিতে আপনাকে স্বাগতম। 🎉\n"
        f"📌 গ্রুপের নিয়ম-কানুন মেনে চলুন এবং নিয়মিত আপডেট উপভোগ করুন।"
    )
    try:
        sent_msg = await bot.send_message(
            chat_id=chat_id,
            text=welcome_text,
            parse_mode="HTML"
        )
        print(f"Welcome sent to {user_name}. Deleting in 30s...")
        await asyncio.sleep(30)
        await bot.delete_message(chat_id=chat_id, message_id=sent_msg.message_id)
        print("Welcome message deleted.")
    except Exception as e:
        print(f"Welcome/Delete Error: {e}")
    finally:
        # ৬০ সেকেন্ড পর মেমোরি থেকে ইউজার আইডি ক্লিয়ার করা
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

# --- অ্যাডমিন কন্ট্রোল প্যানেল ---
def get_admin_keyboard():
    status_label = "🔴 ট্রানজ্যাকশন বন্ধ করুন" if is_tx_active else "🟢 ট্রানজ্যাকশন চালু করুন"
    keyboard = [
        [InlineKeyboardButton(status_label, callback_data="toggle_tx")],
        [InlineKeyboardButton("⚡ এখনই একটি ৩-ধাপের ট্রানজ্যাকশন পাঠান", callback_data="instant_tx")]
    ]
    return InlineKeyboardMarkup(keyboard)

async def handle_private_messages(update: Update, context: ContextTypes.DEFAULT_TYPE):
    global AUTHORIZED_ADMIN_ID
    if update.effective_chat.type != "private":
        return

    sender_id = update.effective_user.id
    text = update.message.text.strip() if update.message.text else ""

    if text == SECRET_ADMIN_KEY:
        AUTHORIZED_ADMIN_ID = sender_id
        status_text = "চালু আছে 🟢" if is_tx_active else "বন্ধ আছে 🔴"
        panel_text = (
            f"👑 <b>মাস্টার অ্যাডমিন লগইন সফল!</b>\n\n"
            f"আপনার আইডি (<code>{sender_id}</code>) মাস্টার অ্যাডমিন হিসেবে সংরক্ষিত।\n"
            f"⚙️ ট্রানজ্যাকশন স্ট্যাটাস: <b>{status_text}</b>\n"
            f"⏱️ ইন্টারভ্যাল: প্রতি ৫ মিনিট পর পর\n\n"
            f"নিচের বাটন দিয়ে নিয়ন্ত্রণ করুন:"
        )
        await update.message.reply_text(panel_text, reply_markup=get_admin_keyboard(), parse_mode="HTML")
        return

    if AUTHORIZED_ADMIN_ID and sender_id == AUTHORIZED_ADMIN_ID:
        status_text = "চালু আছে 🟢" if is_tx_active else "বন্ধ আছে 🔴"
        panel_text = (
            f"🎛 <b>কন্ট্রোল অ্যাডমিন প্যানেল</b>\n\n"
            f"⚙️ বর্তমান ট্রানজ্যাকশন স্ট্যাটাস: <b>{status_text}</b>\n\n"
            f"নিচের বাটন চেপে নিয়ন্ত্রণ করুন:"
        )
        await update.message.reply_text(panel_text, reply_markup=get_admin_keyboard(), parse_mode="HTML")
        return

    await update.message.reply_text("🔒 <i>অ্যাক্সেস সংরক্ষিত।</i>", parse_mode="HTML")

async def admin_callback(update: Update, context: ContextTypes.DEFAULT_TYPE):
    global is_tx_active
    query = update.callback_query
    await query.answer()

    if AUTHORIZED_ADMIN_ID and query.from_user.id != AUTHORIZED_ADMIN_ID:
        await query.message.reply_text("⛔ আপনি এই প্যানেল ব্যবহারের অনুমতিপ্রাপ্ত নন!")
        return

    if query.data == "toggle_tx":
        is_tx_active = not is_tx_active
        status_text = "চালু আছে 🟢" if is_tx_active else "বন্ধ আছে 🔴"
        panel_text = (
            f"🎛 <b>কন্ট্রোল অ্যাডমিন প্যানেল</b>\n\n"
            f"⚙️ বর্তমান ট্রানজ্যাকশন স্ট্যাটাস: <b>{status_text}</b>\n\n"
            f"নিচের বাটন চেপে নিয়ন্ত্রণ করুন:"
        )
        await query.edit_message_text(panel_text, reply_markup=get_admin_keyboard(), parse_mode="HTML")
        
    elif query.data == "instant_tx":
        await query.message.reply_text("⚡ ২য় গ্রুপে ৩ ধাপের আলাদা ট্রানজ্যাকশন মেসেজ পাঠানো শুরু হয়েছে (৩০ সেকেন্ড অন্তর অন্তর)...")
        asyncio.create_task(execute_3_step_transaction(context.bot, TRANSACTION_CHAT_ID))

# --- সার্ভিস শুরু ---
async def post_init(application):
    asyncio.create_task(send_periodic_transactions(application))
    asyncio.create_task(start_dummy_web_server())

def main():
    print("বট চালু হচ্ছে...")
    app = ApplicationBuilder().token(BOT_TOKEN).post_init(post_init).build()

    # প্রাইভেট মেসেজ ও অ্যাডমিন হ্যান্ডলার
    app.add_handler(MessageHandler(filters.ChatType.PRIVATE & filters.TEXT, handle_private_messages))
    app.add_handler(CallbackQueryHandler(admin_callback))

    # জয়েনিং হ্যান্ডলারসমূহ
    app.add_handler(MessageHandler(filters.StatusUpdate.NEW_CHAT_MEMBERS, handle_new_chat_members))
    app.add_handler(ChatMemberHandler(handle_chat_member_updated, ChatMemberHandler.CHAT_MEMBER))

    app.run_polling(allowed_updates=Update.ALL_TYPES)

if __name__ == "__main__":
    main()
