import asyncio
import os
import random
import time
from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup
from telegram.ext import (
    ApplicationBuilder, 
    ChatMemberHandler, 
    MessageHandler,
    CommandHandler, 
    CallbackQueryHandler, 
    filters,
    ContextTypes
)

# --- কনফিগারেশন ---
BOT_TOKEN = "8826168593:AAEobfC2UHJKtDv9XmvcMc1CmOviAVbloQQ"
SUPPORT_USERNAME = "@Talha_juba098"

# অ্যাডমিন প্যানেল খোলার পাসওয়ার্ড
ADMIN_PASSWORD = "talha1234"

# গ্রুপ আইডিসমূহ
WELCOME_CHAT_ID = "-1004471047712"     # ১ম গ্রুপ: যেখানে ওয়েলকাম মেসেজ যাবে
TRANSACTION_CHAT_ID = "-1003991468184"  # ২য় গ্রুপ: যেখানে ট্রানজ্যাকশন পোস্ট হবে

# রিসাইকেল টাইমার লিস্ট (১, ২, ৩, ৪, ৫ মিনিট)
CYCLE_INTERVALS = [60, 120, 180, 240, 300]

is_tx_active = True

NAMES = [
    "Arif Hasan", "Tanvir Ahmed", "Sakib Al Hasan", "Rahim Uddin", 
    "Mehedi Hasan", "Nusrat Jahan", "Sadia Islam", "Farhana Akter", 
    "Jannatul Ferdous", "Ayesha Siddiqua", "Rifat Hossain", "Sumaiya Kabir", 
    "Shakil Khan", "Fahim Shahriar", "Mim Akter", "Tasnim Sultana",
    "Mahmudul Hasan", "Naimur Rahman", "Sabiha Sultana", "Habibur Rahman"
]

PAYMENT_METHODS = [
    "bKash (Personal)",
    "Nagad",
    "Binance (USDT/Pay)"
]

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
    print(f"Server active on port {port}")
    async with server:
        await server.serve_forever()

# --- ১. ট্রানজ্যাকশন মেসেজ তৈরি ---
def get_transaction_data():
    tx_id = f"TX{random.randint(10000000, 99999999)}"
    common_amounts = [50, 100, 150, 200, 300, 450, 500, 800, 900, 1000, 1200, 1300, 1500, 1800, 2000]
    if random.random() < 0.6:
        amount = random.choice(common_amounts)
    else:
        amount = random.randint(50, 2000)
        
    user = random.choice(NAMES)
    method = random.choice(PAYMENT_METHODS)
    status = "SUCCESSFUL ✅"
    
    return (
        f"🔔 <b>New Withdrawal Completed!</b>\n\n"
        f"👤 <b>Customer:</b> <code>{user}</code>\n"
        f"💳 <b>Method:</b> <code>{method}</code>\n"
        f"💰 <b>Amount:</b> <code>৳ {amount:,} BDT</code>\n"
        f"🆔 <b>TrxID:</b> <code>{tx_id}</code>\n"
        f"📊 <b>Status:</b> {status}\n"
        f"⏰ <b>Time:</b> {time.strftime('%I:%M:%S %p')}\n\n"
        f"💬 <b>Support:</b> {SUPPORT_USERNAME}"
    )

# --- ২. স্বয়ংক্রিয় ট্রানজ্যাকশন লুপ ---
async def send_periodic_transactions(application):
    await asyncio.sleep(5)
    while True:
        for wait_seconds in CYCLE_INTERVALS:
            await asyncio.sleep(wait_seconds)
            if is_tx_active:
                try:
                    msg = get_transaction_data()
                    await application.bot.send_message(
                        chat_id=TRANSACTION_CHAT_ID,
                        text=msg,
                        parse_mode="HTML"
                    )
                    print(f"Transaction sent after {wait_seconds // 60} min.")
                except Exception as e:
                    print(f"Error sending transaction: {e}")

# --- ৩. ওয়েলকাম মেসেজ হ্যান্ডলিং (৩০ সেকেন্ড পর স্বয়ংক্রিয়ভাবে ডিলিট) ---
async def send_and_auto_delete_welcome(bot, chat_id, user):
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
        print(f"Welcome message sent to {user_name}. Deleting in 30s...")
        await asyncio.sleep(30)
        await bot.delete_message(chat_id=chat_id, message_id=sent_msg.message_id)
        print("Welcome message deleted.")
    except Exception as e:
        print(f"Welcome/Delete Error: {e}")

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

# --- ৪. সহজ ও নির্ভরযোগ্য অ্যাডমিন প্যানেল ---
def get_admin_keyboard():
    status_label = "🔴 ট্রানজ্যাকশন বন্ধ করুন" if is_tx_active else "🟢 ট্রানজ্যাকশন চালু করুন"
    keyboard = [
        [InlineKeyboardButton(status_label, callback_data="toggle_tx")],
        [InlineKeyboardButton("⚡ এখনই একটি ট্রানজ্যাকশন পাঠান", callback_data="instant_tx")]
    ]
    return InlineKeyboardMarkup(keyboard)

# প্রাইভেটে যেকোনো মেসেজ বা পাসওয়ার্ড আসলে হ্যান্ডেল করা
async def handle_private_messages(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if update.effective_chat.type != "private":
        return

    text = update.message.text.strip() if update.message.text else ""

    # যদি পাসওয়ার্ড পাঠায় অথবা /admin talha1234 লেখে
    if text == ADMIN_PASSWORD or text == f"/admin {ADMIN_PASSWORD}":
        status_text = "চালু আছে 🟢" if is_tx_active else "বন্ধ আছে 🔴"
        panel_text = (
            f"🎛 <b>কন্ট্রোল অ্যাডমিন প্যানেল (লগইন সফল)</b>\n\n"
            f"⚙️ ট্রানজ্যাকশন স্ট্যাটাস: <b>{status_text}</b>\n"
            f"💳 মেথড: বিকাশ | নগদ | বাইন্যান্স\n"
            f"💬 সাপোর্ট: {SUPPORT_USERNAME}\n\n"
            f"নিচের বাটন চেপে নিয়ন্ত্রণ করুন:"
        )
        await update.message.reply_text(panel_text, reply_markup=get_admin_keyboard(), parse_mode="HTML")
    else:
        await update.message.reply_text(
            "🔒 অ্যাডমিন প্যানেল খুলতে পাসওয়ার্ড লিখুন:\n<code>talha1234</code>",
            parse_mode="HTML"
        )

async def admin_callback(update: Update, context: ContextTypes.DEFAULT_TYPE):
    global is_tx_active
    query = update.callback_query
    await query.answer()

    if query.data == "toggle_tx":
        is_tx_active = not is_tx_active
        status_text = "চালু আছে 🟢" if is_tx_active else "বন্ধ আছে 🔴"
        panel_text = (
            f"🎛 <b>কন্ট্রোল অ্যাডমিন প্যানেল</b>\n\n"
            f"⚙️ ট্রানজ্যাকশন স্ট্যাটাস: <b>{status_text}</b>\n"
            f"💳 মেথড: বিকাশ | নগদ | বাইন্যান্স\n"
            f"💬 সাপোর্ট: {SUPPORT_USERNAME}\n\n"
            f"নিচের বাটন চেপে নিয়ন্ত্রণ করুন:"
        )
        await query.edit_message_text(panel_text, reply_markup=get_admin_keyboard(), parse_mode="HTML")
        
    elif query.data == "instant_tx":
        msg = get_transaction_data()
        try:
            await context.bot.send_message(
                chat_id=TRANSACTION_CHAT_ID,
                text=msg,
                parse_mode="HTML"
            )
            await query.message.reply_text("✅ ২য় গ্রুপে তাৎক্ষণিক ট্রানজ্যাকশন সফলভাবে পাঠানো হয়েছে!")
        except Exception as e:
            await query.message.reply_text(f"❌ পাঠানো সম্ভব হয়নি: {e}")

# --- ৫. সার্ভিস ইনিশিয়ালাইজেশন ---
async def post_init(application):
    asyncio.create_task(send_periodic_transactions(application))
    asyncio.create_task(start_dummy_web_server())

def main():
    print("বট চালু হচ্ছে...")
    app = ApplicationBuilder().token(BOT_TOKEN).post_init(post_init).build()

    # অ্যাডমিন বা পাসওয়ার্ড মেসেজ হ্যান্ডলার
    app.add_handler(MessageHandler(filters.ChatType.PRIVATE & filters.TEXT, handle_private_messages))
    app.add_handler(CallbackQueryHandler(admin_callback))

    # মেম্বার জয়েন হ্যান্ডলারসমূহ
    app.add_handler(MessageHandler(filters.StatusUpdate.NEW_CHAT_MEMBERS, handle_new_chat_members))
    app.add_handler(ChatMemberHandler(handle_chat_member_updated, ChatMemberHandler.CHAT_MEMBER))

    app.run_polling(allowed_updates=Update.ALL_TYPES)

if __name__ == "__main__":
    main()
