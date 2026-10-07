import asyncio
import os
import random
import time
from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup
from telegram.ext import (
    ApplicationBuilder, 
    ChatMemberHandler, 
    CommandHandler, 
    CallbackQueryHandler, 
    ContextTypes
)

# বটের কনফিগারেশন
BOT_TOKEN = "8826168593:AAEobfC2UHJKtDv9XmvcMc1CmOviAVbloQQ"

# আপনার সাপোর্ট ইউজারনেম
SUPPORT_USERNAME = "@Talha_juba098"

# ফার্স্ট গ্রুপ (শুধুমাত্র ওয়েলকাম মেসেজ চলবে)
WELCOME_CHAT_ID = "-1004424049305"

# সেকেন্ড গ্রুপ (শুধুমাত্র ট্রানজ্যাকশন চলবে)
TRANSACTION_CHAT_ID = "-1003991468184"

# আপনার টেলিগ্রাম আইডি (প্রয়োজন হলে বসাতে পারেন, না দিলেও সমস্যা নেই)
ADMIN_USER_ID = None  

# রিসাইকেল টাইম লিস্ট (১, ২, ৩, ৪, ৫ মিনিট)
CYCLE_INTERVALS = [60, 120, 180, 240, 300]

# অ্যাডমিন কন্ট্রোল ফ্ল্যাগ
is_tx_active = True

# বাংলাদেশি নামসমূহ
NAMES = [
    "Arif Hasan", "Tanvir Ahmed", "Sakib Al Hasan", "Rahim Uddin", 
    "Mehedi Hasan", "Nusrat Jahan", "Sadia Islam", "Farhana Akter", 
    "Jannatul Ferdous", "Ayesha Siddiqua", "Rifat Hossain", "Sumaiya Kabir", 
    "Shakil Khan", "Fahim Shahriar", "Mim Akter", "Tasnim Sultana",
    "Mahmudul Hasan", "Naimur Rahman", "Sabiha Sultana", "Habibur Rahman"
]

# --- Render পোর্টের জন্য ওয়েব সার্ভার ---
async def start_dummy_web_server():
    port = int(os.environ.get("PORT", 10000))
    async def handle_ping(reader, writer):
        response = b"HTTP/1.1 200 OK\r\nContent-Type: text/plain\r\nContent-Length: 2\r\n\r\nOK"
        writer.write(response)
        await writer.drain()
        writer.close()

    server = await asyncio.start_server(handle_ping, "0.0.0.0", port)
    print(f"Web server running on port {port}")
    async with server:
        await server.serve_forever()

# --- ১. স্বয়ংক্রিয় ট্রানজ্যাকশন জেনারেটর (সাপোর্ট আইডি সহ) ---
def get_transaction_data():
    tx_id = f"TX{random.randint(10000000, 99999999)}"
    
    # ৫০ থেকে ২০০০ টাকার মধ্যে নির্দিষ্ট কিছু সাধারণ স্ল্যাব বা র‍্যান্ডম ভ্যালু
    common_amounts = [50, 100, 150, 200, 300, 450, 500, 800, 900, 1000, 1200, 1300, 1500, 1800, 2000]
    if random.random() < 0.6:
        amount = random.choice(common_amounts)
    else:
        amount = random.randint(50, 2000)
        
    user = random.choice(NAMES)
    status = "SUCCESSFUL ✅"
    
    return (
        f"🔔 *New Transaction Completed!*\n\n"
        f"👤 *Customer:* `{user}`\n"
        f"💰 *Amount:* `৳ {amount:,} BDT`\n"
        f"🆔 *TrxID:* `{tx_id}`\n"
        f"📊 *Status:* {status}\n"
        f"⏰ *Time:* {time.strftime('%I:%M:%S %p')}\n\n"
        f"💬 *Support:* {SUPPORT_USERNAME}"
    )

# --- ২. সেকেন্ড গ্রুপে রিসাইকেল টাইমে ট্রানজ্যাকশন পাঠানো ---
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
                        parse_mode="Markdown"
                    )
                    print(f"Transaction sent to 2nd group after {wait_seconds // 60} min.")
                except Exception as e:
                    print(f"Error sending transaction: {e}")

# --- ৩. ফার্স্ট গ্রুপে ওয়েলকাম মেসেজ এবং ৩০ সেকেন্ড পর ডিলিট ---
async def welcome_new_member(update: Update, context: ContextTypes.DEFAULT_TYPE):
    result = update.chat_member
    chat_id = str(update.effective_chat.id)
    
    # শুধুমাত্র ফার্স্ট গ্রুপে ওয়েলকাম মেসেজ যাবে
    if chat_id != WELCOME_CHAT_ID:
        return

    new_member = result.new_chat_member
    if new_member.status in ["member", "administrator"] and result.old_chat_member.status in ["left", "kicked"]:
        user = new_member.user
        user_name = user.full_name
        
        welcome_text = (
            f"🌸 *আসসালামু আলাইকুম*, [{user_name}](tg://user?id={user.id})!\n\n"
            f"আমাদের কমিউনিটিতে আপনাকে স্বাগতম। 🎉\n"
            f"📌 গ্রুপের নিয়ম-কানুন মেনে চলুন এবং নিয়মিত আপডেট উপভোগ করুন।"
        )
        
        try:
            sent_msg = await context.bot.send_message(
                chat_id=chat_id,
                text=welcome_text,
                parse_mode="Markdown"
            )
            print(f"Welcome sent to {user_name}, will delete in 30s...")

            # ৩০ সেকেন্ড অপেক্ষা করে মেসেজটি স্বয়ংক্রিয়ভাবে মুছে ফেলা
            await asyncio.sleep(30)
            await context.bot.delete_message(
                chat_id=chat_id,
                message_id=sent_msg.message_id
            )
            print("Welcome message deleted.")
        except Exception as e:
            print(f"Error handling welcome/delete message: {e}")

# --- ৪. অ্যাডমিন প্যানেল কমান্ড ও বাটন কন্ট্রোল ---
async def admin_panel(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if ADMIN_USER_ID and update.effective_user.id != ADMIN_USER_ID:
        await update.message.reply_text("⛔ আপনি অ্যাডমিন নন!")
        return

    status_text = "চালু (Active) 🟢" if is_tx_active else "বন্ধ (Paused) 🔴"
    keyboard = [
        [
            InlineKeyboardButton("ট্রানজ্যাকশন অন/অফ 🔄", callback_data="toggle_tx"),
            InlineKeyboardButton("এখনই একটি ট্রানজ্যাকশন পাঠান ⚡", callback_data="instant_tx")
        ]
    ]
    reply_markup = InlineKeyboardMarkup(keyboard)

    await update.message.reply_text(
        f"🛠 *বট কন্ট্রোল প্যানেল*\n\n"
        f"📊 বর্তমান ট্রানজ্যাকশন স্ট্যাটাস: *{status_text}*\n"
        f"📌 সাপোর্ট ইউজারনেম: {SUPPORT_USERNAME}",
        reply_markup=reply_markup,
        parse_mode="Markdown"
    )

async def admin_button_callback(update: Update, context: ContextTypes.DEFAULT_TYPE):
    global is_tx_active
    query = update.callback_query
    await query.answer()

    if ADMIN_USER_ID and query.from_user.id != ADMIN_USER_ID:
        return

    if query.data == "toggle_tx":
        is_tx_active = not is_tx_active
        status_text = "চালু (Active) 🟢" if is_tx_active else "বন্ধ (Paused) 🔴"
        keyboard = [
            [
                InlineKeyboardButton("ট্রানজ্যাকশন অন/অফ 🔄", callback_data="toggle_tx"),
                InlineKeyboardButton("এখনই একটি ট্রানজ্যাকশন পাঠান ⚡", callback_data="instant_tx")
            ]
        ]
        await query.edit_message_text(
            f"🛠 *বট কন্ট্রোল প্যানেল*\n\n"
            f"📊 বর্তমান ট্রানজ্যাকশন স্ট্যাটাস: *{status_text}*",
            reply_markup=InlineKeyboardMarkup(keyboard),
            parse_mode="Markdown"
        )
    elif query.data == "instant_tx":
        msg = get_transaction_data()
        try:
            await context.bot.send_message(
                chat_id=TRANSACTION_CHAT_ID,
                text=msg,
                parse_mode="Markdown"
            )
            await query.message.reply_text("✅ সেকেন্ড গ্রুপে ইনস্ট্যান্ট ট্রানজ্যাকশন পাঠানো হয়েছে!")
        except Exception as e:
            await query.message.reply_text(f"❌ ব্যর্থ হয়েছে: {e}")

# --- ৫. বট রানার ---
async def post_init(application):
    asyncio.create_task(send_periodic_transactions(application))
    asyncio.create_task(start_dummy_web_server())

def main():
    print("বট চালু হচ্ছে...")
    app = ApplicationBuilder().token(BOT_TOKEN).post_init(post_init).build()

    # হ্যান্ডলারসমূহ যুক্ত করা
    app.add_handler(CommandHandler("admin", admin_panel))
    app.add_handler(CallbackQueryHandler(admin_button_callback))
    app.add_handler(ChatMemberHandler(welcome_new_member, ChatMemberHandler.CHAT_MEMBER))

    app.run_polling(allowed_updates=Update.ALL_TYPES)

if __name__ == "__main__":
    main()
