import asyncio
import os
import random
import time
from telegram import Update
from telegram.ext import ApplicationBuilder, ChatMemberHandler, ContextTypes

# আপনার বটের টোকেন
BOT_TOKEN = "8826168593:AAEobfC2UHJKtDv9XmvcMc1CmOviAVbloQQ"

# আপনার দুটি গ্রুপের চ্যাট আইডি
TARGET_CHAT_IDS = [
    "-1004424049305",
    "-1003991468184",
]

INTERVAL_SECONDS = 300  # প্রতি ৫ মিনিট পরপর পোস্ট হবে

# --- Render-এর জন্য ডামি ওয়েব সার্ভার (পোর্ট সমস্যা সমাধানের জন্য) ---
async def start_dummy_web_server():
    port = int(os.environ.get("PORT", 10000))

    async def handle_ping(reader, writer):
        response = b"HTTP/1.1 200 OK\r\nContent-Type: text/plain\r\nContent-Length: 2\r\n\r\nOK"
        writer.write(response)
        await writer.drain()
        writer.close()

    server = await asyncio.start_server(handle_ping, "0.0.0.0", port)
    print(f"Web server started on port {port}")
    async with server:
        await server.serve_forever()

# --- ১. স্বয়ংক্রিয় ট্রানজ্যাকশন মেসেজ তৈরি ---
def get_transaction_data():
    tx_id = f"TX{random.randint(100000, 999999)}"
    amount = round(random.uniform(10.0, 500.0), 2)
    user = f"user_{random.randint(100, 999)}"
    status = "SUCCESS ✅"
    
    return (
        f"🔔 *New Transaction Detected!*\n\n"
        f"👤 *User:* `{user}`\n"
        f"💵 *Amount:* `${amount}`\n"
        f"🆔 *TxID:* `{tx_id}`\n"
        f"📊 *Status:* {status}\n"
        f"⏰ *Time:* {time.strftime('%Y-%m-%d %H:%M:%S')}"
    )

# --- ২. ব্যাকগ্রাউন্ডে দুটি গ্রুপেই মেসেজ পাঠানো ---
async def send_periodic_transactions(application):
    await asyncio.sleep(5)
    while True:
        try:
            msg = get_transaction_data()
            for chat_id in TARGET_CHAT_IDS:
                try:
                    await application.bot.send_message(
                        chat_id=chat_id,
                        text=msg,
                        parse_mode="Markdown"
                    )
                except Exception as send_err:
                    print(f"Error sending to group {chat_id}: {send_err}")
            print("Transactions sent successfully.")
        except Exception as e:
            print(f"Loop error: {e}")
            
        await asyncio.sleep(INTERVAL_SECONDS)

# --- ৩. নতুন মেম্বারদের ওয়েলকাম মেসেজ দেওয়া ---
async def welcome_new_member(update: Update, context: ContextTypes.DEFAULT_TYPE):
    result = update.chat_member
    new_member = result.new_chat_member

    if new_member.status in ["member", "administrator"] and result.old_chat_member.status in ["left", "kicked"]:
        user = new_member.user
        user_name = user.full_name
        
        welcome_text = (
            f"👋 স্বাগতম, [{user_name}](tg://user?id={user.id}) আমাদের গ্রুপে!\n\n"
            f"📌 গ্রুপে নিয়মিত আপডেট ও লেনদেন সংক্রান্ত তথ্য শেয়ার করা হয়।\n"
            f"দয়া করে গ্রুপের নিয়ম মেনে চলুন।"
        )
        
        await context.bot.send_message(
            chat_id=update.effective_chat.id,
            text=welcome_text,
            parse_mode="Markdown"
        )
        print(f"Welcomed: {user_name}")

# --- ৪. ব্যাকগ্রাউন্ড টাস্ক চালু করা ---
async def post_init(application):
    # ট্রানজ্যাকশন ও ওয়েব সার্ভার দুটিকেই ব্যাকগ্রাউন্ডে চালু করা
    asyncio.create_task(send_periodic_transactions(application))
    asyncio.create_task(start_dummy_web_server())

def main():
    print("বট চালু হচ্ছে...")
    app = ApplicationBuilder().token(BOT_TOKEN).post_init(post_init).build()
    app.add_handler(ChatMemberHandler(welcome_new_member, ChatMemberHandler.CHAT_MEMBER))
    app.run_polling(allowed_updates=Update.ALL_TYPES)

if __name__ == "__main__":
    main()
