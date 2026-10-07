
import asyncio
import os
import random
import time
from telegram import Update
from telegram.ext import ApplicationBuilder, ChatMemberHandler, ContextTypes

# আপনার বটের টোকেন এখানে বসান
BOT_TOKEN = "YOUR_BOT_TOKEN_HERE"

# আপনার গ্রুপ আইডি এখানে বসান (যেমন: "-1001234567890")
TARGET_CHAT_IDS = [
    "-100XXXXXXXXXX",
]

INTERVAL_SECONDS = 300  # ৫ মিনিট পর পর ট্রানজ্যাকশন পোস্ট হবে

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

async def send_periodic_transactions(application):
    await asyncio.sleep(5)
    while True:
        try:
            msg = get_transaction_data()
            for chat_id in TARGET_CHAT_IDS:
                await application.bot.send_message(
                    chat_id=chat_id,
                    text=msg,
                    parse_mode="Markdown"
                )
            print("Transactions sent successfully.")
        except Exception as e:
            print(f"Error sending transaction: {e}")
            
        await asyncio.sleep(INTERVAL_SECONDS)

async def welcome_new_member(update: Update, context: ContextTypes.DEFAULT_TYPE):
    result = update.chat_member
    new_member = result.new_chat_member

    if new_member.status in ["member", "administrator"] and result.old_chat_member.status in ["left", "kicked"]:
        user = new_member.user
        user_name = user.full_name
        
        welcome_text = (
            f"👋 স্বাগতম, [{user_name}](tg://user?id={user.id}) আমাদের গ্রুপে!\n\n"
            f"📌 গ্রুপে নিয়মিত আপডেট ও লেনদেন সংক্রান্ত তথ্য শেয়ার করা হয়।"
        )
        
        await context.bot.send_message(
            chat_id=update.effective_chat.id,
            text=welcome_text,
            parse_mode="Markdown"
        )
        print(f"Welcomed: {user_name}")

async def post_init(application):
    asyncio.create_task(send_periodic_transactions(application))

def main():
    print("বট চালু হচ্ছে...")
    app = ApplicationBuilder().token(BOT_TOKEN).post_init(post_init).build()
    app.add_handler(ChatMemberHandler(welcome_new_member, ChatMemberHandler.CHAT_MEMBER))
    app.run_polling(allowed_updates=Update.ALL_TYPES)

if __name__ == "__main__":
    main()
