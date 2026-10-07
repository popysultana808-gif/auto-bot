import asyncio
import os
import random
import time
from telegram import Update
from telegram.ext import ApplicationBuilder, ChatMemberHandler, ContextTypes

# আপনার বটের টোকেন
BOT_TOKEN = "8826168593:AAEobfC2UHJKtDv9XmvcMc1CmOviAVbloQQ"

# আপনার দেওয়া দুটি গ্রুপের আইডি তালিকাভুক্ত করা হয়েছে
TARGET_CHAT_IDS = [
    "-1004424049305",
    "-1003991468184",
]

INTERVAL_SECONDS = 300  # প্রতি ৫ মিনিট (৩০০ সেকেন্ড) পরপর পোস্ট হবে

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

# --- ২. ব্যাকগ্রাউন্ডে দুটি গ্রুপেই স্বয়ংক্রিয়ভাবে মেসেজ পাঠানো ---
async def send_periodic_transactions(application):
    await asyncio.sleep(5)  # বট চালুর ৫ সেকেন্ড পর কার্যক্রম শুরু হবে
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
            print("Transactions sent successfully to target groups.")
        except Exception as e:
            print(f"Unexpected error in loop: {e}")
            
        await asyncio.sleep(INTERVAL_SECONDS)

# --- ৩. নতুন মেম্বার যুক্ত হলে স্বাগতম জানানো ---
async def welcome_new_member(update: Update, context: ContextTypes.DEFAULT_TYPE):
    result = update.chat_member
    new_member = result.new_chat_member

    # কেউ গ্রুপে নতুন যোগ দিলে বা ইনভাইট পেলে
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
        print(f"Welcomed: {user_name} in chat: {update.effective_chat.id}")

# --- ৪. ব্যাকগ্রাউন্ড লুপ চালু করা ---
async def post_init(application):
    asyncio.create_task(send_periodic_transactions(application))

def main():
    print("বট চালু হচ্ছে...")
    app = ApplicationBuilder().token(BOT_TOKEN).post_init(post_init).build()
    app.add_handler(ChatMemberHandler(welcome_new_member, ChatMemberHandler.CHAT_MEMBER))
    app.run_polling(allowed_updates=Update.ALL_TYPES)

if __name__ == "__main__":
    main()
