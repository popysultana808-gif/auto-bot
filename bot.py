import asyncio
import os
import random
import time
from telegram import Update
from telegram.ext import ApplicationBuilder, ChatMemberHandler, ContextTypes

# বটের কনফিগারেশন
BOT_TOKEN = "8826168593:AAEobfC2UHJKtDv9XmvcMc1CmOviAVbloQQ"

# যে গ্রুপে শুধুমাত্র ট্রানজ্যাকশন যাবে
TRANSACTION_CHAT_ID = "-1004424049305"

# আপনার রিসাইকেল টাইম লিস্ট (মিনিটকে সেকেন্ডে রূপান্তর: ১, ২, ৩, ৪, ৫ মিনিট)
CYCLE_INTERVALS = [60, 120, 180, 240, 300]

# সুন্দর বাংলাদেশি ছেলে ও মেয়েদের নাম
NAMES = [
    "Arif Hasan", "Tanvir Ahmed", "Sakib Al Hasan", "Rahim Uddin", 
    "Mehedi Hasan", "Nusrat Jahan", "Sadia Islam", "Farhana Akter", 
    "Jannatul Ferdous", "Ayesha Siddiqua", "Rifat Hossain", "Sumaiya Kabir", 
    "Shakil Khan", "Fahim Shahriar", "Mim Akter", "Tasnim Sultana",
    "Mahmudul Hasan", "Naimur Rahman", "Sabiha Sultana", "Habibur Rahman"
]

# --- Render-এর ওয়েব সার্ভার যাতে পোর্ট এরর না দেয় ---
async def start_dummy_web_server():
    port = int(os.environ.get("PORT", 10000))

    async def handle_ping(reader, writer):
        response = b"HTTP/1.1 200 OK\r\nContent-Type: text/plain\r\nContent-Length: 2\r\n\r\nOK"
        writer.write(response)
        await writer.drain()
        writer.close()

    server = await asyncio.start_server(handle_ping, "0.0.0.0", port)
    print(f"Web server active on port {port}")
    async with server:
        await server.serve_forever()

# --- ১. স্বয়ংক্রিয় ট্রানজ্যাকশন মেসেজ জেনারেটর ---
def get_transaction_data():
    tx_id = f"TX{random.randint(10000000, 99999999)}"
    # টাকায় ব্যালেন্স (যেমন: ৫০০ থেকে ২৫,০০০ টাকার মধ্যে পূর্ণসংখ্যা)
    amount = random.randint(500, 25000)
    user = random.choice(NAMES)
    status = "SUCCESSFUL ✅"
    
    return (
        f"🔔 *New Transaction Completed!*\n\n"
        f"👤 *Customer:* `{user}`\n"
        f"💰 *Amount:* `৳ {amount:,} BDT`\n"
        f"🆔 *TrxID:* `{tx_id}`\n"
        f"📊 *Status:* {status}\n"
        f"⏰ *Time:* {time.strftime('%I:%M:%S %p')}"
    )

# --- ২. রিসাইকেল টাইমিং অনুযায়ী ট্রানজ্যাকশন পাঠানোর লুপ ---
async def send_periodic_transactions(application):
    await asyncio.sleep(5)  # বট চালুর ৫ সেকেন্ড পর ব্যাকগ্রাউন্ড প্রসেস শুরু
    
    while True:
        # প্রতি সাইকেলে ১ মিনিট, ২ মিনিট, ৩ মিনিট, ৪ মিনিট, ৫ মিনিট করে চলবে
        for wait_seconds in CYCLE_INTERVALS:
            try:
                # সময় অনুযায়ী অপেক্ষা (১ম বার ১ মিনিট, ২য় বার ২ মিনিট...)
                await asyncio.sleep(wait_seconds)
                
                msg = get_transaction_data()
                await application.bot.send_message(
                    chat_id=TRANSACTION_CHAT_ID,
                    text=msg,
                    parse_mode="Markdown"
                )
                print(f"Transaction sent after {wait_seconds // 60} min wait.")
            except Exception as e:
                print(f"Error sending transaction: {e}")

# --- ৩. নতুন মেম্বার জয়েন করলে ওয়েলকাম মেসেজ দেওয়া ---
async def welcome_new_member(update: Update, context: ContextTypes.DEFAULT_TYPE):
    result = update.chat_member
    new_member = result.new_chat_member

    # যেকোনো গ্রুপে নতুন মেম্বার ঢুকলে
    if new_member.status in ["member", "administrator"] and result.old_chat_member.status in ["left", "kicked"]:
        user = new_member.user
        user_name = user.full_name
        
        welcome_text = (
            f"👋 স্বাগতম, [{user_name}](tg://user?id={user.id}) ہمارے গ্রুপে!\n\n"
            f"📌 গ্রুপে নিয়মিত আপডেট ও লেনদেন সংক্রান্ত তথ্য শেয়ার করা হয়।\n"
            f"দয়া করে নিয়মাবলি মেনে চলুন।"
        )
        
        await context.bot.send_message(
            chat_id=update.effective_chat.id,
            text=welcome_text,
            parse_mode="Markdown"
        )
        print(f"Welcomed: {user_name} in chat: {update.effective_chat.id}")

# --- ৪. ব্যাকগ্রাউন্ড টাস্ক চালু করা ---
async def post_init(application):
    asyncio.create_task(send_periodic_transactions(application))
    asyncio.create_task(start_dummy_web_server())

def main():
    print("বট চালু হচ্ছে...")
    app = ApplicationBuilder().token(BOT_TOKEN).post_init(post_init).build()
    app.add_handler(ChatMemberHandler(welcome_new_member, ChatMemberHandler.CHAT_MEMBER))
    app.run_polling(allowed_updates=Update.ALL_TYPES)

if __name__ == "__main__":
    main()
