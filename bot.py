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
    ConversationHandler,
    filters,
    ContextTypes
)

# --- কনফিগারেশন ---
BOT_TOKEN = "8826168593:AAE6mhFPKuJGz5TUwua4I4P7h3n0D_2_SeI"
ADMIN_USER_ID = 8919985167  # আপনার নির্দিষ্ট টেলিগ্রাম আইডি

# গ্রুপ আইডিসমূহ
WELCOME_CHAT_ID = "-1004471047712"     # ১ম গ্রুপ: ওয়েলকাম
TRANSACTION_CHAT_ID = "-1003991468184"  # ২য় গ্রুপ: ট্রানজ্যাকশন

# বাটন ইউআরএল
GMAIL_BOT_URL = "https://t.me/NEW_FRESH_GMAILACCOUNTSELL50_bot"

# ৩টি মেসেজ শেষ হওয়ার পর বিরতি (৫ মিনিট = ৩০০ সেকেন্ড)
BREAK_AFTER_3_MESSAGES = 300

is_tx_active = True

# কাস্টম ট্রানজ্যাকশনের কনভারসেশন স্টেটস
WAITING_FOR_NAME, WAITING_FOR_AMOUNT, WAITING_FOR_METHOD = range(3)

# ডিফল্ট কাস্টমার নামসমূহ
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
    print(f"Keep-Alive Server running on port {port}")
    async with server:
        await server.serve_forever()

# --- ট্রানজ্যাকশন বাটন ---
def get_transaction_keyboard():
    keyboard = [
        [InlineKeyboardButton("FAST GMAIL SELL", url=GMAIL_BOT_URL)]
    ]
    return InlineKeyboardMarkup(keyboard)

# --- ৩ ধাপের ট্রানজ্যাকশন এক্সিকিউশন ---
async def execute_3_step_transaction(bot, chat_id, customer_name, amount, method):
    tx_id = f"TX{random.randint(10000000, 99999999)}"
    reply_markup = get_transaction_keyboard()

    # ধাপ ১: পেমেন্ট উইথড্র রিকোয়েস্ট
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

    # ধাপ ২: সিকিউরিটি ভেরিফিকেশন চেক
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

    # ৩০ সেকেন্ড অপেক্ষা
    await asyncio.sleep(30)

    # ধাপ ৩: পেমেন্ট ডিপার্টমেন্ট সম্পন্ন
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
        print(f"Transaction completed for {customer_name}")
    except Exception as e:
        print(f"Error Step 3: {e}")

# --- স্বয়ংক্রিয় ট্রানজ্যাকশন লুপ ---
async def send_periodic_transactions(application):
    await asyncio.sleep(5)
    while True:
        if is_tx_active:
            try:
                name = random.choice(NAMES)
                amt = f"{random.uniform(50.00, 2000.00):.2f}"
                met = random.choice(PAYMENT_METHODS)
                await execute_3_step_transaction(application.bot, TRANSACTION_CHAT_ID, name, amt, met)
            except Exception as e:
                print(f"Auto tx loop error: {e}")
        
        # ৩য় মেসেজ পাঠানোর পর ঠিক ৫ মিনিট অপেক্ষা
        await asyncio.sleep(BREAK_AFTER_3_MESSAGES)

# --- ওয়েলকাম মেসেজ (৩০ সেকেন্ডে ডিলিট) ---
async def send_and_auto_delete_welcome(bot, chat_id, user):
    user_id = user.id
    if user_id in recently_welcomed_users:
        return
    recently_welcomed_users.add(user_id)

    user_link = f'<a href="tg://user?id={user.id}">{user.full_name or "মেম্বার"}</a>'
    welcome_text = (
        f"🌸 <b>আসসালামু আলাইকুম</b>, {user_link}!\n\n"
        f"আমাদের কমিউনিটিতে আপনাকে স্বাগতম। 🎉\n"
        f"📌 গ্রুপের নিয়ম-কানুন মেনে চলুন এবং নিয়মিত আপডেট উপভোগ করুন।"
    )
    try:
        sent_msg = await bot.send_message(chat_id=chat_id, text=welcome_text, parse_mode="HTML")
        await asyncio.sleep(30)
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

# --- মাস্টার অ্যাডমিন ড্যাশবোর্ড ---
def get_dashboard_markup():
    status_label = "🔴 অটো ট্রানজ্যাকশন বন্ধ করুন" if is_tx_active else "🟢 অটো ট্রানজ্যাকশন চালু করুন"
    keyboard = [
        [InlineKeyboardButton(status_label, callback_data="toggle_tx")],
        [InlineKeyboardButton("✍️ কাস্টম ট্রানজ্যাকশন তৈরি করুন", callback_data="start_custom_tx")],
        [InlineKeyboardButton("⚡ র‍্যান্ডম টেস্ট ট্রানজ্যাকশন পাঠান", callback_data="instant_random_tx")]
    ]
    return InlineKeyboardMarkup(keyboard)

def get_dashboard_text():
    status_str = "চালু আছে 🟢" if is_tx_active else "বন্ধ আছে 🔴"
    return (
        f"👑 <b>মাস্টার অ্যাডমিন কন্ট্রোল প্যানেল</b>\n"
        f"━━━━━━━━━━━━━━━━━━━━━━\n"
        f"👤 <b>অ্যাডমিন আইডি:</b> <code>{ADMIN_USER_ID}</code>\n"
        f"⚙️ <b>অটো ট্রানজ্যাকশন:</b> {status_str}\n"
        f"⏱️ <b>টাইমার:</b> প্রতি ৫ মিনিট পর পর\n\n"
        f"নিচের বাটন চেপে যা ইচ্ছা নিয়ন্ত্রণ করুন:"
    )

# /admin কমান্ড
async def admin_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if update.effective_user.id != ADMIN_USER_ID:
        await update.message.reply_text("⛔ আপনি এই বটের অ্যাডমিন নন!")
        return

    await update.message.reply_text(
        get_dashboard_text(),
        reply_markup=get_dashboard_markup(),
        parse_mode="HTML"
    )

# ড্যাশবোর্ড বাটন ক্লিক হ্যান্ডলার
async def dashboard_callback(update: Update, context: ContextTypes.DEFAULT_TYPE):
    global is_tx_active
    query = update.callback_query
    await query.answer()

    if query.from_user.id != ADMIN_USER_ID:
        await query.message.reply_text("⛔ অনুমতি নেই!")
        return

    data = query.data

    if data == "toggle_tx":
        is_tx_active = not is_tx_active
        await query.edit_message_text(
            get_dashboard_text(),
            reply_markup=get_dashboard_markup(),
            parse_mode="HTML"
        )
    elif data == "instant_random_tx":
        await query.message.reply_text("⚡ ২য় গ্রুপে ৩ ধাপের র‍্যান্ডম ট্রানজ্যাকশন পাঠানো শুরু হয়েছে...")
        name = random.choice(NAMES)
        amt = f"{random.uniform(50.00, 2000.00):.2f}"
        met = random.choice(PAYMENT_METHODS)
        asyncio.create_task(execute_3_step_transaction(context.bot, TRANSACTION_CHAT_ID, name, amt, met))

# --- কাস্টম ট্রানজ্যাকশন বাটন কনভারসেশন লজিক ---
async def start_custom_tx(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()

    if query.from_user.id != ADMIN_USER_ID:
        return ConversationHandler.END

    await query.message.reply_text(
        "📝 <b>ধাপ ১: কাস্টমারের নাম লিখুন</b>\n\n"
        "যার উইথড্র পাঠাবেন তার নাম লিখে এখানে সেন্ড করুন (যেমন: <code>Md Karim</code>):",
        parse_mode="HTML"
    )
    return WAITING_FOR_NAME

async def receive_custom_name(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data["custom_name"] = update.message.text.strip()
    await update.message.reply_text(
        "💰 <b>ধাপ ২: টাকার পরিমাণ লিখুন</b>\n\n"
        "কত টাকা পাঠাতে চান লিখে সেন্ড করুন (যেমন: <code>500</code> বা <code>1250</code>):",
        parse_mode="HTML"
    )
    return WAITING_FOR_AMOUNT

async def receive_custom_amount(update: Update, context: ContextTypes.DEFAULT_TYPE):
    amt_text = update.message.text.strip().replace("৳", "")
    try:
        amt_val = float(amt_text)
        context.user_data["custom_amount"] = f"{amt_val:.2f}"
    except ValueError:
        context.user_data["custom_amount"] = amt_text

    # মেথড বাটন
    keyboard = [
        [InlineKeyboardButton("Bkash", callback_data="method_Bkash"),
         InlineKeyboardButton("Nagad", callback_data="method_Nagad")],
        [InlineKeyboardButton("Binance", callback_data="method_Binance")]
    ]
    reply_markup = InlineKeyboardMarkup(keyboard)

    await update.message.reply_text(
        "🏦 <b>ধাপ ৩: পেমেন্ট মাধ্যম বেছে নিন</b>\n\n"
        "নিচের বাটন থেকে মাধ্যম সিলেক্ট করুন:",
        reply_markup=reply_markup,
        parse_mode="HTML"
    )
    return WAITING_FOR_METHOD

async def receive_custom_method(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()

    selected_method = query.data.replace("method_", "")
    name = context.user_data.get("custom_name", "Customer")
    amount = context.user_data.get("custom_amount", "500.00")

    await query.edit_message_text(
        f"✅ <b>কাস্টম ট্রানজ্যাকশন সম্পন্ন ও পাঠানো হচ্ছে!</b>\n\n"
        f"👤 কাস্টমার: <b>{name}</b>\n"
        f"💰 পরিমাণ: <b>৳{amount}</b>\n"
        f"🏦 মাধ্যম: <b>{selected_method}</b>\n\n"
        f"২য় গ্রুপে ৩টি ধাপ (৩০ সেকেন্ড অন্তর) লাইভ শুরু হয়ে গেছে।",
        parse_mode="HTML"
    )

    # ২য় গ্রুপে পাঠানো
    asyncio.create_task(
        execute_3_step_transaction(context.bot, TRANSACTION_CHAT_ID, name, amount, selected_method)
    )

    return ConversationHandler.END

async def cancel_custom_tx(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text("❌ বাতিল করা হয়েছে।")
    return ConversationHandler.END

# --- সার্ভিস শুরু ---
async def post_init(application):
    asyncio.create_task(send_periodic_transactions(application))
    asyncio.create_task(start_dummy_web_server())

def main():
    print("বট চালু হচ্ছে...")
    app = ApplicationBuilder().token(BOT_TOKEN).post_init(post_init).build()

    # কাস্টম ট্রানজ্যাকশন ইন্টারঅ্যাক্টিভ হ্যান্ডলার
    custom_tx_conv = ConversationHandler(
        entry_points=[CallbackQueryHandler(start_custom_tx, pattern="^start_custom_tx$")],
        states={
            WAITING_FOR_NAME: [MessageHandler(filters.TEXT & ~filters.COMMAND, receive_custom_name)],
            WAITING_FOR_AMOUNT: [MessageHandler(filters.TEXT & ~filters.COMMAND, receive_custom_amount)],
            WAITING_FOR_METHOD: [CallbackQueryHandler(receive_custom_method, pattern="^method_")]
        },
        fallbacks=[CommandHandler("cancel", cancel_custom_tx)]
    )

    # অ্যাডমিন কমান্ড ও ড্যাশবোর্ড বাটন
    app.add_handler(CommandHandler("admin", admin_command))
    app.add_handler(custom_tx_conv)
    app.add_handler(CallbackQueryHandler(dashboard_callback))

    # জয়েনিং হ্যান্ডলারসমূহ
    app.add_handler(MessageHandler(filters.StatusUpdate.NEW_CHAT_MEMBERS, handle_new_chat_members))
    app.add_handler(ChatMemberHandler(handle_chat_member_updated, ChatMemberHandler.CHAT_MEMBER))

    app.run_polling(allowed_updates=Update.ALL_TYPES)

if __name__ == "__main__":
    main()
