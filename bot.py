import telebot
from telebot.types import InlineKeyboardMarkup, InlineKeyboardButton
from telebot.apihelper import ApiTelegramException
import pandas as pd
import os
import warnings

warnings.filterwarnings('ignore', category=UserWarning)

TELEGRAM_BOT_TOKEN = "8785580026:AAF0c5sJkYxIJOqs9dtPbWZA-JqGACI39pc"
bot = telebot.TeleBot(TELEGRAM_BOT_TOKEN)

DATABASE_DIR = "Users_Database"
os.makedirs(DATABASE_DIR, exist_ok=True)
user_states = {}

@bot.message_handler(commands=['start'])
def send_welcome(message):
    chat_id = message.chat.id
    if chat_id not in user_states:
        user_states[chat_id] = {}
    user_states[chat_id]['action'] = 'idle'
    
    # رسالة ترحيبية بسيطة، حلوة وتفتح النفس للمستخدم
    welcome_text = (
        "أهلاً وسهلاً بك عزيزي المشترك 🌸\n"
        "منور البوت. اني مساعدك الشخصي للبحث في قواعد بيانات الإكسل بكل سهولة وسرعة.\n\n"
        "كل اللي عليك تسويه، تدزلي ملف الإكسل (Excel) حتى نحفظه بأرشيفك الخاص، ووراها تبحث عن أي اسم أو معلومة تريدها بضغطة زر."
    )
    
    markup = InlineKeyboardMarkup()
    markup.add(InlineKeyboardButton("📂 ملفاتي المحفوظة", callback_data="list_my_files"))
    
    bot.reply_to(message, welcome_text, reply_markup=markup)

@bot.message_handler(content_types=['document'])
def handle_docs(message):
    chat_id = message.chat.id
    user_dir = os.path.join(DATABASE_DIR, str(chat_id))
    os.makedirs(user_dir, exist_ok=True)
    
    file_name = message.document.file_name
    file_path = os.path.join(user_dir, file_name)
    
    try:
        bot.send_chat_action(chat_id, 'upload_document')
        file_info = bot.get_file(message.document.file_id)
        downloaded_file = bot.download_file(file_info.file_path)
        
        with open(file_path, 'wb') as f:
            f.write(downloaded_file)
            
    except Exception as e:
        bot.reply_to(message, "عذراً صار خطأ بتحميل الملف، فدوة ارجع دزه مرة ثانية.")
        return 
        
    if chat_id not in user_states:
        user_states[chat_id] = {}
    user_states[chat_id]['active_file'] = file_path

    markup = InlineKeyboardMarkup(row_width=1)
    markup.add(InlineKeyboardButton("🔍 ابدأ البحث بالملف", callback_data="start_search"))

    bot.reply_to(message, f"تم حفظ الملف يمك بنجاح:\n({file_name})\n\nهسة تكدر تبحث براحتك، اضغط الزر جوه واكتب الشي اليريدك:", reply_markup=markup)

@bot.callback_query_handler(func=lambda call: True)
def handle_query(call):
    chat_id = call.message.chat.id
    
    try:
        if call.data == "list_my_files":
            user_dir = os.path.join(DATABASE_DIR, str(chat_id))
            if not os.path.exists(user_dir) or not os.listdir(user_dir):
                bot.answer_callback_query(call.id, "ما عندك أي ملفات محفوظة حالياً!", show_alert=True)
                return
            
            markup = InlineKeyboardMarkup(row_width=1)
            files = sorted(os.listdir(user_dir))
            for index, f_name in enumerate(files):
                markup.add(InlineKeyboardButton(f"📄 {f_name[:25]}...", callback_data=f"sel_{index}"))
                
            bot.edit_message_text("هذي ملفاتك المحفوظة، اضغط على الملف الي تريده:", chat_id, call.message.message_id, reply_markup=markup)

        elif call.data.startswith("sel_"):
            file_index = int(call.data.split("_")[1])
            user_dir = os.path.join(DATABASE_DIR, str(chat_id))
            files = sorted(os.listdir(user_dir))
            
            if file_index < len(files):
                file_name = files[file_index]
                file_path = os.path.join(user_dir, file_name)
                
                if chat_id not in user_states: user_states[chat_id] = {}
                user_states[chat_id]['active_file'] = file_path
                
                markup = InlineKeyboardMarkup(row_width=1)
                markup.add(InlineKeyboardButton("🔍 ابدأ البحث بالملف", callback_data="start_search"))
                bot.edit_message_text(f"تم اختيار الملف: {file_name}\nهسة اكتب الشي الي تريد تبحث عنه:", chat_id, call.message.message_id, reply_markup=markup)
            else:
                bot.answer_callback_query(call.id, "الملف غير موجود.", show_alert=True)

        elif call.data == "start_search":
            if chat_id not in user_states or 'active_file' not in user_states[chat_id]:
                bot.answer_callback_query(call.id, "يرجى اختيار أو رفع ملف أولاً!", show_alert=True)
                return
            
            user_states[chat_id]['action'] = 'searching'
            bot.send_message(chat_id, "تفضل، اكتب الاسم الكامل، الرقم الإحصائي، أو أي كلمة تريد تبحث عنها:")
            
    except ApiTelegramException as e:
        if "message is not modified" in str(e):
            pass 
        else:
            print(f"Telegram API Error: {e}")

@bot.message_handler(func=lambda message: True)
def handle_text(message):
    chat_id = message.chat.id
    active_file = user_states.get(chat_id, {}).get('active_file', None)

    if not active_file or not os.path.exists(active_file):
        bot.reply_to(message, "حبيبي دزلي ملف إكسل أول شي حتى أگدر أبحثلك بداخله!")
        return

    user_text = message.text.strip()
    bot.send_chat_action(chat_id, 'typing')

    try:
        all_sheets = pd.read_excel(active_file, sheet_name=None, header=None)
        found_results = False
        search_words = user_text.split()
        
        for sheet_name, df in all_sheets.items():
            if df.empty or df.notna().sum().sum() == 0:
                continue
            
            row_texts = df.fillna('').astype(str).apply(lambda row: ' '.join(row), axis=1)
            match_mask = row_texts.apply(lambda text: all(w in text for w in search_words))
            results = df[match_mask]
            
            if not results.empty:
                found_results = True
                headers = df.iloc[0].fillna('').astype(str).tolist() if len(df) > 0 else []
                
                for index, row in results.iterrows():
                    excel_row_num = index + 1
                    
                    basic_info = []
                    work_info = []
                    notes_list = []
                    
                    for col_idx, val in enumerate(row):
                        val_str = str(val).strip()
                        if pd.notna(val) and val_str not in ["", "nan"]:
                            col_name = f"حقل {col_idx+1}"
                            if col_idx < len(headers) and headers[col_idx].strip() not in ["", "nan", "Unnamed"]:
                                col_name = headers[col_idx].strip()
                            
                            if any(k in col_name for k in ["الاسم", "الرتبة", "الرقم", "الجنس", "الوجبة", "العنوان", "ت"]):
                                basic_info.append(f"▪️ {col_name}: {val_str}")
                            elif any(k in col_name for k in ["قاطع", "وزارة", "منطقة", "هدف", "موقع"]):
                                work_info.append(f"🏢 {col_name}: {val_str}")
                            else:
                                notes_list.append(f"- {val_str}")
                    
                    card = (
                        f"بطاقة تعريفية رسمية\n"
                        f"القسم: {sheet_name} | السطر: {excel_row_num}\n"
                        f"------------------------------------------\n"
                        f"معطيات المنتسب الأساسي:\n" + "\n".join(basic_info) + "\n\n"
                    )
                    if work_info:
                        card += "بيانات العمل والتنسيق:\n" + "\n".join(work_info) + "\n\n"
                    if notes_list:
                        card += "سجل التنقلات والملاحظات:\n" + "\n".join(notes_list) + "\n"
                    
                    bot.send_message(chat_id, card)
        
        if not found_results:
            bot.reply_to(message, f"عذراً، ما لقيت أي مطابقة تخص: '{user_text}' بهذا الملف.")

    except Exception as e:
        bot.reply_to(message, "صار خطأ بمعالجة الملف، تأكد منه وارجع دزه.")
        print(f"Error in handle_text: {e}")

if __name__ == '__main__':
    print("البوت شغال بكل سلاسة... اضغط Ctrl+C للإيقاف.")
    bot.infinity_polling()
