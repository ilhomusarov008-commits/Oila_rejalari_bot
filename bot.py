import os
from aiohttp import web
import sqlite3
import ast
from datetime import datetime

from aiogram import Bot, Dispatcher, F
from aiogram.filters import Command, CommandStart
from aiogram.types import Message, CallbackQuery, FSInputFile
from aiogram.utils.keyboard import InlineKeyboardBuilder
from aiogram.fsm.state import State, StatesGroup
from aiogram.fsm.context import FSMContext

from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

# Bot sozlamalari
TOKEN = os.getenv("BOT_TOKEN", "8627810193:AAGPB34GdXHjObWRRpY77qMyDsQrFf6KuE0")
ADMIN_ID = 8671467518
DB_PATH = os.getenv("DB_PATH", "survey_new.db")
EXPORT_PATH = os.getenv("EXPORT_PATH", "Oila_va_nikoh_mezonlari_tadqiqoti.xlsx")

# 1. Demografik savollar
QUESTIONS = [
    ("age_group", "Yoshingiz qaysi oraliqda?", ["17–19", "20–22", "23–25", "26 va undan katta"]),
    ("course", "Kursingiz?", ["1-kurs", "2-kurs", "3-kurs", "4-kurs", "Magistratura / Boshqa"]),
    ("residence", "Doimiy yashash joyingiz?", ["Shahar", "Tuman/Qishloq"]),
    ("marital_status", "Hozirgi oilaviy holatingiz?", ["Turmush qurmagan", "Turmush qurgan"]),
    ("ideal_marriage_age", "Siz uchun nikoh qurishning maqbul yoshi nechada?", None),
    ("age_difference_pref", "Bo‘lajak juftingiz bilan yosh farqini qanday tasavvur qilasiz?",
     ["Mendan kichik", "Tengdosh", "Mendan katta", "Farqi muhim emas"]),
    ("desired_children", "Kelajakda nechta farzandli bo‘lishni xohlaysiz?",
     ["0", "1", "2", "3", "4", "5 va undan ko‘p"]),
]

# 2. Likert mezonlari
LIKERT = [
    ("kind_understanding", "Mehribon va tushunuvchan bo‘lishi"),
    ("dependable_character", "Ishonchli xarakterga ega bo‘lishi"),
    ("intelligent", "Aqlli bo‘lishi"),
    ("emotional_stability", "Hissiy jihatdan barqaror va yetuk bo‘lishi"),
    ("pleasing_disposition", "Yoqimli fe’l-atvorga ega bo‘lishi"),
    ("sociability", "Ijtimoiy va muloqotga kirishuvchan bo‘lishi"),
    ("honesty_responsibility", "Halol va mas’uliyatli bo‘lishi"),
    ("education", "Yaxshi ta’limga ega bo‘lishi"),
    ("self_development", "O‘z ustida ishlashi va rivojlanishga intilishi"),
    ("ambition", "Shuhratparast va maqsadga intiluvchan bo‘lishi"),
    ("industriousness", "Mehnatsevar bo‘lishi"),
    ("financial_prospect", "Yaxshi moliyaviy istiqbolga ega bo‘lishi"),
    ("social_status", "Ijtimoiy mavqei yoki obro‘si yuqori bo‘lishi"),
    ("physical_attractiveness", "Jismoniy jihatdan jozibador bo‘lishi"),
    ("age_compatibility", "Yosh jihatidan sizga mos bo‘lishi"),
    ("health_lifestyle", "Sog‘lom turmush tarziga ega bo‘lishi"),
    ("neatness", "Ozoda va saranjom bo‘lishi"),
    ("marriage_orientation", "Oila qurishga jiddiy munosabatda bo‘lishi"),
    ("family_responsibility", "Oilaviy mas’uliyatni o‘z zimmasiga olishga tayyor bo‘lishi"),
]

# 3. SD Shkalasi
SD_QUESTIONS = [
    ("sd_1", "Hech qachon birovga nisbatan ich-ichimdan g‘azab yoki nafrat sezmaganman.", True),
    ("sd_2", "Gohida bajara olmaydigan va'dalarni berib qo‘yaman.", False),
    ("sd_3", "Har doim o‘z xatolarimni ochiq tan olaman.", True),
    ("sd_4", "Ba'zan atrofimdagilar haqida g‘iybat qilishim yoki g‘iybatni eshitishim mumkin.", False),
    ("sd_5", "Har qanday vaziyatda ham samimiy va xushfe'l bo‘lishga intilaman.", True),
    ("sd_6", "Ba'zan kayfiyatim yomon bo‘lsa, atrofdagilarga qo‘pollik qilib qo‘yaman.", False),
]

MAIN_FACTORS = ["Shaxsiy xarakter", "Intellekt", "Tashqi ko‘rinish", "Oilaviy qadriyatlar", "Ijtimoiy-iqtisodiy imkoniyatlar", "Boshqa"]

def db():
    conn = sqlite3.connect(DB_PATH)
    conn.execute("""CREATE TABLE IF NOT EXISTS responses (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        telegram_id INTEGER,
        name TEXT,
        created_at TEXT,
        data TEXT
    )""")
    conn.commit()
    return conn

def save_response(telegram_id, name, data):
    conn = db()
    conn.execute(
        "INSERT INTO responses (telegram_id,name,created_at,data) VALUES (?,?,?,?)",
        (telegram_id, name, datetime.now().isoformat(timespec="seconds"), repr(data))
    )
    conn.commit()
    conn.close()

def has_completed_response(telegram_id):
    conn = db()
    row = conn.execute("SELECT 1 FROM responses WHERE telegram_id = ? LIMIT 1", (telegram_id,)).fetchone()
    conn.close()
    return row is not None

def calculate_sd_level(sd_score):
    if sd_score <= 1:
        return "past"
    elif sd_score <= 4:
        return "o‘rtacha"
    else:
        return "yuqori"

# EXCEL AVTOMATIK TAHLIL
def export_xlsx():
    conn = db()
    rows = conn.execute("SELECT id,telegram_id,name,created_at,data FROM responses ORDER BY id").fetchall()
    conn.close()

    wb = Workbook()
    
    header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    sec_header_fill = PatternFill(start_color="2D6A4F", end_color="2D6A4F", fill_type="solid")
    font_bold_white = Font(bold=True, color="FFFFFF")
    font_bold = Font(bold=True)

    # 1. RAW DATA SHEET
    ws_raw = wb.active
    ws_raw.title = "Javoblar (Raw)"
    
    headers = ["ID", "Respondent kodi", "Ism", "Sana", "Yosh", "Kurs", "Yashash joyi", "Oilaviy holati", "Maqbul nikoh yoshi", "Yosh farqi preference", "Istalgan farzandlar"]
    headers += [label for _, label in LIKERT]
    headers += ["TOP-5 Reyting", "SD Score (0-6)", "SD Darajasi", "Asosiy yo‘nalish"]
    ws_raw.append(headers)

    # 2. CODED DATA SHEET
    ws_coded = wb.create_sheet("Kodlangan ma'lumotlar")
    coded_headers = ["ID", "Age_Group", "Course", "Residence", "Marital_Status", "Ideal_Marriage_Age", "Age_Diff_Pref", "Desired_Children"]
    coded_headers += [f"Q_{i+1}" for i in range(len(LIKERT))]
    coded_headers += ["SD_Total", "SD_Level", "Main_Factor"]
    ws_coded.append(coded_headers)

    total_rows = len(rows)

    for rid, telegram_id, name, created_at, data_text in rows:
        try:
            data = ast.literal_eval(data_text)
        except Exception:
            data = {}

        resp_code = f"RESP-{rid:04d}"
        
        raw_row = [
            rid, resp_code, name, created_at, 
            data.get("age_group", ""), data.get("course", ""), data.get("residence", ""), data.get("marital_status", ""),
            data.get("ideal_marriage_age", ""), data.get("age_difference_pref", ""), data.get("desired_children", "")
        ]
        for key, _ in LIKERT:
            raw_row.append(data.get(key, ""))
        
        sd_score = data.get("sd_score", 0)
        sd_level = calculate_sd_level(sd_score)
        
        raw_row += [data.get("top5_ranking", ""), sd_score, sd_level, data.get("main_factor", "")]
        ws_raw.append(raw_row)

        coded_row = [
            rid, data.get("age_group", ""), data.get("course", ""), data.get("residence", ""), data.get("marital_status", ""),
            data.get("ideal_marriage_age", ""), data.get("age_difference_pref", ""), data.get("desired_children", "")
        ]
        for key, _ in LIKERT:
            coded_row.append(data.get(key, None))
        coded_row += [sd_score, sd_level, data.get("main_factor", "")]
        ws_coded.append(coded_row)

    # 3. STATISTIKA VARAQI
    ws_stats = wb.create_sheet("📊 Umumiy Statistika")
    
    last_r = max(2, total_rows + 1)

    ws_stats.cell(row=1, column=1, value="TADQIQOTNING UMUMIY STATISTIK XULOSALARI").font = Font(bold=True, size=14, color="1F4E78")
    ws_stats.cell(row=2, column=1, value="Jami respondentlar soni:").font = font_bold
    ws_stats.cell(row=2, column=2, value=f"=COUNTA('Javoblar (Raw)'!A2:A{last_r})")

    ws_stats.cell(row=4, column=1, value="1. DEMOGRAFIK TAQSIMOT (FOIZLARDA)").font = font_bold_white
    ws_stats.cell(row=4, column=1).fill = sec_header_fill

    ws_stats.append(["Kategoriya", "Variant", "Soni", "Ushlagan ulushi (%)"])
    for cell in ws_stats[5]: cell.font = font_bold

    demo_items = [
        ("Yosh guruhi", "17–19", "E"), ("Yosh guruhi", "20–22", "E"), ("Yosh guruhi", "23–25", "E"), ("Yosh guruhi", "26 va undan katta", "E"),
        ("Yashash joyi", "Shahar", "G"), ("Yashash joyi", "Tuman/Qishloq", "G"),
        ("Oilaviy holati", "Turmush qurmagan", "H"), ("Oilaviy holati", "Turmush qurgan", "H")
    ]

    r_idx = 6
    for cat, val, col in demo_items:
        ws_stats.cell(row=r_idx, column=1, value=cat)
        ws_stats.cell(row=r_idx, column=2, value=val)
        ws_stats.cell(row=r_idx, column=3, value=f"=COUNTIF('Javoblar (Raw)'!{col}2:{col}{last_r}, \"{val}\")")
        ws_stats.cell(row=r_idx, column=4, value=f"=IF($B$2>0, C{r_idx}/$B$2, 0)")
        ws_stats.cell(row=r_idx, column=4).number_format = '0.0%'
        r_idx += 1

    r_idx += 1
    ws_stats.cell(row=r_idx, column=1, value="2. JUFT TANLASH MEZONLARINING TAHLILI VA ENG MUHIM KO‘RSATKICHLAR").font = font_bold_white
    ws_stats.cell(row=r_idx, column=1).fill = sec_header_fill
    r_idx += 1

    headers_mezon = [
        "Nomer", "Mezon nomi", "O‘rtacha Ball (Mean)", "Baho", 
        "'Juda muhim (3)' belgilaganlar (Soni)", "'Juda muhim' ulushi (%)", 
        "TOP-5 ga kiritganlar (Soni)", "TOP-5 ulushi (%)"
    ]
    
    for c_idx, h_text in enumerate(headers_mezon, 1):
        cell = ws_stats.cell(row=r_idx, column=c_idx, value=h_text)
        cell.font = font_bold

    r_idx += 1

    start_likert_col = 12
    top5_col_let = get_column_letter(start_likert_col + len(LIKERT))

    for i, (_, label) in enumerate(LIKERT):
        col_let = get_column_letter(start_likert_col + i)
        
        ws_stats.cell(row=r_idx, column=1, value=i+1)
        ws_stats.cell(row=r_idx, column=2, value=label)
        
        ws_stats.cell(row=r_idx, column=3, value=f"=IF($B$2>0, AVERAGE('Javoblar (Raw)'!{col_let}2:{col_let}{last_r}), 0)")
        ws_stats.cell(row=r_idx, column=3).number_format = '0.00'
        
        ws_stats.cell(row=r_idx, column=4, value=f"=IF(C{r_idx}>2.2, \"Juda Muhim\", IF(C{r_idx}>1.5, \"Muhim\", \"Kamroq Muhim\"))")
        
        ws_stats.cell(row=r_idx, column=5, value=f"=COUNTIF('Javoblar (Raw)'!{col_let}2:{col_let}{last_r}, 3)")
        
        ws_stats.cell(row=r_idx, column=6, value=f"=IF($B$2>0, E{r_idx}/$B$2, 0)")
        ws_stats.cell(row=r_idx, column=6).number_format = '0.0%'
        
        ws_stats.cell(row=r_idx, column=7, value=f"=COUNTIF('Javoblar (Raw)'!{top5_col_let}2:{top5_col_let}{last_r}, \"*\"&B{r_idx}&\"*\")")
        
        ws_stats.cell(row=r_idx, column=8, value=f"=IF($B$2>0, G{r_idx}/$B$2, 0)")
        ws_stats.cell(row=r_idx, column=8).number_format = '0.0%'
        
        r_idx += 1

    r_idx += 1
    ws_stats.cell(row=r_idx, column=1, value="3. RESPONDENTLAR SAMIMIYLIGI (SD DARAJASI)").font = font_bold_white
    ws_stats.cell(row=r_idx, column=1).fill = sec_header_fill
    r_idx += 1

    ws_stats.cell(row=r_idx, column=1, value="SD Darajasi").font = font_bold
    ws_stats.cell(row=r_idx, column=2, value="Soni").font = font_bold
    ws_stats.cell(row=r_idx, column=3, value="Ulushi (%)").font = font_bold
    ws_stats.cell(row=r_idx, column=4, value="Metodologik izoh").font = font_bold
    r_idx += 1

    sd_types = [
        ("Past (Samimiy)", "past", "Javoblar juda ishonchli va samimiy"),
        ("O‘rtacha (Meyoriy)", "o‘rtacha", "Javoblar statistik qabul qilinadi"),
        ("Yuqori (Maqbullik yuqori)", "yuqori", "Javoblar o‘zini yaxshi ko‘rsatishga moyil")
    ]

    sd_col_let = get_column_letter(start_likert_col + len(LIKERT) + 2)
    for title, val, comment in sd_types:
        ws_stats.cell(row=r_idx, column=1, value=title)
        ws_stats.cell(row=r_idx, column=2, value=f"=COUNTIF('Javoblar (Raw)'!{sd_col_let}2:{sd_col_let}{last_r}, \"{val}\")")
        ws_stats.cell(row=r_idx, column=3, value=f"=IF($B$2>0, B{r_idx}/$B$2, 0)")
        ws_stats.cell(row=r_idx, column=3).number_format = '0.0%'
        ws_stats.cell(row=r_idx, column=4, value=comment)
        r_idx += 1

    for ws in [ws_raw, ws_coded, ws_stats]:
        ws.freeze_panes = "A2"
        for col_idx, column_cells in enumerate(ws.columns, 1):
            max_len = max(len(str(cell.value or "")) for cell in list(column_cells)[:50])
            ws.column_dimensions[get_column_letter(col_idx)].width = min(max(max_len + 3, 12), 42)

    for cell in ws_raw[1]:
        cell.font = font_bold_white
        cell.fill = header_fill
        cell.alignment = Alignment(wrap_text=True, vertical="center", horizontal="center")

    wb.save(EXPORT_PATH)
    return EXPORT_PATH, len(rows)

def buttons(items, prefix="ans"):
    kb = InlineKeyboardBuilder()
    for i, item in enumerate(items):
        kb.button(text=item, callback_data=f"{prefix}:{i}")
    kb.adjust(2)
    return kb.as_markup()

def likert_buttons():
    kb = InlineKeyboardBuilder()
    labels = [
        ("0 — Umuman muhim emas", 0),
        ("1 — Unchalik muhim emas", 1),
        ("2 — Muhim", 2),
        ("3 — Juda muhim / deyarli zarur", 3),
    ]
    for label, val in labels:
        kb.button(text=label, callback_data=f"likert:{val}")
    kb.adjust(1)
    return kb.as_markup()

def sd_buttons():
    kb = InlineKeyboardBuilder()
    kb.button(text="Ha", callback_data="sd:1")
    kb.button(text="Yo‘q", callback_data="sd:0")
    kb.adjust(2)
    return kb.as_markup()

def admin_keyboard():
    kb = InlineKeyboardBuilder()
    kb.button(text="📊 Statistika", callback_data="admin:stats")
    kb.button(text="📥 Excel yuklab olish (Tahlil bilan)", callback_data="admin:export")
    kb.button(text="🧹 Mening testlarimni tozalash", callback_data="admin:reset")
    kb.button(text="💥 BARCHA respondentlarni o‘chirish", callback_data="admin:reset_all")
    kb.button(text="🔄 So‘rovnomani boshlash", callback_data="admin:start_survey")
    kb.adjust(1)
    return kb.as_markup()

class Survey(StatesGroup):
    name = State()
    question = State()
    likert = State()
    top5 = State()
    sd_scale = State()
    main_factor = State()
    custom_factor = State()

async def start_survey(message: Message, state: FSMContext):
    await state.clear()
    await state.set_state(Survey.name)
    await message.answer(
        "Assalom u alaykum! O`sarov Ilhom tomonidan yaratilgan so`rovnoma botiga xush kelibsiz.\n\n"
        "Iltimos, ismingizni kiriting (Siz haqingizdagi boshqa barcha malumotlar anonimlashtiriladi):"
    )

async def send_next_q(message: Message, state: FSMContext):
    data = await state.get_data()
    idx = data.get("q_index", 0)
    if idx < len(QUESTIONS):
        key, q, opts = QUESTIONS[idx]
        await state.set_state(Survey.question)
        if opts is None:
            await message.answer(
                f"📌 {idx+1}/{len(QUESTIONS)}\n\n{q}\n\n"
                "Javobingizni faqat raqam bilan yozing. Masalan: 25"
            )
        else:
            await message.answer(f"📌 {idx+1}/{len(QUESTIONS)}\n\n{q}", reply_markup=buttons(opts))
    else:
        await send_likert(message, state, 0)

async def send_likert(message: Message, state: FSMContext, idx):
    if idx < len(LIKERT):
        await state.update_data(l_index=idx)
        await state.set_state(Survey.likert)
        await message.answer(
            f"📊 Mezonlar ({idx+1}/{len(LIKERT)})\n\n"
            f"Bo‘lajak juftingizning quyidagi xususiyati siz uchun qanchalik muhim?\n\n"
            f"👉 “{LIKERT[idx][1]}”",
            reply_markup=likert_buttons()
        )
    else:
        await state.set_state(Survey.top5)
        await state.update_data(top5_selected=[])
        await message.answer(
            "🏆 TOP-5 — Eng muhim 5 ta mezonni tanlang\n\n"
            "Yuqoridagi 19 mezon ichidan eng ustuvor hisoblagan 5 tasini ketma-ketlikda bosing.\n\n"
            "⚠️ Birinchi bosilgani 1-o‘rin, ikkinchisi 2-o‘rin deb hisoblanadi.\n\n"
            "Tanlangan: 0/5",
            reply_markup=top5_keyboard([])
        )

def top5_keyboard(selected):
    kb = InlineKeyboardBuilder()
    for i, (_, label) in enumerate(LIKERT):
        prefix = "☑️ " if i in selected else "⬜ "
        kb.button(text=prefix + label, callback_data=f"top5:{i}")
    kb.button(text=f"Tanlangan: {len(selected)}/5", callback_data="top5:count")
    if len(selected) == 5:
        kb.button(text="✅ Tasdiqlash", callback_data="top5:confirm")
    kb.adjust(1)
    return kb.as_markup()

async def send_sd_question(message: Message, state: FSMContext, idx):
    if idx < len(SD_QUESTIONS):
        await state.update_data(sd_index=idx)
        await state.set_state(Survey.sd_scale)
        key, q_text, _ = SD_QUESTIONS[idx]
        await message.answer(
            f"📝 Qo‘shimcha savollar ({idx+1}/{len(SD_QUESTIONS)})\n\n"
            f"Quyidagi tasdiqqa munosabatingizni bildiring:\n\n"
            f"“{q_text}”",
            reply_markup=sd_buttons()
        )
    else:
        await state.set_state(Survey.main_factor)
        await message.answer(
            "🎯 Yakuniy savol\n\n"
            "Umuman olganda, juft tanlashda siz uchun qaysi umumiy yo‘nalish eng muhim?",
            reply_markup=buttons(MAIN_FACTORS, "factor")
        )

async def main():
    if not TOKEN or TOKEN == "YANGI_BOT_TOKENINI_SHUYERGA_YAZING":
        raise RuntimeError("Iltimos, koddagi TOKEN o'zgaruvchisiga yangi bot tokenini joylang!")
    
    # Baza jadvallarini yaratish
    db()

    # Render uchun aiohttp veb-serverini ishga tushirish
    app = web.Application()
    app.router.add_get('/', lambda r: web.Response(text="Bot ishlamoqda"))
    runner = web.AppRunner(app)
    await runner.setup()
    port = int(os.environ.get("PORT", 8080))
    site = web.TCPSite(runner, '0.0.0.0', port)
    await site.start()

    bot = Bot(TOKEN)
    dp = Dispatcher()

    @dp.message(CommandStart())
    async def start(message: Message, state: FSMContext):
        if message.from_user.id != ADMIN_ID and has_completed_response(message.from_user.id):
            await state.clear()
            await message.answer("Siz ushbu so‘rovnomada avval qatnashgansiz. Rahmat!")
            return
        await start_survey(message, state)

    @dp.message(Command("admin"))
    async def admin_panel(message: Message):
        if message.from_user.id != ADMIN_ID:
            return
        await message.answer("⚙️ **ADMIN PANEL**\n\nQuyidagi tugmalardan birini tanlang:", reply_markup=admin_keyboard())

    @dp.callback_query(F.data.startswith("admin:"))
    async def admin_actions(call: CallbackQuery, state: FSMContext):
        if call.from_user.id != ADMIN_ID:
            return
        
        action = call.data.split(":")[1]

        if action == "stats":
            conn = db()
            total = conn.execute("SELECT COUNT(*) FROM responses").fetchone()[0]
            today = datetime.now().date().isoformat()
            today_count = conn.execute("SELECT COUNT(*) FROM responses WHERE substr(created_at, 1, 10) = ?", (today,)).fetchone()[0]
            conn.close()
            await call.answer()
            await call.message.answer(f"📊 TADQIQOT STATISTIKASI\n\nJami respondentlar: {total}\nBugun qatnashganlar: {today_count}")

        elif action == "export":
            path, n = export_xlsx()
            await call.answer()
            await call.message.answer_document(FSInputFile(path), caption=f"Jami respondentlar: {n}\nExcel analitik statistik bazasi tayyor!")

        elif action == "reset":
            conn = db()
            conn.execute("DELETE FROM responses WHERE telegram_id = ?", (call.from_user.id,))
            conn.commit()
            conn.close()
            await state.clear()
            await call.answer("Sizning test javoblaringiz tozalandi!", show_alert=True)
            await call.message.answer("🧹 Faqat sizning test javoblaringiz bazadan o‘chirildi.")

        elif action == "reset_all":
            conn = db()
            conn.execute("DELETE FROM responses")
            conn.commit()
            conn.close()
            await state.clear()
            await call.answer("Barcha respondentlar javoblari to‘liq o‘chirildi!", show_alert=True)
            await call.message.answer("💥 **BAZA TO‘LIQ TOZALANDI!** Barcha respondentlar javoblari o‘chirib tashlandi.")

        elif action == "start_survey":
            await call.answer()
            await start_survey(call.message, state)

    @dp.message(Command("stats"))
    async def stats_cmd(message: Message):
        if message.from_user.id != ADMIN_ID:
            return
        conn = db()
        total = conn.execute("SELECT COUNT(*) FROM responses").fetchone()[0]
        today = datetime.now().date().isoformat()
        today_count = conn.execute("SELECT COUNT(*) FROM responses WHERE substr(created_at, 1, 10) = ?", (today,)).fetchone()[0]
        conn.close()
        await message.answer(f"📊 TADQIQOT STATISTIKASI\n\nJami respondentlar: {total}\nBugun qatnashganlar: {today_count}")

    @dp.message(Command("export"))
    async def export_cmd(message: Message):
        if message.from_user.id != ADMIN_ID:
            return
        path, n = export_xlsx()
        await message.answer_document(FSInputFile(path), caption=f"Jami respondentlar: {n}\nExcel analitik statistik bazasi tayyor!")

    @dp.message(Survey.name)
    async def get_name(message: Message, state: FSMContext):
        name = (message.text or "").strip()
        if not name:
            await message.answer("Iltimos, ismingizni kiriting.")
            return
        await state.update_data(name=name, q_index=0, answers={}, sd_score=0)
        await send_next_q(message, state)

    @dp.callback_query(Survey.question, F.data.startswith("ans:"))
    async def answer_question(call: CallbackQuery, state: FSMContext):
        data = await state.get_data()
        idx = data["q_index"]
        key, q, opts = QUESTIONS[idx]
        if opts is None:
            await call.answer("Bu savolga raqam bilan javob bering.", show_alert=True)
            return
        choice = int(call.data.split(":")[1])
        answers = data.get("answers", {})
        answers[key] = opts[choice]
        await state.update_data(answers=answers, q_index=idx+1)
        await call.answer()
        await call.message.edit_reply_markup(reply_markup=None)
        await send_next_q(call.message, state)

    @dp.message(Survey.question)
    async def answer_text_question(message: Message, state: FSMContext):
        data = await state.get_data()
        idx = data["q_index"]
        key, q, opts = QUESTIONS[idx]
        if opts is not None:
            await message.answer("Javobni pastdagi tugmalardan tanlang.")
            return
        try:
            val = int((message.text or "").strip())
            if not 15 <= val <= 60:
                raise ValueError
        except Exception:
            await message.answer("Iltimos, maqbul nikoh yoshini raqam bilan kiriting. Masalan: 25")
            return
        answers = data.get("answers", {})
        answers[key] = val
        await state.update_data(answers=answers, q_index=idx+1)
        await send_next_q(message, state)

    @dp.callback_query(Survey.likert, F.data.startswith("likert:"))
    async def answer_likert(call: CallbackQuery, state: FSMContext):
        data = await state.get_data()
        idx = data["l_index"]
        key, _ = LIKERT[idx]
        answers = data.get("answers", {})
        answers[key] = int(call.data.split(":")[1])
        await state.update_data(answers=answers)
        await call.answer()
        await call.message.edit_reply_markup(reply_markup=None)
        await send_likert(call.message, state, idx+1)

    @dp.callback_query(Survey.top5, F.data.startswith("top5:"))
    async def choose_top5(call: CallbackQuery, state: FSMContext):
        action = call.data.split(":", 1)[1]
        data = await state.get_data()
        selected = list(data.get("top5_selected", []))

        if action == "count":
            await call.answer(f"Tanlangan: {len(selected)}/5", show_alert=True)
            return

        if action == "confirm":
            if len(selected) != 5:
                await call.answer("Avval 5 ta xususiyatni tanlang.", show_alert=True)
                return
            ordered = [LIKERT[i][1] for i in selected]
            top5_text = " | ".join(f"{rank}-o‘rin: {label}" for rank, label in enumerate(ordered, 1))
            await state.update_data(top5=top5_text)
            await call.answer("TOP-5 saqlandi.")
            await call.message.edit_reply_markup(reply_markup=None)
            await send_sd_question(call.message, state, 0)
            return

        idx = int(action)
        if idx in selected:
            selected.remove(idx)
        else:
            if len(selected) >= 5:
                await call.answer("Ko'pi bilan 5 ta xususiyat tanlashingiz mumkin.", show_alert=True)
                return
            selected.append(idx)

        await state.update_data(top5_selected=selected)
        await call.answer()
        await call.message.edit_reply_markup(reply_markup=top5_keyboard(selected))

    @dp.callback_query(Survey.sd_scale, F.data.startswith("sd:"))
    async def answer_sd(call: CallbackQuery, state: FSMContext):
        data = await state.get_data()
        idx = data["sd_index"]
        ans_val = bool(int(call.data.split(":")[1]))
        
        _, _, expected_val = SD_QUESTIONS[idx]
        current_sd = data.get("sd_score", 0)
        
        if ans_val == expected_val:
            current_sd += 1
            
        await state.update_data(sd_score=current_sd)
        await call.answer()
        await call.message.edit_reply_markup(reply_markup=None)
        await send_sd_question(call.message, state, idx+1)

    @dp.callback_query(Survey.main_factor, F.data.startswith("factor:"))
    async def finish(call: CallbackQuery, state: FSMContext):
        data = await state.get_data()
        factor = MAIN_FACTORS[int(call.data.split(":")[1])]

        if factor == "Boshqa":
            await state.set_state(Survey.custom_factor)
            await call.answer()
            await call.message.edit_reply_markup(reply_markup=None)
            await call.message.answer("Iltimos, siz uchun muhim bo‘lgan boshqa yo‘nalishni yozing:")
            return

        answers = data.get("answers", {})
        answers["top5_ranking"] = data.get("top5", "")
        answers["sd_score"] = data.get("sd_score", 0)
        answers["main_factor"] = factor
        
        save_response(call.from_user.id, data.get("name", ""), answers)
        await call.answer()
        await call.message.edit_reply_markup(reply_markup=None)
        await state.clear()
        await call.message.answer("Rahmat! Javoblaringiz muvaffaqiyatli qabul qilindi.")

    @dp.message(Survey.custom_factor)
    async def finish_custom_factor(message: Message, state: FSMContext):
        factor = (message.text or "").strip()
        if not factor:
            await message.answer("Iltimos, javobingizni yozing.")
            return

        data = await state.get_data()
        answers = data.get("answers", {})
        answers["top5_ranking"] = data.get("top5", "")
        answers["sd_score"] = data.get("sd_score", 0)
        answers["main_factor"] = factor
        
        save_response(message.from_user.id, data.get("name", ""), answers)
        await state.clear()
        await message.answer("Rahmat! Javoblaringiz muvaffaqiyatli qabul qilindi.")

    await dp.start_polling(bot)

if __name__ == "__main__":
    import asyncio
    asyncio.run(main())
