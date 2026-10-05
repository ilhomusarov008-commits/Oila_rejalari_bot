import os
import json
import sqlite3
import asyncio
from datetime import datetime

from aiohttp import web

from aiogram import Bot, Dispatcher, F
from aiogram.filters import Command, CommandStart
from aiogram.types import Message, CallbackQuery
from aiogram.fsm.state import State, StatesGroup
from aiogram.fsm.context import FSMContext
from aiogram.utils.keyboard import InlineKeyboardBuilder

from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter


# ============================================================
# 1. SOZLAMALAR
# ============================================================

TOKEN = os.getenv("BOT_TOKEN")

ADMIN_ID = int(os.getenv("ADMIN_ID", "8671467518"))

DB_PATH = os.getenv(
    "DB_PATH",
    "oila_va_nikoh_tadqiqoti.db"
)

EXPORT_PATH = os.getenv(
    "EXPORT_PATH",
    "Oila_va_nikoh_statistik_tahlil.xlsx"
)

if not TOKEN:
    raise RuntimeError(
        "BOT_TOKEN environment variable o'rnatilmagan!"
    )


# ============================================================
# 2. DEMOGRAFIK SAVOLLAR
#    BU SAVOLLAR 25 TA TESTGA KIRMAYDI
# ============================================================

GENERAL_QUESTIONS = [
    (
        "age_group",
        "Yoshingiz qaysi oraliqda?",
        [
            "17–19",
            "20–22",
            "23–25",
            "26 va undan katta"
        ]
    ),
    (
        "course",
        "Kursingiz?",
        [
            "1-kurs",
            "2-kurs",
            "3-kurs",
            "4-kurs",
            "Magistratura / Boshqa"
        ]
    ),
    (
        "residence",
        "Doimiy yashash joyingiz?",
        [
            "Shahar",
            "Tuman/Qishloq"
        ]
    ),
    (
        "marital_status",
        "Hozirgi oilaviy holatingiz?",
        [
            "Turmush qurmagan",
            "Turmush qurgan"
        ]
    ),
    (
        "desired_children",
        "Kelajakda nechta farzandli bo‘lishni xohlaysiz?",
        [
            "0",
            "1",
            "2",
            "3",
            "4",
            "5 va undan ko‘p"
        ]
    )
]


# ============================================================
# 3. 25 TA TEST SAVOLLARI
#
# HAR BIRINING KEY'I ERKAK VA AYOL UCHUN BIR XIL.
# BU STATA VA JINSLARARO TAQQOSLASH UCHUN QULAY.
# ============================================================

MALE_RESPONDENT_CRITERIA = [

    (
        "kind_understanding",
        "Bo‘lajak ayolingizning mehribon, e’tiborli va tushunuvchan bo‘lishi siz uchun qanchalik muhim?"
    ),

    (
        "faithful_reliable",
        "Bo‘lajak ayolingizning vafodor, sadoqatli va munosabatlarda ishonchli bo‘lishi siz uchun qanchalik muhim?"
    ),

    (
        "honest_responsible",
        "Bo‘lajak ayolingizning halol, mas’uliyatli va va’dasida turadigan bo‘lishi siz uchun qanchalik muhim?"
    ),

    (
        "emotional_stability",
        "Bo‘lajak ayolingizning hissiy jihatdan barqaror, vazmin va sabrli bo‘lishi siz uchun qanchalik muhim?"
    ),

    (
        "pleasant_manner",
        "Bo‘lajak ayolingizning samimiy, xushmuomala va yoqimli fe’l-atvorli bo‘lishi siz uchun qanchalik muhim?"
    ),

    (
        "intelligence",
        "Bo‘lajak ayolingizning aql-idrokli, mustaqil fikrlaydigan va oqilona qaror qabul qila olishi siz uchun qanchalik muhim?"
    ),

    (
        "education_worldview",
        "Bo‘lajak ayolingizning yaxshi ta’limga ega, bilimli va keng dunyoqarashli bo‘lishi siz uchun qanchalik muhim?"
    ),

    (
        "self_development",
        "Bo‘lajak ayolingizning o‘z ustida ishlashi, bilim va ko‘nikmalarini rivojlantirishga intilishi siz uchun qanchalik muhim?"
    ),

    (
        "hardworking",
        "Bo‘lajak ayolingizning maqsadli, tirishqoq va mehnatsevar bo‘lishi siz uchun qanchalik muhim?"
    ),

    (
        "communication",
        "Bo‘lajak ayolingizning siz bilan ochiq muloqot qilishi, fikringizni tinglashi va kelishmovchiliklarni tinch yo‘l bilan hal qila olishi siz uchun qanchalik muhim?"
    ),

    (
        "healthy_lifestyle",
        "Bo‘lajak ayolingizning sog‘lig‘iga e’tibor berishi va sog‘lom turmush tarziga amal qilishi siz uchun qanchalik muhim?"
    ),

    (
        "neatness",
        "Bo‘lajak ayolingizning ozoda, sarishta va saranjom bo‘lishi siz uchun qanchalik muhim?"
    ),

    (
        "physical_attractiveness",
        "Bo‘lajak ayolingizning tashqi ko‘rinishi va jismoniy jozibadorligi siz uchun qanchalik muhim?"
    ),

    (
        "age_compatibility",
        "Bo‘lajak ayolingizning yoshi sizning yoshingizga mos bo‘lishi siz uchun qanchalik muhim?"
    ),

    (
        "family_values",
        "Bo‘lajak ayolingizning oilaviy qadriyatlarni hurmat qilishi va oilani muhim deb bilishi siz uchun qanchalik muhim?"
    ),

    (
        "respect_elders",
        "Bo‘lajak ayolingizning ota-ona, katta yoshdagilar va qarindoshlarga hurmat bilan munosabatda bo‘lishi siz uchun qanchalik muhim?"
    ),

    (
        "marriage_orientation",
        "Bo‘lajak ayolingizning nikoh va oilaviy hayotga jiddiy munosabatda bo‘lishi siz uchun qanchalik muhim?"
    ),

    (
        "family_responsibility",
        "Bo‘lajak ayolingizning oilaviy qarorlar va majburiyatlarda mas’uliyatli ishtirok etishi siz uchun qanchalik muhim?"
    ),

    (
        "children_responsibility",
        "Bo‘lajak ayolingizning farzand tarbiyasida faol ishtirok etishi va farzandlarga mas’uliyat bilan munosabatda bo‘lishi siz uchun qanchalik muhim?"
    ),

    (
        "household_management",
        "Bo‘lajak ayolingizning uy-ro‘zg‘or ishlarini tashkil etish va oilaviy hayotni tartibli yurita olish ko‘nikmalariga ega bo‘lishi siz uchun qanchalik muhim?"
    ),

    (
        "financial_management",
        "Bo‘lajak ayolingizning oilaviy mablag‘larni tejashi va oilaviy budjetni oqilona boshqara olishi siz uchun qanchalik muhim?"
    ),

    (
        "financial_contribution",
        "Bo‘lajak ayolingizning oilaviy daromadga hissa qo‘shish imkoniyati va istagi siz uchun qanchalik muhim?"
    ),

    (
        "career_attitude",
        "Bo‘lajak ayolingizning kasbiy rivojlanishga intilishi va oila bilan ish o‘rtasida muvozanatni saqlashi siz uchun qanchalik muhim?"
    ),

    (
        "social_behavior",
        "Bo‘lajak ayolingizning jamiyatda o‘zini tutishi, obro‘si va atrofdagilar bilan muomala madaniyati siz uchun qanchalik muhim?"
    ),

    (
        "partner_support",
        "Bo‘lajak ayolingizning sizning ta’limingiz, ishingiz, kasbiy rivojlanishingiz va shaxsiy maqsadlaringizni qo‘llab-quvvatlashi siz uchun qanchalik muhim?"
    )
]


FEMALE_RESPONDENT_CRITERIA = [

    (
        "kind_understanding",
        "Bo‘lajak eringizning mehribon, e’tiborli va sizni tushuna oladigan bo‘lishi siz uchun qanchalik muhim?"
    ),

    (
        "faithful_reliable",
        "Bo‘lajak eringizning vafodor, sadoqatli va munosabatlarda ishonchli bo‘lishi siz uchun qanchalik muhim?"
    ),

    (
        "honest_responsible",
        "Bo‘lajak eringizning halol, va’dasida turadigan va mas’uliyatli bo‘lishi siz uchun qanchalik muhim?"
    ),

    (
        "emotional_stability",
        "Bo‘lajak eringizning hissiy jihatdan barqaror, vazmin va qiyin vaziyatlarda o‘zini boshqara oladigan bo‘lishi siz uchun qanchalik muhim?"
    ),

    (
        "pleasant_manner",
        "Bo‘lajak eringizning xushmuomala, samimiy va yoqimli fe’l-atvorli bo‘lishi siz uchun qanchalik muhim?"
    ),

    (
        "intelligence",
        "Bo‘lajak eringizning aql-idrokli, mustaqil fikrlaydigan va muammolarga oqilona yondasha oladigan bo‘lishi siz uchun qanchalik muhim?"
    ),

    (
        "education_worldview",
        "Bo‘lajak eringizning yaxshi ta’limga ega, bilimli va keng dunyoqarashli bo‘lishi siz uchun qanchalik muhim?"
    ),

    (
        "self_development",
        "Bo‘lajak eringizning o‘z ustida ishlashi, bilim va ko‘nikmalarini rivojlantirishga intilishi siz uchun qanchalik muhim?"
    ),

    (
        "hardworking",
        "Bo‘lajak eringizning maqsadli, tashabbuskor va mehnatsevar bo‘lishi siz uchun qanchalik muhim?"
    ),

    (
        "communication",
        "Bo‘lajak eringizning siz bilan ochiq muloqot qilishi, fikringizni tinglashi va kelishmovchiliklarni tinch yo‘l bilan hal qila olishi siz uchun qanchalik muhim?"
    ),

    (
        "healthy_lifestyle",
        "Bo‘lajak eringizning sog‘lig‘iga e’tibor berishi va sog‘lom turmush tarziga amal qilishi siz uchun qanchalik muhim?"
    ),

    (
        "neatness",
        "Bo‘lajak eringizning ozoda, saranjom va o‘ziga e’tiborli bo‘lishi siz uchun qanchalik muhim?"
    ),

    (
        "physical_attractiveness",
        "Bo‘lajak eringizning tashqi ko‘rinishi va jismoniy jozibadorligi siz uchun qanchalik muhim?"
    ),

    (
        "age_compatibility",
        "Bo‘lajak eringizning yoshi sizning yoshingizga mos bo‘lishi siz uchun qanchalik muhim?"
    ),

    (
        "family_values",
        "Bo‘lajak eringizning oilaviy qadriyatlarni hurmat qilishi va oilani muhim deb bilishi siz uchun qanchalik muhim?"
    ),

    (
        "respect_elders",
        "Bo‘lajak eringizning ota-ona, katta yoshdagilar va qarindoshlarga hurmat bilan munosabatda bo‘lishi siz uchun qanchalik muhim?"
    ),

    (
        "marriage_orientation",
        "Bo‘lajak eringizning nikoh va oilaviy hayotga jiddiy tayyorgarlik ko‘rgan bo‘lishi siz uchun qanchalik muhim?"
    ),

    (
        "family_responsibility",
        "Bo‘lajak eringizning oilaviy qarorlar va majburiyatlarda mas’uliyatli ishtirok etishi siz uchun qanchalik muhim?"
    ),

    (
        "children_responsibility",
        "Bo‘lajak eringizning farzand tarbiyasida faol ishtirok etishi va farzandlarga mas’uliyat bilan munosabatda bo‘lishi siz uchun qanchalik muhim?"
    ),

    (
        "household_management",
        "Bo‘lajak eringizning uy-ro‘zg‘or ishlarida ishtirok etishga va oilaviy vazifalarni jufti bilan bo‘lishishga tayyorligi siz uchun qanchalik muhim?"
    ),

    (
        "financial_management",
        "Bo‘lajak eringizning oilaviy daromad va xarajatlarni rejalashtira olishi, pulni oqilona boshqarishi siz uchun qanchalik muhim?"
    ),

    (
        "financial_contribution",
        "Bo‘lajak eringizning barqaror daromad topish imkoniyati va oilaning moddiy farovonligini ta’minlash qobiliyati siz uchun qanchalik muhim?"
    ),

    (
        "career_attitude",
        "Bo‘lajak eringizning kasbiy faoliyatida rivojlanishga intilishi, lekin oila va ish o‘rtasida muvozanatni saqlay olishi siz uchun qanchalik muhim?"
    ),

    (
        "social_behavior",
        "Bo‘lajak eringizning jamiyatdagi obro‘si, atrofdagilar bilan muomala madaniyati va ijtimoiy xulqi siz uchun qanchalik muhim?"
    ),

    (
        "partner_support",
        "Bo‘lajak eringizning sizning ta’lim olishingiz, ishlashingiz, kasbiy rivojlanishingiz va shaxsiy maqsadlaringizni qo‘llab-quvvatlashi siz uchun qanchalik muhim?"
    )
]


# ============================================================
# 4. LIKERT
# ============================================================

LIKERT = [
    ("0 — Umuman muhim emas", 0),
    ("1 — Unchalik muhim emas", 1),
    ("2 — Muhim", 2),
    ("3 — Juda muhim", 3)
]


# ============================================================
# 5. SOCIAL DESIRABILITY
#    BU SAVOLLAR 25 TA TESTGA KIRMAYDI
# ============================================================

SD_QUESTIONS = [

    (
        "sd1",
        "Hech qachon birovga nisbatan ich-ichimdan g‘azab yoki nafrat sezmaganman.",
        True
    ),

    (
        "sd2",
        "Gohida bajara olmaydigan va’dalarni berib qo‘yaman.",
        False
    ),

    (
        "sd3",
        "Har doim o‘z xatolarimni ochiq tan olaman.",
        True
    ),

    (
        "sd4",
        "Ba’zan atrofimdagilar haqida g‘iybat qilishim yoki g‘iybatni eshitishim mumkin.",
        False
    ),

    (
        "sd5",
        "Har qanday vaziyatda ham samimiy va xushfe’l bo‘lishga intilaman.",
        True
    ),

    (
        "sd6",
        "Ba’zan kayfiyatim yomon bo‘lsa, atrofdagilarga qo‘pollik qilib qo‘yaman.",
        False
    )
]


# ============================================================
# 6. YAKUNIY OMILLAR
# ============================================================

MAIN_FACTORS = [
    "Shaxsiy xarakter va odob",
    "Intellekt va ta’lim",
    "Tashqi ko‘rinish",
    "Oilaviy qadriyatlar va tarbiya",
    "Ijtimoiy-iqtisodiy va ro‘zg‘or omillari",
    "Boshqa"
]


# ============================================================
# 7. DATABASE
# ============================================================

def get_db():

    conn = sqlite3.connect(DB_PATH)

    conn.execute("""
        CREATE TABLE IF NOT EXISTS responses (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            telegram_id INTEGER,
            name TEXT,
            gender TEXT,
            created_at TEXT,
            data TEXT
        )
    """)

    conn.commit()

    return conn


def save_response(user_id, name, gender, data):

    conn = get_db()

    conn.execute("""
        INSERT INTO responses
        (
            telegram_id,
            name,
            gender,
            created_at,
            data
        )
        VALUES (?, ?, ?, ?, ?)
    """, (
        user_id,
        name,
        gender,
        datetime.now().isoformat(timespec="seconds"),
        json.dumps(data, ensure_ascii=False)
    ))

    conn.commit()
    conn.close()


def already_completed(user_id):

    conn = get_db()

    row = conn.execute("""
        SELECT id
        FROM responses
        WHERE telegram_id = ?
        LIMIT 1
    """, (user_id,)).fetchone()

    conn.close()

    return row is not None


# ============================================================
# 8. KLAVIATURALAR
# ============================================================

def make_buttons(items, prefix):

    kb = InlineKeyboardBuilder()

    for i, item in enumerate(items):

        kb.button(
            text=item,
            callback_data=f"{prefix}:{i}"
        )

    kb.adjust(1)

    return kb.as_markup()


def gender_keyboard():

    kb = InlineKeyboardBuilder()

    kb.button(
        text="👨 Erkak",
        callback_data="gender:Erkak"
    )

    kb.button(
        text="👩 Ayol",
        callback_data="gender:Ayol"
    )

    kb.adjust(2)

    return kb.as_markup()


def skip_name_keyboard():

    kb = InlineKeyboardBuilder()

    kb.button(
        text="⏭ Ismni ko‘rsatmasdan davom etish",
        callback_data="name:skip"
    )

    return kb.as_markup()


def likert_keyboard():

    kb = InlineKeyboardBuilder()

    for text, value in LIKERT:

        kb.button(
            text=text,
            callback_data=f"likert:{value}"
        )

    kb.adjust(1)

    return kb.as_markup()


def sd_keyboard():

    kb = InlineKeyboardBuilder()

    kb.button(
        text="Ha",
        callback_data="sd:1"
    )

    kb.button(
        text="Yo‘q",
        callback_data="sd:0"
    )

    kb.adjust(2)

    return kb.as_markup()


def top5_keyboard(selected, criteria):

    kb = InlineKeyboardBuilder()

    for i, (_, label) in enumerate(criteria):

        if i in selected:
            prefix = "☑️ "
        else:
            prefix = "⬜ "

        short_label = label.replace(
            "Bo‘lajak ayolingizning ", ""
        ).replace(
            "Bo‘lajak eringizning ", ""
        )

        kb.button(
            text=prefix + short_label[:55],
            callback_data=f"top5:{i}"
        )

    kb.button(
        text=f"📌 Tanlangan: {len(selected)}/5",
        callback_data="top5:count"
    )

    if len(selected) == 5:

        kb.button(
            text="✅ TOP-5 ni tasdiqlash",
            callback_data="top5:confirm"
        )

    kb.adjust(1)

    return kb.as_markup()


def admin_keyboard():

    kb = InlineKeyboardBuilder()

    kb.button(
        text="📊 Statistika",
        callback_data="admin:stats"
    )

    kb.button(
        text="📥 Excel eksport",
        callback_data="admin:export"
    )

    kb.button(
        text="🧹 Mening testimni o‘chirish",
        callback_data="admin:reset"
    )

    kb.button(
        text="💥 BARCHA MA’LUMOTNI O‘CHIRISH",
        callback_data="admin:reset_all"
    )

    kb.button(
        text="🔄 So‘rovnomani boshlash",
        callback_data="admin:start"
    )

    kb.adjust(1)

    return kb.as_markup()


# ============================================================
# 9. FSM STATES
# ============================================================

class Survey(StatesGroup):

    name = State()

    gender = State()

    general = State()

    marriage_age = State()

    likert = State()

    top5 = State()

    sd = State()

    factor = State()

    custom_factor = State()


# ============================================================
# 10. SO‘ROVNOMANI BOSHLASH
# ============================================================

async def start_survey(message, state):

    await state.clear()

    await state.set_state(
        Survey.name
    )

    await message.answer(
        "📋 OILA VA NIKOH MEZONLARI BO‘YICHA SO‘ROVNOMA\n\n"
        "Ushbu so‘rovnoma ilmiy tadqiqot maqsadida o‘tkaziladi.\n\n"
        "Ismingizni kiritishingiz mumkin. "
        "Ism ko‘rsatish majburiy emas.",
        reply_markup=skip_name_keyboard()
    )


# ============================================================
# 11. UMUMIY SAVOLNI YUBORISH
# ============================================================

async def send_general(message, state, index):

    if index >= len(GENERAL_QUESTIONS):

        await state.set_state(
            Survey.marriage_age
        )

        await message.answer(
            "📌 Nikoh qurishning maqbul yoshi\n\n"
            "Siz uchun nikoh qurishning maqbul yoshi nechada?\n\n"
            "Masalan: 25"
        )

        return

    key, question, options = GENERAL_QUESTIONS[index]

    await state.update_data(
        general_index=index
    )

    await state.set_state(
        Survey.general
    )

    await message.answer(
        f"📌 DEMOGRAFIK SAVOL {index + 1}/{len(GENERAL_QUESTIONS)}\n\n"
        f"{question}",
        reply_markup=make_buttons(
            options,
            "general"
        )
    )


# ============================================================
# 12. TEST SAVOLINI YUBORISH
# ============================================================

async def send_test_question(message, state, index):

    data = await state.get_data()

    gender = data.get("gender")

    if gender == "Ayol":

        criteria = FEMALE_RESPONDENT_CRITERIA
        target = "ideal er"

    else:

        criteria = MALE_RESPONDENT_CRITERIA
        target = "ideal ayol"

    if index >= 25:

        await state.update_data(
            top5_selected=[]
        )

        await state.set_state(
            Survey.top5
        )

        await message.answer(
            "🏆 III-BO‘LIM — TOP-5\n\n"
            "Yuqoridagi 25 ta test mezonidan "
            "siz uchun ENG MUHIM 5 TASINI tanlang.\n\n"
            "Tanlash tartibi ham saqlanadi:\n"
            "1-tanlov → 1-o‘rin\n"
            "2-tanlov → 2-o‘rin\n"
            "3-tanlov → 3-o‘rin\n"
            "4-tanlov → 4-o‘rin\n"
            "5-tanlov → 5-o‘rin",
            reply_markup=top5_keyboard(
                [],
                criteria
            )
        )

        return

    key, question = criteria[index]

    await state.update_data(
        likert_index=index
    )

    await state.set_state(
        Survey.likert
    )

    await message.answer(
        f"📊 II-BO‘LIM — TEST {index + 1}/25\n\n"
        f"Bo‘lajak {target}ning ushbu xususiyati "
        f"siz uchun qanchalik muhim?\n\n"
        f"👉 {question}",
        reply_markup=likert_keyboard()
    )


# ============================================================
# 13. SD SAVOLINI YUBORISH
# ============================================================

async def send_sd_question(message, state, index):

    if index >= len(SD_QUESTIONS):

        await state.set_state(
            Survey.factor
        )

        await message.answer(
            "🎯 V-BO‘LIM — YAKUNIY SAVOL\n\n"
            "Umuman olganda, juft tanlashda "
            "siz uchun qaysi yo‘nalish eng muhim?",
            reply_markup=make_buttons(
                MAIN_FACTORS,
                "factor"
            )
        )

        return

    key, question, expected = SD_QUESTIONS[index]

    await state.update_data(
        sd_index=index
    )

    await state.set_state(
        Survey.sd
    )

    await message.answer(
        f"📝 IV-BO‘LIM — {index + 1}/{len(SD_QUESTIONS)}\n\n"
        f"“{question}”",
        reply_markup=sd_keyboard()
    )


# ============================================================
# 14. BOT
# ============================================================

async def main():

    get_db()

    bot = Bot(TOKEN)

    dp = Dispatcher()

    # ========================================================
    # RENDER HEALTH SERVER
    # ========================================================

    app = web.Application()

    async def health(request):

        return web.Response(
            text="Bot ishlamoqda"
        )

    app.router.add_get(
        "/",
        health
    )

    runner = web.AppRunner(app)

    await runner.setup()

    port = int(
        os.environ.get(
            "PORT",
            "8080"
        )
    )

    site = web.TCPSite(
        runner,
        "0.0.0.0",
        port
    )

    await site.start()

    # ========================================================
    # START
    # ========================================================

    @dp.message(CommandStart())
    async def command_start(message: Message, state: FSMContext):

        user_id = message.from_user.id

        if (
            user_id != ADMIN_ID
            and already_completed(user_id)
        ):

            await state.clear()

            await message.answer(
                "Siz ushbu so‘rovnomada avval qatnashgansiz.\n\n"
                "Ishtirokingiz uchun rahmat!"
            )

            return

        await start_survey(
            message,
            state
        )

    # ========================================================
    # ADMIN
    # ========================================================

    @dp.message(Command("admin"))
    async def admin_command(message: Message):

        if message.from_user.id != ADMIN_ID:
            return

        await message.answer(
            "⚙️ ADMIN PANEL\n\n"
            "Kerakli amalni tanlang:",
            reply_markup=admin_keyboard()
        )

    # ========================================================
    # ISM
    # ========================================================

    @dp.callback_query(
        Survey.name,
        F.data == "name:skip"
    )
    async def skip_name(
        call: CallbackQuery,
        state: FSMContext
    ):

        await state.update_data(
            name="Ko‘rsatilmagan"
        )

        await state.set_state(
            Survey.gender
        )

        await call.answer()

        await call.message.edit_reply_markup(
            reply_markup=None
        )

        await call.message.answer(
            "👤 Jinsingizni tanlang:",
            reply_markup=gender_keyboard()
        )

    @dp.message(Survey.name)
    async def get_name(
        message: Message,
        state: FSMContext
    ):

        name = (
            message.text or ""
        ).strip()

        if not name:

            await message.answer(
                "Ismni kiritishingiz yoki "
                "“Ismni ko‘rsatmasdan davom etish” tugmasini bosishingiz mumkin.",
                reply_markup=skip_name_keyboard()
            )

            return

        await state.update_data(
            name=name
        )

        await state.set_state(
            Survey.gender
        )

        await message.answer(
            "👤 Jinsingizni tanlang:",
            reply_markup=gender_keyboard()
        )

    # ========================================================
    # JINS
    # ========================================================

    @dp.callback_query(
        Survey.gender,
        F.data.startswith("gender:")
    )
    async def get_gender(
        call: CallbackQuery,
        state: FSMContext
    ):

        gender = call.data.split(
            ":",
            1
        )[1]

        await state.update_data(
            gender=gender,
            general_index=0,
            general_answers={},
            test_answers={},
            top5_selected=[],
            sd_answers={}
        )

        await call.answer()

        await call.message.edit_reply_markup(
            reply_markup=None
        )

        await send_general(
            call.message,
            state,
            0
        )

    # ========================================================
    # DEMOGRAFIKA
    # ========================================================

    @dp.callback_query(
        Survey.general,
        F.data.startswith("general:")
    )
    async def get_general(
        call: CallbackQuery,
        state: FSMContext
    ):

        index = int(
            call.data.split(":")[1]
        )

        data = await state.get_data()

        questions = GENERAL_QUESTIONS

        key = questions[index][0]

        general_answers = data.get(
            "general_answers",
            {}
        )

        general_answers[key] = questions[index][2][index]

        await state.update_data(
            general_answers=general_answers,
            general_index=index + 1
        )

        await call.answer()

        await call.message.edit_reply_markup(
            reply_markup=None
        )

        await send_general(
            call.message,
            state,
            index + 1
        )

    # ========================================================
    # NIKOH YOSHI
    # ========================================================

    @dp.message(Survey.marriage_age)
    async def get_marriage_age(
        message: Message,
        state: FSMContext
    ):

        text = (
            message.text or ""
        ).strip()

        try:

            age = int(text)

            if age < 15 or age > 60:
                raise ValueError

        except ValueError:

            await message.answer(
                "Iltimos, 15 dan 60 gacha bo‘lgan "
                "yoshni raqam bilan kiriting.\n"
                "Masalan: 25"
            )

            return

        await state.update_data(
            ideal_marriage_age=age
        )

        await send_test_question(
            message,
            state,
            0
        )

    # ========================================================
    # 25 TA LIKERT TEST
    # ========================================================

    @dp.callback_query(
        Survey.likert,
        F.data.startswith("likert:")
    )
    async def get_likert(
        call: CallbackQuery,
        state: FSMContext
    ):

        value = int(
            call.data.split(":")[1]
        )

        data = await state.get_data()

        index = data.get(
            "likert_index",
            0
        )

        gender = data.get(
            "gender"
        )

        if gender == "Ayol":
            criteria = FEMALE_RESPONDENT_CRITERIA
        else:
            criteria = MALE_RESPONDENT_CRITERIA

        key = criteria[index][0]

        test_answers = data.get(
            "test_answers",
            {}
        )

        test_answers[key] = value

        await state.update_data(
            test_answers=test_answers
        )

        await call.answer()

        await call.message.edit_reply_markup(
            reply_markup=None
        )

        await send_test_question(
            call.message,
            state,
            index + 1
        )

    # ========================================================
    # TOP-5
    # ========================================================

    @dp.callback_query(
        Survey.top5,
        F.data.startswith("top5:")
    )
    async def top5_handler(
        call: CallbackQuery,
        state: FSMContext
    ):

        action = call.data.split(
            ":",
            1
        )[1]

        data = await state.get_data()

        gender = data.get(
            "gender"
        )

        if gender == "Ayol":
            criteria = FEMALE_RESPONDENT_CRITERIA
        else:
            criteria = MALE_RESPONDENT_CRITERIA

        selected = data.get(
            "top5_selected",
            []
        )

        # ----------------------------------------------------
        # COUNT
        # ----------------------------------------------------

        if action == "count":

            await call.answer(
                f"Tanlangan: {len(selected)}/5"
            )

            return

        # ----------------------------------------------------
        # CONFIRM
        # ----------------------------------------------------

        if action == "confirm":

            if len(selected) != 5:

                await call.answer(
                    "Avval 5 ta mezon tanlang.",
                    show_alert=True
                )

                return

            top5_data = []

            for rank, index in enumerate(
                selected,
                start=1
            ):

                key, label = criteria[index]

                top5_data.append({
                    "rank": rank,
                    "key": key,
                    "label": label
                })

            await state.update_data(
                top5=top5_data,
                sd_index=0
            )

            await call.answer()

            await call.message.edit_reply_markup(
                reply_markup=None
            )

            await send_sd_question(
                call.message,
                state,
                0
            )

            return

        # ----------------------------------------------------
        # SELECT
        # ----------------------------------------------------

        index = int(action)

        if index in selected:

            selected.remove(index)

        else:

            if len(selected) >= 5:

                await call.answer(
                    "Faqat 5 ta mezon tanlash mumkin.",
                    show_alert=True
                )

                return

            selected.append(index)

        await state.update_data(
            top5_selected=selected
        )

        await call.answer()

        await call.message.edit_reply_markup(
            reply_markup=top5_keyboard(
                selected,
                criteria
            )
        )

    # ========================================================
    # SD
    # ========================================================

    @dp.callback_query(
        Survey.sd,
        F.data.startswith("sd:")
    )
    async def get_sd(
        call: CallbackQuery,
        state: FSMContext
    ):

        answer = int(
            call.data.split(":")[1]
        )

        data = await state.get_data()

        index = data.get(
            "sd_index",
            0
        )

        key, question, expected = SD_QUESTIONS[index]

        sd_answers = data.get(
            "sd_answers",
            {}
        )

        sd_answers[key] = answer

        await state.update_data(
            sd_answers=sd_answers
        )

        await call.answer()

        await call.message.edit_reply_markup(
            reply_markup=None
        )

        await send_sd_question(
            call.message,
            state,
            index + 1
        )

    # ========================================================
    # YAKUNIY OMIL
    # ========================================================

    @dp.callback_query(
        Survey.factor,
        F.data.startswith("factor:")
    )
    async def get_factor(
        call: CallbackQuery,
        state: FSMContext
    ):

        index = int(
            call.data.split(":")[1]
        )

        factor = MAIN_FACTORS[index]

        await call.answer()

        await call.message.edit_reply_markup(
            reply_markup=None
        )

        if factor == "Boshqa":

            await state.set_state(
                Survey.custom_factor
            )

            await call.message.answer(
                "✍️ Siz uchun boshqa qaysi omil muhim?\n\n"
                "Javobingizni yozing:"
            )

            return

        await finish_survey(
            call.message,
            state,
            factor
        )

    # ========================================================
    # BOSHQA OMIL
    # ========================================================

    @dp.message(Survey.custom_factor)
    async def custom_factor(
        message: Message,
        state: FSMContext
    ):

        factor = (
            message.text or ""
        ).strip()

        if not factor:

            await message.answer(
                "Iltimos, javobingizni yozing."
            )

            return

        await finish_survey(
            message,
            state,
            factor
        )

    # ========================================================
    # ADMIN CALLBACKS
    # ========================================================

    @dp.callback_query(
        F.data.startswith("admin:")
    )
    async def admin_actions(
        call: CallbackQuery,
        state: FSMContext
    ):

        if call.from_user.id != ADMIN_ID:

            await call.answer(
                "Sizda ruxsat yo‘q.",
                show_alert=True
            )

            return

        action = call.data.split(
            ":",
            1
        )[1]

        await call.answer()

        # ----------------------------------------------------
        # STATS
        # ----------------------------------------------------

        if action == "stats":

            conn = get_db()

            total = conn.execute(
                "SELECT COUNT(*) FROM responses"
            ).fetchone()[0]

            male = conn.execute(
                "SELECT COUNT(*) FROM responses WHERE gender='Erkak'"
            ).fetchone()[0]

            female = conn.execute(
                "SELECT COUNT(*) FROM responses WHERE gender='Ayol'"
            ).fetchone()[0]

            conn.close()

            await call.message.answer(
                "📊 STATISTIKA\n\n"
                f"Jami respondentlar: {total}\n"
                f"Erkaklar: {male}\n"
                f"Ayollar: {female}"
            )

        # ----------------------------------------------------
        # EXPORT
        # ----------------------------------------------------

        elif action == "export":

            try:

                path = create_excel()

                from aiogram.types import FSInputFile

                await call.message.answer_document(
                    FSInputFile(path),
                    caption=(
                        "📊 To‘liq statistik Excel fayl.\n\n"
                        "Unda Stata uchun kodlangan ma’lumotlar, "
                        "Likert tahlili, TOP-5, SD va jinslararo "
                        "taqqoslash mavjud."
                    )
                )

            except Exception as e:

                await call.message.answer(
                    f"❌ Excel yaratishda xatolik:\n{e}"
                )

        # ----------------------------------------------------
        # RESET CURRENT ADMIN
        # ----------------------------------------------------

        elif action == "reset":

            conn = get_db()

            conn.execute(
                "DELETE FROM responses WHERE telegram_id=?",
                (ADMIN_ID,)
            )

            conn.commit()
            conn.close()

            await call.message.answer(
                "✅ Sizning testingiz o‘chirildi."
            )

        # ----------------------------------------------------
        # RESET ALL
        # ----------------------------------------------------

        elif action == "reset_all":

            conn = get_db()

            conn.execute(
                "DELETE FROM responses"
            )

            conn.commit()
            conn.close()

            await call.message.answer(
                "💥 Barcha respondent ma’lumotlari o‘chirildi."
            )

        # ----------------------------------------------------
        # START
        # ----------------------------------------------------

        elif action == "start":

            await start_survey(
                call.message,
                state
            )

    # ========================================================
    # POLLING
    # ========================================================

    try:

        await dp.start_polling(
            bot
        )

    finally:

        await bot.session.close()

        await runner.cleanup()


# ============================================================
# 15. SO‘ROVNOMANI YAKUNLASH
# ============================================================

async def finish_survey(
    message,
    state,
    factor
):

    data = await state.get_data()

    user_id = message.from_user.id

    name = data.get(
        "name",
        "Ko‘rsatilmagan"
    )

    gender = data.get(
        "gender",
        ""
    )

    data["main_factor"] = factor

    save_response(
        user_id,
        name,
        gender,
        data
    )

    await state.clear()

    await message.answer(
        "✅ SO‘ROVNOMA YAKUNLANDI!\n\n"
        "Ishtirokingiz uchun katta rahmat.\n\n"
        "Sizning javoblaringiz ilmiy tadqiqot "
        "va statistik tahlil uchun saqlab qo‘yildi."
    )


# ============================================================
# 16. EXCEL EKSPORT
# ============================================================

def create_excel():

    conn = get_db()

    rows = conn.execute("""
        SELECT
            id,
            telegram_id,
            name,
            gender,
            created_at,
            data
        FROM responses
        ORDER BY id
    """).fetchall()

    conn.close()

    wb = Workbook()

    # --------------------------------------------------------
    # STYLES
    # --------------------------------------------------------

    header_fill = PatternFill(
        "solid",
        fgColor="1F4E78"
    )

    header_font = Font(
        color="FFFFFF",
        bold=True
    )

    thin = Side(
        style="thin",
        color="B7B7B7"
    )

    border = Border(
        left=thin,
        right=thin,
        top=thin,
        bottom=thin
    )

    # --------------------------------------------------------
    # 01 RAW
    # --------------------------------------------------------

    ws = wb.active

    ws.title = "01_Barcha_Malumotlar"

    raw_headers = [
        "ID",
        "Telegram_ID",
        "Ism",
        "Jins",
        "Sana",
        "Yosh guruhi",
        "Kurs",
        "Yashash joyi",
        "Oilaviy holat",
        "Nikoh yoshi",
        "Farzandlar soni"
    ]

    for i in range(25):

        raw_headers.append(
            f"TEST_{i+1}"
        )

    for i in range(5):

        raw_headers.append(
            f"TOP5_{i+1}"
        )

    raw_headers += [
        "SD_Score",
        "SD_Level",
        "Main_Factor"
    ]

    ws.append(raw_headers)

    for cell in ws[1]:

        cell.fill = header_fill
        cell.font = header_font
        cell.border = border
        cell.alignment = Alignment(
            horizontal="center",
            vertical="center"
        )

    for row in rows:

        data = json.loads(row[5])

        general = data.get(
            "general_answers",
            {}
        )

        tests = data.get(
            "test_answers",
            {}
        )

        top5 = data.get(
            "top5",
            []
        )

        sd_answers = data.get(
            "sd_answers",
            {}
        )

        sd_score = calculate_sd_score(
            sd_answers
        )

        raw = [
            row[0],
            row[1],
            row[2],
            row[3],
            row[4],
            general.get("age_group", ""),
            general.get("course", ""),
            general.get("residence", ""),
            general.get("marital_status", ""),
            data.get("ideal_marriage_age", ""),
            general.get("desired_children", "")
        ]

        criteria = (
            FEMALE_RESPONDENT_CRITERIA
            if row[3] == "Ayol"
            else MALE_RESPONDENT_CRITERIA
        )

        for key, label in criteria:

            raw.append(
                tests.get(key, "")
            )

        for item in top5:

            raw.append(
                item.get("key", "")
            )

        while len(
            raw
        ) < 11 + 25 + 5:

            raw.append("")

        raw.append(
            sd_score
        )

        raw.append(
            sd_level(sd_score)
        )

        raw.append(
            data.get("main_factor", "")
        )

        ws.append(raw)

    # --------------------------------------------------------
    # 02 STATA
    # --------------------------------------------------------

    ws2 = wb.create_sheet(
        "02_Stata_Kodlangan"
    )

    stata_headers = [
        "id",
        "gender",
        "age_group",
        "course",
        "residence",
        "marital_status",
        "ideal_marriage_age",
        "desired_children"
    ]

    for i in range(25):

        stata_headers.append(
            f"q{i+1}"
        )

    stata_headers += [
        "top1",
        "top2",
        "top3",
        "top4",
        "top5",
        "sd_score",
        "sd_level",
        "main_factor"
    ]

    ws2.append(stata_headers)

    for row in rows:

        data = json.loads(row[5])

        general = data.get(
            "general_answers",
            {}
        )

        tests = data.get(
            "test_answers",
            {}
        )

        criteria = (
            FEMALE_RESPONDENT_CRITERIA
            if row[3] == "Ayol"
            else MALE_RESPONDENT_CRITERIA
        )

        values = [
            row[0],
            1 if row[3] == "Erkak" else 2,
            encode_age(
                general.get("age_group", "")
            ),
            encode_course(
                general.get("course", "")
            ),
            encode_residence(
                general.get("residence", "")
            ),
            encode_marital(
                general.get("marital_status", "")
            ),
            data.get(
                "ideal_marriage_age",
                ""
            ),
            encode_children(
                general.get(
                    "desired_children",
                    ""
                )
            )
        ]

        for key, label in criteria:

            values.append(
                tests.get(key, "")
            )

        top5 = data.get(
            "top5",
            []
        )

        for i in range(5):

            if i < len(top5):

                values.append(
                    top5[i]["key"]
                )

            else:

                values.append("")

        sd_score = calculate_sd_score(
            data.get(
                "sd_answers",
                {}
            )
        )

        values.append(
            sd_score
        )

        values.append(
            encode_sd_level(
                sd_level(sd_score)
            )
        )

        values.append(
            encode_factor(
                data.get(
                    "main_factor",
                    ""
                )
            )
        )

        ws2.append(values)

    # --------------------------------------------------------
    # 03 DEMOGRAFIYA
    # --------------------------------------------------------

    ws3 = wb.create_sheet(
        "03_Demografiya"
    )

    ws3.append([
        "O‘zgaruvchi",
        "Qiymat",
        "Soni",
        "Foiz"
    ])

    demographics = [
        ("gender", "Jins"),
        ("age_group", "Yosh guruhi"),
        ("course", "Kurs"),
        ("residence", "Yashash joyi"),
        ("marital_status", "Oilaviy holat"),
        ("desired_children", "Farzandlar soni")
    ]

    for key, title in demographics:

        counter = {}

        for row in rows:

            data = json.loads(row[5])

            if key == "gender":

                value = row[3]

            else:

                value = data.get(
                    "general_answers",
                    {}
                ).get(
                    key,
                    ""
                )

            counter[value] = (
                counter.get(value, 0) + 1
            )

        total = sum(
            counter.values()
        )

        for value, count in counter.items():

            percent = (
                count / total * 100
                if total
                else 0
            )

            ws3.append([
                title,
                value,
                count,
                round(percent, 2)
            ])

    # --------------------------------------------------------
    # 04 LIKERT TAHLIL
    # --------------------------------------------------------

    ws4 = wb.create_sheet(
        "04_Likert_Tahlil"
    )

    ws4.append([
        "Savol",
        "Mezon",
        "N",
        "O‘rtacha",
        "Standart og‘ish",
        "0",
        "1",
        "2",
        "3",
        "3 baho ulushi (%)"
    ])

    for index in range(25):

        key_m, label_m = MALE_RESPONDENT_CRITERIA[index]

        key_f, label_f = FEMALE_RESPONDENT_CRITERIA[index]

        male_values = []
        female_values = []

        for row in rows:

            data = json.loads(row[5])

            value = data.get(
                "test_answers",
                {}
            ).get(
                key_m,
                None
            )

            if value is None:
                continue

            if row[3] == "Erkak":
                male_values.append(value)
            else:
                female_values.append(value)

        all_values = (
            male_values +
            female_values
        )

        if all_values:

            mean = sum(
                all_values
            ) / len(all_values)

            variance = sum(
                (x - mean) ** 2
                for x in all_values
            ) / len(all_values)

            std = variance ** 0.5

        else:

            mean = 0
            std = 0

        counts = [
            all_values.count(i)
            for i in range(4)
        ]

        share_3 = (
            counts[3] /
            len(all_values) *
            100
            if all_values
            else 0
        )

        ws4.append([
            index + 1,
            label_f,
            len(all_values),
            round(mean, 3),
            round(std, 3),
            counts[0],
            counts[1],
            counts[2],
            counts[3],
            round(share_3, 2)
        ])

    # --------------------------------------------------------
    # 05 TOP5
    # --------------------------------------------------------

    ws5 = wb.create_sheet(
        "05_TOP5"
    )

    ws5.append([
        "Mezon",
        "1-o‘rin",
        "2-o‘rin",
        "3-o‘rin",
        "4-o‘rin",
        "5-o‘rin",
        "Jami TOP-5"
    ])

    all_keys = [
        key
        for key, label in MALE_RESPONDENT_CRITERIA
    ]

    all_labels = {
        key: label
        for key, label in FEMALE_RESPONDENT_CRITERIA
    }

    for key in all_keys:

        counts = [0] * 5

        for row in rows:

            data = json.loads(row[5])

            top5 = data.get(
                "top5",
                []
            )

            for item in top5:

                if item.get("key") == key:

                    rank = item.get(
                        "rank",
                        0
                    )

                    if 1 <= rank <= 5:

                        counts[rank - 1] += 1

        ws5.append([
            all_labels.get(
                key,
                key
            ),
            counts[0],
            counts[1],
            counts[2],
            counts[3],
            counts[4],
            sum(counts)
        ])

    # --------------------------------------------------------
    # 06 SD
    # --------------------------------------------------------

    ws6 = wb.create_sheet(
        "06_SD"
    )

    ws6.append([
        "Respondent ID",
        "Jins",
        "SD Score",
        "SD Level"
    ])

    for row in rows:

        data = json.loads(row[5])

        score = calculate_sd_score(
            data.get(
                "sd_answers",
                {}
            )
        )

        ws6.append([
            row[0],
            row[3],
            score,
            sd_level(score)
        ])

    # --------------------------------------------------------
    # 07 JINSLARARO TAQQOSLASH
    # --------------------------------------------------------

    ws7 = wb.create_sheet(
        "07_Jinslararo_Taqqoslash"
    )

    ws7.append([
        "Savol",
        "Erkaklar o‘rtachasi",
        "Ayollar o‘rtachasi",
        "Farq"
    ])

    for index in range(25):

        key = MALE_RESPONDENT_CRITERIA[index][0]

        male_values = []
        female_values = []

        for row in rows:

            data = json.loads(row[5])

            value = data.get(
                "test_answers",
                {}
            ).get(
                key,
                None
            )

            if value is None:
                continue

            if row[3] == "Erkak":
                male_values.append(value)
            else:
                female_values.append(value)

        male_mean = (
            sum(male_values) /
            len(male_values)
            if male_values
            else 0
        )

        female_mean = (
            sum(female_values) /
            len(female_values)
            if female_values
            else 0
        )

        label = FEMALE_RESPONDENT_CRITERIA[index][1]

        ws7.append([
            label,
            round(male_mean, 3),
            round(female_mean, 3),
            round(
                male_mean - female_mean,
                3
            )
        ])

    # --------------------------------------------------------
    # 08 CODEBOOK
    # --------------------------------------------------------

    ws8 = wb.create_sheet(
        "08_Codebook"
    )

    ws8.append([
        "O‘zgaruvchi",
        "Izoh",
        "Kodlash"
    ])

    codebook = [

        ("gender", "Jins", "1=Erkak; 2=Ayol"),

        ("age_group", "Yosh guruhi",
         "1=17–19; 2=20–22; 3=23–25; 4=26+"),

        ("course", "Kurs",
         "1=1-kurs; 2=2-kurs; 3=3-kurs; 4=4-kurs; 5=Magistratura/Boshqa"),

        ("residence", "Yashash joyi",
         "1=Shahar; 2=Tuman/Qishloq"),

        ("marital_status", "Oilaviy holat",
         "1=Turmush qurmagan; 2=Turmush qurgan"),

        ("Likert", "Test savollari",
         "0=Umuman muhim emas; 1=Unchalik muhim emas; 2=Muhim; 3=Juda muhim"),

        ("SD", "Social desirability",
         "0–6 ball; Past/O‘rtacha/Yuqori")
    ]

    for item in codebook:

        ws8.append(item)

    # --------------------------------------------------------
    # 09 XULOSA
    # --------------------------------------------------------

    ws9 = wb.create_sheet(
        "09_Xulosa"
    )

    total = len(rows)

    male = sum(
        1
        for r in rows
        if r[3] == "Erkak"
    )

    female = sum(
        1
        for r in rows
        if r[3] == "Ayol"
    )

    ws9.append([
        "Ko‘rsatkich",
        "Qiymat"
    ])

    ws9.append([
        "Jami respondentlar",
        total
    ])

    ws9.append([
        "Erkak respondentlar",
        male
    ])

    ws9.append([
        "Ayol respondentlar",
        female
    ])

    ws9.append([
        "Asosiy test savollari",
        25
    ])

    ws9.append([
        "TOP-5 mezonlari",
        5
    ])

    ws9.append([
        "SD savollari",
        6
    ])

    # --------------------------------------------------------
    # FORMAT
    # --------------------------------------------------------

    for sheet in wb.worksheets:

        for row in sheet.iter_rows():

            for cell in row:

                cell.border = border

                cell.alignment = Alignment(
                    vertical="top",
                    wrap_text=True
                )

        for cell in sheet[1]:

            cell.fill = header_fill
            cell.font = header_font

        for column in sheet.columns:

            max_length = 0

            column_letter = get_column_letter(
                column[0].column
            )

            for cell in column:

                try:

                    length = len(
                        str(cell.value)
                    )

                    if length > max_length:
                        max_length = length

                except:
                    pass

            sheet.column_dimensions[
                column_letter
            ].width = min(
                max(max_length + 2, 12),
                45
            )

    wb.save(
        EXPORT_PATH
    )

    return EXPORT_PATH


# ============================================================
# 17. YORDAMCHI FUNKSIYALAR
# ============================================================

def calculate_sd_score(sd_answers):

    score = 0

    for key, question, expected in SD_QUESTIONS:

        answer = sd_answers.get(
            key
        )

        if answer is None:
            continue

        if bool(answer) == expected:

            score += 1

    return score


def sd_level(score):

    if score <= 1:
        return "Past"

    if score <= 4:
        return "O‘rtacha"

    return "Yuqori"


def encode_sd_level(level):

    return {
        "Past": 1,
        "O‘rtacha": 2,
        "Yuqori": 3
    }.get(
        level,
        0
    )


def encode_age(value):

    return {
        "17–19": 1,
        "20–22": 2,
        "23–25": 3,
        "26 va undan katta": 4
    }.get(
        value,
        0
    )


def encode_course(value):

    return {
        "1-kurs": 1,
        "2-kurs": 2,
        "3-kurs": 3,
        "4-kurs": 4,
        "Magistratura / Boshqa": 5
    }.get(
        value,
        0
    )


def encode_residence(value):

    return {
        "Shahar": 1,
        "Tuman/Qishloq": 2
    }.get(
        value,
        0
    )


def encode_marital(value):

    return {
        "Turmush qurmagan": 1,
        "Turmush qurgan": 2
    }.get(
        value,
        0
    )


def encode_children(value):

    return {
        "0": 0,
        "1": 1,
        "2": 2,
        "3": 3,
        "4": 4,
        "5 va undan ko‘p": 5
    }.get(
        value,
        0
    )


def encode_factor(value):

    factors = {
        "Shaxsiy xarakter va odob": 1,
        "Intellekt va ta’lim": 2,
        "Tashqi ko‘rinish": 3,
        "Oilaviy qadriyatlar va tarbiya": 4,
        "Ijtimoiy-iqtisodiy va ro‘zg‘or omillari": 5,
        "Boshqa": 6
    }

    return factors.get(
        value,
        6
    )


# ============================================================
# 18. ISHGA TUSHIRISH
# ============================================================

if __name__ == "__main__":

    asyncio.run(
        main()
    )
