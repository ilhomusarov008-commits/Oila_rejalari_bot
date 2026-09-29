import os
import ast
import sqlite3
import asyncio
from datetime import datetime

from aiohttp import web

from aiogram import Bot, Dispatcher, F
from aiogram.filters import Command, CommandStart
from aiogram.types import Message, CallbackQuery, FSInputFile
from aiogram.utils.keyboard import InlineKeyboardBuilder

from aiogram.fsm.state import State, StatesGroup
from aiogram.fsm.context import FSMContext

from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter


# ============================================================
# 1. BOT SOZLAMALARI
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
# 2. UMUMIY DEMOGRAFIK SAVOLLAR
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
        "ideal_marriage_age",
        "Siz uchun nikoh qurishning maqbul yoshi nechada?",
        None
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
# 3. AYOLLAR UCHUN — IDEAL ERKAK MEZONLARI
# ============================================================

FEMALE_RESPONDENT_CRITERIA = [

    (
        "kind_understanding",
        "Mehribon va tushunuvchan bo‘lishi"
    ),

    (
        "reliable_determined",
        "Ishonchli va qat’iyatli xarakterga ega bo‘lishi"
    ),

    (
        "intelligent_educated",
        "Aql-idrokli va ilmli bo‘lishi"
    ),

    (
        "emotional_stable",
        "Hissiy jihatdan barqaror va sabrli bo‘lishi"
    ),

    (
        "pleasant_manner",
        "Yoqimli fe’l-atvor va yaxshi muomalaga ega bo‘lishi"
    ),

    (
        "honest_responsible",
        "Halol, mas’uliyatli va sadoqatli bo‘lishi"
    ),

    (
        "higher_education",
        "Yuqori ma’lumotga / ta’limga ega bo‘lishi"
    ),

    (
        "self_development_career",
        "O‘z ustida ishlashi va martabaga intilishi"
    ),

    (
        "hardworking",
        "Mehnatsevar va harakatchan bo‘lishi"
    ),

    (
        "financial_provider",
        "Oila va ro‘zg‘orni munosib moliyaviy ta’minlay olishi"
    ),

    (
        "social_status",
        "Ijtimoiy mavqei yoki obro‘si yuqori bo‘lishi"
    ),

    (
        "physical_attractiveness",
        "Jismoniy jihatdan jozibador va kelbatli bo‘lishi"
    ),

    (
        "healthy_lifestyle",
        "Sog‘lom turmush tarziga rioya qilishi"
    ),

    (
        "neatness",
        "Ozoda va saranjom bo‘lishi"
    ),

    (
        "family_leadership",
        "Oilaviy mas’uliyatni va yetakchilikni o‘z zimmasiga olishi"
    ),

    (
        "respect_family_values",
        "Kattalarga va oilaviy qadriyatlarga hurmat ko‘rsatishi"
    )
]


# ============================================================
# 4. ERKAKLAR UCHUN — IDEAL AYOL MEZONLARI
# ============================================================

MALE_RESPONDENT_CRITERIA = [

    (
        "gentle_understanding",
        "Muloyim, shirinso‘z va tushunuvchan bo‘lishi"
    ),

    (
        "faithful",
        "Vafodor va sadoqatli bo‘lishi"
    ),

    (
        "intelligent_educated",
        "Aql-idrokli va ilmli/ma’rifatli bo‘lishi"
    ),

    (
        "emotional_stable",
        "Hissiy jihatdan barqaror va sabr-toqatli bo‘lishi"
    ),

    (
        "pleasant_sincere",
        "Yoqimli fe’l-atvor va samimiy bo‘lishi"
    ),

    (
        "honest_responsible",
        "Halol va mas’uliyatli bo‘lishi"
    ),

    (
        "education_worldview",
        "Yaxshi ta’lim va dunyoqarashga ega bo‘lishi"
    ),

    (
        "cooking_household",
        "Pazandalik va ro‘zg‘or tutish ko‘nikmalariga ega bo‘lishi"
    ),

    (
        "financial_management",
        "Tejamkorlik va oilaviy resurslarni to‘g‘ri boshqara olishi"
    ),

    (
        "family_children_priority",
        "Oila va farzandlar tarbiyasini birinchi o‘ringa qo‘yishi"
    ),

    (
        "beauty_attractiveness",
        "Tashqi jozibadorlik va go‘zallikka ega bo‘lishi"
    ),

    (
        "healthy_lifestyle",
        "Sog‘lom turmush tarziga rioya qilishi"
    ),

    (
        "neatness",
        "Ozoda, sarishta va saranjom bo‘lishi"
    ),

    (
        "respect_elders_traditions",
        "Kattalarga va urf-odatlarga hurmat ko‘rsatishi"
    ),

    (
        "dress_social_etiquette",
        "Kiyinish va jamiyatda muomala odobiga rioya qilishi"
    ),

    (
        "husband_role_respect",
        "Erkakning oiladagi o‘rnini va hurmatini joyiga qo‘yishi"
    )
]


# ============================================================
# 5. LIKERT SHKALASI
# ============================================================

LIKERT_LABELS = [
    ("0 — Umuman muhim emas", 0),
    ("1 — Unchalik muhim emas", 1),
    ("2 — Muhim", 2),
    ("3 — Juda muhim", 3)
]


# ============================================================
# 6. TOP-5
# ============================================================

TOP5_LIMIT = 5


# ============================================================
# 7. SD SHKALASI
# ============================================================

SD_QUESTIONS = [

    (
        "sd_1",
        "Hech qachon birovga nisbatan ich-ichimdan g‘azab yoki nafrat sezmaganman.",
        True
    ),

    (
        "sd_2",
        "Gohida bajara olmaydigan va'dalarni berib qo‘yaman.",
        False
    ),

    (
        "sd_3",
        "Har doim o‘z xatolarimni ochiq tan olaman.",
        True
    ),

    (
        "sd_4",
        "Ba'zan atrofimdagilar haqida g‘iybat qilishim yoki g‘iybatni eshitishim mumkin.",
        False
    ),

    (
        "sd_5",
        "Har qanday vaziyatda ham samimiy va xushfe'l bo‘lishga intilaman.",
        True
    ),

    (
        "sd_6",
        "Ba'zan kayfiyatim yomon bo‘lsa, atrofdagilarga qo‘pollik qilib qo‘yaman.",
        False
    )
]


# ============================================================
# 8. YAKUNIY BOSH OMIL
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
# 9. DATABASE
# ============================================================

def db():

    conn = sqlite3.connect(DB_PATH)

    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS responses (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            telegram_id INTEGER,
            name TEXT,
            gender TEXT,
            created_at TEXT,
            data TEXT
        )
        """
    )

    conn.commit()

    return conn


def save_response(
    telegram_id,
    name,
    gender,
    data
):

    conn = db()

    conn.execute(
        """
        INSERT INTO responses
        (
            telegram_id,
            name,
            gender,
            created_at,
            data
        )
        VALUES (?, ?, ?, ?, ?)
        """,
        (
            telegram_id,
            name,
            gender,
            datetime.now().isoformat(timespec="seconds"),
            repr(data)
        )
    )

    conn.commit()
    conn.close()


def has_completed_response(telegram_id):

    conn = db()

    row = conn.execute(
        """
        SELECT 1
        FROM responses
        WHERE telegram_id = ?
        LIMIT 1
        """,
        (telegram_id,)
    ).fetchone()

    conn.close()

    return row is not None


# ============================================================
# 10. SD HISOBLASH
# ============================================================

def calculate_sd_level(score):

    if score <= 1:
        return "Past"

    if score <= 4:
        return "O‘rtacha"

    return "Yuqori"


# ============================================================
# 11. KLAVIATURALAR
# ============================================================

def option_buttons(items, prefix="ans"):

    kb = InlineKeyboardBuilder()

    for i, item in enumerate(items):

        kb.button(
            text=item,
            callback_data=f"{prefix}:{i}"
        )

    kb.adjust(1)

    return kb.as_markup()


def gender_buttons():

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


def likert_buttons():

    kb = InlineKeyboardBuilder()

    for label, value in LIKERT_LABELS:

        kb.button(
            text=label,
            callback_data=f"likert:{value}"
        )

    kb.adjust(1)

    return kb.as_markup()


def sd_buttons():

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

        kb.button(
            text=prefix + label,
            callback_data=f"top5:{i}"
        )

    kb.button(
        text=f"Tanlangan: {len(selected)}/5",
        callback_data="top5:count"
    )

    if len(selected) == 5:

        kb.button(
            text="✅ Tasdiqlash",
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
        text="📥 Excel — to‘liq tahlil",
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
# 12. FSM
# ============================================================

class Survey(StatesGroup):

    name = State()

    gender = State()

    general = State()

    ideal_marriage_age = State()

    likert = State()

    top5 = State()

    sd_scale = State()

    main_factor = State()

    custom_factor = State()


# ============================================================
# 13. SO‘ROVNOMANI BOSHLASH
# ============================================================

async def start_survey(
    message: Message,
    state: FSMContext
):

    await state.clear()

    await state.set_state(
        Survey.name
    )

    await message.answer(
        "Assalomu alaykum!\n\n"
        "Oila va nikoh mezonlari bo‘yicha ilmiy tadqiqot "
        "so‘rovnomasiga xush kelibsiz.\n\n"
        "Iltimos, ismingizni kiriting:"
    )


# ============================================================
# 14. UMUMIY SAVOLLARNI YUBORISH
# ============================================================

async def send_general_question(
    message: Message,
    state: FSMContext
):

    data = await state.get_data()

    index = data.get(
        "general_index",
        0
    )

    if index >= len(GENERAL_QUESTIONS):

        await send_likert_question(
            message,
            state,
            0
        )

        return

    key, question, options = GENERAL_QUESTIONS[index]

    await state.set_state(
        Survey.general
    )

    if options is None:

        await state.set_state(
            Survey.ideal_marriage_age
        )

        await message.answer(
            f"📌 {index + 3}/{len(GENERAL_QUESTIONS) + 2}\n\n"
            f"{question}\n\n"
            "Javobingizni faqat raqam bilan yozing.\n"
            "Masalan: 25"
        )

        return

    await message.answer(
        f"📌 Savol {index + 3}\n\n"
        f"{question}",
        reply_markup=option_buttons(
            options,
            "general"
        )
    )


# ============================================================
# 15. LIKERT SAVOLLARI
# ============================================================

async def send_likert_question(
    message: Message,
    state: FSMContext,
    index
):

    data = await state.get_data()

    gender = data.get(
        "gender"
    )

    if gender == "Ayol":

        criteria = FEMALE_RESPONDENT_CRITERIA

        target = "ideal erkak"

    else:

        criteria = MALE_RESPONDENT_CRITERIA

        target = "ideal ayol"

    if index >= len(criteria):

        await state.set_state(
            Survey.top5
        )

        await state.update_data(
            top5_selected=[]
        )

        await message.answer(
            "🏆 III-BO‘LIM — TOP-5\n\n"
            f"Yuqoridagi {len(criteria)} ta mezon ichidan "
            "siz uchun eng muhim 5 tasini tanlang.\n\n"
            "Birinchi tanlov — 1-o‘rin\n"
            "Ikkinchi tanlov — 2-o‘rin\n"
            "va hokazo.\n\n"
            "Tanlangan: 0/5",
            reply_markup=top5_keyboard(
                [],
                criteria
            )
        )

        return

    await state.update_data(
        likert_index=index
    )

    await state.set_state(
        Survey.likert
    )

    _, label = criteria[index]

    await message.answer(
        f"📊 II-BO‘LIM — {index + 1}/{len(criteria)}\n\n"
        f"Bo‘lajak {target}ning quyidagi xususiyati "
        "siz uchun qanchalik muhim?\n\n"
        f"👉 {label}",
        reply_markup=likert_buttons()
    )


# ============================================================
# 16. SD SAVOLLAR
# ============================================================

async def send_sd_question(
    message: Message,
    state: FSMContext,
    index
):

    if index >= len(SD_QUESTIONS):

        await state.set_state(
            Survey.main_factor
        )

        await message.answer(
            "🎯 V-BO‘LIM — YAKUNIY SAVOL\n\n"
            "Umuman olganda, juft tanlashda siz uchun "
            "qaysi umumiy yo‘nalish eng muhim?",
            reply_markup=option_buttons(
                MAIN_FACTORS,
                "factor"
            )
        )

        return

    await state.update_data(
        sd_index=index
    )

    await state.set_state(
        Survey.sd_scale
    )

    _, question, _ = SD_QUESTIONS[index]

    await message.answer(
        f"📝 IV-BO‘LIM — {index + 1}/{len(SD_QUESTIONS)}\n\n"
        f"“{question}”",
        reply_markup=sd_buttons()
    )


# ============================================================
# 17. BOTNI ISHGA TUSHIRISH
# ============================================================

async def main():

    db()

    # --------------------------------------------------------
    # Render health check
    # --------------------------------------------------------

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

    # --------------------------------------------------------
    # Bot
    # --------------------------------------------------------

    bot = Bot(TOKEN)

    dp = Dispatcher()

    # ========================================================
    # START
    # ========================================================

    @dp.message(CommandStart())
    async def start(
        message: Message,
        state: FSMContext
    ):

        if (
            message.from_user.id != ADMIN_ID
            and has_completed_response(
                message.from_user.id
            )
        ):

            await state.clear()

            await message.answer(
                "Siz ushbu so‘rovnomada avval "
                "qatnashgansiz.\n\n"
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
    async def admin_panel(
        message: Message
    ):

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
                "Iltimos, ismingizni kiriting."
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
            reply_markup=gender_buttons()
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
            answers={},
            sd_score=0
        )

        await call.answer()

        await call.message.edit_reply_markup(
            reply_markup=None
        )

        await send_general_question(
            call.message,
            state
        )

    # ========================================================
    # UMUMIY SAVOLLAR
    # ========================================================

    @dp.callback_query(
        Survey.general,
        F.data.startswith("general:")
    )
    async def answer_general(
        call: CallbackQuery,
        state: FSMContext
    ):

        data = await state.get_data()

        index = data.get(
            "general_index",
            0
        )

        key, question, options = GENERAL_QUESTIONS[index]

        choice = int(
            call.data.split(":")[1]
        )

        value = options[choice]

        answers = data.get(
            "answers",
            {}
        )

        answers[key] = value

        await state.update_data(
            answers=answers,
            general_index=index + 1
        )

        await call.answer()

        await call.message.edit_reply_markup(
            reply_markup=None
        )

        await send_general_question(
            call.message,
            state
        )

    # ========================================================
    # NIKOH YOSHI
    # ========================================================

    @dp.message(Survey.ideal_marriage_age)
    async def answer_marriage_age(
        message: Message,
        state: FSMContext
    ):

        try:

            age = int(
                (message.text or "").strip()
            )

            if age < 15 or age > 60:
                raise ValueError

        except Exception:

            await message.answer(
                "Iltimos, 15–60 oralig‘ida "
                "raqam kiriting.\n\n"
                "Masalan: 25"
            )

            return

        data = await state.get_data()

        answers = data.get(
            "answers",
            {}
        )

        answers[
            "ideal_marriage_age"
        ] = age

        await state.update_data(
            answers=answers,
            general_index=data.get(
                "general_index",
                0
            ) + 1
        )

        await send_general_question(
            message,
            state
        )

    # ========================================================
    # LIKERT
    # ========================================================

    @dp.callback_query(
        Survey.likert,
        F.data.startswith("likert:")
    )
    async def answer_likert(
        call: CallbackQuery,
        state: FSMContext
    ):

        data = await state.get_data()

        index = data[
            "likert_index"
        ]

        gender = data[
            "gender"
        ]

        if gender == "Ayol":

            criteria = FEMALE_RESPONDENT_CRITERIA

        else:

            criteria = MALE_RESPONDENT_CRITERIA

        key, label = criteria[index]

        value = int(
            call.data.split(":")[1]
        )

        answers = data.get(
            "answers",
            {}
        )

        answers[key] = value

        await state.update_data(
            answers=answers
        )

        await call.answer()

        await call.message.edit_reply_markup(
            reply_markup=None
        )

        await send_likert_question(
            call.message,
            state,
            index + 1
        )

    # ========================================================
    # TOP 5
    # ========================================================

    @dp.callback_query(
        Survey.top5,
        F.data.startswith("top5:")
    )
    async def choose_top5(
        call: CallbackQuery,
        state: FSMContext
    ):

        action = call.data.split(
            ":",
            1
        )[1]

        data = await state.get_data()

        gender = data["gender"]

        if gender == "Ayol":

            criteria = FEMALE_RESPONDENT_CRITERIA

        else:

            criteria = MALE_RESPONDENT_CRITERIA

        selected = list(
            data.get(
                "top5_selected",
                []
            )
        )

        if action == "count":

            await call.answer(
                f"Tanlangan: {len(selected)}/5",
                show_alert=True
            )

            return

        if action == "confirm":

            if len(selected) != 5:

                await call.answer(
                    "Avval 5 ta mezonni tanlang.",
                    show_alert=True
                )

                return

            ordered = [
                criteria[i][1]
                for i in selected
            ]

            top5_text = " | ".join(
                f"{i + 1}-o‘rin: {value}"
                for i, value
                in enumerate(ordered)
            )

            await state.update_data(
                top5=top5_text
            )

            await call.answer(
                "TOP-5 saqlandi."
            )

            await call.message.edit_reply_markup(
                reply_markup=None
            )

            await send_sd_question(
                call.message,
                state,
                0
            )

            return

        index = int(action)

        if index in selected:

            selected.remove(index)

        else:

            if len(selected) >= TOP5_LIMIT:

                await call.answer(
                    "Ko‘pi bilan 5 ta mezon tanlash mumkin.",
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
        Survey.sd_scale,
        F.data.startswith("sd:")
    )
    async def answer_sd(
        call: CallbackQuery,
        state: FSMContext
    ):

        data = await state.get_data()

        index = data[
            "sd_index"
        ]

        answer = bool(
            int(
                call.data.split(":")[1]
            )
        )

        _, _, expected = SD_QUESTIONS[index]

        score = data.get(
            "sd_score",
            0
        )

        if answer == expected:

            score += 1

        await state.update_data(
            sd_score=score
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
    # BOSH OMIL
    # ========================================================

    @dp.callback_query(
        Survey.main_factor,
        F.data.startswith("factor:")
    )
    async def finish(
        call: CallbackQuery,
        state: FSMContext
    ):

        data = await state.get_data()

        factor = MAIN_FACTORS[
            int(
                call.data.split(":")[1]
            )
        ]

        if factor == "Boshqa":

            await state.set_state(
                Survey.custom_factor
            )

            await call.answer()

            await call.message.edit_reply_markup(
                reply_markup=None
            )

            await call.message.answer(
                "Iltimos, siz uchun muhim "
                "bo‘lgan boshqa yo‘nalishni yozing:"
            )

            return

        answers = data.get(
            "answers",
            {}
        )

        answers[
            "top5_ranking"
        ] = data.get(
            "top5",
            ""
        )

        answers[
            "sd_score"
        ] = data.get(
            "sd_score",
            0
        )

        answers[
            "sd_level"
        ] = calculate_sd_level(
            data.get(
                "sd_score",
                0
            )
        )

        answers[
            "main_factor"
        ] = factor

        save_response(
            call.from_user.id,
            data.get("name", ""),
            data.get("gender", ""),
            answers
        )

        await call.answer()

        await call.message.edit_reply_markup(
            reply_markup=None
        )

        await state.clear()

        await call.message.answer(
            "✅ Rahmat!\n\n"
            "Javoblaringiz muvaffaqiyatli "
            "qabul qilindi."
        )

    # ========================================================
    # BOSHQA — ERKIN MATN
    # ========================================================

    @dp.message(
        Survey.custom_factor
    )
    async def finish_custom_factor(
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

        data = await state.get_data()

        answers = data.get(
            "answers",
            {}
        )

        answers[
            "top5_ranking"
        ] = data.get(
            "top5",
            ""
        )

        answers[
            "sd_score"
        ] = data.get(
            "sd_score",
            0
        )

        answers[
            "sd_level"
        ] = calculate_sd_level(
            data.get(
                "sd_score",
                0
            )
        )

        answers[
            "main_factor"
        ] = factor

        save_response(
            message.from_user.id,
            data.get(
                "name",
                ""
            ),
            data.get(
                "gender",
                ""
            ),
            answers
        )

        await state.clear()

        await message.answer(
            "✅ Rahmat!\n\n"
            "Javoblaringiz muvaffaqiyatli "
            "qabul qilindi."
        )

    # ========================================================
    # ADMIN ACTIONS
    # ========================================================

    @dp.callback_query(
        F.data.startswith("admin:")
    )
    async def admin_actions(
        call: CallbackQuery,
        state: FSMContext
    ):

        if call.from_user.id != ADMIN_ID:
            return

        action = call.data.split(
            ":",
            1
        )[1]

        # ----------------------------------------------------
        # STATS
        # ----------------------------------------------------

        if action == "stats":

            conn = db()

            total = conn.execute(
                "SELECT COUNT(*) FROM responses"
            ).fetchone()[0]

            male = conn.execute(
                """
                SELECT COUNT(*)
                FROM responses
                WHERE gender = 'Erkak'
                """
            ).fetchone()[0]

            female = conn.execute(
                """
                SELECT COUNT(*)
                FROM responses
                WHERE gender = 'Ayol'
                """
            ).fetchone()[0]

            today = datetime.now().date().isoformat()

            today_count = conn.execute(
                """
                SELECT COUNT(*)
                FROM responses
                WHERE substr(created_at,1,10)=?
                """,
                (today,)
            ).fetchone()[0]

            conn.close()

            await call.answer()

            await call.message.answer(
                "📊 TADQIQOT STATISTIKASI\n\n"
                f"Jami respondentlar: {total}\n"
                f"👨 Erkaklar: {male}\n"
                f"👩 Ayollar: {female}\n"
                f"📅 Bugun: {today_count}"
            )

        # ----------------------------------------------------
        # EXPORT
        # ----------------------------------------------------

        elif action == "export":

            await call.answer(
                "Excel tayyorlanmoqda..."
            )

            path, total = export_xlsx()

            await call.message.answer_document(
                FSInputFile(path),
                caption=(
                    f"📊 To‘liq statistik baza tayyor.\n\n"
                    f"Jami respondent: {total}\n\n"
                    "Excel ichida erkaklar, ayollar, "
                    "jinslararo taqqoslash, "
                    "deskriptiv statistika, "
                    "korrelyatsiya, regressiya "
                    "va Stata uchun kodlangan ma’lumotlar mavjud."
                )
            )

        # ----------------------------------------------------
        # RESET ADMIN
        # ----------------------------------------------------

        elif action == "reset":

            conn = db()

            conn.execute(
                """
                DELETE FROM responses
                WHERE telegram_id = ?
                """,
                (call.from_user.id,)
            )

            conn.commit()
            conn.close()

            await state.clear()

            await call.answer(
                "Sizning javoblaringiz o‘chirildi.",
                show_alert=True
            )

        # ----------------------------------------------------
        # RESET ALL
        # ----------------------------------------------------

        elif action == "reset_all":

            conn = db()

            conn.execute(
                "DELETE FROM responses"
            )

            conn.commit()
            conn.close()

            await state.clear()

            await call.answer(
                "Barcha ma’lumotlar o‘chirildi!",
                show_alert=True
            )

            await call.message.answer(
                "💥 BAZA TO‘LIQ TOZALANDI."
            )

        # ----------------------------------------------------
        # START
        # ----------------------------------------------------

        elif action == "start":

            await call.answer()

            await start_survey(
                call.message,
                state
            )

    # ========================================================
    # /STATS
    # ========================================================

    @dp.message(
        Command("stats")
    )
    async def stats_command(
        message: Message
    ):

        if message.from_user.id != ADMIN_ID:
            return

        conn = db()

        total = conn.execute(
            "SELECT COUNT(*) FROM responses"
        ).fetchone()[0]

        male = conn.execute(
            """
            SELECT COUNT(*)
            FROM responses
            WHERE gender='Erkak'
            """
        ).fetchone()[0]

        female = conn.execute(
            """
            SELECT COUNT(*)
            FROM responses
            WHERE gender='Ayol'
            """
        ).fetchone()[0]

        conn.close()

        await message.answer(
            "📊 STATISTIKA\n\n"
            f"Jami: {total}\n"
            f"👨 Erkak: {male}\n"
            f"👩 Ayol: {female}"
        )

    # ========================================================
    # /EXPORT
    # ========================================================

    @dp.message(
        Command("export")
    )
    async def export_command(
        message: Message
    ):

        if message.from_user.id != ADMIN_ID:
            return

        path, total = export_xlsx()

        await message.answer_document(
            FSInputFile(path),
            caption=(
                f"📊 Excel tayyor.\n"
                f"Jami respondent: {total}"
            )
        )

    # ========================================================
    # POLLING
    # ========================================================

    await dp.start_polling(
        bot
    )


# ============================================================
# 18. EXCEL EKSPORT
# ============================================================

def export_xlsx():

    conn = db()

    rows = conn.execute(
        """
        SELECT
            id,
            telegram_id,
            name,
            gender,
            created_at,
            data
        FROM responses
        ORDER BY id
        """
    ).fetchall()

    conn.close()

    wb = Workbook()

    # ========================================================
    # STYLE
    # ========================================================

    blue_fill = PatternFill(
        start_color="1F4E78",
        end_color="1F4E78",
        fill_type="solid"
    )

    green_fill = PatternFill(
        start_color="2D6A4F",
        end_color="2D6A4F",
        fill_type="solid"
    )

    orange_fill = PatternFill(
        start_color="C65911",
        end_color="C65911",
        fill_type="solid"
    )

    white_font = Font(
        bold=True,
        color="FFFFFF"
    )

    bold_font = Font(
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

    # ========================================================
    # MA’LUMOTLARNI TAYYORLASH
    # ========================================================

    all_records = []

    for row in rows:

        rid, telegram_id, name, gender, created_at, data_text = row

        try:

            data = ast.literal_eval(
                data_text
            )

        except Exception:

            data = {}

        record = {
            "id": rid,
            "telegram_id": telegram_id,
            "name": name,
            "gender": gender,
            "created_at": created_at
        }

        record.update(data)

        all_records.append(
            record
        )

    # ========================================================
    # 01 — BARCHA MA’LUMOTLAR
    # ========================================================

    ws = wb.active

    ws.title = "01_Barcha_Malumotlar"

    headers = [

        "ID",
        "Respondent_Kodi",
        "Ism",
        "Jins",
        "Sana",

        "Yosh_Guruhi",
        "Kurs",
        "Yashash_Joyi",
        "Oilaviy_Holati",
        "Maqbul_Nikoh_Yoshi",
        "Istalgan_Farzandlar"

    ]

    # Ayollar / erkaklar mezonlarini birgalikda saqlaymiz

    all_criteria_keys = []

    for key, label in FEMALE_RESPONDENT_CRITERIA:

        all_criteria_keys.append(
            key
        )

    for key, label in MALE_RESPONDENT_CRITERIA:

        if key not in all_criteria_keys:

            all_criteria_keys.append(
                key
            )

    for key in all_criteria_keys:

        headers.append(
            key
        )

    headers += [

        "TOP5",
        "SD_Score",
        "SD_Level",
        "Main_Factor"

    ]

    ws.append(headers)

    for cell in ws[1]:

        cell.font = white_font
        cell.fill = blue_fill
        cell.alignment = Alignment(
            wrap_text=True,
            horizontal="center",
            vertical="center"
        )

    for record in all_records:

        row = [

            record["id"],

            f"RESP-{record['id']:04d}",

            record.get(
                "name",
                ""
            ),

            record.get(
                "gender",
                ""
            ),

            record.get(
                "created_at",
                ""
            ),

            record.get(
                "age_group",
                ""
            ),

            record.get(
                "course",
                ""
            ),

            record.get(
                "residence",
                ""
            ),

            record.get(
                "marital_status",
                ""
            ),

            record.get(
                "ideal_marriage_age",
                ""
            ),

            record.get(
                "desired_children",
                ""
            )

        ]

        for key in all_criteria_keys:

            row.append(
                record.get(
                    key,
                    ""
                )
            )

        row += [

            record.get(
                "top5_ranking",
                ""
            ),

            record.get(
                "sd_score",
                ""
            ),

            record.get(
                "sd_level",
                ""
            ),

            record.get(
                "main_factor",
                ""
            )

        ]

        ws.append(row)

    # ========================================================
    # 02 — ERKAKLAR
    # ========================================================

    create_gender_sheet(
        wb,
        "02_Erkaklar",
        [
            r
            for r in all_records
            if r.get("gender") == "Erkak"
        ],
        MALE_RESPONDENT_CRITERIA,
        blue_fill,
        white_font
    )

    # ========================================================
    # 03 — AYOLLAR
    # ========================================================

    create_gender_sheet(
        wb,
        "03_Ayollar",
        [
            r
            for r in all_records
            if r.get("gender") == "Ayol"
        ],
        FEMALE_RESPONDENT_CRITERIA,
        green_fill,
        white_font
    )

    # ========================================================
    # 04 — JINS TAQQOSLASH
    # ========================================================

    create_gender_comparison(
        wb,
        all_records,
        blue_fill,
        green_fill,
        white_font,
        bold_font
    )

    # ========================================================
    # 05 — DESKRIPTIV
    # ========================================================

    create_descriptive_sheet(
        wb,
        all_records,
        blue_fill,
        white_font,
        bold_font
    )

    # ========================================================
    # 06 — LIKERT
    # ========================================================

    create_likert_analysis(
        wb,
        all_records,
        blue_fill,
        white_font,
        bold_font
    )

    # ========================================================
    # 07 — TOP5
    # ========================================================

    create_top5_analysis(
        wb,
        all_records,
        blue_fill,
        white_font,
        bold_font
    )

    # ========================================================
    # 08 — SD
    # ========================================================

    create_sd_analysis(
        wb,
        all_records,
        blue_fill,
        white_font,
        bold_font
    )

    # ========================================================
    # 09 — KORRELYATSIYA
    # ========================================================

    create_correlation_sheet(
        wb,
        all_records,
        blue_fill,
        white_font,
        bold_font
    )

    # ========================================================
    # 10 — REGRESSIYA
    # ========================================================

    create_regression_sheet(
        wb,
        all_records,
        blue_fill,
        white_font,
        bold_font
    )

    # ========================================================
    # 11 — STATA DATA
    # ========================================================

    create_stata_data(
        wb,
        all_records,
        blue_fill,
        white_font
    )

    # ========================================================
    # 12 — STATA CODEBOOK
    # ========================================================

    create_stata_codebook(
        wb,
        blue_fill,
        white_font,
        bold_font
    )

    # ========================================================
    # UMUMIY FORMAT
    # ========================================================

    for ws in wb.worksheets:

        ws.freeze_panes = "A2"

        for row in ws.iter_rows():

            for cell in row:

                cell.border = border

                cell.alignment = Alignment(
                    vertical="center",
                    wrap_text=True
                )

        for col_idx, column_cells in enumerate(
            ws.columns,
            1
        ):

            max_len = 10

            for cell in list(
                column_cells
            )[:100]:

                value = str(
                    cell.value or ""
                )

                max_len = max(
                    max_len,
                    len(value)
                )

            ws.column_dimensions[
                get_column_letter(col_idx)
            ].width = min(
                max_len + 2,
                45
            )

    wb.save(
        EXPORT_PATH
    )

    return (
        EXPORT_PATH,
        len(all_records)
    )


# ============================================================
# 19. GENDER SHEET
# ============================================================

def create_gender_sheet(
    wb,
    title,
    records,
    criteria,
    header_fill,
    white_font
):

    ws = wb.create_sheet(
        title
    )

    ws["A1"] = title.upper()

    ws["A1"].font = Font(
        bold=True,
        size=14,
        color="FFFFFF"
    )

    ws["A1"].fill = header_fill

    ws["A3"] = "Respondentlar soni"

    ws["B3"] = len(records)

    headers = [
        "ID",
        "Kod",
        "Ism",
        "Jins",
        "Yosh",
        "Kurs",
        "Yashash joyi",
        "Oilaviy holat",
        "Maqbul nikoh yoshi",
        "Farzandlar"
    ]

    headers += [
        label
        for key, label in criteria
    ]

    headers += [
        "TOP5",
        "SD Score",
        "SD Level",
        "Main Factor"
    ]

    ws.append([])

    ws.append(
        headers
    )

    for cell in ws[5]:

        cell.font = white_font
        cell.fill = header_fill

    for record in records:

        row = [

            record.get(
                "id",
                ""
            ),

            f"RESP-{record.get('id', 0):04d}",

            record.get(
                "name",
                ""
            ),

            record.get(
                "gender",
                ""
            ),

            record.get(
                "age_group",
                ""
            ),

            record.get(
                "course",
                ""
            ),

            record.get(
                "residence",
                ""
            ),

            record.get(
                "marital_status",
                ""
            ),

            record.get(
                "ideal_marriage_age",
                ""
            ),

            record.get(
                "desired_children",
                ""
            )

        ]

        for key, label in criteria:

            row.append(
                record.get(
                    key,
                    ""
                )
            )

        row += [

            record.get(
                "top5_ranking",
                ""
            ),

            record.get(
                "sd_score",
                ""
            ),

            record.get(
                "sd_level",
                ""
            ),

            record.get(
                "main_factor",
                ""
            )

        ]

        ws.append(
            row
        )


# ============================================================
# 20. JINS TAQQOSLASH
# ============================================================

def create_gender_comparison(
    wb,
    records,
    blue_fill,
    green_fill,
    white_font,
    bold_font
):

    ws = wb.create_sheet(
        "04_Jins_Taqqoslash"
    )

    ws["A1"] = (
        "ERKAK VA AYOL RESPONDENTLARINING "
        "JUFT TANLASH MEZONLARI TAQQOSLANISHI"
    )

    ws["A1"].font = Font(
        bold=True,
        size=14
    )

    ws["A3"] = "Mezon"
    ws["B3"] = "Erkaklar o‘rtachasi"
    ws["C3"] = "Ayollar o‘rtachasi"
    ws["D3"] = "Farq (Ayol-Erkak)"

    for cell in ws[3]:

        cell.font = white_font
        cell.fill = blue_fill

    max_len = max(
        len(MALE_RESPONDENT_CRITERIA),
        len(FEMALE_RESPONDENT_CRITERIA)
    )

    for i in range(max_len):

        row = i + 4

        male = (
            MALE_RESPONDENT_CRITERIA[i]
            if i < len(MALE_RESPONDENT_CRITERIA)
            else None
        )

        female = (
            FEMALE_RESPONDENT_CRITERIA[i]
            if i < len(FEMALE_RESPONDENT_CRITERIA)
            else None
        )

        label = ""

        if male:
            label = male[1]

        if female:
            if label:
                label += " / " + female[1]
            else:
                label = female[1]

        ws.cell(
            row=row,
            column=1,
            value=label
        )

        male_key = male[0] if male else ""

        female_key = (
            female[0]
            if female
            else ""
        )

        male_values = [

            r.get(
                male_key
            )
            for r in records
            if r.get("gender") == "Erkak"
            and isinstance(
                r.get(male_key),
                (int, float)
            )
        ]

        female_values = [

            r.get(
                female_key
            )
            for r in records
            if r.get("gender") == "Ayol"
            and isinstance(
                r.get(female_key),
                (int, float)
            )
        ]

        male_range = f"B{row}"
        female_range = f"C{row}"

        if male_values:

            ws.cell(
                row=row,
                column=2,
                value=sum(male_values) / len(male_values)
            )

        else:

            ws.cell(
                row=row,
                column=2,
                value=0
            )

        if female_values:

            ws.cell(
                row=row,
                column=3,
                value=sum(female_values) / len(female_values)
            )

        else:

            ws.cell(
                row=row,
                column=3,
                value=0
            )

        ws.cell(
            row=row,
            column=4,
            value=f"=C{row}-B{row}"
        )

        ws.cell(
            row=row,
            column=2
        ).number_format = "0.00"

        ws.cell(
            row=row,
            column=3
        ).number_format = "0.00"

        ws.cell(
            row=row,
            column=4
        ).number_format = "0.00"


# ============================================================
# 21. DESKRIPTIV
# ============================================================

def create_descriptive_sheet(
    wb,
    records,
    header_fill,
    white_font,
    bold_font
):

    ws = wb.create_sheet(
        "05_Deskriptiv"
    )

    ws["A1"] = (
        "DESKRIPTIV STATISTIKA"
    )

    ws["A1"].font = Font(
        bold=True,
        size=14
    )

    headers = [
        "Ko‘rsatkich",
        "N",
        "Mean",
        "Median",
        "Minimum",
        "Maximum",
        "Std.Dev"
    ]

    for col, value in enumerate(
        headers,
        1
    ):

        cell = ws.cell(
            3,
            col,
            value
        )

        cell.font = white_font
        cell.fill = header_fill

    variables = [

        (
            "Maqbul nikoh yoshi",
            "ideal_marriage_age"
        ),

        (
            "Istalgan farzandlar",
            "desired_children"
        ),

        (
            "SD Score",
            "sd_score"
        )

    ]

    row = 4

    for label, key in variables:

        values = []

        for r in records:

            value = r.get(
                key
            )

            if isinstance(
                value,
                (int, float)
            ):

                values.append(
                    float(value)
                )

        ws.cell(
            row,
            1,
            label
        )

        ws.cell(
            row,
            2,
            len(values)
        )

        if values:

            mean = sum(values) / len(values)

            sorted_values = sorted(
                values
            )

            n = len(
                sorted_values
            )

            if n % 2 == 0:

                median = (
                    sorted_values[n // 2 - 1]
                    +
                    sorted_values[n // 2]
                ) / 2

            else:

                median = sorted_values[
                    n // 2
                ]

            minimum = min(
                values
            )

            maximum = max(
                values
            )

            if len(values) > 1:

                variance = sum(
                    (
                        x - mean
                    ) ** 2
                    for x in values
                ) / (
                    len(values) - 1
                )

                std = variance ** 0.5

            else:

                std = 0

            ws.cell(
                row,
                3,
                mean
            )

            ws.cell(
                row,
                4,
                median
            )

            ws.cell(
                row,
                5,
                minimum
            )

            ws.cell(
                row,
                6,
                maximum
            )

            ws.cell(
                row,
                7,
                std
            )

        row += 1


# ============================================================
# 22. LIKERT TAHLIL
# ============================================================

def create_likert_analysis(
    wb,
    records,
    header_fill,
    white_font,
    bold_font
):

    ws = wb.create_sheet(
        "06_Likert_Tahlil"
    )

    ws["A1"] = (
        "JUFT TANLASH MEZONLARI — LIKERT TAHLILI"
    )

    ws["A1"].font = Font(
        bold=True,
        size=14
    )

    headers = [
        "Jins",
        "Mezon",
        "N",
        "Mean",
        "Std.Dev",
        "0 Soni",
        "1 Soni",
        "2 Soni",
        "3 Soni",
        "3 ulushi (%)"
    ]

    for col, value in enumerate(
        headers,
        1
    ):

        cell = ws.cell(
            3,
            col,
            value
        )

        cell.font = white_font
        cell.fill = header_fill

    row = 4

    for gender, criteria in [
        (
            "Erkak",
            MALE_RESPONDENT_CRITERIA
        ),
        (
            "Ayol",
            FEMALE_RESPONDENT_CRITERIA
        )
    ]:

        gender_records = [
            r
            for r in records
            if r.get("gender") == gender
        ]

        for key, label in criteria:

            values = [

                r.get(key)

                for r in gender_records

                if isinstance(
                    r.get(key),
                    (int, float)
                )

            ]

            ws.cell(
                row,
                1,
                gender
            )

            ws.cell(
                row,
                2,
                label
            )

            ws.cell(
                row,
                3,
                len(values)
            )

            if values:

                mean = sum(values) / len(values)

                if len(values) > 1:

                    variance = sum(
                        (
                            x - mean
                        ) ** 2
                        for x in values
                    ) / (
                        len(values) - 1
                    )

                    std = variance ** 0.5

                else:

                    std = 0

                ws.cell(
                    row,
                    4,
                    mean
                )

                ws.cell(
                    row,
                    5,
                    std
                )

                for score in range(4):

                    ws.cell(
                        row,
                        6 + score,
                        values.count(score)
                    )

                ws.cell(
                    row,
                    10,
                    values.count(3) / len(values)
                )

                ws.cell(
                    row,
                    10
                ).number_format = "0.0%"

            row += 1


# ============================================================
# 23. TOP-5 TAHLIL
# ============================================================

def create_top5_analysis(
    wb,
    records,
    header_fill,
    white_font,
    bold_font
):

    ws = wb.create_sheet(
        "07_TOP5_Tahlil"
    )

    ws["A1"] = (
        "TOP-5 MEZONLAR TAHLILI"
    )

    ws["A1"].font = Font(
        bold=True,
        size=14
    )

    headers = [
        "Jins",
        "Mezon",
        "1-o‘rin",
        "2-o‘rin",
        "3-o‘rin",
        "4-o‘rin",
        "5-o‘rin",
        "Jami TOP-5"
    ]

    for col, value in enumerate(
        headers,
        1
    ):

        cell = ws.cell(
            3,
            col,
            value
        )

        cell.font = white_font
        cell.fill = header_fill

    row = 4

    for gender, criteria in [
        (
            "Erkak",
            MALE_RESPONDENT_CRITERIA
        ),
        (
            "Ayol",
            FEMALE_RESPONDENT_CRITERIA
        )
    ]:

        gender_records = [
            r
            for r in records
            if r.get("gender") == gender
        ]

        for key, label in criteria:

            counts = [0, 0, 0, 0, 0]

            for r in gender_records:

                text = r.get(
                    "top5_ranking",
                    ""
                )

                if not text:
                    continue

                parts = text.split(
                    " | "
                )

                for rank, part in enumerate(
                    parts[:5]
                ):

                    if label in part:

                        counts[rank] += 1

            ws.cell(
                row,
                1,
                gender
            )

            ws.cell(
                row,
                2,
                label
            )

            for i, value in enumerate(
                counts,
                3
            ):

                ws.cell(
                    row,
                    i,
                    value
                )

            ws.cell(
                row,
                8,
                sum(counts)
            )

            row += 1


# ============================================================
# 24. SD TAHLILI
# ============================================================

def create_sd_analysis(
    wb,
    records,
    header_fill,
    white_font,
    bold_font
):

    ws = wb.create_sheet(
        "08_SD_Tahlil"
    )

    ws["A1"] = (
        "IJTIMOIY MAQBULLIK — SD TAHLILI"
    )

    ws["A1"].font = Font(
        bold=True,
        size=14
    )

    headers = [
        "Jins",
        "SD Daraja",
        "Soni",
        "Ulushi (%)"
    ]

    for col, value in enumerate(
        headers,
        1
    ):

        cell = ws.cell(
            3,
            col,
            value
        )

        cell.font = white_font
        cell.fill = header_fill

    row = 4

    for gender in [
        "Erkak",
        "Ayol"
    ]:

        gender_records = [
            r
            for r in records
            if r.get("gender") == gender
        ]

        total = len(
            gender_records
        )

        for level in [
            "Past",
            "O‘rtacha",
            "Yuqori"
        ]:

            count = sum(
                1
                for r in gender_records
                if r.get("sd_level") == level
            )

            ws.cell(
                row,
                1,
                gender
            )

            ws.cell(
                row,
                2,
                level
            )

            ws.cell(
                row,
                3,
                count
            )

            ws.cell(
                row,
                4,
                count / total
                if total
                else 0
            )

            ws.cell(
                row,
                4
            ).number_format = "0.0%"

            row += 1


# ============================================================
# 25. KORRELYATSIYA
# ============================================================

def create_correlation_sheet(
    wb,
    records,
    header_fill,
    white_font,
    bold_font
):

    ws = wb.create_sheet(
        "09_Korrelyatsiya"
    )

    ws["A1"] = (
        "KORRELYATSIYA TAHLILI"
    )

    ws["A1"].font = Font(
        bold=True,
        size=14
    )

    ws["A3"] = (
        "Excel formulasi orqali Pearson korrelyatsiyasi"
    )

    ws["A5"] = "X"
    ws["B5"] = "Y"
    ws["C5"] = "Pearson r"

    for cell in ws[5]:

        cell.font = white_font
        cell.fill = header_fill

    # Demografik o‘zgaruvchi va SD

    ws["A6"] = "Maqbul nikoh yoshi"
    ws["B6"] = "SD Score"

    # Formula uchun raw sheetdagi ustunlar
    # J = ideal marriage age
    # SD Score ustuni keyin aniqlanadi.

    ws["C6"] = (
        "=IFERROR(CORREL("
        "'01_Barcha_Malumotlar'!J2:J10000,"
        "'01_Barcha_Malumotlar'!AM2:AM10000"
        "),0)"
    )

    ws["A7"] = "Maqbul nikoh yoshi"
    ws["B7"] = "Istalgan farzandlar"

    ws["C7"] = (
        "=IFERROR(CORREL("
        "'01_Barcha_Malumotlar'!J2:J10000,"
        "'01_Barcha_Malumotlar'!K2:K10000"
        "),0)"
    )

    ws["A8"] = "SD Score"
    ws["B8"] = "Maqbul nikoh yoshi"

    ws["C8"] = (
        "=IFERROR(CORREL("
        "'01_Barcha_Malumotlar'!AM2:AM10000,"
        "'01_Barcha_Malumotlar'!J2:J10000"
        "),0)"
    )


# ============================================================
# 26. REGRESSIYA
# ============================================================

def create_regression_sheet(
    wb,
    records,
    header_fill,
    white_font,
    bold_font
):

    ws = wb.create_sheet(
        "10_Regressiya"
    )

    ws["A1"] = (
        "EKONOMETRIK REGRESSIYA — OLS"
    )

    ws["A1"].font = Font(
        bold=True,
        size=14
    )

    ws["A3"] = (
        "Model: Ideal nikoh yoshi = "
        "β0 + β1(SD Score) + β2(Istalgan farzandlar) + ε"
    )

    ws["A5"] = "Ko‘rsatkich"
    ws["B5"] = "Natija"

    for cell in ws[5]:

        cell.font = white_font
        cell.fill = header_fill

    ws["A6"] = "Kuzatuvlar soni"

    ws["B6"] = (
        "=COUNT('01_Barcha_Malumotlar'!J2:J10000)"
    )

    ws["A7"] = "R²"

    # Excel LINEST ko‘p o‘lchovli regressiya
    ws["B7"] = (
        "=IFERROR("
        "INDEX(LINEST("
        "'01_Barcha_Malumotlar'!J2:J10000,"
        "'01_Barcha_Malumotlar'!AM2:AM10000,"
        "TRUE,TRUE),3,1),0)"
    )

    ws["A9"] = "Izoh"

    ws["B9"] = (
        "To‘liq ekonometrik baholashni "
        "Stata/R/Python orqali qayta tekshirish tavsiya etiladi."
    )

    ws["A11"] = "Stata modeli"

    ws["B11"] = (
        "reg ideal_marriage_age sd_score desired_children"
    )


# ============================================================
# 27. STATA UCHUN KODLANGAN MA’LUMOTLAR
# ============================================================

def create_stata_data(
    wb,
    records,
    header_fill,
    white_font
):

    ws = wb.create_sheet(
        "11_Stata_Data"
    )

    headers = [

        "id",
        "gender",
        "gender_code",
        "age_group",
        "age_code",
        "course",
        "course_code",
        "residence",
        "residence_code",
        "marital_status",
        "marital_code",
        "ideal_marriage_age",
        "desired_children",
        "sd_score",
        "sd_level",
        "main_factor"

    ]

    # 16 universal-ish criterion columns

    all_keys = []

    for key, label in MALE_RESPONDENT_CRITERIA:

        all_keys.append(key)

    for key, label in FEMALE_RESPONDENT_CRITERIA:

        if key not in all_keys:

            all_keys.append(key)

    for i, key in enumerate(
        all_keys,
        1
    ):

        headers.append(
            f"q{i}"
        )

    headers += [
        "top5"
    ]

    ws.append(
        headers
    )

    for cell in ws[1]:

        cell.font = white_font
        cell.fill = header_fill

    gender_map = {
        "Erkak": 1,
        "Ayol": 2
    }

    age_map = {
        "17–19": 1,
        "20–22": 2,
        "23–25": 3,
        "26 va undan katta": 4
    }

    course_map = {
        "1-kurs": 1,
        "2-kurs": 2,
        "3-kurs": 3,
        "4-kurs": 4,
        "Magistratura / Boshqa": 5
    }

    residence_map = {
        "Shahar": 1,
        "Tuman/Qishloq": 2
    }

    marital_map = {
        "Turmush qurmagan": 0,
        "Turmush qurgan": 1
    }

    for r in records:

        row = [

            r.get(
                "id",
                ""
            ),

            r.get(
                "gender",
                ""
            ),

            gender_map.get(
                r.get("gender"),
                ""
            ),

            r.get(
                "age_group",
                ""
            ),

            age_map.get(
                r.get("age_group"),
                ""
            ),

            r.get(
                "course",
                ""
            ),

            course_map.get(
                r.get("course"),
                ""
            ),

            r.get(
                "residence",
                ""
            ),

            residence_map.get(
                r.get("residence"),
                ""
            ),

            r.get(
                "marital_status",
                ""
            ),

            marital_map.get(
                r.get("marital_status"),
                ""
            ),

            r.get(
                "ideal_marriage_age",
                ""
            ),

            r.get(
                "desired_children",
                ""
            ),

            r.get(
                "sd_score",
                ""
            ),

            r.get(
                "sd_level",
                ""
            ),

            r.get(
                "main_factor",
                ""
            )

        ]

        for key in all_keys:

            row.append(
                r.get(
                    key,
                    ""
                )
            )

        row.append(
            r.get(
                "top5_ranking",
                ""
            )
        )

        ws.append(
            row
        )


# ============================================================
# 28. STATA CODEBOOK
# ============================================================

def create_stata_codebook(
    wb,
    header_fill,
    white_font,
    bold_font
):

    ws = wb.create_sheet(
        "12_Stata_Codebook"
    )

    headers = [
        "Variable",
        "Mazmuni",
        "Kodlash"
    ]

    for col, value in enumerate(
        headers,
        1
    ):

        cell = ws.cell(
            1,
            col,
            value
        )

        cell.font = white_font
        cell.fill = header_fill

    variables = [

        (
            "gender_code",
            "Jins",
            "1=Erkak; 2=Ayol"
        ),

        (
            "age_code",
            "Yosh guruhi",
            "1=17–19; 2=20–22; 3=23–25; 4=26+"
        ),

        (
            "course_code",
            "Kurs",
            "1=1-kurs; 2=2-kurs; 3=3-kurs; 4=4-kurs; 5=Magistr/Boshqa"
        ),

        (
            "residence_code",
            "Yashash joyi",
            "1=Shahar; 2=Tuman/Qishloq"
        ),

        (
            "marital_code",
            "Oilaviy holat",
            "0=Turmush qurmagan; 1=Turmush qurgan"
        ),

        (
            "ideal_marriage_age",
            "Maqbul nikoh yoshi",
            "15–60"
        ),

        (
            "desired_children",
            "Istalgan farzandlar",
            "0–5+"
        ),

        (
            "q1-q16",
            "Juft tanlash mezonlari",
            "0=Umuman muhim emas; 1=Unchalik muhim emas; 2=Muhim; 3=Juda muhim"
        ),

        (
            "sd_score",
            "Ijtimoiy maqbullik skori",
            "0–6"
        ),

        (
            "sd_level",
            "SD darajasi",
            "Past / O‘rtacha / Yuqori"
        )

    ]

    row = 2

    for variable, meaning, coding in variables:

        ws.cell(
            row,
            1,
            variable
        )

        ws.cell(
            row,
            2,
            meaning
        )

        ws.cell(
            row,
            3,
            coding
        )

        row += 1


# ============================================================
# 29. START
# ============================================================

if __name__ == "__main__":

    asyncio.run(
        main()
    )
