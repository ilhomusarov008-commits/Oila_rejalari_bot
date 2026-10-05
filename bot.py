import os
import sqlite3
import ast
import re
import math
from datetime import datetime
from aiohttp import web
from aiogram import Bot, Dispatcher, F
from aiogram.filters import Command, CommandStart
from aiogram.types import Message, CallbackQuery, FSInputFile
from aiogram.utils.keyboard import InlineKeyboardBuilder
from aiogram.fsm.state import State, StatesGroup
from aiogram.fsm.context import FSMContext
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# 1. BOT SOZLAMALARI
TOKEN = os.getenv("BOT_TOKEN")
ADMIN_ID = int(os.getenv("ADMIN_ID", "8671467518"))
DB_PATH = os.getenv("DB_PATH", "oila_va_nikoh_tadqiqoti.db")
EXPORT_PATH = os.getenv("EXPORT_PATH", "Oila_va_nikoh_statistik_tahlil.xlsx")
if not TOKEN:
    raise RuntimeError("BOT_TOKEN environment variable o'rnatilmagan!")

# 2. DEMOGRAFIK SAVOLLAR
GENERAL_QUESTIONS = [("age_group", "Yoshingiz qaysi oraliqda?", ["17–19", "20–22", "23–25",
    "26 va undan katta"]),
    ("course", "Kursingiz?", ["1-kurs", "2-kurs", "3-kurs", "4-kurs", "Magistratura / Boshqa"]),
    ("residence", "Doimiy yashash joyingiz?", ["Shahar", "Tuman/Qishloq"]),
    ("marital_status", "Hozirgi oilaviy holatingiz?", ["Turmush qurmagan", "Turmush qurgan"]),
    ("desired_children", "Kelajakda nechta farzandli bo‘lishni xohlaysiz?", ["0", "1", "2", "3",
    "4", "5 va undan ko‘p"]),
    ]

# 3. 25 TA ASOSIY TEST MEZONI
CRITERIA = [("kind_understanding", "Mehribon, e’tiborli va tushunuvchan bo‘lishi"),
    ("faithfulness", "Vafodor, sadoqatli va ishonchli bo‘lishi"),
    ("honesty_responsibility", "Halol, va’dasida turadigan va mas’uliyatli bo‘lishi"),
    ("emotional_stability",
    "Hissiy jihatdan barqaror, vazmin va qiyin vaziyatlarda o‘zini boshqara olishi"),
    ("pleasant_disposition", "Xushmuomala, samimiy va yoqimli fe’l-atvorga ega bo‘lishi"),
    ("intelligence", "Aql-idrokli, mustaqil fikrlaydigan va muammolarga oqilona yondasha olishi"),
    ("education_worldview", "Yaxshi ta’limga ega, bilimli va keng dunyoqarashli bo‘lishi"),
    ("self_development",
    "O‘z ustida ishlashi, bilim va ko‘nikmalarini rivojlantirishga intilishi"),
    ("ambition_industriousness", "Maqsadli, tashabbuskor va mehnatsevar bo‘lishi"),
    ("communication_conflict",
    "Ochiq muloqot qilishi, fikrni tinglashi va kelishmovchiliklarni tinch yo‘l bilan hal qila olishi"),
    ("health_lifestyle", "Sog‘lig‘iga e’tibor berishi va sog‘lom turmush tarziga amal qilishi"),
    ("neatness", "Ozoda, saranjom va o‘ziga e’tiborli bo‘lishi"),
    ("physical_attractiveness", "Tashqi ko‘rinishi va jismoniy jozibadorligi"),
    ("age_compatibility", "Yoshi respondentning yoshiga mos bo‘lishi"),
    ("family_values", "Oilaviy qadriyatlarni hurmat qilishi va oilani muhim deb bilishi"),
    ("respect_elders_relatives",
    "Ota-ona, katta yoshdagilar va qarindoshlarga hurmat bilan munosabatda bo‘lishi"),
    ("marriage_orientation", "Nikoh va oilaviy hayotga jiddiy tayyorgarlik ko‘rgan bo‘lishi"),
    ("family_responsibility",
    "Oilaviy qarorlar va majburiyatlarda mas’uliyatli ishtirok etishga tayyorligi"),
    ("parenting_responsibility",
    "Farzand tarbiyasida faol ishtirok etishi va farzandlarga mas’uliyat bilan munosabatda bo‘lishi"),
    ("household_participation",
    "Uy-ro‘zg‘or ishlarida ishtirok etishga va oilaviy vazifalarni bo‘lishishga tayyorligi"),
    ("financial_management",
    "Oilaviy daromad va xarajatlarni rejalashtira olishi hamda mablag‘ni oqilona boshqarishi"),
    ("income_stability",
    "Barqaror daromadga ega bo‘lishi yoki oilaviy daromadga munosib hissa qo‘sha olishi"),
    ("work_family_balance",
    "Kasbiy faoliyatda rivojlanishga intilishi va oila bilan ish o‘rtasida muvozanatni saqlay olishi"),
    ("social_reputation", "Jamiyatda o‘zini tutishi, obro‘si va atrofdagilar bilan munosabati"),
    ("partner_support",
    "Respondentning ta’limi, ishi, kasbiy rivojlanishi va shaxsiy maqsadlarini qo‘llab-quvvatlashi"),
    ]

# 4. LIKERT SHKALASI
LIKERT_LABELS = [("0 — Umuman muhim emas", 0),
    ("1 — Unchalik muhim emas", 1),
    ("2 — Muhim", 2),
    ("3 — Juda muhim", 3),
    ]

# 6. YAKUNIY OMILLAR
MAIN_FACTORS = ["Shaxsiy xarakter va odob",
    "Intellekt va ta’lim",
    "Tashqi ko‘rinish",
    "Oilaviy qadriyatlar va tarbiya",
    "Ijtimoiy-iqtisodiy va ro‘zg‘or omillari",
    "Boshqa"]

# 6b. KONSTRUKTLAR (25 mezonning 5 guruhi — Excel tahlili uchun)
CONSTRUCTS = [
    (MAIN_FACTORS[0], "idx_xarakter", ["kind_understanding", "faithfulness",
        "honesty_responsibility", "emotional_stability", "pleasant_disposition",
        "communication_conflict"]),
    (MAIN_FACTORS[1], "idx_intellekt", ["intelligence", "education_worldview",
        "self_development", "ambition_industriousness"]),
    (MAIN_FACTORS[2], "idx_korinish", ["health_lifestyle", "neatness",
        "physical_attractiveness", "age_compatibility"]),
    (MAIN_FACTORS[3], "idx_oilaviy", ["family_values", "respect_elders_relatives",
        "marriage_orientation", "family_responsibility", "parenting_responsibility"]),
    (MAIN_FACTORS[4], "idx_iqtisodiy", ["household_participation", "financial_management",
        "income_stability", "work_family_balance", "social_reputation", "partner_support"]),
]

# 7. FSM

class Survey(StatesGroup):
    name = State()
    gender = State()
    general = State()
    marriage_age = State()
    likert = State()
    top5 = State()
    main_factor = State()
    custom_factor = State()

# 8. DATABASE

def get_conn():
    conn = sqlite3.connect(DB_PATH)
    conn.execute("""
        CREATE TABLE IF NOT EXISTS responses (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            telegram_id INTEGER NOT NULL,
            name TEXT,
            gender TEXT NOT NULL,
            created_at TEXT NOT NULL,
            data TEXT NOT NULL
        )
        """)
    conn.commit()
    return conn

def save_response(telegram_id, name, gender, data):
    conn = get_conn()
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
        """,
        (telegram_id, name, gender, datetime.now().isoformat(timespec="seconds"), repr(data)))
    conn.commit()
    conn.close()

def has_completed_response(telegram_id):
    conn = get_conn()
    row = conn.execute("""
        SELECT 1
        FROM responses
        WHERE telegram_id = ?
        LIMIT 1
        """,
        (telegram_id,)).fetchone()
    conn.close()
    return row is not None

def delete_user_response(telegram_id):
    conn = get_conn()
    conn.execute("DELETE FROM responses WHERE telegram_id = ?", (telegram_id,))
    conn.commit()
    conn.close()

def delete_all_responses():
    conn = get_conn()
    conn.execute("DELETE FROM responses")
    conn.commit()
    conn.close()

def load_rows():
    conn = get_conn()
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
        """).fetchall(
        )
    conn.close()
    result = []
    for (rid, telegram_id, name, gender, created_at, data_text) in rows:
        try:
            data = ast.literal_eval(data_text)
            if not isinstance(data, dict):
                data = {}
        except Exception:
            data = {}
        result.append({"id": rid, "telegram_id": telegram_id, "name": name or "", "gender": gender or "",
            "created_at": created_at or "", "data": data})
    return result

# 9. KLAVIATURALAR

def option_buttons(items, prefix):
    kb = InlineKeyboardBuilder()
    for i, item in enumerate(items):
        kb.button(text=str(item), callback_data=f"{prefix}:{i}")
    kb.adjust(1)
    return kb.as_markup()

def gender_buttons():
    kb = InlineKeyboardBuilder()
    kb.button(text="Erkak", callback_data="gender:Erkak")
    kb.button(text="Ayol", callback_data="gender:Ayol")
    kb.adjust(2)
    return kb.as_markup()

def likert_buttons():
    kb = InlineKeyboardBuilder()
    for label, value in LIKERT_LABELS:
        kb.button(text=label, callback_data=f"likert:{value}")
    kb.adjust(1)
    return kb.as_markup()

def top5_keyboard(selected):
    kb = InlineKeyboardBuilder()
    for i, (_, label) in enumerate(CRITERIA):
        if i in selected:
            prefix = "☑️ "
        else:
            prefix = "⬜ "
        kb.button(text=prefix + label, callback_data=f"top5:{i}")
    kb.button(text=f"Tanlangan: {len(selected)}/5", callback_data="top5:count")
    if len(selected) == 5:
        kb.button(text="Tasdiqlash", callback_data="top5:confirm")
    kb.adjust(1)
    return kb.as_markup()

def admin_keyboard():
    kb = InlineKeyboardBuilder()
    kb.button(text="Statistika", callback_data="admin:stats")
    kb.button(text="Excel — to‘liq tahlil", callback_data="admin:export")
    kb.button(text="Mening testimni o‘chirish", callback_data="admin:reset")
    kb.button(text="BARCHA MA’LUMOTNI O‘CHIRISH", callback_data="admin:reset_all")
    kb.button(text="So‘rovnomani boshlash", callback_data="admin:start")
    kb.adjust(1)
    return kb.as_markup()

# 10. YORDAMCHI FUNKSIYALAR

def target_label(gender):
    if gender == "Ayol":
        return "erkak"
    return "ayol"

async def start_survey(message, state):
    await state.clear()
    await state.set_state(Survey.name)
    await message.answer("Assalomu alaykum!\n\n" "Oila va nikoh mezonlari bo‘yicha "
        "ilmiy tadqiqot so‘rovnomasiga xush kelibsiz.\n\n" "Ismingizni kiriting yoki "
        "«O‘tkazib yuborish» tugmasini bosing.\n" "Ism-familiya berish majburiy emas.",
        reply_markup=option_buttons(["O‘tkazib yuborish"], "skip_name"))

async def send_general(message, state, index):
    if index >= len(GENERAL_QUESTIONS):
        await state.set_state(Survey.marriage_age)
        await message.answer("Siz uchun nikoh qurishning " "maqbul yoshi nechada?\n\n"
            "Javobingizni faqat raqam bilan yozing.\n" "Masalan: 25")
        return
    key, question, options = GENERAL_QUESTIONS[index]
    await state.update_data(general_index=index)
    await state.set_state(Survey.general)
    await message.answer(f"DEMOGRAFIK SAVOL " f"{index + 1}/{len(GENERAL_QUESTIONS)}\n\n"
        f"{question}", reply_markup=option_buttons(options, "general"))

async def send_likert(message, state, index):
    if index >= len(CRITERIA):
        await state.update_data(top5_selected=[])
        await state.set_state(Survey.top5)
        await message.answer("TOP-5 TANLOV\n\n" "Yuqoridagi 25 ta test mezonidan "
            "siz uchun eng muhim 5 tasini tanlang.\n\n" "Birinchi bosganingiz — 1-o‘rin,\n"
            "ikkinchi bosganingiz — 2-o‘rin,\n" "shu tartibda 5-o‘ringacha.\n\n"
            "Tanlangan: 0/5", reply_markup=top5_keyboard([]))
        return
    data = await state.get_data()
    target = target_label(data.get("gender"))
    key, label = CRITERIA[index]
    await state.update_data(likert_index=index)
    await state.set_state(Survey.likert)
    await message.answer(f"TEST SAVOLI " f"{index + 1}/25\n\n" f"Bo‘lajak {target}ning "
        "quyidagi xususiyati siz uchun " "qanchalik muhim?\n\n" f"{label}", reply_markup=likert_buttons(
        ))

async def send_main_factor(message, state):
    await state.set_state(Survey.main_factor)
    await message.answer("YAKUNIY SAVOL\n\n" "Umuman olganda, juft tanlashda "
        "siz uchun qaysi umumiy yo‘nalish " "eng muhim?", reply_markup=option_buttons(
        MAIN_FACTORS, "factor"))

# 11. EXCEL EKSPORT

def style_sheet(ws):
    thin = Side(style="thin")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    for row in ws.iter_rows():
        for cell in row:
            cell.alignment = Alignment(vertical="top", wrap_text=True)
            cell.border = border
    for cell in ws[1]:
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    for col in ws.columns:
        max_len = 0
        letter = get_column_letter(col[0].column)
        for cell in col:
            value = "" if cell.value is None else str(cell.value)
            max_len = max(max_len, len(value))
        ws.column_dimensions[letter].width = min(max(max_len + 2, 12), 45)
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions


# ---------- statistik yordamchilar (tashqi kutubxonasiz) ----------

def mean(v):
    return sum(v) / len(v) if v else None


def stdev(v):
    n = len(v)
    if n < 2:
        return None
    m = sum(v) / n
    return (sum((x - m) ** 2 for x in v) / (n - 1)) ** 0.5


def median(v):
    if not v:
        return None
    s = sorted(v)
    mid = len(s) // 2
    return s[mid] if len(s) % 2 else (s[mid - 1] + s[mid]) / 2


def r2(x, n=4):
    return "" if x is None else round(x, n)


def pct(n, total):
    return round(n / total * 100, 2) if total else ""


def stars(p):
    if p is None:
        return ""
    if p < 0.001:
        return "***"
    if p < 0.01:
        return "**"
    if p < 0.05:
        return "*"
    return ""


def _betacf(a, b, x):
    tiny = 1e-300
    qab, qap, qam = a + b, a + 1, a - 1
    c = 1.0
    d = 1 - qab * x / qap
    d = tiny if abs(d) < tiny else d
    d = 1 / d
    h = d
    for m in range(1, 201):
        m2 = 2 * m
        aa = m * (b - m) * x / ((qam + m2) * (a + m2))
        d = 1 + aa * d
        d = tiny if abs(d) < tiny else d
        c = 1 + aa / c
        c = tiny if abs(c) < tiny else c
        d = 1 / d
        h *= d * c
        aa = -(a + m) * (qab + m) * x / ((a + m2) * (qap + m2))
        d = 1 + aa * d
        d = tiny if abs(d) < tiny else d
        c = 1 + aa / c
        c = tiny if abs(c) < tiny else c
        d = 1 / d
        de = d * c
        h *= de
        if abs(de - 1) < 3e-12:
            break
    return h


def betainc(a, b, x):
    if x <= 0:
        return 0.0
    if x >= 1:
        return 1.0
    bt = math.exp(math.lgamma(a + b) - math.lgamma(a) - math.lgamma(b)
                  + a * math.log(x) + b * math.log(1 - x))
    if x < (a + 1) / (a + b + 2):
        return bt * _betacf(a, b, x) / a
    return 1 - bt * _betacf(b, a, 1 - x) / b


def t_pvalue(t, df):
    return betainc(df / 2, 0.5, df / (df + t * t))


def welch(a, b):
    """Welch t-testi: (t, df, p, Cohen d) yoki None."""
    na, nb = len(a), len(b)
    if na < 2 or nb < 2:
        return None
    ma, mb = mean(a), mean(b)
    va, vb = stdev(a) ** 2, stdev(b) ** 2
    se2 = va / na + vb / nb
    if se2 == 0:
        return None
    t = (ma - mb) / se2 ** 0.5
    df = se2 ** 2 / ((va / na) ** 2 / (na - 1) + (vb / nb) ** 2 / (nb - 1))
    sp = (((na - 1) * va + (nb - 1) * vb) / (na + nb - 2)) ** 0.5
    d = (ma - mb) / sp if sp else None
    return t, df, t_pvalue(t, df), d


def pearson(x, y):
    n = len(x)
    if n < 3:
        return None
    mx, my = mean(x), mean(y)
    sxx = sum((a - mx) ** 2 for a in x)
    syy = sum((b - my) ** 2 for b in y)
    sxy = sum((a - mx) * (b - my) for a, b in zip(x, y))
    if sxx == 0 or syy == 0:
        return None
    return sxy / (sxx * syy) ** 0.5


def corr_p(r, n):
    if r is None or n < 4:
        return None
    if abs(r) >= 1:
        return 0.0
    return t_pvalue(r * ((n - 2) / (1 - r * r)) ** 0.5, n - 2)


def cronbach(matrix):
    """Kronbax alfa. matrix: to‘liq javob bergan respondentlar qatorlari."""
    n = len(matrix)
    if n < 3 or len(matrix[0]) < 2:
        return None
    k = len(matrix[0])
    item_var = sum(stdev(list(c)) ** 2 for c in zip(*matrix))
    total_var = stdev([sum(r) for r in matrix]) ** 2
    if total_var == 0:
        return None
    return k / (k - 1) * (1 - item_var / total_var)


def build_people(rows):
    people = []
    for r in rows:
        d = r["data"]
        ans = d.get("likert_answers", {})
        items = {k: ans[k] for k, _ in CRITERIA if isinstance(ans.get(k), int)}
        complete = len(items) == len(CRITERIA)
        cons = {}
        for name, code, keys in CONSTRUCTS:
            vals = [items[k] for k in keys if k in items]
            cons[name] = mean(vals) if len(vals) == len(keys) else None
        m = re.match(r"\d+", str(d.get("desired_children", "")))
        age = d.get("ideal_marriage_age")
        people.append({
            "id": r["id"], "gender": r["gender"], "d": d, "items": items,
            "complete": complete, "cons": cons,
            "total": mean(list(items.values())) if complete else None,
            "age": age if isinstance(age, int) else None,
            "kids": int(m.group()) if m else None,
        })
    return people


def export_excel(path):
    rows = load_rows()
    people = build_people(rows)
    total = len(rows)
    male = [p for p in people if p["gender"] == "Erkak"]
    female = [p for p in people if p["gender"] == "Ayol"]
    wb = Workbook()
    wb.remove(wb.active)

    def new_sheet(name, header):
        ws = wb.create_sheet(name)
        ws.append(header)
        return ws

    # ---------- 01 — BARCHA MA’LUMOTLAR ----------
    headers = ["id", "telegram_id", "name", "gender", "created_at", "age_group", "course",
               "residence", "marital_status", "ideal_marriage_age", "desired_children"]
    headers += [f"q{i}" for i in range(1, 26)]
    headers += [f"q{i}_label" for i in range(1, 26)]
    headers += [f"top5_rank_{i}" for i in range(1, 6)]
    headers += [code for _, code, _ in CONSTRUCTS] + ["idx_umumiy", "main_factor", "custom_factor"]
    ws = new_sheet("01_Barcha_Malumotlar", headers)
    for row, p in zip(rows, people):
        d = row["data"]
        values = [row["id"], row["telegram_id"], row["name"], row["gender"], row["created_at"]]
        values += [d.get(k, "") for k in ("age_group", "course", "residence", "marital_status",
                                          "ideal_marriage_age", "desired_children")]
        values += [d.get("likert_answers", {}).get(k, "") for k, _ in CRITERIA]
        values += [label for _, label in CRITERIA]
        top5 = d.get("top5_ranks", [])
        values += [top5[i] if i < len(top5) else "" for i in range(5)]
        values += [r2(p["cons"][name]) for name, _, _ in CONSTRUCTS] + [r2(p["total"])]
        values += [d.get("main_factor", ""), d.get("custom_factor", "")]
        ws.append(values)
    style_sheet(ws)

    # ---------- 02 — DEMOGRAFIYA (xom) ----------
    ws = new_sheet("02_Demografiya", ["id", "gender", "age_group", "course", "residence",
                                      "marital_status", "ideal_marriage_age", "desired_children"])
    for row in rows:
        d = row["data"]
        ws.append([row["id"], row["gender"]] + [d.get(k, "") for k in (
            "age_group", "course", "residence", "marital_status",
            "ideal_marriage_age", "desired_children")])
    style_sheet(ws)

    # ---------- 03 — DEMOGRAFIK TAQSIMOT ----------
    ws = new_sheet("03_Demografik_Taqsimot", ["O‘zgaruvchi", "Variant", "N", "Ulush (%)",
                   "Erkak N", "Erkak (%)", "Ayol N", "Ayol (%)"])
    demo = [("gender", ["Erkak", "Ayol"])] + [(k, opts) for k, _, opts in GENERAL_QUESTIONS]
    for key, options in demo:
        for opt in options:
            def val(p):
                return p["gender"] if key == "gender" else p["d"].get(key, "")
            n = sum(1 for p in people if val(p) == opt)
            nm = sum(1 for p in male if val(p) == opt)
            nf = sum(1 for p in female if val(p) == opt)
            ws.append([key, opt, n, pct(n, total), nm, pct(nm, len(male)),
                       nf, pct(nf, len(female))])
    style_sheet(ws)

    # ---------- 04 — LIKERT TAHLILI ----------
    ws = new_sheet("04_Likert_Tahlil", ["q", "construct", "mezon", "N", "O‘rtacha", "Median",
                   "Standart og‘ish", "0 soni", "1 soni", "2 soni", "3 soni",
                   "2-3 ulushi (%)", "3 ulushi (%)", "Reyting (o‘rtacha bo‘yicha)"])
    stats = []
    for key, label in CRITERIA:
        vals = [p["items"][key] for p in people if key in p["items"]]
        stats.append((vals, mean(vals)))
    means_only = [m for _, m in stats if m is not None]
    for idx, ((key, label), (vals, m)) in enumerate(zip(CRITERIA, stats), start=1):
        n = len(vals)
        counts = [vals.count(i) for i in range(4)]
        rank = 1 + sum(1 for x in means_only if m is not None and x > m) if m is not None else ""
        ws.append([f"q{idx}", key, label, n, r2(m), r2(median(vals)), r2(stdev(vals)), *counts,
                   pct(counts[2] + counts[3], n), pct(counts[3], n), rank])
    style_sheet(ws)

    # ---------- 05 — TOP-5 (o‘rinlar bo‘yicha) ----------
    ws = new_sheet("05_TOP5", ["rank", "mezon", "konstrukt", "N", "Ulush (%)"])
    label_to_key = {label: key for key, label in CRITERIA}
    for rank in range(1, 6):
        counts = {}
        for p in people:
            ranks = p["d"].get("top5_ranks", [])
            if len(ranks) >= rank:
                counts[ranks[rank - 1]] = counts.get(ranks[rank - 1], 0) + 1
        for label, count in sorted(counts.items(), key=lambda x: (-x[1], x[0])):
            ws.append([rank, label, label_to_key.get(label, ""), count, pct(count, total)])
    style_sheet(ws)

    # ---------- 06 — TOP-5 UMUMIY (ballar) ----------
    ws = new_sheet("06_TOP5_Umumiy", ["mezon", "konstrukt", "TOP-5 da uchrashi (N)", "Ulush (%)",
                   "Ball (1-o‘rin=5 ... 5-o‘rin=1)", "Erkak N", "Erkak (%)", "Ayol N", "Ayol (%)"])
    table = []
    for key, label in CRITERIA:
        n = score = nm = nf = 0
        for p in people:
            ranks = p["d"].get("top5_ranks", [])
            if label in ranks:
                n += 1
                score += 5 - ranks.index(label)
                if p["gender"] == "Erkak":
                    nm += 1
                elif p["gender"] == "Ayol":
                    nf += 1
        table.append([label, key, n, pct(n, total), score, nm, pct(nm, len(male)),
                      nf, pct(nf, len(female))])
    for line in sorted(table, key=lambda x: (-x[4], -x[2], x[0])):
        ws.append(line)
    style_sheet(ws)

    # ---------- 07 — GENDER TAHLILI (t-test) ----------
    ws = new_sheet("07_Gender_Tahlil", ["q", "konstrukt", "mezon", "Erkak N", "Erkak o‘rtacha",
                   "Erkak SD", "Ayol N", "Ayol o‘rtacha", "Ayol SD", "Farq (Erkak-Ayol)",
                   "t", "df", "p-qiymat", "Cohen d", "Ma’nodorlik", "Bonferroni (p<0.002)"])
    for idx, (key, label) in enumerate(CRITERIA, start=1):
        a = [p["items"][key] for p in male if key in p["items"]]
        b = [p["items"][key] for p in female if key in p["items"]]
        res = welch(a, b)
        t, df, pv, cd = res if res else (None, None, None, None)
        diff = mean(a) - mean(b) if a and b else None
        ws.append([f"q{idx}", key, label, len(a), r2(mean(a)), r2(stdev(a)), len(b),
                   r2(mean(b)), r2(stdev(b)), r2(diff), r2(t), r2(df, 2), r2(pv, 5), r2(cd),
                   stars(pv), "" if pv is None else ("Ha" if pv < 0.002 else "Yo‘q")])
    style_sheet(ws)

    # ---------- 08 — KONSTRUKTLAR (indekslar, ishonchlilik) ----------
    ws = new_sheet("08_Konstruktlar", ["Konstrukt", "Mezonlar soni", "N (to‘liq)", "O‘rtacha",
                   "SD", "Kronbax alfa", "Erkak o‘rtacha", "Ayol o‘rtacha", "Farq (E-A)",
                   "p-qiymat", "Cohen d", "Ma’nodorlik", "Yakuniy omil sifatida tanlagan (%)",
                   "Reyting"])
    cons_rows = []
    factor_total = sum(1 for p in people if p["d"].get("main_factor"))
    for name, code, keys in CONSTRUCTS:
        vals = [p["cons"][name] for p in people if p["cons"][name] is not None]
        a = [p["cons"][name] for p in male if p["cons"][name] is not None]
        b = [p["cons"][name] for p in female if p["cons"][name] is not None]
        res = welch(a, b)
        t, df, pv, cd = res if res else (None, None, None, None)
        matrix = [[p["items"][k] for k in keys] for p in people if p["complete"]]
        chosen = sum(1 for p in people if p["d"].get("main_factor") == name)
        cons_rows.append([name, len(keys), len(vals), mean(vals), stdev(vals),
                          cronbach(matrix) if matrix else None, mean(a), mean(b),
                          mean(a) - mean(b) if a and b else None, pv, cd, stars(pv),
                          pct(chosen, factor_total)])
    all_means = [r[3] for r in cons_rows if r[3] is not None]
    for r in cons_rows:
        rank = 1 + sum(1 for x in all_means if r[3] is not None and x > r[3]) if r[3] is not None else ""
        ws.append([r[0], r[1], r[2], r2(r[3]), r2(r[4]), r2(r[5]), r2(r[6]), r2(r[7]), r2(r[8]),
                   r2(r[9], 5), r2(r[10]), r[11], r[12], rank])
    tot_vals = [p["total"] for p in people if p["total"] is not None]
    ta = [p["total"] for p in male if p["total"] is not None]
    tb = [p["total"] for p in female if p["total"] is not None]
    res = welch(ta, tb)
    t, df, pv, cd = res if res else (None, None, None, None)
    full = [[p["items"][k] for k, _ in CRITERIA] for p in people if p["complete"]]
    ws.append(["UMUMIY SHKALA (25 mezon)", 25, len(tot_vals), r2(mean(tot_vals)),
               r2(stdev(tot_vals)), r2(cronbach(full) if full else None), r2(mean(ta)),
               r2(mean(tb)), r2(mean(ta) - mean(tb) if ta and tb else None), r2(pv, 5),
               r2(cd), stars(pv), "", ""])
    style_sheet(ws)

    # ---------- 09 — GURUHLAR TAQQOSI ----------
    cnames = [n for n, _, _ in CONSTRUCTS]
    ws = new_sheet("09_Guruhlar_Taqqosi", ["O‘zgaruvchi", "Guruh", "N"] + cnames + ["Umumiy indeks"])
    group_vars = [("age_group", [o for k, _, o in GENERAL_QUESTIONS if k == "age_group"][0]),
                  ("course", [o for k, _, o in GENERAL_QUESTIONS if k == "course"][0]),
                  ("residence", ["Shahar", "Tuman/Qishloq"]),
                  ("marital_status", ["Turmush qurmagan", "Turmush qurgan"])]
    for key, options in group_vars:
        groups = []
        for opt in options:
            members = [p for p in people if p["d"].get(key) == opt]
            groups.append(members)
            line = [key, opt, len(members)]
            for name in cnames:
                line.append(r2(mean([p["cons"][name] for p in members if p["cons"][name] is not None])))
            line.append(r2(mean([p["total"] for p in members if p["total"] is not None])))
            ws.append(line)
        if len(groups) == 2:
            line = [key, "p-qiymat (Welch)", ""]
            for name in cnames + [None]:
                x = [(p["total"] if name is None else p["cons"][name]) for p in groups[0]]
                y = [(p["total"] if name is None else p["cons"][name]) for p in groups[1]]
                res = welch([v for v in x if v is not None], [v for v in y if v is not None])
                line.append(r2(res[2], 5) if res else "")
            ws.append(line)
    style_sheet(ws)

    # ---------- 10 — KONSTRUKTLAR KORRELYATSIYASI ----------
    variables = [(name, lambda p, n=name: p["cons"][n]) for name in cnames]
    variables += [("Umumiy indeks", lambda p: p["total"]),
                  ("Nikoh yoshi", lambda p: p["age"]),
                  ("Istalgan farzandlar soni", lambda p: p["kids"])]
    names = [v[0] for v in variables]
    ws = new_sheet("10_Korrelyatsiya", ["Pearson r"] + names)
    pm = []
    for n1, f1 in variables:
        line, pline = [n1], [n1]
        for n2, f2 in variables:
            pairs = [(f1(p), f2(p)) for p in people if f1(p) is not None and f2(p) is not None]
            r = pearson([a for a, _ in pairs], [b for _, b in pairs]) if n1 != n2 else 1.0
            line.append(r2(r))
            pline.append(r2(corr_p(r, len(pairs)) if n1 != n2 else None, 5))
        ws.append(line)
        pm.append(pline)
    ws.append([])
    ws.append(["p-qiymatlar"] + names)
    for pline in pm:
        ws.append(pline)
    style_sheet(ws)

    # ---------- 11 — 25 MEZON KORRELYATSIYASI ----------
    ws = new_sheet("11_Mezon_Korrelyatsiya", ["q / Pearson r"] + [f"q{i}" for i in range(1, 26)])
    for i, (k1, _) in enumerate(CRITERIA, start=1):
        line = [f"q{i} {k1}"]
        for j, (k2, _) in enumerate(CRITERIA, start=1):
            if i == j:
                line.append(1.0)
                continue
            pairs = [(p["items"][k1], p["items"][k2]) for p in people
                     if k1 in p["items"] and k2 in p["items"]]
            line.append(r2(pearson([a for a, _ in pairs], [b for _, b in pairs])))
        ws.append(line)
    style_sheet(ws)

    # ---------- 12 — NIKOH YOSHI VA FARZAND ----------
    ws = new_sheet("12_Nikoh_Yoshi_Farzand", ["Ko‘rsatkich", "Guruh", "N", "O‘rtacha", "Median",
                   "SD", "Min", "Max", "p-qiymat (E vs A)"])
    for title, field in (("Nikohning maqbul yoshi", "age"), ("Istalgan farzandlar soni", "kids")):
        groups = [("Jami", people), ("Erkak", male), ("Ayol", female)]
        for gname, members in groups:
            v = [p[field] for p in members if p[field] is not None]
            ws.append([title, gname, len(v), r2(mean(v)), r2(median(v)), r2(stdev(v)),
                       min(v) if v else "", max(v) if v else "", ""])
        a = [p[field] for p in male if p[field] is not None]
        b = [p[field] for p in female if p[field] is not None]
        res = welch(a, b)
        ws.append([title, "Farq (Erkak-Ayol)", "", r2(mean(a) - mean(b) if a and b else None),
                   "", "", "", "", r2(res[2], 5) if res else ""])
    style_sheet(ws)

    # ---------- 13 — YAKUNIY OMILLAR ----------
    ws = new_sheet("13_Faktorlar", ["Faktor", "N", "Ulush (%)", "Erkak N", "Erkak (%)",
                                    "Ayol N", "Ayol (%)"])
    fm = sum(1 for p in male if p["d"].get("main_factor"))
    ff = sum(1 for p in female if p["d"].get("main_factor"))
    for factor in MAIN_FACTORS:
        n = sum(1 for p in people if p["d"].get("main_factor") == factor)
        nm = sum(1 for p in male if p["d"].get("main_factor") == factor)
        nf = sum(1 for p in female if p["d"].get("main_factor") == factor)
        ws.append([factor, n, pct(n, factor_total), nm, pct(nm, fm), nf, pct(nf, ff)])
    customs = [(p["id"], p["gender"], p["d"].get("custom_factor", "")) for p in people
               if p["d"].get("custom_factor")]
    if customs:
        ws.append([])
        ws.append(["«Boshqa» — erkin javoblar", "id", "Jins"])
        for rid, gender, text in customs:
            ws.append([text, rid, gender])
    style_sheet(ws)

    # ---------- 14 — STATA KODLARI ----------
    ws = new_sheet("14_Stata_Kodlar", ["Stata kodi"])
    for line in [
        "* q1-q25: Likert 0-3 (0=Umuman muhim emas ... 3=Juda muhim)",
        "* idx_xarakter, idx_intellekt, idx_korinish, idx_oilaviy, idx_iqtisodiy, idx_umumiy",
        "",
        "summarize q1-q25, detail",
        "alpha q1-q25",
        "alpha q1 q2 q3 q4 q5 q10",
        "pwcorr q1-q25, sig",
        "tabulate gender",
        "tabstat q1-q25, by(gender) stat(n mean sd)",
        "foreach v of varlist q1-q25 {",
        "    ttest `v', by(gender)",
        "}",
        "tabstat idx_*, by(gender) stat(n mean sd)",
        "foreach v of varlist idx_* {",
        "    ttest `v', by(gender)",
        "}",
        "pwcorr idx_* ideal_marriage_age, sig",
        "regress idx_umumiy ideal_marriage_age i.gender",
        "tabulate main_factor gender, chi2 column",
        "factor q1-q25, pcf",
        "rotate, varimax",
    ]:
        ws.append([line])
    style_sheet(ws)

    # ---------- 15 — CODEBOOK ----------
    ws = new_sheet("15_Codebook", ["O‘zgaruvchi", "Mazmuni", "Tip"])
    fixed = [("id", "Respondent identifikatori", "numeric"),
             ("telegram_id", "Telegram identifikatori", "numeric"),
             ("gender", "Jins", "categorical"),
             ("age_group", "Yosh guruhi", "categorical"),
             ("course", "Kurs", "categorical"),
             ("residence", "Doimiy yashash joyi", "categorical"),
             ("marital_status", "Oilaviy holat", "categorical"),
             ("ideal_marriage_age", "Nikohning maqbul yoshi", "numeric"),
             ("desired_children", "Istalgan farzandlar soni", "categorical"),
             ("main_factor", "Yakuniy asosiy omil", "categorical")]
    for item in fixed:
        ws.append(list(item))
    for i, (key, label) in enumerate(CRITERIA, start=1):
        ws.append([f"q{i}", label, "Likert 0–3"])
    for name, code, keys in CONSTRUCTS:
        ws.append([code, f"{name} indeksi (mezonlar o‘rtachasi, 0–3)", "numeric"])
    ws.append(["idx_umumiy", "Umumiy indeks (25 mezon o‘rtachasi, 0–3)", "numeric"])
    for i in range(1, 6):
        ws.append([f"top5_rank_{i}", f"TOP-5 dagi {i}-o‘rin", "text"])
    style_sheet(ws)

    # ---------- 16 — UMUMIY STATISTIKA ----------
    ws = new_sheet("16_Umumiy_Statistika", ["Ko‘rsatkich", "Qiymat"])
    for line in [("Jami respondentlar", total), ("Erkak respondentlar", len(male)),
                 ("Ayol respondentlar", len(female)),
                 ("25 ta mezonga to‘liq javob berganlar", sum(1 for p in people if p["complete"])),
                 ("Yakuniy omilni tanlaganlar", factor_total),
                 ("Test savollari soni", 25), ("TOP-5 mezonlar soni", 5),
                 ("Konstruktlar soni", len(CONSTRUCTS))]:
        ws.append(list(line))
    style_sheet(ws)

    # ---------- 17 — XOM MA’LUMOT ----------
    ws = new_sheet("17_Xom_SQL_Malumot", ["id", "telegram_id", "name", "gender", "created_at", "data"])
    for row in rows:
        ws.append([row["id"], row["telegram_id"], row["name"], row["gender"],
                   row["created_at"], repr(row["data"])])
    style_sheet(ws)

    wb.save(path)


# 12. ADMIN STATISTIKA

def admin_stats_text():
    rows = load_rows()
    total = len(rows)
    male = sum(1 for r in rows if r["gender"] == "Erkak")
    female = sum(1 for r in rows if r["gender"] == "Ayol")
    return ("ADMIN STATISTIKA\n\n"
            f"Jami respondent: {total}\n"
            f"Erkak: {male}\n"
            f"Ayol: {female}\n"
            "Test savollari: 25\n"
            "TOP-5: 5 ta tanlov")

# 13. MAIN

async def main():
    get_conn().close()

    # RENDER HEALTH CHECK
    app = web.Application()

    async def health(request):
        return web.Response(text="Bot ishlamoqda")
    app.router.add_get("/", health)
    runner = web.AppRunner(app)
    await runner.setup()
    port = int(os.environ.get("PORT", "8080"))
    site = web.TCPSite(runner, "0.0.0.0", port)
    await site.start()

    # TELEGRAM BOT
    bot = Bot(TOKEN)
    dp = Dispatcher()

    # START

    @dp.message(CommandStart())
    async def start(message: Message, state: FSMContext):
        if (message.from_user.id != ADMIN_ID and has_completed_response(message.from_user.id)):
            await state.clear()
            await message.answer("Siz ushbu so‘rovnomada " "avval qatnashgansiz.\n\n"
                "Ishtirokingiz uchun rahmat!")
            return
        await start_survey(message, state)

    # ADMIN

    @dp.message(Command("admin"))
    async def admin_panel(message: Message):
        if (message.from_user.id != ADMIN_ID):
            return
        await message.answer("ADMIN PANEL\n\n" "Kerakli amalni tanlang:", reply_markup=admin_keyboard(
            ))

    @dp.callback_query(F.data == "admin:stats")
    async def admin_stats(call: CallbackQuery):
        if (call.from_user.id != ADMIN_ID):
            await call.answer("Ruxsat yo‘q", show_alert=True)
            return
        await call.answer()
        await call.message.answer(admin_stats_text())

    @dp.callback_query(F.data == "admin:export")
    async def admin_export(call: CallbackQuery):
        if (call.from_user.id != ADMIN_ID):
            await call.answer("Ruxsat yo‘q", show_alert=True)
            return
        await call.answer("Excel tayyorlanmoqda...")
        try:
            export_excel(EXPORT_PATH)
            await call.message.answer_document(document=FSInputFile(EXPORT_PATH), caption=(
                "Excel tayyor.\n" "25 ta test, TOP-5, gender t-test, konstruktlar, korrelyatsiya, Stata "
                "va Codebook kiritilgan."))
        except Exception as e:
            await call.message.answer("Excel eksportida xatolik:\n" f"{type(e).__name__}: {e}")

    @dp.callback_query(F.data == "admin:reset")
    async def admin_reset(call: CallbackQuery):
        if (call.from_user.id != ADMIN_ID):
            await call.answer("Ruxsat yo‘q", show_alert=True)
            return
        delete_user_response(call.from_user.id)
        await call.answer("Test yozuvingiz o‘chirildi.")
        await call.message.answer("Mening testim o‘chirildi. "
            "Endi qayta boshlashingiz mumkin.")

    @dp.callback_query(F.data == "admin:reset_all")
    async def admin_reset_all(call: CallbackQuery):
        if (call.from_user.id != ADMIN_ID):
            await call.answer("Ruxsat yo‘q", show_alert=True)
            return
        delete_all_responses()
        await call.answer("Barcha ma’lumot o‘chirildi.", show_alert=True)
        await call.message.answer("Barcha respondent ma’lumotlari o‘chirildi.")

    @dp.callback_query(F.data == "admin:start")
    async def admin_start(call: CallbackQuery, state: FSMContext):
        if (call.from_user.id != ADMIN_ID):
            await call.answer("Ruxsat yo‘q", show_alert=True)
            return
        await call.answer()
        await start_survey(call.message, state)

    # ISM

    @dp.callback_query(Survey.name, F.data == "skip_name:0")
    async def skip_name(call: CallbackQuery, state: FSMContext):
        await state.update_data(name="")
        await state.set_state(Survey.gender)
        await call.answer()
        await call.message.edit_reply_markup(reply_markup=None)
        await call.message.answer("Jinsingizni tanlang:", reply_markup=gender_buttons())

    @dp.message(Survey.name)
    async def get_name(message: Message, state: FSMContext):
        name = (message.text or "").strip()
        if not name:
            await message.answer("Ismingizni yozing yoki "
                "«O‘tkazib yuborish» tugmasini bosing.")
            return
        await state.update_data(name=name)
        await state.set_state(Survey.gender)
        await message.answer("Jinsingizni tanlang:", reply_markup=gender_buttons())

    # JINS

    @dp.callback_query(Survey.gender, F.data.startswith("gender:"))
    async def get_gender(call: CallbackQuery, state: FSMContext):
        gender = call.data.split(":", 1)[1]
        await state.update_data(gender=gender, general_index=0, general_answers={},
            likert_answers={}, top5_selected=[], top5_ranks=[])
        await call.answer()
        await call.message.edit_reply_markup(reply_markup=None)
        await send_general(call.message, state, 0)

    # DEMOGRAFIYA

    @dp.callback_query(Survey.general, F.data.startswith("general:"))
    async def get_general(call: CallbackQuery, state: FSMContext):
        data = await state.get_data()
        question_index = data.get("general_index", 0)
        option_index = int(call.data.split(":")[1])
        key, question, options = GENERAL_QUESTIONS[question_index]
        selected_answer = options[option_index]
        general_answers = data.get("general_answers", {})
        general_answers[key] = selected_answer
        next_index = question_index + 1
        await state.update_data(general_answers=general_answers, general_index=next_index)
        await call.answer()
        await call.message.edit_reply_markup(reply_markup=None)
        await send_general(call.message, state, next_index)

    # NIKOH YOSHI

    @dp.message(Survey.marriage_age)
    async def get_marriage_age(message: Message, state: FSMContext):
        value = (message.text or "").strip()
        try:
            age = int(value)
            if (age < 15 or age > 80):
                raise ValueError
        except ValueError:
            await message.answer("Iltimos, 15–80 oralig‘idagi " "yoshni raqam bilan kiriting.\n"
                "Masalan: 25")
            return
        await state.update_data(ideal_marriage_age=age)
        await send_likert(message, state, 0)

    # 25 TA TEST

    @dp.callback_query(Survey.likert, F.data.startswith("likert:"))
    async def get_likert(call: CallbackQuery, state: FSMContext):
        data = await state.get_data()
        index = data.get("likert_index", 0)
        value = int(call.data.split(":")[1])
        key, label = CRITERIA[index]
        answers = data.get("likert_answers", {})
        answers[key] = value
        await state.update_data(likert_answers=answers, likert_index=index + 1)
        await call.answer()
        await call.message.edit_reply_markup(reply_markup=None)
        await send_likert(call.message, state, index + 1)

    # TOP-5

    @dp.callback_query(Survey.top5, F.data.startswith("top5:"))
    async def get_top5(call: CallbackQuery, state: FSMContext):
        action = call.data.split(":", 1)[1]
        data = await state.get_data()
        selected = list(data.get("top5_selected", []))
        if action == "count":
            await call.answer(f"Tanlangan: " f"{len(selected)}/5", show_alert=True)
            return
        if action == "confirm":
            if len(selected) != 5:
                await call.answer("Avval 5 ta mezon tanlang.", show_alert=True)
                return
            ranks = [CRITERIA[i][1] for i in selected]
            await state.update_data(top5_ranks=ranks)
            await call.answer()
            await call.message.edit_reply_markup(reply_markup=None)
            await send_main_factor(call.message, state)
            return
        try:
            idx = int(action)
        except ValueError:
            await call.answer()
            return
        if idx in selected:
            selected.remove(idx)
        else:
            if len(selected) >= 5:
                await call.answer("Faqat 5 ta mezon " "tanlash mumkin.", show_alert=True)
                return
            selected.append(idx)
        await state.update_data(top5_selected=selected)
        await call.answer()
        await call.message.edit_reply_markup(reply_markup=top5_keyboard(selected))

    # YAKUNIY OMIL

    @dp.callback_query(Survey.main_factor, F.data.startswith("factor:"))
    async def get_factor(call: CallbackQuery, state: FSMContext):
        idx = int(call.data.split(":")[1])
        factor = MAIN_FACTORS[idx]
        await call.answer()
        await call.message.edit_reply_markup(reply_markup=None)
        if factor == "Boshqa":
            await state.set_state(Survey.custom_factor)
            await call.message.answer("Iltimos, siz uchun " "boshqa eng muhim " "omilni yozing:")
            return
        await state.update_data(main_factor=factor, custom_factor="")
        await finish_survey(call.message, state)

    @dp.message(Survey.custom_factor)
    async def get_custom_factor(message: Message, state: FSMContext):
        custom = (message.text or "").strip()
        if not custom:
            await message.answer("Iltimos, omilni yozib yuboring.")
            return
        await state.update_data(main_factor="Boshqa", custom_factor=custom)
        await finish_survey(message, state)

    # YAKUNLASH

    async def finish_survey(message, state):
        data = await state.get_data()
        final_data = {"age_group": data.get("general_answers", {}).get("age_group", ""),
            "course": data.get("general_answers", {}).get("course", ""), "residence": data.get(
            "general_answers", {}).get("residence", ""), "marital_status": data.get(
            "general_answers", {}).get("marital_status", ""), "desired_children": data.get(
            "general_answers", {}).get("desired_children", ""), "ideal_marriage_age": data.get(
            "ideal_marriage_age", ""), "likert_answers": data.get("likert_answers", {}),
            "top5_ranks": data.get("top5_ranks", []), "main_factor": data.get(
            "main_factor", ""), "custom_factor": data.get("custom_factor", "")}
        save_response(message.from_user.id, data.get("name", ""), data.get("gender", ""),
            final_data)
        await state.clear()
        await message.answer("So‘rovnoma yakunlandi!\n\n" "Ishtirokingiz uchun katta rahmat.\n"
            "Sizning javoblaringiz " "ilmiy-statistik tahlil uchun " "saqlab qo‘yildi.")

    # POLLING
    try:
        await dp.start_polling(bot)
    finally:
        await bot.session.close()
        await runner.cleanup()

# START
if __name__ == "__main__":
    import asyncio
    asyncio.run(main())
    return kb.as_markup()

# 10. YORDAMCHI FUNKSIYALAR

def target_label(gender):
    if gender == "Ayol":
        return "erkak"
    return "ayol"

def calculate_sd_level(score):
    if score <= 1:
        return "Past"
    if score <= 4:
        return "O‘rtacha"
    return "Yuqori"

async def start_survey(message, state):
    await state.clear()
    await state.set_state(Survey.name)
    await message.answer("Assalomu alaykum!\n\n" "Oila va nikoh mezonlari bo‘yicha "
        "ilmiy tadqiqot so‘rovnomasiga xush kelibsiz.\n\n" "Ismingizni kiriting yoki "
        "«O‘tkazib yuborish» tugmasini bosing.\n" "Ism-familiya berish majburiy emas.",
        reply_markup=option_buttons(["O‘tkazib yuborish"], "skip_name"))

async def send_general(message, state, index):
    if index >= len(GENERAL_QUESTIONS):
        await state.set_state(Survey.marriage_age)
        await message.answer("Siz uchun nikoh qurishning " "maqbul yoshi nechada?\n\n"
            "Javobingizni faqat raqam bilan yozing.\n" "Masalan: 25")
        return
    key, question, options = GENERAL_QUESTIONS[index]
    await state.update_data(general_index=index)
    await state.set_state(Survey.general)
    await message.answer(f"DEMOGRAFIK SAVOL " f"{index + 1}/{len(GENERAL_QUESTIONS)}\n\n"
        f"{question}", reply_markup=option_buttons(options, "general"))

async def send_likert(message, state, index):
    if index >= len(CRITERIA):
        await state.update_data(top5_selected=[])
        await state.set_state(Survey.top5)
        await message.answer("TOP-5 TANLOV\n\n" "Yuqoridagi 25 ta test mezonidan "
            "siz uchun eng muhim 5 tasini tanlang.\n\n" "Birinchi bosganingiz — 1-o‘rin,\n"
            "ikkinchi bosganingiz — 2-o‘rin,\n" "shu tartibda 5-o‘ringacha.\n\n"
            "Tanlangan: 0/5", reply_markup=top5_keyboard([]))
        return
    data = await state.get_data()
    target = target_label(data.get("gender"))
    key, label = CRITERIA[index]
    await state.update_data(likert_index=index)
    await state.set_state(Survey.likert)
    await message.answer(f"TEST SAVOLI " f"{index + 1}/25\n\n" f"Bo‘lajak {target}ning "
        "quyidagi xususiyati siz uchun " "qanchalik muhim?\n\n" f"{label}", reply_markup=likert_buttons(
        ))

async def send_sd(message, state, index):
    if index >= len(SD_QUESTIONS):
        await state.set_state(Survey.main_factor)
        await message.answer("YAKUNIY SAVOL\n\n" "Umuman olganda, juft tanlashda "
            "siz uchun qaysi umumiy yo‘nalish " "eng muhim?", reply_markup=option_buttons(
            MAIN_FACTORS, "factor"))
        return
    key, question, expected = SD_QUESTIONS[index]
    await state.update_data(sd_index=index)
    await state.set_state(Survey.sd_scale)
    await message.answer(f"SD SAVOLI " f"{index + 1}/{len(SD_QUESTIONS)}\n\n" f"{question}",
        reply_markup=sd_buttons())

# 11. EXCEL EKSPORT

def style_sheet(ws):
    thin = Side(style="thin")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    for row in ws.iter_rows():
        for cell in row:
            cell.alignment = Alignment(vertical="top", wrap_text=True)
            cell.border = border
    for cell in ws[1]:
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    for col in ws.columns:
        max_len = 0
        letter = get_column_letter(col[0].column)
        for cell in col:
            value = "" if cell.value is None else str(cell.value)
            max_len = max(max_len, len(value))
        ws.column_dimensions[letter].width = min(max(max_len + 2, 12), 45)
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions


# ---------- statistik yordamchilar (tashqi kutubxonasiz) ----------

def mean(v):
    return sum(v) / len(v) if v else None


def stdev(v):
    n = len(v)
    if n < 2:
        return None
    m = sum(v) / n
    return (sum((x - m) ** 2 for x in v) / (n - 1)) ** 0.5


def median(v):
    if not v:
        return None
    s = sorted(v)
    mid = len(s) // 2
    return s[mid] if len(s) % 2 else (s[mid - 1] + s[mid]) / 2


def r2(x, n=4):
    return "" if x is None else round(x, n)


def pct(n, total):
    return round(n / total * 100, 2) if total else ""


def stars(p):
    if p is None:
        return ""
    if p < 0.001:
        return "***"
    if p < 0.01:
        return "**"
    if p < 0.05:
        return "*"
    return ""


def _betacf(a, b, x):
    tiny = 1e-300
    qab, qap, qam = a + b, a + 1, a - 1
    c = 1.0
    d = 1 - qab * x / qap
    d = tiny if abs(d) < tiny else d
    d = 1 / d
    h = d
    for m in range(1, 201):
        m2 = 2 * m
        aa = m * (b - m) * x / ((qam + m2) * (a + m2))
        d = 1 + aa * d
        d = tiny if abs(d) < tiny else d
        c = 1 + aa / c
        c = tiny if abs(c) < tiny else c
        d = 1 / d
        h *= d * c
        aa = -(a + m) * (qab + m) * x / ((a + m2) * (qap + m2))
        d = 1 + aa * d
        d = tiny if abs(d) < tiny else d
        c = 1 + aa / c
        c = tiny if abs(c) < tiny else c
        d = 1 / d
        de = d * c
        h *= de
        if abs(de - 1) < 3e-12:
            break
    return h


def betainc(a, b, x):
    if x <= 0:
        return 0.0
    if x >= 1:
        return 1.0
    bt = math.exp(math.lgamma(a + b) - math.lgamma(a) - math.lgamma(b)
                  + a * math.log(x) + b * math.log(1 - x))
    if x < (a + 1) / (a + b + 2):
        return bt * _betacf(a, b, x) / a
    return 1 - bt * _betacf(b, a, 1 - x) / b


def t_pvalue(t, df):
    return betainc(df / 2, 0.5, df / (df + t * t))


def welch(a, b):
    """Welch t-testi: (t, df, p, Cohen d) yoki None."""
    na, nb = len(a), len(b)
    if na < 2 or nb < 2:
        return None
    ma, mb = mean(a), mean(b)
    va, vb = stdev(a) ** 2, stdev(b) ** 2
    se2 = va / na + vb / nb
    if se2 == 0:
        return None
    t = (ma - mb) / se2 ** 0.5
    df = se2 ** 2 / ((va / na) ** 2 / (na - 1) + (vb / nb) ** 2 / (nb - 1))
    sp = (((na - 1) * va + (nb - 1) * vb) / (na + nb - 2)) ** 0.5
    d = (ma - mb) / sp if sp else None
    return t, df, t_pvalue(t, df), d


def pearson(x, y):
    n = len(x)
    if n < 3:
        return None
    mx, my = mean(x), mean(y)
    sxx = sum((a - mx) ** 2 for a in x)
    syy = sum((b - my) ** 2 for b in y)
    sxy = sum((a - mx) * (b - my) for a, b in zip(x, y))
    if sxx == 0 or syy == 0:
        return None
    return sxy / (sxx * syy) ** 0.5


def corr_p(r, n):
    if r is None or n < 4:
        return None
    if abs(r) >= 1:
        return 0.0
    return t_pvalue(r * ((n - 2) / (1 - r * r)) ** 0.5, n - 2)


def cronbach(matrix):
    """Kronbax alfa. matrix: to‘liq javob bergan respondentlar qatorlari."""
    n = len(matrix)
    if n < 3 or len(matrix[0]) < 2:
        return None
    k = len(matrix[0])
    item_var = sum(stdev(list(c)) ** 2 for c in zip(*matrix))
    total_var = stdev([sum(r) for r in matrix]) ** 2
    if total_var == 0:
        return None
    return k / (k - 1) * (1 - item_var / total_var)


def build_people(rows):
    people = []
    for r in rows:
        d = r["data"]
        ans = d.get("likert_answers", {})
        items = {k: ans[k] for k, _ in CRITERIA if isinstance(ans.get(k), int)}
        complete = len(items) == len(CRITERIA)
        cons = {}
        for name, code, keys in CONSTRUCTS:
            vals = [items[k] for k in keys if k in items]
            cons[name] = mean(vals) if len(vals) == len(keys) else None
        m = re.match(r"\d+", str(d.get("desired_children", "")))
        age = d.get("ideal_marriage_age")
        people.append({
            "id": r["id"], "gender": r["gender"], "d": d, "items": items,
            "complete": complete, "cons": cons,
            "total": mean(list(items.values())) if complete else None,
            "age": age if isinstance(age, int) else None,
            "kids": int(m.group()) if m else None,
        })
    return people


def export_excel(path):
    rows = load_rows()
    people = build_people(rows)
    total = len(rows)
    male = [p for p in people if p["gender"] == "Erkak"]
    female = [p for p in people if p["gender"] == "Ayol"]
    wb = Workbook()
    wb.remove(wb.active)

    def new_sheet(name, header):
        ws = wb.create_sheet(name)
        ws.append(header)
        return ws

    # ---------- 01 — BARCHA MA’LUMOTLAR ----------
    headers = ["id", "telegram_id", "name", "gender", "created_at", "age_group", "course",
               "residence", "marital_status", "ideal_marriage_age", "desired_children"]
    headers += [f"q{i}" for i in range(1, 26)]
    headers += [f"q{i}_label" for i in range(1, 26)]
    headers += [f"top5_rank_{i}" for i in range(1, 6)]
    headers += [code for _, code, _ in CONSTRUCTS] + ["idx_umumiy", "main_factor", "custom_factor"]
    ws = new_sheet("01_Barcha_Malumotlar", headers)
    for row, p in zip(rows, people):
        d = row["data"]
        values = [row["id"], row["telegram_id"], row["name"], row["gender"], row["created_at"]]
        values += [d.get(k, "") for k in ("age_group", "course", "residence", "marital_status",
                                          "ideal_marriage_age", "desired_children")]
        values += [d.get("likert_answers", {}).get(k, "") for k, _ in CRITERIA]
        values += [label for _, label in CRITERIA]
        top5 = d.get("top5_ranks", [])
        values += [top5[i] if i < len(top5) else "" for i in range(5)]
        values += [r2(p["cons"][name]) for name, _, _ in CONSTRUCTS] + [r2(p["total"])]
        values += [d.get("main_factor", ""), d.get("custom_factor", "")]
        ws.append(values)
    style_sheet(ws)

    # ---------- 02 — DEMOGRAFIYA (xom) ----------
    ws = new_sheet("02_Demografiya", ["id", "gender", "age_group", "course", "residence",
                                      "marital_status", "ideal_marriage_age", "desired_children"])
    for row in rows:
        d = row["data"]
        ws.append([row["id"], row["gender"]] + [d.get(k, "") for k in (
            "age_group", "course", "residence", "marital_status",
            "ideal_marriage_age", "desired_children")])
    style_sheet(ws)

    # ---------- 03 — DEMOGRAFIK TAQSIMOT ----------
    ws = new_sheet("03_Demografik_Taqsimot", ["O‘zgaruvchi", "Variant", "N", "Ulush (%)",
                   "Erkak N", "Erkak (%)", "Ayol N", "Ayol (%)"])
    demo = [("gender", ["Erkak", "Ayol"])] + [(k, opts) for k, _, opts in GENERAL_QUESTIONS]
    for key, options in demo:
        for opt in options:
            def val(p):
                return p["gender"] if key == "gender" else p["d"].get(key, "")
            n = sum(1 for p in people if val(p) == opt)
            nm = sum(1 for p in male if val(p) == opt)
            nf = sum(1 for p in female if val(p) == opt)
            ws.append([key, opt, n, pct(n, total), nm, pct(nm, len(male)),
                       nf, pct(nf, len(female))])
    style_sheet(ws)

    # ---------- 04 — LIKERT TAHLILI ----------
    ws = new_sheet("04_Likert_Tahlil", ["q", "construct", "mezon", "N", "O‘rtacha", "Median",
                   "Standart og‘ish", "0 soni", "1 soni", "2 soni", "3 soni",
                   "2-3 ulushi (%)", "3 ulushi (%)", "Reyting (o‘rtacha bo‘yicha)"])
    stats = []
    for key, label in CRITERIA:
        vals = [p["items"][key] for p in people if key in p["items"]]
        stats.append((vals, mean(vals)))
    means_only = [m for _, m in stats if m is not None]
    for idx, ((key, label), (vals, m)) in enumerate(zip(CRITERIA, stats), start=1):
        n = len(vals)
        counts = [vals.count(i) for i in range(4)]
        rank = 1 + sum(1 for x in means_only if m is not None and x > m) if m is not None else ""
        ws.append([f"q{idx}", key, label, n, r2(m), r2(median(vals)), r2(stdev(vals)), *counts,
                   pct(counts[2] + counts[3], n), pct(counts[3], n), rank])
    style_sheet(ws)

    # ---------- 05 — TOP-5 (o‘rinlar bo‘yicha) ----------
    ws = new_sheet("05_TOP5", ["rank", "mezon", "konstrukt", "N", "Ulush (%)"])
    label_to_key = {label: key for key, label in CRITERIA}
    for rank in range(1, 6):
        counts = {}
        for p in people:
            ranks = p["d"].get("top5_ranks", [])
            if len(ranks) >= rank:
                counts[ranks[rank - 1]] = counts.get(ranks[rank - 1], 0) + 1
        for label, count in sorted(counts.items(), key=lambda x: (-x[1], x[0])):
            ws.append([rank, label, label_to_key.get(label, ""), count, pct(count, total)])
    style_sheet(ws)

    # ---------- 06 — TOP-5 UMUMIY (ballar) ----------
    ws = new_sheet("06_TOP5_Umumiy", ["mezon", "konstrukt", "TOP-5 da uchrashi (N)", "Ulush (%)",
                   "Ball (1-o‘rin=5 ... 5-o‘rin=1)", "Erkak N", "Erkak (%)", "Ayol N", "Ayol (%)"])
    table = []
    for key, label in CRITERIA:
        n = score = nm = nf = 0
        for p in people:
            ranks = p["d"].get("top5_ranks", [])
            if label in ranks:
                n += 1
                score += 5 - ranks.index(label)
                if p["gender"] == "Erkak":
                    nm += 1
                elif p["gender"] == "Ayol":
                    nf += 1
        table.append([label, key, n, pct(n, total), score, nm, pct(nm, len(male)),
                      nf, pct(nf, len(female))])
    for line in sorted(table, key=lambda x: (-x[4], -x[2], x[0])):
        ws.append(line)
    style_sheet(ws)

    # ---------- 07 — GENDER TAHLILI (t-test) ----------
    ws = new_sheet("07_Gender_Tahlil", ["q", "konstrukt", "mezon", "Erkak N", "Erkak o‘rtacha",
                   "Erkak SD", "Ayol N", "Ayol o‘rtacha", "Ayol SD", "Farq (Erkak-Ayol)",
                   "t", "df", "p-qiymat", "Cohen d", "Ma’nodorlik", "Bonferroni (p<0.002)"])
    for idx, (key, label) in enumerate(CRITERIA, start=1):
        a = [p["items"][key] for p in male if key in p["items"]]
        b = [p["items"][key] for p in female if key in p["items"]]
        res = welch(a, b)
        t, df, pv, cd = res if res else (None, None, None, None)
        diff = mean(a) - mean(b) if a and b else None
        ws.append([f"q{idx}", key, label, len(a), r2(mean(a)), r2(stdev(a)), len(b),
                   r2(mean(b)), r2(stdev(b)), r2(diff), r2(t), r2(df, 2), r2(pv, 5), r2(cd),
                   stars(pv), "" if pv is None else ("Ha" if pv < 0.002 else "Yo‘q")])
    style_sheet(ws)

    # ---------- 08 — KONSTRUKTLAR (indekslar, ishonchlilik) ----------
    ws = new_sheet("08_Konstruktlar", ["Konstrukt", "Mezonlar soni", "N (to‘liq)", "O‘rtacha",
                   "SD", "Kronbax alfa", "Erkak o‘rtacha", "Ayol o‘rtacha", "Farq (E-A)",
                   "p-qiymat", "Cohen d", "Ma’nodorlik", "Yakuniy omil sifatida tanlagan (%)",
                   "Reyting"])
    cons_rows = []
    factor_total = sum(1 for p in people if p["d"].get("main_factor"))
    for name, code, keys in CONSTRUCTS:
        vals = [p["cons"][name] for p in people if p["cons"][name] is not None]
        a = [p["cons"][name] for p in male if p["cons"][name] is not None]
        b = [p["cons"][name] for p in female if p["cons"][name] is not None]
        res = welch(a, b)
        t, df, pv, cd = res if res else (None, None, None, None)
        matrix = [[p["items"][k] for k in keys] for p in people if p["complete"]]
        chosen = sum(1 for p in people if p["d"].get("main_factor") == name)
        cons_rows.append([name, len(keys), len(vals), mean(vals), stdev(vals),
                          cronbach(matrix) if matrix else None, mean(a), mean(b),
                          mean(a) - mean(b) if a and b else None, pv, cd, stars(pv),
                          pct(chosen, factor_total)])
    all_means = [r[3] for r in cons_rows if r[3] is not None]
    for r in cons_rows:
        rank = 1 + sum(1 for x in all_means if r[3] is not None and x > r[3]) if r[3] is not None else ""
        ws.append([r[0], r[1], r[2], r2(r[3]), r2(r[4]), r2(r[5]), r2(r[6]), r2(r[7]), r2(r[8]),
                   r2(r[9], 5), r2(r[10]), r[11], r[12], rank])
    tot_vals = [p["total"] for p in people if p["total"] is not None]
    ta = [p["total"] for p in male if p["total"] is not None]
    tb = [p["total"] for p in female if p["total"] is not None]
    res = welch(ta, tb)
    t, df, pv, cd = res if res else (None, None, None, None)
    full = [[p["items"][k] for k, _ in CRITERIA] for p in people if p["complete"]]
    ws.append(["UMUMIY SHKALA (25 mezon)", 25, len(tot_vals), r2(mean(tot_vals)),
               r2(stdev(tot_vals)), r2(cronbach(full) if full else None), r2(mean(ta)),
               r2(mean(tb)), r2(mean(ta) - mean(tb) if ta and tb else None), r2(pv, 5),
               r2(cd), stars(pv), "", ""])
    style_sheet(ws)

    # ---------- 09 — GURUHLAR TAQQOSI ----------
    cnames = [n for n, _, _ in CONSTRUCTS]
    ws = new_sheet("09_Guruhlar_Taqqosi", ["O‘zgaruvchi", "Guruh", "N"] + cnames + ["Umumiy indeks"])
    group_vars = [("age_group", [o for k, _, o in GENERAL_QUESTIONS if k == "age_group"][0]),
                  ("course", [o for k, _, o in GENERAL_QUESTIONS if k == "course"][0]),
                  ("residence", ["Shahar", "Tuman/Qishloq"]),
                  ("marital_status", ["Turmush qurmagan", "Turmush qurgan"])]
    for key, options in group_vars:
        groups = []
        for opt in options:
            members = [p for p in people if p["d"].get(key) == opt]
            groups.append(members)
            line = [key, opt, len(members)]
            for name in cnames:
                line.append(r2(mean([p["cons"][name] for p in members if p["cons"][name] is not None])))
            line.append(r2(mean([p["total"] for p in members if p["total"] is not None])))
            ws.append(line)
        if len(groups) == 2:
            line = [key, "p-qiymat (Welch)", ""]
            for name in cnames + [None]:
                x = [(p["total"] if name is None else p["cons"][name]) for p in groups[0]]
                y = [(p["total"] if name is None else p["cons"][name]) for p in groups[1]]
                res = welch([v for v in x if v is not None], [v for v in y if v is not None])
                line.append(r2(res[2], 5) if res else "")
            ws.append(line)
    style_sheet(ws)

    # ---------- 10 — KONSTRUKTLAR KORRELYATSIYASI ----------
    variables = [(name, lambda p, n=name: p["cons"][n]) for name in cnames]
    variables += [("Umumiy indeks", lambda p: p["total"]),
                  ("Nikoh yoshi", lambda p: p["age"]),
                  ("Istalgan farzandlar soni", lambda p: p["kids"])]
    names = [v[0] for v in variables]
    ws = new_sheet("10_Korrelyatsiya", ["Pearson r"] + names)
    pm = []
    for n1, f1 in variables:
        line, pline = [n1], [n1]
        for n2, f2 in variables:
            pairs = [(f1(p), f2(p)) for p in people if f1(p) is not None and f2(p) is not None]
            r = pearson([a for a, _ in pairs], [b for _, b in pairs]) if n1 != n2 else 1.0
            line.append(r2(r))
            pline.append(r2(corr_p(r, len(pairs)) if n1 != n2 else None, 5))
        ws.append(line)
        pm.append(pline)
    ws.append([])
    ws.append(["p-qiymatlar"] + names)
    for pline in pm:
        ws.append(pline)
    style_sheet(ws)

    # ---------- 11 — 25 MEZON KORRELYATSIYASI ----------
    ws = new_sheet("11_Mezon_Korrelyatsiya", ["q / Pearson r"] + [f"q{i}" for i in range(1, 26)])
    for i, (k1, _) in enumerate(CRITERIA, start=1):
        line = [f"q{i} {k1}"]
        for j, (k2, _) in enumerate(CRITERIA, start=1):
            if i == j:
                line.append(1.0)
                continue
            pairs = [(p["items"][k1], p["items"][k2]) for p in people
                     if k1 in p["items"] and k2 in p["items"]]
            line.append(r2(pearson([a for a, _ in pairs], [b for _, b in pairs])))
        ws.append(line)
    style_sheet(ws)

    # ---------- 12 — NIKOH YOSHI VA FARZAND ----------
    ws = new_sheet("12_Nikoh_Yoshi_Farzand", ["Ko‘rsatkich", "Guruh", "N", "O‘rtacha", "Median",
                   "SD", "Min", "Max", "p-qiymat (E vs A)"])
    for title, field in (("Nikohning maqbul yoshi", "age"), ("Istalgan farzandlar soni", "kids")):
        groups = [("Jami", people), ("Erkak", male), ("Ayol", female)]
        for gname, members in groups:
            v = [p[field] for p in members if p[field] is not None]
            ws.append([title, gname, len(v), r2(mean(v)), r2(median(v)), r2(stdev(v)),
                       min(v) if v else "", max(v) if v else "", ""])
        a = [p[field] for p in male if p[field] is not None]
        b = [p[field] for p in female if p[field] is not None]
        res = welch(a, b)
        ws.append([title, "Farq (Erkak-Ayol)", "", r2(mean(a) - mean(b) if a and b else None),
                   "", "", "", "", r2(res[2], 5) if res else ""])
    style_sheet(ws)

    # ---------- 13 — YAKUNIY OMILLAR ----------
    ws = new_sheet("13_Faktorlar", ["Faktor", "N", "Ulush (%)", "Erkak N", "Erkak (%)",
                                    "Ayol N", "Ayol (%)"])
    fm = sum(1 for p in male if p["d"].get("main_factor"))
    ff = sum(1 for p in female if p["d"].get("main_factor"))
    for factor in MAIN_FACTORS:
        n = sum(1 for p in people if p["d"].get("main_factor") == factor)
        nm = sum(1 for p in male if p["d"].get("main_factor") == factor)
        nf = sum(1 for p in female if p["d"].get("main_factor") == factor)
        ws.append([factor, n, pct(n, factor_total), nm, pct(nm, fm), nf, pct(nf, ff)])
    customs = [(p["id"], p["gender"], p["d"].get("custom_factor", "")) for p in people
               if p["d"].get("custom_factor")]
    if customs:
        ws.append([])
        ws.append(["«Boshqa» — erkin javoblar", "id", "Jins"])
        for rid, gender, text in customs:
            ws.append([text, rid, gender])
    style_sheet(ws)

    # ---------- 14 — STATA KODLARI ----------
    ws = new_sheet("14_Stata_Kodlar", ["Stata kodi"])
    for line in [
        "* q1-q25: Likert 0-3 (0=Umuman muhim emas ... 3=Juda muhim)",
        "* idx_xarakter, idx_intellekt, idx_korinish, idx_oilaviy, idx_iqtisodiy, idx_umumiy",
        "",
        "summarize q1-q25, detail",
        "alpha q1-q25",
        "alpha q1 q2 q3 q4 q5 q10",
        "pwcorr q1-q25, sig",
        "tabulate gender",
        "tabstat q1-q25, by(gender) stat(n mean sd)",
        "foreach v of varlist q1-q25 {",
        "    ttest `v', by(gender)",
        "}",
        "tabstat idx_*, by(gender) stat(n mean sd)",
        "foreach v of varlist idx_* {",
        "    ttest `v', by(gender)",
        "}",
        "pwcorr idx_* ideal_marriage_age, sig",
        "regress idx_umumiy ideal_marriage_age i.gender",
        "tabulate main_factor gender, chi2 column",
        "factor q1-q25, pcf",
        "rotate, varimax",
    ]:
        ws.append([line])
    style_sheet(ws)

    # ---------- 15 — CODEBOOK ----------
    ws = new_sheet("15_Codebook", ["O‘zgaruvchi", "Mazmuni", "Tip"])
    fixed = [("id", "Respondent identifikatori", "numeric"),
             ("telegram_id", "Telegram identifikatori", "numeric"),
             ("gender", "Jins", "categorical"),
             ("age_group", "Yosh guruhi", "categorical"),
             ("course", "Kurs", "categorical"),
             ("residence", "Doimiy yashash joyi", "categorical"),
             ("marital_status", "Oilaviy holat", "categorical"),
             ("ideal_marriage_age", "Nikohning maqbul yoshi", "numeric"),
             ("desired_children", "Istalgan farzandlar soni", "categorical"),
             ("main_factor", "Yakuniy asosiy omil", "categorical")]
    for item in fixed:
        ws.append(list(item))
    for i, (key, label) in enumerate(CRITERIA, start=1):
        ws.append([f"q{i}", label, "Likert 0–3"])
    for name, code, keys in CONSTRUCTS:
        ws.append([code, f"{name} indeksi (mezonlar o‘rtachasi, 0–3)", "numeric"])
    ws.append(["idx_umumiy", "Umumiy indeks (25 mezon o‘rtachasi, 0–3)", "numeric"])
    for i in range(1, 6):
        ws.append([f"top5_rank_{i}", f"TOP-5 dagi {i}-o‘rin", "text"])
    style_sheet(ws)

    # ---------- 16 — UMUMIY STATISTIKA ----------
    ws = new_sheet("16_Umumiy_Statistika", ["Ko‘rsatkich", "Qiymat"])
    for line in [("Jami respondentlar", total), ("Erkak respondentlar", len(male)),
                 ("Ayol respondentlar", len(female)),
                 ("25 ta mezonga to‘liq javob berganlar", sum(1 for p in people if p["complete"])),
                 ("Yakuniy omilni tanlaganlar", factor_total),
                 ("Test savollari soni", 25), ("TOP-5 mezonlar soni", 5),
                 ("Konstruktlar soni", len(CONSTRUCTS))]:
        ws.append(list(line))
    style_sheet(ws)

    # ---------- 17 — XOM MA’LUMOT ----------
    ws = new_sheet("17_Xom_SQL_Malumot", ["id", "telegram_id", "name", "gender", "created_at", "data"])
    for row in rows:
        ws.append([row["id"], row["telegram_id"], row["name"], row["gender"],
                   row["created_at"], repr(row["data"])])
    style_sheet(ws)

    wb.save(path)


# 12. ADMIN STATISTIKA

def admin_stats_text():
    rows = load_rows()
    total = len(rows)
    male = sum(1 for r in rows if r["gender"] == "Erkak")
    female = sum(1 for r in rows if r["gender"] == "Ayol")
    return ("ADMIN STATISTIKA\n\n"
            f"Jami respondent: {total}\n"
            f"Erkak: {male}\n"
            f"Ayol: {female}\n"
            "Test savollari: 25\n"
            "TOP-5: 5 ta tanlov")

# 13. MAIN

async def main():
    get_conn().close()

    # RENDER HEALTH CHECK
    app = web.Application()

    async def health(request):
        return web.Response(text="Bot ishlamoqda")
    app.router.add_get("/", health)
    runner = web.AppRunner(app)
    await runner.setup()
    port = int(os.environ.get("PORT", "8080"))
    site = web.TCPSite(runner, "0.0.0.0", port)
    await site.start()

    # TELEGRAM BOT
    bot = Bot(TOKEN)
    dp = Dispatcher()

    # START

    @dp.message(CommandStart())
    async def start(message: Message, state: FSMContext):
        if (message.from_user.id != ADMIN_ID and has_completed_response(message.from_user.id)):
            await state.clear()
            await message.answer("Siz ushbu so‘rovnomada " "avval qatnashgansiz.\n\n"
                "Ishtirokingiz uchun rahmat!")
            return
        await start_survey(message, state)

    # ADMIN

    @dp.message(Command("admin"))
    async def admin_panel(message: Message):
        if (message.from_user.id != ADMIN_ID):
            return
        await message.answer("ADMIN PANEL\n\n" "Kerakli amalni tanlang:", reply_markup=admin_keyboard(
            ))

    @dp.callback_query(F.data == "admin:stats")
    async def admin_stats(call: CallbackQuery):
        if (call.from_user.id != ADMIN_ID):
            await call.answer("Ruxsat yo‘q", show_alert=True)
            return
        await call.answer()
        await call.message.answer(admin_stats_text())

    @dp.callback_query(F.data == "admin:export")
    async def admin_export(call: CallbackQuery):
        if (call.from_user.id != ADMIN_ID):
            await call.answer("Ruxsat yo‘q", show_alert=True)
            return
        await call.answer("Excel tayyorlanmoqda...")
        try:
            export_excel(EXPORT_PATH)
            await call.message.answer_document(document=FSInputFile(EXPORT_PATH), caption=(
                "Excel tayyor.\n" "25 ta test, TOP-5, gender t-test, konstruktlar, korrelyatsiya, Stata "
                "va Codebook kiritilgan."))
        except Exception as e:
            await call.message.answer("Excel eksportida xatolik:\n" f"{type(e).__name__}: {e}")

    @dp.callback_query(F.data == "admin:reset")
    async def admin_reset(call: CallbackQuery):
        if (call.from_user.id != ADMIN_ID):
            await call.answer("Ruxsat yo‘q", show_alert=True)
            return
        delete_user_response(call.from_user.id)
        await call.answer("Test yozuvingiz o‘chirildi.")
        await call.message.answer("Mening testim o‘chirildi. "
            "Endi qayta boshlashingiz mumkin.")

    @dp.callback_query(F.data == "admin:reset_all")
    async def admin_reset_all(call: CallbackQuery):
        if (call.from_user.id != ADMIN_ID):
            await call.answer("Ruxsat yo‘q", show_alert=True)
            return
        delete_all_responses()
        await call.answer("Barcha ma’lumot o‘chirildi.", show_alert=True)
        await call.message.answer("Barcha respondent ma’lumotlari o‘chirildi.")

    @dp.callback_query(F.data == "admin:start")
    async def admin_start(call: CallbackQuery, state: FSMContext):
        if (call.from_user.id != ADMIN_ID):
            await call.answer("Ruxsat yo‘q", show_alert=True)
            return
        await call.answer()
        await start_survey(call.message, state)

    # ISM

    @dp.callback_query(Survey.name, F.data == "skip_name:0")
    async def skip_name(call: CallbackQuery, state: FSMContext):
        await state.update_data(name="")
        await state.set_state(Survey.gender)
        await call.answer()
        await call.message.edit_reply_markup(reply_markup=None)
        await call.message.answer("Jinsingizni tanlang:", reply_markup=gender_buttons())

    @dp.message(Survey.name)
    async def get_name(message: Message, state: FSMContext):
        name = (message.text or "").strip()
        if not name:
            await message.answer("Ismingizni yozing yoki "
                "«O‘tkazib yuborish» tugmasini bosing.")
            return
        await state.update_data(name=name)
        await state.set_state(Survey.gender)
        await message.answer("Jinsingizni tanlang:", reply_markup=gender_buttons())

    # JINS

    @dp.callback_query(Survey.gender, F.data.startswith("gender:"))
    async def get_gender(call: CallbackQuery, state: FSMContext):
        gender = call.data.split(":", 1)[1]
        await state.update_data(gender=gender, general_index=0, general_answers={},
            likert_answers={}, top5_selected=[], top5_ranks=[], sd_answers={}, sd_score=0)
        await call.answer()
        await call.message.edit_reply_markup(reply_markup=None)
        await send_general(call.message, state, 0)

    # DEMOGRAFIYA

    @dp.callback_query(Survey.general, F.data.startswith("general:"))
    async def get_general(call: CallbackQuery, state: FSMContext):
        data = await state.get_data()
        question_index = data.get("general_index", 0)
        option_index = int(call.data.split(":")[1])
        key, question, options = GENERAL_QUESTIONS[question_index]
        selected_answer = options[option_index]
        general_answers = data.get("general_answers", {})
        general_answers[key] = selected_answer
        next_index = question_index + 1
        await state.update_data(general_answers=general_answers, general_index=next_index)
        await call.answer()
        await call.message.edit_reply_markup(reply_markup=None)
        await send_general(call.message, state, next_index)

    # NIKOH YOSHI

    @dp.message(Survey.marriage_age)
    async def get_marriage_age(message: Message, state: FSMContext):
        value = (message.text or "").strip()
        try:
            age = int(value)
            if (age < 15 or age > 80):
                raise ValueError
        except ValueError:
            await message.answer("Iltimos, 15–80 oralig‘idagi " "yoshni raqam bilan kiriting.\n"
                "Masalan: 25")
            return
        await state.update_data(ideal_marriage_age=age)
        await send_likert(message, state, 0)

    # 25 TA TEST

    @dp.callback_query(Survey.likert, F.data.startswith("likert:"))
    async def get_likert(call: CallbackQuery, state: FSMContext):
        data = await state.get_data()
        index = data.get("likert_index", 0)
        value = int(call.data.split(":")[1])
        key, label = CRITERIA[index]
        answers = data.get("likert_answers", {})
        answers[key] = value
        await state.update_data(likert_answers=answers, likert_index=index + 1)
        await call.answer()
        await call.message.edit_reply_markup(reply_markup=None)
        await send_likert(call.message, state, index + 1)

    # TOP-5

    @dp.callback_query(Survey.top5, F.data.startswith("top5:"))
    async def get_top5(call: CallbackQuery, state: FSMContext):
        action = call.data.split(":", 1)[1]
        data = await state.get_data()
        selected = list(data.get("top5_selected", []))
        if action == "count":
            await call.answer(f"Tanlangan: " f"{len(selected)}/5", show_alert=True)
            return
        if action == "confirm":
            if len(selected) != 5:
                await call.answer("Avval 5 ta mezon tanlang.", show_alert=True)
                return
            ranks = [CRITERIA[i][1] for i in selected]
            await state.update_data(top5_ranks=ranks)
            await call.answer()
            await call.message.edit_reply_markup(reply_markup=None)
            await send_sd(call.message, state, 0)
            return
        try:
            idx = int(action)
        except ValueError:
            await call.answer()
            return
        if idx in selected:
            selected.remove(idx)
        else:
            if len(selected) >= 5:
                await call.answer("Faqat 5 ta mezon " "tanlash mumkin.", show_alert=True)
                return
            selected.append(idx)
        await state.update_data(top5_selected=selected)
        await call.answer()
        await call.message.edit_reply_markup(reply_markup=top5_keyboard(selected))

    # SD

    @dp.callback_query(Survey.sd_scale, F.data.startswith("sd:"))
    async def get_sd(call: CallbackQuery, state: FSMContext):
        data = await state.get_data()
        index = data.get("sd_index", 0)
        answer = int(call.data.split(":")[1])
        key, question, expected = SD_QUESTIONS[index]
        sd_answers = data.get("sd_answers", {})
        sd_answers[key] = answer
        score = data.get("sd_score", 0)
        if (bool(answer) == expected):
            score += 1
        await state.update_data(sd_answers=sd_answers, sd_score=score, sd_index=index + 1)
        await call.answer()
        await call.message.edit_reply_markup(reply_markup=None)
        await send_sd(call.message, state, index + 1)

    # YAKUNIY OMIL

    @dp.callback_query(Survey.main_factor, F.data.startswith("factor:"))
    async def get_factor(call: CallbackQuery, state: FSMContext):
        idx = int(call.data.split(":")[1])
        factor = MAIN_FACTORS[idx]
        await call.answer()
        await call.message.edit_reply_markup(reply_markup=None)
        if factor == "Boshqa":
            await state.set_state(Survey.custom_factor)
            await call.message.answer("Iltimos, siz uchun " "boshqa eng muhim " "omilni yozing:")
            return
        await state.update_data(main_factor=factor, custom_factor="")
        await finish_survey(call.message, state)

    @dp.message(Survey.custom_factor)
    async def get_custom_factor(message: Message, state: FSMContext):
        custom = (message.text or "").strip()
        if not custom:
            await message.answer("Iltimos, omilni yozib yuboring.")
            return
        await state.update_data(main_factor="Boshqa", custom_factor=custom)
        await finish_survey(message, state)

    # YAKUNLASH

    async def finish_survey(message, state):
        data = await state.get_data()
        sd_score = data.get("sd_score", 0)
        final_data = {"age_group": data.get("general_answers", {}).get("age_group", ""),
            "course": data.get("general_answers", {}).get("course", ""), "residence": data.get(
            "general_answers", {}).get("residence", ""), "marital_status": data.get(
            "general_answers", {}).get("marital_status", ""), "desired_children": data.get(
            "general_answers", {}).get("desired_children", ""), "ideal_marriage_age": data.get(
            "ideal_marriage_age", ""), "likert_answers": data.get("likert_answers", {}),
            "top5_ranks": data.get("top5_ranks", []), "sd_answers": data.get("sd_answers", {}),
            "sd_score": sd_score, "sd_level": calculate_sd_level(sd_score), "main_factor": data.get(
            "main_factor", ""), "custom_factor": data.get("custom_factor", "")}
        save_response(message.from_user.id, data.get("name", ""), data.get("gender", ""),
            final_data)
        await state.clear()
        await message.answer("So‘rovnoma yakunlandi!\n\n" "Ishtirokingiz uchun katta rahmat.\n"
            "Sizning javoblaringiz " "ilmiy-statistik tahlil uchun " "saqlab qo‘yildi.")

    # POLLING
    try:
        await dp.start_polling(bot)
    finally:
        await bot.session.close()
        await runner.cleanup()

# START
if __name__ == "__main__":
    import asyncio
    asyncio.run(main())
