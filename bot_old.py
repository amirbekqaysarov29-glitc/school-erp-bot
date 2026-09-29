import asyncio
import os
import random
import sqlite3
import shutil
import threading
import html
from pathlib import Path
from datetime import date, datetime, timedelta

import openpyxl
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas

from aiogram import Bot, Dispatcher, F
from aiogram.filters import CommandStart
from aiogram.types import Message, CallbackQuery, InlineKeyboardMarkup, InlineKeyboardButton, FSInputFile

from fastapi import FastAPI, Form, Request
from fastapi.responses import HTMLResponse, RedirectResponse, FileResponse
import uvicorn

try:
    from dotenv import load_dotenv
    load_dotenv()
except Exception:
    pass


# ================= CONFIG =================

BOT_TOKEN = os.getenv("BOT_TOKEN", "").strip()
ADMIN_CODE = os.getenv("ADMIN_CODE", "ADMIN2026")
TEACHER_CODE = os.getenv("TEACHER_CODE", "TEACHER2026")

DB_NAME = os.getenv("DB_NAME", "data/davomat.db")

# Local: http://127.0.0.1:8000
# Deploy: set WEB_HOST=0.0.0.0
WEB_HOST = os.getenv("WEB_HOST", "0.0.0.0")
WEB_PORT = int(os.getenv("WEB_PORT", "8000"))

Path("data").mkdir(exist_ok=True)
Path("exports/excel").mkdir(parents=True, exist_ok=True)
Path("exports/pdf").mkdir(parents=True, exist_ok=True)
Path("backups").mkdir(exist_ok=True)

web_app = FastAPI(title="School ERP Dashboard")


# ================= DATABASE =================

def db():
    conn = sqlite3.connect(DB_NAME)
    conn.row_factory = sqlite3.Row
    return conn


def init_db():
    conn = db()
    c = conn.cursor()

    c.execute("""
    CREATE TABLE IF NOT EXISTS users (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        telegram_id INTEGER UNIQUE,
        role TEXT,
        full_name TEXT,
        created_at TEXT
    )
    """)
    

    c.execute("""
    CREATE TABLE IF NOT EXISTS students (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        full_name TEXT,
        class_name TEXT,
        parent_phone TEXT,
        parent_telegram_id INTEGER,
        parent_code TEXT UNIQUE,
        created_at TEXT
    )
    """)

    c.execute("""
    CREATE TABLE IF NOT EXISTS teachers (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        full_name TEXT,
        telegram_id INTEGER UNIQUE,
        teacher_code TEXT UNIQUE,
        created_at TEXT
    )
    """)

    c.execute("""
    CREATE TABLE IF NOT EXISTS teacher_classes (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        teacher_id INTEGER,
        class_name TEXT
    )
    """)

    c.execute("""
    CREATE TABLE IF NOT EXISTS subjects (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        name TEXT UNIQUE,
        created_at TEXT
    )
    """)

    c.execute("""
    CREATE TABLE IF NOT EXISTS attendance (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        student_id INTEGER,
        full_name TEXT,
        class_name TEXT,
        subject TEXT,
        date TEXT,
        status TEXT,
        teacher_telegram_id INTEGER,
        created_at TEXT
    )
    """)

    # Default subjects. Teacher can add more.
    for subject_name in ["Matematika", "Ona tili", "Ingliz tili"]:
        c.execute(
            "INSERT OR IGNORE INTO subjects(name, created_at) VALUES (?, ?)",
            (subject_name, datetime.now().isoformat())
        )

    conn.commit()
    conn.close()


def unique_code(cur, table, column, prefix):
    while True:
        code = f"{prefix}-{random.randint(1000, 9999)}"
        cur.execute(f"SELECT id FROM {table} WHERE {column} = ?", (code,))
        if not cur.fetchone():
            return code


def set_role(telegram_id, user_role, full_name=None):
    conn = db()
    c = conn.cursor()

    c.execute("""
    INSERT INTO users (telegram_id, role, full_name, created_at)
    VALUES (?, ?, ?, ?)
    ON CONFLICT(telegram_id)
    DO UPDATE SET
        role = excluded.role,
        full_name = COALESCE(excluded.full_name, users.full_name)
    """, (
        telegram_id,
        user_role,
        full_name,
        datetime.now().isoformat()
    ))

    conn.commit()
    conn.close()


def role(telegram_id):
    conn = db()
    c = conn.cursor()

    c.execute("SELECT role FROM users WHERE telegram_id = ?", (telegram_id,))
    row = c.fetchone()

    conn.close()
    return row["role"] if row else None


def ids_by_role(user_role):
    conn = db()
    c = conn.cursor()

    c.execute("SELECT telegram_id FROM users WHERE role = ?", (user_role,))
    rows = c.fetchall()

    conn.close()
    return [row["telegram_id"] for row in rows]


def add_student(full_name, class_name, phone):
    conn = db()
    c = conn.cursor()

    code = unique_code(c, "students", "parent_code", "PAR")

    c.execute("""
    INSERT INTO students (full_name, class_name, parent_phone, parent_code, created_at)
    VALUES (?, ?, ?, ?, ?)
    """, (
        full_name.strip(),
        class_name.strip(),
        phone.strip(),
        code,
        datetime.now().isoformat()
    ))

    conn.commit()
    conn.close()
    return code


def students():
    conn = db()
    c = conn.cursor()

    c.execute("""
    SELECT *
    FROM students
    ORDER BY class_name, full_name
    """)

    rows = c.fetchall()
    conn.close()
    return rows


def student(student_id):
    conn = db()
    c = conn.cursor()

    c.execute("SELECT * FROM students WHERE id = ?", (student_id,))
    row = c.fetchone()

    conn.close()
    return row


def search_students(query):
    conn = db()
    c = conn.cursor()

    like = f"%{query.strip()}%"

    c.execute("""
    SELECT *
    FROM students
    WHERE full_name LIKE ?
       OR class_name LIKE ?
       OR parent_phone LIKE ?
       OR parent_code LIKE ?
    ORDER BY class_name, full_name
    """, (like, like, like, like))

    rows = c.fetchall()
    conn.close()
    return rows


def delete_student(student_id):
    conn = db()
    c = conn.cursor()

    c.execute("DELETE FROM students WHERE id = ?", (student_id,))
    count = c.rowcount

    conn.commit()
    conn.close()
    return count


def classes():
    conn = db()
    c = conn.cursor()

    c.execute("""
    SELECT DISTINCT class_name
    FROM students
    ORDER BY class_name
    """)

    rows = c.fetchall()
    conn.close()
    return [row["class_name"] for row in rows]


def teacher_classes(telegram_id):
    conn = db()
    c = conn.cursor()

    c.execute("SELECT id FROM teachers WHERE telegram_id = ?", (telegram_id,))
    teacher = c.fetchone()

    if not teacher:
        conn.close()
        return []

    c.execute("""
    SELECT class_name
    FROM teacher_classes
    WHERE teacher_id = ?
    ORDER BY class_name
    """, (teacher["id"],))

    rows = c.fetchall()
    conn.close()

    return [row["class_name"] for row in rows]


def classes_for_teacher(telegram_id):
    allowed = teacher_classes(telegram_id)
    return allowed if allowed else classes()


def students_by_class(class_name):
    conn = db()
    c = conn.cursor()

    c.execute("""
    SELECT id, full_name, class_name
    FROM students
    WHERE class_name = ?
    ORDER BY full_name
    """, (class_name,))

    rows = c.fetchall()
    conn.close()
    return rows


def link_parent(code, telegram_id, full_name=None):
    conn = db()
    c = conn.cursor()

    c.execute("""
    UPDATE students
    SET parent_telegram_id = ?
    WHERE UPPER(parent_code) = ?
    """, (
        telegram_id,
        code.strip().upper()
    ))

    count = c.rowcount
    conn.commit()
    conn.close()

    if count:
        set_role(telegram_id, "parent", full_name)

    return count


def parent_children(telegram_id):
    conn = db()
    c = conn.cursor()

    c.execute("""
    SELECT *
    FROM students
    WHERE parent_telegram_id = ?
    ORDER BY class_name, full_name
    """, (telegram_id,))

    rows = c.fetchall()
    conn.close()
    return rows


def add_subject(name):
    conn = db()
    c = conn.cursor()

    clean_name = name.strip()
    if clean_name:
        c.execute("""
        INSERT OR IGNORE INTO subjects (name, created_at)
        VALUES (?, ?)
        """, (
            clean_name,
            datetime.now().isoformat()
        ))

    conn.commit()
    conn.close()


def subjects():
    conn = db()
    c = conn.cursor()

    c.execute("SELECT name FROM subjects ORDER BY name")
    rows = c.fetchall()

    conn.close()
    return [row["name"] for row in rows]


def add_teacher(name, classes_text=""):
    conn = db()
    c = conn.cursor()

    code = unique_code(c, "teachers", "teacher_code", "TCH")

    c.execute("""
    INSERT INTO teachers (full_name, teacher_code, created_at)
    VALUES (?, ?, ?)
    """, (
        name.strip(),
        code,
        datetime.now().isoformat()
    ))

    teacher_id = c.lastrowid

    for class_name in [item.strip() for item in classes_text.split(",") if item.strip()]:
        c.execute("""
        INSERT INTO teacher_classes (teacher_id, class_name)
        VALUES (?, ?)
        """, (
            teacher_id,
            class_name
        ))

    conn.commit()
    conn.close()
    return code


def link_teacher(code, telegram_id, full_name=None):
    conn = db()
    c = conn.cursor()

    c.execute("""
    UPDATE teachers
    SET telegram_id = ?
    WHERE UPPER(teacher_code) = ?
    """, (
        telegram_id,
        code.strip().upper()
    ))

    count = c.rowcount
    conn.commit()
    conn.close()

    if count:
        set_role(telegram_id, "teacher", full_name)

    return count


def teachers():
    conn = db()
    c = conn.cursor()

    c.execute("SELECT * FROM teachers ORDER BY full_name")
    rows = c.fetchall()

    conn.close()
    return rows


def save_attendance(student_id, full_name, class_name, subject_name, status, teacher_telegram_id):
    conn = db()
    c = conn.cursor()

    today = str(date.today())

    c.execute("""
    INSERT INTO attendance (
        student_id,
        full_name,
        class_name,
        subject,
        date,
        status,
        teacher_telegram_id,
        created_at
    )
    VALUES (?, ?, ?, ?, ?, ?, ?, ?)
    """, (
        student_id,
        full_name,
        class_name,
        subject_name,
        today,
        status,
        teacher_telegram_id,
        datetime.now().isoformat()
    ))

    conn.commit()
    conn.close()


def attendance_by_date(target_date):
    conn = db()
    c = conn.cursor()

    c.execute("""
    SELECT *
    FROM attendance
    WHERE date = ?
    ORDER BY class_name, subject, full_name
    """, (target_date,))

    rows = c.fetchall()
    conn.close()
    return rows


def stats():
    conn = db()
    c = conn.cursor()

    queries = {
        "students": "SELECT COUNT(*) AS count FROM students",
        "teachers": "SELECT COUNT(*) AS count FROM teachers",
        "subjects": "SELECT COUNT(*) AS count FROM subjects",
        "present": "SELECT COUNT(*) AS count FROM attendance WHERE status='present'",
        "absent": "SELECT COUNT(*) AS count FROM attendance WHERE status='absent'",
        "late": "SELECT COUNT(*) AS count FROM attendance WHERE status='late'",
    }

    result = {}

    for key, query in queries.items():
        c.execute(query)
        result[key] = c.fetchone()["count"]

    conn.close()
    return result


def top_absent(limit=10):
    conn = db()
    c = conn.cursor()

    c.execute("""
    SELECT student_id, full_name, COUNT(*) AS total
    FROM attendance
    WHERE status = 'absent'
    GROUP BY student_id, full_name
    ORDER BY total DESC
    LIMIT ?
    """, (limit,))

    rows = c.fetchall()
    conn.close()
    return rows


def subject_stats():
    conn = db()
    c = conn.cursor()

    c.execute("""
    SELECT
        subject,
        SUM(CASE WHEN status = 'present' THEN 1 ELSE 0 END) AS present,
        SUM(CASE WHEN status = 'absent' THEN 1 ELSE 0 END) AS absent,
        SUM(CASE WHEN status = 'late' THEN 1 ELSE 0 END) AS late
    FROM attendance
    GROUP BY subject
    ORDER BY absent DESC
    """)

    rows = c.fetchall()
    conn.close()
    return rows


def monthly(year_month):
    conn = db()
    c = conn.cursor()

    c.execute("""
    SELECT
        SUM(CASE WHEN status = 'present' THEN 1 ELSE 0 END) AS present,
        SUM(CASE WHEN status = 'absent' THEN 1 ELSE 0 END) AS absent,
        SUM(CASE WHEN status = 'late' THEN 1 ELSE 0 END) AS late
    FROM attendance
    WHERE date LIKE ?
    """, (f"{year_month}%",))

    row = c.fetchone()
    conn.close()

    present = row["present"] or 0
    absent = row["absent"] or 0
    late = row["late"] or 0
    total = present + absent + late
    percent = round((present / total) * 100, 1) if total else 0

    return present, absent, late, total, percent


def student_summary(student_id):
    conn = db()
    c = conn.cursor()

    c.execute("""
    SELECT
        SUM(CASE WHEN status = 'present' THEN 1 ELSE 0 END) AS present,
        SUM(CASE WHEN status = 'absent' THEN 1 ELSE 0 END) AS absent,
        SUM(CASE WHEN status = 'late' THEN 1 ELSE 0 END) AS late,
        COUNT(*) AS total
    FROM attendance
    WHERE student_id = ?
    """, (student_id,))

    summary = c.fetchone()

    c.execute("""
    SELECT date, subject, status
    FROM attendance
    WHERE student_id = ?
    ORDER BY date DESC, id DESC
    LIMIT 8
    """, (student_id,))

    records = c.fetchall()
    conn.close()

    present = summary["present"] or 0
    absent = summary["absent"] or 0
    late = summary["late"] or 0
    total = summary["total"] or 0
    percent = round((present / total) * 100, 1) if total else 0

    return present, absent, late, total, percent, records


def consecutive_absent(student_id, days=3):
    dates = [str(date.today() - timedelta(days=i)) for i in range(days)]
    placeholders = ",".join(["?"] * len(dates))

    conn = db()
    c = conn.cursor()

    c.execute(f"""
    SELECT COUNT(DISTINCT date) AS count
    FROM attendance
    WHERE student_id = ?
      AND status = 'absent'
      AND date IN ({placeholders})
    """, [student_id] + dates)

    count = c.fetchone()["count"]
    conn.close()

    return count >= days


def parent_ids():
    conn = db()
    c = conn.cursor()

    c.execute("""
    SELECT DISTINCT parent_telegram_id
    FROM students
    WHERE parent_telegram_id IS NOT NULL
    """)

    rows = c.fetchall()
    conn.close()

    return [row["parent_telegram_id"] for row in rows]


# ================= KEYBOARDS =================

def button(text, data):
    return InlineKeyboardButton(text=text, callback_data=data)


def keyboard(rows):
    return InlineKeyboardMarkup(inline_keyboard=rows)


def login_keyboard():
    return keyboard([
        [button("👨‍💼 Admin", "login:admin")],
        [button("👨‍🏫 O‘qituvchi", "login:teacher")],
        [button("👨‍👩‍👧 Ota-ona", "login:parent")],
    ])


def admin_keyboard():
    return keyboard([
        [button("👥 O‘quvchilar", "students")],
        [button("👨‍🏫 Teacherlar", "teachers")],
        [button("📘 Fanlar", "subjects")],
        [button("📊 Hisobotlar", "reports")],
        [button("📈 Analytics", "analytics")],
        [button("📤 Excel export", "excel")],
        [button("📄 PDF report", "pdf")],
        [button("📢 Broadcast", "broadcast")],
        [button("💾 Backup", "backup")],
    ])


def teacher_keyboard():
    return keyboard([
        [button("📚 Davomat olish", "attendance")],
        [button("➕ O‘quvchi qo‘shish", "add_student")],
        [button("👥 O‘quvchilar", "students")],
        [button("📘 Fanlar", "subjects")],
        [button("📊 Hisobotlar", "reports")],
        [button("📈 Analytics", "analytics")],
        [button("📤 Excel export", "excel")],
        [button("📄 PDF report", "pdf")],
        [button("🔍 O‘quvchi qidirish", "search")],
    ])


def parent_keyboard():
    return keyboard([
        [button("👶 Farzandlarim", "children")],
    ])


def menu_keyboard(user_role):
    if user_role == "admin":
        return admin_keyboard()

    if user_role == "teacher":
        return teacher_keyboard()

    if user_role == "parent":
        return parent_keyboard()

    return login_keyboard()


def status_text(status):
    return {
        "present": "✅ Keldi",
        "absent": "❌ Kelmadi",
        "late": "⏰ Kechikdi",
    }.get(status, status)


def status_icon(status):
    return {
        "present": "✅",
        "absent": "❌",
        "late": "⏰",
    }.get(status, "➖")


# ================= BOT =================

bot = Bot(BOT_TOKEN) if BOT_TOKEN else None
dp = Dispatcher()

STATE = {}
SESSION = {}
STATUS_ORDER = ["present", "absent", "late"]


def set_state(user_id, action, data=None):
    STATE[user_id] = {
        "action": action,
        "data": data or {}
    }


def pop_state(user_id):
    STATE.pop(user_id, None)


async def send_menu(message, user_role=None, text=None):
    user_role = user_role or role(message.from_user.id)

    await message.answer(
        text or ("Menyu:" if user_role else "🔐 Rol tanlang:"),
        reply_markup=menu_keyboard(user_role)
    )


async def notify_admins(text):
    if not bot:
        return

    for admin_id in ids_by_role("admin"):
        try:
            await bot.send_message(admin_id, text)
        except Exception as error:
            print("admin notify:", error)


async def notify_absence(student_id, full_name, class_name, subject_name, today):
    if not bot:
        return

    current_student = student(student_id)

    if current_student and current_student["parent_telegram_id"]:
        try:
            await bot.send_message(
                current_student["parent_telegram_id"],
                (
                    "❌ Davomat bildirishnomasi\n\n"
                    f"Farzandingiz {full_name} bugun darsga kelmadi.\n\n"
                    f"🏫 Sinf: {class_name}\n"
                    f"📘 Fan: {subject_name}\n"
                    f"📅 Sana: {today}"
                )
            )
        except Exception as error:
            print("parent notify:", error)

    await notify_admins(
        (
            "❌ Kelmadi\n\n"
            f"👤 {full_name}\n"
            f"🏫 {class_name}\n"
            f"📘 {subject_name}\n"
            f"📅 {today}"
        )
    )

    if consecutive_absent(student_id, 3):
        alert = (
            "⚠️ 3 kunlik ogohlantirish\n\n"
            f"👤 {full_name}\n"
            f"🏫 {class_name}\n"
            "❌ Ketma-ket kelmagan holat bor."
        )

        await notify_admins(alert)

        if current_student and current_student["parent_telegram_id"]:
            try:
                await bot.send_message(current_student["parent_telegram_id"], alert)
            except Exception:
                pass


def profile_text(student_id):
    current_student = student(student_id)

    if not current_student:
        return "❌ O‘quvchi topilmadi."

    present, absent, late, total, percent, records = student_summary(student_id)

    text = (
        "👤 O‘quvchi profili\n\n"
        f"🆔 ID: {current_student['id']}\n"
        f"👤 Ism: {current_student['full_name']}\n"
        f"🏫 Sinf: {current_student['class_name']}\n"
        f"📞 Telefon: {current_student['parent_phone'] or '-'}\n"
        f"🔐 Parent kodi: {current_student['parent_code']}\n\n"
        "📊 Davomat\n"
        f"✅ Keldi: {present}\n"
        f"❌ Kelmadi: {absent}\n"
        f"⏰ Kechikdi: {late}\n"
        f"📈 Foiz: {percent}%\n"
    )

    if records:
        text += "\nOxirgi yozuvlar:\n"
        for record in records:
            text += f"{record['date']} | {record['subject']} | {status_text(record['status'])}\n"

    return text


@dp.message(CommandStart())
async def start(message: Message):
    current_role = role(message.from_user.id)

    if current_role:
        await send_menu(message, current_role)
    else:
        await message.answer(
            "🔐 Botga kirish uchun rolingizni tanlang:",
            reply_markup=login_keyboard()
        )


@dp.callback_query(F.data.startswith("login:"))
async def login_choose(callback: CallbackQuery):
    selected_role = callback.data.split(":")[1]

    set_state(
        callback.from_user.id,
        "login",
        {"role": selected_role}
    )

    if selected_role == "admin":
        await callback.message.answer("👨‍💼 Admin parolini kiriting:")

    elif selected_role == "teacher":
        await callback.message.answer("👨‍🏫 O‘qituvchi paroli yoki TCH-kodni kiriting:")

    else:
        await callback.message.answer("👨‍👩‍👧 Ota-ona PAR-kodni kiriting:")

    await callback.answer()


@dp.message()
async def text_handler(message: Message):
    user_id = message.from_user.id
    current_state = STATE.get(user_id)
    message_text = (message.text or "").strip()

    if not current_state:
        await message.answer("Buyruq tushunarsiz. /start ni bosing.")
        return

    action = current_state["action"]
    data = current_state["data"]

    if action == "login":
        selected_role = data["role"]

        if selected_role == "admin":
            if message_text != ADMIN_CODE:
                await message.answer("❌ Admin paroli xato.")
                return

            set_role(user_id, "admin", message.from_user.full_name)
            pop_state(user_id)
            await send_menu(message, "admin", "✅ Admin sifatida kirdingiz.")
            return

        if selected_role == "teacher":
            if message_text == TEACHER_CODE:
                set_role(user_id, "teacher", message.from_user.full_name)
                pop_state(user_id)
                await send_menu(message, "teacher", "✅ Teacher sifatida kirdingiz.")
                return

            if message_text.upper().startswith("TCH-") and link_teacher(
                message_text,
                user_id,
                message.from_user.full_name
            ):
                pop_state(user_id)
                await send_menu(message, "teacher", "✅ Teacher kodingiz tasdiqlandi.")
                return

            await message.answer("❌ Teacher paroli/kodi xato.")
            return

        if selected_role == "parent":
            if link_parent(message_text, user_id, message.from_user.full_name):
                pop_state(user_id)
                await send_menu(message, "parent", "✅ Parent sifatida kirdingiz.")
                return

            await message.answer("❌ Parent kodi xato.")
            return

    if action == "add_student":
        if role(user_id) != "teacher":
            pop_state(user_id)
            await message.answer("❌ Bu bo‘lim faqat teacher uchun.")
            return

        if message_text.count(",") < 2:
            await message.answer("Format: Ali Karimov, 10-A, +998901234567")
            return

        full_name, class_name, phone = [
            item.strip()
            for item in message_text.split(",", 2)
        ]

        parent_code = add_student(full_name, class_name, phone)
        pop_state(user_id)

        await message.answer(
            (
                "✅ O‘quvchi qo‘shildi\n\n"
                f"👤 {full_name}\n"
                f"🏫 {class_name}\n"
                f"📞 {phone}\n"
                f"🔐 Parent kodi: {parent_code}"
            ),
            reply_markup=teacher_keyboard()
        )
        return

    if action == "search":
        rows = search_students(message_text)
        pop_state(user_id)
        current_role = role(user_id)

        if not rows:
            await message.answer("❌ Topilmadi.", reply_markup=menu_keyboard(current_role))
            return

        text = "🔍 Natijalar:\n\n"
        buttons = []

        for row in rows:
            text += (
                f"🆔 {row['id']} | "
                f"{row['full_name']} | "
                f"{row['class_name']} | "
                f"{row['parent_code']}\n"
            )

            buttons.append([
                button(f"👤 {row['full_name']}", f"profile:{row['id']}")
            ])

        buttons.append([button("⬅️ Menyu", "menu")])

        await message.answer(text, reply_markup=keyboard(buttons))
        return

    if action == "subject":
        add_subject(message_text)
        pop_state(user_id)

        await message.answer(
            f"✅ Fan qo‘shildi: {message_text}",
            reply_markup=menu_keyboard(role(user_id))
        )
        return

    if action == "teacher":
        if role(user_id) != "admin":
            pop_state(user_id)
            await message.answer("❌ Bu bo‘lim faqat admin uchun.")
            return

        parts = [item.strip() for item in message_text.split(",") if item.strip()]
        teacher_name = parts[0]
        class_names = ",".join(parts[1:]) if len(parts) > 1 else ""

        teacher_code = add_teacher(teacher_name, class_names)
        pop_state(user_id)

        await message.answer(
            (
                "✅ Teacher qo‘shildi\n"
                f"👨‍🏫 {teacher_name}\n"
                f"🔐 Kod: {teacher_code}\n"
                f"🏫 Sinflar: {', '.join(parts[1:]) or 'hammasi'}"
            ),
            reply_markup=admin_keyboard()
        )
        return

    if action == "broadcast":
        if role(user_id) != "admin":
            pop_state(user_id)
            await message.answer("❌ Bu bo‘lim faqat admin uchun.")
            return

        success = 0

        for parent_id in parent_ids():
            try:
                await bot.send_message(
                    parent_id,
                    f"📢 Maktab xabari\n\n{message_text}"
                )
                success += 1
            except Exception as error:
                print("broadcast:", error)

        pop_state(user_id)

        await message.answer(
            f"✅ Xabar yuborildi: {success}",
            reply_markup=admin_keyboard()
        )
        return


@dp.callback_query(F.data == "menu")
async def menu(callback: CallbackQuery):
    await callback.message.answer(
        "Menyu:",
        reply_markup=menu_keyboard(role(callback.from_user.id))
    )
    await callback.answer()


@dp.callback_query(F.data == "add_student")
async def callback_add_student(callback: CallbackQuery):
    if role(callback.from_user.id) != "teacher":
        await callback.answer("Faqat teacher.", show_alert=True)
        return

    set_state(callback.from_user.id, "add_student")
    await callback.message.answer("O‘quvchi: Ali Karimov, 10-A, +998901234567")
    await callback.answer()


@dp.callback_query(F.data == "students")
async def callback_students(callback: CallbackQuery):
    if role(callback.from_user.id) not in ["admin", "teacher"]:
        await callback.answer("Ruxsat yo‘q.", show_alert=True)
        return

    rows = students()
    text = "👥 O‘quvchilar:\n\n"
    buttons = []

    for row in rows:
        text += (
            f"🆔 {row['id']} | "
            f"{row['full_name']} | "
            f"{row['class_name']} | "
            f"Kod: {row['parent_code']}\n"
        )

        buttons.append([
            button(f"👤 {row['full_name']}", f"profile:{row['id']}")
        ])

    buttons.append([button("🗑 O‘chirish", "delete_menu")])
    buttons.append([button("⬅️ Menyu", "menu")])

    await callback.message.answer(
        text if rows else "O‘quvchi yo‘q.",
        reply_markup=keyboard(buttons)
    )

    await callback.answer()


@dp.callback_query(F.data == "search")
async def callback_search(callback: CallbackQuery):
    set_state(callback.from_user.id, "search")
    await callback.message.answer("🔍 Ism, sinf, telefon yoki kod kiriting:")
    await callback.answer()


@dp.callback_query(F.data.startswith("profile:"))
async def callback_profile(callback: CallbackQuery):
    student_id = int(callback.data.split(":")[1])
    current_role = role(callback.from_user.id)

    if current_role == "parent":
        child_ids = [child["id"] for child in parent_children(callback.from_user.id)]

        if student_id not in child_ids:
            await callback.answer("Bu profil sizniki emas.", show_alert=True)
            return

    await callback.message.answer(
        profile_text(student_id),
        reply_markup=menu_keyboard(current_role)
    )

    await callback.answer()


@dp.callback_query(F.data == "children")
async def callback_children(callback: CallbackQuery):
    rows = parent_children(callback.from_user.id)
    buttons = []

    for row in rows:
        buttons.append([
            button(
                f"👤 {row['full_name']} — {row['class_name']}",
                f"profile:{row['id']}"
            )
        ])

    buttons.append([button("⬅️ Menyu", "menu")])

    await callback.message.answer(
        "Farzandlaringiz:" if rows else "Farzand topilmadi.",
        reply_markup=keyboard(buttons)
    )

    await callback.answer()


@dp.callback_query(F.data == "delete_menu")
async def callback_delete_menu(callback: CallbackQuery):
    if role(callback.from_user.id) not in ["admin", "teacher"]:
        await callback.answer("Ruxsat yo‘q.", show_alert=True)
        return

    buttons = []

    for row in students():
        buttons.append([
            button(
                f"🗑 {row['full_name']} — {row['class_name']}",
                f"delete:{row['id']}"
            )
        ])

    buttons.append([button("⬅️ Menyu", "menu")])

    await callback.message.answer(
        "O‘chirish:",
        reply_markup=keyboard(buttons)
    )

    await callback.answer()


@dp.callback_query(F.data.startswith("delete:"))
async def callback_delete(callback: CallbackQuery):
    student_id = int(callback.data.split(":")[1])
    delete_student(student_id)

    await callback.message.answer(
        "✅ O‘quvchi o‘chirildi.",
        reply_markup=menu_keyboard(role(callback.from_user.id))
    )

    await callback.answer()


@dp.callback_query(F.data == "subjects")
async def callback_subjects(callback: CallbackQuery):
    if role(callback.from_user.id) not in ["admin", "teacher"]:
        await callback.answer("Ruxsat yo‘q.", show_alert=True)
        return

    subject_list = subjects()

    text = "📘 Fanlar:\n\n"

    if subject_list:
        text += "\n".join(f"• {item}" for item in subject_list)
    else:
        text += "Fan yo‘q."

    await callback.message.answer(
        text,
        reply_markup=keyboard([
            [button("➕ Fan qo‘shish", "add_subject")],
            [button("⬅️ Menyu", "menu")]
        ])
    )

    await callback.answer()


@dp.callback_query(F.data == "add_subject")
async def callback_add_subject(callback: CallbackQuery):
    set_state(callback.from_user.id, "subject")
    await callback.message.answer("Fan nomini kiriting:")
    await callback.answer()


@dp.callback_query(F.data == "teachers")
async def callback_teachers(callback: CallbackQuery):
    if role(callback.from_user.id) != "admin":
        await callback.answer("Faqat admin.", show_alert=True)
        return

    text = "👨‍🏫 Teacherlar:\n\n"

    for row in teachers():
        text += (
            f"🆔 {row['id']} | "
            f"{row['full_name']} | "
            f"{row['teacher_code']} | "
            f"{'✅ Ulangan' if row['telegram_id'] else '⏳ Ulanmagan'}\n"
        )

    await callback.message.answer(
        text,
        reply_markup=keyboard([
            [button("➕ Teacher qo‘shish", "add_teacher")],
            [button("⬅️ Menyu", "menu")]
        ])
    )

    await callback.answer()


@dp.callback_query(F.data == "add_teacher")
async def callback_add_teacher(callback: CallbackQuery):
    set_state(callback.from_user.id, "teacher")
    await callback.message.answer(
        "Format:\n"
        "Ali Ustoz, 9-A, 10-B\n\n"
        "Agar hamma sinflarni ko‘rsin desangiz:\n"
        "Ali Ustoz"
    )
    await callback.answer()


# ================= ATTENDANCE =================

@dp.callback_query(F.data == "attendance")
async def callback_attendance(callback: CallbackQuery):
    if role(callback.from_user.id) != "teacher":
        await callback.answer("Faqat teacher.", show_alert=True)
        return

    class_list = classes_for_teacher(callback.from_user.id)

    if not class_list:
        await callback.message.answer("Hali sinf/o‘quvchi yo‘q.")
        await callback.answer()
        return

    buttons = [
        [button(class_name, f"attendance_class:{class_name}")]
        for class_name in class_list
    ]

    buttons.append([button("⬅️ Menyu", "menu")])

    await callback.message.answer(
        "🏫 Sinfni tanlang:",
        reply_markup=keyboard(buttons)
    )

    await callback.answer()


@dp.callback_query(F.data.startswith("attendance_class:"))
async def callback_attendance_class(callback: CallbackQuery):
    class_name = callback.data.split(":", 1)[1]

    SESSION[callback.from_user.id] = {
        "class": class_name,
        "subject": None,
        "students": {}
    }

    subject_list = subjects()

    if not subject_list:
        await callback.message.answer("Avval fan qo‘shing.")
        await callback.answer()
        return

    buttons = [
        [button(subject_name, f"attendance_subject:{subject_name}")]
        for subject_name in subject_list
    ]

    buttons.append([button("⬅️ Menyu", "menu")])

    await callback.message.answer(
        f"📘 Fanni tanlang\n\n🏫 {class_name}",
        reply_markup=keyboard(buttons)
    )

    await callback.answer()


async def render_attendance(callback: CallbackQuery, edit=True):
    current_session = SESSION[callback.from_user.id]
    rows = students_by_class(current_session["class"])

    text = (
        "📚 Davomat\n\n"
        f"🏫 {current_session['class']}\n"
        f"📘 {current_session['subject']}\n"
        f"📅 {date.today()}\n\n"
    )

    buttons = []

    for row in rows:
        current_status = current_session["students"].get(row["id"], "present")

        text += f"{status_icon(current_status)} {row['full_name']}\n"

        buttons.append([
            button(
                f"{status_icon(current_status)} {row['full_name']}",
                f"attendance_toggle:{row['id']}"
            )
        ])

    buttons.append([button("💾 Saqlash", "attendance_save")])
    buttons.append([button("⬅️ Menyu", "menu")])

    if edit:
        await callback.message.edit_text(
            text,
            reply_markup=keyboard(buttons)
        )
    else:
        await callback.message.answer(
            text,
            reply_markup=keyboard(buttons)
        )


@dp.callback_query(F.data.startswith("attendance_subject:"))
async def callback_attendance_subject(callback: CallbackQuery):
    current_session = SESSION.get(callback.from_user.id)

    if not current_session:
        await callback.answer("Avval sinf tanlang.", show_alert=True)
        return

    current_session["subject"] = callback.data.split(":", 1)[1]
    current_session["students"] = {
        row["id"]: "present"
        for row in students_by_class(current_session["class"])
    }

    await render_attendance(callback, edit=False)
    await callback.answer()


@dp.callback_query(F.data.startswith("attendance_toggle:"))
async def callback_attendance_toggle(callback: CallbackQuery):
    current_session = SESSION.get(callback.from_user.id)

    if not current_session:
        await callback.answer("Davomat sessiyasi yo‘q.", show_alert=True)
        return

    student_id = int(callback.data.split(":")[1])
    current_status = current_session["students"].get(student_id, "present")

    next_status = STATUS_ORDER[
        (STATUS_ORDER.index(current_status) + 1) % len(STATUS_ORDER)
    ]

    current_session["students"][student_id] = next_status

    await render_attendance(callback, edit=True)
    await callback.answer()


@dp.callback_query(F.data == "attendance_save")
async def callback_attendance_save(callback: CallbackQuery):
    current_session = SESSION.get(callback.from_user.id)

    if not current_session:
        await callback.answer("Sessiya yo‘q.", show_alert=True)
        return

    today = str(date.today())

    counts = {
        "present": 0,
        "absent": 0,
        "late": 0
    }

    absent_students = []

    for row in students_by_class(current_session["class"]):
        current_status = current_session["students"].get(row["id"], "present")
        counts[current_status] += 1

        save_attendance(
            row["id"],
            row["full_name"],
            current_session["class"],
            current_session["subject"],
            current_status,
            callback.from_user.id
        )

        if current_status == "absent":
            absent_students.append(row["full_name"])
            await notify_absence(
                row["id"],
                row["full_name"],
                current_session["class"],
                current_session["subject"],
                today
            )

    SESSION.pop(callback.from_user.id, None)

    text = (
        "✅ Davomat saqlandi\n\n"
        f"🏫 {current_session['class']}\n"
        f"📘 {current_session['subject']}\n"
        f"✅ Keldi: {counts['present']}\n"
        f"❌ Kelmadi: {counts['absent']}\n"
        f"⏰ Kechikdi: {counts['late']}"
    )

    if absent_students:
        text += "\n\nKelmaganlar:\n" + "\n".join(absent_students)

    await callback.message.answer(text, reply_markup=teacher_keyboard())
    await callback.answer()


# ================= REPORTS =================

@dp.callback_query(F.data == "reports")
async def callback_reports(callback: CallbackQuery):
    if role(callback.from_user.id) not in ["admin", "teacher"]:
        await callback.answer("Ruxsat yo‘q.", show_alert=True)
        return

    current_stats = stats()
    current_month = str(date.today())[:7]
    present, absent, late, total, percent = monthly(current_month)

    text = (
        "📊 Hisobotlar\n\n"
        f"👥 O‘quvchilar: {current_stats['students']}\n"
        f"👨‍🏫 Teacherlar: {current_stats['teachers']}\n"
        f"📘 Fanlar: {current_stats['subjects']}\n\n"
        f"✅ Keldi: {current_stats['present']}\n"
        f"❌ Kelmadi: {current_stats['absent']}\n"
        f"⏰ Kechikdi: {current_stats['late']}\n\n"
        f"📅 {current_month}: ✅ {present} ❌ {absent} ⏰ {late} 📈 {percent}%\n\n"
        "📈 TOP kelmaganlar:\n"
    )

    absent_rows = top_absent()

    if absent_rows:
        for index, row in enumerate(absent_rows, 1):
            text += f"{index}. {row['full_name']} — {row['total']}\n"
    else:
        text += "Ma’lumot yo‘q.\n"

    await callback.message.answer(
        text,
        reply_markup=keyboard([
            [button("📤 Excel", "excel")],
            [button("📄 PDF", "pdf")],
            [button("⬅️ Menyu", "menu")]
        ])
    )

    await callback.answer()


@dp.callback_query(F.data == "analytics")
async def callback_analytics(callback: CallbackQuery):
    if role(callback.from_user.id) not in ["admin", "teacher"]:
        await callback.answer("Ruxsat yo‘q.", show_alert=True)
        return

    text = "📈 Fanlar statistikasi:\n\n"

    rows = subject_stats()

    if rows:
        for row in rows:
            text += (
                f"• {row['subject']}: "
                f"✅ {row['present'] or 0} | "
                f"❌ {row['absent'] or 0} | "
                f"⏰ {row['late'] or 0}\n"
            )
    else:
        text += "Ma’lumot yo‘q."

    await callback.message.answer(
        text,
        reply_markup=menu_keyboard(role(callback.from_user.id))
    )

    await callback.answer()


@dp.callback_query(F.data == "excel")
async def callback_excel(callback: CallbackQuery):
    if role(callback.from_user.id) not in ["admin", "teacher"]:
        await callback.answer("Ruxsat yo‘q.", show_alert=True)
        return

    rows = attendance_by_date(str(date.today()))

    if not rows:
        await callback.message.answer("Bugungi davomat yo‘q.")
        await callback.answer()
        return

    path = f"exports/excel/attendance_{date.today()}.xlsx"

    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.title = "Davomat"

    sheet.append(["ID", "F.I.SH", "Sinf", "Fan", "Status", "Sana"])

    for row in rows:
        sheet.append([
            row["id"],
            row["full_name"],
            row["class_name"],
            row["subject"],
            status_text(row["status"]),
            row["date"]
        ])

    workbook.save(path)

    await callback.message.answer_document(FSInputFile(path))
    await callback.answer()


@dp.callback_query(F.data == "pdf")
async def callback_pdf(callback: CallbackQuery):
    if role(callback.from_user.id) not in ["admin", "teacher"]:
        await callback.answer("Ruxsat yo‘q.", show_alert=True)
        return

    rows = attendance_by_date(str(date.today()))

    if not rows:
        await callback.message.answer("Bugungi davomat yo‘q.")
        await callback.answer()
        return

    path = f"exports/pdf/attendance_{date.today()}.pdf"

    pdf = canvas.Canvas(path, pagesize=A4)
    y = 800

    pdf.setFont("Helvetica-Bold", 16)
    pdf.drawString(50, y, "School ERP - Davomat hisoboti")
    y -= 30

    pdf.setFont("Helvetica", 10)

    for row in rows:
        if y < 50:
            pdf.showPage()
            y = 800
            pdf.setFont("Helvetica", 10)

        pdf.drawString(
            50,
            y,
            (
                f"ID: {row['id']} | "
                f"{row['full_name']} | "
                f"{row['class_name']} | "
                f"{row['subject']} | "
                f"{status_text(row['status'])} | "
                f"{row['date']}"
            )
        )
        y -= 15

    pdf.save()

    await callback.message.answer_document(FSInputFile(path))
    await callback.answer()


@dp.callback_query(F.data == "broadcast")
async def callback_broadcast(callback: CallbackQuery):
    if role(callback.from_user.id) != "admin":
        await callback.answer("Faqat admin.", show_alert=True)
        return

    set_state(callback.from_user.id, "broadcast")
    await callback.message.answer("Broadcast matnini kiriting:")
    await callback.answer()


@dp.callback_query(F.data == "backup")
async def callback_backup(callback: CallbackQuery):
    if role(callback.from_user.id) != "admin":
        await callback.answer("Faqat admin.", show_alert=True)
        return

    path = f"backups/davomat_{datetime.now().strftime('%Y%m%d_%H%M%S')}.db"
    shutil.copy(DB_NAME, path)

    await callback.message.answer_document(FSInputFile(path))
    await callback.answer()


# ================= WEB DASHBOARD =================

# Web authentication is separate from Telegram authentication.
# The existing Telegram bot code above is intentionally kept unchanged.
WEB_SESSIONS = {}


def esc(value):
    return html.escape(str(value if value is not None else ""))


def badge(status):
    label = status_text(status)
    css = {"present": "present", "absent": "absent", "late": "late"}.get(status, "")
    return f'<span class="badge {css}">{esc(label)}</span>'


def web_role(request: Request):
    token = request.cookies.get("school_erp_web_session")
    return WEB_SESSIONS.get(token) if token else None


def web_guard(request: Request, allowed):
    r = web_role(request)
    if r not in allowed:
        return None
    return r


def admin_only(request: Request):
    return web_guard(request, ["admin"])


def ensure_web_tables():
    conn = db(); c = conn.cursor()
    c.execute('''CREATE TABLE IF NOT EXISTS payments(
        id INTEGER PRIMARY KEY AUTOINCREMENT, student_id INTEGER,
        student_name TEXT, month TEXT, amount REAL DEFAULT 0,
        status TEXT DEFAULT 'pending', paid_at TEXT, note TEXT)''')
    c.execute('''CREATE TABLE IF NOT EXISTS schedules(
        id INTEGER PRIMARY KEY AUTOINCREMENT, day_name TEXT, time TEXT,
        class_name TEXT, subject TEXT, teacher TEXT, room TEXT)''')
    conn.commit(); conn.close()


def web_payment_stats():
    ensure_web_tables(); conn=db(); c=conn.cursor()
    c.execute("SELECT COALESCE(SUM(amount),0) x FROM payments WHERE status='paid'"); paid=c.fetchone()["x"] or 0
    c.execute("SELECT COALESCE(SUM(amount),0) x FROM payments WHERE status!='paid'"); pending=c.fetchone()["x"] or 0
    c.execute("SELECT COUNT(*) x FROM payments WHERE status='paid'"); paid_count=c.fetchone()["x"] or 0
    c.execute("SELECT COUNT(*) x FROM payments WHERE status!='paid'"); pending_count=c.fetchone()["x"] or 0
    conn.close(); return float(paid),float(pending),int(paid_count),int(pending_count)


def web_payments(limit=200):
    ensure_web_tables(); conn=db(); c=conn.cursor()
    c.execute("SELECT * FROM payments ORDER BY id DESC LIMIT ?",(limit,)); rows=c.fetchall(); conn.close(); return rows


def web_schedules():
    ensure_web_tables(); conn=db(); c=conn.cursor()
    c.execute("SELECT * FROM schedules ORDER BY id DESC"); rows=c.fetchall(); conn.close(); return rows


def web_add_payment(student_id, month, amount, status, note=""):
    s=student(student_id)
    if not s: return False
    ensure_web_tables(); conn=db(); c=conn.cursor()
    paid_at=datetime.now().isoformat() if status=='paid' else None
    c.execute("INSERT INTO payments(student_id,student_name,month,amount,status,paid_at,note) VALUES(?,?,?,?,?,?,?)",
              (student_id,s['full_name'],month,float(amount),status,paid_at,note))
    conn.commit(); conn.close(); return True


def web_add_schedule(day_name,time_value,class_name,subject,teacher,room):
    ensure_web_tables(); conn=db(); c=conn.cursor()
    c.execute("INSERT INTO schedules(day_name,time,class_name,subject,teacher,room) VALUES(?,?,?,?,?,?)",
              (day_name,time_value,class_name,subject,teacher,room))
    conn.commit(); conn.close()


def web_delete_payment(payment_id):
    ensure_web_tables(); conn=db(); c=conn.cursor(); c.execute("DELETE FROM payments WHERE id=?",(payment_id,)); conn.commit(); conn.close()


def web_delete_schedule(schedule_id):
    ensure_web_tables(); conn=db(); c=conn.cursor(); c.execute("DELETE FROM schedules WHERE id=?",(schedule_id,)); conn.commit(); conn.close()


def web_delete_teacher(teacher_id):
    conn=db(); c=conn.cursor(); c.execute("DELETE FROM teacher_classes WHERE teacher_id=?",(teacher_id,)); c.execute("DELETE FROM teachers WHERE id=?",(teacher_id,)); conn.commit(); conn.close()


def web_delete_subject(name):
    conn=db(); c=conn.cursor(); c.execute("DELETE FROM subjects WHERE name=?",(name,)); conn.commit(); conn.close()


def report_rows(start_date, end_date):
    conn=db(); c=conn.cursor()
    c.execute('''SELECT date, COUNT(*) total,
        SUM(CASE WHEN status='present' THEN 1 ELSE 0 END) present,
        SUM(CASE WHEN status='absent' THEN 1 ELSE 0 END) absent,
        SUM(CASE WHEN status='late' THEN 1 ELSE 0 END) late
        FROM attendance WHERE date BETWEEN ? AND ? GROUP BY date ORDER BY date''',(str(start_date),str(end_date)))
    rows=c.fetchall(); conn.close(); return rows


def report_period(period):
    today=date.today()
    if period=='weekly':
        start=today-timedelta(days=today.weekday()); end=today
        title='Haftalik hisobot'
    elif period=='monthly':
        start=today.replace(day=1); end=today
        title='Oylik hisobot'
    else:
        start=end=today; title='Kunlik hisobot'
    return start,end,title


def export_report_excel(period):
    start,end,title=report_period(period); rows=report_rows(start,end)
    path=Path('exports/excel')/f'report_{period}_{today_key()}.xlsx'
    wb=openpyxl.Workbook(); ws=wb.active; ws.title=title[:31]
    ws.append(['School ERP',title]); ws.append(['Davr',f'{start} — {end}'])
    ws.append([]); ws.append(['Sana','Jami','Keldi','Kelmadi','Kechikdi','Qatnashuv %'])
    for r in rows:
        total=int(r['total'] or 0); present=int(r['present'] or 0); pct=round(present/total*100,1) if total else 0
        ws.append([r['date'],total,present,int(r['absent'] or 0),int(r['late'] or 0),pct])
    for col in range(1,7): ws.column_dimensions[chr(64+col)].width=20
    wb.save(path); return path


def today_key():
    return datetime.now().strftime('%Y%m%d_%H%M%S')


def export_report_pdf(period):
    start,end,title=report_period(period); rows=report_rows(start,end)
    path=Path('exports/pdf')/f'report_{period}_{today_key()}.pdf'
    pdf=canvas.Canvas(str(path),pagesize=A4); y=800
    pdf.setFont('Helvetica-Bold',16); pdf.drawString(45,y,'School ERP - '+title); y-=24
    pdf.setFont('Helvetica',10); pdf.drawString(45,y,f'Davr: {start} - {end}'); y-=28
    pdf.drawString(45,y,'Sana     Jami     Keldi     Kelmadi     Kechikdi     Qatnashuv %'); y-=18
    if not rows: pdf.drawString(45,y,'Bu davr uchun davomat ma\'lumoti topilmadi.')
    for r in rows:
        if y<55: pdf.showPage(); y=800; pdf.setFont('Helvetica',10)
        total=int(r['total'] or 0); present=int(r['present'] or 0); pct=round(present/total*100,1) if total else 0
        pdf.drawString(45,y,f"{r['date']}   {total:<7} {present:<8} {int(r['absent'] or 0):<10} {int(r['late'] or 0):<10} {pct}%"); y-=16
    pdf.save(); return path


def layout(title, body, role, active='dashboard'):
    admin = role=='admin'
    items=[('dashboard','/','🏠','Dashboard'),('students','/students','👥','O‘quvchilar'),('attendance','/attendance','✅','Davomat'),('schedule','/schedule','📅','Dars jadvali'),('reports','/reports','📊','Hisobotlar')]
    if admin:
        items += [('teachers','/teachers','👨‍🏫','O‘qituvchilar'),('subjects','/subjects','📚','Fanlar'),('payments','/payments','💳','Oylik to‘lovlar')]
    else:
        items += [('payments','/payments','💳','To‘lovlarni ko‘rish')]
    nav=''.join(f'<a class="nav {"on" if active==k else ""}" href="{u}"><span class="nav-icon">{i}</span><span>{l}</span></a>' for k,u,i,l in items)
    return f'''<!doctype html><html lang="uz"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>{esc(title)} — School ERP</title>
    <style>
    *{{box-sizing:border-box}}
    :root{{--bg:#f4f7fb;--ink:#12213f;--muted:#6b7a90;--line:#e6edf5;--blue:#2563eb;--navy:#0b1f4d}}
    body{{margin:0;font-family:Inter,Segoe UI,Arial,sans-serif;background:var(--bg);color:var(--ink)}}
    .side{{position:fixed;inset:0 auto 0 0;width:258px;background:linear-gradient(180deg,#071a45 0%,#0b397f 55%,#0d55b8 100%);color:#fff;padding:22px 15px;box-shadow:10px 0 35px #102b5520;z-index:5}}
    .brand{{display:flex;align-items:center;gap:10px;font-size:22px;font-weight:900;padding:12px;margin-bottom:12px;letter-spacing:-.3px}}
    .brand-mark{{width:42px;height:42px;border-radius:13px;background:#ffffff18;display:grid;place-items:center;font-size:24px;box-shadow:inset 0 1px #ffffff20}}
    .role{{margin:0 8px 18px;padding:12px 13px;border-radius:14px;background:#ffffff12;border:1px solid #ffffff12;font-size:12px;color:#cfe0ff}}
    .role b{{display:block;color:#fff;margin-top:3px;font-size:13px}}
    .nav{{display:flex;align-items:center;gap:10px;text-decoration:none;color:#dbeafe;padding:12px 13px;border-radius:13px;margin:5px 0;font-weight:700;font-size:14px;transition:.18s}}
    .nav-icon{{width:25px;text-align:center;font-size:17px}}
    .nav:hover{{background:#ffffff14;color:#fff;transform:translateX(2px)}}
    .nav.on{{background:linear-gradient(90deg,#2563eb,#3b82f6);color:#fff;box-shadow:0 10px 25px #0002}}
    .logout{{position:absolute;bottom:18px;left:15px;right:15px;background:#ffffff0e}}
    .main{{margin-left:258px;padding:28px 32px 45px;max-width:1600px}}
    .topbar{{display:flex;align-items:flex-start;justify-content:space-between;gap:20px;margin-bottom:22px}}
    .eyebrow{{font-size:12px;font-weight:800;color:#2563eb;text-transform:uppercase;letter-spacing:1.1px;margin-bottom:5px}}
    h1{{margin:0;font-size:31px;letter-spacing:-.7px}} .subtitle,.muted{{color:var(--muted)}}
    .subtitle{{margin-top:7px;font-size:14px}}
    .welcome{{background:linear-gradient(135deg,#0b2d73,#1769e8);color:white;border-radius:24px;padding:23px 25px;box-shadow:0 18px 40px #174ea830;position:relative;overflow:hidden}}
    .welcome:after{{content:"";position:absolute;width:220px;height:220px;border-radius:50%;right:-75px;top:-100px;background:#ffffff12}}
    .welcome h2{{margin:0;font-size:24px}} .welcome p{{margin:7px 0 0;color:#dbeafe;font-size:13px}}
    .quick{{display:flex;gap:9px;flex-wrap:wrap;margin-top:14px;position:relative;z-index:1}}
    .quick a{{text-decoration:none;color:#fff;background:#ffffff16;border:1px solid #ffffff25;padding:8px 11px;border-radius:10px;font-size:12px;font-weight:800}}
    .cards{{display:grid;grid-template-columns:repeat(4,minmax(170px,1fr));gap:15px;margin:20px 0}}
    .stat{{position:relative;overflow:hidden;border-radius:19px;padding:19px;color:white;min-height:125px;box-shadow:0 12px 28px #18345b18}}
    .stat .label{{font-size:13px;font-weight:800;opacity:.92}} .stat .value{{font-size:30px;font-weight:900;margin-top:12px;letter-spacing:-1px}}
    .stat .mini{{font-size:11px;margin-top:7px;opacity:.88}} .stat .icon{{position:absolute;right:15px;top:14px;font-size:28px;opacity:.9}}
    .stat:after{{content:"";position:absolute;width:110px;height:110px;border-radius:50%;right:-35px;bottom:-55px;background:#fff1}}
    .blue{{background:linear-gradient(135deg,#2563eb,#06a6f2)}} .purple{{background:linear-gradient(135deg,#5b21b6,#8b5cf6)}} .green{{background:linear-gradient(135deg,#047857,#22c55e)}} .red{{background:linear-gradient(135deg,#be123c,#f43f5e)}} .orange{{background:linear-gradient(135deg,#c2410c,#f59e0b)}} .cyan{{background:linear-gradient(135deg,#0e7490,#06b6d4)}} .pink{{background:linear-gradient(135deg,#be185d,#ec4899)}} .indigo{{background:linear-gradient(135deg,#3730a3,#6366f1)}}
    .grid2{{display:grid;grid-template-columns:minmax(0,1.55fr) minmax(300px,.85fr);gap:18px}}
    .panel{{background:#fff;border-radius:20px;padding:20px;box-shadow:0 9px 30px #17325d0c;border:1px solid #edf2f7}}
    .panel-head{{display:flex;justify-content:space-between;align-items:center;gap:10px;margin-bottom:15px}} .panel-head h3{{margin:0;font-size:17px}} .panel-head small{{color:var(--muted)}}
    .bar-chart{{height:245px;display:flex;align-items:flex-end;gap:13px;padding:16px 7px 0;border-bottom:1px solid var(--line)}}
    .bar-item{{flex:1;height:100%;display:flex;flex-direction:column;justify-content:flex-end;align-items:center;gap:7px;min-width:35px}}
    .bar-track{{height:175px;width:100%;display:flex;align-items:flex-end;justify-content:center;gap:4px}}
    .bar{{width:42%;min-width:7px;border-radius:7px 7px 2px 2px;transition:.2s;box-shadow:0 5px 12px #2563eb18}}
    .bar:hover{{transform:translateY(-4px)}} .bar.present{{background:linear-gradient(180deg,#22c55e,#16a34a)}} .bar.absent{{background:linear-gradient(180deg,#fb7185,#e11d48)}} .bar.late{{background:linear-gradient(180deg,#fbbf24,#f59e0b)}}
    .bar-label{{font-size:11px;color:#6b7a90;font-weight:800}} .bar-num{{font-size:10px;color:#94a3b8}}
    .legend{{display:flex;gap:15px;flex-wrap:wrap;margin-top:12px;font-size:11px;color:#64748b;font-weight:700}} .dot{{display:inline-block;width:9px;height:9px;border-radius:50%;margin-right:5px}} .d-green{{background:#22c55e}} .d-red{{background:#e11d48}} .d-orange{{background:#f59e0b}}
    .donut-wrap{{display:flex;align-items:center;justify-content:center;gap:22px;min-height:245px}}
    .donut{{width:172px;height:172px;border-radius:50%;background:conic-gradient(#22c55e 0 var(--present),#e11d48 var(--present) var(--absent),#f59e0b var(--absent) 100%);position:relative;display:grid;place-items:center;box-shadow:0 12px 30px #17325d18}}
    .donut:after{{content:"";width:108px;height:108px;border-radius:50%;background:#fff;box-shadow:inset 0 0 0 1px #edf2f7}}
    .donut-center{{position:absolute;z-index:2;text-align:center}} .donut-center strong{{display:block;font-size:26px}} .donut-center span{{font-size:10px;color:#7b8798;font-weight:800}}
    .donut-legend{{display:grid;gap:13px}} .legend-row{{display:flex;align-items:center;gap:8px;font-size:12px;color:#64748b}} .legend-row b{{color:#12213f;margin-left:auto}}
    .rank-list{{display:grid;gap:10px}} .rank{{display:grid;grid-template-columns:32px 1fr auto;align-items:center;gap:10px;padding:11px 12px;border:1px solid #edf2f7;border-radius:13px;background:#fbfdff}}
    .rank-no{{width:30px;height:30px;border-radius:10px;background:#fee2e2;color:#be123c;display:grid;place-items:center;font-weight:900;font-size:12px}} .rank-name{{font-weight:800;font-size:13px}} .rank-class{{font-size:11px;color:#94a3b8;margin-top:2px}} .rank-count{{font-weight:900;color:#e11d48}}
    .action-grid{{display:grid;grid-template-columns:repeat(2,1fr);gap:10px}} .action{{text-decoration:none;padding:14px;border-radius:14px;border:1px solid #e7edf5;background:#f8fafc;color:#172b4d;font-weight:800;font-size:12px;transition:.18s}} .action:hover{{transform:translateY(-2px);box-shadow:0 8px 18px #17325d12;background:#fff}}
    .form{{display:grid;grid-template-columns:repeat(auto-fit,minmax(180px,1fr));gap:10px}} input,select{{width:100%;padding:12px;border:1px solid #d5dfec;border-radius:10px;font-size:14px}} button,.btn{{display:inline-block;border:0;background:#1261e8;color:white;padding:11px 15px;border-radius:10px;text-decoration:none;cursor:pointer;font-weight:700}} .danger{{background:#ef4444}} .green-btn{{background:#059669}} .gray{{background:#64748b}}
    .table{{overflow:auto;background:white;border-radius:16px;box-shadow:0 8px 30px #17325d10}} table{{width:100%;border-collapse:collapse;min-width:700px}} th,td{{padding:12px;border-bottom:1px solid #e7edf5;text-align:left;font-size:14px}} th{{background:#0a285c;color:white}} .badge{{padding:5px 9px;border-radius:999px;font-weight:700;font-size:12px}} .present{{background:#dcfce7;color:#166534}} .absent{{background:#fee2e2;color:#991b1b}} .late{{background:#fef3c7;color:#92400e}} .tabs{{display:flex;gap:8px;flex-wrap:wrap;margin:15px 0}} .tabs a{{padding:9px 13px;background:white;border-radius:10px;text-decoration:none}} .tabs a.sel{{background:#1261e8;color:white}}
    @media(max-width:1100px){{.cards{{grid-template-columns:repeat(2,1fr)}}.grid2{{grid-template-columns:1fr}}}}
    @media(max-width:800px){{.side{{position:static;width:100%}}.logout{{position:static;margin-top:15px}}.main{{margin-left:0;padding:18px}}.cards{{grid-template-columns:1fr 1fr}}.topbar{{display:block}}}}
    @media(max-width:520px){{.cards{{grid-template-columns:1fr}}.main{{padding:14px}}.donut-wrap{{flex-direction:column}}}}
    </style></head><body>
    <aside class="side"><div class="brand"><span class="brand-mark">🏫</span><span>School ERP</span></div><div class="role">Kirish roli:<b>{'👑 ADMIN' if admin else '👨‍🏫 O‘QITUVCHI'}</b></div>{nav}<a class="nav logout" href="/logout"><span class="nav-icon">🚪</span><span>Chiqish</span></a></aside><main class="main">{body}</main></body></html>'''


def table(headers, rows):
    h=''.join(f'<th>{esc(x)}</th>' for x in headers); b=''.join('<tr>'+''.join(f'<td>{x}</td>' for x in row)+'</tr>' for row in rows)
    return f'<div class="table"><table><thead><tr>{h}</tr></thead><tbody>{b}</tbody></table></div>'


@web_app.get('/login', response_class=HTMLResponse)
async def web_login_page():
    return '''<!doctype html><html lang="uz"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>School ERP — Kirish</title><style>body{margin:0;background:linear-gradient(135deg,#071d50,#1261e8);font-family:Arial;min-height:100vh;display:grid;place-items:center}.box{width:390px;max-width:92%;background:#fff;padding:32px;border-radius:22px;box-shadow:0 25px 70px #0005}.logo{text-align:center;font-size:48px}h1{text-align:center}.muted{text-align:center;color:#64748b;margin-bottom:25px}label{display:block;font-weight:700;margin:12px 0 7px}select,input{width:100%;padding:13px;border:1px solid #d4dce8;border-radius:10px;box-sizing:border-box}button{width:100%;padding:14px;margin-top:18px;background:#1261e8;color:#fff;border:0;border-radius:10px;font-weight:800;font-size:16px}</style></head><body><div class="box"><div class="logo">🏫</div><h1>School ERP</h1><div class="muted">Web boshqaruv paneliga kirish</div><form method="post" action="/login"><label>Rol</label><select name="role" required><option value="">Tanlang</option><option value="admin">👑 Admin</option><option value="teacher">👨‍🏫 O‘qituvchi</option></select><label>Kirish kodi</label><input type="password" name="code" placeholder="Kodni kiriting" required><button>🔐 Kirish</button></form></div></body></html>'''


@web_app.post('/login')
async def web_login(role:str=Form(...), code:str=Form(...)):
    role=role.strip().lower(); code=code.strip()
    if role=='admin' and code==ADMIN_CODE: target='/admin'
    elif role=='teacher' and code==TEACHER_CODE: target='/teacher'
    else: return HTMLResponse('<h2 style="text-align:center;margin-top:100px">❌ Kod noto‘g‘ri</h2><p style="text-align:center"><a href="/login">← Qaytish</a></p>',status_code=401)
    import secrets
    token=secrets.token_urlsafe(32); WEB_SESSIONS[token]=role
    r=RedirectResponse(target,status_code=303); r.set_cookie('school_erp_web_session',token,httponly=True,samesite='lax'); return r


@web_app.get('/logout')
async def web_logout(request:Request):
    token=request.cookies.get('school_erp_web_session'); WEB_SESSIONS.pop(token,None) if token else None
    r=RedirectResponse('/login',status_code=303); r.delete_cookie('school_erp_web_session'); return r


@web_app.get('/')
async def web_root(request:Request):
    role=web_role(request); return RedirectResponse('/admin' if role=='admin' else '/teacher' if role=='teacher' else '/login',status_code=303)


def dashboard_data():
    today = date.today()
    conn = db(); c = conn.cursor()
    daily=[]
    for offset in range(6,-1,-1):
        d = today - timedelta(days=offset)
        c.execute('''SELECT
            SUM(CASE WHEN status='present' THEN 1 ELSE 0 END) AS present,
            SUM(CASE WHEN status='absent' THEN 1 ELSE 0 END) AS absent,
            SUM(CASE WHEN status='late' THEN 1 ELSE 0 END) AS late
            FROM attendance WHERE date=?''', (str(d),))
        r=c.fetchone()
        daily.append({'date':str(d),'label':d.strftime('%d/%m'),'present':int(r['present'] or 0),'absent':int(r['absent'] or 0),'late':int(r['late'] or 0)})
    c.execute("SELECT COUNT(*) AS n FROM attendance WHERE date=? AND status='present'", (str(today),)); today_present=int(c.fetchone()['n'] or 0)
    c.execute("SELECT COUNT(*) AS n FROM attendance WHERE date=? AND status='absent'", (str(today),)); today_absent=int(c.fetchone()['n'] or 0)
    c.execute("SELECT COUNT(*) AS n FROM attendance WHERE date=? AND status='late'", (str(today),)); today_late=int(c.fetchone()['n'] or 0)
    conn.close()
    s=stats(); paid,pending,paid_count,pending_count=web_payment_stats()
    absent_list=top_absent(5)
    total_today=today_present+today_absent+today_late
    present_pct=round(today_present/total_today*100,1) if total_today else 0
    absent_pct=round(today_absent/total_today*100,1) if total_today else 0
    late_pct=round(today_late/total_today*100,1) if total_today else 0
    return s,daily,today_present,today_absent,today_late,total_today,present_pct,absent_pct,late_pct,paid,pending,paid_count,pending_count,absent_list


def dashboard_body(role):
    s,daily,tp,ta,tl,tt,pp,ap,lp,paid,pending,paid_count,pending_count,absent_list=dashboard_data()
    max_total=max([x['present']+x['absent']+x['late'] for x in daily]+[1])
    bars=[]
    for x in daily:
        p_h=max(5,round(x['present']/max_total*155)) if x['present'] else 3
        a_h=max(5,round(x['absent']/max_total*155)) if x['absent'] else 3
        l_h=max(5,round(x['late']/max_total*155)) if x['late'] else 3
        bars.append(f'''<div class="bar-item"><div class="bar-num">{x['present']+x['absent']+x['late']}</div><div class="bar-track"><div class="bar present" style="height:{p_h}px" title="Keldi: {x['present']}"></div><div class="bar absent" style="height:{a_h}px" title="Kelmadi: {x['absent']}"></div><div class="bar late" style="height:{l_h}px" title="Kechikdi: {x['late']}"></div></div><div class="bar-label">{x['label']}</div></div>''')
    if tt:
        present_stop=pp
        absent_stop=pp+ap
        donut_style=f'--present:{present_stop}%;--absent:{absent_stop}%;'
    else:
        donut_style='--present:100%;--absent:100%;'
    ranks=''.join(f'''<div class="rank"><div class="rank-no">{i}</div><div><div class="rank-name">{esc(x['full_name'])}</div><div class="rank-class">O‘quvchi ID: {x['student_id']}</div></div><div class="rank-count">{x['total']} marta</div></div>''' for i,x in enumerate(absent_list,1))
    if not ranks: ranks='<div class="muted" style="padding:18px;text-align:center">Hozircha kelmaganlar bo‘yicha ma’lumot yo‘q 🎉</div>'
    actions='''<a class="action" href="/students">👥 O‘quvchilarni boshqarish</a><a class="action" href="/attendance">✅ Davomatni ko‘rish</a><a class="action" href="/schedule">📅 Dars jadvali</a><a class="action" href="/reports">📊 Hisobotlar</a>'''
    if role=='admin': actions+='''<a class="action" href="/payments">💳 Oylik to‘lovlar</a><a class="action" href="/teachers">👨‍🏫 O‘qituvchilar</a>'''
    return f'''
    <div class="topbar"><div><div class="eyebrow">School ERP • Boshqaruv markazi</div><h1>Dashboard</h1><div class="subtitle">Bugungi o‘quv jarayonini bitta oynadan nazorat qiling.</div></div></div>
    <section class="welcome"><h2>{'👑 Admin boshqaruv paneli' if role=='admin' else '👨‍🏫 O‘qituvchi paneli'}</h2><p>Bugun: {date.today().strftime('%Y-%m-%d')} • Davomat, to‘lov va o‘quvchilar holati</p><div class="quick">{actions}</div></section>
    <div class="cards">
      <div class="stat blue"><span class="icon">👥</span><div class="label">O‘quvchilar</div><div class="value">{s['students']}</div><div class="mini">Jami o‘quvchilar</div></div>
      <div class="stat purple"><span class="icon">👨‍🏫</span><div class="label">O‘qituvchilar</div><div class="value">{s['teachers']}</div><div class="mini">Faol o‘qituvchilar</div></div>
      <div class="stat green"><span class="icon">✅</span><div class="label">Bugun keldi</div><div class="value">{tp}</div><div class="mini">Bugungi davomat</div></div>
      <div class="stat red"><span class="icon">❌</span><div class="label">Bugun kelmadi</div><div class="value">{ta}</div><div class="mini">Nazorat talab qilinadi</div></div>
      <div class="stat orange"><span class="icon">⏰</span><div class="label">Kechikdi</div><div class="value">{tl}</div><div class="mini">Bugungi kechikish</div></div>
      <div class="stat cyan"><span class="icon">📚</span><div class="label">Fanlar</div><div class="value">{s['subjects']}</div><div class="mini">Tizimdagi fanlar</div></div>
      <div class="stat pink"><span class="icon">💳</span><div class="label">To‘langan</div><div class="value">{paid:,.0f}</div><div class="mini">so‘m • {paid_count} ta to‘lov</div></div>
      <div class="stat indigo"><span class="icon">⏳</span><div class="label">Kutilayotgan</div><div class="value">{pending:,.0f}</div><div class="mini">so‘m • {pending_count} ta</div></div>
    </div>
    <div class="grid2">
      <section class="panel"><div class="panel-head"><h3>📈 Oxirgi 7 kunlik davomat</h3><small>Yashil — keldi • qizil — kelmadi • sariq — kechikdi</small></div><div class="bar-chart">{''.join(bars)}</div><div class="legend"><span><i class="dot d-green"></i>Keldi</span><span><i class="dot d-red"></i>Kelmadi</span><span><i class="dot d-orange"></i>Kechikdi</span></div></section>
      <section class="panel"><div class="panel-head"><h3>🎯 Bugungi davomat</h3><small>{tt} ta qayd</small></div><div class="donut-wrap"><div class="donut" style="{donut_style}"><div class="donut-center"><strong>{pp}%</strong><span>QATNASHUV</span></div></div><div class="donut-legend"><div class="legend-row"><i class="dot d-green"></i>Keldi <b>{tp}</b></div><div class="legend-row"><i class="dot d-red"></i>Kelmadi <b>{ta}</b></div><div class="legend-row"><i class="dot d-orange"></i>Kechikdi <b>{tl}</b></div></div></div></section>
    </div>
    <div class="grid2" style="margin-top:18px">
      <section class="panel"><div class="panel-head"><h3>🚨 Ko‘p kelmagan o‘quvchilar</h3><small>Umumiy davomat bo‘yicha</small></div><div class="rank-list">{ranks}</div></section>
      <section class="panel"><div class="panel-head"><h3>⚡ Tezkor amallar</h3><small>Bir bosishda kerakli bo‘lim</small></div><div class="action-grid">{actions}</div></section>
    </div>
    '''


@web_app.get('/admin',response_class=HTMLResponse)
async def web_admin(request:Request):
    if not admin_only(request): return RedirectResponse('/login',303)
    return layout('Admin',dashboard_body('admin'),'admin')


@web_app.get('/teacher',response_class=HTMLResponse)
async def web_teacher(request:Request):
    if not web_guard(request,['teacher']): return RedirectResponse('/login',303)
    return layout('O‘qituvchi',dashboard_body('teacher'),'teacher')


@web_app.get('/students',response_class=HTMLResponse)
async def web_students(request:Request):
    role=web_guard(request,['admin','teacher'])
    if not role:return RedirectResponse('/login',303)
    rows=[]
    for x in students(): rows.append([esc(x['id']),esc(x['full_name']),esc(x['class_name']),esc(x['parent_phone']),esc(x['parent_code']),f'<a class="btn danger" href="/students/delete/{x["id"]}">O‘chirish</a>' if role=='admin' else '—'])
    form='''<div class="panel"><h3>➕ O‘quvchi qo‘shish</h3><form class="form" method="post" action="/students/add"><input name="full_name" placeholder="F.I.SH" required><input name="class_name" placeholder="Sinf / guruh" required><input name="phone" placeholder="Ota-ona telefoni"><button>Saqlash</button></form></div>''' if role=='admin' else ''
    body=f'<h1>👥 O‘quvchilar</h1><div class="muted">O‘quvchilar ro‘yxati</div>{form}{table(["ID","F.I.SH","Sinf/Guruh","Telefon","Parent kod","Amal"],rows)}'
    return layout('O‘quvchilar',body,role,'students')


@web_app.post('/students/add')
async def web_student_add(request:Request,full_name:str=Form(...),class_name:str=Form(...),phone:str=Form('')):
    if not admin_only(request): return RedirectResponse('/login',303)
    add_student(full_name,class_name,phone); return RedirectResponse('/students',303)


@web_app.get('/students/delete/{student_id}')
async def web_student_delete(request:Request,student_id:int):
    if not admin_only(request): return RedirectResponse('/login',303)
    delete_student(student_id); return RedirectResponse('/students',303)


@web_app.get('/teachers',response_class=HTMLResponse)
async def web_teachers(request:Request):
    if not admin_only(request): return RedirectResponse('/login',303)
    rows=[]
    for x in teachers(): rows.append([esc(x['id']),esc(x['full_name']),esc(x['telegram_id'] or '—'),esc(x['teacher_code']),esc(', '.join(teacher_classes(x['telegram_id'])) if x['telegram_id'] else '—'),f'<a class="btn danger" href="/teachers/delete/{x["id"]}">O‘chirish</a>'])
    form='''<div class="panel"><h3>➕ O‘qituvchi qo‘shish</h3><form class="form" method="post" action="/teachers/add"><input name="name" placeholder="O‘qituvchi F.I.SH" required><input name="classes_text" placeholder="Guruhlar: 9-A, IELTS-1"><button>Saqlash</button></form><p class="muted">Saqlangandan keyin tizim avtomatik TCH-... kod beradi.</p></div>'''
    return layout('O‘qituvchilar',f'<h1>👨‍🏫 O‘qituvchilar</h1>{form}{table(["ID","F.I.SH","Telegram","Teacher kod","Guruhlar","Amal"],rows)}','admin','teachers')


@web_app.post('/teachers/add')
async def web_teacher_add(request:Request,name:str=Form(...),classes_text:str=Form('')):
    if not admin_only(request): return RedirectResponse('/login',303)
    add_teacher(name,classes_text); return RedirectResponse('/teachers',303)


@web_app.get('/teachers/delete/{teacher_id}')
async def web_teacher_delete(request:Request,teacher_id:int):
    if not admin_only(request): return RedirectResponse('/login',303)
    web_delete_teacher(teacher_id); return RedirectResponse('/teachers',303)


@web_app.get('/subjects',response_class=HTMLResponse)
async def web_subjects(request:Request):
    role=web_guard(request,['admin','teacher'])
    if not role:return RedirectResponse('/login',303)
    rows=[]
    for i,name in enumerate(subjects(),1): rows.append([i,esc(name),f'<a class="btn danger" href="/subjects/delete?name={name}">O‘chirish</a>' if role=='admin' else '—'])
    form='''<div class="panel"><h3>➕ Fan qo‘shish</h3><form class="form" method="post" action="/subjects/add"><input name="name" placeholder="Masalan: Informatika" required><button>Saqlash</button></form></div>''' if role=='admin' else ''
    return layout('Fanlar',f'<h1>📚 Fanlar</h1>{form}{table(["ID","Fan nomi","Amal"],rows)}',role,'subjects')


@web_app.post('/subjects/add')
async def web_subject_add(request:Request,name:str=Form(...)):
    if not admin_only(request): return RedirectResponse('/login',303)
    add_subject(name); return RedirectResponse('/subjects',303)


@web_app.get('/subjects/delete')
async def web_subject_delete(request:Request,name:str):
    if not admin_only(request): return RedirectResponse('/login',303)
    web_delete_subject(name); return RedirectResponse('/subjects',303)


@web_app.get('/attendance',response_class=HTMLResponse)
async def web_attendance(request:Request):
    role=web_guard(request,['admin','teacher'])
    if not role:return RedirectResponse('/login',303)
    rows=[]
    for x in attendance_by_date(str(date.today())): rows.append([x['id'],esc(x['full_name']),esc(x['class_name']),esc(x['subject']),badge(x['status']),x['date']])
    return layout('Davomat',f'<h1>✅ Bugungi davomat</h1><div class="muted">{date.today()}</div>{table(["ID","O‘quvchi","Guruh","Fan","Holat","Sana"],rows)}',role,'attendance')


@web_app.get('/payments',response_class=HTMLResponse)
async def web_payments_page(request:Request):
    role=web_guard(request,['admin','teacher'])
    if not role:return RedirectResponse('/login',303)
    paid,pending,pc,pnc=web_payment_stats(); rows=[]
    for x in web_payments(): rows.append([x['id'],esc(x['student_name']),esc(x['month']),f'{float(x["amount"]):,.0f}',esc(x['status']),esc(x['paid_at'] or '—'),f'<a class="btn danger" href="/payments/delete/{x["id"]}">O‘chirish</a>' if role=='admin' else '—'])
    form='''<div class="panel"><h3>➕ Oylik to‘lov qo‘shish</h3><form class="form" method="post" action="/payments/add"><select name="student_id" required><option value="">O‘quvchi</option>'''+''.join(f'<option value="{x["id"]}">{esc(x["full_name"])} — {esc(x["class_name"])}</option>' for x in students())+'''</select><input name="month" placeholder="2026-09" required><input name="amount" type="number" step="1" placeholder="Summa" required><select name="status"><option value="paid">To‘langan</option><option value="pending">Kutilmoqda</option></select><input name="note" placeholder="Izoh"><button>Saqlash</button></form></div>''' if role=='admin' else ''
    body=f'<h1>💳 Oylik to‘lovlar</h1><div class="cards"><div class="card">To‘langan<b>{paid:,.0f}</b></div><div class="card">Kutilmoqda<b>{pending:,.0f}</b></div></div>{form}{table(["ID","O‘quvchi","Oy","Summa","Status","Sana","Amal"],rows)}'
    return layout('To‘lovlar',body,role,'payments')


@web_app.post('/payments/add')
async def web_payment_add(request:Request,student_id:int=Form(...),month:str=Form(...),amount:float=Form(...),status:str=Form('paid'),note:str=Form('')):
    if not admin_only(request): return RedirectResponse('/login',303)
    web_add_payment(student_id,month,amount,status,note); return RedirectResponse('/payments',303)


@web_app.get('/payments/delete/{payment_id}')
async def web_payment_delete(request:Request,payment_id:int):
    if not admin_only(request): return RedirectResponse('/login',303)
    web_delete_payment(payment_id); return RedirectResponse('/payments',303)


@web_app.get('/schedule',response_class=HTMLResponse)
async def web_schedule_page(request:Request):
    role=web_guard(request,['admin','teacher'])
    if not role:return RedirectResponse('/login',303)
    rows=[]
    for x in web_schedules(): rows.append([esc(x['day_name']),esc(x['time']),esc(x['class_name']),esc(x['subject']),esc(x['teacher']),esc(x['room'] or '—'),f'<a class="btn danger" href="/schedule/delete/{x["id"]}">O‘chirish</a>' if role=='admin' else '—'])
    form='''<div class="panel"><h3>➕ Dars qo‘shish</h3><form class="form" method="post" action="/schedule/add"><input name="day_name" placeholder="Dushanba" required><input name="time_value" placeholder="09:00-10:30" required><input name="class_name" placeholder="9-A" required><input name="subject" placeholder="Ingliz tili" required><input name="teacher" placeholder="O‘qituvchi" required><input name="room" placeholder="Xona"><button>Saqlash</button></form></div>''' if role=='admin' else ''
    return layout('Jadval',f'<h1>📅 Dars jadvali</h1>{form}{table(["Kun","Vaqt","Guruh","Fan","O‘qituvchi","Xona","Amal"],rows)}',role,'schedule')


@web_app.post('/schedule/add')
async def web_schedule_add(request:Request,day_name:str=Form(...),time_value:str=Form(...),class_name:str=Form(...),subject:str=Form(...),teacher:str=Form(...),room:str=Form('')):
    if not admin_only(request): return RedirectResponse('/login',303)
    web_add_schedule(day_name,time_value,class_name,subject,teacher,room); return RedirectResponse('/schedule',303)


@web_app.get('/schedule/delete/{schedule_id}')
async def web_schedule_delete(request:Request,schedule_id:int):
    if not admin_only(request): return RedirectResponse('/login',303)
    web_delete_schedule(schedule_id); return RedirectResponse('/schedule',303)


@web_app.get('/reports',response_class=HTMLResponse)
async def web_reports(request:Request,period:str='daily'):
    role=web_guard(request,['admin','teacher'])
    if not role:return RedirectResponse('/login',303)
    if period not in ('daily','weekly','monthly'): period='daily'
    start,end,title=report_period(period); rows=report_rows(start,end)
    total=sum(int(x['total'] or 0) for x in rows); present=sum(int(x['present'] or 0) for x in rows); absent=sum(int(x['absent'] or 0) for x in rows); late=sum(int(x['late'] or 0) for x in rows); pct=round(present/total*100,1) if total else 0
    tr=[[x['date'],x['total'] or 0,x['present'] or 0,x['absent'] or 0,x['late'] or 0,round((x['present'] or 0)/(x['total'] or 1)*100,1)] for x in rows]
    tabs=''.join(f'<a class="{"sel" if period==p else ""}" href="/reports?period={p}">{t}</a>' for p,t in [('daily','📅 Kunlik'),('weekly','🗓 Haftalik'),('monthly','📆 Oylik')])
    body=f'<h1>📊 {title}</h1><div class="tabs">{tabs}</div><div class="panel"><b>Davr:</b> {start} — {end}<div class="cards"><div class="card">Jami<b>{total}</b></div><div class="card">Keldi<b>{present}</b></div><div class="card">Kelmadi<b>{absent}</b></div><div class="card">Kechikdi<b>{late}</b></div><div class="card">Qatnashuv<b>{pct}%</b></div></div><a class="btn green" href="/export/excel?period={period}">📥 Excel yuklab olish</a> <a class="btn" href="/export/pdf?period={period}">📄 PDF yuklab olish</a></div>{table(["Sana","Jami","Keldi","Kelmadi","Kechikdi","%"],tr)}'
    return layout('Hisobotlar',body,role,'reports')


@web_app.get('/export/excel')
async def web_export_excel(request:Request,period:str='daily'):
    if not web_guard(request,['admin','teacher']): return RedirectResponse('/login',303)
    if period not in ('daily','weekly','monthly'): period='daily'
    path=export_report_excel(period); return FileResponse(str(path),filename=path.name,media_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')


@web_app.get('/export/pdf')
async def web_export_pdf(request:Request,period:str='daily'):
    if not web_guard(request,['admin','teacher']): return RedirectResponse('/login',303)
    if period not in ('daily','weekly','monthly'): period='daily'
    path=export_report_pdf(period); return FileResponse(str(path),filename=path.name,media_type='application/pdf')


def run_web():
    uvicorn.run(
        web_app,
        host=WEB_HOST,
        port=WEB_PORT,
        log_level="warning"
    )


async def main():
    if not BOT_TOKEN:
        raise RuntimeError("BOT_TOKEN .env da yo‘q")

    init_db()

    threading.Thread(
        target=run_web,
        daemon=True
    ).start()

    print("Bot ishga tushdi.")
    print(f"Web local: http://127.0.0.1:{WEB_PORT}")
    print(f"Web host:  http://{WEB_HOST}:{WEB_PORT}")

    await dp.start_polling(bot)


if __name__ == "__main__":
    asyncio.run(main())
