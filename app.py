# -*- coding: utf-8 -*-
"""Maktab monitoring test platformasi (MVP).

Rollar: admin, teacher (o'qituvchi), student (o'quvchi).
Ishga tushirish:  python app.py   ->  http://localhost:5000
"""
import io
import json
import os
import random
import re
import secrets
import sqlite3
from datetime import datetime, timedelta, timezone
from functools import wraps

from flask import (Flask, abort, flash, g, jsonify, redirect, render_template,
                   request, send_file, session, url_for)
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from werkzeug.security import check_password_hash, generate_password_hash

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
INSTANCE = os.path.join(BASE_DIR, "instance")
os.makedirs(INSTANCE, exist_ok=True)
DB_PATH = os.environ.get("DB_PATH", os.path.join(INSTANCE, "monitoring.db"))

try:
    from zoneinfo import ZoneInfo
    TZ = ZoneInfo(os.environ.get("APP_TZ", "Asia/Tashkent"))
except Exception:  # pragma: no cover
    TZ = timezone(timedelta(hours=5))

LETTERS = "ABCD"
FMT = "%Y-%m-%d %H:%M:%S"
GRACE_SECONDS = 3


def now():
    return datetime.now(TZ).replace(tzinfo=None)


def iso(dt):
    return dt.strftime(FMT) if dt else None


def parse(s):
    return datetime.strptime(s, FMT) if s else None


def parse_local_input(s):
    """<input type=datetime-local> qiymatini bazaga mos formatga o'tkazadi."""
    s = (s or "").strip()
    if not s:
        return None
    s = s.replace("T", " ")
    if len(s) == 16:
        s += ":00"
    try:
        datetime.strptime(s, FMT)
    except ValueError:
        return None
    return s


def _secret_key():
    env = os.environ.get("SECRET_KEY")
    if env:
        return env
    path = os.path.join(INSTANCE, "secret.key")
    if not os.path.exists(path):
        with open(path, "w") as f:
            f.write(secrets.token_hex(32))
    with open(path) as f:
        return f.read().strip()


app = Flask(__name__)
app.config.update(
    SECRET_KEY=_secret_key(),
    SESSION_COOKIE_HTTPONLY=True,
    SESSION_COOKIE_SAMESITE="Lax",
    MAX_CONTENT_LENGTH=8 * 1024 * 1024,
)

# --------------------------------------------------------------------------
# Ma'lumotlar bazasi
# --------------------------------------------------------------------------
SCHEMA = """
CREATE TABLE IF NOT EXISTS users(
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  username TEXT UNIQUE NOT NULL,
  password_hash TEXT NOT NULL,
  pw_plain TEXT,
  role TEXT NOT NULL CHECK(role IN ('admin','teacher','student')),
  full_name TEXT NOT NULL,
  class_name TEXT,
  active INTEGER NOT NULL DEFAULT 1,
  created_at TEXT
);
CREATE TABLE IF NOT EXISTS subjects(
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  name TEXT UNIQUE NOT NULL
);
CREATE TABLE IF NOT EXISTS teacher_subjects(
  teacher_id INTEGER NOT NULL REFERENCES users(id) ON DELETE CASCADE,
  subject_id INTEGER NOT NULL REFERENCES subjects(id) ON DELETE CASCADE,
  PRIMARY KEY(teacher_id, subject_id)
);
CREATE TABLE IF NOT EXISTS questions(
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  teacher_id INTEGER NOT NULL REFERENCES users(id) ON DELETE CASCADE,
  subject_id INTEGER NOT NULL REFERENCES subjects(id) ON DELETE CASCADE,
  text TEXT NOT NULL,
  a TEXT, b TEXT, c TEXT, d TEXT,
  correct TEXT NOT NULL,
  deleted INTEGER NOT NULL DEFAULT 0,
  created_at TEXT
);
CREATE TABLE IF NOT EXISTS tests(
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  teacher_id INTEGER NOT NULL REFERENCES users(id) ON DELETE CASCADE,
  subject_id INTEGER NOT NULL REFERENCES subjects(id) ON DELETE CASCADE,
  title TEXT NOT NULL,
  classes TEXT NOT NULL,
  num_questions INTEGER NOT NULL,
  minutes INTEGER NOT NULL,
  start_at TEXT,
  end_at TEXT,
  attempts_allowed INTEGER NOT NULL DEFAULT 1,
  max_warnings INTEGER NOT NULL DEFAULT 3,
  auto_finish INTEGER NOT NULL DEFAULT 1,
  shuffle_q INTEGER NOT NULL DEFAULT 1,
  shuffle_opt INTEGER NOT NULL DEFAULT 1,
  show_result INTEGER NOT NULL DEFAULT 1,
  active INTEGER NOT NULL DEFAULT 1,
  created_at TEXT
);
CREATE TABLE IF NOT EXISTS attempts(
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  test_id INTEGER NOT NULL REFERENCES tests(id) ON DELETE CASCADE,
  student_id INTEGER NOT NULL REFERENCES users(id) ON DELETE CASCADE,
  started_at TEXT NOT NULL,
  finished_at TEXT,
  status TEXT,
  score INTEGER,
  total INTEGER,
  warnings INTEGER NOT NULL DEFAULT 0,
  layout TEXT NOT NULL
);
CREATE TABLE IF NOT EXISTS answers(
  attempt_id INTEGER NOT NULL REFERENCES attempts(id) ON DELETE CASCADE,
  question_id INTEGER NOT NULL,
  choice TEXT,
  PRIMARY KEY(attempt_id, question_id)
);
CREATE TABLE IF NOT EXISTS violations(
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  attempt_id INTEGER NOT NULL REFERENCES attempts(id) ON DELETE CASCADE,
  at TEXT,
  kind TEXT
);
CREATE INDEX IF NOT EXISTS idx_q_ts ON questions(teacher_id, subject_id, deleted);
CREATE INDEX IF NOT EXISTS idx_att_test ON attempts(test_id);
CREATE INDEX IF NOT EXISTS idx_att_student ON attempts(student_id);
"""


def db():
    if "db" not in g:
        conn = sqlite3.connect(DB_PATH, timeout=15)
        conn.row_factory = sqlite3.Row
        conn.execute("PRAGMA foreign_keys=ON")
        conn.execute("PRAGMA journal_mode=WAL")
        g.db = conn
    return g.db


@app.teardown_appcontext
def close_db(_exc):
    conn = g.pop("db", None)
    if conn is not None:
        conn.close()


def q(sql, args=(), one=False):
    rows = db().execute(sql, args).fetchall()
    if one:
        return rows[0] if rows else None
    return rows


def ex(sql, args=()):
    cur = db().execute(sql, args)
    db().commit()
    return cur.lastrowid


def init_db():
    conn = sqlite3.connect(DB_PATH)
    conn.executescript(SCHEMA)
    row = conn.execute("SELECT id FROM users WHERE role='admin' LIMIT 1").fetchone()
    if not row:
        pw = os.environ.get("ADMIN_PASSWORD", "admin123")
        conn.execute(
            "INSERT INTO users(username,password_hash,role,full_name,created_at) VALUES(?,?,?,?,?)",
            ("admin", generate_password_hash(pw), "admin", "Administrator", iso(now())),
        )
        print("[!] Admin yaratildi: login=admin, parol=%s  (kirgach almashtiring)" % pw)
    conn.commit()
    conn.close()


# --------------------------------------------------------------------------
# Xavfsizlik: CSRF, login, rollar
# --------------------------------------------------------------------------
@app.before_request
def csrf_protect():
    if request.method == "POST":
        tok = session.get("csrf")
        sent = request.form.get("csrf") or request.headers.get("X-CSRF")
        if not tok or not sent or not secrets.compare_digest(tok, sent):
            if request.path.startswith("/api/"):
                return jsonify(error="csrf"), 400
            abort(400, "Sessiya muddati tugagan. Sahifani yangilab qayta urinib ko'ring.")


@app.context_processor
def inject_globals():
    if "csrf" not in session:
        session["csrf"] = secrets.token_hex(16)
    return dict(csrf_token=session["csrf"], user=current_user())


@app.template_filter("dt")
def fmt_dt(s):
    if not s:
        return "—"
    try:
        return parse(s).strftime("%d.%m.%Y %H:%M")
    except Exception:
        return s


@app.template_filter("dtinput")
def fmt_dt_input(s):
    return s[:16].replace(" ", "T") if s else ""


def current_user():
    if "user_obj" in g:
        return g.user_obj
    u = None
    if "uid" in session:
        u = q("SELECT * FROM users WHERE id=? AND active=1", (session["uid"],), one=True)
    g.user_obj = u
    return u


def role_required(*roles):
    def deco(f):
        @wraps(f)
        def wrapper(*a, **k):
            u = current_user()
            if not u:
                session.pop("uid", None)
                if request.path.startswith("/api/"):
                    return jsonify(error="auth"), 401
                return redirect(url_for("login", next=request.path))
            if roles and u["role"] not in roles:
                abort(403)
            return f(*a, **k)
        return wrapper
    return deco


FAILS = {}  # (ip, username) -> [urinishlar, bloklangan_vaqt]


@app.route("/login", methods=["GET", "POST"])
def login():
    if request.method == "POST":
        username = request.form.get("username", "").strip().lower()
        password = request.form.get("password", "")
        key = (request.remote_addr, username)
        rec = FAILS.get(key)
        if rec and rec[1] and rec[1] > now():
            flash("Juda ko'p noto'g'ri urinish. Bir necha daqiqadan keyin qayta urinib ko'ring.", "err")
            return render_template("login.html")
        u = q("SELECT * FROM users WHERE username=? AND active=1", (username,), one=True)
        if u and check_password_hash(u["password_hash"], password):
            FAILS.pop(key, None)
            session.clear()
            session["uid"] = u["id"]
            session["csrf"] = secrets.token_hex(16)
            nxt = request.args.get("next", "")
            if nxt.startswith("/") and not nxt.startswith("//"):
                return redirect(nxt)
            return redirect(url_for("home"))
        rec = rec or [0, None]
        rec[0] += 1
        if rec[0] >= 6:
            rec = [0, now() + timedelta(minutes=5)]
        FAILS[key] = rec
        flash("Login yoki parol noto'g'ri.", "err")
    return render_template("login.html")


@app.route("/logout", methods=["POST"])
def logout():
    session.clear()
    return redirect(url_for("login"))


@app.route("/")
def home():
    u = current_user()
    if not u:
        return redirect(url_for("login"))
    return redirect(url_for({"admin": "admin_home", "teacher": "t_home", "student": "s_home"}[u["role"]]))


@app.errorhandler(400)
@app.errorhandler(403)
@app.errorhandler(404)
def err(e):
    return render_template("error.html", e=e), e.code


# --------------------------------------------------------------------------
# Yordamchi funksiyalar
# --------------------------------------------------------------------------
def grade(pct):
    """5 ballik baho (86+ → 5, 71+ → 4, 55+ → 3, qolgani → 2)."""
    return 5 if pct >= 86 else 4 if pct >= 71 else 3 if pct >= 55 else 2


STATUS_TXT = {
    "finished": "Yakunlangan",
    "time": "Vaqt tugadi",
    "warnings": "Ogohlantirish tufayli yakunlandi",
    None: "Jarayonda",
}

CYR = dict(zip("абвгдезийклмнопрстуфхцэыъь", "abvgdezijklmnoprstufxcey''"))
CYR.update({"ё": "yo", "ж": "j", "ч": "ch", "ш": "sh", "щ": "sh", "ю": "yu", "я": "ya",
            "ў": "o", "қ": "q", "ғ": "g", "ҳ": "h"})


def slugify(s):
    out = []
    for ch in s.lower():
        out.append(CYR.get(ch, ch))
    s = "".join(out)
    return re.sub(r"[^a-z0-9]", "", s)


def make_username(full_name):
    parts = [slugify(p) for p in full_name.split() if slugify(p)]
    base = ".".join(parts[:2]) or "oquvchi"
    name, n = base, 1
    while q("SELECT 1 FROM users WHERE username=?", (name,), one=True):
        n += 1
        name = "%s%d" % (base, n)
    return name


def make_password(n=6):
    alphabet = "abcdefghjkmnpqrstuvwxyz23456789"
    return "".join(secrets.choice(alphabet) for _ in range(n))


def as_int(v, default, lo, hi):
    try:
        x = int(v)
    except (TypeError, ValueError):
        return default
    return max(lo, min(hi, x))


# --------------------------------------------------------------------------
# ADMIN
# --------------------------------------------------------------------------
@app.route("/admin")
@role_required("admin")
def admin_home():
    subjects = q("SELECT * FROM subjects ORDER BY name")
    teachers = q("SELECT * FROM users WHERE role='teacher' ORDER BY full_name")
    tsub = {}
    for r in q("SELECT ts.teacher_id, s.id, s.name FROM teacher_subjects ts JOIN subjects s ON s.id=ts.subject_id"):
        tsub.setdefault(r["teacher_id"], []).append(r["id"])
    return render_template("admin.html", subjects=subjects, teachers=teachers, tsub=tsub)


@app.route("/admin/subject", methods=["POST"])
@role_required("admin")
def admin_subject():
    name = request.form.get("name", "").strip()
    if not name:
        flash("Fan nomini kiriting.", "err")
    elif q("SELECT 1 FROM subjects WHERE lower(name)=lower(?)", (name,), one=True):
        flash("Bunday fan allaqachon bor.", "err")
    else:
        ex("INSERT INTO subjects(name) VALUES(?)", (name,))
        flash("Fan qo'shildi.", "ok")
    return redirect(url_for("admin_home"))


@app.route("/admin/teacher", methods=["POST"])
@role_required("admin")
def admin_teacher():
    full_name = request.form.get("full_name", "").strip()
    username = request.form.get("username", "").strip().lower()
    password = request.form.get("password", "")
    sids = request.form.getlist("subject_ids")
    if not full_name or not username or len(password) < 6:
        flash("F.I.Sh., login va kamida 6 belgili parol kiriting.", "err")
    elif q("SELECT 1 FROM users WHERE username=?", (username,), one=True):
        flash("Bu login band.", "err")
    else:
        tid = ex(
            "INSERT INTO users(username,password_hash,role,full_name,created_at) VALUES(?,?,?,?,?)",
            (username, generate_password_hash(password), "teacher", full_name, iso(now())),
        )
        for s in sids:
            ex("INSERT OR IGNORE INTO teacher_subjects VALUES(?,?)", (tid, int(s)))
        flash("O'qituvchi qo'shildi.", "ok")
    return redirect(url_for("admin_home"))


@app.route("/admin/teacher/<int:tid>/subjects", methods=["POST"])
@role_required("admin")
def admin_teacher_subjects(tid):
    if not q("SELECT 1 FROM users WHERE id=? AND role='teacher'", (tid,), one=True):
        abort(404)
    ex("DELETE FROM teacher_subjects WHERE teacher_id=?", (tid,))
    for s in request.form.getlist("subject_ids"):
        ex("INSERT OR IGNORE INTO teacher_subjects VALUES(?,?)", (tid, int(s)))
    flash("Fanlar yangilandi.", "ok")
    return redirect(url_for("admin_home"))


@app.route("/admin/teacher/<int:tid>/password", methods=["POST"])
@role_required("admin")
def admin_teacher_password(tid):
    pw = request.form.get("password", "")
    if len(pw) < 6:
        flash("Parol kamida 6 belgi bo'lsin.", "err")
    else:
        ex("UPDATE users SET password_hash=? WHERE id=? AND role='teacher'", (generate_password_hash(pw), tid))
        flash("Parol almashtirildi.", "ok")
    return redirect(url_for("admin_home"))


@app.route("/admin/teacher/<int:tid>/toggle", methods=["POST"])
@role_required("admin")
def admin_teacher_toggle(tid):
    ex("UPDATE users SET active=1-active WHERE id=? AND role='teacher'", (tid,))
    return redirect(url_for("admin_home"))


@app.route("/password", methods=["POST"])
@role_required("admin", "teacher")
def change_own_password():
    old, new = request.form.get("old", ""), request.form.get("new", "")
    u = current_user()
    if not check_password_hash(u["password_hash"], old):
        flash("Joriy parol noto'g'ri.", "err")
    elif len(new) < 6:
        flash("Yangi parol kamida 6 belgi bo'lsin.", "err")
    else:
        ex("UPDATE users SET password_hash=? WHERE id=?", (generate_password_hash(new), u["id"]))
        flash("Parol almashtirildi.", "ok")
    return redirect(request.referrer or url_for("home"))


# --------------------------------------------------------------------------
# O'QITUVCHI: savollar
# --------------------------------------------------------------------------
def my_subjects():
    return q(
        "SELECT s.* FROM subjects s JOIN teacher_subjects ts ON ts.subject_id=s.id "
        "WHERE ts.teacher_id=? ORDER BY s.name", (current_user()["id"],))


def own_subject(sid):
    r = q("SELECT s.* FROM subjects s JOIN teacher_subjects ts ON ts.subject_id=s.id "
          "WHERE ts.teacher_id=? AND s.id=?", (current_user()["id"], sid), one=True)
    if not r:
        abort(404)
    return r


def own_question(qid):
    r = q("SELECT * FROM questions WHERE id=? AND teacher_id=? AND deleted=0", (qid, current_user()["id"]), one=True)
    if not r:
        abort(404)
    return r


def own_test(tid):
    r = q("SELECT t.*, s.name AS subject FROM tests t JOIN subjects s ON s.id=t.subject_id "
          "WHERE t.id=? AND t.teacher_id=?", (tid, current_user()["id"]), one=True)
    if not r:
        abort(404)
    return r


@app.route("/teacher")
@role_required("teacher")
def t_home():
    uid = current_user()["id"]
    subjects = q(
        "SELECT s.*, (SELECT COUNT(*) FROM questions qq WHERE qq.subject_id=s.id AND qq.teacher_id=? AND qq.deleted=0) AS qcount "
        "FROM subjects s JOIN teacher_subjects ts ON ts.subject_id=s.id WHERE ts.teacher_id=? ORDER BY s.name",
        (uid, uid))
    tests = q(
        "SELECT t.*, s.name AS subject, "
        "(SELECT COUNT(*) FROM attempts a WHERE a.test_id=t.id AND a.finished_at IS NOT NULL) AS done "
        "FROM tests t JOIN subjects s ON s.id=t.subject_id WHERE t.teacher_id=? ORDER BY t.id DESC LIMIT 5", (uid,))
    nstud = q("SELECT COUNT(*) c FROM users WHERE role='student' AND active=1", one=True)["c"]
    return render_template("t_home.html", subjects=subjects, tests=tests, nstud=nstud)


@app.route("/teacher/subject/<int:sid>/questions")
@role_required("teacher")
def t_questions(sid):
    subject = own_subject(sid)
    rows = q("SELECT * FROM questions WHERE teacher_id=? AND subject_id=? AND deleted=0 ORDER BY id DESC",
             (current_user()["id"], sid))
    return render_template("t_questions.html", subject=subject, rows=rows)


def read_question_form():
    f = request.form
    data = {k: f.get(k, "").strip() for k in ("text", "a", "b", "c", "d")}
    data["correct"] = f.get("correct", "").strip().upper()
    err = None
    if not data["text"]:
        err = "Savol matnini kiriting."
    elif not data["a"] or not data["b"]:
        err = "Kamida A va B variantlarini kiriting."
    elif data["correct"] not in LETTERS:
        err = "To'g'ri javobni tanlang."
    elif not data[data["correct"].lower()]:
        err = "To'g'ri javob sifatida bo'sh variant tanlangan."
    elif (data["d"] and not data["c"]):
        err = "D variantini kiritish uchun avval C ni to'ldiring."
    return data, err


@app.route("/teacher/subject/<int:sid>/questions/new", methods=["GET", "POST"])
@role_required("teacher")
def t_question_new(sid):
    subject = own_subject(sid)
    data = {}
    if request.method == "POST":
        data, err = read_question_form()
        if err:
            flash(err, "err")
        else:
            ex("INSERT INTO questions(teacher_id,subject_id,text,a,b,c,d,correct,created_at) VALUES(?,?,?,?,?,?,?,?,?)",
               (current_user()["id"], sid, data["text"], data["a"], data["b"], data["c"], data["d"],
                data["correct"], iso(now())))
            flash("Savol qo'shildi.", "ok")
            if request.form.get("another"):
                return redirect(url_for("t_question_new", sid=sid))
            return redirect(url_for("t_questions", sid=sid))
    return render_template("t_question_form.html", subject=subject, q=data, edit=False)


@app.route("/teacher/questions/<int:qid>/edit", methods=["GET", "POST"])
@role_required("teacher")
def t_question_edit(qid):
    row = own_question(qid)
    subject = own_subject(row["subject_id"])
    data = dict(row)
    if request.method == "POST":
        data, err = read_question_form()
        if err:
            flash(err, "err")
        else:
            ex("UPDATE questions SET text=?,a=?,b=?,c=?,d=?,correct=? WHERE id=?",
               (data["text"], data["a"], data["b"], data["c"], data["d"], data["correct"], qid))
            flash("Savol yangilandi.", "ok")
            return redirect(url_for("t_questions", sid=row["subject_id"]))
    return render_template("t_question_form.html", subject=subject, q=data, edit=True)


@app.route("/teacher/questions/<int:qid>/delete", methods=["POST"])
@role_required("teacher")
def t_question_delete(qid):
    row = own_question(qid)
    ex("UPDATE questions SET deleted=1 WHERE id=?", (qid,))
    flash("Savol o'chirildi.", "ok")
    return redirect(url_for("t_questions", sid=row["subject_id"]))


@app.route("/teacher/subject/<int:sid>/questions/clear", methods=["POST"])
@role_required("teacher")
def t_questions_clear(sid):
    own_subject(sid)
    ex("UPDATE questions SET deleted=1 WHERE teacher_id=? AND subject_id=?", (current_user()["id"], sid))
    flash("Barcha savollar o'chirildi.", "ok")
    return redirect(url_for("t_questions", sid=sid))


# ---- hujjatdan savollarni yuklash ----------------------------------------
def _norm_letter(v):
    v = (v or "").strip().upper()
    m = re.match(r"^([A-D])\b", v)
    if m:
        return m.group(1)
    # kirill A,B,V... ga yo'l qo'ymaymiz; faqat lotin
    return v[:1] if v[:1] in LETTERS else ""


def parse_xlsx(fileobj):
    wb = load_workbook(fileobj, data_only=True)
    ws = wb.active
    rows = []
    for i, r in enumerate(ws.iter_rows(min_row=1, values_only=True), start=1):
        cells = ["" if c is None else str(c).strip() for c in (list(r) + [None] * 6)[:6]]
        if i == 1 and _norm_letter(cells[5]) == "":
            continue  # sarlavha qatori
        if not any(cells):
            continue
        rows.append((i, cells))
    return rows


def parse_docx(fileobj):
    import docx
    d = docx.Document(fileobj)
    rows = []
    # 1) jadval: Savol | A | B | C | D | To'g'ri javob
    for t in d.tables:
        for i, r in enumerate(t.rows, start=1):
            cells = [c.text.strip() for c in r.cells][:6]
            cells += [""] * (6 - len(cells))
            if i == 1 and _norm_letter(cells[5]) == "":
                continue
            if any(cells):
                rows.append((len(rows) + 1, cells))
    if rows:
        return rows
    # 2) matn:  1. Savol / A) ... / B) ... / Javob: B
    lines = []
    for p in d.paragraphs:
        for ln in p.text.split("\n"):
            if ln.strip():
                lines.append(ln.strip())
    cur = None
    opt_re = re.compile(r"^\*?\s*([A-Da-d])\s*[\).:]\s*(.+)$")
    ans_re = re.compile(r"^(?:to['’`ʻ]?g['’`ʻ]?ri\s+)?(?:javob|answer|ans)\s*[:\-–]\s*([A-Da-d])\b", re.I)
    q_re = re.compile(r"^\d+\s*[\.\)]\s*(.+)$")
    out = []

    def flush():
        nonlocal cur
        if cur:
            out.append((len(out) + 1, [cur["t"], cur["A"], cur["B"], cur["C"], cur["D"], cur["ans"]]))
        cur = None

    for ln in lines:
        m = ans_re.match(ln)
        if m and cur:
            cur["ans"] = m.group(1).upper()
            continue
        m = opt_re.match(ln)
        if m and cur:
            L = m.group(1).upper()
            cur[L] = m.group(2).strip()
            if ln.lstrip().startswith("*"):
                cur["ans"] = L
            continue
        m = q_re.match(ln)
        if m:
            flush()
            cur = {"t": m.group(1).strip(), "A": "", "B": "", "C": "", "D": "", "ans": ""}
            continue
        if cur and not cur["A"]:
            cur["t"] += " " + ln  # ko'p qatorli savol matni
    flush()
    return out


@app.route("/teacher/subject/<int:sid>/upload", methods=["POST"])
@role_required("teacher")
def t_upload(sid):
    own_subject(sid)
    f = request.files.get("file")
    if not f or not f.filename:
        flash("Fayl tanlanmadi.", "err")
        return redirect(url_for("t_questions", sid=sid))
    ext = f.filename.rsplit(".", 1)[-1].lower()
    try:
        data = io.BytesIO(f.read())
        if ext == "xlsx":
            rows = parse_xlsx(data)
        elif ext == "docx":
            rows = parse_docx(data)
        else:
            flash("Faqat .xlsx yoki .docx fayl yuklang.", "err")
            return redirect(url_for("t_questions", sid=sid))
    except Exception as e:  # noqa
        flash("Faylni o'qib bo'lmadi: %s" % e, "err")
        return redirect(url_for("t_questions", sid=sid))
    added, errors = 0, []
    uid = current_user()["id"]
    for idx, c in rows:
        text, a, b, cc, d, corr = c
        corr = _norm_letter(corr)
        opts = {"A": a, "B": b, "C": cc, "D": d}
        if not text:
            errors.append("%d-qator: savol matni yo'q" % idx)
        elif not a or not b:
            errors.append("%d-qator: kamida A va B kerak" % idx)
        elif corr not in LETTERS or not opts[corr]:
            errors.append("%d-qator: to'g'ri javob noto'g'ri" % idx)
        elif d and not cc:
            errors.append("%d-qator: D bor, C yo'q" % idx)
        else:
            db().execute(
                "INSERT INTO questions(teacher_id,subject_id,text,a,b,c,d,correct,created_at) VALUES(?,?,?,?,?,?,?,?,?)",
                (uid, sid, text, a, b, cc, d, corr, iso(now())))
            added += 1
    db().commit()
    if added:
        flash("%d ta savol yuklandi." % added, "ok")
    if errors:
        flash("O'tkazib yuborilgan qatorlar (%d): %s%s" % (
            len(errors), "; ".join(errors[:8]), " ..." if len(errors) > 8 else ""), "err")
    if not added and not errors:
        flash("Fayldan savol topilmadi. Shablon formatini tekshiring.", "err")
    return redirect(url_for("t_questions", sid=sid))


@app.route("/teacher/template.xlsx")
@role_required("teacher")
def t_template():
    wb = Workbook()
    ws = wb.active
    ws.title = "Savollar"
    ws.append(["Savol", "A", "B", "C", "D", "To'g'ri javob (A/B/C/D)"])
    ws.append(["2 + 2 nechiga teng?", "3", "4", "5", "6", "B"])
    ws.append(["O'zbekiston poytaxti qaysi shahar?", "Samarqand", "Buxoro", "Toshkent", "Namangan", "C"])
    for c in ws[1]:
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor="2F5BEA")
    for i, w in enumerate([60, 22, 22, 22, 22, 24], start=1):
        ws.column_dimensions[get_column_letter(i)].width = w
    bio = io.BytesIO()
    wb.save(bio)
    bio.seek(0)
    return send_file(bio, as_attachment=True, download_name="savollar_shablon.xlsx",
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


# --------------------------------------------------------------------------
# O'QITUVCHI: o'quvchilar
# --------------------------------------------------------------------------
def class_list():
    return q("SELECT class_name, COUNT(*) AS n FROM users WHERE role='student' AND active=1 "
             "GROUP BY class_name ORDER BY class_name")


@app.route("/teacher/students")
@role_required("teacher")
def t_students():
    cls = request.args.get("class", "")
    classes = class_list()
    students = []
    if cls:
        students = q("SELECT * FROM users WHERE role='student' AND active=1 AND class_name=? ORDER BY full_name", (cls,))
    return render_template("t_students.html", classes=classes, cls=cls, students=students)


@app.route("/teacher/students/add", methods=["POST"])
@role_required("teacher")
def t_students_add():
    cls = request.form.get("class_name", "").strip()
    names = [re.sub(r"^\s*\d+[\.\)]\s*", "", ln).strip() for ln in request.form.get("names", "").splitlines()]
    names = [n for n in names if n]
    if not cls or not names:
        flash("Sinf nomini va kamida bitta ism-familiya kiriting.", "err")
        return redirect(url_for("t_students", **({"class": cls} if cls else {})))
    for n in names:
        pw = make_password()
        ex("INSERT INTO users(username,password_hash,pw_plain,role,full_name,class_name,created_at) VALUES(?,?,?,?,?,?,?)",
           (make_username(n), generate_password_hash(pw), pw, "student", n, cls, iso(now())))
    flash("%d ta o'quvchi qo'shildi. Login-parollar quyida." % len(names), "ok")
    return redirect(url_for("t_students", **{"class": cls}))


@app.route("/teacher/students/<int:sid>/reset", methods=["POST"])
@role_required("teacher")
def t_student_reset(sid):
    s = q("SELECT * FROM users WHERE id=? AND role='student'", (sid,), one=True)
    if not s:
        abort(404)
    pw = make_password()
    ex("UPDATE users SET password_hash=?, pw_plain=? WHERE id=?", (generate_password_hash(pw), pw, sid))
    flash("%s uchun yangi parol: %s" % (s["full_name"], pw), "ok")
    return redirect(url_for("t_students", **{"class": s["class_name"]}))


@app.route("/teacher/students/<int:sid>/delete", methods=["POST"])
@role_required("teacher")
def t_student_delete(sid):
    s = q("SELECT * FROM users WHERE id=? AND role='student'", (sid,), one=True)
    if not s:
        abort(404)
    ex("UPDATE users SET active=0 WHERE id=?", (sid,))
    flash("O'quvchi arxivlandi (natijalari saqlanadi).", "ok")
    return redirect(url_for("t_students", **{"class": s["class_name"]}))


def _autosize(ws, min_w=8, max_w=60):
    for col in ws.columns:
        width = max((len(str(c.value)) if c.value is not None else 0) for c in col)
        ws.column_dimensions[get_column_letter(col[0].column)].width = max(min_w, min(max_w, width + 3))


def _style_header(ws, row=1):
    for c in ws[row]:
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor="2F5BEA")
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.freeze_panes = ws.cell(row=row + 1, column=1)


def _send_wb(wb, name):
    bio = io.BytesIO()
    wb.save(bio)
    bio.seek(0)
    return send_file(bio, as_attachment=True, download_name=name,
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


@app.route("/teacher/students/credentials.xlsx")
@role_required("teacher")
def t_credentials():
    cls = request.args.get("class", "")
    rows = q("SELECT * FROM users WHERE role='student' AND active=1 AND class_name=? ORDER BY full_name", (cls,))
    if not rows:
        abort(404)
    wb = Workbook()
    ws = wb.active
    ws.title = cls[:30] or "Sinf"
    ws.append(["№", "F.I.Sh.", "Sinf", "Login", "Parol"])
    for i, r in enumerate(rows, start=1):
        ws.append([i, r["full_name"], r["class_name"], r["username"], r["pw_plain"] or "(o'zgartirilgan)"])
    _style_header(ws)
    _autosize(ws)
    return _send_wb(wb, "login_parollar_%s.xlsx" % (slugify(cls) or "sinf"))


# --------------------------------------------------------------------------
# O'QITUVCHI: testlar
# --------------------------------------------------------------------------
@app.route("/teacher/tests")
@role_required("teacher")
def t_tests():
    rows = q(
        "SELECT t.*, s.name AS subject, "
        "(SELECT COUNT(*) FROM attempts a WHERE a.test_id=t.id AND a.finished_at IS NOT NULL) AS done, "
        "(SELECT COUNT(*) FROM attempts a WHERE a.test_id=t.id AND a.finished_at IS NULL) AS running "
        "FROM tests t JOIN subjects s ON s.id=t.subject_id WHERE t.teacher_id=? ORDER BY t.id DESC",
        (current_user()["id"],))
    return render_template("t_tests.html", rows=rows)


@app.route("/teacher/tests/new", methods=["GET", "POST"])
@role_required("teacher")
def t_test_new():
    uid = current_user()["id"]
    subjects = my_subjects()
    classes = class_list()
    bank = {r["subject_id"]: r["n"] for r in q(
        "SELECT subject_id, COUNT(*) n FROM questions WHERE teacher_id=? AND deleted=0 GROUP BY subject_id", (uid,))}
    form = {}
    if request.method == "POST":
        f = request.form
        form = f
        title = f.get("title", "").strip()
        sid = as_int(f.get("subject_id"), 0, 0, 10**9)
        sel = [c for c in f.getlist("classes") if c]
        n = as_int(f.get("num_questions"), 0, 0, 500)
        minutes = as_int(f.get("minutes"), 0, 0, 600)
        err = None
        if not title:
            err = "Test nomini kiriting."
        elif not any(s["id"] == sid for s in subjects):
            err = "Fan tanlang."
        elif not sel:
            err = "Kamida bitta sinfni tanlang."
        elif n < 1:
            err = "Savollar soni kamida 1 bo'lsin."
        elif n > bank.get(sid, 0):
            err = "Savollar bankida faqat %d ta savol bor (siz %d ta tanladingiz)." % (bank.get(sid, 0), n)
        elif minutes < 1:
            err = "Vaqtni daqiqada kiriting (kamida 1)."
        start, end = parse_local_input(f.get("start_at")), parse_local_input(f.get("end_at"))
        if not err and start and end and end <= start:
            err = "Tugash vaqti boshlanish vaqtidan keyin bo'lishi kerak."
        if err:
            flash(err, "err")
        else:
            tid = ex(
                "INSERT INTO tests(teacher_id,subject_id,title,classes,num_questions,minutes,start_at,end_at,"
                "attempts_allowed,max_warnings,auto_finish,shuffle_q,shuffle_opt,show_result,active,created_at) "
                "VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,1,?)",
                (uid, sid, title, "|" + "|".join(sel) + "|", n, minutes, start, end,
                 as_int(f.get("attempts_allowed"), 1, 1, 10), as_int(f.get("max_warnings"), 3, 1, 10),
                 1 if f.get("auto_finish") else 0, 1 if f.get("shuffle_q") else 0,
                 1 if f.get("shuffle_opt") else 0, 1 if f.get("show_result") else 0, iso(now())))
            flash("Test yaratildi. O'quvchilar o'z akkauntidan ishlay oladi.", "ok")
            return redirect(url_for("t_tests"))
    return render_template("t_test_form.html", subjects=subjects, classes=classes, bank=bank, form=form,
                           sel=request.form.getlist("classes"))


@app.route("/teacher/tests/<int:tid>/toggle", methods=["POST"])
@role_required("teacher")
def t_test_toggle(tid):
    own_test(tid)
    ex("UPDATE tests SET active=1-active WHERE id=?", (tid,))
    return redirect(url_for("t_tests"))


@app.route("/teacher/tests/<int:tid>/delete", methods=["POST"])
@role_required("teacher")
def t_test_delete(tid):
    own_test(tid)
    ex("DELETE FROM tests WHERE id=?", (tid,))
    flash("Test va uning natijalari o'chirildi.", "ok")
    return redirect(url_for("t_tests"))


def results_for(test):
    for a in q("SELECT id FROM attempts WHERE test_id=? AND finished_at IS NULL", (test["id"],)):
        ensure_current(a["id"])
    rows = q("SELECT a.*, u.full_name, u.class_name FROM attempts a JOIN users u ON u.id=a.student_id "
             "WHERE a.test_id=? ORDER BY u.class_name, u.full_name, a.id", (test["id"],))
    out = []
    for a in rows:
        pct = round(100.0 * a["score"] / a["total"], 1) if a["finished_at"] and a["total"] else None
        dur = None
        if a["finished_at"]:
            dur = round((parse(a["finished_at"]) - parse(a["started_at"])).total_seconds() / 60.0, 1)
        out.append(dict(a=a, pct=pct, grade=grade(pct) if pct is not None else None, dur=dur,
                        status=STATUS_TXT.get(a["status"] if a["finished_at"] else None, "Jarayonda")))
    classes = [c for c in test["classes"].split("|") if c]
    taken = {a["student_id"] for a in rows}
    missing = [s for s in q(
        "SELECT * FROM users WHERE role='student' AND active=1 AND class_name IN (%s) ORDER BY class_name, full_name"
        % ",".join("?" * len(classes)), classes) if s["id"] not in taken] if classes else []
    return out, missing


@app.route("/teacher/tests/<int:tid>/results")
@role_required("teacher")
def t_results(tid):
    test = own_test(tid)
    res, missing = results_for(test)
    fin = [r for r in res if r["pct"] is not None]
    stats = None
    if fin:
        pcts = [r["pct"] for r in fin]
        stats = dict(n=len(fin), avg=round(sum(pcts) / len(pcts), 1), mx=max(pcts), mn=min(pcts),
                     warn=sum(1 for r in fin if r["a"]["warnings"] > 0))
    return render_template("t_results.html", test=test, res=res, missing=missing, stats=stats)


@app.route("/teacher/tests/<int:tid>/results.xlsx")
@role_required("teacher")
def t_results_xlsx(tid):
    test = own_test(tid)
    res, missing = results_for(test)
    wb = Workbook()
    ws = wb.active
    ws.title = "Natijalar"
    ws.append(["Test: %s" % test["title"], "Fan: %s" % test["subject"]])
    ws["A1"].font = Font(bold=True, size=13)
    ws.append([])
    ws.append(["№", "F.I.Sh.", "Sinf", "To'g'ri javoblar", "Savollar soni", "Foiz (%)", "Baho",
               "Ogohlantirishlar", "Boshlangan", "Tugagan", "Davomiyligi (daq)", "Holat"])
    hdr = ws.max_row
    for i, r in enumerate(res, start=1):
        a = r["a"]
        ws.append([i, a["full_name"], a["class_name"], a["score"], a["total"], r["pct"], r["grade"],
                   a["warnings"], fmt_dt(a["started_at"]), fmt_dt(a["finished_at"]), r["dur"], r["status"]])
    _style_header(ws, hdr)
    ws.freeze_panes = ws.cell(row=hdr + 1, column=1)
    _autosize(ws)
    ws2 = wb.create_sheet("Ishtirok etmaganlar")
    ws2.append(["№", "F.I.Sh.", "Sinf"])
    for i, s in enumerate(missing, start=1):
        ws2.append([i, s["full_name"], s["class_name"]])
    _style_header(ws2)
    _autosize(ws2)
    return _send_wb(wb, "natijalar_%s.xlsx" % (slugify(test["title"]) or "test"))


@app.route("/teacher/attempts/<int:aid>")
@role_required("teacher")
def t_attempt(aid):
    att = q("SELECT a.*, u.full_name, u.class_name FROM attempts a JOIN users u ON u.id=a.student_id WHERE a.id=?",
            (aid,), one=True)
    if not att:
        abort(404)
    test = own_test(att["test_id"])
    ensure_current(aid)
    att = q("SELECT a.*, u.full_name, u.class_name FROM attempts a JOIN users u ON u.id=a.student_id WHERE a.id=?",
            (aid,), one=True)
    items = review_items(att)
    viol = q("SELECT * FROM violations WHERE attempt_id=? ORDER BY id", (aid,))
    return render_template("t_attempt.html", test=test, att=att, items=items, viol=viol,
                           status=STATUS_TXT.get(att["status"] if att["finished_at"] else None))


def review_items(att):
    layout = json.loads(att["layout"])
    ans = {r["question_id"]: r["choice"] for r in q("SELECT * FROM answers WHERE attempt_id=?", (att["id"],))}
    items = []
    for it in layout:
        qr = q("SELECT * FROM questions WHERE id=?", (it["qid"],), one=True)
        if not qr:
            continue
        items.append(dict(q=qr, order=it["order"], chosen=ans.get(it["qid"]),
                          ok=ans.get(it["qid"]) == qr["correct"]))
    return items


# --------------------------------------------------------------------------
# O'QUVCHI
# --------------------------------------------------------------------------
def test_deadline(att, test):
    d = parse(att["started_at"]) + timedelta(minutes=test["minutes"])
    if test["end_at"]:
        d = min(d, parse(test["end_at"]))
    return d


def finish_attempt(aid, status):
    att = q("SELECT * FROM attempts WHERE id=?", (aid,), one=True)
    if not att or att["finished_at"]:
        return
    layout = json.loads(att["layout"])
    ans = {r["question_id"]: r["choice"] for r in q("SELECT * FROM answers WHERE attempt_id=?", (aid,))}
    correct = 0
    for it in layout:
        qr = q("SELECT correct FROM questions WHERE id=?", (it["qid"],), one=True)
        if qr and ans.get(it["qid"]) == qr["correct"]:
            correct += 1
    ex("UPDATE attempts SET finished_at=?, status=?, score=?, total=? WHERE id=? AND finished_at IS NULL",
       (iso(now()), status, correct, len(layout), aid))


def ensure_current(aid):
    """Vaqti o'tib ketgan tugallanmagan urinishni yopadi va yangilangan qatorni qaytaradi."""
    att = q("SELECT * FROM attempts WHERE id=?", (aid,), one=True)
    if att and not att["finished_at"]:
        test = q("SELECT * FROM tests WHERE id=?", (att["test_id"],), one=True)
        if now() > test_deadline(att, test) + timedelta(seconds=GRACE_SECONDS):
            finish_attempt(aid, "time")
            att = q("SELECT * FROM attempts WHERE id=?", (aid,), one=True)
    return att


def own_attempt(aid):
    att = q("SELECT * FROM attempts WHERE id=? AND student_id=?", (aid, current_user()["id"]), one=True)
    if not att:
        abort(404)
    return ensure_current(aid)


@app.route("/student")
@role_required("student")
def s_home():
    u = current_user()
    n = now()
    tests = q("SELECT t.*, s.name AS subject, us.full_name AS teacher FROM tests t "
              "JOIN subjects s ON s.id=t.subject_id JOIN users us ON us.id=t.teacher_id "
              "WHERE t.active=1 AND t.classes LIKE ? ORDER BY t.id DESC", ("%|" + (u["class_name"] or "") + "|%",))
    cards = []
    for t in tests:
        atts = [ensure_current(a["id"]) for a in q(
            "SELECT id FROM attempts WHERE test_id=? AND student_id=? ORDER BY id", (t["id"], u["id"]))]
        running = next((a for a in atts if not a["finished_at"]), None)
        done = [a for a in atts if a["finished_at"]]
        start, end = parse(t["start_at"]), parse(t["end_at"])
        if end and n > end and not running:
            state = "closed"
        elif start and n < start:
            state = "soon"
        elif running:
            state = "running"
        elif len(done) >= t["attempts_allowed"]:
            state = "done"
        else:
            state = "open"
        cards.append(dict(t=t, state=state, running=running, done=done))
    return render_template("s_home.html", cards=cards)


def build_layout(test):
    rows = q("SELECT * FROM questions WHERE teacher_id=? AND subject_id=? AND deleted=0",
             (test["teacher_id"], test["subject_id"]))
    random.shuffle(rows)
    rows = rows[:test["num_questions"]]
    if not test["shuffle_q"]:
        rows.sort(key=lambda r: r["id"])
    layout = []
    for r in rows:
        letters = [L for L in LETTERS if r[L.lower()]]
        if test["shuffle_opt"]:
            random.shuffle(letters)
        layout.append({"qid": r["id"], "order": letters})
    return layout


@app.route("/student/test/<int:tid>/start", methods=["POST"])
@role_required("student")
def s_start(tid):
    u = current_user()
    t = q("SELECT * FROM tests WHERE id=? AND active=1 AND classes LIKE ?",
          (tid, "%|" + (u["class_name"] or "") + "|%"), one=True)
    if not t:
        abort(404)
    n = now()
    for a in q("SELECT id FROM attempts WHERE test_id=? AND student_id=? AND finished_at IS NULL", (tid, u["id"])):
        a = ensure_current(a["id"])
        if not a["finished_at"]:
            return redirect(url_for("s_take", aid=a["id"]))
    if t["start_at"] and n < parse(t["start_at"]):
        flash("Test hali boshlanmagan.", "err")
        return redirect(url_for("s_home"))
    if t["end_at"] and n >= parse(t["end_at"]):
        flash("Test yopilgan.", "err")
        return redirect(url_for("s_home"))
    used = q("SELECT COUNT(*) c FROM attempts WHERE test_id=? AND student_id=?", (tid, u["id"]), one=True)["c"]
    if used >= t["attempts_allowed"]:
        flash("Urinishlar soni tugagan.", "err")
        return redirect(url_for("s_home"))
    layout = build_layout(t)
    if not layout:
        flash("Testda savollar yo'q. O'qituvchiga murojaat qiling.", "err")
        return redirect(url_for("s_home"))
    aid = ex("INSERT INTO attempts(test_id,student_id,started_at,layout,total) VALUES(?,?,?,?,?)",
             (tid, u["id"], iso(n), json.dumps(layout), len(layout)))
    return redirect(url_for("s_take", aid=aid))


@app.route("/student/attempt/<int:aid>")
@role_required("student")
def s_take(aid):
    att = own_attempt(aid)
    if att["finished_at"]:
        return redirect(url_for("s_result", aid=aid))
    test = q("SELECT t.*, s.name AS subject FROM tests t JOIN subjects s ON s.id=t.subject_id WHERE t.id=?",
             (att["test_id"],), one=True)
    layout = json.loads(att["layout"])
    ans = {r["question_id"]: r["choice"] for r in q("SELECT * FROM answers WHERE attempt_id=?", (aid,))}
    items = []
    for it in layout:
        qr = q("SELECT * FROM questions WHERE id=?", (it["qid"],), one=True)
        opts = [(L, qr[L.lower()]) for L in it["order"]]
        items.append(dict(id=qr["id"], text=qr["text"], opts=opts, chosen=ans.get(qr["id"])))
    remaining = int((test_deadline(att, test) - now()).total_seconds())
    return render_template("s_take.html", att=att, test=test, items=items, remaining=max(remaining, 0))


@app.route("/student/attempt/<int:aid>/result")
@role_required("student")
def s_result(aid):
    att = own_attempt(aid)
    if not att["finished_at"]:
        return redirect(url_for("s_take", aid=aid))
    test = q("SELECT t.*, s.name AS subject FROM tests t JOIN subjects s ON s.id=t.subject_id WHERE t.id=?",
             (att["test_id"],), one=True)
    pct = round(100.0 * att["score"] / att["total"], 1) if att["total"] else 0
    return render_template("s_result.html", att=att, test=test, pct=pct, grade=grade(pct),
                           status=STATUS_TXT.get(att["status"]))


# ---- API (test paytida) --------------------------------------------------
def api_attempt(aid):
    att = q("SELECT * FROM attempts WHERE id=? AND student_id=?", (aid, current_user()["id"]), one=True)
    if not att:
        return None, None
    att = ensure_current(aid)
    test = q("SELECT * FROM tests WHERE id=?", (att["test_id"],), one=True)
    return att, test


@app.route("/api/attempt/<int:aid>/answer", methods=["POST"])
@role_required("student")
def api_answer(aid):
    att, test = api_attempt(aid)
    if not att:
        return jsonify(error="not_found"), 404
    if att["finished_at"]:
        return jsonify(finished=True)
    data = request.get_json(silent=True) or {}
    qid, choice = data.get("qid"), str(data.get("choice", "")).upper()
    layout = json.loads(att["layout"])
    if choice not in LETTERS or not any(it["qid"] == qid and choice in it["order"] for it in layout):
        return jsonify(error="bad"), 400
    ex("INSERT INTO answers(attempt_id,question_id,choice) VALUES(?,?,?) "
       "ON CONFLICT(attempt_id,question_id) DO UPDATE SET choice=excluded.choice", (aid, qid, choice))
    return jsonify(ok=True)


@app.route("/api/attempt/<int:aid>/violation", methods=["POST"])
@role_required("student")
def api_violation(aid):
    att, test = api_attempt(aid)
    if not att:
        return jsonify(error="not_found"), 404
    if att["finished_at"]:
        return jsonify(finished=True, warnings=att["warnings"], max=test["max_warnings"])
    kind = (request.get_json(silent=True) or {}).get("kind", "blur")[:30]
    warnings = att["warnings"] + 1
    ex("UPDATE attempts SET warnings=? WHERE id=?", (warnings, aid))
    ex("INSERT INTO violations(attempt_id,at,kind) VALUES(?,?,?)", (aid, iso(now()), kind))
    finished = False
    if test["auto_finish"] and warnings >= test["max_warnings"]:
        finish_attempt(aid, "warnings")
        finished = True
    return jsonify(finished=finished, warnings=warnings, max=test["max_warnings"])


@app.route("/api/attempt/<int:aid>/finish", methods=["POST"])
@role_required("student")
def api_finish(aid):
    att, test = api_attempt(aid)
    if not att:
        return jsonify(error="not_found"), 404
    finish_attempt(aid, "finished")
    return jsonify(finished=True, url=url_for("s_result", aid=aid))


init_db()

if __name__ == "__main__":
    host = os.environ.get("HOST", "0.0.0.0")
    port = int(os.environ.get("PORT", "5000"))
    app.run(host=host, port=port, debug=False, threaded=True)
