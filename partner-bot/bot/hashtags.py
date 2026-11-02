"""Учёт процессов из чата партнёров.

Партнёр пишет в чат сообщение по эталону:

    #отчет #план_Сергиенко
    #знакомство 2 ч — презентация экосистемы для Котовой
    #прогрев 1,5 ч — переписка с Соколовой, отправила кейсы
    #сделка 30 мин — Жуков: согласовали пакет, ждём оплату

Бот вытаскивает процесс и часы по каждой строке и складывает в SQLite.
Отредактированное сообщение пересчитывается целиком, поэтому опечатку в часах
можно просто исправить в Telegram.
"""
from __future__ import annotations

import re
import sqlite3
from dataclasses import dataclass
from datetime import date, datetime, timedelta
from pathlib import Path

# ключ процесса -> (название в отчёте, корни слов для хэштегов и ключевых слов)
PROCESSES: dict[str, tuple[str, tuple[str, ...]]] = {
    "intro": ("Знакомство с экосистемой", ("знакомств", "экосистем")),
    "warmup": ("Прогрев", ("прогрев",)),
    "deal": ("Сделки и закрытие", ("сделк", "закрыти")),
}
PROCESS_ORDER = ["intro", "warmup", "deal"]

TAG_RE = re.compile(r"#([\w]+)", re.UNICODE)
PARTNER_TAG_RE = re.compile(r"#(?:план|партнер|партнёр)_([\w]+)", re.IGNORECASE | re.UNICODE)
HOURS_RE = re.compile(r"(\d+(?:[.,]\d+)?)\s*(?:ч\b|час\w*|h\b)", re.IGNORECASE | re.UNICODE)
MINUTES_RE = re.compile(r"(\d+)\s*мин\w*", re.IGNORECASE | re.UNICODE)
HHMM_RE = re.compile(r"\b(\d{1,2}):(\d{2})\b")


def process_of(word: str) -> str | None:
    w = word.lower()
    for key, (_, roots) in PROCESSES.items():
        if any(w.startswith(r) for r in roots):
            return key
    return None


def hours_in(line: str) -> float:
    total = 0.0
    for m in HOURS_RE.finditer(line):
        total += float(m.group(1).replace(",", "."))
    for m in MINUTES_RE.finditer(line):
        total += int(m.group(1)) / 60
    if total == 0:
        for m in HHMM_RE.finditer(line):  # «1:30» — тоже часы
            total += int(m.group(1)) + int(m.group(2)) / 60
    return round(total, 2)


@dataclass
class Entry:
    process: str
    hours: float
    note: str
    people: int | None = None


@dataclass
class ParsedReport:
    partner: str | None
    entries: list[Entry]
    has_tags: bool          # были ли вообще процессные хэштеги

    @property
    def missing_hours(self) -> list[str]:
        return [e.process for e in self.entries if e.hours == 0]


def parse_message(text: str) -> ParsedReport | None:
    """None, если в сообщении нет ни одного процессного хэштега — это не отчёт."""
    if not text:
        return None
    tags = [t.lower() for t in TAG_RE.findall(text)]
    if not any(process_of(t) for t in tags):
        return None

    pm = PARTNER_TAG_RE.search(text)
    partner = pm.group(1).replace("_", " ") if pm else None

    entries: dict[str, Entry] = {}
    current: str | None = None
    for raw_line in text.splitlines():
        line = raw_line.strip()
        if not line:
            continue
        found = None
        for t in TAG_RE.findall(line):
            found = process_of(t) or found
        if not found:
            # строка вида «2. Прогрев — 1 ч» без решётки внутри отчёта
            head = re.sub(r"^[\d.)\-•\s]+", "", line).split(" ")[0]
            found = process_of(head)
        if found:
            current = found
        if current is None:
            continue
        h = hours_in(line)
        note = TAG_RE.sub("", line)
        for rx in (HOURS_RE, MINUTES_RE, HHMM_RE):
            note = rx.sub("", note)
        note = re.sub(r"\s*([—–-])\s*[,;]\s*", r" \1 ", note)
        note = re.sub(r"\s{2,}", " ", re.sub(r"^[\d.)\s]*", "", note)).strip(" -—–:,.")
        e = entries.setdefault(current, Entry(current, 0.0, ""))
        e.hours = round(e.hours + h, 2)
        if note and found:
            e.note = (e.note + "; " if e.note else "") + note
    return ParsedReport(partner=partner, entries=list(entries.values()), has_tags=True)


# ---------- хранилище ----------

SCHEMA = """
CREATE TABLE IF NOT EXISTS entries (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    chat_id INTEGER NOT NULL,
    message_id INTEGER NOT NULL,
    user_id INTEGER,
    partner TEXT NOT NULL,
    process TEXT NOT NULL,
    hours REAL NOT NULL,
    note TEXT,
    day TEXT NOT NULL,          -- YYYY-MM-DD по часовому поясу команды
    created_at TEXT NOT NULL
);
CREATE INDEX IF NOT EXISTS idx_entries_day ON entries(day);
CREATE INDEX IF NOT EXISTS idx_entries_msg ON entries(chat_id, message_id);
CREATE TABLE IF NOT EXISTS partners (
    user_id INTEGER PRIMARY KEY,
    partner TEXT NOT NULL
);
CREATE TABLE IF NOT EXISTS users (
    user_id INTEGER PRIMARY KEY,
    username TEXT,
    full_name TEXT,
    private_ok INTEGER DEFAULT 0   -- написал боту /start, значит можно писать в личку
);
CREATE TABLE IF NOT EXISTS pending (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    chat_id INTEGER, user_id INTEGER, partner TEXT, payload TEXT, created_at TEXT
);
CREATE TABLE IF NOT EXISTS settings (
    key TEXT PRIMARY KEY,
    value TEXT
);
"""


class Store:
    def __init__(self, path: str):
        Path(path).parent.mkdir(parents=True, exist_ok=True)
        self.db = sqlite3.connect(path, check_same_thread=False)
        self.db.executescript(SCHEMA)
        cols = {r[1] for r in self.db.execute("PRAGMA table_info(entries)")}
        with self.db:
            if "kind" not in cols:   # план/факт появились позже — докатываем схему
                self.db.execute("ALTER TABLE entries ADD COLUMN kind TEXT NOT NULL DEFAULT 'fact'")
            if "people" not in cols:
                self.db.execute("ALTER TABLE entries ADD COLUMN people INTEGER")

    def save(self, chat_id: int, message_id: int, user_id: int | None,
             partner: str, day: date, report: ParsedReport, kind: str = "fact") -> None:
        with self.db:
            self.db.execute("DELETE FROM entries WHERE chat_id=? AND message_id=?",
                            (chat_id, message_id))
            for e in report.entries:
                self.db.execute(
                    "INSERT INTO entries(chat_id,message_id,user_id,partner,process,hours,note,day,"
                    "created_at,kind,people) VALUES (?,?,?,?,?,?,?,?,?,?,?)",
                    (chat_id, message_id, user_id, partner, e.process, e.hours, e.note,
                     day.isoformat(), datetime.now().isoformat(timespec="seconds"), kind,
                     getattr(e, "people", None)),
                )

    def summary(self, start: date, end: date, kind: str = "fact") -> dict[str, dict[str, float]]:
        """{партнёр: {процесс: часы}} за период включительно."""
        rows = self.db.execute(
            "SELECT partner, process, SUM(hours) FROM entries WHERE day BETWEEN ? AND ? AND kind=?"
            " GROUP BY partner, process ORDER BY partner",
            (start.isoformat(), end.isoformat(), kind),
        ).fetchall()
        out: dict[str, dict[str, float]] = {}
        for partner, process, hours in rows:
            out.setdefault(partner, {})[process] = round(hours or 0, 2)
        return out

    def remember_partner(self, user_id: int, partner: str) -> None:
        with self.db:
            self.db.execute("INSERT INTO partners(user_id, partner) VALUES (?, ?)"
                            " ON CONFLICT(user_id) DO UPDATE SET partner=excluded.partner",
                            (user_id, partner))

    def partner_for(self, user_id: int) -> str | None:
        row = self.db.execute("SELECT partner FROM partners WHERE user_id=?", (user_id,)).fetchone()
        return row[0] if row else None

    def upsert_user(self, user_id: int, username: str | None, full_name: str,
                    private_ok: bool = False) -> None:
        with self.db:
            self.db.execute(
                "INSERT INTO users(user_id, username, full_name, private_ok) VALUES (?,?,?,?)"
                " ON CONFLICT(user_id) DO UPDATE SET username=excluded.username,"
                " full_name=excluded.full_name, private_ok=MAX(users.private_ok, excluded.private_ok)",
                (user_id, (username or "").lower(), full_name, int(private_ok)))

    def user_by_username(self, username: str) -> tuple[int, bool] | None:
        """(user_id, можно ли писать в личку)."""
        row = self.db.execute("SELECT user_id, private_ok FROM users WHERE username=?",
                              (username.lstrip("@").lower(),)).fetchone()
        return (row[0], bool(row[1])) if row else None

    def add_pending(self, chat_id: int, user_id: int, partner: str, payload: str) -> int:
        with self.db:
            cur = self.db.execute(
                "INSERT INTO pending(chat_id,user_id,partner,payload,created_at) VALUES (?,?,?,?,?)",
                (chat_id, user_id, partner, payload, datetime.now().isoformat(timespec="seconds")))
        return cur.lastrowid

    def pop_pending(self, pid: int) -> tuple | None:
        row = self.db.execute("SELECT chat_id,user_id,partner,payload FROM pending WHERE id=?",
                              (pid,)).fetchone()
        if row:
            with self.db:
                self.db.execute("DELETE FROM pending WHERE id=?", (pid,))
        return row

    def get_setting(self, key: str) -> str | None:
        row = self.db.execute("SELECT value FROM settings WHERE key=?", (key,)).fetchone()
        return row[0] if row else None

    def set_setting(self, key: str, value: str) -> None:
        with self.db:
            self.db.execute("INSERT INTO settings(key, value) VALUES (?, ?)"
                            " ON CONFLICT(key) DO UPDATE SET value=excluded.value", (key, value))

    def all_entries(self) -> list[tuple]:
        """Для выгрузки в Google Таблицу: (день, партнёр, процесс, часы, заметка, username)."""
        return self.db.execute(
            "SELECT e.day, e.partner, e.process, e.hours, e.note, COALESCE(u.username, ''), e.kind,"
            " e.people FROM entries e LEFT JOIN users u ON u.user_id = e.user_id"
            " ORDER BY e.day DESC, e.partner").fetchall()

    def partners_reported(self, day: date) -> set[str]:
        rows = self.db.execute("SELECT DISTINCT partner FROM entries WHERE day=? AND kind='fact'",
                               (day.isoformat(),)).fetchall()
        return {r[0] for r in rows}


def last_days(today: date, days: int) -> tuple[date, date]:
    return today - timedelta(days=days), today - timedelta(days=1)
