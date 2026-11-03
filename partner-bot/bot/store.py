"""Маленькая база бота (SQLite): кто писал боту и настройки, заданные из Telegram."""
from __future__ import annotations

import sqlite3
from pathlib import Path

SCHEMA = """
CREATE TABLE IF NOT EXISTS users (
    user_id INTEGER PRIMARY KEY,
    username TEXT,
    full_name TEXT,
    private_ok INTEGER DEFAULT 0   -- написал боту /start, значит можно писать в личку
);
CREATE TABLE IF NOT EXISTS status_seen (   -- последний виденный статус клиента
    book TEXT, tab TEXT, client TEXT, status TEXT, PRIMARY KEY (book, tab, client)
);
CREATE TABLE IF NOT EXISTS results (       -- засчитанные результаты (смена статуса)
    book TEXT, tab TEXT, client TEXT, cat TEXT, day TEXT
);
CREATE TABLE IF NOT EXISTS goals_archive ( -- цели и результаты прошлых месяцев
    book TEXT, tab TEXT, month TEXT, cat TEXT, goal REAL, result REAL
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

    def get_setting(self, key: str) -> str | None:
        row = self.db.execute("SELECT value FROM settings WHERE key=?", (key,)).fetchone()
        return row[0] if row else None

    def set_setting(self, key: str, value: str) -> None:
        with self.db:
            self.db.execute("INSERT INTO settings(key, value) VALUES (?, ?)"
                            " ON CONFLICT(key) DO UPDATE SET value=excluded.value", (key, value))

    # ---------- результаты месяца ----------

    def seen_statuses(self, book: str) -> dict[tuple[str, str], str]:
        rows = self.db.execute("SELECT tab, client, status FROM status_seen WHERE book=?", (book,)).fetchall()
        return {(t, c): st for t, c, st in rows}

    def save_statuses(self, book: str, seen: dict[tuple[str, str], str]) -> None:
        with self.db:
            self.db.execute("DELETE FROM status_seen WHERE book=?", (book,))
            self.db.executemany("INSERT INTO status_seen VALUES (?,?,?,?)",
                                [(book, t, c, st) for (t, c), st in seen.items()])

    def add_results(self, book: str, items: list[tuple[str, str, str, str]]) -> None:
        with self.db:
            self.db.executemany("INSERT INTO results VALUES (?,?,?,?,?)", [(book, *x) for x in items])

    def month_results(self, book: str, month: str) -> dict[tuple[str, str], int]:
        rows = self.db.execute("SELECT tab, cat, COUNT(*) FROM results WHERE book=? AND day LIKE ? "
                               "GROUP BY tab, cat", (book, f"{month}%")).fetchall()
        return {(t, c): n for t, c, n in rows}

    def archive_goals(self, book: str, month: str, rows: list[tuple[str, str, float, float]]) -> None:
        with self.db:
            self.db.executemany("INSERT INTO goals_archive VALUES (?,?,?,?,?,?)",
                                [(book, t, month, c, g, r) for t, c, g, r in rows])

    def has_result(self, book: str, tab: str, client: str, cat: str, month: str) -> bool:
        return self.db.execute("SELECT 1 FROM results WHERE book=? AND tab=? AND client=? AND cat=? "
                               "AND day LIKE ? LIMIT 1", (book, tab, client, cat, f"{month}%")).fetchone() is not None
