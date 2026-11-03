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
