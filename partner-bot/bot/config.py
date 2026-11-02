"""Настройки читаются из переменных окружения (файл .env на сервере)."""
from __future__ import annotations

import os
from dataclasses import dataclass, field
from datetime import time
from zoneinfo import ZoneInfo


@dataclass
class SheetSource:
    label: str      # подпись в отчёте, например фамилия партнёра
    file_id: str    # id файла на Google Drive (xlsx или Google Таблица)


@dataclass
class Config:
    bot_token: str = ""
    service_account_file: str = ""
    sheets: list[SheetSource] = field(default_factory=list)
    sheet_tab: str = "Воронка лидов"
    report_chat_ids: list[int] = field(default_factory=list)
    admin_ids: list[int] = field(default_factory=list)
    partners_chat_id: int | None = None
    report_time: time = time(9, 0)
    reminder_time: time | None = time(18, 0)
    tz: ZoneInfo = ZoneInfo("Europe/Moscow")
    exclude_status_regex: str = r"отказ"
    db_path: str = "data/bot.db"


def _ids(raw: str) -> list[int]:
    return [int(x) for x in raw.replace(";", ",").split(",") if x.strip()]


def _time(raw: str) -> time | None:
    raw = raw.strip()
    if not raw or raw.lower() in {"off", "нет", "-"}:
        return None
    h, m = raw.split(":")
    return time(int(h), int(m))


def parse_sheets(raw: str) -> list[SheetSource]:
    """SHEETS=Петрова=1abc...;Иванова=1def...  (подпись можно не указывать)."""
    result = []
    for i, chunk in enumerate(x.strip() for x in raw.split(";")):
        if not chunk:
            continue
        if "=" in chunk:
            label, file_id = (p.strip() for p in chunk.split("=", 1))
        else:
            label, file_id = f"Таблица {i + 1}", chunk
        result.append(SheetSource(label=label, file_id=file_id))
    return result


def load_config() -> Config:
    env = os.environ.get
    tz = ZoneInfo(env("TZ_NAME", "Europe/Moscow"))
    partners = env("PARTNERS_CHAT_ID", "").strip()
    return Config(
        bot_token=env("TELEGRAM_BOT_TOKEN", ""),
        service_account_file=env("GOOGLE_SERVICE_ACCOUNT_FILE", "secrets/service-account.json"),
        sheets=parse_sheets(env("SHEETS", "")),
        sheet_tab=env("SHEET_TAB", "Воронка лидов"),
        report_chat_ids=_ids(env("REPORT_CHAT_IDS", "")),
        admin_ids=_ids(env("ADMIN_IDS", "")),
        partners_chat_id=int(partners) if partners else None,
        report_time=_time(env("REPORT_TIME", "09:00")) or time(9, 0),
        reminder_time=_time(env("REMINDER_TIME", "18:00")),
        tz=tz,
        exclude_status_regex=env("EXCLUDE_STATUS_REGEX", r"отказ"),
        db_path=env("DB_PATH", "data/bot.db"),
    )
