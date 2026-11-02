"""Настройки читаются из переменных окружения (файл .env на сервере)."""
from __future__ import annotations

import os
import re
from dataclasses import dataclass, field
from datetime import time
from zoneinfo import ZoneInfo


@dataclass
class Config:
    bot_token: str = ""
    service_account_file: str = "secrets/service-account.json"
    sheet_id: str = ""                       # общая таблица воронки на Google Drive
    sheet_tab: str = "Воронка лидов"
    managers: list[str] = field(default_factory=list)   # @username или числовые id
    partners_file: str = "data/partners.json"
    partners_chat_id: int | None = None      # можно не задавать: /bindchat в чате партнёров
    report_time: time = time(8, 0)
    reminder_time: time | None = time(18, 30)
    weekly_day: int | None = 0               # 0 = понедельник, None = без недельной сводки
    tz: ZoneInfo = ZoneInfo("Europe/Moscow")
    exclude_status_regex: str = r"отказ"
    db_path: str = "data/bot.db"
    write_back: bool = True                  # подсветка и лист с часами (нужна Google Таблица)
    hours_tab: str = "Часы из чата"
    helper_header: str = "След. касание (бот)"


def parse_managers(raw: str) -> list[str]:
    return [x.strip().lstrip("@").lower() for x in re.split(r"[,;\s]+", raw) if x.strip()]


def _time(raw: str) -> time | None:
    raw = raw.strip()
    if not raw or raw.lower() in {"off", "нет", "-"}:
        return None
    h, m = raw.split(":")
    return time(int(h), int(m))


def _bool(raw: str) -> bool:
    return raw.strip().lower() not in {"0", "off", "no", "false", "нет", ""}


def load_config() -> Config:
    env = os.environ.get
    partners_chat = env("PARTNERS_CHAT_ID", "").strip()
    weekly = env("WEEKLY_DAY", "0").strip()
    return Config(
        bot_token=env("TELEGRAM_BOT_TOKEN", ""),
        service_account_file=env("GOOGLE_SERVICE_ACCOUNT_FILE", "secrets/service-account.json"),
        sheet_id=env("SHEET_ID", "").strip(),
        sheet_tab=env("SHEET_TAB", "Воронка лидов"),
        managers=parse_managers(env("MANAGERS", "")),
        partners_file=env("PARTNERS_FILE", "data/partners.json"),
        partners_chat_id=int(partners_chat) if partners_chat else None,
        report_time=_time(env("REPORT_TIME", "08:00")) or time(8, 0),
        reminder_time=_time(env("REMINDER_TIME", "18:30")),
        weekly_day=None if weekly.lower() in {"off", "-", ""} else int(weekly),
        tz=ZoneInfo(env("TZ_NAME", "Europe/Moscow")),
        exclude_status_regex=env("EXCLUDE_STATUS_REGEX", r"отказ"),
        db_path=env("DB_PATH", "data/bot.db"),
        write_back=_bool(env("WRITE_BACK", "on")),
        hours_tab=env("HOURS_TAB", "Часы из чата"),
    )
