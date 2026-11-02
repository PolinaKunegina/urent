"""Сборка сводок. Можно запустить без Телеграма, чтобы проверить доступ к таблице:

    python -m bot.report                    # таблица из SHEET_ID, вывод в консоль
    python -m bot.report --file x.xlsx      # локальный файл
    python -m bot.report --date 2026-11-03 --weekly --evening
"""
from __future__ import annotations

import argparse
import logging
import re
from dataclasses import dataclass, field
from datetime import date, datetime, timedelta

from .analytics import (FunnelStats, compute, not_reported, render_hours, render_morning,
                        render_partner_evening, render_weekly)
from .config import Config, load_config
from .gwrite import SheetsWriter
from .hashtags import PROCESSES, Store
from .partners import UNKNOWN, Directory
from .sheets import (GSHEET_MIME, PARTNERS_TAB, Funnel, fetch_xlsx, google_session, parse_funnel,
                     parse_partners)

log = logging.getLogger(__name__)


@dataclass
class Snapshot:
    funnel: Funnel | None
    stats: FunnelStats | None
    directory: Directory
    is_gsheet: bool = False
    errors: list[str] = field(default_factory=list)
    partners_tab: bool = False   # есть ли в таблице лист «Партнёры»

    def expected_partners(self) -> list[str]:
        if self.directory.partners:
            return self.directory.names()
        if self.stats:
            return [p for p in self.stats.partners if p != UNKNOWN]
        return []


def load_snapshot(cfg: Config, today: date, local_xlsx: bytes | None = None,
                  session=None) -> Snapshot:
    directory = Directory.load(cfg.partners_file)
    partners_tab = False
    try:
        if local_xlsx is not None:
            data, mime = local_xlsx, "local"
        else:
            if not cfg.sheet_id:
                raise ValueError("не задан SHEET_ID")
            data, mime = fetch_xlsx(cfg.sheet_id, cfg.service_account_file, session)
        funnel = parse_funnel(data, cfg.sheet_tab)
        records = parse_partners(data)
        if records is not None:  # лист «Партнёры» главнее файла
            directory, partners_tab = Directory.from_records(records), True
        stats = compute(funnel.leads, today, cfg.exclude_status_regex, directory)
        return Snapshot(funnel, stats, directory, is_gsheet=(mime == GSHEET_MIME),
                        partners_tab=partners_tab)
    except Exception as exc:
        log.exception("Не удалось прочитать таблицу")
        return Snapshot(None, None, directory,
                        errors=[f"Не удалось прочитать таблицу: {exc.__class__.__name__}: {exc}"])


def write_back(cfg: Config, snap: Snapshot, store: Store | None, session=None) -> str | None:
    """Подсветка «нет плана» и лист с часами. Возвращает предупреждение или None."""
    if not cfg.write_back or snap.funnel is None:
        return None
    if not snap.is_gsheet:
        return ("Подсветка и лист с часами выключены: таблица загружена как .xlsx. "
                "Нужно сохранить её как Google Таблицу.")
    try:
        writer = SheetsWriter(session or google_session(cfg.service_account_file), cfg.sheet_id)
        writer.mark_plan(snap.funnel, snap.stats.plan, cfg.helper_header)
        if not snap.partners_tab:
            writer.create_partners_tab(PARTNERS_TAB, distinct_sources(snap.funnel))
        if store is not None:
            rows = [[d, p, PROCESSES[pr][0], h, note or "", f"@{u}" if u else ""]
                    for d, p, pr, h, note, u in store.all_entries()]
            writer.write_hours(cfg.hours_tab, rows)
    except Exception as exc:
        log.exception("Не удалось записать в таблицу")
        code = getattr(getattr(exc, "response", None), "status_code", None)
        if code == 403:
            return "Бот не может писать в таблицу: дайте сервисному аккаунту доступ «Редактор»."
        return f"Не удалось обновить подсветку/часы в таблице: {exc.__class__.__name__}"
    return None


def distinct_sources(funnel: Funnel) -> list[str]:
    """Варианты из колонки «Источник» для первичного заполнения листа «Партнёры»."""
    d, seen = Directory([]), []
    for l in funnel.leads:
        name, _ = d.resolve(l.source)
        if name != UNKNOWN and name not in seen:
            seen.append(name)
    return sorted(seen, key=str.lower)


def morning_report(cfg: Config, snap: Snapshot, store: Store | None, today: date,
                   warning: str | None = None) -> str:
    hours_block = None
    if store is not None:
        y = today - timedelta(days=1)
        hours_block = render_hours(store.summary(y, y), store.summary(today - timedelta(days=7), y),
                                   snap.expected_partners(), y)
    errors = list(snap.errors) + ([warning] if warning else [])
    if snap.stats is None:
        return "📊 Сводка не собрана.\n" + "\n".join(f"❗️ {x}" for x in errors)
    return render_morning(snap.stats, hours_block, errors)


def weekly_report(snap: Snapshot, store: Store | None, today: date) -> str | None:
    if snap.stats is None or snap.funnel is None:
        return None
    start, end = today - timedelta(days=7), today - timedelta(days=1)
    hours = store.summary(start, end) if store else {}
    return render_weekly(snap.funnel.leads, snap.stats, hours, start, end)


def evening_chat_text(snap: Snapshot, store: Store | None, today: date) -> str:
    text = ("Коллеги, напоминание 🌿 До 19:00 обновите таблицу воронки: статус, что происходит "
            "и <b>дату следующего касания</b> у каждого лида. И отчёт по часам с хэштегами "
            "#знакомство, #прогрев, #сделка (эталон — /template).")
    if store is not None:
        silent = not_reported(snap.expected_partners(), store.summary(today, today))
        if silent:
            text += "\n\n🔕 Сегодня ещё нет отчёта от: " + ", ".join(silent)
    return text


def evening_partner_texts(snap: Snapshot, today: date) -> dict[str, str]:
    """{имя партнёра: текст в личку}."""
    if snap.stats is None:
        return {}
    out = {}
    for name, p in snap.stats.partners.items():
        if name == UNKNOWN:
            continue
        t = render_partner_evening(p, today)
        if t:
            out[name] = t
    return out


def _strip_html(text: str) -> str:
    return re.sub(r"</?[a-z]+>", "", text).replace("&lt;", "<").replace("&gt;", ">").replace("&amp;", "&")


def main() -> None:
    p = argparse.ArgumentParser()
    p.add_argument("--file", help="локальный xlsx вместо таблицы на Drive")
    p.add_argument("--date", help="дата отчёта YYYY-MM-DD, по умолчанию сегодня")
    p.add_argument("--weekly", action="store_true", help="показать и недельную сводку")
    p.add_argument("--evening", action="store_true", help="показать вечерние сообщения партнёрам")
    p.add_argument("--html", action="store_true", help="не убирать HTML-разметку")
    args = p.parse_args()

    cfg = load_config()
    today = date.fromisoformat(args.date) if args.date else datetime.now(cfg.tz).date()
    local = open(args.file, "rb").read() if args.file else None
    snap = load_snapshot(cfg, today, local)
    texts = [morning_report(cfg, snap, None, today)]
    if args.weekly:
        texts.append(weekly_report(snap, None, today) or "")
    if args.evening:
        texts.append(evening_chat_text(snap, None, today))
        texts += [f"[{k}]\n{v}" for k, v in evening_partner_texts(snap, today).items()]
    out = "\n\n" + "─" * 30 + "\n\n"
    joined = out.join(texts)
    print(joined if args.html else _strip_html(joined))


if __name__ == "__main__":
    main()
