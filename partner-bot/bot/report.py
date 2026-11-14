"""Загрузка сети из таблиц (основная + песочница новичков) и сборка отчётов.

Без Телеграма, чтобы проверить таблицу:
    python -m bot.report [--date 2026-11-04] [--file сеть.xlsx]
"""
from __future__ import annotations

import argparse
import logging
import re
from dataclasses import dataclass, field
from datetime import date, datetime

from .config import Config, load_config
from .network import (PartnerStats, compute_network, render_morning, render_weekly,
                      summary_rows)
from .partners import Directory
from .setup_sheet import (Sheet, ensure_structure, expand_history, refresh_stage_criteria,
                          repair_month_block, rewrite_guide, upgrade_time_stage, write_summary)
from .sheets import fetch_xlsx, google_session, parse_partners
from .workbook import parse_partner_tabs

log = logging.getLogger(__name__)


def sheet_url(sheet_id: str, gid: int | None = None) -> str:
    return f"https://docs.google.com/spreadsheets/d/{sheet_id}/edit" + (f"#gid={gid}" if gid is not None else "")


@dataclass
class Book:
    sheet_id: str
    sandbox: bool
    directory: Directory
    gids: dict[str, int] = field(default_factory=dict)
    tabs: list = field(default_factory=list)          # PartnerTab этой таблицы


@dataclass
class Network:
    today: date
    stats: list[PartnerStats] = field(default_factory=list)
    books: list[Book] = field(default_factory=list)
    errors: list[str] = field(default_factory=list)

    @property
    def directory(self) -> Directory:
        return Directory([p for b in self.books for p in b.directory.partners])

    def tab_url(self, partner: str) -> str:
        for b in self.books:
            if partner in b.gids:
                return sheet_url(b.sheet_id, b.gids[partner])
        return ""

    def main_url(self) -> str:
        main = next((b for b in self.books if not b.sandbox), None)
        return sheet_url(main.sheet_id) if main else ""


def load_network(cfg: Config, today: date, session=None, local_xlsx: bytes | None = None,
                 with_gids: bool = True) -> Network:
    net = Network(today)
    sources = [(cfg.sheet_id, False), (cfg.sandbox_id, True)]
    tabs = []
    for sid, sandbox in sources:
        if not sid and not (local_xlsx and not sandbox):
            continue
        try:
            if local_xlsx is not None and not sandbox:
                data = local_xlsx
            else:
                session = session or google_session(cfg.service_account_file)
                data, _ = fetch_xlsx(sid, cfg.service_account_file, session)
            book = Book(sid, sandbox, Directory.from_records(parse_partners(data) or []))
            if with_gids and session is not None and sid:
                book.gids = Sheet(session, sid).tabs()
            net.books.append(book)
            book.tabs = parse_partner_tabs(data, sandbox=sandbox)
            tabs += book.tabs
        except Exception as exc:
            log.exception("Не удалось прочитать таблицу %s", sid)
            what = "песочницу" if sandbox else "основную таблицу"
            net.errors.append(f"Не удалось прочитать {what}: {exc.__class__.__name__}: {exc}")
    net.stats = compute_network(tabs, today)
    return net


_CRITERIA_DONE: set[tuple[str, str]] = set()   # (таблица, вкладка) — формулы факта уже обновлены


def sync(cfg: Config, net: Network, session=None) -> list[str]:
    """Создаёт вкладки новым партнёрам и обновляет «Сводку». Возвращает предупреждения."""
    warnings = []
    session = session or google_session(cfg.service_account_file)
    stamp = datetime.now(cfg.tz).strftime("%d.%m.%Y %H:%M")
    for b in net.books:
        try:
            created = ensure_structure(session, b.sheet_id, b.directory, net.today, sandbox=b.sandbox)
            sh = Sheet(session, b.sheet_id)
            upgraded = 0
            for t in b.tabs:
                if t.name not in b.gids:
                    continue
                try:
                    if upgrade_time_stage(sh, t, b.gids[t.name], net.today):   # время и этап у клиентов
                        upgraded += 1
                        continue
                except Exception as exc:
                    log.exception("Не удалось добавить время и этап на вкладку %s", t.name)
                    warnings.append(f"Вкладка {t.name}: не добавились время и этап ({exc.__class__.__name__})")
                    continue
                try:
                    repair_month_block(sh, t, b.gids[t.name], net.today)
                except Exception:
                    log.exception("Не удалось собрать «Мой месяц» на вкладке %s", t.name)
                if (b.sheet_id, t.name) not in _CRITERIA_DONE:
                    try:
                        refresh_stage_criteria(sh, t)
                        _CRITERIA_DONE.add((b.sheet_id, t.name))
                    except Exception:
                        log.exception("Не удалось обновить формулы факта на вкладке %s", t.name)
                if t.history_full_soon():      # колонки касаний заканчиваются — добавляем новые
                    expand_history(sh, t, b.gids[t.name])
            if upgraded:                          # инструкция — под новые колонки
                rewrite_guide(sh, sandbox=b.sandbox)
            if created:
                log.info("Новые вкладки: %s", created)
            mine = [s for s in net.stats if s.sandbox == b.sandbox]
            write_summary(session, b.sheet_id, summary_rows(mine), stamp)
        except Exception as exc:
            log.exception("Не удалось обновить таблицу %s", b.sheet_id)
            code = getattr(getattr(exc, "response", None), "status_code", None)
            warnings.append("Боту нужен доступ «Редактор» к таблице." if code == 403
                            else f"Не удалось обновить таблицу: {exc.__class__.__name__}")
    return warnings


def weekly_snapshot(cfg: Config, net: Network, session=None) -> None:
    """Снимок недели на лист «Итоги недели» — из него готовят презентацию к пятнице."""
    from datetime import timedelta
    session = session or google_session(cfg.service_account_file)
    start, end = net.today - timedelta(days=7), net.today - timedelta(days=1)
    stamp = datetime.now(cfg.tz).strftime("%d.%m.%Y %H:%M")
    for b in net.books:
        mine = [s for s in net.stats if s.sandbox == b.sandbox]
        write_summary(session, b.sheet_id, summary_rows(mine), stamp, tab="Итоги недели",
                      title=f"Итоги недели {start:%d.%m}–{end:%d.%m} (снимок, часы за 7 дней)")


def morning_text(net: Network) -> str:
    from .onboarding import SANDBOX_DAYS, due_for_promotion
    text = render_morning(net.stats, net.today, net.main_url(), net.errors)
    due = [x for b in net.books if b.sandbox for x in due_for_promotion(b.directory, net.today)]
    if due:
        text += (f"\n\n🌱 <b>{SANDBOX_DAYS} дней в песочнице — пора переводить в основную таблицу</b>\n"
                 + "\n".join(f"{n} (с {d:%d.%m}) — /promote {n}" for n, d in due))
    return text


def weekly_text(net: Network) -> str:
    return render_weekly(net.stats, net.today)


def _strip_html(text: str) -> str:
    return re.sub(r"</?[a-z]+>", "", text).replace("&lt;", "<").replace("&gt;", ">").replace("&amp;", "&")


def main() -> None:
    p = argparse.ArgumentParser()
    p.add_argument("--file")
    p.add_argument("--date")
    args = p.parse_args()
    cfg = load_config()
    today = date.fromisoformat(args.date) if args.date else datetime.now(cfg.tz).date()
    local = open(args.file, "rb").read() if args.file else None
    net = load_network(cfg, today, local_xlsx=local, with_gids=not local)
    print(_strip_html(morning_text(net)))
    if today.weekday() == 0:
        print("\n" + _strip_html(weekly_text(net)))


if __name__ == "__main__":
    main()
