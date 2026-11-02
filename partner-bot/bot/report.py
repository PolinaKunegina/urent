"""Сборка отчёта. Можно запустить без Телеграма, чтобы проверить доступ к таблице:

    python -m bot.report                 # таблицы из SHEETS, вывод в консоль
    python -m bot.report --file x.xlsx   # локальный файл
    python -m bot.report --date 2026-11-03
"""
from __future__ import annotations

import argparse
import logging
import re
from datetime import date, datetime, timedelta

from .analytics import compute, render_hours, render_report
from .config import Config, SheetSource, load_config
from .hashtags import Store
from .sheets import fetch_xlsx, parse_funnel

log = logging.getLogger(__name__)


def build_report(cfg: Config, today: date, store: Store | None = None,
                 local_files: dict[str, bytes] | None = None) -> str:
    stats, errors = [], []
    sources = cfg.sheets or [SheetSource(label=k, file_id="") for k in (local_files or {})]
    for src in sources:
        try:
            if local_files and src.label in local_files:
                data = local_files[src.label]
            else:
                data = fetch_xlsx(src.file_id, cfg.service_account_file)
            leads = parse_funnel(data, cfg.sheet_tab)
            stats.append(compute(src.label, leads, today, cfg.exclude_status_regex))
        except Exception as exc:  # одна сломанная таблица не должна гасить весь отчёт
            log.exception("Не удалось прочитать таблицу %s", src.label)
            errors.append(f"{src.label}: не удалось прочитать таблицу ({exc.__class__.__name__}: {exc})")

    hours_block = None
    if store is not None:
        yesterday = today - timedelta(days=1)
        hours_block = render_hours(
            store.summary(yesterday, yesterday),
            store.summary(today - timedelta(days=7), yesterday),
            [s.label for s in cfg.sheets if not re.match(r"^Таблица \d+$", s.label)],
            yesterday,
        )
    return render_report(stats, today, hours_block, errors)


def _strip_html(text: str) -> str:
    return re.sub(r"</?[a-z]+>", "", text).replace("&lt;", "<").replace("&gt;", ">").replace("&amp;", "&")


def main() -> None:
    p = argparse.ArgumentParser()
    p.add_argument("--file", action="append", default=[], help="локальный xlsx (можно несколько)")
    p.add_argument("--date", help="дата отчёта YYYY-MM-DD, по умолчанию сегодня")
    p.add_argument("--html", action="store_true", help="не убирать HTML-разметку")
    args = p.parse_args()

    cfg = load_config()
    today = date.fromisoformat(args.date) if args.date else datetime.now(cfg.tz).date()
    local = None
    if args.file:
        local = {}
        for path in args.file:
            with open(path, "rb") as f:
                local[path.rsplit("/", 1)[-1].removesuffix(".xlsx")] = f.read()
        cfg.sheets = []
    text = build_report(cfg, today, store=None, local_files=local)
    print(text if args.html else _strip_html(text))


if __name__ == "__main__":
    main()
