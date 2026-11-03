"""Создание и поддержка структуры таблицы: вкладки партнёров, «Оглавление», «Сводка».

Бот вызывает ensure_structure раз в час: новый партнёр, добавленный строкой на лист
«Партнёры», через час получает свою вкладку и ссылку в оглавлении. Существующие
вкладки бот не трогает — их заполняют партнёры.

Можно запустить вручную:  python -m bot.setup_sheet SHEET_ID [--prefill-from "Воронка лидов"]
"""
from __future__ import annotations

import logging
import time as _time
from datetime import date
from urllib.parse import quote

from .guide import GUIDE, GUIDE_TAB, guide_requests, guide_values
from .partners import Directory
from .workbook import (PLAN_COL0, PLAN_HEADERS, col_letter, fx, partner_tab_format_requests,
                       partner_tab_values, sep_for_locale)

log = logging.getLogger(__name__)
API = "https://sheets.googleapis.com/v4/spreadsheets"
SPRINT_END = date(2026, 12, 31)
TOC, SUMMARY, PARTNERS = "Оглавление", "Сводка", "Партнёры"


def service_email() -> str:
    """Email сервисного аккаунта бота — из ключа (для защиты заголовков)."""
    import json
    import os
    try:
        path = os.environ.get("GOOGLE_SERVICE_ACCOUNT_FILE", "secrets/service-account.json")
        return json.load(open(path))["client_email"]
    except Exception:
        return ""


class _Retrying:
    """Обёртка над сессией: при 429 (лимит Google 60 запросов в минуту) ждёт и повторяет."""
    def __init__(self, session):
        self._s = session

    def __getattr__(self, method):
        fn = getattr(self._s, method)

        def call(*a, **kw):
            for attempt in range(6):
                r = fn(*a, **kw)
                if getattr(r, "status_code", 200) != 429:
                    return r
                _time.sleep(min(60, 10 * (attempt + 1)))
            return r
        return call


class Sheet:
    def __init__(self, session, spreadsheet_id: str):
        self.s, self.id = _Retrying(session), spreadsheet_id

    def sep(self) -> str:
        if not hasattr(self, "_sep"):
            r = self.s.get(f"{API}/{self.id}", params={"fields": "properties.locale"}, timeout=30)
            r.raise_for_status()
            self._sep = sep_for_locale(r.json()["properties"].get("locale", ""))
        return self._sep

    def tabs(self) -> dict[str, int]:
        r = self.s.get(f"{API}/{self.id}", params={"fields": "sheets.properties(sheetId,title,index)"},
                       timeout=30)
        r.raise_for_status()
        return {sh["properties"]["title"]: sh["properties"]["sheetId"] for sh in r.json()["sheets"]}

    def batch(self, requests: list[dict]) -> dict:
        if not requests:
            return {}
        r = self.s.post(f"{API}/{self.id}:batchUpdate", json={"requests": requests}, timeout=60)
        r.raise_for_status()
        return r.json()

    def put(self, rng: str, values: list[list], raw: bool = False) -> None:
        r = self.s.put(f"{API}/{self.id}/values/{quote(rng)}",
                       params={"valueInputOption": "RAW" if raw else "USER_ENTERED"},
                       json={"range": rng, "values": values}, timeout=60)
        r.raise_for_status()

    def append(self, rng: str, values: list[list]) -> None:
        r = self.s.post(f"{API}/{self.id}/values/{quote(rng)}:append",
                        params={"valueInputOption": "USER_ENTERED", "insertDataOption": "INSERT_ROWS"},
                        json={"values": values}, timeout=30)
        r.raise_for_status()

    def get_values(self, rng: str) -> list[list]:
        r = self.s.get(f"{API}/{self.id}/values/{quote(rng)}", timeout=30)
        r.raise_for_status()
        return r.json().get("values", [])

    def clear(self, rng: str) -> None:
        r = self.s.post(f"{API}/{self.id}/values/{quote(rng)}:clear", json={}, timeout=30)
        r.raise_for_status()

    def add_tab(self, title: str, index: int | None = None, rows: int = 300, cols: int = 26) -> int:
        props = {"title": title, "gridProperties": {"rowCount": rows, "columnCount": cols}}
        if index is not None:
            props["index"] = index
        res = self.batch([{"addSheet": {"properties": props}}])
        return res["replies"][0]["addSheet"]["properties"]["sheetId"]


def create_partner_tab(sh: Sheet, name: str, start: date, clients: list[dict] | None = None,
                       days: dict | None = None, tab_title: str | None = None) -> int:
    title = tab_title or name
    sid = sh.add_tab(title, cols=PLAN_COL0 + len(PLAN_HEADERS) + 1)
    sh.put(f"'{title}'!A1", partner_tab_values(name, start, SPRINT_END, clients, sh.sep(), sid, days))
    sh.batch(partner_tab_format_requests(sid, sh.sep(), service_email()))
    log.info("Создана вкладка партнёра «%s» (%d клиентов)", name, len(clients or []))
    return sid


def write_toc(sh: Sheet, directory: Directory, tabs: dict[str, int], sandbox: bool = False) -> None:
    if TOC not in tabs:
        tabs[TOC] = sh.add_tab(TOC, index=0, rows=200, cols=6)
        sh.batch([{"updateSheetProperties": {"properties": {"sheetId": tabs[TOC], "gridProperties": {
            "frozenRowCount": 6}}, "fields": "gridProperties.frozenRowCount"}},
            {"updateDimensionProperties": {"range": {"sheetId": tabs[TOC], "dimension": "COLUMNS",
                                                     "startIndex": 0, "endIndex": 1},
                                           "properties": {"pixelSize": 230}, "fields": "pixelSize"}},
            {"repeatCell": {"range": {"sheetId": tabs[TOC], "startRowIndex": 0, "endRowIndex": 1},
                            "cell": {"userEnteredFormat": {"textFormat": {"bold": True, "fontSize": 14}}},
                            "fields": "userEnteredFormat.textFormat"}}])
    title = "Песочница новых партнёров" if sandbox else "Партнёрская сеть — CRM и планирование"
    rows = [
        [title],
        ["Найдите себя в списке и нажмите «Открыть →» — это ваша личная CRM."],
        [f"Впервые здесь? Начните с листа «{GUIDE_TAB}» — 5 минут чтения."],
        ["Каждый вечер до 19:00: даты следующих касаний, факт за сегодня, план на завтра."],
        [f"Общая статистика — на листе «{SUMMARY}»."],
        ["Партнёр", "Вкладка", "Telegram"],
    ]
    for p in directory.partners:
        gid = tabs.get(p.name)
        link = (fx(f'=HYPERLINK("#gid={gid}","Открыть →")', sh.sep()) if gid is not None
                else "вкладка появится в течение часа")
        rows.append([p.name, link, "@" + p.telegram.strip().lstrip("@") if p.username else ""])
    sh.clear(f"'{TOC}'!A1:C300")
    sh.put(f"'{TOC}'!A1", rows)


def ensure_structure(session, spreadsheet_id: str, directory: Directory, today: date,
                     prefill: dict[str, list[dict]] | None = None, sandbox: bool = False,
                     force_toc: bool = False) -> list[str]:
    """Создаёт недостающие вкладки партнёров, обновляет оглавление. Возвращает созданные."""
    sh = Sheet(session, spreadsheet_id)
    sh.batch([{"updateSpreadsheetProperties": {"properties": {"timeZone": "Europe/Moscow"},
                                               "fields": "timeZone"}}])  # «сегодня» по Москве
    tabs = sh.tabs()
    if PARTNERS not in tabs:   # новая таблица (например, песочница) — заводим справочник
        from .sheets import PARTNERS_HEADER
        tabs[PARTNERS] = sh.add_tab(PARTNERS, rows=200, cols=4)
        sh.put(f"'{PARTNERS}'!A1", [PARTNERS_HEADER], raw=True)
    created = []
    for p in directory.partners:
        if p.name in tabs:
            continue
        tabs[p.name] = create_partner_tab(sh, p.name, today, (prefill or {}).get(p.name))
        created.append(p.name)
    if GUIDE_TAB not in tabs:
        gid = sh.add_tab(GUIDE_TAB, index=1, rows=len(GUIDE) + 5, cols=2)
        sh.put(f"'{GUIDE_TAB}'!A1", guide_values(), raw=True)
        sh.batch(guide_requests(gid))
        tabs[GUIDE_TAB] = gid
    if created or force_toc or TOC not in tabs:
        write_toc(sh, directory, tabs, sandbox)
    return created


SUMMARY_HEADER = [
    "Партнёр", "Активность, %", "Клиентов в работе", "Без касания 7 дней", "% клиентов с касанием",
    "Дней с фактом (из 7)", "Привлечение 7 дн, ч", "Прогрев 7 дн, ч", "Закрытие 7 дн, ч",
    "Привлечение 7 дн, людей", "Прогрев 7 дн, людей", "Закрытие 7 дн, людей",
    "Привлечение с 1-го числа, ч", "Прогрев с 1-го числа, ч", "Закрытие с 1-го числа, ч",
]


def write_summary(session, spreadsheet_id: str, rows: list[list], updated: str) -> None:
    sh = Sheet(session, spreadsheet_id)
    tabs = sh.tabs()
    if SUMMARY not in tabs:
        sid = sh.add_tab(SUMMARY, index=1, rows=300, cols=len(SUMMARY_HEADER) + 1)
        sh.batch([
            {"updateSheetProperties": {"properties": {"sheetId": sid, "gridProperties": {
                "frozenRowCount": 3, "frozenColumnCount": 1}},
                "fields": "gridProperties.frozenRowCount,gridProperties.frozenColumnCount"}},
            {"repeatCell": {"range": {"sheetId": sid, "startRowIndex": 2, "endRowIndex": 3},
                            "cell": {"userEnteredFormat": {"textFormat": {"bold": True},
                                                           "wrapStrategy": "WRAP"}},
                            "fields": "userEnteredFormat.textFormat,userEnteredFormat.wrapStrategy"}},
            {"updateDimensionProperties": {"range": {"sheetId": sid, "dimension": "COLUMNS",
                                                     "startIndex": 0, "endIndex": 1},
                                           "properties": {"pixelSize": 200}, "fields": "pixelSize"}},
        ])
    sh.clear(f"'{SUMMARY}'!A1:{col_letter(len(SUMMARY_HEADER))}400")
    sh.put(f"'{SUMMARY}'!A1", [["Сводка по партнёрам (обновляет бот)", f"обновлено {updated}"], [],
                                SUMMARY_HEADER, *rows], raw=True)


def main() -> None:
    import argparse
    import os
    from .sheets import fetch_xlsx, google_session, parse_funnel, parse_partners

    logging.basicConfig(level=logging.INFO)
    ap = argparse.ArgumentParser()
    ap.add_argument("sheet_id")
    ap.add_argument("--prefill-from", help="лист со старой воронкой, клиенты разложатся по партнёрам")
    ap.add_argument("--sandbox", action="store_true")
    args = ap.parse_args()
    sa = os.environ.get("GOOGLE_SERVICE_ACCOUNT_FILE", "secrets/service-account.json")
    session = google_session(sa)
    data, _ = fetch_xlsx(args.sheet_id, sa, session)
    directory = Directory.from_records(parse_partners(data) or [])
    prefill = {}
    if args.prefill_from:
        f = parse_funnel(data, args.prefill_from)
        for l in f.leads:
            name, known = directory.resolve(l.source)
            if not known:
                continue
            prefill.setdefault(name, []).append({
                "name": l.name, "source": "мероприятие 29.10",
                "status": l.status, "description": l.description,
                "touches": [{"n": t.n, "date": t.when.isoformat() if t.when else t.raw, "fmt": t.fmt}
                            for t in l.touches]})
    created = ensure_structure(session, args.sheet_id, directory, date.today(), prefill, args.sandbox)
    print("Созданы вкладки:", ", ".join(created) or "нет новых")


if __name__ == "__main__":
    main()


def rebuild_partner_tabs(session, spreadsheet_id: str, today: date, only: list[str] | None = None) -> list[str]:
    """Пересобирает вкладки партнёров по новому макету, сохраняя клиентов и часы.
    Будущие даты из истории касаний становятся «Следующим шагом»."""
    from .sheets import fetch_xlsx
    from .workbook import parse_partner_tabs
    import os
    sa = os.environ.get("GOOGLE_SERVICE_ACCOUNT_FILE", "secrets/service-account.json")
    sh = Sheet(session, spreadsheet_id)
    data, _ = fetch_xlsx(spreadsheet_id, sa, session)
    tabs = sh.tabs()
    done = []
    for t in parse_partner_tabs(data):
        if only and t.name not in only:
            continue
        clients = []
        for l in t.clients:
            future = sorted((x for x in l.touches if x.when and x.when >= today), key=lambda x: x.when)
            nxt = l.next_date or (future[0].when if future else None)
            what = l.next_what or (future[0].fmt if future and not l.next_date else "")
            history = [x for x in l.touches if not (x.when and x.when >= today)]
            clients.append({
                "name": l.name, "source": l.source, "status": l.status, "description": l.description,
                "next_date": nxt.isoformat() if nxt else "", "next_what": what,
                "touches": [{"n": i + 1, "date": x.when.isoformat() if x.when else x.raw, "fmt": x.fmt}
                            for i, x in enumerate(history[:5])]})
        start = min(t.days) if t.days else today
        tmp = f"_новая {t.name}"
        new_id = create_partner_tab(sh, t.name, start, clients, t.days, tab_title=tmp)
        # новая вкладка готова — только теперь убираем старую и переименовываем
        sh.batch([{"deleteSheet": {"sheetId": tabs[t.name]}},
                  {"updateSheetProperties": {"properties": {"sheetId": new_id, "title": t.name},
                                             "fields": "title"}}])
        done.append(t.name)
    return done
