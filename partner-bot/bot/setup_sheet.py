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
from .workbook import (MONTH_COL, PLAN_COL0, PLAN_HEADERS, STATUSES, col_letter, link_cell,
                       month_block_requests, month_block_values, partner_links,
                       partner_tab_format_requests, partner_tab_values, sep_for_locale)

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
    tabs = sh.tabs()
    nav = {"toc": tabs.get(TOC), "guide": tabs.get(GUIDE_TAB)}
    sid = sh.add_tab(title, cols=PLAN_COL0 + len(PLAN_HEADERS) + 1)
    sh.put(f"'{title}'!A1", partner_tab_values(name, start, SPRINT_END, clients, sh.sep(), sid, days))
    sh.batch(partner_tab_format_requests(sid, sh.sep(), service_email())
             + month_block_requests(sid) + partner_links(sid, PLAN_COL0, nav["toc"], nav["guide"]))
    log.info("Создана вкладка партнёра «%s» (%d клиентов)", name, len(clients or []))
    return sid


def ensure_toc_tab(sh: Sheet, tabs: dict[str, int]) -> None:
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


def write_toc(sh: Sheet, directory: Directory, tabs: dict[str, int], sandbox: bool = False) -> None:
    ensure_toc_tab(sh, tabs)
    title = "Песочница новых партнёров" if sandbox else "Партнёрская сеть — CRM и планирование"
    rows = [
        [title],
        ["Найдите себя в списке и нажмите «Открыть →» — это ваша личная CRM."],
        [f"Впервые здесь? Начните с листа «{GUIDE_TAB}» — 5 минут чтения."],
        ["Каждый вечер до 19:00: у каждого клиента следующий шаг, факт за сегодня, план на завтра."],
        [f"Общая статистика — на листе «{SUMMARY}»."],
        ["Партнёр", "Вкладка", "Telegram"],
    ]
    for p in directory.partners:
        gid = tabs.get(p.name)
        link = "Открыть →" if gid is not None else "вкладка появится в течение часа"
        rows.append([p.name, link, "@" + p.telegram.strip().lstrip("@") if p.username else ""])
    sh.clear(f"'{TOC}'!A1:C300")
    sh.put(f"'{TOC}'!A1", rows)
    links = [link_cell(tabs[TOC], 6 + i, 1, [("Открыть →", f"#gid={tabs[p.name]}")])
             for i, p in enumerate(directory.partners) if tabs.get(p.name) is not None]
    if tabs.get(GUIDE_TAB) is not None:
        links.append(link_cell(tabs[TOC], 2, 1, [("📖 Открыть инструкцию", f"#gid={tabs[GUIDE_TAB]}")]))
    sh.batch(links)


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
    ensure_toc_tab(sh, tabs)               # оглавление и инструкция — до вкладок партнёров,
    if GUIDE_TAB not in tabs:              # чтобы в них сразу были ссылки на эти листы
        gid = sh.add_tab(GUIDE_TAB, index=1, rows=len(GUIDE) + 5, cols=2)
        sh.put(f"'{GUIDE_TAB}'!A1", guide_values(sandbox), raw=True)
        sh.batch(guide_requests(gid))
        tabs[GUIDE_TAB] = gid
        sh.batch([link_cell(gid, 0, 1, [("← Оглавление", f"#gid={tabs[TOC]}")])])
    created = []
    for p in directory.partners:
        if p.name in tabs:
            continue
        tabs[p.name] = create_partner_tab(sh, p.name, today, (prefill or {}).get(p.name))
        created.append(p.name)
    if created or force_toc or TOC not in tabs:
        write_toc(sh, directory, tabs, sandbox)
    return created


SUMMARY_HEADER = [
    "Партнёр", "Активность, %", "Клиентов в работе", "Без касания 7 дней", "% клиентов с касанием",
    "Дней с фактом (из 7)", "Привлечение 7 дн, ч", "Прогрев 7 дн, ч", "Закрытие 7 дн, ч",
    "Привлечение 7 дн, людей", "Прогрев 7 дн, людей", "Закрытие 7 дн, людей",
    "Привлечение с 1-го числа, ч", "Прогрев с 1-го числа, ч", "Закрытие с 1-го числа, ч",
    "Новые контакты: факт / цель", "Партнёры 95: факт / цель", "Партнёры 240: факт / цель",
    "Маленький продукт: факт / цель",
]


def write_summary(session, spreadsheet_id: str, rows: list[list], updated: str,
                  tab: str = None, title: str = "Сводка по партнёрам (обновляет бот)") -> None:
    SUMMARY = tab or globals()["SUMMARY"]
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
    sh.put(f"'{SUMMARY}'!A1", [[title, f"обновлено {updated}"], [],
                                SUMMARY_HEADER, *rows], raw=True)
    tabs = sh.tabs()
    nav = [(t, f"#gid={tabs[k]}") for t, k in (("← Оглавление", TOC), ("📖 Инструкция", GUIDE_TAB)) if k in tabs]
    sh.batch([link_cell(tabs[SUMMARY], 0, 2 + i, [x]) for i, x in enumerate(nav)])


def add_nav_links(sh: Sheet, partner_tabs: list) -> int:
    """Переставляет переходы на существующих вкладках (partner_tabs — PartnerTab с plan_col)."""
    tabs = sh.tabs()
    reqs, clear = [], []
    for t in partner_tabs:
        if t.name in tabs and t.plan_col is not None:
            reqs += partner_links(tabs[t.name], t.plan_col, tabs.get(TOC), tabs.get(GUIDE_TAB))
            clear.append(f"'{t.name}'!A3")
    if GUIDE_TAB in tabs and TOC in tabs:
        reqs.append(link_cell(tabs[GUIDE_TAB], 0, 1, [("← Оглавление", f"#gid={tabs[TOC]}")]))
    for rng in clear:
        sh.clear(rng)
    sh.batch(reqs)
    return len(reqs)


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
    """Пересобирает вкладки партнёров по новому макету, сохраняя клиентов, историю и часы."""
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
        for l in t.clients:     # история касаний переносится целиком, как есть
            clients.append({
                "name": l.name, "source": l.source, "status": l.status, "description": l.description,
                "next_date": l.next_date.isoformat() if l.next_date else "", "next_what": l.next_what,
                "touches": [{"n": x.n, "date": x.when.isoformat() if x.when else x.raw, "fmt": x.fmt,
                             "minutes": x.minutes} for x in l.touches]})
        start = min(t.days) if t.days else today
        tmp = f"_новая {t.name}"
        new_id = create_partner_tab(sh, t.name, start, clients, t.days, tab_title=tmp)
        # новая вкладка готова — только теперь убираем старую и переименовываем
        sh.batch([{"deleteSheet": {"sheetId": tabs[t.name]}},
                  {"updateSheetProperties": {"properties": {"sheetId": new_id, "title": t.name},
                                             "fields": "title"}}])
        done.append(t.name)
    return done


HISTORY_ADD = 5   # сколько касаний добавлять за раз


def expand_history(sh: Sheet, tab, gid: int, add: int = HISTORY_ADD) -> None:
    """Добавляет касания («Дата / Формат (этап) / Мин») после последнего; блок часов сдвигается вправо.

    Вставка идёт прямо перед разделителем, поэтому диапазоны формул факта расширяются сами."""
    from .workbook import FIRST_ROW, HEADER_ROW, CLIENT_ROWS, COL_NEXT_DATE, touch_headers
    last = tab.touch_cols[-1]
    triple = len(last) > 3 and last[3] is not None
    step = 3 if triple else 2
    ins = (last[3] if triple else last[2]) + 1
    heads = []
    for k in range(add):
        heads += touch_headers(last[0] + k + 1)[:step] if triple else [f"Дата {last[0] + k + 1}-го касания", "Формат"]
    reqs = [{"insertDimension": {"range": {"sheetId": gid, "dimension": "COLUMNS",
                                           "startIndex": ins, "endIndex": ins + step * add},
                                 "inheritFromBefore": True}},
            _text_row(gid, HEADER_ROW - 1, ins, heads)]
    for k in range(add):
        c = ins + step * k
        rng = {"sheetId": gid, "startRowIndex": FIRST_ROW - 1, "endRowIndex": HEADER_ROW + CLIENT_ROWS,
               "startColumnIndex": c, "endColumnIndex": c + 1}
        reqs.append({"repeatCell": {"range": rng, "cell": {"userEnteredFormat": {
            "numberFormat": {"type": "DATE", "pattern": "dd.mm.yyyy"}}}, "fields": "userEnteredFormat.numberFormat"}})
        reqs.append({"setDataValidation": {"range": rng, "rule": {"condition": {"type": "DATE_IS_VALID"},
                                                                  "strict": True}}})
        if triple:
            mrng = dict(rng, startColumnIndex=c + 2, endColumnIndex=c + 3)
            reqs.append({"setDataValidation": {"range": mrng, "rule": {"condition": {
                "type": "NUMBER_GREATER_THAN_EQ", "values": [{"userEnteredValue": "0"}]}, "strict": True}}})
    if tab.plan_col is not None:   # ссылка «→ к моим часам» должна вести на новый адрес
        new_plan = tab.plan_col + step * add
        reqs.append(link_cell(gid, 0, COL_NEXT_DATE,
                              [("→ к моим часам", f"#gid={gid}&range={col_letter(new_plan)}{HEADER_ROW}")],
                              size=12, bold=True))
    sh.batch(reqs)
    log.info("Вкладка «%s»: добавлено %d касаний", tab.name, add)


def _text_row(gid: int, row: int, col: int, texts: list) -> dict:
    """updateCells: строка значений; строки, начинающиеся с «=», уходят формулами
    (formulaValue разбирается в локали таблицы — для ru_RU разделитель «;»)."""
    def val(x):
        if x is None or x == "":
            return {}
        if isinstance(x, (int, float)):
            return {"userEnteredValue": {"numberValue": x}}
        if str(x).startswith("="):
            return {"userEnteredValue": {"formulaValue": x}}
        return {"userEnteredValue": {"stringValue": str(x)}}
    return {"updateCells": {"rows": [{"values": [val(x) for x in texts]}], "fields": "userEnteredValue",
                            "start": {"sheetId": gid, "rowIndex": row, "columnIndex": col}}}


def needs_time_stage(tab) -> bool:
    """Вкладка старого макета: пары «дата / формат», нет «Время, мин», «что» в колонке E."""
    from .workbook import COL_NEXT_WHAT
    return (tab.time_col is None and tab.next_what_col == COL_NEXT_WHAT and tab.plan_col is not None
            and bool(tab.touch_cols) and all(len(t) < 4 or t[3] is None for t in tab.touch_cols))


def time_stage_upgrade_requests(tab, gid: int, today: date, service_email: str = "",
                                toc_gid: int | None = None, guide_gid: int | None = None,
                                sep: str = ";") -> list[dict]:
    """Одним атомарным batchUpdate: «Время, мин» и «Этап» после «Следующий шаг: что», колонка «Мин»
    после каждого «Формата», формулы факта в блоке часов. Факт, внесённый руками в прошлые дни,
    остаётся как есть; сегодня и дальше — формулы."""
    from .workbook import (COL_TIME, FMT_HEADER, HEADER_ROW, HINT_CLIENTS, HINT_HOURS, MIN_HEADER,
                           P_FACT, STAGE_HEADER, TIME_HEADER, fact_formulas, time_stage_requests)
    pairs = sorted(tab.touch_cols, key=lambda t: t[1])
    reqs = [{"insertDimension": {"range": {"sheetId": gid, "dimension": "COLUMNS",
                                           "startIndex": COL_TIME, "endIndex": COL_TIME + 2},
                                 "inheritFromBefore": True}}]
    for _, d, f, *_ in reversed(pairs):          # справа налево — левые индексы не сдвигаются
        reqs.append({"insertDimension": {"range": {"sheetId": gid, "dimension": "COLUMNS",
                                                   "startIndex": f + 3, "endIndex": f + 4},
                                         "inheritFromBefore": True}})
    new_touch = [(n, d + 2 + k, f + 2 + k, f + 3 + k) for k, (n, d, f, *_) in enumerate(pairs)]
    plan = tab.plan_col + 2 + len(pairs)
    reqs.append(_text_row(gid, HEADER_ROW - 1, COL_TIME, [TIME_HEADER, STAGE_HEADER]))
    for _, _, f, m in new_touch:
        reqs.append(_text_row(gid, HEADER_ROW - 1, f, [FMT_HEADER, MIN_HEADER]))
    reqs.append(_text_row(gid, 1, 0, [HINT_CLIENTS]))
    reqs.append(_text_row(gid, 1, plan, [HINT_HOURS]))
    hist0 = new_touch[0][1]
    for dp in sorted(tab.days.values(), key=lambda x: x.day):
        if not dp.row or (dp.day < today and (dp.fact or dp.people)):
            continue
        reqs.append(_text_row(gid, dp.row - 1, plan + P_FACT, fact_formulas(dp.row, plan, hist0, sep)))
    reqs += time_stage_requests(gid, plan, new_touch, service_email, sep)
    reqs += partner_links(gid, plan, toc_gid, guide_gid)
    return reqs


def month_repair_requests(tab, gid: int, today: date, sep: str, protected: list[dict],
                          width: int) -> list[dict]:
    """Собирает «Мой месяц» обратно, если в блок попала чужая колонка: подписи, цели и результаты
    встают подряд от подписи блока, лишние ячейки справа очищаются, шапка защищается заново."""
    from .workbook import RESULT_CATS, header_protection_ranges
    c0 = tab.month_col
    n = len(RESULT_CATS)
    vals = month_block_values(today, sep)
    vals[1][1:1 + n] = [tab.goals.get(k, "") for k, _, _ in RESULT_CATS]
    vals[2][1:1 + n] = [tab.results.get(k, 0) for k, _, _ in RESULT_CATS]
    reqs = [_text_row(gid, r, c0, row) for r, row in enumerate(vals)]
    reqs += month_block_requests(gid)
    extra0, extra1 = c0 + n + 2, max(tab.month_end, c0 + n + 2)
    if extra1 > extra0:              # хвост старого блока справа — пустой и без заливки
        reqs.append({"updateCells": {"range": {"sheetId": gid, "startRowIndex": 0, "endRowIndex": 3,
                                               "startColumnIndex": extra0, "endColumnIndex": extra1},
                                     "fields": "userEnteredValue,userEnteredFormat.backgroundColor"}})
    for c in range(c0 + 1, c0 + n + 2):   # узкие колонки (минуты) под подписями блока — пошире
        reqs.append({"updateDimensionProperties": {"range": {"sheetId": gid, "dimension": "COLUMNS",
                                                             "startIndex": c, "endIndex": c + 1},
                                                   "properties": {"pixelSize": 100}, "fields": "pixelSize"}})
    editors = None
    for pr in protected:
        rg = pr.get("range", {})
        if pr.get("description") == "bot:заголовки" and rg.get("startRowIndex", 0) < 4 \
                and rg.get("endRowIndex", 99) <= 4:
            reqs.append({"deleteProtectedRange": {"protectedRangeId": pr["protectedRangeId"]}})
            editors = editors or pr.get("editors")
    if editors:
        reqs += [{"addProtectedRange": {"protectedRange": {
            "range": part, "description": "bot:заголовки", "warningOnly": False, "editors": editors}}}
            for part in header_protection_ranges(gid, width)]
    return reqs


def repair_month_block(sh: Sheet, tab, gid: int, today: date) -> bool:
    if tab.month_col is None or tab.month_ok:
        return False
    from .workbook import PLAN_HEADERS
    r = sh.s.get(f"{API}/{sh.id}", params={"fields": "sheets(properties.sheetId,protectedRanges)"}, timeout=30)
    r.raise_for_status()
    protected = next((x.get("protectedRanges", []) for x in r.json().get("sheets", [])
                      if x["properties"]["sheetId"] == gid), [])
    width = (tab.plan_col or 0) + len(PLAN_HEADERS)
    sh.batch(month_repair_requests(tab, gid, today, sh.sep(), protected, width))
    log.info("Вкладка «%s»: блок «Мой месяц» собран заново", tab.name)
    return True


def upgrade_time_stage(sh: Sheet, tab, gid: int, today: date) -> bool:
    if not needs_time_stage(tab):
        return False
    tabs = sh.tabs()
    sh.batch(time_stage_upgrade_requests(tab, gid, today, service_email(), tabs.get(TOC), tabs.get(GUIDE_TAB),
                                         sh.sep()))
    log.info("Вкладка «%s»: добавлены время и этап, факт считается формулами", tab.name)
    return True


def add_month_block(sh: Sheet, tab, gid: int, today: date) -> None:
    """Существующей вкладке: блок «Мой месяц», новый список статусов, цели открыты для правки."""
    from .workbook import COL_STATUS, FIRST_ROW, HEADER_ROW, CLIENT_ROWS
    sh.put(f"'{tab.name}'!{col_letter(MONTH_COL)}1", month_block_values(today, sh.sep()))
    sh.put(f"'{tab.name}'!A2", [["Следующий шаг сам попадает в историю. Красный — нет шага на 7 дней."]], raw=True)
    reqs = month_block_requests(gid) + [{"setDataValidation": {"range": {
        "sheetId": gid, "startRowIndex": FIRST_ROW - 1, "endRowIndex": HEADER_ROW + CLIENT_ROWS,
        "startColumnIndex": COL_STATUS, "endColumnIndex": COL_STATUS + 1}, "rule": {
        "condition": {"type": "ONE_OF_LIST", "values": [{"userEnteredValue": x} for x in STATUSES]},
        "showCustomUi": True, "strict": False}}}]
    from .workbook import header_protection_ranges
    r = sh.s.get(f"{API}/{sh.id}", params={"fields": "sheets(properties.sheetId,protectedRanges(protectedRangeId,description,range,editors))"},
                 timeout=30)
    r.raise_for_status()
    for x in r.json().get("sheets", []):
        if x["properties"]["sheetId"] != gid:
            continue
        for pr in x.get("protectedRanges", []):
            rg = pr["range"]
            if pr.get("description") == "bot:заголовки" and rg.get("startRowIndex", 0) == 0 \
                    and rg.get("endRowIndex") == HEADER_ROW:      # сплошная защита шапки — режем вокруг целей
                width = rg.get("endColumnIndex")
                reqs.append({"deleteProtectedRange": {"protectedRangeId": pr["protectedRangeId"]}})
                for part in header_protection_ranges(gid, width):
                    reqs.append({"addProtectedRange": {"protectedRange": {
                        "range": part, "description": "bot:заголовки", "warningOnly": False,
                        "editors": pr.get("editors", {})}}})
    sh.batch(reqs)


def rewrite_guide(sh: Sheet, sandbox: bool = False) -> int:
    """Обновляет текст «Инструкции» на месте: лист не пересоздаётся, ссылки на него живы."""
    tabs = sh.tabs()
    if GUIDE_TAB not in tabs:
        ensure_toc_tab(sh, tabs)
        gid = sh.add_tab(GUIDE_TAB, index=1, rows=len(GUIDE) + 5, cols=2)
    else:
        gid = tabs[GUIDE_TAB]
        sh.batch([{"updateSheetProperties": {"properties": {"sheetId": gid, "gridProperties": {
            "rowCount": len(GUIDE) + 5, "columnCount": 2}}, "fields": "gridProperties(rowCount,columnCount)"}}])
        sh.clear(f"'{GUIDE_TAB}'!A1:B{len(GUIDE) + 5}")
    sh.put(f"'{GUIDE_TAB}'!A1", guide_values(sandbox), raw=True)
    reqs = guide_requests(gid)
    if TOC in sh.tabs():
        reqs.append(link_cell(gid, 0, 1, [("← Оглавление", f"#gid={sh.tabs()[TOC]}")]))
    sh.batch(reqs)
    return gid
