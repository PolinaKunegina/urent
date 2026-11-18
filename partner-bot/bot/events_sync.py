"""Мероприятия в таблице: создание листов, колонки на вкладках, журнал, «События», гости.

Включается командой /events on (настройка events=on). Пока выключено — ничего не трогает.
"""
from __future__ import annotations

import csv
import io
import logging
from datetime import date

from openpyxl import load_workbook

from .events import (EVENT_COLS, EVENTS_TAB, GUESTS_TAB, LISTS_TAB, REPORT_TAB, build_report,
                     event_columns_requests, events_tab_requests, events_tab_values, lists_values,
                     parse_event_rows, parse_sovetnik_export, parse_tracker, report_format_requests,
                     report_rules, report_values, update_journal, _list_rule, _rng, _fill, _width, HEAD, YELLOW)
from .setup_sheet import API, GUIDE_TAB, TOC, Sheet, _text_row, month_repair_requests, service_email
from .workbook import HEADER_ROW, PLAN_HEADERS, partner_links

log = logging.getLogger(__name__)
DRIVE = "https://www.googleapis.com/drive/v3/files"
GUESTS_HEAD = ["Мероприятие", "Имя и фамилия гостя", "Контакт", "Кто привёл", "Был", "Готово (бот)"]
REPORT_MARK = "События — собирает бот"


def ensure_event_sheets(sh: Sheet, partners: list[str]) -> dict[str, int]:
    """Справочник, скрытые списки, «Гости» и «События»; ручной прототип «Событий» уходит в архив."""
    tabs = sh.tabs()
    sep = sh.sep()
    email = service_email()
    if LISTS_TAB not in tabs:
        gid = sh.add_tab(LISTS_TAB, rows=300, cols=6)
        sh.batch([{"updateSheetProperties": {"properties": {"sheetId": gid, "hidden": True}, "fields": "hidden"}}])
        sh.put(f"'{LISTS_TAB}'!A1", lists_values(sep, partners))
        tabs[LISTS_TAB] = gid
    else:   # партнёры меняются — список «Кто привёл» обновляем
        sh.clear(f"'{LISTS_TAB}'!F2:F300")
        sh.put(f"'{LISTS_TAB}'!F2", [[p] for p in partners] or [[""]], raw=True)
    if EVENTS_TAB not in tabs:
        gid = sh.add_tab(EVENTS_TAB, index=4, rows=520, cols=7)
        sh.put(f"'{EVENTS_TAB}'!A1", events_tab_values(sep))
        sh.batch(events_tab_requests(gid, email))
        tabs[EVENTS_TAB] = gid
    if GUESTS_TAB not in tabs:
        gid = sh.add_tab(GUESTS_TAB, index=5, rows=300, cols=6)
        sh.put(f"'{GUESTS_TAB}'!A1", [["Гости «плюс один»: у гостя ещё нет своей строки ни у одного партнёра. "
                                       "Заполните строку — бот добавит гостя клиентом на вкладку партнёра."],
                                      [], [], GUESTS_HEAD], raw=True)
        L = f"='{LISTS_TAB}'!"
        sh.batch([
            {"updateSheetProperties": {"properties": {"sheetId": gid, "gridProperties": {"frozenRowCount": 4}},
                                       "fields": "gridProperties.frozenRowCount"}},
            _fill(gid, _rng(gid, 3, 4, 0, 6), HEAD, bold=True, wrap=True),
            _fill(gid, _rng(gid, 4, 300, 0, 5), YELLOW),
            _list_rule(gid, 4, 300, 0, f"{L}E2:E500"),
            _list_rule(gid, 4, 300, 3, f"{L}F2:F300"),
            {"setDataValidation": {"range": _rng(gid, 4, 300, 4, 5), "rule": {"condition": {"type": "BOOLEAN"}}}},
            _width(gid, 0, 260), _width(gid, 1, 220), _width(gid, 2, 170), _width(gid, 3, 190),
            _width(gid, 4, 55), _width(gid, 5, 260),
        ])
        tabs[GUESTS_TAB] = gid
    head = (sh.get_values(f"'{REPORT_TAB}'!A1:A1") or [[""]]) if REPORT_TAB in tabs else [[""]]
    if REPORT_TAB in tabs and not str((head[0] or [""])[0]).startswith(REPORT_MARK):
        sh.batch([{"updateSheetProperties": {"properties": {"sheetId": tabs[REPORT_TAB],
                                                            "title": "_События (ручные, до бота)", "hidden": True},
                                             "fields": "title,hidden"}}])
        tabs.pop(REPORT_TAB)
    if REPORT_TAB not in tabs:
        gid = sh.add_tab(REPORT_TAB, index=4, rows=2000, cols=16)
        reqs = report_rules(gid)
        if email:
            reqs.append({"addProtectedRange": {"protectedRange": {
                "range": {"sheetId": gid}, "description": "bot:события", "warningOnly": False,
                "editors": {"users": [email]}}}})
        sh.batch(reqs)
        tabs[REPORT_TAB] = gid
    return tabs


def needs_event_cols(tab) -> bool:
    return tab.event_col is None and tab.desc_col is not None and tab.plan_col is not None


def event_cols_upgrade_requests(tab, gid: int, today: date, sep: str, protected: list[dict],
                                toc_gid=None, guide_gid=None) -> list[dict]:
    """Одним batchUpdate: три колонки после «Что происходит», затем «Мой месяц» собирается заново
    (вставка проходит через его строки 1–3) и обновляются ссылки на часы."""
    at = tab.desc_col + 1
    reqs = [{"insertDimension": {"range": {"sheetId": gid, "dimension": "COLUMNS",
                                           "startIndex": at, "endIndex": at + 3}, "inheritFromBefore": True}},
            _text_row(gid, HEADER_ROW - 1, at, EVENT_COLS)]
    reqs += event_columns_requests(gid, at)
    plan = tab.plan_col + 3
    if tab.month_col is not None:
        import copy
        shifted = copy.copy(tab)
        if tab.month_col < at <= max(tab.month_end, tab.month_col + 1):
            shifted.month_end = tab.month_end + 3
        reqs += month_repair_requests(shifted, gid, today, sep, protected, plan + len(PLAN_HEADERS))
    reqs += partner_links(gid, plan, toc_gid, guide_gid)
    return reqs


def upgrade_event_cols(sh: Sheet, tab, gid: int, today: date) -> bool:
    if not needs_event_cols(tab):
        return False
    r = sh.s.get(f"{API}/{sh.id}", params={"fields": "sheets(properties.sheetId,protectedRanges)"}, timeout=30)
    r.raise_for_status()
    protected = next((x.get("protectedRanges", []) for x in r.json().get("sheets", [])
                      if x["properties"]["sheetId"] == gid), [])
    tabs = sh.tabs()
    sh.batch(event_cols_upgrade_requests(tab, gid, today, sh.sep(), protected, tabs.get(TOC), tabs.get(GUIDE_TAB)))
    log.info("Вкладка «%s»: добавлены колонки мероприятий", tab.name)
    return True


def write_report(sh: Sheet, summary: list[list], people: list[list], stamp: str, gid: int) -> None:
    sh.clear(f"'{REPORT_TAB}'!A1:P2000")
    sh.put(f"'{REPORT_TAB}'!A1", report_values(summary, people, stamp), raw=True)
    sh.batch(report_format_requests(gid, len(summary), len(people)))


# ---------- источники для сверки ----------

def read_tracker(session, tracker_id: str) -> list[list]:
    """Лист «Мероприятие оплаты» из «Трекера комьюнити» (бот должен быть читателем таблицы)."""
    from urllib.parse import quote
    r = session.get(f"{API}/{tracker_id}", params={"fields": "sheets.properties.title"}, timeout=30)
    r.raise_for_status()
    title = next((s["properties"]["title"] for s in r.json()["sheets"]
                  if s["properties"]["title"].strip().lower().startswith("мероприятие")), None)
    if not title:
        return []
    r = session.get(f"{API}/{tracker_id}/values/{quote(repr_range(title))}", timeout=30)
    r.raise_for_status()
    return r.json().get("values", [])


def repr_range(title: str) -> str:
    return "'" + title.replace("'", "''") + "'!A1:J3000"


def read_sovetnik_export(session, folder_id: str) -> list[list] | None:
    """Самый свежий файл в папке выгрузок: Google Таблица, Excel или CSV. None — папки/файлов нет."""
    r = session.get(DRIVE, params={"q": f"'{folder_id}' in parents and trashed=false",
                                   "orderBy": "modifiedTime desc", "pageSize": 5,
                                   "fields": "files(id,name,mimeType,modifiedTime)",
                                   "supportsAllDrives": "true", "includeItemsFromAllDrives": "true"}, timeout=30)
    r.raise_for_status()
    files = r.json().get("files", [])
    if not files:
        return None
    f = files[0]
    if f["mimeType"] == "application/vnd.google-apps.spreadsheet":
        r = session.get(f"{DRIVE}/{f['id']}/export", params={"mimeType": "text/csv"}, timeout=60)
    else:
        r = session.get(f"{DRIVE}/{f['id']}", params={"alt": "media"}, timeout=60)
    r.raise_for_status()
    data = r.content
    if f["name"].lower().endswith((".xlsx", ".xlsm")) or data[:2] == b"PK":
        wb = load_workbook(io.BytesIO(data), read_only=True, data_only=True)
        try:
            return [list(row) for row in wb[wb.sheetnames[0]].iter_rows(values_only=True)]
        finally:
            wb.close()
    for enc in ("utf-8-sig", "cp1251"):
        try:
            text = data.decode(enc)
            break
        except UnicodeDecodeError:
            continue
    dialect = csv.Sniffer().sniff(text[:2000], delimiters=",;\t") if text.strip() else csv.excel
    return list(csv.reader(io.StringIO(text), dialect))


# ---------- гости «плюс один» ----------

def guest_requests(guests: list[list], tabs_by_name: dict, gids: dict) -> tuple[list[dict], list[tuple[int, str]]]:
    """Строки листа «Гости» без отметки бота → клиент на вкладке партнёра.
    Возвращает (данные для values:batchUpdate, [(номер строки «Гостей», отметка)])."""
    data, marks = [], []
    taken = {}
    for i, row in enumerate(guests[4:], start=5):
        get = lambda k: (row[k] if k < len(row) else "") or ""
        event, name, contact, who, came, done = (str(get(k)).strip() for k in range(6))
        if done or not (name and who):
            continue
        tab = tabs_by_name.get(who)
        if tab is None or tab.event_col is None:
            marks.append((i, "⚠ нет вкладки партнёра с колонками мероприятий"))
            continue
        if any(c.name.strip().lower() == name.lower() for c in tab.clients):
            marks.append((i, f"✓ уже есть у {who}"))
            continue
        used = taken.setdefault(who, {c.row for c in tab.clients})
        row_n = next((r for r in range(HEADER_ROW + 1, HEADER_ROW + 101) if r not in used), None)
        if row_n is None:
            marks.append((i, "⚠ у партнёра закончились строки"))
            continue
        used.add(row_n)
        from .workbook import col_letter
        put = lambda col, v: data.append({"range": f"'{who}'!{col_letter(col)}{row_n}", "values": [[v]]})
        put(1, name)
        if tab.contact_col is not None and contact:
            put(tab.contact_col, contact)
        put(tab.event_col, event)
        put(tab.event_col + 1, came.upper() in {"TRUE", "ИСТИНА"})
        marks.append((i, f"✓ добавлен на вкладку «{who}», строка {row_n}"))
    return data, marks


# ---------- целиком ----------

def run(session, store, books, cfg_tz, today: date, tracker_id: str = "", sovetnik_folder: str = "") -> list[str]:
    """books: [(Book, PartnerTab-ы)] из load_network. Возвращает предупреждения."""
    warnings = []
    main = next((b for b in books if not b.sandbox), None)
    if main is None:
        return ["нет основной таблицы"]
    sh = Sheet(session, main.sheet_id)
    partners = [p.name for p in main.directory.partners]
    tabs = ensure_event_sheets(sh, partners)

    events = parse_event_rows(sh.get_values(f"'{EVENTS_TAB}'!A5:F500"))
    ev_by_key = {e.key: e for e in events}

    for b in books:                                           # колонки мероприятий на вкладках
        bsh = sh if b is main else Sheet(session, b.sheet_id)
        for t in b.tabs:
            if t.name in b.gids:
                try:
                    upgrade_event_cols(bsh, t, b.gids[t.name], today)
                except Exception as exc:
                    log.exception("Колонки мероприятий: %s", t.name)
                    warnings.append(f"{t.name}: колонки мероприятий не добавились ({exc.__class__.__name__})")

    # гости «плюс один» → строки у партнёров
    guests = sh.get_values(f"'{GUESTS_TAB}'!A1:F300")
    data, marks = guest_requests(guests, {t.name: t for t in main.tabs}, main.gids)
    if data:
        r = sh.s.post(f"{API}/{sh.id}/values:batchUpdate", json={"valueInputOption": "RAW", "data": data}, timeout=60)
        r.raise_for_status()
    for i, mark in marks:
        sh.put(f"'{GUESTS_TAB}'!F{i}", [[mark]], raw=True)

    pairs = [(b.sheet_id, b.tabs) for b in books]
    journal = update_journal(store.invitations(), pairs, ev_by_key, today)
    store.save_invitations([{k: v for k, v in r.items() if not k.startswith("_")} for r in journal])

    tracker = None
    if tracker_id:
        try:
            tracker = parse_tracker(read_tracker(session, tracker_id),
                                    [p.name for b in books for p in b.directory.partners], today)
        except Exception as exc:
            warnings.append(f"Трекер комьюнити не прочитан ({exc.__class__.__name__}): дайте боту доступ «Читатель»")
    sovetnik = None
    if sovetnik_folder:
        try:
            rows = read_sovetnik_export(session, sovetnik_folder)
            sovetnik = parse_sovetnik_export(rows) if rows else None
        except Exception as exc:
            warnings.append(f"Выгрузка Советника не прочитана ({exc.__class__.__name__}): дайте боту доступ к папке")

    who = {t.name: t.name for b in books for t in b.tabs}
    summary, people = build_report(journal, pairs, events, today, tracker, sovetnik, who)
    from datetime import datetime
    write_report(sh, summary, people, datetime.now(cfg_tz).strftime("%d.%m.%Y %H:%M"), tabs[REPORT_TAB])
    return warnings
