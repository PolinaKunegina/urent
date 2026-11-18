"""Мероприятия: справочник, выбор мероприятия у клиента, журнал приглашений и лист «События».

Как устроено
- «Мероприятия» — справочник. Одна строка на одно мероприятие: дата, название (как в календаре
  Советника), тип, формат, ведущий, план новых. Заполняет Наталья, тип/формат/ведущий — из списков.
  Колонка «Пункт в списке» собирается формулой: «18.11 · Бизнес-встреча».
- «_списки» (скрытый) — то, из чего выбирают: мероприятия текущей недели и будущие (формула),
  типы, форматы, ведущие, все мероприятия и партнёры (для листа «Гости»).
- На вкладке партнёра после «Что происходит»: «Мероприятие» (список), «Был» и «В Советнике» (галочки).
  Сменили мероприятие — «Был» сбрасывает скрипт истории касаний, «В Советнике» остаётся.
- Журнал (база бота) помнит каждое приглашение: клиенту выбрали новое мероприятие — старое
  приглашение со своими отметками остаётся в «Событиях».
- «События» бот переписывает целиком: сверху сводка по мероприятиям, ниже — все приглашённые.
  Руками этот лист никто не правит.
- Сверка: приглашённые на вкладках ↔ строки «Событий» ↔ пришедшие новички по «Трекеру комьюнити»;
  галочка «В Советнике» ↔ выгрузка пользователей Советника (по телефону и нику из «Контакта»).
"""
from __future__ import annotations

import io
import re
from dataclasses import dataclass, field
from datetime import date, datetime, timedelta

from openpyxl import load_workbook

from .sheets import _clean, _norm, to_date
from .workbook import FIRST_ROW, HEADER_ROW, CLIENT_ROWS, STATUSES, fx

EVENTS_TAB, LISTS_TAB, REPORT_TAB, GUESTS_TAB = "Мероприятия", "_списки", "События", "Гости"
EVENT_TYPES = ["Бизнес-встреча", "Бизнес-завтрак", "Тренинг", "Разбор", "Кинотренинг", "Созвон в треке", "Другое"]
EVENT_FORMATS = ["Офлайн", "Онлайн", "Офлайн + онлайн"]
HOSTS = ["Кир Воликов", "Наталья Сергиенко", "Антон Быков"]
EVENTS_HEADERS = ["Дата", "Название (как в календаре Советника)", "Тип", "Формат", "Ведущий",
                  "План новых", "Пункт в списке (сам)"]
EV_FIRST = 5                     # первая строка мероприятий в справочнике
EV_LAST = 500
EVENT_COLS = ["Мероприятие", "Был", "В Советнике"]
PAID_RE = re.compile(r"купил|оплатил", re.I)
RANK = {s.lower(): i for i, s in enumerate(STATUSES)}

GREY = {"red": 0.95, "green": 0.95, "blue": 0.95}
HEAD = {"red": 0.89, "green": 0.94, "blue": 0.99}
YELLOW = {"red": 1.0, "green": 0.976, "blue": 0.85}
WARN = {"red": 0.99, "green": 0.85, "blue": 0.7}
GOOD = {"red": 0.85, "green": 0.94, "blue": 0.83}


@dataclass
class Event:
    day: date
    title: str
    kind: str = ""
    fmt: str = ""
    host: str = ""
    plan: float | None = None

    @property
    def key(self) -> str:
        return event_key(self.day, self.title)


def event_key(day: date, title: str) -> str:
    return f"{day:%d.%m} · {title.strip()}"


# ---------- справочник и списки ----------

def parse_events(xlsx: bytes) -> list[Event]:
    wb = load_workbook(io.BytesIO(xlsx), data_only=True, read_only=True)
    try:
        if EVENTS_TAB not in wb.sheetnames:
            return []
        rows = list(wb[EVENTS_TAB].iter_rows(min_row=EV_FIRST, values_only=True))
    finally:
        wb.close()
    return parse_event_rows(rows)


def parse_event_rows(rows: list) -> list[Event]:
    """Строки справочника начиная с первой строки мероприятий (A дата … F план)."""
    out = []
    for r in rows:
        get = lambda i: r[i] if i < len(r) else None
        d, title = to_date(get(0)), _clean(get(1))
        if not d or not title:
            continue
        plan = get(5)
        try:
            plan = float(str(plan).replace(",", ".")) if plan not in (None, "") else None
        except ValueError:
            plan = None
        out.append(Event(d, title, _clean(get(2)), _clean(get(3)), _clean(get(4)), plan))
    return out


def lists_values(sep: str, partners: list[str]) -> list[list]:
    """Скрытый лист «_списки». A — мероприятия для выбора партнёром (текущая неделя и будущие,
    по дате), B — типы, C — форматы, D — ведущие, E — все мероприятия (новые сверху), F — партнёры."""
    ev = f"'{EVENTS_TAB}'!A{EV_FIRST}:G{EV_LAST}"
    dates, names = f"'{EVENTS_TAB}'!A{EV_FIRST}:A{EV_LAST}", f"'{EVENTS_TAB}'!B{EV_FIRST}:B{EV_LAST}"
    week = fx(f'=IFERROR(INDEX(SORT(FILTER({ev},{dates}>=TODAY()-WEEKDAY(TODAY(),3),{names}<>""),1,TRUE),0,7),"")',
              sep)
    every = fx(f'=IFERROR(INDEX(SORT(FILTER({ev},{names}<>""),1,FALSE),0,7),"")', sep)
    head = ["Мероприятия для выбора (сами)", "Типы", "Форматы", "Ведущие (можно дописать)",
            "Все мероприятия (сами)", "Партнёры (сами)"]
    n = max(len(EVENT_TYPES), len(EVENT_FORMATS), len(HOSTS), len(partners), 1)
    rows = [head, [week, EVENT_TYPES[0], EVENT_FORMATS[0], HOSTS[0], every, partners[0] if partners else ""]]
    for i in range(1, n):
        rows.append(["", *(lst[i] if i < len(lst) else "" for lst in (EVENT_TYPES, EVENT_FORMATS, HOSTS)),
                     "", partners[i] if i < len(partners) else ""])
    return rows


def events_tab_values(sep: str, events: list[Event] | None = None) -> list[list]:
    rows = [["Мероприятия — справочник", "Заполняет Наталья: одна строка — одно мероприятие. "
             "Тип, формат и ведущий — из списка, план — сколько новых людей ждём."],
            ["Название пишите как в календаре Советника. Повторяющиеся события различаются датой."],
            [],
            EVENTS_HEADERS]
    for k, e in enumerate(events or []):
        r = EV_FIRST + k
        rows.append([e.day.isoformat(), e.title, e.kind, e.fmt, e.host, "" if e.plan is None else e.plan,
                     key_formula(r, sep)])
    for r in range(EV_FIRST + len(events or []), EV_FIRST + 60):
        rows.append(["", "", "", "", "", "", key_formula(r, sep)])
    return rows


def key_formula(r: int, sep: str) -> str:
    return fx(f'=IF(OR(A{r}="",B{r}=""),"",TEXT(A{r},"dd.mm")&" · "&TRIM(B{r}))', sep)


def _rng(gid, r0, r1, c0, c1):
    return {"sheetId": gid, "startRowIndex": r0, "endRowIndex": r1, "startColumnIndex": c0, "endColumnIndex": c1}


def _list_rule(gid, r0, r1, c, src: str, strict=True, msg=""):
    rule = {"condition": {"type": "ONE_OF_RANGE", "values": [{"userEnteredValue": src}]},
            "showCustomUi": True, "strict": strict}
    if msg:
        rule["inputMessage"] = msg
    return {"setDataValidation": {"range": _rng(gid, r0, r1, c, c + 1), "rule": rule}}


def _width(gid, c, px):
    return {"updateDimensionProperties": {"range": {"sheetId": gid, "dimension": "COLUMNS",
                                                    "startIndex": c, "endIndex": c + 1},
                                          "properties": {"pixelSize": px}, "fields": "pixelSize"}}


def _fill(gid, rng, color, bold=False, wrap=False):
    fmt = {"backgroundColor": color}
    fields = ["userEnteredFormat.backgroundColor"]
    if bold:
        fmt["textFormat"] = {"bold": True}
        fields.append("userEnteredFormat.textFormat.bold")
    if wrap:
        fmt["wrapStrategy"] = "WRAP"
        fields.append("userEnteredFormat.wrapStrategy")
    return {"repeatCell": {"range": rng, "cell": {"userEnteredFormat": fmt}, "fields": ",".join(fields)}}


def events_tab_requests(gid: int, service_email: str = "", owner_email: str = "") -> list[dict]:
    f0, last = EV_FIRST - 1, EV_LAST
    L = f"='{LISTS_TAB}'!"
    reqs = [
        {"updateSheetProperties": {"properties": {"sheetId": gid, "gridProperties": {"frozenRowCount": 4}},
                                   "fields": "gridProperties.frozenRowCount"}},
        {"repeatCell": {"range": _rng(gid, 0, 1, 0, 1), "cell": {"userEnteredFormat": {
            "textFormat": {"bold": True, "fontSize": 13}}}, "fields": "userEnteredFormat.textFormat"}},
        _fill(gid, _rng(gid, 3, 4, 0, 7), HEAD, bold=True, wrap=True),
        _fill(gid, _rng(gid, f0, last, 0, 6), YELLOW),
        _fill(gid, _rng(gid, f0, last, 6, 7), GREY),
        {"repeatCell": {"range": _rng(gid, f0, last, 0, 1), "cell": {"userEnteredFormat": {
            "numberFormat": {"type": "DATE", "pattern": "dd.mm.yyyy"}}}, "fields": "userEnteredFormat.numberFormat"}},
        {"setDataValidation": {"range": _rng(gid, f0, last, 0, 1), "rule": {
            "condition": {"type": "DATE_IS_VALID"}, "inputMessage": "Дата мероприятия (двойной клик — календарь)",
            "strict": True}}},
        _list_rule(gid, f0, last, 2, f"{L}B2:B30"),
        _list_rule(gid, f0, last, 3, f"{L}C2:C30"),
        _list_rule(gid, f0, last, 4, f"{L}D2:D30", strict=False, msg="Ведущий из списка; новый — допишите на листе «_списки», колонка D"),
        {"setDataValidation": {"range": _rng(gid, f0, last, 5, 6), "rule": {
            "condition": {"type": "NUMBER_GREATER_THAN_EQ", "values": [{"userEnteredValue": "0"}]},
            "inputMessage": "Сколько новых людей ждём (план компании)", "strict": True}}},
        _width(gid, 0, 100), _width(gid, 1, 340), _width(gid, 2, 140), _width(gid, 3, 130),
        _width(gid, 4, 160), _width(gid, 5, 90), _width(gid, 6, 300),
        {"updateDimensionProperties": {"range": {"sheetId": gid, "dimension": "ROWS", "startIndex": 3,
                                                 "endIndex": 4}, "properties": {"pixelSize": 40},
                                       "fields": "pixelSize"}},
    ]
    editors = [e for e in (service_email, owner_email) if e]
    if editors:
        reqs.append({"addProtectedRange": {"protectedRange": {
            "range": _rng(gid, 3, last, 6, 7), "description": "bot:мероприятия", "warningOnly": False,
            "editors": {"users": editors}}}})
    return reqs


# ---------- колонки мероприятий на вкладке партнёра ----------

def event_columns_requests(gid: int, at: int) -> list[dict]:
    """Формат трёх колонок с индекса at: «Мероприятие» (список), «Был», «В Советнике» (галочки)."""
    f0, last = FIRST_ROW - 1, HEADER_ROW + CLIENT_ROWS
    reqs = [
        _list_rule(gid, f0, last, at, f"='{LISTS_TAB}'!A2:A300", strict=False,
                   msg="Мероприятие из списка: текущая неделя и будущие"),
        {"repeatCell": {"range": _rng(gid, f0, last, at, at + 1), "cell": {"userEnteredFormat": {
            "numberFormat": {"type": "TEXT"}}}, "fields": "userEnteredFormat.numberFormat"}},
    ]
    for c in (at + 1, at + 2):
        reqs.append({"setDataValidation": {"range": _rng(gid, f0, last, c, c + 1),
                                           "rule": {"condition": {"type": "BOOLEAN"}}}})
        reqs.append({"repeatCell": {"range": _rng(gid, f0, last, c, c + 1), "cell": {"userEnteredFormat": {
            "horizontalAlignment": "CENTER", "numberFormat": {"type": "TEXT"}}},
            "fields": "userEnteredFormat.horizontalAlignment,userEnteredFormat.numberFormat"}})
    reqs += [_fill(gid, _rng(gid, HEADER_ROW - 1, HEADER_ROW, at, at + 3), GOOD, bold=True, wrap=True),
             _fill(gid, _rng(gid, f0, last, at, at + 3), YELLOW),
             _width(gid, at, 190), _width(gid, at + 1, 55), _width(gid, at + 2, 85)]
    return reqs


# ---------- журнал и отчёт ----------

def _rank(status: str) -> int | None:
    st = (status or "").strip().lower()
    if not st or st.startswith("отказ"):
        return None
    return RANK.get(st)


def update_journal(journal: list[dict], books: list[tuple[str, list]], events: dict[str, Event],
                   today: date) -> list[dict]:
    """books: [(id таблицы, вкладки партнёров)]. Возвращает новый журнал (список словарей)."""
    by_key = {(r["book"], r["tab"], r["client"], r["event"]): dict(r) for r in journal}
    seen_clients = set()
    for book, tabs in books:
        for t in tabs:
            for c in t.clients:
                seen_clients.add((book, t.name, c.name))
                for r in by_key.values():          # прошлые приглашения этого клиента — больше не текущие
                    if (r["book"], r["tab"], r["client"]) == (book, t.name, c.name) and r["event"] != c.event:
                        r["current"] = 0
                if not c.event:
                    continue
                k = (book, t.name, c.name, c.event)
                r = by_key.setdefault(k, {"book": book, "tab": t.name, "client": c.name, "event": c.event,
                                          "first_seen": today.isoformat(), "status_at_event": None})
                r.update(contact=c.contact, attended=c.attended, registered=c.registered, current=1)
    for r in by_key.values():                      # статус в день мероприятия — один раз, когда день наступил
        ev = events.get(r["event"])
        if ev and ev.day <= today and r.get("status_at_event") is None:
            r["status_at_event"] = _status_of(books, r) or ""
        r["_on_tab"] = (r["book"], r["tab"], r["client"]) in seen_clients
    return list(by_key.values())


def _status_of(books, r) -> str:
    for book, tabs in books:
        if book != r["book"]:
            continue
        for t in tabs:
            if t.name == r["tab"]:
                c = next((c for c in t.clients if c.name == r["client"]), None)
                return c.status if c else ""
    return ""


def _client(books, r):
    for book, tabs in books:
        if book == r["book"]:
            for t in tabs:
                if t.name == r["tab"]:
                    return next((c for c in t.clients if c.name == r["client"]), None)
    return None


@dataclass
class Tracker:
    """Пришедшие новички по «Трекеру комьюнити»: день → [(имя, чей гость или '')]."""
    by_day: dict = field(default_factory=dict)


def norm_contact(text: str) -> set[str]:
    """Телефоны (последние 10 цифр) и @ники из ячейки «Контакт» или выгрузки."""
    out = set()
    for m in re.findall(r"\+?\d[\d\s\-()]{8,}\d", text or ""):
        d = re.sub(r"\D", "", m)
        if len(d) >= 10:
            out.add(d[-10:])
    for m in re.findall(r"(?:@|t\.me/)([A-Za-z0-9_]{4,})", text or ""):
        out.add("@" + m.lower())
    return out


def build_report(journal: list[dict], books, events: list[Event], today: date,
                 tracker: Tracker | None = None, sovetnik: set[str] | None = None,
                 partners_by_tab: dict[str, str] | None = None) -> tuple[list[list], list[list]]:
    """(сводка по мероприятиям, приглашённые) — строки для листа «События»."""
    ev_by_key = {e.key: e for e in events}
    people = []
    for r in journal:
        c = _client(books, r)
        ev = ev_by_key.get(r["event"])
        status, stage = (c.status, c.stage) if c else ("", "")
        reg_check = ""
        if sovetnik is not None:
            ids = norm_contact(r.get("contact") or (c.contact if c else ""))
            found = bool(ids & sovetnik)
            if not ids:
                reg_check = "нет контакта — не сверить" if r.get("registered") else ""
            elif found and not r.get("registered"):
                reg_check = "⚠ в выгрузке есть, галочки нет"
            elif r.get("registered") and not found:
                reg_check = "⚠ галочка есть, в выгрузке нет"
        if not r.get("_on_tab", True):
            reg_check = ("⚠ клиента нет на вкладке. " + reg_check).strip()
        if ev and ev.day > today:
            outcome = "ещё впереди"
        elif r.get("attended"):
            outcome = "был и в Советнике" if r.get("registered") else "был, не в Советнике"
        else:
            outcome = "не пришёл"
        people.append({"r": r, "ev": ev, "status": status, "stage": stage, "outcome": outcome,
                       "check": reg_check, "who": (partners_by_tab or {}).get(r["tab"], r["tab"])})

    summary = []
    keys = sorted({p["r"]["event"] for p in people} | set(ev_by_key),
                  key=lambda k: (ev_by_key[k].day if k in ev_by_key else date.min), reverse=True)
    for k in keys:
        ev = ev_by_key.get(k)
        ps = [p for p in people if p["r"]["event"] == k]
        came = [p for p in ps if p["r"].get("attended")]
        reg = [p for p in came if p["r"].get("registered")]
        moved = [p for p in ps if _rank(p["status"]) is not None and p["r"].get("status_at_event") is not None
                 and _rank(p["r"]["status_at_event"]) is not None
                 and _rank(p["status"]) > _rank(p["r"]["status_at_event"])]
        paid = [p for p in ps if PAID_RE.search(p["status"] or "")]
        on_tabs = sum(1 for p in ps if p["r"].get("_on_tab", True))
        tr = None
        if tracker is not None and ev:
            tr = len(tracker.by_day.get(ev.day, []))
        checks = []
        if on_tabs != len(ps):
            checks.append(f"⚠ на вкладках {on_tabs}, в журнале {len(ps)}")
        if tr is not None and ev and ev.day <= today and tr != len(came):
            checks.append(f"⚠ пришло по галочкам {len(came)}, по трекеру {tr}")
        if not ev:
            checks.append("⚠ мероприятия нет в справочнике")
        by_partner = {}
        for p in ps:
            a = by_partner.setdefault(p["who"], [0, 0])
            a[0] += 1
            a[1] += 1 if p["r"].get("attended") else 0
        plan = ev.plan if ev and ev.plan is not None else ""
        pct = f"{round(100 * len(came) / plan)}%" if plan and ev.day <= today else ""
        summary.append([
            ev.day.isoformat() if ev else "", ev.title if ev else k, ev.kind if ev else "", ev.fmt if ev else "",
            plan, len(ps), len(came), "" if tr is None else tr, len(reg), len(came) - len(reg),
            len(moved), len(paid), pct, "; ".join(checks) or "✓",
            ", ".join(f"{w} {a[0]}/{a[1]}" for w, a in sorted(by_partner.items(), key=lambda x: -x[1][0])),
        ])

    rows = []
    for p in sorted(people, key=lambda p: ((p["ev"].day if p["ev"] else date.min), p["who"], p["r"]["client"]),
                    reverse=True):
        r = p["r"]
        rows.append([p["ev"].day.isoformat() if p["ev"] else "", p["ev"].title if p["ev"] else r["event"],
                     r["client"], r.get("contact") or "", p["who"], p["stage"], p["status"],
                     "✓" if r.get("attended") else "", "✓" if r.get("registered") else "", p["outcome"],
                     r.get("status_at_event") or "", p["check"]])
    return summary, rows


SUMMARY_HEAD = ["Дата", "Мероприятие", "Тип", "Формат", "План новых", "Приглашено", "Пришло (галочки)",
                "Пришло новых (трекер)", "Был и в Советнике", "Был, не в Советнике", "Дальше по воронке",
                "Дошли до оплаты", "% плана", "Сверка", "Кто сколько привёл (приглашено/пришло)"]
PEOPLE_HEAD = ["Дата", "Мероприятие", "Клиент", "Контакт", "Пригласил", "Этап", "Текущий статус", "Был",
               "В Советнике", "Итог", "Статус в день мероприятия", "Сверка"]


def report_values(summary: list[list], people: list[list], updated: str) -> list[list]:
    out = [["События — собирает бот из вкладок партнёров, руками не правьте", f"обновлено {updated}"],
           [],
           ["Сводка по мероприятиям"], SUMMARY_HEAD, *summary,
           [], ["Приглашённые: одна строка — одно приглашение"], PEOPLE_HEAD, *people]
    return out


def report_format_requests(gid: int, n_summary: int, n_people: int) -> list[dict]:
    """Заливка шапок и подсветка расхождений (ячейки, начинающиеся с «⚠»)."""
    s_head = 3
    p_head = s_head + 1 + n_summary + 2
    w = max(len(SUMMARY_HEAD), len(PEOPLE_HEAD))
    reqs = [
        {"repeatCell": {"range": _rng(gid, 0, p_head + 1 + n_people + 50, 0, w), "cell": {"userEnteredFormat": {}},
                        "fields": "userEnteredFormat"}},
        {"repeatCell": {"range": _rng(gid, 0, 1, 0, 1), "cell": {"userEnteredFormat": {
            "textFormat": {"bold": True, "fontSize": 13}}}, "fields": "userEnteredFormat.textFormat"}},
        {"repeatCell": {"range": _rng(gid, s_head - 1, s_head, 0, 1), "cell": {"userEnteredFormat": {
            "textFormat": {"bold": True, "fontSize": 12}}}, "fields": "userEnteredFormat.textFormat"}},
        {"repeatCell": {"range": _rng(gid, p_head - 1, p_head, 0, 1), "cell": {"userEnteredFormat": {
            "textFormat": {"bold": True, "fontSize": 12}}}, "fields": "userEnteredFormat.textFormat"}},
        _fill(gid, _rng(gid, s_head, s_head + 1, 0, len(SUMMARY_HEAD)), HEAD, bold=True, wrap=True),
        _fill(gid, _rng(gid, p_head, p_head + 1, 0, len(PEOPLE_HEAD)), HEAD, bold=True, wrap=True),
        {"repeatCell": {"range": _rng(gid, s_head + 1, s_head + 1 + n_summary, 0, 1), "cell": {"userEnteredFormat": {
            "numberFormat": {"type": "DATE", "pattern": "dd.mm.yyyy"}}}, "fields": "userEnteredFormat.numberFormat"}},
        {"repeatCell": {"range": _rng(gid, p_head + 1, p_head + 1 + n_people, 0, 1), "cell": {"userEnteredFormat": {
            "numberFormat": {"type": "DATE", "pattern": "dd.mm.yyyy"}}}, "fields": "userEnteredFormat.numberFormat"}},
    ]
    for c, px in enumerate([95, 260, 120, 110, 80, 90, 90, 100, 95, 100, 95, 90, 75, 260, 330]):
        reqs.append(_width(gid, c, px))
    return reqs


def report_rules(gid: int) -> list[dict]:
    """Условное форматирование «Событий» (ставится один раз, при создании листа)."""
    whole = _rng(gid, 3, 2000, 0, 15)
    return [
        {"addConditionalFormatRule": {"index": 0, "rule": {"ranges": [whole], "booleanRule": {
            "condition": {"type": "TEXT_STARTS_WITH", "values": [{"userEnteredValue": "⚠"}]},
            "format": {"backgroundColor": WARN, "textFormat": {"bold": True}}}}}},
        {"addConditionalFormatRule": {"index": 1, "rule": {"ranges": [whole], "booleanRule": {
            "condition": {"type": "TEXT_EQ", "values": [{"userEnteredValue": "был, не в Советнике"}]},
            "format": {"backgroundColor": {"red": 1.0, "green": 0.95, "blue": 0.75}}}}}},
        {"addConditionalFormatRule": {"index": 2, "rule": {"ranges": [whole], "booleanRule": {
            "condition": {"type": "TEXT_EQ", "values": [{"userEnteredValue": "был и в Советнике"}]},
            "format": {"backgroundColor": GOOD}}}}},
    ]


# ---------- трекер и выгрузка Советника ----------

def parse_tracker(values: list[list], partner_names: list[str], today: date) -> Tracker:
    """Лист «Мероприятие оплаты »: A дата «12.11», C участник, J формат (офлайн/онлайн/не был).
    Считаем новичков: не партнёры и не команда. Гость «Имя (гость Партнёра)» — тоже новичок."""
    from .partners import _key
    team = {_key(n) for n in partner_names}
    tr = Tracker()
    for row in values[1:]:
        get = lambda i: (row[i] if i < len(row) else "") or ""
        raw, who, fmt = str(get(0)).strip(), str(get(2)).strip(), str(get(9)).strip().lower()
        m = re.fullmatch(r"(\d{1,2})\.(\d{1,2})(?:\.(\d{2,4}))?", raw)
        if not m or not who or fmt.startswith("не был"):
            continue
        d, mo = int(m[1]), int(m[2])
        year = int(m[3]) + (2000 if m[3] and len(m[3]) == 2 else 0) if m[3] else today.year
        try:
            day = date(year, mo, d)
        except ValueError:
            continue
        if not m[3] and day > today + timedelta(days=60):     # «12.11» в январе — прошлый год
            day = day.replace(year=year - 1)
        base = re.sub(r"\s*\(.*\)\s*$", "", who)
        host = re.search(r"гост[ья]\s+([^)]+)", who, re.I)
        inner = re.search(r"\(([^)]*)\)", who)
        if not host and (_key(base) in team or (inner and _key(inner[1]) in team)):
            continue          # партнёр или команда, в том числе «Лина Герц (Алина Гайнаншина)»
        tr.by_day.setdefault(day, []).append((base, host[1].strip() if host else ""))
    return tr


def parse_sovetnik_export(rows: list[list]) -> set[str]:
    """Выгрузка пользователей Советника (CSV/Excel): ищем колонки с телефоном и ником по заголовку,
    а если заголовков не узнали — собираем телефоны и @ники из всех ячеек."""
    if not rows:
        return set()
    head = [_norm(h) for h in rows[0]]
    cols = [i for i, h in enumerate(head) if re.search(r"тел|phone|telegram|телеграм|ник|username|login", h)]
    out = set()
    for row in rows[1:]:
        cells = [row[i] for i in cols if i < len(row)] if cols else row
        for v in cells:
            v = "" if v is None else str(v)
            out |= norm_contact(v)
            if re.fullmatch(r"[A-Za-z][A-Za-z0-9_]{4,}", v.strip()):     # ник без @
                out.add("@" + v.strip().lower())
    return out


# ---------- вечернее напоминание партнёру ----------

def evening_lines(tab, events: dict[str, Event], today: date) -> list[str]:
    todays = {}
    for c in tab.clients:
        ev = events.get(c.event)
        if ev and ev.day == today:
            todays.setdefault(ev.key, []).append(c.name + (" ✓" if c.attended else ""))
    out = []
    for k, names in todays.items():
        out.append(f"📍 Сегодня «{events[k].title}»: отметь галочкой «Был», кто пришёл — "
                   + ", ".join(names))
    return out


def updated_stamp(tz) -> str:
    return datetime.now(tz).strftime("%d.%m.%Y %H:%M")
