"""Структура таблицы партнёрской сети (ТЗ от 3 ноября).

Одна Google Таблица, у каждого партнёра своя вкладка с его именем:
  слева  — CRM: клиенты, 5 касаний с датами и форматом, статус, что происходит,
           «Ближайшее касание» считается формулой;
  справа — планирование процессов по дням: план и факт часов по привлечению,
           прогреву и закрытию.
Плюс «Оглавление» (ссылки на вкладки), «Сводка» (её пишет бот) и «Партнёры»
(справочник с @никами).

Здесь же — разбор этих вкладок из xlsx-выгрузки и запросы Sheets API для их создания.
"""
from __future__ import annotations

import io
import re
from dataclasses import dataclass, field
from datetime import date, timedelta

from openpyxl import load_workbook

from .sheets import Lead, Touch, _clean, _norm, to_date

# ---------- макет вкладки партнёра ----------

HEADER_ROW = 4                 # строка заголовков (1-based)
FIRST_ROW = HEADER_ROW + 1
CLIENT_ROWS = 100              # сколько строк под клиентов
CRM_HEADERS = [
    "№", "Имя клиента", "Откуда пришёл",
    "Следующий шаг: дата", "Следующий шаг: что",
    "Текущий статус", "Что происходит",
    "Дата 1-го касания", "Формат", "Дата 2-го касания", "Формат", "Дата 3-го касания", "Формат",
    "Дата 4-го касания", "Формат", "Дата 5-го касания", "Формат",
]
COL_NAME, COL_SOURCE, COL_NEXT_DATE, COL_NEXT_WHAT, COL_STATUS, COL_DESC = 1, 2, 3, 4, 5, 6
TOUCH_COLS = [(n, 7 + 2 * (n - 1), 8 + 2 * (n - 1)) for n in range(1, 6)]   # история: (n, дата, формат)
DATE_COLS = [d for _, d, _ in TOUCH_COLS]

PLAN_COL0 = len(CRM_HEADERS) + 1   # колонка S — начало блока часов (R — узкий разделитель)
PLAN_HEADERS = [
    "Дата",
    "Привлечение — план, ч", "Прогрев — план, ч", "Закрытие — план, ч",
    "Привлечение — факт, ч", "Прогрев — факт, ч", "Закрытие — факт, ч",
    "Привлечение — людей", "Прогрев — людей", "Закрытие — людей",
    "Что делал(а) / комментарий",
]
P_PLAN, P_FACT, P_PEOPLE, P_COMMENT = 1, 4, 7, 10   # смещения от PLAN_COL0
PROCESSES = ["attract", "warmup", "close"]
PROCESS_NAMES = {"attract": "привлечение", "warmup": "прогрев", "close": "закрытие"}

STATUSES = [
    "Новый контакт", "Был на встрече", "Назначено касание", "Думает", "Диагностика",
    "Оплатил вход 5000", "Оплатил пакет 240 / стал ментором", "Пауза", "Отказ",
]
EXCLUDED_RE = r"отказ"                                   # не в работе
CLOSED_RE = r"стал ментором|оплатил пакет|240"           # сделка закрыта, касание не нужно

SERVICE_TABS = {"Оглавление", "Сводка", "Партнёры", "Инструкция"}


@dataclass
class DayPlan:
    day: date
    plan: dict[str, float]     # процесс -> часы (только заполненные)
    fact: dict[str, float]
    comment: str = ""
    people: dict[str, float] = field(default_factory=dict)   # сколько людей в работе (факт)

    @property
    def has_plan(self) -> bool:
        return bool(self.plan)

    @property
    def has_fact(self) -> bool:
        return bool(self.fact)


@dataclass
class PartnerTab:
    name: str
    clients: list[Lead] = field(default_factory=list)
    days: dict[date, DayPlan] = field(default_factory=dict)
    sandbox: bool = False
    touch_cols: list = field(default_factory=list)   # фактические колонки истории [(n, дата, формат)]
    plan_col: int | None = None                      # колонка «Дата» блока часов

    def history_full_soon(self) -> bool:
        """У кого-то из клиентов осталось меньше двух свободных пар касаний."""
        if not self.touch_cols:
            return False
        cap = len(self.touch_cols)
        return any(max((t.n for t in c.touches), default=0) >= cap - 1 for c in self.clients)


# ---------- разбор ----------

def _num(v) -> float | None:
    if v is None or v == "":
        return None
    try:
        return float(str(v).replace(",", "."))
    except ValueError:
        return None


def _is_partner_tab(ws) -> int | None:
    """Строка заголовков, если на листе есть CRM партнёра."""
    for r_idx, row in enumerate(ws.iter_rows(min_row=1, max_row=10, values_only=True), 1):
        cells = {_norm(c) for c in row}
        if "имя клиента" in cells and "текущий статус" in cells:
            return r_idx
    return None


def parse_partner_tabs(xlsx: bytes, sandbox: bool = False) -> list[PartnerTab]:
    wb = load_workbook(io.BytesIO(xlsx), data_only=True, read_only=True)
    out = []
    try:
        for title in wb.sheetnames:
            if title in SERVICE_TABS:
                continue
            ws = wb[title]
            header_row = _is_partner_tab(ws)
            if header_row is None:
                continue
            rows = list(ws.iter_rows(min_row=header_row, values_only=True))
            out.append(_parse_rows(title, rows, sandbox))
    finally:
        wb.close()
    return out


def _parse_rows(title: str, rows: list[tuple], sandbox: bool) -> PartnerTab:
    head = [_norm(c) for c in rows[0]]
    idx = {h: i for i, h in enumerate(head) if h}

    def find(prefix, start=0):
        return next((i for i, h in enumerate(head) if i >= start and h.startswith(prefix)), None)

    c_name, c_src = idx.get("имя клиента"), find("откуда")
    c_status, c_desc = find("текущий статус"), find("что происходит")
    c_next_date, c_next_what = find("следующий шаг: дата"), find("следующий шаг: что")
    touch_cols = []
    for i, h in enumerate(head):
        m = re.match(r"дата\s*(\d+)", h)
        if m:
            f = i + 1 if i + 1 < len(head) and head[i + 1].startswith("формат") else None
            touch_cols.append((int(m.group(1)), i, f))
    c_date = next((i for i, h in enumerate(head) if h == "дата"), None)
    plan_cols = {p: find(f"{PROCESS_NAMES[p]} — план") for p in PROCESSES}
    fact_cols = {p: find(f"{PROCESS_NAMES[p]} — факт") for p in PROCESSES}
    people_cols = {p: find(f"{PROCESS_NAMES[p]} — людей") for p in PROCESSES}
    c_comment = find("что делал")

    def cell(row, i):
        return row[i] if i is not None and i < len(row) else None

    tab = PartnerTab(name=title, sandbox=sandbox, touch_cols=touch_cols, plan_col=c_date)
    for offset, row in enumerate(rows[1:], 1):
        r_num = HEADER_ROW + offset
        name = _clean(cell(row, c_name))
        if name:
            touches = []
            for n, d_i, f_i in touch_cols:
                raw, fmt = cell(row, d_i), cell(row, f_i)
                if raw in (None, "") and fmt in (None, ""):
                    continue
                touches.append(Touch(n=n, when=to_date(raw), raw=_clean(raw), fmt=_clean(fmt)))
            tab.clients.append(Lead(row=r_num, number="", name=name,
                                    source=_clean(cell(row, c_src)),
                                    status=_clean(cell(row, c_status)),
                                    description=_clean(cell(row, c_desc)), touches=touches,
                                    next_date=to_date(cell(row, c_next_date)),
                                    next_what=_clean(cell(row, c_next_what))))
        d = to_date(cell(row, c_date))
        if d:
            plan = {p: v for p in PROCESSES if (v := _num(cell(row, plan_cols[p]))) is not None}
            fact = {p: v for p in PROCESSES if (v := _num(cell(row, fact_cols[p]))) is not None}
            people = {p: v for p in PROCESSES if (v := _num(cell(row, people_cols[p]))) is not None}
            tab.days[d] = DayPlan(d, plan, fact, _clean(cell(row, c_comment)), people)
    return tab


# ---------- создание вкладок (запросы Sheets API) ----------

def col_letter(idx: int) -> str:
    s, n = "", idx + 1
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


YELLOW = {"red": 1.0, "green": 0.976, "blue": 0.85}
HEAD_BG = {"red": 0.93, "green": 0.93, "blue": 0.93}
PLAN_BG = {"red": 0.89, "green": 0.94, "blue": 0.99}
RED_BG = {"red": 0.957, "green": 0.8, "blue": 0.8}
RED_TEXT = {"red": 0.6, "green": 0.0, "blue": 0.0}
TODAY_BG = {"red": 0.85, "green": 0.93, "blue": 0.83}
PEOPLE_BG = {"red": 0.93, "green": 0.89, "blue": 0.98}
LINK_BLUE = {"red": 0.07, "green": 0.33, "blue": 0.8}
TOMORROW_BG = {"red": 0.85, "green": 0.92, "blue": 0.98}
NEXT_BG = {"red": 0.99, "green": 0.85, "blue": 0.7}


def plan_dates(start: date, end: date) -> list[date]:
    return [start + timedelta(days=i) for i in range((end - start).days + 1)]


def sep_for_locale(locale: str) -> str:
    """Разделитель аргументов формул: «;» там, где десятичная запятая (ru_RU и др.)."""
    return "," if (locale or "").split("_")[0] in {"en", "ja", "ko", "zh", "he", "th"} else ";"


def fx(formula: str, sep: str) -> str:
    """Формулы пишем с «,», под локаль меняем на «;» (вне кавычек)."""
    if sep == ",":
        return formula
    out, in_str = [], False
    for ch in formula:
        if ch == '"':
            in_str = not in_str
        out.append(";" if ch == "," and not in_str else ch)
    return "".join(out)


def partner_tab_values(name: str, start: date, end: date, clients: list[dict] | None = None,
                       sep: str = ";", gid: int | None = None,
                       days: dict | None = None) -> list[list]:
    """Значения вкладки целиком (USER_ENTERED): заголовки, ссылки-переходы, даты, клиенты, часы."""
    width = PLAN_COL0 + len(PLAN_HEADERS)
    grid = [["" for _ in range(width)] for _ in range(HEADER_ROW + CLIENT_ROWS)]
    grid[0][0] = f"Клиенты — {name}"
    grid[1][0] = ("Следующий шаг — что и когда запланировано, он сам копируется в касания справа. "
                  "Красным — клиент без шага в ближайшие 7 дней.")
    grid[0][PLAN_COL0] = "Мои часы по дням"
    grid[1][PLAN_COL0] = "Вечером: факт за сегодня (зелёная строка) и план на завтра (голубая)."
    if gid is not None:
        plan_cell = f"{col_letter(PLAN_COL0)}{HEADER_ROW}"
        grid[0][COL_NEXT_DATE] = fx(f'=HYPERLINK("#gid={gid}&range={plan_cell}","→ к моим часам")', sep)
        grid[0][PLAN_COL0 + 3] = fx(f'=HYPERLINK("#gid={gid}&range=A{HEADER_ROW}","← к клиентам")', sep)
    for i, h in enumerate(CRM_HEADERS):
        grid[HEADER_ROW - 1][i] = h
    for i, h in enumerate(PLAN_HEADERS):
        grid[HEADER_ROW - 1][PLAN_COL0 + i] = h
    for k in range(CLIENT_ROWS):
        grid[FIRST_ROW - 1 + k][0] = k + 1
    for k, d in enumerate(plan_dates(start, end)):
        if k >= CLIENT_ROWS:
            break
        row = grid[FIRST_ROW - 1 + k]
        row[PLAN_COL0] = d.isoformat()
        dp = (days or {}).get(d)
        if dp:
            for i, p in enumerate(PROCESSES):
                if p in dp.plan:
                    row[PLAN_COL0 + P_PLAN + i] = dp.plan[p]
                if p in dp.fact:
                    row[PLAN_COL0 + P_FACT + i] = dp.fact[p]
                if p in dp.people:
                    row[PLAN_COL0 + P_PEOPLE + i] = dp.people[p]
            row[PLAN_COL0 + P_COMMENT] = dp.comment
    for k, c in enumerate(clients or []):
        row = grid[FIRST_ROW - 1 + k]
        row[COL_NAME] = c.get("name", "")
        row[COL_SOURCE] = c.get("source", "")
        row[COL_NEXT_DATE] = c.get("next_date", "")
        row[COL_NEXT_WHAT] = c.get("next_what", "")
        row[COL_STATUS] = c.get("status", "")
        row[COL_DESC] = c.get("description", "")
        for t in c.get("touches", []):
            slot = next((tc for tc in TOUCH_COLS if tc[0] == t["n"]), None)
            if slot:
                row[slot[1]] = t["date"]
                row[slot[2]] = t.get("fmt", "")
    return grid


def partner_tab_format_requests(sheet_id: int, sep: str = ";", service_email: str = "") -> list[dict]:
    last = HEADER_ROW + CLIENT_ROWS
    crm_w, plan_w = len(CRM_HEADERS), len(PLAN_HEADERS)

    def rng(r0, r1, c0, c1):
        return {"sheetId": sheet_id, "startRowIndex": r0, "endRowIndex": r1,
                "startColumnIndex": c0, "endColumnIndex": c1}

    def fill(r, color, bold=False, wrap=False):
        fmt = {"backgroundColor": color}
        fields = "userEnteredFormat.backgroundColor"
        if bold:
            fmt["textFormat"] = {"bold": True}
            fields += ",userEnteredFormat.textFormat.bold"
        if wrap:
            fmt["wrapStrategy"] = "WRAP"
            fmt["verticalAlignment"] = "MIDDLE"
            fields += ",userEnteredFormat.wrapStrategy,userEnteredFormat.verticalAlignment"
        return {"repeatCell": {"range": r, "cell": {"userEnteredFormat": fmt}, "fields": fields}}

    def date_fmt(r):
        return {"repeatCell": {"range": r, "cell": {"userEnteredFormat": {
            "numberFormat": {"type": "DATE", "pattern": "dd.mm.yyyy"}}},
            "fields": "userEnteredFormat.numberFormat"}}

    def width(c0, c1, px):
        return {"updateDimensionProperties": {"range": {"sheetId": sheet_id, "dimension": "COLUMNS",
                                                        "startIndex": c0, "endIndex": c1},
                                              "properties": {"pixelSize": px}, "fields": "pixelSize"}}

    def date_rule(c, msg):
        return {"setDataValidation": {"range": rng(f0, last, c, c + 1), "rule": {
            "condition": {"type": "DATE_IS_VALID"}, "inputMessage": msg, "strict": True}}}

    h0, f0 = HEADER_ROW - 1, FIRST_ROW - 1
    nd, st, pl = col_letter(COL_NEXT_DATE), col_letter(COL_STATUS), col_letter(PLAN_COL0)
    no_step = fx(f'=AND($B{FIRST_ROW}<>"",NOT(REGEXMATCH(LOWER(${st}{FIRST_ROW}),"{EXCLUDED_RE}|{CLOSED_RE}")),'
                 f'OR(${nd}{FIRST_ROW}="",${nd}{FIRST_ROW}<TODAY(),${nd}{FIRST_ROW}>TODAY()+7))', sep)
    reqs = [
        {"updateSheetProperties": {"properties": {"sheetId": sheet_id, "gridProperties": {
            "frozenRowCount": HEADER_ROW, "frozenColumnCount": 2}},
            "fields": "gridProperties.frozenRowCount,gridProperties.frozenColumnCount"}},
        {"repeatCell": {"range": rng(0, 1, 0, PLAN_COL0 + plan_w), "cell": {"userEnteredFormat": {
            "textFormat": {"bold": True, "fontSize": 13}}}, "fields": "userEnteredFormat.textFormat"}},
        *[{"repeatCell": {"range": rng(0, 1, c, c + 1), "cell": {"userEnteredFormat": {"textFormat": {
            "foregroundColor": LINK_BLUE, "underline": True, "bold": True, "fontSize": 12}}},
            "fields": "userEnteredFormat.textFormat"}} for c in (COL_NEXT_DATE, PLAN_COL0 + 3)],
        fill(rng(h0, HEADER_ROW, 0, crm_w), HEAD_BG, bold=True, wrap=True),
        fill(rng(h0, HEADER_ROW, COL_NEXT_DATE, COL_NEXT_WHAT + 1), NEXT_BG, bold=True, wrap=True),
        fill(rng(h0, HEADER_ROW, PLAN_COL0, PLAN_COL0 + plan_w), PLAN_BG, bold=True, wrap=True),
        fill(rng(f0, last, 1, crm_w), YELLOW),
        fill(rng(f0, last, PLAN_COL0 + 1, PLAN_COL0 + plan_w), YELLOW),
        date_fmt(rng(f0, last, COL_NEXT_DATE, COL_NEXT_DATE + 1)),
        date_fmt(rng(f0, last, PLAN_COL0, PLAN_COL0 + 1)),
        width(0, 1, 40), width(1, 2, 190), width(2, 3, 140), width(3, 4, 110), width(4, 5, 190),
        width(5, 6, 170), width(6, 7, 260), width(7, crm_w, 100), width(crm_w, crm_w + 1, 24),
        width(PLAN_COL0, PLAN_COL0 + 1, 95), width(PLAN_COL0 + P_PLAN, PLAN_COL0 + P_PEOPLE, 100),
        width(PLAN_COL0 + P_PEOPLE, PLAN_COL0 + P_COMMENT, 90),
        width(PLAN_COL0 + P_COMMENT, PLAN_COL0 + P_COMMENT + 1, 260),
        fill(rng(h0, HEADER_ROW, PLAN_COL0 + P_PEOPLE, PLAN_COL0 + P_COMMENT), PEOPLE_BG, bold=True, wrap=True),
        {"updateDimensionProperties": {"range": {"sheetId": sheet_id, "dimension": "ROWS",
                                                 "startIndex": h0, "endIndex": HEADER_ROW},
                                       "properties": {"pixelSize": 48}, "fields": "pixelSize"}},
        {"setDataValidation": {"range": rng(f0, last, COL_STATUS, COL_STATUS + 1), "rule": {
            "condition": {"type": "ONE_OF_LIST", "values": [{"userEnteredValue": s} for s in STATUSES]},
            "showCustomUi": True, "strict": False}}},
        {"setDataValidation": {"range": rng(f0, last, PLAN_COL0 + P_PLAN, PLAN_COL0 + P_PEOPLE), "rule": {
            "condition": {"type": "NUMBER_GREATER_THAN_EQ", "values": [{"userEnteredValue": "0"}]},
            "inputMessage": "Часы числом, например 1,5", "strict": True}}},
        {"setDataValidation": {"range": rng(f0, last, PLAN_COL0 + P_PEOPLE, PLAN_COL0 + P_COMMENT), "rule": {
            "condition": {"type": "NUMBER_GREATER_THAN_EQ", "values": [{"userEnteredValue": "0"}]},
            "inputMessage": "Сколько людей было в работе по этому процессу сегодня", "strict": True}}},
        date_rule(COL_NEXT_DATE, "Когда следующий шаг? Выберите дату в календаре (двойной клик)"),
        {"addConditionalFormatRule": {"index": 0, "rule": {
            "ranges": [rng(f0, last, COL_NAME, COL_NAME + 1), rng(f0, last, COL_NEXT_DATE, COL_NEXT_DATE + 1)],
            "booleanRule": {"condition": {"type": "CUSTOM_FORMULA", "values": [{"userEnteredValue": no_step}]},
                            "format": {"backgroundColor": RED_BG,
                                       "textFormat": {"foregroundColor": RED_TEXT, "bold": True}}}}}},
        {"addConditionalFormatRule": {"index": 1, "rule": {
            "ranges": [rng(f0, last, PLAN_COL0, PLAN_COL0 + plan_w)],
            "booleanRule": {"condition": {"type": "CUSTOM_FORMULA", "values": [
                {"userEnteredValue": f"=${pl}{FIRST_ROW}=TODAY()"}]}, "format": {"backgroundColor": TODAY_BG}}}}},
        {"addConditionalFormatRule": {"index": 2, "rule": {
            "ranges": [rng(f0, last, PLAN_COL0, PLAN_COL0 + plan_w)],
            "booleanRule": {"condition": {"type": "CUSTOM_FORMULA", "values": [
                {"userEnteredValue": f"=${pl}{FIRST_ROW}=TODAY()+1"}]}, "format": {"backgroundColor": TOMORROW_BG}}}}},
    ]
    for c in DATE_COLS:
        reqs.append(date_fmt(rng(f0, last, c, c + 1)))
        reqs.append(date_rule(c, "Дата прошедшего касания (календарь — двойной клик)"))
    if service_email:   # заголовки и даты в часах правит только бот (владелец таблицы — всегда)
        for r in (rng(0, HEADER_ROW, 0, PLAN_COL0 + plan_w), rng(f0, last, PLAN_COL0, PLAN_COL0 + 1),
                  rng(f0, last, 0, 1)):
            reqs.append({"addProtectedRange": {"protectedRange": {
                "range": r, "description": "bot:заголовки", "warningOnly": False,
                "editors": {"users": [service_email]}}}})
    return reqs
