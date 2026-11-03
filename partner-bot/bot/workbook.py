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
    "Дата 1-го касания", "Формат", "Дата 2-го касания", "Формат", "Дата 3-го касания", "Формат",
    "Дата 4-го касания", "Формат", "Дата 5-го касания", "Формат",
    "Текущий статус", "Что происходит", "Ближайшее касание (авто)",
]
COL_NAME, COL_SOURCE, COL_STATUS, COL_DESC, COL_NEXT = 1, 2, 13, 14, 15
TOUCH_COLS = [(n, 3 + 2 * (n - 1), 4 + 2 * (n - 1)) for n in range(1, 6)]   # (n, дата, формат)
DATE_COLS = [d for _, d, _ in TOUCH_COLS]

PLAN_COL0 = 17                 # колонка R — начало блока планирования
PLAN_HEADERS = [
    "Дата",
    "Привлечение — план, ч", "Прогрев — план, ч", "Закрытие — план, ч",
    "Привлечение — факт, ч", "Прогрев — факт, ч", "Закрытие — факт, ч",
    "Что делал(а) / комментарий",
]
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
    touch_cols = []
    for i, h in enumerate(head):
        m = re.match(r"дата\s*(\d+)", h)
        if m:
            f = i + 1 if i + 1 < len(head) and head[i + 1].startswith("формат") else None
            touch_cols.append((int(m.group(1)), i, f))
    c_date = next((i for i, h in enumerate(head) if h == "дата"), None)
    plan_cols = {p: find(f"{PROCESS_NAMES[p]} — план") for p in PROCESSES}
    fact_cols = {p: find(f"{PROCESS_NAMES[p]} — факт") for p in PROCESSES}
    c_comment = find("что делал")

    def cell(row, i):
        return row[i] if i is not None and i < len(row) else None

    tab = PartnerTab(name=title, sandbox=sandbox)
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
                                    description=_clean(cell(row, c_desc)), touches=touches))
        d = to_date(cell(row, c_date))
        if d:
            plan = {p: v for p in PROCESSES if (v := _num(cell(row, plan_cols[p]))) is not None}
            fact = {p: v for p in PROCESSES if (v := _num(cell(row, fact_cols[p]))) is not None}
            tab.days[d] = DayPlan(d, plan, fact, _clean(cell(row, c_comment)))
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


def next_touch_formula(r: int, sep: str = ";") -> str:
    d0, d1 = col_letter(DATE_COLS[0]), col_letter(DATE_COLS[-1] + 1)
    mins = f'MINIFS({d0}{r}:{d1}{r},{d0}{r}:{d1}{r},">="&TODAY())'
    return fx(f'=IF(B{r}="","",IF({mins}=0,"нет плана",{mins}))', sep)


def partner_tab_values(name: str, start: date, end: date, clients: list[dict] | None = None,
                       sep: str = ";") -> list[list]:
    """Значения вкладки целиком (USER_ENTERED): заголовки, формулы, даты планирования, клиенты."""
    width = PLAN_COL0 + len(PLAN_HEADERS)
    grid = [["" for _ in range(width)] for _ in range(HEADER_ROW + CLIENT_ROWS)]
    grid[0][0] = f"Клиенты — {name}"
    grid[1][0] = ("Жёлтое заполняете вы: клиенты, даты касаний (выбирайте в календаре), статус. "
                  "Красным — клиент без касания в ближайшие 7 дней.")
    grid[0][PLAN_COL0] = "Планирование процессов, часы"
    grid[1][PLAN_COL0] = "Вечером: факт за сегодня и план на завтра. Строка сегодняшнего дня — зелёная."
    for i, h in enumerate(CRM_HEADERS):
        grid[HEADER_ROW - 1][i] = h
    for i, h in enumerate(PLAN_HEADERS):
        grid[HEADER_ROW - 1][PLAN_COL0 + i] = h
    for k in range(CLIENT_ROWS):
        r = FIRST_ROW + k
        row = grid[r - 1]
        row[0] = k + 1
        row[COL_NEXT] = next_touch_formula(r, sep)
    for k, d in enumerate(plan_dates(start, end)):
        if k < CLIENT_ROWS:
            grid[FIRST_ROW - 1 + k][PLAN_COL0] = d.isoformat()
    for k, c in enumerate(clients or []):
        row = grid[FIRST_ROW - 1 + k]
        row[COL_NAME] = c.get("name", "")
        row[COL_SOURCE] = c.get("source", "")
        row[COL_STATUS] = c.get("status", "")
        row[COL_DESC] = c.get("description", "")
        for t in c.get("touches", []):
            slot = next((tc for tc in TOUCH_COLS if tc[0] == t["n"]), None)
            if slot:
                row[slot[1]] = t["date"]
                row[slot[2]] = t.get("fmt", "")
    return grid


def partner_tab_format_requests(sheet_id: int, sep: str = ";") -> list[dict]:
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

    h0, f0 = HEADER_ROW - 1, FIRST_ROW - 1
    no_touch = fx(f'=AND($B{FIRST_ROW}<>"",NOT(REGEXMATCH(LOWER($N{FIRST_ROW}),"{EXCLUDED_RE}|{CLOSED_RE}")),'
                  f'OR($P{FIRST_ROW}="нет плана",$P{FIRST_ROW}>TODAY()+7))', sep)
    plan_letter = col_letter(PLAN_COL0)
    reqs = [
        {"updateSheetProperties": {"properties": {"sheetId": sheet_id, "gridProperties": {
            "frozenRowCount": HEADER_ROW, "frozenColumnCount": 2}},
            "fields": "gridProperties.frozenRowCount,gridProperties.frozenColumnCount"}},
        {"repeatCell": {"range": rng(0, 1, 0, PLAN_COL0 + plan_w), "cell": {"userEnteredFormat": {
            "textFormat": {"bold": True, "fontSize": 13}}}, "fields": "userEnteredFormat.textFormat"}},
        fill(rng(h0, HEADER_ROW, 0, crm_w), HEAD_BG, bold=True, wrap=True),
        fill(rng(h0, HEADER_ROW, PLAN_COL0, PLAN_COL0 + plan_w), PLAN_BG, bold=True, wrap=True),
        fill(rng(f0, last, 1, COL_NEXT), YELLOW),
        fill(rng(f0, last, PLAN_COL0 + 1, PLAN_COL0 + plan_w), YELLOW),
        date_fmt(rng(f0, last, COL_NEXT, COL_NEXT + 1)),
        date_fmt(rng(f0, last, PLAN_COL0, PLAN_COL0 + 1)),
        width(0, 1, 40), width(1, 2, 190), width(2, 3, 140), width(3, 13, 105),
        width(13, 14, 170), width(14, 15, 260), width(15, 16, 120), width(16, 17, 24),
        width(PLAN_COL0, PLAN_COL0 + 1, 95), width(PLAN_COL0 + 1, PLAN_COL0 + 7, 100),
        width(PLAN_COL0 + 7, PLAN_COL0 + 8, 260),
        {"updateDimensionProperties": {"range": {"sheetId": sheet_id, "dimension": "ROWS",
                                                 "startIndex": h0, "endIndex": HEADER_ROW},
                                       "properties": {"pixelSize": 48}, "fields": "pixelSize"}},
        {"setDataValidation": {"range": rng(f0, last, COL_STATUS, COL_STATUS + 1), "rule": {
            "condition": {"type": "ONE_OF_LIST", "values": [{"userEnteredValue": s} for s in STATUSES]},
            "showCustomUi": True, "strict": False}}},
        {"setDataValidation": {"range": rng(f0, last, PLAN_COL0 + 1, PLAN_COL0 + 7), "rule": {
            "condition": {"type": "NUMBER_GREATER_THAN_EQ", "values": [{"userEnteredValue": "0"}]},
            "inputMessage": "Часы числом, например 1,5", "strict": True}}},
        {"addConditionalFormatRule": {"index": 0, "rule": {
            "ranges": [rng(f0, last, COL_NAME, COL_NAME + 1), rng(f0, last, COL_NEXT, COL_NEXT + 1)],
            "booleanRule": {"condition": {"type": "CUSTOM_FORMULA", "values": [{"userEnteredValue": no_touch}]},
                            "format": {"backgroundColor": RED_BG,
                                       "textFormat": {"foregroundColor": RED_TEXT, "bold": True}}}}}},
        {"addConditionalFormatRule": {"index": 1, "rule": {
            "ranges": [rng(f0, last, PLAN_COL0, PLAN_COL0 + plan_w)],
            "booleanRule": {"condition": {"type": "CUSTOM_FORMULA", "values": [
                {"userEnteredValue": f"=${plan_letter}{FIRST_ROW}=TODAY()"}]},
                "format": {"backgroundColor": TODAY_BG}}}}},
    ]
    for c in DATE_COLS:
        reqs.append(date_fmt(rng(f0, last, c, c + 1)))
        reqs.append({"setDataValidation": {"range": rng(f0, last, c, c + 1), "rule": {
            "condition": {"type": "DATE_IS_VALID"}, "inputMessage": "Выберите дату в календаре (двойной клик)",
            "strict": True}}})
    return reqs
