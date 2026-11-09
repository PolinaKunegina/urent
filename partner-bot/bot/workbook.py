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
TIME_HEADER, STAGE_HEADER = "Время, мин", "Этап"
FMT_HEADER, MIN_HEADER = "Формат (этап)", "Мин"


def touch_headers(n: int) -> list[str]:
    return [f"Дата {n}-го касания", FMT_HEADER, MIN_HEADER]


CRM_HEADERS = [
    "№", "Имя клиента", "Откуда пришёл",
    "Следующий шаг: дата", "Следующий шаг: что", TIME_HEADER, STAGE_HEADER,
    "Текущий статус", "Что происходит",
    *[h for n in range(1, 6) for h in touch_headers(n)],
]
COL_NAME, COL_SOURCE, COL_NEXT_DATE, COL_NEXT_WHAT = 1, 2, 3, 4
COL_TIME, COL_STAGE, COL_STATUS, COL_DESC = 5, 6, 7, 8
HIST_COL0 = 9                  # J — первая «Дата касания»
# история: (n, дата, «формат (этап)», минуты)
TOUCH_COLS = [(n, HIST_COL0 + 3 * (n - 1), HIST_COL0 + 3 * (n - 1) + 1, HIST_COL0 + 3 * (n - 1) + 2)
              for n in range(1, 6)]
DATE_COLS = [t[1] for t in TOUCH_COLS]
STAGES = ["Привлечение", "Прогрев", "Закрытие"]
STAGE_OF = {"attract": "привлечение", "warmup": "прогрев", "close": "закрытие"}

PLAN_COL0 = len(CRM_HEADERS) + 1   # колонка Z — начало блока часов (Y — узкий разделитель)
PLAN_HEADERS = [
    "Дата",
    "Привлечение — план, ч", "Прогрев — план, ч", "Закрытие — план, ч", "План: что буду делать",
    "Привлечение — факт, ч", "Прогрев — факт, ч", "Закрытие — факт, ч",
    "Привлечение — людей", "Прогрев — людей", "Закрытие — людей",
    "Факт: что сделал(а)", "% плана",
]
# смещения от PLAN_COL0
P_PLAN, P_PLAN_NOTE, P_FACT, P_PEOPLE, P_FACT_NOTE, P_PCT = 1, 4, 5, 8, 11, 12
PROCESSES = ["attract", "warmup", "close"]
PROCESS_NAMES = {"attract": "привлечение", "warmup": "прогрев", "close": "закрытие"}

STATUSES = [
    "Новый контакт", "Был на встрече", "Приглашён на следующую встречу",
    "Назначена диагностика", "Прошёл диагностику",
    "Купил маленький продукт (не оргвзнос)", "Оплатил 95", "Оплатил 240",
    "Коллаборация / эксперт", "Отказ",
]
# старые статусы → новые (для перевода уже заполненных строк)
STATUS_MIGRATION = {
    "назначено касание": "Приглашён на следующую встречу",
    "думает": "Прошёл диагностику",
    "диагностика": "Назначена диагностика",
    "стал партнёром": "Оплатил 240",
}
EXCLUDED_RE = r"отказ"                                                  # не в работе
CLOSED_RE = r"оплатил 240|стал партн|стал ментор|коллаборац|оплатил пакет"  # шаг не нужен, не красный

# Блок «Мой месяц» (строки 1–3): цели вписывает партнёр, результаты считает бот по смене статусов
RESULT_CATS = [   # (ключ, подпись, по какому статусу засчитывается; None — считается иначе)
    ("leads", "Новые контакты", None),           # первое касание клиента в этом месяце
    ("c95", "Партнёры 95", r"(?<!\d)95(?!\d)"),
    ("c240", "Партнёры 240", r"(?<!\d)240(?!\d)"),
    ("c5000", "Маленький продукт", r"маленький продукт|купил продукт|интенсив|5000"),   # любая покупка, кроме оргвзноса
]
MONTH_COL = 7                  # H: подписи блока; I..L — категории; M — % выполнения
MONTHS_NOM = ["январь", "февраль", "март", "апрель", "май", "июнь", "июль", "август", "сентябрь",
              "октябрь", "ноябрь", "декабрь"]


def result_cats(status: str) -> set[str]:
    st = (status or "").lower()
    return {k for k, _, rx in RESULT_CATS if rx and re.search(rx, st)}


def new_leads_in_month(tab, today: date) -> int:
    """Клиенты, чьё самое раннее касание пришлось на текущий месяц (не позже сегодня)."""
    n = 0
    for c in tab.clients:
        dates = [x.when for x in c.touches if x.when]
        if dates and min(dates).year == today.year and min(dates).month == today.month and min(dates) <= today:
            n += 1
    return n


SERVICE_TABS = {"Оглавление", "Сводка", "Партнёры", "Инструкция", "Итоги недели"}

HINT_CLIENTS = ("Следующий шаг сам попадает в историю вместе с минутами и этапом. "
                "Красный — нет шага на 7 дней.")
HINT_HOURS = ("Факт и людей считает таблица из минут у клиентов. Вы пишете план в голубой строке (завтра) "
              "и «Факт: что сделал(а)» в зелёной (сегодня). Серые цифры — уже назначенные шаги будущих дней.")


@dataclass
class DayPlan:
    day: date
    plan: dict[str, float]     # процесс -> часы (только заполненные)
    fact: dict[str, float]
    fact_note: str = ""                                       # «Факт: что сделал(а)»
    people: dict[str, float] = field(default_factory=dict)   # сколько людей в работе (факт)
    plan_note: str = ""                                       # «План: что буду делать»
    row: int = 0                                              # номер строки на листе (1-based)

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
    touch_cols: list = field(default_factory=list)   # колонки истории [(n, дата, формат, минуты|None)]
    goals: dict = field(default_factory=dict)        # категория -> цель месяца (вписывает партнёр)
    results: dict = field(default_factory=dict)      # категория -> результат месяца (пишет бот)
    month_col: int | None = None                     # колонка подписи «Мой месяц»
    plan_col: int | None = None                      # колонка «Дата» блока часов
    time_col: int | None = None                      # «Время, мин» (нет — вкладка старого макета)
    month_cols: dict = field(default_factory=dict)   # категория -> колонка её подписи в «Мой месяц»
    month_ok: bool = True                            # подписи стоят подряд, как задумано
    month_end: int = 0                               # колонка за последней ячейкой блока
    next_what_col: int | None = None

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
            rows = list(ws.iter_rows(min_row=1, values_only=True))
            tab = _parse_rows(title, rows[header_row - 1:], sandbox)
            _parse_month(tab, rows[:header_row - 1])
            out.append(tab)
    finally:
        wb.close()
    return out


def _parse_month(tab: PartnerTab, top: list[tuple]) -> None:
    """Блок «Мой месяц»: подпись в строке 1, цели строкой ниже, результаты ещё ниже.

    Категорию ищем по её названию: если в блок вставили колонку, цифры всё равно читаются
    из-под своих подписей, а month_ok = False говорит, что блок пора собрать обратно."""
    for r, row in enumerate(top):
        for c, v in enumerate(row):
            if _norm(v).startswith("мой месяц"):
                tab.month_col = c
                labels = {_norm(x): i for i, x in enumerate(row) if i > c and x not in (None, "")}
                ok = True
                for k, (key, title, _) in enumerate(RESULT_CATS):
                    col = labels.get(_norm(title), c + 1 + k)
                    ok = ok and col == c + 1 + k
                    tab.month_cols[key] = col
                    get = lambda rr: top[rr][col] if rr < len(top) and col < len(top[rr]) else None
                    g, res = _num(get(r + 1)), _num(get(r + 2))
                    if g is not None:
                        tab.goals[key] = g
                    if res is not None:
                        tab.results[key] = res
                done = labels.get("выполнено")
                tab.month_ok = ok and done in (None, c + 1 + len(RESULT_CATS))
                tab.month_end = max([*tab.month_cols.values(), done or 0]) + 1
                return


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
            mn = i + 2 if f is not None and i + 2 < len(head) and head[i + 2].startswith("мин") else None
            touch_cols.append((int(m.group(1)), i, f, mn))
    c_date = next((i for i, h in enumerate(head) if h == "дата"), None)
    plan_cols = {p: find(f"{PROCESS_NAMES[p]} — план") for p in PROCESSES}
    fact_cols = {p: find(f"{PROCESS_NAMES[p]} — факт") for p in PROCESSES}
    people_cols = {p: find(f"{PROCESS_NAMES[p]} — людей") for p in PROCESSES}
    c_plan_note = find("план: что")
    c_fact_note = find("факт: что") if find("факт: что") is not None else find("что делал")

    def cell(row, i):
        return row[i] if i is not None and i < len(row) else None

    tab = PartnerTab(name=title, sandbox=sandbox, touch_cols=touch_cols, plan_col=c_date,
                     time_col=find("время"), next_what_col=c_next_what)
    for offset, row in enumerate(rows[1:], 1):
        r_num = HEADER_ROW + offset
        name = _clean(cell(row, c_name))
        if name:
            touches = []
            for n, d_i, f_i, m_i in touch_cols:
                raw, fmt = cell(row, d_i), cell(row, f_i)
                if raw in (None, "") and fmt in (None, ""):
                    continue
                touches.append(Touch(n=n, when=to_date(raw), raw=_clean(raw), fmt=_clean(fmt),
                                     minutes=_num(cell(row, m_i))))
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
            plan_note, fact_note = _clean(cell(row, c_plan_note)), _clean(cell(row, c_fact_note))
            if c_plan_note is None and fact_note.lower().startswith("план"):   # старый общий комментарий
                plan_note, fact_note = fact_note, ""
            tab.days[d] = DayPlan(d, plan, fact, fact_note, people, plan_note, row=r_num)
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
MONTH_BG = {"red": 0.99, "green": 0.95, "blue": 0.8}
FACT_BG = {"red": 0.85, "green": 0.94, "blue": 0.83}
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


def link_cell(sheet_id: int, row: int, col: int, parts: list[tuple[str, str | None]],
              size: int = 11, bold: bool = False) -> dict:
    """Ячейка с обычными (не формульными) ссылками — так переход внутри таблицы работает
    надёжно, как ссылка, вставленная через Ctrl+K. parts = [(текст, uri или None), ...]."""
    text, runs, pos = "", [], 0
    for chunk, uri in parts:
        fmt = {"fontSize": size, "bold": bold}
        if uri:
            fmt.update({"link": {"uri": uri}, "foregroundColor": LINK_BLUE, "underline": True})
        runs.append({"startIndex": pos, "format": fmt})
        text += chunk
        pos += len(chunk)
    return {"updateCells": {
        "rows": [{"values": [{"userEnteredValue": {"stringValue": text}, "textFormatRuns": runs}]}],
        "fields": "userEnteredValue,textFormatRuns",
        "start": {"sheetId": sheet_id, "rowIndex": row, "columnIndex": col}}}


def partner_links(gid: int, plan_col: int, toc_gid: int | None, guide_gid: int | None) -> list[dict]:
    """Переходы на вкладке партнёра: к часам и обратно, к оглавлению и инструкции."""
    reqs = [
        link_cell(gid, 0, COL_NEXT_DATE, [("→ к моим часам", f"#gid={gid}&range={col_letter(plan_col)}{HEADER_ROW}")],
                  size=12, bold=True),
        link_cell(gid, 0, plan_col + 3, [("← к клиентам", f"#gid={gid}&range=A{HEADER_ROW}")], size=12, bold=True),
    ]
    nav = []
    if toc_gid is not None:
        nav.append(("← Оглавление", f"#gid={toc_gid}"))
    if guide_gid is not None:
        nav += [("   ", None), ("📖 Инструкция", f"#gid={guide_gid}")]
    if nav:
        reqs.append(link_cell(gid, 2, COL_NAME, nav))   # колонка B закреплена — ссылки всегда на виду
    return reqs


def header_protection_ranges(sheet_id: int, width: int) -> list[dict]:
    """Строки 1–4 под защитой, кроме ячеек целей месяца (их вписывает партнёр)."""
    g0, g1 = MONTH_COL + 1, MONTH_COL + 1 + len(RESULT_CATS)

    def r(r0, r1, c0, c1):
        return {"sheetId": sheet_id, "startRowIndex": r0, "endRowIndex": r1,
                "startColumnIndex": c0, "endColumnIndex": c1}
    return [r(0, 1, 0, width), r(1, 2, 0, g0), r(1, 2, g1, width), r(2, HEADER_ROW, 0, width)]


def month_block_values(today: date, sep: str = ";") -> list[list]:
    m = MONTHS_NOM[today.month - 1]
    goal_l, goal_r = col_letter(MONTH_COL + 1), col_letter(MONTH_COL + len(RESULT_CATS))
    pct = fx(f'=IF(SUM({goal_l}2:{goal_r}2)=0,"",SUM({goal_l}3:{goal_r}3)/SUM({goal_l}2:{goal_r}2))', sep)
    return [[f"Мой месяц: {m}", *[t for _, t, _ in RESULT_CATS], "Выполнено"],
            ["Цель (впишите)", *[""] * len(RESULT_CATS), ""],
            ["Результат (сам)", *[0] * len(RESULT_CATS), pct]]


def month_block_requests(sheet_id: int) -> list[dict]:
    n = len(RESULT_CATS)

    def rng(r0, r1, c0, c1):
        return {"sheetId": sheet_id, "startRowIndex": r0, "endRowIndex": r1,
                "startColumnIndex": c0, "endColumnIndex": c1}

    def cell(r, fmt, fields):
        return {"repeatCell": {"range": r, "cell": {"userEnteredFormat": fmt}, "fields": fields}}
    return [
        cell(rng(0, 3, MONTH_COL, MONTH_COL + 1), {"textFormat": {"bold": True, "fontSize": 10},
                                                     "backgroundColor": MONTH_BG},
             "userEnteredFormat(textFormat,backgroundColor)"),
        cell(rng(0, 1, MONTH_COL + 1, MONTH_COL + n + 2), {"textFormat": {"bold": True, "fontSize": 9},
                                                             "backgroundColor": MONTH_BG,
                                                             "horizontalAlignment": "CENTER"},
             "userEnteredFormat(textFormat,backgroundColor,horizontalAlignment)"),
        cell(rng(1, 2, MONTH_COL + 1, MONTH_COL + n + 1), {"backgroundColor": YELLOW, "horizontalAlignment": "CENTER",
                                                             "textFormat": {"bold": True, "fontSize": 11}},
             "userEnteredFormat(backgroundColor,horizontalAlignment,textFormat)"),
        cell(rng(2, 3, MONTH_COL + 1, MONTH_COL + n + 1), {"backgroundColor": FACT_BG, "horizontalAlignment": "CENTER",
                                                             "textFormat": {"bold": True, "fontSize": 11}},
             "userEnteredFormat(backgroundColor,horizontalAlignment,textFormat)"),
        cell(rng(2, 3, MONTH_COL + n + 1, MONTH_COL + n + 2), {"numberFormat": {"type": "PERCENT", "pattern": "0%"},
                                                                 "horizontalAlignment": "CENTER",
                                                                 "textFormat": {"bold": True, "fontSize": 11}},
             "userEnteredFormat(numberFormat,horizontalAlignment,textFormat)"),
        {"setDataValidation": {"range": rng(1, 2, MONTH_COL + 1, MONTH_COL + n + 1), "rule": {
            "condition": {"type": "NUMBER_GREATER_THAN_EQ", "values": [{"userEnteredValue": "0"}]},
            "inputMessage": "Сколько хотите за месяц — числом", "strict": True}}},
    ]


def fact_formulas(r: int, plan_col: int, hist0: int = HIST_COL0, sep: str = ";") -> list[str]:
    """Формулы строки дня r (1-based): часы по трём процессам из минут у клиентов, потом людей.

    История — тройки «дата / формат (этап) / мин». Берём три одинаковых диапазона, сдвинутых
    на колонку: даты, форматы, минуты — так SUMIFS складывает минуты касаний этого дня с нужным
    этапом. Диапазоны доходят до разделителя и чуть дальше, чтобы новые колонки касаний,
    вставленные перед разделителем, сами попадали в формулы.
    """
    sep_col = plan_col - 1
    a, b = FIRST_ROW, HEADER_ROW + CLIENT_ROWS

    def rng(c0, c1):
        return f"${col_letter(c0)}${a}:${col_letter(c1)}${b}"
    dates, fmts, mins = rng(hist0, sep_col), rng(hist0 + 1, sep_col + 1), rng(hist0 + 2, sep_col + 2)
    d = f"${col_letter(plan_col)}{r}"
    hours = [fx(f'=IFERROR(1/(1/ROUND(SUMIFS({mins},{dates},{d},{fmts},"*({STAGE_OF[p]})*")/60,2)),"")', sep)
             for p in PROCESSES]
    people = [fx(f'=IFERROR(1/(1/COUNTIFS({dates},{d},{fmts},"*({STAGE_OF[p]})*")),"")', sep)
              for p in PROCESSES]
    return hours + people


def partner_tab_values(name: str, start: date, end: date, clients: list[dict] | None = None,
                       sep: str = ";", gid: int | None = None,
                       days: dict | None = None) -> list[list]:
    """Значения вкладки целиком (USER_ENTERED): заголовки, ссылки-переходы, даты, клиенты, часы."""
    width = PLAN_COL0 + len(PLAN_HEADERS)
    grid = [["" for _ in range(width)] for _ in range(HEADER_ROW + CLIENT_ROWS)]
    grid[0][0] = f"Клиенты — {name}"
    grid[1][0] = HINT_CLIENTS
    for row, cells in enumerate(month_block_values(start, sep)):
        for k, v in enumerate(cells):
            grid[row][MONTH_COL + k] = v
    grid[0][PLAN_COL0] = "Мои часы по дням"
    grid[1][PLAN_COL0] = HINT_HOURS
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
        r = FIRST_ROW + k
        pl = f"{col_letter(PLAN_COL0 + P_PLAN)}{r}:{col_letter(PLAN_COL0 + P_PLAN + 2)}{r}"
        fa = f"{col_letter(PLAN_COL0 + P_FACT)}{r}:{col_letter(PLAN_COL0 + P_FACT + 2)}{r}"
        row[PLAN_COL0 + P_PCT] = fx(f'=IF(OR(SUM({pl})=0,COUNT({fa})=0),"",SUM({fa})/SUM({pl}))', sep)
        formulas = fact_formulas(r, PLAN_COL0, HIST_COL0, sep)
        for i in range(6):
            row[PLAN_COL0 + P_FACT + i] = formulas[i]
        dp = (days or {}).get(d)
        if dp:
            for i, p in enumerate(PROCESSES):
                if p in dp.plan:
                    row[PLAN_COL0 + P_PLAN + i] = dp.plan[p]
            if dp.fact or dp.people:      # факт, внесённый руками до формул, сохраняем как есть
                for i, p in enumerate(PROCESSES):
                    row[PLAN_COL0 + P_FACT + i] = dp.fact.get(p, "")
                    row[PLAN_COL0 + P_PEOPLE + i] = dp.people.get(p, "")
            row[PLAN_COL0 + P_PLAN_NOTE] = dp.plan_note
            row[PLAN_COL0 + P_FACT_NOTE] = dp.fact_note
    for k, c in enumerate(clients or []):
        row = grid[FIRST_ROW - 1 + k]
        row[COL_NAME] = c.get("name", "")
        row[COL_SOURCE] = c.get("source", "")
        row[COL_NEXT_DATE] = c.get("next_date", "")
        row[COL_NEXT_WHAT] = c.get("next_what", "")
        row[COL_STATUS] = c.get("status", "")
        row[COL_DESC] = c.get("description", "")
        row[COL_TIME] = c.get("minutes", "")
        row[COL_STAGE] = c.get("stage", "")
        for t in c.get("touches", []):
            slot = next((tc for tc in TOUCH_COLS if tc[0] == t["n"]), None)
            if slot:
                row[slot[1]] = t["date"]
                row[slot[2]] = t.get("fmt", "")
                row[slot[3]] = "" if t.get("minutes") is None else t["minutes"]
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
        fill(rng(h0, HEADER_ROW, 0, crm_w), HEAD_BG, bold=True, wrap=True),
        fill(rng(h0, HEADER_ROW, COL_NEXT_DATE, COL_STAGE + 1), NEXT_BG, bold=True, wrap=True),
        fill(rng(h0, HEADER_ROW, PLAN_COL0, PLAN_COL0 + plan_w), PLAN_BG, bold=True, wrap=True),
        fill(rng(f0, last, 1, crm_w), YELLOW),
        fill(rng(f0, last, PLAN_COL0 + 1, PLAN_COL0 + plan_w), YELLOW),
        date_fmt(rng(f0, last, COL_NEXT_DATE, COL_NEXT_DATE + 1)),
        date_fmt(rng(f0, last, PLAN_COL0, PLAN_COL0 + 1)),
        width(0, 1, 40), width(1, 2, 190), width(2, 3, 140), width(3, 4, 110), width(4, 5, 190),
        width(COL_STATUS, COL_STATUS + 1, 170), width(COL_DESC, COL_DESC + 1, 260),
        width(crm_w, crm_w + 1, 24),
        width(PLAN_COL0, PLAN_COL0 + 1, 95), width(PLAN_COL0 + P_PLAN, PLAN_COL0 + P_PLAN_NOTE, 95),
        width(PLAN_COL0 + P_PLAN_NOTE, PLAN_COL0 + P_PLAN_NOTE + 1, 220),
        width(PLAN_COL0 + P_FACT, PLAN_COL0 + P_PEOPLE, 95),
        width(PLAN_COL0 + P_PEOPLE, PLAN_COL0 + P_FACT_NOTE, 85),
        width(PLAN_COL0 + P_FACT_NOTE, PLAN_COL0 + P_FACT_NOTE + 1, 240),
        width(PLAN_COL0 + P_PCT, PLAN_COL0 + P_PCT + 1, 75),
        fill(rng(h0, HEADER_ROW, PLAN_COL0 + P_FACT, PLAN_COL0 + P_PEOPLE), FACT_BG, bold=True, wrap=True),
        fill(rng(h0, HEADER_ROW, PLAN_COL0 + P_FACT_NOTE, PLAN_COL0 + P_FACT_NOTE + 1), FACT_BG, bold=True, wrap=True),
        fill(rng(h0, HEADER_ROW, PLAN_COL0 + P_PEOPLE, PLAN_COL0 + P_FACT_NOTE), PEOPLE_BG, bold=True, wrap=True),
        fill(rng(h0, HEADER_ROW, PLAN_COL0 + P_PCT, PLAN_COL0 + P_PCT + 1), HEAD_BG, bold=True, wrap=True),
        fill(rng(f0, last, PLAN_COL0 + P_PCT, PLAN_COL0 + P_PCT + 1), {"red": 1, "green": 1, "blue": 1}),
        {"repeatCell": {"range": rng(f0, last, PLAN_COL0 + P_PCT, PLAN_COL0 + P_PCT + 1), "cell": {
            "userEnteredFormat": {"numberFormat": {"type": "PERCENT", "pattern": "0%"},
                                  "horizontalAlignment": "CENTER"}},
            "fields": "userEnteredFormat.numberFormat,userEnteredFormat.horizontalAlignment"}},
        {"updateDimensionProperties": {"range": {"sheetId": sheet_id, "dimension": "ROWS",
                                                 "startIndex": h0, "endIndex": HEADER_ROW},
                                       "properties": {"pixelSize": 48}, "fields": "pixelSize"}},
        {"setDataValidation": {"range": rng(f0, last, COL_STATUS, COL_STATUS + 1), "rule": {
            "condition": {"type": "ONE_OF_LIST", "values": [{"userEnteredValue": s} for s in STATUSES]},
            "showCustomUi": True, "strict": False}}},
        *[{"setDataValidation": {"range": rng(f0, last, PLAN_COL0 + c0, PLAN_COL0 + c0 + 3), "rule": {
            "condition": {"type": "NUMBER_GREATER_THAN_EQ", "values": [{"userEnteredValue": "0"}]},
            "inputMessage": "Часы числом, например 1,5", "strict": True}}} for c0 in (P_PLAN,)],
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
        for r in (*header_protection_ranges(sheet_id, PLAN_COL0 + plan_w),
                  rng(f0, last, PLAN_COL0, PLAN_COL0 + 1),
                  rng(f0, last, 0, 1), rng(f0, last, PLAN_COL0 + P_PCT, PLAN_COL0 + P_PCT + 1)):
            reqs.append({"addProtectedRange": {"protectedRange": {
                "range": r, "description": "bot:заголовки", "warningOnly": False,
                "editors": {"users": [service_email]}}}})
    reqs += time_stage_requests(sheet_id, PLAN_COL0, TOUCH_COLS, service_email, sep)
    return reqs


def time_stage_requests(sheet_id: int, plan_col: int, touch_cols: list, service_email: str = "",
                        sep: str = ";") -> list[dict]:
    """Колонки «Время, мин» и «Этап», минуты в истории, факт-формулы под защитой и серые будущие дни."""
    f0, last = FIRST_ROW - 1, HEADER_ROW + CLIENT_ROWS

    def rng(c0, c1, r0=f0, r1=last):
        return {"sheetId": sheet_id, "startRowIndex": r0, "endRowIndex": r1,
                "startColumnIndex": c0, "endColumnIndex": c1}

    def width(c, px):
        return {"updateDimensionProperties": {"range": {"sheetId": sheet_id, "dimension": "COLUMNS",
                                                        "startIndex": c, "endIndex": c + 1},
                                              "properties": {"pixelSize": px}, "fields": "pixelSize"}}

    def minutes_rule(c, msg):
        return {"setDataValidation": {"range": rng(c, c + 1), "rule": {
            "condition": {"type": "NUMBER_GREATER_THAN_EQ", "values": [{"userEnteredValue": "0"}]},
            "inputMessage": msg, "strict": True}}}

    fa0, fa1 = plan_col + P_FACT, plan_col + P_FACT_NOTE      # факт, ч ×3 и людей ×3
    tm, stg = col_letter(COL_TIME), col_letter(COL_STAGE)
    reqs = [
        width(COL_TIME, 75), width(COL_STAGE, 120),
        {"repeatCell": {"range": rng(COL_TIME, COL_STAGE + 1, HEADER_ROW - 1, HEADER_ROW), "cell": {
            "userEnteredFormat": {"backgroundColor": NEXT_BG, "textFormat": {"bold": True},
                                  "wrapStrategy": "WRAP", "verticalAlignment": "MIDDLE"}},
            "fields": "userEnteredFormat.backgroundColor,userEnteredFormat.textFormat.bold,"
                      "userEnteredFormat.wrapStrategy,userEnteredFormat.verticalAlignment"}},
        {"repeatCell": {"range": rng(COL_TIME, COL_STAGE + 1), "cell": {"userEnteredFormat": {
            "backgroundColor": YELLOW, "numberFormat": {"type": "NUMBER", "pattern": "0"}}},
            "fields": "userEnteredFormat.backgroundColor,userEnteredFormat.numberFormat"}},
        minutes_rule(COL_TIME, "Сколько минут займёт этот шаг (или занял): 15, 30, 60"),
        {"setDataValidation": {"range": rng(COL_STAGE, COL_STAGE + 1), "rule": {
            "condition": {"type": "ONE_OF_LIST", "values": [{"userEnteredValue": x} for x in STAGES]},
            "inputMessage": "На каком этапе человек: привлечение, прогрев или закрытие",
            "showCustomUi": True, "strict": True}}},
        {"addConditionalFormatRule": {"index": 0, "rule": {   # минуты есть, этапа нет — время не посчитается
            "ranges": [rng(COL_STAGE, COL_STAGE + 1)],
            "booleanRule": {"condition": {"type": "CUSTOM_FORMULA", "values": [
                {"userEnteredValue": fx(f'=AND(${tm}{FIRST_ROW}<>"",${stg}{FIRST_ROW}="")', sep)}]},
                "format": {"backgroundColor": RED_BG}}}}},
        {"addConditionalFormatRule": {"index": 0, "rule": {   # будущие дни: это назначенные шаги, не факт
            "ranges": [rng(fa0, fa1)],
            "booleanRule": {"condition": {"type": "CUSTOM_FORMULA", "values": [
                {"userEnteredValue": f"=${col_letter(plan_col)}{FIRST_ROW}>TODAY()"}]},
                "format": {"textFormat": {"foregroundColor": {"red": 0.6, "green": 0.6, "blue": 0.6},
                                          "italic": True}}}}}},
    ]
    for t in touch_cols:
        if len(t) > 3 and t[3] is not None:
            reqs += [width(t[2], 150), width(t[3], 50), minutes_rule(t[3], "Минуты этого касания")]
        reqs.append(width(t[1], 100))
    reqs.append({"setDataValidation": {"range": rng(fa0, fa1)}})   # в формулах проверка не нужна
    if service_email:
        reqs.append({"addProtectedRange": {"protectedRange": {
            "range": rng(fa0, fa1), "description": "bot:факт-формулы", "warningOnly": False,
            "editors": {"users": [service_email]}}}})
    return reqs
