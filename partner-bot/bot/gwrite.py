"""Запись обратно в Google Таблицу: подсветка лидов без плана и лист с часами из чата.

Работает только с родной Google Таблицей (не с загруженным .xlsx) и только если
сервисному аккаунту дан доступ «Редактор».

Подсветка устроена так, чтобы не трогать заливку, которую ставят партнёры:
бот пишет в служебную колонку «След. касание (бот)» дату ближайшего касания или
«нет плана», а условное форматирование (ставится один раз) красит имя лида
и эту ячейку, пока там «нет плана». Как только партнёр впишет дату, при следующем
обновлении красный уходит сам.
"""
from __future__ import annotations

import logging

from .sheets import PARTNERS_HEADER, Funnel

log = logging.getLogger(__name__)

SHEETS_API = "https://sheets.googleapis.com/v4/spreadsheets"
NO_PLAN = "нет плана"
RED_BG = {"red": 0.957, "green": 0.8, "blue": 0.8}
RED_TEXT = {"red": 0.6, "green": 0.0, "blue": 0.0}


def col_letter(idx: int) -> str:
    """0 -> A, 25 -> Z, 26 -> AA."""
    s, n = "", idx + 1
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def helper_col_index(f: Funnel, helper_header: str) -> tuple[int, bool]:
    """(индекс колонки, есть ли уже заголовок)."""
    target = helper_header.lower()
    if target in f.headers:
        return f.headers.index(target), True
    last = max((i for i, h in enumerate(f.headers) if h), default=0)
    return last + 1, False


class SheetsWriter:
    def __init__(self, session, spreadsheet_id: str):
        self.s = session
        self.id = spreadsheet_id

    def _meta(self) -> dict:
        r = self.s.get(f"{SHEETS_API}/{self.id}", params={
            "fields": "sheets(properties(sheetId,title),conditionalFormats)"}, timeout=30)
        r.raise_for_status()
        return r.json()

    def _sheet(self, title: str) -> dict | None:
        return next((sh for sh in self._meta().get("sheets", [])
                     if sh["properties"]["title"] == title), None)

    def _batch(self, requests: list[dict]) -> None:
        if requests:
            r = self.s.post(f"{SHEETS_API}/{self.id}:batchUpdate",
                            json={"requests": requests}, timeout=30)
            r.raise_for_status()

    def _values(self, rng: str, values: list[list]) -> None:
        url = f"{SHEETS_API}/{self.id}/values/{rng}"
        r = self.s.put(url, params={"valueInputOption": "RAW"},
                       json={"range": rng, "values": values}, timeout=30)
        r.raise_for_status()

    # ---------- подсветка ----------

    def mark_plan(self, f: Funnel, plan: dict[int, str], helper_header: str) -> None:
        """plan: {номер строки: «03.11.2026» | «нет плана» | «отказ»}."""
        h_idx, has_header = helper_col_index(f, helper_header)
        L = col_letter(h_idx)
        tab = f"'{f.tab}'"
        if not has_header:
            self._values(f"{tab}!{L}{f.header_row}", [[helper_header]])
        first = f.header_row + 1
        last = max(f.last_row, max(plan, default=first), first)
        values = [[plan.get(r, "")] for r in range(first, last + 1)]
        self._values(f"{tab}!{L}{first}:{L}{last}", values)
        self._ensure_rule(f, h_idx, first)

    def _ensure_rule(self, f: Funnel, h_idx: int, first_row: int) -> None:
        sheet = self._sheet(f.tab)
        if sheet is None:
            return
        for rule in sheet.get("conditionalFormats", []):
            cond = rule.get("booleanRule", {}).get("condition", {})
            if any(NO_PLAN in v.get("userEnteredValue", "") for v in cond.get("values", [])):
                return  # правило уже стоит
        sheet_id = sheet["properties"]["sheetId"]
        L = col_letter(h_idx)
        formula = f'=${L}{first_row}="{NO_PLAN}"'
        rng = lambda c: {"sheetId": sheet_id, "startRowIndex": first_row - 1,
                         "startColumnIndex": c, "endColumnIndex": c + 1}
        self._batch([{
            "addConditionalFormatRule": {
                "index": 0,
                "rule": {
                    "ranges": [rng(f.name_col), rng(h_idx)],
                    "booleanRule": {
                        "condition": {"type": "CUSTOM_FORMULA",
                                      "values": [{"userEnteredValue": formula}]},
                        "format": {"backgroundColor": RED_BG,
                                   "textFormat": {"foregroundColor": RED_TEXT, "bold": True}},
                    },
                },
            }
        }])
        log.info("Добавлено правило подсветки: %s", formula)

    # ---------- лист с часами ----------

    def write_hours(self, tab: str, rows: list[list]) -> None:
        if self._sheet(tab) is None:
            self._batch([{"addSheet": {"properties": {"title": tab,
                                                      "gridProperties": {"frozenRowCount": 1}}}}])
        header = ["Дата", "Партнёр", "Процесс", "Часы", "Что делали", "Telegram"]
        r = self.s.post(f"{SHEETS_API}/{self.id}/values/'{tab}'!A:F:clear", json={}, timeout=30)
        r.raise_for_status()
        self._values(f"'{tab}'!A1", [header, *rows])

    # ---------- лист «Партнёры» ----------

    def create_partners_tab(self, tab: str, sources: list[str]) -> None:
        """Создаёт справочник один раз; дальше его правят люди, бот только читает."""
        if self._sheet(tab) is not None:
            return
        self._batch([{"addSheet": {"properties": {"title": tab,
                                                  "gridProperties": {"frozenRowCount": 1}}}}])
        rows = [[name, "", ""] for name in sources]
        self._values(f"'{tab}'!A1", [PARTNERS_HEADER, *rows])
        log.info("Создан лист «%s» с %d вариантами источника", tab, len(rows))
