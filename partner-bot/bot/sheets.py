"""Загрузка таблицы воронки с Google Drive и разбор листа «Воронка лидов».

Файл может быть как загруженным .xlsx, так и родной Google Таблицей:
xlsx скачиваем как есть, Google Таблицу экспортируем в xlsx. Дальше один и тот же
разбор через openpyxl, поэтому партнёрам не нужно ничего конвертировать.
"""
from __future__ import annotations

import io
import re
from dataclasses import dataclass, field
from datetime import date, datetime, timedelta

from openpyxl import load_workbook

DRIVE_API = "https://www.googleapis.com/drive/v3/files"
GSHEET_MIME = "application/vnd.google-apps.spreadsheet"
XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
SCOPES = ["https://www.googleapis.com/auth/drive.readonly"]


@dataclass
class Touch:
    n: int                 # номер касания (1..5)
    when: date | None      # распознанная дата
    raw: str               # как написано в ячейке
    fmt: str = ""          # «Формат» рядом с датой


@dataclass
class Lead:
    row: int
    number: str
    name: str
    source: str = ""
    status: str = ""
    description: str = ""
    touches: list[Touch] = field(default_factory=list)

    def dates(self) -> list[date]:
        return [t.when for t in self.touches if t.when]

    def next_touch(self, today: date) -> Touch | None:
        """Ближайшее касание сегодня или позже."""
        future = [t for t in self.touches if t.when and t.when >= today]
        return min(future, key=lambda t: t.when) if future else None

    def last_touch(self, today: date) -> Touch | None:
        past = [t for t in self.touches if t.when and t.when < today]
        return max(past, key=lambda t: t.when) if past else None

    def bad_dates(self) -> list[Touch]:
        return [t for t in self.touches if t.raw and not t.when]


# ---------- загрузка ----------

def fetch_xlsx(file_id: str, service_account_file: str) -> bytes:
    from google.auth.transport.requests import AuthorizedSession
    from google.oauth2 import service_account

    creds = service_account.Credentials.from_service_account_file(
        service_account_file, scopes=SCOPES
    )
    s = AuthorizedSession(creds)
    meta = s.get(f"{DRIVE_API}/{file_id}", params={
        "fields": "mimeType,name", "supportsAllDrives": "true"}, timeout=30)
    meta.raise_for_status()
    mime = meta.json()["mimeType"]
    if mime == GSHEET_MIME:
        r = s.get(f"{DRIVE_API}/{file_id}/export", params={"mimeType": XLSX_MIME}, timeout=60)
    else:
        r = s.get(f"{DRIVE_API}/{file_id}", params={
            "alt": "media", "supportsAllDrives": "true"}, timeout=60)
    r.raise_for_status()
    return r.content


# ---------- разбор ----------

_DATE_FORMATS = ("%d.%m.%Y", "%d.%m.%y", "%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y")


def to_date(value) -> date | None:
    if value is None or value == "":
        return None
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    if isinstance(value, (int, float)) and 20000 < value < 80000:
        return date(1899, 12, 30) + timedelta(days=int(value))
    text = str(value).strip()
    text = text.split(" ")[0] if re.match(r"^\d{4}-\d{2}-\d{2} ", text) else text
    for fmt in _DATE_FORMATS:
        try:
            return datetime.strptime(text, fmt).date()
        except ValueError:
            pass
    return None


def _clean(v) -> str:
    if v is None:
        return ""
    if isinstance(v, (datetime, date)):
        return v.strftime("%d.%m.%Y")
    return str(v).strip()


def _norm(v) -> str:
    return re.sub(r"\s+", " ", _clean(v)).lower()


def _find_header(ws, max_scan: int = 15):
    for r_idx, row in enumerate(ws.iter_rows(min_row=1, max_row=max_scan, values_only=True), 1):
        if any(_norm(c) == "имя лида" for c in row):
            return r_idx, [_norm(c) for c in row]
    raise ValueError("Не нашла строку заголовков с колонкой «Имя лида»")


def parse_funnel(xlsx: bytes, tab: str = "Воронка лидов") -> list[Lead]:
    wb = load_workbook(io.BytesIO(xlsx), data_only=True, read_only=True)
    if tab in wb.sheetnames:
        ws = wb[tab]
    else:  # подстраховка: первый лист, где есть «Имя лида»
        ws = None
        for name in wb.sheetnames:
            try:
                _find_header(wb[name])
                ws = wb[name]
                break
            except ValueError:
                continue
        if ws is None:
            raise ValueError(f"Нет листа «{tab}» и листа с колонкой «Имя лида»")

    header_row, headers = _find_header(ws)
    col = {}
    touch_cols: list[tuple[int, int, int | None]] = []  # (n, date_idx, fmt_idx)
    for i, h in enumerate(headers):
        if h in ("№", "no", "номер"):
            col.setdefault("number", i)
        elif h == "имя лида":
            col["name"] = i
        elif h == "источник":
            col["source"] = i
        elif "статус" in h:
            col.setdefault("status", i)
        elif h.startswith("описание"):
            col.setdefault("description", i)
        else:
            m = re.match(r"дата\s*(\d+)", h)
            if m:
                fmt_idx = i + 1 if i + 1 < len(headers) and headers[i + 1].startswith("формат") else None
                touch_cols.append((int(m.group(1)), i, fmt_idx))

    def cell(row, key):
        idx = col.get(key)
        return row[idx] if idx is not None and idx < len(row) else None

    leads = []
    for r_idx, row in enumerate(ws.iter_rows(min_row=header_row + 1, values_only=True), header_row + 1):
        name = _clean(cell(row, "name"))
        if not name:
            continue
        touches = []
        for n, d_idx, f_idx in touch_cols:
            raw = row[d_idx] if d_idx < len(row) else None
            fmt = row[f_idx] if f_idx is not None and f_idx < len(row) else None
            if raw in (None, "") and fmt in (None, ""):
                continue
            touches.append(Touch(n=n, when=to_date(raw), raw=_clean(raw), fmt=_clean(fmt)))
        leads.append(Lead(
            row=r_idx,
            number=_clean(cell(row, "number")),
            name=name,
            source=_clean(cell(row, "source")),
            status=_clean(cell(row, "status")),
            description=_clean(cell(row, "description")),
            touches=touches,
        ))
    wb.close()
    return leads
