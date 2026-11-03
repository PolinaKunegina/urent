"""Доступ к таблице: каждый партнёр правит только свою вкладку.

- партнёру выдаётся доступ «Редактор» к файлу по почте с листа «Партнёры» (колонка «Почта»);
- его вкладка защищена: править её могут он сам, руководители и бот;
- служебные листы (оглавление, инструкция, сводка, справочник и прочие не-партнёрские)
  правят только руководители и бот;
- руководители — все, у кого есть доступ «Редактор» к файлу, но кто не партнёр
  (Наталья добавляет их обычным «Поделиться»), плюс владелец файла.

Бот вызывает ensure_access раз в час — новые строки на листе «Партнёры» подхватываются сами.
"""
from __future__ import annotations

import logging
import re
from dataclasses import dataclass, field

log = logging.getLogger(__name__)
DRIVE = "https://www.googleapis.com/drive/v3/files"
SHEETS = "https://sheets.googleapis.com/v4/spreadsheets"
MARK = "bot:"


def split_emails(raw: str) -> list[str]:
    return [x.lower() for x in re.split(r"[\s,;/]+", raw or "") if "@" in x and "." in x.split("@")[-1]]


@dataclass
class AccessReport:
    shared: list[str] = field(default_factory=list)       # кому только что выдан доступ
    failed: dict[str, str] = field(default_factory=dict)  # почта -> причина
    protected: list[str] = field(default_factory=list)    # вкладки, где обновлена защита


def _permissions(session, file_id: str) -> list[dict]:
    r = session.get(f"{DRIVE}/{file_id}/permissions", params={
        "fields": "permissions(id,type,role,emailAddress)", "supportsAllDrives": "true"}, timeout=30)
    r.raise_for_status()
    return r.json().get("permissions", [])


def share_writer(session, file_id: str, email: str, notify: bool = False) -> str | None:
    """Выдаёт «Редактор». None — успех, иначе текст причины."""
    r = session.post(f"{DRIVE}/{file_id}/permissions",
                     params={"sendNotificationEmail": str(notify).lower(), "supportsAllDrives": "true"},
                     json={"type": "user", "role": "writer", "emailAddress": email}, timeout=30)
    if r.status_code < 300:
        return None
    try:
        msg = r.json()["error"]["message"]
    except Exception:
        msg = r.text[:200]
    if "does not have a Google Account" in msg or "not a valid" in msg.lower() or "invalid sharing request" in msg.lower():
        return "к этой почте не привязан Google-аккаунт"
    return msg


def ensure_access(session, file_id: str, partners: list[dict], service_email: str,
                  extra_admins: list[str] | None = None, notify: bool = False) -> AccessReport:
    """partners: [{"name", "emails": [...]}] — имя совпадает с названием вкладки."""
    rep = AccessReport()
    perms = _permissions(session, file_id)
    writers = {p["emailAddress"].lower() for p in perms
               if p.get("emailAddress") and p["role"] in ("writer", "owner")}
    owner = {p["emailAddress"].lower() for p in perms if p.get("emailAddress") and p["role"] == "owner"}
    partner_emails = {e for p in partners for e in p["emails"]}

    for p in partners:
        for email in p["emails"]:
            if email in writers:
                continue
            err = share_writer(session, file_id, email, notify)
            if err:
                rep.failed[email] = err
            else:
                rep.shared.append(email)
                writers.add(email)

    admins = sorted((writers - partner_emails - owner - {service_email.lower()}) | set(extra_admins or []))
    base = sorted({service_email.lower(), *admins})

    r = session.get(f"{SHEETS}/{file_id}", params={
        "fields": "sheets(properties(sheetId,title),protectedRanges(protectedRangeId,description,editors,range))"},
        timeout=30)
    r.raise_for_status()
    by_name = {p["name"]: [e for e in p["emails"] if e in writers] for p in partners}
    requests = []
    for sh in r.json().get("sheets", []):
        title, sid = sh["properties"]["title"], sh["properties"]["sheetId"]
        if title.startswith("_"):
            continue
        editors = sorted(set(base) | set(by_name.get(title, [])))
        mine = next((pr for pr in sh.get("protectedRanges", [])
                     if (pr.get("description") or "").startswith(MARK)), None)
        if mine is None:
            requests.append({"addProtectedRange": {"protectedRange": {
                "range": {"sheetId": sid}, "description": f"{MARK}{title}", "warningOnly": False,
                "editors": {"users": editors}}}})
            rep.protected.append(title)
        elif sorted(u.lower() for u in mine.get("editors", {}).get("users", [])) != \
                sorted(set(editors) | owner):
            requests.append({"updateProtectedRange": {"protectedRange": {
                "protectedRangeId": mine["protectedRangeId"], "editors": {"users": editors}},
                "fields": "editors"}})
            rep.protected.append(title)
    if requests:
        r = session.post(f"{SHEETS}/{file_id}:batchUpdate", json={"requests": requests}, timeout=60)
        r.raise_for_status()
    return rep
