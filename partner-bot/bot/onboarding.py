"""Новые партнёры: самозапись в песочницу, добавление руководителем, перевод в основную таблицу.

Песочница — отдельная таблица для новичков на первые 2 недели (они не видят чужих
клиентов). Через 2 недели руководитель переводит партнёра командой /promote: вкладка
со всеми клиентами и часами копируется в основную таблицу.
"""
from __future__ import annotations

import re
from dataclasses import dataclass
from datetime import date, timedelta

from .access import share_writer, split_emails
from .config import Config
from .partners import Directory, _key
from .setup_sheet import PARTNERS, Sheet, ensure_structure, write_toc
from .sheets import PARTNERS_HEADER, fetch_xlsx, google_session, parse_partners
from .workbook import col_letter

API = "https://sheets.googleapis.com/v4/spreadsheets"
SANDBOX_DAYS = 14
FORBIDDEN = re.compile(r"[\[\]*?:/\\'\"]")


class OnboardingError(Exception):
    """Понятная человеку причина, почему не получилось."""


def clean_name(text: str) -> str:
    """«иванова   мария» -> «Иванова Мария». Ошибка, если не похоже на фамилию и имя."""
    name = re.sub(r"\s+", " ", FORBIDDEN.sub("", text or "")).strip()
    words = name.split(" ")
    if not 2 <= len(words) <= 4 or len(name) > 60 or not all(re.fullmatch(r"[\w-]+", w) for w in words) \
            or any(ch.isdigit() for ch in name):
        raise OnboardingError("Напиши, пожалуйста, фамилию и имя, например: Иванова Мария")
    return " ".join(w[:1].upper() + w[1:] for w in words)


def _directory(session, sa: str, sheet_id: str) -> Directory:
    data, _ = fetch_xlsx(sheet_id, sa, session)
    return Directory.from_records(parse_partners(data) or [])


@dataclass
class Added:
    name: str
    tab_url: str
    access: str = ""          # что с доступом по почте: пусто — почту не давали


def parse_email(text: str) -> str:
    """Почта из аргументов команды (не @ник)."""
    found = [e for e in split_emails(text) if not e.startswith("@")]
    return found[0] if found else ""


def _email_col(sh: Sheet) -> str:
    """Буква колонки «Почта» на листе «Партнёры»; если колонки нет — заводит её."""
    head = (sh.get_values(f"'{PARTNERS}'!A1:Z1") or [[]])[0]
    for i, h in enumerate(head):
        if str(h).strip().lower().startswith(("почта", "email", "e-mail")):
            return col_letter(i)
    i = max(len(head), 4)
    sh.put(f"'{PARTNERS}'!{col_letter(i)}1", [[PARTNERS_HEADER[4]]], raw=True)
    return col_letter(i)


def _grant(session, sh: Sheet, name: str, emails: list[str]) -> str:
    """Записывает почту в «Партнёры» и выдаёт «Редактор» на всю таблицу. Возвращает итог для человека."""
    if not emails:
        return ""
    col = _email_col(sh)
    for i, row in enumerate(sh.get_values(f"'{PARTNERS}'!A1:A400"), 1):
        if row and _key(row[0]) == _key(name):
            sh.put(f"'{PARTNERS}'!{col}{i}", [[", ".join(emails)]], raw=True)
            break
    done, failed = [], []
    for e in emails:
        err = share_writer(session, sh.id, e, notify=True)
        (failed.append(f"{e} — {err}") if err else done.append(e))
    out = []
    if done:
        out.append("🔑 Доступ «Редактор» выдан: " + ", ".join(done) + ". Google пришлёт письмо со ссылкой.")
    if failed:
        out.append("⚠️ Не получилось выдать доступ: " + "; ".join(failed))
    return "\n".join(out)


def add_partner(cfg: Config, name: str, telegram: str, today: date, session=None, email: str = "") -> Added:
    if not cfg.sandbox_id:
        raise OnboardingError("Регистрация пока не открыта. Напиши, пожалуйста, Наталье Сергиенко (@Natalya_Sergienko) или Полине Кунегиной (@Polina_Kunegina).")
    session = session or google_session(cfg.service_account_file)
    nick = telegram.strip().lstrip("@")
    for sid in filter(None, [cfg.sandbox_id, cfg.sheet_id]):
        d = _directory(session, cfg.service_account_file, sid)
        same = next((p for p in d.partners
                     if (nick and p.username == nick.lower()) or _key(p.name) == _key(name)), None)
        if same:
            where = "песочнице" if sid == cfg.sandbox_id else "основной таблице"
            raise OnboardingError(f"«{same.name}» уже есть в {where} — вкладка на месте.")
    sh = Sheet(session, cfg.sandbox_id)
    sh.append(f"'{PARTNERS}'!A:D", [[name, f"@{nick}" if nick else "", "", today.isoformat()]])
    d = _directory(session, cfg.service_account_file, cfg.sandbox_id)
    ensure_structure(session, cfg.sandbox_id, d, today, sandbox=True, force_toc=True)
    gid = sh.tabs().get(name)
    access = _grant(session, sh, name, [email] if email else [])
    return Added(name, f"https://docs.google.com/spreadsheets/d/{cfg.sandbox_id}/edit#gid={gid}", access)


def give_access(cfg: Config, name_query: str, email: str, session=None) -> Added:
    """Почта уже добавленному партнёру: ищет его в основной таблице, потом в песочнице."""
    session = session or google_session(cfg.service_account_file)
    for sid in filter(None, [cfg.sheet_id, cfg.sandbox_id]):
        d = _directory(session, cfg.service_account_file, sid)
        p = next((x for x in d.partners if _key(x.name) == _key(name_query)), None)
        if p is None:
            continue
        sh = Sheet(session, sid)
        emails = sorted({*p.emails, email})
        access = _grant(session, sh, p.name, emails)
        gid = sh.tabs().get(p.name)
        return Added(p.name, f"https://docs.google.com/spreadsheets/d/{sid}/edit#gid={gid}", access)
    raise OnboardingError(f"Не нашла «{name_query}» ни в основной таблице, ни в песочнице. "
                          "Имя пишется как название вкладки.")


def promote(cfg: Config, name_query: str, today: date, session=None) -> Added:
    """Переносит вкладку партнёра из песочницы в основную таблицу."""
    if not (cfg.sandbox_id and cfg.sheet_id):
        raise OnboardingError("Нужны обе таблицы: основная (/setsheet) и песочница (/setsandbox).")
    session = session or google_session(cfg.service_account_file)
    sa = cfg.service_account_file
    sand = Sheet(session, cfg.sandbox_id)
    d = _directory(session, sa, cfg.sandbox_id)
    p = next((x for x in d.partners if _key(x.name) == _key(name_query)), None)
    if p is None:
        raise OnboardingError(f"Не нашла «{name_query}» в песочнице. Имя пишется как название вкладки.")
    gid = sand.tabs().get(p.name)
    if gid is None:
        raise OnboardingError(f"У «{p.name}» ещё нет вкладки в песочнице.")
    main = Sheet(session, cfg.sheet_id)
    if p.name in main.tabs():
        raise OnboardingError(f"В основной таблице уже есть вкладка «{p.name}».")

    r = session.post(f"{API}/{cfg.sandbox_id}/sheets/{gid}:copyTo",
                     json={"destinationSpreadsheetId": cfg.sheet_id}, timeout=60)
    r.raise_for_status()
    new_gid = r.json()["sheetId"]
    main.batch([{"updateSheetProperties": {"properties": {"sheetId": new_gid, "title": p.name},
                                           "fields": "title"}}])
    main.append(f"'{PARTNERS}'!A:D", [[p.name, p.telegram, ", ".join(p.aliases),
                                       (p.added or today).isoformat()]])
    access = _grant(session, main, p.name, p.emails)   # доступ к основной таблице — по той же почте

    # убираем из песочницы: строку справочника и вкладку
    rows = sand.get_values(f"'{PARTNERS}'!A1:A400")
    for i, row in enumerate(rows, 1):
        if row and _key(row[0]) == _key(p.name):
            sand.clear(f"'{PARTNERS}'!A{i}:Z{i}")
    sand.batch([{"deleteSheet": {"sheetId": gid}}])

    for sid, sandbox in ((cfg.sheet_id, False), (cfg.sandbox_id, True)):
        sh = Sheet(session, sid)
        write_toc(sh, _directory(session, sa, sid), sh.tabs(), sandbox)
    return Added(p.name, f"https://docs.google.com/spreadsheets/d/{cfg.sheet_id}/edit#gid={new_gid}", access)


def due_for_promotion(directory: Directory, today: date) -> list[tuple[str, date]]:
    return [(p.name, p.added) for p in directory.partners
            if p.added and p.added <= today - timedelta(days=SANDBOX_DAYS)]
