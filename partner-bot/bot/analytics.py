"""Сводки: утренняя для руководителей, вечерняя для партнёров, недельная по понедельникам."""
from __future__ import annotations

import html
import re
from collections import Counter
from dataclasses import dataclass, field
from datetime import date, timedelta

from .hashtags import PROCESS_ORDER
from .partners import UNKNOWN, Directory
from .sheets import Lead, Touch

WEEKDAYS = ["пн", "вт", "ср", "чт", "пт", "сб", "вс"]
MONTHS = ["января", "февраля", "марта", "апреля", "мая", "июня", "июля",
          "августа", "сентября", "октября", "ноября", "декабря"]
SHORT_PROCESS = {"intro": "знакомство", "warmup": "прогрев", "deal": "сделки"}


def human_date(d: date) -> str:
    return f"{WEEKDAYS[d.weekday()]}, {d.day} {MONTHS[d.month - 1]}"


def fmt_hours(h: float) -> str:
    s = f"{h:.1f}".rstrip("0").rstrip(".")
    return s.replace(".", ",") + " ч"


def e(s: str) -> str:
    return html.escape(s, quote=False)


@dataclass
class PartnerStats:
    name: str
    known: bool
    active: list[Lead] = field(default_factory=list)
    today: list[tuple[Lead, Touch]] = field(default_factory=list)
    tomorrow: list[tuple[Lead, Touch]] = field(default_factory=list)
    unplanned: list[Lead] = field(default_factory=list)


@dataclass
class FunnelStats:
    today_date: date
    partners: dict[str, PartnerStats]
    refused: list[Lead]
    bad_dates: list[Lead]
    statuses: Counter
    plan: dict[int, str]          # строка таблицы -> значение служебной колонки
    lead_partner: dict[int, str]  # строка таблицы -> партнёр

    def _all(self, attr):
        return [x for p in self.partners.values() for x in getattr(p, attr)]

    @property
    def active(self):
        return self._all("active")

    @property
    def today(self):
        return [(l, t, self.lead_partner[l.row]) for l, t in self._all("today")]

    @property
    def tomorrow(self):
        return [(l, t, self.lead_partner[l.row]) for l, t in self._all("tomorrow")]

    @property
    def unplanned(self):
        return self._all("unplanned")


def compute(leads: list[Lead], today: date, exclude_status_regex: str,
            directory: Directory | None = None) -> FunnelStats:
    directory = directory or Directory([])
    rx = re.compile(exclude_status_regex, re.IGNORECASE) if exclude_status_regex else None
    tomorrow = today + timedelta(days=1)
    partners: dict[str, PartnerStats] = {}
    refused, plan, lead_partner = [], {}, {}
    for l in leads:
        pname, known = directory.resolve(l.source)
        lead_partner[l.row] = pname
        if rx and l.status and rx.search(l.status):
            refused.append(l)
            plan[l.row] = "отказ"
            continue
        ps = partners.setdefault(pname, PartnerStats(name=pname, known=known))
        ps.active.append(l)
        nt = l.next_touch(today)
        if nt is None:
            ps.unplanned.append(l)
            plan[l.row] = "нет плана"
            continue
        plan[l.row] = "сегодня" if nt.when == today else nt.when.strftime("%d.%m.%Y")
        if nt.when == today:
            ps.today.append((l, nt))
        elif nt.when == tomorrow:
            ps.tomorrow.append((l, nt))
    active = [l for p in partners.values() for l in p.active]
    return FunnelStats(
        today_date=today,
        partners=dict(sorted(partners.items(), key=lambda kv: (kv[0] == UNKNOWN, kv[0].lower()))),
        refused=refused,
        bad_dates=[l for l in active if l.bad_dates()],
        statuses=Counter(l.status or "без статуса" for l in active),
        plan=plan,
        lead_partner=lead_partner,
    )


# ---------- утренняя сводка ----------

def _touch_line(lead: Lead, touch: Touch, partner: str | None = None) -> str:
    bits = [f"касание {touch.n}"]
    if touch.fmt:
        bits.append(touch.fmt)
    who = f" <i>({e(partner)})</i>" if partner else ""
    line = f"  • <b>{e(lead.name)}</b>{who} — {e(', '.join(bits))}"
    if lead.description:
        d = lead.description if len(lead.description) <= 90 else lead.description[:87] + "…"
        line += f"\n    <i>{e(d)}</i>"
    return line


def render_morning(s: FunnelStats, hours_block: str | None = None,
                   errors: list[str] | None = None) -> str:
    today = s.today_date
    tomorrow = today + timedelta(days=1)
    out = [f"📊 <b>Воронка партнёрской сети — {human_date(today)}</b>"]
    out.append(
        f"В работе: <b>{len(s.active)}</b> · сегодня: <b>{len(s.today)}</b>"
        f" · завтра: <b>{len(s.tomorrow)}</b> · 🔴 без плана: <b>{len(s.unplanned)}</b>"
        + (f" · отказов: {len(s.refused)}" if s.refused else ""))

    out.append("")
    out.append(f"📅 <b>Касания сегодня, {today:%d.%m}</b> — {len(s.today)}")
    out += [_touch_line(l, t, p) for l, t, p in s.today] or ["  никого"]
    out.append(f"➡️ <b>Завтра, {tomorrow:%d.%m}</b> — {len(s.tomorrow)}")
    out += [_touch_line(l, t, p) for l, t, p in s.tomorrow] or ["  никого"]

    out.append("")
    out.append("👥 <b>По партнёрам</b>")
    for p in s.partners.values():
        if p.name == UNKNOWN:
            continue
        mark = "" if p.known else " <i>(нет в списке партнёров)</i>"
        line = (f"<b>{e(p.name)}</b>{mark} — в работе {len(p.active)} · сегодня {len(p.today)}"
                f" · завтра {len(p.tomorrow)}")
        if p.unplanned:
            line += f"\n  🔴 без плана {len(p.unplanned)}: " + ", ".join(e(l.name) for l in p.unplanned)
        out.append(line)
    unknown = s.partners.get(UNKNOWN)
    if unknown:
        out.append(f"❓ <b>{UNKNOWN}</b> ({len(unknown.active)}): "
                   + ", ".join(e(l.name) for l in unknown.active))

    if s.bad_dates:
        bad = "; ".join(f"{e(l.name)} (касание {t.n}: «{e(t.raw)}»)"
                        for l in s.bad_dates for t in l.bad_dates())
        out.append("")
        out.append(f"⚠️ Не распознана дата: {bad}")
    if s.statuses:
        out.append("")
        out.append("Статусы: " + ", ".join(f"{e(k)} — {v}" for k, v in s.statuses.most_common()))
    for err in errors or []:
        out.append("")
        out.append(f"❗️ {e(err)}")
    if hours_block:
        out.append("")
        out.append(hours_block)
    return "\n".join(out)


def render_hours(day_data: dict[str, dict[str, float]], week: dict[str, dict[str, float]],
                 expected_partners: list[str], day: date) -> str:
    out = [f"⏱ <b>Часы по процессам за {day:%d.%m}</b> (из чата)"]
    if day_data:
        for partner, procs in sorted(day_data.items()):
            parts = [f"{SHORT_PROCESS[p]} {fmt_hours(procs.get(p, 0))}" for p in PROCESS_ORDER]
            out.append(f"  • {e(partner)}: {' · '.join(parts)} = <b>{fmt_hours(sum(procs.values()))}</b>")
    else:
        out.append("  отчётов с хэштегами не было")
    silent = not_reported(expected_partners, day_data)
    if silent:
        out.append(f"  🔕 Не отчитались: {e(', '.join(silent))}")
    if week:
        totals = {p: sum(week[x].get(p, 0) for x in week) for p in PROCESS_ORDER}
        all_h = sum(totals.values()) or 1
        out.append("За 7 дней по всем: " + " · ".join(
            f"{SHORT_PROCESS[p]} {fmt_hours(totals[p])} ({round(100 * totals[p] / all_h)}%)"
            for p in PROCESS_ORDER))
    return "\n".join(out)


def not_reported(expected: list[str], data: dict[str, dict[str, float]]) -> list[str]:
    reported = {p.lower() for p in data}
    return [p for p in expected if p.lower() not in reported]


# ---------- вечер: партнёру его лиды ----------

def render_partner_evening(p: PartnerStats, today: date) -> str | None:
    if not p.unplanned and not p.tomorrow:
        return None
    tomorrow = today + timedelta(days=1)
    out = [f"🌿 <b>{e(p.name)}, сводка по вашим лидам</b>"]
    if p.unplanned:
        out.append(f"🔴 Нет даты следующего касания ({len(p.unplanned)}):")
        out += [f"  • {e(l.name)}" for l in p.unplanned]
        out.append("Поставьте дату в свободную колонку «Дата N-го касания» до 19:00.")
    if p.tomorrow:
        out.append(f"➡️ Касания завтра, {tomorrow:%d.%m}: " + ", ".join(e(l.name) for l, _ in p.tomorrow))
    return "\n".join(out)


# ---------- недельная сводка ----------

def render_weekly(leads: list[Lead], s: FunnelStats, hours: dict[str, dict[str, float]],
                  start: date, end: date) -> str:
    """Неделя start..end включительно."""
    touches, new_leads = Counter(), Counter()
    for l in leads:
        p = s.lead_partner.get(l.row, UNKNOWN)
        for t in l.touches:
            if t.when and start <= t.when <= end:
                touches[p] += 1
                if t.n == 1:
                    new_leads[p] += 1
    names = sorted(set(touches) | set(hours) | {k for k in s.partners if k != UNKNOWN},
                   key=lambda x: (x == UNKNOWN, x.lower()))
    out = [f"🗓 <b>Итоги недели {start:%d.%m}–{end:%d.%m}</b>",
           f"Касаний: <b>{sum(touches.values())}</b> · новых лидов: <b>{sum(new_leads.values())}</b>"
           f" · часов в чате: <b>{fmt_hours(sum(sum(v.values()) for v in hours.values()))}</b>"]
    for n in names:
        h = hours.get(n, {})
        procs = " · ".join(f"{SHORT_PROCESS[p]} {fmt_hours(h.get(p, 0))}" for p in PROCESS_ORDER)
        out.append(f"<b>{e(n)}</b>: касаний {touches[n]}, новых {new_leads[n]}; {procs}")
    if touches:
        leader, cnt = touches.most_common(1)[0]
        if leader != UNKNOWN:
            out.append(f"🏆 Больше всего касаний: {e(leader)} ({cnt})")
    if hours:
        top = max(hours.items(), key=lambda kv: sum(kv[1].values()))
        out.append(f"⏱ Больше всего часов: {e(top[0])} ({fmt_hours(sum(top[1].values()))})")
    return "\n".join(out)


def split_message(text: str, limit: int = 4000) -> list[str]:
    """Телеграм режет сообщения длиннее 4096 символов — делим по абзацам."""
    if len(text) <= limit:
        return [text]
    chunks, cur = [], ""
    for block in text.split("\n\n"):
        candidate = (cur + "\n\n" + block) if cur else block
        if len(candidate) <= limit:
            cur = candidate
            continue
        if cur:
            chunks.append(cur)
        while len(block) > limit:
            cut = block.rfind("\n", 0, limit)
            cut = cut if cut > 0 else limit
            chunks.append(block[:cut])
            block = block[cut:].lstrip("\n")
        cur = block
    if cur:
        chunks.append(cur)
    return chunks
