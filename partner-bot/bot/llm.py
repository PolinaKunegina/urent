"""Разбор отчётов партнёров (план/факт) через Claude.

Партнёр пишет как ему удобно, без хэштегов и строгого шаблона. Claude возвращает
JSON строго по схеме: часы по трём процессам и изменения по клиентам. Список лидов
партнёра из таблицы передаётся в запрос, чтобы Claude сопоставлял имена с теми,
что уже записаны, а не придумывал новые.
"""
from __future__ import annotations

import json
import logging
import os
from datetime import date

log = logging.getLogger(__name__)

# Haiku: дешёвая и быстрая, для разбора коротких отчётов её хватает.
# Можно сменить без правки кода: CLAUDE_MODEL=claude-sonnet-5-5 в .env
MODEL = os.environ.get("CLAUDE_MODEL", "claude-haiku-5-5")

_nullable_str = {"anyOf": [{"type": "string"}, {"type": "null"}]}

SCHEMA = {
    "type": "object",
    "properties": {
        "processes": {
            "type": "array",
            "items": {
                "type": "object",
                "properties": {
                    "process": {"type": "string", "enum": ["intro", "warmup", "deal"]},
                    "hours": {"type": "number"},
                    "people": {"anyOf": [{"type": "integer"}, {"type": "null"}]},
                    "note": {"type": "string"},
                },
                "required": ["process", "hours", "people", "note"],
                "additionalProperties": False,
            },
        },
        "clients": {
            "type": "array",
            "items": {
                "type": "object",
                "properties": {
                    "name_in_message": {"type": "string"},
                    "matched_lead": _nullable_str,
                    "status": _nullable_str,
                    "description": _nullable_str,
                    "touch_today": {"type": "boolean"},
                    "touch_format": _nullable_str,
                    "next_touch_date": _nullable_str,
                    "no_next_touch_reason": _nullable_str,
                },
                "required": ["name_in_message", "matched_lead", "status", "description",
                             "touch_today", "touch_format", "next_touch_date",
                             "no_next_touch_reason"],
                "additionalProperties": False,
            },
        },
        "questions": {"type": "array", "items": {"type": "string"}},
    },
    "required": ["processes", "clients", "questions"],
    "additionalProperties": False,
}

SYSTEM = """Ты разбираешь ежедневные отчёты партнёров партнёрской сети и превращаешь их в данные для таблицы воронки.

Три процесса работы партнёра:
- intro — знакомство с экосистемой (нетворкинг, встречи, приглашения на мероприятия);
- warmup — прогрев (переписка, кейсы, повторные встречи, созвоны для интереса);
- deal — сделки и закрытие (диагностики, индивидуальные созвоны на покупку, оплаты).

Из сообщения извлеки:
1. processes — часы по каждому упомянутому процессу (минуты переведи в часы, «1:30» = 1.5), сколько людей охвачено, если сказано (иначе null), и коротко, что делал. Процесс не упомянут — не включай его. Явно сказано «0 часов» — включи с hours 0.
2. clients — каждого упомянутого клиента (лида):
   - matched_lead — точное имя из списка лидов партнёра, если это тот же человек (имя и фамилия в любом порядке, уменьшительное имя, опечатка); если такого нет в списке — null;
   - status — текущий этап; если подходит один из существующих статусов, используй его дословно; если этап не назван — null;
   - description — что с человеком происходит, одной короткой фразой;
   - touch_today — было ли касание с ним в день отчёта;
   - touch_format — формат этого касания (встреча, созвон, переписка, мероприятие и т. п.) или null;
   - next_touch_date — дата следующего касания в формате YYYY-MM-DD, если названа (переводи «завтра», «в четверг», «через неделю» в дату относительно даты отчёта); иначе null;
   - no_next_touch_reason — если партнёр объяснил, почему следующего касания нет (отказ, оплатил, пауза), иначе null.
3. questions — что непонятно и стоит уточнить у партнёра (например, клиент без фамилии, неясная дата). Пусто, если всё ясно.

Не придумывай того, чего нет в сообщении. Пиши по-русски."""


def build_prompt(text: str, kind: str, partner: str, today: date,
                 leads: list[dict], statuses: list[str]) -> str:
    return (
        f"Дата отчёта: {today.isoformat()} ({['пн','вт','ср','чт','пт','сб','вс'][today.weekday()]})\n"
        f"Тип сообщения: {'ПЛАН на день' if kind == 'plan' else 'ФАКТ за день'}\n"
        f"Партнёр: {partner}\n\n"
        f"Лиды партнёра в таблице (JSON):\n{json.dumps(leads, ensure_ascii=False)}\n\n"
        f"Существующие статусы: {', '.join(statuses) if statuses else 'нет'}\n\n"
        f"Сообщение партнёра:\n<<<\n{text}\n>>>"
    )


class ReportParser:
    def __init__(self, api_key: str):
        import anthropic
        self.client = anthropic.Anthropic(api_key=api_key)

    def parse(self, text: str, kind: str, partner: str, today: date,
              leads: list[dict], statuses: list[str]) -> dict:
        response = self.client.messages.create(
            model=MODEL,
            max_tokens=16000,
            output_config={"effort": "medium",
                           "format": {"type": "json_schema", "schema": SCHEMA}},
            system=SYSTEM,
            messages=[{"role": "user",
                       "content": build_prompt(text, kind, partner, today, leads, statuses)}],
        )
        if response.stop_reason == "refusal":
            raise RuntimeError("Модель отказалась разбирать сообщение")
        if response.stop_reason == "max_tokens":
            raise RuntimeError("Ответ модели обрезан")
        raw = next(b.text for b in response.content if b.type == "text")
        return json.loads(raw)
