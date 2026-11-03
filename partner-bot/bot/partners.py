"""Справочник партнёров: кто есть кто в колонке «Источник» и в Telegram.

Основной вариант: лист «Партнёры» в той же Google Таблице (бот создаёт его сам и
заполняет вариантами из колонки «Источник», людям остаётся дописать @username
и объединить варианты одного человека). Запасной вариант, если листа нет:
файл data/partners.json (в git не попадает, там живые имена и ники):

    [
      {"name": "Петрова Анна", "aliases": ["Анна Петрова", "Аня"], "telegram": "@anna_p"},
      {"name": "Ментор", "aliases": ["ментор"], "telegram": ""}
    ]

«Имя Фамилия» и «Фамилия Имя» бот и так считает одним человеком, aliases нужны для
сокращений и опечаток. Пустой источник и «?» уходят в блок «Без партнёра — уточнить».
"""
from __future__ import annotations

import json
from datetime import date
import re
from dataclasses import dataclass, field
from pathlib import Path

UNKNOWN = "Без партнёра — уточнить"


def _key(s: str) -> str:
    """Ключ сравнения: без регистра, ё=е, без @ и знаков, слова по алфавиту."""
    s = s.lower().replace("ё", "е")
    words = re.findall(r"[\w]+", s)
    return " ".join(sorted(words))


@dataclass
class Partner:
    name: str
    aliases: list[str] = field(default_factory=list)
    telegram: str = ""        # @username без @, в нижнем регистре
    added: "date | None" = None   # когда добавлен (для песочницы: перевод через 2 недели)

    @property
    def username(self) -> str:
        return self.telegram.strip().lstrip("@").lower()


class Directory:
    def __init__(self, partners: list[Partner]):
        self.partners = partners
        self._by_key: dict[str, Partner] = {}
        self._seen: dict[str, str] = {}   # для имён вне справочника: первое написание
        for p in partners:
            for variant in [p.name, *p.aliases]:
                if variant.strip():
                    self._by_key[_key(variant)] = p

    @classmethod
    def from_records(cls, raw: list[dict]) -> "Directory":
        return cls([Partner(name=x["name"], aliases=x.get("aliases", []),
                            telegram=x.get("telegram", ""), added=x.get("added")) for x in raw])

    @classmethod
    def load(cls, path: str) -> "Directory":
        f = Path(path)
        if not f.exists():
            return cls([])
        return cls.from_records(json.loads(f.read_text(encoding="utf-8")))

    def resolve(self, source: str) -> tuple[str, bool]:
        """(имя партнёра, известен ли он справочнику)."""
        src = (source or "").strip()
        if not src or not re.search(r"\w", src):  # пусто или «?»
            return UNKNOWN, False
        k = _key(src)
        p = self._by_key.get(k)
        if p:
            return p.name, True
        name = self._seen.setdefault(k, src)  # «Имя Фамилия» = «Фамилия Имя»
        return name, not self.partners  # без справочника любое имя считаем партнёром

    def by_username(self, username: str) -> Partner | None:
        u = (username or "").lstrip("@").lower()
        return next((p for p in self.partners if p.username and p.username == u), None)

    def names(self) -> list[str]:
        return [p.name for p in self.partners]
