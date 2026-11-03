"""Тесты на структуре таблицы по ТЗ от 3 ноября. Имена вымышленные: репозиторий публичный."""
import io
import json
from datetime import date, datetime, timedelta

from openpyxl import Workbook

from bot.config import parse_managers
from bot.network import (compute_network, needs_evening_ping, render_evening_chat, render_morning,
                         render_partner_status, render_weekly, split_message, summary_rows)
from bot.partners import UNKNOWN, Directory
from bot.setup_sheet import ensure_structure
from bot.sheets import PARTNERS_HEADER, to_date
from bot.store import Store
from bot.workbook import (COL_NEXT, PLAN_COL0, fx, next_touch_formula, parse_partner_tabs,
                          partner_tab_format_requests, partner_tab_values, sep_for_locale)

TODAY = date(2026, 11, 10)


def partner_sheet(wb, name, clients, days=None, start=date(2026, 11, 3)):
    ws = wb.create_sheet(name)
    grid = partner_tab_values(name, start, date(2026, 12, 31), clients)
    for row in grid:
        ws.append([date.fromisoformat(v) if isinstance(v, str) and len(v) == 10 and v[4] == "-" else
                   (None if isinstance(v, str) and v.startswith("=") else v) for v in row])
    for d, plan, fact in days or []:
        r = 5 + (d - start).days
        for i, v in enumerate(plan + fact):
            if v is not None:
                ws.cell(row=r, column=PLAN_COL0 + 2 + i, value=v)
    return ws


def network_xlsx():
    wb = Workbook()
    wb.active.title = "Оглавление"
    p = wb.create_sheet("Партнёры")
    p.append(PARTNERS_HEADER)
    p.append(["Петрова Анна", "@anna_p", "Аня П."])
    p.append(["Орлов Пётр", "@petr_o", ""])
    p.append(["Новикова Дарья", "", ""])
    d = lambda x: (TODAY + timedelta(days=x)).isoformat()
    partner_sheet(wb, "Петрова Анна", [
        {"name": "Котова Вера", "source": "мероприятие", "touches": [{"n": 1, "date": d(-5)}, {"n": 2, "date": d(2)}]},
        {"name": "Лисина Ада", "touches": [{"n": 1, "date": d(-5)}, {"n": 2, "date": d(12)}]},  # дальше 7 дней
        {"name": "Жуков Илья", "touches": [{"n": 1, "date": d(-9)}]},                           # нет плана
        {"name": "Ильин Олег", "status": "Оплатил пакет 240 / стал ментором", "touches": [{"n": 1, "date": d(-9)}]},
        {"name": "Гусев Семён", "status": "Отказ"},
        {"name": "Соколова Ия", "touches": [{"n": 1, "date": d(1)}]},
    ], days=[(TODAY - timedelta(days=1), [2, 1, 0], [1.5, 1, 0]),
             (TODAY - timedelta(days=2), [1, 1, 1], [1, 2, 1]),
             (TODAY, [None, None, None], [1, 0, 0]),
             (TODAY + timedelta(days=1), [2, 2, 0], [None, None, None])])
    partner_sheet(wb, "Орлов Пётр", [{"name": "Белова Ника", "touches": [{"n": 1, "date": d(3)}]}])
    partner_sheet(wb, "Новикова Дарья", [])
    notes = wb.create_sheet("Лист2")
    notes.append(["имя клиента", "статус"])   # заметки, а не вкладка партнёра
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def stats():
    return {s.name: s for s in compute_network(parse_partner_tabs(network_xlsx()), TODAY)}


def test_parse_partner_tabs():
    tabs = {t.name: t for t in parse_partner_tabs(network_xlsx())}
    assert set(tabs) == {"Петрова Анна", "Орлов Пётр", "Новикова Дарья"}   # «Лист2» не вкладка
    anna = tabs["Петрова Анна"]
    assert [c.name for c in anna.clients][:2] == ["Котова Вера", "Лисина Ада"]
    assert anna.clients[0].source == "мероприятие" and anna.clients[0].row == 5
    assert [(t.n, t.when) for t in anna.clients[0].touches] == [(1, TODAY - timedelta(days=5)),
                                                                 (2, TODAY + timedelta(days=2))]
    y = anna.days[TODAY - timedelta(days=1)]
    assert y.plan == {"attract": 2, "warmup": 1, "close": 0} and y.fact == {"attract": 1.5, "warmup": 1, "close": 0}
    assert min(anna.days) == date(2026, 11, 3) and max(anna.days) == date(2026, 12, 31)


def test_compute_partner():
    s = stats()["Петрова Анна"]
    assert [l.name for l in s.active] == ["Котова Вера", "Лисина Ада", "Жуков Илья", "Соколова Ия"]
    assert [l.name for l in s.closed] == ["Ильин Олег"]
    assert [l.name for l in s.no_touch] == ["Лисина Ада", "Жуков Илья"]
    assert [l.name for l in s.touch_tomorrow] == ["Соколова Ия"]
    assert s.fact_days == 2 and s.window == 7
    assert s.week == {"attract": 2.5, "warmup": 3, "close": 1}
    assert s.today_fact == {"attract": 1, "warmup": 0, "close": 0}
    assert s.tomorrow_plan == {"attract": 2, "warmup": 2, "close": 0}
    assert s.activity == round(100 * (2 / 7 + 2 / 4) / 2)


def test_activity_first_days_only_counts_since_start():
    s = compute_network(parse_partner_tabs(network_xlsx()), date(2026, 11, 4))
    orlov = next(x for x in s if x.name == "Орлов Пётр")
    assert orlov.window == 1 and orlov.fact_days == 0


def test_morning_report():
    text = render_morning(list(stats().values()), TODAY, "https://sheet")
    assert "Партнёров: <b>3</b>" in text
    assert "без касания в ближайшие 7 дней: <b>2</b>" in text
    assert "<b>Петрова Анна</b> (2): Лисина Ада, Жуков Илья" in text
    assert "Ещё не внесли клиентов: Новикова Дарья" in text
    assert "<b>Петрова Анна</b>: привлечение 1,5 (2) · прогрев 1 (1) · закрытие 0 (0)" in text
    assert "Не заполнили факт: Новикова Дарья, Орлов Пётр" in text
    assert "https://sheet" in text


def test_weekly_and_summary():
    st = list(stats().values())
    w = render_weekly(st, TODAY)
    assert "Итоги недели 03.11–09.11" in w and "Петрова Анна" in w
    rows = summary_rows(st)
    anna = next(r for r in rows if r[0] == "Петрова Анна")
    assert anna[2:6] == [4, 2, 50, 2] and anna[6:9] == [2.5, 3, 1]


def test_evening_texts():
    st = stats()
    chat = render_evening_chat(list(st.values()), TODAY)
    assert "Факт за сегодня ещё не заполнили: Новикова Дарья, Орлов Пётр" in chat
    p = render_partner_status(st["Петрова Анна"], TODAY, "https://tab")
    assert "Нет касания в ближайшие 7 дней (2)" in p and "Соколова Ия" in p
    assert "Факт за сегодня: привлечение 1 · прогрев 0 · закрытие 0 ч" in p
    assert "План на завтра: привлечение 2" in p and "https://tab" in p
    assert needs_evening_ping(st["Петрова Анна"])
    assert "нет твоих клиентов" in render_partner_status(st["Новикова Дарья"], TODAY)


def test_formulas_follow_locale():
    assert sep_for_locale("ru_RU") == ";" and sep_for_locale("en_US") == ","
    assert next_touch_formula(5) == ('=IF(B5="";"";IF(MINIFS(D5:M5;D5:M5;">="&TODAY())=0;"нет плана";'
                                     'MINIFS(D5:M5;D5:M5;">="&TODAY())))')
    assert fx('=HYPERLINK("#gid=1","a,b")', ";") == '=HYPERLINK("#gid=1";"a,b")'
    grid = partner_tab_values("Анна", date(2026, 11, 3), date(2026, 12, 31))
    assert grid[3][1] == "Имя клиента" and grid[4][COL_NEXT].startswith("=IF(B5")
    assert grid[4][PLAN_COL0] == "2026-11-03" and grid[62][PLAN_COL0] == "2026-12-31"
    reqs = json.dumps(partner_tab_format_requests(42), ensure_ascii=False)
    assert 'OR($P5=\\"нет плана\\";$P5>TODAY()+7)' in reqs and "DATE_IS_VALID" in reqs


class Resp:
    def __init__(self, data=None):
        self._d, self.status_code = data or {}, 200

    def json(self):
        return self._d

    def raise_for_status(self):
        pass


class FakeSheets:
    def __init__(self, tabs):
        self.tabs, self.calls, self.next_id = dict(tabs), [], 100

    def get(self, url, params=None, timeout=None):
        if params and "locale" in params.get("fields", ""):
            return Resp({"properties": {"locale": "ru_RU"}})
        return Resp({"sheets": [{"properties": {"title": t, "sheetId": i}} for t, i in self.tabs.items()]})

    def post(self, url, json=None, timeout=None):
        self.calls.append(("POST", url, json))
        replies = []
        for r in (json or {}).get("requests", []):
            if "addSheet" in r:
                self.next_id += 1
                self.tabs[r["addSheet"]["properties"]["title"]] = self.next_id
                replies.append({"addSheet": {"properties": {"sheetId": self.next_id}}})
        return Resp({"replies": replies})

    def put(self, url, params=None, json=None, timeout=None):
        self.calls.append(("PUT", url, json))
        return Resp()


def test_ensure_structure_creates_missing_tabs_and_toc():
    from bot.partners import Partner
    fake = FakeSheets({"Партнёры": 1, "Петрова Анна": 2})
    d = Directory([Partner("Петрова Анна", telegram="@anna_p"), Partner("Орлов Пётр", telegram="@petr_o")])
    created = ensure_structure(fake, "ID", d, TODAY)
    assert created == ["Орлов Пётр"]
    puts = {c[2]["range"]: c[2]["values"] for c in fake.calls if c[0] == "PUT"}
    toc = puts["'Оглавление'!A1"]
    assert toc[6] == ["Петрова Анна", '=HYPERLINK("#gid=2";"Открыть →")', "@anna_p"]
    fresh = FakeSheets({"Лист1": 0})
    ensure_structure(fresh, "SB", Directory([]), TODAY, sandbox=True)
    assert "Партнёры" in fresh.tabs and "Оглавление" in fresh.tabs and "Инструкция" in fresh.tabs
    assert ensure_structure(fake, "ID", d, TODAY) == [] and "Инструкция" in fake.tabs
    assert puts["'Орлов Пётр'!A1"][3][1] == "Имя клиента"
    assert ensure_structure(fake, "ID", d, TODAY) == []   # второй раз ничего не создаёт


def test_directory():
    d = Directory([])
    assert d.resolve("Анна Петрова") == ("Анна Петрова", True)
    assert d.resolve("Петрова  анна")[0] == "Анна Петрова"
    assert d.resolve("?") == (UNKNOWN, False)


def test_misc():
    assert to_date("29.10.2026") == date(2026, 10, 29) and to_date(46327) == date(2026, 11, 1)
    assert to_date(datetime(2026, 11, 3, 0, 0)) == date(2026, 11, 3)
    assert parse_managers("@Anton, https://t.me/Polina_K") == ["anton", "polina_k"]
    parts = split_message("\n".join("x" * 100 for _ in range(100)), 4000)
    assert all(len(p) <= 4000 for p in parts)


def test_store(tmp_path):
    st = Store(str(tmp_path / "t.db"))
    st.upsert_user(5, "Anna_P", "Анна", private_ok=True)
    st.upsert_user(5, "Anna_P", "Анна", private_ok=False)
    assert st.user_by_username("@anna_p") == (5, True)
    st.set_setting("sheet_id", "x")
    assert st.get_setting("sheet_id") == "x"


def test_onboarding_helpers():
    import pytest
    from bot.onboarding import OnboardingError, clean_name, due_for_promotion
    from bot.partners import Partner
    assert clean_name("  иванова   мария ") == "Иванова Мария"
    assert clean_name("Петрова-Водкина Анна") == "Петрова-Водкина Анна"
    for bad in ["Мария", "Иванова Мария 1990", "привет как дела у тебя сегодня", "@ivanov"]:
        with pytest.raises(OnboardingError):
            clean_name(bad)
    d = Directory([Partner("Новикова Дарья", added=TODAY - timedelta(days=14)),
                   Partner("Белова Ника", added=TODAY - timedelta(days=3)), Partner("Без Даты")])
    assert due_for_promotion(d, TODAY) == [("Новикова Дарья", TODAY - timedelta(days=14))]


def test_guide_tab_content():
    from bot.guide import GUIDE, guide_values
    text = " ".join(t for _, t in GUIDE)
    assert "прообраз" in text and "Советник" in text and "красным" in text
    assert guide_values()[0][0].startswith("Ваша CRM")
