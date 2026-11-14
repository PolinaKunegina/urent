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
from bot.workbook import (COL_NEXT_DATE, COL_NEXT_WHAT, PLAN_COL0, fx, parse_partner_tabs,
                          partner_tab_format_requests, partner_tab_values, sep_for_locale)

TODAY = date(2026, 11, 10)


def partner_sheet(wb, name, clients, days=None, start=date(2026, 11, 3)):
    ws = wb.create_sheet(name)
    grid = partner_tab_values(name, start, date(2026, 12, 31), clients)
    for row in grid:
        ws.append([date.fromisoformat(v) if isinstance(v, str) and len(v) == 10 and v[4] == "-" else
                   (None if isinstance(v, str) and v.startswith("=") else v) for v in row])
    from bot.workbook import P_FACT, P_PEOPLE, P_PLAN
    for d, plan, fact, *people in days or []:
        r = 5 + (d - start).days
        for off, vals in ((P_PLAN, plan), (P_FACT, fact), (P_PEOPLE, people[0] if people else ())):
            for i, v in enumerate(vals):
                if v is not None:
                    ws.cell(row=r, column=PLAN_COL0 + off + 1 + i, value=v)
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
        {"name": "Котова Вера", "source": "мероприятие", "next_date": d(2), "next_what": "созвон с Антоном",
         "touches": [{"n": 1, "date": d(-5), "fmt": "встреча"}]},
        {"name": "Лисина Ада", "next_date": d(12), "touches": [{"n": 1, "date": d(-5)}]},   # дальше 7 дней
        {"name": "Жуков Илья", "touches": [{"n": 1, "date": d(-9)}]},                       # нет шага
        {"name": "Ильин Олег", "status": "Оплатил пакет 240 / стал ментором", "touches": [{"n": 1, "date": d(-9)}]},
        {"name": "Гусев Семён", "status": "Отказ"},
        {"name": "Соколова Ия", "touches": [{"n": 1, "date": d(-3)}, {"n": 2, "date": d(1)}]},  # старый способ
        {"name": "Белых Ян", "next_date": d(-1), "next_what": "созвон"},                     # просрочен
        {"name": "Сегодня Тест", "next_date": d(0), "next_what": "пригласить на кинотренинг"},
    ], days=[(TODAY - timedelta(days=1), [2, 1, 0], [1.5, 1, 0], (3, 2, None)),
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
    kv = anna.clients[0]
    assert kv.source == "мероприятие" and kv.row == 5
    assert kv.next_date == TODAY + timedelta(days=2) and kv.next_what == "созвон с Антоном"
    assert [(t.n, t.when, t.fmt) for t in kv.touches] == [(1, TODAY - timedelta(days=5), "встреча")]
    y = anna.days[TODAY - timedelta(days=1)]
    assert y.plan == {"attract": 2, "warmup": 1, "close": 0} and y.fact == {"attract": 1.5, "warmup": 1, "close": 0}
    assert y.people == {"attract": 3, "warmup": 2}
    assert min(anna.days) == date(2026, 11, 3) and max(anna.days) == date(2026, 12, 31)


def test_compute_partner():
    s = stats()["Петрова Анна"]
    assert [l.name for l in s.active] == ["Котова Вера", "Лисина Ада", "Жуков Илья", "Соколова Ия", "Белых Ян",
                                          "Сегодня Тест"]
    assert [l.name for l in s.touch_today] == ["Сегодня Тест"]
    assert [l.name for l in s.closed] == ["Ильин Олег"]
    assert [l.name for l in s.no_touch] == ["Лисина Ада", "Жуков Илья", "Белых Ян"]
    assert [l.name for l in s.touch_tomorrow] == ["Соколова Ия"]
    assert s.fact_days == 2 and s.window == 7
    assert s.week == {"attract": 2.5, "warmup": 3, "close": 1}
    assert s.today_fact == {"attract": 1, "warmup": 0, "close": 0}
    assert s.tomorrow_plan == {"attract": 2, "warmup": 2, "close": 0}
    assert s.activity == round(100 * (2 / 7 + 3 / 6) / 2)


def test_activity_first_days_only_counts_since_start():
    s = compute_network(parse_partner_tabs(network_xlsx()), date(2026, 11, 4))
    orlov = next(x for x in s if x.name == "Орлов Пётр")
    assert orlov.window == 1 and orlov.fact_days == 0


def test_morning_report():
    text = render_morning(list(stats().values()), TODAY, "https://sheet")
    assert "Партнёров: <b>3</b>" in text
    assert "без касания в ближайшие 7 дней: <b>3</b>" in text
    assert "<b>Петрова Анна</b> (3): Лисина Ада, Жуков Илья, Белых Ян" in text
    assert "Ещё не внесли клиентов: Новикова Дарья" in text
    assert "<b>Петрова Анна</b>: привлечение 1,5 (2), 3 чел. · прогрев 1 (1), 2 чел. · закрытие 0 (0)" in text
    assert "Не заполнили факт: Новикова Дарья, Орлов Пётр" in text
    assert "https://sheet" in text


def test_weekly_and_summary():
    st = list(stats().values())
    w = render_weekly(st, TODAY)
    assert "Итоги недели 03.11–09.11" in w and "Петрова Анна" in w
    assert "на человека: привлечение 0,8 ч/чел · прогрев 1,5 ч/чел" in w
    rows = summary_rows(st)
    anna = next(r for r in rows if r[0] == "Петрова Анна")
    assert anna[2:6] == [6, 3, 50, 2] and anna[6:9] == [2.5, 3, 1] and anna[9:12] == [3, 2, 0]


def test_evening_texts():
    st = stats()
    chat = render_evening_chat(list(st.values()), TODAY)
    assert "Пока нет минут за сегодня: Новикова Дарья, Орлов Пётр" in chat
    p = render_partner_status(st["Петрова Анна"], TODAY, "https://tab")
    assert "Нет касания в ближайшие 7 дней (3)" in p and "Соколова Ия" in p
    assert "Факт за сегодня: привлечение 1 · прогрев 0 · закрытие 0 ч" in p
    assert "План на завтра: привлечение 2" in p and "https://tab" in p
    assert "Сегодня по плану:\n  • Сегодня Тест — пригласить на кинотренинг" in p
    assert needs_evening_ping(st["Петрова Анна"])
    assert "нет твоих клиентов" in render_partner_status(st["Новикова Дарья"], TODAY)


def test_layout_and_formulas():
    assert sep_for_locale("ru_RU") == ";" and sep_for_locale("en_US") == ","
    assert fx('=HYPERLINK("#gid=1","a,b")', ";") == '=HYPERLINK("#gid=1";"a,b")'
    grid = partner_tab_values("Анна", date(2026, 11, 3), date(2026, 12, 31), gid=77,
                              clients=[{"name": "Котова Вера", "next_date": "2026-11-06",
                                        "next_what": "созвон с Антоном"}])
    head = grid[3]
    assert head[1:9] == ["Имя клиента", "Откуда пришёл", "Следующий шаг: дата", "Следующий шаг: что",
                         "Время, мин", "Этап", "Текущий статус", "Что происходит"]
    assert head[9:12] == ["Дата 1-го касания", "Формат (этап)", "Мин"]
    assert head[PLAN_COL0] == "Дата" and PLAN_COL0 == 25 and head[PLAN_COL0 - 1] == ""
    assert head[PLAN_COL0 + 4] == "План: что буду делать"
    assert head[PLAN_COL0 + 8:PLAN_COL0 + 13] == ["Привлечение — людей", "Прогрев — людей", "Закрытие — людей",
                                                  "Факт: что сделал(а)", "% плана"]
    assert grid[4][PLAN_COL0 + 12] == '=IF(OR(SUM(AA5:AC5)=0;COUNT(AE5:AG5)=0);"";SUM(AE5:AG5)/SUM(AA5:AC5))'
    # факт: минуты касаний этого дня с нужным этапом; диапазоны сдвинуты на колонку и заходят за разделитель
    assert grid[4][PLAN_COL0 + 6] == ('=IFERROR(1/(1/ROUND(SUMIFS($L$5:$AA$104;$J$5:$Y$104;$Z5;'
                                      '$K$5:$Z$104;"*прогрев*")/60;2));"")')
    assert grid[5][PLAN_COL0 + 8] == '=IFERROR(1/(1/COUNTIFS($J$5:$Y$104;$Z6;$K$5:$Z$104;"*привлеч*"));"")'
    assert grid[4][COL_NEXT_DATE] == "2026-11-06" and grid[4][COL_NEXT_WHAT] == "созвон с Антоном"
    from bot.workbook import partner_links
    links = json.dumps(partner_links(77, PLAN_COL0, 1, 2), ensure_ascii=False)
    assert '"uri": "#gid=77&range=Z4"' in links and '"uri": "#gid=77&range=A4"' in links
    assert '"stringValue": "← Оглавление   📖 Инструкция"' in links and '"uri": "#gid=2"' in links
    assert "HYPERLINK" not in json.dumps(grid, ensure_ascii=False)
    assert grid[4][PLAN_COL0] == "2026-11-03" and grid[62][PLAN_COL0] == "2026-12-31"
    reqs = json.dumps(partner_tab_format_requests(42, service_email="bot@x"), ensure_ascii=False)
    assert 'OR($D5=\\"\\";$D5<TODAY();$D5>TODAY()+7)' in reqs and "LOWER($H5)" in reqs
    assert "=$Z5=TODAY()+1" in reqs and "DATE_IS_VALID" in reqs and "=$Z5>TODAY()" in reqs
    assert 'AND($F5<>\\"\\";$G5=\\"\\")' in reqs and '"Прогрев"' in reqs
    assert reqs.count("addProtectedRange") == 8   # шапка кусками вокруг целей + даты, №, %, факт-формулы
    assert "userEnteredFormat(" not in reqs.replace("userEnteredFormat(textFormat,", "")


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
    assert toc[6] == ["Петрова Анна", "Открыть →", "@anna_p"]
    posts = json.dumps([c[2] for c in fake.calls if c[0] == "POST"], ensure_ascii=False)
    assert '"uri": "#gid=2"' in posts
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


def test_month_results(tmp_path):
    from bot.network import month_line
    from bot.results import track
    from bot.sheets import Lead
    from bot.workbook import PartnerTab, result_cats
    assert result_cats("Стал партнёром") == set()
    assert result_cats("Оплатил 95") == {"c95"} and result_cats("Оплатил 240") == {"c240"}
    assert result_cats("Купил маленький продукт (не оргвзнос)") == {"c5000"} and result_cats("Оплатил интенсив 5000") == {"c5000"}
    assert result_cats("Думает") == set() and result_cats("Оплатил 95") == {"c95"}
    st = Store(str(tmp_path / "t.db"))
    tab = PartnerTab("Петрова Анна", clients=[Lead(5, "", "Котова Вера", status="Думает"),
                                              Lead(6, "", "Орлов Пётр", status="Оплатил 95")])
    assert track(st, "B", [tab], TODAY) == []                     # оплата без касаний в месяце — не результат
    tab.clients[0].status = "Оплатил 240"                         # Вера купила 240
    tab.clients.append(Lead(7, "", "Новый Ян", status="Оплатил интенсив 5000"))   # новый сразу с оплатой
    new = track(st, "B", [tab], TODAY)
    assert sorted((c, k) for _, c, k, _ in new) == [("Котова Вера", "c240"), ("Новый Ян", "c5000")]
    assert track(st, "B", [tab], TODAY) == []                     # повторно не засчитывает
    assert st.month_results("B", TODAY.strftime("%Y-%m")) == {("Петрова Анна", "c240"): 1,
                                                              ("Петрова Анна", "c5000"): 1}
    assert month_line({"leads": 3, "c240": 2}, {"c240": 1}) == "новых контактов 0/3 · партнёры 240 1/2"


def test_month_block_parsed():
    from openpyxl import load_workbook
    wb = load_workbook(io.BytesIO(network_xlsx()))
    ws = wb["Петрова Анна"]
    assert ws.cell(row=1, column=8).value.startswith("Мой месяц")
    for k, v in enumerate([3, 5, 2, 4]):
        ws.cell(row=2, column=9 + k, value=v)
    for k, v in enumerate([1, 2, 0, 1]):
        ws.cell(row=3, column=9 + k, value=v)
    buf = io.BytesIO()
    wb.save(buf)
    tab = next(t for t in parse_partner_tabs(buf.getvalue()) if t.name == "Петрова Анна")
    assert tab.goals == {"leads": 3, "c95": 5, "c240": 2, "c5000": 4}
    assert tab.results == {"leads": 1, "c95": 2, "c240": 0, "c5000": 1}
    s = next(x for x in compute_network([tab], TODAY))
    assert ("🎯 Месяц (факт/цель): новых контактов 1/3 · партнёры 95 2/5 · партнёры 240 0/2 · продукты 1/4"
            in render_partner_status(s, TODAY))
    from bot.workbook import new_leads_in_month
    assert new_leads_in_month(tab, TODAY) == 5     # клиенты Анны с первым касанием в ноябре



def test_backfill_counts_payment_with_touch_this_month(tmp_path):
    from bot.results import track
    from bot.sheets import Lead, Touch
    from bot.workbook import PartnerTab
    st = Store(str(tmp_path / "t.db"))
    tab = PartnerTab("Кунегина Полина", clients=[
        Lead(8, "", "Жуков Илья", status="Оплатил интенсив 5000",
             touches=[Touch(1, date(2026, 10, 29), "", "встреча"), Touch(3, date(2026, 11, 2), "", "оплата")]),
        Lead(9, "", "Старый Клиент", status="Оплатил 240", touches=[Touch(1, date(2026, 10, 1), "", "x")])])
    new = track(st, "B", [tab], date(2026, 11, 3))
    assert [(c, k) for _, c, k, _ in new] == [("Жуков Илья", "c5000")]
    assert track(st, "B", [tab], date(2026, 11, 3)) == []        # второй раз не дублирует


OLD_HEAD = ["№", "Имя клиента", "Откуда пришёл", "Следующий шаг: дата", "Следующий шаг: что", "Текущий статус",
            "Что происходит"] + [h for n in range(1, 6) for h in (f"Дата {n}-го касания", "Формат")] + [
            "", "Дата", "Привлечение — план, ч", "Прогрев — план, ч", "Закрытие — план, ч", "План: что буду делать",
            "Привлечение — факт, ч", "Прогрев — факт, ч", "Закрытие — факт, ч", "Привлечение — людей",
            "Прогрев — людей", "Закрытие — людей", "Факт: что сделал(а)", "% плана"]


def old_layout_tab():
    wb = Workbook()
    ws = wb.active
    ws.title = "Старая Вкладка"
    ws.cell(1, 6, "Мой месяц: ноябрь")
    for i, h in enumerate(OLD_HEAD):
        ws.cell(4, i + 1, h)
    ws.cell(5, 2, "Котова Вера")
    ws.cell(5, 8, date(2026, 10, 29))
    ws.cell(5, 9, "встреча")
    for k in range(10):
        ws.cell(5 + k, 19, date(2026, 11, 3) + timedelta(days=k))
    ws.cell(6, 24, 1.5)        # 04.11 — факт, внесённый руками
    ws.cell(11, 24, 2)         # 09.11 — тоже руками, но «сегодня»
    buf = io.BytesIO()
    wb.save(buf)
    return parse_partner_tabs(buf.getvalue())[0]


def test_upgrade_old_tab_to_time_and_stage():
    from bot.setup_sheet import needs_time_stage, time_stage_upgrade_requests
    tab = old_layout_tab()
    assert needs_time_stage(tab) and tab.plan_col == 18
    assert tab.clients[0].touches[0].fmt == "встреча" and tab.clients[0].touches[0].minutes is None
    reqs = time_stage_upgrade_requests(tab, 9, date(2026, 11, 9), "bot@x", 1, 2)
    ins = [r["insertDimension"]["range"] for r in reqs if "insertDimension" in r]
    assert [(r["startIndex"], r["endIndex"]) for r in ins] == [(5, 7), (19, 20), (17, 18), (15, 16), (13, 14), (11, 12)]
    cells = {(r["updateCells"]["start"]["rowIndex"], r["updateCells"]["start"]["columnIndex"]):
             [v.get("userEnteredValue", {}) for v in r["updateCells"]["rows"][0]["values"]]
             for r in reqs if "updateCells" in r}
    assert [v["stringValue"] for v in cells[(3, 5)]] == ["Время, мин", "Этап"]
    assert [v["stringValue"] for v in cells[(3, 10)]] == ["Формат (этап)", "Мин"]   # после 1-го касания
    assert [v["stringValue"] for v in cells[(3, 22)]] == ["Формат (этап)", "Мин"]   # после 5-го
    # блок часов переехал с S на Z; 04.11 с ручным фактом не трогаем, 09.11 (сегодня) — формулы
    assert (5, 30) not in cells and (10, 30) in cells and (4, 30) in cells
    assert cells[(10, 30)][1]["formulaValue"].startswith("=IFERROR(1/(1/ROUND(SUMIFS($L$5:$AA$104;$J$5:$Y$104;$Z11;")
    text = json.dumps(reqs, ensure_ascii=False)
    assert '"uri": "#gid=9&range=Z4"' in text and "bot:факт-формулы" in text
    # после перестройки та же вкладка читается как новая и второй раз не перестраивается
    from bot.workbook import CRM_HEADERS
    wb = Workbook()
    ws = wb.active
    ws.title = "Новая"
    for i, h in enumerate(CRM_HEADERS + ["", "Дата"]):
        ws.cell(4, i + 1, h)
    ws.cell(5, 2, "Котова Вера")
    ws.cell(5, 10, date(2026, 11, 9))
    ws.cell(5, 11, "созвон (прогрев)")
    ws.cell(5, 12, 30)
    buf = io.BytesIO()
    wb.save(buf)
    new = parse_partner_tabs(buf.getvalue())[0]
    assert not needs_time_stage(new) and new.time_col == 5
    assert new.touch_cols[0] == (1, 9, 10, 11)
    t = new.clients[0].touches[0]
    assert (t.when, t.fmt, t.minutes) == (date(2026, 11, 9), "созвон (прогрев)", 30)


def test_expand_history_triples():
    from bot.setup_sheet import expand_history

    class FakeSheet:
        def __init__(self):
            self.reqs = []

        def batch(self, reqs):
            self.reqs += reqs
    from bot.workbook import PartnerTab, TOUCH_COLS
    tab = PartnerTab(name="X", touch_cols=list(TOUCH_COLS), plan_col=25)
    sh = FakeSheet()
    expand_history(sh, tab, 5, add=2)
    ins = sh.reqs[0]["insertDimension"]["range"]
    assert (ins["startIndex"], ins["endIndex"]) == (24, 30)       # перед разделителем — формулы расширятся
    heads = [v["userEnteredValue"]["stringValue"] for v in sh.reqs[1]["updateCells"]["rows"][0]["values"]]
    assert heads == ["Дата 6-го касания", "Формат (этап)", "Мин", "Дата 7-го касания", "Формат (этап)", "Мин"]
    assert "range=AF4" in json.dumps(sh.reqs, ensure_ascii=False)


def test_month_block_repaired_after_column_insert():
    """Так выглядел блок после вставки колонки «Мин»: пустая колонка между 240 и продуктом."""
    from bot.setup_sheet import month_repair_requests
    from bot.workbook import CRM_HEADERS
    wb = Workbook()
    ws = wb.active
    ws.title = "Сдвинутый"
    for i, v in enumerate(["Мой месяц: ноябрь", "Новые контакты", "Партнёры 95", "Партнёры 240", "",
                           "Маленький продукт", "Выполнено"]):
        ws.cell(1, 8 + i, v)
    for i, v in enumerate([3, 5, 2, None, 4]):
        ws.cell(2, 9 + i, v)
    for i, h in enumerate(CRM_HEADERS + ["", "Дата"]):
        ws.cell(4, i + 1, h)
    buf = io.BytesIO()
    wb.save(buf)
    tab = parse_partner_tabs(buf.getvalue())[0]
    assert tab.goals == {"leads": 3, "c95": 5, "c240": 2, "c5000": 4} and not tab.month_ok
    prot = [{"protectedRangeId": 1, "description": "bot:заголовки", "editors": {"users": ["bot@x"]},
             "range": {"startRowIndex": 1, "endRowIndex": 2, "startColumnIndex": 0, "endColumnIndex": 8}},
            {"protectedRangeId": 2, "description": "bot:факт-формулы",
             "range": {"startRowIndex": 4, "endRowIndex": 104}}]
    reqs = month_repair_requests(tab, 9, TODAY, ";", prot, 38)
    rows = {r["updateCells"]["start"]["rowIndex"]: [v.get("userEnteredValue", {}) for v in r["updateCells"]["rows"][0]["values"]]
            for r in reqs if "updateCells" in r and "start" in r["updateCells"]}
    assert [v.get("stringValue") for v in rows[0]] == ["Мой месяц: ноябрь", "Новые контакты", "Партнёры 95",
                                                       "Партнёры 240", "Маленький продукт", "Выполнено"]
    assert [v.get("numberValue") for v in rows[1][1:5]] == [3, 5, 2, 4]
    assert rows[2][5]["formulaValue"] == '=IF(SUM(I2:L2)=0;"";SUM(I3:L3)/SUM(I2:L2))'
    clear = next(r["updateCells"]["range"] for r in reqs if "updateCells" in r and "range" in r["updateCells"])
    assert (clear["startColumnIndex"], clear["endColumnIndex"]) == (13, 14)        # старый «Выполнено» в N
    text = json.dumps(reqs)
    assert '"protectedRangeId": 1' in text and '"protectedRangeId": 2' not in text
    assert text.count("addProtectedRange") == 4
    # собранный блок читается как правильный
    for i, v in enumerate(["Мой месяц: ноябрь", "Новые контакты", "Партнёры 95", "Партнёры 240",
                           "Маленький продукт", "Выполнено", None]):
        ws.cell(1, 8 + i).value = v
    buf = io.BytesIO()
    wb.save(buf)
    assert parse_partner_tabs(buf.getvalue())[0].month_ok


def test_stage_matched_by_word_and_old_formulas_rewritten():
    from bot.setup_sheet import stage_criteria_updates
    from bot.workbook import PartnerTab, fact_formulas
    new = fact_formulas(5, 25, 9, ";")
    assert '"*прогрев*"' in new[1] and '"*привлеч*"' in new[3] and '"*закрыт*"' in new[5]
    old = [f.replace('"*привлеч*"', '"*(привлечение)*"').replace('"*прогрев*"', '"*(прогрев)*"')
           .replace('"*закрыт*"', '"*(закрытие)*"') for f in new]
    tab = PartnerTab(name="Юля", plan_col=25, time_col=5)
    data = stage_criteria_updates(tab, [old, [1.5, "", "", 3, "", ""], new])
    assert [d["range"] for d in data] == [f"'Юля'!{c}5" for c in ("AE", "AF", "AG", "AH", "AI", "AJ")]
    assert [d["values"][0][0] for d in data] == new          # ручные цифры и новые формулы не трогаем
