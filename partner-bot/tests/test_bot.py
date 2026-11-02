"""Тесты на структуре, повторяющей реальную таблицу «Воронка_партнерской_сети.xlsx».
Имена вымышленные: репозиторий публичный."""
import io
import json
from datetime import date, datetime

from openpyxl import Workbook

from bot.analytics import compute, render_morning, split_message
from bot.config import Config, parse_managers
from bot.gwrite import SheetsWriter, col_letter, helper_col_index
from bot.hashtags import Store, parse_message
from bot.partners import UNKNOWN, Directory, Partner
from bot.report import (evening_chat_text, evening_partner_texts, load_snapshot,
                        morning_report, weekly_report)
from bot.sheets import parse_funnel, to_date

TODAY = date(2026, 11, 3)


def make_xlsx(rows):
    wb = Workbook()
    ws = wb.active
    ws.title = "Воронка лидов"
    ws.append(["Воронка лидов Форма к"])
    ws.append(["Жёлтым — ячейки для заполнения."])
    ws.append(["№", "Имя лида", "Источник",
               "Дата 1-го касания", "Формат", "Дата 2-го касания", "Формат",
               "Дата 3-го касания", "Формат", "Дата 4-го касания", "Формат",
               "Дата 5-го касания", "Формат", "Текущий статус",
               "Описание, что происходит", "Описание, что происходит"])
    for r in rows:
        ws.append(r)
    wb.create_sheet("Дашборд").append(["Показатель", "Факт"])
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


ROWS = [
    # как в реальной таблице: дата строкой и настоящая дата вперемешку, источник «?»
    [1, "Котова Вера", "?", "29.10.2026", "Была на встрече", datetime(2026, 11, 3), "Zoom",
     None, None, None, None, None, None, "Касание 1", "Ждёт демо личного кабинета"],
    [3, "Орлов Пётр", "Анна Петрова", "29.10.2026", "что с ним?"],
    [16, "Лисина Ада", "Петрова Анна", "04.11.2026"],      # то же имя в другом порядке
    [23, "Лисина Ева", None, "19.11.2026"],
    [24, "Семён Гусев", "ментор", "29.10.2026", None, "вчера"],
    [25, "Жуков Илья", "Аня П.", "28.10.2026", None, "03.11.2026",
     None, None, None, None, None, None, None, "Отказ"],
    [26],  # пустые строки-заготовки
    [27, None],
]

DIRECTORY = Directory([
    Partner("Петрова Анна", aliases=["Аня П."], telegram="@anna_p"),
    Partner("Ментор", aliases=["ментор"]),
])


def funnel():
    return parse_funnel(make_xlsx(ROWS))


def test_parse_real_layout():
    f = funnel()
    assert [l.name for l in f.leads] == ["Котова Вера", "Орлов Пётр", "Лисина Ада",
                                         "Лисина Ева", "Семён Гусев", "Жуков Илья"]
    assert f.header_row == 3 and f.name_col == 1 and f.tab == "Воронка лидов"
    iv = f.leads[0]
    assert iv.row == 4 and iv.status == "Касание 1"
    assert iv.description == "Ждёт демо личного кабинета"
    assert [(t.n, t.when, t.fmt) for t in iv.touches] == [
        (1, date(2026, 10, 29), "Была на встрече"), (2, date(2026, 11, 3), "Zoom")]


def test_compute_by_partner():
    s = compute(funnel().leads, TODAY, "отказ", DIRECTORY)
    assert len(s.active) == 5 and [l.name for l in s.refused] == ["Жуков Илья"]
    assert list(s.partners) == ["Ментор", "Петрова Анна", UNKNOWN]
    anna = s.partners["Петрова Анна"]
    assert [l.name for l in anna.active] == ["Орлов Пётр", "Лисина Ада"]
    assert [l.name for l in anna.unplanned] == ["Орлов Пётр"]
    assert [(l.name, p) for l, _, p in s.today] == [("Котова Вера", UNKNOWN)]
    assert [(l.name, p) for l, _, p in s.tomorrow] == [("Лисина Ада", "Петрова Анна")]
    assert s.plan == {4: "сегодня", 5: "нет плана", 6: "04.11.2026", 7: "19.11.2026",
                      8: "нет плана", 9: "отказ"}

    text = render_morning(s)
    assert "В работе: <b>5</b> · сегодня: <b>1</b> · завтра: <b>1</b>" in text
    assert "🔴 без плана 1: Орлов Пётр" in text
    assert "Без партнёра — уточнить</b> (2): Котова Вера, Лисина Ева" in text
    assert "Не распознана дата: Семён Гусев" in text


def test_directory_without_file_merges_name_order():
    d = Directory([])
    assert d.resolve("Анна Петрова") == ("Анна Петрова", True)
    assert d.resolve("Петрова  анна")[0] == "Анна Петрова"
    assert d.resolve("?") == (UNKNOWN, False)
    assert d.resolve("") == (UNKNOWN, False)


def test_directory_file_and_unknown(tmp_path):
    path = tmp_path / "partners.json"
    path.write_text(json.dumps([{"name": "Петрова Анна", "aliases": ["Аня"], "telegram": "@Anna_P"}],
                               ensure_ascii=False), encoding="utf-8")
    d = Directory.load(str(path))
    assert d.resolve("аня") == ("Петрова Анна", True)
    assert d.resolve("Иван Новый") == ("Иван Новый", False)  # нет в справочнике — помечаем
    assert d.by_username("anna_p").name == "Петрова Анна"


def test_dates():
    assert to_date("29.10.2026") == date(2026, 10, 29)
    assert to_date("2026-10-02") == date(2026, 10, 2)
    assert to_date(46327) == date(2026, 11, 1)
    assert to_date("вчера") is None


def test_hashtag_message_multiline():
    text = ("#отчет #план_Петрова\n"
            "#знакомство 2 ч — презентация для Котовой\n"
            "#прогрев 1,5 часа — кейсы Соколовой\n"
            "#сделка 30 мин — Жуков, ждём оплату")
    r = parse_message(text)
    assert r.partner == "Петрова"
    assert {x.process: x.hours for x in r.entries} == {"intro": 2.0, "warmup": 1.5, "deal": 0.5}
    assert r.missing_hours == []


def test_hashtag_numbered_format_from_chat():
    # формат из переписки: «1. Процесс знакомства — часы и как выглядит»
    text = ("#знакомство\n"
            "1. Знакомство с экосистемой — 1:30, провела эфир\n"
            "2. Прогрев — 1 ч, переписка\n"
            "3. Закрытие — встреча с клиентом")
    r = parse_message(text)
    assert {x.process: x.hours for x in r.entries} == {"intro": 1.5, "warmup": 1.0, "deal": 0.0}
    assert r.missing_hours == ["deal"]


def test_not_a_report():
    assert parse_message("Всем привет! #новости") is None
    assert parse_message("") is None


def test_store(tmp_path):
    st = Store(str(tmp_path / "t.db"))
    st.save(1, 10, 5, "Петрова Анна", TODAY, parse_message("#прогрев 1 ч"))
    st.save(1, 10, 5, "Петрова Анна", TODAY, parse_message("#прогрев 2 ч"))  # отредактировали
    st.save(1, 11, 5, "Петрова Анна", TODAY, parse_message("#сделка 1 ч"))
    assert st.summary(TODAY, TODAY) == {"Петрова Анна": {"warmup": 2.0, "deal": 1.0}}
    st.upsert_user(5, "Anna_P", "Анна", private_ok=False)
    assert st.user_by_username("@anna_p") == (5, False)
    st.upsert_user(5, "Anna_P", "Анна", private_ok=True)
    st.upsert_user(5, "Anna_P", "Анна", private_ok=False)  # из группы не сбрасывает
    assert st.user_by_username("anna_p") == (5, True)
    assert st.all_entries()[0][5] == "anna_p"
    st.set_setting("partners_chat_id", "-100")
    assert st.get_setting("partners_chat_id") == "-100"


def test_reports_end_to_end(tmp_path):
    (tmp_path / "p.json").write_text(json.dumps(
        [{"name": "Петрова Анна", "aliases": ["Аня П."], "telegram": "@anna_p"},
         {"name": "Ментор", "aliases": ["ментор"]}], ensure_ascii=False), encoding="utf-8")
    cfg = Config(partners_file=str(tmp_path / "p.json"))
    st = Store(str(tmp_path / "t.db"))
    st.save(1, 1, 5, "Петрова Анна", date(2026, 11, 2), parse_message("#прогрев 1,5 ч"))
    snap = load_snapshot(cfg, TODAY, local_xlsx=make_xlsx(ROWS))
    assert snap.expected_partners() == ["Петрова Анна", "Ментор"]

    morning = morning_report(cfg, snap, st, TODAY)
    assert "Петрова Анна: знакомство 0 ч · прогрев 1,5 ч · сделки 0 ч" in morning
    assert "Не отчитались: Ментор" in morning

    weekly = weekly_report(snap, st, TODAY)
    assert "Итоги недели 27.10–02.11" in weekly
    assert "<b>Петрова Анна</b>: касаний 2, новых 2" in weekly  # Орлов 29.10, Жуков 28.10

    assert "ещё нет отчёта от: Петрова Анна, Ментор" in evening_chat_text(snap, st, TODAY)
    ev = evening_partner_texts(snap, TODAY)
    assert set(ev) == {"Петрова Анна", "Ментор"}
    assert "Орлов Пётр" in ev["Петрова Анна"] and "Лисина Ада" in ev["Петрова Анна"]


def test_broken_table_does_not_crash():
    snap = load_snapshot(Config(), TODAY, local_xlsx=b"not an xlsx")
    assert snap.stats is None
    assert "Не удалось прочитать таблицу" in morning_report(Config(), snap, None, TODAY)


class FakeResp:
    def __init__(self, data=None):
        self._data = data or {}
        self.status_code = 200

    def json(self):
        return self._data

    def raise_for_status(self):
        pass


class FakeSession:
    def __init__(self):
        self.calls = []
        self.rules = []

    def get(self, url, params=None, timeout=None):
        self.calls.append(("GET", url, params))
        return FakeResp({"sheets": [
            {"properties": {"sheetId": 7, "title": "Воронка лидов"}, "conditionalFormats": self.rules}]})

    def put(self, url, params=None, json=None, timeout=None):
        self.calls.append(("PUT", url, json))
        return FakeResp()

    def post(self, url, json=None, timeout=None):
        self.calls.append(("POST", url, json))
        return FakeResp()


def test_write_back_highlight():
    f = funnel()
    s = compute(f.leads, TODAY, "отказ", DIRECTORY)
    assert helper_col_index(f, "След. касание (бот)") == (16, False)
    assert col_letter(16) == "Q" and col_letter(26) == "AA"

    sess = FakeSession()
    SheetsWriter(sess, "ID").mark_plan(f, s.plan, "След. касание (бот)")
    puts = [c for c in sess.calls if c[0] == "PUT"]
    assert puts[0][2] == {"range": "'Воронка лидов'!Q3", "values": [["След. касание (бот)"]]}
    body = puts[1][2]
    assert body["range"].startswith("'Воронка лидов'!Q4:Q")
    assert body["values"][:6] == [["сегодня"], ["нет плана"], ["04.11.2026"], ["19.11.2026"],
                                  ["нет плана"], ["отказ"]]
    rule = [c for c in sess.calls if c[0] == "POST"][0][2]["requests"][0]["addConditionalFormatRule"]["rule"]
    assert rule["booleanRule"]["condition"]["values"][0]["userEnteredValue"] == '=$Q4="нет плана"'
    assert [r["startColumnIndex"] for r in rule["ranges"]] == [1, 16]

    # правило уже есть — второй раз не добавляем
    sess2 = FakeSession()
    sess2.rules = [{"booleanRule": {"condition": {"values": [{"userEnteredValue": '=$Q4="нет плана"'}]}}}]
    SheetsWriter(sess2, "ID").mark_plan(f, s.plan, "След. касание (бот)")
    assert not [c for c in sess2.calls if c[0] == "POST"]


def test_parse_managers():
    assert parse_managers("@Natasha, kir ;12345") == ["natasha", "kir", "12345"]
    assert parse_managers("@Anton, https://t.me/Polina_K") == ["anton", "polina_k"]


def test_split_message():
    text = "\n\n".join("x" * 1500 for _ in range(5))
    parts = split_message(text, 4000)
    assert all(len(p) <= 4000 for p in parts) and "".join(parts).count("x") == 7500


def test_partners_tab_in_workbook():
    from openpyxl import load_workbook
    from bot.sheets import PARTNERS_HEADER, parse_partners
    assert parse_partners(make_xlsx(ROWS)) is None  # листа нет
    wb = load_workbook(io.BytesIO(make_xlsx(ROWS)))
    ws = wb.create_sheet("Партнёры")
    ws.append(PARTNERS_HEADER)
    ws.append(["Петрова Анна", "@anna_p", "Аня П., Анна Петрова"])
    ws.append(["Ментор", "", "ментор"])
    buf = io.BytesIO()
    wb.save(buf)
    snap = load_snapshot(Config(partners_file="/nonexistent"), TODAY, local_xlsx=buf.getvalue())
    assert snap.partners_tab and snap.expected_partners() == ["Петрова Анна", "Ментор"]
    assert [l.name for l in snap.stats.partners["Петрова Анна"].active] == ["Орлов Пётр", "Лисина Ада"]
    assert snap.directory.by_username("anna_p").name == "Петрова Анна"


def test_write_back_creates_partners_tab(tmp_path):
    from bot import report
    snap = load_snapshot(Config(partners_file="/nonexistent"), TODAY, local_xlsx=make_xlsx(ROWS))
    snap.is_gsheet = True
    assert report.distinct_sources(snap.funnel) == ["Анна Петрова", "Аня П.", "ментор"]
    sess = FakeSession()
    st = Store(str(tmp_path / "t.db"))
    st.save(1, 1, 5, "Петрова Анна", TODAY, parse_message("#сделка 1 ч — договор"))
    cfg = Config(sheet_id="ID", partners_file="/nonexistent")
    assert report.write_back(cfg, snap, st, session=sess) is None
    puts = {c[2]["range"]: c[2]["values"] for c in sess.calls if c[0] == "PUT"}
    assert puts["'Партнёры'!A1"][1:] == [["Анна Петрова", "", ""], ["Аня П.", "", ""], ["ментор", "", ""]]
    assert puts["'Часы из чата'!A1"][1] == ["2026-11-03", "Петрова Анна", "Сделки и закрытие", 1.0,
                                           "договор", ""]
    snap.is_gsheet = False
    assert "xlsx" in report.write_back(cfg, snap, st, session=sess)
