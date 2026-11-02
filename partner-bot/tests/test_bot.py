"""Тесты на структуре, повторяющей реальную таблицу «Воронка_партнерской_сети.xlsx»."""
import io
from datetime import date, datetime

from openpyxl import Workbook

from bot.analytics import compute, render_report, split_message
from bot.config import parse_sheets
from bot.hashtags import Store, parse_message
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
    # как в реальной таблице: дата строкой и настоящая дата вперемешку
    [1, "Котова Вера", "?", "29.10.2026", "Была на встрече", datetime(2026, 11, 3), "Zoom",
     None, None, None, None, None, None, "Касание 1", "Ждёт демо личного кабинета"],
    [3, "Орлов Пётр", "Партнёр А", "29.10.2026", "что с ним?"],
    [16, "Лисина Ада", None, "04.11.2026"],
    [23, "Лисина Ева", None, "19.11.2026"],
    [24, "Семён Гусев", "Партнёр Б", "29.10.2026", None, "вчера"],
    [25, "Жуков Илья", "Партнёр В", "29.10.2026", None, "03.11.2026",
     None, None, None, None, None, None, None, "Отказ"],
    [26],  # пустые строки-заготовки
    [27, None],
]


def test_parse_real_layout():
    leads = parse_funnel(make_xlsx(ROWS))
    assert [l.name for l in leads] == ["Котова Вера", "Орлов Пётр", "Лисина Ада",
                                       "Лисина Ева", "Семён Гусев", "Жуков Илья"]
    iv = leads[0]
    assert iv.status == "Касание 1"
    assert iv.description == "Ждёт демо личного кабинета"
    assert [(t.n, t.when, t.fmt) for t in iv.touches] == [
        (1, date(2026, 10, 29), "Была на встрече"), (2, date(2026, 11, 3), "Zoom")]


def test_compute_report():
    leads = parse_funnel(make_xlsx(ROWS))
    s = compute("Сергиенко", leads, TODAY, "отказ")
    assert len(s.active) == 5 and [l.name for l in s.refused] == ["Жуков Илья"]
    assert [l.name for l, _ in s.today] == ["Котова Вера"]
    assert [l.name for l, _ in s.tomorrow] == ["Лисина Ада"]
    assert [l.name for l in s.unplanned] == ["Орлов Пётр", "Семён Гусев"]
    assert [l.name for l in s.bad_dates] == ["Семён Гусев"]

    text = render_report([s], TODAY, None, [])
    assert "Лидов в работе: <b>5</b>" in text
    assert "Касания сегодня, 03.11: <b>1</b>" in text
    assert "Без запланированного касания: <b>2</b>" in text


def test_dates():
    assert to_date("29.10.2026") == date(2026, 10, 29)
    assert to_date("2026-10-02") == date(2026, 10, 2)
    assert to_date(46327) == date(2026, 11, 1)
    assert to_date("вчера") is None


def test_hashtag_message_multiline():
    text = ("#отчет #план_Сергиенко\n"
            "#знакомство 2 ч — презентация для Котовой\n"
            "#прогрев 1,5 часа — кейсы Соколовой\n"
            "#сделка 30 мин — Жуков, ждём оплату")
    r = parse_message(text)
    assert r.partner == "Сергиенко"
    assert {e.process: e.hours for e in r.entries} == {"intro": 2.0, "warmup": 1.5, "deal": 0.5}
    assert r.missing_hours == []


def test_hashtag_numbered_format_from_chat():
    # формат из переписки: «1. Процесс знакомства — часы и как выглядит»
    text = ("#знакомство\n"
            "1. Знакомство с экосистемой — 1:30, провела эфир\n"
            "2. Прогрев — 1 ч, переписка\n"
            "3. Закрытие — встреча с клиентом")
    r = parse_message(text)
    hours = {e.process: e.hours for e in r.entries}
    assert hours == {"intro": 1.5, "warmup": 1.0, "deal": 0.0}
    assert r.missing_hours == ["deal"]


def test_not_a_report():
    assert parse_message("Всем привет! #новости") is None
    assert parse_message("") is None


def test_store_edit_replaces(tmp_path):
    st = Store(str(tmp_path / "t.db"))
    st.save(1, 10, 5, "Сергиенко", TODAY, parse_message("#прогрев 1 ч"))
    st.save(1, 10, 5, "Сергиенко", TODAY, parse_message("#прогрев 2 ч"))  # отредактировали
    st.save(1, 11, 5, "Сергиенко", TODAY, parse_message("#сделка 1 ч"))
    assert st.summary(TODAY, TODAY) == {"Сергиенко": {"warmup": 2.0, "deal": 1.0}}
    st.remember_partner(5, "Сергиенко")
    assert st.partner_for(5) == "Сергиенко"


def test_config_sheets():
    s = parse_sheets("Сергиенко=abc; Иванова = def ;xyz")
    assert [(x.label, x.file_id) for x in s] == [
        ("Сергиенко", "abc"), ("Иванова", "def"), ("Таблица 3", "xyz")]


def test_split_message():
    text = "\n\n".join("x" * 1500 for _ in range(5))
    parts = split_message(text, 4000)
    assert all(len(p) <= 4000 for p in parts) and "".join(parts).count("x") == 7500
