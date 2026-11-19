"""Мероприятия: журнал приглашений, сводка, сверка. Имена вымышленные."""
import io
import json
from datetime import date

from openpyxl import Workbook

from bot.events import (Event, build_report, event_key, lists_values, norm_contact, parse_event_rows,
                        parse_sovetnik_export, parse_tracker, update_journal, evening_lines, Tracker)
from bot.events_sync import event_cols_upgrade_requests, guest_requests, needs_event_cols
from bot.workbook import CRM_HEADERS, parse_partner_tabs

TODAY = date(2026, 11, 20)
EV = [Event(date(2026, 11, 19), "Бизнес-встреча", "Бизнес-встреча", "Офлайн", "Кир Воликов", 5),
      Event(date(2026, 11, 26), "Бизнес-завтрак", "Бизнес-завтрак", "Офлайн", "Наталья Сергиенко", 4)]
K1, K2 = EV[0].key, EV[1].key


def tab_with_events(name="Петрова Анна", clients=()):
    """Вкладка нового макета + колонки мероприятий после «Что происходит»."""
    head = list(CRM_HEADERS)
    at = head.index("Что происходит") + 1
    head[at:at] = ["Мероприятие", "Был", "В Советнике"]
    wb = Workbook()
    ws = wb.active
    ws.title = name
    ws.cell(1, 9, "Мой месяц: ноябрь")
    for i, h in enumerate(head + ["", "Дата"]):
        ws.cell(4, i + 1, h)
    for r, c in enumerate(clients, start=5):
        ws.cell(r, 2, c["name"])
        ws.cell(r, 3, c.get("contact", ""))
        ws.cell(r, 9, c.get("status", "Новый контакт"))
        ws.cell(r, at + 1, c.get("event", ""))
        ws.cell(r, at + 2, c.get("came", False))
        ws.cell(r, at + 3, c.get("reg", False))
    buf = io.BytesIO()
    wb.save(buf)
    return parse_partner_tabs(buf.getvalue())[0]


def test_event_key_and_lists():
    assert event_key(date(2026, 11, 19), " Бизнес-встреча ") == "19.11 · Бизнес-встреча"
    rows = lists_values(";", ["Петрова Анна", "Орлов Пётр"])
    assert rows[0][0].startswith("Мероприятия для выбора")
    assert "TODAY()-WEEKDAY(TODAY();3)" in rows[1][0] and "INDEX(SORT(FILTER(" in rows[1][0]
    assert rows[1][5] == "Петрова Анна" and rows[2][5] == "Орлов Пётр"
    evs = parse_event_rows([["19.11.2026", "Бизнес-встреча", "Бизнес-встреча", "Офлайн", "Кир", "5"],
                            ["", "пусто"], ["26.11.2026", "Завтрак"]])
    assert [(e.key, e.plan) for e in evs] == [("19.11 · Бизнес-встреча", 5.0), ("26.11 · Завтрак", None)]


def test_partner_tab_reads_event_columns():
    t = tab_with_events(clients=[{"name": "Котова Вера", "contact": "@kotova", "event": K1, "came": True}])
    c = t.clients[0]
    assert (c.event, c.attended, c.registered, c.contact) == (K1, True, False, "@kotova")
    assert t.event_col is not None and not needs_event_cols(t)


def test_journal_keeps_old_invitation_and_report():
    anna = tab_with_events(clients=[
        {"name": "Котова Вера", "contact": "+7 (916) 123-45-67", "event": K1, "came": True, "reg": True,
         "status": "Был на встрече"},
        {"name": "Орлов Пётр", "contact": "@orlov_p", "event": K1, "came": True, "status": "Был на встрече"},
        {"name": "Лисина Ада", "event": K1, "status": "Новый контакт"}])
    ev = {e.key: e for e in EV}
    j = update_journal([], [("B", [anna])], ev, TODAY)
    assert len(j) == 3 and all(r["status_at_event"] for r in j if r["client"] != "Лисина Ада")
    # Котова прошла дальше: статус вырос; Орлову выбрали новое мероприятие
    anna = tab_with_events(clients=[
        {"name": "Котова Вера", "contact": "+7 (916) 123-45-67", "event": K1, "came": True, "reg": True,
         "status": "Оплатил 95"},
        {"name": "Орлов Пётр", "contact": "@orlov_p", "event": K2, "status": "Назначена диагностика"}])
    j = update_journal(j, [("B", [anna])], ev, TODAY)
    keys = {(r["client"], r["event"]): r for r in j}
    assert keys[("Орлов Пётр", K1)]["attended"] and not keys[("Орлов Пётр", K1)]["current"]   # старое осталось
    assert keys[("Орлов Пётр", K2)]["current"]
    assert not keys[("Лисина Ада", K1)]["_on_tab"]                  # строку удалили с вкладки
    tr = Tracker(by_day={EV[0].day: [("Котова Вера", ""), ("Орлов Пётр", ""), ("Гость", "Петрова Анна")]})
    summary, people = build_report(j, [("B", [anna])], EV, TODAY, tr, {"9161234567"}, {"Петрова Анна": "Петрова Анна"})
    s1 = next(s for s in summary if s[1] == "Бизнес-встреча")
    # план, приглашено, пришло, трекер, был+Сов, был без Сов, дальше по воронке, оплата, % плана
    assert s1[4:13] == [5, 3, 2, 3, 1, 1, 2, 1, "40%"]
    assert "⚠ на вкладках 2, в журнале 3" in s1[13] and "⚠ пришло по галочкам 2, по трекеру 3" in s1[13]
    assert s1[14] == "Петрова Анна 3/2"
    s2 = next(s for s in summary if s[1] == "Бизнес-завтрак")
    assert s2[5] == 1 and s2[12] == ""                               # будущее: % плана не считаем
    outcome = {(p[2], p[1]): p[9] for p in people}
    assert outcome[("Котова Вера", "Бизнес-встреча")] == "был и в Советнике"
    assert outcome[("Орлов Пётр", "Бизнес-встреча")] == "был, не в Советнике"
    assert outcome[("Орлов Пётр", "Бизнес-завтрак")] == "ещё впереди"
    check = {p[2]: p[11] for p in people if p[1] == "Бизнес-встреча"}
    assert check["Котова Вера"] == "" and check["Лисина Ада"].startswith("⚠ клиента нет на вкладке")


def test_registration_check_against_export():
    t = tab_with_events(clients=[{"name": "Котова Вера", "contact": "@Kotova", "event": K1, "came": True},
                                 {"name": "Орлов Пётр", "contact": "8 916 000-11-22", "event": K1, "came": True,
                                  "reg": True}])
    j = update_journal([], [("B", [t])], {e.key: e for e in EV}, TODAY)
    reg = parse_sovetnik_export([["Имя", "Телефон", "Telegram"], ["Вера", "", "kotova"], ["Иван", "+79160000000", ""]])
    assert "@kotova" in reg and "9160000000" in reg
    _, people = build_report(j, [("B", [t])], EV, TODAY, None, reg)
    check = {p[2]: p[11] for p in people}
    assert check["Котова Вера"] == "⚠ в выгрузке есть, галочки нет"
    assert check["Орлов Пётр"] == "⚠ галочка есть, в выгрузке нет"
    assert norm_contact("t.me/Some_Nick, +7 999 111-22-33") == {"@some_nick", "9991112233"}


def test_tracker_counts_newcomers_by_day():
    rows = [["Дата", "Мероприятие", "Участник", "", "", "", "", "", "", "Формат"],
            ["19.11", "Бизнес-встреча, лофт", "Котова Вера", "", "", "", "", "", "", "офлайн"],
            ["19.11", "Бизнес-встреча, лофт", "Петрова Анна", "", "", "", "", "", "", "офлайн"],      # партнёр
            ["19.11", "Бизнес-встреча, лофт", "Олеся (гость Петровой Анны)", "", "", "", "", "", "", "онлайн"],
            ["19.11", "Бизнес-встреча, лофт", "Орлов Пётр", "", "", "", "", "", "", "не был"]]
    tr = parse_tracker(rows, ["Петрова Анна"], TODAY)
    assert tr.by_day[date(2026, 11, 19)] == [("Котова Вера", ""), ("Олеся", "Петровой Анны")]


def test_event_columns_inserted_and_month_block_reassembled():
    head = list(CRM_HEADERS) + ["", "Дата"]
    wb = Workbook()
    ws = wb.active
    ws.title = "Петрова Анна"
    for i, v in enumerate(["Мой месяц: ноябрь", "Новые контакты", "Партнёры 95", "Партнёры 240",
                           "Маленький продукт", "Выполнено"]):
        ws.cell(1, 9 + i, v)
    ws.cell(2, 10, 3)
    for i, h in enumerate(head):
        ws.cell(4, i + 1, h)
    buf = io.BytesIO()
    wb.save(buf)
    t = parse_partner_tabs(buf.getvalue())[0]
    assert needs_event_cols(t) and t.desc_col == 9 and t.plan_col == 26 and t.month_end == 14
    reqs = event_cols_upgrade_requests(t, 7, TODAY, ";", [], 1, 2)
    ins = reqs[0]["insertDimension"]["range"]
    assert (ins["startIndex"], ins["endIndex"]) == (10, 13)          # после «Что происходит» (J)
    text = json.dumps(reqs, ensure_ascii=False)
    assert '"stringValue": "Мероприятие"' in text and '"type": "BOOLEAN"' in text
    assert "='_списки'!A2:A300" in text and '"uri": "#gid=7&range=AD4"' in text
    clear = [r["updateCells"]["range"] for r in reqs if "updateCells" in r and "range" in r["updateCells"]]
    assert (clear[0]["startColumnIndex"], clear[0]["endColumnIndex"]) == (14, 17)   # хвост сдвинутого блока


def test_guest_goes_to_partner_tab():
    t = tab_with_events(clients=[{"name": "Котова Вера"}])
    guests = [["Гости"], [], [], ["Мероприятие", "Имя", "Контакт", "Кто привёл", "Был", "Готово"],
              [K1, "Олеся Новая", "@olesya", "Петрова Анна", "TRUE", ""],
              [K1, "Котова Вера", "", "Петрова Анна", "TRUE", ""],
              [K1, "Без партнёра", "", "", "TRUE", ""],
              [K1, "Уже готово", "", "Петрова Анна", "", "✓ добавлен"]]
    data, marks = guest_requests(guests, {"Петрова Анна": t}, {"Петрова Анна": 5})
    cells = {d["range"]: d["values"][0][0] for d in data}
    assert cells["'Петрова Анна'!B6"] == "Олеся Новая" and cells["'Петрова Анна'!C6"] == "@olesya"
    assert cells["'Петрова Анна'!K6"] == K1 and cells["'Петрова Анна'!L6"] is True
    assert marks == [(5, "✓ добавлен на вкладку «Петрова Анна», строка 6"), (6, "✓ уже есть у Петрова Анна")]


def test_evening_lines_for_event_day():
    t = tab_with_events(clients=[{"name": "Котова Вера", "event": K1, "came": True}, {"name": "Орлов Пётр", "event": K1}])
    assert evening_lines(t, {e.key: e for e in EV}, date(2026, 11, 19)) == [
        "📍 Сегодня «Бизнес-встреча»: отметь галочкой «Был», кто пришёл — Котова Вера ✓, Орлов Пётр"]
    assert evening_lines(t, {e.key: e for e in EV}, TODAY) == []


def test_tracker_skips_partner_written_by_alias():
    rows = [["Дата"], ["12.11", "М", "Лина Герц (Алина Гайнаншина)", "", "", "", "", "", "", "офлайн"],
            ["12.11", "М", "Руслан Татлыев", "", "", "", "", "", "", "офлайн"]]
    tr = parse_tracker(rows, ["Гайнаншина Алина", "Алина Герц, Лина Герц"], TODAY)
    assert tr.by_day[date(2026, 11, 12)] == [("Руслан Татлыев", "")]


def test_event_column_widths_survive_month_repair():
    from bot.workbook import PartnerTab
    t = PartnerTab(name="X", desc_col=9, plan_col=26, month_col=8, month_end=14)
    reqs = event_cols_upgrade_requests(t, 7, TODAY, ";", [], None, None)
    widths = {}
    for r in reqs:
        if "updateDimensionProperties" in r:
            widths[r["updateDimensionProperties"]["range"]["startIndex"]] = r["updateDimensionProperties"]["properties"]["pixelSize"]
    assert (widths[10], widths[11], widths[12]) == (190, 70, 95)


def test_paid_before_event_not_counted_and_untracked_day_is_soft():
    t = tab_with_events(clients=[{"name": "Коваленко Олеся", "event": K1, "came": True,
                                  "status": "Купил маленький продукт (не оргвзнос)"}])
    j = update_journal([], [("B", [t])], {e.key: e for e in EV}, TODAY)
    summary, _ = build_report(j, [("B", [t])], EV, TODAY, Tracker(by_day={}), None)
    s1 = next(s for s in summary if s[1] == "Бизнес-встреча")
    assert s1[11] == 0 and s1[7] == "" and s1[13] == "в трекере нет записей за эту дату"
    s2 = next(s for s in summary if s[1] == "Бизнес-завтрак")
    assert s2[7] == "" and s2[13] == "✓"


def test_event_row_status_not_red_and_not_a_client():
    from bot.network import compute_partner
    from bot.setup_sheet import event_status_requests
    from bot.workbook import new_leads_in_month
    t = tab_with_events(clients=[{"name": "Нетворкинг 18.11", "status": "Событие (не клиент)"},
                                 {"name": "Котова Вера", "status": "Новый контакт"}])
    ps = compute_partner(t, TODAY)
    assert [l.name for l in ps.active] == ["Котова Вера"] and not ps.closed
    assert new_leads_in_month(t, TODAY) == 0
    reqs = json.dumps(event_status_requests(t, 9, ";"), ensure_ascii=False)
    assert '"Событие (не клиент)"' in reqs and 'REGEXMATCH(LOWER($I5);\\"событие\\")' in reqs
    assert '"index": 0' in reqs
