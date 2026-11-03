"""Телеграм-бот партнёрской сети (ТЗ от 3 ноября).

Партнёры ведут клиентов и часы в таблице, бот:
  08:00  руководителям в личку — клиенты без касания в ближайшие 7 дней, часы за вчера
         (факт и план), активность партнёров; по понедельникам — итоги недели;
  18:30  в чат партнёров — напоминание заполнить таблицу и кто ещё не заполнил;
         каждому партнёру в личку — его клиенты без касания, заполнен ли факт и план;
  каждый час — вкладки новым партнёрам (по листу «Партнёры») и лист «Сводка».
"""
from __future__ import annotations

import asyncio
import logging
import re
from datetime import datetime

from telegram import Update
from telegram.constants import ParseMode
from telegram.ext import Application, CommandHandler, ContextTypes, TypeHandler

from .config import Config, load_config
from .network import (needs_evening_ping, render_evening_chat, render_partner_status,
                      split_message)
from .report import load_network, morning_text, sync, weekly_text
from .store import Store

logging.basicConfig(format="%(asctime)s %(levelname)s %(name)s: %(message)s", level=logging.INFO)
logging.getLogger("httpx").setLevel(logging.WARNING)
log = logging.getLogger("partner-bot")

GUIDE = (
    "<b>Как вести таблицу</b> 🌿\n\n"
    "У тебя своя вкладка с твоим именем (ссылка — в «Оглавлении» таблицы или по /mine).\n\n"
    "<b>Слева — твои клиенты.</b> Имя, откуда пришёл, даты и формат касаний (до пяти), статус "
    "и что происходит. Дату выбирай в календаре двойным кликом. Колонка «Ближайшее касание» "
    "считается сама. Красным подсвечен клиент, у которого в ближайшие 7 дней нет касания — "
    "поставь ему дату.\n\n"
    "<b>Справа — часы по дням.</b> Каждый вечер до 19:00: факт за сегодня и план на завтра "
    "по трём процессам — привлечение, прогрев, закрытие. Строка сегодняшнего дня зелёная.\n\n"
    "Вечером я напомню в личку, если что-то не заполнено."
)

HELP = (
    "Я помогаю вести партнёрскую сеть: напоминаю заполнить таблицу и собираю аналитику.\n\n"
    "/mine — моя сводка и ссылка на мою вкладку\n"
    "/guide — как вести таблицу\n\n"
    "Руководителям:\n"
    "/report — утренняя сводка сейчас\n"
    "/week — итоги последних 7 дней\n"
    "/status — проверка настроек\n"
    "/bindchat — отправить в чате (или теме) партнёров, куда слать напоминания\n"
    "/setsheet ссылка — основная таблица\n"
    "/setsandbox ссылка — таблица-песочница новых партнёров\n"
    "/myid — id чата и пользователя"
)


def _cfg(ctx) -> Config:
    return ctx.application.bot_data["cfg"]


def _store(ctx) -> Store:
    return ctx.application.bot_data["store"]


def _today(cfg: Config):
    return datetime.now(cfg.tz).date()


def _apply_overrides(ctx) -> None:
    """Таблицы, заданные из Telegram (/setsheet, /setsandbox), главнее .env."""
    cfg, store = _cfg(ctx), _store(ctx)
    cfg.sheet_id = store.get_setting("sheet_id") or cfg.sheet_id
    cfg.sandbox_id = store.get_setting("sandbox_id") or cfg.sandbox_id


async def _network(ctx, with_gids: bool = True):
    _apply_overrides(ctx)
    cfg = _cfg(ctx)
    net = await asyncio.to_thread(load_network, cfg, _today(cfg), None, None, with_gids)
    ctx.application.bot_data["net"] = net
    return net


def _partners_chat(ctx) -> str | None:
    """«chat_id:thread» для напоминаний (поддерживает и привязку из прошлой версии бота)."""
    store = _store(ctx)
    return (store.get_setting("partners_chat") or store.get_setting("reminder_thread")
            or (f"{store.get_setting('partners_chat_id')}:" if store.get_setting("partners_chat_id") else None))


def _is_manager(ctx, update: Update) -> bool:
    cfg, u = _cfg(ctx), update.effective_user
    if u is None:
        return False
    if not cfg.managers:
        return True  # этап настройки
    return str(u.id) in cfg.managers or (u.username or "").lower() in cfg.managers


def _manager_ids(ctx) -> tuple[list[int], list[str]]:
    ids, missing = [], []
    for m in _cfg(ctx).managers:
        if m.isdigit():
            ids.append(int(m))
            continue
        found = _store(ctx).user_by_username(m)
        if found and found[1]:
            ids.append(found[0])
        else:
            missing.append("@" + m)
    return ids, missing


async def _send(ctx, chat_id: int, text: str, thread_id: int | None = None) -> None:
    for chunk in split_message(text):
        await ctx.bot.send_message(chat_id, chunk, parse_mode=ParseMode.HTML,
                                   message_thread_id=thread_id, disable_web_page_preview=True)


async def _deny(update: Update) -> None:
    await update.effective_message.reply_text("Эта команда доступна руководителям сети.")


# ---------- учёт пользователей ----------

async def track_user(update: Update, ctx: ContextTypes.DEFAULT_TYPE) -> None:
    u, chat = update.effective_user, update.effective_chat
    if u is None or u.is_bot:
        return
    _store(ctx).upsert_user(u.id, u.username, u.full_name,
                            private_ok=bool(chat and chat.type == "private"))


# ---------- команды ----------

async def cmd_start(update: Update, ctx: ContextTypes.DEFAULT_TYPE) -> None:
    lines = [HELP]
    u = update.effective_user
    net = ctx.application.bot_data.get("net")
    partner = net.directory.by_username(u.username or "") if net else None
    if _cfg(ctx).managers and _is_manager(ctx, update):
        lines.insert(0, "✅ Ты в списке руководителей: сводка будет приходить сюда каждое утро в 8:00.\n")
    elif partner:
        lines.insert(0, f"✅ Ты партнёр «{partner.name}»: вечером пришлю твою сводку по таблице.\n")
    await update.effective_message.reply_text("\n".join(lines))


async def cmd_guide(update: Update, ctx: ContextTypes.DEFAULT_TYPE) -> None:
    await update.effective_message.reply_text(GUIDE, parse_mode=ParseMode.HTML)


async def cmd_myid(update: Update, ctx: ContextTypes.DEFAULT_TYPE) -> None:
    await update.effective_message.reply_text(
        f"chat_id: {update.effective_chat.id}\nuser_id: {update.effective_user.id}\n"
        f"username: @{update.effective_user.username or '—'}")


async def cmd_report(update: Update, ctx: ContextTypes.DEFAULT_TYPE) -> None:
    if not _is_manager(ctx, update):
        return await _deny(update)
    await update.effective_message.reply_text("Собираю сводку…")
    net = await _network(ctx)
    await _send(ctx, update.effective_chat.id, morning_text(net))


async def cmd_week(update: Update, ctx: ContextTypes.DEFAULT_TYPE) -> None:
    if not _is_manager(ctx, update):
        return await _deny(update)
    net = await _network(ctx, with_gids=False)
    await _send(ctx, update.effective_chat.id, weekly_text(net))


async def cmd_mine(update: Update, ctx: ContextTypes.DEFAULT_TYPE) -> None:
    net = await _network(ctx)
    partner = net.directory.by_username(update.effective_user.username or "")
    if partner is None:
        await update.effective_message.reply_text(
            "Не нашла тебя на листе «Партнёры». Попроси руководителя добавить твой @username.")
        return
    stats = next((s for s in net.stats if s.name == partner.name), None)
    if stats is None:
        await update.effective_message.reply_text(
            "Твоя вкладка ещё создаётся — загляни через час 🌿")
        return
    await _send(ctx, update.effective_chat.id,
                render_partner_status(stats, net.today, net.tab_url(partner.name)))


async def cmd_bindchat(update: Update, ctx: ContextTypes.DEFAULT_TYPE) -> None:
    if not _is_manager(ctx, update):
        return await _deny(update)
    msg, chat = update.effective_message, update.effective_chat
    if chat.type == "private":
        await msg.reply_text("Эту команду нужно отправить в чате партнёров (можно внутри темы).")
        return
    thread = msg.message_thread_id if msg.is_topic_message else None
    _store(ctx).set_setting("partners_chat", f"{chat.id}:{thread or ''}")
    await msg.reply_text("✅ Готово: вечерние напоминания буду присылать сюда.")


def _sheet_id_from(text: str) -> str | None:
    m = re.search(r"/d/([A-Za-z0-9_-]{20,})", text) or re.fullmatch(r"([A-Za-z0-9_-]{20,})", text)
    return m.group(1) if m else None


async def _set_book(update: Update, ctx, key: str, sandbox: bool) -> None:
    if not _is_manager(ctx, update):
        return await _deny(update)
    sid = _sheet_id_from(" ".join(ctx.args or []).strip())
    if not sid:
        await update.effective_message.reply_text(
            f"Пришли команду со ссылкой: /{'setsandbox' if sandbox else 'setsheet'} https://docs.google.com/spreadsheets/d/…")
        return
    old = _store(ctx).get_setting(key)
    _store(ctx).set_setting(key, sid)
    await update.effective_message.reply_text("Проверяю таблицу и создаю вкладки — это займёт до пары минут…")
    net = await _network(ctx)
    book = next((b for b in net.books if b.sheet_id == sid), None)
    if book is None:
        _store(ctx).set_setting(key, old or "")
        await update.effective_message.reply_text(
            "Не получилось открыть таблицу, оставила прежнюю. Проверь, что адресу "
            "partners-pulse-bot@partners-pulse.iam.gserviceaccount.com дан доступ «Редактор».")
        return
    warnings = await asyncio.to_thread(sync, _cfg(ctx), net)
    net = await _network(ctx)
    what = "Песочница" if sandbox else "Основная таблица"
    text = (f"✅ {what} подключена: партнёров на листе «Партнёры» — {len(book.directory.partners)}.\n"
            "Вкладки, оглавление и сводка готовы.")
    if not book.directory.partners:
        text += "\nДобавь партнёров на лист «Партнёры» (имя, @ник) — через час у них появятся вкладки."
    if warnings:
        text += "\n⚠️ " + "\n⚠️ ".join(warnings)
    await update.effective_message.reply_text(text)


async def cmd_setsheet(update: Update, ctx: ContextTypes.DEFAULT_TYPE) -> None:
    await _set_book(update, ctx, "sheet_id", sandbox=False)


async def cmd_setsandbox(update: Update, ctx: ContextTypes.DEFAULT_TYPE) -> None:
    await _set_book(update, ctx, "sandbox_id", sandbox=True)


async def cmd_status(update: Update, ctx: ContextTypes.DEFAULT_TYPE) -> None:
    if not _is_manager(ctx, update):
        return await _deny(update)
    net = await _network(ctx, with_gids=False)
    ids, missing = _manager_ids(ctx)
    store = _store(ctx)
    lines = ["<b>Проверка настроек</b>"]
    for b in net.books:
        n = len([s for s in net.stats if s.sandbox == b.sandbox])
        lines.append(f"✅ {'Песочница' if b.sandbox else 'Основная таблица'}: вкладок партнёров {n}, "
                     f"в справочнике {len(b.directory.partners)}")
    if not any(b.sandbox for b in net.books):
        lines.append("⚪️ Песочница не подключена: /setsandbox ссылка")
    lines += [f"❌ {x}" for x in net.errors]
    lines.append(f"Руководители: получают {len(ids)}"
                 + (f", ещё не нажали Start: {', '.join(missing)}" if missing else ""))
    no_start = [p.name for p in net.directory.partners
                if p.username and not (store.user_by_username(p.username) or (0, False))[1]]
    lines.append(f"Партнёры без Start у бота (не получат личные напоминания): {len(no_start)}"
                 + (f" — {', '.join(no_start)}" if no_start else ""))
    chat = _partners_chat(ctx)
    lines.append("✅ Чат для напоминаний привязан" if chat else "⚠️ Чат не привязан: /bindchat в чате партнёров")
    cfg = _cfg(ctx)
    rt = cfg.reminder_time.strftime("%H:%M") if cfg.reminder_time else "выкл"
    lines.append(f"Сводка в {cfg.report_time:%H:%M}, напоминание в {rt}")
    await _send(ctx, update.effective_chat.id, "\n".join(lines))


# ---------- расписание ----------

async def job_morning(ctx: ContextTypes.DEFAULT_TYPE) -> None:
    cfg = _cfg(ctx)
    net = await _network(ctx)
    texts = [morning_text(net)]
    if cfg.weekly_day is not None and net.today.weekday() == cfg.weekly_day:
        texts.append(weekly_text(net))
    ids, missing = _manager_ids(ctx)
    if missing:
        log.warning("Руководители не нажали Start, сводку им не отправить: %s", missing)
    for chat_id in ids:
        for t in texts:
            try:
                await _send(ctx, chat_id, t)
            except Exception:
                log.exception("Не удалось отправить сводку %s", chat_id)


async def job_evening(ctx: ContextTypes.DEFAULT_TYPE) -> None:
    store = _store(ctx)
    net = await _network(ctx)
    chat = _partners_chat(ctx)
    if chat:
        chat_id, _, thread = chat.partition(":")
        try:
            await _send(ctx, int(chat_id), render_evening_chat(net.stats, net.today),
                        int(thread) if thread else None)
        except Exception:
            log.exception("Не удалось отправить напоминание в чат")
    directory = net.directory
    for s in net.stats:
        if not needs_evening_ping(s):
            continue
        p = next((x for x in directory.partners if x.name == s.name), None)
        found = store.user_by_username(p.username) if p and p.username else None
        if not found or not found[1]:
            continue  # партнёр не нажал Start — написать первым бот не может
        try:
            await _send(ctx, found[0], render_partner_status(s, net.today, net.tab_url(s.name)))
        except Exception:
            log.exception("Не удалось написать партнёру %s", s.name)


async def job_sync(ctx: ContextTypes.DEFAULT_TYPE) -> None:
    net = await _network(ctx)
    for w in await asyncio.to_thread(sync, _cfg(ctx), net):
        log.warning(w)


def main() -> None:
    cfg = load_config()
    if not cfg.bot_token:
        raise SystemExit("Не задан TELEGRAM_BOT_TOKEN")
    app = Application.builder().token(cfg.bot_token).build()
    app.bot_data["cfg"] = cfg
    app.bot_data["store"] = Store(cfg.db_path)

    app.add_handler(TypeHandler(Update, track_user), group=-1)
    app.add_handler(CommandHandler(["start", "help"], cmd_start))
    app.add_handler(CommandHandler(["guide", "template"], cmd_guide))
    app.add_handler(CommandHandler("myid", cmd_myid))
    app.add_handler(CommandHandler("report", cmd_report))
    app.add_handler(CommandHandler("week", cmd_week))
    app.add_handler(CommandHandler("mine", cmd_mine))
    app.add_handler(CommandHandler("bindchat", cmd_bindchat))
    app.add_handler(CommandHandler("setsheet", cmd_setsheet))
    app.add_handler(CommandHandler("setsandbox", cmd_setsandbox))
    app.add_handler(CommandHandler("status", cmd_status))

    jq = app.job_queue
    jq.run_daily(job_morning, time=cfg.report_time.replace(tzinfo=cfg.tz), name="morning")
    if cfg.reminder_time:
        jq.run_daily(job_evening, time=cfg.reminder_time.replace(tzinfo=cfg.tz), name="evening")
    jq.run_repeating(job_sync, interval=3600, first=30, name="sync")

    log.info("Бот запущен: сводка в %s, напоминание в %s, руководители %s",
             cfg.report_time, cfg.reminder_time, cfg.managers)
    app.run_polling(allowed_updates=Update.ALL_TYPES)


if __name__ == "__main__":
    main()
