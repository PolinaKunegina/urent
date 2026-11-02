"""Телеграм-бот партнёрской сети.

Расписание (по умолчанию, Москва):
  08:00  руководителям в личку — сводка по воронке + часы из чата; по понедельникам ещё итоги недели
  18:30  в чат партнёров — напоминание и кто ещё не отчитался; каждому партнёру в личку — его лиды без плана
  каждый час — обновление подсветки «нет плана» и листа с часами в Google Таблице
"""
from __future__ import annotations

import asyncio
import logging
from datetime import datetime, timedelta

from telegram import Update
from telegram.constants import ParseMode
from telegram.ext import (Application, CommandHandler, ContextTypes,
                          MessageHandler, TypeHandler, filters)

from .analytics import e, fmt_hours, render_partner_evening, split_message
from .config import Config, load_config
from .hashtags import PROCESSES, Store, parse_message
from .partners import Directory
from .report import (evening_chat_text, evening_partner_texts, load_snapshot, morning_report,
                     weekly_report, write_back)

logging.basicConfig(format="%(asctime)s %(levelname)s %(name)s: %(message)s", level=logging.INFO)
logging.getLogger("httpx").setLevel(logging.WARNING)
log = logging.getLogger("partner-bot")

TEMPLATE = (
    "<b>Эталон отчёта в чат</b> (каждый процесс с новой строки):\n\n"
    "<code>#отчет\n"
    "#знакомство 2 ч — презентация экосистемы для Котовой, показала личный кабинет\n"
    "#прогрев 1,5 ч — переписка с Соколовой, отправила 2 кейса\n"
    "#сделка 30 мин — Жуков: согласовали пакет, ждём оплату</code>\n\n"
    "Часы можно писать так: <code>2 ч</code>, <code>1,5 часа</code>, <code>40 мин</code>, <code>1:30</code>.\n"
    "Ошиблись — просто отредактируйте сообщение, бот пересчитает.\n"
    "Отчёт принят, если бот поставил ✍."
)

HELP = (
    "Я собираю аналитику партнёрской сети.\n\n"
    "• Каждое утро присылаю руководителям сводку по таблице воронки.\n"
    "• В чате партнёров считаю часы по хэштегам #знакомство, #прогрев, #сделка.\n"
    "• Вечером напоминаю партнёрам про лиды без следующего касания.\n\n"
    "/template — эталон сообщения с часами\n"
    "/mine — мои лиды без плана и касания на завтра\n"
    "/report — сводка сейчас (руководителям)\n"
    "/week — итоги последних 7 дней (руководителям)\n"
    "/hours 7 — часы по процессам за N дней (руководителям)\n"
    "/status — проверка настроек (руководителям)\n"
    "/bindchat — сделать этот чат чатом партнёров (руководителям)"
)


# ---------- помощники ----------

def _cfg(ctx) -> Config:
    return ctx.application.bot_data["cfg"]


def _store(ctx) -> Store:
    return ctx.application.bot_data["store"]


def _dir(ctx) -> Directory:
    """Справочник из последнего чтения таблицы (лист «Партнёры»), иначе из файла."""
    return ctx.application.bot_data.get("directory") or Directory.load(_cfg(ctx).partners_file)


def _today(cfg: Config):
    return datetime.now(cfg.tz).date()


def _partners_chat(ctx) -> int | None:
    cfg = _cfg(ctx)
    if cfg.partners_chat_id:
        return cfg.partners_chat_id
    v = _store(ctx).get_setting("partners_chat_id")
    return int(v) if v else None


def _is_manager(ctx, update: Update) -> bool:
    cfg = _cfg(ctx)
    u = update.effective_user
    if u is None:
        return False
    if not cfg.managers:
        return True  # этап настройки
    return str(u.id) in cfg.managers or (u.username or "").lower() in cfg.managers


def _manager_ids(ctx) -> tuple[list[int], list[str]]:
    """(id, кому можно писать), (кто ещё не нажал /start)."""
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


async def _send(ctx, chat_id: int, text: str) -> None:
    for chunk in split_message(text):
        await ctx.bot.send_message(chat_id, chunk, parse_mode=ParseMode.HTML,
                                   disable_web_page_preview=True)


async def _snapshot(ctx):
    cfg = _cfg(ctx)
    snap = await asyncio.to_thread(load_snapshot, cfg, _today(cfg))
    if snap.funnel is not None:
        ctx.application.bot_data["directory"] = snap.directory
    return snap


# ---------- учёт пользователей ----------

async def track_user(update: Update, ctx: ContextTypes.DEFAULT_TYPE) -> None:
    u, chat = update.effective_user, update.effective_chat
    if u is None or u.is_bot:
        return
    _store(ctx).upsert_user(u.id, u.username, u.full_name,
                            private_ok=bool(chat and chat.type == "private"))


# ---------- команды ----------

async def cmd_start(update: Update, ctx: ContextTypes.DEFAULT_TYPE) -> None:
    u = update.effective_user
    lines = [HELP]
    partner = _dir(ctx).by_username(u.username or "")
    if _is_manager(ctx, update) and _cfg(ctx).managers:
        lines.insert(0, "✅ Вы в списке руководителей: сводка будет приходить сюда каждое утро.\n")
    elif partner:
        lines.insert(0, f"✅ Вы партнёр «{partner.name}»: вечером сюда будут приходить ваши лиды без плана.\n")
    await update.effective_message.reply_text("\n".join(lines))


async def cmd_myid(update: Update, ctx: ContextTypes.DEFAULT_TYPE) -> None:
    await update.effective_message.reply_text(
        f"chat_id: {update.effective_chat.id}\nuser_id: {update.effective_user.id}\n"
        f"username: @{update.effective_user.username or '—'}")


async def cmd_template(update: Update, ctx: ContextTypes.DEFAULT_TYPE) -> None:
    await update.effective_message.reply_text(TEMPLATE, parse_mode=ParseMode.HTML)


async def _deny(update: Update) -> None:
    await update.effective_message.reply_text("Эта команда доступна руководителям сети.")


async def cmd_report(update: Update, ctx: ContextTypes.DEFAULT_TYPE) -> None:
    if not _is_manager(ctx, update):
        return await _deny(update)
    await update.effective_message.reply_text("Собираю сводку…")
    cfg = _cfg(ctx)
    snap = await _snapshot(ctx)
    await _send(ctx, update.effective_chat.id, morning_report(cfg, snap, _store(ctx), _today(cfg)))


async def cmd_week(update: Update, ctx: ContextTypes.DEFAULT_TYPE) -> None:
    if not _is_manager(ctx, update):
        return await _deny(update)
    snap = await _snapshot(ctx)
    text = weekly_report(snap, _store(ctx), _today(_cfg(ctx)))
    await _send(ctx, update.effective_chat.id, text or "Не удалось прочитать таблицу.")


async def cmd_hours(update: Update, ctx: ContextTypes.DEFAULT_TYPE) -> None:
    if not _is_manager(ctx, update):
        return await _deny(update)
    try:
        days = max(1, min(90, int(ctx.args[0]))) if ctx.args else 7
    except ValueError:
        days = 7
    today = _today(_cfg(ctx))
    start = today - timedelta(days=days)
    data = _store(ctx).summary(start, today)
    if not data:
        await update.effective_message.reply_text(f"За {days} дн. отчётов с хэштегами нет.")
        return
    lines = [f"⏱ <b>Часы по процессам с {start:%d.%m} по {today:%d.%m}</b>"]
    for partner, procs in sorted(data.items()):
        lines.append(f"\n<b>{e(partner)}</b>")
        for key, (title, _) in PROCESSES.items():
            lines.append(f"  {title}: {fmt_hours(procs.get(key, 0))}")
        lines.append(f"  Итого: <b>{fmt_hours(sum(procs.values()))}</b>")
    await _send(ctx, update.effective_chat.id, "\n".join(lines))


async def cmd_mine(update: Update, ctx: ContextTypes.DEFAULT_TYPE) -> None:
    partner = _dir(ctx).by_username(update.effective_user.username or "")
    if partner is None:
        await update.effective_message.reply_text(
            "Не нашла вас в списке партнёров. Попросите руководителя добавить ваш @username.")
        return
    snap = await _snapshot(ctx)
    p = snap.stats.partners.get(partner.name) if snap.stats else None
    text = render_partner_evening(p, _today(_cfg(ctx))) if p else None
    await _send(ctx, update.effective_chat.id,
                text or f"🟢 {e(partner.name)}, у всех ваших лидов есть дата следующего касания.")


async def cmd_bindchat(update: Update, ctx: ContextTypes.DEFAULT_TYPE) -> None:
    if not _is_manager(ctx, update):
        return await _deny(update)
    chat = update.effective_chat
    if chat.type == "private":
        await update.effective_message.reply_text("Эту команду нужно отправить в чате партнёров.")
        return
    _store(ctx).set_setting("partners_chat_id", str(chat.id))
    await update.effective_message.reply_text(
        "✅ Готово: это чат партнёров. Здесь я считаю отчёты с хэштегами и в 18:30 напоминаю про таблицу.")


async def cmd_status(update: Update, ctx: ContextTypes.DEFAULT_TYPE) -> None:
    if not _is_manager(ctx, update):
        return await _deny(update)
    cfg = _cfg(ctx)
    snap = await _snapshot(ctx)
    ids, missing = _manager_ids(ctx)
    d = _dir(ctx)
    no_tg = [p.name for p in d.partners if p.username and not _store(ctx).user_by_username(p.username)]
    lines = ["<b>Проверка настроек</b>"]
    if snap.stats:
        lines.append(f"✅ Таблица читается: {len(snap.funnel.leads)} лидов, лист «{e(snap.funnel.tab)}»")
        lines.append("✅ Это Google Таблица, подсветка работает" if snap.is_gsheet
                     else "⚠️ Это .xlsx: подсветка и лист с часами не работают, нужна Google Таблица")
    else:
        lines += [f"❌ {e(x)}" for x in snap.errors]
    lines.append(f"Руководители: получают {len(ids)}" +
                 (f", ещё не нажали /start: {e(', '.join(missing))}" if missing else ""))
    lines.append(f"Партнёров в справочнике: {len(d.partners)}" +
                 (f", ещё не писали боту: {e(', '.join(no_tg))}" if no_tg else ""))
    chat = _partners_chat(ctx)
    lines.append(f"Чат партнёров: {chat}" if chat else "⚠️ Чат партнёров не привязан: /bindchat в нём")
    rt = cfg.reminder_time.strftime("%H:%M") if cfg.reminder_time else "выкл"
    lines.append(f"Сводка в {cfg.report_time:%H:%M}, напоминание в {rt}")
    await _send(ctx, update.effective_chat.id, "\n".join(lines))


# ---------- сообщения партнёров ----------

def _partner_name(ctx, user, report) -> str:
    d, store = _dir(ctx), _store(ctx)
    by_user = d.by_username(user.username or "")
    if by_user:
        return by_user.name
    if report.partner:
        name, _ = d.resolve(report.partner)
        store.remember_partner(user.id, name)
        return name
    return store.partner_for(user.id) or user.full_name


async def on_message(update: Update, ctx: ContextTypes.DEFAULT_TYPE) -> None:
    msg, user, chat = update.effective_message, update.effective_user, update.effective_chat
    if msg is None or user is None:
        return
    partners_chat = _partners_chat(ctx)
    if chat.type != "private" and partners_chat and chat.id != partners_chat:
        return
    report = parse_message(msg.text or msg.caption or "")
    if report is None:
        return
    partner = _partner_name(ctx, user, report)
    day = msg.date.astimezone(_cfg(ctx).tz).date()
    _store(ctx).save(chat.id, msg.message_id, user.id, partner, day, report)
    log.info("Отчёт %s: %s", partner, [(x.process, x.hours) for x in report.entries])

    missing = report.missing_hours
    try:
        await msg.set_reaction("✍" if not missing else "🤔")
    except Exception:  # реакции могут быть выключены в чате
        pass
    if missing and update.message is not None:  # не дёргаем на каждое редактирование
        names = ", ".join(PROCESSES[m][0].lower() for m in missing)
        await msg.reply_text(
            f"Записала отчёт, но не вижу часов: {names}. "
            "Допишите, например «#прогрев 1,5 ч — …», можно прямо отредактировать сообщение.")


# ---------- расписание ----------

async def job_morning(ctx: ContextTypes.DEFAULT_TYPE) -> None:
    cfg, store = _cfg(ctx), _store(ctx)
    today = _today(cfg)
    snap = await _snapshot(ctx)
    warning = await asyncio.to_thread(write_back, cfg, snap, store)
    texts = [morning_report(cfg, snap, store, today, warning)]
    if cfg.weekly_day is not None and today.weekday() == cfg.weekly_day:
        weekly = weekly_report(snap, store, today)
        if weekly:
            texts.append(weekly)
    ids, missing = _manager_ids(ctx)
    if missing:
        log.warning("Руководители не нажали /start, сводку им не отправить: %s", missing)
    for chat_id in ids:
        for t in texts:
            try:
                await _send(ctx, chat_id, t)
            except Exception:
                log.exception("Не удалось отправить сводку %s", chat_id)


async def job_evening(ctx: ContextTypes.DEFAULT_TYPE) -> None:
    cfg, store = _cfg(ctx), _store(ctx)
    today = _today(cfg)
    snap = await _snapshot(ctx)
    await asyncio.to_thread(write_back, cfg, snap, store)
    chat = _partners_chat(ctx)
    if chat:
        try:
            await _send(ctx, chat, evening_chat_text(snap, store, today))
        except Exception:
            log.exception("Не удалось отправить напоминание в чат")
    d = _dir(ctx)
    for name, text in evening_partner_texts(snap, today).items():
        p = next((x for x in d.partners if x.name == name), None)
        found = store.user_by_username(p.username) if p and p.username else None
        if not found or not found[1]:
            continue  # партнёр ещё не нажал /start в личке — написать первым бот не может
        try:
            await _send(ctx, found[0], text)
        except Exception:
            log.exception("Не удалось написать партнёру %s", name)


async def job_sync(ctx: ContextTypes.DEFAULT_TYPE) -> None:
    cfg = _cfg(ctx)
    snap = await _snapshot(ctx)
    warning = await asyncio.to_thread(write_back, cfg, snap, _store(ctx))
    if warning:
        log.warning(warning)


async def job_refresh_directory(ctx: ContextTypes.DEFAULT_TYPE) -> None:
    await _snapshot(ctx)


def main() -> None:
    cfg = load_config()
    if not cfg.bot_token:
        raise SystemExit("Не задан TELEGRAM_BOT_TOKEN")
    app = Application.builder().token(cfg.bot_token).build()
    app.bot_data["cfg"] = cfg
    app.bot_data["store"] = Store(cfg.db_path)

    app.add_handler(TypeHandler(Update, track_user), group=-1)
    app.add_handler(CommandHandler(["start", "help"], cmd_start))
    app.add_handler(CommandHandler("myid", cmd_myid))
    app.add_handler(CommandHandler("template", cmd_template))
    app.add_handler(CommandHandler("report", cmd_report))
    app.add_handler(CommandHandler("week", cmd_week))
    app.add_handler(CommandHandler("hours", cmd_hours))
    app.add_handler(CommandHandler("mine", cmd_mine))
    app.add_handler(CommandHandler("bindchat", cmd_bindchat))
    app.add_handler(CommandHandler("status", cmd_status))
    app.add_handler(MessageHandler(
        filters.UpdateType.MESSAGES & (filters.TEXT | filters.CAPTION) & ~filters.COMMAND,
        on_message))

    jq = app.job_queue
    jq.run_daily(job_morning, time=cfg.report_time.replace(tzinfo=cfg.tz), name="morning")
    if cfg.reminder_time:
        jq.run_daily(job_evening, time=cfg.reminder_time.replace(tzinfo=cfg.tz), name="evening")
    if cfg.write_back:
        jq.run_repeating(job_sync, interval=3600, first=30, name="sync")
    else:
        jq.run_repeating(job_refresh_directory, interval=3600, first=30, name="directory")

    log.info("Бот запущен: сводка в %s, напоминание в %s, руководители %s",
             cfg.report_time, cfg.reminder_time, cfg.managers)
    app.run_polling(allowed_updates=Update.ALL_TYPES)


if __name__ == "__main__":
    main()
