"""Телеграм-бот: утренний отчёт по воронке + учёт часов по хэштегам из чата партнёров."""
from __future__ import annotations

import asyncio
import logging
from datetime import datetime, timedelta

from telegram import Update
from telegram.constants import ParseMode
from telegram.ext import (Application, CommandHandler, ContextTypes, MessageHandler,
                          filters)

from .analytics import e, fmt_hours, split_message
from .config import Config, load_config
from .hashtags import PROCESSES, Store, parse_message
from .report import build_report

logging.basicConfig(format="%(asctime)s %(levelname)s %(name)s: %(message)s", level=logging.INFO)
logging.getLogger("httpx").setLevel(logging.WARNING)
log = logging.getLogger("partner-bot")

TEMPLATE = (
    "<b>Эталон отчёта в чат</b> (каждый процесс с новой строки):\n\n"
    "<code>#отчет #план_Фамилия\n"
    "#знакомство 2 ч — презентация экосистемы для Котовой, показала личный кабинет\n"
    "#прогрев 1,5 ч — переписка с Соколовой, отправила 2 кейса\n"
    "#сделка 30 мин — Жуков: согласовали пакет, ждём оплату</code>\n\n"
    "Часы можно писать так: <code>2 ч</code>, <code>1,5 часа</code>, <code>40 мин</code>, <code>1:30</code>.\n"
    "Ошиблись — просто отредактируйте сообщение, бот пересчитает.\n"
    "Хэштег <code>#план_Фамилия</code> достаточно указать один раз, дальше бот запомнит."
)

HELP = (
    "Я собираю аналитику партнёрской сети.\n\n"
    "• Каждое утро присылаю сводку по таблице воронки: сколько лидов в работе, "
    "с кем касания сегодня и завтра, у кого не запланирован следующий шаг.\n"
    "• В чате партнёров считаю часы по процессам из сообщений с хэштегами.\n\n"
    "/report — сводка прямо сейчас\n"
    "/hours 7 — часы по процессам за N дней\n"
    "/template — эталон сообщения для партнёров\n"
    "/myid — id чата и пользователя (для настройки)"
)


def _cfg(ctx: ContextTypes.DEFAULT_TYPE) -> Config:
    return ctx.application.bot_data["cfg"]


def _store(ctx: ContextTypes.DEFAULT_TYPE) -> Store:
    return ctx.application.bot_data["store"]


def _is_manager(cfg: Config, update: Update) -> bool:
    if not cfg.report_chat_ids and not cfg.admin_ids:
        return True  # этап настройки: ещё никого не назначили
    allowed = set(cfg.report_chat_ids) | set(cfg.admin_ids)
    chat_id = update.effective_chat.id if update.effective_chat else None
    user_id = update.effective_user.id if update.effective_user else None
    return chat_id in allowed or user_id in allowed


async def _send(ctx, chat_id: int, text: str) -> None:
    for chunk in split_message(text):
        await ctx.bot.send_message(chat_id, chunk, parse_mode=ParseMode.HTML,
                                   disable_web_page_preview=True)


async def _make_report(ctx) -> str:
    cfg = _cfg(ctx)
    today = datetime.now(cfg.tz).date()
    return await asyncio.to_thread(build_report, cfg, today, _store(ctx))


# ---------- команды ----------

async def cmd_start(update: Update, ctx: ContextTypes.DEFAULT_TYPE) -> None:
    await update.effective_message.reply_text(HELP)


async def cmd_myid(update: Update, ctx: ContextTypes.DEFAULT_TYPE) -> None:
    await update.effective_message.reply_text(
        f"chat_id: {update.effective_chat.id}\nuser_id: {update.effective_user.id}")


async def cmd_template(update: Update, ctx: ContextTypes.DEFAULT_TYPE) -> None:
    await update.effective_message.reply_text(TEMPLATE, parse_mode=ParseMode.HTML)


async def cmd_report(update: Update, ctx: ContextTypes.DEFAULT_TYPE) -> None:
    if not _is_manager(_cfg(ctx), update):
        await update.effective_message.reply_text("Сводка доступна только руководителям сети.")
        return
    await update.effective_message.reply_text("Собираю сводку…")
    await _send(ctx, update.effective_chat.id, await _make_report(ctx))


async def cmd_hours(update: Update, ctx: ContextTypes.DEFAULT_TYPE) -> None:
    cfg = _cfg(ctx)
    if not _is_manager(cfg, update):
        await update.effective_message.reply_text("Доступно только руководителям сети.")
        return
    try:
        days = max(1, min(90, int(ctx.args[0]))) if ctx.args else 7
    except ValueError:
        days = 7
    today = datetime.now(cfg.tz).date()
    start, end = today - timedelta(days=days), today
    data = _store(ctx).summary(start, end)
    if not data:
        await update.effective_message.reply_text(f"За {days} дн. отчётов с хэштегами нет.")
        return
    lines = [f"⏱ <b>Часы по процессам с {start:%d.%m} по {end:%d.%m}</b>"]
    for partner, procs in sorted(data.items()):
        lines.append(f"\n<b>{e(partner)}</b>")
        for key, (title, _) in PROCESSES.items():
            lines.append(f"  {title}: {fmt_hours(procs.get(key, 0))}")
        lines.append(f"  Итого: <b>{fmt_hours(sum(procs.values()))}</b>")
    await _send(ctx, update.effective_chat.id, "\n".join(lines))


# ---------- сообщения партнёров ----------

async def on_message(update: Update, ctx: ContextTypes.DEFAULT_TYPE) -> None:
    msg = update.effective_message
    cfg = _cfg(ctx)
    if msg is None or update.effective_user is None:
        return
    chat = update.effective_chat
    if cfg.partners_chat_id and chat.id != cfg.partners_chat_id and chat.type != "private":
        return
    report = parse_message(msg.text or msg.caption or "")
    if report is None:
        return

    store = _store(ctx)
    user = update.effective_user
    if report.partner:
        store.remember_partner(user.id, report.partner)
    partner = report.partner or store.partner_for(user.id) or user.full_name
    day = msg.date.astimezone(cfg.tz).date()
    store.save(chat.id, msg.message_id, user.id, partner, day, report)
    log.info("Отчёт %s: %s", partner, [(x.process, x.hours) for x in report.entries])

    missing = report.missing_hours
    try:
        await msg.set_reaction("✍" if not missing else "🤔")
    except Exception:  # реакции могут быть выключены в чате
        pass
    if missing and update.message is not None:  # подсказываем один раз, не на каждое редактирование
        names = ", ".join(PROCESSES[m][0].lower() for m in missing)
        await msg.reply_text(
            f"Записала отчёт, но не вижу часов: {names}. "
            "Допишите, например «#прогрев 1,5 ч — …», можно прямо отредактировать сообщение.")


# ---------- расписание ----------

async def job_morning_report(ctx: ContextTypes.DEFAULT_TYPE) -> None:
    cfg = _cfg(ctx)
    if not cfg.report_chat_ids:
        log.warning("REPORT_CHAT_IDS пуст — утренний отчёт некому отправить")
        return
    text = await _make_report(ctx)
    for chat_id in cfg.report_chat_ids:
        try:
            await _send(ctx, chat_id, text)
        except Exception:
            log.exception("Не удалось отправить отчёт в %s", chat_id)


async def job_reminder(ctx: ContextTypes.DEFAULT_TYPE) -> None:
    cfg = _cfg(ctx)
    if not cfg.partners_chat_id:
        return
    await ctx.bot.send_message(
        cfg.partners_chat_id,
        "Коллеги, напоминание 🌿 До 19:00 обновите табличку воронки: статус, что происходит "
        "и <b>дату следующего касания</b> у каждого лида. И отчёт по часам с хэштегами "
        "#знакомство, #прогрев, #сделка — эталон по команде /template.",
        parse_mode=ParseMode.HTML,
    )


def main() -> None:
    cfg = load_config()
    if not cfg.bot_token:
        raise SystemExit("Не задан TELEGRAM_BOT_TOKEN")
    app = Application.builder().token(cfg.bot_token).build()
    app.bot_data["cfg"] = cfg
    app.bot_data["store"] = Store(cfg.db_path)

    app.add_handler(CommandHandler(["start", "help"], cmd_start))
    app.add_handler(CommandHandler("myid", cmd_myid))
    app.add_handler(CommandHandler("template", cmd_template))
    app.add_handler(CommandHandler("report", cmd_report))
    app.add_handler(CommandHandler("hours", cmd_hours))
    app.add_handler(MessageHandler(
        filters.UpdateType.MESSAGES & (filters.TEXT | filters.CAPTION) & ~filters.COMMAND,
        on_message))

    jq = app.job_queue
    jq.run_daily(job_morning_report, time=cfg.report_time.replace(tzinfo=cfg.tz), name="morning")
    if cfg.reminder_time:
        jq.run_daily(job_reminder, time=cfg.reminder_time.replace(tzinfo=cfg.tz), name="reminder")

    log.info("Бот запущен: %d табл., отчёт в %s, получатели %s",
             len(cfg.sheets), cfg.report_time, cfg.report_chat_ids)
    app.run_polling(allowed_updates=Update.ALL_TYPES)


if __name__ == "__main__":
    main()
