"""Телеграм-бот партнёрской сети.

Расписание (по умолчанию, Москва):
  08:00  руководителям в личку — сводка по воронке + часы из чата; по понедельникам ещё итоги недели
  18:30  в чат партнёров — напоминание и кто ещё не отчитался; каждому партнёру в личку — его лиды без плана
  каждый час — обновление подсветки «нет плана» и листа с часами в Google Таблице
"""
from __future__ import annotations

import asyncio
import logging
import os
import re
from datetime import datetime, timedelta

from telegram import InlineKeyboardButton, InlineKeyboardMarkup, Update
from telegram.constants import ParseMode
from telegram.ext import (Application, CallbackQueryHandler, CommandHandler, ContextTypes,
                          MessageHandler, TypeHandler, filters)

from .analytics import e, fmt_hours, render_partner_evening, split_message
from .config import Config, load_config
from .hashtags import PROCESSES, Store, parse_message
from .flow import confirm_new_lead, detect_kind, process_report
from .partners import Directory
from .report import (evening_chat_text, evening_partner_texts, load_snapshot, morning_report,
                     weekly_report, write_back)

logging.basicConfig(format="%(asctime)s %(levelname)s %(name)s: %(message)s", level=logging.INFO)
logging.getLogger("httpx").setLevel(logging.WARNING)
log = logging.getLogger("partner-bot")

TEMPLATE = (
    "<b>Как писать план и факт</b>\n\n"
    "Пиши как удобно, своими словами: бот на ИИ сам разберёт сообщение и внесёт в таблицу.\n"
    "Утром — в тему <b>План</b>, вечером до 19:00 — в тему <b>Факт</b> "
    "(или начни сообщение со слова «План» / «Факт»).\n\n"
    "<b>Пример факта:</b>\n"
    "<code>Факт\n"
    "Знакомство с экосистемой — 2 ч: нетворкинг в бизнес-клубе, пригласила 3 человек на четверг.\n"
    "Прогрев — 1,5 ч: переписка с Котовой Верой, отправила кейсы; созвон с Орловым Петром.\n"
    "Сделки — 1 ч: диагностика с Жуковым Ильёй, оплатил вход 5000.\n\n"
    "Котова Вера — интересуется, следующее касание в четверг, встреча с Антоном.\n"
    "Орлов Пётр — вторая встреча прошла, созвон 06.11.\n"
    "Жуков Илья — оплатил вход, следующий шаг — трек с Киром 10.11.</code>\n\n"
    "Главное по каждому клиенту: имя, что происходит, когда следующее касание. "
    "Если клиента ещё нет в таблице, бот спросит, добавить ли его."
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
    "/bindchat — сделать этот чат чатом партнёров (руководителям)\n"
    "/setsheet ссылка — переключить бота на другую таблицу (руководителям)\n"
    "/bindtopic план | факт — отправить в теме форума, чтобы бот читал её (руководителям)\n"
    "/setkey ключ — ключ Claude для разбора сообщений, только в личке (руководителям)"
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


def _parser(ctx):
    """Разборщик на Claude, если задан ключ (в .env или командой /setkey)."""
    key = os.environ.get("ANTHROPIC_API_KEY") or _store(ctx).get_setting("anthropic_key")
    if not key:
        return None
    cached = ctx.application.bot_data.get("parser")
    if cached and cached[0] == key:
        return cached[1]
    from .llm import ReportParser
    parser = ReportParser(key)
    ctx.application.bot_data["parser"] = (key, parser)
    return parser


def _apply_sheet_override(ctx) -> None:
    """Таблицу можно сменить из Telegram командой /setsheet — она главнее .env."""
    override = _store(ctx).get_setting("sheet_id")
    if override:
        _cfg(ctx).sheet_id = override


async def _snapshot(ctx):
    cfg = _cfg(ctx)
    _apply_sheet_override(ctx)
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
        lines.insert(0, "✅ Ты в списке руководителей: сводка будет приходить сюда каждое утро.\n")
    elif partner:
        lines.insert(0, f"✅ Ты партнёр «{partner.name}»: вечером сюда будут приходить твои лиды без плана.\n")
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
            "Не нашла тебя в списке партнёров. Попроси руководителя добавить твой @username.")
        return
    snap = await _snapshot(ctx)
    p = snap.stats.partners.get(partner.name) if snap.stats else None
    text = render_partner_evening(p, _today(_cfg(ctx))) if p else None
    await _send(ctx, update.effective_chat.id,
                text or f"🟢 {e(partner.name)}, у всех твоих лидов есть дата следующего касания.")


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


async def cmd_setsheet(update: Update, ctx: ContextTypes.DEFAULT_TYPE) -> None:
    if not _is_manager(ctx, update):
        return await _deny(update)
    raw = " ".join(ctx.args or []).strip()
    m = re.search(r"/d/([A-Za-z0-9_-]{20,})", raw) or re.fullmatch(r"([A-Za-z0-9_-]{20,})", raw)
    if not m:
        await update.effective_message.reply_text(
            "Пришли команду со ссылкой на таблицу, например:\n/setsheet https://docs.google.com/spreadsheets/d/…/edit")
        return
    cfg = _cfg(ctx)
    old = cfg.sheet_id
    cfg.sheet_id = m.group(1)
    snap = await asyncio.to_thread(load_snapshot, cfg, _today(cfg))
    if snap.funnel is None:
        cfg.sheet_id = old
        await update.effective_message.reply_text(
            "Не получилось открыть эту таблицу, оставила прежнюю.\n"
            f"{snap.errors[0] if snap.errors else ''}\n\n"
            "Проверь, что адресу partners-pulse-bot@partners-pulse.iam.gserviceaccount.com "
            "дан доступ «Редактор».")
        return
    _store(ctx).set_setting("sheet_id", cfg.sheet_id)
    ctx.application.bot_data["directory"] = snap.directory
    kind = ("✅ Это Google Таблица: подсветка и лист с часами заработают в течение часа."
            if snap.is_gsheet else "⚠️ Это всё ещё .xlsx: бот только читает, подсветки не будет.")
    await update.effective_message.reply_text(
        f"Готово, теперь работаю с этой таблицей: {len(snap.funnel.leads)} лидов на листе "
        f"«{snap.funnel.tab}».\n{kind}")
    if snap.is_gsheet and cfg.write_back:
        warning = await asyncio.to_thread(write_back, cfg, snap, _store(ctx))
        if warning:
            await update.effective_message.reply_text(f"⚠️ {warning}")


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
    lines.append("✅ ИИ-разбор плана и факта включён" if _parser(ctx)
                 else "⚠️ ИИ-разбор выключен: пришли ключ Claude командой /setkey в личке")
    store = _store(ctx)
    with store.db:
        topics = store.db.execute("SELECT key, value FROM settings WHERE key LIKE 'topic:%'").fetchall()
    bound = {v for _, v in topics}
    lines.append("Темы: " + ", ".join(f"{'план' if k == 'plan' else 'факт'} ✅" for k in sorted(bound))
                 if bound else "⚠️ Темы План/Факт не привязаны: /bindtopic в каждой теме")
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
    text = msg.text or msg.caption or ""
    store, cfg = _store(ctx), _cfg(ctx)
    thread_id = msg.message_thread_id if msg.is_topic_message else None
    kind = detect_kind(store, chat.id, thread_id, text)
    parser = _parser(ctx)

    if kind and parser:
        legacy = parse_message(text)  # только ради #план_Фамилия, если указан
        partner = _partner_name(ctx, user, legacy or type("R", (), {"partner": None})())
        try:
            await msg.set_reaction("👀")
        except Exception:
            pass
        _apply_sheet_override(ctx)
        try:
            outcome = await asyncio.to_thread(
                process_report, cfg, store, parser, text=text, kind=kind, partner=partner,
                today=msg.date.astimezone(cfg.tz).date(), chat_id=chat.id,
                message_id=msg.message_id, user_id=user.id)
        except Exception:
            log.exception("Не удалось разобрать отчёт")
            await msg.reply_text("Не получилось разобрать сообщение, попробую позже. "
                                 "Если повторится — напиши руководителю.")
            return
        markup = None
        if outcome.pending:
            markup = InlineKeyboardMarkup([
                [InlineKeyboardButton(f"➕ {name[:28]} — новый", callback_data=f"newlead:{pid}:yes"),
                 InlineKeyboardButton("Нет", callback_data=f"newlead:{pid}:no")]
                for pid, name in outcome.pending])
        try:
            await msg.set_reaction("✍")
        except Exception:
            pass
        await msg.reply_text(outcome.reply, reply_markup=markup)
        return

    # запасной путь без ИИ: старые хэштеги с часами
    report = parse_message(text)
    if report is None:
        if kind and not parser:
            log.warning("Пришёл %s, но ключ Claude не задан (/setkey)", kind)
        return
    partner = _partner_name(ctx, user, report)
    day = msg.date.astimezone(cfg.tz).date()
    store.save(chat.id, msg.message_id, user.id, partner, day, report, kind=kind or "fact")
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
            "Допиши, например «#прогрев 1,5 ч — …», можно прямо отредактировать сообщение.")


async def on_newlead(update: Update, ctx: ContextTypes.DEFAULT_TYPE) -> None:
    q = update.callback_query
    _, pid, answer = q.data.split(":")
    store, cfg = _store(ctx), _cfg(ctx)
    if answer == "no":
        store.pop_pending(int(pid))
        await q.answer("Ок, не добавляю")
        await q.message.reply_text(
            "Поняла, не добавляю. Если это клиент из таблицы, напиши его имя так, как оно записано там.")
        return
    await q.answer("Добавляю…")
    _apply_sheet_override(ctx)
    try:
        text = await asyncio.to_thread(confirm_new_lead, cfg, store, int(pid), _today(cfg))
    except Exception:
        log.exception("Не удалось добавить лида")
        text = "Не получилось добавить в таблицу, попробуй позже."
    await q.message.reply_text(text)


async def cmd_bindtopic(update: Update, ctx: ContextTypes.DEFAULT_TYPE) -> None:
    if not _is_manager(ctx, update):
        return await _deny(update)
    msg = update.effective_message
    arg = (ctx.args[0].lower() if ctx.args else "")
    kind = {"план": "plan", "plan": "plan", "факт": "fact", "fact": "fact"}.get(arg)
    if not kind or not msg.is_topic_message:
        await msg.reply_text("Отправь эту команду внутри темы форума: /bindtopic план или /bindtopic факт")
        return
    _store(ctx).set_setting(f"topic:{msg.chat_id}:{msg.message_thread_id}", kind)
    if not _partners_chat(ctx):
        _store(ctx).set_setting("partners_chat_id", str(msg.chat_id))
    await msg.reply_text(f"✅ Эта тема — {'план' if kind == 'plan' else 'факт'}. Читаю отсюда сообщения партнёров.")


async def cmd_setkey(update: Update, ctx: ContextTypes.DEFAULT_TYPE) -> None:
    msg = update.effective_message
    if update.effective_chat.type != "private":
        try:
            await msg.delete()
        except Exception:
            pass
        await ctx.bot.send_message(msg.chat_id, "Ключ можно присылать только в личку боту — сообщение удалила.")
        return
    if not _is_manager(ctx, update):
        return await _deny(update)
    key = (ctx.args[0] if ctx.args else "").strip()
    try:
        await msg.delete()  # не держим ключ в истории чата
    except Exception:
        pass
    if not key.startswith("sk-ant-"):
        await ctx.bot.send_message(msg.chat_id, "Пришли так: /setkey sk-ant-… (ключ из console.anthropic.com)")
        return
    _store(ctx).set_setting("anthropic_key", key)
    ctx.application.bot_data.pop("parser", None)
    await ctx.bot.send_message(msg.chat_id, "✅ Ключ сохранила, сообщение с ним удалила. "
                                             "Теперь разбираю план и факт партнёров через ИИ.")


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
    app.add_handler(CommandHandler("setsheet", cmd_setsheet))
    app.add_handler(CommandHandler("bindtopic", cmd_bindtopic))
    app.add_handler(CommandHandler("setkey", cmd_setkey))
    app.add_handler(CallbackQueryHandler(on_newlead, pattern=r"^newlead:"))
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
