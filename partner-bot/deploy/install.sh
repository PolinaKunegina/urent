#!/usr/bin/env bash
# Установка бота на дроплет DigitalOcean одной командой.
# Откройте дроплет в панели DigitalOcean → Access → Launch Droplet Console и вставьте:
#
#   curl -fsSL https://raw.githubusercontent.com/PolinaKunegina/urent/main/partner-bot/deploy/install.sh | bash
#
# Скрипт спросит токен бота, ссылку на таблицу, ники руководителей и JSON-ключ Google,
# поднимет бота в Docker и включит автообновление: каждые 5 минут сервер проверяет
# GitHub и сам перезапускает бота, если код изменился.
# Повторный запуск безопасен: уже введённые настройки не спрашиваются.
set -euo pipefail

BRANCH="${BRANCH:-main}"
APP=/opt/urent
BOT=$APP/partner-bot

ask() { local v; read -r -p "$1" v </dev/tty; printf '%s' "$v"; }
say() { printf '\n\033[1;32m▶ %s\033[0m\n' "$1"; }

[ "$(id -u)" = 0 ] || { echo "Запустите от root"; exit 1; }

say "Ставлю Docker и git"
command -v git >/dev/null || { apt-get update -qq && apt-get install -y -qq git; }
command -v docker >/dev/null || curl -fsSL https://get.docker.com | sh
if ! docker compose version >/dev/null 2>&1; then
  # Docker из репозитория Ubuntu ставится без compose — доставляем модуль
  apt-get update -qq
  apt-get install -y -qq docker-compose-plugin || apt-get install -y -qq docker-compose-v2
fi

say "Забираю код (ветка $BRANCH)"
if [ -d "$APP/.git" ]; then
  git -C "$APP" fetch -q origin "$BRANCH" && git -C "$APP" checkout -q -B "$BRANCH" "origin/$BRANCH"
else
  git clone -q -b "$BRANCH" https://github.com/PolinaKunegina/urent.git "$APP"
fi
cd "$BOT"
mkdir -p data secrets && chmod 700 secrets

if [ ! -f .env ]; then
  say "Настройки"
  token=$(ask "Токен бота от @BotFather: ")
  link=$(ask "Ссылка на таблицу воронки: ")
  sheet_id=$(printf '%s' "$link" | sed -E 's#.*/d/([A-Za-z0-9_-]+).*#\1#')
  managers=$(ask "Ники руководителей через запятую (например @nat, @kir, @anton): ")
  cp .env.example .env
  sed -i "s|^TELEGRAM_BOT_TOKEN=.*|TELEGRAM_BOT_TOKEN=$token|" .env
  sed -i "s|^SHEET_ID=.*|SHEET_ID=$sheet_id|" .env
  sed -i "s|^MANAGERS=.*|MANAGERS=$managers|" .env
  chmod 600 .env
fi

if [ ! -s secrets/service-account.json ]; then
  say "Ключ Google"
  echo "Вставьте содержимое скачанного JSON-ключа целиком, затем на новой строке нажмите Ctrl+D:"
  cat </dev/tty > secrets/service-account.json
  python3 -m json.tool secrets/service-account.json >/dev/null \
    || { echo "Это не похоже на JSON, удаляю. Запустите скрипт ещё раз."; rm -f secrets/service-account.json; exit 1; }
  chmod 600 secrets/service-account.json
fi

say "Включаю автообновление с GitHub (каждые 5 минут)"
cat > /etc/systemd/system/urent-bot-update.service <<UNIT
[Unit]
Description=Обновление бота партнёрской сети из GitHub
[Service]
Type=oneshot
ExecStart=/bin/bash $BOT/deploy/update.sh
UNIT
cat > /etc/systemd/system/urent-bot-update.timer <<UNIT
[Unit]
Description=Проверка обновлений бота каждые 5 минут
[Timer]
OnBootSec=1min
OnUnitActiveSec=5min
[Install]
WantedBy=timers.target
UNIT
systemctl daemon-reload
systemctl enable --now urent-bot-update.timer >/dev/null

say "Запускаю бота"
docker compose up -d --build

email=$(python3 -c "import json;print(json.load(open('secrets/service-account.json'))['client_email'])" 2>/dev/null || true)
cat <<DONE

✅ Бот запущен. Осталось:
  1. Дать доступ «Редактор» к таблице адресу:  ${email:-<client_email из JSON-ключа>}
  2. Руководителям и партнёрам написать боту /start в личку.
  3. Добавить бота в чат партнёров и отправить там /bindchat (от руководителя).
  4. Проверить всё командой /status в личке с ботом.

Логи:        cd $BOT && docker compose logs -f
Настройки:   nano $BOT/.env   (после правки: docker compose up -d)
DONE
