#!/usr/bin/env bash
# Первичная настройка дроплета DigitalOcean (Ubuntu 22.04/24.04). Запускать от root:
#   bash setup_droplet.sh
# Другая ветка: BRANCH=имя-ветки bash setup_droplet.sh
set -euo pipefail

if ! command -v docker >/dev/null; then
  curl -fsSL https://get.docker.com | sh
fi

APP_DIR=/opt/urent
BRANCH="${BRANCH:-main}"
if [ ! -d "$APP_DIR/.git" ]; then
  git clone -b "$BRANCH" https://github.com/PolinaKunegina/urent.git "$APP_DIR"
fi
cd "$APP_DIR/partner-bot"
mkdir -p data secrets
[ -f .env ] || cp .env.example .env

echo
echo "Дальше:"
echo "  1) nano $APP_DIR/partner-bot/.env                       — токен бота, ID таблиц, получатели"
echo "  2) положите ключ Google в $APP_DIR/partner-bot/secrets/service-account.json"
echo "  3) cd $APP_DIR/partner-bot && docker compose up -d --build"
echo "  4) docker compose logs -f                               — проверить, что бот стартовал"
