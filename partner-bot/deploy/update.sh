#!/usr/bin/env bash
# Обновить бота на сервере до свежей версии из GitHub
set -euo pipefail
cd /opt/urent
git pull
cd partner-bot
docker compose up -d --build
docker compose ps
