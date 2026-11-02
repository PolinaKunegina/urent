#!/usr/bin/env bash
# Автообновление: если в GitHub появился новый код, подтягивает его и пересобирает бота.
# Запускается таймером systemd каждые 5 минут; вручную: FORCE=1 bash update.sh
set -euo pipefail
cd /opt/urent
BRANCH=$(git rev-parse --abbrev-ref HEAD)
git fetch -q origin "$BRANCH"
if [ "$(git rev-parse HEAD)" != "$(git rev-parse "origin/$BRANCH")" ] || [ "${FORCE:-}" = 1 ]; then
  git reset -q --hard "origin/$BRANCH"     # .env, data/ и secrets/ не в git — их не трогает
  cd partner-bot
  docker compose up -d --build
  docker image prune -f >/dev/null
  echo "Бот обновлён до $(git rev-parse --short HEAD)"
else
  cd partner-bot && docker compose up -d >/dev/null   # на случай, если контейнер остановился
fi
