#!/usr/bin/env bash
# Автообновление: если в GitHub появился новый код, подтягивает его и пересобирает бота.
# Запускается таймером systemd каждые 5 минут; вручную: FORCE=1 bash update.sh
#
# Репозиторий может быть приватным: сервер ходит в GitHub по своему ключу (deploy key).
# Ключ создаётся здесь один раз, его публичную часть бот присылает по команде /deploykey —
# её нужно один раз добавить в GitHub: репозиторий → Settings → Deploy keys → Add deploy key.
# Пока ключ не добавлен, а репозиторий публичный, обновления идут по-старому, по https.
set -euo pipefail
cd /opt/urent
DATA=/opt/urent/partner-bot/data
KEY=/root/.ssh/urent_deploy
REPO=PolinaKunegina/urent

mkdir -p "$DATA"
if [ ! -f "$KEY" ] && command -v ssh-keygen >/dev/null; then
  mkdir -p /root/.ssh && chmod 700 /root/.ssh
  ssh-keygen -q -t ed25519 -N "" -C "urent-bot-server" -f "$KEY" || true
fi
[ -f "$KEY.pub" ] && cp "$KEY.pub" "$DATA/deploy_key.pub" || true
export GIT_SSH_COMMAND="ssh -i $KEY -o IdentitiesOnly=yes -o StrictHostKeyChecking=accept-new -o BatchMode=yes -o ConnectTimeout=15"

BRANCH=$(git rev-parse --abbrev-ref HEAD)
REFSPEC="+refs/heads/$BRANCH:refs/remotes/origin/$BRANCH"
via=""
for url in "git@github.com:$REPO.git" "ssh://git@ssh.github.com:443/$REPO.git" "https://github.com/$REPO.git"; do
  if git fetch -q "$url" "$REFSPEC" 2>/dev/null; then
    case "$url" in https*) via="https" ;; *) via="ssh" ;; esac
    break
  fi
done
echo "${via:-fail} $(date '+%d.%m.%Y %H:%M')" > "$DATA/deploy_status"
[ -n "$via" ] || { echo "Не удалось получить обновления из GitHub"; exit 1; }

if [ "$(git rev-parse HEAD)" != "$(git rev-parse "origin/$BRANCH")" ] || [ "${FORCE:-}" = 1 ]; then
  git reset -q --hard "origin/$BRANCH"     # .env, data/ и secrets/ не в git — их не трогает
  cd partner-bot
  docker compose up -d --build
  docker image prune -f >/dev/null
  echo "Бот обновлён до $(git rev-parse --short HEAD)"
else
  cd partner-bot && docker compose up -d >/dev/null   # на случай, если контейнер остановился
fi
