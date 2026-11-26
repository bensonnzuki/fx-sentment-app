#!/usr/bin/env bash
# Run every 15 minutes from cron. scraper.py only rewrites docs/latest.json when the
# data changed or 25+ minutes have passed, so GitHub Pages stays under its build limit.
cd /home/bnzuki/fx-sentment-app || exit 1

# Never let two runs overlap.
exec 9>/tmp/fx-sentment-sync.lock
flock -n 9 || { echo "$(date '+%F %T') previous run still active, skipping"; exit 0; }

echo "=== $(date '+%F %T') ==="
git pull --rebase --autostash -q origin main || echo "git pull failed, continuing with local copy"
/usr/bin/python3 scraper.py || exit 1

git add docs/latest.json
if git diff --staged --quiet; then
  echo "Nothing new to publish"
else
  git commit -q -m "Data update $(date '+%H:%M')" && git push -q origin main && echo "Published"
fi
