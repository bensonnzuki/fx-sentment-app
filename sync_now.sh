#!/usr/bin/env bash
cd /home/bnzuki/fx-sentment-app || exit 1
git pull --rebase -q origin main
/usr/bin/python3 scraper.py || exit 1
CHANGED=$(/usr/bin/python3 -c 'import json; print(json.load(open("docs/latest.json")).get("any_changed", True))')
if [ "$CHANGED" = "True" ]; then
  git add docs/latest.json
  git diff --quiet && git diff --staged --quiet || (git commit -m "Update 15m data" -q && git push -q origin main)
else
  echo "No sentiment or 15M RSI change — skipping git push"
fi
