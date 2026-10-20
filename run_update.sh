#!/bin/bash
# KOSPI200 매트릭스 data.json 갱신 (VM cron: 평일 장중 5분마다 + 장 마감 뒤)
cd /home/user/kospi200-matrix || exit 1
git pull -q --rebase --autostash 2>&1
/home/user/kis_agent/.venv/bin/python3 compute.py || exit 1
git add data.json
if git diff --staged --quiet; then echo "$(date) 변경 없음"; else
  git commit -q -m "data: KOSPI200 매트릭스 $(date +%F_%H:%M)"
  git push -q && echo "$(date) 푸시 완료"
fi
