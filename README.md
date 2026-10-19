# KOSPI200 매트릭스

엑셀 `kospi200matrix`(하나증권 HTS DDE 연동)의 일간·주간·월간·연간 Matrix와 요약 시트를
한국투자증권 Open API 데이터로 옮긴 페이지. https://zeloso21.github.io/kospi200-matrix/

- `compute.py` — KIS 업종 기간별시세(KOSPI200)로 기준선·예상 고저점·이평 매트릭스 계산 → `data.json`
- `run_update.sh` — GCP VM cron(평일 09:05~15:35 30분마다, 15:45 마감분)이 실행해 커밋·푸시
- `index.html` — data.json을 읽어 요약 화면 표시
