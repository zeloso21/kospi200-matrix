"""KOSPI200 매트릭스 계산 (2026-10-19 사용자 요청).

C:\\연습\\KODEX200_트레이딩\\kospi200matrix_*.xlsx(하나증권 HTS DDE로 실시간 값을 받던 엑셀)의
D/W/월간/년간 Matrix와 '요약' 시트 계산을 한국투자증권 Open API(모의서버 업종 기간별시세
FHKUP03500100, 업종코드 2001=KOSPI200) 데이터로 옮긴 것. VM cron이 장중 30분마다 실행해 data.json을 만들고
GitHub Pages(zeloso21/kospi200-matrix)가 그걸 읽어 요약 화면을 그림.

기간은 엑셀과 같게 옵션 만기 주기(2026-10-19 사용자 요청): 주간 = 금요일~목요일(주간옵션 목요일 만기),
월간 = 둘째 목요일(월물 만기) 다음 날 ~ 다음 달 둘째 목요일. 만기일이 휴장이면 그 전 거래일이 만기지만, 그 사이는 전부 휴장이라
'목요일 기준 묶기'로 같은 결과. 주·월 봉은 일봉을 묶어 코드가 만듦(엑셀은 일부 시가를 손으로 넣어 시가가 다를 수 있음). 연간은 달력 연도.
추세 판정도 엑셀은 손으로 적은 SELL/BUY라, 여기선 '현재가가 중심선·5기간 평균 위/아래'로 코드가 판정.
"""
import datetime
import json
import os
import sys

sys.path.insert(0, "/home/user/kis_agent")
import daily_flow_cafe as df  # noqa: E402  (KIS 인증·조회 재사용)

OUT = os.path.join(os.path.dirname(os.path.abspath(__file__)), "data.json")
BACK = {"Y": 8000}
NAMES = {"D": "일간", "W": "주간(금~목, 주간옵션 만기)", "M": "월간(월물 만기 주기)", "Y": "연간"}


def bars(kis, per, today):
    j = kis.get("/uapi/domestic-stock/v1/quotations/inquire-daily-indexchartprice", "FHKUP03500100",
                {"FID_COND_MRKT_DIV_CODE": "U", "FID_INPUT_ISCD": "2001",
                 "FID_INPUT_DATE_1": (today - datetime.timedelta(days=BACK[per])).strftime("%Y%m%d"),
                 "FID_INPUT_DATE_2": today.strftime("%Y%m%d"), "FID_PERIOD_DIV_CODE": per})
    rows = [r for r in j.get("output2") or [] if r.get("stck_bsop_date") and float(r.get("bstp_nmix_oprc") or 0) > 0]
    out = [{"date": r["stck_bsop_date"], "o": float(r["bstp_nmix_oprc"]), "h": float(r["bstp_nmix_hgpr"]),
            "l": float(r["bstp_nmix_lwpr"]), "c": float(r["bstp_nmix_prpr"])} for r in rows]
    return sorted(out, key=lambda b: b["date"])


def daily_long(kis, today, days=430):
    """월간 10기간+기준 1기간을 덮는 일봉. 한 번에 받는 행 수에 한도가 있어 60일씩 나눠 받음."""
    got, end = {}, today
    start = today - datetime.timedelta(days=days)
    while end > start:
        s = max(start, end - datetime.timedelta(days=60))
        j = kis.get("/uapi/domestic-stock/v1/quotations/inquire-daily-indexchartprice", "FHKUP03500100",
                    {"FID_COND_MRKT_DIV_CODE": "U", "FID_INPUT_ISCD": "2001", "FID_INPUT_DATE_1": s.strftime("%Y%m%d"),
                     "FID_INPUT_DATE_2": end.strftime("%Y%m%d"), "FID_PERIOD_DIV_CODE": "D"})
        for r in j.get("output2") or []:
            if r.get("stck_bsop_date") and float(r.get("bstp_nmix_oprc") or 0) > 0:
                got[r["stck_bsop_date"]] = {"date": r["stck_bsop_date"], "o": float(r["bstp_nmix_oprc"]), "h": float(r["bstp_nmix_hgpr"]),
                                            "l": float(r["bstp_nmix_lwpr"]), "c": float(r["bstp_nmix_prpr"])}
        end = s - datetime.timedelta(days=1)
    return [got[k] for k in sorted(got)]


def second_thu(y, m):
    d = datetime.date(y, m, 1)
    return d + datetime.timedelta(days=(3 - d.weekday()) % 7 + 7)


def expiry_key(day, per):
    """그 거래일이 속한 만기 주기의 (명목) 만기일."""
    d = datetime.datetime.strptime(day, "%Y%m%d").date()
    if per == "W":  # 월(0)~목(3)은 이번 주 목요일, 금(4)은 다음 주 목요일
        return d + datetime.timedelta(days=(3 - d.weekday()) % 7)
    e = second_thu(d.year, d.month)
    if d > e:
        y, m = (d.year + 1, 1) if d.month == 12 else (d.year, d.month + 1)
        e = second_thu(y, m)
    return e


def expiry_bars(daily, per):
    groups = {}
    for b in daily:
        groups.setdefault(expiry_key(b["date"], per), []).append(b)
    out = []
    for k in sorted(groups):
        g = groups[k]
        out.append({"date": g[0]["date"], "end": g[-1]["date"], "expiry": k.strftime("%Y%m%d"), "days": len(g),
                    "o": g[0]["o"], "h": max(b["h"] for b in g), "l": min(b["l"] for b in g), "c": g[-1]["c"]})
    return out[1:]  # 맨 앞 주기는 받은 일봉 범위에 잘려 있을 수 있어 버림


def avg(xs):
    return sum(xs) / len(xs) if xs else 0.0


def matrix(bs, base, price):
    """엑셀 Matrix 시트 계산. base=피벗 기준 봉(직전 기간), cur=진행 중인 기간 봉."""
    cur = bs[-1]
    p = (base["h"] + base["l"] + base["c"]) / 3
    r1, s1 = 2 * p - base["l"], 2 * p - base["h"]
    r2, s2 = p + (r1 - s1), p - (r1 - s1)
    mid = (cur["h"] + cur["l"]) / 2          # 이번 기간 고가·저가 중간 = 중심선
    k = mid - p                               # 피벗을 이번 기간 중심으로 옮기는 폭
    best_re, high1, best_sup, low1 = r2 + k, r1 + k, s2 + k, s1 + k
    levels = [("best/re", "최대 저항", best_re), ("high/2", "", (best_re + high1) / 2), ("high/1", "", high1),
              ("resistance", "저항", (best_re + mid) / 2), ("middle", "중심", mid),
              ("support", "지지", (mid + best_sup) / 2), ("low/1", "", low1),
              ("low/2", "", (low1 + best_sup) / 2), ("best/sup", "최대 지지", best_sup)]
    last10 = bs[-10:]
    oh = avg([b["h"] - b["o"] for b in last10])
    ol = avg([b["o"] - b["l"] for b in last10])
    hl = avg([b["h"] - b["l"] for b in last10])
    e_high = (best_re + (cur["o"] + oh) + (mid + hl / 2)) / 3
    e_low = (best_sup + (cur["o"] - ol) + (mid - hl / 2)) / 3
    ma = {}
    for n in (3, 5, 10):
        seg = bs[-n:]
        ma[n] = {f: avg([b[f] for b in seg]) for f in ("h", "c", "l")}
    ma_rows = {f: {n: {"v": ma[n][f], "r": (price - ma[n][f]) / price * 100} for n in (3, 5, 10)} for f in ("h", "c", "l")}
    above_mid, above_ma5 = price >= mid, price >= ma[5]["c"]
    trend = "강세" if above_mid and above_ma5 else ("약세" if not above_mid and not above_ma5 else "중립")
    return {
        "base": base, "cur": cur, "pivot": p, "start_line": p - (base["c"] - cur["o"]),
        "levels": [{"key": a, "label": b, "v": v} for a, b, v in levels],
        "avg_move": {"open_high": oh, "open_low": ol, "high_low": hl, "n": len(last10)},
        "e_high": e_high, "e_low": e_low, "open_high": cur["o"] + oh, "open_low": cur["o"] - ol,
        "ma": ma_rows, "ma_up": max(ma[n]["h"] for n in (3, 5, 10)), "ma_dn": min(ma[n]["l"] for n in (3, 5, 10)),
        "range_now": (cur["h"] - cur["l"]) / cur["o"], "range_avg10": avg([(b["h"] - b["l"]) / b["o"] for b in last10]),
        "trend": trend, "above_mid": above_mid, "above_ma5": above_ma5,
        "recent": list(reversed(bs[-10:])),
    }


def main():
    now = datetime.datetime.now()
    kis = df.Kis()
    data = {"updated": now.strftime("%Y-%m-%d %H:%M"), "periods": {}}
    daily = daily_long(kis, now.date())
    allb = {"D": daily[-60:], "W": expiry_bars(daily, "W"), "M": expiry_bars(daily, "M"), "Y": bars(kis, "Y", now.date())}
    d = allb["D"]
    price = d[-1]["c"]
    closed = now.time() >= datetime.time(15, 35) or now.weekday() >= 5 or d[-1]["date"] != now.strftime("%Y%m%d")
    for per, bs in allb.items():
        if len(bs) < 11:
            raise SystemExit(f"{per} 데이터 부족 {len(bs)}")
        # 일간은 엑셀 D Matrix처럼 늘 오늘(진행 중) 봉 기준 — 피벗=(오늘 고가+저가+현재가)/3라 장중엔 계속 바뀌고 마감 값이 다음 날용.
        # 2026-10-20 사용자 요청으로 장중 '어제 봉' 고정 방식에서 엑셀과 똑같이 변경. 주·월·연은 엑셀도 직전 기간 봉
        base = bs[-1] if per == "D" else bs[-2]
        m = matrix(bs, base, price)
        m["name"] = NAMES[per]
        data["periods"][per] = m
    data["price"] = price
    data["day"] = d[-1]
    data["prev_close"] = d[-2]["c"]
    data["market_closed"] = closed
    y = allb["Y"]
    data["year"] = {"range_now": (y[-1]["h"] - y[-1]["l"]) / y[-1]["o"],
                    "range_avg10": avg([b["h"] - b["l"] for b in y[-10:]]) / avg([b["o"] for b in y[-10:]])}
    tmp = OUT + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=1)
    os.replace(tmp, OUT)
    print(f"{now:%F %T} 갱신 price={price} closed={closed}")


if __name__ == "__main__":
    main()
