#!/usr/bin/env python3
"""保有株の売りライン監視 (Chandelier Exit: 取得後の最高終値 - N x ATR)。

データ: Tiingo EOD API。APIキーはクラウド環境のプロキシが自動付与するため、
コード内では Authorization にダミー値を入れている(プロキシが上書きする)。

使い方:
  python3 monitor.py --mode daily     # 変化があった日だけ通知文を出す(なければ NO_ALERT)
  python3 monitor.py --mode weekly    # 全銘柄の状況まとめ
  python3 monitor.py --mode weekly --asof 2026-08-24   # 過去日付で検証
"""
import argparse
import datetime as dt
import json
import os
import sys
import urllib.request

HERE = os.path.dirname(os.path.abspath(__file__))
API = "https://api.tiingo.com/tiingo/daily/{sym}/prices?startDate={start}&endDate={end}"


def fetch(sym, start, end):
    req = urllib.request.Request(
        API.format(sym=sym.lower(), start=start, end=end),
        headers={"Authorization": "Token " + os.environ.get("TIINGO_TOKEN", "proxy-injected"),
                 "Content-Type": "application/json"})
    with urllib.request.urlopen(req, timeout=30) as r:
        rows = json.load(r)
    return [{"date": x["date"][:10], "open": x["adjOpen"], "high": x["adjHigh"],
             "low": x["adjLow"], "close": x["adjClose"]} for x in rows]


def atr_series(rows, n):
    tr = [None]
    for a, b in zip(rows, rows[1:]):
        tr.append(max(b["high"] - b["low"], abs(b["high"] - a["close"]), abs(b["low"] - a["close"])))
    out = []
    for i in range(len(rows)):
        w = [t for t in tr[max(1, i - n + 1):i + 1] if t is not None]
        out.append(sum(w) / len(w) if len(w) >= n else None)
    return out


def sma(vals, n, i):
    return sum(vals[i - n + 1:i + 1]) / n if i + 1 >= n else None


def expected_session(holidays):
    """今の時点で終値が出ているはずの直近の米国取引日 (ET 16:00 以降なら当日)。"""
    try:
        from zoneinfo import ZoneInfo
        now = dt.datetime.now(ZoneInfo("America/New_York"))
    except Exception:  # noqa: BLE001  tzdataが無い環境では冬時間(UTC-5)で代用
        now = dt.datetime.now(dt.timezone(dt.timedelta(hours=-5)))
    d = now.date() if now.hour >= 16 else now.date() - dt.timedelta(days=1)
    while d.weekday() >= 5 or d.isoformat() in holidays:
        d -= dt.timedelta(days=1)
    return d.isoformat()


def analyze_holding(h, rows, cfg):
    """取得日以降について、ラチェット式(下げない)売りラインを日毎に計算する。"""
    atr = atr_series(rows, cfg["atr_period"])
    m = cfg["atr_mult"]
    peak = None
    stop = None
    hist = []  # (date, close, peak, stop)
    for i, r in enumerate(rows):
        if r["date"] < h["entry"] or atr[i] is None:
            continue
        # 初日は買値を起点にする(最初の売りライン = 買値 - N x ATR)
        peak = max(h["cost"], r["close"]) if peak is None else max(peak, r["close"])
        cand = peak - m * atr[i]
        stop = cand if stop is None else max(stop, cand)
        hist.append((r["date"], r["close"], peak, stop, atr[i]))
    return hist


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--mode", choices=["daily", "weekly"], default="daily")
    ap.add_argument("--asof", help="YYYY-MM-DD (検証用)")
    ap.add_argument("--config", default=os.path.join(HERE, "config.json"))
    a = ap.parse_args()
    cfg = json.load(open(a.config, encoding="utf-8"))
    end = a.asof or dt.date.today().isoformat()
    end_d = dt.date.fromisoformat(end)
    expected = None if a.asof else expected_session(set(cfg.get("market_holidays", [])))

    alerts, lines, errors = [], [], []
    total_cost = total_val = 0.0
    vals = {}

    for h in cfg["holdings"]:
        sym = h["symbol"]
        if h["entry"] > end:  # 検証日時点でまだ買っていない
            continue
        start = (dt.date.fromisoformat(h["entry"]) - dt.timedelta(days=60)).isoformat()
        try:
            rows = fetch(sym, start, end)
        except Exception as e:  # noqa: BLE001
            errors.append(f"{sym}: データ取得失敗 ({e})")
            continue
        hist = analyze_holding(h, rows, cfg)
        if not hist:
            if not rows or rows[-1]["date"] < h["entry"]:
                lines.append(f"{sym}: 購入日({h['entry']})の終値待ち")
            else:
                errors.append(f"{sym}: データ不足")
            continue
        d, close, peak, stop, atr = hist[-1]
        first_day = len(hist) == 1  # 購入初日: 前日との比較ができない
        if first_day:
            pclose, pstop = h["cost"], stop
        else:
            pd_, pclose, _, pstop, _ = hist[-2]
        if expected and d < expected:
            errors.append(f"{sym}: データ未更新 (最新{d}、本来は{expected}の終値があるはず。判定は{d}時点)")
        elif not expected and (end_d - dt.date.fromisoformat(d)).days > 4:
            errors.append(f"{sym}: 最新データが古い ({d})")
        pnl = (close / h["cost"] - 1) * 100
        val = close * h["shares"]
        total_cost += h["cost"] * h["shares"]
        total_val += val
        vals[sym] = val
        stop_pct = (stop / close - 1) * 100
        status = "割れ" if close < stop else "OK"
        lines.append(f"{sym}: ${close:.2f} ({pnl:+.0f}%) 売りライン${stop:.2f} [{status}] 最高値${peak:.2f}")

        # ① 新たに売りラインを割った
        if close < stop and pclose >= pstop:
            alerts.append(f"🔴{sym} 売りライン割れ: 終値${close:.2f} < ${stop:.2f}。ルール上は売り")
        # ② 売りラインが大きく上がった(5営業日前比)
        if len(hist) >= 6:
            s5 = hist[-6][3]
            if stop >= s5 * (1 + cfg["stop_raise_alert_pct"] / 100) and close >= stop:
                alerts.append(f"⬆️{sym} 売りライン上昇 ${s5:.2f}→${stop:.2f}。SBIの逆指値を${stop:.2f}に変更を")
        # ③ 含み損→プラス転換
        if first_day:
            alerts.append(f"🆕{sym} 購入初日の売りライン ${stop:.2f}。SBIで逆指値をこの値段に")
        elif close > h["cost"] and pclose <= h["cost"]:
            alerts.append(f"🟢{sym} プラス転換: ${close:.2f} > 取得${h['cost']:.2f}")
        # ④ 決算接近
        ed = cfg.get("earnings", {}).get(sym)
        if ed:
            days = (dt.date.fromisoformat(ed) - end_d).days
            if 0 <= days <= cfg["earnings_warn_days"]:
                alerts.append(f"📅{sym} 決算まで{days}日 ({ed})。値が飛ぶ可能性")

    # 売却済み銘柄: 再エントリー候補(50日線を上抜け)
    for s in cfg.get("sold", []):
        sym = s["symbol"]
        try:
            rows = fetch(sym, (end_d - dt.timedelta(days=120)).isoformat(), end)
        except Exception as e:  # noqa: BLE001
            errors.append(f"{sym}(売却済): データ取得失敗 ({e})")
            continue
        c = [r["close"] for r in rows]
        n = cfg["reentry_ma_days"]
        if len(c) > n:
            m1, m0 = sma(c, n, len(c) - 1), sma(c, n, len(c) - 2)
            if c[-1] > m1 and c[-2] <= m0:
                alerts.append(f"🔁{sym}(売却済) {n}日線を上抜け ${c[-1]:.2f}。再購入を検討する場面")
            lines.append(f"{sym}(売却済 {s.get('date','')} @${s.get('price','?')}): ${c[-1]:.2f} {n}日線${m1:.2f}")

    # 市場全体: SPYと200日線
    # 見張り候補: 買いのタイミング条件(方針v2の4-1)を新たに満たした日に通知
    for w in cfg.get("watchlist", []):
        sym = w["symbol"]
        try:
            rows = fetch(sym, (end_d - dt.timedelta(days=150)).isoformat(), end)
        except Exception as e:  # noqa: BLE001
            errors.append(f"{sym}(見張り): データ取得失敗 ({e})")
            continue
        c = [r["close"] for r in rows]
        if len(c) < 63:
            errors.append(f"{sym}(見張り): データ不足")
            continue

        def entry_ok(i):
            ma = sma(c, 50, i)
            ma10 = sma(c, 50, i - 10)
            r42 = c[i] / c[i - 42] - 1
            return c[i] > ma and ma >= ma10 * 0.99 and r42 > -0.10, ma, r42

        ok, ma, r42 = entry_ok(len(c) - 1)
        ok_prev, _, _ = entry_ok(len(c) - 2)
        ed = cfg.get("earnings", {}).get(sym)
        near_earn = False
        if ed:
            bd = sum(1 for k in range((dt.date.fromisoformat(ed) - end_d).days + 1)
                     if (end_d + dt.timedelta(days=k)).weekday() < 5)
            near_earn = 0 <= (dt.date.fromisoformat(ed) - end_d).days and bd <= 5
        state = "条件クリア" if ok and not near_earn else ("決算前で待機" if ok else "未達")
        lines.append(f"{sym}(見張り): ${c[-1]:.2f} 50日線${ma:.2f} 2ヶ月{r42*100:+.0f}% → {state}")
        if ok and not ok_prev and not near_earn:
            alerts.append(f"👀{sym}(見張り) 買いのタイミング条件を満たした ${c[-1]:.2f}。"
                          f"4つの絶対条件(決算の中身)は自分で確認を")

    mk = cfg["market_symbol"]
    try:
        rows = fetch(mk, (end_d - dt.timedelta(days=330)).isoformat(), end)
        if expected and rows[-1]["date"] < expected:
            errors.append(f"{mk}: データ未更新 (最新{rows[-1]['date']}、本来は{expected})")
        c = [r["close"] for r in rows]
        n = cfg["market_ma_days"]
        m1, m0 = sma(c, n, len(c) - 1), sma(c, n, len(c) - 2)
        side = "上" if c[-1] > m1 else "下"
        market_line = f"{mk} ${c[-1]:.2f} / {n}日線${m1:.2f} → 線の{side}"
        if (c[-1] > m1) != (c[-2] > m0):
            alerts.append(f"🌐{mk}が{n}日線を{'上' if c[-1] > m1 else '下'}抜け。相場全体の流れが変化")
    except Exception as e:  # noqa: BLE001
        market_line = f"{mk}: 取得失敗 ({e})"
        errors.append(market_line)

    # 集中度
    weight_lines = []
    if total_val:
        for sym, v in sorted(vals.items(), key=lambda x: -x[1]):
            w = v / total_val * 100
            if w > cfg["weight_limit_pct"]:
                weight_lines.append(f"{sym} {w:.0f}%")

    if a.mode == "daily":
        out = alerts + [f"⚠️{e}" for e in errors]
        print("\n".join(out) if out else "NO_ALERT")
        return
    pnl_total = total_val - total_cost
    if total_cost:
        print(f"【週次まとめ {end}】評価${total_val:,.0f} 損益${pnl_total:+,.0f} ({(total_val/total_cost-1)*100:+.1f}%)")
    else:
        print(f"【週次まとめ {end}】評価額は取得できず")
    print(*lines, sep="\n")
    print(market_line)
    if weight_lines:
        print(f"集中超過(>{cfg['weight_limit_pct']:.0f}%): " + ", ".join(weight_lines))
    if alerts:
        print("今週の注意:", *alerts, sep="\n")
    if errors:
        print("⚠️仕組みの異常:", *errors, sep="\n")
    else:
        print("✅データ取得正常")


if __name__ == "__main__":
    sys.exit(main())
