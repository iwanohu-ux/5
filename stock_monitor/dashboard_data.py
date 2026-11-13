#!/usr/bin/env python3
"""ダッシュボード用のデータ(保有・見張り・相場)を1つのJSONにまとめる。

使い方:
  python3 dashboard_data.py --out dashboard.json
各銘柄の値動き(1日/1週/1月/3月)、50日線との位置と傾き、流れの判定、
ルール上の状態(売りライン・買いタイミング条件)、注意点を出す。
"""
import argparse
import datetime as dt
import json
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import monitor as M  # noqa: E402


def pct(a, b):
    return round((a / b - 1) * 100, 1) if b else None


def trend_of(close, ma50, slope):
    if close > ma50 and slope >= 0.5:
        return "上向き"
    if close < ma50 and slope <= -0.5:
        return "下向き"
    return "横ばい"


def base_stats(rows):
    c = [r["close"] for r in rows]
    ma50 = sum(c[-50:]) / 50
    ma50_prev = sum(c[-60:-10]) / 50
    slope_raw = (ma50 / ma50_prev - 1) * 100
    return {"_slope_raw": slope_raw,
        "date": rows[-1]["date"], "close": round(c[-1], 2),
        "d1": pct(c[-1], c[-2]), "w1": pct(c[-1], c[-6]), "m1": pct(c[-1], c[-22]),
        "m2": pct(c[-1], c[-43]), "m3": pct(c[-1], c[-64]),
        "ma50": round(ma50, 2), "ma50_gap": pct(c[-1], ma50), "ma50_slope": pct(ma50, ma50_prev),
        "from_high": pct(c[-1], max(c[-252:])),  # 直近1年の高値から
        "trend": trend_of(c[-1], ma50, slope_raw),
    }


HIST_DAYS = 252  # チャート用に持つ営業日数(約1年)


def hist_of(rows, stop_hist=None):
    """チャート用: 日付・終値・50日線(・売りライン)を直近 HIST_DAYS 日分。"""
    c = [r["close"] for r in rows]
    n = len(rows)
    i0 = max(0, n - HIST_DAYS)
    out = {"h_d": [r["date"] for r in rows[i0:]],
           "h_c": [round(x, 2) for x in c[i0:]],
           "h_m": [round(sum(c[i - 49:i + 1]) / 50, 2) if i >= 49 else None for i in range(i0, n)]}
    if stop_hist:
        st = {d: v for d, v in stop_hist}
        out["h_s"] = [round(st[d], 2) if d in st else None for d in out["h_d"]]
    return out


def cautions(s, ed, today):
    out = []
    if s["w1"] is not None and s["w1"] >= 20:
        out.append(f"1週間で+{s['w1']:.0f}%と急騰。高値づかみに注意")
    if s["w1"] is not None and s["w1"] <= -15:
        out.append(f"1週間で{s['w1']:.0f}%と急落")
    if s["d1"] is not None and s["d1"] <= -8:
        out.append(f"前日比{s['d1']:.0f}%の急落")
    if ed:
        days = (dt.date.fromisoformat(ed) - today).days
        if 0 <= days <= 7:
            out.append(f"決算まで{days}日({ed})。値が飛ぶ可能性")
    return out


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--config", default=os.path.join(os.path.dirname(os.path.abspath(__file__)), "config.json"))
    ap.add_argument("--out", default="dashboard.json")
    a = ap.parse_args()
    cfg = json.load(open(a.config, encoding="utf-8"))
    today = dt.date.today()
    end = today.isoformat()
    earnings = cfg.get("earnings", {})
    rows_out, errors = [], []

    for h in cfg["holdings"]:
        sym = h["symbol"]
        start = min(dt.date.fromisoformat(h["entry"]) - dt.timedelta(days=60), today - dt.timedelta(days=460))
        try:
            rows = M.fetch(sym, start.isoformat(), end)
            s = base_stats(rows)
        except Exception as e:  # noqa: BLE001
            errors.append(f"{sym}: {e}")
            continue
        hist = M.analyze_holding(h, rows, cfg)
        stop = round(hist[-1][3], 2) if hist else None
        peak = round(hist[-1][2], 2) if hist else None
        broken = stop is not None and s["close"] < stop
        rows_out.append({
            "kind": "保有", "symbol": sym, **s,
            "cost": h["cost"], "shares": h["shares"], "pnl_pct": pct(s["close"], h["cost"]),
            "pnl_usd": round((s["close"] - h["cost"]) * h["shares"], 0),
            "value_usd": round(s["close"] * h["shares"], 0),
            "stop": stop, "stop_gap": pct(stop, s["close"]) if stop else None, "peak": peak,
            "atr_mult": h.get("atr_mult", cfg["atr_mult"]),
            "status": "売りライン割れ(ルール上は売り)" if broken else "保有を続ける(売りラインの上)",
            "status_level": "bad" if broken else "ok",
            "cautions": " / ".join(cautions(s, earnings.get(sym), today)),
            "note": "",
            **hist_of(rows, [(x[0], x[3]) for x in hist]),
        })

    for w in cfg.get("watchlist", []):
        sym = w["symbol"]
        try:
            rows = M.fetch(sym, (today - dt.timedelta(days=460)).isoformat(), end)
            s = base_stats(rows)
        except Exception as e:  # noqa: BLE001
            errors.append(f"{sym}: {e}")
            continue
        conds = {"50日線の上": s["ma50_gap"] > 0, "50日線が横ばい以上": s["_slope_raw"] >= -1.0,  # monitor.py と同じ基準(四捨五入前で判定)
                 "2ヶ月で-10%より上": s["m2"] > -10}
        ok = all(conds.values())
        atr = M.atr_series(rows, cfg["atr_period"])[-1]
        mult = 4.0 if "atr_mult 4.0" in w.get("note", "") else cfg["atr_mult"]
        rows_out.append({
            "kind": "見張り", "symbol": sym, **s,
            "status": "買いのタイミング条件クリア" if ok
            else "未達(" + "・".join(k for k, v in conds.items() if not v) + ")",
            "status_level": "ok" if ok else "warn",
            "first_stop_pct": round(-mult * atr / s["close"] * 100, 0), "atr_mult": mult,
            "cautions": " / ".join(cautions(s, earnings.get(sym), today)),
            "note": w.get("note", ""),
            **hist_of(rows),
        })

    try:
        spy = M.fetch(cfg["market_symbol"], (today - dt.timedelta(days=460)).isoformat(), end)
        s = base_stats(spy)
        c = [r["close"] for r in spy]
        ma200 = sum(c[-200:]) / 200
        rows_out.append({"kind": "相場", "symbol": cfg["market_symbol"], **s,
                         "ma200_gap": pct(c[-1], ma200),
                         "status": ("200日線の上" if c[-1] > ma200 else "200日線の下") + "・"
                         + ("50日線の上" if s["ma50_gap"] > 0 else "50日線の下(調整中)"),
                         "status_level": "ok" if c[-1] > ma200 else "bad", "cautions": "", "note": "",
                         **hist_of(spy)})
    except Exception as e:  # noqa: BLE001
        errors.append(f"SPY: {e}")

    asof = dt.datetime.now(dt.timezone.utc).isoformat(timespec="minutes")
    rows_out += [{"kind": "エラー", "symbol": e.split(":")[0], "status": e, "status_level": "bad"} for e in errors]
    for r in rows_out:
        r.pop("_slope_raw", None)
        r["asof"] = asof
    json.dump(rows_out, open(a.out, "w", encoding="utf-8"), ensure_ascii=False, separators=(",", ":"))
    print(f"{len(rows_out)}銘柄 / エラー{len(errors)}件 → {a.out}")


if __name__ == "__main__":
    main()
