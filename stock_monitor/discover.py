#!/usr/bin/env python3
"""新しい投資候補を見つける(SEC全社の財務データ + 新規上場一覧)。

トラック
  若手  : 上場5年以内。直近四半期の売上成長20%以上、営業利益率10%以上
  確立  : 上場年数は問わない。直近4四半期の売上10億ドル以上、成長15%以上、営業利益率15%以上
新顔の判定(記録を持たずにデータだけで判定)
  最新の四半期で条件を満たし、1つ前の四半期では満たしていなかった会社のうち、
  決算の届け出が直近 --fresh-days 日以内のもの。

使い方:
  python3 discover.py --mode new      # 前回の実行以降の新顔と新規上場だけ(なければ NO_NEW)
  python3 discover.py --mode full     # 条件を満たす全社の一覧
"""
import argparse
import datetime as dt
import json
import os
import time
import urllib.request

UA = {"User-Agent": "personal-stock-screener research-script"}
REV_TAGS = ("Revenues", "RevenueFromContractWithCustomerExcludingAssessedTax")
OP_TAG = "OperatingIncomeLoss"
EXCLUDE_SIC = ("Crude Petroleum", "Natural Gas", "Oil", "Petroleum", "Pipe Line", "Real Estate",
               "Mining", "Gold", "Coal", "Deep Sea", "Water Transportation", "Blank Check",
               "Commercial Banks", "Savings Institution", "Fire, Marine", "Life Insurance",
               "Surety Insurance", "Accident & Health Insurance", "Royalty", "Primary Production",
               "Primary Smelting", "Steel", "Electric Services", "Cogeneration", "Metal Mining")
MAX_MARGIN = 0.90   # これを超える営業利益率は一時金などの特殊要因とみなす(4Q合計が出せればそちらで判定)


def get(url, cache, headers=UA):
    if cache and os.path.exists(cache):
        return json.load(open(cache))
    for attempt in range(3):
        try:
            with urllib.request.urlopen(urllib.request.Request(url, headers=headers), timeout=60) as r:
                data = json.load(r)
            if cache:
                json.dump(data, open(cache, "w"))
            time.sleep(0.12)  # SECの上限(1秒10回)を守る
            return data
        except urllib.error.HTTPError as e:
            if e.code == 404:
                return None
            time.sleep(2 * (attempt + 1))
    raise RuntimeError(f"取得失敗: {url}")


def quarters_back(y, q, n):
    out = []
    for _ in range(n):
        out.append(f"CY{y}Q{q}")
        q -= 1
        if q == 0:
            y, q = y - 1, 4
    return out  # 新しい順


def load_frames(cache_dir, today):
    # 直近の四半期から9四半期分(4四半期合計の前年比に必要)
    y, q = today.year, (today.month - 1) // 3  # 今の四半期の1つ前
    if q == 0:
        y, q = y - 1, 4
    qs = quarters_back(y, q, 12)  # 10-12月期の補完に、その年の1〜3月期まで必要
    rev, op = {}, {}
    for p in qs:
        for tag in REV_TAGS:
            d = get(f"https://data.sec.gov/api/xbrl/frames/us-gaap/{tag}/USD/{p}.json",
                    os.path.join(cache_dir, f"{tag}_{p}.json"))
            for r in (d or {}).get("data", []):
                k = (r["cik"], p)
                rev[k] = max(rev.get(k, 0), r["val"])
        d = get(f"https://data.sec.gov/api/xbrl/frames/us-gaap/{OP_TAG}/USD/{p}.json",
                os.path.join(cache_dir, f"{OP_TAG}_{p}.json"))
        for r in (d or {}).get("data", []):
            op[(r["cik"], p)] = (r["val"], r["entityName"])
    # 10-12月期は年次報告にしか載らない会社が多いので「通年 - 1〜9月期」で補う
    for year in sorted({int(p[2:6]) for p in qs if p.endswith("Q4")}):
        fy_rev, fy_op = {}, {}
        for tag in REV_TAGS:
            d = get(f"https://data.sec.gov/api/xbrl/frames/us-gaap/{tag}/USD/CY{year}.json",
                    os.path.join(cache_dir, f"{tag}_CY{year}.json"))
            for r in (d or {}).get("data", []):
                fy_rev[r["cik"]] = max(fy_rev.get(r["cik"], 0), r["val"])
        d = get(f"https://data.sec.gov/api/xbrl/frames/us-gaap/{OP_TAG}/USD/CY{year}.json",
                os.path.join(cache_dir, f"{OP_TAG}_CY{year}.json"))
        for r in (d or {}).get("data", []):
            fy_op[r["cik"]] = (r["val"], r["entityName"])
        q123 = [f"CY{year}Q{k}" for k in (1, 2, 3)]
        q4 = f"CY{year}Q4"
        for cik, v in fy_rev.items():
            if (cik, q4) not in rev and all((cik, p) in rev for p in q123):
                d4 = v - sum(rev[(cik, p)] for p in q123)
                if d4 > 0:
                    rev[(cik, q4)] = d4
        for cik, (v, name) in fy_op.items():
            if (cik, q4) not in op and all((cik, p) in op for p in q123):
                op[(cik, q4)] = (v - sum(op[(cik, p)][0] for p in q123), name)
    return qs, rev, op


def metrics(cik, qs, rev, op, i):
    """qs[i] を最新四半期とした指標。データが欠けていれば None。"""
    # 10-12月期は年次報告にしか載らず単独の四半期値が無い会社が多いので、
    # 判定は「最新四半期 と 前年同期」の2点で行い、4四半期合計は揃う会社だけ参考に出す。
    need = qs[i:i + 8]
    if len(need) < 5:
        return None
    q0, q4 = need[0], need[4]
    if (cik, q0) not in rev or (cik, q4) not in rev or (cik, q0) not in op:
        return None
    r0, r4, o0 = rev[(cik, q0)], rev[(cik, q4)], op[(cik, q0)][0]
    if r0 <= 0 or r4 <= 0:
        return None
    m = {"q": q0, "rev_q": r0, "q_g": r0 / r4 - 1, "q_m": o0 / r0, "ann_rev": r0 * 4}
    if (cik, q4) in op:
        m["q_m_yago"] = op[(cik, q4)][0] / r4
    if len(need) == 8 and all((cik, p) in rev for p in need):
        r = [rev[(cik, p)] for p in need]
        m["ttm_rev"], m["ttm_g"] = sum(r[:4]), sum(r[:4]) / sum(r[4:8]) - 1
        o = [op.get((cik, p), (None,))[0] for p in need[:4]]
        if None not in o:
            m["ttm_m"] = sum(o) / sum(r[:4])
    return m


USE_TTM = [True]   # --quarter-only で False(比較用)


def sane(m):
    """一時金などで数字が跳ねた四半期を除く。4Q合計が出せればそちらで確かめる。"""
    if m["q_g"] > 3.0 or m["q_m"] > MAX_MARGIN:
        return "ttm_g" in m and m["ttm_g"] <= 3.0 and m.get("ttm_m", 1) <= MAX_MARGIN
    return True


def pass_young(m):
    if not m or m["rev_q"] < 20e6:
        return False
    ok = m["q_g"] >= 0.20 and m["q_m"] >= 0.10
    if USE_TTM[0]:
        ok = ok and sane(m)
        if "ttm_g" in m:
            ok = ok and m["ttm_g"] >= 0.15 and m.get("ttm_m", m["q_m"]) >= 0.10
    return ok


def pass_est(m):
    if not m:
        return False
    size = m.get("ttm_rev", m["ann_rev"])
    ok = size >= 1e9 and m["q_g"] >= 0.15 and m["q_m"] >= 0.15
    # 4四半期合計が出せる会社は、合計でも条件を満たすこと(1四半期だけの好調を除く)
    if USE_TTM[0]:
        ok = ok and sane(m)
        if "ttm_g" in m:
            ok = ok and m["ttm_g"] >= 0.15 and m.get("ttm_m", m["q_m"]) >= 0.15
    return ok


def profile(cik, cache_dir, today):
    s = get(f"https://data.sec.gov/submissions/CIK{int(cik):010d}.json",
            os.path.join(cache_dir, f"sub_{cik}.json"))
    rec = s["filings"]["recent"]
    forms, dates = rec["form"], rec["filingDate"]
    older = len(s["filings"].get("files", [])) > 0
    ipo = min([d for f, d in zip(forms, dates) if f in ("424B4", "424B1")], default=None)
    periodic = [d for f, d in zip(forms, dates) if f in ("10-Q", "10-K", "20-F", "40-F")]
    first_p = min(periodic, default=None)
    cutoff = (today - dt.timedelta(days=5 * 365)).isoformat()
    young = (not older) and (ipo or first_p or "0") >= cutoff and (first_p is None or first_p >= cutoff)
    listed = ipo or first_p
    return {"tickers": s.get("tickers") or [], "sic": s.get("sicDescription") or "",
            "spin": young and ipo is None,  # 新規上場の目論見書が無い = 分社化・合併による上場の可能性
            "young": young, "listed": listed, "last_periodic": max(periodic, default=None),
            "days_listed": (today - dt.date.fromisoformat(listed)).days if listed else None}


def new_ipos(today, days):
    """Polygonの新規上場一覧(直近days日)。取れなければ空。"""
    url = ("https://api.polygon.io/vX/reference/ipos?ipo_status=history&order=desc&limit=100"
           f"&listing_date.gte={(today - dt.timedelta(days=days)).isoformat()}")
    try:
        d = get(url, None, {"Authorization": "Bearer " + os.environ.get("POLYGON_KEY", "proxy-injected")})
        return [r for r in d.get("results", []) if r.get("security_type") == "CS"]
    except Exception:  # noqa: BLE001
        return None


def fmt(name, p, m):
    tk = ",".join(p["tickers"][:2]) or "-"
    if "ttm_g" in m:
        ttm = f"4Q成長{m['ttm_g']*100:+.0f}%" + (f" 4Q営利率{m['ttm_m']*100:.0f}%" if "ttm_m" in m else "")
        size = m["ttm_rev"]
    else:
        ttm, size = "4Q合計:データ不足", m["ann_rev"]
    infl = ""
    if "q_m_yago" in m and m["q_m"] - m["q_m_yago"] >= 0.10:
        infl = f" ★利益率の転換 {m['q_m_yago']*100:.0f}%→{m['q_m']*100:.0f}%"
    days = f" 上場{p['days_listed']}日" if p["days_listed"] is not None and p["young"] else ""
    if p.get("spin"):
        days += "(分社化・合併の可能性)"
    return (f"{tk:9} {name[:34]:34} {m['q']} 成長{m['q_g']*100:+.0f}% 営利率{m['q_m']*100:.0f}% | "
            f"{ttm} | 年売上${size/1e9:.1f}B{days} [{p['sic'][:24]}]{infl}")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--mode", choices=["new", "full"], default="new")
    ap.add_argument("--cache", default=os.path.join(os.path.dirname(os.path.abspath(__file__)), ".cache"))
    ap.add_argument("--fresh-days", type=int, help="何日以内に決算を届け出た会社を新顔とするか(既定: 前回の実行以降)")
    ap.add_argument("--config", default=os.path.join(os.path.dirname(os.path.abspath(__file__)), "config.json"),
                    help="保有・見張り中の銘柄は新顔から除く")
    ap.add_argument("--exclude", default="", help="既に把握している銘柄(カンマ区切り)")
    ap.add_argument("--quarter-only", action="store_true", help="4Q合計の確認をしない(比較用)")
    ap.add_argument("--top", type=int, default=25, help="fullモードで各トラックの表示上限")
    a = ap.parse_args()
    USE_TTM[0] = not a.quarter_only
    os.makedirs(a.cache, exist_ok=True)
    today = dt.date.today()
    # 毎朝(米国の平日)に動く前提: 月曜(UTC)は金曜からの3日分、それ以外は1日分
    if a.fresh_days is None:
        a.fresh_days = 3 if dt.datetime.now(dt.timezone.utc).weekday() == 0 else 1
    known = {s.strip().upper() for s in a.exclude.split(",") if s.strip()}
    if os.path.exists(a.config):
        c = json.load(open(a.config, encoding="utf-8"))
        known |= {x["symbol"] for x in c.get("holdings", []) + c.get("watchlist", []) + c.get("sold", [])}

    qs, rev, op = load_frames(a.cache, today)
    ciks = {k[0] for k in rev}
    found = {"young": [], "est": []}
    for cik in ciks:
        # 最新四半期は、その会社のデータがある新しいほう(qs[0] か qs[1])
        i = 0 if (cik, qs[0]) in rev and (cik, qs[0]) in op else 1
        m = metrics(cik, qs, rev, op, i)
        if not (pass_young(m) or pass_est(m)):
            continue
        prev = metrics(cik, qs, rev, op, i + 1)
        p = profile(cik, a.cache, today)
        if not p["tickers"] or any(x in p["sic"] for x in EXCLUDE_SIC):
            continue
        name = op.get((cik, m["q"]), (0, "?"))[1]
        fresh = p["last_periodic"] and (today - dt.date.fromisoformat(p["last_periodic"])).days <= a.fresh_days
        # 前の四半期のデータが無い会社は「新顔」と判定しない(判定できないため)
        if p["young"] and pass_young(m):
            found["young"].append((name, p, m, prev is not None and not pass_young(prev), fresh))
        elif not p["young"] and pass_est(m):
            found["est"].append((name, p, m, prev is not None and not pass_est(prev), fresh))

    out = []
    labels = {"young": "若手(上場5年以内)", "est": "確立企業(数倍狙い)"}
    for key in ("young", "est"):
        rows = sorted(found[key], key=lambda x: -x[2]["q_g"])
        if a.mode == "full":
            out.append(f"■{labels[key]}: {len(rows)}社" + (f"(上位{a.top}社を表示)" if len(rows) > a.top else ""))
            out += [("🆕" if newq else "  ") + fmt(n, p, m) for n, p, m, newq, fr in rows[:a.top]]
        else:
            new = [(n, p, m) for n, p, m, newq, fr in rows
                   if newq and fr and not (set(p["tickers"]) & known)]
            if new:
                out.append(f"🔎{labels[key]}の新顔: {len(new)}社(創業者経営は要確認)")
                out += ["  " + fmt(n, p, m) for n, p, m in new]

    ipos = new_ipos(today, a.fresh_days)
    if ipos is None:
        out.append("⚠️新規上場一覧(Polygon)の取得に失敗")
    elif ipos and a.mode == "new":
        out.append(f"🆕新規上場 {len(ipos)}社(上場直後は買わない。条件を満たせば四半期決算の後に候補入り)")
        out += [f"  {r['ticker']:6} {r.get('issuer_name','')[:40]} {r.get('listing_date')} "
                f"公開価格${r.get('final_issue_price')}" for r in ipos[:15]]
    print("\n".join(out) if out else "NO_NEW")


if __name__ == "__main__":
    main()
