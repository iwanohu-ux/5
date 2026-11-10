#!/usr/bin/env python3
"""保有銘柄・見張り候補のニュースと、SECへの重要な届け出を確認する。

  毎朝 : SECの臨時報告(8-K)や増資の届け出のうち、方針の「前提崩れ」に関わるものを拾う
  週次 : 上記に加えて、各銘柄の主なニュースの見出し(Polygon)を添える
  急落 : --news-for で指定した銘柄は、毎朝でも見出しを取る(急落の理由の確認用)

使い方:
  python3 news.py --mode daily
  python3 news.py --mode daily --news-for HNGE,GCT
  python3 news.py --mode weekly
"""
import argparse
import datetime as dt
import html
import json
import os
import re
import time
import urllib.request

HERE = os.path.dirname(os.path.abspath(__file__))
# SECは連絡先を含む名乗りを求める。連絡先はコードに書かず、環境変数 SEC_CONTACT から読む
CONTACT = os.environ.get("SEC_CONTACT", "").strip()
UA = {"User-Agent": f"personal-stock-screener {CONTACT}".strip() if CONTACT else "personal-stock-screener research-script"}

# 8-Kの項目番号のうち、方針の「前提崩れ」や重大な変化に関わるもの
RED_ITEMS = {
    "1.01": "重要な契約(買収・提携など)", "1.02": "重要な契約の終了", "1.03": "破産",
    "2.05": "事業撤退・リストラ", "2.06": "減損", "3.01": "上場廃止・上場基準の未達",
    "3.02": "新株の発行(希薄化)", "4.01": "監査法人の交代", "4.02": "過去の決算の訂正(重大な警告)",
    "5.01": "支配権の変更", "5.02": "役員・取締役の交代(創業者の退任なら前提崩れ)",
}
INFO_ITEMS = {"2.02": "決算発表"}
RED_FORMS = {
    "S-1": "株式の売り出し・増資の届け出", "S-3": "増資の枠の届け出", "S-3ASR": "増資の枠の届け出",
    "F-3": "増資の枠の届け出", "424B5": "増資・売り出しの実施", "424B3": "株式の売り出し",
    "SC 13D": "大株主(5%超)が経営に関与する意思を表明", "SC TO-T": "株式の公開買付け",
    "DEFM14A": "合併・買収の株主総会資料", "25-NSE": "上場廃止の届け出",
}


def get_json(url, headers):
    for attempt in range(3):
        try:
            with urllib.request.urlopen(urllib.request.Request(url, headers=headers), timeout=30) as r:
                return json.load(r)
        except Exception:  # noqa: BLE001
            time.sleep(2 * (attempt + 1))
    raise RuntimeError(url)


def sec_events(cik, since):
    s = get_json(f"https://data.sec.gov/submissions/CIK{int(cik):010d}.json", UA)
    time.sleep(0.12)
    rec = s["filings"]["recent"]
    out = []
    for form, date, items, acc, doc in zip(rec["form"], rec["filingDate"], rec["items"],
                                           rec["accessionNumber"], rec["primaryDocument"]):
        if date < since:
            break  # 新しい順に並んでいる
        url = f"https://www.sec.gov/Archives/edgar/data/{int(cik)}/{acc.replace('-', '')}/{doc}"
        if form in ("8-K", "8-K/A", "6-K"):
            its = [i.strip() for i in (items or "").split(",") if i.strip()]
            red = [f"{i}({RED_ITEMS[i]})" for i in its if i in RED_ITEMS]
            info = [f"{i}({INFO_ITEMS[i]})" for i in its if i in INFO_ITEMS]
            if red:
                ex = item_excerpt(url, [i for i in its if i in RED_ITEMS])
                what = f"{form} " + ", ".join(red) + (f"\n    【本文の抜粋】{ex}" if ex else "")
                out.append(("red", date, what, url))
            elif info:
                out.append(("info", date, f"{form} " + ", ".join(info), url))
        elif form in RED_FORMS:
            out.append(("red", date, f"{form}({RED_FORMS[form]})", url))
    return out


def fetch_text(url, limit=None):
    """SECの書類の本文をテキストで返す。連絡先が未設定なら読めないので None。"""
    if not CONTACT:
        return None
    try:
        with urllib.request.urlopen(urllib.request.Request(url, headers=UA), timeout=30) as r:
            raw = r.read().decode("utf-8", errors="ignore")
        time.sleep(0.12)
    except Exception:  # noqa: BLE001
        return None
    raw = re.sub(r"(?is)<(script|style).*?</\1>", " ", raw)
    t = html.unescape(re.sub(r"<[^>]+>", " ", raw))
    t = re.sub(r"\s+", " ", t).strip()
    return t[:limit] if limit else t


def item_excerpt(url, items, width=900):
    """8-K本文から、指定した項目(例: 5.02)の部分を抜き出す。"""
    t = fetch_text(url)
    if not t:
        return None
    parts = []
    for i in items:
        m = re.search(rf"Item\s*{re.escape(i)}", t, re.I)
        if m:
            parts.append(t[m.start():m.start() + width])
    return " / ".join(parts) or t[:width]


def press_release(cik, acc, width=2500):
    """決算の8-Kに添付されたプレスリリース(EX-99.1)の冒頭。"""
    base = f"https://www.sec.gov/Archives/edgar/data/{int(cik)}/{acc.replace('-', '')}"
    idx = fetch_text(base + "/index.json")
    if not idx:
        return None
    try:
        files = [f["name"] for f in json.loads(idx)["directory"]["item"]]
    except Exception:  # noqa: BLE001
        return None
    ex = [f for f in files if re.search(r"ex-?99", f, re.I) and f.lower().endswith((".htm", ".html", ".txt"))]
    return fetch_text(f"{base}/{ex[0]}", width) if ex else None


def earnings_filings(cik, since_utc):
    """since_utc 以降にSECが受理した決算発表(8-K/6-K の項目2.02)。受理時刻つき。"""
    s = get_json(f"https://data.sec.gov/submissions/CIK{int(cik):010d}.json", UA)
    time.sleep(0.12)
    rec = s["filings"]["recent"]
    out = []
    for form, acc_t, items, acc, doc in zip(rec["form"], rec["acceptanceDateTime"], rec["items"],
                                            rec["accessionNumber"], rec["primaryDocument"]):
        t = dt.datetime.fromisoformat(acc_t.replace("Z", "+00:00"))
        if t < since_utc:
            break
        if form in ("8-K", "8-K/A", "6-K") and "2.02" in (items or ""):
            url = f"https://www.sec.gov/Archives/edgar/data/{int(cik)}/{acc.replace('-', '')}/{doc}"
            out.append((t, url, press_release(cik, acc)))
    return out


_poly_last = [0.0]


def headlines(sym, since, limit=5):
    wait = 13 - (time.time() - _poly_last[0])  # Polygon無料枠: 1分5回
    if wait > 0:
        time.sleep(wait)
    url = (f"https://api.polygon.io/v2/reference/news?ticker={sym}&order=desc&sort=published_utc"
           f"&limit={limit}&published_utc.gte={since}")
    try:
        j = get_json(url, {"Authorization": "Bearer " + os.environ.get("POLYGON_KEY", "proxy-injected")})
    finally:
        _poly_last[0] = time.time()
    return [(r["published_utc"][:10], r.get("publisher", {}).get("name", ""), r["title"], r.get("article_url", ""))
            for r in j.get("results", [])]


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--mode", choices=["daily", "weekly", "earnings"], default="daily")
    ap.add_argument("--news-for", default="", help="毎朝でも見出しを取る銘柄(カンマ区切り)")
    ap.add_argument("--config", default=os.path.join(HERE, "config.json"))
    ap.add_argument("--days", type=int, help="確認する日数(既定: 毎朝は前回の実行以降、週次は7日)")
    ap.add_argument("--hours", type=int, default=17, help="earningsモード: 何時間前からの決算発表を拾うか")
    a = ap.parse_args()
    cfg = json.load(open(a.config, encoding="utf-8"))

    if a.mode == "earnings":
        # 決算速報: 保有・見張り中の銘柄が直近に決算を出していれば、届け出と見出しを並べる
        now = dt.datetime.now(dt.timezone.utc)
        since_utc = now - dt.timedelta(hours=a.hours)
        out, errors = [], []
        for x, kind in [(h, "保有") for h in cfg["holdings"]] + [(w, "見張り") for w in cfg.get("watchlist", [])]:
            sym, cik = x["symbol"], x.get("cik")
            if not cik:
                continue
            try:
                found = earnings_filings(cik, since_utc)
            except Exception as e:  # noqa: BLE001
                errors.append(f"{sym}: SECの届け出の取得失敗 ({e})")
                continue
            for t, url, pr in found:
                jst = t.astimezone(dt.timezone(dt.timedelta(hours=9))).strftime("%m/%d %H:%M")
                out.append(f"📊{sym}({kind}) 決算発表 {jst}(日本時間)\n    {url}")
                out.append(f"    【決算資料の冒頭】{pr}" if pr else "    【決算資料】本文は未取得(SEC_CONTACT未設定か取得失敗)")
                try:
                    for d, p, title, u in headlines(sym, since_utc.date().isoformat(), limit=6):
                        out.append(f"    {d} [{p}] {title}\n      {u}")
                except Exception as e:  # noqa: BLE001
                    errors.append(f"{sym}: ニュースの取得失敗 ({e})")
        out += [f"⚠️{e}" for e in errors]
        print("\n".join(out) if out else "NO_EARNINGS")
        return

    today = dt.datetime.now(dt.timezone.utc).date()
    days = a.days or (7 if a.mode == "weekly" else (3 if today.weekday() == 0 else 1))
    since = (today - dt.timedelta(days=days)).isoformat()

    entries = [(h["symbol"], h.get("cik"), "保有") for h in cfg["holdings"]]
    entries += [(w["symbol"], w.get("cik"), "見張り") for w in cfg.get("watchlist", [])]
    out, errors = [], []
    for sym, cik, kind in entries:
        if not cik:
            errors.append(f"{sym}: SECの登録番号(cik)が未設定")
            continue
        try:
            for level, date, what, url in sec_events(cik, since):
                mark = "🚨" if level == "red" else "ℹ️"
                out.append(f"{mark}{sym}({kind}) {date} {what}\n    {url}")
        except Exception as e:  # noqa: BLE001
            errors.append(f"{sym}: SECの届け出の取得失敗 ({e})")

    targets = [s for s, _, _ in entries] if a.mode == "weekly" else \
        [s.strip().upper() for s in a.news_for.split(",") if s.strip()]
    for sym in targets:
        try:
            hs = headlines(sym, since)
        except Exception as e:  # noqa: BLE001
            errors.append(f"{sym}: ニュースの取得失敗 ({e})")
            continue
        if hs:
            out.append(f"📰{sym} の見出し({since}以降)")
            out += [f"    {d} [{p}] {t}\n      {u}" for d, p, t, u in hs]

    out += [f"⚠️{e}" for e in errors]
    print("\n".join(out) if out else "NO_NEWS")


if __name__ == "__main__":
    main()
