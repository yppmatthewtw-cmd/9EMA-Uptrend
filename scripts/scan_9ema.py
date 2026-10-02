#!/usr/bin/env python3
"""9EMA-Uptrend scan: US common stocks whose 21-day EMA is rising, whose 9-day
EMA sits above the 21-day EMA (or has just crossed above it), that print higher
lows, whose price has come back to the 9 / 21 EMA, and whose consolidation
leans upward like an ascending triangle.

Measured on the last complete session t (daily bars from Yahoo, see
.github/workflows/fetch_yahoo_eod.yml; an intraday roll-up can fill a session
Yahoo has not published daily bars for yet).

  C1 21EMA 向上       EMA21[t] > EMA21[t-1] and EMA21[t] > EMA21[t-5]
                      (5-session slope > 0); the share of rising steps over the
                      last 10 is recorded for the score.
  C2 9EMA 在 21EMA 上  EMA9[t] > EMA21[t].  The most recent golden cross (first
                      session with EMA9 > EMA21 after a session with EMA9 <=
                      EMA21) is dated; "剛黃金交叉" when it is <= cross_days
                      sessions ago, otherwise "持續在上方".
  C3 一底高於一底      pivot lows on the daily low (lowest of the 3 bars either
                      side); the last two confirmed pivots inside the last
                      hl_look sessions must rise, the latest must be <= hl_recent
                      sessions ago, and no close since may sit below it.  The
                      length of the rising streak is recorded.
  C5 上升三角／重心偏上  over the last tri_win sessions: the volume-weighted
                      average price's position inside the window's high-low
                      range (重心) >= cog_min, the regression slope of the lows > 0, and
                      the regression slope of the highs >= high_slope_min
                      (flat-to-rising top; a falling top is a different pattern).
                      The flatness of the top, the rise of the lows, the 重心 and
                      the number of touches of the top make the pattern score.

  C6 MACD 慢線向上      the MACD signal line (EMA9 of EMA12 − EMA26) is higher
                      than the session before.
  C7 柱狀圖動能轉勢     the histogram's first bar of slowing decline (a 淺紅 bar
                      after 深紅: hist < 0 and rising after falling) or of
                      re-accelerating advance (a 深綠 bar after 淺綠: hist >= 0
                      and rising after falling) printed within the last
                      hist_days sessions, and every bar since has kept rising.
                      Colours follow TradingView's four-colour histogram.

Universe: a bar on t, >= min_hist sessions of history, close >= min_px,
20-session median dollar volume >= min_dv.

Scores (0-100; each term piecewise-linear and clipped so the workbook can
recompute them from the inputs beside them):
  趨勢分數   0.4 lin(EMA21 5日斜率%, 0->0, 2.5->1) + 0.3 lin(EMA21 10步上升比例,
             0.6->0, 1.0->1) + 0.3 lin(RS21 (21日回報 − 全池中位數), 0->0, 15%->1)
  回調質素   0.4 lin(距最近 EMA 的 ATR 倍數, near_atr->0, 0->1)
             + 0.3 lin(近10日 跌日量/升日量, 1.2->0, 0.6->1)
             + 0.3 收市在當日高低區間位置
  底部結構   0.5 lin(連續上升底數目, 1->0, 4->1) + 0.3 lin(最後一底較上一底高 %,
             0->0, 8->1) + 0.2 lin(最後一底距今日數, 25->0, 5->1)
  三角形分數 0.35 lin(|高位斜率| %/日, 0.30->0, 0.05->1) + 0.30 lin(低位斜率 %/日,
             0->0, 0.30->1) + 0.25 lin(重心, 0.40->0, 0.70->1) + 0.10 (頂部觸及次數
             >=2 -> 1, 1 -> 0.5)
  綜合分數   0.25 趨勢 + 0.25 回調質素 + 0.20 底部結構 + 0.30 三角形

Sheets (subsets of the first tier, R2.00 onward; a name can sit on more than one):
  a 剛黃金交叉        the golden cross is <= cross_days sessions ago.
  b 回到 21EMA        the close is within near_pct of EMA21 and either within
                      near_atr x ATR14 of it or a low inside the last touch_days
                      sessions touched that day's EMA21 (low <= EMA x (1 +
                      touch_tol)); the close must not sit more than under_tol
                      below EMA21 (a breakdown is not a pull-back).
  c 回到 9EMA         the same measured against EMA9.  A name near both goes to
                      the sheet of the EMA its close is nearer to (by %), so b
                      holds the deeper pull-backs and c the shallow ones.

Second tier (差一項): fails exactly one of C1/C2/C3/C5/C6/C7 at the strict
thresholds but passes it at the loosened ones.  Third tier (候補): passes the
four price criteria and fails only the MACD ones, i.e. the names to watch for
the histogram's turn.

Env: YAHOO (comma list of daily .csv.gz, oldest first; later files win on a
duplicate symbol/date), SUPP (comma list of intraday roll-up .csv.gz used only
for sessions the daily files lack), SNAP (Nasdaq screener snapshot CSV for
name / sector / industry / market cap), EXCH_DIR (rreichel3/us-stock-symbols
checkout for the exchange), LAST_DATE (override), OUT_JSON.
"""
import csv, json, math, os, re, statistics, sys
from collections import defaultdict

import numpy as np
import pandas as pd

YAHOO = [p for p in os.environ.get("YAHOO", "").split(",") if p]
SUPP = [p for p in os.environ.get("SUPP", "").split(",") if p]
SNAP = os.environ.get("SNAP", "")
EXCH_DIR = os.environ.get("EXCH_DIR", "/home/user/rreichel3/us-stock-symbols")
LAST_DATE = os.environ.get("LAST_DATE", "")
OUT_JSON = os.environ.get("OUT_JSON", "data/screen_r1.json")
if not YAHOO:
    sys.exit("YAHOO=<daily bars .csv.gz[,more]> is required")

P0 = {
    "min_hist": 90, "min_px": 2.0, "min_dv": 1_000_000,
    "slope_lag": 5, "steps_look": 10,
    "cross_days": 5,
    "piv": 3, "hl_look": 60, "hl_recent": 30,
    "near_atr": 0.75, "near_pct": 0.03, "touch_tol": 0.005, "touch_days": 3, "under_tol": -0.03,
    "tri_win": 25, "cog_min": 0.50, "high_slope_min": -0.10, "high_slope_max": 1.0,
    "macd_fast": 12, "macd_slow": 26, "macd_sig": 9, "hist_days": 3,
}
PL = dict(P0, cross_days=8, hl_look=80, hl_recent=40, near_atr=1.25, near_pct=0.05, touch_tol=0.01,
          under_tol=-0.05, cog_min=0.45, high_slope_min=-0.20, high_slope_max=1.5, hist_days=5)
W = {"trend": 0.25, "pb": 0.25, "hl": 0.20, "tri": 0.30}
LIN = {   # (x0 -> 0, x1 -> 1) for every sub-score, mirrored in the workbook formulas
    "t_slope": (0.0, 2.5), "t_steps": (0.6, 1.0), "t_rs": (0.0, 15.0),
    "p_near": (P0["near_atr"], 0.0), "p_vol": (1.2, 0.6),
    "h_n": (1.0, 4.0), "h_rise": (0.0, 8.0), "h_age": (25.0, 5.0),
    "r_flat": (0.30, 0.05), "r_low": (0.0, 0.30), "r_cog": (0.40, 0.70),
}


def lin(x, x0, x1):
    if x is None or (isinstance(x, float) and math.isnan(x)):
        return 0.0
    v = (x - x0) / (x1 - x0)
    return max(0.0, min(1.0, v))


def norm(sym):
    return sym.replace("/", ".").strip().upper()


# ---------------------------------------------------------------- bars
def load(paths):
    frames = []
    for p in paths:
        df = pd.read_csv(p, dtype={"symbol": str, "date": str})
        df["src"] = os.path.basename(p)
        frames.append(df)
    return pd.concat(frames, ignore_index=True) if frames else pd.DataFrame()


daily = load(YAHOO)
daily = daily.drop_duplicates(["symbol", "date"], keep="last")
daily["kind"] = "daily"
supp = load(SUPP)
if len(supp):
    supp = supp.drop_duplicates(["symbol", "date"], keep="last")
    supp["kind"] = "intraday"
    have = set(zip(daily["symbol"], daily["date"]))
    supp = supp[[(s, d) not in have for s, d in zip(supp["symbol"], supp["date"])]]
    bars = pd.concat([daily, supp], ignore_index=True)
else:
    bars = daily
bars = bars.dropna(subset=["close"]).sort_values(["symbol", "date"])
n_sym_all = bars["symbol"].nunique()
cover = bars.groupby("date")["symbol"].nunique()
if LAST_DATE:
    last_date = LAST_DATE
else:
    ok = cover[cover >= 0.6 * n_sym_all]
    last_date = ok.index.max()
bars = bars[bars["date"] <= last_date]
cal = sorted(bars["date"].unique())
T = len(cal)
print(f"{n_sym_all} symbols, {T} sessions {cal[0]} -> {last_date}; "
      f"bars on {last_date}: {int(cover.get(last_date, 0))} "
      f"(intraday roll-up: {int((bars[(bars.date == last_date)].kind == 'intraday').sum())})")

# ---------------------------------------------------------------- metadata
meta = {}
if SNAP:
    with open(SNAP, encoding="utf-8") as f:
        for r in csv.DictReader(f):
            s = norm(r["symbol"])
            try:
                mc = float(r["marketCap"] or 0)
            except ValueError:
                mc = 0.0
            name = re.sub(r"\s*\(Name to be changed[^)]*\)", "", r["name"])
            for tail in (" Common Stock", " Ordinary Shares", " Class A", " Common Shares", " American Depositary Shares"):
                name = name.split(tail)[0]
            meta[s] = {"name": name.strip().rstrip(","), "sector": (r["sector"] or "—").strip(),
                       "industry": (r["industry"] or "—").strip(), "country": r["country"].strip(),
                       "ipoyear": r["ipoyear"].strip(), "mcap": mc}
exch = {}
for ex, fn in (("NASDAQ", "nasdaq/nasdaq_tickers.txt"), ("NYSE", "nyse/nyse_tickers.txt"), ("AMEX", "amex/amex_tickers.txt")):
    p = os.path.join(EXCH_DIR, fn)
    if os.path.exists(p):
        for line in open(p):
            if line.strip():
                exch.setdefault(norm(line), ex)

ZH_SECTOR = {
    "Technology": "科技", "Consumer Discretionary": "非必需消費", "Health Care": "醫療保健",
    "Finance": "金融", "Industrials": "工業", "Consumer Staples": "必需消費",
    "Energy": "能源", "Real Estate": "房地產", "Utilities": "公用事業",
    "Basic Materials": "原材料", "Telecommunications": "電訊", "Miscellaneous": "其他", "—": "—", "": "—",
}


def cap_group(mc):
    if not mc:
        return "x"
    return "a" if mc >= 10e9 else "b" if mc >= 2e9 else "c"


# ---------------------------------------------------------------- indicators
def ema(x, n):
    a = 2.0 / (n + 1)
    out = np.empty_like(x)
    out[0] = x[0]
    for i in range(1, len(x)):
        out[i] = a * x[i] + (1 - a) * out[i - 1]
    return out


def atr(h, l, c, n=14):
    tr = np.empty_like(c)
    tr[0] = h[0] - l[0]
    for i in range(1, len(c)):
        tr[i] = max(h[i] - l[i], abs(h[i] - c[i - 1]), abs(l[i] - c[i - 1]))
    out = np.full_like(c, np.nan)
    if len(c) >= n:
        out[n - 1] = tr[:n].mean()
        for i in range(n, len(c)):
            out[i] = (out[i - 1] * (n - 1) + tr[i]) / n
    return out


def pivots(v, k, low=True):
    """Indices i with v[i] the extreme of v[i-k .. i+k] (strict against the
    outermost bars, so a flat shelf counts once); pivots closer than k bars
    keep the more extreme one."""
    n = len(v)
    raw = []
    for i in range(k, n - k):
        w = v[i - k:i + k + 1]
        if low:
            if v[i] == w.min() and v[i - k] > v[i] and v[i + k] > v[i]:
                raw.append(i)
        else:
            if v[i] == w.max() and v[i - k] < v[i] and v[i + k] < v[i]:
                raw.append(i)
    out = []
    for i in raw:
        if out and i - out[-1] <= k:
            if (low and v[i] < v[out[-1]]) or (not low and v[i] > v[out[-1]]):
                out[-1] = i
        else:
            out.append(i)
    return out


def slope_pct(y, idx, px):
    """Least-squares slope of y over bar index, as % of px per session."""
    if len(idx) < 2:
        return None
    x = np.asarray(idx, float); yy = np.asarray(y, float)
    x = x - x.mean()
    if (x * x).sum() == 0:
        return None
    return float((x * (yy - yy.mean())).sum() / (x * x).sum() / px * 100)


series = {}
for sym, g in bars.groupby("symbol", sort=False):
    g = g.sort_values("date")
    series[sym] = {
        "dates": g["date"].tolist(),
        "o": g["open"].to_numpy(float), "h": g["high"].to_numpy(float), "l": g["low"].to_numpy(float),
        "c": g["close"].to_numpy(float), "v": g["volume"].to_numpy(float), "kind": g["kind"].tolist(),
    }

# ---------------------------------------------------------------- universe
funnel = {"symbols": len(series), "bar_on_last": 0, "hist": 0, "price": 0, "liq": 0}
elig = []
for sym, s in series.items():
    if s["dates"][-1] != last_date:
        continue
    funnel["bar_on_last"] += 1
    if len(s["c"]) < P0["min_hist"]:
        continue
    funnel["hist"] += 1
    if s["c"][-1] < P0["min_px"]:
        continue
    funnel["price"] += 1
    dv = s["c"][-20:] * s["v"][-20:]
    if np.median(dv) < P0["min_dv"]:
        continue
    funnel["liq"] += 1
    elig.append(sym)
print("universe:", funnel)

# ---------------------------------------------------------------- measure
rows = {}
rets21 = {}
for sym in elig:
    s = series[sym]
    c, h, l, o, v = s["c"], s["h"], s["l"], s["o"], s["v"]
    n = len(c); t = n - 1
    e9, e21 = ema(c, 9), ema(c, 21)
    a14 = atr(h, l, c, 14)
    macd_line = ema(c, P0["macd_fast"]) - ema(c, P0["macd_slow"])
    sig = ema(macd_line, P0["macd_sig"])
    hist = macd_line - sig

    def hist_colour(i):            # TradingView's four-colour histogram
        if hist[i] >= 0:
            return "深綠" if hist[i] > hist[i - 1] else "淺綠"
        return "深紅" if hist[i] < hist[i - 1] else "淺紅"

    colours = [hist_colour(i) for i in range(t - 7, t + 1)]
    # the momentum turn: first 淺紅 after 深紅, or first 深綠 after 淺綠 — how many sessions ago, with every bar since still improving
    turn_ago = None
    for k in range(0, 8):
        i = t - k
        improving = all(hist[j] > hist[j - 1] for j in range(i, t + 1))
        if not improving:
            break
        if hist[i] < 0 and hist[i - 1] < hist[i - 2]:          # 深紅 → 淺紅
            turn_ago = k; turn_kind = "深紅→淺紅"; break
        if hist[i] >= 0 and hist[i - 1] <= hist[i - 2] and hist[i - 1] >= 0:   # 淺綠 → 深綠
            turn_ago = k; turn_kind = "淺綠→深綠"; break
    if turn_ago is None:
        turn_kind = "—"
    rets21[sym] = c[t] / c[t - 21] - 1
    # C1
    lag = P0["slope_lag"]
    slope21 = (e21[t] / e21[t - lag] - 1) * 100
    steps = sum(1 for k in range(P0["steps_look"]) if e21[t - k] > e21[t - k - 1]) / P0["steps_look"]
    # C2: last golden cross
    cross = None
    for k in range(t, 0, -1):
        if e9[k] > e21[k] and e9[k - 1] <= e21[k - 1]:
            cross = k; break
    above_days = (t - cross) if cross is not None else None
    # C3: pivot lows
    piv_lo = pivots(l, P0["piv"], low=True)
    piv_hi = pivots(h, P0["piv"], low=False)

    def hl_struct(look, recent):
        inw = [i for i in piv_lo if i >= t - look]
        if len(inw) < 2:
            return None
        streak = 1
        for a, b in zip(inw[::-1][1:], inw[::-1]):   # walk back while rising
            if l[b] > l[a]:
                streak += 1
            else:
                break
        last, prev = inw[-1], inw[-2]
        ok = l[last] > l[prev] and last >= t - recent and c[last + 1:].min() >= l[last] * 0.995 if last + 1 <= t else l[last] > l[prev]
        return {"ok": bool(ok), "streak": streak, "last_i": last, "prev_i": prev,
                "rise": (l[last] / l[prev] - 1) * 100, "age": t - last, "n_in_win": len(inw)}

    hl0 = hl_struct(P0["hl_look"], P0["hl_recent"])
    hlL = hl_struct(PL["hl_look"], PL["hl_recent"])
    # C4: proximity
    A = a14[t] if not math.isnan(a14[t]) else (h[t] - l[t])
    d9 = c[t] / e9[t] - 1; d21 = c[t] / e21[t] - 1
    d9_atr = (c[t] - e9[t]) / A if A > 0 else 0.0
    d21_atr = (c[t] - e21[t]) / A if A > 0 else 0.0
    touch9 = bool(l[t] <= e9[t] * (1 + P0["touch_tol"]))
    touch21 = bool(l[t] <= e21[t] * (1 + P0["touch_tol"]))

    def touched(e, P):          # a low inside the last touch_days sessions at or under that day's EMA
        return any(l[i] <= e[i] * (1 + P["touch_tol"]) for i in range(t - P["touch_days"] + 1, t + 1))
    hi15 = c[t - 15:t + 1].max()
    pb_depth = (c[t] / hi15 - 1) * 100
    # C5: triangle window
    wn = P0["tri_win"]; w0 = t - wn + 1
    wh, wl, wc, wv = h[w0:t + 1], l[w0:t + 1], c[w0:t + 1], v[w0:t + 1]
    rng_hi, rng_lo = wh.max(), wl.min()
    vwap = float((wc * wv).sum() / wv.sum()) if wv.sum() > 0 else float(wc.mean())
    cog = (vwap - rng_lo) / (rng_hi - rng_lo) if rng_hi > rng_lo else 0.5    # 重心: VWAP inside the window's high-low range
    c_hi, c_lo = float(wc.max()), float(wc.min())
    cog_c = (vwap - c_lo) / (c_hi - c_lo) if c_hi > c_lo else 0.5           # the same inside the closes' range (shown, not used)
    ph_w = [i for i in piv_hi if i >= w0]
    pl_w = [i for i in piv_lo if i >= w0]
    sl_hi = slope_pct(h[ph_w], ph_w, c[t]) if len(ph_w) >= 2 else slope_pct(wh, list(range(w0, t + 1)), c[t])
    sl_lo = slope_pct(l[pl_w], pl_w, c[t]) if len(pl_w) >= 2 else slope_pct(wl, list(range(w0, t + 1)), c[t])
    hi_src = "樞軸高" if len(ph_w) >= 2 else "全部高位"
    lo_src = "樞軸低" if len(pl_w) >= 2 else "全部低位"
    top = wh.max()
    touches = int((wh >= top * 0.985).sum())
    half = wn // 2
    cog_first = float(np.median(wc[:half])); cog_second = float(np.median(wc[half:]))
    atr5 = float(np.mean([max(h[i] - l[i], abs(h[i] - c[i - 1]), abs(l[i] - c[i - 1])) for i in range(t - 4, t + 1)]))
    atr20 = float(np.mean([max(h[i] - l[i], abs(h[i] - c[i - 1]), abs(l[i] - c[i - 1])) for i in range(t - 19, t + 1)]))
    # volume on down vs up days (last 10)
    dn, up = [], []
    for i in range(t - 9, t + 1):
        if c[i] < c[i - 1]:
            dn.append(v[i])
        elif c[i] > c[i - 1]:
            up.append(v[i])
    dv_ratio = (np.mean(dn) / np.mean(up)) if dn and up and np.mean(up) > 0 else 1.0
    bar_pos = (c[t] - l[t]) / (h[t] - l[t]) if h[t] > l[t] else 0.5

    def crit(P, hl):
        c1 = e21[t] > e21[t - 1] and e21[t] > e21[t - P["slope_lag"]]
        c2 = e9[t] > e21[t]
        c3 = bool(hl and hl["ok"])
        c5 = (cog >= P["cog_min"] and (sl_lo is not None and sl_lo > 0)
              and (sl_hi is not None and P["high_slope_min"] <= sl_hi <= P["high_slope_max"]))
        c6 = sig[t] > sig[t - 1]                                   # MACD 慢線（訊號線）向上
        c7 = turn_ago is not None and turn_ago <= P["hist_days"]    # 柱狀圖動能減速／加速嘅第一根已出現
        return [bool(c1), bool(c2), c3, bool(c5), bool(c6), bool(c7)]

    def near(P):
        """Sheet conditions (not core criteria): price back at the 9EMA / 21EMA."""
        near9 = abs(d9) <= P["near_pct"] and (abs(d9_atr) <= P["near_atr"] or touched(e9, P))
        near21 = abs(d21) <= P["near_pct"] and (abs(d21_atr) <= P["near_atr"] or touched(e21, P))
        held = d21 >= P["under_tol"]
        return bool(near9 and held), bool(near21 and held)

    f0 = crit(P0, hl0)
    fL = crit(PL, hlL)
    near9_0, near21_0 = near(P0)
    fails0 = [k for k, ok in zip(("C1", "C2", "C3", "C5", "C6", "C7"), f0) if not ok]
    tier = 1 if not fails0 else (2 if len(fails0) == 1 and all(fL) else (3 if set(fails0) <= {"C6", "C7"} else 0))
    which = "9EMA" if abs(d9_atr) <= abs(d21_atr) else "21EMA"
    rows[sym] = {
        "sym": sym, "tier": tier, "fails": fails0, "fails_loose": [k for k, ok in zip(("C1", "C2", "C3", "C5", "C6", "C7"), fL) if not ok],
        "macd": round(float(macd_line[t]), 4), "sig": round(float(sig[t]), 4), "sig_prev": round(float(sig[t - 1]), 4),
        "hist": round(float(hist[t]), 4), "hist_prev": round(float(hist[t - 1]), 4), "hist_prev2": round(float(hist[t - 2]), 4),
        "hist_colour": colours[-1], "hist_seq": " ".join(colours), "turn_ago": turn_ago, "turn_kind": turn_kind,
        "near9": near9_0, "near21": near21_0,
        "close": round(float(c[t]), 4), "prev": round(float(c[t - 1]), 4), "open": round(float(o[t]), 4),
        "high": round(float(h[t]), 4), "low": round(float(l[t]), 4), "vol": float(v[t]),
        "bar_kind": s["kind"][t],
        "ema9": round(float(e9[t]), 4), "ema21": round(float(e21[t]), 4), "ema21_lag": round(float(e21[t - lag]), 4),
        "ema21_prev": round(float(e21[t - 1]), 4), "ema9_prev": round(float(e9[t - 1]), 4),
        "slope21": round(float(slope21), 4), "steps21": round(steps, 2),
        "cross_date": s["dates"][cross] if cross is not None else None, "above_days": above_days,
        "status": ("剛黃金交叉" if above_days is not None and above_days <= P0["cross_days"] else "持續在上方") if e9[t] > e21[t] else "9EMA在下方",
        "atr14": round(float(A), 4), "atr5": round(atr5, 4), "atr20": round(atr20, 4),
        "d9": round(float(d9), 5), "d21": round(float(d21), 5), "d9_atr": round(float(d9_atr), 3), "d21_atr": round(float(d21_atr), 3),
        "touch9": touch9, "touch21": touch21, "touch3_9": bool(touched(e9, P0)), "touch3_21": bool(touched(e21, P0)), "near_which": which,
        "hi15": round(float(hi15), 4), "pb_depth": round(float(pb_depth), 3),
        "hl": None if not hl0 else {"streak": hl0["streak"], "last_date": s["dates"][hl0["last_i"]], "last_low": round(float(l[hl0["last_i"]]), 4),
                                    "prev_date": s["dates"][hl0["prev_i"]], "prev_low": round(float(l[hl0["prev_i"]]), 4),
                                    "rise": round(hl0["rise"], 3), "age": hl0["age"], "n_in_win": hl0["n_in_win"], "ok": hl0["ok"]},
        "piv_lows": [(s["dates"][i], round(float(l[i]), 4)) for i in piv_lo if i >= t - P0["hl_look"]],
        "piv_highs": [(s["dates"][i], round(float(h[i]), 4)) for i in piv_hi if i >= w0],
        "tri": {"win": wn, "hi": round(float(rng_hi), 4), "lo": round(float(rng_lo), 4), "c_hi": round(c_hi, 4), "c_lo": round(c_lo, 4),
                "vwap": round(vwap, 4), "cog": round(float(cog), 4), "cog_c": round(float(cog_c), 4),
                "sl_hi": None if sl_hi is None else round(sl_hi, 4), "sl_lo": None if sl_lo is None else round(sl_lo, 4),
                "hi_src": hi_src, "lo_src": lo_src, "touches": touches,
                "med_first": round(cog_first, 4), "med_second": round(cog_second, 4)},
        "dv_ratio": round(float(dv_ratio), 4), "bar_pos": round(float(bar_pos), 4),
        "ret21": round(float(rets21[sym]) * 100, 3),
        "dv20_med": float(np.median(c[t - 19:t + 1] * v[t - 19:t + 1])),
        "vol20": float(np.mean(v[t - 19:t + 1])),
        "stats_above21": bool(c[t] > e21[t]),
    }
    r = rows[sym]
    m = meta.get(sym, {})
    r.update({"name": m.get("name", sym), "sector": m.get("sector", "—"), "sector_zh": ZH_SECTOR.get(m.get("sector", "—"), m.get("sector", "—")),
              "industry": m.get("industry", "—"), "country": m.get("country", ""), "ipoyear": m.get("ipoyear", ""),
              "mcap": m.get("mcap", 0.0), "cap": cap_group(m.get("mcap", 0.0)), "exch": exch.get(sym, "—")})

med21 = statistics.median(rets21.values())
# pattern label + scores
for sym, r in rows.items():
    r["rs21"] = round((rets21[sym] - med21) * 100, 3)
    tr = r["tri"]
    sh, sl = tr["sl_hi"], tr["sl_lo"]
    if sh is None or sl is None:
        lab = "未成形"
    elif sl > 0 and abs(sh) <= 0.10:
        lab = "上升三角（平頂）"
    elif sl > 0 and sh > 0.10 and sl > sh * 1.25:
        lab = "上傾收窄（楔形）"
    elif sl > 0 and sh > 0.10 and sh > sl * 1.25:
        lab = "上傾擴張"
    elif sl > 0 and sh > 0.10:
        lab = "上升通道"
    elif sl > 0 and sh < -0.10:
        lab = "對稱三角"
    elif sl <= 0 and abs(sh) <= 0.10:
        lab = "矩形／平台"
    else:
        lab = "未成形"
    r["pattern"] = lab
    hl = r["hl"] or {}
    # score from the rounded inputs the workbook carries, so its formulas reproduce these numbers exactly
    r["slope21"] = round((r["ema21"] / r["ema21_lag"] - 1) * 100, 6)
    r["d9"] = round(r["close"] / r["ema9"] - 1, 6); r["d21"] = round(r["close"] / r["ema21"] - 1, 6)
    r["d9_atr"] = round((r["close"] - r["ema9"]) / r["atr14"], 6) if r["atr14"] else 0.0
    r["d21_atr"] = round((r["close"] - r["ema21"]) / r["atr14"], 6) if r["atr14"] else 0.0
    r["bar_pos"] = round((r["close"] - r["low"]) / (r["high"] - r["low"]), 6) if r["high"] > r["low"] else 0.5
    r["pb_depth"] = round((r["close"] / r["hi15"] - 1) * 100, 6)
    if hl:
        hl["rise"] = round((hl["last_low"] / hl["prev_low"] - 1) * 100, 6)
    sub = {
        "t_slope": lin(r["slope21"], *LIN["t_slope"]), "t_steps": lin(r["steps21"], *LIN["t_steps"]), "t_rs": lin(r["rs21"], *LIN["t_rs"]),
        "p_near": lin(min(abs(r["d9_atr"]), abs(r["d21_atr"])), *LIN["p_near"]), "p_vol": lin(r["dv_ratio"], *LIN["p_vol"]), "p_bar": r["bar_pos"],
        "h_n": lin(hl.get("streak"), *LIN["h_n"]), "h_rise": lin(hl.get("rise"), *LIN["h_rise"]), "h_age": lin(hl.get("age"), *LIN["h_age"]),
        "r_flat": lin(abs(sh) if sh is not None else None, *LIN["r_flat"]), "r_low": lin(sl, *LIN["r_low"]), "r_cog": lin(tr["cog"], *LIN["r_cog"]),
        "r_touch": 1.0 if tr["touches"] >= 2 else 0.5,
    }
    r["sub"] = {k: round(v, 4) for k, v in sub.items()}
    r["s_trend"] = round(100 * (0.4 * sub["t_slope"] + 0.3 * sub["t_steps"] + 0.3 * sub["t_rs"]), 2)
    r["s_pb"] = round(100 * (0.4 * sub["p_near"] + 0.3 * sub["p_vol"] + 0.3 * sub["p_bar"]), 2)
    r["s_hl"] = round(100 * (0.5 * sub["h_n"] + 0.3 * sub["h_rise"] + 0.2 * sub["h_age"]), 2)
    r["s_tri"] = round(100 * (0.35 * sub["r_flat"] + 0.30 * sub["r_low"] + 0.25 * sub["r_cog"] + 0.10 * sub["r_touch"]), 2)
    r["score"] = round(W["trend"] * r["s_trend"] + W["pb"] * r["s_pb"] + W["hl"] * r["s_hl"] + W["tri"] * r["s_tri"], 3)

# ---------------------------------------------------------------- funnel of the criteria
crit_names = ["C1", "C2", "C3", "C5", "C6", "C7"]
single = {k: sum(1 for r in rows.values() if k not in r["fails"]) for k in crit_names}
cum = []
alive = list(rows.values())
for k in crit_names:
    alive = [r for r in alive if k not in r["fails"]]
    cum.append(len(alive))
tier1 = sorted([r for r in rows.values() if r["tier"] == 1], key=lambda r: -r["score"])
tier2 = sorted([r for r in rows.values() if r["tier"] == 2], key=lambda r: -r["score"])
tier3 = sorted([r for r in rows.values() if r["tier"] == 3], key=lambda r: -r["score"])   # core four pass, only MACD missing (候補)
for i, r in enumerate(tier3, 1):
    r["rank"] = i
for i, r in enumerate(tier1, 1):
    r["rank"] = i
for i, r in enumerate(tier2, 1):
    r["rank"] = i
# the three sheets: subsets of the first tier
for r in rows.values():
    r["in_a"] = r["status"] == "剛黃金交叉"
    closer21 = abs(r["d21"]) <= abs(r["d9"])          # the EMA the close sits nearer to (by %)
    r["in_b"] = r["near21"] and (closer21 or not r["near9"])
    r["in_c"] = r["near9"] and not r["in_b"]
sheets = {k: [r["sym"] for r in tier1 if r[f"in_{k}"]] for k in ("a", "b", "c")}
print("single-criterion passes:", single, "cumulative:", dict(zip(crit_names, cum)))
print(f"tier1 {len(tier1)}  tier2 {len(tier2)}  tier3 (MACD 候補) {len(tier3)}")
print("sheets a/b/c:", {k: len(v) for k, v in sheets.items()}, " b∩c:", len(set(sheets["b"]) & set(sheets["c"])))

# ---------------------------------------------------------------- breadth of the eligible universe (last 10 sessions)
breadth = []
idx_last = T - 1
for back in range(9, -1, -1):
    d = cal[idx_last - back]
    rets, up, ab21, a9o21, e21up, nh, nl = [], 0, 0, 0, 0, 0, 0
    cnt = 0
    for sym in elig:
        s = series[sym]
        try:
            j = s["dates"].index(d)
        except ValueError:
            continue
        if j < 30:
            continue
        c = s["c"]
        cnt += 1
        rets.append(c[j] / c[j - 1] - 1)
        up += c[j] > c[j - 1]
    # EMA-based breadth needs the full series; compute once per symbol
    breadth.append({"date": d, "n": cnt, "median_ret": statistics.median(rets) if rets else None,
                    "pct_up": up / cnt if cnt else None})
ema_b = {d: {"n": 0, "above21": 0, "a9o21": 0, "e21up": 0, "nh20": 0, "nl20": 0} for d in [b["date"] for b in breadth]}
for sym in elig:
    s = series[sym]; c = s["c"]
    e9, e21 = ema(c, 9), ema(c, 21)
    for b in breadth:
        d = b["date"]
        try:
            j = s["dates"].index(d)
        except ValueError:
            continue
        if j < 30:
            continue
        E = ema_b[d]
        E["n"] += 1
        E["above21"] += c[j] > e21[j]
        E["a9o21"] += e9[j] > e21[j]
        E["e21up"] += e21[j] > e21[j - 1]
        E["nh20"] += c[j] >= c[j - 19:j + 1].max()
        E["nl20"] += c[j] <= c[j - 19:j + 1].min()
for b in breadth:
    E = ema_b[b["date"]]
    if E["n"]:
        b.update({"pct_above21": E["above21"] / E["n"], "pct_9over21": E["a9o21"] / E["n"], "pct_e21_rising": E["e21up"] / E["n"],
                  "nh20": E["nh20"], "nl20": E["nl20"]})

sector_counts = defaultdict(int)
for r in tier1:
    sector_counts[r["sector_zh"]] += 1

out = {
    "meta": {"last_date": last_date, "first_date": cal[0], "sessions": T, "symbols": n_sym_all,
             "funnel": funnel, "eligible": len(elig), "single": single, "cumulative": dict(zip(crit_names, cum)),
             "tier1": len(tier1), "tier2": len(tier2), "tier3": len(tier3), "params": P0, "params_loose": PL, "weights": W, "lin": LIN,
             "median_ret21": med21, "sources": YAHOO + SUPP, "snapshot": SNAP,
             "intraday_fill": int((bars[bars.date == last_date].kind == "intraday").sum())},
    "rows": tier1, "tier2": tier2, "tier3": tier3, "sheets": sheets,
    "breadth": breadth, "sector_counts": dict(sector_counts),
    "all_fail_counts": {k: sum(1 for r in rows.values() if k in r["fails"]) for k in crit_names},
}
os.makedirs(os.path.dirname(OUT_JSON) or ".", exist_ok=True)
json.dump(out, open(OUT_JSON, "w"), ensure_ascii=False, indent=1, default=float)
print("wrote", OUT_JSON)
