#!/usr/bin/env python3
"""Build the 9EMA-Uptrend watchlist (R1.00 onward) as an Excel workbook.

Every value the sheet can derive from other cells is a formula: the day's
move, the distances to EMA9 / EMA21 / the last bottom / the 15-day high, the
EMA21 slope, the ATR multiples, every sub-score and the four component scores,
the composite.  The cap-group pages read the 總表 row by reference, so there is
one source per number.  Measured inputs (closes, EMAs, ATR, pivot lows, slopes,
重心) are blue.  Every ticker cell links to its TradingView chart.
"""
import csv, datetime, json, os
import numpy as np
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.comments import Comment

W = os.environ.get("WORK_DIR", "./data")
SCREEN = os.environ.get("SCREEN_JSON", "screen_r1.json")
MARKET = os.environ.get("MARKET_JSON", "market_r1.json")
MACRO = os.environ.get("MACRO", "")
SNAP = os.environ.get("SNAP", "")
OUT = os.environ["OUT_XLSX"]
REV = os.environ.get("REV", "R1.00")
MODEL_TAG = os.environ.get("MODEL_TAG", "")
BUILT_AT = os.environ.get("BUILT_AT", "")

scr = json.load(open(f"{W}/{SCREEN}"))
mkt = json.load(open(f"{W}/{MARKET}")) if os.path.exists(f"{W}/{MARKET}") else {}
rows = scr["rows"]; tier2 = scr["tier2"]; tier3 = scr.get("tier3", []); SHEETS = scr["sheets"]
META = scr["meta"]; LAST = META["last_date"]; P = META["params"]; PL = META["params_loose"]
WT = META["weights"]; LIN = META["lin"]
ALL = rows + tier2 + tier3
T3_FILL = PatternFill("solid", fgColor="FCE4D6")      # third tier: only MACD missing

FONT = "Arial"
HDR_FILL = PatternFill("solid", fgColor="1F3864")
HDR_FONT = Font(name=FONT, bold=True, color="FFFFFF", size=10)
SUB_FILL = PatternFill("solid", fgColor="D9E2F3")
NEW_FILL = PatternFill("solid", fgColor="FFF2CC")
X_FILL = PatternFill("solid", fgColor="E2EFDA")       # fresh golden cross
T2_FILL = PatternFill("solid", fgColor="F2F2F2")      # second tier
WARN_FILL = PatternFill("solid", fgColor="FCE4D6")
THIN = Side(style="thin", color="BFBFBF")
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
BASE = Font(name=FONT, size=10)
BOLD = Font(name=FONT, size=10, bold=True)
TITLE = Font(name=FONT, size=14, bold=True, color="1F3864")
H2 = Font(name=FONT, size=11, bold=True, color="1F3864")
INPUT = Font(name=FONT, size=10, color="0000FF")
NOTE = Font(name=FONT, size=9, italic=True, color="666666")
LINK_FONT = Font(name=FONT, size=10, bold=True, color="0563C1", underline="single")
SRC_FONT = Font(name=FONT, size=9, color="0563C1", underline="single")
WRAP = Alignment(wrap_text=True, vertical="top")
CENTER = Alignment(horizontal="center", vertical="center")

PCT = '0.00%;[Red](0.00%);-'
PCT1 = '0.0%;[Red](0.0%);-'
PX = '$#,##0.00##'
SC = '0.0'
NUM2 = '0.00'
INT = '#,##0'
CAPZH = {"a": "大型", "b": "中型", "c": "小型", "x": "未分類"}
CRIT_ZH = {"C1": "C1 21EMA向上", "C2": "C2 9EMA在21EMA上", "C3": "C3 一底高於一底", "C5": "C4 上升三角／重心偏上",
           "C6": "C5 MACD慢線向上", "C7": "C6 MACD柱轉勢"}
SHEET_ZH = {"a": "a 剛黃金交叉", "b": "b 回到21EMA", "c": "c 回到9EMA"}


def tv_url(sym, ex):
    s2 = sym.replace("/", ".").lower()
    return (f"https://www.tradingview.com/chart/Q1c5VWwD/?symbol={ex.lower()}%3A{s2}" if ex and ex != "—"
            else f"https://www.tradingview.com/chart/Q1c5VWwD/?symbol={s2}")


def link_ticker(ws_, row_, col_, sym, ex):
    c = ws_.cell(row=row_, column=col_, value=sym)
    c.hyperlink = tv_url(sym, ex); c.font = LINK_FONT; c.border = BOX
    return c


def style_header(ws, headers, widths, row=1, freeze_at=None, height=42):
    for j, (h, wd) in enumerate(zip(headers, widths), 1):
        c = ws.cell(row=row, column=j, value=h)
        c.fill, c.font = HDR_FILL, HDR_FONT
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = BOX
        ws.column_dimensions[get_column_letter(j)].width = wd
    ws.row_dimensions[row].height = height
    ws.freeze_panes = freeze_at or f"D{row + 1}"
    ws.auto_filter.ref = f"A{row}:{get_column_letter(len(headers))}{row}"


def put(ws, r, c, v, fmt=None, font=None, fill=None, align=None, border=True):
    cell = ws.cell(row=r, column=c, value=v)
    cell.font = font or BASE
    if border:
        cell.border = BOX
    if fmt: cell.number_format = fmt
    if fill: cell.fill = fill
    if align: cell.alignment = align
    return cell


def lin_f(expr, x0, x1):
    """Piecewise-linear 0..1 clip of a cell expression, x0 -> 0, x1 -> 1."""
    return f"MIN(1,MAX(0,(({expr})-({x0}))/(({x1})-({x0}))))"


def num_or0(ref):
    return f"IF(ISNUMBER({ref}),{ref},0)"


def fmt_date(d):
    return d or ""


# ================================================================ 總表
wb = Workbook()
ws = wb.active
ws.title = "總表"
COLS = [
    ("tier", "梯隊", 5), ("rank", "排名", 5), ("sym", "代號", 8), ("name", "公司", 26), ("exch", "交易所", 7),
    ("sec", "板塊", 9), ("ind", "行業", 24), ("mcap", "市值(十億美元)", 9), ("cap", "市值組", 6),
    ("close", "收市價", 9), ("prev", "前收", 9), ("chg", "當日%", 7), ("high", "當日高", 9), ("low", "當日低", 9),
    ("vol", "成交量", 10), ("vol20", "20日均量", 10), ("vr", "量比", 6),
    ("ema9", "EMA9", 9), ("ema21", "EMA21", 9), ("ema21l", "EMA21（5日前）", 9), ("slope", "EMA21 5日斜率%", 7), ("steps", "EMA21 10步上升比例", 7),
    ("d9", "距EMA9%", 7), ("d21", "距EMA21%", 7), ("atr", "ATR14", 8), ("atrp", "ATR14%", 7),
    ("d9a", "距EMA9（ATR倍）", 7), ("d21a", "距EMA21（ATR倍）", 7), ("near", "最接近", 7), ("touch", "近3日低位觸及", 8),
    ("status", "9/21狀態", 9), ("in_a", "a 剛金叉", 6), ("in_b", "b 回到21EMA", 7), ("in_c", "c 回到9EMA", 7), ("xdate", "金叉日期", 10), ("xdays", "金叉後日數", 6),
    ("hln", "連續上升底數目", 6), ("hldate", "最後一底日期", 10), ("hllow", "最後一底價", 9), ("hlprev", "上一底價", 9),
    ("hlrise", "底較底升%", 7), ("hlage", "最後一底距今（日）", 7), ("dlow", "距最後一底%", 7),
    ("hi15", "15日最高收市", 9), ("pb", "距15日高%", 7),
    ("macd", "MACD", 8), ("sig", "MACD慢線", 8), ("sigp", "慢線（前日）", 8), ("sigd", "慢線斜率", 8),
    ("hist", "柱（今日）", 8), ("histp", "柱（前日）", 8), ("histp2", "柱（前2日）", 8), ("hcol", "柱顏色", 6), ("turn", "柱轉勢", 9), ("tago", "轉勢幾日前", 6), ("hseq", "柱顏色序列（舊→新）", 26),
    ("pat", "形態（25日）", 13), ("slh", "高位斜率（%/日）", 8), ("sll", "低位斜率（%/日）", 8), ("cog", "重心（VWAP在25日高低區間位置）", 9),
    ("cogc", "重心（收市範圍）", 8), ("touches", "頂部觸及次數", 6), ("whi", "25日最高", 9), ("wlo", "25日最低", 9), ("vwap", "25日VWAP", 9),
    ("dvr", "跌日／升日量比（10日）", 8), ("bpos", "收市在當日區間位置", 8), ("rs", "RS21（對全池中位數）%", 8), ("ret21", "21日回報%", 7),
    ("s_tr", "趨勢分數", 7), ("s_pb", "回調質素", 7), ("s_hl", "底部結構", 7), ("s_tri", "三角形分數", 7), ("score", "綜合分數", 7),
    ("fails", "唔過嘅條件", 12), ("note", "備註", 28),
]
L = {k: get_column_letter(i) for i, (k, _, _) in enumerate(COLS, 1)}
CI = {k: i for i, (k, _, _) in enumerate(COLS, 1)}
style_header(ws, [h for _, h, _ in COLS], [w for _, _, w in COLS], freeze_at="D2")

INPUT_KEYS = {"macd", "sig", "sigp", "hist", "histp", "histp2", "mcap", "close", "prev", "high", "low", "vol", "vol20", "ema9", "ema21", "ema21l", "steps", "atr",
              "hln", "hllow", "hlprev", "hlage", "hi15", "slh", "sll", "cog", "cogc", "touches", "whi", "wlo", "vwap", "dvr", "rs", "ret21"}
FMT = {"macd": '0.000', "sig": '0.000', "sigp": '0.000', "sigd": '+0.000;[Red]-0.000', "hist": '0.000', "histp": '0.000', "histp2": '0.000', "mcap": '#,##0.00', "close": PX, "prev": PX, "high": PX, "low": PX, "vol": INT, "vol20": INT, "vr": NUM2,
       "ema9": PX, "ema21": PX, "ema21l": PX, "slope": PCT, "steps": '0%', "d9": PCT, "d21": PCT, "atr": PX, "atrp": PCT1,
       "d9a": NUM2, "d21a": NUM2, "hllow": PX, "hlprev": PX, "hlrise": PCT, "dlow": PCT, "hi15": PX, "pb": PCT,
       "slh": NUM2, "sll": NUM2, "cog": NUM2, "cogc": NUM2, "whi": PX, "wlo": PX, "vwap": PX, "dvr": NUM2, "bpos": NUM2,
       "rs": '0.0', "ret21": '0.0', "s_tr": SC, "s_pb": SC, "s_hl": SC, "s_tri": SC, "score": SC, "chg": PCT}


def measure_text(r, k):
    hl = r["hl"] or {}; tr = r["tri"]
    if k == "C1":
        return f"EMA21 5日斜率 {r['slope21']:+.2f}%、EMA21 比前一日 {'高' if r['ema21'] > r['ema21_prev'] else '低'}（要兩者都向上）"
    if k == "C2":
        return f"EMA9 {r['ema9']:.2f} vs EMA21 {r['ema21']:.2f}（要 EMA9 高過 EMA21）"
    if k == "C3":
        if not hl:
            return f"近 {P['hl_look']} 日不足兩個樞軸底（要最後兩個底遞升、最後一底 ≤{P['hl_recent']} 日前、其後收市冇跌穿）"
        return (f"最後兩底 {hl['prev_low']:.2f} → {hl['last_low']:.2f}（{hl['rise']:+.1f}%），最後一底 {hl['age']} 日前"
                f"（要遞升、≤{P['hl_recent']} 日前、其後收市冇跌穿）")
    if k == "C4":
        return (f"距EMA9 {r['d9']*100:+.2f}%（{r['d9_atr']:+.2f} ATR）、距EMA21 {r['d21']*100:+.2f}%（{r['d21_atr']:+.2f} ATR）、"
                f"近3日觸及 9EMA {'有' if r['touch3_9'] else '無'}／21EMA {'有' if r['touch3_21'] else '無'}"
                f"（要離其中一條 ≤{P['near_pct']*100:.0f}% 兼 ≤{P['near_atr']} ATR 或近 {P['touch_days']} 日觸及；收市唔低過 EMA21 {abs(P['under_tol'])*100:.0f}%）")
    if k == "C6":
        return f"MACD 慢線 {r['sig_prev']:.3f} → {r['sig']:.3f}（要升）"
    if k == "C7":
        ago = "" if r["turn_ago"] is None else f" {r['turn_ago']} 日前"
        return (f"柱 {r['hist_prev2']:+.3f} → {r['hist_prev']:+.3f} → {r['hist']:+.3f}（{r['hist_seq']}）；轉勢 {r['turn_kind']}{ago}"
                f"（要 ≤{P['hist_days']} 日內出現第一根淺紅／深綠，其後每根都升）")
    if k == "C5":
        return (f"重心 {tr['cog']:.2f}（要 ≥{P['cog_min']}）、高位斜率 {tr['sl_hi']:+.2f}%/日（要 {P['high_slope_min']:+.2f} 至 {P['high_slope_max']:+.2f}）、"
                f"低位斜率 {tr['sl_lo']:+.2f}%/日（要 >0）")
    return ""


rr = 2
ROW_OF = {}
for r in ALL:
    hl = r["hl"] or {}; tr = r["tri"]
    t2 = r["tier"] == 2
    touch_txt = ("9+21EMA" if (r["touch3_9"] and r["touch3_21"]) else "9EMA" if r["touch3_9"] else "21EMA" if r["touch3_21"] else "—")
    vals = {
        "tier": r["tier"], "rank": r["rank"], "sym": r["sym"], "name": r["name"], "exch": r["exch"], "sec": r["sector_zh"], "ind": r["industry"],
        "mcap": round(r["mcap"] / 1e9, 3) if r["mcap"] else None, "cap": CAPZH.get(r["cap"], r["cap"]),
        "close": r["close"], "prev": r["prev"], "high": r["high"], "low": r["low"], "vol": r["vol"], "vol20": round(r["vol20"]),
        "ema9": r["ema9"], "ema21": r["ema21"], "ema21l": r["ema21_lag"], "steps": r["steps21"], "atr": r["atr14"],
        "touch": touch_txt, "status": r["status"], "xdate": fmt_date(r["cross_date"]), "xdays": r["above_days"],
        "in_a": "✓" if r["in_a"] else "", "in_b": "✓" if r["in_b"] else "", "in_c": "✓" if r["in_c"] else "",
        "hln": hl.get("streak", 0), "hldate": hl.get("last_date", ""), "hllow": hl.get("last_low"), "hlprev": hl.get("prev_low"), "hlage": hl.get("age"),
        "hi15": r["hi15"], "pat": r["pattern"],
        "macd": r["macd"], "sig": r["sig"], "sigp": r["sig_prev"], "hist": r["hist"], "histp": r["hist_prev"], "histp2": r["hist_prev2"],
        "hcol": r["hist_colour"], "turn": r["turn_kind"], "tago": r["turn_ago"] if r["turn_ago"] is not None else "—", "hseq": r["hist_seq"], "slh": tr["sl_hi"], "sll": tr["sl_lo"], "cog": tr["cog"], "cogc": tr["cog_c"], "touches": tr["touches"],
        "whi": tr["hi"], "wlo": tr["lo"], "vwap": tr["vwap"], "dvr": r["dv_ratio"], "rs": r["rs21"], "ret21": r["ret21"],
        "fails": "、".join(CRIT_ZH[k] for k in r["fails"]) if r["fails"] else "全部通過",
        "note": ("日內滾算收市（Yahoo 未出日線）" if r["bar_kind"] == "intraday" else ""),
    }
    for k, v in vals.items():
        c = put(ws, rr, CI[k], v, FMT.get(k), INPUT if k in INPUT_KEYS else BASE)
    link_ticker(ws, rr, CI["sym"], r["sym"], r["exch"])
    # formulas
    R = {k: f"{L[k]}{rr}" for k in L}        # cell address of every column on this row
    f = {
        "chg": f"={R['close']}/{R['prev']}-1",
        "vr": f"={R['vol']}/{R['vol20']}",
        "slope": f"={R['ema21']}/{R['ema21l']}-1",
        "d9": f"={R['close']}/{R['ema9']}-1",
        "d21": f"={R['close']}/{R['ema21']}-1",
        "atrp": f"={R['atr']}/{R['close']}",
        "d9a": f"=({R['close']}-{R['ema9']})/{R['atr']}",
        "d21a": f"=({R['close']}-{R['ema21']})/{R['atr']}",
        "near": f"=IF(ABS({R['d9a']})<=ABS({R['d21a']}),\"9EMA\",\"21EMA\")",
        "hlrise": f"=IF(ISNUMBER({R['hlprev']}),{R['hllow']}/{R['hlprev']}-1,\"\")",
        "dlow": f"=IF(ISNUMBER({R['hllow']}),{R['close']}/{R['hllow']}-1,\"\")",
        "pb": f"={R['close']}/{R['hi15']}-1",
        "sigd": f"={R['sig']}-{R['sigp']}",
        "bpos": f"=IF({R['high']}={R['low']},0.5,({R['close']}-{R['low']})/({R['high']}-{R['low']}))",
        "s_tr": ("=100*(0.4*" + lin_f(R["slope"] + "*100", *LIN["t_slope"])
                 + "+0.3*" + lin_f(R["steps"], *LIN["t_steps"])
                 + "+0.3*" + lin_f(R["rs"], *LIN["t_rs"]) + ")"),
        "s_pb": ("=100*(0.4*" + lin_f(f"MIN(ABS({R['d9a']}),ABS({R['d21a']}))", *LIN["p_near"])
                 + "+0.3*" + lin_f(R["dvr"], *LIN["p_vol"])
                 + "+0.3*" + R["bpos"] + ")"),
        "s_hl": ("=100*(0.5*" + lin_f(R["hln"], *LIN["h_n"])
                 + "+0.3*" + lin_f(num_or0(R["hlrise"]) + "*100", *LIN["h_rise"])
                 + "+0.2*" + lin_f(num_or0(R["hlage"]), *LIN["h_age"]) + ")"),
        "s_tri": ("=100*(0.35*" + lin_f(f"ABS({R['slh']})", *LIN["r_flat"])
                  + "+0.3*" + lin_f(R["sll"], *LIN["r_low"])
                  + "+0.25*" + lin_f(R["cog"], *LIN["r_cog"])
                  + f"+0.1*IF({R['touches']}>=2,1,0.5))"),
        "score": f"={WT['trend']}*{R['s_tr']}+{WT['pb']}*{R['s_pb']}+{WT['hl']}*{R['s_hl']}+{WT['tri']}*{R['s_tri']}",
    }
    for k, v in f.items():
        put(ws, rr, CI[k], v, FMT.get(k), BOLD if k == "score" else BASE)
    if t2 or r["tier"] == 3:
        for j in range(1, len(COLS) + 1):
            ws.cell(row=rr, column=j).fill = T2_FILL if t2 else T3_FILL
        ws.cell(row=rr, column=CI["fails"]).comment = Comment(
            "\n".join(measure_text(r, k) for k in r["fails"]), "9EMA scan", width=520, height=120)
    if r["status"] == "剛黃金交叉":
        ws.cell(row=rr, column=CI["status"]).fill = X_FILL
    ROW_OF[r["sym"]] = rr
    rr += 1

n1, n2 = len(rows), len(tier2)
last_row = rr - 1
srow = rr + 1
put(ws, srow, 1, "全表統計", font=BOLD, border=False)
stats = [
    ("第一梯隊", f"=COUNTIF({L['tier']}2:{L['tier']}{last_row},1)", None),
    ("第二梯隊", f"=COUNTIF({L['tier']}2:{L['tier']}{last_row},2)", None),
    ("第三梯隊（只差MACD）", f"=COUNTIF({L['tier']}2:{L['tier']}{last_row},3)", None),
    ("剛黃金交叉", f"=COUNTIF({L['status']}2:{L['status']}{last_row},\"剛黃金交叉\")", None),
    ("b 回到21EMA", f"=COUNTIFS({L['tier']}2:{L['tier']}{last_row},1,{L['in_b']}2:{L['in_b']}{last_row},\"✓\")", None),
    ("c 回到9EMA", f"=COUNTIFS({L['tier']}2:{L['tier']}{last_row},1,{L['in_c']}2:{L['in_c']}{last_row},\"✓\")", None),
    ("上升三角（平頂）", f"=COUNTIF({L['pat']}2:{L['pat']}{last_row},\"上升三角（平頂）\")", None),
    ("當日中位%", f"=MEDIAN({L['chg']}2:{L['chg']}{n1 + 1})", PCT),
    ("距EMA9 中位%", f"=MEDIAN({L['d9']}2:{L['d9']}{n1 + 1})", PCT),
    ("距EMA21 中位%", f"=MEDIAN({L['d21']}2:{L['d21']}{n1 + 1})", PCT),
    ("低過EMA9 行數", f"=COUNTIF({L['d9']}2:{L['d9']}{n1 + 1},\"<0\")", None),
    ("綜合分數 中位", f"=MEDIAN({L['score']}2:{L['score']}{n1 + 1})", SC),
    ("RS21 中位%", f"=MEDIAN({L['rs']}2:{L['rs']}{n1 + 1})", '0.0'),
    ("大型股", f"=COUNTIFS({L['tier']}2:{L['tier']}{last_row},1,{L['cap']}2:{L['cap']}{last_row},\"大型\")", None),
    ("中型股", f"=COUNTIFS({L['tier']}2:{L['tier']}{last_row},1,{L['cap']}2:{L['cap']}{last_row},\"中型\")", None),
    ("小型股", f"=COUNTIFS({L['tier']}2:{L['tier']}{last_row},1,{L['cap']}2:{L['cap']}{last_row},\"小型\")", None),
]
for j, (lbl, formula, fmt) in enumerate(stats, 2):
    put(ws, srow, j, lbl, font=Font(name=FONT, size=9, italic=True), border=False, align=Alignment(wrap_text=True))
    put(ws, srow + 1, j, formula, fmt, BOLD, SUB_FILL)
ws.row_dimensions[srow].height = 28
put(ws, srow + 3, 1,
    (f"口徑：以 {LAST} 收市嘅 Yahoo 日線計；EMA 用標準指數平滑（α=2/(N+1)），ATR14 用 Wilder 平滑。藍字＝量度值（輸入），黑字＝公式；"
     f"當日%、量比、EMA21 斜率、距 EMA／底／15 日高、ATR 倍數、收市區間位置、四個分數同綜合分數全部係公式，改動輸入會自動重算。"
     f"灰底＝第二梯隊（只差一項，「唔過嘅條件」格嘅註解列明量度值同門檻）；綠底狀態＝9EMA 喺 {P['cross_days']} 日內剛黃金交叉 21EMA。"
     f"a／b／c 三欄 ✓ ＝ 該股喺對應分頁（b 同 c 互斥：離兩條 EMA 都近就歸較近嗰條）。橙底＝第三梯隊：價格四項全過、只差 MACD 慢線／柱轉勢（候補，等第一根淺紅／深綠）。"
     f"RS21 同全池中位數 21 日回報（{META['median_ret21']*100:+.2f}%）比較。"),
    font=NOTE, border=False)
ws.cell(row=srow + 3, column=1).alignment = Alignment(wrap_text=False)

# ================================================================ a／b／c 三頁（引用總表）
ABC_COLS = [("rank_in", "頁內排名", 6), ("sym", "代號", 8), ("name", "公司", 26), ("sec", "板塊", 9), ("ind", "行業", 24), ("cap", "市值組", 6), ("mcap", "市值(十億美元)", 9),
            ("close", "收市價", 9), ("chg", "當日%", 7), ("ema9", "EMA9", 9), ("ema21", "EMA21", 9), ("d9", "距EMA9%", 7), ("d21", "距EMA21%", 7),
            ("d9a", "距EMA9（ATR倍）", 7), ("d21a", "距EMA21（ATR倍）", 7), ("touch", "近3日低位觸及", 8), ("slope", "EMA21 5日斜率%", 7),
            ("status", "9/21狀態", 9), ("xdate", "金叉日期", 10), ("xdays", "金叉後日數", 6), ("hln", "連續上升底數目", 6), ("hldate", "最後一底日期", 10), ("dlow", "距最後一底%", 7),
            ("pb", "距15日高%", 7), ("pat", "形態（25日）", 13), ("cog", "重心", 7), ("rs", "RS21%", 7),
            ("hcol", "柱顏色", 6), ("turn", "柱轉勢", 9), ("tago", "轉勢幾日前", 6), ("sigd", "慢線斜率", 8),
            ("s_tr", "趨勢分數", 7), ("s_pb", "回調質素", 7), ("s_hl", "底部結構", 7), ("s_tri", "三角形分數", 7), ("score", "綜合分數", 7),
            ("tier", "梯隊", 5), ("rank", "總表排名", 6), ("fails", "唔過嘅條件", 14)]
ABC_NOTE = {
    "a": f"核心六條件全過，而且 9EMA 喺最近 {P['cross_days']} 個交易日內剛升穿 21EMA（金叉日期欄）。",
    "b": f"核心六條件全過，而且收市離 EMA21 ≤{P['near_pct']*100:.0f}% 兼（≤{P['near_atr']} ATR14 或近 {P['touch_days']} 日最低價觸及 EMA21）、收市唔低過 EMA21 {abs(P['under_tol'])*100:.0f}%；離兩條 EMA 都近嘅股票歸較近嗰條（呢頁係較深嘅回調）。",
    "c": f"核心六條件全過，而且收市離 EMA9 ≤{P['near_pct']*100:.0f}% 兼（≤{P['near_atr']} ATR14 或近 {P['touch_days']} 日最低價觸及 EMA9），而且唔喺 b 頁（呢頁係較淺嘅回調）。",
}
for key in ("a", "b", "c"):
    wp = wb.create_sheet(SHEET_ZH[key])
    style_header(wp, [h for _, h, _ in ABC_COLS], [w for _, _, w in ABC_COLS], freeze_at="C2")
    rp = 2
    for k, r in enumerate([x for x in ALL if x[f"in_{key}"]], 1):
        src = ROW_OF[r["sym"]]
        for j, (ck, _, _) in enumerate(ABC_COLS, 1):
            if ck == "rank_in":
                put(wp, rp, j, k)
            elif ck == "sym":
                link_ticker(wp, rp, j, r["sym"], r["exch"])
            else:
                put(wp, rp, j, f"='總表'!{L[ck]}{src}", FMT.get(ck), BOLD if ck == "score" else BASE)
        if r["tier"] in (2, 3):
            for j in range(1, len(ABC_COLS) + 1):
                wp.cell(row=rp, column=j).fill = T2_FILL if r["tier"] == 2 else T3_FILL
        elif r["status"] == "剛黃金交叉":
            wp.cell(row=rp, column=[i for i, (kk, _, _) in enumerate(ABC_COLS, 1) if kk == "status"][0]).fill = X_FILL
        rp += 1
    n_t1 = sum(1 for x in rows if x[f"in_{key}"]); n_t2 = sum(1 for x in tier2 if x[f"in_{key}"]); n_t3 = sum(1 for x in tier3 if x[f"in_{key}"])
    put(wp, rp + 1, 1, ABC_NOTE[key] + f" 第一梯隊 {n_t1} 隻；灰底＝第二梯隊（只差一項）{n_t2} 隻；橙底＝第三梯隊（價格四項過、只差 MACD 慢線／柱轉勢，候補）{n_t3} 隻；「唔過嘅條件」欄列明。每格引用總表同一行，改總表會跟住變。", font=NOTE, border=False)

# ================================================================ 大型／中型／小型
PAGE_COLS = [("rank_in", "組內排名", 6), ("sym", "代號", 8), ("name", "公司", 26), ("sec", "板塊", 9), ("ind", "行業", 24), ("mcap", "市值(十億美元)", 9),
             ("close", "收市價", 9), ("chg", "當日%", 7), ("ema9", "EMA9", 9), ("ema21", "EMA21", 9), ("d9", "距EMA9%", 7), ("d21", "距EMA21%", 7),
             ("slope", "EMA21 5日斜率%", 7), ("status", "9/21狀態", 9), ("xdays", "金叉後日數", 6), ("hln", "連續上升底數目", 6), ("pb", "距15日高%", 7),
             ("macd", "MACD", 8), ("sig", "MACD慢線", 8), ("sigp", "慢線（前日）", 8), ("sigd", "慢線斜率", 8),
    ("hist", "柱（今日）", 8), ("histp", "柱（前日）", 8), ("histp2", "柱（前2日）", 8), ("hcol", "柱顏色", 6), ("turn", "柱轉勢", 9), ("tago", "轉勢幾日前", 6), ("hseq", "柱顏色序列（舊→新）", 26),
    ("pat", "形態（25日）", 13), ("cog", "重心", 7), ("rs", "RS21%", 7), ("s_tr", "趨勢分數", 7), ("s_pb", "回調質素", 7), ("s_hl", "底部結構", 7),
             ("s_tri", "三角形分數", 7), ("score", "綜合分數", 7), ("rank", "總表排名", 6)]
for cap, title in (("a", "大型股"), ("b", "中型股"), ("c", "小型股")):
    wp = wb.create_sheet(title)
    style_header(wp, [h for _, h, _ in PAGE_COLS], [w for _, _, w in PAGE_COLS], freeze_at="C2")
    rp = 2
    for k, r in enumerate([x for x in rows if x["cap"] == cap], 1):
        src = ROW_OF[r["sym"]]
        for j, (key, _, _) in enumerate(PAGE_COLS, 1):
            if key == "rank_in":
                put(wp, rp, j, k)
            elif key == "sym":
                link_ticker(wp, rp, j, r["sym"], r["exch"])
            else:
                put(wp, rp, j, f"='總表'!{L[key]}{src}", FMT.get(key), BOLD if key == "score" else BASE)
        if r["status"] == "剛黃金交叉":
            wp.cell(row=rp, column=[i for i, (kk, _, _) in enumerate(PAGE_COLS, 1) if kk == "status"][0]).fill = X_FILL
        rp += 1
    lim = {"a": "市值 ≥ $100 億", "b": "市值 $20–100 億", "c": "市值 < $20 億"}[cap]
    put(wp, rp + 1, 1, f"{title}（{lim}）：第一梯隊入面屬呢個市值組嘅股票，每格都係引用總表同一行，改總表會跟住變。", font=NOTE, border=False)

# ================================================================ 差一項
w2 = wb.create_sheet("差一項")
D2 = [("rank", "梯隊2排名", 6), ("sym", "代號", 8), ("name", "公司", 26), ("sec", "板塊", 9), ("cap", "市值組", 6), ("fail", "唔過嘅條件", 18),
      ("meas", "量度值 vs 嚴格門檻", 90), ("loose", "放寬門檻（已通過）", 60), ("score", "綜合分數", 7), ("row", "總表行", 6)]
style_header(w2, [h for _, h, _ in D2], [w for _, _, w in D2], freeze_at="C2", height=30)
LOOSE_TXT = {
    "C1": "同嚴格",
    "C2": "同嚴格",
    "C3": f"回望 {PL['hl_look']} 日、最後一底 ≤{PL['hl_recent']} 日前",
    "C5": f"重心 ≥{PL['cog_min']}、高位斜率 {PL['high_slope_min']:+.2f} 至 {PL['high_slope_max']:+.2f} %/日、低位斜率 >0",
    "C6": "同嚴格", "C7": f"轉勢 ≤{PL['hist_days']} 日內",
}
rd = 2
for r in tier2:
    k = r["fails"][0]
    src = ROW_OF[r["sym"]]
    put(w2, rd, 1, r["rank"]); link_ticker(w2, rd, 2, r["sym"], r["exch"]); put(w2, rd, 3, r["name"]); put(w2, rd, 4, r["sector_zh"]); put(w2, rd, 5, CAPZH.get(r["cap"], r["cap"]))
    put(w2, rd, 6, CRIT_ZH[k]); put(w2, rd, 7, measure_text(r, k), align=WRAP); put(w2, rd, 8, LOOSE_TXT[k], align=WRAP)
    put(w2, rd, 9, f"='總表'!{L['score']}{src}", SC, BOLD); put(w2, rd, 10, src)
    rd += 1
put(w2, rd + 1, 1, f"第二梯隊（{len(tier2)} 隻）：五項入面只唔過一項，而且嗰一項喺放寬門檻下通過。綜合分數引用總表。", font=NOTE, border=False)

# ================================================================ 樞軸底頂
wv = wb.create_sheet("樞軸底頂")
PV = [("tier", "梯隊", 5), ("rank", "排名", 5), ("sym", "代號", 8), ("name", "公司", 24), ("close", "收市價", 9)]
PV += [(f"lo{i}", f"底{i}（{P['hl_look']}日內，舊→新）", 15) for i in range(1, 7)]
PV += [(f"hi{i}", f"頂{i}（{P['tri_win']}日內，舊→新）", 15) for i in range(1, 5)]
style_header(wv, [h for _, h, _ in PV], [w for _, _, w in PV], freeze_at="D2", height=30)
rv = 2
for r in ALL:
    put(wv, rv, 1, r["tier"]); put(wv, rv, 2, r["rank"]); link_ticker(wv, rv, 3, r["sym"], r["exch"]); put(wv, rv, 4, r["name"]); put(wv, rv, 5, r["close"], PX, INPUT)
    los = r["piv_lows"][-6:]; his = r["piv_highs"][-4:]
    for i in range(6):
        put(wv, rv, 6 + i, f"{los[i][0][5:]} @ {los[i][1]:.2f}" if i < len(los) else "", font=INPUT)
    for i in range(4):
        put(wv, rv, 12 + i, f"{his[i][0][5:]} @ {his[i][1]:.2f}" if i < len(his) else "", font=INPUT)
    if r["tier"] == 2:
        for j in range(1, len(PV) + 1):
            wv.cell(row=rv, column=j).fill = T2_FILL
    rv += 1
put(wv, rv + 1, 1,
    (f"樞軸底 = 當日最低價係前後各 {P['piv']} 個交易日入面最低（兩邊最外嗰日都要高過佢），{P['piv']} 日內嘅兩個底只留較低嗰個；"
     f"樞軸頂同理用最高價。C3 要最後兩個底遞升；C5 嘅高位／低位斜率係呢啲樞軸對日數嘅最小二乘斜率（不足兩個樞軸就用窗口內全部高／低位）。"),
    font=NOTE, border=False)

# ================================================================ 篩選規則
wr = wb.create_sheet("篩選規則")
wr.column_dimensions["A"].width = 24; wr.column_dimensions["B"].width = 110; wr.column_dimensions["C"].width = 16; wr.column_dimensions["D"].width = 14
put(wr, 1, 1, f"{REV} 篩選條件：21EMA 向上 ＋ 9EMA 在 21EMA 上 ＋ 一底高於一底 ＋ 上升三角（重心偏上）＋ MACD 慢線向上 ＋ MACD 柱轉勢；分頁 a 剛黃金交叉／b 回到 21EMA／c 回到 9EMA", font=TITLE, border=False)
put(wr, 2, 1, "說明", font=BOLD)
put(wr, 2, 2, (f"全美上市普通股（Nasdaq／NYSE／AMEX，{META['symbols']:,} 隻）用 Yahoo 日線量度，數據終點 {LAST} 收市（{META['sessions']} 個交易日，{META['first_date']} 起）。"
               f"EMA 用標準指數平滑 α=2/(N+1)，ATR14 用 Wilder 平滑。單獨通過＝只計該項；累計＝由 C1 起逐項收窄。"), align=WRAP)
put(wr, 2, 3, "單獨通過", font=BOLD, align=CENTER); put(wr, 2, 4, "累計通過", font=BOLD, align=CENTER)
fn = META["funnel"]
rules = [
    ("股票池", f"當日（{LAST}）有成交、≥{P['min_hist']} 個交易日歷史、收市 ≥${P['min_px']:.0f}、20 日成交額中位數 ≥${P['min_dv']/1e6:.0f}M"
              f"（{fn['symbols']:,} → 有當日 bar {fn['bar_on_last']:,} → 歷史夠 {fn['hist']:,} → 價格 {fn['price']:,} → 成交額 {fn['liq']:,}）", META["eligible"], META["eligible"]),
    ("C1 21EMA 向上", f"EMA21 高過前一日，而且高過 {P['slope_lag']} 日前（5 日斜率 >0）；另記錄最近 {P['steps_look']} 步入面上升嘅比例（入分數）", META["single"]["C1"], META["cumulative"]["C1"]),
    ("C2 9EMA 在 21EMA 上", f"EMA9 > EMA21。最近一次黃金交叉（EMA9 由 ≤ 變 >）喺 {P['cross_days']} 日內＝「剛黃金交叉」（a 頁），否則「持續在上方」",
     META["single"]["C2"], META["cumulative"]["C2"]),
    ("C3 一底高於一底", f"用最低價搵樞軸底（前後各 {P['piv']} 日最低）；近 {P['hl_look']} 日最後兩個底要遞升、最後一底 ≤{P['hl_recent']} 日前、之後收市冇跌穿佢（容 0.5%）。連續遞升底嘅數目入分數",
     META["single"]["C3"], META["cumulative"]["C3"]),
    ("C4 上升三角／重心偏上", f"近 {P['tri_win']} 日：VWAP 喺高低區間嘅位置（重心）≥{P['cog_min']}；樞軸低位嘅最小二乘斜率 >0；樞軸高位斜率喺 {P['high_slope_min']:+.2f} 至 {P['high_slope_max']:+.2f} %/日之間"
                          f"（平頂至微升：頂部下傾係對稱三角、頂部急升係通道，都唔係上升三角）", META["single"]["C5"], META["cumulative"]["C5"]),
    ("C5 MACD 慢線向上", f"MACD（EMA{P['macd_fast']} − EMA{P['macd_slow']}）嘅慢線（訊號線，MACD 嘅 EMA{P['macd_sig']}）高過前一日", META["single"]["C6"], META["cumulative"]["C6"]),
    ("C6 MACD 柱轉勢", f"柱 = MACD − 慢線，顏色跟 TradingView：柱 ≥0 升＝深綠、跌＝淺綠；柱 <0 跌＝深紅、升＝淺紅。要近 {P['hist_days']} 日內出現第一根「深紅之後嘅淺紅」（跌勢減速）或「淺綠之後嘅深綠」（升勢再加速），而且由嗰根起每根都升",
     META["single"]["C7"], META["cumulative"]["C7"]),
    ("第一梯隊（總表）", "六項全過，按綜合分數排", "", META["tier1"]),
    ("a 剛黃金交叉", f"第一梯隊（另加灰底第二梯隊）入面 9EMA 喺 {P['cross_days']} 日內剛升穿 21EMA", "", len(SHEETS["a"])),
    ("b 回到 21EMA", f"第一梯隊入面收市離 EMA21 ≤{P['near_pct']*100:.0f}% 兼（≤{P['near_atr']} ATR14 或近 {P['touch_days']} 日最低價 ≤ 當日 EMA21×{1+P['touch_tol']:.3f}）、收市唔低過 EMA21 {abs(P['under_tol'])*100:.0f}%；離兩條都近就歸較近嗰條", "", len(SHEETS["b"])),
    ("c 回到 9EMA", f"第一梯隊入面收市離 EMA9 ≤{P['near_pct']*100:.0f}% 兼（≤{P['near_atr']} ATR14 或近 {P['touch_days']} 日最低價 ≤ 當日 EMA9×{1+P['touch_tol']:.3f}），而且唔喺 b 頁", "", len(SHEETS["c"])),
    ("第三梯隊（候補）", "價格四項（C1–C4）全過、只差 MACD 慢線／柱轉勢；總表橙底，分頁都列埋，等第一根淺紅／深綠出現", "", META.get("tier3", 0)),
    ("第二梯隊", f"只唔過一項，而且嗰項喺放寬門檻下通過（金叉 ≤{PL['cross_days']} 日；底回望 {PL['hl_look']} 日／最後一底 ≤{PL['hl_recent']} 日；"
                f"重心 ≥{PL['cog_min']}、高位斜率 {PL['high_slope_min']:+.2f} 至 {PL['high_slope_max']:+.2f}；柱轉勢 ≤{PL['hist_days']} 日）；「差一項」頁列明每隻差邊項", "", META["tier2"]),
]
rw_ = 3
for a, b, c, d in rules:
    put(wr, rw_, 1, a, font=BOLD, align=WRAP); put(wr, rw_, 2, b, align=WRAP); put(wr, rw_, 3, c, align=CENTER); put(wr, rw_, 4, d, align=CENTER)
    wr.row_dimensions[rw_].height = 44
    rw_ += 1
rw_ += 1
put(wr, rw_, 1, "評分（每項 0–1，線性、兩頭封頂；總表公式）", font=H2, border=False); rw_ += 1
scoring = [
    ("趨勢分數", f"0.4 × EMA21 5日斜率（{LIN['t_slope'][0]}%→0、{LIN['t_slope'][1]}%→1）+ 0.3 × EMA21 10步上升比例（{LIN['t_steps'][0]:.0%}→0、{LIN['t_steps'][1]:.0%}→1）+ 0.3 × RS21（{LIN['t_rs'][0]:.0f}%→0、{LIN['t_rs'][1]:.0f}%→1；RS21 = 21 日回報 − 全池中位數）"),
    ("回調質素", f"0.4 × 貼近 EMA（距最近一條 EMA 嘅 ATR 倍數：{LIN['p_near'][0]}→0、{LIN['p_near'][1]:.0f}→1）+ 0.3 × 量縮（近 10 日跌日均量 ÷ 升日均量：{LIN['p_vol'][0]}→0、{LIN['p_vol'][1]}→1）+ 0.3 × 收市在當日高低區間位置"),
    ("底部結構", f"0.5 × 連續上升底數目（{LIN['h_n'][0]:.0f}→0、{LIN['h_n'][1]:.0f}→1）+ 0.3 × 最後一底較上一底升幅（{LIN['h_rise'][0]:.0f}%→0、{LIN['h_rise'][1]:.0f}%→1）+ 0.2 × 最後一底距今（{LIN['h_age'][0]:.0f} 日→0、{LIN['h_age'][1]:.0f} 日→1）"),
    ("三角形分數", f"0.35 × 平頂（|高位斜率| {LIN['r_flat'][0]}→0、{LIN['r_flat'][1]}→1 %/日）+ 0.30 × 低位上升（低位斜率 {LIN['r_low'][0]}→0、{LIN['r_low'][1]}→1 %/日）+ 0.25 × 重心（{LIN['r_cog'][0]}→0、{LIN['r_cog'][1]}→1）+ 0.10 × 頂部觸及（≥2 次→1、1 次→0.5；觸及＝最高價喺 25 日最高 1.5% 內）"),
    ("綜合分數", f"{WT['trend']} × 趨勢 + {WT['pb']} × 回調質素 + {WT['hl']} × 底部結構 + {WT['tri']} × 三角形（每頁排名用；先分梯隊）"),
    ("欄位顏色", "藍字＝量度值（輸入）；黑字＝公式；灰底＝第二梯隊；綠底狀態＝剛黃金交叉；市值組：大型 ≥$100 億、中型 $20–100 億、小型 <$20 億（Nasdaq 快照市值）"),
]
for a, b in scoring:
    put(wr, rw_, 1, a, font=BOLD, align=WRAP); put(wr, rw_, 2, b, align=WRAP); wr.row_dimensions[rw_].height = 32; rw_ += 1

# ================================================================ 市況
wm = wb.create_sheet("市況")
for col, wd in zip("ABCDEFGHIJKLMNOP", (26, 22, 11, 9, 9, 9, 9, 11, 11, 11, 9, 9, 8, 9, 9, 9)):
    wm.column_dimensions[col].width = wd
put(wm, 1, 1, f"市況（{LAST} 收市）", font=TITLE, border=False)
if mkt.get("headline"):
    put(wm, 2, 1, mkt["headline"], font=BOLD, border=False)
rm = 4
MACRO_NAMES = [
    ("SPY", "標普 500 ETF"), ("QQQ", "納指 100 ETF"), ("DIA", "道指 ETF"), ("IWM", "羅素 2000 ETF"), ("MDY", "標普中型 400 ETF"), ("RSP", "標普 500 等權 ETF"),
    ("^SOX", "費城半導體指數"), ("^VIX", "VIX"), ("^TNX", "10 年美債息（%）"), ("^IRX", "3 個月美債息（%）"), ("^TYX", "30 年美債息（%）"), ("DX-Y.NYB", "美元指數"),
    ("CL=F", "WTI 原油"), ("BZ=F", "Brent 原油"), ("GC=F", "黃金"), ("SI=F", "白銀"), ("HG=F", "銅"), ("NG=F", "天然氣"),
    ("XLK", "科技"), ("XLC", "通訊服務"), ("XLY", "非必需消費"), ("XLF", "金融"), ("XLI", "工業"), ("XLV", "醫療保健"), ("XLE", "能源"),
    ("XLB", "原材料"), ("XLP", "必需消費"), ("XLU", "公用事業"), ("XLRE", "房地產"),
]
if MACRO and os.path.exists(MACRO):
    mdf = pd.read_csv(MACRO, dtype={"symbol": str, "date": str})
    mdf = mdf[mdf["date"] <= LAST].sort_values(["symbol", "date"])
    HM = ["代號", "名稱", "收市", "當日%", "5日%", "1個月%（21日）", "3個月%（63日）", "EMA9", "EMA21", "EMA21（5日前）", "距EMA21%", "EMA21 5日斜率%", "9>21", "期內最高距%", "期內最低距%", "最後日期"]
    for j, h in enumerate(HM, 1):
        c = put(wm, rm, j, h, font=HDR_FONT, fill=HDR_FILL, align=Alignment(horizontal="center", vertical="center", wrap_text=True))
    wm.row_dimensions[rm].height = 30
    rm += 1

    def ema_np(x, n):
        a = 2.0 / (n + 1); out = np.empty_like(x); out[0] = x[0]
        for i in range(1, len(x)):
            out[i] = a * x[i] + (1 - a) * out[i - 1]
        return out

    for sym, zh in MACRO_NAMES:
        g = mdf[mdf["symbol"] == sym]
        if len(g) < 30:
            continue
        c = g["close"].to_numpy(float); t = len(c) - 1
        e9, e21 = ema_np(c, 9), ema_np(c, 21)
        put(wm, rm, 1, sym, font=BOLD); put(wm, rm, 2, zh)
        put(wm, rm, 3, round(float(c[t]), 4), '#,##0.00', INPUT)
        put(wm, rm, 4, round(float(c[t] / c[t - 1] - 1), 5), PCT, INPUT)
        put(wm, rm, 5, round(float(c[t] / c[t - 5] - 1), 5), PCT, INPUT)
        put(wm, rm, 6, round(float(c[t] / c[t - 21] - 1), 5), PCT, INPUT)
        put(wm, rm, 7, round(float(c[t] / c[t - 63] - 1), 5) if t >= 63 else None, PCT, INPUT)
        put(wm, rm, 8, round(float(e9[t]), 4), '#,##0.00', INPUT); put(wm, rm, 9, round(float(e21[t]), 4), '#,##0.00', INPUT)
        put(wm, rm, 10, round(float(e21[t - 5]), 4), '#,##0.00', INPUT)
        put(wm, rm, 11, f"=C{rm}/I{rm}-1", PCT); put(wm, rm, 12, f"=I{rm}/J{rm}-1", PCT)
        put(wm, rm, 13, f"=IF(H{rm}>I{rm},\"是\",\"否\")", align=CENTER)
        put(wm, rm, 14, round(float(c[t] / c.max() - 1), 5), PCT, INPUT); put(wm, rm, 15, round(float(c[t] / c.min() - 1), 5), PCT, INPUT)
        put(wm, rm, 16, g["date"].iloc[-1])
        rm += 1
    put(wm, rm, 1, f"Yahoo 日線（data/yahoo/{os.path.basename(MACRO)}），{LAST} 或之前最後一個收市；期內＝{mdf['date'].min()} 起。美債息係收益率（%），其百分比變動只係參考。", font=NOTE, border=False)
    rm += 2

# breadth
put(wm, rm, 1, "大市闊度（合資格股票池）", font=H2, border=False); rm += 1
HB = ["日期", "股票數", "當日中位數%", "上升比例", "企喺 EMA21 上", "EMA9 > EMA21", "EMA21 向上", "20 日新高", "20 日新低", "新高−新低"]
for j, h in enumerate(HB, 1):
    put(wm, rm, j, h, font=HDR_FONT, fill=HDR_FILL, align=Alignment(horizontal="center", vertical="center", wrap_text=True))
wm.row_dimensions[rm].height = 30
rm += 1
b0 = rm
for b in scr["breadth"]:
    put(wm, rm, 1, b["date"]); put(wm, rm, 2, b["n"], INT, INPUT)
    put(wm, rm, 3, round(b["median_ret"], 5) if b["median_ret"] is not None else None, PCT, INPUT)
    put(wm, rm, 4, round(b["pct_up"], 4) if b["pct_up"] is not None else None, '0%', INPUT)
    put(wm, rm, 5, round(b.get("pct_above21", 0), 4), '0%', INPUT); put(wm, rm, 6, round(b.get("pct_9over21", 0), 4), '0%', INPUT)
    put(wm, rm, 7, round(b.get("pct_e21_rising", 0), 4), '0%', INPUT); put(wm, rm, 8, b.get("nh20"), INT, INPUT); put(wm, rm, 9, b.get("nl20"), INT, INPUT)
    put(wm, rm, 10, f"=H{rm}-I{rm}", INT)
    rm += 1
put(wm, rm, 1, "合資格股票池（同篩選規則頁嘅股票池）每日：收市對收市中位數、上升比例、企喺 EMA21 上／EMA9>EMA21／EMA21 向上嘅比例、收市創 20 日新高／新低嘅數目。", font=NOTE, border=False)
rm += 2

# narrative blocks


def block(title, items, link=False):
    global rm
    put(wm, rm, 1, title, font=H2, border=False); rm += 1
    for it in items:
        if link:
            c = put(wm, rm, 1, it[0], font=SRC_FONT, border=False); c.hyperlink = it[1]
        else:
            put(wm, rm, 1, it[0], font=BOLD, align=WRAP)
            put(wm, rm, 2, it[1], align=WRAP)
            wm.merge_cells(start_row=rm, start_column=2, end_row=rm, end_column=12)
            wm.row_dimensions[rm].height = max(30, 15 * (1 + len(it[1]) // 95))
            if len(it) > 2:
                put(wm, rm, 13, it[2], align=WRAP); wm.merge_cells(start_row=rm, start_column=13, end_row=rm, end_column=16)
        rm += 1
    rm += 1


if mkt:
    block("近三日大市", mkt.get("days", []))
    block("宏觀領先指標（息口・債息・油價・美元・波動率）", mkt.get("macro", []))
    block("四大力量周期對標（宏觀流動性／政治／技術革命／黑天鵝）", mkt.get("regime", []))
    put(wm, rm, 1, "對呢份名單嘅含義", font=H2, border=False); rm += 1
    for i, t in enumerate(mkt.get("implications", []), 1):
        put(wm, rm, 1, f"{i}", font=BOLD, align=CENTER); put(wm, rm, 2, t, align=WRAP)
        wm.merge_cells(start_row=rm, start_column=2, end_row=rm, end_column=12); wm.row_dimensions[rm].height = max(30, 15 * (1 + len(t) // 95)); rm += 1
    rm += 1
    put(wm, rm, 1, "領先指標觀察表", font=H2, border=False); rm += 1
    for j, h in enumerate(["指標", "現值", "轉勢門檻"], 1):
        put(wm, rm, j, h, font=HDR_FONT, fill=HDR_FILL, align=CENTER)
    wm.merge_cells(start_row=rm, start_column=3, end_row=rm, end_column=8); rm += 1
    for a, b, c in mkt.get("leading", []):
        put(wm, rm, 1, a, align=WRAP); put(wm, rm, 2, b, align=WRAP); put(wm, rm, 3, c, align=WRAP)
        wm.merge_cells(start_row=rm, start_column=3, end_row=rm, end_column=8); rm += 1
    rm += 1
# sector distribution of tier 1 (formulas)
put(wm, rm, 1, "名單板塊分佈（第一梯隊）", font=H2, border=False); rm += 1
secs = sorted({r["sector_zh"] for r in rows}, key=lambda s: -sum(1 for r in rows if r["sector_zh"] == s))
for s in secs:
    put(wm, rm, 1, s); put(wm, rm, 2, f"=COUNTIFS('總表'!{L['tier']}2:{L['tier']}{last_row},1,'總表'!{L['sec']}2:{L['sec']}{last_row},A{rm})", INT)
    rm += 1
rm += 1
if mkt.get("sources"):
    block("來源", mkt["sources"], link=True)

# ================================================================ 數據核對
wd = wb.create_sheet("數據核對")
wd.column_dimensions["A"].width = 26; wd.column_dimensions["B"].width = 100
for col in "CDEFG":
    wd.column_dimensions[col].width = 14
put(wd, 1, 1, "數據來源同核對", font=TITLE, border=False)
rd_ = 3
src_lines = [
    ("股票池", f"data/yahoo/tickers.txt：{META['symbols']:,} 隻（承接 10MA-watchlist 嘅全美 ≥$1 普通股名單，Nasdaq 快照剔除 ETF／基金／優先股／權證／單位）"),
    ("日線", "；".join(os.path.basename(p) for p in META["sources"]) + f" —— {LAST} 有 bar {META['funnel']['bar_on_last']:,} 隻，其中日內 60 分鐘 bar 滾算（Yahoo 未出日線）{META['intraday_fill']} 隻"),
    ("名稱／板塊／市值", f"Nasdaq 收市後快照 {os.path.basename(META['snapshot']) if META['snapshot'] else '—'}（名稱、板塊、行業、市值）；交易所來自 rreichel3/us-stock-symbols"),
    ("宏觀", f"data/yahoo/{os.path.basename(MACRO)}（指數、ETF、商品、債息、美元）" if MACRO else "—"),
    ("抓取方式", "GitHub Actions runner 跑 scripts/fetch_yahoo.py（yfinance，auto_adjust=False）後 commit；研究容器本身連唔到 Yahoo"),
]
for a, b in src_lines:
    put(wd, rd_, 1, a, font=BOLD, align=WRAP); put(wd, rd_, 2, b, align=WRAP); wd.row_dimensions[rd_].height = 30; rd_ += 1
rd_ += 1
if SNAP and os.path.exists(SNAP):
    snap = {}
    with open(SNAP, encoding="utf-8") as f:
        for q in csv.DictReader(f):
            try:
                snap[q["symbol"].strip()] = float(str(q["lastsale"]).strip("$ ").replace(",", ""))
            except ValueError:
                pass
    # which session does the snapshot hold?  compare against the list's close and previous close
    def match(key):
        hits = [(r["sym"], abs(snap[r["sym"]] / r[key] - 1)) for r in ALL if r["sym"] in snap]
        return hits, (sum(1 for _, d in hits if d <= 0.005) / len(hits) if hits else 0)
    h_close, m_close = match("close"); h_prev, m_prev = match("prev")
    which = ("close", LAST, m_close) if m_close >= m_prev else ("prev", "前一交易日", m_prev)
    put(wd, rd_, 1, "Nasdaq 快照 vs Yahoo", font=H2, border=False); rd_ += 1
    put(wd, rd_, 1, os.path.basename(SNAP), font=BOLD)
    put(wd, rd_, 2, (f"快照「最後成交價」同名單 {len(ALL)} 隻嘅 Yahoo {LAST} 收市吻合（0.5% 內）{m_close:.1%}，同前一交易日收市吻合 {m_prev:.1%} "
                     f"→ 快照入面係{'當日' if which[0]=='close' else '前一日（Nasdaq API 到美東早上五點都仲未轉日，自動重新標籤）'}收市，當{which[1]}收市用"), align=WRAP)
    wd.row_dimensions[rd_].height = 30; rd_ += 1
    hits = h_close if which[0] == "close" else h_prev
    worst = sorted(hits, key=lambda x: -x[1])[:5]
    put(wd, rd_, 1, "名單逐隻", font=BOLD)
    put(wd, rd_, 2, f"{len(hits)} 隻有快照；最大差 {max(d for _, d in hits):.3%}（{'、'.join(f'{s} {d:.2%}' for s, d in worst)}）；0.5% 內 {sum(1 for _, d in hits if d <= 0.005)} 隻", align=WRAP)
    rd_ += 2
    for j, h in enumerate(["代號", "總表排名", f"Yahoo {'收市' if which[0]=='close' else '前收'}", "快照最後成交", "差%"], 1):
        put(wd, rd_, j, h, font=HDR_FONT, fill=HDR_FILL, align=CENTER)
    rd_ += 1
    key = "close" if which[0] == "close" else "prev"
    for r in ALL:
        if r["sym"] not in snap:
            continue
        put(wd, rd_, 1, r["sym"]); put(wd, rd_, 2, f"='總表'!{L['rank']}{ROW_OF[r['sym']]}")
        put(wd, rd_, 3, f"='總表'!{L[key]}{ROW_OF[r['sym']]}", PX); put(wd, rd_, 4, snap[r["sym"]], PX, INPUT); put(wd, rd_, 5, f"=D{rd_}/C{rd_}-1", '0.000%')
        rd_ += 1

# ================================================================ 本版更新
wu = wb.create_sheet("本版更新", 0)
wu.column_dimensions["A"].width = 6; wu.column_dimensions["B"].width = 150
put(wu, 1, 1, f"{REV} 更新重點", font=TITLE, border=False)
fresh1 = sum(1 for r in rows if r["status"] == "剛黃金交叉")
tri1 = sum(1 for r in rows if r["pattern"] == "上升三角（平頂）")
cap_n = {k: sum(1 for r in rows if r["cap"] == k) for k in "abc"}
top5 = "、".join(f"{r['sym']}（{r['score']:.0f}）" for r in rows[:5])
na, nb, nc = (len(SHEETS[k]) for k in "abc")
upd = [
    f"{REV}：加入 MACD（{P['macd_fast']}/{P['macd_slow']}/{P['macd_sig']}）兩個核心條件 —— ⑤ 慢線（訊號線）高過前一日 ⑥ 柱狀圖動能轉勢：近 {P['hist_days']} 日內出現第一根「深紅之後嘅淺紅」或「淺綠之後嘅深綠」，其後每根都升（顏色跟 TradingView 四色柱）。連同 ① 21EMA 向上 ② 9EMA 在 21EMA 上 ③ 一底高於一底 ④ 上升三角重心偏上，六項全過＝第一梯隊。總表加咗 MACD、慢線、慢線斜率、三日柱值、柱顏色、轉勢種類／幾日前、柱顏色序列。",
    f"三個分頁（引用總表）：a 剛黃金交叉（{P['cross_days']} 日內金叉）{na} 隻；b 回到 21EMA（離 EMA21 ≤{P['near_pct']*100:.0f}% 兼 ≤{P['near_atr']} ATR 或近 {P['touch_days']} 日觸及）{nb} 隻；c 回到 9EMA（同樣口徑對 EMA9，離兩條都近就歸較近嗰條）{nc} 隻 —— 都係第一梯隊數目；每頁下方另加灰底第二梯隊（只差一項）同橙底第三梯隊（價格四項過、只差 MACD，共 {META.get('tier3', 0)} 隻）畀你候補。b 頁第一梯隊得零隻：回到 21EMA 嘅股票柱狀圖多數仲喺減速中、未出現第一根淺紅／深綠，候補嗰啲就係要等嘅名。",
    f"數據更新至 {LAST} 收市（{META['sessions']} 個交易日，{META['first_date']} 起；Yahoo 日線由 GitHub Actions 抓取）。",
    f"合資格 {META['eligible']:,} 隻 → C1 {META['cumulative']['C1']} → C2 {META['cumulative']['C2']} → C3 {META['cumulative']['C3']} → C4 {META['cumulative']['C5']} → C5 MACD 慢線 {META['cumulative']['C6']} → C6 柱轉勢 第一梯隊 {META['tier1']} 隻（大型 {cap_n['a']}／中型 {cap_n['b']}／小型 {cap_n['c']}；平頂上升三角 {tri1} 隻）。第二梯隊（差一項）{META['tier2']} 隻。",
    f"排名：綜合分數 = {WT['trend']} 趨勢 + {WT['pb']} 回調質素 + {WT['hl']} 底部結構 + {WT['tri']} 三角形（每項 0–100，總表全部係公式）。頭五位：{top5}。",
    f"市況：{mkt.get('headline', '')}（市況頁有指數／ETF／商品／債息嘅 EMA 狀態、合資格股票池闊度、四大力量周期對標同領先指標觀察表）。",
    "頁：總表（兩個梯隊）、a 剛黃金交叉、b 回到21EMA、c 回到9EMA、大型股／中型股／小型股（引用總表）、差一項、樞軸底頂、篩選規則、市況、數據核對。所有代號連 TradingView；藍字係量度值、黑字係公式。",
    f"產出：{os.path.basename(OUT)}（{MODEL_TAG}；{BUILT_AT}）。",
]
for i, t in enumerate(upd, 1):
    put(wu, i + 1, 1, i, font=BOLD, align=CENTER); put(wu, i + 1, 2, t, align=WRAP)
    wu.row_dimensions[i + 1].height = max(30, 15 * (1 + len(t) // 110))

wb.save(OUT)
print("wrote", OUT, "sheets:", wb.sheetnames)
