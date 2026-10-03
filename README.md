# 9EMA-Uptrend Watchlist

全美上市普通股掃描，篩選「**21EMA 向上 ＋ 9EMA 在 21EMA 上方 ＋ 一底高於一底 ＋ 形態接近上升三角（重心偏上）＋ MACD 慢線向上 ＋ MACD 柱轉勢**」嘅股票（R3.00 起；R2.00 冇 MACD 兩項），再分三頁：**a 剛黃金交叉／b 回到 21EMA／c 回到 9EMA**（R2.00 起；R1.00 將「回到 9/21EMA」當核心門檻），
以 **趨勢 × 回調質素 × 底部結構 × 三角形** 綜合分數排序。股票池同數據管線承接
[10MA-watchlist](https://github.com/yppmatthewtw-cmd/10MA-watchlist)（Yahoo 日線由 GitHub Actions runner 抓取後 commit，研究容器本身連唔到 Yahoo）。

## 報告（`reports/`）

| 版本 | 內容 |
|------|------|
| R4.00 | 條件同 R3.00，數據更新至 **2026-10-02 收盤**（193 個交易日；10-02 Yahoo 日線齊 3,850 隻，毋須日內滾算）。交付 Excel（`9EMA-Uptrend_Watchlist_R4.00_claudefable51xhigh_10.03_1043.xlsx`，13 頁）。合資格 2,792 → 21EMA 向上 630 → 9EMA 在上 597 → 一底高於一底 368 → 上升三角 97 → MACD 慢線 44 → 柱轉勢 **第一梯隊 26 隻**（AMN、SEI、DDD、FFIV、ANET、FTNT、PANW、NEU、HNGE、QLYS、CNXN、CRVL…）。分頁：**a 剛黃金交叉 6 隻**（NEU、TENB、NHC、STRL、OPY、BJ）、**b 回到 21EMA 0 隻**（候補 8 隻：HYPD、SMTI、GYRE、ULTA、WAT、HALO、BTGO、HCA）、**c 回到 9EMA 18 隻**。第二梯隊 35 隻、第三梯隊（只差 MACD）71 隻。**市況**：10-02 九月非農只加 2.9 萬、失業率 4.2% → 10 月加息預期大減，10 年息回落至 5.28%、VIX 15.3、WTI $91.3；SPY +0.7%、QQQ 創新高；合資格股票 63% 上升、32% 企喺 21EMA 上（10-01 係 26%），新高 273 對新低 383 —— 由「偏防守」升級為「審慎偏多」，倉位由一半加至七成。Nasdaq 快照仍係 09-30 收市，只用於名稱／板塊／市值。8,040 項公式 pycel 重算 0 錯誤 |
| R3.00 | **加入 MACD（12/26/9）兩個核心條件**：⑤ 慢線（訊號線）高過前一日 ⑥ 柱狀圖動能轉勢 —— 近 3 日內出現第一根「深紅之後嘅淺紅」（跌勢減速）或「淺綠之後嘅深綠」（升勢再加速），其後每根都升（顏色跟 TradingView 四色柱）。數據同 R1／R2（**2026-10-01 收盤**）。交付 Excel（`9EMA-Uptrend_Watchlist_R3.00_claudefable51xhigh_10.02_1708.xlsx`，13 頁）。合資格 2,791 → 573 → 542 → 322 → 上升三角 100 → MACD 慢線 61 → 柱轉勢 **第一梯隊 25 隻**（HNGE、TTMI、MTSI、DT、STX、TWIN、HPE、NVDA、AMN、CNXN…；23 隻係淺綠→深綠、2 隻係深紅→淺紅）。分頁：**a 剛黃金交叉 6 隻**（NEU、CAT、PWR、CLDT、AMG、FNRN）、**b 回到 21EMA 0 隻**、**c 回到 9EMA 17 隻**。第二梯隊（只差一項）30 隻；新增**第三梯隊 74 隻**（價格四項全過、只差 MACD，總表橙底，三個分頁都列埋做候補 —— b 頁候補 10 隻：FRNM、ADPT、FLGT、FTRE、BTGO、WAT、SMTI、BPRN、PKOH、BIIB）。總表加咗 MACD、慢線、慢線斜率、三日柱值、柱顏色、轉勢種類／幾日前、柱顏色序列。8,375 項公式 pycel 重算 0 錯誤 |
| R2.00 | **條件按用戶重訂**：核心四項（21EMA 向上 ＋ 9EMA 在 21EMA 上 ＋ 一底高於一底 ＋ 上升三角重心偏上）做總表，「回到 9/21EMA」改為分頁條件；數據同 R1（**2026-10-01 收盤**）。交付 Excel（`9EMA-Uptrend_Watchlist_R2.00_claudefable51xhigh_10.02_1658.xlsx`，13 頁）。合資格 2,791 → 573 → 542 → 322 → **第一梯隊 100 隻**（大型 26／中型 30／小型 43；平頂上升三角 30）；三個分頁都係第一梯隊子集並引用總表：**a 剛黃金交叉 15 隻**（UCTT、MAZE、NEU、HWKN、GFS、CAT、PWR、M、CLDT、AMBA、AIP、AMG、FNRN、COLM、OPY）、**b 回到 21EMA 10 隻**（FRNM、ADPT、FLGT、FTRE、BTGO、WAT、SMTI、BPRN、PKOH、BIIB —— 離兩條 EMA 都近嘅股票歸較近嗰條，所以 b 係較深回調）、**c 回到 9EMA 76 隻**；14 隻離兩條 EMA 都超過 3%（只喺總表）。第二梯隊（只差上升三角一項）99 隻。頭五位：JILL、PPIH、MG、HNGE、TTMI。市況同 R1。9,908 項公式 pycel 重算 0 錯誤 |
| R1.00 | **首版**，數據更新至 **2026-10-01 收盤**（192 個交易日，2025-12-26 起），交付 Excel（`9EMA-Uptrend_Watchlist_R1.00_claudefable51xhigh_10.02_1641.xlsx`，11 頁）。合資格 2,791 → 21EMA 向上 573 → 9EMA 在上 542 → 一底高於一底 322 → 回到 9/21EMA 223 → 上升三角 **第一梯隊 86 隻**（大型 20／中型 25／小型 40；剛黃金交叉 15、平頂上升三角 27）；每個門檻放寬一級嘅**第二梯隊 80 隻**（差一項頁列明每隻差邊項同量度值）；**剛黃金交叉池 95 隻**（C1＋5 日內金叉＋C4，唔理底／形態，提前睇）。頭五位：JILL、PPIH、MG、TTMI、NVAX。**市況**：標普 7,666 離高位唔遠，但合資格股票只有 26% 企喺 21EMA 上、26% 嘅 21EMA 向上（09-30 仲係 23%），09-30 新低 942 對新高 106、10-01 569 對 202；10 年期美債息 5.24–5.33%（2002 年以來最高）、WTI $90 樓上、聯儲 09-16 加息後市場再價入 10 月加息（約 70%）—— 偏防守，只做高 RS 嘅 9/21EMA 回調、倉位減半（市況頁有四大力量周期對標同領先指標觀察表）。所有代號連 TradingView；可推導欄位全部係公式；6,107 項公式 pycel 重算 0 錯誤，五個分數對篩選程式最大差 0.005 |

## 篩選規則（R3.00）

以最後一個完整交易日 t 嘅 Yahoo 日線量度（EMA 用標準指數平滑 α=2/(N+1)，ATR14 用 Wilder 平滑）。`scripts/scan_9ema.py`：

1. **股票池**：當日有成交、≥90 個交易日歷史、收市 ≥$2、20 日成交額中位數 ≥$1M（3,868 → 2,791）。
2. **C1 21EMA 向上**：EMA21[t] > EMA21[t−1] 而且 > EMA21[t−5]。
3. **C2 9EMA 在 21EMA 上**：EMA9[t] > EMA21[t]；最近一次黃金交叉喺 5 日內＝「剛黃金交叉」，否則「持續在上方」。
4. **C3 一底高於一底**：用最低價搵樞軸底（前後各 3 日最低）；近 60 日最後兩個底遞升、最後一底 ≤30 日前、之後收市冇跌穿佢。
5. **C4 上升三角／重心偏上**：近 25 日 VWAP 喺高低區間嘅位置（重心）≥0.5、樞軸低位斜率 >0、樞軸高位斜率喺 −0.10 至 +1.0 %/日之間（平頂至微升）。
5b. **C5 MACD 慢線向上**：MACD（EMA12 − EMA26）嘅訊號線（MACD 嘅 EMA9）高過前一日。
5c. **C6 MACD 柱轉勢**：柱 = MACD − 慢線；柱 ≥0 升＝深綠、跌＝淺綠，柱 <0 跌＝深紅、升＝淺紅。近 3 日內要出現第一根「深紅之後嘅淺紅」或「淺綠之後嘅深綠」，而且由嗰根起每根都升。
6. **分頁**（第一梯隊子集，另列灰底第二梯隊同橙底第三梯隊做候補）：a 剛黃金交叉＝金叉 ≤5 日；b 回到 21EMA＝收市離 EMA21 ≤3% 兼（≤0.75 ATR14 或近 3 日最低價觸及 EMA21×1.005）、收市唔低過 EMA21 3%；c 回到 9EMA＝同樣口徑對 EMA9；離兩條都近就歸較近嗰條（b 深、c 淺）。
7. **評分**（每項 0–100，線性兩頭封頂，總表全部係公式）：趨勢（EMA21 斜率、上升步比例、RS21）、回調質素（貼近 EMA 嘅 ATR 倍數、跌日／升日量比、收市在當日區間位置）、底部結構（連續上升底數目、底較底升幅、最後一底距今）、三角形（平頂、低位上升、重心、頂部觸及）；綜合 = 0.25 趨勢 + 0.25 回調質素 + 0.20 底部結構 + 0.30 三角形。
8. **第二梯隊**：只唔過一項、而且嗰項喺放寬門檻下通過（底回望 80 日／最後一底 ≤40 日；重心 ≥0.45、高位斜率 −0.20 至 +1.5；柱轉勢 ≤5 日）。**第三梯隊**：價格四項全過、只差 MACD 慢線／柱轉勢（候補）。

## 數據

- `data/yahoo/tickers.txt`：3,868 隻全美 ≥$1 普通股（承接 10MA-watchlist，Nasdaq 快照剔除 ETF／基金／優先股／權證／單位）。
- `data/yahoo/eod_*.csv.gz`：Yahoo 日線（open/high/low/close/adj_close/volume）；`intraday_*_60m.csv.gz`：60 分鐘 bar 滾算成日線，補 Yahoo 未出日線嘅最後一日；`macro_daily_*.csv.gz`：指數、ETF、商品、債息、美元（`data/yahoo/macro_tickers.txt`）。
- 名稱／板塊／行業／市值：10MA-watchlist 嘅 Nasdaq 收市後快照（`data/snapshots/`）；交易所：rreichel3/us-stock-symbols。
- `data/screen_r<版本>.json`：篩選結果（三個梯隊、a/b/c 分頁、闊度、參數）；`data/market_r<版本>.json`：市況頁嘅文字（近三日大市、宏觀領先指標、四大力量周期、含義、觀察表、來源）。

## 重跑

```bash
# 1. 抓數據（GitHub Actions → Actions → Fetch Yahoo daily bars → Run workflow）
#    start=2025-12-26 end=<最後交易日+1> interval=1d                      → data/yahoo/eod_<start>_<end>.csv.gz
#    start=<最後交易日> end=<最後交易日+1> interval=60m                   → intraday 滾算（Yahoo 未出日線時用）
#    list=data/yahoo/macro_tickers.txt out=data/yahoo/macro_daily_<start>_<end>.csv.gz
# 2. 篩選
YAHOO=data/yahoo/eod_2025-12-26_2026-10-04.csv.gz SUPP=data/yahoo/intraday_2026-10-01_2026-10-03_60m.csv.gz \
SNAP=<Nasdaq 快照 csv> OUT_JSON=data/screen_r4.json python3 scripts/scan_9ema.py
# 3. Excel（檔名：9EMA-Uptrend_Watchlist_R<版本>_<model><effort>_<mm.dd>_<hhmm 香港時間>.xlsx）
WORK_DIR=./data SCREEN_JSON=screen_r4.json MARKET_JSON=market_r4.json MACRO=data/yahoo/macro_daily_2025-12-26_2026-10-04.csv.gz \
SNAP=<Nasdaq 快照 csv> REV=R4.00 MODEL_TAG="claudefable51xhigh" BUILT_AT="<時間>" OUT_XLSX=reports/<檔名>.xlsx python3 scripts/build_9ema_xlsx.py
# 4. 核對：pycel 重算全部公式、對比篩選程式嘅分數
python3 scripts/verify_xlsx.py reports/<檔名>.xlsx data/screen_r4.json
```
