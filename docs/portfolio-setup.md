# 台股投資紀錄（方案1）設定說明

本方案：**Google Sheet 記錄交易 + Apps Script 每日收盤後更新 `data/portfolio.json` + GitHub Pages 顯示**。

## 1) 建立 Google Sheet
建立一個工作表，命名為：`Trades`

第一列欄位請固定如下：

| Date | Symbol | Market | Side | Shares | Price | Fee | Tax | Note |
|---|---|---|---|---:|---:|---:|---:|---|
| 2026-03-07 | 2330 | TW | BUY | 1000 | 978 | 1426 | 0 | 首次建倉 |

欄位說明：
- `Date`: 交易日期（YYYY-MM-DD）
- `Symbol`: 股票代號（台股四碼）
- `Market`: 固定 `TW`
- `Side`: `BUY` 或 `SELL`
- `Shares`: 股數
- `Price`: 成交價
- `Fee`: 手續費
- `Tax`: 交易稅（賣出有）
- `Note`: 備註

## 2) 安裝 Apps Script
1. 在 Google Sheet 點選「擴充功能」→「Apps Script」
2. 貼上 `scripts/update_portfolio.gs` 內容
3. 在「專案設定」→「指令碼屬性」設定：
   - `GITHUB_TOKEN`：GitHub PAT（需 repo 權限）
   - `GITHUB_OWNER`：`shufree727jp-a11y`
   - `GITHUB_REPO`：`shufree727jp-a11y.github.io`
   - `GITHUB_BRANCH`：`main`

## 3) 排程（每日收盤後）
在 Apps Script 觸發器新增：
- 執行函式：`dailyUpdateTWPortfolio`
- 時間型觸發：每天
- 建議時間：台灣時間 14:10～15:00

## 4) 資料更新流程
`dailyUpdateTWPortfolio()` 會：
1. 讀取 `Trades`
2. 抓取台股最新價
3. 計算持倉與損益
4. 直接更新 GitHub 上 `data/portfolio.json`
5. GitHub Pages 自動反映到 `portfolio.html`

## 5) 注意事項
- 若當天尚未有成交價，會使用 0 或前值（可依需求再強化）。
- 建議先用少量測試資料跑一次。
- 請勿把 Token 寫死在程式碼，請放在 Script Properties。

## 6) 欄位名稱可用中文（已支援）
目前腳本同時支援英文與中文欄位名稱，對照如下：

- Date / 日期
- Symbol / 股票代號 / 代號
- Market / 市場
- Side / 買賣別 / 買賣（可填 BUY/SELL 或 買/賣）
- Shares / 股數
- Price / 成交價 / 價格
- Fee / 手續費
- Tax / 交易稅 / 稅
- Note / 備註
