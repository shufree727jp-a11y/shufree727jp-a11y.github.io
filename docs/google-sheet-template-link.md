# Google Sheet 範本連結版（台股投資紀錄）

以下是可直接使用的連結：

## 1) 一鍵開新 Google Sheet
- https://sheet.new

開啟後請將工作表命名為：`Trades`

## 2) 匯入範本 CSV（含欄位與範例）
- https://raw.githubusercontent.com/shufree727jp-a11y/shufree727jp-a11y.github.io/main/docs/trades-template.csv

在 Google Sheet 內操作：
1. 檔案 → 匯入
2. 選擇「連結」貼上上面 URL（或先下載再上傳）
3. 匯入到現有工作表 `Trades`

## 3) 欄位（中文可用）
- 日期
- 股票代號
- 市場
- 買賣別（買/賣 或 BUY/SELL）
- 股數
- 成交價
- 手續費
- 交易稅
- 備註

## 4) Apps Script 程式碼位置
- https://raw.githubusercontent.com/shufree727jp-a11y/shufree727jp-a11y.github.io/main/scripts/update_portfolio.gs

貼到你的 Apps Script 專案後，設定 Script Properties：
- `GITHUB_TOKEN`
- `GITHUB_OWNER=shufree727jp-a11y`
- `GITHUB_REPO=shufree727jp-a11y.github.io`
- `GITHUB_BRANCH=main`
