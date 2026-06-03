# united-etf-00988a-portfolio

自動下載統一投信 00988A ETF 投資組合資料，並產出標準化持股與每日差異檔。

資料來源：

- https://www.ezmoney.com.tw/ETF/Fund/Info?fundCode=61YTW

## 輸出檔案

```text
data/raw/00988A_portfolio_YYYYMMDD.xlsx
data/out/00988A_holdings_YYYYMMDD.csv
data/out/00988A_diff_YYYYMMDD.csv
data/out/00988A_diff_YYYYMMDD.md
data/out/00988A_latest.csv
```

## 標準化欄位

```text
code,name,shares
```

## 自動執行

GitHub Actions：

```text
.github/workflows/download_00988a.yml
```

預設週一至週五台灣時間 16:30 執行。
