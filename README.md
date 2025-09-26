# Google Apps Script 版：Google 表單資料自動生成簡報

這個專案僅保留 Google Apps Script（GAS）實作：從 Google 表單（回應在 Google 試算表）讀取資料，進行基本分析，生成圖表，並自動建立 Google 簡報；同時可匯出 CSV 與分析報告到 Google 雲端硬碟。

## 快速開始（僅需 Google 帳號，無需 API 金鑰與 GCP 專案）

1) 開啟 Google Apps Script 並建立專案
- 前往 `https://script.google.com/`
- 建立「空白專案」，命名為「Google 表單到簡報生成器」

2) 複製本專案檔案
- 將 `google_apps_script/Code.gs` 內容貼到 Apps Script 的 `Code.gs`
- 將 `google_apps_script/appsscript.json` 內容貼到專案設定（左側檔案清單中新增 `appsscript.json`）

3) 連結表單回應的試算表
- 在 Google 表單的「回應」頁面，點綠色試算表圖示建立/開啟回應試算表
- 在該試算表中，點「擴充功能 > Apps Script」可直接進入同一專案（或於 GAS 中使用 SpreadsheetApp.getActive）

4) 授權並執行
- 在 Apps Script 編輯器，選擇函數 `main`，按「執行」
- 首次執行依指示授權（選擇帳號 > 進階 > 前往專案 > 允許）

完成後，系統會：
- 在試算表內插入圖表（依 CONFIG 設定）
- 建立 Google 簡報（含標題、摘要、洞察、圖表頁）
- 將 CSV 與分析報告輸出到雲端硬碟

## 主要檔案

- `google_apps_script/Code.gs`：主要程式（讀取資料、分析、圖表、建立簡報、輸出）
- `google_apps_script/appsscript.json`：GAS 專案設定
- `google_apps_script_setup_guide.md`：完整設定指南
- `google_apps_script_examples.md`：常見情境範例
- `README_GoogleAppsScript.md`：GAS 版本的完整說明

## 基本設定（在 Code.gs 中）

```javascript
const CONFIG = {
  presentation: {
    title: 'Google表單資料分析報告',
    subtitle: '自動生成簡報'
  },
  charts: [
    { type: 'bar', title: '類別分布分析', xColumn: 0, yColumn: 1 }
  ],
  output: { createSlides: true, createCharts: true, exportToDrive: true }
};
```
- `xColumn`/`yColumn` 為欄位索引（從 0 開始；第一欄=0）
- 圖表類型支援：`bar`、`pie`、`line`

## 常用指令（在 Apps Script 中執行）
- `test()`：快速檢查讀取/分析/圖表是否正常
- `main()`：完整流程（讀取 → 分析 → 圖表 → 建立簡報 → 輸出）

## 自動化（可選）
- 在 Apps Script「觸發器」新增時間驅動觸發，定期執行 `main`

## 疑難排解
- 權限問題：重新執行並完整授權
- 圖表未生成：確認 `xColumn`/`yColumn` 指向的欄位為有效數值/類別
- 簡報未建立：確保有允許建立/修改雲端硬碟檔案的權限

如需更詳細步驟與範例，請參考：
- `google_apps_script_setup_guide.md`
- `google_apps_script_examples.md`
- `README_GoogleAppsScript.md`