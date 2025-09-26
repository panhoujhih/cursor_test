# Google Apps Script 版本 - Google表單資料自動生成簡報

## 🎯 專案概述

這是一個基於Google Apps Script的解決方案，可以從Google表單的資料庫中讀取資料，進行分析處理，並自動生成圖表和簡報。**不需要Google Cloud專案或API金鑰**，完全在Google環境中執行。

## ✨ 主要特色

- 🚀 **零設定**: 不需要Google Cloud專案或API金鑰
- 📊 **自動分析**: 智能分析資料類型並生成洞察
- 📈 **多種圖表**: 支援長條圖、圓餅圖、折線圖
- 🎨 **自動簡報**: 生成專業的Google Slides簡報
- 💾 **多格式輸出**: CSV匯出、分析報告、圖表檔案
- ⏰ **自動化執行**: 支援定期自動執行
- 🔄 **即時更新**: 資料變更時自動重新生成

## 📁 檔案結構

```
google_apps_script/
├── Code.gs                    # 主要程式碼
├── appsscript.json           # 專案配置
├── google_apps_script_setup_guide.md  # 設定指南
├── google_apps_script_examples.md     # 使用範例
└── README_GoogleAppsScript.md         # 本檔案
```

## 🚀 快速開始

### 第一步：建立Google Apps Script專案

1. 前往 [Google Apps Script](https://script.google.com/)
2. 點擊「新增專案」
3. 輸入專案名稱：「Google表單到簡報生成器」

### 第二步：複製程式碼

1. 將 `Code.gs` 的內容複製到您的專案中
2. 將 `appsscript.json` 的內容複製到專案配置中
3. 儲存專案

### 第三步：準備資料

1. 確保您的Google表單已連結到試算表
2. 試算表包含標題行和資料行
3. 至少有一個數值欄位用於圖表

### 第四步：執行

1. 在Apps Script編輯器中點擊「執行」
2. 選擇 `main` 函數
3. 授權權限
4. 等待執行完成

## 📊 功能說明

### 資料處理
- 自動讀取Google試算表資料
- 智能識別資料類型（數值、文字、日期）
- 自動清理和驗證資料
- 生成資料洞察和統計

### 圖表生成
- **長條圖**: 適合比較不同類別的數值
- **圓餅圖**: 適合顯示比例和占比
- **折線圖**: 適合顯示趨勢和變化

### 簡報建立
- 自動建立Google Slides簡報
- 包含標題頁、資料摘要、洞察分析
- 整合生成的圖表
- 專業的版面設計

### 結果匯出
- CSV格式的原始資料
- 詳細的分析報告
- 圖表檔案
- 簡報連結

## ⚙️ 配置說明

### 基本配置
```javascript
const CONFIG = {
  presentation: {
    title: '您的簡報標題',
    subtitle: '您的副標題'
  },
  charts: [
    {
      type: 'bar',        // 圖表類型
      title: '圖表標題',
      xColumn: 0,         // X軸欄位索引
      yColumn: 1          // Y軸欄位索引
    }
  ],
  output: {
    createSlides: true,   // 是否建立簡報
    createCharts: true,   // 是否建立圖表
    exportToDrive: true   // 是否匯出檔案
  }
};
```

### 圖表類型
- `bar`: 長條圖
- `pie`: 圓餅圖
- `line`: 折線圖

### 欄位索引
- 第一欄（標題）= 索引 0
- 第二欄 = 索引 1
- 第三欄 = 索引 2
- 以此類推...

## 📝 使用範例

### 範例1：員工資料分析
```javascript
const CONFIG = {
  presentation: {
    title: '員工資料分析報告',
    subtitle: '2024年第一季'
  },
  charts: [
    {
      type: 'bar',
      title: '各部門人數統計',
      xColumn: 2, // 部門欄位
      yColumn: 3  // 人數欄位
    }
  ]
};
```

### 範例2：銷售資料分析
```javascript
const CONFIG = {
  presentation: {
    title: '銷售資料分析',
    subtitle: '多維度分析報告'
  },
  charts: [
    {
      type: 'bar',
      title: '各產品銷售額',
      xColumn: 0, // 產品名稱
      yColumn: 1  // 銷售額
    },
    {
      type: 'pie',
      title: '銷售區域占比',
      xColumn: 2, // 銷售區域
      yColumn: 1  // 銷售額
    }
  ]
};
```

## 🔧 進階功能

### 自動化執行
```javascript
// 設定每日自動執行
function setupDailyTrigger() {
  ScriptApp.newTrigger('main')
    .timeBased()
    .everyDays(1)
    .atHour(9)
    .create();
}
```

### 多工作表處理
```javascript
// 處理多個工作表
function processMultipleSheets() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = spreadsheet.getSheets();
  // 處理邏輯...
}
```

### 外部通知
```javascript
// 發送結果到Slack
function sendToSlack(message, webhookUrl) {
  const payload = { text: message };
  UrlFetchApp.fetch(webhookUrl, {
    method: 'POST',
    contentType: 'application/json',
    payload: JSON.stringify(payload)
  });
}
```

## 🛠️ 故障排除

### 常見問題

**Q: 出現「權限被拒絕」錯誤**
A: 重新執行授權流程，確保所有權限都被授予

**Q: 圖表沒有生成**
A: 檢查：
- 欄位索引是否正確
- 資料是否包含數值
- 圖表配置是否正確

**Q: 簡報沒有建立**
A: 檢查：
- 是否啟用了Slides API
- 是否有足夠的權限建立檔案

**Q: 資料讀取失敗**
A: 檢查：
- 試算表是否有資料
- 第一行是否為標題行
- 資料格式是否正確

### 除錯技巧

1. **使用測試函數**: 執行 `test()` 函數檢查基本功能
2. **查看執行日誌**: 在Apps Script編輯器中查看詳細日誌
3. **逐步執行**: 分別執行各個函數確認問題所在
4. **檢查資料格式**: 確認試算表資料格式正確

## 📈 效能優化

### 資料量限制
- 建議每次處理不超過10,000筆記錄
- 圖表數量建議不超過10個
- 簡報投影片數量建議不超過20張

### 執行時間優化
- 使用 `test()` 函數先測試小量資料
- 分批處理大量資料
- 避免在單次執行中處理過多圖表

## 🔒 安全性

### 權限管理
- 只授予必要的權限
- 定期檢查和更新權限設定
- 使用服務帳戶進行自動化執行

### 資料保護
- 所有資料都在Google環境中處理
- 不會將資料傳送到外部服務
- 可以設定資料保留期限

## 📚 學習資源

### 官方文件
- [Google Apps Script 文件](https://developers.google.com/apps-script)
- [Google Sheets API](https://developers.google.com/sheets/api)
- [Google Slides API](https://developers.google.com/slides/api)

### 相關資源
- [Apps Script 範例程式碼](https://developers.google.com/apps-script/samples)
- [Google Workspace 開發者指南](https://developers.google.com/workspace)

## 🤝 貢獻

歡迎提交Issue和Pull Request來改善這個專案：

1. Fork 這個專案
2. 建立您的功能分支
3. 提交您的變更
4. 推送到分支
5. 建立Pull Request

## 📄 授權

本專案採用MIT授權條款。

## 🆘 支援

如果您遇到任何問題：

1. 查看本README的故障排除部分
2. 參考 `google_apps_script_setup_guide.md` 設定指南
3. 查看 `google_apps_script_examples.md` 使用範例
4. 建立GitHub Issue描述您的問題

---

**開始使用**: 按照 `google_apps_script_setup_guide.md` 的步驟完成設定，然後執行 `main()` 函數即可開始使用！

**優勢**: 這個Google Apps Script版本比Python版本更簡單，不需要任何外部設定，完全在Google環境中執行，非常適合快速開始使用。
