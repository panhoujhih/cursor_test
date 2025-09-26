# Google Apps Script 使用範例

## 範例1：基本使用

### 設定
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

### 執行
1. 在Apps Script編輯器中執行 `main()` 函數
2. 系統會自動：
   - 讀取試算表資料
   - 生成長條圖
   - 建立簡報
   - 匯出CSV和分析報告

## 範例2：多圖表分析

### 設定
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
    },
    {
      type: 'line',
      title: '月度銷售趨勢',
      xColumn: 3, // 月份
      yColumn: 1  // 銷售額
    }
  ]
};
```

## 範例3：問卷調查分析

### 試算表結構
```
姓名 | 年齡 | 性別 | 滿意度 | 建議
張三 | 25  | 男   | 4      | 很好
李四 | 30  | 女   | 5      | 優秀
```

### 設定
```javascript
const CONFIG = {
  presentation: {
    title: '客戶滿意度調查報告',
    subtitle: '2024年客戶回饋分析'
  },
  charts: [
    {
      type: 'pie',
      title: '性別分布',
      xColumn: 2, // 性別
      yColumn: 2  // 性別 (用於計數)
    },
    {
      type: 'bar',
      title: '滿意度分布',
      xColumn: 3, // 滿意度
      yColumn: 3  // 滿意度 (用於計數)
    }
  ]
};
```

## 範例4：活動報名統計

### 試算表結構
```
報名時間 | 姓名 | 聯絡方式 | 活動類型 | 人數
2024-01-01 | 張三 | 0912345678 | 講座 | 1
2024-01-02 | 李四 | 0987654321 | 工作坊 | 2
```

### 設定
```javascript
const CONFIG = {
  presentation: {
    title: '活動報名統計報告',
    subtitle: '2024年活動參與分析'
  },
  charts: [
    {
      type: 'bar',
      title: '各活動類型報名人數',
      xColumn: 3, // 活動類型
      yColumn: 4  // 人數
    },
    {
      type: 'line',
      title: '每日報名趨勢',
      xColumn: 0, // 報名時間
      yColumn: 4  // 人數
    }
  ]
};
```

## 範例5：自定義資料處理

### 進階資料處理函數
```javascript
function processDataAdvanced(data) {
  const headers = data.headers;
  const rows = data.rows;
  
  // 自定義統計
  const customStats = {
    totalRecords: rows.length,
    dateRange: null,
    topCategories: {},
    trends: []
  };
  
  // 分析日期範圍
  const dateColumn = headers.findIndex(h => h.includes('日期') || h.includes('時間'));
  if (dateColumn !== -1) {
    const dates = rows.map(row => new Date(row[dateColumn])).filter(d => !isNaN(d));
    if (dates.length > 0) {
      customStats.dateRange = {
        start: new Date(Math.min(...dates)),
        end: new Date(Math.max(...dates))
      };
    }
  }
  
  // 分析類別分布
  const categoryColumn = headers.findIndex(h => h.includes('類別') || h.includes('類型'));
  if (categoryColumn !== -1) {
    const categories = {};
    rows.forEach(row => {
      const category = row[categoryColumn];
      categories[category] = (categories[category] || 0) + 1;
    });
    customStats.topCategories = categories;
  }
  
  return {
    ...data,
    customStats: customStats
  };
}
```

## 範例6：條件式圖表生成

### 智能圖表選擇
```javascript
function generateSmartCharts(processedData) {
  const charts = [];
  const headers = processedData.headers;
  const stats = processedData.stats;
  
  // 根據資料類型自動選擇圖表
  Object.keys(stats.columnTypes).forEach((header, index) => {
    const columnType = stats.columnTypes[header].type;
    
    if (columnType === 'numeric') {
      // 數值欄位生成長條圖
      charts.push({
        type: 'bar',
        title: `${header} 分布`,
        xColumn: index,
        yColumn: index
      });
    } else if (columnType === 'categorical') {
      // 類別欄位生成圓餅圖
      charts.push({
        type: 'pie',
        title: `${header} 占比`,
        xColumn: index,
        yColumn: index
      });
    }
  });
  
  return charts;
}
```

## 範例7：多工作表處理

### 處理多個工作表
```javascript
function processMultipleSheets() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = spreadsheet.getSheets();
  const results = [];
  
  sheets.forEach((sheet, index) => {
    try {
      // 設定當前工作表
      SpreadsheetApp.setActiveSheet(sheet);
      
      // 處理資料
      const data = getCurrentSheetData();
      const processedData = processData(data);
      
      results.push({
        sheetName: sheet.getName(),
        data: processedData,
        success: true
      });
      
    } catch (error) {
      results.push({
        sheetName: sheet.getName(),
        error: error.toString(),
        success: false
      });
    }
  });
  
  return results;
}
```

## 範例8：錯誤處理和日誌

### 完整的錯誤處理
```javascript
function mainWithErrorHandling() {
  const startTime = new Date();
  const log = [];
  
  try {
    log.push(`開始執行: ${startTime.toLocaleString()}`);
    
    // 步驟1: 資料讀取
    log.push('步驟1: 讀取資料...');
    const data = getCurrentSheetData();
    log.push(`✅ 成功讀取 ${data.rows.length} 筆記錄`);
    
    // 步驟2: 資料處理
    log.push('步驟2: 處理資料...');
    const processedData = processData(data);
    log.push(`✅ 資料處理完成`);
    
    // 步驟3: 生成圖表
    log.push('步驟3: 生成圖表...');
    const charts = generateCharts(processedData);
    log.push(`✅ 生成 ${charts.length} 個圖表`);
    
    // 步驟4: 建立簡報
    log.push('步驟4: 建立簡報...');
    const presentation = createPresentation(processedData);
    log.push(`✅ 簡報建立完成`);
    
    const endTime = new Date();
    const duration = (endTime - startTime) / 1000;
    
    log.push(`執行完成: ${endTime.toLocaleString()}`);
    log.push(`總耗時: ${duration} 秒`);
    
    // 儲存執行日誌
    saveExecutionLog(log);
    
    return {
      status: 'success',
      duration: duration,
      chartsGenerated: charts.length,
      presentationCreated: !!presentation
    };
    
  } catch (error) {
    log.push(`❌ 執行失敗: ${error.toString()}`);
    saveExecutionLog(log);
    
    return {
      status: 'error',
      error: error.toString(),
      log: log
    };
  }
}

function saveExecutionLog(log) {
  const timestamp = new Date().toISOString().replace(/[:.]/g, '-');
  const logContent = log.join('\n');
  const logBlob = Utilities.newBlob(logContent, 'text/plain', `execution_log_${timestamp}.txt`);
  DriveApp.getRootFolder().createFile(logBlob);
}
```

## 範例9：定期自動執行

### 設定每日自動執行
```javascript
function setupDailyTrigger() {
  // 刪除現有的觸發器
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'main') {
      ScriptApp.deleteTrigger(trigger);
    }
  });
  
  // 建立新的每日觸發器
  ScriptApp.newTrigger('main')
    .timeBased()
    .everyDays(1)
    .atHour(9) // 每天早上9點執行
    .create();
    
  console.log('✅ 每日自動執行已設定');
}

function setupWeeklyTrigger() {
  ScriptApp.newTrigger('main')
    .timeBased()
    .everyWeeks(1)
    .onWeekDay(ScriptApp.WeekDay.MONDAY)
    .atHour(10)
    .create();
    
  console.log('✅ 每週自動執行已設定');
}
```

## 範例10：與外部系統整合

### 發送結果到Slack
```javascript
function sendToSlack(message, webhookUrl) {
  const payload = {
    text: message,
    username: 'Google Apps Script Bot',
    icon_emoji: ':chart_with_upwards_trend:'
  };
  
  const options = {
    method: 'POST',
    contentType: 'application/json',
    payload: JSON.stringify(payload)
  };
  
  try {
    UrlFetchApp.fetch(webhookUrl, options);
    console.log('✅ 訊息已發送到Slack');
  } catch (error) {
    console.error('❌ 發送到Slack失敗:', error);
  }
}

function mainWithSlackNotification() {
  const result = main();
  
  if (result.status === 'success') {
    const message = `📊 簡報生成完成！\n` +
                   `- 處理記錄數: ${result.dataRecords}\n` +
                   `- 生成圖表數: ${result.chartsGenerated}\n` +
                   `- 簡報建立: ${result.presentationCreated ? '是' : '否'}`;
    
    sendToSlack(message, 'YOUR_SLACK_WEBHOOK_URL');
  } else {
    sendToSlack(`❌ 簡報生成失敗: ${result.message}`, 'YOUR_SLACK_WEBHOOK_URL');
  }
}
```

## 使用建議

1. **開始簡單**: 先使用基本範例，確認功能正常
2. **逐步複雜**: 根據需求添加更多圖表和功能
3. **測試先行**: 使用 `test()` 函數確認設定正確
4. **監控執行**: 查看執行日誌了解執行狀況
5. **備份重要**: 定期備份重要的配置和程式碼

---

這些範例涵蓋了從基本使用到進階功能的所有場景，您可以根據自己的需求選擇合適的範例進行修改和使用。
