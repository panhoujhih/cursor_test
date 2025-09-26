/**
 * Google Apps Script - 任務追蹤資料分析
 * 主要功能：從指定試算表讀取master_table和detail_table，篩選追蹤任務並關聯最新工作內容
 */

// 試算表ID
const SPREADSHEET_ID = '1y0AIZo5yceO6OOwG9igv-gTYYOuN3Kz-a0vZZDF5KO8';

// 全域配置
const CONFIG = {
  // 簡報設定
  presentation: {
    title: '任務追蹤分析報告',
    subtitle: '追蹤中任務的最新工作內容'
  },
  
  // 圖表設定
  charts: [
    {
      type: 'bar',
      title: '任務狀態分布',
      xColumn: 0, // 任務名稱
      yColumn: 1  // 工作內容數量
    }
  ],
  
  // 輸出設定
  output: {
    createSlides: true,
    createCharts: true,
    exportToDrive: true
  }
};

/**
 * 主函數 - 執行完整的資料處理和簡報生成流程
 */
function main() {
  try {
    console.log('🚀 開始執行Google表單資料處理流程');
    
    // 步驟1: 取得當前試算表資料
    const data = getCurrentSheetData();
    if (!data || data.length === 0) {
      throw new Error('無法取得有效資料');
    }
    
    // 步驟2: 處理和分析資料
    const processedData = processData(data);
    
    // 步驟3: 生成圖表
    if (CONFIG.output.createCharts) {
      generateCharts(processedData);
    }
    
    // 步驟4: 建立簡報
    if (CONFIG.output.createSlides) {
      createPresentation(processedData);
    }
    
    // 步驟5: 匯出結果
    if (CONFIG.output.exportToDrive) {
      exportResults(processedData);
    }
    
    console.log('✅ 所有步驟完成！');
    return {
      status: 'success',
      message: '簡報生成完成',
      dataRecords: data.length,
      chartsGenerated: CONFIG.charts.length,
      presentationCreated: CONFIG.output.createSlides
    };
    
  } catch (error) {
    console.error('❌ 執行失敗:', error);
    return {
      status: 'error',
      message: error.toString(),
      dataRecords: 0,
      chartsGenerated: 0,
      presentationCreated: false
    };
  }
}

/**
 * 取得指定試算表的資料
 */
function getCurrentSheetData() {
  try {
    // 開啟指定的試算表
    const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
    
    // 取得master_table工作表
    const masterSheet = spreadsheet.getSheetByName('master_table');
    if (!masterSheet) {
      throw new Error('找不到master_table工作表');
    }
    
    // 取得detail_table工作表
    const detailSheet = spreadsheet.getSheetByName('detail_table');
    if (!detailSheet) {
      throw new Error('找不到detail_table工作表');
    }
    
    // 讀取master_table資料
    const masterData = masterSheet.getDataRange().getValues();
    const masterHeaders = masterData[0];
    const masterRows = masterData.slice(1);
    
    // 讀取detail_table資料
    const detailData = detailSheet.getDataRange().getValues();
    const detailHeaders = detailData[0];
    const detailRows = detailData.slice(1);
    
    console.log(`📊 Master Table: ${masterRows.length} 筆記錄, ${masterHeaders.length} 個欄位`);
    console.log(`📊 Detail Table: ${detailRows.length} 筆記錄, ${detailHeaders.length} 個欄位`);
    console.log(`📋 Master欄位: ${masterHeaders.join(', ')}`);
    console.log(`📋 Detail欄位: ${detailHeaders.join(', ')}`);
    
    // 處理資料：篩選追蹤中的任務並關聯最新工作內容
    const processedData = processTaskData(masterHeaders, masterRows, detailHeaders, detailRows);
    
    return {
      headers: ['任務名稱', '任務編號', '最新工作內容說明'],
      rows: processedData,
      data: [['任務名稱', '任務編號', '最新工作內容說明'], ...processedData],
      masterHeaders: masterHeaders,
      detailHeaders: detailHeaders
    };
    
  } catch (error) {
    console.error('❌ 讀取試算表資料失敗:', error);
    throw error;
  }
}

/**
 * 處理任務資料：篩選追蹤中的任務並關聯最新工作內容
 */
function processTaskData(masterHeaders, masterRows, detailHeaders, detailRows) {
  // 找到欄位索引
  const taskNameIndex = masterHeaders.findIndex(h => h === '任務名稱');
  const taskNumberIndex = masterHeaders.findIndex(h => h === '任務編號');
  const trackingIndex = masterHeaders.findIndex(h => h === '是否追蹤');
  const detailTaskNumberIndex = detailHeaders.findIndex(h => h === '任務編號');
  const workContentIndex = detailHeaders.findIndex(h => h === '工作內容說明');
  
  if (taskNameIndex === -1 || taskNumberIndex === -1 || trackingIndex === -1) {
    throw new Error('master_table中找不到必要欄位：任務名稱、任務編號、是否追蹤');
  }
  
  if (detailTaskNumberIndex === -1 || workContentIndex === -1) {
    throw new Error('detail_table中找不到必要欄位：任務編號、工作內容說明');
  }
  
  // 篩選追蹤中的任務
  const trackingTasks = masterRows.filter(row => row[trackingIndex] === '是');
  console.log(`🔍 找到 ${trackingTasks.length} 個追蹤中的任務`);
  
  // 建立任務編號到最新工作內容的映射
  const taskToLatestWork = {};
  detailRows.forEach(row => {
    const taskNumber = row[detailTaskNumberIndex];
    const workContent = row[workContentIndex];
    
    if (taskNumber && workContent) {
      // 如果該任務還沒有工作內容，或這是更新的工作內容，則更新
      if (!taskToLatestWork[taskNumber] || row[0] > taskToLatestWork[taskNumber].rowIndex) {
        taskToLatestWork[taskNumber] = {
          content: workContent,
          rowIndex: row[0] // 假設第一欄是時間戳或序號
        };
      }
    }
  });
  
  // 組合最終資料
  const result = trackingTasks.map(row => {
    const taskName = row[taskNameIndex];
    const taskNumber = row[taskNumberIndex];
    const latestWork = taskToLatestWork[taskNumber] ? taskToLatestWork[taskNumber].content : '無工作內容記錄';
    
    return [taskName, taskNumber, latestWork];
  });
  
  console.log(`✅ 處理完成，共 ${result.length} 筆追蹤任務資料`);
  return result;
}

/**
 * 處理和分析資料
 */
function processData(data) {
  const headers = data.headers;
  const rows = data.rows;
  
  // 基本統計
  const stats = {
    totalRecords: rows.length,
    totalColumns: headers.length,
    columnTypes: {},
    insights: []
  };
  
  // 分析各欄位
  headers.forEach((header, index) => {
    const columnData = rows.map(row => row[index]).filter(cell => cell !== '');
    const columnType = analyzeColumnType(columnData);
    
    stats.columnTypes[header] = {
      type: columnType,
      nonEmptyCount: columnData.length,
      uniqueCount: [...new Set(columnData)].length,
      sampleValues: columnData.slice(0, 5)
    };
    
    // 生成洞察
    if (columnType === 'numeric') {
      const numericData = columnData.map(val => parseFloat(val)).filter(val => !isNaN(val));
      if (numericData.length > 0) {
        const avg = numericData.reduce((a, b) => a + b, 0) / numericData.length;
        const max = Math.max(...numericData);
        const min = Math.min(...numericData);
        stats.insights.push(`欄位 "${header}" 的數值範圍為 ${min.toFixed(2)} 到 ${max.toFixed(2)}，平均值為 ${avg.toFixed(2)}`);
      }
    } else if (columnType === 'categorical') {
      const valueCounts = {};
      columnData.forEach(val => {
        valueCounts[val] = (valueCounts[val] || 0) + 1;
      });
      const topValue = Object.keys(valueCounts).reduce((a, b) => valueCounts[a] > valueCounts[b] ? a : b);
      const percentage = (valueCounts[topValue] / columnData.length * 100).toFixed(1);
      stats.insights.push(`欄位 "${header}" 中最常見的值是 "${topValue}"，佔 ${percentage}%`);
    }
  });
  
  console.log('🔍 資料分析完成');
  console.log(`📈 總記錄數: ${stats.totalRecords}`);
  console.log(`📊 總欄位數: ${stats.totalColumns}`);
  console.log(`💡 洞察數量: ${stats.insights.length}`);
  
  return {
    ...data,
    stats: stats
  };
}

/**
 * 分析欄位類型
 */
function analyzeColumnType(data) {
  if (data.length === 0) return 'empty';
  
  const numericCount = data.filter(val => !isNaN(parseFloat(val)) && isFinite(val)).length;
  const numericRatio = numericCount / data.length;
  
  if (numericRatio > 0.7) {
    return 'numeric';
  } else if (data.length <= 20) {
    return 'categorical';
  } else {
    return 'text';
  }
}

/**
 * 生成圖表
 */
function generateCharts(processedData) {
  const sheet = SpreadsheetApp.getActiveSheet();
  const charts = [];
  
  CONFIG.charts.forEach((chartConfig, index) => {
    try {
      const chart = createChart(sheet, chartConfig, processedData);
      if (chart) {
        charts.push(chart);
        console.log(`📈 圖表 ${index + 1} 生成成功: ${chartConfig.title}`);
      }
    } catch (error) {
      console.error(`❌ 圖表 ${index + 1} 生成失敗:`, error);
    }
  });
  
  return charts;
}

/**
 * 建立單個圖表
 */
function createChart(sheet, chartConfig, processedData) {
  const { type, title, xColumn, yColumn } = chartConfig;
  const headers = processedData.headers;
  
  // 檢查欄位索引是否有效
  if (xColumn >= headers.length || yColumn >= headers.length) {
    throw new Error(`欄位索引超出範圍: xColumn=${xColumn}, yColumn=${yColumn}`);
  }
  
  const dataRange = sheet.getDataRange();
  const chartBuilder = sheet.newChart();
  
  // 根據圖表類型設定
  switch (type) {
    case 'bar':
      chartBuilder.setChartType(Charts.ChartType.COLUMN);
      break;
    case 'pie':
      chartBuilder.setChartType(Charts.ChartType.PIE);
      break;
    case 'line':
      chartBuilder.setChartType(Charts.ChartType.LINE);
      break;
    default:
      chartBuilder.setChartType(Charts.ChartType.COLUMN);
  }
  
  // 設定資料範圍
  chartBuilder.addRange(dataRange);
  
  // 設定圖表位置和大小
  chartBuilder.setPosition(2 + (processedData.stats.totalColumns + 2) * index, 1, 0, 0)
              .setOption('title', title)
              .setOption('width', 400)
              .setOption('height', 300)
              .setOption('legend', { position: 'bottom' });
  
  // 設定X軸和Y軸
  if (type !== 'pie') {
    chartBuilder.setOption('hAxis', { title: headers[xColumn] });
    chartBuilder.setOption('vAxis', { title: headers[yColumn] });
  }
  
  return chartBuilder.build();
}

/**
 * 建立簡報
 */
function createPresentation(processedData) {
  try {
    // 建立新的Google Slides簡報
    const presentation = SlidesApp.create(CONFIG.presentation.title);
    const slides = presentation.getSlides();
    
    // 標題投影片
    const titleSlide = slides[0];
    const titleShape = titleSlide.getShapes()[0];
    const titleText = titleShape.getText();
    titleText.setText(CONFIG.presentation.title);
    titleText.getTextStyle().setBold(true).setFontSize(24);
    
    // 副標題
    const subtitleShape = titleSlide.insertTextBox(CONFIG.presentation.subtitle, 100, 200, 400, 50);
    subtitleShape.getText().getTextStyle().setFontSize(16);
    
    // 資料摘要投影片
    const summarySlide = slides.pushSlide();
    const summaryTitle = summarySlide.insertTextBox('追蹤任務摘要', 50, 50, 400, 30);
    summaryTitle.getText().getTextStyle().setBold(true).setFontSize(20);
    
    // 摘要內容
    const stats = processedData.stats;
    const summaryContent = [
      `追蹤中任務數: ${stats.totalRecords}`,
      `有工作內容的任務: ${processedData.rows.filter(row => row[2] !== '無工作內容記錄').length}`,
      `無工作內容的任務: ${processedData.rows.filter(row => row[2] === '無工作內容記錄').length}`,
      `生成時間: ${new Date().toLocaleString('zh-TW')}`
    ];
    
    const summaryText = summarySlide.insertTextBox(summaryContent.join('\n'), 50, 100, 400, 150);
    summaryText.getText().getTextStyle().setFontSize(14);
    
    // 任務詳細資料投影片
    const taskSlide = slides.pushSlide();
    const taskTitle = taskSlide.insertTextBox('追蹤任務詳細資料', 50, 50, 400, 30);
    taskTitle.getText().getTextStyle().setBold(true).setFontSize(20);
    
    // 建立任務表格
    let taskContent = '任務名稱 | 任務編號 | 最新工作內容說明\n';
    taskContent += '--- | --- | ---\n';
    
    processedData.rows.forEach(row => {
      const taskName = row[0] || '未命名';
      const taskNumber = row[1] || '無編號';
      const workContent = row[2] || '無內容';
      
      // 限制工作內容長度以避免投影片過長
      const shortContent = workContent.length > 50 ? workContent.substring(0, 50) + '...' : workContent;
      
      taskContent += `${taskName} | ${taskNumber} | ${shortContent}\n`;
    });
    
    const taskText = taskSlide.insertTextBox(taskContent, 50, 100, 700, 400);
    taskText.getText().getTextStyle().setFontSize(12);
    
    // 洞察投影片
    if (stats.insights.length > 0) {
      const insightsSlide = slides.pushSlide();
      const insightsTitle = insightsSlide.insertTextBox('分析洞察', 50, 50, 400, 30);
      insightsTitle.getText().getTextStyle().setBold(true).setFontSize(20);
      
      const insightsContent = stats.insights.slice(0, 5).join('\n• ');
      const insightsText = insightsSlide.insertTextBox('• ' + insightsContent, 50, 100, 400, 200);
      insightsText.getText().getTextStyle().setFontSize(14);
    }
    
    console.log(`🎨 簡報建立成功: ${presentation.getUrl()}`);
    return presentation;
    
  } catch (error) {
    console.error('❌ 建立簡報失敗:', error);
    return null;
  }
}

/**
 * 計算空儲存格數量
 */
function getEmptyCellCount(processedData) {
  let emptyCount = 0;
  processedData.rows.forEach(row => {
    row.forEach(cell => {
      if (cell === '' || cell === null || cell === undefined) {
        emptyCount++;
      }
    });
  });
  return emptyCount;
}

/**
 * 匯出結果
 */
function exportResults(processedData) {
  try {
    const folder = DriveApp.getRootFolder();
    const timestamp = new Date().toISOString().replace(/[:.]/g, '-');
    
    // 匯出CSV
    const csvContent = generateCSV(processedData);
    const csvBlob = Utilities.newBlob(csvContent, 'text/csv', `google_form_data_${timestamp}.csv`);
    const csvFile = folder.createFile(csvBlob);
    
    console.log(`💾 CSV檔案已匯出: ${csvFile.getName()}`);
    
    // 建立分析報告
    const reportContent = generateAnalysisReport(processedData);
    const reportBlob = Utilities.newBlob(reportContent, 'text/plain', `analysis_report_${timestamp}.txt`);
    const reportFile = folder.createFile(reportBlob);
    
    console.log(`📊 分析報告已匯出: ${reportFile.getName()}`);
    
    return {
      csvFile: csvFile,
      reportFile: reportFile
    };
    
  } catch (error) {
    console.error('❌ 匯出結果失敗:', error);
    return null;
  }
}

/**
 * 生成CSV內容
 */
function generateCSV(processedData) {
  const headers = processedData.headers;
  const rows = processedData.rows;
  
  let csvContent = headers.join(',') + '\n';
  rows.forEach(row => {
    const csvRow = row.map(cell => {
      if (typeof cell === 'string' && cell.includes(',')) {
        return `"${cell}"`;
      }
      return cell;
    });
    csvContent += csvRow.join(',') + '\n';
  });
  
  return csvContent;
}

/**
 * 生成分析報告
 */
function generateAnalysisReport(processedData) {
  const stats = processedData.stats;
  let report = `任務追蹤分析報告\n`;
  report += `生成時間: ${new Date().toLocaleString('zh-TW')}\n\n`;
  
  report += `=== 基本資訊 ===\n`;
  report += `追蹤中任務數: ${stats.totalRecords}\n`;
  report += `有工作內容的任務: ${processedData.rows.filter(row => row[2] !== '無工作內容記錄').length}\n`;
  report += `無工作內容的任務: ${processedData.rows.filter(row => row[2] === '無工作內容記錄').length}\n\n`;
  
  report += `=== 追蹤任務詳細資料 ===\n`;
  processedData.rows.forEach((row, index) => {
    const taskName = row[0] || '未命名';
    const taskNumber = row[1] || '無編號';
    const workContent = row[2] || '無內容';
    
    report += `${index + 1}. 任務名稱: ${taskName}\n`;
    report += `   任務編號: ${taskNumber}\n`;
    report += `   最新工作內容: ${workContent}\n\n`;
  });
  
  report += `\n=== 主要洞察 ===\n`;
  stats.insights.forEach((insight, index) => {
    report += `${index + 1}. ${insight}\n`;
  });
  
  return report;
}

/**
 * 設定配置
 */
function updateConfig(newConfig) {
  Object.assign(CONFIG, newConfig);
  console.log('✅ 配置已更新');
}

/**
 * 測試函數
 */
function test() {
  console.log('🧪 開始測試...');
  
  // 測試資料讀取
  const data = getCurrentSheetData();
  console.log('✅ 資料讀取測試通過');
  
  // 測試資料處理
  const processedData = processData(data);
  console.log('✅ 資料處理測試通過');
  
  // 測試圖表生成
  const charts = generateCharts(processedData);
  console.log(`✅ 圖表生成測試通過，生成 ${charts.length} 個圖表`);
  
  console.log('🎉 所有測試通過！');
  return true;
}

/**
 * 測試特定試算表連接
 */
function testSpecificSpreadsheet() {
  console.log('🧪 測試特定試算表連接...');
  
  try {
    // 開啟指定的試算表
    const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
    console.log(`✅ 成功開啟試算表: ${spreadsheet.getName()}`);
    
    // 檢查工作表
    const masterSheet = spreadsheet.getSheetByName('master_table');
    const detailSheet = spreadsheet.getSheetByName('detail_table');
    
    if (masterSheet) {
      console.log(`✅ 找到master_table工作表`);
      const masterData = masterSheet.getDataRange().getValues();
      console.log(`   - 欄位: ${masterData[0].join(', ')}`);
      console.log(`   - 資料筆數: ${masterData.length - 1}`);
    } else {
      console.log('❌ 找不到master_table工作表');
    }
    
    if (detailSheet) {
      console.log(`✅ 找到detail_table工作表`);
      const detailData = detailSheet.getDataRange().getValues();
      console.log(`   - 欄位: ${detailData[0].join(', ')}`);
      console.log(`   - 資料筆數: ${detailData.length - 1}`);
    } else {
      console.log('❌ 找不到detail_table工作表');
    }
    
    // 測試資料處理
    const data = getCurrentSheetData();
    console.log(`✅ 資料處理成功，共 ${data.rows.length} 筆追蹤任務`);
    
    // 顯示前3筆資料
    console.log('📋 前3筆追蹤任務資料:');
    data.rows.slice(0, 3).forEach((row, index) => {
      console.log(`   ${index + 1}. ${row[0]} (${row[1]}) - ${row[2].substring(0, 30)}...`);
    });
    
    return true;
    
  } catch (error) {
    console.error('❌ 測試失敗:', error);
    return false;
  }
}
