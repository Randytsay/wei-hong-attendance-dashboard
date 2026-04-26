/**
 * 韋宏大C 出席儀表板 - GAS 後端
 * 資料來源: Google Sheets "表單回應 1"
 * 
 * 部署方式:
 * 1. 在 GAS 編輯器中建立新專案
 * 2. 貼上此檔案內容
 * 3. 部署 → 新增部署作業 → 類型: 網頁應用程式 → 設定「任何人」存取
 * 4. 複製 the exec URL 到 HTML 中的 google.script.run 呼叫
 */

// ===== 設定 =====
const SPREADSHEET_ID = '1NvM2cZEeLWScclaoO6Lf0JpSNuaBhxms9P4UdZSPJSk';
const SHEET_NAME = '表單回應 1';
const CACHE_TTL = 300; // 5分鐘快取

// ===== 快取 =====
let cache = CacheService.getScriptCache();

// ===== 工具函式 =====

function parseDate(dateStr) {
  if (!dateStr) return null;
  // 支援格式: 2023/1/5, 2023/01/05, 日期物件
  if (dateStr instanceof Date) return dateStr;
  const str = String(dateStr).trim();
  const parts = str.match(/(\d{1,4})[\/\-](\d{1,2})[\/\-](\d{1,4})/);
  if (parts) {
    const [, y, m, d] = parts;
    // 判斷年份格式
    const year = y.length === 2 ? 2000 + parseInt(y) : parseInt(y);
    const month = parseInt(m) - 1;
    const day = parseInt(d);
    return new Date(year, month, day);
  }
  const parsed = new Date(str);
  return isNaN(parsed.getTime()) ? null : parsed;
}

function formatDateKey(date) {
  if (!date || !(date instanceof Date)) return null;
  const y = date.getFullYear();
  const m = String(date.getMonth() + 1).padStart(2, '0');
  const d = String(date.getDate()).padStart(2, '0');
  return `${y}-${m}-${d}`;
}

function normalizeSubgroupName(name) {
  if (!name) return '未知小C';
  return name.trim()
    .replace(/[Cc]$/, 'C')  // 統一小寫c結尾
    .replace(/Ｃ$/, 'C')    // 全形C轉半形
    .replace(/小[cC]$/, '小C')
    .replace(/[ˇ̂ʹ′´`]//g, '') // 移除重音符號
    .trim();
}

// 取得子表格
function getSheet() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  return ss.getSheetByName(SHEET_NAME) || ss.getSheets()[0];
}

// 讀取原始資料（含快取）
function getRawData() {
  const cacheKey = 'rawData_v2';
  const cached = cache.get(cacheKey);
  if (cached) {
    try {
      return JSON.parse(cached);
    } catch (e) {
      // invalid cache, continue
    }
  }

  const sheet = getSheet();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  // 讀取所有資料列 (跳過標題)
  // 欄位: A=時間戳記, B=簽到日期, C=夥伴姓名, D=所屬小C名稱, E=帶來新朋友人數
  const values = sheet.getRange(2, 1, lastRow - 1, 5).getValues();
  
  const data = values.map((row, idx) => ({
    timestamp: row[0],
    signDate: row[1],
    partnerName: String(row[2] || '').trim(),
    subgroupName: normalizeSubgroupName(row[3]),
    newFriendsBrought: parseInt(row[4]) || 0,
    rowNum: idx + 2
  })).filter(r => r.partnerName && r.signDate);

  cache.put(cacheKey, JSON.stringify(data), CACHE_TTL);
  return data;
}

// 依日期分組資料
function groupByDate(data) {
  const grouped = {};
  data.forEach(row => {
    const date = parseDate(row.signDate);
    if (!date) return;
    const key = formatDateKey(date);
    if (!grouped[key]) grouped[key] = [];
    grouped[key].push(row);
  });
  return grouped;
}

// ===== GAS 公開函式 =====

/**
 * 取得所有不重複的簽到日期（依日期分組用）
 */
function getUniqueMeetingDates() {
  try {
    const data = getRawData();
    const byDate = groupByDate(data);
    return Object.keys(byDate).sort((a, b) => new Date(b) - new Date(a));
  } catch (error) {
    console.error('getUniqueMeetingDates error:', error);
    throw error;
  }
}

/**
 * 取得指定日期的分析資料
 * @param {string} dateStr - YYYY-MM-DD 格式
 * @param {boolean} isInitial
 */
function getAnalyticsData(dateStr, isInitial = false) {
  try {
    const data = getRawData();
    const byDate = groupByDate(data);
    const dayData = byDate[dateStr] || [];

    if (dayData.length === 0) {
      return {
        totalPartners: 0,
        totalNewFriends: 0,
        grandTotal: 0,
        subGroupData: []
      };
    }

    // 依小C分組
    const subgroupMap = {};
    dayData.forEach(row => {
      const name = row.subgroupName;
      if (!subgroupMap[name]) {
        subgroupMap[name] = {
          name: name,
          partners: 0,
          newFriends: 0,
          total: 0,
          attendees: []
        };
      }
      // 所有報到的人都算夥伴（新朋友人數欄位目前皆為0）
      subgroupMap[name].partners += 1;
      subgroupMap[name].newFriends += row.newFriendsBrought;
      subgroupMap[name].total += (1 + row.newFriendsBrought);
      subgroupMap[name].attendees.push({
        name: row.partnerName,
        type: row.newFriendsBrought > 0 ? 'friend' : 'partner'
      });
    });

    const subGroupData = Object.values(subgroupMap).sort((a, b) => b.total - a.total);
    const grandTotal = subGroupData.reduce((sum, g) => sum + g.total, 0);
    const totalPartners = subGroupData.reduce((sum, g) => sum + g.partners, 0);
    const totalNewFriends = subGroupData.reduce((sum, g) => sum + g.newFriends, 0);

    return {
      totalPartners,
      totalNewFriends,
      grandTotal,
      subGroupData
    };
  } catch (error) {
    console.error('getAnalyticsData error:', error);
    throw error;
  }
}

/**
 * 趨勢分析資料
 * @param {string} startDateStr - YYYY-MM-DD
 * @param {string} endDateStr - YYYY-MM-DD
 * @param {boolean} isQuickSelect
 */
function getTrendAnalysisData(startDateStr, endDateStr, isQuickSelect = false) {
  try {
    const data = getRawData();
    const byDate = groupByDate(data);
    
    const start = new Date(startDateStr);
    const end = new Date(endDateStr);
    
    // 收集日期範圍內的所有資料
    const filteredData = [];
    Object.keys(byDate).forEach(dateKey => {
      const d = new Date(dateKey);
      if (d >= start && d <= end) {
        filteredData.push(...byDate[dateKey]);
      }
    });

    // 計算統計
    const dateSet = new Set();
    let totalPartners = 0, totalNewFriends = 0, grandTotal = 0;
    
    Object.keys(byDate).forEach(dateKey => {
      const d = new Date(dateKey);
      if (d >= start && d <= end) {
        dateSet.add(dateKey);
        byDate[dateKey].forEach(row => {
          totalPartners += 1;
          totalNewFriends += row.newFriendsBrought;
          grandTotal += (1 + row.newFriendsBrought);
        });
      }
    });

    // 計算成長率（與前期相比）
    const rangeDays = Math.ceil((end - start) / (1000 * 60 * 60 * 24)) + 1;
    const halfRange = Math.floor(rangeDays / 2);
    const midDate = new Date(start.getTime() + halfRange * 24 * 60 * 60 * 1000);
    
    let firstHalfPartners = 0, secondHalfPartners = 0;
    let firstHalfNewFriends = 0, secondHalfNewFriends = 0;
    
    Object.keys(byDate).forEach(dateKey => {
      const d = new Date(dateKey);
      if (d >= start && d < midDate) {
        byDate[dateKey].forEach(row => {
          firstHalfPartners += 1;
          firstHalfNewFriends += row.newFriendsBrought;
        });
      } else if (d >= midDate && d <= end) {
        byDate[dateKey].forEach(row => {
          secondHalfPartners += 1;
          secondHalfNewFriends += row.newFriendsBrought;
        });
      }
    });

    const partnerGrowthRate = firstHalfPartners > 0 
      ? ((secondHalfPartners - firstHalfPartners) / firstHalfPartners * 100) 
      : (secondHalfPartners > 0 ? Infinity : null);
    const newFriendGrowthRate = firstHalfNewFriends > 0 
      ? ((secondHalfNewFriends - firstHalfNewFriends) / firstHalfNewFriends * 100) 
      : (secondHalfNewFriends > 0 ? Infinity : null);

    // 圖表資料：依週分組
    const weeklyData = {};
    Object.keys(byDate).forEach(dateKey => {
      const d = new Date(dateKey);
      if (d >= start && d <= end) {
        // 取得該週的週一
        const day = d.getDay();
        const mondayOffset = day === 0 ? -6 : 1 - day;
        const monday = new Date(d);
        monday.setDate(d.getDate() + mondayOffset);
        const weekKey = formatDateKey(monday);
        
        if (!weeklyData[weekKey]) {
          weeklyData[weekKey] = { partners: 0, newFriends: 0, total: 0, dates: new Set() };
        }
        byDate[dateKey].forEach(row => {
          weeklyData[weekKey].partners += 1;
          weeklyData[weekKey].newFriends += row.newFriendsBrought;
          weeklyData[weekKey].total += (1 + row.newFriendsBrought);
          weeklyData[weekKey].dates.add(dateKey);
        });
      }
    });

    const weekLabels = Object.keys(weeklyData).sort();
    
    // 依小C分組趨勢
    const subgroupTrendMap = {};
    Object.keys(byDate).forEach(dateKey => {
      const d = new Date(dateKey);
      if (d >= start && d <= end) {
        byDate[dateKey].forEach(row => {
          const name = row.subgroupName;
          if (!subgroupTrendMap[name]) {
            subgroupTrendMap[name] = {};
          }
          if (!subgroupTrendMap[name][dateKey]) {
            subgroupTrendMap[name][dateKey] = 0;
          }
          subgroupTrendMap[name][dateKey] += (1 + row.newFriendsBrought);
        });
      }
    });

    const trendData = {
      labels: weekLabels.map(d => {
        const date = new Date(d);
        return `${date.getMonth() + 1}/${date.getDate()}`;
      }),
      datasets: [{
        label: '總出席人數',
        data: weekLabels.map(w => weeklyData[w]?.total || 0)
      }],
      subGroupDatasets: Object.keys(subgroupTrendMap).slice(0, 8).map(name => ({
        label: name,
        data: weekLabels.map(w => {
          // 取得該週每天的總和
          let weekTotal = 0;
          const monday = new Date(w);
          for (let i = 0; i < 7; i++) {
            const checkDate = new Date(monday);
            checkDate.setDate(monday.getDate() + i);
            const dateKey = formatDateKey(checkDate);
            if (subgroupTrendMap[name][dateKey]) {
              weekTotal += subgroupTrendMap[name][dateKey];
            }
          }
          return weekTotal;
        })
      }))
    };

    const analysis = {
      partnerGrowthRate,
      newFriendGrowthRate,
      averageAttendance: dateSet.size > 0 ? grandTotal / dateSet.size : 0
    };

    return {
      totalPartners,
      totalNewFriends,
      grandTotal,
      analysis,
      trendData
    };
  } catch (error) {
    console.error('getTrendAnalysisData error:', error);
    throw error;
  }
}

/**
 * 歷史比較資料
 * @param {string} period - 'week', 'month', 'quarter', 'half', 'year'
 */
function getComparisonData(period) {
  try {
    const data = getRawData();
    const byDate = groupByDate(data);
    const today = new Date();
    
    // 計算當期和前期區間
    const rangeMap = {
      week: 7,
      month: 30,
      quarter: 90,
      half: 180,
      year: 365
    };
    
    const rangeDays = rangeMap[period] || 30;
    const halfRange = Math.floor(rangeDays / 2);
    
    // 當期區間
    const currentEnd = new Date(today);
    const currentStart = new Date(today);
    currentStart.setDate(today.getDate() - rangeDays + 1);
    
    // 前期區間
    const previousEnd = new Date(currentStart);
    previousEnd.setDate(previousEnd.getDate() - 1);
    const previousStart = new Date(previousEnd);
    previousStart.setDate(previousEnd.getDate() - rangeDays + 1);
    
    const periodLabelMap = {
      week: '週',
      month: '月',
      quarter: '季',
      half: '半年',
      year: '年'
    };
    
    function calcPeriod(start, end) {
      let grandTotal = 0, totalPartners = 0, totalNewFriends = 0;
      Object.keys(byDate).forEach(dateKey => {
        const d = new Date(dateKey);
        if (d >= start && d <= end) {
          byDate[dateKey].forEach(row => {
            totalPartners += 1;
            totalNewFriends += row.newFriendsBrought;
            grandTotal += (1 + row.newFriendsBrought);
          });
        }
      });
      return { grandTotal, totalPartners, totalNewFriends };
    }
    
    const currentPeriod = {
      ...calcPeriod(currentStart, currentEnd),
      label: `本期${periodLabelMap[period]}`,
      start: formatDateKey(currentStart),
      end: formatDateKey(currentEnd)
    };
    
    const previousPeriod = {
      ...calcPeriod(previousStart, previousEnd),
      label: `上期${periodLabelMap[period]}`,
      start: formatDateKey(previousStart),
      end: formatDateKey(previousEnd)
    };
    
    const calcRate = (curr, prev) => {
      if (prev === 0) return curr > 0 ? Infinity : null;
      return (curr - prev) / prev * 100;
    };
    
    return {
      currentPeriod,
      previousPeriod,
      growthRates: {
        grandTotal: calcRate(currentPeriod.grandTotal, previousPeriod.grandTotal),
        partner: calcRate(currentPeriod.totalPartners, previousPeriod.totalPartners),
        newFriend: calcRate(currentPeriod.totalNewFriends, previousPeriod.totalNewFriends)
      }
    };
  } catch (error) {
    console.error('getComparisonData error:', error);
    throw error;
  }
}

// ===== 清除快取（除錯用）====
function clearCache() {
  cache.remove('rawData_v2');
  return 'Cache cleared';
}

// ===== 測試函式 =====
function testGetData() {
  const dates = getUniqueMeetingDates();
  console.log('Available dates count:', dates.length);
  console.log('Latest 5 dates:', dates.slice(0, 5));
  
  if (dates.length > 0) {
    const analytics = getAnalyticsData(dates[0]);
    console.log('First date analytics:', JSON.stringify(analytics, null, 2));
  }
}
