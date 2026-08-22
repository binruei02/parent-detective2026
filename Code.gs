/**
 * 數位識讀偵探事務所 - Google Apps Script (GAS) 後端控制器
 * 專為班親會宣講打造：提供網頁服務 (Web App)、即時投票統計、家長提問留言板與家庭約定紀錄
 */

// 1. GET 請求進入點：渲染 HTML 網頁
function doGet(e) {
  var template = HtmlService.createTemplateFromFile('Index');
  return template.evaluate()
    .setTitle('數位識讀偵探事務所｜班親會互動共學網頁')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// 2. 輔助函式：若有拆分 CSS/JS 檔案時可透過 <?!= include('FileName'); ?> 引入
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * 3. 取得或自動建立 Google 試算表（用於現場數據紀錄）
 */
function getOrCreateSpreadsheet() {
  var prop = PropertiesService.getScriptProperties();
  var sheetId = prop.getProperty('SPREADSHEET_ID');
  var ss;
  
  if (sheetId) {
    try {
      ss = SpreadsheetApp.openById(sheetId);
      return ss;
    } catch (err) {
      Logger.log('無法開啟現有試算表，將建立新試算表: ' + err.toString());
    }
  }
  
  // 自動建立新試算表
  ss = SpreadsheetApp.create('班親會數位識讀-現場互動統計');
  prop.setProperty('SPREADSHEET_ID', ss.getId());
  
  // 初始化工作表分頁
  var voteSheet = ss.getSheetByName('工作表1') || ss.insertSheet('即時投票紀錄');
  voteSheet.setName('即時投票紀錄');
  voteSheet.appendRow(['時間戳記', '案件編號', '選擇選項', '使用者身分']);
  voteSheet.getRange('A1:D1').setBackground('#1A2938').setFontColor('#FFFFFF').setFontWeight('bold');
  
  var feedbackSheet = ss.insertSheet('現場家長提問與回饋');
  feedbackSheet.appendRow(['時間戳記', '家長稱呼', '分類標籤', '回饋內容']);
  feedbackSheet.getRange('A1:D1').setBackground('#E67E22').setFontColor('#FFFFFF').setFontWeight('bold');
  
  var pledgeSheet = ss.insertSheet('家庭約定認證名冊');
  pledgeSheet.appendRow(['時間戳記', '學生姓名', '家長簽名', '承諾守則項目']);
  pledgeSheet.getRange('A1:D1').setBackground('#27AE60').setFontColor('#FFFFFF').setFontWeight('bold');
  
  return ss;
}

/**
 * 4. 提交投票答案
 * @param {string} caseId 案件代碼 (如 case1)
 * @param {string} optionKey 選項 (A/B/C/D)
 * @param {string} role 身分
 */
function submitPollAnswer(caseId, optionKey, role) {
  try {
    var ss = getOrCreateSpreadsheet();
    var sheet = ss.getSheetByName('即時投票紀錄');
    sheet.appendRow([new Date(), caseId, optionKey, role || '現場家長']);
    return { success: true };
  } catch (err) {
    Logger.log('submitPollAnswer error: ' + err.toString());
    return { success: false, error: err.toString() };
  }
}

/**
 * 5. 提交現場家長提問與心得便利貼
 * @param {string} name 稱呼
 * @param {string} tag 分類
 * @param {string} message 內容
 */
function submitFeedback(name, tag, message) {
  try {
    var ss = getOrCreateSpreadsheet();
    var sheet = ss.getSheetByName('現場家長提問與回饋');
    sheet.appendRow([new Date(), name, tag, message]);
    return { success: true };
  } catch (err) {
    Logger.log('submitFeedback error: ' + err.toString());
    return { success: false, error: err.toString() };
  }
}

/**
 * 6. 提交家庭約定證書簽名
 * @param {string} studentName 學生姓名
 * @param {string} parentName 家長姓名
 * @param {string} pledgeItems 承諾條款
 */
function submitPledge(studentName, parentName, pledgeItems) {
  try {
    var ss = getOrCreateSpreadsheet();
    var sheet = ss.getSheetByName('家庭約定認證名冊');
    sheet.appendRow([new Date(), studentName, parentName, pledgeItems]);
    return { success: true };
  } catch (err) {
    Logger.log('submitPledge error: ' + err.toString());
    return { success: false, error: err.toString() };
  }
}

/**
 * 7. 取得試算表網址（方便講者點擊查看現場完整統計）
 */
function getSpreadsheetUrl() {
  var ss = getOrCreateSpreadsheet();
  return ss.getUrl();
}
