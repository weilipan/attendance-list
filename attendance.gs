function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('點名工具')
    .addItem('產生點名單', 'createAttendanceSheet')
    .addToUi();
}

function createAttendanceSheet() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
   sheet.clear();
  
  // 插入配分標準說明行
  var explanation = [
    "出席成績計算方式：準時(2分)、公假(1.8分)、事病假(1.7分)、5-15分內(1.5分)、15-25分內(1分)、缺席(0分)。最終成績 = 平均出席分數 × 50"
  ];
  sheet.appendRow(explanation);
  sheet.getRange(1, 1, 1, 4).setFontWeight("bold").setHorizontalAlignment("left").merge().setWrap(true);
  
  // 取得使用者輸入的課程起始日期與週數
  var startDateStr = Browser.inputBox("請輸入課程起始日期 (YYYYMMDD)", Browser.Buttons.OK_CANCEL);
  if (startDateStr == "cancel") return;
  var weeks = Browser.inputBox("請輸入課程週數 (例如: 16)", Browser.Buttons.OK_CANCEL);
  if (weeks == "cancel") return;
  
  var startDate = new Date(startDateStr.substring(0, 4), parseInt(startDateStr.substring(4, 6)) - 1, startDateStr.substring(6, 8));
  weeks = parseInt(weeks);
  
  // 設定標題列
  var headers = ["學號", "班級", "座號", "姓名"];
  for (var i = 0; i < weeks; i++) {
    var currentDate = new Date(startDate);
    currentDate.setDate(currentDate.getDate() + (i * 7));
    headers.push(Utilities.formatDate(currentDate, Session.getScriptTimeZone(), "MM/dd"));
  }
  headers.push("出席成績");
  
  sheet.appendRow(headers);
  
  // 凍結標題列
  sheet.setFrozenRows(2);
  sheet.setFrozenColumns(4);
  
  // 設定出席選單（僅限於週次欄位）
  var rule = SpreadsheetApp.newDataValidation().requireValueInList(["準時", "公假", "事病假", "5-15分內", "15-25分內", "缺席"], true).build();
  sheet.getRange(3, 5, 50, weeks).setDataValidation(rule);
  
  // 在最後一列的出席成績欄設定公式，確保未填入的週次不納入計算
  var lastCol = 4 + weeks; // 確保最後一欄是「出席成績」，不多加
  
  // 先清除該欄的資料驗證規則
  sheet.getRange(3, lastCol+1, 50,2).clearDataValidations();
  for (var i = 3; i <= 52; i++) {
    var formula = `=IF(COUNTA(E${i}:${String.fromCharCode(68 + weeks - 1)}${i})=0, "", AVERAGE(FILTER(IF(E${i}:${String.fromCharCode(68 + weeks - 1)}${i}="準時", 2, IF(E${i}:${String.fromCharCode(68 + weeks - 1)}${i}="公假", 1.8, IF(E${i}:${String.fromCharCode(68 + weeks - 1)}${i}="事病假", 1.7, IF(E${i}:${String.fromCharCode(68 + weeks - 1)}${i}="5-15分內", 1.5, IF(E${i}:${String.fromCharCode(68 + weeks - 1)}${i}="15-25分內", 1, IF(E${i}:${String.fromCharCode(68 + weeks - 1)}${i}="缺席", 0, "")))))), E${i}:${String.fromCharCode(68 + weeks - 1)}${i}<>"")) * 50)`;
    sheet.getRange(i, lastCol+1).setFormula(formula);
  }
  
}
