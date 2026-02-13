// === Config ===
var SPREADSHEET_ID = '1aYwZx1ukgjbxmZC4F8WwMa2YgIrRcanUpOiMncmjzMs';
var HEADER_BG_COLOR = '#78909C';
var ALT_ROW_COLOR = '#EBEFF1';
var DAY_LABELS = ['日', '月', '火', '水', '木', '金', '土'];

// === Helpers ===

/**
 * Returns sheet name in YY年M月 format (e.g. 26年2月)
 */
function getSheetName(date) {
  var year = date.getFullYear() % 100;
  var month = date.getMonth() + 1;
  return year + '年' + month + '月';
}
