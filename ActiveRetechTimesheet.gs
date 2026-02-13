// === Config ===
const SPREADSHEET_ID = '1aYwZx1ukgjbxmZC4F8WwMa2YgIrRcanUpOiMncmjzMs';
const HEADER_BG_COLOR = '#78909C';
const ALT_ROW_COLOR = '#EBEFF1';
const DAY_LABELS = ['日', '月', '火', '水', '木', '金', '土'];

// === Helpers ===

/**
 * Returns sheet name in YY年M月 format (e.g. 26年2月)
 */
function getSheetName(date) {
  const year = date.getFullYear() % 100;
  const month = date.getMonth() + 1;
  return year + '年' + month + '月';
}

/**
 * Finds sheet by name or creates it. Clears all content and formatting if it already exists.
 */
function getOrCreateSheet(spreadsheet, sheetName) {
  const existing = spreadsheet.getSheetByName(sheetName);
  if (existing) {
    existing.clear();
    existing.clearFormats();
    return existing;
  }
  return spreadsheet.insertSheet(sheetName);
}

/**
 * Fetches events from default calendar, excludes all-day events, sorted by start time.
 */
function getCalendarEvents(startDate, endDate) {
  const events = CalendarApp.getDefaultCalendar().getEvents(startDate, endDate);
  return events
    .filter(function(e) { return !e.isAllDayEvent(); })
    .sort(function(a, b) { return a.getStartTime() - b.getStartTime(); });
}

/**
 * Writes header section (rows 1-4): month date, title, and column headers.
 */
function writeHeader(sheet, monthStart) {
  sheet.getRange('B1').setValue(monthStart);
  sheet.getRange('B1').setNumberFormat('yyyy-mm-dd');

  const titleRange = sheet.getRange('D2:G2');
  titleRange.merge();
  titleRange.setValue('作業報告書');
  titleRange.setFontWeight('bold');
  titleRange.setHorizontalAlignment('center');

  const headers = ['日付', '曜日', '出社時間', '退社時間', '休憩時間', '勤務時間', '作業内容'];
  const headerRange = sheet.getRange('B4:H4');
  headerRange.setValues([headers]);
  headerRange.setBackground(HEADER_BG_COLOR);
  headerRange.setFontColor('#FFFFFF');
  headerRange.setFontSize(9);
  headerRange.setHorizontalAlignment('center');
}

/**
 * Writes one row per event starting at row 5. Applies alternating row colors.
 * Returns the last data row number.
 */
function writeEventRows(sheet, events) {
  const startRow = 5;
  events.forEach(function(event, i) {
    const row = startRow + i;
    const start = event.getStartTime();
    const end = event.getEndTime();
    const rowData = [
      start,
      DAY_LABELS[start.getDay()],
      start,
      end,
      '0:00',
      '=E' + row + '-D' + row,
      event.getTitle()
    ];
    const range = sheet.getRange('B' + row + ':H' + row);
    range.setValues([rowData]);

    sheet.getRange('B' + row).setNumberFormat('yyyy-mm-dd');
    sheet.getRange('D' + row + ':G' + row).setNumberFormat('h:mm');

    const bgColor = (i % 2 === 1) ? ALT_ROW_COLOR : '#FFFFFF';
    range.setBackground(bgColor);
  });
  return startRow + events.length - 1;
}

/**
 * Writes totals row with "計" label and SUM of working hours.
 */
function writeTotalsRow(sheet, lastDataRow) {
  const totalRow = lastDataRow + 1;
  sheet.getRange('B' + totalRow).setValue('計');
  sheet.getRange('G' + totalRow).setFormula('=SUM(G5:G' + lastDataRow + ')');
  sheet.getRange('G' + totalRow).setNumberFormat('h:mm');

  const range = sheet.getRange('B' + totalRow + ':H' + totalRow);
  range.setBackground(HEADER_BG_COLOR);
  range.setFontColor('#FFFFFF');
  range.setFontWeight('bold');
}

// === Main ===

/**
 * Generates timesheet for the current month from default calendar events.
 */
function generateTimesheet() {
  const now = new Date();
  const monthStart = new Date(now.getFullYear(), now.getMonth(), 1);
  const monthEnd = new Date(now.getFullYear(), now.getMonth() + 1, 0, 23, 59, 59);

  const events = getCalendarEvents(monthStart, monthEnd);
  if (events.length === 0) {
    Logger.log('No events found for this month.');
    return;
  }

  const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheetName = getSheetName(now);
  const sheet = getOrCreateSheet(spreadsheet, sheetName);

  writeHeader(sheet, monthStart);
  const lastDataRow = writeEventRows(sheet, events);
  writeTotalsRow(sheet, lastDataRow);

  sheet.autoResizeColumns(2, 7);
  Logger.log('Timesheet created: ' + sheetName + ' — ' + spreadsheet.getUrl());
}
