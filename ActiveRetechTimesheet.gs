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
