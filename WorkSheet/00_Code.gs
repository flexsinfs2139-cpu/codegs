// ============================================================
// CODE.GS — WORK TRACKER ENTRY POINT
// ============================================================


// ============================================================
// INITIALIZE WORK TRACKER
// ============================================================

function initializeWorkTracker() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // 1. Create / reset Lists sheet
  createListsSheet();

  // 2. Create current month sheet
  createCurrentMonthSheet();

  // 3. Make current month sheet active
  const timezone = ss.getSpreadsheetTimeZone();
  const today = new Date();

  const sheetName = Utilities.formatDate(
    today,
    timezone,
    'MMMyy'
  ).toUpperCase();

  const sheet = ss.getSheetByName(sheetName);

  if (sheet) {
    ss.setActiveSheet(sheet);
  }
}


// ============================================================
// SETUP MENU
// ============================================================

function setupWorkTracker() {
  createListsSheet();
  createCurrentMonthSheet();
}


// ============================================================
// REBUILD LISTS
// ============================================================

function rebuildLists() {
  createListsSheet();

  const sheet =
    SpreadsheetApp
      .getActiveSpreadsheet()
      .getSheetByName(CONFIG.LISTS_SHEET_NAME);

  if (sheet) {
    SpreadsheetApp
      .getActiveSpreadsheet()
      .setActiveSheet(sheet);
  }
}

