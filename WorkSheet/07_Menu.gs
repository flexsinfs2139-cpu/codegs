// ============================================================
// MENU.GS — UNIFIED MODERN SAAS APPLICATION TOOLBAR MENU
// ============================================================

/**
 * Builds the single, unified 'WorkSheet' menu on spreadsheet open.
 * Groups primary actions, rapid tasks entry, DSR reporting, and advanced utilities
 * into a modern SaaS hierarchy.
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();

  // 1. WorkSheet Menu (Navigation & Core Actions)
  ui.createMenu('WorkSheet')
    .addItem('📊 Open Dashboard', 'openDashboard')
    .addItem('🎯 Open Todo Sheet', 'openTodoSheet')
    .addItem('📅 Create Current Month Sheet', 'createCurrentMonthSheet')
    .addSeparator()
    .addItem('⚡ Setup / Repair Work Tracker', 'setupWorkTracker')
    .addToUi();

  // 2. Setup Menu (Combined Setup & Advanced Administration)
  ui.createMenu('Setup')
    .addItem('⚡ Setup / Repair Work Tracker', 'setupWorkTracker')
    .addItem('🔄 Refresh Dashboard', 'refreshDashboard')
    .addItem('📋 Rebuild Reference Lists', 'rebuildLists')
    .addItem('📅 Create Current Month Sheet', 'createCurrentMonthSheet')
    .addItem('🎯 Enforce Single Todo Quadrant', 'sanitizeAllTodoQuadrants')
    .addSeparator()
    .addItem('🧪 Populate Dummy Data', 'populateDummyData')
    .addToUi();

  // 3. Tasks Menu (Daily Task Logging)
  ui.createMenu('Tasks')
    .addItem('Fill Task for Today', 'fillTaskForToday')
    .addItem('Fill Task for Selected Date', 'fillTaskForSelectedDate')
    .addToUi();

  // 4. DSR Menu (Daily Status Report Generation)
  ui.createMenu('DSR')
    .addItem('Generate DSR for Today', 'generateDSRForToday')
    .addItem('Generate DSR for Selected Date', 'generateDSRForSelectedDate')
    .addToUi();
}


/**
 * Helper to jump directly to the Dashboard sheet (or build if missing).
 */
function openDashboard() {
  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.DASHBOARD_SHEET_NAME);
  if (sheet) {
    ss.setActiveSheet(sheet);
  } else {
    refreshDashboard();
  }
}


/**
 * Helper to jump directly to the Todo sheet (or create if missing).
 */
function openTodoSheet() {
  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.TODO_SHEET_NAME);
  if (sheet) {
    ss.setActiveSheet(sheet);
  } else {
    createTodoSheet();
  }
}