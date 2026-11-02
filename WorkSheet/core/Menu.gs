// ============================================================
// MENU.GS — MODERN SAAS TOP-LEVEL TOOLBAR MENUS
// ============================================================

/**
 * Builds four top-level menus on spreadsheet open:
 * 1. WorkSheet  — Hub for workspace navigation (Dashboard, Todo, Decisions, Month) & Master setup
 * 2. Setup      — System repair, migration, reference lists, and dummy data
 * 3. Tasks      — Daily task logging, decision recording, and visual timeline inspection
 * 4. DSR        — Daily Status Report compilation and generation
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();

  // 1. WorkSheet Menu (Navigation & Core Views)
  ui.createMenu('WorkSheet')
    .addItem('📊 Open Dashboard', 'openDashboard')
    .addItem('🎯 Open Todo Backlog', 'openTodoSheet')
    .addItem('📜 Open Decisions Timeline', 'openDecisionsSheet')
    .addItem('📅 Create Current Month Sheet', 'createCurrentMonthSheet')
    .addSeparator()
    .addItem('⚡ Setup / Repair Work Tracker', 'setupWorkTracker')
    .addToUi();

  // 2. Setup Menu (System Administration & Migration)
  ui.createMenu('Setup')
    .addItem('⚡ Setup / Repair Work Tracker', 'setupWorkTracker')
    .addItem('🔄 Refresh Dashboard', 'refreshDashboard')
    .addItem('📋 Rebuild Reference Lists', 'rebuildLists')
    .addItem('📅 Create Current Month Sheet', 'createCurrentMonthSheet')
    .addItem('🚀 Run Migration (Task IDs & Decisions)', 'runFullSystemMigration')
    .addSeparator()
    .addItem('🧪 Populate Dummy Data', 'populateDummyData')
    .addToUi();

  // 3. Tasks Menu (Daily Task Logging & Timeline Tracking)
  ui.createMenu('Tasks')
    .addItem('Fill Task for Today', 'fillTaskForToday')
    .addItem('Fill Task for Selected Date', 'fillTaskForSelectedDate')
    .addSeparator()
    .addItem('📝 Record Decision / Event', 'openRecordDecisionDialog')
    .addItem('🔍 View Task History & Timeline', 'openTaskHistoryDialog')
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


/**
 * Helper to jump directly to the Decisions sheet (or create if missing).
 */
function openDecisionsSheet() {
  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.DECISIONS_SHEET_NAME);
  if (sheet) {
    ss.setActiveSheet(sheet);
  } else {
    ensureDecisionsSheet(ss);
  }
}