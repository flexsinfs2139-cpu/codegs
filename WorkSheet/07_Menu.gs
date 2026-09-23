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

  // Submenu: Tasks
  const tasksMenu = ui.createMenu('Tasks')
    .addItem('Fill Task for Today', 'fillTaskForToday')
    .addItem('Fill Task for Selected Date', 'fillTaskForSelectedDate');

  // Submenu: DSR
  const dsrMenu = ui.createMenu('DSR')
    .addItem('Generate DSR for Today', 'generateDSRForToday')
    .addItem('Generate DSR for Selected Date', 'generateDSRForSelectedDate');

  // Submenu: Advanced
  const advancedMenu = ui.createMenu('Advanced')
    .addItem('Refresh Dashboard', 'refreshDashboard')
    .addItem('Rebuild Lists', 'rebuildLists')
    .addItem('Create Current Month Sheet', 'createCurrentMonthSheet')
    .addItem('Enforce Single Todo Quadrant', 'sanitizeAllTodoQuadrants')
    .addItem('Populate Dummy Data', 'populateDummyData');

  // Root WorkSheet Menu
  ui.createMenu('WorkSheet')
    .addItem('⚡ Setup / Repair Work Tracker', 'setupWorkTracker')
    .addItem('📊 Open Dashboard', 'openDashboard')
    .addItem('🎯 Open Todo', 'openTodoSheet')
    .addSeparator()
    .addSubMenu(tasksMenu)
    .addSubMenu(dsrMenu)
    .addSeparator()
    .addSubMenu(advancedMenu)
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