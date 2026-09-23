// ============================================================
// CODE.GS — WORK TRACKER ENTRY POINT
// ============================================================


// ============================================================
// INITIALIZE WORK TRACKER
// ============================================================

function initializeWorkTracker() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // 1. Create / reset Lists sheet silently
  createListsSheet(true);

  // 2. Create / setup Todo sheet with Eisenhower Matrix silently
  createTodoSheet(true);

  // 3. Create current month sheet silently
  createCurrentMonthSheet(true);

  // 4. Provision / refresh Command Center Dashboard silently
  refreshDashboard(true, true);

  // 5. Make current month sheet active
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
  createListsSheet(true);
  createTodoSheet(true);
  createCurrentMonthSheet(true);
  refreshDashboard(true, true);
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


// ============================================================
// GLOBAL WORKBOOK TRIGGERS (ON EDIT & ON CHANGE)
// ============================================================

/**
 * Google Apps Script simple trigger: Automatically runs when any cell is edited.
 * Coordinates real-time updates across the entire workbook:
 * 1. Enforces mutually exclusive quadrant checkboxes in the Todo sheet.
 * 2. Automatically updates the Command Center Dashboard with every sheet change.
 *
 * @param {GoogleAppsScript.Events.SheetsOnEdit} e
 */
function onEdit(e) {
  try {
    handleTodoQuadrantExclusiveSelect(e);
  } catch (err) {
    console.warn('Error in handleTodoQuadrantExclusiveSelect:', err);
  }

  try {
    updateDashboardOnChange(e);
  } catch (err) {
    console.warn('Error in updateDashboardOnChange:', err);
  }
}


/**
 * Google Apps Script installable trigger: Automatically runs when sheet structure changes.
 *
 * @param {GoogleAppsScript.Events.SheetsOnChange} e
 */
function onChange(e) {
  try {
    updateDashboardOnChange(e);
  } catch (err) {
    console.warn('Error in onChange:', err);
  }
}

