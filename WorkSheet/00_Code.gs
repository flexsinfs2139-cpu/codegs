// ============================================================
// CODE.GS — MASTER SETUP & ENTRY POINT FOR WORKSHEET SAAS ENGINE
// ============================================================

/**
 * Master Setup / Repair / Upgrade Function for WorkSheet.
 * Safe, idempotent, and non-destructive.
 *
 * Configures the complete Modern SaaS Work Tracker system:
 * 1. Creates required sheets if missing (Lists, Todo, Month, Dashboard)
 * 2. Verifies and updates the Lists sheet & named ranges
 * 3. Verifies and configures the Eisenhower Matrix Todo sheet
 * 4. Verifies and configures the current Month task sheet
 * 5. Builds and refreshes the Command Center Dashboard
 * 6. Configures headers, dropdowns, conditional formatting, column dimensions
 * 7. Applies Varela Round & Roboto Mono typography standards
 * 8. Configures safe gridline behavior
 * 9. Detects and registers workbook triggers without duplicates
 * 10. Strictly preserves all existing user tasks and entries
 *
 * @param {boolean} [suppressAlert=false] Whether to suppress the completion UI dialog
 */
function setupWorkTracker(suppressAlert = false) {
  console.log('[WorkSheet] Setup started');

  const ss = getSpreadsheet();

  // 1. Lists Sheet & Named Ranges
  ensureListsSheet(ss);
  console.log('[WorkSheet] Lists verified');

  // 2. Eisenhower Matrix Todo Sheet
  createTodoSheet(true);
  console.log('[WorkSheet] Todo verified');

  // 3. Current Month Task Sheet
  const currentMonthSheet = createCurrentMonthSheet(true);
  console.log('[WorkSheet] Current month verified');

  // 4. Command Center Dashboard (14 columns, SaaS layout)
  refreshDashboard(true, true);
  console.log('[WorkSheet] Dashboard verified');

  // 5. Validations & Formatting across all Month Sheets
  setupWorkbookValidationsAndFormatting(ss);
  console.log('[WorkSheet] Validations & Formatting configured');

  // 6. Centralized Triggers Verification (No Duplicates)
  setupTriggers(ss);
  console.log('[WorkSheet] Triggers verified');

  // Set active sheet to Current Month (or Dashboard if preferred)
  if (currentMonthSheet) {
    ss.setActiveSheet(currentMonthSheet);
  }

  console.log('[WorkSheet] Setup completed');

  if (!suppressAlert) {
    try {
      SpreadsheetApp.getUi().alert(
        '⚡ WorkSheet Setup Complete\n\n' +
        'Modern SaaS Work Tracker has been safely verified and configured.\n' +
        'All sheets, reference lists, formulas, and triggers are live.'
      );
    } catch (e) {
      ss.toast('WorkSheet setup completed successfully.', '⚡ WorkSheet', 4);
    }
  }
}


/**
 * Initializes the Work Tracker silently (alias for setupWorkTracker with suppressAlert = true).
 */
function initializeWorkTracker() {
  setupWorkTracker(true);
}


/**
 * Rebuilds the reference Lists sheet and updates all named ranges.
 */
function rebuildLists() {
  createListsSheet();

  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.LISTS_SHEET_NAME);

  if (sheet) {
    ss.setActiveSheet(sheet);
  }
}


/**
 * Applies typography, dimensions, dropdowns, conditional formatting,
 * and safe gridline settings across all month sheets in the workbook.
 *
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
 */
function setupWorkbookValidationsAndFormatting(ss) {
  const monthSheets = getAllMonthSheets(ss);

  monthSheets.forEach(sheet => {
    try {
      // Ensure row 1 header integrity without touching row 2+ data
      sheet
        .getRange(CONFIG.HEADER_ROW, 1, 1, CONFIG.HEADERS.length)
        .setValues([CONFIG.HEADERS]);

      formatWorkTracker(sheet);
      setupDropdowns(sheet);
      setupConditionalFormatting(sheet);
      setupWeekendFormatting(sheet);
      sheet.setFrozenRows(1);
      setSheetGridlinesHidden(sheet, true);
    } catch (err) {
      console.warn(`[WorkSheet] Notice updating sheet ${sheet.getName()}:`, err);
    }
  });
}


/**
 * Verifies and configures required workbook triggers without creating duplicates.
 * Checks ScriptApp.getProjectTriggers() before registering installable triggers.
 *
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
 */
function setupTriggers(ss) {
  try {
    if (typeof ScriptApp === 'undefined' || !ScriptApp.getProjectTriggers) {
      return;
    }

    const triggers = ScriptApp.getProjectTriggers();
    const hasOnChange = triggers.some(t => t.getHandlerFunction() === 'onChange');

    if (!hasOnChange) {
      ScriptApp.newTrigger('onChange')
        .forSpreadsheet(ss)
        .onChange()
        .create();
      console.log('[WorkSheet] Registered onChange installable trigger');
    }
  } catch (err) {
    console.warn('[WorkSheet] Trigger registration notice (skipped or unauthorized):', err.message || err);
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
