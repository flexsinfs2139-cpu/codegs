// ============================================================
// TODO.GS — SIMPLIFIED 5-COLUMN TODO SHEET ENGINE
// ============================================================

/**
 * Menu Action / Programmatic: Creates the Todo sheet if missing, or refreshes
 * its simplified 5-column structure, dropdowns, and formatting in place.
 * Preserves all existing Todo task data.
 *
 * @param {boolean} [suppressAlert=false] Whether to suppress completion alert dialogs.
 */
function createTodoSheet(suppressAlert) {
  const ss = getSpreadsheet();
  const isNew = !ss.getSheetByName(CONFIG.TODO_SHEET_NAME);

  ensureListsSheet(ss);
  const sheet = ensureTodoSheet(ss);

  setupTodoStructure(sheet);
  formatTodoSheet(sheet);
  setupTodoDropdowns(sheet);
  ensureAllTodoRowsHaveIds(sheet);
  updateTodoTaskIdsNamedRange(ss);

  ss.setActiveSheet(sheet);

  if (!suppressAlert) {
    SpreadsheetApp.getUi().alert(
      isNew
        ? 'Todo sheet created successfully (5 columns: ID, Task Description, Project, Due Date, Notes).'
        : 'Todo sheet refreshed successfully with simplified 5-column structure.'
    );
  }
}


/**
 * Returns the Todo sheet, creating an empty one if missing.
 */
function ensureTodoSheet(ss) {
  let sheet = ss.getSheetByName(CONFIG.TODO_SHEET_NAME);
  if (sheet) {
    return sheet;
  }
  sheet = ss.insertSheet(CONFIG.TODO_SHEET_NAME);
  return sheet;
}


/**
 * Sets up the simplified 5-column Todo structure:
 * - Row 1: Table Headers [ID, Task Description, Project, Due Date, Notes]
 * - Row 2+: Backlog task entries
 */
function setupTodoStructure(sheet) {
  const headers = CONFIG.TODO_HEADERS;

  sheet
    .getRange(CONFIG.TODO_HEADER_ROW, 1, 1, headers.length)
    .setValues([headers]);
}


/**
 * Applies header styling, column widths, dimensions, and freezes Row 1.
 * Trims the sheet down to exactly 5 columns.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 */
function formatTodoSheet(sheet) {
  trimTodoSheet(sheet);

  const headers = CONFIG.TODO_HEADERS;
  const lastRow = Math.max(sheet.getLastRow(), 2);

  // 1. Header Row (Row 1)
  sheet
    .getRange(CONFIG.TODO_HEADER_ROW, 1, 1, headers.length)
    .setBackground('#0f172a')
    .setFontColor('#ffffff')
    .setFontFamily(CONFIG.FONTS.TEXT)
    .setFontWeight('bold')
    .setFontSize(10)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');

  sheet.setRowHeight(CONFIG.TODO_HEADER_ROW, 28);

  // 2. Data Rows Base Formatting
  const dataRange = sheet.getRange(CONFIG.TODO_FIRST_DATA_ROW, 1, lastRow - 1, headers.length);
  dataRange
    .setFontFamily(CONFIG.FONTS.TEXT)
    .setFontSize(10)
    .setVerticalAlignment('middle');

  // Column A: ID (Monospace, Bold, Center)
  sheet
    .getRange(CONFIG.TODO_FIRST_DATA_ROW, 1, lastRow - 1, 1)
    .setFontFamily(CONFIG.FONTS.DIGITS)
    .setFontWeight('bold')
    .setHorizontalAlignment('center');

  // Column B: Task Description (Left aligned, text wrap)
  sheet
    .getRange(CONFIG.TODO_FIRST_DATA_ROW, 2, lastRow - 1, 1)
    .setHorizontalAlignment('left')
    .setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);

  // Column C: Project (Center aligned)
  sheet
    .getRange(CONFIG.TODO_FIRST_DATA_ROW, 3, lastRow - 1, 1)
    .setHorizontalAlignment('center');

  // Column D: Due Date (Monospace, Center aligned, yyyy-MM-dd)
  sheet
    .getRange(CONFIG.TODO_FIRST_DATA_ROW, 4, lastRow - 1, 1)
    .setFontFamily(CONFIG.FONTS.DIGITS)
    .setHorizontalAlignment('center')
    .setNumberFormat('yyyy-MM-dd');

  // Column E: Notes (Left aligned, text wrap)
  sheet
    .getRange(CONFIG.TODO_FIRST_DATA_ROW, 5, lastRow - 1, 1)
    .setHorizontalAlignment('left')
    .setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);

  // 3. Column Widths
  const widths = (CONFIG.DIMENSIONS && CONFIG.DIMENSIONS.TODO_COL_WIDTHS) || [95, 360, 140, 110, 280];
  widths.forEach((w, idx) => {
    sheet.setColumnWidth(idx + 1, w);
  });

  // 4. Row Heights for Data Rows
  if (lastRow >= CONFIG.TODO_FIRST_DATA_ROW) {
    sheet.setRowHeights(
      CONFIG.TODO_FIRST_DATA_ROW,
      lastRow - 1,
      (CONFIG.DIMENSIONS && CONFIG.DIMENSIONS.TODO_ROW_HEIGHT) || 36
    );
  }

  // 5. Freeze Header Row & Hide Gridlines
  sheet.setFrozenRows(CONFIG.TODO_HEADER_ROW);
  setSheetGridlinesHidden(sheet, true);

  // 6. Subtle Borders
  formatBorders(dataRange);
}


/**
 * Ensures the sheet has enough rows and trims surplus default columns/rows,
 * preserving all existing Todo data.
 */
function trimTodoSheet(sheet) {
  const requiredRows = Math.max(
    sheet.getLastRow(),
    CONFIG.TODO_DROPDOWN_ROWS + 1
  );

  const maxRows = sheet.getMaxRows();
  if (maxRows < requiredRows) {
    sheet.insertRowsAfter(maxRows, requiredRows - maxRows);
  }

  const maxCols = sheet.getMaxColumns();
  if (maxCols < CONFIG.TODO_HEADERS.length) {
    sheet.insertColumnsAfter(maxCols, CONFIG.TODO_HEADERS.length - maxCols);
  }

  trimSheet(sheet, requiredRows, CONFIG.TODO_HEADERS.length);
}


/**
 * Applies Project dropdown validation sourced from the Lists sheet
 * named range over data rows starting from row 2.
 */
function setupTodoDropdowns(sheet) {
  const ss = sheet.getParent();
  const rowCount = Math.max(
    sheet.getLastRow() - 1,
    CONFIG.TODO_DROPDOWN_ROWS
  );

  const rule = createDropdownRule(ss, 'Projects');
  if (rule) {
    sheet
      .getRange(CONFIG.TODO_FIRST_DATA_ROW, 3, rowCount, 1)
      .setDataValidation(rule);
  }
}


/**
 * Scans all rows in the Todo sheet and ensures every non-empty task description
 * has a permanent, unique Task ID.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 */
function ensureAllTodoRowsHaveIds(sheet) {
  if (!sheet) return;
  const lastRow = sheet.getLastRow();
  if (lastRow < CONFIG.TODO_FIRST_DATA_ROW) return;

  const numRows = lastRow - CONFIG.TODO_HEADER_ROW;
  const range = sheet.getRange(CONFIG.TODO_FIRST_DATA_ROW, 1, numRows, 2);
  const values = range.getValues();
  let modified = false;

  for (let i = 0; i < values.length; i++) {
    const id = String(values[i][0] || '').trim();
    const desc = String(values[i][1] || '').trim();

    if (desc && !id) {
      values[i][0] = generateNextTaskId();
      modified = true;
    }
  }

  if (modified) {
    range.setValues(values);
    console.log('[Todo] Auto-assigned missing Task IDs in Todo sheet');
  }
}


/**
 * Updates or creates the 'TodoTaskIds' named range referencing Column A of Todo sheet.
 *
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} [ss]
 */
function updateTodoTaskIdsNamedRange(ss) {
  if (!ss) ss = getSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.TODO_SHEET_NAME);
  if (!sheet) return;

  const lastRow = Math.max(sheet.getLastRow(), 2);
  const range = sheet.getRange(CONFIG.TODO_FIRST_DATA_ROW, 1, lastRow - 1, 1);
  ss.setNamedRange(CONFIG.NAMED_RANGES.TODO_TASK_IDS || 'TodoTaskIds', range);
}
