// ============================================================
// TODO.GS — TODO SHEET SETUP & MANAGEMENT
// ============================================================

/**
 * Menu Action: Creates the Todo sheet if missing, or refreshes its
 * formatting/dropdowns in place. Never touches existing Todo data.
 */
function createTodoSheet() {
  const ss = getSpreadsheet();
  const isNew = !ss.getSheetByName(CONFIG.TODO_SHEET_NAME);

  ensureListsSheet(ss);
  const sheet = ensureTodoSheet(ss);

  formatTodoSheet(sheet);
  setupTodoDropdowns(sheet);

  ss.setActiveSheet(sheet);

  SpreadsheetApp.getUi().alert(
    isNew
      ? 'Todo sheet created successfully.'
      : 'Todo sheet already exists — formatting and dropdowns refreshed.'
  );
}


/**
 * Returns the Todo sheet, creating it with its header row if missing.
 * Existing data is never cleared or overwritten.
 */
function ensureTodoSheet(ss) {
  let sheet = ss.getSheetByName(CONFIG.TODO_SHEET_NAME);

  if (sheet) {
    return sheet;
  }

  sheet = ss.insertSheet(CONFIG.TODO_SHEET_NAME);

  sheet
    .getRange(1, 1, 1, CONFIG.TODO_HEADERS.length)
    .setValues([CONFIG.TODO_HEADERS]);

  return sheet;
}


/**
 * Applies header styling, column widths, and a frozen header row.
 * Trims the sheet down to its essential 3 columns. Only touches
 * formatting/dimensions — safe to re-run against a sheet with existing data.
 */
function formatTodoSheet(sheet) {
  trimTodoSheet(sheet);

  sheet
    .getRange(1, 1, 1, CONFIG.TODO_HEADERS.length)
    .setFontFamily('Arial')
    .setFontWeight('bold')
    .setBackground(CONFIG.COLORS.HEADER)
    .setHorizontalAlignment('center');

  sheet.setColumnWidth(1, 340); // Task Name
  sheet.setColumnWidth(2, 130); // Project
  sheet.setColumnWidth(3, 100); // Priority

  sheet.setRowHeight(1, 28);
  sheet.setFrozenRows(1);
}


/**
 * Removes surplus default columns/rows beyond the essential Task Name,
 * Project, Priority columns, preserving any existing Todo data.
 */
function trimTodoSheet(sheet) {
  const requiredRows = Math.max(
    sheet.getLastRow(),
    CONFIG.TODO_DROPDOWN_ROWS + 1
  );

  trimSheet(sheet, requiredRows, CONFIG.TODO_HEADERS.length);
}


/**
 * Applies Project and Priority dropdown validation sourced from the
 * existing Lists sheet named ranges (same ranges used by the month sheet),
 * over a generous block of rows so newly typed Todo rows keep working.
 */
function setupTodoDropdowns(sheet) {
  const ss = sheet.getParent();

  const rowCount = Math.max(
    sheet.getLastRow() + CONFIG.TODO_DROPDOWN_ROWS,
    CONFIG.TODO_DROPDOWN_ROWS
  );

  const rules = {
    2: createDropdownRule(ss, 'Projects'),
    3: createDropdownRule(ss, 'Priorities')
  };

  Object.entries(rules).forEach(([column, rule]) => {
    if (rule) {
      sheet
        .getRange(2, Number(column), rowCount, 1)
        .setDataValidation(rule);
    }
  });
}
