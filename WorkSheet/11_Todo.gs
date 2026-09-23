// ============================================================
// TODO.GS — TODO SHEET SETUP & MANAGEMENT (EISENHOWER MATRIX)
// ============================================================

/**
 * Menu Action: Creates the Todo sheet if missing, or refreshes its
 * Eisenhower Matrix formatting, checkboxes, and stats row in place.
 * Preserves all existing Todo task data.
 */
function createTodoSheet() {
  const ss = getSpreadsheet();
  const isNew = !ss.getSheetByName(CONFIG.TODO_SHEET_NAME);

  ensureListsSheet(ss);
  const sheet = ensureTodoSheet(ss);

  setupTodoStructure(sheet);
  formatTodoSheet(sheet);
  setupTodoDropdowns(sheet);
  setupTodoCheckboxes(sheet);
  setupTodoConditionalFormatting(sheet);

  ss.setActiveSheet(sheet);

  SpreadsheetApp.getUi().alert(
    isNew
      ? 'Todo sheet created successfully with Eisenhower Matrix.'
      : 'Todo sheet already exists — Eisenhower Matrix formatting, checkboxes, and stats row refreshed.'
  );
}


/**
 * Returns the Todo sheet, creating an empty one if missing.
 * Structure setup and formatting are handled by setupTodoStructure/formatTodoSheet.
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
 * Sets up the Eisenhower Matrix structure:
 * - Row 1: Live Stats row (Total Tasks, Q1, Q2, Q3, Q4)
 * - Row 2: Table Header row (Task Name, Project, Q1: Do, Q2: Schedule, Q3: Delegate, Q4: Don't Do)
 * Handles legacy 1-row header migration seamlessly by inserting a row before Row 1.
 */
function setupTodoStructure(sheet) {
  // If the sheet already exists with legacy 1-row header layout,
  // insert a row before row 1 so data safely shifts down to row 3.
  if (sheet.getLastRow() >= 1) {
    const r1c1 = String(sheet.getRange(1, 1).getValue()).trim();
    if (r1c1 === 'Task Name') {
      sheet.insertRowBefore(1);
    }
  }

  // Row 1: Eisenhower Matrix Stats Row
  const statsRange = sheet.getRange('A1:B1');
  if (!statsRange.isPartOfMerge()) {
    statsRange.merge();
  }
  sheet.getRange('A1').setValue('="📋 Total Tasks: " & COUNTA(A3:A)');
  sheet.getRange(1, 3).setValue('="🔴 Q1: " & COUNTIF(C3:C, TRUE)');
  sheet.getRange(1, 4).setValue('="🔵 Q2: " & COUNTIF(D3:D, TRUE)');
  sheet.getRange(1, 5).setValue('="🟡 Q3: " & COUNTIF(E3:E, TRUE)');
  sheet.getRange(1, 6).setValue('="⚪ Q4: " & COUNTIF(F3:F, TRUE)');

  // Row 2: Table Headers
  sheet
    .getRange(2, 1, 1, CONFIG.TODO_HEADERS.length)
    .setValues([CONFIG.TODO_HEADERS]);
}


/**
 * Applies header styling, column widths, dimensions, and freezes the top 2 rows
 * (Stats row + Table header row). Trims the sheet down to the 6 essential columns.
 */
function formatTodoSheet(sheet) {
  trimTodoSheet(sheet);

  // 1. Stats Row (Row 1)
  const statsRange = sheet.getRange(1, 1, 1, CONFIG.TODO_HEADERS.length);
  statsRange
    .setFontFamily('Arial')
    .setFontWeight('bold')
    .setFontSize(10)
    .setBackground(CONFIG.COLORS.STATS_BG)
    .setVerticalAlignment('middle');

  sheet.getRange(1, 1).setHorizontalAlignment('left');
  sheet.getRange(1, 3, 1, 4).setHorizontalAlignment('center');

  statsRange.setBorder(
    null, null, true, null, null, null,
    CONFIG.COLORS.STATS_BORDER,
    SpreadsheetApp.BorderStyle.SOLID_MEDIUM
  );

  // 2. Table Headers (Row 2)
  const headerRange = sheet.getRange(2, 1, 1, CONFIG.TODO_HEADERS.length);
  headerRange
    .setFontFamily('Arial')
    .setFontWeight('bold')
    .setFontSize(10)
    .setBackground(CONFIG.COLORS.HEADER)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setBorder(
      null, null, true, null, null, null,
      CONFIG.COLORS.BORDER,
      SpreadsheetApp.BorderStyle.SOLID
    );

  // 3. Column Widths
  sheet.setColumnWidth(1, 320); // Task Name
  sheet.setColumnWidth(2, 130); // Project
  sheet.setColumnWidth(3, 115); // Q1: Do
  sheet.setColumnWidth(4, 115); // Q2: Schedule
  sheet.setColumnWidth(5, 115); // Q3: Delegate
  sheet.setColumnWidth(6, 115); // Q4: Don't Do

  // 4. Row Heights
  sheet.setRowHeight(1, 32);
  sheet.setRowHeight(2, 28);

  // 5. Freeze top 2 rows (Stats row + Table headers)
  sheet.setFrozenRows(2);
}


/**
 * Ensures the sheet has enough rows and trims surplus default columns/rows,
 * preserving all existing Todo data.
 */
function trimTodoSheet(sheet) {
  const requiredRows = Math.max(
    sheet.getLastRow(),
    CONFIG.TODO_DROPDOWN_ROWS + 2
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
 * named range over data rows starting from row 3.
 */
function setupTodoDropdowns(sheet) {
  const ss = sheet.getParent();

  const rowCount = Math.max(
    sheet.getLastRow() - 2,
    CONFIG.TODO_DROPDOWN_ROWS
  );

  const rule = createDropdownRule(ss, 'Projects');
  if (rule) {
    sheet
      .getRange(CONFIG.TODO_FIRST_DATA_ROW, 2, rowCount, 1)
      .setDataValidation(rule);
  }
}


/**
 * Sets up native Google Sheets checkboxes across the 4 Eisenhower Matrix
 * quadrant columns (Cols 3–6: Q1 Do, Q2 Schedule, Q3 Delegate, Q4 Don't Do)
 * starting from row 3.
 */
function setupTodoCheckboxes(sheet) {
  const rowCount = Math.max(
    sheet.getLastRow() - 2,
    CONFIG.TODO_DROPDOWN_ROWS
  );

  sheet
    .getRange(CONFIG.TODO_FIRST_DATA_ROW, 3, rowCount, 4)
    .clearDataValidations()
    .insertCheckboxes();
}
