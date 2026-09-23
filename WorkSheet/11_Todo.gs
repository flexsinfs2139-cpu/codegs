// ============================================================
// TODO.GS — TODO SHEET SETUP & MANAGEMENT (EISENHOWER MATRIX)
// ============================================================

/**
 * Menu Action: Creates the Todo sheet if missing, or refreshes its
 * Eisenhower Matrix formatting, checkboxes, and stats row in place.
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
  setupTodoCheckboxes(sheet);
  setupTodoConditionalFormatting(sheet);

  ss.setActiveSheet(sheet);

  if (!suppressAlert) {
    SpreadsheetApp.getUi().alert(
      isNew
        ? 'Todo sheet created successfully with Eisenhower Matrix.'
        : 'Todo sheet already exists — Eisenhower Matrix formatting, checkboxes, and stats row refreshed.'
    );
  }
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
    .setFontFamily(CONFIG.FONTS.TEXT)
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
    .setFontFamily(CONFIG.FONTS.TEXT)
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

  // 3. Data Rows (Row 3+)
  const lastRow = sheet.getLastRow();
  if (lastRow >= CONFIG.TODO_FIRST_DATA_ROW) {
    sheet
      .getRange(CONFIG.TODO_FIRST_DATA_ROW, 1, lastRow - CONFIG.TODO_HEADER_ROW, 2)
      .setFontFamily(CONFIG.FONTS.TEXT)
      .setFontSize(10);

    sheet
      .getRange(CONFIG.TODO_FIRST_DATA_ROW, 1, lastRow - CONFIG.TODO_HEADER_ROW, CONFIG.TODO_HEADERS.length)
      .setBorder(
        null, true, true, true, true, true,
        CONFIG.COLORS.BORDER,
        SpreadsheetApp.BorderStyle.SOLID
      );
  }

  // 4. Column Widths
  sheet.setColumnWidth(1, 340); // Task Name
  sheet.setColumnWidth(2, 140); // Project
  sheet.setColumnWidth(3, 120); // Q1: Do
  sheet.setColumnWidth(4, 120); // Q2: Schedule
  sheet.setColumnWidth(5, 120); // Q3: Delegate
  sheet.setColumnWidth(6, 120); // Q4: Don't Do

  // 4. Row Heights
  sheet.setRowHeight(1, 32);
  sheet.setRowHeight(2, 28);

  // 5. Freeze top 2 rows (Stats row + Table headers) & hide gridlines
  sheet.setFrozenRows(2);
  sheet.setHideGridlines(true);

  // 6. Enforce strict single-selection across all quadrants
  sanitizeAllTodoQuadrants(sheet);
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
 * starting from row 3. Sanitizes all existing rows so only one quadrant is selected.
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

  sanitizeAllTodoQuadrants(sheet);
}


/**
 * Scans every data row in the Todo sheet and enforces strict single selection
 * across the 4 Eisenhower quadrant columns (Cols 3–6). If any row has multiple
 * checkboxes checked, preserves only the first checked quadrant and unchecks the rest.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} [sheet]
 */
function sanitizeAllTodoQuadrants(sheet) {
  if (!sheet) {
    const ss = getSpreadsheet();
    sheet = ss.getSheetByName(CONFIG.TODO_SHEET_NAME);
  }
  if (!sheet) return;

  const lastRow = sheet.getLastRow();
  if (lastRow < CONFIG.TODO_FIRST_DATA_ROW) return;

  const numRows = lastRow - CONFIG.TODO_HEADER_ROW;
  const range = sheet.getRange(CONFIG.TODO_FIRST_DATA_ROW, 3, numRows, 4);
  const values = range.getValues();
  let modified = false;

  for (let r = 0; r < values.length; r++) {
    let firstFound = false;
    for (let c = 0; c < 4; c++) {
      const v = values[r][c];
      const isChecked = v === true || String(v).toUpperCase() === 'TRUE' || v === 1;

      if (isChecked) {
        if (!firstFound) {
          firstFound = true;
          if (values[r][c] !== true) {
            values[r][c] = true;
            modified = true;
          }
        } else {
          values[r][c] = false;
          modified = true;
        }
      } else {
        if (values[r][c] !== false) {
          values[r][c] = false;
          modified = true;
        }
      }
    }
  }

  if (modified) {
    range.setValues(values);
  }
}


/**
 * Handles mutual exclusivity for Eisenhower quadrant checkboxes.
 * When any quadrant checkbox (Q1–Q4) is checked, all other quadrants
 * on that same task row are automatically unchecked in an atomic write.
 *
 * @param {GoogleAppsScript.Events.SheetsOnEdit} [e]
 */
function handleTodoQuadrantExclusiveSelect(e) {
  // If invoked without event object (e.g. manual execution or testing)
  if (!e || !e.range) {
    sanitizeAllTodoQuadrants();
    return;
  }

  const sheet = e.range.getSheet();
  const sheetName = sheet.getName().trim().toLowerCase();
  if (sheetName !== CONFIG.TODO_SHEET_NAME.toLowerCase()) {
    return;
  }

  const startRow = e.range.getRow();
  const startCol = e.range.getColumn();
  const numRows = e.range.getNumRows();
  const numCols = e.range.getNumColumns();
  const endRow = startRow + numRows - 1;
  const endCol = startCol + numCols - 1;

  // Only listen to data rows (row >= 3) and quadrant columns (Cols 3 to 6: C, D, E, F)
  if (endRow < CONFIG.TODO_FIRST_DATA_ROW || endCol < 3 || startCol > 6) {
    return;
  }

  const firstProcessRow = Math.max(startRow, CONFIG.TODO_FIRST_DATA_ROW);
  const totalProcessRows = endRow - firstProcessRow + 1;

  for (let r = 0; r < totalProcessRows; r++) {
    const currentRow = firstProcessRow + r;
    const quadrantRange = sheet.getRange(currentRow, 3, 1, 4);
    const rowValues = quadrantRange.getValues()[0];

    // Single-cell click interaction (typical user clicking a checkbox)
    const isSingleCellEdit = (numRows === 1 && numCols === 1 && startCol >= 3 && startCol <= 6);
    const cellVal = e.range.getValue();
    const isEditedChecked = cellVal === true || String(cellVal).toUpperCase() === 'TRUE' || cellVal === 1 || e.value === 'TRUE';

    if (isSingleCellEdit) {
      const targetColIndex = startCol - 3; // 0 for C, 1 for D, 2 for E, 3 for F

      if (isEditedChecked) {
        // Enforce radio button: Only the clicked quadrant is true, all other 3 are false
        const exclusiveRow = [false, false, false, false];
        exclusiveRow[targetColIndex] = true;
        quadrantRange.setValues([exclusiveRow]);
      } else {
        // User unchecked this quadrant
        quadrantRange.getCell(1, targetColIndex + 1).setValue(false);
      }
    } else {
      // Multi-cell edit, paste, or range spacebar toggle
      const checkedCols = [];
      for (let c = 0; c < 4; c++) {
        const v = rowValues[c];
        if (v === true || String(v).toUpperCase() === 'TRUE' || v === 1) {
          checkedCols.push(c + 3);
        }
      }

      if (checkedCols.length > 1) {
        let activeCol = (startCol >= 3 && startCol <= 6) ? startCol : checkedCols[checkedCols.length - 1];
        if (!checkedCols.includes(activeCol)) {
          activeCol = checkedCols[checkedCols.length - 1];
        }

        const exclusiveRow = [
          activeCol === 3,
          activeCol === 4,
          activeCol === 5,
          activeCol === 6
        ];
        quadrantRange.setValues([exclusiveRow]);
      }
    }
  }
}
