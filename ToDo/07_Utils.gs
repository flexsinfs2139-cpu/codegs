/**
 * Generic sheet helpers (no workspace-specific logic).
 */

function getOrCreateSheet_(ss, name) {
  return ss.getSheetByName(name) || ss.insertSheet(name);
}

/**
 * Resets a sheet: filters, banding, merges, validation, rules, protections, charts.
 * With keepContent, values and formulas survive and only formatting is cleared.
 */
function resetSheet_(sheet, keepContent) {
  const filter = sheet.getFilter();
  if (filter) filter.remove();

  sheet.getBandings().forEach(b => b.remove());
  sheet.getCharts().forEach(c => sheet.removeChart(c));
  sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET).forEach(p => p.remove());

  sheet.setFrozenRows(0);
  sheet.setFrozenColumns(0);

  const all = sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns());
  all.breakApart();
  all.clearDataValidations();
  if (keepContent) sheet.clearFormats();
  else sheet.clear();
  sheet.clearConditionalFormatRules();
  sheet.setHiddenGridlines(false);
}

/** Trims or extends the grid to exactly rows x cols (removes extra rows/columns). */
function fitSheet_(sheet, rows, cols) {
  const maxRows = sheet.getMaxRows();
  const maxCols = sheet.getMaxColumns();

  if (maxRows > rows) sheet.deleteRows(rows + 1, maxRows - rows);
  else if (maxRows < rows) sheet.insertRowsAfter(maxRows, rows - maxRows);

  if (maxCols > cols) sheet.deleteColumns(cols + 1, maxCols - cols);
  else if (maxCols < cols) sheet.insertColumnsAfter(maxCols, cols - maxCols);
}

/** Deletes the given 1-based rows (ascending), bottom-up in contiguous blocks. */
function deleteRows_(sheet, rows) {
  let end = rows.length - 1;
  while (end >= 0) {
    let start = end;
    while (start > 0 && rows[start - 1] === rows[start] - 1) start--;
    sheet.deleteRows(rows[start], end - start + 1);
    end = start - 1;
  }
}

/** 1-based column index to letter (A–Z). */
function columnLetter_(index) {
  return String.fromCharCode(64 + index);
}

/** Merges a range and writes a value into its top-left cell only. */
function mergeWithValue_(range, value) {
  range.merge();
  range.getCell(1, 1).setValue(value);
  return range;
}

function orderSheets_(ss, sheets) {
  sheets.forEach((sheet, i) => {
    ss.setActiveSheet(sheet);
    ss.moveActiveSheet(i + 1);
  });
}

/** Deletes the default "Sheet1" only if it is completely empty. */
function removeBlankDefaultSheet_(ss) {
  const sheet = ss.getSheetByName('Sheet1');
  if (sheet && sheet.getLastRow() === 0 && sheet.getLastColumn() === 0 && ss.getSheets().length > 1) {
    ss.deleteSheet(sheet);
  }
}
