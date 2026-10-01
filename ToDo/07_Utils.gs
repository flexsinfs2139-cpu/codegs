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
  [SpreadsheetApp.ProtectionType.SHEET, SpreadsheetApp.ProtectionType.RANGE].forEach(type =>
    sheet.getProtections(type).forEach(p => p.remove())
  );

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

/** Last row with real content in the first `width` columns; unchecked checkboxes (FALSE) don't count. */
function lastContentRow_(sheet, width) {
  const lastRow = sheet.getLastRow();
  if (lastRow === 0) return 1;
  const values = sheet.getRange(1, 1, lastRow, Math.min(width, sheet.getMaxColumns())).getValues();
  for (let r = values.length - 1; r >= 1; r--) {
    if (values[r].some(v => v !== '' && v !== false)) return r + 1;
  }
  return 1;
}

/**
 * Re-maps an older column layout onto `headers`, matching header names loosely
 * ("Task" fills "Tasks"). Values in columns the new layout doesn't have (e.g. Due,
 * Notes) are kept as a cell note on the task so nothing is lost. "✓" becomes TRUE.
 * Sheets that are already current, or have no task column, are left alone.
 */
function migrateColumns_(sheet, headers) {
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow === 0 || lastCol === 0) return;

  const key = h => String(h).trim().toLowerCase().replace(/s$/, '');
  const oldHeaders = sheet.getRange(1, 1, 1, lastCol).getDisplayValues()[0];
  const old = oldHeaders.map(key);
  const wanted = headers.map(key);
  if (wanted.every((h, i) => old[i] === h)) return;

  const source = wanted.map(h => old.indexOf(h));
  if (source[0] === -1) return;

  if (sheet.getMaxColumns() < headers.length) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), headers.length - sheet.getMaxColumns());
  }

  const all = sheet.getRange(1, 1, lastRow, lastCol);
  if (lastRow > 1) {
    const body = sheet.getRange(2, 1, lastRow - 1, lastCol);
    const values = body.getValues();
    const display = body.getDisplayValues();
    const oldNotes = body.getNotes();
    const extra = old.map((h, i) => i).filter(i => old[i] !== '' && source.indexOf(i) === -1);

    const rows = values.map(r => source.map(i => (i === -1 ? '' : r[i] === '✓' ? true : r[i])));
    const notes = display.map((r, n) => [
      [oldNotes[n][source[0]]]
        .concat(extra.filter(i => r[i] !== '').map(i => `${oldHeaders[i]}: ${r[i]}`))
        .filter(Boolean)
        .join('\n'),
    ]);

    all.clearContent().clearNote().clearDataValidations();
    sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
    sheet.getRange(2, 1, notes.length, 1).setNotes(notes);
  } else {
    all.clearContent();
  }
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
}

/** Black header row and a thin border on every cell of the table. */
function styleTable_(range, headers) {
  range.setBorder(true, true, true, true, true, true, THEME.border, SpreadsheetApp.BorderStyle.SOLID);
  range.getSheet().getRange(range.getRow(), range.getColumn(), 1, headers.length)
    .setValues([headers])
    .setBackground(THEME.header)
    .setFontColor(THEME.headerText)
    .setFontWeight('bold')
    .setHorizontalAlignment('center');
}

/** Turns a range into centered checkboxes, keeping any TRUE values already in it. */
function insertCheckboxes_(range) {
  const values = range.getValues().map(r => r.map(v => v === true));
  range.insertCheckboxes()
    .setValues(values)
    .setHorizontalAlignment('center');
}

/** 1-based column index to letter (A–Z). */
function columnLetter_(index) {
  return String.fromCharCode(64 + index);
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
