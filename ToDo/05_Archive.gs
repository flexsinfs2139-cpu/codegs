/**
 * Archive: moves completed tasks off the TODO sheet into an Archive sheet.
 */

/** Menu action: moves every Done task from TODO to the Archive sheet (after confirmation). */
function archiveCompletedTasks() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const todo = ss.getSheetByName(CONFIG.sheets.todo);
  const lastRow = todo ? todo.getLastRow() : 0;
  if (lastRow < 2) {
    ss.toast('There are no tasks yet.', 'Archive', 4);
    return;
  }

  const width = COLUMNS.length;
  const values = todo.getRange(2, 1, lastRow - 1, width).getValues();
  const doneRows = [];
  values.forEach((r, i) => {
    if (r[COL.task - 1] !== '' && r[COL.status - 1] === 'Done') doneRows.push(i + 2);
  });
  if (!doneRows.length) {
    ss.toast('No completed tasks to archive.', 'Archive', 4);
    return;
  }

  const n = doneRows.length;
  const noun = n === 1 ? 'task' : 'tasks';
  const answer = ui.alert(
    'Archive completed tasks',
    `Move ${n} completed ${noun} to the "${CONFIG.sheets.archive}" sheet?`,
    ui.ButtonSet.OK_CANCEL
  );
  if (answer !== ui.Button.OK) return;

  appendToArchive_(ss, doneRows.map(r => values[r - 2]));

  // A sheet must keep at least one non-frozen row
  if (todo.getMaxRows() - 1 <= n) todo.insertRowsAfter(todo.getMaxRows(), 1);
  deleteRows_(todo, doneRows);
  rebuildTodoSheet_(todo, false); // restores the full grid and formatting

  SpreadsheetApp.flush();
  ss.setActiveSheet(todo);
  ss.toast(`Archived ${n} completed ${noun}.`, 'Eisenhower Matrix', 4);
}

/** Appends TODO rows to the Archive sheet with checkmarks and an archived timestamp. */
function appendToArchive_(ss, taskRows) {
  const archive = getOrCreateArchiveSheet_(ss);
  const archivedAt = new Date();
  const rows = taskRows.map(r => {
    const row = r.slice();
    row[COL.important - 1] = row[COL.important - 1] === true ? '✓' : '';
    row[COL.urgent - 1] = row[COL.urgent - 1] === true ? '✓' : '';
    return row.concat([archivedAt]);
  });

  const start = archive.getLastRow() + 1;
  const missing = start + rows.length - 1 - archive.getMaxRows();
  if (missing > 0) archive.insertRowsAfter(archive.getMaxRows(), missing);
  archive.getRange(start, 1, rows.length, rows[0].length).setValues(rows);
}

function getOrCreateArchiveSheet_(ss) {
  const existing = ss.getSheetByName(CONFIG.sheets.archive);
  if (existing) return existing;

  const sheet = ss.insertSheet(CONFIG.sheets.archive, ss.getSheets().length);
  const headers = COLUMNS.map(c => c.header).concat(['Archived']);
  const width = headers.length;
  const rows = sheet.getMaxRows();

  sheet.setHiddenGridlines(true);
  sheet.getRange(1, 1, rows, width)
    .setFontFamily(CONFIG.font)
    .setFontSize(10)
    .setFontColor(THEME.ink)
    .setVerticalAlignment('middle');
  sheet.getRange(1, 1, 1, width)
    .setValues([headers])
    .setFontWeight('bold')
    .setFontColor('#ffffff')
    .setBackground(THEME.ink);
  sheet.setFrozenRows(1);
  sheet.setRowHeight(1, 40);
  COLUMNS.forEach((c, i) => sheet.setColumnWidth(i + 1, c.width));
  sheet.setColumnWidth(width, 140);

  // Plain text stops task names like "1/2 day" being re-parsed as dates
  [COL.task, COL.notes].forEach(c => sheet.getRange(2, c, rows - 1, 1).setNumberFormat('@'));
  [COL.due, width].forEach(c => sheet.getRange(2, c, rows - 1, 1).setNumberFormat(CONFIG.dateFormat));
  sheet.getRange(1, COL.important, rows, 2).setHorizontalAlignment('center');

  sheet.setTabColor(THEME.faint);
  return sheet;
}
