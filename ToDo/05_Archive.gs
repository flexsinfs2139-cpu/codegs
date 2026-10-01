/**
 * Archive: moves completed tasks off the TODO sheet into an Archive sheet
 * that uses the same table layout plus an Archived date.
 */

/** Menu action: moves every Done task from TODO to the Archive sheet (after confirmation). */
function archiveCompletedTasks() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const todo = ss.getSheetByName(CONFIG.sheets.todo);
  const lastRow = todo ? lastContentRow_(todo, COLUMNS.length) : 0;
  if (lastRow < 2) {
    ss.toast('There are no tasks yet.', 'Archive', 4);
    return;
  }

  const values = todo.getRange(2, 1, lastRow - 1, COLUMNS.length).getValues();
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

  deleteTaskRows_(todo, doneRows);
  rebuildTodoSheet_(todo, false); // restores the full grid and formatting

  SpreadsheetApp.flush();
  ss.setActiveSheet(todo);
  ss.toast(`Archived ${n} completed ${noun}.`, 'Eisenhower Matrix', 4);
}

/** Appends TODO rows to the Archive sheet with an archived timestamp. */
function appendToArchive_(ss, taskRows) {
  const archive = getOrCreateSheet_(ss, CONFIG.sheets.archive);
  rebuildArchiveSheet_(archive);

  const archivedAt = new Date();
  const rows = taskRows.map(r => r.concat([archivedAt]));
  const start = lastContentRow_(archive, rows[0].length) + 1;
  const missing = start + rows.length - 1 - archive.getMaxRows();
  if (missing > 0) archive.insertRowsAfter(archive.getMaxRows(), missing);
  archive.getRange(start, 1, rows.length, rows[0].length).setValues(rows);

  rebuildArchiveSheet_(archive); // formats the new rows and keeps spare rows below
}

/** Resets formatting (keeping content) and re-applies the Archive layout. */
function rebuildArchiveSheet_(sheet) {
  const headers = COLUMNS.map(c => c.header).concat(['Archived']);
  const width = headers.length;
  migrateColumns_(sheet, headers); // before the reset, while dates still show as dates
  resetSheet_(sheet, true);

  const rows = Math.max(CONFIG.todoRows, lastContentRow_(sheet, width) + CONFIG.spareRows);
  const body = rows - 1;
  fitSheet_(sheet, rows, width);
  sheet.setHiddenGridlines(true);

  sheet.getRange(1, 1, rows, width)
    .setFontFamily(CONFIG.font)
    .setFontSize(10)
    .setVerticalAlignment('middle');
  sheet.setRowHeight(1, CONFIG.headerHeight);
  sheet.setRowHeights(2, body, CONFIG.rowHeight);
  COLUMNS.forEach((c, i) => sheet.setColumnWidth(i + 1, c.width));
  sheet.setColumnWidth(width, CONFIG.archivedWidth);
  sheet.setFrozenRows(1);

  styleTable_(sheet.getRange(1, 1, rows, width), headers);
  sheet.getRange(2, COL.task, body, 1).setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
  sheet.getRange(2, COL.status, body, 1).setHorizontalAlignment('center');
  insertCheckboxes_(sheet.getRange(2, COL.important, body, 2));
  sheet.getRange(2, width, body, 1)
    .setNumberFormat(CONFIG.dateFormat)
    .setHorizontalAlignment('center');

  sheet.setConditionalFormatRules(statusRules_([sheet.getRange(2, COL.status, body, 1)]));
  sheet.setTabColor(THEME.archiveTab);
}

/**
 * Deletes the given 1-based rows (ascending) from the task table only, shifting
 * the tasks below up. Whole-row deletes would also cut the quadrant columns.
 */
function deleteTaskRows_(sheet, rows) {
  const width = COLUMNS.length;
  const last = lastContentRow_(sheet, width);
  const range = sheet.getRange(2, 1, last - 1, width);
  const drop = new Set(rows);
  const kept = range.getValues().filter((r, i) => !drop.has(i + 2));
  const notes = range.getNotes().filter((r, i) => !drop.has(i + 2));

  range.clearContent().clearNote();
  if (kept.length) {
    sheet.getRange(2, 1, kept.length, width).setValues(kept).setNotes(notes);
  }
}
