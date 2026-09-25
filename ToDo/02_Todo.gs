/**
 * TODO sheet: the single source of truth for tasks.
 * The Matrix and Dashboard only read from it.
 */

/** Resets formatting (keeping content) and re-applies the TODO layout. */
function rebuildTodoSheet_(sheet, seed) {
  resetSheet_(sheet, true);
  setupTodoSheet_(sheet, seed);
}

function setupTodoSheet_(sheet, seed) {
  migrateTodoColumns_(sheet);

  const width = COLUMNS.length;
  const rows = Math.max(CONFIG.todoRows, sheet.getLastRow());
  const body = rows - 1;
  fitSheet_(sheet, rows, Math.max(width, sheet.getLastColumn()));
  sheet.setHiddenGridlines(true);

  sheet.getRange(1, 1, rows, width)
    .setFontFamily(CONFIG.font)
    .setFontSize(10)
    .setFontColor(THEME.ink)
    .setVerticalAlignment('middle');

  // Header
  sheet.getRange(1, 1, 1, width)
    .setValues([COLUMNS.map(c => c.header)])
    .setFontWeight('bold')
    .setFontColor('#ffffff')
    .setBackground(THEME.ink);
  ['important', 'urgent', 'status', 'due'].forEach(key =>
    sheet.getRange(1, COL[key]).setHorizontalAlignment('center')
  );
  sheet.setFrozenRows(1);

  // Sizes
  sheet.setRowHeight(1, 40);
  sheet.setRowHeights(2, body, 32);
  COLUMNS.forEach((c, i) => sheet.setColumnWidth(i + 1, c.width));

  // Sample tasks (brand-new sheet only)
  if (seed && SAMPLE_TASKS.length) {
    sheet.getRange(2, 1, SAMPLE_TASKS.length, width).setValues(sampleRows_(sheet));
  }

  // Capture existing data: insertCheckboxes() resets every cell to FALSE
  const data = sheet.getRange(2, 1, body, width).getValues();

  // Checkboxes for Important / Urgent, restored to their previous state
  sheet.getRange(2, COL.important, body, 2)
    .insertCheckboxes()
    .setValues(data.map(r => [r[COL.important - 1] === true, r[COL.urgent - 1] === true]))
    .setHorizontalAlignment('center');

  // Status dropdown; tasks without a status get the default
  const statusRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(Object.keys(STATUSES), true)
    .setAllowInvalid(false)
    .build();
  const statusRange = sheet.getRange(2, COL.status, body, 1)
    .setDataValidation(statusRule)
    .setHorizontalAlignment('center')
    .setFontSize(9)
    .setFontWeight('bold');
  const statuses = data.map(r => [r[COL.status - 1]]);
  if (fillDefaultStatus_(data.map(r => [r[COL.task - 1]]), statuses)) statusRange.setValues(statuses);

  // Due date (double-click opens the date picker)
  const dateRule = SpreadsheetApp.newDataValidation()
    .requireDate()
    .setAllowInvalid(false)
    .setHelpText('Enter a date, or double-click to pick one.')
    .build();
  sheet.getRange(2, COL.due, body, 1)
    .setDataValidation(dateRule)
    .setNumberFormat(CONFIG.dateFormat)
    .setHorizontalAlignment('center');

  // Long text stays on one clean line
  sheet.getRange(2, COL.task, body, 1).setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
  sheet.getRange(2, COL.notes, body, 1)
    .setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP)
    .setFontColor(THEME.muted);

  // Zebra rows
  sheet.getRange(2, 1, body, width)
    .applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY, false, false)
    .setFirstRowColor('#ffffff')
    .setSecondRowColor(THEME.surface);

  sheet.setConditionalFormatRules(todoFormatRules_(sheet, body));

  sheet.getRange(1, 1, rows, width).createFilter();
}

/** Done rows muted, overdue / due-today dates highlighted, status chips. */
function todoFormatRules_(sheet, body) {
  const ref = key => `$${columnLetter_(COL[key])}2`;
  const column = key => sheet.getRange(2, COL[key], body, 1);

  const doneRule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied(`=${ref('status')}="Done"`)
    .setFontColor(THEME.faint)
    .setStrikethrough(true)
    .setRanges([column('task'), column('due'), column('notes')])
    .build();

  const overdueRule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied(`=AND(${ref('due')}<>"",${ref('due')}<TODAY(),${ref('status')}<>"Done")`)
    .setFontColor(THEME.danger)
    .setBold(true)
    .setRanges([column('due')])
    .build();

  const dueTodayRule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied(`=AND(${ref('due')}=TODAY(),${ref('status')}<>"Done")`)
    .setFontColor(THEME.warning)
    .setBold(true)
    .setRanges([column('due')])
    .build();

  return [doneRule, overdueRule, dueTodayRule].concat(statusRules_([column('status')]));
}

/** onEdit handler: newly typed (or pasted) tasks get the default status. */
function handleTodoEdit_(range) {
  if (range.getColumn() > COL.task || range.getLastColumn() < COL.task) return;

  const sheet = range.getSheet();
  const first = Math.max(range.getRow(), 2);
  const count = range.getLastRow() - first + 1;
  if (count < 1) return;

  const tasks = sheet.getRange(first, COL.task, count, 1).getValues();
  const statusRange = sheet.getRange(first, COL.status, count, 1);
  const statuses = statusRange.getValues();
  if (fillDefaultStatus_(tasks, statuses)) statusRange.setValues(statuses);
}

/** Upgrades a TODO sheet from the original 5-column layout by inserting the Due column. */
function migrateTodoColumns_(sheet) {
  const lastCol = sheet.getLastColumn();
  if (lastCol === 0) return;
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  if (headers.indexOf('Due') === -1 && headers[COL.due - 1] === 'Notes') {
    sheet.insertColumnBefore(COL.due);
  }
}

/** Sets the default status on rows that have a task but no status. Returns true if any changed. */
function fillDefaultStatus_(tasks, statuses) {
  let changed = false;
  statuses.forEach((row, i) => {
    if (tasks[i][0] !== '' && row[0] === '') {
      row[0] = CONFIG.defaultStatus;
      changed = true;
    }
  });
  return changed;
}

/** SAMPLE_TASKS with due offsets turned into dates in the spreadsheet's time zone. */
function sampleRows_(sheet) {
  const tz = sheet.getParent().getSpreadsheetTimeZone();
  const now = Date.now();
  return SAMPLE_TASKS.map(([task, important, urgent, status, dueInDays, notes]) => {
    const due = dueInDays === ''
      ? ''
      : Utilities.formatDate(new Date(now + dueInDays * 86400000), tz, 'yyyy-MM-dd');
    return [task, important, urgent, status, due, notes];
  });
}
