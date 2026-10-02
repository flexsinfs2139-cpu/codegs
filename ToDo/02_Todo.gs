/**
 * TODO sheet: the task table (A:D) and four live quadrant columns (F:I).
 * The task table is the single source of truth; the quadrant columns only read from it.
 */

/** Resets formatting (keeping content) and re-applies the TODO layout. */
function rebuildTodoSheet_(sheet, seed) {
  migrateColumns_(sheet, COLUMNS.map(c => c.header)); // before the reset, while dates still show as dates
  resetSheet_(sheet, true);
  setupTodoSheet_(sheet, seed);
}

function setupTodoSheet_(sheet, seed) {
  const width = COLUMNS.length;
  const rows = Math.max(CONFIG.todoRows, lastContentRow_(sheet, width) + CONFIG.spareRows);
  const body = rows - 1;
  fitSheet_(sheet, rows, TODO_LAST_COL);
  sheet.setHiddenGridlines(true);

  sheet.getRange(1, 1, rows, TODO_LAST_COL)
    .setFontFamily(CONFIG.font)
    .setFontSize(10)
    .setVerticalAlignment('middle');

  // Sizes
  sheet.setRowHeight(1, CONFIG.headerHeight);
  sheet.setRowHeights(2, body, CONFIG.rowHeight);
  COLUMNS.forEach((c, i) => sheet.setColumnWidth(i + 1, c.width));
  sheet.setColumnWidth(MATRIX_COL - 1, CONFIG.gapWidth);
  sheet.setColumnWidths(MATRIX_COL, QUADRANTS.length, CONFIG.quadrantWidth);
  sheet.setFrozenRows(1);

  // Task table
  styleTable_(sheet.getRange(1, 1, rows, width), COLUMNS.map(c => c.header), THEME.header, THEME.headerText, THEME.border);
  sheet.getRange(2, COL.task, body, 1).setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);

  // Spacer column separator
  sheet.getRange(1, MATRIX_COL - 1, rows, 1)
    .setBackground(THEME.spacerBg)
    .setBorder(false, false, false, false, false, false);

  // Sample tasks (brand-new sheet only)
  if (seed && SAMPLE_TASKS.length) {
    const sample = SAMPLE_TASKS.map(t => COLUMNS.map(c => t[c.key]));
    sheet.getRange(2, 1, sample.length, width).setValues(sample);
  }

  // Capture existing data: insertCheckboxes() resets every cell to FALSE
  const data = sheet.getRange(2, 1, body, width).getValues();
  insertCheckboxes_(sheet.getRange(2, COL.important, body, 2));

  // Status dropdown; tasks without a status get the default
  const statusRange = sheet.getRange(2, COL.status, body, 1)
    .setDataValidation(statusValidation_())
    .setHorizontalAlignment('center');
  const statuses = data.map(r => [r[COL.status - 1]]);
  if (fillDefaultStatus_(data.map(r => [r[COL.task - 1]]), statuses)) statusRange.setValues(statuses);

  // Quadrant columns: distinct vibrant headers and gentle pastel column tints
  QUADRANTS.forEach((q, i) => {
    const colIndex = MATRIX_COL + i;
    // Header (Row 1)
    sheet.getRange(1, colIndex)
      .setValue(q.headerLabel || q.title)
      .setBackground(q.headerBg)
      .setFontColor(q.headerFg || '#ffffff')
      .setFontWeight('bold')
      .setHorizontalAlignment('center');

    // Body (Rows 2..rows)
    sheet.getRange(2, colIndex, body, 1)
      .setBackground(q.bg)
      .setFontColor(q.text)
      .setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);

    // Grid border
    sheet.getRange(1, colIndex, rows, 1)
      .setBorder(true, true, true, true, true, true, THEME.border, SpreadsheetApp.BorderStyle.SOLID);
  });

  sheet.getRange(2, MATRIX_COL, 1, QUADRANTS.length)
    .setFormulas([QUADRANTS.map(quadrantListFormula_)]);
  sheet.getRange(1, MATRIX_COL, rows, QUADRANTS.length).protect()
    .setDescription('Auto-generated. Edit tasks in the Tasks columns.')
    .setWarningOnly(true);

  sheet.setConditionalFormatRules(todoFormatRules_(sheet, body));
  sheet.setTabColor(THEME.todoTab);
}

/** Full-color rules: Done strike-through, Quadrant highlights on tasks, Status chips, Priority highlights, and Quadrant column tasks. */
function todoFormatRules_(sheet, body) {
  const doneRule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied(`=$${columnLetter_(COL.status)}2="Done"`)
    .setFontColor(THEME.faint)
    .setBackground(THEME.doneBg)
    .setStrikethrough(true)
    .setRanges([sheet.getRange(2, COL.task, body, 1)])
    .build();

  return [doneRule]
    .concat(taskQuadrantHighlightRules_(sheet, body))
    .concat(statusRules_([sheet.getRange(2, COL.status, body, 1)]))
    .concat(priorityRules_(sheet, body))
    .concat(quadrantColumnRules_(sheet, body));
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

function statusValidation_() {
  return SpreadsheetApp.newDataValidation()
    .requireValueInList(Object.keys(STATUSES), true)
    .setAllowInvalid(false)
    .build();
}
