/**
 * Task details sidebar: shows the selected TODO row and saves edits back.
 * Sheets has no selection event that can reach an open sidebar, so the
 * sidebar polls getSelectedTask() while it is open.
 */

/** Menu action: opens the sidebar. */
function showTaskPanel() {
  const template = HtmlService.createTemplateFromFile('08_TaskPanelSidebar');
  template.bootstrap = JSON.stringify({
    statuses: STATUSES,
    quadrants: QUADRANTS,
    defaultStatus: CONFIG.defaultStatus,
    pollMs: CONFIG.panelPollMs,
  });
  SpreadsheetApp.getUi().showSidebar(template.evaluate().setTitle('Task details'));
}

/** Sidebar: the task on the user's current row, or the reason there is none. */
function getSelectedTask() {
  const sheet = SpreadsheetApp.getActiveSheet();
  if (sheet.getName() !== CONFIG.sheets.todo) return { row: null, reason: 'sheet' };
  const range = sheet.getActiveRange();
  const row = range ? range.getRow() : 1;
  if (row < 2) return { row: null, reason: 'header' };
  return readTask_(sheet, row);
}

/** Sidebar: writes only the changed fields and returns the fresh row. */
function saveTask(row, patch) {
  const sheet = todoSheetOrThrow_();
  row = Number(row);
  if (!(row >= 2 && row <= sheet.getMaxRows())) throw new Error('That row no longer exists.');
  const cell = key => sheet.getRange(row, COL[key]);

  // Plain-text format stops entries like "1/2 day" or "=x" being re-parsed
  if ('task' in patch) cell('task').setNumberFormat('@').setValue(String(patch.task).trim());
  if ('notes' in patch) cell('notes').setNumberFormat('@').setValue(String(patch.notes));
  if ('important' in patch) cell('important').setValue(patch.important === true);
  if ('urgent' in patch) cell('urgent').setValue(patch.urgent === true);
  if ('status' in patch) cell('status').setValue(STATUSES[patch.status] ? patch.status : '');
  if ('due' in patch) cell('due').setValue(/^\d{4}-\d{2}-\d{2}$/.test(patch.due) ? patch.due : '');

  if (cell('task').getValue() !== '' && cell('status').getValue() === '') {
    cell('status').setValue(CONFIG.defaultStatus);
  }
  return readTask_(sheet, row);
}

/** Sidebar: selects the first empty TODO row (adding rows if the list is full). */
function createTask() {
  const sheet = todoSheetOrThrow_();
  let row = firstEmptyRow_(sheet);
  if (!row) {
    rebuildTodoSheet_(sheet, false); // regrows the grid with CONFIG.spareRows empty rows
    row = firstEmptyRow_(sheet);
  }
  sheet.activate();
  sheet.getRange(row, COL.task).activate();
  return readTask_(sheet, row);
}

/** Sidebar: jumps to the TODO sheet. */
function openTodoSheet() {
  todoSheetOrThrow_().activate();
}

function readTask_(sheet, row) {
  const range = sheet.getRange(row, 1, 1, COLUMNS.length);
  const values = range.getValues()[0];
  const display = range.getDisplayValues()[0];
  const at = key => COL[key] - 1;
  const due = values[at('due')];
  return {
    row,
    task: display[at('task')],
    important: values[at('important')] === true,
    urgent: values[at('urgent')] === true,
    status: display[at('status')],
    due: due instanceof Date
      ? Utilities.formatDate(due, sheet.getParent().getSpreadsheetTimeZone(), 'yyyy-MM-dd')
      : '',
    notes: display[at('notes')],
  };
}

function firstEmptyRow_(sheet) {
  const values = sheet.getRange(2, 1, sheet.getMaxRows() - 1, COLUMNS.length).getValues();
  const index = values.findIndex(r => r.every(v => v === '' || v === false));
  return index === -1 ? null : index + 2;
}

function todoSheetOrThrow_() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.sheets.todo);
  if (!sheet) throw new Error('No TODO sheet yet. Run Eisenhower → Rebuild workspace first.');
  return sheet;
}
