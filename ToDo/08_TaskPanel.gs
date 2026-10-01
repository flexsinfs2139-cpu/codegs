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

/**
 * Sidebar: the task on the user's current row, or the reason there is none.
 * `current` is the task the sidebar shows ({ row, cell }), so a task picked from a
 * quadrant column stays selected even after an edit moves it to another quadrant.
 */
function getSelectedTask(current) {
  const sheet = SpreadsheetApp.getActiveSheet();
  if (sheet.getName() !== CONFIG.sheets.todo) return { row: null, reason: 'sheet' };
  const range = sheet.getActiveRange();
  let row = range ? range.getRow() : 1;
  if (row < 2) return { row: null, reason: 'header' };

  // A task picked in a quadrant column opens that task's own row
  if (range.getColumn() >= MATRIX_COL) {
    const cell = range.getCell(1, 1).getA1Notation();
    row = current && current.cell === cell
      ? Number(current.row)
      : findTaskRow_(sheet, range.getCell(1, 1).getDisplayValue());
    if (!row) return { row: null, reason: 'header' };
    return Object.assign(readTask_(sheet, row), { cell });
  }
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
  if ('important' in patch) cell('important').setValue(patch.important === true);
  if ('urgent' in patch) cell('urgent').setValue(patch.urgent === true);
  if ('status' in patch) cell('status').setValue(STATUSES[patch.status] ? patch.status : '');

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
  return {
    row,
    task: display[at('task')],
    important: values[at('important')] === true,
    urgent: values[at('urgent')] === true,
    status: display[at('status')],
  };
}

/** Row of the first open task with this name (quadrant columns list open tasks only). */
function findTaskRow_(sheet, name) {
  if (!name) return null;
  const values = sheet.getRange(2, 1, sheet.getMaxRows() - 1, COLUMNS.length).getDisplayValues();
  const index = values.findIndex(r => r[COL.task - 1] === name && r[COL.status - 1] !== 'Done');
  return index === -1 ? null : index + 2;
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
