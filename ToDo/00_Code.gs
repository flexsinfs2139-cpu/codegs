/**
 * Eisenhower Matrix: modern SaaS-style workspace for Google Sheets
 * Run setupEisenhowerMatrix() once (or use the "Eisenhower" menu).
 * Re-running rebuilds the layout of every sheet but keeps all tasks on the
 * TODO sheet. Older 5-column TODO sheets are upgraded in place (Due added).
 *
 * Files:
 *   00_Code       entry points, menu and triggers
 *   01_Config     all configuration and constants
 *   02_Todo       TODO sheet (task list)
 *   03_Matrix     Eisenhower Matrix sheet
 *   04_Dashboard  Dashboard sheet
 *   05_Archive    archiving completed tasks
 *   06_Formulas   formula and conditional-format builders
 *   07_Utils      generic sheet helpers
 */

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Eisenhower')
    .addItem('Rebuild workspace', 'setupEisenhowerMatrix')
    .addItem('Archive completed tasks', 'archiveCompletedTasks')
    .addToUi();
}

/** Simple trigger. */
function onEdit(e) {
  if (!e || !e.range) return;
  if (e.range.getSheet().getName() === CONFIG.sheets.todo) handleTodoEdit_(e.range);
}

function setupEisenhowerMatrix() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const isNewTodo = !ss.getSheetByName(CONFIG.sheets.todo);

  const todo = getOrCreateSheet_(ss, CONFIG.sheets.todo);
  const matrix = getOrCreateSheet_(ss, CONFIG.sheets.matrix);
  const dashboard = getOrCreateSheet_(ss, CONFIG.sheets.dashboard);

  rebuildTodoSheet_(todo, isNewTodo && CONFIG.includeSampleTasks);

  resetSheet_(matrix);
  setupMatrixSheet_(matrix);

  resetSheet_(dashboard);
  setupDashboard_(dashboard);

  orderSheets_(ss, [todo, matrix, dashboard]);
  todo.setTabColor(THEME.ink);
  matrix.setTabColor(THEME.accent);
  dashboard.setTabColor('#22c55e');
  removeBlankDefaultSheet_(ss);

  SpreadsheetApp.flush();
  ss.setActiveSheet(matrix);
  ss.toast('Workspace is ready.', 'Eisenhower Matrix', 4);
}
