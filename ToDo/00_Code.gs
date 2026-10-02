/**
 * Eisenhower Matrix: single-sheet task list for Google Sheets
 * Run setupEisenhowerMatrix() once (or use the "Eisenhower" menu).
 * Re-running rebuilds the layout but keeps every task. Sheets from older
 * versions are upgraded in place (Due / Notes move into a note on the task).
 *
 * Files:
 *   00_Code       entry points, menu and triggers
 *   01_Config     all configuration and constants
 *   02_Todo       TODO sheet (task table + quadrant columns)
 *   05_Archive    Archive sheet and archiving completed tasks
 *   06_Formulas   formula and conditional-format builders
 *   07_Utils      generic sheet helpers
 *   08_TaskPanel  task details sidebar (+ 08_TaskPanelSidebar.html)
 */

function onOpen() {
  let ui;
  try {
    ui = SpreadsheetApp.getUi();
  } catch (err) {
    // Run from the editor, a trigger, or a standalone project: there is no sheet UI to attach to
    console.warn('No spreadsheet UI here. Open the spreadsheet (Extensions → Apps Script must own this ' +
                 'project) and reload the tab to get the Eisenhower menu.');
    return;
  }

  ui.createMenu('Eisenhower')
    .addItem('Task details panel', 'showTaskPanel')
    .addSeparator()
    .addItem('Rebuild workspace', 'setupEisenhowerMatrix')
    .addItem('Archive completed tasks', 'archiveCompletedTasks')
    .addToUi();
}

/** Simple trigger. */
function onEdit(e) {
  if (!e || !e.range) return;
  if (e.range.getSheet().getName() === CONFIG.sheets.todo) handleTodoEdit_(e.range);
}

/**
 * Installable open trigger (added by Rebuild workspace): opens the Task details
 * panel whenever the spreadsheet loads in a browser. A simple onOpen can't show
 * a sidebar, and the mobile apps have no sidebars, so it only appears on the web.
 */
function autoOpenTaskPanel() {
  try {
    showTaskPanel();
  } catch (err) {
    console.warn('Task details panel not opened: ' + err.message);
  }
}

function setupEisenhowerMatrix() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const isNewTodo = !ss.getSheetByName(CONFIG.sheets.todo);

  const todo = getOrCreateSheet_(ss, CONFIG.sheets.todo);
  const archive = getOrCreateSheet_(ss, CONFIG.sheets.archive);

  rebuildTodoSheet_(todo, isNewTodo && CONFIG.includeSampleTasks);
  rebuildArchiveSheet_(archive);

  removeLegacySheets_(ss);
  orderSheets_(ss, [todo, archive]);
  todo.setTabColor(THEME.todoTab || THEME.header);
  removeBlankDefaultSheet_(ss);
  syncTaskPanelTrigger_(ss);

  SpreadsheetApp.flush();
  ss.setActiveSheet(todo);
  ss.toast('Workspace is ready.', 'Eisenhower Matrix', 4);
  if (CONFIG.autoOpenPanel) showTaskPanel();
}

/** Deletes the auto-generated Matrix / Dashboard sheets of older versions. */
function removeLegacySheets_(ss) {
  CONFIG.legacySheets.forEach(name => {
    const sheet = ss.getSheetByName(name);
    if (sheet && ss.getSheets().length > 1) ss.deleteSheet(sheet);
  });
}

/** Installs (or removes) the open trigger behind CONFIG.autoOpenPanel, never duplicating it. */
function syncTaskPanelTrigger_(ss) {
  const handler = 'autoOpenTaskPanel';
  const existing = ScriptApp.getUserTriggers(ss).filter(t => t.getHandlerFunction() === handler);
  if (CONFIG.autoOpenPanel && !existing.length) {
    ScriptApp.newTrigger(handler).forSpreadsheet(ss).onOpen().create();
  } else if (!CONFIG.autoOpenPanel) {
    existing.forEach(t => ScriptApp.deleteTrigger(t));
  }
}
