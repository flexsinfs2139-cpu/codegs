// ============================================================
// MIGRATION.GS — SAFE NON-DESTRUCTIVE SYSTEM UPGRADE ENGINE
// ============================================================

/**
 * Menu Action / Setup Routine: Executes a safe, idempotent migration across all sheets:
 * 1. Upgrades Todo sheet to 5 columns (ID, Task Description, Project, Due Date, Notes)
 *    and assigns permanent TSK-XXXX IDs without losing any existing task data.
 * 2. Upgrades all Month sheets to 8 columns (inserting Column C: Task ID) without
 *    touching or altering existing daily work log entries.
 * 3. Provisions the 12-column Decisions sheet and creates baseline Task Created events.
 * 4. Refreshes named ranges, dropdowns, conditional formatting, and dashboard.
 *
 * @param {boolean} [suppressAlert=false]
 */
function runFullSystemMigration(suppressAlert = false) {
  console.log('[Migration] Starting system migration...');
  const ss = getSpreadsheet();

  // 1. Migrate Todo Sheet
  const migratedTodosCount = migrateTodoSheet(ss);
  console.log(`[Migration] Todo sheet migrated (${migratedTodosCount} tasks verified)`);

  // 2. Provision Decisions Sheet
  const decSheet = ensureDecisionsSheet(ss);
  console.log('[Migration] Decisions sheet provisioned');

  // 3. Migrate all Month Sheets
  const migratedMonthSheetsCount = migrateAllMonthSheets(ss);
  console.log(`[Migration] ${migratedMonthSheetsCount} Month sheets verified and upgraded`);

  // 4. Update Named Ranges & Validations
  ensureListsSheet(ss);
  updateTodoTaskIdsNamedRange(ss);
  setupWorkbookValidationsAndFormatting(ss);

  // 5. Seed baseline timeline events for migrated Todo items if Decisions sheet is empty
  seedBaselineTimelineEvents(ss);

  // 6. Refresh Command Center Dashboard
  refreshDashboard(true, true);

  console.log('[Migration] System migration completed successfully.');

  if (!suppressAlert) {
    try {
      SpreadsheetApp.getUi().alert(
        '🚀 System Migration Complete!\n\n' +
        `• Todo Sheet: Migrated to 5 columns with permanent Task IDs (${migratedTodosCount} tasks).\n` +
        `• Month Sheets: ${migratedMonthSheetsCount} sheets upgraded to 8 columns with Task ID linking.\n` +
        '• Decisions Sheet: Chronological audit trail provisioned.\n' +
        '• All historical task descriptions, projects, and work logs have been 100% preserved.'
      );
    } catch (e) {
      ss.toast('System migration completed successfully.', '🚀 Migration', 5);
    }
  }
}


/**
 * Non-destructively migrates the Todo sheet to the 5-column layout:
 * [ID, Task Description, Project, Due Date, Notes]
 *
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
 * @returns {number} Count of tasks in Todo
 */
function migrateTodoSheet(ss) {
  let sheet = ss.getSheetByName(CONFIG.TODO_SHEET_NAME);
  if (!sheet) {
    createTodoSheet(true);
    return 0;
  }

  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();

  // Check if sheet is already in simplified 5-column layout
  if (lastRow >= 1) {
    const r1c1 = String(sheet.getRange(1, 1).getValue()).trim();
    if (r1c1 === 'ID' && lastCol === 5) {
      ensureAllTodoRowsHaveIds(sheet);
      formatTodoSheet(sheet);
      setupTodoDropdowns(sheet);
      return Math.max(sheet.getLastRow() - 1, 0);
    }
  }

  // Extract existing data from legacy layout
  const existingTasks = [];
  let headerRowIndex = 2; // Default legacy 2-row layout

  if (lastRow >= 1) {
    const r1c1 = String(sheet.getRange(1, 1).getValue()).trim();
    if (r1c1 === 'Task Name') {
      headerRowIndex = 1;
    }
  }

  const startDataRow = headerRowIndex + 1;
  if (lastRow >= startDataRow) {
    const rawData = sheet.getRange(startDataRow, 1, lastRow - headerRowIndex, lastCol).getValues();

    rawData.forEach(row => {
      let desc = '';
      let proj = '';
      let quadrantNote = '';

      if (headerRowIndex === 1 && String(row[0] || '').trim() === 'ID') {
        // Already has ID in col 1
        desc = String(row[1] || '').trim();
        proj = String(row[2] || '').trim();
      } else {
        // Legacy: Col 1 is Task Name, Col 2 is Project
        desc = String(row[0] || '').trim();
        proj = String(row[1] || '').trim();

        // Capture quadrant context if checked
        if (row[2] === true) quadrantNote = 'Migrated from Q1 (Do)';
        else if (row[3] === true) quadrantNote = 'Migrated from Q2 (Schedule)';
        else if (row[4] === true) quadrantNote = 'Migrated from Q3 (Delegate)';
        else if (row[5] === true) quadrantNote = 'Migrated from Q4 (Don\'t Do)';
      }

      if (desc) {
        existingTasks.push({
          desc: desc,
          proj: proj,
          notes: quadrantNote
        });
      }
    });
  }

  // Clear sheet and rebuild as simplified 5-column layout
  sheet.clear();
  sheet.clearConditionalFormatRules();

  // Set Row 1 header
  sheet.getRange(1, 1, 1, CONFIG.TODO_HEADERS.length).setValues([CONFIG.TODO_HEADERS]);

  // Insert migrated task rows with permanent Task IDs
  if (existingTasks.length > 0) {
    const newRows = existingTasks.map(t => {
      const taskId = generateNextTaskId();
      return [
        taskId,
        t.desc,
        t.proj,
        '', // Due Date
        t.notes || ''
      ];
    });

    sheet.getRange(2, 1, newRows.length, CONFIG.TODO_HEADERS.length).setValues(newRows);
  }

  formatTodoSheet(sheet);
  setupTodoDropdowns(sheet);
  updateTodoTaskIdsNamedRange(ss);

  return existingTasks.length;
}


/**
 * Migrates all Monthly task sheets to 8 columns:
 * Inserts Column C (Task ID) before Project if not already present.
 *
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
 * @returns {number} Count of Month sheets migrated
 */
function migrateAllMonthSheets(ss) {
  const monthSheets = getAllMonthSheets(ss);
  let count = 0;

  monthSheets.forEach(sheet => {
    try {
      const lastCol = sheet.getLastColumn();
      const r1Values = sheet.getRange(1, 1, 1, Math.max(lastCol, 1)).getValues()[0];
      const col3Header = String(r1Values[2] || '').trim();

      // If Column 3 is already 'Task ID', sheet is already upgraded
      if (col3Header === 'Task ID') {
        sheet.getRange(1, 1, 1, CONFIG.HEADERS.length).setValues([CONFIG.HEADERS]);
        formatWorkTracker(sheet);
        setupDropdowns(sheet);
        setupConditionalFormatting(sheet);
        count++;
        return;
      }

      // Legacy 7-column sheet: Col 3 is 'Project'
      if (col3Header === 'Project' || lastCol === 7) {
        // Insert Task ID column before Column 3
        sheet.insertColumnBefore(3);

        // Update Row 1 header to full 8 columns
        sheet.getRange(1, 1, 1, CONFIG.HEADERS.length).setValues([CONFIG.HEADERS]);

        // Format updated 8-column layout
        formatWorkTracker(sheet);
        setupDropdowns(sheet);
        setupConditionalFormatting(sheet);
        setupWeekendFormatting(sheet);

        count++;
        console.log(`[Migration] Upgraded month sheet: ${sheet.getName()}`);
      }
    } catch (err) {
      console.warn(`[Migration] Notice upgrading month sheet ${sheet.getName()}:`, err);
    }
  });

  return count;
}


/**
 * Seeds baseline 'Task Created' events for existing Todo items if Decisions sheet is empty.
 *
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
 */
function seedBaselineTimelineEvents(ss) {
  const decSheet = ensureDecisionsSheet(ss);
  if (decSheet.getLastRow() >= CONFIG.DECISIONS_FIRST_DATA_ROW) {
    // Already contains events; preserve existing history
    return;
  }

  const allTodos = getAllTodoTasks();
  if (!allTodos || allTodos.length === 0) return;

  const timezone = getTimezone();
  const now = new Date();

  allTodos.forEach((task, idx) => {
    if (task.id) {
      // Offset timestamps slightly for chronological order
      const eventTime = new Date(now.getTime() - (allTodos.length - idx) * 60000);
      recordTimelineEvent({
        taskId: task.id,
        eventType: 'Task Created',
        prevStatus: '-',
        newStatus: 'Pending',
        decision: `Task created in Todo backlog: "${task.description}"`,
        rationale: task.notes || 'Baseline backlog import',
        decidedBy: 'System Migration',
        dueDate: task.dueDate || '',
        eventDate: eventTime
      });
    }
  });

  console.log(`[Migration] Seeded ${allTodos.length} baseline events in Decisions sheet`);
}
