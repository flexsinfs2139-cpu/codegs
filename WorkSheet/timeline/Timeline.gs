// ============================================================
// TIMELINE.GS — DECISIONS & CHRONOLOGICAL TASK TIMELINE ENGINE
// ============================================================

/**
 * Ensures the 'Decisions' sheet exists with proper structure and formatting.
 *
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} [ss]
 * @returns {GoogleAppsScript.Spreadsheet.Sheet}
 */
function ensureDecisionsSheet(ss) {
  if (!ss) ss = getSpreadsheet();
  let sheet = ss.getSheetByName(CONFIG.DECISIONS_SHEET_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.DECISIONS_SHEET_NAME);
  }

  formatDecisionsSheet(sheet);
  return sheet;
}


/**
 * Formats the Decisions sheet with the 12-column chronological timeline schema.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 */
function formatDecisionsSheet(sheet) {
  const ss = sheet.getParent();
  const headers = CONFIG.DECISIONS_HEADERS;

  // Set Header row
  sheet.getRange(CONFIG.DECISIONS_HEADER_ROW, 1, 1, headers.length).setValues([headers]);

  // Style Header
  sheet
    .getRange(CONFIG.DECISIONS_HEADER_ROW, 1, 1, headers.length)
    .setBackground('#0f172a')
    .setFontColor('#ffffff')
    .setFontFamily(CONFIG.FONTS.TEXT)
    .setFontWeight('bold')
    .setFontSize(10)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');

  sheet.setRowHeight(CONFIG.DECISIONS_HEADER_ROW, 28);

  // Column Dimensions
  const widths = (CONFIG.DIMENSIONS && CONFIG.DIMENSIONS.DECISIONS_COL_WIDTHS) || [
    110, 100, 140, 150, 120, 120, 280, 280, 130, 220, 100, 180
  ];
  widths.forEach((width, index) => {
    sheet.setColumnWidth(index + 1, width);
  });

  // Freeze top row & hide gridlines
  sheet.setFrozenRows(CONFIG.DECISIONS_HEADER_ROW);
  setSheetGridlinesHidden(sheet, true);

  // Set Event Type validation on Column D (Col 4)
  const eventTypesRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(CONFIG.EVENT_TYPES, true)
    .setAllowInvalid(false)
    .build();

  const maxRows = Math.max(sheet.getMaxRows() - 1, 100);
  sheet.getRange(CONFIG.DECISIONS_FIRST_DATA_ROW, 4, maxRows, 1).setDataValidation(eventTypesRule);
}


/**
 * Generates a permanent, unique, strictly monotonic Task ID (e.g. TSK-0001).
 * Never reuses existing IDs, including IDs of deleted tasks.
 *
 * @returns {string} Permanent Task ID
 */
function generateNextTaskId() {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
  } catch (err) {
    console.warn('Could not obtain lock for Task ID generation:', err);
  }

  try {
    const props = PropertiesService.getScriptProperties();
    let currentCounter = parseInt(props.getProperty('TASK_ID_COUNTER') || '0', 10);

    // Initial counter discovery if not set
    if (currentCounter === 0) {
      currentCounter = discoverHighestTaskId();
    }

    currentCounter++;
    props.setProperty('TASK_ID_COUNTER', String(currentCounter));

    return 'TSK-' + String(currentCounter).padStart(4, '0');
  } finally {
    try {
      lock.releaseLock();
    } catch (e) {}
  }
}


/**
 * Discovers the highest existing Task ID across all sheets to prevent collisions.
 *
 * @returns {number}
 */
function discoverHighestTaskId() {
  let highest = 0;
  try {
    const ss = getSpreadsheet();

    // Check Todo sheet
    const todoSheet = ss.getSheetByName(CONFIG.TODO_SHEET_NAME);
    if (todoSheet && todoSheet.getLastRow() >= CONFIG.TODO_FIRST_DATA_ROW) {
      const numRows = todoSheet.getLastRow() - CONFIG.TODO_HEADER_ROW;
      const ids = todoSheet.getRange(CONFIG.TODO_FIRST_DATA_ROW, 1, numRows, 1).getValues();
      ids.forEach(r => {
        const match = String(r[0] || '').match(/^TSK-(\d+)$/i);
        if (match) {
          const num = parseInt(match[1], 10);
          if (num > highest) highest = num;
        }
      });
    }

    // Check Decisions sheet
    const decSheet = ss.getSheetByName(CONFIG.DECISIONS_SHEET_NAME);
    if (decSheet && decSheet.getLastRow() >= CONFIG.DECISIONS_FIRST_DATA_ROW) {
      const numRows = decSheet.getLastRow() - CONFIG.DECISIONS_HEADER_ROW;
      const ids = decSheet.getRange(CONFIG.DECISIONS_FIRST_DATA_ROW, 2, numRows, 1).getValues();
      ids.forEach(r => {
        const match = String(r[0] || '').match(/^TSK-(\d+)$/i);
        if (match) {
          const num = parseInt(match[1], 10);
          if (num > highest) highest = num;
        }
      });
    }
  } catch (err) {
    console.warn('Error discovering highest Task ID:', err);
  }
  return highest;
}


/**
 * Generates a permanent, unique, strictly monotonic Timeline Event ID (e.g. EVT-0001).
 *
 * @returns {string} Permanent Event ID
 */
function generateNextTimelineId() {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
  } catch (err) {
    console.warn('Could not obtain lock for Timeline ID generation:', err);
  }

  try {
    const props = PropertiesService.getScriptProperties();
    let currentCounter = parseInt(props.getProperty('TIMELINE_ID_COUNTER') || '0', 10);

    if (currentCounter === 0) {
      currentCounter = discoverHighestTimelineId();
    }

    currentCounter++;
    props.setProperty('TIMELINE_ID_COUNTER', String(currentCounter));

    return 'EVT-' + String(currentCounter).padStart(4, '0');
  } finally {
    try {
      lock.releaseLock();
    } catch (e) {}
  }
}


/**
 * Discovers the highest existing Timeline ID in the Decisions sheet.
 *
 * @returns {number}
 */
function discoverHighestTimelineId() {
  let highest = 0;
  try {
    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG.DECISIONS_SHEET_NAME);
    if (sheet && sheet.getLastRow() >= CONFIG.DECISIONS_FIRST_DATA_ROW) {
      const numRows = sheet.getLastRow() - CONFIG.DECISIONS_HEADER_ROW;
      const ids = sheet.getRange(CONFIG.DECISIONS_FIRST_DATA_ROW, 1, numRows, 1).getValues();
      ids.forEach(r => {
        const match = String(r[0] || '').match(/^EVT-(\d+)$/i);
        if (match) {
          const num = parseInt(match[1], 10);
          if (num > highest) highest = num;
        }
      });
    }
  } catch (err) {
    console.warn('Error discovering highest Timeline ID:', err);
  }
  return highest;
}


/**
 * Records an immutable event row in the Decisions sheet.
 * Handles deduplication, ID generation, timestamping, and formatting.
 *
 * @param {Object} p
 * @param {string} p.taskId
 * @param {string} p.eventType
 * @param {string} [p.prevStatus='-']
 * @param {string} [p.newStatus='-']
 * @param {string} [p.decision='']
 * @param {string} [p.rationale='']
 * @param {string} [p.decidedBy='']
 * @param {string} [p.nextAction='']
 * @param {string} [p.dueDate='']
 * @param {string} [p.reference='']
 * @param {Date} [p.eventDate]
 * @returns {string} The assigned Timeline ID
 */
function recordTimelineEvent(p) {
  if (!p || !p.eventType) {
    throw new Error('Event type is required to record a timeline event.');
  }

  const ss = getSpreadsheet();
  const sheet = ensureDecisionsSheet(ss);
  const timezone = getTimezone();

  const taskId = (p.taskId && String(p.taskId).trim()) || 'GENERAL';
  const eventDateObj = p.eventDate || new Date();
  const formattedDate = Utilities.formatDate(eventDateObj, timezone, 'yyyy-MM-dd HH:mm');

  // Deduplication check: Avoid recording exact duplicate status change within 10 seconds
  const lastRow = sheet.getLastRow();
  if (lastRow >= CONFIG.DECISIONS_FIRST_DATA_ROW) {
    const lastRowValues = sheet.getRange(lastRow, 1, 1, CONFIG.DECISIONS_HEADERS.length).getValues()[0];
    const lastTaskId = String(lastRowValues[1] || '').trim();
    const lastDate = String(lastRowValues[2] || '').trim();
    const lastEventType = String(lastRowValues[3] || '').trim();
    const lastNewStatus = String(lastRowValues[5] || '').trim();

    if (
      lastTaskId === taskId &&
      lastEventType === p.eventType &&
      lastNewStatus === (p.newStatus || '-') &&
      lastDate === formattedDate
    ) {
      console.log(`[Timeline] Skipping duplicate event for ${taskId}: ${p.eventType}`);
      return String(lastRowValues[0]);
    }
  }

  const timelineId = generateNextTimelineId();
  const newRow = lastRow + 1;

  const rowData = [
    timelineId,
    taskId,
    formattedDate,
    p.eventType,
    p.prevStatus || '-',
    p.newStatus || '-',
    p.decision || '',
    p.rationale || '',
    p.decidedBy || '',
    p.nextAction || '',
    p.dueDate || '',
    p.reference || ''
  ];

  sheet.getRange(newRow, 1, 1, rowData.length).setValues([rowData]);
  sheet.setRowHeight(newRow, (CONFIG.DIMENSIONS && CONFIG.DIMENSIONS.DECISIONS_ROW_HEIGHT) || 36);

  // Apply typography
  const textFont = CONFIG.FONTS.TEXT;
  const digitsFont = CONFIG.FONTS.DIGITS;

  sheet
    .getRange(newRow, 1, 1, rowData.length)
    .setFontFamily(textFont)
    .setFontSize(10)
    .setVerticalAlignment('middle')
    .setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);

  // Monospace columns: Timeline ID (1), Task ID (2), Event Date (3), Due Date (11)
  sheet.getRange(newRow, 1).setFontFamily(digitsFont).setHorizontalAlignment('center').setFontWeight('bold');
  sheet.getRange(newRow, 2).setFontFamily(digitsFont).setHorizontalAlignment('center').setFontWeight('bold');
  sheet.getRange(newRow, 3).setFontFamily(digitsFont).setHorizontalAlignment('center');
  sheet.getRange(newRow, 4).setHorizontalAlignment('center');
  sheet.getRange(newRow, 5, 1, 2).setHorizontalAlignment('center');
  sheet.getRange(newRow, 11).setFontFamily(digitsFont).setHorizontalAlignment('center');

  console.log(`[Timeline] Recorded ${timelineId} for ${taskId}: ${p.eventType}`);
  return timelineId;
}


/**
 * Returns all timeline events associated with a specific Task ID.
 *
 * @param {string} taskId
 * @returns {Array<Object>}
 */
function getTaskTimelineHistory(taskId) {
  if (!taskId) return [];
  const cleanId = String(taskId).trim().toUpperCase();

  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.DECISIONS_SHEET_NAME);
  if (!sheet || sheet.getLastRow() < CONFIG.DECISIONS_FIRST_DATA_ROW) {
    return [];
  }

  const numRows = sheet.getLastRow() - CONFIG.DECISIONS_HEADER_ROW;
  const data = sheet.getRange(CONFIG.DECISIONS_FIRST_DATA_ROW, 1, numRows, CONFIG.DECISIONS_HEADERS.length).getValues();

  const events = [];
  data.forEach(row => {
    const rowTaskId = String(row[1] || '').trim().toUpperCase();
    if (rowTaskId === cleanId) {
      events.push({
        timelineId: String(row[0] || ''),
        taskId: rowTaskId,
        eventDate: String(row[2] || ''),
        eventType: String(row[3] || ''),
        prevStatus: String(row[4] || '-'),
        newStatus: String(row[5] || '-'),
        decision: String(row[6] || ''),
        rationale: String(row[7] || ''),
        decidedBy: String(row[8] || ''),
        nextAction: String(row[9] || ''),
        dueDate: String(row[10] || ''),
        reference: String(row[11] || '')
      });
    }
  });

  return events;
}


/**
 * Fetches all available Todo tasks with their current details.
 *
 * @returns {Array<Object>}
 */
function getAllTodoTasks() {
  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.TODO_SHEET_NAME);
  if (!sheet || sheet.getLastRow() < CONFIG.TODO_FIRST_DATA_ROW) {
    return [];
  }

  const numRows = sheet.getLastRow() - CONFIG.TODO_HEADER_ROW;
  const data = sheet.getRange(CONFIG.TODO_FIRST_DATA_ROW, 1, numRows, CONFIG.TODO_HEADERS.length).getValues();

  const tasks = [];
  data.forEach(r => {
    const id = String(r[0] || '').trim();
    const desc = String(r[1] || '').trim();
    const proj = String(r[2] || '').trim();
    const dueDate = r[3] instanceof Date
      ? Utilities.formatDate(r[3], getTimezone(), 'yyyy-MM-dd')
      : String(r[3] || '').trim();
    const notes = String(r[4] || '').trim();

    if (id || desc) {
      tasks.push({
        id: id,
        description: desc,
        project: proj,
        dueDate: dueDate,
        notes: notes,
        label: id ? `${id} — ${desc} (${proj || 'No Project'})` : desc
      });
    }
  });

  return tasks;
}


/**
 * Looks up a single Todo item by Task ID.
 *
 * @param {string} taskId
 * @returns {Object|null}
 */
function getTodoTaskById(taskId) {
  if (!taskId) return null;
  const cleanId = String(taskId).trim().toUpperCase();
  const allTasks = getAllTodoTasks();
  return allTasks.find(t => t.id.toUpperCase() === cleanId) || null;
}


// ============================================================
// REAL-TIME SPREADSHEET EDIT DISPATCHER
// ============================================================

/**
 * Handles all real-time edit operations:
 * 1. Auto-assigns Task ID when a new task is typed in the Todo sheet.
 * 2. Auto-populates Task Description and Project when a Task ID is selected in a Month sheet.
 * 3. Records timeline events when Status or Due Date changes.
 *
 * @param {GoogleAppsScript.Events.SheetsOnEdit} e
 */
function handleSpreadsheetEdit(e) {
  if (!e || !e.range) return;

  const sheet = e.range.getSheet();
  const sheetName = sheet.getName().trim();
  const row = e.range.getRow();
  const col = e.range.getColumn();

  // 1. EDITS ON THE TODO SHEET
  if (sheetName.toLowerCase() === CONFIG.TODO_SHEET_NAME.toLowerCase()) {
    handleTodoSheetEdit(sheet, row, col, e);
    return;
  }

  // 2. EDITS ON MONTH SHEETS (e.g. NOV26, OCT26)
  const isMonthSheet = /^[A-Z]{3}\d{2}$/i.test(sheetName);
  if (isMonthSheet && row >= 2) {
    handleMonthSheetEdit(sheet, row, col, e);
  }
}


/**
 * Handles cell edits on the Todo sheet.
 */
function handleTodoSheetEdit(sheet, row, col, e) {
  if (row < CONFIG.TODO_FIRST_DATA_ROW) return;

  // Col 2: Task Description edited
  if (col === 2) {
    const idCell = sheet.getRange(row, 1);
    const existingId = String(idCell.getValue() || '').trim();
    const taskDesc = String(e.range.getValue() || '').trim();

    // If description entered and ID is missing, auto-assign a permanent Task ID
    if (taskDesc && !existingId) {
      const newId = generateNextTaskId();
      idCell.setValue(newId);
      idCell.setFontFamily(CONFIG.FONTS.DIGITS).setFontWeight('bold').setHorizontalAlignment('center');

      const project = String(sheet.getRange(row, 3).getValue() || '').trim();
      recordTimelineEvent({
        taskId: newId,
        eventType: 'Task Created',
        prevStatus: '-',
        newStatus: 'Pending',
        decision: `Task created in Todo backlog: "${taskDesc}"`,
        rationale: 'Created via direct spreadsheet entry',
        decidedBy: 'User'
      });

      // Update named range so Month dropdown reflects the new ID
      updateTodoTaskIdsNamedRange(sheet.getParent());
    }
  }

  // Col 4: Due Date edited
  if (col === 4) {
    const taskId = String(sheet.getRange(row, 1).getValue() || '').trim();
    const newDueDate = e.range.getValue();
    const formattedDue = newDueDate instanceof Date
      ? Utilities.formatDate(newDueDate, getTimezone(), 'yyyy-MM-dd')
      : String(newDueDate || '').trim();

    if (taskId && formattedDue) {
      recordTimelineEvent({
        taskId: taskId,
        eventType: 'Due Date Changed',
        decision: `Target due date updated to ${formattedDue}`,
        dueDate: formattedDue,
        decidedBy: 'User'
      });
    }
  }
}


/**
 * Handles cell edits on Monthly Work Sheets.
 */
function handleMonthSheetEdit(sheet, row, col, e) {
  const taskIdCol = CONFIG.HEADERS.indexOf('Task ID') + 1; // Col 3
  const projectCol = CONFIG.HEADERS.indexOf('Project') + 1; // Col 4
  const taskDescCol = CONFIG.HEADERS.indexOf('Task') + 1;    // Col 5
  const statusCol = CONFIG.HEADERS.indexOf('Status') + 1;    // Col 8

  // A. TASK ID DROPDOWN SELECTION (Column C)
  if (col === taskIdCol) {
    const rawVal = String(e.range.getValue() || '').trim();
    if (!rawVal) return;

    // Support clean ID (TSK-0001) or combined label (TSK-0001 — Fix Login)
    const match = rawVal.match(/^(TSK-\d+)/i);
    const cleanId = match ? match[1].toUpperCase() : rawVal.toUpperCase();

    // Re-format cell to pure Task ID if label was pasted
    if (rawVal !== cleanId) {
      e.range.setValue(cleanId);
    }

    const todoItem = getTodoTaskById(cleanId);
    if (todoItem) {
      // Auto-populate Task Description and Project
      if (todoItem.description) {
        sheet.getRange(row, taskDescCol).setValue(todoItem.description);
      }
      if (todoItem.project) {
        sheet.getRange(row, projectCol).setValue(todoItem.project);
      }
      console.log(`[Month] Auto-populated ${cleanId} on row ${row}: ${todoItem.description}`);
    }
    return;
  }

  // B. STATUS DROPDOWN SELECTION (Column H)
  if (col === statusCol) {
    const oldStatus = String(e.oldValue || '-').trim();
    const newStatus = String(e.range.getValue() || '').trim();

    if (!newStatus || oldStatus === newStatus) return;

    const taskId = String(sheet.getRange(row, taskIdCol).getValue() || '').trim();
    const taskDesc = String(sheet.getRange(row, taskDescCol).getValue() || '').trim();
    const project = String(sheet.getRange(row, projectCol).getValue() || '').trim();

    let eventType = 'Status Changed';
    if (newStatus === 'Blocked') {
      eventType = 'Blocker Identified';
    } else if (oldStatus === 'Blocked' && newStatus !== 'Blocked') {
      eventType = 'Blocker Resolved';
    } else if (newStatus === 'Completed') {
      eventType = 'Task Completed';
    } else if (newStatus === 'Cancelled') {
      eventType = 'Task Cancelled';
    } else if (oldStatus === 'Completed' && newStatus !== 'Completed') {
      eventType = 'Task Reopened';
    }

    recordTimelineEvent({
      taskId: taskId || 'GENERAL',
      eventType: eventType,
      prevStatus: oldStatus,
      newStatus: newStatus,
      decision: taskId
        ? `Status changed from ${oldStatus} to ${newStatus} on ${taskDesc}`
        : `[${project}] ${taskDesc} status: ${oldStatus} ➔ ${newStatus}`,
      decidedBy: 'User'
    });
  }
}


// ============================================================
// DIALOG CONTROLLERS (RECORD DECISION & VIEW TASK TIMELINE)
// ============================================================

/**
 * Menu Action: Opens the Record Decision / Event modal dialog.
 */
function openRecordDecisionDialog() {
  const allTasks = getAllTodoTasks();
  let htmlOutput;
  const candidateNames = ['timeline/DecisionDialog', 'DecisionDialog', '14_DecisionDialog'];
  for (const name of candidateNames) {
    try {
      const template = HtmlService.createTemplateFromFile(name);
      template.tasks = allTasks;
      template.eventTypes = CONFIG.EVENT_TYPES;
      htmlOutput = template.evaluate();
      if (htmlOutput) break;
    } catch (ignore) {
      // Continue next candidate
    }
  }

  if (!htmlOutput) {
    const fallbackHtml = getFallbackDecisionDialogHtml(allTasks);
    htmlOutput = HtmlService.createHtmlOutput(fallbackHtml);
  }

  htmlOutput.setWidth(680).setHeight(560);
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Record Decision / Timeline Event');
}


/**
 * Client-callable function: Saves a decision or event submitted via modal.
 *
 * @param {Object} formData
 * @returns {Object} Result payload
 */
function submitDecisionEvent(formData) {
  try {
    if (!formData.eventType) {
      return { success: false, message: 'Event type is required.' };
    }

    const timelineId = recordTimelineEvent({
      taskId: formData.taskId,
      eventType: formData.eventType,
      prevStatus: formData.prevStatus || '-',
      newStatus: formData.newStatus || '-',
      decision: formData.decision,
      rationale: formData.rationale,
      decidedBy: formData.decidedBy,
      nextAction: formData.nextAction,
      dueDate: formData.dueDate,
      reference: formData.reference
    });

    return {
      success: true,
      timelineId: timelineId,
      message: `Event ${timelineId} successfully recorded in Decisions timeline.`
    };
  } catch (err) {
    return { success: false, message: err.message || String(err) };
  }
}


/**
 * Menu Action: Opens the interactive Task History / Timeline viewer modal.
 * Automatically inspects the active cell to pre-select a Task ID if available.
 */
function openTaskHistoryDialog(preSelectedTaskId) {
  const ss = getSpreadsheet();
  const sheet = ss.getActiveSheet();

  let targetId = preSelectedTaskId;
  if (!targetId) {
    const activeCell = sheet.getActiveCell();
    const cellVal = String(activeCell.getValue() || '').trim();
    const match = cellVal.match(/^(TSK-\d+)/i);
    if (match) {
      targetId = match[1].toUpperCase();
    }
  }

  const allTasks = getAllTodoTasks();
  if (!targetId && allTasks.length > 0) {
    targetId = allTasks[0].id;
  }

  const events = targetId ? getTaskTimelineHistory(targetId) : [];
  const taskDetails = targetId ? getTodoTaskById(targetId) : null;

  let htmlOutput;
  const candidateNames = ['timeline/TimelineDialog', 'TimelineDialog', '14_TimelineDialog'];
  for (const name of candidateNames) {
    try {
      const template = HtmlService.createTemplateFromFile(name);
      template.tasks = allTasks;
      template.selectedTaskId = targetId || '';
      template.taskDetails = taskDetails;
      template.events = events;
      htmlOutput = template.evaluate();
      if (htmlOutput) break;
    } catch (ignore) {
      // Continue next candidate
    }
  }

  if (!htmlOutput) {
    const fallbackHtml = getFallbackTimelineDialogHtml(allTasks, targetId, events);
    htmlOutput = HtmlService.createHtmlOutput(fallbackHtml);
  }

  htmlOutput.setWidth(780).setHeight(620);
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, `Task Timeline History — ${targetId || 'Overview'}`);
}


/**
 * Client-callable function: Loads timeline events for a selected Task ID inside the modal.
 *
 * @param {string} taskId
 * @returns {Object}
 */
function loadTaskTimelineData(taskId) {
  const events = getTaskTimelineHistory(taskId);
  const taskDetails = getTodoTaskById(taskId);
  return {
    success: true,
    taskId: taskId,
    taskDetails: taskDetails,
    events: events
  };
}


/**
 * Fallback HTML for Record Decision dialog.
 */
function getFallbackDecisionDialogHtml(tasks) {
  return `<div style="font-family:sans-serif;padding:20px;">
    <h3>Record Decision / Event</h3>
    <p>Please ensure 14_DecisionDialog.html is deployed.</p>
  </div>`;
}


/**
 * Fallback HTML for Timeline Viewer dialog.
 */
function getFallbackTimelineDialogHtml(tasks, targetId, events) {
  return `<div style="font-family:sans-serif;padding:20px;">
    <h3>Task History — ${targetId}</h3>
    <p>Please ensure 14_TimelineDialog.html is deployed.</p>
  </div>`;
}
