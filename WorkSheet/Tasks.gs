// ============================================================
// TASKS.GS — FAST TASK ENTRY & BATCH INSERTION ENGINE
// ============================================================

/**
 * Menu Action: Fill task for today.
 * Opens the streamlined task entry dialog for today's date.
 */
function fillTaskForToday() {
  const today = getToday();
  openTaskDialog(today);
}

/**
 * Menu Action: Fill task for selected date in the current month.
 * Opens an interactive visual calendar picker to choose the target date.
 */
function fillTaskForSelectedDate() {
  const ss = getSpreadsheet();
  const sheet = ss.getActiveSheet();
  const timezone = getTimezone();

  let initialDate = getToday();
  const activeRow = sheet.getActiveCell().getRow();

  if (activeRow > 1 && sheet.getName() !== CONFIG.LISTS_SHEET_NAME) {
    const cellValue = sheet.getRange(activeRow, 1).getValue();
    const parsed = parseDateFromCell(cellValue, sheet.getName(), timezone);
    if (parsed) {
      initialDate = parsed;
    }
  }

  openCalendarPicker('tasks', initialDate);
}

/**
 * Opens the streamlined Task Entry modal dialog.
 * Allows adding multiple tasks for a single project with shared category & priority.
 * Status is automatically "Pending", and no Notes field is requested.
 *
 * @param {Date} targetDate
 */
function openTaskDialog(targetDate) {
  const ss = getSpreadsheet();
  const timezone = getTimezone();
  const date = targetDate || getToday();

  // Ensure Lists sheet exists so project/category lists are available
  ensureListsSheet(ss);

  const currentYear = date.getFullYear();
  const currentMonth = date.getMonth();
  const currentDay = date.getDate();
  const daysInMonth = new Date(currentYear, currentMonth + 1, 0).getDate();
  const monthSheetName = getMonthSheetName(date, timezone);

  // Ensure target month sheet exists
  let monthSheet = ss.getSheetByName(monthSheetName);
  if (!monthSheet) {
    createCurrentMonthSheet();
    monthSheet = ss.getSheetByName(monthSheetName);
  }

  const dateInfo = {
    day: currentDay,
    monthIndex: currentMonth,
    year: currentYear,
    daysInMonth: daysInMonth,
    formattedDate: Utilities.formatDate(date, timezone, 'EEEE, dd MMMM yyyy'),
    shortDate: Utilities.formatDate(date, timezone, 'dd/MM/yyyy'),
    isoDate: Utilities.formatDate(date, timezone, 'yyyy-MM-dd'),
    monthSheetName: monthSheetName
  };

  const configLists = {
    projects: (CONFIG.LISTS && CONFIG.LISTS.Projects) || [],
    categories: (CONFIG.LISTS && CONFIG.LISTS.Categories) || [],
    priorities: (CONFIG.LISTS && CONFIG.LISTS.Priorities) || ['Low', 'Medium', 'High', 'Urgent']
  };

  const htmlContent = getTaskDialogHtml(dateInfo, configLists);
  const htmlOutput = HtmlService.createHtmlOutput(htmlContent)
    .setWidth(550)
    .setHeight(585);

  SpreadsheetApp.getUi().showModalDialog(
    htmlOutput,
    `Add Daily Tasks — ${dateInfo.shortDate}`
  );
}

/**
 * Client-callable function: Saves a batch of tasks for a single project
 * directly into the Month sheet.
 * All added tasks automatically receive Status = 'Pending' without Notes.
 *
 * @param {Object} payload
 * @returns {Object} Result object
 */
function saveTasksBatch(payload) {
  const ss = getSpreadsheet();
  const timezone = getTimezone();

  if (!payload || !payload.tasks || !payload.tasks.length) {
    return {
      success: false,
      message: 'Please enter at least one task.'
    };
  }

  const project = String(payload.project || '').trim();
  const category = String(payload.category || '').trim();
  const priority = String(payload.priority || 'Medium').trim();
  const status = 'Pending'; // Always Pending automatically

  // Clean task lines (strip leading bullet characters: -, *, •, 1., etc.)
  const cleanedTasks = [];
  payload.tasks.forEach(t => {
    let clean = String(t || '').trim();
    clean = clean.replace(/^[\-\*\•\d+\.\)]\s*/, '').trim();
    if (clean) {
      cleanedTasks.push(clean);
    }
  });

  if (cleanedTasks.length === 0) {
    return {
      success: false,
      message: 'Please enter at least one valid task description.'
    };
  }

  // Resolve target Date
  const year = parseInt(payload.year, 10);
  const month = parseInt(payload.monthIndex, 10);
  const day = parseInt(payload.day, 10);
  const targetDate = new Date(year, month, day);

  const targetMonthSheetName = getMonthSheetName(targetDate, timezone);
  let monthSheet = ss.getSheetByName(targetMonthSheetName);

  if (!monthSheet) {
    createCurrentMonthSheet();
    monthSheet = ss.getSheetByName(targetMonthSheetName);
  }

  if (!monthSheet) {
    return {
      success: false,
      message: `Month sheet "${targetMonthSheetName}" could not be loaded.`
    };
  }

  // Find existing rows for targetDate in monthSheet
  const monthLastRow = monthSheet.getLastRow();
  const dateColValues = monthSheet
    .getRange(2, 1, Math.max(monthLastRow - 1, 1), 1)
    .getValues();

  const matchingRowIndices = [];
  for (let r = 0; r < dateColValues.length; r++) {
    if (isMatchingDate(dateColValues[r][0], targetDate, timezone)) {
      matchingRowIndices.push(r + 2); // 1-indexed sheet row
    }
  }

  let firstSavedRow = 2;

  if (matchingRowIndices.length > 0) {
    const firstRowIndex = matchingRowIndices[0];
    const existingTask = monthSheet.getRange(firstRowIndex, 4).getValue(); // Col 4 is Task
    const isFirstRowEmpty = !existingTask || String(existingTask).trim() === '';

    let remainingTasks = cleanedTasks;
    let insertAfterRow = matchingRowIndices[matchingRowIndices.length - 1];

    if (isFirstRowEmpty) {
      // 1. Fill first task into the existing template row for this date
      const t1 = cleanedTasks[0];
      monthSheet.getRange(firstRowIndex, 3, 1, 5).setValues([[
        project,
        t1,
        category,
        priority,
        status
      ]]);
      firstSavedRow = firstRowIndex;
      remainingTasks = cleanedTasks.slice(1);
      insertAfterRow = firstRowIndex;
    }

    // 2. If there are more tasks, insert new rows directly after with the SAME Date and Day
    const dateVal = monthSheet.getRange(firstRowIndex, 1).getValue();
    const dayVal = monthSheet.getRange(firstRowIndex, 2).getValue();

    remainingTasks.forEach(taskText => {
      monthSheet.insertRowAfter(insertAfterRow);
      const newRow = insertAfterRow + 1;

      // Columns: Date, Day, Project, Task, Category, Priority, Status (7 columns)
      monthSheet.getRange(newRow, 1, 1, CONFIG.HEADERS.length).setValues([[
        dateVal,
        dayVal,
        project,
        taskText,
        category,
        priority,
        status
      ]]);

      monthSheet.setRowHeight(newRow, 42);

      monthSheet
        .getRange(newRow, 1, 1, CONFIG.HEADERS.length)
        .setFontFamily('Arial')
        .setFontSize(10)
        .setVerticalAlignment('middle')
        .setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);

      monthSheet
        .getRange(newRow, 1, 1, 2)
        .setFontFamily('Roboto Mono')
        .setFontWeight('bold')
        .setHorizontalAlignment('center');

      monthSheet.getRange(newRow, 3).setHorizontalAlignment('center');
      monthSheet.getRange(newRow, 4).setHorizontalAlignment('left');
      monthSheet.getRange(newRow, 5, 1, 3).setHorizontalAlignment('center');

      formatBorders(monthSheet.getRange(newRow, 1, 1, CONFIG.HEADERS.length));
      setupDropdownsForRow(monthSheet, newRow);

      insertAfterRow = newRow;
    });
  } else {
    // Fallback: Append tasks to the bottom of the month sheet
    const fallbackDateLabel = Utilities.formatDate(targetDate, timezone, CONFIG.DATE_FORMAT || 'MMdd');
    const fallbackDayLabel = Utilities.formatDate(targetDate, timezone, 'EEE');

    cleanedTasks.forEach(taskText => {
      const newRow = monthSheet.getLastRow() + 1;
      monthSheet.getRange(newRow, 1, 1, CONFIG.HEADERS.length).setValues([[
        fallbackDateLabel,
        fallbackDayLabel,
        project,
        taskText,
        category,
        priority,
        status
      ]]);

      monthSheet.setRowHeight(newRow, 42);
      setupDropdownsForRow(monthSheet, newRow);
      formatBorders(monthSheet.getRange(newRow, 1, 1, CONFIG.HEADERS.length));
    });
  }

  // Refocus user on the month sheet and highlight the added row
  ss.setActiveSheet(monthSheet);
  monthSheet.setActiveSelection(`D${firstSavedRow}`);

  return {
    success: true,
    count: cleanedTasks.length,
    project: project || 'General',
    monthSheetName: targetMonthSheetName,
    firstRow: firstSavedRow
  };
}

/**
 * Applies dropdown validation rules for the 7-column month sheet.
 * Col 3: Projects, Col 5: Categories, Col 6: Priorities, Col 7: Statuses
 */
function setupDropdowns(sheet) {
  const rowCount = Math.max(sheet.getLastRow() - 1, 20);
  const ss = sheet.getParent();

  const rules = {
    3: createDropdownRule(ss, 'Projects'),
    5: createDropdownRule(ss, 'Categories'),
    6: createDropdownRule(ss, 'Priorities'),
    7: createDropdownRule(ss, 'Statuses')
  };

  Object.entries(rules).forEach(([column, rule]) => {
    if (rule) {
      sheet
        .getRange(2, Number(column), rowCount, 1)
        .setDataValidation(rule);
    }
  });
}

/**
 * Applies dropdown validation rules for a single newly inserted row in the month sheet.
 */
function setupDropdownsForRow(sheet, row) {
  const ss = sheet.getParent();

  const rules = {
    3: createDropdownRule(ss, 'Projects'),
    5: createDropdownRule(ss, 'Categories'),
    6: createDropdownRule(ss, 'Priorities'),
    7: createDropdownRule(ss, 'Statuses')
  };

  Object.entries(rules).forEach(([column, rule]) => {
    if (rule) {
      sheet
        .getRange(row, Number(column), 1, 1)
        .setDataValidation(rule);
    }
  });
}

/**
 * Builds data validation rule for a given Named Range.
 */
function createDropdownRule(ss, namedRange) {
  const range = ss.getRangeByName(namedRange);
  if (!range) {
    return null;
  }

  return SpreadsheetApp
    .newDataValidation()
    .requireValueInRange(range, true)
    .setAllowInvalid(false)
    .build();
}

/**
 * Clears task data while safeguarding against accidental clearing of the Lists sheet.
 * Clears Columns 3 to 7 (Project through Status), resetting Status to 'Pending'.
 */
function clearTasks() {
  const ui = SpreadsheetApp.getUi();
  const ss = getSpreadsheet();
  const sheet = ss.getActiveSheet();

  if (sheet.getName() === CONFIG.LISTS_SHEET_NAME) {
    ui.alert('Cannot clear tasks on the Lists sheet.');
    return;
  }

  const response = ui.alert(
    'Clear Tasks',
    `Clear all task entries from "${sheet.getName()}"? Dates and days will be preserved.`,
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) {
    return;
  }

  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) {
    return;
  }

  // Clear Project through Status (Cols 3 to 7: 5 columns)
  sheet
    .getRange(2, 3, lastRow - 1, 5)
    .clearContent();

  // Restore default status to Pending (Col 7)
  sheet
    .getRange(2, 7, lastRow - 1, 1)
    .setValue('Pending');
}

/**
 * Backwards compatibility helper: redirects to fillTaskForToday().
 */
function addTaskRow() {
  fillTaskForToday();
}

/**
 * Backwards compatibility alias: redirects staging sheet call to openTaskDialog.
 */
function openTaskStagingSheet(targetDate) {
  openTaskDialog(targetDate);
}

/**
 * Generates HTML for the Task Entry Modal Dialog.
 *
 * @param {Object} dateInfo
 * @param {Object} configLists
 * @returns {string} HTML content
 */
function getTaskDialogHtml(dateInfo, configLists) {
  const jsonDateInfo = JSON.stringify(dateInfo).replace(/<\/script/gi, '<\\/script');
  const jsonConfig = JSON.stringify(configLists).replace(/<\/script/gi, '<\\/script');

  return `<!DOCTYPE html>
<html>
<head>
  <base target="_top">
  <style>
    * {
      box-sizing: border-box;
      margin: 0;
      padding: 0;
    }

    html, body {
      margin: 0;
      padding: 0;
      box-sizing: border-box;
      background-color: #f8fafc;
      font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Arial, sans-serif;
      color: #1e293b;
      font-size: 13px;
      line-height: 1.4;
      overflow-y: auto;
    }

    body {
      padding: 12px 14px;
    }

    .card {
      background: #ffffff;
      border: 1px solid #e2e8f0;
      border-radius: 10px;
      padding: 14px 16px;
      margin-bottom: 0;
      box-shadow: 0 1px 2px rgba(0, 0, 0, 0.04);
    }

    .header-row {
      display: flex;
      justify-content: space-between;
      align-items: center;
      margin-bottom: 10px;
    }

    .title-area {
      display: flex;
      align-items: center;
      gap: 8px;
    }

    .title-icon {
      font-size: 18px;
    }

    .title-text {
      font-size: 15px;
      font-weight: 700;
      color: #0f172a;
    }

    .badge-tab {
      background: #ede9fe;
      color: #6366f1;
      font-size: 11px;
      font-weight: 700;
      padding: 2px 8px;
      border-radius: 6px;
    }

    .date-row {
      display: flex;
      align-items: center;
      gap: 10px;
      background: #f1f5f9;
      padding: 8px 12px;
      border-radius: 8px;
      margin-bottom: 12px;
      font-size: 12px;
    }

    .date-label {
      font-weight: 600;
      color: #475569;
    }

    .date-picker-input {
      padding: 4px 8px;
      font-size: 12px;
      font-weight: 600;
      border: 1px solid #cbd5e1;
      border-radius: 6px;
      background: #ffffff;
      color: #0f172a;
      outline: none;
      cursor: pointer;
    }

    .date-picker-input:focus {
      border-color: #6366f1;
      box-shadow: 0 0 0 2px rgba(99, 102, 241, 0.15);
    }

    .date-display-text {
      color: #334155;
      font-weight: 600;
      margin-left: auto;
    }

    .form-grid {
      display: grid;
      grid-template-columns: 1fr 1fr;
      gap: 10px;
      margin-bottom: 12px;
    }

    .form-group {
      display: flex;
      flex-direction: column;
      gap: 4px;
    }

    .form-group.full-width {
      grid-column: span 2;
    }

    label {
      font-size: 11px;
      font-weight: 700;
      text-transform: uppercase;
      letter-spacing: 0.03em;
      color: #64748b;
    }

    select, textarea {
      font-family: inherit;
      font-size: 13px;
      border: 1px solid #cbd5e1;
      border-radius: 6px;
      padding: 7px 10px;
      background-color: #ffffff;
      color: #0f172a;
      outline: none;
      transition: border-color 0.15s ease-in-out, box-shadow 0.15s ease-in-out;
    }

    select:focus, textarea:focus {
      border-color: #6366f1;
      box-shadow: 0 0 0 3px rgba(99, 102, 241, 0.15);
    }

    .status-auto-row {
      display: flex;
      align-items: center;
      justify-content: space-between;
      background: #f8fafc;
      border: 1px dashed #cbd5e1;
      border-radius: 6px;
      padding: 6px 10px;
    }

    .badge-pending {
      background: #fef3c7;
      color: #b45309;
      font-size: 11px;
      font-weight: 700;
      padding: 2px 8px;
      border-radius: 9999px;
      border: 1px solid #fde68a;
    }

    .task-area-label-row {
      display: flex;
      justify-content: space-between;
      align-items: center;
      margin-bottom: 4px;
    }

    .counter-badge {
      font-size: 11px;
      font-weight: 600;
      color: #6366f1;
      background: #ede9fe;
      padding: 1px 7px;
      border-radius: 9999px;
    }

    textarea {
      width: 100%;
      min-height: 110px;
      max-height: 160px;
      resize: vertical;
      line-height: 1.4;
    }

    .tip-text {
      font-size: 11px;
      color: #94a3b8;
      margin-top: 4px;
      display: flex;
      align-items: center;
      gap: 4px;
    }

    .btn-container {
      display: flex;
      gap: 8px;
      margin-top: 14px;
    }

    .btn {
      display: inline-flex;
      align-items: center;
      justify-content: center;
      gap: 6px;
      padding: 9px 14px;
      font-size: 12px;
      font-weight: 600;
      border-radius: 6px;
      cursor: pointer;
      transition: all 0.15s ease-in-out;
      outline: none;
      user-select: none;
      border: 1px solid transparent;
    }

    .btn-primary {
      background-color: #4f46e5;
      color: #ffffff;
      flex: 1.2;
      box-shadow: 0 1px 2px rgba(79, 70, 229, 0.3);
    }

    .btn-primary:hover {
      background-color: #4338ca;
    }

    .btn-secondary {
      background-color: #ffffff;
      color: #4f46e5;
      border-color: #c7d2fe;
      flex: 1.2;
    }

    .btn-secondary:hover {
      background-color: #eef2ff;
      border-color: #a5b4fc;
    }

    .btn-cancel {
      background-color: #ffffff;
      color: #64748b;
      border-color: #e2e8f0;
      flex: 0.6;
    }

    .btn-cancel:hover {
      background-color: #f1f5f9;
      color: #0f172a;
    }

    .btn:disabled {
      opacity: 0.6;
      cursor: not-allowed;
    }

    .status-toast {
      margin-top: 10px;
      padding: 8px 12px;
      border-radius: 6px;
      font-size: 12px;
      display: none;
      text-align: center;
    }

    .status-toast.success {
      display: block;
      background: #f0fdf4;
      color: #166534;
      border: 1px solid #bbf7d0;
    }

    .status-toast.error {
      display: block;
      background: #fef2f2;
      color: #991b1b;
      border: 1px solid #fecaca;
    }
  </style>
</head>
<body>
  <div class="card">
    <div class="header-row">
      <div class="title-area">
        <span class="title-icon">⚡</span>
        <span class="title-text">Add Daily Tasks</span>
      </div>
      <span class="badge-tab" id="monthBadge"></span>
    </div>

    <!-- Date Row with Calendar Picker -->
    <div class="date-row">
      <span class="date-label">📅 Date:</span>
      <input type="date" id="dateInput" class="date-picker-input" value="${dateInfo.isoDate}" onchange="handleDateChange()">
      <span id="dateDisplay" class="date-display-text">${dateInfo.formattedDate}</span>
    </div>

    <!-- Form Controls for Project, Category, Priority -->
    <div class="form-grid">
      <div class="form-group">
        <label for="projectSelect">Project</label>
        <select id="projectSelect"></select>
      </div>

      <div class="form-group">
        <label for="categorySelect">Category</label>
        <select id="categorySelect"></select>
      </div>

      <div class="form-group">
        <label for="prioritySelect">Priority</label>
        <select id="prioritySelect"></select>
      </div>

      <div class="form-group">
        <label>Status</label>
        <div class="status-auto-row">
          <span style="font-size: 11px; color: #64748b;">Default</span>
          <span class="badge-pending">Pending (Auto)</span>
        </div>
      </div>

      <!-- Tasks Textarea -->
      <div class="form-group full-width">
        <div class="task-area-label-row">
          <label for="tasksInput">Tasks (one per line)</label>
          <span id="taskCountBadge" class="counter-badge">0 tasks</span>
        </div>
        <textarea
          id="tasksInput"
          placeholder="Paste or type tasks for this project...&#10;• Example task 1&#10;• Example task 2"
          oninput="updateTaskCount()"
        ></textarea>
        <div class="tip-text">
          <span>⌨️ Shortcut: Press <strong>Ctrl + Enter</strong> to add tasks</span>
        </div>
      </div>
    </div>

    <!-- Action Buttons -->
    <div class="btn-container">
      <button id="saveBtn" class="btn btn-primary" onclick="submitBatch(false)">
        📥 Add Tasks
      </button>
      <button id="saveAndNextBtn" class="btn btn-secondary" onclick="submitBatch(true)">
        + Add & Next Project
      </button>
      <button id="closeBtn" class="btn btn-cancel" onclick="google.script.host.close()">
        Close
      </button>
    </div>

    <div id="statusToast" class="status-toast"></div>
  </div>

  <script>
    const dateInfo = ${jsonDateInfo};
    const configLists = ${jsonConfig};

    // Initialize UI on load
    window.onload = function() {
      initDropdowns();
      document.getElementById('monthBadge').innerText = dateInfo.monthSheetName;
      document.getElementById('tasksInput').focus();
    };

    function handleDateChange() {
      const inputVal = document.getElementById('dateInput').value;
      if (!inputVal) return;
      const parts = inputVal.split('-');
      dateInfo.year = parseInt(parts[0], 10);
      dateInfo.monthIndex = parseInt(parts[1], 10) - 1;
      dateInfo.day = parseInt(parts[2], 10);

      const d = new Date(dateInfo.year, dateInfo.monthIndex, dateInfo.day);
      const options = { weekday: 'short', day: 'numeric', month: 'short', year: 'numeric' };
      document.getElementById('dateDisplay').innerText = d.toLocaleDateString('en-GB', options);

      // Update month badge e.g. "SEP26"
      const monthNames = ['JAN', 'FEB', 'MAR', 'APR', 'MAY', 'JUN', 'JUL', 'AUG', 'SEP', 'OCT', 'NOV', 'DEC'];
      const yy = String(dateInfo.year).slice(-2);
      dateInfo.monthSheetName = monthNames[dateInfo.monthIndex] + yy;
      document.getElementById('monthBadge').innerText = dateInfo.monthSheetName;
    }

    function initDropdowns() {
      // Projects
      const projSelect = document.getElementById('projectSelect');
      projSelect.innerHTML = '';
      (configLists.projects || []).forEach(function(p) {
        const opt = document.createElement('option');
        opt.value = p;
        opt.text = p;
        projSelect.appendChild(opt);
      });

      // Categories
      const catSelect = document.getElementById('categorySelect');
      catSelect.innerHTML = '';
      (configLists.categories || []).forEach(function(c) {
        const opt = document.createElement('option');
        opt.value = c;
        opt.text = c;
        if (c.toLowerCase() === 'development') {
          opt.selected = true;
        }
        catSelect.appendChild(opt);
      });

      // Priorities
      const prioSelect = document.getElementById('prioritySelect');
      prioSelect.innerHTML = '';
      (configLists.priorities || []).forEach(function(pr) {
        const opt = document.createElement('option');
        opt.value = pr;
        opt.text = pr;
        if (pr.toLowerCase() === 'medium') {
          opt.selected = true;
        }
        prioSelect.appendChild(opt);
      });
    }

    function getTaskLines() {
      const text = document.getElementById('tasksInput').value || '';
      return text
        .split('\\n')
        .map(function(line) { return line.trim(); })
        .filter(function(line) { return line.length > 0; });
    }

    function updateTaskCount() {
      const count = getTaskLines().length;
      const badge = document.getElementById('taskCountBadge');
      badge.innerText = count + (count === 1 ? ' task' : ' tasks');
    }

    // Ctrl + Enter shortcut
    document.addEventListener('keydown', function(e) {
      if ((e.ctrlKey || e.metaKey) && e.key === 'Enter') {
        e.preventDefault();
        submitBatch(false);
      }
    });

    function submitBatch(keepOpen) {
      const tasks = getTaskLines();
      const toast = document.getElementById('statusToast');

      if (tasks.length === 0) {
        toast.className = 'status-toast error';
        toast.innerText = 'Please write or paste at least one task.';
        document.getElementById('tasksInput').focus();
        return;
      }

      const project = document.getElementById('projectSelect').value;
      const category = document.getElementById('categorySelect').value;
      const priority = document.getElementById('prioritySelect').value;

      const payload = {
        day: dateInfo.day,
        monthIndex: dateInfo.monthIndex,
        year: dateInfo.year,
        project: project,
        category: category,
        priority: priority,
        tasks: tasks
      };

      const saveBtn = document.getElementById('saveBtn');
      const saveAndNextBtn = document.getElementById('saveAndNextBtn');

      saveBtn.disabled = true;
      saveAndNextBtn.disabled = true;
      saveBtn.innerText = '⏳ Saving...';
      toast.style.display = 'none';

      google.script.run
        .withSuccessHandler(function(res) {
          saveBtn.disabled = false;
          saveAndNextBtn.disabled = false;
          saveBtn.innerText = '📥 Add Tasks';

          if (!res || !res.success) {
            toast.className = 'status-toast error';
            toast.innerText = (res && res.message) || 'Failed to save tasks.';
            return;
          }

          if (keepOpen) {
            // Reset task textarea for next project
            document.getElementById('tasksInput').value = '';
            updateTaskCount();

            toast.className = 'status-toast success';
            toast.innerText = '✓ Added ' + res.count + ' task(s) for ' + res.project + '! Ready for next project.';
            document.getElementById('tasksInput').focus();
          } else {
            toast.className = 'status-toast success';
            toast.innerText = '✓ Saved ' + res.count + ' task(s) to ' + res.monthSheetName + '!';

            setTimeout(function() {
              google.script.host.close();
            }, 900);
          }
        })
        .withFailureHandler(function(err) {
          saveBtn.disabled = false;
          saveAndNextBtn.disabled = false;
          saveBtn.innerText = '📥 Add Tasks';
          toast.className = 'status-toast error';
          toast.innerText = 'Error: ' + (err.message || err);
        })
        .saveTasksBatch(payload);
    }
  </script>
</body>
</html>`;
}