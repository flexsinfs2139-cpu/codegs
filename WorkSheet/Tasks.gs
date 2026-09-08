// ============================================================
// TASKS.GS — TASK STAGING & INSERTION ENGINE
// ============================================================

/**
 * Headers for the temporary Task Entry (Staging) sheet.
 * Date and Day are omitted here as requested; they are inferred from the sheet date.
 */
const STAGING_HEADERS = [
  'Project',
  'Task',
  'Category',
  'Priority',
  'Status',
  'Notes'
];

/**
 * Menu Action: Fill task for today.
 * Opens or creates a staging sheet named "Date Month Year" (e.g. "08 September 2026").
 */
function fillTaskForToday() {
  const timezone = getTimezone();
  const today = getToday();
  openTaskStagingSheet(today);
}

/**
 * Menu Action: Fill task for selected date in the current month.
 * Asks ONLY for a day number between 1 and the total days of the current month.
 */
function fillTaskForSelectedDate() {
  const timezone = getTimezone();
  const today = getToday();
  const currentYear = today.getFullYear();
  const currentMonth = today.getMonth(); // 0-indexed (e.g. 8 for September)

  // Calculate total days in the current month (28, 29, 30, or 31)
  const daysInMonth = new Date(currentYear, currentMonth + 1, 0).getDate();

  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt(
    'Fill Task for Selected Date',
    `Enter day of the current month (1–${daysInMonth}):`,
    ui.ButtonSet.OK_CANCEL
  );

  if (response.getSelectedButton() !== ui.Button.OK) {
    return;
  }

  const input = response.getResponseText().trim();
  const dayNumber = parseInt(input, 10);

  if (isNaN(dayNumber) || dayNumber < 1 || dayNumber > daysInMonth) {
    ui.alert(
      'Invalid Day Number',
      `Please enter a valid number between 1 and ${daysInMonth}.`,
      ui.ButtonSet.OK
    );
    return;
  }

  const targetDate = new Date(currentYear, currentMonth, dayNumber);
  openTaskStagingSheet(targetDate);
}

/**
 * Creates and sets up the temporary "Date Month Year" staging sheet
 * without Date or Day columns, and displays the task entry sidebar.
 *
 * @param {Date} targetDate
 */
function openTaskStagingSheet(targetDate) {
  const ss = getSpreadsheet();
  const timezone = getTimezone();
  const stagingSheetName = getStagingSheetName(targetDate, timezone);
  const targetMonthSheetName = getMonthSheetName(targetDate, timezone);

  // Ensure reference lists exist
  ensureListsSheet(ss);

  // Ensure the target month sheet exists
  let monthSheet = ss.getSheetByName(targetMonthSheetName);
  if (!monthSheet) {
    createCurrentMonthSheet();
    monthSheet = ss.getSheetByName(targetMonthSheetName);
  }

  // If staging sheet already exists, activate it and reopen sidebar
  let stagingSheet = ss.getSheetByName(stagingSheetName);
  if (stagingSheet) {
    ss.setActiveSheet(stagingSheet);
    showTaskStagingSidebar(targetDate, stagingSheetName, targetMonthSheetName);
    return;
  }

  // Create the staging sheet
  stagingSheet = ss.insertSheet(stagingSheetName);

  // Store target date in developer metadata for exact retrieval
  try {
    stagingSheet.addDeveloperMetadata(
      'TARGET_DATE',
      Utilities.formatDate(targetDate, timezone, 'yyyy-MM-dd')
    );
  } catch (e) {}

  // 1. Setup Header (6 task columns, omitting Date and Day)
  stagingSheet
    .getRange(1, 1, 1, STAGING_HEADERS.length)
    .setValues([STAGING_HEADERS]);

  stagingSheet
    .getRange(1, 1, 1, STAGING_HEADERS.length)
    .setBackground(CONFIG.COLORS.HEADER)
    .setFontFamily('Arial')
    .setFontWeight('bold')
    .setHorizontalAlignment('center');

  stagingSheet.setRowHeight(1, 28);

  // 2. Pre-fill initial 5 blank task rows
  const initialRows = [];
  for (let i = 0; i < 5; i++) {
    initialRows.push([
      '',          // Project
      '',          // Task
      '',          // Category
      'Medium',    // Priority
      'Pending',   // Status
      ''           // Notes
    ]);
  }

  stagingSheet
    .getRange(2, 1, initialRows.length, STAGING_HEADERS.length)
    .setValues(initialRows);

  // 3. Formatting dimensions and styling
  // Project(140), Task(340), Category(130), Priority(100), Status(130), Notes(300)
  const stagingWidths = [140, 340, 130, 100, 130, 300];
  stagingWidths.forEach((w, idx) => {
    stagingSheet.setColumnWidth(idx + 1, w);
  });

  stagingSheet.setRowHeights(2, 15, 42);

  stagingSheet.getRange('A2:A').setHorizontalAlignment('center');
  stagingSheet.getRange('B2:B').setHorizontalAlignment('left');
  stagingSheet.getRange('C2:E').setHorizontalAlignment('center');
  stagingSheet.getRange('F2:F').setHorizontalAlignment('left');

  stagingSheet
    .getRange(1, 1, 20, STAGING_HEADERS.length)
    .setFontFamily('Arial')
    .setFontSize(10)
    .setVerticalAlignment('middle')
    .setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);

  formatBorders(stagingSheet.getRange(1, 1, 15, STAGING_HEADERS.length));

  // 4. Setup dropdown validations for the 6 staging columns
  setupDropdownsForStagingSheet(stagingSheet);

  // 5. Freeze header & trim excess
  stagingSheet.setFrozenRows(1);
  trimSheet(stagingSheet, 20, STAGING_HEADERS.length);

  // 6. Focus on the first Task cell (B2)
  ss.setActiveSheet(stagingSheet);
  stagingSheet.setActiveSelection('B2');

  // 7. Open the submission sidebar
  showTaskStagingSidebar(targetDate, stagingSheetName, targetMonthSheetName);
}

/**
 * Applies dropdown validation rules for the staging sheet (6 columns).
 * Col 1: Projects, Col 3: Categories, Col 4: Priorities, Col 5: Statuses
 */
function setupDropdownsForStagingSheet(sheet) {
  const rowCount = Math.max(sheet.getLastRow() - 1, 20);
  const ss = sheet.getParent();

  const rules = {
    1: createDropdownRule(ss, 'Projects'),
    3: createDropdownRule(ss, 'Categories'),
    4: createDropdownRule(ss, 'Priorities'),
    5: createDropdownRule(ss, 'Statuses')
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
 * Opens the Task Entry Sidebar in Google Sheets.
 */
function showTaskStagingSidebar(targetDate, stagingSheetName, monthSheetName) {
  const timezone = getTimezone();
  const formattedDate = Utilities.formatDate(targetDate, timezone, 'dd MMMM yyyy');
  const shortDate = Utilities.formatDate(targetDate, timezone, 'dd/MM/yyyy');

  const htmlContent = getTaskSidebarHtml(
    formattedDate,
    shortDate,
    stagingSheetName,
    monthSheetName
  );

  const htmlOutput = HtmlService.createHtmlOutput(htmlContent)
    .setTitle('Task Entry');

  SpreadsheetApp.getUi().showSidebar(htmlOutput);
}

/**
 * Client-callable function: Saves tasks from the staging sheet into the month sheet,
 * populates Date and Day automatically, inserts additional rows with the same date
 * if more than one task was entered, and deletes the staging sheet.
 *
 * @param {string} stagingSheetName
 * @returns {Object} Result object
 */
function saveTasksFromStagingSheet(stagingSheetName) {
  const ss = getSpreadsheet();
  const timezone = getTimezone();
  const tempSheet = ss.getSheetByName(stagingSheetName);

  if (!tempSheet) {
    return {
      success: false,
      message: `Staging sheet "${stagingSheetName}" was not found.`
    };
  }

  const lastRow = tempSheet.getLastRow();
  if (lastRow < 2) {
    return {
      success: false,
      message: 'No tasks entered. Please enter at least one task before saving.'
    };
  }

  // Read entered tasks from the 6 staging columns:
  // [Project, Task, Category, Priority, Status, Notes]
  const rawData = tempSheet
    .getRange(2, 1, lastRow - 1, STAGING_HEADERS.length)
    .getValues();

  const enteredTasks = [];
  rawData.forEach(row => {
    const taskDesc = row[1] ? String(row[1]).trim() : '';
    if (taskDesc) {
      enteredTasks.push({
        project: row[0] || '',
        task: taskDesc,
        category: row[2] || '',
        priority: row[3] || 'Medium',
        status: row[4] || 'Pending',
        notes: row[5] || ''
      });
    }
  });

  if (enteredTasks.length === 0) {
    return {
      success: false,
      message: 'No tasks found. Please enter at least one task title in the "Task" column.'
    };
  }

  // Resolve target date (via developer metadata or sheet name)
  let targetDate = null;
  try {
    const metaList = tempSheet.getDeveloperMetadata();
    const targetMeta = metaList.find(m => m.getKey() === 'TARGET_DATE');
    if (targetMeta) {
      const parts = targetMeta.getValue().split('-');
      targetDate = new Date(
        parseInt(parts[0], 10),
        parseInt(parts[1], 10) - 1,
        parseInt(parts[2], 10)
      );
    }
  } catch (e) {}

  if (!targetDate) {
    targetDate = parseDateFromStagingSheetName(stagingSheetName) || getToday();
  }

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

  // Find all rows in monthSheet matching targetDate
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
    const existingTask = monthSheet.getRange(firstRowIndex, 4).getValue();
    const isFirstRowEmpty = !existingTask || String(existingTask).trim() === '';

    let remainingTasks = enteredTasks;
    let insertAfterRow = matchingRowIndices[matchingRowIndices.length - 1];

    if (isFirstRowEmpty) {
      // 1. Fill first task into the existing row for this date
      const t1 = enteredTasks[0];
      monthSheet.getRange(firstRowIndex, 3, 1, 6).setValues([[
        t1.project,
        t1.task,
        t1.category,
        t1.priority,
        t1.status,
        t1.notes
      ]]);
      firstSavedRow = firstRowIndex;
      remainingTasks = enteredTasks.slice(1);
      insertAfterRow = firstRowIndex;
    }

    // 2. If there are more tasks for this day, insert new row(s) after with the SAME date and day
    const dateVal = monthSheet.getRange(firstRowIndex, 1).getValue();
    const dayVal = monthSheet.getRange(firstRowIndex, 2).getValue();

    remainingTasks.forEach(task => {
      monthSheet.insertRowAfter(insertAfterRow);
      const newRow = insertAfterRow + 1;

      monthSheet.getRange(newRow, 1, 1, CONFIG.HEADERS.length).setValues([[
        dateVal,
        dayVal,
        task.project,
        task.task,
        task.category,
        task.priority,
        task.status,
        task.notes
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
      monthSheet.getRange(newRow, 8).setHorizontalAlignment('left');

      formatBorders(monthSheet.getRange(newRow, 1, 1, CONFIG.HEADERS.length));
      setupDropdownsForRow(monthSheet, newRow);

      insertAfterRow = newRow;
    });
  } else {
    // Fallback: Append tasks to the bottom of the month sheet
    const fallbackDateLabel = Utilities.formatDate(targetDate, timezone, 'MMdd');
    const fallbackDayLabel = Utilities.formatDate(targetDate, timezone, 'EEE');

    enteredTasks.forEach(task => {
      const newRow = monthSheet.getLastRow() + 1;
      monthSheet.getRange(newRow, 1, 1, CONFIG.HEADERS.length).setValues([[
        fallbackDateLabel,
        fallbackDayLabel,
        task.project,
        task.task,
        task.category,
        task.priority,
        task.status,
        task.notes
      ]]);

      monthSheet.setRowHeight(newRow, 42);
      setupDropdownsForRow(monthSheet, newRow);
      formatBorders(monthSheet.getRange(newRow, 1, 1, CONFIG.HEADERS.length));
    });
  }

  // 3. Delete the temporary "Date Month Year" staging sheet
  ss.deleteSheet(tempSheet);

  // 4. Return user to the month sheet
  ss.setActiveSheet(monthSheet);
  monthSheet.setActiveSelection(`D${firstSavedRow}`);

  return {
    success: true,
    count: enteredTasks.length,
    monthSheetName: targetMonthSheetName,
    firstRow: firstSavedRow
  };
}

/**
 * Parses target date from staging sheet name (e.g. "08 September 2026").
 */
function parseDateFromStagingSheetName(sheetName) {
  if (!sheetName) return null;

  const match = sheetName.match(/^(\d{1,2})\s+([A-Za-z]+)\s+(\d{4})$/);
  if (match) {
    const day = parseInt(match[1], 10);
    const monthName = match[2].toLowerCase();
    const year = parseInt(match[3], 10);
    const monthNames = [
      'january', 'february', 'march', 'april', 'may', 'june',
      'july', 'august', 'september', 'october', 'november', 'december'
    ];
    const monthIndex = monthNames.indexOf(monthName);
    if (monthIndex !== -1) {
      return new Date(year, monthIndex, day);
    }
  }
  return null;
}

/**
 * Client-callable function: Discards and deletes the temporary staging sheet.
 */
function cancelTaskStagingSheet(stagingSheetName) {
  const ss = getSpreadsheet();
  const tempSheet = ss.getSheetByName(stagingSheetName);

  if (tempSheet) {
    ss.deleteSheet(tempSheet);
  }

  return { success: true };
}

/**
 * Applies dropdown validation rules for the 8-column month sheet.
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
 * Legacy helper: Appends a single task row at the bottom.
 */
function addTaskRow() {
  fillTaskForToday();
}

/**
 * Clears task data while safeguarding against accidental clearing of the Lists sheet.
 */
function clearTasks() {
  const ui = SpreadsheetApp.getUi();
  const sheet = getSpreadsheet().getActiveSheet();

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

  // Clear Project through Notes (Cols 3 to 8)
  sheet
    .getRange(2, 3, lastRow - 1, 6)
    .clearContent();

  // Restore default status to Pending
  sheet
    .getRange(2, 7, lastRow - 1, 1)
    .setValue('Pending');
}

/**
 * Generates HTML for the Task Entry Sidebar.
 */
function getTaskSidebarHtml(formattedDate, shortDate, stagingSheetName, monthSheetName) {
  const jsonStaging = JSON.stringify(stagingSheetName).replace(/<\/script/gi, '<\\/script');
  const jsonMonth = JSON.stringify(monthSheetName).replace(/<\/script/gi, '<\\/script');

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

    body {
      font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Arial, sans-serif;
      background-color: #f8fafc;
      color: #1e293b;
      padding: 16px;
      font-size: 13px;
      line-height: 1.5;
    }

    .header-card {
      background: #ffffff;
      border: 1px solid #e2e8f0;
      border-radius: 10px;
      padding: 14px 16px;
      margin-bottom: 16px;
      box-shadow: 0 1px 2px rgba(0,0,0,0.03);
    }

    .title-row {
      display: flex;
      align-items: center;
      gap: 8px;
      font-size: 15px;
      font-weight: 700;
      color: #0f172a;
      margin-bottom: 6px;
    }

    .date-badge {
      display: inline-block;
      background: #ede9fe;
      color: #6366f1;
      font-size: 11px;
      font-weight: 600;
      padding: 2px 8px;
      border-radius: 9999px;
    }

    .meta-text {
      font-size: 12px;
      color: #64748b;
      margin-top: 4px;
    }

    .instructions {
      background: #f1f5f9;
      border-radius: 8px;
      padding: 12px 14px;
      margin-bottom: 20px;
      font-size: 12px;
      color: #334155;
    }

    .instructions ol {
      padding-left: 18px;
      margin-top: 6px;
    }

    .instructions li {
      margin-bottom: 4px;
    }

    .btn-container {
      display: flex;
      flex-direction: column;
      gap: 10px;
    }

    .btn {
      display: flex;
      align-items: center;
      justify-content: center;
      gap: 6px;
      width: 100%;
      padding: 10px 14px;
      font-size: 13px;
      font-weight: 600;
      border-radius: 7px;
      cursor: pointer;
      transition: all 0.15s ease-in-out;
      outline: none;
      user-select: none;
    }

    .btn-save {
      background-color: #5850ec;
      color: #ffffff;
      border: 1px solid #4f46e5;
      box-shadow: 0 1px 3px rgba(88, 80, 236, 0.25);
    }

    .btn-save:hover {
      background-color: #4f46e5;
    }

    .btn-save:disabled {
      opacity: 0.6;
      cursor: not-allowed;
    }

    .btn-cancel {
      background-color: #ffffff;
      color: #64748b;
      border: 1px solid #e2e8f0;
    }

    .btn-cancel:hover {
      background-color: #f8fafc;
      color: #0f172a;
      border-color: #cbd5e1;
    }

    .status-msg {
      margin-top: 14px;
      padding: 10px;
      border-radius: 6px;
      font-size: 12px;
      display: none;
      text-align: center;
    }

    .status-msg.error {
      display: block;
      background: #fef2f2;
      color: #b91c1c;
      border: 1px solid #fecaca;
    }

    .status-msg.success {
      display: block;
      background: #f0fdf4;
      color: #15803d;
      border: 1px solid #bbf7d0;
    }
  </style>
</head>
<body>
  <div class="header-card">
    <div class="title-row">
      <span>📝</span>
      <span>Fill Daily Tasks</span>
    </div>
    <span class="date-badge">${shortDate}</span>
    <div class="meta-text">Target Tab: <strong>${monthSheetName}</strong></div>
    <div class="meta-text">Staging Tab: <strong>${stagingSheetName}</strong></div>
  </div>

  <div class="instructions">
    <strong>How it works:</strong>
    <ol>
      <li>Enter your tasks in the sheet on the left (Project, Task, Category, Priority, Status, Notes).</li>
      <li>Date and Day are handled automatically.</li>
      <li>Fill as many rows as needed for this date.</li>
      <li>Click <strong>Save Tasks to Month Sheet</strong> below when finished.</li>
    </ol>
  </div>

  <div class="btn-container">
    <button id="saveBtn" class="btn btn-save" onclick="handleSave()">
      📥 Save Tasks to Month Sheet
    </button>
    <button id="cancelBtn" class="btn btn-cancel" onclick="handleCancel()">
      ✕ Cancel & Delete Sheet
    </button>
  </div>

  <div id="statusMsg" class="status-msg"></div>

  <script>
    const stagingSheetName = ${jsonStaging};
    const monthSheetName = ${jsonMonth};

    function handleSave() {
      const btn = document.getElementById('saveBtn');
      const cancelBtn = document.getElementById('cancelBtn');
      const msg = document.getElementById('statusMsg');

      btn.disabled = true;
      cancelBtn.disabled = true;
      btn.innerHTML = '⏳ Saving tasks...';
      msg.style.display = 'none';

      google.script.run
        .withSuccessHandler(function(res) {
          if (!res.success) {
            btn.disabled = false;
            cancelBtn.disabled = false;
            btn.innerHTML = '📥 Save Tasks to Month Sheet';
            msg.className = 'status-msg error';
            msg.innerText = res.message || 'Failed to save tasks.';
            return;
          }

          btn.innerHTML = '✓ Tasks Saved!';
          btn.style.backgroundColor = '#16a34a';
          btn.style.borderColor = '#15803d';
          msg.className = 'status-msg success';
          msg.innerText = 'Saved ' + res.count + ' task(s) to ' + res.monthSheetName + '! Staging sheet deleted.';

          setTimeout(function() {
            google.script.host.close();
          }, 1500);
        })
        .withFailureHandler(function(err) {
          btn.disabled = false;
          cancelBtn.disabled = false;
          btn.innerHTML = '📥 Save Tasks to Month Sheet';
          msg.className = 'status-msg error';
          msg.innerText = 'Error: ' + (err.message || err);
        })
        .saveTasksFromStagingSheet(stagingSheetName);
    }

    function handleCancel() {
      if (!confirm('Are you sure you want to discard this sheet and all entries on it?')) {
        return;
      }

      const cancelBtn = document.getElementById('cancelBtn');
      cancelBtn.disabled = true;
      cancelBtn.innerText = 'Deleting...';

      google.script.run
        .withSuccessHandler(function() {
          google.script.host.close();
        })
        .withFailureHandler(function(err) {
          alert('Error deleting sheet: ' + err.message);
        })
        .cancelTaskStagingSheet(stagingSheetName);
    }
  </script>
</body>
</html>`;
}