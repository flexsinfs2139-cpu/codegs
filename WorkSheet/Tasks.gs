// ============================================================
// TASKS.GS — TASK STAGING & INSERTION ENGINE
// ============================================================

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
 * Asks the user for the day/date, then creates the staging sheet.
 */
function fillTaskForSelectedDate() {
  const ss = getSpreadsheet();
  const sheet = ss.getActiveSheet();
  const timezone = getTimezone();
  const ui = SpreadsheetApp.getUi();

  let suggestedDate = null;
  const activeRow = sheet.getActiveCell().getRow();

  if (activeRow > 1 && sheet.getName() !== CONFIG.LISTS_SHEET_NAME) {
    const cellValue = sheet.getRange(activeRow, 1).getValue();
    suggestedDate = parseDateFromCell(cellValue, sheet.getName(), timezone);
  }

  const defaultDateStr = suggestedDate
    ? Utilities.formatDate(suggestedDate, timezone, 'dd/MM/yyyy')
    : Utilities.formatDate(getToday(), timezone, 'dd/MM/yyyy');

  const response = ui.prompt(
    'Fill Task for Selected Date',
    `Enter day of current month (1–31) or date (DD/MM/YYYY) [Default: ${defaultDateStr}]:`,
    ui.ButtonSet.OK_CANCEL
  );

  if (response.getSelectedButton() !== ui.Button.OK) {
    return;
  }

  const input = response.getResponseText().trim() || defaultDateStr;
  const targetDate = resolveDateInput(input, sheet.getName(), timezone);

  if (!targetDate) {
    ui.alert(
      'Invalid Date',
      'Please enter a valid day of the month (1–31) or date (DD/MM/YYYY).',
      ui.ButtonSet.OK
    );
    return;
  }

  openTaskStagingSheet(targetDate);
}

/**
 * Creates and sets up the temporary "Date Month Year" staging sheet
 * and displays the task entry sidebar.
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

  // If staging sheet already exists, activate it and show sidebar
  let stagingSheet = ss.getSheetByName(stagingSheetName);
  if (stagingSheet) {
    ss.setActiveSheet(stagingSheet);
    showTaskStagingSidebar(targetDate, stagingSheetName, targetMonthSheetName);
    return;
  }

  // Create the staging sheet
  stagingSheet = ss.insertSheet(stagingSheetName);

  // 1. Setup Header
  stagingSheet
    .getRange(1, 1, 1, CONFIG.HEADERS.length)
    .setValues([CONFIG.HEADERS]);

  stagingSheet
    .getRange(1, 1, 1, CONFIG.HEADERS.length)
    .setBackground(CONFIG.COLORS.HEADER)
    .setFontFamily('Arial')
    .setFontWeight('bold')
    .setHorizontalAlignment('center');

  stagingSheet.setRowHeight(1, 28);

  // 2. Determine Date and Day labels for the selected date
  const dateLabel = Utilities.formatDate(targetDate, timezone, 'MMdd');
  const dayLabel = Utilities.formatDate(targetDate, timezone, 'EEE');

  // 3. Pre-fill initial 5 rows for user input
  const initialRows = [];
  for (let i = 0; i < 5; i++) {
    initialRows.push([
      dateLabel,
      dayLabel,
      '',
      '',
      '',
      'Medium',
      'Pending',
      ''
    ]);
  }

  stagingSheet
    .getRange(2, 1, initialRows.length, CONFIG.HEADERS.length)
    .setValues(initialRows);

  // 4. Formatting dimensions, fonts, and wrap strategy
  const widths = [75, 60, 130, 315, 125, 90, 125, 300];
  widths.forEach((w, idx) => {
    stagingSheet.setColumnWidth(idx + 1, w);
  });

  stagingSheet.setRowHeights(2, 15, 42);

  stagingSheet
    .getRange('A2:B')
    .setFontFamily('Roboto Mono')
    .setFontWeight('bold')
    .setHorizontalAlignment('center');

  stagingSheet.getRange('C2:C').setHorizontalAlignment('center');
  stagingSheet.getRange('D2:D').setHorizontalAlignment('left');
  stagingSheet.getRange('E2:G').setHorizontalAlignment('center');
  stagingSheet.getRange('H2:H').setHorizontalAlignment('left');

  stagingSheet
    .getRange(1, 1, 20, CONFIG.HEADERS.length)
    .setFontFamily('Arial')
    .setVerticalAlignment('middle')
    .setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);

  formatBorders(stagingSheet.getRange(1, 1, 15, CONFIG.HEADERS.length));

  // 5. Setup dropdown validations for rows 2 through 30
  setupDropdowns(stagingSheet);

  // 6. Freeze header & trim excess
  stagingSheet.setFrozenRows(1);
  trimSheet(stagingSheet, 20, CONFIG.HEADERS.length);

  // 7. Focus on first task cell (D2)
  ss.setActiveSheet(stagingSheet);
  stagingSheet.setActiveSelection('D2');

  // 8. Open the submission sidebar
  showTaskStagingSidebar(targetDate, stagingSheetName, targetMonthSheetName);
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
 * inserts additional rows with the same date if more than one task was entered,
 * and deletes the staging sheet.
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
      message: 'No tasks entered. Please write at least one task before saving.'
    };
  }

  // Read entered tasks
  const rawData = tempSheet
    .getRange(2, 1, lastRow - 1, CONFIG.HEADERS.length)
    .getValues();

  const enteredTasks = [];
  rawData.forEach(row => {
    const taskDesc = row[3] ? String(row[3]).trim() : '';
    if (taskDesc) {
      enteredTasks.push({
        dateLabel: row[0],
        dayLabel: row[1],
        project: row[2] || '',
        task: taskDesc,
        category: row[4] || '',
        priority: row[5] || 'Medium',
        status: row[6] || 'Pending',
        notes: row[7] || ''
      });
    }
  });

  if (enteredTasks.length === 0) {
    return {
      success: false,
      message: 'No tasks found. Please enter at least one task title in the "Task" column.'
    };
  }

  // Parse target date from sheet name (e.g. "08 September 2026")
  const targetDate = resolveDateInput(stagingSheetName, '', timezone) || getToday();
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
      // 1. Fill the first task into the existing row for this date
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

    // 2. If there are more tasks for this day, insert new row(s) after with the same date
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
 * Applies dropdown validation rules for an entire sheet.
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
 * Applies dropdown validation rules for a single newly inserted row.
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
      <li>Enter your tasks in the sheet on the left.</li>
      <li>Fill as many rows as needed for this date.</li>
      <li>Dropdowns are configured for Project, Category, Priority, and Status.</li>
      <li>Click <strong>Save Tasks</strong> below when finished.</li>
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