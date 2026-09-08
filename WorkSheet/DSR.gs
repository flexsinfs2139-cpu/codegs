// ============================================================
// DSR.GS — DAILY STATUS REPORT ENGINE
// ============================================================

/**
 * Generates and displays the Daily Status Report for today.
 */
function generateDSRForToday() {
  const timezone = getTimezone();
  const today = getToday();
  showDSRDialog(today, true);
}

/**
 * Generates and displays the Daily Status Report for a selected date in the month.
 * Automatically suggests date if a row in the month sheet is highlighted,
 * or prompts the user to enter/confirm the date.
 */
function generateDSRForSelectedDate() {
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
    'Generate DSR for Selected Date',
    `Enter date (DD/MM/YYYY or day 1–31) [Default: ${defaultDateStr}]:`,
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
      'Please enter a valid date (DD/MM/YYYY) or day of the month (1–31).',
      ui.ButtonSet.OK
    );
    return;
  }

  showDSRDialog(targetDate, false);
}

/**
 * Compiles DSR data and presents the interactive HTML modal dialog.
 *
 * @param {Date} targetDate - The target date for the report
 * @param {boolean} isToday - Whether this is today's report
 */
function showDSRDialog(targetDate, isToday) {
  const dsrData = compileDSRData(targetDate, isToday);
  const htmlContent = getDSRDialogHtml(dsrData, isToday);

  const htmlOutput = HtmlService.createHtmlOutput(htmlContent)
    .setWidth(720)
    .setHeight(560);

  const dialogTitle = isToday
    ? 'Daily Status Report — Today'
    : `Daily Status Report — ${dsrData.formattedDate}`;

  SpreadsheetApp.getUi().showModalDialog(htmlOutput, dialogTitle);
}

/**
 * Extracts and groups tasks for the given target date from the workbook.
 *
 * @param {Date} targetDate
 * @param {boolean} isToday
 * @returns {Object} Compiled DSR data
 */
function compileDSRData(targetDate, isToday) {
  const ss = getSpreadsheet();
  const timezone = getTimezone();
  const formattedDate = Utilities.formatDate(targetDate, timezone, 'dd/MM/yyyy');

  // Locate the target month sheet (e.g. "SEP26")
  const expectedSheetName = Utilities.formatDate(
    targetDate,
    timezone,
    'MMMyy'
  ).toUpperCase();

  let sheet = ss.getSheetByName(expectedSheetName);

  // Fallback to active sheet if expected sheet does not exist
  if (!sheet) {
    const active = ss.getActiveSheet();
    if (active.getName() !== CONFIG.LISTS_SHEET_NAME) {
      sheet = active;
    }
  }

  const projectGroups = {};
  let totalTasks = 0;

  if (sheet && sheet.getLastRow() > 1) {
    const dataRange = sheet.getRange(
      2,
      1,
      sheet.getLastRow() - 1,
      CONFIG.HEADERS.length
    );
    const rows = dataRange.getValues();

    rows.forEach(row => {
      const cellDate = row[0];
      if (!isMatchingDate(cellDate, targetDate, timezone)) {
        return;
      }

      const rawProject = row[2] ? String(row[2]).trim() : '';
      const rawTask = row[3] ? String(row[3]).trim() : '';
      const rawCategory = row[4] ? String(row[4]).trim() : '';
      const rawPriority = row[5] ? String(row[5]).trim() : '';
      const rawStatus = row[6] ? String(row[6]).trim() : '';
      const rawNotes = row[7] ? String(row[7]).trim() : '';

      // Ignore rows with completely empty task description
      if (!rawTask) {
        return;
      }

      totalTasks++;
      const projectName = rawProject || 'General';

      if (!projectGroups[projectName]) {
        projectGroups[projectName] = {
          completed: [],
          inProgress: [],
          blocked: [],
          plan: []
        };
      }

      const itemDesc = rawCategory ? `${rawTask} (${rawCategory})` : rawTask;
      const statusLower = rawStatus.toLowerCase();

      if (statusLower === 'completed') {
        projectGroups[projectName].completed.push(itemDesc);
      } else if (statusLower === 'in progress') {
        projectGroups[projectName].inProgress.push(itemDesc);
      } else if (statusLower === 'blocked') {
        const blockerDetail = rawNotes ? `${itemDesc} — ${rawNotes}` : itemDesc;
        projectGroups[projectName].blocked.push(blockerDetail);
      } else if (
        statusLower === 'pending' ||
        statusLower === 'hold' ||
        statusLower === 'not started'
      ) {
        projectGroups[projectName].plan.push(itemDesc);
      } else {
        // Any other category/status
        projectGroups[projectName].inProgress.push(itemDesc);
      }
    });
  }

  const projectNames = Object.keys(projectGroups);
  const totalProjects = projectNames.length;

  let reportText = '';

  if (totalProjects === 0) {
    reportText = [
      `Project Name: General`,
      `Date: ${formattedDate}`,
      ``,
      `Tasks Completed:`,
      `- None`,
      ``,
      `Work in Progress:`,
      `- None`,
      ``,
      `Blockers / Issues:`,
      `- None`,
      ``,
      `Plan for Next Working Day:`,
      `- None`
    ].join('\n');
  } else {
    const blocks = projectNames.map(proj => {
      const g = projectGroups[proj];

      const completedLines = g.completed.length
        ? g.completed.map(t => `- ${t}`).join('\n')
        : '- None';

      const inProgressLines = g.inProgress.length
        ? g.inProgress.map(t => `- ${t}`).join('\n')
        : '- None';

      const blockedLines = g.blocked.length
        ? g.blocked.map(t => `- ${t}`).join('\n')
        : '- None';

      const planLines = g.plan.length
        ? g.plan.map(t => `- ${t}`).join('\n')
        : '- None';

      return [
        `Project Name: ${proj}`,
        `Date: ${formattedDate}`,
        ``,
        `Tasks Completed:`,
        completedLines,
        ``,
        `Work in Progress:`,
        inProgressLines,
        ``,
        `Blockers / Issues:`,
        blockedLines,
        ``,
        `Plan for Next Working Day:`,
        planLines
      ].join('\n');
    });

    reportText = blocks.join('\n\n');
  }

  return {
    formattedDate: formattedDate,
    isToday: isToday,
    totalTasks: totalTasks,
    totalProjects: totalProjects,
    projectNames: projectNames.join(', ') || 'General',
    reportText: reportText
  };
}



/**
 * Saves DSR history to a dedicated "DSR" sheet in the spreadsheet.
 */
function saveDSRToSheet(payload) {
  const ss = getSpreadsheet();
  let sheet = ss.getSheetByName('DSR');

  if (!sheet) {
    sheet = ss.insertSheet('DSR');
    const headers = [
      'Saved At',
      'DSR Date',
      'Projects',
      'Total Tasks',
      'Full Report'
    ];

    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

    sheet
      .getRange(1, 1, 1, headers.length)
      .setBackground(CONFIG.COLORS.HEADER)
      .setFontFamily('Arial')
      .setFontWeight('bold')
      .setHorizontalAlignment('center');

    sheet.setRowHeight(1, 28);
    sheet.setColumnWidth(1, 160);
    sheet.setColumnWidth(2, 100);
    sheet.setColumnWidth(3, 160);
    sheet.setColumnWidth(4, 90);
    sheet.setColumnWidth(5, 450);
    sheet.setFrozenRows(1);
  }

  const timezone = getTimezone();
  const savedAt = Utilities.formatDate(
    new Date(),
    timezone,
    'yyyy-MM-dd HH:mm:ss'
  );
  const nextRow = sheet.getLastRow() + 1;

  sheet.getRange(nextRow, 1, 1, 5).setValues([[
    savedAt,
    payload.formattedDate,
    payload.projectNames || 'General',
    payload.totalTasks || 0,
    payload.reportText || ''
  ]]);

  sheet.setRowHeight(nextRow, 42);
  sheet.getRange(nextRow, 1, 1, 5).setVerticalAlignment('middle');

  return { success: true, sheetName: 'DSR', row: nextRow };
}

/**
 * Returns HTML markup for the modal dialog.
 * Styled to match the requested design and screenshot.
 */
function getDSRDialogHtml(dsrData, isToday) {
  const badgeText = isToday ? 'Today' : dsrData.formattedDate;
  const subtitle = `Date: ${dsrData.formattedDate} • ${dsrData.totalTasks} task(s) across ${dsrData.totalProjects} project(s)`;

  // Safely JSON-encode for inline script
  const jsonReportText = JSON.stringify(dsrData.reportText).replace(/<\/script/gi, '<\\/script');
  const jsonFormattedDate = JSON.stringify(dsrData.formattedDate).replace(/<\/script/gi, '<\\/script');
  const jsonPayload = JSON.stringify(dsrData).replace(/<\/script/gi, '<\\/script');

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
      font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, "Helvetica Neue", Arial, sans-serif;
      background-color: #f8fafc;
      color: #1e293b;
      padding: 16px 20px;
      font-size: 14px;
      -webkit-font-smoothing: antialiased;
    }

    .dsr-container {
      display: flex;
      flex-direction: column;
      height: 100%;
    }

    .dsr-header {
      display: flex;
      flex-direction: column;
      gap: 4px;
      padding-bottom: 12px;
      border-bottom: 1px solid #e2e8f0;
    }

    .dsr-title-row {
      display: flex;
      align-items: center;
      gap: 8px;
    }

    .dsr-icon {
      font-size: 17px;
      line-height: 1;
    }

    .dsr-title {
      font-size: 16px;
      font-weight: 700;
      color: #0f172a;
      letter-spacing: -0.01em;
    }

    .dsr-badge {
      display: inline-flex;
      align-items: center;
      background-color: #ede9fe;
      color: #6366f1;
      font-size: 11px;
      font-weight: 600;
      padding: 2px 9px;
      border-radius: 9999px;
      line-height: 1.4;
    }

    .dsr-subtitle {
      font-size: 13px;
      color: #64748b;
      font-weight: 400;
      margin-top: 2px;
    }

    .dsr-card {
      background: #ffffff;
      border: 1px solid #e2e8f0;
      border-radius: 10px;
      padding: 16px 20px;
      margin-top: 14px;
      flex: 1;
      box-shadow: 0 1px 2px rgba(0, 0, 0, 0.03);
    }

    .dsr-text-box {
      font-family: ui-monospace, SFMono-Regular, Menlo, Monaco, Consolas, "Liberation Mono", "Courier New", monospace;
      font-size: 13px;
      line-height: 1.6;
      color: #1e293b;
      white-space: pre-wrap;
      word-break: break-word;
      max-height: 310px;
      overflow-y: auto;
      padding-right: 8px;
    }

    /* Custom Scrollbar */
    .dsr-text-box::-webkit-scrollbar {
      width: 6px;
    }
    .dsr-text-box::-webkit-scrollbar-track {
      background: #f1f5f9;
      border-radius: 4px;
    }
    .dsr-text-box::-webkit-scrollbar-thumb {
      background: #cbd5e1;
      border-radius: 4px;
    }
    .dsr-text-box::-webkit-scrollbar-thumb:hover {
      background: #94a3b8;
    }

    .dsr-footer {
      display: flex;
      justify-content: space-between;
      align-items: center;
      margin-top: 16px;
      padding-top: 4px;
    }

    .dsr-footer-left {
      display: flex;
      align-items: center;
    }

    .dsr-footer-right {
      display: flex;
      align-items: center;
      gap: 8px;
    }

    .btn {
      display: inline-flex;
      align-items: center;
      justify-content: center;
      gap: 6px;
      font-size: 13px;
      font-weight: 500;
      padding: 8px 16px;
      border-radius: 7px;
      cursor: pointer;
      transition: all 0.15s ease-in-out;
      outline: none;
      user-select: none;
      font-family: inherit;
    }

    .btn-primary {
      background-color: #5850ec;
      color: #ffffff;
      border: 1px solid #4f46e5;
      font-weight: 600;
      box-shadow: 0 1px 2px rgba(88, 80, 236, 0.2);
    }

    .btn-primary:hover {
      background-color: #4f46e5;
    }

    .btn-primary:active {
      transform: translateY(1px);
    }

    .btn-secondary {
      background-color: #ffffff;
      color: #334155;
      border: 1px solid #e2e8f0;
      box-shadow: 0 1px 2px rgba(0, 0, 0, 0.04);
    }

    .btn-secondary:hover {
      background-color: #f8fafc;
      border-color: #cbd5e1;
      color: #0f172a;
    }

    .btn-secondary:disabled {
      opacity: 0.6;
      cursor: not-allowed;
    }

    .btn-close {
      background: transparent;
      color: #64748b;
      border: none;
      padding: 8px 12px;
    }

    .btn-close:hover {
      color: #0f172a;
    }
  </style>
</head>
<body>
  <div class="dsr-container">
    <!-- Header -->
    <div class="dsr-header">
      <div class="dsr-title-row">
        <span class="dsr-icon">📋</span>
        <span class="dsr-title">Daily Status Report</span>
        <span class="dsr-badge">${badgeText}</span>
      </div>
      <div class="dsr-subtitle">${subtitle}</div>
    </div>

    <!-- Preformatted Report Box -->
    <div class="dsr-card">
      <div class="dsr-text-box" id="dsrText">${escapeHtml(dsrData.reportText)}</div>
    </div>

    <!-- Actions Bar -->
    <div class="dsr-footer">
      <div class="dsr-footer-left">
        <button id="copyBtn" class="btn btn-primary" onclick="copyToClipboard()">
          📋 Copy to Clipboard
        </button>
      </div>

      <div class="dsr-footer-right">
        <button id="saveBtn" class="btn btn-secondary" onclick="saveToSheet()">
          📄 Save to Sheet
        </button>
        <button id="downloadBtn" class="btn btn-secondary" onclick="downloadTxt()">
          📥 Download .txt
        </button>
        <button class="btn btn-close" onclick="google.script.host.close()">
          Close
        </button>
      </div>
    </div>
  </div>

  <script>
    const reportText = ${jsonReportText};
    const formattedDate = ${jsonFormattedDate};
    const dsrPayload = ${jsonPayload};

    function copyToClipboard() {
      const text = document.getElementById('dsrText').innerText;
      if (navigator.clipboard && navigator.clipboard.writeText) {
        navigator.clipboard.writeText(text).then(showCopied).catch(fallbackCopy);
      } else {
        fallbackCopy();
      }
    }

    function fallbackCopy() {
      const text = document.getElementById('dsrText').innerText;
      const ta = document.createElement('textarea');
      ta.value = text;
      ta.style.position = 'fixed';
      ta.style.opacity = '0';
      document.body.appendChild(ta);
      ta.select();
      try {
        document.execCommand('copy');
        showCopied();
      } catch (err) {
        alert('Failed to copy. Please manually copy from the box.');
      }
      document.body.removeChild(ta);
    }

    function showCopied() {
      const btn = document.getElementById('copyBtn');
      const prevHtml = btn.innerHTML;
      btn.innerHTML = '✓ Copied!';
      btn.style.backgroundColor = '#16a34a';
      btn.style.borderColor = '#15803d';

      setTimeout(function() {
        btn.innerHTML = prevHtml;
        btn.style.backgroundColor = '';
        btn.style.borderColor = '';
      }, 2000);
    }

    function downloadTxt() {
      const text = document.getElementById('dsrText').innerText;
      const cleanDate = formattedDate.replace(/[\/\\:]/g, '-');
      const filename = 'DSR_' + cleanDate + '.txt';

      const blob = new Blob([text], { type: 'text/plain;charset=utf-8' });
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = filename;
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      URL.revokeObjectURL(url);
    }

    function saveToSheet() {
      const btn = document.getElementById('saveBtn');
      btn.disabled = true;
      btn.innerHTML = '⏳ Saving...';

      google.script.run
        .withSuccessHandler(function(res) {
          btn.innerHTML = '✓ Saved to Sheet!';
          btn.style.color = '#16a34a';
          btn.style.borderColor = '#86efac';

          setTimeout(function() {
            btn.disabled = false;
            btn.innerHTML = '📄 Save to Sheet';
            btn.style.color = '';
            btn.style.borderColor = '';
          }, 3000);
        })
        .withFailureHandler(function(err) {
          btn.disabled = false;
          btn.innerHTML = '❌ Error';
          alert('Could not save DSR: ' + (err.message || err));
        })
        .saveDSRToSheet(dsrPayload);
    }
  </script>
</body>
</html>`;
}

/**
 * Escapes HTML entities to prevent script injection in the dialog.
 */
function escapeHtml(text) {
  if (!text) return '';
  return text
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#039;');
}
