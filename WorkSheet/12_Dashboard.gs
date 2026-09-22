// ============================================================
// DASHBOARD.GS — TODO & MONTH STATUS OVERVIEW
// ============================================================

/**
 * Menu Action: (Re)builds the Dashboard sheet from scratch — the Todo list
 * and the current month sheet's task counts by status. The sheet is fully
 * derived data, so it is safe to wipe and rewrite on every refresh.
 */
function refreshDashboard() {
  const ss = getSpreadsheet();
  const sheet = ensureDashboardSheet(ss);

  sheet.clear();
  sheet.clearConditionalFormatRules();

  writeDashboardHeader(sheet);

  const todoEndRow = writeTodoSection(sheet, ss, 4);
  writeMonthStatusSection(sheet, ss, todoEndRow + 2);

  formatDashboardSheet(sheet);

  ss.setActiveSheet(sheet);

  SpreadsheetApp.getUi().alert('Dashboard refreshed.');
}


/**
 * Returns the Dashboard sheet, creating an empty one if missing.
 */
function ensureDashboardSheet(ss) {
  const sheet = ss.getSheetByName(CONFIG.DASHBOARD_SHEET_NAME);

  if (sheet) {
    return sheet;
  }

  return ss.insertSheet(CONFIG.DASHBOARD_SHEET_NAME);
}


/**
 * Writes the title and last-refreshed timestamp.
 */
function writeDashboardHeader(sheet) {
  const timezone = getTimezone();
  const timestamp = Utilities.formatDate(
    getToday(),
    timezone,
    'dd MMM yyyy, HH:mm'
  );

  sheet.getRange(1, 1)
    .setValue('📊 Work Tracker Dashboard')
    .setFontFamily('Arial')
    .setFontSize(14)
    .setFontWeight('bold');

  sheet.getRange(2, 1)
    .setValue(`Last refreshed: ${timestamp}`)
    .setFontFamily('Arial')
    .setFontStyle('italic')
    .setFontColor('#64748b');
}


/**
 * Writes the Todo section (Task Name, Project, Priority) starting at
 * startRow, sourced directly from the existing Todo sheet.
 *
 * @returns {number} The last row written (section end).
 */
function writeTodoSection(sheet, ss, startRow) {
  const rows = getTodoTaskRows(ss);
  const headers = CONFIG.TODO_HEADERS;

  sheet.getRange(startRow, 1, 1, headers.length)
    .merge()
    .setValue(`📝 Todo Tasks (${rows.length})`)
    .setFontFamily('Arial')
    .setFontWeight('bold')
    .setBackground(CONFIG.COLORS.HEADER);

  sheet.getRange(startRow + 1, 1, 1, headers.length)
    .setValues([headers])
    .setFontFamily('Arial')
    .setFontWeight('bold');

  if (rows.length > 0) {
    sheet.getRange(startRow + 2, 1, rows.length, headers.length)
      .setValues(rows);
  }

  return startRow + 1 + Math.max(rows.length, 1);
}


/**
 * Writes the current month sheet's task counts grouped by Status,
 * starting at startRow.
 *
 * @returns {number} The last row written (section end).
 */
function writeMonthStatusSection(sheet, ss, startRow) {
  const timezone = getTimezone();
  const monthSheetName = getMonthSheetName(getToday(), timezone);
  const monthSheet = ss.getSheetByName(monthSheetName);
  const counts = computeMonthStatusCounts(monthSheet);

  const statuses = CONFIG.LISTS.Statuses;
  const total = statuses.reduce((sum, status) => sum + counts[status], 0);

  sheet.getRange(startRow, 1, 1, 2)
    .merge()
    .setValue(`📅 ${monthSheetName} — Tasks by Status`)
    .setFontFamily('Arial')
    .setFontWeight('bold')
    .setBackground(CONFIG.COLORS.HEADER);

  sheet.getRange(startRow + 1, 1, 1, 2)
    .setValues([['Status', 'Count']])
    .setFontFamily('Arial')
    .setFontWeight('bold');

  const dataRows = statuses.map(status => [status, counts[status]]);
  dataRows.push(['Total', total]);

  sheet.getRange(startRow + 2, 1, dataRows.length, 2)
    .setValues(dataRows)
    .setFontFamily('Arial');

  sheet.getRange(startRow + 1 + dataRows.length, 1, 1, 2)
    .setFontWeight('bold');

  sheet.getRange(startRow + 2, 2, dataRows.length, 1)
    .setHorizontalAlignment('center');

  return startRow + 1 + dataRows.length;
}


/**
 * Returns non-empty Todo rows (Task Name, Project, Priority) from the
 * existing Todo sheet, or an empty array if it doesn't exist yet.
 */
function getTodoTaskRows(ss) {
  const sheet = ss.getSheetByName(CONFIG.TODO_SHEET_NAME);

  if (!sheet || sheet.getLastRow() < 2) {
    return [];
  }

  const values = sheet
    .getRange(2, 1, sheet.getLastRow() - 1, CONFIG.TODO_HEADERS.length)
    .getValues();

  return values.filter(row => String(row[0] || '').trim() !== '');
}


/**
 * Counts tasks in the given month sheet by Status, ignoring rows with
 * no Task description. Returns zero counts if the sheet is missing/empty.
 */
function computeMonthStatusCounts(monthSheet) {
  const counts = {};
  CONFIG.LISTS.Statuses.forEach(status => {
    counts[status] = 0;
  });

  if (!monthSheet || monthSheet.getLastRow() < 2) {
    return counts;
  }

  const taskCol = CONFIG.HEADERS.indexOf('Task');
  const statusCol = CONFIG.HEADERS.indexOf('Status');

  const rows = monthSheet
    .getRange(2, 1, monthSheet.getLastRow() - 1, CONFIG.HEADERS.length)
    .getValues();

  rows.forEach(row => {
    const task = String(row[taskCol] || '').trim();
    if (!task) {
      return;
    }

    const status = String(row[statusCol] || '').trim();
    if (Object.prototype.hasOwnProperty.call(counts, status)) {
      counts[status]++;
    }
  });

  return counts;
}


/**
 * Applies column widths, freezes the title rows, and trims the sheet
 * down to the columns actually used by the dashboard.
 */
function formatDashboardSheet(sheet) {
  sheet.setColumnWidth(1, 300);
  sheet.setColumnWidth(2, 130);
  sheet.setColumnWidth(3, 100);

  sheet.setFrozenRows(2);

  trimSheet(
    sheet,
    Math.max(sheet.getLastRow(), 2),
    CONFIG.TODO_HEADERS.length
  );
}
