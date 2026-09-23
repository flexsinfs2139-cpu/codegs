// ============================================================
// DASHBOARD.GS — MODERN SAAS WORK & PROJECT COMMAND CENTER
// ============================================================

/**
 * Menu Action / Programmatic: (Re)builds the modern SaaS Command Center Dashboard from scratch.
 * Aggregates live task metrics across all monthly sheets into 6 high-level KPI cards,
 * a Monthly Breakdown table, a Project Performance table, and an Eisenhower Matrix
 * Todo Command Center with dedicated Todo KPI cards.
 * Uses Varela Round for all text/headers and Roboto Mono for all digits/metrics.
 *
 * @param {boolean} [suppressAlert=false] Whether to suppress the completion dialog
 * @param {boolean} [keepActiveSheet=false] Whether to preserve the current active sheet
 */
function refreshDashboard(suppressAlert = false, keepActiveSheet = false) {
  const ss = getSpreadsheet();
  const currentActiveSheet = keepActiveSheet ? ss.getActiveSheet() : null;
  const sheet = ensureDashboardSheet(ss);

  // Guarantee sufficient row & column headroom before drawing
  const minRows = 40;
  const currentMaxRows = sheet.getMaxRows();
  const currentMaxCols = sheet.getMaxColumns();

  if (currentMaxRows < minRows) {
    sheet.insertRowsAfter(currentMaxRows, minRows - currentMaxRows);
  }
  if (currentMaxCols < CONFIG.DASHBOARD.COLUMNS_COUNT) {
    sheet.insertColumnsAfter(currentMaxCols, CONFIG.DASHBOARD.COLUMNS_COUNT - currentMaxCols);
  }

  sheet.clear();
  sheet.clearConditionalFormatRules();

  const timezone = getTimezone();
  const timestamp = Utilities.formatDate(
    getToday(),
    timezone,
    'dd MMM yyyy, HH:mm'
  );

  const metrics = extractTaskMetrics(ss);

  renderDashboardHeader(sheet, timestamp);
  renderKpiCards(sheet, metrics.totals, metrics.todayTotals);

  const tablesEndRow = renderTablesSection(
    sheet,
    metrics.monthlyData,
    metrics.projectData,
    metrics.totals
  );

  const finalEndRow = renderTodoSection(sheet, ss, tablesEndRow + 1);

  formatDashboardSheet(sheet, finalEndRow);

  if (!keepActiveSheet) {
    ss.setActiveSheet(sheet);
  }

  if (!suppressAlert) {
    try {
      SpreadsheetApp.getUi().alert('Command Center Dashboard refreshed.');
    } catch (err) {
      ss.toast('Command Center Dashboard refreshed.', '⚡ Dashboard', 3);
    }
  }
}


/**
 * Automatically updates the Command Center Dashboard whenever any change occurs
 * across the spreadsheet (e.g. status changes, new tasks, priority updates, or todo edits).
 *
 * Can be triggered automatically by onEdit(e) / onChange(e), or called programmatically.
 *
 * @param {GoogleAppsScript.Events.SheetsOnEdit|GoogleAppsScript.Events.SheetsOnChange} [e]
 */
function updateDashboardOnChange(e) {
  const ss = getSpreadsheet();
  const dashboardSheet = ss.getSheetByName(CONFIG.DASHBOARD_SHEET_NAME);

  // If the Dashboard sheet hasn't been provisioned yet, skip
  if (!dashboardSheet) {
    return;
  }

  // If invoked via an edit event
  if (e && e.range) {
    const sheet = e.range.getSheet();
    const sheetName = sheet.getName();

    // 1. Edits occurring directly on the Dashboard sheet are ignored (read-only view)
    if (sheetName === CONFIG.DASHBOARD_SHEET_NAME) {
      return;
    }

    // 2. Edits occurring on Month sheets (e.g. SEP26), Todo sheet, or Lists sheet
    const isMonthSheet = /^[A-Z]{3}\d{2}$/.test(sheetName.trim().toUpperCase());
    const isTodoSheet = sheetName === CONFIG.TODO_SHEET_NAME;
    const isListsSheet = sheetName === CONFIG.LISTS_SHEET_NAME;

    if (!isMonthSheet && !isTodoSheet && !isListsSheet) {
      return;
    }
  }

  // Refresh dashboard silently in background without stealing active view
  refreshDashboard(true, true);
}


/**
 * Legacy stub: user edits directly on Dashboard are now ignored as Dashboard is a read-only analytics view.
 * All task interactions and checkbox selections happen directly on the dedicated Todo sheet.
 *
 * @param {GoogleAppsScript.Events.SheetsOnEdit} [e]
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} [ss]
 */
function handleDashboardTodoEdit(e, ss) {
  // Read-only dashboard view: tasks are managed on the Todo sheet
}


/**
 * Aliases for backwards compatibility and flexible trigger binding.
 */
const refreshDashboardOnChange = updateDashboardOnChange;
const updateDashboardWithChange = updateDashboardOnChange;


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
 * Renders the top title banner, system metadata status, and section label.
 * Rows 1–3.
 */
function renderDashboardHeader(sheet, timestamp) {
  const colors = CONFIG.DASHBOARD.COLORS;
  const textFont = CONFIG.FONTS.TEXT;

  // Row 1: Main Title Banner
  sheet.getRange(1, 1, 1, CONFIG.DASHBOARD.COLUMNS_COUNT)
    .merge()
    .setValue(CONFIG.DASHBOARD.TITLE)
    .setFontFamily(textFont)
    .setFontSize(14)
    .setFontWeight('bold')
    .setFontColor('#0f172a')
    .setVerticalAlignment('middle');
  sheet.setRowHeight(1, 38);

  // Row 2: Sub-Header Metadata / Status Bar
  const subheaderText = `● SYSTEM LIVE   |   Last Refreshed: ${timestamp}   |   ${CONFIG.DASHBOARD.VERSION}`;
  const subheaderRange = sheet.getRange(2, 1, 1, CONFIG.DASHBOARD.COLUMNS_COUNT);
  subheaderRange
    .merge()
    .setValue(subheaderText)
    .setFontFamily(textFont)
    .setFontSize(9)
    .setFontColor(colors.MUTED_TEXT)
    .setVerticalAlignment('middle');
  subheaderRange.setBorder(
    null, null, true, null, null, null,
    '#e2e8f0',
    SpreadsheetApp.BorderStyle.SOLID
  );
  sheet.setRowHeight(2, 22);

  // Row 3: Section Label
  sheet.getRange(3, 1)
    .setValue('KEY PERFORMANCE INDICATORS')
    .setFontFamily(textFont)
    .setFontSize(9)
    .setFontWeight('bold')
    .setFontColor('#475569')
    .setVerticalAlignment('middle');
  sheet.setRowHeight(3, 26);
}


/**
 * Renders the 6 Modern SaaS KPI Cards across Rows 4–6:
 * - Card 1 (A-B): TOTAL TASKS
 * - Card 2 (C-D): COMPLETED
 * - Card 3 (E-F): IN PROGRESS
 * - Card 4 (G-H): BLOCKED
 * - Card 5 (I-J): PENDING
 * - Card 6 (K-N): COMPLETION RATE
 *
 * Uses Varela Round for titles/subtitles and Roboto Mono for big digits.
 */
function renderKpiCards(sheet, totals, todayTotals = {}) {
  const colors = CONFIG.DASHBOARD.COLORS;
  const textFont = CONFIG.FONTS.TEXT;
  const digitsFont = CONFIG.FONTS.DIGITS;

  const tTotal = todayTotals.total || 0;
  const tDone = todayTotals.done || 0;
  const tActive = todayTotals.active || 0;
  const tBlocked = todayTotals.blocked || 0;
  const tPending = todayTotals.pending || 0;
  const tRate = todayTotals.rate || 0;

  const cardDefs = [
    {
      startCol: 1,
      numCols: 2,
      title: '• TOTAL TASKS',
      value: String(totals.total),
      sub: `${tTotal} logged today`,
      bg: colors.CARD_TOTAL_BG,
      text: colors.CARD_TOTAL_TEXT,
      border: colors.CARD_TOTAL_BORDER
    },
    {
      startCol: 3,
      numCols: 2,
      title: '✓ COMPLETED',
      value: String(totals.done),
      sub: `${tDone} completed today`,
      bg: colors.CARD_DONE_BG,
      text: colors.CARD_DONE_TEXT,
      border: colors.CARD_DONE_BORDER
    },
    {
      startCol: 5,
      numCols: 2,
      title: '⏳ IN PROGRESS',
      value: String(totals.active),
      sub: `${tActive} active today`,
      bg: colors.CARD_ACTIVE_BG,
      text: colors.CARD_ACTIVE_TEXT,
      border: colors.CARD_ACTIVE_BORDER
    },
    {
      startCol: 7,
      numCols: 2,
      title: '⛔ BLOCKED',
      value: String(totals.blocked),
      sub: `${tBlocked} blocked today`,
      bg: colors.CARD_BLOCKED_BG,
      text: colors.CARD_BLOCKED_TEXT,
      border: colors.CARD_BLOCKED_BORDER
    },
    {
      startCol: 9,
      numCols: 2,
      title: '📋 PENDING',
      value: String(totals.pending),
      sub: `${tPending} pending today`,
      bg: colors.CARD_PENDING_BG,
      text: colors.CARD_PENDING_TEXT,
      border: colors.CARD_PENDING_BORDER
    },
    {
      startCol: 11,
      numCols: 4,
      title: '📈 COMPLETION RATE',
      value: `${totals.rate}%`,
      sub: `${tRate}% today's rate`,
      bg: colors.CARD_RATE_BG,
      text: colors.CARD_RATE_TEXT,
      border: colors.CARD_RATE_BORDER
    }
  ];

  cardDefs.forEach(card => {
    // 1. Background over rows 4 to 6
    const cardRange = sheet.getRange(4, card.startCol, 3, card.numCols);
    cardRange.setBackground(card.bg);

    // 2. Row 4: Title (Varela Round)
    const titleRange = sheet.getRange(4, card.startCol, 1, card.numCols);
    titleRange
      .merge()
      .setValue(card.title)
      .setFontFamily(textFont)
      .setFontSize(9)
      .setFontWeight('bold')
      .setFontColor(card.text)
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle');

    // 3. Row 5: Large Metric (Roboto Mono)
    const valRange = sheet.getRange(5, card.startCol, 1, card.numCols);
    valRange
      .merge()
      .setValue(card.value)
      .setFontFamily(digitsFont)
      .setFontSize(22)
      .setFontWeight('bold')
      .setFontColor(card.text)
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle');

    // 4. Row 6: Subtitle / Description (Varela Round)
    const subRange = sheet.getRange(6, card.startCol, 1, card.numCols);
    subRange
      .merge()
      .setValue(card.sub)
      .setFontFamily(textFont)
      .setFontSize(8)
      .setFontColor(colors.MUTED_TEXT)
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle');

    // 5. Borders: Outer thin border + thick colored accent line at bottom
    cardRange.setBorder(
      true, true, false, true, false, false,
      '#e2e8f0',
      SpreadsheetApp.BorderStyle.SOLID
    );
    subRange.setBorder(
      null, null, true, null, null, null,
      card.border,
      SpreadsheetApp.BorderStyle.SOLID_THICK
    );
  });

  sheet.setRowHeight(4, 26);
  sheet.setRowHeight(5, 46);
  sheet.setRowHeight(6, 24);

  // Row 7: Blank Spacer
  sheet.setRowHeight(7, 16);
}


/**
 * Renders the two side-by-side tables:
 * - Left (Cols A-G): MONTHLY BREAKDOWN
 * - Col H: Spacer
 * - Right (Cols I-M): PROJECT PERFORMANCE
 *
 * Uses Varela Round for names/headers and Roboto Mono for digits/metrics.
 *
 * @returns {number} The last row index used for the tables section.
 */
function renderTablesSection(sheet, monthlyData, projectData, totals) {
  const colors = CONFIG.DASHBOARD.COLORS;
  const textFont = CONFIG.FONTS.TEXT;
  const digitsFont = CONFIG.FONTS.DIGITS;

  // Row 8: Section Banners (Varela Round)
  sheet.getRange(8, 1, 1, 7)
    .merge()
    .setValue('  MONTHLY BREAKDOWN')
    .setFontFamily(textFont)
    .setFontSize(9)
    .setFontWeight('bold')
    .setFontColor(colors.BANNER_TEXT)
    .setBackground(colors.BANNER_BG)
    .setVerticalAlignment('middle');

  sheet.getRange(8, 9, 1, 6)
    .merge()
    .setValue('  PROJECT PERFORMANCE')
    .setFontFamily(textFont)
    .setFontSize(9)
    .setFontWeight('bold')
    .setFontColor(colors.BANNER_TEXT)
    .setBackground(colors.BANNER_BG)
    .setVerticalAlignment('middle');

  sheet.setRowHeight(8, 26);

  // Row 9: Table Column Headers (Varela Round)
  const leftHeaders = ['Month', 'Total', 'Done', 'Active', 'Blocked', 'Pending', 'Progress'];
  sheet.getRange(9, 1, 1, 7)
    .setValues([leftHeaders])
    .setFontFamily(textFont)
    .setFontSize(9)
    .setFontWeight('bold')
    .setFontColor(colors.TABLE_HEADER_TEXT)
    .setBackground(colors.TABLE_HEADER_BG)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setBorder(
      null, null, true, null, null, null,
      colors.BORDER,
      SpreadsheetApp.BorderStyle.SOLID
    );

  const rightHeaders = ['Project', 'Total', 'Completed', 'In Progress', 'Blocked', 'Completion %'];
  sheet.getRange(9, 9, 1, 6)
    .setValues([rightHeaders])
    .setFontFamily(textFont)
    .setFontSize(9)
    .setFontWeight('bold')
    .setFontColor(colors.TABLE_HEADER_TEXT)
    .setBackground(colors.TABLE_HEADER_BG)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setBorder(
      null, null, true, null, null, null,
      colors.BORDER,
      SpreadsheetApp.BorderStyle.SOLID
    );

  sheet.setRowHeight(9, 26);

  // Rows 10+: Table Data
  const numMonthRows = Math.max(monthlyData.length, 1);
  const numProjRows = Math.max(projectData.length, 1);
  const dataRowCount = Math.max(numMonthRows, numProjRows);

  // Populate Left Table (Monthly)
  if (monthlyData.length > 0) {
    const leftRows = monthlyData.map(m => [
      m.displayName,
      m.total,
      m.done,
      m.active,
      m.blocked,
      m.pending,
      `${m.rate}%`
    ]);

    sheet.getRange(10, 1, leftRows.length, 7)
      .setValues(leftRows)
      .setFontSize(9)
      .setVerticalAlignment('middle');

    // Month Name (Col A): Varela Round, bold, left-aligned
    sheet.getRange(10, 1, leftRows.length, 1)
      .setFontFamily(textFont)
      .setFontWeight('bold')
      .setHorizontalAlignment('left');

    // Metrics (Cols B-G): Roboto Mono, centered
    const monthMetricsRange = sheet.getRange(10, 2, leftRows.length, 6);
    monthMetricsRange
      .setFontFamily(digitsFont)
      .setHorizontalAlignment('center');

    // Done Column (Col C): bold green accent
    sheet.getRange(10, 3, leftRows.length, 1)
      .setFontWeight('bold')
      .setFontColor(colors.GREEN_ACCENT);

    // Subtle borders on data rows
    sheet.getRange(10, 1, leftRows.length, 7).setBorder(
      null, true, true, true, true, true,
      '#e2e8f0',
      SpreadsheetApp.BorderStyle.SOLID
    );
  } else {
    sheet.getRange(10, 1, 1, 7)
      .setValues([['No data', 0, 0, 0, 0, 0, '0%']])
      .setFontFamily(textFont)
      .setFontSize(9)
      .setHorizontalAlignment('center');
  }

  // Populate Right Table (Projects)
  if (projectData.length > 0) {
    const rightRows = projectData.map(p => [
      p.project,
      p.tasks,
      p.done,
      p.active,
      p.blocked,
      `${p.rate}%`
    ]);

    sheet.getRange(10, 9, rightRows.length, 6)
      .setValues(rightRows)
      .setFontSize(9)
      .setVerticalAlignment('middle');

    // Project Name (Col I): Varela Round, bold, left-aligned
    sheet.getRange(10, 9, rightRows.length, 1)
      .setFontFamily(textFont)
      .setFontWeight('bold')
      .setHorizontalAlignment('left');

    // Project Metrics (Cols J-N): Roboto Mono, centered
    const projMetricsRange = sheet.getRange(10, 10, rightRows.length, 5);
    projMetricsRange
      .setFontFamily(digitsFont)
      .setHorizontalAlignment('center');

    // Completed Column (Col K / Col 11): bold green accent
    sheet.getRange(10, 11, rightRows.length, 1)
      .setFontWeight('bold')
      .setFontColor(colors.GREEN_ACCENT);

    // Completion % Column (Col N / Col 14): bold green accent
    sheet.getRange(10, 14, rightRows.length, 1)
      .setFontWeight('bold')
      .setFontColor(colors.GREEN_ACCENT);

    // Subtle borders on data rows
    sheet.getRange(10, 9, rightRows.length, 6).setBorder(
      null, true, true, true, true, true,
      '#e2e8f0',
      SpreadsheetApp.BorderStyle.SOLID
    );
  } else {
    sheet.getRange(10, 9, 1, 6)
      .setValues([['No data', 0, 0, 0, 0, '0%']])
      .setFontFamily(textFont)
      .setFontSize(9)
      .setHorizontalAlignment('center');
  }

  for (let r = 0; r < dataRowCount; r++) {
    sheet.setRowHeight(10 + r, 24);
  }

  // Bottom Summary Row
  const totalRowIndex = 10 + dataRowCount;

  // Left Summary (TOTAL / AVG)
  const leftSummary = [
    'TOTAL / AVG',
    totals.total,
    totals.done,
    totals.active,
    totals.blocked,
    totals.pending,
    `${totals.rate}%`
  ];
  const leftSummaryRange = sheet.getRange(totalRowIndex, 1, 1, 7);
  leftSummaryRange
    .setValues([leftSummary])
    .setFontSize(9)
    .setFontWeight('bold')
    .setFontColor(colors.TABLE_HEADER_TEXT)
    .setVerticalAlignment('middle');

  // Label: Varela Round
  sheet.getRange(totalRowIndex, 1)
    .setFontFamily(textFont)
    .setHorizontalAlignment('left');

  // Digits: Roboto Mono
  sheet.getRange(totalRowIndex, 2, 1, 6)
    .setFontFamily(digitsFont)
    .setHorizontalAlignment('center');

  sheet.getRange(totalRowIndex, 3).setFontColor(colors.GREEN_ACCENT);

  leftSummaryRange.setBorder(
    true, null, true, null, null, null,
    colors.BORDER,
    SpreadsheetApp.BorderStyle.SOLID
  );

  // Right Summary (TOTAL)
  const rightSummary = [
    'TOTAL',
    totals.total,
    totals.done,
    totals.active,
    totals.blocked,
    `${totals.rate}%`
  ];
  const rightSummaryRange = sheet.getRange(totalRowIndex, 9, 1, 6);
  rightSummaryRange
    .setValues([rightSummary])
    .setFontSize(9)
    .setFontWeight('bold')
    .setFontColor(colors.TABLE_HEADER_TEXT)
    .setVerticalAlignment('middle');

  // Label: Varela Round
  sheet.getRange(totalRowIndex, 9)
    .setFontFamily(textFont)
    .setHorizontalAlignment('left');

  // Digits: Roboto Mono
  sheet.getRange(totalRowIndex, 10, 1, 5)
    .setFontFamily(digitsFont)
    .setHorizontalAlignment('center');

  sheet.getRange(totalRowIndex, 11).setFontColor(colors.GREEN_ACCENT);
  sheet.getRange(totalRowIndex, 14).setFontColor(colors.GREEN_ACCENT);

  rightSummaryRange.setBorder(
    true, null, true, null, null, null,
    colors.BORDER,
    SpreadsheetApp.BorderStyle.SOLID
  );

  sheet.setRowHeight(totalRowIndex, 26);

  return totalRowIndex;
}


/**
 * Renders the Eisenhower Matrix Todo statistics cards in the Dashboard sheet.
 *
 * Uses Varela Round for titles/headers/names and Roboto Mono for digits/metrics.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
 * @param {number} startRow
 * @returns {number} The last row index written.
 */
function renderTodoSection(sheet, ss, startRow) {
  const rows = getTodoTaskRows(ss);
  const colors = CONFIG.DASHBOARD.COLORS;
  const textFont = CONFIG.FONTS.TEXT;
  const digitsFont = CONFIG.FONTS.DIGITS;

  const totalTodos = rows.length;
  const q1Count = rows.filter(r => r[2] === true).length;
  const q2Count = rows.filter(r => r[3] === true).length;
  const q3Count = rows.filter(r => r[4] === true).length;
  const q4Count = rows.filter(r => r[5] === true).length;
  const focusRatio = totalTodos > 0
    ? Math.round(((q1Count + q2Count) / totalTodos) * 100)
    : 0;

  // 1. Spacer Row
  sheet.setRowHeight(startRow, 18);

  // 2. Section Label (Varela Round)
  const labelRow = startRow + 1;
  sheet.getRange(labelRow, 1)
    .setValue('EISENHOWER MATRIX — TODO COMMAND')
    .setFontFamily(textFont)
    .setFontSize(9)
    .setFontWeight('bold')
    .setFontColor('#475569')
    .setVerticalAlignment('middle');
  sheet.setRowHeight(labelRow, 26);

  // 3. Six Todo KPI Cards across Rows (labelRow + 1) to (labelRow + 3)
  const cardStartRow = labelRow + 1;
  const todoCards = [
    {
      startCol: 1,
      numCols: 2,
      title: '• TOTAL TODOS',
      value: String(totalTodos),
      sub: 'Open backlog',
      bg: colors.CARD_PENDING_BG,
      text: colors.CARD_PENDING_TEXT,
      border: colors.CARD_PENDING_BORDER
    },
    {
      startCol: 3,
      numCols: 2,
      title: '🔴 Q1: DO',
      value: String(q1Count),
      sub: 'Urgent & Important',
      bg: colors.CARD_BLOCKED_BG,
      text: colors.CARD_BLOCKED_TEXT,
      border: colors.CARD_BLOCKED_BORDER
    },
    {
      startCol: 5,
      numCols: 2,
      title: '🔵 Q2: SCHEDULE',
      value: String(q2Count),
      sub: 'Important / Plan',
      bg: colors.CARD_TOTAL_BG,
      text: colors.CARD_TOTAL_TEXT,
      border: colors.CARD_TOTAL_BORDER
    },
    {
      startCol: 7,
      numCols: 2,
      title: '🟡 Q3: DELEGATE',
      value: String(q3Count),
      sub: 'Urgent / Delegate',
      bg: colors.CARD_ACTIVE_BG,
      text: colors.CARD_ACTIVE_TEXT,
      border: colors.CARD_ACTIVE_BORDER
    },
    {
      startCol: 9,
      numCols: 2,
      title: '⚪ Q4: DON\'T DO',
      value: String(q4Count),
      sub: 'Drop / eliminate',
      bg: '#f1f5f9',
      text: '#475569',
      border: '#64748b'
    },
    {
      startCol: 11,
      numCols: 4,
      title: '🎯 FOCUS RATIO',
      value: `${focusRatio}%`,
      sub: 'Q1 + Q2 share',
      bg: colors.CARD_DONE_BG,
      text: colors.CARD_DONE_TEXT,
      border: colors.CARD_DONE_BORDER
    }
  ];

  todoCards.forEach(card => {
    // Background across 3 card rows
    const cardRange = sheet.getRange(cardStartRow, card.startCol, 3, card.numCols);
    cardRange.setBackground(card.bg);

    // Title (Row 1 of card: Varela Round)
    const titleRange = sheet.getRange(cardStartRow, card.startCol, 1, card.numCols);
    titleRange
      .merge()
      .setValue(card.title)
      .setFontFamily(textFont)
      .setFontSize(9)
      .setFontWeight('bold')
      .setFontColor(card.text)
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle');

    // Value (Row 2 of card: Roboto Mono for digits)
    const valRange = sheet.getRange(cardStartRow + 1, card.startCol, 1, card.numCols);
    valRange
      .merge()
      .setValue(card.value)
      .setFontFamily(digitsFont)
      .setFontSize(22)
      .setFontWeight('bold')
      .setFontColor(card.text)
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle');

    // Subtitle (Row 3 of card: Varela Round)
    const subRange = sheet.getRange(cardStartRow + 2, card.startCol, 1, card.numCols);
    subRange
      .merge()
      .setValue(card.sub)
      .setFontFamily(textFont)
      .setFontSize(8)
      .setFontColor(colors.MUTED_TEXT)
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle');

    // Borders
    cardRange.setBorder(
      true, true, false, true, false, false,
      '#e2e8f0',
      SpreadsheetApp.BorderStyle.SOLID
    );
    subRange.setBorder(
      null, null, true, null, null, null,
      card.border,
      SpreadsheetApp.BorderStyle.SOLID_THICK
    );
  });

  sheet.setRowHeight(cardStartRow, 26);
  sheet.setRowHeight(cardStartRow + 1, 46);
  sheet.setRowHeight(cardStartRow + 2, 24);

  return cardStartRow + 2;
}


/**
 * Aggregates task data across all month sheets found in the spreadsheet.
 */
function extractTaskMetrics(ss) {
  const monthSheets = getAllMonthSheets(ss);
  const dateCol = CONFIG.HEADERS.indexOf('Date');
  const taskCol = CONFIG.HEADERS.indexOf('Task');
  const projectCol = CONFIG.HEADERS.indexOf('Project');
  const statusCol = CONFIG.HEADERS.indexOf('Status');

  const timezone = getTimezone();
  const today = getToday();
  const currentMonthSheetName = Utilities.formatDate(today, timezone, 'MMMyy').toUpperCase();
  const todayDateStr = Utilities.formatDate(today, timezone, CONFIG.DATE_FORMAT || 'MMdd');

  const monthlyData = [];
  const projectMap = {};

  (CONFIG.LISTS.Projects || []).forEach(proj => {
    projectMap[proj] = { tasks: 0, done: 0, active: 0, blocked: 0, pending: 0 };
  });

  let grandTotal = 0;
  let grandDone = 0;
  let grandActive = 0;
  let grandBlocked = 0;
  let grandPending = 0;

  let todayTotal = 0;
  let todayDone = 0;
  let todayActive = 0;
  let todayBlocked = 0;
  let todayPending = 0;

  monthSheets.forEach(sheet => {
    const sheetName = sheet.getName();
    const displayName = formatMonthDisplayName(sheetName);
    const isCurrentMonth = sheetName.trim().toUpperCase() === currentMonthSheetName;
    const lastRow = sheet.getLastRow();

    let mTotal = 0;
    let mDone = 0;
    let mActive = 0;
    let mBlocked = 0;
    let mPending = 0;

    if (lastRow > 1) {
      const rows = sheet
        .getRange(2, 1, lastRow - 1, CONFIG.HEADERS.length)
        .getValues();

      rows.forEach(row => {
        const task = String(row[taskCol] || '').trim();
        if (!task) return;

        const rowDate = dateCol >= 0 ? String(row[dateCol] || '').trim() : '';
        const project = String(row[projectCol] || '').trim();
        const status = String(row[statusCol] || '').trim();

        mTotal++;
        if (status === 'Completed') {
          mDone++;
        } else if (status === 'In Progress') {
          mActive++;
        } else if (status === 'Blocked') {
          mBlocked++;
        } else {
          mPending++;
        }

        // Today metrics extraction
        if (isCurrentMonth && (rowDate === todayDateStr || rowDate.endsWith(todayDateStr))) {
          todayTotal++;
          if (status === 'Completed') {
            todayDone++;
          } else if (status === 'In Progress') {
            todayActive++;
          } else if (status === 'Blocked') {
            todayBlocked++;
          } else {
            todayPending++;
          }
        }

        if (project) {
          if (!projectMap[project]) {
            projectMap[project] = { tasks: 0, done: 0, active: 0, blocked: 0, pending: 0 };
          }
          projectMap[project].tasks++;
          if (status === 'Completed') {
            projectMap[project].done++;
          } else if (status === 'In Progress') {
            projectMap[project].active++;
          } else if (status === 'Blocked') {
            projectMap[project].blocked++;
          } else {
            projectMap[project].pending++;
          }
        }
      });
    }

    const mRate = mTotal > 0 ? Math.round((mDone / mTotal) * 100) : 0;
    monthlyData.push({
      displayName,
      sheetName,
      total: mTotal,
      done: mDone,
      active: mActive,
      blocked: mBlocked,
      pending: mPending,
      rate: mRate
    });

    grandTotal += mTotal;
    grandDone += mDone;
    grandActive += mActive;
    grandBlocked += mBlocked;
    grandPending += mPending;
  });

  const grandRate = grandTotal > 0 ? Math.round((grandDone / grandTotal) * 100) : 0;
  const todayRate = todayTotal > 0 ? Math.round((todayDone / todayTotal) * 100) : 0;

  const projectData = [];
  Object.keys(projectMap).forEach(proj => {
    const data = projectMap[proj];
    if (data.tasks > 0) {
      const rate = Math.round((data.done / data.tasks) * 100);
      projectData.push({
        project: proj,
        tasks: data.tasks,
        done: data.done,
        active: data.active,
        blocked: data.blocked,
        pending: data.pending,
        rate: rate
      });
    }
  });

  // Deterministic sort: tasks count descending, then alphabetically by project name
  projectData.sort((a, b) => b.tasks - a.tasks || a.project.localeCompare(b.project));

  return {
    monthlyData,
    projectData,
    totals: {
      total: grandTotal,
      done: grandDone,
      active: grandActive,
      blocked: grandBlocked,
      pending: grandPending,
      rate: grandRate
    },
    todayTotals: {
      total: todayTotal,
      done: todayDone,
      active: todayActive,
      blocked: todayBlocked,
      pending: todayPending,
      rate: todayRate
    }
  };
}


/**
 * Returns all month sheets in chronological/tab order matching the MMMyy pattern.
 */
function getAllMonthSheets(ss) {
  const monthSheets = ss.getSheets().filter(s => {
    const name = s.getName().trim().toUpperCase();
    return /^[A-Z]{3}\d{2}$/.test(name);
  });

  if (monthSheets.length > 0) {
    return monthSheets;
  }

  const timezone = getTimezone();
  const currentMonthName = getMonthSheetName(getToday(), timezone);
  const currentSheet = ss.getSheetByName(currentMonthName);
  return currentSheet ? [currentSheet] : [];
}


/**
 * Formats a sheet name like "SEP26" into a clean display title like "Sept 26".
 */
function formatMonthDisplayName(sheetName) {
  const match = sheetName.match(/^([A-Za-z]{3})(\d{2})$/);
  if (!match) return sheetName;

  const monthCode = match[1].toUpperCase();
  const yearCode = match[2];

  const monthMap = {
    JAN: 'Jan',
    FEB: 'Feb',
    MAR: 'Mar',
    APR: 'Apr',
    MAY: 'May',
    JUN: 'Jun',
    JUL: 'Jul',
    AUG: 'Aug',
    SEP: 'Sept',
    OCT: 'Oct',
    NOV: 'Nov',
    DEC: 'Dec'
  };

  const monthName = monthMap[monthCode] || monthCode;
  return `${monthName} ${yearCode}`;
}


/**
 * Returns non-empty Todo rows (Task Name, Project, Q1–Q4) from the
 * existing Todo sheet, or an empty array if it doesn't exist yet.
 */
function getTodoTaskRows(ss) {
  const sheet = ss.getSheetByName(CONFIG.TODO_SHEET_NAME);

  if (!sheet || sheet.getLastRow() < CONFIG.TODO_FIRST_DATA_ROW) {
    return [];
  }

  const numRows = sheet.getLastRow() - CONFIG.TODO_HEADER_ROW;
  if (numRows <= 0) {
    return [];
  }

  const values = sheet
    .getRange(CONFIG.TODO_FIRST_DATA_ROW, 1, numRows, CONFIG.TODO_HEADERS.length)
    .getValues();

  return values.filter(row => String(row[0] || '').trim() !== '');
}


/**
 * Applies column widths, freezes the title rows, and trims the sheet
 * down to exactly the 14 columns used by the modern SaaS dashboard command center.
 */
function formatDashboardSheet(sheet, lastRow) {
  const widths = (CONFIG.DIMENSIONS && CONFIG.DIMENSIONS.DASHBOARD_COL_WIDTHS) || [
    130, // Col 1 (A): Month
    75,  // Col 2 (B): Total
    75,  // Col 3 (C): Done
    75,  // Col 4 (D): Active
    75,  // Col 5 (E): Blocked
    75,  // Col 6 (F): Pending
    85,  // Col 7 (G): Progress
    35,  // Col 8 (H): Spacer Gap
    140, // Col 9 (I): Project
    75,  // Col 10 (J): Total
    75,  // Col 11 (K): Completed
    75,  // Col 12 (L): In Progress
    75,  // Col 13 (M): Blocked
    85   // Col 14 (N): Completion %
  ];

  widths.forEach((width, index) => {
    sheet.setColumnWidth(index + 1, width);
  });

  sheet.setFrozenRows(2);
  setSheetGridlinesHidden(sheet, true);

  const finalRows = Math.max(lastRow, 14);
  trimSheet(sheet, finalRows, CONFIG.DASHBOARD.COLUMNS_COUNT);
}
