// ============================================================
// SAMPLE_DATA.GS — COMPREHENSIVE DUMMY DATA SEED ENGINE
// ============================================================

/**
 * Menu Action / Programmatic: Seeds realistic sample data into both
 * the current Month sheet and the Todo sheet, and refreshes the Dashboard.
 *
 * Populates:
 * - 21 diverse tasks in the Month sheet across various dates, projects,
 *   categories, priorities, and statuses (Completed, In Progress, Blocked, Pending).
 * - 12 prioritized tasks in the Todo sheet mapped to the 4 Eisenhower
 *   Matrix quadrants (Q1: Do, Q2: Schedule, Q3: Delegate, Q4: Don't Do).
 *
 * @param {boolean} [suppressAlert=false] Whether to suppress the completion alert dialog.
 */
function populateDummyData(suppressAlert = false) {
  const ss = getSpreadsheet();
  const timezone = getTimezone();
  const today = getToday();

  // 1. Ensure Lists sheet exists so validation and ranges are healthy
  ensureListsSheet(ss);

  // 2. Populate Todo sheet
  populateTodoDummyData(ss);

  // 3. Populate Month sheet
  populateMonthDummyData(ss, today, timezone);

  // 4. Refresh Dashboard silently
  refreshDashboard(true, true);

  if (!suppressAlert) {
    try {
      SpreadsheetApp.getUi().alert(
        'Dummy Data Seeded Successfully!\n\n' +
        '• Month Sheet: 21 realistic engineering tasks across projects and statuses.\n' +
        '• Todo Sheet: 12 Eisenhower Matrix tasks with live counter stats.\n' +
        '• Command Center Dashboard: Refreshed with live KPIs and Todo Backlog.'
      );
    } catch (err) {
      ss.toast('Dummy data populated and Dashboard refreshed.', '⚡ Sample Data', 4);
    }
  }
}

/**
 * Alias for populateDummyData.
 */
function addDummyData(suppressAlert) {
  populateDummyData(suppressAlert);
}


/**
 * Populates the Todo sheet with 12 structured tasks across the 4 Eisenhower quadrants.
 *
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
 */
function populateTodoDummyData(ss) {
  // Ensure Todo sheet exists with proper structure & formatting
  const sheet = ensureTodoSheet(ss);
  setupTodoStructure(sheet);

  // Clear any existing data rows from row 3 downwards
  const lastRow = sheet.getLastRow();
  if (lastRow >= CONFIG.TODO_FIRST_DATA_ROW) {
    sheet.getRange(
      CONFIG.TODO_FIRST_DATA_ROW,
      1,
      lastRow - CONFIG.TODO_HEADER_ROW,
      CONFIG.TODO_HEADERS.length
    ).clearContent().clearFormat();
  }

  // 12 Realistic Eisenhower Matrix tasks
  const sampleTodos = [
    // Q1: Do (Urgent & Important — Critical Blockers & Security)
    ['Resolve production database latency spike during peak traffic', 'Workbench', true, false, false, false],
    ['Hotfix critical session leakage in OAuth2 callback handler', 'Clinkio', true, false, false, false],
    ['Restore broken CI/CD build artifact deployment pipeline', 'Surfari', true, false, false, false],

    // Q2: Schedule (Important, Not Urgent — Architecture, Planning & Quality)
    ['Architect microservices migration plan for payment processing', 'Memryx', false, true, false, false],
    ['Design cohesive UI design system tokens in Figma and code', 'DroidLens', false, true, false, false],
    ['Implement end-to-end integration test coverage for core flows', 'Workbench', false, true, false, false],
    ['Conduct quarterly security penetration audit and dependency scan', 'Clinkio', false, true, false, false],

    // Q3: Delegate (Urgent, Not Important — Operations & Reporting)
    ['Compile weekly sprint velocity metrics and burn-down chart', 'Workbench', false, false, true, false],
    ['Update OpenAPI / Swagger reference documentation for v2.4', 'Memryx', false, false, true, false],
    ['Triage and categorize incoming community bug reports', 'Surfari', false, false, true, false],

    // Q4: Don't Do (Not Urgent, Not Important — Eliminate / Postpone)
    ['Audit deprecated third-party charting libraries for replacement', 'Other', false, false, false, true],
    ['Archive obsolete project Slack channels and documentation drafts', 'Other', false, false, false, true]
  ];

  const numRows = sampleTodos.length;

  // Insert data rows
  sheet
    .getRange(CONFIG.TODO_FIRST_DATA_ROW, 1, numRows, CONFIG.TODO_HEADERS.length)
    .setValues(sampleTodos);

  // Set font and styling
  const textFont = CONFIG.FONTS.TEXT;
  sheet
    .getRange(CONFIG.TODO_FIRST_DATA_ROW, 1, numRows, 2)
    .setFontFamily(textFont)
    .setFontSize(10)
    .setVerticalAlignment('middle')
    .setHorizontalAlignment('left');

  for (let r = 0; r < numRows; r++) {
    sheet.setRowHeight(CONFIG.TODO_FIRST_DATA_ROW + r, 28);
  }

  // Apply column widths, dropdowns, checkboxes, conditional formatting, and trimming
  formatTodoSheet(sheet);
  setupTodoDropdowns(sheet);
  setupTodoCheckboxes(sheet);
  setupTodoConditionalFormatting(sheet);
}


/**
 * Populates the current Month sheet with 21 realistic engineering tasks
 * distributed across days, projects, categories, priorities, and statuses.
 *
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
 * @param {Date} today
 * @param {string} timezone
 */
function populateMonthDummyData(ss, today, timezone) {
  const year = today.getFullYear();
  const month = today.getMonth();
  const sheetName = Utilities.formatDate(today, timezone, 'MMMyy').toUpperCase();

  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  } else {
    sheet.clear();
    sheet.clearConditionalFormatRules();
  }

  const daysInMonth = new Date(year, month + 1, 0).getDate();

  // 21 Rich sample tasks mapped to days in month
  const sampleTasksByDay = {
    2: [
      { project: 'Workbench', task: 'Design database schema for multi-tenant workspace isolation', category: 'Development', priority: 'High', status: 'Completed' },
      { project: 'Workbench', task: 'Configure PostgreSQL connection pooling and read replicas', category: 'API', priority: 'High', status: 'Completed' }
    ],
    3: [
      { project: 'Clinkio', task: 'Implement OAuth2 Google and GitHub single sign-on flows', category: 'Development', priority: 'Urgent', status: 'Completed' },
      { project: 'Clinkio', task: 'Audit session expiration and JWT token refresh vulnerabilities', category: 'Bug Fix', priority: 'High', status: 'Completed' }
    ],
    4: [
      { project: 'Surfari', task: 'Optimize asset compression and lazy-loading for landing page', category: 'UI/UX', priority: 'Medium', status: 'Completed' },
      { project: 'Surfari', task: 'End-to-end checkout flow validation across mobile viewports', category: 'Testing', priority: 'Medium', status: 'Completed' }
    ],
    7: [
      { project: 'Memryx', task: 'Implement vector embeddings cache for semantic search query engine', category: 'Development', priority: 'High', status: 'Completed' },
      { project: 'Memryx', task: 'Benchmark memory consumption under sustained read load', category: 'Research', priority: 'Medium', status: 'Completed' }
    ],
    8: [
      { project: 'DroidLens', task: 'Develop real-time camera preview frame processing pipeline', category: 'Development', priority: 'Urgent', status: 'Completed' },
      { project: 'DroidLens', task: 'Fix buffer overflow during high-fps video capture stream', category: 'Bug Fix', priority: 'Urgent', status: 'Completed' }
    ],
    10: [
      { project: 'Workbench', task: 'Build responsive Kanban board with drag-and-drop task cards', category: 'UI/UX', priority: 'High', status: 'In Progress' },
      { project: 'Workbench', task: 'Sync state transitions with WebSocket event listeners', category: 'API', priority: 'High', status: 'In Progress' }
    ],
    11: [
      { project: 'Clinkio', task: 'Stripe webhook listener for subscription tier upgrades & downgrades', category: 'Development', priority: 'High', status: 'In Progress' }
    ],
    14: [
      { project: 'Surfari', task: 'Resolve Safari iOS visual glitch on sticky navigation bar', category: 'UI/UX', priority: 'Medium', status: 'In Progress' },
      { project: 'Surfari', task: 'Investigate third-party CDN latency spikes during peak Asian hours', category: 'Research', priority: 'Medium', status: 'Blocked' }
    ],
    16: [
      { project: 'Memryx', task: 'Deploy staging cluster on AWS ECS with auto-scaling triggers', category: 'Deployment', priority: 'Urgent', status: 'Blocked' }
    ],
    18: [
      { project: 'DroidLens', task: 'Prepare v1.2 beta release notes and test flight build distribution', category: 'Meeting', priority: 'Low', status: 'Pending' },
      { project: 'DroidLens', task: 'Write automated integration tests for ML model edge inference', category: 'Testing', priority: 'Medium', status: 'Pending' }
    ],
    21: [
      { project: 'Workbench', task: 'Refactor state management store to eliminate redundant re-renders', category: 'Development', priority: 'Medium', status: 'Pending' }
    ],
    22: [
      { project: 'Clinkio', task: 'Set up Grafana alerts for API 5xx error rate thresholds', category: 'Deployment', priority: 'High', status: 'Pending' }
    ],
    23: [
      { project: 'Other', task: 'Quarterly infrastructure cost optimization and idle resource cleanup', category: 'Other', priority: 'Low', status: 'Pending' }
    ]
  };

  // Set headers
  sheet.getRange(CONFIG.HEADER_ROW, 1, 1, CONFIG.HEADERS.length).setValues([CONFIG.HEADERS]);

  const rows = [];
  for (let day = 1; day <= daysInMonth; day++) {
    const date = new Date(year, month, day);
    const dateLabel = Utilities.formatDate(date, timezone, CONFIG.DATE_FORMAT || 'MMdd');
    const dayLabel = Utilities.formatDate(date, timezone, 'EEE');

    const dayTasks = sampleTasksByDay[day];
    if (dayTasks && dayTasks.length > 0) {
      dayTasks.forEach(t => {
        rows.push([
          dateLabel,
          dayLabel,
          t.project,
          t.task,
          t.category,
          t.priority,
          t.status
        ]);
      });
    } else {
      // Default single empty template row for that date
      rows.push([
        dateLabel,
        dayLabel,
        '',
        '',
        '',
        '',
        'Pending'
      ]);
    }
  }

  // Batch insert all rows
  sheet.getRange(2, 1, rows.length, CONFIG.HEADERS.length).setValues(rows);

  // Apply layout, dropdowns, conditional formatting, and dimensions
  trimSheet(sheet, rows.length + 1, CONFIG.HEADERS.length);
  formatWorkTracker(sheet);
  setupDropdowns(sheet);
  setupConditionalFormatting(sheet);
  setupWeekendFormatting(sheet);

  sheet.setFrozenRows(1);
}
