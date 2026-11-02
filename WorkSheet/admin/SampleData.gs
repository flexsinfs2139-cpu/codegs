// ============================================================
// SAMPLEDATA.GS — 1-CLICK DUMMY DATA SEEDING ENGINE
// ============================================================

/**
 * Menu Action: Seeds the entire workbook with realistic engineering tasks.
 * Populates:
 * - 12 prioritized backlog tasks in the Todo sheet with Task IDs, Due Dates, and Notes.
 * - 21 rich daily tasks in the current Month sheet across categories, priorities, and statuses.
 * - Initial timeline events and decisions in the Decisions sheet.
 * - Refreshes the Command Center Dashboard with live metrics.
 *
 * @param {boolean} [suppressAlert=false] Whether to suppress the completion alert dialog.
 */
function populateDummyData(suppressAlert = false) {
  const ss = getSpreadsheet();
  const timezone = getTimezone();
  const today = getToday();

  // 1. Ensure master Lists sheet
  ensureListsSheet(ss);

  // 2. Populate Todo Backlog (Simplified 5-column layout)
  populateTodoDummyData(ss);

  // 3. Populate Current Month Sheet (8-column layout)
  populateMonthDummyData(ss, today, timezone);

  // 4. Populate Decisions Sheet with baseline events
  populateDecisionsDummyData(ss);

  // 5. Refresh Command Center Dashboard
  refreshDashboard(true, true);

  if (!suppressAlert) {
    SpreadsheetApp.getUi().alert(
      'Sample Data Populated Successfully!\n\n' +
        '• Todo Sheet: 12 backlog tasks with permanent Task IDs (TSK-0001..12), due dates, and notes.\n' +
        '• Month Sheet: 21 daily engineering tasks linked with Task IDs.\n' +
        '• Decisions Sheet: Chronological audit trail with decisions, blockers, and status changes.\n' +
        '• Command Center Dashboard: Refreshed with live KPIs and backlog schedule metrics.'
    );
  }
}


/**
 * Populates the Todo sheet with 12 structured tasks across active projects.
 * Schema: [ID, Task Description, Project, Due Date, Notes]
 *
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
 */
function populateTodoDummyData(ss) {
  const sheet = ensureTodoSheet(ss);
  setupTodoStructure(sheet);

  // Clear existing rows
  const lastRow = sheet.getLastRow();
  if (lastRow >= CONFIG.TODO_FIRST_DATA_ROW) {
    sheet.getRange(
      CONFIG.TODO_FIRST_DATA_ROW,
      1,
      lastRow - CONFIG.TODO_HEADER_ROW,
      CONFIG.TODO_HEADERS.length
    ).clearContent().clearFormat();
  }

  // 12 Realistic engineering backlog tasks
  const sampleTodos = [
    ['TSK-0001', 'Resolve production database latency spike during peak traffic', 'Workbench', '2026-11-04', 'Investigate connection pooling and read replicas'],
    ['TSK-0002', 'Hotfix critical session leakage in OAuth2 callback handler', 'Clinkio', '2026-11-03', 'Security audit patch for token expiration'],
    ['TSK-0003', 'Restore broken CI/CD build artifact deployment pipeline', 'Surfari', '2026-11-05', 'GitHub Actions runner disk space exhaustion'],
    ['TSK-0004', 'Architect microservices migration plan for payment processing', 'Memryx', '2026-11-14', 'Split monolith Stripe integrations into standalone worker'],
    ['TSK-0005', 'Design cohesive UI design system tokens in Figma and code', 'DroidLens', '2026-11-18', 'Align palette with Varela Round and Roboto Mono'],
    ['TSK-0006', 'Implement end-to-end integration test coverage for core flows', 'Workbench', '2026-11-20', 'Cypress test suites for billing checkout'],
    ['TSK-0007', 'Conduct quarterly security penetration audit and dependency scan', 'Clinkio', '2026-11-25', 'Snyk and Dependabot vulnerability triage'],
    ['TSK-0008', 'Compile weekly sprint velocity metrics and burn-down chart', 'Workbench', '2026-11-06', 'Executive stakeholder operational summary'],
    ['TSK-0009', 'Update OpenAPI / Swagger reference documentation for v2.4', 'Memryx', '2026-11-12', 'Document all subscription webhook payloads'],
    ['TSK-0010', 'Triage and categorize incoming community bug reports', 'Surfari', '2026-11-08', 'Filter duplicate client issues'],
    ['TSK-0011', 'Audit deprecated third-party charting libraries for replacement', 'Other', '2026-11-30', 'Evaluate lightweight SVG canvas alternatives'],
    ['TSK-0012', 'Archive obsolete project Slack channels and documentation drafts', 'Other', '', 'Internal hygiene and knowledgebase cleanup']
  ];

  const numRows = sampleTodos.length;
  ensureSheetCapacity(sheet, CONFIG.TODO_FIRST_DATA_ROW + numRows + 5, CONFIG.TODO_HEADERS.length);

  sheet
    .getRange(CONFIG.TODO_FIRST_DATA_ROW, 1, numRows, CONFIG.TODO_HEADERS.length)
    .setValues(sampleTodos);

  formatTodoSheet(sheet);
  setupTodoDropdowns(sheet);
  updateTodoTaskIdsNamedRange(ss);

  // Initialize Task ID counter in ScriptProperties to 12
  PropertiesService.getScriptProperties().setProperty('TASK_ID_COUNTER', '12');
}


/**
 * Populates the current Month sheet with 21 realistic engineering tasks.
 * Schema: [Date, Day, Task ID, Project, Task, Category, Priority, Status]
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

  const sampleTasksByDay = {
    2: [
      { taskId: 'TSK-0001', project: 'Workbench', task: 'Design database schema for multi-tenant workspace isolation', category: 'Development', priority: 'High', status: 'Completed' },
      { taskId: '', project: 'Workbench', task: 'Configure PostgreSQL connection pooling and read replicas', category: 'API', priority: 'High', status: 'Completed' }
    ],
    3: [
      { taskId: 'TSK-0002', project: 'Clinkio', task: 'Implement OAuth2 Google and GitHub single sign-on flows', category: 'Development', priority: 'Urgent', status: 'Completed' },
      { taskId: '', project: 'Clinkio', task: 'Audit session expiration and JWT token refresh vulnerabilities', category: 'Bug Fix', priority: 'High', status: 'Completed' }
    ],
    4: [
      { taskId: 'TSK-0003', project: 'Surfari', task: 'Optimize asset compression and lazy-loading for landing page', category: 'UI/UX', priority: 'Medium', status: 'Completed' },
      { taskId: '', project: 'Surfari', task: 'End-to-end checkout flow validation across mobile viewports', category: 'Testing', priority: 'Medium', status: 'Completed' }
    ],
    7: [
      { taskId: 'TSK-0004', project: 'Memryx', task: 'Implement vector embeddings cache for semantic search query engine', category: 'Development', priority: 'High', status: 'Completed' },
      { taskId: '', project: 'Memryx', task: 'Benchmark memory consumption under sustained read load', category: 'Research', priority: 'Medium', status: 'Completed' }
    ],
    8: [
      { taskId: 'TSK-0005', project: 'DroidLens', task: 'Develop real-time camera preview frame processing pipeline', category: 'Development', priority: 'Urgent', status: 'Completed' },
      { taskId: '', project: 'DroidLens', task: 'Fix buffer overflow during high-fps video capture stream', category: 'Bug Fix', priority: 'Urgent', status: 'Completed' }
    ],
    10: [
      { taskId: 'TSK-0006', project: 'Workbench', task: 'Build responsive Kanban board with drag-and-drop task cards', category: 'UI/UX', priority: 'High', status: 'In Progress' },
      { taskId: '', project: 'Workbench', task: 'Sync state transitions with WebSocket event listeners', category: 'API', priority: 'High', status: 'In Progress' }
    ],
    11: [
      { taskId: 'TSK-0007', project: 'Clinkio', task: 'Stripe webhook listener for subscription tier upgrades & downgrades', category: 'Development', priority: 'High', status: 'In Progress' }
    ],
    14: [
      { taskId: '', project: 'Surfari', task: 'Resolve Safari iOS visual glitch on sticky navigation bar', category: 'UI/UX', priority: 'Medium', status: 'In Progress' },
      { taskId: 'TSK-0008', project: 'Surfari', task: 'Investigate third-party CDN latency spikes during peak Asian hours', category: 'Research', priority: 'Medium', status: 'Blocked' }
    ],
    16: [
      { taskId: 'TSK-0009', project: 'Memryx', task: 'Deploy staging cluster on AWS ECS with auto-scaling triggers', category: 'Deployment', priority: 'Urgent', status: 'Blocked' }
    ],
    18: [
      { taskId: 'TSK-0010', project: 'DroidLens', task: 'Prepare v1.2 beta release notes and test flight build distribution', category: 'Meeting', priority: 'Low', status: 'Pending' },
      { taskId: '', project: 'DroidLens', task: 'Write automated integration tests for ML model edge inference', category: 'Testing', priority: 'Medium', status: 'Pending' }
    ],
    21: [
      { taskId: 'TSK-0011', project: 'Workbench', task: 'Refactor state management store to eliminate redundant re-renders', category: 'Development', priority: 'Medium', status: 'Pending' }
    ],
    22: [
      { taskId: '', project: 'Clinkio', task: 'Set up Grafana alerts for API 5xx error rate thresholds', category: 'Deployment', priority: 'High', status: 'Pending' }
    ],
    23: [
      { taskId: 'TSK-0012', project: 'Other', task: 'Quarterly infrastructure cost optimization and idle resource cleanup', category: 'Other', priority: 'Low', status: 'Pending' }
    ]
  };

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
          t.taskId || '',
          t.project,
          t.task,
          t.category,
          t.priority,
          t.status
        ]);
      });
    } else {
      rows.push([
        dateLabel,
        dayLabel,
        '',
        '',
        '',
        '',
        '',
        'Pending'
      ]);
    }
  }

  ensureSheetCapacity(sheet, rows.length + 5, CONFIG.HEADERS.length);
  sheet.getRange(2, 1, rows.length, CONFIG.HEADERS.length).setValues(rows);

  trimSheet(sheet, rows.length + 1, CONFIG.HEADERS.length);
  formatWorkTracker(sheet);
  setupDropdowns(sheet);
  setupConditionalFormatting(sheet);
  setupWeekendFormatting(sheet);
  sheet.setFrozenRows(1);
  setSheetGridlinesHidden(sheet, true);
}


/**
 * Populates the Decisions sheet with baseline sample decisions, blockers, and events.
 *
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
 */
function populateDecisionsDummyData(ss) {
  const sheet = ensureDecisionsSheet(ss);

  // Clear existing data rows
  const lastRow = sheet.getLastRow();
  if (lastRow >= CONFIG.DECISIONS_FIRST_DATA_ROW) {
    sheet.getRange(
      CONFIG.DECISIONS_FIRST_DATA_ROW,
      1,
      lastRow - CONFIG.DECISIONS_HEADER_ROW,
      CONFIG.DECISIONS_HEADERS.length
    ).clearContent().clearFormat();
  }

  const sampleDecisions = [
    ['EVT-0001', 'TSK-0001', '2026-11-02 09:30', 'Task Created', '-', 'Pending', 'Task created in Todo backlog', 'Initial feature breakdown', 'User', 'Schema design', '2026-11-04', 'ARCH-101'],
    ['EVT-0002', 'TSK-0001', '2026-11-02 10:15', 'Status Changed', 'Pending', 'In Progress', 'Began schema design for isolation', 'Assigned sprint priority', 'Lead Dev', 'Verify PostgreSQL constraints', '', ''],
    ['EVT-0003', 'TSK-0001', '2026-11-02 16:45', 'Decision Made', 'In Progress', 'In Progress', 'Adopt schema-per-tenant rather than database-per-tenant', 'Lower connection footprint and simplified cloud backup', 'Architect', 'Implement Alembic migration script', '2026-11-03', 'ADR-04'],
    ['EVT-0004', 'TSK-0001', '2026-11-02 18:30', 'Task Completed', 'In Progress', 'Completed', 'Schema isolation and read replicas configured', 'All CI tests passed with 100% compliance', 'Lead Dev', '-', '', 'PR #108'],
    ['EVT-0005', 'TSK-0008', '2026-11-04 11:00', 'Blocker Identified', 'In Progress', 'Blocked', 'Third-party CDN latency spikes during peak Asian hours', 'Cloudflare edge node routing failure', 'DevOps Lead', 'Open urgent ticket with Cloudflare enterprise support', '2026-11-05', 'TICKET-9921'],
    ['EVT-0006', 'TSK-0009', '2026-11-05 14:20', 'Blocker Identified', 'Pending', 'Blocked', 'AWS IAM permissions missing for staging ECS task definition', 'Policy restricted by security compliance baseline', 'Backend Dev', 'Request IAM role update from SecOps', '2026-11-06', 'SEC-402']
  ];

  const numRows = sampleDecisions.length;
  ensureSheetCapacity(sheet, CONFIG.DECISIONS_FIRST_DATA_ROW + numRows + 5, CONFIG.DECISIONS_HEADERS.length);
  sheet.getRange(CONFIG.DECISIONS_FIRST_DATA_ROW, 1, numRows, CONFIG.DECISIONS_HEADERS.length).setValues(sampleDecisions);
  formatDecisionsSheet(sheet);

  // Set Timeline counter to 6
  PropertiesService.getScriptProperties().setProperty('TIMELINE_ID_COUNTER', '6');
}
