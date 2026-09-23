const CONFIG = {
  HEADER_ROW: 1,
  DATE_FORMAT: 'MMdd',
  DEFAULT_TASKS_PER_DAY: 1,
  LISTS_SHEET_NAME: 'Lists',
  TODO_SHEET_NAME: 'Todo',
  TODO_DROPDOWN_ROWS: 200,
  TODO_STAT_ROW: 1,
  TODO_HEADER_ROW: 2,
  TODO_FIRST_DATA_ROW: 3,
  DASHBOARD_SHEET_NAME: 'Dashboard',

  TODO_HEADERS: [
    'Task Name',
    'Project',
    'Q1: Do',
    'Q2: Schedule',
    'Q3: Delegate',
    'Q4: Don\'t Do'
  ],

  HEADERS: [
    'Date',
    'Day',
    'Project',
    'Task',
    'Category',
    'Priority',
    'Status'
  ],

  COLORS: {
    HEADER: '#d9ead3',
    BORDER: '#d9d9d9',
    STATS_BG: '#f8f9fa',
    STATS_BORDER: '#e2e8f0',
    Q1_BG: '#fce8e6',
    Q1_TEXT: '#c5221f',
    Q2_BG: '#e8f0fe',
    Q2_TEXT: '#1a73e8',
    Q3_BG: '#fef7e0',
    Q3_TEXT: '#b06000',
    Q4_BG: '#f1f3f4',
    Q4_TEXT: '#5f6368'
  },

  LISTS: {
    Projects: [
      'Clinkio',
      'Surfari',
      'Workbench',
      'Memryx',
      'DroidLens',
      'Other'
    ],

    Categories: [
      'Development',
      'Bug Fix',
      'UI/UX',
      'API',
      'Testing',
      'Research',
      'Deployment',
      'Meeting',
      'Other'
    ],

    Priorities: [
      'Low',
      'Medium',
      'High',
      'Urgent'
    ],

    Statuses: [
      'Pending',
      'In Progress',
      'Blocked',
      'Completed',
      'Cancelled'
    ]
  }
};