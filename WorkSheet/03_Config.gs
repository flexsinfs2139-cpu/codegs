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

  FONTS: {
    TEXT: 'Varela Round',
    DIGITS: 'Roboto Mono'
  },

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

  DASHBOARD: {
    TITLE: '⚡ WORK & PROJECT COMMAND CENTER',
    VERSION: 'Modern SaaS Engine v3.0',
    COLUMNS_COUNT: 13,
    COLORS: {
      CARD_TOTAL_BG: '#eff6ff',
      CARD_TOTAL_BORDER: '#2563eb',
      CARD_TOTAL_TEXT: '#1d4ed8',

      CARD_DONE_BG: '#f0fdf4',
      CARD_DONE_BORDER: '#16a34a',
      CARD_DONE_TEXT: '#15803d',

      CARD_ACTIVE_BG: '#fffbeb',
      CARD_ACTIVE_BORDER: '#d97706',
      CARD_ACTIVE_TEXT: '#b45309',

      CARD_BLOCKED_BG: '#fef2f2',
      CARD_BLOCKED_BORDER: '#dc2626',
      CARD_BLOCKED_TEXT: '#b91c1c',

      CARD_PENDING_BG: '#f5f3ff',
      CARD_PENDING_BORDER: '#7c3aed',
      CARD_PENDING_TEXT: '#6d28d9',

      CARD_RATE_BG: '#f0fdfa',
      CARD_RATE_BORDER: '#0d9488',
      CARD_RATE_TEXT: '#0f766e',

      BANNER_BG: '#e0e7ff',
      BANNER_TEXT: '#4338ca',
      TABLE_HEADER_BG: '#f8fafc',
      TABLE_HEADER_TEXT: '#0f172a',
      GREEN_ACCENT: '#16a34a',
      MUTED_TEXT: '#64748b',
      BORDER: '#cbd5e1'
    }
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