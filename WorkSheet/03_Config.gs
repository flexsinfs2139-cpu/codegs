// ============================================================
// CONFIG.GS — CENTRAL APPLICATION CONFIGURATION
// ============================================================

const CONFIG = {
  // Core Operational Settings
  HEADER_ROW: 1,
  DATE_FORMAT: 'MMdd',
  DEFAULT_TASKS_PER_DAY: 1,

  // Sheet Names & Aliases
  SHEETS: {
    LISTS: 'Lists',
    TODO: 'Todo',
    DASHBOARD: 'Dashboard',
    DSR: 'DSR',
    DECISIONS: 'Decisions'
  },
  LISTS_SHEET_NAME: 'Lists',
  TODO_SHEET_NAME: 'Todo',
  DASHBOARD_SHEET_NAME: 'Dashboard',
  DECISIONS_SHEET_NAME: 'Decisions',

  // Todo Structure Settings (Simplified 5-column layout)
  TODO_DROPDOWN_ROWS: 200,
  TODO_HEADER_ROW: 1,
  TODO_FIRST_DATA_ROW: 2,

  // Decisions Structure Settings (12-column chronological timeline)
  DECISIONS_HEADER_ROW: 1,
  DECISIONS_FIRST_DATA_ROW: 2,

  // Design System Typography
  FONTS: {
    TEXT: 'Varela Round',
    DIGITS: 'Roboto Mono'
  },

  // Sheet Column Headers
  // Monthly Work Sheet: 8 Columns
  HEADERS: [
    'Date',
    'Day',
    'Task ID',
    'Project',
    'Task',
    'Category',
    'Priority',
    'Status'
  ],

  // Simplified Todo Sheet: 5 Columns
  TODO_HEADERS: [
    'ID',
    'Task Description',
    'Project',
    'Due Date',
    'Notes'
  ],

  // Decisions & Chronological Task Timeline Sheet: 12 Columns
  DECISIONS_HEADERS: [
    'Timeline ID',
    'Task ID',
    'Event Date',
    'Event Type',
    'Previous Status',
    'New Status',
    'Decision / Event',
    'Rationale',
    'Decided By',
    'Next Action',
    'Due Date',
    'Reference'
  ],

  // Layout Dimensions & Sizing
  DIMENSIONS: {
    MONTH_COL_WIDTHS: [75, 60, 100, 130, 340, 130, 100, 125],
    HEADER_ROW_HEIGHT: 28,
    DATA_ROW_HEIGHT: 42,
    TODO_COL_WIDTHS: [95, 360, 140, 110, 280],
    TODO_ROW_HEIGHT: 36,
    DECISIONS_COL_WIDTHS: [110, 100, 140, 150, 120, 120, 280, 280, 130, 220, 100, 180],
    DECISIONS_ROW_HEIGHT: 36,
    DASHBOARD_COL_WIDTHS: [130, 75, 75, 75, 75, 75, 85, 35, 140, 75, 75, 75, 75, 85],
    DASHBOARD_COLUMNS_COUNT: 14
  },

  // Restrained Modern SaaS Palette
  COLORS: {
    HEADER_BG: '#f8fafc',
    HEADER_TEXT: '#0f172a',
    HEADER: '#f8fafc', // Backwards compatibility
    BORDER: '#cbd5e1',
    BORDER_LIGHT: '#e2e8f0',
    STATS_BG: '#f8fafc',
    STATS_BORDER: '#e2e8f0',

    // Semantic Status Colors
    STATUS: {
      COMPLETED: { bg: '#f0fdf4', text: '#15803d', border: '#86efac' },
      IN_PROGRESS: { bg: '#eff6ff', text: '#1d4ed8', border: '#93c5fd' },
      BLOCKED: { bg: '#fef2f2', text: '#b91c1c', border: '#fca5a5' },
      PENDING: { bg: '#f8fafc', text: '#475569', border: '#cbd5e1' },
      CANCELLED: { bg: '#f1f5f9', text: '#64748b', border: '#e2e8f0' }
    },

    // Semantic Priority Colors
    PRIORITY: {
      URGENT: { bg: '#fef2f2', text: '#b91c1c' },
      HIGH: { bg: '#fff7ed', text: '#c2410c' },
      MEDIUM: { bg: '#fefce8', text: '#854d0e' },
      LOW: { bg: '#f8fafc', text: '#64748b' }
    }
  },

  // SaaS Command Center Dashboard Configuration
  DASHBOARD: {
    TITLE: '⚡ WORK & PROJECT COMMAND CENTER',
    VERSION: 'Modern SaaS Engine v4.0',
    COLUMNS_COUNT: 14,
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

  // Timeline Event Types
  EVENT_TYPES: [
    'Task Created',
    'Status Changed',
    'Decision Made',
    'Blocker Identified',
    'Blocker Resolved',
    'Due Date Changed',
    'Scope Changed',
    'Priority Changed',
    'Task Reopened',
    'Task Completed',
    'Task Cancelled',
    'Follow-up Action'
  ],

  // Named Range Identifiers
  NAMED_RANGES: {
    PROJECTS: 'ProjectsList',
    CATEGORIES: 'CategoriesList',
    PRIORITIES: 'PrioritiesList',
    STATUSES: 'StatusesList',
    TODO_TASK_IDS: 'TodoTaskIds',
    EVENT_TYPES: 'EventTypesList'
  },

  // Reference Lists & Validation Options
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
    ],

    EventTypes: [
      'Task Created',
      'Status Changed',
      'Decision Made',
      'Blocker Identified',
      'Blocker Resolved',
      'Due Date Changed',
      'Scope Changed',
      'Priority Changed',
      'Task Reopened',
      'Task Completed',
      'Task Cancelled',
      'Follow-up Action'
    ]
  }
};