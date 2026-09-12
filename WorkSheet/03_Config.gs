const CONFIG = {
  HEADER_ROW: 1,
  DATE_FORMAT: 'MMdd',
  DEFAULT_TASKS_PER_DAY: 1,
  LISTS_SHEET_NAME: 'Lists',

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
    BORDER: '#d9d9d9'
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