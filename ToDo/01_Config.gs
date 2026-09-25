/**
 * Central configuration for the Eisenhower workspace.
 * Every top-level constant lives here: other files only declare functions,
 * so Apps Script's file load order can never hit an uninitialized constant.
 */

const CONFIG = {
  sheets: { todo: 'TODO', matrix: 'Eisenhower Matrix', dashboard: 'Dashboard', archive: 'Archive' },
  todoRows: 200,            // minimum rows on the TODO sheet (incl. header)
  spareRows: 50,            // empty rows kept below the last task on rebuild
  slots: 12,                // visible task rows per quadrant
  panelPollMs: 1000,        // how often the task sidebar checks the selected row
  includeSampleTasks: true, // seed sample tasks when the TODO sheet is first created
  defaultStatus: 'To Do',   // filled in automatically when a new task is typed
  dateFormat: 'mmm d, yyyy',
  font: 'Inter',
};

const THEME = {
  ink: '#111827',
  muted: '#6b7280',
  faint: '#9ca3af',
  border: '#e5e7eb',
  divider: '#f3f4f6',
  surface: '#f9fafb',
  accent: '#6366f1',
  danger: '#dc2626',
  warning: '#d97706',
};

// TODO sheet columns, in order. Formulas elsewhere look columns up by key,
// so reordering here is safe as long as Important and Urgent stay adjacent.
const COLUMNS = [
  { key: 'task',      header: 'Task',      width: 340 },
  { key: 'important', header: 'Important', width: 100 },
  { key: 'urgent',    header: 'Urgent',    width: 100 },
  { key: 'status',    header: 'Status',    width: 130 },
  { key: 'due',       header: 'Due',       width: 120 },
  { key: 'notes',     header: 'Notes',     width: 320 },
];

const COL = {}; // key -> 1-based column index
COLUMNS.forEach((c, i) => { COL[c.key] = i + 1; });

const STATUSES = {
  'To Do':       { bg: '#f3f4f6', fg: '#374151' },
  'In Progress': { bg: '#dbeafe', fg: '#1e40af' },
  'Done':        { bg: '#dcfce7', fg: '#166534' },
  'Blocked':     { bg: '#fee2e2', fg: '#991b1b' },
};

const QUADRANTS = [
  { id: 'Q1', title: 'Do First',  caption: 'Important · Urgent',         important: true,  urgent: true,  bg: '#fef2f2', text: '#991b1b', accent: '#ef4444' },
  { id: 'Q2', title: 'Schedule',  caption: 'Important · Not urgent',     important: true,  urgent: false, bg: '#f0fdf4', text: '#166534', accent: '#22c55e' },
  { id: 'Q3', title: 'Delegate',  caption: 'Not important · Urgent',     important: false, urgent: true,  bg: '#fffbeb', text: '#92400e', accent: '#f59e0b' },
  { id: 'Q4', title: 'Eliminate', caption: 'Not important · Not urgent', important: false, urgent: false, bg: '#f9fafb', text: '#374151', accent: '#9ca3af' },
];

const SAMPLE_TASKS = [
  // Task, Important, Urgent, Status, Due (days from today, '' = none), Notes
  ['Fix checkout bug reported by customers', true,  true,  'In Progress', 0,  'Hotfix before Friday release'],
  ['Prepare investor update deck',           true,  true,  'To Do',       1,  'Due tomorrow 9 AM'],
  ['Renew SSL certificate',                  true,  true,  'Blocked',     -2, 'Waiting on DNS access'],
  ['Plan Q4 product roadmap',                true,  false, 'In Progress', 14, 'Draft with PM team'],
  ['Weekly 1:1s with direct reports',        true,  false, 'To Do',       3,  'Block recurring calendar time'],
  ['Learn advanced SQL',                     true,  false, 'To Do',       '', ''],
  ['Reply to vendor pricing emails',         false, true,  'To Do',       2,  'Delegate to Ops'],
  ['Book team offsite logistics',            false, true,  'In Progress', 5,  'Ask office manager'],
  ['Reorganize old Drive folders',           false, false, 'To Do',       '', ''],
  ['Review newsletter subscriptions',        false, false, 'Done',        -5, 'Unsubscribed from 12'],
];

const EMPTY_LIST_TEXT = 'No open tasks';
const MORE_PREFIX = '…and ';

// Quoted sheet reference, safe even if the name contains spaces
const TODO_REF = `'${CONFIG.sheets.todo}'`;
