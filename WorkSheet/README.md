# WorkSheet — Modern SaaS Work & Task Tracker for Google Sheets

An automated, modular Google Apps Script application transforming Google Sheets into a **lightweight, modern SaaS-style Work Tracker** (Linear + Notion aesthetic). It features dynamic monthly calendar sheet generation, a 14-column executive Command Center Dashboard, an Eisenhower Matrix Todo tracker with live formula metrics, single-selection quadrant radio checkboxes, Daily Status Report (DSR) automation, centralized protected reference lists, and standalone production HTML modal dialogs.

---

## Table of Contents

- [Overview](#overview)
- [Architecture & File Structure](#architecture--file-structure)
- [Master Setup (`setupWorkTracker`)](#master-setup-setupworktracker)
- [Core Features & SaaS UX](#core-features--saas-ux)
  - [1. Executive Command Center Dashboard (14 Columns)](#1-executive-command-center-dashboard-14-columns)
  - [2. Eisenhower Matrix Todo Tracker](#2-eisenhower-matrix-todo-tracker)
  - [3. Monthly Task Tracking & Quick Entry](#3-monthly-task-tracking--quick-entry)
  - [4. Daily Status Report (DSR) Engine](#4-daily-status-report-dsr-engine)
  - [5. Visual Calendar Picker](#5-visual-calendar-picker)
  - [6. Unified WorkSheet Toolbar Menu](#6-unified-worksheet-toolbar-menu)
- [Design System & Typography](#design-system--typography)
- [Configuration Guide (`03_Config.gs`)](#configuration-guide-03_configgs)
- [Trigger Architecture & Real-Time Sync](#trigger-architecture--real-time-sync)
- [Developer Guidelines](#developer-guidelines)

---

## Overview

WorkSheet eliminates the friction of manual spreadsheet maintenance. It provisions a complete, interconnected productivity suite inside Google Sheets:
- **Command Center Dashboard**: Live high-level KPI cards answering what is active *Today* vs *This Month*, side-by-side Monthly Breakdown and 6-column Project Performance tables, and Eisenhower Todo statistics.
- **Fast Batch Task Logging**: Paste multi-line tasks or select single dates with automatic `Pending` defaults, keyboard shortcuts (`Ctrl + Enter`), and `+ Add & Next Project` workflow.
- **Eisenhower Decision Matrix**: Quadrant checkboxes (`Q1: Do`, `Q2: Schedule`, `Q3: Delegate`, `Q4: Don't Do`) with mutually-exclusive radio behavior and frozen live stats bar.
- **Single Master Setup**: Idempotent `setupWorkTracker()` entry point that verifies, upgrades, and configures the system without data loss or duplicate triggers.

---

## Architecture & File Structure

The codebase is organized into 8 functional modules comprising 14 Google Apps Script files and 5 standalone production HTML templates:

```text
WorkSheet/
├── Code.gs                           # Master setupWorkTracker() pipeline and centralized triggers (onEdit, onChange)
├── admin/                            # Module 01: Setup, migrations, and developer seeding tools
│   ├── 01_Migration.gs               # Safe, non-destructive migration engine upgrading Todo & Month sheets
│   └── 01_SampleData.gs              # 1-click realistic dummy data seeding engine
├── core/                             # Module 02: Central configuration, reference lists, menus, and utilities
│   ├── 02_Config.gs                  # Central configuration: tokens, schemas, dimensions, colors, fonts
│   ├── 02_Lists.gs                   # Validation reference lists, named ranges, and sheet protection
│   ├── 02_Menu.gs                    # Four top-level toolbar menus: WorkSheet, Setup, Tasks, and DSR
│   └── 02_Utils.gs                   # Helper utilities: timezones, dates, sheet trimming, safe gridlines
├── dashboard/                        # Module 03: Executive Command Center dashboard
│   └── 03_Dashboard.gs               # Modern SaaS Command Center dashboard (14 columns, backlog schedule KPIs)
├── dsr/                              # Module 04: Daily Status Report compilation and calendar picker
│   ├── 04_CalendarPicker.gs          # Controller for interactive visual calendar date picker
│   ├── 04_CalendarPickerDialog.html  # Production modal dialog UI for calendar date selection
│   ├── 04_DSR.gs                     # Daily Status Report generator with Task ID traceability and export
│   └── 04_DSRDialog.html             # Production modal dialog UI for DSR reports with date switching
├── styles/                           # Module 05: Visual formatting and conditional rules
│   ├── 05_ConditionalFormatting.gs   # Semantic status, priority, and weekend formatting rules
│   └── 05_Formatting.gs              # Visual styling, layout dimensions, borders, and safe gridline hiding
├── tasks/                            # Module 06: Daily task logging and monthly sheet engine
│   ├── 06_MonthSheet.gs              # 8-column monthly task sheet creation with Task ID dropdown linking
│   ├── 06_Tasks.gs                   # Task operations, batch insertion, dropdown rules, and modal dialog
│   └── 06_TaskDialog.html            # Production modal dialog UI for multi-task batch entry
├── timeline/                         # Module 07: Decisions sheet and audit engine
│   ├── 07_Timeline.gs                # Chronological Decisions timeline engine and monotonic ID generator
│   ├── 07_DecisionDialog.html        # Production modal dialog UI for recording decisions and timeline events
│   └── 07_TimelineDialog.html        # Interactive visual Task History / Timeline viewer modal dialog
├── todo/                             # Module 08: Backlog management and monotonic task IDs
│   └── 08_Todo.gs                    # Simplified 5-column Todo backlog engine with auto-assigned Task IDs
├── GEMINI.md                         # Antigravity context & project memory
└── README.md                         # Complete developer & user documentation
```

### File Responsibilities by Module

| Module | File | Primary Functions / Role | Description |
| :--- | :--- | :--- | :--- |
| **`root`** | [`Code.gs`](./Code.gs) | `setupWorkTracker`, `initializeWorkTracker`, `rebuildLists`, `onEdit`, `onChange` | Primary operational entry point. Idempotent master setup routine and workbook trigger router. |
| **`admin/`** | [`01_Migration.gs`](./admin/01_Migration.gs) | `runFullSystemMigration`, `migrateTodoSheet`, `migrateAllMonthSheets` | Automated, non-destructive migration engine upgrading Todo to 5 columns and Month sheets to 8 columns. |
| **`admin/`** | [`01_SampleData.gs`](./admin/01_SampleData.gs) | `populateDummyData`, `populateTodoDummyData`, `populateMonthDummyData`, `populateDecisionsDummyData` | Seeds 12 backlog tasks with Task IDs, 21 monthly tasks, and sample timeline events with 1 click. |
| **`core/`** | [`02_Config.gs`](./core/02_Config.gs) | `CONFIG` object | Central configuration for dimensions, typography, colors, headers, and reference lists. |
| **`core/`** | [`02_Lists.gs`](./core/02_Lists.gs) | `createListsSheet`, `ensureListsSheet`, `createNamedRanges`, `protectListsSheet` | Manages reference lists, named ranges consumed by dropdowns, and sheet protection. |
| **`core/`** | [`02_Menu.gs`](./core/02_Menu.gs) | `onOpen` | Injects four top-level toolbar menus: `WorkSheet`, `Setup`, `Tasks`, and `DSR`. |
| **`core/`** | [`02_Utils.gs`](./core/02_Utils.gs) | `setSheetGridlinesHidden`, `trimSheet`, `getSpreadsheet`, `getTimezone`, `getToday` | Core utilities including safe gridline toggling that avoids unsupported Google Sheets API errors. |
| **`dashboard/`** | [`03_Dashboard.gs`](./dashboard/03_Dashboard.gs) | `refreshDashboard`, `updateDashboardOnChange`, `renderKpiCards`, `renderTablesSection`, `renderTodoSection` | Modern 14-column SaaS Command Center aggregating monthly metrics, project stats, and backlog schedule KPIs. |
| **`dsr/`** | [`04_CalendarPicker.gs`](./dsr/04_CalendarPicker.gs) | `openCalendarPicker`, `proceedFromCalendar` | Backend controller for the visual calendar picker; loads `04_CalendarPickerDialog.html` with resilient fallback. |
| **`dsr/`** | [`04_CalendarPickerDialog.html`](./dsr/04_CalendarPickerDialog.html) | Calendar Modal UI | Standalone visual date picker modal with `Varela Round` and `Roboto Mono` fonts, keyboard navigation, and date selection. |
| **`dsr/`** | [`04_DSR.gs`](./dsr/04_DSR.gs) | `generateDSRForToday`, `generateDSRForSelectedDate`, `showDSRDialog`, `saveDSRToSheet` | Compiles daily status reports categorized by project with Task ID traceability, copy, .txt download, and sheet logging. |
| **`dsr/`** | [`04_DSRDialog.html`](./dsr/04_DSRDialog.html) | DSR Modal UI | Interactive DSR modal dialog with live calendar date switcher and 1-click export actions. |
| **`styles/`** | [`05_ConditionalFormatting.gs`](./styles/05_ConditionalFormatting.gs) | `setupConditionalFormatting`, `setupWeekendFormatting` | Applies semantic pastel tags for Statuses (Col H), Priorities (Col G), and weekend rows across columns A–H. |
| **`styles/`** | [`05_Formatting.gs`](./styles/05_Formatting.gs) | `formatWorkTracker`, `formatHeader`, `formatColumns`, `formatDimensions`, `formatBorders` | Applies SaaS typography, subtle `#cbd5e1` borders, dimensions, and safe gridline hiding across 8 columns. |
| **`tasks/`** | [`06_MonthSheet.gs`](./tasks/06_MonthSheet.gs) | `createCurrentMonthSheet`, `createMonthRows` | Idempotently creates or refreshes 8-column monthly task sheets without deleting existing task data. |
| **`tasks/`** | [`06_Tasks.gs`](./tasks/06_Tasks.gs) | `fillTaskForToday`, `fillTaskForSelectedDate`, `saveTasksBatch`, `clearTasks` | Controller for fast multi-task batch logging; loads `06_TaskDialog.html` with resilient fallback. |
| **`tasks/`** | [`06_TaskDialog.html`](./tasks/06_TaskDialog.html) | Task Entry Modal UI | Clean batch task modal with project/category/priority selection, `Ctrl + Enter` shortcut, and `+ Add & Next Project` workflow. |
| **`timeline/`** | [`07_Timeline.gs`](./timeline/07_Timeline.gs) | `ensureDecisionsSheet`, `recordTimelineEvent`, `getTaskTimelineHistory`, `generateNextTaskId`, `handleSpreadsheetEdit` | Central engine for the Decisions timeline, monotonic ID generation, and automatic status change logging. |
| **`timeline/`** | [`07_DecisionDialog.html`](./timeline/07_DecisionDialog.html) | Decision Entry Modal UI | Production modal dialog for manually recording decisions, rationale, owners, and follow-up actions. |
| **`timeline/`** | [`07_TimelineDialog.html`](./timeline/07_TimelineDialog.html) | Task History Modal UI | Interactive chronological visual timeline viewer rendering the full lifecycle history of any task. |
| **`todo/`** | [`08_Todo.gs`](./todo/08_Todo.gs) | `createTodoSheet`, `ensureTodoSheet`, `setupTodoStructure`, `ensureAllTodoRowsHaveIds` | Simplified 5-column Todo backlog tracker with auto-assigned permanent Task IDs (`TSK-XXXX`). |

---

## Master Setup (`setupWorkTracker`)

`setupWorkTracker(suppressAlert)` in [`Code.gs`](./Code.gs) is the **single master entry point** for provisioning, repairing, or upgrading the entire system.

### Key Behaviors:
1. **Idempotent**: Safe to run once or ten times. It never duplicates sheets, triggers, named ranges, or formatting rules.
2. **Non-Destructive**: Strictly preserves existing user tasks and entries on Month and Todo sheets.
3. **Trigger Verification**: Detects existing triggers via `ScriptApp.getProjectTriggers()` to prevent duplicate registrations.
4. **Structured Logging**: Outputs clean milestone logs in the Apps Script execution log:
   ```text
   [WorkSheet] Setup started
   [WorkSheet] Lists verified
   [WorkSheet] Todo verified
   [WorkSheet] Current month verified
   [WorkSheet] Dashboard verified
   [WorkSheet] Validations & Formatting configured
   [WorkSheet] Triggers verified
   [WorkSheet] Setup completed
   ```

---

## Core Features & SaaS UX

### 1. Executive Command Center Dashboard (14 Columns)
The Dashboard serves as the central command center:
- **Top Metadata Bar**: Displays live timestamp, engine version, and system status.
- **6 Modern SaaS KPI Cards**:
  - `TOTAL TASKS`: Overall tasks across all months + subtitle with tasks logged *Today*.
  - `COMPLETED`: Total delivered tasks + subtitle with tasks completed *Today*.
  - `IN PROGRESS`: Total active pipeline + subtitle with tasks in progress *Today*.
  - `BLOCKED`: Total blockers + subtitle with blockers *Today*.
  - `PENDING`: Total open backlog + subtitle with pending tasks *Today*.
  - `COMPLETION RATE`: Grand completion percentage + subtitle with today's rate.
- **Side-by-Side Analytics Section (14 Columns)**:
  - **Monthly Breakdown (Cols 1–7 / A–G)**: `Month`, `Total`, `Done`, `Active`, `Blocked`, `Pending`, `Progress`.
  - **Spacer Gap (Col 8 / H)**: 35px subtle visual separator.
  - **Project Performance (Cols 9–14 / I–N)**: `Project`, `Total`, `Completed`, `In Progress`, `Blocked`, `Completion %`. Sorted deterministically by task volume and alphabetical order.
  - **Summary Totals Row**: Centered monospaced totals with green delivery rate highlights.
- **Backlog Command & Schedule Section (14 Columns)**:
  - 5 dedicated KPI cards: `TOTAL BACKLOG`, `⚠️ DUE / OVERDUE`, `📅 UPCOMING (7D)`, `⚪ NO DUE DATE`, and `📁 ACTIVE PROJECTS`.
- **Clean Canvas**: Seamless white surface with gridlines safely hidden via `setSheetGridlinesHidden`.

### 2. Simplified Todo Backlog & Permanent Task IDs
The dedicated `Todo` sheet implements a clean 5-column project backlog:
- **Columns**: `ID`, `Task Description`, `Project`, `Due Date`, `Notes`.
- **Auto-Assigned Monotonic IDs**: Formatted as `TSK-0001`, `TSK-0002`... Generated via script locks and `PropertiesService` so IDs are permanent and never reused.
- **Project Dropdown**: Validated against `ProjectsList` on the `Lists` sheet.
- **Due Date**: Target deadline formatted in `Roboto Mono` (`yyyy-MM-dd`).

### 3. Monthly Task Tracking & Task ID Linking
- **Standard Schema (8 Columns)**: `Date`, `Day`, `Task ID`, `Project`, `Task`, `Category`, `Priority`, `Status`.
- **Task ID Auto-Population**:
  - Selecting a `Task ID` in Column C automatically fills `Task` (Col E) and `Project` (Col D) from the Todo backlog.
  - Historical work log entries remain preserved even if the original task is completed or removed from Todo.
- **Fast Batch Task Logging Modal (`tasks/06_TaskDialog.html`)**:
  - Accessible via **Tasks** > **Fill Task for Today** or **Fill Task for Selected Date**.
  - Select Project, Category, and Priority once; enter or paste multiple tasks into a textarea.
  - Automatically defaults to `Pending` status and records `Task Created` in the Decisions timeline.
  - Includes `Ctrl + Enter` shortcut and `+ Add & Next Project` continuous entry.

### 4. Decisions & Chronological Task Timeline
- Dedicated `Decisions` sheet providing an immutable audit trail of all task events:
  - Columns: `Timeline ID`, `Task ID`, `Event Date`, `Event Type`, `Previous Status`, `New Status`, `Decision / Event`, `Rationale`, `Decided By`, `Next Action`, `Due Date`, `Reference`.
  - Supports 12 event types including `Task Created`, `Status Changed`, `Decision Made`, `Blocker Identified`, `Blocker Resolved`, `Due Date Changed`, `Scope Changed`, etc.
  - Accessible via **Tasks** > **Record Decision / Event** (`timeline/07_DecisionDialog.html`) and **Tasks** > **View Task History & Timeline** (`timeline/07_TimelineDialog.html`).

### 5. Daily Status Report (DSR) Engine
- Accessible via **DSR** > **Generate DSR for Today** or **Generate DSR for Selected Date**.
- Automatically groups tasks by project into:
  1. `Tasks Completed`
  2. `Work in Progress`
  3. `Blockers / Issues`
  4. `Plan for Next Working Day`
- Formats tasks with their `[Task ID]` prefix for full traceability back to the timeline.
- Styled interactive modal (`dsr/04_DSRDialog.html`) with live date switching, **Copy to Clipboard**, **Save to Sheet**, and **Download .txt**.

### 6. Four Top-Level Toolbar Menus
Clean, modular top-level menu hierarchy directly on the Google Sheets menu bar:

```text
1. WorkSheet (Navigation & Core Views)
   ├── 📊 Open Dashboard           → openDashboard()
   ├── 🎯 Open Todo Backlog        → openTodoSheet()
   ├── 📜 Open Decisions Timeline  → openDecisionsSheet()
   ├── 📅 Create Current Month Sheet → createCurrentMonthSheet()
   ├── ───────────────────────────
   └── ⚡ Setup / Repair Work Tracker → setupWorkTracker()

2. Setup (System Administration & Migration)
   ├── ⚡ Setup / Repair Work Tracker → setupWorkTracker()
   ├── 🔄 Refresh Dashboard        → refreshDashboard()
   ├── 📋 Rebuild Reference Lists  → rebuildLists()
   ├── 📅 Create Current Month Sheet → createCurrentMonthSheet()
   ├── 🚀 Run Migration (Task IDs & Decisions) → runFullSystemMigration()
   ├── ───────────────────────────
   └── 🧪 Populate Dummy Data      → populateDummyData()

3. Tasks (Daily Task Logging & Timeline Tracking)
   ├── Fill Task for Today         → fillTaskForToday()
   ├── Fill Task for Selected Date → fillTaskForSelectedDate()
   ├── ───────────────────────────
   ├── 📝 Record Decision / Event  → openRecordDecisionDialog()
   └── 🔍 View Task History & Timeline → openTaskHistoryDialog()

4. DSR (Daily Status Report Generation)
   ├── Generate DSR for Today      → generateDSRForToday()
   └── Generate DSR for Selected Date → generateDSRForSelectedDate()
```

---

## Design System & Typography

- **Typography**:
  - `Varela Round`: Primary UI font for titles, section headers, labels, and text descriptions.
  - `Roboto Mono`: Monospaced font for all dates, numbers, KPI values, metrics, and percentages.
- **Surfaces & Borders**: Soft near-white surfaces (`#f8fafc`), subtle borders (`#cbd5e1` / `#e2e8f0`).
- **Semantic Palette**:
  - `Completed`: `#f0fdf4` fill, `#15803d` text
  - `In Progress`: `#eff6ff` fill, `#1d4ed8` text
  - `Blocked`: `#fef2f2` fill, `#b91c1c` text
  - `Pending`: `#f8fafc` fill, `#475569` text
  - `Cancelled`: `#f1f5f9` fill, `#64748b` text

---

## Configuration Guide (`core/02_Config.gs`)

All system parameters live centrally in `CONFIG` within [`core/02_Config.gs`](./core/02_Config.gs):

```javascript
const CONFIG = {
  HEADER_ROW: 1,
  DATE_FORMAT: 'MMdd',
  DEFAULT_TASKS_PER_DAY: 1,

  SHEETS: {
    LISTS: 'Lists',
    TODO: 'Todo',
    DASHBOARD: 'Dashboard',
    DSR: 'DSR'
  },

  FONTS: {
    TEXT: 'Varela Round',
    DIGITS: 'Roboto Mono'
  },

  DIMENSIONS: {
    MONTH_COL_WIDTHS: [75, 60, 130, 340, 130, 100, 125],
    HEADER_ROW_HEIGHT: 28,
    DATA_ROW_HEIGHT: 42,
    TODO_COL_WIDTHS: [340, 140, 120, 120, 120, 120],
    DASHBOARD_COL_WIDTHS: [130, 75, 75, 75, 75, 75, 85, 35, 140, 75, 75, 75, 75, 85],
    DASHBOARD_COLUMNS_COUNT: 14
  },

  LISTS: {
    Projects: ['Clinkio', 'Surfari', 'Workbench', 'Memryx', 'DroidLens', 'Other'],
    Categories: ['Development', 'Bug Fix', 'UI/UX', 'API', 'Testing', 'Research', 'Deployment', 'Meeting', 'Other'],
    Priorities: ['Low', 'Medium', 'High', 'Urgent'],
    Statuses: ['Pending', 'In Progress', 'Blocked', 'Completed', 'Cancelled']
  }
};
```

---

## Trigger Architecture & Real-Time Sync

1. **Centralized Router in `Code.gs`**:
   - `onEdit(e)` simple trigger coordinates spreadsheet edits and `updateDashboardOnChange(e)`.
   - `onChange(e)` installable trigger listens for structural sheet changes.
2. **Duplicate Prevention**: `setupTriggers(ss)` checks `ScriptApp.getProjectTriggers()` to guarantee only one `onChange` trigger exists.
3. **Silent Execution**: Background updates run with `suppressAlert = true` and `keepActiveSheet = true` so active user workflows are never interrupted.

---

## Developer Guidelines

- **Configuration First**: Never hardcode colors, dimensions, or list items; always declare them in `core/02_Config.gs`.
- **Batch Processing**: Use bulk `getValues()` and `setValues()`; avoid spreadsheet API calls inside iteration loops.
- **Safe Gridlines**: Always use `setSheetGridlinesHidden(sheet, hidden)` in `core/02_Utils.gs`. Never call `sheet.setHideGridlines()` as it is not a native method on `Sheet`.
- **Preserve User Data**: Setup and maintenance scripts must always treat existing user sheets non-destructively.
