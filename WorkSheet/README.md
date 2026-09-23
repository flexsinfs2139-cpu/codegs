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

The codebase is organized into 14 modular Google Apps Script files and 3 standalone production HTML templates:

```text
WorkSheet/
├── 00_Code.gs                  # Master setupWorkTracker() pipeline and centralized triggers (onEdit, onChange)
├── 01_CalendarPicker.gs        # Controller for interactive visual calendar date picker
├── 01_CalendarPickerDialog.html # Production modal dialog UI for calendar date selection
├── 02_ConditionalFormatting.gs # Semantic status, priority, weekend, and Eisenhower formatting rules
├── 03_Config.gs                # Central configuration: tokens, schemas, dimensions, colors, fonts
├── 04_DSR.gs                   # Daily Status Report generator, parser, and clipboard exporter
├── 04_DSRDialog.html           # Production modal dialog UI for DSR reports with date switching
├── 05_Formatting.gs            # Visual styling, layout dimensions, borders, and safe gridline hiding
├── 06_Lists.gs                 # Validation reference lists, named ranges, and sheet protection
├── 07_Menu.gs                  # Single unified 'WorkSheet' menu with submenus and direct openers
├── 08_MonthSheet.gs            # Idempotent monthly sheet creation and non-destructive refresh
├── 09_Tasks.gs                 # Task operations, batch insertion, dropdown rules, and modal dialog
├── 09_TaskDialog.html          # Production modal dialog UI for multi-task batch entry
├── 10_Utils.gs                 # Helper utilities: timezones, dates, sheet trimming, safe gridlines
├── 11_Todo.gs                  # Eisenhower Matrix Todo tracker with live stats bar & radio checkboxes
├── 12_Dashboard.gs             # Modern SaaS Command Center dashboard (14 columns, live metrics)
├── 13_SampleData.gs            # 1-click realistic dummy data seeding engine
└── README.md                   # Complete developer & user documentation
```

### File Responsibilities

| File | Primary Functions / Role | Description |
| :--- | :--- | :--- |
| **`00_Code.gs`** | `setupWorkTracker`, `initializeWorkTracker`, `rebuildLists`, `onEdit`, `onChange` | Primary operational entry point. Idempotent master setup routine and workbook trigger router. |
| **`01_CalendarPicker.gs`** | `openCalendarPicker`, `proceedFromCalendar` | Backend controller for the visual calendar picker; loads `01_CalendarPickerDialog.html` with resilient fallback. |
| **`01_CalendarPickerDialog.html`** | Calendar Modal UI | Standalone visual date picker modal with `Varela Round` and `Roboto Mono` fonts, keyboard navigation, and double-click date selection. |
| **`02_ConditionalFormatting.gs`** | `setupConditionalFormatting`, `setupWeekendFormatting`, `setupTodoConditionalFormatting` | Applies semantic pastel tags for Statuses, Priorities, weekend rows, and Eisenhower Matrix quadrants. |
| **`03_Config.gs`** | `CONFIG` object | Central configuration for dimensions, typography, colors, headers, and reference lists. |
| **`04_DSR.gs`** | `generateDSRForToday`, `generateDSRForSelectedDate`, `showDSRDialog`, `saveDSRToSheet` | Compiles daily status reports categorized by project with clipboard copy, .txt download, and sheet logging. |
| **`04_DSRDialog.html`** | DSR Modal UI | Interactive DSR modal dialog with live calendar date switcher and 1-click export actions. |
| **`05_Formatting.gs`** | `formatWorkTracker`, `formatHeader`, `formatColumns`, `formatDimensions`, `formatBorders` | Applies SaaS typography, subtle `#cbd5e1` borders, dimensions, and safe gridline hiding. |
| **`06_Lists.gs`** | `createListsSheet`, `ensureListsSheet`, `createNamedRanges`, `protectListsSheet` | Manages reference lists, named ranges consumed by dropdowns, and sheet protection. |
| **`07_Menu.gs`** | `onOpen`, `openDashboard`, `openTodoSheet` | Injects the unified `WorkSheet` toolbar menu with `Tasks`, `DSR`, and `Advanced` submenus. |
| **`08_MonthSheet.gs`** | `createCurrentMonthSheet`, `createMonthRows` | Idempotently creates or refreshes monthly task sheets without deleting existing task data. |
| **`09_Tasks.gs`** | `fillTaskForToday`, `fillTaskForSelectedDate`, `saveTasksBatch`, `clearTasks` | Controller for fast multi-task batch logging; loads `09_TaskDialog.html` with resilient fallback. |
| **`09_TaskDialog.html`** | Task Entry Modal UI | Clean batch task modal with project/category/priority selection, `Ctrl + Enter` shortcut, and `+ Add & Next Project` workflow. |
| **`10_Utils.gs`** | `setSheetGridlinesHidden`, `trimSheet`, `getSpreadsheet`, `getTimezone`, `getToday` | Core utilities including safe gridline toggling that avoids unsupported Google Sheets API errors. |
| **`11_Todo.gs`** | `createTodoSheet`, `ensureTodoSheet`, `setupTodoStructure`, `handleTodoQuadrantExclusiveSelect` | Eisenhower Matrix Todo tracker with Row 1 live formula counters and mutually exclusive radio checkboxes. |
| **`12_Dashboard.gs`** | `refreshDashboard`, `updateDashboardOnChange`, `renderKpiCards`, `renderTablesSection`, `renderTodoSection` | Modern 14-column SaaS Command Center aggregating monthly and today's metrics across the entire workbook. |
| **`13_SampleData.gs`** | `populateDummyData`, `populateTodoDummyData`, `populateMonthDummyData` | Seeds 21 realistic engineering tasks and 12 Eisenhower matrix items with 1 click. |

---

## Master Setup (`setupWorkTracker`)

`setupWorkTracker(suppressAlert)` in [`00_Code.gs`](./00_Code.gs) is the **single master entry point** for provisioning, repairing, or upgrading the entire system.

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
- **Eisenhower Matrix Todo Command Section**:
  - 6 dedicated KPI cards: `TOTAL TODOS`, `🔴 Q1: DO`, `🔵 Q2: SCHEDULE`, `🟡 Q3: DELEGATE`, `⚪ Q4: DON'T DO`, and `🎯 FOCUS RATIO` (Q1+Q2 share).
- **Clean Canvas**: Seamless white surface with gridlines safely hidden via `setSheetGridlinesHidden`.

### 2. Eisenhower Matrix Todo Tracker
The dedicated `Todo` sheet implements the Eisenhower decision framework:
- **Row 1 Live Stats Bar**: Frozen at the top displaying real-time formula-driven task counters for `Total Tasks`, `Q1`, `Q2`, `Q3`, and `Q4`.
- **Row 2 Headers**: `Task Name`, `Project`, `Q1: Do`, `Q2: Schedule`, `Q3: Delegate`, `Q4: Don't Do`.
- **Mutually Exclusive Radio Checkboxes**: When a quadrant checkbox is ticked, `onEdit(e)` automatically unticks any previous quadrant for that row.
- **Project Dropdown**: Validated against the `Lists` sheet.

### 3. Monthly Task Tracking & Quick Entry
- **Standard Schema**: `Date`, `Day`, `Project`, `Task`, `Category`, `Priority`, `Status`.
- **Fast Batch Task Logging Modal (`09_TaskDialog.html`)**:
  - Accessible via **WorkSheet** > **Tasks** > **Fill Task for Today** or **Fill Task for Selected Date**.
  - Select Project, Category, and Priority once; enter or paste multiple tasks into a textarea (one per line).
  - Automatically defaults to `Pending` status.
  - Includes `Ctrl + Enter` shortcut and `+ Add & Next Project` continuous entry.
- **Safe Row Insertion**: Fills the first empty task row for that date or inserts subsequent rows preserving date formatting and dropdown rules.

### 4. Daily Status Report (DSR) Engine
- Accessible via **WorkSheet** > **DSR** > **Generate DSR for Today** or **Generate DSR for Selected Date**.
- Automatically groups tasks by project into:
  1. `Tasks Completed`
  2. `Work in Progress`
  3. `Blockers / Issues`
  4. `Plan for Next Working Day`
- Styled interactive modal (`04_DSRDialog.html`) with live date switching, **Copy to Clipboard**, **Save to Sheet**, and **Download .txt**.

### 5. Visual Calendar Picker
- Accessible via **WorkSheet** > **Tasks** > **Fill Task for Selected Date** or **WorkSheet** > **DSR** > **Generate DSR for Selected Date**.
- Interactive calendar widget (`01_CalendarPickerDialog.html`) allowing date selection with arrow keys, click, or double-click to proceed directly.

### 6. Unified WorkSheet Toolbar Menu
Consolidated single top-level menu hierarchy:

```text
WorkSheet
├── ⚡ Setup / Repair Work Tracker  → setupWorkTracker()
├── 📊 Open Dashboard              → openDashboard()
├── 🎯 Open Todo                   → openTodoSheet()
├── ─────────────────────────────
├── Tasks ▶
│   ├── Fill Task for Today        → fillTaskForToday()
│   └── Fill Task for Selected Date→ fillTaskForSelectedDate()
├── DSR ▶
│   ├── Generate DSR for Today     → generateDSRForToday()
│   └── Generate DSR for Selected Date → generateDSRForSelectedDate()
├── ─────────────────────────────
└── Advanced ▶
    ├── Refresh Dashboard          → refreshDashboard()
    ├── Rebuild Lists              → rebuildLists()
    ├── Create Current Month Sheet → createCurrentMonthSheet()
    ├── Enforce Single Todo Quadrant→ sanitizeAllTodoQuadrants()
    └── Populate Dummy Data        → populateDummyData()
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

## Configuration Guide (`03_Config.gs`)

All system parameters live centrally in `CONFIG` within [`03_Config.gs`](./03_Config.gs):

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

1. **Centralized Router in `00_Code.gs`**:
   - `onEdit(e)` simple trigger coordinates `handleTodoQuadrantExclusiveSelect(e)` and `updateDashboardOnChange(e)`.
   - `onChange(e)` installable trigger listens for structural sheet changes.
2. **Duplicate Prevention**: `setupTriggers(ss)` checks `ScriptApp.getProjectTriggers()` to guarantee only one `onChange` trigger exists.
3. **Silent Execution**: Background updates run with `suppressAlert = true` and `keepActiveSheet = true` so active user workflows are never interrupted.

---

## Developer Guidelines

- **Configuration First**: Never hardcode colors, dimensions, or list items; always declare them in `03_Config.gs`.
- **Batch Processing**: Use bulk `getValues()` and `setValues()`; avoid spreadsheet API calls inside iteration loops.
- **Safe Gridlines**: Always use `setSheetGridlinesHidden(sheet, hidden)` in `10_Utils.gs`. Never call `sheet.setHideGridlines()` as it is not a native method on `Sheet`.
- **Preserve User Data**: Setup and maintenance scripts must always treat existing user sheets non-destructively.
