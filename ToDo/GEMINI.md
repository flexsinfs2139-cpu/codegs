# Antigravity Context & Project Memory — Eisenhower Matrix ToDo

This file serves as persistent memory and context for Antigravity (the equivalent of `CLAUDE.md` / `.claude` in Claude). It is automatically discovered and loaded into the agent's context window on every turn to retain conversation history, architectural decisions, and project conventions across sessions.

---

## 1. Project Summary

- **Project**: Eisenhower Matrix ToDo (Google Apps Script)
- **Host Platform**: Google Sheets
- **Location**: `/home/rvkt/Documents/codegs/ToDo`
- **Purpose**: Modern SaaS single-sheet Eisenhower Matrix workspace featuring dynamic task management (`Do Now`, `Schedule`, `Delegate`, `Delete`), live formula-filtered quadrant columns, interactive task sidebar with live polling, automatic status tracking, rich vibrant SaaS color system, archive engine, and non-destructive rebuilds.

---

## 2. Key Architecture & File Mapping

| File | Primary Role | Key Functions / Objects |
| :--- | :--- | :--- |
| [`ToDo/00_Code.gs`](./00_Code.gs) | Master entry points, trigger routing & workspace setup | `onOpen`, `onEdit`, `autoOpenTaskPanel`, `setupEisenhowerMatrix`, `removeLegacySheets_`, `syncTaskPanelTrigger_` |
| [`ToDo/01_Config.gs`](./01_Config.gs) | Central configuration, theme colors & schema | `CONFIG`, `THEME`, `COLUMNS`, `COL`, `STATUSES`, `QUADRANTS`, `SAMPLE_TASKS` |
| [`ToDo/02_Todo.gs`](./02_Todo.gs) | TODO sheet engine & table builder | `rebuildTodoSheet_`, `setupTodoSheet_`, `todoFormatRules_`, `handleTodoEdit_`, `fillDefaultStatus_`, `statusValidation_` |
| [`ToDo/05_Archive.gs`](./05_Archive.gs) | Archiving engine & historical task storage | `archiveCompletedTasks`, `appendToArchive_`, `rebuildArchiveSheet_`, `archiveFormatRules_`, `deleteTaskRows_` |
| [`ToDo/06_Formulas.gs`](./06_Formulas.gs) | Dynamic formula & conditional formatting builders | `quadrantListFormula_`, `taskCol_`, `statusRules_`, `priorityRules_`, `taskQuadrantHighlightRules_`, `quadrantColumnRules_` |
| [`ToDo/07_Utils.gs`](./07_Utils.gs) | Generic sheet manipulation helpers | `getOrCreateSheet_`, `resetSheet_`, `fitSheet_`, `lastContentRow_`, `migrateColumns_`, `styleTable_`, `insertCheckboxes_`, `columnLetter_`, `orderSheets_`, `removeBlankDefaultSheet_` |
| [`ToDo/08_TaskPanel.gs`](./08_TaskPanel.gs) | Task details sidebar controller | `showTaskPanel`, `getSelectedTask`, `saveTask`, `createTask`, `openTodoSheet`, `readTask_`, `findTaskRow_`, `firstEmptyRow_` |
| [`ToDo/08_TaskPanelSidebar.html`](./08_TaskPanelSidebar.html) | Interactive task details panel UI | Embedded responsive sidebar with live debounced autosave, keyboard shortcuts (`Ctrl + Enter`), colored quadrant badges, and priority toggles |

---

## 3. Retained Conversation Context & Architectural Decisions

### 1. Single-Sheet Eisenhower Layout Architecture
- **Unified Workspace**: Replaced legacy multi-tab structure (`Eisenhower Matrix`, `Dashboard`) with a single high-efficiency sheet (`TODO`).
- **Core Task Table (Cols A:D / 1–4)**:
  - `Tasks` (Width 340px): Free-form task name with automatic line break removal.
  - `Status` (Width 105px): Dropdown validated (`To Do`, `In Progress`, `Done`, `Blocked`). New tasks automatically default to `To Do`.
  - `Important` (Width 75px): Native Google Sheets checkbox.
  - `Urgent` (Width 75px): Native Google Sheets checkbox.
- **Visual Spacer (Col E / 5)**:
  - 16px wide column (`THEME.spacerBg: '#f8fafc'`) providing a clean breathing margin between the input table and the matrix.
- **Live Quadrant Columns (Cols F:I / 6–9)**:
  - `Q1 · Do Now` (Urgent=TRUE, Important=TRUE)
  - `Q2 · Schedule` (Urgent=FALSE, Important=TRUE)
  - `Q3 · Delegate` (Urgent=TRUE, Important=FALSE)
  - `Q4 · Delete` (Urgent=FALSE, Important=FALSE)
  - Populated dynamically via `=IFERROR(FILTER($A$2:$A, ...), "")`. Protected with warning against accidental manual editing.
  - Clicking any task in a quadrant column automatically detects and loads the source task row in the Task Details sidebar.

### 2. Vibrant Modern SaaS Color System & Dimensions
- **Executive Slate Header & Clean Grid**:
  - Task Table header: `#0f172a` (Modern Slate 900) with crisp `#ffffff` text.
  - Generous dimensions: 38px executive header height, 26px data row height for comfortable SaaS padding.
  - Typography: Modern geometric sans-serif (`Inter`, `Plus Jakarta Sans`).
  - Grid borders: `#cbd5e1` (Clean Slate 300) replacing harsh pure-black lines.
  - Sheet Tab colors: `#4f46e5` (Vibrant Indigo) for `TODO`, `#64748b` (Cool Slate) for `Archive`.
- **Distinct Signature Quadrant Banners & Pastel Column Tints**:
  - **Q1 (Do Now)**: Vibrant Crimson banner (`#dc2626`), Rose column tint (`#fef2f2`), Dark Rose text (`#991b1b`).
  - **Q2 (Schedule)**: Vibrant Emerald banner (`#16a34a`), Emerald column tint (`#f0fdf4`), Dark Emerald text (`#166534`).
  - **Q3 (Delegate)**: Vibrant Amber banner (`#d97706`), Amber column tint (`#fffbeb`), Dark Amber text (`#92400e`).
  - **Q4 (Delete)**: Vibrant Indigo banner (`#6366f1`), Lavender column tint (`#f5f3ff`), Dark Indigo text (`#4338ca`).
- **Dynamic Task Row Highlighting**:
  - Open tasks in Column A automatically adopt their corresponding quadrant's background and text colors based on checkbox state.
  - Completed (`Done`) tasks are automatically struck through, dimmed to `#94a3b8`, and tinted with `#f8fafc`.
- **Active Priority Checkbox Cell Tints**:
  - Checking `Important` (Col C) tints the cell soft warm amber (`#fef3c7`, text `#92400e`).
  - Checking `Urgent` (Col D) tints the cell soft warm rose (`#fee2e2`, text `#991b1b`).
- **Status Pills**:
  - `To Do`: Soft Indigo badge (`#e0e7ff`, text `#3730a3`).
  - `In Progress`: Royal Blue badge (`#dbeafe`, text `#1e40af`).
  - `Done`: Emerald Green badge (`#dcfce7`, text `#166534`).
  - `Blocked`: Crimson Rose badge (`#fee2e2`, text `#991b1b`).

### 3. Modern SaaS Task Details Sidebar (`08_TaskPanelSidebar.html`)
- **Typography & Theme**: Built with `Plus Jakarta Sans` and `JetBrains Mono` with clean SaaS CSS design tokens.
- **Top Brand & Sync Navbar**: Features `⚡ Eisenhower` logo badge, row counter pill (`Row #3`), and real-time pulsing sync indicator (`● Saved` / `● Saving...`).
- **Interactive 2x2 Eisenhower Decision Matrix (Visual Picker)**:
  - 4-quadrant interactive decision cards in a 2x2 grid (`Do Now`, `Schedule`, `Delegate`, `Delete`).
  - Single-click quadrant selection: clicking any tile automatically syncs both `Important` and `Urgent` states.
  - Active tile lights up with a radiant border, glowing background, and elevated shadow.
- **Priority Fine-Tuning Switches**: Granular `Important` and `Urgent` toggle switches stay in bi-directional sync with the 2x2 matrix.
- **Real-Time Polling Engine**: Polls `getSelectedTask()` every 1,000ms while open to mirror sheet row selection.
- **Bi-Directional Debounced Autosave**: Edits in the sidebar flush to the sheet after 700ms or on blur/enter without manual save buttons.
- **Keyboard Shortcuts Bar**: Visual shortcuts guide (`[Ctrl + Enter] Save` · `[Esc] Blur`).
- **Floating Action Dock**: Large primary gradient CTA (`✓ Mark Done` / `↺ Reopen Task`) and secondary `+ New Task` button.
- **Auto-Open on Spreadsheet Load**: `syncTaskPanelTrigger_` creates an installable `onOpen` trigger so the sidebar opens automatically on desktop browsers.

### 4. Safe Task Archiving Engine (`05_Archive.gs`)
- **Non-Destructive Archiving**: Finds tasks with `Status === 'Done'`, prompts for user confirmation, appends rows to `Archive` with an `Archived` timestamp column.
- **Partial Row Deletion**: Uses `deleteTaskRows_` to delete only columns A:D from the `TODO` sheet, leaving live quadrant formula columns (F:I) intact.
- **Archive Sheet Styling**: Styled with Deep Slate header (`#1e293b`), clean borders, date formatting, and status pill conditional formatting.

---

## 4. Development & Coding Guidelines

- **Configuration First**: All dimensions, colors, statuses, and quadrant definitions must reside in `CONFIG`, `THEME`, `STATUSES`, or `QUADRANTS` in [`01_Config.gs`](./01_Config.gs). Never hardcode color hexes or column widths in implementation files.
- **Batch Spreadsheet Operations**: Always read and write ranges in bulk using `getValues()` and `setValues()`. Minimize Google Apps Script API calls.
- **Non-Destructive Rebuilds**: Re-running `setupEisenhowerMatrix()` or `rebuildTodoSheet_` must preserve existing task rows, checkboxes, and notes via `migrateColumns_`.
- **Safe Grid Trimming**: Always invoke `fitSheet_` to trim unused rows and columns down to active rows + `CONFIG.spareRows`.
- **Formulas & Column Indices**: Quadrant formulas in `06_Formulas.gs` depend on `COLUMNS` keys (`taskCol_('task')`, `taskCol_('important')`, etc.). Ensure `Important` and `Urgent` stay adjacent.
