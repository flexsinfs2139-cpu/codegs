# Antigravity Context & Project Memory — WorkSheet

This file serves as persistent memory and context for Antigravity (the equivalent of `CLAUDE.md` / `.claude` in Claude). It is automatically discovered and loaded into the agent's context window on every turn to retain conversation history, architectural decisions, and project conventions across sessions.

> **Related Project Memory**:
> - [Eisenhower Matrix ToDo Memory](./ToDo/GEMINI.md) — Dedicated context & architecture memory for `/home/rvkt/Documents/codegs/ToDo`.

---

## 1. Project Summary

- **Project**: WorkSheet (Google Apps Script)
- **Host Platform**: Google Sheets
- **Location**: `/home/rvkt/Documents/codegs/WorkSheet`
- **Purpose**: Automated modern SaaS work tracker managing daily tasks, project assignments, priorities, and statuses with dynamic monthly sheet generation, protected reference lists, executive dashboard, and custom UI formatting.

---

## 2. Key Architecture & File Mapping

| File | Primary Role | Key Functions / Objects |
| :--- | :--- | :--- |
| [`WorkSheet/00_Code.gs`](./WorkSheet/00_Code.gs) | Master entry points & trigger routing | `setupWorkTracker`, `initializeWorkTracker`, `rebuildLists`, `setupTriggers`, `onEdit`, `onChange` |
| [`WorkSheet/01_CalendarPicker.gs`](./WorkSheet/01_CalendarPicker.gs) | Visual calendar date picker controller | `openCalendarPicker`, `proceedFromCalendar`, `getCalendarPickerHtml` |
| [`WorkSheet/01_CalendarPickerDialog.html`](./WorkSheet/01_CalendarPickerDialog.html) | Visual calendar picker modal UI | Dedicated HTML template with Google Fonts and keyboard navigation |
| [`WorkSheet/02_ConditionalFormatting.gs`](./WorkSheet/02_ConditionalFormatting.gs) | Dynamic semantic color rules | `setupConditionalFormatting`, `setupWeekendFormatting`, `setupTodoConditionalFormatting` |
| [`WorkSheet/03_Config.gs`](./WorkSheet/03_Config.gs) | Central configuration | `CONFIG` (Headers, Lists, Colors, Dimensions, Fonts, Dashboard) |
| [`WorkSheet/04_DSR.gs`](./WorkSheet/04_DSR.gs) | Daily Status Report engine | `generateDSRForToday`, `generateDSRForSelectedDate`, `showDSRDialog`, `saveDSRToSheet` |
| [`WorkSheet/04_DSRDialog.html`](./WorkSheet/04_DSRDialog.html) | DSR interactive modal UI | Dedicated HTML template with live date switcher and export buttons |
| [`WorkSheet/05_Formatting.gs`](./WorkSheet/05_Formatting.gs) | Visual layout & styling | `formatWorkTracker`, `formatHeader`, `formatColumns`, `formatDimensions`, `formatBorders` |
| [`WorkSheet/06_Lists.gs`](./WorkSheet/06_Lists.gs) | Reference lists & protection | `createListsSheet`, `ensureListsSheet`, `createNamedRanges`, `protectListsSheet` |
| [`WorkSheet/07_Menu.gs`](./WorkSheet/07_Menu.gs) | Toolbar UI menus | `onOpen` (four top-level menus: `WorkSheet`, `Setup`, `Tasks`, `DSR`) |
| [`WorkSheet/08_MonthSheet.gs`](./WorkSheet/08_MonthSheet.gs) | Monthly sheet engine | `createCurrentMonthSheet`, `createMonthRows` |
| [`WorkSheet/09_Tasks.gs`](./WorkSheet/09_Tasks.gs) | Task operations & batch insertion | `fillTaskForToday`, `fillTaskForSelectedDate`, `setupDropdowns`, `saveTasksBatch`, `clearTasks` |
| [`WorkSheet/09_TaskDialog.html`](./WorkSheet/09_TaskDialog.html) | Multi-task batch entry modal UI | Dedicated HTML template with `Ctrl + Enter` and `+ Add & Next Project` |
| [`WorkSheet/10_Utils.gs`](./WorkSheet/10_Utils.gs) | Helper utilities | `trimSheet`, `getSpreadsheet`, `getTimezone`, `getToday`, `setSheetGridlinesHidden` |
| [`WorkSheet/11_Todo.gs`](./WorkSheet/11_Todo.gs) | Simplified 5-column Todo backlog | `createTodoSheet`, `ensureTodoSheet`, `setupTodoStructure`, `formatTodoSheet`, `ensureAllTodoRowsHaveIds`, `updateTodoTaskIdsNamedRange` |
| [`WorkSheet/12_Dashboard.gs`](./WorkSheet/12_Dashboard.gs) | Modern SaaS Command Center (14 Cols) | `refreshDashboard`, `updateDashboardOnChange`, `renderKpiCards`, `renderTablesSection`, `renderTodoSection` |
| [`WorkSheet/13_SampleData.gs`](./WorkSheet/13_SampleData.gs) | Sample & dummy data engine | `populateDummyData`, `populateTodoDummyData`, `populateMonthDummyData`, `populateDecisionsDummyData` |
| [`WorkSheet/14_Timeline.gs`](./WorkSheet/14_Timeline.gs) | Decisions & Task Timeline audit engine | `ensureDecisionsSheet`, `formatDecisionsSheet`, `generateNextTaskId`, `generateNextTimelineId`, `recordTimelineEvent`, `getTaskTimelineHistory`, `getAllTodoTasks`, `handleSpreadsheetEdit`, `submitDecisionEvent`, `loadTaskTimelineData` |
| [`WorkSheet/14_DecisionDialog.html`](./WorkSheet/14_DecisionDialog.html) | Decision & event logging modal UI | Dedicated HTML template for recording decisions linked to Task IDs |
| [`WorkSheet/14_TimelineDialog.html`](./WorkSheet/14_TimelineDialog.html) | Visual task timeline & audit modal UI | Dedicated HTML template for viewing task audit trails and decisions |
| [`WorkSheet/15_Migration.gs`](./WorkSheet/15_Migration.gs) | Non-destructive system migration engine | `runFullSystemMigration`, `migrateTodoSheet`, `migrateAllMonthSheets`, `seedBaselineTimelineEvents` |
| [`WorkSheet/README.md`](./WorkSheet/README.md) | User & developer documentation | Comprehensive documentation of the WorkSheet modern SaaS system |

---

## 3. Retained Conversation Context & Prior Review Findings

### What Has Been Completed:
1. **Full Codebase Review**: Completed a deep-dive analysis of all Apps Script files.
2. **Documentation**: Authored the full [`WorkSheet/README.md`](./WorkSheet/README.md).
3. **DSR Generation & Modal Dialog**:
   - Implemented `04_DSR.gs` to extract tasks by date, categorize by project into Tasks Completed, Work in Progress, Blockers/Issues, and Plan for Next Working Day.
   - Built styled HTML modal dialog matching the user's UI specification (`Copy to Clipboard`, `Save to Sheet`, `Download .txt`, `Close`).
4. **Streamlined Task Entry & Batch Insertion**:
   - Schema unified across monthly sheets with 8 columns: `Date`, `Day`, `Task ID`, `Project`, `Task`, `Category`, `Priority`, `Status`.
   - All newly generated rows and added tasks default automatically to `Pending` status.
   - Preserves keyboard shortcut `Ctrl + Enter` for instant submission and `+ Add & Next Project` workflow.
5. **Resolved Architectural Review Items**:
   - Date Format Divergence resolved: Added unified `CONFIG.DATE_FORMAT: 'MMdd'`.
   - `clearTasks()` Target Safety resolved: Clears task columns on month sheets and is blocked on `Lists`, `Todo`, `Decisions`, and `Dashboard` sheets.
6. **Visual Calendar Date Picker**:
   - Implemented `01_CalendarPicker.gs` and `01_CalendarPickerDialog.html` providing an interactive monthly calendar widget for choosing dates apart from today.
   - Connected to both Task entry and DSR report generation with live in-modal date switching.
7. **Two-Digit Alphabetical File Ordering**:
   - Standard 0-prefixed two-digit numbering (`00_Code.gs` through `15_Migration.gs`) ensuring structured alphabetical and numerical ordering in the Google Apps Script IDE.
8. **Simplified 5-Column Todo Backlog & Monotonic Task IDs**:
   - Built simplified `11_Todo.gs` with schema: `ID`, `Task Description`, `Project`, `Due Date`, `Notes`.
   - Removed Eisenhower Matrix quadrants, checkboxes, and quadrant formulas.
   - Strictly monotonic, permanent Task IDs (`TSK-0001`, `TSK-0002`...) stored in `PropertiesService` under script locks so IDs are never reused or mutated even if tasks are deleted.
   - Linked to monthly sheets via `Task ID` dropdown; selecting a Task ID auto-populates `Task` and `Project` while allowing ad-hoc entries.
9. **Decisions Sheet & Dual-Path Timeline Audit Engine**:
   - Implemented `14_Timeline.gs` managing 12-column `Decisions` sheet: `Timeline ID`, `Task ID`, `Event Date`, `Event Type`, `Previous Status`, `New Status`, `Decision / Event`, `Rationale`, `Decided By`, `Next Action`, `Due Date`, `Reference`.
   - Dual-path event logging: captures interactive cell edits on Month sheets (status changes, task creation) via `handleSpreadsheetEdit(e)` in `onEdit`, and script operations via direct calls to `recordTimelineEvent`.
   - Modal dialogs for manual decision entry (`14_DecisionDialog.html`) and visual task timeline exploration (`14_TimelineDialog.html`).
10. **Typography Standards (`Varela Round` & `Roboto Mono`)**:
    - Centralized in `CONFIG.FONTS = { TEXT: 'Varela Round', DIGITS: 'Roboto Mono' }`.
    - Applied universally across all sheets: `Varela Round` for titles, headers, labels, and text descriptions; `Roboto Mono` for numbers, KPI digits, dates, percentages, and live counters.
11. **Real-Time Dashboard Auto-Update Engine**:
    - Centralized global `onEdit(e)` and `onChange(e)` triggers in `00_Code.gs`.
    - Implemented `updateDashboardOnChange(e)` in `12_Dashboard.gs` to automatically refresh the Dashboard whenever edits occur across Month sheets, Todo sheet, Decisions sheet, or Lists sheet.
    - Operates silently (`suppressAlert = true`) and preserves the user's active sheet (`keepActiveSheet = true`) so data entry is uninterrupted.
12. **Safe Gridline Toggling (`setSheetGridlinesHidden`)**:
    - Replaced unsupported `sheet.setHideGridlines()` calls with the safe helper `setSheetGridlinesHidden(sheet, hidden)` in `10_Utils.gs`.
    - Leverages the Google Sheets Advanced API (`Sheets.Spreadsheets.batchUpdate`) if enabled, and fails gracefully without throwing `TypeError: sheet.setHideGridlines is not a function` in standard Apps Script runtimes.
13. **Dedicated Production HTML Templates**:
    - Separated presentation layer from `.gs` scripts into standalone, syntax-highlighted `.html` templates (`01_CalendarPickerDialog.html`, `09_TaskDialog.html`, `04_DSRDialog.html`, `14_DecisionDialog.html`, `14_TimelineDialog.html`).
    - Harmonized typography with `Varela Round` and `Roboto Mono` via Google Fonts.
    - Resilient fallback loaders: checks for template files first, falling back smoothly to inline generator if deployed as `.gs` only.
14. **Modern SaaS Work Tracker Re-Architecture (`setupWorkTracker`)**:
    - Created single master idempotent setup function `setupWorkTracker()` in `00_Code.gs`.
    - Non-destructive execution: configures/upgrades the entire system (Lists, Todo, Decisions, Month, Dashboard, formatting, validations, triggers) without data loss.
    - Structured logging pipeline: `[WorkSheet] Setup started`, `[WorkSheet] Lists verified`, `[WorkSheet] Todo verified`, `[WorkSheet] Decisions verified`, `[WorkSheet] Current month verified`, `[WorkSheet] Dashboard verified`, `[WorkSheet] Validations & Formatting configured`, `[WorkSheet] Triggers verified`, `[WorkSheet] Setup completed`.
    - Centralized trigger setup (`setupTriggers`) with duplicate detection via `ScriptApp.getProjectTriggers()`.
    - Executive 14-Column Command Center in `12_Dashboard.gs`: 6 KPI cards, side-by-side Monthly Breakdown (Cols 1–7) and 6-column Project Performance (Cols 9–14), plus 5 Backlog Schedule KPI cards.
    - Modular Four Top-Level Toolbar Menus in `07_Menu.gs`: `WorkSheet` (Hub/Navigation), `Setup` (combined Setup, Migration & Advanced admin tools), `Tasks` (daily task logging), and `DSR` (daily status reporting).
15. **Non-Destructive Migration Engine (`15_Migration.gs`)**:
    - Provides `runFullSystemMigration()`, `migrateTodoSheet()`, `migrateAllMonthSheets()`, and `seedBaselineTimelineEvents()`.
    - Detects older schemas (e.g. Eisenhower Todo sheets or 7-column month sheets), inserts Task IDs, backfills IDs safely, provisions `Decisions`, and seeds baseline creation events.

---

## 4. Development & Coding Guidelines

- **Configuration First**: Never hardcode list values, dimensions, or colors; always define them inside `CONFIG` in [`WorkSheet/03_Config.gs`](./WorkSheet/03_Config.gs).
- **Batch Operations**: Use `setValues()` and `getValues()` in bulk; avoid calling Google Sheets APIs inside loops.
- **Timezone Handling**: Always derive the timezone via `getTimezone()` (`getSpreadsheetTimeZone()`) rather than assuming UTC or local server time.
- **Sheet Bounds**: Always run `trimSheet()` on newly inserted sheets to remove surplus empty rows and columns for sheet responsiveness.
- **Data Validation Integrity**: When modifying `CONFIG.LISTS`, always execute `createListsSheet()` or `rebuildLists()` to refresh Google Sheets Named Ranges.
- **Safe Gridlines**: Always use `setSheetGridlinesHidden(sheet, hidden)` in `10_Utils.gs`. Never invoke `sheet.setHideGridlines()`.
