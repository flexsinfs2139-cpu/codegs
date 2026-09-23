# Antigravity Context & Project Memory — WorkSheet

This file serves as persistent memory and context for Antigravity (the equivalent of `CLAUDE.md` / `.claude` in Claude). It is automatically discovered and loaded into the agent's context window on every turn to retain conversation history, architectural decisions, and project conventions across sessions.

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
| [`WorkSheet/07_Menu.gs`](./WorkSheet/07_Menu.gs) | Toolbar UI menus | `onOpen` (unified `WorkSheet` menu with `Tasks`, `DSR`, and `Advanced`) |
| [`WorkSheet/08_MonthSheet.gs`](./WorkSheet/08_MonthSheet.gs) | Monthly sheet engine | `createCurrentMonthSheet`, `createMonthRows` |
| [`WorkSheet/09_Tasks.gs`](./WorkSheet/09_Tasks.gs) | Task operations & batch insertion | `fillTaskForToday`, `fillTaskForSelectedDate`, `setupDropdowns`, `saveTasksBatch`, `clearTasks` |
| [`WorkSheet/09_TaskDialog.html`](./WorkSheet/09_TaskDialog.html) | Multi-task batch entry modal UI | Dedicated HTML template with `Ctrl + Enter` and `+ Add & Next Project` |
| [`WorkSheet/10_Utils.gs`](./WorkSheet/10_Utils.gs) | Helper utilities | `trimSheet`, `getSpreadsheet`, `getTimezone`, `getToday`, `setSheetGridlinesHidden` |
| [`WorkSheet/11_Todo.gs`](./WorkSheet/11_Todo.gs) | Eisenhower Matrix Todo tracker | `createTodoSheet`, `ensureTodoSheet`, `setupTodoStructure`, `formatTodoSheet`, `handleTodoQuadrantExclusiveSelect` |
| [`WorkSheet/12_Dashboard.gs`](./WorkSheet/12_Dashboard.gs) | Modern SaaS Command Center (14 Cols) | `refreshDashboard`, `updateDashboardOnChange`, `renderKpiCards`, `renderTablesSection`, `renderTodoSection` |
| [`WorkSheet/13_SampleData.gs`](./WorkSheet/13_SampleData.gs) | Sample & dummy data engine | `populateDummyData`, `populateTodoDummyData`, `populateMonthDummyData` |
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
   - Removed `Notes` column across the entire workbook (`03_Config.gs`, `08_MonthSheet.gs`, `05_Formatting.gs`, `02_ConditionalFormatting.gs`, `09_Tasks.gs`).
   - Unified schema to 7 columns: `Date`, `Day`, `Project`, `Task`, `Category`, `Priority`, `Status`.
   - All newly generated rows and added tasks default automatically to `Pending` status.
   - Preserves keyboard shortcut `Ctrl + Enter` for instant submission and `+ Add & Next Project` workflow.
5. **Resolved Architectural Review Items**:
   - Date Format Divergence resolved: Added unified `CONFIG.DATE_FORMAT: 'MMdd'`.
   - `clearTasks()` Target Safety resolved: Clears columns 3-7 on month sheets and is blocked on `Lists`, `Todo`, and `Dashboard` sheets.
6. **Visual Calendar Date Picker**:
   - Implemented `01_CalendarPicker.gs` and `01_CalendarPickerDialog.html` providing an interactive monthly calendar widget for choosing dates apart from today.
   - Connected to both Task entry and DSR report generation with live in-modal date switching.
7. **Two-Digit Alphabetical File Ordering**:
   - Standard 0-prefixed two-digit numbering (`00_Code.gs` through `13_SampleData.gs`) ensuring structured alphabetical and numerical ordering in the Google Apps Script IDE.
8. **Eisenhower Matrix Todo Sheet & Live Stats Row**:
   - Built `11_Todo.gs` implementing the Eisenhower Matrix decision framework.
   - Table Schema (Row 2 headers): `Task Name`, `Project`, `Q1: Do`, `Q2: Schedule`, `Q3: Delegate`, `Q4: Don't Do`.
   - Row 1 Live Stats Bar: Frozen at the top displaying real-time formula-driven task counters for `Total Tasks`, `Q1 (Do)`, `Q2 (Schedule)`, `Q3 (Delegate)`, and `Q4 (Don't Do)`.
   - Native Google Sheets checkboxes for all 4 quadrants with conditional formatting colors.
   - Project dropdown validation from `Lists` sheet.
9. **Mutually Exclusive Quadrant Radio-Button Behavior**:
   - Implemented `onEdit(e)` trigger in `11_Todo.gs` ensuring only one Eisenhower quadrant column (Q1, Q2, Q3, or Q4) can be selected per task row, automatically unchecking previous selections.
10. **Typography Standards (`Varela Round` & `Roboto Mono`)**:
    - Centralized in `CONFIG.FONTS = { TEXT: 'Varela Round', DIGITS: 'Roboto Mono' }`.
    - Applied universally across all sheets: `Varela Round` for titles, headers, labels, and text descriptions; `Roboto Mono` for numbers, KPI digits, dates, percentages, and live counters.
11. **Real-Time Dashboard Auto-Update Engine**:
    - Centralized global `onEdit(e)` and `onChange(e)` triggers in `00_Code.gs`.
    - Implemented `updateDashboardOnChange(e)` in `12_Dashboard.gs` to automatically refresh the Dashboard whenever edits occur across Month sheets, Todo sheet, or Lists sheet.
    - Operates silently (`suppressAlert = true`) and preserves the user's active sheet (`keepActiveSheet = true`) so data entry is uninterrupted.
12. **Safe Gridline Toggling (`setSheetGridlinesHidden`)**:
    - Replaced unsupported `sheet.setHideGridlines()` calls with the safe helper `setSheetGridlinesHidden(sheet, hidden)` in `10_Utils.gs`.
    - Leverages the Google Sheets Advanced API (`Sheets.Spreadsheets.batchUpdate`) if enabled, and fails gracefully without throwing `TypeError: sheet.setHideGridlines is not a function` in standard Apps Script runtimes.
13. **Dedicated Production HTML Templates**:
    - Separated presentation layer from `.gs` scripts into standalone, syntax-highlighted `.html` templates (`01_CalendarPickerDialog.html`, `09_TaskDialog.html`, `04_DSRDialog.html`).
    - Harmonized typography with `Varela Round` and `Roboto Mono` via Google Fonts.
    - Resilient fallback loaders: checks for template files first, falling back smoothly to inline generator if deployed as `.gs` only.
14. **Modern SaaS Work Tracker Re-Architecture (`setupWorkTracker`)**:
    - Created single master idempotent setup function `setupWorkTracker()` in `00_Code.gs`.
    - Non-destructive execution: configures/upgrades the entire system (Lists, Todo, Month, Dashboard, formatting, validations, triggers) without data loss.
    - Structured logging pipeline: `[WorkSheet] Setup started`, `[WorkSheet] Lists verified`, `[WorkSheet] Todo verified`, `[WorkSheet] Current month verified`, `[WorkSheet] Dashboard verified`, `[WorkSheet] Validations & Formatting configured`, `[WorkSheet] Triggers verified`, `[WorkSheet] Setup completed`.
    - Centralized trigger setup (`setupTriggers`) with duplicate detection via `ScriptApp.getProjectTriggers()`.
    - Executive 14-Column Command Center in `12_Dashboard.gs`: 6 KPI cards displaying grand totals and today's counts as secondary subtitles, side-by-side Monthly Breakdown (Cols 1–7) and 6-column Project Performance (Cols 9–14: `Project`, `Total`, `Completed`, `In Progress`, `Blocked`, `Completion %`) sorted deterministically, plus 6 Eisenhower Todo cards.
    - Unified single top-level `WorkSheet` toolbar menu in `07_Menu.gs` with structured submenus (`Tasks`, `DSR`, `Advanced`).
    - Idempotent Month Sheet upgrade in `08_MonthSheet.gs` preserving row 2+ user tasks.

---

## 4. Development & Coding Guidelines

- **Configuration First**: Never hardcode list values, dimensions, or colors; always define them inside `CONFIG` in [`WorkSheet/03_Config.gs`](./WorkSheet/03_Config.gs).
- **Batch Operations**: Use `setValues()` and `getValues()` in bulk; avoid calling Google Sheets APIs inside loops.
- **Timezone Handling**: Always derive the timezone via `getTimezone()` (`getSpreadsheetTimeZone()`) rather than assuming UTC or local server time.
- **Sheet Bounds**: Always run `trimSheet()` on newly inserted sheets to remove surplus empty rows and columns for sheet responsiveness.
- **Data Validation Integrity**: When modifying `CONFIG.LISTS`, always execute `createListsSheet()` or `rebuildLists()` to refresh Google Sheets Named Ranges.
- **Safe Gridlines**: Always use `setSheetGridlinesHidden(sheet, hidden)` in `10_Utils.gs`. Never invoke `sheet.setHideGridlines()`.
