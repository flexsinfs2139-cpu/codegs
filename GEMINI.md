# Antigravity Context & Project Memory — WorkSheet

This file serves as persistent memory and context for Antigravity (the equivalent of `CLAUDE.md` / `.claude` in Claude). It is automatically discovered and loaded into the agent's context window on every turn to retain conversation history, architectural decisions, and project conventions across sessions.

---

## 1. Project Summary

- **Project**: WorkSheet (Google Apps Script)
- **Host Platform**: Google Sheets
- **Location**: `/home/rvkt/Documents/codegs/WorkSheet`
- **Purpose**: Automated work tracker managing daily tasks, project assignments, priorities, and statuses with dynamic monthly sheet generation, protected reference lists, and custom UI formatting.

---

## 2. Key Architecture & File Mapping

| File | Primary Role | Key Functions / Objects |
| :--- | :--- | :--- |
| [`WorkSheet/00_Code.gs`](./WorkSheet/00_Code.gs) | Main entry points | `initializeWorkTracker`, `setupWorkTracker`, `rebuildLists` |
| [`WorkSheet/01_CalendarPicker.gs`](./WorkSheet/01_CalendarPicker.gs) | Visual calendar date picker | `openCalendarPicker`, `proceedFromCalendar`, `getCalendarPickerHtml` |
| [`WorkSheet/02_ConditionalFormatting.gs`](./WorkSheet/02_ConditionalFormatting.gs) | Dynamic color rules | `setupConditionalFormatting`, `setupWeekendFormatting` |
| [`WorkSheet/03_Config.gs`](./WorkSheet/03_Config.gs) | Central configuration | `CONFIG` (Headers, Lists, Colors, Row Heights) |
| [`WorkSheet/04_DSR.gs`](./WorkSheet/04_DSR.gs) | Daily Status Report engine | `generateDSRForToday`, `generateDSRForSelectedDate`, `showDSRDialog`, `saveDSRToSheet` |
| [`WorkSheet/05_Formatting.gs`](./WorkSheet/05_Formatting.gs) | Visual layout & styling | `formatWorkTracker`, `formatHeader`, `formatColumns`, `formatDimensions`, `formatBorders` |
| [`WorkSheet/06_Lists.gs`](./WorkSheet/06_Lists.gs) | Reference lists & protection | `createListsSheet`, `ensureListsSheet`, `createNamedRanges`, `protectListsSheet` |
| [`WorkSheet/07_Menu.gs`](./WorkSheet/07_Menu.gs) | Toolbar UI menus | `onOpen` (creates `Month Sheet`, `Setup`, `Tasks`, `DSR`) |
| [`WorkSheet/08_MonthSheet.gs`](./WorkSheet/08_MonthSheet.gs) | Monthly sheet engine | `createCurrentMonthSheet`, `createMonthRows` |
| [`WorkSheet/09_Tasks.gs`](./WorkSheet/09_Tasks.gs) | Task operations | `setupDropdowns`, `createDropdownRule`, `addTaskRow`, `clearTasks` |
| [`WorkSheet/10_Utils.gs`](./WorkSheet/10_Utils.gs) | Helper utilities | `trimSheet`, `getSpreadsheet`, `getTimezone`, `getToday` |
| [`WorkSheet/11_Todo.gs`](./WorkSheet/11_Todo.gs) | Eisenhower Matrix Todo tracker | `createTodoSheet`, `ensureTodoSheet`, `setupTodoStructure`, `formatTodoSheet`, `setupTodoCheckboxes`, `setupTodoDropdowns` |
| [`WorkSheet/12_Dashboard.gs`](./WorkSheet/12_Dashboard.gs) | Modern SaaS Command Center | `refreshDashboard`, `renderDashboardHeader`, `renderKpiCards`, `renderTablesSection`, `renderTodoSection` |
| [`WorkSheet/13_SampleData.gs`](./WorkSheet/13_SampleData.gs) | Sample & dummy data engine | `populateDummyData`, `addDummyData`, `populateTodoDummyData`, `populateMonthDummyData` |
| [`WorkSheet/README.md`](./WorkSheet/README.md) | User & developer documentation | Comprehensive documentation of the WorkSheet system |

---

## 3. Retained Conversation Context & Prior Review Findings

### What Has Been Completed:
1. **Full Codebase Review**: Completed a deep-dive analysis of all Apps Script files.
2. **Documentation**: Authored the full [`WorkSheet/README.md`](./WorkSheet/README.md).
3. **DSR Generation & Modal Dialog**:
   - Implemented `04_DSR.gs` to extract tasks by date, categorize by project into Tasks Completed, Work in Progress, Blockers/Issues, and Plan for Next Working Day.
   - Built styled HTML modal dialog matching the user's UI specification (`Copy to Clipboard`, `Save to Sheet`, `Download .txt`, `Close`).
   - Updated `07_Menu.gs` to have the two requested options: `Generate DSR for today` and `Generate DSR for selected date in the month`.
4. **Streamlined Task Entry & Batch Insertion**:
   - Removed `Notes` column across the entire workbook (`03_Config.gs`, `08_MonthSheet.gs`, `05_Formatting.gs`, `02_ConditionalFormatting.gs`, `09_Tasks.gs`).
   - Unified schema to 7 columns: `Date`, `Day`, `Project`, `Task`, `Category`, `Priority`, `Status`.
   - All newly generated rows and added tasks default automatically to `Pending` status.
   - Built a fast, intuitive modal dialog for task entry (`Tasks -> Fill task for today` and `Fill task for selected date in the current month`).
   - Allows selecting `Project`, `Category`, and `Priority` once, and entering/pasting multiple tasks in a textarea (one per line, auto-cleaning bullet points).
   - Includes `+ Add & Next Project` button to immediately log tasks for another project without closing the modal.
   - Preserves keyboard shortcut `Ctrl + Enter` for instant submission.
   - Automatically populates Date and Day, updates the existing date row or inserts subsequent rows with the same Date/Day, and applies formatting and dropdowns.
5. **Resolved Architectural Review Items**:
   - Date Format Divergence resolved: Added unified `CONFIG.DATE_FORMAT: 'MMdd'`.
   - `clearTasks()` Target Safety resolved: Clears columns 3-7 on month sheets and is blocked on `Lists`, `Todo`, and `Dashboard` sheets.
6. **Visual Calendar Date Picker**:
   - Implemented `01_CalendarPicker.gs` providing an interactive monthly calendar widget for choosing dates apart from today.
   - Connected to both `Tasks -> Fill task for selected date in the current month` and `DSR -> Generate DSR for selected date in the month`.
   - Added native calendar datepicker input directly into the Task Entry modal and the DSR modal (allowing real-time report refreshing across dates).
7. **Two-Digit Alphabetical File Ordering**:
   - Standard 0-prefixed two-digit numbering (`00_Code.gs` through `12_Dashboard.gs`) ensuring structured alphabetical and numerical ordering in the Google Apps Script IDE.
8. **Eisenhower Matrix Todo Sheet & Live Stats Row**:
   - Built `11_Todo.gs` implementing the Eisenhower Matrix decision framework.
   - Table Schema (Row 2 headers): `Task Name`, `Project`, `Q1: Do`, `Q2: Schedule`, `Q3: Delegate`, `Q4: Don't Do`.
   - Row 1 Live Stats Bar: Frozen at the top displaying real-time formula-driven task counters for `Total Tasks`, `Q1 (Do)`, `Q2 (Schedule)`, `Q3 (Delegate)`, and `Q4 (Don't Do)`.
   - Native Google Sheets checkboxes for all 4 quadrants with conditional formatting colors (Q1 soft red, Q2 soft blue, Q3 soft yellow, Q4 soft gray).
   - Project dropdown validation from `Lists` sheet.
   - Integrated into `12_Dashboard.gs` with quadrant counts and interactive checkboxes.
9. **Mutually Exclusive Quadrant Radio-Button Behavior**:
   - Implemented `onEdit(e)` trigger in `11_Todo.gs` ensuring only one Eisenhower quadrant column (Q1, Q2, Q3, or Q4) can be selected per task row, automatically unchecking previous selections.
10. **Silent Initialization & Integrated Todo Setup**:
    - `initializeWorkTracker()` in `00_Code.gs` automatically provisions `Lists`, `Todo` (with Eisenhower Matrix), and the `Current Month` sheet silently without any blocking alert dialogs (`suppressAlert = true`).
11. **Typography Standards (`Varela Round` & `Roboto Mono`)**:
    - Centralized in `CONFIG.FONTS = { TEXT: 'Varela Round', DIGITS: 'Roboto Mono' }`.
    - Applied universally across all sheets: `Varela Round` for titles, headers, labels, and text descriptions; `Roboto Mono` for numbers, KPI digits, dates, percentages, and live counters.
12. **Real-Time Dashboard Auto-Update Engine**:
    - Centralized global `onEdit(e)` and `onChange(e)` triggers in `00_Code.gs`.
    - Implemented `updateDashboardOnChange(e)` in `12_Dashboard.gs` to automatically refresh the Dashboard whenever edits occur across Month sheets, Todo sheet, or Lists sheet.
    - Operates silently (`suppressAlert = true`) and preserves the user's active sheet (`keepActiveSheet = true`) so data entry is uninterrupted.
    - Synchronizes edits made directly on the Dashboard's Todo Backlog table back to the `Todo` sheet.
    - Linked `saveTasksBatch()` and `clearTasks()` in `09_Tasks.gs` and `initializeWorkTracker()` in `00_Code.gs` for programmatic freshness.
13. **Comprehensive Sample/Dummy Data Seeding & Todo Backlog Horizontal Spacing**:
    - Built [`13_SampleData.gs`](./WorkSheet/13_SampleData.gs) to seed realistic dummy data across Month and Todo sheets with 1 click.
    - Month sheet seeded with 21 engineering tasks spanning all projects, categories, priorities, and statuses (`Completed`, `In Progress`, `Blocked`, `Pending`).
    - Todo sheet seeded with 12 prioritized tasks mapped to Eisenhower Matrix quadrants (Q1: 3, Q2: 4, Q3: 3, Q4: 2) with strict mutual exclusivity.
    - Added `Setup -> Populate Dummy Data` menu action in `07_Menu.gs`.
    - Solved Todo Backlog horizontal cramping on Dashboard: increased column widths (Col A: 280px, Col B: 130px, Col C: 105px, Cols D–F: 120px) preventing any header overlapping or truncation.
    - Aligned Todo Backlog banner and table cleanly to 6 columns (A–F) with subtle borders and left-aligned task/project headers.
14. **Clean Canvas UI (Hidden Gridlines & Formula Bar Architecture)**:
    - Automatically hides gridlines across Dashboard, Todo, and Month sheets via `sheet.setHideGridlines(true)`.
    - Added subtle `#e2e8f0` row borders to Monthly Breakdown, Project Performance, and Todo Backlog so data is structured and readable on a seamless canvas.
    - Added `View -> Hide Gridlines (All Sheets)` and `View -> Show Gridlines (All Sheets)` toolbar menu actions in `07_Menu.gs` and `10_Utils.gs`.
    - Documented Google Sheets browser UI architecture for the formula bar: formula bar visibility is a user-level browser setting (toggled via `View -> Show -> Formula bar` or `Ctrl + Shift + F` for Full Screen mode) which Google Apps Script cannot access via API.
15. **Dashboard Streamlining (Todo Backlog Table Removal)**:
    - Removed the redundant Todo Backlog table from `12_Dashboard.gs` to eliminate duplicate task rows, conflicting checkbox states, and artificial column bloating.
    - Preserved the high-level Eisenhower Matrix KPI summary cards (`Total Todos`, `Q1: Do`, `Q2: Schedule`, `Q3: Delegate`, `Q4: Don't Do`, `Focus Ratio`).
    - Standardized column widths for the Monthly Breakdown (Cols A–G) and Project Performance (Cols I–M) tables to balanced proportions (Cols A: 130px, B–F: 80px, G: 95px, H: 45px, I: 150px, J–L: 80px, M: 95px).
    - Converted direct edits on the Dashboard to be cleanly ignored since active tasks and single-selection checkbox toggles are managed on the dedicated `Todo` sheet.
16. **Removed View Toolbar Menu**:
    - Removed the `View` menu (`Hide Gridlines (All Sheets)`, `Show Gridlines (All Sheets)`) from `07_Menu.gs` to keep the custom Google Sheets menu toolbar concise and clean. Automatic gridline hiding remains handled directly by sheet generators.
17. **Safe Gridline Toggling (`setSheetGridlinesHidden`)**:
    - Replaced unsupported `sheet.setHideGridlines()` calls with the safe helper `setSheetGridlinesHidden(sheet, hidden)` in `10_Utils.gs`.
    - Leverages the Google Sheets Advanced API (`Sheets.Spreadsheets.batchUpdate`) if enabled, and fails gracefully without throwing `TypeError: sheet.setHideGridlines is not a function` in standard Apps Script runtimes.

---

## 4. Development & Coding Guidelines

- **Configuration First**: Never hardcode list values or colors; always define them inside `CONFIG` in [`WorkSheet/03_Config.gs`](./WorkSheet/03_Config.gs).
- **Batch Operations**: Use `setValues()` and `getValues()` in bulk; avoid calling Google Sheets APIs inside loops.
- **Timezone Handling**: Always derive the timezone via `getTimezone()` (`getSpreadsheetTimeZone()`) rather than assuming UTC or local server time.
- **Sheet Bounds**: Always run `trimSheet()` on newly inserted sheets to remove surplus empty rows and columns for sheet responsiveness.
- **Data Validation Integrity**: When modifying `CONFIG.LISTS`, always execute `createListsSheet()` or `rebuildLists()` to refresh Google Sheets Named Ranges.
