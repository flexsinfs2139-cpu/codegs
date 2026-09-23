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
| [`WorkSheet/12_Dashboard.gs`](./WorkSheet/12_Dashboard.gs) | Todo & Month status overview | `refreshDashboard`, `writeDashboardHeader`, `writeTodoSection`, `writeMonthStatusSection` |
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

---

## 4. Development & Coding Guidelines

- **Configuration First**: Never hardcode list values or colors; always define them inside `CONFIG` in [`WorkSheet/03_Config.gs`](./WorkSheet/03_Config.gs).
- **Batch Operations**: Use `setValues()` and `getValues()` in bulk; avoid calling Google Sheets APIs inside loops.
- **Timezone Handling**: Always derive the timezone via `getTimezone()` (`getSpreadsheetTimeZone()`) rather than assuming UTC or local server time.
- **Sheet Bounds**: Always run `trimSheet()` on newly inserted sheets to remove surplus empty rows and columns for sheet responsiveness.
- **Data Validation Integrity**: When modifying `CONFIG.LISTS`, always execute `createListsSheet()` or `rebuildLists()` to refresh Google Sheets Named Ranges.
