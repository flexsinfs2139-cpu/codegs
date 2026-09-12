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
| [`WorkSheet/01_Config.gs`](./WorkSheet/01_Config.gs) | Central configuration | `CONFIG` (Headers, Lists, Colors, Row Heights) |
| [`WorkSheet/02_Utils.gs`](./WorkSheet/02_Utils.gs) | Helper utilities | `trimSheet`, `getSpreadsheet`, `getTimezone`, `getToday` |
| [`WorkSheet/03_Lists.gs`](./WorkSheet/03_Lists.gs) | Reference lists & protection | `createListsSheet`, `ensureListsSheet`, `createNamedRanges`, `protectListsSheet` |
| [`WorkSheet/04_MonthSheet.gs`](./WorkSheet/04_MonthSheet.gs) | Monthly sheet engine | `createCurrentMonthSheet`, `createMonthRows` |
| [`WorkSheet/05_Tasks.gs`](./WorkSheet/05_Tasks.gs) | Task operations | `setupDropdowns`, `createDropdownRule`, `addTaskRow`, `clearTasks` |
| [`WorkSheet/06_Formatting.gs`](./WorkSheet/06_Formatting.gs) | Visual layout & styling | `formatWorkTracker`, `formatHeader`, `formatColumns`, `formatDimensions`, `formatBorders` |
| [`WorkSheet/07_ConditionalFormatting.gs`](./WorkSheet/07_ConditionalFormatting.gs) | Dynamic color rules | `setupConditionalFormatting`, `setupWeekendFormatting` |
| [`WorkSheet/08_Menu.gs`](./WorkSheet/08_Menu.gs) | Toolbar UI menus | `onOpen` (creates `Month Sheet`, `Setup`, `Tasks`, `DSR`) |
| [`WorkSheet/09_DSR.gs`](./WorkSheet/09_DSR.gs) | Daily Status Report engine | `generateDSRForToday`, `generateDSRForSelectedDate`, `showDSRDialog`, `saveDSRToSheet` |
| [`WorkSheet/10_CalendarPicker.gs`](./WorkSheet/10_CalendarPicker.gs) | Visual calendar date picker | `openCalendarPicker`, `proceedFromCalendar`, `getCalendarPickerHtml` |
| [`WorkSheet/README.md`](./WorkSheet/README.md) | User & developer documentation | Comprehensive documentation of the WorkSheet system |

---

## 3. Retained Conversation Context & Prior Review Findings

### What Has Been Completed:
1. **Full Codebase Review**: Completed a deep-dive analysis of all Apps Script files.
2. **Documentation**: Authored the full [`WorkSheet/README.md`](./WorkSheet/README.md).
3. **DSR Generation & Modal Dialog**:
   - Implemented `09_DSR.gs` to extract tasks by date, categorize by project into Tasks Completed, Work in Progress, Blockers/Issues, and Plan for Next Working Day.
   - Built styled HTML modal dialog matching the user's UI specification (`Copy to Clipboard`, `Save to Sheet`, `Download .txt`, `Close`).
   - Updated `08_Menu.gs` to have the two requested options: `Generate DSR for today` and `Generate DSR for selected date in the month`.
4. **Streamlined Task Entry & Batch Insertion**:
   - Removed `Notes` column across the entire workbook (`01_Config.gs`, `04_MonthSheet.gs`, `06_Formatting.gs`, `07_ConditionalFormatting.gs`, `05_Tasks.gs`).
   - Unified schema to 7 columns: `Date`, `Day`, `Project`, `Task`, `Category`, `Priority`, `Status`.
   - All newly generated rows and added tasks default automatically to `Pending` status.
   - Built a fast, intuitive modal dialog for task entry (`Tasks -> Fill task for today` and `Fill task for selected date in the current month`).
   - Allows selecting `Project`, `Category`, and `Priority` once, and entering/pasting multiple tasks in a textarea (one per line, auto-cleaning bullet points).
   - Includes `+ Add & Next Project` button to immediately log tasks for another project without closing the modal.
   - Preserves keyboard shortcut `Ctrl + Enter` for instant submission.
   - Automatically populates Date and Day, updates the existing date row or inserts subsequent rows with the same Date/Day, and applies formatting and dropdowns.
5. **Resolved Architectural Review Items**:
   - Date Format Divergence resolved: Added unified `CONFIG.DATE_FORMAT: 'MMdd'`.
   - `clearTasks()` Target Safety resolved: Clears columns 3-7 on month sheets and is blocked on the `Lists` sheet.
6. **Visual Calendar Date Picker**:
   - Implemented `10_CalendarPicker.gs` providing an interactive monthly calendar widget for choosing dates apart from today.
   - Connected to both `Tasks -> Fill task for selected date in the current month` and `DSR -> Generate DSR for selected date in the month`.
   - Added native calendar datepicker input directly into the Task Entry modal and the DSR modal (allowing real-time report refreshing across dates).
7. **Two-Digit File Ordering**:
   - Prefixed all `.gs` files with standard numeric ordering (`00_Code.gs` through `10_CalendarPicker.gs`) ensuring `Code.gs` appears first as `00_Code.gs` and all other files follow structured numerical ordering in the Google Apps Script IDE.

---

## 4. Development & Coding Guidelines

- **Configuration First**: Never hardcode list values or colors; always define them inside `CONFIG` in [`WorkSheet/01_Config.gs`](./WorkSheet/01_Config.gs).
- **Batch Operations**: Use `setValues()` and `getValues()` in bulk; avoid calling Google Sheets APIs inside loops.
- **Timezone Handling**: Always derive the timezone via `getTimezone()` (`getSpreadsheetTimeZone()`) rather than assuming UTC or local server time.
- **Sheet Bounds**: Always run `trimSheet()` on newly inserted sheets to remove surplus empty rows and columns for sheet responsiveness.
- **Data Validation Integrity**: When modifying `CONFIG.LISTS`, always execute `createListsSheet()` or `rebuildLists()` to refresh Google Sheets Named Ranges.
