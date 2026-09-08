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
| [`WorkSheet/Config.gs`](./WorkSheet/Config.gs) | Central configuration | `CONFIG` (Headers, Lists, Colors, Row Heights) |
| [`WorkSheet/Code.gs`](./WorkSheet/Code.gs) | Main entry points | `initializeWorkTracker`, `setupWorkTracker`, `rebuildLists`, `formatCurrentSheet` |
| [`WorkSheet/MonthSheet.gs`](./WorkSheet/MonthSheet.gs) | Monthly sheet engine | `createCurrentMonthSheet`, `createMonthRows` |
| [`WorkSheet/Tasks.gs`](./WorkSheet/Tasks.gs) | Task operations | `setupDropdowns`, `createDropdownRule`, `addTaskRow`, `clearTasks` |
| [`WorkSheet/Lists.gs`](./WorkSheet/Lists.gs) | Reference lists & protection | `createListsSheet`, `ensureListsSheet`, `createNamedRanges`, `protectListsSheet` |
| [`WorkSheet/Formatting.gs`](./WorkSheet/Formatting.gs) | Visual layout & styling | `formatWorkTracker`, `formatHeader`, `formatColumns`, `formatDimensions`, `formatBorders` |
| [`WorkSheet/ConditionalFormatting.gs`](./WorkSheet/ConditionalFormatting.gs) | Dynamic color rules | `setupConditionalFormatting`, `setupWeekendFormatting` |
| [`WorkSheet/Menu.gs`](./WorkSheet/Menu.gs) | Toolbar UI menus | `onOpen` (creates `Month Sheet`, `Setup`, `Tasks`, `DSR`) |
| [`WorkSheet/DSR.gs`](./WorkSheet/DSR.gs) | Daily Status Report engine | `generateDSRForToday`, `generateDSRForSelectedDate`, `showDSRDialog`, `saveDSRToSheet` |
| [`WorkSheet/Utils.gs`](./WorkSheet/Utils.gs) | Helper utilities | `trimSheet`, `getSpreadsheet`, `getTimezone`, `getToday` |
| [`WorkSheet/README.md`](./WorkSheet/README.md) | User & developer documentation | Comprehensive documentation of the WorkSheet system |

---

## 3. Retained Conversation Context & Prior Review Findings

### What Has Been Completed:
1. **Full Codebase Review**: Completed a deep-dive analysis of all Apps Script files.
2. **Documentation**: Authored the full [`WorkSheet/README.md`](./WorkSheet/README.md).
3. **DSR Generation & Modal Dialog**:
   - Implemented `DSR.gs` to extract tasks by date, categorize by project into Tasks Completed, Work in Progress, Blockers/Issues, and Plan for Next Working Day.
   - Built styled HTML modal dialog matching the user's UI specification (`Copy to Clipboard`, `Save to Sheet`, `Download .txt`, `Close`).
   - Updated `Menu.gs` to only have the two requested options: `Generate DSR for today` and `Generate DSR for selected date in the month`.
4. **Task Staging & Insertion Workflow**:
   - Updated `Tasks` menu to strictly have two options: `Fill task for today` and `Fill task for selected date in the current month`.
   - Creates a temporary staging sheet named `Date Month Year` (e.g. `08 September 2026`) for entering multiple tasks.
   - Supports inserting new row(s) after the date with the same date when multiple tasks are added.
   - Automatically deletes the staging sheet and refocuses the month sheet after saving.

### Outstanding Review Findings & Next Steps:
1. **Date Format Divergence**:
   - `MonthSheet.gs` formats dates as `'MMdd'` (e.g. `0908`).
   - `Tasks.gs:addTaskRow()` formats dates as `'ddMMM'` (e.g. `08SEP`).
   - *Target Fix*: Unify into a single `CONFIG.DATE_FORMAT` property in `Config.gs`.
3. **`formatCurrentSheet()` Weekend Formatting**:
   - `formatCurrentSheet()` in `Code.gs` omits `setupWeekendFormatting(sheet)`. Calling `formatCurrentSheet` loses Saturday/Sunday highlights.
4. **`clearTasks()` Target Safety**:
   - `clearTasks()` runs on `getActiveSheet()` without verifying if the active sheet is `Lists`. If run on the Lists sheet, it erases options and sets status to 'Pending'.

---

## 4. Development & Coding Guidelines

- **Configuration First**: Never hardcode list values or colors; always define them inside `CONFIG` in [`Config.gs`](./WorkSheet/Config.gs).
- **Batch Operations**: Use `setValues()` and `getValues()` in bulk; avoid calling Google Sheets APIs inside loops.
- **Timezone Handling**: Always derive the timezone via `getTimezone()` (`getSpreadsheetTimeZone()`) rather than assuming UTC or local server time.
- **Sheet Bounds**: Always run `trimSheet()` on newly inserted sheets to remove surplus empty rows and columns for sheet responsiveness.
- **Data Validation Integrity**: When modifying `CONFIG.LISTS`, always execute `createListsSheet()` or `rebuildLists()` to refresh Google Sheets Named Ranges.
