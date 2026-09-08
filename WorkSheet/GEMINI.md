# Antigravity Context & Project Memory — WorkSheet

This file serves as persistent memory and context for Antigravity (the equivalent of `CLAUDE.md` / `.claude` in Claude). It is automatically loaded into the agent's context window on every turn to retain conversation history, architectural decisions, and project conventions.

---

## 1. Project Summary

- **Project**: WorkSheet (Google Apps Script)
- **Host Platform**: Google Sheets
- **Target Repository**: `/home/rvkt/Documents/codegs/WorkSheet`
- **Purpose**: Automated work tracker managing daily tasks, project assignments, priorities, and statuses with dynamic monthly sheet generation, protected reference lists, and custom UI formatting.

---

## 2. Key Architecture & File Mapping

| File | Primary Role | Key Functions / Objects |
| :--- | :--- | :--- |
| [`Config.gs`](./Config.gs) | Central configuration | `CONFIG` (Headers, Lists, Colors, Row Heights) |
| [`Code.gs`](./Code.gs) | Main entry points | `initializeWorkTracker`, `setupWorkTracker`, `rebuildLists` |
| [`MonthSheet.gs`](./MonthSheet.gs) | Monthly sheet engine | `createCurrentMonthSheet`, `createMonthRows` |
| [`Tasks.gs`](./Tasks.gs) | Task operations | `setupDropdowns`, `createDropdownRule`, `addTaskRow`, `clearTasks` |
| [`Lists.gs`](./Lists.gs) | Reference lists & protection | `createListsSheet`, `ensureListsSheet`, `createNamedRanges`, `protectListsSheet` |
| [`Formatting.gs`](./Formatting.gs) | Visual layout & styling | `formatWorkTracker`, `formatHeader`, `formatColumns`, `formatDimensions`, `formatBorders` |
| [`ConditionalFormatting.gs`](./ConditionalFormatting.gs) | Dynamic color rules | `setupConditionalFormatting`, `setupWeekendFormatting` |
| [`Menu.gs`](./Menu.gs) | Toolbar UI menus | `onOpen` (creates `Month Sheet`, `Setup`, `Tasks`, `DSR`) |
| [`DSR.gs`](./DSR.gs) | Daily Status Report engine | `generateDSRForToday`, `generateDSRForSelectedDate`, `showDSRDialog`, `saveDSRToSheet` |
| [`Utils.gs`](./Utils.gs) | Helper utilities | `trimSheet`, `getSpreadsheet`, `getTimezone`, `getToday` |
| [`README.md`](./README.md) | User & developer documentation | Comprehensive documentation of the WorkSheet system |

---

## 3. Retained Conversation Context & Prior Review Findings

### What Has Been Completed:
1. **Full Codebase Review**: Completed a deep-dive analysis of all Apps Script files.
2. **Documentation**: Authored the full [`README.md`](./README.md) for the `WorkSheet` directory.
3. **DSR Generation & Modal Dialog**:
   - Implemented `DSR.gs` to extract tasks by date, categorize by project into Tasks Completed, Work in Progress, Blockers/Issues, and Plan for Next Working Day.
   - Built styled HTML modal dialog matching the user's UI specification (`Copy to Clipboard`, `Save to Sheet`, `Download .txt`, `Close`).
   - Updated `Menu.gs` to only have the two requested options: `Generate DSR for today` and `Generate DSR for selected date in the month`.
4. **Task Staging & Insertion Workflow**:
   - Updated `Tasks` menu to strictly have two options: `Fill task for today` and `Fill task for selected date in the current month`.
   - Prompts strictly for a day number (`1–30` or `1–31`) based on the current month's days.
   - Creates a temporary staging sheet named `Date Month Year` (e.g. `08 September 2026`) containing only task fields (`Project`, `Task`, `Category`, `Priority`, `Status`, `Notes`), omitting redundant Date and Day columns.
   - Automatically populates Date and Day upon saving to the month sheet.
   - Supports inserting new row(s) after the date with the same date when multiple tasks are added.
   - Automatically deletes the staging sheet and refocuses the month sheet after saving.

### Outstanding Review Findings & Next Steps:
1. **Date Format Divergence**:
   - `MonthSheet.gs` formats dates as `'MMdd'` (e.g. `0908`).
   - `Tasks.gs:addTaskRow()` formats dates as `'ddMMM'` (e.g. `08SEP`).
   - *Target Fix*: Unify into a single `CONFIG.DATE_FORMAT` property in `Config.gs`.
3. **`clearTasks()` Target Safety**:
   - `clearTasks()` runs on `getActiveSheet()` without verifying if the active sheet is `Lists`. If run on the Lists sheet, it erases options and sets status to 'Pending'.

---

## 4. Development & Coding Guidelines

- **Configuration First**: Never hardcode list values or colors; always define them inside `CONFIG` in [`Config.gs`](./Config.gs).
- **Batch Operations**: Use `setValues()` and `getValues()` in bulk; avoid calling Google Sheets APIs inside loops.
- **Timezone Handling**: Always derive the timezone via `getTimezone()` (`getSpreadsheetTimeZone()`) rather than assuming UTC or local server time.
- **Sheet Bounds**: Always run `trimSheet()` on newly inserted sheets to remove surplus empty rows and columns for sheet responsiveness.
- **Data Validation Integrity**: When modifying `CONFIG.LISTS`, always execute `createListsSheet()` or `rebuildLists()` to refresh Google Sheets Named Ranges.
