// ============================================================
// UTILS.GS — SHARED UTILITIES & DATE HELPERS
// ============================================================

/**
 * Trims excess rows and columns from a sheet to improve performance.
 */
function trimSheet(sheet, lastRow, lastCol) {
  const maxRows = sheet.getMaxRows();
  const maxCols = sheet.getMaxColumns();

  if (maxCols > lastCol) {
    sheet.deleteColumns(
      lastCol + 1,
      maxCols - lastCol
    );
  }

  if (maxRows > lastRow) {
    sheet.deleteRows(
      lastRow + 1,
      maxRows - lastRow
    );
  }
}

/**
 * Returns active spreadsheet instance.
 */
function getSpreadsheet() {
  return SpreadsheetApp.getActiveSpreadsheet();
}

/**
 * Returns spreadsheet timezone.
 */
function getTimezone() {
  return getSpreadsheet().getSpreadsheetTimeZone();
}

/**
 * Returns current date instance.
 */
function getToday() {
  return new Date();
}

/**
 * Returns month sheet name format, e.g. "SEP26".
 */
function getMonthSheetName(date, timezone) {
  return Utilities.formatDate(
    date || getToday(),
    timezone || getTimezone(),
    'MMMyy'
  ).toUpperCase();
}

/**
 * Returns staging sheet name format, e.g. "08 September 2026".
 */
function getStagingSheetName(date, timezone) {
  return Utilities.formatDate(
    date || getToday(),
    timezone || getTimezone(),
    'dd MMMM yyyy'
  );
}

/**
 * Checks if a cell value matches the target Date.
 * Handles Date objects, 'MMdd', 'ddMMM', 'DD/MM/YYYY', numbers, etc.
 */
function isMatchingDate(cellValue, targetDate, timezone) {
  if (cellValue === null || cellValue === undefined || cellValue === '') {
    return false;
  }

  const targetDay = targetDate.getDate();
  const targetMonth = targetDate.getMonth();
  const targetYear = targetDate.getFullYear();

  if (cellValue instanceof Date) {
    return (
      cellValue.getDate() === targetDay &&
      cellValue.getMonth() === targetMonth &&
      cellValue.getFullYear() === targetYear
    );
  }

  if (typeof cellValue === 'number') {
    // Check if day of month matches directly (1-31)
    if (cellValue === targetDay) return true;

    // Check if numeric MMdd (e.g. 908 for Sept 8)
    const month = Math.floor(cellValue / 100);
    const day = cellValue % 100;
    if (month === targetMonth + 1 && day === targetDay) return true;
  }

  const str = String(cellValue).trim();

  // Match MMdd format (e.g. "0908")
  const mmdd = Utilities.formatDate(targetDate, timezone, 'MMdd');
  if (str === mmdd) return true;

  // Match ddMMM format (e.g. "08SEP")
  const ddmmm = Utilities.formatDate(targetDate, timezone, 'ddMMM').toUpperCase();
  if (str.toUpperCase() === ddmmm) return true;

  // Match dd/MM/yyyy or dd-MM-yyyy
  const ddmmyyyy1 = Utilities.formatDate(targetDate, timezone, 'dd/MM/yyyy');
  const ddmmyyyy2 = Utilities.formatDate(targetDate, timezone, 'dd-MM-yyyy');
  if (str === ddmmyyyy1 || str === ddmmyyyy2) return true;

  // Match single day number string e.g. "8" or "08"
  if (/^\d{1,2}$/.test(str)) {
    if (parseInt(str, 10) === targetDay) return true;
  }

  return false;
}

/**
 * Parses date from cell value and sheet name (e.g. "SEP26").
 */
function parseDateFromCell(cellValue, sheetName, timezone) {
  if (!cellValue) return null;

  if (cellValue instanceof Date) {
    return cellValue;
  }

  const str = String(cellValue).trim();
  const now = getToday();
  let year = now.getFullYear();
  let month = now.getMonth();

  // Try extracting month and year from sheet name (e.g. "SEP26")
  const sheetMatch = sheetName ? sheetName.match(/^([A-Za-z]{3})(\d{2})$/) : null;
  if (sheetMatch) {
    const monthNames = [
      'JAN', 'FEB', 'MAR', 'APR', 'MAY', 'JUN',
      'JUL', 'AUG', 'SEP', 'OCT', 'NOV', 'DEC'
    ];
    const mIndex = monthNames.indexOf(sheetMatch[1].toUpperCase());
    if (mIndex !== -1) {
      month = mIndex;
    }
    year = 2000 + parseInt(sheetMatch[2], 10);
  }

  // Check MMdd format (4 digits, e.g. "0908")
  if (/^\d{4}$/.test(str)) {
    const m = parseInt(str.substring(0, 2), 10) - 1;
    const d = parseInt(str.substring(2, 4), 10);
    return new Date(year, m, d);
  }

  // Check ddMMM format (e.g. "08SEP")
  const ddmmmMatch = str.match(/^(\d{1,2})([A-Za-z]{3})$/);
  if (ddmmmMatch) {
    const d = parseInt(ddmmmMatch[1], 10);
    const monthNames = [
      'JAN', 'FEB', 'MAR', 'APR', 'MAY', 'JUN',
      'JUL', 'AUG', 'SEP', 'OCT', 'NOV', 'DEC'
    ];
    const m = monthNames.indexOf(ddmmmMatch[2].toUpperCase());
    if (m !== -1) {
      return new Date(year, m, d);
    }
  }

  // Check day number (1-31)
  if (/^\d{1,2}$/.test(str)) {
    return new Date(year, month, parseInt(str, 10));
  }

  // Check DD/MM/YYYY or DD-MM-YYYY
  const dateParts = str.split(/[\/\-\.]/);
  if (dateParts.length === 3) {
    const d = parseInt(dateParts[0], 10);
    const m = parseInt(dateParts[1], 10) - 1;
    let y = parseInt(dateParts[2], 10);
    if (y < 100) y += 2000;
    return new Date(y, m, d);
  }

  return null;
}

/**
 * Resolves user text input into a Date object for the current month.
 */
function resolveDateInput(input, sheetName, timezone) {
  if (!input) return null;

  const now = getToday();
  let year = now.getFullYear();
  let month = now.getMonth();

  // Try extracting month & year from sheet name (e.g. "SEP26")
  const sheetMatch = sheetName ? sheetName.match(/^([A-Za-z]{3})(\d{2})$/) : null;
  if (sheetMatch) {
    const monthNames = [
      'JAN', 'FEB', 'MAR', 'APR', 'MAY', 'JUN',
      'JUL', 'AUG', 'SEP', 'OCT', 'NOV', 'DEC'
    ];
    const mIndex = monthNames.indexOf(sheetMatch[1].toUpperCase());
    if (mIndex !== -1) {
      month = mIndex;
    }
    year = 2000 + parseInt(sheetMatch[2], 10);
  }

  const str = String(input).trim();

  // Case 1: Simple day number (1–31)
  if (/^\d{1,2}$/.test(str)) {
    const day = parseInt(str, 10);
    if (day >= 1 && day <= 31) {
      return new Date(year, month, day);
    }
  }

  // Case 2: DD/MM/YYYY or DD-MM-YYYY
  const parts = str.split(/[\/\-\.]/);
  if (parts.length === 3) {
    const day = parseInt(parts[0], 10);
    const m = parseInt(parts[1], 10) - 1;
    let y = parseInt(parts[2], 10);
    if (y < 100) y += 2000;
    return new Date(y, m, day);
  }

  // Case 3: DD/MM (assume year from current context)
  if (parts.length === 2) {
    const day = parseInt(parts[0], 10);
    const m = parseInt(parts[1], 10) - 1;
    return new Date(year, m, day);
  }

  return null;
}