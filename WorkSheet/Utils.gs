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


function getSpreadsheet() {
  return SpreadsheetApp.getActiveSpreadsheet();
}


function getTimezone() {
  return getSpreadsheet().getSpreadsheetTimeZone();
}


function getToday() {
  return new Date();
}