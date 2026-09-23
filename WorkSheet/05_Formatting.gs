function formatWorkTracker(sheet) {
  if (!sheet) {
    sheet =
      getSpreadsheet()
        .getActiveSheet();
  }

  const lastRow =
    Math.max(sheet.getLastRow(), 2);

  const range =
    sheet.getRange(
      1,
      1,
      lastRow,
      CONFIG.HEADERS.length
    );

  range
    .setFontFamily(CONFIG.FONTS.TEXT)
    .setFontSize(10)
    .setVerticalAlignment('middle')
    .setWrapStrategy(
      SpreadsheetApp.WrapStrategy.WRAP
    );

  formatHeader(sheet);
  formatColumns(sheet);
  formatDimensions(sheet);
  formatBorders(range);
  setSheetGridlinesHidden(sheet, true);
}


function formatHeader(sheet) {
  sheet
    .getRange(
      1,
      1,
      1,
      CONFIG.HEADERS.length
    )
    .setBackground(CONFIG.COLORS.HEADER)
    .setFontFamily(CONFIG.FONTS.TEXT)
    .setFontWeight('bold')
    .setHorizontalAlignment('center');
}


function formatColumns(sheet) {
  sheet
    .getRange('A:A')
    .setFontFamily(CONFIG.FONTS.DIGITS)
    .setFontSize(10)
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setNumberFormat('0000');

  sheet
    .getRange('B:B')
    .setFontFamily(CONFIG.FONTS.TEXT)
    .setHorizontalAlignment('center');

  sheet.getRange('C:C')
    .setHorizontalAlignment('center');

  sheet.getRange('D:D')
    .setHorizontalAlignment('left');

  sheet.getRange('E:G')
    .setHorizontalAlignment('center');
}


function formatDimensions(sheet) {
  const widths = [
    75,   // Date
    60,   // Day
    130,  // Project
    340,  // Task
    130,  // Category
    100,  // Priority
    125   // Status
  ];

  widths.forEach((width, index) => {
    sheet.setColumnWidth(
      index + 1,
      width
    );
  });

  sheet.setRowHeight(1, 28);

  const lastRow = sheet.getLastRow();

  if (lastRow > 1) {
    sheet.setRowHeights(
      2,
      lastRow - 1,
      42
    );
  }
}


function formatBorders(range) {
  range.setBorder(
    true,
    true,
    true,
    true,
    true,
    true,
    CONFIG.COLORS.BORDER,
    SpreadsheetApp.BorderStyle.SOLID
  );
}