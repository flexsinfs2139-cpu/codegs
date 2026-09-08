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
    .setFontFamily('Arial')
    .setFontSize(10)
    .setVerticalAlignment('middle')
    .setWrapStrategy(
      SpreadsheetApp.WrapStrategy.WRAP
    );

  formatHeader(sheet);
  formatColumns(sheet);
  formatDimensions(sheet);
  formatBorders(range);
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
    .setFontFamily('Arial')
    .setFontWeight('bold')
    .setHorizontalAlignment('center');
}


function formatColumns(sheet) {
  sheet
    .getRange('A:A')
    .setFontFamily('Roboto Mono')
    .setFontSize(10)
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setNumberFormat('0000');

  sheet
    .getRange('B:B')
    .setFontFamily('Roboto Mono')
    .setHorizontalAlignment('center');

  sheet.getRange('C:C')
    .setHorizontalAlignment('center');

  sheet.getRange('D:D')
    .setHorizontalAlignment('left');

  sheet.getRange('E:G')
    .setHorizontalAlignment('center');

  sheet.getRange('H:H')
    .setHorizontalAlignment('left');
}


function formatDimensions(sheet) {
  const widths = [
    75,
    60,
    130,
    315,
    125,
    90,
    125,
    300
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