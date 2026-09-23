function createCurrentMonthSheet(suppressAlert) {
  const ss = getSpreadsheet();
  const timezone = getTimezone();
  const today = getToday();

  const year = today.getFullYear();
  const month = today.getMonth();

  const sheetName = Utilities.formatDate(
    today,
    timezone,
    'MMMyy'
  ).toUpperCase();

  let sheet = ss.getSheetByName(sheetName);
  const isExisting = Boolean(sheet);

  ensureListsSheet(ss);

  if (isExisting) {
    // Idempotent upgrade: preserve existing data, refresh headers & rules
    sheet
      .getRange(CONFIG.HEADER_ROW, 1, 1, CONFIG.HEADERS.length)
      .setValues([CONFIG.HEADERS]);

    formatWorkTracker(sheet);
    setupDropdowns(sheet);
    setupConditionalFormatting(sheet);
    setupWeekendFormatting(sheet);
    sheet.setFrozenRows(1);
    setSheetGridlinesHidden(sheet, true);

    ss.setActiveSheet(sheet);

    if (!suppressAlert) {
      SpreadsheetApp.getUi().alert(
        `${sheetName} verified and refreshed successfully.`
      );
    }

    return sheet;
  }

  sheet = ss.insertSheet(sheetName);

  createMonthRows(
    sheet,
    year,
    month,
    timezone
  );

  trimSheet(
    sheet,
    sheet.getLastRow(),
    CONFIG.HEADERS.length
  );

  formatWorkTracker(sheet);
  setupDropdowns(sheet);
  setupConditionalFormatting(sheet);
  setupWeekendFormatting(sheet);

  sheet.setFrozenRows(1);
  setSheetGridlinesHidden(sheet, true);

  ss.setActiveSheet(sheet);

  if (!suppressAlert) {
    SpreadsheetApp.getUi().alert(
      `${sheetName} created successfully.`
    );
  }

  return sheet;
}


function createMonthRows(
  sheet,
  year,
  month,
  timezone
) {
  sheet
    .getRange(
      CONFIG.HEADER_ROW,
      1,
      1,
      CONFIG.HEADERS.length
    )
    .setValues([
      CONFIG.HEADERS
    ]);

  const daysInMonth =
    new Date(
      year,
      month + 1,
      0
    ).getDate();

  const rows = [];

  for (
    let day = 1;
    day <= daysInMonth;
    day++
  ) {
    const date =
      new Date(year, month, day);

    const dateLabel =
      Utilities.formatDate(
        date,
        timezone,
        CONFIG.DATE_FORMAT || 'MMdd'
      );

    const dayLabel =
      Utilities.formatDate(
        date,
        timezone,
        'EEE'
      );

    for (
      let task = 0;
      task < CONFIG.DEFAULT_TASKS_PER_DAY;
      task++
    ) {
      rows.push([
        dateLabel,
        dayLabel,
        '',
        '',
        '',
        '',
        'Pending'
      ]);
    }
  }

  sheet
    .getRange(
      2,
      1,
      rows.length,
      CONFIG.HEADERS.length
    )
    .setValues(rows);
}