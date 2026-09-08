function setupDropdowns(sheet) {
  const rowCount =
    sheet.getLastRow() - 1;

  if (rowCount <= 0) {
    return;
  }

  const ss = sheet.getParent();

  const rules = {
    3: createDropdownRule(
      ss,
      'Projects'
    ),

    5: createDropdownRule(
      ss,
      'Categories'
    ),

    6: createDropdownRule(
      ss,
      'Priorities'
    ),

    7: createDropdownRule(
      ss,
      'Statuses'
    )
  };

  Object.entries(rules)
    .forEach(([column, rule]) => {
      sheet
        .getRange(
          2,
          Number(column),
          rowCount,
          1
        )
        .setDataValidation(rule);
    });
}


function createDropdownRule(
  ss,
  namedRange
) {
  return SpreadsheetApp
    .newDataValidation()
    .requireValueInRange(
      ss.getRangeByName(namedRange),
      true
    )
    .setAllowInvalid(false)
    .build();
}


function addTaskRow() {
  const sheet =
    getSpreadsheet()
      .getActiveSheet();

  const row =
    sheet.getLastRow() + 1;

  const timezone =
    getTimezone();

  const today =
    getToday();

  const dateLabel =
    Utilities.formatDate(
      today,
      timezone,
      'ddMMM'
    ).toUpperCase();

  const dayLabel =
    Utilities.formatDate(
      today,
      timezone,
      'EEE'
    );

  sheet
    .getRange(row, 1, 1, 8)
    .setValues([[
      dateLabel,
      dayLabel,
      '',
      '',
      '',
      '',
      'Pending',
      ''
    ]]);

  sheet.setRowHeight(
    row,
    42
  );

  setupDropdowns(sheet);

  sheet
    .getRange(
      row,
      1,
      1,
      CONFIG.HEADERS.length
    )
    .setWrapStrategy(
      SpreadsheetApp.WrapStrategy.WRAP
    );

  sheet.setActiveSelection(
    `D${row}`
  );
}


function clearTasks() {
  const ui =
    SpreadsheetApp.getUi();

  const response =
    ui.alert(
      'Clear Tasks',
      'Clear all task information?',
      ui.ButtonSet.YES_NO
    );

  if (
    response !== ui.Button.YES
  ) {
    return;
  }

  const sheet =
    getSpreadsheet()
      .getActiveSheet();

  const lastRow =
    sheet.getLastRow();

  if (lastRow <= 1) {
    return;
  }

  // Clear Project through Notes.
  sheet
    .getRange(
      2,
      3,
      lastRow - 1,
      6
    )
    .clearContent();

  // Restore default status.
  sheet
    .getRange(
      2,
      7,
      lastRow - 1,
      1
    )
    .setValue('Pending');
}