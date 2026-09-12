function createListsSheet() {
  const ss = getSpreadsheet();

  const sheet = resetListsSheet(ss);

  writeLists(sheet);
  formatListsSheet(sheet);
  createNamedRanges(ss, sheet);
  trimListsSheet(sheet);
  protectListsSheet(sheet);

  ss.setActiveSheet(sheet);

  SpreadsheetApp.getUi().alert(
    'Lists sheet created successfully.'
  );
}


function ensureListsSheet(ss) {
  let sheet = ss.getSheetByName(
    CONFIG.LISTS_SHEET_NAME
  );

  if (sheet) {
    return sheet;
  }

  createListsSheet();

  return ss.getSheetByName(
    CONFIG.LISTS_SHEET_NAME
  );
}


function resetListsSheet(ss) {
  let sheet = ss.getSheetByName(
    CONFIG.LISTS_SHEET_NAME
  );

  if (sheet) {
    removeListsProtection(sheet);
    removeNamedRanges(ss);
    sheet.clear();
    return sheet;
  }

  return ss.insertSheet(
    CONFIG.LISTS_SHEET_NAME
  );
}


function writeLists(sheet) {
  const headers = Object.keys(CONFIG.LISTS);

  sheet
    .getRange(
      1,
      1,
      1,
      headers.length
    )
    .setValues([headers]);

  headers.forEach((name, index) => {
    const values = CONFIG.LISTS[name];

    sheet
      .getRange(
        2,
        index + 1,
        values.length,
        1
      )
      .setValues(
        values.map(value => [value])
      );
  });
}


function formatListsSheet(sheet) {
  const lastRow =
    1 + Math.max(
      ...Object.values(CONFIG.LISTS)
        .map(list => list.length)
    );

  const lastCol =
    Object.keys(CONFIG.LISTS).length;

  sheet
    .getRange(
      1,
      1,
      1,
      lastCol
    )
    .setBackground(CONFIG.COLORS.HEADER)
    .setFontFamily('Arial')
    .setFontWeight('bold')
    .setHorizontalAlignment('center');

  sheet
    .getRange(
      1,
      1,
      lastRow,
      lastCol
    )
    .setFontFamily('Arial')
    .setVerticalAlignment('middle');

  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, lastCol);
}


function createNamedRanges(ss, sheet) {
  const lists = CONFIG.LISTS;

  Object.keys(lists).forEach((name, index) => {
    const values = lists[name];

    ss.setNamedRange(
      name,
      sheet.getRange(
        2,
        index + 1,
        values.length,
        1
      )
    );
  });
}


function removeNamedRanges(ss) {
  const names = Object.keys(CONFIG.LISTS);

  ss.getNamedRanges()
    .filter(namedRange =>
      names.includes(namedRange.getName())
    )
    .forEach(namedRange =>
      namedRange.remove()
    );
}


function trimListsSheet(sheet) {
  const requiredRows =
    1 + Math.max(
      ...Object.values(CONFIG.LISTS)
        .map(list => list.length)
    );

  const requiredColumns =
    Object.keys(CONFIG.LISTS).length;

  trimSheet(
    sheet,
    requiredRows,
    requiredColumns
  );
}


function protectListsSheet(sheet) {
  removeListsProtection(sheet);

  const protection = sheet.protect();

  protection.setDescription(
    'Protected Lists Sheet'
  );

  protection.removeEditors(
    protection.getEditors()
  );

  if (protection.canDomainEdit()) {
    protection.setDomainEdit(false);
  }
}


function removeListsProtection(sheet) {
  sheet
    .getProtections(
      SpreadsheetApp.ProtectionType.SHEET
    )
    .forEach(protection => {
      if (
        protection.getDescription() ===
        'Protected Lists Sheet'
      ) {
        protection.remove();
      }
    });
}