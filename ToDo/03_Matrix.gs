/**
 * Eisenhower Matrix sheet: four live quadrant cards fed from TODO.
 */

function setupMatrixSheet_(sheet) {
  const S = CONFIG.slots;
  const top = 5;                 // first quadrant row (Q1/Q2 header)
  const bottom = top + S + 3;    // Q3/Q4 header row
  const lastRow = bottom + S + 2;
  const lastCol = 8;

  fitSheet_(sheet, lastRow, lastCol);
  sheet.setHiddenGridlines(true);

  sheet.getRange(1, 1, lastRow, lastCol)
    .setFontFamily(CONFIG.font)
    .setFontSize(10)
    .setFontColor(THEME.ink)
    .setVerticalAlignment('middle');

  // Layout: A = axis label, B:D = left quadrant (task, due, status), E = gutter, F:H = right quadrant
  [40, 250, 80, 110, 20, 250, 80, 110].forEach((w, i) => sheet.setColumnWidth(i + 1, w));
  sheet.setRowHeights(1, lastRow, 30);
  sheet.setRowHeight(1, 48);
  sheet.setRowHeight(2, 22);
  sheet.setRowHeight(3, 12);
  sheet.setRowHeight(4, 26);
  sheet.setRowHeight(top + S + 2, 16);
  sheet.setRowHeight(lastRow, 16);

  // Title
  mergeWithValue_(sheet.getRange('B1:H1'), 'Eisenhower Matrix')
    .setFontSize(18)
    .setFontWeight('bold');
  mergeWithValue_(sheet.getRange('B2:H2'), 'Prioritize by importance and urgency · open tasks only, soonest due first, updates live')
    .setFontColor(THEME.muted)
    .setFontSize(9);

  // Axis labels
  [['B4:D4', 'URGENT'], ['F4:H4', 'NOT URGENT']].forEach(([a1, text]) => {
    mergeWithValue_(sheet.getRange(a1), text)
      .setFontSize(9)
      .setFontWeight('bold')
      .setFontColor(THEME.muted)
      .setHorizontalAlignment('center');
  });

  [[top, 'IMPORTANT'], [bottom, 'NOT IMPORTANT']].forEach(([row, text]) => {
    mergeWithValue_(sheet.getRange(row, 1, S + 2, 1), text)
      .setTextRotation(90)
      .setFontSize(9)
      .setFontWeight('bold')
      .setFontColor(THEME.muted)
      .setHorizontalAlignment('center');
  });

  // Quadrants
  const positions = [[top, 2], [top, 6], [bottom, 2], [bottom, 6]];
  const listRanges = [];
  const statusRanges = [];
  const overdueRules = [];

  QUADRANTS.forEach((q, i) => {
    const [row, col] = positions[i];
    buildQuadrant_(sheet, q, row, col);
    listRanges.push(sheet.getRange(row + 2, col, S, 1));
    statusRanges.push(sheet.getRange(row + 2, col + 2, S, 1));
    overdueRules.push(overdueRule_(sheet.getRange(row + 2, col + 1, S, 1)));
  });

  const placeholder = builder => builder
    .setFontColor(THEME.faint)
    .setItalic(true)
    .setRanges(listRanges)
    .build();

  sheet.setConditionalFormatRules([
    placeholder(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo(EMPTY_LIST_TEXT)),
    placeholder(SpreadsheetApp.newConditionalFormatRule().whenTextStartsWith(MORE_PREFIX)),
  ].concat(overdueRules, statusRules_(statusRanges)));

  sheet.protect()
    .setDescription('Auto-generated. Edit tasks on the TODO sheet.')
    .setWarningOnly(true);
}

function buildQuadrant_(sheet, q, row, col) {
  const S = CONFIG.slots;
  const W = 3; // task, due, status

  // Header with live open-task count
  const header = sheet.getRange(row, col, 1, W).merge();
  header.getCell(1, 1).setFormula(
    `="${q.id}  ·  ${q.title.toUpperCase()}     " & ${openCountFormula_(q)} & " open"`
  );
  header
    .setBackground(q.bg)
    .setFontColor(q.text)
    .setFontSize(12)
    .setFontWeight('bold')
    .setHorizontalAlignment('left')
    .setBorder(true, null, null, null, null, null, q.accent, SpreadsheetApp.BorderStyle.SOLID_THICK);
  sheet.setRowHeight(row, 38);

  // Caption
  mergeWithValue_(sheet.getRange(row + 1, col, 1, W), q.caption)
    .setBackground(q.bg)
    .setFontColor(q.text)
    .setFontSize(9);
  sheet.setRowHeight(row + 1, 20);

  // Task list (Task + Due + Status), capped to the slot count so it never spills
  sheet.getRange(row + 2, col, S, W)
    .setBackground('#ffffff')
    .setBorder(null, null, null, null, null, true, THEME.divider, SpreadsheetApp.BorderStyle.SOLID);

  sheet.getRange(row + 2, col, S, 1).setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
  sheet.getRange(row + 2, col + 1, S, 1)
    .setNumberFormat('mmm d')
    .setHorizontalAlignment('center')
    .setFontSize(9)
    .setFontColor(THEME.muted);
  sheet.getRange(row + 2, col + 2, S, 1)
    .setHorizontalAlignment('center')
    .setFontSize(9)
    .setFontWeight('bold');

  sheet.getRange(row + 2, col).setFormula(matrixListFormula_(q));

  // Card outline
  sheet.getRange(row, col, S + 2, W)
    .setBorder(true, true, true, true, null, null, THEME.border, SpreadsheetApp.BorderStyle.SOLID);
}
