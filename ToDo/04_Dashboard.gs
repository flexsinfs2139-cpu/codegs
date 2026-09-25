/**
 * Dashboard sheet: KPI cards, quadrant/status panels and a "Focus now" list.
 * Columns J:K hold hidden helper counts that the visible cells reference.
 */

function setupDashboard_(sheet) {
  const A = todoCol_('task');
  const D = todoCol_('status');
  const E = todoCol_('due');
  const P = '  '; // visual left padding
  const CANVAS = '#f4f5f7';
  const lastRow = 25;
  const lastCol = 11; // A–I visible, J–K hidden helper data
  const statusNames = Object.keys(STATUSES);
  const BAR_COLORS = { 'To Do': '#9ca3af', 'In Progress': '#3b82f6', 'Done': '#22c55e', 'Blocked': '#ef4444' };
  const SOLID = SpreadsheetApp.BorderStyle.SOLID;
  const THICK = SpreadsheetApp.BorderStyle.SOLID_THICK;

  fitSheet_(sheet, lastRow, lastCol);
  sheet.showColumns(1, lastCol);
  sheet.setHiddenGridlines(true);

  // Canvas
  sheet.getRange(1, 1, lastRow, lastCol)
    .setFontFamily(CONFIG.font)
    .setFontSize(10)
    .setFontColor(THEME.ink)
    .setVerticalAlignment('middle')
    .setBackground(CANVAS);

  // Grid: margin | card | gutter | card | gutter | card | gutter | card | margin | data | data
  [28, 180, 16, 180, 16, 180, 16, 180, 28, 140, 60].forEach((w, i) => sheet.setColumnWidth(i + 1, w));
  sheet.setRowHeights(1, lastRow, 30);
  const heights = { 1: 16, 2: 40, 3: 20, 4: 20, 5: 28, 6: 46, 7: 22, 8: 10, 9: 24, 10: 38, 18: 24, 19: 38, 25: 24 };
  Object.keys(heights).forEach(r => sheet.setRowHeight(Number(r), heights[r]));
  sheet.setRowHeights(20, 5, 32);

  // ----------------------------------------------------------
  // Hidden data (J:K) — feeds the KPI cards and panels
  // ----------------------------------------------------------
  const quadrantData = [['Quadrant', 'Open']].concat(
    QUADRANTS.map(q => [`${q.id} ${q.title}`, `=${openCountFormula_(q)}`])
  );
  sheet.getRange(1, 10, quadrantData.length, 2).setValues(quadrantData);

  const statusData = [['Status', 'Count']].concat(
    statusNames.map(s => [s, `=COUNTIFS(${A},"<>",${D},"${s}")`])
  );
  sheet.getRange(7, 10, statusData.length, 2).setValues(statusData);
  const statusCell = s => `K${8 + statusNames.indexOf(s)}`;

  const overdue = 'K13';
  sheet.getRange('J13:K13').setValues([
    ['Overdue', `=COUNTIFS(${A},"<>",${D},"<>Done",${E},"<"&TODAY())`],
  ]);

  // ----------------------------------------------------------
  // Header
  // ----------------------------------------------------------
  mergeWithValue_(sheet.getRange('B2:F2'), 'Dashboard')
    .setFontSize(20)
    .setFontWeight('bold');
  sheet.getRange('H2')
    .setFormula('="Updated " & TEXT(NOW(),"mmm d, h:mm AM/PM")')
    .setFontSize(9)
    .setFontColor(THEME.muted)
    .setHorizontalAlignment('right');
  mergeWithValue_(sheet.getRange('B3:F3'), 'Live overview of your Eisenhower priorities')
    .setFontSize(9)
    .setFontColor(THEME.muted);

  // ----------------------------------------------------------
  // KPI cards (rows 5–8)
  // ----------------------------------------------------------
  const done = statusCell('Done');
  const inProgress = statusCell('In Progress');
  const toDo = statusCell('To Do');
  const blocked = statusCell('Blocked');

  const cards = [
    { col: 2, label: 'TOTAL TASKS', accent: THEME.accent, format: `"${P}"0`,
      value: `=COUNTA(${A})`,
      sub: `="${P}" & (B6-${done}) & " open  ·  " & ${done} & " done"` },
    { col: 4, label: 'IN PROGRESS', accent: '#3b82f6', format: `"${P}"0`,
      value: `=${inProgress}`,
      sub: `="${P}" & ${toDo} & " waiting to start"` },
    { col: 6, label: 'BLOCKED', accent: '#ef4444', format: `"${P}"0`,
      value: `=${blocked}`,
      sub: `=IF(AND(${blocked}=0,${overdue}=0),"${P}All clear",` +
           `"${P}" & IF(${blocked}>0,"Needs attention","No blockers") & IF(${overdue}>0,"  ·  " & ${overdue} & " overdue",""))` },
    { col: 8, label: 'COMPLETION', accent: '#22c55e', format: `"${P}"0%`,
      value: `=IFERROR(${done}/B6,0)`,
      sub: `=SPARKLINE(H6,{"charttype","bar";"max",1;"color1","#22c55e";"color2","${THEME.divider}"})` },
  ];

  cards.forEach(c => {
    sheet.getRange(5, c.col, 4, 1)
      .setBackground('#ffffff')
      .setBorder(true, true, true, true, null, null, THEME.border, SOLID);
    sheet.getRange(5, c.col)
      .setValue(P + c.label)
      .setFontSize(9)
      .setFontWeight('bold')
      .setFontColor(THEME.muted)
      .setBorder(true, null, null, null, null, null, c.accent, THICK);
    sheet.getRange(6, c.col)
      .setFormula(c.value)
      .setNumberFormat(c.format)
      .setFontSize(26)
      .setFontWeight('bold')
      .setHorizontalAlignment('left');
    sheet.getRange(7, c.col)
      .setFormula(c.sub)
      .setFontSize(9)
      .setFontColor(THEME.muted);
  });

  // ----------------------------------------------------------
  // Panel: Open by quadrant (chips + inline bars, no chart)
  // ----------------------------------------------------------
  panel_(sheet, 'B10:D17', P + 'Open by quadrant');

  const openTotal = `($B$6-${done})`;
  QUADRANTS.forEach((q, i) => {
    const r = 12 + i;
    const cell = `K${2 + i}`;
    sheet.getRange(r, 2)
      .setValue(`${q.id}  ·  ${q.title}`)
      .setBackground(q.bg)
      .setFontColor(q.text)
      .setFontSize(9)
      .setFontWeight('bold')
      .setHorizontalAlignment('center');
    sheet.getRange(r, 4)
      .setFormula(`=IF(${cell}=0,"0",REPT("█",MAX(1,ROUND(IFERROR(${cell}/${openTotal},0)*14))) & "  " & ${cell})`)
      .setFontColor(q.accent)
      .setFontWeight('bold');
  });

  sheet.getRange('B17:D17').merge().getCell(1, 1)
    .setFormula(`="${P}" & ${openTotal} & " open across 4 quadrants"`);
  sheet.getRange('B17').setFontSize(9).setFontColor(THEME.muted);

  // ----------------------------------------------------------
  // Panel: Status breakdown (chips + inline bars)
  // ----------------------------------------------------------
  panel_(sheet, 'F10:H17', P + 'Status breakdown');

  statusNames.forEach((s, i) => {
    const r = 12 + i;
    const cell = statusCell(s);
    sheet.getRange(r, 6)
      .setValue(s)
      .setBackground(STATUSES[s].bg)
      .setFontColor(STATUSES[s].fg)
      .setFontSize(9)
      .setFontWeight('bold')
      .setHorizontalAlignment('center');
    sheet.getRange(r, 8)
      .setFormula(`=IF(${cell}=0,"0",REPT("█",MAX(1,ROUND(IFERROR(${cell}/$B$6,0)*14))) & "  " & ${cell})`)
      .setFontColor(BAR_COLORS[s] || THEME.muted)
      .setFontWeight('bold');
  });

  sheet.getRange('F17:H17').merge().getCell(1, 1)
    .setFormula(`="${P}" & B6 & " tasks tracked"`);
  sheet.getRange('F17').setFontSize(9).setFontColor(THEME.muted);

  // ----------------------------------------------------------
  // Panel: Focus now (open Q1 tasks, soonest due first)
  // ----------------------------------------------------------
  const q1 = QUADRANTS[0];
  const emptyMsg = `${P}Nothing urgent. Nice work!`;
  const focusList = `SORT(FILTER({ARRAYFORMULA("${P}•  "&${A}),${E},${D}},${openTaskConditions_(q1)}),2,TRUE)`;
  const focusColumn = (i, fallback) =>
    `=IFERROR(ARRAY_CONSTRAIN(INDEX(${focusList},0,${i}),5,1),"${fallback}")`;

  const focusHeader = panel_(sheet, 'B19:H24', '');
  focusHeader.getCell(1, 1).setFormula(
    `="${P}Focus now  ·  ${q1.id} ${q1.title}   (" & K2 & " open)"`
  );
  sheet.getRange('B19:B24')
    .setBorder(null, true, null, null, null, null, q1.accent, THICK);

  sheet.getRange('B20').setFormula(focusColumn(1, emptyMsg));
  sheet.getRange('B20:B24').setWrapStrategy(SpreadsheetApp.WrapStrategy.OVERFLOW);

  sheet.getRange('F20').setFormula(focusColumn(2, ''));
  sheet.getRange('F20:F24')
    .setNumberFormat('"Due "mmm d')
    .setHorizontalAlignment('center')
    .setFontSize(9)
    .setFontColor(THEME.muted);

  sheet.getRange('H20').setFormula(focusColumn(3, ''));
  sheet.getRange('H20:H24')
    .setHorizontalAlignment('center')
    .setFontSize(9)
    .setFontWeight('bold');

  sheet.getRange('B20:H24')
    .setBorder(null, null, null, null, null, true, THEME.divider, SOLID);

  // ----------------------------------------------------------
  // Conditional formatting
  // ----------------------------------------------------------
  const attentionAlert = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied(`=OR($F$6>0,${overdue}>0)`)
    .setFontColor(THEME.danger)
    .setBold(true)
    .setRanges([sheet.getRange('F7')])
    .build();

  const emptyFocus = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo(emptyMsg)
    .setFontColor(THEME.faint)
    .setItalic(true)
    .setRanges([sheet.getRange('B20')])
    .build();

  sheet.setConditionalFormatRules(
    [attentionAlert, emptyFocus, overdueRule_(sheet.getRange('F20:F24'))]
      .concat(statusRules_([sheet.getRange('H20:H24')]))
  );

  sheet.hideColumns(10, 2);

  sheet.protect()
    .setDescription('Auto-generated. Edit tasks on the TODO sheet.')
    .setWarningOnly(true);
}

/** White card with a bold header row; returns the merged header range. */
function panel_(sheet, a1, title) {
  const range = sheet.getRange(a1);
  range
    .setBackground('#ffffff')
    .setBorder(true, true, true, true, null, null, THEME.border, SpreadsheetApp.BorderStyle.SOLID);

  const header = sheet.getRange(range.getRow(), range.getColumn(), 1, range.getNumColumns()).merge();
  if (title) header.getCell(1, 1).setValue(title);
  header
    .setFontSize(11)
    .setFontWeight('bold')
    .setBorder(null, null, true, null, null, null, THEME.divider, SpreadsheetApp.BorderStyle.SOLID);
  return header;
}
