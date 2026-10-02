/**
 * Formula and conditional-format builders for the TODO and Archive sheets.
 */

/** Open-ended task-table column on the same sheet, e.g. $A$2:$A */
function taskCol_(key) {
  const letter = columnLetter_(COL[key]);
  return `$${letter}$2:$${letter}`;
}

function bool_(value) {
  return value ? 'TRUE' : 'FALSE';
}

/** Open tasks of one quadrant, in task-table order (blank when there are none). */
function quadrantListFormula_(q) {
  const conditions = [
    `${taskCol_('task')}<>""`,
    `${taskCol_('important')}=${bool_(q.important)}`,
    `${taskCol_('urgent')}=${bool_(q.urgent)}`,
    `${taskCol_('status')}<>"Done"`,
  ].join(',');
  return `=IFERROR(FILTER(${taskCol_('task')},${conditions}),"")`;
}

function statusRules_(ranges) {
  return Object.keys(STATUSES).map(status =>
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo(status)
      .setBackground(STATUSES[status].bg)
      .setFontColor(STATUSES[status].fg)
      .setRanges(ranges)
      .build()
  );
}

/** Checkbox active highlight rules: soft warm amber for Important, soft warm rose for Urgent. */
function priorityRules_(sheet, body) {
  const impCol = columnLetter_(COL.important);
  const urgCol = columnLetter_(COL.urgent);
  const impRange = sheet.getRange(2, COL.important, body, 1);
  const urgRange = sheet.getRange(2, COL.urgent, body, 1);

  const impRule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied(`=$${impCol}2=TRUE`)
    .setBackground(THEME.importantBg)
    .setFontColor(THEME.importantFg)
    .setRanges([impRange])
    .build();

  const urgRule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied(`=$${urgCol}2=TRUE`)
    .setBackground(THEME.urgentBg)
    .setFontColor(THEME.urgentFg)
    .setRanges([urgRange])
    .build();

  return [impRule, urgRule];
}

/** Highlights open tasks in column A with their corresponding quadrant color theme. */
function taskQuadrantHighlightRules_(sheet, body) {
  const taskCol = columnLetter_(COL.task);
  const statusCol = columnLetter_(COL.status);
  const impCol = columnLetter_(COL.important);
  const urgCol = columnLetter_(COL.urgent);
  const taskRange = sheet.getRange(2, COL.task, body, 1);

  return QUADRANTS.map((q, i) => {
    const formula = `=AND($${taskCol}2<>"",$${statusCol}2<>"Done",$${impCol}2=${bool_(q.important)},$${urgCol}2=${bool_(q.urgent)})`;
    const builder = SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(formula)
      .setBackground(q.bg)
      .setFontColor(q.text)
      .setRanges([taskRange]);
    if (i === 0) builder.setBold(true);
    return builder.build();
  });
}

/** Enhances typography and styling for live quadrant columns (F:I) when a task is filtered into them. */
function quadrantColumnRules_(sheet, body) {
  return QUADRANTS.map((q, i) => {
    const col = MATRIX_COL + i;
    const colLetter = columnLetter_(col);
    const range = sheet.getRange(2, col, body, 1);
    const builder = SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(`=$${colLetter}2<>""`)
      .setFontColor(q.text)
      .setBackground(q.bg)
      .setRanges([range]);
    if (i === 0) builder.setBold(true);
    return builder.build();
  });
}
