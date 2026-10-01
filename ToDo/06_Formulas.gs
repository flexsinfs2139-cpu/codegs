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
