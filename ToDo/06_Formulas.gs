/**
 * Formula and conditional-format builders shared by the Matrix and Dashboard.
 */

/** Full-column reference on the TODO sheet, e.g. 'TODO'!E2:E */
function todoCol_(key) {
  const letter = columnLetter_(COL[key]);
  return `${TODO_REF}!${letter}2:${letter}`;
}

function bool_(value) {
  return value ? 'TRUE' : 'FALSE';
}

function openCountFormula_(q) {
  return `COUNTIFS(${todoCol_('task')},"<>",` +
         `${todoCol_('important')},${bool_(q.important)},` +
         `${todoCol_('urgent')},${bool_(q.urgent)},` +
         `${todoCol_('status')},"<>Done")`;
}

/** FILTER conditions selecting a quadrant's open tasks. */
function openTaskConditions_(q) {
  return [
    `${todoCol_('task')}<>""`,
    `${todoCol_('important')}=${bool_(q.important)}`,
    `${todoCol_('urgent')}=${bool_(q.urgent)}`,
    `${todoCol_('status')}<>"Done"`,
  ].join(',');
}

/**
 * Task / Due / Status for a quadrant, soonest due first (undated last).
 * When tasks exceed the slot count, the last slot shows "…and N more".
 */
function matrixListFormula_(q) {
  const S = CONFIG.slots;
  const cols = `{${todoCol_('task')},${todoCol_('due')},${todoCol_('status')}}`;
  return `=IFERROR(LET(t,SORT(FILTER(${cols},${openTaskConditions_(q)}),2,TRUE),n,ROWS(t),` +
         `IF(n<=${S},t,{ARRAY_CONSTRAIN(t,${S - 1},3);"${MORE_PREFIX}"&(n-${S - 1})&" more on TODO","",""})),` +
         `"${EMPTY_LIST_TEXT}")`;
}

/** Red, bold dates that are already past; range must list open tasks only. */
function overdueRule_(range) {
  const cell = range.getCell(1, 1).getA1Notation();
  return SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied(`=AND(ISNUMBER(${cell}),${cell}<TODAY())`)
    .setFontColor(THEME.danger)
    .setBold(true)
    .setRanges([range])
    .build();
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
