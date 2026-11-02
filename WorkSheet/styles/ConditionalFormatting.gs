// ============================================================
// CONDITIONALFORMATTING.GS — DYNAMIC SAAS COLOR RULES
// ============================================================

/**
 * Configures semantic status and priority conditional formatting on an 8-column Month sheet.
 * Priority: Column G
 * Status: Column H
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 */
function setupConditionalFormatting(sheet) {
  const rules = [];
  const statusColors = CONFIG.COLORS.STATUS;
  const prioColors = CONFIG.COLORS.PRIORITY;

  // -------------------------
  // PRIORITY (Column G)
  // -------------------------

  rules.push(
    textRule(
      sheet.getRange('G2:G'),
      'Urgent',
      prioColors.URGENT.bg,
      prioColors.URGENT.text
    )
  );

  rules.push(
    textRule(
      sheet.getRange('G2:G'),
      'High',
      prioColors.HIGH.bg,
      prioColors.HIGH.text
    )
  );

  rules.push(
    textRule(
      sheet.getRange('G2:G'),
      'Medium',
      prioColors.MEDIUM.bg,
      prioColors.MEDIUM.text
    )
  );

  rules.push(
    textRule(
      sheet.getRange('G2:G'),
      'Low',
      prioColors.LOW.bg,
      prioColors.LOW.text
    )
  );

  // -------------------------
  // STATUS (Column H)
  // -------------------------

  rules.push(
    textRule(
      sheet.getRange('H2:H'),
      'In Progress',
      statusColors.IN_PROGRESS.bg,
      statusColors.IN_PROGRESS.text
    )
  );

  rules.push(
    textRule(
      sheet.getRange('H2:H'),
      'Completed',
      statusColors.COMPLETED.bg,
      statusColors.COMPLETED.text
    )
  );

  rules.push(
    textRule(
      sheet.getRange('H2:H'),
      'Blocked',
      statusColors.BLOCKED.bg,
      statusColors.BLOCKED.text
    )
  );

  rules.push(
    textRule(
      sheet.getRange('H2:H'),
      'Pending',
      statusColors.PENDING.bg,
      statusColors.PENDING.text
    )
  );

  rules.push(
    textRule(
      sheet.getRange('H2:H'),
      'Cancelled',
      statusColors.CANCELLED.bg,
      statusColors.CANCELLED.text
    )
  );

  sheet.setConditionalFormatRules(rules);
}


/**
 * Helper to build an exact-match text conditional formatting rule.
 */
function textRule(range, text, background, fontColor) {
  return SpreadsheetApp
    .newConditionalFormatRule()
    .whenTextEqualTo(text)
    .setBackground(background)
    .setFontColor(fontColor)
    .setRanges([range])
    .build();
}


/**
 * Highlights Saturday and Sunday rows with subtle, restrained weekend shading.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 */
function setupWeekendFormatting(sheet) {
  const rules = sheet.getConditionalFormatRules() || [];

  // Saturday: soft muted gray across columns A to H
  rules.push(
    SpreadsheetApp
      .newConditionalFormatRule()
      .whenFormulaSatisfied('=$B2="Sat"')
      .setBackground('#f8fafc')
      .setFontColor('#64748b')
      .setRanges([sheet.getRange('A2:H')])
      .build()
  );

  // Sunday: soft muted warm tint across columns A to H
  rules.push(
    SpreadsheetApp
      .newConditionalFormatRule()
      .whenFormulaSatisfied('=$B2="Sun"')
      .setBackground('#fef2f2')
      .setFontColor('#b91c1c')
      .setRanges([sheet.getRange('A2:H')])
      .build()
  );

  sheet.setConditionalFormatRules(rules);
}