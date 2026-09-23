// ============================================================
// CONDITIONALFORMATTING.GS — DYNAMIC SAAS COLOR RULES
// ============================================================

/**
 * Configures semantic status and priority conditional formatting on a Month sheet.
 * Uses restrained SaaS pastel backgrounds with readable dark text defined in CONFIG.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 */
function setupConditionalFormatting(sheet) {
  const rules = [];
  const statusColors = CONFIG.COLORS.STATUS;
  const prioColors = CONFIG.COLORS.PRIORITY;

  // -------------------------
  // PRIORITY (Column F)
  // -------------------------

  rules.push(
    textRule(
      sheet.getRange('F2:F'),
      'Urgent',
      prioColors.URGENT.bg,
      prioColors.URGENT.text
    )
  );

  rules.push(
    textRule(
      sheet.getRange('F2:F'),
      'High',
      prioColors.HIGH.bg,
      prioColors.HIGH.text
    )
  );

  rules.push(
    textRule(
      sheet.getRange('F2:F'),
      'Medium',
      prioColors.MEDIUM.bg,
      prioColors.MEDIUM.text
    )
  );

  rules.push(
    textRule(
      sheet.getRange('F2:F'),
      'Low',
      prioColors.LOW.bg,
      prioColors.LOW.text
    )
  );

  // -------------------------
  // STATUS (Column G)
  // -------------------------

  rules.push(
    textRule(
      sheet.getRange('G2:G'),
      'In Progress',
      statusColors.IN_PROGRESS.bg,
      statusColors.IN_PROGRESS.text
    )
  );

  rules.push(
    textRule(
      sheet.getRange('G2:G'),
      'Completed',
      statusColors.COMPLETED.bg,
      statusColors.COMPLETED.text
    )
  );

  rules.push(
    textRule(
      sheet.getRange('G2:G'),
      'Blocked',
      statusColors.BLOCKED.bg,
      statusColors.BLOCKED.text
    )
  );

  rules.push(
    textRule(
      sheet.getRange('G2:G'),
      'Pending',
      statusColors.PENDING.bg,
      statusColors.PENDING.text
    )
  );

  rules.push(
    textRule(
      sheet.getRange('G2:G'),
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

  // Saturday: soft muted gray
  rules.push(
    SpreadsheetApp
      .newConditionalFormatRule()
      .whenFormulaSatisfied('=$B2="Sat"')
      .setBackground('#f8fafc')
      .setFontColor('#64748b')
      .setRanges([sheet.getRange('A2:G')])
      .build()
  );

  // Sunday: soft muted warm tint
  rules.push(
    SpreadsheetApp
      .newConditionalFormatRule()
      .whenFormulaSatisfied('=$B2="Sun"')
      .setBackground('#fef2f2')
      .setFontColor('#b91c1c')
      .setRanges([sheet.getRange('A2:G')])
      .build()
  );

  sheet.setConditionalFormatRules(rules);
}


/**
 * Applies visual highlight rules to the Eisenhower Matrix checkboxes in the Todo sheet.
 * When a checkbox is checked (TRUE), its cell receives that quadrant's accent color:
 * - Q1: Do (Urgent & Important) -> Soft Red
 * - Q2: Schedule (Important, Not Urgent) -> Soft Blue
 * - Q3: Delegate (Urgent, Not Important) -> Soft Yellow
 * - Q4: Don't Do (Neither) -> Soft Gray
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 */
function setupTodoConditionalFormatting(sheet) {
  const rules = [];

  // Q1: Do (Column C) — Soft Red
  rules.push(
    SpreadsheetApp
      .newConditionalFormatRule()
      .whenFormulaSatisfied('=$C3=TRUE')
      .setBackground(CONFIG.COLORS.Q1_BG)
      .setFontColor(CONFIG.COLORS.Q1_TEXT)
      .setRanges([sheet.getRange('C3:C')])
      .build()
  );

  // Q2: Schedule (Column D) — Soft Blue
  rules.push(
    SpreadsheetApp
      .newConditionalFormatRule()
      .whenFormulaSatisfied('=$D3=TRUE')
      .setBackground(CONFIG.COLORS.Q2_BG)
      .setFontColor(CONFIG.COLORS.Q2_TEXT)
      .setRanges([sheet.getRange('D3:D')])
      .build()
  );

  // Q3: Delegate (Column E) — Soft Yellow
  rules.push(
    SpreadsheetApp
      .newConditionalFormatRule()
      .whenFormulaSatisfied('=$E3=TRUE')
      .setBackground(CONFIG.COLORS.Q3_BG)
      .setFontColor(CONFIG.COLORS.Q3_TEXT)
      .setRanges([sheet.getRange('E3:E')])
      .build()
  );

  // Q4: Don't Do (Column F) — Soft Gray
  rules.push(
    SpreadsheetApp
      .newConditionalFormatRule()
      .whenFormulaSatisfied('=$F3=TRUE')
      .setBackground(CONFIG.COLORS.Q4_BG)
      .setFontColor(CONFIG.COLORS.Q4_TEXT)
      .setRanges([sheet.getRange('F3:F')])
      .build()
  );

  sheet.setConditionalFormatRules(rules);
}