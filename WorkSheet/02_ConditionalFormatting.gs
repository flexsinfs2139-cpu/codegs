function setupConditionalFormatting(sheet) {
  const rules = [];

  // -------------------------
  // PRIORITY
  // -------------------------

  rules.push(
    textRule(
      sheet.getRange('F2:F'),
      'Urgent',
      '#fce8e6',
      '#ff0000'
    )
  );

  rules.push(
    textRule(
      sheet.getRange('F2:F'),
      'High',
      '#fff2cc',
      '#b45f06'
    )
  );

  rules.push(
    textRule(
      sheet.getRange('F2:F'),
      'Medium',
      '#fff2cc',
      '#7f6000'
    )
  );

  // -------------------------
  // STATUS
  // -------------------------

  rules.push(
    textRule(
      sheet.getRange('G2:G'),
      'In Progress',
      '#cfe2f3',
      '#1155cc'
    )
  );

  rules.push(
    textRule(
      sheet.getRange('G2:G'),
      'Completed',
      '#d9ead3',
      '#008000'
    )
  );

  rules.push(
    textRule(
      sheet.getRange('G2:G'),
      'Blocked',
      '#f4cccc',
      '#cc0000'
    )
  );

  rules.push(
    textRule(
      sheet.getRange('G2:G'),
      'Cancelled',
      '#eeeeee',
      '#666666'
    )
  );

  sheet.setConditionalFormatRules(
    rules
  );
}


function textRule(
  range,
  text,
  background,
  fontColor
) {
  return SpreadsheetApp
    .newConditionalFormatRule()
    .whenTextEqualTo(text)
    .setBackground(background)
    .setFontColor(fontColor)
    .setRanges([range])
    .build();
}

function setupWeekendFormatting(sheet) {
  const rules = sheet.getConditionalFormatRules();

  // Saturday
  rules.push(
    SpreadsheetApp
      .newConditionalFormatRule()
      .whenFormulaSatisfied('=$B2="Sat"')
      .setBackground('#f3f4f6')
      .setFontColor('#6b7280')
      .setRanges([
        sheet.getRange('A2:G')
      ])
      .build()
  );

  // Sunday
  rules.push(
    SpreadsheetApp
      .newConditionalFormatRule()
      .whenFormulaSatisfied('=$B2="Sun"')
      .setBackground('#fce8e6')
      .setFontColor('#cc0000')
      .setRanges([
        sheet.getRange('A2:G')
      ])
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