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