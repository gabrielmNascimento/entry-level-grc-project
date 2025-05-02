function styleRiskRegister() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Sheet1");
  const lastHeader = "L1"
  
  // Style headers
  const headerRange = sheet.getRange(`A1:${lastHeader}`);
  headerRange
    .setBackground("#2c3e50")
    .setFontColor("white")
    .setFontWeight("bold")
    .setHorizontalAlignment("center");

  // Apply alternating row colors
  const lastRow = sheet.getLastRow();
  for (let row = 2; row <= lastRow; row++) {
    const rowRange = sheet.getRange(`A${row}:I${row}`);
    rowRange.setBackground(row % 2 === 0 ? "#f9f9f9" : "#ffffff"); // Light gray/white
  }

  // Auto-resize columns
  sheet.autoResizeColumns(1, 9);

  // 4. Conditional formatting for Risk Level
  const riskLevelRange = sheet.getRange("H2:H" + lastRow);
  const rules = sheet.getConditionalFormatRules();
  
  rules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo("Critical")
      .setBackground("#ff0000") // Red
      .setFontColor("white")
      .setBold(true)
      .setRanges([riskLevelRange])
      .build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo("High")
      .setBackground("#ff9900") // Orange
      .setRanges([riskLevelRange])
      .build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo("Medium")
      .setBackground("#ffff00") // Yellow
      .setRanges([riskLevelRange])
      .build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo("Low")
      .setBackground("#00cc00") // Green
      .setRanges([riskLevelRange])
      .build()
  );
  
  sheet.setConditionalFormatRules(rules);

  // Calculate Risk Score (Column G = Likelihood * Impact)
  const scoreFormulaRange = sheet.getRange("G2:G" + lastRow);
  scoreFormulaRange.setFormula('=VALUE(REGEXEXTRACT(E2, "\\d+")) * VALUE(REGEXEXTRACT(F2, "\\d+"))');

  // Auto-set Risk Level
  const levelFormulaRange = sheet.getRange("H2:H" + lastRow);
  levelFormulaRange.setFormula('=IFS(G2>=7, "Critical", G2>=5, "High", G2>=3, "Medium", TRUE, "Low")');

  // Wrap text in long columns
  sheet.getRange("B2:B" + lastRow).setWrap(true);
  sheet.getRange("I2:I" + lastRow).setWrap(true);
  sheet.getRange("J2:J" + lastRow).setWrap(true);

  // Freeze header row
  sheet.setFrozenRows(1);
}
