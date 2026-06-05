function GeneralGmvTablazat() {

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  sheet.clear();
  sheet.clearFormats();

  // === FEJLÉC ===
  sheet.getRange("A1:B1").merge();
  sheet.getRange("C1:D1").merge();

  sheet.getRange("A1").setValue("GMV STORES");
  sheet.getRange("C1").setValue("GMV / PY/PL%");

  sheet.getRange("A2").setValue("TOP Stores");
  sheet.getRange("B2").setValue("GMV");
  sheet.getRange("C2").setValue("v. PY");
  sheet.getRange("D2").setValue("v. PL");

  // === FŐ CIKLUS ===
  for (var i = 3; i <= 12; i++) {

    // A oszlop
    sheet.getRange("A" + i).setFormula(
      '="T " & ROW(A' + (i - 2) + ') & " - " & MID(\'Plan 2026\'!A' + i + '; 4; 3)'
    );

    // B oszlop
    sheet.getRange("B" + i).setFormula(
      '=Sales!H' + (i + 3)
    );

    // C oszlop
    sheet.getRange("C" + i).setValue(0);

    // D oszlop
    sheet.getRange("D" + i).setFormula(
      '=IFERROR( (B' + i + ' / (IF(MONTH(TODAY())>1; SUM(OFFSET(\'Plan 2026\'!B' + i + '; 0; 0; 1; MONTH(TODAY())-1)); 0) ' +
      '+ (INDEX(\'Plan 2026\'!B' + i + ':M' + i + '; 1; MONTH(TODAY())) / DAY(EOMONTH(TODAY(); 0))) * DAY(TODAY()))) - 1 ; 0)'
    );
  }

  // === ÖSSZESÍTŐ ===
  sheet.getRange("A13").setValue("Total");        // ← biztosan "Total"
  sheet.getRange("B13").setFormula("=SUM(B3:B12)");
  sheet.getRange("C13").setFormula("=SUM(C3:C12)");
  sheet.getRange("D13").setFormula("=AVERAGE(D3:D12)");

  // === FORMAZÁSOK ===
  sheet.getRange("A1:D13")
    .setFontFamily("Arial")
    .setBorder(true, true, true, true, true, true, "#d9d9d9", SpreadsheetApp.BorderStyle.SOLID);

  sheet.getRange("A1:D1")
    .setFontWeight("bold")
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle")
    .setBackground("#ffffff");

  sheet.getRange("A2:D2")
    .setFontWeight("bold")
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle");

  sheet.getRange("A2").setBackground("#4a86e8").setFontColor("#ffffff");
  
  // Világosabb kék: RGB(143, 180, 255) → #8fb4ff
  sheet.getRange("B2:D2").setBackground("#8fb4ff").setFontColor("#000000");
  
  // A3:A12 is ugyanaz a világosabb kék szín
  sheet.getRange("A3:A12")
    .setBackground("#8fb4ff")
    .setFontWeight("bold")
    .setHorizontalAlignment("center");

  sheet.getRange("B3:D12").setBackground("#efefef");
  sheet.getRange("A13:D13").setBackground("#8fb4ff").setFontWeight("bold");

  sheet.getRange("B3:B13").setNumberFormat("#,##0").setHorizontalAlignment("right");
  sheet.getRange("C3:C13").setNumberFormat("0").setHorizontalAlignment("right");
  sheet.getRange("D3:D13").setNumberFormat("0.00%").setHorizontalAlignment("right");

  sheet.setColumnWidth(1, 140);
  sheet.setColumnWidth(2, 95);
  sheet.setColumnWidth(3, 70);
  sheet.setColumnWidth(4, 80);

  for (var r = 1; r <= 13; r++) {
    sheet.setRowHeight(r, 28);
  }

  // === ÚJRAÉLESZTÉS ===
  SpreadsheetApp.flush();
  Utilities.sleep(400);

  var range = sheet.getRange("A3:D13");
  var formulas = range.getFormulas();
  range.setFormulas(formulas);

  SpreadsheetApp.flush();
}
