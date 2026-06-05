function GeneralGmvTablazat() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getActiveSheet();
    var ordersSheet = ss.getSheetByName("Orders Coupon Nr");
    
    if (!ordersSheet) {
        SpreadsheetApp.getUi().alert('Hiba: Nem található "Orders Coupon Nr" nevű munkalap!');
        return;
    }

    sheet.clear();
    sheet.clearFormats();

    // ==================== NÉV LEKÉPEZÉS ====================
    const nameMap = {
        "2090951000172": "Viktor D.",
        "2090951001865": "Leila B.",
        "2090951000196": "Tibor K.",
        "2090951500474": "Barna R.",
        "2090951501938": "Dávid GY.",
        "2090951500504": "Krisztina T.",
        "2090953501530": "Máté Sz.",
        "2090953502025": "Tamás K.",
        "2090951001988": "Mercédesz F.",
        "2090953301352": "Zsolt Q.",
        "2090953302014": "Gábor K.",
        "2090950900107": "Tímea Sz.",
        "2090950901791": "Bence K.",
        "2090950900138": "Gábor T.",
        "2090954202030": "Tamás T.",
        "2090954202047": "Tamás J.",
        "2090954202054": "Balázs S.",
        "2090954202061": "Dániel F.",
        "2090951601966": "Gábor Cs.",
        "2090951601959": "Botond N.",
        "2090951600518": "Zsolt H.",
        "2090952401893": "Szilvia P.",
        "2090952400889": "Imre L.",
        "2090951001889": "Gabriella K.",
        "2090953701688": "Veronika Zs-B.",
        "2090953701602": "Tünde K.",
        "2090953701947": "Tünde W-H.",
        "2090953701879": "Zsanett Sz."
    };

    // ==================== ELSŐ TÁBLA (GMV STORES) ====================
    sheet.getRange("A1:B1").merge().setValue("GMV STORES");
    sheet.getRange("C1:D1").merge().setValue("GMV / PY/PL%");

    sheet.getRange("A2").setValue("TOP Stores");
    sheet.getRange("B2").setValue("GMV");
    sheet.getRange("C2").setValue("v. PY");
    sheet.getRange("D2").setValue("v. PL");

    // Első tábla adatai (tömegesen)
    const firstTableData = [];
    for (var i = 3; i <= 12; i++) {
        firstTableData.push([
            `= "T " & ROW(A${i-2}) & " - " & MID('Plan 2026'!A${i}; 4; 3)`,
            `=Sales!H${i + 3}`,
            "0",
            `=IFERROR( (B${i} / (IF(MONTH(TODAY())>1; SUM(OFFSET('Plan 2026'!B${i}; 0; 0; 1; MONTH(TODAY())-1)); 0) + (INDEX('Plan 2026'!B${i}:M${i}; 1; MONTH(TODAY())) / DAY(EOMONTH(TODAY(); 0))) * DAY(TODAY()))) - 1 ; 0)`
        ]);
    }
    sheet.getRange("A3:D12").setFormulas(firstTableData);

    sheet.getRange("A13").setValue("Total");
    sheet.getRange("B13").setFormula("=SUM(B3:B12)");
    sheet.getRange("C13").setFormula("0");
    sheet.getRange("D13").setFormula("=AVERAGE(D3:D12)");

    // ==================== MÁSODIK TÁBLA (TOP 20 DS) - TÖMEGESEN ====================
    const startCol = 6;  // F
    const startRow = 2;

    sheet.getRange(startRow, startCol, 1, 4).merge().setValue("TOP 20 DS");

    sheet.getRange(startRow + 1, startCol, 1, 4).setValues([["TOP 20 DS", "GMV", "v. PL", "Quota"]]);

    // TOP 20 DS adatok tömbként
    const top20Data = [];
    for (var i = 0; i < 20; i++) {
        var id = ordersSheet.getRange(5 + i, 1).getValue();
        var displayName = nameMap[id] || id;

        top20Data.push([
            displayName,
            `='Orders Coupon Nr'!G${5 + i}`,
            0,
            0
        ]);
    }

    // Egyetlen művelettel írjuk ki az egész táblát
    sheet.getRange(startRow + 2, startCol, 20, 4).setValues(top20Data);

    // Összesítő
    const totalRow = startRow + 22;
    sheet.getRange(totalRow, startCol).setValue("Total");
    sheet.getRange(totalRow, startCol + 1).setFormula("=SUM(G4:G23)");
    sheet.getRange(totalRow, startCol + 2).setValue(0);
    sheet.getRange(totalRow, startCol + 3).setValue(0);

    // ==================== FORMÁZÁSOK ====================
    sheet.getRange("A1:I" + totalRow)
        .setFontFamily("Arial")
        .setBorder(true, true, true, true, true, true, "#d9d9d9", SpreadsheetApp.BorderStyle.SOLID);

    // Fejlécek
    sheet.getRange("A1:D1").setFontWeight("bold").setHorizontalAlignment("center").setBackground("#ffffff");
    sheet.getRange("A2:D2").setFontWeight("bold").setHorizontalAlignment("center").setVerticalAlignment("middle")
          .setBackground("#4a86e8").setFontColor("#ffffff");   // A2 kék, de B2-D2 világosabb

    sheet.getRange("B2:D2").setBackground("#8fb4ff").setFontColor("#000000");

    sheet.getRange("F2:I2").setFontWeight("bold").setHorizontalAlignment("center").setBackground("#4a86e8").setFontColor("#ffffff");
    sheet.getRange("F3:I3").setFontWeight("bold").setHorizontalAlignment("center").setBackground("#8fb4ff").setFontColor("#000000");

    // Adatsorok
    sheet.getRange("A3:A12").setBackground("#8fb4ff").setFontWeight("bold").setHorizontalAlignment("center");
    sheet.getRange("F4:F23").setBackground("#8fb4ff").setFontWeight("bold").setHorizontalAlignment("center");

    sheet.getRange("B3:D12").setBackground("#efefef");
    sheet.getRange("G4:I23").setBackground("#efefef");

    sheet.getRange("A13:D13").setBackground("#8fb4ff").setFontWeight("bold");
    sheet.getRange("F24:I24").setBackground("#8fb4ff").setFontWeight("bold");

    // Formátumok
    sheet.getRange("B3:B13").setNumberFormat("#,##0");
    sheet.getRange("G4:G23").setNumberFormat("#,##0");
    sheet.getRange("C3:D13").setNumberFormat("0.00%");
    sheet.getRange("H4:I24").setNumberFormat("0");

    // Oszlopszélességek
    sheet.setColumnWidth(1, 140);
    sheet.setColumnWidth(2, 95);
    sheet.setColumnWidth(3, 70);
    sheet.setColumnWidth(4, 80);
    sheet.setColumnWidth(6, 160);
    sheet.setColumnWidth(7, 95);
    sheet.setColumnWidth(8, 70);
    sheet.setColumnWidth(9, 80);

    // Sor magasság
    for (var r = 1; r <= totalRow; r++) {
        sheet.setRowHeight(r, 28);
    }



    SpreadsheetApp.flush();
}
