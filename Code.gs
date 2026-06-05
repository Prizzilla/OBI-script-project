function GeneralGmvTablazat() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getActiveSheet();
    var ordersSheet = ss.getSheetByName("Orders Coupon Nr");
    var wgrSheet = ss.getSheetByName("WGR");
    var salesSheet = ss.getSheetByName("Sales");

    if (!ordersSheet) {
        SpreadsheetApp.getUi().alert('Hiba: Nem található "Orders Coupon Nr" nevű munkalap!');
        return;
    }
    if (!wgrSheet) {
        SpreadsheetApp.getUi().alert('Hiba: Nem található "WGR" nevű munkalap!');
        return;
    }
    if (!salesSheet) {
        SpreadsheetApp.getUi().alert('Hiba: Nem található "Sales" nevű munkalap!');
        return;
    }

    sheet.clear();
    sheet.clearFormats();

    // ==================== DINAMIKUS OSZLOPOK ====================
    const salesTotalCol   = getTotalColumn(salesSheet, 5, 3);      // Sales: C5-től
    const ordersTotalCol  = getTotalColumn(ordersSheet, 3, 2);     // Orders: B3-tól
    const wgrGmvCol       = getTotalColumn(wgrSheet, 4, 3);        // WGR GMV: C4-től
    const wgrProfitCol    = getProfitColumn(wgrSheet);             // WGR Profit: első üres cella előtti oszlop a 4. sorban

    // ==================== SZÍNEK ====================
    const darkBlue = "#4a86e8";
    const lightBlue = "#8fb4ff";

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
        "2090953701879": "Zsanett Sz.",
        "2090951500481": "Tarjányi Krisztina",
        "2090952400919":"Sólyom Balázs"
    };

    // ==================== ELSŐ TÁBLA (GMV STORES) ====================
    sheet.getRange("A1:B1").merge().setValue("GMV STORES");
    sheet.getRange("C1:D1").merge().setValue("GMV / PY/PL%");

    sheet.getRange("A2").setValue("TOP Stores");
    sheet.getRange("B2").setValue("GMV");
    sheet.getRange("C2").setValue("v. PY");
    sheet.getRange("D2").setValue("v. PL");

    const firstTableData = [];
    for (var i = 3; i <= 12; i++) {
        firstTableData.push([
            `= "T " & ROW(A${i-2}) & " - " & MID('Plan 2026'!A${i}; 4; 3)`,
            `=Sales!${salesTotalCol}${(i + 3)}`,
            "0",
            `=IFERROR( (B${i} / (IF(MONTH(TODAY())>1; SUM(OFFSET('Plan 2026'!B${i}; 0; 0; 1; MONTH(TODAY())-1)); 0) + (INDEX('Plan 2026'!B${i}:M${i}; 1; MONTH(TODAY())) / DAY(EOMONTH(TODAY(); 0))) * DAY(TODAY()))) - 1 ; 0)`
        ]);
    }
    sheet.getRange("A3:D12").setFormulas(firstTableData);

    sheet.getRange("A13").setValue("Total");
    sheet.getRange("B13").setFormula("=SUM(B3:B12)");
    sheet.getRange("C13").setFormula("0");
    sheet.getRange("D13").setFormula("=AVERAGE(D3:D12)");

    // ==================== MÁSODIK TÁBLA (TOP 20 DS) ====================
    const startCol2 = 6;
    const startRow2 = 2;

    sheet.getRange(startRow2, startCol2, 1, 2).merge().setValue("GMV Designers");
    sheet.getRange(startRow2, startCol2 + 2, 1, 2).merge().setValue("GMV / PY/PL%");

    sheet.getRange(startRow2 + 1, startCol2, 1, 4).setValues([["TOP 20 DS", "GMV", "v. PL", "Quota"]]);

    const top20Data = [];
    for (var i = 0; i < 20; i++) {
        var id = ordersSheet.getRange(5 + i, 1).getValue();
        var displayName = nameMap[id] || id;
        top20Data.push([displayName, `='Orders Coupon Nr'!${ordersTotalCol}${5 + i}`, 0, 0]);
    }
    sheet.getRange(startRow2 + 2, startCol2, 20, 4).setValues(top20Data);

    const totalRow2 = startRow2 + 22;
    sheet.getRange(totalRow2, startCol2).setValue("Total");
    sheet.getRange(totalRow2, startCol2 + 1).setFormula("=SUM(G4:G23)");
    sheet.getRange(totalRow2, startCol2 + 2).setValue(0);
    sheet.getRange(totalRow2, startCol2 + 3).setValue(0);

    // ==================== HARMADIK TÁBLA (WGR) ====================
    const startCol3 = 11;
    const startRow3 = 1;

    sheet.getRange(startRow3, startCol3, 1, 2).merge().setValue("CAT49");
    sheet.getRange(startRow3, startCol3 + 2, 1, 2).merge().setValue("GMV / PY/PL%");

    sheet.getRange(startRow3 + 1, startCol3).setValue("WGR");
    sheet.getRange(startRow3 + 1, startCol3 + 1).setValue("GMV");
    sheet.getRange(startRow3 + 1, startCol3 + 2).setValue("Margin");
    sheet.getRange(startRow3 + 1, startCol3 + 3).setValue("Profit");

    const wgrData = [];
    for (var i = 0; i < 11; i++) {
        var r = startRow3 + 2 + i;   // a fő sheet sor száma (3-tól kezdődik)

        var germanName = wgrSheet.getRange("B" + (5 + i)).getValue().toString().trim();
        var englishName = germanName ? LanguageApp.translate(germanName, 'de', 'en') : germanName;

        wgrData.push([
            englishName,
            `='WGR'!${wgrGmvCol}${(5 + i)}`,
            `=IFERROR(IF(L${r}=0; 0; N${r} / L${r}); 0)`,     // ← Margin most a saját L és N oszlopból
            `='WGR'!${wgrProfitCol}${(5 + i)}`
        ]);
    }
    sheet.getRange(startRow3 + 2, startCol3, 11, 4).setValues(wgrData);

    const totalRow3 = startRow3 + 13;
    sheet.getRange(totalRow3, startCol3).setValue("Total");
    sheet.getRange(totalRow3, startCol3 + 1).setFormula("=SUM(L3:L13)");
    sheet.getRange(totalRow3, startCol3 + 2).setFormula("=IFERROR(IF(L14=0;0; N14 / L14); 0)");
    sheet.getRange(totalRow3, startCol3 + 3).setFormula("=SUM(N3:N13)");

    // ==================== FORMÁZÁSOK ====================
    const maxRow = Math.max(totalRow2, totalRow3);

    sheet.getRange("A1:N" + maxRow)
        .setFontFamily("Arial")
        .setBorder(true, true, true, true, true, true, "#d9d9d9", SpreadsheetApp.BorderStyle.SOLID);

    sheet.getRange("A1:D1").setFontSize(14).setHorizontalAlignment("center").setFontWeight("bold");
    sheet.getRange("F2:I2").setFontSize(14).setHorizontalAlignment("center").setFontWeight("bold");
    sheet.getRange("K1:N1").setFontSize(14).setHorizontalAlignment("center").setFontWeight("bold");

    sheet.getRange("A1:B1").setBackground(darkBlue).setFontColor("#ffffff");
    sheet.getRange("F3").setBackground(darkBlue).setFontColor("#ffffff").setFontWeight("bold").setHorizontalAlignment("center");
    sheet.getRange("K1:L1").setBackground(darkBlue).setFontColor("#ffffff");
    sheet.getRange("L2:N2").setBackground(lightBlue).setFontColor("#000000").setFontWeight("bold").setHorizontalAlignment("center");
    sheet.getRange("K2").setBackground(darkBlue).setFontColor("#ffffff").setFontWeight("bold").setHorizontalAlignment("center");

    sheet.getRange("A2:D2").setFontWeight("bold").setHorizontalAlignment("center").setVerticalAlignment("middle");
    sheet.getRange("A2").setBackground(darkBlue).setFontColor("#ffffff");
    sheet.getRange("B2:D2").setBackground(lightBlue).setFontColor("#000000");

    sheet.getRange("F2:G2").setBackground(darkBlue).setFontColor("#ffffff").setFontWeight("bold");
    sheet.getRange("H2:I2").setBackground("#ffffff").setFontColor("#000000").setFontWeight("bold");
    sheet.getRange("G3:I3").setFontWeight("bold").setHorizontalAlignment("center").setBackground(lightBlue).setFontColor("#000000");

    sheet.getRange("A3:A12").setBackground(lightBlue).setFontWeight("bold").setHorizontalAlignment("center");
    sheet.getRange("F4:F23").setBackground(lightBlue).setFontWeight("bold").setHorizontalAlignment("center");
    sheet.getRange("K3:K13").setBackground(lightBlue).setFontWeight("bold").setHorizontalAlignment("center").setFontSize(9);

    sheet.getRange("B3:B12").setBackground("#ffffff");
    sheet.getRange("G4:G23").setBackground("#ffffff");
    sheet.getRange("L3:L13").setBackground("#ffffff");
    sheet.getRange("N3:N13").setBackground("#ffffff");

    sheet.getRange("C3:D12").setBackground("#efefef");
    sheet.getRange("H4:I23").setBackground("#efefef");
    sheet.getRange("M3:M13").setBackground("#efefef");

    sheet.getRange("A13").setBackground(darkBlue).setFontColor("#ffffff").setFontWeight("bold");
    sheet.getRange("B13:D13").setBackground(lightBlue).setFontColor("#000000").setFontWeight("bold");

    sheet.getRange("F" + totalRow2).setBackground(darkBlue).setFontColor("#ffffff").setFontWeight("bold");
    sheet.getRange("G" + totalRow2 + ":I" + totalRow2).setBackground(lightBlue).setFontColor("#000000").setFontWeight("bold");

    sheet.getRange("K" + totalRow3).setBackground(darkBlue).setFontColor("#ffffff").setFontWeight("bold");
    sheet.getRange("L" + totalRow3 + ":N" + totalRow3).setBackground(lightBlue).setFontColor("#000000").setFontWeight("bold");

    sheet.getRange("B3:B13").setNumberFormat("#,##0").setHorizontalAlignment("right");
    sheet.getRange("G4:G23").setNumberFormat("#,##0").setHorizontalAlignment("right");
    sheet.getRange("L3:L" + totalRow3).setNumberFormat("#,##0").setHorizontalAlignment("right");
    sheet.getRange("N3:N" + totalRow3).setNumberFormat("#,##0").setHorizontalAlignment("right");

    sheet.getRange("C3:D13").setNumberFormat("0.00%").setHorizontalAlignment("right");
    sheet.getRange("M3:M" + totalRow3).setNumberFormat("0.0%").setHorizontalAlignment("right");

    sheet.setColumnWidth(1, 140);
    sheet.setColumnWidth(2, 95);
    sheet.setColumnWidth(3, 70);
    sheet.setColumnWidth(4, 80);
    sheet.setColumnWidth(6, 160);
    sheet.setColumnWidth(7, 95);
    sheet.setColumnWidth(8, 70);
    sheet.setColumnWidth(9, 80);
    sheet.setColumnWidth(11, 180);
    sheet.setColumnWidth(12, 95);
    sheet.setColumnWidth(13, 80);
    sheet.setColumnWidth(14, 95);

    for (var r = 1; r <= maxRow; r++) {
        sheet.setRowHeight(r, 28);
    }

    SpreadsheetApp.flush();
}

// ==================== SEGÉDFÜGGVÉNYEK ====================
function getTotalColumn(sheet, startRow, startCol) {
    let col = startCol;
    while (true) {
        const value = sheet.getRange(startRow, col).getValue().toString().trim();
        if (value === "Total") {
            break;
        }
        col++;
        if (col > 200) break;
    }
    return columnToLetter(col);
}

function getProfitColumn(sheet) {
    let col = 3;
    while (true) {
        const value = sheet.getRange(4, col).getValue().toString().trim();
        if (value === "") {
            col--;
            break;
        }
        col++;
        if (col > 200) {
            col = 14;
            break;
        }
    }
    return columnToLetter(col);
}

function columnToLetter(column) {
    let temp, letter = '';
    while (column > 0) {
        temp = (column - 1) % 26;
        letter = String.fromCharCode(temp + 65) + letter;
        column = Math.floor((column - temp - 1) / 26);
    }
    return letter;
}
