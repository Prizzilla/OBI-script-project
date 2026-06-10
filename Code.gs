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
    const salesTotalCol   = getTotalColumn(salesSheet, 5, 3);
    const ordersTotalCol  = getTotalColumn(ordersSheet, 3, 2);
    const wgrGmvCol       = getTotalColumn(wgrSheet, 4, 3);
    const wgrProfitCol    = getProfitColumn(wgrSheet);


    // ==================== ELŐZŐ HÓNAP DINAMIKUS OSZLOPOK ====================
    const previousMonth = getPreviousMonthInfo();
    const hasPreviousMonth = previousMonth.month > 0;

    const salesPrevMonthCol = hasPreviousMonth
        ? getMonthColumnByHeader(salesSheet, 5, 3, previousMonth.yyyymm, false)
        : null;

    const ordersPrevMonthCol = hasPreviousMonth
        ? getMonthColumnByHeader(ordersSheet, 3, 2, previousMonth.yyyymm, false)
        : null;

    const wgrPrevGmvCol = hasPreviousMonth
        ? getMonthColumnByHeader(wgrSheet, 4, 3, previousMonth.yyyymm, false)
        : null;

    const wgrPrevProfitCol = hasPreviousMonth
        ? getMonthColumnByHeader(wgrSheet, 4, 3, previousMonth.yyyymm, true)
        : null;

    const planPrevMonthCol = hasPreviousMonth
        ? columnToLetter(previousMonth.month + 1)
        : null;

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
        "2090951500481": "Semsei Imre",
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
        var r = startRow3 + 2 + i;
        var germanName = wgrSheet.getRange("B" + (5 + i)).getValue().toString().trim();
        var englishName = germanName ? LanguageApp.translate(germanName, 'de', 'en') : germanName;

        wgrData.push([
            englishName,
            `='WGR'!${wgrGmvCol}${(5 + i)}`,
            `=IFERROR(IF(L${r}=0; 0; N${r} / L${r}); 0)`,
            `='WGR'!${wgrProfitCol}${(5 + i)}`
        ]);
    }
    sheet.getRange(startRow3 + 2, startCol3, 11, 4).setValues(wgrData);

    const totalRow3 = startRow3 + 13;
    sheet.getRange(totalRow3, startCol3).setValue("Total");
    sheet.getRange(totalRow3, startCol3 + 1).setFormula("=SUM(L3:L13)");
    sheet.getRange(totalRow3, startCol3 + 2).setFormula("=IFERROR(IF(L14=0;0; N14 / L14); 0)");
    sheet.getRange(totalRow3, startCol3 + 3).setFormula("=SUM(N3:N13)");

    // ==================== PLUSZ ELSŐ TÁBLA - ELŐZŐ HÓNAP (GMV STORES) ====================
    const prevStartRow1 = 30;

    sheet.getRange(prevStartRow1, 1, 1, 2).merge().setValue("GMV STORES");
    sheet.getRange(prevStartRow1, 3, 1, 2).merge().setValue("GMV / PY/PL%");

    sheet.getRange(prevStartRow1 + 1, 1).setValue("TOP Stores");
    sheet.getRange(prevStartRow1 + 1, 2).setValue("GMV");
    sheet.getRange(prevStartRow1 + 1, 3).setValue("v. PY");
    sheet.getRange(prevStartRow1 + 1, 4).setValue("v. PL");

    const prevFirstTableData = [];

for (var i = 0; i < 10; i++) {
    var planRow = 3 + i;
    var salesRow = 6 + i;

    var gmvValue = hasPreviousMonth && salesPrevMonthCol
        ? Number(salesSheet.getRange(salesRow, letterToColumn(salesPrevMonthCol)).getValue()) || 0
        : 0;

    prevFirstTableData.push({
        gmv: gmvValue,
        planRow: planRow,
        salesRow: salesRow
    });
}

prevFirstTableData.sort(function(a, b) {
    return b.gmv - a.gmv;
});

var sortedPrevFirstRows = prevFirstTableData.map(function(item, index) {
    var targetRow = prevStartRow1 + 2 + index;

    return [
        `= "T " & ${index + 1} & " - " & MID('Plan 2026'!A${item.planRow}; 4; 3)`,
        hasPreviousMonth && salesPrevMonthCol ? `=Sales!${salesPrevMonthCol}${item.salesRow}` : "0",
        "0",
        hasPreviousMonth && planPrevMonthCol
            ? `=IFERROR(B${targetRow} / 'Plan 2026'!${planPrevMonthCol}${item.planRow} - 1; 0)`
            : "0"
    ];
});

sheet.getRange(prevStartRow1 + 2, 1, 10, 4).setFormulas(sortedPrevFirstRows);

    const prevTotalRow1 = prevStartRow1 + 12;
    sheet.getRange(prevTotalRow1, 1).setValue("Total");
    sheet.getRange(prevTotalRow1, 2).setFormula(`=SUM(B${prevStartRow1 + 2}:B${prevStartRow1 + 11})`);
    sheet.getRange(prevTotalRow1, 3).setFormula("0");
    sheet.getRange(prevTotalRow1, 4).setFormula(`=AVERAGE(D${prevStartRow1 + 2}:D${prevStartRow1 + 11})`);

    // ==================== PLUSZ MÁSODIK TÁBLA - ELŐZŐ HÓNAP (TOP 20 DS) ====================
    const prevStartCol2 = 6;
    const prevStartRow2 = 30;

    sheet.getRange(prevStartRow2, prevStartCol2, 1, 2).merge().setValue("GMV Designers");
    sheet.getRange(prevStartRow2, prevStartCol2 + 2, 1, 2).merge().setValue("GMV / PY/PL%");

    sheet.getRange(prevStartRow2 + 1, prevStartCol2, 1, 4).setValues([["TOP 20 DS", "GMV", "v. PL", "Quota"]]);

   const prevTop20Data = [];

for (var i = 0; i < 20; i++) {
    var sourceRow = 5 + i;
    var id = ordersSheet.getRange(sourceRow, 1).getValue();
    var displayName = nameMap[id] || id;

    var gmvValue = hasPreviousMonth && ordersPrevMonthCol
        ? Number(ordersSheet.getRange(sourceRow, letterToColumn(ordersPrevMonthCol)).getValue()) || 0
        : 0;

    prevTop20Data.push({
        gmv: gmvValue,
        row: [
            displayName,
            hasPreviousMonth && ordersPrevMonthCol ? `='Orders Coupon Nr'!${ordersPrevMonthCol}${sourceRow}` : "0",
            0,
            0
        ]
    });
}

prevTop20Data.sort(function(a, b) {
    return b.gmv - a.gmv;
});

sheet.getRange(prevStartRow2 + 2, prevStartCol2, 20, 4).setValues(
    prevTop20Data.map(function(item) {
        return item.row;
    })
);

    const prevTotalRow2 = prevStartRow2 + 22;
    sheet.getRange(prevTotalRow2, prevStartCol2).setValue("Total");
    sheet.getRange(prevTotalRow2, prevStartCol2 + 1).setFormula(`=SUM(G${prevStartRow2 + 2}:G${prevStartRow2 + 21})`);
    sheet.getRange(prevTotalRow2, prevStartCol2 + 2).setValue(0);
    sheet.getRange(prevTotalRow2, prevStartCol2 + 3).setValue(0);

    // ==================== PLUSZ HARMADIK TÁBLA - ELŐZŐ HÓNAP (WGR) ====================
    const prevStartCol3 = 11;
    const prevStartRow3 = 30;

    sheet.getRange(prevStartRow3, prevStartCol3, 1, 2).merge().setValue("CAT49");
    sheet.getRange(prevStartRow3, prevStartCol3 + 2, 1, 2).merge().setValue("GMV / PY/PL%");

    sheet.getRange(prevStartRow3 + 1, prevStartCol3).setValue("WGR");
    sheet.getRange(prevStartRow3 + 1, prevStartCol3 + 1).setValue("GMV");
    sheet.getRange(prevStartRow3 + 1, prevStartCol3 + 2).setValue("Margin");
    sheet.getRange(prevStartRow3 + 1, prevStartCol3 + 3).setValue("Profit");

    const prevWgrData = [];

for (var i = 0; i < 11; i++) {
    var sourceRow = 5 + i;
    var germanName = wgrSheet.getRange("B" + sourceRow).getValue().toString().trim();
    var englishName = germanName ? LanguageApp.translate(germanName, 'de', 'en') : germanName;

    var gmvValue = hasPreviousMonth && wgrPrevGmvCol
        ? Number(wgrSheet.getRange(sourceRow, letterToColumn(wgrPrevGmvCol)).getValue()) || 0
        : 0;

    prevWgrData.push({
        gmv: gmvValue,
        sourceRow: sourceRow,
        name: englishName
    });
}

prevWgrData.sort(function(a, b) {
    return b.gmv - a.gmv;
});

var sortedPrevWgrRows = prevWgrData.map(function(item, index) {
    var targetRow = prevStartRow3 + 2 + index;

    return [
        item.name,
        hasPreviousMonth && wgrPrevGmvCol ? `='WGR'!${wgrPrevGmvCol}${item.sourceRow}` : "0",
        `=IFERROR(IF(L${targetRow}=0; 0; N${targetRow} / L${targetRow}); 0)`,
        hasPreviousMonth && wgrPrevProfitCol ? `='WGR'!${wgrPrevProfitCol}${item.sourceRow}` : "0"
    ];
});

sheet.getRange(prevStartRow3 + 2, prevStartCol3, 11, 4).setValues(sortedPrevWgrRows);

    const prevTotalRow3 = prevStartRow3 + 13;
    sheet.getRange(prevTotalRow3, prevStartCol3).setValue("Total");
    sheet.getRange(prevTotalRow3, prevStartCol3 + 1).setFormula(`=SUM(L${prevStartRow3 + 2}:L${prevStartRow3 + 12})`);
    sheet.getRange(prevTotalRow3, prevStartCol3 + 2).setFormula(`=IFERROR(IF(L${prevTotalRow3}=0;0; N${prevTotalRow3} / L${prevTotalRow3}); 0)`);
    sheet.getRange(prevTotalRow3, prevStartCol3 + 3).setFormula(`=SUM(N${prevStartRow3 + 2}:N${prevStartRow3 + 12})`);

    // ==================== FORMÁZÁSOK ====================
    const maxRow = Math.max(totalRow2, totalRow3, prevTotalRow1, prevTotalRow2, prevTotalRow3);

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

    // ==================== ELŐZŐ HÓNAP TÁBLÁK FORMÁZÁSA ====================
    sheet.getRange("A30:D30").setFontSize(14).setHorizontalAlignment("center").setFontWeight("bold");
    sheet.getRange("F30:I30").setFontSize(14).setHorizontalAlignment("center").setFontWeight("bold");
    sheet.getRange("K30:N30").setFontSize(14).setHorizontalAlignment("center").setFontWeight("bold");

    sheet.getRange("A30:B30").setBackground(darkBlue).setFontColor("#ffffff");
    sheet.getRange("C30:D30").setBackground(lightBlue).setFontColor("#000000").setFontWeight("bold");

    sheet.getRange("A31:D31").setFontWeight("bold").setHorizontalAlignment("center").setVerticalAlignment("middle");
    sheet.getRange("A31").setBackground(darkBlue).setFontColor("#ffffff");
    sheet.getRange("B31:D31").setBackground(lightBlue).setFontColor("#000000");

    sheet.getRange("A32:A41").setBackground(lightBlue).setFontWeight("bold").setHorizontalAlignment("center");
    sheet.getRange("B32:B41").setBackground("#ffffff");
    sheet.getRange("C32:D41").setBackground("#efefef");

    sheet.getRange("A" + prevTotalRow1).setBackground(darkBlue).setFontColor("#ffffff").setFontWeight("bold");
    sheet.getRange("B" + prevTotalRow1 + ":D" + prevTotalRow1).setBackground(lightBlue).setFontColor("#000000").setFontWeight("bold");

    sheet.getRange("F30:G30").setBackground(darkBlue).setFontColor("#ffffff").setFontWeight("bold");
    sheet.getRange("H30:I30").setBackground("#ffffff").setFontColor("#000000").setFontWeight("bold");

    sheet.getRange("F31").setBackground(darkBlue).setFontColor("#ffffff").setFontWeight("bold").setHorizontalAlignment("center");
    sheet.getRange("G31:I31").setFontWeight("bold").setHorizontalAlignment("center").setBackground(lightBlue).setFontColor("#000000");

    sheet.getRange("F32:F51").setBackground(lightBlue).setFontWeight("bold").setHorizontalAlignment("center");
    sheet.getRange("G32:G51").setBackground("#ffffff");
    sheet.getRange("H32:I51").setBackground("#efefef");

    sheet.getRange("F" + prevTotalRow2).setBackground(darkBlue).setFontColor("#ffffff").setFontWeight("bold");
    sheet.getRange("G" + prevTotalRow2 + ":I" + prevTotalRow2).setBackground(lightBlue).setFontColor("#000000").setFontWeight("bold");

    sheet.getRange("K30:L30").setBackground(darkBlue).setFontColor("#ffffff");
    sheet.getRange("M30:N30").setBackground(lightBlue).setFontColor("#000000").setFontWeight("bold");

    sheet.getRange("K31").setBackground(darkBlue).setFontColor("#ffffff").setFontWeight("bold").setHorizontalAlignment("center");
    sheet.getRange("L31:N31").setBackground(lightBlue).setFontColor("#000000").setFontWeight("bold").setHorizontalAlignment("center");

    sheet.getRange("K32:K42").setBackground(lightBlue).setFontWeight("bold").setHorizontalAlignment("center").setFontSize(9);
    sheet.getRange("L32:L42").setBackground("#ffffff");
    sheet.getRange("N32:N42").setBackground("#ffffff");
    sheet.getRange("M32:M42").setBackground("#efefef");

    sheet.getRange("K" + prevTotalRow3).setBackground(darkBlue).setFontColor("#ffffff").setFontWeight("bold");
    sheet.getRange("L" + prevTotalRow3 + ":N" + prevTotalRow3).setBackground(lightBlue).setFontColor("#000000").setFontWeight("bold");

    sheet.getRange("B3:B13").setNumberFormat("#,##0").setHorizontalAlignment("right");
    sheet.getRange("G4:G23").setNumberFormat("#,##0").setHorizontalAlignment("right");
    sheet.getRange("L3:L" + totalRow3).setNumberFormat("#,##0").setHorizontalAlignment("right");
    sheet.getRange("N3:N" + totalRow3).setNumberFormat("#,##0").setHorizontalAlignment("right");

    sheet.getRange("C3:D13").setNumberFormat("0.00%").setHorizontalAlignment("right");
    sheet.getRange("M3:M" + totalRow3).setNumberFormat("0.0%").setHorizontalAlignment("right");

    sheet.getRange("B32:B" + prevTotalRow1).setNumberFormat("#,##0").setHorizontalAlignment("right");
    sheet.getRange("G32:G" + prevTotalRow2).setNumberFormat("#,##0").setHorizontalAlignment("right");
    sheet.getRange("L32:L" + prevTotalRow3).setNumberFormat("#,##0").setHorizontalAlignment("right");
    sheet.getRange("N32:N" + prevTotalRow3).setNumberFormat("#,##0").setHorizontalAlignment("right");

    sheet.getRange("C32:D" + prevTotalRow1).setNumberFormat("0.00%").setHorizontalAlignment("right");
    sheet.getRange("M32:M" + prevTotalRow3).setNumberFormat("0.0%").setHorizontalAlignment("right");

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

    // ==================== CHARTDATA + DIAGRAM ====================
       
    updateChartData();        // a régi diagram
    createHUMonthlyChart();   // ← az új diagram

   // PPT formázás
if (sheet.getName() === "Táblázatok") {

    sheet.getDataRange()
        .setFontSize(7)
        .setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP)
        .setVerticalAlignment("middle");

    for (var r = 1; r <= sheet.getMaxRows(); r++) {
        sheet.setRowHeight(r, 21);
    }

    var lastCol = sheet.getLastColumn();

    for (var c = 1; c <= lastCol; c++) {
        sheet.autoResizeColumn(c);
    }
}

// ==================== TÁBLÁZATOK - CÍMFORMÁZÁS ====================
if (sheet.getName() === "Táblázatok") {

    [
        "A1:B1",
        "C1:D1",
        "F2:G2",
        "H2:I2",
        "K1:L1",
        "M1:N1",
        "A30:B30",
        "C30:D30",
        "F30:G30",
        "H30:I30",
        "K30:L30",
        "M30:N30"
    ].forEach(function(rangeA1) {

        sheet.getRange(rangeA1)
            .setFontSize(10);

    });

    [
        "C1:D1",
        "H2:I2",
        "M1:N1",
        "C30:D30",
        "H30:I30",
        "M30:N30"
    ].forEach(function(rangeA1) {

        sheet.getRange(rangeA1)
            .setBackground("#ffffff")
            .setFontColor("#000000");

    });
}

SpreadsheetApp.flush();
}
// ====================== CHARTDATA ÉS DIAGRAM ======================

function updateChartData() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var salesSheet = ss.getSheetByName("Sales");
    var chartSheet = ss.getSheetByName("ChartData");

    if (!chartSheet) {
        chartSheet = ss.insertSheet("ChartData");
    }

    chartSheet.clear();

    // Dinamikus hónapok keresése
    const headerRow = 5;
    let col = 3;
    let monthColumns = [];

    while (true) {
        const header = salesSheet
            .getRange(headerRow, col)
            .getValue()
            .toString()
            .trim();

        if (header === "Total" || header === "" || col > 100) {
            break;
        }

        monthColumns.push({
            col: col,
            name: header
        });

        col++;
    }

    // ==========================================
    // STORE NEVEK
    // ==========================================
    const dataStartRow = 6;
    const numRows =23;

    for (let i = 0; i < numRows; i++) {
        const row = dataStartRow + i;

        chartSheet
            .getRange(i + 2, 1)
            .setValue(
                salesSheet.getRange(row, 2).getValue()
            );
    }

    // ==========================================
    // HÓNAP ADATOK
    // ==========================================
    for (let i = 0; i < numRows; i++) {

        const row = dataStartRow + i;

        for (let j = 0; j < monthColumns.length; j++) {

            const letter = columnToLetter(
                monthColumns[j].col
            );

            chartSheet
                .getRange(i + 2, j + 2)
                .setFormula(
                    `='Sales'!${letter}${row}`
                );
        }
    }

    createColumnChart(chartSheet, monthColumns.length);
}

// =====================================================
// FIX HÓNAPOK
// =====================================================
function getFixedMonths() {
    return [
        { name: "Január",     color: "#ff9aa2" },   // puha rózsaszín
        { name: "Február",    color: "#ff8c8c" },   // puha piros
        { name: "Március",    color: "#ffe38c" },   // puha sárga
        { name: "Április",    color: "#fff3a3" },   // világos sárga
        { name: "Május",      color: "#a3e4a3" },   // puha zöld
        { name: "Június",     color: "#9cd4ff" },   // puha kék
        { name: "Július",     color: "#c9a3e0" },   // puha lila
        { name: "Augusztus",  color: "#7ed6d1" },   // puha türkiz
        { name: "Szeptember", color: "#ffbb7d" },   // puha narancs
        { name: "Október",    color: "#8c9bb8" },   // szürke-kék
        { name: "November",   color: "#7ed6a8" },   // puha zöldes
        { name: "December",   color: "#ff9a9a" }    // puha piros
    ];
}

// =====================================================
// SERIES DEFINÍCIÓ
// =====================================================

function getSeriesWithLabels(numMonths) {

    const months = getFixedMonths();

    const series = {};

    for (let i = 0; i < Math.min(numMonths, 12); i++) {

        series[i] = {
            color: months[i].color,

            // Google Charts ezt használja,
            // ha támogatott az adott chart típusnál
            labelInLegend: months[i].name,

            label: months[i].name
        };
    }

    return series;
}

// =====================================================
// DIAGRAM
// =====================================================
// ====================== HU MONTHLY SUMMARY DIAGRAM (Offers vs Orders) ======================
function createHUMonthlyChart() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var dataSheet = ss.getSheetByName("HU monthly summary 2026");
    var chartSheet = ss.getSheetByName("ChartData");

    if (!dataSheet) {
        SpreadsheetApp.getUi().alert('Hiba: Nem található "HU monthly summary 2026" nevű munkalap!');
        return;
    }

    if (!chartSheet) {
        chartSheet = ss.insertSheet("ChartData");
    }

    var currentMonth = new Date().getMonth() + 1;
    var numStores = 10;

    var helperStartRow = 1;
    var helperStartCol = 12; // L oszlop

    chartSheet
        .getRange(helperStartRow, helperStartCol, numStores + 1, 5)
        .clearContent();

    chartSheet
        .getRange(helperStartRow, helperStartCol, 1, 5)
        .setValues([
            ["Store", "OFFERS", "OFFERS label", "ORDERS", "ORDERS label"]
        ]);

    var chartData = [];

    for (var i = 0; i < numStores; i++) {
        var storeName = dataSheet.getRange(3 + i, 1).getValue();

        var offersValues = dataSheet
            .getRange(3 + i, 2, 1, currentMonth)
            .getValues()[0];

        var ordersValues = dataSheet
            .getRange(17 + i, 2, 1, currentMonth)
            .getValues()[0];

        var offersSum = offersValues.reduce(function(sum, value) {
            return sum + (Number(value) || 0);
        }, 0);

        var ordersSum = ordersValues.reduce(function(sum, value) {
            return sum + (Number(value) || 0);
        }, 0);

        chartData.push([
            storeName,
            offersSum,
            formatEuroLabel(offersSum),
            ordersSum,
            formatEuroLabel(ordersSum)
        ]);
    }

    chartSheet
        .getRange(helperStartRow + 1, helperStartCol, chartData.length, 5)
        .setValues(chartData);

    var dataRange = chartSheet.getRange(
        helperStartRow,
        helperStartCol,
        numStores + 1,
        5
    );

    var chart = chartSheet.newChart()
        .setChartType(Charts.ChartType.COLUMN)
        .addRange(dataRange)
        .setNumHeaders(1)

        .setPosition(13, 12, 0, 0)

        .setOption('title', 'Offers vs Orders by Store')

        .setOption('backgroundColor', '#0B5394')
        .setOption('chartArea', {
            backgroundColor: '#0B5394',
            left: 80,
            top: 80
        })

        .setOption('hAxis', {
            title: 'Stores',
            textStyle: { color: '#ffffff', fontSize: 11 },
            titleTextStyle: { color: '#ffffff' },
            slantedText: true,
            slantedTextAngle: 45
        })

        .setOption('vAxis', {
            title: 'Amount',
            format: '#,##0€',
            textStyle: { color: '#ffffff' },
            titleTextStyle: { color: '#ffffff' },
            gridlines: { color: '#2f6fb0' },
            minorGridlines: { count: 2 },
            viewWindow: { min: 0 }
        })

        .setOption('legend', {
            position: 'top',
            alignment: 'center',
            textStyle: {
                color: '#ffffff',
                fontSize: 12
            }
        })

        .setOption('annotations', {
    alwaysOutside: true,
    stem: { color: 'none' },
    textStyle: {
        color: '#fff2cc',
        fontSize: 12,
        bold: true
    }
})
.setOption('displayAnnotations', true)

        .setOption('is3D', true)

        .setOption('series', {
            0: { color: '#fff2cc' }, // OFFERS - tojásos fehér
            1: { color: '#e25f00' }  // ORDERS - narancssárga
        })

        .setOption('height', 520)
        .setOption('width', 1050)
        .setOption('isStacked', false)
        .setOption('useFirstColumnAsDomain', true)

        .build();

    chartSheet.insertChart(chart);
}

// Euro formázás az oszlopok tetejére
function formatEuroLabel(value) {
    return Math.round(value)
        .toString()
        .replace(/\B(?=(\d{3})+(?!\d))/g, " ") + "€";
}
function createColumnChart(chartSheet, numMonths) {

    const existingCharts =
        chartSheet.getCharts();

    existingCharts.forEach(chart => {
        chartSheet.removeChart(chart);
    });

    // 2. sortól indulunk
    // nincs szükség fejlécre
    const range = chartSheet.getRange(
        2,
        1,
        23,
        numMonths + 1
    );

    const chart = chartSheet
        .newChart()
        .setChartType(
            Charts.ChartType.COLUMN
        )
        .addRange(range)

        .setPosition(
            2,
            1,
            0,
            0
        )

        .setOption(
            "title",
            "Turnover Sales by Store"
        )

        .setOption('backgroundColor', '#0B5394')

.setOption('chartArea', {
    backgroundColor: '#0B5394'
})

.setOption('legend', {
    position: 'top',
    alignment: 'center',
    textStyle: {
        color: '#ffffff',
        fontSize: 12
    }
})

.setOption('hAxis', {
    title: 'Stores',
    titleTextStyle: {
        color: '#ffffff'
    },
    textStyle: {
        color: '#ffffff'
    }
})

.setOption('vAxis', {
    title: 'Amount',
    titleTextStyle: {
        color: '#ffffff'
    },
    textStyle: {
        color: '#ffffff'
    },
    gridlines: {
        color: '#2f6fb0'
    }
})

        .setOption("hAxis", {
            title: "Stores"
        })

        .setOption("vAxis", {
            title: "Amount"
        })

        .setOption("legend", {
            position: "top",
            alignment: "center",
            textStyle: {
                fontSize: 12
            }
        })

        .setOption(
            "height",
            550
        )

        .setOption(
            "width",
            1050
        )

        .setOption(
            "isStacked",
            false
        )

        .setOption('is3D', true)

        .setOption(
            "series",
            getSeriesWithLabels(numMonths)
        )

        .setOption(
            "useFirstColumnAsDomain",
            true
        )

        .build();

    chartSheet.insertChart(chart);
}
// ==================== SEGÉDFÜGGVÉNYEK ====================
function getTotalColumn(sheet, startRow, startCol) {
    let col = startCol;
    while (true) {
        const value = sheet.getRange(startRow, col).getValue().toString().trim();
        if (value === "Total") break;
        col++;
        if (col > 200) break;
    }
    return columnToLetter(col);
}



function letterToColumn(letter) {
    var column = 0;

    for (var i = 0; i < letter.length; i++) {
        column = column * 26 + letter.charCodeAt(i) - 64;
    }

    return column;
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

function getPreviousMonthInfo() {
    const today = new Date();
    const currentMonth = today.getMonth() + 1;

    if (currentMonth === 1) {
        return {
            month: 0,
            yyyymm: null
        };
    }

    const previousMonth = currentMonth - 1;
    const year = today.getFullYear();
    const monthText = previousMonth < 10 ? "0" + previousMonth : "" + previousMonth;

    return {
        month: previousMonth,
        yyyymm: "" + year + monthText
    };
}

function getMonthColumnByHeader(sheet, headerRow, startCol, yyyymm, useLastMatch) {
    let foundCol = null;
    let col = startCol;

    while (col <= 200) {
        const value = sheet.getRange(headerRow, col).getValue().toString().trim();

        if (value === "" && foundCol !== null) {
            break;
        }

        if (value === "Total" && foundCol !== null && !useLastMatch) {
            break;
        }

        if (value === yyyymm) {
            foundCol = col;

            if (!useLastMatch) {
                break;
            }
        }

        col++;
    }

    return foundCol ? columnToLetter(foundCol) : null;
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
