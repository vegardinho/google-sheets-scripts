function onOpen() {
    hltCrMnth();
    ui.createMenu('KU Supermeny')
        .addItem('Sjekk ark alle arrangmentslenker', 'fiksAlleLenker')
        .addItem('Nytt arr på samme dato', 'delDato')
        .addItem('Gå til i dag', 'hltCrMnth')
        .addItem('Lag nytt hovedark', 'lagHippo')
        .addItem('Slå sammen arrangementer', 'mergeArrs')
        .addItem('Løs opp sammenslåing', 'løsOppMerge')
        .addToUi();
}


function hltCrMnth() {
    const TODAY = new Date();
    const THIS_MONTH = TODAY.getMonth();
    const THIS_DATE = TODAY.getDate();

    var thisSheet = hippo;
    var row = START_ROW;
    var hippoDay;
    var monthRange;
    var tekstDag;

    // var startDato = lagDato(START_ROW).split(".");
    // var sluttDato = lagDato(LAST_ROW).split(".");

    const START_DATE = thisSheet.getRange(START_DATE_CELL).getValue();
    const END_DATE = thisSheet.getRange(END_DATE_CELL).getValue();

    //Sjekk om vi er i en hippo
    if (START_DATE === "" || END_DATE === "") {
        return;
    }

    //Ikke kjør hvis utenfor hippokalenderen
    if (TODAY < new Date(START_DATE) || TODAY > new Date(END_DATE)) {
        return;
    }

    for (var i = 0; i < 6; i++) {
        monthRange = thisSheet.getRange(row, MONTH_COLUMN);
        if (!monthRange.isPartOfMerge()) {
            break;
        }
        var mnthCells = monthRange.getMergedRanges()[0];
        var måned = monthToNumber[mnthCells.getDisplayValue()];
        if (måned === THIS_MONTH + 1) {
            break;
        }
        row = mnthCells.getLastRow() + 1;
    }

    for (var i = row; i <= mnthCells.getLastRow(); i++) {
        hippoDay = thisSheet.getRange(i, DATE_COLUMN);

        if (hippoDay.isPartOfMerge()) {
            tekstDag = hippoDay.getMergedRanges()[0].getValue();
        } else {
            tekstDag = hippoDay.getValue();
        }

        if (tekstDag === THIS_DATE) {
            row = i;
            break;
        }
    }

    ss.setActiveSelection(thisSheet.getRange(row, DATE_COLUMN, 1, 7));
}  



function fiksAlleLenker() {  
    //TODO: for-loop hvor fiksLenke() kalles for alle celler
    for (var i = START_ROW; i <= ANT_RADER; i++) {
        fiksLenke(hippo.getRange(i, ARR_COLUMN));
    }        
    SpreadsheetApp.getUi().alert('Alle lenker er oppdatert!');
}


function slettAlleArk() {  
    var ui = SpreadsheetApp.getUi();
    var result = ui.alert(
        'Bekreft',
        'Er du sikker på at du vil slette alle arrangement-ark?',
        ui.ButtonSet.YES_NO);

    // Process the user's response.
    if (result == ui.Button.YES) {
        var sheets = ss.getSheets();

        for (var i = 3; i < sheets.length; i++) {
            if (sheets[i] != null) {
                ss.deleteSheet(sheets[i]);
            }
        }
    } else {
        ui.alert('Sletting avbrutt');
        return;
    }
    SpreadsheetApp.getUi().alert('Alle arrangement-ark slettet');
}


function delDato() {
    var radNr = ss.getActiveCell().getRow();
    hippo.insertRowAfter(radNr);

    //merge dag, dato og uke
    var dag = hippo.getRange(radNr, 4);
    var dato = hippo.getRange(radNr, 3);
    var ukeStartRad = hippo.getRange(radNr, 2).getMergedRanges()[0].getRowIndex();
    var ukeSluttRad = hippo.getRange(radNr, 2).getMergedRanges()[0].getLastRow();

    if (dag.isPartOfMerge()) {

        var startRadNr = dag.getMergedRanges()[0].getRowIndex();
        var sluttRadNr = dag.getMergedRanges()[0].getLastRow();
        var antRad = (sluttRadNr - startRadNr) + 2;

        if (dag.isPartOfMerge() && (radNr === sluttRadNr)) {
            hippo.getRange(startRadNr, 3, antRad).mergeVertically();
            hippo.getRange(startRadNr, 4, antRad).mergeVertically(); 
        } 
    } else {
        hippo.getRange(radNr, 4, 2).mergeVertically();
        hippo.getRange(radNr, 3, 2).mergeVertically();
    }

    if (ukeSluttRad === radNr) {
        var antUkeRad = (ukeSluttRad - ukeStartRad) + 2;
        hippo.getRange(ukeStartRad, 2, antUkeRad).mergeVertically();
    }
}


function lagHippo() {
    var thisSheet = ss.insertSheet();
    //   var startDato = new Date();
    //   var sluttDato = new Date(2020, 5, 24);

    thisSheet.clear();

    var responseStart = ui.prompt('Velg en startdato (åååå-mm-dd)');
    if (responseStart.getSelectedButton() != ui.Button.OK) {
        return;
    }

    var responseEnd = ui.prompt('Velg en sluttdato (åååå-mm-dd)');
    if (responseEnd.getSelectedButton() != ui.Button.OK) {
        return;
    }

    var startStreng = responseStart.getResponseText();
    var sluttStreng = responseEnd.getResponseText()
    var startDato = new Date(startStreng);
    var sluttDato = new Date(sluttStreng);

    // Lagre datostrenger for senere
    kopi.getRange(START_DATE_ROW, START_DATE_COL).setValue(startStreng);
    kopi.getRange(END_DATE_ROW, END_DATE_COL).setValue(sluttStreng);

    var denneDatoen = startDato;
    var thisMonth, thisDate, thisDay, thisWeek;
    var lastMonth;
    var monthCell;
    var weekRange, weekdayRange;
    var farge;
    var monthRow = START_ROW;
    var row = START_ROW;
    var lastMonth = 12;
    var semester;

    var columnNames = ["Mnd", "Uke", "Dato", "Dag", "Arrangement", "Lokale", "Ansvarlig", "På jobb", "Kommentar"];
    for (var i = 0; i < columnNames.length; i++) {
        thisSheet.getRange(START_ROW - 1, (i + 1)).setValue(columnNames[i]).setHorizontalAlignment("center");
    }

    if (startDato.getMonth() < 6) {
        semester = "Vår";
    } else {
        semester = "Høst";
    }

    var streng = Utilities.formatString("%s %s Tentativ Hippo", semester, startDato.getFullYear());
    thisSheet.getRange(1, 1)
        .setValue(streng)
        .setFontSize(24)
        .setFontWeight("bold");
    thisSheet.getRange(1, 1, (START_ROW - 2), NUM_COLS).merge();
    thisSheet.setFrozenRows(START_ROW - 1);

    thisSheet.getRange(START_ROW, WEEK_COLUMN).setValue(denneDatoen.getWeek());

    while (denneDatoen <= sluttDato) {
        thisMonth = denneDatoen.getMonth();
        thisWeek = denneDatoen.getWeek();
        thisDate = denneDatoen.getDate();
        thisDay = denneDatoen.getDay();

        farge = farger[thisMonth % ANT_FARGE_GRP][thisWeek % ANT_FARGE_WEEK];

        thisSheet.getRange(row, DATE_COLUMN, 1, (NUM_COLS - DATE_COLUMN + 1)).setBackground(farge);
        thisSheet.getRange(row, DATE_COLUMN).setValue(thisDate);
        thisSheet.getRange(row, DAY_COLUMN).setValue(ukedager[thisDay]);

        if (thisDay != THURSDAY) {
            denneDatoen.setDate(thisDate + 1);
        }

        switch (thisDay) {
            case MONDAY:
                thisSheet.getRange(row, WEEK_COLUMN).setValue(thisWeek);
                break;
            case THURSDAY:
                denneDatoen.setDate(thisDate + 3);
                break;
            case SUNDAY:
                if ((row - START_ROW) < DAYS_WEEK) {
                    weekRange = thisSheet.getRange(START_ROW, WEEK_COLUMN, (row - START_ROW + 1), 1);
                } else {
                    weekRange = thisSheet.getRange(row - 4, WEEK_COLUMN, DAYS_WEEK, 1);
                }
                handleWeek(weekRange);
        }

        if (thisMonth != lastMonth) { 
            setMonthContent(thisSheet, row, thisMonth);
            mergeMonth(row, monthRow, thisMonth, lastMonth, thisSheet);
            monthRow = row;
        }

        lastMonth = thisMonth;
        row++;
    }

    mergeMonth(row, monthRow, thisMonth, lastMonth, thisSheet);

    if (thisDay > 0) {
        handleWeek(thisSheet.getRange(row - thisDay, WEEK_COLUMN, thisDay, 1));
    }

    thisSheet.getRange(START_ROW, MONTH_COLUMN, (row - START_ROW), NUM_COLS).setBorder(true, true, true, true, true, true);
    thisSheet.getRange(START_ROW - 1, MONTH_COLUMN, 1, NUM_COLS).setBorder(true, true, false, true, true, true);

    //Lagre siste rad
    kopi.getRange(LAST_ROW_INFO_ROW, LAST_ROW_INFO_COL).setValue(row - 1);
}

function handleWeek(weekRange) {
    weekRange
        .mergeVertically()
        .setBackground("#efefef")
        .setFontWeight("bold")
        .setHorizontalAlignment("center")
        .setVerticalAlignment("middle");
}

function setMonthContent(thisSheet, row, thisMonth) {
    thisSheet.getRange(row, MONTH_COLUMN)
        .setValue(indexToMonth[thisMonth])
        .setFontSize(30)
        .setVerticalText(true)
        .setVerticalAlignment("middle");
}

function mergeMonth(row, monthRow, thisMonth, lastMonth, thisSheet) {
    if (monthRow === row) {
        return;
    }

    var monthRange = thisSheet.getRange(monthRow, MONTH_COLUMN, (row - monthRow), 1);
    var farge = farger[lastMonth % ANT_FARGE_GRP][2];
    monthRange.mergeVertically();
    monthRange.setBackground(farge);
}


function mergeArrs() {
  
  if (ss.getActiveSheet() !== hippo) {
      return;
  }

    response = ui.alert("Bekreft", "Er du sikker på at du vil slå sammen? Kun verdier på den øverste raden vil lagres etter sammenslåingen.", ui.ButtonSet.YES_NO_CANCEL);
    if (response != ui.Button.YES) {
        return;
    }

    var cells = ss.getActiveRange()
    var startRow = cells.getRow();
    var numRows = cells.getNumRows();
    var NUM_COLS = 3;

    kopi.getRange(startRow + 1, ARR_COLUMN, numRows - 1, 1).clearContent();
    hippo.getRange(startRow, ARR_COLUMN, numRows, NUM_COLS).mergeVertically();
}

// Break apart selected cell and reset formatting. Check also if date column cell on row is merged; delete excess row if true.
function løsOppMerge(celle) {
    var thisCell = ss.getActiveCell(); 
    var firstRow = thisCell.getRowIndex();
    var dayCell = hippo.getRange(firstRow, DAY_COLUMN);

    //Return if either wrong sheet or no expected cells merged
    if ((ss.getActiveSheet() !== hippo) || (!dayCell.isPartOfMerge() && !thisCell.isPartOfMerge())) {
        return;
    }
  
    if (dayCell.isPartOfMerge()) {
      ss.deleteRow(firstRow);
    } 
  
    if (thisCell.isPartOfMerge()) {
      var lastRow = thisCell.getMergedRanges()[0].getLastRow();
      var numRows = lastRow - firstRow + 1;
      var numCols = 4;
      var mergeRange = hippo.getRange(firstRow, ARR_COLUMN, numRows, numCols);
      mergeRange.breakApart();
      mergeRange.setBorder(true, true, true, true, true, true);
    }
}