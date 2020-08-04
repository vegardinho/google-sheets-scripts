function onOpen() {
    hltCrMnth();
    ui.createMenu('KU Supermeny')
        .addItem('Sjekk ark alle arrangmentslenker', 'checkAllLinks')
        .addItem('Nytt arr på samme dato', 'addDateRow')
        .addItem('Gå til i dag', 'hltCrMnth')
        .addItem('Lag nytt hovedark', 'crtMainSht')
        .addItem('Slå sammen arrangementer', 'mergeEvs')
        .addItem('Løs opp sammenslåing', 'undoMerge')
        .addToUi();
}


// Evaluate calendar and current date and highlight current date if in scope
function hltCrMnth() {
    // Date-library months from 0-11
    const TODAY = new Date();
    const THIS_MONTH = TODAY.getMonth() + 1; 
    const THIS_DATE = TODAY.getDate();

    var row = M_START_ROW;
    var hippoDay;
    var monthRange;
    var dayText;

    // Test cases:
    // var startDate = lagDato(M_START_ROW).split(".");
    // var endDate = lagDato(LAST_ROW).split(".");

    if (ss.getActiveSheet().getName() !== hippo.getName()) {
        return;
    }

    // Don't run if date outside calendar defined in main sheet
    if (TODAY < new Date(M_START_DATE) || TODAY > new Date(M_END_DATE)) {
        return;
    }
 

    // Jump from month to month until either right month or not in merged cell
    // TODO: use while-loop instead; avoids setting maximum iterations
    for (var i = 0; i < 12; i++) {
        monthRange = hippo.getRange(row, M_MNTH_COL_NUM);
        if (!monthRange.isPartOfMerge()) {
            break;
        }
        var topMonthCell = monthRange.getMergedRanges()[0];
        var month = monthToNumber[topMonthCell.getDisplayValue()];
        if (month === THIS_MONTH) {
            break;
        }
        row = topMonthCell.getLastRow() + 1;
    }

    // Loop until day-cell matches today's date
    for (var i = row; i <= topMonthCell.getLastRow(); i++) {
        hippoDay = hippo.getRange(i, M_DATE_COL_NUM);

        if (hippoDay.isPartOfMerge()) {
            dayText = hippoDay.getMergedRanges()[0].getValue();
        } else {
            dayText = hippoDay.getValue();
        }

        if (dayText === THIS_DATE) {
            row = i;
            break;
        }
    }

    ss.setActiveSelection(hippo.getRange(row, M_DATE_COL_NUM, 1, 7));
}  


// Run fiksLenke on every event-cell in main sheet
function checkAllLinks() {  
    //TODO: for-loop hvor fiksLenke() kalles for alle celler
    for (var i = M_START_ROW; i <= M_NUM_ROWS; i++) {
        setEventLink(hippo.getRange(i, M_EVENT_COL_NUM));
    }        
    ui.alert('Alle lenker er oppdatert!');
}


// Delete all sheets, except first M_NUM_SH_DFLT sheets
function dltAllSheets() {  
    var result = ui.alert(
            'Bekreft',
            'Er du sikker på at du vil slette alle arrangement-ark?',
            ui.ButtonSet.YES_NO);

    if (result !== ui.Button.YES) {
        ui.alert('Sletting avbrutt');
        return;
    }

    var sheets = ss.getSheets();

    for (var i = M_NUM_SH_DFLT; i < sheets.length; i++) {
        if (sheets[i] != null) {
            ss.deleteSheet(sheets[i]);
        }
    }
    ui.alert('Alle arrangement-ark slettet');
}


// Add additional row beneath selected cell, to allow for more events
// on one date. Merge appropriate cells after insertion.
function addDateRow() {
    if (ss.getActiveSheet().getName() != hippo.getName()) {
        return;
    }

    var firstRow = ss.getActiveCell().getRow();
    hippo.insertRowAfter(firstRow);
    removeFormatting(hippo.getRange(firstRow + 1, M_EVENT_COL_NUM), '');

    // Merge rows for columns: day, dateNum, weekday and comments
    var day = hippo.getRange(firstRow, M_DAY_COL_NUM);
    var weekStartRow = hippo.getRange(firstRow, M_WEEK_COL_NUM).getMergedRanges()[0].getRowIndex();
    var weekEndRow = hippo.getRange(firstRow, M_WEEK_COL_NUM).getMergedRanges()[0].getLastRow();
    var numRows = 2;

    if (day.isPartOfMerge() && (firstRow === lastRow)) {

        firstRow = day.getMergedRanges()[0].getRowIndex();
        var lastRow = day.getMergedRanges()[0].getLastRow();
        numRows = (lastRow - firstRow) + 2;
    }

    hippo.getRange(firstRow, M_DAY_COL_NUM, numRows).mergeVertically();
    hippo.getRange(firstRow, M_DATE_COL_NUM, numRows).mergeVertically();
    hippo.getRange(firstRow, M_COMM_COL_NUM, numRows).mergeVertically();

    if (weekEndRow === firstRow) {
        var antUkeRad = (weekEndRow - weekStartRow) + 2;
        hippo.getRange(weekStartRow, 2, antUkeRad).mergeVertically();
    }
}

// Create main sheet (hippo) from user input dates
function crtMainSht() {
    var thisSheet = ss.insertSheet();
    thisSheet.clear();

    var responseStart = ui.prompt('Velg en startDate (åååå-mm-dd)');
    if (responseStart.getSelectedButton() != ui.Button.OK) {
        return;
    }

    var responseEnd = ui.prompt('Velg en endDate (åååå-mm-dd)');
    if (responseEnd.getSelectedButton() != ui.Button.OK) {
        return;
    }

    var startString = responseStart.getResponseText();
    var endString = responseEnd.getResponseText();
    var startDate = new Date(startString);
    var endDate = new Date(endString);
    var lastRow; 

    // Lagre datostrenger for senere
    copy.getRange(START_DATE_ROW, START_DATE_COL).setValue(startString);
    copy.getRange(END_DATE_ROW, END_DATE_COL).setValue(endString);

    // Test case:
    // var startDate = new Date();
    // var endDate = new Date(2020, 5, 24);

    setTitle(startDate, thisSheet);
    setColTitles();
    lastRow = createCal(thisSheet, startDate, endDate);
    setStyle();

    // Save last row number for later use
    copy.getRange(LAST_ROW_INFO_ROW, LAST_ROW_INFO_COL).setValue(lastRow - 1);
}

// Define column titles
function setColTitles() {
    var columnNames = ["Mnd", "Uke", "Dato", "Dag", "Arrangement", "Lokale", "Ansvarlig", "På jobb", "Kommentar"];
    for (var i = 0; i < columnNames.length; i++) {
        thisSheet.getRange(M_START_ROW - 1, (i + 1)).setValue(columnNames[i]).setHorizontalAlignment("center");
    }
}

// Set border style and vertucal alignment on calendar cells
function setStyle(sheet, rw) {
    var thisSheet = sheet;
    var row = rw;

    thisSheet.getRange(M_START_ROW, M_MNTH_COL_NUM, (row - M_START_ROW), M_NUM_COLS).setBorder(true, true, true, true, true, true);
    thisSheet.getRange(M_START_ROW - 1, M_MNTH_COL_NUM, 1, M_NUM_COLS).setBorder(true, true, false, true, true, true);
    
    thisSheet.getRange(M_START_ROW, M_MNTH_COL_NUM, (row - M_START_ROW), M_NUM_COLS).setVerticalAlignment("middle");
}

// Define content and colors in cells of calendar
function createCal(strtDt, endDt, sheet) {
    var thisSheet = sheet;
    var endDate = endDt;
    var startDate = strtDt;

    var color;
    var weekRange, weekdayRange;
    var monthRow = M_START_ROW;
    var row = M_START_ROW;
    var thisMonth, thisDateNum, thisDay, thisWeek
        var thisDate = startDate;
    // Ensure prevMonth != thisMonth
    var prevMonth = thisDate.getMonth() - 1;

    thisSheet.getRange(M_START_ROW, M_WEEK_COL_NUM).setValue(thisDate.getWeek());

    while (thisDate <= endDate) {
        thisMonth = thisDate.getMonth();
        thisWeek = thisDate.getWeek();
        thisDateNum = thisDate.getDate();
        thisDay = thisDate.getDay();

        // Automatically select alternating colors, based on month and week
        color = colors[thisMonth % NUM_CLRS_GRP][thisWeek % NUM_CLRS_WK];

        thisSheet.getRange(row, M_DATE_COL_NUM, 1, (M_NUM_COLS - M_DATE_COL_NUM + 1)).setBackground(color);
        thisSheet.getRange(row, M_DATE_COL_NUM).setValue(thisDateNum);
        thisSheet.getRange(row, M_DAY_COL_NUM).setValue(weekdays[thisDay]);

        // Set weeknumber on corresponding column if monday; skip to sunday on next iteartion, if thursday;
        // merge and set style for week-column, if sunday
        switch (thisDay) {
            case MONDAY:
                thisSheet.getRange(row, M_WEEK_COL_NUM).setValue(thisWeek);
                break;
            case THURSDAY:
                thisDate.setDate(thisDateNum + 3);
                break;
            case SUNDAY:
                if ((row - M_START_ROW) < DAYS_PER_WEEK) {
                    weekRange = thisSheet.getRange(M_START_ROW, M_WEEK_COL_NUM, (row - M_START_ROW + 1), 1);
                } else {
                    weekRange = thisSheet.getRange(row - 4, M_WEEK_COL_NUM, DAYS_PER_WEEK, 1);
                }
                handleWeek(weekRange);
        }

        // Iterate to next day unless thursday
        if (thisDay != THURSDAY) {
            thisDate.setDate(thisDateNum + 1);
        }

        // If new month, merge and style old month-column
        if (thisMonth != prevMonth) { 
            setMonthContent(thisSheet, row, thisMonth);
            mergeMonth(row, monthRow, prevMonth);
            monthRow = row;
        }

        prevMonth = thisMonth;
        row++;
    }

    // Handle remaining, uncompleted week and month
    mergeMonth(row, monthRow, prevMonth);
    if (thisDay > 0) {
        handleWeek(thisSheet.getRange(row - thisDay, M_WEEK_COL_NUM, thisDay, 1));
    }

    return row;
}

// Set main title in main sheet
function setTitle(strtDt, sheet) {
    var startDate = strtDt;
    var thisSheet = sheet;
    var semester;

    if (startDate.getMonth() < 6) {
        semester = "Vår";
    } else {
        semester = "Høst";
    }

    var streng = Utilities.formatString("%s %s Tentativ Hippo", semester, startDate.getFullYear());
    thisSheet.getRange(1, 1)
        .setValue(streng)
        .setFontSize(24)
        .setFontWeight("bold");
    thisSheet.getRange(1, 1, (M_START_ROW - 2), M_NUM_COLS).merge();
    thisSheet.setFrozenRows(M_START_ROW - 1);
}

function handleWeek(weekRange) {
    weekRange
        .mergeVertically()
        .setBackground("#efefef")
        .setFontWeight("bold")
        .setHorizontalAlignment("center")
}

function setMonthContent(thisSheet, row, thisMonth) {
    thisSheet.getRange(row, M_MNTH_COL_NUM)
        .setValue(indexToMonth[thisMonth])
        .setFontSize(30)
        .setVerticalText(true)
}

function mergeMonth(row, monthRow, prevMonth) {
    if (monthRow === row) {
        return;
    }

    var monthRange = thisSheet.getRange(monthRow, M_MNTH_COL_NUM, (row - monthRow), 1);
    var color = colors[prevMonth % NUM_CLRS_GRP][2];
    monthRange.mergeVertically();
    monthRange.setBackground(color);
}


// Merge rows for event that continues for several days
function mergeEvs() {
    if (ss.getActiveSheet().getName() !== hippo.getName()) {
        return;
    }

    var cells = ss.getActiveRange();
    var startRow = cells.getRow();
    var numRows = cells.getNumRows();
    const NUM_COLS = 3;

    if (numRows <= 1) {
        ui.alert("Feilmelding", "Vennligst velg minst to rader som skal slås sammen", ui.ButtonSet.OK);
        return;
    }


    response = ui.alert("Bekreft", "Er du sikker på at du vil slå sammen? Kun verdier på den øverste raden vil lagres etter sammenslåingen.", ui.ButtonSet.YES_NO_CANCEL);
    if (response != ui.Button.YES) {
        return;
    }

    copy.getRange(startRow + 1, M_EVENT_COL_NUM, numRows - 1, 1).clearContent();
    hippo.getRange(startRow, M_EVENT_COL_NUM, numRows, NUM_COLS).mergeVertically();
}

// Break apart selected cell and reset formatting. Check also if date column cell on row is merged; delete excess row if true.
function undoMerge(celle) {
    var thisCell = ss.getActiveCell(); 
    var firstRow = thisCell.getRowIndex();
    var dayCell = hippo.getRange(firstRow, M_DAY_COL_NUM);

    //Return if either wrong sheet or no expected cells merged
    if ((ss.getActiveSheet().getName() !== hippo.getName()) || (!dayCell.isPartOfMerge() && !thisCell.isPartOfMerge())) {
        return;
    }

    if (dayCell.isPartOfMerge()) {
        ss.deleteRow(dayCell.getMergedRanges()[0].getLastRow());
    } 

    if (thisCell.isPartOfMerge()) {
        var lastRow = thisCell.getMergedRanges()[0].getLastRow();
        var numRows = lastRow - firstRow + 1;
        var numCols = 4;
        var mergeRange = hippo.getRange(firstRow, M_EVENT_COL_NUM, numRows, numCols);
        mergeRange.breakApart();
        mergeRange.setBorder(true, true, true, true, true, true);
    }
}