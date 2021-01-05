function onOpen() {
        ui.createMenu('KU Supermeny')
        .addItem('Sjekk alle arrangmentslenker', 'checkAllLinks')
    .addItem('Oppdater PR/Hotell', 'updateSheet')
    //.addItem('Fjern arr på samme dato', 'rmvDateRow')
    //.addItem('Legg til ny rad', 'addDateRow')
    .addItem('Gå til i dag', 'hltCrMnth')
    .addItem('Lag nytt hovedark', 'crtMainSht')
    //.addItem('Løs opp sammenslåing', 'undoMerge')
    .addToUi();
}

/**
* Evaluate calendar and current date and highlight current date if in scope
* 
* @return {undefined} 
*/
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
            break;
        }
    }
    row = i;

    ss.setActiveSelection(hippo.getRange(row, M_DATE_COL_NUM, 1, 7));
}  


/**
* Run fiksLenke on every event-cell in main sheet
* 
* @return {undefined} 
*/
function checkAllLinks() {  
    for (var i = M_START_ROW; i <= mNumRows; i++) {
        var loopCell = hippo.getRange(i, M_EVENT_COL_NUM);
        if (loopCell.isPartOfMerge()) {
            i = loopCell.getLastRow();
        }
        setEventLink(loopCell);
    }
    ui.alert('Alle lenker er oppdatert!');
}


/**
* Delete all sheets, except first M_NUM_SH_DFLT sheets
* 
* @return {undefined} 
*/
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

/**
* Removes rows belonging to event of active cell 
* 
* @return {undefined} 
*/
function rmvDateRow() {
    if (ss.getActiveSheet().getName() != hippo.getName()) {
        return;
    }

    var cell = ss.getActiveCell();
    var cellRow = cell.getRow();
    var day = hippo.getRange(cellRow, M_DAY_COL_NUM);
    var dayIndex = day.getRowIndex();
    var dayContent;
    var dateNum = hippo.getRange(cellRow, M_DATE_COL_NUM);
    var dateNumContent;
    var numRows = 1;
    var firstRow = cellRow;
    var lastRow = cellRow;


    dayContent = day.getDisplayValue();
    dateNumContent = dateNum.getDisplayValue();

    if (cell.isPartOfMerge()) {
        firstRow = cell.getMergedRanges()[0].getRowIndex();
        lastRow = cell.getMergedRanges()[0].getLastRow();
        numRows = (lastRow - firstRow) + 1;
        //Don't delete if event is all columns. Correct action is to do 'løs opp sammenslåing'
        if (day.getLastRow() === lastRow && day.getRowIndex() === firstRow) {
            return;
        }
    }  

    hippo.deleteRows(firstRow, numRows);

    if (day.isPartOfMerge()) {
        day = day.getMergedRanges()[0];
        dateNum = dateNum.getMergedRanges()[0];
    }

    day.setValue(dayContent);
    dateNum.setValue(dateNumContent);

}

/**
* Add additional row beneath selected cell, to allow for more events 
* on one date. Merge appropriate cells after insertion.
*
* @return {undefined} 
*/
function addDateRow(sheet=null, row=null) {
    var thisSheet;
    if (sheet == null) {
        if (ss.getActiveSheet().getName() != hippo.getName()) {
            return;
        }
        thisSheet = hippo;  
    } else {
        thisSheet = sheet;
    }



    var cellRow;
    if (row == null) {
        cellRow = ss.getActiveCell().getRow();
    } else {
        cellRow = row;
    }

    var firstRow = cellRow;
    var lastRow;
    var numRows = 2;
    var day = thisSheet.getRange(cellRow, M_DAY_COL_NUM);

    var weekCell = thisSheet.getRange(cellRow, M_WEEK_COL_NUM);
    var weekStartRow = cellRow;
    var weekEndRow = cellRow;

    var mnthCell = thisSheet.getRange(cellRow, M_MNTH_COL_NUM);
    var mnthStartRow = cellRow;
    var mnthEndRow = cellRow;

    if (day.isPartOfMerge()) {
        day = day.getMergedRanges()[0];
        firstRow = day.getMergedRanges()[0].getRowIndex();
        lastRow = day.getLastRow();
        numRows = (lastRow - firstRow) + 2;
    }

    if (weekCell.isPartOfMerge()) {
        weekStartRow = weekCell.getMergedRanges()[0].getRowIndex();
        weekEndRow = weekCell.getMergedRanges()[0].getLastRow();
    }

    if (mnthCell.isPartOfMerge()) {
        mnthStartRow = mnthCell.getMergedRanges()[0].getRowIndex();
        mnthEndRow = mnthCell.getMergedRanges()[0].getLastRow();
    }

    lastRow = day.getLastRow();

    thisSheet.insertRowAfter(cellRow);
    resetCellValue(thisSheet.getRange(cellRow + 1, M_EVENT_COL_NUM), '');

    thisSheet.getRange(firstRow, M_DAY_COL_NUM, numRows).mergeVertically();
    thisSheet.getRange(firstRow, M_DATE_COL_NUM, numRows).mergeVertically();
    thisSheet.getRange(firstRow, M_COMM_COL_NUM, numRows).mergeVertically();

    if (cellRow === weekEndRow) {
        numRows = (weekEndRow - weekStartRow) + 2;
        thisSheet.getRange(weekStartRow, M_WEEK_COL_NUM, numRows).mergeVertically();
    }

    if (cellRow === mnthEndRow) {
        numRows = (mnthEndRow - mnthStartRow) + 2;
        thisSheet.getRange(mnthStartRow, M_MNTH_COL_NUM, numRows).mergeVertically();
    }

    thisSheet.getRange(cellRow + 1, M_EVENT_COL_NUM).setBorder(true, true, true, true, true, true);
    mNumRows++;
    copy.getRange(CP_LAST_ROW_CELL).setValue(mNumRows);
}



/**
* Create main sheet (hippo) from user input dates
* @return {undefined} 
*/
function crtMainSht() {
    var thisSheet = ss.insertSheet();
    thisSheet.clear();

    var responseStart = ui.prompt('Velg en startdato (åååå-mm-dd)');
    if (responseStart.getSelectedButton() != ui.Button.OK) {
        return;
    }

    var responseEnd = ui.prompt('Velg en sluttdato (åååå-mm-dd)');
    if (responseEnd.getSelectedButton() != ui.Button.OK) {
        return;
    }

    var startString = responseStart.getResponseText();
    var endString = responseEnd.getResponseText();
    var startDate = new Date(startString);
    var endDate = new Date(endString);
    var lastRow; 

    // Lagre datostrenger for senere
    copy.getRange(CP_STRT_DATE_CLL_MAIN).setValue(startString);
    copy.getRange(CP_END_DATE_CLL_MAIN).setValue(endString);

    // Test case:
    // var startDate = new Date();
    // var endDate = new Date(2020, 5, 24);

    setTitle(startDate, thisSheet);
    setColTitles(thisSheet);
    lastRow = createCal(thisSheet, startDate, endDate);

    // Save last row number for later use
    copy.getRange(CP_LAST_ROW_CELL).setValue(lastRow - 1);

    for (let i = M_START_ROW; i < (lastRow * 2) - (M_START_ROW); i+=2) {
        addDateRow(thisSheet, i);
    }

    setStyle(thisSheet, (lastRow * 2) - (M_START_ROW));
}


/**
* Define column titles
*
* @return {undefined} 
*/
function setColTitles(thisSheet) {
    var columnNames = ["Mnd", "Uke", "Dato", "Dag", "Arrangement", "Lokale", "Ansvarlig", "På jobb", "Kommentar"];
    for (var i = 0; i < columnNames.length; i++) {
        thisSheet.getRange(M_START_ROW - 1, (i + 1)).setValue(columnNames[i]).setHorizontalAlignment("center");
    }
}

/**
* Set border style and vertucal alignment on calendar cells
* 
* @param {Sheet}   sheet  
* @param {Integer} endRow 
*/
function setStyle(sheet, endRow) {
    var thisSheet = sheet;
    var row = endRow;

    thisSheet.getRange(M_START_ROW, M_MNTH_COL_NUM, (row - M_START_ROW), M_NUM_COLS).setBorder(true, true, true, true, true, true);
    thisSheet.getRange(M_START_ROW - 1, M_MNTH_COL_NUM, 1, M_NUM_COLS).setBorder(true, true, false, true, true, true);

    thisSheet.getRange(M_START_ROW, M_MNTH_COL_NUM, (row - M_START_ROW), M_NUM_COLS).setVerticalAlignment("middle");
}

/**
* Define content and colors in cells of calendar
* @param  {String} strtDt 
* @param  {String} endDt 
* @param  {Sheet}  sheet
* @return {Integer}        Last row of calendar
*/
function createCal(sheet, strtDt, endDt) {
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
        mergeMonth(thisSheet, row, monthRow, prevMonth);
        monthRow = row;
    }

    prevMonth = thisMonth;
    row++;
    }

    // Handle remaining, uncompleted week and month
    mergeMonth(thisSheet, row, monthRow, prevMonth);
    if (thisDay > 0) {
        handleWeek(thisSheet.getRange(row - thisDay, M_WEEK_COL_NUM, thisDay, 1));
    }

    return row;
}

/**
* Set main title in main sheet
* 
* @param {String}  strtDt 
* @param {Sheet}   sheet 
*/
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


/**
* Sets style and merges range.
* 
* @param  {Range}  weekRange 
* @return {undefined}           
*/
function handleWeek(weekRange) {
    weekRange
    .mergeVertically()
    .setBackground("#efefef")
    .setFontWeight("bold")
    .setHorizontalAlignment("center")
}

/**
* Formats range according to month specifications.
* 
* @param {Sheet} thisSheet 
* @param {Integer} row      
* @param {Integer} thisMonth 
*/
function setMonthContent(thisSheet, row, thisMonth) {
    thisSheet.getRange(row, M_MNTH_COL_NUM)
    .setValue(indexToMonth[thisMonth])
    .setFontSize(30)
    .setVerticalText(true)
}


/**
* Merges cells, and sets format.
* 
* @param  {Integer} row       
* @param  {Integer} monthRow  
* @param  {Integer} prevMonth
* @return {undefined}         
*/
function mergeMonth(thisSheet, row, monthRow, prevMonth) {
    if (monthRow === row) {
        return;
    }

    var monthRange = thisSheet.getRange(monthRow, M_MNTH_COL_NUM, (row - monthRow), 1);
    var color = colors[prevMonth % NUM_CLRS_GRP][2];
    monthRange.mergeVertically();
    monthRange.setBackground(color);
}


/**
* Merges rows for event that continues for several days
* 
* @return {undefined} 
*/
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


/**
* Break apart selected cell and reset formatting. 
* Check also if date column cell on row is merged; delete excess row if true.
* 
* @return {undefined}       
*/
function undoMerge() {
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


/**
 * Checks wether sheet is hotel or PR, then calls function to update list in sheet
 * 
 * @return {undefined} 
 */
function updateSheet() {
    var thisSheet = ss.getActiveSheet();

    if (thisSheet.getName() == pr.getName()) {
        SharedFunctions.checkIfNewEvents(addEventToPrHotel, copy, thisSheet, PR_START_ROW, PR_LAST_ROW, PR_DATE_COL_NUM,
            PR_LAST_COL_NUM, PR_EVENT_COL_NUM, PR_DATE_COL_NUM, CP_START_ROW, CP_LAST_ROW, CP_EVENT_COL_NUM, CP_DATE_COL_NUM, 4); 
    } else if (thisSheet.getName() == hotel.getName()) {
        SharedFunctions.checkIfNewEvents(addEventToPrHotel, copy, thisSheet, H_START_ROW, H_LAST_ROW, H_DATE_COL_NUM,
            H_LAST_COL_NUM, H_EVENT_COL_NUM, H_DATE_COL_NUM, CP_START_ROW, CP_LAST_ROW, CP_EVENT_COL_NUM, CP_DATE_COL_NUM, 4); 
    } else {
        ui.alert("Vennligst velg PR- eller Hotell-arket før du forsøker å oppdatere.");
    }

    fixPRHotelCells();
}


/**
 * Sets or links values in either hotel- or PR sheet, from event.
 * @param {Sheet} dstSheet  Hotel or PR sheet   
 * @param {Integer} sheetRow    Next empty row in @dstSheet to put event values
 * @param {Integer} copyRow     Row in copy sheet with event
 */
function addEventToPrHotel(sheetRow, copyRow, dstSheet) {
    var evName = copy.getRange(copyRow, CP_EVENT_COL_NUM).getDisplayValue();
    var copyCell = copy.getRange(copyRow, CP_EVENT_COL_NUM);
    var evSheet = ss.getSheetByName(evName);
    var evDate = copy.getRange(copyRow, CP_DATE_COL_NUM).getValue();
    var dstEventColNum, dstDateColNum, evFormula;

    if (evSheet == null) {
        evSheet = getEvSheet(evName);
        ss.setActiveSheet(dstSheet);
        setUpNewSheet(evSheet, evName, hippo.getRange(copyRow, M_EVENT_COL_NUM), evDate);
    }

    ss.setActiveSheet(dstSheet);
    evFormula = copyCell.getFormula();

    if (dstSheet.getName() == pr.getName()) {
        dstEventColNum = PR_EVENT_COL_NUM;
        dstDateColNum = PR_DATE_COL_NUM;

        dst = [dstSheet.getRange(sheetRow, PR_RLS_COL_NUM), dstSheet.getRange(sheetRow, PR_TCKTS_COL_NUM), 
        dstSheet.getRange(sheetRow, PR_GRFCS_COL_NUM), dstSheet.getRange(sheetRow, PR_PHOTOS_COL_NUM)];
        src = [evSheet.getRange(TMPLT_RLS), evSheet.getRange(TMPLT_RLS_TCKTS), evSheet.getRange(TMPLT_GRFCS), 
        evSheet.getRange(TMPLT_PHOTOS)];
    } else {
        dstEventColNum = H_EVENT_COL_NUM;
        dstDateColNum = H_DATE_COL_NUM;

        dst = [dstSheet.getRange(sheetRow, H_STATUS_COL_NUM), dstSheet.getRange(sheetRow, H_GUESTS_COL_NUM)];
        src = [evSheet.getRange(TMPLT_HOTEL), evSheet.getRange(TMPLT_GUESTS)];
    }

    dstSheet.getRange(sheetRow, dstEventColNum).setFormula(evFormula);
    dstSheet.getRange(sheetRow, dstDateColNum).setValue(evDate);

    linkNatively(dst, src);
}


/**
 * Calls @SharedFunction.fixCellErrors for either Hotel or PR sheet
 * 
 * @return {undefined} 
 */
/*function fixPRHotelCells() {
    var thisSheet = ss.getActiveSheet();

    if (thisSheet.getName() == pr.getName()) {
        SharedFunctions.fixCellErrors(pr.getRange(PR_START_ROW, PR_GRFCS_COL_NUM, mNumRows, (PR_TCKTS_COL_NUM - PR_GRFCS_COL_NUM + 1)));
    } else if (thisSheet.getName() == hotel.getName()) {
        SharedFunctions.fixCellErrors(hotel.getRange(H_START_ROW, H_STATUS_COL_NUM, mNumRows, (H_GUESTS_COL_NUM - H_STATUS_COL_NUM + 1)));
    } else {
        ui.alert("Vennligst velg PR- eller Hotell-arket før du forsøker å oppdatere.");
    }   
}*/











