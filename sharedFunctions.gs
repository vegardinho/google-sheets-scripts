/**
 * Iterates through all event cells in source sheet, and compares them to the destination sheet. 
 * Based on comparison result, either creates sheet, deletes entry/moves entries in destination sheet, 
 * or does nothing. Clears the remaining unused rows in destination sheet.
 * 
 * @param  {function}   newEventAction      Function to be called upon new entry in @findNextEvent
 * @param  {Sheet}      srcSheet            Function to with source content (main sheet)
 * @param  {Sheet}      dstSheet            Sheet to be filled with events from @srcSheet
 * @param  {Integer}    DST_FIRST_ROW       
 * @param  {Integer}    DST_LAST_ROW      
 * @param  {Integer}    DST_START_COL_NUM 
 * @param  {Integer}    DST_LAST_COL_NUM  
 * @param  {Integer}    DST_EVENT_COL_NUM 
 * @param  {Integer}    SRC_FIRST_ROW    
 * @param  {Integer}    SRC_LAST_ROW      
 * @param  {Integer}    SRC_EVENT_COL_NUM 
 * @param  {Integer}    MAX_CHECKS          See @findEventCells function 
 * 
 * @return {undefined} 
 */
function checkIfNewEvents(newEventAction, srcSheet, dstSheet, DST_FIRST_ROW, DST_LAST_ROW, DST_START_COL_NUM, 
        DST_LAST_COL_NUM, DST_EVENT_COL_NUM, DST_DATE_COL_NUM, SRC_FIRST_ROW, SRC_LAST_ROW, SRC_EVENT_COL_NUM, 
        SRC_DATE_COL_NUM, MAX_CHECKS) {
    var dstSheetRow = DST_FIRST_ROW;
    var srcRow = SRC_FIRST_ROW;

    // Run while not reached defined end in source sheet, or defined limit in destination sheet
    while ((srcRow <= SRC_LAST_ROW) && (dstSheetRow <= DST_LAST_ROW)) {
        srcRow = findNextEvent(srcRow, dstSheetRow, newEventAction, srcSheet, dstSheet, 
            DST_FIRST_ROW, DST_LAST_ROW, DST_START_COL_NUM, DST_LAST_COL_NUM, DST_EVENT_COL_NUM, 
            DST_DATE_COL_NUM, SRC_FIRST_ROW, SRC_LAST_ROW, SRC_EVENT_COL_NUM, SRC_DATE_COL_NUM, MAX_CHECKS); 
        dstSheetRow++;
    }

    dstSheet.getRange(dstSheetRow - 1, DST_START_COL_NUM, (DST_LAST_ROW - dstSheetRow + 1), DST_LAST_COL_NUM).clearContent();
}


/**
 * Finds next non-empty event cell in source sheet, puts event in right place in destination sheet, 
 * and calls on creation of event sheet if needed.
 * 
 * @param  {Integer} srcRow             Current row (event) in @srcSheet 
 * @param  {Integer} dstSheetRow        Next non-matched row (event) in @dstSheet
 * @param  {Function} newEventAction    
 * @param  {Sheet} srcSheet             
 * @param  {Sheet} dstSheet          
 * @param  {Integer} DST_FIRST_ROW     
 * @param  {Integer} DST_LAST_ROW      
 * @param  {Integer} DST_START_COL_NUM 
 * @param  {Integer} DST_LAST_COL_NUM  
 * @param  {Integer} DST_EVENT_COL_NUM 
 * @param  {Integer} SRC_FIRST_ROW     
 * @param  {Integer} SRC_LAST_ROW      
 * @param  {Integer} SRC_EVENT_COL_NUM 
 * @param  {Integer} MAX_CHECKS        
 * 
 * @return {Integer}                    Next row in @srcSheet (with corresponding event)   
 */
function findNextEvent(srcRow, dstSheetRow, newEventAction, srcSheet, dstSheet, 
            DST_FIRST_ROW, DST_LAST_ROW, DST_START_COL_NUM, DST_LAST_COL_NUM, DST_EVENT_COL_NUM, 
            DST_DATE_COL_NUM, SRC_FIRST_ROW, SRC_LAST_ROW, SRC_EVENT_COL_NUM, SRC_DATE_COL_NUM, MAX_CHECKS) {
    var srcRow = srcRow;
    var incrRows = -1;
    var to, from;


    while ((incrRows == -1) && (srcRow <= SRC_LAST_ROW)) {
                incrRows = findEventCells(srcRow, dstSheetRow, dstSheet, srcSheet, MAX_CHECKS, 
            SRC_EVENT_COL_NUM, SRC_DATE_COL_NUM, DST_EVENT_COL_NUM, DST_DATE_COL_NUM, DST_LAST_ROW);
        
        if (incrRows > 0) {
            // Event already added to destination sheet, but there are more events in destination sheet than there should be.
            // (Some have probably been deleted in source since last run.)
            // Overwrite events from dstSheetRow to dstSheetRow + i with the subsequent events.
            from = dstSheet.getRange(dstSheetRow + incrRows, DST_START_COL_NUM, (DST_LAST_ROW - dstSheetRow), DST_LAST_COL_NUM);
            to = dstSheet.getRange(dstSheetRow, DST_START_COL_NUM, (DST_LAST_ROW - dstSheetRow - incrRows), DST_LAST_COL_NUM);
            from.copyTo(to);

        } else if (incrRows < -1) {
        // Event cell in source sheet not empty, but no matches in destination sheet. Assuming new event.
            if (incrRows == -3) {
                // There are subsequent events in destination sheet. Move content down
                from = dstSheet.getRange(dstSheetRow, DST_START_COL_NUM, (DST_LAST_ROW - dstSheetRow), DST_LAST_COL_NUM);
                to = dstSheet.getRange((dstSheetRow + 1), DST_START_COL_NUM);
                from.copyTo(to);
                dstSheet.getRange(dstSheetRow, DST_START_COL_NUM, 1, DST_LAST_COL_NUM).clearContent();
            }

            newEventAction(dstSheetRow, srcRow, dstSheet);
        }

        srcRow++;
    }

    return srcRow;
}

/**
 * 
 * @param  {integer} dstRow  Row in destination sheet
 * @param  {integer} srcRow Row in source sheet
 * 
 */

/**
 * Searches through at most @MAX_CHECKS rows in destination sheet, to determine wether event is new, 
 * old, or skipped events in between.
 * 
 * @param  {Integer} srcRow            
 * @param  {Integer} dstRow           
 * @param  {Integer} dstSheet        
 * @param  {Integer} srcSheet       
 * @param  {Integer} MAX_CHECKS         Maximum number of iterations before assuming event has not 
 *                                      previously been added to @dstSheet   
 * @param  {Integer} SRC_EVENT_COL_NUM 
 * @param  {Integer} DST_EVENT_COL_NUM 
 * @param  {Integer} DST_LAST_ROW      
 * 
 * @return {integer}          -1: empty source sheet cell
                              -2: no more cells in destination sheet with content
                              -3: no matches of name from source cell in destination sheet
 */
function findEventCells(srcRow, dstRow, dstSheet, srcSheet, MAX_CHECKS, SRC_EVENT_COL_NUM, 
        SRC_DATE_COL_NUM, DST_EVENT_COL_NUM, DST_DATE_COL_NUM, DST_LAST_ROW) {
    var dstSheetRow = dstRow;
    var srcRow = srcRow;
    var srcEventCell = srcSheet.getRange(srcRow, SRC_EVENT_COL_NUM);
    var srcDate = new Date(srcSheet.getRange(srcRow, SRC_DATE_COL_NUM).getValue() + "Z");
    var dstEventCell, dstDate, dstEventCellVal, srcEventCellVal;
        
    for (var i = 0; i < MAX_CHECKS; i++) {
        dstEventCell = dstSheet.getRange(dstSheetRow + i, DST_EVENT_COL_NUM);
        dstDate = new Date(dstSheet.getRange(dstSheetRow + i, DST_DATE_COL_NUM).getValue() + "Z");
        dstEventCellVal = dstEventCell.getDisplayValue();
        srcEventCellVal = srcEventCell.getDisplayValue();
        
        // No event in row-cell in source sheet
        if (srcEventCell.isBlank()) {
            return -1;
        }
        // No point in iterating anymore, as no more events filled out in destination sheet;
        // or content in @srcSheet has moved down, so date cell link to @dstSheet must be updated.
        if ((dstEventCell.isBlank() && i === 0) || 
            ((dstEventCellVal == srcEventCellVal) && (dstDate.valueOf() != srcDate.valueOf() || 
            dstEventCell.getFormula() === "" || dstSheet.getRange(dstRow, DST_EVENT_COL_NUM + 1).getValue() == "#REF!"))) {
            return -2;
        }
        // Event already added to destination sheet, or ran out of legal rows in destination sheet
        if (dstEventCellVal == srcEventCellVal || dstRow + i > DST_LAST_ROW) {
            return i;
        }
    }

    return -3;
}



/**
 * Checks if reference of cells in @range is faulty, and attempts to reset the same formula
 * 
 * @param  {Range} range 
 * @return {undefined}       
 */
function fixCellErrors(range) {
    var range = range;
    var sheet = range.getSheet();
    var ss = sheet.getParent();

    var firstRow = range.getRow();
    var lastRow = range.getLastRow();
    var firstCol = range.getColumn();
    var lastCol = range.getLastColumn();

    var cell, formula;
            
    for (let i = firstRow; i < lastRow + 1; i++) {
        for (let l = firstCol; l < lastCol + 1; l++) {
            cell = sheet.getRange(i, l);
            formula = cell.getFormula();
            if (cell.getValue() == "#REF!") {
                cell.setFormula(formula);
            }
        }
    }
}





