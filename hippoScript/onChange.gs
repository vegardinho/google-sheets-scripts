/**
 * Slett gamle instanser av trigger, og lag ny
 *
 * @return {undefined}
 */
 function createSpreadsheetOnChangeTrigger() {
    ScriptApp.getProjectTriggers().slice()
    .forEach (function (d) {
        if (d.getHandlerFunction() === 'newChange') ScriptApp.deleteTrigger(d);
    });

    ScriptApp.newTrigger('newChange').forSpreadsheet(ss).onChange().create();
}


/**
 * Call setEventLink function on every row in changed range
 *
 * @return {undefined} 
 */
 function newChange() {
    if (ss.getActiveSheet().getName() !== hippo.getName()) {
        return;
    }
    
    var thisRange = ss.getActiveRange();
    var firstRow = thisRange.getRow();
    var numRows = thisRange.getNumRows();
    
    if (numRows > 1 && !thisRange.isPartOfMerge()) {         
        for (var i = 1; i <= numRows; i++) {
            setEventLink(thisRange.getCell(i, 1));
        }
    } else {
        setEventLink(thisRange);
    }
}

/**
 * Deletes event in copy sheet, and calendar event corresponding to cell.
 * Remove formatting if existing.
 * 
 * @param  {Range}  cell        Event cell
 * @param  {String} name        Name of event
 * @param  {String} startDate   
 * @param  {String} endDate   
 * @return {undefined}           
 */
 function rstEvVal(cell, name, startDate, endDate) {
    var start = startDate;
    var end = endDate;
    
    clearCopyValues(cell.getRow());
    updateCalendarEvent(name, start, end, prvName=null, dlt=true);
    if (cell.getFormula() != "") {
        resetCellValue(cell, name);
    }
}

/**
 * Returns dates from cell rows.
 * 
 * @param  {Integer}    row     
 * @param  {Integer}    lastRow 
 * @return {String[]}               Dates
 */
 function getDates(row, lastRow) {
    var row = row;
    var lastRow = lastRow;
    var startDate = makeDate(row);
    var endDate;
    
    if (row === lastRow) {
        endDate = startDate;
    } else {
        endDate = makeDate(lastRow);
    }
    

    return [startDate, endDate];
}


/**
 * Creates link for event cell to event sheet (and create calendar event), if relevant.
 * Changes event sheet and/or references in copy sheet (as well as calendar), if existing.
 * 
 * @param {Integer}     valgtCelle      Cell with change
 */
 function setEventLink(valgtCelle) {
    var cell = valgtCelle;
    var evName = replaceIllegalChars(cell.getDisplayValue());

    var formula = cell.getFormula();
    var txtStyle = cell.getRichTextValue().getTextStyle();
    var col = cell.getColumn();
    var row = cell.getRow();
    var lastRow = row;
    var prv_nm_cll = copy.getRange(CP_EVENT + row);
    var prv_nm_val = prv_nm_cll.getDisplayValue();
    var prv_nm_frml = prv_nm_cll.getFormula();
    var prv_name = null;
    var sheet, date, sheet, sheetDate, str, response, dstCells, protections, endDate, dates, crtNew;
    
    //Return if wrong column or both empty and no values previously saved
    if ((cell.getColumn() !== M_EVENT_COL_NUM) || (evName === "" && prv_nm_val === "")) {
        return;
    } 

    if (cell.isPartOfMerge()) {
        cell = cell.getMergedRanges()[0];
    }
    
    lastRow = cell.getLastRow();
    dates = getDates(row, lastRow);
    date = dates[0];
    endDate = dates[1];

    //Return if empty or not bold, erase formatting if link exists
    if (!txtStyle.isBold()) {
        rstEvVal(cell, evName, date, endDate);
        return;
    }
    
    sheet = getEvSheet(evName);
    if (sheet) {
        // Set UTC time
        sheetDate = new Date(sheet.getRange(TMPLT_DATE).getValue() + "Z");
    }
    // Visually go back to main sheet, so user doesn't see sheet for long
    ss.setActiveSheet(hippo);
    ss.setActiveCell(cell);
        // Make user confirm new date, if conflicting dates between main sheet and adhering existing sheet 
    // Reset if event has duplicate name by error and not by intent (wishing to change date)
    if (sheet != null && sheetDate != "Invalid Date" && date.valueOf() != sheetDate.valueOf()) {
        var response = confirmNewDate(evName, sheetDate.toLocaleDateString(), date.toLocaleDateString());  
        if (response === ui.Button.YES) {
            updateCalendarEvent(evName, sheetDate, sheetDate, prvName=null, dlt=true);
            dltDplctEv(evName, sheetDate);
        } else {
            resetCellValue(cell, evName);
            return;
        }
    }
    
    //If something existed before the change, clean up if change is deletion; otherwise, ask if change is new name or new event
    if (prv_nm_val != "") {
        if (evName === "") {
            dltEvent(col, row, cell, prv_nm_cll, prv_nm_val, evName, date, endDate);
            return;
            //Only change name if same event
        }
        if (evName != prv_nm_val) {         
            str = Utilities.formatString("Er \"%s\" det samme arrangementet som \"%s\"?", prv_nm_val, evName, prv_nm_val);
            response = ui.alert(str, ui.ButtonSet.YES_NO_CANCEL);
            if (response == ui.Button.YES) {
                chngNameOfEv(prv_nm_val, evName);
                prv_name = prv_nm_val;
            } else if (response == ui.Button.NO) {
                prv_nm_cll.clearContent();
              /*
                ui.alert('Arrangementet \"' + prv_nm_val + '\" er nå slettet. Dersom du ønsker å bruke regnearket tilknyttet det slettede arrangementet (\"' + prv_nm_val + '\") senere, taster du inn dette navnet for valgt datocelle ' +
                'for å gjenopprette arrangementsdataene.');*/
            } else {
                cell.setFormula(prv_nm_frml);
            }
        }
    }  
    
    sheet = ss.getSheetByName(evName);
    formula = "=HYPERLINK(\"#gid=" + sheet.getSheetId() + "\";" + "\"" + evName + "\")";
    setLinkStyle(cell, formula);
    dstCells = setSheetValues(cell, sheet, evName, date);

    updateCopy(date, formula, cell.getRow());
    protectEvSheet(sheet, dstCells);
    updateCalendarEvent(evName, date, endDate, prv_name, dlt=false);
}

function setSheetLink(cell, formula) {
    var nameStr = cell.getValue();
    var formula = formula = "=HYPERLINK(\"#gid=" + sheet.getSheetId() + "\";" + "\"" + nameStr + "\")";
    setLinkStyle(cell, formula);
}


/**
 * Replaces characters that break links (e.g. "'")
 *
 * @param  {String} stroing  The text to be analyzed
 *
 * @return {String}          The text with breaking characters removed
 */
 function replaceIllegalChars(stroing) {

  var streng = "";
  var changed = false;
  var rmvChars = ['\'', '"'];
  var rplcmntChars = ['´', '´'];


  
  for (var i = 0; i < stroing.length; i++) {
      for (var l = 0; l < rmvChars.length; l++) {
         if (stroing[i] == rmvChars[l]) {
           streng += rplcmntChars[l];
           changed = true;
           break;
       }
   }
   if (!changed) {
      streng += stroing[i];
  }
  changed = false;
}

return streng;
}

/**
 * @param  {String} evName    
 * @param  {String} sheetDate 
 * @param  {String} date      
 * @return {String}          Response from user
 */
 function confirmNewDate(evName, sheetDate, date) {
    var evName = evName;
    var sheetDate = sheetDate;
    var date = date;

    str = "Det ser ut som arrangementsarket med navnet \'" + evName + "\' står registrert med datoen " + sheetDate + ". " +
    "Ønsker du å endre datoen på dette arket fra " + sheetDate + " til " + date + "?";
    return  ui.alert("Bekreft ny dato", str, ui.ButtonSet.YES_NO_CANCEL);

}

/**
 * Change name of sheet, and cell value in sheet, to fit new event name.
 * First, delete any existing sheets with the new name (avoid duplicates).
 * 
 * @param  {String} prv_nm_val 
 * @param  {String} newName    
 * @return {undefined}            [description]
 */
 function chngNameOfEv(prv_nm_val, newName) {
    var arv_nm_val = arv_nm_val;
    var evName = newName;
    var exstSheet = ss.getSheetByName(newName);
    var sheet = ss.getSheetByName(prv_nm_val);

    if (sheet === null) {
        getEvSheet(evName);
    } else {
        if (exstSheet != null) {
            ss.deleteSheet(exstSheet);
        }
        sheet.setName(evName);
        sheet.getRange(TMPLT_EV_NAME).setValue(evName);
    }
}

/**
 * Clears all references to event: in main sheet, in copy, and in calendar.
 *  
 * @param  {Integer}    column      Event cell column
 * @param  {Row}        row         Event cell row
 * @param  {Cell}       cell        Event cell       
 * @param  {Range}      prv_nm_cll  Copy sheet event cell
 * @param  {String}     prv_nm_val 
 * @param  {String}     evName     
 * @param  {Integer}    date       
 * @param  {Integer}    endDate    
 * @return {undefined}            
 */
 function dltEvent(column, row, cell, prv_nm_cll, prv_nm_val, evName, date, endDate) {
    var col = column;
    var row = row;
    var cell = cell;
    var prv_nm_cll = prv_nm_cll;
    var prv_nm_val = prv_nm_val;
    var evName = evName;
    var date = date;
    var endDate = endDate;

  /*
    result = ui.alert(
            'Bekreft sletting',
            'Er du sikker på at du vil slette arrangementet \"' + prv_nm_val + '\"?',
            ui.ButtonSet.YES_NO);
            if (result === ui.Button.YES) {*/
              
                clearCopyValues(row);
                resetCellValue(cell, "");

                updateCalendarEvent(evName, date, endDate, prvName=null, dlt=true); 
  /*
    } else {
        cell.setFormula(prv_nm_cll.getFormula());
    }*/
}

/**
 * Adds protection for all "dstClls" in "sht"
 * 
 * @param  {Sheet}      sht     
 * @param  {Range[]}    dstClls 
 * @return {undefined}         
 */
 function protectEvSheet(sht, dstClls) {
    var sheet = sht;
    var dstCells = dstClls;
    var protections, length;
    var rangeLength = dstCells.length;
    // Protect linked cells in event sheet. 
    // Some cells are concatenated, as adding protection seems to add a lot of overhead
    protections = sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE);
    if (protections == null) {
        length = 0;
    } else {
        length = protections.length;
    }

    if (length < rangeLength) {
        for (let i = 0; i < 6; i++) {
            protectRange(dstCells[i]);
        }
    }
}

/**
 * Sets formula and formula style for "cell"
 * @param {Range}   cell    
 * @param {String}  formula 
 */
 function setLinkStyle(cell, formula) {

    cell.setFormula(formula);
    cell.setFontLine('none');
    cell.setFontColor("black");
}

/**
 * Set "formula" and "date" as formula and date for row, in copy sheet
 * 
 * @param  {String} date    
 * @param  {String} formula 
 * @param  {Integer} row    
 * @return {undefined}       
 */
 function updateCopy(date, formula, row) {
    copy.getRange(CP_DATE + row).setValue(date.toLocaleDateString());
    copy.getRange(CP_EVENT + row).setFormula(formula);
}

/**
 * Sets values in event sheet, either by linkink or hardcoding.
 * 
 * @param {Range}   ev_nm_cll   Event cell
 * @param {Sheet}   sht       
 * @param {String}  evName    
 * @param {String}  date      
 */
 function setSheetValues(ev_nm_cll, sht, evName, date) {
    var sheet = sht;
    var row = ev_nm_cll.getRow();
    var col = ev_nm_cll.getColumn();
    var ev_name = evName;
    var ev_date = date.toLocaleDateString();
    var name_range = sheet.getRange(TMPLT_EV_NAME);
    var date_range = sheet.getRange(TMPLT_DATE);

    var src = [hippo.getRange(row, M_VENUE_COL_NUM), hippo.getRange(row, M_RSPN_COL_NUM), 
    hippo.getRange(row, M_WRKNG_COL_NUM), hippo.getRange(row, M_COMM_COL_NUM)];

    var dst = [sheet.getRange(TMPLT_VENUE), sheet.getRange(TMPLT_RSPNSBLE), 
    sheet.getRange(TMPLT_AT_WORK), sheet.getRange(TMPLT_COMMENTS)];

    // Event name and date not linked, but hardcoded. Runs on every change, so not an issue.
    name_range.setValue(ev_name);
    date_range.setValue(ev_date);

    linkNatively(dst, src);
    return [name_range, date_range].concat(dst);
}


function linkNatively(dst, src) {
    for (var i = 0; i < dst.length; i++) {
        dst[i].setFormula(Utilities.formatString('=\'%s\'!$%s$%s', src[i].getSheet().getName(), 
            src[i].getA1Notation()[0], src[i].getRow()));
    };
}

/**
 * Returns event sheet.
 * Creates new if not existing.
 * 
 * @param  {String} evName
 * @return {Sheet}      
 */
 function getEvSheet(evName) {
    var sheet = ss.getSheetByName(evName);

    // Create new sheet if non-existent
    if (sheet === null && evName != "") {
        sheet = ss.insertSheet(evName, ss.getSheets().length, {template: tmplt});
    } 

    return sheet;
}

/**
 * Delete all occurences of "name" in main sheet and copy sheet with "delDate"
 * 
 * @param  {String} name        Event name
 * @param  {String} delDate     Date of event to be deleted
 * @return {undefined}         
 */
 function dltDplctEv(name, delDate) {
    for (var i = M_START_ROW; i <= mNumRows; i++) {
        let cell = copy.getRange(CP_EVENT + i).getDisplayValue();
        let date = copy.getRange(CP_DATE + i).getDisplayValue().substring(1);
        let hippoCell = hippo.getRange(i, M_EVENT_COL_NUM);
        let hippoName = hippoCell.getDisplayValue();

        if (cell === name && date === delDate) {
            clearCopyValues(i);
            if (hippoName === name) {
                resetCellValue(hippoCell, "");
            }
            return;
        }
    }
}


/**
 * Clear main sheet EVENT and DATE cells.
 * 
 * @param  {Integer} row 
 * @return {undefined}     
 */
 function clearCopyValues(row) {
    copy.getRange(CP_EVENT + row).clearContent();
    copy.getRange(CP_DATE + row).clearContent();
}


/**
 * Un-bolds cell text, and replaces value with "name"
 * 
 * @param  {Range}  cell 
 * @param  {String} name    New value of cell
 * @return {undefined}      
 */
 function resetCellValue(cell, name) {
    cell.setFontWeight("normal");
    cell.setValue(name);
}

/**
 * Protect "rng" from changes, with first editor of copy sheet as sole editor.
 * Denies any domain editing allowances.
 * 
 * @param  {Range} rng 
 * @return {undefined}   
 */
 function protectRange(rng) {
    var range = rng
    var me = copy.getProtections(SpreadsheetApp.ProtectionType.SHEET)[0].getEditors()[0];
    var protection = range.protect().setDescription('Ny beskyttelse');
    protection.addEditor(me);
    protection.removeEditors(protection.getEditors());

    if (protection.canDomainEdit()) {
        protection.setDomainEdit(false);
    }
}


/**
 * Return date on format "DD.MM" from row in main sheet

 * @param  {Integer}    row     Row in main sheet
 * @return {String}             Formatted date string
 */
 function makeDate(row) {
    var col = M_EVENT_COL_NUM;
    var month = monthToNumber[hippo.getRange(row, M_MNTH_COL_NUM).getMergedRanges()[0].getDisplayValue()].toString();
    var dateNum = hippo.getRange(row, col-2);
    var year = M_END_DATE.substring(0,4);
    var str;

    if (dateNum.isPartOfMerge()) {
        dateNum = dateNum.getMergedRanges()[0].getDisplayValue();
    } else {
        dateNum = dateNum.getDisplayValue();
    }

    // Return date object in UTC
    str = `${year}-${month}-${dateNum}Z`;

    return new Date(str);
}