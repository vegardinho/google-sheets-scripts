// Slett gamle instanser av trigger, og lag ny
function createSpreadsheetOnChangeTrigger() {
    ScriptApp.getProjectTriggers().slice()
        .forEach (function (d) {
                if (d.getHandlerFunction() === 'newChange') ScriptApp.deleteTrigger(d);
                });

    ScriptApp.newTrigger('newChange').forSpreadsheet(ss).onChange().create();
}


// Call setEventLink function on every row in changed range
function newChange() {
    if (ss.getActiveSheet().getName() !== hippo.getName()) {
        return;
    }
    var thisRange = ss.getActiveRange();
    var numRows = thisRange.getNumRows();
    var firstRow = thisRange.getRow();
    if (numRows > 1) {
        for (var i = 1; i <= numRows; i++) {
            setEventLink(thisRange.getCell(i, 1));
        }
    } else {
        setEventLink(thisRange);
    }
}

function rstEvVal(cell, newName, startDate, endDate) {
    var start = startDate;
    var end = endDate;
  
    clearOldValues(cell.getRow());
    updateCalendarEvent(newName, start, end, dlt=true);
    if (cell.getFormula() != "") {
        removeFormatting(cell, newName);
    }
}

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

// Check if cell is confirmed event; if so, link to adhering sheet
function setEventLink(valgtCelle) {
    var cell = valgtCelle;
    var evName = cell.getDisplayValue();
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

    //Return if wrong column or no values previously saved
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
        rstEvVal(cell, evName, date, endDate)
            return;
    }

    //If something existed before the change, clean up if change is deletion; otherwise, ask if change is new name or new event
    if (prv_nm_val != "") {
        if (evName === "") {
            dltEvent(col, row, cell, prv_nm_cll, prv_nm_val, evName, date, endDate);
            return;
            //Only change name if same event
        } else if (evName != prv_nm_val) {         
            str = Utilities.formatString("Er \"%s\" det samme arrangementet som \"%s\"?", prv_nm_val, evName, prv_nm_val);
            response = ui.alert(str, ui.ButtonSet.YES_NO_CANCEL);
            if (response == ui.Button.YES) {
                chngNameOfEv(prv_nm_val, evName);
                prv_name = prv_nm_val;
            } else if (response == ui.Button.NO) {
                prv_nm_cll.clearContent();
                ui.alert('Arrangementet \"' + prv_nm_val + '\" er nå slettet. Dersom du ønsker å bruke regnearket tilknyttet det slettede arrangementet (\"' + prv_nm_val + '\") senere, taster du inn dette navnet for valgt datocelle ' +
                        'for å gjenopprette arrangementsdataene.');
            } else {
                cell.setFormula(prv_nm_frml);
            }
        }
    }  
  
    
    sheet = getEvSheet(evName);
    sheetDate = sheet.getRange(TMPLT_DATE).getDisplayValue().substring(1);
    // Visually go back to main sheet, so user doesn't see sheet for long
    ss.setActiveSheet(hippo);
    ss.setActiveCell(cell);
  
    // Make user confirm new date, if conflicting dates between main sheet and adhering existing sheet 
    // Reset if event has duplicate name by error and not by intent (wishing to change date)
    if (sheet != null && sheetDate != "" && date != sheetDate) {
        var response = confirmNewDate(evName, sheetDate, date);  
        if (response === ui.Button.YES) {
            dltDplctEv(evName, sheetDate);
        } else {
            removeFormatting(cell, evName);
            return;
        }
    }

    formula = "=HYPERLINK(\"#gid=" + sheet.getSheetId() + "\";" + "\"" + cell.getValue() + "\")";
    setLinkStyle(cell, formula);
    dstCells = linkSheetValues(cell, sheet);

    // Event name and date not linked, but hardcoded. Runs on every change, so not an issue.
    dstCells[0].setValue(evName);
    dstCells[1].setValue(date);

    updateCopy(date, formula, cell.getRow());
    protectEvSheet(sheet, dstCells);
    updateCalendarEvent(evName, date, endDate, prv_name, dlt=false);

}

function confirmNewDate(evName, sheetDate, date) {
    var evName = evName;
    var sheetDate = sheetDate;
    var date = date;

    str = "Det ser ut som arrangementsarket med navnet \'" + evName + "\' står registrert med datoen " + sheetDate + ". " +
        "Ønsker du å endre datoen på dette arket fra " + sheetDate + " til " + date + "? Hvis du har flere arrangementer med samme navn, bør du " +
        "slå sammen cellene vertikalt for å samle i ett ark, eller endre navnet på ett av dem for å få to forskjellige arrangementsark (f.eks \'" + evName + " 2" + "\')";
    return  ui.alert("Bekreft ny dato", str, ui.ButtonSet.YES_NO_CANCEL);

}

function chngNameOfEv(prv_nm_val, newName) {
    var arv_nm_val = arv_nm_val;
    var evName = newName;
    var sheet = sheet;

    sheet = ss.getSheetByName(prv_nm_val);
    if (sheet === null) {
       getEvSheet(evName);
    } else {
       sheet.setName(evName);
       sheet.getRange(TMPLT_EV_NAME).setValue(evName);
    }
}

function dltEvent(column, row, cell, prv_nm_cll, prv_nm_val, evName, date, endDate) {
    var col = column;
    var row = row;
    var cell = cell;
    var prv_nm_cll = prv_nm_cll;
    var prv_nm_val = prv_nm_val;
    var evName = evName;
    var date = date;
    var endDate = endDate;

    result = ui.alert(
            'Bekreft sletting',
            'Er du sikker på at du vil slette arrangementet \"' + prv_nm_val + '\"?',
            ui.ButtonSet.YES_NO);
    if (result === ui.Button.YES) {
        clearOldValues(row);
        hippo.getRange(row, col, 1, 4).clearContent();
        removeFormatting(cell, "");
        updateCalendarEvent(evName, date, endDate, dlt=true);
    } else {
        cell.setFormula(prv_nm_cll.getFormula());
    }

}

function protectEvSheet(sht, dstClls) {
    var sheet = sht;
    var dstCells = dstClls;
    // Protect linked cells in event sheet. 
    // Some cells are concatenated, as adding protection seems to add a lot of overhead (processing time during script run)
    protections = sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE);
    if (protections.length < dstCells.lenght) {
        for (let i = 0; i < dstCells.length; i++) {
            protectRange(dstCells[i]) 
        }
    }
}

//Change link style
function setLinkStyle(cell, formula) {
    Logger.log("setter lenke");
    cell.setFormula(formula);
    cell.setFontLine('none');
    cell.setFontColor("black");
}
//Update copy-sheet link to event sheet, and insert date
function updateCopy(date, formula, row) {
    copy.getRange(CP_DATE + row).setValue(date);
    copy.getRange(CP_EVENT + row).setFormula(formula);
}

// Link cells in src to dst
function linkSheetValues(ev_nm_cll, sheet) {
    var row = ev_nm_cll.getRow();
    var col = ev_nm_cll.getColumn();

    var src = [ev_nm_cll, hippo.getRange(row, col-2), hippo.getRange(row, col+1), hippo.getRange(row, col+2), 
        hippo.getRange(row, col+3), hippo.getRange(row, col+4), pr.getRange(PR_FACE + row), pr.getRange(PR_BILL + row), 
        pr.getRange(PR_GRFCS + row), hotel.getRange(H_STATUS + row)];  

    var dst = [sheet.getRange(TMPLT_EV_NAME), sheet.getRange(TMPLT_DATE), sheet.getRange(TMPLT_VENUE), sheet.getRange(TMPLT_RSPNSBLE), 
        sheet.getRange(TMPLT_AT_WORK), sheet.getRange(TMPLT_COMMENTS), sheet.getRange(TMPLT_RLS_FB), sheet.getRange(TMPLT_RLS_TCKTS),
        sheet.getRange(TMPLT_GRFCS), sheet.getRange(TMPLT_HOTEL)];

    for (var i = 2; i < dst.length; i++) {
        dst[i].setFormula(Utilities.formatString('=\'%s\'!%s', src[i].getSheet().getName(), src[i].getA1Notation()));
    };

    return dst
}

function getEvSheet(evName) {
    var sheet = ss.getSheetByName(evName);

    // Create new sheet if non-existent
    if (sheet === null && evName != "") {
        sheet = ss.insertSheet(evName, ss.getSheets().length, {template: tmplt});
    } 

    return sheet;
}

function dltDplctEv(name, delDate) {
    for (var i = M_START_ROW; i <= M_NUM_ROWS; i++) {
        let cell = copy.getRange(CP_EVENT + i).getDisplayValue();
        let date = copy.getRange(CP_DATE + i).getDisplayValue().substring(1);
        let hippoCell = hippo.getRange(i, M_EVENT_COL_NUM);
        let hippoName = hippoCell.getDisplayValue();

        if (cell === name && date === delDate) {
            console.log(cell);
            console.log(date);
            console.log(hippoName);
            clearOldValues(i);
            if (hippoName === name) {
                removeFormatting(hippoCell, "");
            }
            return;
        }
    }
}


//Deletes date and event name cells in copy sheet
function clearOldValues(row) {
    copy.getRange(CP_EVENT + row).clearContent();
    copy.getRange(CP_DATE + row).clearContent();
}

function removeFormatting(cell, name) {
    cell.setFontWeight("normal");
    cell.setValue(name);
}

// Protect range @områdeStreng in sheet @sheet, then remove all other users from the list of editors.
// Ensure the current user is an editor before removing others. Otherwise, if the user's edit
// permission comes from a group, the script throws an exception upon removing the group.
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

function makeDate(row) {
    var col = M_EVENT_COL_NUM;
    var mnth = monthToNumber[hippo.getRange(row, col-4).getMergedRanges()[0].getDisplayValue()].toString();
    var dateNum = hippo.getRange(row, col-2);

    if (dateNum.isPartOfMerge()) {
        dateNum = dateNum.getMergedRanges()[0].getDisplayValue();
    } else {
        dateNum = dateNum.getDisplayValue();
    }

    if (dateNum.length === 1) {
        dateNum = 0 + dateNum;
    }

    if (mnth.length === 1) {
        mnth = "0" + mnth;
    }

    return dateNum + "." + mnth;
}