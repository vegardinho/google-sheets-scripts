/**
 * Creates custom menu and checks that link to hippo exists in hipo copy.
 * @return {undefined}
 */
function onOpen() {
    ui.createMenu('KU Supermeny')
    .addItem('Oppdater arrangementsliste fra Hippo', 'checkIfNewEvents')
    .addItem('Nullstill rapporteringsdokumentet', 'setUpAsNew')
    .addItem('Fiks arrangement', 'fixEventOnRow')
    .addToUi();
    checkHippoLink();
}

/**
 * Sets up document as new one. Clears the current one.
 * @return {undefined}
 */
function setUpAsNew() {
    var result;
    var result = ui.alert(
        'Bekreft',
        'Er du sikker på at du vil slette alle arrangementer (inkludert ark)?',
        ui.ButtonSet.YES_NO);

    // Process the user's response.
    if (result == ui.Button.YES) {
        lastRow = M_FIRST_RAP_ROW;

        // Delete all data on main sheet, and then delete all event sheets
        mainSheet.getRange(M_FIRST_RAP_ROW, M_DATE_COL_NUM, (M_LAST_RAP_ROW - lastRow + 1), M_LAST_COL_NUM).clearContent();
        deleteSheets(true);

    } else {
        ui.alert('Sletting avbrutt');
        return;
    }

    ui.alert('Alt innhold slettet');
    hippoCopy.clearContents();
    checkHippoLink();
    ui.alert("Lenke til ny hippo lagt til! Henter arrangementer... (Dette kan ta noe tid.)");
    Utilities.sleep(6000);
    checkIfNewEvents();
    ui.alert("Alle arrangementer lagt til.");
}

/**
 * Checks if link from hippo is specified and imported from hippo, in hippo copy.
 * @return {undefined}
 */
function checkHippoLink() {
    var hippoLink, hippoFormula;

    hcLinkCellFormula = hippoCopy.getRange(HC_LINK_CELL_RANGE).getFormula();
    if (hcLinkCellFormula != "") {
        return;
    }

    userLink = ui.prompt('For å få lagt inn arrangementer fra hippo, trenger jeg lenken til hippo-dokumentet. Eks: https://docs.google.com/spreadsheets/d/13OA-4EL-6t8w/edit#gid=0');
    if (userLink.getSelectedButton() != ui.Button.OK) {
        return;
    }

    hippoFormula = Utilities.formatString('=IMPORTRANGE("%s";"Kopi!A3:E")', userLink.getResponseText());
    hcLinkCell.setFormula(hippoFormula);
    hcLinkCellFormula = hippoFormula;
}


/**
 * Loops through all event cells in main sheet, and calls function for check (or creation) of corresponding sheet.
 * @return {undefined}
 */
function checkAllLinks() {  
    for (var i = 4; i < M_LAST_RAP_ROW; i++) {
        setUpEventSheet(i, M_EVENT_COL_NUM);
    }

    SpreadsheetApp.getUi().alert('Alle lenker er oppdatert!');
}


/**
 * Determines if element is in Array
 * @param  {Array[String]}
 * @param  {Object type in Array}
 * @return {Boolean}
 */
function isIn(sheetArray, element) {
    var list = sheetArray;
    var elmnt = element;

    for (var i = 0; i < list.length; i++) {
        if (element === sheetArray[i]) {
            return true;
        }
    }

    return false;
}

/**
 * Deletes all sheets exept those specified in KEEP_SHEETS
 * @param  {boolean}    blockPrompt     Determine wether to ask for user confirmation or not
 * @return {undefined}
 */
function deleteSheets(blockPrompt) {
    var result, dateLink;

    if (!blockPrompt) {
        result = ui.alert(
        'Bekreft',
        'Er du sikker på at du vil slette alle arrangementsark',
            ui.ButtonSet.YES_NO);
    }

    // Process the user's response.
    if ((result == ui.Button.YES) || (blockPrompt)) {

        for (var i = 0; i < sheets.length; i++) {
            if (isIn(KEEP_SHEETS, sheets[i].getName())) {
                continue;
            }
            ss.deleteSheet(sheets[i])
        }

    } else {
        ui.alert('Sletting avbrutt');
        return;
    }

    if (!blockPrompt) {
        SpreadsheetApp.getUi().alert('Alle ark slettet');
    }
}




/** 
 * Manually fixes/re-runs script for events already linked to main sheet.
 *
 * @param  {String} mainSheetRow  Row for broken cell in main sheet
 * @param  {String} hippoCopyRow  Row for related correct cell in hippo copy
 *
 * @return {Null}
 */
function fixEventOnRow(mainSheetRow, hippoCopyRow) {
    var response = ui.prompt('Hvilket arrangementsnummer er det snakk om?');
    var eventNum = Number(response.getResponseText());
    var hippoSheetRow, eventName, mainSheetRow;

    if (response.getSelectedButton() != ui.Button.OK || eventNum == NaN) {
        return;
    }

    mainSheetRow = eventNum + 3;
    eventName = mainSheet.getRange(mainSheetRow, M_EVENT_COL_NUM).getDisplayValue();

    for (let i = 1; i < HC_LAST_HIP_ROW; i++) {
        var cell = hippoCopy.getRange(i, HC_EVENT_COL_NUM);
        if (cell.getDisplayValue() == eventName) {
            hippoSheetRow = i;
            break;
        }
    }

    setUpEventSheet(mainSheetRow, hippoSheetRow);
}

/** 
 * Remove all range protections (that the user has permission to edit) in the current spreadsheet
 *
 * @param    {Null}
 *
 * @return   {Null}
 */
function deleteProtections() {
    var sheet = ss.getActiveSheet();
    var rangeProtections = sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE);
    var sheetProtections = sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET);

    for (var i = 0; i < rangeProtections.length; i++) {
        if (rangeprotections[i].canedit()) {
            rangeprotections[i].remove();
        }
    }

    for (var i = 0; i < sheetProtections.length; i++) {
        if (sheetProtections[i].canEdit()) {
            sheetProtections[i].remove();
        }
    }
}


/**
 * Gets sheet by name of event, if existing, or creates new.
 *
 * @param  {String} name    Name of event (and sheet).
 *
 * @return {Sheet}          The event sheet.
 */
function getSheet(name) {
    var newSheet, mainSheetLink;

    var eventName = name;
    var sheet = ss.getSheetByName(eventName);

    if (sheet === null) {
        newSheet = ss.insertSheet(eventName, ss.getSheets().length, {template: template});
        newSheet.getRange(TMPLT_EV_NAME_CELL).getMergedRanges()[0].setValue(eventName);
        sheet = newSheet;
    }

    return sheet;
}

/**
 * Updates link in sheet to main sheet (to be used if broken)
 * @param  {sheet}  The sheet to have its link updated.
 * @return {undefined}
 */
function updateLinkToHippo(sheet) {
    var mainSheetLink, newSheet;
    var newSheet = sheet;

    mainSheetLink = "=HYPERLINK(\"#gid=" + mainSheet.getSheetId() + "\";" + "\"" + "TILBAKE TIL HOVEDARK" + "\")";
    newSheet.getRange(TMPLT_MAIN_FORMULA_CELL).setFormula(mainSheetLink);
}

/**
 * Set up sheet with info provided by arguments. Create new report sheet if not existing,
 * then set up linking, formulas and protections regardless of prior existence.
 *
 *
 * @return   {Null}
 */
function setUpEventSheet(mainSheetRow, hippoCopyRow, dstSheet=null) {

    var cell = mainSheet.getRange(mainSheetRow, M_EVENT_COL_NUM);
    var evName = hippoCopy.getRange(hippoCopyRow, HC_EVENT_COL_NUM).getDisplayValue();
    var responsible = hippoCopy.getRange(hippoCopyRow, HC_RSPNSBLE_COL_NUM);
    var venue = hippoCopy.getRange(hippoCopyRow, HC_VENUE_COL_NUM);
    var date = hippoCopy.getRange(hippoCopyRow, HC_DATE_COL_NUM);

    cell.setValue(evName);
    var eventName = evName;
    var sheet = getSheet(eventName);
    
    // Go back visually to the sheet the user came from
    ss.setActiveSheet(mainSheet);
    mainSheet.setActiveSelection(cell);

    linkAppropriateCells(sheet, date, cell, responsible, venue);
    linkCellToSheet(sheet, cell);
    protectSheetAreas(sheet);
}


/** 
 * Protect range, then remove all other users from the list of editors.
 * Ensure the current user is an editor before removing others. Otherwise, if the user's edit
 * permission comes from a group, the script throws an exception upon removing the group.
 * @param   {Sheet}      eventSheet     Sheet to have its areas protected
 * @return  {undefined}
 */
function protectSheetAreas(eventSheet) {
    var sheet = eventSheet;

    var protections = sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE);

    if (protections.length === 0) {

        beskyttOmråde(TMPLT_EV_NAME_CELL, sheet);
        beskyttOmråde(TMPLT_INFO_CELLS, sheet);
        beskyttOmråde(TMPLT_TOTAL_CELL, sheet);
        beskyttOmråde(TMPLT_INCOME_CELL, sheet);
    }
}


/**
 * Pastes hyperlink to eventSheet in eventCell, and configures styling.
 *
 * @param  {Sheet}  eventSheet  Event sheet
 * @param  {Range}  eventCell   Event cell in main sheet
 * @return {undefined}
 */
function linkCellToSheet(eventSheet, eventCell) {
    var sheet = eventSheet;
    var cell = eventCell;

    var formula = "=HYPERLINK(\"#gid=" + sheet.getSheetId() + "\";" + "\"" + sheet.getName() + "\")";
    cell.setFormula(formula);
    cell.setFontLine('underline');
    cell.setFontColor('#1155cc'); 
}

/**
 * Sets up automatic linking of values. 1) From event sheet cells to main sheet, 2) responsible and venue from
 * hippo copy to event sheet, 3) date from main sheet to event sheet
 *
 * @param  {Sheet}  eventSheet      Event sheet
 * @param  {Range}  eventDate       Event date cell in hippo copy 
 * @param  {Range}  eventCell       Event cell in main sheet
 * @param  {Range}  eventRspnsble   Event responsible cell in hippo copy
 * @param  {Range}  eventVenue      Event venue cell in hippo copy
 * @return {undefined}
 */
function linkAppropriateCells(eventSheet, eventDate, eventCell, eventRspnsble, eventVenue) {
    var sheet = eventSheet;
    var date = eventDate;
    var responsible = eventRspnsble;
    var venue = eventVenue;
    var col = eventCell.getColumn();
    var row = eventCell.getRow();


    var range, sheetName;

    var src = [date, sheet.getRange(TMPLT_FB_PRTCPNT_CELL), sheet.getRange(TMPLT_FB_INTRSTD_CELL),
                sheet.getRange(TMPLT_INCOME_CELL), sheet.getRange(TMPLT_MEMB_CELL), sheet.getRange(TMPLT_N_MEMB_CELL), 
                sheet.getRange(TMPLT_OTHER_CELL), sheet.getRange(TMPLT_FILM_CELL), 
                responsible, venue, date];

            

    var dst = [mainSheet.getRange(row, M_DATE_COL_NUM), mainSheet.getRange(row, M_FB_PART_COL_NUM), 
                mainSheet.getRange(row, M_FB_INT_COL_NUM), mainSheet.getRange(row, M_INCOME_COL_NUM), 
                mainSheet.getRange(row, M_MEMB_COL_NUM), mainSheet.getRange(row, M_N_MEMB_COL_NUM),
                mainSheet.getRange(row, M_OTHER_COL_NUM), mainSheet.getRange(row, M_FILM_COL_NUM), 
                sheet.getRange(TMPLT_RSPNSBLE_CELL), sheet.getRange(TMPLT_VENUE_CELL), sheet.getRange(TMPLT_DATE_CELL)];


    for (var i = 0; i < dst.length; i++) {
        if (src[i] == undefined) {
            continue;
        }

        range = src[i].getA1Notation();
        sheetName = src[i].getSheet().getName();
        dst[i].setFormula(Utilities.formatString('=\'%s\'!$%s$%s', sheetName, range.substring(0,1), range.substring(1)));
    }
}


/**
 * Adds protection for range in specified sheet. Sets editor/editor rights to current editor of main sheet protections.
 * @param  {String}     protectRange    Range in sheet that is to be protected
 * @param  {Sheet}      eventRange      Event sheet
 * @return {undefined}
 */
function beskyttOmråde(protectRange, eventSheet) {
    var sheet = eventSheet;
    var range = protectRange;

    var me = mainSheet.getProtections(SpreadsheetApp.ProtectionType.SHEET)[0].getEditors()[0];  
    var protection = sheet.getRange(range).protect().setDescription('Ny beskyttelse');
    // Make sure admin is only editor
    protection.addEditor(me);
    protection.removeEditors(protection.getEditors());

    // Disable edit rights for whole domain
    if (protection.canDomainEdit()) {
        protection.setDomainEdit(false);
    }
}


function checkIfNewEvents() {
    SharedFunctions.checkIfNewEvents(setUpEventSheet, hippoCopy, mainSheet, M_FIRST_RAP_ROW, M_LAST_RAP_ROW, M_DATE_COL_NUM,
        M_LAST_COL_NUM, M_EVENT_COL_NUM, M_DATE_COL_NUM, HC_FIRST_HIP_ROW, HC_LAST_HIP_ROW, HC_EVENT_COL_NUM, HC_DATE_COL_NUM, 4); 
}

