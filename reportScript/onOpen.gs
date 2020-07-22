function onOpen() {
    checkHippoLink();
    sjekkHippo();
    ui.createMenu('Egendefinert')
    .addItem('Oppdater arrangementsliste fra Hippo', 'sjekkHippo')
    .addItem('Sett opp som nytt rapporteringsdokument', 'clearSheet')
    .addToUi();
}

function clearSheet() {
    var result = ui.alert(
        'Bekreft',
        'Er du sikker på at du vil slette alle arrangementer (inkludert ark)?',
        ui.ButtonSet.YES_NO);

    // Process the user's response.
    if (result == ui.Button.YES) {
        var sisteRad = FIRST_RAP_ROW;

        //Slett alle referanser til slettede ark, og slett arkene
        hovedark.getRange(FIRST_RAP_ROW, 3, (LAST_RAP_ROW - sisteRad + 1), 9).clearContent();
        deleteSheets(true);

    } else {
        ui.alert('Sletting avbrutt');
        return;
    }

    ui.alert('Alt innhold slettet');
    hippo.clearContents();
    checkHippoLink();
    ui.alert("Lenke til ny hippo lagt til! Henter arrangementer...");
    sjekkHippo();
    ui.alert("Alle arrangementer lagt til.");
}


function checkHippoLink() {
    if (hippo.getRange(1, 1).getFormula() != "") {
        return;
    }

    var hippoLenke = ui.prompt('For å få lagt inn arrangementer fra hippo, trenger jeg lenken til hippo-dokumentet. Eks: https://docs.google.com/spreadsheets/d/13OA-4EL-6t8w/edit#gid=0');
    if (hippoLenke.getSelectedButton() != ui.Button.OK) {
        return;
    }

    var hippoFormula = Utilities.formatString('=IMPORTRANGE("%s";"Kopi!A3:E")', hippoLenke.getResponseText());
    hippo.getRange(1, 1).setFormula(hippoFormula);
}


function fiksAlleLenker() {  
    //TODO: for-loop hvor fiksLenke() kalles for alle celler
    for (var i = 4; i < LAST_RAP_ROW; i++) {
        fiksLenke(hovedark.getRange(i, 4));
    }

    SpreadsheetApp.getUi().alert('Alle lenker er oppdatert!');
}


function deleteSheets(blockPrompt) {
    var ui = SpreadsheetApp.getUi();
    var result;

    if (!blockPrompt) {
        var result = ui.alert(
        'Bekreft',
        'Er du sikker på at du vil slette alle arrangementsark',
            ui.ButtonSet.YES_NO);
    }
    // Process the user's response.
    if ((result == ui.Button.YES) || (blockPrompt)) {
        var sisteRad = FIRST_RAP_ROW;

        if (sheets.length > 3) {

            var datoRef = sheets[3].getRange("C6").getFormula();
            sisteRad = parseInt(datoRef.substr(datoRef.length - 1));

            for (var i = 3; i < sheets.length; i++) {
                ss.deleteSheet(sheets[i])
            }
        }

    } else {
        ui.alert('Sletting avbrutt');
        return;
    }

    if (!blockPrompt) {
        SpreadsheetApp.getUi().alert('Alle ark slettet');
    }
}


function sjekkHippo() {
    //Finn siste rad med innhold i rapporteringsskjema
    //Finn tilsvarende rad i hippo
    //Let gjennom uthevede arr fra dato (ev gjennom sheets), og legg til i rapporteringsskjema

    var startRow = 4;
    var rapRow = startRow;
    var prevCell = hovedark.getRange((rapRow) - 1, 4);
    var thisCell, thisDate;

    var funnet;
    var lastHipRow = hippo.getRange("E1").getDisplayValue();
    var hipRow = 2;
    var arrNavn = hippo.getRange(hipRow, 2);
    var arrDato = hippo.getRange(hipRow, 1);
  
    while ((prevCell.getDisplayValue() != "") && (rapRow <= LAST_RAP_ROW) && (hipRow <= lastHipRow)) {
        thisCell = hovedark.getRange(rapRow, ARR_COL);
        thisDate = hovedark.getRange(rapRow, DATE_COL);
        funnet = false;

        while (!funnet && (hipRow <= lastHipRow)) {
            if (!arrDato.isBlank() && !arrNavn.isBlank()) {
                funnet = true;
            }

            if (funnet && ((arrNavn.getDisplayValue() != thisCell.getDisplayValue()) || (arrDato.getDisplayValue() != thisDate.getDisplayValue()))) {
                //move other stuff down if not last row (if commented out: overwrite existing with order on hippokopi)
                //        if (thisCell.getDisplayValue() != "") {
                //          var src = hovedark.getRange(rapRow, 3, (LAST_RAP_ROW - rapRow), 9);
                //          var dst = hovedark.getRange((rapRow + 1), 3, (LAST_RAP_ROW - rapRow + 1), 9);
                //          src.copyTo(dst, {contentsOnly:true});
                //        } 

                thisCell.setValue(arrNavn.getDisplayValue());
                var a1 = arrDato.getA1Notation();
                hovedark.getRange(rapRow, DATE_COL).setFormula(Utilities.formatString('=\'%s\'!$%s$%s', hippo.getName(), a1.substring(0,1), a1.substring(1)));
                fiksLenke(thisCell, hippo.getRange(hipRow, 3), hippo.getRange(hipRow, 4));
            }

            hipRow++;
            arrDato = hippo.getRange(hipRow, 1);
            arrNavn = hippo.getRange(hipRow, 2);
        }

        prevCell = thisCell;
        rapRow++;
    }

    hovedark.getRange(rapRow - 1, DATE_COL, LAST_RAP_ROW - rapRow + 1, 9).clearContent();
}

// Remove all range protections in the spreadsheet that the user has permission to edit.
function slettBeskyttelser() {
    var sheet = ss.getActiveSheet();
    var rangeProtections = sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE);
    var sheetProtections = sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET);

    for (var i = 0; i < rangeProtections.length; i++) {
        if (rangeProtections[i].canEdit()) {
            rangeProtections[i].remove();
        }
    }

    for (var i = 0; i < sheetProtections.length; i++) {
        if (sheetProtections[i].canEdit()) {
            sheetProtections[i].remove();
        }
    }
}