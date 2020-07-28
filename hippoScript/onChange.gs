// Slett gamle instanser av trigger, og lag ny
function createSpreadsheetOnChangeTrigger() {
    ScriptApp.getProjectTriggers().slice()
        .forEach (function (d) {
                if (d.getHandlerFunction() === 'nyEndring') ScriptApp.deleteTrigger(d);
                });

    var ss = SpreadsheetApp.getActive();
    ScriptApp.newTrigger('nyEndring').forSpreadsheet(ss).onChange().create();
}


function nyEndring() {
    if (ss.getActiveSheet().getName() !== "Hovedark") {
      return;
    }
  
        var thisRange = ss.getActiveRange();
        var numRows = thisRange.getNumRows();
        var firstRow = thisRange.getRow();
        if (numRows > 1) {
            for (var i = 1; i <= numRows; i++) {
                fiksLenke(thisRange.getCell(i, 1));
            }
        } else {
            fiksLenke(thisRange);
        }           
}



function fiksLenke(valgtCelle) {
    var celle = valgtCelle;

    if (celle.getColumn() !== 5) {
        return;
    } 

    if (celle.isPartOfMerge()) {
        celle = celle.getMergedRanges()[0];
    }

    var arrNavn = celle.getDisplayValue();
    var formel = celle.getFormula();
    var tekstStil = celle.getRichTextValue().getTextStyle();
    var col = celle.getColumn();
    var row = celle.getRow();
    var gammeltNavn = kopi.getRange(CP_EVENT + row);
    var gammelDato = kopi.getRange(CP_DATE + row);
    var gammelVerdi = gammeltNavn.getDisplayValue();
    var gammelFormel = gammeltNavn.getFormula();
    
    //Return if empty or not bold, erase formatting if link exists
    if (!tekstStil.isBold()) {
        gammeltNavn.clearContent();
        gammelDato.clearContent();
        if (formel != "") {
            removeFormatting(celle, arrNavn);
        }
        return;
    }

    var dato = lagDato(row);
    var ark = ss.getSheetByName(arrNavn);

    // Create new sheet if non-existent, warn if date doesn't match event sheet date
    if (ark === null && arrNavn != "") {
        ark = ss.insertSheet(arrNavn, ss.getSheets().length, {template: mal});

    } else if (ark != null && dato != ark.getRange(DATO_CELLE).getDisplayValue().substring(1)) {
        var arkDato = ark.getRange(DATO_CELLE).getDisplayValue().substring(1);
        var streng = "Det ser ut som arrangementsarket med navnet \'" + arrNavn + "\' står registrert med datoen " + arkDato + ". " +
          "Ønsker du å endre datoen på dette arket fra " + arkDato + " til " + dato + "? Hvis du har flere arrangementer med samme navn, bør du " +
            "slå sammen cellene vertikalt for å samle i ett ark, eller endre navnet på ett av dem for å få to forskjellige arrangementsark (f.eks \'" + arrNavn + " 2" + "\')";
        var response = ui.alert("Bekreft ny dato", streng, ui.ButtonSet.YES_NO_CANCEL);

        //Reset if event has duplicate name by error and not by intent (wishing to change date)
        if (response === ui.Button.YES) {
          console.log("hei");
           dltDplctEv(arrNavn, arkDato);
          console.log("hei igjen");
        } else {
          celle.setFontWeight('normal');
          return;
        }
    }

    //If something existed before the change, clean up if change is deletion; otherwise, ask if change is new name or new event
    if (gammelVerdi != "") {
        if (arrNavn === "") {
            var result = ui.alert(
                    'Bekreft sletting',
                    'Er du sikker på at du vil slette arrangementet \"' + gammelVerdi + '\"?',
                    ui.ButtonSet.YES_NO);
            if (result == ui.Button.YES) {
                removeFormatting(celle, gammeltNavn);
                clearOldValues(row);
                hippo.getRange(row, col, 1, 4).clearContent();
            } else {
                celle.setFormula(gammeltNavn.getFormula());
            }
            return;

        } else if (arrNavn != gammelVerdi) {
            var streng = Utilities.formatString("Arrangementet på denne raden var tidligere \"%s\". Er \"%s\" det samme arrangementet som \"%s\"?", gammelVerdi, arrNavn, gammelVerdi);
            var response = ui.alert(streng, ui.ButtonSet.YES_NO);

            //Only change name if same event
            if (response == ui.Button.YES) {
                var regex = /(?<=;").*(?="\))/;
                var nyFormel = gammelFormel.replace(regex, arrNavn);
                celle.setFormula(nyFormel);
                gammeltNavn.setFormula(nyFormel);
                ark = ss.getSheetByName(gammelVerdi);
                ark.setName(arrNavn);
                ark.getRange(ARR_NAVN_ARK).setValue(arrNavn);
                ss.setActiveSheet(hippo);

                return;

            } else if (response == ui.Button.NO) {
                gammeltNavn.clearContent();
                ui.alert('Arrangementet \"' + gammelVerdi + '\" er nå slettet. Dersom du ønsker å bruke regnearket tilknyttet det slettede arrangementet (\"' + gammelVerdi + '\") senere, taster du inn dette navnet for valgt datocelle ' +
                        'for å gjenopprette arrangementsdataene.');
            } else {
                celle.setFormula(gammelFormel);
            }
        }
    }

    ss.setActiveSheet(hippo);
    hippo.setActiveSelection(celle);

    //Lenk sammen celler fra hippo, til arrangement-arket og pr-arket
    var src = [celle, hippo.getRange(row, col-2), hippo.getRange(row, col+1), hippo.getRange(row, col+2), 
        hippo.getRange(row, col+3), hippo.getRange(row, col+4), pr.getRange(PR_FACE + row), pr.getRange(PR_BILL + row), 
               pr.getRange(PR_GRAF + row), h_sheet.getRange(H_STATUS + row)];  
  
    var dst = [ark.getRange(ARR_NAVN_ARK), ark.getRange(DATO_CELLE), ark.getRange(LOKALE), ark.getRange(ANSVARLIG), 
        ark.getRange(PÅ_JOBB), ark.getRange(COMMENTS), ark.getRange(FACE_SLIPP), ark.getRange(BILL_SLIPP),
              ark.getRange(GRAFIKK), ark.getRange(STAT_HOTEL)];

    for (var i = 2; i < dst.length; i++) {
        dst[i].setFormula(Utilities.formatString('=\'%s\'!%s', src[i].getSheet().getName(), src[i].getA1Notation()));
    };

    //Arrnavn og dato endres under skriptkjøring, og trenger ikke lenking.
    dst[0].setValue(src[0].getDisplayValue());
    dst[1].setValue(dato);

    //Update copy-sheet link to event sheet, and insert date
    var formel = "=HYPERLINK(\"#gid=" + ark.getSheetId() + "\";" + "\"" + celle.getValue() + "\")";
    gammelDato.setValue(dato);
    gammeltNavn.setFormula(formel);

    //Change link style
    celle.setFormula(formel);
    celle.setFontLine('none');
    celle.setFontColor("black");


    // Protect linked cells in event sheet. 
    // Some cells are concatenated, as adding protection seems to add a lot of overhead (processing time during script run)
    var protections = ark.getProtections(SpreadsheetApp.ProtectionType.RANGE);

    if (protections.length < 8) {

        beskyttOmråde(ARR_NAVN_ARK, ark);
        beskyttOmråde(DATO_CELLE + ":" + ANSVARLIG, ark);
        beskyttOmråde(PÅ_JOBB, ark);
        beskyttOmråde(COMMENTS, ark);
        beskyttOmråde(FACE_SLIPP, ark);
        beskyttOmråde(BILL_SLIPP, ark);
        beskyttOmråde(GRAFIKK, ark);
        beskyttOmråde(STAT_HOTEL, ark);
    }
}

function dltDplctEv(name, delDate) {
  console.log(name);
  console.log(delDate);
  for (var i = START_ROW; i <= ANT_RADER; i++) {
    let cell = kopi.getRange(CP_EVENT + i).getDisplayValue();
    let date = kopi.getRange(CP_DATE + i).getDisplayValue().substring(1);
    let hippoCell = hippo.getRange(i, ARR_COLUMN);
    let hippoName = hippoCell.getDisplayValue();
    console.log(cell);
    console.log(date);
    
    if (cell === name && date === delDate) {
      console.log(cell);
      console.log(date);
      console.log(hippoName);
      clearOldValues(i);
      if (hippoName === name) {
        hippoCell.clearContent();
      }
      return;
    }
  }
}
        

//Deletes date and event name cells in copy sheet
function clearOldValues(row) {
    kopi.getRange(CP_EVENT + row).clearContent();
    kopi.getRange(CP_DATE + row).clearContent();
}

function removeFormatting(cell, name) {
    cell.setFontWeight("normal");
    cell.setValue(name);
}

 // Protect range @områdeStreng in sheet @ark, then remove all other users from the list of editors.
 // Ensure the current user is an editor before removing others. Otherwise, if the user's edit
 // permission comes from a group, the script throws an exception upon removing the group.
function beskyttOmråde(områdeStreng, ark) {
    var me = kopi.getProtections(SpreadsheetApp.ProtectionType.SHEET)[0].getEditors()[0];
    var protection = ark.getRange(områdeStreng).protect().setDescription('Ny beskyttelse');
    protection.addEditor(me);
    protection.removeEditors(protection.getEditors());

    if (protection.canDomainEdit()) {
        protection.setDomainEdit(false);
    }
}

function lagDato(row) {
    var col = ARR_COLUMN;
    var måned = monthToNumber[hippo.getRange(row, col-4).getMergedRanges()[0].getDisplayValue()].toString();
    var tallDag = hippo.getRange(row, col-2);

    if (tallDag.isPartOfMerge()) {
        tallDag = tallDag.getMergedRanges()[0].getDisplayValue();
    } else {
        tallDag = tallDag.getDisplayValue();
    }

    if (tallDag.length === 1) {
        tallDag = 0 + tallDag;
    }
  
    if (måned.length === 1) {
        måned = "0" + måned;
    }

    return tallDag + "." + måned;
}