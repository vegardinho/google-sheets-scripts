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
    ss = SpreadsheetApp.getActiveSpreadsheet();
    if (ss.getActiveSheet().getName() === "Hovedark") {
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
    var gammeltNavn = kopi.getRange(row, 2);
    var gammelDato = kopi.getRange(row, 1);
    var gammelVerdi = gammeltNavn.getDisplayValue();
    var gammelFormel = gammeltNavn.getFormula();
    
    console.log(formel);
    console.log(arrNavn);
    
  //Return if empty or not bold, erase formatting if link exists
  if (!tekstStil.isBold() || (arrNavn === "" && gammelVerdi === "")) {
    if (formel != "") {
        
        removeFormatting(celle, arrNavn);
    }
    return;
  }
  
  var dato = lagDato(row);
  var ark = ss.getSheetByName(arrNavn);

   // Legg til hvis ikke eksisterende
   if (ark === null && arrNavn != "") {

    var nyttArk = ss.insertSheet(arrNavn, ss.getSheets().length, {template: mal});
    var hippoLenke = "=HYPERLINK(\"#gid=" + hippo.getSheetId() + "\";" + "\"" + "TILBAKE TIL HIPPO" + "\")";
    nyttArk.getRange("A1").setFormula(hippoLenke);
    ark = nyttArk;

   } else if (ark != null && dato != ark.getRange(DATO_CELLE).getDisplayValue().substring(1)) {
    var arkDato = ark.getRange(DATO_CELLE).getDisplayValue().substring(1);
    var streng = "Det ser ut som arrangementsarket med navnet \'" + arrNavn + "\' står registrert med datoen " + arkDato + ". " + 
    "Ønsker du å endre datoen på dette arket fra " + arkDato + " til " + dato + "? Hvis du har flere arrangementer med samme navn, bør du " +
    "slå sammen cellene vertikalt for å samle i ett ark, eller endre navnet på ett av dem for å få to forskjellige arrangementsark (f.eks \'" + arrNavn + " 2" + "\')";
    var response = ui.alert("Bekreft ny dato", streng, ui.ButtonSet.YES_NO_CANCEL);
    
    if (response != ui.Button.YES) {
        removeFormatting(celle, gammeltNavn);
        celle.setFontWeight('normal');
        return;
    }
   }
   
   if (gammelVerdi != "") {
    if (arrNavn === "") {
        var result = ui.alert(
            'Bekreft sletting',
            'Er du sikker på at du vil slette arrangementet \"' + gammelVerdi + "\"? " +
            'Hvis du etter sletting ønsker å bruke regnearket tilknyttet dette arrangementet senere, kan du taste inn navnet \"' + gammelVerdi + '\" for valgt celle.',
            ui.ButtonSet.YES_NO);
        if (result == ui.Button.YES) {
            removeFormatting(celle, gammeltNavn);
            gammeltNavn.clearContent();
            hippo.getRange(row, col, 1, 4).clearContent();
        } else {
            celle.setFormula(gammeltNavn.getFormula());
        }
        return;

    } else if (arrNavn != gammelVerdi) {

        var streng = Utilities.formatString("Arrangementet på denne raden var tidligere \"%s\". Er \"%s\" det samme arrangementet som \"%s\"?", gammelVerdi, arrNavn, gammelVerdi);
        var response = ui.alert(streng, ui.ButtonSet.YES_NO);

        if (response == ui.Button.YES) {
        //Bare endre navnet, så return
        var regex = /(?<=;").*(?="\))/;
        var nyFormel = gammelFormel.replace(regex, arrNavn);
        celle.setFormula(nyFormel);
        gammeltNavn.setFormula(nyFormel);
        var ark = ss.getSheetByName(gammelVerdi);
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

   //Lenk sammen celler fra hippo, til arrangement-arket

   var dst = [ark.getRange(ARR_NAVN_ARK), ark.getRange(DATO_CELLE), ark.getRange(LOKALE), ark.getRange(ANSVARLIG), ark.getRange(PÅ_JOBB), ark.getRange(COMMENTS)]; 
   var src = [celle, hippo.getRange(row, col-2), hippo.getRange(row, col+1), hippo.getRange(row, col+2), hippo.getRange(row, col+3), hippo.getRange(row, col+4)];  

   //Arrnavn og dato endres under skriptkjøring, og trenger ikke lenking.
   dst[0].setValue(src[0].getDisplayValue());
   dst[1].setValue(dato);

   for (var i = 2; i < dst.length; i++) {
    dst[i].setFormula(Utilities.formatString('=\'%s\'!%s', hippo.getName(), src[i].getA1Notation()));
   };

   var formel = "=HYPERLINK(\"#gid=" + ark.getSheetId() + "\";" + "\"" + celle.getValue() + "\")";

   //referer dato og navn, så den kan nås av rapporteringsskriptet
   gammelDato.setValue(dato);
   gammeltNavn.setFormula(formel);

   //Endre stil på lenken
   celle.setFormula(formel);
   celle.setFontLine('none');
   celle.setFontColor("black");


   // Protect range, then remove all other users from the list of editors.
   // Ensure the current user is an editor before removing others. Otherwise, if the user's edit
   // permission comes from a group, the script throws an exception upon removing the group.
   var protections = ark.getProtections(SpreadsheetApp.ProtectionType.RANGE);

   if (protections.length < 4) {

    beskyttOmråde(ARR_NAVN_ARK, ark);
    beskyttOmråde(DATO_CELLE + ":" + ANSVARLIG, ark);
    beskyttOmråde(PÅ_JOBB, ark);
    beskyttOmråde(COMMENTS, ark);

   }
}

function removeFormatting(cell, name) {
    cell.setFontWeight("normal");
    cell.setValue(name);
}

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
