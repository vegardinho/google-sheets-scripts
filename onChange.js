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
      fiksLenke(ss.getActiveCell());
   }
}


function fiksLenke(valgtCelle) {

   var celle = valgtCelle;
   var arrNavn = celle.getDisplayValue();
   var formel = celle.getFormula();
   var tekstStil = celle.getRichTextValue().getTextStyle();
   var col = celle.getColumn();
   var row = celle.getRow();
   var gammeltNavn = kopi.getRange(row, 2);
   var gammelDato = kopi.getRange(row, 1);
   var gammelVerdi = gammeltNavn.getDisplayValue();
   var gammelFormel = gammeltNavn.getFormula();

   const ARR_NAVN_ARK = "B2";
   const DATO_CELLE = "A4";
   const LOKALE = "B4";
   const ANSVARLIG = "C4";
   const PÅ_JOBB = "J11";
   const COMMENTS = "F16";

   //  Return null hvis ikke riktig kolonne eller ikke bold
   if ((celle.getColumn() !== 5) || !tekstStil.isBold() || (arrNavn === "" && gammelVerdi === "")) {
      return;
   }

   if (gammelVerdi != "") {
      if (arrNavn === "") {
	 var result = ui.alert(
	    'Bekreft sletting',
	    'Er du sikker på at du vil slette arrangementet? ' +
	    'Hvis du etter sletting ønsker å bruke regnearket tilknyttet dette arrangementet senere, kan du taste inn navnet \"' + gammelVerdi + '\".',
	    ui.ButtonSet.YES_NO);
	 if (result == ui.Button.YES) {
	    gammeltNavn.clearContent();
	    hippo.getRange(row, col, 1, 5).clearContent();
	    celle.setFontWeight("normal");
	 } else {
	    celle.setFormula(gammeltNavn.getFormula());
	 }

	 return;

      } else if (arrNavn != gammelVerdi) {

	 var streng = Utilities.formatString("Er \"%s\" det samme arrangementet som \"%s\"?", gammelVerdi, arrNavn);
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

	    return;

	 } else {
	    gammeltNavn.clearContent();
	    ui.alert('Dersom du senere ønsker å bruke regnearket tilknyttet \"' + gammelVerdi + '\", taster du inn dette navnet for valgt datocelle ' +
	       'for å gjenopprette arrangementsdataene.');
	 }
      }
   }

   // Legg til hvis ikke eksisterende
   if (ss.getSheetByName(arrNavn) === null) {

      var nyttArk = ss.insertSheet(arrNavn, ss.getSheets().length, {template: mal});
      var hippoLenke = "=HYPERLINK(\"#gid=" + hippo.getSheetId() + "\";" + "\"" + "TILBAKE TIL HIPPO" + "\")";
      nyttArk.getRange("A1").setFormula(hippoLenke);

   }

   ss.setActiveSheet(hippo);
   hippo.setActiveSelection(celle);

   //Lenk sammen celler fra hippo, til arrangement-arket
   var ark = ss.getSheetByName(arrNavn);
   var dato = lagDato(row);

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
   var måned = hippo.getRange(row, col-4).getMergedRanges()[0].getDisplayValue();
   var tallDag = hippo.getRange(row, col-2);

   if (tallDag.isPartOfMerge()) {
      tallDag = tallDag.getMergedRanges()[0].getDisplayValue();
   } else {
      tallDag = tallDag.getDisplayValue();
   }

   return tallDag + "." + monthToNumber[måned].toString();
}

