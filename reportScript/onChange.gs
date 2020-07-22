// Slett gamle instanser av trigger, og lag ny
function createSpreadsheetOnChangeTrigger() {
    ScriptApp.getProjectTriggers().slice()
        .forEach (function (d) {
                if (d.getHandlerFunction() === 'fiksEnLenke') ScriptApp.deleteTrigger(d);
                });

    var ss = SpreadsheetApp.getActive();
    ScriptApp.newTrigger('fiksEnLenke').forSpreadsheet(ss).onChange().create();
}


function fiksEnLenke() {
    fiksLenke(SpreadsheetApp.getActiveSpreadsheet().getActiveCell()); 
}

function fiksLenke(valgtCelle, lokale, ansvarlig) {

    var celle = valgtCelle;

    //  Return null hvis ikke riktig kolonne eller tom
    if ((celle.getColumn() !== 4) || celle.getDisplayValue() === "") {
        return;
    }

    var arrNavn = celle.getDisplayValue();
    var formel = celle.getFormula();
    var tekstStil = celle.getRichTextValue().getTextStyle();
    var ark = ss.getSheetByName(arrNavn);

    // Legg til hvis ikke eksisterende
    if (ark === null) {
        var nyttArk = ss.insertSheet(arrNavn, ss.getSheets().length, {template: mal});
        var hovedarkLenke = "=HYPERLINK(\"#gid=" + hovedark.getSheetId() + "\";" + "\"" + "TILBAKE TIL HOVEDARK" + "\")";
        nyttArk.getRange("A1").setFormula(hovedarkLenke);
        ark = nyttArk;
    }

    ss.setActiveSheet(hovedark);
    hovedark.setActiveSelection(celle);

    //Lenk sammen celler fra mal til hovedark (omvendt fra hippo-skript)
    var col = celle.getColumn();
    var row = celle.getRow();

    var dato = ark.getRange("$C$6");
    var ansv = ark.getRange("$C$7");
    var loka = ark.getRange("$C$8");
    var fb_del = ark.getRange("$C$9");
    var fb_int = ark.getRange("$C$10");
    var avst = ark.getRange("$C$11");
    var memb = ark.getRange("$C$14");
    var n_memb = ark.getRange("$C$15");
    var free = ark.getRange("$C$16");
    var film = ark.getRange("$C$17");

    var src = [dato, fb_del, fb_int, avst, memb, n_memb, free, film]; 
    var dst = [hovedark.getRange(row, col-1), hovedark.getRange(row, col+5), hovedark.getRange(row, col+6), 
        hovedark.getRange(row, col+7), hovedark.getRange(row, col+2), hovedark.getRange(row, col+3), 
        hovedark.getRange(row, col+4), hovedark.getRange(row, col+1)];  

    //Arrnavn endres under skriptkjøring, og trenger ikke lenking.
    ark.getRange("B2").getMergedRanges()[0].setValue(arrNavn);

    //Dato lenkes fra hovedark
    var a1 = dst[0].getA1Notation();
    dato.setFormula(Utilities.formatString('=\'%s\'!$%s$%s', hovedark.getName(), a1.substring(0,1), a1.substring(1)));

    //Lenker til ansvarlig og lokale hvis parametere er med
    if (ansvarlig != undefined) {
        a1 = ansvarlig.getA1Notation();
        ansv.setFormula(Utilities.formatString('=\'%s\'!$%s$%s', hippo.getName(), a1.substring(0,1), a1.substring(1)));
    }
    if (ansvarlig != undefined) {
        a1 = lokale.getA1Notation();
        loka.setFormula(Utilities.formatString('=\'%s\'!$%s$%s', hippo.getName(), a1.substring(0,1), a1.substring(1)));
    }

    for (var i = 1; i < dst.length; i++) {
        a1 = src[i].getA1Notation();
        dst[i].setFormula(Utilities.formatString('=\'%s\'!$%s$%s', ark.getName(), a1.substring(0,1), a1.substring(1)));
    };

    var formel = "=HYPERLINK(\"#gid=" + ark.getSheetId() + "\";" + "\"" + celle.getValue() + "\")";
    celle.setFormula(formel);
    celle.setFontLine('underline');
    celle.setFontColor('#1155cc');

    // Protect range, then remove all other users from the list of editors.
    // Ensure the current user is an editor before removing others. Otherwise, if the user's edit
    // permission comes from a group, the script throws an exception upon removing the group.
    var protections = ark.getProtections(SpreadsheetApp.ProtectionType.RANGE);

    if (protections.length === 0) {

        beskyttOmråde("B2", ark);
        beskyttOmråde("C6:C8", ark);
    }
}

function beskyttOmråde(områdeStreng, ark) {
    var me = hovedark.getProtections(SpreadsheetApp.ProtectionType.SHEET)[0].getEditors()[0];  
    var protection = ark.getRange(områdeStreng).protect().setDescription('Ny beskyttelse');
    protection.addEditor(me);
    protection.removeEditors(protection.getEditors());

    if (protection.canDomainEdit()) {
        protection.setDomainEdit(false);
    }
}
