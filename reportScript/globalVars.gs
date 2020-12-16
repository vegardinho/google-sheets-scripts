// GENERAL VARIABLES
var ui = SpreadsheetApp.getUi();

var ss = SpreadsheetApp.getActiveSpreadsheet();
var sheets = ss.getSheets();

var mainSheet = ss.getSheetByName("Hovedark");
var hippoCopy = ss.getSheetByName("Hippokopi");
var template = ss.getSheetByName("MAL");

const KEEP_SHEETS = ["Hovedark", "Hippokopi", "MAL"];


// MAIN SHEET (RAPPORTERINGSOVERSIKT)
const M_NAME = mainSheet.getName();

const M_DATE_COL_NUM = 3;
const M_EVENT_COL_NUM = 4;

const M_FILM_COL_NUM = 5;
const M_MEMB_COL_NUM = 6;
const M_N_MEMB_COL_NUM = 7;
const M_OTHER_COL_NUM = 8;
const M_FB_PART_COL_NUM = 9;
const M_FB_INT_COL_NUM = 10;
const M_INCOME_COL_NUM = 11;

const M_LAST_COL_NUM = 11;

const M_LAST_RAP_ROW = 63;
const M_FIRST_RAP_ROW = 4;


// COPY OF HIPPO (HIPPOKOPI)
const HC_LINK_CELL_RANGE = "A1";
const HC_LAST_HIP_ROW_CELL = "E1";
const HC_FIRST_HIP_ROW = 2;
const HC_LAST_HIP_ROW = hippoCopy.getRange(HC_LAST_HIP_ROW_CELL).getDisplayValue();

const HC_DATE_COL_NUM = 1;
const HC_EVENT_COL_NUM = 2;
const HC_VENUE_COL_NUM = 3;
const HC_RSPNSBLE_COL_NUM = 4;

var hcLinkCell = hippoCopy.getRange(HC_LINK_CELL_RANGE);
var hcLinkCellFormula = hcLinkCell.getFormula();


/**** TEMPLATE (MAL) ****/
const TMPLT_MAIN_FORMULA_CELL = "A1";
const TMPLT_EV_NAME_CELL = "B2";
const TMPLT_INFO_CELLS = "C6:C8";

const TMPLT_DATE_CELL = "$C$6";
const TMPLT_RSPNSBLE_CELL = "$C$7";
const TMPLT_VENUE_CELL = "$C$8";
const TMPLT_FB_INTRSTD_CELL = "$C$9";
const TMPLT_FB_PRTCPNT_CELL = "$C$10";
const TMPLT_Z_CELL = "$C$12";
const TMPLT_BILLIG_CELL = "$C$13";
const TMPLT_INCOME_CELL = "$C$14";

const TMPLT_MEMB_CELL = "$C$16";
const TMPLT_N_MEMB_CELL = "$C$17";
const TMPLT_OTHER_CELL = "$C$18";
const TMPLT_FILM_CELL = "$C$19";

const TMPLT_TOTAL_CELL = "$C$20";