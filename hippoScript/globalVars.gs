// Date-relevant
var monthToNumber = {
    "JANUAR": 1,
    "FEBRUAR": 2,
    "MARS": 3,
    "APRIL": 4,
    "MAI": 5,
    "JUNI": 6,
    "JULI": 7,
    "AUGUST": 8,
    "SEPTEMBER": 9,
    "OKTOBER": 10,
    "NOVEMBER": 11,
    "DESEMBER": 12,
};

var indexToMonth = {
    0: "JANUAR",
    1: "FEBRUAR",
    2: "MARS",
    3: "APRIL",
    4: "MAI",
    5: "JUNI",
    6: "JULI",
    7: "AUGUST",
    8: "SEPTEMBER",
    9: "OKTOBER",
    10: "NOVEMBER",
    11: "DESEMBER",
};

var weekdays = {
    1: "Man",
    2: "Tirs",
    3: "Ons",
    4: "Tors",
    5: "Fre",
    6: "Lør",
    0: "Søn",
};

const SUNDAY = 0;
const MONDAY = 1;
const THURSDAY = 4;
const DAYS_PER_WEEK = 5;

//Colors
// [lightest, middle, darkest]
var colors = [
    ["#d9ead3", "#b6d7a8", "#93c47d"], 
    ["#cfe2f3", "#9fc5e8", "#6fa8dc"], 
    ["#d5a6bd", "#c27ba0", "#a64d79"], 
    ["#ea9999", "#e06666", "#cc0000"] ];

const NUM_CLRS_WK = 2;
const NUM_CLRS_GRP = colors.length;


/*** General variables ***/
var ss = SpreadsheetApp.getActiveSpreadsheet();
var sheets = ss.getSheets();

var hippo = ss.getSheetByName("Hovedark");
var tmplt = ss.getSheetByName("MAL ARR.");
var copy = ss.getSheetByName("Kopi");
var pr = ss.getSheetByName("PR-plan");
var hotel = ss.getSheetByName("Hotellbooking");

var ui = SpreadsheetApp.getUi();
const CP_LAST_ROW_CELL = "E3";
var mNumRows = copy.getRange(CP_LAST_ROW_CELL).getValue();


/*** COPY-SHEET ***/
const CP_DATE = "A";
const CP_EVENT = "B";
const CP_STRT_DATE_CLL_MAIN = "E4";
const CP_END_DATE_CLL_MAIN = "E5";

const CP_EVENT_COL_NUM = 2;
const CP_DATE_COL_NUM = 1;

const CP_START_ROW = 4;
const CP_LAST_ROW = CP_START_ROW + mNumRows;


//Other
const CAL_ID = 'g5ac3hoiqkfibq8a53s7s6mdlc@group.calendar.google.com';


//Template sheet
const TMPLT_EV_NAME = "B2";
const TMPLT_DATE = "A4";
const TMPLT_VENUE = "B4";
const TMPLT_RSPNSBLE = "C4";
const TMPLT_AT_WORK = "D4";
const TMPLT_COMMENTS = "D24";
const TMPLT_HOTEL = "C14";
const TMPLT_GUESTS = "G11"
const TMPLT_PHOTOS = "A18";
const TMPLT_GRFCS = "B18";
const TMPLT_RLS = "D18";
const TMPLT_RLS_TCKTS = "E18";


/*** HOTEL BOOKING SHEET ***/
const H_DATE_COL_NUM = 1;
const H_EVENT_COL_NUM = 2;
const H_STATUS_COL_NUM = 3;
const H_GUESTS_COL_NUM = 4;
const H_ROOMS_COL_NUM = 5;
const H_COMM_COL_NUM = 6;

const H_START_ROW = 4;
const H_LAST_ROW = H_START_ROW + mNumRows;
const H_LAST_COL_NUM = H_COMM_COL_NUM;





//Main sheet (Hovedark/hippo)
const M_START_ROW = 4;
const M_LAST_ROW = M_START_ROW + mNumRows;

const M_NUM_SH_DFLT = 7;
const M_NUM_COLS = 9;
const M_START_DATE = copy.getRange(CP_STRT_DATE_CLL_MAIN).getDisplayValue();
const M_END_DATE = copy.getRange(CP_END_DATE_CLL_MAIN).getDisplayValue();

const M_MNTH_COL_NUM = 1;
const M_WEEK_COL_NUM = 2;
const M_DATE_COL_NUM = 3;
const M_DAY_COL_NUM = 4;
const M_EVENT_COL_NUM = 5;
const M_VENUE_COL_NUM = 6;
const M_RSPN_COL_NUM = 7;
const M_WRKNG_COL_NUM = 8;
const M_COMM_COL_NUM = 9;


/*** PR SHEET ***/
const PR_DATE_COL_NUM = 1;
const PR_EVENT_COL_NUM = 2;
const PR_GRFCS_COL_NUM = 3;
const PR_PHOTOS_COL_NUM = 4;
const PR_RLS_COL_NUM = 5;
const PR_TCKTS_COL_NUM = 6;
const PR_COMM_COL_NUM = 7;

const PR_LAST_COL_NUM = PR_COMM_COL_NUM;

const PR_START_ROW = 4;
const PR_LAST_ROW = PR_START_ROW + mNumRows;

