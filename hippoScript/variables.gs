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

// [lightest, middle, darkest]
var colors = [
    ["#d9ead3", "#b6d7a8", "#93c47d"], 
    ["#cfe2f3", "#9fc5e8", "#6fa8dc"], 
    ["#d5a6bd", "#c27ba0", "#a64d79"], 
    ["#ea9999", "#e06666", "#cc0000"] ];

const NUM_CLRS_WK = 2;
const NUM_CLRS_GRP = colors.length;


//General variables
var ss = SpreadsheetApp.getActiveSpreadsheet();
var sheets = ss.getSheets();
var hippo = ss.getSheetByName("Hovedark");
var tmplt = ss.getSheetByName("MAL ARR.");
var copy = ss.getSheetByName("Kopi");
var pr = ss.getSheetByName("PR-plan");
var hotel = ss.getSheetByName("Hotellbooking");
var ui = SpreadsheetApp.getUi();

//Mal-celler
const TMPLT_EV_NAME = "B2";
const TMPLT_DATE = "A4";
const TMPLT_VENUE = "B4";
const TMPLT_RSPNSBLE = "C4";
const TMPLT_AT_WORK = "J11";
const TMPLT_COMMENTS = "F16";
const TMPLT_RLS_FB = "D12";
const TMPLT_RLS_TCKTS = "D14";
const TMPLT_GRFCS = "F12";
const TMPLT_HOTEL = "I8";

//PR-ark
const PR_GRFCS = "C";
const PR_FACE = "D";
const PR_KOMM = "E";
const PR_BILL = "F";

//Hotellbooking
const H_DATE = "A";
const H_EVENT = "B";
const H_STATUS = "C";
const H_NUM = "D";
const H_NAMES = "E";

//Copy-sheet
const CP_DATE = "A";
const CP_EVENT = "B";
const CP_STRT_DATE_CLL_MAIN = "E4";
const CP_END_DATE_CLL_MAIN = "E5";

//Main sheet (Hovedark/hippo)
const M_NUM_ROWS = kopi.getRange("E3").getValue();
const M_START_ROW = 4;
const M_NUM_SH_DFLT = 7;
const M_NUM_COLS = 9;
const M_START_DATE = thisSheet.getRange(CP_STRT_DATE_CLL_MAIN).getValue();
const M_END_DATE = thisSheet.getRange(CP_END_DATE_CLL_MAIN).getValue();

const M_MNTH_COL_NUM = 1;
const M_WEEK_COL_NUM = 2;
const M_DATE_COL_NUM = 3;
const M_DAY_COL_NUM = 4;
const M_EVENT_COL_NUM = 5;


/*
Old stuff:
const LAST_ROW_INFO_ROW = 500;
const LAST_ROW_INFO_COL = 20;
const LAST_ROW = hippo.getRange(LAST_ROW_INFO_ROW, LAST_ROW_INFO_COL).getValue();
*/