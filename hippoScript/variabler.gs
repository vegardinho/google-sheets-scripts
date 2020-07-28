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

var ukedager = {
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
const DAYS_WEEK = 5;

// [lyseste, middels, mørkeste]
var farger = [
    ["#d9ead3", "#b6d7a8", "#93c47d"], 
    ["#cfe2f3", "#9fc5e8", "#6fa8dc"], 
    ["#d5a6bd", "#c27ba0", "#a64d79"], 
    ["#ea9999", "#e06666", "#cc0000"] ];

const ANT_FARGE_WEEK = 2;
const ANT_FARGE_GRP = farger.length;


//General variables
var ss = SpreadsheetApp.getActiveSpreadsheet();
var sheets = ss.getSheets();
var hippo = ss.getSheetByName("Hovedark");
var mal = ss.getSheetByName("MAL ARR.");
var kopi = ss.getSheetByName("Kopi");
var pr = ss.getSheetByName("PR-plan");
var h_sheet = ss.getSheetByName("Hotellbooking");
var ui = SpreadsheetApp.getUi();

//Mal-celler
const ARR_NAVN_ARK = "B2";
const DATO_CELLE = "A4";
const LOKALE = "B4";
const ANSVARLIG = "C4";
const PÅ_JOBB = "J11";
const COMMENTS = "F16";
const FACE_SLIPP = "D12";
const BILL_SLIPP = "D14";
const GRAFIKK = "F12";
const STAT_HOTEL = "I8";

//PR-ark
const PR_GRAF = "C";
const PR_FACE = "D";
const PR_KOMM = "E";
const PR_BILL = "F";

//Hotellbooking
const H_DATE = "A";
const H_EVENT = "B";
const H_STATUS = "C";
const H_NUM = "D";
const H_NAMES = "E";

//Main sheet (Hovedark/hippo)
const ANT_RADER = kopi.getRange("E3").getValue();
const START_ROW = 4;
const STATIC_SHEETS = 5;
const NUM_COLS = 9;
const START_DATE_CELL = "E4";
const END_DATE_CELL = "E5";

const MONTH_COLUMN = 1;
const WEEK_COLUMN = 2;
const DATE_COLUMN = 3;
const DAY_COLUMN = 4;
const ARR_COLUMN = 5;


//Copy-sheet
const CP_DATE = "A";
const CP_EVENT = "B";

/*
const LAST_ROW_INFO_ROW = 500;
const LAST_ROW_INFO_COL = 20;
const LAST_ROW = hippo.getRange(LAST_ROW_INFO_ROW, LAST_ROW_INFO_COL).getValue();
*/