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

// [lyseste, middels, mørkeste]
var farger = [
    ["#d9ead3", "#b6d7a8", "#93c47d"], 
    ["#cfe2f3", "#9fc5e8", "#6fa8dc"], 
    ["#d5a6bd", "#c27ba0", "#a64d79"], 
    ["#ea9999", "#e06666", "#cc0000"] ];

var ss = SpreadsheetApp.getActiveSpreadsheet();
var sheets = ss.getSheets();
var hippo = ss.getSheetByName("Hovedark");;
var mal = ss.getSheetByName("MAL ARR.");
var kopi = ss.getSheetByName("Kopi");
var ui = SpreadsheetApp.getUi();

const ANT_RADER = 150;
const START_ROW = 4;
const STATIC_SHEETS = 5;
const NUM_COLS = 9;

const MONTH_COLUMN = 1;
const WEEK_COLUMN = 2;
const DATE_COLUMN = 3;
const DAY_COLUMN = 4;
const ARR_COLUMN = 5;

const LAST_ROW_INFO_ROW = 500;
const LAST_ROW_INFO_COL = 20;
const LAST_ROW = hippo.getRange(LAST_ROW_INFO_ROW, LAST_ROW_INFO_COL).getValue();

const SUNDAY = 0;
const MONDAY = 1;
const THURSDAY = 4;
const DAYS_WEEK = 5;

const ANT_FARGE_WEEK = 2;
const ANT_FARGE_GRP = farger.length;
