/*
Copyright 2025 Andrey Gonzalez (for McGill Students Running Club)

Licensed under the Apache License, Version 2.0 (the "License");
you may not use this file except in compliance with the License.
You may obtain a copy of the License at

    https://www.apache.org/licenses/LICENSE-2.0

Unless required by applicable law or agreed to in writing, software
distributed under the License is distributed on an "AS IS" BASIS,
WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
See the License for the specific language governing permissions and
limitations under the License.
*/

/** 
 * Name of attendance sheet for the current semester
 * TO UPDATE EACH SEMESTER 
 */
const ATTENDANCE_SHEET_NAME = 'HR Attendance F25';

/**
 * Name of semester
 * TO UPDATE EACH SEMESTER 
 */
const SEMESTER_NAME = 'Fall 2025';

/**
 * ID of attendance sheet for the current semester
 * TO UPDATE EACH SEMESTER 
 */
const ATTENDANCE_SHEET_ID = '1kUevgOCN1wCdbNiVY412-7ejnlSjtIyKNHFVLV9KK1Q';

/**
 * Attendance sheet object for the current semester
 * TO UPDATE EACH SEMESTER 
 */
const ATTENDANCE_SHEET = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ATTENDANCE_SHEET_NAME);

/**
 * Retrieves the attendance sheet for the current semester.
 * Ensures proper sheet reference when accessing as a library from an external script.
 *
 * @return {GoogleAppsScript.Spreadsheet.Sheet} The attendance sheet object.
 */
const GET_ATTENDANCE_SHEET_ = () => {
  return (ATTENDANCE_SHEET) ?? SpreadsheetApp.openById(ATTENDANCE_SHEET_ID).getSheetByName(ATTENDANCE_SHEET_NAME);
}

/**
 * Mapping of column letters to numbers
 */
const COL = {
  A: 1,
  B: 2,
  C: 3,
  D: 4,
  E: 5,
  F: 6,
  G: 7,
  H: 8,
  I: 9,
  J: 10,
  K: 11,
  L: 12,
  M: 13,
  N: 14,
  O: 15,
  P: 16,
  Q: 17,
  R: 18,
  S: 19,
  T: 20,
  U: 21,
  V: 22,
  W: 23,
  X: 24,
  Y: 25,
  Z: 26
}

/**
 * Mapping of columns in semester attendance sheet to column number (1-indexed)
 */
const SEM_ATTENDANCE_COLS = {
  TIMESTAMP: COL.A,
  EMAIL: COL.B,
  HEADRUNNERS: COL.C,
  HEADRUN: COL.D,
  RUN_LEVEL: COL.E,
  B_ATTENDEES: COL.F, // Beginner
  E_ATTENDEES: COL.G, // Easy
  I_ATTENDEES: COL.H, // Intermediate
  A_ATTENDEES: COL.I, // Advanced
  CONFIRMATION: COL.J,
  DISTANCE: COL.K,
  COMMENTS: COL.L,
  TRANSFER_STATUS: COL.M,
  PLATFORM: COL.N,
  NOT_FOUND: COL.O,
}

/**
 * Timezone of the script
 */
const TIMEZONE = getUserTimeZone_();

/**
 * Maps run levels to column with attendees for that level
 */
const ATTENDEE_MAP = {
  'beginner': SEM_ATTENDANCE_COLS.B_ATTENDEES,
  'easy': SEM_ATTENDANCE_COLS.E_ATTENDEES,
  'intermediate': SEM_ATTENDANCE_COLS.I_ATTENDEES,
  'advanced': SEM_ATTENDANCE_COLS.A_ATTENDEES,
};

/**
 * Number of run levels
 */
const NUM_LEVELS = Object.keys(ATTENDEE_MAP).length;

/**
 * String indicating that there are no attendees for a run level
 */
const EMPTY_ATTENDEE_FLAG = 'None';

// EXTERNAL SHEETS USED IN SCRIPTS
/** 
 * Name of sheet (in membership spreadsheet) with master registry
 */
const MEMBERSHIP_SHEET_NAME = 'MASTER';

/**
 * URL of membership spreadsheet
 */
const MEMBERSHIP_URL = "https://docs.google.com/spreadsheets/d/1qvoL3mJXCvj3m7Y70sI-FAktCiSWqEmkDxfZWz0lFu4/";

/**
 * Column in membership list with member email
 */
const MEMBER_EMAIL_COL = COL.A;

/**
 * Column in membership list with member key (ID?)
 */
const MEMBER_SEARCH_KEY_COL = COL.V;

/** 
 * Name of sheet with events log in points ledger spreadsheet
 */
const LOG_SHEET_NAME = 'Event Log';

/**
 * URL of points ledger spreadsheet
 */
const POINTS_LEDGER_URL = "https://docs.google.com/spreadsheets/d/1sar-Pmfb_Nar0Lc9u8-rXyllLvQMqBFlSwolCoHX-_4/";

/** 
 * ID of spreadsheet with head run schedule and list of head runners
 */
const HEADRUN_SHEET_ID = '14uhswSruvsR2TT94CPYsPfCPfuvUmHeYZ4xqZ1IbWTg';

/**
 * Name of sheet with compiled head runs and head runners
 */
const COMPILED_SHEET_NAME = "Compiled";

/**
 * Sheet object with compiled head runs and head runners
 */
const GET_COMPILED_SHEET_ = () => SpreadsheetApp.openById(HEADRUN_SHEET_ID).getSheetByName(COMPILED_SHEET_NAME);

/**
 * Name of sheet with list of head runners and their info
 */
const HEADRUNNER_SHEET_NAME = "List of Head Runners";

/**
 * Sheet object with list of head runners and their info
 */
const GET_HEADRUNNER_SHEET_ = () => SpreadsheetApp.openById(HEADRUN_SHEET_ID).getSheetByName(HEADRUNNER_SHEET_NAME);

/**
 * Maps information to script property name with that information
 */
const SCRIPT_PROPERTY = {
  isCheckingAttendance: 'IS_CHECKING_ATTENDANCE',
  calendarTriggers: 'calendarTriggers',
  webAppId: 'WEB_APP_ID',   // ⚠️⚠️ UPDATE WEB_APP_ID FOR NEW 'POINTS LEDGER CODE' DEPLOYMENTS ⚠️⚠️
  webAppKey: 'WEB_APP_KEY',
};

/**
 * Script properties
 */
let PROP_STORE = null;

/** 
 * Get property store or create if not found
 */
const GET_PROP_STORE_ = () => {
  return PROP_STORE ?? PropertiesService.getScriptProperties();
}

/**
 * Name of email template used to send attendance copy (without '.html')
 */
const COPY_EMAIL_TEMPLATE = 'Copy-Email';

/**
 * Name of email template used to send reminder (without '.html')
 */
const REMINDER_EMAIL_TEMPLATE = 'Reminder-Email';

/**
 * Gets link of Google Form connected to attendance sheet
 * 
 * @return {string}  Link to attendance form
 */
const GET_ATTENDANCE_FORM_LINK_ = () => ATTENDANCE_SHEET.getFormUrl();

/**
 * Email of club president
 */
const PRESIDENT_EMAIL = 'alexis.demetriou@mail.mcgill.ca';

/**
 * Email of club VP Internal (VP Headruns)
 */
const VP_INTERNAL_EMAIL = 'nicolas.morrison@mail.mcgill.ca';

/**
 * Club email address
 */
const CLUB_EMAIL = 'mcrunningclub@ssmu.ca';

/**
 * Email address used to access the attendance app
 */
const APP_EMAIL = 'mcgillstudentsrunningclub@gmail.com';

/**
 * Name of script property that has headrunner info
 */
const HEADRUNNER_STORE_NAME = 'headrunners';

/**
 * Name of script property that has headrun info/schedule
 */
const HEADRUN_STORE_NAME = 'headruns';

/**
 * ID of sheet with app imports
 */
const IMPORT_SHEET_ID = 82376152;

/**
 * Sheet object with app imports
 */
const IMPORT_SHEET = SpreadsheetApp.getActiveSpreadsheet().getSheetById(IMPORT_SHEET_ID);

/**
 * ALLOWS PROPER SHEET REF WHEN ACCESSING AS LIBRARY FROM EXTERNAL SCRIPT
 * SpreadsheetApp.getActiveSpreadsheet() DOES NOT WORK IN EXTERNAL SCRIPT
 */
const GET_IMPORT_SHEET_ = () => {
  return (IMPORT_SHEET) ?? SpreadsheetApp.openById(ATTENDANCE_SHEET_ID).getSheetById(IMPORT_SHEET_ID);
}

const IS_IMPORTED_COL = 10;   // Update if columns modified

/**
 * MAPPING FROM MASTER ATTENDANCE SHEET TO SEMESTER SHEET
 */
const IMPORT_MAP = {
  'timestamp': SEM_ATTENDANCE_COLS.TIMESTAMP,
  'headrunners': SEM_ATTENDANCE_COLS.HEADRUNNERS,
  'headRun': SEM_ATTENDANCE_COLS.HEADRUN,
  'runLevel': SEM_ATTENDANCE_COLS.RUN_LEVEL,
  'confirmation': SEM_ATTENDANCE_COLS.CONFIRMATION,
  'distance': SEM_ATTENDANCE_COLS.DISTANCE,
  'comments': SEM_ATTENDANCE_COLS.COMMENTS,
  'platform': SEM_ATTENDANCE_COLS.PLATFORM,
  'attendees': SEM_ATTENDANCE_COLS.B_ATTENDEES,
}

const CALENDAR_STORE = SCRIPT_PROPERTY.calendarTriggers;
const TRIGGER_FUNC = checkAttendanceSubmission.name;
const TRIGGER_BASE_ID = 'attendanceTrigger';
const TRIGGER_OFFSET = 60 * 60 * 1000;  // 1 hour in ms


/**
 * Users authorized to use the McRUN menu.
 *
 * Prevents unwanted data overwrite in Gsheet.
 *
 * @constant {string[]} PERM_USER_ - List of authorized user emails.
 */
const PERM_USER_ = [
  CLUB_EMAIL,
  'ademetriou8@gmail.com',
  'andreysebastian10.g@gmail.com',
  'monaliu832@gmail.com'
  // ADD NEW TECH MEMBERS!!
];