// LIST OF COLUMNS IN SHEET_NAME
const TIMESTAMP_COL = 1;
const EMAIL_COL = 2;
const HEADRUNNERS_COL = 3;
const HEADRUN_COL = 4;
const RUN_LEVEL_COL = 5;
const ATTENDEES_BEGINNER_COL = 6;
const ATTENDEES_INTERMEDIATE_COL = 7;
const ATTENDEES_ADVANCED_COL = 8;
const CONFIRMATION_COL = 9;
const DISTANCE_COL = 10;
const COMMENTS_COL = 11;
const TRANSFER_STATUS_COL = 12;
const PLATFORM_COL = 13;
const NAMES_NOT_FOUND_COL = 14;


const ATTENDANCE_SHEET_NAME = 'HR Attendance W25';
const ATTENDANCE_SS_ID = '1SnaD9UO4idXXb07X8EakiItOOORw5UuEOg0dX_an3T4';
const ATTENDANCE_SHEET = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ATTENDANCE_SHEET_NAME);

// ALLOWS PROPER SHEET REF WHEN ACCESSING AS LIBRARY FROM EXTERNAL SCRIPT
// SpreadsheetApp.getActiveSpreadsheet() DOES NOT WORK IN EXTERNAL SCRIPT
const GET_ATTENDANCE_SHEET_ = () => {
  return (ATTENDANCE_SHEET) ?? SpreadsheetApp.openById(ATTENDANCE_SS_ID).getSheetByName(ATTENDANCE_SHEET_NAME);
}

const TIMEZONE = getUserTimeZone_();

// RUN LEVELS
const ATTENDEE_MAP = {
  'beginner': ATTENDEES_BEGINNER_COL,
  //'easy': ATTENDEES_BEGINNER_COL,
  'intermediate': ATTENDEES_INTERMEDIATE_COL,
  'advanced': ATTENDEES_ADVANCED_COL,
};

const LEVEL_COUNT = Object.keys(ATTENDEE_MAP).length;
const EMPTY_ATTENDEE_FLAG = 'None';

const MEMBER_EMAIL_COL = 1;   // Found in 'Members' sheet
const MEMBER_SEARCH_KEY_COL = 6;  // Found in 'Members' sheet

// EXTERNAL SHEETS USED IN SCRIPTS
const MASTER_NAME = 'MASTER';
const SEMESTER_NAME = 'Winter 2025';
const MEMBERSHIP_URL = "https://docs.google.com/spreadsheets/d/1qvoL3mJXCvj3m7Y70sI-FAktCiSWqEmkDxfZWz0lFu4/";

// LEDGER SPREADSHEET
const LOG_SHEET_NAME = 'Event Log';
const POINTS_LEDGER_URL = "https://docs.google.com/spreadsheets/d/13ps2HsOz-ZLg8xc0RYhKl7eg3BOs1MYVrwS0jxP3FTc/";

// SCRIPT PROPERTIES; MAKE SURE THAT NAMES MATCHES BANK
const SCRIPT_PROPERTY = {
  isCheckingAttendance: 'IS_CHECKING_ATTENDANCE',
  calendarTriggers: 'calendarTriggers',
};

// NAME OF HTML TEMPLATES. ENSURE CORRECT FILE NAME!!
const COPY_EMAIL_HTML_FILE = 'Confirmation-Email';
const REMINDER_EMAIL_HTML_FILE = 'Reminder-Email';

const GET_ATTENDANCE_GFORM_LINK_ = () => ATTENDANCE_SHEET.getFormUrl();


/** GET HEADRUN SCHEDULE AND HEADRUNNER INFO */
const GET_HEADRUN_SS = () => SpreadsheetApp.openById('1FvnfUHO2Xj-Q5j6-B2Wz9DgxEXFX-OoCh6Pf49Kqyac');

const COMPILED_SHEET_NAME = "Compiled";
const GET_COMPILED_SHEET_ = () => GET_HEADRUN_SS().getSheetByName(COMPILED_SHEET_NAME);

const HEADRUNNER_SHEET_NAME = "List of Head Runners";
const GET_HEADRUNNER_SHEET_ = () => GET_HEADRUN_SS().getSheetByName(HEADRUNNER_SHEET_NAME);


/** GET PROP STORE OR GENERATE IF NULL  */
let PROP_STORE = null;
const GET_PROP_STORE_ = () => {
  return PROP_STORE ?? PropertiesService.getScriptProperties();
}


/**
 * Returns timezone for currently running script.
 *
 * Prevents incorrect time formatting during time changes like Daylight Savings Time.
 *
 * @return {string}  Timezone as geographical location (e.g.`'America/Montreal'`).
 */

function getUserTimeZone_() {
  return Session.getScriptTimeZone();
}


/**
 * Returns email of current user executing Google Apps Script functions.
 *
 * Prevents incorrect account executing Google automations (e.g. McRUN bot.)
 *
 * @return {string}  Email of current user.
 */

function getCurrentUserEmail_() {
  return Session.getActiveUser().toString();
}


/**
 * Registers column positions from `ATTENDANCE_SHEET`.
 *
 * Prevents user from updating column variables manually.
 *
 * CURRENTLY IN REVIEW!
 *
 * @author [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>)
 * @date  Oct 23, 2024
 * @update  Oct 23, 2024
 */

function getColumnPosition() {
  var rangeList = ATTENDANCE_SHEET.getNamedRanges();
  var dRange = ATTENDANCE_SHEET.getNamedRanges()[0].getRange();

  for (var i = 0; i < rangeList.length; i++) {
    Logger.log(rangeList[i].getName());
  }
}
