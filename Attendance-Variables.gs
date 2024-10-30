// SHEET CONSTANTS
const SHEET_NAME = 'HR Attendance F24';
const ATTENDANCE_SHEET = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);

// LIST OF COLUMNS IN SHEET_NAME
const TIMESTAMP_COL = 1;
const EMAIL_COL = 2;
const HEADRUNNERS = 3;
const HEADRUN_COL = 4;
const RUN_LEVEL_COL = 5;
const ATTENDEES_BEGINNER_COL = 6;
const ATTENDEES_INTERMEDIATE_COL = 7;
const ATTENDEES_ADVANCED_COL = 8;
const CONFIRMATION_COL = 9;
const DISTANCE_COL = 10;
const COMMENTS_COL = 11;
const IS_COPY_SENT_COL = 12;
const PLATFORM_COL = 13;

const TIMEZONE = getUserTimeZone_();
const LEVEL_COUNT = 3;  // Beginner/Easy, Intermediate, Hard

// EXTERNAL SHEETS USED IN SCRIPTS
const MASTER_NAME = 'MASTER';
const SEMESTER_NAME = 'Fall 2024';
const MEMBERSHIP_URL = "https://docs.google.com/spreadsheets/d/1qvoL3mJXCvj3m7Y70sI-FAktCiSWqEmkDxfZWz0lFu4/edit?usp=sharing";

const LEDGER_NAME = 'Head Run Attendance';
const LEDGER_URL = "https://docs.google.com/spreadsheets/d/13ps2HsOz-ZLg8xc0RYhKl7eg3BOs1MYVrwS0jxP3FTc/edit?usp=sharing";


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
  return Session.getActiveUser();
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

  for (var i=0; i < rangeList.length; i++){
    Logger.log(rangeList[i].getName());
  }
}


