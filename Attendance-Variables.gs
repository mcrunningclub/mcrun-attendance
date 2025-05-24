// LIST OF COLUMNS IN SHEET_NAME
const TIMESTAMP_COL = 1;
const EMAIL_COL = 2;
const HEADRUNNERS_COL = 3;
const HEADRUN_COL = 4;
const RUN_LEVEL_COL = 5;
const ATTENDEES_BEGINNER_COL = 6;
const ATTENDEES_EASY_COL = 7;
const ATTENDEES_INTERMEDIATE_COL = 8;
const ATTENDEES_ADVANCED_COL = 9;
const CONFIRMATION_COL = 10;
const DISTANCE_COL = 11;
const COMMENTS_COL = 12;
const TRANSFER_STATUS_COL = 13;
const PLATFORM_COL = 14;
const NAMES_NOT_FOUND_COL = 15;


/** TO UPDATE EACH SEMESTER */
const ATTENDANCE_SHEET_NAME = 'HR Attendance S25';
const SEMESTER_NAME = 'Summer 2025';

const ATTENDANCE_SS_ID = '1SnaD9UO4idXXb07X8EakiItOOORw5UuEOg0dX_an3T4';
const ATTENDANCE_SHEET = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ATTENDANCE_SHEET_NAME);

/**
 * Retrieves the attendance sheet for the current semester.
 * Ensures proper sheet reference when accessing as a library from an external script.
 *
 * @return {GoogleAppsScript.Spreadsheet.Sheet} The attendance sheet object.
 */
const GET_ATTENDANCE_SHEET_ = () => {
  return (ATTENDANCE_SHEET) ?? SpreadsheetApp.openById(ATTENDANCE_SS_ID).getSheetByName(ATTENDANCE_SHEET_NAME);
}

const TIMEZONE = getUserTimeZone_();

// RUN LEVELS
const ATTENDEE_MAP = {
  'beginner': ATTENDEES_BEGINNER_COL,
  'easy': ATTENDEES_EASY_COL,
  'intermediate': ATTENDEES_INTERMEDIATE_COL,
  'advanced': ATTENDEES_ADVANCED_COL,
};

const LEVEL_COUNT = Object.keys(ATTENDEE_MAP).length;
const EMPTY_ATTENDEE_FLAG = 'None';

const MEMBER_EMAIL_COL = 1;   // Found in 'Members' sheet
const MEMBER_SEARCH_KEY_COL = 6;  // Found in 'Members' sheet

// EXTERNAL SHEETS USED IN SCRIPTS
const MASTER_NAME = 'MASTER';
const MEMBERSHIP_URL = "https://docs.google.com/spreadsheets/d/1qvoL3mJXCvj3m7Y70sI-FAktCiSWqEmkDxfZWz0lFu4/";

// LEDGER SPREADSHEET
const LOG_SHEET_NAME = 'Event Log';
const POINTS_LEDGER_URL = "https://docs.google.com/spreadsheets/d/1DwmnZgLftSqegfsoFA5fekuT0sosgCntVMmTylbj8o4/";

// SCRIPT PROPERTIES; MAKE SURE THAT NAMES MATCHES BANK
const SCRIPT_PROPERTY = {
  isCheckingAttendance: 'IS_CHECKING_ATTENDANCE',
  calendarTriggers: 'calendarTriggers',
};


// NAME OF HTML TEMPLATES. ENSURE CORRECT FILE NAME!!
const COPY_EMAIL_HTML_FILE = 'Copy-Email';
const REMINDER_EMAIL_HTML_FILE = 'Reminder-Email';

const GET_ATTENDANCE_GFORM_LINK_ = () => ATTENDANCE_SHEET.getFormUrl();


/** GET HEADRUN SCHEDULE AND HEADRUNNER INFO */
const GET_HEADRUN_SS_ = () => SpreadsheetApp.openById('1Hx4R4gkMjQ71Jj1oeaxS6G6uMvoHzSKkhBLVVi3yHBs');

const COMPILED_SHEET_NAME = "Compiled";
const GET_COMPILED_SHEET_ = () => GET_HEADRUN_SS_().getSheetByName(COMPILED_SHEET_NAME);

const HEADRUNNER_SHEET_NAME = "List of Head Runners";
const GET_HEADRUNNER_SHEET_ = () => GET_HEADRUN_SS_().getSheetByName(HEADRUNNER_SHEET_NAME);


/** GET PROP STORE OR GENERATE IF NULL  */
let PROP_STORE = null;
const GET_PROP_STORE_ = () => {
  return PROP_STORE ?? PropertiesService.getScriptProperties();
}


/**
 * Returns the attendance sheet name for the current semester.
 *
 * @return {string} The name of the attendance sheet.
 * @customfunction
 */

function GET_SEMESTER_SHEET_NAME() {
  return ATTENDANCE_SHEET_NAME;
}

/**
 * Returns timezone for currently running script.
 *
 * Prevents incorrect time formatting during time changes like Daylight Savings Time.
 *
 * @return {string} Timezone as a geographical location (e.g., `'America/Montreal'`).
 */

function getUserTimeZone_() {
  return Session.getScriptTimeZone();
}


/**
 * Returns the email of the current user executing Google Apps Script functions.
 * Useful for ensuring the correct account is executing Google automations.
 *
 * @return {string} Email of the current user.
 */

function getCurrentUserEmail_() {
  return Session.getActiveUser().toString();
}


/**
 * Sorts the `ATTENDANCE_SHEET` by submission time.
 * Excludes the header row from sorting.
 *
 * @author [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>)
 * @date  Nov 1, 2024
 * @update  Apr 7, 2025
 */

function sortAttendanceForm() {
  const sheet = GET_ATTENDANCE_SHEET_();

  const numRows = sheet.getLastRow() - 1;  // Remove header row from count
  const numCols = sheet.getLastColumn();
  const range = sheet.getRange(2, 1, numRows, numCols);

  // Sorts values by `Timestamp` without the header row
  range.sort([{ column: 1, ascending: true }]);
}


/**
 * Changes the attendance status of all members to "not present."
 * Helper function for `consolidateMemberData()`.
 *
 * @trigger New head run or McRUN attendance submission.
 *
 * @author [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>)
 * @date  Oct 9, 2023
 * @update  Oct 29, 2024
 */

function removePresenceChecks() {
  const sheetURL = MEMBERSHIP_URL;
  const ss = SpreadsheetApp.openByUrl(sheetURL);

  const masterSheetName = MASTER_NAME;
  const sheet = ss.getSheetByName(masterSheetName);

  let rangeAttendance;
  const rangeList = sheet.getNamedRanges();

  for (let i = 0; i < rangeList.length; i++) {
    if (rangeList[i].getName() === "attendanceStatus") {
      rangeAttendance = rangeList[i];
      break;
    }
  }

  rangeAttendance.getRange().uncheck(); // Remove all Presence checks
}

/**
 * Formats specific columns of the `HR Attendance` sheet for better readability.
 * Includes freezing panes, bold formatting, text wrapping, alignment, and column resizing.
 *
 * @trigger New Google form or app submission.
 *
 * @author [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>)
 * @date  Oct 9, 2023
 * @update  May 13, 2025
 */

function formatSpecificColumns_() {
  const sheet = GET_ATTENDANCE_SHEET_();

  // Helper function to improve readability
  const getThisRange = (ranges) =>
    Array.isArray(ranges) ? sheet.getRangeList(ranges) : sheet.getRange(ranges);

  // 1. Freeze panes
  sheet.setFrozenRows(1);
  sheet.setFrozenColumns(1);

  // 2. Bold formatting
  getThisRange([
    'A1:O1',  // Header Row
    'A2:A',   // Timestamp
    'D2:D',   // Headrun
    'M2:O'    // Transfer Status + ... + Not Found
  ]).setFontWeight('bold');

  // 3. Font size adjustments
  getThisRange(['A2:A', 'D2:D', 'N2:N']).setFontSize(11); // for Headrun + Submission Platform
  getThisRange(['C2:C', 'F2:I']).setFontSize(9);  // Headrunners + Attendees

  // 4. Text wrapping
  getThisRange(['B2:E', 'J2:L']).setWrap(true);
  getThisRange('F2:I').setWrap(false);  // Attendees

  // 5. Horizontal and vertical alignment
  getThisRange(['E2:E', 'M2:N']).setHorizontalAlignment('center');  // Headrun + Transfer Status + Submission Platform

  getThisRange([
    'D2:I',   // Headrun Details + Attendees
    'M2:N',   // Transfer Status + Submission Platform
  ]).setVerticalAlignment('middle');

  // 6. Update banding colours by extending range
  const dataRange = sheet.getRange(1,1);
  const banding = dataRange.getBandings()[0];
  banding.setRange(sheet.getDataRange());

  // 7. Resize columns using `sizeMap`
  const sizeMap = {
    [TIMESTAMP_COL]: 150,
    [EMAIL_COL]: 240,
    [HEADRUNNERS_COL]: 240,
    [HEADRUN_COL]: 155,
    [RUN_LEVEL_COL]: 170,
    [ATTENDEES_BEGINNER_COL]: 160,
    [ATTENDEES_EASY_COL]: 160,
    [ATTENDEES_INTERMEDIATE_COL]: 160,
    [ATTENDEES_ADVANCED_COL]: 160,
    [CONFIRMATION_COL]: 300,
    [DISTANCE_COL]: 160,
    [COMMENTS_COL]: 355,
    [TRANSFER_STATUS_COL]: 135,
    [PLATFORM_COL]: 160,
    [NAMES_NOT_FOUND_COL]: 225
  }

  Object.entries(sizeMap).forEach(([col, width]) => {
    sheet.setColumnWidth(col, width);
  });
}

