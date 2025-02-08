// IMPORT SHEET CONSTANTS
const IMPORT_SHEET_ID = '82376152';
const IMPORT_SHEET = SpreadsheetApp.getActiveSpreadsheet().getSheetById(IMPORT_SHEET_ID);

// MAPPING FROM MASTER ATTENDANCE SHEET TO SEMESTER SHEET
const IMPORT_MAP = {
  'timestamp' : TIMESTAMP_COL,
  'headrunners' : HEADRUNNERS_COL,
  'headRun' : HEADRUN_COL,
  'runLevel' : RUN_LEVEL_COL,
  'confirmation' : CONFIRMATION_COL,
  'distance' : DISTANCE_COL,
  'comments' : COMMENTS_COL,
  'platform' : PLATFORM_COL,
  'attendees' : ATTENDEES_BEGINNER_COL,
}

// USED TO IMPORT NEW ATTENDANCE SUBMISSION FROM APP
function onChange(e) {
  // Get details of edit event's sheet
  console.log(e);
  const thisSource = e.source;
  
  // Try-catch to prevent errors when sheetId cannot be found
  try {
    const thisSheetID = thisSource.getSheetId();
    const thisLastRow = thisSource.getLastRow();

    if (thisSheetID == IMPORT_SHEET_ID) {
      const importSheet = thisSource.getSheetById(thisSheetID);
      const registrationObj = importSheet.getRange(thisLastRow, 1).getValue();

      const lastRow = copyToSemesterSheet(registrationObj);
      // TRIGGER CODE BELOW
      //....................

    }
  }
  catch (error) {
    console.log(error);
  }

}

function transferLastImport() {
  const thisLastRow = IMPORT_SHEET.getLastRow();
  transferThisRow(thisLastRow);
}

function transferThisRow(row) {
  const registrationObj = IMPORT_SHEET.getRange(row, 1).getValue();
  const lastRow = copyToSemesterSheet(registrationObj);
  onFormSubmit(lastRow);
}


/** 
 * Transfer new attendance submission from `Import` to semester sheet.
 * 
 * @param {Object} attendance  Information on attendance submission.

 * @param {integer} [row=getLastSubmission()]  Target row in `Attendance_Sheet`.
 * 
 * @return {integer}  Latest row in `Attendance_Sheet`.
 * 
 * @author [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>)
 * @date  Feb 8, 2025
 * @update  Feb 8, 2025
 * 
 */

function copyToSemesterSheet(attendance, row=ATTENDANCE_SHEET.getLastRow()) {
  const attendanceSheet = ATTENDANCE_SHEET;
  const importMap = IMPORT_MAP;

  const attendanceObj = JSON.parse(attendance);

  const startRow = row + 1;
  const colSize = attendanceSheet.getLastColumn();

  const valuesByIndex = Array(colSize);   // Array.length = colSize

  for (const [key, value] of Object.entries(attendanceObj)) {
    if (key in importMap) {
      let indexInMain = importMap[key] - 1;   // Set 1-index to 0-index for `setValues()`
      let holder = String(value).replace(/,+\s*$/, '');   // Remove trailing commas and spaces
      valuesByIndex[indexInMain] = holder.replace(/;/g, '\n');    // Replace semi-colon with newline
    }
  }

  // Set values of registration
  const rangeToImport = attendanceSheet.getRange(startRow, 1, 1, colSize);
  rangeToImport.setValues([valuesByIndex]);

  return startRow;
}


/**
 * Find row index of last submission in reverse using while-loop.
 * 
 * Used to prevent native `sheet.getLastRow()` from returning empty row.
 * 
 * @return {integer}  Returns 1-index of last row in GSheet.
 *  
 * @author [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>)
 * @date  Feb 8, 2025
 * @update  Feb 8, 2025
 */

function getLastSubmission_() {
  const sheet = ATTENDANCE_SHEET;
  const startRow = 1;
  const numRow = sheet.getLastRow();
  
  // Fetch all values in the TIMESTAMP_COL
  const values = sheet.getRange(startRow, COLUMN_MAP.TIMESTAMP, numRow).getValues();
  let lastRow = values.length;

  // Loop through the values in reverse order
  while (values[lastRow - 1][0] === "") {
    lastRow--;
  }

  return lastRow;
}


function testMigrate() {
  const ex = `{
    "timestamp":"2025-02-08 10:07:41",
    "headrunners":"Isabella V.;Liam G.;Liam M.;Theo G.;Zisheng H.",
    "headRun":"Saturday - 10:00AM",
    "runLevel":"Beginner",
    "attendees":"Dante D'Alessandro:dante.dalessandro@mail.mcgill.ca;Lisa Stewart:lisa.stewart@mail.mcgill.ca;Romeo Hor:romeo.hor@mail.mcgill.ca;Solal Michon:solal.michon@mail.mcgill.ca;Lakshya Sethi:lakshya.sethi@mail.mcgill.ca",
    "confirmation":true,
    "distance":"5k",
    "comments":"",
    "platform":"McRUN App"
  }`;

  const newRowIndex = copyToSemesterSheet(ex);
  Logger.log(newRowIndex);
}

