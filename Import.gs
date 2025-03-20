// IMPORT SHEET CONSTANTS
const IMPORT_SHEET_ID = 82376152;
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
// TRIGGERED BY ZAPIER AUTOMATION OR BY MASTER ATTENDANCE SHEET

// UPDATE : ZAPIER AUTOMATION DOES NOT TRIGGER INSTANTLY
// SO `ONCHANGE` NO LONGER NEEDED, AND TRIGGER IS WHEN CHECKING FOR ATTENDANCE

/** 
 * Process latest imported attendance submission via McRUN app.
 * 
 * Verifies if import is JSON-formatted string or GSheet multi-column import.
 * 
 * @param {SpreadsheetApp.Sheet} sourceSheet  Sheet with latest submission.
 * 
 * @author [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>)
 * @date  Feb 10, 2025
 * @update  Mar 20, 2025
 * 
 */

function processOnChange(sourceSheet) {
  const thisLastRow = sourceSheet.getLastRow();
  const thisColSize = sourceSheet.getLastColumn();
  const latestImport = sourceSheet.getSheetValues(thisLastRow, 1, 1, thisColSize)[0];   // Get last row

  const keys = sourceSheet.getSheetValues(1, 1, 1, thisColSize)[0];  // Get header row
  let submissionStr;

  // Case 1: JSON-formatted import (single-column)
  if (latestImport[1] === "") {
    console.log("Entered case 1 in onChange()!", `thisLastRow: ${thisLastRow}`);
    submissionStr = (latestImport[0] !== "") 
      ? latestImport[0] 
      : sourceSheet.getSheetValues(thisLastRow - 1, 1, 1, thisColSize)[0];    // Try with second-last row
  }

  // Case 2: Multi-column import (e.g., from Zapier)
  else {
    console.log("Entered case 2 in `onChange()`!");
    submissionStr = packageAttendance(keys, latestImport);
  }

  // Useful debugging message
  console.log(submissionStr);

  // Verify if already imported
  const isImportedCol = keys.indexOf('isImported') + 1;
  const isImported = sourceSheet.getRange(thisLastRow, isImportedCol).getValue();
  console.log(`isImported: ${isImported}`);

  if (isImported) {
    Logger.log(`This submission has been already been transferred`);
    return;
  }

  // Otherwise, continue importing latest submission
  const attendanceObj = JSON.parse(submissionStr);
  const lastSemesterRow = copyToSemesterSheet(attendanceObj);

  // Log successful transfer
  toggleSuccessfulImport(thisLastRow, isImportedCol);
  
  // TRIGGER MAINTENANCE FUNCTIONS
  if((attendanceObj['platform']).toLowerCase() === 'mcrun app'){
    console.log("Entering onAppSubmission() now...");
    onAppSubmission(lastSemesterRow);
  }

}


function toggleSuccessfulImport(row, colIndex = null) {
  const sheet = IMPORT_SHEET;
  let isImportedCol = colIndex;

  if (!colIndex) {
    const header = sheet.getSheetValues(1, 1, 1, sheet.getLastColumn())[0];
    isImportedCol = header.indexOf('isImported') + 1; // Get 0-index, then make 1-indexed
  }

  const isImportedRange = sheet.getRange(row, isImportedCol);
  isImportedRange.setValue(true);
}


/** 
 * Compare the input timestamps.
 * 
 * @param {string} ts1  Timestamp 1.
 * @param {string} ts2  Timestamp 2.
 * 
 * @return {Boolean}  Returns result of comparaison.
 * 
 * @author [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>)
 * @date  Mar 20, 2025
 * @update  Mar 20, 2025
 * 
 */

function isSameTimestamp(ts1, ts2) {
  return Date(ts1) === Date(ts2);
}


function transferThisRow_(row) {
  const registrationObj = JSON.parse(IMPORT_SHEET.getRange(row, 1).getValue());
  const latestSemesterRow = copyToSemesterSheet(registrationObj);
  toggleSuccessfulImport(row);

  //onAppSubmission(latestSemesterRow);
  //onFormSubmission(latestSemesterRow);
}

function transferLastImport() {
  const thisLastRow = IMPORT_SHEET.getLastRow();
  transferThisRow_(thisLastRow);
}


/** 
 * Transfer new attendance submission from `Import` to semester sheet.
 * 
 * @param {Object<JSON>} attendanceJSON   Attendance information as JSON object.
 * 
 * @param {integer} [row=getLastSubmission()]  Target row in `Attendance_Sheet`.
 * 
 * @return {integer}  Latest row in `Attendance_Sheet`.
 * 
 * @author [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>)
 * @date  Feb 8, 2025
 * @update  March 13, 2025
 * 
 */

function copyToSemesterSheet(attendanceJSON, row=ATTENDANCE_SHEET.getLastRow()) {
  const attendanceSheet = ATTENDANCE_SHEET;
  const importMap = IMPORT_MAP;

  const startRow = row + 1;
  const colSize = attendanceSheet.getLastColumn();

  const valuesByIndex = Array(colSize);   // Array.length = colSize

  // Format timestamp correctly and replace
  const timestampValue = attendanceJSON['timestamp'];
  attendanceJSON['timestamp'] = formatTimestamp(timestampValue);

  for (const [key, value] of Object.entries(attendanceJSON)) {
    if (key in importMap) {
      let indexInMain = importMap[key] - 1;   // Set 1-index to 0-index for `setValues()`
      let holder = String(value).replace(/,+\s*$/, '');   // Remove trailing commas and spaces
      valuesByIndex[indexInMain] = holder.replace(/;/g, '\n');    // Replace semi-colon with newline
    }
  }

  // Process attendees for all run levels
  const attendeeMap = ATTENDEE_MAP;

  // Get formatted runLevel and attendees value by index
  const runLevelIndex = importMap['runLevel'] - 1;
  const runLevel = (valuesByIndex[runLevelIndex]).toLowerCase();

  const attendeeIndex = importMap['attendees'] - 1;
  const attendees = valuesByIndex[attendeeIndex].replace(/\s?,\s?/g, '\n');

  for (const [level, levelIndex] of Object.entries(attendeeMap)) {
    const arrIndex = levelIndex - 1;   // Transform 1-index to 0-index for array
    valuesByIndex[arrIndex] = (runLevel === level) ? attendees : EMPTY_ATTENDEE_FLAG;
  }

  // Set values of registration
  const rangeToImport = attendanceSheet.getRange(startRow, 1, 1, colSize);
  rangeToImport.setValues([valuesByIndex]);

  return startRow;
}


/** 
 * Create JSON-formatted string of key-value pairs for attendance submission.
 * 
 * @param {string[]} keyArr  Array of keys storing header row values.
 * 
 * @param {string[]} valArr  Values of attendance submission to map.
 * 
 * @return {string}  A JSON string of attendance submission as key-value pairs.
 * 
 * @author [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>)
 * @date  Feb 9, 2025
 * @update  Feb 9, 2025
 * 
 */

function packageAttendance(keyArr, valArr) {
  if (keyArr.length !== valArr.length) {
    const errMessage = `keyArr and valArr must have the same length.
      keyArr: ${keyArr}
      valArr: ${valArr}`
    ;
    throw new Error(errMessage);
  }

  // Create JSON string
  const obj = Object.fromEntries(keyArr.map((key, i) => [key, valArr[i]]));
  return JSON.stringify(obj);
}


/** 
 * Format timestamp to format as `yyyy-MM-dd hh:mm:ss`.
 * 
 * Raw format cannot be understood by GSheet.
 * 
 * @param {string} raw  Datetime value to be formatted.
 * 
 * @return {Date}  A Date object with correct format.
 * 
 * @author [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>)
 * @date  Feb 9, 2025
 * @update  Feb 10, 2025
 * 
 */

function formatTimestamp(raw) {
  const date = new Date(raw);
  const options =  {
    year: 'numeric',
    month: '2-digit',
    day: '2-digit',
    hour: '2-digit',
    minute: '2-digit',
    second: '2-digit',
    hour12: false // 24-hour format
  };

  return date.toLocaleString('en-CA', options).replace(',', '');  // remove comma between date and time
}
