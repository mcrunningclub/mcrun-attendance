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

// IMPORT SHEET CONSTANTS
const IMPORT_SHEET_ID = 82376152;
const IMPORT_SHEET = SpreadsheetApp.getActiveSpreadsheet().getSheetById(IMPORT_SHEET_ID);

// ALLOWS PROPER SHEET REF WHEN ACCESSING AS LIBRARY FROM EXTERNAL SCRIPT
// SpreadsheetApp.getActiveSpreadsheet() DOES NOT WORK IN EXTERNAL SCRIPT
const GET_IMPORT_SHEET_ = () => {
  return (IMPORT_SHEET) ?? SpreadsheetApp.openById(ATTENDANCE_SHEET_ID).getSheetById(IMPORT_SHEET_ID);
}

const IS_IMPORTED_COL = 10;   // Update if columns modified

// MAPPING FROM MASTER ATTENDANCE SHEET TO SEMESTER SHEET
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

// USED TO IMPORT NEW ATTENDANCE SUBMISSION FROM APP
// TRIGGERED BY ZAPIER AUTOMATION OR BY MASTER ATTENDANCE SHEET

// UPDATE : ZAPIER AUTOMATION DOES NOT TRIGGER INSTANTLY
// SO `ONCHANGE` NO LONGER NEEDED, AND TRIGGER USED WHEN CHECKING FOR ATTENDANCE

/** 
 * Process latest imported attendance submission via McRUN app.
 * 
 * Verifies if import is JSON-formatted string or GSheet multi-column import.
 * 
 * @param {SpreadsheetApp.Sheet} sourceSheet  Sheet with latest submission.
 * 
 * @author [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>)
 * @date  Feb 10, 2025
 * @update  Sep 28, 2025
 */

function processImportFromApp(importObj) {
  const importSheet = GET_IMPORT_SHEET_();

  // Log debugging messages
  const funcName = processImportFromApp.name;
  logAsAC_(` Processing following import...`, funcName, false);
  console.log(importObj);

  let newSemesterRow = null;
  try {
    // First add to import sheet as backup
    importSheet.appendRow([importObj]);

    // Now process input
    const attendanceObj = JSON.parse(importObj);
    logAsAC_(`Now trying to export values to row #${newSemesterRow} in Attendance Sheet`, funcName);
    newSemesterRow = copyToSemesterSheet_(attendanceObj);
    logAsAC_(`Successfully imported values!`, funcName);

    // Log successful transfer to attendance sheet
    const newImportRow = importSheet.getLastRow();
    toggleSuccessfulImport_(newImportRow, IS_IMPORTED_COL);
  }
  catch (e) {
    logAsAC_(`Unable to fully import 'importObj' in Attendance Sheet`, funcName);
    throw Error(`${e.message} Import failed...`);
  }

  // Finally apply post-import functions
  try {
    if(newSemesterRow) {
      logAsAC_(`Now trying to apply 'onAppSubmission' for row #${newSemesterRow}`, funcName);
      onAppSubmission(newSemesterRow);
    };
  }
  catch (e) {
    logAsAC_(`Unable to apply '${onAppSubmission.name}' for 'importObj'`, funcName, false);
    throw e;
  }
}



/**
 * Transfers the last imported attendance submission to the semester sheet.
 *
 * @trigger  New head run or McRUN attendance submission.
 */

function transferLastImport() {
  const thisLastRow = getLastRow_(IMPORT_SHEET);
  transferThisRow_(thisLastRow);
}

function transferThisRow_(row) {
  const attendanceObj = JSON.parse(IMPORT_SHEET.getRange(row, 1).getValue());
  const attendanceTimestamp = attendanceObj['timestamp'];

  // Check if attendanceObj already imported in attendance sheet. Exit if true
  const isFound = checkExistingTimestamp_(attendanceTimestamp, 3);   // Check last 3 rows
  if (isFound) return;

  // Transfer if attendance submission not found.
  copyToSemesterSheet_(attendanceObj);
  toggleSuccessfulImport_(row);
}


function toggleSuccessfulImport_(row, colIndex = null) {
  const sheet = GET_IMPORT_SHEET_();
  let isImportedCol = colIndex;

  if (!colIndex) {
    const header = sheet.getSheetValues(1, 1, 1, sheet.getLastColumn())[0];
    isImportedCol = header.indexOf('isImported') + 1;  // 0-index to 1-indexed for `.getRange()`
  }

  const isImportedRange = sheet.getRange(row, isImportedCol);
  isImportedRange.setValue(true);
  logAsAC_(`Toggled successful import in row #${row}`, toggleSuccessfulImport_.name);
}


/** 
 * Check is submission already added by comparing timestamps.
 * 
 * @param {string} timestampToCompare  Input timestamp
 * @param {integer} [numOfRow = 5]  Number of rows to check from the bottom.
 *                                  Defaults to 5.
 * 
 * @return {Boolean}  Returns true if found in attendance sheet.
 * 
 * @author [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>)
 * @date  Mar 20, 2025
 * @update  Mar 20, 2025
 */

function checkExistingTimestamp_(timestampToCompare, numOfRow = 5) {
  const sheet = ATTENDANCE_SHEET;

  // Get dimensions of array
  const numCol = 1;
  const timestampCol = SEM_ATTENDANCE_COLS.TIMESTAMP;
  const endRow = sheet.getLastRow();
  const startRow = endRow - numOfRow + 1;

  // Parse `timestampToCompare` as Date
  const compareAsDate = new Date(timestampToCompare);

  // Get latest timestamp values from attendance sheet as 1d array
  const latestTimestampValues = sheet.getSheetValues(startRow, timestampCol, numOfRow, numCol).flat();

  // Compare timestamps in attendance sheet until found, else return false
  const isFound = latestTimestampValues.some(ts => isSameTimestamp_(ts, compareAsDate));
  console.log(`Timestamp '${compareAsDate}' found in attendance sheet: ${isFound}`);
  return isFound;
}


/** 
 * Transfer new attendance submission from `Import` to semester sheet.
 * 
 * @param {Object<JSON>} attendanceJSON   Attendance information as JSON object.
 * @param {integer} [row=getLastRow_()]  Target row in `Attendance_Sheet`.
 * 
 * @return {integer}  Latest row in `Attendance_Sheet`.
 * 
 * @author [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>)
 * @date  Feb 8, 2025
 * @update  Sep 18, 2025
 */

function copyToSemesterSheet_(attendanceJSON, row = getLastRow_()) {
  // Start with debugging message
  logAsAC_(`Starting execution for row #${row}...`, copyToSemesterSheet_.name);

  const attendanceSheet = GET_ATTENDANCE_SHEET_();
  const importMap = IMPORT_MAP;

  const startRow = row + 1;
  const colSize = attendanceSheet.getLastColumn();

  const valuesByIndex = Array(colSize);   // Array.length = colSize

  // Format timestamp correctly and replace
  const timestampValue = attendanceJSON['timestamp'];
  attendanceJSON['timestamp'] = formatTimestamp_(timestampValue);

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

  // Log and return startRow
  logAsAC_(`Set registration '${timestampValue}' in row #${startRow}`, copyToSemesterSheet_.name);
  return startRow;
}


/** 
 * Create JSON-formatted string of key-value pairs for attendance submission.
 * 
 * @param {string[]} keyArr  Array of keys storing header row values.
 * @param {string[]} valArr  Values of attendance submission to map.
 * @return {string}  A JSON string of attendance submission as key-value pairs.
 * 
 * @author [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>)
 * @date  Feb 9, 2025
 * @update  Feb 9, 2025
 */

function packageAttendance_(keyArr, valArr) {
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