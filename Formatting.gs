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
 * Adds `Google Form` as source of attendance submission.
 *
 * @trigger New Google Form submission.
 *
 * @author [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>)
 * @date  Oct 17, 2023
 * @update  Feb 9, 2025
 */

function addMissingPlatform_(row = ATTENDANCE_SHEET.getLastRow()) {
  const sheet = ATTENDANCE_SHEET;
  const rangePlatform = sheet.getRange(row, SEM_ATTENDANCE_COLS.PLATFORM);
  rangePlatform.setValue('Google Form');
}



/**
 * Global wrapper function that runs the following on the sheet:
 *  - formats head run column for each row
 *  - formats headrunner names for each row
 *  - formats attendees' names for each row
 *  - formats confirmations in each row
 *
 * Row number is 1-indexed in GSheet. Header row skipped. Top-to-bottom execution.
 *
 * @author [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>)
 * @date  Dec 11, 2024
 * @update  Dec 11, 2024
 */

function formatSemesterAttendance() {
  runOnSheet_(formatHeadrunInRow_.name);   // Removes hyphen-space if applicable
  runOnSheet_(formatHeadrunnersInRow_.name);  // Applies uniform formatting to headrunners
  runOnSheet_(formatConfirmationInRow_.name);  // Modifies bool to user-friendly message
  runOnSheet_(formatAttendeeNamesInRow_.name);  // Applies uniform formatting to attendees
}


/**
 * Formats confirmation bool in `row` into user-friendly string.
 *
 * @param {integer} [row=ATTENDANCE_SHEET.getLastRow()]  The row in the `ATTENDANCE_SHEET` sheet (1-indexed).
 *                                                       Defaults to the last row in the sheet.
 *
 * @author [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>)
 * @date  Dec 8, 2024
 * @update  Apr 7, 2025
 */

function formatConfirmationInRow_(row = ATTENDANCE_SHEET.getLastRow()) {
  const sheet = GET_ATTENDANCE_SHEET_();
  const confirmationCol = SEM_ATTENDANCE_COLS.CONFIRMATION;

  // Get confirmation col and value using `row`
  const rangeConfirmation = sheet.getRange(row, confirmationCol);
  const confirmationResp = rangeConfirmation.getValue().toString();    // Options: TRUE or FALSE;

  // Ensure that current value is bool to prevent overwrite
  const isBool = (confirmationResp === 'true' || confirmationResp === 'false');
  if (!isBool) return;

  // Format and set value according to TRUE/FALSE response
  const formattedValue = (confirmationResp === 'true') ? 'Yes' : 'No (explain in comment section)';
  rangeConfirmation.setValue(formattedValue);

  // Log debugging message
  logAsAC_(
    `Confirmation (raw): ${confirmationResp}   (formatted): ${formattedValue}`,
    formatConfirmationInRow_.name,
    false
  );
}


/**
 * Wrapper function for `formatAttendeeNamesInRow` and `formatHeadRunnerInRow`
 * for **ALL** submissions in GSheet.
 *
 * Row number is 1-indexed in GSheet. Header row skipped. Top-to-bottom execution.
 */

function formatAllNames() {
  const funcA = formatAttendeeNamesInRow_.name;
  const funcB = formatHeadrunnersInRow_.name;
  runOnSheet_(funcA, funcB);  // Run both functions
}

/**
 * Wrapper function for `formatAttendeeNamesInRow` and `formatHeadRunnerInRow`.
 *
 * Formats headrunner and attendee names in target `row`.
 *
 * @param {integer} [row=ATTENDANCE_SHEET.getLastRow()]  The row in the `ATTENDANCE_SHEET` sheet (1-indexed).
 *                                                       Defaults to the last row in the sheet.
 */

function formatNamesInRow_(row = ATTENDANCE_SHEET.getLastRow()) {
  logAsAC_('Now attempting to format headrunner and attendee names', formatNamesInRow_.name, false);
  formatAttendeeNamesInRow_(row);
  formatHeadrunnersInRow_(row);
  SpreadsheetApp.flush();   // Apply all changes
}


/**
 * Formats attendee names from `row` into uniform view, sorted and separated by newline.
 *
 * @param {integer} [row=ATTENDANCE_SHEET.getLastRow()]  The row in the `ATTENDANCE_SHEET` sheet (1-indexed).
 *                                                       Defaults to the last row in the sheet.
 *
 * @author [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>)
 * @date  Oct 24, 2024
 * @update  Apr 7, 2025
 *
 * ```javascript
 * // Sample Script ➜ Format names in row `13`.
 * const rowToFormat = 13;
 * formatNamesInRow(rowToFormat);
 * ```
 */

function formatAttendeeNamesInRow_(row = ATTENDANCE_SHEET.getLastRow()) {
  const sheet = GET_ATTENDANCE_SHEET_();
  const numColToGet = NUM_LEVELS;

  logAsAC_(`Starting name formatting in row #${row}`, formatAttendeeNamesInRow_.name, false);

  // Get attendee names starting from beginner col
  const nameRange = sheet.getRange(row, SEM_ATTENDANCE_COLS.B_ATTENDEES, 1, numColToGet);  // Attendees columns
  var namesArr = nameRange.getValues()[0];    // 1D Array of size 3 (Beginner, Intermediate, Advanced)

  for (var i = 0; i < namesArr.length; i++) {
    var trimmedArr = namesArr[i].trim();

    // Case 1: Cell is non-empty and does not contains "None"
    if (trimmedArr.length != 0 && trimmedArr !== EMPTY_ATTENDEE_FLAG) {

      // Replace "n/a" (case insensitive) with EMPTY_ATTENDEE_FLAG value : "None"
      var cellValue = trimmedArr.replace(/n\/a/gi, EMPTY_ATTENDEE_FLAG);

      // Exit if cell contains email and ':' delimiter -> already formatted.
      if (cellValue.includes('@') && cellValue.includes(':')) continue;

      // Split by commas or newline characters
      var names = cellValue.split(/[,|\n]+/);

      // Remove whitespace, strip accents and capitalize names
      var formattedNames = names.map(name => name
        .trim()
        .normalize("NFD").replace(/[\u0300-\u036f]/g, "")  // Strip accents
        .replace(/[\u2018\u2019']/g, "") // Remove apostrophes (`, ', ’)
        .toLowerCase()
        .replace(/\b\w/g, l => l.toUpperCase()) // Capitalize each name
      );

      // Sort names alphabetically
      formattedNames.sort();

      // Join back with newline characters
      namesArr[i] = formattedNames.join('\n');
    }

    // Case 2: Cell is empty
    else {
      namesArr[i] = EMPTY_ATTENDEE_FLAG;
    }

  }
  // Replace values with formatted names
  nameRange.setValues([namesArr]);    // setValues() requires 2D array
  logAsAC_(
    `Completed formatting of attendee names\n${namesArr}`, 
    formatAttendeeNamesInRow_.name, 
    false
  );
}


/**
 * Formats headrunner names from `row` into uniform view, separated by newline.
 *
 * Updated format is '`${firstName} ${lastNameLetter}.`'
 *
 * @param {integer} [row=ATTENDANCE_SHEET.getLastRow()]  The row in the `ATTENDANCE_SHEET` sheet (1-indexed).
 *                                                       Defaults to the last row in the sheet.
 *
 * @param {integer} numRow  Number of rows to format from `startRow`.
 *
 * @author [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>)
 * @date  Dec 10, 2024
 * @update  Apr 7, 2024
 *
 * ```javascript
 * // Sample Script ➜ Format names in row `7`.
 * const rowToFormat = 7;
 * formatHeadrunnerInRow(rowToFormat);
 *
 * // Sample Script ➜ Format names from row `3` to `9`.
 * const startRow = 3;
 * const numRow = 9 - startRow;
 * formatHeadrunnerInRow(startRow, numRow);
 * ```
 */

function formatHeadrunnersInRow_(startRow = ATTENDANCE_SHEET.getLastRow(), numRow = 1) {
  const sheet = GET_ATTENDANCE_SHEET_();
  const headrunnerCol = SEM_ATTENDANCE_COLS.HEADRUNNERS;

  // Get all the values in `HEADRUNNERS_COL` in bulk
  const rangeHeadRunner = sheet.getRange(startRow, headrunnerCol, numRow);
  const rawValues = rangeHeadRunner.getValues();

  // Callback function to process the raw value into the formatted format
  function processRow(row) {
    const headrunners = row[0]  // Get first column from 2D array
      .split(/[,|\n]+/)         // Split by commas or newlines
      .map(formatHeadrunnerName_)   // Format each name using formatName()
      .join('\n');       // Join the names with a newline

    return [headrunners]; // Return as a 2D array for .setValues()
  };

  // Map over each row to process and format by applying `processRow()`
  const formattedNames = rawValues.map(processRow);   // apply processRow()

  // Update the sheet with formatted names
  rangeHeadRunner.setValues(formattedNames);
  logAsAC_(
    `Completed formatting of headrunner names\n${formattedNames.join(';')}`,
    formatHeadrunnersInRow_.name,
    false
  );
}

/**
 * Callback function to clean and format a single headrunner name
 * 
 * @param {string} name  Original string
 * @return {string}  Formatted string
 */
function formatHeadrunnerName_(name) {
  const cleanedName = name
    .trim()
    .normalize("NFD").replace(/[\u0300-\u036f]/g, "") // Remove accents
    .toLowerCase()
    .replace(/\b\w/g, letter => letter.toUpperCase()); // Capitalize each proper name

  // Split into first, second (and rest if applicable)
  const [firstPart, secondPart, ...rest] = cleanedName.split(' ');

  // Get initial of last name (and prepend second part if applicable)
  const lastPart = rest.length === 0 ? 
    getInitial(secondPart) :
   `${secondPart} ${getInitial(rest.join(''))}`
  ;

  return `${firstPart} ${lastPart}`;  // Return formatted name

  function getInitial(name) {
    const initial = (name.charAt(0) || '').toUpperCase();
    return initial ? `${initial}.` : '';
  }
};


/**
 * Removes hyphen-space in headrun from `row` if applicable.
 *
 * @param {integer} [startRow=ATTENDANCE_SHEET.getLastRow()]
 *                      The row in the `ATTENDANCE_SHEET` sheet (1-indexed).
 *                      Defaults to the last row in the sheet.
 *
 * @param {integer} [numRow=1] Number of rows to format from `startRow`
 *
 * @author [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>)
 * @date  Dec 10, 2024
 * @update  Apr 7, 2025
 */
function formatHeadrunInRow_(startRow = ATTENDANCE_SHEET.getLastRow(), numRow = 1) {
  const sheet = GET_ATTENDANCE_SHEET_();

  // Get the cell value, and remove hyphen-space in each cell
  const rangeToFormat = sheet.getRange(startRow, SEM_ATTENDANCE_COLS.HEADRUN, numRow);
  var values = rangeToFormat.getValues();

  // Bulk format if applicable
  var formattedHeadRun = values.map(row => {
    let cleanValue = row[0].toString().replace(/- /g, "");
    return [cleanValue] // must return as 2d
  });

  // Replace with formatted value
  rangeToFormat.setValues(formattedHeadRun);  // setValues requires 2d array
}

/**
 * Boiler plate function `functionName` to execute on complete sheet.
 *
 * Also executes `functionName2` if non-empty.
 *
 * @param {string}  functionName  Name of function to execute.
 * @param {string}  [functionName2=""]  Name of function to execute.
 *                                      Defaults to empty string.
 *
 * @author [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>)
 * @date  Dec 11, 2024
 * @update  Dec 11, 2024
 */

function runOnSheet_(functionName, functionName2 = "") {
  const sheet = ATTENDANCE_SHEET;
  const startRow = 2  // Skip header row
  const numRows = sheet.getLastRow();

  for (var row = startRow; row <= numRows; row++) {
    this[functionName](row);

    // Only executes `functionName2` if non-empty.
    if (functionName2) {
      this[functionName2](row);
    }
  }
}

/**
 * Sorts the `ATTENDANCE_SHEET` by submission time.
 * Excludes the header row from sorting.
 *
 * @author [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>)
 * @date  Nov 1, 2024
 * @update  Apr 7, 2025
 */

function sortSemesterAttendance() {
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

  const masterSheetName = MEMBERSHIP_SHEET_NAME;
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
  const allCols = Object.entries(SEM_ATTENDANCE_COLS).length;
  const lastRow = getLastRow_();

  // Helper functions
  function getHeaderCells(columns) {
    const str_cells = columns.map((colNumber) => {
      const letter = String.fromCharCode(64 + colNumber);
      return letter + '1';
    })
    return sheet.getRangeList(str_cells);
  }

  function getHeaderRow() {
    return sheet.getRange(1, 1, 1, allCols);
  }

  function getColumnsExceptHeader(columns) {
    const str_columns = columns.map((colNumber) => {
      const letter = String.fromCharCode(64 + colNumber);
      return letter + '2:' + letter + lastRow;
    })
    return sheet.getRangeList(str_columns);
  }

  function getColumnsIncludingHeader(columns) {
    const str_columns = columns.map((colNumber) => {
      const letter = String.fromCharCode(64 + colNumber);
      return letter + '1:' + letter;
    })
    return sheet.getRangeList(str_columns);
  }

  // 1. Freeze panes
  sheet.setFrozenRows(1);
  sheet.setFrozenColumns(1);

  // 2. Bold formatting
  getHeaderRow().setFontWeight('bold');
  getColumnsExceptHeader([
    SEM_ATTENDANCE_COLS.TIMESTAMP,
    SEM_ATTENDANCE_COLS.HEADRUN,
    SEM_ATTENDANCE_COLS.TRANSFER_STATUS,
    SEM_ATTENDANCE_COLS.PLATFORM,
    SEM_ATTENDANCE_COLS.NOT_FOUND
  ]).setFontWeight('bold');

  // 3. Font size adjustments
  getColumnsExceptHeader([
    SEM_ATTENDANCE_COLS.TIMESTAMP,
    SEM_ATTENDANCE_COLS.HEADRUN,
    SEM_ATTENDANCE_COLS.PLATFORM
  ]).setFontSize(11);
  getColumnsExceptHeader([
    SEM_ATTENDANCE_COLS.HEADRUNNERS,
    SEM_ATTENDANCE_COLS.A_ATTENDEES,
    SEM_ATTENDANCE_COLS.B_ATTENDEES,
    SEM_ATTENDANCE_COLS.E_ATTENDEES,
    SEM_ATTENDANCE_COLS.I_ATTENDEES
  ]).setFontSize(9);  // Headrunners + Attendees

  // 4. Text wrapping
  getColumnsExceptHeader([
    SEM_ATTENDANCE_COLS.EMAIL,
    SEM_ATTENDANCE_COLS.HEADRUNNERS,
    SEM_ATTENDANCE_COLS.HEADRUN,
    SEM_ATTENDANCE_COLS.RUN_LEVEL,
    SEM_ATTENDANCE_COLS.CONFIRMATION,
    SEM_ATTENDANCE_COLS.DISTANCE,
    SEM_ATTENDANCE_COLS.COMMENTS
  ]).setWrap(true);
  getColumnsExceptHeader([
    SEM_ATTENDANCE_COLS.HEADRUNNERS,
    SEM_ATTENDANCE_COLS.A_ATTENDEES,
    SEM_ATTENDANCE_COLS.B_ATTENDEES,
    SEM_ATTENDANCE_COLS.E_ATTENDEES,
    SEM_ATTENDANCE_COLS.I_ATTENDEES
  ]).setWrap(false);  // Attendees

  // 5. Horizontal and vertical alignment
  getColumnsExceptHeader([
    SEM_ATTENDANCE_COLS.HEADRUN,
    SEM_ATTENDANCE_COLS.TRANSFER_STATUS,
    SEM_ATTENDANCE_COLS.PLATFORM
  ]).setHorizontalAlignment('center');  // Headrun + Transfer Status + Submission Platform
  getColumnsExceptHeader([
    SEM_ATTENDANCE_COLS.HEADRUN,
    SEM_ATTENDANCE_COLS.RUN_LEVEL,
    SEM_ATTENDANCE_COLS.HEADRUNNERS,
    SEM_ATTENDANCE_COLS.A_ATTENDEES,
    SEM_ATTENDANCE_COLS.B_ATTENDEES,
    SEM_ATTENDANCE_COLS.E_ATTENDEES,
    SEM_ATTENDANCE_COLS.I_ATTENDEES,
    SEM_ATTENDANCE_COLS.TRANSFER_STATUS,
    SEM_ATTENDANCE_COLS.PLATFORM
  ]).setVerticalAlignment('middle');

  // 6. Update banding colours by extending range
  const dataRange = sheet.getRange(1,1);
  const banding = dataRange.getBandings()[0];
  banding.setRange(sheet.getDataRange());

  // 7. Resize columns using `sizeMap`
  const sizeMap = {
    [SEM_ATTENDANCE_COLS.TIMESTAMP]: 150,
    [SEM_ATTENDANCE_COLS.EMAIL]: 240,
    [SEM_ATTENDANCE_COLS.HEADRUNNERS]: 240,
    [SEM_ATTENDANCE_COLS.HEADRUN]: 155,
    [SEM_ATTENDANCE_COLS.RUN_LEVEL]: 170,
    [SEM_ATTENDANCE_COLS.B_ATTENDEES]: 160,
    [SEM_ATTENDANCE_COLS.E_ATTENDEES]: 160,
    [SEM_ATTENDANCE_COLS.I_ATTENDEES]: 160,
    [SEM_ATTENDANCE_COLS.A_ATTENDEES]: 160,
    [SEM_ATTENDANCE_COLS.CONFIRMATION]: 300,
    [SEM_ATTENDANCE_COLS.DISTANCE]: 160,
    [SEM_ATTENDANCE_COLS.COMMENTS]: 355,
    [SEM_ATTENDANCE_COLS.TRANSFER_STATUS]: 135,
    [SEM_ATTENDANCE_COLS.PLATFORM]: 160,
    [SEM_ATTENDANCE_COLS.NOT_FOUNT]: 225
  }

  Object.entries(sizeMap).forEach(([col, width]) => {
    sheet.setColumnWidth(col, width);
  });
}