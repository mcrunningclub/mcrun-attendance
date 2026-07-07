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
 *  - **formatAllHeadRun()**
 *  - **formatAllHeadRunner()**
 *  - **formatAllAttendeeNames()**
 *  - **formatAllConfirmations()**
 *
 * Row number is 1-indexed in GSheet. Header row skipped. Top-to-bottom execution.
 *
 * @author [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>)
 * @date  Dec 11, 2024
 * @update  Dec 11, 2024
 */

function cleanSheetData() {
  formatAllHeadRun();   // Removes hyphen-space if applicable
  formatAllHeadRunner();  // Applies uniform formatting to headrunners
  formatAllConfirmations();  // Modifies bool to user-friendly message
  formatAllAttendeeNames();  // Applies uniform formatting to attendees
}

/**
 * Formats all headrun entries in the attendance sheet.
 *
 * Removes hyphen-space if applicable.
 */

function formatAllHeadRun() {
  runOnSheet_(formatHeadrunInRow_.name);
}

/**
 * Wrapper function for `formatAllConfirmation` for **ALL** submissions.
 *
 * Row number is 1-indexed in GSheet. Header row skipped. Top-to-bottom execution.
 */

function formatAllConfirmations() {
  runOnSheet_(formatConfirmationInRow_.name);
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
  const funcB = formatHeadrunnerInRow_.name;
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
  formatHeadrunnerInRow_(row);
  SpreadsheetApp.flush();   // Apply all changes
}


/**
 * Wrapper function for `formatAttendeeNamesInRow` for *ALL* submissions.
 *
 * Row number is 1-indexed in GSheet. Header row skipped. Top-to-bottom execution.
 */

function formatAllAttendeeNames() {
  runOnSheet_(formatAttendeeNamesInRow_.name);
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
 * Create email using details from input `emailDetails' for internal use
 *
 * @param {Map<string>} emailDetails  Information needed to populate email body.
 * @return {string}  HTML code for email.
 *
 * @author [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>)
 * @date  Oct 29, 2024
 * @update  May 15, 2025
 *
 * ```js
 * // Sample Script ➜ Create email using info.
 * const emailDetails = {
 *    title : 'Monday - 6pm',
 *    distance : '5km',
 *    attendees : `['- Beginner: Bob Burger', '- Easy: Marge Simpson, Mabel Pines']`,
 *    confirmation : No,
 *    comments : 'Bob will pay fee next time.'
 * };
 * const emailHTML = createEmailCopy(emailDetails);
 * ```
 */

function createEmailCopy_(emailDetails) {
  // Check for non-empty key-value object
  const size = Object.keys(emailDetails).length;

  if (size < 5) {
    const objectPrint = JSON.stringify(emailDetails);
    throw Error(`Confirmation email cannot be created due to incorrect mapping of \`emailDetails\` argument:
      ${objectPrint}`);
  }

  // Load HTML template and replace placeholders
  const templateName = COPY_EMAIL_TEMPLATE;
  const template = HtmlService.createTemplateFromFile(templateName);

  template.TITLE = emailDetails.title;
  template.DISTANCE = emailDetails.distance;
  template.ATTENDEES = emailDetails.attendees;
  template.CONFIRMATION = emailDetails.confirmation;
  template.COMMENTS = emailDetails.comments;

  return template.evaluate().getContent();  // Returns string content from populated html template
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