/**
 * Adds `Google Form` as source of attendance submission.
 *
 * @trigger New Google Form submission.
 *
 * @author [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>)
 * @date  Oct 17, 2023
 * @update  Feb 9, 2025
 *
 */

function addMissingPlatform_(row=ATTENDANCE_SHEET.getLastRow()) {
  const sheet = ATTENDANCE_SHEET;
  const rangePlatform = sheet.getRange(row, PLATFORM_COL);
  rangePlatform.setValue('Google Form');
}


/**
 * Sets sendEmail column to `true` so emailSubmission() can proceed.
 *
 * @trigger New Google Form submission.
 *
 * @author [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>)
 * @date  Feb 9, 2025
 * @update  Feb 9, 2025
 *
 */

function setCopySent_(row=ATTENDANCE_SHEET.getLastRow()) {
  const sheet = ATTENDANCE_SHEET;
  const rangeIsCopySent = sheet.getRange(row, IS_COPY_SENT_COL);

  // Since GForm automatically sends copy to submitter, set isCopySent `true`.
  rangeIsCopySent.insertCheckboxes().setValue(true);
}


/**
 * Sorts `ATTENDANCE_SHEET` according to submission time.
 *
 * @trigger  Edit time.
 *
 * @author [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>)
 * @date  Nov 1, 2024
 * @update  Nov 1, 2024
 *
 */

function sortAttendanceForm() {
  const sheet = ATTENDANCE_SHEET;

  const numRows = sheet.getLastRow() - 1;     // Remove header row from count
  const numCols = sheet.getLastColumn();
  const range = sheet.getRange(2, 1, numRows, numCols);

  // Sorts values by `Timestamp` without the header row
  range.sort([{ column: 1, ascending: true }]);
}


/**
 * Change attendance status of all members to not present.
 *
 * Helper function for `consolidateMemberData()`.
 *
 * @trigger New head run or McRUN attendance submission.
 *
 * @author [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>)
 * @date  Oct 9, 2023
 * @update  Oct 29, 2024
 *
 */

function removePresenceChecks() {
  // `Membership Collected (main)` Google Sheet
  const sheetURL = MEMBERSHIP_URL;
  const ss = SpreadsheetApp.openByUrl(sheetURL);

  // `MASTER` sheet in `Membership Collected (main)`
  const masterSheetName = MASTER_NAME;
  const sheet = ss.getSheetByName(masterSheetName);

  var rangeAttendance;
  var rangeList = sheet.getNamedRanges();

  for (var i = 0; i < rangeList.length; i++) {
    if (rangeList[i].getName() == "attendanceStatus") {
      rangeAttendance = rangeList[i];
      break;
    }
  }

  rangeAttendance.getRange().uncheck(); // remove all Presence checks
}


function prettifySheet() {
  formatSpecificColumns();
  hideAllAttendeeEmail();
}

/**
 * Formats certain columns of `HR Attendance` sheet.
 *
 * Modifies confirmation bool into user-friendly message.
 *
 * @trigger New Google form or app submission.
 *
 * @author [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>)
 * @date  Oct 9, 2023
 * @update  Mar 24, 2025
 */

function formatSpecificColumns() {
  const sheet = ATTENDANCE_SHEET;

  // Helper fuction to improve readability
  const getThisRange = (ranges) => 
    Array.isArray(ranges) ? sheet.getRangeList(ranges) : sheet.getRange(ranges);

  // 1. Freeze panes
  sheet.setFrozenRows(1);
  sheet.setFrozenColumns(1);
  
  // 2. Bold formatting
  getThisRange([
    'A1:N1',  // Header Row
    'A2:A',   // Timestamp
    'D2:D',   // Headrun
    'L2:M'    // Copy Sent + Submission Platform
  ]).setFontWeight('bold');

  // 3. Font size adjustments
  getThisRange(['A2:A', 'D2:D', 'M2:M']).setFontSize(11); // for Headrun + Submission Platform
  getThisRange(['C2:C', 'F2:H']).setFontSize(9);  // Headrunners + Attendees

  // 4. Text wrapping
  getThisRange(['B2:E', 'I2:K']).setWrap(true);  // Referral + Waiver + Payment Choice
  getThisRange('F2:H').setWrap(false);  // Attendees

  // 5. Horizontal and vertical alignment
  getThisRange(['E2:E', 'L2:M']).setHorizontalAlignment('center');  // Headrun + Copy Sent + Submission Platform
 
  getThisRange([
    'D2:H',   // Headrun Details + Attendees
    'L2:M',   // Copy Sent + Submission Platform
  ]).setVerticalAlignment('middle');
  
  // 6. Update banding colours
  const dataRange = sheet.getDataRange();

  // Need to remove current banding, before applying it to current range
  // Apply BLUE banding with distinct header and footer colours.
  dataRange.getBandings().forEach(banding => banding.remove());
  dataRange.applyRowBanding(SpreadsheetApp.BandingTheme.BLUE, true, true);

  // 7. Resize columns using `sizeMap`
  const sizeMap = {
    [TIMESTAMP_COL]: 150,
    [EMAIL_COL]: 240,
    [HEADRUNNERS_COL]: 240,
    [HEADRUN_COL]: 155,
    [RUN_LEVEL_COL]: 170,
    [ATTENDEES_BEGINNER_COL]: 160,
    [ATTENDEES_INTERMEDIATE_COL]: 160,
    [ATTENDEES_ADVANCED_COL]: 160,
    [CONFIRMATION_COL]: 300,
    [DISTANCE_COL]: 160,
    [COMMENTS_COL]: 355,
    [IS_COPY_SENT_COL]: 135,
    [PLATFORM_COL]: 160,
    [NAMES_NOT_FOUND_COL]: 225,
  }

  Object.entries(sizeMap).forEach(([col, width]) => {
    sheet.setColumnWidth(col, width);
  });

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
 * @update  Feb 9, 2025
 */

function formatConfirmationInRow_(row = ATTENDANCE_SHEET.getLastRow()) {
  const sheet = ATTENDANCE_SHEET;
  const confirmationCol = CONFIRMATION_COL;

  // Get confirmation col and value using `row`
  const rangeConfirmation = sheet.getRange(row, confirmationCol);
  const confirmationResp = rangeConfirmation.getValue().toString();    // Options: TRUE or FALSE;

  // Ensure that current value is bool to prevent overwrite
  var isBool = (confirmationResp === 'true' || confirmationResp === 'false');
  if (!isBool) return;

  // Format and set value according to TRUE/FALSE response
  const formattedValue = (confirmationResp === 'true') ? 'Yes' : 'No (explain in comment section)';
  rangeConfirmation.setValue(formattedValue);
  
  // Useful debugging messages
  console.log(`Confirmation Response (raw): ${confirmationResp}`);
  console.log(`Confirmation Response (formatted): ${formattedValue}`);
}

/**
 * Wrapper function for `formatAttendeeNamesInRow` and `formatHeadRunnerInRow`
 * for **ALL** submissions in GSheet.
 *
 * Row number is 1-indexed in GSheet. Header row skipped. Top-to-bottom execution.
 */

function formatAllNames() {
  const funcA = formatAttendeeNamesInRow_.name;
  const funcB = formatHeadRunnerInRow_.name;
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
  console.log('Now attempting to format headrunner and attendee names');
  formatAttendeeNamesInRow_(row);
  formatHeadRunnerInRow_(row);
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
 * @update  March 13, 2025
 *
 * ```javascript
 * // Sample Script ➜ Format names in row `13`.
 * const rowToFormat = 13;
 * formatNamesInRow(rowToFormat);
 * ```
 */

function formatAttendeeNamesInRow_(row = ATTENDANCE_SHEET.getLastRow()) {
  const sheet = ATTENDANCE_SHEET;
  const numColToGet = LEVEL_COUNT;

  // Get attendee names starting from beginner col
  const nameRange = sheet.getRange(row, ATTENDEES_BEGINNER_COL, 1, numColToGet);  // Attendees columns
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
}


/**
 * Create email using details from input `emailDetails' for internal use
 *
 * @param {Map<string>} emailDetails  Information needed to populate email body.
 * @return {string}  HTML code for email.
 *
 * @author [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>)
 * @date  Oct 29, 2024
 * @update  Jan 15, 2025
 *
 * ```javascript
 * // Sample Script ➜ Create email using info.
 * const emailDetails = {
 *    name : 'Monday - 6pm',
 *    distance : '5km',
 *    attendees : 'Bob Burger, Marge Simpson, Mabel Pines',
 *    toEmail : 'alex-the-headrunner@mcrun.ca',
 *    confirmation : false,
 *    notes : 'Bob will pay fee next time.'
 * };
 * var emailHTML = createEmailCopy(emailDetails);
 * ```
 */

function createEmailCopy_(emailDetails) {
  // Check for non-empty key-value object
  const size = Object.keys(emailDetails).length;

  if (size != 6) {
    const objectPrint = JSON.stringify(emailDetails);
    throw Error(`Confirmation email cannot be created due to incorrect mapping of \`emailDetails\` argument:
      ${objectPrint}`);
  }

  // Load HTML template and replace placeholders
  const templateName = COPY_EMAIL_HTML_FILE;
  const template = HtmlService.createTemplateFromFile(templateName);

  template.NAME = emailDetails.name;
  template.DISTANCE = emailDetails.distance;
  template.ATTENDEES = emailDetails.attendees;
  template.EMAIL = emailDetails.toEmail;
  template.CONFIRMATION = emailDetails.confirmation;
  template.NOTES = emailDetails.notes;

  return template.evaluate().getContent();  // Returns string content from populated html template
}


/**
 * Formats all entries in `memberMap` then sorts by searchKey.
 *
 * Removes whitespace and hyphens, strip accents, and capitalize names.
 *
 * @param {string[][]} memberMap  Array of searchkey and their emails.
 * @return {string[]}  Sorted array of formatted names.
 *
 * @author [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>)
 * @date  Nov 1, 2024
 * @update  Dec 14, 2024
 *
 * ```javascript
 * // Sample Script ➜ Format, then sort names.
 * const rawData = [["Francine de-Blé", "francine.de-ble@mail.com"],
 *                  ["BOb-Burger belChEr ", "bob.belcher@mail.com"]];
 * const result = formatAndSortMemberMap_(rawData);
 * Logger.log(result)  // [["Bob Burger Belcher", "bob.belcher@mail.com"],
 *                         [ "Francine De ble", "francine.de-ble@mail.com"]]
 * ```
 */
function formatAndSortMemberMap_(memberMap, searchKeyIndex, emailIndex) {
  const formattedMap = memberMap.map(row => {
    const memberEmail = row[emailIndex];
    const formattedSearchKey = row[searchKeyIndex]
      .trim()
      .normalize("NFD").replace(/[\u0300-\u036f]/g, "")   // Strip accents
      .replace(/[\u2018\u2019']/g, "") // Remove apostrophes (`, ', ’)
      .toLowerCase()
      .replace(/\b\w/g, l => l.toUpperCase())   // Capitalize each word

    // Combine formatted searchkey and email
    return [formattedSearchKey, memberEmail];
  });

  // Sort by formatted searchKey
  formattedMap.sort((a, b) => a[0].localeCompare(b[0]));
  return formattedMap;
}


/**
 * Formats all entries in `names`, swaps lastName and firstName before sorting.
 *
 * Removes whitespace and apostrophes, strip accents and capitalize names.
 *
 * @param {string[]} names  Array of names to format.
 * @return {string[]}  Sorted array of formatted names.
 *
 * @author [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>)
 * @date  Dec 6, 2024
 * @update  Dec 11, 2024
 *
 * ```javascript
 * // Sample Script ➜ Format, swap first and last name, then sort.
 * const rawNames = ["BOb-Burger bulChEr ", "Francine de-Blé"];
 * const result = swapAndFormatName_(rawNames);
 * Logger.log(result)  // ["Bulcher, Bob Burger", "De ble, Francine"]
 * ```
 */

function swapAndFormatName_(names) {
  const formattedNames = names.map(name => {
    let nameParts = name
      .trim()
      .normalize("NFD").replace(/[\u0300-\u036f]/g, "")   // Strip accents
      .replace(/[\u2018\u2019']/g, "")  // Remove apostrophes (`, ', ’)
      .toLowerCase()
      .replace(/\b\w/g, l => l.toUpperCase())   // Capitalize each name
      .split(/\s+/) // Split by spaces
      ;

    // Replace hyphens with spaces. Can only perform after splitting first and last name.
    nameParts = nameParts.map(name => name.replace(/-/g, " "));

    // If first name is not hyphenated, only left-most substring stored in first name
    const firstName = nameParts[0];
    const lastName = nameParts[nameParts.length - 1];
    return `${lastName}, ${firstName}`; // Format as "LastName, FirstName"
  });

  return formattedNames.sort();
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
