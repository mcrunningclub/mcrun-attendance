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
  const rangePlatform = sheet.getRange(row, PLATFORM_COL);
  rangePlatform.setValue('Google Form');
}


/**
 * Converts a string to title case.
 *
 * @param {string} inputString  The string to be converted to title case.
 * @return {string}  The title-cased string.
 */

function toTitleCase_(inputString) {
  return inputString.replace(/\w\S*/g, word => {
    return word.charAt(0).toUpperCase() + word.substr(1).toLowerCase();
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
  const confirmationCol = CONFIRMATION_COL;

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
  console.log(`[AC] Confirmation Response (raw): ${confirmationResp}   (formatted): ${formattedValue}`);
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
  console.log('[AC] Now attempting to format headrunner and attendee names');
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
  console.log(`[AC] Completed formatting of attendee names`, namesArr);
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
  const templateName = COPY_EMAIL_HTML_FILE;
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
