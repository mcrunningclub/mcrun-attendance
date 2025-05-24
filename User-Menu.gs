/**
 * Users authorized to use the McRUN menu.
 *
 * Prevents unwanted data overwrite in Gsheet.
 *
 * @constant {string[]} PERM_USER_ - List of authorized user emails.
 */
const PERM_USER_ = [
  CLUB_EMAIL,
  'ademetriou8@gmail.com',
  'andreysebastian10.g@gmail.com',
  'monaliu832@gmail.com'
];

/**
 * Logs the user attempting to use the custom McRUN menu.
 *
 * If the input is empty, the email is extracted using `getCurrentUserEmail_()`.
 *
 * @trigger User choice in custom menu.
 *
 * @param {string} [email=""]  Email of the active user. Defaults to an empty string.
 *
 * @author [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>)
 * @date  Nov 21, 2024
 * @update  Dec 6, 2024
 */

function logMenuAttempt_(email = "") {
  const userEmail = email ? email : getCurrentUserEmail_();
  Logger.log(`McRUN menu access attempt by: ${userEmail}`);
}

/**
 * Creates a custom menu to run frequently used scripts in Google App Script.
 *
 * Extracts function names using the `name` property to allow for refactoring.
 *
 * Note: Authorization checks cannot be performed here, as unauthorized users
 * would not see the menu due to Google Apps Script limitations.
 *
 * @trigger Open Google Spreadsheet.
 *
 * @author [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>)
 * @date  Dec 5, 2024
 * @update  Apr 1, 2025
 */

function onOpen() {
  const ui = SpreadsheetApp.getUi();

  ui.createMenu('🏃‍♂️ McRUN Attendance Menu')
    .addItem('📢 Custom menu. Click for help.', helpUI_.name)
    .addSeparator()

    .addSubMenu(ui.createMenu('Formatting Menu')
      .addItem('Sort By Timestamp', sortByTimestampUI_.name)
      .addItem('Prettify Attendance Sheet', prettifySheetUI_.name)
      .addItem('Clean Sheet Data', cleanSheetDataUI_.name)
      .addItem('Format Names in Row', formatNamesInRowUI_.name)
      .addItem('Format All Names', formatAllNamesUI_.name)
    )

    .addSubMenu(ui.createMenu('Attendance Menu')
      .addItem('Remove Presence Checks', removePresenceCheckUI_.name)
      .addItem('Find Unregistered in Row', findUnregisteredAttendeesUI_.name)
      .addItem('Find All Unregistered Attendees', findAllUnregisteredUI_.name)
      .addItem('Check For Missing Submission', checkMissingAttendanceUI_.name)
    )

    .addSubMenu(ui.createMenu('Trigger Menu')
      .addItem('Submit By Form', onFormSubmitUI_.name)
      .addItem('Submit By App', onAppSubmitUI_.name)
      .addItem('Turn On/Off Submission Checker', toggleAttendanceCheckUI_.name)
    )

    .addSubMenu(ui.createMenu('Export Menu')
      .addItem('Import from App Records', importAppRecordUI_.name)
      .addItem('Export to Points Ledger', exportToPointsLedgerUI_.name)
    )
    .addToUi();

  checkValidScriptProperties(); // Verify validity of `SCRIPT_PROPERTY`
}


/**
 * Displays a help message for the custom McRUN menu.
 *
 * Accessible to all users.
 */

function helpUI_() {
  const ui = SpreadsheetApp.getUi();

  const helpMessage = `
    📋 McRUN Attendance Menu Help

    - This menu is only accessible to authorized members.

    - Scripts are applied to the attendance sheet.

    - Please contact the admin if you need access or assistance.
  `;

  // Display the help message
  ui.alert("McRUN Menu Help", helpMessage.trim(), ui.ButtonSet.OK);
}


/**
 * Displays a confirmation dialog and executes a function if the user is authorized.
 *
 * @trigger User choice in custom menu.
 *
 * @param {string} functionName  Name of the function to execute.
 * @param {string} [additionalMsg=""]  Custom message to display during execution. Defaults to an empty string.
 * @param {string} [funcArg=""]  Argument to pass to the function. Defaults to an empty string.
 *
 * @return {string}  Return value of the executed function.
 *
 * @author [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>)
 * @date  Dec 5, 2024
 * @update  Dec 6, 2024
 */

function confirmAndRunUserChoice_(functionName, additionalMsg = "", funcArg = "") {
  const ui = SpreadsheetApp.getUi();
  const userEmail = getCurrentUserEmail_();

  // Check if the user is authorized
  if (!PERM_USER_.includes(userEmail)) {
    const warningMsgHeader = "🛑 You are not authorized 🛑";
    const warningMsgBody = "Please contact the exec team if you believe this is an error.";

    ui.alert(warningMsgHeader, warningMsgBody, ui.ButtonSet.OK);
    return;
  }

  // Continue execution if the user is authorized
  let message = `
    ⚙️ Now executing ${functionName}().

    🚨 Press cancel to stop.
  `;

  // Append additional message if non-empty
  message += additionalMsg ? `\n🔔 ${additionalMsg}` : "";

  // Save user response
  const response = ui.alert(message, ui.ButtonSet.OK_CANCEL);
  let retValue = "";

  if (response == ui.Button.OK) {
    // Execute the function (with argument if non-empty)
    retValue = funcArg ? this[functionName](funcArg) : this[functionName]();
  } else {
    ui.alert('Execution cancelled...');
  }

  // Log the attempt in the console using the active user's email
  logMenuAttempt_(userEmail);

  // Return the value from the executed function if required
  return retValue;
}


/**
 * Scripts for `Formatting` menu items.
 *
 * Extracting function name using `name` property to allow for refactoring.
 *
 * Some functions include a custom message for user.
 */

function sortByTimestampUI_() {
  const functionName = sortAttendanceForm.name;
  const customMsg = "This sheet will be sorted by timestamp of submission."
  confirmAndRunUserChoice_(functionName, customMsg);
}

function cleanSheetDataUI_() {
  functionName = cleanSheetData.name;
  const customMsg = "This will clean and formal all the sheet data. \
  \n\nWARNING! Wide-sheet formatting may take some time."

  confirmAndRunUserChoice_(functionName, customMsg);
}

function prettifySheetUI_() {
  const functionName = prettifySheet.name;
  const customMsg = "This will improve the sheet formatting, and hide the attendee emails."
  confirmAndRunUserChoice_(functionName, customMsg);
}

function formatAllNamesUI_() {
  functionName = formatAllNames.name;
  const customMsg = "This will format all the names found in this sheet. \
  \n\nWARNING! Wide-sheet formatting may take some time."

  confirmAndRunUserChoice_(functionName, customMsg);
}

/**
 * This UI function can target a specific row, or the last row if input is omitted.
 */

function formatNamesInRowUI_() {
  const returnObj = requestRowInput_();  // returnObj = {row : int, msg : string}
  const selectedRow = returnObj.row;

  if (!selectedRow) return;  // User input is not valid

  // Execute Function with argument
  const functionName = formatNamesInRow_.name;
  confirmAndRunUserChoice_(functionName, returnObj.msg, selectedRow);
}


/**
 * Scripts for `Attendance` menu items.
 *
 * Extracting function name using `name` property to allow for refactoring.
 *
 * Some functions include a custom message for user.
 */

function removePresenceCheckUI_() {
  const functionName = removePresenceChecks.name;
  const customMsg = "This will remove the presence checkmarks in the membership registry."
  confirmAndRunUserChoice_(functionName, customMsg);
}

function checkMissingAttendanceUI_() {
  const functionName = checkMissingAttendance.name;
  const customMsg = "WARNING. This will send an email reminder to headrunners if attendance is missing."
  confirmAndRunUserChoice_(functionName, customMsg);
}

function toggleAttendanceCheckUI_() {
  const functionName = toggleAttendanceCheck_.name;
  const customMsg = "Turn on to allow McRUN bot (checkMissingAttendance) to check for \
  any missing attendance submissions.\n\nRefer to function documentation for further explanation."

  // Save the return value of `toggleAttendanceCheck`
  const isCheckingAttendance = confirmAndRunUserChoice_(functionName, customMsg);   // saved as str

  // Display info message and new flag value to user
  const ui = SpreadsheetApp.getUi();
  const checkerState = isCheckingAttendance == "true" ? "on ✅" : "off ❌";
  const header = `Submission checker has be successfully turned ${checkerState}.`;

  const message = isCheckingAttendance == "true" ?
    "McRUN Bot will automatically check for submissions. Run again to *stop* future deployments." :
    "McRUN Bot will *not* check for submissions. Run again to *allow* automatic deployments.";
  ui.alert(header, message, ui.ButtonSet.OK);
}

function findAllUnregisteredUI_() {
  functionName = getAllUnregisteredMembers_.name;
  const customMsg = "This will search for *all* unregistered attendees, and add them to `Not Found` column. \
  \n\nWARNING! Wide-sheet formatting will may take some time."

  confirmAndRunUserChoice_(functionName, customMsg);
}

/**
 * This UI function can target a specific row, or the last row if input is omitted.
 */

function findUnregisteredAttendeesUI_() {
  const returnObj = requestRowInput_();  // returnObj = {row : int, msg : string}
  const selectedRow = returnObj.row;

  if (!selectedRow) return;  // User input is not valid

  // Assemble notification message
  const fullMsg = `↪️ This will search for unregistered attendees in ${selectedRow}, and add them to \`Not Found\` column.`;

  // Execute Function with argument
  const functionName = getUnregisteredMembersInRow_.name;
  confirmAndRunUserChoice_(functionName, fullMsg, selectedRow);
}


/**
 * Scripts for `Triggers` menu items.
 *
 * Extracting function name using `name` property to allow for refactoring.
 *
 * Some functions include a custom message for user.
 */

function onFormSubmitUI_() {
  const returnObj = requestRowInput_();  // returnObj = {row : int, msg : string}
  const selectedRow = returnObj.row;

  if (!selectedRow) return;  // User input is not valid

  // Assemble notification message
  const firstMsg = "↪️ Most functions will be triggered when there is a new attendance submission via Google Form.";
  const fullMsg = (returnObj.msg ? `${returnObj.msg}\n\n` : '') + firstMsg;

  // Execute Function with argument
  const functionName = onFormSubmissionInRow_.name;
  confirmAndRunUserChoice_(functionName, fullMsg, selectedRow);
}


function onAppSubmitUI_() {
  const functionName = onAppSubmission.name;
  const customMsg = "This will run functions triggered when new attendance is submitted using the app."
  confirmAndRunUserChoice_(functionName, customMsg);
}


/**
 * Scripts for `Transfer` menu items.
 *
 * Extracting function name using `name` property to allow for refactoring.
 *
 * Some functions include a custom message for user.
 */

function exportToPointsLedgerUI_() {
  const returnObj = requestRowInput_();  // returnObj = {row : int, msg : string}
  const selectedRow = returnObj.row;

  if (!selectedRow) return;  // User input is not valid

  // Assemble notification message
  const firstMsg = "↪️ Exporting attendance record to points ledger...";
  const fullMsg = (returnObj.msg ? `${returnObj.msg}\n\n` : '') + firstMsg;

  // Execute Function with argument
  const functionName = transferSubmissionToLedger.name;
  confirmAndRunUserChoice_(functionName, fullMsg, selectedRow);
}


function importAppRecordUI_() {
  const returnObj = requestRowInput_();  // returnObj = {row : int, msg : string}
  const selectedRow = returnObj.row;

  if (!selectedRow) return;  // User input is not valid

  // Assemble notification message
  const firstMsg = "↪️ Importing attendance record from McRUN app...";
  const fullMsg = (returnObj.msg ? `${returnObj.msg}\n\n` : '') + firstMsg;

  // Execute Function with argument
  const functionName = transferThisRow_.name;
  confirmAndRunUserChoice_(functionName, fullMsg, selectedRow);
}


/**
 * Helper Functions for user input, etc.
 */

function requestRowInput_() {
  const ui = SpreadsheetApp.getUi();
  const headerMsg = "Which row do you want to target?";
  const textMsg = "Enter the row number, or leave empty for the last row.";

  const response = ui.prompt(headerMsg, textMsg, ui.ButtonSet.OK);
  const responseText = response.getResponseText().trim();

  return processRowInput_(responseText, ui);
}


/**
 * Returns result of reponse processing for row input.
 *
 * Helper function for UI functions for McRUN menu.
 *
 * @param {string} userResponse  User response text from `SpreadsheetApp.getUi().prompt`
 * @param {GoogleAppsScript.Base.Ui} ui  User interface in Google Sheets
 * @return {Result} `Result`  Packaged result of processing.
 * 
 * ### Properties of Return Object
 * - ```Result.row {integer}``` — Parsed integer value of `userResponse`.
 * 
 * - ```Result.msg {string}``` — Custom message to display to the user.
 * 
 * ### Metadata
 * @author [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>)
 * @date  Mar 24, 2025
 * @update  Mar 24, 2025
 */

function processRowInput_(userResponse, ui) {
  const rowNumber = Number.parseInt(userResponse);
  const returnObj = { row: null, msg: '' };

  if (userResponse === "") {
    // User did not enter a row number; check last row only.
    returnObj.row = '';
    returnObj.msg = "This will only target the last row.";
  }
  else if (isValidRow_(rowNumber)) {
    // Row is valid, can continue execution.
    returnObj.row = rowNumber;
    returnObj.msg = `This will only target row ${rowNumber}.`
  }
  else {
    // Input value is invalid row. Stop execution.
    ui.alert("Incorrect row number, please try again with a valid row number.");
  }

  return returnObj;
}


/**
 * Returns true if row is int and found in `ATTENDANCE_SHEET`.
 *
 * Helper function for UI functions for McRUN menu.
 *
 * @param {number}  The row number in `ATTENDANCE_SHEET` 1-indexed.
 * @return {boolean}  Returns true if valid row in sheet.
 *
 * @author [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>) & ChatGPT
 * @date  Dec 6, 2024
 * @update  Dec 6, 2024
 */

function isValidRow_(row) {
  const sheet = ATTENDANCE_SHEET;
  const lastRow = sheet.getLastRow();
  const rowInt = parseInt(row);

  return (Number.isInteger(rowInt) && rowInt >= 0 && rowInt <= lastRow);
}


/**
 * Verifies that `SCRIPT_PROPERTY` bank matches script properties in 'Project Settings'.
 *
 * @author [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>)
 * @date  Dec 11, 2024
 * @update  Dec 11, 2024
 */

function checkValidScriptProperties() {
  const scriptProperties = PropertiesService.getScriptProperties();
  const keys = scriptProperties.getKeys();
  const userDefinedProperties = Object.values(SCRIPT_PROPERTY);

  const ui = SpreadsheetApp.getUi();
  const errorTitle = "⚠️ WARNING TO DEVELOPER! ⚠️";

  // Verify same size in both property banks
  if (keys.length != userDefinedProperties.length) {
    let errorMessage = "Script Properties in 'Project Settings' does not match 'SCRIPT_PROPERTY' in Google Apps Script";
    ui.alert(errorTitle, errorMessage, ui.ButtonSet.OK);

    throw Error(errorMessage);
  }

  // Compare script properties in 'Project Settings' with user-defined 'SCRIPT_PROPERTY' object.
  keys.forEach(key => {
    let isIncluded = userDefinedProperties.includes(key);
    if (!isIncluded) {
      let errorMessage = `\`${key}\` in 'Project Settings' is not found in 'SCRIPT_PROPERTY' in Google Apps Script`;
      ui.alert(errorTitle, errorMessage, ui.ButtonSet.OK);

      throw Error(errorMessage);
    }
  });
}
