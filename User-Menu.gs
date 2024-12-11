/**
 * Users authorized to use the McRUN menu.
 * 
 * Prevents unwanted data overwrite in Gsheet.
 */
const PERM_USER_ = [
  'mcrunningclub@ssmu.ca',
  'ademetriou8@gmail.com',
  'andreysebastian10.g@gmail.com',
  'gagnonjikael@gmail.com',
  'thecharlesvillegas@gmail.com',
];


/**
 * Log user attempting to use custom McRUN menu.
 * 
 * If input empty, then extract email using `getCurrentUserEmail_()`.
 * 
 * @trigger User choice in custom menu.
 * 
 * @param {string} [email=""]  Email of active user. 
 *                             Defaults to empty string.
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
 * Creates custom menu to run frequently used scripts in Google App Script.
 * 
 * Extracting function name using `name` property to allow for refactoring.
 * 
 * Cannot check if user authorized here, or custom menu will not be 
 * displayed due to Google Apps Script limitation.
 * 
 * @trigger Open Google Spreadsheet.
 * 
 * @author [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>)
 * @date  Dec 5, 2024
 * @update  Dec 8, 2024
 */

function onOpen() {
  const ui = SpreadsheetApp.getUi();

  ui.createMenu('🏃‍♂️ McRUN Attendance Menu')
    .addItem('📢 Custom menu. Click for help.', helpUI_.name)
    .addSeparator()

    .addSubMenu(ui.createMenu('Formatting Menu')
      .addItem('Sort By Timestamp', sortByTimestampUI_.name)
      .addItem('Format Sheet View', formatSheetUI_.name)
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
  .addToUi();

  checkValidScriptProperties(); // verify validity of `SCRIPT_PROPERTY`

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
 * Boiler plate function to display custom UI to user.
 * 
 * Executes function `functionName` with optional argument `funcArg`.
 * 
 * Verifies if user is authorized before executing script.
 * 
 * @trigger User choice in custom menu.
 * 
 * @param {string}  functionName  Name of function to execute.
 * @param {string}  [additionalMsg=""]  Custom message for executing function.
 *                                      Defaults to empty string.
 * @param {string}  [funcArg=""]  Function argument to pass with `functionName`.
 *                                Defaults to empty string.
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

  // Check if authorized user to prevent illegal execution
  if (!PERM_USER_.includes(userEmail)) {
    const warningMsgHeader = "🛑 You are not authorized 🛑"
    const warningMsgBody = "Please contact the exec team if you believe this is an error.";

    ui.alert(warningMsgHeader, warningMsgBody, ui.ButtonSet.OK);
    return;
  }

  // Continue execution if user is authorized
  var message = `
    ⚙️ Now executing ${functionName}().
  
    🚨 Press cancel to stop.
  `;

  // Append additional message if non-empty
  message += additionalMsg ? `\n🔔 ${additionalMsg}` : "";

  // Save user response
  const response = ui.alert(message, ui.ButtonSet.OK_CANCEL);
  let retValue = "";

  if (response == ui.Button.OK) {
    if (funcArg) {
      retValue = this[functionName](funcArg);   // executing function `functionName` with arg
    }
    else {
      retValue = this[functionName]();   // executing function with name `functionName` w/o arg
    }
  }
  else {
    ui.alert('Execution cancelled...');
  }

  // Log attempt in console using active user email
  logMenuAttempt_(userEmail);

  // Return value from executed function if required
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

function formatSheetUI_() {
  const functionName = formatSpecificColumns.name;
  const customMsg = "This will only modify the sheet formatting (view), and improve confirmation value."
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
  const ui = SpreadsheetApp.getUi();
  const headerMsg = "Which row do you want to format?";
  const textMsg = "Enter the row number, or leave empty for the last row.";

  const response = ui.prompt(headerMsg, textMsg, ui.ButtonSet.OK);
  const responseText = response.getResponseText().trim();

  var functionName = formatNamesInRow_.name;
  var customMsg = "";

  const rowNumber = Number.parseInt(responseText);

  if (responseText == "") {
    // User did not enter a row number; check last row only.
    customMsg = "This will only format the names in the last row."
  }
  else if (isValidRow_(rowNumber)) {
    // Row is valid, can continue execution.
    customMsg = `This will only format the names in row ${rowNumber}.`
  }
  else {
    // Input value is invalid row. Stop execution.
    ui.alert("Incorrect row number, please try again with a valid row number.");
    return;
  }

  // Run respective function depending if-statement above
  confirmAndRunUserChoice_(functionName, customMsg, rowNumber);
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
  functionName = getAllUnregisteredMembers.name;
  const customMsg = "This will search for *all* unregistered attendees, and add them to `Not Found` column. \
  \n\nWARNING! Wide-sheet formatting will may take some time."

  confirmAndRunUserChoice_(functionName, customMsg);
}

/**
 * This UI function can target a specific row, or the last row if input is omitted.
 */

function findUnregisteredAttendeesUI_() {
  const ui = SpreadsheetApp.getUi();
  const headerMsg = "Which row do you want to check?";
  const textMsg = "Enter the row number, or leave empty to check last row.";

  const response = ui.prompt(headerMsg, textMsg, ui.ButtonSet.OK);
  const responseText = response.getResponseText().trim();
  const rowNumber = Number.parseInt(responseText);

  var functionName = getUnregisteredMembers_.name;    // Function searches last row if no arg provided
  var customMsg = "";

  if (responseText == "") {
    // User did not enter a row number; check last row only.
    customMsg = "This will search for unregistered attendees in the last row, and add them to \`Not Found\` column."
  }
  else if (isValidRow_(rowNumber)) {
    // Row is valid, can continue execution.
    customMsg = `This will search for unregistered attendees in row ${responseText} and add them to \`Not Found\` column.`
  }
  else {
    // Input value is invalid row. Stop execution.
    ui.alert("Incorrect row number, please try again with a valid row number.");
    return;
  }

  // Run respective function depending if-statement result above
  confirmAndRunUserChoice_(functionName, customMsg, rowNumber);
}


/** 
 * Scripts for `Triggers` menu items.
 * 
 * Extracting function name using `name` property to allow for refactoring.
 * 
 * Some functions include a custom message for user.
 */

function onFormSubmitUI_() {
  const functionName = onFormSubmission.name;
  const customMsg = "This will run functions triggered when new attendance is submitted using the Google Form."
  confirmAndRunUserChoice_(functionName, customMsg);
}

function onAppSubmitUI_() {
  const functionName = onAppSubmission.name;
  const customMsg = "This will run functions triggered when new attendance is submitted using the app."
  confirmAndRunUserChoice_(functionName, customMsg);
}


/**
 * Returns true if row is int and found in `ATTENDANCE_SHEET`.
 * 
 * Helper function for UI functions for McRUN menu.
 * 
 * @param {number}  The row number in `ATTENDANCE_SHEET` 1-indexed.
 * 
 * @author [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>) & ChatGPT
 * @date  Dec 6, 2024
 * @update Dec 6, 2024
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
 * @author [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>) & ChatGPT
 * @date  Dec 11, 2024
 * @update Dec 11, 2024
 */

function checkValidScriptProperties() {
  const scriptProperties = PropertiesService.getScriptProperties();
  const keys = scriptProperties.getKeys();
  const userDefinedProperties = Object.values(SCRIPT_PROPERTY);
  
  const ui = SpreadsheetApp.getUi();
  const errorTitle = "⚠️ WARNING TO DEVELOPER! ⚠️";

  // Verify same size in both property banks
  if(keys.length != userDefinedProperties.length) {
    let errorMessage = "Script Properties in 'Project Settings' does not match 'SCRIPT_PROPERTY' in Google Apps Script";
    ui.alert(errorTitle, errorMessage, ui.ButtonSet.OK);
    
    throw Error(errorMessage);
  }

  // Compare script properties in 'Project Settings' with user-defined 'SCRIPT_PROPERTY' object.
  keys.forEach(key => {
    let isIncluded = userDefinedProperties.includes(key);
    if(!isIncluded) {
      let errorMessage = `\`${key}\` in 'Project Settings' is not found in 'SCRIPT_PROPERTY' in Google Apps Script`;
      ui.alert(errorTitle, errorMessage, ui.ButtonSet.OK);

      throw Error(errorMessage);
    }
  });
}
