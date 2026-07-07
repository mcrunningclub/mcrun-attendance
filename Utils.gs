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
  const userEmail = Session.getActiveUser().toString();
  logAsAC_(`Current user email '${userEmail}'`, getCurrentUserEmail_.name);
  return userEmail;
}

/**
 * Logs message in a standard and comprehensible format.
 * @param {string} msg  Message to log
 * @param {string} funcName  Name of the function to log if applicable. Defaults to ""
 * @param {boolean} useLogger  If true, use the Logger class, otherwise use console
 */

function logAsAC_(msg, funcName = "", useLogger = true) {
  const message = `[AC#${funcName}] ${msg}`;
  useLogger ? Logger.log(message) : console.log(message);
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
 * Find row index of last submission in reverse using while-loop.
 * 
 * Used to prevent native `sheet.getLastRow()` from returning empty row.
 * 
 * @param {Spreadsheet.sheet} [sheet = GET_ATTENDANCE_SHEET_()] Target sheet
 * @return {integer}  Returns 1-index of last row in GSheet.
 *  
 * @author [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>)
 * @date  Feb 8, 2025
 * @update  Apr 10, 2025
 */

function getLastRow_(sheet = GET_ATTENDANCE_SHEET_()) {
  const startRow = 1;
  const numRow = sheet.getLastRow();

  // Fetch all values in the TIMESTAMP_COL
  const values = sheet.getSheetValues(startRow, SEM_ATTENDANCE_COLS.TIMESTAMP, numRow, 1);
  let lastRow = values.length;

  // Loop through the values in reverse order
  while (values[lastRow - 1][0] === "") {
    lastRow--;
  }

  return lastRow;
}


/** 
 * Format timestamp to format as `yyyy-MM-dd hh:mm:ss`.
 * 
 * Raw format cannot be understood by GSheet.
 * 
 * @param {string} raw  Datetime value to be formatted.
 * @return {Date}  A Date object with correct format.
 * 
 * @author [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>)
 * @date  Feb 9, 2025
 * @update  Feb 10, 2025
 */

function formatTimestamp_(raw) {
  const date = new Date(raw);
  const options = {
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


/** 
 * Compare the input timestamps.
 * 
 * @param {string} timestamp1  Timestamp 1
 * @param {string} timestamp2  Timestamp 2
 * 
 * @return {Boolean}  Returns result of comparaison.
 * 
 * @author [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>)
 * @date  Mar 20, 2025
 * @update  Mar 20, 2025
 */

function isSameTimestamp_(timestamp1, timestamp2) {
  const ts1 = (timestamp1 instanceof Date) ? timestamp1 : new Date(timestamp1);
  const ts2 = (timestamp2 instanceof Date) ? timestamp2 : new Date(timestamp2);
  return ts1.getTime() === ts2.getTime();
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