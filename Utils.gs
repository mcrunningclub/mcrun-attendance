

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
 */

function logAsAC_(msg, funcName = "", useLogger = true) {
  const message = `[AC#${funcName}] ${msg}`;
  useLogger ? Logger.log(message) : console.log(message);
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