/**
 * Adds `Google Form` as source of attendance submission. 
 * 
 * Sets sendEmail column to `true` so emailSubmission() can proceed.
 * 
 * @trigger New head run or mcrun event submission.
 * 
 * @author [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>)
 * @date  Oct 17, 2023
 * @update  Oct 17, 2023
 * 
 */

function addMissingFormInfo() {
  const sheet = ATTENDANCE_SHEET;
  const lastRow = sheet.getLastRow();

  const rangePlatform = sheet.getRange(lastRow, PLATFORM_COL);
  rangePlatform.setValue('Google Form');

  const rangeSendEmail = sheet.getRange(lastRow, IS_COPY_SENT_COL);   // cell for email confirmation
  rangeSendEmail.setValue(true);
  rangeSendEmail.insertCheckboxes();
}


/**
 * Change attendance status of all members to not present.
 * 
 * Helper function for `consolidateMemberData()`.
 * 
 * @trigger New head run or mcrun event submission.
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
  
  for (var i=0; i < rangeList.length; i++){
    if (rangeList[i].getName() == "attendanceStatus") {
      rangeAttendance = rangeList[i];
      break;
    }
  }

  rangeAttendance.getRange().uncheck(); // remove all Presence checks
}


/**
 * Formats certain columns of `HR Attendance` sheet.
 *
 * @trigger New Google form or app submission.
 *
 * @author [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>)
 * @date  Oct 9, 2023
 * @update  Oct 23, 2024
 */

function formatSpecificColumns() {
  const sheet = ATTENDANCE_SHEET;
  
  const rangeListToBold = sheet.getRangeList(['A2:A', 'D2:D', 'L2:M']);
  rangeListToBold.setFontWeight('bold');  // Set ranges to bold

  const rangeListToWrap = sheet.getRangeList(['B2:G', 'I2:K']);
  rangeListToWrap.setWrap(true);  // Turn on wrap

  const rangeAttendees = sheet.getRange('F2:H');
  rangeAttendees.setFontSize(9);  // Reduce font size for `Attendees` column

  const rangeHeadRun = sheet.getRange('D2:D');
  rangeHeadRun.setFontSize(11);   // Increase font size for `Head Run` column

  const rangeListToCenter = sheet.getRangeList(['L2:M']); 
  rangeListToCenter.setHorizontalAlignment('center');
  rangeListToCenter.setVerticalAlignment('middle');   // Center and align to middle

  const rangePlatform = sheet.getRange('M2:M');
  rangePlatform.setFontSize(11);  // Increase font size for `Submission Platform` column

  // Gets non-empty range
  const range = sheet.getRange(1, 1, sheet.getLastRow(), sheet.getLastColumn());
  range.getBandings().forEach(banding => banding.remove());   // Need to remove current banding, before applying it to current range
  range.applyRowBanding(SpreadsheetApp.BandingTheme.BLUE, true, true);    // Apply BLUE banding with distinct header and footer colours.

  //const rangeSendEmail = sheet.getRange('L2:L');   // cells for email confirmation
  //rangeSendEmail.insertCheckboxes();
}


/**
 * Formats attendee names from 'row' into uniform view, sorted and separated by newline.
 *
 * @param {integer} row  The index in the `HR Attendance` sheet.
 * 
 * @author [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>)
 * @date  Oct 24, 2024
 * @update  Oct 24, 2024
 * 
 * ```javascript
 * // Sample Script ➜ Format names in row `13`.
 * const rowToFormat = 13;
 * formatNamesInRow(rowToFormat);
 * ```
 */

function formatNamesInRow(row) {
  const sheet = ATTENDANCE_SHEET;
  var nameRange = sheet.getRange(row, ATTENDEES_BEGINNER_COL, 1, 3);  // Attendees columns
  var namesArr = nameRange.getValues()[0];    // 1D Array of size 3 (Beginner, Intermediate, Advanced)

  for (var i = 0; i < namesArr.length; i++) {
    // Only process non-empty cells
    if (namesArr[i]) { 
      // Replace "n/a" (case insensitive) with "None"
      var cellValue = namesArr[i].replace(/n\/a/gi, "None");
        
      // Split by commas or newline characters
      var names = cellValue.split(/[,|\n]+/); 

      // Remove whitespace and capitalize names
      var formattedNames = names.map(function(name) {
        return name
              .trim()
              .toLowerCase()
              .replace(/\b\w/g, function(l) { return l.toUpperCase(); });
      });
        
      // Sort names alphabetically
      formattedNames.sort();
        
      // Join back with newline characters
      namesArr[i] = formattedNames.join('\n');
      }
    }
  
  // Replace values with formatted names
  nameRange.setValues([namesArr]);    // setValues requires 2D array
}


