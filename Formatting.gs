/**
 * Adds `Google Form` as source of attendance submission. 
 * 
 * Sets sendEmail column to `true` so emailSubmission() can proceed.
 * 
 * @trigger New Google Form submission.
 * 
 * @author [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>)
 * @date  Oct 17, 2023
 * @update  Oct 17, 2023
 * 
 */

function addMissingFormInfo() {
  const sheet = ATTENDANCE_SHEET;
  const lastRow = sheet.getLastRow();

  const rangeIsCopySent = sheet.getRange(lastRow, IS_COPY_SENT_COL);
  
  // Since GForm automatically sends copy to submitter, set isCopySent `true`.
  rangeIsCopySent
    .insertCheckboxes()
    .setValue(true);

  const rangePlatform = sheet.getRange(lastRow, PLATFORM_COL);
  rangePlatform.setValue('Google Form');
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
  
  // Sort all the way to the last row, without the header row
  const range = sheet.getRange(2, 1, numRows, numCols);
  
  // Sorts values by `Timestamp`
  range.sort([{column: 1, ascending: true}]);
  return;
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
}


/**
 * Wrapper function for `formatNamesInRow` for lastest submission.
 * 
 * @author [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>)
 * @date  Oct 30, 2024
 * @update  Oct 30, 2024
 */

function formatLastestNames() {
  const sheet = ATTENDANCE_SHEET;
  const lastRow = sheet.getLastRow();

  formatNamesInRow(lastRow);
}

/**
 * Wrapper function for `formatNamesInRow` for *ALL* submissions.
 * 
 * @author [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>)
 * @date  Dec 5, 2024
 * @update  Dec 5, 2024
 */

function formatAllNames() {
  const sheet = ATTENDANCE_SHEET;
  const lastRow = sheet.getLastRow();

  for(var row=lastRow; row < 0; row++) {
    formatNamesInRow(row);
  }
}


/**
 * Formats attendee names from `row` into uniform view, sorted and separated by newline.
 *
 * @param {integer} row  The index in the `HR Attendance` sheet (1-indexed).
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
  const numColToGet = LEVEL_COUNT;

  // Get attendee names starting from beginner col
  const nameRange = sheet.getRange(row, ATTENDEES_BEGINNER_COL, 1, numColToGet);  // Attendees columns
  var namesArr = nameRange.getValues()[0];    // 1D Array of size 3 (Beginner, Intermediate, Advanced)

  for (var i = 0; i < namesArr.length; i++) {
    var trimmedArr = namesArr[i].trim();
    Logger.log(trimmedArr.length);

    // Case 1: Cell is non-empty
    if (trimmedArr.length != 0) { 
      
      // Replace "n/a" (case insensitive) with "None"
      var cellValue = trimmedArr.replace(/n\/a/gi, "None");
        
      // Split by commas or newline characters
      var names = cellValue.split(/[,|\n]+/); 

      // Remove whitespace, strip accents and capitalize names
      var formattedNames = names.map(name => name
        .trim()
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
      namesArr[i] = "None";
    }

  }
  // Replace values with formatted names
  nameRange.setValues([namesArr]);    // setValues() requires 2D array
}


/**
 * Creates email using details from input `emailDetails`.
 *
 * @param {Map<string>} emailDetails  Information needed to populate email.
 * @return {string}  HTML code for email.
 * 
 * @author [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>)
 * @date  Oct 29, 2024
 * @update  Oct 29, 2024
 * 
 * ```javascript
 * // Sample Script ➜ Create email using info.
 * const emailDetails = { 
 *    name : 'Monday - 6pm',
 *    distance : '5km',
 *    attendees : 'Bob Burger, Marge Simpson, Mabel Pines',
 *    toEmail : 'Head Runner Alex',
 *    confirmation : false,
 *    notes : 'Bob will pay fee next time.'
 * };
 * var emailHTML = createEmailCopy(emailDetails);
 * ```
 */

function createEmailCopy(emailDetails) {
  // Check for non-empty key-value object
  const size = Object.keys(emailDetails).length;
  if (size != 6) return null;
  
  const emailBodyHTML = " \
  <html> \
    <head> \
      <title>Submission Details</title> \
    </head> \
    <body> \
      <p> \
        Hi, \
      </p> \
      <p> \
        Here is a copy of the latest submission: \
      </p> \
      <p> \
        <strong>&nbsp;&nbsp;&nbsp;Head Run: </strong>" + emailDetails.name + "\
      </p> \
      <p> \
        <strong>&nbsp;&nbsp;&nbsp;Distance: </strong>" + emailDetails.distance + "\
      </p> \
      <p>\
        <strong>&nbsp;&nbsp;&nbsp;Attendees: </strong>" + emailDetails.attendees + "\
      </p> \
      <p> \
        <strong>&nbsp;&nbsp;&nbsp;Submitted by: </strong> " + emailDetails.toEmail + "\
      </p> \
      <p> \
          &nbsp;&nbsp; \
        <strong><em>I declare all attendees have provided their waiver and paid the one-time member fee: </em></strong>" + 
          emailDetails.confirmation + " \
      </p> \
      <p> \
        <strong>&nbsp;&nbsp;&nbsp;Comments: </strong> " + emailDetails.notes + "\
      </p> \
      <p> \
        <br> \
        - McRUN Bot \
      </p> \
    </body> \
  </html>";
  
  return emailBodyHTML;

}


/**
 * Formats all entries in `names ` then sorts.
 * 
 * Removes whitespace, strip accents and capitalize names.
 *
 * @param {string[]} names  Array of names to format.
 * @return {string[]}  Sorted array of formatted names.
 * 
 * @author [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>)
 * @date  Nov 1, 2024
 * @update  Nov 1, 2024
 * 
 * ```javascript
 * // Sample Script ➜ Format, then sort names.
 * const rawNames = ["BOb burger", "Francine deBlé"];
 * const result = formatAndSortNames(rawNames);
 * Logger.log(result)  // [Bob Burger, Francine Deble]
 * ```
 */
function formatAndSortNames(names) {
  const formattedNames = names.map(name => 
    name
      .trim()
      .normalize("NFD").replace(/[\u0300-\u036f]/g, "")  // Strip accents
      .toLowerCase()
      .replace(/\b\w/g, l => l.toUpperCase()) // Capitalize each name
  );

  return formattedNames.sort();
}


