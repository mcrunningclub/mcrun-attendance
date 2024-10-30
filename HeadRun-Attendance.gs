/**
 * Functions to execute after form submission.
 * 
 * @trigger Form Submission.
 */

function onFormSubmission() {
  addMissingFormInfo();
  formatLastestNames();
  //emailSubmission();    // IN-REVIEW
  formatSpecificColumns();
  //copyToLedger();       // IN-REVIEW
}


/**
 * Functions to execute after McRUN app submission.
 * 
 * @trigger McRUN App Submission.
 */
function onAppSubmission() {
  removePresenceChecks();
  formatLastestNames();
  //emailSubmission();    // IN-REVIEW
  formatSpecificColumns();
  //copyToLedger();       // IN-REVIEW
}


/**
 * Verifies if new submission in `HR Attendance` sheet from app.
 * 
 * Since app cannot create a trigger when submitting, `onAppSubmission()` 
 * will only run if latest submission is less than `[timeLimit]` old. 
 * 
 * @trigger Edit time in `HR Attendance` sheet.
 *  
 * @author [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>)
 * @date  Oct 17, 2023
 * @update  Oct 17, 2023
 */

function onEditCheck() {
  Utilities.sleep(5*1000);  // Let latest submission sync for 5 seconds

  const sheet = ATTENDANCE_SHEET;
  const timestamp = new Date(sheet.getRange(sheet.getLastRow(), 1).getValue());
  const timeLimit = 45 * 1000   // 45 sec = 45 * 1000 millisec

  if (Date.now() - timestamp > timeLimit) return;   // exit if over time limit

  const latestPlatform = sheet.getRange(sheet.getLastRow(), PLATFORM_COL).getValue();

  if (latestPlatform == "McRUN App") {
    onAppSubmission();
  }
}


/**
 * Consolidate multiple submission of same headrun into single row.
 * 
 * CURRENTLY IN REVIEW!
 * 
 * @author [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>) & ChatGPT
 * @date  Oct 24, 2024
 * @update  Oct 24, 2024
 */

function consolidateSubmissions() {
  const sheet = ATTENDANCE_SHEET;

  // Data range, assuming headers are on row 1
  var dataRange = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()); 
  var data = dataRange.getValues();

  // Create a map to store consolidated rows
  var consolidated = {};
  
  for (var i = 0; i < data.length; i++) {
    var date = new Date(data[i][0]).toDateString(); // Convert the timestamp in column A (1st column) to a date string
    var matchString = data[i][3]; // Column D (4th column)
    
    var key = date + '|' + matchString; // Create a unique key for matching
    
    // If a row with the same key already exists, consolidate the data
    if (consolidated[key]) {
      for (var j = 1; j < data[i].length; j++) {
        if (j !== 3) { // Skip column D since we're matching based on it
          // Concatenate the new data, separated by commas or newlines, avoiding duplicates
          if (consolidated[key][j] && data[i][j]) {
            consolidated[key][j] += ', ' + data[i][j];
          } else if (data[i][j]) {
            consolidated[key][j] = data[i][j];
          }
        }
      }
    } else {
      // If no matching row exists, store the current row as it is
      consolidated[key] = data[i];
    }
  }
  
  // Clear the existing data and set the consolidated rows
  sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).clear(); // Clear data range
  
  var newData = Object.values(consolidated); // Convert the consolidated object to an array of rows
  sheet.getRange(2, 1, newData.length, newData[0].length).setValues(newData); // Write the new consolidated data
}


/**
 * Copy newest attendance submission to ledger spreadsheet.
 * 
 * CURRENTLY IN REVIEW!
 * 
 * @author [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>) & ChatGPT
 * @date  Oct 30, 2023
 * @update  Oct 29, 2024
 */

function copyToLedger() {
  const sourceSheet = ATTENDANCE_SHEET;
  const ledgerName = LEDGER_NAME;
  const sheetUrl = LEDGER_URL;

  var destinationSpreadsheet = SpreadsheetApp.openByUrl(sheetUrl);
  var destinationSheet = destinationSpreadsheet.getSheetByName(ledgerName);
  var sourceData = sourceSheet.getRange(sourceSheet.getLastRow(), 1, 1, 5).getValues()[0];

  destinationSheet.appendRow(sourceData);
}


/**
 * Send a copy of attendance submission to headrunners, President & VP Internal.
 * 
 * Attendees are separated by level
 * 
 * @trigger Attendance submissions.
 *  
 * @author [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>)
 * @date  Oct 9, 2023
 * @update  Oct 30, 2024
 */

function emailSubmission() {
  // Error Management: prevent wrong user sending email
  //if ( getCurrentUserEmail() != 'mcrunningclub@ssmu.ca' ) return;   // REMOVE AFTER TESTING !

  const sheet = ATTENDANCE_SHEET;
  const lastRow = sheet.getLastRow();
  const latestSubmission = sheet.getRange(lastRow, 1, 1, sheet.getLastColumn());

  // Save values in 0-indexed array, then transform into 1-indexed by appending empty
  // string to the front. Now, access is easier e.g [EMAIL_COL] vs [EMAIL_COL-1]
  const values = latestSubmission.getValues()[0];
  values.unshift("");   // append "" to front

  var timestamp = new Date(values[TIMESTAMP_COL]);
  const formattedDate = Utilities.formatDate(timestamp, TIMEZONE, 'MM/dd/yyyy');

  // Replace newline with comma-space
  const allAttendees = [
    ATTENDEES_BEGINNER_COL, 
    ATTENDEES_INTERMEDIATE_COL, 
    ATTENDEES_ADVANCED_COL
  ].map(level => values[level].replaceAll('\n', ', '));

  // Read only (cannot edit values in sheet)
  const headrun = { 
    name:         values[HEADRUN_COL],
    distance:     values[DISTANCE_COL],
    attendees:    allAttendees,
    toEmail:      values[EMAIL_COL],
    confirmation: values[CONFIRMATION_COL],
    notes:        values[COMMENTS_COL]
  };

  // Read and edit sheet values
  const rangeConfirmation = sheet.getRange(lastRow, CONFIRMATION_COL);
  const rangeIsCopySent = sheet.getRange(lastRow, IS_COPY_SENT_COL);
  
  // Only send if submitter wants copy && email has not been sent yet
  if(rangeIsCopySent.getValue()) return;

  // Replace newline delimiter with comma-space if non-empty or matches "None"
  const attendeesStr = headrun.attendees.toString();

  if(attendeesStr.length > 1) {
    headrun.attendees = attendeesStr.replaceAll('\n', ', ');
  }
  else headrun.attendees = 'None';  // otherwise replace empty string by 'none'
  
  headrun.confirmation = (headrun.confirmation ? 'Yes' : 'No (explain in comment section)' );
  rangeConfirmation.setValue(headrun.confirmation);

  
  const emailBodyHTML = createEmailCopy(headrun);

  var message = {
    to: headrun.toEmail,
    bcc: emailPresident,
    cc: "mcrunningclub@ssmu.ca" + ", " + emailVPinternal,
    subject: "McRUN Attendance Form (" + formattedDate + ")",
    htmlBody: emailBodyHTML,
    noReply: true,
    name: "McRUN Attendance Bot"
  }

  //MailApp.sendEmail(message);   // REMOVE AFTER TEST!

  // As of Oct 2024, MailApp is void and cannot return a confirmation.
  // Assume sent and set isCopySend as true.
  rangeIsCopySent.setValue(true);
  rangeIsCopySent.insertCheckboxes();
}


/**
 * Check for missing submission after scheduled headrun.
 * 
 * @trigger 30-60 mins after schedule in `getHeadRunString()`.
 * 
 * CURRENTLY IN REVIEW!
 * 
 * @UPDATE-EACH-SEMESTER `getHeadRunString()`, `getHeadRunnerEmail()`.
 * 
 * @TODO  change to `checkMissingAttendance(var headRunAMorPM)`.
 * @TODO  modify headrun source to GSheet-> i.e. modify `getHeadRunString()` and `getHeadRunnerEmail()`.
 * 
 * @author [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>)
 * @date  Oct 15, 2023
 * @update  Oct 10, 2024
 */

function checkMissingAttendance() {
  const sheet = ATTENDANCE_SHEET;
  
  // Gets values of all timelogs
  var numRows = sheet.getLastRow() - 1;
  var submissionDates = sheet.getRange(2, TIMESTAMP_COL, numRows).getValues();
  var submissionHeadRuns = sheet.getRange(2, HEADRUN_COL, numRows).getValues();

  // Get date at trigger time and compare with timestamp of existing submissions
  const today = new Date();
  const formattedToday = Utilities.formatDate(today, TIMEZONE, 'yyyy-MM-dd a');   // e.g. '2024-10-27 PM'

  // Formats trigger datetime to get head runner emails
  const headRunDay = Utilities.formatDate(today, TIMEZONE, 'EEEEa');  // e.g. 'MondayAM'
  const headRunDetail = getHeadRunString(headRunDay); // e.g 'Monday - 9am'

  // Error handling
  if (headRunDetail.length <= 0) {
    // Create an instance of ExecutionError with a custom message
    var errorMessage = "No headrunner has been found for " + headRunDay;    
    throw new Error(errorMessage); // Throw the ExecutionError
  }

  // Start checking from end of head run attendance submissions
  // Exit loop when submission found or until list exhausted
  var isSubmitted = false;

  for(var i = numRows-1; i>= 0 && !isSubmitted; i--) {
    var submissionDate = Utilities.formatDate(new Date(submissionDates[i]), TIMEZONE, 'yyyy-MM-dd a');

    // Get detailed head run to compare with today's headRunString
    if (submissionDate === formattedToday) {
      var submissionHeadRun = submissionHeadRuns[i].join();
      isSubmitted = (submissionHeadRun === headRunDetail);
    }
  }

  // Verify if form has been submitted. Otherwise send an email reminder.
  if (isSubmitted) return;    

  // Get head runners email using target headrun
  const headRunnerEmail = getHeadRunnerEmail(headRunDay).join();

  const remainderEmailBodyHTML = " \
  <html> \
    <head> \
      <title>Missing Submission Form</title> \
    </head> \
    <body> \
      <p> \
        Hi, \
      </p> \
      <p> \
        This is a friendly reminder to submit today's headrun attendance. \
      </p> \
      <p> \
        <strong>Log attendance using the McRUN Appm or click <a href= https://docs.google.com/forms/d/e/1FAIpQLSf_4zdnyY4I4vSxaatAaxxgsU38hb862arFDU9wTbSpnoXdKA/viewform\> here</a> to access the F24 attendance form or </strong> \
      </p> \
      <p> \
        Please ignore this message if the headrun has been cancelled or your group has already submitted the attendance form. \
      </p> \
      <p> \
        <br> \
        Thank you for all your help! McRun only runs because of you.\
      </p> \
      <p> \
        <br> \
        - McRUN Bot \
      </p> \
      <br> \
    </body> \
  </html>";

  var reminderEmail = {
    to: headRunnerEmail,
    bcc: emailPresident,
    cc: "mcrunningclub@ssmu.ca" + ", " + emailVPinternal,
    subject: "McRUN Missing Attendance Form - " + headRunDetail,
    htmlBody: remainderEmailBodyHTML,
    noReply: true,
    name: "McRUN Attendance Bot"
  }

  MailApp.sendEmail(reminderEmail);
  return;

  // Date formatting examples
  const todayWeekDay = Utilities.formatDate(today, TIMEZONE, 'EEEE');
  const todayDate = Utilities.formatDate(today, TIMEZONE, 'dd');
}

/**
 * Find attendees that are unregistered members.
 * 
 * CURRENTLY IN REVIEW!
 * 
 * @author [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>)
 * @date  Oct 30, 2024
 * @update  Oct 30, 2024
 */

function getUnregisteredMembers(){
  const sheet = ATTENDANCE_SHEET;
  const NAMES_NOT_FOUND_COL = 14;
  const lastRow = sheet.getLastRow();

  const unfoundNameRange = sheet.getRange(lastRow, NAMES_NOT_FOUND_COL);

  const numColToGet = LEVEL_COUNT;

  // Get attendee names starting from beginner col
  const nameRange = sheet.getRange(lastRow, ATTENDEES_BEGINNER_COL, 1, numColToGet);  // Attendees columns

  // 1D Array of size 3 (Beginner, Intermediate, Advanced)
  const allNames = nameRange.getValues()[0];

  // Split names in each level into separate entry in array, then sort
  const sortedNames = allNames
    .filter(level => !level.includes("None")) // Skip levels with "none"
    .flatMap(level => level.split('\n'))
    .sort();

  const memberSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Members");
  const rangeMembers = memberSheet.getRange(2, 1, 223);
  const members = rangeMembers.getValues().map(row => row[0]); // Get sorted member names in a 1D array

  const cannotFindArray = [];
  let index = 0;

  // Iterate through sortedNames to find matching members
  sortedNames.forEach(attendee => {
    while (index < members.length) {
      const memberName = members[index];

      if (attendee === memberName) {
        index++; // Move to next member in sorted order
        return;  // Continue to the next attendee
      } 
      else if (attendee < memberName) {
        cannotFindArray.push(attendee); // Attendee not in members
        return;  // Continue to the next attendee
      }

      index++;
    }

    // If index exceeds members, remaining attendees are unfound
    if (index >= members.length) cannotFindArray.push(attendee);
  });

  // Log unfound names
  unfoundNameRange.setValue(cannotFindArray.join(", "));

}


/**
 * Helper function to find unmatched attendees
 * 
 * CURRENTLY IN REVIEW!
 * 
 * @author [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>)
 * @date  Oct 30, 2024
 * @update  Oct 30, 2024
 */


function findUnmatchedAttendees(attendees, members) {
  const cannotFindArray = [];
  let index = 0;

  attendees.forEach(attendee => {
    while (index < members.length) {
      const memberName = members[index];

      if (attendee === memberName) {
        Logger.log("Found: " + attendee);
        index++; // Move to next member in sorted order
        return;  // Continue to the next attendee
      } else if (attendee < memberName) {
        cannotFindArray.push(attendee); // Attendee not in members
        return;  // Continue to the next attendee
      }

      index++;
    }

    // If index exceeds members, remaining attendees are unfound
    if (index >= members.length) cannotFindArray.push(attendee);
  });

  return cannotFindArray;
}


/** 
 * @author ChatGPT
 */
function copyRowToAnotherSpreadsheet_() {
  var sourceSpreadsheet = SpreadsheetApp.openById("SourceSpreadsheetID"); // Replace with the ID of your source spreadsheet
  var destinationSpreadsheet = SpreadsheetApp.openById("DestinationSpreadsheetID"); // Replace with the ID of your destination spreadsheet

  var sourceSheet = sourceSpreadsheet.getSheetByName("SourceSheetName"); // Replace with the name of your source sheet
  var destinationSheet = destinationSpreadsheet.getSheetByName("DestinationSheetName"); // Replace with the name of your destination sheet

  var rowIndexToCopy = 2; // Replace with the row index you want to copy (e.g., row 2)
  var sourceData = sourceSheet.getRange(rowIndexToCopy, 1, 1, sourceSheet.getLastColumn()).getValues();

  destinationSheet.appendRow(sourceData[0]);
}


function deadCode_() {
  return;

  var emailBody = 
    "Here is a copy of your submission: \n\n- HEAD RUN: " + headRun + "\n- DISTANCE: " + distance + "\n\n------ ATTENDEES ------\n" + attendees + "\n\n*I declare all attendees have provided their waiver and paid the one-time member fee*  > " + confirmation + "\n\nComments: " + notes + "\n\nKeep up the amazing work!\n\nBest,\nMcRUN Team"
  ;

  var headRunTime = getHeadRunTime(todayWeekDay);
  if(headRunTime.length < 1) return;  // exit if no head run today

  var dateTime, thresholdTime;

  for(const time of headRunTime) {
    dateTime = getDateTime(time);   // convert to Date object
    thresholdTime = getThresholdTime(time);  // add 2 hours

    Logger.log(today);
    Logger.log(thresholdTime);

    if (today.setHours(today.getHours() + 2) ) {};
  }

  var test = new Date(submissionDates[0]).getDate();

  for (var i = 0; i < data.length; i++) {
    var cellValue = data[i][0];
    if (cellValue instanceof Date) {
      cellValue.setHours(0, 0, 0, 0); // Set the time to midnight for comparison
      if (cellValue.getTime() !== today.getTime()) {
        // If the date in the cell doesn't match today's date
        // Send a notification email
      }
    }
  }

  headRunTime.forEach(
    function(item) { Logger.log(item); }
  );


  function getDateTime(timeString) {
    var dateTime = new Date();

    var parts = timeString.split(':');
    var hours = parseInt(parts[0], 10);
    var minutes = parseInt(parts[1], 10);

    dateTime.setHours(hours, minutes, 0, 0); // Set the time

    return dateTime;
  }


  function getThresholdTime(startTime) {
    var dateTime = new Date();

    var parts = startTime.split(':');
    var hours = parseInt(parts[0], 10);
    var minutes = parseInt(parts[1], 10);

    dateTime.setHours(hours + 2, minutes, 0, 0); // Set the time
    return dateTime;
  }
}