/**
 * Functions to execute after form submission.
 * 
 * @trigger Form Submission.
 */

function onFormSubmission() {
  addMissingFormInfo();
  formatLastestNames();
  getUnregisteredMembers();
  
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
  getUnregisteredMembers();
  
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
 * @update  Nov 1, 2024
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

  sortAttendanceForm();   // Sort after adding information to submission
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

  const reminderEmailBodyHTML = REMINDER_EMAIL_HTML;

  var reminderEmail = {
    to: headRunnerEmail,
    bcc: emailPresident,
    cc: "mcrunningclub@ssmu.ca" + ", " + emailVPinternal,
    subject: "McRUN Missing Attendance Form - " + headRunDetail,
    htmlBody: reminderEmailBodyHTML,
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
 * Find attendees in `row` of `ATTENDANCE_SHEET `that are unregistered members.
 * 
 * Sets unregistered members in `NOT_FOUND_COL`.
 * 
 * List of members found in `Members` sheet.
 * 
 * CURRENTLY IN REVIEW!
 * 
 * @param {number} [row=ATTENDANCE_SHEET.getLastRow()]  The row number in `ATTENDANCE_SHEET` 1-indexed.
 *                                                      Defaults to the last row in the sheet.
 * 
 * @author [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>) & ChatGPT
 * @date  Oct 30, 2024
 * @update  Nov 1, 2024
 */

function getUnregisteredMembers(row=ATTENDANCE_SHEET.getLastRow()){
  const sheet = ATTENDANCE_SHEET;

  const unfoundNameRange = sheet.getRange(row, NAMES_NOT_FOUND_COL);

  const numColToGet = LEVEL_COUNT;

  // Get attendee names starting from beginner col
  const nameRange = sheet.getRange(row, ATTENDEES_BEGINNER_COL, 1, numColToGet);  // Attendees columns

  // 1D Array of size 3 (Beginner, Intermediate, Advanced)
  const allNames = nameRange
    .getValues()[0]
    .filter(level => !level.includes("None")) // Skip levels with "none"
    .flatMap(level => level.split('\n'))    // Split names in each level into separate entry in array
  ;   

  // Remove whitespace, strip accents and capitalize names
  const sortedNames = formatAndSortNames(allNames);

  // Get existing member registry in `Members`
  const memberSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Members");

  const memberCount = memberSheet.getLastRow() - 1;   // Do not count header row
  const numCols = sheet.getLastColumn();
  
  var membersRange = memberSheet.getRange(2, 1, memberCount, numCols);

  // Get array of members' full name
  const members = membersRange.getValues()
    .map(row => row[2])     // Get member full names in a 1D array
    .filter(name => name)  // Remove empty rows
  ;

  const formattedMembers = formatAndSortNames(members);

  // Use the helper function on sorted items
  const unregistered = findUnregistered_(sortedNames, formattedMembers);
  unfoundNameRange.setValue(unregistered.join(", "));

  // Log unfound names
  unfoundNameRange.setValue(unregistered.join(", "));

  return;
}


/**
 * Helper function to find unregistered attendees
 * 
 * CURRENTLY IN REVIEW!
 * 
 * @param {string[]} attendees  All attendees of the head run.
 * @param {string} members  All registered members.
 * @return {string[]}  Returns attendees not found in `members`.
 * 
 * @TODO Move this to `Membership (Main)` and call as library
 * 
 * @author [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>) & ChatGPT
 * @date  Oct 30, 2024
 * @update  Nov 1, 2024
 */

function findUnregistered_(attendees, members) {
  const unregistered = [];
  let index = 0;

  attendees.forEach(attendee => {
    // Move through `members` array starting from the last found `index`
    while (index < members.length) {
      const memberName = members[index];

      if (attendee === memberName) {
        index++; // Move to the next member in sorted order
        return;  // Break out of the while loop, continue to next attendee

      } else if (attendee < memberName) {
        unregistered.push(attendee); // Attendee not in members
        return;  // Continue to the next attendee
      }

      index++;
    }

    // If index exceeds `members`, mark remaining attendees as unfound
    if (index >= members.length) unregistered.push(attendee);
  });

  return unregistered;
}


function deadCode_() {
  return;

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

}