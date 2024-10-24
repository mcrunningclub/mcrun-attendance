/* SHEET CONSTANTS */
const SHEET_NAME = 'HR Attendance F24';
const ATTENDANCE_SHEET = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);

const MEMBERSHIP_NAME = 'Fall 2024';
const MASTER_NAME = 'MASTER';
const TIMEZONE = getUserTimeZone();

// List of columns in SHEET_NAME
const TIMESTAMP_COL = 1;
const EMAIL_COL = 2;
const HEADRUNNERS = 3;
const HEADRUN_COL = 4;
const RUN_LEVEL_COL = 5;
const ATTENDEES_BEGINNER_COL = 6;
const ATTENDEES_INTERMEDIATE_COL = 7;
const ATTENDEES_ADVANCED_COL = 8;
const CONFIRMATION_COL = 9;
const DISTANCE_COL = 10;
const COMMENTS_COL = 11;
const IS_COPY_SENT_COL = 12;
const PLATFORM_COL = 13;


function getUserTimeZone() {
  return Session.getScriptTimeZone();
}

function getColumnPosition() {
  var rangeList = ATTENDANCE_SHEET.getNamedRanges();
  var dRange = ATTENDANCE_SHEET.getNamedRanges()[0].getRange();

  for (var i=0; i < rangeList.length; i++){
    Logger.log(rangeList[i].getName());
  }
}

function onFormSubmission() {
  addMissingFormInfo();
  //emailSubmission();    // IN-REVIEW
  formatSpecificColumns();
  copyToLedger();
}

function onAppSubmission() {
  removePresenceChecks();
  //emailSubmission();    // IN-REVIEW
  formatSpecificColumns();
  copyToLedger();
}

/**
 * @author: Andrey S Gonzalez
 * @date: Feb 9, 2024
 * @update: Feb 9, 2024
 * 
 * Returns email of current user.
 * Prevents incorrect account executing Google automations e.g. McRUN bot.
 * 
 */
function getCurrentUserEmail() {
  return Session.getActiveUser();
}


/**
 * @author: Andrey S Gonzalez
 * @date: Oct 17, 2023
 * @update: Oct 17, 2023
 * 
 * Since app cannot create a trigger when submitting, onAppSubmission() will only run if latest submission
 * is less than [timeLimit] old. 
 * 
 * Triggered at edit time. Flag is timeLimit
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
 * @author: Andrey S Gonzalez
 * @date: Oct 17, 2023
 * @update: Oct 17, 2023
 * 
 * Adds additional information when Google Form is used. Sets sendEmail column to `true` so emailSubmission() can proceed.
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
 * @author: Andrey S Gonzalez
 * @date: Oct 9, 2023
 * @update: October 23, 2024
 * 
 * Change attendance status of all members to not present. 
 * Triggered after new head run or mcrun event.
 */
function removePresenceChecks() {
  
  // `Membership Collected (main)` Google Form
  var ss = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/1qvoL3mJXCvj3m7Y70sI-FAktCiSWqEmkDxfZWz0lFu4/edit?usp=sharing");

  var sheet = ss.getSheetByName(MASTER_NAME);
  var rangeAttendance, rangeList = sheet.getNamedRanges();
  
  for (var i=0; i < rangeList.length; i++){
    if (rangeList[i].getName() == "attendanceStatus") {
      rangeAttendance = rangeList[i];
      break;
    }
  }

  rangeAttendance.getRange().uncheck(); // remove all Presence checks
}


/**
 * @author: Andrey S Gonzalez
 * @date: Oct 9, 2023
 * @update: Oct 23, 2024
 * 
 * Format certain columns. Triggered by form or app submission.
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
 * @author: Andrey S Gonzalez & ChatGPT
 * @date: Oct 24, 2024
 * @update: Oct 24, 2024
 * 
 * Format attendee names from 'row' into uniform view, sorted and separated by newline.
 * 
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


/**
 * @author: Andrey S Gonzalez & ChatGPT
 * @date: Oct 24, 2024
 * @update: Oct 24, 2024
 * 
 * Consolidate multiple submission of same headrun into single row
 * CURRENTLY IN PROGRESS
 * 
 */
function consolidateSubmissions() {
  const sheet = ATTENDANCE_SHEET;
  var dataRange = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()); // Data range, assuming headers are on row 1
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
 * @author: Andrey S Gonzalez
 * @date: Oct 9, 2023
 * @update: Feb 9, 2024
 * 
 * Send a copy of attendance submission to submitter & VP Internal. Triggered after app submission
 */
function emailSubmission() {
  // Error Management: prevent wrong user sending email
  //if ( getCurrentUserEmail() != 'mcrunningclub@ssmu.ca' ) return;

  const sheet = ATTENDANCE_SHEET;
  const lastRow = sheet.getLastRow();
  const latestSubmission = sheet.getRange(lastRow, 1, 1, sheet.getLastColumn());

  const values = latestSubmission.getValues()[0];

  var timestamp = new Date(values[0]);
  const formattedDate = Utilities.formatDate(timestamp, TIMEZONE, 'MM/dd/yyyy');

  // Read only (cannot edit values in sheet)
  const toEmail = values[2];
  const headRun = values[3];
  var attendees = values[4];     // stored as string with '\n' delimeters
  var formattedAttendees;           // stores formatted string in `attendees`
  var confirmation = values[5];
  const distance = values[6];
  const notes = values[7];

  // Read and edit sheet values
  const rangeConfirmation = sheet.getRange(lastRow, 6);
  const rangeSendEmail = sheet.getRange(lastRow, 9);
  const rangeCopyIsSent = sheet.getRange(lastRow, 10);

  rangeSendEmail.insertCheckboxes();
  
  // Only send if submitter wants copy && email has not been sent yet
  if(rangeCopyIsSent.getValue()) return;

  // Format string `attendees` by splitting by '\n', trimming whitespace then flatten array
  if(attendees.toString().length > 1) {
    const splitArray = attendees.split('\n');  // split the string into an array;
    formattedAttendees = splitArray.map(str => str.trim());   // trim whitespace from every string in array
    attendees = formattedAttendees.join(", ");       // combine all array elements into single string
  }
  else attendees = 'none';  // otherwise replace empty string by 'none'
  
  confirmation = (confirmation ? 'Yes' : 'No (explain in comment section)' );
  rangeConfirmation.setValue(confirmation);

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
        <strong>&nbsp;&nbsp;&nbsp;Head Run: </strong>" + headRun + "\
      </p> \
      <p> \
        <strong>&nbsp;&nbsp;&nbsp;Distance: </strong>" + distance + "\
      </p> \
      <p>\
        <strong>&nbsp;&nbsp;&nbsp;Attendees: </strong>" + attendees + "\
      </p> \
      <p> \
        <strong>&nbsp;&nbsp;&nbsp;Submitted by: </strong> " + toEmail + "\
      </p> \
      <p> \
          &nbsp;&nbsp; \
        <strong><em>I declare all attendees have provided their waiver and paid the one-time member fee: </em></strong>" + 
          confirmation + " \
      </p> \
      <p> \
        <strong>&nbsp;&nbsp;&nbsp;Comments: </strong> " + notes + "\
      </p> \
      <p> \
        <br> \
        - McRUN Bot \
      </p> \
    </body> \
  </html>";

  Logger.log(emailBodyHTML);

  
  var message = {
    to: toEmail,
    bcc: emailPresident,
    cc: "mcrunningclub@ssmu.ca" + ", " + emailVPinternal,
    subject: "McRUN Attendance Form (" + formattedDate + ")",
    htmlBody: emailBodyHTML,
    noReply: true,
    name: "McRUN Attendance Bot"
  }

  //MailApp.sendEmail(message);

  rangeCopyIsSent.setValue(true);
  rangeCopyIsSent.insertCheckboxes();
}

/**
 * UPDATE FUNCTION FOR STATIC REFERENCE
 * 
 * @author: Andrey S Gonzalez
 * @date: Oct 30, 2023
 * @update: Oct 30, 2023
 * 
 * Copy newest submission to ledger spreadsheet
 */

function copyToLedger() {
  return // currently out of service (may-25)
  const sourceSheet = ATTENDANCE_SHEET;
  const sheetUrl = "https://docs.google.com/spreadsheets/d/1J-nSg2QLNYkVWc0PplfwQWM8fyujE1Dv_PURL6kBNXI/edit?usp=sharing";

  var destinationSpreadsheet = SpreadsheetApp.openByUrl(sheetUrl);
  var destinationSheet = destinationSpreadsheet.getSheetByName("Head Run Attendance");
  var sourceData = sourceSheet.getRange(sourceSheet.getLastRow(), 1, 1, 5).getValues();

  destinationSheet.appendRow(sourceData[0]);
}


/**
 * @author: Andrey S Gonzalez
 * @date: Oct 15, 2023
 * @update: Oct 10, 2024
 * 
 * Check for missing submission. Triggers set according to following schedule:
 *  Tuesday: 6:00pm
    Wednesday: 6:00pm
    Thursday: 7:30am 
    Saturday: 10:00am
    Sunday: 6:00pm
 *
 * FUNCTIONS TO UPDATE EACH SEMESTER: getHeadRunString(), getHeadRunnerEmail()
 * 
 * TODO: [] change to -> checkMissingAttendance(var headRunAMorPM)
 *       [] modify headrun source to GSheet-> i.e. modify getHeadRunString() and getHeadRunnerEmail()
 */

function checkMissingAttendance() {

  const sheet = ATTENDANCE_SHEET;
  
  // Gets values of all timelogs
  var numRows = sheet.getLastRow() - 1;
  var submissionDates = sheet.getRange(2, TIMESTAMP_COL, numRows).getValues();
  var submissionHeadRuns = sheet.getRange(2, HEADRUN_COL, numRows).getValues();

  // Get date at trigger time and compare with timestamp of existing submissions
  const today = new Date();
  const formattedToday = Utilities.formatDate(today, TIMEZONE, 'yyyy-MM-dd a');

  // Formats trigger datetime to get head runner emails
  const headRunDay = Utilities.formatDate(today, TIMEZONE, 'EEEEa');  // e.g. 'MondayAM'
  const headRunDetail = getHeadRunString(headRunDay);

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
        <strong>Click <a href= https://docs.google.com/forms/d/e/1FAIpQLSf_4zdnyY4I4vSxaatAaxxgsU38hb862arFDU9wTbSpnoXdKA/viewform\> here</a> to access the F24 attendance form. </strong> \
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
 * ChatGPT Example
 */
function copyRowToAnotherSpreadsheet() {
  var sourceSpreadsheet = SpreadsheetApp.openById("SourceSpreadsheetID"); // Replace with the ID of your source spreadsheet
  var destinationSpreadsheet = SpreadsheetApp.openById("DestinationSpreadsheetID"); // Replace with the ID of your destination spreadsheet

  var sourceSheet = sourceSpreadsheet.getSheetByName("SourceSheetName"); // Replace with the name of your source sheet
  var destinationSheet = destinationSpreadsheet.getSheetByName("DestinationSheetName"); // Replace with the name of your destination sheet

  var rowIndexToCopy = 2; // Replace with the row index you want to copy (e.g., row 2)
  var sourceData = sourceSheet.getRange(rowIndexToCopy, 1, 1, sourceSheet.getLastColumn()).getValues();

  destinationSheet.appendRow(sourceData[0]);
}


function deadCode() {
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