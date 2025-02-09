/**
 * Functions to execute after form submission.
 *
 * @trigger Attendance form Submission.
 */

function onFormSubmission() {
  addMissingFormInfo();
  formatNamesInRow_();     // formats names in last row
  getUnregisteredMembersInRow_();

  //emailSubmission();    // IN-REVIEW
  formatSpecificColumns();
}


/**
 * Functions to execute after McRUN app submission.
 *
 * @trigger McRUN App Attendance Submission.
 */
function onAppSubmission(row=ATTENDANCE_SHEET.getLastRow()) {
  //removePresenceChecks();
  formatConfirmationInRow_(row);

  //emailSubmission();    // IN-REVIEW
  hideAttendeeEmailInRow_(row);
  formatSpecificColumns();
  hideAttendeeEmailInRow_(row-1); // Rehide emails because second-last col is not footer anymore
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
 * @update  Feb 9, 2025
 */

function onEditCheck() {
  Utilities.sleep(5 * 1000);  // Let latest submission sync for 5 seconds

  const sheet = ATTENDANCE_SHEET;
  const timestamp = new Date(sheet.getRange(sheet.getLastRow(), 1).getValue());
  const timeLimit = 45 * 1000   // 45 sec = 45 * 1000 millisec

  // Exit if over time limit
  if (Date.now() - timestamp > timeLimit) {
    Logger.log(`Timestamp: ${timestamp} is older than ${timeLimit}.
    onEditCheck() exited.`)
    return;
  }

  const latestPlatform = sheet.getRange(sheet.getLastRow(), PLATFORM_COL).getValue();

  if (latestPlatform.toLowerCase() === "mcrun app") {
    Logger.log(`Latest Platform: ${latestPlatform}`);
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
 * Attendees are separated by level.
 *
 * @trigger Attendance submissions.
 *
 * @author [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>)
 * @date  Oct 9, 2023
 * @update  Oct 30, 2024
 */

function emailSubmission() {
  // Error Management: prevent wrong user sending email
  if (getCurrentUserEmail() != 'mcrunningclub@ssmu.ca') return;

  const sheet = ATTENDANCE_SHEET;
  const lastRow = sheet.getLastRow();
  const latestSubmission = sheet.getRange(lastRow, 1, 1, sheet.getLastColumn());

  // Save values in 0-indexed array, then transform into 1-indexed by appending empty
  // string to the front. Now, access is easier e.g [EMAIL_COL] vs [EMAIL_COL-1]
  const values = latestSubmission.getValues()[0];
  values.unshift('');   // append '' to front for 1-indexed access

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
    name: values[HEADRUN_COL],
    distance: values[DISTANCE_COL],
    attendees: allAttendees,
    toEmail: values[EMAIL_COL],
    confirmation: values[CONFIRMATION_COL],
    notes: values[COMMENTS_COL]
  };

  // Read and edit sheet values
  const rangeConfirmation = sheet.getRange(lastRow, CONFIRMATION_COL);
  const rangeIsCopySent = sheet.getRange(lastRow, IS_COPY_SENT_COL);

  // Only send if submitter wants copy && email has not been sent yet
  if (rangeIsCopySent.getValue()) return;

  // Replace newline delimiter with comma-space if non-empty or matches "None"
  const attendeesStr = headrun.attendees.toString();

  if (attendeesStr.length > 1) {
    headrun.attendees = attendeesStr.replaceAll('\n', ', ');
  }
  else headrun.attendees = EMPTY_ATTENDEE_FLAG;  // otherwise replace empty string by 'none'

  headrun.confirmation = (headrun.confirmation ? 'Yes' : 'No (explain in comment section)');
  rangeConfirmation.setValue(headrun.confirmation);

  const emailBodyHTML = createEmailCopy_(headrun);

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
 * Toggles flag to run `checkAttendance()` by updating value in `ScriptProperties` bank.
 *
 * @trigger User choice in custom menu.
 *
 * @author [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>)
 * @date  Dec 5, 2024
 * @update  Dec 6, 2024
 */

function toggleAttendanceCheck_() {
  const scriptProperties = PropertiesService.getScriptProperties();
  const propertyName = SCRIPT_PROPERTY.isCheckingAttendance;  // User defined in `Attendance-Variables.gs`

  const isChecking = scriptProperties.getProperty(propertyName);  // !! converts str to bool
  const toggledState = (isChecking == "true") ? "false" : "true";   // toggle bool, but save as str
  scriptProperties.setProperty(propertyName, toggledState);    // function requires property as str

  return toggledState;
}

/**
 * Check for missing submission after scheduled headrun.
 *
 * Service property `IS_CHECKING_ATTENDANCE` must be set to `true`.
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
 * @update  Jan 24, 2025
 */

function checkMissingAttendance() {
  const scriptProperties = PropertiesService.getScriptProperties();
  const propertyName = SCRIPT_PROPERTY.isCheckingAttendance;  // User defined in `Attendance-Variables.gs`
  const isCheckingAllowed = scriptProperties.getProperty(propertyName).toString();

  if (isCheckingAllowed !== "true") {
    Logger.log("Error: `verifyAttendance()` is not allowed to run. Set script property to true.");
    return;
  }

  verifyAttendance_();
}

function verifyAttendance_() {
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
    // Create an instance of Error with a custom message
    var errorMessage = "No headrunner has been found for " + headRunDay;
    throw new Error(errorMessage); // Throw the ExecutionError
  }

  // Start checking from end of head run attendance submissions
  // Exit loop when submission found or until list exhausted
  var isSubmitted = false;

  for (var i = numRows - 1; i >= 0 && !isSubmitted; i--) {
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

  // Load HTML template and replace placeholders
  const templateName = REMINDER_EMAIL_HTML_FILE;
  const template = HtmlService.createTemplateFromFile(templateName);

  template.LINK = ATTENDANCE_GFORM_LINK;
  template.SEMESTER = SEMESTER_NAME;

  // Returns string content from populated html template
  const reminderEmailBodyHTML = template.evaluate().getContent();

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
}


/**
 * Wrapper function for `getUnregisteredMembers` for *ALL* rows.
 *
 * Row number is 1-indexed in GSheet. Executes top to bottom. Header row skipped.
 *
 * @author [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>)
 * @date  Dec 6, 2024
 * @update  Dec 14, 2024
 */

function getAllUnregisteredMembers() {
  runOnSheet_(getUnregisteredMembersInRow_.name);
}

/**
 * Find attendees in `row` of `ATTENDANCE_SHEET `that are unregistered members.
 *
 * Sets unregistered members in `NOT_FOUND_COL`.
 *
 * List of members found in `Members` sheet.
 *
 * @param {number} [row=ATTENDANCE_SHEET.getLastRow()]  The row number in `ATTENDANCE_SHEET` 1-indexed.
 *                                                      Defaults to the last row in the sheet.
 *
 * @author [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>) & ChatGPT
 * @date  Oct 30, 2024
 * @update  Dec 14, 2024
 */

function getUnregisteredMembersInRow_(row = ATTENDANCE_SHEET.getLastRow()) {
  const sheet = ATTENDANCE_SHEET;
  const unfoundNameRange = sheet.getRange(row, NAMES_NOT_FOUND_COL);

  // Get attendee names starting from beginner col to advanced col
  const numColToGet = LEVEL_COUNT;
  const attendeeRange = sheet.getRange(row, ATTENDEES_BEGINNER_COL, 1, numColToGet);  // Attendees columns

  // 1D Array of size `LEVEL_COUNT` (Beginner, Intermediate, Advanced -> 3)
  // If email appended to name, remove before creating list of all attendees
  const allNames = attendeeRange.getValues()[0]
    .filter(level => !level.includes(EMPTY_ATTENDEE_FLAG)) // Skip levels with "None"
    .flatMap(level =>
      level.split('\n')                       // Split names in each level into separate entry in array
        .map(entry => entry.includes(':')
          ? entry.split(':')[0].trim()        // Remove email from entry if applicable
          : entry.trim()                      // Otherwise, trim name only
        )
    );

  // Remove whitespace, strip accents and capitalize names
  // Swap order of attendee names to `lastName, firstName`
  const sortedNames = swapAndFormatName_(allNames);

  // Get existing member registry in `Members` sheet
  const memberSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Members");

  const memberEmailCol = MEMBER_EMAIL_COL - 1;    // 1-index to 0-index
  const memberKeyIndex = MEMBER_SEARCH_KEY_COL - 1;   // 1-index to 0-index

  const startCol = 1;
  const startRow = 2;   // Skip header row
  const numCols = MEMBER_SEARCH_KEY_COL;
  const memberCount = memberSheet.getLastRow() - 1;   // Do not count header row

  var memberSheetRange = memberSheet.getRange(startRow, startCol, memberCount, numCols)

  // Get array of member names to use as search key, combined with email for later use
  const memberMap = memberSheetRange.getValues()
    .map(row => [row[memberKeyIndex], row[memberEmailCol]])  // Combine memberKey and email
    .filter(row => row[0] && row[1])  // Filter rows with empty names or emails
    ;

  const cleanMemberMap = formatAndSortMemberMap_(memberMap, 0, 1);    // searchKeyIndex: 0, emailIndex: 1

  // Use the helper function on sorted items
  const registeredAndUnregistered = findUnregistered_(sortedNames, cleanMemberMap);

  // Separate data for registered and unregistered
  const unregistered = registeredAndUnregistered['unregistered'];
  const registered = registeredAndUnregistered['registered'];

  // Log unfound names
  unfoundNameRange.setValue(unregistered.join("\n"));

  // Append email to registered attendees
  appendMemberEmail(row, registered);

  // Hide emails in row, and highlight unregistered attendee
  hideAttendeeEmailInRow_(row);
}


/**
 * Helper function to find unregistered attendees.
 *
 * @param {string[]} attendees  All attendees of the head run (sorted).
 * @param {string[][]} memberMap  All search keys of registered members (sorted) and emails.
 * @return {Map<Map<String,String>, List>}  Returns attendees not found in `members`.
 *
 * @author [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>) & ChatGPT
 * @date  Oct 30, 2024
 * @update  Dec 14, 2024
 */

function findUnregistered_(attendees, memberMap) {
  const unregistered = [];
  const registeredMap = {}    // Saves member name-email pair

  const SEARCH_KEY_INDEX = 0;
  const EMAIL_INDEX = 1;

  let index = 0;

  for (const attendee of attendees) {
    // Split attendee name into last and first name
    const [attendeeLastName, attendeeFirstName] = attendee.split(",").map(s => s.trim());

    let isFound = false;

    // Check members starting from the current index
    while (index < memberMap.length) {
      const memberSearchKey = memberMap[index][SEARCH_KEY_INDEX];
      const [memberLastName, memberFirstNames] = memberSearchKey.split(",").map(s => s.trim());
      const searchFirstNameList = memberFirstNames.split("|").map(s => s.trim());   // only if preferredName exists

      // Compare last names and check if first name matches any in the list
      if (attendeeLastName === memberLastName && searchFirstNameList.includes(attendeeFirstName)) {
        isFound = true;

        // Create entry using existing memberSearchKey
        const memberEmail = memberMap[index][EMAIL_INDEX];    // Get member email
        registeredMap[memberSearchKey] = memberEmail;    // Push name-email pair to object

        index++; // Move to the next member
        break;
      }

      // If attendee's last name is less than the current member's last name
      if (attendeeLastName < memberLastName) {
        break; // Stop searching as attendees are sorted alphabetically
      }

      index++;
    }

    // If attendee not found, add to unregistered array, and put back hyphen in names
    if (!isFound) {
      unregistered.push(`${attendeeFirstName.replace(' ', '-')} ${attendeeLastName.replace(' ', '-')}`);
    }
  }


  const returnObject = {
    'registered': registeredMap,
    'unregistered': unregistered.sort()  // sorted list of unregistered attendees
  };

  return returnObject;
}



/**
 * 
 */
