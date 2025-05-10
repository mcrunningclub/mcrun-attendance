// PREVIOUS FUNCTIONS: onEditCheck(), consolidateSubmissions()

/**
 * Functions to execute after form submission.
 * 
 * To use as a trigger function, it cannot have parameters.
 * Otherwise, a runtime exception is raised for new form submission.
 *
 * @trigger Attendance Form or McRUN App Submission.
 */

function onFormSubmission() {
  // Since GForm might not add submission to bottom, sort beforehand
  sortAttendanceForm();
  SpreadsheetApp.flush();

  const row = getLastSubmission_();  // Get submission row index
  console.log(`Latest row number: ${row}`);

  onFormSubmissionInRow_(row);
  transferAndFormat_(row);
}


function onFormSubmissionInRow_(row) {
  addMissingPlatform_(row);    // Sets platform to 'Google Form'
  bulkFormatting_(row);
  getUnregisteredMembersInRow_(row);    // Find any unregistered members
}


/**
 * Functions to execute after McRUN app submission.
 *
 * @trigger McRUN App Attendance Submission.
 */

function onAppSubmission(row = ATTENDANCE_SHEET.getLastRow()) {
  console.log(`[AC] Starting 'onAppSubmission' for row ${row}`);
  bulkFormatting_(row);
  transferAndFormat_(row);

  sortAttendanceForm();
  console.log(`[AC] Completed 'onAppSubmission' successfully!`);
}


function bulkFormatting_(row) {
  formatConfirmationInRow_(row);  // Transforms bool to user-friendly message
  formatNamesInRow_(row);     // Formats names in last row
}

function transferAndFormat_(row) {
  const logRow = transferSubmissionToLedger(row);
  triggerEmailInLedger_(logRow)
  formatSpecificColumns();
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

function getLastSubmission_(sheet = GET_ATTENDANCE_SHEET_()) {
  const startRow = 1;
  const numRow = sheet.getLastRow();

  // Fetch all values in the TIMESTAMP_COL
  const values = sheet.getSheetValues(startRow, TIMESTAMP_COL, numRow, 1);
  let lastRow = values.length;

  // Loop through the values in reverse order
  while (values[lastRow - 1][0] === "") {
    lastRow--;
  }

  return lastRow;
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
 * @update  Apr 7, 2025
 */

function emailSubmission() {
  // Error Management: prevent wrong user sending email
  if (getCurrentUserEmail() != 'mcrunningclub@ssmu.ca') return;

  const sheet = GET_ATTENDANCE_SHEET_();
  const lastRow = getLastSubmission_();

  // Save values in 0-indexed array, then transform into 1-indexed by appending empty
  // string to the front. Now, access is easier e.g [EMAIL_COL] vs [EMAIL_COL-1]
  const submission = sheet.getSheetValues(lastRow, 1, 1, sheet.getLastColumn())[0];
  submission.unshift('');   // append '' to front for 1-indexed access

  var timestamp = new Date(submission[TIMESTAMP_COL]);
  const formattedDate = Utilities.formatDate(timestamp, TIMEZONE, 'MM/dd/yyyy');

  // Replace newline with comma-space
  const allAttendees = Object.values(ATTENDEE_MAP).map(
    level => submission[level].replaceAll('\n', ', ')
  );

  // Read only (cannot edit values in sheet)
  const headrun = {
    name: submission[HEADRUN_COL],
    distance: submission[DISTANCE_COL],
    attendees: allAttendees,
    toEmail: submission[EMAIL_COL],
    confirmation: submission[CONFIRMATION_COL],
    notes: submission[COMMENTS_COL]
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
    bcc: PRESIDENT_EMAIL,
    cc: "mcrunningclub@ssmu.ca" + ", " + VP_INTERNAL_EMAIL,
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
 * @author [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>)
 * @date  Oct 15, 2023
 * @update  May 4, 2025
 */

function checkMissingAttendance() {
  // const scriptProperties = PropertiesService.getScriptProperties();
  // const propertyName = SCRIPT_PROPERTY.isCheckingAttendance;  // User defined in `Attendance-Variables.gs`
  // const isCheckingAllowed = scriptProperties.getProperty(propertyName).toString();

  // if (isCheckingAllowed !== "true") {
  //   throw new Error("`verifyAttendance()` is not allowed to run. Set script property to true.");
  // }

  const today = new Date(new Date().getTime() + 23 * 60 * 60 * 1000);   // new Date();
  const currentWeekday = today.getDay();

  const currentDaySchedule = getScheduleFromStore_(currentWeekday);
  const currentTimeKey = getMatchedTimeKey(today, currentDaySchedule);

  const weekdayStr = getWeekday_(currentWeekday);
  const headrunTitle = toTitleCase(weekdayStr) + ' ' + currentTimeKey;

  // Get emails using runSchedule
  const runScheduleLevels = currentDaySchedule[currentTimeKey];
  const emailsByLevel = getHeadrunnerEmailFromStore_(runScheduleLevels);
  const emailObj = { 'emails' : emailsByLevel, 'headrunTitle' : headrunTitle };

  // Save result of attendance verification, and get title for email
  const matchedTimeKey = verifyAttendance_(currentWeekday);

  // Send copy of submission if true. Otherwise send an email reminder
  (currentTimeKey === matchedTimeKey) ? sendSubmissionCopy_(emailObj) : sendEmailReminder_(emailObj);
  console.log(`Executed 'checkMissingAttendance' with\n`, emailObj);
}


function checkForNewImport_() {
  const importSheet = IMPORT_SHEET;
  const numRow = importSheet.getLastRow();
  const numCol = importSheet.getLastColumn();

  // Check the last 5 rows
  const numRowToCheck = 5;
  const startRow = numRow - numRowToCheck;

  // Get range but do not sort sheet. Non-imported submissions most likely at bottom.
  const rangeToCheck = importSheet.getRange(startRow, 1, numRowToCheck, numCol);

  throw new Error('Function is incomplete. Please review.');
}


function verifyAttendance_(currentWeekday) {
  const sheet = ATTENDANCE_SHEET;

  // Gets values of all timelogs
  const numRows = sheet.getLastRow() - 1;
  const submissionDates = sheet.getSheetValues(2, TIMESTAMP_COL, numRows, 1);

  // Get date at trigger time and compare with timestamp of existing submissions
  return findMatchingTimeKey();

  /** Helper function */
  function findMatchingTimeKey() {
    let timeKey = null;

    // Start checking from end of head run attendance submissions
    // Exit loop when submission found or until list exhausted
    for (let i = numRows - 1; i >= 0 && !timeKey; i--) {
      
      const submissionDate = submissionDates[i][0];
      const thisWeekday = submissionDate.getDay();

      if (thisWeekday === currentWeekday) {
        const runSchedule = getScheduleFromStore_(thisWeekday);
        timeKey = getMatchedTimeKey(submissionDate, runSchedule);
      }
    }
    
    return timeKey;
  }

  function normalizeHeadrun(headrun) {
    [dayOfWeek, time,] = headrun.toLowerCase().split(/\s*-\s*|\s+/);

    // Add missing ':00' if needed
    if (/^\d{1,2}(am|pm)$/i.test(time)) {
      time = time.replace(/(am|pm)$/i, ':00$1');
    }
    return `${dayOfWeek} - ${time}`;
  }
}


function sendEmailReminder_(headrunTitle) {
  [dayOfWeek, time,] = headrunTitle.split(/\s*-\s*|\s+/);
  const amPmOfDay = time.match(/(am|pm)/i);
  const headrunDay = (dayOfWeek + amPmOfDay[0]);    // e.g. 'MondayAM'

  // Get head runners email using input headrun
  const headRunnerEmail = getHeadRunnerEmail_(headrunDay).join();

  // Load HTML template and replace placeholders
  const templateName = REMINDER_EMAIL_HTML_FILE;
  const template = HtmlService.createTemplateFromFile(templateName);

  template.LINK = GET_ATTENDANCE_GFORM_LINK_();
  template.SEMESTER = SEMESTER_NAME;

  // Returns string content from populated html template
  const reminderEmailBodyHTML = template.evaluate().getContent();

  var reminderEmail = {
    to: headRunnerEmail,
    bcc: PRESIDENT_EMAIL,
    cc: "mcrunningclub@ssmu.ca" + ", " + VP_INTERNAL_EMAIL,
    subject: "McRUN Missing Attendance Form - " + headrunTitle,
    htmlBody: reminderEmailBodyHTML,
    noReply: true,
    name: "McRUN Attendance Bot"
  }

  MailApp.sendEmail(reminderEmail);
  console.log(`Reminder sent successfully for missing attendance submission (${headrunTitle})`);
}


function sendSubmissionCopy_() {
  console.log("Please complete 'sendSubmissionCopy'");
  return;
}


/**
 * Wrapper function for `getUnregisteredMembers` for *ALL* rows.
 * 
 * Row number is 1-indexed in GSheet. Executes top to bottom. Header row skipped.
 * 
 */

function getAllUnregisteredMembers_() {
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
 * @update  Apr 1, 2025
 */

function getUnregisteredMembersInRow_(row = ATTENDANCE_SHEET.getLastRow()) {
  const sheet = ATTENDANCE_SHEET;
  const numColToGet = LEVEL_COUNT;

  // Get attendee names starting from beginner col to advanced col
  const allAttendees = sheet.getSheetValues(row, ATTENDEES_BEGINNER_COL, 1, numColToGet)[0];
  const registered = [];
  const unregistered = [];

  const memberMap = getMemberMap_();
  const cleanMemberMap = formatAndSortMemberMap_(memberMap, 0, 1);    // searchKeyIndex: 0, emailIndex: 1

  // Function to prepare a name: remove whitespace, strip accents, capitalize, and swap names
  const prepareThisName = compose_(formatThisName_, reverseThisName_);

  allAttendees.forEach(level => {
    const namesWithEmail = [];
    const namesWithoutEmail = [];

    // Skip if level contains the EMPTY_ATTENDEE_FLAG
    if (level.includes(EMPTY_ATTENDEE_FLAG)) {
      registered.push('');
      unregistered.push('');
      return;
    };

    // Process and separate names in the level
    level.split('\n').forEach(name => {
      if (name.includes(':')) {
        namesWithEmail.push(name); // Name with email
      } else {
        const nameToFind = prepareThisName(name);
        namesWithoutEmail.push(nameToFind); // Name without email
      }
    });

    // Try to find names without email in registrations (memberMap)
    // And append their respective email as in `namesWithEmail`
    const {
      'unregistered': foundUnregistered,
      'registered': foundRegistered
    } = findUnregistered_(namesWithoutEmail.sort(), cleanMemberMap);

    // Combine and sort both arrays
    const mergedRegistered = [...namesWithEmail, ...foundRegistered];

    registered.push(mergedRegistered.sort());
    unregistered.push(foundUnregistered);
  });

  // Log unfound names
  setNamesNotFound_(row, unregistered.join("\n"));

  // Append email to registered attendees
  appendMemberEmail_(row, registered, unregistered);
}


function setNamesNotFound_(row, notFoundArr) {
  const sheet = ATTENDANCE_SHEET;
  const unfoundNameRange = sheet.getRange(row, NAMES_NOT_FOUND_COL);
  unfoundNameRange.setValue(notFoundArr);
}


// Group functions to apply on `input`
function compose_(...fns) {
  return (input) => fns.reduce((v, f) => f(v), input);
}


// Remove whitespace, strip accents and capitalize names
// Swap order of attendee names to `lastName, firstName`
function formatThisName_(name) {
  return name
    .trim()
    .normalize("NFD").replace(/[\u0300-\u036f]/g, "")  // Strip accents
    .replace(/[\u2018\u2019']/g, "") // Remove apostrophes (`, ', ’)
    .toLowerCase()
    .replace(/\b\w/g, l => l.toUpperCase()) // Capitalize each name
    ;
}


// Format as `LastName, FirstName`
function reverseThisName_(name) {
  let nameParts = name.split(/\s+/)   // Split by spaces;

  if (nameParts.length === 1) {
    return name;
  }

  // Replace hyphens with spaces. Can only perform after splitting first and last name.
  nameParts = nameParts.map(p => p.replace(/-/g, ' '));

  // If first name is not hyphenated, only left-most substring stored in first name
  const firstName = nameParts[0];
  const lastName = nameParts[nameParts.length - 1];
  return `${lastName}, ${firstName}`;
}


function getMemberMap_() {
  // Get existing member registry in `Members` sheet
  const memberSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Members");

  const memberEmailCol = MEMBER_EMAIL_COL - 1;    // GSheet to array (1-index to 0-index)
  const memberKeyIndex = MEMBER_SEARCH_KEY_COL - 1;

  const startCol = 1;
  const startRow = 2;   // Skip header row
  const numCols = MEMBER_SEARCH_KEY_COL;
  const memberCount = memberSheet.getLastRow() - 1;   // Do not count header row

  const memberSheetValues = memberSheet.getSheetValues(startRow, startCol, memberCount, numCols);

  // Get array of member names to use as search key, combined with email
  // Step 1. Combine memberKey and email
  // Step 2. Filter rows with empty names or emails
  const memberMap = memberSheetValues
    .map(row => [row[memberKeyIndex], row[memberEmailCol]])
    .filter(row => row[0] && row[1])
    ;

  return memberMap;
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
 * @update  Feb 9, 2025
 */

function findUnregistered_(attendees, memberMap) {
  const unregistered = [];
  const registeredMap = {}    // Saves member name-email pair

  if (attendees.length < 1) {
    return { 'registered': [], 'unregistered': [] };
  }

  const SEARCH_KEY_INDEX = 0;
  const EMAIL_INDEX = 1;

  let index = 0;

  for (const attendee of attendees) {
    // Split attendee name into last and first name
    const [attendeeLastName, attendeeFirstName = ""] = attendee.split(",").map(s => s.trim());

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

  const registered = Object.keys(registeredMap).map(name => {
    const [lastName, firstName] = name.split(', ');
    const email = registeredMap[name];
    return `${firstName} ${lastName}:${email}`;
  });


  const returnObject = {
    'registered': registered || [],
    'unregistered': unregistered.sort() || []  // sorted list of unregistered attendees
  };

  return returnObject;
}

