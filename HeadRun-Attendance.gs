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
  formatSpecificColumns_();
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
  const startRow = 2;   // Skip header row
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
 * @update  May 11, 2025
 */

function checkMissingAttendance() {
  const scriptProperties = PropertiesService.getScriptProperties();
  const propertyName = SCRIPT_PROPERTY.isCheckingAttendance;  // User defined in `Attendance-Variables.gs`
  const isCheckingAllowed = scriptProperties.getProperty(propertyName).toString();

  if (isCheckingAllowed !== "true") {
    throw new Error("`verifyAttendance()` is not allowed to run. Set script property to true.");
  }

  const today = new Date();  // new Date(new Date().getTime() + 23 * 60 * 60 * 1000);
  const currentWeekday = today.getDay();

  const currentDaySchedule = getScheduleFromStore_(currentWeekday);
  const currentTimeKey = getMatchedTimeKey_(today, currentDaySchedule);

  // Verify if valid timekey
  if (!currentTimeKey) {
    throw new Error(`No timekey found for ${today} with run schedule ${currentDaySchedule}`);
  }

  const weekdayStr = getWeekday_(currentWeekday);
  const headrunTitle = toTitleCase_(weekdayStr) + ' ' + currentTimeKey;    // e.g. 'Tuesday - 9am'

  // Get emails using run schedule for current day
  const runScheduleLevels = currentDaySchedule[currentTimeKey];

  // Headrunner emails separated by levels e.g. {'easy' : [emails], 'advanced' : [emails], ...}
  const emailsByLevel = getHeadrunnerEmailFromStore_(runScheduleLevels);
  const emailObj = { 'emailsByLevel' : emailsByLevel, 'headrunTitle' : headrunTitle };

  // Save result of attendance verification, and get title for email
  const { 'timeKey' : matchedTimeKey, 'submission' : submission } = verifyAttendance_(currentWeekday);

  // Send copy of submission if true. Otherwise send an email reminder to headrunners
  (currentTimeKey === matchedTimeKey) ? sendSubmissionCopy_(emailObj, submission) : sendEmailReminder_(emailObj);
  Logger.log(`Executed 'checkMissingAttendance' with\n`, emailObj);


  /** Helper function */
  function verifyAttendance_(currentWeekday) {
    const sheet = ATTENDANCE_SHEET;

    // Gets values of all timelogs
    const numRows = sheet.getLastRow() - 1;
    const numCols = COMMENTS_COL;
    const submissionArr = sheet.getSheetValues(2, TIMESTAMP_COL, numRows, numCols);

    // Get date at trigger time and compare with timestamp of existing submissions
    return findMatchingTimeKey();

    /** Helper function */
    function findMatchingTimeKey() {
      const ret = { timeKey : null, submission : null }

      // Start checking from end of head run attendance submissions
      // Exit loop when submission found or until list exhausted
      for (let i = numRows - 1; i >= 0 && !ret.timeKey; i--) {
        
        const submissionDate = submissionArr[i][0];   // Date index = 0
        const thisWeekday = submissionDate.getDay();

        if (thisWeekday === currentWeekday) {
          const runSchedule = getScheduleFromStore_(thisWeekday);
          ret.timeKey = getMatchedTimeKey_(submissionDate, runSchedule);
          ret.submission = submissionArr[i];    // Need values for confirmation email
        }
      }
      
      return ret;
    }
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



function copyTest() {
  const weekdayStr = 'tuesday';
  const currentTimeKey = '6pm';

  const headrunTitle = toTitleCase_(weekdayStr) + ' ' + currentTimeKey;    // e.g. 'Tuesday - 9am'
  const runSchedule = getScheduleFromStore_(weekdayStr);
  const emailsByLevel = getHeadrunnerEmailFromStore_(runSchedule[currentTimeKey]);

  const emailObj = { 'emailsByLevel' : emailsByLevel, 'headrunTitle' : headrunTitle };
  const numCols = COMMENTS_COL;
  const submission = GET_ATTENDANCE_SHEET_().getSheetValues(2, TIMESTAMP_COL, 1, numCols)[0];

  sendSubmissionCopy_(emailObj, submission);
}


function sendBotEmail(subject, recipient, htmlBody) {
  const reminderEmail = {
    to: recipient,
    bcc: PRESIDENT_EMAIL,
    cc: CLUB_EMAIL + "," + VP_INTERNAL_EMAIL,
    subject: subject,
    htmlBody: htmlBody,
    noReply: true,
    name: "McRUN Attendance Bot"
  }

  MailApp.sendEmail(reminderEmail);
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
 * @update  May 14, 2025
 */

function sendSubmissionCopy_({ emailsByLevel, headrunTitle}, submission) {
  // Error Management: prevent wrong user sending email
  if (getCurrentUserEmail_() != CLUB_EMAIL) return;

  // Make submission 1-indexed
  submission.unshift('');

  // Create regex to extract name from name-email pairings, e.g. `Bob Burger:bob@mail.com` -> `Bob Burger`
  const nameRegex = /^(.*?):/gm;
  const extractNames = (nameEmail) => [...nameEmail.matchAll(nameRegex)].map(m => m[1]).join(', ');

  // Format attendees as `['Easy: Bob Burger, Cat Fox', 'Intermediate: None', Advanced:' Catherine Rex']`
  const allAttendees = Object.entries(ATTENDEE_MAP).map(([level, index]) => {
    const label = toTitleCase_(level);
    const levelAttendee = extractNames(submission[index]) || EMPTY_ATTENDEE_FLAG;
    return `- ${label}: ${levelAttendee}`;
  });

  // Prepare values to populate copy email template
  const headrun = {
    title : headrunTitle,
    distance : submission[DISTANCE_COL],
    attendees : allAttendees,
    toEmail : Object.values(emailsByLevel).join(','),
    confirmation : submission[CONFIRMATION_COL],
    comments : submission[COMMENTS_COL] || 'None'
  };

  // Create html code and populate placeholders
  const copyEmailHTML = createEmailCopy_(headrun);

  // Send email with following arguments
  const subject = "McRUN Attendance Form (" + headrunTitle + ")"
  sendBotEmail(subject, headrun.toEmail, copyEmailHTML);
  Logger.log(`Successfully sent copy of attendance submission for (${headrunTitle})`);

  // const message = {
  //   to: headrun.toEmail,
  //   bcc: PRESIDENT_EMAIL,
  //   cc: CLUB_EMAIL + "," + VP_INTERNAL_EMAIL,
  //   subject: "McRUN Attendance Form (" + headrunTitle + ")",
  //   htmlBody: emailBodyHTML,
  //   noReply: true,
  //   name: "McRUN Attendance Bot"
  // }

  // MailApp.sendEmail(message);
  // console.log(`Successfully sent copy of attendance submission for (${headrunTitle})`);
}


function sendEmailReminder_({ emailsByLevel, headrunTitle }) {
  // Error Management: prevent wrong user sending email
  if (getCurrentUserEmail_() != CLUB_EMAIL) throw Error('Please change to McRUN account');

  // Load HTML template and replace placeholders
  const templateName = REMINDER_EMAIL_HTML_FILE;
  const template = HtmlService.createTemplateFromFile(templateName);

  template.LINK = GET_ATTENDANCE_GFORM_LINK_();
  template.SEMESTER = SEMESTER_NAME;

  // Returns string content from populated html template
  const reminderEmailHTML = template.evaluate().getContent();
  const subject = "McRUN Missing Attendance Form - " + headrunTitle;
  const recipients = Object.values(emailsByLevel).join(",");

  // Send reminder email with following paramaters
  sendBotEmail(subject, recipients, reminderEmailHTML);

  // const reminderEmail = {
  //   to: Object.values(emailsByLevel).join(","),
  //   bcc: PRESIDENT_EMAIL,
  //   cc: CLUB_EMAIL + "," + VP_INTERNAL_EMAIL,
  //   subject: "McRUN Missing Attendance Form - " + headrunTitle,
  //   htmlBody: reminderEmailHTML,
  //   noReply: true,
  //   name: "McRUN Attendance Bot"
  // }

  // MailApp.sendEmail(reminderEmail);
  console.log(`Reminder sent successfully for missing attendance submission (${headrunTitle})`);
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
    .replace(/\b\w/g, l => l.toUpperCase());  // Capitalize each name
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

