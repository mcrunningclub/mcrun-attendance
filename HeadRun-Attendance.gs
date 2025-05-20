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
 * @warning Service property `IS_CHECKING_ATTENDANCE` must be set to `true`.
 *
 * @trigger 30-60 mins after headrun schedule.
 *
 * @author [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>)
 * @date  Oct 15, 2023
 * @update  May 15, 2025
 */

function checkMissingAttendance() {
  // First, check if attendance can be verified
  checkIfLegal();

  const today = new Date();  // new Date(new Date().getTime() + 23 * 60 * 60 * 1000);
  const currentWeekday = today.getDay();

  const currentDaySchedule = getScheduleFromStore_(currentWeekday);
  const currentTimeKey = getMatchedTimeKey_(today, currentDaySchedule);

  // Verify if valid timekey
  if (!currentTimeKey) {
    throw new Error(`No timekey found for ${today} with run schedule\n\n${JSON.stringify(currentDaySchedule)}\n\n`);
  }

  const weekdayStr = getWeekday_(currentWeekday);
  const headrunTitle = toTitleCase_(weekdayStr) + ' ' + currentTimeKey;    // e.g. 'Tuesday 9am'

  // Get emails using run schedule for current day, then proceed to actual verification
  const runScheduleLevels = currentDaySchedule[currentTimeKey];
  const { 'timeKey' : matchedTimeKey, 'submission' : submission } = verifyAttendance_(currentWeekday);

  // Headrunner emails separated by levels e.g. {'easy' : [emails], 'advanced' : [emails], ...}
  const emailsByLevel = getHeadrunnerEmailFromStore_(runScheduleLevels);
  const emailObj = { 'emailsByLevel' : emailsByLevel, 'headrunTitle' : headrunTitle };

  // Send copy of submission if true. Otherwise send an email reminder to headrunners
  (currentTimeKey === matchedTimeKey) ? sendSubmissionCopy_(emailObj, submission) : sendEmailReminder_(emailObj);
  Logger.log(`Executed 'checkMissingAttendance' with ${JSON.stringify(emailObj)}`);


  /** Helper functions */
  function checkIfLegal() {
    const scriptProperties = GET_PROP_STORE_();
    const propertyName = SCRIPT_PROPERTY.isCheckingAttendance;  // User defined in `Attendance-Variables.gs`
    const isCheckingAllowed = scriptProperties.getProperty(propertyName).toString();

    if (isCheckingAllowed !== "true") {
      throw new Error("'verifyAttendance()' is not allowed to run. Set script property to true.");
    }
  }

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
}

function sendBotEmail_(subject, recipient, htmlBody) {
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
 * @update  May 15, 2025
 */

function sendSubmissionCopy_({ emailsByLevel, headrunTitle}, submission) {
  // Error Management: prevent wrong user sending email
  if (getCurrentUserEmail_() != CLUB_EMAIL) throw Error('Please change to McRUN account');

  submission.unshift('');   // Make submission 1-indexed

  // Prepare values to populate copy email template
  const headrun = {
    title : headrunTitle,
    distance : submission[DISTANCE_COL],
    attendees : prepareAttendees(),
    confirmation : submission[CONFIRMATION_COL],
    comments : submission[COMMENTS_COL] || 'None'
  };

  // Create html code by populating with `headrun` values
  const copyEmailHTML = createEmailCopy_(headrun);

  // Send email using email bot helper function
  sendBotEmail_(
    "McRUN Attendance Copy (" + headrunTitle + ")",  // Subject
    Object.values(emailsByLevel).join(','),   // Headrunner recipients
    copyEmailHTML   // HTML body
  ); 

  Logger.log(`Successfully sent copy of attendance submission for '${headrunTitle}'`);

  /** Helper Function */
  function prepareAttendees() {
    // Create regex to extract name from name-email pairings, e.g. `Bob Burger:bob@mail.com` -> `Bob Burger`
    const nameRegex = /^(.*?):/gm;
    const extractNames = (nameEmail) => [...nameEmail.matchAll(nameRegex)].map(m => m[1]).join(', ');

    // Format attendees as `['- Easy: Bob Burger, Cat Fox', '- Intermediate: None', '- Advanced: Catherine Brown']`
    return Object.entries(ATTENDEE_MAP).map(([level, sheetIndex]) => {
      const label = toTitleCase_(level);
      const levelAttendee = extractNames(submission[sheetIndex]) || EMPTY_ATTENDEE_FLAG;
      return `- ${label}: ${levelAttendee}`;
    });
  }
}


/**
 * Send a reminder email to headrunners when attendance for respective headrun not found.
 *
 * @author [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>)
 * @date  May 2, 2025
 * @update  May 15, 2025
 */

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
  sendBotEmail_(subject, recipients, reminderEmailHTML);
  Logger.log(`Reminder sent successfully for missing attendance submission (${headrunTitle})`);
}

