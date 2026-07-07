/**
 * Registers column positions from `ATTENDANCE_SHEET`.
 *
 * Prevents user from updating column variables manually.
 *
 * CURRENTLY IN REVIEW!
 *
 * @author [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>)
 * @date  Oct 23, 2024
 * @update  Oct 23, 2024
 */

function getColumnPosition() {
  var rangeList = ATTENDANCE_SHEET.getNamedRanges();
  var dRange = ATTENDANCE_SHEET.getNamedRanges()[0].getRange();

  for (var i = 0; i < rangeList.length; i++) {
    Logger.log(rangeList[i].getName());
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


/**
 * MyNewType definition
 * @typedef {Object} MyNewType
 * @property {function} logFirst
 * @property {function} logSecond
 */

/**
 * @param {number} first
 * @param {number} second
 * @returns MyNewType
 */
function gen_(first, second) {
  /**
   * logs first argument
   * @param {number} times
   */
  function logFirst(times) {
    for (let i = 0; i < times; i++) {
      console.log(first);
    }
  }

  /**
   * logs second argument
   * @param {number} times
   */
  function logSecond(times) {
    for (let i = 0; i < times; i++) {
      console.log(second);
    }
  }

  return {
    logFirst,
    logSecond
  };
};



/**
 * @typedef {Object} Result2
 * @property {number} row - The parsed integer value of `userResponse`.
 * @property {string} msg - A custom message to display to the user.
 */


// TO COMPLETE!!
function sendStatsEmailFromExternal(logSheet, row, activity) {
  // Prevent email sent by wrong user
  if (getCurrentUserEmail_() != MCRUN_EMAIL) {
    throw new Error('Please switch to the McRUN Google Account before sending emails');
  }

  // Otherwise send email with extracted stats
  activity['points'] = getEventPointsInRow_(row);

  // Extract email and store in arr
  const recipientArr =
    attendees.split('\n').reduce((acc, entry) => {
      const [, email] = entry.split(':');
      acc.push(email);
      return acc;
    }, []
  );

  // Print log and save return status of `emailMemberStats`
  console.log(activityStats);
  logStatus_(returnStatus, logSheet, row);
  Logger.log(`Successfully executed 'sendStatsEmail' and logged messages in sheet`);
}



function hideAllAttendeeEmail() {
  const sheet = ATTENDANCE_SHEET;
  const startRow = 2  // Skip header row
  const numRows = sheet.getLastRow() - 1;   // Remove header row from count
  const endRow = startRow + numRows;

  for (var row = startRow; row < endRow; row++) {
    hideAttendeeEmailInRow_(row);
  }
}


function hideAttendeeEmailInRow_(row = ATTENDANCE_SHEET.getLastRow()) {
  Object.values(ATTENDEE_MAP).forEach(col => hideAttendeeEmailInCell_(col, row));
}


function hideAttendeeEmailInCell_(column, row = ATTENDANCE_SHEET.getLastRow()) {
  const sheet = ATTENDANCE_SHEET;
  const lastRow = ATTENDANCE_SHEET.getLastRow();

  const attendeeRange = sheet.getRange(row, column);
  const cellValue = attendeeRange.getValue();

  if (!cellValue || cellValue === EMPTY_ATTENDEE_FLAG) return;   // No attendees for this level

  // Get the cell's background color
  const banding = attendeeRange.getBandings()[0];   // Only 1 banding
  const bandingColours = {
    'colourEvenRow': banding.getFirstRowColorObject(),
    'colourOddRow': banding.getSecondRowColorObject(),
    'colourFooter': banding.getFooterRowColorObject(),

    getColour: function (row) {
      return this.colourOddRow;
      if (row == lastRow) { return this.colourFooter }
      else if (row % 2 == 0) { return this.colourEvenRow }
      else { return this.colourOddRow }
    }
  }

  let cellBackgroundColour = bandingColours.getColour(row);

  // Create a RichTextValueBuilder for the cell
  const richTextBuilder = SpreadsheetApp.newRichTextValue().setText(cellValue);
  const isRegisteredTextStyle = SpreadsheetApp.newTextStyle()
    .setItalic(true)
    .setForegroundColorObject(cellBackgroundColour)
    .build()
    ;
  const isUnregisteredTextStyle = SpreadsheetApp.newTextStyle()
    .setForegroundColor('red')
    .build()
    ;

  // Split the cell value by line breaks
  const lines = cellValue.split("\n");

  // Iterate through each line and format the email portion
  let currentIndex = 0;
  const delimiter = ":";

  for (const line of lines) {
    const delimiterIndex = line.indexOf(delimiter);

    if (delimiterIndex !== -1) {
      // Find the email (after the delimiter)
      const email = line.substring(delimiterIndex + 1).trim();
      if (email) {
        const start = currentIndex + delimiterIndex; // Start index of delimiter
        const end = start + email.length + 1; // End index of the email

        // Apply text color and italic formatting to the email
        richTextBuilder.setTextStyle(start, end, isRegisteredTextStyle);
      }
    }
    else {
      const start = currentIndex;
      const end = start + line.length + (line.includes('\n') ? 1 : 0);
      richTextBuilder.setTextStyle(start, end, isUnregisteredTextStyle);
    }
    // Update currentIndex to account for the line length and newline character
    currentIndex += line.length + 1;
  }

  // Build and set the RichTextValue for the cell
  const richTextValue = richTextBuilder.build();
  attendeeRange.setRichTextValue(richTextValue);
}
