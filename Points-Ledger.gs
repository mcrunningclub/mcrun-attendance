/**
 * Appends email to attendee name if found. Otherwise, do not add to name.
 *
 * Loops through all levels found in `row`. Sets new cell values in the end.
 *
 * @param {integer} row  Row in `ATTENDANCE_SHEET` to append email.
 * @param {string[][]} registered  All search keys of registered members (sorted) and emails.
 *
 * @author [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>) & ChatGPT
 * @date  Dec 14, 2024
 * @update  April 1, 2025
 */

function appendMemberEmail_(row, registered, unregistered) {
  const sheet = ATTENDANCE_SHEET;
  const numRowToGet = 1;
  const numColToGet = LEVEL_COUNT;

  // Get attendee range starting from beginner col to advanced col
  const attendeeRange = sheet.getRange(row, ATTENDEES_BEGINNER_COL, numRowToGet, numColToGet);  // Attendees columns

  //const allAttendees = attendeeRange.getValues()[0]; // Single row of attendees
  const updatedAttendees = [];    // Resulting values to set in sheet

  // Iterate through levels and add emails
  for (let col = 0; col < numColToGet; col++) {
    const levelRegistered = registered[col];
    const levelUnregistered = unregistered[col];

    const registeredCount = levelRegistered ? levelRegistered.length : 0;
    const unregisteredCount = levelUnregistered ? levelUnregistered.length : 0;

    // Skip levels with no attendees
    if (registeredCount === 0 && unregisteredCount === 0) {
      updatedAttendees.push(EMPTY_ATTENDEE_FLAG);
      continue;
    }
    // Update these members since no unregistered
    else if (unregisteredCount === 0) {
      updatedAttendees[col] = levelRegistered.join('\n');
      continue;
    }
    else if (registeredCount === 0) {
      updatedAttendees[col] = levelUnregistered.join('\n');
      continue;
    }
    // Merge attendees in `col` and sort before setting back in sheet
    const attendees = [...levelRegistered, ...levelUnregistered].sort();

    // Join back into a string and add to the results
    updatedAttendees.push(attendees.join('\n'));
  }

  // Write the updated attendees back to the sheet
  attendeeRange.setValues([updatedAttendees]);
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


function transferSubmissionToLedger(row = getLastSubmission_()) {
  // STEP 1: Package all non-empty submission levels in single 2d arr
  const packagedEvents = packageRowForLedger_(row);

  // STEP 1b: Only transfer if attendees count > 0
  if (packagedEvents.length === 0) return;

  // STEP 2: Send submission using library and store new row index.
  // This triggers automations in the recipient sheet.
  try {
    const logNewRow = sendNewSubmission_(packagedEvents);
    storeStravaInLogSheet_(logNewRow);
  }

  // STEP 2b: Error occured, send using `openByUrl`. Downside: automations not triggered
  catch (e) {
    Logger.log(e);    // Display error message from 'sendNewSubmission'
    Logger.log(`Unable to transfer submission with library. Now trying with 'openByUrl'...`);

    // `Points Ledger` Google Sheet
    const ss = SpreadsheetApp.openByUrl(POINTS_LEDGER_URL);
    const logSheet = ss.getSheetByName(LOG_SHEET_NAME);
    const logNewRow = logSheet.getLastRow() + 1;

    const packageNumRows = packagedEvents.length;
    const packageNumCols = packagedEvents[0].length;

    // Set values using defined range dimensions
    logSheet.getRange(logNewRow, 1, packageNumRows, packageNumCols).setValues(packagedEvents);

    // Display successful message for Step 2b, and error message from Step 2.
    Logger.log(`Successfully transferred event attendance submission to Ledger row ${logNewRow}`);
  }
}


function packageRowForLedger_(row) {
  const sheet = GET_ATTENDANCE_SHEET_();

  // Define dimenstion of range
  const startCol = TIMESTAMP_COL;
  const numCols = DISTANCE_COL - startCol + 1;

  // Fetch values from the row, convert to 1-indexed by unshifting
  // Access is easier e.g [EMAIL_COL] vs [EMAIL_COL-1]
  const rowValues = sheet.getSheetValues(row, startCol, 1, numCols)[0];
  rowValues.unshift("");   // Padding for 1-indexed access

  // Identify attendee columns with actual data (not marked "None" or empty)
  const validAttendeeCols = Object.values(ATTENDEE_MAP).filter(level => {
    const attendee = rowValues[level];
    return attendee && !attendee.includes(EMPTY_ATTENDEE_FLAG);
  });

  // Build list of events for all valid attendees
  const exportTimestamp = Utilities.formatDate(new Date(), TIMEZONE, 'yyyy-MM-dd HH:mm:ss');
  const eventLabel = `Headrun ${rowValues[RUN_LEVEL_COL]}\n${rowValues[HEADRUN_COL]}`;
  const eventTimestamp = rowValues[TIMESTAMP_COL];
  const distance = rowValues[DISTANCE_COL];

  const events = validAttendeeCols.map(colIndex => [
    exportTimestamp,        // Export Timestamp
    eventLabel,             // Event Name
    eventTimestamp,         // Event Timestamp
    rowValues[colIndex],    // Member Name + Email
    distance                // Distance
    // Points will be added in recipient sheet
  ]);

  return events;
}


function sendNewSubmission_(submissionArr) {
  const funcName = PointsLedgerCode.storeImportFromAttendanceSheet.name;
  return executePointsLedgerFunction_(funcName, [submissionArr]);
}

function storeStravaInLogSheet_(logRow) {
  const funcName = PointsLedgerCode.findAndStoreStravaActivity.name;
  return executePointsLedgerFunction_(funcName, [logRow]);
}

function triggerEmailInLedger_(logRow) {
  const funcName = PointsLedgerCode.sendStatsEmail.name;
  executePointsLedgerFunction_(funcName, [undefined, logRow]);
}

function executePointsLedgerFunction_(funcName, args) {
  console.log(`\n---START OF '${funcName}()' LOG MESSAGES\n\n`);
  const retValue = PointsLedgerCode[funcName].apply(args);
  console.log(`\n---END OF '${funcName}()' LOG MESSAGES\n\n`);
  return retValue;
}

