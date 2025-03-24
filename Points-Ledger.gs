/**
 * Appends email to attendee name if found. Otherwise, do not add to name.
 *
 * Loops through all levels found in `row`. Sets new cell values in the end.
 *
 * @param {integer} row  Row in `ATTENDANCE_SHEET` to append email.
 * @param {string[][]} memberMap  All search keys of registered members (sorted) and emails.
 *
 * @author [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>) & ChatGPT
 * @date  Dec 14, 2024
 * @update  Feb 9, 2025
 */

function appendMemberEmail(row, memberMap) {
  const sheet = ATTENDANCE_SHEET;
  const numRowToGet = 1;
  const numColToGet = LEVEL_COUNT;

  // Get attendee range starting from beginner col to advanced col
  const attendeeRange = sheet.getRange(row, ATTENDEES_BEGINNER_COL, numRowToGet, numColToGet);  // Attendees columns

  const allAttendees = attendeeRange.getValues()[0]; // Single row of attendees
  const updatedAttendees = [];    // Resulting values to set in sheet

  // Iterate through levels and add emails
  for(let col=0; col < numColToGet; col++) {
    let attendeesInLevel = allAttendees[col].split('\n');  // Split by newline
    
    // Skip levels with no attendees
    if(attendeesInLevel.includes(EMPTY_ATTENDEE_FLAG)) {
      updatedAttendees.push(EMPTY_ATTENDEE_FLAG);
      continue;
    }

    const memberSearchKey = memberMap[SEARCH_KEY_INDEX];


    // Compare last names and check if first name matches any in the list
    //if (attendeeLastName === memberLastName && searchFirstNameList.includes(attendeeFirstName)) {return;}


    // Format each attendee with their email if available
    const formattedAttendee = attendeesInLevel.map(name => {

      const [memberLastName, memberFirstNames] = memberSearchKey.split(",").map(s => s.trim());
      const searchFirstNameList = memberFirstNames.split("|").map(s => s.trim());   // only if preferredName exists

      if (name in memberMap) {
        return `${name}:${memberMap[name]}`;
      }
      return name; // Leave the name as-is if no email found
    });

    // Join back into a string and add to the results
    updatedAttendees.push(formattedAttendee.join('\n'));
  }

  // Write the updated attendees back to the sheet
  attendeeRange.setValues([updatedAttendees]);
}


function hideAllAttendeeEmail() {
  const sheet = ATTENDANCE_SHEET;
  const startRow = 2  // Skip header row
  const numRows = sheet.getLastRow() - 1;   // Remove header row from count
  const endRow = startRow + numRows;

  for(var row = startRow; row < endRow; row++) {
    hideAttendeeEmailInRow_(row);
  }
}


function hideAttendeeEmailInRow_(row = ATTENDANCE_SHEET.getLastRow()) {
  Object.values(ATTENDEE_MAP).forEach(col => hideAttendeeEmailInCell_(col, row));
}


function hideAttendeeEmailInCell_(column, row=ATTENDANCE_SHEET.getLastRow()) {
  const sheet = ATTENDANCE_SHEET;
  const lastRow = ATTENDANCE_SHEET.getLastRow();

  const attendeeRange = sheet.getRange(row, column);
  const cellValue = attendeeRange.getValue();

  if(!cellValue || cellValue === EMPTY_ATTENDEE_FLAG) return;   // No attendees for this level

  // Get the cell's background color
  const banding = attendeeRange.getBandings()[0];   // Only 1 banding
  const bandingColours = {
    'colourEvenRow': banding.getFirstRowColorObject(),
    'colourOddRow' : banding.getSecondRowColorObject(),
    'colourFooter' : banding.getFooterRowColorObject(),

    getColour : function(row) {
      return this.colourOddRow;
      if(row == lastRow)     {return this.colourFooter}
      else if(row % 2 == 0)  {return this.colourEvenRow}
      else                   {return this.colourOddRow}
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

    if(delimiterIndex !== -1) {
      // Find the email (after the delimiter)
      const email = line.substring(delimiterIndex + 1).trim();
      if(email) {
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


function transferAllSubmissions() {
  const startRow = 50; // ATTENDANCE_SHEET.getLastRow()

  for (let row = startRow; row > 1; row--) {
    transferSubmissionToLedger(row);
  }
}


function transferSubmissionToLedger(row=ATTENDANCE_SHEET.getLastRow()) {
  const sheet = ATTENDANCE_SHEET;

  // `Points Ledger` Google Sheet
  const sheetURL = LEDGER_URL;
  const ss = SpreadsheetApp.openByUrl(sheetURL);
  const ledgerSheet = ss.getSheetByName(LEDGET_SHEET_NAME);
  var ledgerLastRow = ledgerSheet.getLastRow();   // Increment per event transfer

  // Select columns to transfer from `sheet`
  const startCol = TIMESTAMP_COL;
  const numCol = DISTANCE_COL - startCol + 1;  // GSheet is 1-indexed
  const numRow = 1;

  // Range is `EMAIL_COL` to `DISTANCE_COL`
  // Save values in 0-indexed array, then transform into 1-indexed by appending empty
  // string to the front. Now, access is easier e.g [EMAIL_COL] vs [EMAIL_COL-1]
  const values = sheet.getSheetValues(row, startCol, numRow, numCol)[0];
  values.unshift("");   // append "" to front

  const formattedNow = Utilities.formatDate(new Date(), TIMEZONE, 'yyyy-MM-dd HH:mm');

  // TODO : APPEND HEADRUN LEVEL TO EVENT NAME (USED TO DETERMINE POINTS TO GIVE)

  const allAttendeesCol = Object.values(ATTENDEE_MAP)
    .filter(level => !values[level].includes(EMPTY_ATTENDEE_FLAG)   // Skip levels with "None"
  );

  for(var level of allAttendeesCol) {
    // Format in `Event Log` sheet in `Points Ledger`
    // Import-Timestamp   Event   Event-TS   MemberEmail   Distance   Points
    const eventToTransfer = [
      formattedNow,           // Import Timestamp
      values[HEADRUN_COL],    // Event name
      values[TIMESTAMP_COL],  // Event Timestamp
      values[level],          // Member Emails
      values[DISTANCE_COL],   // Distance
      // Note: Points col added in `Points Ledger`
    ]

    const colSizeOfTransfer = eventToTransfer.length;
    const rangeNewLog = ledgerSheet.getRange(ledgerLastRow++, 1, 1, colSizeOfTransfer);
    rangeNewLog.setValues([eventToTransfer]);
  }
}

