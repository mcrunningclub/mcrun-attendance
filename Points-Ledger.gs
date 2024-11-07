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
  return;
  const sourceSheet = ATTENDANCE_SHEET;
  const ledgerName = LEDGER_NAME;
  const sheetUrl = LEDGER_URL;

  var destinationSpreadsheet = SpreadsheetApp.openByUrl(sheetUrl);
  var destinationSheet = destinationSpreadsheet.getSheetByName(ledgerName);
  var sourceData = sourceSheet.getRange(sourceSheet.getLastRow(), 1, 1, 5).getValues()[0];

  destinationSheet.appendRow(sourceData);
}


/**
 * Function to send email to each member updating them on their points
 * 
 * @trigger The 1st and 14th of every month
 * 
 * @author [Charles Villegas](<charles.villegas@mail.mcgill.ca>) & ChatGPT
 * @date  Nov 5, 2024
 * @update  Nov 5, 2024
 */


function pointsEmail() {
  const sheet = ATTENDANCE_SHEET;
  const lastRow = sheet.getLastRow();

  if (getCurrentUserEmail() != 'mcrunningclub@ssmu.ca') return;   // prevent email sent by wrong user

  const points = SpreadsheetApp.openByUrl(LEDGER_URL).getSheetByName("Member Points");
  const emails = SpreadsheetApp.openByUrl(MEMBERSHIP_URL).getSheetByName("MASTER");

  // Define the columns to check for attendees
  const attendeeColumns = [
    ATTENDEES_BEGINNER_COL, 
    ATTENDEES_INTERMEDIATE_COL, 
    ATTENDEES_ADVANCED_COL
  ];

  // Collect all unique values in one step
  const uniqueRecipients = new Set(
    attendeeColumns.flatMap(level => {
      // Get all values in the current column and split by newline
      return sheet.getRange(2, level, lastRow, 1).getValues()
        .flat() // Flatten the 2D array to 1D
        .map(value => value.split('\n')) // Split by newline
        .flat(); // Flatten the nested arrays
    })
  );

  // Convert the Set to an Array of unique recipients
  const uniqueRecipientsArray = [...uniqueRecipients].map(value => value.trim()).filter(Boolean);

  // Get all names and point values from points, and names and emails from emails
  const pointsData = points.getRange(2, 5, points.getLastRow() - 1, 2).getValues();
  const namesData = emails.getRange(2, 1, emails.getLastRow() - 1, 3).getValues();
  
  // Create a mapping of full names to points
  const pointsMap = {};
  pointsData.forEach(([fullName, points]) => {
    pointsMap[fullName.trim()] = points; // Store points with full name as the key
  });

  // Create a mapping of first and last names to emails
  const emailMap = {};
  namesData.forEach(([email, firstName, lastName]) => {
    const fullName = `${firstName.trim()} ${lastName.trim()}`; // Combine first and last name
    emailMap[fullName] = email; // Store email with full name as the key
  });

  // Loop through the full names array and email that member regarding their current points
  uniqueRecipientsArray.forEach(fullName => {
    const trimmedName = fullName.trim();
    const points = pointsMap[trimmedName] ?? 0;
    const email = emailMap[trimmedName]; // Get email for the full name

    if (email) {
      // Construct and send the email
      const subject = `Your Points Update`;
      const message = `Hello ${trimmedName},\n\nYou have ${points} points.\n\nBest,\nMcGill Students Running Club`;
      
      /*
      MailApp.sendEmail({
        to: email,
        subject: subject,
        body: message
      });
      */

      // log confirmation for the sent email with values for each variable
      Logger.log(`Email sent to ${trimmedName} at ${email} with ${points} points.`);
    } else {
      Logger.log(`No email found for ${trimmedName}.`);
    }
  });
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

