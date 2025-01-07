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
 * @update  Jan 7, 2025
 */

function pointsEmail() {
  const sheet = ATTENDANCE_SHEET;
  const lastRow = sheet.getLastRow();

  // if (getCurrentUserEmail() != 'mcrunningclub@ssmu.ca') return;   // prevent email sent by wrong user

  const points = SpreadsheetApp.openByUrl(LEDGER_URL).getSheetByName("Member Points");

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
  const pointsData = points.getRange(2, 1, points.getLastRow() - 1, 6).getValues();
  
  // Create a mapping of full names to points
  const pointsMap = {};
  pointsData.forEach(([email, , , ,fullName, points]) => {
    pointsMap[fullName.trim()] = [email, points]; // Store points with full name as the key
  });

  // Loop through the full names array and email that member regarding their current points
  uniqueRecipientsArray.forEach(fullName => {
    const trimmedName = fullName.trim();

    if (!pointsMap[trimmedName]) return;     // skips to next iteration if no email is found

    const points = pointsMap[trimmedName][1] ?? 0;
    const email = pointsMap[trimmedName][0]; // Get email for the full name
    const firstName = trimmedName.split(" ")[0];

    if (email) {
      // Construct and send the email
      const subject = `Your Points Update`;

      const pointsEmailHTML = `
        <!DOCTYPE html>
        <html>
          <body style="font-family: Arial, Helvetica, sans-serif;">
            <div style="margin: 10% 5%; text-align: center">
                <img src="https://mcrun.ssmu.ca/wp-content/uploads/2023/04/McRun-Logo-Circular.png" style="width: 20%" >
                <div style="text-align: left; width: max-content; margin: 0 auto">
                <h1>Hello, ${firstName}!</h1>
                <h3>You currently have:</h3>
                <h2>${points} points</h2>
                <p>Thanks for running with us, hope you keep up the great pace!</p>
                <p>- McGill Students Running Club</p>
                </div>
            </div>
          </body>
        </html>
      `;
      
      
      MailApp.sendEmail({
        to: email,
        subject: subject,
        htmlBody: pointsEmailHTML
      });
      

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

