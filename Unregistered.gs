/**
 * Wrapper function for `getUnregisteredMembers` for *ALL* rows.
 *
 * Executes the function for all rows in the attendance sheet, skipping the header row.
 */

function getAllUnregisteredMembers_() {
  runOnSheet_(getUnregisteredMembersInRow_.name);
}


/**
 * Find attendees in a specific row of the attendance sheet who are unregistered members.
 *
 * Sets unregistered members in the `NOT_FOUND_COL` column.
 * 
 * List of members found in `Members` sheet.
 *
 * @param {number} [row=ATTENDANCE_SHEET.getLastRow()] - The row number in the attendance sheet (1-indexed).
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

  /** Helper Function */
  function setNamesNotFound_(row, notFoundArr) {
    const sheet = ATTENDANCE_SHEET;
    const unfoundNameRange = sheet.getRange(row, NAMES_NOT_FOUND_COL);
    unfoundNameRange.setValue(notFoundArr);
  }
}


/**
 * Helper function to find unregistered attendees.
 *
 * Compares attendees against the member map to identify unregistered members.
 *
 * @param {string[]} attendees - All attendees of the head run (sorted).
 * @param {string[][]} memberMap - All search keys of registered members (sorted) and emails.
 * @return {Object} - An object containing registered and unregistered attendees.
 *                    { registered: string[], unregistered: string[] }
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


/** HELPER FUNCTIONS TO FORMAT */

// Group functions to apply on `input`
function compose_(...fns) {
  return (input) => fns.reduce((v, f) => f(v), input);
}


/**
 * Formats a name by removing whitespace, stripping accents, and capitalizing names.
 *
 * @param {string} name - The name to format.
 * @return {string} - The formatted name.
 */

function formatThisName_(name) {
  return name
    .trim()
    .normalize("NFD").replace(/[\u0300-\u036f]/g, "")  // Strip accents
    .replace(/[\u2018\u2019']/g, "") // Remove apostrophes (`, ', ’)
    .toLowerCase()
    .replace(/\b\w/g, l => l.toUpperCase());  // Capitalize each name
}


/**
 * Reverses the order of a name to `LastName, FirstName` format.
 *
 * @param {string} name - The name to reverse.
 * @return {string} - The reversed name.
 */

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


/**
 * Formats and sorts all entries in the member map by search key.
 *
 * Removes whitespace, hyphens, and accents, and capitalizes names.
 *
 * @param {string[][]} memberMap - Array of search keys and their emails.
 * @param {number} searchKeyIndex - The index of the search key in the member map.
 * @param {number} emailIndex - The index of the email in the member map.
 * @return {string[][]} - A sorted array of formatted names and emails.
 *
 * @author [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>)
 * @date  Nov 1, 2024
 * @update  Dec 14, 2024
 *
 * ```js
 * // Sample Script ➜ Format, then sort names.
 * const rawData = [["Francine de-Blé", "francine.de-ble@mail.com"],
 *                  ["BOb-Burger belChEr ", "bob.belcher@mail.com"]];
 * const result = formatAndSortMemberMap_(rawData);
 * Logger.log(result)  // [["Bob Burger Belcher", "bob.belcher@mail.com"],
 *                         [ "Francine De ble", "francine.de-ble@mail.com"]]
 * ```
 */
function formatAndSortMemberMap_(memberMap, searchKeyIndex, emailIndex) {
  const formattedMap = memberMap.map(row => {
    const memberEmail = row[emailIndex];
    const formattedSearchKey = row[searchKeyIndex]
      .trim()
      .normalize("NFD").replace(/[\u0300-\u036f]/g, "")   // Strip accents
      .replace(/[\u2018\u2019']/g, "") // Remove apostrophes (`, ', ’)
      .toLowerCase()
      .replace(/\b\w/g, l => l.toUpperCase())   // Capitalize each word

    // Combine formatted searchkey and email
    return [formattedSearchKey, memberEmail];
  });

  // Sort by formatted searchKey
  formattedMap.sort((a, b) => a[0].localeCompare(b[0]));
  return formattedMap;
}


/**
 * Formats and sorts an array of names, swapping last and first names.
 *
 * Removes whitespace, apostrophes, and accents, and capitalizes names.
 *
 * @param {string[]} names - Array of names to format.
 * @return {string[]} - A sorted array of formatted names.
 *
 * @author [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>)
 * @date  Dec 6, 2024
 * @update  Dec 11, 2024
 *
 * ```javascript
 * // Sample Script ➜ Format, swap first and last name, then sort.
 * const rawNames = ["BOb-Burger bulChEr ", "Francine de-Blé"];
 * const result = swapAndFormatName_(rawNames);
 * Logger.log(result)  // ["Bulcher, Bob Burger", "De ble, Francine"]
 * ```
 */

function swapAndFormatName_(names) {
  const formattedNames = names.map(name => {
    let nameParts = name
      .trim()
      .normalize("NFD").replace(/[\u0300-\u036f]/g, "")   // Strip accents
      .replace(/[\u2018\u2019']/g, "")  // Remove apostrophes (`, ', ’)
      .toLowerCase()
      .replace(/\b\w/g, l => l.toUpperCase())   // Capitalize each name
      .split(/\s+/) // Split by spaces
      ;

    // Replace hyphens with spaces. Can only perform after splitting first and last name.
    nameParts = nameParts.map(name => name.replace(/-/g, " "));

    // If first name is not hyphenated, only left-most substring stored in first name
    const firstName = nameParts[0];
    const lastName = nameParts[nameParts.length - 1];
    return `${lastName}, ${firstName}`; // Format as "LastName, FirstName"
  });

  return formattedNames.sort();
}


/**
 * Retrieves the member map from the `Members` sheet.
 *
 * Combines member search keys and emails, filtering out empty rows.
 *
 * @return {string[][]} - An array of member search keys and emails.
 */

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
    .filter(row => row[0] && row[1]);

  return memberMap;
}

