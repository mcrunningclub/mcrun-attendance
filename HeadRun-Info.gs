// Emails of current execs
const PRESIDENT_EMAIL = 'alexis.demetriou@mail.mcgill.ca';
const VP_INTERNAL_EMAIL = 'emmanuelle.blais@mail.mcgill.ca';

/**
 * Return headrun day and time from headrun code input `headRunDay`.
 *
 * @param {string}  headRunDay  The headrun code representing specific headrun (e.g., `'SundayPM'`).
 * @return {string}  String of headrun day and time. (e.g., `'Sunday - 6pm'`)
 *
 * Current head runs for semester:
 *
 * Tuesday   :  6:00pm
 * Wednesday :  6:00pm
 * Thursday  :  7:30am
 * Saturday  :  10:00am
 * Sunday    :  6:00pm
 *
 * @author [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>)
 * @date  Nov 13, 2023
 * @update  Sep 24, 2024
 *
 * ```javascript
 * // Sample Script ➜ Getting headrun datetime for Sunday evening run.
 * const headrun = getHeadRunnerEmail('SundayPM');
 * Logger.log(headrun) // 'Sunday - 6pm'
 * ```
 */

function getHeadrunTitle(headRunDay) {
  switch (headRunDay) {
    case 'TuesdayPM': return 'Tuesday - 6:00pm';
    case 'WednesdayPM': return 'Wednesday - 6:00pm';
    case 'ThursdayAM': return 'Thursday - 7:30am';
    case 'SaturdayAM': return 'Saturday - 10:00am';
    case 'SundayPM': return 'Sunday - 6:00pm';

    default : throw new Error(`No headrunner has been found for ${headRunDay}`);
  }
}


/**
 * Wrapper function for `formatHeadRunnerInRow` to apply on *ALL* submissions.
 *
 * Row number is 1-indexed in GSheet. Header row skipped. Top-to-bottom execution.
 */

function formatAllHeadRunner() {
  runOnSheet_(formatHeadRunnerInRow_.name);
}

/**
 * Formats headrunner names from `row` into uniform view, separated by newline.
 *
 * Updated format is '`${firstName} ${lastNameLetter}.`'
 *
 * @param {integer} [row=ATTENDANCE_SHEET.getLastRow()]  The row in the `ATTENDANCE_SHEET` sheet (1-indexed).
 *                                                       Defaults to the last row in the sheet.
 *
 * @param {integer} numRow  Number of rows to format from `startRow`.
 *
 * @author [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>)
 * @date  Dec 10, 2024
 * @update  Dec 11, 2024
 *
 * ```javascript
 * // Sample Script ➜ Format names in row `7`.
 * const rowToFormat = 7;
 * formatHeadRunnerInRow(rowToFormat);
 *
 * // Sample Script ➜ Format names from row `3` to `9`.
 * const startRow = 3;
 * const numRow = 9 - startRow;
 * formatHeadRunnerInRow(startRow, numRow);
 * ```
 */

function formatHeadRunnerInRow_(startRow = ATTENDANCE_SHEET.getLastRow(), numRow = 1) {
  const sheet = ATTENDANCE_SHEET;
  const headrunnerCol = HEADRUNNERS_COL;

  // Get all the values in `HEADRUNNERS_COL` in bulk
  const rangeHeadRunner = sheet.getRange(startRow, headrunnerCol, numRow);
  const rawValues = rangeHeadRunner.getValues();

  // Callback function to clean and format a single headrunner name
  function formatName(name) {
    const cleanedName = name
      .trim()
      .normalize("NFD").replace(/[\u0300-\u036f]/g, "") // Remove accents
      .toLowerCase()
      .replace(/\b\w/g, letter => letter.toUpperCase()); // Capitalize each proper name

    // Split into first and last names
    const [firstName, lastName = ""] = cleanedName.split(' ');
    const lastInitial = lastName.charAt(0).toUpperCase();  // Get first letter of last name
    return `${firstName} ${lastInitial}.`;  // Return formatted name
  };

  // Callback function to process the raw value into the formatted format
  function processRow(row) {
    const headrunners = row[0]  // Get first column from 2D array
      .split(/[,|\n]+/)         // Split by commas or newlines
      .map(formatName)   // Format each name using formatName()
      .join('\n');       // Join the names with a newline

    return [headrunners]; // Return as a 2D array for .setValues()
  };

  // Map over each row to process and format by applying `processRow()`
  const formattedNames = rawValues.map(processRow);   // apply processRow()

  // Update the sheet with formatted names
  rangeHeadRunner.setValues(formattedNames);
  console.log(`Completed formatting of headrunner names`,  formattedNames);
}


/**
 * Returns the headrunners' emails according to input `headrun`.
 *
 * @param {string}  headrun  The headrun code representing specific headrun (e.g., `'SundayPM'`).
 * @return {string[]}  Array of headrunner emails for respective headrun.
 *                      (e.g., `['headrunner1@example.com', 'headrunner2@example.com', ...]`)
 *
 * @author [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>)
 * @date  Nov 13, 2023
 * @update  Sep 29, 2024
 *
 * ```javascript
 * // Sample Script ➜ Getting headrunner emails for Sunday evening run.
 * const headrunnerEmails = getHeadRunnerEmail('SundayPM');
 * ```
 */

function getHeadRunnerEmail(headrun) {
  // Head Runner Emails
  const aidenLee = 'jihong.lee@mail.mcgill.ca';
  const alyssaAbouChakra = 'alyssa.abouchakra@mail.mcgill.ca';
  const camilaCognac = 'camila.cognac@mail.mcgill.ca';
  const charlesVillegas = 'charles.villegas@mail.mcgill.ca';
  const edmundPaquin = 'edmund.paquin@mail.mcgill.ca';
  const isabellaVignuzzi = 'isabella.vignuzzi@mail.mcgill.ca';
  const kateRichards = 'katherine.richards@mail.mcgill.ca';
  const liamGrant = 'liam.grant@mail.mcgill.ca';
  const liamMurphy = 'liam.murphy3@mail.mcgill.ca';
  const lizzyVreendeburg = 'elizabeth.vreedenburgh@mail.mcgill.ca';
  const michaelRafferty = 'michael.rafferty@mail.mcgill.ca';
  const sachiKapoor = 'sachi.kapoor@mail.mcgill.ca';
  const sophiaLongo = 'sophia.longo@mail.mcgill.ca';
  const theoGhanem = 'theo.ghanem@mail.mcgill.ca';
  const zishengHong = 'zisheng.hong@mail.mcgill.ca';


  // Head Runners associated to each head run
  const tuesdayHeadRunner = [
    kateRichards,
    liamMurphy,
    zishengHong,
  ];

  const wednesdayHeadRunner = [
    lizzyVreendeburg,
    edmundPaquin,
    sophiaLongo,
    michaelRafferty,
  ];

  const thursdayHeadRunner = [
    alyssaAbouChakra,
    sachiKapoor,
    liamGrant,
  ];

  const saturdayHeadRunner = [
    michaelRafferty,
    liamMurphy,
    isabellaVignuzzi,
    theoGhanem,
    liamGrant,
  ];

  const sundayHeadRunner = [
    charlesVillegas,
    kateRichards,
    edmundPaquin,
    sophiaLongo,
    camilaCognac,
    aidenLee,
  ];

  const thisHeadrun = headrun.toLowerCase();
  // Easier to decode from input `headrun`
  switch (thisHeadrun) {
    case 'tuesdaypm': return tuesdayHeadRunner;
    case 'wednesdaypm': return wednesdayHeadRunner;
    case 'thursdayam': return thursdayHeadRunner;
    case 'saturdayam': return saturdayHeadRunner;
    case 'sundaypm': return sundayHeadRunner;

    default: throw Error(`No headrun found for ${thisHeadrun}`);
  }
}


/**
 * Wrapper function for `formatHeadRunInRow` to apply on *ALL* submissions.
 *
 * Row number is 1-indexed in GSheet. Header row skipped. Top-to-bottom execution.
 */

function formatAllHeadRun() {
  runOnSheet_(formatHeadRunInRow_.name);
}

/**
 * Removes hyphen-space in headrun from `row` if applicable.
 *
 * @param {integer} [startRow=ATTENDANCE_SHEET.getLastRow()]
 *                      The row in the `ATTENDANCE_SHEET` sheet (1-indexed).
 *                      Defaults to the last row in the sheet.
 *
 * @param {integer} [numRow=1] Number of rows to format from `startRow`
 *
 * @author [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>)
 * @date  Dec 10, 2024
 * @update  Dec 11, 2024
 */

function formatHeadRunInRow_(startRow = ATTENDANCE_SHEET.getLastRow(), numRow = 1) {
  const sheet = ATTENDANCE_SHEET;
  const headrunCol = HEADRUN_COL;

  // Get the cell value, and remove hyphen-space in each cell
  const rangeToFormat = sheet.getRange(startRow, headrunCol, numRow);
  var values = rangeToFormat.getValues();

  // Bulk format if applicable
  var formattedHeadRun = values.map(row => {
    let cleanValue = row[0].toString().replace(/- /g, "");
    return [cleanValue] // must return as 2d
  });

  // Replace with formatted value
  rangeToFormat.setValues(formattedHeadRun);  // setValues requires 2d array
}
