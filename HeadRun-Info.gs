// Emails of current execs
const emailPresident = 'alexis.demetriou@mail.mcgill.ca';
const emailVPinternal = 'emmanuelle.blais@mail.mcgill.ca';

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

function getHeadRunString(headRunDay) {
  switch(headRunDay) {
    case 'TuesdayPM'  : return 'Tuesday - 6pm';
    case 'WednesdayPM': return 'Wednesday - 6pm';
    case 'ThursdayAM' : return 'Thursday - 7:30am';
    case 'SaturdayAM' : return 'Saturday - 10am';
    case 'SundayPM'   : return 'Sunday - 6pm';

  default : return '';
  }

}

/**
 * Wrapper function for `formatHeadRunnerInRow` to apply on *ALL* submissions.
 * 
 * Row number is 1-indexed in GSheet. Header row skipped.
 */

function formatAllHeadRunner() {
  const sheet = ATTENDANCE_SHEET;
  const startRow = 2  // Skip header row
  const numRow = sheet.getLastRow() - 1;  // Remove header row from count

  formatHeadRunnerInRow_(startRow, numRow);
}


/**
 * Formats headrunner names from `row` into uniform view and separated by newline.
 * 
 * New format is `${firstName} ${lastNameLetter}.`
 *
 * @param {integer} [row=ATTENDANCE_SHEET.getLastRow()]  The row in the `ATTENDANCE_SHEET` sheet (1-indexed).
 *                                                       Defaults to the last row in the sheet.
 * 
 * @author [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>)
 * @date  Dec 10, 2024
 * @update  Dec 11, 2024
 * 
 * ```javascript
 * // Sample Script ➜ Format names in row `7`.
 * const rowToFormat = 7;
 * formatHeadRunnerInRow(rowToFormat);
 * ```
 */

function formatHeadRunnerInRow_(startRow=ATTENDANCE_SHEET.getLastRow(), numRow=1) {
  const sheet = ATTENDANCE_SHEET;
  const headrunnerCol = HEADRUNNERS_COL;
  
  // Get all the values in `HEADRUNNERS_COL`
  const rangeHeadRunner = sheet.getRange(startRow, headrunnerCol, numRow);
  const rawValues = rangeHeadRunner.getValues();

  // Split by commas or newline characters and clean up each name
  const cleanNames = rawValues.map(row => {
    let headrunners = row[0].split(/[,|\n]+/); 
    let ret = headrunners.map(name => name
      .trim()
      .normalize("NFD").replace(/[\u0300-\u036f]/g, "")  // Strip accents
      .toLowerCase()
      .replace(/\b\w/g, letter => letter.toUpperCase())   // Capitalize each word
    );

    return ret;
  });

  const formattedNames = cleanNames.map(names => {
    let ret = names.map(name => {
      const [firstName, lastName = ""] = name.split(' ');  // Defaults lastName to empty string
      const lastNameLetter = lastName ? lastName.charAt(0).toUpperCase() : '';
      const retString = `${firstName} ${lastNameLetter}.`;
      return retString.trim();
    }).join('\n'); // Join names with newlines

    return [ret];   // Every row must be 1d array for setValues
  });

  // Replace with formatted names
  rangeHeadRunner.setValues(formattedNames);
}


/**
 * Returns the headrunners' emails according to input `headrun`.
 * 
 * @param {string}  headrun  The headrun code representing specific headrun (e.g., `'SundayPM'`).
 * @return {string[]}  Array of headrunner emails for respective headrun. (e.g., `['headrunner1@example.com',           
 *                     'headrunner2@example.com', ...]`)
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
  const abigailFinch = 'abigail.finch@mail.mcgill.ca';
  const aidenLee = 'jihong.lee@mail.mcgill.ca';
  const alexanderHebiton = 'alexander.hebiton@mail.mcgill.ca';
  const ameliaRilling = 'amelia.rilling@mail.mcgill.ca';
  const bridgetAndersen = 'bridget.andersen@mail.mcgill.ca';
  const camilaCognac = 'camila.cognac@mail.mcgill.ca';
  const charlesVillegas = 'charles.villegas@mail.mcgill.ca';
  const edmundPaquin = 'edmund.paquin@mail.mcgill.ca';
  const emersonDarling = 'emerson.darling@mail.mcgill.ca';
  const filipSnitil = 'filip.snitil@mail.mcgill.ca';
  const bellaVignuzzi = 'isabella.vignuzzi@mail.mcgill.ca';
  const jamesDiPaola = 'james.dipaola@mail.mcgill.ca';
  const julietteAdeline = 'juliette.adeline@mail.mcgill.ca';
  const justinCote = 'justin.cote2@mail.mcgill.ca';
  const kateRichards = 'katherine.richards@mail.mcgill.ca';
  const lakshyaSethi = 'lakshya.sethi@mail.mcgill.ca';
  const liamGrant = 'liam.grant@mail.mcgill.ca';
  const liamMurphy = 'liam.murphy3@mail.mcgill.ca';
  const madisonHughes = 'madison.hughes@mail.mcgill.ca';
  const michaelRafferty = 'michael.rafferty@mail.mcgill.ca';
  const nicolasMorrison = 'nicolas.morrison@mail.mcgill.ca';
  const pooyaPilehChiha = 'pooya.pilehchiha@mail.mcgill.ca';
  const prabhjeetSingh = 'prabhjeet.singh@mail.mcgill.ca';
  const rachelMattingly = 'rachel.mattingly@mail.mcgill.ca';
  const roriSa = 'rori.sa@mail.mcgill.ca';
  const sophiaLongo = 'sophia.longo@mail.mcgill.ca';
  const tessLedieu = 'tess.ledieu@mail.mcgill.ca';
  const theoGhanem = 'theo.ghanem@mail.mcgill.ca';

  // Head Runners associated to each head run
  const tuesdayHeadRunner = [
    tessLedieu,
    julietteAdeline,
    jamesDiPaola, 
    michaelRafferty, 
    liamMurphy, 
    bridgetAndersen
    ];

  const wednesdayHeadRunner = [
    kateRichards, 
    nicolasMorrison, 
    sophiaLongo, 
    camilaCognac, 
    alexanderHebiton
    ];

  const thursdayHeadRunner = [
    charlesVillegas, 
    ameliaRilling, 
    emersonDarling, 
    justinCote, 
    liamGrant
    ];

  const saturdayHeadRunner = [
    abigailFinch, 
    rachelMattingly, 
    filipSnitil, 
    theoGhanem, 
    bellaVignuzzi, 
    lakshyaSethi
    ];

  const sundayHeadRunner = [
    prabhjeetSingh, 
    edmundPaquin, 
    roriSa, 
    madisonHughes, 
    pooyaPilehChiha, 
    aidenLee
    ];

  // Easier to decode from input `headrun`
  switch (headrun) {
  case 'TuesdayPM'   : return tuesdayHeadRunner;
  case 'WednesdayPM': return wednesdayHeadRunner;
  case 'ThursdayAM' : return thursdayHeadRunner;
  case 'SaturdayAM': return saturdayHeadRunner;
  case 'SundayPM': return sundayHeadRunner;

  default : return '';
  }

}


/**
 * Wrapper function for `formatHeadRunInRow` to apply on *ALL* submissions.
 * 
 * Row number is 1-indexed in GSheet. Header row skipped.
 */

function formatAllHeadRun() {
  const sheet = ATTENDANCE_SHEET;
  const startRow = 2;   // Skip header row
  const numRow = sheet.getLastRow() - 1;  // Remove header row from count

  formatHeadRunInRow_(startRow, numRow);    // This formats sheet from first submission to last row
}

/**
 * Removes hyphen-space in headrun from `row` if applicable.
 * 
 * @param {integer} [startRow=ATTENDANCE_SHEET.getLastRow()]  
 *                      The row in the `ATTENDANCE_SHEET` sheet (1-indexed).
 *                      Defaults to the last row in the sheet.
 * 
 * @param {integer} numRow  Number of rows to format from `startRow`
 * 
 * @author [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>)
 * @date  Dec 10, 2024
 * @update  Dec 11, 2024
 */

function formatHeadRunInRow_(startRow=ATTENDANCE_SHEET.getLastRow(), numRow=1) {
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

