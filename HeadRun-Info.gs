// Emails of current execs
const PRESIDENT_EMAIL = 'alexis.demetriou@mail.mcgill.ca';
const VP_INTERNAL_EMAIL = 'emmanuelle.blais@mail.mcgill.ca';

const CALENDAR_STORE = SCRIPT_PROPERTY.calendarTriggers;
const HEADRUNNER_STORE = 'headrunners';
const HEADRUN_STORE = 'headruns';

const TRIGGER_FUNC = checkMissingAttendance.name;


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

function getHeadrunTitle_(headRunDay) {
  switch (headRunDay) {
    case 'TuesdayPM': return 'Tuesday - 6:00pm';
    case 'WednesdayPM': return 'Wednesday - 6:00pm';
    case 'ThursdayAM': return 'Thursday - 7:30am';
    case 'SaturdayAM': return 'Saturday - 10:00am';
    case 'SundayPM': return 'Sunday - 6:00pm';

    default: throw new Error(`No headrunner has been found for ${headRunDay}`);
  }
}

function getAllHeadruns() {
  const docProp = PropertiesService.getDocumentProperties();
  return docProp.getProperty(HEADRUN_STORE);
}

function getAllHeadrunners() {
  const docProp = PropertiesService.getDocumentProperties();
  return docProp.getProperty(HEADRUNNER_STORE);
}


/** DATA STRUCTURE FOR HEADRUNS AND HEADRUNNERS: */

/*
headruns = {
  sunday: {
    '10:00am' : {level : 'easy', headrunners : [name1, name2, ...]},
    '2:15pm': {level : 'intermediate', headrunners : [name2]}
  ]
}

headrunners = {
  name1: {email1, strava1},
  name2: {email2, strava2}
}
*/


function readAndStoreRunData() {
  const sheet = GET_HEADRUNNER_SHEET_();
  const data = sheet.getDataRange().getValues();

  // Pop header from data
  const header = data.shift();

  // Get columns indices
  const THIS_NAME_COL = getColIndex('name');
  const THIS_EMAIL_COL = getColIndex('email');
  const THIS_STRAVA_COL = getColIndex('strava');
  const THIS_LEVEL_COL = getColIndex('level');

  const headrunInfo = {};

  // Package headrun info
  const headrunners = data.reduce((acc, row) => {
    let info = {
      'email' : row[THIS_EMAIL_COL],
      'strava' : extractStravaId(row[THIS_STRAVA_COL])
    }

    // Key must be identical in both objects (headruns + headrunners)
    const nameKey = createVarName(row[THIS_NAME_COL]);

    // Append (or create) headrun schedule entry to 'headrunInfo'
    appendHeadrunInfo(row, nameKey);
    
    // Add headrunner info and return obj
    acc[nameKey] = info;
    return acc;
  }, {});

  console.log(headrunners);
  console.log(headrunInfo);

  /** Helper function to extract data */
  function getColIndex(target) {
    return header.findIndex(val => val.toLowerCase() == target);
  }

  function createVarName(name) {
    name = name.replace(/[ \-]/g, '');
    return name.charAt(0).toLowerCase() + name.slice(1);
  }

  function extractStravaId(input) {
    const match = input.match(/\/athletes\/(\d+)/);
    const athleteId = match ? match[1] : input;
    return athleteId;
  }

  /*
  let headrunnerSchedule = 'Wednesday 6pm (Beginner); Sunday 8am (Intermediate); Sunday 6pm (Beginner)';
  {
    'wednesday' : { 
      '6pm' : { level: 'beginner', people : []}
    },
    'sunday' : {
      '8am' : { level : 'intermediate', people : [] },
      '6pm' : { level : 'beginner', people : [] }
    }
  }
  */

  function appendHeadrunInfo(row, headrunner) {
    const levels = row[THIS_LEVEL_COL].split(';');
    // Entry info: `[weekday] [time-12h] ([run level])`
    levels.forEach(entry => {
      const match = entry.trim().match(/^(\w+)\s+([\d:apm]+)\s+\((\w+)\)$/i);
      if (match) {
        const [_, day, time, level] = match;    // Get match using dereferencing
        const dayKey = day.toLowerCase();
        const timeKey = time.toLowerCase();
        const levelVal = level.toLowerCase();

        if (!headrunInfo[dayKey]) {
          headrunInfo[dayKey] = {};
        }

        // Add time to dayKey if not found
        if (!headrunInfo[dayKey][timeKey]) {
           headrunInfo[dayKey][timeKey] = {};
        }

        const people = headrunInfo[dayKey][timeKey][headrunner] ?? [];
        headrunInfo[dayKey][timeKey] = { level: levelVal, headrunner: people.concat(headrunner)};
      }
    });
  }
}


function readAndStoreHeadruns() {
  const sheet = GET_COMPILED_SHEET_();
  const data = sheet.getDataRange().getValues();
  console.log(data);
  return;
  
  const headruns = data.shift();
  const count = headruns.length;

  headruns.forEach((hr, i) => {
    if (i < 1) return;
    headrunners = hr[count][i];
    console.log(headruns);
  })
  
  return;


  const docProp = PropertiesService.getDocumentProperties();
  docProp.setProperty(headrunners);
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
 * @update  Apr 7, 2024
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
  const sheet = GET_ATTENDANCE_SHEET_();
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
  console.log(`[AC] Completed formatting of headrunner names`, formattedNames);
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

function getHeadRunnerEmail_(headrun) {
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
 * @update  Apr 7, 2025
 */

function formatHeadRunInRow_(startRow = ATTENDANCE_SHEET.getLastRow(), numRow = 1) {
  const sheet = GET_ATTENDANCE_SHEET_();
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


/**
 * Adds new events as time-based triggers and removed expired ones
 * 
 * @trigger  Every Sunday at 1am.
 */

function updateWeeklyCalendarTriggers() {
  createDailyAttendanceTrigger_();
  deleteExpiredCalendarTriggers_();
}

/**
 * Add new McRUN event from calendar to Apps Script trigger for today.
 * 
 * @trigger  Updated calendar.
 *
 * @author [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>)
 * @date  Apr 17, 2025
 * @update  Apr 17, 2025
 */

function addSingleEventTrigger() {
  const now = new Date();
  const midnight = new Date(new Date().setHours(23, 59, 59, 59));

  const calendar = CalendarApp.getDefaultCalendar();
  const events = calendar.getEvents(now, midnight);
  events.forEach(e => createAndStoreTrigger_(e));
}


/**
 * Get events from calendar and create time-based triggers.
 *
 * @author [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>) + ChatGPT
 * @date  Apr 17, 2025
 * @update  Apr 27, 2025
 */

function createDailyAttendanceTrigger_() {
  const calendar = CalendarApp.getDefaultCalendar();

  const now = new Date();
  const startOfWeek = getStartOfWeek(now); // Sunday
  const endOfWeek = new Date(startOfWeek);
  endOfWeek.setDate(endOfWeek.getDate() + 7); // Saturday end

  const events = calendar.getEvents(startOfWeek, endOfWeek);

  const filteredEvents = events.filter(event =>
    !event.isAllDayEvent() &&
    event.getStartTime() > now
  );

  filteredEvents.forEach(event => createAndStoreTrigger_(event));

  // Helper: Gets the Sunday of the current week
  function getStartOfWeek(date) {
    const start = new Date(date);
    const day = start.getDay(); // 0 = Sunday, 1 = Monday, etc.
    start.setDate(start.getDate() - day);
    start.setHours(0, 0, 0, 0);
    return start;
  }
}


function updateCalendarTriggers() {
  // Get events from day start (midnight)
  const now = new Date();
  const start = new Date().setHours(0, 0, 0, 0);

  const calendar = CalendarApp.getDefaultCalendar();
  const events = calendar.getEvents(start, now);





  const offset = now - 10*60 * 1000;    // Search 6 sec ago

  

  const cancelledRegex = /cancel{1,2}ed/i;

  for (const event of events) {
    if (offset < event.getLastUpdated()) {
      console.log(event.getDescription());
      console.log(event.getTitle);
      console.log(`This event has been cancelled: ${isCancelled(event)}`)
    }
  }

  function isCancelled(event) {
    const str = event.getDescription() + event.getTitle();
    return cancelledRegex.test(str);
  }
}

function mTest() {
  const now = new Date()
  const offset = now - 6 * 1000;    // Search 6 sec ago
}


/**
 * Add time-based trigger using event information from Calendar.
 *
 * @author [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>) + ChatGPT
 * @date  Apr 15, 2025
 * @update  Apr 17, 2025
 */

function createAndStoreTrigger_(event) {
  const props = GET_PROP_STORE_();
  const stored = JSON.parse(props.getProperty(CALENDAR_STORE) || "{}");

  const offset = 60 * 60 * 1000;
  const startTime = new Date(event.getStartTime().getTime() + offset);

  // Only add trigger if new
  if (isExistingTrigger_(startTime, stored)) return;

  const trigger = ScriptApp.newTrigger(TRIGGER_FUNC)
    .timeBased()
    .at(startTime)
    .create();

  stored[trigger.getUniqueId()] = startTime.toISOString();

  // Store updated calendar triggers
  props.setProperty(CALENDAR_STORE, JSON.stringify(stored));
  Logger.log(`Trigger created and stored for "${event.getTitle()}" at ${startTime}`);

  // Helper function
  function isExistingTrigger_(time, stored) {
    const triggerTimes = Object.values(stored);
    return (time in triggerTimes);
  }
}


/**
 * Removes expired calendar triggers and updates store in Properties.
 *
 * @author [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>) + ChatGPT
 * @date  Apr 15, 2025
 * @update  Apr 17, 2025
 */

function deleteExpiredCalendarTriggers_() {
  const now = new Date();
  const props = GET_PROP_STORE_();
  const stored = JSON.parse(props.getProperty(CALENDAR_STORE) || "{}");

  const triggers = ScriptApp.getProjectTriggers();
  const updated = {};

  triggers.forEach(trigger => {
    const id = trigger.getUniqueId();
    const scheduledTime = stored[id] ? new Date(stored[id]) : null;

    if (scheduledTime && scheduledTime < now) {
      ScriptApp.deleteTrigger(trigger);
      Logger.log(`Deleted expired calendar trigger: ${id} for ${scheduledTime}`);
    } else if (scheduledTime) {
      updated[id] = stored[id];
    }
  });

  props.setProperty(CALENDAR_STORE, JSON.stringify(updated));
  console.log(`Updated store ${CALENDAR_STORE} with values`, updated);
}


function testRepeatTrigger() {
  // Trigger every 6 hours.
  ScriptApp.newTrigger('myFunction')
      .timeBased()
      .everyHours(6)
      .create();
  // Trigger every Monday at 09:00.
  ScriptApp.newTrigger('myFunction')
      .timeBased()
      .onWeekDay(ScriptApp.WeekDay.MONDAY)
      .atHour(9)
      .create();
}
