// Emails of current execs
const PRESIDENT_EMAIL = 'alexis.demetriou@mail.mcgill.ca';
const VP_INTERNAL_EMAIL = 'emmanuelle.blais@mail.mcgill.ca';

const CALENDAR_STORE = SCRIPT_PROPERTY.calendarTriggers;
const HEADRUNNER_STORE_NAME = 'headrunners';
const HEADRUN_STORE_NAME = 'headruns';

const TRIGGER_FUNC = checkMissingAttendance.name;

function storeObject_(key, obj) {
  const docProp = PropertiesService.getDocumentProperties();
  docProp.setProperty(key, JSON.stringify(obj));
}

let ALL_HEADRUNS = null;

function getAllHeadruns_() {
  if (!ALL_HEADRUNS){  
    const docProp = PropertiesService.getDocumentProperties();
    ALL_HEADRUNS = JSON.parse(docProp.getProperty(HEADRUN_STORE_NAME));
  }
  return ALL_HEADRUNS;
}

function getAllHeadrunners_() {
  const docProp = PropertiesService.getDocumentProperties();
  return JSON.parse(docProp.getProperty(HEADRUNNER_STORE_NAME));
}


/** Returns day code formatted as `weekday` in lowercase. Index [0-6] (Sunday = 0) */
function getWeekday_(weekdayIndex) {
  const days = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];
  return days[weekdayIndex].toLowerCase();
}


/** 
 * Find schedule for current weekday, either as string representation, 
 * or js integer equivalent (1 = 'monday').
 * 
 * @author [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>)
 * @date  May 2, 2025
 * @update  May 5, 2025
 */

function getScheduleFromStore_(currentWeekday) {
  const runSchedule = getAllHeadruns_();
  const isString = typeof currentWeekday === 'string';
  const weekString = isString ? currentWeekday.toLowerCase() : getWeekday_(currentWeekday);
  return runSchedule[weekString] ?? null;   // Run schedule for current weekday
}


/**
 * Return headrun day and time from headrun code input `headRunDay`.
 *
 * @param {string} headRunDay  The headrun code representing specific headrun (e.g., `'SundayPM'`).
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

    default: throw new Error(`No headrunner has been found for '${headRunDay}'`);
  }
}


function getMatchedTimeKey(submissionDate, runSchedule, offsetHours = 2) {
  const timeKey = Object.keys(runSchedule).find((timeStr) => {
    const timeMatch = timeStr.match(/(\d+(?::\d+)?)(am|pm)/i);

    // Parse time string to hours and minutes
    const [, timePart, meridian] = timeMatch;
    const [hourStr, minuteStr = '0'] = timePart.split(':');
    const hours = parseInt(hourStr, 10);
    const minutes = parseInt(minuteStr, 10);

    // Convert to number for easy comparaison
    const unixTimestamp = convertToUnix(hours, minutes, meridian.toLowerCase());
    const offsetMilli = offsetHours * 60 * 60 * 1000;

    const lowerBound = unixTimestamp - offsetMilli;
    const upperBound = unixTimestamp + offsetMilli;

    Logger.log(`lowerBound: ${new Date(lowerBound)}`);
    Logger.log(`upperBound: ${new Date(upperBound)}`);

    return (submissionDate >= lowerBound && submissionDate <= upperBound);
  });

  return timeKey;


  /** Helper function to get timestamp in unix */
  function convertToUnix(time12h, minutes = 0, meridian) {
    let hours = time12h;

    if (meridian === 'pm' && hours !== 12) hours += 12;
    if (meridian === 'am' && hours === 12) hours = 0;
    return new Date(1715440100000 + 31536000000).setHours(hours, minutes, 0, 0);
    //return new Date().setHours(hours, minutes, 0, 0);
  }
}

/** 
 * Returns emails of headrunners for a run, divided by levels.
 * 
 * Replaced initial `getHeadRunnerEmail()`, which was hard-coded and required updating.
 * 
 * @author [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>)
 * @date  May 2, 2025
 * @update  May 5, 2025
 * 
 * ```js
 * const run = getScheduleFromStore_('monday')['8am'];
 * const emails = getHeadrunnerEmailFromStore_(run);
 * Logger.log(emails)   // { beginner : ['bob@mail.com], advanced : ['jane@mail.com'] };
 * ```
 * 
*/

function getHeadrunnerEmailFromStore_(runScheduleLevels) {
  const headrunnerStore = getAllHeadrunners_();
  const allEmails = {};

  for (const level in runScheduleLevels) {
    const levelHeadrunners = runScheduleLevels[level];

    const levelEmails = levelHeadrunners.reduce((arr, headrunner) => {
      const email = headrunnerStore[headrunner].email ?? '';
      arr.push(email);
      return arr;
    }, []);

    allEmails[level] = levelEmails;
  }

  return allEmails;
}


/** Display all headrun and headrunner data */

function prettyPrintRunData() {
  prettyPrintHeadrunnerObj();
  prettyPrintHeadrunObj();

  function prettyPrintHeadrunnerObj(headrunnerObj = getAllHeadrunners_()) {
    const output = Object.entries(headrunnerObj).reduce((acc, [name, prop]) => {
      acc.push(`${name}:\n  -email: '${prop.email}'\n  -strava: '${prop.stravaId ?? ''}'`);
      return acc;
    }, []);

    console.log(output.join('\n'));
  }

  function prettyPrintHeadrunObj(headrunObj = getAllHeadruns_()) {
    let output = '';
    for (const day in headrunObj) {
      output += day + ':\n';

      const times = headrunObj[day];
      for (const time in times) {
        output += `  '${time}':\n`;

        var levels = times[time];
        for (const level in levels) {
          const people = levels[level];
          output += `    -${level}: [ '${people.join(', ')}' ]\n`;
        }
      }
    }
    console.log(output);
  }
}


/**
 * Parses headrunner information in Headrunner sheet and stores it in Properties store.
 * 
 * Only execute when information needs updating.
 * 
 * @author [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>) + ChatGPT
 * @date  May 2, 2025
 * @update  May 3, 2025
 * 
 * ### Sample data structures
 * ```js
 * { sunday : {
 *    '10:00am' : { 'easy' : ['bobBurger', 'janeDoe'] },
 *    '2:15pm': { 'intermediate' : ['janeDoe'] }
 * }}
 * 
 * { 'bobBurger' : { email : 'bob.burger@mail.com', strava : '123456789'} }
 * ```
 */

function readAndStoreRunData() {
  const sheet = GET_HEADRUNNER_SHEET_();
  const data = sheet.getDataRange().getValues();

  // Pop header from data, and map column names to indices
  const header = data.shift();
  const colIndex = getColIndices(['name', 'email', 'strava', 'level'], header);

  // Package both objects at the same time
  const headrunObj = {};
  const headrunnerObj = data.reduce((acc, row) => {
    
    // Keys must be identical in both objects (headruns + headrunners)
    const nameKey = formatNameKey(row[colIndex.name]);
    const email = row[colIndex.email];
    const stravaId = extractStravaId(row[colIndex.strava]);
    const levelStr = row[colIndex.level];

    // Append this runner’s schedule info to the headrun object
    appendHeadrunInfo_(levelStr, nameKey, headrunObj);
    
    // Store runner info
    acc[nameKey] = {'email' : email, 'stravaId' : stravaId };
    return acc;
  }, {});

  // Save information to properties store
  storeObject_(HEADRUNNER_STORE_NAME, headrunnerObj);
  storeObject_(HEADRUN_STORE_NAME, headrunObj);

  Logger.log(`Completed parsing and storage of run data from '${SEMESTER_NAME}' sheet`);

  /** Helper functions to extract data */
   function getColIndices(targets, headerRow) {
    const indices = {};

    targets.forEach(key => {
      const index = headerRow.findIndex(h => h.toLowerCase() === key);
      if (index === -1) {
        throw new Error(`Column '${key}' not found in header row.`);
      }
      indices[key] = index;
    });
    return indices;
  }

  function formatNameKey(name) {
    if (!name) return '';
    return name.replace(/[ \-]/g, '').replace(/^./, c => c.toLowerCase());
  }

  function extractStravaId(input) {
    const match = input ? input.match(/(\d+)$/) : null;
    return match ? match[1] : input;
  }
}


/**
 * Appends a headrunner to a nested schedule object.
 * 
 * Helper function for `readAndStoreRunData`.
 * 
 * @param {string} levelsStr  Headrun schedule string delimited by `;`.
 * @param {string} thisHeadrunner  The name of the headrunner to add.
 * @param {Object} headrunObj  Stores all headrun information (day, time, level, headrunners).
 * 
 * @author [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>) + ChatGPT
 * @date  May 2, 2025
 * @update  May 2, 2025
 * 
 * ```javascript
 * // Sample Script ➜ Stores headrunner schedule with name.
 * var headrunnerSchedule = 'Wednesday 6pm (Beginner); Sunday 8am (Intermediate); Sunday 6pm (Beginner)';
 * appendHeadrunInfo(headrunnerSchedule, 'Bob');   // Appends to `headrunObj`
 * 
 * Logger.log(headrunObj);
 * // { 'wednesday' : { '6pm' : { 'beginner' : ['Bob'] } },
 * //   'sunday' : { '8am' : { 'intermediate' : ['Bob'] }, '6pm' : { 'beginner' : ['Bob'] }  }
 * ```
 */

function appendHeadrunInfo_(levelsStr, thisHeadrunner, headrunObj) {
  const levels = levelsStr.split(/\s*;\s*/);

  levels.forEach(entry => {
    // Entry formatted as: `[weekday] [time-12h] ([run level])`
    const match = entry.match(/^(\w+)\s+([\d:apm]+)\s+\(([\w\s]+)\)$/i);
    
    if (match) {
      const [_, day, time, level] = match;
      
      // Create data structure (if first time)
      const dayObj = ensureKey(headrunObj, day, {});
      const timeObj = ensureKey(dayObj, time, {});
      const levelArr = ensureKey(timeObj, level, []);

      // Push to array of headrunners
      levelArr.push(thisHeadrunner);
    }
    else {
      console.error(`Invalid entry skipped: "${entry}"`);
    }
  });

  function ensureKey(obj, key, defaultValue) {
    key = key.toLowerCase();
    if (!obj[key]) obj[key] = defaultValue;
    return obj[key];
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
