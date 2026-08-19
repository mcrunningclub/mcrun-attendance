/*
Copyright 2025 Andrey Gonzalez (for McGill Students Running Club)

Licensed under the Apache License, Version 2.0 (the "License");
you may not use this file except in compliance with the License.
You may obtain a copy of the License at

    https://www.apache.org/licenses/LICENSE-2.0

Unless required by applicable law or agreed to in writing, software
distributed under the License is distributed on an "AS IS" BASIS,
WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
See the License for the specific language governing permissions and
limitations under the License.
*/
/**
 * Stores an object in the document properties store.
 *
 * @param {string} key - The key under which the object will be stored.
 * @param {Object} obj - The object to store.
 */
function storeProperty_(key, obj) {
  const docProp = PropertiesService.getScriptProperties();
  docProp.setProperty(key, JSON.stringify(obj));
}

let ALL_HEADRUNS = null;

/**
 * Retrieves all headruns from the properties store.
 *
 * @return {Object} - An object containing all headruns.
 */
function getAllHeadruns_() {
  return ALL_HEADRUNS ?? initializeHeadruns();

  function initializeHeadruns() {
    const docProp = PropertiesService.getScriptProperties();
    ALL_HEADRUNS = JSON.parse(docProp.getProperty(HEADRUN_STORE_NAME));
    return ALL_HEADRUNS;
  }
}

/**
 * Retrieves headrunners from the properties store
 * 
 * @return {Object}  An object containing all headrunners
 */
function getAllHeadrunners_() {
  const docProp = PropertiesService.getScriptProperties();
  return JSON.parse(docProp.getProperty(HEADRUNNER_STORE_NAME));
}

/** 
 * Returns day of week (starting on Sunday) given index of the day
 * 
 * @param {number} index  Index of weekday to get, eg. 1
 * @return {string}  Name of weekday, eg. "Monday"
 */
function getWeekdayAsString_(index) {
  const days = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];
  return days[index];
}

/** 
 * Return headrun schedule for given day of week.
 * 
 * @param {string|number} weekday  Day of week to get schedule for. Can be string representation or js equivalent (1 = 'monday').
 * @return {Object}  JSON of run schedule for the given weekday. null if getAllHeadruns_ doesn't return anything
 * 
 * @author [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>)
 * @date  May 2, 2025
 * @update  Sep 28, 2025
 */
function getScheduleFromStore_(weekday) {
  const runSchedule = getAllHeadruns_();
  if (!runSchedule) return null;

  const isString = typeof weekday === 'string';
  const weekdayStr = (isString ? weekday : getWeekdayAsString_(weekday)).toLowerCase();

  // Verify valid run schedule for current weekday, else throw error
  if (runSchedule[weekdayStr]) {
    return runSchedule[weekdayStr];
  }
  throw new Error(`No run schedule found for ${weekday}`);
}


/**
 * Finds timekey in runSchedule within [submissionDate - offsetHours, submissionDate + offsetHours].
 *
 * @param {Date} submissionDate  Date object of submission time.
 * @param {Object} runSchedule  Run schedule to search.
 * @param {integer} [offsetHours=2]  Offset time to search for submission.
 *                                   Defaults to 2 hours.
 * 
 * @return {string}  Matched time key. e.g. `'6pm'`
 *
 * @author [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>)
 * @date  May 4, 2025
 * @update  Jun 1, 2025
 */
function getMatchedTimeKey_(submissionDate, runSchedule, offsetHours = 2) {
  if (!runSchedule) return null;

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
    
    // Debug messages
    //Logger.log(`lowerBound: ${new Date(lowerBound)}`);
    //Logger.log(`upperBound: ${new Date(upperBound)}`);

    return (submissionDate >= lowerBound && submissionDate <= upperBound);
  });

  // Check timekey found, else return error
  if (timeKey) {
    return timeKey;
  }
  throw new Error(`No timekey found for ${submissionDate} with run schedule\n\n${JSON.stringify(runSchedule)}\n\n`);

  /** Helper function to get timestamp in unix */
  function convertToUnix(hour12h, minutes = 0, meridian) {
    if (meridian === 'pm' && hour12h !== 12) hour12h += 12;
    if (meridian === 'am' && hour12h === 12) hour12h = 0;
    return submissionDate.setHours(hour12h, minutes, 0, 0);
  }
}


/** 
 * Returns email address of headrunners for a run, divided by levels.
 * 
 * Replaced initial `getHeadRunnerEmail()`, which was hard-coded and required updating.
 * 
 * @param {string[]} runLevels  Headrun levels to get headrunner emails for.
 * @return {Object}  Given run levels as keys and list of emails as values.
 * 
 * @author [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>)
 * @date  May 2, 2025
 * @update  May 5, 2025
 * 
 * ```js
 * const runs = getScheduleFromStore_('monday');
 * const emails = getHeadrunnerEmailsFromSchedule_(runs['8am']);
 * Logger.log(emails)   // { beginner : ['bob@mail.com'], advanced : ['jane@mail.com'] };
 * ```
 */
function getHeadrunnerEmailsFromSchedule_(runLevels) {
  const headrunnerStore = getAllHeadrunners_();
  const allEmails = {};

  for (const level in runLevels) {
    const levelHeadrunners = runLevels[level];

    const levelEmails = levelHeadrunners.reduce((arr, headrunner) => {
      const email = headrunnerStore[headrunner].email ?? '';
      arr.push(email);
      return arr;
    }, []);

    allEmails[level] = levelEmails;
  }
  return allEmails;
}


/** 
 * Iterates array of headrunner names and returns array of email address if found.
 * Names are formatted as `firstName [middleName] initialLastName.`
 * 
 * @param {string[]} names  Names of headrunners to get emails for.
 * @return {string[]}  List of emails found. If no emails found, return empty list.
 * 
 * @author [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>)
 * @date  Sep 27, 2025
 * @update  Sep 28, 2025
 * 
 * ```js
 * const headrunners = ['Bob B.', 'Jane D.', 'Bart S.'];
 * const emails = getHeadrunnerEmailFromName_(headrunners);
 * Logger.log(emails)   // ['bob@mail.com', 'bart@mail.com'] };
 * ```
 */
function getHeadrunnerEmailsFromNames_(names) {
  // Get all headrunner info (e.g. nameKey, email, strava, ...)
  if (!names) return;

  try {
    const headrunnerStore = getAllHeadrunners_();

    // Reduce list of headrunners for their emails
    return names.reduce((acc, nameKey) => {
      const email = headrunnerStore[nameKey]?.email || "";
      if (email) acc.push(email);
      return acc;
    }, []);
  }
  catch(e) {
    logAsAC_(`Unable to get headrunner email for names '${names}'`, getHeadrunnerEmailsFromNames_.name);
    logAsAC_(`Catch error: ${e.message}`, getHeadrunnerEmailsFromNames_.name);
  }

  return [];
}


/**
 * Adds emails to string with headrunner names.
 * 
 * @param {string} names  Headrunner names delimited by newlines
 * @param {string} delimiter  (Optional) Delimiter used to separate headrunners, default \n 
 * @return {string}  Headrunner info as `name:email` delimited by newlines.
 * 
 * @author [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>)
 * @date  Sep 27, 2025
 * @update  Sep 27, 2025
 * 
 * ```js
 * const headrunners = "Bob B.\nJane D.\nBart S.";
 * const nameEmails = appendHeadrunnerEmail_(headrunners);
 * Logger.log(nameEmails)   // "Bob B.:bob@mail.com\nJane D.\nBart S.:bart@mail.com";
 * ```
 */
function appendHeadrunnerEmail_(namesStr, delimiter = '\n') {
  // Get all headrunner info (e.g. nameKey, email, strava, ...)
  if (!namesStr) return;
  const headrunnerStore = getAllHeadrunners_();

  // If store cannot be read (i.e. cannot access Script Properties, then return namesString)
  // `getAllHeadrunners_()` will not work from external scripts
  if (!headrunnerStore) return namesStr;

  // Split names into array
  const names = namesStr.split(delimiter);

  // Append email to name if found in store
  const namesAndEmails = [];
  for(let i = 0; i < names.length; i++) {
    const email = headrunnerStore[names[i]]?.email || "";
    const nameEmail = names[i] + (email ? `:${email}` : '');
    namesAndEmails.push(nameEmail);
  }
  return namesAndEmails.join(delimiter);
}


/** 
 * Display all headrun and headrunner data in user-friendly log
 */
function prettyPrintRunData() {
  prettyPrintHeadrunnerObj();
  prettyPrintHeadrunObj();

  /** Headrunner printer  */
  function prettyPrintHeadrunnerObj(headrunnerObj = getAllHeadrunners_()) {
    const output = Object.entries(headrunnerObj).reduce((acc, [name, prop]) => {
      acc.push(`${name}:\n  -email: '${prop.email}'\n  -strava: '${prop.stravaId ?? ''}'`);
      return acc;
    }, []);

    console.log(output.join('\n'));
  }

  /** Headrun printer  */
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
 * @WARNING  Only execute when information needs updating.
 * 
 * @author [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>) + ChatGPT
 * @date  May 2, 2025
 * @update  May 3, 2025
 * 
 * ### Sample data structures
 * ```js
 * { sunday : {
 *    '10:00am' : { 'easy' : ['Bob B.', 'Jane D.'] },
 *    '2:15pm': { 'intermediate' : ['Jane D.'] }
 * }}
 * { 'Bob B.' : { email : 'bob.burger@mail.com', strava : '123456789'} }
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
    const nameKey = formatHeadrunnerName_(row[colIndex.name]);
    const email = row[colIndex.email];
    const stravaId = extractStravaId(row[colIndex.strava]);
    const levelStr = row[colIndex.level];

    // Append this runner's schedule info to the headrun object
    appendHeadrunInfo_(levelStr, nameKey, headrunObj);
    
    // Store runner info
    acc[nameKey] = {'email' : email, 'stravaId' : stravaId };
    return acc;
  }, {});

  // Save information to properties store
  storeProperty_(HEADRUNNER_STORE_NAME, headrunnerObj);
  storeProperty_(HEADRUN_STORE_NAME, headrunObj);

  Logger.log(`Completed parsing and storage of run data for '${SEMESTER_NAME}'`);
  prettyPrintRunData();

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

  function formatNameKey_(name) {
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