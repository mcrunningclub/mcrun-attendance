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

const CALENDAR_STORE = SCRIPT_PROPERTY.calendarTriggers;
const TRIGGER_FUNC = checkAttendanceSubmission.name;
const TRIGGER_BASE_ID = 'attendanceTrigger';
const TRIGGER_OFFSET = 60 * 60 * 1000;  // 1 hour in ms

/**
 * Adds new events as time-based triggers and removed expired ones
 * 
 * @trigger  Every Sunday at 1am.
 */

function updateWeeklyCalendarTriggers() {
  // Error Management: ensure correct calendar is used
  if (getCurrentUserEmail_() != CLUB_EMAIL) throw Error('Please change to McRUN account');

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

  PropertiesService.getScriptProperties().setProperty('testEvent', events[0]);
}




function cTest() {
  const calendar = CalendarApp.getDefaultCalendar();
  const day = new Date('2025-04-22 1:00:00');
  const day2 = new Date('2025-04-23 23:00:00');

  const events = calendar.getEvents(day, day2);
  events.forEach(e => {
    console.log(
      e.getDescription(),
      e.getTag('headrunner'),
      e.getTitle()
    )
  })
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

/**
 * Gets the start of the day for a given date.
 *
 * @param {Date} date - The date for which to get the start of the day.
 * @return {Date} - A new Date object set to the start of the given day.
 */

function getStartOfDay_(date) {
  const start = new Date(date);
  start.setHours(0, 0, 0, 0);
  return start;
}

function getEndOfDay_(date) {
  const start = new Date(date);
  start.setHours(23, 59, 59, 59);
  return start;
} 


function updateCalendarTriggers() {
  // Get events from 12am to 11:59pm
  const now = new Date();
  const start = getStartOfDay_(now);
  const end = getEndOfDay_(now);

  const calendar = CalendarApp.getDefaultCalendar();
  const events = calendar.getEvents(start, end);

  const offset = now - 10*60 * 1000;    // Search 6 sec ago

  for (const event of events) {
    if (offset < event.getLastUpdated() && isCancelled(event)) {
      const triggerId = event.getTag('id');
      cleanUpTrigger(triggerId);
      console.log(`This event has been cancelled: ${isCancelled(event)}`);
    }
  }

  function isCancelled(event) {
    const str = event.getDescription() + event.getTitle();
    const cancelledRegex = /cancel/i;
    return cancelledRegex.test(str);
  }
}



/**
 * Add time-based trigger using event information from Calendar.
 * 
 * @param {CalendarEvent} event  Scheduled event as trigger target
 *
 * @author [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>) + ChatGPT
 * @date  Apr 15, 2025
 * @update  Jun 2, 2025
 */

function createAndStoreTrigger_(event) {
  const scriptProperties = PropertiesService.getScriptProperties();
  const startTime = new Date(event.getStartTime().getTime() + TRIGGER_OFFSET);

  // Only add trigger if new
  //if (isExistingTrigger_(startTime)) return;

  const trigger = ScriptApp.newTrigger(TRIGGER_FUNC)
    .timeBased()
    .at(startTime)
    .create();

  // Store trigger details using 'memberName' as key
  const triggerData = JSON.stringify({
    triggerId: trigger.getUniqueId(),
    timedate : event.getStartTime(),
    title : event.getTitle(),
    description : event.getDescription()
  });

  // Label trigger key with member name, and log trigger data
  const key = TRIGGER_BASE_ID + (trigger.getUniqueId());
  
  scriptProperties.setProperty(key, triggerData);
  Logger.log(`Created new trigger '${key}':\n\n${triggerData}`);

  // Helper function
  function isExistingTrigger_(time) {
    const triggerTimes = Object.values(stored);
    return (time in triggerTimes);
  }
}



function runSubmissionChecker() {
  const scriptProperties = PropertiesService.getScriptProperties();
  const allProps = scriptProperties.getProperties();

  for (let key in allProps) {
    if (!key.startsWith(TRIGGER_BASE_ID)) continue;
    const triggerData = JSON.parse(allProps[key]);
    const { timedate, triggerId, title, description } = triggerData;

    const allLevels = Object.keys(ATTENDEE_MAP).join('|');
    const level = title.match(new RegExp(allLevels, 'i'))[0];

    // Verify if trigger time is in the future
    const today = new Date();
    if (new Date(timedate) > today) continue;

    const isSubmitted = checkAttendanceSubmission(timedate, level);
    if (isSubmitted) {
      cleanUpTrigger(key, triggerId);
      Logger.log(`Cleaning up trigger ${key}\n\n${triggerData}`);
    }
    else {
      const emailObj = { emailsByLevel, title };

      // Send reminder of email
      sendEmailReminder_(emailObj);
    }
  }
}


function checkThisEvent(timedate) {
  const currentWeekday = timedate.getDay();
  const currentDaySchedule = getScheduleFromStore_(currentWeekday);
  const currentTimeKey = getMatchedTimeKey_(timedate, currentDaySchedule);

  // Get emails using run schedule for current day, then proceed to actual verification
  const runScheduleLevels = currentDaySchedule[currentTimeKey];
  const { 'timeKey' : matchedTimeKey, 'submission' : submission } = verifyAttendance_(currentWeekday);
}


/** Helper: remove trigger and data in script properties */
function cleanUpTrigger(key, triggerId) {
  deleteTriggerById(triggerId);
  PropertiesService.getScriptProperties().deleteProperty(key);
}

/**
 * Deletes a trigger by its unique ID.
 *
 * This function iterates through all project triggers to find and delete the one
 * with the specified unique ID. If the trigger is not found, it throws an error.
 *
 * @param {string} triggerId - The unique ID of the trigger to delete.
 */
function deleteTriggerById(triggerId) {
  const triggers = ScriptApp.getProjectTriggers();

  for (let trigger of triggers) {
    if (trigger.getUniqueId() === triggerId) {
      ScriptApp.deleteTrigger(trigger);
      Logger.log(`Trigger with id ${triggerId} deleted!`);
      return;
    }
  }
  // If we reach here, the trigger was not found
  throw new Error(`⚠️ Trigger with id ${triggerId} not found`);
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
