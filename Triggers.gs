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
 * Adds new events as time-based triggers and removed expired ones
 * 
 * @trigger  Every Sunday at 1am.
 */

function updateWeeklyCalendarTriggers() {
  // Error Management: ensure correct calendar is used
  if (getCurrentUserEmail_() != CLUB_EMAIL) throw Error('Please change to McRUN account');

  createCalendarTriggersForWeek_();
  deleteExpiredCalendarTriggers_();
}

/**
 * Get events for current week from calendar and create time-based triggers.
 *
 * @author [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>) + ChatGPT
 * @date  Apr 17, 2025
 * @update  Apr 27, 2025
 */

function createCalendarTriggersForWeek_() {
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

  filteredEvents.forEach(event => createCalendarTrigger_(event));

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
 * Get cancelled events for today from calendar and remove their triggers.
 */
function cleanUpCalendarTriggersForToday() {
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
      deleteTrigger_(triggerId, null);
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
 * Add new McRUN event(s) from calendar to Apps Script trigger for today.
 * 
 * @trigger  Updated calendar.
 *
 * @author [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>)
 * @date  Apr 17, 2025
 * @update  Apr 17, 2025
 */

function createCalendarTriggersForToday() {
  const now = new Date();
  const midnight = new Date(new Date().setHours(23, 59, 59, 59));

  const calendar = CalendarApp.getDefaultCalendar();
  const events = calendar.getEvents(now, midnight);
  events.forEach(e => createCalendarTrigger_(e));

  PropertiesService.getScriptProperties().setProperty('testEvent', events[0]);
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

function createCalendarTrigger_(event) {
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

/**
 * Check if attendance has been submitted and send reminder email if not.
 */
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

    const isSubmitted = checkMissingAttendance(timedate, level);
    if (isSubmitted) {
      deleteTrigger_(triggerId, key);
      Logger.log(`Cleaning up trigger ${key}\n\n${triggerData}`);
    }
    else {
      const emailObj = { emailsByLevel, title };

      // Send reminder of email
      sendAttendanceReminder_(emailObj);
    }
  }
}

/**
 * Deletes a trigger by its unique ID and removes its data from script properties if needed.
 *
 * This function iterates through all project triggers to find and delete the one
 * with the specified unique ID. If the trigger is not found, it throws an error.
 *
 * @param {string} id - The unique ID of the trigger to delete.
 * @param {string} key - (Optional) The key of trigger's associated script property.
 */
function deleteTrigger_(id, key = null) {
  const triggers = ScriptApp.getProjectTriggers();

  for (let trigger of triggers) {
    if (trigger.getUniqueId() === id) {
      ScriptApp.deleteTrigger(trigger);
      if (key) {
        PropertiesService.getScriptProperties().deleteProperty(key);
      }
      Logger.log(`Trigger with id ${id} deleted!`);
      return;
    }
  }
  // If we reach here, the trigger was not found
  throw new Error(`⚠️ Trigger with id ${id} not found`);
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