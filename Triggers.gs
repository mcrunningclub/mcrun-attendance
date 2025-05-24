const CALENDAR_STORE = SCRIPT_PROPERTY.calendarTriggers;
const TRIGGER_FUNC = checkMissingAttendance.name;

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


function updateCalendarTriggers() {
  // Get events from day start (midnight)
  const now = new Date();
  const start = getStartOfDay_(now);

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