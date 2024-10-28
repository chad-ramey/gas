/**
 * To learn how to use this script, refer to the documentation:
 * https://developers.google.com/apps-script/samples/automations/vacation-calendar
 * 
 * Script Name: User Calendar Event Copier
 * Description: This script copies specific events from individual user calendars to a shared team calendar 
 *              based on defined keywords. The events are identified by keywords such as "PTO" or "Out of office" 
 *              and are imported into the specified team calendar.
 *              
 * Key Features:
 * 1. Searches user calendars for events containing specific keywords in the event title.
 * 2. Filters events by a specified date range and checks if the event occurs on a weekday.
 * 3. Imports the filtered events to the team calendar, appending the user's name to the event summary.
 * 
 * Usage Instructions:
 * 1. Set the 'TEAM_CALENDAR_ID2' variable to the ID of the team calendar where events should be copied.
 * 2. Define an array of user email addresses in the 'USER_EMAILS2' variable to specify whose calendars 
 *    should be monitored.
 * 3. Set the keywords in 'KEYWORDS2' that will be used to filter relevant events.
 * 4. Deploy the script with appropriate triggers (e.g., hourly) to ensure regular syncing.
 * 
 * Author: Chad Ramey
 * Date: October 28, 2024
 * Updated the script to use individual email addresses instead of a Google Group
 */

// Set the ID of the team calendar to add events to. You can find the calendar's ID on the settings page.
let TEAM_CALENDAR_ID2 = ''; // Calendar

// Set an array of individual email addresses for the script to action on.
let USER_EMAILS2 = [
  '',
  '',
  '',
  ''
];

let KEYWORDS2 = ['pto', 'ooo', 'out of office', 'offline', 'vacation', 'PTO', 'Out of office'];
let MONTHS_IN_ADVANCE2 = 3;

/**
 * Sets up the script to run automatically every hour.
 */
function setup2() {
  let triggers = ScriptApp.getProjectTriggers();
  if (triggers.length > 0) {
    throw new Error('Triggers are already setup.');
  }
  ScriptApp.newTrigger('sync2').timeBased().everyHours(1).create();
  // Runs the first sync2 immediately.
  sync2();
}

/**
 * Looks through the user's public calendars and adds any
 * 'vacation' or 'out of office' events to the team calendar.
 */
function sync2() {
  // Defines the calendar event date range to search.
  let today = new Date();
  let maxDate = new Date();
  maxDate.setMonth(maxDate.getMonth() + MONTHS_IN_ADVANCE2);

  // Determines the time the script was last run.
  let lastRun = PropertiesService.getScriptProperties().getProperty('lastRun');
  lastRun = lastRun ? new Date(lastRun) : null;

  // For each user, find events having one or more of the keywords in the event
  // summary in the specified date range. Imports each of those to the team calendar.
  let count = 0;
  USER_EMAILS2.forEach(function(email) {
    let username = email.split('@')[0];
    KEYWORDS2.forEach(function(keyword) {
      let events = findEvents2(email, keyword, today, maxDate, lastRun);
      events.forEach(function(event) {
        importEvent2(username, event);
        count++;
      }); // End foreach event.
    }); // End foreach keyword.
  }); // End foreach user.

  PropertiesService.getScriptProperties().setProperty('lastRun', today);
  console.log('Imported ' + count + ' events');
}

/**
 * Imports the given event from the user's calendar into the shared team
 * calendar.
 * @param {string} username The team member that is attending the event.
 * @param {Calendar.Event} event The event to import.
 */
function importEvent2(username, event) {
  event.summary = '[' + username + '] ' + event.summary;
  event.organizer = {
    id: TEAM_CALENDAR_ID2,
  };
  event.attendees = [];

  // If the event is not of type 'default', it can't be imported, so it needs
  // to be changed.
  if (event.eventType != 'default') {
    event.eventType = 'default';
    delete event.outOfOfficeProperties;
    delete event.focusTimeProperties;
  }

  console.log('Importing: %s', event.summary);
  try {
    Calendar.Events.import(event, TEAM_CALENDAR_ID2);
  } catch (e) {
    console.error('Error attempting to import event: %s. Skipping.', e.toString());
  }
}

/**
 * In a given user's calendar, looks for occurrences of the given keyword
 * in events within the specified date range and returns any such events
 * found.
 * @param {string} email The email address to retrieve events for.
 * @param {string} keyword The keyword to look for.
 * @param {Date} start The starting date of the range to examine.
 * @param {Date} end The ending date of the range to examine.
 * @param {Date} optSince A date indicating the last time this script was run.
 * @return {Calendar.Event[]} An array of calendar events.
 */
function findEvents2(email, keyword, start, end, optSince) {
  let params = {
    q: keyword,
    timeMin: formatDateAsRFC3339_2(start),
    timeMax: formatDateAsRFC3339_2(end),
    showDeleted: true,
  };
  if (optSince) {
    // This prevents the script from examining events that have not been
    // modified since the specified date (that is, the last time the
    // script was run).
    params.updatedMin = formatDateAsRFC3339_2(optSince);
  }
  let pageToken = null;
  let events = [];
  do {
    params.pageToken = pageToken;
    let response;
    try {
      response = Calendar.Events.list(email, params);
    } catch (e) {
      console.error('Error retrieving events for %s, %s: %s; skipping', email, keyword, e.toString());
      continue;
    }
    events = events.concat(response.items.filter(function(item) {
      return shouldImportEvent2(email, keyword, item);
    }));
    pageToken = response.nextPageToken;
  } while (pageToken);
  return events;
}

/**
 * Determines if the given event should be imported into the shared team
 * calendar and ensures that it falls on a weekday (Monday - Friday).
 * @param {string} email The email address of the user attending the event.
 * @param {string} keyword The keyword being searched for.
 * @param {Calendar.Event} event The event being considered.
 * @return {boolean} True if the event should be imported.
 */
function shouldImportEvent2(email, keyword, event) {
  // Filters out events where the keyword did not appear in the summary
  // (that is, the keyword appeared in a different field, and are thus
  // not likely to be relevant).
  if (event.summary.toLowerCase().indexOf(keyword) < 0) {
    return false;
  }

  // Check if the event falls on a weekday (Monday - Friday).
  let eventStartDate = new Date(event.start.dateTime || event.start.date);
  let dayOfWeek = eventStartDate.getUTCDay(); // 0 = Sunday, 6 = Saturday
  if (dayOfWeek === 0 || dayOfWeek === 6) {
    // Exclude events occurring on Saturday (6) or Sunday (0).
    return false;
  }

  if (!event.organizer || event.organizer.email === email) {
    // If the user is the creator of the event, always import it.
    return true;
  }

  // Only imports events the user has accepted.
  if (!event.attendees) return false;
  let matching = event.attendees.filter(function (attendee) {
    return attendee.self;
  });
  return matching.length > 0 && matching[0].responseStatus === 'accepted';
}

/**
 * Returns an RFC3339 formatted date String corresponding to the given
 * Date object.
 * @param {Date} date a Date.
 * @return {string} a formatted date string.
 */
function formatDateAsRFC3339_2(date) {
  return Utilities.formatDate(date, 'UTC', 'yyyy-MM-dd\'T\'HH:mm:ssZ');
}
