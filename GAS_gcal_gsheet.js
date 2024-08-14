/**
 * Google Apps Script for automating the creation of calendar events from a Google Sheets document.
 * This script reads data from the "Publishing Schedule 2024" sheet and creates events in the
 * specified team members' Google Calendars. The events are created at 10:00 AM UK time and
 * include a popup reminder 10 minutes before the event.
 *
 * Key functionalities:
 *  - Events are only created for rows marked as "Dubbed" in the specified column (Column D, index 3).
 *  - The script skips past dates and creates events only for future dates, using the date in
 *    Column B (index 1).
 *  - The script processes rows starting from row 52.
 *  - The script uses script properties to store created event IDs to avoid duplicate entries.
 *  - Additional event details can be included from other columns if available.
 *
 * Important notes:
 *  - The script assumes that all users have access to the calendars and that their email addresses
 *    are correctly specified in the `emailAddresses` array.
 *  - The script does not use domain-wide delegation, so the events are created using the credentials
 *    of the user running the script.
 *  - Ensure the date format in the Google Sheets document is consistent and recognized by JavaScript.
 *  - Users must have the appropriate permissions to access and modify the calendars.
 *
 * To configure:
 *  - Update the `spreadsheetId` and `sheetName` with your own Google Sheets ID and sheet name.
 *  - Modify the `emailAddresses` array with the email addresses of the team members who should
 *    receive the calendar events.
 *
 * This script is intended to streamline calendar management for the team, ensuring everyone is
 * informed of key post-launch checks for dubbed content.
 * 
 * Author: Chad Ramey
 * Date: July 30, 2024
 */

function createCalendarEvents() {
  // IDs of the spreadsheet and the tab
  var spreadsheetId = 'docID'; // Replace with your Google Sheets ID
  var sheetName = 'sheeName';

  // Open the spreadsheet and the specific tab
  var sheet = SpreadsheetApp.openById(spreadsheetId).getSheetByName(sheetName);

  // Get the data range starting from row 52
  var dataRange = sheet.getRange('A52:H');
  var data = dataRange.getValues();

  // Email addresses of the team members
  var emailAddresses = ['email_1@domain.com', 'email_2@omain.com', 'email_3omain.com']; // Add new email here

  // Get today's date
  var currentDate = new Date();
  var today = new Date(currentDate.getFullYear(), currentDate.getMonth(), currentDate.getDate());

  Logger.log('Current Date: ' + today);

  var scriptProperties = PropertiesService.getScriptProperties();

  // Iterate through the rows
  for (var i = 0; i < data.length; i++) {
    if (data[i][3] === 'Dubbed') { // Check if the class is Dubbed (column D, index 3)
      var dateStr = data[i][1]; // Date (column B, index 1)

      // Parse the date string and set the event time to 10:00 AM UK time
      var eventDate = new Date(dateStr);
      eventDate.setHours(10, 0, 0, 0); // 10:00 AM UK time

      Logger.log('Event Date: ' + eventDate);

      // Compare only the date parts
      var eventDateOnly = new Date(eventDate.getFullYear(), eventDate.getMonth(), eventDate.getDate());

      if (eventDateOnly < today) {
        Logger.log('Skipping event for ' + dateStr + ' as it is in the past.');
        continue;
      }

      // Create the event title
      var eventTitle = 'Post-launch checks (dubbing) ' + data[i][7]; // Column H, index 7

      // Additional info
      var additionalInfo = '';
      if (data[i][2]) { // Column C, index 2
        additionalInfo += 'Column C: ' + data[i][2] + '\n';
      }
      if (data[i][4]) { // Column E, index 4
        additionalInfo += 'Column E: ' + data[i][4] + '\n';
      }
      if (data[i][5]) { // Column F, index 5
        additionalInfo += 'Column F: ' + data[i][5] + '\n';
      }

      // Create the event description
      var eventDescription = eventTitle + '\n' + additionalInfo;

      // Create the event in the calendar of each team member
      emailAddresses.forEach(function(email) {
        try {
          var eventId = dateStr + '-' + email; // Unique ID for each event and user
          if (!scriptProperties.getProperty(eventId)) {
            var calendar = CalendarApp.getCalendarById(email);
            if (!calendar) {
              throw new Error('Calendar not found for email: ' + email);
            }
            var event = calendar.createEvent(eventTitle, eventDate, eventDate, {description: eventDescription, timeZone: 'Europe/London'});

            // Add popup reminder
            event.addPopupReminder(10); // Popup reminder 10 minutes before the event

            Logger.log('Event created for email ' + email + ' with title: ' + eventTitle);
            scriptProperties.setProperty(eventId, event.getId());
          } else {
            Logger.log('Event for ' + eventId + ' already exists. Skipping.');
          }
        } catch (e) {
          Logger.log('Error creating event for email ' + email + ': ' + e.message);
        }
      });
    }
  }
}
