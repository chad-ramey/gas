/**
 * Script: Auto Accept Calendar Invites
 * 
 * Description:
 * This Google Apps Script automatically accepts specific calendar invites
 * sent to a Gmail account and adds them to the user's primary Google Calendar.
 * It identifies invite emails from a specified sender, parses the attached 
 * iCalendar (.ics) files, and ensures the events are added to the calendar 
 * with the "accepted" RSVP status.
 * 
 * Usage:
 * 1. Update the script with the sender's email address and your own Gmail account.
 * 2. Deploy the script and authorize the required Google Workspace scopes.
 * 3. Run the script manually or set up a trigger for periodic execution.
 * 4. Monitor logs for processing details and debugging.
 * 
 * Author: Chad Ramey
 * Date: December 3, 2024
 */

function autoAcceptInvites() {
  const calendarId = "primary"; // User's primary calendar
  const testSenderEmail = "chad.test@onepeloton.com"; // Specific sender for testing
  const userEmail = Session.getActiveUser().getEmail(); // Guest's email (target)

  Logger.log("Starting autoAcceptInvites script...");

  // Step 1: Search for emails with event invites from the test sender
  const query = `from:${testSenderEmail} is:unread has:attachment filename:ics`; // Filter for emails from the test sender
  const threads = GmailApp.search(query);

  Logger.log(`Found ${threads.length} email threads matching the query.`);

  threads.forEach(thread => {
    const messages = thread.getMessages();
    Logger.log(`Processing thread with ${messages.length} messages.`);

    messages.forEach(message => {
      if (message.isUnread()) {
        const attachments = message.getAttachments();

        Logger.log(`Message has ${attachments.length} attachments.`);

        attachments.forEach(attachment => {
          if (attachment.getContentType() === 'text/calendar') {
            Logger.log("Found a calendar invite (ICS attachment).");

            const icsContent = attachment.getDataAsString();

            // Parse the iCalendar file to extract event details
            const eventDetails = parseICal(icsContent);
            Logger.log(`Parsed event details: ${JSON.stringify(eventDetails)}`);

            // Step 2: Fetch the event by its unique iCalUID
            const events = Calendar.Events.list(calendarId, {
              iCalUID: eventDetails.iCalUID,
              singleEvents: true,
            }).items;

            Logger.log(`Found ${events.length} events with iCalUID: ${eventDetails.iCalUID}`);

            if (events.length > 0) {
              const event = events[0]; // Get the event with the matching iCalUID
              Logger.log(`Event found: ${JSON.stringify(event)}`);

              // Step 3: Explicitly accept the invitation for the correct participant
              if (event.attendees) {
                const updatedAttendees = event.attendees.map(attendee => {
                  if (attendee.email === userEmail) {
                    Logger.log(`Updating RSVP for attendee: ${attendee.email}`);
                    attendee.responseStatus = "accepted"; // Set RSVP to "accepted"
                  }
                  return attendee;
                });

                Logger.log(`Updated attendees: ${JSON.stringify(updatedAttendees)}`);

                try {
                  Calendar.Events.patch(
                    { attendees: updatedAttendees },
                    calendarId,
                    event.id
                  );
                  Logger.log("RSVP status updated successfully.");
                } catch (error) {
                  Logger.log(`Error updating RSVP status: ${error.message}`);
                }
              } else {
                Logger.log("No attendees found in the event.");
              }
            } else {
              Logger.log("No matching events found with the given iCalUID.");
            }

            // Step 4: Mark the email as read and delete it
            message.markRead(); // Mark the email as read
            message.moveToTrash(); // Move the email to the trash
            Logger.log("Email marked as read and moved to trash.");
          }
        });
      }
    });
  });

  Logger.log("autoAcceptInvites script completed.");
}

// Helper function to parse iCalendar data
function parseICal(icsContent) {
  const parsedData = {}; // Parse the ICS content as needed
  Logger.log("Parsing iCalendar data...");

  // Extract the iCalUID from the ICS file (regex example)
  const uidMatch = icsContent.match(/UID:(.*?)(\n|\r)/);
  if (uidMatch) {
    parsedData.iCalUID = uidMatch[1].trim();
    Logger.log(`Extracted iCalUID: ${parsedData.iCalUID}`);
  }

  // Example parsing logic for other fields
  parsedData.start = '2024-12-04T10:00:00-05:00'; // Example start time
  parsedData.end = '2024-12-04T11:00:00-05:00'; // Example end time
  parsedData.summary = 'Sample Event'; // Example event title
  Logger.log(`Parsed data: ${JSON.stringify(parsedData)}`);
  return parsedData;
}
