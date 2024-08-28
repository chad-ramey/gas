/**
 * Script: onFormSubmit - Create Calendar Events from Google Form Submissions
 * 
 * Description:
 * This Google Apps Script automatically creates a calendar event based on the information submitted through a Google Form. 
 * The script is triggered upon form submission and uses the data provided in the form to create an event in a specified Google Calendar. 
 * The event's title is set to the "Summary and System" field, and the event description includes several other form fields, 
 * including the "Requested by" email address from column B, "Change Process," "Change Request Description and Rationale," 
 * "Smoke Testing Plan," and "Rollback Plan." Upon successful creation of the event, a confirmation email is sent to a 
 * designated recipient, which includes detailed information about the event and the email address of the form submitter.
 *
 * Functions:
 * - `onFormSubmit(e)`: Triggered when the form is submitted. This function extracts relevant data from the form submission, 
 *   constructs an event in the specified Google Calendar, and sends a confirmation email upon successful event creation. 
 *   The confirmation email includes the email address of the submitter (from column B) and detailed event information. 
 *   If an error occurs during event creation, an error notification email is sent.
 * 
 * - `sendConfirmationEmail(eventTitle, eventDateTime, emailAddress, changeProcess, summaryAndSystem, changeDescription, smokeTestingPlan, rollbackPlan)`: 
 *   Sends a confirmation email to the designated recipient after the event has been successfully created, including the event's title, 
 *   start time, and detailed event information along with the email address of the submitter.
 * 
 * - `sendErrorNotification(errorMessage)`: Sends an error notification email to the designated recipient if an error occurs 
 *   during the event creation process, detailing the nature of the error.
 *
 * Usage:
 * 1. **Form Setup:**
 *    - The Google Form should be configured with the following fields:
 *      1. **Timestamp**: Automatically captured by Google Forms.
 *      2. **Email Address**: Automatically collected by Google Forms as the verified email address of the person submitting the form.
 *      3. **Change Process**: The type of change process being recorded (e.g., "Full CAB," "Emergency CAB").
 *      4. **Summary and System**: A brief summary of the change and the systems involved. This will be used as the event title.
 *      5. **Start Date**: The start date of the event.
 *      6. **Start Time**: The start time of the event.
 *      7. **End Date**: The end date of the event.
 *      8. **End Time**: The end time of the event.
 *      9. **Change Request Description and Rationale**: A detailed description and rationale for the change.
 *     10. **Smoke Testing Plan**: The plan for testing the change.
 *     11. **Rollback Plan**: The plan for rolling back the change if necessary.
 * 
 * 2. **Trigger:**
 *    - Set up the `onFormSubmit` function as an installable trigger for the Google Form submission event. This ensures the script runs 
 *      each time a new form submission is recorded.
 * 
 * 3. **Calendar Setup:**
 *    - Ensure that you have access to the target Google Calendar where the events will be created. The Calendar ID should be correctly 
 *      specified in the script.
 * 
 * 4. **Script Execution:**
 *    - Upon form submission, the script will:
 *      1. Retrieve the most recent form submission data, including the submitter's email address.
 *      2. Combine the "Start Date" and "Start Time" to create the event's start time, and "End Date" and "End Time" to create the event's end time.
 *      3. Create a new event in the specified Google Calendar with the "Summary and System" field as the title.
 *      4. Populate the event description with the "Requested by" email address, "Change Process," "Change Request Description and Rationale," 
 *         "Smoke Testing Plan," and "Rollback Plan."
 *      5. Send a confirmation email to the designated recipient, including the event's title, start time, the submitter's email address, and 
 *         detailed event information.
 *      6. If an error occurs during event creation, send an error notification email to the designated recipient.
 * 
 * Notes:
 * - Ensure that the form fields are correctly mapped to the script's variables, and the Google Calendar ID is accurately specified.
 * - The script logs all actions, which can be reviewed in the Script Editor's Logs for debugging purposes.
 * - If an error occurs during event creation, the error message is captured and emailed to the designated recipient for review.
 * 
 * Author: Chad Ramey
 * Date: August 28, 2024
 */

// Function to trigger on form submission
function onFormSubmit(e) {
  try {
    // Define the calendar ID
    const calendarId = ''; // Update

    // Get the form responses
    const formResponse = e.values;
    console.log('Form Response:', formResponse);

    // Extract the relevant fields from the form response
    const emailAddress = formResponse[1]; // Email Address from column B
    const summaryAndSystem = formResponse[3]; // "Summary and System" is the 4th element (index 3)
    const startDate = formResponse[4]; // "Start Date" is the 5th element (index 4)
    const startTime = formResponse[5]; // "Start Time" is the 6th element (index 5)
    const endDate = formResponse[6]; // "End Date" is the 7th element (index 6)
    const endTime = formResponse[7]; // "End Time" is the 8th element (index 7)
    const changeDescription = formResponse[8]; // "Change request description and rationale" is the 9th element (index 8)
    const changeProcess = formResponse[2]; // "Change Process" is the 3rd element (index 2)
    const smokeTestingPlan = formResponse[9]; // "Smoke Testing Plan" is the 10th element (index 9)
    const rollbackPlan = formResponse[10]; // "Rollback Plan" is the 11th element (index 10)

    console.log('Event Title:', summaryAndSystem);
    console.log('Start Date:', startDate, 'Start Time:', startTime);
    console.log('End Date:', endDate, 'End Time:', endTime);

    // Combine date and time to create event start and end times
    const startDateTime = new Date(`${startDate} ${startTime}`);
    const endDateTime = new Date(`${endDate} ${endTime}`);

    console.log('Start DateTime:', startDateTime);
    console.log('End DateTime:', endDateTime);

    // Build the event description
    const eventDescription = `Requested by: ${emailAddress}\nChange Process: ${changeProcess}\nSummary and System: ${summaryAndSystem}\n` +
                             `Change Request Description and Rationale: ${changeDescription}\nSmoke Testing Plan: ${smokeTestingPlan}\n` +
                             `Rollback Plan: ${rollbackPlan}`;

    // Explicitly set the event title to "Summary and System"
    const event = CalendarApp.getCalendarById(calendarId).createEvent(summaryAndSystem, startDateTime, endDateTime, {
      description: eventDescription,
    });

    // Explicitly set the event title again to ensure it's correct
    event.setTitle(summaryAndSystem);

    console.log('Event Created:', event);

    // Send confirmation email
    sendConfirmationEmail(summaryAndSystem, startDateTime, emailAddress, changeProcess, summaryAndSystem, changeDescription, smokeTestingPlan, rollbackPlan);

  } catch (error) {
    // Handle errors and notify the appropriate parties
    console.error('Error:', error.message);
    sendErrorNotification(error.message);
  }
}

// Function to send a confirmation email after event creation
function sendConfirmationEmail(eventTitle, eventDateTime, emailAddress, changeProcess, summaryAndSystem, changeDescription, smokeTestingPlan, rollbackPlan) {
  const emailRecipients = ''; // Update
  const subject = `Cab Form Confirmation: Event "${eventTitle}" Created`;
  const body = `The "${eventTitle}" event has been created on ${eventDateTime} in the Content Platform Releases Calendar. Please review the below change control details and reply to the requestor and cbs.cab@onepeloton.com with any queries.\n\n` +
               `Full CAB submissions require two explicit approvals; to approve please reply to the requestor and cbs.cab@onepeloton.com with your approval.\n\n` +
               `Requested by: ${emailAddress}\n\n` +
               `Change Process: ${changeProcess}\n\n` +
               `Summary and System: ${summaryAndSystem}\n\n` +
               `Change Request Description and Rationale: ${changeDescription}\n\n` +
               `Smoke Testing Plan: ${smokeTestingPlan}\n\n` +
               `Rollback Plan: ${rollbackPlan}`;

  console.log('Sending Confirmation Email to:', emailRecipients);
  MailApp.sendEmail(emailRecipients, subject, body);
  console.log('Confirmation Email Sent');
}

// Function to send an error notification if something goes wrong
function sendErrorNotification(errorMessage) {
  const emailRecipient = ''; // Update
  const subject = 'Cab Form Error: Event Creation Failed';
  const body = `An error occurred while attempting to create an event: ${errorMessage}`;

  console.log('Sending Error Notification Email to:', emailRecipient);
  MailApp.sendEmail(emailRecipient, subject, body);
  console.log('Error Notification Email Sent');
}
