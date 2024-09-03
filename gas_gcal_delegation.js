/**
 * Script: onFormSubmit - Calendar Delegation via Service Account
 * 
 * Description:
 * This Google Apps Script automates the process of delegating calendar access from one user to another using a service account. 
 * It is triggered upon the submission of a Google Form, which records the details necessary for the delegation 
 * (i.e., the calendar owner, the authorized user, and the submitter's email address). The script then attempts to add 
 * the authorized user as an owner of the specified calendar. Upon success or failure, it updates the form's corresponding 
 * spreadsheet with a timestamp and sends an email notification to the submitter.
 *
 * Functions:
 * - `onFormSubmit(e)`: Triggered when the form is submitted. This function retrieves the data from the most recent submission, 
 *   checks if the submitter is authorized, and if not, marks the submission as "Not Authorized" and sends an email notification. 
 *   If the submitter is authorized, the function attempts to add the authorized user as a calendar owner, updates the spreadsheet 
 *   with a completion timestamp if successful, and sends an email notification to the form submitter.
 * 
 * - `addDelegateToCalendarUsingServiceAccount(calendarOwnerEmail, authorizedUserEmail)`: This function uses the Calendar API 
 *   (Advanced Service) to add the specified authorized user as an owner of the calendar owned by the specified user. 
 *   It returns `true` if the operation is successful and `false` otherwise.
 * 
 * - `sendEmailNotification(yourEmail, authorizedUserEmail, success)`: Sends an email notification to the form submitter, 
 *   informing them of the success or failure of the calendar delegation process. The email content is customized based on the outcome.
 *
 * Usage:
 * 1. **Form Setup:**
 *    - The Google Form should be configured with the following fields:
 *      1. **Timestamp**: Automatically captured by Google Forms.
 *      2. **Email Address**: Automatically collected by Google Forms as the verified email address of the person submitting the form.
 *      3. **Calendar Owner Email Address**: The email address of the person who owns the calendar.
 *      4. **Authorized Calendar User Email Address**: The email address of the person to whom calendar access is being delegated.
 *    - The form should be linked to a Google Sheet where submissions are recorded.
 * 
 * 2. **Trigger:**
 *    - Set up the `onFormSubmit` function as an installable trigger for the Google Form submission event. This ensures the script runs 
 *      each time a new form submission is recorded.
 * 
 * 3. **Calendar API Setup:**
 *    - Ensure the Calendar API (Advanced Service) is enabled in the Google Apps Script project. The service account must have the necessary 
 *      permissions to manage calendar access for the users specified.
 * 
 * 4. **Script Execution:**
 *    - Upon form submission, the script will:
 *      1. Retrieve the most recent form submission data.
 *      2. Verify that the submitter is authorized to use the form.
 *      3. If the user is not authorized, mark the submission as "Not Authorized" in the "Completed" column and send an email notification.
 *      4. If the user is authorized, attempt to add the authorized user as an owner of the calendar specified by the calendar owner's email.
 *      5. Update the corresponding row in the linked Google Sheet with a timestamp if the operation is successful.
 *      6. Send an email notification to the submitter, indicating whether the operation was successful or if there was an issue.
 * 
 * Notes:
 * - The script restricts form access to only one specific email address (besides the owner). Unauthorized users will have their submissions 
 *   marked as "Not Authorized" in the "Completed" column, and they will receive an email notification informing them that they are not authorized 
 *   to use the form.
 * - Ensure that the service account has the appropriate permissions to manage calendar ACLs (Access Control Lists) for the users involved.
 * - The script logs success or failure messages for the delegation process, which can be reviewed in the Script Editor's Logs.
 * - In case of an error during the delegation process, the script sends a failure notification with instructions to contact the administrator.
 * 
 * Author: Chad Ramey
 * Date: September 3, 2024
 */

function onFormSubmit(e) {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var lastRow = sheet.getLastRow();
  
    // Get data from the last submitted form
    var timestamp = e.values[0]; // Timestamp
    var yourEmail = e.values[1]; // Email Address
    var calendarOwnerEmail = e.values[2]; // Calendar Owner Email Address
    var authorizedUserEmail = e.values[3]; // Authorized Calendar User Email Address

    // Specify the allowed email addresses
    var allowedEmails = ["Email_1", "Email_2"]; // Add more emails as needed

    // Check if the submitter's email is allowed
    if (!allowedEmails.includes(yourEmail)) {
        // If not, mark as "Not Authorized" and exit the function
        sheet.getRange(lastRow, 5).setValue("Not Authorized");
        MailApp.sendEmail(yourEmail, "Unauthorized Access", "You are not authorized to use this form.");
        return;
    }

    try {
        // Use the service account to add the authorized user as a calendar manager (owner)
        var success = addDelegateToCalendarUsingServiceAccount(calendarOwnerEmail, authorizedUserEmail);
    
        // If successful, update the Completed column with date and time
        if (success) {
            var completedTimestamp = new Date();
            var formattedTimestamp = Utilities.formatDate(completedTimestamp, Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
            sheet.getRange(lastRow, 5).setValue(formattedTimestamp);

            // Send success email
            sendEmailNotification(yourEmail, authorizedUserEmail, true);
        } else {
            // Send failure email
            sendEmailNotification(yourEmail, authorizedUserEmail, false);
        }
    } catch (error) {
        Logger.log("Error: " + error.message);
        MailApp.sendEmail("Email_3", "Script Failure Alert: Gcal Delegation", "The script encountered an error: " + error.message);
    }
}
  
function addDelegateToCalendarUsingServiceAccount(calendarOwnerEmail, authorizedUserEmail) {
    try {
        // Add the authorized user as a calendar owner using Calendar API (Advanced Service)
        var resource = {
            scope: {
                type: 'user',
                value: authorizedUserEmail
            },
            role: 'owner'  // Set role to 'owner' for full control
        };
        
        Calendar.Acl.insert(resource, calendarOwnerEmail);
  
        Logger.log('Authorized user ' + authorizedUserEmail + ' was successfully added as an owner to ' + calendarOwnerEmail + '\'s calendar.');
        return true;
    } catch (error) {
        Logger.log('Error adding authorized user as an owner to calendar: ' + error.message);
        return false;
    }
}

function sendEmailNotification(yourEmail, authorizedUserEmail, success) {
    var subject, body;

    if (success) {
        subject = "Calendar Delegation Successful";
        body = "Hello,\n\nThe calendar delegation process has been successfully completed. " +
               authorizedUserEmail + " has been granted access to the calendar owned by " +
               "the specified user.\n\nPlease note that " + authorizedUserEmail + " will receive an email to add " +
               "the calendar to their list. They need to click 'Add' to complete the process.\n\n" +
               "If you have any questions, please let me know.\n\nBest regards,\nName";
    } else {
        subject = "Calendar Delegation Failed";
        body = "Hello,\n\nThe calendar delegation process encountered an issue and was not successful. " +
               "Please contact Name at Email_3 for further assistance.\n\n" +
               "Thank you,\nName";
    }

    MailApp.sendEmail(yourEmail, subject, body);
}
