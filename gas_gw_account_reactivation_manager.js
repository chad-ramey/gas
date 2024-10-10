/**
 * Script: onFormSubmit - Account Monitoring and Suspension via Google Apps Script
 * 
 * Description:
 * This Google Apps Script automates the process of monitoring and managing user account suspension and archiving in Google Workspace. 
 * It is triggered upon the submission of a Google Form, which records the details necessary for monitoring or suspending a user account 
 * (i.e., the email address of the account to be monitored, the timeframe for monitoring, the archived status, and the submitter's email address). 
 * The script checks if the submitter is authorized, updates the corresponding spreadsheet, and sends email notifications to the submitter 
 * based on the success or failure of the operation.
 *
 * Functions:
 * - `onFormSubmit(e)`: Triggered when the form is submitted. It retrieves the data from the most recent submission, checks if the submitter is authorized, 
 *   and either adds the account to a monitoring list (while tracking its archived status) or suspends the account based on the form input. 
 *   The script updates the spreadsheet with a timestamp and sends an email notification to the submitter.
 * 
 * - `calculateDeactivationDate(timeframe)`: Calculates the deactivation date based on the selected timeframe (e.g., 30 days, 60 days, indefinitely, or N/A).
 * 
 * - `addAccountToMonitor(userEmail, deactivateDate, formSubmitterEmail, wasArchived)`: Adds the specified account to a monitoring list along with its deactivation date 
 *   and tracks whether the account was archived.
 * 
 * - `removeAccountFromMonitor(userEmail)`: Removes the specified account from the monitoring list.
 * 
 * - `getArchivedStatus(userEmail)`: Retrieves the archived status of the account from the "Monitored Accounts" sheet.
 * 
 * - `monitorAccounts()`: Periodically checks the monitored accounts and suspends them if their deactivation date has been reached, notifying the submitter via email.
 * 
 * - `manageUserSuspension(userEmail, suspend, wasArchived)`: Suspends or unsuspends the specified user account. If the account was archived before being monitored, it will be 
 *   re-archived when the monitoring ends.
 * 
 * - `sendUnauthorizedEmail(yourEmail)`: Sends an email notification to users who attempt to use the form without authorization.
 * 
 * - `sendEmailNotification(yourEmail, targetEmail, success)`: Sends an email notification to the submitter, informing them of the success or failure of the monitoring/suspension operation.
 *
 * Usage:
 * 1. **Form Setup:**
 *    - The Google Form should be configured with the following fields:
 *      1. **Timestamp**: Automatically captured by Google Forms.
 *      2. **Email Address**: Automatically collected by Google Forms as the verified email address of the person submitting the form.
 *      3. **Action**: Select between "Add" (for monitoring) or "Remove" (for suspension).
 *      4. **Account Email Address**: The email address of the user account to be monitored or suspended.
 *      5. **Timeframe**: Choose "30 days," "60 days," "Indefinitely," or "N/A."
 *    - The form should be linked to a Google Sheet where submissions are recorded.
 * 
 * 2. **Trigger:**
 *  - Form Submission Trigger:
 *    - Set up the `onFormSubmit` function as an installable trigger for the Google Form submission event. This ensures the script runs each time a new form submission is recorded.
 *  - Time-Based Trigger:
 *    - Set up a time-driven trigger for the monitorAccounts function. This trigger should be configured to run periodically (e.g., daily or hourly) to check the status of monitored accounts 
 *      and suspend them if their deactivation date has been reached, while considering if the account was archived and ensuring re-archiving if necessary.
 * 
 * 3. **Permissions:**
 *    - The script restricts form access to specific email addresses. Unauthorized users will have their submissions marked as "Not Authorized" in the "Completed" column and will receive 
 *      an email notification.
 * 
 * 4. **Script Execution:**
 *    - Upon form submission, the script will:
 *      1. Retrieve the most recent form submission data.
 *      2. Verify the submitter’s authorization.
 *      3. If the user is not authorized, mark the submission as "Not Authorized" and send an email notification.
 *      4. If authorized, the script will either add the account to the monitoring list (tracking whether the account was archived) or suspend the account based on the form input.
 *      5. Update the corresponding row in the linked Google Sheet with a timestamp if the operation is successful.
 *      6. Send an email notification to the submitter, indicating whether the operation was successful or if there was an issue.
 * 
 * Notes:
 * - The "Monitored Accounts" sheet includes the following columns:
 *   - **Column A**: `Monitored Accounts` - Holds the email addresses of the accounts being monitored.
 *   - **Column B**: `Deactivation Date` - Stores the deactivation date or "Indefinitely" for accounts monitored without a set timeframe.
 *   - **Column C**: `Form Submitter Email` - Records the email address of the person who submitted the form.
 *   - **Column D**: `Archived Y/N` - Indicates whether the account was archived when it was added to the monitoring list.
 * - The script includes functionality to remove accounts from the monitoring list after the specified timeframe, suspend them, and re-archive them if they were archived before.
 * - Ensure the necessary permissions and APIs (like the Admin SDK) are enabled for the service account used by the script.
 * - The script logs success or failure messages for the operations, which can be reviewed in the Script Editor's Logs.
 * - In case of an error during the process, the script sends a failure notification with instructions to contact the administrator.
 * 
 * Action Items:
 * - Archived accounts tracked for re-archiving.
 * - Move to Deactivated OU.
 * 
 * Author: Chad Ramey
 * Date: October 10, 2024
 */

function onFormSubmit(e) {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var lastRow = sheet.getLastRow();
  
    var timestamp = e.values[0];
    var yourEmail = e.values[1]; // Email Address
    var targetEmail = e.values[3]; // Target Email Address
    var timeframe = e.values[4]; // Timeframe: 30 days, 60 days, Indefinitely, or N/A
    var allowedEmails = ["chad.ramey@onepeloton.com", "tim.rayburn@onepeloton.com", "nick.henry@onepeloton.com", "fabian.graham@onepeloton.com"]; // List of allowed emails

    Logger.log("Form submitted by: " + yourEmail);
    
    if (!allowedEmails.includes(yourEmail)) {
        sheet.getRange(lastRow, 6).setValue("Not Authorized");
        sendUnauthorizedEmail(yourEmail);
        return;
    }

    try {
        if (e.values[2].toLowerCase() === "add") {
            Logger.log("Adding account to monitoring: " + targetEmail);
            var userStatus = manageUserSuspension(targetEmail, false); // Unsuspend user
            if (timeframe.toLowerCase() !== 'n/a') {
                var deactivateDate = calculateDeactivationDate(timeframe);
                addAccountToMonitor(targetEmail, deactivateDate, yourEmail, userStatus.wasArchived); // Add to monitoring list with email and archived status
                Logger.log("Added account to monitoring list: " + targetEmail);
            }
        } else if (e.values[2].toLowerCase() === "remove") {
            Logger.log("Removing account from monitoring: " + targetEmail);
            var wasArchived = getArchivedStatus(targetEmail); // Retrieve archived status
            manageUserSuspension(targetEmail, true, wasArchived); // Suspend user, rearchive if necessary
            removeAccountFromMonitor(targetEmail); // Remove from monitoring list
        }
        sheet.getRange(lastRow, 6).setValue(new Date().toLocaleString());
        sendEmailNotification(yourEmail, targetEmail, true);
    } catch (error) {
        Logger.log("Error: " + error.message);
        sendEmailNotification(yourEmail, targetEmail, false);
    }
}

function calculateDeactivationDate(timeframe) {
    var date = new Date();
    switch (timeframe.toLowerCase()) {
        case '30 days':
            date.setDate(date.getDate() + 30);
            break;
        case '60 days':
            date.setDate(date.getDate() + 60);
            break;
        case 'indefinitely':
            date = null; // No deactivation
            break;
        case 'n/a':
            return null; // No monitoring
    }
    return date;
}

function addAccountToMonitor(userEmail, deactivateDate, formSubmitterEmail, wasArchived) {
    try {
        var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Monitored Accounts');
        if (!sheet) {
            Logger.log('Monitored Accounts sheet not found.');
            return;
        }
        Logger.log('Appending to Monitored Accounts sheet: ' + userEmail + ', ' + deactivateDate + ', Archived: ' + wasArchived);
        sheet.appendRow([userEmail, deactivateDate ? deactivateDate.toLocaleString() : 'Indefinitely', formSubmitterEmail, wasArchived ? 'Y' : 'N']);
        Logger.log('Successfully added to Monitored Accounts sheet.');
    } catch (error) {
        Logger.log('Error adding to Monitored Accounts sheet: ' + error.message);
    }
}

function removeAccountFromMonitor(userEmail) {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Monitored Accounts');
    var data = sheet.getDataRange().getValues();
    for (var i = 0; i < data.length; i++) {
        if (data[i][0] === userEmail) {
            sheet.deleteRow(i + 1);
            break;
        }
    }
}

function getArchivedStatus(userEmail) {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Monitored Accounts');
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
        if (data[i][0] === userEmail) {
            return data[i][3] === 'Y'; // Archived status column
        }
    }
    return false; // Default to not archived if not found
}

function manageUserSuspension(userEmail, suspend, wasArchived = false) {
    try {
        // Fetch user information
        var user = AdminDirectory.Users.get(userEmail, {
            projection: "full"
        });

        var isArchived = user.archived;
        var isSuspended = user.suspended;

        // If the account is archived and the action is to unsuspend, unarchive first
        if (!suspend && isArchived) {
            Logger.log('User ' + userEmail + ' is archived. Unarchiving before unsuspending.');
            var unarchiveUpdate = {
                archived: false
            };
            AdminDirectory.Users.patch(unarchiveUpdate, userEmail); // Unarchive the user first
            Utilities.sleep(2000); // Pause for 2 seconds to allow the unarchive to complete
        }

        // Now unsuspend the user if necessary
        if (!suspend && isSuspended) {
            Logger.log('User ' + userEmail + ' is suspended. Unsuspending the user.');
            var unsuspendUpdate = {
                suspended: false
            };
            AdminDirectory.Users.patch(unsuspendUpdate, userEmail); // Unsuspend the user
            Logger.log('User ' + userEmail + ' has been unsuspended.');
        }

        // If the action is to suspend the account, suspend and re-archive if necessary
        if (suspend && !isSuspended) {
            Logger.log('Suspending user ' + userEmail);
            var suspendUpdate = {
                suspended: true
            };
            AdminDirectory.Users.patch(suspendUpdate, userEmail); // Suspend the user
            Logger.log('User ' + userEmail + ' has been suspended.');

            // If the account was archived before, re-archive it
            if (wasArchived || isArchived) {
                Logger.log('Re-archiving user ' + userEmail);
                var archiveUpdate = {
                    archived: true
                };
                AdminDirectory.Users.patch(archiveUpdate, userEmail); // Re-archive the user
                Logger.log('User ' + userEmail + ' has been re-archived.');
            }
        }

        // Return the initial archived state so it can be logged
        return {
            wasArchived: isArchived
        };

    } catch (error) {
        Logger.log('Error managing user suspension: ' + error.message);
    }
}

function sendUnauthorizedEmail(yourEmail) {
    var subject = "Unauthorized Access Attempt";
    var body = `Hello,\n\nYou attempted to use a form for which you do not have authorization. Please contact the administrator if you believe this is an error.\n\nThank you,\nChad Ramey`;
    MailApp.sendEmail(yourEmail, subject, body);
}

function sendEmailNotification(yourEmail, targetEmail, success) {
    var subject, body;

    if (success) {
        subject = "Form Submission Successful";
        body = `Hello,\n\nThe form submission for the account ${targetEmail} was successful. The account will be monitored based on the specified timeframe.\n\nThank you,\nChad Ramey`;
    } else {
        subject = "Form Submission Failed";
        body = `Hello,\n\nThe form submission for the account ${targetEmail} encountered an issue. Please contact Chad Ramey for further assistance.\n\nThank you,\nChad Ramey`;
    }

    MailApp.sendEmail(yourEmail, subject, body);
}
