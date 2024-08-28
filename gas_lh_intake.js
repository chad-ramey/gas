/**
 * Script: saveAttachmentsToDrive - Save DISCO Hold Report Attachments to Google Drive
 * 
 * Description:
 * This Google Apps Script is designed to automate the process of saving specific email attachments 
 * to a designated Google Drive folder. The script searches for emails in your Gmail inbox that match 
 * specific criteria (including labels, subject line, and sender). If matching emails are found and they 
 * contain attachments that follow a specific naming pattern, the attachments are saved to a specified 
 * Google Drive folder.
 *
 * Functions:
 * - `saveAttachmentsToDrive()`: This is the main function that:
 *   1. Searches for emails in Gmail that match the specified criteria.
 *   2. Checks each email for attachments that meet a specific naming pattern.
 *   3. Saves the matching attachments to a designated Google Drive folder.
 *   4. (Optional) Moves the processed emails to trash or archives them.
 * 
 * Usage:
 * 1. **Email Search Criteria:**
 *    - The script searches for emails that match the following criteria:
 *      - **Label:** `_lh` and `inbox`.
 *      - **Subject:** `"DISCO Hold Report"`.
 *      - **Sender:** ``.
 *    - The emails must contain attachments that match the following naming pattern:
 *      - `report_Member on Hold_YYYY-MM-DD_HH-MM-SS.xlsx`
 *    - The naming pattern is defined using a regular expression to ensure that only correctly 
 *      formatted attachments are processed.
 * 
 * 2. **Google Drive Folder:**
 *    - The script saves the matching attachments to a specific Google Drive folder.
 *    - The folder is identified by its ID, which is hardcoded in the script:
 *      - `folder = DriveApp.getFolderById("10uKSe_AHHv3JBVmJjsTPP6IKytiU4G2x");`
 *    - Ensure that the correct folder ID is used and that the service account or user running 
 *      the script has write access to the folder.
 * 
 * 3. **Optional Email Handling:**
 *    - After saving the attachments, the script includes a commented-out line to move the processed 
 *      emails to trash:
 *      - `message.moveToTrash();`
 *    - Uncomment this line if you want the emails to be archived or deleted after processing.
 * 
 * Notes:
 * - **Permissions:** The script requires access to your Gmail account and Google Drive. Make sure that the necessary OAuth permissions are granted.
 * - **Folder ID:** Replace the hardcoded folder ID with the ID of your target Google Drive folder where the attachments should be saved.
 * - **Customization:** You can customize the search query and the naming pattern as needed to match different types of emails or attachments.
 * 
 * Author: Chad Ramey
 * Date: August 28, 2024
 */

function saveAttachmentsToDrive() {
  var threads = GmailApp.search('label:_lh label:inbox subject:"DISCO Hold Report" from:'); // Update
  
  for (var i = 0; i < threads.length; i++) {
    var messages = threads[i].getMessages();
    
    for (var j = 0; j < messages.length; j++) {
      var message = messages[j];
      var attachments = message.getAttachments();
      
      // Check if the email meets the specific criteria
      if (attachments.length > 0) {
        for (var k = 0; k < attachments.length; k++) {
          var attachment = attachments[k];
          var attachmentName = attachment.getName();
          
          // Check if the attachment follows the specific naming pattern
          if (attachmentName.match(/^report_Member on Hold_\d{4}-\d{2}-\d{2}_\d{2}-\d{2}-\d{2}\.xlsx$/)) {
            var folder = DriveApp.getFolderById(""); // Replace with the ID of your Shared Drive folder
            
            // Download the attachment content as bytes
            var attachmentData = attachment.getBytes();
            
            // Create the file with the specified MIME type
            folder.createFile(attachmentName, attachmentData, 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
            
            // Move the email to trash (archive it)
            // message.moveToTrash();
          }
        }
      }
    }
  }
}
