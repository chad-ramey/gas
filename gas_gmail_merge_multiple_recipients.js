/**
 * Google Apps Script - Mail Merge from Google Sheets
 *
 * This script performs a mail merge using data from a Google Sheets spreadsheet
 * and a Gmail draft template. It adds a custom menu to the sheet, allowing users
 * to trigger the email-sending process.
 * 
 * Main Components:
 * - Constants for column names that indicate recipients and email status.
 * - An onOpen() function to add a custom "Mail Merge" menu to the spreadsheet UI.
 * - sendEmails() function to handle the main mail merge logic:
 *    - It fetches the Gmail draft based on the subject line entered by the user.
 *    - Each row is checked for an existing "Email Sent" timestamp.
 *    - If no timestamp is found, an email is sent to the specified recipient(s),
 *      and the timestamp or error message is recorded in the "Email Sent" column.
 * - Helper functions to retrieve the Gmail draft template, fill in placeholders
 *   with row-specific data, and escape data for compatibility.
 *
 * Usage:
 * 1. Add the script to your Google Sheets project.
 * 2. Customize the "Recipients" and "Email Sent" column names as needed.
 * 3. Prepare a Gmail draft with placeholders (e.g., {{FirstName}}, {{LastName}})
 *    that match column headers in your sheet.
 * 4. Run the script from the "Mail Merge" menu to start sending emails.
 *
 * Constants:
 * - RECIPIENT_COL: The column name where recipient email addresses are stored.
 * - EMAIL_SENT_COL: The column name where the email status/timestamp is recorded.
 */

// Constants and onOpen function
const RECIPIENT_COL = "Recipients";
const EMAIL_SENT_COL = "Email Sent";

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Mail Merge')
    .addItem('Send Emails', 'sendEmails')
    .addToUi();
}

// Main function to send emails
function sendEmails(subjectLine, sheet = SpreadsheetApp.getActiveSheet()) {
  if (!subjectLine) {
    subjectLine = Browser.inputBox("Mail Merge",
      "Type or copy/paste the subject line of the Gmail " +
      "draft message you would like to mail merge with:",
      Browser.Buttons.OK_CANCEL);

    if (subjectLine === "cancel" || subjectLine == "") {
      return;
    }
  }

  const emailTemplate = getGmailTemplateFromDrafts_(subjectLine);
  const dataRange = sheet.getDataRange();
  const data = dataRange.getDisplayValues();
  const heads = data.shift();
  const emailSentColIdx = heads.indexOf(EMAIL_SENT_COL);
  const obj = data.map(r => (heads.reduce((o, k, i) => (o[k] = r[i] || '', o), {})));
  const out = [];

  obj.forEach(function (row, rowIdx) {
    if (row[EMAIL_SENT_COL] == '') {
      const recipientString = row[RECIPIENT_COL].replace(/\s+/g, ', ');

      try {
        const msgObj = fillInTemplateFromObject_(emailTemplate.message, row);
        GmailApp.sendEmail(recipientString, msgObj.subject, msgObj.text, {
          htmlBody: msgObj.html,
          attachments: emailTemplate.attachments,
          inlineImages: emailTemplate.inlineImages
        });
        out.push([new Date()]);
      } catch (e) {
        out.push([e.message]);
      }
    } else {
      out.push([row[EMAIL_SENT_COL]]);
    }
  });

  sheet.getRange(2, emailSentColIdx + 1, out.length).setValues(out);
}

// Helper functions
function getGmailTemplateFromDrafts_(subject_line) {
  try {
    const drafts = GmailApp.getDrafts();
    const draft = drafts.filter(subjectFilter_(subject_line))[0];
    const msg = draft.getMessage();
    const allInlineImages = draft.getMessage().getAttachments({includeInlineImages: true,includeAttachments:false});
    const attachments = draft.getMessage().getAttachments({includeInlineImages: false});
    const htmlBody = msg.getBody(); 
    const img_obj = allInlineImages.reduce((obj, i) => (obj[i.getName()] = i, obj) ,{});
    const imgexp = RegExp('<img.*?src="cid:(.*?)".*?alt="(.*?)"[^\>]+>', 'g');
    const matches = [...htmlBody.matchAll(imgexp)];
    const inlineImagesObj = {};
    matches.forEach(match => inlineImagesObj[match[1]] = img_obj[match[2]]);

    return {message: {subject: subject_line, text: msg.getPlainBody(), html:htmlBody}, 
            attachments: attachments, inlineImages: inlineImagesObj };
  } catch(e) {
    throw new Error("Oops - can't find Gmail draft");
  }

  function subjectFilter_(subject_line){
    return function(element) {
      if (element.getMessage().getSubject() === subject_line) {
        return element;
      }
    }
  }
}

function fillInTemplateFromObject_(template, data) {
  let template_string = JSON.stringify(template);
  template_string = template_string.replace(/{{[^{}]+}}/g, key => {
    return escapeData_(data[key.replace(/[{}]+/g, "")] || "");
  });
  return JSON.parse(template_string);
}

function escapeData_(str) {
  return str
    .replace(/[\\]/g, '\\\\')
    .replace(/[\"]/g, '\\\"')
    .replace(/[\/]/g, '\\/')
    .replace(/[\b]/g, '\\b')
    .replace(/[\f]/g, '\\f')
    .replace(/[\n]/g, '\\n')
    .replace(/[\r]/g, '\\r')
    .replace(/[\t]/g, '\\t');
}
