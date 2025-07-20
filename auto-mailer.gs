/**
 * AUTO MAILER script. Based on Google's mail merge script.
 */

/**
 * SCRIPT CONFIG
 */
const RECIPIENT_COL  = "Recipient";
const EMAIL_SENT_COL = "Email Sent";
const BATCH_SIZE = 50;

/** Creates the menu item "Auto Mailer" for user to run scripts on drop-down. */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Auto Mailer')
    .addItem('Start Auto Mailer', 'startAutoMailer')
    .addSeparator()
    .addItem('Reset Auto Mailer', 'resetMailer')
    .addToUi();
}

/**
 * Gets the subject line from the user and creates the time-based trigger.
 */
function startAutoMailer() {
  const subjectLine = Browser.inputBox("Start Auto Mailer",
                                       "Enter the subject line of your Gmail draft:",
                                       Browser.Buttons.OK_CANCEL);

  if (subjectLine === "cancel" || subjectLine === "") return;

  const properties = PropertiesService.getUserProperties();
  properties.setProperty('autoMailerSubject', subjectLine);
  properties.setProperty('lastProcessedRow', '1');

  deleteTriggers_();

  ScriptApp.newTrigger('sendEmails')
    .timeBased()
    .everyMinutes(5)
    .create();

  SpreadsheetApp.getUi().alert(
    'Auto Mailer Started',
    'The first emails will go out within 5 minutes and will continue automatically in batches. You can safely close this window now.',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

/**
 * Sends emails in batches.
 */
function sendEmails() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const properties = PropertiesService.getUserProperties();
  const currentUser = Session.getActiveUser().getEmail(); // Get current user's email for notifications

  const subjectLine = properties.getProperty('autoMailerSubject');
  let lastProcessedRow = parseInt(properties.getProperty('lastProcessedRow'), 10);

  if (!subjectLine) {
    deleteTriggers_();
    return;
  }

  const emailTemplate = getGmailTemplateFromDrafts_(subjectLine);
  const data = sheet.getDataRange().getDisplayValues();
  const heads = data[0];
  const emailSentColIdx = heads.indexOf(EMAIL_SENT_COL);

  let emailsSentInThisBatch = 0;

  for (let i = lastProcessedRow; i < data.length; i++) {
    if (emailsSentInThisBatch >= BATCH_SIZE) break;

    const rowData = data[i];
    const emailStatus = rowData[emailSentColIdx];

    if (emailStatus === '') {
      const rowObject = heads.reduce((o, k, j) => {
        o[k] = rowData[j] || '';
        return o;
      }, {});

      try {
        const msgObj = fillInTemplateFromObject_(emailTemplate.message, rowObject);

        GmailApp.sendEmail(rowObject[RECIPIENT_COL], msgObj.subject, msgObj.text, {
          htmlBody: msgObj.html,
          attachments: emailTemplate.attachments,
          inlineImages: emailTemplate.inlineImages
        });

        sheet.getRange(i + 1, emailSentColIdx + 1).setValue(new Date());
        emailsSentInThisBatch++;
        Utilities.sleep(1000);

      } catch (e) {
        if (e.message.includes('Service invoked too many times')) {

          // Email notification to the user with the error
          const subject = 'Auto Mailer Error: Quota Reached';
          const body = 'The Auto Mailer script has stopped because you have reached your daily email sending quota. Please reset the mailer and try again in 24 hours.';
          MailApp.sendEmail(currentUser, subject, body);
          deleteTriggers_();
          return;
        } else {
          sheet.getRange(i + 1, emailSentColIdx + 1).setValue(e.message);
        }
      }
    }
    lastProcessedRow = i;
    properties.setProperty('lastProcessedRow', lastProcessedRow.toString());
  }

  if (lastProcessedRow >= data.length - 1) {
    resetMailer();
    // Email notification to the user
    const subject = 'Auto Mailer Complete';
    const body = 'The Auto Mailer script has successfully processed all emails.';
    MailApp.sendEmail(currentUser, subject, body);
  }
}
/**
 * Deletes all triggers and clears user properties.
 */
function resetMailer() {
  deleteTriggers_();
  PropertiesService.getUserProperties().deleteAllProperties();
}

/**
 * Deletes all of the script's triggers.
 */
function deleteTriggers_() {
  const triggers = ScriptApp.getProjectTriggers();
  for (let i = 0; i < triggers.length; i++) {
    ScriptApp.deleteTrigger(triggers[i]);
  }
}


// --- UTILITY FUNCTIONS ---

function getGmailTemplateFromDrafts_(subject_line){
  try {
    const drafts = GmailApp.getDrafts();
    const draft = drafts.filter(subjectFilter_(subject_line))[0];
    const msg = draft.getMessage();
    const allInlineImages = msg.getAttachments({includeInlineImages: true,includeAttachments:false});
    const attachments = msg.getAttachments({includeInlineImages: false});
    const htmlBody = msg.getBody();
    const img_obj = allInlineImages.reduce((obj, i) => (obj[i.getName()] = i, obj) ,{});
    const imgexp = RegExp('<img.*?src="cid:(.*?)".*?alt="(.*?)"[^\>]+>', 'g');
    const matches = [...htmlBody.matchAll(imgexp)];
    const inlineImagesObj = {};
    matches.forEach(match => inlineImagesObj[match[1]] = img_obj[match[2]]);
    return {message: {subject: subject_line, text: msg.getPlainBody(), html:htmlBody},
            attachments: attachments, inlineImages: inlineImagesObj };
  } catch(e) {
    throw new Error("Oops - can't find Gmail draft with subject: " + subject_line);
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
  return  JSON.parse(template_string);
}

function escapeData_(str) {
  return str.toString()
    .replace(/[\\]/g, '\\\\')
    .replace(/[\"]/g, '\\\"')
    .replace(/[\/]/g, '\\/')
    .replace(/[\b]/g, '\\b')
    .replace(/[\f]/g, '\\f')
    .replace(/[\n]/g, '\\n')
    .replace(/[\r]/g, '\\r')
    .replace(/[\t]/g, '\\t');
};