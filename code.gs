// Function to add a sender email
function addSenderEmail(newEmail) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Email List");
  var lastRow = sheet.getLastRow();
  
  // Insert new email at the first empty row in the email list sheet
  sheet.getRange(lastRow + 1, 1).setValue(newEmail);
}

// Function to get the list of sender emails
function getSenderEmails() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Email List");
  var emailData = sheet.getDataRange().getValues();
  var emailList = [];
  
  // Loop through the data and get the emails (assuming emails are in column 1)
  for (var i = 0; i < emailData.length; i++) {
    if (emailData[i][0] && emailData[i][0] !== "") {
      emailList.push(emailData[i][0]);
    }
  }
  
  return emailList;
}

// Function to schedule emails
function scheduleEmails(emailColumn, subjectColumn, bodyColumn, startRow, endRow, senderEmail, dailyLimit, randomDelayMin, randomDelayMax, scheduleDate, scheduleTime) {
  // Access the sheet where emails are stored (assuming it's called "Emails")
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Emails");
  var data = sheet.getRange(startRow, emailColumn, endRow - startRow + 1, 3).getValues(); // Get emails, subjects, bodies
  
  // Iterate through the data and schedule each email
  for (var i = 0; i < data.length; i++) {
    var recipient = data[i][0]; // Email
    var subject = data[i][1]; // Subject
    var body = data[i][2]; // Body
    
    // Apply the random delay
    var delay = Math.floor(Math.random() * (randomDelayMax - randomDelayMin + 1) + randomDelayMin) * 60 * 1000; // Delay in milliseconds
    var dateTime = new Date(scheduleDate + " " + scheduleTime); // Date and time to send email
    
    // Schedule the email (this will run after the delay)
    ScriptApp.newTrigger("sendScheduledEmail")
      .timeBased()
      .at(dateTime)
      .after(delay) // Adding random delay
      .create();
  }
}

// Function to send a scheduled email
function sendScheduledEmail() {
  // Logic to send the email
  // (You'll need to store email details somewhere to fetch and send them when the trigger executes)
  var emailDetails = getScheduledEmailDetails(); // Dummy function to get scheduled emails
  
  // Send the email (dummy email sending)
  MailApp.sendEmail({
    to: emailDetails.recipient,
    subject: emailDetails.subject,
    body: emailDetails.body,
    from: emailDetails.sender
  });
}

// Function to clear scheduled emails
function clearScheduledEmails() {
  var triggers = ScriptApp.getProjectTriggers();
  
  // Iterate over existing triggers and delete them
  for (var i = 0; i < triggers.length; i++) {
    ScriptApp.deleteTrigger(triggers[i]);
  }
}

// Function to send a test email
function sendTestEmail(senderEmail) {
  var testRecipient = "test@example.com"; // Replace with your test recipient email
  
  // Send the test email
  MailApp.sendEmail({
    to: testRecipient,
    subject: "Test Email",
    body: "This is a test email sent from Google Apps Script.",
    from: senderEmail
  });
}

// Function to get the current sender email list for the frontend
function getCurrentSenderEmail() {
  // Assuming the sender email is stored somewhere
  var properties = PropertiesService.getScriptProperties();
  var senderEmail = properties.getProperty("senderEmail");
  
  if (!senderEmail) {
    senderEmail = "default@example.com"; // Provide a default if no sender email is set
  }
  
  return senderEmail;
}

// Function to set the sender email in script properties (e.g., when added via frontend)
function setSenderEmail(senderEmail) {
  var properties = PropertiesService.getScriptProperties();
  properties.setProperty("senderEmail", senderEmail);
}
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Sheet Draft by Josh')
    .addItem('Open Email Dialog', 'openEmailDialog')
    .addToUi();
}


function openEmailDialog() {
  const htmlOutput = HtmlService.createHtmlOutputFromFile('EmailDialog')
    .setWidth(600)
    .setHeight(800);
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Sheet Draft by Josh');
}


function getSheetHeaders() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  return headers;
}


function getPreviewData(column) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const data = sheet.getRange(2, column, 1, 1).getValues()[0][0];
  // Increased substring length to show more preview data
  return data ? data.substring(0, 20) : '';
}


function getSendAsEmails() {
  const aliases = GmailApp.getAliases();
  aliases.unshift(Session.getActiveUser().getEmail());
  return aliases;
}


function sendTestEmail(senderEmail) {
  try {
    GmailApp.sendEmail(Session.getActiveUser().getEmail(), 'Test Email', 'This is a test email from ' + senderEmail);
    return 'Test email sent successfully.';
  } catch (e) {
    // Improved error logging
    Logger.log('Error sending test email: ' + e.message);
    return 'Error sending test email: ' + e.message;
  }
}


function saveDrafts(emailColumn, subjectColumn, bodyColumn, startRow, endRow, senderEmail, dailyLimit, randomDelayMin, randomDelayMax) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const dataRange = sheet.getRange(startRow, 1, endRow - startRow + 1, sheet.getLastColumn());
  const data = dataRange.getValues();
  let draftCount = 0;
  let dailySentCount = 0;


  try {
    data.forEach((row, index) => {
      const email = row[emailColumn];
      const subject = row[subjectColumn];
      const body = row[bodyColumn];


      // Limiting dailySentCount within the current script execution.
      // For a true daily limit across executions, persistent storage is needed.
      if (dailySentCount >= dailyLimit) {
        return; // Exit forEach loop iteration, not the entire function
      }


      if (email && subject && body) {
        if (randomDelayMin && randomDelayMax) {
          const delay = Math.floor(Math.random() * (randomDelayMax - randomDelayMin + 1)) + randomDelayMin;
          Utilities.sleep(delay * 1000);
        }


        GmailApp.createDraft(email, subject, body, { from: senderEmail });
        draftCount++;
        dailySentCount++;
      }
    });


    return {
      draftCount,
      dailySentCount,
      message: 'Drafts saved successfully.'
    };
  } catch (e) {
    // Improved error logging
    Logger.log('Error saving drafts: ' + e.message);
    return {
      draftCount,
      dailySentCount,
      message: 'Error saving drafts: ' + e.message
    };
  }
}


function sendEmails(emailColumn, subjectColumn, bodyColumn, startRow, endRow, senderEmail, dailyLimit, randomDelayMin, randomDelayMax) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const dataRange = sheet.getRange(startRow, 1, endRow - startRow + 1, sheet.getLastColumn());
  const data = dataRange.getValues();
  let sentCount = 0;
  let dailySentCount = 0;


  try {
    data.forEach((row, index) => {
      const email = row[emailColumn];
      const subject = row[subjectColumn];
      const body = row[bodyColumn];


      // Limiting dailySentCount within the current script execution.
      // For a true daily limit across executions, persistent storage is needed.
      if (dailySentCount >= dailyLimit) {
        return; // Exit forEach loop iteration, not the entire function
      }


      if (email && subject && body) {
        if (randomDelayMin && randomDelayMax) {
          const delay = Math.floor(Math.random() * (randomDelayMax - randomDelayMin + 1)) + randomDelayMax;
          Utilities.sleep(delay * 1000);
        }


        GmailApp.sendEmail(email, subject, body, { from: senderEmail });
        sentCount++;
        dailySentCount++;
      }
    });


    return {
      sentCount,
      dailySentCount,
      message: 'Emails sent successfully.'
    };
  } catch (e) {
    // Improved error logging
    Logger.log('Error sending emails: ' + e.message);
    return {
      sentCount,
      dailySentCount,
      message: 'Error sending emails: ' + e.message
    };
  }
}


function scheduleEmails(emailColumn, subjectColumn, bodyColumn, startRow, endRow, senderEmail, dailyLimit, randomDelayMin, randomDelayMax, scheduleDate, scheduleTime) {
  try {
    // Schedule time is expected in HH:mm 24-hour format
    const triggerId = ScriptApp.newTrigger('processScheduledEmails')
      .timeBased()
      .at(new Date(`${scheduleDate}T${scheduleTime}:00Z`))
      .create()
      .getUniqueId(); // getUniqueId() is called after create() but not used before


    PropertiesService.getScriptProperties().setProperty('EMAIL_TRIGGER_ID', triggerId);
    PropertiesService.getScriptProperties().setProperty('EMAIL_COLUMN', emailColumn);
    PropertiesService.getScriptProperties().setProperty('SUBJECT_COLUMN', subjectColumn);
    PropertiesService.getScriptProperties().setProperty('BODY_COLUMN', bodyColumn);
    PropertiesService.getScriptProperties().setProperty('START_ROW', startRow);
    PropertiesService.getScriptProperties().setProperty('END_ROW', endRow);
    PropertiesService.getScriptProperties().setProperty('SENDER_EMAIL', senderEmail);
    PropertiesService.getScriptProperties().setProperty('DAILY_LIMIT', dailyLimit);
    PropertiesService.getScriptProperties().setProperty('RANDOM_DELAY_MIN', randomDelayMin);
    PropertiesService.getScriptProperties().setProperty('RANDOM_DELAY_MAX', randomDelayMax);


    return 'Emails scheduled successfully.';
  } catch (e) {
    // Improved error logging
    Logger.log('Error scheduling emails: ' + e.message);
    return 'Error scheduling emails: ' + e.message;
  }
}


function processScheduledEmails() {
  const properties = PropertiesService.getScriptProperties();
  const emailColumn = properties.getProperty('EMAIL_COLUMN');
  const subjectColumn = properties.getProperty('SUBJECT_COLUMN');
  const bodyColumn = properties.getProperty('BODY_COLUMN');
  const startRow = properties.getProperty('START_ROW');
  const endRow = properties.getProperty('END_ROW');
  const senderEmail = properties.getProperty('SENDER_EMAIL');
  const dailyLimit = properties.getProperty('DAILY_LIMIT');
  const randomDelayMin = properties.getProperty('RANDOM_DELAY_MIN');
  const randomDelayMax = properties.getProperty('RANDOM_DELAY_MAX');


  sendEmails(emailColumn, subjectColumn, bodyColumn, startRow, endRow, senderEmail, dailyLimit, randomDelayMin, randomDelayMax);
  ScriptApp.getProjectTriggers().forEach(trigger => {
    if (trigger.getUniqueId() === properties.getProperty('EMAIL_TRIGGER_ID')) {
      ScriptApp.deleteTrigger(trigger);
    }
  });
  properties.deleteAllProperties();
}


function clearScheduledEmails() {
  const triggerId = PropertiesService.getScriptProperties().getProperty('EMAIL_TRIGGER_ID');
  if (triggerId) {
    ScriptApp.getProjectTriggers().forEach(trigger => {
      if (trigger.getUniqueId() === triggerId) {
        ScriptApp.deleteTrigger(trigger);
      }
    });
    PropertiesService.getScriptProperties().deleteAllProperties();
    return 'Scheduled emails cleared.';
  } else {
    return 'No scheduled emails found.';
  }
}


function addSenderEmail(email) {
  const emails = PropertiesService.getUserProperties().getProperty('SENDER_EMAILS');
  const senderEmails = emails ? JSON.parse(emails) : [];
  senderEmails.push(email);
  PropertiesService.getUserProperties().setProperty('SENDER_EMAILS', JSON.stringify(senderEmails));
}


function getSenderEmails() {
  const emails = PropertiesService.getUserProperties().getProperty('SENDER_EMAILS');
  return emails ? JSON.parse(emails) : [];
}


function removeSenderEmail(email) {
  const emails = PropertiesService.getUserProperties().getProperty('SENDER_EMAILS');
  let senderEmails = emails ? JSON.parse(emails) : [];
  senderEmails = senderEmails.filter(e => e !== email);
  PropertiesService.getUserProperties().setProperty('SENDER_EMAILS', JSON.stringify(senderEmails));
}


// New function to get current user email
function getCurrentUserEmail() {
  return Session.getActiveUser().getEmail();
}
