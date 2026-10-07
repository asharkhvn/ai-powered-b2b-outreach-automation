const SHEET_NAME = "Sheet1";

// Production safety limit.
// Start low. Increase later after you verify deliverability.
const DAILY_SEND_LIMIT = 20;

const COL = {
  COMPANY: 1,
  WEBSITE: 2,
  INDUSTRY: 3,
  EMPLOYEE_SIZE: 4,
  CONTACT_NAME: 5,
  CONTACT_TITLE: 6,
  CONTACT_EMAIL: 7,
  LINKEDIN: 8,
  EVIDENCE: 9,
  PAIN_POINT: 10,
  OPPORTUNITY: 11,
  SCORE: 12,
  ANGLE: 13,
  SUBJECT: 14,
  BODY: 15,
  STATUS: 16,
  DATE_CONTACTED: 17,
  FOLLOWUP_DATE: 18,
  FOLLOWUP_1: 19,
  FOLLOWUP_2: 20,
  NOTES: 21,
  APPROVED: 22,
  LAST_ACTION: 23
};


// ============================================================
// MENU
// ============================================================

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("CCaaS Outreach")
    .addItem("Test Email", "testEmail")
    .addItem("Prepare Emails", "prepareEmails")
    .addItem("Send Approved Emails", "sendApprovedEmails")
    .addItem("Create Follow-Up Dates", "createFollowUpDates")
    .addItem("Check for Replies", "checkForReplies")
    .addItem("Process Follow-Ups", "processFollowUps")
    .addItem("Install Production Trigger", "installProductionTrigger")
    .addItem("Remove Production Trigger", "removeProductionTrigger")
    .addToUi();
}


// ============================================================
// TEST EMAIL
// ============================================================

function testEmail() {

  const email = Session.getActiveUser().getEmail();

  GmailApp.sendEmail(
    email,
    "CCaaS Outreach Test",
    "This is a test of your automated CCaaS outreach system."
  );

  SpreadsheetApp.getUi().alert(
    "Test email sent to: " + email
  );
}


// ============================================================
// PREPARE EMAILS
// ============================================================

function prepareEmails() {

  const sheet =
    SpreadsheetApp.getActiveSpreadsheet()
      .getSheetByName(SHEET_NAME);

  const data =
    sheet.getDataRange().getValues();

  let prepared = 0;

  for (let i = 1; i < data.length; i++) {

    const company =
      data[i][COL.COMPANY - 1];

    const email =
      data[i][COL.CONTACT_EMAIL - 1];

    const subject =
      data[i][COL.SUBJECT - 1];

    const body =
      data[i][COL.BODY - 1];

    const status =
      data[i][COL.STATUS - 1];

    if (
      company &&
      email &&
      subject &&
      body &&
      (!status || status === "New")
    ) {

      sheet
        .getRange(i + 1, COL.STATUS)
        .setValue("Ready");

      prepared++;
    }
  }

  SpreadsheetApp.getUi().alert(
    prepared + " prospect(s) are ready."
  );
}


// ============================================================
// DAILY SEND COUNT
// ============================================================

function getTodaySendCount() {

  const sheet =
    SpreadsheetApp.getActiveSpreadsheet()
      .getSheetByName(SHEET_NAME);

  const data =
    sheet.getDataRange().getValues();

  const today =
    new Date();

  today.setHours(0, 0, 0, 0);

  let count = 0;

  for (let i = 1; i < data.length; i++) {

    const date =
      data[i][COL.DATE_CONTACTED - 1];

    if (!date) {
      continue;
    }

    const sentDate =
      new Date(date);

    sentDate.setHours(0, 0, 0, 0);

    if (
      sentDate.getTime() === today.getTime()
    ) {
      count++;
    }
  }

  return count;
}


// ============================================================
// SEND INITIAL EMAILS
// ============================================================

function sendApprovedEmails() {

  const sheet =
    SpreadsheetApp.getActiveSpreadsheet()
      .getSheetByName(SHEET_NAME);

  const data =
    sheet.getDataRange().getValues();

  let sent = 0;

  let dailyCount =
    getTodaySendCount();

  for (let i = 1; i < data.length; i++) {

    if (dailyCount >= DAILY_SEND_LIMIT) {
      break;
    }

    const row = i + 1;

    const email =
      String(
        data[i][COL.CONTACT_EMAIL - 1]
      ).trim();

    const subject =
      String(
        data[i][COL.SUBJECT - 1]
      ).trim();

    const body =
      String(
        data[i][COL.BODY - 1]
      ).trim();

    const status =
      data[i][COL.STATUS - 1];

    const approved =
      String(
        data[i][COL.APPROVED - 1]
      ).toUpperCase();

    if (
      email &&
      subject &&
      body &&
      status === "Ready" &&
      approved === "YES"
    ) {

      try {

        GmailApp.sendEmail(
          email,
          subject,
          body
        );

        sheet
          .getRange(row, COL.STATUS)
          .setValue("Contacted");

        sheet
          .getRange(row, COL.DATE_CONTACTED)
          .setValue(new Date());

        sheet
          .getRange(row, COL.LAST_ACTION)
          .setValue(
            "Initial outreach sent"
          );

        sent++;
        dailyCount++;

      } catch (error) {

        sheet
          .getRange(row, COL.LAST_ACTION)
          .setValue(
            "ERROR: " + error.message
          );
      }
    }
  }

  SpreadsheetApp.getUi().alert(
    sent +
    " approved initial email(s) sent.\n\n" +
    "Daily limit: " +
    DAILY_SEND_LIMIT
  );
}


// ============================================================
// CREATE FOLLOW-UP DATES
// ============================================================

function createFollowUpDates() {

  const sheet =
    SpreadsheetApp.getActiveSpreadsheet()
      .getSheetByName(SHEET_NAME);

  const data =
    sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {

    const status =
      data[i][COL.STATUS - 1];

    const contacted =
      data[i][COL.DATE_CONTACTED - 1];

    const followup =
      data[i][COL.FOLLOWUP_DATE - 1];

    if (
      status === "Contacted" &&
      contacted &&
      !followup
    ) {

      const date =
        new Date(contacted);

      date.setDate(
        date.getDate() + 5
      );

      sheet
        .getRange(
          i + 1,
          COL.FOLLOWUP_DATE
        )
        .setValue(date);
    }
  }

  SpreadsheetApp.getUi().alert(
    "Follow-up dates created."
  );
}


// ============================================================
// REPLY DETECTION
// ============================================================

function checkForReplies() {

  const result =
    checkForRepliesInternal();

  SpreadsheetApp.getUi().alert(
    result +
    " prospect reply/replies detected."
  );
}


function checkForRepliesInternal() {

  const sheet =
    SpreadsheetApp.getActiveSpreadsheet()
      .getSheetByName(SHEET_NAME);

  const data =
    sheet.getDataRange().getValues();

  let repliesFound = 0;

  for (let i = 1; i < data.length; i++) {

    const row = i + 1;

    const email =
      String(
        data[i][COL.CONTACT_EMAIL - 1]
      ).trim();

    const status =
      data[i][COL.STATUS - 1];

    const contacted =
      data[i][COL.DATE_CONTACTED - 1];

    if (
      !email ||
      !contacted
    ) {
      continue;
    }

    if (
      status !== "Contacted" &&
      status !== "Follow-Up 1"
    ) {
      continue;
    }

    const contactDate =
      new Date(contacted);

    const afterDate =
      formatGmailDate(contactDate);

    const query =
      'from:(' +
      email +
      ') after:' +
      afterDate;

    let threads = [];

    try {

      threads =
        GmailApp.search(query);

    } catch (error) {

      sheet
        .getRange(row, COL.LAST_ACTION)
        .setValue(
          "Reply check error: " +
          error.message
        );

      continue;
    }

    let replyFound = false;

    for (const thread of threads) {

      const messages =
        thread.getMessages();

      for (const message of messages) {

        const sender =
          message
            .getFrom()
            .toLowerCase();

        const senderEmail =
          extractEmailAddress(sender)
            .toLowerCase();

        if (
          senderEmail ===
          email.toLowerCase() &&
          message.getDate() >=
          contactDate
        ) {

          replyFound = true;
          break;
        }
      }

      if (replyFound) {
        break;
      }
    }

    if (replyFound) {

      sheet
        .getRange(row, COL.STATUS)
        .setValue("Responded");

      sheet
        .getRange(
          row,
          COL.FOLLOWUP_DATE
        )
        .clearContent();

      sheet
        .getRange(
          row,
          COL.LAST_ACTION
        )
        .setValue(
          "Reply detected"
        );

      repliesFound++;
    }
  }

  return repliesFound;
}


// ============================================================
// GMAIL DATE FORMAT
// ============================================================

function formatGmailDate(date) {

  const year =
    date.getFullYear();

  const month =
    String(
      date.getMonth() + 1
    ).padStart(2, "0");

  const day =
    String(
      date.getDate()
    ).padStart(2, "0");

  return (
    year +
    "/" +
    month +
    "/" +
    day
  );
}


// ============================================================
// EXTRACT EMAIL ADDRESS
// ============================================================

function extractEmailAddress(fromField) {

  const match =
    fromField.match(
      /<([^>]+)>/
    );

  if (match) {
    return match[1].trim();
  }

  return fromField.trim();
}


// ============================================================
// PROCESS FOLLOW-UPS
// ============================================================

function processFollowUps() {

  // ALWAYS check for replies first.
  checkForRepliesInternal();

  const sheet =
    SpreadsheetApp
      .getActiveSpreadsheet()
      .getSheetByName(SHEET_NAME);

  const data =
    sheet.getDataRange().getValues();

  const today =
    new Date();

  let sent = 0;

  let dailyCount =
    getTodaySendCount();

  for (let i = 1; i < data.length; i++) {

    if (
      dailyCount >= DAILY_SEND_LIMIT
    ) {
      break;
    }

    const row = i + 1;

    const email =
      String(
        data[i][COL.CONTACT_EMAIL - 1]
      ).trim();

    const status =
      data[i][COL.STATUS - 1];

    const followupDate =
      data[i][COL.FOLLOWUP_DATE - 1];

    const followup1 =
      String(
        data[i][COL.FOLLOWUP_1 - 1] || ""
      ).trim();

    const followup2 =
      String(
        data[i][COL.FOLLOWUP_2 - 1] || ""
      ).trim();

    const stopStatuses = [
      "Responded",
      "Meeting",
      "Proposal",
      "Won",
      "Lost",
      "Not a Fit"
    ];

    if (
      stopStatuses.includes(status)
    ) {
      continue;
    }

    if (!email || !followupDate) {
      continue;
    }


    // ========================================================
    // FOLLOW-UP #1
    // ========================================================

    if (
      status === "Contacted" &&
      today >=
      new Date(followupDate) &&
      followup1
    ) {

      try {

        GmailApp.sendEmail(
          email,
          "Re: Quick question about customer support",
          followup1
        );

        sheet
          .getRange(row, COL.STATUS)
          .setValue("Follow-Up 1");

        sheet
          .getRange(row, COL.LAST_ACTION)
          .setValue(
            "Follow-Up 1 sent"
          );

        const nextDate =
          new Date(today);

        nextDate.setDate(
          nextDate.getDate() + 7
        );

        sheet
          .getRange(
            row,
            COL.FOLLOWUP_DATE
          )
          .setValue(nextDate);

        sent++;
        dailyCount++;

      } catch (error) {

        sheet
          .getRange(
            row,
            COL.LAST_ACTION
          )
          .setValue(
            "ERROR: " +
            error.message
          );
      }

      continue;
    }


    // ========================================================
    // FOLLOW-UP #2
    // ========================================================

    if (
      status === "Follow-Up 1" &&
      today >=
      new Date(followupDate) &&
      followup2
    ) {

      try {

        GmailApp.sendEmail(
          email,
          "Re: Quick question about customer support",
          followup2
        );

        sheet
          .getRange(row, COL.STATUS)
          .setValue("Follow-Up 2");

        sheet
          .getRange(row, COL.LAST_ACTION)
          .setValue(
            "Follow-Up 2 sent"
          );

        sent++;
        dailyCount++;

      } catch (error) {

        sheet
          .getRange(
            row,
            COL.LAST_ACTION
          )
          .setValue(
            "ERROR: " +
            error.message
          );
      }
    }
  }

  SpreadsheetApp.getUi().alert(
    sent +
    " follow-up email(s) sent."
  );
}


// ============================================================
// PRODUCTION TRIGGER
// ============================================================

function installProductionTrigger() {

  removeProductionTrigger();

  ScriptApp.newTrigger(
    "processFollowUps"
  )
    .timeBased()
    .everyHours(1)
    .create();

  SpreadsheetApp.getUi().alert(
    "Production trigger installed.\n\n" +
    "The system will check for replies and process due follow-ups approximately once per hour."
  );
}


// ============================================================
// REMOVE PRODUCTION TRIGGER
// ============================================================

function removeProductionTrigger() {

  const triggers =
    ScriptApp.getProjectTriggers();

  for (const trigger of triggers) {

    if (
      trigger.getHandlerFunction() ===
      "processFollowUps"
    ) {

      ScriptApp.deleteTrigger(
        trigger
      );
    }
  }

  SpreadsheetApp.getUi().alert(
    "Production trigger removed."
  );
}
