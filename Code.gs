var googleFormUrl = 'https://docs.google.com/forms/d/e/1FAIpQLSe3jpDnAKA8GmQeAqtu8Cp3e7vWw7hLt2KlRSq-G_sv5VWZlQ/viewform'; // Replace with your Google Form URL

function doGet(e) {
  var htmlOutput = HtmlService.createTemplateFromFile("index");
  return htmlOutput.evaluate().setTitle("COLLAPSIBLE FORM")
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function onFormSubmit(e) {
  var sheetName = "Leave Request";
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  var activeRow = e.range.getRow();

  // Fetch values directly from columns A to Z in a single line
  var [
    columnA, columnB, columnC, columnD, columnE, columnF, columnG, columnH, columnI, columnJ,
    columnK, columnL, columnM, columnN, columnO, columnP, columnQ, columnR, columnS,
    columnT, columnU, columnV, columnW, columnX, columnY, columnZ
  ] = sheet.getRange(activeRow, 1, 1, 26).getValues()[0];

  sheet.getRange(activeRow, 10).setValue("Pending");

  var duration = calculateDateDifference(columnE, columnF); 
  sheet.getRange(activeRow, 9).setValue(duration); 

  var leaveBalance = getLeaveBalance(columnB, columnD); 
  
  if (leaveBalance === undefined || leaveBalance === 0) {
    sendSufficientBalanceEmail(columnB, columnC, columnD, duration);
  } else {
    if (duration > leaveBalance) {
      sendInsufficientBalanceEmail(columnB, columnC, columnD, duration);
      sheet.getRange(activeRow, 10).setValue("Auto-Reject");
    } else {
      sendSufficientBalanceEmail(columnB, columnC, columnD, duration);
    }
  }
}

function onColumnChangeApprovedReject(e) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Leave Request");
  var activeCell = e.source.getActiveRange();

  if (activeCell.getColumn() === 10) {
    var row = activeCell.getRow();

    var email = sheet.getRange(row, 2).getValue(); 
    var name = sheet.getRange(row, 3).getValue();
    var leaveType = sheet.getRange(row, 4).getValue();
    var columnE = sheet.getRange(row, 5).getValue(); 
    var columnF = sheet.getRange(row, 6).getValue(); 
    var fromDate = new Date(columnE);
    var toDate = new Date(columnF);

    var formattedFromDate = formatDate(fromDate);
    var formattedToDate = formatDate(toDate);

    var columnChange = activeCell.getValue();
    if (columnChange === "Approved") {
      var mailSubject = "Leave Application Approved - " + name;
      var mailBody =
        "Dear " + name + ",<br><br>" +
        "Your leave application has been approved. We understand the importance of taking time off to recharge and attend to personal matters, and we are pleased to accommodate your request.<br>" +
        "We appreciate your dedication and hard work within the organization and look forward to your return, refreshed and ready to contribute your best.<br><br>" +
        "Leave Date: " + formattedFromDate + " to " + formattedToDate + "<br><br>" +
        "Enjoy your well-deserved break!<br><br>" +
        "Your current leave balance for " + leaveType + " is " + getLeaveBalance(email, leaveType) + " days.<br>" +
        "You can apply for another leave using our <a href='" + googleFormUrl + "'>leave application form</a>.<br><br>" +
        "Thank You,<br>" +
        "HR Department.<br>";

      sendMail(email, mailSubject, mailBody);
      createCalendarEvent(fromDate, toDate, eventTitle, eventDescription);

      var duration = calculateDateDifference(fromDate, toDate);
      updateLeaveBalanceAndUsed(email, leaveType, duration); 
    } else if (columnChange === "Reject") {
      var mailSubject = "Leave Application Rejected - " + name;
      var mailBody =
        "Dear " + name + ",<br><br>" +
        "Regret to inform you that your leave application, dated " + formattedFromDate + " to " + formattedToDate + ", has been reviewed and, unfortunately, it cannot be approved at this time.<br>" +
        "We understand that this may be disappointing, and we apologize for any inconvenience this may cause. We highly encourage you to discuss potential alternatives with your supervisor or team to manage your workload during this period or consider resubmitting your request for alternative dates if possible.<br><br>" +
        "Thank you for your understanding.<br><br>" +
        "Your current leave balance for " + leaveType + " is " + getLeaveBalance(email, leaveType) + " days.<br>" +
        "You can apply for another leave using our <a href='" + googleFormUrl + "'>leave application form</a>.<br><br>" +
        "Thank You,<br>" +
        "HR Department.<br>";

      sendMail(email, mailSubject, mailBody);
    } else if (columnChange === "Auto-Reject") {
      var mailSubject = "Leave Application Auto-Rejected - " + name;
      var mailBody =
        "Dear " + name + ",<br><br>" +
        "Your leave application, dated " + formattedFromDate + " to " + formattedToDate + ", has been auto-rejected due to insufficient leave balance.<br>" +
        "Please review your leave balance and consider adjusting your future requests accordingly.<br><br>" +
        "Thank you for your understanding.<br><br>" +
        "Your current leave balance for " + leaveType + " is " + getLeaveBalance(email, leaveType) + " days.<br>" +
        "You can apply for another leave using our <a href='" + googleFormUrl + "'>leave application form</a>.<br><br>" +
        "Thank You,<br>" +
        "HR Department.<br>";

      sendMail(email, mailSubject, mailBody);
    }
  }
}

function getLeaveBalance(email, leaveType) {
  var leaveBalanceSheetName = "Leave Balance"; 
  var leaveBalanceSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(leaveBalanceSheetName);
  var leaveBalanceData = leaveBalanceSheet.getDataRange().getValues();

  for (var i = 1; i < leaveBalanceData.length; i++) {
    // Assuming row 1 contains headers
    var sheetEmail = leaveBalanceData[i][0];
    var sheetLeaveType = leaveBalanceData[i][1];
    var sheetLeaveBalance = leaveBalanceData[i][4]; 

    if (sheetEmail === email && sheetLeaveType === leaveType) {
      return sheetLeaveBalance;
    }
  }

  return undefined;
}

function getLeaveType(email) {
  var leaveBalanceSheetName = "Leave Balance"; // Name of the "Leave Balance" sheet
  var leaveBalanceSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(leaveBalanceSheetName);
  var leaveBalanceData = leaveBalanceSheet.getDataRange().getValues();
  var leaveTypes = [];

  for (var i = 1; i < leaveBalanceData.length; i++) {
    // Assuming row 1 contains headers
    var sheetEmail = leaveBalanceData[i][0]; // Email is in the first column
    var sheetLeaveType = leaveBalanceData[i][1]; // Leave type is in the second column

    if (sheetEmail === email) {
      leaveTypes.push(sheetLeaveType);
    }
  }

  return leaveTypes; 
}

function sendInsufficientBalanceEmail(email, name, leaveType, duration) {
  var mailSubject = "Insufficient Leave Balance - " + leaveType;
  var leaveBalance = getLeaveBalance(email, leaveType); 
  var mailBody =
    "Dear " + name + ",<br><br>" +
    "Your request for " + leaveType + " leave has been received, but the requested duration (" + duration + " days) exceeds your available leave balance.<br><br>" +
    "Please review your leave balance and consider adjusting your request accordingly.<br><br>" +
    "Your current leave balance for " + leaveType + " is " + leaveBalance + " days.<br>" +
    "You can apply for another leave using our <a href='" + googleFormUrl + "'>leave application form</a>.<br><br>" +
    "Thank you,<br>" +
    "HR Department";

  sendMail(email, mailSubject, mailBody);
}

function sendSufficientBalanceEmail(email, name, leaveType, duration) {
  var mailSubject = "Leave Request Submitted - Pending";
  var leaveBalance = getLeaveBalance(email, leaveType); 
  var mailBody =
    "Dear " + name + ",<br><br>" +
    "Your request for " + leaveType + " has been submitted. The requested duration is " + duration + " days.<br>" +
    "After approval, you will receive an email notification.<br><br>" +
    "You can apply for another leave using our <a href='" + googleFormUrl + "'>leave application form</a>.<br><br>" +
    "Thank you,<br>" +
    "HR Department";

  sendMail(email, mailSubject, mailBody);
}

function updateLeaveBalanceAndUsed(email, leaveType, duration) {
  var leaveBalanceSheetName = "Leave Balance";
  var leaveBalanceSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(leaveBalanceSheetName);
  var leaveBalanceData = leaveBalanceSheet.getDataRange().getValues();

  for (var i = 1; i < leaveBalanceData.length; i++) {
    var sheetEmail = leaveBalanceData[i][0];
    var sheetLeaveType = leaveBalanceData[i][1];
    var sheetLeaveUsed = leaveBalanceData[i][2]; 
    var sheetLeaveBalance = leaveBalanceData[i][4];

    if (sheetEmail === email && sheetLeaveType === leaveType) {
      var newLeaveUsed = sheetLeaveUsed + duration;
      var newLeaveBalance = sheetLeaveBalance - duration;

      Logger.log("Updating leave used for " + email + " - Leave Type: " + leaveType);
      Logger.log("Old Leave Used: " + sheetLeaveUsed + ", New Leave Used: " + newLeaveUsed);
      Logger.log("Old Leave Balance: " + sheetLeaveBalance + ", New Leave Balance: " + newLeaveBalance);

      leaveBalanceSheet.getRange(i + 1, 3).setValue(newLeaveUsed);
      leaveBalanceSheet.getRange(i + 1, 5).setValue(newLeaveBalance);
      break;
    }
  }
}

function formatDate(date) {
  var options = { weekday: 'short', year: 'numeric', month: 'short', day: 'numeric' };
  return date.toLocaleDateString('en-SG', options);
}

function sendMail(sendTo, mailSubject, mailBody) {
  MailApp.sendEmail({
    to: sendTo,
    subject: mailSubject,
    htmlBody: mailBody,
  });
}

function calculateDateDifference(startDate, endDate) {
  var startTimestamp = new Date(startDate).getTime();
  var endTimestamp = new Date(endDate).getTime();
  if (isNaN(startTimestamp) || isNaN(endTimestamp) || startTimestamp > endTimestamp) {
    return "Invalid Date";
  }
  var millisecondsInADay = 1000 * 60 * 60 * 24;
  var differenceInDays = Math.floor((endTimestamp - startTimestamp) / millisecondsInADay) + 1;
  return differenceInDays;
}

function addEmployeeData(formData) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Leave Balance");
  var lastRow = sheet.getLastRow() + 1; // Get the next empty row
  sheet.appendRow([formData.email, formData.leaveType, 0, formData.totalLeave, ""]);
  
  var emailCell = sheet.getRange(lastRow, 1); 
  var leaveTypeCell = sheet.getRange(lastRow, 2); 
  var usedCell = sheet.getRange(lastRow, 3); 
  var totalCell = sheet.getRange(lastRow, 4); 
  var balanceCell = sheet.getRange(lastRow, 5); 
  
  var formulaUsed = `=IFERROR(SUMIFS('Leave Request'!$I$2:$I, 'Leave Request'!$B$2:$B, ${emailCell.getA1Notation()}, 'Leave Request'!$D$2:$D, ${leaveTypeCell.getA1Notation()}, 'Leave Request'!$J$2:$J, "Approved"), 0)`;
  usedCell.setFormula(formulaUsed);

  var formulaBalance = `=${totalCell.getA1Notation()} - ${usedCell.getA1Notation()}`;
  balanceCell.setFormula(formulaBalance);
}

function openAddEmployeeForm() {
  var html = HtmlService.createHtmlOutputFromFile('AddEmployeeForm')
      .setWidth(400)
      .setHeight(300);
  SpreadsheetApp.getUi().showModalDialog(html, 'Add Employee Information');
}
