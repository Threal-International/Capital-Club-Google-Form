function onFormSubmit(e) {
  updateStatus(e);

  // After updating the status, introduce a delay and send the email.
  Utilities.sleep(5000) // Adjust the delay as needed
  sendEmail()
}

function updateStatus(e) {
  // Get the user's Id and applied loan amount from the submitted form.
  // The following are the index values 0-date, 1-name, 2-id, 3-amount, 4-loanduration, 5-gurantors, 6-email - this follows the form inputs
  var userId = e.values[2].toString(); // assuming user id is the third column
  var userEmail = e.values[5].toString();
  var appliedAmount = e.values[3]; // assuming the applied amount is the fourth column

  // Get the user's limit from the members list sheet: 0-name, 1-email, 2-id, 3-limit
  var membersSheet = PropertiesService.getScriptProperties().getProperty("MEMBER_SHEET_NAME");
  if (!membersSheet) {
    Logger.log("Error: MEMBER_SHEET_NAME script property not set!");
    return; // Stop execution
  }
  var limitSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(membersSheet);;
  var limitData = limitSheet.getDataRange().getValues();

  // Check if the user is a member
  var isMember = false;
  for (var i = 1; i < limitData.length; i++) {
    // We check that both the userId and email belong to the user hence they are a member of the group
    if (limitData[i][2].toString() == userId && limitData[i][1].toString() == userEmail) {
      isMember = true;
      break;
    }
  }

  // Check for outstanding loans
  // The spreadsheet column numbers are: 0-date, 1-name, 2-id, 3-amount, 4-duration, 5-email, 6-loanstatus, 7-repaymentstatus, 8-rejectionreason
  var idColumn = 2
  var statusColumn = 6
  var repaymentColumn = 7
  var sheetID = PropertiesService.getScriptProperties().getProperty("FORM_SHEET_ID");
  if (!sheetID) {
    Logger.log("Error: FORM_SHEET_ID script property not set!");
    return; // Stop execution
  }
  var formResponseSheet = SpreadsheetApp.openById(sheetID);
  var responseData = formResponseSheet.getDataRange().getValues();
  var hasOutstandingLoan = false;
  for (var i = 1; i < responseData.length; i++) {
    if (responseData[i][idColumn].toString() == userId) {
      if (responseData[i][statusColumn] == "Accepted" && responseData[i][repaymentColumn] != "Paid") {
        hasOutstandingLoan = true;
        break;
      }
    }
  }

  var sheet = e.source.getActiveSheet();
  var row = e.range.getRow();

  // If the user is not a member or has entered wrong id, reject the application
  if (!isMember) {
    sheet.getRange(row, 7).setValue("Rejected");
    sheet.getRange(row, 8).setValue("Application rejected")
    sheet.getRange(row, 9).setValue("Not a member of the chama")
    return; // Exit the function
  }

  // Find the user's limit and update the status
  if (!hasOutstandingLoan) {
    var userLimit = 0;
    for (var i = 1; i < limitData.length; i++) {
      if (limitData[i][2].toString() == userId) {
        userLimit = limitData[i][3]; // Assuming limit is in the third column of the limit sheet
        break;
      }
    }

    // Compare the applied amount to the user's limit and update the status
    if (appliedAmount <= userLimit) {
      sheet.getRange(row, 7).setValue("Accepted"); // Assuming status column is the 7th column
      sheet.getRange(row, 8).setValue("Awaiting remittance")
    } else {
      sheet.getRange(row, 7).setValue("Rejected");
      sheet.getRange(row, 8).setValue("Application rejected")
      sheet.getRange(row, 9).setValue("Requested amount is above your borrowing limit")
    }
  } else {
    // Reject the application due to outstanding loan
    sheet.getRange(row, 7).setValue("Rejected");
    sheet.getRange(row, 8).setValue("Application rejected")
    sheet.getRange(row, 9).setValue("You have an outstanding loan")
  }
}
