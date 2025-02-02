function sendEmail() {
  var sheetGID =
    PropertiesService.getScriptProperties().getProperty("FORM_SHEET_GID");
  if (!sheetGID) {
    Logger.log("Error: FORM_SHEET_GID script property not set!");
    return; // Stop execution
  }
  var formresp = getSheetById(sheetGID);

  // collect the most recent response
  var recentResponse = formresp
    .getRange(formresp.getLastRow(), 1, 1, formresp.getLastColumn())
    .getDisplayValues();

  // The spreadsheet column numbers are: 0-date, 1-email, 2-name, 3-id, 4-amount, 5-duration, 6-loanstatus, 7-repaymentstatus, 8-rejectionreason
  var loanStatus = recentResponse[0][6]; // assuming the loanStatus the 7th column
  var email = recentResponse[0][1]; // assuming the email the 2nd column
  var name = recentResponse[0][2]; // assuming the name the 3rd column
  var id = recentResponse[0][3]; // assuming the id is the 4th column
  var loanAmount = recentResponse[0][4]; // assuming the loanAmount the 5th column
  var loanDuration = recentResponse[0][5]; // assuming the loanDuration the 6th column
  var reason = recentResponse[0][8]; // assuming the reason the 9th column
  var htmlTemplate;
  if (loanStatus == "Accepted") {
    htmlTemplate = HtmlService.createTemplateFromFile("acceptedEmail");

    // define the html variables
    htmlTemplate.name = name;
    htmlTemplate.loanAmount = loanAmount;
    htmlTemplate.loanDuration = loanDuration;
  } else if (loanStatus == "Rejected") {
    htmlTemplate = HtmlService.createTemplateFromFile("rejectedEmail");

    // define the html variables
    htmlTemplate.name = name;
    htmlTemplate.loanAmount = loanAmount;
    htmlTemplate.loanDuration = loanDuration;
    htmlTemplate.reason = reason;
  }

  // evaluate the template and return an htmloutput object
  var htmlForEmail = htmlTemplate.evaluate().getContent();

  GmailApp.sendEmail(
    email,
    "Loan Application Status",
    "This email contains html",
    { htmlBody: htmlForEmail }
  );

  // Send an email to the club treasurer
  var treasEmail =
    PropertiesService.getScriptProperties().getProperty("TREASURER_EMAIL");
  if (!treasEmail) {
    Logger.log("Error: TREASURER_EMAIL script property not set!");
    return; // Stop execution
  }
  htmlTemplate = HtmlService.createTemplateFromFile("treasurerEmail");
  htmlTemplate.name = name;
  htmlTemplate.email = email;
  htmlTemplate.id = id;
  htmlTemplate.loanAmount = loanAmount;
  htmlTemplate.loanDuration = loanDuration;
  htmlTemplate.loanStatus = loanStatus;
  var htmlForEmail = htmlTemplate.evaluate().getContent();
  GmailApp.sendEmail(
    treasEmail,
    "Nofitication for a loan application",
    "This email contains html",
    { htmlBody: htmlForEmail }
  );
}
