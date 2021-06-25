function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("HM Email")
    .addItem("HM Email", "OnbEmail")
    .addToUi();
}
function OnbEmail() {
  //declare global Variables
  var emailbody = HtmlService.createHtmlOutputFromFile("MgrEmail").getContent(); // pulling in content for HTML
  var Recipients;
  var RecCC;
  var RecBCC;
  var fullDeptList = [];
  var managerList = [];
  var d = new Date();
  d.setDate(d.getDate() + 5);
  var startDate = d.toLocaleDateString("default", {
    month: "long",
    day: "numeric",
    year: "numeric",
  });

  //get data from active sheet
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = SpreadsheetApp.getActiveSheet(); // Use data from the active sheet
  var startRow = 2; // First row of data to process
  var numRows = sheet.getLastRow() - 1; // Number of rows to process
  var lastColumn = sheet.getLastColumn(); // Last column
  var dataRange = sheet.getRange(startRow, 1, numRows, lastColumn); // Fetch the data range of the active sheet
  var data = dataRange.getValues(); // Fetch values for each row in the rangeFind and replace

  //Get recipient list from recipient table
  var ccList = spreadsheet.getSheets()[0];
  var numRowsRec = ccList.getLastRow() - 1; // Number of rows in the recipients table

  // build table in HTML
  var recipientTable =
    "" +
    "<tr><td>Department</td>" +
    "<td>Name</td>" +
    "<td>Email" +
    "</td><td>Title" +
    "</td><td>Position ID" +
    "</td><td>Manager" +
    "</td><td>Location" +
    "</td><td>Onboarding Location" +
    "</td></tr>";

  /*build table in Table Object
  
 */

  //DETERMINE UNIQUE DEPTS
  //get list of all depts
  for (var i = 0; i < data.length; ++i) {
    var row = data[i];
    fullDeptList[i] = row[0];
  }

  var uniqueDeptList = [];

  //remove duplicate values

  for (m = 0; m < fullDeptList.length; ++m) {
    if (fullDeptList.indexOf(fullDeptList[m]) >= m) {
      uniqueDeptList.push(fullDeptList[m]);
    }
  }

  var managerList = [];

  //loops through department list to build email
  for (dCount = 0; dCount < uniqueDeptList.length; ++dCount) {
    for (
      tCount = 0;
      tCount < numRows;
      ++tCount //loops through each row to add mgr email to managerList
    ) {
      var rowA = data[tCount]; // RowA is active row in data (the entire spreadsheet)

      //checks if dept in row matches current department email build
      if (rowA[0] == uniqueDeptList[dCount]) {
        managerList.push(data[tCount][9]);

        //using recipient table HTML
        recipientTable +=
          "<tr><td>" +
          rowA[0] +
          "</td><td>" +
          rowA[2] +
          "</td><td>" +
          rowA[3] +
          "</td><td>" +
          rowA[4] +
          "</td><td>" +
          rowA[5] +
          "</td><td>" +
          rowA[6] +
          "</td><td>" +
          rowA[7] +
          "</td><td>" +
          rowA[8] +
          "</td></tr>";
      }

      recipientTable += "";
    }

    //Dynamic logic for choosing recipients, cc and bcc list

    for (r = 0; r < numRowsRec; r++) {
      if (uniqueDeptList[dCount] == OpsCC[r][0]) {
        if (OpsCC[r][2] == "Managers") {
          Recipients = managerList;
          RecCC = OpsCC[r][3];
          RecBCC = "";
          Logger.log("Regular Run");
        } else if (OpsCC[r][4] == "Managers") {
          Recipients = managerList;
          RecCC = OpsCC[r][3];
          RecBCC = "test@gmail.com";

          Logger.log("Manager BCC");
        } else {
          Logger.log("Recipients undefined");
          Recipients = "";
          RecCC = "";
          RecBCC = "";
        }
      }
    }


    //this script uses the HTML method

    var HTMLBody = HtmlService.createHtmlOutputFromFile("MgrEmail.html");

    var emailBody = HTMLBody.getContent();

    //replace text with table
    emailBody = emailBody.replace(
      "THIS IS WHERE THE TABLE SHOULD GO",
      recipientTable
    );
    emailBody = emailBody.replace("INSERTSTARTDATE", startDate);

    GmailApp.createDraft(
      Recipients,
      uniqueDeptList[dCount] + " New Hires Starting Monday " + startDate,
      emailBody,
      { cc: RecCC, bcc: RecBCC, htmlBody: emailBody }
    );

    managerList = []; //remove values from managerList to start new email thread
    Recipients = "";
    RecCC = "";
    RecBCC = "";

    recipientTable =
      "" +
      "<tr><th>Department</th>" +
      "<th>Name</th>" +
      "<th>Email" +
      "</th><th>Title" +
      "</th><th>Position ID" +
      "</th><th>Manager" +
      "</th><th>Location" +
      "</th><th>Onboarding Location" +
      "</th></tr>";

    tableHeader = true;
  } //end dept loop
} //end script
