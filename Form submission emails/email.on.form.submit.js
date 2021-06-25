// get values from spreadsheet
function NewRequest(e) {
  var TIMESTAMP = e.values[0];
  var response = e.values[1];
  var preference = e.values[2];
  var B = e.values[3];

  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  var request = sheet.getSheets()[3];
  var startRow = 2; // First row of data to process
  var numRows = request.getLastRow() - 1; // Number of rows to process
  var lastColumn = request.getLastColumn(); // Last column
  var dataRange = request.getRange(startRow, 1, numRows, lastColumn); // create an array of data from columns A and B
  var data = dataRange.getValues(); // Fetch values for each row in the rangeFind and replace
  //var employee=form.inputBox(response)

  var TO = [];
  var name = "";
  var EEID = "";
  var manager = "";
  var manageremail = "";

  for (i = 0; i < data.length; ++i) {
    var row = data[i];

    if (row[0] === response) {
      // if a match in column B is found, break the loop
      TO.push(data[i][17]);
      HRBP = data[i][19];
      name = data[i][2];
      manager = data[i][14];
      manageremail = data[i][15];
      EEID = data[i][1];
    }
  }

  var template = HtmlService.createHtmlOutputFromFile("email");
  var template = template.getContent();
  var template = template.replace("{{name}}", name);
  var template = template.replace("{{EEID}}", EEID);
  var template = template.replace("{{Email}}", response);
  var template = template.replace("{{Mgr}}", manager);
  var template = template.replace("{{MgrE}}", manageremail);
  var template = template.replace("{{A}}", A);
  var template = template.replace("{{B}}", B);

  if (response != "") {
    GmailApp.createDraft(TO, "New Request submitted by: " + name, template, {
      cc: HRBP,
      cc: "test@gmail.com",
      htmlBody: template,
    });
  }

  var Letter_Created = "Letter_Created";
}
