// create a menu that displays when the sheet is opened and can trigger the  below formula
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("Generate Letters")
    // second item is the formula
    .addItem("Generate Letters", "Bonus_Letters")
    .addToUi();
}

// create global variables that can be used
var d = new Date();
d.setDate(d.getDate() + 3);
var DATE = d.toLocaleDateString("default", {
  month: "long",
  day: "numeric",
  year: "numeric",
});

// this is useful to set status after script completes running for this line
var Letter_Created = "Letter_Created";
var folder = DriveApp.getFolderById("google-folder-id-for-doc-template");

// this will be the new folder for PDFs generated to live
var payDayFolder = folder.createFolder(DATE).getId();

// this is the parent folder where we will create a subfolder
var docFolder = DriveApp.getFolderById("parent-folder-for-new-templates");

// create a new folder that is created for the week/time of run
var draftDoc = docFolder.createFolder("Bonus Letter PDFs" + DATE).getId();

function Bonus_Letters() {
  //values is an array of form values
  var sheet = SpreadsheetApp.getActiveSheet();
  var startRow = 2; // First row of data to process
  var numRows = sheet.getLastRow() - 1; // Number of rows to process
  var lastColumn = sheet.getLastColumn() + 1; // Last column
  var dataRange = sheet.getRange(startRow, 1, numRows, lastColumn); // Fetch the data range of the active sheet
  var data = dataRange.getValues(); // Fetch values for each row in the range
  var TODAY_DATE = Utilities.formatDate(new Date(), "PST", "MM/dd/yyyy");

  for (var i = 0; i < data.length; ++i) {
    // create a variable for the row
    var row = data[i];

    // this represents the data in each column row[i][col]
    // apps script is 0 based so the first col is 0
    var KEYNAME = row[1];
    var EMPLOYEE = row[7];
    var EEID = row[0];
    var MONEY = row[14];
    var EFFDATE = row[13];
    var MANAGER = row[12];
    var PLAN = row[17];
    var CURR = row[16];
    var STATUS = row[25];
    // var AMOUNT = Utilities.formatString("$%d,%02d%1.2f", MONEY/1000, MONEY%1000/10,MONEY%10)
    // test outputs with
    //Logger.log(PLAN);

    if (PLAN == "Spot Bonus" && STATUS != "Letter_Created") {
      //file is the template file, and you get it by ID
      var file = DriveApp.getFileById("google doc template file");

      //We can make a copy of the template, name it, and optionally tell it what folder to live in
      //file.makeCopy will return a Google Drive file object
      var folder = DriveApp.getFolderById(draftDoc);
      var copy = file.makeCopy(
        EEID + "_" + KEYNAME + "_SpotBonusLetter",
        folder
      );

      //Once we've got the new file created, we need to open it as a document by using its ID
      var doc = DocumentApp.openById(copy.getId());

      //Since everything we need to change is in the body, we need to get that
      var body = doc.getBody();
      var style = {};
      style[DocumentApp.Attribute.FORMAT] = "Currency";

      //Then we call all of our replaceText methods and set from the data in the sheet
      body.replaceText("{{Date}}", TODAY_DATE);
      body.replaceText("{{Name}}", EMPLOYEE);
      body.replaceText("{{Amount}}", MONEY);
      body.replaceText("{{PayDate}}", DATE);
      body.replaceText("{{Manager Name}}", MANAGER);
      body.replaceText("{{Curr}}", CURR);

      //Lastly we save and close the document to persist our changes
      doc.saveAndClose();

      // add a created text
      sheet.getRange(startRow + i, lastColumn).setValue("Letter_Created");

      // create and save the PDF
      var blobFile = doc.getAs("application/pdf");
      var pdf = DriveApp.createFile(blobFile);
      //
      var pdfFolder = DriveApp.getFolderById(payDayFolder);
      pdfFolder.addFile(pdf);
    }
  }
}
