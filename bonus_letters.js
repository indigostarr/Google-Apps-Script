function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("Generate Letters")
    .addItem("Generate Letters", "Bonus_Letters")
    .addToUi();
}

var d = new Date();
d.setDate(d.getDate() + 3);
var DATE = d.toLocaleDateString("default", {
  month: "long",
  day: "numeric",
  year: "numeric",
});
var Letter_Created = "Letter_Created";
var folder = DriveApp.getFolderById("1x4011Fknq1dWi7GchHlSf2KadurlKRcT");
var payDayFolder = folder.createFolder(DATE).getId();
var docFolder = DriveApp.getFolderById("1EGeK1f-DJxafkMzHIRuuLxKLldXcMy7I");
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
  //function getFriday(TODAY_DATE) {
  // Copy provided date or use current date
  //date = date? new Date(+date) : new Date();

  // Get the day number
  //var dayNum = date.getDay();

  //Set date to next Monday
  //date.setDate(date.getDate() + (dayNum? 5 - dayNum: 1));
  //return date;
  //}
  // function createFolder() {
  //   var folder = DriveApp.getFolderById('1x4011Fknq1dWi7GchHlSf2KadurlKRcT')
  //   var payDayFolder = folder.createFolder(TODAY_DATE);
  // }

  for (var i = 0; i < data.length; ++i) {
    var row = data[i];

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
    //Logger.log(PLAN);

    if (PLAN == "Spot Bonus" && STATUS != "Letter_Created") {
      //file is the template file, and you get it by ID
      var file = DriveApp.getFileById(
        "1C8tzllCEUB8oyBjwrOFFXizy4IX7ZZuR7of7TtWN3-g"
      );

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

      //Then we call all of our replaceText methods
      body.replaceText("{{Date}}", TODAY_DATE);
      body.replaceText("{{Name}}", EMPLOYEE);
      body.replaceText("{{Amount}}", MONEY);
      body.replaceText("{{PayDate}}", DATE);
      body.replaceText("{{Manager Name}}", MANAGER);
      body.replaceText("{{Curr}}", CURR);

      //Lastly we save and close the document to persist our changes
      doc.saveAndClose();
      sheet.getRange(startRow + i, lastColumn).setValue("Letter_Created");
      var blobFile = doc.getAs("application/pdf");
      var pdf = DriveApp.createFile(blobFile);
      var pdfFolder = DriveApp.getFolderById(payDayFolder);
      pdfFolder.addFile(pdf);
    }
  }
}
