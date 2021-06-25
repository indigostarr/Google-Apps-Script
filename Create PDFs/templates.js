function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("Generate Letters")
    .addItem("Transfer Letter", "Transfer_Letters")
    .addToUi();
}

function Transfer_Letters() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var startRow = 2; // First row of data to process
  var numRows = sheet.getLastRow() - 1; // Number of rows to process
  var lastColumn = sheet.getLastColumn() + 1; // Last column
  var dataRange = sheet.getRange(startRow, 1, numRows, lastColumn); // Fetch the data range of the active sheet
  var data = dataRange.getValues(); // Fetch values for each row in the range
  var blobFile = doc.getAs("application/pdf");
  var pdf = DriveApp.createFile(blobFile);
  var pdfFolder = DriveApp.getFolderById("1EglHWE_1kZTj5sXxS9YVu6vRQlK1YJRR");

  for (var i = 0; i < data.length; ++i) {
    var row = data[i];

    var ENTRY_DATE = row[0];
    //var TODAY_DATE = new Date();
    var TODAY_DATE = Utilities.formatDate(new Date(), "PST", "MM/dd/yyyy");
    var COMP_CHANGE = row[1];
    var EQUITY_CHANGE = row[2];
    var START_DATE = Utilities.formatDate(
      new Date(row[3]),
      "GMT",
      "MM/dd/yyyy"
    );
    var SALARY_TYPE = row[4];
    var WORKER_NAME = row[5];
    var JOB_TITLE = row[6];
    var LEVEL = row[7];
    var OFFICE = row[8];
    var MANAGER = row[9];
    var MANAGER_TITLE = row[10];
    var MONEY = row[11];
    var COMP = Utilities.formatString(
      "$%d,%02d%1.2f",
      MONEY / 1000,
      (MONEY % 1000) / 10,
      MONEY % 10
    );
    var P_EQUITY_Amt = row[12];
    var P_EQUITY = Utilities.formatString(
      "$%d,%02d%1.2f",
      P_EQUITY_Amt / 1000,
      (P_EQUITY_Amt % 1000) / 10,
      P_EQUITY_Amt % 10
    );
    var T_EQUITY_Amt = row[13];
    var T_EQUITY = Utilities.formatString(
      "$%d,%02d%1.2f",
      T_EQUITY_Amt / 1000,
      (T_EQUITY_Amt % 1000) / 10,
      T_EQUITY_Amt % 10
    );
    var STATUS = row[14];

    // Salary - No Pay Change - No Equity
    if (
      COMP_CHANGE == "No" &&
      EQUITY_CHANGE == "No" &&
      SALARY_TYPE == "Salary" &&
      STATUS == ""
    ) {
      //if condition for type of letter to makeCopy
      //file is the template file, and you get it by ID
      var file = DriveApp.getFileById(
        "1qMicO4eLkNDYDgYF8oEtFRwyapsjDoAhWzkenmavCm0"
      );
      // var finalFolder = DriveApp.getFolderById('1EglHWE_1kZTj5sXxS9YVu6vRQlK1YJRR');
      //We can make a copy of the template, name it, and optionally tell it what folder to live in
      //file.makeCopy will return a Google Drive file object
      var folder = DriveApp.getFolderById("1EUrpy-8nxQNrWfdxHHX9S5lGiEByCaEx");
      var copy = file.makeCopy(WORKER_NAME + " Transfer letter", folder);

      //Once we've got the new file created, we need to open it as a document by using its ID
      var doc = DocumentApp.openById(copy.getId());

      //Since everything we need to change is in the body, we need to get that
      var body = doc.getBody();

      //Then we call all of our replaceText methods
      body.replaceText("{{TODAY_DATE}}", TODAY_DATE);
      body.replaceText("{{WORKER_NAME}}", WORKER_NAME);
      body.replaceText("{{START DATE}}", START_DATE);
      body.replaceText("{{OFFER_JOB_TITLE}}", JOB_TITLE);
      body.replaceText("{{MANAGER_NAME}}", MANAGER);
      body.replaceText("{{JOB_LEVEL}}", LEVEL);
      body.replaceText("{{SALARY}}", COMP);

      return doc;
    }
    // Salary - Pay Change - Equity
    else if (
      COMP_CHANGE == "Yes" &&
      EQUITY_CHANGE == "Yes" &&
      SALARY_TYPE == "Salary" &&
      STATUS == ""
    ) {
      //file is the template file, and you get it by ID
      var file = DriveApp.getFileById(
        "1a3E4l5RvBcq-fttdJ1-8JXDFHmrcUjZ8mJu1rJDnHqg"
      );
      //We can make a copy of the template, name it, and optionally tell it what folder to live in
      //file.makeCopy will return a Google Drive file object
      var folder = DriveApp.getFolderById("1EUrpy-8nxQNrWfdxHHX9S5lGiEByCaEx");
      var copy = file.makeCopy(WORKER_NAME + " Transfer letter", folder);

      //Once we've got the new file created, we need to open it as a document by using its ID
      var doc = DocumentApp.openById(copy.getId());

      //Since everything we need to change is in the body, we need to get that
      var body = doc.getBody();

      //Then we call all of our replaceText methods
      body.replaceText("{{TODAY_DATE}}", TODAY_DATE);
      body.replaceText("{{WORKER_NAME}}", WORKER_NAME);
      body.replaceText("{{SALARY}}", COMP);
      body.replaceText("{{OFFER_JOB_TITLE}}", JOB_TITLE);
      body.replaceText("{{MANAGER_NAME}}", MANAGER);
      body.replaceText("{{JOB_LEVEL}}", LEVEL);
      body.replaceText("{{OFFICE_LOCATION}}", OFFICE);
      body.replaceText("{{START_DATE}}", START_DATE);
      body.replaceText("{{CUSTOM_MANAGER_TITLE}}", MANAGER_TITLE);
      body.replaceText("{{PRORATED_EQUITY}}", P_EQUITY);
      body.replaceText("{{EQUITY_VALUE}}", T_EQUITY);

      return doc;
    }

    // Salary - Pay Change - No Equity
    else if (
      COMP_CHANGE == "Yes" &&
      EQUITY_CHANGE == "No" &&
      SALARY_TYPE == "Salary" &&
      STATUS == ""
    ) {
      //file is the template file, and you get it by ID
      var file = DriveApp.getFileById(
        "1DKhe9XVu8fYxY47JAZ_S1j1AfVMYAgsjQCjSi3n7D3k"
      );

      //We can make a copy of the template, name it, and optionally tell it what folder to live in
      //file.makeCopy will return a Google Drive file object
      var folder = DriveApp.getFolderById("1EUrpy-8nxQNrWfdxHHX9S5lGiEByCaEx");
      var copy = file.makeCopy(WORKER_NAME + " Transfer letter", folder);

      //Once we've got the new file created, we need to open it as a document by using its ID
      var doc = DocumentApp.openById(copy.getId());

      //Since everything we need to change is in the body, we need to get that
      var body = doc.getBody();

      //Then we call all of our replaceText methods
      body.replaceText("{{TODAY_DATE}}", TODAY_DATE);
      body.replaceText("{{WORKER_NAME}}", WORKER_NAME);
      body.replaceText("{{OFFER_JOB_TITLE}}", JOB_TITLE);
      body.replaceText("{{MANAGER_NAME}}", MANAGER);
      body.replaceText("{{JOB_LEVEL}}", LEVEL);
      body.replaceText("{{OFFICE_LOCATION}}", OFFICE);
      body.replaceText("{{START_DATE}}", START_DATE);
      body.replaceText("{{CUSTOM_MANAGER_TITLE}}", MANAGER_TITLE);
      body.replaceText("{{SALARY}}", COMP);

      return doc;
    }

    // Salary - No Pay Change - Equity
    else if (
      COMP_CHANGE == "No" &&
      EQUITY_CHANGE == "Yes" &&
      SALARY_TYPE == "Salary" &&
      STATUS == ""
    ) {
      //file is the template file, and you get it by ID
      var file = DriveApp.getFileById(
        "1-cjVymQkp9XwDvUaj4MMOuLilOFOdOaUTXM7klmA2OI"
      );

      //We can make a copy of the template, name it, and optionally tell it what folder to live in
      //file.makeCopy will return a Google Drive file object
      var folder = DriveApp.getFolderById("1EUrpy-8nxQNrWfdxHHX9S5lGiEByCaEx");
      var copy = file.makeCopy(WORKER_NAME + " Transfer letter", folder);

      //Once we've got the new file created, we need to open it as a document by using its ID
      var doc = DocumentApp.openById(copy.getId());

      //Since everything we need to change is in the body, we need to get that
      var body = doc.getBody();

      //Then we call all of our replaceText methods
      body.replaceText("{{TODAY_DATE}}", TODAY_DATE);
      body.replaceText("{{WORKER_NAME}}", WORKER_NAME);
      body.replaceText("{{OFFER_JOB_TITLE}}", JOB_TITLE);
      body.replaceText("{{MANAGER_NAME}}", MANAGER);
      body.replaceText("{{JOB_LEVEL}}", LEVEL);
      body.replaceText("{{OFFICE_LOCATION}}", OFFICE);
      body.replaceText("{{START_DATE}}", START_DATE);
      body.replaceText("{{CUSTOM_MANAGER_TITLE}}", MANAGER_TITLE);
      body.replaceText("{{SALARY}}", COMP);
      body.replaceText("{{PRORATED_EQUITY}}", P_EQUITY);
      body.replaceText("{{EQUITY_VALUE}}", T_EQUITY);

      return doc;
    }
    // Hourly - Pay Change - No Equity
    else if (
      COMP_CHANGE == "Yes" &&
      EQUITY_CHANGE == "Yes" &&
      SALARY_TYPE == "Hourly" &&
      STATUS == ""
    ) {
      //file is the template file, and you get it by ID
      var file = DriveApp.getFileById(
        "1hEhHtN7Qx1tZU71H3X7xzMLs0iSUO4Su95anr8dVIwQ"
      );

      //We can make a copy of the template, name it, and optionally tell it what folder to live in
      //file.makeCopy will return a Google Drive file object
      var folder = DriveApp.getFolderById("1EUrpy-8nxQNrWfdxHHX9S5lGiEByCaEx");
      var copy = file.makeCopy(WORKER_NAME + " Transfer letter", folder);

      //Once we've got the new file created, we need to open it as a document by using its ID
      var doc = DocumentApp.openById(copy.getId());

      //Since everything we need to change is in the body, we need to get that
      var body = doc.getBody();

      //Then we call all of our replaceText methods
      body.replaceText("{{TODAY_DATE}}", TODAY_DATE);
      body.replaceText("{{WORKER_NAME}}", WORKER_NAME);
      body.replaceText("{{SALARY}}", MONEY);
      body.replaceText("{{OFFER_JOB_TITLE}}", JOB_TITLE);
      body.replaceText("{{MANAGER_NAME}}", MANAGER);
      body.replaceText("{{JOB_LEVEL}}", LEVEL);
      body.replaceText("{{OFFICE_LOCATION}}", OFFICE);
      body.replaceText("{{START_DATE}}", START_DATE);
      body.replaceText("{{CUSTOM_MANAGER_TITLE}}", MANAGER_TITLE);
      body.replaceText("{{PRORATED_EQUITY}}", P_EQUITY);
      body.replaceText("{{EQUITY_VALUE}}", T_EQUITY);

      return doc;
    }

    // Hourly - Pay Change - No Equity
    else if (
      COMP_CHANGE == "Yes" &&
      EQUITY_CHANGE == "No" &&
      SALARY_TYPE == "Hourly" &&
      STATUS == ""
    ) {
      //file is the template file, and you get it by ID
      var file = DriveApp.getFileById(
        "1KpvVUHl1YKesYHQswGPBUeN3m_cTNHVD0nykuYokjtI"
      );

      //We can make a copy of the template, name it, and optionally tell it what folder to live in
      //file.makeCopy will return a Google Drive file object
      var folder = DriveApp.getFolderById("1EUrpy-8nxQNrWfdxHHX9S5lGiEByCaEx");
      var copy = file.makeCopy(WORKER_NAME + " Transfer letter", folder);

      //Once we've got the new file created, we need to open it as a document by using its ID
      var doc = DocumentApp.openById(copy.getId());

      //Since everything we need to change is in the body, we need to get that
      var body = doc.getBody();

      //Then we call all of our replaceText methods
      body.replaceText("{{TODAY_DATE}}", TODAY_DATE);
      body.replaceText("{{WORKER_NAME}}", WORKER_NAME);
      body.replaceText("{{OFFER_JOB_TITLE}}", JOB_TITLE);
      body.replaceText("{{MANAGER_NAME}}", MANAGER);
      body.replaceText("{{JOB_LEVEL}}", LEVEL);
      body.replaceText("{{OFFICE_LOCATION}}", OFFICE);
      body.replaceText("{{START_DATE}}", START_DATE);
      body.replaceText("{{CUSTOM_MANAGER_TITLE}}", MANAGER_TITLE);
      body.replaceText("{{SALARY}}", MONEY);

      return doc;
    }
    // Hourly - No Pay Change - No Equity
    else if (
      COMP_CHANGE == "No" &&
      EQUITY_CHANGE == "No" &&
      SALARY_TYPE == "Hourly" &&
      STATUS == ""
    ) {
      //file is the template file, and you get it by ID
      var file = DriveApp.getFileById(
        "1NndtAGeVutcbR_xSIapnCZGwNEkJ_RhPmr6tkqQ3Po4"
      );

      //We can make a copy of the template, name it, and optionally tell it what folder to live in
      //file.makeCopy will return a Google Drive file object
      var folder = DriveApp.getFolderById("1EUrpy-8nxQNrWfdxHHX9S5lGiEByCaEx");
      var copy = file.makeCopy(WORKER_NAME + " Transfer letter", folder);

      //Once we've got the new file created, we need to open it as a document by using its ID
      var doc = DocumentApp.openById(copy.getId());

      //Since everything we need to change is in the body, we need to get that
      var body = doc.getBody();

      //Then we call all of our replaceText methods
      body.replaceText("{{TODAY_DATE}}", TODAY_DATE);
      body.replaceText("{{WORKER_NAME}}", WORKER_NAME);
      body.replaceText("{{OFFER_JOB_TITLE}}", JOB_TITLE);
      body.replaceText("{{MANAGER_NAME}}", MANAGER);
      body.replaceText("{{JOB_LEVEL}}", LEVEL);
      body.replaceText("{{OFFICE_LOCATION}}", OFFICE);
      body.replaceText("{{START_DATE}}", START_DATE);
      body.replaceText("{{CUSTOM_MANAGER_TITLE}}", MANAGER_TITLE);
      body.replaceText("{{SALARY}}", MONEY);

      return doc;
    }
    //Lastly we save and close the document to persist our changes
    doc.saveAndClose();
    sheet
      .getRange(startRow + i, sheet.getLastColumn())
      .setValue("Letter_Created");

    pdfFolder.addFile(pdf);
  }
}
