// get values from spreadsheet
function NewPTRequest(e) {
  var TIMESTAMP = e.values[0];
  var response = e.values[1];
  var A = e.values[2];
  var B = e.values[3];
  var C = e.values[4];
  var D = e.values[5];
  var E = e.values[6];
  var F = e.values[7];
  var G = e.values[8]; 
  
  
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  var PBP = sheet.getSheets()[3];
  var startRow = 2;                            // First row of data to process
  var numRows = PBP.getLastRow() - 1;        // Number of rows to process
  var lastColumn = PBP.getLastColumn();      // Last column
  var dataRange = PBP.getRange(startRow, 1, numRows, lastColumn) // create an array of data from columns A and B
  var data = dataRange.getValues();            // Fetch values for each row in the rangeFind and replace
  //var employee=form.inputBox(response)
  
  var TO = [];
  var APBP = "";
  var name = "";
  var EEID = "";
  var manager = "";
  var manageremail = "";
  
  for(i=0; i<data.length; ++i){
    var row = data[i];
    
   //var TO = data[i][17];
    //var APBP = data[i][19];
    if (row[0] === response) {// if a match in column B is found, break the loop
      TO.push(data[i][17]);
      APBP = data[i][19];
      //APBP.push(data[i][19]);
      //name.push(data[i][2]);
      name = data[i][2];
      manager = data[i][14];
      manageremail = data[i][15];
      //EEID.push(data[i][1]);}
      EEID = data[i][1];}
  }
  //Logger.log(response);
  
 
  var template = HtmlService.createHtmlOutputFromFile('email');
  var template = template.getContent();
  var template = template.replace("{{name}}",name);
  var template = template.replace("{{EEID}}",EEID);
  var template = template.replace("{{Email}}", response);
  var template = template.replace("{{Mgr}}",manager);
  var template = template.replace("{{MgrE}}",manageremail);
  var template = template.replace("{{A}}",A);
  var template = template.replace("{{B}}",B);
  var template = template.replace("{{C}}",C);
  var template = template.replace("{{D}}",D);
  var template = template.replace("{{E}}",E);
  var template = template.replace("{{F}}",F);
  var template = template.replace("{{G}}",G);
  
 //var cC = APBP[foundindex][19];
  
  //Logger.log(TO);
  //Logger.log(APBP);// show column A
  if (response != "") {  
    GmailApp.createDraft(TO,"New Covid-19 Part Time Schedule Request submitted by: "+name, template, {cc: APBP,cc:'ikelly@lyft.com',htmlBody:template});}
  
  var Letter_Created = "Letter_Created";
  
}

