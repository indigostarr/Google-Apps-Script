function addColumn() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var updateSheet = [{ name: "Add Column", functionName: "addColumn" }];
  ss.addMenu("Add Column", addColumn);
}

//insert column to existing tab
function addColumn() {
  var source = SpreadsheetApp.getActiveSpreadsheet();
  var newrange = source.getSheets()[1];
  var header = newrange.getRange("K5:L5").getFormulas();
  var form = newrange.getRange("K8:L37").getFormulas();
  var summary = newrange.getRange("K4:L4").getValues();
  var text = newrange.getRange("K6:L7").getValues();
  var sourceFile = DriveApp.getFileById(source.getId());
  var sourceFolder = sourceFile.getParents().next();
  var folderFiles = sourceFolder.getFiles();
  var thisFile;

  // use this to update files in the current folder
  while (folderFiles.hasNext()) {
    thisFile = folderFiles.next();
    if (thisFile.getName() !== sourceFile.getName()) {
      var currentSS = SpreadsheetApp.openById(thisFile.getId());
      var destsheet = currentSS.getSheets()[1];
      destsheet.getRange("H3:K3").breakApart();
      destsheet.getRange("K5:L5").setValues(header);
      destsheet.getRange("K8:L37").setValues(form);
      destsheet.getRange("K4:L4").setValues(summary);
      destsheet.getRange("K6:L7").setValues(text);
      destsheet.getRange("H3:L3").merge();
      destsheet.getRange("K6:K7").merge();
      destsheet.getRange("L6:L7").merge();
      destsheet.getRange("H3:L7").setBorder(true, true, true, true, true, true);
    }
  }
}
