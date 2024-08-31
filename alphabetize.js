function alphabetizeSheets() {
  // Get the active spreadsheet
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  
  // Get all the sheets and their names
  var sheets = spreadsheet.getSheets();
  var sheetNames = sheets.map(sheet => sheet.getName());

  // Sort the names alphabetically
  sheetNames.sort(function(a, b) {
    return a.toLowerCase().localeCompare(b.toLowerCase());
  });

  // Reorder the sheets based on the sorted names
  for (var i = 0; i < sheetNames.length; i++) {
    var sheet = spreadsheet.getSheetByName(sheetNames[i]);
    spreadsheet.setActiveSheet(sheet);
    spreadsheet.moveActiveSheet(i + 1);
  }
  
  // Set the first sheet as the active sheet
  spreadsheet.setActiveSheet(spreadsheet.getSheets()[0]);
}
