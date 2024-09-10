function alphabetizeSheetsWithPinned() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  
  // List of sheet names to pin at the top in the desired order
  var pinnedSheets = ["Pin1", "Pin2", "Pin3"];
  
  // Get all the sheets and split into pinned and others
  var sheetNames = spreadsheet.getSheets().map(sheet => sheet.getName());
  var otherSheetNames = sheetNames.filter(name => !pinnedSheets.includes(name)).sort((a, b) => a.toLowerCase().localeCompare(b.toLowerCase()));

  // Combine pinned sheets with the sorted others
  var sortedSheetNames = pinnedSheets.concat(otherSheetNames);

  // Reorder sheets based on sorted names
  sortedSheetNames.forEach((name, index) => {
    var sheet = spreadsheet.getSheetByName(name);
    spreadsheet.setActiveSheet(sheet);
    spreadsheet.moveActiveSheet(index + 1);
  });

  // Set the first sheet as the active sheet
  spreadsheet.setActiveSheet(spreadsheet.getSheets()[0]);
}
