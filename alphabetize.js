function alphabetizeSheetsWithPinned() {
  // Get the active spreadsheet
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  
  // List of sheet names to pin at the top in the desired order
  var pinnedSheets = ["pin1", "pin2", "pin3"];
  
  // Get all the sheets and their names
  var sheets = spreadsheet.getSheets();
  
  // Create arrays for pinned and other sheets
  var pinnedSheetNames = [];
  var otherSheetNames = [];

  // Sort sheets into pinned and other categories
  sheets.forEach(sheet => {
    var name = sheet.getName();
    if (pinnedSheets.some(pinned => name.toLowerCase().includes(pinned.toLowerCase()))) {
      pinnedSheetNames.push(name);
    } else {
      otherSheetNames.push(name);
    }
  });

  // Sort the non-pinned sheet names alphabetically
  otherSheetNames.sort((a, b) => a.toLowerCase().localeCompare(b.toLowerCase()));

  // Combine the pinned sheets and the alphabetized sheets
  var sortedSheetNames = [...pinnedSheetNames, ...otherSheetNames];

  // Reorder the sheets based on the sorted names
  for (var i = 0; i < sortedSheetNames.length; i++) {
    var sheet = spreadsheet.getSheetByName(sortedSheetNames[i]);
    spreadsheet.setActiveSheet(sheet);
    spreadsheet.moveActiveSheet(i + 1);
  }
  
  // Set the first sheet as the active sheet
  spreadsheet.setActiveSheet(spreadsheet.getSheets()[0]);
}
