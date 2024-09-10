function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Custom Import')
    .addItem('Import Data', 'importData')
    .addToUi();
}

function importData() {
  var { sourceSheet, ui } = getSourceSheet();
  if (!sourceSheet) return;

  var sourceData = sourceSheet.getDataRange().getValues();
  var columnsToKeep = [0, 1, 2];  // Specify the columns you want to keep

  var filteredData = sourceData
    .map(line => columnsToKeep.map(index => line[index]))
    .filter(line => line.some(cell => cell !== ""));

  insertFilteredData(filteredData);
}

function getSourceSheet() {
  var ui = SpreadsheetApp.getUi();
  
  // Prompt for the URL of the Google Sheet
  var urlResponse = ui.prompt("Enter the URL of the Google Sheet", 
                              "Example: https://docs.google.com/spreadsheets/d/your-source-sheet-id/edit#gid=sheetId", 
                              ui.ButtonSet.OK_CANCEL);
  
  if (urlResponse.getSelectedButton() != ui.Button.OK) return { sourceSheet: null, ui };
  var url = urlResponse.getResponseText().trim();
  
  // Extract the spreadsheetId and sheetId (gid) from the URL
  var spreadsheetId = url.match(/spreadsheets\/d\/([a-zA-Z0-9-_]+)/)[1];
  var sheetIdMatch = url.match(/gid=([0-9]+)/);
  var sheetId = sheetIdMatch ? sheetIdMatch[1] : null;
  
  // Open the source spreadsheet
  var sourceSpreadsheet = SpreadsheetApp.openById(spreadsheetId);
  
  // Get the sheet by gid (sheetId)
  var sourceSheet = sourceSpreadsheet.getSheets().filter(sheet => sheet.getSheetId().toString() === sheetId)[0];
  
  if (!sourceSheet) {
    ui.alert("The specified sheet does not exist. Please check the URL and try again.");
    return { sourceSheet: null, ui };
  }

  return { sourceSheet, ui };
}

function insertFilteredData(filteredData) {
  if (filteredData.length > 0) {
    var destinationSheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    destinationSheet.getRange(2, 1, destinationSheet.getLastRow(), destinationSheet.getLastColumn()).clearContent();
    destinationSheet.getRange(2, 1, filteredData.length, filteredData[0].length).setValues(filteredData);
  }
}
