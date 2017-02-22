<<<<<<< HEAD
// sheet clear
function clearsheet() {
  var sheet = SpreadsheetApp.getActive().getSheetByName('rowdata');
  //var range = sheet.getRange(10, 1, 10, sheet.getRange(7,5).getValue());
  Logger.log(sheet.getMaxColumns());
  var range = sheet.getRange("a10:g" + sheet.getMaxRows());
  range.clearContent();
}

// This logs the value in the very last cell of this sheet
// var lastRow = sheet.getLastRow();
// var lastColumn = sheet.getLastColumn();
// var lastCell = sheet.getRange(lastRow, lastColumn);
// Logger.log(lastCell.getValue());
// var ss = SpreadsheetApp.getActiveSpreadsheet();
// var sheet = ss.getSheets()[0];
// var range = sheet.getRange("A1:D10");
=======
// sheet clear
function clearsheet() {
  var sheet = SpreadsheetApp.getActive().getSheetByName('rowdata');
  //var range = sheet.getRange(10, 1, 10, sheet.getRange(7,5).getValue());
  Logger.log(sheet.getMaxColumns());
  var range = sheet.getRange("a10:g" + sheet.getMaxRows());
  range.clearContent();
}

// This logs the value in the very last cell of this sheet
// var lastRow = sheet.getLastRow();
// var lastColumn = sheet.getLastColumn();
// var lastCell = sheet.getRange(lastRow, lastColumn);
// Logger.log(lastCell.getValue());
// var ss = SpreadsheetApp.getActiveSpreadsheet();
// var sheet = ss.getSheets()[0];
// var range = sheet.getRange("A1:D10");
>>>>>>> de8581560c4251176a3df6d79a7371ba9ae110f9
// range.clearContent();