// sheet clear
function clearsheet() {
  var sheet_name = SpreadsheetApp.getActiveSheet().getName();  
  var sheet = SpreadsheetApp.getActive().getSheetByName(sheet_name);
  var range = sheet.getRange("a10:g" + sheet.getMaxRows());
  range.clearContent();
}