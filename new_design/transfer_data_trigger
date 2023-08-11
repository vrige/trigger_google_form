// send incoming forms to the real useful sheet
// notice that in this way we can overcome the constraint of "1 sheet for 1 form"
// otherwise in the real sheet (the destinationSpreadSheet) there would so many useless sheets
function transferData(e) {

  var data = e.range.getValues();

  // Open source and destination spreadsheets by ID (you can take it from the URL)
  var destinationSpreadsheet = SpreadsheetApp.openById('ID_of_the_sheet_which_is_secret');
  
  // Open the respective sheets within the spreadsheets
  var destinationSheet = destinationSpreadsheet.getSheetByName('Incoming Frequencies');

  var lastRow = destinationSheet.getLastRow();
  
  var range = destinationSheet.getRange(lastRow + 1, 1, 1, e.values.length);

  range.setValues(data);
}
