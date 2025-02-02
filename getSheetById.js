function getSheetById(sheet_id) {

  //access the workbook
  var wb = SpreadsheetApp.getActiveSpreadsheet();

  // access all the sheets in the workbook
  var sheets = wb.getSheets();

  // loop through all the sheets
  for (i in sheets) {
    // look through all the sheets
    if (sheets[i].getSheetId() == sheet_id) {
      // get the sheet name
      var sheetName = sheets[i].getSheetName();
    }
  }

  // return the native getSheetByName function
  return wb.getSheetByName(sheetName);
}
