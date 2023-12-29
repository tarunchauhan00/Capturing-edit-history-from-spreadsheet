function onOpen(e){
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('capture')
  .addItem('Capture History', 'manuallyRunScript')
  .addToUi();
}





function onEdit(e) {
  var spreadsheetId = "1vLmLCT4-aYt3udAify1aNrBQvo4YmpL5sB4KfNaMh1I"; // Replace with your spreadsheet ID
  var monitoredSheetName = "source"; // Replace with the name of the sheet you want to monitor
  var logSheetName = "capture history"; // Replace with the name of the sheet where you want to store the results

  // Define the range in the source sheet that you want to monitor
  var monitoredRange = "A2:g"; // Replace with the desired range (e.g., "A2:F" for columns A to F starting from row 2)

  // Check if the edited sheet and spreadsheet match the specified ones
  if (e.source.getId() === spreadsheetId && e.source.getSheetName() === monitoredSheetName) {
    var monitoredSheet = e.source.getSheetByName(monitoredSheetName);
    var logSheet = e.source.getSheetByName(logSheetName);
    var range = e.range;
    var user = Session.getActiveUser().getEmail();
    var timestamp = new Date().toLocaleString();
    var row = range.getRow();
    var col = range.getColumn();
    var oldValue = e.oldValue;
    var newValue = e.value;

    // Check if the edit is within the monitored range
    if (
      range.getRow() >= monitoredSheet.getRange(monitoredRange).getRow() &&
      range.getRow() <= monitoredSheet.getRange(monitoredRange).getLastRow() &&
      range.getColumn() >= monitoredSheet.getRange(monitoredRange).getColumn() &&
      range.getColumn() <= monitoredSheet.getRange(monitoredRange).getLastColumn()
    ) {
      // Get the next available row in the log sheet
      var nextRow = logSheet.getLastRow() + 1;

      // Write information to specific columns in the log sheet
      logSheet.getRange(nextRow, 1).setValue(monitoredSheetName); // Monitored Sheet Name
      logSheet.getRange(nextRow, 2).setValue(user); // User
      logSheet.getRange(nextRow, 3).setValue(timestamp); // Timestamp
      logSheet.getRange(nextRow, 4).setValue(oldValue); // Old Value
      logSheet.getRange(nextRow, 5).setValue(newValue); // New Value
    }
  }
}









function manuallyRunScript() {
  var spreadsheetId = "1vLmLCT4-aYt3udAify1aNrBQvo4YmpL5sB4KfNaMh1I";
  var monitoredSheetName = "source";
  var logSheetName = "capture history";
  var monitoredRange = "A2:F";

  var monitoredSheet = SpreadsheetApp.openById(spreadsheetId).getSheetByName(monitoredSheetName);
  var logSheet = SpreadsheetApp.openById(spreadsheetId).getSheetByName(logSheetName);

  var range = monitoredSheet.getRange(monitoredRange);
  var user = Session.getActiveUser().getEmail();
  var timestamp = new Date().toLocaleString();

  // Get values in the monitored range
  var values = range.getValues();

  // Loop through each cell in the monitored range
  for (var i = 0; i < values.length; i++) {
    for (var j = 0; j < values[i].length; j++) {
      var value = values[i][j];

      // Skip empty cells
      if (value !== "") {
        // Get the next available row in the log sheet
        var nextRow = logSheet.getLastRow() + 1;

        // Write information to specific columns in the log sheet
        logSheet.getRange(nextRow, 1).setValue(monitoredSheetName); // Monitored Sheet Name
        logSheet.getRange(nextRow, 2).setValue(user); // User
        logSheet.getRange(nextRow, 3).setValue(timestamp); // Timestamp
        logSheet.getRange(nextRow, 4).setValue(value); // Value
      }
    }
  }
}
