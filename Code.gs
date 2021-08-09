/**
 * A special function that runs when the spreadsheet is first
 * opened or reloaded. onOpen() is used to add custom menu
 * items to the spreadsheet.
 */
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Labels')
    .addItem('Open Print Menu', 'showSidebarRemoteDYMO')
    // .addItem('print (local js file)', 'showSidebarLocalDYMO')
    .addSeparator()
    .addItem('Format last response as text', 'writeLastResponse')
    .addItem('Format selected response as text', 'writeActiveResponse')
    .addItem('Format all responses as text', 'writeAllResponses')
    .addToUi();
    
}

/**
 * Takes in values from a range of one row (2D array), returns it formatted as a label string literal. Change the label format here.
 */
function rowValuesToLabelFormat(values){
  Logger.log(values);
  var formattedDate = Utilities.formatDate(values[0][0], "PST", "MM/dd/yyyy");
  Logger.log(formattedDate);
  var firstName = values[0][1];
  var lastName = values[0][2];
  var gender = values[0][3];
  var formattedBirthdate = Utilities.formatDate(values[0][4], "PST", "MM/dd/yyyy");
  var label = `${formattedDate}\n${lastName}, ${firstName}\n${gender}\nDOB: ${formattedBirthdate}`;
  
  return label;
}

/**
 * Format the form repsonse currently selected (cell or row in the "Registration" sheet) as label text.
 */
function getFormattedActiveRow(){
  var registration = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Registration");
  var rowIndex = registration.getActiveCell().getRowIndex();
  var range = registration.getRange(rowIndex, 2, 1, 5);

  var values = range.getValues()
  var label = rowValuesToLabelFormat(values);

  Logger.log(label);
  return label;

}

/**
 * Formats the form response at the bottom of the "Registration" sheet as label text.
 */
function getFormattedLast(){
  var registration = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Registration");
  var lastRowIndex = registration.getLastRow();
  var range = registration.getRange(lastRowIndex, 2, 1, 5);

  var values = range.getValues();
  var label = rowValuesToLabelFormat(values);
  Logger.log(label);

  return label;
}

//Methods below were used as a temporary workaround for printing before sidebar was completed - label text could be copy/pasted into DYMO connect to print.

/**
 * Appends the last patient response, formatted as a blood test label, to the first column in the Labels sheet.
 */
function writeLastResponse(){
  var label = getFormattedLast();
  var labels = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Labels");
  labels.appendRow([label]);
  labels.autoResizeColumn(1);
}

/**
 * Appends the selected (by row or cell) patient response, formatted as a blood test label, to the first column in the Labels sheet.
 */
function writeActiveResponse(){
  var label = getFormattedActiveRow();
  var labels = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Labels");
  labels.appendRow([label]);
  labels.autoResizeColumn(1);
}

/**
 * Formats all the patient responses as blood test labels, appends to the first column of the Labels sheet
 */

function writeAllResponses(){
    var registration = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Registration");
    var labels = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Labels");

    //spreadsheet indexes start at 1
    var values = registration.getRange(2, 2, registration.getLastRow() - 1, 5).getValues();
    
    for(let i = 0; i <values.length; i++){
      // Logger.log(values[0]);
      var label = rowValuesToLabelFormat([values[i]]);
      // Logger.log(label);
      labels.appendRow([label]);
      labels.autoResizeColumn(1);
    }
  }


