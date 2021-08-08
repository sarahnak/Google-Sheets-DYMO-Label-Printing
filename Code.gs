/**
 * A special function that runs when the spreadsheet is first
 * opened or reloaded. onOpen() is used to add custom menu
 * items to the spreadsheet.
 */
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Labels')
    .addItem('Print Labels', 'showSidebarRemoteDYMO')
    // .addItem('print (local js file)', 'showSidebarLocalDYMO')
    .addSeparator()
    .addItem('Format last response', 'writeLastPatient')
    .addItem('Format all responses', 'allPatients')
    .addToUi();
    
}

/**
 * Format the patient response currently selected (cell or row) as label text with newline characters.
 */
function getFormattedActiveRow(){
  var registration = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Registration");
  var rowIndex = registration.getActiveCell().getRowIndex();
  Logger.log(rowIndex);

  var range = registration.getRange(rowIndex, 2, 1, 5);
  var values = range.getValues()
  var formattedDate = Utilities.formatDate(values[0][0], "PST", "MM/dd/yyyy");

  var firstName = values[0][1];
  var lastName = values[0][2];
  var gender = values[0][3];
  var formattedBirthdate = Utilities.formatDate(values[0][4], "PST", "MM/dd/yyyy");
  var label = `${formattedDate}\n${lastName}, ${firstName}\n${gender}\nDOB: ${formattedBirthdate}`;

  return label;

}

/**
 * Formats data fields in the same way as getFormattedLast, but with \n characters. Necessary to set as label text.
 */
function getFormattedLastNewline(){
  var registration = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Registration");
  var lastRowIndex = registration.getLastRow();

  var range = registration.getRange(lastRowIndex, 2, 1, 5);
  var values = range.getValues();

  var formattedDate = Utilities.formatDate(values[0][0], "PST", "MM/dd/yyyy");
  Logger.log(formattedDate);

  var firstName = values[0][1];
  // Logger.log(firstName);
  var lastName = values[0][2];
  var gender = values[0][3];
  var formattedBirthdate = Utilities.formatDate(values[0][4], "PST", "MM/dd/yyyy");
  var label = `${formattedDate}\n${lastName}, ${firstName}\n${gender}\nDOB: ${formattedBirthdate}`;
  Logger.log(label);

  return label;
}

/**
 * Write the last patient response, formatted as a blood test label, to the A1 cell in the Labels sheet
 * Anything in the cell will be overwritten.
 * 
 */
function writeLastPatient(){
  var label = getFormattedLast();
  var labels = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Labels");
  labels.getRange("A1").setValue(label);
  labels.autoResizeColumn(1);

}

/**
 * Formats all the patient responses as blood test labels, writes to the first column of the Labels sheet
 */

function allPatients(){
    var registration = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Registration");
    var labels = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Labels");

    //spreadsheet indexes start at 1
    var values = registration.getRange(2, 2, registration.getLastRow() - 1, 5).getValues();

    
    for(var row in values){
      Logger.log(values[row][0]);
      var formattedDate = Utilities.formatDate(values[row][0], "PST", "MM/dd/yyyy");
      var firstName = values[row][1];
      var lastName = values[row][2];
      var gender = values[row][3];
      var formattedBirthdate = Utilities.formatDate(values[row][4], "PST", "MM/dd/yyyy");
      var label = 
`${formattedDate}
${lastName}, ${firstName} 
${gender}
DOB: ${formattedBirthdate}`;
      Logger.log(label);
      var lastRow = labels.getLastRow();
      labels.getRange(lastRow+1, 1).setValue(label);
    }
  }

/**
 * Formats the last patient response text as a blood test label text (JavaScript string literal) and returns it
 *  This was used as the temp fix of copy pasting an address-label formatted text into DYMO connect for printing.
*/  

function getFormattedLast(){
  // Gets the active sheet.
  var registration = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Registration");
  var lastRowIndex = registration.getLastRow();
  // Logger.log(lastRowIndex);

  var range = registration.getRange(lastRowIndex, 2, 1, 5);
  var values = range.getValues();


  // var date = values[0][0];
  // Logger.log(date);
  var formattedDate = Utilities.formatDate(values[0][0], "PST", "MM/dd/yyyy");
  Logger.log(formattedDate);

  var firstName = values[0][1];
  // Logger.log(firstName);
  var lastName = values[0][2];
  var gender = values[0][3];
  var formattedBirthdate = Utilities.formatDate(values[0][4], "PST", "MM/dd/yyyy");
  var label = 
`${formattedDate}
${lastName}, ${firstName} 
${gender}
DOB: ${formattedBirthdate}`;
  Logger.log(label);

  return label;
}

/**trying to refactor
 * do this later
 */
function formatRangeAsLabel(range){
  var values = range.getValues();
  for(var row in values){
      Logger.log(values[row][0]);
      var formattedDate = Utilities.formatDate(values[row][0], "PST", "MM/dd/yyyy");
      var firstName = values[row][1];
      var lastName = values[row][2];
      var gender = values[row][3];
      var formattedBirthdate = Utilities.formatDate(values[row][4], "PST", "MM/dd/yyyy");
      var label = 
`${formattedDate}
${lastName}, ${firstName} 
${gender}
DOB: ${formattedBirthdate}`;
      Logger.log(label);
      var lastRow = labels.getLastRow();
      labels.getRange(lastRow+1, 1).setValue(label);
    }
  
}

