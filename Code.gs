/**
 * A special function that runs when the spreadsheet is first
 * opened or reloaded. onOpen() is used to add custom menu
 * items to the spreadsheet.
 */
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Labels')
    .addItem('Open Print Menu', 'showSidebar')
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
  //temp fix for gender responses in Mongolian
  var gender;
  if(values[0][3] == "Эр"){
    gender = "Male";
  }
  else if(values[0][3] == "Эм"){
    gender = "Female";
  }
  else{
    gender = values[0][3];
  }
  var formattedBirthdate = Utilities.formatDate(values[0][4], "PST", "MM/dd/yyyy");
  var label = `${formattedDate}\n${lastName}, ${firstName}\n${gender}\nDOB: ${formattedBirthdate}`;
  
  return label;
}


/**
 * Format the form response currently selected (cell or row in any sheet) as blood sample label text.
 */
function getFormattedActiveRow(){
  var registration = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var rowIndex = registration.getActiveCell().getRowIndex();
  var range = registration.getRange(rowIndex, 2, 1, 5);

  var values = range.getValues()
  var label = rowValuesToLabelFormat(values);

  Logger.log(label);
  return label;

}

/**
 * Formats the form response at the bottom of the currently selected sheet as blood sample label text.
 */
function getFormattedLast(){
  var registration = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var lastRowIndex = registration.getLastRow();
  var range = registration.getRange(lastRowIndex, 2, 1, 5);

  var values = range.getValues();
  var label = rowValuesToLabelFormat(values);
  Logger.log(label);

  return label;
}

/**
 * Takes in values from a range of one row (2D array), returns it formatted as a mailing address string literal.
 */
function rowValuesToAddressLabelFormat(values){
  Logger.log(values);
  var firstName = values[0][0];
  var lastName = values[0][1];
  var street = values[0][5];
  var city = values[0][6];
  var state = values[0][7];
  var zipcode = values[0][8]
  var label = `${firstName} ${lastName}\n${street}\n${city}, ${state}  ${zipcode}`;
  
  return label;
}

/**
 * Format the last form response as mailing address label text.
 */

function getAddressFormatLastRow(){
  var registration = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var lastRowIndex = registration.getLastRow();
  var range = registration.getRange(lastRowIndex, 3, 1, 9);

  var values = range.getValues();
  var label = rowValuesToAddressLabelFormat(values);
  Logger.log(label);
  return label;
}

/**
 * Format the form response currently selected (cell or row in the active sheet) as mailing address label text.
 */
function getAddressFormatActiveRow(){
  var registration = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var rowIndex = registration.getActiveCell().getRowIndex();
  var range = registration.getRange(rowIndex, 3, 1, 9);

  var values = range.getValues()
  var label = rowValuesToAddressLabelFormat(values);

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
    var registration = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
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


