// Global variables
var ss = SpreadsheetApp.getActiveSpreadsheet();

// On open
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Convert')
  .addItem('SBV -> Sheet', 'convertSbv2Sheet')
  .addItem('Sheet -> SBV', 'convertSheet2Sbv')
  .addSeparator()
  .addItem('Delete Sheets', 'deleteAllSheets')
  .addToUi();
}

/**
 * SBV -> Spreadsheet
 * Convert text data of SBV file that is pasted in sheet 'sbv' into a spreadsheet table.
 */
function convertSbv2Sheet() {
  const ui = SpreadsheetApp.getUi();
  const sbv = ss.getSheetByName('sbv');  
  var captions = sbv.getRange(2, 1, sbv.getLastRow()-1).getValues();
  var table = [];
  var prevRow = '';
  var tableRow = -1;
  var tableCol = 0;
  const now = new Date();
  const sheetName = 'sbv2sheet' + Utilities.formatDate(now, ss.getSpreadsheetTimeZone(), 'yyyyMMddHHmmss');
  var header = [];
  header[0] = ['time', 'caption'];
  const headerStyle = SpreadsheetApp.newTextStyle().setBold(true).build();
  
  try {
    if (captions.length < 1) {
      throw new Error('No captions available for converting.');
    }
    
    for (var i = 0; i < captions.length; i++) {
      var captionLine = captions[i][0];
      if (captionLine.match(/^\d:\d{2}:\d{2}\.\d{3},\d:\d{2}:\d{2}\.\d{3}$/) !== null) {
        // When the row is a time record
        tableRow += 1;
        tableCol = 0;
        table[tableRow] = [captionLine,''];
        prevRow = 'time';
      } else if (prevRow == 'time') {
        // When the previous row is a time record, i.e., when this row is the first row of text
        tableCol = 1;
        table[tableRow][1] = captionLine;
        prevRow = 'cap';
      } else if (prevRow = 'cap' && captionLine !== '') {
        // When the previous row is text and this row is not a blank, i.e., when this row is the second row of text
        tableRow += 1;
        table[tableRow] = ['', captionLine];
      } else if (captionLine == '') {
        // When the row is blank
        continue;
      }
    }

    // Create new sheet and set contents of array 'table'.
    var newSheet = ss.insertSheet(sheetName, 0);
    var sheetHeader = newSheet.getRange(1, 1, 1, header[0].length)
    .setValues(header)
    .setHorizontalAlignment('center')
    .setTextStyle(headerStyle);
    var sheetData = newSheet.getRange(2, 1, table.length, header[0].length)
    .setValues(table)
    .setVerticalAlignment('top');
    newSheet.setFrozenRows(1);
    
    // Alert
    ui.alert('Complete', 'SBV converted to spreadsheet.', ui.ButtonSet.OK);
  } catch (e) {
    var log = errorMessage(e);
    ui.alert(log);
  }
}

/**
 * Spreadsheet -> SBV
 * Convert time-caption table into a simple SBV file-formatted text that can be used for YouTube.
 */
function convertSheet2Sbv() {
  const ui = SpreadsheetApp.getUi();
  const sheet = ss.getSheetByName('spreadsheet');  
  var table = sheet.getRange(3, 1, sheet.getLastRow()-1, 2).getValues();
  Logger.log(table);
  const now = new Date();
  const sheetName = 'sheet2sbv' + Utilities.formatDate(now, ss.getSpreadsheetTimeZone(), 'yyyyMMddHHmmss');
  var sbvArray = [];
  var sbvText = '';
  
  try {
    for (var i = 0; i < table.length; i++) {
      var caption = table[i];
      for (var j = 0; j < caption.length; j++ ) {
        var captionElem = caption[j];
        if (captionElem == '' || captionElem == null) {
          continue;
        } else {
          sbvArray.push(captionElem);
        }
      }
      sbvArray.push('');
    }
    sbvText = sbvArray.join('\n');

    // Create new sheet and set contents of 'sbvText'.
    var newSheet = ss.insertSheet(sheetName, 0);
    var sheetData = newSheet.getRange(1, 1)
    .setValue(sbvText)
    .setVerticalAlignment('top');
    newSheet.setFrozenRows(2);
    
    // Alert
    ui.alert('Complete', 'Spreadsheet converted into SBV-format text.', ui.ButtonSet.OK);
  } catch (e) {
    var log = errorMessage(e);
    ui.alert(log);
  }
}

/**
 * Function to delete all sheets in this spreadsheet
 */
function deleteAllSheets() {
  const ui = SpreadsheetApp.getUi();
  const exceptionSheet1 = ss.getSheetByName('Form->');
  const exceptionSheet2 = ss.getSheetByName('sbv');
  const exceptionSheet3 = ss.getSheetByName('spreadsheet');
  const exceptionSheets = [exceptionSheet1, exceptionSheet2, exceptionSheet3];
  const confirmMessage = 'Are you sure you want to delete all sheets in this spreadsheet?';
  
  try {
    var confirm = ui.alert(confirmMessage, ui.ButtonSet.OK_CANCEL);
    if (confirm !== ui.Button.OK) {
      throw new Error('Canceled');
    }
    deleteSheets(exceptionSheets);
    ui.alert('All deleted.')
  } catch(e) {
    const log = errorMessage(e);
    ui.alert(log);
  }
}

/****************************************
// Background Functions
/****************************************

/**
 * Standarized error message for this script
 * @param {Object} e Error object returned by try-catch
 * @return {string} message Standarized error message
 */
function errorMessage(e) {
  var message = 'Error : line - ' + e.lineNumber + '\n[' + e.name + '] ' + e.message + '\n' + e.stack;
  return message;
}

/**
 * Delete all sheets in this spreadsheet except for the designated sheet IDs
 * @param {Array} exceptionSheets [Optional] Array of sheet objects to not delete
 */
function deleteSheets(exceptionSheets) {
  exceptionSheets = exceptionSheets || [];
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();
  var deleteSheets = sheets.filter(ds => exceptionSheets.indexOf(ds) == -1);
  Logger.log(deleteSheets);
  if (deleteSheets.length !== 0) {
    for (var i = 0; i < deleteSheets.length; i++) {
      ss.deleteSheet(deleteSheets[i]);
    }
  }
  /*
  for (var i = 0; i < sheets.length; i++) {
    var sheet = sheets[i];
    for (var j = 0; j < exceptionSheetIds.length; j++) {
      var exceptionSheetId = exceptionSheetIds[j];
      if (sheet.getSheetId() == exceptionSheetId) {
        continue;
      } else {
        ss.deleteSheet(sheet);
      }
    }
  }*/ 
}
