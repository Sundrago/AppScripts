/**
 * Converts spreadsheet data into a structured JSON object.
 */
function convertDataIntoJson() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const {lastColumn, lastRow} = {lastColumn: sheet.getLastColumn(), lastRow: sheet.getLastRow()};
  
  const properties = sheet.getRange(1, 2, 1, lastColumn - 1).getValues()[0];
  const dataTypes = sheet.getRange(3, 2, 1, lastColumn - 1).getValues()[0];
  const inclusionCriteria = sheet.getRange(2, 2, 1, lastColumn - 1).getValues()[0];
  const dataRows = sheet.getRange(4, 2, lastRow - 3, lastColumn - 1).getValues();
  const keys = sheet.getRange(4, 1, lastRow - 3, 1).getValues();

  let dictionary = {};

  keys.forEach((keyRow, rowIndex) => {
    const key = keyRow[0];
    if (key !== "" && !isNaN(key)) {
      let obj = {};
      let arrays = {};

      properties.forEach((property, propIndex) => {
        if (inclusionCriteria[propIndex] === "all") {
          const value = dataRows[rowIndex][propIndex];
          const dataType = dataTypes[propIndex];

          if (property.includes('[') && property.includes(']')) { 
            const arrayName = property.split('[')[0]; 
            const arrayIndex = parseInt(property.match(/\[(\d+)\]/)[1], 10);

            arrays[arrayName] = arrays[arrayName] || []; 
            if (dataType === 'int[]' && value !== "") arrays[arrayName][arrayIndex] = parseInt(value, 10);
            else if (value !== "") arrays[arrayName][arrayIndex] = value;
          } else if (value !== "") {
            obj[property] = dataType === 'int' ? parseInt(value, 10) : value;
          }
        }
      });

      Object.assign(obj, arrays);

      if (Object.keys(obj).length > 0) {
        dictionary[parseInt(key, 10)] = obj;
      }
    }
  });

  const jsonString = JSON.stringify(dictionary, null, 2);
  Logger.log(jsonString);
  createJsonFile(jsonString, "exportedData.json"); // Changed to automatically save the JSON and show in dialog
}

/**
 * Displays a dialog with options to copy and download the JSON string.
 * @param {string} jsonString - JSON string to display.
 * @param {string} [url] - URL of the downloadable JSON file.
 */
function showCopyDialog(jsonString, url) {
    const htmlTemplate = HtmlService.createTemplateFromFile('DialogTemplate');
    htmlTemplate.jsonString = jsonString;
    htmlTemplate.url = url;
    const htmlOutput = htmlTemplate.evaluate().setWidth(400).setHeight(250).setTitle('Copy JSON & Download JSON');
    SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Copy JSON & Download JSON');
}

/**
 * Creates a JSON file in Google Drive and opens a dialog for copying and downloading.
 * @param {string} jsonString - JSON content to save.
 * @param {string} fileName - name for the new file.
 */
function createJsonFile(jsonString, fileName) {
  const file = DriveApp.createFile(fileName, jsonString, MimeType.PLAIN_TEXT);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  showCopyDialog(jsonString, file.getUrl());
  Logger.log(file.getUrl());
}

/**
 * Adds a custom menu to the Google Sheets
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('선용툴')
    .addItem('Json 생성', 'convertDataIntoJson')
    .addToUi();
}
