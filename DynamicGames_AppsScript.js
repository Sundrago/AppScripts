/**
 * Converts dialogue data from a sheet into a JSON format.
 */
function sheetToJson() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("기본 대사");
  const data = sheet.getDataRange().getValues();
  let jsonData = {};
  const keys = data[0];

  for (let i = 3; i < data.length; i++) {
    const row = data[i];
    const propertyName = row[1];
    if (!propertyName) continue; // Skip if propertyName is empty or null

    for (let j = 2; j < keys.length; j++) {
      const key = keys[j];
      if (!jsonData[key]) {
        jsonData[key] = {};
      }
      if (['rank', 'preferredGame', 'unpreferredGame'].includes(propertyName)) {
        jsonData[key][propertyName] = row[j];
      } else {
        jsonData[key][propertyName] = row[j] ? `[${key}_${propertyName}]${row[j]}` : "";
      }
    }
  }

  const jsonString = JSON.stringify(jsonData, null, 2);
  showCopyDialog(jsonString); // Show dialog for copying the JSON string
}

/**
 * Converts weather data from a sheet into a JSON format.
 */
function weatherSheetToJson() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("날씨");
  if (!sheet) {
    Logger.log("Sheet not found!");
    return;
  }

  const startRow = 5;
  const startCol = 3;
  const endCol = 13;
  const typeColumn = 1;
  const idxRow = 3;

  const dataRange = sheet.getRange(startRow, startCol, sheet.getLastRow() - startRow + 1, endCol - startCol + 1);
  const typeData = sheet.getRange(startRow, typeColumn, sheet.getLastRow() - startRow + 1).getValues();
  const idxData = sheet.getRange(idxRow, startCol, 1, endCol - startCol + 1).getValues()[0];

  let jsonData = [];

  for (let i = 0; i < typeData.length; i++) {
    for (let j = 0; j < idxData.length; j++) {
      const cellValue = dataRange.getValues()[i][j];
      if (cellValue) { // Check if cell is not empty
        jsonData.push({
          type: typeData[i][0],
          idx: idxData[j],
          data: `[${typeData[i][0]}_weather_${idxData[j]}]${cellValue}`
        });
      }
    }
  }

  const jsonString = JSON.stringify(jsonData, null, 2);
  showCopyDialog(jsonString); // Show dialog for copying the JSON string
}

/**
 * Converts dialogue data into String Table (for Localization)
 */
function iterateSheet() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("기본 대사");
  if (!sheet) {
    Logger.log("Sheet not found!");
    return;
  }

  const startRow = 4;
  const startColumn = 4;
  const dataRange = sheet.getRange(startRow, startColumn, sheet.getLastRow() - startRow + 1, sheet.getLastColumn() - startColumn + 1);
  const keys = sheet.getRange(1, startColumn, 1, dataRange.getNumColumns()).getValues()[0];
  const properties = sheet.getRange(startRow, 2, dataRange.getNumRows(), 1).getValues();

  dataRange.getValues().forEach((row, rowIndex) => {
    row.forEach((value, colIndex) => {
      const property = properties[rowIndex][0];
      const key = keys[colIndex];
      if (value && !['preferredGame', 'unpreferredGame', 'rank'].includes(property)) {
        Logger.log(`Key: ${key}, Property: ${property}, Value: ${value}`);
        addKeyToSheet(`${key}_${property}`, value);
      }
    });
  });
}

/**
 * Converts weather data into String Table (for Localization)
 */
function weatherAddKeys() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("날씨");
  if (!sheet) {
    Logger.log("Sheet not found!");
    return;
  }

  const startRow = 5;
  const startColumn = 3;
  const endColumn = 13;
  const dataRange = sheet.getRange(startRow, startColumn, sheet.getLastRow() - startRow + 1, endColumn - startColumn + 1);
  const keys = sheet.getRange(startRow, 1, dataRange.getNumRows(), 1).getValues();
  const properties = sheet.getRange(3, startColumn, 1, dataRange.getNumColumns()).getValues()[0];

  dataRange.getValues().forEach((row, rowIndex) => {
    row.forEach((value, colIndex) => {
      if (value) { // Check if the cell is not empty
        const key = keys[rowIndex][0];
        const property = properties[colIndex];
        Logger.log(`Composite Key: ${key}_weather_${property}, Value: ${value}`);
        addKeyToSheet(`${key}_weather_${property}`, value);
      }
    });
  });
}

/**
 * Adds a key-value pair to a specified sheet, updating if exists or adding if not.
 * @param {string} key - The key to add or update.
 * @param {string} value - The value associated with the key.
 */
function addKeyToSheet(key, value) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("PetDialogue");
  if (!sheet) {
    Logger.log("Sheet not found!");
    return;
  }

  const keys = sheet.getRange("A:A").getValues();
  let keyFound = false;

  for (let i = 0; i < keys.length; i++) {
    if (keys[i][0] === key) {
      keyFound = true;
      const currentValue = sheet.getRange(i + 1, 2).getValue();
      if (currentValue !== value) {
        sheet.getRange(i + 1, 2).setValue(value);
      }
      break;
    }
  }

  if (!keyFound) {
    const lastRow = sheet.getLastRow();
    sheet.getRange(lastRow + 1, 1).setValue(key);
    sheet.getRange(lastRow + 1, 2).setValue(value);
  } else {
    Logger.log("Key already exists.");
  }
}

/**
 * Shows a dialog with a text area for copying JSON string.
 * @param {string} jsonString - The JSON string to display in the dialog.
 */
function showCopyDialog(jsonString) {
  const htmlOutput = HtmlService.createHtmlOutput(
    `<textarea id="jsonText" style="width:100%;height:200px;">${jsonString}</textarea>
    <button style="width: 60%; height: 40px; margin-left: 20%; margin-right: 20%; text-align: center; font-size: 16px;" onclick="copyToClipboard()">Copy to Clipboard</button>
    <script>function copyToClipboard() {var copyText = document.getElementById("jsonText"); copyText.select(); document.execCommand("copy");}</script>`
  ).setWidth(400).setHeight(250);
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Export JSON');
}

/**
 * Include new menu items on google sheets.
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('선용툴')
    .addItem('[기본대사] Json 생성', 'sheetToJson')
    .addItem('[기본대사] Add KEYS', 'iterateSheet')
    .addSeparator()
    .addItem('[날씨] Json 생성', 'weatherSheetToJson')
    .addItem('[날씨] Add KEYS', 'weatherAddKeys')
    .addSeparator()
    .addItem('Call ChatGPT 3.5Turbo', 'callGPT35')
    .addItem('Call ChatGPT 4.0', 'callGPT40')
    .addToUi();
}
