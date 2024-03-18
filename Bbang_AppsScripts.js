/**
 * Generates C# code from spreadsheet data.
 * @param {number} startRow - starting row for reading data
 * @param {number} endRow - ending row for reading data
 * @returns {string} generated C# code.
 */
function generateCSharpCode(startRow = 3, endRow = 1000) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const progress = sheet.getRange("A1");

  // Adjust the row range to ensure it's within bounds
  startRow = Math.max(startRow, 1);
  endRow = Math.min(endRow, sheet.getLastRow());

  const data = sheet.getRange(startRow, 1, endRow - startRow + 1, sheet.getLastColumn()).getValues();
  let csharpCode = `public void ${sheet.getName()}(string ID)\n{\n    print("EVENT ID ${sheet.getName()} : " + ID);\n    pmtControl.Reset();\n    pmtControl.imageMode = true;\n\n    switch(ID)\n    {\n`;
  
  let currentCase = "";
  data.forEach((row, index) => {
    updateProgress(progress, index + 1, data.length, "코드 변환");

    if (row[0] !== "" && row[0] !== currentCase) {
      if (currentCase !== "") {
        csharpCode += '            break;\n\n';
      }
      currentCase = row[0];
      csharpCode += `        case "${currentCase}":\n`;
    }

    csharpCode += processCommand(row, startRow + index);
  });

  // Finalize C# code
  csharpCode += '            break;\n    }\n}';
  progress.clearContent(); 
  Logger.log(csharpCode);
  SpreadsheetApp.getUi().alert('코드 변환 완료');
  return csharpCode;
}

/**
 * Updates a progress cell with the current operation's status.
 * @param {GoogleAppsScript.Spreadsheet.Range} progressCell - cell to update.
 * @param {number} currentIndex - current index of the task.
 * @param {number} total - total number of tasks.
 * @param {string} message - message for the progress update.
 */
function updateProgress(progressCell, currentIndex, total, message) {
  const percentage = Math.round((currentIndex / total) * 100);
  progressCell.setValue(`${message} ${currentIndex} / ${total} (${percentage}%)`);
}

/**
 * Processes a command and generate a corresponding C# code.
 * @param {Array} row - spreadsheet row containing the command and parameters.
 * @param {number} line number from where the command originates.
 * @returns {string} C# code line for the command.
 */
function processCommand(row, lineNumber) {
  const command = row[1];
  switch (command) {
    case "AddString":
      return `            pmtControl.AddString("${row[2]}", "${row[3]}");\n`;
    case "AddOption":
      return `            pmtControl.AddOption("${row[3]}", "${SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getName()}", "${row[4]}");\n`;
    case "AddNextAction":
      return `            pmtControl.AddNextAction("${row[3] || SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getName()}", "${row[4]}");\n`;
    case "StoreOutAction":
      return `            string storeOutAction = "${row[3]}";\n            pmtControl.AddNextAction("main", "store_out");\n`;
    case "AddBbang":
      return `            AddBBang(${row[3]});\n`;
    case "Custom":
      return row[3].split("\n").map(line => `            ${line}\n`).join('');
    default:
      Logger.log(`No command found on line ${lineNumber} : ${command}`);
      return `            // No command found on line ${lineNumber} : ${command}\n`;
  }
}

/**
 * Updates hyperlinks in column F based on column E.
 */
function updateLinksInColumnF() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const lastRow = sheet.getLastRow();
  const data = sheet.getRange("E3:E" + lastRow).getValues();
  const progress = sheet.getRange("A1");

  data.forEach((cell, index) => {
    updateProgress(progress, index + 1, data.length, "인덱스 업데이트");
    const cellValue = cell[0];
    const cellOutput = sheet.getRange(index + 3, 6); // Column F

    if (cellValue === "") {
      cellOutput.clearContent();
    } else {
      const address = findCellAddress(cellValue);
      cellOutput.setFormula(`=HYPERLINK("https://docs.google.com/spreadsheets/d/1vnVZIWT2fYS4G4XNV8qeBGQ6GA3vfaEEbCJuwcgcqJs/edit#gid=0&range=${address}", "${address}")`);
    }
  });

  progress.clearContent();
  SpreadsheetApp.getUi().alert('인덱스 업데이트 완료');
}

/**
 * Finds cell address of a given input in column A and returns.
 * @param {string} input - The value to find in column A.
 * @returns {string} The cell address if found; "ERROR" otherwise.
 */
function findCellAddress(input) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const columnA = sheet.getRange("A:A").getValues();

  for (let i = 0; i < columnA.length; i++) {
    if (columnA[i][0] === input) {
      return `A${i + 1}`;
    }
  }
  return "ERROR";
}

/**
 * Adds custom menu items to the Google Sheets UI.
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('선용툴')
    .addItem('인덱스링크 업데이트', 'updateLinksInColumnF')
    .addItem('코드작성', 'generateCSharpCodeBegin')
    .addToUi();
}

/**
 * Initiates C# code generation and displays output in the spreadsheet.
 */
function generateCSharpCodeBegin() {
  const code = generateCSharpCode(3, 1000);
  outputLongString(code);
}

/**
 * Outputs a long string into consecutive cells due to character limits.
 * @param {string} longString - The long string to output.
 */
function outputLongString(longString) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  let startRow = 3; 
  const column = 10; 

  for (let i = 0; i < longString.length; i += 50000) { 
    const textChunk = longString.substring(i, i + 50000);
    sheet.getRange(startRow++, column).setValue(textChunk);
  }
}
