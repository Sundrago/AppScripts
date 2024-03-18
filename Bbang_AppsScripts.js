function generateCSharpCode(startRow = 3, endRow = 1000) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const progress = sheet.getRange("A1");

  startRow = Math.max(startRow, 1);
  endRow = Math.min(endRow, sheet.getLastRow());

  const dataRange = sheet.getRange(startRow, 1, endRow - startRow + 1, sheet.getLastColumn());
  const data = dataRange.getValues();

  let csharpCode = `public void ${sheet.getName()}(string ID)\n{\n    print("EVENT ID ${sheet.getName()} : " + ID);\n    pmtControl.Reset();\n    pmtControl.imageMode = true;\n\n    switch(ID)\n    {\n`;
  let currentCase = "";

  data.forEach((row, index) => {
    updateProgress(progress, index, data.length, "코드 변환");

    if (row[0] !== "" && row[0] !== currentCase) {
      if (currentCase !== "") csharpCode += '            break;\n\n';
      currentCase = row[0];
      csharpCode += `        case "${currentCase}":\n`;
    }

    csharpCode += processCommand(row, startRow + index);
  });

  // Finalize C# code
  csharpCode += '            break;\n    }\n}';
  progress.setValue(""); 
  Logger.log(csharpCode);
  SpreadsheetApp.getUi().alert('코드 변환 완료');
  return csharpCode;
}

function updateProgress(progressCell, currentIndex, total, message) {
  const percentage = Math.round((currentIndex / total) * 100);
  progressCell.setValue(`${message} ${currentIndex} / ${total} (${percentage}%)`);
}

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

function updateLinksInColumnF() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const lastRow = sheet.getLastRow(); // Get the last row with data
  const dataRange = sheet.getRange("E3:E" + lastRow);
  const data = dataRange.getValues(); // Get all values in column E
  const progress = sheet.getRange("A1"); 

  data.forEach((cell, i) => {
    updateProgress(progress, i, data.length, "인덱스 업데이트");
    const cellData = sheet.getRange(i + 3, 5); // Col E
    const cellOutput = sheet.getRange(i + 3, 6); // Col F

    if (cell[0] === "") {
      cellOutput.clearContent(); // Clear the cell if the source is empty
    } else {
      const cellValue = cellData.getValue();
      const idx = findCellAddress(cellValue);
      cellOutput.setFormula(`=HYPERLINK("https://docs.google.com/spreadsheets/d/1vnVZIWT2fYS4G4XNV8qeBGQ6GA3vfaEEbCJuwcgcqJs/edit#gid=0&range=${idx}", "${idx}")`);
    }
  });
  
  progress.setValue(""); // Clear progress indicator
  SpreadsheetApp.getUi().alert('인덱스 업데이트 완료');
}

function findCellAddress(input) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const columnA = sheet.getRange("A:A").getValues(); // Col A

  for (let i = 0; i < columnA.length; i++) {
    if (columnA[i][0] === input) {
      return `A${i + 1}`;
    }
  }
  return "ERROR";
}

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('선용툴')
    .addItem('인덱스링크 업데이트', 'updateLinksInColumnF')
    .addItem('코드작성', 'generateCSharpCodeBegin')
    .addToUi();
}

function generateCSharpCodeBegin() {
  const code = generateCSharpCode(3, 2000);
  outputLongString(code);
}

function outputLongString(longString) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const maxChars = 50000; // Maximum characters per cell
  let startRow = 3; // Starting row
  let column = 10; // Col J

  for (let i = 0; i < longString.length; i += maxChars) {
    const textChunk = longString.substring(i, i + maxChars);
    sheet.getRange(startRow++, column).setValue(textChunk);
  }
}
