function GenerateCSharpCode(startRow=3, endRow=400) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var progress = sheet.getRange("A1"); 

  startRow = Math.max(startRow, 1); 
  endRow = Math.min(endRow, sheet.getLastRow());

  var data = sheet.getRange(startRow, 1, endRow - startRow + 1, sheet.getLastColumn()).getValues();

  var csharpCode = "public void " +  sheet.getName() + "(string ID)\n{\n";
  csharpCode += '    print("EVENT ID ' + sheet.getName() + ' : " + ID);\n';
  csharpCode += "    pmtComtrol.Reset();\n";
  csharpCode += "    pmtComtrol.imageMode = true;\n\n";
  csharpCode += "    switch(ID)\n    {\n";

  var currentCase = "";

  for (var i = 0; i < data.length; i++) {
    progress.setValue("코드 변환 " + i + " / " + data.length + " (" + Math.round((i / data.length) * 100) + "%)");
    var row = data[i];

    // 새로운 ID 시작
    if (row[0] !== "" && row[0] !== currentCase) {
      if (currentCase !== "") {
        csharpCode += '            break;\n\n';
      }
      currentCase = row[0];
      csharpCode += '        case "' + currentCase + '":\n';
    }

    // 명령어 처리
    var command = row[1];

    switch (command) {
      case "AddString":
        csharpCode += '            pmtComtrol.AddString("' + row[2] + '", "' + row[3] + '");\n';
        break;
      case "AddOption":
        csharpCode += '            pmtComtrol.AddOption("' + row[3] + '", "' + sheet.getName() + '", "' + row[4] + '");\n';
        break;
      case "AddNextAction":
        if(row[3] == "") 
          csharpCode += '            pmtComtrol.AddNextAction("' + sheet.getName() + '", "' + row[4] + '");\n';
        else
          csharpCode += '            pmtComtrol.AddNextAction("' + row[3] + '", "' + row[4] + '");\n';
        break;
      case "StoreOutAction":
        csharpCode += '            storeOutAction = "' + row[3] + '";\n';
        csharpCode += '            pmtComtrol.AddNextAction("main", "store_out");\n';
        break;
      case "AddBbang":
        csharpCode += '            AddBBang(' + row[3] + ');\n';
        break;
      case "__________" :
        break;
      case "__" :
        break;
      case "Custom" :
        var lines = row[3].split("\n"); 
        for (var j = 0; j < lines.length; j++) {
          csharpCode += '            ' + lines[j] + "\n";
        }
        break;
        
      default :
        csharpCode += "            // " +  "No command found on line " + (startRow + i) + " : " + command + "\n";
        Logger.log("No command found on line " + (startRow + i) + " : " + command);
        break;
    }
  }

  // 마지막 case 종료
  csharpCode += '            break;\n';
  csharpCode += "    }\n";
  csharpCode += "}";

  Browser.msgBox('코드변환완료');
  Logger.log(csharpCode);
  progress.setValue("");
  return csharpCode;
}

function UpdateLinksInColumnF() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var lastRow = sheet.getLastRow(); // Get the last row with data
  var columnF = sheet.getRange("E3:E" + lastRow).getValues();

  var progress = sheet.getRange("A1"); 

  for (var i = 0; i < columnF.length; i++) {
    progress.setValue("인덱스 업데이트 " + i + " / " + columnF.length + " (" + Math.round((i / columnF.length) * 100) + "%)");
    var cellData = sheet.getRange(i + 3, 5); // Column E is the 5th column
    var cellOutput = sheet.getRange(i + 3, 6); // Column E is the 5th column

    if (columnF[i][0] === "") {
      // If the cell is empty, remove the link
      cellOutput.setFormula("");
    } else {
      // If the cell is not empty, add a link to sample.net
      var cellValue = cellData.getValue();
      var idx = findCellAddress(cellValue);
      cellOutput.setFormula('=HYPERLINK("https://docs.google.com/spreadsheets/d/1vnVZIWT2fYS4G4XNV8qeBGQ6GA3vfaEEbCJuwcgcqJs/edit#gid=0&range=' + idx + '", "' + idx + '")');
    }
  }
  progress.setValue("");
  Browser.msgBox('인덱스 업데이트 완료');
}

function FindCellAddress(input) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var columnA = sheet.getRange("A:A").getValues(); // Get all values in Column A

  for (var i = 0; i < columnA.length; i++) {
    if (columnA[i][0] === input) { // Check if the cell matches the STRING
      return "A" + (i + 1); // Return the address (1-indexed)
    }
  }
  return "ERROR"; // Return "Not found" if the string is not in Column A
}

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  // Or DocumentApp, SlidesApp or FormApp.
  ui.createMenu('선용툴')
      .addItem('인덱스링크 업데이트', 'updateLinksInColumnF')
      // .addSeparator()
      .addItem('코드작성', 'GenerateCSharpCodeBegin')
      .addToUi();
}

function GenerateCSharpCodeBegin() {
  // SpreadsheetApp.getUi() // Or DocumentApp, SlidesApp or FormApp.
  //    .alert('You clicked the first menu item!');
     outputLongString(GenerateCSharpCode(3,1000));
    // generateCSharpCode(3,447);
}

function OutputLongString(longString) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var maxChars = 50000; // Maximum characters per cell
  var startRow = 3; // Starting row for output
  var column = 10; // Column number to output (e.g., 1 for column A)

  for (var i = 0; i < longString.length; i += maxChars) {
    var textChunk = longString.substring(i, Math.min(i + maxChars, longString.length));
    sheet.getRange(startRow, column++).setValue(textChunk);
  }
}
