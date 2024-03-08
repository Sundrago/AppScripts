function sheetToJson() {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("기본 대사");
    var data = sheet.getDataRange().getValues();
    var jsonData = {};

    var keys = data[0]; 

    for (var i = 3; i < data.length; i++) {
        var row = data[i];
        var propertyName = row[1]; 

        if(propertyName == "" || propertyName == null) continue;
        
        for (var j = 2; j < keys.length; j++) { 
            var key = keys[j];
            
            
            if (!jsonData[key]) {
                jsonData[key] = {};
            }
            if(propertyName == "rank" || propertyName == "preferredGame" || propertyName == "unpreferredGame") {
              jsonData[key][propertyName] = row[j];
            }
            else {
              if(row[j] != "") jsonData[key][propertyName] = "[" + key + "_" + propertyName + "]" +row[j];
              else jsonData[key][propertyName] = "";
            }
        }
    }

    var jsonString = JSON.stringify(jsonData, null, 2);
    showCopyDialog(jsonString);
}

function WeatherSheetToJson() {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("날씨");
    if (!sheet) {
        Logger.log("Sheet not found!");
        return;
    }

    var startRow = 5; 
    var startCol = 3;
    var endCol = 13;
    var typeColumn = 1;
    var idxRow = 3;

    var dataRange = sheet.getRange(startRow, startCol, sheet.getLastRow() - startRow + 1, endCol - startCol + 1);
    var typeData = sheet.getRange(startRow, typeColumn, sheet.getLastRow() - startRow + 1, 1).getValues();
    var idxData = sheet.getRange(idxRow, startCol, 1, endCol - startCol + 1).getValues()[0];

    var dataValues = dataRange.getValues();
    var jsonData = [];

    for (var i = 0; i < dataValues.length; i++) {
        for (var j = 0; j < dataValues[i].length; j++) {
            var cellValue = dataValues[i][j];
            if (cellValue !== "" && cellValue !== null) {
                var jsonObject = {
                    type: typeData[i][0],
                    idx: idxData[j],
                    data: "[" + typeData[i][0] + "_weather_" + idxData[j] + "]" +cellValue
                };
                jsonData.push(jsonObject);
            }
        }
    }

    var jsonString = JSON.stringify(jsonData, null, 2);
    Logger.log(jsonString);
    showCopyDialog(jsonString);
}


function iterateSheet() {
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = spreadsheet.getSheetByName("기본 대사");
    if (!sheet) {
        Logger.log("Sheet not found!");
        return;
    }

    var startRow = 4;
    var startColumn = 4;
    var lastRow = sheet.getLastRow();
    var lastColumn = sheet.getLastColumn();

    var keys = sheet.getRange(1, startColumn, 1, lastColumn - startColumn + 1).getValues()[0];
    var properties = sheet.getRange(startRow, 2, lastRow - startRow + 1, 1).getValues();

    var range = sheet.getRange(startRow, startColumn, lastRow - startRow + 1, lastColumn - startColumn + 1);
    var values = range.getValues();

    for (var x = 0; x < values.length; x++) {
        for (var y = 0; y < values[x].length; y++) {
            var value = values[x][y];
            var key = keys[y];
            var property = properties[x][0];

            if(key == "" || property == "" || value == "" || property == "preferredGame" || property == "unpreferredGame" || property == "rank") continue

            Logger.log("Key: " + key + ", Property: " + property + ", Value: " + value);
            AddKeyToSheet(key+"_"+property, value);
        }
    }
}

function WeatherAddKeys() {
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = spreadsheet.getSheetByName("날씨");
    if (!sheet) {
        Logger.log("Sheet not found!");
        return;
    }

    var startRow = 5;
    var startColumn = 3;
    var endColumn = 13;
    var lastRow = sheet.getLastRow();

    var keys = sheet.getRange(startRow, 1, lastRow - startRow + 1, 1).getValues();
    var properties = sheet.getRange(3, startColumn, 1, endColumn - startColumn + 1).getValues()[0];

    var range = sheet.getRange(startRow, startColumn, lastRow - startRow + 1, endColumn - startColumn + 1);
    var values = range.getValues();

    for (var x = 0; x < values.length; x++) {
        for (var y = 0; y < values[x].length; y++) {
            var value = values[x][y];
            var key = keys[x][0];
            var property = properties[y];

            if (value === "" || value === null) continue;

            var compositeKey = key + "_weather_" + property;
            Logger.log("Composite Key: " + compositeKey + ", Value: " + value);
            AddKeyToSheet(compositeKey, value);
        }
    }
}



function AddKeyToSheet(key, value) {
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = spreadsheet.getSheetByName("PetDialogue");
    if (!sheet) {
        Logger.log("Sheet not found!");
        return;
    }

    var keys = sheet.getRange("A:A").getValues();
    var keyFound = false;

    // Check if the key already exists
    for (var i = 0; i < keys.length; i++) {
        if (keys[i][0] === key) {
            keyFound = true;
            if(sheet.getRange(i+1, 2).getValue !== value) {
              sheet.getRange(i+1, 2).setValue(value);
            }
            break;
        }
    }

    if (!keyFound) {
        var lastRow = sheet.getLastRow();
        sheet.getRange(lastRow + 1, 1).setValue(key);
        sheet.getRange(lastRow + 1, 2).setValue(value);
    } else {
        Logger.log("Key already exists.");
    }
}

function CallGPT35(promptText) {

  var apiEndpoint = "https://api.openai.com/v1/chat/completions";
    var apiKey = ""; //API key

    var payload = {
        model: "gpt-3.5-turbo",
        messages: [
            { role: "system", content: "Your task : Translate sentence into multiple languages, follow the instructions bellow. Instructions : 1. Return each sentence into SEMICOLON separated single line. 2. If there is a tag in the sentence, remain <tag>, </tag> on the same word as given according to each languages. 3. Translated result should be IN THE SAME ORDER as given bellow. Chinese (Simplified)(zh-CN)	Chinese (Traditional)(zh-hant)	Dutch(nl)	English(en)	French(fr)	German(de)	Japanese(ja)	Korean(ko)	Portuguese(pt)	Russian(ru)	Spanish(es)" },
            { role: "user", content: promptText }
        ]
    };

    var options = {
        method: "post",
        contentType: "application/json",
        headers: {
            "Authorization": "Bearer " + apiKey
        },
        payload: JSON.stringify(payload),
        muteHttpExceptions: true 
    };

    var response = UrlFetchApp.fetch(apiEndpoint, options);
    var responseJson = JSON.parse(response.getContentText());

    if (response.getResponseCode() !== 200) {
        throw new Error('Error from OpenAI: ' + responseJson.error.message);
    }

    Logger.log(responseJson.choices[0].message.content);
    return responseJson.choices[0].message.content;
}

function CallGPT40(promptText) {

  var apiEndpoint = "https://api.openai.com/v1/chat/completions";
    var apiKey = ""; //API key

    var payload = {
        model: "gpt-4", 
        messages: [
            { role: "system", content: "Your task : Translate sentence into multiple languages, follow the instructions bellow. Instructions : 1. Return each sentence into SEMICOLON separated single line. 2. If there is a tag in the sentence, remain <tag>, </tag> on the same word as given according to each languages. 3. Translated result should be IN THE SAME ORDER as given bellow. Chinese (Simplified)(zh-CN)	Chinese (Traditional)(zh-hant)	Dutch(nl)	English(en)	French(fr)	German(de)	Japanese(ja)	Korean(ko)	Portuguese(pt)	Russian(ru)	Spanish(es)" },
            { role: "user", content: promptText }
        ]
    };

    var options = {
        method: "post",
        contentType: "application/json",
        headers: {
            "Authorization": "Bearer " + apiKey
        },
        payload: JSON.stringify(payload),
        muteHttpExceptions: true
    };

    var response = UrlFetchApp.fetch(apiEndpoint, options);
    var responseJson = JSON.parse(response.getContentText());

    if (response.getResponseCode() !== 200) {
        throw new Error('Error from OpenAI: ' + responseJson.error.message);
    }

    Logger.log(responseJson.choices[0].message.content);
    return responseJson.choices[0].message.content;
}

function showCopyDialog(jsonString) {
    var htmlOutput = HtmlService
        .createHtmlOutput('<textarea id="jsonText" style="width:100%;height:200px;">'
                          + jsonString +
                          '</textarea><button style="width: 60%; height: 40px; margin-left: 20%; margin-right: 20%; text-align: center; font-size: 16px;" onclick="copyToClipboard()">Copy to Clipboard</button>'
                          + '<script>function copyToClipboard() {var copyText = document.getElementById("jsonText"); copyText.select(); document.execCommand("copy");}</script>')
        .setWidth(400)
        .setHeight(250);
    SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Export JSON');
}



function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('선용툴')
      .addItem('[기본대사] Json 생성', 'sheetToJson')
      .addItem('[기본대사] Add KEYS', 'iterateSheet')
      .addSeparator()
      .addItem('[날씨] Json 생성', 'WeatherSheetToJson')
      .addItem('[날씨] Add KEYS', 'WeatherAddKeys')
      .addSeparator()
      .addItem('Call ChatGPT 3.5Turbo', 'CallGPT35')
      .addItem('Call ChatGPT 4.0', 'CallGPT40')
      .addToUi();
}
