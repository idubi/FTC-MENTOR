function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const spreadsheetId = data.targetSheetId || e.parameter.targetSheetId;
    const sheetName = data.sheetName || e.parameter.sheetName || "Sheet1";
    const headers = data.headers;
    var ss = SpreadsheetApp.openById(data.targetSheetId);


    if (data.action === 'init') {
        var oldSheet = ss.getSheetByName(data.oldSheetName);
        if (oldSheet) {
          // 1. Rename the old sheet to archive it
          oldSheet.setName(data.newSheetName);
          
          // 2. Create a new fresh sheet with the original name
          var newSheet = ss.insertSheet(data.oldSheetName, 0);
          
          // 3. Copy headers from the archived sheet to the new one
          var tempHeaders = oldSheet.getRange(1, 1, 1, oldSheet.getLastColumn()).getValues();
          newSheet.getRange(1, 1, 1, tempHeaders[0].length).setValues(tempHeaders);
          
          return ContentService.createTextOutput(JSON.stringify({ status: "success", message: "Sheet archived and reset" }))
            .setMimeType(ContentService.MimeType.JSON);
        }
        return ContentService.createTextOutput(JSON.stringify({ status: "error", message: "Original sheet not found" }))
          .setMimeType(ContentService.MimeType.JSON);
      }
      else {
          let sheet = ss.getSheetByName(sheetName);

          // Create sheet if it doesn't exist
          if (!sheet) {
            sheet = ss.insertSheet(sheetName);
            if (headers && headers.length > 0) {
              sheet.appendRow(headers);
              sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold").setBackground("#f3f3f3");
            }
          }

          // Get current headers from the sheet to ensure correct column mapping
          const currentHeaders = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
          const newRow = currentHeaders.map(header => data[header] !== undefined ? data[header] : "");

          sheet.appendRow(newRow);
          return ContentService.createTextOutput("Success").setMimeType(ContentService.MimeType.TEXT);

      }
      

  } catch (error) {
    return ContentService.createTextOutput("Error: " + error.toString()).setMimeType(ContentService.MimeType.TEXT);
  }
}

function doGet(e) {
  try {
    const spreadsheetId = e.parameter.targetSheetId;
    const sheetName = e.parameter.sheetName || "Sheet1";

    const ss = SpreadsheetApp.openById(spreadsheetId);
    const sheet = ss.getSheetByName(sheetName);

    if (!sheet) {
      return ContentService.createTextOutput(JSON.stringify({ error: "Sheet not found: " + sheetName }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const rows = data.slice(1);

    const jsonArray = rows.map(row => {
      const obj = {};
      headers.forEach((header, index) => {
        obj[header] = row[index];
      });
      return obj;
    });

    return ContentService.createTextOutput(JSON.stringify(jsonArray))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({ error: error.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}


function showRowPopup() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var row = sheet.getActiveRange().getRow();
  var data = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues()[0];
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  
  var html = '<div style="font-family: Arial; padding: 10px;">';
  for (var i = 0; i < headers.length; i++) {
    html += '<b>' + headers[i] + ':</b> ' + data[i] + '<br><br>';
  }
  html += '</div>';
  
  var ui = HtmlService.createHtmlOutput(html)
      .setTitle('Row ' + row + ' Details')
      .setWidth(400)
      .setHeight(500);
  SpreadsheetApp.getUi().showModalDialog(ui, 'Record View');
}