/*
  Google Apps Script - ZZ Labs Email Subscription
*/

function doGet() {
  return handleCors(JSON.stringify({ status: 'OK', message: 'ZZ Labs Email API is running' }));
}

function doPost(e) {
  return handlePostRequest(e);
}

function handleCors(jsonResponse) {
  return ContentService
    .createTextOutput(jsonResponse)
    .setMimeType(ContentService.MimeType.JSON)
    .setHeader("Access-Control-Allow-Origin", "*")
    .setHeader("Access-Control-Allow-Methods", "POST, GET, OPTIONS")
    .setHeader("Access-Control-Allow-Headers", "Content-Type");
}

function handlePostRequest(e) {
  try {
    // Use exact sheet name
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Registros ZZLABS.SITE");
    
    if (!sheet) {
      return handleCors(JSON.stringify({ success: false, message: 'Sheet not found: Registros ZZLABS.SITE' }));
    }
    
    var data = JSON.parse(e.postData.contents);
    var email = data.email;
    
    if (!email) {
      return handleCors(JSON.stringify({ success: false, message: 'Email is required' }));
    }
    
    var emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
    if (!emailRegex.test(email)) {
      return handleCors(JSON.stringify({ success: false, message: 'Invalid email format' }));
    }
    
    var lastRow = sheet.getLastRow();
    var lastColumn = sheet.getLastColumn();
    
    // Find "Correos" column
    var headers = sheet.getRange(1, 1, 1, lastColumn).getValues()[0];
    var correoColumnIndex = 1;
    
    for (var i = 0; i < headers.length; i++) {
      var header = headers[i] ? headers[i].toString().toLowerCase() : '';
      if (header.includes('correo') || header.includes('email')) {
        correoColumnIndex = i + 1;
        break;
      }
    }
    
    // Check duplicates
    var emailExists = false;
    if (lastRow > 1) {
      var existingEmails = sheet.getRange(2, correoColumnIndex, lastRow - 1, 1).getValues();
      for (var j = 0; j < existingEmails.length; j++) {
        if (existingEmails[j][0] && existingEmails[j][0].toString().toLowerCase() === email.toLowerCase()) {
          emailExists = true;
          break;
        }
      }
    }
    
    if (emailExists) {
      return handleCors(JSON.stringify({ success: true, message: 'Email already subscribed!' }));
    }
    
    // Add new email with timestamp
    var timestamp = new Date();
    sheet.getRange(lastRow + 1, 1).setValue(timestamp);
    sheet.getRange(lastRow + 1, correoColumnIndex).setValue(email);
    
    return handleCors(JSON.stringify({ success: true, message: 'Email subscribed!', row: lastRow + 1 }));
    
  } catch (error) {
    return handleCors(JSON.stringify({ success: false, message: error.toString() }));
  }
}
