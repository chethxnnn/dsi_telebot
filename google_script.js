/**
 * Google Apps Script for Placement Scraper Sheet Integration
 * Paste this in your Google Sheet -> Extensions -> Apps Script
 * Then click Deploy -> New deployment -> Select type: Web App
 *   - Execute as: Me
 *   - Who has access: Anyone
 * Copy the Web App URL into your Vercel / Web App settings.
 */

function doPost(e) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var data = JSON.parse(e.postData.contents);
    
    // 1. Initialize headers if sheet is empty
    if (sheet.getLastRow() === 0) {
      var headers = [
        'Date', 'Company', 'Role', 'Salary (Probation)', 'Salary (Post-Probation)', 
        'Salary (Raw)', 'Location', 'Eligibility', 'Registration Link', 
        'Deadline', 'Needs Review', 'Message Link', 'Raw Message'
      ];
      sheet.appendRow(headers);
      var headerRange = sheet.getRange(1, 1, 1, headers.length);
      headerRange.setFontWeight("bold");
      headerRange.setBackground("#2F5496");
      headerRange.setFontColor("#FFFFFF");
      sheet.setFrozenRows(1);
    }
    
    // 2. Append rows
    if (data.rows && data.rows.length > 0) {
      data.rows.forEach(function(row) {
        sheet.appendRow([
          row.date || '',
          row.company || '',
          row.role || '',
          row.salary_probation || '',
          row.salary_post_probation || '',
          row.salary_raw || '',
          row.location || '',
          row.eligibility || '',
          row.registration_link || '',
          row.deadline || '',
          row.needs_review || '',
          row.message_link || '',
          row.raw_message || ''
        ]);
      });
    }
    
    // 3. Save last_id in cell Z1
    if (data.last_id) {
      sheet.getRange("Z1").setValue(data.last_id);
    }
    
    return ContentService.createTextOutput(JSON.stringify({
      "status": "success",
      "rows_added": data.rows ? data.rows.length : 0,
      "last_id": data.last_id || 0
    })).setMimeType(ContentService.MimeType.JSON);
    
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({
      "status": "error",
      "message": err.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet(e) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var lastId = sheet.getRange("Z1").getValue() || 0;
  return ContentService.createTextOutput(JSON.stringify({
    "status": "success",
    "total_rows": Math.max(0, sheet.getLastRow() - 1),
    "last_id": parseInt(lastId) || 0
  })).setMimeType(ContentService.MimeType.JSON);
}
