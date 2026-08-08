/**
 * Google Apps Script for Placement Scraper Sheet Integration
 * Automatically maintains 5 Category Sheets:
 *   1. All Placements (Master)
 *   2. IT & Software
 *   3. Sales & Business Dev
 *   4. Marketing
 *   5. Core & Other
 */

function setupHeaders(sheet) {
  var headers = [
    'Date', 'Company', 'Role', 'Category', 'Salary (Probation)', 'Salary (Post-Probation)', 
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

function getOrInitSheet(ss, name) {
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
  }
  if (sheet.getLastRow() === 0) {
    setupHeaders(sheet);
  }
  return sheet;
}

function doPost(e) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var data = JSON.parse(e.postData.contents);
    
    // Ensure all 5 category sheets exist
    var allSheet = getOrInitSheet(ss, 'All Placements');
    var itSheet = getOrInitSheet(ss, 'IT & Software');
    var salesSheet = getOrInitSheet(ss, 'Sales & Business Dev');
    var mktSheet = getOrInitSheet(ss, 'Marketing');
    var coreSheet = getOrInitSheet(ss, 'Core & Other');

    if (data.rows && data.rows.length > 0) {
      // Group rows by target sheet
      var grouped = {
        'All Placements': [],
        'IT & Software': [],
        'Sales & Business Dev': [],
        'Marketing': [],
        'Core & Other': []
      };

      data.rows.forEach(function(row) {
        var rowArr = [
          row.date || '',
          row.company || '',
          row.role || '',
          row.category || '',
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
        ];
        
        grouped['All Placements'].push(rowArr);
        
        var cat = row.category || 'Core & Other';
        if (grouped[cat]) {
          grouped[cat].push(rowArr);
        } else {
          grouped['Core & Other'].push(rowArr);
        }
      });

      // Ultra-fast bulk setValues insert for each sheet
      Object.keys(grouped).forEach(function(sheetName) {
        var rowsToInsert = grouped[sheetName];
        if (rowsToInsert.length > 0) {
          var targetSheet = getOrInitSheet(ss, sheetName);
          targetSheet.getRange(targetSheet.getLastRow() + 1, 1, rowsToInsert.length, 14).setValues(rowsToInsert);
        }
      });
    }

    // Save last_id in cell Z1 of master sheet
    if (data.last_id) {
      allSheet.getRange("Z1").setValue(data.last_id);
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
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var allSheet = getOrInitSheet(ss, 'All Placements');
  var lastId = allSheet.getRange("Z1").getValue() || 0;
  return ContentService.createTextOutput(JSON.stringify({
    "status": "success",
    "total_rows": Math.max(0, allSheet.getLastRow() - 1),
    "last_id": parseInt(lastId) || 0
  })).setMimeType(ContentService.MimeType.JSON);
}
