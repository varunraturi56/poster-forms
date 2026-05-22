/**
 * Google Apps Script — Poster Voting & Feedback Backend
 *
 * SETUP:
 * 1. Create a new Google Spreadsheet
 * 2. Create two sheet tabs:
 *    - "Votes" with headers:
 *      Timestamp | Device ID | Voter Name | INFUZE Vote | INFUZE Impressions | INFUZE Comment | XCARCITY Vote | XCARCITY Impressions | XCARCITY Comment | Overall Comment
 *    - "Feedback" with headers:
 *      Timestamp | Device ID | Poster ID | Project | Poster | Impressions | Improvements | Comment | Reviewer Name | Is Update
 * 3. Extensions → Apps Script → paste this file into Code.gs
 * 4. Deploy → New Deployment → Web app
 *    - Execute as: Me
 *    - Who has access: Anyone
 * 5. Copy the deployment URL into SCRIPT_URL in vote.html, feedback.html, and results.html
 */

function doPost(e) {
  try {
    var body = JSON.parse(e.postData.contents);
    var sheetName = body.sheet;
    var data = body.data;

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(sheetName);

    if (!sheet) {
      return ContentService
        .createTextOutput(JSON.stringify({ error: 'Sheet not found: ' + sheetName }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    var row;
    if (sheetName === 'Votes') {
      row = [
        data.timestamp || new Date().toISOString(),
        data.deviceId || '',
        data.voterName || 'Anonymous',
        data.infuzeVote || '',
        data.infuzeImpressions || '',
        data.infuzeComment || '',
        data.xcarcityVote || '',
        data.xcarcityImpressions || '',
        data.xcarcityComment || '',
        data.overallComment || ''
      ];
    } else if (sheetName === 'Feedback') {
      row = [
        data.timestamp || new Date().toISOString(),
        data.deviceId || '',
        data.posterId || '',
        data.project || '',
        data.poster || '',
        data.impressions || '',
        data.improvements || '',
        data.comment || '',
        data.reviewerName || 'Anonymous',
        data.isUpdate ? 'Yes' : 'No'
      ];
    } else {
      return ContentService
        .createTextOutput(JSON.stringify({ error: 'Unknown sheet: ' + sheetName }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    sheet.appendRow(row);

    return ContentService
      .createTextOutput(JSON.stringify({ status: 'ok' }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ error: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet(e) {
  try {
    var sheetName = e.parameter.sheet;

    if (!sheetName) {
      return ContentService
        .createTextOutput(JSON.stringify({ error: 'Missing sheet parameter' }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(sheetName);

    if (!sheet) {
      return ContentService
        .createTextOutput(JSON.stringify({ error: 'Sheet not found: ' + sheetName }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    var data = sheet.getDataRange().getValues();
    if (data.length <= 1) {
      return ContentService
        .createTextOutput(JSON.stringify([]))
        .setMimeType(ContentService.MimeType.JSON);
    }

    var rows = data.slice(1).map(function(row) {
      var obj = {};
      if (sheetName === 'Votes') {
        obj.timestamp = row[0] || '';
        obj.deviceId = row[1] || '';
        obj.voterName = row[2] || '';
        obj.infuzeVote = row[3] || '';
        obj.infuzeImpressions = row[4] || '';
        obj.infuzeComment = row[5] || '';
        obj.xcarcityVote = row[6] || '';
        obj.xcarcityImpressions = row[7] || '';
        obj.xcarcityComment = row[8] || '';
        obj.overallComment = row[9] || '';
      } else if (sheetName === 'Feedback') {
        obj.timestamp = row[0] || '';
        obj.deviceId = row[1] || '';
        obj.posterId = row[2] || '';
        obj.project = row[3] || '';
        obj.poster = row[4] || '';
        obj.impressions = row[5] || '';
        obj.improvements = row[6] || '';
        obj.comment = row[7] || '';
        obj.reviewerName = row[8] || '';
        obj.isUpdate = row[9] || '';
      }
      return obj;
    });

    return ContentService
      .createTextOutput(JSON.stringify(rows))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ error: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
