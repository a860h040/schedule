/**
 * NeoChrono PTO read API patch for the existing Google Apps Script deployment.
 *
 * IMPORTANT:
 * 1) Keep your existing CONFIG, PTO_REQUEST_HEADERS, getPtoRequestSpreadsheet_(),
 *    and the rest of the PTO web app code.
 * 2) Replace the existing doGet() with the doGet(e) below.
 * 3) Add the helper functions in this file.
 * 4) Deploy a NEW VERSION of the EXISTING web-app deployment so the /exec URL
 *    stays the same.
 */

function doGet(e) {
  var action = String(
    e && e.parameter && e.parameter.action
      ? e.parameter.action
      : ''
  ).trim();

  if (action === 'neochronoPto') {
    return neoChronoPtoReadApi_(e);
  }

  // Existing web app behavior stays unchanged.
  return HtmlService
    .createHtmlOutputFromFile('Index')
    .setTitle('Pharmacy Schedule')
    .addMetaTag(
      'viewport',
      'width=device-width, initial-scale=1, maximum-scale=1'
    );
}


function neoChronoPtoReadApi_(e) {
  try {
    var ss = getPtoRequestSpreadsheet_();
    var sheetName = String(
      e && e.parameter && e.parameter.sheet
        ? e.parameter.sheet
        : CONFIG.PTO_REQUEST_SHEET
    ).trim();

    // Only allow the configured PTO sheet.
    if (sheetName !== CONFIG.PTO_REQUEST_SHEET) {
      throw new Error('Invalid PTO sheet.');
    }

    var sheet = ss.getSheetByName(CONFIG.PTO_REQUEST_SHEET);

    if (!sheet) {
      throw new Error(
        'Sheet not found: ' + CONFIG.PTO_REQUEST_SHEET
      );
    }

    preparePtoRequestSheet_(sheet);

    var values = sheet.getDataRange().getValues();
    var headers = values.length
      ? values[0].map(function(v){ return String(v || '').trim(); })
      : PTO_REQUEST_HEADERS.slice();

    var timezone = ss.getSpreadsheetTimeZone() ||
      Session.getScriptTimeZone() ||
      'America/New_York';

    var rows = [];

    for (var r = 1; r < values.length; r++) {
      var row = values[r];
      var obj = {};
      var hasValue = false;

      for (var c = 0; c < headers.length; c++) {
        var header = headers[c];
        var value = row[c];

        if (value !== '' && value !== null && value !== undefined) {
          hasValue = true;
        }

        obj[header] = neoChronoPtoApiValue_(
          header,
          value,
          timezone
        );
      }

      if (hasValue) {
        rows.push(obj);
      }
    }

    return neoChronoPtoApiResponse_(
      {
        success: true,
        sheet: CONFIG.PTO_REQUEST_SHEET,
        headers: headers,
        rows: rows,
        rowCount: rows.length,
        generatedAt: new Date().toISOString()
      },
      e && e.parameter ? e.parameter.callback : ''
    );

  } catch (error) {
    return neoChronoPtoApiResponse_(
      {
        success: false,
        sheet: CONFIG.PTO_REQUEST_SHEET,
        message: error && error.message
          ? error.message
          : String(error)
      },
      e && e.parameter ? e.parameter.callback : ''
    );
  }
}


function neoChronoPtoApiValue_(header, value, timezone) {
  if (!(value instanceof Date)) {
    return value === null || value === undefined
      ? ''
      : value;
  }

  var dateOnlyHeaders = {
    'Date': true,
    'Start Date': true,
    'End Date': true,
    'Weekend Saturday': true,
    'Weekend Sunday': true
  };

  if (dateOnlyHeaders[header]) {
    return Utilities.formatDate(
      value,
      timezone,
      'yyyy-MM-dd'
    );
  }

  return value.toISOString();
}


function neoChronoPtoApiResponse_(payload, callback) {
  var json = JSON.stringify(payload);
  var cb = String(callback || '').trim();

  // JSONP lets GitHub Pages read the Apps Script deployment without relying
  // on cross-origin fetch/CORS behavior.
  if (
    cb &&
    /^[A-Za-z_$][A-Za-z0-9_$]*$/.test(cb)
  ) {
    return ContentService
      .createTextOutput(
        cb + '(' + json + ');'
      )
      .setMimeType(
        ContentService.MimeType.JAVASCRIPT
      );
  }

  return ContentService
    .createTextOutput(json)
    .setMimeType(
      ContentService.MimeType.JSON
    );
}
