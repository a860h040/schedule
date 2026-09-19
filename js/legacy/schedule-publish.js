/**
 * NeoChrono Schedule Publisher
 *
 * Sends the selected schedule period from the private GitHub workbook to the
 * pharmacist-facing Google Apps Script web app. The receiver writes the data
 * into a sheet named exactly "Schedule".
 *
 * Browser transport intentionally uses a normal HTML form targeted to a hidden
 * iframe. This avoids Google Apps Script CORS limitations while still allowing
 * the receiver to post a verified result back to NeoChrono.
 */

const PHARMACIST_SCHEDULE_RECEIVER = Object.freeze({
  WEB_APP_URL: 'https://script.google.com/macros/s/AKfycbyKpDf3Fe6TyvX0vxrQ6_5O18f1DT2ZiEciQbMEC_VObA2WE_COLzXXzwRRMjR17clK/exec',
  SHEET_NAME: 'Schedule',
  TIMEOUT_MS: 45000
});

function neoChronoPublishDateKey_(value) {
  try {
    if (value instanceof Date) return formatDateKey_(startOfDay_(value));
    const d = asDate_(value);
    return d ? formatDateKey_(startOfDay_(d)) : '';
  } catch (ignore) {
    return String(value || '').slice(0,10);
  }
}

function neoChronoPublishCell_(header, value) {
  if (value instanceof Date) {
    if (
      header === 'Date' ||
      header === 'Start Date' ||
      header === 'End Date' ||
      header === 'Weekend Saturday' ||
      header === 'Weekend Sunday' ||
      header === 'Effective Start' ||
      header === 'Effective End'
    ) {
      return formatDateKey_(startOfDay_(value));
    }
    return value.toISOString();
  }

  if (value === undefined || value === null) return '';
  return value;
}

function neoChronoScheduleMatrixForPeriod_(startDate, endDate) {
  const ss = getDb_();
  const sheetName =
    (typeof APP !== 'undefined' && APP.SHEETS && APP.SHEETS.SCHEDULE)
      ? APP.SHEETS.SCHEDULE
      : 'Schedule';

  const sh = ss.getSheetByName(sheetName);
  if (!sh) throw new Error('NeoChrono Schedule sheet was not found.');

  const values = sh.getDataRange().getValues();
  if (!values.length) throw new Error('The NeoChrono Schedule sheet is empty.');

  const headers = (values[0] || []).map(function(v){ return clean_(v); });
  const dateIndex = headers.indexOf('Date');

  if (dateIndex < 0) {
    throw new Error('The Schedule sheet is missing the Date column.');
  }

  const startKey = neoChronoPublishDateKey_(startDate);
  const endKey = neoChronoPublishDateKey_(endDate);

  if (!startKey || !endKey) {
    throw new Error('A valid Start Date and End Date are required.');
  }
  if (endKey < startKey) {
    throw new Error('End Date must be on or after Start Date.');
  }

  const rows = values.slice(1).filter(function(row){
    const dk = neoChronoPublishDateKey_(row[dateIndex]);
    return dk && dk >= startKey && dk <= endKey;
  });

  if (!rows.length) {
    throw new Error(
      'No Schedule rows were found from ' + startKey + ' through ' + endKey + '.'
    );
  }

  const matrix = [
    headers,
    ...rows.map(function(row){
      return headers.map(function(header, i){
        return neoChronoPublishCell_(header, row[i]);
      });
    })
  ];

  return {
    matrix: matrix,
    startDate: startKey,
    endDate: endKey,
    rowCount: rows.length,
    columnCount: headers.length
  };
}

function neoChronoPostScheduleToReceiver_(payload) {
  if (
    typeof document === 'undefined' ||
    typeof window === 'undefined'
  ) {
    throw new Error(
      'Direct pharmacist schedule publishing requires the NeoChrono web app.'
    );
  }

  return new Promise(function(resolve, reject){
    const publishId = String(payload.publishId || '');
    const frameName =
      'neoScheduleReceiver_' +
      publishId.replace(/[^A-Za-z0-9_]/g,'_');

    const iframe = document.createElement('iframe');
    iframe.name = frameName;
    iframe.id = frameName;
    iframe.style.display = 'none';
    iframe.setAttribute('aria-hidden','true');

    const form = document.createElement('form');
    form.method = 'POST';
    form.action = PHARMACIST_SCHEDULE_RECEIVER.WEB_APP_URL;
    form.target = frameName;
    form.style.display = 'none';

    function hidden(name, value) {
      const input = document.createElement('input');
      input.type = 'hidden';
      input.name = name;
      input.value = String(value == null ? '' : value);
      form.appendChild(input);
    }

    hidden('action','replaceSchedule');
    hidden('publishId',publishId);
    hidden('sheet',PHARMACIST_SCHEDULE_RECEIVER.SHEET_NAME);
    hidden('payload',JSON.stringify(payload));

    let done = false;

    function cleanup() {
      window.removeEventListener('message', onMessage);
      try { form.remove(); } catch (ignore) {}
      try { iframe.remove(); } catch (ignore) {}
    }

    function finishError(err) {
      if (done) return;
      done = true;
      clearTimeout(timer);
      cleanup();
      reject(err instanceof Error ? err : new Error(String(err)));
    }

    function finishSuccess(result) {
      if (done) return;
      done = true;
      clearTimeout(timer);
      cleanup();
      resolve(result);
    }

    function onMessage(event) {
      if (event.source !== iframe.contentWindow) return;
      const data = event.data;
      if (!data || data.type !== 'NEOCHRONO_SCHEDULE_PUBLISH_RESULT') return;
      if (String(data.publishId || '') !== publishId) return;

      if (data.ok) {
        finishSuccess(data);
      } else {
        finishError(
          new Error(
            data.message ||
            'The pharmacists Schedule sheet rejected the upload.'
          )
        );
      }
    }

    window.addEventListener('message', onMessage);
    document.body.appendChild(iframe);
    document.body.appendChild(form);

    const timer = setTimeout(function(){
      finishError(
        new Error(
          'The pharmacists Schedule receiver did not confirm the upload within ' +
          Math.round(PHARMACIST_SCHEDULE_RECEIVER.TIMEOUT_MS / 1000) +
          ' seconds. Make sure Code4.gs is deployed in the receiver Apps Script project.'
        )
      );
    }, PHARMACIST_SCHEDULE_RECEIVER.TIMEOUT_MS);

    try {
      form.submit();
    } catch (e) {
      finishError(e);
    }
  });
}

/**
 * Called by the existing "Finalize & Send" button.
 *
 * 1) Finalizes the selected period in NeoChrono.
 * 2) Reads that exact date range from the GitHub-backed Schedule sheet.
 * 3) Sends it to the pharmacist-facing Apps Script receiver.
 * 4) Receiver clears/replaces the sheet named exactly "Schedule".
 */
async function finalizeAndSendToPharmacistsSchedule(token, startDate, endDate) {
  const ctx = requireAdmin_(token);

  const startKey = neoChronoPublishDateKey_(startDate);
  const endKey = neoChronoPublishDateKey_(endDate);

  if (!startKey || !endKey) {
    throw new Error('Start Date and End Date are required.');
  }

  // "Finalize & Send" should actually finalize before publishing.
  if (typeof finalizeSchedule === 'function') {
    finalizeSchedule(token, startKey, endKey, true);
  }

  const schedule = neoChronoScheduleMatrixForPeriod_(startKey, endKey);
  const publishId =
    'PUB_' +
    Date.now() +
    '_' +
    Utilities.getUuid().slice(0,8).toUpperCase();

  const payload = {
    action: 'replaceSchedule',
    publishId: publishId,
    sheet: PHARMACIST_SCHEDULE_RECEIVER.SHEET_NAME,
    startDate: schedule.startDate,
    endDate: schedule.endDate,
    rows: schedule.rowCount,
    columns: schedule.columnCount,
    generatedAt: new Date().toISOString(),
    values: schedule.matrix
  };

  let receipt;
  try {
    receipt = await neoChronoPostScheduleToReceiver_(payload);
  } catch (err) {
    throw new Error(
      'Send failed: ' +
      (err && err.message ? err.message : String(err))
    );
  }

  if (typeof audit_ === 'function') {
    audit_(
      'SEND_SCHEDULE_TO_GOOGLE',
      schedule.startDate + ' through ' + schedule.endDate,
      '', '', '', '', '',
      'No', '',
      'Published ' + schedule.rowCount +
        ' Schedule row(s) to Google receiver web app. Target sheet: ' +
        PHARMACIST_SCHEDULE_RECEIVER.SHEET_NAME +
        '. Publish ID: ' + publishId + '.',
      ctx.username
    );
  }

  return {
    ok: true,
    transferredRows: Number(receipt.transferredRows || schedule.rowCount),
    transferredColumns: Number(receipt.transferredColumns || schedule.columnCount),
    receiverSpreadsheetName:
      receipt.spreadsheetName ||
      'Pharmacists Schedule',
    receiverSheetName:
      receipt.sheetName ||
      PHARMACIST_SCHEDULE_RECEIVER.SHEET_NAME,
    receiverUrl:
      receipt.receiverUrl ||
      PHARMACIST_SCHEDULE_RECEIVER.WEB_APP_URL,
    publishId: publishId,
    validationWasRequired: false,
    message:
      receipt.message ||
      'Schedule was finalized and uploaded to the Google Sheet named Schedule.'
  };
}

/**
 * Browser-side connectivity check. This verifies that the receiver URL is
 * configured; the real write is verified during Finalize & Send through the
 * postMessage receipt returned by Code4.gs.
 */
function testPharmacistScheduleReceiverConnection() {
  return {
    ok: true,
    webAppUrl: PHARMACIST_SCHEDULE_RECEIVER.WEB_APP_URL,
    sheetName: PHARMACIST_SCHEDULE_RECEIVER.SHEET_NAME,
    message:
      'Receiver is configured. Use Finalize & Send to verify the deployed Code4.gs receiver.'
  };
}
