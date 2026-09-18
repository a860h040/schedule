/**
 * DeleteSchedule.gs
 *
 * ADD THIS AS A NEW .gs FILE IN THE SAME GOOGLE APPS SCRIPT PROJECT.
 * DO NOT replace or edit the existing Code.gs.
 *
 * Uses the existing scheduler helpers already defined in Code.gs:
 *   requireAdmin_()
 *   getDb_()
 *   APP.SHEETS.SCHEDULE
 *   asDate_()
 *   startOfDay_()
 *   dateKey_()
 *   clean_()
 *   audit_()
 *
 * Deletes Schedule rows ONLY for the selected Start Date through End Date.
 * The Schedule header is preserved.
 * All other sheets and configuration remain unchanged.
 */

function deleteSchedulePeriod(token, startDate, endDate) {
  const ctx = requireAdmin_(token);

  const start = startOfDay_(asDate_(startDate));
  const end = startOfDay_(asDate_(endDate));

  if (!start || !end) {
    throw new Error('Start Date and End Date are required.');
  }

  if (end.getTime() < start.getTime()) {
    throw new Error('End Date cannot be before Start Date.');
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    const ss = getDb_();
    const sh = ss.getSheetByName(APP.SHEETS.SCHEDULE);

    if (!sh) {
      throw new Error('Schedule sheet was not found.');
    }

    if (sh.getLastRow() < 2) {
      return {
        ok: true,
        deletedRows: 0,
        startDate: dateKey_(start),
        endDate: dateKey_(end),
        message: 'The Schedule sheet is already empty.'
      };
    }

    const range = sh.getDataRange();
    const values = range.getValues();
    const headers = values[0].map(clean_);
    const dateIndex = headers.indexOf('Date');

    if (dateIndex < 0) {
      throw new Error('The Schedule sheet does not contain a Date column.');
    }

    const keep = [values[0]];
    let deletedRows = 0;
    const deletedGenerationIds = new Set();

    for (let i = 1; i < values.length; i++) {
      const row = values[i];

      // Ignore fully blank rows.
      if (!row.some(v => clean_(v) !== '')) {
        continue;
      }

      const rowDate = startOfDay_(asDate_(row[dateIndex]));

      if (
        rowDate &&
        rowDate.getTime() >= start.getTime() &&
        rowDate.getTime() <= end.getTime()
      ) {
        deletedRows++;

        const generationIndex = headers.indexOf('Generation ID');
        if (generationIndex >= 0 && clean_(row[generationIndex])) {
          deletedGenerationIds.add(clean_(row[generationIndex]));
        }

        continue;
      }

      keep.push(row);
    }

    if (deletedRows > 0) {
      // Clear schedule content only. Formatting, widths, frozen rows, etc. stay intact.
      sh.getDataRange().clearContent();

      // Restore header and all schedule rows outside the deleted period in one batch.
      sh.getRange(1, 1, keep.length, keep[0].length).setValues(keep);

      SpreadsheetApp.flush();
    }

    audit_(
      'SCHEDULE_DELETED',
      dateKey_(start) + ' through ' + dateKey_(end),
      '',
      '',
      '',
      deletedRows + ' schedule row(s)',
      '',
      'Yes',
      'Administrator confirmed permanent schedule deletion',
      'Deleted selected Schedule period. Generation IDs affected: ' +
        Array.from(deletedGenerationIds).join(', '),
      ctx.username
    );

    return {
      ok: true,
      deletedRows: deletedRows,
      startDate: dateKey_(start),
      endDate: dateKey_(end),
      message: deletedRows
        ? deletedRows + ' schedule row(s) deleted.'
        : 'No Schedule rows were found in the selected date range.'
    };

  } finally {
    lock.releaseLock();
  }
}