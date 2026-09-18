/**
 * Pharmacy Scheduler - CSV Printable Schedule Export Add-on
 * ----------------------------------------------------------
 * IMPORTANT:
 * This file is an ADD-ON for the existing Pharmacy Scheduler project.
 * Do not replace the scheduler's main Code.gs with only this file.
 * In Apps Script, add a new script file (recommended name: CSVExport.gs)
 * and paste this code into it, OR paste these functions at the bottom of
 * the existing Code.gs.
 *
 * It depends on the scheduler's existing global functions/objects:
 *   APP, requireAdmin_, readTable_, getSettingsMap_, asDate_,
 *   formatDateKey_, addDays_, clean_, yes_, getTz_, weekendGroupForDate_.
 */

/**
 * Returns a CSV laid out like the uploaded printable Excel schedule:
 *   Row 1: SCHEDULE PERIOD
 *   Row 2: A/B/C weekend rotation labels
 *   Row 3: dates
 *   Row 4: day abbreviations
 *   Remaining rows: pharmacist name + one cell per date
 *
 * Cell values:
 *   - scheduled shift code(s), e.g. IM or CC2/E2
 *   - P for approved PTO when not scheduled
 *   - X for an OFF day
 *   - bottom UNFILLED row lists unfilled required shifts for each date
 *
 * CSV cannot preserve Excel colors, borders, merged cells, or column widths,
 * but this keeps the same wide matrix/layout when opened in Excel.
 */
function exportPrintableScheduleCsv(token, startDate, endDate) {
  requireAdmin_(token);

  var start = csvExportDate_(startDate);
  var end = csvExportDate_(endDate);
  if (!start || !end) throw new Error('Select a valid Start Date and End Date before exporting.');
  if (start.getTime() > end.getTime()) throw new Error('Start Date cannot be after End Date.');

  var dates = [];
  for (var d = new Date(start.getTime()); d.getTime() <= end.getTime(); d = addDays_(d, 1)) {
    dates.push(new Date(d.getTime()));
  }
  if (!dates.length) throw new Error('No dates were found in the selected export period.');
  if (dates.length > 100) throw new Error('CSV export is limited to 100 days at a time.');

  var users = readTable_(APP.SHEETS.USERS)
    .filter(function(u) {
      return clean_(u.Role).toLowerCase() !== 'administrator';
    });

  var schedule = readTable_(APP.SHEETS.SCHEDULE).filter(function(r) {
    var rd = csvExportDate_(r.Date);
    return rd && rd.getTime() >= start.getTime() && rd.getTime() <= end.getTime();
  });

  var requests = readTable_(APP.SHEETS.REQUESTS);
  var rawSettings = getSettingsMap_();
  var weekendSettings = csvExportWeekendSettings_(rawSettings);

  // Preserve Users-sheet order, but also include any pharmacist who appears
  // in the saved schedule even if their Users row is currently inactive/missing.
  var userByKey = {};
  var userKeyByUsername = {};
  var userKeyByName = {};
  var orderedUsers = [];
  users.forEach(function(u) {
    var key = csvExportUserKey_(u.Username, u['Pharmacist Name']);
    if (!userByKey[key]) {
      userByKey[key] = u;
      orderedUsers.push(u);
    }
    var uname = String(u.Username === null || u.Username === undefined ? '' : u.Username).trim();
    var nkey = csvExportNameKey_(u['Pharmacist Name']);
    if (uname) userKeyByUsername[uname] = key;
    if (nkey) userKeyByName[nkey] = key;
  });

  schedule.forEach(function(r) {
    if (clean_(r.Status).toUpperCase() === 'UNFILLED') return;
    var name = clean_(r['Assigned Pharmacist']);
    if (!name || name.toUpperCase() === 'UNFILLED') return;
    var key = csvExportResolveUserKey_(r.Username, name, userKeyByUsername, userKeyByName);
    if (!userByKey[key]) {
      var synthetic = {
        Username: r.Username,
        'Pharmacist Name': name,
        Active: 'No',
        Role: 'Pharmacist'
      };
      userByKey[key] = synthetic;
      orderedUsers.push(synthetic);
      var su = String(r.Username === null || r.Username === undefined ? '' : r.Username).trim();
      var sn = csvExportNameKey_(name);
      if (su) userKeyByUsername[su] = key;
      if (sn) userKeyByName[sn] = key;
    }
  });

  // Index filled assignments by pharmacist/date.
  var assignmentsByUserDate = {};
  schedule.forEach(function(r) {
    if (clean_(r.Status).toUpperCase() === 'UNFILLED') return;
    var name = clean_(r['Assigned Pharmacist']);
    if (!name || name.toUpperCase() === 'UNFILLED') return;
    var dateKey = csvExportDateKey_(r.Date);
    if (!dateKey) return;
    var userKey = csvExportResolveUserKey_(r.Username, name, userKeyByUsername, userKeyByName);
    var k = userKey + '|' + dateKey;
    if (!assignmentsByUserDate[k]) assignmentsByUserDate[k] = [];
    assignmentsByUserDate[k].push({
      shift: clean_(r.Shift),
      slot: Number(r.Slot || 1),
      warning: clean_(r.Warning)
    });
  });

  // Index unfilled required slots by date for a final UNFILLED summary row.
  var unfilledByDate = {};
  schedule.forEach(function(r) {
    if (clean_(r.Status).toUpperCase() !== 'UNFILLED') return;
    var dateKey = csvExportDateKey_(r.Date);
    if (!dateKey) return;
    if (!unfilledByDate[dateKey]) unfilledByDate[dateKey] = [];
    var label = clean_(r.Shift);
    if (Number(r.Slot || 1) > 1) label += ' #' + Number(r.Slot || 1);
    unfilledByDate[dateKey].push(label);
  });

  // Index approved PTO. If a pharmacist has no assignment on that date,
  // the printable matrix will show P instead of X.
  var ptoByUserDate = {};
  requests.forEach(function(r) {
    if (clean_(r['Record Type']).toUpperCase() !== 'PTO') return;
    if (clean_(r.Status).toUpperCase() !== 'APPROVED') return;

    var reqStart = csvExportDate_(r.Date || r['Start Date']);
    var reqEnd = csvExportDate_(r.Date || r['End Date'] || r['Start Date']);
    if (!reqStart || !reqEnd) return;

    var userKey = csvExportResolveUserKey_(r.Username, r.Pharmacist, userKeyByUsername, userKeyByName);
    for (var pd = new Date(reqStart.getTime()); pd.getTime() <= reqEnd.getTime(); pd = addDays_(pd, 1)) {
      ptoByUserDate[userKey + '|' + csvExportDateKey_(pd)] = true;
    }
  });

  var rows = [];

  // Excel-template style heading rows.
  rows.push(['SCHEDULE PERIOD'].concat(dates.map(function() { return ''; })));
  rows.push([''].concat(dates.map(function(date) {
    var dow = date.getDay();
    if (dow !== 0 && dow !== 6) return '';
    try {
      return weekendGroupForDate_(date, weekendSettings) || '';
    } catch (e) {
      return '';
    }
  })));
  rows.push([''].concat(dates.map(function(date) {
    return Utilities.formatDate(date, getTz_(), 'M/d');
  })));
  rows.push([''].concat(dates.map(function(date) {
    return Utilities.formatDate(date, getTz_(), 'EEE').toUpperCase();
  })));

  orderedUsers.forEach(function(u) {
    var name = clean_(u['Pharmacist Name']) || clean_(u.Username) || 'Unknown Pharmacist';
    var userKey = csvExportResolveUserKey_(u.Username, name, userKeyByUsername, userKeyByName);
    var row = [name];

    dates.forEach(function(date) {
      var dateKey = csvExportDateKey_(date);
      var list = assignmentsByUserDate[userKey + '|' + dateKey] || [];
      if (list.length) {
        list.sort(function(a, b) {
          if (a.slot !== b.slot) return a.slot - b.slot;
          return a.shift.localeCompare(b.shift);
        });
        var shifts = [];
        list.forEach(function(a) {
          if (a.shift && shifts.indexOf(a.shift) < 0) shifts.push(a.shift);
        });
        row.push(shifts.join('/'));
      } else if (ptoByUserDate[userKey + '|' + dateKey]) {
        row.push('P');
      } else {
        row.push('X');
      }
    });

    rows.push(row);
  });

  // Blank separator + unfilled summary row. This keeps open required shifts
  // visible in the exported printable schedule without assigning them to a person.
  rows.push([''].concat(dates.map(function() { return ''; })));
  rows.push(['UNFILLED'].concat(dates.map(function(date) {
    var items = unfilledByDate[csvExportDateKey_(date)] || [];
    return items.join('/');
  })));

  var csv = rows.map(function(row) {
    return row.map(csvExportEscape_).join(',');
  }).join('\r\n');

  var filename = 'Pharmacist_Schedule_' +
    Utilities.formatDate(start, getTz_(), 'yyyy-MM-dd') + '_to_' +
    Utilities.formatDate(end, getTz_(), 'yyyy-MM-dd') + '.csv';

  return {
    ok: true,
    filename: filename,
    csv: csv,
    startDate: csvExportDateKey_(start),
    endDate: csvExportDateKey_(end),
    pharmacistCount: orderedUsers.length,
    dayCount: dates.length,
    unfilledCount: schedule.filter(function(r) {
      return clean_(r.Status).toUpperCase() === 'UNFILLED';
    }).length
  };
}

/** Build only the weekend settings shape needed by weekendGroupForDate_. */
function csvExportWeekendSettings_(raw) {
  raw = raw || {};
  var rotation = String(raw['Weekend Rotation'] || 'A,B,C')
    .split(',')
    .map(function(x) { return String(x || '').trim().toUpperCase(); })
    .filter(function(x) { return !!x; });
  if (!rotation.length) rotation = ['A', 'B', 'C'];

  var anchorDate = csvExportDate_(raw['Weekend Anchor Date']);
  if (!anchorDate) {
    // A harmless fallback used only for the export header if the setting is blank.
    anchorDate = new Date();
    while (anchorDate.getDay() !== 6) anchorDate = addDays_(anchorDate, 1);
  }

  return {
    weekendRotation: rotation,
    weekendAnchorDate: anchorDate,
    weekendAnchorGroup: String(raw['Weekend Anchor Group'] || 'A').trim().toUpperCase()
  };
}

function csvExportDate_(value) {
  if (value instanceof Date && !isNaN(value.getTime())) {
    return new Date(value.getFullYear(), value.getMonth(), value.getDate());
  }
  if (value === null || value === undefined || value === '') return null;

  var s = String(value).trim();
  var iso = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (iso) {
    var d1 = new Date(Number(iso[1]), Number(iso[2]) - 1, Number(iso[3]));
    return isNaN(d1.getTime()) ? null : d1;
  }

  var d2 = asDate_(value);
  if (!d2 || isNaN(d2.getTime())) return null;
  return new Date(d2.getFullYear(), d2.getMonth(), d2.getDate());
}

function csvExportDateKey_(value) {
  var d = csvExportDate_(value);
  return d ? Utilities.formatDate(d, getTz_(), 'yyyy-MM-dd') : '';
}

function csvExportUserKey_(username, name) {
  var u = String(username === null || username === undefined ? '' : username).trim();
  if (u) return 'U:' + u;
  return 'N:' + csvExportNameKey_(name);
}

function csvExportNameKey_(name) {
  return String(name || '').trim().toLowerCase().replace(/\s+/g, ' ');
}

function csvExportResolveUserKey_(username, name, byUsername, byName) {
  var u = String(username === null || username === undefined ? '' : username).trim();
  if (u && byUsername && byUsername[u]) return byUsername[u];
  var n = csvExportNameKey_(name);
  if (n && byName && byName[n]) return byName[n];
  return csvExportUserKey_(username, name);
}

function csvExportEscape_(value) {
  var s = value === null || value === undefined ? '' : String(value);
  if (/[",\r\n]/.test(s)) s = '"' + s.replace(/"/g, '""') + '"';
  return s;
}