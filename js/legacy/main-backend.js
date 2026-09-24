/**
 * Pharmacy Employee Scheduling System
 * Google Apps Script backend
 * Version: 2026-09-06
 * https://script.google.com/macros/s/AKfycbyA16BJntH3Zr5_pZ7E8CG_lLXwdtTVB95tMIizdqk-IaJgtMeDh_U3pNmFn4QpEQvs/exec
 * All scheduling configuration is read from Google Sheets. Run
 * setupPharmacyScheduler() once from the bound spreadsheet.
 */

const APP = Object.freeze({
  NAME: 'Pharmacy Employee Scheduling System',
  VERSION: '2026.09.06.30',
  SESSION_HOURS: 8,
  SHEETS: {
    ADMINS: 'Admin Accounts',
    USERS: 'Users',
    SKILLS: 'Employee Skills',
    SHIFTS: 'Shifts',
    REQUIREMENTS: 'Staffing Requirements',
    REQUESTS: 'PTO / Availability Requests',
    WEEKLY_AVAILABILITY: 'Employee Weekly Availability',
    SCHEDULE: 'Schedule',
    SETTINGS: 'Settings',
    AUDIT: 'Audit Log',
    SESSIONS: 'Sessions'
  },
  HEADERS: {
    ADMINS: ['Admin ID','Admin Name','Username','Temporary Password','Password Hash','Password Salt','Active','Must Change Password','Updated At','Updated By'],
    USERS: [
      'Employee ID','Pharmacist Name','Username','Temporary Password','Password Hash','Password Salt','Active','Role',
      'Schedule Type','Employment Type','Preceptor','Resident','Weekend Group','Weekly Hour Maximum','Target Weekly Hours',
      'Maximum Evening Shifts Per Month','Preferred Start Time','Preferred End Time','Custom Hours Enabled',
      'Custom Hours Mode','Preferred Shift Type','Off-Day Coverage Skills','Weekend Eligible','Evening Eligible','Night Eligible',
      'Resident Covers Regular','Resident Weekends','Resident Evenings','Resident Nights','Rotation Anchor Date',
      'Weekend Rotation Anchor Date','SevenOn Weekly Handling','Notes','Must Change Password','Updated At','Updated By'
    ],
    SKILLS: ['Employee ID','Pharmacist Name','Username','Skill','Active','Updated At','Updated By'],
    SHIFTS: ['Shift','Meaning','Start','End','Hours','Credited Hours','Type','Skill','Weekend','Active','Priority','Updated At','Updated By'],
    REQUIREMENTS: ['Shift','Sunday','Monday','Tuesday','Wednesday','Thursday','Friday','Saturday','Active','Updated At','Updated By'],
    REQUESTS: ['Record Type','Record ID','Pharmacist','Username','Date','Start Date','End Date','Weekend Saturday','Weekend Sunday','Available','Status','Comment','Submitted At','Reviewed By','Reviewed At','Updated At','Updated By'],
    WEEKLY_AVAILABILITY: ['Availability ID','Employee ID','Pharmacist Name','Username','Day','Available','Start Time','End Time','Rule Type','Effective Start','Effective End','Notes','Active','Updated At','Updated By'],
    SCHEDULE: ['Generation ID','Assignment ID','Date','Day','Shift','Slot','Assigned Pharmacist','Username','Hours','Credited Hours','Required Skill','Coverage For Pharmacist','Coverage For Username','Coverage Reason','Shift Type','Weekend','Weekend Group','Holiday','Locked','Manual','Status','Warning','Updated At','Updated By','Finalized At'],
    SETTINGS: ['Setting','Value','Description','Updated At','Updated By'],
    AUDIT: ['Timestamp','User','Action','Date','Pharmacist','Old Shift','New Shift','Old Value','New Value','Override','Override Reason','Details'],
    SESSIONS: ['Token','Username','Created At','Expires At','Last Seen At']
  }
});

let _DB_CACHE = null;
let _TZ_CACHE = null;
let _GENERATED_ASSIGNMENT_SEQ = 0;

function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle(APP.NAME)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/** ----------------------------- SETUP ----------------------------- */

function setupPharmacyScheduler() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('Open the Google Sheet first, then run setupPharmacyScheduler() from its bound Apps Script project.');
  PropertiesService.getScriptProperties().setProperty('PHARMACY_SCHEDULER_SPREADSHEET_ID', ss.getId());

  Object.keys(APP.SHEETS).forEach(k => ensureSheet_(ss, APP.SHEETS[k], APP.HEADERS[k]));
  seedSettings_(ss);
  // v12 hard work-pattern migration: Sunday-Saturday weeks, exactly five
  // workdays for regular pharmacists, and no more than five consecutive days.
  const workPatternResult = syncWorkPatternSettings_(ss, 'SYSTEM');
  // System-managed weekend anchors are created from Weekend Group + the global
  // Weekend Anchor Date/Group settings. This repairs existing pharmacist rows
  // and means administrators do not have to type the 21-day anchor manually.
  const weekendAnchorResult = syncWeekendRotationAnchorDates_(ss, 'SYSTEM');
  seedShifts_(ss);
  seedStaffingRequirements_(ss);
  const adminResult = seedFourAdmins_(ss);
  const pharmacistResult = disablePharmacistLogins_(ss);
  hideLegacyPharmacistCredentialColumns_(ss);
  formatSheets_(ss);
  // v29 repairs legacy PTO rows independently PER REQUESTED OFF DATE RANGE so the
  // first two requests for each calendar date are automatically Approved.
  const ptoAutoApprovalResult = reconcilePtoAutoApprovals_('SYSTEM');

  const result = {
    ok: true,
    message: 'Pharmacy scheduler setup/repair completed. Users includes Off-Day Coverage Skills; those skills are used only when a primary pharmacist has an algorithm-generated OFF day. v29 calculates PTO independently for EACH requested OFF date and prioritizes Start Date / End Date over the legacy Date field. For every calendar date, request #1 and request #2 are automatically Approved first-come, first-served; request #3 and above for that same date remain Pending for administrator review. Existing first/second Pending requests are automatically repaired when the app loads. Evening-to-Day/Morning next-day transitions remain blocked. Pharmacists are scheduling records only; only administrators can sign in.',
    spreadsheetId: ss.getId(),
    adminAccounts: adminResult.accounts,
    pharmacistAccountsConverted: pharmacistResult.count,
    weekendRotationAnchorsUpdated: weekendAnchorResult.updated,
    workPatternSettingsUpdated: workPatternResult.updated,
    ptoAutoApprovalsUpdated: ptoAutoApprovalResult.updated
  };
  Logger.log(JSON.stringify(result));
  adminResult.created.forEach(function(a) {
    Logger.log('ADMIN LOGIN -> Username: ' + a.username + ' | Temporary password: ' + a.temporaryPassword);
  });
  return result;
}


/**
 * Explicitly bind this Apps Script project to the spreadsheet that is open
 * when this function is run from the Apps Script editor. Use this after
 * copying the spreadsheet or Apps Script project so the deployed web app
 * does not keep writing to an older spreadsheet ID stored in Script Properties.
 */
function connectSchedulerToCurrentSpreadsheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('Open the Google Sheet that should store the schedule, then run connectSchedulerToCurrentSpreadsheet() from its bound Apps Script project.');
  PropertiesService.getScriptProperties().setProperty('PHARMACY_SCHEDULER_SPREADSHEET_ID', ss.getId());
  _DB_CACHE = ss;
  _TZ_CACHE = null;
  SpreadsheetApp.flush();
  return {ok:true,spreadsheetId:ss.getId(),spreadsheetName:ss.getName(),spreadsheetUrl:ss.getUrl()};
}

function getDatabaseInfo_() {
  const ss = getDb_();
  return {
    id: String(ss.getId()),
    name: String(ss.getName()),
    url: String(ss.getUrl())
  };
}

function resetInitialAdminPassword() {
  const admins = readTable_(APP.SHEETS.ADMINS);
  const admin = admins.find(r => yesDefault_(r.Active,true));
  if (!admin) throw new Error('No administrator account exists. Run setupPharmacyScheduler() first.');
  const password = generateTemporaryPassword_();
  const salt = Utilities.getUuid();
  updateRowByKey_(APP.SHEETS.ADMINS, 'Admin ID', admin['Admin ID'], {
    'Temporary Password': password,
    'Password Salt': salt,
    'Password Hash': hashPassword_(password, salt),
    'Must Change Password': 'Yes',
    'Updated At': new Date(),
    'Updated By': 'SYSTEM'
  });
  Logger.log('RESET ADMIN LOGIN -> Username: ' + admin.Username + ' | Temporary password: ' + password);
  return {ok:true, username:admin.Username, temporaryPassword:password};
}

function seedFourAdmins_(ss) {
  const sh = ss.getSheetByName(APP.SHEETS.ADMINS);
  const current = readTableFromSheet_(sh);
  const existingById = {};
  current.forEach(r => existingById[clean_(r['Admin ID'])] = r);

  // Preserve the original administrator credentials from the legacy Users sheet for Admin 1 when possible.
  const legacyUsers = readTableFromSheet_(ss.getSheetByName(APP.SHEETS.USERS));
  const legacyAdmin = legacyUsers.find(r => clean_(r.Role).toLowerCase() === 'administrator' && clean_(r.Username));
  const definitions = [
    ['ADMIN-001','Schedule Administrator','admin'],
    ['ADMIN-002','Schedule Administrator 2','admin2'],
    ['ADMIN-003','Schedule Administrator 3','admin3'],
    ['ADMIN-004','Schedule Administrator 4','admin4']
  ];
  const created = [];
  definitions.forEach(function(d, idx) {
    if (existingById[d[0]]) return;
    let username = d[2], temp = '', salt = '', hash = '', mustChange = 'Yes';
    if (idx === 0 && legacyAdmin) {
      username = clean_(legacyAdmin.Username) || d[2];
      temp = clean_(legacyAdmin['Temporary Password']);
      salt = clean_(legacyAdmin['Password Salt']);
      hash = clean_(legacyAdmin['Password Hash']);
      mustChange = yes_(legacyAdmin['Must Change Password']) ? 'Yes' : 'No';
    }
    if (!salt || !hash) {
      temp = generateTemporaryPassword_();
      salt = Utilities.getUuid();
      hash = hashPassword_(temp,salt);
      mustChange = 'Yes';
    }
    appendObjectRow_(APP.SHEETS.ADMINS, {
      'Admin ID':d[0], 'Admin Name':d[1], 'Username':username,
      'Temporary Password':temp, 'Password Hash':hash, 'Password Salt':salt,
      'Active':'Yes', 'Must Change Password':mustChange,
      'Updated At':new Date(), 'Updated By':'SYSTEM'
    });
    created.push({adminId:d[0], username:username, temporaryPassword:temp});
  });
  return {accounts:readTableFromSheet_(sh).map(publicAdmin_), created:created};
}

function disablePharmacistLogins_(ss) {
  const sh = ss.getSheetByName(APP.SHEETS.USERS);
  if (!sh || sh.getLastRow() < 2) return {count:0};
  const rows = readTableFromSheet_(sh);
  let count = 0;
  rows.forEach(function(u) {
    if (clean_(u.Role).toLowerCase() === 'administrator') return;
    const id = clean_(u['Employee ID']);
    if (!id) return;
    const internalKey = clean_(u.Username) || ('EMPLOYEE::' + id);
    const needsUpdate = clean_(u['Temporary Password']) || clean_(u['Password Hash']) || clean_(u['Password Salt']) || !clean_(u.Username) || yes_(u['Must Change Password']);
    if (!needsUpdate) return;
    updateRowByKey_(APP.SHEETS.USERS,'Employee ID',id,{
      'Username':internalKey,
      'Temporary Password':'',
      'Password Hash':'',
      'Password Salt':'',
      'Must Change Password':'No',
      'Role':'Pharmacist',
      'Updated At':new Date(),
      'Updated By':'SYSTEM'
    });
    count++;
  });
  return {count:count};
}

function hideLegacyPharmacistCredentialColumns_(ss) {
  const sh = ss.getSheetByName(APP.SHEETS.USERS);
  if (!sh || sh.getLastColumn() < 1) return;
  const headers = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0].map(clean_);
  ['Username','Temporary Password','Password Hash','Password Salt','Must Change Password'].forEach(function(h) {
    const idx = headers.indexOf(h);
    if (idx >= 0) {
      try { sh.hideColumns(idx+1); } catch(e) {}
    }
  });
}

/** Legacy compatibility: pharmacists no longer have login passwords. */
function backfillTemporaryPasswords_(ss) {
  return {count:0};
}

function resetAllAdminPasswords() {
  const admins = readTable_(APP.SHEETS.ADMINS);
  let count = 0;
  const results = [];
  admins.forEach(function(admin) {
    if (!yesDefault_(admin.Active,true)) return;
    const id = clean_(admin['Admin ID']);
    if (!id) return;
    const password = generateTemporaryPassword_();
    const salt = Utilities.getUuid();
    updateRowByKey_(APP.SHEETS.ADMINS,'Admin ID',id,{
      'Temporary Password':password,
      'Password Salt':salt,
      'Password Hash':hashPassword_(password,salt),
      'Must Change Password':'Yes',
      'Updated At':new Date(),
      'Updated By':'SYSTEM'
    });
    count++;
    results.push({adminId:id,username:admin.Username,temporaryPassword:password});
  });
  Logger.log(JSON.stringify(results));
  return {ok:true,count:count,admins:results};
}

// Kept so an older button/script call does not fail. It now resets ADMIN accounts only.
function resetAllUsersToTemporaryPasswords() {
  return resetAllAdminPasswords();
}

function ensureSheet_(ss, name, headers) {
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  const lastCol = Math.max(sh.getLastColumn(), headers.length);
  let current = lastCol ? sh.getRange(1, 1, 1, lastCol).getValues()[0] : [];
  const existing = current.map(clean_).filter(Boolean);
  if (!existing.length) {
    sh.getRange(1,1,1,headers.length).setValues([headers]);
  } else {
    const merged = existing.slice();
    headers.forEach(h => { if (!merged.includes(h)) merged.push(h); });
    sh.getRange(1,1,1,merged.length).setValues([merged]);
  }
  sh.setFrozenRows(1);
  return sh;
}

function seedSettings_(ss) {
  const defaults = [
    ['Default Schedule Days','56','Default generation length: 8 complete Sunday-Saturday weeks'],
    ['Weekly Hours Limit','40','Default weekly maximum if employee-specific value is blank'],
    ['Maximum Hours Per Schedule Period','320','Legacy hard maximum; kept for compatibility'],
    ['Required Hours Per Schedule Period','320','Each active pharmacist must finish the generated two-month schedule with exactly this many credited hours'],
    ['Resident E2 Shift Code','E2','Residents are scheduled to this evening shift once per week'],
    ['Resident E2 Shifts Per Week','1','Required number of resident E2 shifts per scheduling week'],
    ['SevenOn Block Scheduling','Yes','Keep 7-on/7-off employees on their dedicated shift for their ON block'],
    ['Week Start','Sunday','Beginning of scheduling work week; v13 five-day logic requires Sunday'],
    ['Regular Workdays Per Week','5','Every regular pharmacist must work exactly this many days in each complete Sunday-Saturday week'],
    ['Maximum Consecutive Workdays','5','Hard maximum consecutive workdays for regular pharmacists'],
    ['Block Evening To Morning Transition','Yes','Hard rule: an evening shift cannot be followed on the next calendar day by a Day/Morning shift; the next day must be OFF or another evening shift'],
    ['Required Weekend Assignment','Yes','Weekend Group A/B/C Saturday and Sunday assignments are reserved before weekday scheduling'],
    ['Residents Required Assigned Weekend','Yes','Residents must work both days of their assigned A/B/C weekend before weekday scheduling'],
    ['Preceptor Weekday Evening Allowed','No','Preceptors may not work evening-type shifts Monday through Friday'],
    ['Preceptor Weekend Evening Allowed','Yes','Preceptors may work evening-type shifts on Saturday and Sunday'],
    ['Require Paired ED Coverage','Yes','When EDD and EDE are both required on a date, reserve two different ED-qualified pharmacists before general scheduling'],
    ['ED Day Shift Code','EDD','Day shift used by paired Emergency Department coverage'],
    ['ED Evening Shift Code','EDE','Evening shift used by paired Emergency Department coverage'],
    ['Maximum Evening Shifts','7','Default calendar-month evening shift maximum'],
    ['Weekend Rotation','A,B,C','Comma-separated weekend rotation groups'],
    ['Weekend Anchor Date',formatDateKey_(nextOrSameSaturday_(new Date())),'Anchor Saturday for weekend rotation'],
    ['Weekend Anchor Group','A','Weekend group assigned to the anchor Saturday'],
    ['Allow Weekend Fallback','No','If Yes, other weekend groups may be used with a warning'],
    ['Allow Admin Rule Override','Yes','If Yes, administrators may save manual rule violations with a reason'],
    ['Default Regular Schedule Days','5','Expected workdays per week for regular employees'],
    ['Default Regular Weekly Hours','40','Default target weekly hours'],
    ['Full Schedule Visible To Employees','Yes','If No, employees see only their own schedule'],
    ['Session Hours','8','Login session duration in hours'],
    ['PTO Auto-Approve Per Day','2','First-come, first-served: automatically approve the first two PTO requests covering the same calendar day; later requests require administrator review'],
    ['PTO Third Request Requires Admin','Yes','When the daily auto-approval limit is reached, new PTO requests remain Pending until an administrator approves or rejects them'],
    ['Schedule Title','Pharmacy Staffing Schedule','Title shown in the web application']
  ];
  const sh = ss.getSheetByName(APP.SHEETS.SETTINGS);
  const rows = readTableFromSheet_(sh);
  const existing = new Set(rows.map(r => clean_(r.Setting)));
  const now = new Date();
  const add = defaults.filter(r => !existing.has(r[0])).map(r => [r[0],r[1],r[2],now,'SYSTEM']);
  if (add.length) sh.getRange(sh.getLastRow()+1,1,add.length,APP.HEADERS.SETTINGS.length).setValues(add);
}


/**
 * v14 work-pattern migration.
 *
 * This intentionally repairs existing Settings rows instead of only seeding
 * missing values. The weekend/weekday algorithm is defined on Sunday-Saturday
 * weeks and requires five workdays for every regular (non-7-on/7-off)
 * pharmacist. A 56-day default period gives eight complete work weeks.
 */
function syncWorkPatternSettings_(ss, actor) {
  ss = ss || getSpreadsheet_();
  actor = actor || 'SYSTEM';
  const sh = ss.getSheetByName(APP.SHEETS.SETTINGS);
  if (!sh) return {updated:0};
  const required = {
    'Default Schedule Days':'56',
    'Week Start':'Sunday',
    'Regular Workdays Per Week':'5',
    'Maximum Consecutive Workdays':'5',
    'Block Evening To Morning Transition':'Yes',
    'Required Weekend Assignment':'Yes',
    'Residents Required Assigned Weekend':'Yes',
    'Preceptor Weekday Evening Allowed':'No',
    'Preceptor Weekend Evening Allowed':'Yes',
    'Require Paired ED Coverage':'Yes',
    'ED Day Shift Code':'EDD',
    'ED Evening Shift Code':'EDE',
    'Default Regular Schedule Days':'5'
  };
  const rows = readTableFromSheet_(sh);
  const byKey = {};
  rows.forEach(r=>byKey[clean_(r.Setting)] = r);
  let updated = 0;
  Object.keys(required).forEach(key=>{
    const desired = required[key];
    const row = byKey[key];
    if (row) {
      if (clean_(row.Value) !== desired) {
        updateRowByKey_(APP.SHEETS.SETTINGS,'Setting',key,{
          'Value':desired,
          'Updated At':new Date(),
          'Updated By':actor
        });
        updated++;
      }
    } else {
      appendObjectRow_(APP.SHEETS.SETTINGS,{
        'Setting':key,
        'Value':desired,
        'Description': key==='Week Start' ? 'Beginning of scheduling work week; v12 requires Sunday' :
          key==='Regular Workdays Per Week' ? 'Every regular pharmacist must work exactly five days per complete Sunday-Saturday week' :
          key==='Maximum Consecutive Workdays' ? 'Hard maximum consecutive workdays for regular pharmacists' :
          key==='Block Evening To Morning Transition' ? 'Hard rule: after an Evening shift, the next calendar day cannot be a Day/Morning shift; use OFF or another Evening shift' :
          key==='Required Weekend Assignment' ? 'Reserve assigned A/B/C weekend shifts before weekday scheduling' :
          key==='Residents Required Assigned Weekend' ? 'Residents must work both days of their assigned A/B/C weekend' :
          key==='Preceptor Weekday Evening Allowed' ? 'Hard rule: preceptors cannot work evening shifts Monday-Friday' :
          key==='Preceptor Weekend Evening Allowed' ? 'Preceptors may work evening shifts Saturday-Sunday' :
          key==='Require Paired ED Coverage' ? 'Reserve EDD and EDE as paired daily ED coverage using two different pharmacists' :
          key==='ED Day Shift Code' ? 'Day shift used by paired Emergency Department coverage' :
          key==='ED Evening Shift Code' ? 'Evening shift used by paired Emergency Department coverage' :
          key==='Default Schedule Days' ? 'Eight complete Sunday-Saturday weeks' : 'Expected regular workdays per week',
        'Updated At':new Date(),
        'Updated By':actor
      });
      updated++;
    }
  });
  return {updated:updated};
}

/**
 * Populate the system-managed Weekend Rotation Anchor Date column.
 *
 * Group A/B/C anchors are derived from the global Weekend Anchor Date and
 * Weekend Anchor Group settings. Every non-7-on/7-off employee in the same
 * weekend group receives the same Saturday anchor and therefore works one
 * weekend every three weeks. True 7-on/7-off employees are controlled by
 * Rotation Anchor Date instead and do not use the A/B/C weekend anchor.
 */
function syncWeekendRotationAnchorDates_(ss, actor) {
  ss = ss || getSpreadsheet_();
  actor = actor || 'SYSTEM';
  const sh = ss.getSheetByName(APP.SHEETS.USERS);
  if (!sh || sh.getLastRow() < 2) return {updated:0};

  const raw = settingsMapFromSheet_(ss.getSheetByName(APP.SHEETS.SETTINGS));
  const rotation = clean_(raw['Weekend Rotation'] || 'A,B,C').split(',').map(s=>s.trim()).filter(Boolean);
  const settings = {
    weekendRotation: rotation.length ? rotation : ['A','B','C'],
    weekendAnchorDate: startOfDay_(asDate_(raw['Weekend Anchor Date']) || nextOrSameSaturday_(new Date())),
    weekendAnchorGroup: clean_(raw['Weekend Anchor Group'] || 'A')
  };

  const values = sh.getDataRange().getValues();
  const headers = values[0].map(clean_);
  const groupCol = headers.indexOf('Weekend Group');
  const schedCol = headers.indexOf('Schedule Type');
  const anchorCol = headers.indexOf('Weekend Rotation Anchor Date');
  const updatedAtCol = headers.indexOf('Updated At');
  const updatedByCol = headers.indexOf('Updated By');
  if (groupCol < 0 || anchorCol < 0) return {updated:0};

  let updated = 0;
  const now = new Date();
  for (let r=1; r<values.length; r++) {
    const group = clean_(values[r][groupCol]);
    const scheduleType = schedCol >= 0 ? clean_(values[r][schedCol]) : '';
    const fakeUser = {'Schedule Type':scheduleType};
    let desired = '';
    if (group && settings.weekendRotation.includes(group) && !isSevenOn_(fakeUser)) {
      desired = weekendAnchorForGroup_(group,settings) || '';
    }

    const current = values[r][anchorCol];
    const currentKey = asDate_(current) ? formatDateKey_(asDate_(current)) : '';
    const desiredKey = asDate_(desired) ? formatDateKey_(asDate_(desired)) : '';
    if (currentKey === desiredKey && ((currentKey && current instanceof Date) || (!currentKey && !clean_(current)))) continue;

    values[r][anchorCol] = desired;
    if (updatedAtCol >= 0) values[r][updatedAtCol] = now;
    if (updatedByCol >= 0) values[r][updatedByCol] = actor;
    updated++;
  }

  if (updated) sh.getRange(2,1,values.length-1,headers.length).setValues(values.slice(1));
  return {updated:updated};
}

function settingsMapFromSheet_(sh) {
  const rows = readTableFromSheet_(sh);
  const out = {};
  rows.forEach(r=>out[clean_(r.Setting)] = r.Value);
  return out;
}


function seedShifts_(ss) {
  const seed = [
    ['C7','Central Pharmacy 0700','07:00','15:30',8.5,'','Day','C7','No','Yes',40],
    ['C8','Central Pharmacy 0800','08:00','16:30',8.5,'','Day','C8','No','Yes',40],
    ['CARD','Cardiology','07:00','15:30',8.5,'','Day','CARD','No','Yes',20],
    ['CC1','Critical Care','07:00','15:30',8.5,'','Day','CC1','No','Yes',10],
    ['CC2','Critical Care','07:00','15:30',8.5,'','Day','CC2','No','Yes',10],
    ['EDD','ED Day','07:00','15:30',8.5,'','Day','EDD','No','Yes',15],
    ['IM','Internal Medicine','07:00','15:30',8.5,'','Day','IM','No','Yes',20],
    ['ONC','Oncology','07:00','15:30',8.5,'','Day','ONC','No','Yes',10],
    ['E1','Evenings','13:00','21:30',8.5,'','Evening','E1','No','Yes',25],
    ['E2','Evenings','13:00','21:30',8.5,'','Evening','E2','No','Yes',25],
    ['EDE','ED Evenings','14:00','22:30',8.5,'','Evening','EDE','No','Yes',15],
    ['E','Evening Central','11:00','21:30',10.5,'','Evening','E','Yes','Yes',15],
    ['WC7','Weekend Central','08:00','16:30',8.5,'','Day','WC7','Yes','Yes',15],
    ['WD1','Weekend Clinical','07:00','15:30',8.5,'','Day','WD1','Yes','Yes',10],
    ['WD2','Weekend Clinical','07:00','15:30',8.5,'','Day','WD2','Yes','Yes',10],
    ['WE1','Evenings','13:00','21:30',8.5,'','Evening','WE1','No','No',25],
    ['WEDE','Evenings ED','13:00','21:30',8.5,'','Evening','WEDE','Yes','Yes',15],
    ['WMC','Weekend Midday Central Pharmacy','10:00','18:30',8.5,'','Evening','WMC','Yes','Yes',20],
    ['N1','Nights','20:30','07:00',10.5,'','Night','N1','Yes','Yes',5],
    ['N2','Nights','21:00','07:30',10.5,'','Night','N2','Yes','Yes',5],
    ['SUP','Supervisor','','','','','Other','SUP','No','No',50],
    ['RES','Resident','','','','','Training','RES','No','No',50],
    ['ED11','ED Day/Evening Overlap','11:00','19:00',8,'','Day','ED11','No','No',30],
    ['TDP','Training / TDP','','','','','Training','TDP','No','No',50]
  ];
  const sh = ss.getSheetByName(APP.SHEETS.SHIFTS);
  const existing = new Set(readTableFromSheet_(sh).map(r => clean_(r.Shift)));
  const now = new Date();
  const add = seed.filter(r => !existing.has(r[0])).map(r => r.concat([now,'SYSTEM']));
  if (add.length) sh.getRange(sh.getLastRow()+1,1,add.length,APP.HEADERS.SHIFTS.length).setValues(add);
}

function seedStaffingRequirements_(ss) {
  const raw = {
    C7:[0,1,1,1,1,1,0], C8:[0,1,1,1,1,1,0], CARD:[0,1,1,1,1,1,0],
    CC1:[0,1,1,1,1,1,0], CC2:[0,1,1,1,1,1,0], EDD:[0,1,1,1,1,1,0],
    IM:[0,1,1,1,1,1,0], ONC:[0,1,1,1,1,1,0], E1:[0,1,1,1,1,1,0],
    E2:[0,1,1,1,1,1,0], EDE:[0,1,1,1,1,1,0], E:[1,1,1,1,1,1,1],
    WC7:[1,0,0,0,0,0,1], WD1:[1,0,0,0,0,0,1], WD2:[1,0,0,0,0,0,1],
    WE1:[0,0,0,0,0,0,0], WEDE:[1,0,0,0,0,0,1], WMC:[1,0,0,0,0,0,1],
    N1:[1,1,1,1,1,1,1], N2:[1,1,1,1,1,1,1], SUP:[0,0,0,0,0,0,0],
    RES:[0,0,0,0,0,0,0], ED11:[0,0,0,0,0,0,0], TDP:[0,0,0,0,0,0,0]
  };
  const sh = ss.getSheetByName(APP.SHEETS.REQUIREMENTS);
  const existing = new Set(readTableFromSheet_(sh).map(r => clean_(r.Shift)));
  const now = new Date();
  const add = Object.keys(raw).filter(k => !existing.has(k)).map(k => [k].concat(raw[k]).concat(['Yes',now,'SYSTEM']));
  if (add.length) sh.getRange(sh.getLastRow()+1,1,add.length,APP.HEADERS.REQUIREMENTS.length).setValues(add);
}

function seedInitialAdmin_(ss) {
  const r = seedFourAdmins_(ss);
  const first = r.created[0] || {};
  return {created:!!r.created.length, username:first.username || 'admin', temporaryPassword:first.temporaryPassword || ''};
}

function formatSheets_(ss) {
  Object.values(APP.SHEETS).forEach(name => {
    const sh = ss.getSheetByName(name);
    if (!sh) return;
    sh.getRange(1,1,1,Math.max(1,sh.getLastColumn())).setFontWeight('bold').setWrap(true);
    sh.autoResizeColumns(1, Math.max(1, Math.min(sh.getLastColumn(), 15)));
  });
}

/** ------------------------- AUTHENTICATION ------------------------ */

function login(username, password) {
  username = clean_(username).toLowerCase();
  if (!username || !password) throw new Error('Username and password are required.');

  const admins = readTable_(APP.SHEETS.ADMINS);
  const admin = admins.find(r => clean_(r.Username).toLowerCase() === username && yesDefault_(r.Active,true));
  if (!admin) {
    Utilities.sleep(250);
    throw new Error('Invalid administrator username or password.');
  }
  const salt = clean_(admin['Password Salt']);
  const hash = clean_(admin['Password Hash']);
  if (!salt || !hash) throw new Error('This administrator does not have a valid password. Reset the administrator password.');
  if (hashPassword_(String(password),salt) !== hash) {
    Utilities.sleep(250);
    throw new Error('Invalid administrator username or password.');
  }

  const settings = getSettingsMap_();
  const hours = num_(settings['Session Hours'], APP.SESSION_HOURS);
  const token = Utilities.getUuid() + Utilities.getUuid().replace(/-/g,'');
  const now = new Date();
  const expires = new Date(now.getTime() + hours * 3600000);
  cleanupSessions_();
  getDb_().getSheetByName(APP.SHEETS.SESSIONS).appendRow([token, admin.Username, now, expires, now]);
  audit_('LOGIN','',admin['Admin Name'],'','','','','No','','Administrator logged in',admin.Username);
  return {
    ok:true,
    token:String(token),
    mustChangePassword:yes_(admin['Must Change Password']),
    version:APP.VERSION
  };
}

function serverPing() {
  return {
    ok:true,
    appName:APP.NAME,
    version:APP.VERSION,
    timestamp:Utilities.formatDate(new Date(), getTz_(), 'yyyy-MM-dd HH:mm:ss')
  };
}

function testSchedulerBackend() {
  const ss = getDb_();
  const result = {
    ok: true,
    version: APP.VERSION,
    spreadsheetName: ss.getName(),
    adminSheetExists: !!ss.getSheetByName(APP.SHEETS.ADMINS),
    sessionSheetExists: !!ss.getSheetByName(APP.SHEETS.SESSIONS),
    weeklyAvailabilitySheetExists: !!ss.getSheetByName(APP.SHEETS.WEEKLY_AVAILABILITY)
  };
  Logger.log(JSON.stringify(result, null, 2));
  return result;
}

function logout(token) {
  const session = getSession_(token, false);
  if (!session) return {ok:true};
  deleteSession_(token);
  return {ok:true};
}

function changePassword(token, currentPassword, newPassword) {
  const ctx = requireAuth_(token);
  if (!newPassword || String(newPassword).length < 10) throw new Error('New password must be at least 10 characters.');
  const admins = readTable_(APP.SHEETS.ADMINS);
  const admin = admins.find(r => clean_(r.Username) === ctx.username);
  if (!admin) throw new Error('Administrator account not found.');
  if (hashPassword_(String(currentPassword), clean_(admin['Password Salt'])) !== clean_(admin['Password Hash'])) throw new Error('Current password is incorrect.');
  const salt = Utilities.getUuid();
  updateRowByKey_(APP.SHEETS.ADMINS,'Admin ID',admin['Admin ID'],{
    'Temporary Password':'',
    'Password Salt':salt,
    'Password Hash':hashPassword_(String(newPassword),salt),
    'Must Change Password':'No',
    'Updated At':new Date(),
    'Updated By':ctx.username
  });
  audit_('PASSWORD_CHANGED','',admin['Admin Name'],'','','','','No','','Administrator password changed',ctx.username);
  return {ok:true};
}

function requireAuth_(token) {
  const session = getSession_(token, true);
  if (!session) throw new Error('Your session has expired. Please sign in again.');
  const admins = readTable_(APP.SHEETS.ADMINS);
  const admin = admins.find(r => clean_(r.Username) === clean_(session.Username) && yesDefault_(r.Active,true));
  if (!admin) throw new Error('Administrator account is inactive or unavailable.');
  return {username:admin.Username, user:adminSessionUser_(admin), admin:admin, isAdmin:true};
}

function requireAdmin_(token) {
  return requireAuth_(token);
}

function adminSessionUser_(admin) {
  return {
    'Admin ID':admin['Admin ID'],
    'Pharmacist Name':admin['Admin Name'],
    'Admin Name':admin['Admin Name'],
    'Username':admin.Username,
    'Role':'Administrator',
    'Active':admin.Active
  };
}

function publicAdmin_(a) {
  const x = Object.assign({},a);
  delete x['Temporary Password'];
  delete x['Password Hash'];
  delete x['Password Salt'];
  return x;
}

function getSession_(token, touch) {
  token = clean_(token);
  if (!token) return null;
  const rows = readTable_(APP.SHEETS.SESSIONS);
  const s = rows.find(r => clean_(r.Token) === token);
  if (!s) return null;
  const exp = asDate_(s['Expires At']);
  if (!exp || exp.getTime() <= Date.now()) {
    deleteSession_(token);
    return null;
  }
  if (touch) updateRowByKey_(APP.SHEETS.SESSIONS,'Token',token,{'Last Seen At':new Date()});
  return s;
}

function deleteSession_(token) {
  const sh = getDb_().getSheetByName(APP.SHEETS.SESSIONS);
  const data = sh.getDataRange().getValues();
  for (let i=data.length-1;i>=1;i--) if (clean_(data[i][0]) === clean_(token)) sh.deleteRow(i+1);
}

function cleanupSessions_() {
  const sh = getDb_().getSheetByName(APP.SHEETS.SESSIONS);
  if (sh.getLastRow() < 2) return;
  const data = sh.getDataRange().getValues();
  const keep = [data[0]];
  const now = Date.now();
  for (let i=1;i<data.length;i++) {
    const exp = asDate_(data[i][3]);
    if (exp && exp.getTime() > now) keep.push(data[i]);
  }
  if (keep.length !== data.length) {
    sh.clearContents();
    sh.getRange(1,1,keep.length,keep[0].length).setValues(keep);
  }
}

/** ---------------------------- APP DATA --------------------------- */

function getAppData(token) {
  const ctx = requireAuth_(token);

  /*
   * Keep Shifts and Staffing Requirements synchronized. This repairs older
   * custom shifts that were created before automatic propagation existed.
   */
  ensureAllShiftStaffingRequirements_('SYSTEM');

  // Keep PTO auto-approval status synchronized on every app load/refresh.
  // This also upgrades legacy Pending rows created before v27.
  reconcilePtoAutoApprovals_('SYSTEM');
  const users = readTable_(APP.SHEETS.USERS).filter(r => clean_(r.Role).toLowerCase() !== 'administrator');
  const prnAvailabilitySourceLoaded = !!getDb_().getSheetByName('My Availability');
  const prnAvailability = normalizePrnAvailabilityRows_(
    readTable_('My Availability'),
    users
  );
  const admins = readTable_(APP.SHEETS.ADMINS).map(publicAdmin_);
  const skills = readTable_(APP.SHEETS.SKILLS);
  const shifts = readTable_(APP.SHEETS.SHIFTS);
  const requirements = readTable_(APP.SHEETS.REQUIREMENTS);
  const requests = readTable_(APP.SHEETS.REQUESTS);
  const weeklyAvailability = readTable_(APP.SHEETS.WEEKLY_AVAILABILITY);
  // v30 migration: v17-v29 chunked generation could create duplicate
  // Assignment IDs because the per-request sequence restarted each week.
  // Repair those legacy duplicates before the browser receives schedule data.
  // This is idempotent: once IDs are unique, subsequent loads make no changes.
  ensureUniqueScheduleAssignmentIds_();
  const schedule = readTable_(APP.SHEETS.SCHEDULE);
  const settings = getSettingsMap_();
  const health = computeScheduleHealth_(schedule);
  const stats = computeEmployeeStats_(schedule, users);
  return serialize_({
    appName:APP.NAME,
    version:APP.VERSION,
    database:getDatabaseInfo_(),
    user:ctx.user,
    isAdmin:true,
    admins:admins,
    users:users.map(publicUser_),
    skills:skills,
    shifts:shifts,
    requirements:requirements,
    requests:requests,
    weeklyAvailability:weeklyAvailability,
    prnAvailability:prnAvailability,
    prnAvailabilitySourceLoaded:prnAvailabilitySourceLoaded,
    schedule:schedule,
    settings:settings,
    health:health,
    stats:stats,
    configValidation:validateConfiguration_()
  });
}

function publicUser_(u) {
  const x = Object.assign({}, u);
  delete x['Temporary Password'];
  delete x['Password Hash'];
  delete x['Password Salt'];
  return x;
}

/** ----------------------- PRN MY AVAILABILITY --------------------- */

function isPrnEmployee_(u) {
  return clean_(u && u['Employment Type']).toUpperCase()==='PRN' ||
    clean_(u && u['Schedule Type']).toUpperCase()==='PRN';
}

function prnRowValue_(row, aliases) {
  row=row||{};
  const keys=Object.keys(row);
  const map={};
  keys.forEach(k=>map[clean_(k).toLowerCase()]=k);

  for(const alias of aliases||[]){
    const actual=map[clean_(alias).toLowerCase()];
    if(actual!==undefined){
      const value=row[actual];
      if(value!==''&&value!==null&&value!==undefined)return value;
    }
  }

  return '';
}

function prnAvailabilityDateKey_(value) {
  const d=asDate_(value);
  return d?formatDateKey_(startOfDay_(d)):'';
}

function prnAvailabilityHeaderDateKey_(header) {
  const raw=clean_(header);
  if(!raw)return '';

  let m=raw.match(/^(\d{4})[-\/]([01]?\d)[-\/]([0-3]?\d)$/);
  if(m)return m[1]+'-'+String(Number(m[2])).padStart(2,'0')+'-'+String(Number(m[3])).padStart(2,'0');

  m=raw.match(/^([01]?\d)[-\/]([0-3]?\d)[-\/](\d{4})$/);
  if(m)return m[3]+'-'+String(Number(m[1])).padStart(2,'0')+'-'+String(Number(m[2])).padStart(2,'0');

  return '';
}

function prnAvailabilityTruth_(value, defaultValue) {
  const raw=clean_(value).toLowerCase();
  if(!raw)return !!defaultValue;
  if(['no','n','false','0','off','unavailable','not available','cannot work','cant work'].includes(raw))return false;
  if(['yes','y','true','1','available','avail','x','can work','open'].includes(raw))return true;
  return !!defaultValue;
}

function prnAvailabilityCellShift_(value, shiftCodes) {
  const raw=clean_(value).toUpperCase();
  if(!raw)return '';

  const direct=(shiftCodes||[]).find(code=>code===raw);
  if(direct)return direct;

  const tokens=raw.split(/[,;|\/\s]+/).map(x=>x.trim()).filter(Boolean);
  const matches=tokens.filter(x=>(shiftCodes||[]).includes(x));
  return matches.length?matches.join(','):'';
}

function normalizePrnAvailabilityRows_(rawRows, users) {
  rawRows=Array.isArray(rawRows)?rawRows:[];
  users=Array.isArray(users)?users:[];

  const prnUsers=users.filter(isPrnEmployee_);
  const byUsername={};
  const byEmployeeId={};
  const byName={};

  prnUsers.forEach(u=>{
    const username=clean_(u.Username);
    const employeeId=clean_(u['Employee ID']);
    const name=clean_(u['Pharmacist Name']);

    if(username)byUsername[username.toLowerCase()]=u;
    if(employeeId)byEmployeeId[employeeId.toLowerCase()]=u;
    if(name)byName[name.toLowerCase()]=u;
  });

  const shiftCodes=readTable_(APP.SHEETS.SHIFTS)
    .filter(s=>yesDefault_(s.Active,true))
    .map(s=>clean_(s.Shift).toUpperCase())
    .filter(Boolean);

  function resolveUser(row){
    const username=clean_(prnRowValue_(row,['Username','User Name','User','Email'])).toLowerCase();
    const employeeId=clean_(prnRowValue_(row,['Employee ID','Employee Id','EmployeeID','ID'])).toLowerCase();
    const name=clean_(prnRowValue_(row,['Pharmacist','Pharmacist Name','Employee','Employee Name','Name'])).toLowerCase();

    return (username&&byUsername[username]) ||
      (employeeId&&byEmployeeId[employeeId]) ||
      (name&&byName[name]) ||
      null;
  }

  const out=[];
  const seen=new Set();

  function pushAvailability(user, dateKey, row, available, shiftOverride){
    if(!user||!dateKey)return;

    const username=clean_(user.Username);
    const shift=clean_(shiftOverride || prnRowValue_(row,[
      'Shift','Shift Type','Preferred Shift','Available Shift','Preferred Shift Type'
    ])).toUpperCase();

    const start=clean_(prnRowValue_(row,['Start Time','Available Start','Start']));
    const end=clean_(prnRowValue_(row,['End Time','Available End','End']));
    const key=[
      username.toLowerCase(),
      dateKey,
      shift,
      start,
      end,
      available?'Y':'N'
    ].join('|');

    if(seen.has(key))return;
    seen.add(key);

    out.push({
      'Pharmacist':clean_(user['Pharmacist Name']),
      'Username':username,
      'Employee ID':clean_(user['Employee ID']),
      'Date':dateKey,
      'Available':available?'Yes':'No',
      'Shift':shift,
      'Start Time':start,
      'End Time':end,
      'Source':'My Availability'
    });
  }

  rawRows.forEach(row=>{
    const user=resolveUser(row);
    if(!user)return;

    const activeValue=prnRowValue_(row,['Active','Enabled']);
    if(activeValue!==''&&!prnAvailabilityTruth_(activeValue,true))return;

    const status=clean_(prnRowValue_(row,['Status'])).toLowerCase();
    if(['deleted','inactive','cancelled','canceled','rejected'].includes(status))return;

    const explicitAvailable=prnRowValue_(row,['Available','Availability','Can Work','Can Work?','Working']);
    const available=prnAvailabilityTruth_(explicitAvailable,true);

    const single=prnRowValue_(row,['Availability Date','Available Date','Work Date','Shift Date','Date']);
    let startDate=prnRowValue_(row,['Start Date','Available Start Date']);
    let endDate=prnRowValue_(row,['End Date','Available End Date']);

    if(single){
      pushAvailability(user,prnAvailabilityDateKey_(single),row,available,'');
      return;
    }

    if(startDate||endDate){
      let start=prnAvailabilityDateKey_(startDate||endDate);
      let end=prnAvailabilityDateKey_(endDate||startDate);

      if(start&&end&&end>=start){
        let d=startOfDay_(asDate_(start));
        const finish=startOfDay_(asDate_(end));
        let guard=0;

        while(d&&finish&&d<=finish&&guard<370){
          pushAvailability(user,formatDateKey_(d),row,available,'');
          d=addDays_(d,1);
          guard++;
        }
      }

      return;
    }

    // Also support a wide monthly sheet where each date is a column and the
    // cell contains Yes/Available/X or an optional shift code such as CC1.
    Object.keys(row||{}).forEach(header=>{
      const dateKey=prnAvailabilityHeaderDateKey_(header);
      if(!dateKey)return;

      const cell=row[header];
      if(cell===''||cell===null||cell===undefined)return;

      const cellText=clean_(cell);
      const cellAvailable=prnAvailabilityTruth_(cellText,true);
      const cellShift=prnAvailabilityCellShift_(cellText,shiftCodes);

      pushAvailability(user,dateKey,row,cellAvailable,cellShift);
    });
  });

  return out.sort((a,b)=>
    clean_(a.Date).localeCompare(clean_(b.Date)) ||
    clean_(a.Pharmacist).localeCompare(clean_(b.Pharmacist))
  );
}

/** ------------------------- CONFIG VALIDATION --------------------- */

function validateConfiguration(token) {
  requireAdmin_(token);
  return serialize_(validateConfiguration_());
}

function validateConfiguration_() {
  const model = loadSchedulingModel_();
  const errors = [], warnings = [];
  if(clean_(model.settings.weekStart).toLowerCase()!=='sunday') errors.push('Week Start must be Sunday for the five-day weekend-compensation algorithm. Run setupPharmacyScheduler() to repair Settings.');
  if(model.settings.regularWorkdaysPerWeek!==5) errors.push('Regular Workdays Per Week must be 5. Run setupPharmacyScheduler() to repair Settings.');
  if(model.settings.maxConsecutiveWorkdays!==5) errors.push('Maximum Consecutive Workdays must be 5. Run setupPharmacyScheduler() to repair Settings.');
  if(!model.settings.blockEveningToMorningTransition) errors.push('Block Evening To Morning Transition must be Yes. After an evening shift, the next day must be OFF or another evening shift, not a day/morning shift. Run setupPharmacyScheduler() to repair Settings.');
  if(model.settings.preceptorWeekdayEveningAllowed) errors.push('Preceptor Weekday Evening Allowed must be No. Preceptors may work evening shifts on weekends only. Run setupPharmacyScheduler() to repair Settings.');
  if(!model.settings.preceptorWeekendEveningAllowed) errors.push('Preceptor Weekend Evening Allowed must be Yes. Run setupPharmacyScheduler() to repair Settings.');
  const groups = model.settings.weekendRotation;

  model.users.forEach(u => {
    if (!yes_(u.Active)) return;
    if (!numOrNull_(u['Weekly Hour Maximum'])) errors.push('Employee ' + u['Pharmacist Name'] + ' is missing Weekly Hour Maximum.');
    const weekendGroup = clean_(u['Weekend Group']);
    if (model.settings.residentsRequiredAssignedWeekend && yes_(u.Resident) && !isSevenOn_(u) && !weekendGroup) errors.push('Resident ' + u['Pharmacist Name'] + ' must have Weekend Group A, B, or C because residents are required to work their assigned weekend.');
    if (weekendGroup && !groups.includes(weekendGroup)) errors.push('Employee ' + u['Pharmacist Name'] + ' has invalid Weekend Group ' + u['Weekend Group'] + '.');
    if (weekendGroup && !isSevenOn_(u) && !yesDefault_(u['Weekend Eligible'],true)) errors.push('Employee ' + u['Pharmacist Name'] + ' is in Weekend Group ' + weekendGroup + ' but Weekend Eligible is No. Required weekend assignments cannot be satisfied.');
    if (isSevenOn_(u) && !asDate_(u['Rotation Anchor Date'])) errors.push('7-on/7-off employee ' + u['Pharmacist Name'] + ' has no Rotation Anchor Date.');
    if (!isSevenOn_(u) && weekendGroup) {
      const weekendAnchor = asDate_(u['Weekend Rotation Anchor Date']);
      if (!weekendAnchor) {
        errors.push('Employee ' + u['Pharmacist Name'] + ' is in Weekend Group ' + weekendGroup + ' but has no Weekend Rotation Anchor Date. Run setupPharmacyScheduler() to repair it.');
      } else {
        if (dayIndex_(weekendAnchor) !== 6) errors.push('Employee ' + u['Pharmacist Name'] + ' Weekend Rotation Anchor Date must be a Saturday.');
        const anchorGroup = weekendGroupForDate_(weekendAnchor,model.settings);
        if (anchorGroup !== weekendGroup) errors.push('Employee ' + u['Pharmacist Name'] + ' Weekend Rotation Anchor Date does not match Weekend Group ' + weekendGroup + '.');
      }
    }
  });

  model.requirements.forEach(req => {
    if (!yesDefault_(req.Active, true)) return;
    const shift = model.shiftMap[clean_(req.Shift)];
    const total = dayNames_().reduce((n,d)=>n+Math.max(0,Math.floor(num_(req[d],0))),0);
    if (!shift) {
      if (total > 0) errors.push('Staffing requirement references unknown shift ' + req.Shift + '.');
      return;
    }
    if (total > 0 && !yes_(shift.Active)) errors.push('Required shift ' + req.Shift + ' is inactive.');
    if (total > 0 && (!clean_(shift.Start) || !clean_(shift.End))) errors.push('Required shift ' + req.Shift + ' has no complete start/end time.');
    if (total > 0 && !numOrNull_(shift['Credited Hours']) && !numOrNull_(shift.Hours)) errors.push('Required shift ' + req.Shift + ' has neither Credited Hours nor Hours.');
    if (total > 0 && clean_(shift.Skill)) {
      const qualified = model.users.some(u => yes_(u.Active) && model.skillsByUser[clean_(u.Username)] && model.skillsByUser[clean_(u.Username)].has(clean_(shift.Skill)));
      if (!qualified) errors.push('Shift ' + req.Shift + ' requires skill ' + shift.Skill + ' but no active pharmacist has that skill.');
    }
  });


  if(model.settings.requirePairedEdCoverage){
    const dayCode=clean_(model.settings.edDayShiftCode||'EDD').toUpperCase();
    const eveCode=clean_(model.settings.edEveningShiftCode||'EDE').toUpperCase();
    const dayShift=model.shiftMap[dayCode], eveShift=model.shiftMap[eveCode];
    if(!dayShift||!yes_(dayShift.Active)) errors.push('Paired ED coverage requires an active ED Day shift code '+dayCode+'.');
    if(!eveShift||!yes_(eveShift.Active)) errors.push('Paired ED coverage requires an active ED Evening shift code '+eveCode+'.');
    if(dayShift&&eveShift){
      const eveSkill=clean_(eveShift.Skill);
      const weekdayEdeEligible=model.users.filter(u=>{
        if(!yes_(u.Active))return false;
        const skills=model.skillsByUser[clean_(u.Username)];
        if(eveSkill&&!(skills&&skills.has(eveSkill)))return false;
        if(!yesDefault_(u['Evening Eligible'],true))return false;
        if(yes_(u.Preceptor)&&!model.settings.preceptorWeekdayEveningAllowed)return false;
        if(isSevenOn_(u))return false;
        return true;
      });
      if(!weekdayEdeEligible.length){
        warnings.push('Paired ED coverage warning: no active non-preceptor pharmacist is currently eligible for weekday '+eveCode+'. Under the rule that preceptors may work evening shifts on weekends only, weekday '+eveCode+' will remain UNFILLED until a qualified non-preceptor is available.');
      }
    }
  }

  model.weeklyAvailability.forEach(r => {
    if (!yesDefault_(r.Active,true)) return;
    const username=clean_(r.Username);
    const u=model.usersByUsername[username] || model.usersByName[clean_(r['Pharmacist Name'])];
    if(!u){errors.push('Weekly availability row '+clean_(r['Availability ID'])+' references an unknown employee.');return;}
    const day=clean_(r.Day);
    if(!dayNames_().includes(day)) errors.push('Weekly availability for '+u['Pharmacist Name']+' has invalid day '+day+'.');
    const rule=clean_(r['Rule Type']||'HARD').toUpperCase();
    if(!['HARD','SOFT'].includes(rule)) errors.push('Weekly availability for '+u['Pharmacist Name']+' has invalid Rule Type '+rule+'.');
    const available=yes_(r.Available);
    const st=timeMinutes_(r['Start Time']),en=timeMinutes_(r['End Time']);
    if(available && ((st===null)!=(en===null))) errors.push('Weekly availability for '+u['Pharmacist Name']+' on '+day+' must have both Start Time and End Time, or leave both blank for all-day availability.');
    const es=asDate_(r['Effective Start']),ee=asDate_(r['Effective End']);
    if(es&&ee&&formatDateKey_(ee)<formatDateKey_(es)) errors.push('Weekly availability for '+u['Pharmacist Name']+' on '+day+' has Effective End before Effective Start.');
  });

  model.shifts.forEach(s => {
    if (yes_(s.Active) && clean_(s.Type).toLowerCase() !== 'training' && (!clean_(s.Start) || !clean_(s.End))) warnings.push('Active shift ' + s.Shift + ' has incomplete start/end time.');
    if (yes_(s.Active) && !numOrNull_(s['Credited Hours']) && numOrNull_(s.Hours)) warnings.push('Shift ' + s.Shift + ' has blank Credited Hours; Hours will be used for weekly limits.');
  });

  if (!asDate_(model.settings.raw['Weekend Anchor Date'])) errors.push('Weekend Anchor Date is missing or invalid.');
  if (!groups.includes(model.settings.weekendAnchorGroup)) errors.push('Weekend Anchor Group must be one of: ' + groups.join(', ') + '.');
  return {errors:unique_(errors), warnings:unique_(warnings)};
}

/** ---------------------- SCHEDULE GENERATION ---------------------- */


function isProtectedPreScheduleRow_(r) {
  const status=clean_(r&&r.Status).toUpperCase();
  const filled=status!=='UNFILLED' &&
    clean_(r&&r.Username) &&
    clean_(r&&r['Assigned Pharmacist']).toUpperCase()!=='UNFILLED';
  if(!filled)return false;
  return yes_(r.Locked) || yes_(r.Manual) || status==='MANUAL' || status==='LOCKED';
}

function preflightScheduleGeneration(token, options) {
  requireAdmin_(token);
  options=options||{};
  const cfg=validateConfiguration_();
  if(cfg.errors.length) return {ok:false,stage:'CONFIGURATION',errors:cfg.errors,warnings:cfg.warnings};
  const model=loadSchedulingModel_();
  const defaultDays=num_(model.settings.raw['Default Schedule Days'],56);
  const start=startOfDay_(asDate_(options.startDate)||new Date());
  const end=startOfDay_(asDate_(options.endDate)||addDays_(start,defaultDays-1));
  if(end<start) return {ok:false,stage:'PERIOD',errors:['End Date must be on or after Start Date.'],warnings:[]};
  const totalDays=daysBetween_(start,end)+1;
  if(totalDays>181) return {ok:false,stage:'PERIOD',errors:['For safety, one generation cannot exceed 181 days.'],warnings:[]};

  const periodWarnings=(cfg.warnings||[]).slice();
  const hasRegularEmployees=model.users.some(u=>yes_(u.Active)&&regularFiveDayRuleApplies_(u));
  if(
    hasRegularEmployees &&
    (dayIndex_(start)!==0 || dayIndex_(end)!==6 || totalDays%7!==0)
  ){
    periodWarnings.unshift(
      'The selected range begins or ends in a partial Sunday-Saturday week. The exact five-workday rule will be enforced only for complete Sunday-Saturday weeks fully contained inside the selected range.'
    );
  }

  const existing=readTable_(APP.SHEETS.SCHEDULE);
  const overlap=existing.filter(r=>inDateRange_(asDate_(r.Date),start,end));
  const finalized=overlap.filter(r=>asDate_(r['Finalized At']));
  const locked=overlap.filter(r=>yes_(r.Locked));
  const protectedRows=overlap.filter(isProtectedPreScheduleRow_);
  const replaceableRows=overlap.filter(r=>!isProtectedPreScheduleRow_(r));
  const overlapDates=[...new Set(overlap.map(r=>formatDateKey_(startOfDay_(asDate_(r.Date)))).filter(Boolean))].sort();

  if(replaceableRows.length){
    periodWarnings.unshift(
      'REGENERATION WARNING: '+replaceableRows.length+' existing generated/unprotected schedule row(s) on '+overlapDates.length+
      ' date(s) fall inside the selected range and will be replaced if the administrator confirms. Manual and locked assignments are protected.'
    );
  }
  if(protectedRows.length){
    periodWarnings.unshift(
      protectedRows.length+' manual/locked pre-scheduling assignment(s) are protected and will be used as fixed constraints during generation.'
    );
  }

  const tz=getTz_();
  const batchId='GEN_'+Utilities.formatDate(start,tz,'yyyyMMdd')+'_'+Utilities.formatDate(end,tz,'yyyyMMdd')+'_'+Utilities.formatDate(new Date(),tz,'HHmmss');
  return serialize_({
    ok:true,
    batchGenerationId:batchId,
    startDate:formatDateKey_(start),
    endDate:formatDateKey_(end),
    totalDays:totalDays,
    totalWeeks:Math.ceil(totalDays/7),
    overlapCount:overlap.length,
    overlapDateCount:overlapDates.length,
    finalizedOverlapCount:finalized.length,
    lockedOverlapCount:locked.length,
    protectedCount:protectedRows.length,
    manualProtectedCount:protectedRows.filter(r=>yes_(r.Manual)||['MANUAL','LOCKED'].includes(clean_(r.Status).toUpperCase())).length,
    replaceableCount:replaceableRows.length,
    requiresOverwriteConfirmation:replaceableRows.length>0,
    warnings:periodWarnings
  });
}

function generateSchedule(token, options) {
  const ctx = requireAdmin_(token);
  options = options || {};
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) throw new Error('Another administrator is currently modifying or generating the schedule. Try again after that operation finishes.');
  try {
    const generationStartedAt=Date.now();
    _GENERATED_ASSIGNMENT_SEQ=0;
    const cfg = validateConfiguration_();
    if (cfg.errors.length) return {ok:false, stage:'CONFIGURATION', errors:cfg.errors, warnings:cfg.warnings};

    const model = loadSchedulingModel_();
    // v16: stop optional repair/coverage work before Apps Script's hard execution limit.
    // Core assignment is still completed; only expensive repair search is curtailed if needed.
    model.runtimeDeadline=generationStartedAt+(yes_(options.chunked)?45000:240000);
    const defaultDays = num_(model.settings.raw['Default Schedule Days'],56);
    const start = startOfDay_(asDate_(options.startDate) || new Date());
    const end = startOfDay_(asDate_(options.endDate) || addDays_(start, defaultDays-1));
    if (end < start) throw new Error('End Date must be on or after Start Date.');
    if (daysBetween_(start,end) > 180) throw new Error('For safety, one generation cannot exceed 181 days.');

    // The administrator may generate any calendar date range. Partial
    // Sunday-Saturday boundary weeks are allowed. Exact five-workday validation
    // still applies to every complete Sunday-Saturday week fully inside the range.
    const mode = 'OVERWRITE';
    const existing = readTable_(APP.SHEETS.SCHEDULE);
    const overlap = existing.filter(r => inDateRange_(asDate_(r.Date),start,end));
    const preserved = overlap.filter(isProtectedPreScheduleRow_);
    const replaceable = overlap.filter(r=>!isProtectedPreScheduleRow_(r));

    // Manual and locked pre-scheduling assignments are FIXED constraints.
    // Only generated/unprotected rows are eligible for regeneration.
    if (replaceable.length && !yes_(options.overwriteConfirmed)) {
      return {
        ok:false,
        stage:'OVERLAP_CONFIRMATION_REQUIRED',
        errors:[
          'The selected dates contain '+replaceable.length+
          ' existing generated/unprotected schedule row(s). Confirm regeneration before continuing. Manual and locked assignments will remain protected.'
        ],
        warnings:[]
      };
    }

    // The allocator is seeded with every manual or locked assignment so their
    // workdays, hours, evenings, weekends, resident duties, rotations and rest
    // transitions are counted BEFORE any new shift is assigned.
    const generationId = clean_(options.batchGenerationId) || makeGenerationId_(start,end);
    const slots = createRequiredSlots_(model,start,end,generationId);
    const preservedKeys = new Set(preserved.map(r => slotKey_(dateKey_(r.Date),clean_(r.Shift),num_(r.Slot,1))));
    const generatedSlots = slots.filter(s => !preservedKeys.has(s.slotKey));

    const contextStart = generationContextStart_(start, model.settings.weekStart);
    const contextEnd = generationContextEnd_(end, model.settings.weekStart);
    const fixedContext = existing.filter(r => {
      const d = asDate_(r.Date);
      return d && !inDateRange_(d,start,end) && inDateRange_(d,contextStart,contextEnd) && clean_(r.Status) !== 'UNFILLED' && clean_(r.Username);
    }).concat(preserved);
    const state = createState_(model, fixedContext, start, end);
    const slotResults = preserved.map(r => scheduleRowToAssignment_(r, model)).filter(Boolean);
    const lockedSnapshot = preserved.filter(assignmentIsLocked_).map(r => lockedFingerprint_(r));

    // Mandatory-pattern passes happen before the normal fairness allocator.
    // 1) True 7-on/7-off employees stay on one dedicated shift for each ON block.
    // 2) Residents are placed on their REQUIRED A/B/C weekend first.
    // 3) Remaining A/B/C weekend-team assignments are reserved and protected.
    // 4) EDD + EDE are treated as paired daily ED coverage using two different
    //    eligible pharmacists whenever both shifts are required.
    // 5) Every resident receives one E2 per week; E2 counts as one of the five
    //    workdays for regular residents.
    const reservedSlotKeys=new Set();
    if (model.settings.sevenOnBlockScheduling) {
      preassignSevenOnSevenOff_(generatedSlots,slotResults,model,state,ctx.username,reservedSlotKeys);
    }
    if (model.settings.requiredWeekendAssignment) {
      if (model.settings.residentsRequiredAssignedWeekend) {
        preassignResidentWeekendsPriority_(generatedSlots,slotResults,model,state,ctx.username,reservedSlotKeys,start,end);
      }
      preassignWeekendTeams_(generatedSlots,slotResults,model,state,ctx.username,reservedSlotKeys,start,end);
    }
    if (model.settings.requirePairedEdCoverage) {
      preassignPairedEdCoverage_(generatedSlots,slotResults,model,state,ctx.username,reservedSlotKeys,start,end);
    }
    preassignResidentE2_(generatedSlots,slotResults,model,state,ctx.username,reservedSlotKeys,start,end);

    const remainingSlots=generatedSlots.filter(s=>!reservedSlotKeys.has(s.slotKey));

    // Manager skill-priority ordering.
    // Preferred/Home Unit owners who are available stay first. If a home-unit
    // owner is unavailable, backup pharmacists are considered according to the
    // manager's Employee Skills Priority (1 first, then 2, 3...). Scarcity and
    // the original deterministic shift priorities remain tie-breakers.
    remainingSlots.forEach(s => {
      s.skillPriorityRank =
        typeof skillPrioritySlotRank_ === 'function'
          ? skillPrioritySlotRank_(s,model,state)
          : 99999;
      s.scarcity = countStaticCandidates_(s, model);
      s.categoryRank = shiftCategoryRank_(s.shift);
    });
    remainingSlots.sort((a,b) =>
      a.skillPriorityRank-b.skillPriorityRank ||
      a.scarcity-b.scarcity ||
      a.categoryRank-b.categoryRank ||
      num_(a.shift.Priority,50)-num_(b.shift.Priority,50) ||
      a.dateKey.localeCompare(b.dateKey) ||
      a.shiftCode.localeCompare(b.shiftCode) ||
      a.slot-b.slot
    );

    for (let i=0;i<remainingSlots.length;i++) {
      const slot=remainingSlots[i];
      const assigned = assignBestCandidate_(slot, model, state, ctx.username);
      if (assigned) slotResults.push(assigned);
      else slotResults.push(makeUnfilledAssignment_(slot, model, state, ctx.username));
    }

    runRepairPasses_(slotResults, model, state, ctx.username);

    // Regular pharmacists are primary staffing. Before optional backup coverage,
    // fill their required weekly workload from UNFILLED slots or PRN-held slots.
    rebalanceRegularWeeklyWorkload_(slotResults,model,state,ctx.username,start,end);

    // v15: OFF-day coverage skills live ONLY on the Users sheet. They are not
    // ordinary Employee Skills and therefore never make a pharmacist eligible
    // for a normal-day assignment. After the primary five-day/week pattern is
    // known, use those special coverage skills only when the primary/home
    // pharmacist has an algorithm-generated OFF day. A same-week non-protected
    // assignment may be swapped so the covering pharmacist still stays at five
    // workdays instead of receiving a sixth day.
    let coverageChanged=false;
    for (let coveragePass=0; coveragePass<2; coveragePass++) {
      if(Date.now()>model.runtimeDeadline)break;
      const changed=assignGeneratedOffDayCoverage_(slotResults,model,state,ctx.username,start,end);
      if(!changed)break;
      coverageChanged=true;
    }
    // One consolidated repair is enough after coverage swaps; v15 ran a full
    // repair after every coverage pass, which was the main timeout multiplier.
    if(coverageChanged && Date.now()<=model.runtimeDeadline) runRepairPasses_(slotResults,model,state,ctx.username);

    // A coverage repair may move assignments, so do one final regular-workload
    // pass before validation.
    rebalanceRegularWeeklyWorkload_(slotResults,model,state,ctx.username,start,end);

    refreshUnfilledReasons_(slotResults, model, state);
    applyGeneratedOffDayCoverageLabels_(slotResults, model, state, start, end);

    const postValidation = validateGeneratedAssignments_(slotResults, model, start, end, {skipPeriodHours:yes_(options.chunked)});
    const lockedNow = slotResults.filter(assignmentIsLocked_).map(r => lockedFingerprint_(r));
    const lockedErrors = compareLockedSnapshots_(lockedSnapshot, lockedNow);
    if (lockedErrors.length) throw new Error('Locked assignment protection failed: ' + lockedErrors.join('; '));

    saveSchedulePeriod_(existing, slotResults, start, end, mode);
    audit_('SCHEDULE_GENERATED', formatDateKey_(start), '', '', '', '', generationId, 'No', '',
      JSON.stringify({start:formatDateKey_(start),end:formatDateKey_(end),mode:mode,protectedPreAssignments:preserved.length,replacedRows:replaceable.length,filled:slotResults.filter(r=>r.status!=='UNFILLED').length,unfilled:slotResults.filter(r=>r.status==='UNFILLED').length}), ctx.username);

    const health = computeScheduleHealth_(slotResults.map(assignmentToScheduleObject_));
    return serialize_({
      ok:true,
      generationId:generationId,
      startDate:formatDateKey_(start),
      endDate:formatDateKey_(end),
      filled:health.filled,
      unfilled:health.unfilled,
      required:health.required,
      coveragePercent:health.coveragePercent,
      feasible:postValidation.errors.length === 0,
      message:postValidation.errors.length===0
        ? (health.unfilled===0
            ? 'Schedule generated successfully. Regular pharmacists meet the five-day weekly pattern and required weekend rotation.'
            : 'Regular pharmacists meet the five-day weekly pattern where feasible; some required shifts remain UNFILLED.')
        : 'Schedule generated, but one or more five-day, required-weekend, evening-to-morning transition, consecutive-day, skill, or regular-hours rules could not be satisfied.',
      warnings:cfg.warnings.concat(postValidation.warnings).concat(Date.now()>model.runtimeDeadline?['Performance safeguard: optional deep repair stopped early for this week-sized generation chunk. Required slots not solved by that point remain UNFILLED with an explanation.']:[]),
      errors:postValidation.errors,
      generationMs:Date.now()-generationStartedAt,
      exactHoursTarget:num_(model.settings.periodHoursTarget,320),
      pharmacistHourTotals:buildPeriodHourSummary_(slotResults,model,start,end),
      unfilledDetails:slotResults.filter(r=>r.status==='UNFILLED').map(r=>({date:r.dateKey,shift:r.shiftCode,slot:r.slot,reason:r.warning}))
    });
  } finally {
    lock.releaseLock();
  }
}

function loadSchedulingModel_() {
  const users = readTable_(APP.SHEETS.USERS).filter(r => clean_(r.Role).toLowerCase() !== 'administrator');
  const skills = readTable_(APP.SHEETS.SKILLS);
  const shifts = readTable_(APP.SHEETS.SHIFTS);
  const requirements = readTable_(APP.SHEETS.REQUIREMENTS);
  const requests = readTable_(APP.SHEETS.REQUESTS);
  const weeklyAvailability = readTable_(APP.SHEETS.WEEKLY_AVAILABILITY);
  const prnAvailabilitySourceLoaded = !!getDb_().getSheetByName('My Availability');
  const rawPrnAvailability = readTable_('My Availability');
  const rawSettings = getSettingsMap_();
  const shiftMap = {};
  shifts.forEach(sh => shiftMap[clean_(sh.Shift)] = sh);
  const usersByUsername = {}, usersByName = {}, usersByNameLower = {};
  users.forEach(u => {
    const username=clean_(u.Username);
    const pharmacistName=clean_(u['Pharmacist Name']);
    usersByUsername[username] = u;
    usersByName[pharmacistName] = u;
    usersByNameLower[pharmacistName.toLowerCase()] = u;
    const prefRaw=clean_(u['Preferred Shift Type']);
    u._preferredShiftTokens = new Set(prefRaw.toLowerCase().split(/[\\/,;|]+/).map(x=>x.trim()).filter(Boolean));
    u._preferredShiftCodeTokens = new Set(prefRaw.toUpperCase().split(/[\\/,;|]+/).map(x=>x.trim()).filter(Boolean));
    const offRaw=clean_(u['Off-Day Coverage Skills']);
    u._offDayCoverageSkillSet = new Set(offRaw.toUpperCase().split(/[\\/,;|]+/).map(x=>x.trim()).filter(Boolean));
  });
  const skillsByUser = {};
  skills.filter(r=>yesDefault_(r.Active,true)).forEach(r => {
    const username = clean_(r.Username) || (usersByName[clean_(r['Pharmacist Name'])] || {}).Username;
    if (!username) return;
    if (!skillsByUser[username]) skillsByUser[username] = new Set();
    skillsByUser[username].add(clean_(r.Skill));
  });
  const rotation = clean_(rawSettings['Weekend Rotation'] || 'A,B,C').split(',').map(x=>x.trim()).filter(Boolean);
  const settings = {
    raw: rawSettings,
    weeklyDefault: num_(rawSettings['Weekly Hours Limit'],40),
    periodHoursTarget: num_(rawSettings['Required Hours Per Schedule Period'],num_(rawSettings['Maximum Hours Per Schedule Period'],320)),
    periodHoursMax: num_(rawSettings['Required Hours Per Schedule Period'],num_(rawSettings['Maximum Hours Per Schedule Period'],320)),
    residentE2ShiftCode: clean_(rawSettings['Resident E2 Shift Code'] || 'E2').toUpperCase(),
    residentE2PerWeek: Math.max(0,Math.floor(num_(rawSettings['Resident E2 Shifts Per Week'],1))),
    sevenOnBlockScheduling: yesDefault_(rawSettings['SevenOn Block Scheduling'],true),
    eveningDefault: num_(rawSettings['Maximum Evening Shifts'],7),
    weekStart: clean_(rawSettings['Week Start'] || 'Sunday'),
    regularWorkdaysPerWeek: Math.max(1,Math.floor(num_(rawSettings['Regular Workdays Per Week'],5))),
    maxConsecutiveWorkdays: Math.max(1,Math.floor(num_(rawSettings['Maximum Consecutive Workdays'],5))),
    blockEveningToMorningTransition: yesDefault_(rawSettings['Block Evening To Morning Transition'],true),
    requiredWeekendAssignment: yesDefault_(rawSettings['Required Weekend Assignment'],true),
    residentsRequiredAssignedWeekend: yesDefault_(rawSettings['Residents Required Assigned Weekend'],true),
    preceptorWeekdayEveningAllowed: yes_(rawSettings['Preceptor Weekday Evening Allowed']),
    preceptorWeekendEveningAllowed: yesDefault_(rawSettings['Preceptor Weekend Evening Allowed'],true),
    requirePairedEdCoverage: yesDefault_(rawSettings['Require Paired ED Coverage'],true),
    edDayShiftCode: clean_(rawSettings['ED Day Shift Code'] || 'EDD').toUpperCase(),
    edEveningShiftCode: clean_(rawSettings['ED Evening Shift Code'] || 'EDE').toUpperCase(),
    weekendRotation: rotation.length ? rotation : ['A','B','C'],
    weekendAnchorDate: startOfDay_(asDate_(rawSettings['Weekend Anchor Date']) || nextOrSameSaturday_(new Date())),
    weekendAnchorGroup: clean_(rawSettings['Weekend Anchor Group'] || 'A'),
    allowWeekendFallback: yes_(rawSettings['Allow Weekend Fallback']),
    allowAdminOverride: yesDefault_(rawSettings['Allow Admin Rule Override'],true)
  };
  const weeklyAvailabilityByUser = {};
  weeklyAvailability.filter(r=>yesDefault_(r.Active,true)).forEach(r=>{
    const username=clean_(r.Username)||(usersByName[clean_(r['Pharmacist Name'])]||{}).Username;
    if(!username)return;
    if(!weeklyAvailabilityByUser[username])weeklyAvailabilityByUser[username]=[];
    weeklyAvailabilityByUser[username].push(r);
  });

  const prnAvailability=normalizePrnAvailabilityRows_(rawPrnAvailability,users);
  const prnAvailabilityByUser={};
  const prnAvailabilityDateCountByUser={};

  prnAvailability.forEach(r=>{
    const username=clean_(r.Username);
    if(!username)return;
    if(!prnAvailabilityByUser[username])prnAvailabilityByUser[username]=[];
    prnAvailabilityByUser[username].push(r);
  });

  Object.keys(prnAvailabilityByUser).forEach(username=>{
    prnAvailabilityDateCountByUser[username]=new Set(
      prnAvailabilityByUser[username]
        .filter(r=>yesDefault_(r.Available,true))
        .map(r=>clean_(r.Date))
        .filter(Boolean)
    ).size;
  });

  // v16 performance: index date/PTO rules by employee once instead of scanning
  // every request for every candidate/shift eligibility check.
  const requestsByUser={};
  users.forEach(u=>requestsByUser[clean_(u.Username)]=[]);
  requests.forEach(r=>{
    let username=clean_(r.Username);
    if(!username || !usersByUsername[username]){
      const requestName=clean_(r.Pharmacist);
      const matchedUser=usersByName[requestName]||usersByNameLower[requestName.toLowerCase()];
      username=clean_((matchedUser||{}).Username);
    }
    if(username && requestsByUser[username]) requestsByUser[username].push(r);
  });

  const model={
    users,skills,shifts,requirements,requests,requestsByUser,
    weeklyAvailability,weeklyAvailabilityByUser,
    prnAvailability,prnAvailabilityByUser,prnAvailabilityDateCountByUser,prnAvailabilitySourceLoaded,
    shiftMap,usersByUsername,usersByName,skillsByUser,settings,
    _ptoCache:{},_regularOffCache:{},_availabilityCache:{},_weeklyAvailabilityCache:{},_prnAvailabilityCache:{},_staticCandidateCache:{},
    runtimeDeadline:0
  };
  model.activeUsers=users.filter(u=>yes_(u.Active));

  // Performance index used by Help Fill Open Shifts:
  // required skill -> only active pharmacists who possess that skill.
  // This prevents the helper from running eligibility logic against every
  // pharmacist for every open shift.
  model.activeUsersBySkill={};
  model.activeUsersByOffDaySkill={};

  model.activeUsers.forEach(u=>{
    const username=clean_(u.Username);
    const normalSkills=skillsByUser[username]||new Set();

    normalSkills.forEach(skill=>{
      const key=clean_(skill).toUpperCase();
      if(!key)return;
      if(!model.activeUsersBySkill[key])model.activeUsersBySkill[key]=[];
      model.activeUsersBySkill[key].push(u);
    });

    offDayCoverageSkillSet_(u).forEach(skill=>{
      const key=clean_(skill).toUpperCase();
      if(!key)return;
      if(!model.activeUsersByOffDaySkill[key])model.activeUsersByOffDaySkill[key]=[];
      model.activeUsersByOffDaySkill[key].push(u);
    });
  });

  model.regularUsers=model.activeUsers.filter(u=>regularFiveDayRuleApplies_(u));
  model.residents=model.activeUsers.filter(u=>yes_(u.Resident));
  model.sevenOnShiftByUser=buildSevenOnShiftMap_(model);
  return model;
}

function createRequiredSlots_(model,start,end,generationId) {
  const reqMap = {};
  model.requirements.filter(r=>yesDefault_(r.Active,true)).forEach(r=>reqMap[clean_(r.Shift)] = r);
  const slots = [];
  for (let d=new Date(start); d<=end; d=addDays_(d,1)) {
    const dayName = dayName_(d);
    Object.keys(reqMap).forEach(code => {
      const req = reqMap[code];
      const count = Math.max(0,Math.floor(num_(req[dayName],0)));
      const shift = model.shiftMap[code];
      if (!shift || !yes_(shift.Active) || count < 1) return;
      for (let slot=1;slot<=count;slot++) {
        const dk = formatDateKey_(d);
        slots.push({
          generationId:generationId,date:new Date(d),dateKey:dk,day:dayName,shiftCode:code,slot:slot,shift:shift,
          requiredSkill:clean_(shift.Skill),weekend:isWeekendDate_(d),weekendGroup:isWeekendDate_(d)?weekendGroupForDate_(d,model.settings):'',
          slotKey:slotKey_(dk,code,slot)
        });
      }
    });
  }
  return slots;
}

function countStaticCandidates_(slot, model) {
  const cacheKey=slot.dateKey+'|'+slot.shiftCode;
  if(model._staticCandidateCache && model._staticCandidateCache[cacheKey]!==undefined) return model._staticCandidateCache[cacheKey];
  let n=0;
  skillQualifiedUsersForSlot_(slot,model).forEach(u => {
    if (!yes_(u.Active)) return;
    if (slot.weekend && !yesDefault_(u['Weekend Eligible'],true)) return;
    if (preceptorEveningBlocked_(u,slot,model)) return;
    if (isSevenOn_(u) && (!sevenOnIsOnDay_(u,slot.date) || !sevenOnSlotMatches_(u,slot,model))) return;
    const wa=weeklyAvailabilityStatus_(u,slot.date,slot.shift,model);
    if(wa.blocked)return;
    n++;
  });
  if(model._staticCandidateCache) model._staticCandidateCache[cacheKey]=n;
  return n;
}

function shiftCategoryRank_(shift) {
  const t = clean_(shift.Type).toLowerCase();
  if (t === 'night') return 1;
  if (yes_(shift.Weekend)) return 2;
  if (['cc1','cc2','onc','edd','ede','card','im'].includes(clean_(shift.Shift).toLowerCase())) return 3;
  if (t === 'evening') return 4;
  return 5;
}

function createState_(model, preservedRows, periodStart, periodEnd) {
  const state = {
    assignments: [],
    byEmployeeDate: {},
    assignedDateKeysByEmployee: {},
    intervalsByEmployee: {},
    weeklyHours: {},
    weeklyDays: {},
    monthlyEvenings: {},
    totalAssignments: {},
    totalHours: {},
    periodHours: {},
    periodStart: periodStart ? startOfDay_(periodStart) : null,
    periodEnd: periodEnd ? startOfDay_(periodEnd) : null,
    weekendAssignments: {},
    nightAssignments: {},
    residentE2ByWeek: {},
    lastTypeByEmployeeDate: {},
    weekendRoleOwner: {}
  };
  (preservedRows||[]).forEach(r => {
    const a = scheduleRowToAssignment_(r,model);
    if (a && a.username && a.status !== 'UNFILLED') addAssignmentToState_(state,a,model);
  });
  return state;
}

function scheduleRowToAssignment_(r, model) {
  const d = asDate_(r.Date);
  const shift = model.shiftMap[clean_(r.Shift)];
  if (!d || !shift) return null;
  return {
    generationId: clean_(r['Generation ID']),
    assignmentId: clean_(r['Assignment ID']) || Utilities.getUuid(),
    date:d,dateKey:formatDateKey_(d),day:dayName_(d),shiftCode:clean_(r.Shift),slot:num_(r.Slot,1),shift:shift,
    pharmacist:clean_(r['Assigned Pharmacist']),username:clean_(r.Username),hours:num_(r.Hours,num_(shift.Hours,0)),creditedHours:num_(r['Credited Hours'],creditedHours_(shift)),
    requiredSkill:clean_(r['Required Skill'] || shift.Skill),coverageForPharmacist:clean_(r['Coverage For Pharmacist']),coverageForUsername:clean_(r['Coverage For Username']),coverageReason:clean_(r['Coverage Reason']),shiftType:clean_(r['Shift Type'] || shift.Type),weekend:yes_(r.Weekend),weekendGroup:clean_(r['Weekend Group']),
    holiday:clean_(r.Holiday),locked:yes_(r.Locked),manual:yes_(r.Manual),status:clean_(r.Status || 'ASSIGNED'),warning:clean_(r.Warning),
    updatedAt:asDate_(r['Updated At']) || new Date(),updatedBy:clean_(r['Updated By']),finalizedAt:asDate_(r['Finalized At'])
  };
}

function assignBestCandidate_(slot, model, state, actor) {
  const candidates = [];
  model.users.forEach(u => {
    const e = eligibility_(u,slot,model,state,false);
    if (!e.ok) return;
    candidates.push({user:u, score:scoreCandidate_(u,slot,model,state,e), warnings:e.warnings});
  });
  if (!candidates.length) return null;
  candidates.sort((a,b)=>b.score-a.score || clean_(a.user.Username).localeCompare(clean_(b.user.Username)));
  const pick = candidates[0];
  const a = makeAssigned_(slot,pick.user,pick.warnings,actor);
  addAssignmentToState_(state,a,model);
  return a;
}

function prnAvailabilityStatus_(u,date,shift,model) {
  if(!isPrnEmployee_(u)){
    return {
      blocked:false,
      preferred:false,
      reason:'',
      availableDateCount:0,
      matchingRows:[]
    };
  }

  // PRN is opt-in only. If My Availability has not loaded, or the PRN has
  // not submitted availability for this date/shift, the algorithm must NOT
  // schedule the PRN. This prevents PRNs from becoming fallback regular staff.
  if(!model||!model.prnAvailabilitySourceLoaded){
    return {
      blocked:true,
      preferred:false,
      reason:'PRN_AVAILABILITY_NOT_LOADED',
      availableDateCount:0,
      matchingRows:[]
    };
  }

  const username=clean_(u.Username);
  const dk=formatDateKey_(date);
  const shiftCode=clean_(shift&&shift.Shift).toUpperCase();
  const shiftType=clean_(shift&&shift.Type).toUpperCase();
  const cacheKey=username+'|'+dk+'|'+shiftCode;

  if(model._prnAvailabilityCache&&model._prnAvailabilityCache[cacheKey]){
    return model._prnAvailabilityCache[cacheKey];
  }

  const rows=(model.prnAvailabilityByUser&&model.prnAvailabilityByUser[username])||[];
  const onDate=rows.filter(r=>clean_(r.Date)===dk);

  function rowMatchesShift_(r){
    if(!yesDefault_(r.Available,true))return false;

    const requested=clean_(r.Shift).toUpperCase();
    if(requested){
      const tokens=requested.split(/[,;|\/\s]+/).map(x=>x.trim()).filter(Boolean);
      const generic=tokens.some(x=>x==='ANY'||x==='ALL'||x==='OPEN'||x==='AVAILABLE');
      if(!generic&&!tokens.includes(shiftCode)&&!tokens.includes(shiftType))return false;
    }

    const rs=timeMinutes_(r['Start Time']);
    const re=timeMinutes_(r['End Time']);
    if(rs!==null||re!==null){
      const ss=timeMinutes_(shift&&shift.Start);
      const se=timeMinutes_(shift&&shift.End);
      if(rs===null||re===null||ss===null||se===null)return false;
      if(!intervalContains_(rs,re,ss,se))return false;
    }

    return true;
  }

  const matching=onDate.filter(rowMatchesShift_);
  const result={
    blocked:matching.length===0,
    preferred:matching.length>0,
    reason:matching.length?'':'PRN_NOT_AVAILABLE',
    availableDateCount:num_(
      model.prnAvailabilityDateCountByUser&&model.prnAvailabilityDateCountByUser[username],
      0
    ),
    matchingRows:matching
  };

  if(model._prnAvailabilityCache)model._prnAvailabilityCache[cacheKey]=result;
  return result;
}

function eligibility_(u, slot, model, state, manualMode) {
  const reasons=[], warnings=[];
  const username = clean_(u.Username);
  if (!yes_(u.Active)) reasons.push('INACTIVE');
  if (!hasRequiredSkillForSlot_(u,slot,model)) reasons.push('MISSING_SKILL');
  if (isBlockedByPto_(u,slot.date,model)) reasons.push('PTO');
  if (isBlockedByRegularOff_(u,slot.date,model)) reasons.push('REGULAR_OFF');
  const avail = availabilityStatus_(u,slot.date,model);
  if (avail.blocked) reasons.push('UNAVAILABLE');
  const weeklyAvail = weeklyAvailabilityStatus_(u,slot.date,slot.shift,model);
  if (weeklyAvail.blocked) reasons.push(weeklyAvail.reason || 'WEEKLY_AVAILABILITY');

  const prnAvail = prnAvailabilityStatus_(u,slot.date,slot.shift,model);
  if (prnAvail.blocked) reasons.push(prnAvail.reason || 'PRN_NOT_AVAILABLE');

  if (state.byEmployeeDate[username+'|'+slot.dateKey]) reasons.push('ALREADY_SCHEDULED');
  if (hasTimeConflict_(username,slot,state)) reasons.push('TIME_CONFLICT');
  // v25 hard transition rule: an Evening shift may be followed by OFF or
  // another Evening shift, but never by a Day/Morning shift on the next
  // calendar day. The check is bidirectional so it also protects a required
  // Day/Morning shift tomorrow from an Evening assignment today.
  if (model.settings.blockEveningToMorningTransition && wouldCreateEveningToMorningTransition_(username,slot,state)) {
    reasons.push('EVENING_TO_MORNING');
  }

  const weekKey = username+'|'+weekStartKey_(slot.date,model.settings.weekStart);
  const newHours = num_(state.weeklyHours[weekKey],0) + creditedHours_(slot.shift);
  const maxHours = employeeWeeklyMax_(u,model);
  // A true 7-on/7-off block intentionally exceeds the ordinary 40-hour weekly rule.
  // The block rule takes precedence so the employee is not split across random days.
  if (!isSevenOn_(u) && newHours > maxHours + 0.0001) reasons.push('WEEKLY_HOURS');

  // v12: regular pharmacists work EXACTLY five days in a Sunday-Saturday week.
  // Eligibility enforces the hard upper bound; final validation enforces the
  // corresponding lower bound (exactly five, not four).
  if (regularFiveDayRuleApplies_(u)) {
    const currentDays=num_(state.weeklyDays[weekKey],0);
    const requiredDays=regularRequiredWorkdaysForWeek_(u,slot.date,model);
    if (currentDays + 1 > requiredDays) reasons.push('WEEKLY_DAYS');
    if (wouldExceedConsecutiveDays_(username,slot.date,state,model.settings.maxConsecutiveWorkdays)) reasons.push('CONSECUTIVE_DAYS');
  }

  // Hard 2-month / generated-period cap.
  // Uses Credited Hours (or shift Hours when Credited Hours is blank).
  if (
    state.periodStart &&
    state.periodEnd &&
    inDateRange_(slot.date, state.periodStart, state.periodEnd)
  ) {
    const proposedPeriodHours =
      num_(state.periodHours[username],0) +
      creditedHours_(slot.shift);

    if (
      !isSevenOn_(u) &&
      proposedPeriodHours >
      num_(model.settings.periodHoursTarget,320) +
      0.0001
    ) {
      reasons.push('TWO_MONTH_HOURS');
    }
  }

  const type = clean_(slot.shift.Type).toLowerCase();
  const residentE2=isResidentMandatoryE2_(u,slot,model);
  const sevenDedicated=isSevenOn_(u) && sevenOnSlotMatches_(u,slot,model);
  if (type === 'evening') {
    if (preceptorEveningBlocked_(u,slot,model)) reasons.push('PRECEPTOR_WEEKDAY_EVENING');
    if (!yesDefault_(u['Evening Eligible'],true) && !residentE2) reasons.push('EVENING_NOT_ELIGIBLE');
    const mk = username+'|'+monthKey_(slot.date);
    // Dedicated 7-on E employees and the mandatory resident E2 shift are not
    // broken by the ordinary monthly evening-shift cap.
    if (!sevenDedicated && !residentE2 && num_(state.monthlyEvenings[mk],0) + 1 > employeeEveningMax_(u,model)) reasons.push('EVENING_LIMIT');
  }
  if (type === 'night' && !yesDefault_(u['Night Eligible'],true)) reasons.push('NIGHT_NOT_ELIGIBLE');

  if (residentE2) {
    const e2wk=username+'|'+weekStartKey_(slot.date,model.settings.weekStart);
    if (num_(state.residentE2ByWeek[e2wk],0) >= model.settings.residentE2PerWeek) reasons.push('RESIDENT_E2_WEEKLY_LIMIT');
  }

  if (slot.weekend || isWeekendDate_(slot.date)) {
    if (!yesDefault_(u['Weekend Eligible'],true)) reasons.push('WEEKEND_NOT_ELIGIBLE');
    const expected = slot.weekendGroup || weekendGroupForDate_(slot.date,model.settings);
    const actual = clean_(u['Weekend Group']);
    if (isSevenOn_(u)) {
      // 7-on/7-off rotations are controlled by their 14-day anchor, not A/B/C weekends.
    } else if (isPrnEmployee_(u)) {
      // PRN pharmacists are governed by My Availability rather than the A/B/C
      // weekend rotation. Weekend Eligible still applies above.
    } else if (yes_(u.Resident)) {
      // Residents may only cover the weekend assigned to their A/B/C group,
      // and the employee-specific 21-day weekend anchor must also be ON.
      if (!actual || actual !== expected) reasons.push('RESIDENT_WRONG_WEEKEND_GROUP');
      else if (!employeeWeekendIsOn_(u,slot.date,model.settings)) reasons.push('WEEKEND_ANCHOR_OFF');
    } else if (expected && actual !== expected) {
      if (model.settings.allowWeekendFallback) warnings.push('WEEKEND ROTATION EXCEPTION requested, but the employee-specific 21-day weekend anchor is a hard rule: expected Group '+expected+', employee Group '+(actual||'blank')+'.');
      else reasons.push('WRONG_WEEKEND_GROUP');
      // The new per-employee anchor is intentionally hard. Even when the old
      // fallback setting is Yes, a weekend-team pharmacist cannot be pulled
      // into either of the two OFF weekends in the 3-week cycle.
      reasons.push('WEEKEND_ANCHOR_OFF');
    } else if (expected && actual === expected && !employeeWeekendIsOn_(u,slot.date,model.settings)) {
      reasons.push('WEEKEND_ANCHOR_OFF');
    }
  }

  if (isSevenOn_(u)) {
    if (!sevenOnIsOnDay_(u,slot.date)) reasons.push('SEVEN_OFF');
    else if (!sevenOnSlotMatches_(u,slot,model)) reasons.push('SEVEN_WRONG_SHIFT');
  }
  const residentReason = residentEligibilityReason_(u,slot.shift,slot.date,model);
  if (residentReason) reasons.push(residentReason);

  if (yes_(u['Custom Hours Enabled']) && clean_(u['Custom Hours Mode']).toUpperCase()==='HARD' && !shiftFitsPreferredHours_(slot.shift,u)) reasons.push('CUSTOM_HOURS');

  if (manualMode && reasons.length) return {ok:false,reasons:reasons,warnings:warnings};
  return {
    ok:reasons.length===0,
    reasons:reasons,
    warnings:warnings,
    availabilityPreferred:avail.preferred,
    weeklyPreferred:weeklyAvail.preferred,
    weeklySoftMismatch:weeklyAvail.softMismatch,
    prnAvailabilityPreferred:prnAvail.preferred,
    prnAvailabilityCount:prnAvail.availableDateCount
  };
}


function preceptorEveningBlocked_(u,slot,model) {
  if (!u || !slot || !slot.shift || !yes_(u.Preceptor)) return false;
  if (clean_(slot.shift.Type).toLowerCase() !== 'evening') return false;
  if (isResidentMandatoryE2_(u,slot,model)) return false;
  if (isWeekendDate_(slot.date)) return !model.settings.preceptorWeekendEveningAllowed;
  return !model.settings.preceptorWeekdayEveningAllowed;
}

/**
 * v25 Evening -> Morning/Day transition protection.
 *
 * Allowed:
 *   Evening -> Evening
 *   Day     -> Evening
 *   Evening -> OFF
 *
 * Blocked:
 *   Evening -> Day/Morning on the next calendar day
 *
 * We use Shift.Type = 'Day' as the morning/day family. This matches the
 * existing scheduler classification for early shifts such as C7, IM, CARD,
 * CC1/CC2, EDD, WD1/WD2, etc. The check is bidirectional because mandatory
 * weekend/day assignments may already exist in state before an evening shift
 * is considered.
 */
function wouldCreateEveningToMorningTransition_(username,slot,state) {
  username=clean_(username);
  if(!username||!slot||!slot.date||!slot.shift||!state)return false;
  const proposedType=clean_(slot.shift.Type).toLowerCase();
  if(proposedType!=='day' && proposedType!=='evening')return false;
  const prevKey=username+'|'+formatDateKey_(addDays_(slot.date,-1));
  const nextKey=username+'|'+formatDateKey_(addDays_(slot.date,1));
  const prev=state.byEmployeeDate[prevKey];
  const next=state.byEmployeeDate[nextKey];
  const prevType=prev?clean_(prev.shiftType || (prev.shift||{}).Type).toLowerCase():'';
  const nextType=next?clean_(next.shiftType || (next.shift||{}).Type).toLowerCase():'';

  // Yesterday evening -> proposed morning/day today.
  if(proposedType==='day' && prevType==='evening')return true;
  // Proposed evening today -> already-reserved morning/day tomorrow.
  if(proposedType==='evening' && nextType==='day')return true;
  return false;
}

function preferredShiftMatches_(u,slot) {
  if(!u||!slot)return false;
  const raw=clean_(u['Preferred Shift Type']).toLowerCase();
  if(!raw)return false;
  const code=clean_(slot.shiftCode || (slot.shift||{}).Shift).toLowerCase();
  const type=clean_((slot.shift||{}).Type).toLowerCase();
  const tokens=u._preferredShiftTokens || new Set(raw.split(/[\\/,;|]+/).map(x=>x.trim()).filter(Boolean));
  return tokens.has(code) || tokens.has(type) || raw===code || raw===type;
}


/**
 * True only when Preferred Shift Type explicitly names the shift CODE.
 * Generic values such as "Day" are not role ownership.
 */
function preferredShiftCodeMatches_(u,slot) {
  if(!u||!slot)return false;
  const raw=clean_(u['Preferred Shift Type']).toUpperCase();
  if(!raw)return false;
  const code=clean_(slot.shiftCode || (slot.shift||{}).Shift).toUpperCase();
  if(!code)return false;
  const tokens=u._preferredShiftCodeTokens || new Set(raw.split(/[\\/,;|]+/).map(x=>x.trim()).filter(Boolean));
  return tokens.has(code);
}

/**
 * Coverage-only qualifications are stored on Users -> Off-Day Coverage Skills.
 * Example: IM,CARD,CC1. These values are intentionally NOT merged into
 * Employee Skills and are ignored by the ordinary scheduler.
 */
function offDayCoverageSkillSet_(u) {
  if(!u)return new Set();
  if(u._offDayCoverageSkillSet instanceof Set)return u._offDayCoverageSkillSet;
  const raw=clean_(u['Off-Day Coverage Skills']).toUpperCase();
  u._offDayCoverageSkillSet=new Set(raw.split(/[\\/,;|]+/).map(x=>x.trim()).filter(Boolean));
  return u._offDayCoverageSkillSet;
}

function hasOffDayCoverageSkillForSlot_(u,slot) {
  if(!u||!slot)return false;
  const skills=offDayCoverageSkillSet_(u);
  if(!skills.size)return false;
  const codes=[
    clean_(slot.shiftCode || (slot.shift||{}).Shift).toUpperCase(),
    clean_(slot.requiredSkill).toUpperCase(),
    clean_((slot.shift||{}).Skill).toUpperCase()
  ].filter(Boolean);
  return codes.some(c=>skills.has(c));
}

/**
 * Apply every normal hard rule, but replace the ordinary Employee Skills check
 * with the Users-sheet OFF-day coverage qualification. This is the only place
 * where Off-Day Coverage Skills can grant eligibility.
 */
function offDayCoverageEligibility_(u,slot,model,state) {
  const base=eligibility_(u,slot,model,state,false);
  const reasons=(base.reasons||[]).filter(r=>r!=='MISSING_SKILL');
  if(!hasOffDayCoverageSkillForSlot_(u,slot))reasons.push('MISSING_OFF_DAY_COVERAGE_SKILL');
  return {
    ok:reasons.length===0,
    reasons:reasons,
    warnings:(base.warnings||[]).slice(),
    availabilityPreferred:base.availabilityPreferred,
    weeklyPreferred:base.weeklyPreferred,
    weeklySoftMismatch:base.weeklySoftMismatch
  };
}

function primaryHomeOwnersForSlot_(slot,model) {
  if(!slot||!model)return [];
  return model.users
    .filter(u=>yes_(u.Active)&&regularFiveDayRuleApplies_(u)&&preferredShiftCodeMatches_(u,slot))
    .sort((a,b)=>clean_(a['Pharmacist Name']).localeCompare(clean_(b['Pharmacist Name'])));
}

function hardUnavailableForHomeShift_(u,slot,model) {
  if(!u||!slot)return true;
  if(isBlockedByPto_(u,slot.date,model))return true;
  const av=availabilityStatus_(u,slot.date,model);
  if(av.blocked)return true;
  const wa=weeklyAvailabilityStatus_(u,slot.date,slot.shift,model);
  if(wa.blocked)return true;
  return false;
}

/**
 * A generated OFF day means the primary pharmacist has no assignment on this
 * date, is not on PTO/hard-unavailable, and already has the required five
 * workdays in this Sunday-Saturday week.
 */
function isGeneratedOffDayForHomeOwner_(u,slot,state,model) {
  if(!u||!slot||!state||!regularFiveDayRuleApplies_(u))return false;
  const username=clean_(u.Username);
  if(!username)return false;
  if(state.byEmployeeDate[username+'|'+slot.dateKey])return false;
  if(hardUnavailableForHomeShift_(u,slot,model))return false;
  const wk=username+'|'+weekStartKey_(slot.date,model.settings.weekStart);
  return num_(state.weeklyDays[wk],0)>=model.settings.regularWorkdaysPerWeek;
}

/**
 * v24: Resolve a manual assignment that is allowed ONLY because the selected
 * pharmacist has the shift in Users -> Off-Day Coverage Skills.
 *
 * The coverage-only skill is never treated as a normal Employee Skill. It is
 * accepted only when all of the following are true:
 *   1) the selected pharmacist does not already have the normal required skill;
 *   2) the selected pharmacist has the shift/required skill in Off-Day Coverage Skills;
 *   3) exactly one primary/home owner exists for the shift;
 *   4) that owner has an algorithm-generated OFF day on this date; and
 *   5) the covering pharmacist passes every other hard scheduling rule.
 */
function resolveManualOffDayCoverageContext_(u,slot,model,state) {
  const out={
    applies:false,
    coverageForPharmacist:'',
    coverageForUsername:'',
    coverageReason:'',
    message:'',
    eligibility:null
  };
  if(!u||!slot||!model||!state)return out;

  // If the pharmacist has the ordinary Employee Skill, this is a normal
  // assignment and must not be labeled as OFF-day-only coverage.
  if(hasRequiredSkillForSlot_(u,slot,model))return out;
  if(!hasOffDayCoverageSkillForSlot_(u,slot))return out;

  const owners=primaryHomeOwnersForSlot_(slot,model)
    .filter(owner=>clean_(owner.Username)!==clean_(u.Username));
  if(!owners.length){
    out.message='This pharmacist has an OFF-day coverage skill, but there is no primary/home pharmacist configured for this shift.';
    return out;
  }

  // A shift may have more than one normal owner (for example the paired ED
  // EDD/EDE pharmacists). Manual OFF-day coverage is still safe when exactly
  // one of those owners is actually on an algorithm-generated OFF day.
  const offOwners=owners.filter(owner=>isGeneratedOffDayForHomeOwner_(owner,slot,state,model));
  if(offOwners.length!==1){
    out.message=offOwners.length===0
      ? 'OFF-day coverage is configured, but none of the primary/home pharmacists has an algorithm-generated OFF day for this shift/date.'
      : 'More than one primary/home pharmacist is OFF for this shift/date, so the system cannot determine which pharmacist is being covered automatically.';
    return out;
  }
  const owner=offOwners[0];

  const e=offDayCoverageEligibility_(u,slot,model,state);
  out.eligibility=e;
  if(!e.ok){
    out.message='The OFF-day coverage skill is configured, but another hard scheduling rule prevents this assignment.';
    return out;
  }

  out.applies=true;
  out.coverageForPharmacist=clean_(owner['Pharmacist Name']);
  out.coverageForUsername=clean_(owner.Username);
  out.coverageReason='GENERATED OFF DAY';
  out.message='Covering '+out.coverageForPharmacist+' on an algorithm-generated OFF day.';
  return out;
}

/**
 * Trace substitute coverage for one clearly-defined specialty owner.
 * If multiple employees share the same preferred role (for example EDD/EDE),
 * coverage is not inferred automatically because there is no single owner.
 */
function applyGeneratedOffDayCoverageLabels_(results,model,state,start,end) {
  (results||[]).forEach(a=>{
    if(!a)return;
    a.coverageForPharmacist='';
    a.coverageForUsername='';
    a.coverageReason='';
    if(a.status==='UNFILLED'||!a.username||!inDateRange_(a.date,start,end))return;
    if(isPatternProtectedAssignment_(a))return;
    const slot=assignmentAsSlot_(a);
    const owners=primaryHomeOwnersForSlot_(slot,model);
    if(owners.length!==1)return;
    const owner=owners[0];
    if(clean_(owner.Username)===clean_(a.username))return;
    if(!isGeneratedOffDayForHomeOwner_(owner,slot,state,model))return;
    const coveringUser=model.usersByUsername[clean_(a.username)];
    // Only an explicit Users-sheet OFF-day coverage skill receives this label.
    if(!coveringUser||!hasOffDayCoverageSkillForSlot_(coveringUser,slot))return;
    a.coverageForPharmacist=clean_(owner['Pharmacist Name']);
    a.coverageForUsername=clean_(owner.Username);
    a.coverageReason='GENERATED OFF DAY';
  });
}

function makeOffDayCoverageAssignment_(slot,u,owner,model,state,actor,note) {
  const e=offDayCoverageEligibility_(u,slot,model,state);
  if(!e.ok)return null;
  const warnings=(e.warnings||[]).slice();
  warnings.push('OFF-DAY COVERAGE ONLY');
  if(note)warnings.push(note);
  const a=makeAssigned_(slot,u,warnings,actor);
  a.coverageForPharmacist=clean_(owner['Pharmacist Name']);
  a.coverageForUsername=clean_(owner.Username);
  a.coverageReason='GENERATED OFF DAY';
  addAssignmentToState_(state,a,model);
  return a;
}

function offDayCoverageCandidates_(slot,owner,model,state) {
  const out=[];
  model.users.forEach(u=>{
    if(!yes_(u.Active))return;
    if(clean_(u.Username)===clean_(owner.Username))return;
    if(!hasOffDayCoverageSkillForSlot_(u,slot))return;
    const e=offDayCoverageEligibility_(u,slot,model,state);
    if(e.ok)out.push({user:u,elig:e,score:scoreCandidate_(u,slot,model,state,e)+1800});
  });
  out.sort((a,b)=>b.score-a.score||clean_(a.user.Username).localeCompare(clean_(b.user.Username)));
  return out;
}

function sameSchedulingWeek_(a,b,weekStart) {
  return !!a&&!!b&&weekStartKey_(a,weekStart)===weekStartKey_(b,weekStart);
}

/**
 * If a covering pharmacist is already at five workdays, move one ordinary,
 * non-protected assignment in the same Sunday-Saturday week to the OFF-day
 * specialty. This keeps the coverer at five days instead of creating day six.
 */
function trySwapForOffDayCoverage_(unfilled,slot,owner,coverer,results,model,state,actor) {
  const username=clean_(coverer.Username);
  const donors=(results||[]).filter(a=>
    a&&a.status!=='UNFILLED'&&a.username===username&&
    !a.locked&&!a.manual&&!isPatternProtectedAssignment_(a)&&
    clean_(a.coverageReason).toUpperCase()!=='GENERATED OFF DAY'&&
    sameSchedulingWeek_(a.date,slot.date,model.settings.weekStart)
  ).sort((a,b)=>{
    const asame=a.dateKey===slot.dateKey?0:1, bsame=b.dateKey===slot.dateKey?0:1;
    if(asame!==bsame)return asame-bsame;
    const ahome=preferredShiftCodeMatches_(coverer,assignmentAsSlot_(a))?1:0;
    const bhome=preferredShiftCodeMatches_(coverer,assignmentAsSlot_(b))?1:0;
    if(ahome!==bhome)return ahome-bhome;
    return a.dateKey.localeCompare(b.dateKey)||a.shiftCode.localeCompare(b.shiftCode);
  });

  for(const donor of donors){
    removeAssignmentFromState_(state,donor,model);
    const e=offDayCoverageEligibility_(coverer,slot,model,state);
    if(!e.ok){
      addAssignmentToState_(state,donor,model);
      continue;
    }

    const moved=makeAssigned_(slot,coverer,(e.warnings||[]).concat([
      'OFF-DAY COVERAGE ONLY',
      'OFF-DAY COVERAGE SWAP: moved from '+donor.shiftCode+' to cover '+clean_(owner['Pharmacist Name'])+'.'
    ]),actor);
    moved.coverageForPharmacist=clean_(owner['Pharmacist Name']);
    moved.coverageForUsername=clean_(owner.Username);
    moved.coverageReason='GENERATED OFF DAY';
    addAssignmentToState_(state,moved,model);

    replaceAssignmentObject_(unfilled,moved);
    const vacatedSlot=assignmentAsSlot_(donor);
    const replacement=assignBestCandidate_(vacatedSlot,model,state,actor);
    if(replacement) replaceAssignmentObject_(donor,replacement);
    else replaceAssignmentObject_(donor,makeUnfilledAssignment_(vacatedSlot,model,state,actor));
    return true;
  }
  return false;
}

/**
 * Use Users -> Off-Day Coverage Skills only when the single primary/home
 * pharmacist for the shift has a genuine algorithm-generated OFF day.
 */
function assignGeneratedOffDayCoverage_(results,model,state,actor,start,end) {
  let changed=false;
  const deadline=num_(model.runtimeDeadline,0);
  const targets=(results||[]).filter(a=>a&&a.status==='UNFILLED'&&inDateRange_(a.date,start,end));
  const coverers=(model.activeUsers||model.users).filter(u=>yes_(u.Active));
  for(const unfilled of targets){
    if(deadline>0 && Date.now()>deadline)break;
    const slot=assignmentAsSlot_(unfilled);
    const owners=primaryHomeOwnersForSlot_(slot,model);
    if(owners.length!==1)continue;
    const owner=owners[0];
    if(!isGeneratedOffDayForHomeOwner_(owner,slot,state,model))continue;

    const direct=offDayCoverageCandidates_(slot,owner,model,state);
    if(direct.length){
      const a=makeOffDayCoverageAssignment_(slot,direct[0].user,owner,model,state,actor,
        'Covering '+clean_(owner['Pharmacist Name'])+' on an algorithm-generated OFF day.');
      if(a){replaceAssignmentObject_(unfilled,a);changed=true;continue;}
    }

    const possible=coverers
      .filter(u=>clean_(u.Username)!==clean_(owner.Username)&&hasOffDayCoverageSkillForSlot_(u,slot))
      .sort((a,b)=>clean_(a['Pharmacist Name']).localeCompare(clean_(b['Pharmacist Name'])));
    for(const coverer of possible){
      if(deadline>0 && Date.now()>deadline)break;
      if(trySwapForOffDayCoverage_(unfilled,slot,owner,coverer,results,model,state,actor)){
        changed=true;
        break;
      }
    }
  }
  return changed;
}

function candidateListForSlot_(slot,model,state) {
  const out=[];
  model.users.forEach(u=>{
    const e=eligibility_(u,slot,model,state,false);
    if(!e.ok)return;
    out.push({user:u,score:scoreCandidate_(u,slot,model,state,e),warnings:e.warnings||[]});
  });
  out.sort((a,b)=>b.score-a.score||clean_(a.user.Username).localeCompare(clean_(b.user.Username)));
  return out;
}

function scoreCandidate_(u,slot,model,state,elig) {
  const username=clean_(u.Username);
  let score=1000;
  const pref=clean_(u['Preferred Shift Type']).toLowerCase();
  const type=clean_(slot.shift.Type).toLowerCase();
  if (preferredShiftMatches_(u,slot)) score += 40;
  // Exact preferred shift codes define the employee's primary/home role.
  // This keeps specialty owners (IM, CC1, CC2, ONC, etc.) on their own role
  // on normal workdays, while another qualified pharmacist can cover the role
  // on an algorithm-generated OFF day.
  if (preferredShiftCodeMatches_(u,slot)) score += 1400;
  if (yes_(u['Custom Hours Enabled']) && shiftFitsPreferredHours_(slot.shift,u)) score += 25;
  if (elig.availabilityPreferred) score += 20;
  if (elig.weeklyPreferred) score += 35;
  if (elig.weeklySoftMismatch) score -= 25;

  const wk = username+'|'+weekStartKey_(slot.date,model.settings.weekStart);
  if (regularFiveDayRuleApplies_(u)) {
    const currentDays=num_(state.weeklyDays[wk],0);
    const requiredDays=regularRequiredWorkdaysForWeek_(u,slot.date,model);
    const deficit=Math.max(0,requiredDays-currentDays);

    // Regular pharmacists are the primary workforce. An eligible regular who
    // still owes workdays this week must rank above PRN backup coverage.
    score += 100000;
    score += deficit*10000;
    if (deficit===1) score += 5000;
  }
  const currentWeek = num_(state.weeklyHours[wk],0);

  if(isPrnEmployee_(u)){
    const availableCount=num_(
      model.prnAvailabilityDateCountByUser&&model.prnAvailabilityDateCountByUser[username],
      0
    );
    const assignedCount=num_(state.totalAssignments[username],0);

    // PRN is backup coverage only. Availability makes the PRN eligible, but
    // never gives them priority over a regular pharmacist who still needs work.
    score -= 100000;
    score += Math.max(0,availableCount-assignedCount)*10;
    score -= assignedCount*35;
    if(elig.prnAvailabilityPreferred)score += 100;
  }else{
    const target = num_(u['Target Weekly Hours'], num_(model.settings.raw['Default Regular Weekly Hours'],40));
    score += Math.max(-40, Math.min(60,(target-currentWeek)*2));
  }

  // Exact 320-hour schedule-period objective applies to regular pharmacists,
  // not PRN pharmacists whose schedule is driven by My Availability.
  if (
    !isPrnEmployee_(u) &&
    state.periodStart &&
    state.periodEnd &&
    inDateRange_(slot.date,state.periodStart,state.periodEnd)
  ) {
    const periodTarget = num_(model.settings.periodHoursTarget,320);
    const currentPeriodHours = num_(state.periodHours[username],0);
    const shiftCredit = creditedHours_(slot.shift);
    const remainingBefore = periodTarget - currentPeriodHours;
    const remainingAfter = periodTarget - (currentPeriodHours + shiftCredit);

    // Employees further below 320 receive much higher assignment priority.
    score += Math.max(0,remainingBefore) * 8;

    // Highest priority when this shift lands exactly on 320.
    if (Math.abs(remainingAfter) < 0.0001) {
      score += 10000;
    } else {
      score += Math.max(0,500 - Math.abs(remainingAfter) * 4);
    }
  }

  score -= num_(state.totalAssignments[username],0)*1.5;

  if (type==='evening') score -= num_(state.monthlyEvenings[username+'|'+monthKey_(slot.date)],0)*10;
  if (type==='night') score -= num_(state.nightAssignments[username],0)*5;
  if (slot.weekend || isWeekendDate_(slot.date)) score -= num_(state.weekendAssignments[username],0)*4;

  const priorKey=username+'|'+formatDateKey_(addDays_(slot.date,-1));
  const priorType=state.lastTypeByEmployeeDate[priorKey];
  if (priorType) {
    if (priorType===type) score += 12;
    if ((priorType==='night' && type!=='night') || (priorType==='evening' && type==='day')) score -= 25;
  }

  if (isWeekendDate_(slot.date)) {
    const sat = weekendSaturday_(slot.date);
    const roleKey=formatDateKey_(sat)+'|'+slot.shiftCode;
    if (state.weekendRoleOwner[roleKey]===username) score += 70;
  }

  if (yes_(u.Resident) && isResidentMandatoryE2_(u,slot,model)) score += 5000;
  if (yes_(u.Resident) && isWeekendDate_(slot.date) && clean_(u['Weekend Group'])===slot.weekendGroup) score += 2500;
  if (elig.warnings && elig.warnings.some(w=>w.indexOf('WEEKEND ROTATION EXCEPTION')===0)) score -= 150;
  score += deterministicTie_(username+'|'+slot.slotKey);
  return score;
}

function nextGeneratedAssignmentId_(slot) {
  // v30: Assignment IDs must be unique across chunked weekly generation calls.
  // Older versions reset a sequence counter for every Apps Script request, so
  // week 1 and week 3 could both contain an ID such as GA-<generation>-17.
  // The calendar then opened the FIRST row with that duplicated ID, which made
  // a click on (for example) 09/21 ONC display an unrelated earlier assignment.
  // Date + Shift + Slot uniquely identify a required staffing slot, so include
  // them directly in the generated ID. This stays unique even when the 8-week
  // schedule is generated as eight independent server calls.
  const datePart=formatDateKey_(slot.date).replace(/-/g,'');
  const shiftPart=clean_(slot.shiftCode).replace(/[^A-Za-z0-9_-]+/g,'_') || 'SHIFT';
  const slotPart='S'+String(num_(slot.slot,1));
  return ['GA',clean_(slot.generationId),datePart,shiftPart,slotPart].join('-');
}
function makeAssigned_(slot,u,warnings,actor) {
  return {
    generationId:slot.generationId,assignmentId:nextGeneratedAssignmentId_(slot),date:slot.date,dateKey:slot.dateKey,day:slot.day,
    shiftCode:slot.shiftCode,slot:slot.slot,shift:slot.shift,pharmacist:clean_(u['Pharmacist Name']),username:clean_(u.Username),
    hours:num_(slot.shift.Hours,0),creditedHours:creditedHours_(slot.shift),requiredSkill:slot.requiredSkill,
    coverageForPharmacist:'',coverageForUsername:'',coverageReason:'',shiftType:clean_(slot.shift.Type),
    weekend:isWeekendDate_(slot.date),weekendGroup:slot.weekendGroup,holiday:'',locked:false,manual:false,status:'ASSIGNED',
    warning:(warnings||[]).join(' | '),updatedAt:new Date(),updatedBy:actor,finalizedAt:null
  };
}

function makeUnfilledAssignment_(slot,model,state,actor) {
  return {
    generationId:slot.generationId,assignmentId:nextGeneratedAssignmentId_(slot),date:slot.date,dateKey:slot.dateKey,day:slot.day,
    shiftCode:slot.shiftCode,slot:slot.slot,shift:slot.shift,pharmacist:'UNFILLED',username:'',hours:num_(slot.shift.Hours,0),creditedHours:creditedHours_(slot.shift),
    requiredSkill:slot.requiredSkill,coverageForPharmacist:'',coverageForUsername:'',coverageReason:'',shiftType:clean_(slot.shift.Type),weekend:isWeekendDate_(slot.date),weekendGroup:slot.weekendGroup,holiday:'',
    locked:false,manual:false,status:'UNFILLED',warning:'UNFILLED — pending final eligibility analysis.',updatedAt:new Date(),updatedBy:actor,finalizedAt:null
  };
}

function addAssignmentToState_(state,a,model) {
  if (!a || !a.username || a.status==='UNFILLED') return;
  const u=a.username, dk=a.dateKey;
  state.assignments.push(a);
  state.byEmployeeDate[u+'|'+dk]=a;
  if (!state.assignedDateKeysByEmployee[u]) state.assignedDateKeysByEmployee[u]=new Set();
  state.assignedDateKeysByEmployee[u].add(dk);
  if (!state.intervalsByEmployee[u]) state.intervalsByEmployee[u]=[];
  const interval=assignmentInterval_(a.date,a.shift);
  if (interval) state.intervalsByEmployee[u].push({start:interval.start,end:interval.end,assignment:a});
  const wk=u+'|'+weekStartKey_(a.date,model.settings.weekStart);
  state.weeklyHours[wk]=num_(state.weeklyHours[wk],0)+num_(a.creditedHours,creditedHours_(a.shift));
  state.weeklyDays[wk]=num_(state.weeklyDays[wk],0)+1;
  state.totalAssignments[u]=num_(state.totalAssignments[u],0)+1;
  state.totalHours[u]=num_(state.totalHours[u],0)+num_(a.creditedHours,0);
  if (
    state.periodStart &&
    state.periodEnd &&
    inDateRange_(a.date, state.periodStart, state.periodEnd)
  ) {
    state.periodHours[u] =
      num_(state.periodHours[u],0) +
      num_(a.creditedHours,creditedHours_(a.shift));
  }
  const type=clean_(a.shiftType || a.shift.Type).toLowerCase();
  if (type==='evening') {
    const mk=u+'|'+monthKey_(a.date); state.monthlyEvenings[mk]=num_(state.monthlyEvenings[mk],0)+1;
  }
  if (type==='night') state.nightAssignments[u]=num_(state.nightAssignments[u],0)+1;
  if (a.weekend || isWeekendDate_(a.date)) state.weekendAssignments[u]=num_(state.weekendAssignments[u],0)+1;
  const emp=model.usersByUsername[u];
  if (emp && yes_(emp.Resident) && clean_(a.shiftCode).toUpperCase()===model.settings.residentE2ShiftCode) {
    const e2wk=u+'|'+weekStartKey_(a.date,model.settings.weekStart);
    state.residentE2ByWeek[e2wk]=num_(state.residentE2ByWeek[e2wk],0)+1;
  }
  state.lastTypeByEmployeeDate[u+'|'+dk]=type;
  if (isWeekendDate_(a.date)) state.weekendRoleOwner[formatDateKey_(weekendSaturday_(a.date))+'|'+a.shiftCode]=u;
}

function removeAssignmentFromState_(state,a,model) {
  if (!a || !a.username || a.status==='UNFILLED') return;
  const u=a.username, dk=a.dateKey;
  delete state.byEmployeeDate[u+'|'+dk];
  if(state.assignedDateKeysByEmployee[u]) state.assignedDateKeysByEmployee[u].delete(dk);
  state.assignments=state.assignments.filter(x=>x.assignmentId!==a.assignmentId);
  if (state.intervalsByEmployee[u]) state.intervalsByEmployee[u]=state.intervalsByEmployee[u].filter(x=>x.assignment.assignmentId!==a.assignmentId);
  const wk=u+'|'+weekStartKey_(a.date,model.settings.weekStart);
  state.weeklyHours[wk]=Math.max(0,num_(state.weeklyHours[wk],0)-num_(a.creditedHours,creditedHours_(a.shift)));
  state.weeklyDays[wk]=Math.max(0,num_(state.weeklyDays[wk],0)-1);
  state.totalAssignments[u]=Math.max(0,num_(state.totalAssignments[u],0)-1);
  state.totalHours[u]=Math.max(0,num_(state.totalHours[u],0)-num_(a.creditedHours,0));
  if (
    state.periodStart &&
    state.periodEnd &&
    inDateRange_(a.date, state.periodStart, state.periodEnd)
  ) {
    state.periodHours[u] = Math.max(
      0,
      num_(state.periodHours[u],0) -
      num_(a.creditedHours,creditedHours_(a.shift))
    );
  }
  const type=clean_(a.shiftType || a.shift.Type).toLowerCase();
  if (type==='evening') { const mk=u+'|'+monthKey_(a.date); state.monthlyEvenings[mk]=Math.max(0,num_(state.monthlyEvenings[mk],0)-1); }
  if (type==='night') state.nightAssignments[u]=Math.max(0,num_(state.nightAssignments[u],0)-1);
  if (a.weekend || isWeekendDate_(a.date)) state.weekendAssignments[u]=Math.max(0,num_(state.weekendAssignments[u],0)-1);
  const emp=model.usersByUsername[u];
  if (emp && yes_(emp.Resident) && clean_(a.shiftCode).toUpperCase()===model.settings.residentE2ShiftCode) {
    const e2wk=u+'|'+weekStartKey_(a.date,model.settings.weekStart);
    state.residentE2ByWeek[e2wk]=Math.max(0,num_(state.residentE2ByWeek[e2wk],0)-1);
  }
  delete state.lastTypeByEmployeeDate[u+'|'+dk];
  if (isWeekendDate_(a.date)) {
    const k=formatDateKey_(weekendSaturday_(a.date))+'|'+a.shiftCode;
    if (state.weekendRoleOwner[k]===u) delete state.weekendRoleOwner[k];
  }
}

function hasTimeConflict_(username,slot,state) {
  const interval=assignmentInterval_(slot.date,slot.shift);
  if (!interval) return false;
  return (state.intervalsByEmployee[username]||[]).some(x => interval.start < x.end && interval.end > x.start);
}

function assignSpecificUser_(slot,u,model,state,actor,note) {
  const e=eligibility_(u,slot,model,state,false);
  if(!e.ok)return null;
  const warnings=(e.warnings||[]).slice();
  if(note)warnings.push(note);
  const a=makeAssigned_(slot,u,warnings,actor);
  addAssignmentToState_(state,a,model);
  return a;
}

function preassignSevenOnSevenOff_(slots,results,model,state,actor,reserved) {
  const seven=model.users
    .filter(u=>yes_(u.Active)&&isSevenOn_(u))
    .sort((a,b)=>clean_(a['Pharmacist Name']).localeCompare(clean_(b['Pharmacist Name'])));

  seven.forEach(u=>{
    const username=clean_(u.Username);
    const code=sevenOnAssignedShiftCode_(u,model);
    if(!code)return;

    const dedicated=slots
      .filter(s=>s.shiftCode===code&&sevenOnIsOnDay_(u,s.date))
      .sort((a,b)=>a.dateKey.localeCompare(b.dateKey)||a.slot-b.slot);
    if(!dedicated.length)return;

    // Split the visible schedule into consecutive ON segments. A segment is
    // normally seven days, except when the requested schedule starts or ends
    // in the middle of a 7-day ON block.
    const segments=[];
    let current=[];
    dedicated.forEach(slot=>{
      if(!current.length){current=[slot];return;}
      const prev=current[current.length-1];
      if(daysBetween_(prev.date,slot.date)===1){current.push(slot);}
      else {segments.push(current);current=[slot];}
    });
    if(current.length)segments.push(current);

    segments.forEach(segment=>{
      // First evaluate the entire block with the 320-hour cap temporarily
      // removed. This prevents the cap from cutting a 7-on block in half.
      const noPeriodState=Object.assign({},state,{periodStart:null,periodEnd:null});
      const eligible=segment.filter(slot=>{
        if(reserved.has(slot.slotKey))return false;
        if(state.byEmployeeDate[username+'|'+slot.dateKey])return false;
        return eligibility_(u,slot,model,noPeriodState,false).ok;
      });

      if(!eligible.length)return;

      // True 7-on/7-off employees are governed by the complete block pattern,
      // not the regular five-day/320-hour rule. This prevents a period-hours
      // target from cutting a 7-day ON block into a partial block.
      eligible.forEach(slot=>{
        if(reserved.has(slot.slotKey))return;
        const a=assignSpecificUser_(slot,u,model,state,actor,'7-ON/7-OFF BLOCK');
        if(a){
          results.push(a);
          reserved.add(slot.slotKey);
        }
      });
    });
  });
}

function weekStartsInRange_(start,end,weekStartName) {
  const first=startOfDay_(asDate_(weekStartKey_(start,weekStartName)));
  const out=[];
  for(let d=first;d<=end;d=addDays_(d,7))out.push(d);
  return out;
}

function preassignResidentE2_(slots,results,model,state,actor,reserved,start,end) {
  const residents=model.users.filter(u=>yes_(u.Active)&&yes_(u.Resident));
  const code=model.settings.residentE2ShiftCode;
  const perWeek=model.settings.residentE2PerWeek;
  if(!code||perWeek<1)return;

  const weeks=weekStartsInRange_(start,end,model.settings.weekStart);
  weeks.forEach(ws=>{
    const we=addDays_(ws,6);
    const rangeStart=ws<start?start:ws;
    const rangeEnd=we>end?end:we;
    const e2Slots=slots.filter(s=>!reserved.has(s.slotKey)&&s.shiftCode===code&&inDateRange_(s.date,rangeStart,rangeEnd));
    if(!e2Slots.length)return;

    const dayLoad={};
    results.filter(a=>a.status!=='UNFILLED'&&a.shiftCode===code&&inDateRange_(a.date,rangeStart,rangeEnd)).forEach(a=>dayLoad[a.dateKey]=num_(dayLoad[a.dateKey],0)+1);

    residents.slice().sort((a,b)=>clean_(a['Pharmacist Name']).localeCompare(clean_(b['Pharmacist Name']))).forEach(u=>{
      const username=clean_(u.Username);
      const wk=username+'|'+weekStartKey_(rangeStart,model.settings.weekStart);
      let have=num_(state.residentE2ByWeek[wk],0);
      while(have<perWeek){
        const candidates=e2Slots
          .filter(s=>!reserved.has(s.slotKey)&&!state.byEmployeeDate[username+'|'+s.dateKey])
          .sort((a,b)=>num_(dayLoad[a.dateKey],0)-num_(dayLoad[b.dateKey],0)||a.dateKey.localeCompare(b.dateKey));
        let assigned=null;
        for(const slot of candidates){
          assigned=assignSpecificUser_(slot,u,model,state,actor,'RESIDENT REQUIRED E2');
          if(assigned){
            results.push(assigned);
            reserved.add(slot.slotKey);
            dayLoad[slot.dateKey]=num_(dayLoad[slot.dateKey],0)+1;
            have++;
            break;
          }
        }
        if(!assigned)break;
      }
    });
  });
}


function preassignResidentWeekendsPriority_(slots,results,model,state,actor,reserved,start,end) {
  const residents=model.users
    .filter(u=>yes_(u.Active)&&yes_(u.Resident)&&regularFiveDayRuleApplies_(u)&&clean_(u['Weekend Group']))
    .filter(u=>model.settings.weekendRotation.includes(clean_(u['Weekend Group'])))
    .sort((a,b)=>clean_(a['Pharmacist Name']).localeCompare(clean_(b['Pharmacist Name'])));
  if(!residents.length)return;

  const saturdayKeys=[...new Set(slots.filter(s=>s.weekend).map(s=>formatDateKey_(weekendSaturday_(s.date))))].sort();
  saturdayKeys.forEach(satKey=>{
    const sat=startOfDay_(asDate_(satKey));
    const expected=weekendGroupForDate_(sat,model.settings);
    const requiredResidents=residents.filter(u=>
      clean_(u['Weekend Group'])===expected && employeeWeekendIsOn_(u,sat,model.settings)
    );
    if(!requiredResidents.length)return;

    [sat,addDays_(sat,1)].forEach(day=>{
      if(day<start||day>end)return;
      const dk=formatDateKey_(day);
      const daySlots=slots.filter(s=>s.dateKey===dk&&s.weekend&&!reserved.has(s.slotKey));
      requiredResidents.forEach(u=>{
        const username=clean_(u.Username);
        if(state.byEmployeeDate[username+'|'+dk])return;
        const candidates=daySlots
          .filter(s=>!reserved.has(s.slotKey))
          .map(s=>{
            const e=eligibility_(u,s,model,state,false);
            return {slot:s,e:e,scarcity:countStaticCandidates_(s,model),preferred:preferredShiftMatches_(u,s)?0:1};
          })
          .filter(x=>x.e.ok)
          .sort((a,b)=>a.preferred-b.preferred||a.scarcity-b.scarcity||shiftCategoryRank_(a.slot.shift)-shiftCategoryRank_(b.slot.shift)||num_(a.slot.shift.Priority,50)-num_(b.slot.shift.Priority,50)||a.slot.shiftCode.localeCompare(b.slot.shiftCode));
        if(!candidates.length)return;
        const pick=candidates[0].slot;
        const a=assignSpecificUser_(pick,u,model,state,actor,'REQUIRED RESIDENT WEEKEND GROUP '+expected);
        if(a){
          results.push(a);
          reserved.add(pick.slotKey);
        }
      });
    });
  });
}

function preassignPairedEdCoverage_(slots,results,model,state,actor,reserved,start,end) {
  const dayCode=clean_(model.settings.edDayShiftCode||'EDD').toUpperCase();
  const eveCode=clean_(model.settings.edEveningShiftCode||'EDE').toUpperCase();
  if(!dayCode||!eveCode)return;

  const dateKeys=[...new Set(slots
    .filter(s=>!isWeekendDate_(s.date)&&(s.shiftCode===dayCode||s.shiftCode===eveCode))
    .map(s=>s.dateKey))].sort();

  dateKeys.forEach(dk=>{
    const daySlots=slots.filter(s=>s.dateKey===dk&&s.shiftCode===dayCode&&!reserved.has(s.slotKey)).sort((a,b)=>a.slot-b.slot);
    const eveSlots=slots.filter(s=>s.dateKey===dk&&s.shiftCode===eveCode&&!reserved.has(s.slotKey)).sort((a,b)=>a.slot-b.slot);
    const pairCount=Math.min(daySlots.length,eveSlots.length);

    for(let i=0;i<pairCount;i++){
      const ds=daySlots[i], es=eveSlots[i];
      const dc=candidateListForSlot_(ds,model,state);
      const ec=candidateListForSlot_(es,model,state);
      let best=null;
      dc.forEach(d=>ec.forEach(e=>{
        if(clean_(d.user.Username)===clean_(e.user.Username))return;
        let score=d.score+e.score;
        // Give a strong but still rule-respecting preference to pharmacists whose
        // profile explicitly identifies EDD/EDE as their preferred work area.
        if(preferredShiftMatches_(d.user,ds))score+=500;
        if(preferredShiftMatches_(e.user,es))score+=500;
        if(!best||score>best.score)best={d:d,e:e,score:score};
      }));

      if(best){
        const ea=assignSpecificUser_(es,best.e.user,model,state,actor,'PAIRED ED COVERAGE '+dayCode+'/'+eveCode+' - EVENING');
        const da=assignSpecificUser_(ds,best.d.user,model,state,actor,'PAIRED ED COVERAGE '+dayCode+'/'+eveCode+' - DAY');
        if(ea){results.push(ea);reserved.add(es.slotKey);}
        if(da){results.push(da);reserved.add(ds.slotKey);}
      } else {
        // If the full pair is infeasible, still preserve any legal coverage that
        // can be obtained. Validation will explicitly report the missing half.
        const ec2=candidateListForSlot_(es,model,state);
        if(ec2.length){
          const ea=assignSpecificUser_(es,ec2[0].user,model,state,actor,'ED REQUIRED COVERAGE '+eveCode);
          if(ea){results.push(ea);reserved.add(es.slotKey);}
        }
        const exclude=state.byEmployeeDate;
        const dc2=candidateListForSlot_(ds,model,state);
        if(dc2.length){
          const da=assignSpecificUser_(ds,dc2[0].user,model,state,actor,'ED REQUIRED COVERAGE '+dayCode);
          if(da){results.push(da);reserved.add(ds.slotKey);}
        }
      }
    }
  });
}

function preassignWeekendTeams_(slots,results,model,state,actor,reserved,start,end) {
  const teamUsers=model.users
    .filter(u=>yes_(u.Active)&&regularFiveDayRuleApplies_(u)&&clean_(u['Weekend Group']))
    .filter(u=>model.settings.weekendRotation.includes(clean_(u['Weekend Group'])))
    .sort((a,b)=>clean_(a['Pharmacist Name']).localeCompare(clean_(b['Pharmacist Name'])));
  if(!teamUsers.length)return;

  // Build one record for every physical weekend represented in the generated
  // slots. Saturday and Sunday belong to adjacent scheduling weeks, but they
  // remain one required A/B/C weekend pair.
  const saturdayKeys=[...new Set(slots.filter(s=>s.weekend).map(s=>formatDateKey_(weekendSaturday_(s.date))))].sort();
  saturdayKeys.forEach(satKey=>{
    const sat=startOfDay_(asDate_(satKey));
    const expected=weekendGroupForDate_(sat,model.settings);
    const groupUsers=teamUsers
      .filter(u=>clean_(u['Weekend Group'])===expected)
      .filter(u=>employeeWeekendIsOn_(u,sat,model.settings))
      .sort((a,b)=>(yes_(b.Resident)?1:0)-(yes_(a.Resident)?1:0)||clean_(a['Pharmacist Name']).localeCompare(clean_(b['Pharmacist Name'])));
    if(!groupUsers.length)return;

    [sat,addDays_(sat,1)].forEach(day=>{
      if(day<start||day>end)return;
      const dk=formatDateKey_(day);
      const daySlots=slots.filter(s=>s.dateKey===dk&&s.weekend&&!reserved.has(s.slotKey));
      if(!daySlots.length)return;

      groupUsers.forEach(u=>{
        const username=clean_(u.Username);
        // A preserved/manual assignment on the date already satisfies the
        // workday count, but validation will still verify that the employee
        // actually worked the required weekend group.
        if(state.byEmployeeDate[username+'|'+dk])return;

        const priorPair=results.find(a=>a.status!=='UNFILLED'&&a.username===username&&formatDateKey_(weekendSaturday_(a.date))===satKey);
        const preferredCode=priorPair?clean_(priorPair.shiftCode):'';
        const candidates=daySlots
          .filter(s=>!reserved.has(s.slotKey))
          .map(s=>({slot:s,scarcity:countStaticCandidates_(s,model),same:preferredCode&&s.shiftCode===preferredCode?0:1}))
          .sort((a,b)=>a.same-b.same||a.scarcity-b.scarcity||shiftCategoryRank_(a.slot.shift)-shiftCategoryRank_(b.slot.shift)||num_(a.slot.shift.Priority,50)-num_(b.slot.shift.Priority,50)||a.slot.shiftCode.localeCompare(b.slot.shiftCode));

        for(const c of candidates){
          const note='REQUIRED WEEKEND GROUP '+expected;
          const a=assignSpecificUser_(c.slot,u,model,state,actor,note);
          if(a){
            results.push(a);
            reserved.add(c.slot.slotKey);
            break;
          }
        }
      });
    });
  });
}

// Compatibility wrapper for any older script/button that still calls the
// resident-only function name. v12 uses the all-team function above.
function preassignResidentWeekends_(slots,results,model,state,actor,reserved) {
  const dates=slots.map(s=>s.date).filter(Boolean);
  if(!dates.length)return;
  const start=new Date(Math.min.apply(null,dates.map(d=>d.getTime())));
  const end=new Date(Math.max.apply(null,dates.map(d=>d.getTime())));
  return preassignWeekendTeams_(slots,results,model,state,actor,reserved,start,end);
}

function runRepairPasses_(results,model,state,actor) {
  const deadline=num_(model.runtimeDeadline,0);
  const timedOut=()=>deadline>0 && Date.now()>deadline;
  // Pass 1: retry any unfilled slot after the initial distribution.
  const firstUnfilled=results.filter(a=>a.status==='UNFILLED');
  for(const u of firstUnfilled){
    if(timedOut())return;
    const slot=assignmentAsSlot_(u);
    const direct=assignBestCandidate_(slot,model,state,actor);
    if (direct) replaceAssignmentObject_(u,direct);
  }

  // Pass 2: same-day two-step repair chain. Build the date index once instead
  // of filtering the whole result set for every unfilled slot.
  const assignedByDate={};
  results.forEach(a=>{
    if(a.status==='UNFILLED'||a.locked||a.manual||isPatternProtectedAssignment_(a))return;
    if(!assignedByDate[a.dateKey])assignedByDate[a.dateKey]=[];
    assignedByDate[a.dateKey].push(a);
  });
  const secondUnfilled=results.filter(a=>a.status==='UNFILLED');
  for(const unfilled of secondUnfilled){
    if(timedOut())return;
    const targetSlot=assignmentAsSlot_(unfilled);
    const sameDay=(assignedByDate[unfilled.dateKey]||[]).slice();
    for (let i=0;i<sameDay.length;i++) {
      if(timedOut())return;
      const donor=sameDay[i];
      if(donor.status==='UNFILLED'||donor.locked||donor.manual||isPatternProtectedAssignment_(donor))continue;
      const donorUser=model.usersByUsername[donor.username];
      if (!donorUser) continue;
      removeAssignmentFromState_(state,donor,model);
      const donorToTarget=eligibility_(donorUser,targetSlot,model,state,false);
      if (!donorToTarget.ok) { addAssignmentToState_(state,donor,model); continue; }
      const moved=makeAssigned_(targetSlot,donorUser,donorToTarget.warnings.concat(['REPAIR: reassigned from '+donor.shiftCode+' to cover scarce shift.']),actor);
      addAssignmentToState_(state,moved,model);
      const vacatedSlot=assignmentAsSlot_(donor);
      const replacement=assignBestCandidate_(vacatedSlot,model,state,actor);
      if (replacement) {
        replaceAssignmentObject_(unfilled,moved);
        replaceAssignmentObject_(donor,replacement);
        break;
      }
      removeAssignmentFromState_(state,moved,model);
      addAssignmentToState_(state,donor,model);
    }
  }
}

function rebalanceRegularWeeklyWorkload_(results,model,state,actor,start,end) {
  const deadline=num_(model.runtimeDeadline,0);
  const timedOut=()=>deadline>0&&Date.now()>deadline;
  const regularUsers=model.users.filter(u=>yes_(u.Active)&&regularFiveDayRuleApplies_(u));
  if(!regularUsers.length)return 0;

  let changes=0;
  const weeks=weekStartsInRange_(start,end,model.settings.weekStart);

  for(const ws of weeks){
    if(timedOut())break;
    const we=addDays_(ws,6);

    // For a monthly/custom range, only enforce dates actually inside the
    // generation window. Complete weeks get the exact weekly requirement.
    const complete=ws>=start&&we<=end;
    if(!complete)continue;

    for(const u of regularUsers){
      if(timedOut())break;
      const username=clean_(u.Username);
      const wk=username+'|'+weekStartKey_(ws,model.settings.weekStart);
      const target=regularRequiredWorkdaysForWeek_(u,ws,model);
      let current=num_(state.weeklyDays[wk],0);

      while(current<target&&!timedOut()){
        let filled=false;

        // First use an actually UNFILLED required slot.
        const open=results
          .filter(a=>
            a.status==='UNFILLED' &&
            inDateRange_(a.date,ws,we)
          )
          .sort((a,b)=>a.dateKey.localeCompare(b.dateKey)||a.shiftCode.localeCompare(b.shiftCode));

        for(const unfilled of open){
          const slot=assignmentAsSlot_(unfilled);
          const e=eligibility_(u,slot,model,state,false);
          if(!e.ok)continue;

          const assigned=makeAssigned_(
            slot,
            u,
            (e.warnings||[]).concat(['REGULAR WORKLOAD: filled to satisfy required weekly workdays.']),
            actor
          );

          addAssignmentToState_(state,assigned,model);
          replaceAssignmentObject_(unfilled,assigned);
          current++;
          changes++;
          filled=true;
          break;
        }

        if(filled)continue;

        // Next reclaim a non-protected PRN assignment. PRN is supplemental and
        // may not occupy a slot while an eligible regular pharmacist is short.
        const prnAssignments=results
          .filter(a=>{
            if(a.status==='UNFILLED'||a.locked||a.manual||isPatternProtectedAssignment_(a))return false;
            if(!inDateRange_(a.date,ws,we))return false;
            const donor=model.usersByUsername[a.username];
            return donor&&isPrnEmployee_(donor);
          })
          .sort((a,b)=>a.dateKey.localeCompare(b.dateKey)||a.shiftCode.localeCompare(b.shiftCode));

        for(const donor of prnAssignments){
          const slot=assignmentAsSlot_(donor);
          const e=eligibility_(u,slot,model,state,false);
          if(!e.ok)continue;

          const donorUser=model.usersByUsername[donor.username];
          removeAssignmentFromState_(state,donor,model);

          const assigned=makeAssigned_(
            slot,
            u,
            (e.warnings||[]).concat(['REGULAR WORKLOAD: regular pharmacist replaced PRN backup to satisfy weekly work requirement.']),
            actor
          );

          addAssignmentToState_(state,assigned,model);
          replaceAssignmentObject_(donor,assigned);
          current++;
          changes++;
          filled=true;
          break;
        }

        if(!filled)break;
      }
    }
  }

  return changes;
}

function refreshUnfilledReasons_(results,model,state) {
  results.filter(a=>a.status==='UNFILLED').forEach(a=>a.warning=explainUnfilled_(assignmentAsSlot_(a),model,state));
}

function explainUnfilled_(slot,model,state) {
  const owners=primaryHomeOwnersForSlot_(slot,model);
  if(owners.length===1 && isGeneratedOffDayForHomeOwner_(owners[0],slot,state,model)) {
    const backups=offDaySkillQualifiedUsersForSlot_(slot,model)
      .filter(u=>clean_(u.Username)!==clean_(owners[0].Username)&&hasOffDayCoverageSkillForSlot_(u,slot));
    if(backups.length){
      const blocked={};
      backups.forEach(u=>{
        const e=offDayCoverageEligibility_(u,slot,model,state);
        const key=e.ok?'OTHER':((e.reasons||[])[0]||'OTHER');
        blocked[key]=num_(blocked[key],0)+1;
      });
      const labels={PTO:'on approved PTO',REGULAR_OFF:'on approved Regular Off',UNAVAILABLE:'unavailable',ALREADY_SCHEDULED:'already scheduled',TIME_CONFLICT:'time conflict',WEEKLY_HOURS:'over weekly hours',WEEKLY_DAYS:'already at the five-workday weekly limit',CONSECUTIVE_DAYS:'would create more than five consecutive workdays',EVENING_TO_MORNING:'evening-to-morning transition requires the next day OFF or another evening shift',TWO_MONTH_HOURS:'would exceed the exact 320-hour schedule-period target',PRECEPTOR_WEEKDAY_EVENING:'preceptor weekday-evening restriction',EVENING_LIMIT:'at evening limit',WRONG_WEEKEND_GROUP:'wrong weekend group',WEEKEND_ANCHOR_OFF:'outside weekend rotation',WEEKEND_NOT_ELIGIBLE:'not weekend eligible',EVENING_NOT_ELIGIBLE:'not evening eligible',NIGHT_NOT_ELIGIBLE:'not night eligible',RESIDENT_RESTRICTION:'resident restriction',OTHER:'otherwise not assignable'};
      const pieces=Object.keys(blocked).map(k=>blocked[k]+' '+(labels[k]||k.toLowerCase().replace(/_/g,' ')));
      return 'UNFILLED — '+clean_(owners[0]['Pharmacist Name'])+' has an algorithm-generated OFF day; '+backups.length+' pharmacist(s) are configured in Users -> Off-Day Coverage Skills for '+clean_(slot.shiftCode)+': '+pieces.join('; ')+'.';
    }
    return 'UNFILLED — '+clean_(owners[0]['Pharmacist Name'])+' has an algorithm-generated OFF day and no pharmacist is configured in Users -> Off-Day Coverage Skills for '+clean_(slot.shiftCode)+'.';
  }
  const skillQualified=skillQualifiedUsersForSlot_(slot,model);
  if (!skillQualified.length) return 'UNFILLED — 0 active employees have required skill '+(slot.requiredSkill||'(none)')+'.';
  const counts={};
  skillQualified.forEach(u=>{
    const e=eligibility_(u,slot,model,state,false);
    if (e.ok) counts['OTHER']=num_(counts.OTHER,0)+1;
    else {
      const primary=e.reasons[0]||'OTHER'; counts[primary]=num_(counts[primary],0)+1;
    }
  });
  const labels={PTO:'on approved PTO',REGULAR_OFF:'on approved Regular Off',UNAVAILABLE:'unavailable',ALREADY_SCHEDULED:'already scheduled',TIME_CONFLICT:'time conflict',WEEKLY_HOURS:'over weekly hours',WEEKLY_DAYS:'already at the five-workday weekly limit',CONSECUTIVE_DAYS:'would create more than five consecutive workdays',EVENING_TO_MORNING:'evening-to-morning transition requires the next day OFF or another evening shift',TWO_MONTH_HOURS:'would exceed the exact 320-hour schedule-period target',PRECEPTOR_WEEKDAY_EVENING:'preceptor weekday-evening restriction',PRECEPTOR_EVENING:'preceptor evening restriction',EVENING_LIMIT:'at evening limit',WRONG_WEEKEND_GROUP:'wrong weekend group',RESIDENT_WRONG_WEEKEND_GROUP:'resident is outside the assigned weekend group',WEEKEND_ANCHOR_OFF:'outside the employee 21-day weekend rotation anchor',SEVEN_OFF:'in 7-on/7-off OFF period',SEVEN_WRONG_SHIFT:'7-on/7-off employee is restricted to the dedicated shift',RESIDENT_E2_WEEKLY_LIMIT:'resident already has the required E2 shift for this week',CUSTOM_HOURS:'outside hard custom hours',WEEKLY_AVAILABILITY_DAY:'not available on this weekday',WEEKLY_AVAILABILITY_TIME:'outside recurring weekly available hours',PRN_NOT_AVAILABLE:'not listed as available in My Availability',PRN_AVAILABILITY_NOT_LOADED:'My Availability has not loaded; PRN scheduling is blocked',WEEKEND_NOT_ELIGIBLE:'not weekend eligible',EVENING_NOT_ELIGIBLE:'not evening eligible',NIGHT_NOT_ELIGIBLE:'not night eligible',RESIDENT_RESTRICTION:'resident restriction',OTHER:'otherwise not assignable'};
  const pieces=Object.keys(counts).map(k=>counts[k]+' '+(labels[k]||k.toLowerCase().replace(/_/g,' ')));
  return 'UNFILLED — '+skillQualified.length+' active employees have the skill: '+pieces.join('; ')+'.';
}

function assignmentAsSlot_(a) {
  return {generationId:a.generationId,date:a.date,dateKey:a.dateKey,day:a.day,shiftCode:a.shiftCode,slot:a.slot,shift:a.shift,requiredSkill:a.requiredSkill,weekend:a.weekend,weekendGroup:a.weekendGroup,slotKey:slotKey_(a.dateKey,a.shiftCode,a.slot)};
}

function replaceAssignmentObject_(target,source) {
  Object.keys(target).forEach(k=>delete target[k]);
  Object.keys(source).forEach(k=>target[k]=source[k]);
}

function validateGeneratedAssignments_(assignments,model,start,end,validationOptions) {
  validationOptions=validationOptions||{};
  const errors=[],warnings=[];
  const generatedAssigned=assignments.filter(a=>a.status!=='UNFILLED'&&a.username);
  const contextStart=generationContextStart_(start,model.settings.weekStart),contextEnd=generationContextEnd_(end,model.settings.weekStart);
  const external=readTable_(APP.SHEETS.SCHEDULE).filter(r=>{const d=asDate_(r.Date);return d&&!inDateRange_(d,start,end)&&inDateRange_(d,contextStart,contextEnd)&&clean_(r.Status)!=='UNFILLED'&&clean_(r.Username);}).map(r=>scheduleRowToAssignment_(r,model)).filter(Boolean);
  const assigned=generatedAssigned.concat(external);
  generatedAssigned.forEach(a=>{
    const u=model.usersByUsername[a.username];
    if (!u) errors.push(a.dateKey+' '+a.shiftCode+': assigned user not found.');
    else {
      // Static checks not requiring state totals.
      const validationSlot=assignmentAsSlot_(a);
      const offDayCoverageAssignment=clean_(a.coverageReason).toUpperCase()==='GENERATED OFF DAY' && hasOffDayCoverageSkillForSlot_(u,validationSlot);
      if (!hasRequiredSkillForSlot_(u,validationSlot,model) && !offDayCoverageAssignment) errors.push(a.dateKey+' '+a.shiftCode+': '+a.pharmacist+' lacks '+a.requiredSkill+' and is not configured for OFF-day coverage.');
      if (isBlockedByPto_(u,a.date,model)) errors.push(a.dateKey+' '+a.shiftCode+': '+a.pharmacist+' is on approved PTO.');
      if (isBlockedByRegularOff_(u,a.date,model)) errors.push(a.dateKey+' '+a.shiftCode+': '+a.pharmacist+' has an approved Regular Off request.');
      const wa=weeklyAvailabilityStatus_(u,a.date,a.shift,model);
      if(wa.blocked&&!a.manual) errors.push(a.dateKey+' '+a.shiftCode+': '+a.pharmacist+' violates recurring weekly availability.');

      const prnAvail=prnAvailabilityStatus_(u,a.date,a.shift,model);
      if(prnAvail.blocked&&!a.manual) errors.push(a.dateKey+' '+a.shiftCode+': '+a.pharmacist+' is PRN and is not available for this date/shift in My Availability.');

      if (preceptorEveningBlocked_(u,validationSlot,model)) errors.push(a.dateKey+' '+a.shiftCode+': preceptor assigned a weekday evening shift; preceptors may work evening shifts on weekends only.');
      if (isSevenOn_(u)&&!sevenOnIsOnDay_(u,a.date)&&!a.manual) errors.push(a.dateKey+' '+a.shiftCode+': 7-on/7-off employee scheduled during OFF period.');
      if (isSevenOn_(u)&&sevenOnIsOnDay_(u,a.date)&&!sevenOnSlotMatches_(u,validationSlot,model)&&!a.manual) errors.push(a.dateKey+' '+a.shiftCode+': 7-on/7-off employee '+a.pharmacist+' must stay on dedicated shift '+sevenOnAssignedShiftCode_(u,model)+'.');
      if (regularFiveDayRuleApplies_(u)&&clean_(u['Weekend Group'])&&isWeekendDate_(a.date)&&clean_(u['Weekend Group'])!==a.weekendGroup) errors.push(a.dateKey+' '+a.shiftCode+': '+a.pharmacist+' is scheduled outside required Weekend Group '+clean_(u['Weekend Group'])+'.');
    }
  });
  // Validate v15 OFF-day coverage relationships against the completed schedule.
  const completedState=createState_(model,[],start,end);
  assigned.forEach(a=>addAssignmentToState_(completedState,a,model));
  generatedAssigned.filter(a=>clean_(a.coverageReason).toUpperCase()==='GENERATED OFF DAY').forEach(a=>{
    const coverer=model.usersByUsername[a.username];
    const owner=model.usersByUsername[clean_(a.coverageForUsername)];
    const slot=assignmentAsSlot_(a);
    if(!owner) errors.push(a.dateKey+' '+a.shiftCode+': OFF-day coverage owner is missing.');
    if(owner&&clean_(owner.Username)===clean_(a.username)) errors.push(a.dateKey+' '+a.shiftCode+': a pharmacist cannot cover their own OFF day.');
    if(coverer&&!hasOffDayCoverageSkillForSlot_(coverer,slot)) errors.push(a.dateKey+' '+a.shiftCode+': '+a.pharmacist+' is not listed in Users -> Off-Day Coverage Skills for this shift.');
    if(owner&&!isGeneratedOffDayForHomeOwner_(owner,slot,completedState,model)) errors.push(a.dateKey+' '+a.shiftCode+': OFF-day coverage is only allowed when '+clean_(owner['Pharmacist Name'])+' has an algorithm-generated OFF day.');
  });

  // Recompute aggregate rules.
  const hours={}, days={}, evenings={}, byEmp={};
  assigned.forEach(a=>{
    const u=model.usersByUsername[a.username]; if(!u)return;
    const wk=a.username+'|'+weekStartKey_(a.date,model.settings.weekStart);
    hours[wk]=num_(hours[wk],0)+num_(a.creditedHours,0);
    days[wk]=num_(days[wk],0)+1;
    if (!isSevenOn_(u) && hours[wk] > employeeWeeklyMax_(u,model)+0.0001 && !a.manual) errors.push(a.pharmacist+' exceeds weekly maximum in week '+wk.split('|')[1]+'.');
    if (regularFiveDayRuleApplies_(u) && days[wk] > model.settings.regularWorkdaysPerWeek && !a.manual) errors.push(a.pharmacist+' exceeds the '+model.settings.regularWorkdaysPerWeek+'-workday limit in week '+wk.split('|')[1]+'.');
    if (clean_(a.shiftType).toLowerCase()==='evening') {
      const mk=a.username+'|'+monthKey_(a.date); evenings[mk]=num_(evenings[mk],0)+1;
      const vaSlot=assignmentAsSlot_(a);
      const bypassEvening=isSevenOn_(u)&&sevenOnSlotMatches_(u,vaSlot,model) || isResidentMandatoryE2_(u,vaSlot,model);
      if (!bypassEvening && evenings[mk] > employeeEveningMax_(u,model) && !a.manual) errors.push(a.pharmacist+' exceeds evening maximum in '+mk.split('|')[1]+'.');
    }
    if (!byEmp[a.username]) byEmp[a.username]=[];
    const iv=assignmentInterval_(a.date,a.shift); if(iv) byEmp[a.username].push({a,iv});
  });
  Object.keys(byEmp).forEach(u=>{
    const arr=byEmp[u].sort((x,y)=>x.iv.start-y.iv.start);
    for(let i=1;i<arr.length;i++) if(arr[i].iv.start<arr[i-1].iv.end) errors.push(arr[i].a.pharmacist+' has overlapping shifts around '+arr[i].a.dateKey+'.');
  });

  // v25 validate the hard Evening -> Morning/Day next-day rule across the
  // generated period AND the surrounding context loaded from the saved sheet.
  // This catches old/manual rows too, including a Friday evening before a
  // required Saturday morning weekend assignment.
  if(model.settings.blockEveningToMorningTransition){
    const byUserDate={};
    assigned.forEach(a=>{
      if(!a||!a.username||!a.date)return;
      if(!byUserDate[a.username])byUserDate[a.username]={};
      byUserDate[a.username][a.dateKey]=a;
    });
    Object.keys(byUserDate).forEach(username=>{
      const map=byUserDate[username];
      Object.keys(map).sort().forEach(dk=>{
        const first=map[dk];
        const nextDate=addDays_(first.date,1);
        const second=map[formatDateKey_(nextDate)];
        if(!second)return;
        const firstType=clean_(first.shiftType || (first.shift||{}).Type).toLowerCase();
        const secondType=clean_(second.shiftType || (second.shift||{}).Type).toLowerCase();
        if(firstType==='evening' && secondType==='day') {
          errors.push(first.pharmacist+' works evening shift '+first.shiftCode+' on '+first.dateKey+' and Day/Morning shift '+second.shiftCode+' on '+second.dateKey+'. Evening-to-morning transitions are not allowed; '+second.dateKey+' must be OFF or another evening shift.');
        }
      });
    });
  }


  // Regular work-pattern validation. Custom date ranges may begin/end midweek,
  // so exact five-workday validation applies only to complete Sunday-Saturday
  // weeks fully contained in the selected range. Weekend days count normally.
  const regularUsers=model.users.filter(u=>yes_(u.Active)&&regularFiveDayRuleApplies_(u));
  weekStartsInRange_(start,end,model.settings.weekStart).forEach(ws=>{
    const we=addDays_(ws,6);
    if(ws<start||we>end)return;
    regularUsers.forEach(u=>{
      const username=clean_(u.Username);
      const wk=username+'|'+weekStartKey_(ws,model.settings.weekStart);
      const count=num_(days[wk],0);
      const requiredDays=regularRequiredWorkdaysForWeek_(u,ws,model);
      const protectedDays=regularProtectedOffDaysForWeek_(u,ws,model);
      if(count!==requiredDays){
        errors.push(
          clean_(u['Pharmacist Name'])+' has '+count+
          ' workday(s) in week '+formatDateKey_(ws)+' through '+formatDateKey_(we)+
          '; exactly '+requiredDays+' workday(s) are required'+
          (protectedDays?' after '+protectedDays+' approved PTO/Regular Off day(s)':'')+'.'
        );
      }

      const actualHours=Math.round(num_(hours[wk],0)*10)/10;
      const expectedHours=Math.round(requiredDays*8*10)/10;
      if(Math.abs(actualHours-expectedHours)>0.0001){
        errors.push(
          clean_(u['Pharmacist Name'])+' has '+actualHours+
          ' scheduled hour(s) in week '+formatDateKey_(ws)+' through '+formatDateKey_(we)+
          '; '+expectedHours+' scheduled hour(s) are required'+
          (protectedDays?' after approved PTO/Regular Off':'')+'.'
        );
      }
    });
  });

  regularUsers.forEach(u=>{
    const username=clean_(u.Username);
    const dateSet=(completedState.assignedDateKeysByEmployee&&completedState.assignedDateKeysByEmployee[username])||new Set();
    const dates=[...dateSet].map(asDate_).filter(Boolean);
    const longest=longestConsecutiveDateRun_(dates);
    if(longest>model.settings.maxConsecutiveWorkdays){
      errors.push(clean_(u['Pharmacist Name'])+' has '+longest+' consecutive workdays; maximum allowed is '+model.settings.maxConsecutiveWorkdays+'.');
    }
  });

  // Required A/B/C weekend assignments are structural, not preferences. Every
  // grouped regular pharmacist must work both visible days of every assigned
  // weekend, and those assignments are counted inside the five-day weekly rule.
  const validationSaturdayKeys=[...new Set(assignments.filter(a=>a.weekend).map(a=>formatDateKey_(weekendSaturday_(a.date))))].sort();
  regularUsers.filter(u=>clean_(u['Weekend Group'])).forEach(u=>{
    const username=clean_(u.Username), group=clean_(u['Weekend Group']);
    if(!model.settings.weekendRotation.includes(group))return;
    validationSaturdayKeys.forEach(satKey=>{
      const sat=startOfDay_(asDate_(satKey));
      if(!employeeWeekendIsOn_(u,sat,model.settings))return;
      if(weekendGroupForDate_(sat,model.settings)!==group)return;
      [sat,addDays_(sat,1)].forEach(day=>{
        if(day<start||day>end)return;
        const dk=formatDateKey_(day);
        const hit=completedState.byEmployeeDate[username+'|'+dk];
        if(!hit || !isWeekendDate_(hit.date) || hit.weekendGroup!==group) errors.push(clean_(u['Pharmacist Name'])+' is Weekend Group '+group+' and must work the required weekend day '+dk+'.');
      });
    });
  });


  // Paired Emergency Department coverage: whenever both configured ED shifts
  // are required on a weekday, both must be filled by two different pharmacists.
  if(model.settings.requirePairedEdCoverage){
    const dayCode=clean_(model.settings.edDayShiftCode||'EDD').toUpperCase();
    const eveCode=clean_(model.settings.edEveningShiftCode||'EDE').toUpperCase();
    const dates=[...new Set(assignments.filter(a=>!isWeekendDate_(a.date)&&(a.shiftCode===dayCode||a.shiftCode===eveCode)).map(a=>a.dateKey))].sort();
    dates.forEach(dk=>{
      const requiredDay=assignments.filter(a=>a.dateKey===dk&&a.shiftCode===dayCode).length;
      const requiredEve=assignments.filter(a=>a.dateKey===dk&&a.shiftCode===eveCode).length;
      if(!requiredDay||!requiredEve)return;
      const dayFilled=generatedAssigned.filter(a=>a.dateKey===dk&&a.shiftCode===dayCode&&a.username);
      const eveFilled=generatedAssigned.filter(a=>a.dateKey===dk&&a.shiftCode===eveCode&&a.username);
      if(dayFilled.length<requiredDay) errors.push(dk+': paired ED coverage is missing '+dayCode+' coverage.');
      if(eveFilled.length<requiredEve) errors.push(dk+': paired ED coverage is missing '+eveCode+' coverage.');
      const usedDay=new Set(dayFilled.map(a=>a.username));
      eveFilled.forEach(a=>{
        if(usedDay.has(a.username)) errors.push(dk+': paired ED coverage must use two different pharmacists for '+dayCode+' and '+eveCode+'.');
      });
    });
  }

  // Residents: exactly one E2 per scheduling week whenever that week contains
  // an E2 staffing slot inside the generated date range.
  const e2Code=model.settings.residentE2ShiftCode;
  const e2PerWeek=model.settings.residentE2PerWeek;
  if(e2Code && e2PerWeek>0){
    const residents=model.users.filter(u=>yes_(u.Active)&&yes_(u.Resident));
    weekStartsInRange_(start,end,model.settings.weekStart).forEach(ws=>{
      const we=addDays_(ws,6);
      const rs=ws<start?start:ws, re=we>end?end:we;
      const hasE2Requirement=assignments.some(a=>a.shiftCode===e2Code&&inDateRange_(a.date,rs,re));
      if(!hasE2Requirement)return;
      residents.forEach(u=>{
        const username=clean_(u.Username);
        const count=generatedAssigned.filter(a=>a.username===username&&a.shiftCode===e2Code&&inDateRange_(a.date,rs,re)).length;
        if(count!==e2PerWeek){
          errors.push(clean_(u['Pharmacist Name'])+' has '+count+' '+e2Code+' shift(s) in week '+weekStartKey_(rs,model.settings.weekStart)+'; resident rule requires exactly '+e2PerWeek+'.');
        }
      });
    });
  }

  // Validate the regular-pharmacist period-hour requirement inside the selected generated period.
  const selectedPeriodDays=daysBetween_(start,end)+1;
  const configuredPeriodDays=num_(model.settings.raw['Default Schedule Days'],56);
  const enforceExactPeriodHours=
    !validationOptions.skipPeriodHours &&
    selectedPeriodDays>=configuredPeriodDays;

  if(enforceExactPeriodHours){
  // True 7-on/7-off pharmacists are governed by their block rotation instead.
  const periodHours={};

  model.users
    .filter(u=>yes_(u.Active)&&!isSevenOn_(u)&&!isPrnEmployee_(u))
    .forEach(u=>{
      periodHours[clean_(u.Username)] = 0;
    });

  generatedAssigned.forEach(a=>{
    if (!a.username || !inDateRange_(a.date,start,end) || periodHours[a.username]===undefined) return;

    periodHours[a.username] =
      num_(periodHours[a.username],0) +
      num_(a.creditedHours,creditedHours_(a.shift));
  });

  const exactTarget = num_(model.settings.periodHoursTarget,320);

  Object.keys(periodHours).forEach(username=>{
    const total = Math.round(num_(periodHours[username],0)*10)/10;
    const u = model.usersByUsername[username];
    const name = (u&&u['Pharmacist Name']) || username;

    if (total < exactTarget - 0.0001) {
      errors.push(
        name+
        ' has only '+total+
        ' hours in the generated schedule period and must have exactly '+
        exactTarget+' hours.'
      );
    }

    if (total > exactTarget + 0.0001) {
      errors.push(
        name+
        ' has '+total+
        ' hours in the generated schedule period and exceeds the exact '+
        exactTarget+'-hour requirement.'
      );
    }
  });

  }

  if(
    !validationOptions.skipPeriodHours &&
    selectedPeriodDays<configuredPeriodDays
  ){
    warnings.push(
      'Selected range is '+selectedPeriodDays+' day(s), shorter than the configured '+
      configuredPeriodDays+'-day schedule period. The exact '+
      num_(model.settings.periodHoursTarget,320)+
      '-hour period target is not enforced for this shorter custom range.'
    );
  }

  assignments.filter(a=>a.status==='UNFILLED').forEach(a=>warnings.push(a.dateKey+' '+a.shiftCode+' slot '+a.slot+' is UNFILLED.'));
  // Weekend-team structural rules. Every non-7-on/7-off employee with an
  // A/B/C group must have a Saturday anchor that belongs to that same group.
  model.users.filter(u=>yes_(u.Active)&&!isSevenOn_(u)&&!isPrnEmployee_(u)&&clean_(u['Weekend Group'])).forEach(u=>{
    const group=clean_(u['Weekend Group']);
    if(!model.settings.weekendRotation.includes(group)){
      errors.push((yes_(u.Resident)?'Resident ':'Employee ')+u['Pharmacist Name']+' must have Weekend Group '+model.settings.weekendRotation.join('/')+'.');
      return;
    }
    const anchor=asDate_(u['Weekend Rotation Anchor Date']);
    if(!anchor){
      errors.push(u['Pharmacist Name']+' is missing Weekend Rotation Anchor Date. Run setupPharmacyScheduler().');
      return;
    }
    if(dayIndex_(anchor)!==6) errors.push(u['Pharmacist Name']+' Weekend Rotation Anchor Date must be a Saturday.');
    if(weekendGroupForDate_(anchor,model.settings)!==group) errors.push(u['Pharmacist Name']+' Weekend Rotation Anchor Date does not match Weekend Group '+group+'.');
  });

  // 7-on/7-off employees must resolve to one dedicated shift and must have an anchor.
  const sevenUsers=model.users.filter(u=>yes_(u.Active)&&isSevenOn_(u));
  sevenUsers.forEach(u=>{
    const code=sevenOnAssignedShiftCode_(u,model);
    if(!code) errors.push('7-on/7-off employee '+u['Pharmacist Name']+' needs a dedicated shift. Set Preferred Shift Type to N1, N2, or E (or give only one of those skills).');
  });

  // For employees sharing the same 7-on/7-off shift, their ON phases must not overlap.
  const sevenGroups={};
  sevenUsers.forEach(u=>{
    const code=sevenOnAssignedShiftCode_(u,model);
    if(!code)return;
    if(!sevenGroups[code])sevenGroups[code]=[];
    sevenGroups[code].push(u);
  });
  Object.keys(sevenGroups).forEach(code=>{
    const group=sevenGroups[code];
    for(let i=0;i<group.length;i++)for(let j=i+1;j<group.length;j++){
      const a=asDate_(group[i]['Rotation Anchor Date']),b=asDate_(group[j]['Rotation Anchor Date']);
      if(!a||!b)continue;
      const mod=((daysBetween_(startOfDay_(a),startOfDay_(b))%14)+14)%14;
      if(mod!==7){
        errors.push('7-on/7-off '+code+' employees '+group[i]['Pharmacist Name']+' and '+group[j]['Pharmacist Name']+' must have opposite rotations (anchor dates 7 days apart modulo 14).');
      }
    }
  });

  return {errors:unique_(errors),warnings:unique_(warnings)};
}

function buildPeriodHourSummary_(assignments,model,start,end) {
  const totals={};

  model.users
    .filter(u=>yes_(u.Active))
    .forEach(u=>{
      totals[clean_(u.Username)]={
        username:clean_(u.Username),
        pharmacist:clean_(u['Pharmacist Name']),
        hours:0,
        target:isPrnEmployee_(u)?null:num_(model.settings.periodHoursTarget,320),
        exact:isPrnEmployee_(u)
      };
    });

  assignments
    .filter(a=>a.status!=='UNFILLED'&&a.username&&inDateRange_(a.date,start,end))
    .forEach(a=>{
      if(!totals[a.username]) return;
      totals[a.username].hours += num_(a.creditedHours,creditedHours_(a.shift));
    });

  return Object.keys(totals).map(k=>{
    const x=totals[k];
    x.hours=Math.round(x.hours*10)/10;
    const u=model.usersByUsername[x.username];
    x.exact=(isSevenOn_(u)||isPrnEmployee_(u))?true:Math.abs(x.hours-x.target)<0.0001;
    x.pattern=isSevenOn_(u)?'7-on/7-off':isPrnEmployee_(u)?'PRN availability':'5 days/week';
    return x;
  });
}

function saveSchedulePeriod_(existing,results,start,end,mode) {
  const sh=getDb_().getSheetByName(APP.SHEETS.SCHEDULE);
  const keep=existing.filter(r=>!inDateRange_(asDate_(r.Date),start,end));
  const out=keep.concat(results.map(assignmentToScheduleObject_));
  const values=[APP.HEADERS.SCHEDULE].concat(out.map(r=>APP.HEADERS.SCHEDULE.map(h=>sheetValue_(r[h]))));
  sh.clearContents();
  sh.getRange(1,1,values.length,values[0].length).setValues(values);
  sh.setFrozenRows(1);
  if (values.length>1) sh.getRange(2,3,values.length-1,1).setNumberFormat('m/d/yyyy');
}

function assignmentToScheduleObject_(a) {
  return {
    'Generation ID':a.generationId,'Assignment ID':a.assignmentId,'Date':a.date,'Day':a.day,'Shift':a.shiftCode,'Slot':a.slot,
    'Assigned Pharmacist':a.status==='UNFILLED'?'UNFILLED':a.pharmacist,'Username':a.username,'Hours':a.hours,'Credited Hours':a.creditedHours,
    'Required Skill':a.requiredSkill,'Coverage For Pharmacist':a.coverageForPharmacist||'','Coverage For Username':a.coverageForUsername||'','Coverage Reason':a.coverageReason||'',
    'Shift Type':a.shiftType,'Weekend':a.weekend?'Yes':'No','Weekend Group':a.weekendGroup,'Holiday':a.holiday||'',
    'Locked':a.locked?'Yes':'No','Manual':a.manual?'Yes':'No','Status':a.status,'Warning':a.warning||'','Updated At':a.updatedAt||new Date(),
    'Updated By':a.updatedBy||'','Finalized At':a.finalizedAt||''
  };
}

function assignmentIsLocked_(r) {
  if(!r)return false;
  return yes_(r.Locked) ||
    yes_(r.locked) ||
    clean_(r.Status || r.status).toUpperCase()==='LOCKED';
}

function lockedFingerprint_(r) {
  return [
    clean_(r['Assignment ID'] || r.assignmentId),
    dateKey_(r.Date || r.date),
    clean_(r.Shift || r.shiftCode),
    String(num_(r.Slot!==undefined?r.Slot:r.slot,1)),
    clean_(r.Username || r.username),
    clean_(r['Assigned Pharmacist'] || r.pharmacist)
  ].join('|');
}

function compareLockedSnapshots_(before,after) {
  const b=before.slice().sort(),a=after.slice().sort();
  if(JSON.stringify(a)===JSON.stringify(b))return [];

  const beforeSet=new Set(b);
  const afterSet=new Set(a);
  const missing=b.filter(x=>!afterSet.has(x));
  const added=a.filter(x=>!beforeSet.has(x));
  const details=[];

  if(missing.length)details.push(missing.length+' locked assignment(s) missing or changed');
  if(added.length)details.push(added.length+' unexpected locked assignment(s) appeared');

  return details.length?details:['One or more locked assignments changed.'];
}


function openShiftForceRuleMeta_(code) {
  const map={
    INACTIVE:{label:'Employee is inactive',weight:1200,risk:'CRITICAL'},
    MISSING_SKILL:{label:'Missing required skill',weight:1000,risk:'CRITICAL'},
    PTO:{label:'Approved PTO',weight:950,risk:'CRITICAL'},
    REGULAR_OFF:{label:'Approved Regular Off',weight:950,risk:'CRITICAL'},
    ALREADY_SCHEDULED:{label:'Already scheduled that day',weight:900,risk:'CRITICAL'},
    TIME_CONFLICT:{label:'Time conflict / overlapping shift',weight:900,risk:'CRITICAL'},
    UNAVAILABLE:{label:'Marked unavailable',weight:750,risk:'HIGH'},
    WEEKLY_AVAILABILITY_DAY:{label:'Recurring weekly day is unavailable',weight:725,risk:'HIGH'},
    WEEKLY_AVAILABILITY_TIME:{label:'Outside recurring available hours',weight:700,risk:'HIGH'},
    SEVEN_OFF:{label:'7-on/7-off employee is in OFF block',weight:650,risk:'HIGH'},
    SEVEN_WRONG_SHIFT:{label:'7-on/7-off dedicated-shift rule',weight:650,risk:'HIGH'},
    EVENING_NOT_ELIGIBLE:{label:'Not eligible for evening shifts',weight:600,risk:'HIGH'},
    NIGHT_NOT_ELIGIBLE:{label:'Not eligible for night shifts',weight:600,risk:'HIGH'},
    WEEKEND_NOT_ELIGIBLE:{label:'Not eligible for weekend shifts',weight:600,risk:'HIGH'},
    RESIDENT_RESTRICTION:{label:'Resident assignment restriction',weight:550,risk:'HIGH'},
    PRECEPTOR_WEEKDAY_EVENING:{label:'Preceptor weekday-evening restriction',weight:450,risk:'MODERATE'},
    PRECEPTOR_EVENING:{label:'Preceptor evening restriction',weight:450,risk:'MODERATE'},
    EVENING_TO_MORNING:{label:'Evening → next-day morning/day rule',weight:425,risk:'MODERATE'},
    CONSECUTIVE_DAYS:{label:'Maximum consecutive workdays',weight:375,risk:'MODERATE'},
    WEEKLY_DAYS:{label:'Five-workday weekly limit',weight:350,risk:'MODERATE'},
    WEEKLY_HOURS:{label:'Weekly-hour maximum',weight:325,risk:'MODERATE'},
    RESIDENT_WRONG_WEEKEND_GROUP:{label:'Resident weekend-group rule',weight:300,risk:'MODERATE'},
    WRONG_WEEKEND_GROUP:{label:'Weekend-group rotation',weight:275,risk:'MODERATE'},
    WEEKEND_ANCHOR_OFF:{label:'One-weekend-every-three-weeks rotation',weight:275,risk:'MODERATE'},
    TWO_MONTH_HOURS:{label:'Schedule-period hour target',weight:250,risk:'MODERATE'},
    RESIDENT_E2_WEEKLY_LIMIT:{label:'Resident E2 weekly limit',weight:200,risk:'LOW'},
    EVENING_LIMIT:{label:'Monthly evening-shift limit',weight:175,risk:'LOW'},
    CUSTOM_HOURS:{label:'Outside pharmacist custom work hours',weight:175,risk:'LOW'}
  };

  return map[clean_(code).toUpperCase()] || {
    label:reasonToWarning_(code),
    weight:300,
    risk:'MODERATE'
  };
}

function openShiftForceRisk_(rules) {
  const weights=(rules||[]).map(x=>num_(x.weight,0));
  const max=weights.length?Math.max.apply(null,weights):0;
  if(max>=800)return 'CRITICAL';
  if(max>=500)return 'HIGH';
  if(max>=250)return 'MODERATE';
  return max>0?'LOW':'SAFE';
}

function buildOpenShiftForceSuggestions_(openRows,model,allRows) {
  openRows=Array.isArray(openRows)?openRows:[];
  if(!openRows.length)return [];

  /*
   * PERFORMANCE: build the schedule state ONCE for the entire selected set.
   * The old implementation rebuilt ~1,500 saved assignments separately for
   * every open shift and then rescanned the full schedule for every pharmacist.
   * In the GitHub browser runtime that blocked Chrome's main thread.
   */
  const generationIds=new Set(
    openRows
      .map(r=>clean_(r['Generation ID']))
      .filter(id=>id&&id.indexOf('MANUAL_')!==0)
  );

  const periodRows=generationIds.size
    ? allRows.filter(r=>generationIds.has(clean_(r['Generation ID']))&&asDate_(r.Date))
    : allRows.filter(r=>asDate_(r.Date));

  const periodDates=periodRows
    .map(r=>startOfDay_(asDate_(r.Date)))
    .filter(Boolean)
    .sort((a,b)=>a-b);

  const openDates=openRows
    .map(r=>startOfDay_(asDate_(r.Date)))
    .filter(Boolean)
    .sort((a,b)=>a-b);

  const periodStart=periodDates.length
    ? periodDates[0]
    : (openDates[0]||startOfDay_(new Date()));
  const periodEnd=periodDates.length
    ? periodDates[periodDates.length-1]
    : (openDates[openDates.length-1]||periodStart);

  const fixedRows=allRows.filter(r=>
    clean_(r.Status).toUpperCase()!=='UNFILLED' &&
    clean_(r['Assigned Pharmacist']).toUpperCase()!=='UNFILLED' &&
    clean_(r.Username)
  );

  const state=createState_(model,fixedRows,periodStart,periodEnd);

  return openRows.map(row=>{
    const d=startOfDay_(asDate_(row.Date));
    const shiftCode=clean_(row.Shift).toUpperCase();
    const shift=model.shiftMap[shiftCode]||model.shiftMap[clean_(row.Shift)];
    const assignmentId=clean_(row['Assignment ID']);
    const slotNumber=num_(row.Slot,1);

    const base={
      assignmentId:assignmentId,
      generationId:clean_(row['Generation ID']),
      date:d?formatDateKey_(d):dateKey_(row.Date),
      shift:shiftCode,
      slot:slotNumber,
      requiredSkill:clean_(row['Required Skill']||(shift?shift.Skill:'')),
      locked:yes_(row.Locked),
      finalized:!!asDate_(row['Finalized At']),
      originalReason:clean_(row.Warning),
      candidates:[]
    };

    if(base.finalized){
      base.actionRequired='UNFINALIZE';
      base.actionMessage='Unfinalize this schedule period before an administrator can override this shift.';
      return base;
    }

    if(base.locked){
      base.actionRequired='UNLOCK';
      base.actionMessage='Unlock this open shift before using an administrator override.';
      return base;
    }

    if(!d||!shift){
      base.actionRequired='FIX_SHIFT';
      base.actionMessage='The shift definition or date is invalid, so NeoChrono cannot build an override recommendation.';
      return base;
    }

    const slot={
      generationId:clean_(row['Generation ID'])||'MANUAL',
      date:d,
      dateKey:formatDateKey_(d),
      day:dayName_(d),
      shiftCode:shiftCode,
      slot:slotNumber,
      shift:shift,
      requiredSkill:clean_(row['Required Skill']||shift.Skill),
      weekend:yes_(row.Weekend)||isWeekendDate_(d),
      weekendGroup:clean_(row['Weekend Group'])||(isWeekendDate_(d)?weekendGroupForDate_(d,model.settings):''),
      slotKey:slotKey_(formatDateKey_(d),shiftCode,slotNumber)
    };

    const candidates=[];

    // Override candidates are also restricted to pharmacists who actually
    // possess the required skill before any expensive rule analysis runs.
    skillQualifiedUsersForSlot_(slot,model).forEach(u=>{
      const normal=eligibility_(u,slot,model,state,true);
      let effective=normal;
      let coverage={
        applies:false,
        coverageForPharmacist:'',
        coverageForUsername:'',
        coverageReason:'',
        message:''
      };

      if(
        (normal.reasons||[]).indexOf('MISSING_SKILL')>=0 &&
        hasOffDayCoverageSkillForSlot_(u,slot)
      ){
        const resolved=resolveManualOffDayCoverageContext_(u,slot,model,state);
        coverage={
          applies:!!resolved.applies,
          coverageForPharmacist:clean_(resolved.coverageForPharmacist),
          coverageForUsername:clean_(resolved.coverageForUsername),
          coverageReason:clean_(resolved.coverageReason),
          message:clean_(resolved.message)
        };
        if(resolved.applies&&resolved.eligibility){
          effective=resolved.eligibility;
        }
      }

      const reasonCodes=(effective.reasons||[]).slice();
      const rules=reasonCodes.map(code=>{
        const meta=openShiftForceRuleMeta_(code);
        return {
          code:code,
          label:meta.label,
          weight:meta.weight,
          risk:meta.risk,
          detail:reasonToWarning_(code)
        };
      });

      const warnings=(effective.warnings||[]).slice();
      if(
        !coverage.applies &&
        coverage.message &&
        (normal.reasons||[]).indexOf('MISSING_SKILL')>=0
      ){
        warnings.push(coverage.message);
      }

      const hoursSummary=openShiftProjectedHoursFromState_(
        u,
        shift,
        d,
        state,
        model
      );

      let overrideCost=rules.reduce((sum,x)=>sum+num_(x.weight,0),0);
      overrideCost+=warnings.length*60;

      const hasNormalSkill=hasRequiredSkillForSlot_(u,slot,model);
      if(preferredShiftMatches_(u,slot))overrideCost-=20;

      const normalScore=scoreCandidate_(
        u,
        slot,
        model,
        state,
        {
          ok:true,
          reasons:reasonCodes,
          warnings:warnings,
          availabilityPreferred:effective.availabilityPreferred,
          weeklyPreferred:effective.weeklyPreferred,
          weeklySoftMismatch:effective.weeklySoftMismatch
        }
      );

      const sameDayConflict=
        reasonCodes.indexOf('ALREADY_SCHEDULED')>=0 ||
        reasonCodes.indexOf('TIME_CONFLICT')>=0;

      candidates.push({
        username:clean_(u.Username),
        pharmacist:clean_(u['Pharmacist Name']),
        preferred:preferredShiftMatches_(u,slot),
        hasNormalSkill:hasNormalSkill,
        coverageOnly:!!coverage.applies,
        coverageForPharmacist:coverage.applies?coverage.coverageForPharmacist:'',
        reasonCodes:reasonCodes,
        rules:rules,
        warnings:warnings,
        risk:openShiftForceRisk_(rules),
        overrideCost:Math.max(0,overrideCost),
        schedulerScore:normalScore,
        hoursSummary:hoursSummary,
        sameDayConflict:sameDayConflict
      });
    });

    /*
     * Recommendation rule for Help Fill Open Shifts:
     * A pharmacist who is already working that date / has a time conflict is
     * LAST RESORT, even if they possess the required skill. Keep them in the
     * dropdown so the administrator can still force the assignment, but never
     * recommend them while another skilled pharmacist is not already working.
     */
    candidates.sort((a,b)=>
      Number(!!a.sameDayConflict)-Number(!!b.sameDayConflict) ||
      a.overrideCost-b.overrideCost ||
      b.schedulerScore-a.schedulerScore ||
      clean_(a.pharmacist).localeCompare(clean_(b.pharmacist))
    );

    base.actionRequired='OVERRIDE';
    base.candidates=candidates.slice(0,5);
    base.recommended=base.candidates.length?base.candidates[0]:null;

    if(base.recommended){
      const r=base.recommended;
      base.recommendation=r.sameDayConflict
        ? r.pharmacist+' is a last-resort skilled option because every higher-ranked skilled pharmacist is already working or otherwise more restricted.'
        : (r.rules||[]).length
          ? r.pharmacist+' is the lowest-impact skilled override currently available and would break '+r.rules.length+' rule'+(r.rules.length===1?'':'s')+'.'
          : r.pharmacist+' currently passes the scheduling rules.';
    }else{
      base.recommendation='No active pharmacist is available for an override recommendation.';
    }

    return base;
  });
}

function getOpenShiftOverrideSuggestions(token,assignmentIds) {
  requireAdmin_(token);
  const ids=new Set(
    (Array.isArray(assignmentIds)?assignmentIds:[])
      .map(clean_)
      .filter(Boolean)
  );

  if(!ids.size){
    return {ok:false,suggestions:[],message:'Select at least one open shift.'};
  }

  const allRows=readTable_(APP.SHEETS.SCHEDULE);
  const rows=allRows.filter(r=>ids.has(clean_(r['Assignment ID'])));

  const missing=[...ids].filter(id=>!rows.some(r=>clean_(r['Assignment ID'])===id));
  if(missing.length){
    return {
      ok:false,
      suggestions:[],
      message:missing.length+' selected open shift(s) are no longer present. Refresh the helper and try again.'
    };
  }

  const model=loadSchedulingModel_();
  return serialize_({
    ok:true,
    suggestions:buildOpenShiftForceSuggestions_(rows,model,allRows)
  });
}


function openShiftProjectedHoursFromState_(u,shift,date,state,model) {
  const username=clean_(u.Username);
  const proposed=creditedHours_(shift);
  const weekStart=startOfDay_(asDate_(weekStartKey_(date,model.settings.weekStart)));
  const weekEnd=addDays_(weekStart,6);
  const weekKey=username+'|'+weekStartKey_(date,model.settings.weekStart);
  const weekCurrent=num_(state.weeklyHours[weekKey],0);
  const weekDaysCurrent=num_(state.weeklyDays[weekKey],0);
  const weeklyMax=employeeWeeklyMax_(u,model);
  const periodCurrent=num_(state.periodHours[username],0);
  const periodTarget=num_(model.settings.periodHoursTarget,320);

  function r1_(n){return Math.round(num_(n,0)*10)/10;}

  return {
    pharmacist:clean_(u['Pharmacist Name']),
    proposedShift:clean_(shift.Shift),
    proposedHours:r1_(proposed),
    week:{
      start:formatDateKey_(weekStart),
      end:formatDateKey_(weekEnd),
      current:r1_(weekCurrent),
      projected:r1_(weekCurrent+proposed),
      maximum:r1_(weeklyMax),
      currentDays:weekDaysCurrent,
      projectedDays:weekDaysCurrent+1,
      requiredDays:regularFiveDayRuleApplies_(u)?model.settings.regularWorkdaysPerWeek:null
    },
    period:{
      start:state.periodStart?formatDateKey_(state.periodStart):'',
      end:state.periodEnd?formatDateKey_(state.periodEnd):'',
      current:r1_(periodCurrent),
      projected:r1_(periodCurrent+proposed),
      target:r1_(periodTarget)
    }
  };
}

function evaluateBulkOpenShiftOverrides_(selections,model,allRows) {
  selections=Array.isArray(selections)?selections:[];
  if(!selections.length){
    return {ok:false,errors:['Select at least one open shift to override.'],items:[]};
  }

  const errors=[];
  const prepared=[];

  selections.forEach((choice,index)=>{
    const assignmentId=clean_(choice&&choice.assignmentId);
    const username=clean_(choice&&choice.username);
    const row=assignmentId
      ? allRows.find(r=>clean_(r['Assignment ID'])===assignmentId)
      : null;

    if(!row){
      errors.push('Selection '+(index+1)+': open-shift row was not found.');
      return;
    }

    const status=clean_(row.Status).toUpperCase();
    const assigned=clean_(row['Assigned Pharmacist']).toUpperCase();
    if(status!=='UNFILLED'&&assigned!=='UNFILLED'){
      errors.push(dateKey_(row.Date)+' '+clean_(row.Shift)+': this shift is no longer UNFILLED.');
      return;
    }

    if(asDate_(row['Finalized At'])){
      errors.push(dateKey_(row.Date)+' '+clean_(row.Shift)+': unfinalize the schedule before overriding this shift.');
      return;
    }

    if(yes_(row.Locked)){
      errors.push(dateKey_(row.Date)+' '+clean_(row.Shift)+': unlock this shift before overriding it.');
      return;
    }

    const d=startOfDay_(asDate_(row.Date));
    const shift=model.shiftMap[clean_(row.Shift).toUpperCase()] || model.shiftMap[clean_(row.Shift)];
    const u=model.usersByUsername[username];

    if(!d||!shift){
      errors.push(dateKey_(row.Date)+' '+clean_(row.Shift)+': invalid date or shift definition.');
      return;
    }

    if(!u||!yes_(u.Active)){
      errors.push(dateKey_(row.Date)+' '+clean_(row.Shift)+': selected pharmacist is not an active scheduling record.');
      return;
    }

    const manualPeriod=resolveManualSchedulePeriod_(
      {
        assignmentId:assignmentId,
        generationId:clean_(row['Generation ID']),
        date:formatDateKey_(d),
        shift:clean_(shift.Shift),
        slot:num_(row.Slot,1)
      },
      allRows,
      d,
      model
    );

    prepared.push({
      inputIndex:index,
      choice:choice,
      row:row,
      user:u,
      shift:shift,
      date:d,
      manualPeriod:manualPeriod
    });
  });

  if(errors.length){
    return {ok:false,errors:errors,items:[]};
  }

  prepared.sort((a,b)=>
    formatDateKey_(a.date).localeCompare(formatDateKey_(b.date)) ||
    clean_(a.shift.Shift).localeCompare(clean_(b.shift.Shift)) ||
    num_(a.row.Slot,1)-num_(b.row.Slot,1)
  );

  const periodStarts=prepared.map(x=>x.manualPeriod.start).filter(Boolean).sort((a,b)=>a-b);
  const periodEnds=prepared.map(x=>x.manualPeriod.end).filter(Boolean).sort((a,b)=>a-b);
  const periodStart=periodStarts.length?periodStarts[0]:prepared[0].date;
  const periodEnd=periodEnds.length?periodEnds[periodEnds.length-1]:prepared[prepared.length-1].date;

  const selectedIds=new Set(prepared.map(x=>clean_(x.row['Assignment ID'])));
  const baseSchedule=allRows.filter(r=>
    !selectedIds.has(clean_(r['Assignment ID'])) &&
    clean_(r.Status).toUpperCase()!=='UNFILLED' &&
    clean_(r['Assigned Pharmacist']).toUpperCase()!=='UNFILLED' &&
    clean_(r.Username)
  );

  const state=createState_(model,baseSchedule,periodStart,periodEnd);
  const items=[];

  prepared.forEach(x=>{
    const row=x.row;
    const u=x.user;
    const shift=x.shift;
    const d=x.date;
    const slotNumber=num_(row.Slot,1);
    const shiftCode=clean_(shift.Shift).toUpperCase();

    const slot={
      generationId:clean_(row['Generation ID'])||'MANUAL',
      date:d,
      dateKey:formatDateKey_(d),
      day:dayName_(d),
      shiftCode:shiftCode,
      slot:slotNumber,
      shift:shift,
      requiredSkill:clean_(row['Required Skill']||shift.Skill),
      weekend:yes_(row.Weekend)||isWeekendDate_(d),
      weekendGroup:clean_(row['Weekend Group'])||(isWeekendDate_(d)?weekendGroupForDate_(d,model.settings):''),
      slotKey:slotKey_(formatDateKey_(d),shiftCode,slotNumber)
    };

    const normal=eligibility_(u,slot,model,state,true);
    let effective=normal;
    let coverage={
      applies:false,
      coverageForPharmacist:'',
      coverageForUsername:'',
      coverageReason:'',
      message:''
    };

    if(
      (normal.reasons||[]).indexOf('MISSING_SKILL')>=0 &&
      hasOffDayCoverageSkillForSlot_(u,slot)
    ){
      const resolved=resolveManualOffDayCoverageContext_(u,slot,model,state);
      coverage={
        applies:!!resolved.applies,
        coverageForPharmacist:clean_(resolved.coverageForPharmacist),
        coverageForUsername:clean_(resolved.coverageForUsername),
        coverageReason:clean_(resolved.coverageReason),
        message:clean_(resolved.message)
      };
      if(resolved.applies&&resolved.eligibility){
        effective=resolved.eligibility;
      }
    }

    const reasonCodes=(effective.reasons||[]).slice();
    const rules=reasonCodes.map(code=>{
      const meta=openShiftForceRuleMeta_(code);
      return {
        code:code,
        label:meta.label,
        weight:meta.weight,
        risk:meta.risk,
        detail:reasonToWarning_(code)
      };
    });

    const warnings=(effective.warnings||[]).slice();
    if(
      !coverage.applies &&
      coverage.message &&
      (normal.reasons||[]).indexOf('MISSING_SKILL')>=0
    ){
      warnings.push(coverage.message);
    }

    const hoursSummary=openShiftProjectedHoursFromState_(u,shift,d,state,model);

    const assigned=makeAssigned_(
      slot,
      u,
      warnings,
      'BULK_OVERRIDE_PREVIEW'
    );

    assigned.assignmentId=clean_(row['Assignment ID']);
    assigned.generationId=clean_(row['Generation ID'])||assigned.generationId;
    assigned.coverageForPharmacist=coverage.applies?coverage.coverageForPharmacist:'';
    assigned.coverageForUsername=coverage.applies?coverage.coverageForUsername:'';
    assigned.coverageReason=coverage.applies?coverage.coverageReason:'';
    assigned.holiday=clean_(row.Holiday);
    assigned.manual=true;
    assigned.status='MANUAL';

    /*
     * Add the hypothetical forced assignment even when it breaks rules. This
     * makes every later selected shift see the workload and conflicts created
     * by earlier selections in the same bulk override.
     */
    addAssignmentToState_(state,assigned,model);

    items.push({
      inputIndex:x.inputIndex,
      assignmentId:clean_(row['Assignment ID']),
      generationId:clean_(row['Generation ID']),
      date:formatDateKey_(d),
      shift:shiftCode,
      slot:slotNumber,
      requiredSkill:clean_(slot.requiredSkill),
      username:clean_(u.Username),
      pharmacist:clean_(u['Pharmacist Name']),
      reasonCodes:reasonCodes,
      rules:rules,
      warnings:warnings,
      risk:openShiftForceRisk_(rules),
      hoursSummary:hoursSummary,
      coverage:coverage,
      originalWarning:clean_(row.Warning)
    });
  });

  return {
    ok:true,
    errors:[],
    count:items.length,
    items:items
  };
}

function previewBulkOpenShiftOverrides(token,selections) {
  requireAdmin_(token);
  const model=loadSchedulingModel_();
  const allRows=readTable_(APP.SHEETS.SCHEDULE);
  return serialize_(evaluateBulkOpenShiftOverrides_(selections,model,allRows));
}

function saveBulkOpenShiftOverrides(token,selections,overrideReason) {
  const ctx=requireAdmin_(token);
  overrideReason=clean_(overrideReason);
  selections=Array.isArray(selections)?selections:[];

  if(!selections.length){
    throw new Error('Select at least one shift to override.');
  }

  const lock=LockService.getScriptLock();
  if(!lock.tryLock(30000)){
    throw new Error('Schedule is currently being modified by another administrator.');
  }

  try{
    ensureUniqueScheduleAssignmentIds_();

    const model=loadSchedulingModel_();
    let allRows=readTable_(APP.SHEETS.SCHEDULE);

    /*
     * FORCE OVERRIDE MODE:
     * The administrator already reviewed the impact before reaching Confirm.
     * Do NOT run eligibility/rule evaluation again here. Re-evaluating at save
     * time can only introduce another blocker and defeats the meaning of an
     * explicit administrator override.
     *
     * Only structural requirements remain: the target row, pharmacist record,
     * date, and shift definition must exist so a valid Schedule row can be
     * written. Hours/weekend/PTO/consecutive-day/transition/etc. rules are
     * warnings, never save gates.
     */
    const applied=[];

    selections.forEach((selection,index)=>{
      selection=selection||{};

      let assignmentId=clean_(selection.assignmentId);
      let current=assignmentId
        ? findRowByKey_(APP.SHEETS.SCHEDULE,'Assignment ID',assignmentId)
        : null;

      /*
       * Fallback lookup makes the forced override resilient if an older open
       * row received an Assignment ID during the ID-repair pass.
       */
      if(!current){
        const wantedDate=dateKey_(selection.date);
        const wantedShift=clean_(selection.shift).toUpperCase();
        const wantedSlot=num_(selection.slot,1);

        current=allRows.find(r=>
          dateKey_(r.Date)===wantedDate &&
          clean_(r.Shift).toUpperCase()===wantedShift &&
          num_(r.Slot,1)===wantedSlot
        )||null;

        if(current){
          assignmentId=clean_(current['Assignment ID']);
        }
      }

      if(!current){
        throw new Error(
          'Selected shift '+(index+1)+' could not be found in the Schedule sheet.'
        );
      }

      const username=clean_(selection.username);
      const u=model.usersByUsername[username];

      if(!u){
        throw new Error(
          dateKey_(current.Date)+' '+clean_(current.Shift)+
          ': selected pharmacist record was not found.'
        );
      }

      const shiftCode=clean_(selection.shift||current.Shift).toUpperCase();
      const shift=model.shiftMap[shiftCode]||model.shiftMap[clean_(selection.shift||current.Shift)];

      if(!shift){
        throw new Error(
          dateKey_(current.Date)+' '+shiftCode+
          ': shift definition was not found.'
        );
      }

      const d=startOfDay_(asDate_(selection.date||current.Date));
      if(!d){
        throw new Error(
          clean_(current.Shift)+': selected shift date is invalid.'
        );
      }

      const info=null;
      const coverage={};

      const warningParts=[
        overrideReason ? ('ADMIN FORCE OVERRIDE: '+overrideReason) : 'ADMIN FORCE OVERRIDE',
        'FORCED REGARDLESS OF SCHEDULING RULES'
      ];

      if(info){
        (info.rules||[]).forEach(rule=>{
          warningParts.push(rule.detail||rule.label||rule.code);
        });
        (info.warnings||[]).forEach(w=>warningParts.push(w));
      }

      const previousWarning=clean_(current.Warning);
      if(previousWarning){
        warningParts.push('PREVIOUS OPEN-SHIFT REASON: '+previousWarning);
      }

      if(coverage.applies){
        warningParts.unshift(
          'OFF-DAY COVERAGE ONLY: Covering '+
          clean_(coverage.coverageForPharmacist)+
          ' on an algorithm-generated OFF day.'
        );
      }

      /*
       * This is the authoritative administrator override write.
       * Locked/finalized/status/hour/weekend/etc. scheduling rules are not
       * checked here because the administrator already reviewed and confirmed
       * the override. The row is explicitly converted to a MANUAL assignment.
       */
      const row={
        'Generation ID':
          clean_(current['Generation ID']) ||
          clean_(selection.generationId) ||
          'MANUAL_'+Utilities.formatDate(new Date(),getTz_(),'yyyyMMdd_HHmmss'),
        'Assignment ID':assignmentId,
        'Date':d,
        'Day':dayName_(d),
        'Shift':clean_(shift.Shift),
        'Slot':num_(selection.slot,current.Slot||1),
        'Assigned Pharmacist':clean_(u['Pharmacist Name']),
        'Username':clean_(u.Username),
        'Hours':num_(shift.Hours,0),
        'Credited Hours':creditedHours_(shift),
        'Required Skill':clean_(shift.Skill),
        'Coverage For Pharmacist':coverage.applies?clean_(coverage.coverageForPharmacist):'',
        'Coverage For Username':coverage.applies?clean_(coverage.coverageForUsername):'',
        'Coverage Reason':coverage.applies?clean_(coverage.coverageReason):'',
        'Shift Type':clean_(shift.Type),
        'Weekend':isWeekendDate_(d)?'Yes':'No',
        'Weekend Group':isWeekendDate_(d)?weekendGroupForDate_(d,model.settings):'',
        'Holiday':clean_(current.Holiday),
        'Locked':'No',
        'Manual':'Yes',
        'Status':'MANUAL',
        'Warning':warningParts.join(' | '),
        'Updated At':new Date(),
        'Updated By':ctx.username,
        'Finalized At':''
      };

      const updated=updateRowByKey_(
        APP.SHEETS.SCHEDULE,
        'Assignment ID',
        assignmentId,
        row
      );

      if(!updated){
        throw new Error(
          dateKey_(d)+' '+clean_(shift.Shift)+
          ': the forced override could not be written.'
        );
      }

      applied.push({
        assignmentId:assignmentId,
        date:formatDateKey_(d),
        shift:clean_(shift.Shift),
        slot:num_(selection.slot,current.Slot||1),
        username:clean_(u.Username),
        pharmacist:clean_(u['Pharmacist Name']),
        reasonCodes:[],
        rules:[],
        warnings:info?(info.warnings||[]):[]
      });
    });

    reconcileFilledVsUnfilledScheduleRows_();
    SpreadsheetApp.flush();

    applied.forEach(item=>{
      const verify=findRowByKey_(
        APP.SHEETS.SCHEDULE,
        'Assignment ID',
        item.assignmentId
      );

      if(
        !verify ||
        clean_(verify.Username)!==clean_(item.username) ||
        clean_(verify.Status).toUpperCase()==='UNFILLED'
      ){
        throw new Error(
          item.date+' '+item.shift+
          ': the forced override could not be verified after saving.'
        );
      }

      const details=[];
      (item.rules||[]).forEach(rule=>{
        details.push(rule.detail||rule.label||rule.code);
      });
      (item.warnings||[]).forEach(w=>details.push(w));

      audit_(
        'ASSIGNMENT_CHANGED',
        item.date,
        item.pharmacist,
        item.shift,
        item.shift,
        'UNFILLED',
        item.pharmacist,
        'Yes',
        overrideReason,
        'FORCED BULK OVERRIDE | '+details.join(' | '),
        ctx.username
      );
    });

    audit_(
      'BULK_OPEN_SHIFT_FORCE_OVERRIDE',
      applied.length?applied[0].date:'',
      '',
      '',
      '',
      '',
      String(applied.length),
      'Yes',
      overrideReason,
      JSON.stringify(applied.map(item=>({
        date:item.date,
        shift:item.shift,
        slot:item.slot,
        pharmacist:item.pharmacist,
        rules:item.reasonCodes||[]
      }))),
      ctx.username
    );

    SpreadsheetApp.flush();

    return serialize_({
      ok:true,
      forced:true,
      appliedCount:applied.length,
      overrideReason:overrideReason,
      items:applied,
      message:
        applied.length+
        ' selected shift(s) were force-assigned as administrator overrides.'
    });
  } finally {
    lock.releaseLock();
  }
}


/**
 * Preview or apply safe assignments for rows that are already UNFILLED.
 *
 * IMPORTANT:
 * - Existing filled assignments are never moved, swapped, deleted, or replaced.
 * - The same eligibility_() and scoreCandidate_() functions used by the normal
 *   scheduler are used here.
 * - Finalized or locked UNFILLED rows are reported but not changed.
 * - mode = PREVIEW builds a plan only.
 * - mode = APPLY writes only the planned UNFILLED rows.
 */
function fillExistingUnfilledShifts(token,startDate,endDate,mode) {
  const ctx=requireAdmin_(token);
  mode=clean_(mode||'PREVIEW').toUpperCase();
  const apply=mode==='APPLY';

  const lock=LockService.getScriptLock();
  if(!lock.tryLock(apply?30000:10000)){
    throw new Error('Schedule is currently being modified by another administrator.');
  }

  try{
    // Opening Help Fill Open Shifts is read-only. Do not run the full Schedule
    // ID migration unless the helper is actually applying changes.
    if(apply)ensureUniqueScheduleAssignmentIds_();

    const model=loadSchedulingModel_();
    const allRows=readTable_(APP.SHEETS.SCHEDULE);

    const start=startOfDay_(asDate_(startDate));
    const end=startOfDay_(asDate_(endDate));

    if(!start||!end)throw new Error('A valid start date and end date are required.');
    if(end<start)throw new Error('End Date must be on or after Start Date.');

    const selectedOpen=allRows.filter(r=>{
      const d=asDate_(r.Date);
      if(!d||!inDateRange_(d,start,end))return false;
      return clean_(r.Status).toUpperCase()==='UNFILLED' ||
        clean_(r['Assigned Pharmacist']).toUpperCase()==='UNFILLED';
    });

    const skippedFinalized=selectedOpen.filter(r=>!!asDate_(r['Finalized At']));
    const skippedLocked=selectedOpen.filter(r=>!asDate_(r['Finalized At'])&&yes_(r.Locked));
    const editable=selectedOpen.filter(r=>!asDate_(r['Finalized At'])&&!yes_(r.Locked));

    /*
     * Resolve the original generated period so the exact period-hours rule is
     * evaluated against the same schedule window that produced these rows.
     */
    const generationIds=new Set(
      editable
        .map(r=>clean_(r['Generation ID']))
        .filter(id=>id&&id.indexOf('MANUAL_')!==0)
    );

    const periodRows=generationIds.size
      ? allRows.filter(r=>generationIds.has(clean_(r['Generation ID']))&&asDate_(r.Date))
      : allRows.filter(r=>inDateRange_(asDate_(r.Date),start,end));

    const periodDates=periodRows
      .map(r=>startOfDay_(asDate_(r.Date)))
      .filter(Boolean)
      .sort((a,b)=>a-b);

    const periodStart=periodDates.length?periodDates[0]:start;
    const periodEnd=periodDates.length?periodDates[periodDates.length-1]:end;

    /*
     * Load every existing FILLED row into state. Weekly/monthly maps are keyed
     * by their own date periods, while periodHours counts only periodStart/end.
     * This lets the helper respect neighboring-day transitions and week limits.
     */
    const fixedRows=allRows.filter(r=>
      clean_(r.Status).toUpperCase()!=='UNFILLED' &&
      clean_(r['Assigned Pharmacist']).toUpperCase()!=='UNFILLED' &&
      clean_(r.Username)
    );

    const state=createState_(model,fixedRows,periodStart,periodEnd);

    const slots=editable.map(r=>{
      const d=startOfDay_(asDate_(r.Date));
      const code=clean_(r.Shift).toUpperCase();
      const shift=model.shiftMap[code];
      if(!d||!shift)return null;

      const dk=formatDateKey_(d);
      return {
        originalRow:r,
        generationId:clean_(r['Generation ID'])||('FILL_'+Utilities.formatDate(new Date(),getTz_(),'yyyyMMdd_HHmmss')),
        date:d,
        dateKey:dk,
        day:dayName_(d),
        shiftCode:code,
        slot:num_(r.Slot,1),
        shift:shift,
        requiredSkill:clean_(r['Required Skill']||shift.Skill),
        weekend:yes_(r.Weekend)||isWeekendDate_(d),
        weekendGroup:clean_(r['Weekend Group'])||(isWeekendDate_(d)?weekendGroupForDate_(d,model.settings):''),
        slotKey:slotKey_(dk,code,num_(r.Slot,1))
      };
    }).filter(Boolean);

    /*
     * Scarce slots first. This prevents an easy/general open shift from using
     * the only pharmacist who could have covered a harder specialty shift.
     */
    slots.forEach(slot=>{
      slot._staticCandidates=countStaticCandidates_(slot,model);
      slot._categoryRank=shiftCategoryRank_(slot.shift);
    });

    slots.sort((a,b)=>
      a._staticCandidates-b._staticCandidates ||
      a._categoryRank-b._categoryRank ||
      num_(a.shift.Priority,50)-num_(b.shift.Priority,50) ||
      a.dateKey.localeCompare(b.dateKey) ||
      a.shiftCode.localeCompare(b.shiftCode) ||
      a.slot-b.slot
    );

    const plan=[];
    const unresolved=[];

    slots.forEach(slot=>{
      let candidates=[];

      // Stage 1: pull only pharmacists who have this shift's skill.
      // Stage 2: run the full scheduling-rule comparison only on that short list.
      skillQualifiedUsersForSlot_(slot,model).forEach(u=>{
        const e=eligibility_(u,slot,model,state,false);
        if(!e.ok)return;
        candidates.push({
          user:u,
          eligibility:e,
          coverage:false,
          score:scoreCandidate_(u,slot,model,state,e)
        });
      });

      /*
       * If no normal-skill candidate exists, use the existing generated-OFF-day
       * coverage mechanism when that exact rule applies.
       */
      if(!candidates.length){
        const owners=primaryHomeOwnersForSlot_(slot,model);
        if(owners.length===1&&isGeneratedOffDayForHomeOwner_(owners[0],slot,state,model)){
          offDaySkillQualifiedUsersForSlot_(slot,model).forEach(u=>{
            if(clean_(u.Username)===clean_(owners[0].Username))return;
            if(!hasOffDayCoverageSkillForSlot_(u,slot))return;
            const e=offDayCoverageEligibility_(u,slot,model,state);
            if(!e.ok)return;
            candidates.push({
              user:u,
              eligibility:e,
              coverage:true,
              coverageOwner:owners[0],
              score:scoreCandidate_(u,slot,model,state,e)-5
            });
          });
        }
      }

      candidates.sort((a,b)=>
        b.score-a.score ||
        clean_(a.user['Pharmacist Name']).localeCompare(clean_(b.user['Pharmacist Name']))
      );

      if(!candidates.length){
        unresolved.push({
          assignmentId:clean_(slot.originalRow['Assignment ID']),
          date:slot.dateKey,
          shift:slot.shiftCode,
          slot:slot.slot,
          reason:explainUnfilled_(slot,model,state)
        });
        return;
      }

      const pick=candidates[0];
      const warnings=(pick.eligibility.warnings||[]).slice();
      warnings.push('OPEN SHIFT HELPER: filled an existing UNFILLED slot without moving any filled assignment.');

      const assigned=makeAssigned_(slot,pick.user,warnings,ctx.username);

      /*
       * Keep the exact row identity already present in Schedule.
       */
      assigned.assignmentId=clean_(slot.originalRow['Assignment ID'])||assigned.assignmentId;
      assigned.generationId=clean_(slot.originalRow['Generation ID'])||assigned.generationId;
      assigned.holiday=clean_(slot.originalRow.Holiday);
      assigned.locked=false;
      assigned.manual=false;
      assigned.status='ASSIGNED';
      assigned.finalizedAt=null;

      if(pick.coverage&&pick.coverageOwner){
        assigned.coverageForPharmacist=clean_(pick.coverageOwner['Pharmacist Name']);
        assigned.coverageForUsername=clean_(pick.coverageOwner.Username);
        assigned.coverageReason='GENERATED OFF DAY';
        assigned.warning='OFF-DAY COVERAGE ONLY: Covering '+
          clean_(pick.coverageOwner['Pharmacist Name'])+
          ' on an algorithm-generated OFF day. | '+assigned.warning;
      }

      addAssignmentToState_(state,assigned,model);

      plan.push({
        assignment:assigned,
        date:slot.dateKey,
        shift:slot.shiftCode,
        slot:slot.slot,
        pharmacist:assigned.pharmacist,
        username:assigned.username,
        coverage:!!pick.coverage,
        alternatives:candidates.slice(1,4).map(x=>clean_(x.user['Pharmacist Name'])),
        originalWarning:clean_(slot.originalRow.Warning)
      });
    });

    if(apply&&plan.length){
      plan.forEach(item=>{
        const obj=assignmentToScheduleObject_(item.assignment);
        const ok=updateRowByKey_(
          APP.SHEETS.SCHEDULE,
          'Assignment ID',
          item.assignment.assignmentId,
          obj
        );
        if(!ok){
          throw new Error(
            'Could not update open shift '+item.date+' '+item.shift+
            ' slot '+item.slot+'. The Schedule sheet changed while the helper was running.'
          );
        }
      });

      reconcileFilledVsUnfilledScheduleRows_();
      SpreadsheetApp.flush();

      audit_(
        'OPEN_SHIFTS_AUTO_FILLED',
        formatDateKey_(start),
        '',
        '',
        '',
        '',
        String(plan.length),
        'No',
        '',
        JSON.stringify({
          start:formatDateKey_(start),
          end:formatDateKey_(end),
          filled:plan.map(x=>({
            date:x.date,
            shift:x.shift,
            slot:x.slot,
            pharmacist:x.pharmacist
          })),
          unresolved:unresolved.length,
          skippedFinalized:skippedFinalized.length,
          skippedLocked:skippedLocked.length
        }),
        ctx.username
      );
    }

    const suggestions=plan.map(x=>({
      date:x.date,
      shift:x.shift,
      slot:x.slot,
      pharmacist:x.pharmacist,
      coverage:x.coverage,
      alternatives:x.alternatives,
      previousReason:x.originalWarning
    }));

    /*
     * PERFORMANCE: the first Help Fill Open Shifts click returns only the open
     * rows that need an override. Candidate/rule analysis is loaded lazily only
     * after the administrator selects one or more rows.
     */
    let overrideTargets=[];
    if(!apply){
      const unresolvedIds=new Set(
        unresolved.map(x=>clean_(x.assignmentId)).filter(Boolean)
      );

      overrideTargets=selectedOpen
        .filter(r=>
          !!asDate_(r['Finalized At']) ||
          yes_(r.Locked) ||
          unresolvedIds.has(clean_(r['Assignment ID']))
        )
        .map(r=>({
          assignmentId:clean_(r['Assignment ID']),
          generationId:clean_(r['Generation ID']),
          date:dateKey_(r.Date),
          shift:clean_(r.Shift).toUpperCase(),
          slot:num_(r.Slot,1),
          requiredSkill:clean_(r['Required Skill']),
          originalReason:clean_(r.Warning),
          actionRequired:asDate_(r['Finalized At'])
            ? 'UNFINALIZE'
            : (yes_(r.Locked)?'UNLOCK':'OVERRIDE'),
          actionMessage:asDate_(r['Finalized At'])
            ? 'Unfinalize this schedule period before overriding this shift.'
            : (yes_(r.Locked)
                ? 'Unlock this open shift before overriding it.'
                : '')
        }));
    }

    return serialize_({
      ok:true,
      mode:apply?'APPLY':'PREVIEW',
      startDate:formatDateKey_(start),
      endDate:formatDateKey_(end),
      openCount:selectedOpen.length,
      editableOpenCount:editable.length,
      fillableCount:plan.length,
      appliedCount:apply?plan.length:0,
      unresolvedCount:unresolved.length,
      skippedFinalizedCount:skippedFinalized.length,
      skippedLockedCount:skippedLocked.length,
      suggestions:suggestions,
      unresolved:unresolved,
      overrideSuggestions:[],
      overrideTargets:overrideTargets,
      message:apply
        ? plan.length+' existing UNFILLED shift(s) were safely assigned. No existing filled assignment was moved.'
        : plan.length+' of '+editable.length+' editable UNFILLED shift(s) can currently be filled without breaking scheduling rules.'
    });
  } finally {
    lock.releaseLock();
  }
}


function validateSavedSchedule(token,startDate,endDate) {
  requireAdmin_(token);
  const model=loadSchedulingModel_();
  const repairedManualCoverage=repairManualOffDayCoverageMetadata_(model);
  const rows=readTable_(APP.SHEETS.SCHEDULE);
  let dates=rows.map(r=>asDate_(r.Date)).filter(Boolean).sort((a,b)=>a-b);
  const start=startOfDay_(asDate_(startDate)||(dates[0]||new Date()));
  const end=startOfDay_(asDate_(endDate)||(dates[dates.length-1]||start));
  const selected=rows.filter(r=>inDateRange_(asDate_(r.Date),start,end)).map(r=>scheduleRowToAssignment_(r,model)).filter(Boolean);
  const report=validateGeneratedAssignments_(selected,model,start,end);
  const health=computeScheduleHealth_(rows.filter(r=>inDateRange_(asDate_(r.Date),start,end)));
  return serialize_({ok:report.errors.length===0&&health.unfilled===0,startDate:formatDateKey_(start),endDate:formatDateKey_(end),health:health,errors:report.errors,warnings:report.warnings,repairedManualCoverage:repairedManualCoverage});
}

/** ------------------------- MANUAL SCHEDULE ----------------------- */

function resolveManualSchedulePeriod_(payload, schedule, date, model) {
  const requestedGeneration = clean_(payload && payload.generationId);
  let periodRows = [];

  if (requestedGeneration && requestedGeneration.indexOf('MANUAL_') !== 0) {
    periodRows = schedule.filter(r =>
      clean_(r['Generation ID']) === requestedGeneration &&
      asDate_(r.Date)
    );
  }

  // When a new manual assignment is added from the calendar, use the
  // generation already present on that date.
  if (!periodRows.length) {
    const sameDate = schedule.find(r =>
      dateKey_(r.Date) === formatDateKey_(date) &&
      clean_(r['Generation ID']) &&
      clean_(r['Generation ID']).indexOf('MANUAL_') !== 0
    );

    if (sameDate) {
      const generationId = clean_(sameDate['Generation ID']);
      periodRows = schedule.filter(r =>
        clean_(r['Generation ID']) === generationId &&
        asDate_(r.Date)
      );
    }
  }

  if (periodRows.length) {
    const dates = periodRows
      .map(r => startOfDay_(asDate_(r.Date)))
      .filter(Boolean)
      .sort((a,b) => a-b);

    if (dates.length) {
      return {
        start: dates[0],
        end: dates[dates.length-1]
      };
    }
  }

  // Fallback for a brand-new period with no saved generation yet.
  const defaultDays = num_(
    model.settings.raw['Default Schedule Days'],
    60
  );

  return {
    start: startOfDay_(date),
    end: addDays_(startOfDay_(date), defaultDays - 1)
  };
}

function validateManualAssignment(token,payload) {
  requireAdmin_(token);
  return serialize_(validateManualAssignment_(payload));
}

function manualAssignmentHoursSummary_(u,shift,date,schedule,ignoreId,manualPeriod,model) {
  const username=clean_(u.Username);
  const proposedHours=creditedHours_(shift);
  const weekStart=startOfDay_(asDate_(weekStartKey_(date,model.settings.weekStart)));
  const weekEnd=addDays_(weekStart,6);
  const monthStart=firstOfMonth_(date);
  const monthEnd=lastOfMonth_(date);
  const periodStart=startOfDay_(manualPeriod.start);
  const periodEnd=startOfDay_(manualPeriod.end);

  const relevant=schedule.filter(r=>{
    if (clean_(r['Assignment ID'])===clean_(ignoreId)) return false;
    if (clean_(r.Status)==='UNFILLED') return false;    if (clean_(r.Username)!==username) return false;
    return !!asDate_(r.Date);
  });

  function sumRange_(start,end) {
    return relevant.reduce((sum,r)=>{
      const d=startOfDay_(asDate_(r.Date));
      return inDateRange_(d,start,end) ? sum+num_(r['Credited Hours'],0) : sum;
    },0);
  }
  function r1_(n){return Math.round(num_(n,0)*10)/10;}

  const weekCurrent=r1_(sumRange_(weekStart,weekEnd));
  const weekDaysCurrent=relevant.filter(r=>{const d=startOfDay_(asDate_(r.Date));return inDateRange_(d,weekStart,weekEnd);}).length;
  const monthCurrent=r1_(sumRange_(monthStart,monthEnd));
  const periodCurrent=r1_(sumRange_(periodStart,periodEnd));
  const weeklyMax=num_(u['Weekly Hour Maximum'],num_(model.settings.raw['Weekly Hours Limit'],40));
  const periodTarget=num_(model.settings.periodHoursTarget,num_(model.settings.raw['Required Hours Per Schedule Period'],320));

  return {
    pharmacist:clean_(u['Pharmacist Name']),
    proposedShift:clean_(shift.Shift),
    proposedHours:r1_(proposedHours),
    week:{
      start:formatDateKey_(weekStart),end:formatDateKey_(weekEnd),
      current:weekCurrent,projected:r1_(weekCurrent+proposedHours),maximum:r1_(weeklyMax),
      currentDays:weekDaysCurrent,projectedDays:weekDaysCurrent+1,requiredDays:regularFiveDayRuleApplies_(u)?model.settings.regularWorkdaysPerWeek:null
    },
    month:{
      start:formatDateKey_(monthStart),end:formatDateKey_(monthEnd),
      current:monthCurrent,projected:r1_(monthCurrent+proposedHours)
    },
    period:{
      start:formatDateKey_(periodStart),end:formatDateKey_(periodEnd),
      current:periodCurrent,projected:r1_(periodCurrent+proposedHours),target:r1_(periodTarget)
    }
  };
}

function validateManualAssignment_(payload) {
  const model=loadSchedulingModel_();
  const date=startOfDay_(asDate_(payload.date));
  const shift=model.shiftMap[clean_(payload.shift)];
  const u=model.usersByUsername[clean_(payload.username)] || model.usersByName[clean_(payload.pharmacist)];
  if (!date||!shift||!u) return {valid:false,warnings:['Date, shift, and pharmacist are required.']};
  const schedule=readTable_(APP.SHEETS.SCHEDULE);
  const ignoreId=clean_(payload.assignmentId);
  const manualPeriod=resolveManualSchedulePeriod_(payload,schedule,date,model);
  const baseSchedule=schedule.filter(r=>clean_(r['Assignment ID'])!==ignoreId && clean_(r.Status)!=='UNFILLED');
  const state=createState_(
    model,
    baseSchedule,
    manualPeriod.start,
    manualPeriod.end
  );
  const slot={generationId:clean_(payload.generationId)||'MANUAL',date:date,dateKey:formatDateKey_(date),day:dayName_(date),shiftCode:clean_(shift.Shift),slot:num_(payload.slot,1),shift:shift,requiredSkill:clean_(shift.Skill),weekend:isWeekendDate_(date),weekendGroup:isWeekendDate_(date)?weekendGroupForDate_(date,model.settings):'',slotKey:slotKey_(formatDateKey_(date),clean_(shift.Shift),num_(payload.slot,1))};
  const normalEligibility=eligibility_(u,slot,model,state,true);
  let e=normalEligibility;
  let coverage={applies:false,coverageForPharmacist:'',coverageForUsername:'',coverageReason:'',message:''};

  // v24: when the only ordinary failure is the missing Employee Skill, check
  // whether this pharmacist is explicitly authorized in Users -> Off-Day
  // Coverage Skills and whether the primary/home pharmacist is truly OFF.
  if((normalEligibility.reasons||[]).indexOf('MISSING_SKILL')>=0 && hasOffDayCoverageSkillForSlot_(u,slot)){
    const resolved=resolveManualOffDayCoverageContext_(u,slot,model,state);
    coverage={
      applies:!!resolved.applies,
      coverageForPharmacist:clean_(resolved.coverageForPharmacist),
      coverageForUsername:clean_(resolved.coverageForUsername),
      coverageReason:clean_(resolved.coverageReason),
      message:clean_(resolved.message)
    };
    if(resolved.applies && resolved.eligibility) e=resolved.eligibility;
  }

  const warnings=(e.reasons||[]).map(reasonToWarning_).concat(e.warnings||[]);
  if(!coverage.applies && coverage.message && (normalEligibility.reasons||[]).indexOf('MISSING_SKILL')>=0){
    warnings.push(coverage.message);
  }
  const hoursSummary=manualAssignmentHoursSummary_(u,shift,date,schedule,ignoreId,manualPeriod,model);
  return {valid:warnings.length===0,warnings:warnings,reasonCodes:e.reasons||[],slot:slot,user:publicUser_(u),hoursSummary:hoursSummary,coverage:coverage};
}

function saveManualAssignment(token,payload) {
  const ctx=requireAdmin_(token);
  const lock=LockService.getScriptLock();
  if(!lock.tryLock(15000)) throw new Error('Schedule is currently being modified by another administrator.');
  try {
    const check=validateManualAssignment_(payload);
    const override=yes_(payload.override);
    const model=loadSchedulingModel_();

    /*
     * MANUAL ADMIN ASSIGNMENTS MAY OVERRIDE SCHEDULING RULES.
     *
     * Automatic schedule generation still treats the normal eligibility rules
     * as hard constraints. This exception applies only when an administrator
     * manually edits/assigns a schedule slot.
     *
     * Flow:
     *   1) validateManualAssignment_ reports every violated rule;
     *   2) first Save returns those warnings to the UI;
     *   3) the administrator must explicitly choose Override & Save;
     *   4) an override reason is optional;
     *   5) warnings are always preserved, and any optional reason is saved in
     *      the Schedule/Audit data.
     *
     * Structural problems such as an invalid date, missing pharmacist/shift,
     * or trying to edit a finalized period are still blocked later because
     * those are not scheduling-rule overrides.
     */
    if (check.warnings.length && !override) {
      return serialize_({
        ok:false,
        requiresOverride:true,
        warnings:check.warnings,
        reasonCodes:check.reasonCodes||[],
        hoursSummary:check.hoursSummary,
        message:'This manual assignment violates one or more scheduling rules. Review the warnings, then choose Override & Save if you still want this assignment. An override reason is optional.'
      });
    }

    const rows=readTable_(APP.SHEETS.SCHEDULE);
    const u=model.usersByUsername[clean_(payload.username)] || model.usersByName[clean_(payload.pharmacist)];
    const shift=model.shiftMap[clean_(payload.shift)];
    const d=startOfDay_(asDate_(payload.date));
    if(!u||!shift||!d) throw new Error('Invalid manual assignment.');

    // Prefer updating the existing required-slot row in place. Some older
    // frontends did not send Assignment ID when an UNFILLED card was opened,
    // which caused a new manual row to be appended while the old UNFILLED row
    // remained in the physical Schedule sheet. Resolve that here by finding the
    // exact Date + Shift + Slot UNFILLED row and reusing its Assignment ID.
    const targetSlot=num_(payload.slot,1);
    let id=clean_(payload.assignmentId);
    let old=id ? rows.find(r=>clean_(r['Assignment ID'])===id) : null;
    if(!old){
      old=rows.find(r=>{
        const rd=startOfDay_(asDate_(r.Date));
        return rd && formatDateKey_(rd)===formatDateKey_(d) &&
          clean_(r.Shift)===clean_(shift.Shift) &&
          num_(r.Slot,1)===targetSlot &&
          (clean_(r.Status).toUpperCase()==='UNFILLED' || clean_(r['Assigned Pharmacist']).toUpperCase()==='UNFILLED');
      }) || null;
      if(old) id=clean_(old['Assignment ID']);
    }
    if(!id) id=Utilities.getUuid();
    if(old && asDate_(old['Finalized At'])) throw new Error('This assignment is in a finalized schedule. Unfinalize the period before editing it.');
    const coverage=check.coverage||{};
    const rowWarnings=(check.warnings||[]).slice();
    if(override && rowWarnings.length){
      const optionalReason=clean_(payload.overrideReason);
      rowWarnings.unshift(optionalReason ? ('ADMIN MANUAL OVERRIDE: '+optionalReason) : 'ADMIN MANUAL OVERRIDE');
    }
    if(coverage.applies){
      rowWarnings.unshift('OFF-DAY COVERAGE ONLY: Covering '+clean_(coverage.coverageForPharmacist)+' on an algorithm-generated OFF day.');
    }
    const row={
      'Generation ID':clean_(payload.generationId)||(old?clean_(old['Generation ID']):'MANUAL_'+Utilities.formatDate(new Date(),getTz_(),'yyyyMMdd_HHmmss')),
      'Assignment ID':id,'Date':d,'Day':dayName_(d),'Shift':shift.Shift,'Slot':num_(payload.slot,old?num_(old.Slot,1):1),
      'Assigned Pharmacist':u['Pharmacist Name'],'Username':u.Username,'Hours':num_(shift.Hours,0),'Credited Hours':creditedHours_(shift),'Required Skill':clean_(shift.Skill),
      // v24: coverage metadata is recomputed from the current selected
      // pharmacist and shift. Do not preserve stale coverage metadata when a
      // normal-skilled pharmacist replaces a previous OFF-day coverer.
      'Coverage For Pharmacist':coverage.applies?clean_(coverage.coverageForPharmacist):'',
      'Coverage For Username':coverage.applies?clean_(coverage.coverageForUsername):'',
      'Coverage Reason':coverage.applies?clean_(coverage.coverageReason):'',
      'Shift Type':clean_(shift.Type),'Weekend':isWeekendDate_(d)?'Yes':'No','Weekend Group':isWeekendDate_(d)?weekendGroupForDate_(d,model.settings):'',
      'Holiday':old?old.Holiday:'','Locked':yes_(payload.locked)?'Yes':'No','Manual':'Yes','Status':yes_(payload.locked)?'LOCKED':'MANUAL',
      'Warning':rowWarnings.join(' | '),'Updated At':new Date(),'Updated By':ctx.username,'Finalized At':old?old['Finalized At']:''
    };
    // Persist the assignment to the Schedule sheet first. If this was an
    // UNFILLED row opened from the calendar, the same Assignment ID is updated
    // in place rather than creating a browser-only replacement.
    upsertScheduleRow_(row);

    // If this manual save filled a required slot that previously existed as
    // UNFILLED, remove any duplicate UNFILLED row for the exact same
    // date + shift + slot. This keeps one canonical row per required slot.
    const removedForSlot=removeDuplicateUnfilledRowsForSlot_(d, shift.Shift, num_(row.Slot,1), id);
    // Safety-net reconciliation across the sheet. If a filled row exists for a
    // Date + Shift + Slot, no stale UNFILLED copy for that same slot may remain.
    const removedGlobal=reconcileFilledVsUnfilledScheduleRows_();

    // Force pending Spreadsheet service writes to complete before telling the
    // browser the save succeeded. Then re-read the physical Schedule sheet and
    // verify that the row is actually there with the selected pharmacist.
    SpreadsheetApp.flush();
    const verify=findRowByKey_(APP.SHEETS.SCHEDULE,'Assignment ID',id);
    if(!verify || clean_(verify['Assigned Pharmacist'])!==clean_(u['Pharmacist Name']) || clean_(verify.Status).toUpperCase()==='UNFILLED') {
      throw new Error('The calendar assignment could not be verified in the Google Sheet. Run connectSchedulerToCurrentSpreadsheet() from the Apps Script editor for the Sheet you want this web app to use, then try again.');
    }
    if(coverage.applies && (
      clean_(verify['Coverage For Username'])!==clean_(coverage.coverageForUsername) ||
      clean_(verify['Coverage Reason']).toUpperCase()!=='GENERATED OFF DAY'
    )){
      throw new Error('The pharmacist assignment was saved, but the OFF-day coverage metadata was not written to the Schedule sheet. Please try the save again.');
    }

    const auditDetails=(check.warnings||[]).slice();
    if(coverage.applies) auditDetails.unshift('OFF-DAY COVERAGE: '+clean_(u['Pharmacist Name'])+' covering '+clean_(coverage.coverageForPharmacist)+' for '+clean_(shift.Shift)+'.');
    audit_('ASSIGNMENT_CHANGED',formatDateKey_(d),u['Pharmacist Name'],old?old.Shift:'',shift.Shift,old?old['Assigned Pharmacist']:'',u['Pharmacist Name'],override?'Yes':'No',clean_(payload.overrideReason),auditDetails.join(' | '),ctx.username);
    SpreadsheetApp.flush();
    return serialize_({ok:true,warnings:check.warnings,hoursSummary:check.hoursSummary,assignment:verify,database:getDatabaseInfo_(),removedStaleUnfilled:num_(removedForSlot,0)+num_(removedGlobal,0)});
  } finally {lock.releaseLock();}
}

function removeAssignment(token,assignmentId,reason) {
  const ctx=requireAdmin_(token);
  const rows=readTable_(APP.SHEETS.SCHEDULE);
  const row=rows.find(r=>clean_(r['Assignment ID'])===clean_(assignmentId));
  if(!row) throw new Error('Assignment not found.');
  if(asDate_(row['Finalized At'])) throw new Error('This assignment is in a finalized schedule. Unfinalize the period before removing it.');
  if(yes_(row.Locked)) throw new Error('Unlock the assignment before removing it.');
  const required=isRequiredSlot_(asDate_(row.Date),clean_(row.Shift),num_(row.Slot,1));
  if(required) {
    updateRowByKey_(APP.SHEETS.SCHEDULE,'Assignment ID',assignmentId,{
      'Assigned Pharmacist':'UNFILLED','Username':'','Coverage For Pharmacist':'','Coverage For Username':'','Coverage Reason':'','Manual':'Yes','Status':'UNFILLED','Warning':'MANUALLY REMOVED'+(clean_(reason)?': '+clean_(reason):''),'Updated At':new Date(),'Updated By':ctx.username
    });
    collapseDuplicateUnfilledRowsForSlot_(asDate_(row.Date),clean_(row.Shift),num_(row.Slot,1),assignmentId);
  } else deleteRowByKey_(APP.SHEETS.SCHEDULE,'Assignment ID',assignmentId);
  audit_('ASSIGNMENT_REMOVED',dateKey_(row.Date),row['Assigned Pharmacist'],row.Shift,'',row['Assigned Pharmacist'],'','No','',clean_(reason),ctx.username);
  return {ok:true};
}

function setAssignmentLock(token,assignmentId,locked) {
  const ctx=requireAdmin_(token);
  const row=findRowByKey_(APP.SHEETS.SCHEDULE,'Assignment ID',assignmentId);
  if(!row) throw new Error('Assignment not found.');
  if(asDate_(row['Finalized At'])) throw new Error('This assignment is in a finalized schedule. Unfinalize the period before changing its lock.');
  updateRowByKey_(APP.SHEETS.SCHEDULE,'Assignment ID',assignmentId,{
    'Locked':locked?'Yes':'No','Status':locked?'LOCKED':(yes_(row.Manual)?'MANUAL':'ASSIGNED'),'Updated At':new Date(),'Updated By':ctx.username
  });
  audit_(locked?'ASSIGNMENT_LOCKED':'ASSIGNMENT_UNLOCKED',dateKey_(row.Date),row['Assigned Pharmacist'],row.Shift,row.Shift,'','', 'No','', '',ctx.username);
  return {ok:true};
}

function finalizeSchedule(token,startDate,endDate,finalize) {
  const ctx=requireAdmin_(token);
  const start=startOfDay_(asDate_(startDate)),end=startOfDay_(asDate_(endDate));
  if(!start||!end) throw new Error('Start and end dates are required.');

  // Finalization is the hard gate for the complete v14 work pattern. It checks
  // exact five-day weeks, required A/B/C weekends, the five-consecutive-day
  // maximum, resident E2, skills/availability, and the regular pharmacist
  // period-hour target. True 7-on/7-off employees are validated by their block
  // pattern rather than the regular five-day/320-hour rule.
  if (finalize) {
    const model=loadSchedulingModel_();
    repairManualOffDayCoverageMetadata_(model);
    const rows=readTable_(APP.SHEETS.SCHEDULE);
    const assignments=rows
      .filter(r=>inDateRange_(asDate_(r.Date),start,end))
      .map(r=>scheduleRowToAssignment_(r,model))
      .filter(Boolean);
    const check=validateGeneratedAssignments_(assignments,model,start,end);
    if(check.errors.length){
      throw new Error('Cannot finalize. '+check.errors.slice(0,12).join('; ')+(check.errors.length>12?' ...':''));
    }
  }

  const sh=getDb_().getSheetByName(APP.SHEETS.SCHEDULE);
  const data=sh.getDataRange().getValues();
  const headers=data[0].map(clean_); const dateIdx=headers.indexOf('Date'),finIdx=headers.indexOf('Finalized At'),byIdx=headers.indexOf('Updated By'),atIdx=headers.indexOf('Updated At');
  const now=new Date(); let count=0;
  for(let i=1;i<data.length;i++) {
    const d=asDate_(data[i][dateIdx]); if(inDateRange_(d,start,end)){ data[i][finIdx]=finalize?now:''; data[i][byIdx]=ctx.username; data[i][atIdx]=now; count++; }
  }
  if(data.length>1) sh.getRange(1,1,data.length,data[0].length).setValues(data);
  audit_(finalize?'SCHEDULE_FINALIZED':'SCHEDULE_UNFINALIZED',formatDateKey_(start),'','','','','','No','','Through '+formatDateKey_(end),ctx.username);
  return {ok:true,count:count};
}

function isRequiredSlot_(date,shiftCode,slot) {
  if(!date)return false;
  const req=readTable_(APP.SHEETS.REQUIREMENTS).find(r=>clean_(r.Shift)===shiftCode&&yesDefault_(r.Active,true));
  if(!req)return false;
  return num_(req[dayName_(date)],0)>=num_(slot,1);
}

/** --------------------------- CRUD ADMIN -------------------------- */

function saveEmployee(token,data) {
  const ctx = requireAdmin_(token);
  data = data || {};
  const rows = readTable_(APP.SHEETS.USERS);
  const id = clean_(data['Employee ID']) || ('EMP-' + Utilities.getUuid().slice(0,8).toUpperCase());
  const old = rows.find(r => clean_(r['Employee ID']) === id);
  const values = {};
  APP.HEADERS.USERS.forEach(h => { if (data[h] !== undefined) values[h] = data[h]; });
  values['Employee ID'] = id;
  values['Pharmacist Name'] = clean_(data['Pharmacist Name']);
  if (!values['Pharmacist Name']) throw new Error('Pharmacist Name is required.');
  // Pharmacists do not log in. Username is only an internal scheduling key and is never shown as a credential.
  values.Username = old && clean_(old.Username) ? old.Username : ('EMPLOYEE::' + id);
  values.Role = 'Pharmacist';
  values['Temporary Password'] = '';
  values['Password Hash'] = '';
  values['Password Salt'] = '';
  values['Must Change Password'] = 'No';

  // Weekend Rotation Anchor Date is system-managed. Selecting A/B/C is enough;
  // the backend calculates the correct Saturday in the repeating 21-day cycle.
  const mergedScheduleType = clean_(values['Schedule Type'] !== undefined ? values['Schedule Type'] : (old ? old['Schedule Type'] : 'Regular'));
  const mergedWeekendGroup = clean_(values['Weekend Group'] !== undefined ? values['Weekend Group'] : (old ? old['Weekend Group'] : ''));
  if (mergedScheduleType.toLowerCase().indexOf('7') >= 0) {
    values['Weekend Rotation Anchor Date'] = '';
  } else if (mergedWeekendGroup) {
    const rawSettings = getSettingsMap_();
    const rotation = clean_(rawSettings['Weekend Rotation'] || 'A,B,C').split(',').map(s=>s.trim()).filter(Boolean);
    const weekendSettings = {
      weekendRotation: rotation.length ? rotation : ['A','B','C'],
      weekendAnchorDate: startOfDay_(asDate_(rawSettings['Weekend Anchor Date']) || nextOrSameSaturday_(new Date())),
      weekendAnchorGroup: clean_(rawSettings['Weekend Anchor Group'] || 'A')
    };
    values['Weekend Rotation Anchor Date'] = weekendAnchorForGroup_(mergedWeekendGroup,weekendSettings) || '';
  } else {
    values['Weekend Rotation Anchor Date'] = '';
  }

  values['Updated At'] = new Date();
  values['Updated By'] = ctx.username;

  if (!old) {
    APP.HEADERS.USERS.forEach(h => { if (values[h] === undefined) values[h] = ''; });
    if (!values.Active) values.Active = 'Yes';
    if (!values['Schedule Type']) values['Schedule Type'] = 'Regular';
    if (!values['Employment Type']) values['Employment Type'] = 'Regular';
    if (!values['Weekly Hour Maximum']) values['Weekly Hour Maximum'] = num_(getSettingsMap_()['Weekly Hours Limit'],40);
    if (!values['Target Weekly Hours']) values['Target Weekly Hours'] = 40;
    if (!values['Maximum Evening Shifts Per Month']) values['Maximum Evening Shifts Per Month'] = 7;
    if (!values['Weekend Eligible']) values['Weekend Eligible'] = 'Yes';
    if (!values['Evening Eligible']) values['Evening Eligible'] = 'Yes';
    if (!values['Night Eligible']) values['Night Eligible'] = 'Yes';
    appendObjectRow_(APP.SHEETS.USERS,values);
    audit_('EMPLOYEE_CREATED','',values['Pharmacist Name'],'','','',id,'No','','Pharmacist record created; no login account',ctx.username);
    return {ok:true,employeeId:id};
  }
  updateRowByKey_(APP.SHEETS.USERS,'Employee ID',id,values);
  audit_('EMPLOYEE_UPDATED','',values['Pharmacist Name']||old['Pharmacist Name'],'','','',JSON.stringify(publicUser_(old)),'No','',JSON.stringify(values),ctx.username);
  return {ok:true,employeeId:id};
}


function deleteRowsMatching_(sheetName,predicate) {
  const sh=getDb_().getSheetByName(sheetName);
  if(!sh || sh.getLastRow()<2)return 0;
  const data=sh.getDataRange().getValues();
  const headers=(data[0]||[]).map(clean_);
  const rows=[];
  for(let i=1;i<data.length;i++){
    const obj={};
    headers.forEach((h,j)=>obj[h]=data[i][j]);
    if(predicate(obj))rows.push(i+1);
  }
  rows.sort((a,b)=>b-a).forEach(r=>sh.deleteRow(r));
  return rows.length;
}

function deleteEmployeeProfile(token,employeeId) {
  const ctx=requireAdmin_(token);
  const id=clean_(employeeId);
  if(!id)throw new Error('Employee ID is required.');

  const employee=findRowByKey_(APP.SHEETS.USERS,'Employee ID',id);
  if(!employee)throw new Error('Pharmacist profile not found.');

  const username=clean_(employee.Username);
  const pharmacist=clean_(employee['Pharmacist Name']);
  const now=new Date();

  // Remove the pharmacist from any non-finalized working schedule while
  // preserving finalized schedule history.
  let assignmentsUnfilled=0;
  const scheduleSheet=getDb_().getSheetByName(APP.SHEETS.SCHEDULE);
  if(scheduleSheet && scheduleSheet.getLastRow()>=2){
    const data=scheduleSheet.getDataRange().getValues();
    const headers=(data[0]||[]).map(clean_);
    const idx={};headers.forEach((h,i)=>idx[h]=i);
    for(let i=1;i<data.length;i++){
      if(clean_(data[i][idx.Username])!==username)continue;
      if(idx['Finalized At']!==undefined && clean_(data[i][idx['Finalized At']]))continue;
      if(idx['Assigned Pharmacist']!==undefined)data[i][idx['Assigned Pharmacist']]='UNFILLED';
      if(idx.Username!==undefined)data[i][idx.Username]='';
      if(idx.Status!==undefined)data[i][idx.Status]='UNFILLED';
      if(idx.Warning!==undefined)data[i][idx.Warning]='Pharmacist profile deleted: '+pharmacist;
      if(idx['Updated At']!==undefined)data[i][idx['Updated At']]=now;
      if(idx['Updated By']!==undefined)data[i][idx['Updated By']]=ctx.username;
      assignmentsUnfilled++;
    }
    if(assignmentsUnfilled){
      scheduleSheet.getRange(2,1,data.length-1,data[0].length).setValues(data.slice(1));
    }
  }

  const skillsRemoved=deleteRowsMatching_(APP.SHEETS.SKILLS,r=>
    clean_(r['Employee ID'])===id || (username && clean_(r.Username)===username)
  );
  const availabilityRemoved=deleteRowsMatching_(APP.SHEETS.WEEKLY_AVAILABILITY,r=>
    clean_(r['Employee ID'])===id || (username && clean_(r.Username)===username)
  );

  let preceptorRowsRemoved=0;
  if(getDb_().getSheetByName('Preceptor Calendar')){
    preceptorRowsRemoved=deleteRowsMatching_('Preceptor Calendar',r=>
      clean_(r['Employee ID'])===id || (username && clean_(r.Username)===username)
    );
  }

  deleteRowByKey_(APP.SHEETS.USERS,'Employee ID',id);
  try{ if(typeof _PRECEPTOR_CALENDAR_RUNTIME_CACHE_!=='undefined') _PRECEPTOR_CALENDAR_RUNTIME_CACHE_=null; }catch(_e){}

  audit_(
    'EMPLOYEE_PROFILE_DELETED',
    '',
    pharmacist,
    '',
    '',
    username,
    id,
    'No',
    '',
    'Profile deleted. Non-finalized assignments changed to UNFILLED: '+assignmentsUnfilled+
      '; skills removed: '+skillsRemoved+
      '; weekly availability rows removed: '+availabilityRemoved+
      '; preceptor calendar rows removed: '+preceptorRowsRemoved+
      '. Finalized schedule history preserved.',
    ctx.username
  );

  return {
    ok:true,
    employeeId:id,
    pharmacist:pharmacist,
    assignmentsUnfilled:assignmentsUnfilled,
    skillsRemoved:skillsRemoved,
    availabilityRemoved:availabilityRemoved,
    preceptorRowsRemoved:preceptorRowsRemoved
  };
}

function resetEmployeePassword(token,employeeId) {
  requireAdmin_(token);
  throw new Error('Pharmacists do not have usernames or passwords. Password resets are only available for administrator accounts.');
}

function saveAdminAccount(token,data) {
  const ctx = requireAdmin_(token);
  data = data || {};
  const rows = readTable_(APP.SHEETS.ADMINS);
  const id = clean_(data['Admin ID']);
  if (!id) throw new Error('Admin slot is required. The system uses four fixed administrator slots.');
  const old = rows.find(r => clean_(r['Admin ID']) === id);
  if (!old) throw new Error('Administrator slot not found. Run setupPharmacyScheduler() to create the four administrator slots.');
  const username = clean_(data.Username).toLowerCase();
  const name = clean_(data['Admin Name']);
  if (!username || !name) throw new Error('Administrator name and username are required.');
  if (rows.some(r => clean_(r.Username).toLowerCase() === username && clean_(r['Admin ID']) !== id)) throw new Error('Administrator username already exists.');
  updateRowByKey_(APP.SHEETS.ADMINS,'Admin ID',id,{
    'Admin Name':name,
    'Username':username,
    'Active':yesDefault_(data.Active,true)?'Yes':'No',
    'Updated At':new Date(),
    'Updated By':ctx.username
  });
  audit_('ADMIN_UPDATED','',name,'','','',username,'No','',id,ctx.username);
  return {ok:true};
}

function resetAdminPassword(token,adminId) {
  const ctx = requireAdmin_(token);
  const row = findRowByKey_(APP.SHEETS.ADMINS,'Admin ID',adminId);
  if (!row) throw new Error('Administrator not found.');
  const password = generateTemporaryPassword_();
  const salt = Utilities.getUuid();
  updateRowByKey_(APP.SHEETS.ADMINS,'Admin ID',adminId,{
    'Temporary Password':password,
    'Password Salt':salt,
    'Password Hash':hashPassword_(password,salt),
    'Must Change Password':'Yes',
    'Updated At':new Date(),
    'Updated By':ctx.username
  });
  audit_('ADMIN_PASSWORD_RESET','',row['Admin Name'],'','','','','No','','Temporary password reset',ctx.username);
  return {ok:true,username:row.Username,temporaryPassword:password};
}

function saveSkill(token,data) {
  const ctx=requireAdmin_(token); data=data||{};
  const u=findUser_(data.Username,data['Pharmacist Name']); if(!u)throw new Error('Employee not found.');
  const skill=clean_(data.Skill).toUpperCase(); if(!skill)throw new Error('Skill is required.');
  const rows=readTable_(APP.SHEETS.SKILLS); const existing=rows.find(r=>clean_(r.Username)===clean_(u.Username)&&clean_(r.Skill).toUpperCase()===skill);
  if(existing) updateCompositeRow_(APP.SHEETS.SKILLS, r=>clean_(r.Username)===clean_(u.Username)&&clean_(r.Skill).toUpperCase()===skill, {'Active':'Yes','Updated At':new Date(),'Updated By':ctx.username});
  else appendObjectRow_(APP.SHEETS.SKILLS,{'Employee ID':u['Employee ID'],'Pharmacist Name':u['Pharmacist Name'],'Username':u.Username,'Skill':skill,'Active':'Yes','Updated At':new Date(),'Updated By':ctx.username});
  audit_('EMPLOYEE_SKILL_CHANGED','',u['Pharmacist Name'],'','','',skill,'No','','Skill added/activated',ctx.username); return {ok:true};
}

function removeSkill(token,username,skill) {
  const ctx=requireAdmin_(token); updateCompositeRow_(APP.SHEETS.SKILLS,r=>clean_(r.Username)===clean_(username)&&clean_(r.Skill).toUpperCase()===clean_(skill).toUpperCase(),{'Active':'No','Updated At':new Date(),'Updated By':ctx.username});
  audit_('EMPLOYEE_SKILL_CHANGED','','','','','',skill,'No','','Skill deactivated for '+username,ctx.username); return {ok:true};
}

function ensureStaffingRequirementForShift_(code,actor,shiftActive) {
  code=clean_(code).toUpperCase();
  if(!code)return {created:false};

  const existing=findRowByKey_(APP.SHEETS.REQUIREMENTS,'Shift',code);
  if(existing)return {created:false,row:existing};

  const now=new Date();
  const row={
    'Shift':code,
    'Sunday':0,
    'Monday':0,
    'Tuesday':0,
    'Wednesday':0,
    'Thursday':0,
    'Friday':0,
    'Saturday':0,
    /*
     * A new shift begins with zero required positions, so adding a Shift
     * definition cannot unexpectedly create schedule demand. The admin can
     * then set the desired staffing counts in the Staffing tab.
     */
    'Active':yesDefault_(shiftActive,true)?'Yes':'No',
    'Updated At':now,
    'Updated By':clean_(actor)||'SYSTEM'
  };

  appendObjectRow_(APP.SHEETS.REQUIREMENTS,row);
  return {created:true,row:row};
}

function ensureAllShiftStaffingRequirements_(actor) {
  const shifts=readTable_(APP.SHEETS.SHIFTS);
  let created=0;

  shifts.forEach(function(shift){
    const code=clean_(shift.Shift).toUpperCase();
    if(!code)return;
    const result=ensureStaffingRequirementForShift_(
      code,
      actor||'SYSTEM',
      shift.Active
    );
    if(result.created)created++;
  });

  return {created:created};
}

function saveShift(token,data) {
  const ctx=requireAdmin_(token);
  data=data||{};

  const code=clean_(data.Shift).toUpperCase();
  if(!code)throw new Error('Shift code is required.');

  const existing=findRowByKey_(APP.SHEETS.SHIFTS,'Shift',code);
  const values={};

  APP.HEADERS.SHIFTS.forEach(function(h){
    if(data[h]!==undefined)values[h]=data[h];
  });

  values.Shift=code;
  values['Updated At']=new Date();
  values['Updated By']=ctx.username;

  if(existing){
    updateRowByKey_(APP.SHEETS.SHIFTS,'Shift',code,values);
  }else{
    appendObjectRow_(APP.SHEETS.SHIFTS,values);
  }

  /*
   * Every Shift definition must also exist in Staffing Requirements.
   * New rows start at zero positions for every day and therefore do not
   * alter the generated schedule until an administrator sets staffing counts.
   */
  const staffing=ensureStaffingRequirementForShift_(
    code,
    ctx.username,
    values.Active
  );

  audit_(
    'SHIFT_UPDATED',
    '',
    '',
    existing?existing.Shift:'',
    code,
    existing?JSON.stringify(existing):'',
    JSON.stringify(values),
    'No',
    '',
    staffing.created
      ? 'Shift saved; matching Staffing Requirements row created with zero daily counts.'
      : 'Shift saved; Staffing Requirements row already exists.',
    ctx.username
  );

  return {
    ok:true,
    shift:code,
    staffingRequirementCreated:!!staffing.created
  };
}


function removeShiftType(token,shiftCode) {
  const ctx=requireAdmin_(token);
  const code=clean_(shiftCode).toUpperCase();
  if(!code)throw new Error('Shift code is required.');

  const shift=findRowByKey_(APP.SHEETS.SHIFTS,'Shift',code);
  if(!shift)throw new Error('Shift '+code+' was not found.');

  const requiredSkill=clean_(shift.Skill||code).toUpperCase();

  // Remove the definition and its staffing demand first so future generation
  // cannot recreate this shift.
  deleteRowByKey_(APP.SHEETS.SHIFTS,'Shift',code);
  deleteRowByKey_(APP.SHEETS.REQUIREMENTS,'Shift',code);

  // Remove non-finalized schedule rows for this shift. Finalized history stays.
  const scheduleRowsRemoved=deleteRowsMatching_(APP.SHEETS.SCHEDULE,r=>
    clean_(r.Shift).toUpperCase()===code && !clean_(r['Finalized At'])
  );

  // If no remaining shift uses this required skill, remove the now-orphaned
  // skill assignments from employees.
  const skillStillUsed=readTable_(APP.SHEETS.SHIFTS).some(s=>
    clean_(s.Skill||s.Shift).toUpperCase()===requiredSkill
  );
  let employeeSkillRowsRemoved=0;
  if(!skillStillUsed){
    employeeSkillRowsRemoved=deleteRowsMatching_(APP.SHEETS.SKILLS,r=>
      clean_(r.Skill).toUpperCase()===requiredSkill
    );
  }

  // Remove the deleted shift from pharmacist preference/coverage lists.
  const users=readTable_(APP.SHEETS.USERS);
  let usersCleaned=0;
  users.forEach(u=>{
    const id=clean_(u['Employee ID']);
    if(!id)return;

    const preferred=clean_(u['Preferred Shift Type'])
      .split(/[,;|]/)
      .map(x=>clean_(x))
      .filter(Boolean)
      .filter(x=>x.toUpperCase()!==code);

    const offDay=clean_(u['Off-Day Coverage Skills'])
      .split(/[,;|]/)
      .map(x=>clean_(x))
      .filter(Boolean)
      .filter(x=>x.toUpperCase()!==requiredSkill);

    const oldPreferred=clean_(u['Preferred Shift Type']);
    const oldOffDay=clean_(u['Off-Day Coverage Skills']);
    const newPreferred=preferred.join(', ');
    const newOffDay=offDay.join(', ');

    if(oldPreferred!==newPreferred || (!skillStillUsed && oldOffDay!==newOffDay)){
      const updates={
        'Preferred Shift Type':newPreferred,
        'Updated At':new Date(),
        'Updated By':ctx.username
      };
      if(!skillStillUsed)updates['Off-Day Coverage Skills']=newOffDay;
      updateRowByKey_(APP.SHEETS.USERS,'Employee ID',id,updates);
      usersCleaned++;
    }
  });

  audit_(
    'SHIFT_TYPE_REMOVED',
    '',
    '',
    code,
    '',
    JSON.stringify(shift),
    '',
    'No',
    '',
    'Shift definition and staffing requirement removed. Non-finalized schedule rows removed: '+
      scheduleRowsRemoved+'; orphaned employee skill rows removed: '+employeeSkillRowsRemoved+
      '; pharmacist preference rows cleaned: '+usersCleaned+
      '. Finalized schedule history preserved.',
    ctx.username
  );

  return {
    ok:true,
    shift:code,
    scheduleRowsRemoved:scheduleRowsRemoved,
    employeeSkillRowsRemoved:employeeSkillRowsRemoved,
    usersCleaned:usersCleaned
  };
}

function saveStaffingRequirement(token,data) {
  const ctx=requireAdmin_(token); data=data||{}; const code=clean_(data.Shift).toUpperCase(); if(!code)throw new Error('Shift is required.');
  if(!findRowByKey_(APP.SHEETS.SHIFTS,'Shift',code))throw new Error('Shift '+code+' does not exist in Shifts.');
  const values={}; APP.HEADERS.REQUIREMENTS.forEach(h=>{if(data[h]!==undefined)values[h]=data[h];}); values.Shift=code; values['Updated At']=new Date();values['Updated By']=ctx.username;
  if(findRowByKey_(APP.SHEETS.REQUIREMENTS,'Shift',code))updateRowByKey_(APP.SHEETS.REQUIREMENTS,'Shift',code,values); else appendObjectRow_(APP.SHEETS.REQUIREMENTS,values);
  audit_('STAFFING_REQUIREMENT_UPDATED','','','','','',code,'No','',JSON.stringify(values),ctx.username);return{ok:true};
}

function saveSetting(token,key,value) {
  const ctx=requireAdmin_(token); key=clean_(key); if(!key)throw new Error('Setting name is required.');
  const row=findRowByKey_(APP.SHEETS.SETTINGS,'Setting',key); const old=row?row.Value:'';
  if(row)updateRowByKey_(APP.SHEETS.SETTINGS,'Setting',key,{'Value':value,'Updated At':new Date(),'Updated By':ctx.username});
  else appendObjectRow_(APP.SHEETS.SETTINGS,{'Setting':key,'Value':value,'Description':'Custom setting','Updated At':new Date(),'Updated By':ctx.username});
  let anchorSync={updated:0};
  if(['Weekend Rotation','Weekend Anchor Date','Weekend Anchor Group'].includes(key)) {
    anchorSync=syncWeekendRotationAnchorDates_(getSpreadsheet_(),ctx.username);
  }
  audit_('SETTING_CHANGED','','','','',old,value,'No','',key,ctx.username);return{ok:true,weekendRotationAnchorsUpdated:anchorSync.updated};
}

function reviewRequest(token,recordId,status,comment) {
  const ctx=requireAdmin_(token);
  const row=findRowByKey_(APP.SHEETS.REQUESTS,'Record ID',recordId);
  if(!row)throw new Error('Request not found.');
  if(clean_(row.Status).toLowerCase()==='approved' && clean_(status).toLowerCase()==='pending'){
    throw new Error('An approved PTO / Regular Off request cannot be returned to Pending. Use Reject if the approval must be removed.');
  }
  updateRowByKey_(APP.SHEETS.REQUESTS,'Record ID',recordId,{
    'Status':status,
    'Comment':comment!==undefined?comment:row.Comment,
    'Reviewed By':ctx.username,
    'Reviewed At':new Date(),
    'Updated At':new Date(),
    'Updated By':ctx.username
  });
  // If one of the first two requests is rejected, the next eligible request
  // automatically moves into an auto-approved slot.
  const rebalance=reconcilePtoAutoApprovals_('SYSTEM');
  audit_('REQUEST_REVIEWED',dateKey_(row.Date||row['Start Date']),row.Pharmacist,'','','',status,'No','',recordId,ctx.username);
  return{ok:true,ptoAutoApprovalsUpdated:rebalance.updated};
}

function saveEmployeeRequest(token,data) {
  const ctx = requireAdmin_(token);
  data = data || {};
  const target = findUser_(data.Username || data['Employee ID'], data.Pharmacist);
  if (!target) throw new Error('Select a pharmacist.');

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const now = new Date();
    const id = clean_(data['Record ID']) || ('REQ-' + Utilities.formatDate(now,getTz_(),'yyyyMMddHHmmss') + '-' + Math.floor(Math.random()*900+100));
    const old = findRowByKey_(APP.SHEETS.REQUESTS,'Record ID',id);
    const recordType = clean_(data['Record Type']||'PTO').toUpperCase();
    const row = {
      'Record Type':recordType,
      'Record ID':id,
      'Pharmacist':target['Pharmacist Name'],
      'Username':target.Username,
      'Date':asDate_(data.Date)||'',
      'Start Date':asDate_(data['Start Date'])||'',
      'End Date':asDate_(data['End Date'])||'',
      'Weekend Saturday':data['Weekend Saturday']||'',
      'Weekend Sunday':data['Weekend Sunday']||'',
      'Available':data.Available||'',
      'Status':'',
      'Comment':data.Comment||'',
      'Submitted At':old && old['Submitted At'] ? old['Submitted At'] : now,
      'Reviewed By':'',
      'Reviewed At':'',
      'Updated At':now,
      'Updated By':ctx.username
    };

    const queuedTimeOff = recordType === 'PTO' || isRegularOffRecordType_(recordType);
    if (queuedTimeOff) {
      row['Record Type'] = recordType === 'PTO' ? 'PTO' : 'REGULAR OFF';
      const requestLabel = row['Record Type'] === 'PTO' ? 'PTO' : 'Regular Off';
      const dateKeys = ptoRequestDateKeys_(row);
      if (!dateKeys.length) throw new Error(requestLabel + ' requires a Single Date or a Start Date / End Date.');
      if (row['Start Date'] && row['End Date'] && formatDateKey_(row['End Date']) < formatDateKey_(row['Start Date'])) {
        throw new Error(requestLabel + ' End Date cannot be before Start Date.');
      }

      const duplicate = findDuplicatePtoRequest_(target,row,id);
      if (duplicate) {
        throw new Error('This pharmacist already has a PTO / Regular Off request that overlaps ' + duplicate.date + ' (Record ' + duplicate.recordId + ').');
      }

      // A rejected PTO / Regular Off request edited or resubmitted is a fresh
      // request and moves to the end of the first-come, first-served queue.
      if (old && clean_(old.Status).toLowerCase() === 'rejected') {
        row['Submitted At'] = now;
      }

      // Write the request first, then deterministically rebuild the combined
      // PTO + Regular Off queue. The earliest two requests on each requested
      // OFF date are auto-approved; request #3+ remains Pending until an
      // administrator explicitly approves it.
      row.Status = old && clean_(old.Status).toLowerCase() === 'approved' && clean_(old['Reviewed By']).toUpperCase() !== 'SYSTEM AUTO-APPROVAL'
        ? 'Approved'
        : 'Pending';
      if (row.Status === 'Approved') {
        row['Reviewed By'] = old['Reviewed By'];
        row['Reviewed At'] = old['Reviewed At'];
      }

      if (old) updateRowByKey_(APP.SHEETS.REQUESTS,'Record ID',id,row);
      else appendObjectRow_(APP.SHEETS.REQUESTS,row);

      const rebalance = reconcilePtoAutoApprovals_('SYSTEM');
      const saved = findRowByKey_(APP.SHEETS.REQUESTS,'Record ID',id) || row;
      const queueState = getPtoQueueStateForRecord_(id);
      const autoApproved = clean_(saved.Status).toLowerCase()==='approved' && clean_(saved['Reviewed By']).toUpperCase()==='SYSTEM AUTO-APPROVAL';
      const blockedDates = queueState ? queueState.blockedDates : [];
      const detail = autoApproved
        ? ('AUTO APPROVED: this ' + requestLabel + ' request is within the first ' + ptoAutoApprovalLimit_() + ' combined PTO / Regular Off request(s) on every requested OFF date. No administrator action is required.')
        : clean_(saved.Status).toLowerCase()==='approved'
          ? ('APPROVED by administrator: this ' + requestLabel + ' approval is protected and will be enforced by the scheduler.')
          : ('PENDING ADMIN REVIEW: this ' + requestLabel + ' request is number 3 or later on ' + (blockedDates.length?blockedDates.join(', '):'one or more covered dates') + '.');
      audit_(old?'REQUEST_UPDATED':'REQUEST_ENTERED',dateKeys[0],row.Pharmacist,'','','',row['Record Type'],'No','',id+' | '+detail,ctx.username);
      SpreadsheetApp.flush();
      return {ok:true,recordId:id,status:saved.Status,autoApproved:autoApproved,blockedDates:blockedDates,message:detail,ptoAutoApprovalsUpdated:rebalance.updated};
    }

    // Non-PTO date rules keep the administrator-selected status behavior.
    row.Status = clean_(data.Status || (old ? old.Status : 'Approved')) || 'Approved';
    row['Reviewed By'] = ctx.username;
    row['Reviewed At'] = now;
    if (old) updateRowByKey_(APP.SHEETS.REQUESTS,'Record ID',id,row);
    else appendObjectRow_(APP.SHEETS.REQUESTS,row);
    audit_(old?'REQUEST_UPDATED':'REQUEST_ENTERED',dateKey_(row.Date||row['Start Date']),row.Pharmacist,'','','',row['Record Type'],'No','',id,ctx.username);
    SpreadsheetApp.flush();
    return {ok:true,recordId:id,status:row.Status};
  } finally {
    lock.releaseLock();
  }
}

function ptoAutoApprovalLimit_() {
  const settings = getSettingsMap_();
  return Math.max(0, Math.floor(num_(settings['PTO Auto-Approve Per Day'],2)));
}

function ptoRequestDateKeys_(row) {
  // v29 IMPORTANT: Start Date / End Date are the requested PTO days.
  // The legacy Date column may contain the date the request was entered, so
  // it MUST NOT override an actual PTO range. Only use Date when no range
  // fields are present.
  let start = asDate_(row['Start Date']);
  let end = asDate_(row['End Date']);
  if (start || end) {
    if (start && !end) end = start;
    if (!start && end) start = end;
    if (!start || !end) return [];
    if (formatDateKey_(end) < formatDateKey_(start)) return [];
    const out = [];
    for (let d = startOfDay_(start); formatDateKey_(d) <= formatDateKey_(end); d = addDays_(d,1)) {
      out.push(formatDateKey_(d));
      if (out.length > 370) throw new Error('Time-off request is too long.');
    }
    return out;
  }

  const single = asDate_(row.Date);
  return single ? [formatDateKey_(single)] : [];
}

function findDuplicatePtoRequest_(target,row,excludeRecordId) {
  const requested = new Set(ptoRequestDateKeys_(row));
  if (!requested.size) return null;
  const username = clean_(target.Username);
  const pharmacist = clean_(target['Pharmacist Name']);
  const rows = readTable_(APP.SHEETS.REQUESTS);
  for (const r of rows) {
    if (clean_(r['Record ID']) === clean_(excludeRecordId)) continue;
    if (clean_(r['Record Type']).toUpperCase() !== 'PTO' && !isRegularOffRecordType_(r['Record Type'])) continue;
    if (clean_(r.Status).toLowerCase() === 'rejected') continue;
    if (clean_(r.Username) !== username && clean_(r.Pharmacist) !== pharmacist) continue;
    const overlap = ptoRequestDateKeys_(r).find(k=>requested.has(k));
    if (overlap) return {recordId:clean_(r['Record ID']),date:overlap};
  }
  return null;
}

function ptoQueueSubmittedMs_(row,index) {
  const submitted=asDate_(row['Submitted At'])||asDate_(row['Updated At']);
  return submitted ? submitted.getTime() : (index||0);
}

/**
 * Build deterministic first-come, first-served positions for all active PTO
 * and Regular Off requests in one shared per-date queue. Rejected requests do
 * not consume a slot. A multi-day request is
 * automatically approved only when its position is 1 or 2 (or configured
 * limit) on EVERY date that it covers.
 */
function buildPtoQueueState_(rows,limit) {
  limit = Math.max(0, Math.floor(num_(limit,2)));

  // IMPORTANT v28 RULE:
  // Every requested OFF date has its OWN independent first-come queue.
  // Example:
  //   9/10 -> first 2 requests auto-approved, 3rd+ pending
  //   9/11 -> first 2 requests auto-approved, 3rd+ pending
  // Requests for 9/10 do NOT consume either auto-approval slot for 9/11.
  const items = (rows || []).map(function(r,i) {
    return {
      row:r,
      index:i,
      id:clean_(r['Record ID']),
      submitted:ptoQueueSubmittedMs_(r,i),
      dates:ptoRequestDateKeys_(r)
    };
  }).filter(function(x) {
    return x.id &&
      (clean_(x.row['Record Type']).toUpperCase()==='PTO' || isRegularOffRecordType_(x.row['Record Type'])) &&
      clean_(x.row.Status).toLowerCase()!=='rejected' &&
      x.dates.length>0;
  });

  // Build a completely separate queue for each requested calendar day.
  const byDate = {};
  items.forEach(function(item) {
    item.dates.forEach(function(dateKey) {
      if (!byDate[dateKey]) byDate[dateKey] = [];
      byDate[dateKey].push(item);
    });
  });

  // rankByRecord[id][date] = 1,2,3... for THAT DATE only.
  const rankByRecord = {};
  const autoWinnerByRecord = {};
  items.forEach(function(item) {
    rankByRecord[item.id] = {};
    autoWinnerByRecord[item.id] = {};
  });

  Object.keys(byDate).forEach(function(dateKey) {
    const queue = byDate[dateKey].slice().sort(function(a,b) {
      return a.submitted-b.submitted || a.index-b.index || a.id.localeCompare(b.id);
    });
    queue.forEach(function(item,idx) {
      const position = idx + 1;
      rankByRecord[item.id][dateKey] = position;
      autoWinnerByRecord[item.id][dateKey] = position <= limit;
    });
  });

  const state = {};
  items.forEach(function(item) {
    const positions = rankByRecord[item.id] || {};
    const blockedDates = item.dates.filter(function(dateKey) {
      return !(autoWinnerByRecord[item.id] && autoWinnerByRecord[item.id][dateKey]);
    });
    state[item.id] = {
      // A multi-day PTO request can be auto-approved only when it is within
      // the first two requests on EVERY requested OFF date.
      autoEligible:item.dates.length>0 && blockedDates.length===0,
      blockedDates:blockedDates,
      positions:positions,
      dates:item.dates.slice(),
      limit:limit
    };
  });

  return state;
}

function getPtoQueueStateForRecord_(recordId) {
  const rows=readTable_(APP.SHEETS.REQUESTS);
  const state=buildPtoQueueState_(rows,ptoAutoApprovalLimit_());
  return state[clean_(recordId)]||null;
}

/**
 * Synchronize PTO + Regular Off statuses independently for each requested OFF
 * date. The first two combined requests FOR EACH DATE are SYSTEM AUTO-APPROVED.
 * Request 3+ remains Pending unless an administrator has explicitly approved it. Approved requests are sticky:
 * reconciliation can promote Pending requests but can never demote an Approved
 * request. If an earlier request is rejected, the next Pending request may be
 * promoted automatically.
 */
function reconcilePtoAutoApprovals_(updatedBy) {
  const rows=readTable_(APP.SHEETS.REQUESTS);
  const limit=ptoAutoApprovalLimit_();
  const state=buildPtoQueueState_(rows,limit);
  const now=new Date();
  let updated=0, promoted=0, demoted=0;

  rows.forEach(r=>{
    if(clean_(r['Record Type']).toUpperCase()!=='PTO' && !isRegularOffRecordType_(r['Record Type']))return;
    const id=clean_(r['Record ID']);
    if(!id)return;
    const status=clean_(r.Status).toLowerCase();
    if(status==='rejected')return;
    const q=state[id];
    if(!q)return;
    const reviewer=clean_(r['Reviewed By']).toUpperCase();
    const systemAuto=reviewer==='SYSTEM AUTO-APPROVAL';

    if(q.autoEligible) {
      if(status!=='approved') {
        updateRowByKey_(APP.SHEETS.REQUESTS,'Record ID',id,{
          'Status':'Approved',
          'Reviewed By':'SYSTEM AUTO-APPROVAL',
          'Reviewed At':now,
          'Updated At':now,
          'Updated By':updatedBy||'SYSTEM'
        });
        updated++; promoted++;
      }
      // Manually approved first/second requests remain manually approved; no
      // need to rewrite their reviewer metadata.
      return;
    }

    // Approval is permanent unless an administrator explicitly rejects or
    // deletes the request. Queue recalculation may promote Pending requests,
    // but it must never demote an already Approved request.
    if(status!=='approved' && status!=='pending') {
      updateRowByKey_(APP.SHEETS.REQUESTS,'Record ID',id,{
        'Status':'Pending',
        'Reviewed By':'',
        'Reviewed At':'',
        'Updated At':now,
        'Updated By':updatedBy||'SYSTEM'
      });
      updated++;
    }
  });

  if(updated) SpreadsheetApp.flush();
  return {ok:true,updated:updated,promoted:promoted,demoted:demoted,limit:limit};
}

/** Public maintenance function; not needed for normal use because v27 runs
 * reconciliation automatically on load and whenever a PTO request changes. */
function repairPtoAutoApprovals() {
  const result=reconcilePtoAutoApprovals_('SYSTEM');
  Logger.log(JSON.stringify(result));
  return result;
}

/** v28 explicit maintenance alias: rebuild first-two AUTO approvals per requested OFF date. */
function rebuildPtoFirstTwoPerRequestedDay() {
  const result = reconcilePtoAutoApprovals_('SYSTEM');
  Logger.log(JSON.stringify(result));
  return result;
}

/** v29 maintenance helper. Rebuilds approval queues using Start Date / End Date
 * as the requested PTO dates. Safe to run on existing data. */
function rebuildPtoFirstTwoPerActualRequestedDates() {
  const result = reconcilePtoAutoApprovals_('SYSTEM');
  Logger.log(JSON.stringify(result));
  return result;
}

function evaluatePtoAutoApproval_(row,excludeRecordId) {
  // Compatibility helper for any older code paths. Evaluate the proposed row
  // by inserting it into the current queue with its Submitted At timestamp.
  const rows=readTable_(APP.SHEETS.REQUESTS).filter(r=>clean_(r['Record ID'])!==clean_(excludeRecordId));
  const proposed=Object.assign({},row);
  proposed['Record ID']=clean_(excludeRecordId)||('PROPOSED-'+Utilities.getUuid());
  proposed['Record Type']='PTO';
  proposed.Status='Pending';
  if(!proposed['Submitted At']) proposed['Submitted At']=new Date();
  rows.push(proposed);
  const state=buildPtoQueueState_(rows,ptoAutoApprovalLimit_())[proposed['Record ID']];
  return {
    status:state&&state.autoEligible?'Approved':'Pending',
    blockedDates:state?state.blockedDates:ptoRequestDateKeys_(proposed),
    positions:state?state.positions:{},
    limit:ptoAutoApprovalLimit_()
  };
}

function submitRequest(token,data) {
  // Compatibility wrapper: all current logins are administrators.
  return saveEmployeeRequest(token,data);
}

function saveWeeklyAvailability(token,data) {
  const ctx=requireAdmin_(token); data=data||{};
  let targetUser=findUser_(data.Username||data['Employee ID'],data['Pharmacist Name']);
  if(!targetUser) throw new Error('Employee not found.');

  const day=clean_(data.Day);
  if(!dayNames_().includes(day)) throw new Error('Day must be Sunday through Saturday.');
  const available=yesDefault_(data.Available,true)?'Yes':'No';
  const rule=clean_(data['Rule Type']||'HARD').toUpperCase();
  if(!['HARD','SOFT'].includes(rule)) throw new Error('Rule Type must be HARD or SOFT.');

  let start=normalizeTimeString_(data['Start Time']);
  let end=normalizeTimeString_(data['End Time']);
  if(available==='No'){start='';end='';}
  if((start&&!end)||(!start&&end)) throw new Error('Enter both Start Time and End Time, or leave both blank for all-day availability.');

  const effectiveStart=asDate_(data['Effective Start']);
  const effectiveEnd=asDate_(data['Effective End']);
  if(effectiveStart&&effectiveEnd&&formatDateKey_(effectiveEnd)<formatDateKey_(effectiveStart)) throw new Error('Effective End cannot be before Effective Start.');

  const id=clean_(data['Availability ID'])||('WA-'+Utilities.getUuid().slice(0,8).toUpperCase());
  const existing=findRowByKey_(APP.SHEETS.WEEKLY_AVAILABILITY,'Availability ID',id);

  const row={
    'Availability ID':id,
    'Employee ID':targetUser['Employee ID'],
    'Pharmacist Name':targetUser['Pharmacist Name'],
    'Username':targetUser.Username,
    'Day':day,
    'Available':available,
    'Start Time':start,
    'End Time':end,
    'Rule Type':rule,
    'Effective Start':effectiveStart||'',
    'Effective End':effectiveEnd||'',
    'Notes':clean_(data.Notes),
    'Active':yesDefault_(data.Active,true)?'Yes':'No',
    'Updated At':new Date(),
    'Updated By':ctx.username
  };
  if(existing) updateRowByKey_(APP.SHEETS.WEEKLY_AVAILABILITY,'Availability ID',id,row);
  else appendObjectRow_(APP.SHEETS.WEEKLY_AVAILABILITY,row);
  audit_('WEEKLY_AVAILABILITY_CHANGED','',targetUser['Pharmacist Name'],'','',existing?JSON.stringify(existing):'',JSON.stringify(row),'No','',day+' recurring availability',ctx.username);
  return serialize_({ok:true,row:row});
}

function saveWeeklyAvailabilityPattern(token,employeeId,pattern) {
  const ctx=requireAdmin_(token);
  const target=findUser_(employeeId,'');
  if(!target) throw new Error('Employee not found.');
  pattern=Array.isArray(pattern)?pattern:[];
  const lock=LockService.getScriptLock();
  if(!lock.tryLock(15000)) throw new Error('Weekly availability is currently being updated. Try again.');
  try {
    const existing=readTable_(APP.SHEETS.WEEKLY_AVAILABILITY).filter(r=>{const sameEmployee=(clean_(r['Employee ID'])&&clean_(r['Employee ID'])===clean_(target['Employee ID']))||(!clean_(r['Employee ID'])&&clean_(r.Username)===clean_(target.Username));return sameEmployee&&yesDefault_(r.Active,true)&&!asDate_(r['Effective Start'])&&!asDate_(r['Effective End'])&&clean_(r.Notes).indexOf('[WEEKLY PATTERN]')===0;});
    existing.forEach(r=>updateRowByKey_(APP.SHEETS.WEEKLY_AVAILABILITY,'Availability ID',r['Availability ID'],{'Active':'No','Updated At':new Date(),'Updated By':ctx.username}));
    pattern.forEach(item=>{
      const day=clean_(item.Day);
      if(!dayNames_().includes(day)) return;
      const available=yesDefault_(item.Available,true)?'Yes':'No';
      let st=normalizeTimeString_(item['Start Time']), en=normalizeTimeString_(item['End Time']);
      if(available==='No'){st='';en='';}
      if((st&&!en)||(!st&&en)) throw new Error(day+': enter both Start Time and End Time.');
      appendObjectRow_(APP.SHEETS.WEEKLY_AVAILABILITY,{
        'Availability ID':'WA-'+Utilities.getUuid().slice(0,8).toUpperCase(),
        'Employee ID':target['Employee ID'],'Pharmacist Name':target['Pharmacist Name'],'Username':target.Username,
        'Day':day,'Available':available,'Start Time':st,'End Time':en,'Rule Type':clean_(item['Rule Type']||'HARD').toUpperCase()==='SOFT'?'SOFT':'HARD',
        'Effective Start':'','Effective End':'','Notes':'[WEEKLY PATTERN] '+clean_(item.Notes),'Active':'Yes','Updated At':new Date(),'Updated By':ctx.username
      });
    });
    audit_('WEEKLY_PATTERN_SAVED','',target['Pharmacist Name'],'','','','', 'No','',JSON.stringify(pattern),ctx.username);
    return {ok:true};
  } finally { lock.releaseLock(); }
}

function deleteWeeklyAvailability(token,availabilityId) {
  const ctx=requireAdmin_(token);
  const row=findRowByKey_(APP.SHEETS.WEEKLY_AVAILABILITY,'Availability ID',availabilityId);
  if(!row) throw new Error('Weekly availability rule not found.');
  updateRowByKey_(APP.SHEETS.WEEKLY_AVAILABILITY,'Availability ID',availabilityId,{'Active':'No','Updated At':new Date(),'Updated By':ctx.username});
  audit_('WEEKLY_AVAILABILITY_REMOVED','',row['Pharmacist Name'],'','','',availabilityId,'No','','Rule deactivated',ctx.username);
  return {ok:true};
}

function getAuditLog(token,limit) {
  requireAdmin_(token); const rows=readTable_(APP.SHEETS.AUDIT); return serialize_(rows.slice(-Math.min(num_(limit,200),1000)).reverse());
}

/** --------------------------- RULE HELPERS ------------------------ */

function isBlockedByPto_(u,date,model) {
  const username=clean_(u.Username), dk=formatDateKey_(date), key=username+'|'+dk;
  if(model._ptoCache && model._ptoCache[key]!==undefined)return model._ptoCache[key];
  const rows=(model.requestsByUser&&model.requestsByUser[username]) || model.requests || [];
  const blocked=rows.some(r => {
    if (clean_(r.Status).toLowerCase()!=='approved') return false;
    if (clean_(r['Record Type']).toUpperCase()!=='PTO') return false;
    if ((!model.requestsByUser) && !requestMatchesUser_(r,u)) return false;
    return requestCoversDate_(r,date);
  });
  if(model._ptoCache)model._ptoCache[key]=blocked;
  return blocked;
}

function isRegularOffRecordType_(value) {
  const type=clean_(value).toUpperCase().replace(/[\-_]+/g,' ').replace(/\s+/g,' ').trim();
  return type==='REGULAR OFF' || type==='REGULAR OFF REQUEST' || type==='REGULAROFF';
}

function isBlockedByRegularOff_(u,date,model) {
  const username=clean_(u.Username), dk=formatDateKey_(date), key=username+'|'+dk;
  if(model._regularOffCache && model._regularOffCache[key]!==undefined)return model._regularOffCache[key];
  const rows=(model.requestsByUser&&model.requestsByUser[username]) || model.requests || [];
  const blocked=rows.some(r => {
    if(clean_(r.Status).toLowerCase()!=='approved')return false;
    if(!isRegularOffRecordType_(r['Record Type']))return false;
    if((!model.requestsByUser) && !requestMatchesUser_(r,u))return false;
    return requestCoversDate_(r,date);
  });
  if(model._regularOffCache)model._regularOffCache[key]=blocked;
  return blocked;
}

function availabilityStatus_(u,date,model) {
  const username=clean_(u.Username), dk=formatDateKey_(date), key=username+'|'+dk;
  if(model._availabilityCache && model._availabilityCache[key])return model._availabilityCache[key];
  let blocked=false,preferred=false;
  const rows=(model.requestsByUser&&model.requestsByUser[username]) || model.requests || [];
  rows.forEach(r=>{
    if(clean_(r.Status).toLowerCase()!=='approved')return;
    if((!model.requestsByUser) && !requestMatchesUser_(r,u))return;
    const type=clean_(r['Record Type']).toUpperCase();
    if(type==='PTO'||isRegularOffRecordType_(type)||!requestCoversDate_(r,date))return;
    if(type==='UNAVAILABLE'||clean_(r.Available).toLowerCase()==='no')blocked=true;
    if(type==='AVAILABILITY'||type==='PREFERRED'||clean_(r.Available).toLowerCase()==='yes')preferred=true;
  });
  const result={blocked,preferred};
  if(model._availabilityCache)model._availabilityCache[key]=result;
  return result;
}

function weeklyAvailabilityStatus_(u,date,shift,model) {
  const username=clean_(u.Username);
  const dk=formatDateKey_(date);
  const shiftCode=clean_((shift||{}).Shift)||clean_((shift||{}).Meaning)||clean_((shift||{}).Type);
  const cacheKey=username+'|'+dk+'|'+shiftCode;
  if(model._weeklyAvailabilityCache && model._weeklyAvailabilityCache[cacheKey])return model._weeklyAvailabilityCache[cacheKey];
  const rows=(model.weeklyAvailabilityByUser&&model.weeklyAvailabilityByUser[username]?model.weeklyAvailabilityByUser[username]:[]).filter(r=>{
    if(!yesDefault_(r.Active,true)) return false;
    if(clean_(r.Day)!==dayName_(date)) return false;
    const es=startOfDay_(asDate_(r['Effective Start'])), ee=startOfDay_(asDate_(r['Effective End'])), d=startOfDay_(date);
    if(es&&formatDateKey_(d)<formatDateKey_(es)) return false;
    if(ee&&formatDateKey_(d)>formatDateKey_(ee)) return false;
    return true;
  });
  let result;
  if(!rows.length) result={blocked:false,reason:'',preferred:false,softMismatch:false,matched:false};
  else {
    const hard=rows.filter(r=>clean_(r['Rule Type']||'HARD').toUpperCase()==='HARD');
    const soft=rows.filter(r=>clean_(r['Rule Type']||'HARD').toUpperCase()==='SOFT');    if(hard.some(r=>!yes_(r.Available))) result={blocked:true,reason:'WEEKLY_AVAILABILITY_DAY',preferred:false,softMismatch:false,matched:true};
    else {
      const hardYes=hard.filter(r=>yes_(r.Available));
      if(hardYes.length && !hardYes.some(r=>weeklyRuleFitsShift_(r,shift))) result={blocked:true,reason:'WEEKLY_AVAILABILITY_TIME',preferred:false,softMismatch:false,matched:true};
      else {
        let preferred=false, softMismatch=false;
        const softYes=soft.filter(r=>yes_(r.Available));
        if(soft.some(r=>!yes_(r.Available))) softMismatch=true;
        if(softYes.length){
          preferred=softYes.some(r=>weeklyRuleFitsShift_(r,shift));
          if(!preferred) softMismatch=true;
        }
        if(hardYes.length && hardYes.some(r=>weeklyRuleFitsShift_(r,shift))) preferred=true;
        result={blocked:false,reason:'',preferred,softMismatch,matched:true};
      }
    }
  }
  if(model._weeklyAvailabilityCache)model._weeklyAvailabilityCache[cacheKey]=result;
  return result;
}

function weeklyRuleFitsShift_(rule,shift) {
  const rs=timeMinutes_(rule['Start Time']), re=timeMinutes_(rule['End Time']);
  if(rs===null&&re===null) return true;
  const ss=timeMinutes_(shift.Start), se=timeMinutes_(shift.End);
  if(ss===null||se===null||rs===null||re===null) return false;
  return intervalContains_(rs,re,ss,se);
}

function requestMatchesUser_(r,u){
  return clean_(r.Username).toLowerCase()===clean_(u.Username).toLowerCase() ||
    clean_(r.Pharmacist).toLowerCase()===clean_(u['Pharmacist Name']).toLowerCase();
}
function requestCoversDate_(r,date){
  const d=startOfDay_(date);
  const s=startOfDay_(asDate_(r['Start Date']));
  const e=startOfDay_(asDate_(r['End Date']));

  // IMPORTANT: when Start/End exist, they are the requested OFF dates.
  // Some legacy Google rows use Date as the date the request was entered.
  // Do not accidentally block both the submission date and the actual request.
  if(s||e){
    const start=s||e;
    const end=e||s;
    return !!start&&!!end&&d>=start&&d<=end;
  }

  const single=startOfDay_(asDate_(r.Date));
  return !!single&&sameDate_(single,d);
}

function residentEligibilityReason_(u,shift,date,model) {
  if (!yes_(u.Resident)) return '';
  const type=clean_(shift.Type).toLowerCase(),code=clean_(shift.Shift).toUpperCase();
  if(code==='RES'||code==='TDP')return '';
  // E2 is a required resident duty: exactly once per week.
  if (code===(model ? model.settings.residentE2ShiftCode : 'E2')) return '';
  // Weekend work is controlled by the resident's A/B/C group, not the old
  // Resident Weekends toggle. Group mismatch is handled in eligibility_.
  if (isWeekendDate_(date)) {
    if(type==='night'&&!yesDefault_(u['Resident Nights'],false))return 'RESIDENT_RESTRICTION';
    return '';
  }
  if(type==='evening'&&!yesDefault_(u['Resident Evenings'],false))return 'RESIDENT_RESTRICTION';
  if(type==='night'&&!yesDefault_(u['Resident Nights'],false))return 'RESIDENT_RESTRICTION';
  if(!yesDefault_(u['Resident Covers Regular'],false))return 'RESIDENT_RESTRICTION';
  return '';
}

/**
 * Return ONLY active pharmacists who possess the required Employee Skill for
 * this slot. This is the first stage of Help Fill Open Shifts.
 *
 * E2 keeps the existing resident special rule: active residents are treated as
 * qualified for the required resident E2 duty even when a separate skill row
 * has not been added.
 */
function skillQualifiedUsersForSlot_(slot,model) {
  const active=(model&&model.activeUsers)||[];
  if(!slot)return [];

  const required=clean_(slot.requiredSkill || (slot.shift||{}).Skill).toUpperCase();
  const shiftCode=clean_(slot.shiftCode || (slot.shift||{}).Shift).toUpperCase();

  if(!required){
    return active.slice();
  }

  const out=[];
  const seen=new Set();

  ((model.activeUsersBySkill||{})[required]||[]).forEach(u=>{
    const username=clean_(u.Username);
    if(!username||seen.has(username))return;
    seen.add(username);
    out.push(u);
  });

  if(
    model.settings &&
    shiftCode===clean_(model.settings.residentE2ShiftCode).toUpperCase()
  ){
    (model.residents||active.filter(u=>yes_(u.Resident))).forEach(u=>{
      const username=clean_(u.Username);
      if(!username||seen.has(username))return;
      seen.add(username);
      out.push(u);
    });
  }

  return out;
}

/**
 * Dedicated Users -> Off-Day Coverage Skills lookup. This is only used by the
 * existing generated-OFF-day coverage path; it never turns an ordinary
 * pharmacist without the skill into a normal candidate.
 */
function offDaySkillQualifiedUsersForSlot_(slot,model) {
  if(!slot||!model)return [];
  const keys=[
    clean_(slot.shiftCode || (slot.shift||{}).Shift).toUpperCase(),
    clean_(slot.requiredSkill || (slot.shift||{}).Skill).toUpperCase()
  ].filter(Boolean);

  const out=[];
  const seen=new Set();

  keys.forEach(key=>{
    ((model.activeUsersByOffDaySkill||{})[key]||[]).forEach(u=>{
      const username=clean_(u.Username);
      if(!username||seen.has(username))return;
      seen.add(username);
      out.push(u);
    });
  });

  return out;
}

function hasRequiredSkillForSlot_(u,slot,model) {
  if (!slot.requiredSkill) return true;
  // E2 is part of the resident staffing pattern, so residents are considered
  // qualified for E2 even when the separate skill row has not yet been added.
  if (isResidentMandatoryE2_(u,slot,model)) return true;

  const skills=model.skillsByUser[clean_(u.Username)];
  if(!skills||!skills.size)return false;

  const required=clean_(slot.requiredSkill).toUpperCase();

  // Fast path for standardized skill codes.
  if(skills.has(slot.requiredSkill)||skills.has(required))return true;

  // Compatibility path for older rows with mixed capitalization.
  for(const skill of skills){
    if(clean_(skill).toUpperCase()===required)return true;
  }

  return false;
}

function isResidentMandatoryE2_(u,slot,model) {
  return !!(u && yes_(u.Resident) && slot && clean_(slot.shiftCode || (slot.shift||{}).Shift).toUpperCase()===model.settings.residentE2ShiftCode);
}

function buildSevenOnShiftMap_(model) {
  const map={};
  const seven=model.users.filter(u=>yes_(u.Active)&&isSevenOn_(u));
  const unresolved=[];

  seven.forEach(u=>{
    const username=clean_(u.Username);
    const pref=clean_(u['Preferred Shift Type']).toUpperCase();
    const skills=model.skillsByUser[username]||new Set();

    // An exact shift code in Preferred Shift Type is the strongest instruction.
    if (pref && model.shiftMap[pref] && yes_(model.shiftMap[pref].Active)) {
      map[username]=pref;
      return;
    }

    const special=['N1','N2','E'].filter(code=>model.shiftMap[code]&&yes_(model.shiftMap[code].Active)&&skills.has(code));
    if (special.length===1) {
      map[username]=special[0];
      return;
    }

    if ((pref==='EVENING'||pref==='E') && model.shiftMap.E) {
      map[username]='E';
      return;
    }

    unresolved.push(u);
  });

  // For ambiguous night employees who can work both N1 and N2, pair them by
  // 7-on/7-off phase. Each phase gets one N1 and one N2 whenever possible.
  const night=unresolved.filter(u=>{
    const skills=model.skillsByUser[clean_(u.Username)]||new Set();
    const pref=clean_(u['Preferred Shift Type']).toUpperCase();
    return pref==='NIGHT'||skills.has('N1')||skills.has('N2');
  });

  const anchors=night.map(u=>asDate_(u['Rotation Anchor Date'])).filter(Boolean).sort((a,b)=>a-b);
  const base=anchors.length ? startOfDay_(anchors[0]) : null;
  const phases={0:[],1:[]};
  night.forEach(u=>{
    const a=asDate_(u['Rotation Anchor Date']);
    let phase=0;
    if (base&&a) {
      const mod=((daysBetween_(base,startOfDay_(a))%14)+14)%14;
      phase=mod<7?0:1;
    }
    phases[phase].push(u);
  });

  [0,1].forEach(phase=>{
    const arr=phases[phase].sort((a,b)=>clean_(a['Pharmacist Name']).localeCompare(clean_(b['Pharmacist Name'])));
    const counts={N1:0,N2:0};
    arr.forEach(u=>{
      const username=clean_(u.Username);
      if(map[username])return;
      const skills=model.skillsByUser[username]||new Set();
      const choices=['N1','N2'].filter(c=>model.shiftMap[c]&&yes_(model.shiftMap[c].Active)&&(skills.has(c)||clean_(u['Preferred Shift Type']).toUpperCase()==='NIGHT'));
      if(!choices.length)return;
      choices.sort((a,b)=>counts[a]-counts[b]||a.localeCompare(b));
      map[username]=choices[0];
      counts[choices[0]]++;
    });
  });

  // Remaining 7-on employees with E skill are dedicated E employees.
  unresolved.forEach(u=>{
    const username=clean_(u.Username);
    if(map[username])return;
    const skills=model.skillsByUser[username]||new Set();
    if(model.shiftMap.E && yes_(model.shiftMap.E.Active) && skills.has('E')) map[username]='E';
  });

  return map;
}

function sevenOnAssignedShiftCode_(u,model) {
  return clean_((model.sevenOnShiftByUser||{})[clean_(u.Username)]).toUpperCase();
}

function sevenOnSlotMatches_(u,slot,model) {
  const code=sevenOnAssignedShiftCode_(u,model);
  return !!code && code===clean_(slot.shiftCode || (slot.shift||{}).Shift).toUpperCase();
}


function regularFiveDayRuleApplies_(u){
  return !!u && !isSevenOn_(u) && !isPrnEmployee_(u);
}

function regularProtectedOffDaysForWeek_(u,weekStart,model){
  if(!u||!weekStart||!model)return 0;
  const seen=new Set();

  for(let i=0;i<7;i++){
    const d=addDays_(weekStart,i);
    if(isBlockedByPto_(u,d,model)||isBlockedByRegularOff_(u,d,model)){
      seen.add(formatDateKey_(d));
    }
  }

  return seen.size;
}

function regularRequiredWorkdaysForWeek_(u,date,model){
  if(!regularFiveDayRuleApplies_(u))return null;
  const ws=startOfDay_(asDate_(weekStartKey_(date,model.settings.weekStart)));
  const protectedDays=regularProtectedOffDaysForWeek_(u,ws,model);

  // PTO / approved Regular Off replaces one normal workday. Weekend work does
  // NOT reduce the five-day target; it simply occupies one of the five days,
  // causing the allocator to give a weekday OFF.
  return Math.max(
    0,
    model.settings.regularWorkdaysPerWeek-Math.min(model.settings.regularWorkdaysPerWeek,protectedDays)
  );
}

function wouldExceedConsecutiveDays_(username,date,state,maxDays){
  username=clean_(username);
  const candidate=startOfDay_(date);
  const existing=(state.assignedDateKeysByEmployee&&state.assignedDateKeysByEmployee[username]) || new Set();
  const candidateKey=formatDateKey_(candidate);
  let count=1;
  for(let d=addDays_(candidate,-1);existing.has(formatDateKey_(d));d=addDays_(d,-1))count++;
  for(let d=addDays_(candidate,1);existing.has(formatDateKey_(d));d=addDays_(d,1))count++;
  // If candidate is already assigned, do not double count it. Eligibility normally
  // catches ALREADY_SCHEDULED first, but this keeps the helper correct in isolation.
  if(existing.has(candidateKey))count--;
  return count>Math.max(1,num_(maxDays,5));
}

function longestConsecutiveDateRun_(dates){
  const keys=[...new Set((dates||[]).map(d=>formatDateKey_(startOfDay_(d))))].sort();
  if(!keys.length)return 0;
  let best=1,current=1;
  for(let i=1;i<keys.length;i++){
    const prev=asDate_(keys[i-1]),cur=asDate_(keys[i]);
    if(prev&&cur&&daysBetween_(prev,cur)===1)current++;
    else current=1;
    if(current>best)best=current;
  }
  return best;
}

function isPatternProtectedAssignment_(a){
  const w=clean_(a&&a.warning).toUpperCase();
  return w.indexOf('REQUIRED WEEKEND GROUP')>=0 || w.indexOf('REQUIRED RESIDENT WEEKEND')>=0 || w.indexOf('7-ON/7-OFF BLOCK')>=0 || w.indexOf('RESIDENT REQUIRED E2')>=0 || w.indexOf('PAIRED ED COVERAGE')>=0 || w.indexOf('ED REQUIRED COVERAGE')>=0;
}

function isSevenOn_(u){return clean_(u['Schedule Type']).toLowerCase().replace(/\s/g,'').indexOf('7-on')>=0 || clean_(u['Schedule Type']).toLowerCase().indexOf('7 on')>=0 || clean_(u['Schedule Type']).toLowerCase()==='7on7off';}
function sevenOnIsOnDay_(u,date){const a=startOfDay_(asDate_(u['Rotation Anchor Date']));if(!a)return false;const diff=daysBetween_(a,startOfDay_(date));const mod=((diff%14)+14)%14;return mod<7;}
function employeeWeeklyMax_(u,model){return num_(u['Weekly Hour Maximum'],model.settings.weeklyDefault);}
function employeeEveningMax_(u,model){return num_(u['Maximum Evening Shifts Per Month'],model.settings.eveningDefault);}
function creditedHours_(shift){const c=numOrNull_(shift['Credited Hours']);return c===null?num_(shift.Hours,0):c;}

function normalizeTimeString_(v) {
  if(v instanceof Date) return Utilities.formatDate(v,getTz_(),'HH:mm');
  const s=String(v===undefined||v===null?'':v).trim();
  if(!s) return '';
  const m=s.match(/^(\d{1,2}):(\d{2})(?:\s*([AP]M))?$/i);
  if(!m) throw new Error('Time must be entered as HH:MM, for example 07:00 or 15:00.');
  let h=Number(m[1]), min=Number(m[2]);
  if(min<0||min>59) throw new Error('Invalid time: '+s);
  if(m[3]){if(h<1||h>12)throw new Error('Invalid time: '+s);if(h===12)h=0;if(m[3].toUpperCase()==='PM')h+=12;}
  else if(h<0||h>23) throw new Error('Invalid time: '+s);
  return String(h).padStart(2,'0')+':'+String(min).padStart(2,'0');
}

function shiftFitsPreferredHours_(shift,u) {
  const ss=timeMinutes_(shift.Start),se=timeMinutes_(shift.End),ps=timeMinutes_(u['Preferred Start Time']),pe=timeMinutes_(u['Preferred End Time']);
  if(ss===null||se===null||ps===null||pe===null)return false;
  return intervalContains_(ps,pe,ss,se);
}
function intervalContains_(outerStart,outerEnd,innerStart,innerEnd){
  let oe=outerEnd,ie=innerEnd;if(oe<=outerStart)oe+=1440;if(ie<=innerStart)ie+=1440;
  let is=innerStart;if(is<outerStart&&ie>1440)is+=1440;
  return is>=outerStart&&ie<=oe;
}

function assignmentInterval_(date,shift){
  const sm=timeMinutes_(shift.Start),em=timeMinutes_(shift.End); if(sm===null||em===null)return null;
  const base=startOfDay_(date); const start=new Date(base.getTime()+sm*60000); let end=new Date(base.getTime()+em*60000); if(end<=start)end=new Date(end.getTime()+86400000); return{start,end};
}

function weekendGroupForDate_(date,settings){
  const groups=settings.weekendRotation,anchor=weekendSaturday_(settings.weekendAnchorDate),anchorGroup=settings.weekendAnchorGroup;
  const sat=weekendSaturday_(date),weeks=Math.floor(daysBetween_(anchor,sat)/7); let anchorIdx=Math.max(0,groups.indexOf(anchorGroup)); let idx=((anchorIdx+weeks)%groups.length+groups.length)%groups.length; return groups[idx];
}
function weekendAnchorForGroup_(group,settings){
  const groups=settings.weekendRotation||['A','B','C'];
  const targetIdx=groups.indexOf(clean_(group));
  if(targetIdx<0)return null;
  const anchor=weekendSaturday_(settings.weekendAnchorDate);
  let anchorIdx=groups.indexOf(clean_(settings.weekendAnchorGroup));
  if(anchorIdx<0)anchorIdx=0;
  const offset=((targetIdx-anchorIdx)%groups.length+groups.length)%groups.length;
  return addDays_(anchor,offset*7);
}
function employeeWeekendIsOn_(u,date,settings){
  if(isSevenOn_(u))return true;
  const group=clean_(u['Weekend Group']);
  if(!group)return false;
  const anchor=startOfDay_(asDate_(u['Weekend Rotation Anchor Date'])||weekendAnchorForGroup_(group,settings));
  if(!anchor)return false;
  const sat=weekendSaturday_(date);
  const cycleDays=Math.max(1,(settings.weekendRotation||['A','B','C']).length)*7;
  const diff=daysBetween_(anchor,sat);
  return ((diff%cycleDays)+cycleDays)%cycleDays===0;
}
function weekendSaturday_(date){const d=startOfDay_(date),day=dayIndex_(d);if(day===6)return d;if(day===0)return addDays_(d,-1);return addDays_(d,6-day);}
function isWeekendDate_(date){const d=dayIndex_(date);return d===0||d===6;}

function reasonToWarning_(r){const m={INACTIVE:'Employee is inactive.',MISSING_SKILL:'Employee does not possess the required skill.',PTO:'Employee is on approved PTO.',REGULAR_OFF:'Employee has an approved Regular Off request.',UNAVAILABLE:'Employee is unavailable on this date.',ALREADY_SCHEDULED:'Employee already has a shift on this date.',TIME_CONFLICT:'Assignment overlaps another scheduled shift.',WEEKLY_HOURS:'Assignment would exceed the employee weekly-hour maximum.',WEEKLY_DAYS:'Assignment would exceed the required five workdays in this Sunday-Saturday week.',CONSECUTIVE_DAYS:'Assignment would create more than five consecutive workdays.',EVENING_TO_MORNING:'Evening-to-morning transition is not allowed. If the pharmacist works an evening shift, the next calendar day must be OFF or another evening shift, not a Day/Morning shift.',TWO_MONTH_HOURS:'Assignment would push the pharmacist above the exact 320-hour target for the generated schedule period.',PRECEPTOR_WEEKDAY_EVENING:'Preceptor cannot work evening shifts Monday-Friday; evening shifts are allowed for preceptors only on Saturday/Sunday.',PRECEPTOR_EVENING:'Preceptor cannot normally work evening shifts.',EVENING_LIMIT:'Assignment would exceed the monthly evening-shift maximum.',WRONG_WEEKEND_GROUP:'Employee is not in the scheduled weekend group.',RESIDENT_WRONG_WEEKEND_GROUP:'Resident is not in the weekend group scheduled for this date.',WEEKEND_ANCHOR_OFF:'Employee is outside the assigned one-weekend-every-three-weeks rotation.',SEVEN_OFF:'7-on/7-off employee is in an OFF period.',SEVEN_WRONG_SHIFT:'7-on/7-off employee must remain on the dedicated N1, N2, or E shift during the ON block.',RESIDENT_E2_WEEKLY_LIMIT:'Resident already has the required E2 shift for this week.',CUSTOM_HOURS:'Shift is outside the employee hard custom work hours.',WEEKLY_AVAILABILITY_DAY:'Employee is not available on this weekday under recurring weekly availability.',WEEKLY_AVAILABILITY_TIME:'Shift is outside the employee recurring weekly available hours.',PRN_NOT_AVAILABLE:'PRN pharmacist is not available for this date/shift in My Availability.',PRN_AVAILABILITY_NOT_LOADED:'My Availability has not loaded, so PRN scheduling is blocked.',WEEKEND_NOT_ELIGIBLE:'Employee is not weekend eligible.',EVENING_NOT_ELIGIBLE:'Employee is not evening eligible.',NIGHT_NOT_ELIGIBLE:'Employee is not night eligible.',RESIDENT_RESTRICTION:'Resident settings do not allow this assignment.'};return m[r]||r;}

/** ---------------------- DASHBOARD / REPORTING -------------------- */

function computeScheduleHealth_(scheduleRows) {
  const required=scheduleRows.length,unfilled=scheduleRows.filter(r=>clean_(r.Status)==='UNFILLED'||clean_(r['Assigned Pharmacist'])==='UNFILLED').length,filled=required-unfilled;
  const warnings=scheduleRows.filter(r=>clean_(r.Warning)).length;
  return {required,filled,unfilled,coveragePercent:required?Math.round(filled/required*1000)/10:100,ruleWarnings:warnings};
}

function computeEmployeeStats_(scheduleRows,users) {
  const settings=getSettingsMap_(), weekStart=clean_(settings['Week Start']||'Sunday'), defaultMax=num_(settings['Weekly Hours Limit'],40);
  const map={}; users.filter(u=>yes_(u.Active)).forEach(u=>map[clean_(u.Username)]={username:u.Username,name:u['Pharmacist Name'],weeklyMax:num_(u['Weekly Hour Maximum'],defaultMax),totalHours:0,shifts:0,dayShifts:0,eveningShifts:0,nightShifts:0,weekendShifts:0,weeklyHours:{},weekendGroup:u['Weekend Group'],scheduleType:u['Schedule Type'],preceptor:u.Preceptor,resident:u.Resident});
  scheduleRows.forEach(r=>{
    const u=clean_(r.Username);if(!u||clean_(r.Status)==='UNFILLED')return;if(!map[u])map[u]={username:u,name:r['Assigned Pharmacist'],totalHours:0,shifts:0,dayShifts:0,eveningShifts:0,nightShifts:0,weekendShifts:0,weeklyHours:{}};
    const x=map[u];x.totalHours+=num_(r['Credited Hours'],0);x.shifts++;
    const t=clean_(r['Shift Type']).toLowerCase();if(t==='day')x.dayShifts++;if(t==='evening')x.eveningShifts++;if(t==='night')x.nightShifts++;if(yes_(r.Weekend)||isWeekendDate_(asDate_(r.Date)||new Date('2000-01-03')))x.weekendShifts++;
    const d=asDate_(r.Date); if(d){const wk=weekStartKey_(d,weekStart);x.weeklyHours[wk]=num_(x.weeklyHours[wk],0)+num_(r['Credited Hours'],0);}
  });
  return Object.keys(map).map(k=>map[k]);
}

function getFairnessReport(token) {
  const ctx=requireAuth_(token); let schedule=readTable_(APP.SHEETS.SCHEDULE),users=readTable_(APP.SHEETS.USERS);
  if(!ctx.isAdmin){schedule=schedule.filter(r=>clean_(r.Username)===ctx.username);users=users.filter(r=>clean_(r.Username)===ctx.username);}
  const stats=computeEmployeeStats_(schedule,users); const avg=stats.length?stats.reduce((s,x)=>s+x.totalHours,0)/stats.length:0;
  stats.forEach(x=>{x.hoursVsAverage=Math.round((x.totalHours-avg)*10)/10;x.workloadFlag=x.totalHours>avg*1.25?'HIGH':(x.totalHours<avg*0.75?'LOW':'NORMAL');});
  return serialize_({averageHours:Math.round(avg*10)/10,employees:stats});
}

/** ------------------------- SHEET UTILITIES ----------------------- */

function getDb_() {
  if (_DB_CACHE) return _DB_CACHE;
  const id=PropertiesService.getScriptProperties().getProperty('PHARMACY_SCHEDULER_SPREADSHEET_ID');
  if(id) _DB_CACHE=SpreadsheetApp.openById(id);
  else {
    const ss=SpreadsheetApp.getActiveSpreadsheet();
    if(!ss)throw new Error('Scheduler spreadsheet is not configured. Run setupPharmacyScheduler().');
    _DB_CACHE=ss;
  }
  return _DB_CACHE;
}
function readTable_(name){const sh=getDb_().getSheetByName(name);if(!sh)return[];return readTableFromSheet_(sh);}
function readTableFromSheet_(sh){if(!sh||sh.getLastRow()<2)return[];const data=sh.getDataRange().getValues();const headers=data[0].map(clean_);return data.slice(1).filter(r=>r.some(v=>clean_(v)!=='')).map(r=>{const o={};headers.forEach((h,i)=>o[h]=r[i]);return o;});}
function getSettingsMap_(){const m={};readTable_(APP.SHEETS.SETTINGS).forEach(r=>m[clean_(r.Setting)]=r.Value);return m;}
function appendObjectRow_(sheetName,obj){const sh=getDb_().getSheetByName(sheetName);const headers=sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0].map(clean_);sh.appendRow(headers.map(h=>sheetValue_(obj[h])));}
function findRowByKey_(sheetName,keyHeader,keyValue){return readTable_(sheetName).find(r=>clean_(r[keyHeader])===clean_(keyValue));}
function updateRowByKey_(sheetName,keyHeader,keyValue,updates){
  const sh=getDb_().getSheetByName(sheetName),data=sh.getDataRange().getValues(),headers=data[0].map(clean_),ki=headers.indexOf(keyHeader);if(ki<0)throw new Error('Missing key column '+keyHeader+' in '+sheetName);
  for(let i=1;i<data.length;i++)if(clean_(data[i][ki])===clean_(keyValue)){Object.keys(updates).forEach(h=>{const idx=headers.indexOf(h);if(idx>=0)data[i][idx]=sheetValue_(updates[h]);});sh.getRange(i+1,1,1,headers.length).setValues([data[i]]);return true;}return false;
}
function updateCompositeRow_(sheetName,predicate,updates){
  const sh=getDb_().getSheetByName(sheetName),data=sh.getDataRange().getValues(),headers=data[0].map(clean_);for(let i=1;i<data.length;i++){const o={};headers.forEach((h,j)=>o[h]=data[i][j]);if(predicate(o)){Object.keys(updates).forEach(h=>{const idx=headers.indexOf(h);if(idx>=0)data[i][idx]=sheetValue_(updates[h]);});sh.getRange(i+1,1,1,headers.length).setValues([data[i]]);return true;}}return false;
}
function deleteRowByKey_(sheetName,keyHeader,keyValue){const sh=getDb_().getSheetByName(sheetName),data=sh.getDataRange().getValues(),headers=data[0].map(clean_),ki=headers.indexOf(keyHeader);for(let i=data.length-1;i>=1;i--)if(clean_(data[i][ki])===clean_(keyValue)){sh.deleteRow(i+1);return true;}return false;}
function upsertScheduleRow_(obj){const id=clean_(obj['Assignment ID']);if(findRowByKey_(APP.SHEETS.SCHEDULE,'Assignment ID',id))updateRowByKey_(APP.SHEETS.SCHEDULE,'Assignment ID',id,obj);else appendObjectRow_(APP.SHEETS.SCHEDULE,obj);}


/**
 * v30 legacy migration: guarantee every physical Schedule row has a unique
 * Assignment ID. Chunked generation in older versions reused sequence-based
 * IDs across weeks. The first occurrence keeps its ID; later duplicates (and
 * blank IDs) receive a new stable ID written directly to the sheet.
 *
 * Assignment ID is an internal row identifier only; changing a duplicated ID
 * does not change Date, Shift, pharmacist, lock state, coverage, or any other
 * schedule data.
 */
function ensureUniqueScheduleAssignmentIds_() {
  const sh=getDb_().getSheetByName(APP.SHEETS.SCHEDULE);
  if(!sh || sh.getLastRow()<2) return {changed:0};
  const data=sh.getDataRange().getValues();
  const headers=data[0].map(clean_);
  const idIdx=headers.indexOf('Assignment ID');
  const dateIdx=headers.indexOf('Date');
  const shiftIdx=headers.indexOf('Shift');
  const slotIdx=headers.indexOf('Slot');
  const genIdx=headers.indexOf('Generation ID');
  if(idIdx<0) return {changed:0};

  const seen=new Set();
  const changedRows=[];
  for(let i=1;i<data.length;i++){
    if(!data[i].some(v=>clean_(v)!=='')) continue;
    let id=clean_(data[i][idIdx]);
    if(id && !seen.has(id)){
      seen.add(id);
      continue;
    }

    const d=dateIdx>=0?asDate_(data[i][dateIdx]):null;
    const datePart=d?formatDateKey_(d).replace(/-/g,''):'NODATE';
    const shiftPart=(shiftIdx>=0?clean_(data[i][shiftIdx]):'SHIFT').replace(/[^A-Za-z0-9_-]+/g,'_') || 'SHIFT';
    const slotPart='S'+String(slotIdx>=0?num_(data[i][slotIdx],1):1);
    const genPart=(genIdx>=0?clean_(data[i][genIdx]):'GEN').replace(/[^A-Za-z0-9_-]+/g,'_').slice(-42) || 'GEN';
    let replacement=['FIX',genPart,datePart,shiftPart,slotPart,Utilities.getUuid().replace(/-/g,'').slice(0,8).toUpperCase()].join('-');
    while(seen.has(replacement)) replacement='FIX-'+Utilities.getUuid();
    data[i][idIdx]=replacement;
    seen.add(replacement);
    changedRows.push(i+1);
  }

  // Write only the Assignment ID cells that changed so no other values or
  // formatting in the Schedule sheet are touched.
  changedRows.forEach(rowNumber=>{
    sh.getRange(rowNumber,idIdx+1).setValue(data[rowNumber-1][idIdx]);
  });
  if(changedRows.length) SpreadsheetApp.flush();
  return {changed:changedRows.length};
}

/**
 * Optional editor utility. Run once manually if you want to repair legacy
 * duplicate Assignment IDs immediately without waiting for the web app load.
 */
function repairDuplicateScheduleAssignmentIds() {
  return ensureUniqueScheduleAssignmentIds_();
}

/**
 * Remove stale duplicate UNFILLED rows for one exact staffing slot.
 * A required slot is identified by Date + Shift + Slot. When an admin fills
 * an UNFILLED calendar card manually, there must not be a second red UNFILLED
 * card left behind for that same slot.
 */
function removeDuplicateUnfilledRowsForSlot_(date, shiftCode, slotNumber, keepAssignmentId) {
  const sh=getDb_().getSheetByName(APP.SHEETS.SCHEDULE);
  if(!sh || sh.getLastRow()<2) return 0;
  const data=sh.getDataRange().getValues();
  const headers=data[0].map(clean_);
  const dateIdx=headers.indexOf('Date');
  const shiftIdx=headers.indexOf('Shift');
  const slotIdx=headers.indexOf('Slot');
  const statusIdx=headers.indexOf('Status');
  const assignedIdx=headers.indexOf('Assigned Pharmacist');
  const idIdx=headers.indexOf('Assignment ID');
  if([dateIdx,shiftIdx,slotIdx,statusIdx,idIdx].some(i=>i<0)) return 0;

  const targetDate=formatDateKey_(startOfDay_(asDate_(date)));
  const targetShift=clean_(shiftCode);
  const targetSlot=num_(slotNumber,1);
  const keepId=clean_(keepAssignmentId);
  const rowsToDelete=[];

  for(let i=1;i<data.length;i++){
    const row=data[i];
    if(formatDateKey_(startOfDay_(asDate_(row[dateIdx])))!==targetDate) continue;
    if(clean_(row[shiftIdx])!==targetShift) continue;
    if(num_(row[slotIdx],1)!==targetSlot) continue;
    if(clean_(row[idIdx])===keepId) continue;
    const status=clean_(row[statusIdx]).toUpperCase();
    const assigned=assignedIdx>=0?clean_(row[assignedIdx]).toUpperCase():'';
    if(status==='UNFILLED' || assigned==='UNFILLED') rowsToDelete.push(i+1);
  }

  rowsToDelete.sort((a,b)=>b-a).forEach(r=>sh.deleteRow(r));
  return rowsToDelete.length;
}

/**
 * Collapse duplicate UNFILLED rows for a required slot after a manual removal.
 * Keeps the row identified by keepAssignmentId and removes any other UNFILLED
 * copies for the same Date + Shift + Slot.
 */
function collapseDuplicateUnfilledRowsForSlot_(date, shiftCode, slotNumber, keepAssignmentId) {
  return removeDuplicateUnfilledRowsForSlot_(date, shiftCode, slotNumber, keepAssignmentId);
}

/**
 * Sheet-wide reconciliation for stale UNFILLED rows.
 * If at least one non-UNFILLED assignment exists for the exact same
 * Date + Shift + Slot, all UNFILLED copies for that slot are removed.
 * This is intentionally conservative: it never deletes a filled row.
 */
function reconcileFilledVsUnfilledScheduleRows_() {
  const sh=getDb_().getSheetByName(APP.SHEETS.SCHEDULE);
  if(!sh || sh.getLastRow()<2) return 0;
  const data=sh.getDataRange().getValues();
  const headers=data[0].map(clean_);
  const dateIdx=headers.indexOf('Date');
  const shiftIdx=headers.indexOf('Shift');
  const slotIdx=headers.indexOf('Slot');
  const statusIdx=headers.indexOf('Status');
  const assignedIdx=headers.indexOf('Assigned Pharmacist');
  if([dateIdx,shiftIdx,slotIdx,statusIdx,assignedIdx].some(i=>i<0)) return 0;

  const groups={};
  for(let i=1;i<data.length;i++){
    const d=startOfDay_(asDate_(data[i][dateIdx]));
    if(!d) continue;
    const key=formatDateKey_(d)+'||'+clean_(data[i][shiftIdx])+'||'+num_(data[i][slotIdx],1);
    if(!groups[key]) groups[key]={filled:false,unfilledRows:[]};
    const status=clean_(data[i][statusIdx]).toUpperCase();
    const assigned=clean_(data[i][assignedIdx]).toUpperCase();
    const isUnfilled=status==='UNFILLED' || assigned==='UNFILLED' || !assigned;
    if(isUnfilled) groups[key].unfilledRows.push(i+1);
    else groups[key].filled=true;
  }

  const rowsToDelete=[];
  Object.keys(groups).forEach(k=>{
    const g=groups[k];
    if(g.filled && g.unfilledRows.length) rowsToDelete.push.apply(rowsToDelete,g.unfilledRows);
  });
  rowsToDelete.sort((a,b)=>b-a).forEach(r=>sh.deleteRow(r));
  if(rowsToDelete.length) SpreadsheetApp.flush();
  return rowsToDelete.length;
}

/**
 * v24: Repair legacy manual assignments that were saved before the manual
 * editor knew how to write OFF-day coverage metadata. This does NOT grant new
 * skills. It only labels an already-saved manual assignment when the assigned
 * pharmacist has the matching Users -> Off-Day Coverage Skills entry and the
 * unique primary/home pharmacist is on an algorithm-generated OFF day.
 */
function repairManualOffDayCoverageMetadata_(model) {
  model=model||loadSchedulingModel_();
  const sh=getDb_().getSheetByName(APP.SHEETS.SCHEDULE);
  if(!sh||sh.getLastRow()<2)return 0;
  const data=sh.getDataRange().getValues();
  const headers=data[0].map(clean_);
  const idx={};headers.forEach((h,i)=>idx[h]=i);
  const required=['Date','Shift','Slot','Assigned Pharmacist','Username','Manual','Status','Coverage For Pharmacist','Coverage For Username','Coverage Reason','Warning'];
  if(required.some(h=>idx[h]===undefined))return 0;

  const objects=[];
  for(let i=1;i<data.length;i++){
    const o={};headers.forEach((h,j)=>o[h]=data[i][j]);
    if(o.Date)objects.push(o);
  }
  const assignments=objects
    .filter(r=>clean_(r.Status).toUpperCase()!=='UNFILLED'&&clean_(r.Username))
    .map(r=>scheduleRowToAssignment_(r,model))
    .filter(Boolean);
  if(!assignments.length)return 0;

  const dates=assignments.map(a=>a.date).filter(Boolean).sort((a,b)=>a-b);
  const state=createState_(model,[],dates[0],dates[dates.length-1]);
  assignments.forEach(a=>addAssignmentToState_(state,a,model));

  let changed=0;
  for(let i=1;i<data.length;i++){
    const row={};headers.forEach((h,j)=>row[h]=data[i][j]);
    if(!yes_(row.Manual))continue;
    if(clean_(row.Status).toUpperCase()==='UNFILLED'||!clean_(row.Username))continue;
    if(clean_(row['Coverage Reason']).toUpperCase()==='GENERATED OFF DAY'&&clean_(row['Coverage For Username']))continue;

    const u=model.usersByUsername[clean_(row.Username)]||model.usersByName[clean_(row['Assigned Pharmacist'])];
    const a=scheduleRowToAssignment_(row,model);
    if(!u||!a)continue;
    const slot=assignmentAsSlot_(a);
    if(hasRequiredSkillForSlot_(u,slot,model))continue;
    if(!hasOffDayCoverageSkillForSlot_(u,slot))continue;

    const owners=primaryHomeOwnersForSlot_(slot,model)
      .filter(owner=>clean_(owner.Username)!==clean_(u.Username));
    const offOwners=owners.filter(owner=>isGeneratedOffDayForHomeOwner_(owner,slot,state,model));
    if(offOwners.length!==1)continue;
    const owner=offOwners[0];

    data[i][idx['Coverage For Pharmacist']]=clean_(owner['Pharmacist Name']);
    data[i][idx['Coverage For Username']]=clean_(owner.Username);
    data[i][idx['Coverage Reason']]='GENERATED OFF DAY';
    const existingWarning=clean_(data[i][idx['Warning']]);
    const marker='OFF-DAY COVERAGE ONLY: Covering '+clean_(owner['Pharmacist Name'])+' on an algorithm-generated OFF day.';
    if(existingWarning.indexOf(marker)<0)data[i][idx['Warning']]=existingWarning?existingWarning+' | '+marker:marker;
    changed++;
  }

  if(changed){
    sh.getRange(2,1,data.length-1,data[0].length).setValues(data.slice(1));
    SpreadsheetApp.flush();
  }
  return changed;
}

function repairManualOffDayCoverageMetadata() {
  const lock=LockService.getScriptLock();
  lock.waitLock(30000);
  try{
    const repaired=repairManualOffDayCoverageMetadata_(loadSchedulingModel_());
    return {ok:true,repaired:repaired,message:'Repaired OFF-day coverage metadata on '+repaired+' manual Schedule row(s).'};
  } finally {
    lock.releaseLock();
  }
}

/**
 * Run this once from the Apps Script editor after installing v24 if the
 * existing Schedule sheet already contains stale UNFILLED rows beside filled
 * manual assignments. It is safe to run repeatedly.
 */
function repairStaleUnfilledScheduleRows() {
  const lock=LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const removed=reconcileFilledVsUnfilledScheduleRows_();
    return {ok:true,removed:removed,message:'Removed '+removed+' stale UNFILLED row(s) from the Schedule sheet.'};
  } finally {
    lock.releaseLock();
  }
}
function findUser_(usernameOrId,name){const users=readTable_(APP.SHEETS.USERS).filter(u=>clean_(u.Role).toLowerCase()!=='administrator');return users.find(u=>clean_(u.Username)===clean_(usernameOrId))||users.find(u=>clean_(u['Employee ID'])===clean_(usernameOrId))||users.find(u=>clean_(u['Pharmacist Name'])===clean_(name));}

function audit_(action,date,pharmacist,oldShift,newShift,oldValue,newValue,override,overrideReason,details,user){
  appendObjectRow_(APP.SHEETS.AUDIT,{'Timestamp':new Date(),'User':user||'SYSTEM','Action':action,'Date':date||'','Pharmacist':pharmacist||'','Old Shift':oldShift||'','New Shift':newShift||'','Old Value':oldValue||'','New Value':newValue||'','Override':override||'No','Override Reason':overrideReason||'','Details':details||''});
}

/** -------------------------- DATE / VALUE ------------------------- */

function getTz_(){
  if(_TZ_CACHE)return String(_TZ_CACHE);
  try{_TZ_CACHE=getDb_().getSpreadsheetTimeZone()||Session.getScriptTimeZone()||'America/New_York';}
  catch(e){_TZ_CACHE=Session.getScriptTimeZone()||'America/New_York';}
  _TZ_CACHE=String(_TZ_CACHE||'America/New_York');
  return _TZ_CACHE;
}
function tzOffsetMinutesAt_(date,tz){
  const z=Utilities.formatDate(date,tz,'Z');
  const m=String(z).match(/^([+-])(\d{2})(\d{2})$/);
  if(!m)return 0;
  const mins=Number(m[2])*60+Number(m[3]);
  return m[1]==='-'?-mins:mins;
}
function dateFromKey_(key){
  const m=String(key||'').match(/^(\d{4})-(\d{2})-(\d{2})$/);if(!m)return null;
  const y=Number(m[1]),mo=Number(m[2])-1,d=Number(m[3]),tz=getTz_(),utc=Date.UTC(y,mo,d,0,0,0);
  let probe=new Date(utc),off=tzOffsetMinutesAt_(probe,tz),candidate=new Date(utc-off*60000);
  const off2=tzOffsetMinutesAt_(candidate,tz);if(off2!==off)candidate=new Date(utc-off2*60000);
  return candidate;
}
function asDate_(v){
  if(v instanceof Date&&!isNaN(v))return new Date(v);
  if(v===null||v===undefined||v==='')return null;
  const s=String(v).trim();
  if(/^\d{4}-\d{2}-\d{2}$/.test(s))return dateFromKey_(s);
  const d=new Date(s);return isNaN(d)?null:d;
}
function formatDateKey_(d){return Utilities.formatDate(d,getTz_(),'yyyy-MM-dd');}
function startOfDay_(d){if(!d)return null;return dateFromKey_(formatDateKey_(d));}
function addDays_(d,n){
  const k=formatDateKey_(d),m=k.match(/^(\d{4})-(\d{2})-(\d{2})$/),u=new Date(Date.UTC(Number(m[1]),Number(m[2])-1,Number(m[3])+n));
  return dateFromKey_(u.getUTCFullYear()+'-'+String(u.getUTCMonth()+1).padStart(2,'0')+'-'+String(u.getUTCDate()).padStart(2,'0'));
}
function daysBetween_(a,b){
  const ka=formatDateKey_(a).split('-').map(Number),kb=formatDateKey_(b).split('-').map(Number);
  return Math.round((Date.UTC(kb[0],kb[1]-1,kb[2])-Date.UTC(ka[0],ka[1]-1,ka[2]))/86400000);
}
function sameDate_(a,b){return !!a&&!!b&&formatDateKey_(a)===formatDateKey_(b);}
function inDateRange_(d,s,e){if(!d||!s||!e)return false;const k=formatDateKey_(d),ks=formatDateKey_(s),ke=formatDateKey_(e);return k>=ks&&k<=ke;}
function dateKey_(v){const d=asDate_(v);return d?formatDateKey_(d):clean_(v);}
function dayName_(d){return Utilities.formatDate(d,getTz_(),'EEEE');}
function dayNames_(){return ['Sunday','Monday','Tuesday','Wednesday','Thursday','Friday','Saturday'];}
function dayIndex_(d){return dayNames_().indexOf(dayName_(d));}
function monthKey_(d){return Utilities.formatDate(d,getTz_(),'yyyy-MM');}
function weekStartKey_(date,weekStartName){const names=dayNames_();let target=names.indexOf(clean_(weekStartName));if(target<0)target=1;const idx=dayIndex_(date),diff=(idx-target+7)%7;return formatDateKey_(addDays_(date,-diff));}
function nextOrSameSaturday_(d){const diff=(6-dayIndex_(d)+7)%7;return addDays_(d,diff);}
function slotKey_(dateKey,shift,slot){return dateKey+'|'+shift+'|'+slot;}
function makeGenerationId_(s,e){return 'GEN_'+Utilities.formatDate(s,getTz_(),'yyyyMMdd')+'_'+Utilities.formatDate(e,getTz_(),'yyyyMMdd')+'_'+Utilities.formatDate(new Date(),getTz_(),'HHmmss');}
function firstOfMonth_(d){const k=formatDateKey_(d).split('-');return dateFromKey_(k[0]+'-'+k[1]+'-01');}
function lastOfMonth_(d){const k=formatDateKey_(d).split('-').map(Number),u=new Date(Date.UTC(k[0],k[1],0));return dateFromKey_(u.getUTCFullYear()+'-'+String(u.getUTCMonth()+1).padStart(2,'0')+'-'+String(u.getUTCDate()).padStart(2,'0'));}
function generationContextStart_(start,weekStart){const ws=asDate_(weekStartKey_(start,weekStart));const ms=firstOfMonth_(start);const base=formatDateKey_(ws)<formatDateKey_(ms)?ws:ms;return addDays_(base,-7);}
function generationContextEnd_(end,weekStart){const ws=asDate_(weekStartKey_(end,weekStart));const we=addDays_(ws,6),me=lastOfMonth_(end);const base=formatDateKey_(we)>formatDateKey_(me)?we:me;return addDays_(base,7);}
function timeMinutes_(v){
  if(v instanceof Date){const parts=Utilities.formatDate(v,getTz_(),'HH:mm').split(':').map(Number);return parts[0]*60+parts[1];}
  const s=clean_(v);if(!s)return null;const m=s.match(/^(\d{1,2}):(\d{2})(?:\s*([AP]M))?$/i);if(!m)return null;let h=Number(m[1]),min=Number(m[2]);if(m[3]){if(h===12)h=0;if(m[3].toUpperCase()==='PM')h+=12;}return h*60+min;
}

function clean_(v){if(v===null||v===undefined)return'';if(v instanceof Date)return v;return String(v).trim();}
function yes_(v){return ['yes','y','true','1','on'].includes(String(v).trim().toLowerCase());}
function yesDefault_(v,def){return clean_(v)===''?def:yes_(v);}
function num_(v,def){const n=Number(v);return isFinite(n)?n:def;}
function numOrNull_(v){if(v===null||v===undefined||String(v).trim()==='')return null;const n=Number(v);return isFinite(n)?n:null;}
function unique_(arr){return Array.from(new Set(arr));}
function deterministicTie_(s){let h=0;for(let i=0;i<s.length;i++)h=(h*31+s.charCodeAt(i))>>>0;return (h%1000)/1000000;}
function sheetValue_(v){if(v===undefined||v===null)return'';return v;}
function serialize_(value,keyName){
  if(value instanceof Date){
    const dateOnly=['Date','Start Date','End Date','Rotation Anchor Date','Weekend Rotation Anchor Date','Weekend Saturday','Weekend Sunday','Effective Start','Effective End'];
    return dateOnly.includes(clean_(keyName)) ? formatDateKey_(value) : Utilities.formatDate(value,getTz_(),'yyyy-MM-dd HH:mm:ss');
  }
  if(Array.isArray(value))return value.map(v=>serialize_(v,keyName));
  if(value&&typeof value==='object'){const o={};Object.keys(value).forEach(k=>o[k]=serialize_(value[k],k));return o;}
  return value;
}
function hashPassword_(password,salt){const bytes=Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256,String(salt)+'|'+String(password),Utilities.Charset.UTF_8);return Utilities.base64EncodeWebSafe(bytes);}
function generateTemporaryPassword_(){return 'Rx!'+Utilities.getUuid().replace(/-/g,'').slice(0,14)+'9a';}
