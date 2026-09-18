/**
 * ONE-TIME MIGRATION HELPER ONLY.
 * Run this inside the CURRENT Google Sheet-bound NeoChrono Apps Script project.
 * It does NOT become part of the GitHub runtime.
 *
 * It creates NeoChrono_GitHub_Migration.json in your Google Drive.
 * Then open import.html in the new GitHub build and import that JSON file.
 */
function exportNeoChronoForGitHub() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('Open the current NeoChrono Google Sheet first.');
  const tz = ss.getSpreadsheetTimeZone() || Session.getScriptTimeZone() || 'America/New_York';

  const admins = readTableForMigration_(ss, 'Admin Accounts');
  const users = readTableForMigration_(ss, 'Users');
  const skills = readTableForMigration_(ss, 'Employee Skills');
  const shifts = readTableForMigration_(ss, 'Shifts');
  const staffing = readTableForMigration_(ss, 'Staffing Requirements');
  const requests = readTableForMigration_(ss, 'PTO / Availability Requests');
  const weekly = readTableForMigration_(ss, 'Employee Weekly Availability');
  const schedule = readTableForMigration_(ss, 'Schedule');
  const settingsRows = readTableForMigration_(ss, 'Settings');
  const audit = readTableForMigration_(ss, 'Audit Log');
  const preceptorRows = firstExistingTableForMigration_(ss, ['Preceptor Calendar','Preceptor Schedule']);
  const swapRows = firstExistingTableForMigration_(ss, ['Swap Market','Shift Swaps','Swaps']);

  const settingsMap = {};
  settingsRows.forEach(function(r){ settingsMap[String(r.Setting || '').trim()] = r.Value; });
  const bool = function(v, fallback) {
    if (v === '' || v === null || v === undefined) return !!fallback;
    return /^(yes|true|1|y)$/i.test(String(v).trim());
  };
  const num = function(v, fallback) { const n = Number(v); return isFinite(n) ? n : fallback; };
  const text = function(v) { return String(v === null || v === undefined ? '' : v).trim(); };
  const date = function(v) {
    if (!v) return '';
    if (Object.prototype.toString.call(v) === '[object Date]' && !isNaN(v)) return Utilities.formatDate(v, tz, 'yyyy-MM-dd');
    const s = String(v).trim();
    const d = new Date(s);
    return isNaN(d) ? s : Utilities.formatDate(d, tz, 'yyyy-MM-dd');
  };
  const iso = function(v) {
    if (!v) return '';
    if (Object.prototype.toString.call(v) === '[object Date]' && !isNaN(v)) return v.toISOString();
    const d = new Date(v);
    return isNaN(d) ? String(v) : d.toISOString();
  };
  const arr = function(v) { return text(v).split(/[,;|\/\n]+/).map(function(x){return x.trim().toUpperCase();}).filter(Boolean); };

  const db = {
    meta: {schemaVersion:1, appVersion:'migrated-from-apps-script', createdAt:new Date().toISOString(), updatedAt:new Date().toISOString(), sourceSpreadsheetId:ss.getId(), sourceSpreadsheetName:ss.getName()},
    settings: {
      appName: text(settingsMap['Schedule Title']) || 'NeoChrono',
      scheduleDays: num(settingsMap['Default Schedule Days'], 56),
      weekStart: text(settingsMap['Week Start']) || 'Sunday',
      requiredHoursPerSchedulePeriod: num(settingsMap['Required Hours Per Schedule Period'], 320),
      maximumHoursPerSchedulePeriod: num(settingsMap['Maximum Hours Per Schedule Period'], 320),
      weeklyHoursLimit: num(settingsMap['Weekly Hours Limit'], 40),
      regularWorkdaysPerWeek: num(settingsMap['Regular Workdays Per Week'], 5),
      maximumConsecutiveWorkdays: num(settingsMap['Maximum Consecutive Workdays'], 5),
      maximumEveningShifts: num(settingsMap['Maximum Evening Shifts'], 7),
      weekendRotation: arr(settingsMap['Weekend Rotation'] || 'A,B,C'),
      weekendAnchorDate: date(settingsMap['Weekend Anchor Date']),
      weekendAnchorGroup: text(settingsMap['Weekend Anchor Group']) || 'A',
      allowWeekendFallback: bool(settingsMap['Allow Weekend Fallback'], false),
      requiredWeekendAssignment: bool(settingsMap['Required Weekend Assignment'], true),
      residentE2ShiftCode: text(settingsMap['Resident E2 Shift Code']) || 'E2',
      residentE2ShiftsPerWeek: num(settingsMap['Resident E2 Shifts Per Week'], 1),
      residentsRequiredAssignedWeekend: bool(settingsMap['Residents Required Assigned Weekend'], true),
      sevenOnBlockScheduling: bool(settingsMap['SevenOn Block Scheduling'], true),
      blockEveningToMorningTransition: bool(settingsMap['Block Evening To Morning Transition'], true),
      ptoAutoApprovalLimitPerDate: num(settingsMap['PTO Auto-Approve Per Day'], 2),
      allowAdminRuleOverride: bool(settingsMap['Allow Admin Rule Override'], true),
      showFullScheduleToPharmacists: bool(settingsMap['Full Schedule Visible To Employees'], true),
      preceptorEveningTargetPerMonthWhenOff: 7,
      preceptorHomeUnitRequiredWhenOn: true,
      preceptorOffAllowsEvenings: true,
      eddEdePairRequired: bool(settingsMap['Require Paired ED Coverage'], true),
      edDayShiftCode: text(settingsMap['ED Day Shift Code']) || 'EDD',
      edEveningShiftCode: text(settingsMap['ED Evening Shift Code']) || 'EDE',
      minimumRestHours: 8,
      scheduleTimezone: tz
    },
    admins: admins.map(function(r){return {
      id:text(r['Admin ID']) || Utilities.getUuid(), name:text(r['Admin Name']), username:text(r.Username), role:'Administrator',
      active:bool(r.Active,true), passwordHash:text(r['Password Hash']), passwordSalt:text(r['Password Salt']),
      passwordAlgorithm:'', passwordIterations:null, mustChangePassword:bool(r['Must Change Password'],false),
      lastLogin:iso(r['Last Login']), updatedAt:iso(r['Updated At']), updatedBy:text(r['Updated By'])
    };}).filter(function(r){return r.username;}),
    users: users.map(function(r){return {
      id:text(r['Employee ID']) || Utilities.getUuid(), pharmacistName:text(r['Pharmacist Name']), name:text(r['Pharmacist Name']), username:text(r.Username), role:text(r.Role)||'Pharmacist',
      active:bool(r.Active,true), scheduleType:text(r['Schedule Type'])||'Regular', preceptor:bool(r.Preceptor,false), resident:bool(r.Resident,false),
      weekendGroup:text(r['Weekend Group']), weekendRotationAnchorDate:date(r['Weekend Rotation Anchor Date']), weeklyHourMaximum:num(r['Weekly Hour Maximum'],40), targetWeeklyHours:num(r['Target Weekly Hours'],40),
      maximumEveningShiftsPerMonth:num(r['Maximum Evening Shifts Per Month'],7), preferredStartTime:text(r['Preferred Start Time']), preferredEndTime:text(r['Preferred End Time']),
      customHoursEnabled:bool(r['Custom Hours Enabled'],false), customHoursMode:text(r['Custom Hours Mode'])||'SOFT', preferredShiftType:text(r['Preferred Shift Type']), offDayCoverageSkills:arr(r['Off-Day Coverage Skills']),
      weekendEligible:bool(r['Weekend Eligible'],true), eveningEligible:bool(r['Evening Eligible'],true), nightEligible:bool(r['Night Eligible'],true), residentCoversRegular:bool(r['Resident Covers Regular'],false),
      residentWeekends:bool(r['Resident Weekends'],true), residentEvenings:bool(r['Resident Evenings'],true), residentNights:bool(r['Resident Nights'],false), rotationAnchorDate:date(r['Rotation Anchor Date']),
      sevenOnWeeklyHandling:text(r['SevenOn Weekly Handling'])||'ENFORCE_MAX', notes:text(r.Notes), mustChangePassword:bool(r['Must Change Password'],false),
      passwordHash:text(r['Password Hash']), passwordSalt:text(r['Password Salt']), passwordAlgorithm:'', passwordIterations:null,
      lastLogin:iso(r['Last Login']), updatedAt:iso(r['Updated At']), updatedBy:text(r['Updated By'])
    };}).filter(function(r){return r.username || r.pharmacistName;}),
    skills: skills.map(function(r){return {id:text(r['Skill ID'])||Utilities.getUuid(),employeeId:text(r['Employee ID']),pharmacistName:text(r['Pharmacist Name']),username:text(r.Username),skill:text(r.Skill).toUpperCase(),active:bool(r.Active,true),updatedAt:iso(r['Updated At']),updatedBy:text(r['Updated By'])};}).filter(function(r){return r.username&&r.skill;}),
    shifts: shifts.map(function(r){return {shift:text(r.Shift).toUpperCase(),meaning:text(r.Meaning),start:text(r.Start),end:text(r.End),hours:num(r.Hours,0),creditedHours:num(r['Credited Hours'],num(r.Hours,0)),type:text(r.Type)||'Day',skill:text(r.Skill).toUpperCase(),weekend:bool(r.Weekend,false),active:bool(r.Active,true),priority:num(r.Priority,50),updatedAt:iso(r['Updated At']),updatedBy:text(r['Updated By'])};}).filter(function(r){return r.shift;}),
    staffingRequirements: staffing.map(function(r){return {id:Utilities.getUuid(),shift:text(r.Shift).toUpperCase(),Sunday:num(r.Sunday,0),Monday:num(r.Monday,0),Tuesday:num(r.Tuesday,0),Wednesday:num(r.Wednesday,0),Thursday:num(r.Thursday,0),Friday:num(r.Friday,0),Saturday:num(r.Saturday,0),active:bool(r.Active,true),updatedAt:iso(r['Updated At']),updatedBy:text(r['Updated By'])};}).filter(function(r){return r.shift;}),
    requests: requests.map(function(r){return {recordType:text(r['Record Type']),recordId:text(r['Record ID'])||Utilities.getUuid(),pharmacist:text(r.Pharmacist),username:text(r.Username),date:date(r.Date),startDate:date(r['Start Date']),endDate:date(r['End Date']),weekendSaturday:date(r['Weekend Saturday']),weekendSunday:date(r['Weekend Sunday']),available:r.Available===''?'':bool(r.Available,false),status:text(r.Status),comment:text(r.Comment),submittedAt:iso(r['Submitted At']),reviewedBy:text(r['Reviewed By']),reviewedAt:iso(r['Reviewed At']),updatedAt:iso(r['Updated At']),updatedBy:text(r['Updated By'])};}),
    weeklyAvailability: weekly.map(function(r){return {id:text(r['Availability ID'])||Utilities.getUuid(),employeeId:text(r['Employee ID']),pharmacistName:text(r['Pharmacist Name']),username:text(r.Username),day:text(r.Day).toUpperCase(),available:bool(r.Available,true),startTime:text(r['Start Time']),endTime:text(r['End Time']),ruleType:text(r['Rule Type'])||'AVAILABLE',effectiveStart:date(r['Effective Start']),effectiveEnd:date(r['Effective End']),notes:text(r.Notes),active:bool(r.Active,true),updatedAt:iso(r['Updated At']),updatedBy:text(r['Updated By'])};}),
    preceptorCalendar: preceptorRows.map(function(r){return {id:text(r.ID||r['Record ID'])||Utilities.getUuid(),username:text(r.Username),pharmacistName:text(r['Pharmacist Name']||r.Pharmacist),date:date(r.Date),precepting:bool(r.Precepting, String(r.Status).toUpperCase()==='ON'),status:text(r.Status)|| (bool(r.Precepting,false)?'ON':'OFF'),rotation:text(r.Rotation),notes:text(r.Notes),updatedAt:iso(r['Updated At']),updatedBy:text(r['Updated By'])};}),
    schedule: schedule.map(function(r){return {generationId:text(r['Generation ID']),assignmentId:text(r['Assignment ID'])||Utilities.getUuid(),date:date(r.Date),day:text(r.Day),shift:text(r.Shift),slot:num(r.Slot,1),assignedPharmacist:text(r['Assigned Pharmacist']),username:text(r.Username),hours:num(r.Hours,0),creditedHours:num(r['Credited Hours'],num(r.Hours,0)),requiredSkill:text(r['Required Skill']),coverageForPharmacist:text(r['Coverage For Pharmacist']),coverageForUsername:text(r['Coverage For Username']),coverageReason:text(r['Coverage Reason']),shiftType:text(r['Shift Type']),weekend:bool(r.Weekend,false),weekendGroup:text(r['Weekend Group']),holiday:text(r.Holiday),locked:bool(r.Locked,false),manual:bool(r.Manual,false),status:text(r.Status),warning:text(r.Warning),updatedAt:iso(r['Updated At']),updatedBy:text(r['Updated By']),finalizedAt:iso(r['Finalized At'])};}),
    schedulePeriods: [],
    swaps: swapRows,
    audit: audit.map(function(r){return {id:Utilities.getUuid(),timestamp:iso(r.Timestamp),user:text(r.User),action:text(r.Action),date:date(r.Date),pharmacist:text(r.Pharmacist),oldShift:text(r['Old Shift']),newShift:text(r['New Shift']),oldValue:text(r['Old Value']),newValue:text(r['New Value']),override:bool(r.Override,false),overrideReason:text(r['Override Reason']),details:text(r.Details)};}),
    sessions: [],
    publishedSchedule: []
  };

  const byGen = {};
  db.schedule.forEach(function(r){
    if (!r.generationId) return;
    if (!byGen[r.generationId]) byGen[r.generationId] = [];
    byGen[r.generationId].push(r);
  });
  Object.keys(byGen).forEach(function(g){
    const rows=byGen[g], dates=rows.map(function(r){return r.date;}).filter(Boolean).sort();
    if (!dates.length) return;
    const finalized=rows.filter(function(r){return !!r.finalizedAt;});
    db.schedulePeriods.push({generationId:g,startDate:dates[0],endDate:dates[dates.length-1],status:finalized.length?'FINALIZED':'DRAFT',createdBy:'MIGRATION',createdAt:new Date().toISOString(),finalizedBy:'',finalizedAt:finalized.length?finalized[0].finalizedAt:''});
  });
  const latestFinalized=db.schedulePeriods.filter(function(p){return p.status==='FINALIZED';}).sort(function(a,b){return String(b.finalizedAt).localeCompare(String(a.finalizedAt));})[0];
  if(latestFinalized) db.publishedSchedule=db.schedule.filter(function(r){return r.generationId===latestFinalized.generationId;});

  const json = JSON.stringify(db, null, 2);
  const file = DriveApp.createFile('NeoChrono_GitHub_Migration.json', json, MimeType.PLAIN_TEXT);
  const result = {ok:true,fileId:file.getId(),fileName:file.getName(),url:file.getUrl(),bytes:json.length};
  Logger.log(JSON.stringify(result));
  return result;
}

function readTableForMigration_(ss, name) {
  const sh = ss.getSheetByName(name);
  if (!sh || sh.getLastRow() < 2 || sh.getLastColumn() < 1) return [];
  const values = sh.getDataRange().getValues();
  const headers = values[0].map(function(v){return String(v || '').trim();});
  return values.slice(1).filter(function(r){return r.some(function(v){return String(v === null || v === undefined ? '' : v).trim() !== '';});}).map(function(r){const o={};headers.forEach(function(h,i){if(h)o[h]=r[i];});return o;});
}

function firstExistingTableForMigration_(ss, names) {
  for (let i=0;i<names.length;i++) {
    const rows=readTableForMigration_(ss,names[i]);
    if(rows.length || ss.getSheetByName(names[i])) return rows;
  }
  return [];
}
