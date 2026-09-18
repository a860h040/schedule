import { uid, clean, yes, num, dateKey } from './schema.js';

const ALIASES={
  users:['Users','Pharmacists','Employees'],
  skills:['Employee Skills','Skills'],
  shifts:['Shifts','Shift Definitions'],
  staffingRequirements:['Staffing Requirements','Staffing','Requirements'],
  requests:['PTO / Availability Requests','PTO Availability Requests','PTO','Requests'],
  weeklyAvailability:['Employee Weekly Availability','Weekly Availability'],
  preceptorCalendar:['Preceptor Calendar','Preceptor Schedule'],
  settingsRows:['Settings'],
  schedule:['Schedule'],
  audit:['Audit Log','Audit'],
  swaps:['Swap Market','Shift Swaps','Swaps']
};
const SETUP=['users','skills','shifts','staffingRequirements','requests','weeklyAvailability','preceptorCalendar'];

const txt=v=>String(v??'').trim();
const bool=(v,d=false)=>v===''||v==null?!!d:(typeof v==='boolean'?v:/^(yes|true|1|y|on)$/i.test(txt(v)));
const n=(v,d=0)=>Number.isFinite(Number(v))?Number(v):d;
const list=v=>txt(v).split(/[,;|/\n]+/).map(x=>x.trim().toUpperCase()).filter(Boolean);
const iso=v=>{if(!v)return'';if(v instanceof Date&&!isNaN(v))return v.toISOString();const d=new Date(v);return isNaN(d)?txt(v):d.toISOString();};
const dte=v=>{if(!v)return'';if(v instanceof Date&&!isNaN(v))return dateKey(v);const s=txt(v);const m=s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);if(m)return `${m[1]}-${m[2].padStart(2,'0')}-${m[3].padStart(2,'0')}`;const d=new Date(s);return isNaN(d)?s:dateKey(d);};
function get(r,...names){const keys=Object.fromEntries(Object.keys(r).map(k=>[k.trim().toLowerCase(),k]));for(const x of names){const k=keys[String(x).trim().toLowerCase()];if(k!==undefined)return r[k];}return'';}
function findSheet(wb,names){const m=Object.fromEntries(wb.SheetNames.map(x=>[x.trim().toLowerCase(),x]));for(const n of names)if(m[n.toLowerCase()])return m[n.toLowerCase()];return'';}
function rows(wb,name){return name?window.XLSX.utils.sheet_to_json(wb.Sheets[name],{defval:'',raw:true,blankrows:false}):[];}
function key(sec,r){switch(sec){case'users':return clean(r.username).toLowerCase()||clean(r.pharmacistName).toLowerCase();case'skills':return clean(r.username).toLowerCase()+'|'+clean(r.skill).toUpperCase();case'shifts':case'staffingRequirements':return clean(r.shift).toUpperCase();case'requests':return clean(r.recordId)||[r.recordType,r.username,r.startDate||r.date,r.endDate].map(clean).join('|');case'weeklyAvailability':return clean(r.id)||[r.username,r.day,r.effectiveStart,r.effectiveEnd].map(clean).join('|');case'preceptorCalendar':return clean(r.id)||clean(r.username)+'|'+clean(r.date);case'schedule':return clean(r.assignmentId)||[r.generationId,r.date,r.shift,r.slot].map(clean).join('|');default:return clean(r.id)||JSON.stringify(r);}}
function merge(existing,incoming,sec){const out=[...(existing||[])],m=new Map(out.map(x=>[key(sec,x),x]));for(const x of incoming){const k=key(sec,x);if(k&&m.has(k))Object.assign(m.get(k),x);else{out.push(x);if(k)m.set(k,x);}}return out;}

function mapUsers(a,current){
  const old=new Map((current||[]).map(x=>[clean(x.username).toLowerCase(),x]));
  return a.map(r=>{
    const username=txt(get(r,'Username')), prev=old.get(username.toLowerCase())||{};
    return {
      ...prev,
      id:txt(get(r,'Employee ID'))||prev.id||uid('EMP'),
      pharmacistName:txt(get(r,'Pharmacist Name','Name')),
      name:txt(get(r,'Pharmacist Name','Name')),
      username,
      role:'Pharmacist',
      active:bool(get(r,'Active'),true),
      scheduleType:txt(get(r,'Schedule Type'))||'Regular',
      preceptor:bool(get(r,'Preceptor'),false),
      resident:bool(get(r,'Resident'),false),
      weekendGroup:txt(get(r,'Weekend Group')).toUpperCase(),
      weekendRotationAnchorDate:dte(get(r,'Weekend Rotation Anchor Date')),
      weeklyHourMaximum:n(get(r,'Weekly Hour Maximum'),40),
      targetWeeklyHours:n(get(r,'Target Weekly Hours'),40),
      maximumEveningShiftsPerMonth:n(get(r,'Maximum Evening Shifts Per Month'),7),
      preferredStartTime:txt(get(r,'Preferred Start Time')),
      preferredEndTime:txt(get(r,'Preferred End Time')),
      customHoursEnabled:bool(get(r,'Custom Hours Enabled'),false),
      customHoursMode:txt(get(r,'Custom Hours Mode'))||'SOFT',
      preferredShiftType:txt(get(r,'Preferred Shift Type')),
      offDayCoverageSkills:list(get(r,'Off-Day Coverage Skills')),
      weekendEligible:bool(get(r,'Weekend Eligible'),true),
      eveningEligible:bool(get(r,'Evening Eligible'),true),
      nightEligible:bool(get(r,'Night Eligible'),true),
      residentCoversRegular:bool(get(r,'Resident Covers Regular'),false),
      residentWeekends:bool(get(r,'Resident Weekends'),true),
      residentEvenings:bool(get(r,'Resident Evenings'),true),
      residentNights:bool(get(r,'Resident Nights'),false),
      rotationAnchorDate:dte(get(r,'Rotation Anchor Date')),
      sevenOnWeeklyHandling:txt(get(r,'SevenOn Weekly Handling'))||'ENFORCE_MAX',
      notes:txt(get(r,'Notes')),
      updatedAt:new Date().toISOString(),
      updatedBy:'WORKBOOK IMPORT'
    };
  }).filter(x=>x.username||x.pharmacistName);
}
const mapSkills=a=>a.map(r=>({id:txt(get(r,'Skill ID'))||uid('SKL'),employeeId:txt(get(r,'Employee ID')),pharmacistName:txt(get(r,'Pharmacist Name')),username:txt(get(r,'Username')),skill:txt(get(r,'Skill')).toUpperCase(),active:bool(get(r,'Active'),true),updatedAt:new Date().toISOString(),updatedBy:'WORKBOOK IMPORT'})).filter(x=>x.username&&x.skill);
const mapShifts=a=>a.map(r=>({shift:txt(get(r,'Shift')).toUpperCase(),meaning:txt(get(r,'Meaning')),start:txt(get(r,'Start')),end:txt(get(r,'End')),hours:n(get(r,'Hours')),creditedHours:n(get(r,'Credited Hours'),n(get(r,'Hours'))),type:txt(get(r,'Type'))||'Day',skill:txt(get(r,'Skill')).toUpperCase(),weekend:bool(get(r,'Weekend')),active:bool(get(r,'Active'),true),priority:n(get(r,'Priority'),50),updatedAt:new Date().toISOString(),updatedBy:'WORKBOOK IMPORT'})).filter(x=>x.shift);
const mapStaff=a=>a.map(r=>({id:uid('REQ'),shift:txt(get(r,'Shift')).toUpperCase(),Sunday:n(get(r,'Sunday')),Monday:n(get(r,'Monday')),Tuesday:n(get(r,'Tuesday')),Wednesday:n(get(r,'Wednesday')),Thursday:n(get(r,'Thursday')),Friday:n(get(r,'Friday')),Saturday:n(get(r,'Saturday')),active:bool(get(r,'Active'),true),updatedAt:new Date().toISOString(),updatedBy:'WORKBOOK IMPORT'})).filter(x=>x.shift);
const mapReq=a=>a.map(r=>({recordType:txt(get(r,'Record Type'))||'PTO',recordId:txt(get(r,'Record ID'))||uid('REQ'),pharmacist:txt(get(r,'Pharmacist','Pharmacist Name')),username:txt(get(r,'Username')),date:dte(get(r,'Date')),startDate:dte(get(r,'Start Date')),endDate:dte(get(r,'End Date')),weekendSaturday:dte(get(r,'Weekend Saturday')),weekendSunday:dte(get(r,'Weekend Sunday')),available:get(r,'Available')===''?'':bool(get(r,'Available')),status:txt(get(r,'Status'))||'Pending',comment:txt(get(r,'Comment')),submittedAt:iso(get(r,'Submitted At'))||new Date().toISOString(),reviewedBy:txt(get(r,'Reviewed By')),reviewedAt:iso(get(r,'Reviewed At')),updatedAt:new Date().toISOString(),updatedBy:'WORKBOOK IMPORT'}));
const mapWeekly=a=>a.map(r=>({id:txt(get(r,'Availability ID'))||uid('AVA'),employeeId:txt(get(r,'Employee ID')),pharmacistName:txt(get(r,'Pharmacist Name')),username:txt(get(r,'Username')),day:txt(get(r,'Day')).toUpperCase(),available:bool(get(r,'Available'),true),startTime:txt(get(r,'Start Time')),endTime:txt(get(r,'End Time')),ruleType:txt(get(r,'Rule Type'))||'AVAILABLE',effectiveStart:dte(get(r,'Effective Start')),effectiveEnd:dte(get(r,'Effective End')),notes:txt(get(r,'Notes')),active:bool(get(r,'Active'),true),updatedAt:new Date().toISOString(),updatedBy:'WORKBOOK IMPORT'})).filter(x=>x.username&&x.day);
const mapPrec=a=>a.map(r=>({id:txt(get(r,'ID','Record ID'))||uid('PRE'),username:txt(get(r,'Username')),pharmacistName:txt(get(r,'Pharmacist Name','Pharmacist')),date:dte(get(r,'Date')),precepting:bool(get(r,'Precepting'),txt(get(r,'Status')).toUpperCase()==='ON'),status:txt(get(r,'Status'))||(bool(get(r,'Precepting'))?'ON':'OFF'),rotation:txt(get(r,'Rotation')),notes:txt(get(r,'Notes')),updatedAt:new Date().toISOString(),updatedBy:'WORKBOOK IMPORT'})).filter(x=>x.username&&x.date);
const mapSched=a=>a.map(r=>({generationId:txt(get(r,'Generation ID')),assignmentId:txt(get(r,'Assignment ID'))||uid('ASN'),date:dte(get(r,'Date')),day:txt(get(r,'Day')),shift:txt(get(r,'Shift')).toUpperCase(),slot:n(get(r,'Slot'),1),assignedPharmacist:txt(get(r,'Assigned Pharmacist')),username:txt(get(r,'Username')),hours:n(get(r,'Hours')),creditedHours:n(get(r,'Credited Hours'),n(get(r,'Hours'))),requiredSkill:txt(get(r,'Required Skill')).toUpperCase(),coverageForPharmacist:txt(get(r,'Coverage For Pharmacist')),coverageForUsername:txt(get(r,'Coverage For Username')),coverageReason:txt(get(r,'Coverage Reason')),shiftType:txt(get(r,'Shift Type')),weekend:bool(get(r,'Weekend')),weekendGroup:txt(get(r,'Weekend Group')),holiday:txt(get(r,'Holiday')),locked:bool(get(r,'Locked')),manual:bool(get(r,'Manual')),status:txt(get(r,'Status'))||'ASSIGNED',warning:txt(get(r,'Warning')),updatedAt:new Date().toISOString(),updatedBy:'WORKBOOK IMPORT',finalizedAt:iso(get(r,'Finalized At'))})).filter(x=>x.date&&x.shift);

function mapSettings(a,current){
  const s={...current},m={};for(const r of a)m[txt(get(r,'Setting'))]=get(r,'Value');
  const set=(p,l,f)=>{if(Object.prototype.hasOwnProperty.call(m,l))s[p]=f(m[l]);};
  set('appName','Schedule Title',v=>txt(v)||'NeoChrono'); set('scheduleDays','Default Schedule Days',v=>n(v,56));
  set('weekStart','Week Start',v=>txt(v)||'Sunday'); set('weeklyHoursLimit','Weekly Hours Limit',v=>n(v,40));
  set('requiredHoursPerSchedulePeriod','Required Hours Per Schedule Period',v=>n(v,320)); set('maximumHoursPerSchedulePeriod','Maximum Hours Per Schedule Period',v=>n(v,320));
  set('regularWorkdaysPerWeek','Regular Workdays Per Week',v=>n(v,5)); set('maximumConsecutiveWorkdays','Maximum Consecutive Workdays',v=>n(v,5));
  set('maximumEveningShifts','Maximum Evening Shifts',v=>n(v,7)); set('weekendRotation','Weekend Rotation',v=>list(v||'A,B,C'));
  set('weekendAnchorDate','Weekend Anchor Date',dte); set('weekendAnchorGroup','Weekend Anchor Group',v=>txt(v)||'A');
  set('allowWeekendFallback','Allow Weekend Fallback',v=>bool(v)); set('requiredWeekendAssignment','Required Weekend Assignment',v=>bool(v,true));
  set('residentE2ShiftCode','Resident E2 Shift Code',v=>txt(v)||'E2'); set('residentE2ShiftsPerWeek','Resident E2 Shifts Per Week',v=>n(v,1));
  set('residentsRequiredAssignedWeekend','Residents Required Assigned Weekend',v=>bool(v,true)); set('sevenOnBlockScheduling','SevenOn Block Scheduling',v=>bool(v,true));
  set('blockEveningToMorningTransition','Block Evening To Morning Transition',v=>bool(v,true)); set('ptoAutoApprovalLimitPerDate','PTO Auto-Approve Per Day',v=>n(v,2));
  set('allowAdminRuleOverride','Allow Admin Rule Override',v=>bool(v,true)); set('showFullScheduleToPharmacists','Full Schedule Visible To Employees',v=>bool(v,true));
  set('eddEdePairRequired','Require Paired ED Coverage',v=>bool(v,true)); set('edDayShiftCode','ED Day Shift Code',v=>txt(v)||'EDD'); set('edEveningShiftCode','ED Evening Shift Code',v=>txt(v)||'EDE');
  return s;
}

function periods(schedule){
  const by={};for(const r of schedule){if(r.generationId)(by[r.generationId]??=[]).push(r);}
  return Object.entries(by).map(([generationId,a])=>{const ds=a.map(x=>x.date).filter(Boolean).sort(),f=a.find(x=>x.finalizedAt);return{generationId,startDate:ds[0]||'',endDate:ds.at(-1)||'',status:f?'FINALIZED':'DRAFT',createdBy:'WORKBOOK_IMPORT',createdAt:new Date().toISOString(),finalizedBy:'',finalizedAt:f?.finalizedAt||''};});
}

export async function parseNeoChronoWorkbook(file){
  if(!window.XLSX)throw new Error('Excel reader did not load. Refresh and try again.');
  const wb=window.XLSX.read(await file.arrayBuffer(),{type:'array',cellDates:true});
  const tables={},recognized=[],seen=new Set();
  for(const [key,names] of Object.entries(ALIASES)){const name=findSheet(wb,names);tables[key]=rows(wb,name);if(name){recognized.push({key,sheet:name,rows:tables[key].length});seen.add(name);}}
  return {fileName:file.name,tables,recognized,unrecognized:wb.SheetNames.filter(x=>!seen.has(x))};
}

export function applyNeoChronoWorkbookImport(db,parsed,{mode='replace',includeHistory=false}={}){
  const t=parsed.tables||{}, mapped={
    users:mapUsers(t.users||[],db.users),skills:mapSkills(t.skills||[]),shifts:mapShifts(t.shifts||[]),staffingRequirements:mapStaff(t.staffingRequirements||[]),
    requests:mapReq(t.requests||[]),weeklyAvailability:mapWeekly(t.weeklyAvailability||[]),preceptorCalendar:mapPrec(t.preceptorCalendar||[]),schedule:mapSched(t.schedule||[])
  },counts={};
  for(const k of SETUP){if(!(t[k]||[]).length){counts[k]=0;continue;}db[k]=mode==='upsert'?merge(db[k]||[],mapped[k],k):mapped[k];counts[k]=mapped[k].length;}
  if((t.settingsRows||[]).length){db.settings=mapSettings(t.settingsRows,db.settings);counts.settings=t.settingsRows.length;}else counts.settings=0;
  if(includeHistory&&(t.schedule||[]).length){db.schedule=mode==='upsert'?merge(db.schedule||[],mapped.schedule,'schedule'):mapped.schedule;db.schedulePeriods=periods(db.schedule);counts.schedule=mapped.schedule.length;}
  db.meta=db.meta||{};db.meta.lastWorkbookImportAt=new Date().toISOString();db.meta.lastWorkbookImportFile=parsed.fileName;db.meta.lastWorkbookImportMode=mode;
  return {counts,recognized:parsed.recognized,unrecognized:parsed.unrecognized,mode,includeHistory};
}
