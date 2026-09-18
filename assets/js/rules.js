import { addDays, clean, dateKey, dayName, daysBetween, eachDate, isWeekend, monthKey, num, parseDate, weekStartSunday, yes } from './utils.js';

export function settingsMap(rows){const out={};(rows||[]).forEach(r=>out[clean(r.Setting)]=r.Value);return out;}
export function shiftMap(rows){const out={};(rows||[]).forEach(r=>{if(clean(r.Shift))out[clean(r.Shift).toUpperCase()]=r;});return out;}
export function userMaps(users){const byId={},byName={},byUsername={};(users||[]).forEach(u=>{if(u['Employee ID'])byId[clean(u['Employee ID'])]=u;if(u['Pharmacist Name'])byName[clean(u['Pharmacist Name'])]=u;if(u.Username)byUsername[clean(u.Username)]=u;});return {byId,byName,byUsername};}
export function skillsByUser(skills,users){const maps=userMaps(users),out={};(skills||[]).filter(r=>yes(r.Active??'Yes')).forEach(r=>{const key=clean(r.Username)||clean(maps.byId[clean(r['Employee ID'])]?.Username);if(!key)return;(out[key]||(out[key]=new Set())).add(clean(r.Skill).toUpperCase());});return out;}
export function creditedHours(shift){const c=Number(shift?.['Credited Hours']);return Number.isFinite(c)&&c>0?c:num(shift?.Hours,0);}
export function shiftType(shift){return clean(shift?.Type).toLowerCase();}
export function isEveningShift(shift){return shiftType(shift)==='evening';}
export function isDayShift(shift){return ['day','morning'].includes(shiftType(shift));}
export function isNightShift(shift){return shiftType(shift)==='night';}
export function isRegularUser(u){return clean(u?.['Schedule Type']).toUpperCase()==='REGULAR';}
export function isPrnUser(u){return clean(u?.['Schedule Type']).toUpperCase()==='PRN';}
export function isSevenUser(u){return clean(u?.['Schedule Type']).toUpperCase().replace(/[^A-Z0-9]/g,'')==='7ON7OFF';}

export function weekendGroupForDate(date,settings){
  const seq=clean(settings['Weekend Rotation']||'A,B,C').split(',').map(x=>x.trim().toUpperCase()).filter(Boolean); if(!seq.length)return '';
  const anchor=parseDate(settings['Weekend Anchor Date']); if(!anchor)return '';
  const anchorGroup=clean(settings['Weekend Anchor Group']||seq[0]).toUpperCase(); const anchorIdx=Math.max(0,seq.indexOf(anchorGroup));
  const d=parseDate(date); const sat=addDays(d,(6-d.getDay()+7)%7); const weeks=Math.floor(daysBetween(anchor,sat)/7); const idx=((anchorIdx+weeks)%seq.length+seq.length)%seq.length; return seq[idx];
}
export function isSevenOnDate(user,date){
  const anchor=parseDate(user?.['Rotation Anchor Date']); if(!anchor)return false; const diff=daysBetween(anchor,date); const mod=((diff%14)+14)%14; return mod<7;
}
export function preceptorModeForDate(user,date,calendar){
  if(!yes(user?.Preceptor))return null; const uid=clean(user.Username), id=clean(user['Employee ID']), dk=dateKey(date);
  const rec=(calendar||[]).filter(r=>yes(r.Active??'Yes')).find(r=> (clean(r.Username)===uid||clean(r['Employee ID'])===id||clean(r.Pharmacist)===clean(user['Pharmacist Name'])) && dateKey(r.Date)===dk);
  return rec?clean(rec.Mode||rec.Precepting||'ON').toUpperCase():'ON';
}
export function approvedPtoRanges(requests){
  const out={}; (requests||[]).forEach(r=>{if(clean(r['Record Type']).toUpperCase()!=='PTO'||clean(r.Status).toUpperCase()!=='APPROVED')return;const key=clean(r.Username)||clean(r.Pharmacist);const s=parseDate(r['Start Date']||r.Date),e=parseDate(r['End Date']||r.Date||r['Start Date']);if(!key||!s||!e)return;(out[key]||(out[key]=[])).push([s,e]);});return out;
}
export function isOnPto(user,date,pto){const keys=[clean(user.Username),clean(user['Pharmacist Name'])];return keys.some(k=>(pto[k]||[]).some(([s,e])=>parseDate(date)>=s&&parseDate(date)<=e));}
export function recurringAvailabilityStatus(user,date,shift,weeklyRows){
  const day=dayName(date); const rows=(weeklyRows||[]).filter(r=>yes(r.Active??'Yes') && (clean(r.Username)===clean(user.Username)||clean(r['Employee ID'])===clean(user['Employee ID'])) && clean(r.Day).toUpperCase()===day);
  if(!rows.length)return {ok:true};
  if(rows.some(r=>!yes(r.Available)))return {ok:false,reason:'WEEKLY_AVAILABILITY_DAY'};
  const timed=rows.filter(r=>clean(r['Start Time'])||clean(r['End Time'])); if(!timed.length)return {ok:true};
  const start=clean(shift?.Start),end=clean(shift?.End); if(!start||!end)return {ok:true};
  const ok=timed.some(r=>(!clean(r['Start Time'])||start>=clean(r['Start Time']))&&(!clean(r['End Time'])||end<=clean(r['End Time']))); return ok?{ok:true}:{ok:false,reason:'WEEKLY_AVAILABILITY_TIME'};
}
export function dateAvailabilityStatus(user,date,requests){
  const dk=dateKey(date); const rows=(requests||[]).filter(r=>clean(r['Record Type']).toUpperCase()!=='PTO' && clean(r.Status).toUpperCase()!=='REJECTED' && (clean(r.Username)===clean(user.Username)||clean(r.Pharmacist)===clean(user['Pharmacist Name'])));
  for(const r of rows){const s=dateKey(r['Start Date']||r.Date),e=dateKey(r['End Date']||r.Date||r['Start Date']);if(s&&e&&dk>=s&&dk<=e&&clean(r.Available)&&!yes(r.Available))return {ok:false,reason:'UNAVAILABLE'};}
  return {ok:true};
}

export function ptoDatesForRequest(r){const s=parseDate(r['Start Date']||r.Date),e=parseDate(r['End Date']||r.Date||r['Start Date']);return s&&e?eachDate(s,e).map(dateKey):[];}
export function reconcilePtoAutoApprovals(requests,limit=2){
  const out=structuredClone(requests||[]); const pto=out.map((r,index)=>({r,index,id:clean(r['Record ID']),dates:ptoDatesForRequest(r),submitted:new Date(r['Submitted At']||0).getTime()||0})).filter(x=>clean(x.r['Record Type']).toUpperCase()==='PTO'&&clean(x.r.Status).toUpperCase()!=='REJECTED'&&x.dates.length);
  const byDate={};pto.forEach(x=>x.dates.forEach(d=>(byDate[d]||(byDate[d]=[])).push(x)));
  const winners={};Object.entries(byDate).forEach(([d,items])=>items.sort((a,b)=>a.submitted-b.submitted||a.index-b.index).forEach((x,i)=>{(winners[x.id]||(winners[x.id]={}))[d]=i<limit;}));
  pto.forEach(x=>{const eligible=x.dates.every(d=>winners[x.id]?.[d]);const status=clean(x.r.Status).toUpperCase();const systemAuto=clean(x.r['Reviewed By']).toUpperCase()==='SYSTEM AUTO-APPROVAL';if(eligible&&status!=='APPROVED'){x.r.Status='Approved';x.r['Reviewed By']='SYSTEM AUTO-APPROVAL';x.r['Reviewed At']=new Date().toISOString();}else if(!eligible&&status==='APPROVED'&&systemAuto){x.r.Status='Pending';x.r['Reviewed By']='';x.r['Reviewed At']='';}});
  return out;
}

export function buildModel(data){
  const settings=settingsMap(data.settings),shifts=shiftMap(data.shifts),maps=userMaps(data.users),skills=skillsByUser(data.skills,data.users),pto=approvedPtoRanges(data.requests);
  return {data,settings,shifts,...maps,skills,pto};
}

export function validateConfiguration(data){
  const model=buildModel(data),errors=[],warnings=[]; const seq=clean(model.settings['Weekend Rotation']||'A,B,C').split(',').map(x=>x.trim()).filter(Boolean);
  const activeUsers=(data.users||[]).filter(u=>yes(u.Active)); const activeShifts=(data.shifts||[]).filter(s=>yes(s.Active)); const required=(data.staffing||[]).filter(r=>yes(r.Active??'Yes'));
  required.forEach(r=>{const sh=model.shifts[clean(r.Shift).toUpperCase()];const total=['Sunday','Monday','Tuesday','Wednesday','Thursday','Friday','Saturday'].reduce((s,d)=>s+num(r[d],0),0);if(total&&!sh)errors.push(`Staffing requirement references unknown shift ${clean(r.Shift)}.`);if(total&&sh&&!yes(sh.Active))errors.push(`Required shift ${clean(r.Shift)} is inactive.`);if(total&&sh&&!creditedHours(sh))errors.push(`Required shift ${clean(r.Shift)} has no credited hours.`);if(total&&sh&&clean(sh.Skill)){const qualified=activeUsers.filter(u=>model.skills[clean(u.Username)]?.has(clean(sh.Skill).toUpperCase()));if(!qualified.length)errors.push(`Shift ${sh.Shift} requires ${sh.Skill}, but no active pharmacist has that skill.`);}});
  activeUsers.forEach(u=>{if(isRegularUser(u)&&!num(u['Weekly Hour Maximum'],0))errors.push(`${u['Pharmacist Name']} is missing Weekly Hour Maximum.`);if(clean(u['Weekend Group'])&&!seq.includes(clean(u['Weekend Group'])))errors.push(`${u['Pharmacist Name']} has invalid Weekend Group ${u['Weekend Group']}.`);if(isSevenUser(u)&&!parseDate(u['Rotation Anchor Date']))errors.push(`${u['Pharmacist Name']} is 7-on/7-off but has no Rotation Anchor Date.`);if(isSevenUser(u)&&!clean(u['Preferred Shift Type']))errors.push(`${u['Pharmacist Name']} is 7-on/7-off but has no dedicated Preferred Shift Type.`);});
  activeShifts.forEach(s=>{if((clean(s.Type)==='Day'||clean(s.Type)==='Evening'||clean(s.Type)==='Night')&&(!clean(s.Start)||!clean(s.End)))warnings.push(`Shift ${s.Shift} has no complete start/end time.`);});
  if(clean(model.settings['Week Start']||'Sunday').toLowerCase()!=='sunday')errors.push('Week Start must be Sunday for the exact five-day rule.');
  if(num(model.settings['Regular Workdays Per Week'],5)!==5)errors.push('Regular Workdays Per Week must be 5.');
  if(num(model.settings['Maximum Consecutive Workdays'],5)!==5)warnings.push('Maximum Consecutive Workdays is not 5.');
  return {ok:errors.length===0,errors,warnings};
}

export function scheduleHealth(schedule){const rows=schedule||[],required=rows.length,unfilled=rows.filter(r=>clean(r.Status).toUpperCase()==='UNFILLED'||clean(r['Assigned Pharmacist']).toUpperCase()==='UNFILLED'||!clean(r['Assigned Pharmacist'])).length;const warnings=rows.filter(r=>clean(r.Warning)).length;return {required,filled:required-unfilled,unfilled,coveragePercent:required?Math.round((required-unfilled)*1000/required)/10:100,ruleWarnings:warnings};}

export function employeeStats(schedule,users){const map={};(users||[]).forEach(u=>map[clean(u.Username)]={username:u.Username,name:u['Pharmacist Name'],totalHours:0,shifts:0,dayShifts:0,eveningShifts:0,nightShifts:0,weekendShifts:0,weeklyHours:{},weekendGroup:u['Weekend Group'],scheduleType:u['Schedule Type'],preceptor:u.Preceptor,resident:u.Resident});(schedule||[]).forEach(r=>{const k=clean(r.Username),x=map[k];if(!x||clean(r.Status).toUpperCase()==='UNFILLED')return;const h=num(r['Credited Hours'],0);x.totalHours+=h;x.shifts++;const t=clean(r['Shift Type']).toLowerCase();if(t==='day')x.dayShifts++;if(t==='evening')x.eveningShifts++;if(t==='night')x.nightShifts++;if(yes(r.Weekend))x.weekendShifts++;const wk=dateKey(weekStartSunday(r.Date));x.weeklyHours[wk]=num(x.weeklyHours[wk],0)+h;});return Object.values(map);}
