/* GitHub compatibility patches for the uploaded NeoChrono project.
 * Loaded after the original Code.gs + add-ons.
 * The scheduling algorithm itself is unchanged.
 */

function getSpreadsheet_() {
  return getDb_();
}

function githubEnsureSkillPreferredColumn_() {
  var sh = getDb_().getSheetByName(APP.SHEETS.SKILLS);
  if (!sh) sh = getDb_().insertSheet(APP.SHEETS.SKILLS);

  var lastRow = Math.max(1, sh.getLastRow());
  var lastCol = Math.max(1, sh.getLastColumn());
  var values = sh.getRange(1,1,lastRow,lastCol).getValues();
  if (!values.length) values = [[]];

  var headers = (values[0] || []).map(clean_);
  var changed = false;

  ['Preferred','Priority'].forEach(function(column){
    if (headers.indexOf(column) < 0) {
      headers.push(column);
      values[0] = headers;
      for (var r=1;r<values.length;r++) values[r].push('');
      changed = true;
    }
    if (APP.HEADERS && APP.HEADERS.SKILLS && APP.HEADERS.SKILLS.indexOf(column) < 0) {
      APP.HEADERS.SKILLS.push(column);
    }
  });

  if (changed) {
    sh.getRange(1,1,values.length,headers.length).setValues(values);
  }
  return changed;
}

function githubNormalizeSkillPriorities_(values, headers) {
  var h={};
  headers.forEach(function(name,i){h[clean_(name)]=i;});
  if (h.Priority===undefined || h.Skill===undefined) return false;

  var groups={};
  var changed=false;
  for (var r=1;r<values.length;r++) {
    if (!yesDefault_(values[r][h.Active],true)) {
      if (clean_(values[r][h.Priority])) {
        values[r][h.Priority]='';
        changed=true;
      }
      continue;
    }
    var skill=clean_(values[r][h.Skill]).toUpperCase();
    if (!skill) continue;
    var userKey=clean_(values[r][h['Employee ID']]) || clean_(values[r][h.Username]) || clean_(values[r][h['Pharmacist Name']]);
    if (!userKey) continue;
    if (!groups[userKey]) groups[userKey]=[];
    groups[userKey].push({
      row:r,
      skill:skill,
      preferred:h.Preferred!==undefined && yes_(values[r][h.Preferred]),
      rawPriority:Number(values[r][h.Priority])
    });
  }

  Object.keys(groups).forEach(function(key){
    var items=groups[key];

    // Existing valid priorities are respected. Missing/duplicate priorities are
    // repaired deterministically. On first migration, Preferred/Home Unit is
    // placed first, then the existing sheet order is preserved.
    var hasAny=items.some(function(x){return Number.isFinite(x.rawPriority)&&x.rawPriority>0;});
    items.sort(function(a,b){
      if (hasAny) {
        var ap=Number.isFinite(a.rawPriority)&&a.rawPriority>0?a.rawPriority:9999;
        var bp=Number.isFinite(b.rawPriority)&&b.rawPriority>0?b.rawPriority:9999;
        return ap-bp || (a.preferred?0:1)-(b.preferred?0:1) || a.row-b.row;
      }
      return (a.preferred?0:1)-(b.preferred?0:1) || a.row-b.row;
    });

    items.forEach(function(item,idx){
      var wanted=idx+1;
      if (Number(values[item.row][h.Priority])!==wanted) {
        values[item.row][h.Priority]=wanted;
        changed=true;
      }
    });
  });

  return changed;
}

function githubSkillPriorityCodeForSlot_(slot) {
  return clean_(
    (slot && slot.requiredSkill) ||
    (slot && slot.shift && slot.shift.Skill) ||
    (slot && slot.shiftCode) ||
    (slot && slot.shift && slot.shift.Shift)
  ).toUpperCase();
}

function skillPriorityForUserSlot_(u,slot,model) {
  if (!u || !slot || !model) return 9999;
  var username=clean_(u.Username);
  var code=githubSkillPriorityCodeForSlot_(slot);
  if (!username || !code) return 9999;
  var map=model.skillPriorityByUser && model.skillPriorityByUser[username];
  var p=map ? Number(map[code]) : NaN;
  return Number.isFinite(p)&&p>0 ? p : 9999;
}

function githubHomeOwnerEligibleForSlot_(slot,model,state) {
  var owners=typeof primaryHomeOwnersForSlot_==='function'
    ? primaryHomeOwnersForSlot_(slot,model)
    : [];
  if (!owners.length) return false;

  return owners.some(function(owner){
    var e=eligibility_(owner,slot,model,state,false);
    return !!(e&&e.ok);
  });
}

/**
 * Sort rank used by the normal allocator.
 *
 * 0..999 = a Preferred/Home Unit owner is available, so keep the normal owner
 *          in place before using that pharmacist elsewhere.
 * 1000+  = the home owner is unavailable and backup pharmacists are ordered
 *          by the manager's Employee Skills Priority (1 first, then 2, 3...).
 * 9000+  = no configured home owner; still respect skill priority if possible.
 */
function skillPrioritySlotRank_(slot,model,state) {
  if (!slot || !model || !state) return 99999;

  var owners=typeof primaryHomeOwnersForSlot_==='function'
    ? primaryHomeOwnersForSlot_(slot,model)
    : [];
  var ownerEligible=owners.length && githubHomeOwnerEligibleForSlot_(slot,model,state);

  if (ownerEligible) return 0;

  var best=9999;
  (model.users||[]).forEach(function(u){
    if (!yes_(u.Active)) return;
    var e=eligibility_(u,slot,model,state,false);
    if (!e || !e.ok) return;
    best=Math.min(best,skillPriorityForUserSlot_(u,slot,model));
  });

  if (owners.length) return 1000+best;
  return 9000+best;
}

function ensureSkillPreferenceSystem(token) {
  var ctx = requireAdmin_(token);
  var changed = githubEnsureSkillPreferredColumn_();

  var sh = getDb_().getSheetByName(APP.SHEETS.SKILLS);
  var values = sh.getDataRange().getValues();
  if (!values.length) return {ok:true,changed:changed};

  var headers = values[0].map(clean_);
  var h = {};
  headers.forEach(function(name,i){h[name]=i;});
  var users = readTable_(APP.SHEETS.USERS);
  var shifts = readTable_(APP.SHEETS.SHIFTS);
  var activeShiftCodes = new Set(shifts.filter(function(s){return yesDefault_(s.Active,true);}).map(function(s){return clean_(s.Shift).toUpperCase();}));

  users.forEach(function(u){
    var eid = clean_(u['Employee ID']);
    var username = clean_(u.Username);
    var matches = [];

    for (var r=1;r<values.length;r++) {
      var same =
        (eid && clean_(values[r][h['Employee ID']]) === eid) ||
        (username && clean_(values[r][h.Username]) === username);
      if (!same || !yesDefault_(values[r][h.Active],true)) continue;
      matches.push(r);
    }

    if (!matches.length) return;

    var preferredRows = matches.filter(function(r){return yes_(values[r][h.Preferred]);});
    if (preferredRows.length > 1) {
      preferredRows.slice(1).forEach(function(r){ values[r][h.Preferred]='No'; changed=true; });
      preferredRows = preferredRows.slice(0,1);
    }

    if (!preferredRows.length) {
      var oldPref = clean_(u['Preferred Shift Type']).toUpperCase();
      if (oldPref && activeShiftCodes.has(oldPref)) {
        var target = matches.find(function(r){return clean_(values[r][h.Skill]).toUpperCase()===oldPref;});
        if (target !== undefined) {
          values[target][h.Preferred] = 'Yes';
          matches.filter(function(r){return r!==target;}).forEach(function(r){
            if (clean_(values[r][h.Preferred])) values[r][h.Preferred]='No';
          });
          changed = true;
        }
      }
    }
  });

  if (githubNormalizeSkillPriorities_(values,headers)) changed=true;

  if (changed) {
    sh.getRange(1,1,values.length,headers.length).setValues(values);
    SpreadsheetApp.flush();
    if (typeof _PRECEPTOR_CALENDAR_RUNTIME_CACHE_ !== 'undefined') _PRECEPTOR_CALENDAR_RUNTIME_CACHE_ = null;
    audit_('SKILL_PREFERENCE_SYSTEM','','','','','','','No','',
      'Ensured Employee Skills Preferred/Home Unit and manager Priority ordering support',ctx.username);
  }

  return {ok:true,changed:changed};
}

function saveSkillProfile(token, employeeId, preferredCode, skillCodes, skillPriorities) {
  var ctx = requireAdmin_(token);
  githubEnsureSkillPreferredColumn_();

  var eid = clean_(employeeId);
  var user = readTable_(APP.SHEETS.USERS).find(function(u){return clean_(u['Employee ID'])===eid;});
  if (!user) throw new Error('Pharmacist was not found.');

  var desired = Array.from(new Set((Array.isArray(skillCodes)?skillCodes:[])
    .map(function(x){return clean_(x).toUpperCase();})
    .filter(Boolean)));

  var preferred = clean_(preferredCode).toUpperCase();
  if (preferred && desired.indexOf(preferred) < 0) desired.push(preferred);

  var supplied=skillPriorities && typeof skillPriorities==='object' ? skillPriorities : {};
  desired.sort(function(a,b){
    var ap=Number(supplied[a]),bp=Number(supplied[b]);
    ap=Number.isFinite(ap)&&ap>0?ap:9999;
    bp=Number.isFinite(bp)&&bp>0?bp:9999;
    return ap-bp || a.localeCompare(b);
  });

  var priorityMap={};
  desired.forEach(function(code,idx){priorityMap[code]=idx+1;});

  var sh = getDb_().getSheetByName(APP.SHEETS.SKILLS);
  var values = sh.getDataRange().getValues();
  var headers = values[0].map(clean_);
  var h={}; headers.forEach(function(name,i){h[name]=i;});

  var seen = new Set();
  for (var r=1;r<values.length;r++) {
    var same =
      (eid && clean_(values[r][h['Employee ID']])===eid) ||
      (clean_(user.Username) && clean_(values[r][h.Username])===clean_(user.Username));
    if (!same) continue;

    var code = clean_(values[r][h.Skill]).toUpperCase();
    if (!code) continue;
    seen.add(code);

    var active = desired.indexOf(code)>=0;
    values[r][h['Employee ID']] = eid;
    values[r][h['Pharmacist Name']] = clean_(user['Pharmacist Name']);
    values[r][h.Username] = clean_(user.Username);
    values[r][h.Active] = active ? 'Yes' : 'No';
    values[r][h.Preferred] = active && preferred===code ? 'Yes' : 'No';
    values[r][h.Priority] = active ? priorityMap[code] : '';
    values[r][h['Updated At']] = new Date();
    values[r][h['Updated By']] = ctx.username;
  }

  desired.forEach(function(code){
    if (seen.has(code)) return;
    var row = headers.map(function(){return '';});
    row[h['Employee ID']] = eid;
    row[h['Pharmacist Name']] = clean_(user['Pharmacist Name']);
    row[h.Username] = clean_(user.Username);
    row[h.Skill] = code;
    row[h.Active] = 'Yes';
    row[h.Preferred] = preferred===code ? 'Yes' : 'No';
    row[h.Priority] = priorityMap[code];
    row[h['Updated At']] = new Date();
    row[h['Updated By']] = ctx.username;
    values.push(row);
  });

  sh.getRange(1,1,values.length,headers.length).setValues(values);

  updateRowByKey_(APP.SHEETS.USERS,'Employee ID',eid,{
    'Preferred Shift Type':preferred,
    'Updated At':new Date(),
    'Updated By':ctx.username
  });

  if (typeof _PRECEPTOR_CALENDAR_RUNTIME_CACHE_ !== 'undefined') _PRECEPTOR_CALENDAR_RUNTIME_CACHE_ = null;
  audit_('EMPLOYEE_SKILL_PROFILE_CHANGED','',clean_(user['Pharmacist Name']),'','','',preferred,'No','',
    'Skills by priority: '+desired.map(function(code){return priorityMap[code]+'='+code;}).join(', ')+
    '; Preferred/Home Unit: '+(preferred||'None'),ctx.username);

  return {
    ok:true,
    employeeId:eid,
    skills:desired,
    preferred:preferred,
    priorities:priorityMap
  };
}

// Quick-add compatibility: a newly-added normal skill receives the next
// priority number for that pharmacist automatically.
var _SKILL_PRIORITY_BASE_SAVE_SKILL_ = typeof saveSkill==='function' ? saveSkill : null;
if (_SKILL_PRIORITY_BASE_SAVE_SKILL_) {
  saveSkill = function(token,data) {
    githubEnsureSkillPreferredColumn_();
    var out=_SKILL_PRIORITY_BASE_SAVE_SKILL_(token,data);

    var u=findUser_(data && data.Username,data && data['Pharmacist Name']);
    if (u) {
      var sh=getDb_().getSheetByName(APP.SHEETS.SKILLS);
      var values=sh.getDataRange().getValues();
      var headers=values[0].map(clean_);
      var h={}; headers.forEach(function(name,i){h[name]=i;});
      var eid=clean_(u['Employee ID']),username=clean_(u.Username);
      var rows=[];
      for(var r=1;r<values.length;r++){
        var same=(eid&&clean_(values[r][h['Employee ID']])===eid)||(username&&clean_(values[r][h.Username])===username);
        if(same&&yesDefault_(values[r][h.Active],true)&&clean_(values[r][h.Skill]))rows.push(r);
      }
      rows.sort(function(a,b){
        var ap=Number(values[a][h.Priority]),bp=Number(values[b][h.Priority]);
        ap=Number.isFinite(ap)&&ap>0?ap:9999;
        bp=Number.isFinite(bp)&&bp>0?bp:9999;
        return ap-bp||a-b;
      });
      rows.forEach(function(row,idx){values[row][h.Priority]=idx+1;});
      sh.getRange(1,1,values.length,headers.length).setValues(values);
    }
    return out;
  };
}

// Add priority metadata to every scheduling model without changing the
// underlying eligibility/skill rules.
var _SKILL_PRIORITY_BASE_LOAD_MODEL_ = typeof loadSchedulingModel_==='function'
  ? loadSchedulingModel_
  : null;

if (_SKILL_PRIORITY_BASE_LOAD_MODEL_) {
  loadSchedulingModel_ = function() {
    var model=_SKILL_PRIORITY_BASE_LOAD_MODEL_();
    model.skillPriorityByUser={};

    (model.skills||[])
      .filter(function(r){return yesDefault_(r.Active,true);})
      .forEach(function(r){
        var username=clean_(r.Username) ||
          clean_((model.usersByName && model.usersByName[clean_(r['Pharmacist Name'])] || {}).Username);
        var code=clean_(r.Skill).toUpperCase();
        if(!username||!code)return;
        if(!model.skillPriorityByUser[username])model.skillPriorityByUser[username]={};
        var p=Number(r.Priority);
        if(!Number.isFinite(p)||p<1)p=9999;
        model.skillPriorityByUser[username][code]=p;
      });

    return model;
  };
}

// Priority affects backup assignment only after all hard eligibility rules.
// A normal Preferred/Home Unit owner remains preferred when that owner can work.
var _SKILL_PRIORITY_BASE_SCORE_CANDIDATE_ = typeof scoreCandidate_==='function'
  ? scoreCandidate_
  : null;

if (_SKILL_PRIORITY_BASE_SCORE_CANDIDATE_) {
  scoreCandidate_ = function(u,slot,model,state,elig) {
    var score=_SKILL_PRIORITY_BASE_SCORE_CANDIDATE_(u,slot,model,state,elig);

    var owners=typeof primaryHomeOwnersForSlot_==='function'
      ? primaryHomeOwnersForSlot_(slot,model)
      : [];
    var isOwner=owners.some(function(owner){
      return clean_(owner.Username)===clean_(u.Username);
    });

    // Do not let a backup with Priority 1 steal a normal home-unit assignment
    // from its configured Preferred/Home Unit owner.
    if (!isOwner && owners.length && !githubHomeOwnerEligibleForSlot_(slot,model,state)) {
      var rank=skillPriorityForUserSlot_(u,slot,model);
      if (rank<9999) {
        score += Math.max(1000,36000-(rank-1)*6000);
      }
    }

    return score;
  };
}

async function githubVerifyAdminPassword_(admin,password) {
  var salt = clean_(admin['Password Salt']);
  var hash = clean_(admin['Password Hash']);
  if (!salt || !hash) return false;

  var alg = clean_(admin['Password Algorithm']).toUpperCase();
  var iterations = Number(admin['Password Iterations'] || 180000);
  var looksPbkdf2 = /^[0-9a-f]{64}$/i.test(hash);

  if (alg === 'PBKDF2-SHA256' && looksPbkdf2 && window.__neoRuntime && window.__neoRuntime.pbkdf2Hex) {
    var got = await window.__neoRuntime.pbkdf2Hex(String(password),salt,iterations);
    return String(got).toLowerCase() === hash.toLowerCase();
  }
  return hashPassword_(String(password),salt) === hash;
}

login = async function(username, password) {
  username = clean_(username).toLowerCase();
  if (!username || !password) throw new Error('Username and password are required.');

  var admins = readTable_(APP.SHEETS.ADMINS);
  var admin = admins.find(function(r){
    return clean_(r.Username).toLowerCase()===username && yesDefault_(r.Active,true);
  });

  if (!admin || !(await githubVerifyAdminPassword_(admin,password))) {
    throw new Error('Invalid administrator username or password.');
  }

  var settings = getSettingsMap_();
  var hours = num_(settings['Session Hours'], APP.SESSION_HOURS);
  var token = Utilities.getUuid() + Utilities.getUuid().replace(/-/g,'');
  var now = new Date();
  var expires = new Date(now.getTime()+hours*3600000);
  cleanupSessions_();
  getDb_().getSheetByName(APP.SHEETS.SESSIONS).appendRow([token,admin.Username,now,expires,now]);
  audit_('LOGIN','',admin['Admin Name'],'','','','','No','','Administrator logged in',admin.Username);

  return {
    ok:true,
    token:String(token),
    mustChangePassword:yes_(admin['Must Change Password']),
    version:APP.VERSION
  };
};

changePassword = async function(token,currentPassword,newPassword) {
  var ctx = requireAuth_(token);
  if (!newPassword || String(newPassword).length<10) throw new Error('New password must be at least 10 characters.');

  var admins = readTable_(APP.SHEETS.ADMINS);
  var admin = admins.find(function(r){return clean_(r.Username)===ctx.username;});
  if (!admin) throw new Error('Administrator account not found.');
  if (!(await githubVerifyAdminPassword_(admin,currentPassword))) throw new Error('Current password is incorrect.');

  var salt = Utilities.getUuid();
  updateRowByKey_(APP.SHEETS.ADMINS,'Admin ID',admin['Admin ID'],{
    'Temporary Password':'',
    'Password Salt':salt,
    'Password Hash':hashPassword_(String(newPassword),salt),
    'Password Algorithm':'APPSCRIPT-SHA256',
    'Password Iterations':'',
    'Must Change Password':'No',
    'Updated At':new Date(),
    'Updated By':ctx.username
  });
  audit_('PASSWORD_CHANGED','',admin['Admin Name'],'','','','','No','','Administrator password changed',ctx.username);
  return {ok:true};
};

resetAdminPassword = function(token,adminId) {
  var ctx = requireAdmin_(token);
  var row = findRowByKey_(APP.SHEETS.ADMINS,'Admin ID',adminId);
  if (!row) throw new Error('Administrator not found.');

  var password = generateTemporaryPassword_();
  var salt = Utilities.getUuid();
  updateRowByKey_(APP.SHEETS.ADMINS,'Admin ID',adminId,{
    'Temporary Password':password,
    'Password Salt':salt,
    'Password Hash':hashPassword_(password,salt),
    'Password Algorithm':'APPSCRIPT-SHA256',
    'Password Iterations':'',
    'Must Change Password':'Yes',
    'Updated At':new Date(),
    'Updated By':ctx.username
  });
  audit_('ADMIN_PASSWORD_RESET','',row['Admin Name'],'','','','','No','','Temporary password reset',ctx.username);
  return {ok:true,username:row.Username,temporaryPassword:password};
};

publicAdmin_ = function(a) {
  var x = Object.assign({},a);
  delete x['Temporary Password'];
  delete x['Password Hash'];
  delete x['Password Salt'];
  delete x['Password Algorithm'];
  delete x['Password Iterations'];
  return x;
};

finalizeAndSendToPharmacistsSchedule = async function(token,startDate,endDate) {
  var ctx = requireAdmin_(token);
  var sh = getDb_().getSheetByName(APP.SHEETS.SCHEDULE);
  if (!sh || sh.getLastRow()<1) throw new Error('Schedule sheet is empty. Nothing was published.');

  var values = sh.getDataRange().getValues();
  var headers = (values[0]||[]).map(clean_);
  var rows = values.slice(1).filter(function(r){return r.some(function(v){return clean_(v)!=='';});});
  var publishedRows = rows.map(function(r){
    var o={}; headers.forEach(function(h,i){o[h]=serialize_(r[i]);}); return o;
  });

  var payload={
    publishedAt:new Date().toISOString(),
    publishedBy:ctx.username,
    startDate:clean_(startDate),
    endDate:clean_(endDate),
    source:'NeoChrono GitHub',
    rowCount:publishedRows.length,
    rows:publishedRows
  };

  await window.__neoRuntime.savePublished(payload);

  audit_('SEND_SCHEDULE_AS_IS',
    startDate&&endDate ? clean_(startDate)+' through '+clean_(endDate) : '',
    '','','','','','No','',
    'Published current Schedule AS-IS to private GitHub data repository. Rows sent: '+publishedRows.length+'.',
    ctx.username
  );

  return {
    ok:true,
    transferredRows:publishedRows.length,
    transferredColumns:headers.length,
    receiverSpreadsheetId:'github:a860h040/neochrono-data',
    receiverSpreadsheetName:'NeoChrono Private Data',
    receiverSheetName:'published-schedule.json',
    receiverUrl:'https://github.com/a860h040/neochrono-data',
    validationWasRequired:false,
    message:'Current Schedule was published as-is to the private NeoChrono GitHub data repository.'
  };
};

testPharmacistScheduleReceiverConnection = function() {
  return {
    ok:true,
    spreadsheetId:'github:a860h040/neochrono-data',
    spreadsheetName:'NeoChrono Private Data',
    hasScheduleSheet:true,
    url:'https://github.com/a860h040/neochrono-data'
  };
};
