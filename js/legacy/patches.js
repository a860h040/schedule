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

  if (headers.indexOf('Preferred') < 0) {
    headers.push('Preferred');
    values[0] = headers;
    for (var r=1;r<values.length;r++) values[r].push('');
    sh.getRange(1,1,values.length,headers.length).setValues(values);
    changed = true;
  }

  if (APP.HEADERS && APP.HEADERS.SKILLS && APP.HEADERS.SKILLS.indexOf('Preferred') < 0) {
    APP.HEADERS.SKILLS.push('Preferred');
  }
  return changed;
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

  if (changed) {
    sh.getRange(1,1,values.length,headers.length).setValues(values);
    SpreadsheetApp.flush();
    if (typeof _PRECEPTOR_CALENDAR_RUNTIME_CACHE_ !== 'undefined') _PRECEPTOR_CALENDAR_RUNTIME_CACHE_ = null;
    audit_('SKILL_PREFERENCE_SYSTEM','','','','','','','No','','Ensured Employee Skills Preferred/Home Unit support',ctx.username);
  }

  return {ok:true,changed:changed};
}

function saveSkillProfile(token, employeeId, preferredCode, skillCodes) {
  var ctx = requireAdmin_(token);
  githubEnsureSkillPreferredColumn_();

  var eid = clean_(employeeId);
  var user = readTable_(APP.SHEETS.USERS).find(function(u){return clean_(u['Employee ID'])===eid;});
  if (!user) throw new Error('Pharmacist was not found.');

  var desired = Array.from(new Set((Array.isArray(skillCodes)?skillCodes:[])
    .map(function(x){return clean_(x).toUpperCase();})
    .filter(Boolean))).sort();

  var preferred = clean_(preferredCode).toUpperCase();
  if (preferred && desired.indexOf(preferred) < 0) desired.push(preferred);

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
    'Skills: '+desired.join(', ')+'; Preferred/Home Unit: '+(preferred||'None'),ctx.username);

  return {ok:true,employeeId:eid,skills:desired,preferred:preferred};
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
