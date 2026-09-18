import { uid, clean } from './schema.js';
import { readDatabase, updateDatabase } from './github-store.js';

const SESSION_KEY='neochronoSessionV1';
const PBKDF2_ITERATIONS=180000;

function bytesToHex(bytes){return [...new Uint8Array(bytes)].map(b=>b.toString(16).padStart(2,'0')).join('');}
function hexToBytes(hex){const out=new Uint8Array(hex.length/2);for(let i=0;i<out.length;i++)out[i]=parseInt(hex.slice(i*2,i*2+2),16);return out;}
function randomHex(n=16){const b=new Uint8Array(n);crypto.getRandomValues(b);return bytesToHex(b);}

export async function hashPassword(password,saltHex=randomHex(16),iterations=PBKDF2_ITERATIONS){
  const key=await crypto.subtle.importKey('raw',new TextEncoder().encode(String(password)),'PBKDF2',false,['deriveBits']);
  const bits=await crypto.subtle.deriveBits({name:'PBKDF2',salt:hexToBytes(saltHex),iterations,hash:'SHA-256'},key,256);
  return {salt:saltHex,hash:bytesToHex(bits),iterations,algorithm:'PBKDF2-SHA256'};
}

async function legacyHashPassword(password,salt){
  const bytes=await crypto.subtle.digest('SHA-256',new TextEncoder().encode(`${salt}|${password}`));
  let bin='';new Uint8Array(bytes).forEach(b=>bin+=String.fromCharCode(b));
  return btoa(bin).replace(/\+/g,'-').replace(/\//g,'_').replace(/=+$/,'');
}

export async function verifyPassword(password,record){
  if(!record?.passwordHash||!record?.passwordSalt)return false;
  if(record.passwordAlgorithm==='PBKDF2-SHA256'||record.passwordIterations){
    const x=await hashPassword(password,record.passwordSalt,Number(record.passwordIterations||PBKDF2_ITERATIONS));
    return x.hash===record.passwordHash;
  }
  return (await legacyHashPassword(String(password),record.passwordSalt))===record.passwordHash;
}

export async function createCredentialRecord({name,username,password,role='Administrator',active=true,mustChangePassword=false}){
  const h=await hashPassword(password);
  return {
    id:uid(role==='Administrator'?'ADM':'EMP'),
    name:clean(name), username:clean(username), role, active,
    passwordHash:h.hash,passwordSalt:h.salt,passwordIterations:h.iterations,passwordAlgorithm:h.algorithm,
    mustChangePassword,lastLogin:'',updatedAt:new Date().toISOString(),updatedBy:'SYSTEM'
  };
}

export function currentSession(){
  try{const s=JSON.parse(sessionStorage.getItem(SESSION_KEY)||'null');if(!s)return null;if(new Date(s.expiresAt)<=new Date()){sessionStorage.removeItem(SESSION_KEY);return null;}return s;}catch{return null;}
}
export function logout(){sessionStorage.removeItem(SESSION_KEY);}

export async function login(username,password){
  const {db}=await readDatabase();
  const u=clean(username).toLowerCase();
  let rec=db.admins.find(x=>clean(x.username).toLowerCase()===u&&x.active!==false);
  let accountType='admin';
  if(!rec){rec=db.users.find(x=>clean(x.username).toLowerCase()===u&&x.active!==false);accountType='user';}
  if(!rec||!(await verifyPassword(password,rec)))throw new Error('Invalid username or password.');
  const upgradedHash=(!rec.passwordAlgorithm&&!rec.passwordIterations)?await hashPassword(password):null;
  const now=new Date(); const expires=new Date(now.getTime()+8*60*60*1000);
  const session={token:uid('SES'),username:rec.username,name:rec.name||rec.pharmacistName||rec.username,role:rec.role||'Pharmacist',isAdmin:(rec.role||'').toLowerCase()==='administrator'||accountType==='admin',accountType,createdAt:now.toISOString(),expiresAt:expires.toISOString(),mustChangePassword:!!rec.mustChangePassword};
  sessionStorage.setItem(SESSION_KEY,JSON.stringify(session));
  updateDatabase(d=>{
    const list=accountType==='admin'?d.admins:d.users;
    const x=list.find(v=>clean(v.username).toLowerCase()===u);if(x){x.lastLogin=now.toISOString();x.updatedAt=now.toISOString();if(upgradedHash)Object.assign(x,{passwordHash:upgradedHash.hash,passwordSalt:upgradedHash.salt,passwordIterations:upgradedHash.iterations,passwordAlgorithm:upgradedHash.algorithm});}
    d.audit.push({id:uid('AUD'),timestamp:now.toISOString(),user:rec.username,action:'LOGIN',details:'Successful login'});
  },`Login ${rec.username}`).catch(()=>{});
  return session;
}

export async function changePassword(currentPassword,newPassword){
  const s=currentSession();if(!s)throw new Error('Session expired.');
  if(String(newPassword||'').length<8)throw new Error('New password must be at least 8 characters.');
  const {db}=await readDatabase();
  const list=s.accountType==='admin'?db.admins:db.users;
  const rec=list.find(x=>clean(x.username).toLowerCase()===clean(s.username).toLowerCase());
  if(!rec||!(await verifyPassword(currentPassword,rec)))throw new Error('Current password is incorrect.');
  const h=await hashPassword(newPassword);
  await updateDatabase(d=>{
    const arr=s.accountType==='admin'?d.admins:d.users;const x=arr.find(v=>clean(v.username).toLowerCase()===clean(s.username).toLowerCase());
    Object.assign(x,{passwordHash:h.hash,passwordSalt:h.salt,passwordIterations:h.iterations,passwordAlgorithm:h.algorithm,mustChangePassword:false,updatedAt:new Date().toISOString(),updatedBy:s.username});
    d.audit.push({id:uid('AUD'),timestamp:new Date().toISOString(),user:s.username,action:'PASSWORD_CHANGED',details:'Password changed'});
  },`Change password ${s.username}`);
  s.mustChangePassword=false;sessionStorage.setItem(SESSION_KEY,JSON.stringify(s));
}

export async function adminResetPassword(targetUsername,newPassword){
  const s=currentSession();if(!s?.isAdmin)throw new Error('Administrator access required.');
  const h=await hashPassword(newPassword);
  await updateDatabase(d=>{
    const all=[...d.admins,...d.users]; const x=all.find(v=>clean(v.username).toLowerCase()===clean(targetUsername).toLowerCase());
    if(!x)throw new Error('Account not found.');
    Object.assign(x,{passwordHash:h.hash,passwordSalt:h.salt,passwordIterations:h.iterations,passwordAlgorithm:h.algorithm,mustChangePassword:true,updatedAt:new Date().toISOString(),updatedBy:s.username});
    d.audit.push({id:uid('AUD'),timestamp:new Date().toISOString(),user:s.username,action:'PASSWORD_RESET',pharmacist:x.name||x.pharmacistName||'',details:`Reset ${x.username}`});
  },`Reset password ${targetUsername}`);
}
