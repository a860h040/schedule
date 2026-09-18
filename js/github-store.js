import { normalizeDatabase } from './schema.js';

const CONFIG_KEY='neochronoGithubConfigV1';
const DEFAULT_DB_PATH='data/database.json';
const DEFAULT_PUBLISHED_PATH='data/published-schedule.json';

function encodeUtf8Base64(text){
  const bytes=new TextEncoder().encode(text); let bin=''; bytes.forEach(b=>bin+=String.fromCharCode(b)); return btoa(bin);
}
function decodeUtf8Base64(b64){
  const bin=atob(String(b64||'').replace(/\n/g,'')); const bytes=Uint8Array.from(bin,c=>c.charCodeAt(0)); return new TextDecoder().decode(bytes);
}

export function getGithubConfig(){
  try{return JSON.parse(localStorage.getItem(CONFIG_KEY)||'null');}catch{return null;}
}
export function saveGithubConfig(cfg){
  const clean={
    owner:String(cfg.owner||'').trim(),
    repo:String(cfg.repo||'').trim(),
    branch:String(cfg.branch||'main').trim()||'main',
    token:String(cfg.token||'').trim(),
    dbPath:String(cfg.dbPath||DEFAULT_DB_PATH).trim()||DEFAULT_DB_PATH,
    publishedPath:String(cfg.publishedPath||DEFAULT_PUBLISHED_PATH).trim()||DEFAULT_PUBLISHED_PATH
  };
  localStorage.setItem(CONFIG_KEY,JSON.stringify(clean));
  return clean;
}
export function clearGithubConfig(){localStorage.removeItem(CONFIG_KEY);}
export function hasGithubConfig(){const c=getGithubConfig(); return !!(c&&c.owner&&c.repo&&c.token);}

function apiUrl(cfg,path){return `https://api.github.com/repos/${encodeURIComponent(cfg.owner)}/${encodeURIComponent(cfg.repo)}/contents/${path.split('/').map(encodeURIComponent).join('/')}`;}
function headers(cfg){return {'Accept':'application/vnd.github+json','Authorization':`Bearer ${cfg.token}`,'X-GitHub-Api-Version':'2022-11-28','Content-Type':'application/json'};}

async function ghFetch(url,opts={}){
  const res=await fetch(url,opts);
  const txt=await res.text(); let body=null; try{body=txt?JSON.parse(txt):null;}catch{body=txt;}
  if(!res.ok){const msg=body&&body.message?body.message:`GitHub request failed (${res.status})`; const err=new Error(msg); err.status=res.status; err.body=body; throw err;}
  return body;
}

export async function testGithubConnection(cfg=getGithubConfig()){
  if(!cfg) throw new Error('This device is not provisioned for NeoChrono.');
  const repo=await ghFetch(`https://api.github.com/repos/${encodeURIComponent(cfg.owner)}/${encodeURIComponent(cfg.repo)}`,{headers:headers(cfg)});
  return {ok:true,fullName:repo.full_name,private:repo.private,defaultBranch:repo.default_branch};
}

export async function readJsonFile(path, cfg=getGithubConfig(), allowMissing=false){
  if(!cfg) throw new Error('This device is not provisioned. Ask an administrator to provision it first.');
  const url=`${apiUrl(cfg,path)}?ref=${encodeURIComponent(cfg.branch)}`;
  try{
    const file=await ghFetch(url,{headers:headers(cfg)});
    if(Array.isArray(file)) throw new Error(`${path} is a folder, not a JSON file.`);
    return {data:JSON.parse(decodeUtf8Base64(file.content||'')),sha:file.sha};
  }catch(err){if(allowMissing&&err.status===404)return {data:null,sha:null};throw err;}
}

export async function writeJsonFile(path,data,message,sha=null,cfg=getGithubConfig()){
  if(!cfg) throw new Error('This device is not provisioned.');
  const body={message:message||`NeoChrono update ${path}`,content:encodeUtf8Base64(JSON.stringify(data,null,2)),branch:cfg.branch};
  if(sha) body.sha=sha;
  return ghFetch(apiUrl(cfg,path),{method:'PUT',headers:headers(cfg),body:JSON.stringify(body)});
}

export async function readDatabase(){
  const cfg=getGithubConfig(); if(!cfg)throw new Error('This device is not provisioned.');
  const file=await readJsonFile(cfg.dbPath,cfg,false);
  return {db:normalizeDatabase(file.data),sha:file.sha};
}

export async function createDatabase(db){
  const cfg=getGithubConfig(); if(!cfg)throw new Error('Save the GitHub repository settings first.');
  const existing=await readJsonFile(cfg.dbPath,cfg,true);
  if(existing.sha) throw new Error(`Database already exists at ${cfg.dbPath}.`);
  await writeJsonFile(cfg.dbPath,normalizeDatabase(db),'Initialize NeoChrono database',null,cfg);
  return readDatabase();
}

export async function updateDatabase(mutator, commitMessage='Update NeoChrono database', maxRetries=4){
  const cfg=getGithubConfig(); if(!cfg)throw new Error('This device is not provisioned.');
  let lastError;
  for(let attempt=0;attempt<maxRetries;attempt++){
    const {db,sha}=await readDatabase();
    const working=structuredClone(db);
    const result=await mutator(working);
    working.meta=working.meta||{};
    working.meta.updatedAt=new Date().toISOString();
    try{
      await writeJsonFile(cfg.dbPath,working,commitMessage,sha,cfg);
      return {db:working,result};
    }catch(err){
      lastError=err;
      if(err.status!==409 && err.status!==422) throw err;
      await new Promise(r=>setTimeout(r,150*(attempt+1)));
    }
  }
  throw new Error(`NeoChrono could not save because another user changed the database at the same time. Reload and try again. ${lastError?.message||''}`);
}

export async function publishSchedule(rows,meta={}){
  const cfg=getGithubConfig();
  const current=await readJsonFile(cfg.publishedPath,cfg,true);
  const payload={publishedAt:new Date().toISOString(),...meta,rows};
  await writeJsonFile(cfg.publishedPath,payload,'Publish pharmacist schedule',current.sha,cfg);
  return payload;
}

export async function downloadRawJson(path){return readJsonFile(path,getGithubConfig(),false);}
