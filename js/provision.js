import { saveGithubConfig, testGithubConnection, readJsonFile, createDatabase } from './github-store.js';
import { blankDatabase } from './schema.js';
import { createCredentialRecord } from './auth.js';
const $=id=>document.getElementById(id);
$('form').addEventListener('submit',async e=>{e.preventDefault();$('msg').textContent='';$('saveBtn').disabled=true;try{
  const cfg=saveGithubConfig({owner:$('owner').value,repo:$('repo').value,branch:$('branch').value,token:$('token').value,dbPath:$('dbPath').value,publishedPath:$('publishedPath').value});
  await testGithubConnection(cfg);
  const existing=await readJsonFile(cfg.dbPath,cfg,true);
  if(!existing.sha){
    const name=$('adminName').value.trim(),username=$('adminUsername').value.trim(),password=$('adminPassword').value;
    if(!name||!username||password.length<8)throw new Error('For a new database, enter the first administrator name, username, and a password of at least 8 characters.');
    if(password!==$('adminConfirm').value)throw new Error('Administrator passwords do not match.');
    const db=blankDatabase(); db.admins.push(await createCredentialRecord({name,username,password,role:'Administrator',active:true}));
    db.audit.push({id:'AUD-BOOTSTRAP',timestamp:new Date().toISOString(),user:username,action:'DATABASE_INITIALIZED',details:'GitHub NeoChrono database initialized.'});
    await createDatabase(db);
    $('msg').style.color='#147a4c';$('msg').textContent='Device provisioned and database created. Open NeoChrono and sign in with the administrator username and password.';
  } else {$('msg').style.color='#147a4c';$('msg').textContent='Device provisioned. Existing NeoChrono database found. Open NeoChrono and sign in normally.';}
}catch(err){$('msg').style.color='#a61b1b';$('msg').textContent=err.message;}finally{$('saveBtn').disabled=false;}});
