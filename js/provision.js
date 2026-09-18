import { saveGithubConfig, testGithubConnection, readJsonFile } from './github-store.js';

const $=id=>document.getElementById(id);

$('form').addEventListener('submit',async e=>{
  e.preventDefault();
  $('msg').textContent='';
  $('saveBtn').disabled=true;
  try{
    const cfg=saveGithubConfig({
      owner:$('owner').value,
      repo:$('repo').value,
      branch:$('branch').value,
      token:$('token').value,
      dbPath:$('dbPath').value,
      publishedPath:$('publishedPath').value
    });

    await testGithubConnection(cfg);

    const existing=await readJsonFile(cfg.dbPath,cfg,true);
    if(!existing.sha){
      throw new Error('NeoChrono database was not found. Public account creation is disabled. Ask the system owner to initialize the private database.');
    }

    $('msg').style.color='#147a4c';
    $('msg').textContent='Device connected successfully. Return to NeoChrono and sign in with an existing username and password.';
  }catch(err){
    $('msg').style.color='#a61b1b';
    $('msg').textContent=err.message;
  }finally{
    $('saveBtn').disabled=false;
  }
});