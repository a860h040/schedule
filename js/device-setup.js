import { saveGithubConfig, testGithubConnection, readJsonFile } from './github-store.js?v=20260918-setup2';

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

    const existing=await readJsonFile('data/workbook.json',cfg,true);
    if(!existing.sha){
      throw new Error('The existing NeoChrono database was not found at data/workbook.json.');
    }

    $('msg').style.color='#147a4c';
    $('msg').textContent='Connected successfully. Return to NeoChrono and sign in.';
  }catch(err){
    $('msg').style.color='#a61b1b';
    $('msg').textContent=err.message;
  }finally{
    $('saveBtn').disabled=false;
  }
});