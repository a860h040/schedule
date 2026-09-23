(function(){
  'use strict';

  const ENDPOINT='https://script.google.com/macros/s/AKfycbzd-2TtruGPlfUbRlS7EDFlm7jgBlBUgfM66PJX03zMpZOHFoXBSXBd-rJMX7s71fXk/exec';
  const SHEET='PTO / Availability Requests';
  const READ_ACTION='neochronoPto';
  const WRITE_ACTION='neochronoPtoWrite';
  const MESSAGE_TYPE='NEOCHRONO_PTO_RECEIVER';

  const HEADERS=[
    'Record Type','Record ID','Pharmacist','Username','Date','Start Date','End Date',
    'Weekend Saturday','Weekend Sunday','Available','Status','Comment','Submitted At',
    'Reviewed By','Reviewed At','Updated At','Updated By'
  ];

  function actor_(){
    try{
      const u=window.State&&State.data&&State.data.user?State.data.user:{};
      return String(u.Username||u.username||u['Pharmacist Name']||u.Name||'NeoChrono').trim()||'NeoChrono';
    }catch(_e){return 'NeoChrono';}
  }

  function pharmacistName_(username,fallback){
    const given=String(fallback||'').trim();
    if(given)return given;
    const userKey=String(username||'').trim().toLowerCase();
    if(!userKey)return '';

    // First try the current UI data when it is exposed.
    try{
      const users=window.State&&window.State.data&&Array.isArray(window.State.data.users)
        ? window.State.data.users
        : [];
      const match=users.find(u=>String(u.Username||u.username||'').trim().toLowerCase()===userKey);
      const name=match?String(match['Pharmacist Name']||match.Pharmacist||match.Name||'').trim():'';
      if(name)return name;
    }catch(_e){}

    // During a GitHub runtime mutation the authoritative workbook is loaded.
    // Read Users directly so PTO can never lose the pharmacist name simply
    // because State is a lexical global rather than window.State.
    try{
      if(window.__neoRuntime&&typeof window.__neoRuntime.currentBook==='function'){
        const book=window.__neoRuntime.currentBook();
        const sh=book&&book.getSheetByName?book.getSheetByName('Users'):null;
        const matrix=sh&&sh.getDataRange?sh.getDataRange().getValues():[];
        if(matrix&&matrix.length){
          const headers=(matrix[0]||[]).map(x=>String(x??'').trim());
          const ui=headers.indexOf('Username');
          const ni=headers.indexOf('Pharmacist Name');
          if(ui>=0&&ni>=0){
            const row=matrix.slice(1).find(r=>String(r[ui]??'').trim().toLowerCase()===userKey);
            if(row){
              const name=String(row[ni]??'').trim();
              if(name)return name;
            }
          }
        }
      }
    }catch(_e){}

    return '';
  }

  function makeRecordId_(){
    const d=new Date(),pad=n=>String(n).padStart(2,'0');
    const stamp=''+d.getFullYear()+pad(d.getMonth()+1)+pad(d.getDate())+pad(d.getHours())+pad(d.getMinutes())+pad(d.getSeconds());
    return 'REQ-'+stamp+'-'+Math.floor(Math.random()*900+100);
  }

  function requestId_(){
    if(window.crypto&&crypto.randomUUID)return crypto.randomUUID();
    return 'PTO-'+Date.now()+'-'+Math.random().toString(36).slice(2);
  }

  function snapshotMatrix_(payload){
    if(!payload||payload.ok===false||payload.success===false){
      throw new Error(payload&&payload.message?payload.message:'Google PTO receiver returned an error.');
    }
    const headers=Array.isArray(payload.headers)&&payload.headers.length
      ? payload.headers.map(x=>String(x??'').trim())
      : HEADERS.slice();
    const rows=Array.isArray(payload.rows)?payload.rows:[];
    if(rows.length&&Array.isArray(rows[0])){
      return [headers].concat(rows.map(row=>{
        const out=row.slice(0,headers.length);
        while(out.length<headers.length)out.push('');
        return out;
      }));
    }
    return [headers,...rows.map(row=>headers.map(h=>row&&row[h]!==undefined?row[h]:''))];
  }

  function postBridge_(action,payload,timeoutMs){
    return new Promise((resolve,reject)=>{
      const requestId=requestId_();
      const frameName='neoPtoFrame_'+Date.now()+'_'+Math.random().toString(36).slice(2);
      const iframe=document.createElement('iframe');
      iframe.name=frameName;
      iframe.style.display='none';

      const form=document.createElement('form');
      form.method='POST';
      form.action=ENDPOINT;
      form.target=frameName;
      form.style.display='none';

      function addField(name,value){
        const input=document.createElement('input');
        input.type='hidden';
        input.name=name;
        input.value=String(value===null||value===undefined?'':value);
        form.appendChild(input);
      }

      addField('action',String(action||WRITE_ACTION));
      addField('sheet',SHEET);
      addField('requestId',requestId);
      addField('payload',JSON.stringify(payload||{}));

      let finished=false,timer=null;

      function cleanup(){
        if(finished)return;
        finished=true;
        if(timer)clearTimeout(timer);
        window.removeEventListener('message',onMessage);
        try{form.remove();}catch(_e){}
        try{iframe.remove();}catch(_e){}
      }

      function fail(message){cleanup();reject(new Error(message));}

      function onMessage(event){
        const data=event&&event.data;
        if(!data||data.type!==MESSAGE_TYPE)return;
        if(String(data.requestId||'')!==String(requestId))return;
        cleanup();
        if(data.ok===false||data.success===false){
          reject(new Error(data.message||'Google PTO receiver rejected the request.'));
          return;
        }
        try{
          data.matrix=snapshotMatrix_(data);
          resolve(data);
        }catch(e){reject(e);}
      }

      window.addEventListener('message',onMessage);
      timer=setTimeout(()=>{
        fail('Google did not confirm the PTO change. Make sure the latest Code5.gs is deployed as a new version of the existing Apps Script web app.');
      },Number(timeoutMs||25000));

      iframe.onerror=()=>fail('The Google PTO receiver could not be reached.');
      document.body.appendChild(iframe);
      document.body.appendChild(form);

      try{form.submit();}catch(e){fail(e&&e.message?e.message:String(e));}
    });
  }

  async function saveRequest(data){
    const row={...(data||{})};
    row['Record Type']=String(row['Record Type']||'PTO').trim().toUpperCase();
    if(row['Record Type']!=='PTO')throw new Error('Google PTO bridge only accepts PTO records.');
    row['Record ID']=String(row['Record ID']||'').trim()||makeRecordId_();
    row.Pharmacist=pharmacistName_(row.Username,row.Pharmacist);
    row['Updated By']=String(row['Updated By']||actor_()).trim()||actor_();

    const result=await postBridge_(WRITE_ACTION,{operation:'save',row,actor:actor_()});
    return {
      ok:true,
      recordId:row['Record ID'],
      status:String(result.status||''),
      autoApproved:!!result.autoApproved,
      blockedDates:Array.isArray(result.blockedDates)?result.blockedDates:[],
      matrix:result.matrix,
      message:result.message||'PTO saved to Google Sheet and ready to sync to neochrono-data.'
    };
  }

  async function reviewRequest(recordId,status,comment){
    const id=String(recordId||'').trim();
    if(!id)throw new Error('Request ID is required.');
    const result=await postBridge_(WRITE_ACTION,{
      operation:'review',
      recordId:id,
      status:String(status||'').trim(),
      comment:comment===undefined?'':String(comment),
      actor:actor_()
    });
    return {
      ok:true,
      recordId:id,
      status:String(result.status||status||''),
      matrix:result.matrix,
      message:result.message||'PTO review saved to Google Sheet and ready to sync to neochrono-data.'
    };
  }

  async function removeRequest(recordId){
    const id=String(recordId||'').trim();
    if(!id)throw new Error('Request ID is required.');
    const result=await postBridge_(WRITE_ACTION,{operation:'delete',recordId:id,actor:actor_()});
    return {
      ok:true,
      recordId:id,
      matrix:result.matrix,
      message:result.message||'PTO deleted from Google Sheet and ready to sync to neochrono-data.'
    };
  }

  async function fetchSnapshot(){
    const result=await postBridge_(READ_ACTION,{},20000);
    return {
      ok:true,
      matrix:result.matrix,
      rowCount:Number(result.rowCount||Math.max(0,(result.matrix||[]).length-1)),
      generatedAt:String(result.generatedAt||''),
      message:result.message||'PTO snapshot loaded from Google Sheet.'
    };
  }

  function status(){
    return {endpoint:ENDPOINT,sheet:SHEET,source:'neochrono-data',upstream:'google-sheet'};
  }

  window.__neoGooglePtoSource={
    endpoint:ENDPOINT,
    sheet:SHEET,
    headers:HEADERS.slice(),
    saveRequest,
    reviewRequest,
    removeRequest,
    fetchSnapshot,
    status,
    load:async()=>null,
    isExternalRecord:()=>false,
    isExternalPtoRecord:()=>false
  };
})();
