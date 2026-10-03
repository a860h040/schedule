(function(){
  'use strict';

  const ENDPOINT='https://script.google.com/macros/s/AKfycbzd-2TtruGPlfUbRlS7EDFlm7jgBlBUgfM66PJX03zMpZOHFoXBSXBd-rJMX7s71fXk/exec';
  const SPREADSHEET_ID='1flTBzOIM_dbDODjC-S-DHoViAhHASEyNWInZBxab5To';
  const SHEET='PRN Availability';
  const READ_ACTION='neochronoPrnAvailability';
  const MESSAGE_TYPE='NEOCHRONO_PRN_AVAILABILITY';

  function requestId_(){
    if(window.crypto&&crypto.randomUUID)return crypto.randomUUID();
    return 'PRNAVAIL-'+Date.now()+'-'+Math.random().toString(36).slice(2);
  }

  function snapshotMatrix_(payload){
    if(!payload||payload.ok===false||payload.success===false){
      throw new Error(payload&&payload.message?payload.message:'Google PRN Availability receiver returned an error.');
    }

    const headers=Array.isArray(payload.headers)
      ? payload.headers.map(x=>String(x??'').trim())
      : [];
    const rows=Array.isArray(payload.rows)?payload.rows:[];

    if(!headers.length)return [[]];

    if(rows.length&&Array.isArray(rows[0])){
      return [headers].concat(rows.map(row=>{
        const out=row.slice(0,headers.length);
        while(out.length<headers.length)out.push('');
        return out;
      }));
    }

    return [
      headers,
      ...rows.map(row=>headers.map(h=>row&&row[h]!==undefined?row[h]:''))
    ];
  }

  function fetchSnapshot(){
    return new Promise((resolve,reject)=>{
      const requestId=requestId_();
      const frameName='neoPrnAvailabilityFrame_'+Date.now()+'_'+Math.random().toString(36).slice(2);
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

      addField('action',READ_ACTION);
      addField('spreadsheetId',SPREADSHEET_ID);
      addField('sheet',SHEET);
      addField('requestId',requestId);
      addField('payload',JSON.stringify({
        spreadsheetId:SPREADSHEET_ID,
        sheet:SHEET
      }));

      let finished=false;
      let timer=null;

      function cleanup(){
        if(finished)return;
        finished=true;
        if(timer)clearTimeout(timer);
        window.removeEventListener('message',onMessage);
        try{form.remove();}catch(_e){}
        try{iframe.remove();}catch(_e){}
      }

      function fail(message){
        cleanup();
        reject(new Error(message));
      }

      function onMessage(event){
        const data=event&&event.data;
        if(!data||data.type!==MESSAGE_TYPE)return;
        if(String(data.requestId||'')!==String(requestId))return;
        cleanup();

        if(data.ok===false||data.success===false){
          reject(new Error(data.message||'Google PRN Availability receiver rejected the request.'));
          return;
        }

        try{
          resolve({
            ok:true,
            matrix:snapshotMatrix_(data),
            rowCount:Number(data.rowCount||0),
            generatedAt:String(data.generatedAt||''),
            message:data.message||'PRN availability loaded from Google Sheet "PRN Availability".'
          });
        }catch(e){
          reject(e);
        }
      }

      window.addEventListener('message',onMessage);
      timer=setTimeout(()=>{
        fail('Google did not return PRN Availability. Deploy the latest Code5.gs as a new version of the existing Apps Script web app.');
      },20000);

      iframe.onerror=()=>fail('The Google PRN Availability receiver could not be reached.');
      document.body.appendChild(iframe);
      document.body.appendChild(form);

      try{
        form.submit();
      }catch(e){
        fail(e&&e.message?e.message:String(e));
      }
    });
  }

  function installProgramSignature_(){
    if(!document.getElementById('neoProgramSignatureStyle')){
      const style=document.createElement('style');
      style.id='neoProgramSignatureStyle';
      style.textContent=`
        .neo-program-signature{
          font-size:10px;
          line-height:1.45;
          color:#8ea6b5;
          text-align:center;
          padding:10px 8px 4px;
          opacity:.95;
        }
        .neo-program-signature strong{
          color:#dce8ee;
          font-weight:700;
        }
        .login-card .neo-program-signature{
          margin-top:18px;
          padding-top:14px;
          border-top:1px solid #dbe3ee;
          color:#64748b;
        }
        .login-card .neo-program-signature strong{
          color:#334155;
        }
      `;
      document.head.appendChild(style);
    }

    const sidebar=document.querySelector('.sidebar');
    if(sidebar&&!sidebar.querySelector('.neo-program-signature')){
      const sig=document.createElement('div');
      sig.className='neo-program-signature';
      sig.innerHTML='This program was designed by<br><strong>Alhamza Hamza, PharmD</strong>';
      sidebar.appendChild(sig);
    }

    const loginCard=document.querySelector('.login-card');
    if(loginCard&&!loginCard.querySelector('.neo-program-signature')){
      const sig=document.createElement('div');
      sig.className='neo-program-signature';
      sig.innerHTML='This program was designed by <strong>Alhamza Hamza, PharmD</strong>';
      loginCard.appendChild(sig);
    }
  }

  function watchForProgramSignature_(){
    installProgramSignature_();
    const observer=new MutationObserver(()=>installProgramSignature_());
    observer.observe(document.documentElement,{childList:true,subtree:true});
  }

  window.__neoPrnAvailabilitySource={
    endpoint:ENDPOINT,
    spreadsheetId:SPREADSHEET_ID,
    sheet:SHEET,
    fetchSnapshot,
    status:()=>({
      endpoint:ENDPOINT,
      spreadsheetId:SPREADSHEET_ID,
      sheet:SHEET,
      source:'google-sheet'
    })
  };

  if(document.readyState==='loading'){
    document.addEventListener('DOMContentLoaded',watchForProgramSignature_,{once:true});
  }else{
    watchForProgramSignature_();
  }
})();