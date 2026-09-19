
(function(){
  'use strict';

  const RECEIVER_URL='https://script.google.com/macros/s/AKfycbyKpDf3Fe6TyvX0vxrQ6_5O18f1DT2ZiEciQbMEC_VObA2WE_COLzXXzwRRMjR17clK/exec';
  const SHEET_NAME='Schedule';

  function cleanDate_(value){
    if(value==null||value==='')return '';
    if(value instanceof Date&&!isNaN(value))return value.toISOString();
    return String(value);
  }

  function rowDateKey_(row){
    const raw=cleanDate_(row&&row.Date);
    return raw?raw.slice(0,10):'';
  }

  function scheduleRows_(startDate,endDate){
    return (State.data.schedule||[]).filter(function(r){
      const dk=rowDateKey_(r);
      if(!dk)return false;
      if(startDate&&dk<startDate)return false;
      if(endDate&&dk>endDate)return false;
      return true;
    });
  }

  function scheduleHeaders_(rows){
    const preferred=[
      'Assignment ID','Date','Day','Shift','Slot','Required Skill',
      'Assigned Pharmacist','Username','Hours','Credited Hours','Shift Type',
      'Weekend','Weekend Group','Holiday','Status','Locked','Manual',
      'Warning','Reason','Generation ID','Finalized','Finalized At',
      'Coverage For Pharmacist','Coverage Type','Updated At','Updated By'
    ];
    const set=new Set();
    (rows||[]).forEach(r=>Object.keys(r||{}).forEach(k=>set.add(k)));
    const out=preferred.filter(h=>set.has(h));
    [...set].forEach(h=>{if(!out.includes(h))out.push(h);});
    return out;
  }

  function payload_(startDate,endDate){
    const rows=scheduleRows_(startDate,endDate);
    if(!rows.length)throw new Error('No schedule rows were found in the selected date range.');

    const headers=scheduleHeaders_(rows);
    return {
      action:'replaceSchedule',
      sheet:SHEET_NAME,
      startDate:startDate||'',
      endDate:endDate||'',
      sentAt:new Date().toISOString(),
      headers:headers,
      rows:rows.map(r=>headers.map(h=>cleanDate_(r[h])))
    };
  }

  function postForm_(payload){
    return new Promise((resolve,reject)=>{
      const frameName='neoSchedulePost_'+Date.now()+'_'+Math.random().toString(36).slice(2);
      const iframe=document.createElement('iframe');
      iframe.name=frameName;
      iframe.style.display='none';

      const form=document.createElement('form');
      form.method='POST';
      form.action=RECEIVER_URL;
      form.target=frameName;
      form.style.display='none';

      function add(name,value){
        const input=document.createElement('input');
        input.type='hidden';
        input.name=name;
        input.value=String(value==null?'':value);
        form.appendChild(input);
      }

      add('action','replaceSchedule');
      add('sheet',SHEET_NAME);
      add('sentAt',payload.sentAt||'');
      add('payload',JSON.stringify(payload));

      let submitted=false;
      const timer=setTimeout(()=>{
        try{form.remove();}catch(_e){}
        try{iframe.remove();}catch(_e){}
        if(submitted)resolve();
        else reject(new Error('The Google Schedule receiver could not be submitted.'));
      },1200);

      iframe.onload=()=>{
        if(!submitted)return;
        clearTimeout(timer);
        setTimeout(()=>{
          try{form.remove();}catch(_e){}
          try{iframe.remove();}catch(_e){}
          resolve();
        },150);
      };

      iframe.onerror=()=>{
        clearTimeout(timer);
        try{form.remove();}catch(_e){}
        try{iframe.remove();}catch(_e){}
        reject(new Error('The Google Schedule receiver could not be reached.'));
      };

      document.body.appendChild(iframe);
      document.body.appendChild(form);

      try{
        submitted=true;
        form.submit();
      }catch(e){
        clearTimeout(timer);
        try{form.remove();}catch(_e){}
        try{iframe.remove();}catch(_e){}
        reject(e);
      }
    });
  }

  function statusJsonp_(expectedSentAt){
    return new Promise((resolve,reject)=>{
      const cb='__neoScheduleReceiverCb_'+Date.now()+'_'+Math.random().toString(36).slice(2);
      const script=document.createElement('script');
      const timer=setTimeout(()=>{
        cleanup();
        reject(new Error('The receiver did not confirm the schedule transfer.'));
      },15000);

      function cleanup(){
        clearTimeout(timer);
        try{delete window[cb];}catch(_e){window[cb]=undefined;}
        if(script.parentNode)script.parentNode.removeChild(script);
      }

      window[cb]=function(data){
        cleanup();
        if(!data||data.success===false){
          reject(new Error(data&&data.message?data.message:'Receiver verification failed.'));
          return;
        }
        resolve(data);
      };

      script.onerror=function(){
        cleanup();
        reject(new Error('Could not verify the pharmacists Schedule sheet.'));
      };

      const qs=new URLSearchParams({
        action:'status',
        sheet:SHEET_NAME,
        callback:cb,
        expectedSentAt:expectedSentAt||'',
        _:String(Date.now())
      });
      script.src=RECEIVER_URL+'?'+qs.toString();
      script.async=true;
      document.head.appendChild(script);
    });
  }

  async function sendSchedule(startDate,endDate){
    const payload=payload_(startDate,endDate);

    await postForm_(payload);

    // Give Apps Script a moment to finish writing before verification.
    let lastError=null;
    for(let attempt=0;attempt<7;attempt++){
      await new Promise(r=>setTimeout(r,1000+attempt*700));
      try{
        const status=await statusJsonp_(payload.sentAt);
        const expectedRows=payload.rows.length;
        const receiverRows=Number(status.dataRows||0);

        const receiverSentAt=String(status.lastSentAt||'');
        const exactUpload=receiverSentAt===String(payload.sentAt||'');

        if(receiverRows===expectedRows&&exactUpload){
          return {
            ok:true,
            transferredRows:expectedRows,
            transferredColumns:payload.headers.length,
            receiverSheetName:SHEET_NAME,
            receiverSpreadsheetName:status.spreadsheetName||'Pharmacists Schedule',
            receiverUrl:status.spreadsheetUrl||RECEIVER_URL,
            receiverLastReceivedAt:status.lastReceivedAt||'',
            receiverLastSentAt:receiverSentAt,
            message:'Schedule verified in Google Sheet.'
          };
        }

        if(!exactUpload){
          lastError=new Error(
            'Google still shows the previous upload. The new Schedule payload was not confirmed by Code4.gs.'
          );
        }else{
          lastError=new Error(
            'Google confirmed the new upload, but the Schedule sheet has '+receiverRows+
            ' row(s) while NeoChrono sent '+expectedRows+'.'
          );
        }
      }catch(e){
        lastError=e;
      }
    }

    throw lastError||new Error('The Google receiver did not confirm this exact schedule upload.');
  }

  window.__neoScheduleReceiver={
    url:RECEIVER_URL,
    sheet:SHEET_NAME,
    sendSchedule:sendSchedule
  };
})();
