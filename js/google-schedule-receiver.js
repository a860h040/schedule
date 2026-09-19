
(function(){
  'use strict';

  const RECEIVER_URL='https://script.google.com/macros/s/AKfycbyKpDf3Fe6TyvX0vxrQ6_5O18f1DT2ZiEciQbMEC_VObA2WE_COLzXXzwRRMjR17clK/exec';
  const SHEET_NAME='Schedule';

  const HEADERS=[
    'Generation ID',
    'Assignment ID',
    'Date',
    'Day',
    'Shift',
    'Slot',
    'Assigned Pharmacist',
    'Username',
    'Hours',
    'Credited Hours',
    'Required Skill',
    'Coverage For Pharmacist',
    'Coverage For Username',
    'Coverage Reason',
    'Shift Type',
    'Weekend',
    'Weekend Group',
    'Holiday',
    'Locked',
    'Manual',
    'Status',
    'Warning',
    'Updated At',
    'Updated By',
    'Finalized At'
  ];

  function clean_(value){
    if(value===null||value===undefined)return '';
    if(value instanceof Date&&!isNaN(value))return value.toISOString();
    return String(value);
  }

  function dateKey_(row){
    const raw=clean_(row&&row.Date);
    return raw?raw.slice(0,10):'';
  }

  function scheduleRows_(startDate,endDate){
    return (State.data.schedule||[]).filter(function(r){
      const dk=dateKey_(r);
      if(!dk)return false;
      if(startDate&&dk<startDate)return false;
      if(endDate&&dk>endDate)return false;
      return true;
    });
  }

  function makeTransferId_(){
    if(window.crypto&&crypto.randomUUID)return crypto.randomUUID();
    return 'TX-'+Date.now()+'-'+Math.random().toString(36).slice(2);
  }

  function payload_(startDate,endDate){
    const rows=scheduleRows_(startDate,endDate);

    if(!rows.length){
      throw new Error('No schedule rows were found in the selected date range.');
    }

    return {
      action:'replaceSchedule',
      sheetName:SHEET_NAME,
      transferId:makeTransferId_(),
      startDate:startDate||'',
      endDate:endDate||'',
      sentAt:new Date().toISOString(),
      headers:HEADERS.slice(),
      rows:rows.map(function(r){
        return HEADERS.map(function(h){
          return clean_(r[h]);
        });
      })
    };
  }

  function postPayload_(payload){
    return new Promise(function(resolve,reject){
      const frameName='neoSchedulePost_'+Date.now()+'_'+Math.random().toString(36).slice(2);
      const iframe=document.createElement('iframe');
      iframe.name=frameName;
      iframe.style.display='none';

      const form=document.createElement('form');
      form.method='POST';
      form.action=RECEIVER_URL;
      form.target=frameName;
      form.style.display='none';

      function addField(name,value){
        const input=document.createElement('input');
        input.type='hidden';
        input.name=name;
        input.value=String(value===null||value===undefined?'':value);
        form.appendChild(input);
      }

      addField('action','replaceSchedule');
      addField('sheet',SHEET_NAME);
      addField('transferId',payload.transferId);
      addField('payload',JSON.stringify(payload));

      let finished=false;

      function cleanup(){
        if(finished)return;
        finished=true;
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
        if(!data||data.type!=='NEOCHRONO_SCHEDULE_RECEIVER')return;
        if(String(data.transferId||'')!==String(payload.transferId||''))return;

        cleanup();

        if(!data.ok){
          reject(new Error(data.message||'Google Schedule receiver rejected the transfer.'));
          return;
        }

        resolve(data);
      }

      window.addEventListener('message',onMessage);

      const timer=setTimeout(function(){
        if(finished)return;
        fail(
          'The Google Schedule receiver did not confirm the transfer. '+
          'Make sure Code4.gs is deployed as a new version of the existing web app.'
        );
      },30000);

      const oldCleanup=cleanup;
      cleanup=function(){
        clearTimeout(timer);
        oldCleanup();
      };

      iframe.onerror=function(){
        fail('The Google Schedule receiver could not be reached.');
      };

      document.body.appendChild(iframe);
      document.body.appendChild(form);

      try{
        form.submit();
      }catch(e){
        fail(e&&e.message?e.message:String(e));
      }
    });
  }

  async function sendSchedule(startDate,endDate){
    const payload=payload_(startDate,endDate);
    const result=await postPayload_(payload);

    const expectedRows=payload.rows.length;
    const gotRows=Number(result.transferredRows||0);
    const gotCols=Number(result.transferredColumns||0);

    if(gotRows!==expectedRows){
      throw new Error(
        'Google confirmed the transfer, but row count does not match. '+
        'NeoChrono sent '+expectedRows+' row(s); Google wrote '+gotRows+'.'
      );
    }

    if(gotCols!==HEADERS.length){
      throw new Error(
        'Google confirmed the transfer, but expected 25 columns and received '+gotCols+'.'
      );
    }

    return {
      ok:true,
      transferredRows:gotRows,
      transferredColumns:gotCols,
      receiverSheetName:result.receiverSheetName||SHEET_NAME,
      receiverSpreadsheetName:result.receiverSpreadsheetName||'Pharmacists Schedule',
      receiverUrl:result.receiverUrl||RECEIVER_URL,
      transferId:payload.transferId,
      message:result.message||'Schedule written to Google Sheet.'
    };
  }

  window.__neoScheduleReceiver={
    url:RECEIVER_URL,
    sheet:SHEET_NAME,
    headers:HEADERS.slice(),
    sendSchedule:sendSchedule
  };
})();
