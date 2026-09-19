
(function(){
  'use strict';

  const legacyRenderDashboard=window.renderDashboard;
  const legacyGenerateScheduleUI=window.generateScheduleUI;

  function state_(){
    State.monthlySchedule=State.monthlySchedule||{
      generateMonth:'',
      statsMonth:'',
      exportMonths:[],
      deleteMonths:[]
    };
    return State.monthlySchedule;
  }

  function two_(n){return String(n).padStart(2,'0');}

  function currentMonthKey_(){
    const d=new Date();
    return d.getFullYear()+'-'+two_(d.getMonth()+1);
  }

  function monthLabel_(key){
    const m=String(key||'').match(/^(\d{4})-(\d{2})$/);
    if(!m)return key||'';
    return new Intl.DateTimeFormat('en-US',{month:'long',year:'numeric'}).format(
      new Date(Number(m[1]),Number(m[2])-1,1)
    );
  }

  function monthBounds_(key){
    const m=String(key||'').match(/^(\d{4})-(\d{2})$/);
    if(!m)return null;
    const y=Number(m[1]),mo=Number(m[2])-1;
    if(mo<0||mo>11)return null;
    const first=new Date(y,mo,1);
    const last=new Date(y,mo+1,0);
    return {
      key:key,
      first:first,
      last:last,
      start:y+'-'+two_(mo+1)+'-01',
      end:y+'-'+two_(mo+1)+'-'+two_(last.getDate()),
      days:last.getDate()
    };
  }

  function rowDateKey_(r){
    const raw=r&&r.Date;
    if(raw instanceof Date&&!isNaN(raw))return dateKey(raw);
    const s=String(raw||'').trim();
    const m=s.match(/^(\d{4})-(\d{2})-(\d{2})/);
    if(m)return m[1]+'-'+m[2]+'-'+m[3];
    const d=parseDate(s);
    return d&&!isNaN(d)?dateKey(d):'';
  }

  function rowMonthKey_(r){
    const dk=rowDateKey_(r);
    return dk?dk.slice(0,7):'';
  }

  function scheduleMonths_(){
    return [...new Set(
      (State.data&&State.data.schedule||[])
        .map(rowMonthKey_)
        .filter(Boolean)
    )].sort();
  }

  function syncSelectedMonth_(key){
    const s=state_();
    const bounds=monthBounds_(key||s.generateMonth||currentMonthKey_());
    if(!bounds)return null;
    s.generateMonth=bounds.key;

    const monthEl=document.getElementById('genMonth');
    const startEl=document.getElementById('genStart');
    const endEl=document.getElementById('genEnd');
    const preview=document.getElementById('genMonthRange');

    if(monthEl&&monthEl.value!==bounds.key)monthEl.value=bounds.key;
    if(startEl)startEl.value=bounds.start;
    if(endEl)endEl.value=bounds.end;
    if(preview)preview.textContent=monthLabel_(bounds.key)+' · '+bounds.start+' through '+bounds.end;

    return bounds;
  }

  window.changeGenerationMonth_=function(value){
    syncSelectedMonth_(value);
  };

  // Existing generation/validation/finalize functions still reference genStart
  // and genEnd. Keep those IDs, but make them represent exactly one calendar month.
  window.updateScheduleEndDateFromStart_=function(){
    const monthEl=document.getElementById('genMonth');
    if(monthEl&&monthEl.value){
      syncSelectedMonth_(monthEl.value);
      return;
    }

    const startEl=document.getElementById('genStart');
    if(!startEl||!startEl.value)return;
    const key=String(startEl.value).slice(0,7);
    syncSelectedMonth_(key);
  };

  function monthHealth_(key){
    const rows=(State.data.schedule||[]).filter(r=>rowMonthKey_(r)===key);
    const required=rows.length;
    const unfilled=rows.filter(r=>
      String(r.Status||'').trim().toUpperCase()==='UNFILLED' ||
      String(r['Assigned Pharmacist']||'').trim().toUpperCase()==='UNFILLED'
    ).length;
    const filled=Math.max(0,required-unfilled);
    const warnings=rows.filter(r=>String(r.Warning||'').trim()).length;
    return {
      required,filled,unfilled,warnings,
      coverage:required?Math.round(filled*1000/required)/10:100
    };
  }

  function monthStats_(key){
    const users=(State.data.users||[])
      .filter(u=>String(u.Role||'Pharmacist').toLowerCase()!=='administrator'&&yes(u.Active));
    const rows=(State.data.schedule||[]).filter(r=>{
      if(rowMonthKey_(r)!==key)return false;
      if(String(r.Status||'').trim().toUpperCase()==='UNFILLED')return false;
      const name=String(r['Assigned Pharmacist']||'').trim();
      return !!name&&name.toUpperCase()!=='UNFILLED';
    });

    const by={};
    users.forEach(u=>{
      const username=String(u.Username||'').trim();
      const name=String(u['Pharmacist Name']||username||'Unknown');
      const keyName=username?'U:'+username:'N:'+name.toLowerCase();
      by[keyName]={
        username,name,
        totalHours:0,shifts:0,dayShifts:0,eveningShifts:0,nightShifts:0,weekendShifts:0,
        weekendGroup:u['Weekend Group']||'',
        scheduleType:u['Schedule Type']||'',
        weeklyMax:num(u['Weekly Hour Maximum'],40),
        weeklyHours:{}
      };
    });

    rows.forEach(r=>{
      const username=String(r.Username||'').trim();
      const name=String(r['Assigned Pharmacist']||'').trim();
      const keyName=username?'U:'+username:'N:'+name.toLowerCase();
      if(!by[keyName]){
        by[keyName]={
          username,name,totalHours:0,shifts:0,dayShifts:0,eveningShifts:0,nightShifts:0,weekendShifts:0,
          weekendGroup:r['Weekend Group']||'',scheduleType:'',weeklyMax:40,weeklyHours:{}
        };
      }
      const x=by[keyName];
      const hours=num(r['Credited Hours']!==''&&r['Credited Hours']!=null?r['Credited Hours']:r.Hours,0);
      x.totalHours+=hours;
      x.shifts++;
      const t=String(r['Shift Type']||'').trim().toLowerCase();
      if(t==='day')x.dayShifts++;
      if(t==='evening')x.eveningShifts++;
      if(t==='night')x.nightShifts++;
      if(yes(r.Weekend))x.weekendShifts++;

      const dk=rowDateKey_(r);
      if(dk){
        const d=parseDate(dk);
        if(d&&!isNaN(d)){
          const ws=addDays(d,-d.getDay());
          const wk=dateKey(ws);
          x.weeklyHours[wk]=num(x.weeklyHours[wk],0)+hours;
        }
      }
    });

    return Object.values(by).sort((a,b)=>String(a.name).localeCompare(String(b.name)));
  }

  function monthlyStatsTable_(stats){
    return '<div class="table-wrap"><table class="table monthly-stats-table">'+
      '<thead><tr><th>Pharmacist</th><th>Hours</th><th>Shifts</th><th>Day</th><th>Evening</th><th>Night</th><th>Weekend</th><th>Weekly hours</th><th>Group</th><th>Schedule</th></tr></thead>'+
      '<tbody>'+
      stats.map(s=>'<tr>'+
        '<td><b>'+esc(s.name)+'</b></td>'+
        '<td>'+fmtHours(s.totalHours)+'</td>'+
        '<td>'+s.shifts+'</td>'+
        '<td>'+s.dayShifts+'</td>'+
        '<td>'+s.eveningShifts+'</td>'+
        '<td>'+s.nightShifts+'</td>'+
        '<td>'+s.weekendShifts+'</td>'+
        '<td>'+Object.entries(s.weeklyHours||{}).map(([k,v])=>
          '<span class="badge '+(v>num(s.weeklyMax,40)?'badge-danger':'badge-ok')+'">'+
            esc(k)+': '+fmtHours(v)+' / '+fmtHours(s.weeklyMax||40)+
          '</span>'
        ).join(' ')+'</td>'+
        '<td>'+esc(s.weekendGroup||'')+'</td>'+
        '<td>'+esc(s.scheduleType||'')+'</td>'+
      '</tr>').join('')+
      '</tbody></table></div>';
  }

  function renderMonthlyStats_(key){
    const months=scheduleMonths_();
    const host=document.getElementById('monthlyStatsHost');
    if(!host)return;

    if(!months.length){
      host.innerHTML='<div class="monthly-empty">No generated schedule months yet.</div>';
      return;
    }

    const s=state_();
    if(!months.includes(key))key=months.includes(s.statsMonth)?s.statsMonth:months[0];
    s.statsMonth=key;

    const health=monthHealth_(key);
    const stats=monthStats_(key);

    host.innerHTML=
      '<div class="monthly-stats-toolbar">'+
        '<div class="monthly-month-tabs">'+
          months.map(m=>'<button type="button" class="monthly-month-tab '+(m===key?'active':'')+'" onclick="setMonthlyStatsMonth_(\''+attr(m)+'\')">'+esc(monthLabel_(m))+'</button>').join('')+
        '</div>'+
      '</div>'+
      '<div class="monthly-health-grid">'+
        '<div class="monthly-health-card"><span>Month</span><strong>'+esc(monthLabel_(key))+'</strong></div>'+
        '<div class="monthly-health-card"><span>Required positions</span><strong>'+health.required+'</strong></div>'+
        '<div class="monthly-health-card good"><span>Filled</span><strong>'+health.filled+'</strong></div>'+
        '<div class="monthly-health-card '+(health.unfilled?'bad':'good')+'"><span>Unfilled</span><strong>'+health.unfilled+'</strong></div>'+
        '<div class="monthly-health-card"><span>Coverage</span><strong>'+health.coverage+'%</strong></div>'+
        '<div class="monthly-health-card '+(health.warnings?'warn':'good')+'"><span>Warnings</span><strong>'+health.warnings+'</strong></div>'+
      '</div>'+
      monthlyStatsTable_(stats);
  }

  window.setMonthlyStatsMonth_=function(key){
    state_().statsMonth=key;
    renderMonthlyStats_(key);
  };

  function savedMonthsHtml_(){
    const months=scheduleMonths_();
    if(!months.length){
      return '<div class="monthly-empty">No schedule months have been generated yet.</div>';
    }

    return '<div class="saved-month-grid">'+months.map(key=>{
      const h=monthHealth_(key);
      return '<button type="button" class="saved-month-card" onclick="setMonthlyStatsMonth_(\''+attr(key)+'\');document.getElementById(\'monthlyStatsHost\')?.scrollIntoView({behavior:\'smooth\',block:\'start\'})">'+
        '<span>'+esc(monthLabel_(key))+'</span>'+
        '<strong>'+h.filled+' / '+h.required+' filled</strong>'+
        '<small>'+h.unfilled+' unfilled · '+h.coverage+'%</small>'+
      '</button>';
    }).join('')+'</div>';
  }

  function installMonthlyDashboard_(){
    const toolbar=document.querySelector('.premium-toolbar');
    if(toolbar){
      const s=state_();
      const months=scheduleMonths_();
      const defaultMonth=s.generateMonth||
        (months.includes(currentMonthKey_())?currentMonthKey_():(months[0]||currentMonthKey_()));

      toolbar.innerHTML=
        '<div class="field monthly-generation-field">'+
          '<label>Generate Month</label>'+
          '<input id="genMonth" type="month" class="control" value="'+attr(defaultMonth)+'" onchange="changeGenerationMonth_(this.value)">'+
          '<div id="genMonthRange" class="small muted"></div>'+
        '</div>'+
        '<input id="genStart" type="hidden">'+
        '<input id="genEnd" type="hidden">';

      syncSelectedMonth_(defaultMonth);
    }

    const sectionHeadings=[...document.querySelectorAll('.section-heading')];
    const controlHeading=sectionHeadings.find(x=>x.querySelector('h2')&&x.querySelector('h2').textContent.trim()==='Build and publish');
    if(controlHeading){
      const note=controlHeading.querySelector('.section-note');
      if(note)note.textContent='Generate one calendar month at a time';
    }

    const generateBtn=document.getElementById('generateScheduleBtn');
    if(generateBtn)generateBtn.textContent='Generate Month';

    const exportBtn=document.getElementById('exportScheduleCsvBtn');
    if(exportBtn)exportBtn.textContent='Export Months CSV';

    const deleteBtn=document.getElementById('deleteScheduleBtn');
    if(deleteBtn)deleteBtn.textContent='Delete Months';

    const controls=document.querySelector('.control-panel');
    if(controls&&!document.getElementById('savedScheduleMonths')){
      const block=document.createElement('div');
      block.id='savedScheduleMonths';
      block.className='saved-months-panel';
      block.innerHTML='<div class="saved-months-head"><div><b>Generated Months</b><span>Each month is independent. Regenerating one month does not replace the others.</span></div></div>'+savedMonthsHtml_();
      controls.appendChild(block);
    }

    const spaced=[...document.querySelectorAll('.section-heading-spaced')];
    const workload=spaced.find(x=>x.querySelector('h2')&&x.querySelector('h2').textContent.trim()==='Pharmacist distribution');
    if(workload){
      workload.innerHTML='<div><span class="section-kicker">MONTHLY WORKLOAD</span><h2>Pharmacist distribution by month</h2></div><span class="section-note">Choose one generated month</span>';
      const oldTable=workload.nextElementSibling;
      if(oldTable&&oldTable.classList.contains('table-wrap')){
        const host=document.createElement('div');
        host.id='monthlyStatsHost';
        oldTable.replaceWith(host);
      }else if(!document.getElementById('monthlyStatsHost')){
        const host=document.createElement('div');
        host.id='monthlyStatsHost';
        workload.insertAdjacentElement('afterend',host);
      }
      const months=scheduleMonths_();
      const preferred=state_().statsMonth||
        (months.includes(currentMonthKey_())?currentMonthKey_():(months[0]||''));
      renderMonthlyStats_(preferred);
    }
  }

  window.renderDashboard=function(){
    legacyRenderDashboard();
    if(State.data&&State.data.isAdmin)installMonthlyDashboard_();
  };

  window.generateScheduleUI=async function(){
    syncSelectedMonth_(document.getElementById('genMonth')?.value||state_().generateMonth||currentMonthKey_());
    try{
      await legacyGenerateScheduleUI();
    }finally{
      const btn=document.getElementById('generateScheduleBtn');
      if(btn)btn.textContent='Generate Month';
      const panel=document.getElementById('savedScheduleMonths');
      if(panel){
        panel.innerHTML='<div class="saved-months-head"><div><b>Generated Months</b><span>Each month is independent. Regenerating one month does not replace the others.</span></div></div>'+savedMonthsHtml_();
      }
      const months=scheduleMonths_();
      if(months.length)renderMonthlyStats_(state_().generateMonth||months[0]);
    }
  };

  function monthChoiceRows_(selected,mode){
    const months=scheduleMonths_();
    return months.map(key=>{
      const h=monthHealth_(key);
      const checked=selected.includes(key)?'checked':'';
      return '<label class="month-choice-row">'+
        '<input type="checkbox" class="'+mode+'-month-check" value="'+attr(key)+'" '+checked+' onchange="monthlyChoiceChanged_(\''+mode+'\',this)">'+
        '<span><b>'+esc(monthLabel_(key))+'</b><small>'+h.filled+' filled · '+h.unfilled+' unfilled · '+h.required+' positions</small></span>'+
      '</label>';
    }).join('');
  }

  window.monthlyChoiceChanged_=function(mode,el){
    const cls='.'+mode+'-month-check';
    const selected=[...document.querySelectorAll(cls+':checked')].map(x=>x.value).sort();

    if(mode==='export'&&selected.length>2){
      el.checked=false;
      toast('CSV export can combine up to two months (about 60 days).','error');
      return;
    }

    if(mode==='export'&&selected.length===2){
      const a=monthBounds_(selected[0]),b=monthBounds_(selected[1]);
      const next=new Date(a.first.getFullYear(),a.first.getMonth()+1,1);
      const nextKey=next.getFullYear()+'-'+two_(next.getMonth()+1);
      if(b.key!==nextKey){
        el.checked=false;
        toast('Choose two consecutive months for one CSV file.','error');
        return;
      }
    }

    state_()[mode+'Months']=selected;
  };

  window.exportScheduleCsvUI=function(){
    const months=scheduleMonths_();
    if(!months.length){
      toast('Generate at least one schedule month before exporting.','error');
      return;
    }

    const s=state_();
    let selected=(s.exportMonths||[]).filter(x=>months.includes(x));
    if(!selected.length){
      const current=state_().generateMonth;
      const idx=months.indexOf(current);
      if(idx>=0){
        selected=[current];
        if(months[idx+1])selected.push(months[idx+1]);
      }else{
        selected=months.slice(0,Math.min(2,months.length));
      }
      s.exportMonths=selected;
    }

    const body=
      '<div class="monthly-action-intro"><b>Select the month(s) to export.</b><p>Choose one month or two consecutive months. Two calendar months are the approximately 60-day CSV schedule.</p></div>'+
      '<div class="month-choice-list">'+monthChoiceRows_(selected,'export')+'</div>'+
      '<div class="small muted mt12">Only the selected month(s) are included. Other generated months remain unchanged.</div>';

    openModal('Export Schedule CSV by Month',body,[
      {label:'Cancel',cls:'btn-secondary',action:'closeModal()'},
      {label:'Export Selected Month(s)',cls:'btn-primary',action:'exportSelectedScheduleMonths_()'}
    ]);
  };

  window.exportSelectedScheduleMonths_=async function(){
    const selected=[...document.querySelectorAll('.export-month-check:checked')].map(x=>x.value).sort();
    if(!selected.length){toast('Select at least one month to export.','error');return;}
    if(selected.length>2){toast('Select no more than two months.','error');return;}

    if(selected.length===2){
      const a=monthBounds_(selected[0]);
      const next=new Date(a.first.getFullYear(),a.first.getMonth()+1,1);
      const nextKey=next.getFullYear()+'-'+two_(next.getMonth()+1);
      if(selected[1]!==nextKey){
        toast('The two export months must be consecutive.','error');
        return;
      }
    }

    const first=monthBounds_(selected[0]);
    const last=monthBounds_(selected[selected.length-1]);
    const start=first.start,end=last.end;

    try{
      const r=await server('exportPrintableScheduleCsv',State.token,start,end);
      if(!r||!r.ok||!r.csv)throw new Error((r&&r.message)||'CSV export did not return a file.');

      const blob=new Blob(['\uFEFF'+r.csv],{type:'text/csv;charset=utf-8;'});
      const url=URL.createObjectURL(blob);
      const a=document.createElement('a');
      a.href=url;
      a.download=r.filename||('Pharmacist_Schedule_'+start+'_to_'+end+'.csv');
      a.style.display='none';
      document.body.appendChild(a);
      a.click();
      a.remove();
      setTimeout(()=>URL.revokeObjectURL(url),1500);

      closeModal();
      toast('CSV exported: '+selected.map(monthLabel_).join(' + ')+' ('+r.dayCount+' days)','success');
    }catch(e){
      toast(e.message,'error');
    }
  };

  window.deleteScheduleUI=function(){
    const months=scheduleMonths_();
    if(!months.length){
      toast('There are no generated schedule months to delete.','error');
      return;
    }

    state_().deleteMonths=[];
    const body=
      '<div class="monthly-action-intro danger"><b>Delete selected schedule months only.</b><p>You can remove one month without affecting any other saved month. For example, deleting March leaves January, February, April, May, and later months untouched.</p></div>'+
      '<div class="month-choice-list">'+monthChoiceRows_([], 'delete')+'</div>';

    openModal('Delete Schedule Months',body,[
      {label:'Cancel',cls:'btn-secondary',action:'closeModal()'},
      {label:'Delete Selected Month(s)',cls:'btn-danger',action:'deleteSelectedScheduleMonths_()'}
    ]);
  };

  window.deleteSelectedScheduleMonths_=async function(){
    const selected=[...document.querySelectorAll('.delete-month-check:checked')].map(x=>x.value).sort();
    if(!selected.length){toast('Select at least one month to delete.','error');return;}

    const labels=selected.map(monthLabel_);
    if(!confirm(
      'DELETE SCHEDULE MONTH'+(selected.length===1?'':'S')+'?\n\n'+
      labels.join('\n')+
      '\n\nOnly these calendar months will be deleted. All other schedule months will remain unchanged.\n\nContinue?'
    ))return;

    let deleted=0;
    try{
      for(let i=0;i<selected.length;i++){
        const b=monthBounds_(selected[i]);
        const r=await server('deleteSchedulePeriod',State.token,b.start,b.end);
        deleted+=Number(r&&r.deletedRows||0);
      }

      closeModal();
      await refreshData(false);
      renderDashboard();
      const result=document.getElementById('generationResult');
      if(result){
        result.innerHTML='<div class="alert alert-ok"><b>Selected month'+(selected.length===1?'':'s')+' deleted.</b><br>'+
          esc(String(deleted))+' schedule row(s) removed from '+esc(labels.join(', '))+'. Other months were not changed.</div>';
      }
      toast(labels.join(', ')+' deleted','success');
    }catch(e){
      toast(e.message,'error');
    }
  };

  // Expose helpers for debugging and future UI additions.
  window.__neoMonthlySchedule={
    months:scheduleMonths_,
    bounds:monthBounds_,
    health:monthHealth_,
    stats:monthStats_
  };
})();
