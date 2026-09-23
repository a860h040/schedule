
(function(){
  'use strict';

  function paperMonthDates_(month){
    var arr=[];
    var last=new Date(month.getFullYear(),month.getMonth()+1,0).getDate();
    for(var i=1;i<=last;i++)arr.push(new Date(month.getFullYear(),month.getMonth(),i));
    return arr;
  }

  function paperRequestType_(value){
    return String(value||'')
      .toUpperCase()
      .replace(/[\-_]+/g,' ')
      .replace(/\s+/g,' ')
      .trim();
  }

  function paperApprovedTimeOff_(username,name,dk){
    var foundRegularOff=null;
    var rows=State.data.requests||[];

    for(var i=0;i<rows.length;i++){
      var r=rows[i];
      if(String(r.Status||'').toUpperCase()!=='APPROVED')continue;

      var requestUser=String(r.Username==null?'':r.Username).trim().toLowerCase();
      var requestName=String(r.Pharmacist||'').trim().toLowerCase();
      var wantedUser=String(username==null?'':username).trim().toLowerCase();
      var wantedName=String(name||'').trim().toLowerCase();
      var sameUser=
        (!!wantedUser&&requestUser===wantedUser) ||
        (!!wantedName&&requestName===wantedName);
      if(!sameUser)continue;

      var start=dateKey(r['Start Date']||r.Date||r['End Date']);
      var end=dateKey(r['End Date']||r.Date||r['Start Date']);
      if(!start||!end)continue;
      if(end<start){var tmp=start;start=end;end=tmp;}
      if(dk<start||dk>end)continue;

      var type=paperRequestType_(r['Record Type']);
      if(type==='PTO'){
        return {code:'P',cls:'paper-pto',title:'Approved PTO',record:r};
      }
      if(type==='REGULAR OFF'||type==='REGULAR OFF REQUEST'||type==='REGULAROFF'){
        foundRegularOff={code:'R',cls:'paper-regular-off',title:'Approved Regular Off',record:r};
      }
    }

    return foundRegularOff;
  }

  function paperApprovedPto_(username,name,dk){
    var x=paperApprovedTimeOff_(username,name,dk);
    return !!x&&x.code==='P';
  }

  function paperWeekendGroupForDate_(dk){
    var rows=(State.data.schedule||[]).filter(function(r){
      return r.Date===dk && yes(r.Weekend) && String(r['Weekend Group']||'').trim();
    });
    return rows.length?String(rows[0]['Weekend Group']).trim():'';
  }

  function paperDateIsHoliday_(dk){
    return (State.data.schedule||[]).some(function(r){
      return r.Date===dk && yes(r.Holiday);
    });
  }

  function paperUserSkillCodes_(user){
    var username=String(user&&user.Username==null?'':user.Username);
    var name=String(user&&user['Pharmacist Name']||'');

    return (State.data.skills||[])
      .filter(function(s){
        if(String(s.Active||'Yes').toLowerCase()==='no')return false;

        return (
          String(s.Username==null?'':s.Username)===username ||
          String(s['Pharmacist Name']||'')===name
        );
      })
      .map(function(s){
        return String(s.Skill||'').trim().toUpperCase();
      })
      .filter(Boolean);
  }

  function paperUserGroup_(user){
    var preferred=String(user&&user['Preferred Shift Type']||'')
      .trim()
      .toUpperCase();

    var scheduleType=String(user&&user['Schedule Type']||'')
      .trim()
      .toLowerCase()
      .replace(/\s+/g,'');

    var employmentType=String(user&&user['Employment Type']||'')
      .trim()
      .toUpperCase();

    var skills=paperUserSkillCodes_(user);

    /*
     * PRN pharmacists always appear in their own section at the bottom,
     * immediately before the UNFILLED row. This takes priority over every
     * other visual group so a PRN pharmacist is never mixed into Night,
     * Resident, Regular, or 7-on/7-off sections.
     */
    if(employmentType==='PRN'){
      return {
        rank:4,
        key:'PRN',
        label:'PRN Pharmacists'
      };
    }

    /*
     * Put dedicated night pharmacists together first.
     * Detect them from either Preferred Shift Type or their actual N1/N2 skill.
     */
    var isNight=
      preferred==='N1' ||
      preferred==='N2' ||
      preferred==='NIGHT' ||
      skills.indexOf('N1')>=0 ||
      skills.indexOf('N2')>=0;

    if(isNight){
      return {
        rank:0,
        key:'NIGHT',
        label:'Night Pharmacists'
      };
    }

    /*
     * Residents are kept together even though they are regular-schedule users.
     */
    if(yes(user&&user.Resident)){
      return {
        rank:1,
        key:'RESIDENT',
        label:'Residents'
      };
    }

    /*
     * Normal day/evening pharmacists are grouped together.
     * Preceptors remain in this Regular group unless they are Night/7-on/7-off.
     */
    var isSevenOn=
      scheduleType.indexOf('7-on')>=0 ||
      scheduleType.indexOf('7on')>=0 ||
      scheduleType.indexOf('7-on/7-off')>=0 ||
      scheduleType.indexOf('7on/7off')>=0;

    if(!isSevenOn){
      return {
        rank:2,
        key:'REGULAR',
        label:'Regular Pharmacists'
      };
    }

    /*
     * Non-night 7-on/7-off pharmacists (for example dedicated E pharmacists)
     * are kept together as their own group.
     */
    return {
      rank:3,
      key:'SEVEN_ON',
      label:'7-on / 7-off Pharmacists'
    };
  }

  function paperUserSortKey_(user){
    var preferred=String(user&&user['Preferred Shift Type']||'')
      .trim()
      .toUpperCase();

    if(preferred)return preferred;

    var skills=paperUserSkillCodes_(user);
    return skills.length?skills[0]:'ZZZ';
  }

  function paperPharmacistType_(user){
    var employment=String(user&&user['Employment Type']||'').trim().toUpperCase();
    var scheduleType=String(user&&user['Schedule Type']||'').trim().toUpperCase().replace(/\s+/g,'');
    var preferred=String(user&&user['Preferred Shift Type']||'').trim().toUpperCase();
    var skills=paperUserSkillCodes_(user);
    var sevenOn=scheduleType.indexOf('7-ON')>=0||scheduleType.indexOf('7ON')>=0;

    if(yes(user&&user.Resident))return 'RESIDENT';
    if(employment==='PRN'||scheduleType==='PRN')return 'PRN';
    if(preferred==='N1'||preferred==='N2'||preferred==='NIGHT'||(sevenOn&&(skills.indexOf('N1')>=0||skills.indexOf('N2')>=0)))return 'NIGHT';
    if(preferred==='E'||preferred==='E1'||preferred==='E2'||preferred==='EVENING'||(sevenOn&&(skills.indexOf('E')>=0||skills.indexOf('E1')>=0||skills.indexOf('E2')>=0)))return 'EVENING';
    return 'REGULAR';
  }

  function paperVisibleUsers_(){
    var f=State.calendarFilters||{};
    var selected=selectedCalendarEmployeeKeys();
    var users=(State.data.users||[]).filter(function(u){
      return yes(u.Active) && String(u.Role||'Pharmacist').toLowerCase()!=='administrator';
    });

    if(selected.length){
      users=users.filter(function(u){
        return selected.includes(String(u.Username==null?'':u.Username));
      });
    }
    if(f.pharmacistType){
      users=users.filter(function(u){return paperPharmacistType_(u)===String(f.pharmacistType||'').toUpperCase();});
    }
    if(f.preceptor){
      users=users.filter(function(u){return String(u.Preceptor||'')===f.preceptor;});
    }

    /*
     * Paper Schedule order:
     *   1. Night pharmacists
     *   2. Residents
     *   3. Regular pharmacists
     *   4. Non-night 7-on/7-off pharmacists
     *   5. PRN pharmacists
     *
     * UNFILLED is rendered after all pharmacist groups, so PRN is the
     * final pharmacist section immediately above UNFILLED.
     *
     * Within each section:
     *   preferred/home shift first, then pharmacist name.
     */
    return users.sort(function(a,b){
      var ga=paperUserGroup_(a);
      var gb=paperUserGroup_(b);

      if(ga.rank!==gb.rank)return ga.rank-gb.rank;

      var sa=paperUserSortKey_(a);
      var sb=paperUserSortKey_(b);

      if(sa!==sb)return sa.localeCompare(sb);

      return String(a['Pharmacist Name']||'')
        .localeCompare(String(b['Pharmacist Name']||''));
    });
  }

  function paperCellData_(user,dk){
    var protectedDay=paperApprovedTimeOff_(user.Username,user['Pharmacist Name'],dk);
    var rows=calendarRowsForDate(dk).filter(function(r){
      return r.Status!=='UNFILLED' &&
        String(r.Username==null?'':r.Username)===String(user.Username==null?'':user.Username);
    });

    if(rows.length){
      var codes=rows.map(function(r){return String(r.Shift||'');}).filter(Boolean);
      var r=rows[0];
      var cls='has-assignment';

      if(String(r['Shift Type']||'')==='Evening')cls+=' paper-evening';
      if(String(r['Shift Type']||'')==='Night')cls+=' paper-night';
      if(r['Coverage For Pharmacist'])cls+=' paper-coverage';
      if(yes(r.Manual) || String(r.Status||'').toUpperCase()==='MANUAL' || String(r.Status||'').toUpperCase()==='LOCKED'){
        cls+=' paper-manual';
      }
      if(yes(r.Locked) || String(r.Status||'').toUpperCase()==='LOCKED'){
        cls+=' paper-locked';
      }
      if(protectedDay){
        cls+=' paper-conflict';
      }

      var title=(protectedDay?('CONFLICT: '+protectedDay.title+' exists on this date, but a manual assignment is being retained.\n'):'')+rows.map(function(x){
        return String(x.Shift||'')+' — '+String(x['Assigned Pharmacist']||'')+
          (x['Coverage For Pharmacist']?' | covering '+x['Coverage For Pharmacist']:'')+
          (x.Warning?' | '+x.Warning:'');
      }).join('\n');

      return {
        text:codes.join('/'),
        cls:cls,
        title:title,
        id:r['Assignment ID']||'',
        locked:yes(r.Locked)||String(r.Status||'').toUpperCase()==='LOCKED',
        manual:yes(r.Manual)||['MANUAL','LOCKED'].indexOf(String(r.Status||'').toUpperCase())>=0,
        protectedDay:protectedDay
      };
    }

    var filters=State.calendarFilters||{};
    var filtersActive=!!(
      filters.status || filters.shift || filters.type ||
      filters.skill || filters.weekendGroup
    );

    if(protectedDay){
      return {text:protectedDay.code,cls:protectedDay.cls,title:protectedDay.title,id:'',protectedDay:protectedDay};
    }

    return {
      text:filtersActive?'':'X',
      cls:'paper-off paper-assignable',
      title:filtersActive?'No matching assignment':'Unassigned / available for scheduling — click to pre-assign a shift',
      id:'',
      canAssign:!filtersActive
    };
  }

  function paperManualAssignmentsForUserMonth_(user,month){
    var start=dateKey(new Date(month.getFullYear(),month.getMonth(),1));
    var end=dateKey(new Date(month.getFullYear(),month.getMonth()+1,0));
    return (State.data.schedule||[]).filter(function(r){
      var dk=String(r.Date||'').slice(0,10);
      var status=String(r.Status||'').toUpperCase();
      return dk>=start&&dk<=end &&
        String(r.Username==null?'':r.Username)===String(user.Username==null?'':user.Username) &&
        status!=='UNFILLED' &&
        String(r['Assigned Pharmacist']||'').toUpperCase()!=='UNFILLED' &&
        (yes(r.Manual)||status==='MANUAL'||status==='LOCKED');
    });
  }

  function paperAvailableXCount_(user,month){
    return paperMonthDates_(month).reduce(function(total,dt){
      var dk=dateKey(dt);
      var assigned=calendarRowsForDate(dk).some(function(r){
        return r.Status!=='UNFILLED' &&
          String(r.Username==null?'':r.Username)===String(user.Username==null?'':user.Username);
      });
      if(assigned)return total;
      if(paperApprovedTimeOff_(user.Username,user['Pharmacist Name'],dk))return total;
      return total+1;
    },0);
  }

  function paperAssignmentTotals_(month,users){
    if(!users.length)return '';
    var totalManual=0;
    var cards=users.map(function(user){
      var manual=paperManualAssignmentsForUserMonth_(user,month).length;
      var locked=paperManualAssignmentsForUserMonth_(user,month).filter(function(r){
        return yes(r.Locked)||String(r.Status||'').toUpperCase()==='LOCKED';
      }).length;
      var available=paperAvailableXCount_(user,month);
      totalManual+=manual;
      return '<div class="paper-total-item"><b>'+esc(user['Pharmacist Name']||'')+'</b><span>'+
        manual+' manual'+(locked?' · '+locked+' locked':'')+' · '+available+' X available</span></div>';
    }).join('');

    return '<details class="paper-assignment-totals" open>'+
      '<summary>Pre-scheduling totals — <b>'+totalManual+'</b> manual assignment'+(totalManual===1?'':'s')+'</summary>'+
      '<div class="paper-total-grid">'+cards+'</div>'+
    '</details>';
  }

  function paperUnfilledForDate_(dk){
    return calendarRowsForDate(dk).filter(function(r){
      return r.Status==='UNFILLED';
    });
  }

  function paperUnfilledCell_(rows){
    if(!rows.length)return '';
    return rows.map(function(r){
      return '<span class="paper-unfilled-chip" title="'+
        attr(String(r.Shift||'')+(r.Warning?' — '+r.Warning:''))+
        '" onclick="openAssignmentModal(\''+attr(r['Assignment ID']||'')+'\')">'+
        esc(r.Shift||'UNFILLED')+
        '</span>';
    }).join('');
  }


  function paperShiftCoverageSummary_(month){
    var start=dateKey(new Date(month.getFullYear(),month.getMonth(),1));
    var end=dateKey(new Date(month.getFullYear(),month.getMonth()+1,0));

    var rows=(State.data.schedule||[]).filter(function(r){
      var dk=String(r.Date||'').slice(0,10);
      var status=String(r.Status||'').toUpperCase();
      return dk>=start && dk<=end && status!=='CANCELLED' && String(r.Shift||'').trim();
    });

    var byShift={};
    rows.forEach(function(r){
      var code=String(r.Shift||'').trim().toUpperCase();
      if(!byShift[code])byShift[code]={shift:code,total:0,filled:0,left:0};
      byShift[code].total++;

      var assigned=String(r['Assigned Pharmacist']||r.Username||'').trim();
      var status=String(r.Status||'').toUpperCase();
      var isFilled=status!=='UNFILLED' && assigned && assigned.toUpperCase()!=='UNFILLED';
      if(isFilled)byShift[code].filled++;
    });

    Object.keys(byShift).forEach(function(code){
      byShift[code].left=Math.max(0,byShift[code].total-byShift[code].filled);
    });

    var shiftMeta={};
    (State.data.shifts||[]).forEach(function(s,i){
      var code=String(s.Shift||'').trim().toUpperCase();
      if(!code)return;
      shiftMeta[code]={
        priority:num(s.Priority,50),
        start:String(s.Start||''),
        index:i
      };
    });

    return Object.values(byShift).sort(function(a,b){
      var A=shiftMeta[a.shift]||{priority:50,start:'',index:9999};
      var B=shiftMeta[b.shift]||{priority:50,start:'',index:9999};
      return A.priority-B.priority ||
        A.start.localeCompare(B.start) ||
        A.index-B.index ||
        a.shift.localeCompare(b.shift);
    });
  }

  function paperShiftCoverageBoxes_(month){
    var summary=paperShiftCoverageSummary_(month);
    if(!summary.length){
      return '<div class="paper-coverage-empty">No generated shift positions for this month.</div>';
    }

    return summary.map(function(x){
      var pct=x.total?Math.round((x.filled/x.total)*100):0;
      var complete=x.left===0;
      return '<div class="paper-coverage-box '+(complete?'complete':'needs-coverage')+'">'+
        '<div class="paper-coverage-shift">'+esc(x.shift)+'</div>'+
        '<div class="paper-coverage-count"><b>'+x.filled+'</b> of <b>'+x.total+'</b> filled</div>'+
        '<div class="paper-coverage-left">'+
          (complete?'Fully covered':('<b>'+x.left+'</b> left'))+
        '</div>'+
        '<div class="paper-coverage-bar"><span style="width:'+pct+'%"></span></div>'+
      '</div>';
    }).join('');
  }

  window.renderPaperSchedule=function(){
    var m=State.month;
    var d=State.data;
    var filt=State.calendarFilters;
    var dates=paperMonthDates_(m);
    var users=paperVisibleUsers_();

    var html=
      '<div class="paper-page-shell">'+
      '<div class="calendar-head paper-page-head">'+
        '<div>'+
          '<div class="calendar-title">'+esc(monthTitle(m))+' — Paper Schedule</div>'+
          '<div class="muted small">Excel-style view: pharmacists down the left, dates across the top. X = unassigned/available, P = approved PTO, R = approved Regular Off.</div>'+
        '</div>'+
      '</div>';

    html+=
      '<div class="paper-coverage-section">'+
        '<div class="paper-coverage-section-title">Shift Coverage</div>'+
        '<div class="paper-coverage-grid">'+paperShiftCoverageBoxes_(m)+'</div>'+
      '</div>';

    html+=
      '<div class="paper-filter-card">'+
        '<div class="paper-filter-row">'+
          selectFilter('pharmacistType','Pharmacist Type',['Regular','PRN','Night','Evening','Resident'],filt.pharmacistType)+
          selectFilter('preceptor','Preceptor',['Yes','No'],filt.preceptor)+
          '<button class="btn btn-ghost btn-sm paper-clear-btn" onclick="clearPaperScheduleFilters_()">Clear filters</button>'+
          '<button class="btn btn-secondary btn-sm paper-filter-action-btn" onclick="paperScheduleChangeMonth_(-1)">← Previous</button>'+
          '<button class="btn btn-secondary btn-sm paper-filter-action-btn" onclick="State.month=firstOfMonth(new Date());renderPaperSchedule()">Today</button>'+
          '<button class="btn btn-secondary btn-sm paper-filter-action-btn" onclick="paperScheduleChangeMonth_(1)">Next →</button>'+
          '<button class="btn btn-primary btn-sm paper-filter-action-btn" onclick="printPaperSchedule_()">Print / Save PDF</button>'+
          (d.isAdmin?'<button class="btn btn-primary btn-sm paper-filter-action-btn" onclick="openAssignmentModal(null)">+ Manual assignment</button>':'')+
          (d.isAdmin?'<button id="paperGenerateScheduleBtn" class="btn btn-primary btn-sm paper-filter-action-btn" onclick="paperGenerateSchedule_()">Generate Schedule</button>':'')+
        '</div>'+
      '</div>';

    html+=
      '<div class="paper-legend">'+
        '<span class="key"><span class="swatch" style="background:#fff"></span>X = available</span>'+
        '<span class="key"><span class="swatch" style="background:var(--paper-pto)"></span>P = PTO</span>'+
        '<span class="key"><span class="swatch" style="background:var(--paper-regular-off)"></span>R = Regular Off</span>'+
        '<span class="key"><span class="swatch paper-locked-swatch"></span>Locked manual</span>'+
        '<span class="key"><span class="swatch" style="background:var(--paper-evening)"></span>Evening</span>'+
        '<span class="key"><span class="swatch" style="background:var(--paper-night)"></span>Night</span>'+
        '<span class="key"><span class="swatch" style="background:var(--paper-coverage)"></span>OFF-day coverage</span>'+
        '<span class="key"><span class="swatch" style="background:#ffe6e6"></span>Unfilled</span>'+
      '</div>';

    html+='<div id="paperPrintArea" class="paper-shell"><table class="paper-table">'+
      '<colgroup><col class="paper-name-column">'+
      dates.map(function(){return '<col class="paper-date-column">';}).join('')+
      '</colgroup><thead>';

    html+=
      '<tr class="paper-title-row">'+
        '<th class="name-col">SCHEDULE PERIOD</th>'+
        '<th colspan="'+dates.length+'">'+esc(monthTitle(m))+'</th>'+
      '</tr>';

    html+=
      '<tr class="paper-group-row">'+
        '<th class="name-col">Weekend Group</th>'+
        dates.map(function(dt){
          var dk=dateKey(dt);
          var g=paperWeekendGroupForDate_(dk);
          return '<th class="date-col '+(isWeekend(dt)?'weekend-head ':'')+
            (paperDateIsHoliday_(dk)?'holiday-col':'')+'">'+esc(g)+'</th>';
        }).join('')+
      '</tr>';

    html+=
      '<tr class="paper-date-row">'+
        '<th class="name-col">Pharmacist</th>'+
        dates.map(function(dt){
          var dk=dateKey(dt);
          return '<th class="date-col '+(paperDateIsHoliday_(dk)?'holiday-col':'')+'">'+
            (dt.getMonth()+1)+'/'+dt.getDate()+
          '</th>';
        }).join('')+
      '</tr>';

    html+=
      '<tr class="paper-day-row">'+
        '<th class="name-col">Name</th>'+
        dates.map(function(dt){
          var dk=dateKey(dt);
          return '<th class="date-col '+(paperDateIsHoliday_(dk)?'holiday-col':'')+'">'+
            ['SUN','MON','TUE','WED','THU','FRI','SAT'][dt.getDay()]+
          '</th>';
        }).join('')+
      '</tr></thead><tbody>';

    if(State.calendarFilters.status!=='UNFILLED'){
      var lastPaperGroupKey='';

      users.forEach(function(user){
        var group=paperUserGroup_(user);

        if(group.key!==lastPaperGroupKey){
          lastPaperGroupKey=group.key;

          html+=
            '<tr class="paper-category-row">'+
              '<td colspan="'+(dates.length+1)+'">'+
                esc(group.label)+
              '</td>'+
            '</tr>';
        }

        html+='<tr><td class="name-col">'+esc(user['Pharmacist Name']||'')+'</td>';

        dates.forEach(function(dt){
          var dk=dateKey(dt);
          var cell=paperCellData_(user,dk);
          var holiday=paperDateIsHoliday_(dk);
          var weekend=isWeekend(dt);
          var onclick='';
          if(cell.id){
            onclick=' onclick="openAssignmentModal(\''+attr(cell.id)+'\')"';
          }else if(cell.canAssign&&d.isAdmin){
            onclick=' onclick="paperOpenManualCell_(\''+attr(String(user.Username==null?'':user.Username))+'\',\''+attr(dk)+'\')"';
          }

          html+='<td class="paper-cell date-col '+
            (weekend?'weekend ':'')+
            (holiday?'holiday-col ':'')+
            cell.cls+
            '" title="'+attr(cell.title||'')+'"'+onclick+'>'+
            esc(cell.text)+
          '</td>';
        });

        html+='</tr>';
      });
    }

    var selected=selectedCalendarEmployeeKeys();
    var showUnfilled=
      State.calendarFilters.status==='UNFILLED' ||
      (!selected.length && State.calendarFilters.status!=='FILLED') ||
      (selected.length && State.calendarFilters.showUnfilledWithSelected && State.calendarFilters.status!=='FILLED');

    if(showUnfilled){
      html+='<tr class="unfilled-row"><td class="name-col">UNFILLED</td>';

      dates.forEach(function(dt){
        var dk=dateKey(dt);
        var rows=paperUnfilledForDate_(dk);
        html+='<td class="paper-cell paper-unfilled date-col '+
          (isWeekend(dt)?'weekend ':'')+
          (paperDateIsHoliday_(dk)?'holiday-col':'')+
          '">'+paperUnfilledCell_(rows)+'</td>';
      });

      html+='</tr>';
    }

    html+=
      '</tbody></table></div>'+
      '<div class="paper-note">Tip: click any shift code to review or edit that assignment. The left pharmacist column stays fixed while you scroll horizontally, like a printed Excel schedule.</div>'+
      '</div>';

    $('content').innerHTML=html;
    bindPaperScheduleFilters_();
  };

  function bindPaperScheduleFilters_(){
    document.querySelectorAll('.cal-filter').forEach(function(el){
      el.onchange=function(){
        State.calendarFilters[el.dataset.filter]=el.value;
        renderPaperSchedule();
      };
    });

    document.querySelectorAll('.cal-employee-filter').forEach(function(el){
      el.onchange=function(){
        var slots=normalizedEmployeeFilterSlots();
        slots[num(el.dataset.employeeSlot,0)]=String(el.value||'');
        State.calendarFilters.employees=slots;
        renderPaperSchedule();
      };
    });
  }

  window.togglePaperShowUnfilled_=function(v){
    State.calendarFilters.showUnfilledWithSelected=!!v;
    renderPaperSchedule();
  };

  window.clearPaperScheduleFilters_=function(){
    State.calendarFilters.employees=['','','','','',''];
    State.calendarFilters.showUnfilledWithSelected=false;
    ['status','shift','type','skill','weekendGroup','pharmacistType','preceptor'].forEach(function(k){
      State.calendarFilters[k]='';
    });
    renderPaperSchedule();
  };

  window.paperOpenManualCell_=function(username,dk){
    if(!State.data||!State.data.isAdmin)return;
    var user=(State.data.users||[]).find(function(u){
      return String(u.Username==null?'':u.Username)===String(username==null?'':username);
    });
    if(!user)return;

    if(paperApprovedTimeOff_(user.Username,user['Pharmacist Name'],dk)){
      toast('This date is protected by approved time off and cannot be automatically scheduled.','error');
      return;
    }

    var active=(State.data.shifts||[]).filter(function(s){return yes(s.Active);});
    var preferred=String(user['Preferred Shift Type']||'').trim().toUpperCase();
    var chosen=active.find(function(s){return String(s.Shift||'').trim().toUpperCase()===preferred;});
    var shift=chosen?chosen.Shift:(active.length?active[0].Shift:'');
    openAssignmentModal(null,dk,shift,1,String(user.Username==null?'':user.Username));
  };

  window.paperGenerateSchedule_=async function(){
    if(!State.data||!State.data.isAdmin)return;

    var month=State.month;
    var start=dateKey(new Date(month.getFullYear(),month.getMonth(),1));
    var end=dateKey(new Date(month.getFullYear(),month.getMonth()+1,0));
    var btn=$('paperGenerateScheduleBtn');
    var users=paperVisibleUsers_();
    var manualTotal=users.reduce(function(sum,u){return sum+paperManualAssignmentsForUserMonth_(u,month).length;},0);

    if(!confirm(
      'Generate the remaining schedule for '+monthTitle(month)+'?\n\n'+
      manualTotal+' visible manual assignment(s) will be treated as fixed scheduling constraints. '+
      'Approved PTO (P), approved Regular Off (R), and all manual/locked assignments will not be overwritten.'
    ))return;

    try{
      if(btn){btn.disabled=true;btn.textContent='Checking...';}

      var pre=await server('preflightScheduleGeneration',State.token,{startDate:start,endDate:end});
      if(!pre.ok)throw new Error((pre.errors||[]).join('\n')||'Schedule preflight failed.');

      var overwriteConfirmed=false;
      if(pre.requiresOverwriteConfirmation){
        var replaceable=Number(pre.replaceableCount||0);
        var protectedCount=Number(pre.protectedCount||0);
        if(!confirm(
          'This month contains '+replaceable+' generated/unprotected schedule row(s) that will be regenerated.\n'+
          protectedCount+' manual/locked row(s) will remain protected.\n\nContinue?'
        ))return;
        overwriteConfirmed=true;
      }

      if(btn)btn.textContent='Generating...';
      var result=await server('generateSchedule',State.token,{
        startDate:start,
        endDate:end,
        mode:'OVERWRITE',
        chunked:'No',
        overwriteConfirmed:overwriteConfirmed?'Yes':'No',
        batchGenerationId:pre.batchGenerationId
      });

      if(!result.ok)throw new Error((result.errors||[]).join('\n')||'Schedule generation failed.');

      await refreshData(false);
      renderPaperSchedule();

      var errors=result.errors||[];
      var warnings=result.warnings||[];
      var body='<div class="alert '+(errors.length?'alert-warn':'alert-ok')+'"><b>'+
        (errors.length?'Schedule generated with conflicts to review.':'Schedule generated around the protected pre-assignments.')+
        '</b><br>'+esc(String(result.filled||0))+' of '+esc(String(result.required||0))+' required shift positions filled.</div>';
      if(errors.length){
        body+='<div class="alert alert-danger"><b>Conflicts</b><br>'+errors.slice(0,30).map(esc).join('<br>')+'</div>';
      }
      if(warnings.length){
        body+='<div class="alert alert-warn"><b>Warnings</b><br>'+warnings.slice(0,30).map(esc).join('<br>')+'</div>';
      }
      openModal('Paper Schedule Generation',body,[{label:'Close',cls:'btn-secondary',action:'closeModal()'}]);
    }catch(e){
      toast(e&&e.message?e.message:String(e),'error');
    }finally{
      if(btn){btn.disabled=false;btn.textContent='Generate Schedule';}
    }
  };

  window.paperScheduleChangeMonth_=function(n){
    State.month=addMonths(State.month,n);
    renderPaperSchedule();
  };

  window.printPaperSchedule_=function(){
    var oldTitle=document.title;
    document.title='Pharmacy Paper Schedule - '+monthTitle(State.month);
    document.body.classList.add('printing-paper-schedule');
    window.print();
    setTimeout(function(){
      document.body.classList.remove('printing-paper-schedule');
      document.title=oldTitle;
    },300);
  };
})();
