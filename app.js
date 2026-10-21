/* MyLife Calendar 1.0 — no backend secrets; Google OAuth token exists in RAM only. */
(() => {
'use strict';
const $ = id => document.getElementById(id);
const KEYS = {config:'mylife.calendar.config.v1',demo:'mylife.calendar.demo.v1',syncReport:'mylife.calendar.syncReport.v1'};
const HEADERS=['id','owner','scope','title','start','end','allDay','notes','sync','calendarId','googleEventId','source','deleted','createdAt','updatedAt','location','revision','syncState','googleEtag','lastError'];
const SCOPE='openid email https://www.googleapis.com/auth/spreadsheets https://www.googleapis.com/auth/calendar https://www.googleapis.com/auth/drive.metadata.readonly';
const TZ='Asia/Tokyo';
const defaults={gatewayUrl:window.MYLIFE_DEPLOYMENT?.gatewayUrl||'',clientId:window.MYLIFE_DEPLOYMENT?.clientId||'',sharedSheetId:'',privateSheetId:'',sharedCalId:'',privateCalId:'',workCalId:'',visibleCalendarIds:[],visibleTouched:false,googleCalendarEnabled:true};
const parseJson=(key,fallback)=>{try{return JSON.parse(localStorage.getItem(key))||fallback;}catch{return fallback;}};
let config={...defaults,...parseJson(KEYS.config,{})};for(const k of ['clientId','gatewayUrl'])if(window.MYLIFE_DEPLOYMENT?.[k])config[k]=window.MYLIFE_DEPLOYMENT[k];
const {dateKey,keyAdd,makeLocalISO,timeLocal,workTravel,alarmTimes}=MyLifeCore;
const today=new Date(dateKey()+'T12:00:00+09:00');const state={month:new Date(today.getFullYear(),today.getMonth(),1),selected:dateKey(today),filter:'all',token:'',email:'',mode:'demo',records:[],external:[],visible:[],allCalendars:[],fetchedCalIds:new Set(),busy:false,editing:null,view:'day',expiresAt:0};
const safe=str=>String(str??'').replaceAll('&','&amp;').replaceAll('<','&lt;').replaceAll('>','&gt;').replaceAll('"','&quot;').replaceAll("'",'&#39;');
const pad=n=>String(n).padStart(2,'0');
function civilKey(d){return `${d.getFullYear()}-${pad(d.getMonth()+1)}-${pad(d.getDate())}`;}
function monthText(d){return `${d.getFullYear()}年 ${d.getMonth()+1}月`;}
function inLocal(iso){if(!iso)return '';return /^\d{4}-\d{2}-\d{2}$/.test(iso)?iso:dateKey(iso);}
function showMsg(msg,error=false){const el=$('message');el.textContent=msg;el.classList.toggle('error',error);el.classList.remove('hidden');clearTimeout(showMsg.t);showMsg.t=setTimeout(()=>el.classList.add('hidden'),7000);}
function saveConfig(){for(const k of Object.keys(defaults)){if($(k))config[k]=$(k).value.trim();}localStorage.setItem(KEYS.config,JSON.stringify(config));}
function syncInputs(){for(const k of Object.keys(defaults)){if($(k))$(k).value=config[k]||'';} const signedIn=isSignedIn();$('loggedEmail').textContent=signedIn?`ログイン中：${state.email}（MyLifeの予定保存・勤務取得に使用）`:'未ログイン';$('calendarLink').textContent=config.googleCalendarEnabled===false?'Googleカレンダー連携：OFF（「すべて同期」は勤務だけ取得します）':'Googleカレンダー連携：ON';$('loginBtn').classList.toggle('hidden',signedIn);$('disconnectGoogle').classList.toggle('hidden',config.googleCalendarEnabled===false);$('enableGoogle').classList.toggle('hidden',config.googleCalendarEnabled!==false);}
function setBusy(b){state.busy=b;$('saveEvent').disabled=b;$('deleteEvent').disabled=b;for(const id of ['syncAllBtn','settingsSyncAll'])$(id).disabled=b||!!syncRunner?.busy;}
function ensureCloud(){if(state.mode!=='cloud'||!state.token||Date.now()>=state.expiresAt)throw Error('Googleで接続してください。');}
async function api(url,init={}){
 ensureCloud();const res=await fetch(url,{...init,headers:{'Authorization':`Bearer ${state.token}`,...(init.body?{'Content-Type':'application/json'}:{}),...(init.headers||{})}});
 const txt=await res.text();let obj={};try{obj=txt?JSON.parse(txt):{};}catch{obj={message:txt.slice(0,150)};}
 if(!res.ok){const desc=obj.error?.message||obj.message||res.statusText;throw Error(`${desc}（HTTP ${res.status}）${res.status===401?'。Googleに再接続してください。':''}`);}
 return obj;
}
const sheetsRoot=id=>`https://sheets.googleapis.com/v4/spreadsheets/${encodeURIComponent(id)}`;
const a1=range=>encodeURIComponent(range).replaceAll('%21','!');
function normalizeRow(v,rowNo,store){const o={};HEADERS.forEach((key,i)=>o[key]=v[i]??'');o.allDay=o.allDay==='true'||o.allDay===true;o.sync=o.sync==='true'||o.sync===true;o.deleted=o.deleted==='true'||o.deleted===true;o.revision=+o.revision||0;o._row=rowNo;o._store=store;o.origin='sheet';o.kind=o.source==='shift'?'shift':o.scope==='shared'?'shared':'private';return o;}
function rowValues(r){return HEADERS.map(k=>['allDay','sync','deleted'].includes(k)?String(!!r[k]):String(r[k]??''));}
async function sheetRead(id,store){if(!id)return [];const data=await api(`${sheetsRoot(id)}/values/${a1('events!A:T')}?valueRenderOption=FORMATTED_VALUE`);return(data.values||[]).slice(1).map((v,i)=>normalizeRow(v,i+2,store)).filter(e=>e.id&&(!e.deleted||e.lastError));}
async function gateway(action,scope,record,expectedRevision=0){
 ensureCloud();if(action!=='bootstrap'&&state.protocol!==3)throw Error('同期サーバーの接続確認が必要です。');if(!/^https:\/\/script\.google\.com\/macros\/s\/[A-Za-z0-9_-]+\/exec$/.test(config.gatewayUrl))throw Error('設定にGASのウェブアプリURLが必要です。');
 const res=await fetch(config.gatewayUrl,{method:'POST',headers:{'Content-Type':'text/plain;charset=utf-8'},body:JSON.stringify({token:state.token,protocol:3,action,scope,record,expectedRevision}),redirect:'follow',signal:AbortSignal.timeout(90000)});
 if(!res.ok)throw Error('同期サーバーに接続できません。');const data=await res.json();if(!data.ok)throw Error(data.error||'同期サーバーでエラーが発生しました。');return data;
}
async function reconcileScope(scope){let afterId='',failed=0;do{const r=await gateway('reconcile',scope,{afterId});failed+=r.failed||0;afterId=r.nextAfter||'';}while(afterId);if(failed)throw Error(`未同期または競合の予定が${failed}件あります。`);}
async function bootstrap(){const data=await gateway('bootstrap');if(data.protocol!==3)throw Error('同期サーバーの更新が必要です。既存データは変更していません。');state.protocol=3;Object.assign(config,data.config);state.email=data.email;syncInputs();}
function toGoogleEvent(r){const o={summary:r.title,location:r.location||'',description:r.notes||'',start:r.allDay?{date:r.start}:{dateTime:r.start,timeZone:TZ},end:r.allDay?{date:r.end}:{dateTime:r.end,timeZone:TZ},extendedProperties:{shared:{mylifeId:r.id}}};return o;}
const gcalUrl=(cal,event='')=>`https://www.googleapis.com/calendar/v3/calendars/${encodeURIComponent(cal)}/events${event?'/'+encodeURIComponent(event):''}`;
function eventFromGoogle(g,cal){const isWork=cal.id===config.workCalId;return {id:`google:${cal.id}:${g.id}`,owner:g.creator?.email||'',title:g.summary||'(タイトルなし)',scope:cal.id===config.sharedCalId?'shared':'private',start:g.start?.dateTime||g.start?.date||'',end:g.end?.dateTime||g.end?.date||'',allDay:!!g.start?.date,notes:g.description||'',location:g.location||'',googleEtag:g.etag, sync:true,calendarId:cal.id,googleEventId:g.id,source:'google',deleted:false,origin:'google',kind:isWork?'shift':'google',_canEdit:!isWork&&['owner','writer'].includes(cal.accessRole),_calendarName:cal.summary||'Googleカレンダー',updatedAt:g.updated||''};}
async function fetchCalendarEvents(cal,from,to){let next='';const out=[];do{const p=new URLSearchParams({timeMin:from.toISOString(),timeMax:to.toISOString(),singleEvents:'true',showDeleted:'false',maxResults:'1000',orderBy:'startTime'});if(next)p.set('pageToken',next);const res=await api(`${gcalUrl(cal.id)}?${p}`);for(const e of res.items||[]){if(e.status!=='cancelled'&&e.start&&e.end)out.push(eventFromGoogle(e,cal));}next=res.nextPageToken||'';}while(next);return out;}
async function loadCalendarList(){if(state.mode!=='cloud')return;let results=[];let next='';do{const p=new URLSearchParams({maxResults:'250'});if(next)p.set('pageToken',next);const r=await api(`https://www.googleapis.com/calendar/v3/users/me/calendarList?${p}`);results.push(...(r.items||[]));next=r.nextPageToken||'';}while(next);state.allCalendars=results;}
function checkedCalendarIds(){const primary=state.allCalendars.find(c=>c.primary)?.id||'primary';const ids=new Set(config.visibleTouched?config.visibleCalendarIds:[primary]);for(const x of [config.sharedCalId,config.workCalId,config.privateCalId])if(x)ids.add(x);return ids;}
async function loadCloud(){ensureCloud();saveConfig();if(!config.sharedSheetId||!config.privateSheetId)throw Error('初期設定が未完了です。');
 const [a,b]=await Promise.all([sheetRead(config.sharedSheetId,'shared'),sheetRead(config.privateSheetId,'private')]);
 state.records=[...a,...b];state.visible=mergeEvents();render();
}
async function syncGoogleCalendars(){
 ensureCloud();await loadCalendarList();
 const scopes=await Promise.allSettled(['shared','private'].map(reconcileScope));
 const from=new Date(state.month.getFullYear(),state.month.getMonth()-1,1),to=new Date(state.month.getFullYear(),state.month.getMonth()+2,1);
 const listed=state.allCalendars.filter(c=>checkedCalendarIds().has(c.id));
 // A calendar shared via ACL is not always in the partner's calendarList; still read it.
 const calendars=[...listed,...[config.sharedCalId,config.privateCalId].filter(id=>id&&!listed.some(c=>c.id===id)).map(id=>({id,accessRole:'writer',summary:id===config.sharedCalId?'MyLife 共有':'MyLife 個人'}))];
 const received=await Promise.allSettled(calendars.map(c=>fetchCalendarEvents(c,from,to)));
 const retained=state.external.filter(e=>!calendars.some((c,i)=>c.id===e.calendarId&&received[i].status==='fulfilled'));
 state.external=[...retained,...received.filter(r=>r.status==='fulfilled').flatMap(r=>r.value)];
 await loadCloud();
 const errors=[...scopes,...received].filter(x=>x.status==='rejected');
 if(errors.length)throw Error(errors.map(x=>x.reason.message).join('／'));
 if(state.records.some(r=>r.sync&&r.lastError))throw Error('競合または未同期の予定があります。予定を開いて確認してください。');
}

function loadDemo(){state.fetchedCalIds=new Set();state.records=(parseJson(KEYS.demo,[])||[]).filter(x=>!x.deleted).map((x,i)=>({...x,_row:i+2,_store:x.scope,origin:'sheet',kind:x.scope}));state.external=[];state.visible=mergeEvents();render();}
function mergeEvents(){
 const used=new Set(),out=[];const map=new Map(state.external.map(e=>[`${e.calendarId}::${e.googleEventId}`,e]));
 for(const r of state.records){if(r.deleted)continue;const key=`${canonicalCal(r.calendarId)}::${r.googleEventId}`;const g=map.get(key);if(g)used.add(key);
 // Pending local intent always remains visible, including a failed remote delete/update.
 out.push({...r,kind:r.source==='shift'?'shift':r.scope,origin:'sheet',_canEdit:r.source!=='shift'});}
 for(const e of state.external)if(!used.has(`${e.calendarId}::${e.googleEventId}`))out.push(e);return out;
}
async function refresh(){if(state.busy)return;setBusy(true);try{if(state.mode==='cloud')await loadCloud();else loadDemo();return true;}catch(e){showMsg(`予定を読み込めません：${e.message}`,true);return false;}finally{setBusy(false);}}
async function connect(){saveConfig();if(!config.clientId||!config.clientId.endsWith('.apps.googleusercontent.com')){showMsg('Google接続の初期設定が準備中です。今は体験モードを利用できます。',true);openSettings();return;}if(!window.google?.accounts?.oauth2){showMsg('Google認証ライブラリを読み込めません。ネット接続を確認してください。',true);return;}
 const client=google.accounts.oauth2.initTokenClient({client_id:config.clientId,scope:SCOPE,error_callback:error=>showMsg('認証画面を開けませんでした。Safariから開いて接続をお試しください。',true),callback:async response=>{if(response.error||!response.access_token){showMsg('Google認証に失敗しました。',true);return;}state.token=response.access_token;state.expiresAt=Date.now()+(+response.expires_in||3600)*1000;state.mode='cloud';state.protocol=0;state.email='';state.records=[];state.external=[];state.visible=[];render();try{const info=await api('https://www.googleapis.com/oauth2/v3/userinfo');state.email=info.email||'';$('connectionPill').textContent='Google接続中';await bootstrap();syncInputs();if(await refresh())showMsg('接続しました。「すべて同期」で外部予定を反映できます。');}catch(e){showMsg(`接続後の読み込みに失敗：${e.message}`,true);}}});client.requestAccessToken({prompt:'select_account'});
}
function eventsForDay(day){return state.visible.filter(e=>{if(state.filter!=='all'&&e.kind!==state.filter)return false;let start,end;if(e.allDay){start=e.start.slice(0,10);end=e.end.slice(0,10);return day>=start&&day<end;}start=new Date(e.start);end=new Date(e.end);if(!Number.isFinite(+start)||!Number.isFinite(+end))return false;const a=new Date(`${day}T00:00:00+09:00`),b=new Date(`${keyAdd(day,1)}T00:00:00+09:00`);return start<b&&end>a;}).sort((a,b)=>a.start.localeCompare(b.start));}
function eventTime(e){return e.allDay?'終日':`${timeLocal(e.start)}〜${timeLocal(e.end)}`;}
function render(){const m=state.month;$('monthTitle').textContent=monthText(m);const start=new Date(m.getFullYear(),m.getMonth(),1);start.setDate(start.getDate()-start.getDay());const root=$('calendarGrid');root.replaceChildren();
 for(let i=0;i<42;i++){const d=new Date(start);d.setDate(start.getDate()+i);const key=civilKey(d);const list=eventsForDay(key);const cell=document.createElement('div');cell.className=`day${d.getMonth()!==m.getMonth()?' outside':''}${key===state.selected?' selected':''}${key===dateKey(new Date())?' today':''}`;cell.dataset.day=key;const num=document.createElement('span');num.className='daynum';num.textContent=d.getDate();cell.appendChild(num);for(const e of list.slice(0,2)){const chip=document.createElement('div');chip.className=`miniEvent ${e.kind}`;chip.textContent=e.kind==='shift'?`勤務 ${e.title.replace('【勤務】','')}`:e.title;cell.appendChild(chip);}if(list.length>2){const more=document.createElement('div');more.className='more';more.textContent=`＋${list.length-2}件`;cell.appendChild(more);}cell.addEventListener('click',()=>{state.selected=key;if(d.getMonth()!==m.getMonth())state.month=new Date(d.getFullYear(),d.getMonth(),1);render();});root.appendChild(cell);}
 const sel=new Date(`${state.selected}T12:00:00`);$('selectedDayTitle').textContent=`${sel.getMonth()+1}月${sel.getDate()}日（${'日月火水木金土'[sel.getDay()]}）の予定`;
 const dayRoot=$('dayEvents');dayRoot.replaceChildren();const weekStart=keyAdd(state.selected,-new Date(state.selected+'T12:00:00+09:00').getUTCDay());const shownDays=state.view==='week'?Array.from({length:7},(_,i)=>keyAdd(weekStart,i)):[state.selected];const dayEvents=shownDays.flatMap(day=>eventsForDay(day).map(e=>({...e,_shownDay:day})));if(state.view==='week')$('selectedDayTitle').textContent=weekStart+' からの1週間';if(!dayEvents.length){dayRoot.innerHTML='<div class="empty">この日の予定はありません。<br>「＋追加」から登録できます。</div>';}else for(const e of dayEvents){const btn=document.createElement('button');btn.type='button';btn.className=`eventCard ${e.kind}`;btn.innerHTML=`<span class="eventColor"></span><span class="eventContent"><strong>${safe(e.title)}</strong><small>${state.view==='week'?safe(e._shownDay)+' ':''}${safe(eventTime(e))}${e.syncState?.startsWith('pending')?' · 同期待ち':''}${e.lastError?' · 要確認':''}　·　${safe(e.kind==='shift'?'勤務（編集不可）':e.kind==='shared'?'共有':e.kind==='private'?'個人':`Google · ${e._calendarName||''}`)}</small></span><span class="eventArrow">›</span>`;btn.addEventListener('click',()=>openEvent(e));dayRoot.appendChild(btn);}
 renderAlarmPreview();$('connectionPill').textContent=state.mode==='cloud'?'Google接続中':'体験モード';$('greeting').textContent=state.email?`${state.email} で接続中`:'共有も、じぶんだけの予定も。';
}
function getTomorrowShift(){const tomorrow=keyAdd(dateKey(new Date()),1);return state.visible.filter(e=>e.kind==='shift'&&inLocal(e.start)===tomorrow&&!e.allDay).sort((a,b)=>a.start.localeCompare(b.start))[0];}
function renderAlarmPreview(){const shift=getTomorrowShift(),root=$('alarmPreview');if(!shift){root.innerHTML='<div class="alarmNote">明日の勤務シフトはありません（または未取得です）。<br>アラームはiPhoneショートカットが21時に確認します。</div>';return;}const times=alarmTimes(shift);root.innerHTML=`<div class="alarmNote">${safe(shift.title)} · ${safe(timeLocal(shift.start))}開始</div><div class="alarmTime">${times?times.join(' / '):'勤務先の移動時間が未設定'}</div><div class="alarmNote">${times?'朝の準備40分＋通勤時間で計算':'勤務先名と通勤時間の対応をご確認ください。'}<br>※ここでの表示は確認用。実際のアラーム作成はショートカットが行います。</div>`;}
function showDialog(el){if(!el.open)el.showModal();}
function clearEventForm(){for(const el of $('eventForm').querySelectorAll('input,textarea,select'))el.disabled=false;state.editing=null;$('eventForm').reset();$('resolveSync').classList.add('hidden');$('eventId').value='';$('eventOrigin').value='';$('eventScope').disabled=false;$('syncGoogle').disabled=false;$('eventDialogTitle').textContent='予定を追加';$('deleteEvent').classList.add('hidden');$('editHint').classList.add('hidden');$('saveEvent').classList.remove('hidden');$('startDate').value=state.selected;$('endDate').value=state.selected;$('startTime').value='10:00';$('endTime').value='11:00';$('eventScope').value='shared';$('allDay').checked=false;$('syncGoogle').checked=false;allDayToggle();scopeToggle();}
function allDayToggle(){const disabled=$('allDay').checked;$('startTime').disabled=disabled;$('endTime').disabled=disabled;}
function scopeToggle(){$('scopeHelp').textContent=$('eventScope').value==='shared'?'共有予定は、2人が追加・編集・削除できます。':'非共有予定は本人専用Googleシートに保存され、相手には表示されません。';}
function openEvent(e=null){clearEventForm();if(e){state.editing=e;if(e.lastError){$('resolveSync').classList.remove('hidden');$('editHint').classList.remove('hidden');$('editHint').textContent=e.lastError;}$('eventDialogTitle').textContent=e.kind==='shift'?'勤務シフト':e.origin==='google'?'Googleカレンダーの予定':'予定を編集';$('eventId').value=e.id;$('eventOrigin').value=e.origin;$('eventTitle').value=e.title;$('eventScope').value=e.scope==='private'?'private':'shared';$('startDate').value=inLocal(e.start);$('endDate').value=e.allDay?keyAdd(inLocal(e.end),-1):inLocal(e.end);$('startTime').value=timeLocal(e.start)||'10:00';$('endTime').value=timeLocal(e.end)||'11:00';$('allDay').checked=e.allDay;$('eventNotes').value=e.notes||'';$('eventLocation').value=e.location||'';$('syncGoogle').checked=!!e.sync;
 if(e.kind==='shift'||e.origin==='google')$('eventScope').disabled=true;if(e.kind==='shift')$('syncGoogle').disabled=true;
 if(e.kind==='shift'||(e.origin==='google'&&!e._canEdit)){$('editHint').classList.remove('hidden');$('editHint').textContent='この予定は読み取り専用です。変更は元のシステムで行ってください。';for(const el of $('eventForm').querySelectorAll('input,textarea,select'))el.disabled=true;$('saveEvent').classList.add('hidden');}
 else{$('deleteEvent').classList.remove('hidden');if(e.origin==='google'){$('editHint').classList.remove('hidden');$('editHint').textContent='既存のGoogle予定です。変更・削除は保存後、「すべて同期」を押したときに反映されます。';}}
 if(e.origin==='sheet')$('eventScope').disabled=true;
 }else{for(const el of $('eventForm').querySelectorAll('input,textarea,select'))el.disabled=false;}
 allDayToggle();if(e?.kind==='shift'||(e?.origin==='google'&&!e._canEdit)){for(const el of $('eventForm').querySelectorAll('input,textarea,select'))el.disabled=true;}scopeToggle();showDialog($('eventDialog'));}
function buildFromForm(){const allDay=$('allDay').checked;const sd=$('startDate').value,ed=$('endDate').value;if(!sd||!ed)throw Error('開始日と終了日を入力してください。');let start,end;
 if(allDay){start=sd;end=keyAdd(ed,1);if(ed<sd)throw Error('終了日を開始日以降にしてください。');}
 else{start=makeLocalISO(sd,$('startTime').value);end=makeLocalISO(ed,$('endTime').value);if(new Date(end)<=new Date(start))throw Error('終了時刻は開始より後にしてください。');}
 return{title:$('eventTitle').value.trim(),scope:$('eventScope').value,start,end,allDay,notes:$('eventNotes').value.trim(),location:$('eventLocation').value.trim(),sync:$('syncGoogle').checked};}
function targetCal(scope){return scope==='shared'?config.sharedCalId:config.privateCalId;}
function canonicalCal(id){return id==='primary'?(state.allCalendars.find(c=>c.primary)?.id||'primary'):id;}
async function saveEvent(e){e.preventDefault();if(state.busy)return;let draft;try{draft=buildFromForm();if(!draft.title)throw Error('タイトルを入力してください。');}catch(err){showMsg(err.message,true);return;}
 const old=state.editing;if(old?.kind==='shift'||(old?.origin==='google'&&!old._canEdit))return;setBusy(true);
 try{
 if(state.mode==='demo'){
 const items=parseJson(KEYS.demo,[]);if(old){const idx=items.findIndex(x=>x.id===old.id);if(idx<0)throw Error('元の予定が見つかりません。');items[idx]={...items[idx],...draft,updatedAt:new Date().toISOString()};}else items.push({...draft,id:crypto.randomUUID(),owner:'demo',calendarId:'',googleEventId:'',source:'manual',deleted:false,createdAt:new Date().toISOString(),updatedAt:new Date().toISOString()});localStorage.setItem(KEYS.demo,JSON.stringify(items));
 }else if(old?.origin==='google'){
 await stageExternal(old,draft,false);
 }else{
 const r={...old,...draft,id:old?.id||($('eventId').value||crypto.randomUUID())};$('eventId').value=r.id;
 const result=await gateway('save',draft.scope,r,old?.revision||0);
 if(result.record.lastError)showMsg('予定は保存済みです。Google同期は未完了：'+result.record.lastError,true);
 }
 $('eventDialog').close();showMsg('予定を保存しました。');setBusy(false);await refresh();
 }catch(err){showMsg(`保存に失敗しました：${err.message}`,true);}finally{setBusy(false);}
}
async function removeEvent(){const e=state.editing;if(!e||state.busy)return;if(!window.confirm(`「${e.title}」を削除しますか？`))return;setBusy(true);try{
 if(state.mode==='demo'){const items=parseJson(KEYS.demo,[]);localStorage.setItem(KEYS.demo,JSON.stringify(items.filter(x=>x.id!==e.id)));}
 else if(e.origin==='google'){if(!e._canEdit)throw Error('削除権限がありません。');await stageExternal(e,e,true);}
 else{await gateway('delete',e.scope,e,e.revision||0);}
 $('eventDialog').close();showMsg('予定を削除しました。');setBusy(false);await refresh();
 }catch(err){showMsg(`削除に失敗しました：${err.message}`,true);}finally{setBusy(false);}}
function openSettings(){syncInputs();showDialog($('settingsDialog'));renderSyncReport();}
function demoMode(){state.mode='demo';state.token='';state.email='';state.records=[];state.external=[];state.visible=[];$('settingsDialog').close();loadDemo();showMsg('体験モードに切り替えました。データはこの端末だけに保存されます。');}
function nextMonth(diff){state.month=new Date(state.month.getFullYear(),state.month.getMonth()+diff,1);state.selected=civilKey(new Date(state.month.getFullYear(),state.month.getMonth(),1));refresh();}
async function shareApp(){
 const url=new URL(location.pathname,location.origin).href;
 try{if(navigator.share)await navigator.share({title:'MyLife Calendar',text:'ふたりのカレンダー。Googleで接続して利用を開始できます。',url});else if(navigator.clipboard){await navigator.clipboard.writeText(url);showMsg('共有リンクをコピーしました。相手に送ってください。');}else showMsg('アドレス欄のURLをコピーして共有してください。');}catch(err){if(err.name!=='AbortError')showMsg('共有できませんでした。アドレス欄のURLをコピーしてください。',true);}
}
async function resolveSync(){const e=state.editing;if(!e||!e.lastError||state.busy)return;if(!confirm('MyLife側の未同期の変更を取り消し、現在のGoogle側の内容を採用します。Google側が削除済みならMyLifeの内容を同期OFFで保持します。よろしいですか？'))return;setBusy(true);try{await gateway('queueResolve',e.scope,e,e.revision);$('eventDialog').close();setBusy(false);await refresh();}catch(err){showMsg(err.message,true);}finally{setBusy(false);}}
function isSignedIn(){return state.mode==='cloud'&&!!state.token&&Date.now()<state.expiresAt;}
function isGoogleConnected(){return isSignedIn()&&config.googleCalendarEnabled!==false;}
function disconnectGoogle(){config.googleCalendarEnabled=false;localStorage.setItem(KEYS.config,JSON.stringify(config));state.external=[];state.visible=mergeEvents();render();syncInputs();showMsg('Googleカレンダーとの連携を解除しました。保存済み予定は保持しています。');}
function enableGoogle(){config.googleCalendarEnabled=true;localStorage.setItem(KEYS.config,JSON.stringify(config));syncInputs();showMsg('Googleカレンダー連携をONにしました。「すべて同期」を押したときに反映します。');}
async function stageExternal(old,draft,deleted){const scope=old.calendarId===config.sharedCalId?'shared':'private';return gateway('stageExternal',scope,{...draft,id:crypto.randomUUID(),scope,calendarId:old.calendarId,googleEventId:old.googleEventId,googleEtag:old.googleEtag,deleted},0);}

function renderSyncReport(){const report=parseJson(KEYS.syncReport,{});if(!report.at){$('lastSync').textContent='最終同期：未実行';return;}$('lastSync').textContent='最終同期：'+new Intl.DateTimeFormat('ja-JP',{timeZone:TZ,year:'numeric',month:'2-digit',day:'2-digit',hour:'2-digit',minute:'2-digit',hourCycle:'h23'}).format(new Date(report.at));$('syncSummary').textContent=report.summary||'';}
const syncRunner=MyLifeSync.createRunner({
 isGoogleConnected,
 shift:async()=>{ensureCloud();return gateway('syncShift','shared');},
 google:syncGoogleCalendars,
 onBusy:b=>{for(const id of ['syncAllBtn','settingsSyncAll']){$(id).disabled=b||state.busy;$(id).textContent=b?'同期中…':'↻ すべて同期';}},
 onResult:report=>{const label={success:'成功',failed:'失敗',skipped:'未接続'};const shiftValue=report.results.shift.value,shiftLabel=report.results.shift.status==='skipped'?'対象外':label[report.results.shift.status]+(shiftValue&&report.results.shift.status==='success'?`（追加${shiftValue.created}・変更${shiftValue.updated}・削除${shiftValue.deleted}）`:'');const summary=(report.status==='success'?'同期完了':report.status==='partial'?'一部失敗':'同期失敗')+' · 勤務：'+shiftLabel+'／Google：'+(report.results.google.status==='skipped'&&isSignedIn()?'連携OFF':label[report.results.google.status]);localStorage.setItem(KEYS.syncReport,JSON.stringify({at:report.at,status:report.status,summary}));renderSyncReport();$('syncStatus').textContent=summary;const detail=Object.values(report.results).filter(x=>x.status==='failed').map(x=>x.message).join('／');showMsg(summary+(detail?'：'+detail:''),report.status!=='success');refresh();}
});
function bind(){
 $('quickLogin').addEventListener('click',connect);$('shareApp').addEventListener('click',shareApp);
 $('resolveSync').addEventListener('click',resolveSync);
 $('syncAllBtn').addEventListener('click',()=>syncRunner.run());$('settingsSyncAll').addEventListener('click',()=>syncRunner.run());$('disconnectGoogle').addEventListener('click',disconnectGoogle);$('enableGoogle').addEventListener('click',enableGoogle);
 $('viewDay').addEventListener('click',()=>{state.view='day';render();});$('viewWeek').addEventListener('click',()=>{state.view='week';render();});
 $('openSettings').addEventListener('click',openSettings);$('navSettings').addEventListener('click',openSettings);$('closeSettings').addEventListener('click',()=>$('settingsDialog').close());
 $('loginBtn').addEventListener('click',connect);
 $('prevMonth').addEventListener('click',()=>nextMonth(-1));$('nextMonth').addEventListener('click',()=>nextMonth(1));$('todayBtn').addEventListener('click',()=>{state.month=new Date(today.getFullYear(),today.getMonth(),1);state.selected=dateKey(new Date());refresh();});
 $('filters').addEventListener('click',e=>{const b=e.target.closest('button[data-filter]');if(!b)return;state.filter=b.dataset.filter;for(const x of $('filters').querySelectorAll('button'))x.classList.toggle('selected',x===b);render();});
 for(const id of ['addEvent','addSelected','navAdd'])$(id).addEventListener('click',()=>openEvent());$('navToday').addEventListener('click',()=>{state.month=new Date(new Date().getFullYear(),new Date().getMonth(),1);state.selected=dateKey(new Date());refresh();});
 $('eventForm').addEventListener('submit',saveEvent);$('cancelEvent').addEventListener('click',()=>$('eventDialog').close());$('closeEvent').addEventListener('click',()=>$('eventDialog').close());$('deleteEvent').addEventListener('click',removeEvent);$('allDay').addEventListener('change',allDayToggle);$('eventScope').addEventListener('change',scopeToggle);

}
bind();syncInputs();loadDemo();renderSyncReport();if(!config.clientId||!config.gatewayUrl)$('onboardingStatus').textContent='Google接続の初期設定は準備中です。体験モードは利用できます。設定完了後は、共有リンクとGoogleログインで開始できます。';if('serviceWorker'in navigator&&location.protocol==='https:')navigator.serviceWorker.register('./service-worker.js').catch(()=>{});
})();
