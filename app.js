/* MyLife Calendar 1.0 — no backend secrets; Google OAuth token exists in RAM only. */
(() => {
'use strict';
const $ = id => document.getElementById(id);
const KEYS = {config:'mylife.calendar.config.v1',demo:'mylife.calendar.demo.v1'};
const HEADERS=['id','owner','scope','title','start','end','allDay','notes','sync','calendarId','googleEventId','source','deleted','createdAt','updatedAt','location','revision','syncState','googleEtag','lastError'];
const SCOPE='openid email https://www.googleapis.com/auth/spreadsheets https://www.googleapis.com/auth/calendar https://www.googleapis.com/auth/drive.metadata.readonly';
const TZ='Asia/Tokyo';
const defaults={gatewayUrl:'',clientId:'',sharedSheetId:'',privateSheetId:'',sharedCalId:'',privateCalId:'',workCalId:'',visibleCalendarIds:[],visibleTouched:false};
const parseJson=(key,fallback)=>{try{return JSON.parse(localStorage.getItem(key))||fallback;}catch{return fallback;}};
let config={...defaults,...parseJson(KEYS.config,{})};
const {dateKey,keyAdd,makeLocalISO,timeLocal,workTravel,alarmTimes,parseEventText}=MyLifeCore;
const today=new Date(dateKey()+'T12:00:00+09:00');const state={month:new Date(today.getFullYear(),today.getMonth(),1),selected:dateKey(today),filter:'all',token:'',email:'',mode:'demo',records:[],external:[],visible:[],allCalendars:[],fetchedCalIds:new Set(),busy:false,editing:null,image:null,view:'day',expiresAt:0,ocrBusy:false};
const safe=str=>String(str??'').replaceAll('&','&amp;').replaceAll('<','&lt;').replaceAll('>','&gt;').replaceAll('"','&quot;').replaceAll("'",'&#39;');
const pad=n=>String(n).padStart(2,'0');
function civilKey(d){return `${d.getFullYear()}-${pad(d.getMonth()+1)}-${pad(d.getDate())}`;}
function monthText(d){return `${d.getFullYear()}年 ${d.getMonth()+1}月`;}
function inLocal(iso){if(!iso)return '';return /^\d{4}-\d{2}-\d{2}$/.test(iso)?iso:dateKey(iso);}
function showMsg(msg,error=false){const el=$('message');el.textContent=msg;el.classList.toggle('error',error);el.classList.remove('hidden');clearTimeout(showMsg.t);showMsg.t=setTimeout(()=>el.classList.add('hidden'),7000);}
function saveConfig(){for(const k of Object.keys(defaults)){if($(k))config[k]=$(k).value.trim();}localStorage.setItem(KEYS.config,JSON.stringify(config));}
function syncInputs(){for(const k of Object.keys(defaults)){if($(k))$(k).value=config[k]||'';} $('loggedEmail').textContent=state.email?`接続中：${state.email}（Google認可は一定時間後に更新が必要です）`:'Googleには未接続です。';}
function setBusy(b){state.busy=b;$('saveEvent').disabled=b;$('deleteEvent').disabled=b;$('refreshBtn').disabled=b;}
function ensureCloud(){if(state.mode!=='cloud'||!state.token||Date.now()>=state.expiresAt)throw Error('Googleで接続してください。');}
async function api(url,init={}){
 ensureCloud();const res=await fetch(url,{...init,headers:{'Authorization':`Bearer ${state.token}`,...(init.body?{'Content-Type':'application/json'}:{}),...(init.headers||{})}});
 const txt=await res.text();let obj={};try{obj=txt?JSON.parse(txt):{};}catch{obj={message:txt.slice(0,150)};}
 if(!res.ok){const desc=obj.error?.message||obj.message||res.statusText;throw Error(`${desc}（HTTP ${res.status}）${res.status===401?'。Googleに再接続してください。':''}`);}
 return obj;
}
const sheetsRoot=id=>`https://sheets.googleapis.com/v4/spreadsheets/${encodeURIComponent(id)}`;
const a1=range=>encodeURIComponent(range).replaceAll('%21','!');
function normalizeRow(v,rowNo,store){const o={};HEADERS.forEach((key,i)=>o[key]=v[i]??'');o.allDay=o.allDay==='true'||o.allDay===true;o.sync=o.sync==='true'||o.sync===true;o.deleted=o.deleted==='true'||o.deleted===true;o.revision=+o.revision||0;o._row=rowNo;o._store=store;o.origin='sheet';o.kind=o.scope==='shared'?'shared':'private';return o;}
function rowValues(r){return HEADERS.map(k=>['allDay','sync','deleted'].includes(k)?String(!!r[k]):String(r[k]??''));}
async function sheetRead(id,store){if(!id)return [];const data=await api(`${sheetsRoot(id)}/values/${a1('events!A:T')}?valueRenderOption=FORMATTED_VALUE`);return(data.values||[]).slice(1).map((v,i)=>normalizeRow(v,i+2,store)).filter(e=>e.id&&(!e.deleted||e.lastError));}
async function gateway(action,scope,record,expectedRevision=0){
 ensureCloud();if(!/^https:\/\/script\.google\.com\/macros\/s\/[A-Za-z0-9_-]+\/exec$/.test(config.gatewayUrl))throw Error('設定にGASのウェブアプリURLが必要です。');
 const res=await fetch(config.gatewayUrl,{method:'POST',headers:{'Content-Type':'text/plain;charset=utf-8'},body:JSON.stringify({token:state.token,action,scope,record,expectedRevision}),redirect:'follow',signal:AbortSignal.timeout(90000)});
 if(!res.ok)throw Error('同期サーバーに接続できません。');const data=await res.json();if(!data.ok)throw Error(data.error||'同期サーバーでエラーが発生しました。');return data;
}
async function reconcileScope(scope){let afterId='';do{const r=await gateway('reconcile',scope,{afterId});afterId=r.nextAfter||'';}while(afterId);}
async function bootstrap(){const data=await gateway('bootstrap');Object.assign(config,data.config);state.email=data.email;syncInputs();}
async function makeSheet(kind){saveConfig();ensureCloud();const title=kind==='shared'?'MyLife_共有予定_DB':'MyLife_個人予定_DB';const doc=await api('https://sheets.googleapis.com/v4/spreadsheets',{method:'POST',body:JSON.stringify({properties:{title,locale:'ja_JP',timeZone:TZ},sheets:[{properties:{title:'events',gridProperties:{frozenRowCount:1}}}]})});
 await api(`${sheetsRoot(doc.spreadsheetId)}/values/${a1('events!A1:T1')}?valueInputOption=RAW`,{method:'PUT',body:JSON.stringify({values:[HEADERS]})});
 const key=kind==='shared'?'sharedSheetId':'privateSheetId';config[key]=doc.spreadsheetId;localStorage.setItem(KEYS.config,JSON.stringify(config));syncInputs();showMsg('シートを作成しました。共有用の場合はGoogleスプレッドシートで相手に編集権限を付与してください。');window.open(doc.spreadsheetUrl||`https://docs.google.com/spreadsheets/d/${doc.spreadsheetId}/edit`,'_blank','noopener');
}
async function makeCal(kind){saveConfig();ensureCloud();const names={shared:'MyLife 共有',private:'MyLife 個人',work:'MyLife 勤務'};const cal=await api('https://www.googleapis.com/calendar/v3/calendars',{method:'POST',body:JSON.stringify({summary:names[kind],timeZone:TZ,description:kind==='work'?'勤務シフトシートからの同期専用（MyLife側では読み取りのみ）':'MyLife Calendar'})});
 const key={shared:'sharedCalId',private:'privateCalId',work:'workCalId'}[kind];config[key]=cal.id;localStorage.setItem(KEYS.config,JSON.stringify(config));syncInputs();showMsg(`「${names[kind]}」を作成しました。${kind==='shared'?'相手に予定変更権限を付与してください。':' '} `);await loadCalendarList();
}
function toGoogleEvent(r){const o={summary:r.title,location:r.location||'',description:r.notes||'',start:r.allDay?{date:r.start}:{dateTime:r.start,timeZone:TZ},end:r.allDay?{date:r.end}:{dateTime:r.end,timeZone:TZ},extendedProperties:{shared:{mylifeId:r.id}}};return o;}
const gcalUrl=(cal,event='')=>`https://www.googleapis.com/calendar/v3/calendars/${encodeURIComponent(cal)}/events${event?'/'+encodeURIComponent(event):''}`;
function eventFromGoogle(g,cal){const isWork=cal.id===config.workCalId;return {id:`google:${cal.id}:${g.id}`,owner:g.creator?.email||'',title:g.summary||'(タイトルなし)',scope:'google',start:g.start?.dateTime||g.start?.date||'',end:g.end?.dateTime||g.end?.date||'',allDay:!!g.start?.date,notes:g.description||'',location:g.location||'',googleEtag:g.etag, sync:true,calendarId:cal.id,googleEventId:g.id,source:'google',deleted:false,origin:'google',kind:isWork?'shift':'google',_canEdit:!isWork&&['owner','writer'].includes(cal.accessRole),_calendarName:cal.summary||'Googleカレンダー',updatedAt:g.updated||''};}
async function fetchCalendarEvents(cal,from,to){let next='';const out=[];do{const p=new URLSearchParams({timeMin:from.toISOString(),timeMax:to.toISOString(),singleEvents:'true',showDeleted:'false',maxResults:'1000',orderBy:'startTime'});if(next)p.set('pageToken',next);const res=await api(`${gcalUrl(cal.id)}?${p}`);for(const e of res.items||[]){if(e.status!=='cancelled'&&e.start&&e.end)out.push(eventFromGoogle(e,cal));}next=res.nextPageToken||'';}while(next);return out;}
async function loadCalendarList(){if(state.mode!=='cloud')return;let results=[];let next='';do{const p=new URLSearchParams({maxResults:'250'});if(next)p.set('pageToken',next);const r=await api(`https://www.googleapis.com/calendar/v3/users/me/calendarList?${p}`);results.push(...(r.items||[]));next=r.nextPageToken||'';}while(next);state.allCalendars=results;renderCalendarChecks();}
function checkedCalendarIds(){const primary=state.allCalendars.find(c=>c.primary)?.id||'primary';const ids=new Set(config.visibleTouched?config.visibleCalendarIds:[primary]);for(const x of [config.sharedCalId,config.workCalId,config.privateCalId])if(x)ids.add(x);return ids;}
function renderCalendarChecks(){const root=$('calendarChecks');root.replaceChildren();const selected=checkedCalendarIds();for(const c of state.allCalendars){if(c.id===config.workCalId)continue;const label=document.createElement('label');label.className='checkItem';const input=document.createElement('input');input.type='checkbox';input.checked=selected.has(c.id);input.addEventListener('change',()=>{const ids=new Set(config.visibleTouched?config.visibleCalendarIds:[state.allCalendars.find(c=>c.primary)?.id||'primary']);if(input.checked)ids.add(c.id);else ids.delete(c.id);config.visibleCalendarIds=[...ids];config.visibleTouched=true;localStorage.setItem(KEYS.config,JSON.stringify(config));});const txt=document.createElement('span');txt.textContent=c.summary||c.id;label.append(input,txt);root.appendChild(label);}if(!state.allCalendars.length)root.textContent='カレンダーが見つかりません。';}
async function loadCloud(){ensureCloud();saveConfig();await bootstrap();await reconcileScope('shared');await reconcileScope('private');if(!config.sharedSheetId||!config.privateSheetId)throw Error('「設定」で共有シートと本人専用シートを作成・指定してください。');
 const from=new Date(state.month.getFullYear(),state.month.getMonth()-1,1);const to=new Date(state.month.getFullYear(),state.month.getMonth()+2,1);
 const [a,b]=await Promise.all([sheetRead(config.sharedSheetId,'shared'),sheetRead(config.privateSheetId,'private')]);
 state.records=[...a,...b];let calendars=state.allCalendars.filter(c=>checkedCalendarIds().has(c.id));
 // Primary is listed with an actual email ID, not necessarily literal "primary".
 if(!config.visibleTouched){const primary=state.allCalendars.find(c=>c.primary);if(primary&&!calendars.some(c=>c.id===primary.id))calendars.push(primary);}
 if(config.workCalId&&!calendars.some(c=>c.id===config.workCalId))calendars.push({id:config.workCalId,summary:'MyLife 勤務',accessRole:'reader'});
 const successes=await Promise.allSettled(calendars.map(c=>fetchCalendarEvents(c,from,to)));state.fetchedCalIds=new Set(calendars.filter((c,i)=>successes[i].status==='fulfilled').map(c=>c.id));state.external=successes.filter(x=>x.status==='fulfilled').flatMap(x=>x.value);
 const fails=successes.filter(x=>x.status==='rejected');if(fails.length)showMsg(`一部のGoogleカレンダーを読み込めませんでした：${fails[0].reason.message}`,true);
 state.visible=mergeEvents();render();const errors=state.records.filter(r=>r.lastError);$('syncStatus').replaceChildren();if(!errors.length)$('syncStatus').textContent='同期完了';for(const r of errors){const b=document.createElement('button');b.type='button';b.className='secondary';b.textContent=`${r.title}：同期を確認`;b.addEventListener('click',()=>openEvent(r));$('syncStatus').appendChild(b);}
}
function loadDemo(){state.fetchedCalIds=new Set();state.records=(parseJson(KEYS.demo,[])||[]).filter(x=>!x.deleted).map((x,i)=>({...x,_row:i+2,_store:x.scope,origin:'sheet',kind:x.scope}));state.external=[];state.visible=mergeEvents();render();}
function mergeEvents(){
 const used=new Set(),out=[];const map=new Map(state.external.map(e=>[`${e.calendarId}::${e.googleEventId}`,e]));
 for(const r of state.records){if(r.deleted)continue;const key=`${canonicalCal(r.calendarId)}::${r.googleEventId}`;const g=map.get(key);if(g)used.add(key);
 // Pending local intent always remains visible, including a failed remote delete/update.
 out.push({...r,kind:r.scope,origin:'sheet',_canEdit:true});}
 for(const e of state.external)if(!used.has(`${e.calendarId}::${e.googleEventId}`))out.push(e);return out;
}
async function refresh(){if(state.busy)return;setBusy(true);try{if(state.mode==='cloud')await loadCloud();else loadDemo();return true;}catch(e){showMsg(`同期できません：${e.message}`,true);return false;}finally{setBusy(false);}}
async function connect(){saveConfig();if(!config.clientId||!config.clientId.endsWith('.apps.googleusercontent.com')){showMsg('Google OAuth クライアントIDを入力してください。',true);return;}if(!window.google?.accounts?.oauth2){showMsg('Google認証ライブラリを読み込めません。ネット接続を確認してください。',true);return;}
 const client=google.accounts.oauth2.initTokenClient({client_id:config.clientId,scope:SCOPE,error_callback:error=>showMsg('認証画面を開けませんでした。Safariから開いて接続をお試しください。',true),callback:async response=>{if(response.error||!response.access_token){showMsg('Google認証に失敗しました。',true);return;}state.token=response.access_token;state.expiresAt=Date.now()+(+response.expires_in||3600)*1000;state.mode='cloud';state.records=[];state.external=[];state.visible=[];render();try{const info=await api('https://www.googleapis.com/oauth2/v3/userinfo');state.email=info.email||'';$('connectionPill').textContent='Google接続中';syncInputs();await loadCalendarList();if(await refresh())showMsg('Googleへ接続しました。');}catch(e){showMsg(`接続後の読み込みに失敗：${e.message}`,true);}}});client.requestAccessToken({prompt:'select_account'});
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
 if(e.kind==='shift'||e.origin==='google'){$('eventScope').disabled=true;$('syncGoogle').disabled=true;}
 if(e.kind==='shift'||(e.origin==='google'&&!e._canEdit)){$('editHint').classList.remove('hidden');$('editHint').textContent='この予定は読み取り専用です。変更は元のシステムで行ってください。';for(const el of $('eventForm').querySelectorAll('input,textarea,select'))el.disabled=true;$('saveEvent').classList.add('hidden');}
 else{$('deleteEvent').classList.remove('hidden');if(e.origin==='google'){$('editHint').classList.remove('hidden');$('editHint').textContent='既存のGoogle予定です。変更・削除はGoogleカレンダーに直接反映されます。';}}
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
 await api(gcalUrl(old.calendarId,old.googleEventId),{method:'PATCH',headers:{'If-Match':old.googleEtag},body:JSON.stringify(toGoogleEvent({...old,...draft}))});
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
 else if(e.origin==='google'){if(!e._canEdit)throw Error('削除権限がありません。');await api(gcalUrl(e.calendarId,e.googleEventId),{method:'DELETE',headers:{'If-Match':e.googleEtag}});}
 else{await gateway('delete',e.scope,e,e.revision||0);}
 $('eventDialog').close();showMsg('予定を削除しました。');setBusy(false);await refresh();
 }catch(err){showMsg(`削除に失敗しました：${err.message}`,true);}finally{setBusy(false);}}
function openSettings(){syncInputs();showDialog($('settingsDialog'));if(state.mode==='cloud')renderCalendarChecks();}
function demoMode(){state.mode='demo';state.token='';state.email='';state.records=[];state.external=[];state.visible=[];$('settingsDialog').close();loadDemo();showMsg('体験モードに切り替えました。データはこの端末だけに保存されます。');}
function nextMonth(diff){state.month=new Date(state.month.getFullYear(),state.month.getMonth()+diff,1);state.selected=civilKey(new Date(state.month.getFullYear(),state.month.getMonth(),1));refresh();}
async function recognizeImage(){if(!state.image||state.ocrBusy)return;state.ocrBusy=true;$('ocrState').textContent='日本語OCRの準備・解析中です。初回は言語データの読み込みが必要です。';$('retryOcr').disabled=true;try{
 if(!window.Tesseract){await new Promise((resolve,reject)=>{const s=document.createElement('script');s.src='https://cdn.jsdelivr.net/npm/tesseract.js@5/dist/tesseract.min.js';s.onload=resolve;s.onerror=()=>reject(Error('OCRライブラリを読み込めませんでした。'));document.head.appendChild(s);});}
 const worker=await Tesseract.createWorker('jpn+eng');try{const out=await worker.recognize(state.image);$('recognizedText').value=out.data.text||'';$('ocrState').textContent='認識が終わりました。内容を修正して「予定入力に反映」を押してください。';}finally{await worker.terminate();}
 }catch(e){$('ocrState').textContent='OCRに失敗しました。画像の文字は下欄に手入力して続行できます。詳細：'+e.message;}finally{state.ocrBusy=false;$('retryOcr').disabled=false;}}
function useOcr(){const content=$('recognizedText').value.trim();if(!content){showMsg('読み取った文章を入力してください。',true);return;}const extracted=parseEventText(content);$('imageDialog').close();openEvent();$('eventTitle').value=extracted.title;if(extracted.date&&Number.isFinite(+new Date(extracted.date+'T12:00:00'))){$('startDate').value=extracted.date;$('endDate').value=extracted.date;}if(extracted.start)$('startTime').value=extracted.start;if(extracted.end)$('endTime').value=extracted.end;else if(extracted.start){const [h,m]=extracted.start.split(':').map(Number);const n=h*60+m+60;$('endTime').value=`${pad(Math.floor(n/60)%24)}:${pad(n%60)}`;if(n>=1440)$('endDate').value=keyAdd($('startDate').value,1);}$('eventNotes').value=extracted.notes;showMsg('認識した候補を入力しました。日時と内容を必ず確認してください。');}
function handleImage(file){if(!file)return;if(state.ocrBusy){showMsg('解析が完了してから次の画像を選択してください。');return;}if(file.size>20*1024*1024){showMsg('画像は20MB以下にしてください。',true);return;}if(!file.type.startsWith('image/')){showMsg('画像ファイルを選択してください。',true);return;}state.image=file;const object=URL.createObjectURL(file);$('imagePreview').src=object;$('imagePreview').onload=()=>URL.revokeObjectURL(object);$('recognizedText').value='';$('ocrState').textContent='文字認識を開始します。';showDialog($('imageDialog'));recognizeImage();}
async function resolveSync(){const e=state.editing;if(!e||!e.lastError||state.busy)return;if(!confirm('MyLife側の未同期の変更を取り消し、現在のGoogle側の内容を採用します。Google側が削除済みならMyLifeの内容を同期OFFで保持します。よろしいですか？'))return;setBusy(true);try{await gateway('resolve',e.scope,e,e.revision);$('eventDialog').close();setBusy(false);await refresh();}catch(err){showMsg(err.message,true);}finally{setBusy(false);}}
function bind(){
 $('resolveSync').addEventListener('click',resolveSync);
 $('viewDay').addEventListener('click',()=>{state.view='day';render();});$('viewWeek').addEventListener('click',()=>{state.view='week';render();});
 document.addEventListener('visibilitychange',()=>{if(document.visibilityState==='visible'&&state.mode==='cloud'&&!state.busy)refresh();});
 setInterval(()=>{if(document.visibilityState==='visible'&&state.mode==='cloud'&&!state.busy&&Date.now()<state.expiresAt)refresh();},300000);
 $('openSettings').addEventListener('click',openSettings);$('navSettings').addEventListener('click',openSettings);$('closeSettings').addEventListener('click',()=>$('settingsDialog').close());$('closeAndSaveSettings').addEventListener('click',()=>{saveConfig();$('settingsDialog').close();refresh();showMsg('設定を保存しました。');});
 $('loginBtn').addEventListener('click',connect);$('demoBtn').addEventListener('click',demoMode);
 $('createSharedSheet').addEventListener('click',()=>makeSheet('shared').catch(e=>showMsg(e.message,true)));$('createPrivateSheet').addEventListener('click',()=>makeSheet('private').catch(e=>showMsg(e.message,true)));
 for(const [id,kind] of [['createSharedCal','shared'],['createPrivateCal','private'],['createWorkCal','work']])$(id).addEventListener('click',()=>makeCal(kind).catch(e=>showMsg(e.message,true)));
 $('copyShiftInstructions').addEventListener('click',()=>{showMsg('ZIP内の「README_設定手順.md」と「shift-sync/Code.gs」をご覧ください。');});$('openShiftInfo').addEventListener('click',openSettings);
 $('prevMonth').addEventListener('click',()=>nextMonth(-1));$('nextMonth').addEventListener('click',()=>nextMonth(1));$('todayBtn').addEventListener('click',()=>{state.month=new Date(today.getFullYear(),today.getMonth(),1);state.selected=dateKey(new Date());refresh();});$('refreshBtn').addEventListener('click',refresh);
 $('filters').addEventListener('click',e=>{const b=e.target.closest('button[data-filter]');if(!b)return;state.filter=b.dataset.filter;for(const x of $('filters').querySelectorAll('button'))x.classList.toggle('selected',x===b);render();});
 for(const id of ['addEvent','addSelected','navAdd'])$(id).addEventListener('click',()=>openEvent());$('navToday').addEventListener('click',()=>{state.month=new Date(new Date().getFullYear(),new Date().getMonth(),1);state.selected=dateKey(new Date());refresh();});
 $('eventForm').addEventListener('submit',saveEvent);$('cancelEvent').addEventListener('click',()=>$('eventDialog').close());$('closeEvent').addEventListener('click',()=>$('eventDialog').close());$('deleteEvent').addEventListener('click',removeEvent);$('allDay').addEventListener('change',allDayToggle);$('eventScope').addEventListener('change',scopeToggle);
 $('imageFile').addEventListener('change',e=>{handleImage(e.target.files[0]);e.target.value='';});$('closeImage').addEventListener('click',()=>$('imageDialog').close());$('retryOcr').addEventListener('click',recognizeImage);$('useOcr').addEventListener('click',useOcr);
}
bind();syncInputs();loadDemo();if('serviceWorker'in navigator&&location.protocol==='https:')navigator.serviceWorker.register('./service-worker.js').catch(()=>{});
})();
