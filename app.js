/* MyLife Calendar 1.0 — no backend secrets; Google OAuth token exists in RAM only. */
(() => {
'use strict';
const $ = id => document.getElementById(id);
const KEYS = {config:'mylife.calendar.config.v1',demo:'mylife.calendar.demo.v1',syncReport:'mylife.calendar.syncReport.v1',cache:'mylife.calendar.cache.v1',auth:'mylife.calendar.auth.v1',lastError:'mylife.calendar.lastError.v1'};
const HEADERS=['id','owner','scope','title','start','end','allDay','notes','sync','calendarId','googleEventId','source','deleted','createdAt','updatedAt','location','revision','syncState','googleEtag','lastError'];
const SCOPE='openid email https://www.googleapis.com/auth/spreadsheets https://www.googleapis.com/auth/calendar https://www.googleapis.com/auth/drive.metadata.readonly';
const TZ='Asia/Tokyo';
const defaults={gatewayUrl:window.MYLIFE_DEPLOYMENT?.gatewayUrl||'',clientId:window.MYLIFE_DEPLOYMENT?.clientId||'',sharedSheetId:'',privateSheetId:'',sharedCalId:'',privateCalId:'',workCalId:'',visibleCalendarIds:[],visibleTouched:false,googleCalendarEnabled:true};
const parseJson=(key,fallback)=>{try{return JSON.parse(localStorage.getItem(key))||fallback;}catch{return fallback;}};
let config={...defaults,...parseJson(KEYS.config,{})};for(const k of ['clientId','gatewayUrl'])if(window.MYLIFE_DEPLOYMENT?.[k])config[k]=window.MYLIFE_DEPLOYMENT[k];
const {dateKey,keyAdd,makeLocalISO,timeLocal,workTravel,alarmTimes}=MyLifeCore;
const today=new Date(dateKey()+'T12:00:00+09:00');const state={month:new Date(today.getFullYear(),today.getMonth(),1),selected:dateKey(today),filter:'private',token:'',email:'',mode:'demo',records:[],external:[],visible:[],allCalendars:[],fetchedCalIds:new Set(),busy:false,editing:null,expiresAt:0,weekStart:'',weekCount:78,index:new Map()};
// The calendar is one continuous list of weeks starting on Monday; start three months back and extend on scroll.
state.weekStart=(k=>keyAdd(k,-((new Date(k+'T12:00:00+09:00').getUTCDay()+6)%7)))(dateKey(new Date(today.getFullYear(),today.getMonth()-3,1)));
const safe=str=>String(str??'').replaceAll('&','&amp;').replaceAll('<','&lt;').replaceAll('>','&gt;').replaceAll('"','&quot;').replaceAll("'",'&#39;');
const pad=n=>String(n).padStart(2,'0');
function civilKey(d){return `${d.getFullYear()}-${pad(d.getMonth()+1)}-${pad(d.getDate())}`;}
function monthText(d){return `${d.getFullYear()}年 ${d.getMonth()+1}月`;}
function inLocal(iso){if(!iso)return '';return /^\d{4}-\d{2}-\d{2}$/.test(iso)?iso:dateKey(iso);}
// Only problems are shown; routine 保存/同期 confirmations stay silent.
function showMsg(msg,error=false){const el=$('message');if(!error){el.classList.add('hidden');return;}el.textContent=msg;el.classList.toggle('error',error);el.classList.remove('hidden');clearTimeout(showMsg.t);showMsg.t=setTimeout(()=>el.classList.add('hidden'),7000);}
function saveConfig(){for(const k of Object.keys(defaults)){if($(k))config[k]=$(k).value.trim();}localStorage.setItem(KEYS.config,JSON.stringify(config));}
function syncInputs(){for(const k of Object.keys(defaults)){if($(k))$(k).value=config[k]||'';} const signedIn=isSignedIn();$('loggedEmail').textContent=signedIn?`ログイン中：${state.email}（MyLifeの予定保存・勤務取得に使用）`:'未ログイン';$('calendarLink').textContent='MyLifeの予定はGoogleカレンダーには送りません。同期は勤務シフトの取り込みだけです。';$('loginBtn').classList.toggle('hidden',signedIn);$('disconnectGoogle').classList.add('hidden');$('enableGoogle').classList.add('hidden');}
function setBusy(b){state.busy=b;$('saveEvent').disabled=b;$('deleteEvent').disabled=b;for(const id of ['settingsSyncAll'])$(id).disabled=b||!!syncRunner?.busy;}
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
 state.records=[...a,...b];state.visible=mergeEvents();render();saveCache();
}
async function syncGoogleCalendars(){
 ensureCloud();await loadCalendarList();
 const scopes=await Promise.allSettled(['shared','private'].map(reconcileScope));
 const from=new Date(state.weekStart+'T00:00:00+09:00'),to=new Date(keyAdd(state.weekStart,state.weekCount*7)+'T00:00:00+09:00');
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
 // A shift memo is the viewer's own private row keyed by the shift id; it is shown inside the shift, never as its own event.
 const memos=new Map(state.records.filter(r=>r.source==='shift-memo'&&r._store==='private').map(r=>[r.id,r]));
 for(const r of state.records){if(r.source==='shift-memo')continue;const key=`${canonicalCal(r.calendarId)}::${r.googleEventId}`;const g=map.get(key);if(g)used.add(key);if(r.deleted)continue;/* a pending delete also hides its Google copy */
 // Pending local intent always remains visible, including a failed remote delete/update.
 const memo=r.source==='shift'?memos.get(r.id):null;
 out.push({...r,kind:r.source==='shift'?'shift':r.scope,origin:'sheet',_canEdit:r.source!=='shift',_memo:memo?.notes||'',_memoRev:memo?.revision||0});}
 for(const e of state.external)if(!used.has(`${e.calendarId}::${e.googleEventId}`))out.push(e);return out;
}
/** Put one saved/provisional row into the local list and redraw, instead of re-reading both sheets after every save. */
// The last loaded MyLife data stays on this phone, so reopening the app shows it without signing in or syncing again.
function saveCache(){if(state.mode!=='cloud'||!state.email)return;try{localStorage.setItem(KEYS.cache,JSON.stringify({email:state.email,at:new Date().toISOString(),records:state.records.filter(r=>r.syncState!=='saving')}));}catch{}}
function loadCached(){const c=parseJson(KEYS.cache,null);if(!c?.email||!Array.isArray(c.records)||!(config.clientId&&config.gatewayUrl))return false;state.mode='cloud';state.email=c.email;state.records=c.records;state.external=[];state.visible=mergeEvents();render();return true;}
function upsertLocal(rec,store){const n={...rec,_store:store,origin:'sheet',kind:rec.source==='shift'?'shift':rec.scope==='shared'?'shared':'private'};const i=state.records.findIndex(x=>x.id===n.id&&x._store===store);if(i>=0)state.records[i]={...state.records[i],...n};else state.records.push(n);state.visible=mergeEvents();render();saveCache();}
// The last problem stays visible in settings so it can be read (and shown in a screenshot) after the red message fades.
function showLastError(){const e=parseJson(KEYS.lastError,null);if($('lastError'))$('lastError').textContent=e?.msg?`最後のエラー（${new Intl.DateTimeFormat('ja-JP',{timeZone:TZ,month:'numeric',day:'numeric',hour:'2-digit',minute:'2-digit',hourCycle:'h23'}).format(new Date(e.at))}）：${e.msg}`:'';}
function noteError(msg){try{localStorage.setItem(KEYS.lastError,JSON.stringify({at:new Date().toISOString(),msg}));}catch{}showLastError();}
function clearLastError(){try{localStorage.removeItem(KEYS.lastError);}catch{}showLastError();}
// Keep the Google access token for its own lifetime (about 1 hour) so reopening the app soon after does not ask again.
function saveAuth(){try{localStorage.setItem(KEYS.auth,JSON.stringify({token:state.token,expiresAt:state.expiresAt,email:state.email}));}catch{}}
function clearAuth(){try{localStorage.removeItem(KEYS.auth);}catch{}}
function restoreAuth(){const a=parseJson(KEYS.auth,null);if(!a?.token||!(a.expiresAt>Date.now()+60000)||!(config.clientId&&config.gatewayUrl))return false;state.token=a.token;state.expiresAt=a.expiresAt;state.email=a.email||state.email;state.mode='cloud';return true;}
async function refresh(){if(state.busy)return;setBusy(true);try{if(state.mode==='cloud')await loadCloud();else loadDemo();return true;}catch(e){showMsg(`予定を読み込めません：${e.message}`,true);noteError(`予定を読み込めません：${e.message}`);return false;}finally{setBusy(false);}}
async function connect(after){saveConfig();if(!config.clientId||!config.clientId.endsWith('.apps.googleusercontent.com')){showMsg('Google接続の初期設定が準備中です。今は体験モードを利用できます。',true);openSettings();return;}if(!window.google?.accounts?.oauth2){showMsg('Google認証ライブラリを読み込めません。ネット接続を確認してください。',true);return;}
 const client=google.accounts.oauth2.initTokenClient({client_id:config.clientId,scope:SCOPE,error_callback:error=>showMsg('認証画面を開けませんでした。Safariから開いて接続をお試しください。',true),callback:async response=>{if(response.error||!response.access_token){showMsg('Google認証に失敗しました。',true);noteError('Google認証に失敗しました。');return;}
 // iPhone shows one checkbox per permission; an unchecked one lets sign-in succeed but breaks loading later.
 const need=SCOPE.split(' ').filter(s=>s.startsWith('https://'));if(google.accounts.oauth2.hasGrantedAllScopes&&!google.accounts.oauth2.hasGrantedAllScopes(response,...need)){const msg='Googleの許可画面で、すべての項目にチェックを入れて許可してください。設定の「Googleアカウントに接続」からやり直せます。';showMsg(msg,true);noteError(msg);return;}
 state.token=response.access_token;state.expiresAt=Date.now()+(+response.expires_in||3600)*1000;state.mode='cloud';state.protocol=0;state.email='';/* keep showing the cached calendar until fresh data arrives */try{const info=await api('https://www.googleapis.com/oauth2/v3/userinfo');state.email=info.email||'';$('connectionPill').textContent='Google接続中';await bootstrap();saveAuth();syncInputs();if(await refresh()){clearLastError();if(typeof after==='function')await after();}}catch(e){showMsg(`接続後の読み込みに失敗：${e.message}`,true);noteError(`接続後の読み込みに失敗：${e.message}`);}}});
 // Reconnecting a known account skips the account chooser.
 client.requestAccessToken(state.email?{prompt:'',login_hint:state.email}:{prompt:'select_account'});
}
// 個人 = everything that is mine (my private, my shared, my shifts, my Google). 共有 = every shared event, mine and the partner's.
function isMine(e){if(!state.email)return true;return e.scope==='private'||String(e.owner||'').toLowerCase()===state.email.toLowerCase();}
function isPartner(e){return e.scope==='shared'&&!isMine(e);}
function inFilter(e){return state.filter==='shared'?e.scope==='shared':isMine(e);}
/** Days (JST keys) an event covers; all-day end is exclusive. */
function dayKeysOf(e){let a,b;if(e.allDay){a=String(e.start).slice(0,10);b=keyAdd(String(e.end).slice(0,10),-1);}else{const s=new Date(e.start),t=new Date(e.end);if(!Number.isFinite(+s)||!Number.isFinite(+t))return [];a=dateKey(s);b=dateKey(new Date(Math.max(+s,+t-1)));}if(!a||b<a)b=a;const out=[];for(let k=a,i=0;k<=b&&i<370;k=keyAdd(k,1),i++)out.push(k);return out;}
// One pass per render: day → events for the current 個人/共有 filter (rendering hundreds of days must not rescan every event per day).
function buildIndex(){const idx=new Map();for(const e of state.visible){if(!inFilter(e))continue;e._days=dayKeysOf(e);for(const k of e._days){if(!idx.has(k))idx.set(k,[]);idx.get(k).push(e);}}for(const l of idx.values())l.sort((a,b)=>(b.scope==='shared')-(a.scope==='shared')||String(a.start).localeCompare(String(b.start)));state.index=idx;}
function eventsForDay(day){return state.index.get(day)||[];}
function eventCard(e,showDay,onOpen){const btn=document.createElement('button');btn.type='button';btn.className=`eventCard ${e.kind}${isPartner(e)?' partner':''}`;btn.innerHTML=`<span class="eventColor"></span><span class="eventContent"><strong>${safe(e.title)}</strong><small>${showDay?safe(e._shownDay)+' ':''}${safe(eventTime(e))}${e.syncState==='saving'?' · 保存中':''}${e.lastError?' · 要確認':''}${e._memo?' · メモあり':''}　·　${isPartner(e)?'相手の予定 · ':''}${safe(e.kind==='shift'?'勤務（編集不可）':e.kind==='shared'?'共有':e.kind==='private'?'個人':`Google · ${e._calendarName||''}`)}</small></span><span class="eventArrow">›</span>`;btn.addEventListener('click',()=>onOpen(e));return btn;}
const dayLabel=key=>{const d=new Date(`${key}T12:00:00`);return `${d.getMonth()+1}月${d.getDate()}日（${'日月火水木金土'[d.getDay()]}）`;};
/** Tapping a date pops up that day's events (current 個人/共有 filter), with add and previous/next day. */
function markSelected(){for(const row of $('calendarGrid').children)for(const c of row.children||[])if(c.dataset?.day)c.className=c.className.replace(/ selected/g,'')+(c.dataset.day===state.selected?' selected':'');}
function openDay(key){state.selected=key;markSelected();
 $('dayDialogTitle').textContent=dayLabel(key)+'の予定';const root=$('dayDialogEvents');root.replaceChildren();const list=eventsForDay(key);
 if(!list.length)root.innerHTML='<div class="empty">この日の予定はありません。</div>';else for(const e of list)root.appendChild(eventCard(e,false,x=>{$('dayDialog').close();openEvent(x);}));
 showDialog($('dayDialog'));}
function eventTime(e){return e.allDay?'終日':`${timeLocal(e.start)}〜${timeLocal(e.end)}`;}
const LANES=5;
function render(){buildIndex();const root=$('calendarGrid');root.replaceChildren();const todayKey=dateKey(new Date());
 for(let w=0;w<state.weekCount;w++)root.appendChild(renderWeek(keyAdd(state.weekStart,w*7),todayKey));
 updateMonthTitle();renderAlarmPreview();$('connectionPill').textContent=isSignedIn()?'Google接続中':state.mode==='cloud'?'保存済みの予定を表示中':'体験モード';$('navScopeLabel').textContent=state.filter==='shared'?'共有':'個人';
}
/** One week row: 7 day cells behind, events laid out in up to 5 lanes; a multi-day event is one bar across its days. */
function renderWeek(ws,todayKey){const row=document.createElement('div');row.className='week';row.dataset.week=ws;const days=Array.from({length:7},(_,i)=>keyAdd(ws,i));
 const monthEnd=days.find(k=>keyAdd(k,1).slice(8)==='01');if(monthEnd)row.dataset.monthEnd=monthEnd.slice(0,7);
 days.forEach((k,i)=>{const mon=+k.slice(5,7);const cell=document.createElement('div');cell.className=`day m${mon%2}${k===todayKey?' today':''}${k===state.selected?' selected':''}`;cell.dataset.day=k;cell.style.gridColumn=String(i+1);const num=document.createElement('span');num.className='daynum';num.textContent=k.slice(8)==='01'?`${mon}/1`:String(+k.slice(8));cell.appendChild(num);cell.addEventListener('click',()=>openDay(k));row.appendChild(cell);});
 const seen=new Set(),evs=[];for(const k of days)for(const e of eventsForDay(k))if(!seen.has(e)){seen.add(e);evs.push(e);}
 const span=e=>{const c=e._days.map(k=>days.indexOf(k)).filter(i=>i>=0);return [Math.min(...c),Math.max(...c)];};
 evs.sort((a,b)=>{const[as,ae]=span(a),[bs,be]=span(b);return (b.scope==='shared')-(a.scope==='shared')||(be-bs)-(ae-as)||as-bs||String(a.start).localeCompare(String(b.start));});/* 共有 first, then longer bars */
 const lanes=Array.from({length:LANES},()=>Array(7).fill(false)),hidden=Array(7).fill(0);
 for(const e of evs){const[s,t]=span(e);let lane=-1;for(let l=0;l<LANES&&lane<0;l++){let free=true;for(let c=s;c<=t;c++)if(lanes[l][c])free=false;if(free)lane=l;}
  if(lane<0){for(let c=s;c<=t;c++)hidden[c]++;continue;}for(let c=s;c<=t;c++)lanes[lane][c]=true;
  const chip=document.createElement('div');chip.className=`miniEvent ${e.kind}${isPartner(e)?' partner':''}${e._days.length>1?' span':''}${e._days[0]<days[0]?' contL':''}${e._days[e._days.length-1]>days[6]?' contR':''}${e.syncState==='saving'?' saving':''}`;
  chip.style.gridColumn=`${s+1} / ${t+2}`;chip.style.gridRow=String(lane+2);chip.textContent=e.title;row.appendChild(chip);}
 hidden.forEach((n,i)=>{if(!n)return;const more=document.createElement('div');more.className='more';more.style.gridColumn=String(i+1);more.style.gridRow=String(LANES+2);more.textContent=`+${n}`;row.appendChild(more);});
 return row;}
// Header month: the latest month whose last week is on screen (falls back to the first visible week).
function visibleMonth(){const rows=$('calendarGrid').children;const top=$('calTop')?.getBoundingClientRect?.().bottom||0,bottom=(window.innerHeight||800)-($('bottomNav')?.getBoundingClientRect?.().height||0);let best='',first='';
 for(const r of rows){if(!r.getBoundingClientRect)break;const b=r.getBoundingClientRect();if(b.bottom<=top||b.top>=bottom)continue;if(!first)first=r.dataset.week;if(r.dataset.monthEnd&&r.dataset.monthEnd>best)best=r.dataset.monthEnd;}
 return best||(first?keyAdd(first,3):dateKey()).slice(0,7);}
function updateMonthTitle(){const ym=visibleMonth();$('yearTitle').textContent=ym.slice(0,4);$('monthTitle').textContent=`${+ym.slice(5,7)}月`;}
function scrollToDay(key){const row=[...$('calendarGrid').children].find(r=>r.dataset.week<=key&&key<=keyAdd(r.dataset.week,6));if(!row||typeof window.scrollTo!=='function')return;window.scrollTo(0,Math.max(0,row.offsetTop-($('calTop')?.offsetHeight||0)-8));updateMonthTitle();}
// Scrolling moves the months; near either end, more weeks are added (keeping the view still when prepending).
function onScroll(){if(onScroll.pending)return;onScroll.pending=true;(window.requestAnimationFrame||setTimeout)(()=>{onScroll.pending=false;updateMonthTitle();const doc=document.documentElement;if(!doc)return;
 if(window.scrollY<300){const before=doc.scrollHeight;state.weekStart=keyAdd(state.weekStart,-26*7);state.weekCount+=26;render();window.scrollBy(0,doc.scrollHeight-before);}
 else if(window.scrollY+window.innerHeight>doc.scrollHeight-500){state.weekCount+=26;render();}});}
function getTomorrowShift(){const tomorrow=keyAdd(dateKey(new Date()),1);return state.visible.filter(e=>e.kind==='shift'&&inLocal(e.start)===tomorrow&&!e.allDay).sort((a,b)=>a.start.localeCompare(b.start))[0];}
function renderAlarmPreview(){const shift=getTomorrowShift(),root=$('alarmPreview');if(!shift){root.innerHTML='<div class="alarmNote">明日の勤務シフトはありません（または未取得です）。<br>アラームはiPhoneショートカットが21時に確認します。</div>';return;}const times=alarmTimes(shift);root.innerHTML=`<div class="alarmNote">${safe(shift.title)} · ${safe(timeLocal(shift.start))}開始</div><div class="alarmTime">${times?times.join(' / '):'勤務先の移動時間が未設定'}</div><div class="alarmNote">${times?'朝の準備40分＋通勤時間で計算':'勤務先名と通勤時間の対応をご確認ください。'}<br>※ここでの表示は確認用。実際のアラーム作成はショートカットが行います。</div>`;}
function showDialog(el){if(!el.open)el.showModal();}
function clearEventForm(){for(const el of $('eventForm').querySelectorAll('input,textarea,select'))el.disabled=false;state.editing=null;$('eventForm').reset();$('resolveSync').classList.add('hidden');$('eventId').value='';$('eventOrigin').value='';$('eventScope').disabled=false;$('syncGoogle').disabled=false;$('eventDialogTitle').textContent='予定を追加';$('eventNotesLabel').textContent='メモ';$('deleteEvent').classList.add('hidden');$('editHint').classList.add('hidden');$('saveEvent').classList.remove('hidden');$('startDate').value=state.selected;$('endDate').value=state.selected;$('startTime').value='10:00';$('endTime').value='11:00';$('eventScope').value='private';$('allDay').checked=true;$('syncGoogle').checked=false;$('locationField').classList.add('hidden');$('notesField').classList.remove('hidden');allDayToggle();scopeToggle();}
// Time fields stay tappable while 終日 is on (only dimmed); touching one turns 終日 off.
function allDayToggle(){const on=$('allDay').checked;$('startTime').classList.toggle('dimmed',on);$('endTime').classList.toggle('dimmed',on);}
function timeTouched(){if(!$('allDay').checked)return;$('allDay').checked=false;allDayToggle();}
function scopeToggle(){$('scopeHelp').textContent=$('eventScope').value==='shared'?'共有予定は、2人が追加・編集・削除できます。':'非共有予定は本人専用Googleシートに保存され、相手には表示されません。';}
function openEvent(e=null){clearEventForm();if(e){state.editing=e;if(e.lastError){$('resolveSync').classList.remove('hidden');$('editHint').classList.remove('hidden');$('editHint').textContent=e.lastError;}$('eventDialogTitle').textContent=e.kind==='shift'?'勤務シフト':e.origin==='google'?'Googleカレンダーの予定':'予定を編集';$('eventId').value=e.id;$('eventOrigin').value=e.origin;$('eventTitle').value=e.title;$('eventScope').value=e.scope==='private'?'private':'shared';$('startDate').value=inLocal(e.start);$('endDate').value=e.allDay?keyAdd(inLocal(e.end),-1):inLocal(e.end);$('startTime').value=timeLocal(e.start)||'10:00';$('endTime').value=timeLocal(e.end)||'11:00';$('allDay').checked=e.allDay;$('eventNotes').value=e.notes||'';$('eventLocation').value=e.location||'';$('syncGoogle').checked=!!e.sync;
 if(e.kind==='shift'||e.origin==='google')$('eventScope').disabled=true;if(e.kind==='shift')$('syncGoogle').disabled=true;
 if(e.kind==='shift'||(e.origin==='google'&&!e._canEdit)){$('editHint').classList.remove('hidden');$('editHint').textContent='この予定は読み取り専用です。変更は元のシステムで行ってください。';for(const el of $('eventForm').querySelectorAll('input,textarea,select'))el.disabled=true;$('saveEvent').classList.add('hidden');}
 else{$('deleteEvent').classList.remove('hidden');if(e.origin==='google'){$('editHint').classList.remove('hidden');$('editHint').textContent='既存のGoogle予定です。変更・削除は保存後、「すべて同期」を押したときに反映されます。';}}
 if(e.origin==='sheet')$('eventScope').disabled=true;
 }else{for(const el of $('eventForm').querySelectorAll('input,textarea,select'))el.disabled=false;}
 allDayToggle();if(e?.kind==='shift'||(e?.origin==='google'&&!e._canEdit)){for(const el of $('eventForm').querySelectorAll('input,textarea,select'))el.disabled=true;}
 // Shift fields stay read-only; only the viewer's own memo is editable and it never reaches the partner or Google.
 if(e?.kind==='shift'&&state.mode==='cloud'){$('notesField').classList.remove('hidden');$('eventNotesLabel').textContent='自分だけのメモ';$('eventNotes').value=e._memo||'';$('eventNotes').disabled=false;$('saveEvent').classList.remove('hidden');if(!e.lastError)$('editHint').textContent='勤務は辻岡シフトから自動で取り込みます（編集不可）。メモは自分だけに表示され、相手やGoogleカレンダーには送られません。';}
 scopeToggle();showDialog($('eventDialog'));}
async function saveShiftMemo(){const e=state.editing;setBusy(true);
 try{const result=await gateway('saveShiftMemo','private',{id:e.id,notes:$('eventNotes').value.trim()},e._memoRev||0);$('eventDialog').close();upsertLocal(result.record,'private');showMsg('勤務メモを保存しました（自分だけに表示）。');}
 catch(err){showMsg(`保存に失敗しました：${err.message}`,true);}finally{setBusy(false);}}
function buildFromForm(){const allDay=$('allDay').checked;const sd=$('startDate').value,ed=$('endDate').value;if(!sd||!ed)throw Error('開始日と終了日を入力してください。');let start,end;
 if(allDay){start=sd;end=keyAdd(ed,1);if(ed<sd)throw Error('終了日を開始日以降にしてください。');}
 else{start=makeLocalISO(sd,$('startTime').value);end=makeLocalISO(ed,$('endTime').value);if(new Date(end)<=new Date(start))throw Error('終了時刻は開始より後にしてください。');}
 return{title:$('eventTitle').value.trim(),scope:$('eventScope').value,start,end,allDay,notes:$('eventNotes').value.trim(),location:$('eventLocation').value.trim(),sync:$('syncGoogle').checked};}
function targetCal(scope){return scope==='shared'?config.sharedCalId:config.privateCalId;}
function canonicalCal(id){return id==='primary'?(state.allCalendars.find(c=>c.primary)?.id||'primary'):id;}
async function saveEvent(e){e.preventDefault();if(state.busy)return;
 // Showing cached data without a token: sign in from this tap, then save the same form.
 if(state.mode==='cloud'&&!isSignedIn()){connect(()=>saveEvent({preventDefault(){}}));return;}if(state.editing?.kind==='shift'){if(state.mode==='cloud')await saveShiftMemo();return;}let draft;try{draft=buildFromForm();if(!draft.title)throw Error('タイトルを入力してください。');}catch(err){showMsg(err.message,true);return;}
 const old=state.editing;if(old?.kind==='shift'||(old?.origin==='google'&&!old._canEdit))return;setBusy(true);
 try{
 if(state.mode==='demo'){
 const items=parseJson(KEYS.demo,[]);if(old){const idx=items.findIndex(x=>x.id===old.id);if(idx<0)throw Error('元の予定が見つかりません。');items[idx]={...items[idx],...draft,updatedAt:new Date().toISOString()};}else items.push({...draft,id:crypto.randomUUID(),owner:'demo',calendarId:'',googleEventId:'',source:'manual',deleted:false,createdAt:new Date().toISOString(),updatedAt:new Date().toISOString()});localStorage.setItem(KEYS.demo,JSON.stringify(items));
 }else if(old?.origin==='google'){
 const result=await stageExternal(old,draft,false);$('eventDialog').close();upsertLocal(result.record,result.record.scope);showMsg('予定を保存しました。');return;
 }else{
 // Show the event at once and close the dialog; the gateway round trip (a few seconds on Apps Script) finishes in the background.
 const r={...old,...draft,id:old?.id||($('eventId').value||crypto.randomUUID())};$('eventId').value=r.id;const snapshot=state.records.slice();
 $('eventDialog').close();upsertLocal({...r,owner:old?.owner||state.email,source:old?.source||'manual',revision:old?.revision||0,syncState:'saving',deleted:false},draft.scope);showMsg('保存しています…');
 try{const result=await gateway('save',draft.scope,r,old?.revision||0);upsertLocal(result.record,draft.scope);showMsg(result.record.lastError?'予定は保存済みです。Google同期は未完了：'+result.record.lastError:'予定を保存しました。',!!result.record.lastError);}
 catch(err){state.records=snapshot;state.visible=mergeEvents();render();showMsg(`保存に失敗しました：${err.message}`,true);}
 return;
 }
 $('eventDialog').close();showMsg('予定を保存しました。');setBusy(false);await refresh();
 }catch(err){showMsg(`保存に失敗しました：${err.message}`,true);}finally{setBusy(false);}
}
async function removeEvent(_,confirmed=false){const e=state.editing;if(!e||state.busy)return;if(!confirmed&&!window.confirm(`「${e.title}」を削除しますか？`))return;
 if(state.mode==='cloud'&&!isSignedIn()){connect(()=>removeEvent(null,true));return;}setBusy(true);try{
 if(state.mode==='demo'){const items=parseJson(KEYS.demo,[]);localStorage.setItem(KEYS.demo,JSON.stringify(items.filter(x=>x.id!==e.id)));}
 else if(e.origin==='google'){if(!e._canEdit)throw Error('削除権限がありません。');const result=await stageExternal(e,e,true);$('eventDialog').close();state.external=state.external.filter(x=>!(x.calendarId===e.calendarId&&x.googleEventId===e.googleEventId));upsertLocal(result.record,result.record.scope);showMsg('予定を削除しました。「すべて同期」でGoogleにも反映します。');return;}
 else{
 // Remove from view at once; restore if the gateway refuses.
 const snapshot=state.records.slice();$('eventDialog').close();upsertLocal({...e,deleted:true},e._store||e.scope);showMsg('削除しています…');
 try{const result=await gateway('delete',e.scope,e,e.revision||0);upsertLocal(result.record,e._store||e.scope);showMsg('予定を削除しました。');}
 catch(err){state.records=snapshot;state.visible=mergeEvents();render();showMsg(`削除に失敗しました：${err.message}`,true);}
 return;}
 $('eventDialog').close();showMsg('予定を削除しました。');setBusy(false);await refresh();
 }catch(err){showMsg(`削除に失敗しました：${err.message}`,true);}finally{setBusy(false);}}
function openSettings(){syncInputs();showDialog($('settingsDialog'));renderSyncReport();}
function demoMode(){clearAuth();state.mode='demo';state.token='';state.email='';state.records=[];state.external=[];state.visible=[];$('settingsDialog').close();loadDemo();showMsg('体験モードに切り替えました。データはこの端末だけに保存されます。');}
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
 // 同期 = import shifts from the shift spreadsheet only; MyLife never pushes to Google Calendar.
 isGoogleConnected:()=>false,
 shift:async()=>{ensureCloud();return gateway('syncShift','shared');},
 google:syncGoogleCalendars,
 onBusy:b=>{for(const id of ['settingsSyncAll']){$(id).disabled=b||state.busy;$(id).textContent=b?'同期中…':'↻ 勤務シフトを同期';}},
 onResult:report=>{const label={success:'成功',failed:'失敗',skipped:'未接続'};const shiftValue=report.results.shift.value,shiftLabel=report.results.shift.status==='skipped'?'対象外':label[report.results.shift.status]+(shiftValue&&report.results.shift.status==='success'?`（追加${shiftValue.created}・変更${shiftValue.updated}・削除${shiftValue.deleted}）`:'');const summary=(report.status==='success'?'同期完了':report.status==='partial'?'一部失敗':'同期失敗')+' · 勤務：'+shiftLabel;localStorage.setItem(KEYS.syncReport,JSON.stringify({at:report.at,status:report.status,summary}));renderSyncReport();$('syncStatus').textContent=summary;const detail=Object.values(report.results).filter(x=>x.status==='failed').map(x=>x.message).join('／');showMsg(summary+(detail?'：'+detail:''),report.status!=='success');if(report.status!=='success')noteError(summary+(detail?'：'+detail:''));refresh();}
});
function bind(){
 $('shareApp').addEventListener('click',shareApp);
 $('resolveSync').addEventListener('click',resolveSync);
 // The token lives in RAM only, so an iPhone reload signs out; pressing sync then reconnects first and syncs right after.
 $('settingsSyncAll').addEventListener('click',()=>isSignedIn()||!(config.clientId&&config.gatewayUrl)?syncRunner.run():connect(()=>syncRunner.run()));$('disconnectGoogle').addEventListener('click',disconnectGoogle);$('enableGoogle').addEventListener('click',enableGoogle);
 $('navSettings').addEventListener('click',openSettings);$('closeSettings').addEventListener('click',()=>$('settingsDialog').close());
 $('loginBtn').addEventListener('click',()=>connect());
 $('todayBtn').addEventListener('click',()=>{state.selected=dateKey(new Date());markSelected();scrollToDay(state.selected);});
 // Bottom menu: one button toggles 個人 ⇄ 共有.
 $('navScope').addEventListener('click',()=>{state.filter=state.filter==='shared'?'private':'shared';render();});
 $('closeDay').addEventListener('click',()=>$('dayDialog').close());$('dayAdd').addEventListener('click',()=>{$('dayDialog').close();openEvent();});$('dayPrev').addEventListener('click',()=>openDay(keyAdd(state.selected,-1)));$('dayNext').addEventListener('click',()=>openDay(keyAdd(state.selected,1)));
 $('navAdd').addEventListener('click',()=>openEvent());
 window.addEventListener?.('scroll',onScroll,{passive:true});window.addEventListener?.('resize',updateMonthTitle);
 $('eventForm').addEventListener('submit',saveEvent);$('cancelEvent').addEventListener('click',()=>$('eventDialog').close());$('closeEvent').addEventListener('click',()=>$('eventDialog').close());$('deleteEvent').addEventListener('click',removeEvent);$('allDay').addEventListener('change',allDayToggle);for(const id of ['startTime','endTime'])for(const ev of ['focus','click','input'])$(id).addEventListener(ev,timeTouched);$('eventScope').addEventListener('change',scopeToggle);

}
bind();syncInputs();showLastError();if(!loadCached())loadDemo();
// A still-valid saved sign-in refreshes the calendar right away, without any tap.
if(restoreAuth())(async()=>{try{await bootstrap();saveAuth();await refresh();}catch(e){noteError(`自動更新に失敗：${e.message}`);}})();
// While signed in, pick up the partner's changes without pressing anything: on returning to the app and every 5 minutes while it is open.
function autoRefresh(){if(isSignedIn()&&state.protocol===3&&document.visibilityState!=='hidden')refresh();}
document.addEventListener?.('visibilitychange',()=>{if(document.visibilityState==='visible')autoRefresh();});setInterval(autoRefresh,5*60*1000);renderSyncReport();setTimeout(()=>scrollToDay(dateKey()),0);if(!config.clientId||!config.gatewayUrl)$('onboardingStatus').textContent='Google接続の初期設定は準備中です。体験モードは利用できます。設定完了後は、共有リンクとGoogleログインで開始できます。';if('serviceWorker'in navigator&&location.protocol==='https:')navigator.serviceWorker.register('./service-worker.js').catch(()=>{});
})();
