/** MyLife Calendar gateway v2. No token is persisted or logged.
 * Deploy as owner / anyone. Requests authenticate Google access tokens; all
 * Sheets/Calendar/Drive calls use the caller token, NEVER owner's authority.
 * Script properties: MYLIFE_CONFIG (see config.example.json).
 */
const H=['id','owner','scope','title','start','end','allDay','notes','sync','calendarId','googleEventId','source','deleted','createdAt','updatedAt','location','revision','syncState','googleEtag','lastError'];
function doGet(){return json_({ok:true,version:3,protocol:3});}
function json_(v){return ContentService.createTextOutput(JSON.stringify(v)).setMimeType(ContentService.MimeType.JSON);}
function call_(token,url,method,body,etag){
 const headers={Authorization:'Bearer '+token};if(etag)headers['If-Match']=etag;
 const options={method:method||'get',headers,muteHttpExceptions:true};if(body!==undefined){options.contentType='application/json';options.payload=JSON.stringify(body);}
 const res=UrlFetchApp.fetch(url,options),code=res.getResponseCode();let data={};try{data=JSON.parse(res.getContentText()||'{}');}catch(_){}
 if(code>=300){const err=new Error('Google HTTP '+code+': '+(data.error?.message||data.error_description||'request failed'));err.code=code;throw err;}return data;
}
function cfg_(){const c=JSON.parse(PropertiesService.getScriptProperties().getProperty('MYLIFE_CONFIG')||'{}');if(!c.clientId||!c.users||Object.keys(c.users).length!==2||!c.sharedSheetId)throw Error('サーバー設定が未完了です。');return c;}
function authenticate_(token,c){
 if(typeof token!=='string'||token.length>4096)throw Error('認証が必要です。');
 const t=call_(token,'https://oauth2.googleapis.com/tokeninfo?access_token='+encodeURIComponent(token));
 if(t.aud!==c.clientId||+t.expires_in<=0)throw Error('認証先または有効期限が不正です。');
 const u=call_(token,'https://www.googleapis.com/oauth2/v3/userinfo');
 if(!u.email_verified||!c.users[u.email])throw Error('許可された2人のアカウントで接続してください。');return u.email;
}
/** Provision only the authenticated member's private stores. No ACL grants. */
function provisionPrivate_(token,email){
 const lock=LockService.getScriptLock();if(!lock.tryLock(25000))throw Error('初回設定が混み合っています。再試行してください。');
 try{
  const c=cfg_(),own=c.users[email];if(!own)throw Error('登録されていないアカウントです。');
  const persist=()=>PropertiesService.getScriptProperties().setProperty('MYLIFE_CONFIG',JSON.stringify(c));
  if(!own.privateSheetId){
   const doc=call_(token,'https://sheets.googleapis.com/v4/spreadsheets','post',{properties:{title:'MyLife_個人予定_DB',locale:'ja_JP',timeZone:'Asia/Tokyo'},sheets:[{properties:{title:'events',gridProperties:{frozenRowCount:1}}}]});
   own.privateSheetId=doc.spreadsheetId;own.provisioningHeader=true;persist();
  }
  if(own.provisioningHeader){
   const before=call_(token,root_(own.privateSheetId)+'/values/events!A1:T1');
   if((before.values||[]).length&&JSON.stringify(before.values[0])!==JSON.stringify(H))throw Error('個人DBのヘッダーが変更されています。');
   if(!(before.values||[]).length)call_(token,root_(own.privateSheetId)+'/values/events!A1:T1?valueInputOption=RAW','put',{values:[H]});
   delete own.provisioningHeader;persist();
  }
  if(!own.privateCalId){const cal=call_(token,'https://www.googleapis.com/calendar/v3/calendars','post',{summary:'MyLife 個人',timeZone:'Asia/Tokyo',description:'本人専用。相手への共有は行いません。'});own.privateCalId=cal.id;persist();}
  return c;
 }finally{lock.releaseLock();}
}
function permissions_(token,id){const a=call_(token,'https://www.googleapis.com/drive/v3/files/'+encodeURIComponent(id)+'/permissions?fields=permissions(type,role,emailAddress,deleted),nextPageToken&pageSize=100');if(a.nextPageToken)throw Error('権限数が多すぎます。');return a.permissions||[];}
function policy_(token,c,email){
 const own=c.users[email];const ids=Object.values(c.users).map(x=>x.privateSheetId).filter(Boolean);
 if(!own.privateSheetId||!own.privateCalId||!c.sharedCalId||ids.includes(c.sharedSheetId)||new Set(ids).size!==ids.length||Object.values(c.users).some(x=>x.privateCalId===c.sharedCalId)||new Set(Object.values(c.users).map(x=>x.privateCalId).filter(Boolean)).size!==Object.values(c.users).filter(x=>x.privateCalId).length)throw Error('共有・個人の保存先を別々に設定してください。');
 const privatePermissions=permissions_(token,own.privateSheetId);
 if(privatePermissions.length!==1||privatePermissions[0].type!=='user'||privatePermissions[0].emailAddress!==email||privatePermissions[0].role!=='owner')throw Error('個人シートは本人のみが所有する非共有ファイルにしてください。');
 const shared=permissions_(token,c.sharedSheetId);
 if(shared.length!==2||shared.some(p=>p.type!=='user'||!c.users[p.emailAddress]||!['owner','writer'].includes(p.role)))throw Error('共有シートの編集権限を2人だけにしてください。');
 const acl=call_(token,'https://www.googleapis.com/calendar/v3/calendars/'+encodeURIComponent(own.privateCalId)+'/acl');
 if(acl.nextPageToken||!acl.items?.length||acl.items.some(p=>p.scope?.type!=='user'||p.scope?.value!==email||p.role!=='owner'))throw Error('個人カレンダーは本人だけに限定してください。');
 const sharedAcl=call_(token,'https://www.googleapis.com/calendar/v3/calendars/'+encodeURIComponent(c.sharedCalId)+'/acl');
 if(sharedAcl.nextPageToken||sharedAcl.items?.length!==2||sharedAcl.items.some(p=>p.scope?.type!=='user'||!c.users[p.scope?.value]||!['owner','writer'].includes(p.role)))throw Error('共有カレンダーは2人のみにしてください。');
 return {sharedSheetId:c.sharedSheetId,privateSheetId:own.privateSheetId,sharedCalId:c.sharedCalId,privateCalId:own.privateCalId,workCalId:c.workCalId||''};
}
function root_(id){return 'https://sheets.googleapis.com/v4/spreadsheets/'+encodeURIComponent(id);}
function rows_(token,id){const d=call_(token,root_(id)+'/values/events!A:T');const rows=d.values||[];if(JSON.stringify(rows[0])!==JSON.stringify(H))throw Error('v2ヘッダーが必要です。既存DBの自動書換えはしません。');return rows.slice(1).map((v,i)=>{const o={_row:i+2};H.forEach((k,j)=>o[k]=v[j]??'');for(const k of ['sync','deleted','allDay'])o[k]=String(o[k])==='true';o.revision=+o.revision||0;return o;});}
function write_(token,id,r){const values=[H.map(k=>String(r[k]??''))];if(r._row)call_(token,root_(id)+'/values/events!A'+r._row+':T'+r._row+'?valueInputOption=RAW','put',{values});else{const out=call_(token,root_(id)+'/values/events!A:T:append?valueInputOption=RAW&insertDataOption=INSERT_ROWS','post',{values});r._row=+out.updates.updatedRange.match(/!A(\d+)/)[1];}return r;}
function googleUrl_(r){return 'https://www.googleapis.com/calendar/v3/calendars/'+encodeURIComponent(r.calendarId)+'/events/'+encodeURIComponent(r.googleEventId);}
function payload_(r){return{summary:r.title,location:r.location||'',description:r.notes||'',start:r.allDay?{date:r.start}:{dateTime:r.start,timeZone:'Asia/Tokyo'},end:r.allDay?{date:r.end}:{dateTime:r.end,timeZone:'Asia/Tokyo'},extendedProperties:{shared:{mylifeId:r.id}}};}
function pull_(r,g){r.title=g.summary||'(タイトルなし)';r.start=g.start.dateTime||g.start.date;r.end=g.end.dateTime||g.end.date;r.allDay=!!g.start.date;r.notes=g.description||'';r.location=g.location||'';r.googleEtag=g.etag;r.updatedAt=g.updated||new Date().toISOString();r.revision++;r.syncState='synced';return r;}
function flush_(token,id,r){
 if(!r.sync||!r.calendarId)return r;
 try{
  let g;try{g=call_(token,googleUrl_(r));}catch(e){if(![404,410].includes(e.code))throw e;}
  if(r.deleted){if(g&&g.status!=='cancelled'){if(!r.googleEtag||g.etag!==r.googleEtag)throw Error('Google側で変更されています。削除競合のため停止しました。');call_(token,googleUrl_(r),'delete',undefined,r.googleEtag);}r.syncState='synced';}
  else if(r.syncState==='pending-create'){
   if(g&&g.status!=='cancelled'){if(g.extendedProperties?.shared?.mylifeId!==r.id)throw Error('Google IDが重複しました。');if(g.summary!==r.title||JSON.stringify(g.start)!==JSON.stringify(payload_(r).start)||JSON.stringify(g.end)!==JSON.stringify(payload_(r).end)||String(g.description||'')!==String(r.notes||'')||String(g.location||'')!==String(r.location||''))throw Error('Google側の同名ID予定と内容が異なります。両方を保持して停止しました。');r.googleEtag=g.etag;r.syncState='synced';}
   else{if(g?.status==='cancelled')throw Error('Google側で削除済みです。自動再作成しません。');g=call_(token,googleUrl_(r).slice(0,googleUrl_(r).lastIndexOf('/')),'post',{...payload_(r),id:r.googleEventId});r.googleEtag=g.etag;r.syncState='synced';}
  }else if(r.syncState==='pending-update'){
   if(!g||g.status==='cancelled')throw Error('Google側で削除されています。MyLifeの内容は保持しました。');
   if(!r.googleEtag||g.etag!==r.googleEtag)throw Error('Google側と同時に変更されています。上書きせず内容を保持しました。');
   g=call_(token,googleUrl_(r),'patch',payload_(r),r.googleEtag);r.googleEtag=g.etag;r.syncState='synced';
  }else if(g&&g.status!=='cancelled'&&g.etag!==r.googleEtag)r=pull_(r,g);
  else if(!g||g.status==='cancelled'){r.sync=false;r.syncState='google-deleted';r.lastError='Google側で削除されました。MyLifeの内容は保持しています。';}
  if(r.syncState==='synced')r.lastError='';
 }catch(e){r.lastError=String(e.message).slice(0,300);/* pending state retained for safe retry */}
 return write_(token,id,r);
}
function validate_(r){
 if(!/^[0-9a-f-]{36}$/.test(r.id)||!['shared','private'].includes(r.scope)||typeof r.title!=='string'||!r.title.trim()||r.title.length>200||String(r.notes||'').length>3000||String(r.location||'').length>300)throw Error('予定の形式が不正です。');
 if(typeof r.allDay!=='boolean'||typeof r.sync!=='boolean')throw Error('予定の設定値が不正です。');
 if(r.allDay){const valid=s=>/^\d{4}-\d{2}-\d{2}$/.test(s)&&new Date(s).toISOString().slice(0,10)===s;if(!valid(r.start)||!valid(r.end)||r.end<=r.start)throw Error('終日の日付が不正です。');}
 else if(!/T.*(?:Z|[+-]\d{2}:\d{2})$/.test(r.start)||!Number.isFinite(+new Date(r.start))||!Number.isFinite(+new Date(r.end))||new Date(r.end)<=new Date(r.start))throw Error('予定日時が不正です。');
}
function doPost(e){
 let lock;
 try{
  if(!e.postData||e.postData.contents.length>20000)throw Error('リクエストが不正です。');
  const b=JSON.parse(e.postData.contents);let c=cfg_();const email=authenticate_(b.token,c);if(b.protocol!==3)throw Error('アプリを最新版に更新してください。');if(b.action==='bootstrap')c=provisionPrivate_(b.token,email);const p=policy_(b.token,c,email);
  if(b.action==='bootstrap')return json_({ok:true,config:p,email,version:3,protocol:3});
  if(!['save','delete','reconcile','queueResolve','stageExternal','syncShift'].includes(b.action)||!['shared','private'].includes(b.scope))throw Error('操作が不正です。');
  const id=b.scope==='shared'?p.sharedSheetId:p.privateSheetId;
  lock=LockService.getScriptLock();if(!lock.tryLock(25000))throw Error('処理が混み合っています。再試行してください。');
  const list=rows_(b.token,id);
  if(b.action==='syncShift'){if(b.scope!=='shared')throw Error('勤務は共有保存先のみです。');return json_({ok:true,...syncShift_(b.token,c,email,list)});}
  if(b.action==='reconcile'){
   const all=list.filter(r=>r.sync&&(!r.deleted||r.syncState!=='synced')&&(!b.record?.afterId||r.id>b.record.afterId)).sort((a,b)=>a.id.localeCompare(b.id));const targets=all.slice(0,40);
   let failed=0;targets.forEach(r=>{if(r.syncState==='pending-resolve'){try{let g;try{g=call_(b.token,googleUrl_(r));}catch(e){if(![404,410].includes(e.code))throw e;}r.deleted=false;r.lastError='';if(g&&g.status!=='cancelled')pull_(r,g);else{r.sync=false;r.syncState='local';}write_(b.token,id,r);}catch(e){r.lastError=e.message;write_(b.token,id,r);}}else flush_(b.token,id,r);if(r.lastError)failed++;});return json_({ok:true,failed,count:targets.length,nextAfter:all.length>40?targets[targets.length-1].id:''});
  }
  const input=b.record;validate_(input);if(input.scope!==b.scope)throw Error('保存先が不正です。');
  const old=list.find(r=>r.id===input.id);
  if(old?.source==='shift')throw Error('勤務予定は読み取り専用です。');
  if(b.action==='stageExternal')return json_({ok:true,record:stageExternal_(b.token,c,email,id,list,input)});
  if(old?.revision!==b.expectedRevision&&!(old===undefined&&b.expectedRevision===0))throw Error('別の端末で更新されています。再同期してから編集してください。');
  if(b.action==='queueResolve'||b.action==='resolve'){
   if(!old||!old.lastError)throw Error('解決する競合はありません。');old.syncState='pending-resolve';old.revision++;return json_({ok:true,record:write_(b.token,id,old)});
  }
  if(old?.deleted)throw Error('削除済みの予定です。');

  if(b.action==='delete'&&!old)throw Error('予定がありません。');
  const now=new Date().toISOString(),r={...old};
  for(const k of ['id','scope','title','start','end','allDay','notes','location','sync'])r[k]=input[k]??'';
  r.owner=old?.owner||email;r.createdAt=old?.createdAt||now;r.updatedAt=now;r.source=old?.source||'manual';r.revision=(old?.revision||0)+1;r.deleted=b.action==='delete';r.lastError='';
  if(r.deleted){r.sync=!!old.sync;r.syncState=r.sync?'pending-delete':'local';}
  else if(r.sync){
   r.calendarId=old?.source==='google-import'?old.calendarId:b.scope==='shared'?p.sharedCalId:p.privateCalId;
   if(old?.calendarId&&old.calendarId!==r.calendarId)throw Error('同期先変更は対応していません。');
   r.googleEventId=old?.googleEventId||'ml'+r.id.replace(/-/g,'');r.googleEtag=old?.googleEtag||'';
   r.syncState=old?.googleEtag?'pending-update':'pending-create';
  }else{r.syncState='local';/* Detach keeps original Google event intact. */}
  write_(b.token,id,r);/* Google calls are deferred to explicit reconcile only. */
  return json_({ok:true,record:r});
 }catch(err){return json_({ok:false,error:String(err.message).slice(0,400)});}finally{if(lock)lock.releaseLock();}
}

function stageExternal_(token,c,email,id,list,input){
 const shared=input.scope==='shared';
 if((shared?(id!==c.sharedSheetId||input.calendarId!==c.sharedCalId):(id!==c.users[email].privateSheetId))||input.calendarId===c.workCalId||!input.calendarId||!input.googleEventId||!input.googleEtag)throw Error('Google予定の保存先が不正です。');
 const acl=call_(token,'https://www.googleapis.com/calendar/v3/calendars/'+encodeURIComponent(input.calendarId)+'/acl');
 if(shared){if(acl.nextPageToken||acl.items?.length!==2||acl.items.some(a=>a.scope?.type!=='user'||!c.users[a.scope?.value]||!['writer','owner'].includes(a.role)))throw Error('共有カレンダーの権限が不正です。');}
 else if(acl.nextPageToken||acl.items?.length!==1||acl.items[0].scope?.type!=='user'||acl.items[0].scope?.value!==email||acl.items[0].role!=='owner')throw Error('このGoogleカレンダーの非共有を確認できません。');
 if(list.some(r=>r.calendarId===input.calendarId&&r.googleEventId===input.googleEventId))throw Error('取り込み済みの予定です。画面を更新して再編集してください。');
 const now=new Date().toISOString(),r={...input,owner:email,scope:input.scope,sync:!!input.sync,source:'google-import',createdAt:now,updatedAt:now,revision:1,syncState:input.sync?(input.deleted?'pending-delete':'pending-update'):'local',lastError:''};
 return write_(token,id,r);
}
function syncShift_(token,c,email,list){
 if(!c.shiftSourceId)throw Error('勤務シフトの接続設定が未完了です。');
 const days=readShiftDays_(token,c.shiftSourceId),existing=new Map();
 for(const r of list.filter(r=>r.source==='shift')){const key=r.start.slice(0,10);if(existing.has(key))throw Error('既存勤務に重複があります。データを保持して停止しました。');existing.set(key,r);}
 let created=0,updated=0,deleted=0,unchanged=0;const now=new Date().toISOString();
 days.forEach((shift,key)=>{
  const old=existing.get(key);
  if(!shift){if(old&&!old.deleted){old.deleted=true;old.revision++;old.updatedAt=now;old.syncState=old.sync?'pending-delete':'local';write_(token,c.sharedSheetId,old);deleted++;}return;}
  if(old&&!old.deleted&&old.title===shift.title&&old.start===shift.start&&old.end===shift.end){unchanged++;return;}
  const r={...old,id:old?.id||Utilities.getUuid(),owner:old?.owner||email,scope:'shared',source:'shift',title:shift.title,start:shift.start,end:shift.end,allDay:false,location:'',notes:'',sync:true,calendarId:c.sharedCalId,googleEventId:old?.googleEventId||'',googleEtag:old?.googleEtag||'',deleted:false,createdAt:old?.createdAt||now,updatedAt:now,revision:(old?.revision||0)+1,lastError:''};
  r.googleEventId=r.googleEventId||'ml'+r.id.replace(/-/g,'');r.syncState=old?.googleEtag?'pending-update':'pending-create';write_(token,c.sharedSheetId,r);old?updated++:created++;
 });
 return {created,updated,deleted,unchanged};
}
function readShiftDays_(token,sourceId){
 const base=root_(sourceId),meta=call_(token,base+'?fields=sheets(properties(title,gridProperties(rowCount)))'),days=new Map();
 const work=['水戸内原','水戸内原店','東海','東海店','下妻','下妻店','エリアMT','その他'],holiday=['','H','休','公休','休日','有休','有給','休み希望','勤務先'];
 for(const sheet of meta.sheets||[]){
  const name=sheet.properties.title,ym=name.match(/^(\d{4})年(\d{1,2})月度$/);if(!ym)continue;const ty=+ym[1],tm=+ym[2];
  for(let row=2;row+3<=Math.min(sheet.properties.gridProperties.rowCount,100);row+=5){
   const range="'"+name.replace(/'/g,"''")+"'!A"+row+':G'+(row+3),v=call_(token,base+'/values/'+encodeURIComponent(range)+'?valueRenderOption=FORMATTED_VALUE').values||[];
   for(let col=0;col<7;col++){
    const dm=String(v[0]?.[col]||'').match(/^\s*(\d{1,2})\/(\d{1,2})/);if(!dm)continue;const m=+dm[1],d=+dm[2],y=tm===1&&m===12?ty-1:ty;
    if(!((m===tm&&d<=20)||(m===(tm===1?12:tm-1)&&d>=21)))continue;
    const key=y+'-'+String(m).padStart(2,'0')+'-'+String(d).padStart(2,'0');if(!Number.isFinite(+new Date(key))||new Date(key).toISOString().slice(0,10)!==key)continue;
    const shop=String(v[1]?.[col]||'').trim(),a=String(v[2]?.[col]||'').trim(),b=String(v[3]?.[col]||'').trim();
    if(shop==='勤務先'||(!shop&&!a&&!b))continue;
    if(!work.includes(shop)){if(!holiday.includes(shop))throw Error('未登録の勤務先：'+key);days.set(key,null);continue;}
    if(['','開始時間'].includes(a)&&['','終了時間'].includes(b))continue;
    if(!/^([01]?\d|2[0-3]):[0-5]\d$/.test(a)||!/^([01]?\d|2[0-3]):[0-5]\d$/.test(b))throw Error('勤務時刻が不正：'+key);
    const iso=t=>key+'T'+t.padStart(5,'0')+':00+09:00',start=iso(a),end=iso(b);if(end<=start)throw Error('勤務時刻が不正：'+key);
    days.set(key,{title:'【勤務】'+(shop.replace(/店$/,'')+'店').replace('エリアMT店','エリアMT').replace('その他店','その他'),start,end});
   }
  }
 }
 if(!days.size)throw Error('勤務日付を読み取れません。既存予定は保持しました。');return days;
}
