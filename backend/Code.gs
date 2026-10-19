/** MyLife Calendar gateway v2. No token is persisted or logged.
 * Deploy as owner / anyone. Requests authenticate Google access tokens; all
 * Sheets/Calendar/Drive calls use the caller token, NEVER owner's authority.
 * Script properties: MYLIFE_CONFIG (see config.example.json).
 */
const H=['id','owner','scope','title','start','end','allDay','notes','sync','calendarId','googleEventId','source','deleted','createdAt','updatedAt','location','revision','syncState','googleEtag','lastError'];
function doGet(){return json_({ok:true,version:2});}
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
function permissions_(token,id){const a=call_(token,'https://www.googleapis.com/drive/v3/files/'+encodeURIComponent(id)+'/permissions?fields=permissions(type,role,emailAddress,deleted),nextPageToken&pageSize=100');if(a.nextPageToken)throw Error('権限数が多すぎます。');return a.permissions||[];}
function policy_(token,c,email){
 const own=c.users[email];const ids=Object.values(c.users).map(x=>x.privateSheetId);
 if(!own.privateSheetId||!own.privateCalId||!c.sharedCalId||ids.includes(c.sharedSheetId)||new Set(ids).size!==2||Object.values(c.users).some(x=>x.privateCalId===c.sharedCalId)||new Set(Object.values(c.users).map(x=>x.privateCalId)).size!==2)throw Error('共有・個人の保存先を別々に設定してください。');
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
   if(g&&g.status!=='cancelled'){if(g.extendedProperties?.shared?.mylifeId!==r.id)throw Error('Google IDが重複しました。');r=pull_(r,g);}
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
  const b=JSON.parse(e.postData.contents),c=cfg_(),email=authenticate_(b.token,c),p=policy_(b.token,c,email);
  if(b.action==='bootstrap')return json_({ok:true,config:p,email,version:2});
  if(!['save','delete','reconcile','resolve'].includes(b.action)||!['shared','private'].includes(b.scope))throw Error('操作が不正です。');
  const id=b.scope==='shared'?p.sharedSheetId:p.privateSheetId;
  lock=LockService.getScriptLock();if(!lock.tryLock(25000))throw Error('処理が混み合っています。再試行してください。');
  const list=rows_(b.token,id);
  if(b.action==='reconcile'){
   const all=list.filter(r=>r.sync&&(!r.deleted||r.syncState!=='synced')&&(!b.record?.afterId||r.id>b.record.afterId)).sort((a,b)=>a.id.localeCompare(b.id));const targets=all.slice(0,40);
   targets.forEach(r=>flush_(b.token,id,r));return json_({ok:true,count:targets.length,nextAfter:all.length>40?targets[targets.length-1].id:''});
  }
  const input=b.record;validate_(input);if(input.scope!==b.scope)throw Error('保存先が不正です。');
  const old=list.find(r=>r.id===input.id);
  if(old?.revision!==b.expectedRevision&&!(old===undefined&&b.expectedRevision===0))throw Error('別の端末で更新されています。再同期してから編集してください。');
  if(b.action==='resolve'){
   if(!old||!old.lastError)throw Error('解決する競合はありません。');let g;try{g=call_(b.token,googleUrl_(old));}catch(err){if(![404,410].includes(err.code))throw err;}
   old.deleted=false;old.lastError='';old.revision++;
   if(g&&g.status!=='cancelled')pull_(old,g);else{old.sync=false;old.syncState='local';old.googleEtag='';}
   return json_({ok:true,record:write_(b.token,id,old)});
  }
  if(old?.deleted)throw Error('削除済みの予定です。');
  if(old?.syncState?.startsWith('pending'))throw Error('未完了の同期があります。先に同期を再試行してください。');
  if(b.action==='delete'&&!old)throw Error('予定がありません。');
  const now=new Date().toISOString(),r={...old};
  for(const k of ['id','scope','title','start','end','allDay','notes','location','sync'])r[k]=input[k]??'';
  r.owner=old?.owner||email;r.createdAt=old?.createdAt||now;r.updatedAt=now;r.source=old?.source||'manual';r.revision=(old?.revision||0)+1;r.deleted=b.action==='delete';r.lastError='';
  if(r.deleted){r.sync=!!old.sync;r.syncState=r.sync?'pending-delete':'local';}
  else if(r.sync){
   r.calendarId=b.scope==='shared'?p.sharedCalId:p.privateCalId;
   if(old?.calendarId&&old.calendarId!==r.calendarId)throw Error('同期先変更は対応していません。');
   r.googleEventId=old?.googleEventId||'ml'+r.id.replace(/-/g,'');r.googleEtag=old?.googleEtag||'';
   r.syncState=old?.googleEventId?'pending-update':'pending-create';
  }else{r.syncState='local';/* Detach keeps original Google event intact. */}
  write_(b.token,id,r);if(r.sync)flush_(b.token,id,r);
  return json_({ok:true,record:r});
 }catch(err){return json_({ok:false,error:String(err.message).slice(0,400)});}finally{if(lock)lock.releaseLock();}
}
