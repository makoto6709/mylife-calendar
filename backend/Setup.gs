/** MyLife Calendar 初期設定 v2.3（本人が一度だけ実行）
 * 本体gatewayとは別の新しいApps Scriptプロジェクトに貼り付けて実行する。
 * 共有DB（Sheets）と共有カレンダーを用意し、相手1人だけに共有して、
 * gatewayのScript Property「MYLIFE_CONFIG」に貼るJSONをログへ出力する。
 * 既存IDを指定した場合は検査だけ行い、データ・見出しを書き換えない。
 * 元の勤務シフト表は名前を読むだけで、一切書き込まない。
 * この中の値（メール・ID）はGitHubへコミットしないこと。
 */
const SETUP={
 OWNER_EMAIL:'',     // 本人のGoogleメール（このスクリプトを実行しているアカウント）
 PARTNER_EMAIL:'',   // 恵里子さんのGoogleメール（本人が確認したもの）
 CLIENT_ID:'',       // Google CloudのOAuthクライアントID（〜.apps.googleusercontent.com）
 SHIFT_SOURCE_ID:'', // 「辻岡シフト」のスプレッドシートID
 SHARED_SHEET_ID:'', // 既存の共有DBがあれば。空なら新規作成
 SHARED_CAL_ID:''    // 既存の共有カレンダーがあれば。空なら新規作成
};
const SETUP_HEADERS=['id','owner','scope','title','start','end','allDay','notes','sync','calendarId','googleEventId','source','deleted','createdAt','updatedAt','location','revision','syncState','googleEtag','lastError'];

function setupMyLifeShared(){
 const owner=SETUP.OWNER_EMAIL.trim().toLowerCase(),partner=SETUP.PARTNER_EMAIL.trim().toLowerCase();
 if(!owner||!partner||owner===partner)throw Error('OWNER_EMAILとPARTNER_EMAILに別々の正しいメールを入力してください。');
 if(Session.getEffectiveUser().getEmail().toLowerCase()!==owner)throw Error('OWNER_EMAILのアカウントで実行してください（現在：'+Session.getEffectiveUser().getEmail()+'）。');
 if(!/\.apps\.googleusercontent\.com$/.test(SETUP.CLIENT_ID.trim()))throw Error('CLIENT_IDを入力してください。');
 if(!SETUP.SHIFT_SOURCE_ID.trim())throw Error('SHIFT_SOURCE_IDを入力してください。');
 const shiftName=SpreadsheetApp.openById(SETUP.SHIFT_SOURCE_ID.trim()).getName(); // 読み取りのみ
 const sheetId=SETUP.SHARED_SHEET_ID.trim()||createSharedSheet_(partner);
 checkSharedSheet_(sheetId,owner,partner);
 const calId=SETUP.SHARED_CAL_ID.trim()||createSharedCalendar_(partner);
 checkSharedCalendar_(calId,owner,partner);
 const config={clientId:SETUP.CLIENT_ID.trim(),sharedSheetId:sheetId,sharedCalId:calId,workCalId:'',users:{[owner]:{privateSheetId:'',privateCalId:''},[partner]:{privateSheetId:'',privateCalId:''}},shiftSourceId:SETUP.SHIFT_SOURCE_ID.trim()};
 Logger.log('勤務元：'+shiftName+'（読み取りのみ確認）');
 Logger.log('共有DB：https://docs.google.com/spreadsheets/d/'+sheetId+'/edit');
 Logger.log('次の1行をgatewayのスクリプトプロパティ MYLIFE_CONFIG に貼り付けてください：\n'+JSON.stringify(config));
}
function createSharedSheet_(partner){
 const ss=SpreadsheetApp.create('MyLife_共有予定_DB');const sh=ss.getSheets()[0];
 ss.setSpreadsheetTimeZone('Asia/Tokyo');sh.setName('events');sh.getRange(1,1,1,SETUP_HEADERS.length).setValues([SETUP_HEADERS]);sh.setFrozenRows(1);
 ss.addEditor(partner);return ss.getId();
}
function checkSharedSheet_(id,owner,partner){
 const ss=SpreadsheetApp.openById(id),sh=ss.getSheetByName('events');
 if(!sh)throw Error('共有DBに「events」シートがありません。既存ファイルは変更していません。');
 const head=sh.getRange(1,1,1,SETUP_HEADERS.length).getDisplayValues()[0];
 if(JSON.stringify(head)!==JSON.stringify(SETUP_HEADERS))throw Error('共有DBの見出しがv2形式（events!A1:T1）ではありません。自動書換えはしません：'+head.join(','));
 const file=DriveApp.getFileById(id);
 if(file.getSharingAccess()!==DriveApp.Access.PRIVATE)throw Error('共有DBのリンク共有をオフ（制限付き）にしてください。');
 const editors=[file.getOwner(),...file.getEditors()].map(u=>u.getEmail().toLowerCase()),viewers=[...file.getViewers()].map(u=>u.getEmail().toLowerCase()).filter(x=>!editors.includes(x));
 if(new Set(editors).size!==2||!editors.includes(owner)||!editors.includes(partner)||viewers.length)throw Error('共有DBの編集者を本人と相手の2人だけにしてください（現在：'+editors.concat(viewers).join(', ')+'）。');
}
function createSharedCalendar_(partner){
 const cal=CalendarApp.createCalendar('MyLife 共有',{timeZone:'Asia/Tokyo',summary:'MyLife Calendarの共有予定（2人のみ）'});
 // 相手もACLを確認できるよう「変更および共有の管理権限」(owner)で共有する。
 calendarApi_('/calendars/'+encodeURIComponent(cal.getId())+'/acl?sendNotifications=true','post',{role:'owner',scope:{type:'user',value:partner}});
 return cal.getId();
}
function checkSharedCalendar_(id,owner,partner){
 const acl=(calendarApi_('/calendars/'+encodeURIComponent(id)+'/acl','get').items||[]).filter(x=>x.role!=='none');
 const users=acl.map(x=>x.scope&&x.scope.type==='user'?String(x.scope.value).toLowerCase():'*'+(x.scope&&x.scope.type));
 if(acl.length!==2||!users.includes(owner)||!users.includes(partner)||acl.some(x=>x.role!=='owner'))throw Error('共有カレンダーは本人と相手の2人だけに「変更および共有の管理権限」で共有してください（現在：'+acl.map(x=>(x.scope&&x.scope.value||x.scope&&x.scope.type)+':'+x.role).join(', ')+'）。');
}
function calendarApi_(path,method,body){
 const options={method,headers:{Authorization:'Bearer '+ScriptApp.getOAuthToken()},muteHttpExceptions:true};if(body){options.contentType='application/json';options.payload=JSON.stringify(body);}
 const res=UrlFetchApp.fetch('https://www.googleapis.com/calendar/v3'+path,options);if(res.getResponseCode()>=300)throw Error('Calendar API '+res.getResponseCode()+'：'+res.getContentText().slice(0,300));return JSON.parse(res.getContentText()||'{}');
}
