/**
 * MyLife Shift Bridge v1.0
 * 独立した Google Apps Script プロジェクトに貼る。
 * 元の「辻岡シフト」Google Sheets を一切書き換えない。
 * 既存実シートの月度別・5行カレンダー形式 (A:G) に対応。
 * 備考行は一切読み込まず、共有カレンダーにも転記しない。
 */
const SHIFT_CONFIG = Object.freeze({
  // 読み取り元「辻岡シフト」: 既存アプリの元データ
  SHEET_ID: 'ScriptPropertiesのSHIFT_SOURCE_IDで設定',
  // PWA設定で作った「MyLife 勤務」カレンダーのIDに置き換えてください
  WORK_CALENDAR_ID: 'ScriptPropertiesのSHIFT_WORK_CALENDAR_IDで設定',
  TAB_NAME_PATTERN: /^\d{4}年\d{1,2}月度$/,
  TIMEZONE: 'Asia/Tokyo',
  MARKER_PREFIX: 'MYLIFE_SHIFT_V1:',
  // 対象店舗とその他勤務区分。H・休み希望・「勤務先」等は対象にしない。
  WORK_TYPES: ['水戸内原','水戸内原店','東海','東海店','下妻','下妻店','エリアMT','その他']
});

/** 手動実行 / 時間主導トリガー共通。 */
function syncMyLifeShifts(){
  const props=PropertiesService.getScriptProperties();const sourceId=props.getProperty('SHIFT_SOURCE_ID'),calendarId=props.getProperty('SHIFT_WORK_CALENDAR_ID');if(!sourceId||!calendarId)throw Error('SHIFT_SOURCE_IDとSHIFT_WORK_CALENDAR_IDをスクリプトプロパティに設定してください。');
  const lock=LockService.getScriptLock();
  if(!lock.tryLock(30000))return;
  try{
    const source=SpreadsheetApp.openById(sourceId);
    const calendar=CalendarApp.getCalendarById(calendarId);
    if(!calendar)throw new Error('勤務カレンダーがありません。IDと権限を確認してください。');
    const days=new Map();
    source.getSheets().filter(s=>SHIFT_CONFIG.TAB_NAME_PATTERN.test(s.getName())).forEach(sheet=>{
      if(sheet.getLastRow()<5)return;
      const found=sheet.getName().match(/^(\d{4})年(\d{1,2})月度$/);
      const tabYear=Number(found[1]),tabMonth=Number(found[2]);
      // 1行目は曜日。各週は2行目から5行単位。
      // 日付、勤務先、開始、終了の4行だけを取得し、備考行をそもそも読み取らない。
      for(let sheetRow=2;sheetRow+3<=Math.min(sheet.getLastRow(),100);sheetRow+=5){
        const block=sheet.getRange(sheetRow,1,4,7).getDisplayValues();
        const dates=block[0]||[],shops=block[1]||[],starts=block[2]||[],ends=block[3]||[];
        const isDateHeader=dates.filter(cell=>/^\s*\d{1,2}\/\d{1,2}/.test(cell)).length;
        if(isDateHeader===0)continue;
        for(let col=0;col<7;col++){
          const dm=String(dates[col]||'').match(/^\s*(\d{1,2})\/(\d{1,2})/);
          if(!dm)continue;
          const m=+dm[1],d=+dm[2];
          // 月度が2027年1月なら2026年12月21日〜2027年1月20日
          const y=(tabMonth===1&&m===12)?tabYear-1:tabYear;
          if(!((m===tabMonth&&d<=20)||(m===(tabMonth===1?12:tabMonth-1)&&d>=21)))continue;
          const check=new Date(`${y}-${String(m).padStart(2,'0')}-${String(d).padStart(2,'0')}T12:00:00+09:00`);
          if(Utilities.formatDate(check,SHIFT_CONFIG.TIMEZONE,'yyyy-MM-dd')!==`${y}-${String(m).padStart(2,'0')}-${String(d).padStart(2,'0')}`)continue;
          const key=Utilities.formatDate(check,SHIFT_CONFIG.TIMEZONE,'yyyy-MM-dd');
          const shop=String(shops[col]||'').trim();
          const startText=String(starts[col]||'').trim();
          const endText=String(ends[col]||'').trim();
          const start=asDateTime_(y,m,d,startText),end=asDateTime_(y,m,d,endText);
          // 日付行があるのに勤務先が「勤務先」や「H」の場合は、その日が休み。
          const isWork=SHIFT_CONFIG.WORK_TYPES.some(type=>type===shop);
          if(!isWork){if(!['','H','休','公休','休日','有休','有給','休み希望','勤務先'].includes(shop))throw Error('未登録の勤務先です：'+key+' '+shop);days.set(key,null);continue;}
          if(!start||!end||end<=start)throw Error('勤務時刻が不正です：'+key+'（既存予定を削除せず停止しました）');
          days.set(key,{key,title:`【勤務】${shop.replace(/店$/,'')}店`.replace('エリアMT店','エリアMT').replace('その他店','その他'),start,end});
        }
      }
    });
    if(!days.size)throw new Error('日付を読み取れませんでした。シート構造を確認してください。');
    const sorted=[...days.keys()].sort();const begin=new Date(`${sorted[0]}T00:00:00+09:00`);const finish=new Date(`${sorted[sorted.length-1]}T00:00:00+09:00`);finish.setTime(finish.getTime()+2*86400000);
    // 元データにある日付分だけ管理し、他のGoogleカレンダー予定は絶対に変更しない。
    const existing=new Map();
    calendar.getEvents(begin,finish).forEach(event=>{
      const match=String(event.getDescription()||'').match(/^MYLIFE_SHIFT_V1:(\d{4}-\d{2}-\d{2})$/);
      if(match&&days.has(match[1])){
        if(!existing.has(match[1]))existing.set(match[1],[]);
        existing.get(match[1]).push(event);
      }
    });
    let created=0,updated=0,deleted=0,untouched=0;
    days.forEach((shift,key)=>{
      const found=existing.get(key)||[];
      if(!shift){found.forEach(e=>{e.deleteEvent();deleted++;});return;}
      const primary=found.shift();
      if(primary){
        const changed=primary.getTitle()!==shift.title||primary.getStartTime().getTime()!==shift.start.getTime()||primary.getEndTime().getTime()!==shift.end.getTime();
        if(changed){primary.setTitle(shift.title);primary.setTime(shift.start,shift.end);updated++;}else untouched++;
      }else{calendar.createEvent(shift.title,shift.start,shift.end,{description:SHIFT_CONFIG.MARKER_PREFIX+key});created++;}
      found.forEach(e=>{e.deleteEvent();deleted++;});
    });
    Logger.log(JSON.stringify({dates:days.size,created,updated,deleted,untouched,calendar:calendarId}));
    return{dates:days.size,created,updated,deleted,untouched};
  }finally{lock.releaseLock();}
}

/** 最初に1度手動実行して、15分ごとの時間トリガーを設定。 */
function installMyLifeTrigger(){
  ScriptApp.getProjectTriggers().filter(t=>t.getHandlerFunction()==='syncMyLifeShifts').forEach(t=>ScriptApp.deleteTrigger(t));
  syncMyLifeShifts();
  ScriptApp.newTrigger('syncMyLifeShifts').timeBased().everyMinutes(15).create();
}

function asDateTime_(year,month,day,input){
  const match=String(input||'').match(/^(\d{1,2}):([0-5]\d)$/);
  if(!match)return null;
  const h=+match[1],m=+match[2];if(h>23)return null;
  // スクリプトのプロジェクトタイムゾーンを「Asia/Tokyo」に必ず設定すること。
  return new Date(`${year}-${String(month).padStart(2,'0')}-${String(day).padStart(2,'0')}T${String(h).padStart(2,'0')}:${String(m).padStart(2,'0')}:00+09:00`);
}
