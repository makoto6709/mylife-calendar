/* Pure domain logic. All calendar dates are Asia/Tokyo. */
(function(root){
'use strict';
const pad=n=>String(n).padStart(2,'0');
function dateKey(value=new Date()){return new Intl.DateTimeFormat('sv-SE',{timeZone:'Asia/Tokyo',year:'numeric',month:'2-digit',day:'2-digit'}).format(new Date(value));}
function timeLocal(value){if(!value||/^\d{4}-\d{2}-\d{2}$/.test(value))return '';return new Intl.DateTimeFormat('en-GB',{timeZone:'Asia/Tokyo',hour:'2-digit',minute:'2-digit',hourCycle:'h23'}).format(new Date(value));}
function keyAdd(key,n){const d=new Date(key+'T12:00:00+09:00');d.setUTCDate(d.getUTCDate()+n);return dateKey(d);}
function validDate(s){return /^\d{4}-\d{2}-\d{2}$/.test(s)&&Number.isFinite(+new Date(s+'T00:00:00+09:00'))&&dateKey(s+'T00:00:00+09:00')===s;}
function makeLocalISO(date,time){if(!validDate(date)||!/^([01]\d|2[0-3]):[0-5]\d$/.test(time))throw Error('日付または時刻が正しくありません。');return new Date(date+'T'+time+':00+09:00').toISOString();}
function workTravel(title){return title.includes('水戸内原')?15:title.includes('東海')?45:title.includes('下妻')?90:null;}
function alarmTimes(e){const travel=workTravel(e.title);if(travel===null||e.allDay)return null;const baseline=+new Date(e.start)-(travel+40)*60000;return [-10,-5,0].map(n=>timeLocal(baseline+n*60000));}
function parseEventText(text,year=+dateKey().slice(0,4)){
 const normalized=text.normalize('NFKC');let match=normalized.match(/(20\d\d)[\s年./-]+(\d{1,2})[\s月./-]+(\d{1,2})\s*日?/),date='';
 if(match)date=`${match[1]}-${pad(+match[2])}-${pad(+match[3])}`;else{match=normalized.match(/(\d{1,2})\s*[月/.-]\s*(\d{1,2})\s*日?/);if(match)date=`${year}-${pad(+match[1])}-${pad(+match[2])}`;}
 if(!validDate(date))date='';
 const times=[...normalized.matchAll(/(?<!\d)(午前|午後)?\s*([01]?\d|2[0-3])\s*(?:[:時]\s*([0-5]?\d)(?:分)?|時)(?!\d)/g)].map(m=>{let h=+m[2];if(m[1]==='午後'&&h<12)h+=12;if(m[1]==='午前'&&h===12)h=0;return `${pad(h)}:${pad(+(m[3]||0))}`;}).slice(0,2);
 const title=normalized.split(/\r?\n/).map(s=>s.trim()).find(s=>s.length>=3&&!/^[\d\s年月日曜日月火水木金土()/.：:\-〜~時分]+$/.test(s))||'画像から取り込んだ予定';
 return{date,start:times[0]||'',end:times[1]||'',title:title.slice(0,100),notes:text.slice(0,3000)};
}
const api={dateKey,timeLocal,keyAdd,validDate,makeLocalISO,workTravel,alarmTimes,parseEventText};root.MyLifeCore=api;if(typeof module!=='undefined')module.exports=api;
})(typeof window==='undefined'?globalThis:window);
