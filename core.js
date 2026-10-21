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
const api={dateKey,timeLocal,keyAdd,validDate,makeLocalISO,workTravel,alarmTimes};root.MyLifeCore=api;if(typeof module!=='undefined')module.exports=api;
})(typeof window==='undefined'?globalThis:window);
