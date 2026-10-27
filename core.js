/* Pure domain logic. All calendar dates are Asia/Tokyo. */
(function(root){
'use strict';
const pad=n=>String(n).padStart(2,'0');
function dateKey(value=new Date()){return new Intl.DateTimeFormat('sv-SE',{timeZone:'Asia/Tokyo',year:'numeric',month:'2-digit',day:'2-digit'}).format(new Date(value));}
function timeLocal(value){if(!value||/^\d{4}-\d{2}-\d{2}$/.test(value))return '';return new Intl.DateTimeFormat('en-GB',{timeZone:'Asia/Tokyo',hour:'2-digit',minute:'2-digit',hourCycle:'h23'}).format(new Date(value));}
function keyAdd(key,n){const d=new Date(key+'T12:00:00+09:00');d.setUTCDate(d.getUTCDate()+n);return dateKey(d);}
function validDate(s){return /^\d{4}-\d{2}-\d{2}$/.test(s)&&Number.isFinite(+new Date(s+'T00:00:00+09:00'))&&dateKey(s+'T00:00:00+09:00')===s;}
function makeLocalISO(date,time){if(!validDate(date)||!/^([01]\d|2[0-3]):[0-5]\d$/.test(time))throw Error('日付または時刻が正しくありません。');return new Date(date+'T'+time+':00+09:00').toISOString();}
/** Each chosen day is its own all-day event; gaps must never become a multi-day span. */
function allDayCopies(draft,dates){
 if(!String(draft.title||'').trim())throw Error('タイトルを入力してください。');
 if(!['shared','private'].includes(draft.scope))throw Error('公開範囲を選択してください。');
 const days=[...new Set(dates)].sort();
 if(!days.length)throw Error('登録する日を選んでください。');
 if(days.length>366||days.some(d=>!validDate(d)))throw Error('登録する日付を確認してください。');
 return days.map(day=>({...draft,title:draft.title.trim(),start:day,end:keyAdd(day,1),allDay:true}));
}
function workTravel(title){return title.includes('水戸内原')?15:title.includes('東海')?45:title.includes('下妻')?90:null;}
function alarmTimes(e){const travel=workTravel(e.title);if(travel===null||e.allDay)return null;const baseline=+new Date(e.start)-(travel+40)*60000;return [-10,-5,0].map(n=>timeLocal(baseline+n*60000));}
/** 気象庁 茨城県 forecast (bosai/forecast/data/forecast/080000.json) → {'YYYY-MM-DD':{code,min,max,pop}}.
 * The weekly prefecture forecast fills 7 days; the 3-day forecast for the area (北部 080010) overrides its codes and rain chance;
 * temperatures come from the Mito station (40201). Blank strings in the feed mean "not forecast". */
function parseJmaForecast(json,area='080010',station='40201'){
 const out={},day=t=>String(t).slice(0,10),num=v=>v===''||v===null||v===undefined||!Number.isFinite(+v)?null:+v,put=(k,f)=>{out[k]={...out[k],...f};};
 const [near,week]=Array.isArray(json)?json:[],ws=week?.timeSeries||[],ns=near?.timeSeries||[],find=(ts,code)=>ts?.areas?.find(a=>a.area?.code===code)||null;
 const wc=ws[0]?.areas?.[0];(ws[0]?.timeDefines||[]).forEach((t,i)=>{const code=wc?.weatherCodes?.[i],pop=num(wc?.pops?.[i]);put(day(t),{...(code?{code:String(code)}:{}),...(pop!==null?{pop}:{})});});
 const wt=find(ws[1],station);(ws[1]?.timeDefines||[]).forEach((t,i)=>{const mn=num(wt?.tempsMin?.[i]),mx=num(wt?.tempsMax?.[i]);put(day(t),{...(mn!==null?{min:mn}:{}),...(mx!==null?{max:mx}:{})});});
 const nc=find(ns[0],area);(ns[0]?.timeDefines||[]).forEach((t,i)=>{const code=nc?.weatherCodes?.[i];if(code)put(day(t),{code:String(code)});});
 const np=find(ns[1],area),nearPop={};(ns[1]?.timeDefines||[]).forEach((t,i)=>{const v=num(np?.pops?.[i]);if(v===null)return;const k=day(t);nearPop[k]=Math.max(nearPop[k]??-1,v);});for(const [k,v] of Object.entries(nearPop))put(k,{pop:v});
 const nt=find(ns[2],station);(ns[2]?.timeDefines||[]).forEach((t,i)=>{const v=num(nt?.temps?.[i]);if(v===null)return;put(day(t),+String(t).slice(11,13)<9?{min:v}:{max:v});});
 return out;
}
/** JMA weather codes: the first digit is the main weather (1 晴れ, 2 くもり, 3 雨, 4 雪). */
function weatherIcon(code){return {1:'☀️',2:'☁️',3:'☔',4:'⛄'}[String(code||'')[0]]||'';}
function weatherLabel(code){return {1:'晴れ',2:'くもり',3:'雨',4:'雪'}[String(code||'')[0]]||'';}
const api={dateKey,timeLocal,keyAdd,validDate,makeLocalISO,allDayCopies,workTravel,alarmTimes,parseJmaForecast,weatherIcon,weatherLabel};root.MyLifeCore=api;if(typeof module!=='undefined')module.exports=api;
})(typeof window==='undefined'?globalThis:window);

