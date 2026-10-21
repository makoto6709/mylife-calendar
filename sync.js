/* Manual coordinator: callers invoke run only from an explicit user action. */
(function(root){
'use strict';
function createRunner({shift,google,isGoogleConnected,clock=()=>new Date().toISOString(),onBusy=()=>{},onResult=()=>{}}){
 let busy=false;
 return {get busy(){return busy;},async run(){
  if(busy)return null;busy=true;onBusy(true);const results={};
  try{
   for(const [name,task] of [['shift',shift],['google',google]]){
    if(name==='google'&&!isGoogleConnected()){results[name]={status:'skipped',message:'未接続'};continue;}
    try{results[name]={status:'success',value:await task()};}catch(e){results[name]={status:'failed',message:e.message||'失敗'};}
   }
   const attempted=Object.values(results).filter(x=>x.status!=='skipped'),ok=attempted.filter(x=>x.status==='success').length;
   const report={at:clock(),status:ok===attempted.length?'success':ok?'partial':'failed',results};onResult(report);return report;
  }finally{busy=false;onBusy(false);}
 }};
}
const api={createRunner};root.MyLifeSync=api;if(typeof module!=='undefined')module.exports=api;
})(typeof window==='undefined'?globalThis:window);
