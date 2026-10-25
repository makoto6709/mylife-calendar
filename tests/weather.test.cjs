const {test}=require('node:test');const assert=require('node:assert/strict');
const {parseJmaForecast,weatherIcon,weatherLabel}=require('../core.js');
// Same shape as https://www.jma.go.jp/bosai/forecast/data/forecast/080000.json (茨城県), trimmed.
const feed=()=>[
 {timeSeries:[
  {timeDefines:['2026-10-25T17:00:00+09:00','2026-10-26T00:00:00+09:00','2026-10-27T00:00:00+09:00'],areas:[{area:{code:'080010'},weatherCodes:['313','302','202']},{area:{code:'080020'},weatherCodes:['313','300','202']}]},
  {timeDefines:['2026-10-25T18:00:00+09:00','2026-10-26T00:00:00+09:00','2026-10-26T06:00:00+09:00','2026-10-26T12:00:00+09:00','2026-10-26T18:00:00+09:00'],areas:[{area:{code:'080010'},pops:['30','30','60','70','60']},{area:{code:'080020'},pops:['40','30','60','90','70']}]},
  {timeDefines:['2026-10-26T00:00:00+09:00','2026-10-26T09:00:00+09:00'],areas:[{area:{code:'40201'},temps:['16','17']},{area:{code:'40341'},temps:['10','30']}]}]},
 {timeSeries:[
  {timeDefines:['2026-10-26T00:00:00+09:00','2026-10-27T00:00:00+09:00','2026-10-28T00:00:00+09:00','2026-11-01T00:00:00+09:00'],areas:[{area:{code:'080000'},weatherCodes:['302','202','101','101'],pops:['','50','20','20']}]},
  {timeDefines:['2026-10-26T00:00:00+09:00','2026-10-27T00:00:00+09:00','2026-10-28T00:00:00+09:00','2026-11-01T00:00:00+09:00'],areas:[{area:{code:'40201'},tempsMin:['','15','9','11'],tempsMax:['','22','18','20']}]}]}];
test('JMA forecast: 3-day area codes override the weekly ones, rain chance is the day max, temps from 水戸',()=>{const d=parseJmaForecast(feed());
 assert.deepEqual(d['2026-10-25'],{code:'313',pop:30});
 assert.deepEqual(d['2026-10-26'],{code:'302',pop:70,min:16,max:17},'tomorrow: near codes/pops, 水戸 morning min and daytime max');
 assert.deepEqual(d['2026-10-27'],{code:'202',pop:50,min:15,max:22});
 assert.deepEqual(d['2026-11-01'],{code:'101',pop:20,min:11,max:20},'weekly only');});
test('JMA forecast: blanks and broken input give no data instead of throwing',()=>{assert.deepEqual(parseJmaForecast(null),{});assert.deepEqual(parseJmaForecast({ok:true}),{});const d=parseJmaForecast(feed());assert.equal('min' in d['2026-10-25'],false);});
test('weather code to icon and label',()=>{assert.equal(weatherIcon('101'),'☀️');assert.equal(weatherLabel('202'),'くもり');assert.equal(weatherIcon('313'),'☔');assert.equal(weatherLabel('400'),'雪');assert.equal(weatherIcon(''),'');});
