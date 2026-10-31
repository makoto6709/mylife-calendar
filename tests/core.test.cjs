const {test}=require('node:test');const assert=require('node:assert/strict');const c=require('../core.js');
test('Tokyo date and time independent of host timezone',()=>{assert.equal(c.dateKey('2026-10-19T16:00:00Z'),'2026-10-20');assert.equal(c.timeLocal('2026-10-19T00:00:00Z'),'09:00');assert.equal(c.makeLocalISO('2026-10-19','09:00'),'2026-10-19T00:00:00.000Z');});
test('month/year transition and invalid date',()=>{assert.equal(c.keyAdd('2026-12-31',1),'2027-01-01');assert.equal(c.validDate('2026-02-30'),false);assert.throws(()=>c.makeLocalISO('2026-02-30','10:00'));});
test('wake alarms for all three shops',()=>{for(const [title,expected] of [['水戸内原',['07:55','08:00','08:05']],['東海',['07:25','07:30','07:35']],['下妻',['06:40','06:45','06:50']]])assert.deepEqual(c.alarmTimes({title,start:'2026-10-20T09:00:00+09:00'}),expected);assert.equal(c.alarmTimes({title:'エリアMT',start:'2026-10-20T09:00:00+09:00'}),null);});

test('bulk all-day holidays are separate days across gaps and year boundaries',()=>{
 const c=require('../core.js');const draft={title:' 休日 ',scope:'shared',notes:'共通メモ',sync:false};
 const rows=c.allDayCopies(draft,['2027-01-02','2026-12-31','2026-12-31']);
 assert.equal(rows.length,2);assert.equal(rows[0].start,'2026-12-31');assert.equal(rows[0].end,'2027-01-01');assert.equal(rows[1].start,'2027-01-02');assert.equal(rows[1].end,'2027-01-03');assert.ok(rows.every(r=>r.allDay&&r.title==='休日'&&r.scope==='shared'));
 assert.throws(()=>c.allDayCopies(draft,[]),/日を選んで/);assert.throws(()=>c.allDayCopies(draft,['2026-02-30']),/日付/);
});

test('one-hour duration follows JST dates through midnight and year end',()=>{
 const c=require('../core.js');assert.deepEqual(c.oneHourLater('2026-12-31','23:30'),{date:'2027-01-01',time:'00:30'});assert.deepEqual(c.oneHourLater('2026-11-01','09:05'),{date:'2026-11-01',time:'10:05'});assert.throws(()=>c.oneHourLater('2026-11-01','24:00'),/日付または時刻/);
});
