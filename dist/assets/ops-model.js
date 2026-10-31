(function(root){
'use strict';
const M=root.CRModel;
const tz='Asia/Jakarta';
function today(){
 const bits=Object.fromEntries(new Intl.DateTimeFormat('en-US',{timeZone:tz,year:'numeric',month:'2-digit',day:'2-digit'}).formatToParts(new Date()).map(v=>[v.type,v.value]));
 return [bits.year,bits.month,bits.day].join('-');
}
function dueKey(v){
 if(v==null||v==='')return null;
 if(typeof v==='number'&&v>=20000&&v<=90000)return new Date(Date.UTC(1899,11,30)+Math.floor(v)*86400000).toISOString().slice(0,10);
 const s=String(v).trim();let m=s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
 if(m){const y=+m[1],mo=+m[2],d=+m[3];return validate(y,mo,d);}
 m=s.match(/^(\d{1,2})[\/.-](\d{1,2})[\/.-](20\d{2})$/);
 if(m)return validate(+m[3],+m[2],+m[1]);
 return null;
}
function validate(y,m,d){const test=new Date(Date.UTC(y,m-1,d));return test.getUTCFullYear()===y&&test.getUTCMonth()+1===m&&test.getUTCDate()===d?[y,String(m).padStart(2,'0'),String(d).padStart(2,'0')].join('-'):null;}
function needsAction(r){const g=M.statusGroup(r.status);return g!=='closed'&&g!=='cancel';}
function followupBucket(r,now=today()){
 if(!needsAction(r))return 'complete';
 const d=dueKey(r.dueDate);
 if(!d)return 'unscheduled';
 if(d<now)return 'overdue';
 if(d===now)return 'today';
 return 'upcoming';
}
const ranking={overdue:0,today:1,upcoming:2,unscheduled:3,complete:4};
function compareFollowups(a,b){
 const aa=followupBucket(a),bb=followupBucket(b);
 const ag=ranking[aa]-ranking[bb];if(ag)return ag;
 const da=dueKey(a.dueDate)||'9999',db=dueKey(b.dueDate)||'9999';if(da!==db)return da.localeCompare(db);
 const status=v=>M.statusGroup(v.status)==='active'?0:1;
 return status(a)-status(b)||Number(!String(a.action||'').trim())-Number(!String(b.action||'').trim())||String(a.title).localeCompare(String(b.title),'id');
}
function dueLabel(r,now=today()){
 const d=dueKey(r.dueDate);
 if(!d)return 'Due date belum ada';
 if(d<now)return 'Overdue · '+d;
 if(d===now)return 'Due today';
 return 'Due '+d;
}
function reconBucket(r){
 if(r.manualRekon==='Belum Rekon')return 'unrecorded';
 if(r.manualRekon==='Sudah Rekon')return 'local';
 if(M.hasRekon(r))return 'recorded';
 return M.statusGroup(r.status)==='closed'?'missing':'unrecorded';
}
function isVerified(r){return Boolean(r.verification&&r.verification.evidenceReference&&r.verification.verifiedBy&&r.verification.verifiedAt);}
root.CROps={today,dueKey,needsAction,followupBucket,compareFollowups,dueLabel,reconBucket,isVerified};
})(window);
