import vm from 'node:vm';import {readFileSync} from 'node:fs';import assert from 'node:assert/strict';
const read=x=>readFileSync(new URL('../dist/'+x,import.meta.url),'utf8');
const storage={getItem:()=>null,setItem:()=>{},removeItem:()=>{}};
const w={CR_CONFIG:{clientId:'00000000-0000-0000-0000-000000000001',tenantId:'organizations',siteHost:'telkomind.sharepoint.com',sitePath:'/sites/TSAWarrior',preferredLibrary:'Documents',workbookPath:'Document/Change Request TSEL/Progress CR/Progress CR TSEL.xlsx'}};
const c={window:w,localStorage:storage,sessionStorage:storage,location:{origin:'https://example.com',pathname:'/'},crypto:globalThis.crypto,TextEncoder,URLSearchParams,btoa,Date,Math,Number,JSON,encodeURIComponent};vm.createContext(c);
vm.runInContext(read('assets/data-model.js'),c);vm.runInContext(read('assets/graph.js'),c);
const g=new w.CRGraph.GraphClient();g.token='fake-test-only';g.expiresAt=Date.now()+1e7;
const rows=[[''],['No','CR ID','Description','Status','%','Estimasi BAST','Last Update','Action Needed'],[1,'CR001','Project A','CLOSED',.5,null,'Update lama','Tindak lanjut lama']];
let patches=[];let remoteChange=false;
g.request=async(method,path,data)=>{
 if(method==='GET'&&path.startsWith('/sites/')&&path.endsWith('/drives'))return {value:[{id:'d1',name:'Documents'}]};
 if(method==='GET'&&path.startsWith('/sites/'))return {id:'s1'};
 if(method==='GET'&&path.includes('/root:/'))return {id:'item1',file:{mimeType:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'},name:'Progress CR TSEL.xlsx'};
 if(method==='GET'&&path.includes('/usedRange'))return {address:"'CR Ultimate'!A1:H3",values:rows};
 if(method==='PATCH'){
  const match=path.match(/address='([A-Z])(\d+)'/);assert.ok(match);const col=match[1].charCodeAt(0)-65,row=Number(match[2])-1;
  rows[row][col]=data.values[0][0];patches.push({col,row,value:data.values[0][0]});return {values:data.values};
 }
 throw new Error('Unexpected '+method+' '+path);
};
const fetched=await g.readSheet('CR Ultimate');let parsed=w.CRModel.parseRows('CR Ultimate',fetched);
assert.equal(parsed.length,1);assert.equal(parsed[0].status,'CLOSED');
const result=await g.saveFields(parsed[0],parsed[0],{status:'INPROGRESS',action:'Tindak lanjut baru'});
assert.equal(result.written,2);assert.equal(rows[2][3],'INPROGRESS');assert.equal(rows[2][7],'Tindak lanjut baru');
parsed=w.CRModel.parseRows('CR Ultimate',await g.readSheet('CR Ultimate'));
// Someone else changes the same cell; no overwrite must be allowed.
const old={...parsed[0],cells:[...parsed[0].cells]};rows[2][3]='DONE';
await assert.rejects(()=>g.saveFields(old,old,{status:'CLOSED'}),/Konflik data/);
assert.equal(rows[2][3],'DONE');assert.equal(patches.length,2);
// Clearing a percentage must be represented by an empty cell, not 0%.
const latest=w.CRModel.parseRows('CR Ultimate',await g.readSheet('CR Ultimate'))[0];
const cleared=await g.saveFields(latest,latest,{progress:''});
assert.equal(cleared.written,1);assert.equal(rows[2][4],'');
assert.equal(patches.length,3);
console.log('PASS mocked Graph — workbook resolution, two-cell write/read verification, conflicting edit blocked');
