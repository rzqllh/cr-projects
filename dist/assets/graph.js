/* Microsoft Graph delegated OAuth 2.0 PKCE; tokens are kept in memory only. */
(function(root){
 'use strict';
 const store='crv3_pkce_';
 const siteDefault=root.CR_CONFIG||{};
 const URI='https://graph.microsoft.com/v1.0';
 const scopes='openid profile User.Read Files.ReadWrite Sites.Read.All';
 const b64=x=>btoa(String.fromCharCode(...new Uint8Array(x))).replace(/\+/g,'-').replace(/\//g,'_').replace(/=+$/,'');
 const random=()=>b64(crypto.getRandomValues(new Uint8Array(32)));
 const endpoint=(tenant,path)=>`https://login.microsoftonline.com/${encodeURIComponent(tenant||'organizations')}/oauth2/v2.0/${path}`;
 function conf(){const stored=JSON.parse(localStorage.getItem('crv3_public_auth_config')||'{}');return {...siteDefault,...stored};}
 function uri(){return location.origin+location.pathname;}
 const cleanQuery=()=>{if(location.search.includes('code=')||location.search.includes('error=')) history.replaceState({},'',uri()+location.hash);};
 class GraphClient{
  constructor(){this.token=null;this.expiresAt=0;this.workbook=null;this.isActive=false;this.lastSync=null;}
  config(){return conf();}
  configured(){return Boolean(conf().clientId && /^[0-9a-f-]{32,36}$/i.test(conf().clientId));}
  saveConfig(clientId,tenantId){localStorage.setItem('crv3_public_auth_config',JSON.stringify({clientId:String(clientId||'').trim(),tenantId:String(tenantId||'organizations').trim()}));}
  async signIn(){
   if(!this.configured())throw new Error('Application (client) ID belum dikonfigurasi oleh organisasi.');
   if(!crypto?.subtle)throw new Error('Login Microsoft membutuhkan HTTPS atau localhost.');
   const verifier=random()+random();const digest=await crypto.subtle.digest('SHA-256',new TextEncoder().encode(verifier));const state=random();
   sessionStorage.setItem(store+'verifier',verifier);sessionStorage.setItem(store+'state',state);
   const c=conf();const q=new URLSearchParams({client_id:c.clientId,response_type:'code',redirect_uri:uri(),response_mode:'query',scope:scopes,code_challenge:b64(digest),code_challenge_method:'S256',state,prompt:'select_account'});
   location.assign(endpoint(c.tenantId,'authorize')+'?'+q);
  }
  async handleRedirect(){
   const p=new URLSearchParams(location.search), err=p.get('error');
   if(err){const msg=p.get('error_description')||err;cleanQuery();throw new Error('Microsoft login gagal: '+msg.slice(0,240));}
   const code=p.get('code');if(!code)return false;
   const actual=p.get('state'),expected=sessionStorage.getItem(store+'state');const verifier=sessionStorage.getItem(store+'verifier');
   sessionStorage.removeItem(store+'state');sessionStorage.removeItem(store+'verifier');cleanQuery();
   if(!actual||!expected||actual!==expected||!verifier)throw new Error('State OAuth tidak valid; login dibatalkan.');
   const c=conf();
   const res=await fetch(endpoint(c.tenantId,'token'),{method:'POST',headers:{'Content-Type':'application/x-www-form-urlencoded'},body:new URLSearchParams({client_id:c.clientId,grant_type:'authorization_code',code,redirect_uri:uri(),code_verifier:verifier,scope:scopes})});
   const body=await res.json().catch(()=>({}));
   if(!res.ok||!body.access_token)throw new Error('Microsoft menolak token: '+String(body.error_description||res.status).slice(0,250));
   this.token=body.access_token;this.expiresAt=Date.now()+Math.max(1,Number(body.expires_in||3600)-60)*1000;
   return true;
  }
  connected(){return Boolean(this.token && Date.now()<this.expiresAt);}
  disconnect(){this.token=null;this.expiresAt=0;this.workbook=null;this.isActive=false;this.lastSync=null;}
  async request(method,path,data){
   if(!this.connected())throw new Error('Sesi Microsoft kedaluwarsa. Hubungkan ulang untuk melanjutkan.');
   const response=await fetch(URI+path,{method,headers:{'Authorization':`Bearer ${this.token}`,'Accept':'application/json',...(data!==undefined?{'Content-Type':'application/json'}:{})},...(data!==undefined?{body:JSON.stringify(data)}:{})});
   if(response.status===204)return {};
   const body=await response.json().catch(()=>({}));
   if(!response.ok)throw new Error(`Graph ${response.status}: ${String(body.error?.message||response.statusText||'Request gagal').slice(0,230)}`);
   return body;
  }
  async resolveWorkbook(){
   if(this.workbook)return this.workbook;
   const c=conf();
   const site=await this.request('GET',`/sites/${encodeURIComponent(c.siteHost)}:${c.sitePath}`);
   const drives=await this.request('GET',`/sites/${encodeURIComponent(site.id)}/drives`);
   const list=[...(drives.value||[])].sort((a,b)=>(a.name===c.preferredLibrary?-1:0)-(b.name===c.preferredLibrary?-1:0));
   const candidates=[c.workbookPath,'Change Request TSEL/Progress CR/Progress CR TSEL.xlsx','Shared Documents/'+c.workbookPath].filter(Boolean);
   const errors=[];
   for(const drive of list){
    for(const path of candidates){
     try{
      const route=path.split('/').map(encodeURIComponent).join('/');
      const item=await this.request('GET',`/drives/${encodeURIComponent(drive.id)}/root:/${route}`);
      if(item?.file&&item?.id){this.workbook={driveId:drive.id,itemId:item.id,name:item.name,eTag:item.eTag};return this.workbook;}
     }catch(e){errors.push(e.message);}
    }
   }
   throw new Error('Workbook tidak ditemukan pada document library TSAWarrior atau tidak diizinkan oleh akun ini. '+(errors.at(-1)||''));
  }
  worksheetUrl(sheet){return `/drives/${encodeURIComponent(this.workbook.driveId)}/items/${encodeURIComponent(this.workbook.itemId)}/workbook/worksheets('${String(sheet).replace(/'/g,"''")}')`;}
  async readSheet(sheet){
   await this.resolveWorkbook();
   const payload=await this.request('GET',this.worksheetUrl(sheet)+'/usedRange(valuesOnly=true)');
   const address=String(payload.address||''),m=address.match(/!\$?([A-Z]+)\$?(\d+)/i);
   let offsetRow=m?Number(m[2])-1:0,offsetCol=m?m[1].toUpperCase().split('').reduce((a,ch)=>a*26+ch.charCodeAt(0)-64,0)-1:0;
   const values=(payload.values||[]).map(x=>Array.isArray(x)?x:[]);
   const rows=Array.from({length:offsetRow+values.length},()=>[]);
   values.forEach((r,i)=>r.forEach((v,j)=>rows[offsetRow+i][offsetCol+j]=v));
   return rows;
  }
  async readAll(){
   if(this.isActive)throw new Error('Sinkronisasi lain masih berjalan.');
   this.isActive=true;
   try{
    const recordSets=[];
    for(const sheet of Object.keys(root.CRModel.SHEETS)){
     const rows=await this.readSheet(sheet);
     const records=root.CRModel.parseRows(sheet,rows);
     if(sheet==='CR Ultimate' && records.length===0)throw new Error('Sheet CR Ultimate terbaca kosong; pembaruan dibatalkan.');
     recordSets.push(...records);
    }
    this.lastSync=new Date();return recordSets;
   }finally{this.isActive=false;}
  }
  async saveFields(record,base,fields){
   if(!this.connected())throw new Error('Hubungkan Microsoft terlebih dahulu untuk menyimpan ke Excel.');
   const indices=root.CRModel.getEditableFields(record.sheet);
   const changes=Object.entries(fields).filter(([name,value])=>indices[name]!==undefined && value!==(name==='month'?base.reconMonth:base[name]));
   if(!changes.length)return {written:0};
   if(changes.some(([key])=>key==='status') && !indices.status && record.sheet==='UPCOMING CR')throw new Error('Sheet Upcoming tidak memiliki kolom status.');
   const currentRows=await this.readSheet(record.sheet);
   const current=root.CRModel.stableMatch(record,root.CRModel.parseRows(record.sheet,currentRows));
   // Reject if somebody changed any edited source cell between the last read and this save.
   for(const [field] of changes){const idx=indices[field];if(String(current.cells[idx]??'')!==String(base.cells[idx]??''))throw new Error(`Konflik data pada ${field}: sumber Excel sudah berubah. Refresh terlebih dahulu; tidak ada data ditimpa.`);}
   let written=0;
   for(const [field,value] of changes){
    const idx=indices[field],col=(idx+1);let c='',n=col;while(n>0){n--;c=String.fromCharCode(65+n%26)+c;n=Math.floor(n/26);}
    let v;
    if(field==='progress')v=(value===''||value==null)?'':Math.max(0,Math.min(100,Number(value)))/100;
    else if(field==='yearRekon')v=value===''||value==null?'':Number(value);
    else if(field==='dueDate' && /^20\d{2}-\d{2}-\d{2}$/.test(value)){v=Math.round((Date.parse(value+'T00:00:00Z')-Date.UTC(1899,11,30))/86400000);}
    else v=root.CRModel.safeTextForExcel(value);
    try {await this.request('PATCH',this.worksheetUrl(record.sheet)+`/range(address='${c}${current.rowHint}')`,{values:[[v]]});written++;}
    catch(e){throw new Error((written?`${written} sel sudah dikirim sebelum gagal. Periksa Excel; jangan mengulangi tanpa refresh. `:'')+e.message);}
   }
   const verifyRows=await this.readSheet(record.sheet);
   const verified=root.CRModel.stableMatch(record,root.CRModel.parseRows(record.sheet,verifyRows));
   for(const [field,value] of changes){
    const idx=indices[field];const v=verified.cells[idx];
    if(field==='progress'){if(root.CRModel.progressPct(v)!==Number(value))throw new Error(`Penulisan Excel belum terverifikasi pada ${field}; ${written} sel terlanjur dikirim. Periksa workbook.`);}
    else if(field==='dueDate' && /^20\d{2}-\d{2}-\d{2}$/.test(value)){const actual=Number(v);const expected=Math.round((Date.parse(value+'T00:00:00Z')-Date.UTC(1899,11,30))/86400000);if(actual!==expected)throw new Error(`Tenggat belum terverifikasi di Excel; ${written} sel dikirim. Periksa workbook.`);}
    else if(String(v??'').replace(/^'/,'')!==String(value??''))throw new Error(`Penulisan Excel belum terverifikasi pada ${field}; ${written} sel terlanjur dikirim. Periksa workbook.`);
   }
   return {written};
  }
 }
 root.CRGraph={GraphClient};
})(window);
