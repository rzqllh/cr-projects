(function(){
 'use strict';
 const $=id=>document.getElementById(id),M=window.CRModel,Graph=window.CRGraph,O=window.CROps,UX=window.CRUX;
 const SNAP=window.CR_WORKBOOK_SNAPSHOT;
 const HAS_SNAPSHOT=Boolean(SNAP?.records?.length);
 const SOURCE=window.CR_CONFIG?.sourceUrl || ($('sourceLink').getAttribute('href') === '#' ? '' : $('sourceLink').href);
 if(SOURCE)$('sourceLink').href=SOURCE;
 if(!SOURCE){$('sourceLink').setAttribute('aria-disabled','true');$('sourceLink').addEventListener('click',e=>e.preventDefault());}
 const KEYS=['status','progress','action','note','month','yearRekon','dueDate'];
 const safeLoad=()=>{try{return JSON.parse(localStorage.getItem('crv3_local_drafts')||'{}')}catch{return {}}};
 const state={page:'overview',sheet:'CR Ultimate',segment:'',year:'',status:'',recon:'',query:'',metric:'',number:1,pageSize:18,mode:'snapshot',records:SNAP.records,importedAt:null,drafts:safeLoad(),opening:null,lastFocus:null,graph:new Graph.GraphClient(),busy:false,syncError:null};
 const esc=s=>String(s??'').replace(/[&<>"']/g,x=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[x]));
 const asText=s=>String(s??'').trim();
 const money=v=>{const n=Number(v);return !Number.isFinite(n)||v==null||v===''?'—':new Intl.NumberFormat('id-ID',{style:'currency',currency:'IDR',maximumFractionDigits:0}).format(n)};
 const humanTime=v=>new Intl.DateTimeFormat('id-ID',{dateStyle:'medium',timeStyle:'short'}).format(v);
 const hasDrafts=()=>Object.keys(state.drafts).length>0;
 let countSource=null,countIndex=new Map();function identityFrequency(r){if(countSource!==state.records){countSource=state.records;countIndex=new Map();for(const x of state.records){const k=M.draftIdentity(x);countIndex.set(k,(countIndex.get(k)||0)+1);}}return countIndex.get(M.draftIdentity(r))||0;}
 const notice=(msg,type='')=>{const n=document.createElement('div');n.className='toast '+type;n.textContent=msg;$('toastStack').append(n);setTimeout(()=>n.remove(),6500)};
 // Anchor drafts to CR identity, not the mutable Excel row index.
 function draftKey(rec){return M.draftIdentity(rec);}
 function draftFor(rec){
  const key=draftKey(rec);
  const unique=identityFrequency(rec)===1;
  const d=unique?state.drafts[key]:null;
  // Old row-number-based drafts may be displayed only while offline; applying them
  // after a live Excel reorder could attach another person's note to the wrong CR.
  const legacy=state.mode==='snapshot'?state.drafts[rec.key]:null;
  const draft=d||legacy;
  if(!draft)return rec;
  const amended=state.mode==='live'
   ? {...rec,manualRekon:draft.manualRekon||null,docLink:draft.docLink||'',yearCRManual:draft.yearCRManual||''}
   : {...rec,...draft,reconMonth:draft.month!==undefined?draft.month:rec.reconMonth};
  return {...amended,yearCR:amended.yearCRManual?Number(amended.yearCRManual):rec.yearCR};
 }
 function data(){return state.records.map(draftFor)}
 function scoped(){return data().filter(r=>r.sheet===state.sheet&&(!state.segment||String(r.segment||'')===state.segment)&&(!state.year||(state.year==='unknown'?!r.yearCR:String(r.yearCR)===state.year)));}
 function filtered(){
  return scoped().filter(r=>{
   const status=M.statusGroup(r.status);
   if(state.metric==='active'&&!O.needsAction(r))return false;
   if(state.metric==='closed'&&status!=='closed')return false;
   if(state.metric==='recon'&&!M.hasRekon(r))return false;
   if(state.page==='actions'&&!(['active','unknown','done'].includes(status)))return false;
   if(state.page==='recon'&&!M.hasRekon(r))return false;
   if(state.status&&M.statusDisplay(r.status)!==state.status)return false;
   if(state.recon==='has'&&!M.hasRekon(r))return false;
   if(state.recon==='none'&&M.hasRekon(r))return false;
   const q=M.normalized(state.query);
   if(q&&!M.normalized([r.title,r.crId,r.action,r.note,r.requestor,r.number].join(' ')).includes(q))return false;
   return true;
  }).sort((a,b)=>{
   const pri=r=>({active:0,unknown:1,done:2,closed:3,cancel:4})[M.statusGroup(r.status)]??5;
   return pri(a)-pri(b) || (Number(b.yearCR)||0)-(Number(a.yearCR)||0) || Number(b.rowHint)-Number(a.rowHint);
  });
 }
 function populate(id,options,selected,emptyLabel){
  const s=$(id);s.replaceChildren();
  if(emptyLabel)s.add(new Option(emptyLabel,''));
  options.forEach(v=>s.add(new Option(String(v),String(v))));
  s.value=selected;
 }
 function updateScopes(){
  const rows=data().filter(r=>r.sheet===state.sheet);
  const segments=[...new Set(rows.map(r=>r.segment).filter(x=>x!=null&&String(x).trim()))].sort();
  $('segment').disabled=!segments.length;
  if(state.segment&&!segments.some(x=>String(x)===state.segment))state.segment='';
  populate('segment',segments,state.segment,segments.length?'Semua segmen':'Tidak ada segmen');
  const years=[...new Set(rows.map(r=>r.yearCR).filter(Boolean))].sort((a,b)=>b-a);
  $('yearCR').replaceChildren(new Option('Semua tahun',''),...years.map(v=>new Option(String(v),String(v))));
  $('yearCR').add(new Option('Belum diketahui','unknown'));
  $('yearCR').value=state.year;
  $('workstream').value=state.sheet;
 }
 function updateStatuses(){
  const values=[...new Set(scoped().map(r=>M.statusDisplay(r.status)))].sort((a,b)=>a.localeCompare(b,'id'));
  $('statusFilter').replaceChildren(new Option('Semua status',''),...values.map(x=>new Option(x,x)));
  if(state.status&&!values.includes(state.status))state.status='';
  $('statusFilter').value=state.status;
  $('reconFilter').value=state.recon;
 }
 function renderMetrics(){
  const rows=scoped();const ct=x=>rows.filter(x).length;
  const metrics=[['total','Total CR',rows.length,'Pada workstream terpilih','layout'],['active','Needs follow-up',ct(O.needsAction),'Belum Closed / Canceled','checklist'],['closed','Closed',ct(r=>M.statusGroup(r.status)==='closed'),'Status pekerjaan','check'],['recon','Recon records',ct(M.hasRekon),'Belum tentu verified','reconcile']];
  $('metrics').innerHTML=metrics.map(([type,label,val,detail,icon])=>`<button type="button" class="metric ${state.metric===type?'selected':''}" data-metric="${type}" data-type="${type}" aria-pressed="${state.metric===type}"><span class="metric-top">${label}<span class="metric-icon" aria-hidden="true">${UX.icon(icon)}</span></span><strong class="metric-value">${new Intl.NumberFormat('id-ID').format(val)}</strong><span class="metric-foot">${detail}<b aria-hidden="true">↗</b></span></button>`).join('');
  $('metrics').querySelectorAll('[data-metric]').forEach(el=>el.addEventListener('click',()=>{const m=el.dataset.metric;setPage('register');state.metric=m==='total'?'':m;renderTable();scrollToRegister();}));
 }
 function renderOverview(){
  const rows=scoped();$('distributionLabel').textContent=rows.length+' entri';
  const groups=[['closed','Closed','#3e8a75'],['active','In progress','#dfa24c'],['done','Done','#739bd5'],['cancel','Canceled','#b4bfcb'],['unknown','Belum diisi','#dc756f']];
  $('stackedBar').innerHTML=groups.map(([cat])=>`<span class="bar-${cat}" style="width:${rows.length?(rows.filter(r=>M.statusGroup(r.status)===cat).length/rows.length*100):0}%"></span>`).join('');
  $('statusLegend').innerHTML=groups.map(([cat,label,color])=>`<div class="legend-line"><span class="legend-dot" style="background:${color}"></span><span class="legend-label">${label}</span><b>${rows.filter(r=>M.statusGroup(r.status)===cat).length}</b></div>`).join('');
  const attention=rows.filter(O.needsAction).sort(O.compareFollowups).slice(0,5);
  $('attentionList').innerHTML=attention.length?attention.map(r=>`<button class="attention-item" data-open="${esc(r.key)}"><span><b>${esc(r.title)}</b><small>${esc(O.dueLabel(r))} · ${esc(M.statusDisplay(r.status))}</small></span><span class="attention-arrow">↗</span></button>`).join(''):`<div class="attention-item"><span><b>Tidak ada item aktif yang terdeteksi</b><small>Periksa register untuk seluruh pekerjaan.</small></span></div>`;
  $('attentionList').querySelectorAll('[data-open]').forEach(el=>el.addEventListener('click',()=>openDetail(el.dataset.open)));
 }
 function progressMarkup(r){const p=M.progressPct(r.progress);return p===null?'—':`<div class="progress-meter"><div class="progress-track"><span style="width:${p}%"></span></div>${p}%</div>`}
 function renderTable(){
  const rows=filtered(),pages=Math.max(1,Math.ceil(rows.length/state.pageSize));state.number=Math.min(state.number,pages);
  const slice=rows.slice((state.number-1)*state.pageSize,state.number*state.pageSize);
  $('emptyState').hidden=rows.length>0;
  $('crTable').hidden=rows.length===0;
  $('tableBody').innerHTML=slice.map(r=>{
   const group=M.statusGroup(r.status),year=r.yearCR??'—',note=r.action||r.note||'Belum ada tindak lanjut';
   return `<tr><td>${esc(r.number??'—')}</td><td><span class="cr-title cell-truncate" title="${esc(r.title)}">${esc(r.title)}</span><span class="cr-sub">${esc(r.crId||r.segment||r.sheet)}</span></td><td>${esc(year)}</td><td><span class="status-pill ${group}">${esc(M.statusDisplay(r.status))}</span></td><td>${progressMarkup(r)}</td><td><span class="recon-copy ${M.hasRekon(r)?'':'missing'}">${esc(M.reconLabel(r))}</span></td><td><span class="cell-truncate" title="${esc(note)}">${esc(note)}</span></td><td><button class="open-row" data-open="${esc(r.key)}" aria-label="Buka detail ${esc(r.title)}">↗</button></td></tr>`;
  }).join('');
  $('tableCount').textContent=`${slice.length?((state.number-1)*state.pageSize+1):0}–${Math.min(state.number*state.pageSize,rows.length)} dari ${rows.length} pekerjaan`;
  $('pageCount').textContent=`${state.number} / ${pages}`;
  $('prevPage').disabled=state.number<=1;$('nextPage').disabled=state.number>=pages;
  $('tableBody').querySelectorAll('[data-open]').forEach(el=>el.addEventListener('click',()=>openDetail(el.dataset.open)));
 }

 function renderFollowups(){
  const rows=scoped().filter(O.needsAction).sort(O.compareFollowups);
  const groups=[['overdue','Overdue','Melewati due date'],['today','Due today','Jatuh tempo hari ini'],['upcoming','Upcoming','Due date mendatang'],['unscheduled','Needs scheduling','Due date tidak tersedia atau tidak valid']];
  $('followupsCount').textContent=rows.length+' CR perlu review';
  $('followupsSummary').innerHTML=groups.map(([key,label,hint])=>'<div class="ops-stat"><span class="ops-label">'+label+'</span><strong class="ops-number">'+rows.filter(r=>O.followupBucket(r)===key).length+'</strong><span class="ops-hint">'+hint+'</span></div>').join('');
  $('followupsGroups').innerHTML=groups.map(([key,label])=>{
   const entries=rows.filter(r=>O.followupBucket(r)===key);
   return '<section class="ops-group"><div class="ops-group-head"><h3>'+label+'</h3><span>'+entries.length+' CR</span></div><div class="ops-list">'+(entries.length?entries.map(r=>'<button type="button" class="ops-item" data-open="'+esc(r.key)+'"><div><h4>'+esc(r.title)+'</h4><p>'+esc(String(r.action||'Next action belum tercatat').slice(0,220))+'</p></div><span class="ops-secondary">'+esc(r.crId||r.sheet)+' · '+esc(M.statusDisplay(r.status))+'<br>'+esc(O.dueLabel(r))+'</span><span class="ops-tag">Buka CR →</span></button>').join(''):'<div class="ops-empty">Tidak ada CR dalam kelompok ini.</div>')+'</div></section>';
  }).join('');
  $('followupsGroups').querySelectorAll('[data-open]').forEach(el=>el.addEventListener('click',()=>openDetail(el.dataset.open)));
 }
 function renderReconciliation(){
  const rows=scoped(),groups=[['missing','Closed · no recon record','Closed tanpa catatan rekon'],['recorded','Recorded · pending verification','Periode tercatat, bukti final belum terverifikasi'],['local','Local confirmation','Konfirmasi lokal belum tersinkron'],['unrecorded','Not recorded','Tidak ada catatan rekon']];
  $('reconCount').textContent=rows.length+' CR · '+state.sheet;
  $('reconSummary').innerHTML=groups.map(([key,label])=>'<div class="ops-stat"><span class="ops-label">'+label+'</span><strong class="ops-number">'+rows.filter(r=>O.reconBucket(r)===key).length+'</strong><span class="ops-hint">'+(key==='recorded'?'Bukan status verified':key==='local'?'Draft browser':'Perlu ditinjau')+'</span></div>').join('');
  $('reconGroups').innerHTML=groups.map(([key,label,desc])=>{
   const entries=rows.filter(r=>O.reconBucket(r)===key);
   return '<section class="ops-group"><div class="ops-group-head"><div><h3>'+label+'</h3><span>'+desc+'</span></div><span>'+entries.length+' CR</span></div><div class="ops-list">'+(entries.length?entries.map(r=>'<button type="button" class="ops-item" data-open="'+esc(r.key)+'"><div><h4>'+esc(r.title)+'</h4><p>'+esc(r.crId||r.sheet)+' · '+esc(M.statusDisplay(r.status))+'</p></div><span class="ops-secondary">'+esc(M.reconLabel(r))+'<br>'+esc(r.yearRekon||'Tahun rekon belum ada')+'</span><span class="ops-tag">Periksa →</span></button>').join(''):'<div class="ops-empty">Tidak ada CR dalam kelompok ini.</div>')+'</div></section>';
  }).join('');
  $('reconGroups').querySelectorAll('[data-open]').forEach(el=>el.addEventListener('click',()=>openDetail(el.dataset.open)));
 }

 function renderNav(){
  const title={overview:'Overview',register:'CR Register',actions:'Follow-ups',recon:'Reconciliation'};
  const headings={overview:['Overview','Prioritas tindak lanjut dan kondisi operasional CR.'],register:['CR Register','Cari, filter, dan perbarui Change Request.'],actions:['Follow-ups','Tindakan berikutnya berdasarkan due date dan status pekerjaan.'],recon:['Reconciliation','Periksa periode, nilai, dan bukti rekonsiliasi.']};
  $('breadcrumb').textContent=title[state.page];$('pageTitle').innerHTML=headings[state.page][0];$('pageSubtitle').textContent=headings[state.page][1];
  $('registerTitle').textContent='CR Register';
  $('registerKicker').textContent='REGISTER / '+state.sheet.toUpperCase();
  $('overviewSection').hidden=state.page!=='overview';$('metrics').hidden=state.page!=='overview';$('registerSection').hidden=state.page!=='register';$('followupsSection').hidden=state.page!=='actions';$('reconciliationSection').hidden=state.page!=='recon';document.body.dataset.page=state.page;
  document.querySelectorAll('.nav-item').forEach(n=>{const active=n.dataset.nav===state.page;n.classList.toggle('active',active);n.setAttribute('aria-current',active?'page':'false')});
 }
 function renderConnection(){
  const sourceLocation=$('settingsDialog').querySelector('.settings-row a');
  if(sourceLocation&&SOURCE){sourceLocation.href=SOURCE;sourceLocation.textContent='Buka Excel SharePoint ↗';}
  const manual=state.mode==='manual',connected=state.graph.connected(),last=state.graph.lastSync,failed=Boolean(state.syncError),stale=state.mode==='live'&&(failed||!connected||!last||Date.now()-last.getTime()>Math.max(2.5*Number(state.graph.config().refreshSeconds||60)*1000,180000)),live=state.mode==='live'&&!stale,pending=state.busy;
  $('syncState').className='sync-state '+(live?'live':stale||failed?'error':manual?'manual':'');$('syncState').querySelector('span').textContent=pending?'Menyinkronkan…':live?'Live · Excel Online':failed?'Sinkronisasi gagal':stale?'Data lama · sinkron terhenti':manual?'Manual · Excel lokal':HAS_SNAPSHOT?'Snapshot · offline':'Belum ada data · koneksi diperlukan';
  $('sidebarMode').textContent=live?'Excel live':stale?'Data terakhir tersimpan':manual?'Excel impor manual':'Snapshot lokal';$('sidebarDetail').textContent=live?'Sinkron berkala saat terbuka':stale?'Perlu autentikasi atau refresh':manual?'Data hanya tersedia di tab ini':HAS_SNAPSHOT?'Belum terhubung ke workbook':'Sambungkan Microsoft untuk melihat data';
  $('dataBanner').className='data-banner '+(live?'live':stale||failed?'error':manual?'manual':'');$('bannerTitle').textContent=live?'Data live dari Excel SharePoint':failed?'Sinkron gagal — data belum terverifikasi terbaru':stale?'Sinkron terhenti — data terakhir mungkin usang':manual?'Impor Excel manual — belum tersinkron':HAS_SNAPSHOT?'Mode snapshot — tidak tersinkron otomatis':'Data belum dimuat — koneksi Microsoft diperlukan';
  $('bannerText').textContent=live?`Terakhir diperbarui ${humanTime(last)}. Perubahan harus dikonfirmasi oleh SharePoint.`:failed?`Koneksi gagal: ${state.syncError}. Data terakhir belum tentu terbaru; cek akses Microsoft dan sinkronkan ulang.`:stale?`${last?'Data terakhir dari '+humanTime(last)+'.':'Tidak ada waktu sinkron terakhir.'} Hubungkan kembali untuk memastikan data terbaru.`:manual?`${state.records.length} CR dimuat dari Excel ${humanTime(state.importedAt||new Date())}. Tidak diunggah ke GitHub; edit hanya draft di tab ini. Impor ulang untuk mengambil pembaruan Excel.`:HAS_SNAPSHOT?'Data referensi dari snapshot lokal. Edit belum disimpan ke Excel sampai koneksi Microsoft aktif.':'Source repository ini tidak menyertakan data CR perusahaan. Hubungkan akun Microsoft yang berizin untuk membaca workbook asli.';
  $('connTitle').textContent=live?'Terhubung ke Excel Online':stale?'Sinkron terhenti / sesi habis':manual?'Mode Excel manual':'Excel Online · belum terhubung';$('connSubtitle').textContent=live?'Graph mengakses workbook TSAWarrior dengan izin akun Anda.':manual?'File Excel diproses di tab ini; tidak ada sinkron otomatis atau upload server.':'Tidak ada koneksi workbook sampai Graph memperoleh akses.';
  $('lastSync').textContent=state.graph.lastSync?humanTime(state.graph.lastSync):'Belum pernah';
  $('draftCount').textContent=Object.keys(state.drafts).length+' draft';
  $('snapshotInfo').textContent=manual?`${state.records.length} entri (manual)`:HAS_SNAPSHOT?`${SNAP.records.length} entri`:'Tidak disertakan';
  $('footerSource').textContent=live?'Excel Online · '+humanTime(last):stale?'Sinkron terhenti'+(last?' · '+humanTime(last):''):manual?'Excel manual · '+state.records.length+' entri':HAS_SNAPSHOT?'Snapshot lokal':'Belum ada data';
  $('syncButton').disabled=pending;$('syncButton').querySelector('span:last-child').textContent=manual||!state.graph.configured()?'Impor Excel':'Periksa update';
  $('disconnectButton').disabled=!connected;
 }
 function render(){updateScopes();updateStatuses();renderNav();if(state.page==='overview'){renderMetrics();renderOverview();}else if(state.page==='register'){renderTable();}else if(state.page==='actions'){renderFollowups();}else if(state.page==='recon'){renderReconciliation();}UX.refreshSelects();renderConnection();}
 function scrollToRegister(){requestAnimationFrame(()=>$('registerSection').scrollIntoView({block:'start',behavior:window.matchMedia('(prefers-reduced-motion: reduce)').matches?'instant':'smooth'}));}
 function setPage(next){if(state.page===next)return;state.page=next;state.metric='';state.number=1;state.status='';state.recon='';state.query='';$('search').value='';UX.closeSelects();render();window.scrollTo({top:0,behavior:'instant'});$('sidebar').classList.remove('open');$('menuToggle').setAttribute('aria-expanded','false')}
 function openDetail(key){
  const base=state.records.find(r=>r.key===key);if(!base)return notice('CR ini tidak lagi tersedia.','error');
  const rec=draftFor(base),f=M.getEditableFields(rec.sheet);state.opening={base,rec};state.lastFocus=document.activeElement;
  $('drawerTitle').textContent=rec.title;$('drawerKey').textContent=`${rec.sheet} / Baris sumber ${rec.rowHint}`;
  $('drawerMeta').replaceChildren(...[rec.crId?`ID ${rec.crId}`:null,rec.yearCR?`Tahun CR ${rec.yearCR}`:'Tahun belum diketahui',`Status ${M.statusDisplay(rec.status)}`,rec.requestor?`PIC/Requestor ${rec.requestor}`:null,rec.cost!=null?money(rec.cost):null].filter(Boolean).map(x=>{const el=document.createElement('span');el.textContent=x;return el;}));
  const s=$('editStatus');s.replaceChildren();
  const opts=['CLOSED','INPROGRESS','DONE','NOT STARTED','CANCELED','OPEN',''];
  if(rec.status&&!opts.includes(rec.status))opts.unshift(String(rec.status));
  opts.forEach(x=>s.add(new Option(x||'Belum diisi',x)));
  s.value=rec.status||'';
  $('editProgress').value=M.progressPct(rec.progress)??'';
  $('editAction').value=rec.action??'';$('editNote').value=rec.note??'';
  $('editMonth').value=rec.reconMonth??'';$('editYearRekon').value=rec.yearRekon??'';
  const due=rec.dueDate; $('editDueDate').value=(typeof due==='number' && due>=40000 && due<90000)?new Date(Date.UTC(1899,11,30)+Math.round(due)*86400000).toISOString().slice(0,10):(due??'');
  $('editYearCRManual').value=rec.yearCRManual??'';$('editManualRekon').value=rec.manualRekon??'';$('editDocLink').value=rec.docLink??'';
  [['editStatus','status'],['editProgress','progress'],['editAction','action'],['editNote','note'],['editMonth','month'],['editYearRekon','yearRekon'],['editDueDate','dueDate']].forEach(([id,field])=>$(id).closest('label').hidden=f[field]===undefined);
  $('drawerNotice').className='drawer-notice '+(state.mode==='live'?'live':'');
  $('drawerNotice').textContent=state.mode==='live'?'Kolom sumber akan diperbarui lewat Graph setelah pengecekan konflik. Status rekon manual, override Tahun CR, dan link dokumen tetap anotasi lokal.':state.mode==='manual'?'Mode manual: draft hanya disimpan di tab ini. Ekspor CSV sebelum menutup tab. Excel SharePoint tidak berubah.':'Mode snapshot: seluruh perubahan tersimpan sebagai draft lokal. Excel SharePoint tidak ikut berubah.';
  $('drawerExcelLink').href=SOURCE||'#';$('drawerExcelLink').hidden=!SOURCE;const doc=$('drawerDocLink');doc.hidden=!/^https:\/\//i.test(rec.docLink||'');if(!doc.hidden)doc.href=rec.docLink;
  $('drawerScrim').hidden=false;$('detailDrawer').hidden=false;document.body.style.overflow='hidden';
  $('closeDrawer').focus();
 }
 function closeDetail(){state.opening=null;$('detailDrawer').hidden=true;$('drawerScrim').hidden=true;document.body.style.overflow='';state.lastFocus?.focus?.();}
 function saveLocal(base,changes){
  const key=draftKey(base);
  if(identityFrequency(base)!==1)throw new Error('Identitas CR tidak unik; anotasi lokal tidak disimpan agar tidak salah pekerjaan.');
  const previous=state.drafts[key]||{};
  const updated={...state.drafts,[key]:{...previous,...changes}};
  if(state.mode==='manual'){state.drafts=updated;return;}
  try{localStorage.setItem('crv3_local_drafts',JSON.stringify(updated));state.drafts=updated;}catch{throw new Error('Penyimpanan browser penuh. Ekspor CSV sebagai cadangan.');}
 }
 async function saveDetail(){
  const opening=state.opening;if(!opening)return;
  const {base,rec}=opening;const f=M.getEditableFields(rec.sheet);
  const updates={status:$('editStatus').value,progress:$('editProgress').value,action:$('editAction').value,note:$('editNote').value,month:$('editMonth').value,yearRekon:$('editYearRekon').value,dueDate:$('editDueDate').value.trim()};
  const meta={manualRekon:$('editManualRekon').value,yearCRManual:$('editYearCRManual').value,docLink:$('editDocLink').value.trim()};
  if(meta.docLink && !/^https:\/\//i.test(meta.docLink))return notice('Tautan dokumen harus diawali https://.','error');
  if(meta.yearCRManual && !/^20\d{2}$/.test(meta.yearCRManual))return notice('Override Tahun CR harus tahun valid.','error');
  if(updates.dueDate && !/^20\d{2}-\d{2}-\d{2}$/.test(updates.dueDate) && updates.dueDate!==String(rec.dueDate??''))return notice('Tenggat baru harus format YYYY-MM-DD.','error');
  if(f.progress!==undefined&&updates.progress!==''&&(!Number.isInteger(Number(updates.progress))||Number(updates.progress)<0||Number(updates.progress)>100))return notice('Progres harus berupa angka 0–100.','error');
  if(f.yearRekon!==undefined&&updates.yearRekon!==''&&(!/^20\d\d$/.test(updates.yearRekon)||Number(updates.yearRekon)<2020||Number(updates.yearRekon)>2100))return notice('Tahun rekon harus 2020–2100 atau kosong.','error');
  const changes=Object.fromEntries(Object.entries(updates).filter(([key])=>f[key]!==undefined));
  const originalDue=rec.dueDate;const dueDisplay=(typeof originalDue==='number'&&originalDue>=40000&&originalDue<90000)?new Date(Date.UTC(1899,11,30)+Math.round(originalDue)*86400000).toISOString().slice(0,10):String(originalDue??'');
  const baseAsText={status:asText(rec.status),progress:String(M.progressPct(rec.progress)??''),action:asText(rec.action),note:asText(rec.note),month:asText(rec.reconMonth),yearRekon:String(rec.yearRekon??''),dueDate:dueDisplay};
  const actually=Object.fromEntries(Object.entries(changes).filter(([k,v])=>String(v).trim()!==baseAsText[k]));
  const metaChanges=Object.fromEntries(Object.entries(meta).filter(([k,v])=>String(v)!==String(rec[k]??'')));
  if(!Object.keys(actually).length&&!Object.keys(metaChanges).length)return closeDetail();
  $('saveEdit').disabled=true;$('saveEdit').textContent='Menyimpan…';
  try{
   if(state.mode==='live'){
    if(Object.keys(actually).length)await state.graph.saveFields(base,base,actually);
    if(Object.keys(metaChanges).length)saveLocal(base,metaChanges);
    closeDetail();notice(Object.keys(actually).length?'Sel sumber terverifikasi di SharePoint. Anotasi tambahan tetap lokal.':'Anotasi lokal tersimpan; Excel tidak diubah.','success');
    if(Object.keys(actually).length)await syncNow(true);else render();
   }else{
    saveLocal(base,{...actually,...metaChanges});closeDetail();notice('Draft lokal tersimpan. Excel SharePoint belum diperbarui.','success');render();
   }
  }catch(e){notice(String(e.message||e).slice(0,260),'error');}
  finally{$('saveEdit').disabled=false;$('saveEdit').textContent='Simpan perubahan';}
 }
 function escapeCSV(rows){
  const cols=['key','sheet','rowHint','crId','title','yearCR','status','progress','reconMonth','yearRekon','action','note','requestor','segment','cost','manualRekon','yearCRManual','docLink','dueDate'];
  const lines=[cols.map(M.csvSafe).join(','),...rows.map(r=>cols.map(k=>M.csvSafe(k==='progress'?M.progressPct(r.progress):r[k])).join(','))];
  return '\uFEFF'+lines.join('\r\n');
 }
 function exportCSV(){
  const blob=new Blob([escapeCSV(filtered())],{type:'text/csv;charset=utf-8;'});
  const a=document.createElement('a');a.href=URL.createObjectURL(blob);a.download='CR_Workdesk_'+state.sheet.replace(/\s/g,'_')+'_'+new Date().toISOString().slice(0,10)+'.csv';a.click();setTimeout(()=>URL.revokeObjectURL(a.href),1500);
  notice('CSV berisi data sesuai filter berhasil diekspor.','success');
 }
 async function importCSV(file){
  if(!file)return;
  try{
   const text=await file.text(),matrix=M.parseCSV(text);
   const headers=(matrix.shift()||[]).map(x=>x.trim());const keyidx=headers.indexOf('key');
   if(keyidx===-1)throw new Error('CSV harus memiliki kolom key hasil ekspor dashboard ini.');
   const ididx=headers.indexOf('crId'),titleidx=headers.indexOf('title');
   if(state.mode==='manual'&&(ididx===-1||titleidx===-1))throw new Error('CSV manual memerlukan kolom crId dan title agar perubahan tidak salah CR.');
   const eligible=['status','progress','action','note','reconMonth','yearRekon','dueDate','manualRekon','yearCRManual','docLink'];let count=0;
   for(const cells of matrix){const key=cells[keyidx];const base=state.records.find(r=>r.key===key);if(!base)continue;
    if(state.mode==='manual'&&(M.normalized(base.title)!==M.normalized(cells[titleidx])||M.normalized(base.crId)!==M.normalized(cells[ididx])))continue;
    const f=M.getEditableFields(base.sheet),updates={};
    for(const field of eligible){const idx=headers.indexOf(field);const target=field==='reconMonth'?'month':field;
     if(idx!==-1&&cells[idx]!==undefined&&(f[target]!==undefined||['manualRekon','yearCRManual','docLink'].includes(field)))updates[target]=cells[idx];}
    if(Object.keys(updates).length){const identity=draftKey(base);if(state.records.filter(r=>draftKey(r)===identity).length!==1)continue;state.drafts[identity]={...(state.drafts[identity]||{}),...updates};count++;}
   }
   if(state.mode!=='manual')localStorage.setItem('crv3_local_drafts',JSON.stringify(state.drafts));render();
   notice(`${count} CR diimpor sebagai draft lokal; tidak menulis ke Excel.`,count?'success':'error');
  }catch(e){notice('Impor gagal: '+e.message,'error')}
  $('csvFile').value='';
 }
 async function importWorkbook(file){
  if(!file)return;
  try{
   if(state.mode==='live')throw new Error('Putuskan koneksi live di Pengaturan sebelum beralih ke impor manual.');
   if(!window.CRWorkbookFile)throw new Error('Modul pembaca Excel belum dimuat. Refresh halaman.');
   const {records,counts}=await window.CRWorkbookFile.parseWorkbook(file,M);
   const sheets=Object.keys(counts);
   state.records=records;state.mode='manual';state.importedAt=new Date();state.syncError=null;
   if(!sheets.includes(state.sheet))state.sheet=sheets[0];
   state.segment='';state.year='';state.status='';state.recon='';state.metric='';state.query='';state.number=1;
   $('search').value='';closeSettings();render();
   notice('Excel dimuat lokal: '+records.length+' CR ('+sheets.map(k=>k+': '+counts[k]).join(', ')+'). Tidak ada upload server.','success');
  }catch(e){notice('Impor Excel gagal: '+String(e.message||e).slice(0,220),'error');}
  finally{$('workbookFile').value='';}
 }
 function openSettings(){state.lastFocus=document.activeElement;
  const c=state.graph.config();$('clientId').value=c.clientId||'';$('tenantId').value=c.tenantId||'organizations';$('settingsError').hidden=true;
  $('settingsScrim').hidden=false;$('settingsDialog').hidden=false;document.body.style.overflow='hidden';renderConnection();$('closeSettings').focus();}
 function closeSettings(){$('settingsDialog').hidden=true;$('settingsScrim').hidden=true;document.body.style.overflow='';state.lastFocus?.focus?.();}
 function errSettings(message){$('settingsError').hidden=false;$('settingsError').textContent=message;}
 async function connect(){
  state.graph.saveConfig($('clientId').value,$('tenantId').value);
  try{await state.graph.signIn()}catch(e){errSettings(e.message)}
 }
 async function syncNow(silent=false){
  if(state.busy)return;
  if(!state.graph.connected()){
   if(silent)return;
   if(state.graph.configured()){openSettings();errSettings('Login Microsoft diperlukan. Klik Hubungkan Microsoft untuk melanjutkan.');}
   else openSettings();
   return;
  }
  state.busy=true;renderConnection();
  try{
   const rows=await state.graph.readAll();
   if(!rows.length)throw new Error('Workbook terbaca kosong; snapshot dipertahankan.');
   state.records=rows;state.mode='live';state.syncError=null;
   state.number=1;
   render();if(!silent)notice(`Sinkron berhasil: ${rows.filter(r=>r.sheet===state.sheet).length} CR pada ${state.sheet}.`,'success');
  }catch(e){
   state.syncError=String(e.message||e).slice(0,240);
   if(!silent)notice('Sinkron gagal: '+e.message,'error');
   $('bannerTitle').textContent=state.mode==='live'?'Sinkron terbaru gagal; data terakhir mungkin usang':'Koneksi Graph belum tersedia; memakai snapshot lokal';
   $('bannerText').textContent=String(e.message||e).slice(0,240);
  }finally{state.busy=false;renderConnection()}
 }
 function closeTopModal(){if(!$('settingsDialog').hidden)closeSettings();else if(!$('detailDrawer').hidden)closeDetail();else $('sidebar').classList.remove('open')}
 function setupEvents(){
  document.querySelectorAll('[data-nav]').forEach(el=>el.addEventListener('click',()=>setPage(el.dataset.nav)));
  $('workstream').addEventListener('change',e=>{state.sheet=e.target.value;state.segment='';state.year='';state.status='';state.recon='';state.metric='';state.number=1;render()});
  $('segment').addEventListener('change',e=>{state.segment=e.target.value;state.number=1;state.metric='';render()});
  $('yearCR').addEventListener('change',e=>{state.year=e.target.value;state.number=1;state.metric='';render()});
  $('statusFilter').addEventListener('change',e=>{state.status=e.target.value;state.number=1;render()});
  $('reconFilter').addEventListener('change',e=>{state.recon=e.target.value;state.number=1;render()});
  $('search').addEventListener('input',e=>{state.query=e.target.value;state.number=1;renderTable()});
  $('resetScope').addEventListener('click',()=>{state.segment='';state.year='';state.metric='';state.status='';state.recon='';state.number=1;state.query='';$('search').value='';render()});
  $('clearFilters').addEventListener('click',()=>{state.query='';state.status='';state.recon='';state.metric='';state.number=1;$('search').value='';render()});
  $('prevPage').addEventListener('click',()=>{state.number--;renderTable()});$('nextPage').addEventListener('click',()=>{state.number++;renderTable()});
  $('downloadCSV').addEventListener('click',exportCSV);
  $('heroImport').addEventListener('click',()=>$('workbookFile').click());
  $('densityToggle').addEventListener('click',()=>{const compact=!document.body.classList.contains('density-compact');document.body.classList.toggle('density-compact',compact);$('densityToggle').setAttribute('aria-pressed',String(compact));$('densityToggle').textContent=compact?'Comfortable view':'Compact view';});
  $('copyRedirect').addEventListener('click',async()=>{try{await navigator.clipboard.writeText($('redirectUri').textContent);notice('Redirect URI disalin.','success');}catch{notice('Salin URI secara manual.','error');}});
  $('workbookTrigger').addEventListener('click',()=>$('workbookFile').click());$('workbookSettingsButton').addEventListener('click',()=>$('workbookFile').click());$('workbookFile').addEventListener('change',e=>importWorkbook(e.target.files[0]));
  $('importTrigger').addEventListener('click',()=>$('csvFile').click());$('csvFile').addEventListener('change',e=>importCSV(e.target.files[0]));
  $('closeDrawer').addEventListener('click',closeDetail);$('drawerScrim').addEventListener('click',closeDetail);$('cancelEdit').addEventListener('click',closeDetail);$('saveEdit').addEventListener('click',saveDetail);
  ['settingsButton','openSettingsSidebar','bannerSettings'].forEach(id=>$(id).addEventListener('click',openSettings));
  $('settingsScrim').addEventListener('click',closeSettings);$('closeSettings').addEventListener('click',closeSettings);
  $('connectButton').addEventListener('click',connect);
  $('disconnectButton').addEventListener('click',()=>{state.graph.disconnect();state.records=SNAP.records;state.mode='snapshot';state.importedAt=null;state.syncError=null;closeSettings();render();notice('Microsoft diputus; kembali ke snapshot lokal.');});
  $('syncButton').addEventListener('click',()=>state.mode==='manual'||!state.graph.configured()?$('workbookFile').click():syncNow());
  $('menuToggle').addEventListener('click',()=>{const opened=$('sidebar').classList.toggle('open');$('menuToggle').setAttribute('aria-expanded',String(opened))});
  document.addEventListener('keydown',e=>{
   if(e.key==='Escape')closeTopModal();
   if(e.key==='/'&&!['INPUT','TEXTAREA','SELECT'].includes(document.activeElement?.tagName)){e.preventDefault();if(state.page!=='register')setPage('register');$('search').focus();}
   if(e.key==='Tab'&&(!$('settingsDialog').hidden||!$('detailDrawer').hidden)){
    const modal=!$('settingsDialog').hidden?$('settingsDialog'):$('detailDrawer');
    const els=[...modal.querySelectorAll('button:not(:disabled),input:not(:disabled),select:not(:disabled),textarea:not(:disabled),a[href]')].filter(x=>x.offsetParent!==null),last=els.at(-1),first=els[0];
    if(e.shiftKey&&document.activeElement===first){e.preventDefault();last?.focus();}else if(!e.shiftKey&&document.activeElement===last){e.preventDefault();first?.focus();}
   }
  });
 }
 async function init(){
  if(!SNAP||!M||!Graph){document.body.textContent='CR Workdesk tidak dapat dimuat: file aplikasi tidak lengkap.';return;}
  UX.init();setupEvents();render();
  try{const wasLogin=await state.graph.handleRedirect();if(wasLogin){notice('Microsoft terautentikasi. Memeriksa workbook…','success');await syncNow();}}
  catch(e){openSettings();errSettings(e.message)}
  setInterval(()=>{if(state.mode==='live'&&state.graph.connected()&&!document.hidden&&!state.busy)syncNow(true);else if(state.mode==='live')renderConnection()},(Number(state.graph.config().refreshSeconds)||60)*1000);
 }
 init();
 // Minimal observable test interface, no privileged access or credential exposure.
 window.CRApp={getState:()=>({mode:state.mode,sheet:state.sheet,page:state.page,records:state.records.length,drafts:Object.keys(state.drafts).length,filtered:filtered().length}),sync:()=>syncNow(true)};
})();
