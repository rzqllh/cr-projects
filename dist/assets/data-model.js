(function (root) {
  'use strict';
  const SHEETS = {
    'CR Ultimate': {header:2,title:2,id:1,no:0,status:3,progress:4,action:7,note:6,date:20,year:null,type:8,segment:1,cost:21,month:23,yearRekon:24,recon:null,due:9,requestor:18},
    'CR Interim': {header:2,title:2,id:1,no:0,status:3,progress:6,action:5,note:4,date:15,year:null,type:25,segment:null,cost:null,month:null,yearRekon:null,recon:null,due:7,requestor:13},
    'Cost CR': {header:1,title:4,id:1,no:0,status:5,progress:null,action:17,note:null,date:null,year:2,type:10,segment:null,cost:16,month:22,yearRekon:null,recon:21,due:null,requestor:8},
    'UPCOMING CR': {header:2,title:2,id:null,no:0,status:null,progress:null,action:12,note:13,date:1,year:null,type:3,segment:null,cost:null,month:null,yearRekon:null,recon:null,due:9,requestor:5}
  };
  const raw = v => v === null || v === undefined || String(v).trim() === '' ? null : v;
  const normalized = s => String(s || '').toLocaleLowerCase('id').replace(/\s+/g,' ').trim();
  function yearOf(v) {
    if (v == null || v === '') return null;
    if (typeof v === 'number') {
      if (v >= 2020 && v <= 2100) return Math.floor(v);
      if (v >= 40000 && v <= 90000) return new Date(Date.UTC(1899,11,30) + Math.round(v)*86400000).getUTCFullYear();
    }
    const m = String(v).match(/20[2-9]\d/); return m ? Number(m[0]) : null;
  }
  function statusGroup(s) {
    const v = normalized(s).replace(/\s+/g,'');
    if (/closed|close|selesai|^complete$/.test(v)) return 'closed';
    if (/cancel|drop|reject/.test(v)) return 'cancel';
    if (/^done$|completed|finish/.test(v)) return 'done';
    if (/progress|ongoing|wip|^open$|process|started|pending/.test(v) && !/notstarted/.test(v)) return 'active';
    if (/notstarted|belummulai|notstart/.test(v)) return 'active';
    return 'unknown';
  }
  function statusDisplay(s) {
    const v = normalized(s).replace(/\s+/g,'');
    if (v.includes('closed')) return 'Closed';
    if (v.includes('cancel')) return 'Canceled';
    if (v === 'done') return 'Done';
    if (v.includes('notstarted')) return 'Not started';
    if (v.includes('progress') || v === 'ongoing') return 'In progress';
    if (v === 'open') return 'Open';
    return s ? String(s) : 'Belum diisi';
  }
  function progressPct(v) {
    if (v === null || v === undefined || v === '') return null;
    const n = Number(v);
    if (!Number.isFinite(n)) return null;
    return Math.min(100, Math.max(0, Math.round(n <= 1 ? n * 100 : n)));
  }
  function hasRekon(r) {
    return Boolean(r.manualRekon === 'Sudah Rekon' || (r.reconMonth != null && String(r.reconMonth).trim() && String(r.reconMonth) !== '-') || (r.yearRekon != null && String(r.yearRekon).trim()) || (r.reconRaw != null && String(r.reconRaw).trim() && String(r.reconRaw) !== '-'));
  }
  function reconLabel(r) {
    if(r.manualRekon==='Sudah Rekon')return 'Sudah Rekon (lokal)';
    if(r.manualRekon==='Belum Rekon')return 'Belum Rekon (lokal)';
    if (!hasRekon(r)) return 'Belum tercatat';
    if (r.reconRaw && /^(sudah\s*rekon|reconciled|done)$/i.test(String(r.reconRaw).trim())) return 'Terkonfirmasi';
    if (r.reconMonth || r.yearRekon) return [r.reconMonth, r.yearRekon].filter(x => x !== null && x !== undefined && String(x).trim()).join(' · ');
    return 'Ada catatan rekon';
  }
  function parseRows(sheet, rows) {
    const c = SHEETS[sheet]; if (!c) return [];
    const result=[];
    for (let i=c.header;i<rows.length;i++) {
      const a=Array.from({length:26},(_,j)=>raw(rows[i]?.[j]));
      const rowNum=i+1, title=a[c.title];
      if (!title || ['total','subtotal','description','note'].includes(normalized(title)) || String(title).startsWith('Note :')) continue;
      if (sheet === 'CR Ultimate' && (rowNum > 10000 || !Number.isFinite(Number(a[0])))) continue;
      if (sheet !== 'CR Ultimate' && !Number.isFinite(Number(a[0]))) continue;
      if (sheet === 'CR Ultimate' && Number(a[0]) === 0) continue;
      if (sheet !== 'CR Ultimate' && Number(a[0]) === 0) continue;
      const get=k=>c[k] == null ? null : a[c[k]];
      let yearCR=yearOf(get('date')) || yearOf(get('year'));
      if (!yearCR && sheet==='CR Ultimate') yearCR=yearOf(a[17]);
      const month=get('month') === '-' ? null : get('month');
      result.push({key:`${sheet}@${rowNum}`,sheet,rowHint:rowNum,number:get('no'),crId:get('id'),title:String(title),status:get('status'),progress:get('progress'),action:get('action'),note:get('note'),yearCR,yearRekon:yearOf(get('yearRekon')),reconMonth:month,reconRaw:get('recon'),requestDate:get('date'),dueDate:get('due'),requestor:get('requestor'),type:get('type'),segment:get('segment'),cost:get('cost'),cells:a});
    }
    return result;
  }
  function getEditableFields(sheet) {
    const c=SHEETS[sheet];if(!c)return {};
    const keys=['status','progress','action','note','month','yearRekon','dueDate'];
    return Object.fromEntries(keys.filter(k=>c[k==='dueDate'?'due':k] !== null).map(k=>[k,c[k==='dueDate'?'due':k]]));
  }
  function stableMatch(record,candidates) {
    // Never use row hint alone: reordering the workbook must not redirect an edit.
    const title=normalized(record.title);let list=candidates.filter(x=>normalized(x.title)===title);
    if (record.crId != null && String(record.crId).trim()) list=list.filter(x=>normalized(x.crId)===normalized(record.crId));
    if (list.length !== 1) throw new Error(list.length ? 'Ada beberapa baris cocok. Edit diblokir agar tidak salah CR.' : 'CR ini telah berpindah/berubah atau dihapus. Refresh dahulu.');
    return list[0];
  }
  function safeTextForExcel(value) {
    let v=String(value == null ? '' : value);
    // Protect plain-text writes from being interpreted as spreadsheet formulas.
    if (/^[\u0009\u000d\u000a\s]*[=+@-]/.test(v)) v="'"+v;
    return v;
  }
  function csvSafe(value) {
    const v=safeTextForExcel(value == null ? '' : value);
    return `"${String(v).replace(/"/g,'""')}"`;
  }
  function parseCSV(text) {
    const s=String(text).replace(/^\uFEFF/,'');
    const lines=[];let row=[],value='',q=false;
    for(let i=0;i<s.length;i++){
      const c=s[i];
      if(q){if(c==='"' && s[i+1]==='"'){value+='"';i++;}else if(c==='"')q=false;else value+=c;}
      else if(c==='"')q=true;
      else if(c===','){row.push(value);value='';}
      else if(c==='\n'){row.push(value.replace(/\r$/,''));lines.push(row);row=[];value='';}
      else value+=c;
    }
    if(q)throw new Error('CSV tidak valid: tanda kutip belum ditutup.');
    if(row.length || value){row.push(value.replace(/\r$/,''));lines.push(row);}
    return lines;
  }
  root.CRModel={SHEETS,normalized,yearOf,statusGroup,statusDisplay,progressPct,hasRekon,reconLabel,parseRows,getEditableFields,stableMatch,safeTextForExcel,csvSafe,parseCSV};
})(window);
