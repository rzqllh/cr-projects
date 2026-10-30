(function(root){
  'use strict';
  // Minimal XLSX reader. Only browser-standard ZIP decompression and XML APIs.
  // No network, third-party scripts, macro execution, or file upload.
  const MAX_FILE=16*1024*1024, MAX_ARCHIVE=96*1024*1024, MAX_ENTRY=32*1024*1024;
  const decoder=new TextDecoder('utf-8');
  const u16=(d,i)=>d.getUint16(i,true),u32=(d,i)=>d.getUint32(i,true);
  const safeName=n=>!n.startsWith('/')&&!n.includes('\\')&&!n.split('/').includes('..');
  function zipIndex(bytes){
    const d=new DataView(bytes.buffer,bytes.byteOffset,bytes.byteLength),len=bytes.length;
    let end=-1;for(let i=len-22;i>=Math.max(0,len-65557);i--){if(u32(d,i)===0x06054b50){end=i;break;}}
    if(end<0)throw Error('Bukan workbook XLSX yang valid (struktur ZIP tidak ditemukan).');
    const count=u16(d,end+10),start=u32(d,end+16);let pos=start,total=0;
    if(count>2000||start>=len)throw Error('Arsip XLSX tidak valid atau terlalu besar.');
    const files=new Map();
    for(let i=0;i<count;i++){
      if(pos+46>len||u32(d,pos)!==0x02014b50)throw Error('Daftar file XLSX rusak.');
      const method=u16(d,pos+10),compressed=u32(d,pos+20),size=u32(d,pos+24),nlen=u16(d,pos+28),extra=u16(d,pos+30),comment=u16(d,pos+32),off=u32(d,pos+42);
      const name=decoder.decode(bytes.subarray(pos+46,pos+46+nlen));pos+=46+nlen+extra+comment;total+=size;
      if(!safeName(name)||size>MAX_ENTRY||total>MAX_ARCHIVE||off+30>len)throw Error('Arsip XLSX melebihi batas aman.');
      if(files.has(name))throw Error('Arsip XLSX mengandung nama file ganda.');
      if(method!==0&&method!==8)continue;
      files.set(name,{method,compressed,size,off});
    }
    async function read(path){
      const x=files.get(path);if(!x)return null;
      if(u32(d,x.off)!==0x04034b50)throw Error('Header ZIP tidak cocok.');
      const dataOffset=x.off+30+u16(d,x.off+26)+u16(d,x.off+28);
      if(dataOffset+x.compressed>len)throw Error('Data ZIP terpotong.');
      const slice=bytes.slice(dataOffset,dataOffset+x.compressed);
      let out;
      if(x.method===0)out=slice;
      else {
        if(typeof DecompressionStream!=='function')throw Error('Browser ini belum mendukung pembacaan XLSX lokal. Gunakan Chrome/Edge/Safari terbaru.');
        try{out=new Uint8Array(await new Response(new Blob([slice]).stream().pipeThrough(new DecompressionStream('deflate-raw'))).arrayBuffer());}
        catch{throw Error('Gagal membaca kompresi XLSX. Coba simpan ulang file melalui Excel.');}
      }
      if(out.byteLength!==x.size)throw Error('Ukuran file XLSX tidak konsisten.');
      return decoder.decode(out);
    }
    return {read};
  }
  const children=(e,name)=>Array.from(e?.children||[]).filter(x=>x.localName===name);
  const descendants=(e,name)=>Array.from(e?.getElementsByTagName('*')||[]).filter(x=>x.localName===name);
  function parseXML(xml){
    if(xml==null)throw Error('Bagian wajib XLSX tidak ditemukan.');
    const doc=new DOMParser().parseFromString(xml,'application/xml');
    if(descendants(doc,'parsererror').length)throw Error('XML pada XLSX tidak valid.');
    return doc;
  }
  function sharedStrings(xml){
    if(!xml)return [];
    const doc=parseXML(xml);
    return descendants(doc,'si').map(si=>descendants(si,'t').map(t=>t.textContent||'').join(''));
  }
  function colIndex(a1){
    const m=/^([A-Z]{1,3})\d+$/.exec(a1||'');if(!m)return -1;
    let n=0;for(const c of m[1])n=n*26+c.charCodeAt(0)-64;
    return n-1;
  }
  function sheetRows(xml,strings){
    const doc=parseXML(xml),result=[];
    for(const row of descendants(doc,'row')){
      const rn=Number(row.getAttribute('r'));if(!Number.isSafeInteger(rn)||rn<1||rn>20000)continue;
      const cells=[];
      for(const c of children(row,'c')){
        const k=colIndex(c.getAttribute('r'));if(k<0||k>=26)continue;
        const typ=c.getAttribute('t');let val=null;
        if(typ==='inlineStr')val=descendants(c,'t').map(t=>t.textContent||'').join('');
        else{
          const text=children(c,'v')[0]?.textContent;
          if(text===undefined||text===null)continue;
          if(typ==='s'){const id=Number(text);if(!Number.isSafeInteger(id)||id<0||id>=strings.length)throw Error('Indeks shared string XLSX tidak valid.');val=strings[id];}
          else if(typ==='b')val=text==='1';
          else if(typ==='str'||typ==='e')val=text;
          else val=text===''?null:Number.isFinite(Number(text))?Number(text):text;
        }
        cells[k]=val;
      }
      result[rn-1]=cells;
    }
    return result;
  }
  async function parseWorkbook(file,model){
    if(!file||!(/\.xlsx$/i.test(file.name||'')))throw Error('Pilih file Excel .xlsx (bukan .xls atau .xlsm).');
    if(file.size>MAX_FILE)throw Error('Ukuran file lebih dari 16 MB; tidak dapat dibaca di browser ini.');
    const bytes=new Uint8Array(await file.arrayBuffer()),zip=zipIndex(bytes);
    const book=parseXML(await zip.read('xl/workbook.xml'));
    const relationships=parseXML(await zip.read('xl/_rels/workbook.xml.rels'));
    const rels=new Map(descendants(relationships,'Relationship').map(e=>[e.getAttribute('Id'),e.getAttribute('Target')]));
    const strings=sharedStrings(await zip.read('xl/sharedStrings.xml'));
    const records=[],counts={};
    for(const sheet of descendants(book,'sheet')){
      const name=sheet.getAttribute('name');if(!model.SHEETS[name])continue;
      const rid=sheet.getAttribute('r:id')||sheet.getAttributeNS('http://schemas.openxmlformats.org/officeDocument/2006/relationships','id');
      const target=rels.get(rid);if(!target)throw Error('Referensi worksheet tidak ditemukan: '+name);
      const path=target.startsWith('/')?target.slice(1):target.startsWith('xl/')?target:'xl/'+target;
      if(!safeName(path))throw Error('Referensi worksheet tidak valid.');
      const rows=sheetRows(await zip.read(path),strings);
      const parsed=model.parseRows(name,rows);
      records.push(...parsed);counts[name]=parsed.length;
    }
    if(!Object.keys(counts).length)throw Error('Sheet CR Ultimate / CR Interim / Cost CR / UPCOMING CR tidak ditemukan.');
    if(records.length===0)throw Error('Workbook tidak memiliki baris CR yang dikenali.');
    if(records.length>20000)throw Error('Workbook berisi terlalu banyak entri untuk mode manual.');
    return {records,counts};
  }
  root.CRWorkbookFile={parseWorkbook,zipIndex,parseXML,sheetRows};
})(window);
