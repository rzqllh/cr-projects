import {readFileSync} from 'node:fs';
import {deflateRawSync} from 'node:zlib';
import {strict as assert} from 'node:assert';
import vm from 'node:vm';
const enc=new TextEncoder();
const u16=(d,o,v)=>d.writeUInt16LE(v,o), u32=(d,o,v)=>d.writeUInt32LE(v,o);
function syntheticZip(files){
 const locals=[],central=[];let offset=0;
 for(const [name,xml] of Object.entries(files)){
  const raw=Buffer.from(xml,'utf8'),z=deflateRawSync(raw),nm=Buffer.from(name,'utf8');
  const local=Buffer.alloc(30);u32(local,0,0x04034b50);u16(local,8,8);u32(local,18,z.length);u32(local,22,raw.length);u16(local,26,nm.length);
  const item=Buffer.concat([local,nm,z]);locals.push(item);
  const cen=Buffer.alloc(46);u32(cen,0,0x02014b50);u16(cen,10,8);u32(cen,20,z.length);u32(cen,24,raw.length);u16(cen,28,nm.length);u32(cen,42,offset);central.push(Buffer.concat([cen,nm]));offset+=item.length;
 }
 const c=Buffer.concat(central),end=Buffer.alloc(22);u32(end,0,0x06054b50);u16(end,8,central.length);u16(end,10,central.length);u32(end,12,c.length);u32(end,16,offset);return Buffer.concat([...locals,c,end]);
}
// A deliberately narrow XML DOM test shim. Production uses the browser's DOMParser.
class SimpleNode {
 constructor(name,attrs={},parent=null){this.localName=name;this.attrs=attrs;this.parent=parent;this.children=[];this.directText='';}
 getAttribute(name){return this.attrs[name]??null;}
 getAttributeNS(_ns,name){return this.attrs['r:'+name]??this.attrs[name]??null;}
 getElementsByTagName(name){const out=[];const visit=x=>{for(const c of x.children){if(name==='*'||c.localName===name)out.push(c);visit(c)}};visit(this);return out;}
 get textContent(){return this.directText+this.children.map(x=>x.textContent).join('');}
}
class SimpleDOMParser {
 parseFromString(text){const root=new SimpleNode('root'),stack=[root];const pattern=/<[^>]*>|[^<]+/g;for(const token of text.match(pattern)||[]){if(/^<\?/.test(token)||/^<!/.test(token))continue;if(/^<\//.test(token)){if(stack.length>1)stack.pop();continue;}if(token.startsWith('<')){const m=/^<([\w:-]+)/.exec(token);if(!m)continue;const attrs={};for(const a of token.matchAll(/([\w:-]+)="([^"]*)"/g))attrs[a[1]]=a[2];const el=new SimpleNode(m[1].split(':').pop(),attrs,stack.at(-1));stack.at(-1).children.push(el);if(!/\/>$/.test(token))stack.push(el);}else stack.at(-1).directText+=token.replace(/&amp;/g,'&');}return root;}
}
const window={};const ctx={window,console,TextDecoder,Blob,Response,DecompressionStream,DOMParser:SimpleDOMParser,Uint8Array,DataView,Map,Number,Array,Error};
vm.createContext(ctx);
vm.runInContext(readFileSync(new URL('../dist/assets/data-model.js',import.meta.url),'utf8'),ctx);
vm.runInContext(readFileSync(new URL('../dist/assets/workbook-file.js',import.meta.url),'utf8'),ctx);
const zip=syntheticZip({
 'xl/workbook.xml':'<?xml version="1.0"?><workbook xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="CR Ultimate" r:id="rId1" /></sheets></workbook>',
 'xl/_rels/workbook.xml.rels':'<Relationships><Relationship Id="rId1" Target="worksheets/sheet1.xml" /></Relationships>',
 'xl/sharedStrings.xml':'<sst><si><t>SYNTH-01</t></si><si><t>Testing local workbook import</t></si><si><t>INPROGRESS</t></si></sst>',
 'xl/worksheets/sheet1.xml':'<worksheet><sheetData><row r="3"><c r="A3"><v>1</v></c><c r="B3" t="s"><v>0</v></c><c r="C3" t="s"><v>1</v></c><c r="D3" t="s"><v>2</v></c><c r="E3"><v>0.5</v></c></row></sheetData></worksheet>'
});
const file={name:'Synthetic.xlsx',size:zip.length,arrayBuffer:async()=>zip.buffer.slice(zip.byteOffset,zip.byteOffset+zip.byteLength)};
const result=await window.CRWorkbookFile.parseWorkbook(file,window.CRModel);
assert.equal(result.records.length,1);
assert.equal(result.counts['CR Ultimate'],1);
assert.equal(result.records[0].crId,'SYNTH-01');
assert.equal(result.records[0].rowHint,3);
assert.equal(window.CRModel.progressPct(result.records[0].progress),50);
await assert.rejects(()=>window.CRWorkbookFile.parseWorkbook({...file,name:'fake.xls'},window.CRModel),/xlsx/);
await assert.rejects(()=>window.CRWorkbookFile.parseWorkbook({...file,size:17*1024*1024},window.CRModel),/16 MB/);
console.log('PASS local XLSX import — synthetic ZIP, XML mapping, row identity, safety limits');
