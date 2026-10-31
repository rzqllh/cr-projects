(function(root){
'use strict';
const paths={
layout:'<rect x="3" y="3" width="7" height="7" rx="1.2"/><rect x="14" y="3" width="7" height="7" rx="1.2"/><rect x="3" y="14" width="7" height="7" rx="1.2"/><rect x="14" y="14" width="7" height="7" rx="1.2"/>',
table:'<rect x="3" y="4" width="18" height="16" rx="2"/><path d="M3 10h18M3 15h18M9 4v16"/>',
checklist:'<path d="m3 7 2 2 3-4M12 7h9M3 17l2 2 3-4M12 17h9"/>',
reconcile:'<path d="M4 7h16m-4-4 4 4-4 4M20 17H4m4-4-4 4 4 4"/>',
refresh:'<path d="M20 11a8 8 0 0 0-14-5L4 8m0-5v5h5M4 13a8 8 0 0 0 14 5l2-2m0 5v-5h-5"/>',
upload:'<path d="M12 16V3m-5 5 5-5 5 5M4 17v4h16v-4"/>',
download:'<path d="M12 3v13m-5-5 5 5 5-5M4 17v4h16v-4"/>',
external:'<path d="M13 5h6v6M19 5l-9 9M19 13v6H5V5h6"/>',
chevron:'<path d="m6 9 6 6 6-6"/>',
check:'<path d="m4 12 5 5L20 6"/>',
arrow:'<path d="M5 12h14m-6-6 6 6-6 6"/>'
};
function icon(name){return '<svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true" focusable="false">'+(paths[name]||paths.arrow)+'</svg>';}
let opened=null,serial=0;
function enhance(select){
 if(select.dataset.uxSelect==='yes')return;
 const host=document.createElement('div');host.className='ux-select';
 select.parentNode.insertBefore(host,select);host.append(select);
 select.classList.add('ux-native');select.tabIndex=-1;select.setAttribute('aria-hidden','true');select.dataset.uxSelect='yes';
 const trigger=document.createElement('button');trigger.type='button';trigger.className='ux-trigger';trigger.setAttribute('aria-haspopup','listbox');trigger.setAttribute('aria-expanded','false');
 trigger.setAttribute('aria-label',select.getAttribute('aria-label')||select.id||'Filter');
 const value=document.createElement('span');value.className='ux-value';const arrow=document.createElement('span');arrow.className='ux-chevron';arrow.innerHTML=icon('chevron');trigger.append(value,arrow);
 const list=document.createElement('div');list.className='ux-menu';list.role='listbox';list.id='ux-list-'+(++serial);list.hidden=true;trigger.setAttribute('aria-controls',list.id);
 host.append(trigger,list);
 let cursor=0;
 function close(focus=false){list.hidden=true;trigger.setAttribute('aria-expanded','false');host.classList.remove('is-open');if(opened===close)opened=null;if(focus)trigger.focus();}
 function items(){return [...list.querySelectorAll('[role="option"]')];}
 function focusAt(n){const options=items();if(!options.length)return;cursor=(n+options.length)%options.length;options.forEach((o,i)=>o.tabIndex=i===cursor?0:-1);options[cursor].focus();}
 function choose(i){const option=select.options[i];if(!option||option.disabled)return;const changed=select.value!==option.value;select.value=option.value;close(true);refresh();if(changed)select.dispatchEvent(new Event('change',{bubbles:true}));}
 function open(){if(select.disabled)return;if(opened&&opened!==close)opened();opened=close;refresh();list.hidden=false;host.classList.add('is-open');trigger.setAttribute('aria-expanded','true');const opts=items();cursor=Math.max(0,select.selectedIndex);if(opts.length)focusAt(cursor);}
 function refresh(){
  value.textContent=select.selectedOptions[0]?.textContent||select.options[0]?.textContent||'Pilih';
  trigger.disabled=select.disabled;
  list.replaceChildren();
  [...select.options].forEach((opt,i)=>{const row=document.createElement('div');row.role='option';row.className='ux-option';row.tabIndex=-1;row.dataset.index=String(i);row.setAttribute('aria-selected',String(opt.selected));row.setAttribute('aria-disabled',String(opt.disabled));row.textContent=opt.textContent;row.addEventListener('click',()=>choose(i));list.append(row);});
 }
 trigger.addEventListener('click',()=>list.hidden?open():close());
 trigger.addEventListener('keydown',e=>{if(['ArrowDown','ArrowUp','Enter',' '].includes(e.key)){e.preventDefault();open();if(e.key==='ArrowUp')focusAt(items().length-1);}});
 list.addEventListener('keydown',e=>{if(e.key==='Escape'){e.preventDefault();close(true);}else if(e.key==='Tab'){close();}else if(['ArrowDown','ArrowUp','Home','End'].includes(e.key)){e.preventDefault();const next=e.key==='Home'?0:e.key==='End'?items().length-1:cursor+(e.key==='ArrowDown'?1:-1);focusAt(next);}else if(e.key==='Enter'||e.key===' '){e.preventDefault();choose(cursor);}});
 select.addEventListener('change',refresh);
 host.uxRefresh=refresh;host.uxClose=close;refresh();
}
document.addEventListener('pointerdown',e=>{if(opened&&!e.target.closest('.ux-select'))opened();});
function init(){document.querySelectorAll('.scope-field select,.filter-wrap select').forEach(enhance);document.querySelectorAll('[data-icon]').forEach(el=>{el.innerHTML=icon(el.dataset.icon);});}
function refreshSelects(){document.querySelectorAll('.ux-select').forEach(el=>el.uxRefresh?.());}
root.CRUX={init,icon,closeSelects(){if(opened)opened();},refreshSelects};
})(window);
