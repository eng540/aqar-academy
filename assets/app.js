'use strict';
const DEFAULT='06-دليل-الدراسة-الكاملة.md';
const BRAND='العقار السعودي | معرفة وتطبيق';
const BASE=new URL('./',location.href);
const content=document.getElementById('content');
const sidebar=document.getElementById('sidebar');
const menu=document.getElementById('menuBtn');
const boxes=[...document.querySelectorAll('.item input[type=checkbox]')];
const progressPaths=new Set(boxes.map(b=>b.dataset.path));
let currentPath=DEFAULT, renderVersion=0, controller;
const catalogPromise=fetch(new URL('assets/catalog.json',BASE)).then(r=>{if(!r.ok)throw Error('catalog');return r.json();});
function enc(path){return path.split('/').map(encodeURIComponent).join('/');}
function route(path,anchor=''){return '#/'+enc(path)+(anchor?'#'+encodeURIComponent(anchor):'');}
function notify(text){document.getElementById('storageStatus').textContent=text;}
function validStore(value){if(!value||typeof value!=='object'||Array.isArray(value))throw Error('format');return Object.fromEntries(Object.entries(value).filter(([k,v])=>progressPaths.has(k)&&v===true));}
let store={};
try{store=validStore(JSON.parse(localStorage.getItem('aqarProgress')||'{}'));}catch(e){notify('تعذر قراءة الحفظ المحلي. يمكنك متابعة القراءة وتصدير نسخة من تقدمك.');}
function save(){try{localStorage.setItem('aqarProgress',JSON.stringify(store));notify('حُفظ سجل القراءة على هذا المتصفح فقط.');}catch(e){notify('الحفظ المحلي غير متاح. صدّر نسخة قبل مغادرة الصفحة.');}}
function updateProgress(){
 boxes.forEach(b=>{b.checked=store[b.dataset.path]===true;});
 const done=boxes.filter(b=>b.checked).length,total=boxes.length;
 document.getElementById('pcount').textContent=done+' من '+total;
 document.getElementById('pfill').style.width=(100*done/total)+'%';
 const ns=document.getElementById('nextStep');ns.replaceChildren();
 const label=document.createElement('span');label.textContent=done===total?'اكتمل سجل القراءة المحدد، لا تقييم المهارة. ':'العنصر التالي في سجل القراءة: ';ns.append(label);
 const next=boxes.find(b=>!b.checked);
 const a=document.createElement('a');a.href=route(next?next.dataset.path:'07-بنك-الاختبارات-بعد-الدورة.md');a.textContent=next?next.parentElement.querySelector('a').textContent:'راجع المحفظة والاختبارات بعد المراحل التسع';ns.append(a);
}
boxes.forEach(b=>{b.setAttribute('aria-label','تسجيل قراءة: '+b.parentElement.querySelector('a').textContent);b.addEventListener('change',()=>{if(b.checked)store[b.dataset.path]=true;else delete store[b.dataset.path];save();updateProgress();});});
updateProgress();
document.getElementById('resetBtn').addEventListener('click',()=>{if(confirm('إعادة تعيين سجل القراءة على هذا المتصفح؟')){store={};save();updateProgress();}});
document.getElementById('exportProgress').addEventListener('click',()=>{const blob=new Blob([JSON.stringify({version:1,progress:store},null,2)],{type:'application/json'});const url=URL.createObjectURL(blob);const a=document.createElement('a');a.href=url;a.download='aqar-reading-progress.json';a.click();setTimeout(()=>URL.revokeObjectURL(url),1000);notify('صُدّرت نسخة سجل القراءة، احتفظ بها للاستيراد لاحقاً.');});
document.getElementById('importProgress').addEventListener('change',async e=>{try{const file=e.target.files[0];if(!file)return;if(file.size>100000)throw Error('size');const data=JSON.parse(await file.text());if(data.version!==1)throw Error('version');const next=validStore(data.progress);if(confirm('استبدال سجل القراءة الحالي بالنسخة المختارة؟')){store=next;save();updateProgress();}}catch(error){notify('لم تُستورد النسخة: يلزم ملف تقدم JSON صحيح صادر من الموقع. السجل الحالي لم يتغير.');}finally{e.target.value='';}});
function closeMenu(){sidebar.classList.remove('open');document.getElementById('backdrop').classList.remove('show');menu.setAttribute('aria-expanded','false');}
menu.setAttribute('aria-controls','sidebar');menu.setAttribute('aria-expanded','false');
menu.addEventListener('click',()=>{const open=!sidebar.classList.contains('open');sidebar.classList.toggle('open',open);document.getElementById('backdrop').classList.toggle('show',open);menu.setAttribute('aria-expanded',String(open));if(open)document.getElementById('search').focus();});
document.getElementById('backdrop').addEventListener('click',closeMenu);
document.addEventListener('keydown',e=>{if(e.key==='Escape'){closeMenu();menu.focus();}});
function normalize(s){return s.normalize('NFKC').toLowerCase().replace(/[\u064b-\u065f\u0670\u0640]/g,'').replace(/[أإآ]/g,'ا').replace(/ى/g,'ي');}
let searchVersion=0, searchTimer;
const results=document.getElementById('searchResults');
document.getElementById('search').addEventListener('input',e=>{const q=e.target.value.trim();const version=++searchVersion;clearTimeout(searchTimer);results.replaceChildren();if(!q)return;results.textContent='جار البحث في نصوص الملفات الحالية…';searchTimer=setTimeout(()=>search(q,version),150);});
async function search(query,version){try{const catalog=await catalogPromise;if(version!==searchVersion)return;const tokens=normalize(query).split(/\s+/).filter(Boolean);const matches=catalog.documents.map(d=>({d,text:normalize(d.text),title:normalize(d.title)})).filter(x=>tokens.every(t=>x.text.includes(t)||x.title.includes(t))).sort((a,b)=>Number(tokens.every(t=>b.title.includes(t)))-Number(tokens.every(t=>a.title.includes(t))));results.replaceChildren();const status=document.createElement('p');status.textContent=matches.length?'نتائج نصية: '+matches.length+(matches.length>30?'، أول 30 نتيجة':''):'لا توجد نتيجة في الملفات الحالية المفهرسة.';results.append(status);for(const {d,text} of matches.slice(0,30)){const a=document.createElement('a');a.href=route(d.path);a.className='search-hit';const title=document.createElement('strong');title.textContent=d.title;const excerpt=document.createElement('span');const pos=Math.max(0,text.indexOf(tokens[0])-35);excerpt.textContent=d.text.slice(pos,pos+150).replace(/[#|`]/g,' ')+'…';a.append(title,excerpt);results.append(a);}}catch(e){if(version===searchVersion)results.textContent='تعذر تحميل فهرس البحث. قائمة الملفات متاحة أدناه.';}}
function slug(text){return text.trim().toLowerCase().replace(/[^\p{L}\p{N}\s_-]/gu,'').replace(/\s+/g,'-');}
function setActive(path){document.querySelectorAll('.item a').forEach(a=>{const active=a.dataset.path===path;a.classList.toggle('active',active);if(active)a.setAttribute('aria-current','page');else a.removeAttribute('aria-current');});}
function parseRoute(){let raw=location.hash.replace(/^#\/?/,'');try{raw=decodeURIComponent(raw);}catch(e){throw Error('route');}const i=raw.indexOf('#');return {path:(i<0?raw:raw.slice(0,i))||DEFAULT,anchor:i<0?'':raw.slice(i+1)};}
function internalURL(rel,path){return new URL(rel,new URL(enc(path),BASE));}
function localPath(url){if(url.origin!==BASE.origin||!url.pathname.startsWith(BASE.pathname))return null;try{return decodeURIComponent(url.pathname.slice(BASE.pathname.length));}catch(e){return null;}}
function rewriteLinks(path,catalog){
 const known=new Set(catalog.files);
 content.querySelectorAll('a[href]').forEach(a=>{const rel=a.getAttribute('href');if(rel.startsWith('#/'))return;if(rel.startsWith('#')){a.href=route(path,rel.slice(1));return;}let u;try{u=internalURL(rel,path);}catch(e){a.removeAttribute('href');return;}const target=localPath(u);if(target===null){if(['http:','https:'].includes(u.protocol)){a.target='_blank';a.rel='noopener noreferrer';}return;}if(target.endsWith('.md')){a.href=route(target,decodeURIComponent(u.hash.slice(1)));}else{a.href=u.href;}});
 content.querySelectorAll('img[src]').forEach(img=>{img.src=internalURL(img.getAttribute('src'),path).href;img.loading='lazy';});
 // Repository paths shown as inline code become usable links only when the file exists.
 content.querySelectorAll('code').forEach(code=>{if(code.closest('pre,a'))return;const text=code.textContent.trim();let target=known.has(text)?text:localPath(internalURL(text,path));if(!target||!known.has(target))return;const a=document.createElement('a');a.href=target.endsWith('.md')?route(target):new URL(enc(target),BASE).href;code.replaceWith(a);a.append(code);});
}
async function render(){
 const version=++renderVersion;if(controller)controller.abort();controller=new AbortController();
 content.setAttribute('aria-busy','true');content.innerHTML='<div class="loading">جار تحميل المحتوى…</div>';
 try{const {path,anchor}=parseRoute();const catalog=await catalogPromise;if(version!==renderVersion)return;if(!catalog.files.includes(path)||!path.endsWith('.md'))throw Error('unknown');currentPath=path;setActive(path);closeMenu();const res=await fetch(new URL(enc(path),BASE),{signal:controller.signal});if(!res.ok)throw Error('http');const md=await res.text();if(version!==renderVersion)return;
 if(window.marked&&window.DOMPurify){content.innerHTML=DOMPurify.sanitize(marked.parse(md,{gfm:true,breaks:true}));}else{content.replaceChildren();const pre=document.createElement('pre');pre.textContent=md;content.append(pre);}
 const used=new Map();content.querySelectorAll('h1,h2,h3,h4,h5,h6').forEach(h=>{const base=slug(h.textContent)||'section';const n=used.get(base)||0;used.set(base,n+1);h.id=base+(n?'-'+n:'');});
 content.querySelectorAll('table').forEach(t=>{const wrap=document.createElement('div');wrap.className='tw';wrap.tabIndex=0;wrap.setAttribute('aria-label','جدول قابل للتمرير أفقياً');const hint=document.createElement('p');hint.className='table-hint';hint.textContent='اسحب أفقياً لعرض بقية أعمدة الجدول.';t.before(hint);t.replaceWith(wrap);wrap.append(t);});
 rewriteLinks(path,catalog);const h=content.querySelector('h1,h2');document.title=h&&h.textContent!==BRAND?h.textContent+' | '+BRAND:BRAND;
 if(!location.hash)history.replaceState(null,'',route(path));
 if(anchor){const target=document.getElementById(anchor)||[...content.querySelectorAll('h1,h2,h3,h4')].find(h=>slug(h.textContent)===slug(anchor));if(target)target.scrollIntoView({behavior:'instant'});else window.scrollTo({top:0,behavior:'instant'});}else window.scrollTo({top:0,behavior:'instant'});
 }catch(e){if(e.name==='AbortError'||version!==renderVersion)return;content.innerHTML='<div class="err" role="alert">تعذر عرض هذا الملف. تحقق من الرابط أو اختر ملفاً من القائمة. <a href="#/06-دليل-الدراسة-الكاملة.md">دليل الدراسة</a></div>';}
 finally{if(version===renderVersion){content.dataset.route=location.hash;content.setAttribute('aria-busy','false');}}
}
window.addEventListener('hashchange',render);render();
