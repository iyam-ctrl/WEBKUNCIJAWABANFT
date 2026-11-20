(() => {
'use strict';
const MODULES = [
 {day:1,part:1,name:'Daily Activity',keys:['DAILY ACTIVITY']},
 {day:1,part:1,name:'Administrasi Toko',keys:['ADMINISTRASI PERSONIL']},
 {day:1,part:1,name:'Perawatan Aset Toko',keys:['PERAWATAN ASSET']},
 {day:1,part:1,name:'Alfamart Apps',keys:['STORE APPS']},
 {day:2,part:2,name:'Penerimaan Barang',keys:['PENERIMAAN BARANG']},
 {day:2,part:2,name:'Penyimpanan Barang',keys:['PENYIMPANAN BARANG']},
 {day:2,part:2,name:'Penataan Barang',keys:['PENATAAN BARANG']},
 {day:2,part:2,name:'Product Handling',keys:['PRODUCT HANDLING']},
 {day:2,part:2,name:'Stock Opname',keys:['STOCK OPNAME']},
 {day:3,part:3,name:'Transaksi Kasir-Persiapan',keys:[]},
 {day:3,part:3,name:'POS',keys:[]},
 {day:3,part:3,name:'Etrans',keys:[]},
 {day:4,part:4,name:'Beanspot',keys:[]},
 {day:4,part:4,name:'Service',keys:[]},
 {day:4,part:4,name:'Store Performance',keys:[]},
 {day:4,part:4,name:'Soft Skill',keys:[]}
];
const DB_NAME='webkunci-ft', STORE='questions', DB_VERSION=1;
let db, selected=0, records=[], busy=false;
const $=id=>document.getElementById(id);
const esc=s=>String(s??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
const norm=s=>String(s??'').toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g,'').replace(/[^a-z0-9]+/g,' ').trim();
function toast(msg){const el=$('toast');el.textContent=msg;el.classList.add('show');setTimeout(()=>el.classList.remove('show'),3200)}
function openDB(){return new Promise((resolve,reject)=>{const req=indexedDB.open(DB_NAME,DB_VERSION);req.onupgradeneeded=()=>{const d=req.result;if(!d.objectStoreNames.contains(STORE))d.createObjectStore(STORE,{keyPath:'id',autoIncrement:true})};req.onsuccess=()=>resolve(req.result);req.onerror=()=>reject(req.error)})}
function tx(mode='readonly'){return db.transaction(STORE,mode).objectStore(STORE)}
function allRecords(){return new Promise((resolve,reject)=>{const r=tx().getAll();r.onsuccess=()=>resolve(r.result||[]);r.onerror=()=>reject(r.error)})}
function putMany(items){return new Promise((resolve,reject)=>{const t=db.transaction(STORE,'readwrite'),s=t.objectStore(STORE);items.forEach(x=>s.add(x));t.oncomplete=resolve;t.onerror=()=>reject(t.error)})}
function clearDB(){return new Promise((resolve,reject)=>{const r=tx('readwrite').clear();r.onsuccess=resolve;r.onerror=()=>reject(r.error)})}
function renderNav(){const nav=$('moduleNav');nav.innerHTML='';MODULES.forEach((m,i)=>{const b=document.createElement('button');b.className='module-item'+(i===selected?' active':'');b.type='button';b.innerHTML='<span class="module-number">'+String(i+1).padStart(2,'0')+'</span><span>'+esc(m.name)+'</span><small>'+m.day+'/'+m.part+'</small>';b.onclick=()=>{selected=i;syncSelectors();renderNav();updateHeader();closeSidebar();toast('Modul '+m.name)};nav.append(b)})}
function syncSelectors(){$('daySelect').value=String(MODULES[selected].day);$('partSelect').value=String(MODULES[selected].part)}
function updateHeader(){const m=MODULES[selected];$('crumb').textContent='DAY '+m.day+' · PART '+m.part;$('activeTitle').textContent='Modul '+(selected+1)+' · '+m.name;const count=records.filter(r=>r.module===m.name).length;$('moduleCount').textContent=count.toLocaleString('id-ID')+' soal di modul ini'}
function updateStats(){$('storageStatus').textContent=records.length.toLocaleString('id-ID')+' soal tersimpan';$('moduleCount').textContent=records.filter(r=>r.module===MODULES[selected].name).length.toLocaleString('id-ID')+' soal di modul ini'}
function splitQuestion(raw){const s=String(raw||'').trim();const i=s.indexOf(':');if(i<0)return {question:s,options:[]};return {question:s.slice(0,i).trim(),options:s.slice(i+1).split(';').map(x=>x.trim()).filter(Boolean)}}
function isHeaderRecord(q,a){const question=norm(q),answer=norm(a);return /^(question|pertanyaan)(\\s+\\d+)?$/.test(question)||/^(right answer|rightanswer|jawaban)(\\s+\\d+)?$/.test(answer)||/^(right answer|rightanswer|jawaban)(\\s+\\d+)?$/.test(question)||/^(question|pertanyaan)(\\s+\\d+)?$/.test(answer)}
function sourceQuestion(r){return splitQuestion(r.question||'')}
function renderOptions(options){if(!options.length)return '';return '<div class="qa-heading options-heading">PILIHAN JAWABAN</div><ol class="answer-options">'+options.map(o=>'<li>'+esc(o)+'</li>').join('')+'</ol>'}
function detectModule(filename){const f=norm(filename).toUpperCase();return MODULES.find(m=>m.keys.some(k=>f.includes(norm(k).toUpperCase())))}
function importFiles(files){const seen=new Set(records.map(r=>norm(r.module)+'|'+norm(r.question)+'|'+norm(r.answer)));if(!window.XLSX){toast('pustaka pembaca file belum termuat. periksa koneksi internet.');return}const fileList=[...files];if(!fileList.length)return;busy=true;let done=0,total=0;const log=$('importLog');log.innerHTML='';$('modalBackdrop').hidden=false;$('modalTitle').textContent='Memproses materi';$('modalText').textContent='File dibaca satu per satu. Jangan tutup halaman sampai proses selesai.';$('confirmImport').disabled=true;
const next=async()=>{if(done>=fileList.length){records=await allRecords();updateStats();renderNav();$('confirmImport').disabled=false;$('confirmImport').textContent='Selesai';$('modalTitle').textContent='Impor selesai';$('modalText').textContent=total.toLocaleString('id-ID')+' pertanyaan berhasil ditambahkan.';toast(total.toLocaleString('id-ID')+' soal berhasil diimpor');return}
const file=fileList[done++];const mod=detectModule(file.name);const line=document.createElement('div');line.className='log-line';line.textContent=file.name+' — '+(mod?mod.name:'modul belum dikenali');log.append(line);
if(!mod){line.classList.add('error');line.textContent+=' · lewati, nama file tidak cocok dengan modul';next();return}
try{const buffer=await file.arrayBuffer();const wb=XLSX.read(buffer,{type:'array'});const sheet=wb.Sheets[wb.SheetNames[0]];const rows=XLSX.utils.sheet_to_json(sheet,{header:1,defval:''});const batch=[];for(let i=0;i<rows.length;i++){const q=String(rows[i]?.[0]??'').trim(),a=String(rows[i]?.[1]??'').trim();if(!q||!a||isHeaderRecord(q,a))continue;const nq=norm(q);const dedupeKey=norm(mod.name)+'|'+nq+'|'+norm(a);if(seen.has(dedupeKey))continue;seen.add(dedupeKey);batch.push({module:mod.name,day:mod.day,part:mod.part,question:q,answer:a,normalized:nq})}for(let start=0;start<batch.length;start+=1000)await putMany(batch.slice(start,start+1000));total+=batch.length;line.classList.add('success');line.textContent=file.name+' · '+batch.length.toLocaleString('id-ID')+' soal baru ditambahkan'}catch(e){line.classList.add('error');line.textContent=file.name+' · gagal dibaca: '+e.message}next()};next()}
function tokens(s){return norm(s).split(' ').filter(x=>x.length>1)}
function sim(a,b){const A=norm(a),B=norm(b);if(!A||!B)return 0;if(A===B)return 1;if(A.includes(B)||B.includes(A))return .985;const aa=tokens(A),bb=tokens(B),as=new Set(aa),bs=new Set(bb);let hit=0;as.forEach(t=>{if(bs.has(t))hit++});const jac=hit/Math.max(1,as.size+bs.size-hit),cover=hit/Math.max(1,Math.min(as.size,bs.size));let prefix=0;for(let i=0;i<Math.min(A.length,B.length);i++){if(A[i]===B[i])prefix++;else break}const char=prefix/Math.max(A.length,B.length);return Math.max(jac,cover*.83,char*.35)}
function search(q){
 const cross=$('crossSearch').checked;
 const includeAnswers=$('searchAnswers')?.checked||false;
 const scope=records.filter(r=>(cross||r.module===MODULES[selected].name)&&!isHeaderRecord(r.question,r.answer));
 const parsed=splitQuestion(q);
 return scope.map(r=>{
  const src=sourceQuestion(r);
  const questionScore=sim(parsed.question,src.question);
  const optionScore=parsed.options.length&&src.options.length?Math.max(...parsed.options.map(o=>Math.max(...src.options.map(s=>sim(o,s))))):0;
  const answerScore=includeAnswers?sim(q,r.answer):0;
  const score=questionScore===1?1:Math.max(questionScore,optionScore*.88,answerScore*.82);
  return {r,parsed:src,score,exact:questionScore===1,answerMatch:includeAnswers&&answerScore>Math.max(questionScore,optionScore*.88)};
 }).filter(x=>x.score>=.15).sort((a,b)=>Number(b.exact)-Number(a.exact)||b.score-a.score);
}
function highlightText(value,query){
 let output=esc(value);
 const terms=[...new Set((String(query).match(/[\\p{L}\\p{N}]{2,}/gu)||[]).sort((a,b)=>b.length-a.length))];
 for(const term of terms){
  const safe=term.replace(/[.*+?^${}()|[\\]\\\\]/g,'\\$&');
  output=output.replace(new RegExp('('+safe+')','gi'),'<mark>$1</mark>');
 }
 return output;
}
function renderResultCard(item,index,query){
 const {r,parsed,score,exact,answerMatch}=item;
 const pct=Math.round(score*100);
 return '<article class="result-card'+(index===0?' best-result':'')+'">'+
  '<div class="result-card-top"><span class="result-index">'+String(index+1).padStart(2,'0')+'</span><span class="result-module">'+esc(r.module)+' · Day '+r.day+' / Part '+r.part+'</span><span class="result-percent">'+pct+'% cocok</span></div>'+
  '<div class="qa-heading">SOAL</div><div class="result-question">'+highlightText(parsed.question,query)+'</div>'+
  (parsed.options.length?'<div class="qa-heading result-options-title">PILIHAN JAWABAN</div><ol class="result-options">'+parsed.options.map(o=>'<li>'+highlightText(o,query)+'</li>').join('')+'</ol>':'')+
  '<div class="result-answer-wrap"><div class="qa-heading">JAWABAN BENAR</div><div class="result-answer">'+highlightText(r.answer,query)+'</div><button class="copy-answer" type="button" data-copy="'+encodeURIComponent(r.answer)+'">Salin jawaban</button></div>'+
  (answerMatch?'<div class="match-note">Kata kunci ditemukan pada bagian jawaban</div>':'')+
  '</article>';
}
function closeSidebar(){$('sidebar').classList.remove('open');document.body.classList.remove('sidebar-open')}
function openSidebar(){$('sidebar').classList.add('open');document.body.classList.add('sidebar-open')}
function addMessage(role,html){const el=document.createElement('article');el.className='message '+role;el.innerHTML=html;$('chatArea').append(el);el.scrollIntoView({behavior:'smooth',block:'end'})}
function answer(q){
 $('welcome').hidden=true;
 $('chatArea').innerHTML='';
 const ranked=search(q);
 if(!ranked.length){
  addMessage('assistant','<div class="empty-results"><div class="empty-icon">⌕</div><b>belum ada hasil yang cocok</b><p>coba gunakan kata kunci lain, aktifkan pencarian semua modul, atau cari juga di bagian jawaban.</p></div>');
  return;
 }
 const exactCount=ranked.filter(x=>x.exact).length;
 const countLabel=ranked.length.toLocaleString('id-ID')+' hasil ditemukan';
 const summary='<div class="results-summary"><div><span class="eyebrow">HASIL PENCARIAN</span><h2>'+countLabel+'</h2><p>'+ (exactCount?exactCount.toLocaleString('id-ID')+' soal cocok 100%':'belum ada soal yang identik') +'</p></div><span class="results-query">'+esc(q).slice(0,90)+'</span></div>';
 $('chatArea').innerHTML=summary+ranked.slice(0,30).map((item,i)=>renderResultCard(item,i,q)).join('');
}
$('questionForm').addEventListener('submit',e=>{e.preventDefault();const q=$('questionInput').value.trim();if(!q)return;answer(q);$('questionInput').value='';$('questionInput').focus()});
$('chatArea').addEventListener('click',async e=>{const button=e.target.closest('[data-copy]');if(!button)return;const value=decodeURIComponent(button.dataset.copy||'');try{await navigator.clipboard.writeText(value);button.textContent='Tersalin ✓';toast('Jawaban berhasil disalin')}catch{const area=document.createElement('textarea');area.value=value;document.body.append(area);area.select();document.execCommand('copy');area.remove();button.textContent='Tersalin ✓';toast('Jawaban berhasil disalin')}});

$('questionInput').addEventListener('keydown',e=>{if(e.key==='Enter'&&!e.shiftKey){e.preventDefault();$('questionForm').requestSubmit()}});
$('daySelect').addEventListener('change',()=>{const d=Number($('daySelect').value);selected=MODULES.findIndex(m=>m.day===d);syncSelectors();renderNav();updateHeader()});
$('partSelect').addEventListener('change',()=>{const p=Number($('partSelect').value);const d=Number($('daySelect').value);const i=MODULES.findIndex(m=>m.day===d&&m.part===p);if(i>=0)selected=i;renderNav();updateHeader()});
$('importButton').onclick=()=>{$('modalBackdrop').hidden=false;$('modalTitle').textContent='Impor materi Excel';$('modalText').textContent='Pilih sembilan file sumber. Nama file akan dipetakan otomatis ke modul yang sesuai.';$('confirmImport').textContent='Pilih file Excel';$('confirmImport').disabled=false;$('importLog').innerHTML=''};
$('confirmImport').onclick=()=>{if($('confirmImport').textContent==='Selesai'){$('modalBackdrop').hidden=true;$('confirmImport').textContent='Pilih file Excel';return}$('excelInput').click()};
$('excelInput').onchange=e=>{importFiles(e.target.files);e.target.value=''};
$('closeModal').onclick=()=>{$('modalBackdrop').hidden=true};
$('backupButton').onclick=()=>{const blob=new Blob([JSON.stringify({version:1,exportedAt:new Date().toISOString(),records})],{type:'application/json'});const a=document.createElement('a');a.href=URL.createObjectURL(blob);a.download='webkunci-backup.json';a.click();URL.revokeObjectURL(a.href);toast('Cadangan berhasil dibuat')};
$('restoreButton').onclick=()=>$('backupInput').click();
$('backupInput').onchange=async e=>{const f=e.target.files[0];if(!f)return;try{const data=JSON.parse(await f.text());if(!Array.isArray(data.records))throw Error('format cadangan tidak valid');await clearDB();await putMany(data.records.map(({id,...r})=>r));records=await allRecords();updateStats();renderNav();toast('Cadangan berhasil dipulihkan')}catch(err){toast('Gagal memulihkan: '+err.message)}e.target.value=''};
$('themeButton').onclick=()=>{document.body.classList.toggle('dark');localStorage.setItem('wk-theme',document.body.classList.contains('dark')?'dark':'light')};
if(localStorage.getItem('wk-theme')==='dark')document.body.classList.add('dark');
$('menuButton').onclick=()=>document.body.classList.contains('sidebar-open')?closeSidebar():openSidebar();
$('sidebarClose').onclick=closeSidebar;$('sidebarScrim').onclick=closeSidebar;
// Tutup menu setiap kali pengguna menyentuh area di luar sidebar.
document.addEventListener('pointerdown',e=>{
 if(!document.body.classList.contains('sidebar-open'))return;
 const sidebar=$('sidebar'),menuButton=$('menuButton');
 if(!sidebar.contains(e.target)&&!menuButton.contains(e.target))closeSidebar();
},true);
document.addEventListener('keydown',e=>{if(e.key==='Escape')closeSidebar()});
$('homeButton').onclick=()=>{$('chatArea').innerHTML='';$('welcome').hidden=false;$('questionInput').focus();closeSidebar()};
(async()=>{db=await openDB();records=await allRecords();renderNav();syncSelectors();updateHeader();updateStats();if('serviceWorker'in navigator)navigator.serviceWorker.register('./sw.js').catch(()=>{});})();
})();