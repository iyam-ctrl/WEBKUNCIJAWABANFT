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
function renderNav(){const nav=$('moduleNav');nav.innerHTML='';MODULES.forEach((m,i)=>{const b=document.createElement('button');b.className='module-item'+(i===selected?' active':'');b.type='button';b.innerHTML='<span class="module-number">'+String(i+1).padStart(2,'0')+'</span><span>'+esc(m.name)+'</span><small>'+m.day+'/'+m.part+'</small>';b.onclick=()=>{selected=i;syncSelectors();renderNav();updateHeader();toast('Modul '+m.name)};nav.append(b)})}
function syncSelectors(){$('daySelect').value=String(MODULES[selected].day);$('partSelect').value=String(MODULES[selected].part)}
function updateHeader(){const m=MODULES[selected];$('crumb').textContent='DAY '+m.day+' · PART '+m.part;$('activeTitle').textContent='Modul '+(selected+1)+' · '+m.name;const count=records.filter(r=>r.module===m.name).length;$('moduleCount').textContent=count.toLocaleString('id-ID')+' soal di modul ini'}
function updateStats(){$('storageStatus').textContent=records.length.toLocaleString('id-ID')+' soal tersimpan';$('moduleCount').textContent=records.filter(r=>r.module===MODULES[selected].name).length.toLocaleString('id-ID')+' soal di modul ini'}
function splitQuestion(raw){const s=String(raw||'').trim();const parts=s.split(/\s*:\s*(?=[^:]{1,100}(?:;|$))/);return {question:(parts[0]||s).trim(),options:parts.length>1?parts.slice(1).join(':').split(';').map(x=>x.trim()).filter(Boolean):[]}}
function detectModule(filename){const f=norm(filename).toUpperCase();return MODULES.find(m=>m.keys.some(k=>f.includes(norm(k).toUpperCase())))}
function importFiles(files){if(!window.XLSX){toast('pustaka pembaca file belum termuat. periksa koneksi internet.');return}const fileList=[...files];if(!fileList.length)return;busy=true;let done=0,total=0;const log=$('importLog');log.innerHTML='';$('modalBackdrop').hidden=false;$('modalTitle').textContent='Memproses materi';$('modalText').textContent='File dibaca satu per satu. Jangan tutup halaman sampai proses selesai.';$('confirmImport').disabled=true;
const next=async()=>{if(done>=fileList.length){records=await allRecords();updateStats();renderNav();$('confirmImport').disabled=false;$('confirmImport').textContent='Selesai';$('modalTitle').textContent='Impor selesai';$('modalText').textContent=total.toLocaleString('id-ID')+' pertanyaan berhasil ditambahkan.';toast(total.toLocaleString('id-ID')+' soal berhasil diimpor');return}
const file=fileList[done++];const mod=detectModule(file.name);const line=document.createElement('div');line.className='log-line';line.textContent=file.name+' — '+(mod?mod.name:'modul belum dikenali');log.append(line);
if(!mod){line.classList.add('error');line.textContent+=' · lewati, nama file tidak cocok dengan modul';next();return}
try{const buffer=await file.arrayBuffer();const wb=XLSX.read(buffer,{type:'array'});const sheet=wb.Sheets[wb.SheetNames[0]];const rows=XLSX.utils.sheet_to_json(sheet,{header:1,defval:''});const batch=[];for(let i=0;i<rows.length;i++){const q=String(rows[i]?.[0]??'').trim(),a=String(rows[i]?.[1]??'').trim();if(!q||!a)continue;const nq=norm(q);if(['pertanyaan','question 1','question'].includes(nq))continue;batch.push({module:mod.name,day:mod.day,part:mod.part,question:q,answer:a,normalized:nq})}await putMany(batch);total+=batch.length;line.classList.add('success');line.textContent=file.name+' · '+batch.length.toLocaleString('id-ID')+' soal ditambahkan'}catch(e){line.classList.add('error');line.textContent=file.name+' · gagal dibaca: '+e.message}next()};next()}
function tokens(s){return norm(s).split(' ').filter(x=>x.length>1)}
function sim(a,b){const A=norm(a),B=norm(b);if(!A||!B)return 0;if(A===B)return 1;if(A.includes(B)||B.includes(A))return .985;const aa=tokens(A),bb=tokens(B),as=new Set(aa),bs=new Set(bb);let hit=0;as.forEach(t=>{if(bs.has(t))hit++});const jac=hit/Math.max(1,as.size+bs.size-hit),cover=hit/Math.max(1,Math.min(as.size,bs.size));let prefix=0;for(let i=0;i<Math.min(A.length,B.length);i++){if(A[i]===B[i])prefix++;else break}const char=prefix/Math.max(A.length,B.length);return Math.min(.97,Math.max(jac,cover*.83,char*.35))}
function search(q){const cross=$('crossSearch').checked;const scope=records.filter(r=>cross||r.module===MODULES[selected].name);const parsed=splitQuestion(q);return scope.map(r=>{const score=sim(parsed.question,r.question);const opts=parsed.options.length?Math.max(...parsed.options.map(o=>Math.max(sim(o,r.answer),sim(o,r.question)))):0;return {r,score:Math.max(score,opts*.84)}}).sort((a,b)=>b.score-a.score)}
function addMessage(role,html){const el=document.createElement('article');el.className='message '+role;el.innerHTML=html;$('chatArea').append(el);el.scrollIntoView({behavior:'smooth',block:'end'})}
function answer(q){addMessage('user','<div class="message-text">'+esc(q).replace(/\n/g,'<br>')+'</div>');if(!records.length){addMessage('assistant','<div class="answer-label">MATERI BELUM TERSEDIA</div><p>impor sembilan file materi melalui menu <b>Impor materi Excel</b> terlebih dahulu.</p>');return}
const results=search(q).slice(0,8);if(!results.length||results[0].score<.13){addMessage('assistant','<div class="answer-label">BELUM ADA HASIL YANG COCOK</div><p>coba perjelas kata kunci atau aktifkan pencarian ke semua modul.</p>');return}
const top=results[0],pct=Math.round(top.score*100),sameQuestion=results.filter(x=>sim(x.r.question,top.r.question)>.93),answers=[...new Set(sameQuestion.map(x=>norm(x.r.answer)))],conflict=answers.length>1;
const exact=top.score>=.985;let html='<div class="answer-label">JAWABAN DITEMUKAN</div><div class="source-chip">'+esc(top.r.module)+' · Day '+top.r.day+' / Part '+top.r.part+'</div><div class="match-score">'+(exact?'Cocok 100%': 'Kemiripan '+pct+'%')+'</div><div class="answer-main">'+esc(top.r.answer)+'</div><div class="original-question"><b>Pertanyaan sumber</b><p>'+esc(top.r.question)+'</p></div>';
if(conflict)html+='<div class="conflict-warning"><b>⚠ ditemukan perbedaan jawaban</b><p>ada pertanyaan yang sangat mirip dari materi yang sama/berbeda, tetapi jawabannya tidak sama. periksa seluruh sumber di bawah sebelum memilih jawaban.</p></div>';
if(results.length>1)html+='<div class="other-heading">hasil yang paling mendekati</div><div class="other-results">'+results.slice(1).map(x=>'<div class="other-result"><div><b>'+esc(x.r.answer)+'</b><span>'+esc(x.r.module)+' · '+Math.round(x.score*100)+'%</span></div><p>'+esc(x.r.question)+'</p></div>').join('')+'</div>';
addMessage('assistant',html)}
$('questionForm').addEventListener('submit',e=>{e.preventDefault();const q=$('questionInput').value.trim();if(!q)return;answer(q);$('questionInput').value='';$('questionInput').focus()});
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
$('menuButton').onclick=()=>document.body.classList.toggle('sidebar-open');
$('homeButton').onclick=()=>{$('chatArea').innerHTML='';$('welcome').hidden=false;$('questionInput').focus()};
(async()=>{db=await openDB();records=await allRecords();renderNav();syncSelectors();updateHeader();updateStats();if('serviceWorker'in navigator)navigator.serviceWorker.register('./sw.js').catch(()=>{});})();
})();