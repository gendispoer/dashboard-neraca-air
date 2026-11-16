/* ============================================================
   Dashboard Neraca Air - dashboard.js
   Logika data (WB), filter Year/Month, halaman Tabel/Grafik/Detail,
   impor .xlsx, penyimpanan, dan animasi welcome. Satu file untuk semua halaman.
   ============================================================ */
/*CORE-START*/
const PAGE=document.body.dataset.page||'home';   // home | tabel | grafik | detail
const G=9.81;
const MON=['Jan','Feb','Mar','Apr','Mei','Jun','Jul','Ags','Sep','Okt','Nov','Des'];
const MONF=['Januari','Februari','Maret','April','Mei','Juni','Juli','Agustus','September','Oktober','November','Desember'];
const VOLSRC=['Umbulan','Sumber Air','IP Ngagel I','IP Ngagel II','IP Ngagel III','IP KP I','IP KP II','IP KP III'];
const OPS=[['psb','Pemasangan Baru (PSB)','PTT/PTB'],['bukaan','Bukaan Kembali','Penertiban'],['tutupan','Tutupan Dinas','Penertiban'],
 ['aanboring','Aanboring','PTT/PTB'],['pindah','Pindah Letak Meter','PTT/PTB'],['rendah','Merendahkan Letak Meter','PTT/PTB'],
 ['tinggi','Meninggikan Letak Meter','PTT/PTB'],['gantiReg','Ganti Meter Reguler','PTT/PTB'],['stopKran','Ganti Stop Kran','PTT/PTB'],['kopling','Kopling Bocor','PTT/PTB']];
const GMB_SIZES=[1,1.5,2,3,4,6,8,16,18], CM_SIZES=[1,1.5,2,3,4,6,8];
const TB_KEYS=['tkGratis','pmk','asumsiPmk',...OPS.map(o=>o[0]),'gantiBesar','cuciMeter','cuciStrainer','flushing','tersier','penghubung'];
const LINE_KEYS=['input','rekening','suplisi','restitusi','tkKom','tkSos','sweeping',...TB_KEYS,'curi','airTerjual','salah','bocorSR','bocorTD','luapTangki','luapReservoir'];
const sum=a=>a.reduce((s,x)=>s+(+x||0),0);
const qv=(ptt,ptb,p,dm,t)=>0.6*((+ptt||0)+(+ptb||0))*Math.sqrt(2*G*(+p||0))*(0.75*Math.PI*dm*dm)*((+t||0)*60);
const rowVol=(r,inch)=>qv(r.ptt,r.ptb,r.p,(+r.d||0)*(inch?0.0254:1),r.t);

function calcLines(i){
  const L={};
  L.input=sum(i.volume);
  L.rekening=+i.rekening||0; L.suplisi=+i.suplisi||0; L.restitusi=-(+i.restitusi||0);
  L.tkKom=+i.tkKomersial||0; L.tkSos=+i.tkSosial||0; L.sweeping=+i.sweeping||0;
  L.tkGratis=+i.tkGratis||0; L.pmk=+i.pmk||0; L.asumsiPmk=+i.asumsiPmk||0;
  OPS.forEach(([k])=>{L[k]=rowVol(i.ops[k],false)});
  // Sama dengan rumus di workbook: Ganti Meter Besar = 7 baris pertama, Cuci Meter dan Cuci Strainer = 6 baris pertama.
  L.gantiBesar=sum(i.gmb.slice(0,7).map(r=>rowVol(r,true)));
  L.cuciMeter=sum(i.cm.slice(0,6).map(r=>rowVol(r,true)));
  L.cuciStrainer=sum(i.cs.slice(0,6).map(r=>rowVol(r,true)));
  L.flushing=sum([i.flush.distim,i.flush.disbar]); L.tersier=sum([i.tersier.distim,i.tersier.disbar]); L.penghubung=sum([i.penghubung.distim,i.penghubung.disbar]);
  L.curi=sum([i.curi.timur,i.curi.barat]);
  L.airTerjual=L.rekening+L.suplisi-L.restitusi;
  L.salah=(+i.salahPct||0)*L.airTerjual/100;
  L.bocorSR=sum([i.bocorSR.distim,i.bocorSR.disbar]); L.bocorTD=sum([i.bocorTD.distim,i.bocorTD.disbar]);
  L.luapTangki=+i.luapTangki||0; L.luapReservoir=+i.luapReservoir||0;
  return L;
}
function sumLines(arr){const o={};LINE_KEYS.forEach(k=>o[k]=0);arr.forEach(L=>LINE_KEYS.forEach(k=>o[k]+=L[k]||0));return o}
function summarize(L){
  const d=(a,b)=>b?a/b:0;
  const S={L};
  S.input=L.input;
  S.airBer=L.rekening+L.suplisi+L.restitusi+L.tkKom+L.tkSos+L.sweeping;
  S.nrw=S.input-S.airBer; S.nrwPct=d(S.nrw,S.input)*100; S.revPct=S.input?100-S.nrwPct:0;
  S.tb=sum(TB_KEYS.map(k=>L[k])); S.tbPct=d(S.tb,S.nrw)*100;
  S.resmi=S.airBer+S.tb;
  S.nf=L.curi+L.salah; S.nfPct=d(S.nf,S.nrw)*100;
  S.fs=L.bocorSR+L.bocorTD+L.luapTangki+L.luapReservoir; S.fsPct=d(S.fs,S.nrw)*100;
  S.keh=S.input-S.resmi;
  S.belum=S.keh-S.nf+S.fs; S.belumPct=d(S.belum,S.nrw)*100; // rumus sama dengan workbook: C43-D38+D45
  return S;
}
function parseWB(wb){
  const nm=wb.SheetNames.find(n=>n.trim().toUpperCase()==='WB');
  if(!nm) throw new Error('Sheet "WB" tidak ditemukan di file ini');
  const ws=wb.Sheets[nm];
  const v=a=>{const c=ws[a];if(!c)return 0;const n=typeof c.v==='number'?c.v:parseFloat(c.v);return isFinite(n)?n:0};
  const row=r=>({ptt:v('D'+r),ptb:v('E'+r),p:v('F'+r),d:v('G'+r),t:v('H'+r)});
  const pair=(r,a='D',b='E')=>({distim:v(a+r),disbar:v(b+r)});
  const inp={volume:VOLSRC.map((_,i)=>v('D'+(56+i))),
    rekening:v('D73'),suplisi:v('D74'),restitusi:v('D75'),
    tkKomersial:v('D80'),tkSosial:v('D81'),sweeping:v('D82'),
    tkGratis:v('D103'),pmk:v('D104'),asumsiPmk:v('D102'),ops:{},
    gmb:GMB_SIZES.map((_,i)=>row(109+i)),cm:CM_SIZES.map((_,i)=>row(119+i)),cs:CM_SIZES.map((_,i)=>row(127+i)),
    flush:pair(139),tersier:pair(140),penghubung:pair(141),
    curi:{timur:v('D147'),barat:v('E147')},salahPct:v('D152'),
    bocorSR:pair(162),bocorTD:pair(163),luapTangki:v('F164'),luapReservoir:v('F165')};
  OPS.forEach(([k],i)=>{inp.ops[k]=row(92+i)});
  return inp;
}
/*CORE-END*/

/* ================= helpers ================= */
const $=(s,r=document)=>r.querySelector(s), $$=(s,r=document)=>[...r.querySelectorAll(s)];
const esc=s=>String(s).replace(/[&<>"]/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c]));
const clone=o=>JSON.parse(JSON.stringify(o));
const key=(y,m)=>y+'-'+String(m).padStart(2,'0');
const pk=k=>({y:+k.slice(0,4),m:+k.slice(5)});
const lab=k=>{const{y,m}=pk(k);return MON[m-1]+' '+y};
const base=k=>{const{y,m}=pk(k);return 'WB_'+MON[m-1]+'_'+String(y).slice(2)};
const nf=(d)=>new Intl.NumberFormat('en-US',{minimumFractionDigits:d,maximumFractionDigits:d});
const NF0=nf(0),NF2=nf(2);
function fv(n,d=0){if(!isFinite(n)||n===0)return '-';const s=(d?NF2:NF0).format(Math.abs(n));return n<0?'('+s+')':s}
const fp=n=>NF2.format(isFinite(n)?n:0)+' %';
const fshort=n=>{const a=Math.abs(n);return a>=1e6?(n/1e6).toFixed(a>=1e7?1:2).replace(/\.?0+$/,'')+' jt':a>=1e3?Math.round(n/1e3)+' rb':String(Math.round(n))};
const getP=(o,p)=>p.split('.').reduce((a,k)=>a==null?a:a[k],o);
const setP=(o,p,v)=>{const ks=p.split('.');const l=ks.pop();ks.reduce((a,k)=>a[k],o)[l]=v};

let toastT;
function toast(msg,err){const t=$('#toast');t.textContent=msg;t.className='toast'+(err?' err':'');t.hidden=false;clearTimeout(toastT);toastT=setTimeout(()=>{t.hidden=true},err?7000:3500)}

/* ================= state ================= */
const S={years:new Set(),sel:new Set(),avail:new Set(),mode:'none',page:(PAGE==='home'?'tabel':PAGE),editKey:null,draft:null,dirty:false,booting:true,canWrite:true,driveMsg:''};
const cache={};  // key -> {inp,src}|null (null = file tidak ada)

/* Pilihan tahun/bulan dan file yang diimpor dibawa antar halaman lewat sessionStorage */
const SS='neracaAir.state';
function stLoad(){
  try{const o=JSON.parse(sessionStorage.getItem(SS)||'null');if(!o)return;
    (o.years||[]).forEach(y=>S.years.add(y));(o.sel||[]).forEach(k=>S.sel.add(k));
    }catch(e){}
}
function stSave(){
  try{sessionStorage.setItem(SS,JSON.stringify({years:[...S.years],sel:[...S.sel]}))}catch(e){}
}
addEventListener('pagehide',stSave);
let seq=0;

/* ================= database lokal (localStorage browser) =================
   Semua bulan disimpan di satu kunci: {months:{ "2026-08": <data input>, ... }}
   Tidak butuh server: cocok untuk GitHub Pages. Data hanya ada di browser ini. */
const DEMO_STORE='neracaAir.demo.v1';
const MV={Ags:['Ags','Agu','Aug'],Mei:['Mei','May'],Okt:['Okt','Oct'],Des:['Des','Dec']};
function bases(k){const{y,m}=pk(k);const yy=String(y).slice(2);return (MV[MON[m-1]]||[MON[m-1]]).map(n=>'WB_'+n+'_'+yy)}
let memStore={months:{}};   // cadangan bila localStorage diblokir
function readStore(){
  try{const o=JSON.parse(localStorage.getItem(DEMO_STORE)||'null');if(o&&o.months)return o}catch(e){}
  return memStore;
}
function writeStore(o){
  memStore=o;
  try{localStorage.setItem(DEMO_STORE,JSON.stringify(o));return true}catch(e){return false}
}

/* ================= data layer ================= */
async function loadMonth(k){
  if(cache[k])return cache[k];
  const inp=readStore().months[k];
  cache[k]=inp?{inp:clone(inp),src:'local'}:null;
  return cache[k];
}
async function refreshAvail(){
  S.avail=new Set(Object.keys(readStore().months));
  Object.keys(cache).forEach(k=>{if(cache[k])S.avail.add(k)});
}
async function persist(k,inp){
  const st=readStore();st.months[k]=clone(inp);st.savedAt=Date.now();
  if(!writeStore(st))throw new Error('Penyimpanan browser penuh atau diblokir');
  return 'browser ini';
}
async function removeMonth(k){
  const st=readStore();delete st.months[k];writeStore(st);delete cache[k];S.sel.delete(k);
}
async function connect(){
  S.mode='local';S.canWrite=true;
  await refreshAvail();
  if(!S.sel.size){
    const last=[...S.avail].sort().pop();
    if(last){const{y}=pk(last);S.years.add(y);S.sel.add(last)}
    else{const d=new Date();S.years.add(d.getFullYear());S.sel.add(key(d.getFullYear(),d.getMonth()+1))}
  }
  S.booting=false;setStatus();renderFilters();refresh();
}

/* ================= status / filter UI ================= */
function setStatus(){
  const t=$('#srcText'),b=$('#btnDrive');if(b)b.hidden=true;
  const n=S.avail.size;
  t.innerHTML='<span class="dot ok"></span>Mode demo · data tersimpan di browser ini'+(n?' ('+n+' bulan)':' · impor file WB_Bulan_Tahun.xlsx di halaman Detail');
}
function yearsList(){
  const now=new Date().getFullYear();let lo=Math.min(now-1,...[...S.avail].map(k=>pk(k).y)),hi=Math.max(now,...[...S.avail].map(k=>pk(k).y));
  const a=[];for(let y=lo;y<=hi;y++)a.push(y);return a;
}
function renderFilters(){
  const yc=$('#yearChips');
  yc.innerHTML=yearsList().map(y=>`<button type="button" class="chip" data-y="${y}" aria-pressed="${S.years.has(y)}">${y}</button>`).join('');
  const mp=$('#monthPanel');const ys=[...S.years].sort();
  mp.innerHTML=ys.length?ys.map(y=>`<h4><span>${y}</span><span><button type="button" data-all="${y}">Semua</button><button type="button" data-none="${y}">Kosongkan</button></span></h4><div class="chips">`+
    MON.map((n,i)=>{const k=key(y,i+1);return `<button type="button" class="chip" data-k="${k}" aria-pressed="${S.sel.has(k)}" title="${MONF[i]} ${y}">${n}${S.avail.has(k)?'<i title="File tersedia"></i>':''}</button>`}).join('')+`</div>`).join('')
    :'<p class="hint">Pilih Year dulu.</p>';
  const yl=[...S.years].sort();
  $('#lblYear').textContent=yl.length===0?'':yl.length===1?yl[0]:yl.length+' tahun';
  const ml=[...S.sel].sort();
  $('#lblMonth').textContent=ml.length===0?'':ml.length===1?lab(ml[0]):ml.length+' bulan';
}
function closeDD(except){$$('.dd').forEach(d=>{if(d!==except){d.classList.remove('open');$('.panel',d).hidden=true;$('button',d).setAttribute('aria-expanded','false')}})}
function toggleDD(d){const o=!d.classList.contains('open');closeDD(d);d.classList.toggle('open',o);$('.panel',d).hidden=!o;$('button',d).setAttribute('aria-expanded',String(o))}
if(PAGE!=='home'){
$('#btnYear').onclick=()=>toggleDD($('#ddYear'));
$('#btnMonth').onclick=()=>toggleDD($('#ddMonth'));
document.addEventListener('click',e=>{if(!e.target.closest('.dd'))closeDD()});
document.addEventListener('keydown',e=>{if(e.key==='Escape')closeDD()});
$('#yearChips').addEventListener('click',e=>{
  const b=e.target.closest('[data-y]');if(!b)return;const y=+b.dataset.y;
  if(S.years.has(y)){S.years.delete(y);[...S.sel].forEach(k=>{if(pk(k).y===y)S.sel.delete(k)})}else S.years.add(y);
  renderFilters();refresh();
});
$('#monthPanel').addEventListener('click',e=>{
  const b=e.target.closest('button');if(!b)return;
  if(b.dataset.k){S.sel.has(b.dataset.k)?S.sel.delete(b.dataset.k):S.sel.add(b.dataset.k)}
  else if(b.dataset.all){for(let m=1;m<=12;m++)S.sel.add(key(+b.dataset.all,m))}
  else if(b.dataset.none){for(let m=1;m<=12;m++)S.sel.delete(key(+b.dataset.none,m))}
  else return;
  renderFilters();refresh();
});
}
if($('#btnImport'))$('#btnImport').onclick=()=>$('#fileIn').click();
if($('#fileIn'))$('#fileIn').onchange=async e=>{
  const f=e.target.files[0];e.target.value='';if(!f)return;
  try{
    if(typeof XLSX==='undefined')throw new Error('Pustaka pembaca Excel belum termuat');
    const inp=parseWB(XLSX.read(new Uint8Array(await f.arrayBuffer()),{type:'array'}));
    const m=/WB_([A-Za-z]{3})_(\d{2})/i.exec(f.name);let k=null;
    if(m){const i=MON.findIndex(n=>(MV[n]||[n]).some(a=>a.toLowerCase()===m[1].toLowerCase()));if(i>=0)k=key(2000+ +m[2],i+1)}
    if(!k)throw new Error('Nama file harus berformat WB_Ags_26.xlsx agar bulannya terbaca');
    cache[k]={inp,src:'local'};S.sel.add(k);S.years.add(pk(k).y);
    let where=null;try{where=await persist(k,inp)}catch(er){toast('File terbaca, tetapi gagal disimpan: '+er.message,true)}
    if(where)toast(lab(k)+' diimpor dan disimpan ke '+where);else if(!$('#toast').classList.contains('err'))toast(lab(k)+' diimpor (belum tersimpan permanen)');
    await refreshAvail();renderFilters();refresh();
  }catch(er){toast(er.message,true)}
};

/* ================= routing (tiap halaman = file HTML sendiri di folder pages/) ================= */
const PAGES=['tabel','grafik','detail'];
function route(){
  if(!PAGES.includes(PAGE))return;
  S.page=PAGE;
  $$('.dk').forEach(t=>t.setAttribute('aria-current',t.dataset.page===PAGE?'page':'false'));
  const sec=$('#pg-'+PAGE);if(sec)sec.hidden=false;
  renderPage();
}
if(PAGE!=='home'){
  $$('.dk').forEach(t=>t.onclick=()=>{
    if(matchMedia('(hover:none)').matches)$('#dockZone').classList.remove('open');
    if(t.dataset.page!==PAGE)location.href=t.dataset.page+'.html';
  });
  $('#dockZone').addEventListener('pointerdown',e=>{if(e.pointerType!=='mouse')$('#dockZone').classList.add('open')});
  document.addEventListener('pointerdown',e=>{if(!e.target.closest('#dockZone'))$('#dockZone').classList.remove('open')});
}

/* ================= refresh + render ================= */
async function refresh(){
  stSave();
  const my=++seq;
  const ks=[...S.sel].sort();
  await Promise.all(ks.map(loadMonth));
  if(my!==seq)return;
  renderPage();
}
const present=()=>[...S.sel].sort().filter(k=>cache[k]);
const missing=()=>[...S.sel].sort().filter(k=>cache[k]===null);
function period(){
  const p=present();if(!p.length)return '';
  return p.length===1?MONF[pk(p[0]).m-1]+' '+pk(p[0]).y:p.map(lab).join(', ')+' · akumulasi '+p.length+' bulan';
}
function renderNotice(){
  const m=missing();
  $('#notice').innerHTML=m.length?`<div class="note"><strong>File not found.</strong> ${m.map(k=>esc(base(k))+' ('+lab(k)+')').join(', ')} belum ada di penyimpanan browser. Bulan ini tidak ikut dihitung.</div>`:'';
}
function renderPage(){
  if(!S.years||$('#app').hidden)return;
  renderNotice();
  if(S.page==='tabel')renderTabel();
  else if(S.page==='grafik')renderGrafik();
  else renderDetail();
}
function emptyState(){
  if(S.booting)return '<div class="loading">Memuat data…</div>';
  if(!S.sel.size)return '<div class="empty"><h2>Pilih bulan</h2><p>Pilih Year lalu Month di kanan atas.</p></div>';
  const m=missing();
  return `<div class="empty"><h2>File not found</h2><p>${m.map(k=>esc(base(k))).join(', ')||'File'} tidak ditemukan. Pilih bulan lain atau impor file .xlsx.</p></div>`;
}
function currentSummary(){return summarize(sumLines(present().map(k=>calcLines(cache[k].inp))))}

/* ----- Tabel ----- */
function rows(items,dec=0){return items.map(([l,k,ind])=>k?`<div class="rw${ind?' ind':''}"><span>${l}</span><span class="v" data-v="${k}" data-d="${dec}"></span></div>`:`<div class="rw lbl"><span>${l}</span><span></span></div>`).join('')}
function sheet(title,items,dec){return `<div class="blk"><div class="th">${title}</div><div class="sub"><span>Sumber Data</span><span>Volume (m³)</span></div>${rows(items,dec)}</div>`}
let tabelBuilt=false;
function buildTabel(){
  const opsRows=[['Tangki Air Gratis','tkGratis'],['PMK','pmk'],['Asumsi PMK Tidak Terlapor','asumsiPmk'],['Kegiatan Operasional:',null],
    ...OPS.map(o=>['- '+o[1],o[0],1]),['- Ganti Meter Besar','gantiBesar',1],['- Cuci Meter','cuciMeter',1],['- Cuci Strainer','cuciStrainer',1],
    ['- Flushing/Cuci Pipa','flushing',1],['- Rehab Pipa Tersier','tersier',1],['- Rehab Pipa Penghubung','penghubung',1]];
  const g=(cls,t,n,u,p)=>`<div class="g ${cls}"><b>${t}</b><span class="n" data-v="${n}"></span><span class="u">m³</span>${p?`<span class="p" data-p="${p}"></span>`:''}</div>`;
  $('#tabelBody').innerHTML=`<div class="nrw" id="nrwGrid">
    ${g('a-inp','Volume Input Sistem','input')}
    ${g('a-res','Konsumsi Resmi','resmi')}
    ${g('a-ber','Konsumsi Resmi Berekening','airBer')}
    <div class="a-ta">${sheet('Konsumsi Bermeter Berekening',[['Rekening Pemakaian Air Pelanggan','rekening'],['Suplisi','suplisi'],['Restitusi','restitusi']])}${sheet('Konsumsi Tak Bermeter Berekening',[['Tangki Air Komersial','tkKom'],['Tangki Air Sosial','tkSos'],['Sweeping','sweeping']])}</div>
    ${g('a-tak','Konsumsi Resmi Tak Berekening','tb','', 'tbPct')}
    <div class="a-tb">${sheet('Konsumsi Tak Bermeter Tak Berekening',opsRows)}</div>
    ${g('a-keh','Kehilangan Air','keh')}
    ${g('a-nf','Kehilangan Air Non Fisik','nf','','nfPct')}
    <div class="a-tc">${sheet('Konsumsi Tak Resmi',[['Pencurian Air','curi']],2)}${sheet('Ketidakakuratan Meter &amp; Penanganan Data',[['Estimasi Kesalahan Bacaan Meter','salah']],2)}</div>
    ${g('a-fs','Kehilangan Air Fisik','fs','','fsPct')}
    <div class="a-td">${sheet('Kebocoran Fisik',[['Kebocoran Pipa Penghubung SR','bocorSR'],['Kebocoran Pipa Transmisi &amp; Distribusi','bocorTD'],['Luapan Pengisian Tangki Air','luapTangki'],['Luapan Reservoir','luapReservoir']],2)}</div>
    ${g('a-bd','Kehilangan Air Belum Terdeteksi','belum','','belumPct')}
    ${g('a-rA','Air Berekening','airBer','','revPct')}
    ${g('a-rB','NRW','nrw','','nrwPct')}
  </div>`;
}
function renderTabel(){
  $('#periodTabel').textContent=period();
  const body=$('#tabelBody');
  if(!present().length){body.innerHTML=emptyState();tabelBuilt=false;return}
  if(!tabelBuilt||!$('#nrwGrid')){buildTabel();tabelBuilt=true}
  const s=currentSummary(),L=s.L;
  const val={...L,airBer:s.airBer,resmi:s.resmi,tb:s.tb,keh:s.keh,nf:s.nf,fs:s.fs,belum:s.belum,nrw:s.nrw};
  $$('[data-v]',body).forEach(el=>{el.textContent=fv(val[el.dataset.v],+el.dataset.d||0)});
  $$('[data-p]',body).forEach(el=>{el.textContent=fp(s[el.dataset.p])});
  // baris tak berekening yang aslinya tampil merah di workbook diberi penekanan lembut saat bernilai
}

/* ----- Grafik ----- */
const charts={};
const T={};
const css=v=>getComputedStyle(document.documentElement).getPropertyValue(v).trim();
function mk(id,cfg){if(charts[id]){charts[id].destroy();delete charts[id]}const c=$('#'+id);if(c)charts[id]=new Chart(c,cfg)}
const valLabels={id:'valLabels',afterDatasetsDraw(ch,_a,o){
  if(!o||!o.on)return;const ctx=ch.ctx;ctx.save();ctx.fillStyle=T.text;ctx.font='600 11px "Plus Jakarta Sans", sans-serif';ctx.textBaseline='middle';
  ch.data.datasets.forEach((ds,di)=>{const m=ch.getDatasetMeta(di);if(m.hidden)return;
    m.data.forEach((bar,i)=>{const v=ds.data[i];if(v==null||!isFinite(v))return;ctx.textAlign='left';ctx.fillText(o.fmt(v),bar.x+6,bar.y)})});
  ctx.restore()}};
function hbar(id,labels,datasets,legend){
  mk(id,{type:'bar',data:{labels,datasets},plugins:[valLabels],options:{indexAxis:'y',responsive:true,maintainAspectRatio:false,animation:{duration:350},
    plugins:{legend:{display:legend,position:'bottom',labels:{color:T.text,boxWidth:12,boxHeight:12}},
      tooltip:{callbacks:{label:c=>' '+c.dataset.label+': '+NF0.format(c.parsed.x)+' m³'}},valLabels:{on:true,fmt:v=>fshort(v)}},
    scales:{x:{stacked:legend,beginAtZero:true,grace:'22%',ticks:{color:T.muted,callback:v=>fshort(v),maxRotation:0,autoSkip:true,maxTicksLimit:6},grid:{color:T.line}},
            y:{stacked:legend,ticks:{color:T.text,autoSkip:false},grid:{display:false}}},
    datasets:{bar:{borderRadius:4,barPercentage:.72,categoryPercentage:.9}}}});
}
function renderGrafik(){
  $('#periodGrafik').textContent=period();
  const body=$('#grafikBody');
  if(!present().length){Object.keys(charts).forEach(k=>{charts[k].destroy();delete charts[k]});body.innerHTML=emptyState();return}
  if(typeof Chart==='undefined'){body.innerHTML='<div class="empty"><h2>Grafik tidak termuat</h2><p>Pustaka grafik gagal dimuat. Muat ulang halaman.</p></div>';return}
  Object.keys(charts).forEach(k=>{charts[k].destroy();delete charts[k]});
  Object.assign(T,{text:css('--text'),muted:css('--muted'),line:css('--line'),c1:css('--c1'),c2:css('--c2'),c3:css('--c3'),surface:css('--surface')});
  Chart.defaults.font.family='"Plus Jakarta Sans", system-ui, sans-serif';Chart.defaults.color=T.muted;
  const ks=present(),s=currentSummary(),L=s.L;
  const per=ks.map(k=>({key:k,label:lab(k),inp:cache[k].inp,L:calcLines(cache[k].inp),...summarize(calcLines(cache[k].inp))}));
  const multi=ks.length>1;
  const total=(fn)=>ks.reduce((a,k)=>a+fn(cache[k].inp,calcLines(cache[k].inp)),0);
  const chartCard=(id,title,desc,wide=false,tall=false)=>`<section class="cc ${wide?'wide':''}"><div class="saas-card-head"><div><span class="eyebrow">ANALISIS DATASET</span><h3>${title}</h3><p>${desc}</p></div></div><div class="cv ${tall?'tall':''}"><canvas id="${id}" role="img" aria-label="${title}"></canvas></div></section>`;
  body.innerHTML=`
  <div class="saas-kpis">
    <article class="kpi k1"><div class="kpi-top"><span>Volume input sistem</span><span class="kpi-icon">↗</span></div><strong>${NF0.format(s.input)}</strong><small>m³ · total periode terpilih</small></article>
    <article class="kpi k2"><div class="kpi-top"><span>Air berekening</span><span class="kpi-icon">◌</span></div><strong>${NF0.format(s.airBer)}</strong><small>m³ · ${fp(s.revPct)} dari input</small></article>
    <article class="kpi k3"><div class="kpi-top"><span>Non-Revenue Water</span><span class="kpi-icon">⌁</span></div><strong>${NF0.format(s.nrw)}</strong><small>m³ · ${fp(s.nrwPct)} dari input</small></article>
    <article class="kpi k4"><div class="kpi-top"><span>Kehilangan belum terdeteksi</span><span class="kpi-icon">◎</span></div><strong>${NF0.format(s.belum)}</strong><small>m³ · ${fp(s.belumPct)} dari NRW</small></article>
  </div>
  <div class="saas-main-grid">
    <section class="cc trend-panel"><div class="saas-card-head"><div><span class="eyebrow">PERFORMA BULANAN</span><h3>Tren Neraca Air</h3><p>Perubahan air berekening dan NRW pada bulan terpilih</p></div><span class="period-pill">${ks.length} bulan</span></div>
      <div class="trend-legend"><span><i class="legend-dot blue"></i>Air berekening</span><span><i class="legend-dot red"></i>NRW</span></div>
      <div class="cv trend-cv"><canvas id="trendMain" role="img" aria-label="Tren neraca air bulanan"></canvas></div>
      ${multi?'':'<div class="chart-hint">Pilih minimal dua bulan untuk melihat fluktuasi antarperiode.</div>'}</section>
    <aside class="cc composition-panel"><div class="saas-card-head"><div><span class="eyebrow">KOMPOSISI</span><h3>Air Berekening vs NRW</h3><p>Proporsi terhadap volume input sistem</p></div></div>
      <div class="cv donut-cv"><canvas id="ch1" role="img" aria-label="Komposisi neraca air"></canvas></div>
      <div class="composition-legend"><div><span><i class="legend-dot blue"></i>Air berekening</span><b>${fp(s.revPct)}</b></div><div><span><i class="legend-dot red"></i>NRW</span><b>${fp(s.nrwPct)}</b></div></div>
      <div class="mini-total"><span>Total volume input</span><strong>${NF0.format(s.input)} <small>m³</small></strong></div></aside>
  </div>

  <div class="chart-section"><div><h2>01 · Sumber Air</h2><p>Kontribusi setiap sumber terhadap Volume Input Sistem</p></div></div>
  <div class="chart-grid">${chartCard('srcRank','Volume Input per Sumber','Peringkat delapan sumber air (m³)')}${chartCard('srcShare','Proporsi Sumber Air','Persentase kontribusi sumber terhadap total input')}</div>

  <div class="chart-section"><div><h2>02 · Konsumsi Resmi</h2><p>Komposisi pemanfaatan air yang tercatat dalam neraca</p></div></div>
  <div class="chart-grid">${chartCard('consumptionMix','Komposisi Konsumsi Resmi','Bermeter berekening, tak bermeter berekening, dan tak bermeter tak berekening')}${chartCard('consumptionItems','Rincian Konsumsi','Komponen konsumsi resmi berdasarkan volume (m³)',false,true)}</div>

  <div class="chart-section"><div><h2>03 · Kehilangan Air</h2><p>Rincian kehilangan nonfisik, fisik, dan perbandingan wilayah</p></div></div>
  <div class="chart-grid">${chartCard('lossRank','Komponen Kehilangan Air','Pencurian, kesalahan bacaan meter, kebocoran, dan luapan',false,true)}${chartCard('lossArea','Perbandingan Wilayah Timur dan Barat','Komponen yang tersedia dengan pemisahan wilayah (m³)',false,true)}</div>

  <div class="chart-section"><div><h2>04 · Kegiatan Operasional</h2><p>Aktivitas operasional dan volume hasil estimasi perhitungan</p></div></div>
  <div class="chart-grid">${chartCard('opsVolume','Volume per Kegiatan Operasional','Volume hasil perhitungan setiap jenis kegiatan (m³)',false,true)}${chartCard('opsCount','Jumlah PTT dan PTB','Jumlah pekerjaan berdasarkan jenis kegiatan')}${chartCard('meterSize','Volume Berdasarkan Diameter Meter','Ganti meter besar, cuci meter, dan cuci strainer (m³)',true,true)}</div>
  <p class="chart-note">Catatan interpretasi: nilai volume kegiatan operasional merupakan estimasi berdasarkan parameter PTT, PTB, tekanan, diameter, dan waktu pada workbook. Nilai kehilangan belum terdeteksi mengikuti rumus neraca pada file sumber. Komponen yang tidak memiliki pemisahan Timur/Barat tidak ditampilkan pada grafik wilayah.</p>`;

  const axisCommon={ticks:{color:T.muted,font:{size:10},maxRotation:0},grid:{color:T.line,drawBorder:false}};
  mk('trendMain',{type:'line',data:{labels:per.map(p=>p.label),datasets:[
    {label:'Air berekening',data:per.map(p=>p.airBer),borderColor:T.c1,backgroundColor:'rgba(206,231,243,.45)',borderWidth:3,pointRadius:multi?4:6,pointHoverRadius:7,pointBackgroundColor:T.c1,pointBorderColor:'#FFFFFF',pointBorderWidth:2,tension:.42,fill:true},
    {label:'NRW',data:per.map(p=>p.nrw),borderColor:T.c2,backgroundColor:'rgba(169,27,24,.08)',borderWidth:3,pointRadius:multi?4:6,pointHoverRadius:7,pointBackgroundColor:T.c2,pointBorderColor:'#FFFFFF',pointBorderWidth:2,tension:.42,fill:false}
  ]},options:{responsive:true,maintainAspectRatio:false,interaction:{mode:'index',intersect:false},plugins:{legend:{display:false},tooltip:{backgroundColor:'#181717',titleColor:'#F8FAED',bodyColor:'#F8FAED',padding:11,callbacks:{label:c=>' '+c.dataset.label+': '+NF0.format(c.parsed.y)+' m³'}}},scales:{x:{...axisCommon,grid:{display:false}},y:{...axisCommon,beginAtZero:true,ticks:{...axisCommon.ticks,callback:v=>fshort(v)}}}}});
  mk('ch1',{type:'doughnut',data:{labels:['Air berekening','NRW'],datasets:[{data:[Math.max(0,s.airBer),Math.max(0,s.nrw)],backgroundColor:[T.c1,T.c2],borderColor:'#FFFFFF',borderWidth:4,hoverOffset:5}]},options:{responsive:true,maintainAspectRatio:false,cutout:'72%',plugins:{legend:{display:false},tooltip:{callbacks:{label:c=>' '+c.label+': '+NF0.format(c.parsed)+' m³'}}}}});

  const srcVals=VOLSRC.map((n,i)=>ks.reduce((a,k)=>a+(+cache[k].inp.volume[i]||0),0));
  hbar('srcRank',VOLSRC.map((n,i)=>n),[{label:'Volume',data:srcVals,backgroundColor:T.c1}],false);
  mk('srcShare',{type:'doughnut',data:{labels:VOLSRC,datasets:[{data:srcVals,backgroundColor:['#2A82B8','#A91B18','#7DBBD5','#C65C55','#181717','#B7791F','#8CA7B5','#D98B83'],borderColor:'#fff',borderWidth:3}]},options:{responsive:true,maintainAspectRatio:false,cutout:'58%',plugins:{legend:{position:'bottom',labels:{color:T.text,boxWidth:10,font:{size:9}}},tooltip:{callbacks:{label:c=>` ${c.label}: ${NF0.format(c.parsed)} m³ (${fp(s.input?c.parsed/s.input*100:0)})`}}}}});

  const cons=[s.airBer-s.tb,s.tb,s.tb];
  // Konsumsi resmi: Air Berekening dan Tak Bermeter Tak Berekening; komponen bermeter/tak bermeter berekening dirinci terpisah.
  const meter=total((i,l)=>l.rekening+l.suplisi+l.restitusi+l.tkKom+l.tkSos+l.sweeping);
  const unmeterBill=total((i,l)=>l.tkKom+l.tkSos+l.sweeping);
  const unmeterNoBill=total((i,l)=>sum(TB_KEYS.map(k=>l[k]||0)));
  const meteredBill=total((i,l)=>l.rekening+l.suplisi+l.restitusi);
  mk('consumptionMix',{type:'doughnut',data:{labels:['Konsumsi bermeter berekening','Konsumsi tak bermeter berekening','Konsumsi tak bermeter tak berekening'],datasets:[{data:[Math.max(0,meteredBill),Math.max(0,unmeterBill),Math.max(0,unmeterNoBill)],backgroundColor:[T.c1,'#B7DCEB',T.c2],borderColor:'#fff',borderWidth:3}]},options:{responsive:true,maintainAspectRatio:false,cutout:'60%',plugins:{legend:{position:'bottom',labels:{color:T.text,boxWidth:10,font:{size:9}}},tooltip:{callbacks:{label:c=>` ${c.label}: ${NF0.format(c.parsed)} m³`}}}}});
  const cItems=[['Rekening pelanggan',total((i,l)=>l.rekening)],['Suplisi',total((i,l)=>l.suplisi)],['Restitusi',total((i,l)=>l.restitusi)],['Tangki komersial',total((i,l)=>l.tkKom)],['Tangki sosial',total((i,l)=>l.tkSos)],['Sweeping',total((i,l)=>l.sweeping)],['Konsumsi tak bermeter tak berekening',unmeterNoBill]];
  hbar('consumptionItems',cItems.map(x=>x[0]),[{label:'Volume',data:cItems.map(x=>x[1]),backgroundColor:T.c1}],false);

  const losses=[['Pencurian air',L.curi],['Kesalahan bacaan meter',L.salah],['Kebocoran pipa penghubung SR',L.bocorSR],['Kebocoran transmisi & distribusi',L.bocorTD],['Luapan tangki',L.luapTangki],['Luapan reservoir',L.luapReservoir],['Kehilangan belum terdeteksi',s.belum]];
  hbar('lossRank',losses.map(x=>x[0]),[{label:'Volume',data:losses.map(x=>x[1]),backgroundColor:['#A91B18','#C65C55','#2A82B8','#7DBBD5','#B7791F','#181717','#8D8580']}],false);
  const areaDefs=[['Pencurian air',i=>i.curi.timur,i=>i.curi.barat],['Kebocoran SR',i=>i.bocorSR.distim,i=>i.bocorSR.disbar],['Kebocoran transmisi & distribusi',i=>i.bocorTD.distim,i=>i.bocorTD.disbar],['Flushing/Cuci Pipa',i=>i.flush.distim,i=>i.flush.disbar],['Rehab pipa tersier',i=>i.tersier.distim,i=>i.tersier.disbar],['Rehab pipa penghubung',i=>i.penghubung.distim,i=>i.penghubung.disbar]];
  mk('lossArea',{type:'bar',data:{labels:areaDefs.map(x=>x[0]),datasets:[{label:'Timur',data:areaDefs.map(x=>total(i=>x[1](i))),backgroundColor:T.c1},{label:'Barat',data:areaDefs.map(x=>total(i=>x[2](i))),backgroundColor:T.c2}]},options:{responsive:true,maintainAspectRatio:false,indexAxis:'y',plugins:{legend:{position:'bottom',labels:{color:T.text}},tooltip:{callbacks:{label:c=>` ${c.dataset.label}: ${NF0.format(c.parsed.x)} m³`}}},scales:{x:{beginAtZero:true,ticks:{color:T.muted,callback:v=>fshort(v)},grid:{color:T.line}},y:{ticks:{color:T.text},grid:{display:false}}}}});

  const opRows=OPS.map(([k,label])=>[label,total((i,l)=>l[k]||0)]);
  opRows.push(['Ganti Meter Besar',total((i,l)=>l.gantiBesar)],['Cuci Meter',total((i,l)=>l.cuciMeter)],['Cuci Strainer',total((i,l)=>l.cuciStrainer)],['Flushing/Cuci Pipa',total((i,l)=>l.flushing)],['Rehab Pipa Tersier',total((i,l)=>l.tersier)],['Rehab Pipa Penghubung',total((i,l)=>l.penghubung)]);
  hbar('opsVolume',opRows.map(x=>x[0]),[{label:'Volume',data:opRows.map(x=>x[1]),backgroundColor:T.c1}],false);
  const opsCounts=OPS.map(([k,label])=>({label,ptt:total(i=>i.ops[k].ptt),ptb:total(i=>i.ops[k].ptb)}));
  mk('opsCount',{type:'bar',data:{labels:opsCounts.map(x=>x.label),datasets:[{label:'PTT',data:opsCounts.map(x=>x.ptt),backgroundColor:T.c1},{label:'PTB',data:opsCounts.map(x=>x.ptb),backgroundColor:T.c2}]},options:{responsive:true,maintainAspectRatio:false,indexAxis:'y',plugins:{legend:{position:'bottom',labels:{color:T.text}},tooltip:{callbacks:{label:c=>` ${c.dataset.label}: ${NF0.format(c.parsed.x)}`}}},scales:{x:{beginAtZero:true,ticks:{color:T.muted},grid:{color:T.line}},y:{ticks:{color:T.text},grid:{display:false}}}}});
  const sizeCats=[...GMB_SIZES.map((z,idx)=>({label:`Ganti ${z}\"`,get:inp=>rowVol(inp.gmb[idx],true)})),...CM_SIZES.map((z,idx)=>({label:`Cuci meter ${z}\"`,get:inp=>rowVol(inp.cm[idx],true)})),...CM_SIZES.map((z,idx)=>({label:`Cuci strainer ${z}\"`,get:inp=>rowVol(inp.cs[idx],true)}))];
  hbar('meterSize',sizeCats.map(x=>x.label),[{label:'Volume',data:sizeCats.map(x=>total(i=>x.get(i))),backgroundColor:sizeCats.map((x,i)=>i< GMB_SIZES.length?T.c2:i< GMB_SIZES.length+CM_SIZES.length?T.c1:'#B7DCEB')}],false);
}
matchMedia('(prefers-color-scheme: dark)').addEventListener&&matchMedia('(prefers-color-scheme: dark)').addEventListener('change',()=>{if(S.page==='grafik')renderGrafik()});
new MutationObserver(()=>{if(S.page==='grafik')renderGrafik()}).observe(document.documentElement,{attributes:true,attributeFilter:['data-theme']});

/* ----- Detail (edit) ----- */
const inI=(p,l,w)=>`<input class="in" type="number" step="any" inputmode="decimal" data-p="${p}" aria-label="${esc(l)}" value="${getP(S.draft,p)}">`;
const picTag=t=>t?`<span class="pic">${t}</span>`:'';
function simple(title,items){
  return `<section class="card"><h3>${title}</h3><div class="scroll"><table class="dt"><thead><tr><th>Uraian</th><th>PIC</th><th class="r">Jumlah</th></tr></thead><tbody>`+
  items.map(([p,l,pic,note])=>`<tr><th>${l}${note?` <small style="color:var(--muted);font-weight:400">${note}</small>`:''}</th><td>${picTag(pic)}</td><td>${inI(p,l)}</td></tr>`).join('')+`</tbody></table></div></section>`;
}
function flowTable(title,subtitle,rowsDef,inch,sumKey,limit){
  return `<h4>${subtitle}</h4><div class="scroll"><table class="dt"><thead><tr><th>${inch?'Ukuran':'Kegiatan'}</th><th>PIC</th><th class="r">PTT</th><th class="r">PTB</th><th class="r">Tekanan (m)</th><th class="r">Diameter (${inch?'inci':'m'})</th><th class="r">Waktu (menit)</th><th class="r">Volume (m³)</th></tr></thead><tbody>`+
  rowsDef.map(([p,l,pic],idx)=>{const skip=limit!=null&&idx>=limit;return `<tr class="${skip?'skip':''}"><th>${l}${skip?' <small>(tidak ikut total)</small>':''}</th><td>${picTag(pic)}</td>`+
    ['ptt','ptb','p','d','t'].map(f=>`<td>${inI(p+'.'+f,l+' '+f)}</td>`).join('')+`<td class="cv-out" data-rv="${p}|${inch?1:0}"></td></tr>`}).join('')+
  (sumKey?`<tr class="tot"><th colspan="7">Total ${sumKey[1]}</th><td class="cv-out" data-L="${sumKey[0]}"></td></tr>`:'')+`</tbody></table></div>`;
}
function pairTable(title,defs){
  return `<section class="card"><h3>${title}</h3><div class="scroll"><table class="dt"><thead><tr><th>Uraian</th><th>PIC</th><th class="r">${defs.cols[0]}</th><th class="r">${defs.cols[1]}</th><th class="r">Total</th></tr></thead><tbody>`+
  defs.rows.map(([p,l,pic,a,b,Lk])=>`<tr><th>${l}</th><td>${picTag(pic)}</td><td>${inI(p+'.'+a,l+' '+defs.cols[0])}</td><td>${inI(p+'.'+b,l+' '+defs.cols[1])}</td><td class="cv-out" data-L="${Lk}"></td></tr>`).join('')+`</tbody></table></div></section>`;
}
function buildDetail(){
  const html=[];
  html.push(`<section class="card"><h3>Volume Input</h3><div class="scroll"><table class="dt"><thead><tr><th>Uraian</th><th>Keterangan</th><th class="r">Jumlah (m³)</th></tr></thead><tbody>`+
    VOLSRC.map((n,i)=>`<tr><th>${n}</th><td><span class="pic">Laporan Bagian Produksi</span></td><td>${inI('volume.'+i,n)}</td></tr>`).join('')+
    `<tr class="tot"><th colspan="2">Volume Input Sistem</th><td class="cv-out" data-L="input"></td></tr></tbody></table></div></section>`);
  html.push(simple('Konsumsi Bermeter Berekening',[['rekening','Rekening Pemakaian Air Pelanggan','Pemakaian Air'],['suplisi','Suplisi','Pemakaian Air'],['restitusi','Restitusi','Pemakaian Air','dikurangkan dari air berekening']]));
  html.push(simple('Konsumsi Tak Bermeter Berekening',[['tkKomersial','Tangki Air Komersial','Hublang'],['tkSosial','Tangki Air Sosial','Hublang'],['sweeping','Sweeping','Penertiban']]));
  html.push(simple('Konsumsi Tak Bermeter Tak Berekening',[['tkGratis','Tangki Air Gratis','Hublang'],['pmk','PMK','PKA'],['asumsiPmk','Asumsi PMK Tidak Terlapor','PKA']]));
  html.push(`<section class="card"><h3>Kegiatan Operasional</h3>`+
    flowTable('','Pemeliharaan distribusi',OPS.map(o=>['ops.'+o[0],o[1],o[2]]),false,null)+
    flowTable('','Ganti Meter Besar',GMB_SIZES.map((z,i)=>['gmb.'+i,z+'"','PTT/PTB']),true,['gantiBesar','Ganti Meter Besar'],7)+
    flowTable('','Cuci Meter',CM_SIZES.map((z,i)=>['cm.'+i,z+'"','PTT/PTB']),true,['cuciMeter','Cuci Meter'],6)+
    flowTable('','Cuci Strainer',CM_SIZES.map((z,i)=>['cs.'+i,z+'"','PTT/PTB']),true,['cuciStrainer','Cuci Strainer'],6)+`</section>`);
  html.push(pairTable('Flushing dan Rehab Pipa',{cols:['Distim','Disbar'],rows:[['flush','Flushing/Cuci Pipa','Distim/Disbar','distim','disbar','flushing'],['tersier','Rehab Pipa Tersier','Distim/Disbar','distim','disbar','tersier'],['penghubung','Rehab Pipa Penghubung','Distim/Disbar','distim','disbar','penghubung']]}));
  html.push(pairTable('Kehilangan Air Non Fisik: Pencurian Air',{cols:['Timur','Barat'],rows:[['curi','Pencurian Air','Penertiban','timur','barat','curi']]}));
  html.push(`<section class="card"><h3>Kehilangan Air Non Fisik: Ketidakakuratan Meter</h3><div class="scroll"><table class="dt"><thead><tr><th>Uraian</th><th>PIC</th><th class="r">Kesalahan (%)</th><th class="r">Air terjual (m³)</th><th class="r">Estimasi (m³)</th></tr></thead><tbody><tr><th>Estimasi Kesalahan Bacaan Meter</th><td>${picTag('Hublang')}</td><td>${inI('salahPct','Persen kesalahan bacaan')}</td><td class="cv-out" data-L="airTerjual"></td><td class="cv-out" data-L="salah"></td></tr></tbody></table></div></section>`);
  html.push(pairTable('Kehilangan Air Fisik: Kebocoran',{cols:['Distim','Disbar'],rows:[['bocorSR','Kebocoran Pipa Penghubung SR','Distim/Disbar','distim','disbar','bocorSR'],['bocorTD','Kebocoran Pipa Transmisi &amp; Distribusi','Distim/Disbar','distim','disbar','bocorTD']]}));
  html.push(simple('Kehilangan Air Fisik: Luapan',[['luapTangki','Luapan Pengisian Tangki Air','-'],['luapReservoir','Luapan Reservoir','-']]));
  return html.join('');
}
function updateComputed(){
  const L=calcLines(S.draft),s=summarize(L);
  $$('#detailBody [data-L]').forEach(el=>{el.textContent=NF2.format(L[el.dataset.L]||0)});
  $$('#detailBody [data-rv]').forEach(el=>{const[p,inch]=el.dataset.rv.split('|');el.textContent=NF2.format(rowVol(getP(S.draft,p),inch==='1'))});
  const st=$('#dSum');if(st)st.innerHTML=`<div><span>Volume input</span><b>${NF0.format(s.input)}</b></div><div><span>Air berekening</span><b>${NF0.format(s.airBer)}</b></div><div><span>NRW (m³)</span><b>${NF0.format(s.nrw)}</b></div><div><span>NRW (%)</span><b>${fp(s.nrwPct)}</b></div>`;
}
function renderDetail(){
  const body=$('#detailBody');
  const keys=[...new Set([...present(),...(S.dirty&&S.editKey?[S.editKey]:[])])].sort();
  if(!keys.length){body.innerHTML=emptyState();S.draft=null;S.editKey=null;S.dirty=false;return}
  if(!S.editKey||!keys.includes(S.editKey)){S.editKey=keys[keys.length-1];S.draft=null;S.dirty=false}
  if(!S.draft)S.draft=clone(cache[S.editKey].inp);
  if(body.dataset.built===S.editKey+'|'+keys.join(',')&&body.firstChild){updateComputed();return}
  const lock=false;
  body.innerHTML=`<div class="dtools"><div class="dd left" id="ddEdit"><button type="button" id="btnEdit" aria-haspopup="true" aria-expanded="false"${S.dirty?' disabled':''}>Edit bulan <small>${lab(S.editKey)}</small></button>
        <div class="panel" role="group" aria-label="Pilih bulan yang diedit" hidden><h4>Bulan</h4><div class="chips c3">${keys.map(k=>`<button type="button" class="chip" data-ek="${k}" aria-pressed="${k===S.editKey}">${lab(k)}</button>`).join('')}</div></div></div>
      <span class="hint">${S.dirty?'Simpan atau batalkan perubahan sebelum pindah bulan.':keys.length>1?'Pilih bulan yang mau diedit.':''}</span></div>
    <div class="sumstrip" id="dSum"></div>${buildDetail()}
    <div class="savebar"><span id="saveMsg">${lock?'Akun ini hanya bisa melihat data.':'Belum ada perubahan.'}</span><div class="acts"><button class="btn" type="button" id="btnReset" disabled>Batalkan</button><button class="btn primary" type="button" id="btnSave" disabled>Save</button></div></div>`;
  body.dataset.built=S.editKey+'|'+keys.join(',');
  updateComputed();markDirty();
}
function markDirty(){
  const d=S.dirty;const sv=$('#btnSave'),rs=$('#btnReset'),m=$('#saveMsg'),sel=$('#btnEdit');
  if(!sv)return;sv.disabled=!d;rs.disabled=!d;if(sel)sel.disabled=d;
  if(d){m.textContent='Ada perubahan yang belum disimpan.';m.className='dirty'}else{m.className='';if(!/Tersimpan|hanya/.test(m.textContent))m.textContent='Belum ada perubahan.'}
}
$('#detailBody')?.addEventListener('input',e=>{
  const el=e.target.closest('.in');if(!el||!S.draft)return;
  const v=el.value===''?0:parseFloat(el.value);setP(S.draft,el.dataset.p,isFinite(v)?v:0);
  const orig=getP(cache[S.editKey].inp,el.dataset.p);el.classList.toggle('chg',(+orig||0)!==(isFinite(v)?v:0));
  S.dirty=JSON.stringify(S.draft)!==JSON.stringify(cache[S.editKey].inp);
  updateComputed();markDirty();
});
$('#detailBody')?.addEventListener('change',e=>{
});
$('#detailBody')?.addEventListener('click',async e=>{
  const b=e.target.closest('button');if(!b)return;
  if(b.id==='btnEdit'){toggleDD($('#ddEdit'));return}
  if(b.dataset.ek){S.editKey=b.dataset.ek;S.draft=null;S.dirty=false;$('#detailBody').dataset.built='';closeDD();renderDetail();return}
  if(b.id==='btnReset'){S.draft=null;S.dirty=false;$('#detailBody').dataset.built='';renderDetail()}
  if(b.id==='btnSave'){
    b.disabled=true;$('#saveMsg').textContent='Menyimpan…';
    try{
      const where=await persist(S.editKey,S.draft);
      cache[S.editKey]={inp:clone(S.draft),src:'local'};S.dirty=false;S.avail.add(S.editKey);
      $$('#detailBody .in.chg').forEach(i=>i.classList.remove('chg'));
      $('#saveMsg').textContent=where?'Tersimpan ke '+where+'.':'Tersimpan sementara di halaman ini saja.';
      toast(where?lab(S.editKey)+' tersimpan ke '+where:'Tidak ada penyimpanan aktif, perubahan hanya di halaman ini',!where);
      markDirty();renderFilters();
    }catch(er){
      $('#saveMsg').textContent='Gagal menyimpan.';b.disabled=false;
      toast('Gagal menyimpan: '+(er&&er.message?er.message:'kesalahan tak dikenal'),true);
    }
  }
});

/* ================= welcome ================= */
let raf=0;
function enterApp(){
  const w=$('#welcome');w.classList.add('leave');cancelAnimationFrame(raf);
  setTimeout(()=>{location.href='pages/tabel.html'},450);
}
if($('#startBtn'))$('#startBtn').onclick=enterApp;
(function water(){
  if(!$('#sparks'))return;
  const c=$('#sparks'),x=c.getContext('2d');let W=0,H=0,t=0;const R=[];
  const rm=matchMedia('(prefers-reduced-motion: reduce)').matches;
  const L=[
    {y:.52,a:.035,f:1.6,s:.012,c:'rgba(206,231,243,.55)'},
    {y:.60,a:.040,f:1.2,s:-.010,c:'rgba(120,180,214,.30)'},
    {y:.68,a:.045,f:1.9,s:.014,c:'rgba(169,27,24,.07)'},
    {y:.76,a:.040,f:1.4,s:-.016,c:'rgba(70,140,185,.28)'},
    {y:.86,a:.035,f:2.1,s:.011,c:'rgba(248,250,237,.55)'}
  ];
  function size(){const d=Math.min(devicePixelRatio||1,2);W=c.clientWidth;H=c.clientHeight;c.width=W*d;c.height=H*d;x.setTransform(d,0,0,d,0,0)}
  function ripple(px,py){R.push({x:px,y:py,r:4,a:.5});if(R.length>24)R.shift()}
  const wy=(l,i,px)=>l.y*H+Math.sin(px/W*Math.PI*2*l.f+t*l.s*3+i)*l.a*H+Math.sin(px/W*Math.PI*5+t*l.s*5)*l.a*H*.35;
  function frame(){
    t++;x.clearRect(0,0,W,H);
    const g=x.createRadialGradient(W*.3,H*.2,0,W*.3,H*.2,Math.max(W,H)*.6);
    g.addColorStop(0,'rgba(255,255,255,.65)');g.addColorStop(1,'rgba(255,255,255,0)');x.fillStyle=g;x.fillRect(0,0,W,H);
    L.forEach((l,i)=>{
      x.beginPath();x.moveTo(0,H);
      for(let px=0;px<=W;px+=8)x.lineTo(px,wy(l,i,px));
      x.lineTo(W,H);x.closePath();x.fillStyle=l.c;x.fill();
      x.beginPath();
      for(let px=0;px<=W;px+=8){const py=wy(l,i,px);px?x.lineTo(px,py):x.moveTo(px,py)}
      x.strokeStyle='rgba(255,255,255,.55)';x.lineWidth=1.2;x.stroke();
    });
    for(let i=R.length-1;i>=0;i--){
      const r=R[i];r.r+=1.6;r.a-=.0065;
      if(r.a<=0){R.splice(i,1);continue}
      x.beginPath();x.ellipse(r.x,r.y,r.r,r.r*.42,0,0,6.2832);
      x.strokeStyle='rgba(169,27,24,'+(r.a*.7)+')';x.lineWidth=1.4;x.stroke();
      x.beginPath();x.ellipse(r.x,r.y,r.r*.6,r.r*.25,0,0,6.2832);
      x.strokeStyle='rgba(255,255,255,'+r.a+')';x.lineWidth=1.2;x.stroke();
    }
    if(!rm)raf=requestAnimationFrame(frame);
  }
  addEventListener('resize',size);size();
  if(!rm){setInterval(()=>{if(W)ripple(Math.random()*W,H*(.55+Math.random()*.4))},900)}
  frame();
  let last=0;
  $('#welcome').addEventListener('pointermove',e=>{if(rm)return;const n=performance.now();if(n-last<110)return;last=n;const r=c.getBoundingClientRect();ripple(e.clientX-r.left,e.clientY-r.top)});
  $('#welcome').addEventListener('pointerdown',e=>{const r=c.getBoundingClientRect();ripple(e.clientX-r.left,e.clientY-r.top)});
})();

/* ================= boot ================= */
if(PAGE!=='home'){stLoad();route();connect()}
