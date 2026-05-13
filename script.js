let S = { teamName: 'Enactus Team', university: '', defaultHours: 2, workDays: [1,2,3,4,5], projects: [], members: [], att: {}, meetings: [] };
let wOff = 0, shCtx = null;

let docxLoaded = false;
let docxLoading = true;

// Массив ссылок: если первая не работает, пробуем вторую, затем третью
const cdns = [
  "https://cdnjs.cloudflare.com/ajax/libs/docx/8.5.0/docx.umd.min.js",
  "https://cdn.jsdelivr.net/npm/docx@8.5.0/build/index.umd.js",
  "https://unpkg.com/docx@8.5.0/build/index.umd.js"
];

function loadDocxLibrary(index = 0) {
  if (index >= cdns.length) {
    // Если все ссылки не сработали
    docxLoading = false;
    docxLoaded = false;
    updateDlWordBtnState();
    return;
  }

  updateDlWordBtnState(); // Показываем "Загрузка..."

  const script = document.createElement('script');
  script.src = cdns[index];
  script.async = true;

  script.onload = () => {
    docxLoaded = true;
    docxLoading = false;
    updateDlWordBtnState(); // Показываем "Скачать Word"
    console.log("Библиотека docx успешно загружена из:", cdns[index]);
  };

  script.onerror = () => {
    console.warn("Не удалось загрузить из:", cdns[index], "Пробуем следующий источник...");
    loadDocxLibrary(index + 1); // Пробуем следующую ссылку
  };

  document.head.appendChild(script);
}

// Вызываем загрузку при старте
loadDocxLibrary();

function updateDlWordBtnState() {
  const btn = document.querySelector('#pg-report .btn-main');
  if (!btn) return;
  
  if (docxLoaded) {
    btn.disabled = false;
    btn.textContent = '📄 Скачать Word';
  } else if (docxLoading) {
    btn.disabled = true;
    btn.textContent = '⏳ Загрузка генератора...';
  } else {
    btn.disabled = true;
    btn.textContent = '⚠️ Ошибка загрузки';
  }
}

const CLR=[['#dbeafe','#1e40af'],['#fce7f3','#9d174d'],['#dcfce7','#14532d'],['#fef3c7','#92400e'],['#ede9fe','#4c1d95'],['#ffedd5','#7c2d12'],['#cffafe','#164e63'],['#f3f4f6','#374151']];
const DN=['Пн','Вт','Ср','Чт','Пт','Сб','Вс'], MN=['янв','фев','мар','апр','май','июн','июл','авг','сен','окт','ноя','дек'];

function sv() { try { localStorage.setItem('en_att', JSON.stringify(S)); } catch(e){} }
function ld() { 
  try { const d = localStorage.getItem('en_att'); if(d) Object.assign(S, JSON.parse(d)); } catch(e){} 
  S.members=S.members||[]; S.att=S.att||{}; S.meetings=S.meetings||[]; S.projects=S.projects||[]; 
}

function gp(id, btn) {
  document.querySelectorAll('.page, .bn').forEach(el => el.classList.remove('on'));
  document.getElementById('pg-'+id).classList.add('on');
  btn.classList.add('on');
  document.getElementById('cnt').scrollTop = 0;
  
  if(id==='report') renderReport();
  if(id==='member') { refSel(); renderInd(); }
  if(id==='meeting') renderMtgList();
  if(id==='more') renderMore();
}

function wkDates(o=wOff) {
  const n=new Date(), dow=n.getDay()||7, m=new Date(n);
  m.setDate(n.getDate()-dow+1+o*7); m.setHours(0,0,0,0);
  return Array.from({length:7}, (_,i)=>{ const d=new Date(m); d.setDate(m.getDate()+i); return d; });
}
function wkKey(o=wOff) {
  const d=wkDates(o)[0], y=d.getFullYear(), j=new Date(Date.UTC(y,0,1));
  const w=Math.ceil((((new Date(Date.UTC(y,d.getMonth(),d.getDate()))-j)/864e5)+j.getUTCDay()+1)/7);
  return `${y}-W${String(w).padStart(2,'0')}`;
}
function dk(d) { return d.toISOString().slice(0,10); }
function fd(d) { return d.toLocaleDateString('ru-RU',{day:'2-digit',month:'2-digit'}); }
function setWkUI() {
  const ds=wkDates();
  document.getElementById('wkT').textContent=`Неделя ${wkKey().split('-W')[1]}, ${ds[0].getFullYear()}`;
  document.getElementById('wkS').textContent=`${ds[0].getDate()} ${MN[ds[0].getMonth()]} — ${ds[6].getDate()} ${MN[ds[6].getMonth()]}`;
}
function cw(d) { wOff+=d; setWkUI(); renderJournal(); }
function goNow() { wOff=0; setWkUI(); renderJournal(); }

function ini(n) { return n.split(' ').map(w=>w[0]).join('').slice(0,2).toUpperCase(); }
function clr(i) { return CLR[i%CLR.length]; }
function hrs(s,e) {
  if(!s||!e) return 0;
  const[sh,sm]=s.split(':').map(Number), [eh,em]=e.split(':').map(Number), h=(eh*60+em-sh*60-sm)/60; 
  return h>0?h:0;
}
function rb(r) { return {president:'<span class="rb rb-pr">President</span>',vp:'<span class="rb rb-vp">VP</span>',member:'<span class="rb rb-mb">Member</span>',advisor:'<span class="rb rb-ad">Advisor</span>'}[r]||''; }

function addMember() {
  const n=document.getElementById('nName').value.trim();
  if(!n) return toast('Введите имя ⚠️');
  S.members.push({id:Date.now()+'', name:n, role:document.getElementById('nRole').value});
  document.getElementById('nName').value=''; sv(); renderMembers(); renderJournal(); refSel(); toast('Добавлен: '+n+' ✅');
}
function delMember(id) {
  if(!confirm('Удалить участника?')) return;
  S.members = S.members.filter(m=>m.id!==id);
  for(const wk in S.att) delete S.att[wk][id];
  sv(); renderMembers(); renderJournal(); refSel(); toast('Удалено 🗑');
}
function renderMembers() {
  const el=document.getElementById('memberList');
  if(!S.members.length) return el.innerHTML='<div class="empty"><div class="ei">👥</div>Добавьте участников</div>';
  el.innerHTML=S.members.map((m,i)=>{
    const[bg,fg]=clr(i);
    return`<div class="prow"><div class="ava" style="background:${bg};color:${fg}">${ini(m.name)}</div><div style="flex:1;min-width:0"><div style="font-weight:700;font-size:14px">${m.name}</div><div style="margin-top:2px">${rb(m.role)}</div></div><button class="del" onclick="delMember('${m.id}')">✕</button></div>`;
  }).join('');
}

function addProj() {
  const n=document.getElementById('projInp').value.trim();
  if(!n) return toast('Введите название проекта');
  S.projects.push({id:Date.now()+'', name:n}); document.getElementById('projInp').value=''; sv(); renderProjList(); toast(n+' ✅');
}
function delProj(id) { S.projects=S.projects.filter(p=>p.id!==id); sv(); renderProjList(); }
function renderProjList() {
  document.getElementById('projList').innerHTML=S.projects.map(p=>`<div style="display:flex;align-items:center;gap:8px;padding:7px 0;border-bottom:1px solid #f3f4f6"><div style="flex:1;font-size:14px;font-weight:600">📁 ${p.name}</div><button class="del" onclick="delProj('${p.id}')">✕</button></div>`).join('');
}
function projChipsHTML(selId='') {
  if(!S.projects.length) return `<span style="font-size:12px;color:var(--muted)">Нет проектов</span>`;
  return S.projects.map(p=>`<div class="pchip ${p.id===selId?'on':''}" onclick="selProj(this,'${p.id}')">${p.name}</div>`).join('');
}
function selProj(el,id) { el.closest('.pch').querySelectorAll('.pchip').forEach(c=>c.classList.remove('on')); el.classList.add('on'); }
function getSelProj(cid) { const s=document.querySelector(`#${cid} .pchip.on`); return s?s.getAttribute('onclick').match(/'([^']+)'/g)[1].replace(/'/g,''):''; }

function getRec(mid, dkey) {
  const wk=wkKey();
  if(!S.att[wk]) S.att[wk]={};
  if(!S.att[wk][mid]) S.att[wk][mid]={};
  if(!S.att[wk][mid][dkey]) S.att[wk][mid][dkey]={status:'', start:'17:00', end:'19:00', project:''};
  return S.att[wk][mid][dkey];
}
function cycle(mid, dkey) {
  const r=getRec(mid, dkey), c=['','present','absent'];
  r.status = c[(c.indexOf(r.status)+1)%c.length];
  sv(); r.status==='present' ? openSheet(mid, dkey) : renderJournal();
}
function renderJournal() {
  const dates=wkDates(), workD=dates.filter((_,i)=>S.workDays.includes(i+1)), today=dk(new Date());
  let p=0, a=0, h=0;
  S.members.forEach(mb=>workD.forEach(d=>{
    const r=getRec(mb.id,dk(d));
    if(r.status==='present'){ p++; h+=hrs(r.start,r.end); } else if(r.status==='absent') a++;
  }));
  
  const wkMtgs=S.meetings.filter(m=>workD.map(d=>dk(d)).includes(m.date));
  document.getElementById('statsG').innerHTML=`<div class="sc"><div class="sc-ic" style="background:var(--green-lt)">✅</div><div><div class="sc-num" style="color:var(--green)">${p}</div><div class="sc-lbl">Присутствуют</div></div></div><div class="sc"><div class="sc-ic" style="background:var(--red-lt)">❌</div><div><div class="sc-num" style="color:var(--red)">${a}</div><div class="sc-lbl">Отсутствуют</div></div></div><div class="sc"><div class="sc-ic" style="background:var(--amber-lt)">🗣️</div><div><div class="sc-num" style="color:var(--amber)">${wkMtgs.length}</div><div class="sc-lbl">Собраний</div></div></div><div class="sc"><div class="sc-ic" style="background:var(--or-lt)">⏱️</div><div><div class="sc-num" style="color:var(--or)">${h.toFixed(1)}</div><div class="sc-lbl">Часов</div></div></div>`;

  const banner=document.getElementById('mtgBanner');
  banner.innerHTML=wkMtgs.length ? wkMtgs.map(m=>{
    const proj=S.projects.find(p=>p.id===m.project);
    return`<div class="mtg-card"><div class="mtg-title">🗣️ ${m.topic||'Собрание'}</div><div class="mtg-meta">${m.date} · ${m.start}–${m.end}${proj?' · 📁 '+proj.name:''} · Присутствовали: ${m.attendees.length} чел.</div></div>`;
  }).join('') : '';

  if(!S.members.length) return document.getElementById('journalList').innerHTML=`<div class="empty"><div class="ei">👥</div>Добавьте участников во вкладке «Ещё»</div>`;

  document.getElementById('journalList').innerHTML=S.members.map((m,mi)=>{
    const[bg,fg]=clr(mi); let tH=0;
    const chips=workD.map(d=>{
      const dkey=dk(d), r=getRec(m.id,dkey), st=r.status, isT=dkey===today;
      const lbl={present:'П',absent:'О','':`<span style="font-size:10px;color:#ccc">·</span>`}[st];
      if(st==='present') tH+=hrs(r.start,r.end);
      return`<div class="dc" onclick="cycle('${m.id}','${dkey}')"><span class="dname">${DN[d.getDay()-1||6]}</span><div class="chip ${st}${isT?' today-r':''}">${lbl}</div><span class="ddate">${fd(d)}</span></div>`;
    }).join('');
    return`<div class="mcard"><div class="mcard-top"><div class="ava" style="background:${bg};color:${fg}">${ini(m.name)}</div><div style="flex:1;min-width:0"><div class="mname">${m.name}</div><div class="mrole">${rb(m.role)}</div></div><span class="mhrs">${tH.toFixed(1)} ч</span></div><div class="chips-row">${chips}</div></div>`;
  }).join('');
}

function openSheet(mid, dkey) {
  const r=getRec(mid,dkey), m=S.members.find(x=>x.id===mid), d=new Date(dkey);
  shCtx={mid,dkey};
  document.getElementById('tsTitle').textContent=`${m?.name} · ${DN[d.getDay()-1||6]} ${fd(d)}`;
  document.getElementById('tsS').value=r.start||'17:00';
  document.getElementById('tsE').value=r.end||'19:00';
  document.getElementById('tsPr').innerHTML=projChipsHTML(r.project);
  document.getElementById('tSheet').classList.add('on'); document.getElementById('ov').classList.add('on');
}
function openMtgSheet() {
  document.getElementById('mDate').value=new Date().toISOString().slice(0,10);
  document.getElementById('mTopic').value=''; document.getElementById('mPr').innerHTML=projChipsHTML();
  document.getElementById('mSheet').classList.add('on'); document.getElementById('ov').classList.add('on');
}
function closeSheet() {
  document.querySelectorAll('.bsheet, .overlay').forEach(b=>b.classList.remove('on')); shCtx=null;
}
function saveSheet() {
  if(!shCtx) return;
  const r=getRec(shCtx.mid, shCtx.dkey);
  r.start=document.getElementById('tsS').value; r.end=document.getElementById('tsE').value; r.project=getSelProj('tsPr');
  sv(); closeSheet(); renderJournal();
}
function saveMtg() {
  const date=document.getElementById('mDate').value;
  if(!date) return toast('Выберите дату ⚠️');
  S.meetings.push({id:Date.now()+'', date, start:document.getElementById('mStart').value, end:document.getElementById('mEnd').value, topic:document.getElementById('mTopic').value.trim(), project:getSelProj('mPr'), attendees:[]});
  sv(); closeSheet(); renderMtgList(); toast('Собрание добавлено ✅');
}

function renderMtgList() {
  const el=document.getElementById('mtgList');
  if(!S.meetings.length) return el.innerHTML=`<div class="empty"><div class="ei">🗓️</div>Собраний ещё нет</div>`;
  el.innerHTML=S.meetings.slice().sort((a,b)=>b.date.localeCompare(a.date)).map(m=>{
    const proj=S.projects.find(p=>p.id===m.project);
    return`<div class="mtg-card"><div style="display:flex;align-items:flex-start;justify-content:space-between;gap:8px"><div><div class="mtg-title">${m.topic||'Собрание'}</div><div class="mtg-meta">${m.date} · ${m.start}–${m.end}${proj?' · 📁 '+proj.name:''}</div></div><button style="background:var(--red-lt);color:var(--red);border:none;border-radius:8px;padding:6px 10px;font-size:12px;font-weight:700;cursor:pointer;" onclick="delMtg('${m.id}')">Удалить</button></div><div style="font-size:12px;font-weight:700;color:var(--muted);margin-bottom:8px">Отметить присутствие (${m.attendees.length}/${S.members.length}):</div><div class="att-grid">${S.members.length?S.members.map((mb,mi)=>{const[bg,fg]=clr(mi),on=m.attendees.includes(mb.id);return`<div class="att-chip ${on?'on':''}" onclick="togAtt('${m.id}','${mb.id}')"><div class="att-chip-ava" style="background:${bg};color:${fg}">${ini(mb.name)}</div><span class="att-chip-name">${mb.name.split(' ')[0]}</span>${on?'<span class="att-check">✓</span>':''}</div>`;}).join(''):'<span style="font-size:12px;color:var(--muted)">Добавьте участников</span>'}</div></div>`;
  }).join('');
}
function delMtg(id) { S.meetings=S.meetings.filter(m=>m.id!==id); sv(); renderMtgList(); }
function togAtt(mtgId,mbId) {
  const m=S.meetings.find(x=>x.id===mtgId); if(!m) return;
  m.attendees.includes(mbId) ? m.attendees=m.attendees.filter(x=>x!==mbId) : m.attendees.push(mbId);
  sv(); renderMtgList();
}

function refSel() {
  const s=document.getElementById('indSel'), v=s.value;
  s.innerHTML='<option value="">— выберите участника —</option>'+S.members.map(m=>`<option value="${m.id}">${m.name}</option>`).join('');
  if(v) s.value=v;
}
function renderInd() {
  const id=document.getElementById('indSel').value, el=document.getElementById('indContent');
  if(!id) return el.innerHTML='';
  const m=S.members.find(x=>x.id===id); if(!m) return;
  const mi=S.members.indexOf(m), [bg,fg]=clr(mi);
  let p=0, a=0, h=0, days=0; const rows=[];
  
  Object.keys(S.att).sort().forEach(wk=>{
    const wd=(S.att[wk]||{})[id]||{}; let wp=0,wa=0,wh=0,prSet=new Set();
    Object.values(wd).forEach(r=>{
      if(r.status==='present'){ wp++;wh+=hrs(r.start,r.end); p++;h+=hrs(r.start,r.end); days++; if(r.project)prSet.add(r.project); }
      else if(r.status==='absent'){ wa++;a++;days++; }
    });
    if(wp+wa) rows.push({wk,wp,wa,wh,prjs:[...prSet]});
  });
  
  const mtgCnt=S.meetings.filter(m=>m.attendees.includes(id)).length, rate=days?Math.round(p/days*100):0;
  el.innerHTML=`<div class="ind-hdr"><div class="ind-ava" style="background:${bg};color:${fg}">${ini(m.name)}</div><div style="flex:1"><div style="font-size:16px;font-weight:800">${m.name}</div><div style="margin-top:4px">${rb(m.role)}</div></div><div style="text-align:right"><div class="pct-n">${rate}%</div><div style="font-size:11px;color:var(--muted)">посещаемость</div><div class="pct-bar"><div class="pct-fill" style="width:${rate}%"></div></div></div></div><div class="sg" style="margin-bottom:10px"><div class="sc"><div class="sc-ic" style="background:var(--green-lt)">✅</div><div><div class="sc-num" style="color:var(--green)">${p}</div><div class="sc-lbl">Присутствий</div></div></div><div class="sc"><div class="sc-ic" style="background:var(--red-lt)">❌</div><div><div class="sc-num" style="color:var(--red)">${a}</div><div class="sc-lbl">Отсутствий</div></div></div><div class="sc"><div class="sc-ic" style="background:var(--amber-lt)">🗣️</div><div><div class="sc-num" style="color:var(--amber)">${mtgCnt}</div><div class="sc-lbl">Собраний</div></div></div><div class="sc"><div class="sc-ic" style="background:var(--or-lt)">⏱️</div><div><div class="sc-num" style="color:var(--or)">${h.toFixed(1)}</div><div class="sc-lbl">Часов</div></div></div></div>${rows.length?rows.map(r=>{const pn=r.prjs.map(pid=>{const pr=S.projects.find(x=>x.id===pid);return pr?'📁 '+pr.name:'';}).filter(Boolean).join(', ');return`<div class="wk-item"><div class="wk-lbl">${r.wk}</div><span class="badge bg">П ${r.wp}</span><span class="badge br">О ${r.wa}</span><span class="badge bo">${r.wh.toFixed(1)} ч</span>${pn?`<div style="width:100%;font-size:11px;color:var(--muted);margin-top:2px">${pn}</div>`:''}</div>`;}).join(''):`<div class="empty"><div class="ei">📭</div>Нет данных</div>`}`;
}

function renderReport() {
  const el=document.getElementById('repContent'), wks=Object.keys(S.att).sort();
  if(!wks.length||!S.members.length) return el.innerHTML='<div class="empty"><div class="ei">📭</div>Нет данных</div>';
  el.innerHTML=wks.map(wk=>{
    const wkD=S.att[wk]||{}, allDK=[...new Set(Object.values(wkD).flatMap(d=>Object.keys(d)))].sort();
    const wkMtgs=S.meetings.filter(m=>allDK.includes(m.date));
    const rows=S.members.map((m,mi)=>{
      const days=wkD[m.id]||{}; let p=0,a=0,h=0;
      allDK.forEach(dkey=>{const r=days[dkey]||{status:''};if(r.status==='present'){p++;h+=hrs(r.start||'17:00',r.end||'19:00');}else if(r.status==='absent')a++;});
      if(!p&&!a) return'';
      const[bg,fg]=clr(mi);
      return`<div class="rp-row"><div class="ava" style="background:${bg};color:${fg};width:26px;height:26px;font-size:10px;border-radius:7px;flex-shrink:0">${ini(m.name)}</div><div class="rp-name">${m.name}</div>${rb(m.role)}<span class="badge bg">П ${p}</span><span class="badge br">О ${a}</span><span class="badge bo">${h.toFixed(1)}ч</span></div>`;
    }).join('');
    if(!rows.trim()) return'';
    return`<div class="rp-wk"><div class="rp-wk-t">📅 ${wk}</div>${wkMtgs.map(m=>`<div class="rp-mtg">🗣️ ${m.topic||'Собрание'} · ${m.date} · Присутствовали: ${m.attendees.length} чел.</div>`).join('')}${rows}</div>`;
  }).filter(Boolean).join('')||'<div class="empty"><div class="ei">📭</div>Нет данных</div>';
}

async function dlWord() {
  if (!docxLoaded) return toast('Ждите загрузки библиотеки... ⏳');
  if(typeof docx==='undefined') return toast('Библиотека не загружена ⚠️');
  const{Document,Packer,Paragraph,TextRun,Table,TableRow,TableCell,AlignmentType,WidthType,BorderStyle,ShadingType,HeadingLevel,PageBreak}=docx;
  const brd={style:BorderStyle.SINGLE,size:1,color:'DDDDDD'},borders={top:brd,bottom:brd,left:brd,right:brd};
  const cell=(txt,o={})=>new TableCell({borders,width:{size:o.w||1500,type:WidthType.DXA},margins:{top:80,bottom:80,left:120,right:120},shading:o.fill?{fill:o.fill,type:ShadingType.CLEAR}:undefined,children:[new Paragraph({alignment:o.c?AlignmentType.CENTER:AlignmentType.LEFT,children:[new TextRun({text:String(txt),bold:!!o.b,size:o.sz||20,font:'Arial',color:'000000'})]})]});

  const ch=[
    new Paragraph({heading:HeadingLevel.HEADING_1,children:[new TextRun({text:`Отчёт посещаемости — Enactus ${S.teamName}`,font:'Arial',bold:true,size:36})]}),
    new Paragraph({children:[new TextRun({text:`${S.university?S.university+' · ':''}Дата: ${new Date().toLocaleDateString('ru-RU')}`,font:'Arial',size:20,color:'888888'})]}),
    new Paragraph({children:[]})
  ];

  for(const wk of Object.keys(S.att).sort()){
    const wkD=S.att[wk]||{};
    const allDK=[...new Set(Object.values(wkD).flatMap(d=>Object.keys(d)).filter(dkey=>{const d=new Date(dkey);return S.workDays.includes(d.getDay()||7);}))].sort();
    if(!allDK.length) continue;
    ch.push(new Paragraph({heading:HeadingLevel.HEADING_2,children:[new TextRun({text:`Неделя: ${wk}`,font:'Arial',bold:true,size:28})]}));
    
    const wkMtgs=S.meetings.filter(m=>allDK.includes(m.date));
    wkMtgs.forEach(m=>{const proj=S.projects.find(p=>p.id===m.project);const names=m.attendees.map(id=>{const mb=S.members.find(x=>x.id===id);return mb?mb.name:'';}).filter(Boolean).join(', ');ch.push(new Paragraph({children:[new TextRun({text:`🗣️ ${m.topic||'Собрание'} · ${m.date} · ${m.start}–${m.end}${proj?' · '+proj.name:''} · Присутствовали: ${names||'—'}`,font:'Arial',size:20,color:'FF6B00'})]}));});
    
    const cw=Math.floor((9360-3000)/Math.max(allDK.length,1));
    const hRow=new TableRow({children:[cell('Участник (роль)',{w:3000,b:true,fill:'FFF4EC'}),...allDK.map(dkey=>{const d=new Date(dkey);return cell(`${DN[d.getDay()-1||6]} ${fd(d)}`,{w:cw,b:true,fill:'FFF4EC',c:true});}),cell('Часов',{w:900,b:true,fill:'FFF4EC',c:true})]});
    
    const dRows=S.members.map(m=>{
      const days=wkD[m.id]||{}; let h=0;
      const cells=allDK.map(dkey=>{const r=days[dkey]||{status:''};let t='—',fill='FFFFFF';if(r.status==='present'){t='П';fill='F0FDF4';h+=hrs(r.start||'17:00',r.end||'19:00');}else if(r.status==='absent'){t='О';fill='FFF5F5';}return cell(t,{w:cw,c:true,fill});});
      return new TableRow({children:[cell(`${m.name} (${m.role})`,{w:3000}),...cells,cell(h.toFixed(1)+'ч',{w:900,c:true,b:true})]});
    });
    ch.push(new Table({width:{size:9360,type:WidthType.DXA},columnWidths:[3000,...allDK.map(()=>cw),900],rows:[hRow,...dRows]}),new Paragraph({children:[]}),new Paragraph({children:[new PageBreak()]}));
  }

  if(S.meetings.length){
    ch.push(new Paragraph({heading:HeadingLevel.HEADING_2,children:[new TextRun({text:'Собрания команды',font:'Arial',bold:true,size:28})]}));
    S.meetings.sort((a,b)=>a.date.localeCompare(b.date)).forEach(m=>{
      const proj=S.projects.find(p=>p.id===m.project);
      const names=m.attendees.map(id=>{const mb=S.members.find(x=>x.id===id);return mb?mb.name:'';}).filter(Boolean).join(', ');
      ch.push(new Paragraph({children:[new TextRun({text:`${m.date} · ${m.start}–${m.end}${proj?' · 📁 '+proj.name:''}`,font:'Arial',bold:true,size:22})]}));
      if(m.topic)ch.push(new Paragraph({children:[new TextRun({text:`Тема: ${m.topic}`,font:'Arial',size:20})]}));
      ch.push(new Paragraph({children:[new TextRun({text:`Присутствовали (${m.attendees.length}): ${names||'—'}`,font:'Arial',size:20})]}));
      ch.push(new Paragraph({children:[]}));
    });
    ch.push(new Paragraph({children:[new PageBreak()]}));
  }

  ch.push(new Paragraph({heading:HeadingLevel.HEADING_2,children:[new TextRun({text:'Итоговая сводка',font:'Arial',bold:true,size:28})]}));
  const gR=S.members.map(m=>{
    let p=0,a=0,h=0,days=0;
    Object.values(S.att).forEach(wkD=>{Object.values((wkD[m.id])||{}).forEach(r=>{if(r.status==='present'){p++;h+=hrs(r.start||'17:00',r.end||'19:00');days++;}else if(r.status==='absent'){a++;days++;}});});
    const mtgC=S.meetings.filter(x=>x.attendees.includes(m.id)).length, rate=days?Math.round(p/days*100):0;
    return new TableRow({children:[cell(`${m.name} (${m.role})`,{w:2800}),cell(String(p),{w:900,c:true,fill:'F0FDF4'}),cell(String(a),{w:900,c:true,fill:'FFF5F5'}),cell(String(mtgC),{w:900,c:true,fill:'FFF4EC'}),cell(h.toFixed(1)+'ч',{w:1000,c:true}),cell(rate+'%',{w:860,c:true,b:true})]});
  });
  ch.push(new Table({width:{size:9360,type:WidthType.DXA},columnWidths:[2800,900,900,900,1000,860],rows:[new TableRow({children:[cell('Участник',{w:2800,b:true,fill:'FFF4EC'}),cell('Присут.',{w:900,b:true,fill:'FFF4EC',c:true}),cell('Отсутст.',{w:900,b:true,fill:'FFF4EC',c:true}),cell('Собраний',{w:900,b:true,fill:'FFF4EC',c:true}),cell('Часов',{w:1000,b:true,fill:'FFF4EC',c:true}),cell('% посещ.',{w:860,b:true,fill:'FFF4EC',c:true})]}), ...gR]}));

  const doc=new Document({styles:{default:{document:{run:{font:'Arial',size:22}}},paragraphStyles:[{id:'Heading1',name:'Heading 1',basedOn:'Normal',next:'Normal',quickFormat:true,run:{size:36,bold:true,font:'Arial'},paragraph:{spacing:{before:240,after:240},outlineLevel:0}},{id:'Heading2',name:'Heading 2',basedOn:'Normal',next:'Normal',quickFormat:true,run:{size:28,bold:true,font:'Arial'},paragraph:{spacing:{before:200,after:160},outlineLevel:1}}]},sections:[{properties:{page:{size:{width:12240,height:15840},margin:{top:1080,right:1080,bottom:1080,left:1080}}},children:ch}]});
  const blob=await Packer.toBlob(doc), url=URL.createObjectURL(blob), a=document.createElement('a');
  a.href=url; a.download=`Enactus_${S.teamName}_Посещаемость_${new Date().toLocaleDateString('ru-RU').replace(/\./g,'-')}.docx`; a.click(); URL.revokeObjectURL(url);
  toast('Файл скачан ✅');
}

function renderMore() {
  document.getElementById('sTeam').value=S.teamName;
  document.getElementById('sUni').value=S.university||'';
  document.getElementById('sHrs').value=S.defaultHours;
  renderWdC(); renderProjList(); renderMembers();
}
function renderWdC() {
  const el=document.getElementById('wdC'); if(!el) return;
  el.innerHTML=DN.map((_,i)=>{const d=i+1,on=S.workDays.includes(d);return`<div class="wdc ${on?'on':''}" onclick="togWd(${d},this)">${DN[i]}</div>`;}).join('');
}
function togWd(d,el) {
  if(S.workDays.includes(d)) S.workDays=S.workDays.filter(x=>x!==d); else S.workDays.push(d);
  S.workDays.sort(); el.classList.toggle('on',S.workDays.includes(d));
}
function saveSet() {
  S.teamName=document.getElementById('sTeam').value.trim()||'Enactus Team';
  S.university=document.getElementById('sUni').value.trim();
  S.defaultHours=parseInt(document.getElementById('sHrs').value)||2;
  document.getElementById('hTeam').textContent=S.teamName;
  document.getElementById('hUni').textContent=S.university||'Учёт посещаемости';
  sv(); renderJournal(); toast('Сохранено ✅');
}
function clearAtt() { if(!confirm('Очистить все отметки посещаемости?')) return; S.att={}; sv(); renderJournal(); toast('Очищено 🗑'); }

function expJ() {
  const b=new Blob([JSON.stringify(S,null,2)],{type:'application/json'}), u=URL.createObjectURL(b), a=document.createElement('a');
  a.href=u; a.download=`enactus_backup_${Date.now()}.json`; a.click(); URL.revokeObjectURL(u); toast('Экспортировано 📤');
}
function impJ(e) {
  const f=e.target.files[0]; if(!f) return;
  const r=new FileReader();
  r.onload=ev=>{ try{ Object.assign(S,JSON.parse(ev.target.result)); sv(); init(); toast('Импортировано 📥'); }catch{ toast('Ошибка файла ❌'); } };
  r.readAsText(f); e.target.value='';
}

function toast(msg) {
  const el=document.getElementById('toast');
  el.textContent=msg; el.classList.add('on');
  clearTimeout(el._t); el._t=setTimeout(()=>el.classList.remove('on'), 2400);
}

function init() {
  ld();
  document.getElementById('hTeam').textContent=S.teamName;
  document.getElementById('hUni').textContent=S.university||'Учёт посещаемости';
  setWkUI(); renderJournal(); refSel(); renderWdC();
}

init();

if (!localStorage.getItem('en_att_initialized')) {
  S.teamName = 'Enactus ИГУ';
  S.university = 'ИГУ им. Касыма Тыныстанова';
  S.projects = [{id:'p1',name:'EcoStart'},{id:'p2',name:'FoodShare KG'}];
  S.members = [
    {id:'m1',name:'Айгерим Бекова',role:'president'},
    {id:'m2',name:'Нурлан Осмонов',role:'vp'},
    {id:'m3',name:'Гульмира Токоева',role:'member'},
    {id:'m4',name:'Адиль Сатыбалдиев',role:'member'},
    {id:'m5',name:'Айпери Мамытова',role:'member'}
  ];
  document.getElementById('hTeam').textContent=S.teamName;
  document.getElementById('hUni').textContent=S.university;
  localStorage.setItem('en_att_initialized', 'true'); 
  sv(); renderJournal(); refSel();
}