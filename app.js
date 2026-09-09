import { initializeApp } from "https://www.gstatic.com/firebasejs/12.0.0/firebase-app.js";
import { getAuth, onAuthStateChanged, signInWithEmailAndPassword, signOut } from "https://www.gstatic.com/firebasejs/12.0.0/firebase-auth.js";
import { getFirestore, collection, doc, getDoc, getDocs, addDoc, updateDoc, deleteDoc, query, orderBy, serverTimestamp, writeBatch } from "https://www.gstatic.com/firebasejs/12.0.0/firebase-firestore.js";

const APP_VERSION = "1.14.0-idempotent-dedup";

const firebaseConfig = {
  apiKey: "AIzaSyDhzrMhSPA_S8keOjU6QL2Tath3jFBY9Vs",
  authDomain: "rekap-po.firebaseapp.com",
  projectId: "rekap-po",
  storageBucket: "rekap-po.firebasestorage.app",
  messagingSenderId: "465019621219",
  appId: "1:465019621219:web:34cf437815d45363e54d56"
};

const app = initializeApp(firebaseConfig);
const auth = getAuth(app);
const db = getFirestore(app);

const GROUPS = ["BZOA","BZOB","BZOC","BZOD","BZOE","BBOA","BBOB","BBOC","BBOD","BBOE","BWPOA","BWPOB","BWPOC","BWPOD","BWPOE"];
const MONTHS = ["Januari","Februari","Maret","April","Mei","Juni","Juli","Agustus","September","Oktober","November","Desember"];
const COLORS = ["#60a5fa","#22d3ee","#34d399","#fbbf24","#f472b6","#a78bfa","#fb923c","#818cf8","#10b981","#fb7185","#38bdf8","#c084fc","#84cc16","#4ade80","#94a3b8"];
const YEAR_START = 2021;
const currentYear = new Date().getFullYear();

let state = {
  user: null,
  profile: null,
  page: "dashboard",
  year: currentYear,
  month: new Date().getMonth() + 1,
  registerGroup: null,
  allPOs: [],
  periodPOs: [],
  currentEditId: null,
  quickAction: null,
  pendingConfirm: null,
  pendingVoice: null,
  scanner: { stream: null, detector: null, fallbackReader: null, fallbackPromise: null, raf: 0, running: false, scanning: false },
  barcodeSvg: "",
  barcodePO: "",
  importPreview: null,
  xlsxLibraryPromise: null,
  duplicateGroups: []
};

const $ = (id) => document.getElementById(id);

window.addEventListener("error", (event) => {
  const el = $("bootError");
  if (el && event?.message) {
    el.style.display = "block";
    el.style.color = "#b91c1c";
    el.textContent = "Kesalahan aplikasi: " + event.message;
  }
});
window.addEventListener("unhandledrejection", (event) => {
  const el = $("bootError");
  const msg = event?.reason?.message || String(event?.reason || "Terjadi kesalahan.");
  if (el) {
    el.style.display = "block";
    el.style.color = "#b91c1c";
    el.textContent = "Kesalahan aplikasi: " + msg;
  }
  console.error(event?.reason);
});

function showLoading(show) { $("loading").classList.toggle("hidden", !show); }
function toast(message, type="") {
  const el = document.createElement("div");
  el.className = "toast" + (type ? " " + type : "");
  el.textContent = message;
  $("toastWrap").appendChild(el);
  setTimeout(() => el.remove(), 3600);
}
function todayISO() {
  const d = new Date();
  return `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,"0")}-${String(d.getDate()).padStart(2,"0")}`;
}
function formatDate(value) {
  if (!value) return "-";
  const [y,m,d] = String(value).split("-");
  return y && m && d ? `${d}/${m}/${y}` : String(value);
}
function escapeHTML(value) {
  return String(value ?? "").replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/"/g,"&quot;").replace(/'/g,"&#039;");
}
function normalizePO(value) {
  let s = String(value ?? "").trim().toUpperCase().replace(/\s+/g, "");
  if (!s) return "";
  if (/^PO-?\d+$/.test(s)) return `PO-${s.replace(/^PO-?/i, "").padStart(5,"0")}`;
  if (/^\d+$/.test(s)) return `PO-${s.padStart(5,"0")}`;
  return s;
}
function numericTail(value) {
  const m = String(value ?? "").match(/(\d+)$/);
  return m ? Number(m[1]) : null;
}
function comparePO(a,b) {
  const na = numericTail(a.noPO), nb = numericTail(b.noPO);
  if (na !== null && nb !== null && na !== nb) return na - nb;
  return String(a.noPO||"").localeCompare(String(b.noPO||""), undefined, {numeric:true, sensitivity:"base"});
}
function getStatus(po) { return po.tglKembali ? "SUDAH_KEMBALI" : "BELUM_KEMBALI"; }
function statusHTML(status) { return status === "SUDAH_KEMBALI" ? '<span class="status status-in">🟢 Sudah Kembali</span>' : '<span class="status status-out">🔴 Belum Kembali</span>'; }
function getRegisterCode(group, year, month) { return `${group}${String(year).slice(-2)}-${String.fromCharCode(64 + month)}`; }
function getLegacyCode(group, year, month) { return `${group}${String(year).slice(-2)}${String.fromCharCode(64 + month)}`; }
function matchesPeriod(po, group, year, month) {
  if (po.registerGroup && po.year && po.month) return String(po.registerGroup).toUpperCase() === group && Number(po.year) === Number(year) && Number(po.month) === Number(month);
  const code = String(po.registerCode || "").toUpperCase().replace(/[^A-Z0-9]/g, "");
  const expected = `${group}${String(year).slice(-2)}${String.fromCharCode(64 + month)}`.toUpperCase();
  if (code === expected) return true;
  const date = String(po.tglKirim || "");
  if (/^\d{4}-\d{2}-\d{2}$/.test(date)) return date.startsWith(`${year}-${String(month).padStart(2,"0")}`) && (po.registerCode ? code.startsWith(group) : true);
  return false;
}
function resolveStoredPeriod(po) {
  if (po.registerGroup && po.year && po.month) return {group:String(po.registerGroup).toUpperCase(), year:Number(po.year), month:Number(po.month)};
  const raw = String(po.registerCode || "").toUpperCase().replace(/[^A-Z0-9]/g, "");
  for (const g of GROUPS) {
    if (raw.startsWith(g)) {
      const yearPart = raw.slice(g.length, g.length + 2);
      const monthLetter = raw.slice(g.length + 2, g.length + 3);
      const y = Number(yearPart);
      const m = monthLetter >= "A" && monthLetter <= "L" ? monthLetter.charCodeAt(0) - 64 : 0;
      if (y && m) return {group:g, year:2000+y, month:m};
      const dt = String(po.tglKirim || "");
      if (/^\d{4}-\d{2}-\d{2}$/.test(dt)) return {group:g, year:Number(dt.slice(0,4)), month:Number(dt.slice(5,7))};
    }
  }
  const dt = String(po.tglKirim || "");
  if (/^\d{4}-\d{2}-\d{2}$/.test(dt)) return {group:"",year:Number(dt.slice(0,4)),month:Number(dt.slice(5,7))};
  return {group:"",year:0,month:0};
}
function currentCode() { return state.registerGroup ? getRegisterCode(state.registerGroup, state.year, state.month) : ""; }
function isAdmin() { return state.profile?.role === "admin"; }

function periodStorageKey(){
  return `poTracking.period.${state.user?.uid || "guest"}`;
}
function loadSavedPeriod(){
  try{
    const raw=localStorage.getItem(periodStorageKey());
    if(!raw)return false;
    const data=JSON.parse(raw);
    const y=Number(data?.year),m=Number(data?.month);
    if(y>=YEAR_START && y<=currentYear+1 && m>=1 && m<=12){state.year=y;state.month=m;return true;}
  }catch(err){console.warn("Gagal membaca periode tersimpan",err);}
  return false;
}
function saveCurrentPeriod(){
  try{localStorage.setItem(periodStorageKey(),JSON.stringify({year:state.year,month:state.month}));}catch(err){console.warn("Gagal menyimpan periode",err);}
}
function parseImportDate(value){
  if(value===null || value===undefined || value==="") return "";
  if(value instanceof Date && !Number.isNaN(value.getTime())) return `${value.getFullYear()}-${String(value.getMonth()+1).padStart(2,"0")}-${String(value.getDate()).padStart(2,"0")}`;
  if(typeof value==="number" && Number.isFinite(value)){
    // Excel serial date (1900 date system).
    const utc = new Date(Date.UTC(1899,11,30) + Math.round(value)*86400000);
    if(!Number.isNaN(utc.getTime())) return `${utc.getUTCFullYear()}-${String(utc.getUTCMonth()+1).padStart(2,"0")}-${String(utc.getUTCDate()).padStart(2,"0")}`;
  }
  const s=String(value).trim();
  if(!s)return "";
  let m=s.match(/^(\d{4})[-\/.](\d{1,2})[-\/.](\d{1,2})$/);
  if(m)return `${m[1]}-${String(Number(m[2])).padStart(2,"0")}-${String(Number(m[3])).padStart(2,"0")}`;
  m=s.match(/^(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{4})$/);
  if(m)return `${m[3]}-${String(Number(m[2])).padStart(2,"0")}-${String(Number(m[1])).padStart(2,"0")}`;
  const d=new Date(s);
  if(!Number.isNaN(d.getTime()))return `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,"0")}-${String(d.getDate()).padStart(2,"0")}`;
  return "";
}
function normalizeHeader(value){return String(value??"").trim().toLowerCase().replace(/[._\-\/]+/g," ").replace(/\s+/g," ");}
function findHeader(row,variants){
  const map={};
  Object.keys(row||{}).forEach(k=>map[normalizeHeader(k)]=k);
  for(const v of variants){const key=normalizeHeader(v);if(map[key])return map[key];}
  return null;
}
function parseRegisterInput(value,sendDate){
  const raw=String(value??"").trim().toUpperCase();
  const compact=raw.replace(/[^A-Z0-9]/g,"");
  for(const g of GROUPS){
    if(compact.startsWith(g) && compact.length>=g.length+3){
      const yy=Number(compact.slice(g.length,g.length+2));
      const letter=compact.slice(g.length+2,g.length+3);
      const m=letter>="A"&&letter<="L"?letter.charCodeAt(0)-64:0;
      if(yy>=0&&m){return {group:g,year:2000+yy,month:m,code:getRegisterCode(g,2000+yy,m)};}
    }
  }
  const date=sendDate?parseImportDate(sendDate):"";
  if(date){const year=Number(date.slice(0,4)),month=Number(date.slice(5,7));for(const g of GROUPS){if(compact===g) return {group:g,year,month,code:getRegisterCode(g,year,month)};}}
  return null;
}
function inferRegisterFromPO(noPO,sendDate){
  const raw=String(noPO??"").trim().toUpperCase().replace(/\s+/g,"");
  const date=parseImportDate(sendDate);
  for(const g of GROUPS){
    if(raw.startsWith(g)){
      const rest=raw.slice(g.length).replace(/[^A-Z0-9]/g,"");
      const yy=Number(rest.slice(0,2)),letter=rest.slice(2,3);
      const m=letter>="A"&&letter<="L"?letter.charCodeAt(0)-64:0;
      if(yy>=0&&m)return {group:g,year:2000+yy,month:m,code:getRegisterCode(g,2000+yy,m)};
      if(date){const y=Number(date.slice(0,4)),mo=Number(date.slice(5,7));return {group:g,year:y,month:mo,code:getRegisterCode(g,y,mo)};}
    }
  }
  return null;
}
async function loadXLSXLibrary(){
  if(state.xlsxLibraryPromise)return state.xlsxLibraryPromise;
  if(window.XLSX)return Promise.resolve(window.XLSX);
  state.xlsxLibraryPromise=(async()=>{
    try{
      const mod=await import("https://cdn.sheetjs.com/xlsx-0.20.3/package/xlsx.mjs");
      return mod?.default?.read ? mod.default : mod;
    }catch(primaryError){
      console.warn("ESM SheetJS gagal, mencoba UMD fallback.",primaryError);
      return await new Promise((resolve,reject)=>{
        const existing=document.querySelector('script[data-xlsx-loader="1"]');
        if(existing){existing.addEventListener("load",()=>resolve(window.XLSX));existing.addEventListener("error",()=>reject(primaryError));return;}
        const script=document.createElement("script");script.src="https://cdn.sheetjs.com/xlsx-0.20.3/package/dist/xlsx.full.min.js";script.async=true;script.dataset.xlsxLoader="1";
        script.onload=()=>window.XLSX?resolve(window.XLSX):reject(new Error("Library Excel tidak menyediakan objek XLSX."));
        script.onerror=()=>reject(primaryError);document.head.appendChild(script);
      });
    }
  })();
  try{return await state.xlsxLibraryPromise;}catch(err){state.xlsxLibraryPromise=null;throw err;}
}
function hexToRgba(hex, alpha) {
  const m = String(hex || "").replace("#", "");
  if (m.length !== 6) return `rgba(37,99,235,${alpha})`;
  const r = parseInt(m.slice(0,2),16);
  const g = parseInt(m.slice(2,4),16);
  const b = parseInt(m.slice(4,6),16);
  return `rgba(${r},${g},${b},${alpha})`;
}
function groupAccent(group) {
  const idx = Math.max(0, GROUPS.indexOf(group));
  return COLORS[idx % COLORS.length];
}
function displayRegisterCode(po) {
  const period = resolveStoredPeriod(po);
  const group = period.group || String(po?.registerGroup || "").toUpperCase();
  const year = period.year || Number(po?.year || 0);
  const month = period.month || Number(po?.month || 0);
  if (group && year >= 2000 && month >= 1 && month <= 12) return getRegisterCode(group, year, month);
  const raw = String(po?.registerCode || "").trim().toUpperCase();
  const compact = raw.replace(/[^A-Z0-9]/g, "");
  for (const g of GROUPS) {
    if (compact.startsWith(g) && compact.length >= g.length + 3) {
      const yy = Number(compact.slice(g.length, g.length + 2));
      const ml = compact.slice(g.length + 2, g.length + 3);
      if (yy && ml >= "A" && ml <= "L") return getRegisterCode(g, 2000 + yy, ml.charCodeAt(0) - 64);
    }
  }
  return raw || "-";
}

function populatePeriods() {
  $("yearSelect").innerHTML = "";
  for (let y = YEAR_START; y <= currentYear + 1; y++) {
    const op = document.createElement("option"); op.value = String(y); op.textContent = String(y); if (y === state.year) op.selected = true; $("yearSelect").appendChild(op);
  }
  $("monthSelect").innerHTML = "";
  MONTHS.forEach((name, i) => { const op=document.createElement("option"); op.value=String(i+1); op.textContent=`${String.fromCharCode(65+i)} — ${name}`; if (i+1===state.month) op.selected=true; $("monthSelect").appendChild(op); });
  $("periodLabel").textContent = `${MONTHS[state.month-1]} ${state.year}`;
  $("registerSectionTitle").textContent = `Register ${MONTHS[state.month-1]} ${state.year}`;
}
function populateRegisterSelect(selected="") {
  $("poRegister").innerHTML = GROUPS.map(g=>`<option value="${g}">${escapeHTML(getRegisterCode(g,state.year,state.month))}</option>`).join("");
  if (selected) $("poRegister").value = selected;
}
function populateReportFilters() {
  $("reportYear").innerHTML = "<option value=\"ALL\">Semua Tahun</option>";
  for (let y=YEAR_START;y<=currentYear+1;y++){ const op=document.createElement("option");op.value=String(y);op.textContent=String(y);$("reportYear").appendChild(op); }
  $("reportYear").value=String(state.year);
  $("reportMonth").innerHTML = "<option value=\"ALL\">Semua Bulan</option>" + MONTHS.map((m,i)=>`<option value="${i+1}">${String.fromCharCode(65+i)} — ${m}</option>`).join("");
  $("reportGroup").innerHTML = "<option value=\"ALL\">Semua Register</option>" + GROUPS.map(g=>`<option value="${g}">${escapeHTML(getRegisterCode(g,state.year,state.month))}</option>`).join("");
  $("reportStatus").value="ALL";
  $("reportMonth").value=String(state.month);
  if($("reportDateField")) $("reportDateField").value="send";
}
function renderRegisterNav() {
  $("registerNav").innerHTML = GROUPS.map(g=>`<button type="button" class="nav-btn" data-group="${g}">📁 ${g}</button>`).join("");
  $("registerNav").querySelectorAll("[data-group]").forEach(btn=>btn.addEventListener("click",()=>openRegister(btn.dataset.group)));
}
function renderRegisterCards() {
  const wrap = $("registerGrid");
  wrap.innerHTML = "";
  GROUPS.forEach((g, i) => {
    const code = getRegisterCode(g, state.year, state.month);
    const legacy = getLegacyCode(g, state.year, state.month);
    const rows = state.allPOs.filter(po => matchesPeriod(po, g, state.year, state.month));
    const out = rows.filter(po => getStatus(po) === "BELUM_KEMBALI").length;
    const card = document.createElement("div");
    card.className = "register-card";
    const accent = COLORS[i % COLORS.length];
    const glow = hexToRgba(accent, 0.46);
    card.style.setProperty("--accent", accent);
    card.style.setProperty("--glow", glow);
    card.style.borderLeftColor = accent;
    card.style.background = `linear-gradient(135deg, ${hexToRgba(accent, 0.22)} 0%, ${hexToRgba(accent, 0.08)} 44%, #ffffff 84%)`;
    card.innerHTML = `<div class="register-code">${escapeHTML(code)}</div><div class="register-meta">${rows.length} PO tercatat</div>${out ? `<div class="register-alert">🔴 ${out} belum kembali</div>` : `<div class="register-ok">✓ Tidak ada yang tertahan</div>`}`;
    card.title = `${code} (${legacy})`;
    card.addEventListener("click", () => openRegister(g));
    wrap.appendChild(card);
  });
}
function applyFilter(rows, search, from, to, status) {
  const q=String(search||"").trim().toLowerCase();
  return rows.filter(po=>{
    if(q){ const hay=[po.noPO,po.namaCustomer,po.registerCode,po.keterangan,po.registerGroup].join(" ").toLowerCase(); if(!hay.includes(q)) return false; }
    if(from && String(po.tglKirim||"")<from) return false;
    if(to && String(po.tglKirim||"")>to) return false;
    if(status!=="ALL" && getStatus(po)!==status) return false;
    return true;
  }).sort(comparePO);
}
function renderTable(rows, container, includeRegister=true) {
  if(!rows.length){container.innerHTML='<div class="empty">📭 Tidak ada data PO.</div>';return;}
  let html=`<table><thead><tr><th>No</th>${includeRegister?"<th>Register</th>":""}<th>Tgl Kirim</th><th>Tgl Kembali</th><th>No PO</th><th>Nama Customer</th><th>Keterangan</th><th>Status</th><th>Aksi</th></tr></thead><tbody>`;
  rows.forEach((po,i)=>{
    const reg=displayRegisterCode(po);
    const st=getStatus(po);
    html+=`<tr><td>${i+1}</td>${includeRegister?`<td>${escapeHTML(reg)}</td>`:""}<td>${formatDate(po.tglKirim)}</td><td>${formatDate(po.tglKembali)}</td><td><strong>${escapeHTML(po.noPO)}</strong></td><td>${escapeHTML(po.namaCustomer)}</td><td>${escapeHTML(po.keterangan||"-")}</td><td>${statusHTML(st)}</td><td><div class="row-actions">${st==="BELUM_KEMBALI"?`<button type="button" class="btn btn-success btn-small" data-act="in" data-id="${po.id}">IN</button>`:`<button type="button" class="btn btn-warning btn-small" data-act="out" data-id="${po.id}">OUT</button>`}<button type="button" class="btn btn-outline btn-small" data-act="edit" data-id="${po.id}">Edit</button>${isAdmin()?`<button type="button" class="btn btn-danger btn-small" data-act="delete" data-id="${po.id}">Hapus</button>`:""}</div></td></tr>`;
  });
  html+='</tbody></table>'; container.innerHTML=html;
  container.querySelectorAll("[data-act]").forEach(btn=>btn.addEventListener("click",()=>tableAction(btn.dataset.act,btn.dataset.id)));
}
function updateStats() {
  const total=state.periodPOs.length, returned=state.periodPOs.filter(po=>getStatus(po)==="SUDAH_KEMBALI").length, out=total-returned, today=state.periodPOs.filter(po=>po.tglKirim===todayISO()).length;
  $("statTotal").textContent=String(total);$("statReturned").textContent=String(returned);$("statOutstanding").textContent=String(out);$("statToday").textContent=String(today);
}
function setPeriodData() {
  state.periodPOs=state.allPOs.filter(po=>{
    if(state.registerGroup) return matchesPeriod(po,state.registerGroup,state.year,state.month);
    return resolveStoredPeriod(po).year===state.year && resolveStoredPeriod(po).month===state.month;
  }).sort(comparePO);
  updateStats(); renderRegisterCards(); populatePeriods();
}
function navigate(page) {
  state.page=page;
  document.querySelectorAll(".page").forEach(el=>el.classList.add("hidden"));
  const target=$("page-"+page); if(target) target.classList.remove("hidden");
  document.querySelectorAll(".nav-btn[data-page]").forEach(btn=>btn.classList.toggle("active",btn.dataset.page===page));
  $("sidebar").classList.remove("open");
  if(page==="dashboard") renderDashboard();
  if(page==="register") renderCurrentRegister();
  if(page==="all-po") renderAllPO();
  if(page==="report") renderReport();
  if(page==="audit") loadAudit();
  if(page==="users") loadUsers();
  if(page==="data") loadImportHistory();
}
function renderDashboard(){updateStats();}
async function openRegister(group){ state.registerGroup=group; $("registerTitle").textContent=getRegisterCode(group,state.year,state.month); populateRegisterSelect(group); navigate("register"); }
function renderCurrentRegister(){
  if(!state.registerGroup)return;
  const rows=state.allPOs.filter(po=>matchesPeriod(po,state.registerGroup,state.year,state.month));
  const filtered=applyFilter(rows,$("regSearch").value,$("regFrom").value,$("regTo").value,$("regStatus").value);
  renderTable(filtered,$("registerTable"),false);
}
function renderAllPO(){
  const filtered=applyFilter(state.allPOs,$("allSearch").value,$("allFrom").value,$("allTo").value,$("allStatus").value); renderTable(filtered,$("allTable"),true);
}
function reportFilteredRows(){
  let rows=[...state.allPOs]; const y=$("reportYear").value, m=$("reportMonth").value, g=$("reportGroup").value, st=$("reportStatus").value, q=$("reportSearch").value, from=$("reportFrom").value, to=$("reportTo").value, dateField=$("reportDateField")?.value||"send";
  if(y!=="ALL") rows=rows.filter(po=>resolveStoredPeriod(po).year===Number(y));
  if(m!=="ALL") rows=rows.filter(po=>resolveStoredPeriod(po).month===Number(m));
  if(g!=="ALL") rows=rows.filter(po=>resolveStoredPeriod(po).group===g);
  if(from || to){
    const key=dateField==="return"?"tglKembali":"tglKirim";
    rows=rows.filter(po=>{const d=String(po[key]||""); if(!d)return false; if(from&&d<from)return false; if(to&&d>to)return false; return true;});
  }
  return applyFilter(rows,q,"","",st);
}
function renderReport(){const rows=reportFilteredRows();$("reportTotal").textContent=rows.length;$("reportReturned").textContent=rows.filter(po=>getStatus(po)==="SUDAH_KEMBALI").length;$("reportOutstanding").textContent=rows.filter(po=>getStatus(po)==="BELUM_KEMBALI").length;renderTable(rows,$("reportTable"),true);}

async function loadPOs(){
  const snap=await getDocs(collection(db,"po_documents"));
  state.allPOs=snap.docs.map(d=>({id:d.id,...d.data()}));
  state.allPOs.sort(comparePO); setPeriodData();
}
async function refresh(){showLoading(true);try{await loadPOs();if(state.page==="register")renderCurrentRegister();if(state.page==="all-po")renderAllPO();if(state.page==="report")renderReport();if(state.page==="dashboard")renderDashboard();}catch(err){console.error(err);toast("Gagal memuat data PO.","error");}finally{showLoading(false)}}
async function writeAudit(action,targetPO="",registerCode="",details=""){
  if(!state.user)return;
  try{await addDoc(collection(db,"audit_logs"),{userId:state.user.uid,userName:state.profile?.name||state.user.email||"-",action,targetPO,registerCode,details,timestamp:serverTimestamp()});}catch(err){console.warn("Audit gagal",err);}
}
function openPOForm(po=null){
  state.currentEditId=po?.id||null; $("poModalTitle").textContent=po?"Edit PO":"Tambah PO"; populateRegisterSelect(po?.registerGroup||state.registerGroup||GROUPS[0]);
  const p=po?resolveStoredPeriod(po):{year:state.year,month:state.month};
  $("poSendDate").value=po?.tglKirim||todayISO();$("poReturnDate").value=po?.tglKembali||"";$("poNo").value=po?.noPO||"";$("poCustomer").value=po?.namaCustomer||"";$("poNote").value=po?.keterangan||"";$("poModal").classList.remove("hidden");setTimeout(()=>$("poNo").focus(),40);
  if(po && (p.year!==state.year || p.month!==state.month)) toast("Data lama dibuka; periode asli tetap dipertahankan.","warn");
}
function closeModal(id){$(id).classList.add("hidden");}
async function savePO(e){
  e.preventDefault();
  const id=state.currentEditId;
  const group=$("poRegister").value;
  const tglKirim=$("poSendDate").value;
  const tglKembali=$("poReturnDate").value||null;
  const noPO=normalizePO($("poNo").value);
  const customer=$("poCustomer").value.trim();
  const note=$("poNote").value.trim();

  if(!group||!tglKirim||!noPO||!customer){toast("Lengkapi field wajib.","error");return}
  if(tglKembali && tglKembali<tglKirim){toast("Tgl Kembali tidak boleh sebelum Tgl Kirim.","error");return}

  const duplicate=state.allPOs.find(po=>normalizePO(po.noPO)===noPO && po.id!==id);
  if(duplicate){toast(`${noPO} sudah terdaftar di ${duplicate.registerCode||resolveStoredPeriod(duplicate).group}.`,"error");return}

  const period={year:Number(tglKirim.slice(0,4)),month:Number(tglKirim.slice(5,7))};
  const registerCode=getRegisterCode(group,period.year,period.month);
  const data={
    registerGroup:group,
    year:period.year,
    month:period.month,
    registerCode,
    tglKirim,
    tglKembali,
    noPO,
    uniqueKey:noPO,
    namaCustomer:customer,
    keterangan:note,
    status:tglKembali?"SUDAH_KEMBALI":"BELUM_KEMBALI",
    updatedAt:serverTimestamp(),
    updatedBy:state.user.uid
  };

  try{
    showLoading(true);
    const batch=writeBatch(db);

    if(id){
      const oldPO=state.allPOs.find(po=>po.id===id);
      if(!oldPO){throw new Error("PO yang diedit tidak ditemukan.")}
      const oldKey=oldPO.uniqueKey||normalizePO(oldPO.noPO);

      if(normalizePO(oldPO.noPO)!==noPO && !isAdmin()){
        throw new Error("Staff tidak diperbolehkan mengganti No PO. Gunakan Edit data lain atau minta Admin mengubah No PO.");
      }

      const ref=doc(db,"po_documents",id);
      batch.update(ref,data);

      if(oldKey && oldKey!==noPO){
        batch.delete(doc(db,"po_unique",oldKey));
      }

      batch.set(doc(db,"po_unique",noPO),{
        noPO,
        poId:id,
        updatedAt:serverTimestamp(),
        updatedBy:state.user.uid
      });

      await batch.commit();
      await writeAudit("UPDATE_PO",noPO,registerCode,"Mengubah data PO");
      toast(`${noPO} berhasil diperbarui.`,"success");
    }else{
      const ref=doc(collection(db,"po_documents"));
      batch.set(ref,{
        ...data,
        createdAt:serverTimestamp(),
        createdBy:state.user.uid
      });
      batch.set(doc(db,"po_unique",noPO),{
        noPO,
        poId:ref.id,
        createdAt:serverTimestamp(),
        createdBy:state.user.uid
      });
      await batch.commit();
      await writeAudit("CREATE_PO",noPO,registerCode,"Menambahkan PO baru");
      toast(`${noPO} berhasil ditambahkan.`,"success");
    }

    closeModal("poModal");
    await refresh();
    navigate("dashboard");
  }catch(err){
    console.error(err);
    const message=String(err?.message||"");
    toast(message.startsWith("Staff tidak")?message:"Gagal menyimpan data PO.","error");
  }finally{
    showLoading(false);
  }
}

async function updateStatus(po,action,source="Quick Action"){
  const target=action==="IN"?todayISO():null;
  const newStatus=target?"SUDAH_KEMBALI":"BELUM_KEMBALI";

  try{
    showLoading(true);
    const batch=writeBatch(db);
    const poRef=doc(db,"po_documents",po.id);
    const key=po.uniqueKey||normalizePO(po.noPO);

    batch.update(poRef,{
      uniqueKey:key,
      tglKembali:target,
      status:newStatus,
      updatedAt:serverTimestamp(),
      updatedBy:state.user.uid
    });

    batch.set(doc(db,"po_unique",key),{
      noPO:po.noPO,
      poId:po.id,
      updatedAt:serverTimestamp(),
      updatedBy:state.user.uid
    });

    await batch.commit();
    await writeAudit(action==="IN"?"PO_IN":"PO_OUT",po.noPO,po.registerCode||currentCode(),`${source} — ${action}`);
    toast(`${po.noPO} berhasil diproses ${action}.`,"success");
    await refresh();
  }catch(err){
    console.error(err);
    toast("Gagal memperbarui status PO.","error");
  }finally{
    showLoading(false);
  }
}

function openConfirmSingle(po,action){state.pendingConfirm={type:"single",items:[{po,action}]};$("confirmTitle").textContent=action==="IN"?"Konfirmasi IN":"Konfirmasi OUT";$("confirmBody").innerHTML=`<p>Anda akan mengubah:</p><div class="scan-result"><strong>${escapeHTML(po.noPO)}</strong><div>${escapeHTML(po.namaCustomer)}</div><div class="hint">${action==="IN"?`Tgl Kembali = ${formatDate(todayISO())}`:"Tgl Kembali dikosongkan"}</div><div style="margin-top:7px">${action==="IN"?'<span class="status status-in">🟢 Sudah Kembali</span>':'<span class="status status-out">🔴 Belum Kembali</span>'}</div></div>`;$("confirmModal").classList.remove("hidden");}

async function runConfirm(){
  const pending=state.pendingConfirm;
  if(!pending)return;
  closeModal("confirmModal");

  if(pending.type==="single"){
    await updateStatus(pending.items[0].po,pending.items[0].action,pending.items[0].source||"Konfirmasi");
    return;
  }

  if(pending.type==="batch"){
    try{
      showLoading(true);
      const batch=writeBatch(db);
      const today=todayISO();
      const seenKeys=new Set();

      pending.items.forEach(({po,action})=>{
        const ref=doc(db,"po_documents",po.id);
        const key=po.uniqueKey||normalizePO(po.noPO);
        if(seenKeys.has(key))throw new Error(`Duplikat No PO terdeteksi dalam batch: ${po.noPO}`);
        seenKeys.add(key);

        batch.update(ref,{
          uniqueKey:key,
          tglKembali:action==="IN"?today:null,
          status:action==="IN"?"SUDAH_KEMBALI":"BELUM_KEMBALI",
          updatedAt:serverTimestamp(),
          updatedBy:state.user.uid
        });

        batch.set(doc(db,"po_unique",key),{
          noPO:po.noPO,
          poId:po.id,
          updatedAt:serverTimestamp(),
          updatedBy:state.user.uid
        });
      });

      await batch.commit();

      for(const item of pending.items){
        await writeAudit(item.action==="IN"?"VOICE_BATCH_IN":"VOICE_BATCH_OUT",item.po.noPO,item.po.registerCode||"",`Batch Voice — ${pending.items.length} PO`);
      }

      toast(`${pending.items.length} PO berhasil diproses.`,"success");
      await refresh();
    }catch(err){
      console.error(err);
      toast("Batch gagal diproses. Tidak ada perubahan sebagian yang diterapkan.","error");
    }finally{
      showLoading(false);
    }
  }
}

async function tableAction(action,id){
  const po=state.allPOs.find(x=>x.id===id);
  if(!po)return;

  if(action==="edit"){
    openPOForm(po);
    return;
  }

  if(action==="delete"){
    if(!isAdmin()){toast("Hanya Admin yang dapat menghapus.","error");return}
    if(!confirm(`Hapus ${po.noPO}?`))return;

    try{
      showLoading(true);
      const batch=writeBatch(db);
      batch.delete(doc(db,"po_documents",po.id));
      const key=po.uniqueKey||normalizePO(po.noPO);
      if(key)batch.delete(doc(db,"po_unique",key));
      await batch.commit();
      await writeAudit("DELETE_PO",po.noPO,po.registerCode||"","Menghapus data PO");
      toast(`${po.noPO} berhasil dihapus.`,"success");
      await refresh();
    }catch(err){
      console.error(err);
      toast("Gagal menghapus PO.","error");
    }finally{
      showLoading(false);
    }
    return;
  }

  if(action==="in"||action==="out")openConfirmSingle(po,action.toUpperCase());
}

function openQuick(action){state.quickAction=action;$("quickTitle").textContent=action==="IN"?"📥 PO Kembali / IN":"📤 PO Keluar / OUT";$("quickHint").textContent=action==="IN"?"Tanggal kembali otomatis hari ini.":"Tanggal kembali dikosongkan.";$("quickPo").value="";$("quickModal").classList.remove("hidden");setTimeout(()=>$("quickPo").focus(),40);}
function compactPO(value){
  return String(value ?? '').trim().toUpperCase().replace(/\s+/g,'');
}
function findPOFromInput(value){
  const raw=String(value ?? '').trim();
  if(!raw) return {matches:[]};
  const compact=compactPO(raw);
  const normalized=normalizePO(raw);

  // 1) Exact match against the stored No PO.
  const exact=state.allPOs.filter(po=>compactPO(po.noPO)===compact);
  if(exact.length===1) return {po:exact[0],mode:'exact'};
  if(exact.length>1) return {matches:exact,mode:'exact-ambiguous'};

  // 2) Exact match after normalizing PO-001 / PO 001 / 001.
  if(normalized){
    const normalizedMatches=state.allPOs.filter(po=>normalizePO(po.noPO)===normalized);
    if(normalizedMatches.length===1) return {po:normalizedMatches[0],mode:'normalized'};
    if(normalizedMatches.length>1) return {matches:normalizedMatches,mode:'normalized-ambiguous'};
  }

  // 3) Short numeric input (001 / 129 / PO 001) may match by numeric tail,
  //    but only when it is unique across the entire dataset.
  const tail=numericTail(raw);
  if(tail!==null){
    const matches=state.allPOs.filter(po=>numericTail(po.noPO)===tail);
    if(matches.length===1) return {po:matches[0],mode:'tail-unique'};
    return {matches,mode:matches.length?'tail-ambiguous':'tail-none'};
  }
  return {matches:[],mode:'none'};
}
async function runQuick(){
  const v=$('quickPo').value;
  const found=findPOFromInput(v);
  if(found.po){
    closeModal('quickModal');
    const current=getStatus(found.po);
    if((state.quickAction==='IN' && current==='SUDAH_KEMBALI') || (state.quickAction==='OUT' && current==='BELUM_KEMBALI')){
      toast(`${found.po.noPO} sudah berstatus ${current==='SUDAH_KEMBALI'?'Sudah Kembali':'Belum Kembali'}. Tidak ada perubahan diperlukan.`,'warn');
      return;
    }
    openConfirmSingle(found.po,state.quickAction,'Quick Action');
    return;
  }
  const matches=found.matches||[];
  if(matches.length===0){toast('PO tidak ditemukan.','error');return;}
  toast(`Nomor ${v} ambigu. Ada ${matches.length} PO dengan nomor akhir sama. Gunakan No PO lengkap.`,'warn');
}

const BASIC_NUMBER_WORDS={
  nol:0,satu:1,sebuah:1,dua:2,tiga:3,empat:4,lima:5,enam:6,tujuh:7,delapan:8,sembilan:9,
  sepuluh:10,sebelas:11,
  dua_belas:12,tiga_belas:13,empat_belas:14,lima_belas:15,enam_belas:16,tujuh_belas:17,delapan_belas:18,sembilan_belas:19,
  dua_puluh:20,tiga_puluh:30,empat_puluh:40,lima_puluh:50,enam_puluh:60,tujuh_puluh:70,delapan_puluh:80,sembilan_puluh:90,
  seratus:100,dua_ratus:200,tiga_ratus:300,empat_ratus:400,lima_ratus:500,enam_ratus:600,tujuh_ratus:700,delapan_ratus:800,sembilan_ratus:900,
  seribu:1000,dua_ribu:2000,tiga_ribu:3000,empat_ribu:4000,lima_ribu:5000
};
function numberWordValue(word){
  return Object.prototype.hasOwnProperty.call(BASIC_NUMBER_WORDS,word)?BASIC_NUMBER_WORDS[word]:null;
}
function parseIndonesianNumberPhrase(phrase){
  const tokens=String(phrase||'').toLowerCase().replace(/,/g,' ').trim().split(/\s+/).filter(Boolean);
  if(!tokens.length) return null;
  let total=0,current=0,seen=false;
  for(let i=0;i<tokens.length;i++){
    const token=tokens[i].replace(/-/g,'_');
    if(token==='dan') continue;
    const value=numberWordValue(token);
    if(value!==null){ current += value; seen=true; continue; }
    if(token==='puluh'){ current=(current||1)*10; seen=true; continue; }
    if(token==='ratus'){ current=(current||1)*100; seen=true; continue; }
    if(token==='ribu'){ total += (current||1)*1000; current=0; seen=true; continue; }
    return null;
  }
  const result=total+current;
  return seen && result>0 ? result : null;
}
function voiceActionFromText(text){
  const lower=String(text||'').toLowerCase();
  const inHit=/\b(sudah\s+kembali|sudah\s+masuk|kembali|masuk|diterima(?:\s+kembali)?|in)\b/i.test(lower);
  const outHit=/\b(keluar|kirim|dikirim|dikirimkan|out)\b/i.test(lower);
  if(inHit&&outHit){
    return {error:"Perintah mengandung dua aksi berbeda (IN dan OUT). Ucapkan salah satunya saja."};
  }
  if(inHit){
    return {action:'IN'};
  }
  if(outHit){
    return {action:'OUT'};
  }
  return {error:"Aksi belum dikenali. Gunakan kata 'masuk/kembali' atau 'keluar/kirim'."};
}

function extractNumericVoiceTokens(text){
  const raw=String(text||'');
  const out=[];
  const seen=new Set();
  // Ambil angka yang berdiri sendiri, termasuk setelah kata PO.
  // Angka yang menempel pada huruf (mis. BZOA26A001) tidak diambil di sini
  // karena akan ditangani lebih dahulu oleh full-code matcher.
  const re=/(?<![A-Z0-9])(?:PO\s*[-:]?\s*)?(\d{1,6})(?![A-Z0-9])/gi;
  for(const m of raw.matchAll(re)){
    const token=m[1];
    // Jangan menganggap tahun sebagai nomor PO.
    if(/^20\d{2}$/.test(token)) continue;
    const key=token.replace(/^0+(?=\d)/,'')||'0';
    if(!seen.has(key)){
      seen.add(key);
      out.push(token);
    }
  }
  return out;
}

function extractVoicePOs(text){
  const raw=String(text||'').trim();
  const candidates=[];
  const seen=new Set();

  // Full PO codes first: BZOA26A001, BZOA26-A001, BBOA26A003-T, etc.
  const fullMatches=[...raw.matchAll(/\b[A-Z]{2,}(?:\s*[0-9]{2})\s*[A-L]?\s*-?\s*[0-9]+(?:-[A-Z]+)?\b/gi)].map(m=>m[0].trim());
  for(const token of fullMatches){
    const key=compactPO(token);
    if(key&&!seen.has(key)){seen.add(key);candidates.push(token);}
  }

  // Plain numeric sequence, which is the main mobile-friendly command form:
  // "PO 129, 193, 194 masuk".
  for(const token of extractNumericVoiceTokens(raw)){
    const key=normalizePO(token)||compactPO(token);
    if(key&&!seen.has(key)){seen.add(key);candidates.push(token);}
  }

  // Indonesian spoken numbers. Parse only number-like runs and exclude
  // ordinary sentence words so the parser does not invent PO numbers.
  const wordPattern=/\b(?:nol|satu|sebuah|dua|tiga|empat|lima|enam|tujuh|delapan|sembilan|sepuluh|sebelas)(?:\s+(?:dua|tiga|empat|lima|enam|tujuh|delapan|sembilan|puluh|ratus|ribu|belas|nol|satu|seratus|seribu)){0,6}\b/gi;
  for(const m of raw.matchAll(wordPattern)){
    const phrase=m[0];
    const value=parseIndonesianNumberPhrase(phrase);
    if(value!==null && value>=1 && value<=999999){
      const token=String(value);
      const key=normalizePO(token);
      if(key&&!seen.has(key)){seen.add(key);candidates.push(token);}
    }
  }
  return candidates;
}

function buildVoiceBatchPreview(parsed){
  if(!parsed) return '';
  let html=`<div><strong>Perintah:</strong> ${escapeHTML(parsed.text)}</div>`;
  html+=`<div class="hint" style="margin:5px 0 10px">Aksi: <strong>${escapeHTML(parsed.action)}</strong> • Terdeteksi <strong>${parsed.tokens.length}</strong> nomor</div>`;
  parsed.results.forEach((r,i)=>{
    html+=`<div style="padding:9px 0;border-top:1px solid var(--border)"><strong>${i+1}. ${escapeHTML(r.po.noPO)}</strong><div class="hint">${escapeHTML(r.po.namaCustomer||'-')} • ${escapeHTML(displayRegisterCode(r.po))} • ${statusHTML(getStatus(r.po))}</div></div>`;
  });
  parsed.unresolved.forEach((u)=>{
    const detail=u.reason||((u.matches||[]).length?'Nomor ambigu; gunakan No PO lengkap.':'PO tidak ditemukan.');
    html+=`<div style="padding:9px 0;border-top:1px solid var(--border);color:#991b1b"><strong>⚠ ${escapeHTML(u.token)} tidak aman diproses</strong><div class="hint">${escapeHTML(detail)}</div>`;
    if(u.matches?.length){
      html+=`<div class="hint" style="margin-top:5px">Pilihan: ${u.matches.slice(0,6).map(p=>escapeHTML(p.noPO)).join(' • ')}</div>`;
    }
    html+='</div>';
  });
  return html;
}

function parseVoiceCommand(text){
  const raw=String(text||'').trim();
  if(!raw) return null;

  const actionInfo=voiceActionFromText(raw);
  if(actionInfo.error) return {error:actionInfo.error,text:raw};
  const action=actionInfo.action;

  const cleaned=raw
    .replace(/\btahun\s+\d{4}\b/gi,' ')
    .replace(/\b(?:bulan|periode)\s+(?:januari|februari|maret|april|mei|juni|juli|agustus|september|oktober|november|desember)\b/gi,' ');

  const tokens=extractVoicePOs(cleaned);
  if(!tokens.length) return {error:'Nomor PO tidak ditemukan dalam perintah.',text:raw};
  if(tokens.length>100) return {error:'Maksimal 100 PO dapat diproses dalam satu perintah Voice.',text:raw};

  const results=[];
  const unresolved=[];
  const seenIds=new Set();

  for(const token of tokens){
    const found=findPOFromInput(token);
    if(found.po){
      if(seenIds.has(found.po.id)) continue;
      seenIds.add(found.po.id);
      const current=getStatus(found.po);
      const noop=(action==='IN'&&current==='SUDAH_KEMBALI')||(action==='OUT'&&current==='BELUM_KEMBALI');
      if(noop){
        unresolved.push({token,matches:[found.po],reason:action==='IN'?'PO sudah kembali; tidak ada perubahan diperlukan.':'PO masih belum kembali; tidak ada perubahan diperlukan.'});
      }else{
        results.push({po:found.po,action});
      }
    }else{
      unresolved.push({token,matches:found.matches||[],reason:(found.matches||[]).length?'Nomor PO ambigu. Gunakan No PO lengkap agar tidak salah memilih.':'PO tidak ditemukan.'});
    }
  }

  return {text:raw,action,results,unresolved,tokens};
}

function renderVoiceResult(parsed){
  if(!parsed){$("voiceResult").classList.add("hidden");return;}
  if(parsed.error){
    $("voiceResult").classList.remove("hidden");
    $("voiceResult").innerHTML=`<div style="color:#991b1b;font-weight:800">⚠️ ${escapeHTML(parsed.error)}</div>`;
    $("voiceProcess").classList.add("hidden");
    state.pendingVoice=null;
    return;
  }

  $("voiceResult").classList.remove("hidden");
  $("voiceResult").innerHTML=buildVoiceBatchPreview(parsed);

  // A batch is allowed only when every token resolved uniquely and every
  // resolved PO actually needs the requested state change.
  if(parsed.results.length>0 && parsed.unresolved.length===0){
    state.pendingVoice=parsed;
    $("voiceProcess").textContent=`✓ Proses ${parsed.results.length} PO`;
    $("voiceProcess").classList.remove("hidden");
  }else{
    state.pendingVoice=null;
    $("voiceProcess").classList.add("hidden");
  }
}

function openVoice(){
  stopVoice();
  $("voiceModal").classList.remove("hidden");
  $("voiceResult").classList.add("hidden");
  $("voiceProcess").classList.add("hidden");
  $("voiceTranscript").innerHTML='<div class="muted">Hasil suara akan muncul di sini.</div>';
  $("voiceState").textContent='Contoh: “PO 129, 193, 194 masuk”';
  $("voiceInput").value='';
}

let recognition=null;
let recognitionFinalText='';
function startVoice(){
  const SR=window.SpeechRecognition||window.webkitSpeechRecognition;
  if(!SR){
    toast('Browser ini tidak mendukung Voice Command. Gunakan tab Ketik.','warn');
    $("voiceTabType").click();
    return;
  }
  stopVoice();
  recognition=new SR();
  recognition.lang='id-ID';
  recognition.interimResults=true;
  recognition.continuous=false;
  recognitionFinalText='';
  $("voiceState").textContent='🔴 Mendengarkan...';
  $("voiceStart").disabled=true;
  recognition.onresult=e=>{
    let transcript='';
    for(let i=e.resultIndex;i<e.results.length;i++) transcript+=e.results[i][0].transcript;
    transcript=transcript.trim();
    $("voiceTranscript").innerHTML=`<strong>${escapeHTML(transcript||'...')}</strong>`;
    if(e.results[e.results.length-1].isFinal){
      recognitionFinalText=transcript;
      try{
        renderVoiceResult(parseVoiceCommand(transcript));
      }catch(err){
        console.error('Voice parse error:',err);
        state.pendingVoice=null;
        $("voiceProcess").classList.add('hidden');
        $("voiceResult").classList.remove('hidden');
        $("voiceResult").innerHTML=`<div style="color:#991b1b;font-weight:800">⚠️ Gagal menganalisa perintah Voice: ${escapeHTML(err?.message||String(err))}</div>`;
      }
    }
  };
  recognition.onerror=e=>{
    console.warn('Voice recognition:',e);
    $("voiceState").textContent=`Gagal mendengarkan: ${e.error||'unknown'}`;
    $("voiceStart").disabled=false;
  };
  recognition.onend=()=>{
    $("voiceStart").disabled=false;
    if(!recognitionFinalText){
      $("voiceState").textContent='Tidak ada perintah yang dikenali. Coba lagi atau gunakan tab Ketik.';
    }else{
      $("voiceState").textContent='Selesai. Periksa hasil sebelum diproses.';
    }
    recognition=null;
  };
  try{ recognition.start(); }
  catch(err){
    console.warn(err);
    $("voiceStart").disabled=false;
    $("voiceState").textContent='Gagal memulai microphone.';
  }
}
function stopVoice(){
  if(recognition){try{recognition.onend=null;recognition.stop();}catch{}recognition=null;}
  if($("voiceStart")) $("voiceStart").disabled=false;
}
async function processTypedVoice(){
  const value=$("voiceInput").value.trim();
  if(!value){toast('Masukkan perintah Voice terlebih dahulu.','warn');return;}
  try{
    renderVoiceResult(parseVoiceCommand(value));
  }catch(err){
    console.error('Voice parse error:',err);
    state.pendingVoice=null;
    $("voiceProcess").classList.add('hidden');
    $("voiceResult").classList.remove('hidden');
    $("voiceResult").innerHTML=`<div style="color:#991b1b;font-weight:800">⚠️ Gagal menganalisa perintah Voice: ${escapeHTML(err?.message||String(err))}</div>`;
  }
}

async function loadZXingFallback(){
  if(state.scanner.fallbackReader) return state.scanner.fallbackReader;
  if(state.scanner.fallbackPromise) return state.scanner.fallbackPromise;
  state.scanner.fallbackPromise=(async()=>{
    try{
      const mod=await import('https://cdn.jsdelivr.net/npm/@zxing/browser@0.2.1/+esm');
      const Reader=mod.BrowserMultiFormatReader;
      if(!Reader) throw new Error('ZXing BrowserMultiFormatReader tidak tersedia.');
      state.scanner.fallbackReader=new Reader();
      return state.scanner.fallbackReader;
    }catch(err){
      console.warn('Gagal memuat ZXing fallback:',err);
      return null;
    }finally{state.scanner.fallbackPromise=null;}
  })();
  return state.scanner.fallbackPromise;
}
function setScannerStatus(message,type='info'){
  const el=$('scannerStatus');
  if(!el)return;
  const icons={info:'🔎',success:'✅',warn:'⚠️',error:'❌'};
  el.className=`scanner-status ${type}`;
  el.textContent=`${icons[type]||'🔎'} ${message}`;
}
async function startZXingFallback(){
  const reader=await loadZXingFallback();
  if(!reader || !state.scanner.running) return false;
  try{
    setScannerStatus('Scanner kompatibel aktif. Mencari barcode...','info');
    const video=$('cameraVideo');
    state.scanner.fallbackPromise=reader.decodeFromVideoElementContinuously(video,(result)=>{
      if(result && state.scanner.running){
        const text=String(result.getText?.()||result.text||'').trim();
        if(text){
          $('scanInput').value=text;
          setScannerStatus('Barcode ditemukan. Kamera dihentikan.','success');
          stopCamera();
          handleScanSearch(text);
        }
      }
    }).catch(err=>{
      if(state.scanner.running) console.debug('ZXing scan loop:',err);
    }).finally(()=>{state.scanner.fallbackPromise=null;});
    return true;
  }catch(err){
    console.warn('ZXing fallback gagal dijalankan:',err);
    state.scanner.fallbackPromise=null;
    return false;
  }
}
async function startCamera(){
  if(state.scanner.running || state.scanner.stream) return;
  try{
    if(!navigator.mediaDevices?.getUserMedia) throw new Error('Browser tidak mendukung akses kamera.');
    stopCamera();
    const stream=await navigator.mediaDevices.getUserMedia({video:{facingMode:{ideal:'environment'},width:{ideal:1280},height:{ideal:720}},audio:false});
    state.scanner.stream=stream;
    const video=$('cameraVideo');
    video.srcObject=stream;
    video.classList.remove('hidden');
    const placeholder=$('cameraArea').querySelector('.camera-placeholder');
    if(placeholder) placeholder.remove();
    try{await video.play();}catch{}
    state.scanner.running=true;
    $('cameraStart').disabled=true;
    $('cameraStop').disabled=false;
    $('cameraArea').classList.add('scanning-active');
    setScannerStatus('Kamera aktif. Mencari barcode...','info');

    let detectorReady=false;
    if('BarcodeDetector' in window){
      try{
        const supported=await BarcodeDetector.getSupportedFormats();
        const preferred=['code_128','code_39','ean_13','ean_8','upc_a','qr_code'];
        const formats=preferred.filter(f=>supported.includes(f));
        state.scanner.detector=formats.length?new BarcodeDetector({formats}):new BarcodeDetector();
        detectorReady=!!state.scanner.detector;
      }catch(err){
        console.warn('BarcodeDetector tidak dapat digunakan:',err);
        state.scanner.detector=null;
      }
    }else{
      state.scanner.detector=null;
    }

    if(detectorReady){
      setScannerStatus('BarcodeDetector aktif. Mencari barcode...','info');
      scanLoop();
    }else{
      const started=await startZXingFallback();
      if(!started){
        setScannerStatus('Pembaca barcode otomatis tidak tersedia. Gunakan hasil scan manual.','warn');
        toast('Pembaca barcode otomatis tidak tersedia di browser ini.','warn');
      }
    }
  }catch(err){
    console.error(err);
    stopCamera();
    const msg=err?.name==='NotAllowedError'?'Izin kamera ditolak. Izinkan kamera pada browser lalu coba lagi.':(err?.message||String(err));
    toast(`Kamera tidak dapat dibuka: ${msg}`,'error');
  }
}
function stopCamera(){
  state.scanner.running=false;
  state.scanner.scanning=false;
  if(state.scanner.raf) cancelAnimationFrame(state.scanner.raf);
  state.scanner.raf=0;
  if(state.scanner.fallbackReader){
    try{state.scanner.fallbackReader.reset();}catch{}
  }
  state.scanner.fallbackPromise=null;
  if(state.scanner.stream){
    try{state.scanner.stream.getTracks().forEach(track=>track.stop());}catch{}
    state.scanner.stream=null;
  }
  state.scanner.detector=null;
  const video=$('cameraVideo');
  if(video){video.pause?.();video.srcObject=null;video.classList.add('hidden');}
  const area=$('cameraArea');
  if(area){
    area.classList.remove('scanning-active');
    if(!area.querySelector('.camera-placeholder')){
      const placeholder=document.createElement('div');
      placeholder.className='camera-placeholder';
      placeholder.innerHTML='Kamera belum aktif.<br>Tekan “Buka Kamera”.';
      area.appendChild(placeholder);
    }
  }
  setScannerStatus('Kamera berhenti.','info');
  $('cameraStart').disabled=false;
  $('cameraStop').disabled=true;
}
async function scanLoop(){
  if(!state.scanner.running) return;
  if(state.scanner.detector && !state.scanner.scanning && $('cameraVideo').readyState>=2){
    state.scanner.scanning=true;
    try{
      const codes=await state.scanner.detector.detect($('cameraVideo'));
      if(codes?.length){
        const value=String(codes[0].rawValue||'').trim();
        if(value){
          $('scanInput').value=value;
          setScannerStatus('Barcode ditemukan. Kamera dihentikan.','success');
          stopCamera();
          await handleScanSearch(value);
          return;
        }
      }
    }catch(err){console.debug('BarcodeDetector detect:',err)}
    finally{state.scanner.scanning=false;}
  }
  state.scanner.raf=requestAnimationFrame(scanLoop);
}
async function handleScanSearch(value=$('scanInput').value){
  const found=findPOFromInput(value);
  if(found.po){renderScanResult(found.po);return}
  if(found.matches?.length){toast(`Hasil scan ambigu: ${found.matches.length} PO. Gunakan kode lengkap.`,'warn');return}
  toast('PO hasil scan tidak ditemukan.','error');
}
function renderScanResult(po){$("scanResult").classList.remove("hidden");$("scanResult").innerHTML=`<strong>${escapeHTML(po.noPO)}</strong><div>${escapeHTML(po.namaCustomer)}</div><div class="hint">${escapeHTML(po.registerCode||resolveStoredPeriod(po).group||"-")} • Tgl Kirim ${formatDate(po.tglKirim)} • Tgl Kembali ${formatDate(po.tglKembali)}</div><div style="margin-top:8px">${statusHTML(getStatus(po))}</div><div class="actions" style="margin-top:10px"><button type="button" class="btn ${getStatus(po)==="BELUM_KEMBALI"?"btn-success":"btn-warning"}" id="scanActionBtn">${getStatus(po)==="BELUM_KEMBALI"?"IN":"OUT"}</button><button type="button" class="btn btn-outline" id="scanEditBtn">Edit</button></div>`;const scanActionButton = $("scanActionBtn");
  const scanEditButton = $("scanEditBtn");
  if (scanActionButton) scanActionButton.addEventListener("click",()=>openConfirmSingle(po,getStatus(po)==="BELUM_KEMBALI"?"IN":"OUT"));
  if (scanEditButton) scanEditButton.addEventListener("click",()=>{closeModal("scanModal");openPOForm(po);});}
function makeBarcodeSvg(text){
  const value=String(text||"").trim(); if(!value)return "";
  const patterns=["11011001100","11001101100","11001100110","10010011000","10010001100","10001001100","10011001000","10011000100","10001100100","11001001000","11001000100","11000100100","10110011100","10011011100","10011001110","10111001100","10011101100","10011100110","11001110010","11001011100","11001001110","11011100100","11001110100","11101101110","11101001100","11100101100","11100100110","11101100100","11100110100","11100110010","11011011000","11011000110","11000110110","10100011000","10001011000","10001000110","10110001000","10001101000","10001100010","11010001000","11000101000","11000100010","10110111000","10110001110","10001101110","10111011000","10111000110","10001110110","11101110110","11010001110","11000101110","11011101000","11011100010","11011101110","11101011000","11101000110","11100010110","11101101000","11101100010","11100011010","11101111010","11001000010","11110001010","10100110000","10100001100","10010110000","10010000110","10000101100","10000100110","10110010000","10110000100","10011010000","10011000010","10000110100","10000110010","11000010010","11001010000","11110111010","11000010100","10001111010","10100111100","10010111100","10010011110","10111100100","10011110100","10011110010","11110100100","11110010100","11110010010","11011011110","11011110110","11110110110","10101111000","10100011110","10001011110","10111101000","10111100010","11110101000","11110100010","10111011110","10111101110","11101011110","11110101110","11010000100","11010010000","11010011100","1100011101011"];
  const checksum=(v)=>{let sum=104;for(let i=0;i<v.length;i++){const code=v.charCodeAt(i)-32;sum+=code*(i+1)}return sum%103};
  const chars=value.split("").map(ch=>{const c=ch.charCodeAt(0);return c>=32&&c<=127?c:63;});
  const codes=chars.map(c=>c-32); const all=[104,...codes,checksum(value),106]; let bits="";all.forEach(c=>{bits+=patterns[c]||patterns[63]});
  const module=3, height=90, quiet=18, width=bits.length*module+quiet*2; let bars=""; for(let i=0;i<bits.length;i++){if(bits[i]==="1")bars+=`<rect x="${quiet+i*module}" y="10" width="${module}" height="${height}"/>`}
  return `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 ${width} 122" role="img" aria-label="Barcode ${escapeHTML(value)}"><rect width="100%" height="100%" fill="white"/>${bars}<text x="50%" y="118" text-anchor="middle" font-family="Arial,sans-serif" font-size="16" font-weight="700">${escapeHTML(value)}</text></svg>`;
}
function findGeneratePO(){
  const value=$("genPoInput").value.trim();
  if(!value){toast("Masukkan No PO.","error");return null;}
  const found=findPOFromInput(value);
  if(found.po) return found.po;
  if(found.matches?.length){toast(`No PO ambigu: ditemukan ${found.matches.length} PO. Gunakan No PO lengkap.`,"warn");return null;}
  toast("PO tidak ditemukan.","error");
  return null;
}
function generateBarcode(){
  const po=findGeneratePO();
  if(!po) return;
  state.barcodePO=po.noPO;
  state.barcodeSvg=makeBarcodeSvg(po.noPO);
  $("barcodeArea").innerHTML=state.barcodeSvg;
  $("barcodeMeta").textContent="";
  $("genInfo").textContent=`${po.namaCustomer} • ${displayRegisterCode(po)}`;
  $("barcodePrint").disabled=false;
  $("barcodeSave").disabled=false;
}
function printBarcode(){
  if(!state.barcodeSvg){toast("Generate barcode terlebih dahulu.","warn");return;}
  const iframe=document.createElement("iframe");
  iframe.setAttribute("aria-hidden","true");
  iframe.style.position="fixed";
  iframe.style.left="-10000px";
  iframe.style.top="0";
  iframe.style.width="1px";
  iframe.style.height="1px";
  iframe.style.border="0";
  document.body.appendChild(iframe);
  const doc=iframe.contentDocument || iframe.contentWindow.document;
  const cleanup=()=>{setTimeout(()=>iframe.remove(),1200)};
  try{
    doc.open();
    doc.write(`<!doctype html><html lang="id"><head><meta charset="utf-8"><title>Barcode ${escapeHTML(state.barcodePO)}</title><style>@page{size:A4 portrait;margin:15mm}html,body{margin:0;padding:0;background:#fff;color:#000;font-family:Arial,sans-serif}.wrap{display:flex;flex-direction:column;align-items:center;justify-content:flex-start;padding:15mm 10mm;text-align:center}.wrap svg{display:block;width:50%;max-width:90mm;height:auto}.code{display:none}</style></head><body><div class="wrap">${state.barcodeSvg}<div class="code">${escapeHTML(state.barcodePO)}</div></div></body></html>`);
    doc.close();
    setTimeout(()=>{
      try{iframe.contentWindow.focus();iframe.contentWindow.print();cleanup();}
      catch(err){console.error(err);cleanup();toast("Gagal membuka dialog cetak barcode.","error");}
    },300);
  }catch(err){
    console.error(err);
    cleanup();
    toast("Gagal menyiapkan barcode untuk dicetak.","error");
  }
}
function saveBarcode(){if(!state.barcodeSvg)return;const blob=new Blob([state.barcodeSvg],{type:"image/svg+xml;charset=utf-8"});const url=URL.createObjectURL(blob);const a=document.createElement("a");a.href=url;a.download=`Barcode_${state.barcodePO.replace(/[^A-Z0-9_-]/gi,"_")}.svg`;document.body.appendChild(a);a.click();a.remove();setTimeout(()=>URL.revokeObjectURL(url),1000);}

async function loadAudit(){const el=$("auditTable");el.innerHTML='<div class="empty">⏳ Memuat audit log...</div>';try{const snap=await getDocs(query(collection(db,"audit_logs"),orderBy("timestamp","desc")));const rows=snap.docs.slice(0,200).map(d=>({id:d.id,...d.data()}));if(!rows.length){el.innerHTML='<div class="empty">Belum ada audit log.</div>';return}let html='<table><thead><tr><th>Waktu</th><th>User</th><th>Aksi</th><th>No PO</th><th>Register</th><th>Keterangan</th></tr></thead><tbody>';rows.forEach(r=>{const t=r.timestamp?.toDate?r.timestamp.toDate().toLocaleString("id-ID"):"-";html+=`<tr><td>${escapeHTML(t)}</td><td>${escapeHTML(r.userName||"-")}</td><td><strong>${escapeHTML(r.action||"-")}</strong></td><td>${escapeHTML(r.targetPO||"-")}</td><td>${escapeHTML(r.registerCode||"-")}</td><td>${escapeHTML(r.details||"-")}</td></tr>`});html+='</tbody></table>';el.innerHTML=html}catch(err){console.error(err);el.innerHTML='<div class="empty">Gagal memuat audit log.</div>';}}
async function loadUsers(){const el=$("usersTable");el.innerHTML='<div class="empty">⏳ Memuat pengguna...</div>';try{const snap=await getDocs(collection(db,"users"));const rows=snap.docs.map(d=>({id:d.id,...d.data()})).sort((a,b)=>String(a.name||a.email||"").localeCompare(String(b.name||b.email||"")));if(!rows.length){el.innerHTML='<div class="empty">Belum ada profil pengguna.</div>';return}let html='<table><thead><tr><th>Nama</th><th>Email</th><th>Role</th><th>Status</th><th>Aksi</th></tr></thead><tbody>';rows.forEach(r=>{const active=r.active!==false;html+=`<tr><td>${escapeHTML(r.name||"-")}</td><td>${escapeHTML(r.email||"-")}</td><td>${escapeHTML(r.role||"-")}</td><td>${active?'<span class="status status-in">Aktif</span>':'<span class="status status-out">Nonaktif</span>'}</td><td>${r.id===state.user.uid?'<span class="muted">Akun aktif saat ini</span>':`<button type="button" class="btn btn-small ${active?"btn-warning":"btn-success"}" data-user="${r.id}" data-active="${active}">${active?"Nonaktifkan":"Aktifkan"}</button>`}</td></tr>`});html+='</tbody></table>';el.innerHTML=html;el.querySelectorAll("[data-user]").forEach(btn=>btn.addEventListener("click",()=>toggleUser(btn.dataset.user,btn.dataset.active!=="true")));}catch(err){console.error(err);el.innerHTML='<div class="empty">Gagal memuat pengguna.</div>';}}
async function toggleUser(uid,active){if(!isAdmin())return;try{await updateDoc(doc(db,"users",uid),{active,updatedAt:serverTimestamp(),updatedBy:state.user.uid});await writeAudit(active?"ACTIVATE_USER":"DEACTIVATE_USER",uid,"","Mengubah status user");toast(active?"User diaktifkan.":"User dinonaktifkan.","success");await loadUsers();}catch(err){console.error(err);toast("Gagal mengubah status user.","error");}}

function csvEscape(value){
  var s=String(value === null || value === undefined ? '' : value);
  return '"'+s.replace(/"/g,'""')+'"';
}

function reportExportRows(){
  var rows=reportFilteredRows();
  var out=[];
  out.push(['No','Register','Tgl Kirim','Tgl Kembali','No PO','Nama Customer','Keterangan','Status']);
  for(var i=0;i<rows.length;i++){
    var po=rows[i];
    var reg=displayRegisterCode(po);
    out.push([
      i+1,
      reg,
      formatDate(po.tglKirim),
      formatDate(po.tglKembali),
      po.noPO || '',
      po.namaCustomer || '',
      po.keterangan || '',
      getStatus(po)==='SUDAH_KEMBALI' ? 'Sudah Kembali' : 'Belum Kembali'
    ]);
  }
  return out;
}

function exportExcel(){
  if(!isAdmin()){toast('Hanya Admin yang dapat export.','error');return;}
  var rows=reportExportRows();
  var csv=[];
  for(var i=0;i<rows.length;i++){
    var line=[];
    for(var j=0;j<rows[i].length;j++) line.push(csvEscape(rows[i][j]));
    csv.push(line.join(','));
  }
  var blob=new Blob(['\\ufeff'+csv.join('\\r\\n')],{type:'text/csv;charset=utf-8'});
  var url=URL.createObjectURL(blob);
  var a=document.createElement('a');
  a.href=url;
  a.download='Laporan_PO_'+todayISO().replace(/-/g,'')+'.csv';
  document.body.appendChild(a);
  a.click();
  a.remove();
  setTimeout(function(){URL.revokeObjectURL(url);},1000);
  toast('Export berhasil dibuat dalam format CSV yang dapat dibuka di Excel.','success');
}

function printReport(){
  const rows = reportFilteredRows();
  const y = $("reportYear").value;
  const m = $("reportMonth").value;
  const g = $("reportGroup").value;
  const dateField = $("reportDateField")?.value || "send";
  const dateFieldLabel = dateField === "return" ? "Tgl Kembali" : "Tgl Kirim";
  const dateFrom = $("reportFrom").value;
  const dateTo = $("reportTo").value;
  const title = "Laporan PO - " +
    (y === "ALL" ? "Semua Tahun" : y) + " - " +
    (m === "ALL" ? "Semua Bulan" : MONTHS[Number(m) - 1]) + " - " +
    (g === "ALL" ? "Semua Register" : g);

  const returned = rows.filter(po => getStatus(po) === "SUDAH_KEMBALI").length;
  const outstanding = rows.length - returned;

  const bodyRows = rows.length ? rows.map((po, i) => {
    const reg = displayRegisterCode(po);
    const status = getStatus(po);
    const statusColor = status === "SUDAH_KEMBALI" ? "#15803d" : "#b91c1c";
    const statusBg = status === "SUDAH_KEMBALI" ? "#dcfce7" : "#fee2e2";
    return `<tr>
      <td>${i + 1}</td>
      <td style="font-weight:800">${escapeHTML(reg)}</td>
      <td>${escapeHTML(formatDate(po.tglKirim))}</td>
      <td>${escapeHTML(formatDate(po.tglKembali))}</td>
      <td><strong>${escapeHTML(po.noPO || "")}</strong></td>
      <td>${escapeHTML(po.namaCustomer || "")}</td>
      <td>${escapeHTML(po.keterangan || "")}</td>
      <td style="color:${statusColor};background:${statusBg};font-weight:800">${status === "SUDAH_KEMBALI" ? "Sudah Kembali" : "Belum Kembali"}</td>
    </tr>`;
  }).join("") : `<tr><td colspan="8">Tidak ada data</td></tr>`;

  const html = `<!doctype html>
<html lang="id">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>${escapeHTML(title)}</title>
<style>
  *{box-sizing:border-box;-webkit-print-color-adjust:exact!important;print-color-adjust:exact!important}
  html,body{margin:0;padding:0;background:#fff;color:#0f172a;font-family:Arial,Helvetica,sans-serif}
  body{padding:10mm}
  h1{font-size:20px;margin:0 0 5px}
  .meta{font-size:12px;color:#475569;margin-bottom:12px}
  .summary{display:grid;grid-template-columns:repeat(3,1fr);gap:8px;margin-bottom:12px}
  .pill{border:1px solid #94a3b8;border-radius:8px;padding:8px 10px;background:#e2e8f0;text-align:center}
  .pill.green{background:#dcfce7}.pill.red{background:#fee2e2}
  .pill span{display:block;font-size:10px;font-weight:800;text-transform:uppercase;color:#334155}.pill strong{display:block;font-size:18px;margin-top:2px}
  table{width:100%;border-collapse:collapse;table-layout:fixed}
  th,td{border:1px solid #64748b;padding:7px 8px;font-size:10pt;text-align:center;vertical-align:middle;overflow-wrap:anywhere}
  th{background:#bfdbfe!important;color:#0f172a!important;font-weight:800}
  th:nth-child(1){width:4%}th:nth-child(2){width:13%}th:nth-child(3){width:10%}th:nth-child(4){width:10%}
  th:nth-child(5){width:13%}th:nth-child(6){width:16%}th:nth-child(7){width:20%}th:nth-child(8){width:14%}
  thead{display:table-header-group}tbody tr{break-inside:avoid}
  @page{size:A4 landscape;margin:0}
</style>
</head>
<body>
  <h1>${escapeHTML(title)}</h1>
  <div class="meta">Dicetak ${escapeHTML(new Date().toLocaleString("id-ID"))} • Filter tanggal: ${escapeHTML(dateFieldLabel)}${dateFrom||dateTo?` • ${escapeHTML(dateFrom||"awal")} s/d ${escapeHTML(dateTo||"akhir")}`:""}</div>
  <div class="summary">
    <div class="pill"><span>Total PO</span><strong>${rows.length}</strong></div>
    <div class="pill green"><span>Sudah Kembali</span><strong>${returned}</strong></div>
    <div class="pill red"><span>Belum Kembali</span><strong>${outstanding}</strong></div>
  </div>
  <table>
    <thead><tr>
      <th>No</th><th>Register</th><th>Tgl Kirim</th><th>Tgl Kembali</th><th>No PO</th><th>Nama Customer</th><th>Keterangan</th><th>Status</th>
    </tr></thead>
    <tbody>${bodyRows}</tbody>
  </table>
</body>
</html>`;

  let popup = null;
  try { popup = window.open("", "PO_TRACKING_PRINT"); } catch (e) { popup = null; }

  if (popup && !popup.closed) {
    try {
      popup.document.open();
      popup.document.write(html);
      popup.document.close();
      popup.focus();
      setTimeout(() => { try { popup.print(); } catch(e) { console.error(e); } }, 150);
      return;
    } catch (error) {
      console.error(error);
      try { popup.close(); } catch(e) {}
    }
  }

  // Fallback tanpa popup: gunakan iframe terpisah.
  const iframe = document.createElement("iframe");
  iframe.setAttribute("aria-hidden", "true");
  iframe.style.position = "fixed";
  iframe.style.left = "-10000px";
  iframe.style.top = "0";
  iframe.style.width = "1px";
  iframe.style.height = "1px";
  iframe.style.border = "0";
  iframe.style.opacity = "0";
  document.body.appendChild(iframe);

  const cleanup = () => {
    try { iframe.remove(); } catch(e) {}
  };

  try {
    const idoc = iframe.contentDocument || iframe.contentWindow.document;
    idoc.open();
    idoc.write(html);
    idoc.close();
    setTimeout(() => {
      try {
        iframe.contentWindow.focus();
        iframe.contentWindow.print();
        setTimeout(cleanup, 1000);
      } catch(error) {
        console.error(error);
        cleanup();
        toast("Gagal membuka dialog cetak.","error");
      }
    }, 250);
  } catch(error) {
    console.error(error);
    cleanup();
    toast("Gagal menyiapkan laporan untuk dicetak.","error");
  }
}


function getImportErrorCategories(errors){
  const counts=new Map();
  (errors||[]).forEach(item=>{
    const reason=String(item.reason||"Tidak diketahui");
    const parts=reason.split("; ").map(x=>x.trim()).filter(Boolean);
    if(!parts.length)parts.push("Tidak diketahui");
    parts.forEach(part=>counts.set(part,(counts.get(part)||0)+1));
  });
  return [...counts.entries()].sort((a,b)=>b[1]-a[1]||a[0].localeCompare(b[0]));
}
function importDiagnosticsKey(){return `poTrackingImportDiagnostics:${state.user?.uid||"anonymous"}`;}
function saveImportDiagnostics(){
  try{
    const p=state.importPreview;
    if(!p){sessionStorage.removeItem(importDiagnosticsKey());return;}
    sessionStorage.setItem(importDiagnosticsKey(),JSON.stringify(p));
  }catch(err){console.warn("Gagal menyimpan diagnostics import",err);}
}
function loadImportDiagnostics(){
  try{
    const raw=sessionStorage.getItem(importDiagnosticsKey());
    if(!raw)return;
    const p=JSON.parse(raw);
    if(p&&Array.isArray(p.ready)&&Array.isArray(p.duplicates)&&Array.isArray(p.errors)){
      state.importPreview=p;
      renderImportPreview();
    }
  }catch(err){console.warn("Gagal memulihkan diagnostics import",err);}
}
async function exportImportRows(rows,filename,sheetName){
  if(!rows?.length){toast("Tidak ada data untuk diexport.","warn");return;}
  try{
    showLoading(true);
    const XLSX=await loadXLSXLibrary();
    const wb=XLSX.utils.book_new();
    const data=rows.map(r=>({
      Baris:r.rowNumber||"",
      Status:r.reason?"Error":"Siap Import",
      Register:r.registerCode||"",
      "Tgl Kirim":r.tglKirim||"",
      "Tgl Kembali":r.tglKembali||"",
      "No PO":r.noPO||"",
      "Nama Customer":r.namaCustomer||"",
      Keterangan:r.keterangan||"",
      "Alasan/Error":r.reason||""
    }));
    XLSX.utils.book_append_sheet(wb,XLSX.utils.json_to_sheet(data),sheetName||"Data");
    XLSX.writeFile(wb,filename);
    toast(`${rows.length} baris berhasil diexport.`,"success");
  }catch(err){console.error(err);toast("Gagal membuat file Excel diagnostics.","error");}
  finally{showLoading(false)}
}
function renderImportDiagnostics(p){
  const summary=$("importDiagnosticsSummary");
  const errorTable=$("importErrorTable");
  const errorBtn=$("downloadImportErrorsBtn");
  const readyBtn=$("downloadImportReadyBtn");
  if(!summary||!errorTable)return;
  if(!p){summary.innerHTML="";errorTable.innerHTML="";errorTable.classList.add("hidden");if(errorBtn)errorBtn.disabled=true;if(readyBtn)readyBtn.disabled=true;return;}
  const cats=getImportErrorCategories(p.errors);
  let html=`<div><strong>Diagnostik Import</strong> — ${p.errors.length} baris error tersimpan selama sesi ini. ${p.errors.length?"Perbaiki Excel lalu Preview ulang; data error tidak ikut di-import.":"Tidak ada error."}</div>`;
  if(cats.length){html+='<div class="import-diagnostics-grid">';cats.forEach(([name,count])=>{html+=`<div class="import-diagnostic-item"><strong>${count}×</strong> ${escapeHTML(name)}</div>`});html+='</div>';}
  summary.innerHTML=html;
  if(errorBtn)errorBtn.disabled=!p.errors.length;
  if(readyBtn)readyBtn.disabled=!p.ready.length;
  if(p.errors.length){
    let eh=`<div class="import-error-table-head">❌ Daftar semua ${p.errors.length} error import <span class="muted">— seluruh baris error ditampilkan</span></div><table><thead><tr><th>Baris</th><th>No PO</th><th>Register</th><th>Tgl Kirim</th><th>Tgl Kembali</th><th>Customer</th><th>Keterangan</th><th>Alasan Error</th></tr></thead><tbody>`;
    p.errors.forEach(r=>{eh+=`<tr><td>${r.rowNumber}</td><td><strong>${escapeHTML(r.noPO||"")}</strong></td><td>${escapeHTML(r.registerCode||"-")}</td><td>${escapeHTML(formatDate(r.tglKirim))}</td><td>${escapeHTML(formatDate(r.tglKembali))}</td><td>${escapeHTML(r.namaCustomer||"")}</td><td>${escapeHTML(r.keterangan||"")}</td><td class="import-error">${escapeHTML(r.reason||"Tidak diketahui")}</td></tr>`});
    eh+='</tbody></table>';errorTable.innerHTML=eh;errorTable.classList.remove("hidden");
  }else{errorTable.innerHTML="";errorTable.classList.add("hidden");}
}
function importRowCellClass(status){return status==="ready"?"import-ok":status==="duplicate"?"import-skip":"import-error";}
function renderImportPreview(){
  const p=state.importPreview;
  const panel=$("importPreviewPanel");
  if(!p){panel.classList.add("hidden");$("dataImportBtn").disabled=true;return;}
  panel.classList.remove("hidden");
  $("importPreviewFile").textContent=p.fileName||"";
  $("importTotal").textContent=String(p.total);
  $("importReady").textContent=String(p.ready.length);
  $("importDuplicate").textContent=String(p.duplicates.length);
  $("importError").textContent=String(p.errors.length);
  const combined=[...p.ready.map(x=>({...x,_state:"ready"})),...p.duplicates.map(x=>({...x,_state:"duplicate"})),...p.errors.map(x=>({...x,_state:"error"}))].slice(0,250);
  if(!combined.length){$("importPreviewTable").innerHTML='<div class="empty">Tidak ada baris yang dapat diproses.</div>';}else{
    let html='<table><thead><tr><th>Baris</th><th>Status</th><th>Register</th><th>Tgl Kirim</th><th>Tgl Kembali</th><th>No PO</th><th>Customer</th><th>Keterangan</th><th>Catatan</th></tr></thead><tbody>';
    combined.forEach(r=>{const note=r.reason||"Siap";html+=`<tr><td>${r.rowNumber}</td><td class="${importRowCellClass(r._state)}">${r._state==="ready"?"Siap Import":r._state==="duplicate"?"Duplikat": "Error"}</td><td>${escapeHTML(r.registerCode||"-")}</td><td>${escapeHTML(formatDate(r.tglKirim))}</td><td>${escapeHTML(formatDate(r.tglKembali))}</td><td><strong>${escapeHTML(r.noPO||"")}</strong></td><td>${escapeHTML(r.namaCustomer||"")}</td><td>${escapeHTML(r.keterangan||"")}</td><td>${escapeHTML(note)}</td></tr>`});
    html+='</tbody></table>';$("importPreviewTable").innerHTML=html;
  }
  $("dataImportBtn").disabled=!(p.ready.length>0);
  $("dataImportBtn").textContent=p.ready.length?`📥 Import ${p.ready.length} Data Valid`:'📥 Import Data';
  renderImportDiagnostics(p);
  const hint=$("importHistoricalHint");
  if(hint){
    hint.textContent=p.missingSendCount?`${p.missingSendCount} baris tidak memiliki Tgl Kirim. Data tetap dapat di-import tanpa mengisi tanggal palsu; status akan mengikuti Tgl Kembali jika tersedia.`:"";
    hint.classList.toggle("hidden",!p.missingSendCount);
  }
}
async function previewImportFile(){
  if(!isAdmin())return;
  const file=$("dataImportFile").files?.[0];
  if(!file){toast("Pilih file Excel terlebih dahulu.","warn");return;}
  try{
    showLoading(true);
    const XLSX=await loadXLSXLibrary();
    const buffer=await file.arrayBuffer();
    const workbook=XLSX.read(buffer,{type:"array",cellDates:true});
    const sheetName=workbook.SheetNames?.[0];
    if(!sheetName)throw new Error("Sheet Excel tidak ditemukan.");
    const sheet=workbook.Sheets[sheetName];
    const rows=XLSX.utils.sheet_to_json(sheet,{defval:"",raw:true});
    if(!rows.length)throw new Error("File Excel tidak memiliki data.");
    const sample=rows[0];
    const headers={
      send:findHeader(sample,["Tgl Kirim","Tanggal Kirim","Tanggal Pengiriman","Tgl Pengiriman"]),
      ret:findHeader(sample,["Tgl Kembali","Tanggal Kembali","Tanggal Penerimaan"]),
      po:findHeader(sample,["No PO","No. PO","PO","Nomor PO"]),
      customer:findHeader(sample,["Nama Customer","Customer","Nama Pelanggan"]),
      note:findHeader(sample,["Keterangan","Keterangan PO","Catatan"]),
      register:findHeader(sample,["Register","Register Code","Kode Register"])
    };
    if(!headers.po || !headers.customer)throw new Error("Kolom wajib belum lengkap. Minimal: No PO dan Nama Customer. Tgl Kirim boleh kosong untuk data historis.");
    const ready=[],duplicates=[],errors=[],seen=new Set();
    let missingSendCount=0;
    rows.forEach((row,index)=>{
      const rowNumber=index+2;
      const rawSend=row[headers.send];
      const rawReturn=headers.ret?row[headers.ret]:"";
      const tglKirim=parseImportDate(rawSend);
      const tglKembali=parseImportDate(rawReturn);
      const rawSendPresent=rawSend!==null && rawSend!==undefined && String(rawSend).trim()!=="";
      const rawReturnPresent=rawReturn!==null && rawReturn!==undefined && String(rawReturn).trim()!=="";
      const noPO=String(row[headers.po]??"").trim().toUpperCase();
      const customer=String(row[headers.customer]??"").trim();
      const note=headers.note?String(row[headers.note]??"").trim():"";
      const suppliedRegister=headers.register?String(row[headers.register]??"").trim():"";
      const referenceDate=tglKirim||tglKembali||"";
      let regInfo=suppliedRegister?parseRegisterInput(suppliedRegister,referenceDate):null;
      if(!regInfo)regInfo=inferRegisterFromPO(noPO,referenceDate);
      const rowError=[];
      if(rawSendPresent && !tglKirim)rowError.push("Tgl Kirim tidak valid");
      if(rawReturnPresent && !tglKembali)rowError.push("Tgl Kembali tidak valid");
      if(tglKirim && tglKembali && tglKembali<tglKirim)rowError.push("Tgl Kembali sebelum Tgl Kirim");
      if(!noPO)rowError.push("No PO kosong");
      if(!customer)rowError.push("Nama Customer kosong");
      if(suppliedRegister && !parseRegisterInput(suppliedRegister,referenceDate))rowError.push("Format Register tidak dikenali");
      if(regInfo && tglKirim && (regInfo.year!==Number(tglKirim.slice(0,4)) || regInfo.month!==Number(tglKirim.slice(5,7))))rowError.push("Register tidak sesuai dengan bulan/tahun Tgl Kirim");
      if(!regInfo){
        rowError.push(tglKirim||tglKembali?"Register tidak dikenali. Isi Register atau gunakan No PO lengkap yang memuat kode register.":"Periode tidak dapat ditentukan. Isi Register atau gunakan No PO lengkap yang memuat tahun/bulan register.");
      }
      if(rowError.length){errors.push({rowNumber,tglKirim,tglKembali,noPO,namaCustomer:customer,keterangan:note,registerCode:regInfo?.code||"",reason:rowError.join("; ")});return;}
      const unique=normalizePO(noPO);
      if(!unique){errors.push({rowNumber,tglKirim,tglKembali,noPO,namaCustomer:customer,keterangan:note,registerCode:regInfo.code,reason:"No PO tidak dapat dinormalisasi"});return;}
      if(seen.has(unique)){duplicates.push({rowNumber,tglKirim,tglKembali,noPO,namaCustomer:customer,keterangan:note,registerCode:regInfo.code,reason:"Duplikat dalam file import"});return;}
      seen.add(unique);
      const existing=state.allPOs.find(po=>normalizePO(po.noPO)===unique);
      if(existing){duplicates.push({rowNumber,tglKirim,tglKembali,noPO,namaCustomer:customer,keterangan:note,registerCode:regInfo.code,reason:`Sudah ada di aplikasi (${displayRegisterCode(existing)})`});return;}
      // The precedence used for historical rows is: Register/No PO, then Tgl Kirim, then Tgl Kembali.
      const actualPeriodSource=suppliedRegister?"Register":(noPO && inferRegisterFromPO(noPO,referenceDate)?"No PO":(tglKirim?"Tgl Kirim":"Tgl Kembali"));
      if(!tglKirim)missingSendCount++;
      ready.push({rowNumber,tglKirim,tglKembali,noPO,namaCustomer:customer,keterangan:note,registerGroup:regInfo.group,year:regInfo.year,month:regInfo.month,registerCode:regInfo.code,uniqueKey:unique,periodSource:actualPeriodSource});
    });
    state.importPreview={fileName:file.name,total:rows.length,ready,duplicates,errors,missingSendCount,createdAt:Date.now()};
    saveImportDiagnostics();
    renderImportPreview();
    if(errors.length){
      toast(`Preview selesai: ${ready.length} siap, ${duplicates.length} duplikat, ${errors.length} error.`,"warn");
    }else{
      const historyNote=missingSendCount?` (${missingSendCount} tanpa Tgl Kirim, tetap aman)`:"";
      toast(`Preview selesai: ${ready.length} data siap di-import${historyNote}.`,"success");
    }
  }catch(err){console.error(err);toast(err?.message||"Gagal membaca file Excel.","error");state.importPreview=null;saveImportDiagnostics();renderImportPreview();}
  finally{showLoading(false)}
}
async function findExistingByUniqueKey(uniqueKey){
  try{
    const pointerSnap=await getDoc(doc(db,"po_unique",uniqueKey));
    if(pointerSnap.exists()){
      const pdata=pointerSnap.data()||{};
      return {kind:"index",poId:String(pdata.poId||""),data:pdata};
    }
  }catch(err){console.warn("Gagal membaca indeks po_unique",uniqueKey,err);throw err;}
  const matches=state.allPOs.filter(po=>normalizePO(po.noPO)===uniqueKey);
  if(matches.length){
    const canonical=pickCanonicalFromLocal(matches);
    return {kind:"local",poId:canonical.id,data:{noPO:uniqueKey,poId:canonical.id}};
  }
  return null;
}
function pickCanonicalFromLocal(matches){
  return [...matches].sort((a,b)=>{
    const ar=a.tglKembali?1:0, br=b.tglKembali?1:0;
    if(ar!==br)return br-ar;
    const at=a.createdAt?.toMillis?a.createdAt.toMillis():Number.MAX_SAFE_INTEGER;
    const bt=b.createdAt?.toMillis?b.createdAt.toMillis():Number.MAX_SAFE_INTEGER;
    if(at!==bt)return at-bt;
    return String(a.id||"").localeCompare(String(b.id||""));
  })[0];
}
async function importPreviewData(){
  if(!isAdmin())return;
  const p=state.importPreview;
  if(!p || !p.ready.length){toast("Tidak ada data valid untuk di-import.","warn");return;}
  const errorNote=p.errors.length?` ${p.errors.length} baris error akan dilewati.`:"";
  if(!confirm(`Import maksimal ${p.ready.length} data valid dari ${p.fileName}?${errorNote} Data yang sudah ada di Firestore akan otomatis dilewati sehingga aman bila proses diulang.`))return;
  try{
    showLoading(true);
    $("dataImportBtn").disabled=true;
    $("importProgressWrap").classList.remove("hidden");
    const chunkSize=4;
    const total=p.ready.length;
    let processed=0, imported=0, skipped=0, repaired=0;
    for(let start=0;start<total;start+=chunkSize){
      const chunk=p.ready.slice(start,start+chunkSize);
      const batch=writeBatch(db);
      let writes=0;
      const newItems=[];
      for(const item of chunk){
        const found=await findExistingByUniqueKey(item.uniqueKey);
        if(found){
          skipped++;
          if(found.kind==="local" && found.poId){
            batch.set(doc(db,"po_unique",item.uniqueKey),{noPO:item.uniqueKey,poId:found.poId,createdAt:serverTimestamp(),createdBy:state.user.uid});
            writes++; repaired++;
          }
          continue;
        }
        newItems.push(item);
      }
      for(const item of newItems){
        const ref=doc(collection(db,"po_documents"));
        batch.set(ref,{registerGroup:item.registerGroup,year:item.year,month:item.month,registerCode:item.registerCode,tglKirim:item.tglKirim,tglKembali:item.tglKembali||null,noPO:item.noPO,uniqueKey:item.uniqueKey,namaCustomer:item.namaCustomer,keterangan:item.keterangan,status:item.tglKembali?"SUDAH_KEMBALI":"BELUM_KEMBALI",createdAt:serverTimestamp(),createdBy:state.user.uid,updatedAt:serverTimestamp(),updatedBy:state.user.uid});
        batch.set(doc(db,"po_unique",item.uniqueKey),{noPO:item.uniqueKey,poId:ref.id,createdAt:serverTimestamp(),createdBy:state.user.uid});
        writes+=2; imported++;
      }
      if(writes) await batch.commit();
      processed=Math.min(start+chunk.length,total);
      const pct=Math.round(processed/total*100);
      $("importProgressBar").style.width=`${pct}%`;
      $("importProgressText").textContent=`${processed} / ${total} diperiksa (${pct}%) • Baru ${imported} • Dilewati ${skipped}`;
    }
    await writeAudit("IMPORT_EXCEL","","",`Import ${p.fileName}: ${imported} PO baru, ${skipped} sudah ada/dilewati, ${repaired} indeks po_unique diperbaiki${p.errors.length?`, ${p.errors.length} baris error tidak di-import`:""}.`);
    toast(`Import selesai: ${imported} baru, ${skipped} dilewati${p.errors.length?`, ${p.errors.length} error tidak di-import`:""}.`,"success");
    state.importPreview=null;saveImportDiagnostics();$("dataImportFile").value="";$("importProgressWrap").classList.add("hidden");renderImportPreview();await refresh();await loadImportHistory();await scanDuplicateGroups();
  }catch(err){
    console.error(err);
    toast(`Import berhenti: ${err?.code||"error"} — ${err?.message||"gagal menulis Firestore"}. Batch yang sudah committed tetap aman; jalankan ulang untuk melanjutkan tanpa menggandakan PO.` ,"error");
  }finally{showLoading(false);$("dataImportBtn").disabled=!(state.importPreview?.ready?.length>0);}
}
function canonicalScore(po){
  let score=0;
  if(po.tglKembali)score+=100;
  if(po.tglKirim)score+=20;
  if(po.namaCustomer)score+=5;
  if(po.keterangan)score+=1;
  return score;
}
async function scanDuplicateGroups(){
  if(!isAdmin())return [];
  const map=new Map();
  for(const po of state.allPOs){
    const key=normalizePO(po.noPO);
    if(!key)continue;
    if(!map.has(key))map.set(key,[]);
    map.get(key).push(po);
  }
  const groups=[...map.entries()].filter(([,arr])=>arr.length>1).map(([key,items])=>({key,items}));
  groups.sort((a,b)=>b.items.length-a.items.length||a.key.localeCompare(b.key));
  state.duplicateGroups=groups;
  renderDuplicateAudit(groups);
  return groups;
}
function duplicateCreatedTime(po){return po.createdAt?.toMillis?po.createdAt.toMillis():Number.MAX_SAFE_INTEGER;}
function chooseCanonicalForGroup(items){
  return [...items].sort((a,b)=>canonicalScore(b)-canonicalScore(a)||duplicateCreatedTime(a)-duplicateCreatedTime(b)||String(a.id).localeCompare(String(b.id)))[0];
}
function renderDuplicateAudit(groups=state.duplicateGroups||[]){
  const wrap=$("duplicateAuditPanel"); const body=$("duplicateAuditTable");
  if(!wrap||!body)return;
  wrap.classList.remove("hidden");
  if(!groups.length){body.innerHTML='<div class="empty">✅ Tidak ditemukan duplikat No PO berdasarkan data yang saat ini dimuat.</div>';return;}
  let html=`<div class="hint" style="margin-bottom:10px">Ditemukan <strong>${groups.length}</strong> No PO yang memiliki lebih dari satu dokumen. Sistem menyarankan mempertahankan data terbaik (prioritas Tgl Kembali, lalu Tgl Kirim, lalu waktu dibuat). Tidak ada penghapusan otomatis sebelum Anda menekan tombol tindakan.</div><table><thead><tr><th>No PO</th><th>Jumlah</th><th>Dokumen</th><th>Rekomendasi</th><th>Aksi</th></tr></thead><tbody>`;
  groups.forEach(g=>{
    const canonical=chooseCanonicalForGroup(g.items);
    const docs=g.items.map(po=>`${escapeHTML(po.id)}${po.id===canonical.id?' ⭐':''}`).join('<br>');
    html+=`<tr><td><strong>${escapeHTML(g.key)}</strong></td><td>${g.items.length}</td><td>${docs}</td><td>${escapeHTML(canonical.id)} <span class="muted">(dipertahankan)</span></td><td><button type="button" class="btn btn-danger btn-small" data-dedupe-key="${escapeHTML(g.key)}">Pertahankan & Hapus Duplikat</button></td></tr>`;
  });
  html+='</tbody></table>';body.innerHTML=html;
  body.querySelectorAll('[data-dedupe-key]').forEach(btn=>btn.addEventListener('click',()=>cleanupDuplicateGroup(btn.dataset.dedupeKey)));
}
async function cleanupDuplicateGroup(uniqueKey){
  if(!isAdmin())return;
  const group=(state.duplicateGroups||[]).find(g=>g.key===uniqueKey); if(!group)return;
  const canonical=chooseCanonicalForGroup(group.items);
  const duplicates=group.items.filter(po=>po.id!==canonical.id);
  if(!duplicates.length){toast("Tidak ada dokumen duplikat yang perlu dihapus.","warn");return;}
  if(!confirm(`No PO ${uniqueKey} memiliki ${group.items.length} dokumen. Pertahankan ${canonical.id} dan hapus ${duplicates.length} duplikat?`))return;
  try{
    showLoading(true);
    const batch=writeBatch(db);
    batch.set(doc(db,"po_unique",uniqueKey),{noPO:uniqueKey,poId:canonical.id,updatedAt:serverTimestamp(),updatedBy:state.user.uid});
    for(const po of duplicates) batch.delete(doc(db,"po_documents",po.id));
    await batch.commit();
    await writeAudit("DEDUP_PO",uniqueKey,canonical.registerCode||"",`Membersihkan ${duplicates.length} duplikat; dipertahankan ${canonical.id}.`);
    toast(`${uniqueKey}: ${duplicates.length} duplikat dihapus.`,"success");
    await refresh(); await scanDuplicateGroups();
  }catch(err){console.error(err);toast(`Gagal membersihkan ${uniqueKey}: ${err?.code||"error"} — ${err?.message||""}`,"error");}
  finally{showLoading(false)}
}

function exportRowsAsObjects(rows){
  return rows.map((po,i)=>({No:i+1,Register:displayRegisterCode(po),"Tgl Kirim":formatDate(po.tglKirim),"Tgl Kembali":formatDate(po.tglKembali),"No PO":po.noPO||"","Nama Customer":po.namaCustomer||"",Keterangan:po.keterangan||"",Status:getStatus(po)==="SUDAH_KEMBALI"?"Sudah Kembali":"Belum Kembali"}));
}
async function exportAllDataExcel(){
  if(!isAdmin())return;
  try{
    showLoading(true);const XLSX=await loadXLSXLibrary();const wb=XLSX.utils.book_new();const data=exportRowsAsObjects([...state.allPOs].sort(comparePO));
    const ws=XLSX.utils.json_to_sheet(data);XLSX.utils.book_append_sheet(wb,ws,"PO Data");
    const meta=[{Keterangan:"Export seluruh data PO",Tanggal:todayISO(),Jumlah:data.length}];XLSX.utils.book_append_sheet(wb,XLSX.utils.json_to_sheet(meta),"Info");
    XLSX.writeFile(wb,`PO_Tracking_All_${todayISO().replace(/-/g,"")}.xlsx`);toast(`Export ${data.length} PO berhasil.`,"success");
  }catch(err){console.error(err);toast("Gagal membuat file Excel. Pastikan koneksi internet tersedia untuk library Excel.","error");}
  finally{showLoading(false)}
}
async function downloadImportTemplate(){
  try{
    showLoading(true);const XLSX=await loadXLSXLibrary();const wb=XLSX.utils.book_new();
    const rows=[
      {"Tgl Kirim":"01/01/2021","Tgl Kembali":"05/01/2021","No PO":"BZOA21A001","Nama Customer":"Contoh Customer","Keterangan":"Contoh data","Register":"BZOA21-A"},
      {"Tgl Kirim":"","Tgl Kembali":"15/02/2021","No PO":"BZOB21B002","Nama Customer":"Contoh Histori Tanpa Tgl Kirim","Keterangan":"Tgl Kirim tidak tersedia di arsip","Register":"BZOB21-B"}
    ];
    const ws=XLSX.utils.json_to_sheet(rows);XLSX.utils.book_append_sheet(wb,ws,"Template Import");
    const info=[{"Kolom":"Tgl Kirim","Wajib":"Tidak untuk histori","Keterangan":"Boleh kosong untuk data lama. Jika kosong, jangan diisi tanggal palsu."},{"Kolom":"Tgl Kembali","Wajib":"Tidak","Keterangan":"Boleh diisi walaupun Tgl Kirim kosong. Jika ada, status otomatis Sudah Kembali."},{"Kolom":"No PO","Wajib":"Ya","Keterangan":"No PO asli, jangan diubah."},{"Kolom":"Nama Customer","Wajib":"Ya","Keterangan":"Nama customer."},{"Kolom":"Keterangan","Wajib":"Tidak","Keterangan":"Catatan PO."},{"Kolom":"Register","Wajib":"Disarankan","Keterangan":"Contoh BZOA21-A atau BZOA21A. Jika Tgl Kirim kosong, Register sangat membantu menentukan periode."}];XLSX.utils.book_append_sheet(wb,XLSX.utils.json_to_sheet(info),"Petunjuk");
    XLSX.writeFile(wb,"Template_Import_PO.xlsx");toast("Template Excel berhasil dibuat.","success");
  }catch(err){console.error(err);toast("Gagal membuat template Excel.","error");}
  finally{showLoading(false)}
}
async function loadImportHistory(){
  if(!isAdmin())return;
  const el=$("importHistoryTable");if(!el)return;el.innerHTML='<div class="empty">⏳ Memuat riwayat import...</div>';
  try{
    const snap=await getDocs(query(collection(db,"audit_logs"),orderBy("timestamp","desc")));
    const rows=snap.docs.map(d=>({id:d.id,...d.data()})).filter(r=>r.action==="IMPORT_EXCEL").slice(0,50);
    if(!rows.length){el.innerHTML='<div class="empty">Belum ada riwayat import Excel.</div>';return;}
    let html='<table><thead><tr><th>Waktu</th><th>User</th><th>Aksi</th><th>Keterangan</th></tr></thead><tbody>';
    rows.forEach(r=>{const t=r.timestamp?.toDate?r.timestamp.toDate().toLocaleString("id-ID"):"-";html+=`<tr><td>${escapeHTML(t)}</td><td>${escapeHTML(r.userName||"-")}</td><td><strong>${escapeHTML(r.action)}</strong></td><td>${escapeHTML(r.details||"-")}</td></tr>`});
    html+='</tbody></table>';el.innerHTML=html;
  }catch(err){console.error(err);el.innerHTML='<div class="empty">Gagal memuat riwayat import.</div>';}
}

async function initUser(user){state.user=user;const snap=await getDoc(doc(db,"users",user.uid));if(!snap.exists()){await signOut(auth);toast("Akun belum memiliki profil users di Firestore.","error");return}state.profile=snap.data();if(state.profile.active===false){await signOut(auth);toast("Akun Anda dinonaktifkan.","error");return}loadSavedPeriod();$("loginScreen").classList.add("hidden");$("appScreen").classList.remove("hidden");$("topUserName").textContent=state.profile.name||user.email;$("topUserRole").textContent=state.profile.role||"staff";document.querySelectorAll(".admin-only").forEach(el=>el.classList.toggle("hidden",!isAdmin()));populatePeriods();populateReportFilters();renderRegisterNav();if(isAdmin())loadImportDiagnostics();await refresh();if(isAdmin())await scanDuplicateGroups();navigate("dashboard");}

$("loginForm").addEventListener("submit",async e=>{e.preventDefault();$("loginBtn").disabled=true;$("loginBtn").textContent="⏳ Login...";try{await signInWithEmailAndPassword(auth,$("loginEmail").value.trim(),$("loginPassword").value)}catch(err){console.error(err);toast(err.code==="auth/invalid-credential"?"Email atau password salah.":"Login gagal. Cek email, password, dan konfigurasi Firebase.","error");}finally{$("loginBtn").disabled=false;$("loginBtn").textContent="🔐 Login";}});
$("logoutBtn").addEventListener("click",()=>signOut(auth));
$("mobileMenuBtn").addEventListener("click",()=>$("sidebar").classList.toggle("open"));
document.querySelectorAll(".nav-btn[data-page]").forEach(btn=>btn.addEventListener("click",()=>navigate(btn.dataset.page)));
$("yearSelect").addEventListener("change",async()=>{state.year=Number($("yearSelect").value);state.registerGroup=null;saveCurrentPeriod();populateReportFilters();await refresh();});
$("monthSelect").addEventListener("change",async()=>{state.month=Number($("monthSelect").value);state.registerGroup=null;saveCurrentPeriod();populateReportFilters();await refresh();});
$("dashboardAdd").addEventListener("click",()=>openPOForm());$("manualQuick").addEventListener("click",()=>openPOForm());$("voiceQuick").addEventListener("click",openVoice);$("scanQuick").addEventListener("click",()=>{stopCamera();$("scanModal").classList.remove("hidden");});
$("registerBack").addEventListener("click",()=>navigate("dashboard"));$("registerAdd").addEventListener("click",()=>openPOForm());
["regSearch","regFrom","regTo","regStatus"].forEach(id=>$(id).addEventListener("input",renderCurrentRegister));$("regReset").addEventListener("click",()=>{["regSearch","regFrom","regTo"].forEach(id=>$(id).value="");$("regStatus").value="ALL";renderCurrentRegister()});
["allSearch","allFrom","allTo","allStatus"].forEach(id=>$(id).addEventListener("input",renderAllPO));$("allReset").addEventListener("click",()=>{["allSearch","allFrom","allTo"].forEach(id=>$(id).value="");$("allStatus").value="ALL";renderAllPO()});
["reportYear","reportMonth","reportGroup","reportStatus","reportFrom","reportTo","reportSearch","reportDateField"].forEach(id=>$(id).addEventListener("input",renderReport));$("reportReset").addEventListener("click",()=>{populateReportFilters();$("reportFrom").value="";$("reportTo").value="";$("reportSearch").value="";$("reportDateField").value="send";renderReport()});$("reportPrint").addEventListener("click",printReport);$("reportExport").addEventListener("click",exportExcel);
$("poClose").addEventListener("click",()=>closeModal("poModal"));$("poCancel").addEventListener("click",()=>closeModal("poModal"));$("poForm").addEventListener("submit",savePO);
$("quickClose").addEventListener("click",()=>closeModal("quickModal"));$("quickCancel").addEventListener("click",()=>closeModal("quickModal"));$("quickRun").addEventListener("click",runQuick);$("quickPo").addEventListener("keydown",e=>{if(e.key==="Enter"){e.preventDefault();runQuick()}});
$("confirmClose").addEventListener("click",()=>closeModal("confirmModal"));$("confirmCancel").addEventListener("click",()=>closeModal("confirmModal"));$("confirmRun").addEventListener("click",runConfirm);
$("voiceClose").addEventListener("click",()=>{stopVoice();closeModal("voiceModal")});$("voiceCancel").addEventListener("click",()=>{stopVoice();closeModal("voiceModal")});$("voiceStart").addEventListener("click",startVoice);$("voiceAnalyze").addEventListener("click",processTypedVoice);$("voiceInput").addEventListener("keydown",e=>{if(e.key==="Enter"){e.preventDefault();processTypedVoice()}});$("voiceProcess").addEventListener("click",()=>{if(state.pendingVoice){state.pendingConfirm={type:"batch",items:state.pendingVoice.results};closeModal("voiceModal");$("confirmTitle").textContent=`Konfirmasi Batch ${state.pendingVoice.results.length} PO`;$("confirmBody").innerHTML=state.pendingVoice.results.map((r,i)=>`<div style="padding:8px 0;border-bottom:1px solid var(--border)"><strong>${i+1}. ${escapeHTML(r.po.noPO)}</strong><div class="hint">${escapeHTML(r.po.namaCustomer)} • ${r.action}</div></div>`).join("");$("confirmModal").classList.remove("hidden")}});
$("voiceTabMic").addEventListener("click",()=>{$("voiceTabMic").classList.add("active");$("voiceTabType").classList.remove("active");$("voiceMicPane").classList.remove("hidden");$("voiceTypePane").classList.add("hidden")});$("voiceTabType").addEventListener("click",()=>{$("voiceTabType").classList.add("active");$("voiceTabMic").classList.remove("active");$("voiceMicPane").classList.add("hidden");$("voiceTypePane").classList.remove("hidden")});
$("cameraStart").addEventListener("click",startCamera);$("cameraStop").addEventListener("click",stopCamera);$("scanClose").addEventListener("click",()=>{stopCamera();closeModal("scanModal")});$("scanCancel").addEventListener("click",()=>{stopCamera();closeModal("scanModal")});
function syncScanFooter(){
  const genActive=!$("genPane").classList.contains("hidden");
  const footerSearch=$("scanSearch");
  if(genActive){footerSearch.classList.add("hidden");}
  else{footerSearch.classList.remove("hidden");footerSearch.textContent="🔎 Cari PO";}
}
$("scanSearch").addEventListener("click",()=>handleScanSearch());$("scanInput").addEventListener("keydown",e=>{if(e.key==="Enter"){e.preventDefault();handleScanSearch()}});
$("scanTab").addEventListener("click",()=>{$("scanTab").classList.add("active");$("genTab").classList.remove("active");$("scanPane").classList.remove("hidden");$("genPane").classList.add("hidden");syncScanFooter();});$("genTab").addEventListener("click",()=>{$("genTab").classList.add("active");$("scanTab").classList.remove("active");$("genPane").classList.remove("hidden");$("scanPane").classList.add("hidden");syncScanFooter();});$("genFind").addEventListener("click",()=>{const po=findGeneratePO();if(po){$("genInfo").textContent=`Ditemukan: ${po.namaCustomer} • ${displayRegisterCode(po)}`;$("barcodeMeta").textContent=po.noPO;}});$("genMake").addEventListener("click",generateBarcode);$("barcodePrint").addEventListener("click",printBarcode);$("barcodeSave").addEventListener("click",saveBarcode);$("genPoInput").addEventListener("keydown",e=>{if(e.key==="Enter"){e.preventDefault();generateBarcode()}});syncScanFooter();
["poModal","quickModal","confirmModal","voiceModal","scanModal"].forEach(id=>$(id).addEventListener("click",e=>{if(e.target.id===id){if(id==="scanModal")stopCamera();if(id==="voiceModal")stopVoice();closeModal(id)}}));
window.addEventListener("beforeunload",()=>{stopCamera();stopVoice()});
document.addEventListener("click",e=>{
  const sidebar=$("sidebar");
  const menu=$("mobileMenuBtn");
  if(window.innerWidth<=780 && sidebar.classList.contains("open") && !sidebar.contains(e.target) && !menu.contains(e.target)) sidebar.classList.remove("open");
});
document.addEventListener("visibilitychange",()=>{ if(document.hidden) { stopCamera(); stopVoice(); } });



$("dataPreviewBtn").addEventListener("click",previewImportFile);
$("dataImportBtn").addEventListener("click",importPreviewData);
$("dataExportBtn").addEventListener("click",exportAllDataExcel);
$("dataTemplateBtn").addEventListener("click",downloadImportTemplate);
$("dataImportFile").addEventListener("change",()=>{state.importPreview=null;saveImportDiagnostics();renderImportPreview();$("importPreviewFile").textContent=$("dataImportFile").files?.[0]?.name||"";});

$("downloadImportErrorsBtn").addEventListener("click",()=>{const p=state.importPreview;if(!p)return;const stamp=new Date().toISOString().slice(0,10).replace(/-/g,"");exportImportRows(p.errors,`PO_Import_Errors_${stamp}.xlsx`,"Errors")});
$("downloadImportReadyBtn").addEventListener("click",()=>{const p=state.importPreview;if(!p)return;const stamp=new Date().toISOString().slice(0,10).replace(/-/g,"");exportImportRows(p.ready,`PO_Import_Ready_${stamp}.xlsx`,"Ready Import")});
$("duplicateScanBtn").addEventListener("click",()=>scanDuplicateGroups());

onAuthStateChanged(auth,user=>{if(user)initUser(user).catch(err=>{console.error(err);toast("Gagal memuat profil pengguna.","error")});else{$("loginScreen").classList.remove("hidden");$("appScreen").classList.add("hidden");}});
