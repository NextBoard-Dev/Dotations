// SUiVI DE CHANTIERS - état embarqu forc (pas de localStorage)

// Projet : Rénovation Bureau Pastorale + 1 tâche date



const el = (id)=>document.getElementById(id);
function setInputValue(id, value){ const n = el(id); if(n) n.value = value; return n; }
function ensureProjectHeaderNodes(){
  const projectView = el("viewProject") || document.body;
  if(!el("projectTitle")){
    const h = document.createElement("h2");
    h.id = "projectTitle";
    h.style.display = "none";
    projectView.appendChild(h);
  }
  if(!el("projectSub")){
    const p = document.createElement("p");
    p.id = "projectSub";
    p.style.display = "none";
    projectView.appendChild(p);
  }
}

const STORAGE_KEY = "suivi_chantiers_state_v1";

// Monitoring simple: capture des erreurs + message utilisateur lisible.
const APP_ERROR_BUFFER_MAX = 50;
let __lastErrorSig = "";
let __lastErrorTs = 0;
window.__appErrors = window.__appErrors || [];

function _formatErrorMessage(errLike){
  if(!errLike) return "Erreur inconnue";
  if(typeof errLike === "string") return errLike;
  if(errLike.message) return String(errLike.message);
  try{ return JSON.stringify(errLike); }catch(e){ return String(errLike); }
}

function showAppErrorBanner(message){
  const banner = el("appErrorBanner");
  const text = el("appErrorText");
  if(!banner || !text) return;
  text.textContent = message || "Une erreur technique est survenue.";
  banner.classList.remove("hidden");
}

function reportAppError(errLike, context="runtime"){
  const msg = _formatErrorMessage(errLike);
  const sig = `${context}|${msg}`;
  const now = Date.now();
  // anti-spam: ignore la meme erreur pendant 2 secondes
  if(sig === __lastErrorSig && (now - __lastErrorTs) < 2000) return;
  __lastErrorSig = sig;
  __lastErrorTs = now;

  const item = { ts: new Date().toISOString(), context, message: msg };
  window.__appErrors.push(item);
  if(window.__appErrors.length > APP_ERROR_BUFFER_MAX){
    window.__appErrors.splice(0, window.__appErrors.length - APP_ERROR_BUFFER_MAX);
  }

  console.error("[app-error]", context, msg, errLike);
  showAppErrorBanner(`Une erreur est survenue (${context}). ${msg}`);
}

function softCatch(errLike, context="soft"){
  try{
    const msg = _formatErrorMessage(errLike);
    console.warn(`[soft-error] ${context}: ${msg}`);
  }catch(_e){
    // no-op volontaire: ne jamais casser l'UI pour une remontée d'erreur
  }
}
window.reportAppError = reportAppError;
window.addEventListener("error", (ev)=>{
  const msg = String(ev?.message || "").trim();
  const src = String(ev?.filename || "").trim();
  // "Script error." vient souvent d'un script cross-origin sans détail exploitable.
  // On l'ignore pour ne pas polluer l'UI avec une fausse alerte bloquante.
  if(msg === "Script error." && !ev?.error && !src){
    console.warn("[window.error] Script error cross-origin ignorée");
    return;
  }
  const details = ev?.error || `${msg}${src ? ` @ ${src}:${ev?.lineno||0}:${ev?.colno||0}` : ""}` || "Erreur JavaScript";
  reportAppError(details, "window.error");
});
window.addEventListener("unhandledrejection", (ev)=>{
  reportAppError(ev?.reason || "Promesse rejetee", "promise");
});

document.addEventListener("DOMContentLoaded", ()=>{
  el("appErrorClose")?.addEventListener("click", ()=> el("appErrorBanner")?.classList.add("hidden"));
  el("appErrorReload")?.addEventListener("click", ()=> window.location.reload());
});



/* =========================================================

   SUPABASE GREFFE MINIMALE (NE TOUCHE PAS A L'UI)

   - Pas de module

   - Pas de refactor

   - Supabase est appele APRES saveState()

   - Chargement Supabase APRES premier rendu UI

   Fonctions autorisees (globals) :

     window.supabaseLogin(email, password)

     window.saveAppStateToSupabase(state)

     window.loadAppStateFromSupabase()

========================================================= */



// ---- CONFIG (TES VALEURS) ----

const SUPABASE_URL  = "https://uioqchhbakcvemknqikh.supabase.co";

const SUPABASE_KEY  = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6InVpb3FjaGhiYWtjdmVta25xaWtoIiwicm9sZSI6ImFub24iLCJpYXQiOjE3Njk3NjA4MTUsImV4cCI6MjA4NTMzNjgxNX0.W345e_uwKLGaFcP9KAZq0kNECBUSFluh2ErgHaHeO5w";

const SUPABASE_TABLE = "app_states";
const SUPABASE_USERS_TABLE = "dashboard_users";
const SUPABASE_LOGINS_TABLE = "dashboard_logins";
const SUPABASE_SESSIONS_TABLE = "dashboard_sessions";


// Auto-login (pour ne PAS utiliser la console)

const SUPABASE_AUTO_EMAIL = "sebastien_duc@outlook.fr";

const SUPABASE_AUTO_PASSWORD = "Mililum@tt45";



// ---- client ----

let _sb = null;

function _getSupabaseClient(){

  try{

    if(_sb) return _sb;

    if(!window.supabase || !window.supabase.createClient){

      console.warn("Supabase CDN non charge");

      return null;

    }

    _sb = window.supabase.createClient(SUPABASE_URL, SUPABASE_KEY);

    return _sb;

  }catch(e){

    console.warn("Supabase init failed", e);

    return null;

  }

}



// ---- helper : session ----

async function _ensureSession(){

  const sb = _getSupabaseClient();

  if(!sb) return null;



  try{

    const s = await sb.auth.getSession();

    if(s && s.data && s.data.session) return s.data.session;



    // pas de session -> auto login (si creds presentes)

    if(SUPABASE_AUTO_EMAIL && SUPABASE_AUTO_PASSWORD){

      const res = await sb.auth.signInWithPassword({

        email: SUPABASE_AUTO_EMAIL,

        password: SUPABASE_AUTO_PASSWORD

      });

      if(res && res.data && res.data.session) return res.data.session;

    }

  }catch(e){

    console.warn("Supabase session failed", e);

  }

  return null;

}



/* ===============================

   API GLOBALE AUTORISEE

================================ */

window.supabaseLogin = async function(email, password){

  const sb = _getSupabaseClient();

  if(!sb) return false;

  try{

    const res = await sb.auth.signInWithPassword({ email, password });

    return !!(res && res.data && res.data.session);

  }catch(e){

    console.warn("supabaseLogin failed", e);

    return false;

  }

};



window.saveAppStateToSupabase = async function(stateObj){
  const sb = _getSupabaseClient();
  if(!sb) return false;


  const session = await _ensureSession();

  if(!session || !session.user) return false;



  try{

    const payload = {

      user_id: session.user.id,

      state_json: stateObj,

      updated_at: new Date().toISOString()

    };

    const { error } = await sb.from(SUPABASE_TABLE).upsert(payload, { onConflict: "user_id" });

    if(error){ console.warn("Supabase upsert error", error); return false; }

    return true;

  }catch(e){

    console.warn("saveAppStateToSupabase failed", e);
    console.error("[SUPABASE ERROR]", e);

    const statusEl = document.getElementById("saveStatusMessage") || document.getElementById("saveToastDetail");
    if (statusEl) {
      statusEl.textContent = "Erreur de sauvegarde cloud.";
      statusEl.style.color = "red";
    }

    setTimeout(() => {
      const toast = document.getElementById("saveToast");
      if(toast) toast.classList.remove("show");
    }, 4000);

    return false;

  }

};

// ---- users (simple, sans RLS) ----
async function saveUsersToSupabase(users){
  const sb = _getSupabaseClient();
  if(!sb) return false;
  const session = await _ensureSession();
  if(!session || !session.user) return false;
  try{
    const payload = {
      user_id: session.user.id,
      users_json: users,
      updated_at: new Date().toISOString()
    };
    const { error } = await sb.from(SUPABASE_USERS_TABLE).upsert(payload, { onConflict: "user_id" });
    if(error){ console.warn("Supabase users upsert error", error); return false; }
    return true;
  }catch(e){
    console.warn("saveUsersToSupabase failed", e);
    return false;
  }
}

// expose for login overlay (index.html)
window.saveUsersToSupabase = saveUsersToSupabase;

async function logLoginToSupabase(payload){
  const sb = _getSupabaseClient();
  if(!sb) return false;
  const session = await _ensureSession();
  const userId = session?.user?.id || "anon";
  try{
    const row = {
      user_id: userId,
      email: payload?.email || "",
      name: payload?.name || "",
      role: payload?.role || "user",
      ts: payload?.ts || new Date().toISOString()
    };
    const { error } = await sb.from(SUPABASE_LOGINS_TABLE).insert(row);
    if(error){
      console.warn("Supabase logins insert error", error);
      try{ localStorage.setItem("login_log_last_error", error.message || "insert_failed"); }catch(e){ softCatch(e); }
      return false;
    }
    try{ localStorage.removeItem("login_log_last_error"); }catch(e){ softCatch(e); }
    return true;
  }catch(e){
    console.warn("logLoginToSupabase failed", e);
    try{ localStorage.setItem("login_log_last_error", e?.message || "insert_failed"); }catch(err){}
    return false;
  }
}
window.logUserLogin = async function(payload){
  try{
    return await logLoginToSupabase(payload);
  }catch(e){
    return false;
  }
};

async function sha256Hex(str){
  const enc = new TextEncoder().encode(str || "");
  const buf = await crypto.subtle.digest("SHA-256", enc);
  return [...new Uint8Array(buf)].map(b=>b.toString(16).padStart(2,"0")).join("");
}

async function createSessionToken(token, payload, ttlDays=30){
  const sb = _getSupabaseClient();
  if(!sb) return false;
  const tokenHash = await sha256Hex(token);
  const expires = new Date();
  expires.setDate(expires.getDate() + (ttlDays || 30));
  try{
    const row = {
      token_hash: tokenHash,
      email: payload?.email || "",
      name: payload?.name || "",
      role: payload?.role || "user",
      expires_at: expires.toISOString(),
      created_at: new Date().toISOString()
    };
    const { error } = await sb.from(SUPABASE_SESSIONS_TABLE).insert(row);
    if(error){ console.warn("Supabase sessions insert error", error); return false; }
    return true;
  }catch(e){
    console.warn("createSessionToken failed", e);
    return false;
  }
}

async function validateSessionToken(token, renewDays=30){
  const sb = _getSupabaseClient();
  if(!sb) return null;
  try{
    const tokenHash = await sha256Hex(token);
    const { data, error } = await sb
      .from(SUPABASE_SESSIONS_TABLE)
      .select("email,name,role,expires_at")
      .eq("token_hash", tokenHash)
      .maybeSingle();
    if(error){ console.warn("Supabase sessions select error", error); return null; }
    if(!data) return null;
    const exp = data.expires_at ? new Date(data.expires_at) : null;
    if(!exp || isNaN(exp) || exp < new Date()) return null;
    if(renewDays){
      const next = new Date();
      next.setDate(next.getDate() + renewDays);
      await sb.from(SUPABASE_SESSIONS_TABLE)
        .update({ expires_at: next.toISOString() })
        .eq("token_hash", tokenHash);
    }
    return { email: data.email || "", name: data.name || "", role: data.role || "user" };
  }catch(e){
    console.warn("validateSessionToken failed", e);
    return null;
  }
}

window.createSessionToken = createSessionToken;
window.validateSessionToken = validateSessionToken;

async function cleanupExpiredSessions(){
  const sb = _getSupabaseClient();
  if(!sb) return false;
  try{
    const nowIso = new Date().toISOString();
    const { error } = await sb
      .from(SUPABASE_SESSIONS_TABLE)
      .delete()
      .lt("expires_at", nowIso);
    if(error){ console.warn("Supabase sessions cleanup error", error); return false; }
    return true;
  }catch(e){
    console.warn("cleanupExpiredSessions failed", e);
    return false;
  }
}
window.cleanupExpiredSessions = cleanupExpiredSessions;

async function loadLoginsFromSupabase(startISO, endISO){
  const sb = _getSupabaseClient();
  if(!sb) return [];
  const session = await _ensureSession();
  const userId = session?.user?.id || null;
  try{
    let q = sb.from(SUPABASE_LOGINS_TABLE)
      .select("email,name,role,ts")
      .order("ts", { ascending: true });
    if(userId) q = q.in("user_id", [userId, "anon"]);
    if(startISO) q = q.gte("ts", startISO);
    if(endISO) q = q.lte("ts", endISO);
    const { data, error } = await q;
    if(error){ console.warn("Supabase logins select error", error); return []; }
    return data || [];
  }catch(e){
    console.warn("loadLoginsFromSupabase failed", e);
    return [];
  }
}

async function loadUsersFromSupabase(force=false){
  const sb = _getSupabaseClient();
  if(!sb) return false;
  const session = await _ensureSession();
  if(!session || !session.user) return false;
  try{
    if(!force){
      const localUsers = loadUsers();
      if(localUsers && localUsers.length > 0 && !isHostedGithubPages()){
        return false;
      }
    }
    const { data, error } = await sb
      .from(SUPABASE_USERS_TABLE)
      .select("users_json, updated_at")
      .eq("user_id", session.user.id)
      .maybeSingle();
    if(error){ console.warn("Supabase users select error", error); return false; }
    if(!data || !data.users_json) return false;
    const normalized = (data.users_json || []).map(u=>{
      if(!u || typeof u !== "object") return u;
      if(!u.theme) u.theme = "sable";
      if(!u.id) u.id = uid();
      return u;
    });
    saveUsers(normalized);
    if(typeof window.populateLoginUsers === "function") window.populateLoginUsers();
    return true;
  }catch(e){
    console.warn("loadUsersFromSupabase failed", e);
    return false;
  }
}
window.forceLoadUsersFromSupabase = async function(){
  try{
    return await loadUsersFromSupabase(true);
  }catch(e){
    return false;
  }
};


window.loadAppStateFromSupabase = async function(){

  const sb = _getSupabaseClient();

  if(!sb) return false;



  const session = await _ensureSession();

  if(!session || !session.user) return false;



  try{

    const { data, error } = await sb

      .from(SUPABASE_TABLE)

      .select("state_json, updated_at")

      .eq("user_id", session.user.id)

      .maybeSingle();



    if(error){ console.warn("Supabase select error", error); return false; }

    if(!data || !data.state_json) return false;



    // IMPORTANT : on remplace UNIQUEMENT l'eétat global, puis on rend

    state = normalizeState(data.state_json);

    renderAll();

    clearDirty();

    return true;

  }catch(e){

    console.warn("loadAppStateFromSupabase failed", e);

    return false;

  }

};



// ---- auto-load apres 1er rendu UI ----

let _supabaseAutoloadScheduled = false;
function _scheduleSupabaseAutoLoad(){
  if(_supabaseAutoloadScheduled) return;
  _supabaseAutoloadScheduled = true;

  // pas d'await au chargement initial : on laisse l'UI se rendre d'abord
  setTimeout(function(){
    try{
      window.loadAppStateFromSupabase();
      loadUsersFromSupabase();
    }catch(e){ softCatch(e); }
  }, 120);
}


const uid = ()=> Math.random().toString(16).slice(2,10) + Date.now().toString(16);

const normId = (v)=> (v===undefined || v===null) ? "" : String(v).trim();



let state = null;

let selectedProjectId = null;

let selectedTaskId = null;

let taskOrderMap = {};

let selectedStatusSet = new Set();

let sortMaster = {key:"start", dir:"asc"};

let sortProject = {key:"start", dir:"asc"};

let tabsSortMode = "progress_asc"; // default de tri projets : avancement 0% -> 100%

let unsavedChanges = false;
let lastUndoSnapshot = null;
let _stateVersion = 0;
let _filteredCache = { key:"", version:-1, tasks:null };
let _missingHoursFlow = null;

let isLocked = true; // verrou logique = droits utilisateur (admin = false)
const isHostedGithubPages = ()=>{
  try{
    const host = (location.hostname || "").toLowerCase();
    return host.endsWith("github.io");
  }catch(e){ return false; }
};
window.isHostedGithubPages = isHostedGithubPages;

let workloadRangeType = "all"; // all | custom | school | civil

let workloadRangeStart = "";

let workloadRangeEnd = "";

let workloadRangeYear = "";

let workloadRangeTypeProject = "all";
let workloadRangeStartProject = "";
let workloadRangeEndProject = "";
let workloadRangeYearProject = "";
let ganttExportContext = "master"; // master | project
let unifiedExportSelectedProjectIds = [];

function resetMasterWorkloadFilters(){
  workloadRangeType = "all";
  workloadRangeYear = "";
  workloadRangeStart = "";
  workloadRangeEnd = "";
  const typeNode = el("workloadRangeType");
  const yearNode = el("workloadRangeYear");
  const startNode = el("workloadRangeStart");
  const endNode = el("workloadRangeEnd");
  if(typeNode) typeNode.value = "all";
  if(yearNode) yearNode.value = "";
  if(startNode) startNode.value = "";
  if(endNode) endNode.value = "";
}

function resetProjectWorkloadFilters(){
  workloadRangeTypeProject = "all";
  workloadRangeYearProject = "";
  workloadRangeStartProject = "";
  workloadRangeEndProject = "";
  const typeNode = el("workloadRangeTypeProject");
  const yearNode = el("workloadRangeYearProject");
  const startNode = el("workloadRangeStartProject");
  const endNode = el("workloadRangeEndProject");
  if(typeNode) typeNode.value = "all";
  if(yearNode) yearNode.value = "";
  if(startNode) startNode.value = "";
  if(endNode) endNode.value = "";
}
const CONFIG_KEY = "dashboard_config_v1";
const USERS_KEY = "dashboard_users_v1";
let ganttColVisibility = {

  masterVendor: true,

  masterStatus: true,

  projectVendor: true,

  projectStatus: true

};



const DEFAULT_STATUSES = [

  {v:"CHANTIER_COMPLET", label:"Chantier complet"},

  {v:"ELECTRICITE", label:"Électricité"},

  {v:"PEINTURE", label:"Peinture"},

  {v:"SOL", label:"Sol"},

  {v:"PLACO", label:"Placo / cloisons"},

  {v:"FAUX_PLAFOND", label:"Faux plafond"},

  {v:"AMENAGEMENTS", label:"Aménagements"},

  {v:"MOBILIER", label:"Mobilier"},

  {v:"PLOMBERIE", label:"Plomberie"},

  {v:"PREPARATION", label:"Préparation"},

  {v:"TDV",          label:"TDV"},

  {v:"MACONNERIE",   label:"Maçonnerie"},

  {v:"HUISSERIES", label:"Huisseries"},

  {v:"RESEAUX",      label:"Réseaux"},

  {v:"TOITURE",      label:"Toiture / étanchéité"},
  {v:"TERRASSEMENT", label:"Terrassement"},
  {v:"ETUDE",        label:"Étude"},

];

let STATUSES = DEFAULT_STATUSES.map(s=>({ ...s }));

const sortedStatuses = ()=> [...STATUSES].sort((a,b)=> a.label.localeCompare(b.label,"fr",{sensitivity:"base"}));



const STATUS_COLORS = {

  CHANTIER_COMPLET: "#1e3a8a",

  ELECTRICITE:      "#d97706",

  PEINTURE:         "#2563eb",

  SOL:              "#0f766e",

  PLACO:            "#7c3aed",

  FAUX_PLAFOND:     "#b45309",

  AMENAGEMENTS:     "#db2777",

  MOBILIER:         "#9333ea",

  PLOMBERIE:        "#15803d",

  PREPARATION:      "#16a34a",

  TDV:              "#f97316",

  MACONNERIE:       "#a16207",

  HUISSERIES:      "#6b7280",


  RESEAUX:          "#0ea5b0",

  TOITURE:          "#0d9488",

  TERRASSEMENT:     "#8b5a2b",

  ETUDE:            "#0284c7",
};
const STATUS_PALETTE = [
  "#0ea5e9",
  "#22c55e",
  "#f59e0b",
  "#ef4444",
  "#8b5cf6",
  "#14b8a6",
  "#84cc16",
  "#f97316",
  "#06b6d4",
  "#e11d48"
];

const THEMES = [
  { id:"clair", label:"Clair neutre", swatch:["#f8fafc","#e2e8f0"] },
  { id:"sable", label:"Sable doux", swatch:["#f2e8d5","#d2b48c"] },
  { id:"ardoise", label:"Ardoise", swatch:["#94a3b8","#475569"] },
  { id:"chantier", label:"Chantier orange", swatch:["#f59e0b","#b45309"] },
  { id:"orange_brule", label:"Orange brûlé", swatch:["#9a3412","#f97316"] },
  { id:"petrole", label:"Vert pétrole", swatch:["#0f766e","#14b8a6"] },
  { id:"nuit", label:"Bleu nuit", swatch:["#0ea5e9","#1e3a8a"] },
  { id:"acier_bleu", label:"Acier bleu", swatch:["#0b1b2b","#1e3a8a"] },
  { id:"marine", label:"Bleu marine", swatch:["#0f172a","#2563eb"] },
  { id:"vert_nuit", label:"Vert nuit", swatch:["#0b1f18","#14532d"] },
  { id:"cobalt_rouge", label:"Cobalt & Rouge", swatch:["#0f3d8a","#c81e1e"] },
  { id:"olive_or", label:"Olive & Or", swatch:["#2f3e1f","#d4a017"] },
  { id:"noir_cyan", label:"Noir & Cyan", swatch:["#0b0f1a","#00c2d1"] },
  { id:"gris", label:"Gris profond", swatch:["#374151","#111827"] },
  { id:"industriel", label:"Industriel sombre", swatch:["#6b7280","#111827"] },
  { id:"carbone", label:"Carbone contrasté", swatch:["#1f2937","#0f172a"] },
  { id:"carbone_nuit", label:"Carbone profond", swatch:["#0a0b0f","#1f2937"] },
  { id:"obsidienne", label:"Obsidienne", swatch:["#0a0a0a","#1f2937"] },
  { id:"noir_or", label:"Noir & Or", swatch:["#111111","#f59e0b"] },
  { id:"noir", label:"Noir absolu", swatch:["#0b0b0b","#f8fafc"] }
];

const THEME_ACCENTS = {
  clair:"#d66b1f",
  sable:"#d66b1f",
  ardoise:"#2563eb",
  chantier:"#f97316",
  orange_brule:"#f97316",
  petrole:"#14b8a6",
  nuit:"#22d3ee",
  acier_bleu:"#3b82f6",
  marine:"#38bdf8",
  vert_nuit:"#22c55e",
  cobalt_rouge:"#ef4444",
  olive_or:"#f59e0b",
  noir_cyan:"#06b6d4",
  gris:"#9ca3af",
  industriel:"#94a3b8",
  carbone:"#3b82f6",
  carbone_nuit:"#64748b",
  obsidienne:"#94a3b8",
  noir_or:"#f59e0b",
  noir:"#f8fafc"
};



const statusColor = (v)=>{
  const key = (v||"").toUpperCase();
  if(STATUS_COLORS[key]) return STATUS_COLORS[key];
  if(!key) return "#1f2937";
  let hash = 0;
  for(let i=0;i<key.length;i++) hash = ((hash<<5)-hash) + key.charCodeAt(i);
  const idx = Math.abs(hash) % STATUS_PALETTE.length;
  return STATUS_PALETTE[idx] || "#1f2937";
};

const statusDot = (v)=> `<span class="icon-dot" style="background:${statusColor(v)};border-color:${statusColor(v)}"></span>`;

const parseStatuses = (s)=> (s||"").split(",").map(x=>x.trim()).filter(Boolean);

const taskProgress = (t)=>{
  if(!t?.start || !t?.end) return 0;
  const s = new Date(t.start+"T00:00:00");
  const e = new Date(t.end+"T00:00:00");
  if(isNaN(s) || isNaN(e) || e < s) return 0;
  const today = new Date();
  today.setHours(0,0,0,0);
  if(today <= s) return 0;
  if(today >= e) return 100;
  const total = countWeekdays(s, e);
  if(!total) return 0;
  const elapsed = countWeekdays(s, today);
  return Math.max(0, Math.min(100, Math.round((elapsed / total) * 100)));
};

function setTaskProgressUI(val){
  const v = Math.max(0, Math.min(100, Math.round(val || 0)));
  const fill = el("taskProgressFill");
  const label = el("taskProgressLabel");
  if(fill) fill.style.width = `${v}%`;
  if(label) label.textContent = `${v}%`;
}
function setProjectProgressUI(val){
  const v = Math.max(0, Math.min(100, Math.round(val || 0)));
  const fill = el("projectProgressFill");
  const label = el("projectProgressLabel");
  if(fill) fill.style.width = `${v}%`;
  if(label) label.textContent = `${v}%`;
}

function calcProgressFromInputs(){
  const s = unformatDate(el("t_start")?.value || "");
  const e = unformatDate(el("t_end")?.value || "");
  return taskProgress({start:s, end:e});
}

const deepClone = (obj)=> JSON.parse(JSON.stringify(obj));

const siteColor = (_site="")=>"transparent";

const attrEscape = (s="")=> s.replace(/&/g,"&amp;").replace(/"/g,"&quot;").replace(/</g,"&lt;").replace(/>/g,"&gt;");

const uiUpperNoAccent = (value="")=>
  String(value)
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .toUpperCase();

function applyUiUpperNoAccent(root=document.body){
  if(!root) return;
  if(root.nodeType === Node.TEXT_NODE){
    const parent = root.parentElement;
    if(parent){
      const tag = parent.tagName;
      if(tag==="SCRIPT" || tag==="STYLE" || tag==="NOSCRIPT" || tag==="TEXTAREA") return;
      if(parent.closest("[contenteditable='true']")) return;
    }
    if(root.nodeValue && root.nodeValue.trim()){
      root.nodeValue = uiUpperNoAccent(root.nodeValue);
    }
    return;
  }
  const walker = document.createTreeWalker(root, NodeFilter.SHOW_TEXT, {
    acceptNode(node){
      if(!node || !node.nodeValue || !node.nodeValue.trim()) return NodeFilter.FILTER_REJECT;
      const parent = node.parentElement;
      if(!parent) return NodeFilter.FILTER_REJECT;
      const tag = parent.tagName;
      if(tag==="SCRIPT" || tag==="STYLE" || tag==="NOSCRIPT" || tag==="TEXTAREA") return NodeFilter.FILTER_REJECT;
      if(parent.closest("[contenteditable='true']")) return NodeFilter.FILTER_REJECT;
      return NodeFilter.FILTER_ACCEPT;
    }
  });
  let node = null;
  while((node = walker.nextNode())){
    node.nodeValue = uiUpperNoAccent(node.nodeValue);
  }
  if(typeof root.querySelectorAll !== "function") return;
  root.querySelectorAll("input[placeholder], textarea[placeholder]").forEach((n)=>{
    n.placeholder = uiUpperNoAccent(n.placeholder || "");
  });
  root.querySelectorAll("option").forEach((opt)=>{
    opt.textContent = uiUpperNoAccent(opt.textContent || "");
  });
}

let uiUpperObserver = null;
function ensureUiUpperNoAccentObserver(){
  if(uiUpperObserver || !document.body) return;
  uiUpperObserver = new MutationObserver((mutations)=>{
    for(const mutation of mutations){
      mutation.addedNodes?.forEach((node)=>{
        applyUiUpperNoAccent(node);
      });
    }
  });
  uiUpperObserver.observe(document.body, { childList:true, subtree:true });
}

function updateTopbarHeight(){

  const tb = document.querySelector(".topbar");

  if(!tb) return;

  document.documentElement.style.setProperty("--topbar-h", `${tb.offsetHeight}px`);

}

function updateSidebarTop(){
  const tb = document.querySelector(".topbar");
  if(!tb) return;
  const top = Math.max(0, tb.offsetHeight + 14);
  window.__sidebarTopLocked = top;
  document.documentElement.style.setProperty("--sidebar-top", `${top}px`);
}
function updateSidebarScrollState(){
  const sb = document.querySelector(".sidebar");
  if(!sb) return;
  // Sidebar ne doit plus scroller : seul le bloc Onglets scrolle
  sb.classList.remove("sidebar-scroll");
}

function scrollViewToTop(){
  try{ window.scrollTo(0,0); }catch(e){ softCatch(e); }
  document.querySelectorAll(".tablewrap").forEach(el=>{
    el.scrollTop = 0;
    el.scrollLeft = 0;
  });
  document.querySelectorAll(".tabs-scroll").forEach(el=>{
    el.scrollTop = 0;
  });
}

// verrouille la position de la sidebar une fois la mise en page stabilisée
function _lockSidebarAfterLayout(){
  updateSidebarTop();
  updateSidebarScrollState();
  applySidebarTopLock();
}
window.addEventListener("load", _lockSidebarAfterLayout);
function applySidebarTopLock(){
  if(typeof window.__sidebarTopLocked !== "number") return;
  document.documentElement.style.setProperty("--sidebar-top", `${window.__sidebarTopLocked}px`);
}

function updateTaskDatesWarning(){

  const warn = el("t_dates_warn");

  const start = el("t_start")?.value;

  const end = el("t_end")?.value;

  if(!warn) return;

  if(start && end && end < start){

    warn.classList.remove("hidden");

  }else{

    warn.classList.add("hidden");

  }

}

function formatShortDate(d){

  const dd = String(d.getDate()).padStart(2,"0");

  const mm = String(d.getMonth()+1).padStart(2,"0");

  const yy = String(d.getFullYear()).slice(-2);

  return `${dd}-${mm}-${yy}`;

}

function formatShortDateTwoLinesHTML(d){
  const dd = String(d.getDate()).padStart(2,"0");
  const mm = String(d.getMonth()+1).padStart(2,"0");
  const yy = String(d.getFullYear()).slice(-2);
  return `<span class="wk-date-top">${dd}-${mm}</span><span class="wk-date-bottom">${yy}</span>`;
}

function isTodayInWeek(weekStart){

  const today = new Date();

  const s = startOfWeek(today);

  return +s === +weekStart;

}

function setToggleBtnState(id, isOn){

  const b = el(id);

  if(!b) return;

  b.classList.toggle("btn-primary", !!isOn);

  b.classList.toggle("btn-ghost", !isOn);

}

function applyGanttColumnVisibility(){

  const masterTable = document.querySelector("#masterGantt table");

  if(masterTable){

    masterTable.classList.toggle("hide-vendor", !ganttColVisibility.masterVendor);

    masterTable.classList.toggle("hide-status", !ganttColVisibility.masterStatus);

  }

  const projectTable = document.querySelector("#gantt table");

  if(projectTable){

    projectTable.classList.toggle("hide-vendor", !ganttColVisibility.projectVendor);

    projectTable.classList.toggle("hide-status", !ganttColVisibility.projectStatus);

  }

  setToggleBtnState("btnToggleMasterVendor", ganttColVisibility.masterVendor);

  setToggleBtnState("btnToggleMasterStatus", ganttColVisibility.masterStatus);

  setToggleBtnState("btnToggleProjectVendor", ganttColVisibility.projectVendor);

  setToggleBtnState("btnToggleProjectStatus", ganttColVisibility.projectStatus);

}

function openTaskFromGantt(taskId){
  const task = state.tasks.find(x=>x.id===taskId);
  if(!task) return;
  selectedTaskId = taskId;
  navigateTo(task.projectId, taskId, true);
}
function closeAllOverlays(){
  try{ showVendorDropdown(false); }catch(e){ softCatch(e); }
  try{ showDescriptionDropdown(false); }catch(e){ softCatch(e); }
  try{ toggleStatusMenu(false); }catch(e){ softCatch(e); }
  const vBox = el("vendorDropdown");
  if(vBox){ vBox.style.display="none"; vBox.classList.remove("open"); }
  const dBox = el("descDropdown");
  if(dBox){ dBox.style.display="none"; dBox.classList.remove("open"); }
  const vPanel = el("vendorManagerPanel");
  if(vPanel) vPanel.style.display="none";
  const dPanel = el("descManagerPanel");
  if(dPanel) dPanel.style.display="none";
  const statusMenu = el("t_status_menu");
  if(statusMenu) statusMenu.classList.add("hidden");
  const overlay = el("descOverlay");
  if(overlay) overlay.classList.remove("show");
  document.querySelectorAll(".vendor-open").forEach(n=>n.classList.remove("vendor-open"));
  document.querySelectorAll(".desc-open").forEach(n=>n.classList.remove("desc-open"));
  document.querySelectorAll(".desc-panel-open").forEach(n=>n.classList.remove("desc-panel-open"));
  floatingMap.forEach((anchor, el)=>{ closeFloating(el); });
}
function loadConfig(){
  try{
    const raw = localStorage.getItem(CONFIG_KEY);
    return raw ? JSON.parse(raw) : {};
  }catch(e){
    return {};
  }
}
function saveConfig(cfg){
  try{
    localStorage.setItem(CONFIG_KEY, JSON.stringify(cfg||{}));
  }catch(e){ softCatch(e); }
}
function normalizeStatusId(label){
  const base = (label||"").trim();
  if(!base) return "";
  const ascii = base.normalize("NFD").replace(/[\u0300-\u036f]/g,"");
  return ascii.toUpperCase().replace(/[^A-Z0-9]+/g,"_").replace(/^_+|_+$/g,"");
}
function normalizeStatusList(list){
  const out=[];
  const seen=new Set();
  (list||[]).forEach(item=>{
    if(!item) return;
    const label = (item.label || item.name || "").toString().trim();
    if(!label) return;
    const vRaw = (item.v || "").toString().trim();
    const v = (vRaw ? vRaw : normalizeStatusId(label)).toUpperCase();
    if(!v) return;
    if(seen.has(v)) return;
    seen.add(v);
    out.push({v, label});
  });
  return out;
}
function loadStatusConfig(){
  const cfg = loadConfig();
  if(Array.isArray(cfg.statuses) && cfg.statuses.length){
    STATUSES = normalizeStatusList(cfg.statuses);
  }else{
    STATUSES = deepClone(DEFAULT_STATUSES);
  }
}
function saveStatusConfig(){
  const cfg = loadConfig();
  cfg.statuses = STATUSES;
  saveConfig(cfg);
}
function getHoursConfig(){
  return { internal: 4, external: 4, rsg: 2 };
}
function loadUsers(){
  try{
    const raw = localStorage.getItem(USERS_KEY);
    return raw ? JSON.parse(raw) : [];
  }catch(e){
    return [];
  }
}
function saveUsers(list){
  try{
    localStorage.setItem(USERS_KEY, JSON.stringify(list||[]));
    try{
      if(typeof window.populateLoginUsers === "function") window.populateLoginUsers();
    }catch(e){ softCatch(e); }
    try{ saveUsersToSupabase(list||[]); }catch(e){ softCatch(e); }
  }catch(e){ softCatch(e); }
}

function getCurrentUserName(){
  return sessionStorage.getItem("current_user") || "";
}
function getCurrentUserEmail(){
  return sessionStorage.getItem("current_email") || "";
}
function getCurrentUserRecord(){
  const email = getCurrentUserEmail();
  const name = getCurrentUserName();
  const users = loadUsers();
  if(email){
    const found = users.find(u=>(u.email||"").toLowerCase()===email.toLowerCase());
    if(found) return found;
  }
  if(!name) return null;
  return users.find(u=>u.name===name) || null;
}
function applyTheme(themeId){
  const id = themeId || "sable";
  document.documentElement.setAttribute("data-theme", id);
  const grid = el("themeGrid");
  if(grid){
    grid.querySelectorAll(".theme-swatch").forEach(n=>{
      n.classList.toggle("active", n.dataset.theme===id);
    });
  }
}
function applyThemeForCurrentUser(){
  const u = getCurrentUserRecord();
  const sessionTheme = sessionStorage.getItem("current_theme") || "";
  const theme = (u && u.theme) ? u.theme : (sessionTheme || "sable");
  applyTheme(theme);
}
function setCurrentUserTheme(themeId){
  const name = getCurrentUserName();
  const email = getCurrentUserEmail();
  try{ sessionStorage.setItem("current_theme", themeId || "sable"); }catch(e){ softCatch(e); }
  if(!name){
    applyTheme(themeId);
    return;
  }
  const users = loadUsers();
  let idx = -1;
  if(email){
    idx = users.findIndex(u=>(u.email||"").toLowerCase()===email.toLowerCase());
  }
  if(idx < 0){
    idx = users.findIndex(u=>u.name===name);
  }
  if(idx>=0){
    users[idx].theme = themeId;
    saveUsers(users);
  }
  applyTheme(themeId);
}
function initThemePicker(){
  const grid = el("themeGrid");
  const picker = el("themePicker");
  const toggle = el("themeToggle");
  if(!grid || !picker || !toggle) return;
  const hexToRgb = (hex)=>{
    const v = (hex || "").replace("#","").trim();
    if(v.length !== 6) return {r:0,g:0,b:0};
    return {
      r: parseInt(v.slice(0,2),16),
      g: parseInt(v.slice(2,4),16),
      b: parseInt(v.slice(4,6),16)
    };
  };
  const luma = (hex)=>{
    const {r,g,b} = hexToRgb(hex);
    return 0.2126*r + 0.7152*g + 0.0722*b;
  };
  const themesSorted = [...THEMES].sort((a,b)=>{
    const ac = a.swatch || ["#ffffff","#ffffff"];
    const bc = b.swatch || ["#ffffff","#ffffff"];
    const al = (luma(ac[0]) + luma(ac[1])) / 2;
    const bl = (luma(bc[0]) + luma(bc[1])) / 2;
    return bl - al; // clair -> foncé
  });
  grid.innerHTML = themesSorted.map(t=>{
    const colors = t.swatch || ["#e2e8f0","#94a3b8"];
    const accent = THEME_ACCENTS[t.id] || colors[1] || "#94a3b8";
    const style = `background:linear-gradient(135deg, ${colors[0]} 0%, ${colors[1]} 55%, ${accent} 100%);`;
    return `<button type="button" class="theme-swatch" data-theme="${t.id}" title="${t.label}" style="${style}"></button>`;
  }).join("");

  grid.querySelectorAll(".theme-swatch").forEach(btn=>{
    btn.addEventListener("click", ()=> setCurrentUserTheme(btn.dataset.theme || "sable"));
    btn.addEventListener("mouseenter", ()=>{
      const t = btn.dataset.theme || "sable";
      applyTheme(t);
    });
    btn.addEventListener("mouseleave", ()=>{
      applyThemeForCurrentUser();
    });
  });
  toggle.addEventListener("click", (e)=>{
    e.stopPropagation();
    picker.classList.toggle("open");
    if(!picker.classList.contains("open")) applyThemeForCurrentUser();
  });
  document.addEventListener("click", (e)=>{
    if(!picker.classList.contains("open")) return;
    if(picker.contains(e.target)) return;
    picker.classList.remove("open");
    applyThemeForCurrentUser();
  });
  applyThemeForCurrentUser();
}
function getCurrentRole(){
  return sessionStorage.getItem("current_role") || "user";
}
function updateRoleUI(){
  const role = getCurrentRole();
  const cfgBtn = el("btnConfig");
  if(cfgBtn) cfgBtn.style.display = (role==="admin") ? "inline-flex" : "none";
  const topUser = el("topbarUser");
  if(topUser){
    const name = sessionStorage.getItem("current_user") || "Invité";
    const email = sessionStorage.getItem("current_email") || "";
    const roleLabel = role==="admin" ? "Admin" : "Utilisateur";
    const emailPart = email ? ` - ${email}` : "";
    topUser.textContent = `Utilisateur connecté: ${name}${emailPart} - ${roleLabel}`;
  }
  applyThemeForCurrentUser();
}
function applyRoleAccess(){
  const role = getCurrentRole();
  // UI métier accessible à tous; "Configuration" reste réservé aux admins.
  isLocked = false;
  const lockClass = "is-disabled";
  const ids = [
    "btnAddProject","btnAddTask",
    "btnSaveProject","btnDeleteProject",
    "btnSaveTask","btnNewTask","btnDuplicateTask","btnDeleteTask"
  ];
  ids.forEach(id=>{
    const n=el(id);
    if(!n) return;
    n.classList.toggle(lockClass, false);
    n.removeAttribute("disabled");
  });
  const cfgBtn = el("btnConfig");
  if(cfgBtn){
    cfgBtn.style.display = role==="admin" ? "inline-flex" : "none";
  }
  const switchBtn = el("btnSwitchUser");
  if(switchBtn){
    switchBtn.style.display = "inline-flex";
  }
  const logoutBtn = el("btnLogout");
  if(logoutBtn){
    logoutBtn.style.display = "inline-flex";
  }
  const manageBtn = el("btnManageVendors");
  if(manageBtn){
    manageBtn.classList.toggle(lockClass, false);
    manageBtn.removeAttribute("disabled");
  }
  const manageDescBtn = el("btnManageDescriptions");
  if(manageDescBtn){
    manageDescBtn.classList.toggle(lockClass, false);
    manageDescBtn.removeAttribute("disabled");
  }
  const tabCloses = document.querySelectorAll(".tab-close");
  tabCloses.forEach(n=>{
    n.classList.toggle(lockClass, false);
    n.removeAttribute("aria-disabled");
  });
  const dangerBtns = document.querySelectorAll("button.btn-danger");
  dangerBtns.forEach(btn=>{
    btn.classList.toggle(lockClass, false);
    btn.removeAttribute("disabled");
  });
  const live = el("masterLive");
  if(live){
    live.classList.toggle("is-disabled", false);
  }
  const plive = el("projectLive");
  if(plive){
    plive.classList.toggle("is-disabled", false);
  }
  updateRoleUI();
}
window.applyRoleAccess = applyRoleAccess;
async function hashPassword(str){
  const enc = new TextEncoder().encode(str || "");
  const buf = await crypto.subtle.digest("SHA-256", enc);
  return [...new Uint8Array(buf)].map(b=>b.toString(16).padStart(2,"0")).join("");
}
function renderUsersList(){
  const list = el("cfg_users_list");
  if(!list) return;
  const users = loadUsers();
  if(users.length===0){
    list.innerHTML = `<div class="config-user-meta">Aucun utilisateur</div>`;
    return;
  }
  list.innerHTML = users.map(u=>`
    <div class="config-user-item">
      <div>
        <div><strong>${attrEscape(u.name||"")}</strong></div>
        <div class="config-user-meta">${attrEscape(u.email||"") || ""}  ${u.role==="admin" ? "Admin" : "Utilisateur"}</div>
      </div>
      <div class="config-user-actions">
        <button class="btn btn-ghost cfg-user-edit" data-user-id="${attrEscape(u.id||"")}">Modifier</button>
        <button class="btn btn-ghost cfg-user-pass" data-user-id="${attrEscape(u.id||"")}">Changer mdp</button>
        <button class="btn btn-danger cfg-user-del" data-user-id="${attrEscape(u.id||"")}">Supprimer</button>
      </div>
    </div>
  `).join("");
  list.querySelectorAll(".cfg-user-pass").forEach(btn=>{
    btn.onclick = async ()=>{
      const id = btn.dataset.userId;
      const users = loadUsers();
      const u = users.find(x=>x.id===id);
      if(!u) return;
      const next = prompt(`Nouveau mot de passe pour ${u.name} :`);
      if(!next) return;
      const h = await hashPassword(next);
      users.forEach(x=>{ if(x.id===id) x.hash = h; });
      saveUsers(users);
      renderUsersList();
    };
  });
  list.querySelectorAll(".cfg-user-edit").forEach(btn=>{
    btn.onclick = ()=>{
      const id = btn.dataset.userId;
      const users = loadUsers();
      const idx = users.findIndex(u=>u.id===id);
      if(idx<0) return;
      const newName = prompt("Nom d'utilisateur :", users[idx].name) || "";
      if(!newName) return;
      const newEmail = prompt("Email :", users[idx].email||"") || "";
      const newRole = (prompt("Rôle (admin/user) :", users[idx].role||"user") || users[idx].role || "user").toLowerCase();
      const role = newRole === "admin" ? "admin" : "user";
      if(newEmail && users.some((u,i)=> i!==idx && (u.email||"").toLowerCase()===newEmail.toLowerCase())){
        alert("Email déjà utilisé."); return;
      }
      users[idx].name = newName;
      users[idx].email = newEmail;
      users[idx].role = role;
      saveUsers(users);
      const current = sessionStorage.getItem("current_user") || "";
      const currentEmail = sessionStorage.getItem("current_email") || "";
      if((currentEmail && users[idx].email && currentEmail.toLowerCase()===users[idx].email.toLowerCase()) || (current && current === users[idx].name)){
        sessionStorage.setItem("current_user", newName);
        sessionStorage.setItem("current_role", role);
        if(users[idx].email) sessionStorage.setItem("current_email", users[idx].email);
        updateRoleUI();
      }
      renderUsersList();
    };
  });
  list.querySelectorAll(".cfg-user-del").forEach(btn=>{
    btn.onclick = ()=>{
      const id = btn.dataset.userId;
      const users = loadUsers();
      const u = users.find(x=>x.id===id);
      if(!u) return;
      if(!confirm(`Supprimer ${u.name} ?`)) return;
      const admins = users.filter(u=>u.role==="admin");
      if(admins.length<=1 && admins.some(x=>x.id===id)){
        alert("Impossible de supprimer le dernier admin.");
        return;
      }
      const next = users.filter(x=>x.id!==id);
      users.length = 0;
      users.push(...next);
      saveUsers(users);
      renderUsersList();
    };
  });
}

function refreshStatusUi(){
  buildStatusMenu();
  updateStatusDisplay();
  renderFilters();
  if(selectedProjectId) renderProject();
  else renderMaster();
}

function renderConfigStatusList(){
  const list = el("cfg_status_list");
  if(!list) return;
  if(!STATUSES || STATUSES.length===0){
    list.innerHTML = `<div class="config-user-meta">Aucun statut</div>`;
    return;
  }
  const sorted = [...STATUSES].sort((a,b)=> (a.label||"").localeCompare((b.label||""),"fr",{sensitivity:"base"}));
  list.innerHTML = sorted.map(s=>`
    <div class="config-user-item">
      <div>
        <div><strong>${attrEscape(s.label||"")}</strong></div>
        <div class="config-user-meta">${attrEscape(s.v||"")}</div>
      </div>
      <div class="config-user-actions">
        <button class="btn btn-ghost cfg-status-edit" data-v="${attrEscape(s.v||"")}">Renommer</button>
        <button class="btn btn-danger cfg-status-del" data-v="${attrEscape(s.v||"")}">Supprimer</button>
      </div>
    </div>
  `).join("");

  list.querySelectorAll(".cfg-status-edit").forEach(btn=>{
    btn.onclick = ()=>{
      const v = (btn.dataset.v || "").toUpperCase();
      const s = STATUSES.find(x=>x.v===v);
      if(!s) return;
      const next = prompt("Nouveau libellé du statut :", s.label || "") || "";
      const trimmed = next.trim();
      if(!trimmed) return;
      s.label = trimmed;
      saveStatusConfig();
      renderConfigStatusList();
      refreshStatusUi();
    };
  });

  list.querySelectorAll(".cfg-status-del").forEach(btn=>{
    btn.onclick = ()=>{
      const v = (btn.dataset.v || "").toUpperCase();
      const s = STATUSES.find(x=>x.v===v);
      if(!s) return;
      if(!confirm(`Supprimer le statut "${s.label}" ?`)) return;
      STATUSES = STATUSES.filter(x=>x.v!==v);
      // retirer de toutes les tâches
      (state?.tasks||[]).forEach(t=>{
        const list = parseStatuses(t.status).filter(x=>x.toUpperCase()!==v);
        t.status = list.join(",");
      });
      saveStatusConfig();
      renderConfigStatusList();
      setStatusSelection("");
      refreshStatusUi();
    };
  });
}

function renderConfigVendorsList(){
  const list = el("cfg_vendor_list");
  if(!list) return;
  const registry = loadVendorsRegistry();
  const deleted = new Set(loadDeletedVendors().map(x=>x.toLowerCase()));
  const vendors = dedupVendors(registry).filter(v=>!deleted.has(v.toLowerCase()))
    .sort((a,b)=>a.localeCompare(b,"fr",{sensitivity:"base"}));
  vendorsCache = vendors.slice();
  if(vendors.length===0){
    list.innerHTML = `<div class="config-user-meta">Aucun prestataire</div>`;
    return;
  }
  list.innerHTML = vendors.map(v=>`
    <div class="config-user-item">
      <div>
        <div><strong>${attrEscape(v)}</strong></div>
        <div class="config-user-meta">Prestataire externe</div>
      </div>
      <div class="config-user-actions">
        <button class="btn btn-ghost cfg-vendor-edit" data-v="${attrEscape(v)}">Renommer</button>
        <button class="btn btn-danger cfg-vendor-del" data-v="${attrEscape(v)}">Supprimer</button>
      </div>
    </div>
  `).join("");

  list.querySelectorAll(".cfg-vendor-edit").forEach(btn=>{
    btn.onclick = ()=>{
      const oldName = btn.dataset.v || "";
      const next = prompt("Nouveau nom du prestataire :", oldName) || "";
      const trimmed = next.trim();
      if(!trimmed) return;
      let deleted = loadDeletedVendors().filter(x=> x.toLowerCase()!==oldName.toLowerCase());
      saveDeletedVendors(deleted);
      vendorsCache = vendorsCache.map(x=> x.toLowerCase()===oldName.toLowerCase() ? trimmed : x);
      vendorsCache = dedupVendors(vendorsCache).sort((a,b)=>a.localeCompare(b,"fr",{sensitivity:"base"}));
      saveVendorsRegistry(vendorsCache);
      (state?.tasks||[]).forEach(t=>{
        if((t.vendor||"").toLowerCase()===oldName.toLowerCase()) t.vendor = trimmed;
      });
      renderConfigVendorsList();
      refreshVendorsList();
      renderAll();
    };
  });

  list.querySelectorAll(".cfg-vendor-del").forEach(btn=>{
    btn.onclick = ()=>{
      const name = btn.dataset.v || "";
      if(!name) return;
      if(!confirm(`Supprimer le prestataire "${name}" ?`)) return;
      vendorsCache = vendorsCache.filter(x=>x.toLowerCase()!==name.toLowerCase());
      vendorsCache = dedupVendors(vendorsCache);
      saveVendorsRegistry(vendorsCache);
      const deleted = Array.from(new Set([...loadDeletedVendors(), name]));
      saveDeletedVendors(deleted);
      (state?.tasks||[]).forEach(t=>{
        if((t.vendor||"").toLowerCase()===name.toLowerCase()) t.vendor = "";
      });
      saveState();
      renderConfigVendorsList();
      refreshVendorsList();
      renderAll();
    };
  });
}
function openConfigModal(){
  if(getCurrentRole()!=="admin") return;
  const modal = el("configModal");
  if(!modal) return;
  const cfg = loadConfig();
  el("cfg_name").value = cfg.name || "";
  el("cfg_http").value = cfg.http || "";
  el("cfg_front").value = cfg.front || "";
  el("cfg_back").value = cfg.back || "";
  const linkify = (inputId, linkId)=>{
    const input = el(inputId);
    const link = el(linkId);
    if(!link) return;
    const val = (input?.value || "").trim();
    if(val){
      link.href = val;
      link.textContent = val;
      link.classList.remove("hidden");
    }else{
      link.removeAttribute("href");
      link.textContent = "Ouvrir";
      link.classList.add("hidden");
    }
  };
  linkify("cfg_http","cfg_http_link");
  linkify("cfg_front","cfg_front_link");
  linkify("cfg_back","cfg_back_link");
  const timeDateInput = el("t_time_date_input");
  timeDateInput?.addEventListener("change", ()=>{
    const t = state?.tasks?.find(x=>x.id===selectedTaskId && x.projectId===selectedProjectId);
    updateTimeLogUI(t || null, true);
  });
  timeDateInput?.addEventListener("input", ()=>{
    const t = state?.tasks?.find(x=>x.id===selectedTaskId && x.projectId===selectedProjectId);
    updateTimeLogUI(t || null, true);
  });
  const role = getCurrentRole();
  const usersSection = el("cfg_users_section");
  if(usersSection){
    const wrap = usersSection.closest(".cfg-accordion");
    if(wrap) wrap.style.display = role==="admin" ? "block" : "none";
    usersSection.style.display = "";
  }
  if(role==="admin") renderUsersList();
  const statusSection = el("cfg_status_section");
  if(statusSection){
    const wrap = statusSection.closest(".cfg-accordion");
    if(wrap) wrap.style.display = role==="admin" ? "block" : "none";
    statusSection.style.display = "";
  }
  if(role==="admin"){
    loadStatusConfig();
    renderConfigStatusList();
  }
  const vendorSection = el("cfg_vendor_section");
  if(vendorSection){
    const wrap = vendorSection.closest(".cfg-accordion");
    if(wrap) wrap.style.display = role==="admin" ? "block" : "none";
    vendorSection.style.display = "";
  }
  if(role==="admin"){
    refreshVendorsList();
    renderConfigVendorsList();
  }
  const loginSection = el("cfg_login_section");
  if(loginSection){
    const wrap = loginSection.closest(".cfg-accordion");
    if(wrap) wrap.style.display = role==="admin" ? "block" : "none";
    loginSection.style.display = "";
  }
  if(role==="admin") initLoginJournalUI();
  initVacationConfigUI();
  modal.classList.remove("hidden");
  modal.style.display = "flex";
  modal.setAttribute("aria-hidden","false");
}
function closeConfigModal(){
  const modal = el("configModal");
  if(!modal) return;
  hideModalSafely(modal);
}
function resetModalScrollTop(modal){
  if(!modal) return;
  try{
    modal.scrollTop = 0;
    modal.querySelectorAll(".modal-card,.modal-body,.modal-content,.config-col,.config-users-list,.export-modules-list,#exportPdfModulesList").forEach((n)=>{
      if(n && typeof n.scrollTop === "number") n.scrollTop = 0;
    });
  }catch(e){ softCatch(e); }
}
function showModalSafely(modal){
  if(!modal) return;
  modal.classList.remove("hidden");
  modal.style.display = "flex";
  modal.setAttribute("aria-hidden","false");
  resetModalScrollTop(modal);
  try{
    requestAnimationFrame(()=> resetModalScrollTop(modal));
    setTimeout(()=> resetModalScrollTop(modal), 60);
  }catch(e){ softCatch(e); }
}
function hideModalSafely(modal, focusFallbackSelector=""){
  if(!modal) return;
  try{
    const active = document.activeElement;
    if(active && modal.contains(active) && typeof active.blur === "function"){
      active.blur();
    }
  }catch(e){ softCatch(e); }
  try{
    if(focusFallbackSelector){
      const target = document.querySelector(focusFallbackSelector);
      if(target && typeof target.focus === "function") target.focus();
    }
  }catch(e){ softCatch(e); }
  modal.classList.add("hidden");
  modal.style.display = "none";
  modal.setAttribute("aria-hidden","true");
}

function initVacationConfigUI(){
  const yearInput = el("cfg_vac_year");
  const schoolInput = el("cfg_vac_school");
  const internalInput = el("cfg_vac_internal");
  const list = el("cfg_vac_years");
  if(!yearInput || !schoolInput || !internalInput) return;
  const years = getVacationYears();
  if(list){
    list.innerHTML = years.map(y=>`<option value="${y}"></option>`).join("");
  }
  const setYear = (y)=>{
    if(!y) return;
    yearInput.value = y;
    schoolInput.value = (VACANCES_ZONE_B_WEEKS[y] || []).join(",");
    internalInput.value = (VACANCES_INTERNE_WEEKS[y] || []).join(",");
  };
  if(!yearInput.value){
    const defYear = getSchoolYearKey(new Date());
    setYear(years.includes(defYear) ? defYear : (years[0] || defYear));
  }else{
    setYear(yearInput.value.trim());
  }
}
function navigateTo(projectId=null, taskId=null, push=true){
  selectedProjectId = projectId || null;
  selectedTaskId = taskId || null;
  if(push){
    try{
      history.pushState({projectId:selectedProjectId, taskId:selectedTaskId}, "");
    }catch(e){ softCatch(e); }
  }
  closeAllOverlays();
  setTimeout(()=> closeAllOverlays(), 0);
  const view = document.querySelector(".main");
  if(view){
    view.classList.remove("view-fade");
    void view.offsetWidth;
    view.classList.add("view-fade");
  }
  renderAll();
  setTimeout(()=> scrollViewToTop(), 0);
}
function selectTaskInProject(taskId){
  selectedTaskId = taskId || null;
  renderProject();
}
const FLOAT_Z = 1000000;

const floatingMap = new Map();

function positionFloating(el, anchor){

  if(!el || !anchor) return;

  const rect = anchor.getBoundingClientRect();

  const top = rect.bottom + 4;

  const maxH = Math.max(120, window.innerHeight - top - 12);

  el.style.position = "fixed";

  el.style.left = `${rect.left}px`;

  el.style.top = `${top}px`;

  el.style.width = `${rect.width}px`;

  el.style.maxHeight = `${maxH}px`;

  el.style.zIndex = `${FLOAT_Z}`;

}

function openFloating(el, anchor){

  if(!el || !anchor) return;

  if(!el.__floatParent){

    el.__floatParent = el.parentElement;

    el.__floatNext = el.nextSibling;

  }

  document.body.appendChild(el);

  positionFloating(el, anchor);

  floatingMap.set(el, anchor);

}

function closeFloating(el){

  if(!el || !el.__floatParent) return;

  floatingMap.delete(el);

  el.style.position = "";

  el.style.left = "";

  el.style.top = "";

  el.style.width = "";

  el.style.maxHeight = "";

  el.style.zIndex = "";

  if(el.__floatNext) el.__floatParent.insertBefore(el, el.__floatNext);

  else el.__floatParent.appendChild(el);

}

let vendorsCache = [];

let descCache = [];

const VENDOR_STORE_KEY = "vendors_registry";

const VENDOR_DELETE_KEY = "vendors_deleted";

const DESC_STORE_KEY = "descriptions_registry";

const DESC_DELETE_KEY = "descriptions_deleted";

const normalizeVendor = (v="")=> v.trim();

const normalizeDesc = (v="")=> v.trim();

const dedupVendors = (arr=[])=>{

  const seen=new Set();

  const out=[];

  arr.forEach(v=>{

    const norm = normalizeVendor(v);

    if(!norm) return;

    const key = norm.toLowerCase();

    if(seen.has(key)) return;

    seen.add(key);

    out.push(norm);

  });

  return out;

};

const dedupDescriptions = (arr=[])=>{

  const seen=new Set();

  const out=[];

  arr.forEach(v=>{

    const norm = normalizeDesc(v);

    if(!norm) return;

    const key = norm.toLowerCase();

    if(seen.has(key)) return;

    seen.add(key);

    out.push(norm);

  });

  return out;

};

const ownerType = (o="")=>{

  const k=o.toLowerCase();

  if(k.includes("rsg/ri")) return "rsg";
  if(k.includes("rsg")) return "rsg";
  if(k.includes("ri")) return "ri";

  const hasInt = k.includes("interne");

  const hasExt = k.includes("externe");

  // Plus de catégorie "mixte" : on priorise "interne" si exclusif, sinon "externe".

  if(hasInt && !hasExt) return "interne";

  if(hasExt) return "externe";

  return "inconnu";

};

const ownerBadge = (o="", labelOverride="")=>{

  const k = o.toLowerCase();
  const label = (labelOverride || o || "").toString();

  // Palette aligne avec le graphique de charge

  let color = "#16a34a"; // interne par défaut

  if(k.includes("rsg/ri") || k.includes("rsg")) color = "#2563eb"; // RSG
  if(k.includes("ri")) color = "#7c3aed"; // RI
  if(k.includes("interne") && k.includes("externe")) color = "#b45309"; // mix -> externe

  else if(k.includes("externe")) color = "#b45309"; // prestataire externe

  else if(k.includes("interne")) color = "#16a34a"; // INTERNE

  return `<span class="badge owner" style="background:${color};border-color:${color};color:#fff;">${label}</span>`;

};

function ownerBadgeForTask(t){
  if(!t) return "";
  const owner = t.owner || "";
  if(!owner) return "";
  const typ = ownerType(owner);
  let label = owner;
  if(typ === "interne"){
    label = "INTERNE";
  }
  if(typ === "rsg"){
    label = "RSG";
  }
  if(typ === "ri"){
    label = "RI";
  }
  if(typ === "externe"){
    const v = (t.vendor || "").trim();
    label = v || "Prestataire non renseigné";
  }
  return ownerBadge(owner, label);
}

const SITE_PHOTOS = {
  "CDM": "assets/sites/CDM.jpg",
  "LGT": "assets/sites/LGT.jpg",
  "COLLÈGE": "assets/sites/College.jpg",
  "COLLEGE": "assets/sites/College.jpg",
  "ÉCOLE": "assets/sites/Ecole.jpg",
  "ECOLE": "assets/sites/Ecole.jpg",
  "NDC": "assets/sites/NDC.jpg"
};

function updateSitePhoto(site){
  const img = el("sitePhoto");
  const empty = el("sitePhotoEmpty");
  if(!img || !empty) return;
  const key = (site||"").trim().toUpperCase();
  const src = SITE_PHOTOS[key];
  if(!src){
    img.src = "";
    img.style.display = "none";
    empty.style.display = "block";
    return;
  }
  img.onerror = () => {
    img.style.display = "none";
    empty.style.display = "block";
  };
  img.onload = () => {
    img.style.display = "block";
    empty.style.display = "none";
  };
  img.src = src;
}
const ownerColor = (o="")=>{
  const typ = ownerType(o);
  if(typ === "interne") return "#16a34a";
  if(typ === "externe") return "#b45309";
  if(typ === "rsg") return "#2563eb";
  if(typ === "ri") return "#7c3aed";
  return "#4b5563";
};



const vendorBadge = (v="")=>{

  const k = v.toLowerCase();

  if(k.includes("rsg/ri") || k.includes("rsg")){
    return `<span class="badge owner" style="background:#2563eb;border-color:#2563eb;color:#fff;">${v}</span>`;
  }
  if(k.includes("ri")){
    return `<span class="badge owner" style="background:#7c3aed;border-color:#7c3aed;color:#fff;">${v}</span>`;
  }

  const isInternal = k.includes("interne");
  const isExternal = k.includes("externe") || (!isInternal && v);

  const color = isInternal ? "#16a34a" : (isExternal ? "#b45309" : "#4b5563");

  return `<span class="badge owner" style="background:${color};border-color:${color};color:#fff;">${v}</span>`;

};



function refreshVendorsList(){

  const input = el("t_vendor");

  if(!input) return;

  const registry = loadVendorsRegistry();

  const deleted = new Set(loadDeletedVendors().map(x=>x.toLowerCase()));

  const fromTasks = (state?.tasks||[])

    .map(t=> normalizeVendor(t.vendor||""))

    .filter(Boolean)

    .filter(v=> !deleted.has(v.toLowerCase()));

  vendorsCache = dedupVendors([...registry, ...fromTasks])

    .filter(v=> !deleted.has(v.toLowerCase()))

    .sort((a,b)=>a.localeCompare(b,"fr",{sensitivity:"base"}));

  const current = normalizeVendor(input.value);

  if(current && !vendorsCache.map(x=>x.toLowerCase()).includes(current.toLowerCase())){

    vendorsCache.unshift(current);

  }

  saveVendorsRegistry(vendorsCache);

  renderVendorDropdown(current);

}



function refreshDescriptionsList(){

  const input = el("t_room");

  if(!input) return;

  const registry = loadDescriptionsRegistry();

  const deleted = new Set(loadDeletedDescriptions().map(x=>x.toLowerCase()));

  const fromTasks = (state?.tasks||[])

    .map(t=> normalizeDesc(t.roomNumber||""))

    .filter(Boolean)

    .filter(v=> !deleted.has(v.toLowerCase()));

  descCache = dedupDescriptions([...registry, ...fromTasks])

    .filter(v=> !deleted.has(v.toLowerCase()))

    .sort((a,b)=>a.localeCompare(b,"fr",{sensitivity:"base"}));

  const current = normalizeDesc(input.value);

  if(current && !descCache.map(x=>x.toLowerCase()).includes(current.toLowerCase())){

    descCache.unshift(current);

  }

  saveDescriptionsRegistry(descCache);

  renderDescriptionDropdown(current);

}



function setupVendorPicker(){

  const input = el("t_vendor");

  const manageBtn = el("btnManageVendors");

  const panel = el("vendorManagerPanel");

  const wrap = input ? input.closest(".vendor-wrap") : null;

  const card = input ? input.closest(".card") : null;

  const setManagerOpen = (open)=>{

    if(panel) panel.style.display = open ? "block" : "none";

    wrap?.classList.toggle("vendor-open", !!open);

    card?.classList.toggle("vendor-panel-open", !!open);

    if(open && input && panel) openFloating(panel, input);

    else closeFloating(panel);

  };

  if(!input) return;

  const openList = ()=>{

    renderVendorDropdown(input.value);

    showVendorDropdown(true);

  };

  input.addEventListener("click", openList);

  input.addEventListener("focus", openList);

  input.addEventListener("keydown",(e)=>{

    if(e.key==="ArrowDown" || e.key==="F4"){

      e.preventDefault();

      openList();

    }else if(e.key==="Escape"){

      showVendorDropdown(false);

    }

  });

  input.addEventListener("input", ()=>{

    renderVendorDropdown(input.value);

    showVendorDropdown(true);

  });

  document.addEventListener("click",(e)=>{

    const box = el("vendorDropdown");

    if(!box || !input) return;

    if(!box.contains(e.target) && e.target!==input){ showVendorDropdown(false); }

    if(panel && !panel.contains(e.target) && e.target!==manageBtn){ setManagerOpen(false); }

  });

  const box = el("vendorDropdown");

  if(box){

    box.addEventListener("mousedown",(e)=>e.preventDefault()); // empcher blur avant le click

  }

  if(manageBtn){

    manageBtn.disabled = isLocked;

    manageBtn.onclick=(e)=>{

      e.stopPropagation();

      if(isLocked) return;

      const visible = panel?.style.display==="block";

      if(visible){

        setManagerOpen(false);

      }else{

        renderVendorManager();

        setManagerOpen(true);

        showVendorDropdown(false);

      }

    };

  }

}



function setupDescriptionPicker(){

  const input = el("t_room");

  const manageBtn = el("btnManageDescriptions");

  const panel = el("descManagerPanel");

  const wrap = input ? input.closest(".desc-wrap") : null;

  const card = input ? input.closest(".card") : null;

  const setManagerOpen = (open)=>{

    if(panel) panel.style.display = open ? "block" : "none";

    wrap?.classList.toggle("desc-open", !!open);

    card?.classList.toggle("desc-panel-open", !!open);

    if(open && input && panel) openFloating(panel, input);

    else closeFloating(panel);

  };

  if(!input) return;

  const openList = ()=>{

    renderDescriptionDropdown(input.value);

    showDescriptionDropdown(true);

  };

  input.addEventListener("click", openList);

  input.addEventListener("focus", openList);

  input.addEventListener("keydown",(e)=>{

    if(e.key==="ArrowDown" || e.key==="F4"){

      e.preventDefault();

      openList();

    }else if(e.key==="Escape"){

      showDescriptionDropdown(false);

    }

  });

  input.addEventListener("input", ()=>{

    renderDescriptionDropdown(input.value);

    showDescriptionDropdown(true);

  });

  document.addEventListener("click",(e)=>{

    const box = el("descDropdown");

    if(!box || !input) return;

    if(!box.contains(e.target) && e.target!==input){ showDescriptionDropdown(false); }

    if(panel && !panel.contains(e.target) && e.target!==manageBtn){ setManagerOpen(false); }

  });

  const box = el("descDropdown");

  if(box){

    box.addEventListener("mousedown",(e)=>e.preventDefault()); // empcher blur avant le click

  }

  if(manageBtn){

    manageBtn.disabled = isLocked;

    manageBtn.onclick=(e)=>{

      e.stopPropagation();

      if(isLocked) return;

      const visible = panel?.style.display==="block";

      if(visible){

        setManagerOpen(false);

      }else{

        renderDescriptionManager();

        setManagerOpen(true);

        showDescriptionDropdown(false);

      }

    };

  }

}



function renderVendorDropdown(filter=""){

  const box = el("vendorDropdown");

  const input = el("t_vendor");

  if(!box || !input) return;

  // s'assurer que les prestataires supprims ne rapparaissent pas

  const deleted = new Set(loadDeletedVendors().map(x=>x.toLowerCase()));

  const q = normalizeVendor(filter||"").toLowerCase();

  const list = vendorsCache

    .filter(v=> !deleted.has(v.toLowerCase()))

    .filter(v=>!q || v.toLowerCase().includes(q))

    .slice(0,50);

  if(list.length===0){

    box.innerHTML = `<div class="vendor-empty">Aucun résultat</div>`;

  }else{

    box.innerHTML = list.map(v=>`<div class="vendor-item" role="option" tabindex="0">${attrEscape(v)}</div>`).join("");

  }

  box.querySelectorAll(".vendor-item").forEach(item=>{

    item.onclick=()=>{ input.value=item.textContent; showVendorDropdown(false); };

    item.onkeydown=(e)=>{ if(e.key==="Enter"){ input.value=item.textContent; showVendorDropdown(false); } };

  });

  showVendorDropdown(list.length>0 && document.activeElement===input);

}



function showVendorDropdown(show){

  const box = el("vendorDropdown");

  if(!box) return;

  box.style.display = show ? "block" : "none";

  box.classList.toggle("open", !!show);

  const input = el("t_vendor");

  if(show && input) openFloating(box, input);

  else closeFloating(box);

}



function renderVendorManager(){

  const panel = el("vendorManagerPanel");

  if(!panel) return;

  if(vendorsCache.length===0){

    panel.innerHTML = `<div class="vendor-empty">Aucun prestataire enregistré</div>`;

    return;

  }

  panel.innerHTML = dedupVendors(vendorsCache).map(v=>`

    <div class="vendor-row">

      <span class="vendor-name">${attrEscape(v)}</span>

      <div class="vendor-actions">

        <button class="btn btn-ghost vendor-rename" data-v="${attrEscape(v)}">Renommer</button>

        <button class="btn btn-danger vendor-delete" data-v="${attrEscape(v)}">Supprimer</button>

      </div>

    </div>

  `).join("");

  panel.querySelectorAll(".vendor-rename").forEach(btn=>{

    btn.onclick=()=>{

      const oldName = btn.dataset.v || "";

      const newName = prompt("Nouveau nom du prestataire :", oldName) || "";

      const trimmed = newName.trim();

      if(!trimmed) return;

      // retirer de la liste des supprims si prsent

      let deleted = loadDeletedVendors().filter(x=> x.toLowerCase()!==oldName.toLowerCase());

      saveDeletedVendors(deleted);

      vendorsCache = vendorsCache.map(x=> x.toLowerCase()===oldName.toLowerCase() ? trimmed : x);

      vendorsCache = dedupVendors(vendorsCache);

      vendorsCache = Array.from(new Set(vendorsCache)).sort((a,b)=>a.localeCompare(b,"fr",{sensitivity:"base"}));

      saveVendorsRegistry(vendorsCache);

      // mettre  jour toutes les tâches utilisant l'ancien nom

      (state?.tasks||[]).forEach(t=>{

        if((t.vendor||"").toLowerCase()===oldName.toLowerCase()) t.vendor = trimmed;

      });

      renderVendorDropdown(el("t_vendor")?.value||"");

      renderVendorManager();

      renderAll();

    };

  });

  panel.querySelectorAll(".vendor-delete").forEach(btn=>{

    btn.onclick=()=>{

      const name = btn.dataset.v || "";

      if(!name) return;

      if(!confirm(`Supprimer le prestataire "${name}" ?`)) return;

      vendorsCache = vendorsCache.filter(x=>x.toLowerCase()!==name.toLowerCase());

      vendorsCache = dedupVendors(vendorsCache);

      saveVendorsRegistry(vendorsCache);

      // ajouter  la liste des supprims pour filtrage futur

      const deleted = Array.from(new Set([...loadDeletedVendors(), name]));

      saveDeletedVendors(deleted);

      // retirer ce prestataire des tâches existantes

      (state?.tasks||[]).forEach(t=>{

        if((t.vendor||"").toLowerCase()===name.toLowerCase()) t.vendor = "";

      });

      saveState();

      renderVendorDropdown(el("t_vendor")?.value||"");

      renderVendorManager();

      renderAll();

    };

  });

}



function renderDescriptionDropdown(filter=""){

  const box = el("descDropdown");

  const input = el("t_room");

  if(!box || !input) return;

  const deleted = new Set(loadDeletedDescriptions().map(x=>x.toLowerCase()));

  const q = normalizeDesc(filter||"").toLowerCase();

  const list = descCache

    .filter(v=> !deleted.has(v.toLowerCase()))

    .filter(v=>!q || v.toLowerCase().includes(q))

    .slice(0,50);

  if(list.length===0){

    box.innerHTML = `<div class="vendor-empty">Aucun résultat</div>`;

  }else{

    box.innerHTML = list.map(v=>`<div class="vendor-item" role="option" tabindex="0">${attrEscape(v)}</div>`).join("");

  }

  box.querySelectorAll(".vendor-item").forEach(item=>{

    item.onclick=()=>{ input.value=item.textContent; showDescriptionDropdown(false); };

    item.onkeydown=(e)=>{ if(e.key==="Enter"){ input.value=item.textContent; showDescriptionDropdown(false); } };

  });

  showDescriptionDropdown(list.length>0 && document.activeElement===input);

}



function showDescriptionDropdown(show){

  const box = el("descDropdown");

  if(!box) return;

  const input = el("t_room");

  const wrap = input ? input.closest(".desc-wrap") : null;

  const card = input ? input.closest(".card") : null;

  box.style.display = show ? "block" : "none";

  box.classList.toggle("open", !!show);

  wrap?.classList.toggle("desc-open", !!show);

  card?.classList.toggle("desc-panel-open", !!show);

  if(show && input) openFloating(box, input);

  else closeFloating(box);

}



function renderDescriptionManager(){

  const panel = el("descManagerPanel");

  if(!panel) return;

  if(descCache.length===0){

    panel.innerHTML = `<div class="vendor-empty">Aucune description enregistrée</div>`;

    return;

  }

  panel.innerHTML = dedupDescriptions(descCache).map(v=>`

    <div class="vendor-row">

      <span class="vendor-name">${attrEscape(v)}</span>

      <div class="vendor-actions">

        <button class="btn btn-ghost desc-rename" data-v="${attrEscape(v)}">Renommer</button>

        <button class="btn btn-danger desc-delete" data-v="${attrEscape(v)}">Supprimer</button>

      </div>

    </div>

  `).join("");

  panel.querySelectorAll(".desc-rename").forEach(btn=>{

    btn.onclick=()=>{

      const oldName = btn.dataset.v || "";

      const newName = prompt("Nouvelle description :", oldName) || "";

      const trimmed = newName.trim();

      if(!trimmed) return;

      let deleted = loadDeletedDescriptions().filter(x=> x.toLowerCase()!==oldName.toLowerCase());

      saveDeletedDescriptions(deleted);

      descCache = descCache.map(x=> x.toLowerCase()===oldName.toLowerCase() ? trimmed : x);

      descCache = dedupDescriptions(descCache);

      descCache = Array.from(new Set(descCache)).sort((a,b)=>a.localeCompare(b,"fr",{sensitivity:"base"}));

      saveDescriptionsRegistry(descCache);

      (state?.tasks||[]).forEach(t=>{

        if((t.roomNumber||"").toLowerCase()===oldName.toLowerCase()) t.roomNumber = trimmed;

      });

      markDirty();

      renderDescriptionDropdown(el("t_room")?.value||"");

      renderProject();

    };

  });

  panel.querySelectorAll(".desc-delete").forEach(btn=>{

    btn.onclick=()=>{

      const name = btn.dataset.v || "";

      if(!name) return;

      if(!confirm(`Supprimer "${name}" ?`)) return;

      descCache = descCache.filter(x=>x.toLowerCase()!==name.toLowerCase());

      descCache = dedupDescriptions(descCache);

      saveDescriptionsRegistry(descCache);

      const deleted = Array.from(new Set([...loadDeletedDescriptions(), name]));

      saveDeletedDescriptions(deleted);

      (state?.tasks||[]).forEach(t=>{

        if((t.roomNumber||"").toLowerCase()===name.toLowerCase()) t.roomNumber = "";

      });

      markDirty();

      renderDescriptionDropdown(el("t_room")?.value||"");

      renderProject();

    };

  });

}



function normalizeState(raw){

  if(!raw){
    const base = defaultState();
    if(!Array.isArray(base.timeLogs)) base.timeLogs = [];
    const orphans = detectOrphanTimeLogs(base);
    base.orphanTimeLogs = orphans;
    if(orphans.length > 0){
      console.warn("[INTEGRITY] Orphan timeLogs detected:", orphans.length);
    }
    return base;
  }

  const normalizeStatus = (s)=> (s||"").split(",").filter(Boolean).map(v=>{
    const up = String(v || "").toUpperCase();
    if(up==="PREPA") return "PREPARATION";
    if(up==="HUIS_SER") return "HUISSERIES";
    return up;
  }).join(",");

  const normProjects = (raw.projects||[]).map(p=>({...p, id:normId(p.id)}));

  const normTasks = (raw.tasks||[]).map(t=>{
    const ownerNorm = (String(t.owner||"").toUpperCase()==="RSG/RI") ? "RSG" : (t.owner||"");
    let vendorNorm = (t.vendor||"").toString().trim();
    if(String(ownerNorm).toLowerCase().includes("prestataire externe") && !vendorNorm){
      vendorNorm = "PRESTATAIRE NON RENSEIGNE";
    }
    return {
      ...t,
      projectId:normId(t.projectId),
      status: normalizeStatus(t.status),
      owner: ownerNorm,
      vendor: vendorNorm
    };
  });

  const normLogs = (raw.timeLogs||[]).map(l=>({
    id: l.id || uid(),
    taskId: normId(l.taskId),
    projectId: normId(l.projectId),
    userKey: (l.userKey || "").toString(),
    userName: (l.userName || "").toString(),
    userEmail: (l.userEmail || "").toString(),
    role: (l.role || "").toString(),
    date: (l.date || "").toString().slice(0,10),
    minutes: Number.isFinite(+l.minutes) ? Math.max(0, Math.round(+l.minutes)) : 0,
    note: (l.note || "").toString(),
    createdAt: l.createdAt || "",
    updatedAt: l.updatedAt || ""
  })).filter(l=>l.taskId && l.date);

  const deleted = new Set(loadDeletedVendors().map(x=>x.toLowerCase()));
  normTasks.forEach(t=>{
    if(t.vendor && deleted.has(t.vendor.toLowerCase())) t.vendor = "";
  });

  const tasksById = new Map(normTasks.map(t=>[t.id, t]));
  const keptLogs = [];
  const orphanBuffer = [];
  normLogs.forEach(l=>{
    const task = tasksById.get(l.taskId);
    if(!task){
      orphanBuffer.push(l);
      return;
    }
    const fixed = {...l};
    if(task.projectId) fixed.projectId = task.projectId;
    if(task.start && fixed.date < task.start){
      orphanBuffer.push(fixed);
      return;
    }
    if(task.end && fixed.date > task.end){
      orphanBuffer.push(fixed);
      return;
    }
    const expectedRole = getTaskRoleKey(task);
    const actualRole = normalizeTimeLogRole(fixed);
    if(expectedRole && actualRole !== expectedRole){
      fixed.roleKey = expectedRole;
      fixed.role = String(expectedRole).toUpperCase();
    }
    keptLogs.push(fixed);
  });

  const dedupMap = new Map();
  keptLogs.forEach(l=>{
    const rk = normalizeTimeLogRole(l);
    const key = `${l.taskId}|${l.date}|${rk}`;
    const prev = dedupMap.get(key);
    if(!prev){
      dedupMap.set(key, l);
      return;
    }
    const prevTs = new Date(prev.updatedAt || prev.createdAt || 0).getTime();
    const curTs = new Date(l.updatedAt || l.createdAt || 0).getTime();
    if(curTs >= prevTs) dedupMap.set(key, l);
  });

  const state = {projects:normProjects, tasks:normTasks, ui: raw.ui||{}, timeLogs: Array.from(dedupMap.values())};
  const detectedOrphans = detectOrphanTimeLogs(state);
  const archivedOrphans = Array.isArray(raw.orphanTimeLogs) ? raw.orphanTimeLogs : [];
  const orphans = [...archivedOrphans, ...orphanBuffer, ...detectedOrphans];
  state.orphanTimeLogs = orphans;
  if(detectedOrphans.length > 0){
    console.warn("[INTEGRITY] Orphan timeLogs detected:", detectedOrphans.length);
  }
  return state;

}
function detectOrphanTimeLogs(state) {
  const taskIds = new Set((state.tasks || []).map(t => t.id));
  const projectIds = new Set((state.projects || []).map(p => p.id));

  const orphans = [];

  for (const log of (state.timeLogs || [])) {
    if (!taskIds.has(log.taskId) || !projectIds.has(log.projectId)) {
      orphans.push(log);
    }
  }

  return orphans;
}

const formatDate = (s)=>{

  if(!s) return "";

  const parts = s.split("-");

  if(parts.length!==3) return s;

  return `${parts[2]}/${parts[1]}/${parts[0]}`;

};

const unformatDate = (fr)=>{

  if(!fr) return "";

  const parts = fr.split("/");

  if(parts.length!==3) return fr;

  const [jj,mm,aa] = parts;

  return `${aa}-${mm}-${jj}`;

};

function toInputDate(val){
  if(!val) return "";
  if(val instanceof Date){
    return val.toISOString().slice(0,10);
  }
  const s = String(val).trim();
  if(!s) return "";
  if(/^\d{4}-\d{2}-\d{2}$/.test(s)) return s;
  const m = s.match(/^(\d{2})\/(\d{2})\/(\d{4})$/);
  if(m) return `${m[3]}-${m[2]}-${m[1]}`;
  return s;
}

function toLocalDateKey(d){
  const y = d.getFullYear();
  const m = String(d.getMonth()+1).padStart(2,"0");
  const day = String(d.getDate()).padStart(2,"0");
  return `${y}-${m}-${day}`;
}
function getYesterdayKey(){
  const d = new Date();
  d.setDate(d.getDate()-1);
  return toLocalDateKey(d);
}
function getSelectedLogDate(){
  const input = el("t_time_date_input");
  const raw = (input?.value || "").trim();
  if(raw) return raw;
  return getYesterdayKey();
}
function isTaskActiveOn(t, dateKey){
  if(!t || !t.start || !t.end || !dateKey) return false;
  return t.start <= dateKey && t.end >= dateKey;
}



// -------- Sélection unique Statuts / Corps d'état --------

function buildStatusMenu(){

  const menu = el("t_status_menu");

  if(!menu) return;

  let h="";

  sortedStatuses().forEach(s=>{
    h+=`<div class="ms-item" data-v="${s.v}">
          <span class="ms-label">${s.label}</span>
        </div>`;
  });

  menu.innerHTML=h;

  menu.querySelectorAll(".ms-item").forEach(item=>{

    item.onclick=(e)=>{

      e.stopPropagation();

      const v=item.dataset.v;

      if(selectedStatusSet.has(v)) selectedStatusSet.clear();

      else{
        selectedStatusSet.clear();
        selectedStatusSet.add(v);
      }

      updateStatusDisplay();
      toggleStatusMenu(false);

    };

  });

}



function updateStatusDisplay(){

  const display = el("t_status_display");

  const menu = el("t_status_menu");

  if(!display || !menu) return;

  // visuels des coche

  menu.querySelectorAll(".ms-item").forEach(item=>{

    const v=item.dataset.v;

    if(selectedStatusSet.has(v)) item.classList.add("selected");

    else item.classList.remove("selected");

  });

  if(selectedStatusSet.size===0){

    display.textContent="Sélectionner";

  }else{

    const labels = STATUSES.filter(s=>selectedStatusSet.has(s.v)).map(s=>s.label);

    display.textContent = labels.join(", ");

  }

}



function setStatusSelection(values){

  const first = (values||"").split(",").filter(Boolean)[0] || "";
  selectedStatusSet = new Set(first ? [first] : []);

  updateStatusDisplay();

}



function toggleStatusMenu(show){

  const menu = el("t_status_menu");

  if(!menu) return;

  const shouldShow = show!==undefined ? show : menu.classList.contains("hidden");

  if(shouldShow) menu.classList.remove("hidden");

  else menu.classList.add("hidden");

}



const EMBEDDED_BACKUP = {

  projects: [

    { id:"3e86100919c04fb8456", name:"Bureau Pastorale", site:"CDM", constraints:"", subproject:"Rénovation" },

    { id:"0c644af019c05700845", name:"Internat Saint Gervais", site:"LGT", constraints:"", subproject:"Rénovation CH 011" }

  ],

  tasks: [

    { id:"c807465d19c05012673", projectId:"3e86100919c04fb8456", roomNumber:"Rénovation", status:"ELECTRICITE,PEINTURE,MOBILIER,AMENAGEMENTS", owner:"INTERNE", start:"2026-02-02", end:"2026-02-27", notes:"" },

    { id:"f490f0e019c0571bee2", projectId:"0c644af019c05700845", roomNumber:"Rénovation CH 011", status:"PEINTURE,TDV,AMENAGEMENTS", owner:"INTERNE", start:"2026-02-02", end:"2026-02-14", notes:"" },

    { id:"840b3cb519c05732884", projectId:"0c644af019c05700845", roomNumber:"Rénovation CH 010", status:"PEINTURE,AMENAGEMENTS,TDV", owner:"INTERNE", start:"2026-02-16", end:"2026-02-28", notes:"" }

  ],

  ui: { activeTab:"3e86100919c04fb8456", filters:{} }

};



function defaultState(){

  return deepClone(EMBEDDED_BACKUP);

}



function loadVendorsRegistry(){

  try{

    const raw = localStorage.getItem(VENDOR_STORE_KEY);

    if(!raw) return [];

    const arr = JSON.parse(raw);

    if(Array.isArray(arr)) return dedupVendors(arr);

    return [];

  }catch(e){

    console.warn("Unable to load vendor registry", e);

    return [];

  }

}



function loadDeletedVendors(){

  try{

    const raw = localStorage.getItem(VENDOR_DELETE_KEY);

    if(!raw) return [];

    const arr = JSON.parse(raw);

    if(Array.isArray(arr)) return arr.filter(Boolean);

    return [];

  }catch(e){

    console.warn("Unable to load deleted vendors", e);

    return [];

  }

}



function saveDeletedVendors(list){

  try{

    localStorage.setItem(VENDOR_DELETE_KEY, JSON.stringify(list));

  }catch(e){

    console.warn("Unable to save deleted vendors", e);

  }

}



function saveVendorsRegistry(list){

  try{

    localStorage.setItem(VENDOR_STORE_KEY, JSON.stringify(dedupVendors(list)));

  }catch(e){

    console.warn("Unable to save vendor registry", e);

  }

}



function loadDescriptionsRegistry(){

  try{

    const raw = localStorage.getItem(DESC_STORE_KEY);

    if(!raw) return [];

    const arr = JSON.parse(raw);

    if(Array.isArray(arr)) return dedupDescriptions(arr);

    return [];

  }catch(e){

    console.warn("Unable to load descriptions registry", e);

    return [];

  }

}



function loadDeletedDescriptions(){

  try{

    const raw = localStorage.getItem(DESC_DELETE_KEY);

    if(!raw) return [];

    const arr = JSON.parse(raw);

    if(Array.isArray(arr)) return arr.filter(Boolean);

    return [];

  }catch(e){

    console.warn("Unable to load deleted descriptions", e);

    return [];

  }

}



function saveDeletedDescriptions(list){

  try{

    localStorage.setItem(DESC_DELETE_KEY, JSON.stringify(list));

  }catch(e){

    console.warn("Unable to save deleted descriptions", e);

  }

}



function saveDescriptionsRegistry(list){

  try{

    localStorage.setItem(DESC_STORE_KEY, JSON.stringify(dedupDescriptions(list)));

  }catch(e){

    console.warn("Unable to save descriptions registry", e);

  }

}



function load(){

  const skipFileFetch = (window.location && window.location.protocol === "file:");

  const backupPromise = skipFileFetch

    ? Promise.reject("skip-file-fetch")

    : fetch(`suivi_chantiers_backup.json?v=${Date.now()}`, {cache:"no-store"});

  // 1) tenter le fichier de backup du projet (persistant disque)

  backupPromise

    .then(resp=> resp.ok ? resp.json() : null)

    .then(data=>{

      if(data){

        state = normalizeState(data);

        renderAll();

        clearDirty();

        _scheduleSupabaseAutoLoad();

        return;

      }

      // 2) sinon tenter le localStorage

      try{

        const raw = localStorage.getItem(STORAGE_KEY);

        if(raw){

          state = normalizeState(JSON.parse(raw));

          renderAll();

          clearDirty();

          _scheduleSupabaseAutoLoad();

          return;

        }

      }catch(e){ softCatch(e); }

      // 3) fallback état embarqu

      state = normalizeState(defaultState());

      renderAll();

      clearDirty();

      _scheduleSupabaseAutoLoad();

    })

    .catch(()=>{

      // si fetch choue, on tente localStorage puis default

      try{

        const raw = localStorage.getItem(STORAGE_KEY);

        if(raw){

          state = normalizeState(JSON.parse(raw));

          renderAll();

          clearDirty();

          _scheduleSupabaseAutoLoad();

          return;

        }

      }catch(e){ softCatch(e); }

      state = normalizeState(defaultState());

      renderAll();

      clearDirty();

      _scheduleSupabaseAutoLoad();

    });

}



let _suppressSupabaseSave = false;
let _saveToastTimer = null;

function showSaveToast(type, title, detail){
  const toast = el("saveToast");
  if(!toast) return;
  const icon = el("saveToastIcon");
  const titleEl = el("saveToastTitle");
  const detailEl = el("saveToastDetail");
  toast.classList.remove("is-error");
  if(type === "error") toast.classList.add("is-error");
  if(icon) icon.textContent = (type === "error") ? "ERR" : "OK";
  if(titleEl) titleEl.textContent = title || "Sauvegarde";
  if(detailEl) detailEl.style.color = "";
  if(detailEl) detailEl.textContent = detail || "";
  const bar = el("saveToastProgressBar");
  const duration = (type === "ok") ? 3000 : 4000;
  if(bar){
    bar.style.animation = "none";
    void bar.offsetWidth;
    bar.style.animation = `saveToastCountdown ${duration}ms linear forwards`;
  }
  toast.classList.add("show");
  if(_saveToastTimer) clearTimeout(_saveToastTimer);
  _saveToastTimer = setTimeout(()=> toast.classList.remove("show"), duration);
}

let _lastDataQualityReport = null;

function collectDataQualityIssues(currentState=state){
  const s = currentState || {};
  const tasks = Array.isArray(s.tasks) ? s.tasks : [];
  const logs = Array.isArray(s.timeLogs) ? s.timeLogs : [];
  const taskById = new Map(tasks.map(t=>[t.id, t]));

  let invalidDates = 0;
  let externalWithoutVendor = 0;
  let legacyStatus = 0;
  let orphanLogs = 0;
  let logsOutsideTaskRange = 0;

  tasks.forEach(t=>{
    const start = (t?.start || "").toString();
    const end = (t?.end || "").toString();
    if(!start || !end || end < start) invalidDates += 1;

    const owner = (t?.owner || "").toString().toLowerCase();
    const vendor = (t?.vendor || "").toString().trim();
    if(owner.includes("prestataire externe") && !vendor) externalWithoutVendor += 1;

    const statuses = String(t?.status || "").split(",").map(x=>x.trim().toUpperCase()).filter(Boolean);
    if(statuses.includes("HUIS_SER")) legacyStatus += 1;
  });

  logs.forEach(l=>{
    const task = taskById.get(l?.taskId);
    if(!task){ orphanLogs += 1; return; }
    const d = (l?.date || "").toString().slice(0,10);
    if(!d) return;
    const ts = (task?.start || "").toString();
    const te = (task?.end || "").toString();
    if(ts && d < ts) logsOutsideTaskRange += 1;
    else if(te && d > te) logsOutsideTaskRange += 1;
  });

  const issues = [];
  if(invalidDates > 0) issues.push(`${invalidDates} tâche(s) avec dates invalides`);
  if(externalWithoutVendor > 0) issues.push(`${externalWithoutVendor} tâche(s) externes sans prestataire`);
  if(legacyStatus > 0) issues.push(`${legacyStatus} tâche(s) en statut obsolète HUIS_SER`);
  if(orphanLogs > 0) issues.push(`${orphanLogs} log(s) orphelins`);
  if(logsOutsideTaskRange > 0) issues.push(`${logsOutsideTaskRange} log(s) hors période de tâche`);

  return {
    ok: issues.length === 0,
    issues,
    counts: { invalidDates, externalWithoutVendor, legacyStatus, orphanLogs, logsOutsideTaskRange }
  };
}

function formatQualityIssuesForToast(report){
  if(!report || report.ok) return "Aucune incohérence métier détectée";
  const base = report.issues.slice(0,3).join(" | ");
  if(report.issues.length <= 3) return base;
  return `${base} | +${report.issues.length - 3} autre(s)`;
}

function updateDataQualityBanner(notify=false){
  const brandSub = el("brandSub");
  if(!brandSub) return;

  const today = new Date();
  const fmt = today.toLocaleDateString("fr-FR",{weekday:"long", day:"2-digit", month:"long", year:"numeric"});
  const report = collectDataQualityIssues(state);
  _lastDataQualityReport = report;

  const badgeLabel = report.ok ? "Qualité données: OK" : `Qualité données: ${report.issues.length} incohérence(s)`;
  const badgeStyle = report.ok
    ? "color:#16a34a;border:1px solid #16a34a33;background:#16a34a14;padding:2px 8px;border-radius:10px;cursor:pointer;"
    : "color:#b91c1c;border:1px solid #b91c1c33;background:#b91c1c14;padding:2px 8px;border-radius:10px;cursor:pointer;";

  brandSub.innerHTML = `Tableau maître  Projets  Gantt  <span class="brand-date">${fmt}</span>  <span id="dataQualityBadge" style="${badgeStyle}">${badgeLabel}</span>`;
  const badge = el("dataQualityBadge");
  if(badge){
    badge.onclick = ()=>{
      const r = _lastDataQualityReport || collectDataQualityIssues(state);
      showSaveToast(r.ok ? "ok" : "error", "Contrôle qualité", formatQualityIssuesForToast(r));
    };
  }

  if(notify && !report.ok){
    showSaveToast("error", "Contrôle qualité", formatQualityIssuesForToast(report));
  }
}
function applyDataQualityCleanup(){
  const before = collectDataQualityIssues(state);
  const cleaned = normalizeState(deepClone(state));
  const removedFromLogs = Math.max(0, (state?.timeLogs?.length || 0) - (cleaned?.timeLogs?.length || 0));
  state.projects = cleaned.projects || [];
  state.tasks = cleaned.tasks || [];
  state.timeLogs = cleaned.timeLogs || [];
  state.ui = cleaned.ui || state.ui || {};
  state.orphanTimeLogs = cleaned.orphanTimeLogs || [];

  const after = collectDataQualityIssues(state);
  markDirty();
  renderAll();

  const detail = `Avant: ${before.issues.length} | Après: ${after.issues.length} | Logs retirés: ${removedFromLogs}`;
  showSaveToast(after.ok ? "ok" : "error", "Nettoyage terminé", detail);
  return { before, after, removedFromLogs };
}

function exportDataQualityReportPdf(){
  const report = collectDataQualityIssues(state);
  const today = new Date().toLocaleDateString("fr-FR",{day:"2-digit",month:"short",year:"numeric"});

  setPrintPageFormat("A4 portrait", "6mm");
  document.body.classList.add("print-mode");

  const tpl = ensurePrintTemplate();
  if(!tpl) return;

  let container = document.getElementById("printInjection");
  if(!container){
    container = document.createElement("div");
    container.id = "printInjection";
    document.body.prepend(container);
  }

  container.innerHTML = tpl.innerHTML;
  const header = container.querySelector("#printHeader");
  const meta = container.querySelector("#printMeta");
  const legend = container.querySelector("#printLegend");

  if(header){
    header.querySelector("h1").textContent = "Rapport qualité des données";
  }
  if(meta){
    const rows = [
      ["Date export", today],
      ["État", report.ok ? "OK" : "Incohérences détectées"],
      ["Nombre d'anomalies", String(report.issues.length)],
      ["Tâches", String((state?.tasks || []).length)],
      ["Logs temps", String((state?.timeLogs || []).length)]
    ];
    meta.innerHTML = rows.map(([k,v])=>`<div><strong>${k}</strong><br>${attrEscape(v)}</div>`).join("");
  }
  if(legend) legend.innerHTML = "";

  container.querySelectorAll(".print-dynamic").forEach(n=>n.remove());
  const wrap = document.createElement("div");
  wrap.className = "print-dynamic";
  const card = document.createElement("div");
  card.className = "card print-block";

  const lines = report.ok
    ? ["Aucune incohérence métier détectée."]
    : report.issues;

  const listHtml = lines.map((x,i)=>`<li>${attrEscape(String(i+1))}. ${attrEscape(x)}</li>`).join("");
  card.innerHTML = `
    <div class="card-title">Contrôle qualité</div>
    <div style="padding:10px 14px;">
      <ul style="margin:0;padding-left:18px;line-height:1.5;">${listHtml}</ul>
    </div>
  `;
  wrap.appendChild(card);
  container.querySelector(".print-order")?.appendChild(wrap);

  setTimeout(()=>{
    maximizePrintContainer(container);
    openPreparedPrintInNewWindow();
  }, 0);
}
function animateMetricCounters(root){
  if(!root) return;
  const vals = root.querySelectorAll(".metric-val");
  vals.forEach(node=>{
    const raw = (node.textContent || "").trim();
    const m = raw.match(/^(\d+)(\s*[^\d].*)?$/);
    if(!m) return;
    const end = Number(m[1] || 0);
    const suffix = m[2] || "";
    const start = 0;
    const duration = 520;
    const t0 = performance.now();
    const step = (now)=>{
      const p = Math.min(1, (now - t0) / duration);
      const eased = 1 - Math.pow(1 - p, 3);
      const val = Math.round(start + (end - start) * eased);
      node.textContent = `${val}${suffix}`;
      if(p < 1) requestAnimationFrame(step);
    };
    requestAnimationFrame(step);
  });
}

const _badgeChangeSignatures = new Map();
function animateBadgeChanges(root){
  if(!root) return;
  const nodes = root.querySelectorAll(".num-badge, .badge.owner, .panel-chip .metric-val");
  nodes.forEach((node, idx)=>{
    const key = `${root.id || "root"}:${idx}:${node.className}`;
    const sig = `${(node.textContent || "").trim()}|${node.getAttribute("style") || ""}`;
    const prev = _badgeChangeSignatures.get(key);
    if(prev !== undefined && prev !== sig){
      node.classList.remove("badge-change-pulse");
      void node.offsetWidth;
      node.classList.add("badge-change-pulse");
    }
    _badgeChangeSignatures.set(key, sig);
  });
}

function animateCardsInView(viewId){
  const view = el(viewId);
  if(!view) return;
  const cards = view.querySelectorAll(".card, .tablewrap");
  cards.forEach((node, idx)=>{
    node.style.setProperty("--card-fade-delay", `${Math.min(idx * 0.025, 0.18)}s`);
    node.classList.remove("card-fade-in");
    void node.offsetWidth;
    node.classList.add("card-fade-in");
  });
}

function saveState(opts={}){

  try{

    localStorage.setItem(STORAGE_KEY, JSON.stringify(state));

    clearDirty();

    // Supabase greffe : APRES sauvegarde locale

    const skipSupabase = !!opts.skipSupabase || _suppressSupabaseSave;
    if(!skipSupabase){
      try{ if(window.saveAppStateToSupabase) window.saveAppStateToSupabase(state); }catch(e){ softCatch(e); }
    }

  }catch(e){

    console.warn("save failed", e);

  }

}



function updateSaveButton(){
  const saveButtonIds = [
    "btnSave",
    "btnSaveProject",
    "btnSaveTask",
    "btnConfigSaveTop",
    "btnConfigSave",
    "btnVacSave"
  ];
  const buttons = saveButtonIds.map((id)=>el(id)).filter(Boolean);
  if(!buttons.length) return;
  buttons.forEach((btn)=>{
    btn.classList.remove("btn-danger","btn-success");
    if(unsavedChanges){
      btn.classList.add("btn-primary");
      btn.classList.remove("btn-save-idle");
    }else{
      btn.classList.remove("btn-primary");
      btn.classList.add("btn-save-idle");
    }
  });

}

function markDirty(){
  unsavedChanges = true;
  _stateVersion += 1;
  updateSaveButton();
}
function saveUIState(){
  try{
    if(!state) return;
    state.ui = state.ui || {};
    state.ui.filters = {
      site: el("filterSite")?.value || "",
      project: el("filterProject")?.value || "",
      status: el("filterStatus")?.value || "",
      search: el("filterSearch")?.value || "",
      startAfter: el("filterStartAfter")?.value || "",
      endBefore: el("filterEndBefore")?.value || ""
    };
  }catch(e){ softCatch(e); }
}

function saveUndoSnapshot(){
  try{
    lastUndoSnapshot = {
      projects: deepClone(state.projects || []),
      tasks: deepClone(state.tasks || []),
      timeLogs: deepClone(state.timeLogs || []),
      selectedProjectId,
      selectedTaskId
    };
  }catch(e){
    console.warn("Undo snapshot failed", e);
  }
}

function restoreUndoSnapshot(){
  if(!lastUndoSnapshot) return false;
  state.projects = deepClone(lastUndoSnapshot.projects || []);
  state.tasks = deepClone(lastUndoSnapshot.tasks || []);
  state.timeLogs = deepClone(lastUndoSnapshot.timeLogs || []);
  selectedProjectId = lastUndoSnapshot.selectedProjectId || null;
  selectedTaskId = lastUndoSnapshot.selectedTaskId || null;
  lastUndoSnapshot = null;
  markDirty();
  renderAll();
  return true;
}

function clearDirty(){ unsavedChanges = false; updateSaveButton(); }



function downloadBackup(){

  try{

    const data = JSON.stringify(state, null, 2);

    const blob = new Blob([data], {type:"application/json"});

    const a = document.createElement("a");

    a.href = URL.createObjectURL(blob);

    a.download = "suivi_chantiers_backup.json";

    document.body.appendChild(a);

    a.click();

    a.remove();

  }catch(e){

    console.warn("download backup failed", e);

  }

}



function writeBackupToDisk(){

  try{

    const data = JSON.stringify(state, null, 2);

    if(window.showSaveFilePicker){

      (async ()=>{

        const handle = await window.showSaveFilePicker({

          suggestedName: "suivi_chantiers_backup.json",

          types:[{description:"JSON", accept:{"application/json":[".json"]}}]

        });

        const writable = await handle.createWritable();

        await writable.write(data);

        await writable.close();

      })();

    }

  }catch(e){

    console.warn("writeBackupToDisk failed", e);

  }

}



function flashSaved(){

  const btn = el("btnSave");

  if(!btn) return;

  const old = btn.textContent;

  btn.textContent = " Sauvegard";

  btn.classList.add("pulse");

  setTimeout(()=>{ btn.textContent = old; btn.classList.remove("pulse"); },1200);

}



window.updateRoleUI = updateRoleUI;


function statusLabels(values){

  return parseStatuses(values).map(v=> (STATUSES.find(s=>s.v===v)?.label || v)).join(", ");

}

function toDateInput(d){

  if(!d) return "";

  const x = new Date(d.getTime());

  return x.toISOString().slice(0,10);

}

function parseInputDate(v){

  if(!v) return null;

  const d = new Date(v+"T00:00:00");

  return isNaN(d) ? null : d;

}

function getTasksDateBounds(tasks){

  let min=null, max=null;

  tasks.forEach((t,rowIdx)=>{

    if(!t.start || !t.end) return;

    const s = parseInputDate(t.start);

    const e = parseInputDate(t.end);

    if(!s || !e) return;

    if(!min || s < min) min = s;

    if(!max || e > max) max = e;

  });

  return {min,max};

}

function getWorkloadRange(tasks, boundsTasks=tasks, stateRef=null){

  const {min, max} = getTasksDateBounds(boundsTasks);

  if(!min || !max) return {start:null,end:null};

  const type = (stateRef?.type ?? workloadRangeType) || "all";

  if(type==="all"){

    return {start:min, end:max};

  }

  if(type==="custom"){

    const s = parseInputDate(stateRef?.start ?? workloadRangeStart) || min;

    const e = parseInputDate(stateRef?.end ?? workloadRangeEnd) || max;

    return {start:s, end:e};

  }

  const year = parseInt((stateRef?.year ?? workloadRangeYear) || String(min.getFullYear()),10);

  if(type==="civil"){

    return {start:new Date(year,0,1), end:new Date(year,11,31)};

  }

  // school: 1er sept -> 31 aot

  return {start:new Date(year,8,1), end:new Date(year+1,7,31)};

}

let loginRangeStart = "";
let loginRangeEnd = "";
let loginLogSortKey = "ts";
let loginLogSortDir = "desc";

function toISODateStart(d){
  if(!d) return "";
  const x = new Date(d.getTime());
  x.setHours(0,0,0,0);
  return x.toISOString();
}
function toISODateEnd(d){
  if(!d) return "";
  const x = new Date(d.getTime());
  x.setHours(23,59,59,999);
  return x.toISOString();
}
function toLocalISODate(d){
  if(!d || isNaN(d)) return "";
  const y = d.getFullYear();
  const m = String(d.getMonth()+1).padStart(2,"0");
  const day = String(d.getDate()).padStart(2,"0");
  return `${y}-${m}-${day}`;
}

function hexToRgba(hex, alpha){
  const v = (hex || "").replace("#","").trim();
  if(v.length !== 6) return `rgba(15,23,42,${alpha})`;
  const r = parseInt(v.slice(0,2),16);
  const g = parseInt(v.slice(2,4),16);
  const b = parseInt(v.slice(4,6),16);
  return `rgba(${r},${g},${b},${alpha})`;
}

function clampLoginRange(start, end, maxDays=35){
  const s = new Date(start.getTime());
  const e = new Date(end.getTime());
  const diff = Math.floor((e - s) / 86400000);
  if(diff <= maxDays) return {start:s, end:e, clamped:false};
  const ns = new Date(e.getTime());
  ns.setDate(ns.getDate() - maxDays);
  return {start:ns, end:e, clamped:true};
}

function buildLoginHeatmap(container, events, rangeStart, rangeEnd){
  if(!container) return;
  const start = rangeStart ? new Date(rangeStart+"T00:00:00") : new Date();
  const end = rangeEnd ? new Date(rangeEnd+"T00:00:00") : new Date();
  if(end < start){
    const tmp = new Date(start.getTime());
    start.setTime(end.getTime());
    end.setTime(tmp.getTime());
  }
  const clamp = clampLoginRange(start, end, 34);
  const startKey = toLocalISODate(clamp.start);
  const endKey = toLocalISODate(clamp.end);
  const byUser = {};
  (events || []).forEach(ev=>{
    const ts = ev.ts ? new Date(ev.ts) : null;
    if(!ts || isNaN(ts)) return;
    const key = toLocalISODate(ts);
    const userKey = (ev.email || ev.name || "inconnu").toLowerCase();
    if(!byUser[userKey]) byUser[userKey] = {};
    byUser[userKey][key] = (byUser[userKey][key] || 0) + 1;
  });
  const days = [];
  const cursor = new Date(clamp.start.getTime());
  while(cursor <= clamp.end){
    days.push(new Date(cursor.getTime()));
    cursor.setDate(cursor.getDate()+1);
  }
  const startWeek = startOfWeek(clamp.start);
  const endWeek = startOfWeek(clamp.end);
  const weeks = [];
  for(let w=new Date(startWeek); w<=endWeek; w=addDays(w,7)){
    weeks.push(new Date(w.getTime()));
  }
  const monthLabels = weeks.map((w,i)=>{
    const label = w.toLocaleDateString("fr-FR",{month:"short"});
    const prev = i>0 ? weeks[i-1].toLocaleDateString("fr-FR",{month:"short"}) : "";
    return (label !== prev) ? label : "";
  });

  const userKeys = Object.keys(byUser);
  userKeys.sort((a,b)=>{
    const suma = Object.values(byUser[a] || {}).reduce((x,y)=>x+y,0);
    const sumb = Object.values(byUser[b] || {}).reduce((x,y)=>x+y,0);
    if(sumb !== suma) return sumb - suma;
    return a.localeCompare(b,"fr",{sensitivity:"base"});
  });
  const palette = ["#2563eb","#16a34a","#f59e0b","#db2777","#0ea5b7","#7c3aed","#dc2626","#059669"];
  const userColors = {};
  userKeys.forEach((u,i)=>{ userColors[u] = palette[i % palette.length]; });

  const headerCells = [];
  weeks.forEach(w=>{
    for(let row=0; row<7; row++){
      const d = addDays(w, row);
      const key = toLocalISODate(d);
      const inRange = (key >= startKey && key <= endKey);
      const short = d.toLocaleDateString("fr-FR",{day:"2-digit",month:"2-digit"});
      const hcls = inRange ? "in" : "out";
      headerCells.push(`<span class="login-col-head ${hcls}">${short}</span>`);
    }
  });

  const userRows = userKeys.map(u=>{
    let max = 1;
    Object.values(byUser[u] || {}).forEach(v=>{ if(v>max) max=v; });
    const total = Object.values(byUser[u] || {}).reduce((a,b)=>a+b,0);
    const cells = [];
    weeks.forEach(w=>{
      for(let row=0; row<7; row++){
        const d = addDays(w, row);
        const key = toLocalISODate(d);
        const inRange = (key >= startKey && key <= endKey);
        const count = inRange ? ((byUser[u] && byUser[u][key]) || 0) : 0;
        const level = inRange ? Math.min(4, Math.ceil((count / max) * 4)) : 0;
        const title = inRange ? `${d.toLocaleDateString("fr-FR")}: ${count} connexion(s)` : "";
        const cls = inRange ? `lv${level}` : "out";
        const label = count > 0 ? count : "";
        cells.push(`<span class="login-cell ${cls}" style="background:${hexToRgba(userColors[u], 0.18 + (0.16*level))};" title="${title}">${label}</span>`);
      }
    });
    const label = u.includes("@") ? u : u.charAt(0).toUpperCase()+u.slice(1);
    return `<div class="login-user-row"><div class="login-user-name"><span class="login-user-dot" style="background:${userColors[u]};"></span>${label}</div><div class="login-user-count">${total}</div><div class="login-user-grid-wrap"><div class="login-user-grid">${cells.join("")}</div></div></div>`;
  }).join("");

  const legend = [0, 1, 2, 3, 4];
  container.innerHTML = `
    <div class="login-heatmap">
      <div class="login-months">
        ${monthLabels.map(m=>`<div>${m}</div>`).join("")}
      </div>
      <div class="login-col-head-row">
        <div class="login-col-head-grid">${headerCells.join("")}</div>
      </div>
      <div class="login-users-stack">
        ${userRows || `<div class="login-empty">Aucune connexion dans la période.</div>`}
      </div>
      <div class="login-legend">
        <span>Moins</span>
        ${legend.map((i)=>`<span class="login-cell lv${i}"></span>`).join("")}
        <span>Plus</span>
      </div>
    </div>
  `;
  return {clamped: clamp.clamped, start: clamp.start, end: clamp.end, totalUsers: userKeys.length};
}

async function initLoginJournalUI(){
  const wrap = el("cfg_login_heatmap");
  const startInput = el("cfg_login_start");
  const endInput = el("cfg_login_end");
  const status = el("cfg_login_status");
  const logBox = el("cfg_login_log");
  const logHead = document.querySelector(".login-log-head");
  if(!wrap || !startInput || !endInput) return;
  if(!loginRangeEnd){
    const end = new Date();
    loginRangeEnd = end.toISOString().slice(0,10);
  }
  if(!loginRangeStart){
    const start = new Date();
    start.setDate(start.getDate()-29);
    loginRangeStart = start.toISOString().slice(0,10);
  }
  if(!startInput.value) startInput.value = loginRangeStart;
  if(!endInput.value) endInput.value = loginRangeEnd;
  const events = await loadLoginsFromSupabase(
    toISODateStart(parseInputDate(startInput.value)),
    toISODateEnd(parseInputDate(endInput.value))
  );
  const info = buildLoginHeatmap(wrap, events, startInput.value, endInput.value);
  if(status){
    const err = localStorage.getItem("login_log_last_error");
    const count = (events || []).length;
    if(err){
      status.textContent = `Erreur Supabase: ${err}`;
    }else{
      const extra = info?.clamped ? " (affichage limité à 30 jours)" : "";
      status.textContent = `Connexions: ${count}${extra}`;
    }
    status.classList.toggle("is-error", !!err);
  }
  if(logBox){
    const rows = (events || []).slice();
    const dir = loginLogSortDir === "asc" ? 1 : -1;
    rows.sort((a,b)=>{
      const va = (a[loginLogSortKey] || "").toString().toLowerCase();
      const vb = (b[loginLogSortKey] || "").toString().toLowerCase();
      if(va < vb) return -1 * dir;
      if(va > vb) return 1 * dir;
      return 0;
    });
    const html = rows.slice(0, 200).map((ev, idx)=>{
      const d = ev.ts ? new Date(ev.ts) : null;
      const dateStr = d && !isNaN(d) ? d.toLocaleString("fr-FR") : "";
      const name = ev.name || "";
      const email = ev.email || "";
      const role = (ev.role||"") === "admin" ? "Admin" : "Utilisateur";
      const delay = Math.min(idx * 24, 300);
      return `<div class="login-log-row" style="animation-delay:${delay}ms"><span>${dateStr}</span><span>${name}</span><span>${email}</span><span>${role}</span><span>1</span></div>`;
    }).join("");
    logBox.innerHTML = html || `<div class="login-empty">Aucune connexion dans la période.</div>`;
  }
  if(logHead){
    logHead.querySelectorAll(".login-log-sort").forEach(btn=>{
      btn.classList.toggle("active", btn.dataset.sort === loginLogSortKey);
      btn.classList.toggle("asc", btn.dataset.sort === loginLogSortKey && loginLogSortDir === "asc");
      btn.classList.toggle("desc", btn.dataset.sort === loginLogSortKey && loginLogSortDir === "desc");
    });
  }
}


function getMasterGanttExportRange(tasksAllOverride=null){
  const tasks = (tasksAllOverride || filteredTasks()).filter(t=>t.start && t.end);
  const {min, max} = getTasksDateBounds(tasks);
  if(!min || !max) return null;
  const typeNode = el("ganttExportRangeType");
  const yearNode = el("ganttExportRangeYear");
  const startNode = el("ganttExportRangeStart");
  const endNode = el("ganttExportRangeEnd");
  const type = typeNode ? typeNode.value : "all";
  if(type === "all"){
    const start = new Date(min.getFullYear(), 0, 1);
    const end = new Date(max.getFullYear(), 11, 31);
    end.setHours(23,59,59,999);
    return {start, end};
  }
  if(type === "custom"){
    let s = parseInputDate(startNode?.value) || min;
    let e = parseInputDate(endNode?.value) || max;
    if(e < s){ const tmp = s; s = e; e = tmp; }
    return {start:s, end:e};
  }
  let year = parseInt(yearNode?.value || String(min.getFullYear()),10);
  if(isNaN(year)){
    year = min.getFullYear();
  }
  const endOfDay = (d)=>{ const x=new Date(d.getTime()); x.setHours(23,59,59,999); return x; };
  if(type === "civil"){
    return {start:new Date(year,0,1), end:endOfDay(new Date(year,11,31))};
  }
  // school year: 1er sept -> 31 aout (annee suivante)
  const schoolYear = (min.getMonth() >= 8) ? min.getFullYear() : (min.getFullYear() - 1);
  if(isNaN(year)) year = schoolYear;
  return {start:new Date(year,8,1), end:endOfDay(new Date(year+1,7,31))};
}

function initGanttExportRangeUI(tasksAllOverride=null){
  const tasks = (tasksAllOverride || filteredTasks()).filter(t=>t.start && t.end);
  const {min, max} = getTasksDateBounds(tasks);
  const typeNode = el("ganttExportRangeType");
  const yearNode = el("ganttExportRangeYear");
  const startNode = el("ganttExportRangeStart");
  const endNode = el("ganttExportRangeEnd");
  if(!typeNode || !yearNode || !startNode || !endNode) return;
  if(!min || !max){
    startNode.value = "";
    endNode.value = "";
    yearNode.innerHTML = "";
    return;
  }
  const minYear = min.getFullYear() - 1;
  const maxYear = max.getFullYear() + 1;
  const prevType = yearNode.dataset.rangeType || "";
  let opts = "";
  const isSchool = (typeNode.value === "school");
  for(let y=minYear; y<=maxYear; y++){
    const label = isSchool ? `${y}-${y+1}` : `${y}`;
    opts += `<option value="${y}">${label}</option>`;
  }
  yearNode.innerHTML = opts;
  if(!yearNode.value || prevType !== typeNode.value){
    let defaultYear = min.getFullYear();
    if(typeNode.value === "school"){
      defaultYear = (min.getMonth() >= 8) ? min.getFullYear() : (min.getFullYear() - 1);
    }
    yearNode.value = String(defaultYear);
  }
  yearNode.dataset.rangeType = typeNode.value;
  if(!startNode.value) startNode.value = toDateInput(min);
  if(!endNode.value) endNode.value = toDateInput(max);
  const showDates = (typeNode.value === "custom");
  const showYear = (typeNode.value === "civil" || typeNode.value === "school");
  startNode.style.display = showDates ? "inline-block" : "none";
  endNode.style.display = showDates ? "inline-block" : "none";
  yearNode.style.display = showYear ? "inline-block" : "none";
  renderGanttExportSites();
}

function syncWorkloadFilterUI(tasks, boundsTasks=tasks, uiIds=null, stateRef=null){

  const typeSel = el("workloadRangeType");

  const yearSel = el("workloadRangeYear");

  const startInput = el("workloadRangeStart");

  const endInput = el("workloadRangeEnd");

  const typeNode = uiIds ? el(uiIds.type) : typeSel;

  const yearNode = uiIds ? el(uiIds.year) : yearSel;

  const startNode = uiIds ? el(uiIds.start) : startInput;

  const endNode = uiIds ? el(uiIds.end) : endInput;

  if(!typeNode || !yearNode || !startNode || !endNode) return;

  const {min, max} = getTasksDateBounds(boundsTasks);

  if(!min || !max) return;

  const minYear = min.getFullYear() - 1;

  const maxYear = max.getFullYear() + 1;

  const st = stateRef || {type:workloadRangeType, year:workloadRangeYear, start:workloadRangeStart, end:workloadRangeEnd};
  const type = st.type || "all";
  const prevType = yearNode.dataset.rangeType || "";
  const defaultSchoolYear = (min.getMonth() >= 8) ? min.getFullYear() : (min.getFullYear() - 1);

  if(!st.start) st.start = toDateInput(min);
  if(!st.end) st.end = toDateInput(max);
  if(!st.year || prevType !== type){
    st.year = String(type === "school" ? defaultSchoolYear : min.getFullYear());
  }

  // options années
  let opts="";
  for(let y=minYear; y<=maxYear; y++){
    const label = (type === "school") ? `${y}-${y+1}` : `${y}`;
    opts += `<option value="${y}">${label}</option>`;
  }
  yearNode.innerHTML = opts;
  yearNode.value = st.year;
  yearNode.dataset.rangeType = type;
  typeNode.value = type;

  startNode.value = st.start;

  endNode.value = st.end;

  const showDates = (typeNode.value==="custom");

  const showYear = (typeNode.value==="civil" || typeNode.value==="school");

  startNode.style.display = showDates ? "inline-block" : "none";

  endNode.style.display = showDates ? "inline-block" : "none";

  yearNode.style.display = showYear ? "inline-block" : "none";

  if(stateRef){

    stateRef.type = typeNode.value;

    stateRef.year = yearNode.value;

    stateRef.start = startNode.value;

    stateRef.end = endNode.value;

  }

}

function isWeekday(d){

  const day = d.getDay();

  return day >= 1 && day <= 5; // lundi-vendredi

}

function countWeekdays(start, end){

  if(!start || !end || end < start) return 0;

  let count = 0;

  for(let d=new Date(start); d<=end; d.setDate(d.getDate()+1)){

    if(isWeekday(d)) count += 1;

  }

  return count;

}

function durationDays(start,end){

  if(!start || !end) return "";

  const s=new Date(start+"T00:00:00");

  const e=new Date(end+"T00:00:00");

  if(isNaN(s) || isNaN(e) || e<s) return "";

  const days = countWeekdays(s, e);

  return days>0 ? days : "";

}

function hoursPerDayForOwner(owner){
  const typ = ownerType(owner);
  const h = getHoursConfig();
  if(typ === "rsg") return h.rsg;
  if(typ === "ri") return (h.ri !== undefined ? h.ri : h.rsg);
  if(typ === "externe") return h.external;
  return h.internal;
}

function durationLabelForTask(task){
  if(!task) return "";
  const days = durationDays(task.start, task.end);
  if(!days && days !== 0) return "";
  const totals = getTaskTimeTotals(task);
  const realLabel = `réel ${formatHoursMinutes(totals.totalMinutes || 0)}`;
  return `${days} j (${realLabel})`;
}

function taskTitle(t){

  const p = state?.projects?.find(x=>x.id===t.projectId);

  const projectName = (p?.name||"Projet").trim();

  const sub = (p?.subproject||"").trim();

  const desc = (t.roomNumber||"").trim();

  if(sub && desc) return `${projectName} - ${sub} - ${desc}`;

  if(sub) return `${projectName} - ${sub}`;

  if(desc) return `${projectName} - ${desc}`;

  return projectName;

}



function ganttLaneTitle(t){

  const p = state?.projects?.find(x=>x.id===t.projectId);

  const projectName = (p?.name || "Sans projet").trim() || "Sans projet";

  const desc = (t.roomNumber || "").trim();

  return desc ? `${projectName} - ${desc}` : projectName;

}



function computeTaskOrderMap(){

  const map={};

  state.projects.forEach(p=>{

    const tasks=state.tasks.filter(t=>t.projectId===p.id);

    tasks.sort((a,b)=>{

      const sa = Date.parse(a.start||"9999-12-31");

      const sb = Date.parse(b.start||"9999-12-31");

      if(sa!==sb) return sa-sb;

      const ea = Date.parse(a.end||"9999-12-31");

      const eb = Date.parse(b.end||"9999-12-31");

      if(ea!==eb) return ea-eb;

      return (a.roomNumber||"").localeCompare(b.roomNumber||"");

    });

    tasks.forEach((t,i)=>{ map[t.id]=i+1; });

  });

  taskOrderMap=map;

}



// Gantt helpers

function startOfWeek(d){

  const x=new Date(d.getTime());

  const day=(x.getDay()+6)%7; // lundi=0

  x.setDate(x.getDate()-day);

  x.setHours(0,0,0,0);

  return x;

}

function endOfWorkWeek(d){

  const x=new Date(d.getTime());

  // vendredi = lundi + 4 jours

  x.setDate(x.getDate()+4);

  x.setHours(23,59,59,999);

  return x;

}

function endOfWeek(d){
  const x = startOfWeek(d);
  x.setDate(x.getDate()+6);
  x.setHours(23,59,59,999);
  return x;
}

function addDays(d,n){ const x=new Date(d.getTime()); x.setDate(x.getDate()+n); return x; }

// --- Vacances scolaires (Zone B) par NUMÉROS de semaines ---
// Format : { "2025-2026": [8,9,16,17,28,29,30,31,32,33,34,35,43,44,52] }
const DEFAULT_VACANCES_ZONE_B_WEEKS = {
  "2025-2026": [8,9,16,17,28,29,30,31,32,33,34,35,43,44,52]
};
// --- Vacances internes (entreprise) par NUMÉROS de semaines ---
// Format : { "2025-2026": [9,17,30,31,32,33] }
const DEFAULT_VACANCES_INTERNE_WEEKS = {
  "2025-2026": [9,17,30,31,32,33]
};
let VACANCES_ZONE_B_WEEKS = deepClone(DEFAULT_VACANCES_ZONE_B_WEEKS);
let VACANCES_INTERNE_WEEKS = deepClone(DEFAULT_VACANCES_INTERNE_WEEKS);
function getSchoolYearKey(d){
  const y = d.getFullYear();
  const m = d.getMonth(); // 0=janv
  // année scolaire : septembre -> août
  return (m >= 8) ? `${y}-${y+1}` : `${y-1}-${y}`;
}

function taskTitleProjectView(t){
  const p = state?.projects?.find(x=>x.id===t.projectId);
  const sub = (p?.subproject||"").trim();
  const desc = (t.roomNumber||"").trim();
  if(sub && desc) return `${sub} - ${desc}`;
  if(sub) return sub;
  if(desc) return desc;
  return (p?.name||"Projet").trim();
}
function isVacationWeek(weekStart){
  const info = isoWeekInfo(weekStart);
  const schoolYear = getSchoolYearKey(weekStart);
  const list = VACANCES_ZONE_B_WEEKS[schoolYear] || [];
  return list.includes(info.week);
}
function isInternalVacationWeek(weekStart){
  const info = isoWeekInfo(weekStart);
  const schoolYear = getSchoolYearKey(weekStart);
  const list = VACANCES_INTERNE_WEEKS[schoolYear] || [];
  return list.includes(info.week);
}

function normalizeWeekList(raw){
  if(!raw) return [];
  const nums = (Array.isArray(raw) ? raw : raw.toString().split(/[,; \n]+/g))
    .map(v=>parseInt(v,10))
    .filter(v=>Number.isFinite(v) && v>=1 && v<=53);
  return Array.from(new Set(nums)).sort((a,b)=>a-b);
}
function normalizeVacationMap(map){
  const out = {};
  if(!map || typeof map!=="object") return out;
  Object.keys(map).forEach(k=>{
    const weeks = normalizeWeekList(map[k]);
    if(weeks.length) out[k]=weeks;
  });
  return out;
}
function applyVacationConfig(){
  const cfg = loadConfig();
  const school = normalizeVacationMap(cfg.vacances_school || {});
  const internal = normalizeVacationMap(cfg.vacances_internal || {});
  VACANCES_ZONE_B_WEEKS = Object.assign({}, deepClone(DEFAULT_VACANCES_ZONE_B_WEEKS), school);
  VACANCES_INTERNE_WEEKS = Object.assign({}, deepClone(DEFAULT_VACANCES_INTERNE_WEEKS), internal);
}
function getVacationYears(){
  const years = new Set([
    ...Object.keys(VACANCES_ZONE_B_WEEKS||{}),
    ...Object.keys(VACANCES_INTERNE_WEEKS||{})
  ]);
  return Array.from(years).sort();
}

function overlapDays(aStart,aEnd,bStart,bEnd){

  const start = Math.max(aStart.getTime(), bStart.getTime());

  const end   = Math.min(aEnd.getTime(), bEnd.getTime());

  if(end < start) return 0;

  const diff = (end - start)/(1000*60*60*24);

  return Math.floor(diff)+1;

}

function barGeometry(taskStart, taskEnd, weekStart){

  const weekEnd = addDays(weekStart,4); // semaine ouvre

  const start = taskStart > weekStart ? taskStart : weekStart;

  const end = taskEnd < weekEnd ? taskEnd : weekEnd;

  const days = countWeekdays(start, end);

  if(days<=0) return {days:0,width:0,offset:0};

  let offsetDays = 0;

  if(taskStart > weekStart){

    const offsetEnd = addDays(taskStart, -1);

    const offsetLimit = offsetEnd < weekEnd ? offsetEnd : weekEnd;

    offsetDays = countWeekdays(weekStart, offsetLimit);

  }

  const offsetPct  = Math.min(100, (offsetDays/5)*100);

  let widthPct = (days/5)*100;

  // éviter dépassement au-delà de la cellule

  if(offsetPct + widthPct > 100) widthPct = 100 - offsetPct;

  widthPct = Math.max(12, Math.min(100, widthPct));

  return {days, width:widthPct, offset:offsetPct};

}

function isoWeekInfo(d){

  const date = new Date(Date.UTC(d.getFullYear(), d.getMonth(), d.getDate()));

  const day = date.getUTCDay() || 7;

  date.setUTCDate(date.getUTCDate() + 4 - day);

  const yearStart = new Date(Date.UTC(date.getUTCFullYear(),0,1));

  const weekNo = Math.ceil(((date - yearStart)/86400000 +1)/7);

  return {week:weekNo, year:date.getUTCFullYear()};

}



function weekKey(d){

  const info=isoWeekInfo(d);

  return `${info.year}-S${String(info.week).padStart(2,"0")}`;

}

function keyToLabel(key, mode){

  if(mode==="day"){

    const [y,m,da]=key.split("-");

    return `${da}/${m}`;

  }

  // week

  const parts=key.split("-S");

  if(parts.length===2) return `S${parts[1]}/${String(parts[0]).slice(2)}`;

  return key;

}



function computeWorkloadData(tasks, mode="week", rangeStart=null, rangeEnd=null){

  const map = new Map(); // key -> {internal, external, rsg, ri, total, anchor}

  const ids = new Set((tasks || []).map(t=>t.id));
  const roleByTask = new Map((tasks || []).map(t=>[t.id, getTaskRoleKey(t)]));
  const rangeByTask = new Map((tasks || []).map(t=>[t.id, {start:t.start||"", end:t.end||""}]));
  if(ids.size === 0) return [];

  getCanonicalTimeLogs().forEach(l=>{
    if(!ids.has(l.taskId)) return;
    if(!l.date) return;
    const d = new Date(l.date+"T00:00:00");
    if(isNaN(d)) return;
    const range = rangeByTask.get(l.taskId);
    if(range?.start && l.date < range.start) return;
    if(range?.end && l.date > range.end) return;
    if(rangeStart && d < rangeStart) return;
    if(rangeEnd && d > rangeEnd) return;

    const key = mode==="day" ? l.date : weekKey(d);
    const anchor = mode==="day" ? d.getTime() : startOfWeek(d).getTime();
    if(!map.has(key)) map.set(key,{internal:0,external:0,rsg:0,ri:0,total:0,anchor});
    const slot = map.get(key);

    const roleExpected = roleByTask.get(l.taskId);
    const role = normalizeTimeLogRole(l);
    if(roleExpected && role !== roleExpected) return;
    const hours = (Number(l.minutes||0) / 60);
    if(!hours) return;
    const mult = (typeof roleHoursMultiplier === "function")
      ? roleHoursMultiplier(role)
      : (role === "interne" ? 2 : 1);
    const weightedHours = hours * mult;
    if(role==="rsg") slot.rsg+=weightedHours;
    else if(role==="ri") slot.ri+=weightedHours;
    else if(role==="interne") slot.internal+=weightedHours;
    else slot.external+=weightedHours;
    slot.total = slot.internal + slot.external + slot.rsg + slot.ri;
  });

  const arr = Array.from(map.entries()).map(([key,val])=>({...val,key}));
  arr.sort((a,b)=> a.anchor - b.anchor);
  return arr;

}



function niceMax(v){

  if(v<=5) return 5;

  const pow = Math.pow(10, Math.floor(Math.log10(v)));

  const mul = Math.ceil(v / pow);

  if(mul<=2) return 2*pow;

  if(mul<=5) return 5*pow;

  return 10*pow;

}



function renderGantt(projectId){

  const wrap = el("gantt");

  const legend = el("legend");

  if(legend){

    legend.innerHTML = sortedStatuses().map(s=>{

      const c = STATUS_COLORS[s.v] || "#2563eb";

      return `<span class="legend-item"><span class="legend-dot" style="background:${c};border-color:${c}"></span><span style="color:#111827;font-weight:600;">${s.label}</span></span>`;

    }).join("");

  }

  if(!wrap) return;

  const tasks = state.tasks.filter(t=>t.projectId===projectId && t.start && t.end);
  const missingMap = buildMissingDaysMap(tasks);

  if(tasks.length===0){

    wrap.innerHTML="<div class='gantt-empty'>Aucune tâche date.</div>";

    return;

  }

  const minStart = tasks.map(t=>new Date(t.start+"T00:00:00")).reduce((a,b)=>a<b?a:b);

  const maxEnd   = tasks.map(t=>new Date(t.end+"T00:00:00")).reduce((a,b)=>a>b?a:b);

  const weeks=[];
  for(let w=startOfWeek(minStart); w<=addDays(startOfWeek(maxEnd),0); w=addDays(w,7)) weeks.push(new Date(w));
  const vacWeeks = weeks.map(w=>isVacationWeek(w));
  const internalVacWeeks = weeks.map(w=>isInternalVacationWeek(w));



  // tri pour garder un ordre stable

  tasks.sort((a,b)=>{
    const sa=Date.parse(a.start||"9999-12-31"), sb=Date.parse(b.start||"9999-12-31");
    if(sa!==sb) return sa-sb;
    const ea=Date.parse(a.end||"9999-12-31"), eb=Date.parse(b.end||"9999-12-31");
    if(ea!==eb) return ea-eb;
    return taskTitle(a).localeCompare(taskTitle(b),"fr",{sensitivity:"base"});
  });



  let html="<div class='tablewrap gantt-table'><table class='table' style='--gcol0:70px;--gcol1:200px;--gcol2:120px;--gcol3:120px'>";

  html+="<thead><tr><th class='gantt-col-site' style='width:70px'>Site / Zone</th><th class='gantt-task-col-project gantt-col-task'>Nom</th><th class='gantt-col-vendor' style='width:120px'>Prestataire</th><th class='gantt-col-status' style='width:120px'>Statut</th>";

  weeks.forEach((w,i)=>{

    const info=isoWeekInfo(w);

    const wEnd=endOfWorkWeek(w);

    const range=`${w.toLocaleDateString("fr-FR",{day:"2-digit"})}-${wEnd.toLocaleDateString("fr-FR",{day:"2-digit"})}/${wEnd.toLocaleDateString("fr-FR",{month:"2-digit",year:"2-digit"})}`;

    const weekLabel = `S${String(info.week).padStart(2,"0")}`;

    const mondayLabel = formatShortDateTwoLinesHTML(w);

    const todayClass = isTodayInWeek(w) ? " week-today" : "";
    const vacClass = vacWeeks[i] ? " vac-week" : "";
    const internalVacClass = internalVacWeeks[i] ? " vac-week-internal" : "";
    html+=`<th class="week-cell${todayClass}${vacClass}${internalVacClass}" data-range="${range}" style='width:72px;color:#111827'>${weekLabel}<div class="gantt-week-date">${mondayLabel}</div></th>`;

  });

  html+="</tr></thead><tbody>";



  // 1 ligne par tâche (plus de regroupement)

  tasks.forEach((t,rowIdx)=>{

    const statuses = parseStatuses(t.status).map(v=>v.toUpperCase());

    const mainStatus = statuses[0] || "";

    const color = ownerColor(t.owner);

    const ownerBadges = t.owner ? ownerBadgeForTask(t) : "";

    const vendorBadges = (()=> {

      const set = new Set();

      if(t.vendor) set.add(t.vendor);

      const typ = ownerType(t.owner);
      if(typ === "interne") set.add("INTERNE");
      if(typ === "rsg") set.add("RSG");
      if(typ === "ri") set.add("RI");
      if(typ === "externe" && !t.vendor) set.add("Prestataire non renseigné");

      if(set.size===0) return "<span class='text-muted'></span>";

      return Array.from(set).sort((a,b)=>a.localeCompare(b,"fr",{sensitivity:"base"}))

        .map(v=>vendorBadge(v)).join(" ");

    })();



    const todayKey = new Date().toISOString().slice(0,10);
    const isToday = !!(t.start && t.end && t.start<=todayKey && t.end>=todayKey);
    const isLate = !!(t.end && t.end < todayKey);
    const rowClass = t.id===selectedTaskId ? "gantt-row gantt-row-active" : "gantt-row";
    const rowClassWithToday = `${rowClass}${isToday ? " today-row" : ""}${isLate ? " late-row" : ""}`;
    html+=`<tr class="${rowClassWithToday}" data-task="${t.id}" onclick="openTaskFromGantt('${t.id}')">`;
    const p = state?.projects?.find(x=>x.id===t.projectId);

    const sub = (p?.subproject || "").trim();

    const taskDesc = (t.roomNumber || "").trim();

    const label = [sub, taskDesc].filter(Boolean).join("  ");

    const siteLabel = (p?.site || "").trim();
    html+=`<td class="gantt-col-site">${attrEscape(siteLabel || "")}</td>`;
    const miss = missingMap.get(t.id) || 0;
    const missDot = miss>0 ? `<span class="missing-dot" title="Heures réelles manquantes (${miss} j)"></span>` : "";
    html+=`<td class="gantt-task-col-project gantt-col-task">${missDot}<b><span class="num-badge" style="--badge-color:${color};--badge-text:#fff;">${taskOrderMap[t.id]||""}</span></b> <span class="gantt-task-name">${attrEscape(label)}</span></td>`;

    html+=`<td class="gantt-vendor-cell gantt-col-vendor"><div class="vendor-stack">${vendorBadges}</div></td>`;

    html+=`<td class="gantt-status-cell gantt-col-status"><div class="gantt-status-stack"><div class="status-row"><span>${statusLabels(mainStatus)}</span></div></div></td>`;



    weeks.forEach((w,i)=>{

      const sDate=new Date(t.start+"T00:00:00");

      const eDate=new Date(t.end+"T00:00:00");

      const geo=barGeometry(sDate,eDate,w);

      if(geo.days>0){

        const title = t.vendor ? ` title="Prestataire : ${attrEscape(t.vendor)}"` : "";

        const vacClass = vacWeeks[i] ? " vac-week" : "";
        const internalVacClass = internalVacWeeks[i] ? " vac-week-internal" : "";
        const barDelay = (rowIdx * 0.03 + i * 0.015).toFixed(3);
        html+=`<td class="gantt-cell${vacClass}${internalVacClass}"><div class="gantt-cell-inner"><div class="bar-wrapper"><div class="gantt-bar bar-click" data-task="${t.id}" data-status="${mainStatus}"${title} style="width:${geo.width}%;margin-left:${geo.offset}%;background:${color};border-color:${color};--bar-delay:${barDelay}s"><span class="gantt-days">${geo.days} j</span></div></div></div></td>`;

      }else{

        const vacClass = vacWeeks[i] ? " vac-week" : "";
        const internalVacClass = internalVacWeeks[i] ? " vac-week-internal" : "";
        html+=`<td class="gantt-cell${vacClass}${internalVacClass}"><div class="gantt-cell-inner"><div class="gantt-spacer"></div></div></td>`;

      }

    });



    html+="</tr>";

  });



  html+="</tbody></table></div>";

  wrap.innerHTML=html;

  applyGanttColumnVisibility();

  wrap.querySelectorAll(".bar-click")?.forEach(bar=>{

    bar.onclick=()=>{

      const taskId = bar.dataset.task;

      const task = state.tasks.find(x=>x.id===taskId);

      if(!task) return;

      selectedProjectId = task.projectId;

      selectedTaskId = taskId;

      renderProject();

    };

  });

  wrap.querySelectorAll("tbody tr[data-task] td")?.forEach(td=>{

    td.onclick=(e)=>{

      if(e.target && e.target.closest(".bar-click")) return;

      const row = e.currentTarget?.parentElement;

      if(!row || !row.dataset.task) return;

      const taskId = row.dataset.task;

      const task = state.tasks.find(x=>x.id===taskId);

      if(!task) return;

      selectedProjectId = task.projectId;

      selectedTaskId = taskId;

      renderProject();

    };

  });

}



function renderProjectTasks(projectId){

  const tbody = el("projectTasksTable")?.querySelector("tbody");

  if(!tbody) return;

  const tasks = state.tasks.filter(t=>t.projectId===projectId);

  const sorted = sortTasks(tasks, sortProject);

  if(sorted.length===0){

    tbody.innerHTML="<tr><td colspan='7' class='empty-row'>Aucune tâche</td></tr>";

    return;

  }

  let h="";
  const missingMap = buildMissingDaysMap(sorted);

  sorted.forEach(t=>{

    const statuses = parseStatuses(t.status).map(v=>v.toUpperCase());

    const c = ownerColor(t.owner);

    const ownerBadgeHtml = t.owner ? ownerBadgeForTask(t) : "";
    const durLabel = durationLabelForTask(t);
    const todayKey = new Date().toISOString().slice(0,10);
    const isToday = !!(t.start && t.end && t.start<=todayKey && t.end>=todayKey);
    const isLate = !!(t.end && t.end < todayKey);
    const isSelected = t.id===selectedTaskId;
    const rowClass = `${isSelected ? "row-selected " : ""}${isToday ? "today-row " : ""}${isLate ? "late-row" : ""}`.trim();

    const miss = missingMap.get(t.id) || 0;
    const missDot = miss>0 ? `<span class="missing-dot" title="Heures réelles manquantes (${miss} j)"></span>` : "";
    h+=`<tr class="${rowClass}" data-task="${t.id}">

      <td>${missDot}<span class="num-badge" style="--badge-color:${c};--badge-text:#fff;">${taskOrderMap[t.id]||""}</span></td>

      <td><span class="icon-picto"></span> ${taskTitleProjectView(t)}</td>

      <td class="status-cell"><span class="status-left">${statusDot(statuses[0])}${statusLabels(t.status||"")}</span>${ownerBadgeHtml||""}</td>

      <td>${formatDate(t.start)||""}${isToday ? `<span class="today-dot" title="En cours aujourd'hui"></span>` : ""}</td>

      <td>${formatDate(t.end)||""}</td>

      <td>${taskProgress(t)}%</td>
      <td>${durLabel}</td>

    </tr>`;

  });

  tbody.innerHTML=h;

  tbody.querySelectorAll("tr").forEach(row=>{
    row.onclick=()=>{
      if(!row.dataset.task) return;
      selectTaskInProject(row.dataset.task);
    };
  });
  updateSortIndicators("projectTasksTable", sortProject);

  const pf = el("projectFiltersBadge");

  if(pf){

    const active = (sortProject.key!=="start" || sortProject.dir!=="asc");

    updateBadge(pf, active, "Tri/filtre actif", "Tri par défaut");
    updateResetSortButtonVisual(el("btnResetSortProject"), active);

  }

}



// Nouvelle version : 1 tâche = 1 ligne dans le gantt maître

function buildMasterGanttHTMLForRange(rangeStart=null, rangeEnd=null, tasksOverride=null){
  const tasksAll = (tasksOverride || filteredTasks()).filter(t=>t.start && t.end);
  if(tasksAll.length===0) return "<div class='gantt-empty'>Aucune tâche date.</div>";

  const rs = rangeStart || null;
  const re = rangeEnd || null;
  const tasks = (!rs || !re) ? tasksAll : tasksAll.filter(t=>{
    const s = new Date(t.start+"T00:00:00");
    const e = new Date(t.end+"T00:00:00");
    return e >= rs && s <= re;
  });

  if(tasks.length===0) return "<div class='gantt-empty'>Aucune tâche date.</div>";
  const missingMap = buildMissingDaysMap(tasks);

  const minStart = rs || tasks.map(t=>new Date(t.start+"T00:00:00")).reduce((a,b)=>a<b?a:b);
  const maxEnd   = re || tasks.map(t=>new Date(t.end+"T00:00:00")).reduce((a,b)=>a>b?a:b);
  // Afficher la totalité des semaines concernées par la période choisie
  const displayStart = rs ? startOfWeek(rs) : minStart;
  const displayEnd = re ? endOfWeek(re) : endOfWeek(maxEnd);

  const weeks=[];
  for(let w=startOfWeek(displayStart); w<=addDays(startOfWeek(displayEnd),0); w=addDays(w,7)) weeks.push(new Date(w));
  const vacWeeks = weeks.map(w=>isVacationWeek(w));
  const internalVacWeeks = weeks.map(w=>isInternalVacationWeek(w));

  tasks.sort((a,b)=>{
    const sa=Date.parse(a.start||"9999-12-31"), sb=Date.parse(b.start||"9999-12-31");
    if(sa!==sb) return sa-sb;
    const ea=Date.parse(a.end||"9999-12-31"), eb=Date.parse(b.end||"9999-12-31");
    if(ea!==eb) return ea-eb;
    return taskTitle(a).localeCompare(taskTitle(b),"fr",{sensitivity:"base"});
  });

  const hideVendor = !ganttColVisibility.masterVendor;
  const hideStatus = !ganttColVisibility.masterStatus;
  const tableClass = `table${hideVendor ? " hide-vendor" : ""}${hideStatus ? " hide-status" : ""}`;

  let html=`<div class='tablewrap gantt-table'><table class='${tableClass}' style='--gcol0:70px;--gcol1:120px;--gcol2:90px;--gcol3:90px'>`;
  html+="<thead><tr><th class='gantt-col-site' style='width:70px'>Site / Zone</th><th class='gantt-col-task' style='width:120px'>Tâche</th><th class='gantt-col-vendor' style='width:90px'>Prestataire</th><th class='gantt-col-status' style='width:90px'>Statut</th>";

  weeks.forEach((w,i)=>{
    const info=isoWeekInfo(w);
    const wEnd=endOfWorkWeek(w);
    const range=`${w.toLocaleDateString("fr-FR",{day:"2-digit"})}-${wEnd.toLocaleDateString("fr-FR",{day:"2-digit"})}/${wEnd.toLocaleDateString("fr-FR",{month:"2-digit",year:"2-digit"})}`;
    const weekLabel = `${info.week}`;
    const mondayLabel = formatShortDateTwoLinesHTML(w);
    const todayClass = isTodayInWeek(w) ? " week-today" : "";
    const vacClass = vacWeeks[i] ? " vac-week" : "";
    const internalVacClass = internalVacWeeks[i] ? " vac-week-internal" : "";
    html+=`<th class="week-cell${todayClass}${vacClass}${internalVacClass}" data-range="${range}" style='width:20px;color:#111827'>${weekLabel}<div class="gantt-week-date">${mondayLabel}</div></th>`;
  });

  html+="</tr></thead><tbody>";

  tasks.forEach((t,rowIdx)=>{
    const statuses = parseStatuses(t.status).map(v=>v.toUpperCase());
    const mainStatus = statuses[0] || "";
    const color = ownerColor(t.owner);
    const p = state?.projects?.find(x=>x.id===t.projectId);
    const projectName = (p?.name || "Projet").trim() || "Projet";
    const vendorBadges = (()=>{
      const set = new Set();
      if(t.vendor) set.add(t.vendor);
      const typ = ownerType(t.owner);
      if(typ === "interne") set.add("INTERNE");
      if(typ === "rsg") set.add("RSG");
      if(typ === "ri") set.add("RI");
      if(typ === "externe" && !t.vendor) set.add("Prestataire non renseigné");
      if(set.size===0) return "<span class='text-muted'></span>";
      return Array.from(set).sort((a,b)=>a.localeCompare(b,"fr",{sensitivity:"base"})).map(v=>vendorBadge(v)).join(" ");
    })();

    const todayKey = new Date().toISOString().slice(0,10);
    const isToday = !!(t.start && t.end && t.start<=todayKey && t.end>=todayKey);
    const isLate = !!(t.end && t.end < todayKey);
    const rowClassWithToday = `gantt-row${isToday ? " today-row" : ""}${isLate ? " late-row" : ""}`;
    html+=`<tr class="${rowClassWithToday}" data-task="${t.id}">`;
    html+=`<td class="gantt-col-site">${attrEscape((p?.site || "").trim())}</td>`;
    const miss = missingMap.get(t.id) || 0;
    const missDot = miss>0 ? `<span class="missing-dot" title="Heures réelles manquantes (${miss} j)"></span>` : "";
    html+=`<td class="gantt-col-task">${missDot}<span class="num-badge" style="--badge-color:${color};--badge-text:#fff;">${taskOrderMap[t.id]||""}</span> <span class="gantt-task-name">${attrEscape(projectName)}</span></td>`;
    html+=`<td class="gantt-vendor-cell gantt-col-vendor"><div class="vendor-stack">${vendorBadges}</div></td>`;
    html+=`<td class="gantt-status-cell gantt-col-status"><div class="gantt-status-stack"><div class="status-row"><span>${statusLabels(mainStatus)}</span></div></div></td>`;

    weeks.forEach((w,i)=>{
      const sDate=new Date(t.start+"T00:00:00");
      const eDate=new Date(t.end+"T00:00:00");
      const geo=barGeometry(sDate,eDate,w);
      if(geo.days>0){
        const title = t.vendor ? ` title="Prestataire : ${attrEscape(t.vendor)}"` : "";
        const vacClass = vacWeeks[i] ? " vac-week" : "";
        const internalVacClass = internalVacWeeks[i] ? " vac-week-internal" : "";
        const barDelay = (rowIdx * 0.03 + i * 0.015).toFixed(3);
        html+=`<td class="gantt-cell${vacClass}${internalVacClass}"><div class="gantt-cell-inner"><div class="bar-wrapper"><div class="gantt-bar" data-task="${t.id}" data-status="${mainStatus}"${title} style="width:${geo.width}%;margin-left:${geo.offset}%;background:${color};border-color:${color};--bar-delay:${barDelay}s"><span class="gantt-days">${geo.days} j</span></div></div></div></td>`;
      }else{
        const vacClass = vacWeeks[i] ? " vac-week" : "";
        const internalVacClass = internalVacWeeks[i] ? " vac-week-internal" : "";
        html+=`<td class="gantt-cell${vacClass}${internalVacClass}"><div class="gantt-cell-inner"><div class="gantt-spacer"></div></div></td>`;
      }
    });
    html+="</tr>";
  });

  html+="</tbody></table></div>";
  return html;
}

function buildProjectGanttHTMLForRange(rangeStart=null, rangeEnd=null, tasksOverride=null, plainOverride=null){
  const tasksAll = (tasksOverride || []).filter(t=>t.start && t.end);
  if(tasksAll.length===0) return "<div class='gantt-empty'>Aucune tâche date.</div>";
  const rs = rangeStart || null;
  const re = rangeEnd || null;
  const tasks = (!rs || !re) ? tasksAll : tasksAll.filter(t=>{
    const s = new Date(t.start+"T00:00:00");
    const e = new Date(t.end+"T00:00:00");
    return e >= rs && s <= re;
  });
  if(tasks.length===0) return "<div class='gantt-empty'>Aucune tâche date.</div>";
  const missingMap = buildMissingDaysMap(tasks);

  const minStart = rs || tasks.map(t=>new Date(t.start+"T00:00:00")).reduce((a,b)=>a<b?a:b);
  const maxEnd   = re || tasks.map(t=>new Date(t.end+"T00:00:00")).reduce((a,b)=>a>b?a:b);
  const displayStart = rs ? startOfWeek(rs) : minStart;
  const displayEnd = re ? endOfWeek(re) : endOfWeek(maxEnd);

  const weeks=[];
  for(let w=startOfWeek(displayStart); w<=addDays(startOfWeek(displayEnd),0); w=addDays(w,7)) weeks.push(new Date(w));
  const vacWeeks = weeks.map(w=>isVacationWeek(w));
  const internalVacWeeks = weeks.map(w=>isInternalVacationWeek(w));

  tasks.sort((a,b)=>{
    const oa=(taskOrderMap[a.id]||9999)-(taskOrderMap[b.id]||9999);
    if(oa!==0) return oa;
    const sa=Date.parse(a.start||"9999-12-31"), sb=Date.parse(b.start||"9999-12-31");
    if(sa!==sb) return sa-sb;
    return taskTitle(a).localeCompare(taskTitle(b));
  });

  const hideVendor = !ganttColVisibility.projectVendor;
  const hideStatus = !ganttColVisibility.projectStatus;
  const tableClass = `table${hideVendor ? " hide-vendor" : ""}${hideStatus ? " hide-status" : ""}`;

  const isProjectPrint = (plainOverride===null)
    ? (typeof document !== "undefined" && document.body && document.body.classList.contains("print-gantt-project"))
    : !!plainOverride;

  const buildTable = (subsetRows, rowOffset=0, plainMode=false)=>{
    const tableCssClass = plainMode ? "table gantt-export-plain" : tableClass;
    let html=`<div class='tablewrap gantt-table${plainMode ? " gantt-print-plain" : ""}'><table class='${tableCssClass}' style='--gcol1:120px;--gcol2:90px;--gcol3:90px'>`;
    html+="<thead><tr><th class='gantt-task-col-project gantt-col-task'>Tâche</th><th class='gantt-col-vendor' style='width:90px'>Prestataire</th><th class='gantt-col-status' style='width:90px'>Statut</th>";

    weeks.forEach((w,i)=>{
      const info=isoWeekInfo(w);
      const wEnd=endOfWorkWeek(w);
      const range=`${w.toLocaleDateString("fr-FR",{day:"2-digit"})}-${wEnd.toLocaleDateString("fr-FR",{day:"2-digit"})}/${wEnd.toLocaleDateString("fr-FR",{month:"2-digit",year:"2-digit"})}`;
      const weekLabel = `${info.week}`;
      const mondayLabel = formatShortDateTwoLinesHTML(w);
      const todayClass = isTodayInWeek(w) ? " week-today" : "";
      const vacClass = vacWeeks[i] ? " vac-week" : "";
      const internalVacClass = internalVacWeeks[i] ? " vac-week-internal" : "";
      html+=`<th class="week-cell${todayClass}${vacClass}${internalVacClass}" data-range="${range}" style='width:20px;color:#111827'>${weekLabel}<div class="gantt-week-date">${mondayLabel}</div></th>`;
    });

    html+="</tr></thead><tbody>";

    subsetRows.forEach((t,rowIdx)=>{
      const statuses = parseStatuses(t.status).map(v=>v.toUpperCase());
      const mainStatus = statuses[0] || "";
      const color = ownerColor(t.owner);
      const vendorBadges = (()=> {
        const set = new Set();
        if(t.vendor) set.add(t.vendor);
        const typ = ownerType(t.owner);
        if(typ === "interne") set.add("INTERNE");
        if(typ === "rsg") set.add("RSG");
        if(typ === "ri") set.add("RI");
        if(typ === "externe" && !t.vendor) set.add("Prestataire non renseigné");
        if(set.size===0) return "<span class='text-muted'></span>";
        return Array.from(set).sort((a,b)=>a.localeCompare(b,"fr",{sensitivity:"base"})).map(v=>vendorBadge(v)).join(" ");
      })();

      const todayKey = new Date().toISOString().slice(0,10);
      const isToday = !!(t.start && t.end && t.start<=todayKey && t.end>=todayKey);
      const isLate = !!(t.end && t.end < todayKey);
      const rowClassWithToday = `gantt-row${isToday ? " today-row" : ""}${isLate ? " late-row" : ""}`;
      html+=`<tr class="${rowClassWithToday}" data-task="${t.id}">`;
      const p = state?.projects?.find(x=>x.id===t.projectId);
      const sub = (p?.subproject || "").trim();
      const taskDesc = (t.roomNumber || "").trim();
      const label = [sub, taskDesc].filter(Boolean).join("  ");
      const miss = missingMap.get(t.id) || 0;
      const missDot = miss>0 ? `<span class="missing-dot" title="Heures réelles manquantes (${miss} j)"></span>` : "";
      html+=`<td class="gantt-task-col-project gantt-col-task">${missDot}<b><span class="num-badge" style="--badge-color:${color};--badge-text:#fff;">${taskOrderMap[t.id]||""}</span></b> <span class="gantt-task-name">${attrEscape(label)}</span></td>`;
      html+=`<td class="gantt-vendor-cell gantt-col-vendor"><div class="vendor-stack">${vendorBadges}</div></td>`;
      html+=`<td class="gantt-status-cell gantt-col-status"><div class="gantt-status-stack"><div class="status-row"><span>${statusLabels(mainStatus)}</span></div></div></td>`;

      weeks.forEach((w,i)=>{
        const sDate=new Date(t.start+"T00:00:00");
        const eDate=new Date(t.end+"T00:00:00");
        const geo=barGeometry(sDate,eDate,w);
        const vacClass = vacWeeks[i] ? " vac-week" : "";
        const internalVacClass = internalVacWeeks[i] ? " vac-week-internal" : "";
        if(geo.days>0){
          if(plainMode){
            html+=`<td class="gantt-cell${vacClass}${internalVacClass}"><div class="gantt-cell-inner"><div class="bar-wrapper"><div class="gantt-print-bar" style="width:${geo.width}%;margin-left:${geo.offset}%;background:${color};border-color:${color};"></div></div></div></td>`;
          }else{
            const title = t.vendor ? ` title="Prestataire : ${attrEscape(t.vendor)}"` : "";
            const barDelay = ((rowOffset + rowIdx) * 0.03 + i * 0.015).toFixed(3);
            html+=`<td class="gantt-cell${vacClass}${internalVacClass}"><div class="gantt-cell-inner"><div class="bar-wrapper"><div class="gantt-bar" data-task="${t.id}" data-status="${mainStatus}"${title} style="width:${geo.width}%;margin-left:${geo.offset}%;background:${color};border-color:${color};--bar-delay:${barDelay}s"><span class="gantt-days">${geo.days} j</span></div></div></div></td>`;
          }
        }else{
          html+=`<td class="gantt-cell${vacClass}${internalVacClass}"><div class="gantt-cell-inner"><div class="gantt-spacer"></div></div></td>`;
        }
      });
      html+="</tr>";
    });

    html+="</tbody></table></div>";
    return html;
  };

  if(!isProjectPrint){
    return buildTable(tasks, 0, false);
  }

  return buildTable(tasks, 0, true);
}
function renderMasterGantt(){

  const wrap = el("masterGantt");

  if(!wrap) return;

  const tasks = filteredTasks().filter(t=>t.start && t.end);

  if(tasks.length===0){

    wrap.innerHTML = "<div class='gantt-empty'>Aucune tâche date.</div>";

    return;

  }

  const minStart = tasks.map(t=>new Date(t.start+"T00:00:00")).reduce((a,b)=>a<b?a:b);

  const maxEnd   = tasks.map(t=>new Date(t.end+"T00:00:00")).reduce((a,b)=>a>b?a:b);

  const weeks=[];
  for(let w=startOfWeek(minStart); w<=addDays(startOfWeek(maxEnd),0); w=addDays(w,7)) weeks.push(new Date(w));
  const vacWeeks = weeks.map(w=>isVacationWeek(w));
  const internalVacWeeks = weeks.map(w=>isInternalVacationWeek(w));



  tasks.sort((a,b)=>{
    const sa=Date.parse(a.start||"9999-12-31"), sb=Date.parse(b.start||"9999-12-31");
    if(sa!==sb) return sa-sb;
    const ea=Date.parse(a.end||"9999-12-31"), eb=Date.parse(b.end||"9999-12-31");
    if(ea!==eb) return ea-eb;
    return taskTitle(a).localeCompare(taskTitle(b),"fr",{sensitivity:"base"});
  });



  let html="<div class='tablewrap gantt-table'><table class='table' style='--gcol0:70px;--gcol1:150px;--gcol2:140px;--gcol3:120px'>";

  html+="<thead><tr><th class='gantt-col-site' style='width:70px'>Site / Zone</th><th class='gantt-col-task' style='width:150px'>Nom</th><th class='gantt-col-vendor' style='width:140px'>Prestataire</th><th class='gantt-col-status' style='width:120px'>Statut</th>";

  weeks.forEach((w,i)=>{

    const info=isoWeekInfo(w);

    const wEnd=endOfWorkWeek(w);

    const range=`${w.toLocaleDateString("fr-FR",{day:"2-digit"})}-${wEnd.toLocaleDateString("fr-FR",{day:"2-digit"})}/${wEnd.toLocaleDateString("fr-FR",{month:"2-digit",year:"2-digit"})}`;

    const weekLabel = `S${String(info.week).padStart(2,"0")}`;

    const mondayLabel = formatShortDateTwoLinesHTML(w);

    const todayClass = isTodayInWeek(w) ? " week-today" : "";
    const vacClass = vacWeeks[i] ? " vac-week" : "";
    const internalVacClass = internalVacWeeks[i] ? " vac-week-internal" : "";
    html+=`<th class="week-cell${todayClass}${vacClass}${internalVacClass}" data-range="${range}" style='width:72px;color:#111827'>${weekLabel}<div class="gantt-week-date">${mondayLabel}</div></th>`;

  });

  html+="</tr></thead><tbody>";



  tasks.forEach((t,rowIdx)=>{

    const statuses = parseStatuses(t.status).map(v=>v.toUpperCase());

    const mainStatus = statuses[0] || "";

    const color = ownerColor(t.owner);

    const p = state?.projects?.find(x=>x.id===t.projectId);

    const projectName = (p?.name || "Projet").trim() || "Projet";
    const siteLabel = (p?.site || "").trim();

    const vendorBadges = (()=> {

      const set = new Set();

      if(t.vendor) set.add(t.vendor);

      const typ = ownerType(t.owner);
      if(typ === "interne") set.add("INTERNE");
      if(typ === "rsg") set.add("RSG");
      if(typ === "ri") set.add("RI");
      if(typ === "externe" && !t.vendor) set.add("Prestataire non renseigné");

      if(set.size===0) return "<span class='text-muted'></span>";

      return Array.from(set).sort((a,b)=>a.localeCompare(b,"fr",{sensitivity:"base"}))

        .map(v=>vendorBadge(v)).join(" ");

    })();



    const todayKey = new Date().toISOString().slice(0,10);
    const isToday = !!(t.start && t.end && t.start<=todayKey && t.end>=todayKey);
    const isLate = !!(t.end && t.end < todayKey);
    const rowClass = t.id===selectedTaskId ? "gantt-row gantt-row-active" : "gantt-row";
    const rowClassWithToday = `${rowClass}${isToday ? " today-row" : ""}${isLate ? " late-row" : ""}`;
    html+=`<tr class="${rowClassWithToday}" data-task="${t.id}" onclick="openTaskFromGantt('${t.id}')">`;
    html+=`<td class="gantt-col-site">${attrEscape(siteLabel || "")}</td>`;
    html+=`<td class="gantt-col-task"><span class="num-badge" style="--badge-color:${color};--badge-text:#fff;">${taskOrderMap[t.id]||""}</span> <span class="gantt-task-name">${attrEscape(projectName)}</span></td>`;

    html+=`<td class="gantt-vendor-cell gantt-col-vendor"><div class="vendor-stack">${vendorBadges}</div></td>`;

    html+=`<td class="gantt-status-cell gantt-col-status"><div class="gantt-status-stack"><div class="status-row"><span>${statusLabels(mainStatus)}</span></div></div></td>`;



    weeks.forEach((w,i)=>{

      const sDate=new Date(t.start+"T00:00:00");

      const eDate=new Date(t.end+"T00:00:00");

      const geo=barGeometry(sDate,eDate,w);

      if(geo.days>0){

        const title = t.vendor ? ` title="Prestataire : ${attrEscape(t.vendor)}"` : "";

        const vacClass = vacWeeks[i] ? " vac-week" : "";
        const internalVacClass = internalVacWeeks[i] ? " vac-week-internal" : "";
        const barDelay = (rowIdx * 0.03 + i * 0.015).toFixed(3);
        html+=`<td class="gantt-cell${vacClass}${internalVacClass}"><div class="gantt-cell-inner"><div class="bar-wrapper"><div class="gantt-bar bar-click" data-task="${t.id}" data-status="${mainStatus}"${title} style="width:${geo.width}%;margin-left:${geo.offset}%;background:${color};border-color:${color};--bar-delay:${barDelay}s"><span class="gantt-days">${geo.days} j</span></div></div></div></td>`;

      }else{

        const vacClass = vacWeeks[i] ? " vac-week" : "";
        const internalVacClass = internalVacWeeks[i] ? " vac-week-internal" : "";
        html+=`<td class="gantt-cell${vacClass}${internalVacClass}"><div class="gantt-cell-inner"><div class="gantt-spacer"></div></div></td>`;

      }

    });



    html+="</tr>";

  });



  html+="</tbody></table></div>";

  wrap.innerHTML=html;

  applyGanttColumnVisibility();

  wrap.querySelectorAll(".bar-click")?.forEach(bar=>{

    bar.onclick=(e)=>{

      if(e && e.stopPropagation) e.stopPropagation();

      const taskId = bar.dataset.task;

      const task = state.tasks.find(x=>x.id===taskId);

      if(!task) return;

      openTaskFromGantt(taskId);

    };

  });

}






function renderWorkloadChartFor(tasks, chartId, pieId, uiIds=null, stateRef=null, boundsTasks=null, autoReset=false){

  const mode = "week";

  const bounds = boundsTasks || tasks;

  syncWorkloadFilterUI(tasks, bounds, uiIds, stateRef);

  const range = getWorkloadRange(tasks, bounds, stateRef);

  let data = computeWorkloadData(tasks, mode, range.start, range.end);

  if(autoReset && tasks.length>0 && data.length===0 && stateRef){

    // reset to "all" if filter excludes all data for this project

    stateRef.type = "all";

    stateRef.start = "";

    stateRef.end = "";

    stateRef.year = "";

    syncWorkloadFilterUI(tasks, bounds, uiIds, stateRef);

    const range2 = getWorkloadRange(tasks, bounds, stateRef);

    data = computeWorkloadData(tasks, mode, range2.start, range2.end);

  }

  const svg = el(chartId);

  if(!svg) return;

  const pieSvg = pieId ? el(pieId) : null;

  const idPrefix = (chartId || "workload").replace(/[^a-zA-Z0-9_-]/g,"");

  const gradIntId = `${idPrefix}-grad-int`;

  const gradExtId = `${idPrefix}-grad-ext`;
  const gradRsgId = `${idPrefix}-grad-rsg`;
  const gradRiId = `${idPrefix}-grad-ri`;
  const pieShadowId = `${idPrefix}-pieShadow`;

  const brushedId = `${idPrefix}-brushed`;

  const brushedBaseId = `${idPrefix}-brushed-base`;

  const barShadowId = `${idPrefix}-barShadow`;

  const w=900, h=320, m={l:60,r:24,t:44,b:64};

  const fontFamily = `"Segoe UI", Arial, sans-serif`;

  svg.setAttribute("viewBox",`0 0 ${w} ${h}`);

  svg.style.fontFamily = fontFamily;

  svg.setAttribute("font-family", fontFamily);

  svg.innerHTML="";

  if(pieSvg) pieSvg.innerHTML="";

  if(data.length===0){

    svg.innerHTML = `<text x="${w/2}" y="${h/2}" text-anchor="middle" fill="#6b7280" font-size="12">Aucune tâche date</text>`;

    if(pieSvg){

      pieSvg.setAttribute("viewBox", "0 0 720 360");

      pieSvg.innerHTML = `<text x="360" y="180" text-anchor="middle" fill="#6b7280" font-size="12">Aucune donnée</text>`;

    }

    return;

  }

  const maxVal = niceMax(Math.max(...data.map(d=>d.total),1));

  const chartW = w - m.l - m.r;

  const chartH = h - m.t - m.b;

  const isDay = mode === "day";

  const denseLayout = data.length > 22;
  const groupGap = denseLayout ? 2 : 8;

  const groupW = Math.max(8, Math.min(90, (chartW / data.length) - groupGap));

  const innerGap = denseLayout ? 1 : 4;

  const barW = Math.max(1.2, (groupW - innerGap*3) / 4);

  const denseMode = data.length > 18;
  const axisLabelStep = denseMode ? Math.ceil(data.length / 14) : 1;
  const hideValueLabels = true;
  const labelMinValue = maxVal >= 60 ? 5 : (maxVal >= 25 ? 3 : 1);
  const maxLabelsPerGroup = maxVal >= 25 ? 2 : 4;

  const xStart = m.l;

  let grid="";

  const ticks=4;
  const fmtHours = (v)=>{
    if(!isFinite(v)) return "0";
    const r = Math.round(v * 10) / 10;
    return Number.isInteger(r) ? String(r) : r.toFixed(1);
  };

  for(let i=0;i<=ticks;i++){

    const y = m.t + chartH - (i/ticks)*chartH;

    const val = (i/ticks)*maxVal;

    grid+=`<line class="wl-grid" x1="${m.l}" y1="${y}" x2="${w-m.r}" y2="${y}"></line>`;

    grid+=`<text class="wl-axis" x="${m.l-10}" y="${y+4}" text-anchor="end">${fmtHours(val)} h</text>`;

  }

  // vertical guides

  data.forEach((d,idx)=>{

    const gx = xStart + idx*(groupW+groupGap) + groupW/2;

    grid+=`<line class="wl-grid-vert" x1="${gx}" y1="${m.t}" x2="${gx}" y2="${m.t+chartH}"></line>`;

  });



  let bars="";

  data.forEach((d,idx)=>{

    const gx = xStart + idx*(groupW+groupGap);

    const xInt = gx;
    const xExt = gx + barW + innerGap;
    const xRsg = gx + (barW + innerGap) * 2;
    const xRi  = gx + (barW + innerGap) * 3;

    let hInt = (d.internal/maxVal)*chartH;
    let hExt = (d.external/maxVal)*chartH;
    let hRsg = (d.rsg/maxVal)*chartH;
    let hRi  = (d.ri/maxVal)*chartH;

    if(d.internal > 0 && hInt < 6) hInt = 6;
    if(d.external > 0 && hExt < 6) hExt = 6;
    if(d.rsg > 0 && hRsg < 6) hRsg = 6;
    if(d.ri > 0 && hRi < 6) hRi = 6;

    const yBase = m.t + chartH;

    const yInt = yBase - hInt;
    const yExt = yBase - hExt;
    const yRsg = yBase - hRsg;
    const yRi = yBase - hRi;

    const dInt = (idx * 0.06 + 0.00).toFixed(2);
    const dExt = (idx * 0.06 + 0.02).toFixed(2);
    const dRsg = (idx * 0.06 + 0.04).toFixed(2);
    const dRi  = (idx * 0.06 + 0.06).toFixed(2);
    bars+=`<rect class="wl-bar-internal wl-anim-bar" style="--wl-delay:${dInt}s" fill="url(#${gradIntId})" filter="url(#${barShadowId})" x="${xInt}" y="${yInt}" width="${barW}" height="${hInt}" rx="3" ry="3"></rect>`;
    bars+=`<rect class="wl-bar-external wl-anim-bar" style="--wl-delay:${dExt}s" fill="url(#${gradExtId})" filter="url(#${barShadowId})" x="${xExt}" y="${yExt}" width="${barW}" height="${hExt}" rx="3" ry="3"></rect>`;
    bars+=`<rect class="wl-bar-rsg wl-anim-bar" style="--wl-delay:${dRsg}s" fill="url(#${gradRsgId})" filter="url(#${barShadowId})" x="${xRsg}" y="${yRsg}" width="${barW}" height="${hRsg}" rx="3" ry="3"></rect>`;
    bars+=`<rect class="wl-bar-ri wl-anim-bar" style="--wl-delay:${dRi}s" fill="url(#${gradRiId})" filter="url(#${barShadowId})" x="${xRi}" y="${yRi}" width="${barW}" height="${hRi}" rx="3" ry="3"></rect>`;

    const lbl = keyToLabel(d.key, mode);

    const lx = gx + groupW/2;

    const ly = h - m.b + 14;

    const showXAxisLabel = (idx % axisLabelStep) === 0 || idx === data.length - 1;
    if(showXAxisLabel){
      bars+=`<text class="wl-axis" x="${lx}" y="${ly}" text-anchor="middle">${lbl}</text>`;
    }

    const valueYInt = Math.max(m.t + 12, yBase - Math.max(hInt, 0) - 8);
    const valueYExt = Math.max(m.t + 12, yBase - Math.max(hExt, 0) - 8);
    const valueYRsg = Math.max(m.t + 12, yBase - Math.max(hRsg, 0) - 8);
    const valueYRi = Math.max(m.t + 12, yBase - Math.max(hRi, 0) - 8);

    if(!hideValueLabels){
      const labelCandidates = [
        { v: Number(d.internal) || 0, x: xInt + barW/2, y: valueYInt },
        { v: Number(d.external) || 0, x: xExt + barW/2, y: valueYExt },
        { v: Number(d.rsg) || 0, x: xRsg + barW/2, y: valueYRsg },
        { v: Number(d.ri) || 0, x: xRi + barW/2, y: valueYRi }
      ]
        .filter((c)=>c.v >= labelMinValue)
        .sort((a,b)=>b.v-a.v)
        .slice(0, maxLabelsPerGroup);
      labelCandidates.forEach((c)=>{
        bars+=`<text class="wl-value" x="${c.x}" y="${c.y}" text-anchor="middle">${fmtHours(c.v)} h</text>`;
      });
    }

  });

  const totalInt = data.reduce((s,d)=>s+d.internal,0);
  const totalExt = data.reduce((s,d)=>s+d.external,0);
  const totalRsg = data.reduce((s,d)=>s+d.rsg,0);
  const totalRi  = data.reduce((s,d)=>s+d.ri,0);

  const legend=`<g transform="translate(${w-280},12)">

    <rect x="0" y="0" width="12" height="12" rx="3" fill="url(#${gradIntId})"></rect>

    <text class="wl-axis" x="18" y="11">Interne ${fmtHours(totalInt)} h</text>

    <rect x="0" y="20" width="12" height="12" rx="3" fill="url(#${gradExtId})"></rect>

    <text class="wl-axis" x="18" y="31">Externe ${fmtHours(totalExt)} h</text>

    <rect x="0" y="40" width="12" height="12" rx="3" fill="url(#${gradRsgId})"></rect>

    <text class="wl-axis" x="18" y="51">RSG ${fmtHours(totalRsg)} h</text>

    <rect x="0" y="60" width="12" height="12" rx="3" fill="url(#${gradRiId})"></rect>

    <text class="wl-axis" x="18" y="71">RI ${fmtHours(totalRi)} h</text>

  </g>`;

  const legendOverlay = `

    <g transform="translate(${w-470},12)">

      <rect x="-6" y="-6" width="430" height="26" rx="8" ry="8" fill="rgba(255,255,255,0.92)" stroke="#e5e7eb"/>

      <rect x="0" y="0" width="12" height="12" rx="3" fill="url(#${gradIntId})"></rect>

      <text class="wl-axis" x="18" y="11">Interne ${fmtHours(totalInt)} h</text>

      <rect x="120" y="0" width="12" height="12" rx="3" fill="url(#${gradExtId})"></rect>

      <text class="wl-axis" x="138" y="11">Externe ${fmtHours(totalExt)} h</text>

      <rect x="240" y="0" width="12" height="12" rx="3" fill="url(#${gradRsgId})"></rect>

      <text class="wl-axis" x="258" y="11">RSG ${fmtHours(totalRsg)} h</text>

      <rect x="320" y="0" width="12" height="12" rx="3" fill="url(#${gradRiId})"></rect>

      <text class="wl-axis" x="338" y="11">RI ${fmtHours(totalRi)} h</text>

    </g>`;

  const defs = `

    <defs>

      <linearGradient id="${brushedBaseId}" x1="0" x2="1" y1="0" y2="0">

        <stop offset="0%" stop-color="#f2f3f5"/>

        <stop offset="45%" stop-color="#e3e6ea"/>

        <stop offset="65%" stop-color="#f7f8fa"/>

        <stop offset="100%" stop-color="#d9dde2"/>

      </linearGradient>

      <pattern id="${brushedId}" width="6" height="6" patternUnits="userSpaceOnUse">

        <rect width="6" height="6" fill="url(#${brushedBaseId})"/>

        <path d="M0 1 H6 M0 3 H6 M0 5 H6" stroke="#cfd3d9" stroke-width="0.4" opacity="0.45"/>

      </pattern>

      <linearGradient id="${gradIntId}" x1="0" x2="0" y1="0" y2="1">

        <stop offset="0%" stop-color="#8bbf9f" stop-opacity="0.95"/>

        <stop offset="55%" stop-color="#78ac8d" stop-opacity="0.9"/>

        <stop offset="100%" stop-color="#679979" stop-opacity="0.86"/>

      </linearGradient>

      <linearGradient id="${gradExtId}" x1="0" x2="0" y1="0" y2="1">

        <stop offset="0%" stop-color="#d6b27b" stop-opacity="0.95"/>

        <stop offset="55%" stop-color="#c49b64" stop-opacity="0.9"/>

        <stop offset="100%" stop-color="#b18754" stop-opacity="0.86"/>

      </linearGradient>
      <linearGradient id="${gradRsgId}" x1="0" x2="0" y1="0" y2="1">
        <stop offset="0%" stop-color="#93abc7" stop-opacity="0.95"/>
        <stop offset="55%" stop-color="#7f97b4" stop-opacity="0.9"/>
        <stop offset="100%" stop-color="#6f86a2" stop-opacity="0.86"/>
      </linearGradient>
      <linearGradient id="${gradRiId}" x1="0" x2="0" y1="0" y2="1">
        <stop offset="0%" stop-color="#b7a6ca" stop-opacity="0.95"/>
        <stop offset="55%" stop-color="#a18dbc" stop-opacity="0.9"/>
        <stop offset="100%" stop-color="#8d79aa" stop-opacity="0.86"/>
      </linearGradient>

      <filter id="${barShadowId}" x="-20%" y="-20%" width="140%" height="160%">

        <feDropShadow dx="0" dy="1.2" stdDeviation="1.1" flood-color="#0b1424" flood-opacity="0.12"/>

      </filter>
      <filter id="${pieShadowId}" x="-30%" y="-30%" width="180%" height="190%">
        <feDropShadow dx="0" dy="1.8" stdDeviation="1.6" flood-color="#0b1424" flood-opacity="0.14"/>
      </filter>

    </defs>

  `;

  svg.innerHTML = `${defs}<rect class="wl-bg" x="0" y="0" width="${w}" height="${h}" fill="url(#${brushedId})"></rect><g>${grid}</g><g>${bars}</g>${legendOverlay}`;



  if(pieSvg){

    const pw=720, ph=360;

    pieSvg.setAttribute("viewBox", `0 0 ${pw} ${ph}`);

    pieSvg.style.fontFamily = fontFamily;

    pieSvg.setAttribute("font-family", fontFamily);

    const pieTotal = Math.max(1, totalInt + totalExt + totalRsg + totalRi);

    const cx = pw/2;
    const cy = 148;
    const r = 104;
    const titleY = ph - 56;
    const legendY = ph - 36;
    const minLabelY = 24;
    const maxLabelY = ph - 96;

    const polar = (cx, cy, r, a)=>{

      const rad = (a - 90) * Math.PI / 180;

      return { x: cx + r * Math.cos(rad), y: cy + r * Math.sin(rad) };

    };

    const arcPath = (cx, cy, r, start, end)=>{

      const s = polar(cx, cy, r, end);

      const e = polar(cx, cy, r, start);

      const large = end - start <= 180 ? 0 : 1;

      return `M ${cx} ${cy} L ${s.x} ${s.y} A ${r} ${r} 0 ${large} 0 ${e.x} ${e.y} Z`;

    };

    let pieMarkup = `<rect class="wl-bg" x="0" y="0" width="${pw}" height="${ph}" fill="url(#${brushedId})"></rect>`;

    if(totalInt + totalExt + totalRsg + totalRi > 0){

      const gap = 8;
      const segments = [
        { key:"interne", label:"Interne", value: totalInt, grad: gradIntId },
        { key:"externe", label:"Externe", value: totalExt, grad: gradExtId },
        { key:"rsg", label:"RSG", value: totalRsg, grad: gradRsgId },
        { key:"ri", label:"RI", value: totalRi, grad: gradRiId },
      ].filter(s=>s.value>0);

      if(segments.length === 1){
        const s = segments[0];
          const labelPos = polar(cx, cy, r * 1.14, 90);
        pieMarkup += `
          <circle class="wl-anim-pie-seg" style="--pie-delay:0.02s" cx="${cx}" cy="${cy}" r="${r}" fill="url(#${s.grad})" filter="url(#${pieShadowId})"></circle>
          <line x1="${cx}" y1="${cy - r * 0.95}" x2="${cx}" y2="${cy - r * 1.06}" stroke="#94a3b8" stroke-width="1" />
          <text class="wl-axis wl-anim-pie-label" style="--pie-delay:0.12s" x="${labelPos.x}" y="${labelPos.y}" text-anchor="middle">${s.label}</text>
          <text class="wl-value wl-anim-pie-label" style="--pie-delay:0.18s" x="${labelPos.x}" y="${labelPos.y + 14}" text-anchor="middle">${fmtHours(s.value)} h</text>
        `;
      }else{
        let cursor = 0;
        segments.forEach((seg,segIdx)=>{
          const angle = (seg.value / pieTotal) * 360;
          const startA = cursor;
          const endA = cursor + angle;
          const midA = startA + angle / 2;
          const off = polar(0,0,gap,midA);
          const path = arcPath(cx+off.x, cy+off.y, r, startA, endA);
          const innerPos = polar(cx+off.x, cy+off.y, r*0.95, midA);
          const outerPos = polar(cx+off.x, cy+off.y, r*1.14, midA);
          const labelY = Math.max(minLabelY, Math.min(maxLabelY, outerPos.y));
          const anchor = outerPos.x < cx ? "end" : "start";
          const segDelay = (segIdx * 0.12 + 0.02).toFixed(2);
          const labDelay = (segIdx * 0.12 + 0.14).toFixed(2);
          pieMarkup += `
            <path class="wl-anim-pie-seg" style="--pie-delay:${segDelay}s" d="${path}" fill="url(#${seg.grad})" filter="url(#${pieShadowId})"></path>
            <line x1="${innerPos.x}" y1="${innerPos.y}" x2="${outerPos.x}" y2="${outerPos.y}" stroke="#94a3b8" stroke-width="1" />
            <text class="wl-axis wl-anim-pie-label" style="--pie-delay:${labDelay}s" x="${outerPos.x}" y="${labelY}" text-anchor="${anchor}">${seg.label}</text>
            <text class="wl-value wl-anim-pie-label" style="--pie-delay:${(Number(labDelay)+0.04).toFixed(2)}s" x="${outerPos.x}" y="${labelY + 14}" text-anchor="${anchor}">${fmtHours(seg.value)} h</text>
          `;
          cursor = endA;
        });
      }

      pieMarkup += `
        <text class="wl-axis" x="${cx}" y="${titleY}" text-anchor="middle">Répartition Interne / Externe / RSG / RI</text>
        <g transform="translate(${cx-250},${legendY})">
          <rect x="0" y="0" width="12" height="12" rx="3" fill="url(#${gradIntId})"></rect>
          <text class="wl-axis" x="18" y="11">Interne ${fmtHours(totalInt)} h</text>
          <rect x="150" y="0" width="12" height="12" rx="3" fill="url(#${gradExtId})"></rect>
          <text class="wl-axis" x="168" y="11">Externe ${fmtHours(totalExt)} h</text>
          <rect x="300" y="0" width="12" height="12" rx="3" fill="url(#${gradRsgId})"></rect>
          <text class="wl-axis" x="318" y="11">RSG ${fmtHours(totalRsg)} h</text>
          <rect x="420" y="0" width="12" height="12" rx="3" fill="url(#${gradRiId})"></rect>
          <text class="wl-axis" x="438" y="11">RI ${fmtHours(totalRi)} h</text>
        </g>
      `;
    }

    pieSvg.innerHTML = `${defs}${pieMarkup}`;

  }

}



function renderWorkloadChart(tasks){

  const allTasks = state?.tasks || tasks || [];
  const masterRange = {type:workloadRangeType, year:workloadRangeYear, start:workloadRangeStart, end:workloadRangeEnd};
  renderWorkloadChartFor(
    allTasks,
    "workloadChart",
    "workloadPie",
    null,
    masterRange,
    allTasks,
    true
  );
  workloadRangeType = masterRange.type;
  workloadRangeYear = masterRange.year;
  workloadRangeStart = masterRange.start;
  workloadRangeEnd = masterRange.end;

}



function renderFilters(){

  const fsSite = el("filterSite");
  if(fsSite){
    const baseSites = Object.keys(SITE_PHOTOS || {});
    const projSites = (state.projects || []).map(p=>String(p?.site||"").trim()).filter(Boolean);
    const all = [...baseSites, ...projSites];
    const canonMap = new Map();
    all.forEach(s=>{
      const raw = String(s || "").trim();
      if(!raw) return;
      const key = raw.normalize("NFD").replace(/[\u0300-\u036f]/g,"").toUpperCase();
      if(!canonMap.has(key)) canonMap.set(key, raw);
    });
    const sites = Array.from(canonMap.values())
      .sort((a,b)=>a.localeCompare(b,"fr",{sensitivity:"base"}));
    const current = fsSite.value || "";
    let h=`<option value="" selected>Tous</option>`;
    sites.forEach(s=>{ h+=`<option value="${attrEscape(s)}">${attrEscape(s)}</option>`; });
    fsSite.innerHTML=h;
    if(current && sites.includes(current)) fsSite.value = current;
    else fsSite.value = "";
  }

  const sel = el("filterProject");

  if(sel){

    let h=`<option value="" selected>Tous</option>`;

    state.projects.forEach(p=>{ h+=`<option value="${p.id}">${p.name||"Sans nom"}</option>`; });

    sel.innerHTML=h;
    if(!sel.value) sel.value = "";

  }

  const ss = el("filterStatus");

  if(ss){

    let h=`<option value="" selected>Tous</option>`;

    sortedStatuses().forEach(s=>{ h+=`<option value="${s.v}">${s.label}</option>`; });

    ss.innerHTML=h;
    if(!ss.value) ss.value = "";

  }

}



function renderTabs(){

  const tabs = el("tabs");
  const tabsMaster = el("tabsMaster");
  const tabsSortBtn = el("btnTabsSortProgress");
  const tabsSortResetBtn = el("btnTabsSortReset");

  if(!tabs) return;

  if(tabsSortBtn){
    tabsSortBtn.textContent = (tabsSortMode === "progress_desc") ? "TRI 100%->0%" : "TRI 0%->100%";
  }
  if(tabsSortResetBtn){
    tabsSortResetBtn.disabled = (tabsSortMode === "progress_asc");
  }

  const projectIcon = (name="")=>{

    const n=name.toLowerCase();

    if(n.includes("pastorale")) return "";

    return "";

  };

  const masterBtn = `<button class="tab tab-master ${selectedProjectId?"":"active"}" data-tab="MASTER"><span class="tab-icon"></span> Tableau maître</button>`;

  const getProjectCompletion = (projectId)=>{
    const tasks = state.tasks.filter(t=>t.projectId===projectId && t.start && t.end);
    if(!tasks.length) return 0;
    let sum = 0;
    let weight = 0;
    tasks.forEach(t=>{
      const w = Math.max(1, durationDays(t.start, t.end));
      sum += taskProgress(t) * w;
      weight += w;
    });
    return weight ? Math.round(sum / weight) : 0;
  };

  const projectsSorted = [...state.projects].sort((a,b)=>{
    if(tabsSortMode === "progress_asc" || tabsSortMode === "progress_desc"){
      const aProg = getProjectCompletion(a.id);
      const bProg = getProjectCompletion(b.id);
      if(aProg !== bProg){
        return tabsSortMode === "progress_asc" ? (aProg - bProg) : (bProg - aProg);
      }
    }

    // date de début minimale des tâches de chaque projet
    const aDates = state.tasks.filter(t=>t.projectId===a.id && t.start).map(t=>Date.parse(t.start));
    const bDates = state.tasks.filter(t=>t.projectId===b.id && t.start).map(t=>Date.parse(t.start));
    const aMin = aDates.length ? Math.min(...aDates) : Infinity;
    const bMin = bDates.length ? Math.min(...bDates) : Infinity;

    if(aMin!==bMin) return aMin - bMin;
    return (a.name||"").localeCompare(b.name||"");
  });

  let h="";
  const total = projectsSorted.length || 1;
  projectsSorted.forEach((p,idx)=>{
    const hue = 0 + (120 * (idx/(total-1 || 1))); // rouge -> vert
    const progress = getProjectCompletion(p.id);
    h+=`<button class="tab ${selectedProjectId===p.id?"active":""}" data-tab="${p.id}" style="--tab-hue:${hue};--tab-progress:${progress}%;--tab-progress-color:hsl(${hue} 72% 45%);"><span class="tab-progress-fill" style="width:${progress}%;background:hsl(${hue} 78% 66% / .42);"></span><span>${p.name||"Projet"}</span><span class="tab-close" data-close="${p.id}" aria-label="Supprimer le projet"></span></button>`;
  });

  if(tabsMaster){ tabsMaster.innerHTML = masterBtn; }
  tabs.innerHTML=h;

  const allTabs = [
    ...(tabsMaster ? Array.from(tabsMaster.querySelectorAll("button")) : []),
    ...Array.from(tabs.querySelectorAll("button"))
  ];
  allTabs.forEach(btn=>{
    btn.onclick=()=>{
      const tab=btn.dataset.tab;
      if(tab==="MASTER"){ navigateTo(null, null, true); }
      else { navigateTo(tab, null, true); }
    };
  });
  tabs.querySelectorAll(".tab-close").forEach(close=>{

    close.onclick=(e)=>{

      e.stopPropagation();

      if(isLocked) return;

      const pid = close.dataset.close;

      if(!pid) return;

      const name = state.projects.find(p=>p.id===pid)?.name || "ce projet";

      if(!confirm(`Supprimer définitivement ${name} et toutes ses tâches ?`)) return;

      state.projects = state.projects.filter(p=>p.id!==pid);

      state.tasks    = state.tasks.filter(t=>t.projectId!==pid);

      if(selectedProjectId===pid) { selectedProjectId=null; selectedTaskId=null; }

      saveState();

      renderAll();

      el("btnNewTask")?.classList.remove("btn-armed");

    };

  });

}



function renderKPIs(tasks){

  const box = el("kpis");

  if(!box) return;

  const total=tasks.length;

  const byStatus={};

  tasks.forEach(t=>{ byStatus[t.status]= (byStatus[t.status]||0)+1; });

  let h=`<div class="kpi"><span class="kpi-dot" style="--kpi-dot:#64748b"></span>Total:&nbsp;<b>${total}</b></div>`;

  STATUSES.forEach(s=>{

    h+=`<div class="kpi"><span class="kpi-dot" style="--kpi-dot:${statusColor(s.v)}"></span>${s.label}:&nbsp;<b>${byStatus[s.v]||0}</b></div>`;

  });

  box.innerHTML=h;

}



// Tri générique pour tableaux (master & projet)

function sortTasks(list, cfg){

  if(!cfg) return [...list];

  const dir = cfg.dir==="desc" ? -1 : 1;

  const get = (t)=>{

    switch(cfg.key){

      case "site": {

        const p = state.projects.find(x=>x.id===t.projectId);

        return (p?.site||"").toLowerCase();

      }

      case "project": {

        const p = state.projects.find(x=>x.id===t.projectId);

        return (p?.name||"").toLowerCase();

      }

      case "task": return (taskTitle(t)||"").toLowerCase();

      case "status": return (parseStatuses(t.status)[0]||"").toLowerCase();

      case "start": return Date.parse(t.start||"9999-12-31");

      case "end": return Date.parse(t.end||"9999-12-31");

      case "owner": return (t.owner||"").toLowerCase();

      case "duration": return durationDays(t.start,t.end);
      case "progress": return taskProgress(t);

      case "num": return taskOrderMap[t.id]||9999;

      default: return 0;

    }

  };

  return [...list].sort((a,b)=>{

    const va=get(a), vb=get(b);

    if(va<vb) return -1*dir;

    if(va>vb) return 1*dir;

    return 0;

  });

}



function updateSortIndicators(tableId, cfg){

  const table = el(tableId);

  if(!table) return;

  table.querySelectorAll("th[data-sort]").forEach(th=>{

    th.classList.remove("sorted-asc","sorted-desc");

    if(th.dataset.sort === cfg.key){

      th.classList.add(cfg.dir==="desc" ? "sorted-desc" : "sorted-asc");

    }

  });

}



function renderMasterMetrics(tasks){

  const metrics = el("masterMetrics");

  if(!metrics) return;

  const dated = tasks.filter(t=>t.start && t.end);

  if(dated.length===0){

    metrics.innerHTML="";

    return;

  }

  const allDays = new Set();

  const internalDays = new Set();
  const externalDays = new Set();
  const rsgDays = new Set();
  const riDays = new Set();

  dated.forEach(t=>{

    const s=new Date(t.start+"T00:00:00");

    const e=new Date(t.end+"T00:00:00");

    if(isNaN(s)||isNaN(e)||e<s) return;

    const typ = ownerType(t.owner);
    const ownsInternal = typ === "interne";
    const ownsExternal = typ === "externe" || typ === "inconnu";
    const ownsRsg = typ === "rsg";
    const ownsRi = typ === "ri";

    for(let d=new Date(s); d<=e; d.setDate(d.getDate()+1)){

      const key=d.toISOString().slice(0,10);

      allDays.add(key);

      if(ownsInternal) internalDays.add(key);
      if(ownsExternal) externalDays.add(key);
      if(ownsRsg) rsgDays.add(key);
      if(ownsRi) riDays.add(key);

    }

  });

  const totalDays = allDays.size;

  const real = getRealMinutesForTasks(dated);
  let progSumMaster = 0;
  let progWeightMaster = 0;
  dated.forEach(t=>{
    const w = Math.max(1, durationDays(t.start, t.end));
    progSumMaster += taskProgress(t) * w;
    progWeightMaster += w;
  });
  const avgProgressMaster = progWeightMaster ? Math.round(progSumMaster / progWeightMaster) : 0;

  metrics.innerHTML = `

    <span class="panel-chip">Durée totale : <span class="metric-val">${totalDays||0} j</span></span>
    <span class="panel-chip">Avancement : <span class="metric-val">${avgProgressMaster}%</span></span>

    <span class="panel-chip">Heures réelles : <span class="metric-val">${formatHoursMinutes(real.totalMinutes||0)}</span></span>

    <span class="panel-chip" style="background:#e8eef8;color:#1f2937;border-color:#c7d2fe;">Interne : <span class="metric-val">${internalDays.size||0} j</span> <span class="metric-val">${formatHoursMinutes(real.internalMinutes||0)}</span></span>

    <span class="panel-chip" style="background:#fff1e6;color:#1f2937;border-color:#fdba74;">Externe : <span class="metric-val">${externalDays.size||0} j</span> <span class="metric-val">${formatHoursMinutes(real.externalMinutes||0)}</span></span>
    <span class="panel-chip" style="background:#e8f0ff;color:#1f2937;border-color:#93c5fd;">RSG : <span class="metric-val">${rsgDays.size||0} j</span> <span class="metric-val">${formatHoursMinutes(real.rsgMinutes||0)}</span></span>
    <span class="panel-chip" style="background:#f2eaff;color:#1f2937;border-color:#c4b5fd;">RI : <span class="metric-val">${riDays.size||0} j</span> <span class="metric-val">${formatHoursMinutes(real.riMinutes||0)}</span></span>
  `;
  animateMetricCounters(metrics);

}

function canonSiteKey(raw){
  return String(raw || "")
    .trim()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g,"")
    .toUpperCase();
}
function getAllSitesList(){
  const baseSites = Object.keys(SITE_PHOTOS || {});
  const projSites = (state.projects || []).map(p=>String(p?.site||"").trim()).filter(Boolean);
  const all = [...baseSites, ...projSites];
  const canonMap = new Map();
  all.forEach(s=>{
    const raw = String(s || "").trim();
    if(!raw) return;
    const key = canonSiteKey(raw);
    if(!canonMap.has(key)) canonMap.set(key, raw);
  });
  return Array.from(canonMap.values())
    .sort((a,b)=>a.localeCompare(b,"fr",{sensitivity:"base"}));
}
function getSelectedExportSites(){
  const wrap = el("ganttExportSites");
  if(!wrap) return [];
  return Array.from(wrap.querySelectorAll("input[data-site]"))
    .filter(cb=>cb.checked)
    .map(cb=>cb.dataset.site);
}
function renderGanttExportSites(){
  const wrap = el("ganttExportSites");
  const label = el("ganttExportSitesLabel");
  if(!wrap) return;
  const isProject = ganttExportContext === "project";
  if(label) label.style.display = isProject ? "none" : "";
  wrap.style.display = isProject ? "none" : "flex";
  if(isProject) return;

  const sites = getAllSitesList();
  const prev = new Set();
  wrap.querySelectorAll("input[data-site]").forEach(cb=>{ if(cb.checked) prev.add(cb.dataset.site); });
  const usePrev = prev.size > 0;
  let h = `<label><input type="checkbox" id="ganttExportSitesAll" checked> Tous</label>`;
  sites.forEach(s=>{
    const checked = !usePrev || prev.has(s);
    h += `<label><input type="checkbox" data-site="${attrEscape(s)}" ${checked ? "checked" : ""}> ${attrEscape(s)}</label>`;
  });
  wrap.innerHTML = h;

  const allBox = el("ganttExportSitesAll");
  const siteBoxes = ()=>Array.from(wrap.querySelectorAll("input[data-site]"));
  const updateAll = ()=>{
    const boxes = siteBoxes();
    const checked = boxes.filter(b=>b.checked).length;
    if(allBox){
      allBox.checked = checked === boxes.length;
      allBox.indeterminate = checked > 0 && checked < boxes.length;
    }
  };
  updateAll();
  if(allBox){
    allBox.onchange = ()=>{
      const checked = allBox.checked;
      siteBoxes().forEach(b=>{ b.checked = checked; });
      updateAll();
    };
  }
  siteBoxes().forEach(b=>{
    b.onchange = ()=> updateAll();
  });
}



function filteredTasks(){

  const fsite = el("filterSite")?.value || "";
  const fp = el("filterProject")?.value || "";
  const fs = el("filterStatus")?.value || "";
  const q  = (el("filterSearch")?.value || "").toLowerCase().trim();
  const startAfter = el("filterStartAfter")?.value || "";
  const endBefore  = el("filterEndBefore")?.value || "";

  const key = `${_stateVersion}|${fsite}|${fp}|${fs}|${q}|${startAfter}|${endBefore}`;
  if(_filteredCache.key === key && _filteredCache.tasks){
    return _filteredCache.tasks;
  }

  const result = state.tasks.filter(t=>{

    if(fp && t.projectId!==fp) return false;

    if(fs && !parseStatuses(t.status).includes(fs)) return false;

    const p = state.projects.find(x=>x.id===t.projectId);
    const pSite = String(p?.site || "");
    if(fsite && pSite !== fsite) return false;

    if(q){

      const hay=(taskTitle(t)+" "+(p?.name||"")+" "+(p?.site||"")+" "+(t.owner||"")+" "+parseStatuses(t.status).join(" ")).toLowerCase();

      if(!hay.includes(q)) return false;

    }

    if(startAfter && (!t.start || t.start < startAfter)) return false;

    if(endBefore && (!t.end || t.end > endBefore)) return false;

    return true;

  });

  // Filet de secours : si les filtres vident tout alors qu'on a des données, on retourne toutes les tâches

  let out = result;
  if(out.length===0 && state.tasks.length>0) out = state.tasks;
  _filteredCache = { key, version:_stateVersion, tasks: out };
  return out;

}



function updateBadge(node, active, textActive="Tri/filtre actif", textInactive="Tri par défaut"){

  if(!node) return;

  node.textContent = active ? textActive : textInactive;

  node.classList.toggle("inactive", !active);

}

function updateResetSortButtonVisual(btn, active){
  if(!btn) return;
  btn.classList.toggle("btn-primary", !!active);
  btn.classList.toggle("btn-ghost", !active);
  btn.classList.remove("btn-danger");
}



function filtersActive(){

  const fsite = el("filterSite")?.value || "";
  const fp = el("filterProject")?.value || "";

  const fs = el("filterStatus")?.value || "";

  const q  = (el("filterSearch")?.value || "").trim();

  const startAfter = el("filterStartAfter")?.value || "";

  const endBefore  = el("filterEndBefore")?.value || "";
  const onlyMissing = !!el("toggleMissingOnly")?.checked;

  return !!(fsite || fp || fs || q || startAfter || endBefore || onlyMissing);

}

function updateSidebarFilterIndicator(){
  const icon = el("filtersActiveIcon");
  const resetBtn = document.querySelector(".filters-reset-row .btn");
  if(!icon) return;
  const active = filtersActive();
  icon.classList.toggle("active", active);
  if(resetBtn){
    resetBtn.classList.toggle("btn-danger", active);
  }
}



function renderMaster(){
  computeTaskOrderMap();
  renderTabs();
  closeAllOverlays();
  el("viewMaster")?.classList.remove("hidden");
  el("viewProject")?.classList.add("hidden");

  const tbody = el("masterTable")?.querySelector("tbody");

  if(!tbody) return;

  const tasks = filteredTasks();

  renderKPIs(tasks);

  renderMasterMetrics(tasks);

  // Charge de travail

  renderWorkloadChart(tasks);

  // Bandeau live global (toutes tâches en cours aujourd'hui)

  const masterLive = el("masterLive");

  if(masterLive){

    const todayKey = new Date().toISOString().slice(0,10);

    const inProgress = tasks

      .filter(t=>t.start && t.end && t.start<=todayKey && t.end>=todayKey)

      .sort((a,b)=> (taskOrderMap[a.id]||999)-(taskOrderMap[b.id]||999));

    if(inProgress.length===0){

      masterLive.innerHTML = `<span class="live-title">Projet non démarré</span>`;

    }else{

      const badges = inProgress.map(t=>{

        const num = taskOrderMap[t.id]||"";

        const status = parseStatuses(t.status)[0] || "";

        const color = ownerColor(t.owner);

        const label = STATUSES.find(s=>s.v===status)?.label || status || "En cours";

        const proj = state.projects.find(x=>x.id===t.projectId);

        const projName = proj?.name || "Projet";

        return `<span class="live-item"><span class="num-badge" style="--badge-color:${color};--badge-text:#fff;">${num}</span> ${projName}  ${label}</span>`;

      }).join(" ");

      masterLive.innerHTML = `<span class="live-title">Projet démarré  Tâches en cours :</span> ${badges}`;

    }

  }

  const sorted = sortTasks(tasks, sortMaster);
  const missingMap = buildMissingDaysMap(sorted);
  const todayKey = new Date().toISOString().slice(0,10);
  const allTasks = Array.isArray(state?.tasks) ? state.tasks : [];
  const missingMapAll = buildMissingDaysMap(allTasks);
  const missingHoursCount = allTasks.reduce((acc, t)=> acc + ((missingMapAll.get(t.id) || 0) > 0 ? 1 : 0), 0);
  const onlyMissingEnabled = !!el("toggleMissingOnly")?.checked;
  const visibleTasks = onlyMissingEnabled
    ? sorted.filter(t=> (missingMap.get(t.id) || 0) > 0)
    : sorted;
  const missingOnlyToggleWrap = el("missingOnlyToggleWrap");
  if(missingOnlyToggleWrap){
    missingOnlyToggleWrap.classList.toggle("inactive", !onlyMissingEnabled);
  }
  const missingHoursBadge = el("missingHoursBadge");
  if(missingHoursBadge){
    updateBadge(
      missingHoursBadge,
      missingHoursCount > 0,
      `${missingHoursCount} tâche(s) à compléter`,
      "Heures réelles: OK"
    );
  }
  const processMissingBtn = el("btnProcessMissingHours");
  if(processMissingBtn){
    processMissingBtn.classList.toggle("missing-hours-attention", missingHoursCount > 0);
    processMissingBtn.disabled = missingHoursCount <= 0;
    processMissingBtn.textContent = "Completer Heures Réelles";
    processMissingBtn.setAttribute("aria-label", `Completer Heures Réelles (${missingHoursCount})`);
    if(missingHoursCount > 0){
      const ae = document.activeElement;
      const typingTarget = !!(ae && (ae.tagName === "INPUT" || ae.tagName === "TEXTAREA" || ae.tagName === "SELECT" || ae.isContentEditable));
      const modalOpen = isHoursTaskModalOpen();
      if(!typingTarget && !modalOpen){
        setTimeout(()=>{
          try{ processMissingBtn.focus({preventScroll:true}); }catch(e){ softCatch(e); }
        }, 0);
      }
    }
  }

  if(visibleTasks.length===0){

    tbody.innerHTML = onlyMissingEnabled
      ? "<tr><td colspan='8' class='empty-row'>Aucune tâche à compléter.</td></tr>"
      : "<tr><td colspan='8' class='empty-row'>Aucune tâche.</td></tr>";

    return;

  }

  let h="";

  visibleTasks.forEach(t=>{

    const p = state.projects.find(x=>x.id===t.projectId);

    const statuses = parseStatuses(t.status).map(v=>v.toUpperCase());

    const c = ownerColor(t.owner);

    const rowBg = siteColor(p?.site);
    const sub = (p?.subproject || "").trim();
    const projLabel = sub ? `${p?.name||"Sans projet"} - ${sub}` : (p?.name||"Sans projet");
    const taskLabel = (t.roomNumber||"").trim();

    const isToday = !!(t.start && t.end && t.start<=todayKey && t.end>=todayKey);
    const isLate = !!(t.end && t.end < todayKey);
    const rowClass = `${isToday ? "today-row " : ""}${isLate ? "late-row" : ""}`.trim();
    const miss = missingMap.get(t.id) || 0;
    const missDot = miss>0 ? `<span class="missing-dot" title="Heures réelles manquantes (${miss} j)"></span>` : "";
    h+=`<tr class="${rowClass}" data-project="${t.projectId}" data-task="${t.id}" style="--site-bg:${rowBg};background:var(--site-bg);">

      <td>${p?.site||""}</td>

      <td>${projLabel}</td>

      <td>${missDot}<span class="num-badge" style="--badge-color:${c};--badge-text:#fff;">${taskOrderMap[t.id]||""}</span> <span class="icon-picto"></span> ${taskLabel}</td>

      <td class="status-cell"><span class="status-left">${statusDot(statuses[0])}${statusLabels(t.status||"")}</span>${t.owner?ownerBadgeForTask(t):""}</td>

      <td>${formatDate(t.start)||""}${isToday ? `<span class="today-dot" title="En cours aujourd'hui"></span>` : ""}</td>

      <td>${formatDate(t.end)||""}</td>
      <td>${taskProgress(t)}%</td>
      <td>${durationLabelForTask(t)}</td>

    </tr>`;

  });

  tbody.innerHTML=h;



  // Gantt global sous le tableau maître

  renderMasterGantt();

  updateSortIndicators("masterTable", sortMaster);

  const fb = el("filtersBadge");

  if(fb){

    const active = filtersActive() || sortMaster.key!=="start" || sortMaster.dir!=="asc";

    updateBadge(fb, active, "Tri/filtre actif", "Tri par défaut");
    updateResetSortButtonVisual(el("btnResetSortMaster"), active);

  }
  updateSidebarFilterIndicator();
  const siteBadge = el("siteBadge");
  if(siteBadge){
    const fp = el("filterProject")?.value || "";
    if(fp){
      const proj = state.projects.find(p=>p.id===fp);
      const site = (proj?.site || "").trim();
      if(site){
        siteBadge.textContent = `Site actif : ${site}`;
        siteBadge.style.display = "inline-flex";
      }else{
        siteBadge.style.display = "none";
      }
    }else{
      siteBadge.style.display = "none";
    }
  }
  const masterHoursReport = el("masterHoursReport");
  if(masterHoursReport){
    masterHoursReport.innerHTML = buildMasterRealHoursReportInnerHTML();
  }

  animateBadgeChanges(el("viewMaster"));
  animateCardsInView("viewMaster");

}

function getTimeLogs(){
  if(!state) return [];
  if(!Array.isArray(state.timeLogs)) state.timeLogs = [];
  return state.timeLogs;
}
function getCurrentUserKey(){
  const email = getCurrentUserEmail();
  const name = getCurrentUserName();
  return (email || name || "").trim();
}
function resolveTimeLogRole(name, email){
  const n = (name || "").toLowerCase();
  const e = (email || "").toLowerCase();
  if(n.includes("sébastien duc") || n.includes("sebastien duc")) return "rsg";
  if(e === "sebastien_duc@outlook.fr" || e === "sebastien.duc@scse.fr") return "rsg";
  return "interne";
}
function normalizeTimeLogRole(log){
  const raw = String(log?.role || log?.roleKey || "").toLowerCase();
  if(raw.includes("rsg/ri")) return "rsg";
  if(raw.includes("rsg")) return "rsg";
  if(raw.includes("ri")) return "ri";
  if(raw.includes("externe")) return "externe";
  if(raw.includes("interne")) return "interne";
  return resolveTimeLogRole(log?.userName || "", log?.userEmail || "");
}
function getTaskRoleKey(t){
  const hasVendor = (t?.vendor || "").trim();
  if(hasVendor) return "externe";
  const typ = ownerType(t?.owner);
  if(typ === "rsg") return "rsg";
  if(typ === "ri") return "ri";
  if(typ === "externe" || typ === "inconnu") return "externe";
  return "interne";
}
function roleLabel(roleKey){
  if(roleKey==="rsg") return "RSG";
  if(roleKey==="ri") return "RI";
  if(roleKey==="externe") return "EXTERNE";
  return "INTERNE";
}
function getCanonicalTimeLogs(){
  const logs = getTimeLogs();
  const map = new Map(); // taskId|date|roleKey -> log
  logs.forEach(l=>{
    if(!l || !l.taskId || !l.date) return;
    const roleKey = normalizeTimeLogRole(l);
    const key = `${l.taskId}|${l.date}|${roleKey}`;
    const existing = map.get(key);
    if(!existing){
      map.set(key, l);
      return;
    }
    const prevTs = new Date(existing.updatedAt || existing.createdAt || 0).getTime();
    const curTs = new Date(l.updatedAt || l.createdAt || 0).getTime();
    if(curTs >= prevTs) map.set(key, l);
  });
  return Array.from(map.values());
}
function findTimeLog(taskId, dateKey, userKey, userEmail="", userName=""){
  return getTimeLogs().find(l=>
    l.taskId===taskId &&
    l.date===dateKey &&
    (
      (userKey && l.userKey===userKey) ||
      (userEmail && l.userEmail===userEmail) ||
      (userName && l.userName===userName)
    )
  ) || null;
}
function findTimeLogByRole(taskId, dateKey, roleKey){
  const logs = getCanonicalTimeLogs();
  return logs.find(l=>l.taskId===taskId && l.date===dateKey && normalizeTimeLogRole(l)===roleKey) || null;
}
function upsertTimeLog(taskId, projectId, minutes, note="", dateKeyOverride=null, roleKeyOverride=""){
  const userKey = getCurrentUserKey();
  const userName = getCurrentUserName();
  const userEmail = getCurrentUserEmail();
  const role = roleKeyOverride || resolveTimeLogRole(userName, userEmail);
  const dateKey = dateKeyOverride || getSelectedLogDate();
  if(!userKey) return null;
  const logs = getTimeLogs();
  const duplicates = logs.filter(l=>l.taskId===taskId && l.date===dateKey && normalizeTimeLogRole(l)===role);
  const existing = duplicates[0] || null;
  const now = new Date().toISOString();
  if(existing){
    existing.minutes = minutes;
    existing.note = note || existing.note || "";
    existing.updatedAt = now;
    existing.role = String(role || "").toUpperCase();
    existing.roleKey = role;
    for(let i=duplicates.length-1;i>=1;i--){
      const idx = logs.indexOf(duplicates[i]);
      if(idx>=0) logs.splice(idx,1);
    }
    return existing;
  }
  const entry = {
    id: uid(),
    taskId,
    projectId,
    userKey,
    userName,
    userEmail,
    role: String(role || "").toUpperCase(),
    roleKey: role,
    date: dateKey,
    minutes,
    note: note || "",
    createdAt: now,
    updatedAt: now
  };
  logs.push(entry);
  return entry;
}
function purgeTaskLogsByAssignedRole(task){
  if(!task || !task.id) return 0;
  const roleKey = getTaskRoleKey(task);
  const logs = getTimeLogs();
  const before = logs.length;
  state.timeLogs = logs.filter(l=>{
    if(!l || l.taskId !== task.id) return true;
    return normalizeTimeLogRole(l) === roleKey;
  });
  return before - state.timeLogs.length;
}
function formatHoursMinutes(totalMinutes){
  const h = Math.floor(totalMinutes / 60);
  const m = totalMinutes % 60;
  if(m === 0) return `${h} h`;
  return `${h} h ${m} min`;
}
function roleHoursMultiplier(roleKey){
  return roleKey === "interne" ? 2 : 1;
}
function getTaskTimeTotals(taskRef){
  const taskId = (typeof taskRef === "string") ? taskRef : taskRef?.id;
  const roleKey = (typeof taskRef === "object" && taskRef) ? getTaskRoleKey(taskRef) : "";
  const rangeStart = (typeof taskRef === "object" && taskRef && taskRef.start) ? taskRef.start : "";
  const rangeEnd = (typeof taskRef === "object" && taskRef && taskRef.end) ? taskRef.end : "";
  const totals = new Map();
  getCanonicalTimeLogs().filter(l=>{
    if(l.taskId!==taskId) return false;
    if(!roleKey) return true;
    return normalizeTimeLogRole(l)===roleKey;
  }).forEach(l=>{
    if(rangeStart && l.date < rangeStart) return;
    if(rangeEnd && l.date > rangeEnd) return;
    const normalizedRole = normalizeTimeLogRole(l);
    const key = roleLabel(normalizedRole);
    const weightedMinutes = Math.round((Number(l.minutes || 0)) * roleHoursMultiplier(normalizedRole));
    totals.set(key, (totals.get(key)||0) + weightedMinutes);
  });
  const items = Array.from(totals.entries()).map(([name, minutes])=>({name, minutes}));
  items.sort((a,b)=>b.minutes - a.minutes);
  const totalMinutes = items.reduce((acc,it)=>acc+it.minutes,0);
  return {items, totalMinutes};
}
function getRealMinutesForTasks(tasks){
  const ids = new Set((tasks || []).map(t=>t.id));
  const roleByTask = new Map((tasks || []).map(t=>[t.id, getTaskRoleKey(t)]));
  const rangeByTask = new Map((tasks || []).map(t=>[t.id, {start:t.start||"", end:t.end||""}]));
  let totalMinutes = 0;
  let internalMinutes = 0;
  let externalMinutes = 0;
  let rsgMinutes = 0;
  let riMinutes = 0;
  getCanonicalTimeLogs().forEach(l=>{
    if(!ids.has(l.taskId)) return;
    const range = rangeByTask.get(l.taskId);
    if(range?.start && l.date < range.start) return;
    if(range?.end && l.date > range.end) return;
    const roleExpected = roleByTask.get(l.taskId);
    const m = Number(l.minutes || 0);
    if(!m) return;
    const role = normalizeTimeLogRole(l);
    if(roleExpected && role !== roleExpected) return;
    const weightedMinutes = Math.round(m * roleHoursMultiplier(role));
    totalMinutes += weightedMinutes;
    if(role === "externe") externalMinutes += weightedMinutes;
    else if(role === "rsg") rsgMinutes += weightedMinutes;
    else if(role === "ri") riMinutes += weightedMinutes;
    else internalMinutes += weightedMinutes;
  });
  return { totalMinutes, internalMinutes, externalMinutes, rsgMinutes, riMinutes };
}
function countMissingDaysForUser(t, userKey){
  if(!t || !t.start || !t.end || !userKey) return 0;
  const roleKey = getTaskRoleKey(t);
  const start = new Date(t.start+"T00:00:00");
  const end = new Date(t.end+"T00:00:00");
  if(isNaN(start) || isNaN(end) || end < start) return 0;
  const yKey = getYesterdayKey();
  const limit = new Date(yKey+"T00:00:00");
  let missing = 0;
  for(let d=new Date(start); d<=end && d<=limit; d.setDate(d.getDate()+1)){
    if(!isWeekday(d)) continue;
    const key = toLocalDateKey(d);
    if(!findTimeLogByRole(t.id, key, roleKey)) missing++;
  }
  return missing;
}
function countMissingDaysForTask(t){
  return getMissingDaysList(t).length;
}
function getMissingDaysList(t){
  if(!t || !t.start || !t.end) return [];
  const roleKey = getTaskRoleKey(t);
  const start = new Date(t.start+"T00:00:00");
  const end = new Date(t.end+"T00:00:00");
  if(isNaN(start) || isNaN(end) || end < start) return [];
  const todayKey = toLocalDateKey(new Date());
  const out = [];
  for(let d=new Date(start); d<=end; d.setDate(d.getDate()+1)){
    if(!isWeekday(d)) continue;
    const key = toLocalDateKey(d);
    if(key >= todayKey) continue;
    if(!findTimeLogByRole(t.id, key, roleKey)) out.push(key);
  }
  return out;
}
function buildMissingDaysMap(tasks){
  const map = new Map();
  (tasks || []).forEach(t=>{
    map.set(t.id, countMissingDaysForTask(t));
  });
  return map;
}
function getMissingTasksForMasterFlow(){
  const sorted = sortTasks((state?.tasks || []), sortMaster);
  const missingMap = buildMissingDaysMap(sorted);
  return sorted
    .filter(t=> (missingMap.get(t.id) || 0) > 0)
    .map(t=>({ projectId: t.projectId, taskId: t.id }));
}
function finishMissingHoursFlow(){
  _missingHoursFlow = null;
  selectedProjectId = null;
  selectedTaskId = null;
  const toggleMissingOnly = el("toggleMissingOnly");
  if(toggleMissingOnly) toggleMissingOnly.checked = false;
  renderMaster();
}
function openMissingHoursFlowStep(){
  if(!_missingHoursFlow || !_missingHoursFlow.tasks.length) return;
  while(_missingHoursFlow.index < _missingHoursFlow.tasks.length){
    const step = _missingHoursFlow.tasks[_missingHoursFlow.index];
    const t = state?.tasks?.find(x=>x.id===step.taskId && x.projectId===step.projectId);
    if(!t){
      _missingHoursFlow.index += 1;
      continue;
    }
    selectedProjectId = step.projectId;
    selectedTaskId = step.taskId;
    renderProject();
    setTimeout(()=>{
      if(!_missingHoursFlow) return;
      const selected = getSelectedTaskForHoursModal();
      if(!selected){
        _missingHoursFlow.index += 1;
        openMissingHoursFlowStep();
        return;
      }
      openHoursTaskModal();
    }, 0);
    return;
  }
  showSaveToast("ok", "Parcours terminé", "Toutes les tâches à compléter ont été traitées.");
  finishMissingHoursFlow();
}
function startMissingHoursFlow(){
  const tasks = getMissingTasksForMasterFlow();
  if(!tasks.length){
    showSaveToast("ok", "A compléter", "Aucune tâche à compléter.");
    return;
  }
  _missingHoursFlow = { tasks, index: 0 };
  showSaveToast("ok", "Parcours lancé", `${tasks.length} tâche(s) à compléter`);
  openMissingHoursFlowStep();
}
function advanceMissingHoursFlow(){
  if(!_missingHoursFlow) return;
  _missingHoursFlow.index += 1;
  openMissingHoursFlowStep();
}
function updateTimeLogUI(t, forceAlert=false){
  const dateInput = el("t_time_date_input");
  const input = el("t_time_hours");
  const statusEl = el("t_time_status");
  const btn = el("btnSaveTimeLog");
  const summaryEl = el("t_time_summary");
  const yKey = getYesterdayKey();
  if(dateInput && !dateInput.value) dateInput.value = yKey;
  if(!t || !input || !statusEl || !btn){
    if(input) input.value = "";
    if(statusEl) statusEl.textContent = "";
    if(summaryEl) summaryEl.textContent = "";
    if(btn) btn.disabled = true;
    return;
  }
  const active = isTaskActiveOn(t, yKey);
  const userKey = getCurrentUserKey();
  const userEmail = getCurrentUserEmail();
  const userName = getCurrentUserName();
  if(!userKey){
    input.value = "";
    input.disabled = true;
    btn.disabled = true;
    statusEl.textContent = "Utilisateur non défini";
    if(summaryEl) summaryEl.textContent = "";
    return;
  }
  if(dateInput){
    const lastTask = dateInput.dataset.taskId || "";
    if(!dateInput.value || lastTask !== t.id){
      dateInput.value = t.start || yKey;
      dateInput.dataset.taskId = t.id;
    }
  }
  const selectedDate = dateInput?.value || yKey;
  if(dateInput){
    dateInput.min = t.start || "";
    dateInput.max = t.end || "";
  }
  input.disabled = false;
  const selectedDateObj = new Date(selectedDate+"T00:00:00");
  const inRange = isTaskActiveOn(t, selectedDate);
  const isWeek = isWeekday(selectedDateObj);
  btn.disabled = !(inRange && isWeek);
  const roleKey = getTaskRoleKey(t);
  const log = findTimeLogByRole(t.id, selectedDate, roleKey);
  statusEl.classList.remove("missing");
  if(log){
    const hours = (log.minutes/60);
    input.value = (Math.round(hours*100)/100).toString();
    statusEl.textContent = "Déjà renseigné";
  }else{
    input.value = "";
    if(inRange && isWeek){
      statusEl.classList.add("missing");
      statusEl.innerHTML = `<span class="time-icon missing"><svg viewBox="0 0 24 24" aria-hidden="true"><circle cx="12" cy="12" r="8"></circle><path d="M12 7v5l3 2"></path></svg></span>Heures réelles manquantes`;
    }else{
      statusEl.textContent = "Hors période";
    }
  }
  if(inRange && !isWeek){
    statusEl.textContent = "Week-end";
  }
  const missing = countMissingDaysForUser(t, userKey);
  if(summaryEl){
    const dayLogs = getCanonicalTimeLogs().filter(l=>l.taskId===t.id && l.date===selectedDate && normalizeTimeLogRole(l)===roleKey);
    const totalsDay = new Map();
    dayLogs.forEach(l=>{
      const role = normalizeTimeLogRole(l);
      const label = roleLabel(role);
      totalsDay.set(label, (totalsDay.get(label)||0) + (l.minutes||0));
    });
    const partsDay = Array.from(totalsDay.entries()).map(([name, minutes])=>`${name}: ${formatHoursMinutes(minutes)}`);
    const totalDayMinutes = Array.from(totalsDay.values()).reduce((a,b)=>a+b,0);
    const totalTask = getTaskTimeTotals(t);
    const missingList = getMissingDaysList(t).map(d=>{
      const dt = new Date(d+"T00:00:00");
      return formatShortDate(dt);
    });
    const missLabel = missingList.length ? ` | Jours manquants : ${missingList.join(", ")} (${missingList.length})` : "";
    if(partsDay.length){
      summaryEl.textContent = `${partsDay.join(" | ")} | Total jour: ${formatHoursMinutes(totalDayMinutes)} | Total tâche: ${formatHoursMinutes(totalTask.totalMinutes)}${missLabel}`;
    }else{
      summaryEl.textContent = `Total tâche: ${formatHoursMinutes(totalTask.totalMinutes)}${missLabel}`;
    }
  }
}

function getSelectedTaskForHoursModal(){
  if(!selectedProjectId || !selectedTaskId) return null;
  return state?.tasks?.find(x=>x.id===selectedTaskId && x.projectId===selectedProjectId) || null;
}
function isHoursTaskModalOpen(){
  const modal = el("hoursTaskModal");
  return !!(modal && !modal.classList.contains("hidden"));
}

function formatHoursDecimal(minutes){
  const v = Math.round(((minutes || 0) / 60) * 100) / 100;
  return String(v).replace(".", ",") + "h";
}
function ensureHoursModalCalendarDom(){
  const modalCard = document.querySelector("#hoursTaskModal .modal-card");
  if(!modalCard) return null;
  let wrap = el("hm_calendarWrap");
  if(!wrap){
    wrap = document.createElement("div");
    wrap.id = "hm_calendarWrap";
    wrap.style.marginTop = "10px";
    wrap.style.border = "1px solid var(--line)";
    wrap.style.borderRadius = "10px";
    wrap.style.padding = "10px";
    wrap.style.background = "#f8fafc";

    const legend = document.createElement("div");
    legend.id = "hm_calendarLegend";
    legend.style.display = "flex";
    legend.style.gap = "10px";
    legend.style.flexWrap = "wrap";
    legend.style.marginBottom = "8px";
    legend.style.fontSize = "12px";
    legend.innerHTML = [
      '<span><span style="display:inline-block;width:10px;height:10px;border-radius:50%;background:#22c55e;margin-right:6px"></span>Renseigné</span>',
      '<span><span style="display:inline-block;width:10px;height:10px;border-radius:50%;background:#f59e0b;margin-right:6px"></span>Manquant</span>',
      '<span><span style="display:inline-block;width:10px;height:10px;border-radius:50%;background:#64748b;margin-right:6px"></span>À venir</span>'
    ].join("");

    const grid = document.createElement("div");
    grid.id = "hm_calendar";
    grid.style.display = "block";
    grid.style.gap = "8px";

    wrap.appendChild(legend);
    wrap.appendChild(grid);

    const summary = el("hm_summary");
    if(summary && summary.parentElement){
      summary.parentElement.insertBefore(wrap, summary.nextSibling);
    }else{
      modalCard.appendChild(wrap);
    }
  }
  return wrap;
}
function renderHoursTaskCalendar(t){
  const wrap = ensureHoursModalCalendarDom();
  const grid = el("hm_calendar");
  if(!wrap || !grid) return;
  if(!t || !t.start || !t.end){
    grid.innerHTML = "";
    return;
  }

  const roleKey = getTaskRoleKey(t);
  const todayKey = toLocalDateKey(new Date());
  const logs = getCanonicalTimeLogs().filter(l=>l.taskId===t.id && normalizeTimeLogRole(l)===roleKey);
  const byDate = new Map();
  logs.forEach(l=>byDate.set(l.date, l));

  const start = new Date(t.start + "T00:00:00");
  const end = new Date(t.end + "T00:00:00");
  if(isNaN(start) || isNaN(end) || end < start){
    grid.innerHTML = "";
    return;
  }

  const startDayIdx = (start.getDay() + 6) % 7; // lundi=0
  const endDayIdx = (end.getDay() + 6) % 7;
  const firstMonday = new Date(start);
  firstMonday.setDate(start.getDate() - startDayIdx);
  const lastFriday = new Date(end);
  lastFriday.setDate(end.getDate() + (4 - endDayIdx));

  const selectedDate = (el("hm_date")?.value || "").trim();
  const rows = [];
  const dayNames = ["Lun", "Mar", "Mer", "Jeu", "Ven"];

  for(let ws = new Date(firstMonday); ws <= lastFriday; ws.setDate(ws.getDate()+7)){
    const weekNo = String(isoWeekInfo(ws).week).padStart(2, "0");
    const weekLabel = "S" + weekNo;
    const cards = [];

    for(let i=0; i<5; i++){
      const d = new Date(ws);
      d.setDate(ws.getDate() + i);
      const key = toLocalDateKey(d);
      const log = byDate.get(key) || null;
      const inTaskRange = key >= t.start && key <= t.end;
      const stateType = inTaskRange ? (log ? "filled" : (key < todayKey ? "missing" : "future")) : "outside";
      const dateLabel = d.toLocaleDateString("fr-FR", { day:"2-digit", month:"2-digit" });
      const hoursValue = log ? String(Math.round(((log.minutes || 0) / 60) * 100) / 100).replace(".", ",") : "";

      let bg = "#ffffff";
      let border = "#d0d7e2";
      let text = "#0f172a";
      if(stateType === "filled"){
        bg = "#ecfdf3";
        border = "#7dd3a3";
      }else if(stateType === "missing"){
        bg = "#fff7ed";
        border = "#fdba74";
      }else if(stateType === "outside"){
        bg = "#e5e7eb";
        border = "#cbd5e1";
        text = "#64748b";
      }

      const isEditable = inTaskRange && key <= todayKey;
      const selectedBorder = (selectedDate === key && isEditable) ? "2px solid #2563eb" : ("1px solid " + border);
      const sub = hoursValue ? (hoursValue + "h") : (stateType === "missing" ? "manquant" : (stateType === "outside" ? "hors tâche" : "—"));
      const activeAttr = isEditable ? "1" : "0";
      const disabledAttr = isEditable ? "" : "disabled";

      cards.push(
        '<div class="hm-day" data-date="' + key + '" data-active="' + activeAttr + '" data-state-base="' + stateType + '" style="text-align:left;border:' + selectedBorder + ';background:' + bg + ';color:' + text + ';border-radius:8px;padding:3px 5px;min-height:52px;cursor:' + (inTaskRange ? 'pointer' : 'default') + ';display:flex;flex-direction:column;gap:1px">' +
        '<div style="font-size:11px;line-height:1.1;opacity:.85">' + dayNames[i] + ' • ' + weekLabel + '</div>' +
        '<div style="font-size:12px;font-weight:700;line-height:1.2">' + dateLabel + '</div>' +
        '<input type="text" inputmode="decimal" class="hm-day-input" data-date="' + key + '" data-active="' + activeAttr + '" value="' + hoursValue + '" placeholder="h" ' + disabledAttr + ' style="margin-top:1px;height:18px;border:1px solid #cbd5e1;border-radius:6px;padding:1px 6px;background:#fff;color:#111827;font-size:11px" />' +
        '<div class="hm-day-sub" style="font-size:11px;opacity:.92">' + sub + '</div>' +
        '</div>'
      );
    }

    let weekDays = 0;
    const weekTotalMinutes = cards.reduce((acc, _c, idx)=>{
      const d2 = new Date(ws);
      d2.setDate(ws.getDate() + idx);
      const k2 = toLocalDateKey(d2);
      const inTask2 = k2 >= t.start && k2 <= t.end;
      if(!inTask2) return acc;
      weekDays += 1;
      const l2 = byDate.get(k2) || null;
      return acc + ((l2 && l2.minutes) ? l2.minutes : 0);
    }, 0);
    const weekAvgMinutes = weekDays > 0 ? Math.round(weekTotalMinutes / weekDays) : 0;

    rows.push(
      '<div class="hm-week-row" style="display:grid;grid-template-columns:64px repeat(5, minmax(82px,1fr)) 156px;gap:5px;align-items:stretch;margin-bottom:4px">' +
      '<div style="border:1px solid #cbd5e1;border-radius:8px;background:#eef2ff;color:#1e293b;display:flex;align-items:center;justify-content:center;font-weight:700;font-size:12px">' + weekLabel + '</div>' +
      cards.join("") +
      '<div style="border:1px solid #cbd5e1;border-radius:8px;background:#f8fafc;color:#0f172a;padding:5px 7px;display:flex;flex-direction:column;justify-content:center;gap:2px">' +
      '<div style="font-size:11px;color:#475569">Total semaine</div>' +
      '<div style="font-size:12px;font-weight:700">' + formatHoursMinutes(weekTotalMinutes) + '</div>' +
      '<div style="font-size:11px;color:#475569">Moyenne semaine</div>' +
      '<div style="font-size:12px;font-weight:700">' + formatHoursMinutes(weekAvgMinutes) + '</div>' +
      '</div>' +
      '</div>'
    );
  }

  grid.innerHTML = rows.join("");
}

function collectHoursTaskCalendarEntries(t){
  const grid = el("hm_calendar");
  if(!t || !grid) return [];
  const todayKey = toLocalDateKey(new Date());
  const entries = [];
  grid.querySelectorAll(".hm-day-input[data-date][data-active='1']").forEach((input)=>{
    const date = (input.getAttribute("data-date") || "").trim();
    const raw = (input.value || "").toString().replace(",", ".").trim();
    if(!date || raw === "") return;
    if(date > todayKey) return;
    const hours = parseFloat(raw);
    if(!isFinite(hours) || hours < 0) return;
    entries.push({ date, hours, minutes: Math.round(hours * 60) });
  });
  return entries;
}
function getHoursDraftForDate(dateKey){
  if(!dateKey) return "";
  const grid = el("hm_calendar");
  if(!grid) return "";
  const input = grid.querySelector(`.hm-day-input[data-date="${dateKey}"]`);
  return (input?.value || "").toString();
}

function refreshHoursDayCardVisual(input){
  const dayInput = input?.closest?.(".hm-day-input[data-date]");
  if(!dayInput) return;
  const card = dayInput.closest(".hm-day[data-date]");
  if(!card) return;
  const day = (dayInput.getAttribute("data-date") || "").trim();
  const active = (dayInput.getAttribute("data-active") || card.getAttribute("data-active") || "0") === "1";
  const baseState = (card.getAttribute("data-state-base") || "future").toLowerCase();
  const hmDate = el("hm_date");
  const selectedDay = (hmDate?.value || "").trim();
  const isSelected = !!day && day === selectedDay && active;

  const raw = (dayInput.value || "").toString().replace(",", ".").trim();
  const parsed = parseFloat(raw);
  const hasValidHours = raw !== "" && isFinite(parsed) && parsed >= 0;
  const state = hasValidHours ? "filled" : baseState;

  let bg = "#ffffff";
  let border = "#d0d7e2";
  let text = "#0f172a";
  if(state === "filled"){
    bg = "#ecfdf3";
    border = "#7dd3a3";
  }else if(state === "missing"){
    bg = "#fff7ed";
    border = "#fdba74";
  }else if(state === "outside"){
    bg = "#e5e7eb";
    border = "#cbd5e1";
    text = "#64748b";
  }

  card.style.background = bg;
  card.style.color = text;
  card.style.border = isSelected ? "2px solid #2563eb" : ("1px solid " + border);

  const sub = card.querySelector(".hm-day-sub");
  if(sub){
    if(hasValidHours){
      const shown = String(Math.round(parsed * 100) / 100).replace(".", ",");
      sub.textContent = shown + "h";
    }else if(state === "missing"){
      sub.textContent = "manquant";
    }else if(state === "outside"){
      sub.textContent = "hors tâche";
    }else{
      sub.textContent = "—";
    }
  }
}


function computeTaskWeeklySummary(t){
  if(!t || !t.start || !t.end) return { rows: [], totalMinutes: 0 };
  const roleKey = getTaskRoleKey(t);
  const roleMultiplier = roleHoursMultiplier(roleKey);
  const logs = getCanonicalTimeLogs().filter(l=>l.taskId===t.id && normalizeTimeLogRole(l)===roleKey);
  const byDate = new Map();
  logs.forEach((l)=> byDate.set(l.date, l.minutes || 0));

  const start = new Date(t.start + "T00:00:00");
  const end = new Date(t.end + "T00:00:00");
  if(isNaN(start) || isNaN(end) || end < start) return { rows: [], totalMinutes: 0 };

  const rowsMap = new Map();
  let totalMinutes = 0;
  for(let d=new Date(start); d<=end; d.setDate(d.getDate()+1)){
    if(!isWeekday(d)) continue;
    const key = toLocalDateKey(d);
    const info = isoWeekInfo(d);
    const wk = `${info.year}-S${String(info.week).padStart(2, "0")}`;
    const label = `S${String(info.week).padStart(2, "0")} (${info.year})`;
    if(!rowsMap.has(wk)) rowsMap.set(wk, { key:wk, label, days:0, total:0 });
    const row = rowsMap.get(wk);
    row.days += 1;
    const dayMinutes = Math.round((byDate.get(key) || 0) * roleMultiplier);
    row.total += dayMinutes;
    totalMinutes += dayMinutes;
  }
  return { rows: Array.from(rowsMap.values()), totalMinutes };
}

function renderHoursTaskWeeklySummary(t, draftEntries=null){
  const box = el("hm_summary");
  if(!box) return;
  if(!t){ box.innerHTML = ""; return; }
  let totalMinutes = 0;
  if(Array.isArray(draftEntries)){
    const roleMultiplier = roleHoursMultiplier(getTaskRoleKey(t));
    const byDate = new Map();
    draftEntries.forEach((e)=>{
      const date = (e?.date || "").toString().trim();
      if(!date) return;
      if(!isTaskActiveOn(t, date)) return;
      const d = new Date(date + "T00:00:00");
      if(!isWeekday(d)) return;
      byDate.set(date, Math.round(Math.max(0, Number(e?.minutes || 0)) * roleMultiplier));
    });
    totalMinutes = Array.from(byDate.values()).reduce((a,b)=>a+b,0);
  }else{
    const data = computeTaskWeeklySummary(t);
    totalMinutes = data.totalMinutes || 0;
  }
  box.innerHTML = `<div style="margin-top:8px;border:1px solid #cbd5e1;border-radius:10px;background:#fff;padding:10px;color:#0f172a;font-weight:700;font-size:13px">Total tâche : ${formatHoursMinutes(totalMinutes)}</div>`;
}

function syncHoursTaskStatusFromMain(){
  const statusInput = el("t_time_status");
  const hmStatus = el("hm_status");
  if(!hmStatus) return;
  hmStatus.className = (statusInput?.className || "time-log-status");
  hmStatus.innerHTML = statusInput?.innerHTML || "";
  if((hmStatus.textContent || "").toLowerCase().includes("non renseign")){
    hmStatus.textContent = "";
  }
}

function syncHoursTaskStatusFromCalendarDraft(t, dayKey, rawValue){
  const hmStatus = el("hm_status");
  if(!hmStatus) return;
  hmStatus.className = "time-log-status";
  if(!t || !dayKey){
    hmStatus.textContent = "";
    return;
  }
  const inRange = isTaskActiveOn(t, dayKey);
  if(!inRange){
    hmStatus.textContent = "Hors période";
    return;
  }
  const dateObj = new Date(dayKey + "T00:00:00");
  if(!isWeekday(dateObj)){
    hmStatus.textContent = "Week-end";
    return;
  }
  const todayKey = toLocalDateKey(new Date());
  if(dayKey > todayKey){
    hmStatus.textContent = "À venir";
    return;
  }
  const raw = (rawValue == null ? "" : String(rawValue)).replace(",", ".").trim();
  if(raw === ""){
    hmStatus.classList.add("missing");
    hmStatus.innerHTML = `<span class="time-icon missing"><svg viewBox="0 0 24 24" aria-hidden="true"><circle cx="12" cy="12" r="8"></circle><path d="M12 7v5l3 2"></path></svg></span>Heures réelles manquantes`;
    return;
  }
  const hours = parseFloat(raw);
  if(!isFinite(hours) || hours < 0){
    hmStatus.classList.add("missing");
    hmStatus.textContent = "Temps invalide";
    return;
  }
  hmStatus.textContent = "Déjà renseigné";
  if((hmStatus.textContent || "").toLowerCase().includes("non renseign")){
    hmStatus.textContent = "";
  }
}

function applyHoursSaveButtonVisualState(btn){
  if(!btn) return;
  // Palette harmonisée UI: bouton Valider bleu-gris (plus de vert forcé).
  btn.style.setProperty("background", "#3f6170", "important");
  btn.style.setProperty("border-color", "#365563", "important");
  btn.style.setProperty("color", "#ffffff", "important");
  btn.style.setProperty("-webkit-text-fill-color", "#ffffff", "important");
  btn.style.setProperty("text-shadow", "none", "important");
  if(btn.disabled || btn.classList.contains("is-disabled")){
    btn.style.setProperty("background", "#8ea2ad", "important");
    btn.style.setProperty("border-color", "#7f939e", "important");
    btn.style.setProperty("color", "#f5f8fa", "important");
    btn.style.setProperty("-webkit-text-fill-color", "#f5f8fa", "important");
    btn.style.setProperty("opacity", "1", "important");
    btn.style.setProperty("filter", "none", "important");
    btn.style.setProperty("box-shadow", "none", "important");
  }else{
    btn.style.removeProperty("opacity");
    btn.style.removeProperty("filter");
    btn.style.removeProperty("box-shadow");
  }
}

function syncHoursTaskModal(taskOverride=null){
  const modal = el("hoursTaskModal");
  if(!modal) return;
  const t = taskOverride || getSelectedTaskForHoursModal();
  const p = t ? (state?.projects || []).find(x=>x.id===t.projectId) : null;

  const hmProject = el("hm_project");
  const hmTask = el("hm_task");
  const hmOwnerBadge = el("hm_owner_badge");
  const hmPeriod = el("hm_period");
  const hmDate = el("hm_date");
  const hmHours = el("hm_hours");
  const hmStatus = el("hm_status");
  const hmSummary = el("hm_summary");
  const btnSave = el("btnSaveHoursModal");

  if(!t){
    if(hmProject) hmProject.value = "";
    if(hmTask) hmTask.value = "";
    if(hmOwnerBadge) hmOwnerBadge.innerHTML = "";
    if(hmPeriod) hmPeriod.value = "";
    if(hmDate) hmDate.value = "";
    if(hmHours) hmHours.value = "";
    if(hmStatus) hmStatus.textContent = "";
    if(hmSummary) hmSummary.textContent = "";
    if(btnSave){
      btnSave.disabled = true;
      applyHoursSaveButtonVisualState(btnSave);
    }
    return;
  }

  const dateInput = el("t_time_date_input");
  const hoursInput = el("t_time_hours");
  const statusInput = el("t_time_status");
  const desc = (t.roomNumber || "").trim();
  const owner = roleLabel(getTaskRoleKey(t));
  const vendor = (t.vendor || "").trim();

  if(hmProject) hmProject.value = (p?.name || "").trim();
  if(hmTask) hmTask.value = desc ? (((t.status || "").trim()) + " - " + desc) : ((t.status || "").trim() || "Tâche");
  if(hmOwnerBadge){
    const badgeHtml = ownerBadgeForTask(t) || ownerBadge(owner, vendor ? `${owner} - ${vendor}` : owner);
    hmOwnerBadge.innerHTML = badgeHtml || `<span>${attrEscape(vendor ? `${owner} - ${vendor}` : owner)}</span>`;
  }
  if(hmPeriod) hmPeriod.value = (formatDate(t.start || "") + " -> " + formatDate(t.end || ""));
  if(hmDate){
    const todayKey = toLocalDateKey(new Date());
    const maxAllowed = (!t.end || t.end > todayKey) ? todayKey : t.end;
    hmDate.min = t.start || "";
    hmDate.max = maxAllowed || "";
    hmDate.value = (dateInput?.value || t.start || "");
    if(hmDate.value && maxAllowed && hmDate.value > maxAllowed) hmDate.value = maxAllowed;
  }
  if(hmHours) hmHours.value = (hoursInput?.value || "").toString();
  syncHoursTaskStatusFromMain();
  renderHoursTaskWeeklySummary(t);
  if(btnSave){
    btnSave.disabled = false;
    applyHoursSaveButtonVisualState(btnSave);
  }
  renderHoursTaskCalendar(t);
}
function openHoursTaskModal(){
  const modal = el("hoursTaskModal");
  if(!modal) return;
  const t = getSelectedTaskForHoursModal();
  if(!t){
    alert("Sélectionne une tâche.");
    return;
  }
  updateTimeLogUI(t, true);
  syncHoursTaskModal(t);
  showModalSafely(modal);
}
function closeHoursTaskModal(stopFlow=true){
  const modal = el("hoursTaskModal");
  hideModalSafely(modal, "#btnOpenHoursModal");
  if(stopFlow && _missingHoursFlow){
    _missingHoursFlow = null;
    showSaveToast("ok", "Parcours arrêté", "Traitement des tâches à compléter interrompu.");
  }
}
function saveHoursTaskModal(){
  const t = getSelectedTaskForHoursModal();
  if(!t){
    alert("Sélectionne une tâche.");
    return;
  }
  const hmDate = el("hm_date");
  const hmHours = el("hm_hours");
  const dateInput = el("t_time_date_input");
  const hoursInput = el("t_time_hours");
  if(!dateInput || !hoursInput || !hmDate || !hmHours){
    alert("Saisie des heures indisponible.");
    return;
  }

  const selectedDate = (hmDate.value || "").trim();
  if(!selectedDate || !isTaskActiveOn(t, selectedDate)){
    alert("La date est hors période de la tâche.");
    return;
  }
  const dateObj = new Date(selectedDate + "T00:00:00");
  if(!isWeekday(dateObj)){
    alert("La date tombe un week-end.");
    return;
  }
  const todayKey = toLocalDateKey(new Date());
  if(selectedDate > todayKey){
    alert("La date est dans le futur.");
    return;
  }

  const entries = collectHoursTaskCalendarEntries(t);
  const grid = el("hm_calendar");
  const emptyDates = [];
  (grid ? Array.from(grid.querySelectorAll(".hm-day-input[data-date][data-active='1']")) : []).forEach((input)=>{
    const date = (input.getAttribute("data-date") || "").trim();
    if(!date) return;
    const raw = (input.value || "").toString().replace(",", ".").trim();
    if(raw === "") emptyDates.push(date);
  });

  const rawHours = (hmHours.value || "").toString().replace(",", ".").trim();
  if(rawHours){
    const hours = parseFloat(rawHours);
    if(!isFinite(hours) || hours < 0){
      alert("Temps invalide.");
      return;
    }
    if(!entries.some((x)=>x.date === selectedDate)){
      entries.push({ date: selectedDate, hours, minutes: Math.round(hours * 60) });
    }
  }
  if(!entries.length && !emptyDates.length){
    alert("Saisis le temps passé (heures).");
    return;
  }

  saveUndoSnapshot();
  const roleKey = getTaskRoleKey(t);
  entries.forEach((entry)=>{
    if(!isTaskActiveOn(t, entry.date)) return;
    const dObj = new Date(entry.date + "T00:00:00");
    if(!isWeekday(dObj)) return;
    upsertTimeLog(t.id, t.projectId, entry.minutes, "", entry.date, roleKey);
  });
  if(emptyDates.length){
    const logs = getTimeLogs();
    for(let i = logs.length - 1; i >= 0; i--){
      const l = logs[i];
      if(!l || l.taskId !== t.id) continue;
      if(!emptyDates.includes(l.date)) continue;
      if(normalizeTimeLogRole(l) !== roleKey) continue;
      logs.splice(i, 1);
    }
  }
  closeHoursTaskModal(false);
  markDirty();
  dateInput.value = selectedDate;
  hoursInput.value = rawHours;
  updateTimeLogUI(t, true);
  renderMaster();
  renderProject();
  saveState();
  if(_missingHoursFlow){
    advanceMissingHoursFlow();
  }
}
function checkTimeLogReminders(){
  const userKey = getCurrentUserKey();
  if(!userKey) return;
  const yKey = getYesterdayKey();
  const flag = `timeLogReminder_${yKey}_${userKey}`;
  if(sessionStorage.getItem(flag)) return;
  const tasks = (state?.tasks || []).filter(t=>t.start && t.end && isTaskActiveOn(t, yKey));
  if(!tasks.length) return;
  let missing = 0;
  tasks.forEach(t=>{
    if(!findTimeLog(t.id, yKey, userKey)) missing++;
  });
  if(missing > 0){
    showSaveToast("ok", "Rappel temps (veille)", `${missing} tâche(s) à compléter`);
  }
  sessionStorage.setItem(flag, "1");
}



function renderProject(){
  computeTaskOrderMap();
  renderTabs();
  closeAllOverlays();
  ensureProjectHeaderNodes();
  const p=state.projects.find(x=>x.id===selectedProjectId);
  if(!p){ selectedProjectId=null; renderMaster(); return; }

  el("viewMaster")?.classList.add("hidden");

  el("viewProject")?.classList.remove("hidden");

  const projectTitleNode = el("projectTitle");
  if(projectTitleNode) projectTitleNode.textContent = `Projet : ${p.name||"Sans nom"}`;

  const projectSubNode = el("projectSub");
  if(projectSubNode) projectSubNode.textContent = p.site || "Détails • Gantt";
  updateSitePhoto(p.site || "");

  const projectSummary = el("projectSummary");
  let projectProgress = 0;
  if(projectSummary){
    const tasksAll = state.tasks.filter(t=>t.projectId===p.id);
    const tasksDated = tasksAll.filter(t=>t.start && t.end);
    const todayKey = new Date().toISOString().slice(0,10);
    const inProgress = tasksDated.filter(t=>t.start<=todayKey && t.end>=todayKey).length;
    const minStart = tasksDated.length ? tasksDated.map(t=>t.start).sort()[0] : "";
    const maxEnd = tasksDated.length ? tasksDated.map(t=>t.end).sort().slice(-1)[0] : "";
    const endLabel = maxEnd ? formatDate(maxEnd) : "—";
    const isFinished = !!(maxEnd && maxEnd < todayKey);
    const finishedLabel = isFinished ? " • Projet terminé" : "";
    if(minStart && maxEnd){
      projectProgress = taskProgress({start:minStart, end:maxEnd});
    }
    projectSummary.textContent = `Résumé : ${tasksAll.length} tâches • ${inProgress} en cours • Fin prévue : ${endLabel}${finishedLabel}`;
  }
  setProjectProgressUI(projectProgress);

  // métriques projet : durée totale + équivalent heures (6h/j)

  const projTasks = state.tasks.filter(t=>t.projectId===p.id && t.start && t.end);

  const allDays = new Set();

  const internalDays = new Set();
  const externalDays = new Set();
  const rsgDays = new Set();
  const riDays = new Set();

  projTasks.forEach(t=>{

    const s=new Date(t.start+"T00:00:00");

    const e=new Date(t.end+"T00:00:00");

    if(isNaN(s)||isNaN(e) || e<s) return;

    const typ = ownerType(t.owner);
    const ownsInternal = typ === "interne";
    const ownsExternal = typ === "externe" || typ === "inconnu";
    const ownsRsg = typ === "rsg";
    const ownsRi = typ === "ri";

    for(let d=new Date(s); d<=e; d.setDate(d.getDate()+1)){

      const key=d.toISOString().slice(0,10);

      allDays.add(key);

      if(ownsInternal) internalDays.add(key);
      if(ownsExternal) externalDays.add(key);
      if(ownsRsg) rsgDays.add(key);
      if(ownsRi) riDays.add(key);

    }

  });

  const totalDays = allDays.size;

  const real = getRealMinutesForTasks(projTasks);
  let progSumProj = 0;
  let progWeightProj = 0;
  projTasks.forEach(t=>{
    const w = Math.max(1, durationDays(t.start, t.end));
    progSumProj += taskProgress(t) * w;
    progWeightProj += w;
  });
  const avgProgressProj = progWeightProj ? Math.round(progSumProj / progWeightProj) : 0;

  const metrics = el("projectMetrics");

  if(metrics){

    metrics.innerHTML = `

      <span class="panel-chip">Durée totale : <span class="metric-val">${totalDays || 0} j</span></span>

      <span class="panel-chip">Avancement : <span class="metric-val">${avgProgressProj}%</span></span>

      <span class="panel-chip">Heures réelles : <span class="metric-val">${formatHoursMinutes(real.totalMinutes||0)}</span></span>

      <span class="panel-chip" style="background:#e8eef8;color:#1f2937;border-color:#c7d2fe;">Interne : <span class="metric-val">${internalDays.size||0} j</span> <span class="metric-val">${formatHoursMinutes(real.internalMinutes||0)}</span></span>

      <span class="panel-chip" style="background:#fff1e6;color:#1f2937;border-color:#fdba74;">Externe : <span class="metric-val">${externalDays.size||0} j</span> <span class="metric-val">${formatHoursMinutes(real.externalMinutes||0)}</span></span>
      <span class="panel-chip" style="background:#e8f0ff;color:#1f2937;border-color:#93c5fd;">RSG : <span class="metric-val">${rsgDays.size||0} j</span> <span class="metric-val">${formatHoursMinutes(real.rsgMinutes||0)}</span></span>
      <span class="panel-chip" style="background:#f2eaff;color:#1f2937;border-color:#c4b5fd;">RI : <span class="metric-val">${riDays.size||0} j</span> <span class="metric-val">${formatHoursMinutes(real.riMinutes||0)}</span></span>

    `;

  }



  // Bandeau live : tâches en cours  la date du jour

  const live = el("projectLive");

  if(live){

    const todayKey = new Date().toISOString().slice(0,10);
    const datedProjectTasks = state.tasks.filter(t=>t.projectId===p.id && t.start && t.end);
    const hasStarted = datedProjectTasks.some(t=>t.start<=todayKey);
    const allFinished = datedProjectTasks.length>0 && datedProjectTasks.every(t=>t.end<todayKey || taskProgress(t)>=100);

    const inProgress = state.tasks

      .filter(t=>t.projectId===p.id && t.start && t.end && t.start<=todayKey && t.end>=todayKey)

      .sort((a,b)=> (taskOrderMap[a.id]||999)-(taskOrderMap[b.id]||999));

    if(inProgress.length===0){
      if(allFinished){
        live.innerHTML = `<span class="live-title">Projet terminé</span>`;
      }else if(!hasStarted){
        live.innerHTML = `<span class="live-title">Projet non démarré</span>`;
      }else{
        live.innerHTML = `<span class="live-title">Projet en attente</span>`;
      }

    }else{

      const badges = inProgress.map(t=>{

        const num = taskOrderMap[t.id]||"";

        const status = parseStatuses(t.status)[0] || "";

        const color = ownerColor(t.owner);

        const label = STATUSES.find(s=>s.v===status)?.label || status || "En cours";

        const projName = p.name || "Projet";

        return `<span class="live-item"><span class="num-badge" style="--badge-color:${color};--badge-text:#fff;">${num}</span> ${projName}  ${label}</span>`;

      }).join(" ");

      live.innerHTML = `<span class="live-title">Projet démarré  Tâches en cours :</span> ${badges}`;

    }

  }

  setInputValue("p_name", p.name||"");

  setInputValue("p_subproject", p.subproject||"");

  const siteSelect = el("p_site");
  if(siteSelect){
    const sites = ["CDM","Collège","École","LGT","NDC"].sort((a,b)=>a.localeCompare(b, "fr"));
    siteSelect.innerHTML = sites.map(s=>`<option value="${s}">${s}</option>`).join("");
    siteSelect.value = p.site || "";
    if(!siteSelect.value && p.site){ siteSelect.value = p.site; }
    const syncPhoto = () => updateSitePhoto(siteSelect.value);
    siteSelect.onchange = syncPhoto;
    siteSelect.oninput = syncPhoto;
    updateSitePhoto(siteSelect.value || p.site || "");
  }
  const constraintsInput = el("p_constraints");
  if(constraintsInput){
    constraintsInput.value = p.constraints||"";
  }



  let t=null;

  if(selectedTaskId){

    t = state.tasks.find(x=>x.id===selectedTaskId && x.projectId===p.id) || null;

  }

  if(!t){

    t = state.tasks.find(x=>x.projectId===p.id) || null;

  }

  selectedTaskId = t?.id || null;

  const badge = el("t_num_badge");

  if(badge){

    const num = selectedTaskId ? (taskOrderMap[selectedTaskId] || "") : "";

    badge.textContent = num;

    badge.style.display = num ? "inline-flex" : "none";

  }

  if(!selectedTaskId){

    el("btnNewTask")?.classList.add("btn-armed");

  }else{

    el("btnNewTask")?.classList.remove("btn-armed");

  }



  if(t){

    const desc = (t.roomNumber && t.roomNumber.trim()) || p.subproject || "";

    setInputValue("t_room", desc);

    const ownerVal = (t.owner || "");
    setInputValue("t_owner", ownerVal.toUpperCase()==="RSG/RI" ? "RSG" : ownerVal);

    setInputValue("t_vendor", t.vendor||"");

    const startVal = toInputDate(t.start);
    const endVal = toInputDate(t.end);
    setInputValue("t_start", startVal);

    setInputValue("t_end", endVal);
    setTaskProgressUI(taskProgress(t));
    updateTimeLogUI(t, true);
    if(window.__fpStart){ try{ window.__fpStart.setDate(startVal || null, true, "Y-m-d"); }catch(e){ softCatch(e); } }
    if(window.__fpEnd){ try{ window.__fpEnd.setDate(endVal || null, true, "Y-m-d"); }catch(e){ softCatch(e); } }

    setStatusSelection(t.status||"");

  }else{

    setInputValue("t_room", ""); setInputValue("t_owner", ""); setInputValue("t_vendor", ""); setInputValue("t_start", ""); setInputValue("t_end", "");
    setTaskProgressUI(0);
    updateTimeLogUI(null);
    if(window.__fpStart){ try{ window.__fpStart.setDate(null); }catch(e){ softCatch(e); } }
    if(window.__fpEnd){ try{ window.__fpEnd.setDate(null); }catch(e){ softCatch(e); } }

    setStatusSelection("");

  }



  renderGantt(p.id);

  renderProjectTasks(p.id);

  const projectTasks = state.tasks.filter(t=>t.projectId===p.id);

  const projRange = {type:workloadRangeTypeProject, year:workloadRangeYearProject, start:workloadRangeStartProject, end:workloadRangeEndProject};

  renderWorkloadChartFor(

    projectTasks,

    "workloadChartProject",

    "workloadPieProject",

    {type:"workloadRangeTypeProject", year:"workloadRangeYearProject", start:"workloadRangeStartProject", end:"workloadRangeEndProject"},

    projRange,

    projectTasks,

    true

  );

  workloadRangeTypeProject = projRange.type;

  workloadRangeYearProject = projRange.year;

  workloadRangeStartProject = projRange.start;

  workloadRangeEndProject = projRange.end;

  refreshVendorsList();

  refreshDescriptionsList();
  const projectHoursReport = el("projectHoursReport");
  if(projectHoursReport){
    projectHoursReport.innerHTML = buildProjectRealHoursReportInnerHTML(p.id);
  }
  animateBadgeChanges(el("viewProject"));
  animateCardsInView("viewProject");

}



function renderAll(){
  // filet de sécurité : si localStorage est vide (ex : fichier ouvert en navigation privée), on recharge l'état par défaut
  if(!state || !Array.isArray(state.projects) || state.projects.length===0){
    state = defaultState();
  }
  closeAllOverlays();
  refreshVendorsList();
  refreshDescriptionsList();
  resetProjectWorkloadFilters();
  // rinitialiser les filtres visibles pour éviter un filtrage bloquant

  ["filterSite","filterProject","filterStatus","filterSearch","filterStartAfter","filterEndBefore"].forEach(id=>{

    const n=el(id);

    if(n) n.value="";

  });
  const toggleMissingOnly = el("toggleMissingOnly");
  if(toggleMissingOnly) toggleMissingOnly.checked = false;
  _filteredCache = { key:"", version:-1, tasks:null };

  renderFilters();

  renderTabs();

  if(selectedProjectId) renderProject();

  else renderMaster();

  updateSidebarTop();
  updateSidebarScrollState();

  applySidebarTopLock();
  checkTimeLogReminders();
  updateDataQualityBanner(false);
  applyUiUpperNoAccent();
  ensureUiUpperNoAccentObserver();

}




function getExportFiltersPayload(){
  const site = el("filterSite")?.value || "";
  const project = el("filterProject")?.value || "";
  const status = el("filterStatus")?.value || "";
  const search = (el("filterSearch")?.value || "").trim();
  const startAfter = el("filterStartAfter")?.value || "";
  const endBefore = el("filterEndBefore")?.value || "";
  const projectName = selectedProjectId ? (state.projects.find(p=>p.id===selectedProjectId)?.name || "") : "";
  return { site, project, projectName, status, search, startAfter, endBefore };
}

function getProjectsSortedForExport(){
  return [...(state?.projects || [])].sort((a,b)=>
    (a?.name || "").localeCompare((b?.name || ""), "fr", { sensitivity:"base" })
  );
}

function getUnifiedDefaultExportProjectIds(){
  const projects = getProjectsSortedForExport();
  if(!projects.length) return [];
  if(selectedProjectId && projects.some(p=>p.id===selectedProjectId)) return [selectedProjectId];
  return [projects[0].id];
}

function normalizeUnifiedExportProjectIds(ids, opts={}){
  const allowEmpty = !!opts.allowEmpty;
  const allowed = new Set(getProjectsSortedForExport().map(p=>p.id));
  const uniq = Array.from(new Set((ids || []).map(normId).filter(Boolean)));
  const out = uniq.filter(id=>allowed.has(id));
  if(out.length) return out;
  return allowEmpty ? [] : getUnifiedDefaultExportProjectIds();
}

function getUnifiedExportSelectedProjectIdsFromUi(root){
  const scope = root || document;
  const nodes = Array.from(scope.querySelectorAll("#exportPdfModulesList input[data-export-project-id]:checked"));
  const ids = nodes.map(n=>normId(n.getAttribute("data-export-project-id")));
  return normalizeUnifiedExportProjectIds(ids, { allowEmpty:true });
}

function getActiveProjectIdsForToday(){
  const todayKey = new Date().toISOString().slice(0,10);
  const activeSet = new Set(
    (state.tasks || [])
      .filter((t)=>t?.projectId && t.start && t.end && t.start <= todayKey && t.end >= todayKey)
      .map((t)=>t.projectId)
  );
  return normalizeUnifiedExportProjectIds(Array.from(activeSet), { allowEmpty:true });
}

function getUnifiedExportModuleDefinitions(){
  const isProject = !el("viewProject")?.classList.contains("hidden");
  if(isProject){
    return [
      { key:"project_header", label:"En-tête projet", selector:"#viewProject .panel-head", wide:false },
      { key:"project_tasks", label:"Tableau des tâches", selector:"#viewProject .tablewrap.project-tasks", wide:true },
      { key:"project_gantt", label:"Gantt hebdo", selector:"#viewProject #gantt", wide:true },
      { key:"project_workload", label:"Charge de travail (projet)", selector:"#workloadChartProjectWrap", wide:true },
      { key:"project_pie", label:"Répartition Interne / Externe / RSG / RI (projet)", selector:"#workloadPieProjectWrap", wide:false },
      { key:"project_hours", label:"Analyse heures réelles (projet)", selector:"#projectHoursReportCard", wide:true },
      { key:"project_hours_internal_only", label:"Analyse heures réelles (projet) - sans prestataires externes", selector:"#projectHoursReportCard", wide:true }
    ];
  }
  return [
    { key:"master_table", label:"Tableau maître", selector:"#viewMaster .tablewrap", wide:true },
    { key:"master_gantt", label:"Gantt global", selector:"#viewMaster #masterGantt", wide:true },
    { key:"master_pie", label:"Répartition Interne / Externe / RSG / RI (tableau maître)", selector:"#workloadPieWrap", wide:false },
    { key:"master_hours", label:"Analyse heures réelles", selector:"#masterHoursReportCard", wide:true },
    { key:"master_hours_internal_only", label:"Analyse heures réelles - sans prestataires externes", selector:"#masterHoursReportCard", wide:true }
  ];
}

function cloneNodeForUnifiedExport(sourceNode){
  if(!sourceNode) return null;
  const clone = sourceNode.cloneNode(true);
  clone.querySelectorAll("button,.btn,.theme-picker,.icon-btn,[data-no-export]").forEach(n=>n.remove());
  clone.querySelectorAll("input,select,textarea").forEach(n=>{
    const span = document.createElement("span");
    const value = n.value || n.getAttribute("value") || "";
    span.textContent = value;
    span.className = "pdf-input-value";
    n.replaceWith(span);
  });
  const sourceCanvas = sourceNode.querySelectorAll("canvas");
  const cloneCanvas = clone.querySelectorAll("canvas");
  cloneCanvas.forEach((c, idx)=>{
    const src = sourceCanvas[idx];
    if(!src) return;
    try{
      const img = document.createElement("img");
      img.src = src.toDataURL("image/png", 1.0);
      img.style.width = src.width ? `${src.width}px` : "100%";
      img.style.maxWidth = "100%";
      img.style.height = "auto";
      c.replaceWith(img);
    }catch(_e){/* noop */}
  });
  return clone;
}

function buildProjectTasksExportHTML(projectId){
  const tasks = sortTasks(state.tasks.filter(t=>t.projectId===projectId), sortProject);
  if(!tasks.length){
    return "<div class='tablewrap project-tasks-export'><table class='table'><tbody><tr><td class='empty-row'>Aucune tâche</td></tr></tbody></table></div>";
  }

  const missingMap = buildMissingDaysMap(tasks);
  let rows = "";
  tasks.forEach(t=>{
    const statuses = parseStatuses(t.status).map(v=>v.toUpperCase());
    const c = ownerColor(t.owner);
    const ownerBadgeHtml = t.owner ? ownerBadgeForTask(t) : "";
    const durLabel = durationLabelForTask(t);
    const todayKey = new Date().toISOString().slice(0,10);
    const isToday = !!(t.start && t.end && t.start<=todayKey && t.end>=todayKey);
    const isLate = !!(t.end && t.end < todayKey);
    const isSelected = t.id===selectedTaskId;
    const rowClass = `${isSelected ? "row-selected " : ""}${isToday ? "today-row " : ""}${isLate ? "late-row" : ""}`.trim();
    const miss = missingMap.get(t.id) || 0;
    const missDot = miss>0 ? `<span class="missing-dot" title="Heures réelles manquantes (${miss} j)"></span>` : "";
    rows += `<tr class="${rowClass}">
      <td>${missDot}<span class="num-badge" style="--badge-color:${c};--badge-text:#fff;">${taskOrderMap[t.id]||""}</span></td>
      <td><span class="icon-picto"></span> ${taskTitleProjectView(t)}</td>
      <td class="status-cell"><span class="status-left">${statusDot(statuses[0])}${statusLabels(t.status||"")}</span>${ownerBadgeHtml||""}</td>
      <td>${formatDate(t.start)||""}</td>
      <td>${formatDate(t.end)||""}</td>
      <td>${taskProgress(t)}%</td>
      <td>${durLabel}</td>
    </tr>`;
  });

  return `<div class="tablewrap project-tasks-export"><table class="table" id="projectTasksExportTable">
    <thead>
      <tr>
        <th style="width:70px">N</th>
        <th>Description</th>
        <th style="width:320px">Statuts</th>
        <th style="width:120px">Début</th>
        <th style="width:120px">Fin</th>
        <th style="width:100px">Avancement</th>
        <th style="width:130px">Durée (jours)</th>
      </tr>
    </thead>
    <tbody>${rows}</tbody>
  </table></div>`;
}


function buildMasterTableExportHTML(){
  const sorted = sortTasks(filteredTasks(), sortMaster);
  if(!sorted.length){
    return "<div class='tablewrap master-table-export'><table class='table'><tbody><tr><td class='empty-row'>Aucune tâche.</td></tr></tbody></table></div>";
  }

  const missingMap = buildMissingDaysMap(sorted);
  const todayKey = new Date().toISOString().slice(0,10);
  let rows = "";

  sorted.forEach(t=>{
    const p = state.projects.find(x=>x.id===t.projectId);
    const statuses = parseStatuses(t.status).map(v=>v.toUpperCase());
    const c = ownerColor(t.owner);
    const rowBg = siteColor(p?.site);
    const sub = (p?.subproject || "").trim();
    const projLabel = sub ? `${p?.name||"Sans projet"} - ${sub}` : (p?.name||"Sans projet");
    const taskLabel = (t.roomNumber||"").trim();
    const isToday = !!(t.start && t.end && t.start<=todayKey && t.end>=todayKey);
    const isLate = !!(t.end && t.end < todayKey);
    const rowClass = `${isToday ? "today-row " : ""}${isLate ? "late-row" : ""}`.trim();
    const miss = missingMap.get(t.id) || 0;
    const missDot = miss>0 ? `<span class="missing-dot" title="Heures réelles manquantes (${miss} j)"></span>` : "";

    rows += `<tr class="${rowClass}" style="--site-bg:${rowBg};background:var(--site-bg);">
      <td>${p?.site||""}</td>
      <td>${projLabel}</td>
      <td>${missDot}<span class="num-badge" style="--badge-color:${c};--badge-text:#fff;">${taskOrderMap[t.id]||""}</span> <span class="icon-picto"></span> ${taskLabel}</td>
      <td class="status-cell"><span class="status-left">${statusDot(statuses[0])}${statusLabels(t.status||"")}</span>${t.owner?ownerBadgeForTask(t):""}</td>
      <td>${formatDate(t.start)||""}${isToday ? `<span class="today-dot" title="En cours aujourd'hui"></span>` : ""}</td>
      <td>${formatDate(t.end)||""}</td>
      <td>${taskProgress(t)}%</td>
      <td>${durationLabelForTask(t)}</td>
    </tr>`;
  });

  return `<div class="tablewrap master-table-export"><table class="table" id="masterTableExportTable">
    <thead>
      <tr>
        <th style="width:80px">Site</th>
        <th>Projet</th>
        <th>Tâche</th>
        <th style="width:300px">Statut</th>
        <th style="width:120px">Début</th>
        <th style="width:120px">Fin</th>
        <th style="width:100px">Avancement</th>
        <th style="width:130px">Durée (jours)</th>
      </tr>
    </thead>
    <tbody>${rows}</tbody>
  </table></div>`;
}



function buildMasterExportHeaderHTML(pdfTheme){
  const tasks = filteredTasks();
  const dated = tasks.filter(t=>t.start && t.end);
  const start = dated.length ? dated.reduce((min, t)=> (!min || t.start < min ? t.start : min), "") : "";
  const end = dated.length ? dated.reduce((max, t)=> (!max || t.end > max ? t.end : max), "") : "";
  const sites = Array.from(new Set(tasks.map(t=>{
    const p = state.projects.find(x=>x.id===t.projectId);
    return (p?.site || "").trim();
  }).filter(Boolean))).sort((a,b)=>a.localeCompare(b,"fr",{sensitivity:"base"}));
  const siteLabel = sites.length ? sites.join(", ") : "Tous";

  return `
    <div class="project-export-header-card" style="border:1px solid ${attrEscape(pdfTheme.line)};border-radius:8px;padding:6px 8px;background:${attrEscape(pdfTheme.panel)};background-image:linear-gradient(180deg, ${attrEscape(pdfTheme.panel)} 0%, ${attrEscape(pdfTheme.accentSoft)} 100%);margin-bottom:6px;box-shadow:0 3px 10px rgba(15,23,42,0.18);">
      <div class="project-export-header-grid" style="display:grid;grid-template-columns:1.6fr 1fr .9fr 1fr 1fr 1fr;gap:8px;align-items:start;border-left:4px solid ${attrEscape(pdfTheme.accent)};padding-left:7px;">
        <div><div style="font-size:12px;color:${attrEscape(pdfTheme.muted)};letter-spacing:.15px;margin-bottom:2px;font-weight:700;">Contexte</div><div style="font-weight:800;font-size:16px;line-height:1.2;color:${attrEscape(pdfTheme.text)};">Tableau maître</div></div>
        <div><div style="font-size:12px;color:${attrEscape(pdfTheme.muted)};letter-spacing:.15px;margin-bottom:2px;font-weight:700;">Site / Zone</div><div style="font-weight:800;font-size:16px;line-height:1.2;color:${attrEscape(pdfTheme.text)};">${attrEscape(siteLabel)}</div></div>
        <div><div style="font-size:12px;color:${attrEscape(pdfTheme.muted)};letter-spacing:.15px;margin-bottom:2px;font-weight:700;">Tâches</div><div style="font-weight:800;font-size:16px;line-height:1.2;color:${attrEscape(pdfTheme.text)};">${tasks.length}</div></div>
        <div><div style="font-size:12px;color:${attrEscape(pdfTheme.muted)};letter-spacing:.15px;margin-bottom:2px;font-weight:700;">Début période</div><div style="font-weight:800;font-size:15px;line-height:1.2;color:${attrEscape(pdfTheme.text)};">${attrEscape(start ? formatDate(start) : "-")}</div></div>
        <div><div style="font-size:12px;color:${attrEscape(pdfTheme.muted)};letter-spacing:.15px;margin-bottom:2px;font-weight:700;">Fin période</div><div style="font-weight:800;font-size:15px;line-height:1.2;color:${attrEscape(pdfTheme.text)};">${attrEscape(end ? formatDate(end) : "-")}</div></div>
        <div><div style="font-size:12px;color:${attrEscape(pdfTheme.muted)};letter-spacing:.15px;margin-bottom:2px;font-weight:700;">Date export</div><div style="font-weight:800;font-size:15px;line-height:1.2;color:${attrEscape(pdfTheme.text)};">${attrEscape(new Date().toLocaleDateString("fr-FR", { day:"2-digit", month:"2-digit", year:"numeric" }))}</div></div>
      </div>
    </div>
  `;
}

function buildGroupedProjectTasksExportHTML(projectIds){
  const idSet = new Set((projectIds || []).map(normId).filter(Boolean));
  const tasksRaw = (state.tasks || []).filter(t=> idSet.has(t.projectId));
  if(!tasksRaw.length){
    return "<div class='tablewrap project-tasks-export'><table class='table'><tbody><tr><td class='empty-row'>Aucune tâche</td></tr></tbody></table></div>";
  }
  const tasks = sortTasks(tasksRaw, sortMaster);
  const missingMap = buildMissingDaysMap(tasks);
  const todayKey = new Date().toISOString().slice(0,10);
  let rows = "";
  tasks.forEach((t)=>{
    const p = state.projects.find(x=>x.id===t.projectId);
    const statuses = parseStatuses(t.status).map(v=>v.toUpperCase());
    const c = ownerColor(t.owner);
    const ownerBadgeHtml = t.owner ? ownerBadgeForTask(t) : "";
    const durLabel = durationLabelForTask(t);
    const isToday = !!(t.start && t.end && t.start<=todayKey && t.end>=todayKey);
    const isLate = !!(t.end && t.end < todayKey);
    const rowClass = `${isToday ? "today-row " : ""}${isLate ? "late-row" : ""}`.trim();
    const miss = missingMap.get(t.id) || 0;
    const missDot = miss>0 ? `<span class="missing-dot" title="Heures réelles manquantes (${miss} j)"></span>` : "";
    rows += `<tr class="${rowClass}">
      <td>${attrEscape(p?.name || "Projet")}</td>
      <td>${missDot}<span class="num-badge" style="--badge-color:${c};--badge-text:#fff;">${taskOrderMap[t.id]||""}</span></td>
      <td><span class="icon-picto"></span> ${taskTitleProjectView(t)}</td>
      <td class="status-cell"><span class="status-left">${statusDot(statuses[0])}${statusLabels(t.status||"")}</span>${ownerBadgeHtml||""}</td>
      <td>${formatDate(t.start)||""}</td>
      <td>${formatDate(t.end)||""}</td>
      <td>${taskProgress(t)}%</td>
      <td>${durLabel}</td>
    </tr>`;
  });
  return `<div class="tablewrap project-tasks-export"><table class="table" id="projectTasksExportTable">
    <thead>
      <tr>
        <th style="width:220px">Chantier</th>
        <th style="width:70px">N</th>
        <th>Description</th>
        <th style="width:320px">Statuts</th>
        <th style="width:120px">Début</th>
        <th style="width:120px">Fin</th>
        <th style="width:100px">Avancement</th>
        <th style="width:130px">Durée (jours)</th>
      </tr>
    </thead>
    <tbody>${rows}</tbody>
  </table></div>`;
}

function buildProjectExportHeaderHTML(project, pdfTheme){
  if(!project) return "";
  const projectTasks = state.tasks.filter(t=>t.projectId===project.id);
  const projectTasksDated = projectTasks.filter(t=>t.start && t.end);
  const firstTaskStart = projectTasksDated.length
    ? projectTasksDated.reduce((min, t)=> (!min || t.start < min ? t.start : min), "")
    : "";
  const lastTaskEnd = projectTasksDated.length
    ? projectTasksDated.reduce((max, t)=> (!max || t.end > max ? t.end : max), "")
    : "";
  return `
    <div class="project-export-header-card" style="border:1px solid ${attrEscape(pdfTheme.line)};border-radius:8px;padding:6px 8px;background:${attrEscape(pdfTheme.panel)};background-image:linear-gradient(180deg, ${attrEscape(pdfTheme.panel)} 0%, ${attrEscape(pdfTheme.accentSoft)} 100%);margin-bottom:6px;box-shadow:0 3px 10px rgba(15,23,42,0.18);">
      <div class="project-export-header-grid" style="display:grid;grid-template-columns:1.6fr 1fr 1fr .75fr 1fr 1fr;gap:8px;align-items:start;border-left:4px solid ${attrEscape(pdfTheme.accent)};padding-left:7px;">
        <div><div style="font-size:12px;color:${attrEscape(pdfTheme.muted)};letter-spacing:.15px;margin-bottom:2px;font-weight:700;">Chantier</div><div style="font-weight:800;font-size:16px;line-height:1.2;color:${attrEscape(pdfTheme.text)};">${attrEscape(project?.name || "-")}</div></div>
        <div><div style="font-size:12px;color:${attrEscape(pdfTheme.muted)};letter-spacing:.15px;margin-bottom:2px;font-weight:700;">Site / Zone</div><div style="font-weight:800;font-size:16px;line-height:1.2;color:${attrEscape(pdfTheme.text)};">${attrEscape(project?.site || "-")}</div></div>
        <div><div style="font-size:12px;color:${attrEscape(pdfTheme.muted)};letter-spacing:.15px;margin-bottom:2px;font-weight:700;">Sous-projet</div><div style="font-weight:800;font-size:16px;line-height:1.2;color:${attrEscape(pdfTheme.text)};">${attrEscape(project?.subproject || "-")}</div></div>
        <div><div style="font-size:12px;color:${attrEscape(pdfTheme.muted)};letter-spacing:.15px;margin-bottom:2px;font-weight:700;">Tâches</div><div style="font-weight:800;font-size:16px;line-height:1.2;color:${attrEscape(pdfTheme.text)};">${projectTasks.length}</div></div>
        <div><div style="font-size:12px;color:${attrEscape(pdfTheme.muted)};letter-spacing:.15px;margin-bottom:2px;font-weight:700;">Début projet</div><div style="font-weight:800;font-size:15px;line-height:1.2;color:${attrEscape(pdfTheme.text)};">${attrEscape(firstTaskStart ? formatDate(firstTaskStart) : "-")}</div></div>
        <div><div style="font-size:12px;color:${attrEscape(pdfTheme.muted)};letter-spacing:.15px;margin-bottom:2px;font-weight:700;">Fin projet</div><div style="font-weight:800;font-size:15px;line-height:1.2;color:${attrEscape(pdfTheme.text)};">${attrEscape(lastTaskEnd ? formatDate(lastTaskEnd) : "-")}</div></div>
      </div>
    </div>
  `;
}

function buildGroupedProjectsExportHeaderHTML(projects, pdfTheme){
  const list = (projects || []).filter(Boolean);
  if(!list.length) return "";
  const ids = new Set(list.map(p=>p.id));
  const tasks = (state.tasks || []).filter(t=>ids.has(t.projectId));
  const dated = tasks.filter(t=>t.start && t.end);
  const start = dated.length ? dated.reduce((min, t)=> (!min || t.start < min ? t.start : min), "") : "";
  const end = dated.length ? dated.reduce((max, t)=> (!max || t.end > max ? t.end : max), "") : "";
  const sites = Array.from(new Set(list.map(p=>(p?.site || "").trim()).filter(Boolean))).sort((a,b)=>a.localeCompare(b, "fr", { sensitivity:"base" }));
  const names = list.map(p=>p?.name || "Chantier").join(", ");
  return `
    <div class="project-export-header-card" style="border:1px solid ${attrEscape(pdfTheme.line)};border-radius:8px;padding:6px 8px;background:${attrEscape(pdfTheme.panel)};background-image:linear-gradient(180deg, ${attrEscape(pdfTheme.panel)} 0%, ${attrEscape(pdfTheme.accentSoft)} 100%);margin-bottom:6px;box-shadow:0 3px 10px rgba(15,23,42,0.18);">
      <div class="project-export-header-grid" style="display:grid;grid-template-columns:1.5fr 1fr .9fr 1fr 1fr 1fr;gap:8px;align-items:start;border-left:4px solid ${attrEscape(pdfTheme.accent)};padding-left:7px;">
        <div><div style="font-size:12px;color:${attrEscape(pdfTheme.muted)};letter-spacing:.15px;margin-bottom:2px;font-weight:700;">Chantiers regroupés</div><div style="font-weight:800;font-size:15px;line-height:1.2;color:${attrEscape(pdfTheme.text)};">${attrEscape(names)}</div></div>
        <div><div style="font-size:12px;color:${attrEscape(pdfTheme.muted)};letter-spacing:.15px;margin-bottom:2px;font-weight:700;">Sites / Zones</div><div style="font-weight:800;font-size:15px;line-height:1.2;color:${attrEscape(pdfTheme.text)};">${attrEscape(sites.join(", ") || "-")}</div></div>
        <div><div style="font-size:12px;color:${attrEscape(pdfTheme.muted)};letter-spacing:.15px;margin-bottom:2px;font-weight:700;">Chantiers</div><div style="font-weight:800;font-size:16px;line-height:1.2;color:${attrEscape(pdfTheme.text)};">${list.length}</div></div>
        <div><div style="font-size:12px;color:${attrEscape(pdfTheme.muted)};letter-spacing:.15px;margin-bottom:2px;font-weight:700;">Tâches</div><div style="font-weight:800;font-size:16px;line-height:1.2;color:${attrEscape(pdfTheme.text)};">${tasks.length}</div></div>
        <div><div style="font-size:12px;color:${attrEscape(pdfTheme.muted)};letter-spacing:.15px;margin-bottom:2px;font-weight:700;">Début période</div><div style="font-weight:800;font-size:15px;line-height:1.2;color:${attrEscape(pdfTheme.text)};">${attrEscape(start ? formatDate(start) : "-")}</div></div>
        <div><div style="font-size:12px;color:${attrEscape(pdfTheme.muted)};letter-spacing:.15px;margin-bottom:2px;font-weight:700;">Fin période</div><div style="font-weight:800;font-size:15px;line-height:1.2;color:${attrEscape(pdfTheme.text)};">${attrEscape(end ? formatDate(end) : "-")}</div></div>
      </div>
    </div>
  `;
}

function renderUnifiedExportModulesList(){
  const list = el("exportPdfModulesList");
  if(!list) return;
  let defs = getUnifiedExportModuleDefinitions().filter(d=>document.querySelector(d.selector));
  const isProject = !el("viewProject")?.classList.contains("hidden");
  if(isProject){
    // L'en-tête projet est toujours injecté automatiquement sur chaque partie exportée.
    // "project_pie" embarque déjà les 2 graphiques, donc on masque l'entrée redondante.
    defs = defs.filter(d=>d.key !== "project_header" && d.key !== "project_workload");
  }

  const isProjectHoursVariant = (k)=>k === "project_hours" || k === "project_hours_internal_only";
  const isMasterHoursVariant = (k)=>k === "master_hours" || k === "master_hours_internal_only";
  const isHoursVariant = (k)=>isProjectHoursVariant(k) || isMasterHoursVariant(k);
  const hintHtml = defs.some(d=>isHoursVariant(d.key))
    ? `<div class="export-module-sub" style="margin:0 0 6px 2px;display:block;">Pour l'analyse des heures, choisissez une seule version.</div>`
    : "";

  const projectPickerHtml = (()=> {
    const projects = getProjectsSortedForExport();
    unifiedExportSelectedProjectIds = normalizeUnifiedExportProjectIds(unifiedExportSelectedProjectIds);
    const selected = new Set(unifiedExportSelectedProjectIds);
    const rows = projects.map((p)=>`
      <label class="export-module-row export-project-row">
        <input type="checkbox" data-export-project-id="${attrEscape(p.id)}" ${selected.has(p.id) ? "checked" : ""}>
        <span>
          <span class="export-module-label">${attrEscape(p.name || "Chantier")}</span>
          <span class="export-module-sub">${attrEscape((p.site || "-") + (p.subproject ? ` • ${p.subproject}` : ""))}</span>
        </span>
      </label>
    `).join("");
    return `
      <div class="export-projects-head">
        <span class="export-module-label">Chantiers à regrouper</span>
        <span class="export-projects-actions">
          <button type="button" class="btn btn-ghost btn-xs" id="btnExportProjectsActiveOnly">Actif uniquement</button>
          <button type="button" class="btn btn-ghost btn-xs" id="btnExportProjectsAll">Tous</button>
          <button type="button" class="btn btn-ghost btn-xs" id="btnExportProjectsNone">Aucun</button>
        </span>
      </div>
      ${rows}
    `;
  })();

  list.innerHTML = projectPickerHtml + hintHtml + defs.map((d)=>{
    const exclusive = isHoursVariant(d.key);
    const inputType = exclusive ? "radio" : "checkbox";
    const inputName = isProjectHoursVariant(d.key) ? "project_hours_variant" : (isMasterHoursVariant(d.key) ? "master_hours_variant" : "");
    const sub = exclusive
      ? ((d.key === "project_hours" || d.key === "master_hours") ? "Version complète" : "Version interne uniquement")
      : (d.wide ? "Format large" : "Format standard");
    return `
      <label class="export-module-row${exclusive ? " export-module-row-exclusive" : ""}">
        <input type="${inputType}" ${inputName ? `name="${inputName}"` : ""} data-export-module-key="${attrEscape(d.key)}">
        <span>
          <span class="export-module-label">${attrEscape(d.label)}</span>
          <span class="export-module-sub">${attrEscape(sub)}</span>
        </span>
      </label>
    `;
  }).join("");

  list.querySelectorAll("input[name=\"project_hours_variant\"],input[name=\"master_hours_variant\"]").forEach((n)=>{
    n.addEventListener("change", updateUnifiedHoursExclusiveUI);
  });

  const projectChecks = Array.from(list.querySelectorAll("input[data-export-project-id]"));
  const applyProjectSelectionToUi = ()=>{
    unifiedExportSelectedProjectIds = normalizeUnifiedExportProjectIds(unifiedExportSelectedProjectIds, { allowEmpty:true });
    const selected = new Set(unifiedExportSelectedProjectIds);
    projectChecks.forEach((n)=>{ n.checked = selected.has(normId(n.getAttribute("data-export-project-id"))); });
  };
  const syncProjectSelectionFromUi = ()=>{
    const ids = projectChecks.filter(n=>n.checked).map(n=>normId(n.getAttribute("data-export-project-id")));
    unifiedExportSelectedProjectIds = normalizeUnifiedExportProjectIds(ids, { allowEmpty:true });
    applyProjectSelectionToUi();
  };
  projectChecks.forEach((n)=> n.addEventListener("change", syncProjectSelectionFromUi));
  list.querySelector("#btnExportProjectsActiveOnly")?.addEventListener("click", ()=>{
    unifiedExportSelectedProjectIds = getActiveProjectIdsForToday();
    applyProjectSelectionToUi();
  });
  list.querySelector("#btnExportProjectsAll")?.addEventListener("click", ()=>{
    unifiedExportSelectedProjectIds = normalizeUnifiedExportProjectIds(getProjectsSortedForExport().map(p=>p.id));
    applyProjectSelectionToUi();
  });
  list.querySelector("#btnExportProjectsNone")?.addEventListener("click", ()=>{
    unifiedExportSelectedProjectIds = [];
    applyProjectSelectionToUi();
  });
  applyProjectSelectionToUi();
  updateUnifiedHoursExclusiveUI();
}
function updateUnifiedHoursExclusiveUI(){
  const list = el("exportPdfModulesList");
  if(!list) return;
  const groups = [
    [
      list.querySelector('input[data-export-module-key="project_hours"]'),
      list.querySelector('input[data-export-module-key="project_hours_internal_only"]')
    ],
    [
      list.querySelector('input[data-export-module-key="master_hours"]'),
      list.querySelector('input[data-export-module-key="master_hours_internal_only"]')
    ]
  ];

  groups.forEach(([rFull, rInternal])=>{
    if(!rFull || !rInternal) return;
    const rowFull = rFull.closest('.export-module-row');
    const rowInternal = rInternal.closest('.export-module-row');
    if(rowFull){
      rowFull.classList.toggle('export-module-row-locked', !!rInternal.checked);
      rowFull.setAttribute('title', '');
    }
    if(rowInternal){
      rowInternal.classList.toggle('export-module-row-locked', !!rFull.checked);
      rowInternal.setAttribute('title', '');
    }
  });
}
function openUnifiedPdfModal(){
  const modal = el("exportPdfModal");
  if(!modal) return;
  unifiedExportSelectedProjectIds = getUnifiedDefaultExportProjectIds();
  renderUnifiedExportModulesList();
  showModalSafely(modal);
}

function closeUnifiedPdfModal(){
  const modal = el("exportPdfModal");
  hideModalSafely(modal, "#btnExportPdfUnified");
}

const HOURS_VARIANT_KEYS = new Set([
  "project_hours",
  "project_hours_internal_only",
  "master_hours",
  "master_hours_internal_only"
]);

function getCheckedExportModuleKeys(listRoot){
  return Array.from(listRoot.querySelectorAll("input[data-export-module-key]:checked"))
    .map((n)=>n.getAttribute("data-export-module-key") || "");
}

function resolveHoursVariantsFromModalState(listRoot, selectedKeys){
  const projectHoursFullChecked = !!listRoot.querySelector('input[data-export-module-key="project_hours"]')?.checked;
  const projectHoursInternalChecked = !!listRoot.querySelector('input[data-export-module-key="project_hours_internal_only"]')?.checked;
  const masterHoursFullChecked = !!listRoot.querySelector('input[data-export-module-key="master_hours"]')?.checked;
  const masterHoursInternalChecked = !!listRoot.querySelector('input[data-export-module-key="master_hours_internal_only"]')?.checked;

  if(projectHoursFullChecked && projectHoursInternalChecked){
    throw new Error("Sélection incohérente: choisissez une seule variante pour l'analyse heures projet.");
  }
  if(masterHoursFullChecked && masterHoursInternalChecked){
    throw new Error("Sélection incohérente: choisissez une seule variante pour l'analyse heures tableau maître.");
  }

  const selectedProjectHoursVariant = projectHoursInternalChecked
    ? "project_hours_internal_only"
    : (projectHoursFullChecked ? "project_hours" : "");
  const selectedMasterHoursVariant = masterHoursInternalChecked
    ? "master_hours_internal_only"
    : (masterHoursFullChecked ? "master_hours" : "");

  const cleanKeys = selectedKeys.filter((k)=>!HOURS_VARIANT_KEYS.has(k));
  if(selectedProjectHoursVariant) cleanKeys.push(selectedProjectHoursVariant);
  if(selectedMasterHoursVariant) cleanKeys.push(selectedMasterHoursVariant);

  return {
    selectedProjectHoursVariant,
    selectedMasterHoursVariant,
    selectedKeys: Array.from(new Set(cleanKeys))
  };
}

function buildSelectedExportDefinitions(defs, selectedKeys, selectedProjectHoursVariant, selectedMasterHoursVariant){
  const selectedDefs = defs.filter((d)=>selectedKeys.includes(d.key) && !HOURS_VARIANT_KEYS.has(d.key));
  if(selectedProjectHoursVariant){
    const def = defs.find((d)=>d.key === selectedProjectHoursVariant);
    if(def) selectedDefs.push(def);
  }
  if(selectedMasterHoursVariant){
    const def = defs.find((d)=>d.key === selectedMasterHoursVariant);
    if(def) selectedDefs.push(def);
  }
  return selectedDefs;
}

function runUnifiedPdfExport(){
  try{
    const modal = el("exportPdfModal");
    const listRoot = modal?.querySelector("#exportPdfModulesList") || el("exportPdfModulesList") || document;
    const defs = getUnifiedExportModuleDefinitions();
    const rawSelectedKeys = getCheckedExportModuleKeys(listRoot);
    const {
      selectedProjectHoursVariant,
      selectedMasterHoursVariant,
      selectedKeys
    } = resolveHoursVariantsFromModalState(listRoot, rawSelectedKeys);

    if(selectedKeys.length===0){
      alert("Sélectionnez au moins un module.");
      return;
    }

    const isProjectMode = !el("viewProject")?.classList.contains("hidden");
    const selectedProjectIds = getUnifiedExportSelectedProjectIdsFromUi(listRoot);
    if(!selectedProjectIds.length){
      alert("Sélectionnez au moins un chantier.");
      return;
    }
    const selectedProjects = selectedProjectIds
      .map(id=> state.projects.find(p=>p.id===id))
      .filter(Boolean);
    const currentProject = isProjectMode ? (selectedProjects[0] || null) : null;

    const selectedDefs = buildSelectedExportDefinitions(
      defs,
      selectedKeys,
      selectedProjectHoursVariant,
      selectedMasterHoursVariant
    );
    const mergeProjectHeaderWithTasks = isProjectMode && selectedKeys.includes("project_tasks");
    const rootStyle = getComputedStyle(document.documentElement);
    const pdfTheme = {
      accent: (rootStyle.getPropertyValue("--accent") || "#2563eb").trim(),
      accentSoft: (rootStyle.getPropertyValue("--accent-soft") || "#93c5fd").trim(),
      panel: (rootStyle.getPropertyValue("--panel") || "#f8fafc").trim(),
      line: (rootStyle.getPropertyValue("--line") || "rgba(0,0,0,0.14)").trim(),
      text: (rootStyle.getPropertyValue("--text") || "#0f172a").trim(),
      muted: (rootStyle.getPropertyValue("--muted") || "#475569").trim()
    };

    const checkProjectGanttVsTasks = isProjectMode && selectedKeys.includes("project_tasks") && selectedKeys.includes("project_gantt");
    if(checkProjectGanttVsTasks){
      for(const project of selectedProjects){
        const projectTasksAll = state.tasks.filter(t=>t.projectId===project.id);
        const projectTasksDated = projectTasksAll.filter(t=>t.start && t.end);
        if(projectTasksDated.length !== projectTasksAll.length){
          const missingDates = projectTasksAll.length - projectTasksDated.length;
          alert(
            `Export bloqué : incohérence tâches / Gantt.\n\n` +
            `Chantier: ${project.name || "Projet"}\n` +
            `- Tableau des tâches : ${projectTasksAll.length} ligne(s)\n` +
            `- Gantt hebdo : ${projectTasksDated.length} ligne(s)\n` +
            `- Tâches sans dates : ${missingDates}\n\n` +
            `Action requise : renseignez les dates manquantes ou décochez "Gantt hebdo".`
          );
          return;
        }
      }
    }

    const checkMasterGanttVsTable = !isProjectMode && selectedKeys.includes("master_table") && selectedKeys.includes("master_gantt");
    if(checkMasterGanttVsTable){
      const allMasterTasks = state.tasks.filter((t)=>selectedProjectIds.includes(t.projectId));
      const datedMasterTasks = allMasterTasks.filter(t=>t.start && t.end);
      if(datedMasterTasks.length !== allMasterTasks.length){
        const missingDates = allMasterTasks.length - datedMasterTasks.length;
        alert(
          `Export bloqué : incohérence tâches / Gantt.

` +
          `- Tableau maître : ${allMasterTasks.length} ligne(s)
` +
          `- Gantt global : ${datedMasterTasks.length} ligne(s)
` +
          `- Tâches sans dates : ${missingDates}

` +
          `Action requise : renseignez les dates manquantes ou décochez "Gantt global".`
        );
        return;
      }
    }
    const masterHeaderHtml = buildMasterExportHeaderHTML(pdfTheme);

    const modules = [];
    selectedDefs.forEach(def=>{
      if(def.key === "project_workload" && selectedKeys.includes("project_pie")){
        return;
      }
      if(def.key === "project_header"){
        if(mergeProjectHeaderWithTasks) return;
        modules.push({
          key:def.key,
          label:"",
          wide:false,
          html:buildGroupedProjectsExportHeaderHTML(selectedProjects, pdfTheme),
          forceNewPage:true
        });
        return;
      }

      if(def.key === "project_tasks" && selectedProjects.length){
        const tableHtml = buildGroupedProjectTasksExportHTML(selectedProjectIds);
        const tasksTitle = `<div class="card-title">Tableau des tâches (cumul chantiers)</div>`;
        const groupedHeader = buildGroupedProjectsExportHeaderHTML(selectedProjects, pdfTheme);
        modules.push({
          key:def.key,
          label:"",
          wide:true,
          html:`${groupedHeader}${tasksTitle}${tableHtml}`,
          noAutoProjectHeader:true,
          forceNewPage:true
        });
        return;
      }


      if(def.key === "master_table"){
        const tableTitle = `<div class="card-title">Tableau maître (cumul chantiers)</div>`;
        const tableHtml = buildGroupedProjectTasksExportHTML(selectedProjectIds);
        const groupedHeader = buildGroupedProjectsExportHeaderHTML(selectedProjects, pdfTheme);
        modules.push({
          key:def.key,
          label:"",
          wide:true,
          html:`${groupedHeader}${tableTitle}${tableHtml}`,
          noAutoProjectHeader:true,
          forceNewPage:true
        });
        return;
      }

      if(def.key === "master_gantt"){
        const ganttTasks = state.tasks.filter((t)=>selectedProjectIds.includes(t.projectId) && t.start && t.end);
        const ganttHtml = buildProjectGanttHTMLForRange(null, null, ganttTasks, true);
        const ganttTitle = `<div class="card-title">Gantt global (cumul chantiers)</div>`;
        const groupedHeader = buildGroupedProjectsExportHeaderHTML(selectedProjects, pdfTheme);
        modules.push({
          key:def.key,
          label:"",
          wide:true,
          html:`${groupedHeader}${ganttTitle}${ganttHtml}`,
          noAutoProjectHeader:true,
          forceNewPage:true
        });
        return;
      }

      if(def.key === "master_pie"){
        const sectionHtml = buildGroupedProjectsRepartitionExportInnerHTML(
          selectedProjectIds,
          true,
          "Répartition Interne / Externe / RSG / RI (tableau maître - cumul chantiers)"
        );
        const groupedHeader = buildGroupedProjectsExportHeaderHTML(selectedProjects, pdfTheme);
        modules.push({
          key:def.key,
          label:"",
          wide:true,
          html:`${groupedHeader}${sectionHtml}`,
          noAutoProjectHeader:true,
          forceNewPage:true
        });
        return;
      }

      if(def.key === "master_hours"){
        if(selectedMasterHoursVariant !== "master_hours") return;
        const groupedHeader = buildGroupedProjectsExportHeaderHTML(selectedProjects, pdfTheme);
        modules.push({
          key:def.key,
          label:"",
          wide:true,
          html:`${groupedHeader}${buildGroupedProjectsRealHoursReportInnerHTML(selectedProjectIds, true, "Analyse heures réelles (tableau maître - cumul chantiers)")}`,
          noAutoProjectHeader:true,
          forceNewPage:true
        });
        return;
      }

      if(def.key === "master_hours_internal_only"){
        if(selectedMasterHoursVariant !== "master_hours_internal_only") return;
        const groupedHeader = buildGroupedProjectsExportHeaderHTML(selectedProjects, pdfTheme);
        modules.push({
          key:def.key,
          label:"",
          wide:true,
          html:`${groupedHeader}${buildGroupedProjectsRealHoursReportInnerHTML(selectedProjectIds, false, "Analyse heures réelles (tableau maître - cumul chantiers) - sans prestataires externes")}`,
          noAutoProjectHeader:true,
          forceNewPage:true
        });
        return;
      }

      if(def.key === "project_gantt" && selectedProjects.length){
        const projectTasksForGantt = state.tasks.filter(t=> selectedProjectIds.includes(t.projectId) && t.start && t.end);
        const ganttHtml = buildProjectGanttHTMLForRange(null, null, projectTasksForGantt, true);
        const ganttTitle = `<div class="card-title">Gantt hebdo (cumul chantiers)</div>`;
        const groupedHeader = buildGroupedProjectsExportHeaderHTML(selectedProjects, pdfTheme);
        modules.push({
          key:def.key,
          label:"",
          wide:true,
          html:`${groupedHeader}${ganttTitle}${ganttHtml}`,
          noAutoProjectHeader:true,
          forceNewPage:true
        });
        return;
      }

      if(def.key === "project_pie" && selectedProjects.length){
        const sectionHtml = buildGroupedProjectsRepartitionExportInnerHTML(
          selectedProjectIds,
          true,
          "Répartition Interne / Externe / RSG / RI (cumul chantiers)"
        );
        const groupedHeader = buildGroupedProjectsExportHeaderHTML(selectedProjects, pdfTheme);
        modules.push({
          key:def.key,
          label:"",
          wide:true,
          html:`${groupedHeader}${sectionHtml}`,
          noAutoProjectHeader:true,
          forceNewPage:true
        });
        return;
      }


      if((def.key === "project_hours" || def.key === "project_hours_internal_only") && selectedProjects.length){
        if(def.key !== selectedProjectHoursVariant) return;
        const includeExternal = def.key !== "project_hours_internal_only";
        const reportTitle = includeExternal
          ? "Analyse heures réelles (cumul chantiers)"
          : "Analyse heures réelles (cumul chantiers) - sans prestataires externes";
        const hoursHtml = buildGroupedProjectsRealHoursReportInnerHTML(selectedProjectIds, includeExternal, reportTitle);
        const groupedHeader = buildGroupedProjectsExportHeaderHTML(selectedProjects, pdfTheme);
        modules.push({
          key:def.key,
          label:"",
          wide:true,
          html:`${groupedHeader}${hoursHtml}`,
          noAutoProjectHeader:true,
          forceNewPage:true
        });
        return;
      }

      const rawNode = document.querySelector(def.selector);
      if(!rawNode) return;
      const node = rawNode.closest(".card") || rawNode;
      const cloned = cloneNodeForUnifiedExport(node);
      if(!cloned) return;
      modules.push({
        key: def.key,
        label: def.label,
        wide: !!def.wide,
        html: cloned.outerHTML
      });
    });

    if(modules.length===0){
      alert("Aucun module exportable trouvé.");
      return;
    }

    const viewer = window.open("about:blank", "_blank");
    if(!viewer){
      alert("Popup bloquée. Autorisez les popups puis réessayez.");
      return;
    }

    setPrintPageFormat("A4 landscape", "2mm");
    document.body.classList.add("print-mode");

    const tpl = ensurePrintTemplate();
    if(!tpl){
      alert("Template d'export introuvable.");
      return;
    }

    let container = document.getElementById("printInjection");
    if(!container){
      container = document.createElement("div");
      container.id = "printInjection";
      document.body.prepend(container);
    }

    container.innerHTML = tpl.innerHTML;

    const header = container.querySelector("#printHeader");
    const meta = container.querySelector("#printMeta");
    const legend = container.querySelector("#printLegend");
    const filters = getExportFiltersPayload();
    const filtersLabel = Object.entries(filters)
      .filter(([,v])=>String(v || "").trim())
      .map(([k,v])=>`${k}: ${v}`)
      .join(" • ") || "Aucun filtre";

    if(header){
      header.querySelector("h1").textContent = (el("brandTitle")?.textContent || "Suivi de Chantiers").trim();
    }
    if(meta){
      const mode = isProjectMode
        ? (selectedProjects.length > 1 ? `${selectedProjects.length} chantiers regroupés` : ((el("projectTitle")?.textContent || currentProject?.name || "Projet").trim()))
        : "Tableau maître";
      const metaRows = [
        ["Contexte", mode],
        ["Date export", new Date().toLocaleDateString("fr-FR", { day:"2-digit", month:"short", year:"numeric" })],
        ["Filtres", filtersLabel],
        ["Modules", modules.length]
      ];
      meta.innerHTML = metaRows.map(([k,v])=>`<div><strong>${k}</strong><br>${attrEscape(String(v ?? ""))}</div>`).join("");
    }
    if(legend) legend.innerHTML = "";

    container.querySelectorAll(".print-dynamic").forEach(n=>n.remove());
    const wrap = document.createElement("div");
    wrap.className = "print-dynamic";
    modules.forEach((m)=>{
      const block = document.createElement("div");
      block.className = `card print-block${m.forceNewPage ? " force-new-page" : ""}`;
      const titleHtml = m.label ? `<div class="card-title">${attrEscape(m.label)}</div>` : "";
      block.innerHTML = `${titleHtml}${m.html || ""}`;
      wrap.appendChild(block);
    });
    container.querySelector(".print-order")?.appendChild(wrap);

    maximizePrintContainer(container);
    openPreparedPrintInNewWindow("Export PDF", viewer);

    closeUnifiedPdfModal();
  }catch(err){
    reportAppError(err, "export.pdf.modal");
    alert("Erreur export PDF: " + (err?.message || err));
  }
}
window.runUnifiedPdfExport = runUnifiedPdfExport;

function bind(){

  loadStatusConfig();
  applyVacationConfig();
  buildStatusMenu();

  setStatusSelection("");

  el("t_status_display")?.addEventListener("click",(e)=>{ e.stopPropagation(); toggleStatusMenu(true); });

  document.addEventListener("click",(e)=>{

    const wrap = el("t_status_wrap");

    if(wrap && !wrap.contains(e.target)){ toggleStatusMenu(false); }

  });



  updateTopbarHeight();
  updateSidebarTop();
  applySidebarTopLock();
  applyRoleAccess();
  initThemePicker();
  const performLogout = ()=>{
    try{
      sessionStorage.removeItem("unlocked");
      sessionStorage.removeItem("current_user");
      sessionStorage.removeItem("current_role");
      sessionStorage.removeItem("current_email");
      sessionStorage.removeItem("current_theme");
      localStorage.removeItem("login_session_token_v1");
    }catch(e){ softCatch(e); }
    const lock = document.getElementById("lockscreen");
    if(lock) lock.classList.remove("hidden");
    applyRoleAccess();
    try{ window.refreshLoginUsers?.(); }catch(e){ softCatch(e); }
  };
  const switchBtn = el("btnSwitchUser");
  if(switchBtn){
    switchBtn.addEventListener("click", performLogout);
  }
  const logoutBtn = el("btnLogout");
  if(logoutBtn){
    logoutBtn.addEventListener("click", performLogout);
  }
  window.addEventListener("resize", ()=>{
    updateTopbarHeight();
    updateSidebarTop();
    applySidebarTopLock();
  });
  el("btnConfig")?.addEventListener("click", ()=>{
    if(getCurrentRole()!=="admin") return;
    openConfigModal();
  });
  el("btnHelp")?.addEventListener("click", ()=>{
    const modal = el("helpModal");
    if(!modal) return;
    showModalSafely(modal);
  });
  el("btnHelpClose")?.addEventListener("click", ()=>{
    const modal = el("helpModal");
    if(!modal) return;
    hideModalSafely(modal, "#btnHelp");
  });
  el("helpModal")?.addEventListener("click",(e)=>{
    if(e.target && e.target.id==="helpModal"){
      const modal = el("helpModal");
      hideModalSafely(modal, "#btnHelp");
    }
  });
  el("btnStatusAdd")?.addEventListener("click", ()=>{
    if(getCurrentRole()!=="admin") return;
    const input = el("cfg_status_label");
    const label = (input?.value || "").trim();
    if(!label) return;
    const v = normalizeStatusId(label).toUpperCase();
    if(!v) return;
    if(STATUSES.some(s=>s.v===v)){
      alert("Statut déjà existant.");
      return;
    }
    STATUSES.push({v, label});
    saveStatusConfig();
    if(input) input.value = "";
    renderConfigStatusList();
    refreshStatusUi();
  });
  el("btnVendorAdd")?.addEventListener("click", ()=>{
    if(getCurrentRole()!=="admin") return;
    const input = el("cfg_vendor_name");
    const name = (input?.value || "").trim();
    if(!name) return;
    const registry = loadVendorsRegistry();
    if(registry.some(v=>v.toLowerCase()===name.toLowerCase())){
      alert("Prestataire déjà existant.");
      return;
    }
    const next = dedupVendors([...registry, name]).sort((a,b)=>a.localeCompare(b,"fr",{sensitivity:"base"}));
    saveVendorsRegistry(next);
    if(input) input.value = "";
    refreshVendorsList();
    renderConfigVendorsList();
    renderAll();
  });
  el("cfg_vac_year")?.addEventListener("change", ()=>{
    initVacationConfigUI();
  });
  el("btnVacSave")?.addEventListener("click", ()=>{
    if(getCurrentRole()!=="admin") return;
    const year = (el("cfg_vac_year")?.value || "").trim();
    if(!year) return;
    const schoolWeeks = normalizeWeekList(el("cfg_vac_school")?.value || "");
    const internalWeeks = normalizeWeekList(el("cfg_vac_internal")?.value || "");
    const cfg = loadConfig();
    const school = normalizeVacationMap(cfg.vacances_school || {});
    const internal = normalizeVacationMap(cfg.vacances_internal || {});
    if(schoolWeeks.length) school[year] = schoolWeeks; else delete school[year];
    if(internalWeeks.length) internal[year] = internalWeeks; else delete internal[year];
    cfg.vacances_school = school;
    cfg.vacances_internal = internal;
    saveConfig(cfg);
    applyVacationConfig();
    initVacationConfigUI();
    renderAll();
  });
  el("btnVacDelete")?.addEventListener("click", ()=>{
    if(getCurrentRole()!=="admin") return;
    const year = (el("cfg_vac_year")?.value || "").trim();
    if(!year) return;
    if(!confirm(`Supprimer les semaines pour ${year} ?`)) return;
    const cfg = loadConfig();
    const school = normalizeVacationMap(cfg.vacances_school || {});
    const internal = normalizeVacationMap(cfg.vacances_internal || {});
    delete school[year];
    delete internal[year];
    cfg.vacances_school = school;
    cfg.vacances_internal = internal;
    saveConfig(cfg);
    applyVacationConfig();
    initVacationConfigUI();
    renderAll();
  });
  el("btnConfigSave")?.addEventListener("click", ()=>{
    if(getCurrentRole()!=="admin") return;
    const prev = loadConfig();
    const cfg = {
      name: el("cfg_name")?.value || "",
      http: el("cfg_http")?.value || "",
      front: el("cfg_front")?.value || "",
      back: el("cfg_back")?.value || "",
      statuses: prev.statuses || STATUSES,
      vacances_school: prev.vacances_school || {},
      vacances_internal: prev.vacances_internal || {}
    };
    saveConfig(cfg);
    applyVacationConfig();
    closeConfigModal();
  });
  el("btnConfigSaveTop")?.addEventListener("click", ()=>{
    el("btnConfigSave")?.click();
  });
  el("btnQualityClean")?.addEventListener("click", ()=>{
    if(getCurrentRole()!=="admin") return;
    const report = collectDataQualityIssues(state);
    if(report.ok){
      showSaveToast("ok", "Nettoyage", "Aucune incohérence à corriger");
      return;
    }
    if(!confirm(`Nettoyer les incohérences détectées (${report.issues.length}) ?`)) return;
    applyDataQualityCleanup();
  });

  el("btnQualityExportPdf")?.addEventListener("click", ()=>{
    if(getCurrentRole()!=="admin") return;
    exportDataQualityReportPdf();
  });
  el("btnTabsSortProgress")?.addEventListener("click", ()=>{
    tabsSortMode = (tabsSortMode === "progress_asc") ? "progress_desc" : "progress_asc";
    renderTabs();
  });
  el("btnTabsSortReset")?.addEventListener("click", ()=>{
    tabsSortMode = "progress_asc";
    renderTabs();
  });
  el("btnConfigCloseTop")?.addEventListener("click", ()=>{
    el("btnConfigClose")?.click();
  });
  ["cfg_http","cfg_front","cfg_back"].forEach((id)=>{
    const input = el(id);
    if(!input) return;
    input.addEventListener("input", ()=>{
      const map = {cfg_http:"cfg_http_link", cfg_front:"cfg_front_link", cfg_back:"cfg_back_link"};
      const linkId = map[id];
      const link = el(linkId);
      const val = (input.value || "").trim();
      if(!link) return;
      if(val){
        link.href = val;
        link.textContent = val;
        link.classList.remove("hidden");
      }else{
        link.removeAttribute("href");
        link.textContent = "Ouvrir";
        link.classList.add("hidden");
      }
    });
  });
  el("btnUserAdd")?.addEventListener("click", async ()=>{
    if(getCurrentRole()!=="admin") return;
    const name = (el("cfg_user_name")?.value || "").trim();
    const email = (el("cfg_user_email")?.value || "").trim();
    const role = el("cfg_user_role")?.value || "user";
    const pass = el("cfg_user_pass")?.value || "";
    if(!name || !pass){ alert("Nom et mot de passe requis."); return; }
    const users = loadUsers();
    if(email && users.some(u=>(u.email||"").toLowerCase()===email.toLowerCase())){
      alert("Email déjà existant."); return;
    }
    const hash = await hashPassword(pass);
    users.push({id: uid(), name, email, role, hash, theme:"sable"});
    saveUsers(users);
    el("cfg_user_name").value = "";
    el("cfg_user_email").value = "";
    el("cfg_user_pass").value = "";
    renderUsersList();
  });
  el("btnConfigClose")?.addEventListener("click", ()=> closeConfigModal());
  el("configModal")?.addEventListener("click",(e)=>{
    if(e.target && e.target.id==="configModal") closeConfigModal();
  });
  el("cfg_login_start")?.addEventListener("change", ()=>{
    loginRangeStart = el("cfg_login_start")?.value || "";
    initLoginJournalUI();
  });
  el("cfg_login_end")?.addEventListener("change", ()=>{
    loginRangeEnd = el("cfg_login_end")?.value || "";
    initLoginJournalUI();
  });
  document.querySelectorAll(".login-log-sort").forEach(btn=>{
    btn.addEventListener("click", ()=>{
      const key = btn.dataset.sort || "ts";
      if(loginLogSortKey === key){
        loginLogSortDir = (loginLogSortDir === "asc") ? "desc" : "asc";
      }else{
        loginLogSortKey = key;
        loginLogSortDir = "asc";
      }
      initLoginJournalUI();
    });
  });
  el("cfg_login_reset")?.addEventListener("click", (e)=>{
    e.preventDefault();
    const end = new Date();
    const start = new Date();
    start.setDate(start.getDate()-29);
    loginRangeStart = start.toISOString().slice(0,10);
    loginRangeEnd = end.toISOString().slice(0,10);
    const startInput = el("cfg_login_start");
    const endInput = el("cfg_login_end");
    if(startInput) startInput.value = loginRangeStart;
    if(endInput) endInput.value = loginRangeEnd;
    initLoginJournalUI();
  });
  el("btnSave")?.addEventListener("click", async ()=>{
    const quality = collectDataQualityIssues(state);
    if(!quality.ok){
      showSaveToast("error", "Sauvegarde bloquée", formatQualityIssuesForToast(quality));
      updateDataQualityBanner(false);
      markDirty();
      return;
    }

    _suppressSupabaseSave = true;
    saveState({skipSupabase:true});
    _suppressSupabaseSave = false;

    let supabaseOk = false;
    let usersOk = false;
    try{ if(window.saveAppStateToSupabase) supabaseOk = await window.saveAppStateToSupabase(state); }catch(e){ softCatch(e); }
    try{ usersOk = await saveUsersToSupabase(loadUsers()); }catch(e){ softCatch(e); }

    // Backup local JSON volontairement désactivé: sauvegarde via Supabase uniquement.
    const backupEnabled = false;

    const detailParts = [];
    detailParts.push(`Supabase: ${supabaseOk ? "OK" : "ERREUR"}`);
    if(usersOk === false) detailParts.push(`Users: ERREUR`);
    if(getCurrentRole() === "admin"){
      detailParts.push("Backup local: désactivé");
    }
    showSaveToast(supabaseOk ? "ok" : "error", "Sauvegarde terminée", detailParts.join(" | "));

    flashSaved();
    renderAll();
    el("btnNewTask")?.classList.remove("btn-armed");
  });

  el("btnExportPdfUnified")?.addEventListener("click", ()=>{
    openUnifiedPdfModal();
  });
  el("btnExportPdfCancel")?.addEventListener("click", ()=>{
    closeUnifiedPdfModal();
  });

  
  const hmModalRoot = el("hoursTaskModal");
  if(!(hmModalRoot && hmModalRoot.dataset.hmListenersBound === "1")){
    if(hmModalRoot) hmModalRoot.dataset.hmListenersBound = "1";
    el("btnOpenHoursModal")?.addEventListener("click", ()=> openHoursTaskModal());
    el("btnCloseHoursModal")?.addEventListener("click", ()=> closeHoursTaskModal());
    el("btnSaveHoursModal")?.addEventListener("click", ()=> saveHoursTaskModal());
    el("hoursTaskModal")?.addEventListener("click", (e)=>{
    if(e.target && e.target.id === "hoursTaskModal") closeHoursTaskModal();
    });
    el("hm_date")?.addEventListener("change", ()=>{
    const t = getSelectedTaskForHoursModal();
    const dateInput = el("t_time_date_input");
    const hmDate = el("hm_date");
    const hmHours = el("hm_hours");
    if(!t || !dateInput || !hmDate) return;
    dateInput.value = hmDate.value || "";
    if(hmHours) hmHours.value = getHoursDraftForDate(hmDate.value || "");
    updateTimeLogUI(t, true);
    syncHoursTaskStatusFromCalendarDraft(t, hmDate.value || "", hmHours?.value || "");
    renderHoursTaskWeeklySummary(t, collectHoursTaskCalendarEntries(t));
    });
    el("hm_date")?.addEventListener("input", ()=>{
    const t = getSelectedTaskForHoursModal();
    const dateInput = el("t_time_date_input");
    const hmDate = el("hm_date");
    const hmHours = el("hm_hours");
    if(!t || !dateInput || !hmDate) return;
    dateInput.value = hmDate.value || "";
    if(hmHours) hmHours.value = getHoursDraftForDate(hmDate.value || "");
    updateTimeLogUI(t, true);
    syncHoursTaskStatusFromCalendarDraft(t, hmDate.value || "", hmHours?.value || "");
    renderHoursTaskWeeklySummary(t, collectHoursTaskCalendarEntries(t));
    });

    el("hoursTaskModal")?.addEventListener("click", (e)=>{
    if(!e.target?.closest?.("#hm_calendar")) return;
    if(e.target?.closest?.(".hm-day-input")) return;
    const btn = e.target?.closest?.("#hm_calendar .hm-day[data-date][data-active='1']");
    if(!btn) return;
    const t = getSelectedTaskForHoursModal();
    const hmDate = el("hm_date");
    const dateInput = el("t_time_date_input");
    if(!t || !hmDate || !dateInput) return;
    const day = btn.getAttribute("data-date") || "";
    hmDate.value = day;
    dateInput.value = day;
    const hmHours = el("hm_hours");
    updateTimeLogUI(t, true);
    const dayInput = btn.querySelector(".hm-day-input[data-date]");
    if(hmHours) hmHours.value = dayInput ? (dayInput.value || "") : "";
    syncHoursTaskStatusFromCalendarDraft(t, day, dayInput ? dayInput.value : "");
    renderHoursTaskWeeklySummary(t, collectHoursTaskCalendarEntries(t));
    });
    el("hoursTaskModal")?.addEventListener("input", (e)=>{
    if(!e.target?.closest?.("#hm_calendar")) return;
    const input = e.target?.closest?.("#hm_calendar .hm-day-input[data-date]");
    if(!input) return;
    const day = input.getAttribute("data-date") || "";
    const hmDate = el("hm_date");
    const hmHours = el("hm_hours");
    const dateInput = el("t_time_date_input");
    if(hmDate) hmDate.value = day;
    if(dateInput) dateInput.value = day;
    if(hmHours) hmHours.value = input.value || "";
    refreshHoursDayCardVisual(input);
    const t = getSelectedTaskForHoursModal();
    if(t){
      syncHoursTaskStatusFromCalendarDraft(t, day, input.value || "");
      renderHoursTaskWeeklySummary(t, collectHoursTaskCalendarEntries(t));
    }
    });
    el("hoursTaskModal")?.addEventListener("focusin", (e)=>{
    if(!e.target?.closest?.("#hm_calendar")) return;
    const input = e.target?.closest?.("#hm_calendar .hm-day-input[data-date][data-active='1']");
    if(!input) return;
    const t = getSelectedTaskForHoursModal();
    const day = input.getAttribute("data-date") || "";
    const hmDate = el("hm_date");
    const dateInput = el("t_time_date_input");
    if(!t || !day) return;
    if(hmDate) hmDate.value = day;
    if(dateInput) dateInput.value = day;
    const hmHours = el("hm_hours");
    if(hmHours) hmHours.value = input.value || "";
    updateTimeLogUI(t, true);
    syncHoursTaskStatusFromCalendarDraft(t, day, input.value || "");
    renderHoursTaskWeeklySummary(t, collectHoursTaskCalendarEntries(t));
    });
    el("hoursTaskModal")?.addEventListener("keydown", (e)=>{
    if(e.key !== "Enter" || e.isComposing) return;
    const input = e.target?.closest?.("#hm_calendar .hm-day-input[data-date][data-active='1']");
    if(!input) return;
    e.preventDefault();
    const day = input.getAttribute("data-date") || "";
    const hmDate = el("hm_date");
    const hmHours = el("hm_hours");
    const dateInput = el("t_time_date_input");
    if(hmDate) hmDate.value = day;
    if(dateInput) dateInput.value = day;
    if(hmHours) hmHours.value = input.value || "";
    const t = getSelectedTaskForHoursModal();
    if(t){
      syncHoursTaskStatusFromCalendarDraft(t, day, input.value || "");
      renderHoursTaskWeeklySummary(t, collectHoursTaskCalendarEntries(t));
    }
    saveHoursTaskModal();
    });
  }
  // bouton impression PDF (utilise print.css)

  el("btnBack")?.addEventListener("click", ()=>{

    selectedProjectId=null; selectedTaskId=null;

    renderAll();

  });
  el("btnToggleMasterVendor")?.addEventListener("click", ()=>{

    ganttColVisibility.masterVendor = !ganttColVisibility.masterVendor;

    applyGanttColumnVisibility();

  });

  el("btnToggleMasterStatus")?.addEventListener("click", ()=>{

    ganttColVisibility.masterStatus = !ganttColVisibility.masterStatus;

    applyGanttColumnVisibility();

  });

  el("btnToggleProjectVendor")?.addEventListener("click", ()=>{

    ganttColVisibility.projectVendor = !ganttColVisibility.projectVendor;

    applyGanttColumnVisibility();

  });

  el("btnToggleProjectStatus")?.addEventListener("click", ()=>{

    ganttColVisibility.projectStatus = !ganttColVisibility.projectStatus;

    applyGanttColumnVisibility();

  });
  el("btnAddProject")?.addEventListener("click", ()=>{

    if(isLocked) return;

    saveUndoSnapshot();

    const id = uid();

    const name = "Nouveau projet";

    state.projects.push({id,name,site:"",constraints:"",subproject:""});

    selectedProjectId = id;

    selectedTaskId = null;

    markDirty();

    renderAll();

  });

  el("btnSaveProject")?.addEventListener("click", ()=>{

    if(isLocked) return;

    if(!selectedProjectId) return;

    try{
      toggleStatusMenu(false);
      document.querySelectorAll(".vendor-dropdown,.desc-dropdown").forEach(n=>n.classList.remove("open"));
      document.activeElement && document.activeElement.blur && document.activeElement.blur();
    }catch(e){ softCatch(e); }

    saveUndoSnapshot();

    const p = state.projects.find(x=>x.id===selectedProjectId);

    if(!p) return;

    p.name        = el("p_name").value.trim();

    p.subproject  = el("p_subproject").value.trim();

    const siteSelect = el("p_site");
    p.site        = siteSelect ? siteSelect.value.trim() : "";
    const constraintsInput = el("p_constraints");
    p.constraints = constraintsInput ? constraintsInput.value.trim() : (p.constraints||"");

    renderTabs();

    markDirty();

    renderProject();

  });

  el("btnDeleteProject")?.addEventListener("click", ()=>{

    if(isLocked) return;

    if(!selectedProjectId) return;

    if(!confirm("Supprimer ce projet et toutes ses tâches ? Cette action est définitive.")) return;

    saveUndoSnapshot();

    state.projects = state.projects.filter(p=>p.id!==selectedProjectId);

    state.tasks    = state.tasks.filter(t=>t.projectId!==selectedProjectId);

    selectedProjectId=null;

    selectedTaskId=null;

    markDirty();

    renderAll();

  });

  el("btnAddTask")?.addEventListener("click", ()=>{

    if(isLocked) return;

    if(!selectedProjectId) return;

    saveUndoSnapshot();

    const id=uid();

    state.tasks.push({id,projectId:selectedProjectId,roomNumber:"",status:"",owner:"",vendor:"",start:"",end:"",notes:""});

    selectedTaskId=id;

    markDirty();

    renderProject();

  });

  el("btnNewTask")?.addEventListener("click", ()=>{

    if(isLocked) return;

    // Nouvelle tâche : vider tous les champs, pré-remplir uniquement la date de début (aujourd'hui)

    selectedTaskId=null;

    el("btnNewTask")?.classList.add("btn-armed");

    const todayVal = toInputDate(new Date());
    const tRoom = el("t_room"); if(tRoom) tRoom.value = "";
    const tOwner = el("t_owner"); if(tOwner) tOwner.value = "";
    const tVendor = el("t_vendor"); if(tVendor) tVendor.value = "";
    const tStart = el("t_start"); if(tStart) tStart.value = todayVal;
    const tEnd = el("t_end"); if(tEnd) tEnd.value = "";
    if(window.__fpStart){ try{ window.__fpStart.setDate(todayVal || null, true, "Y-m-d"); }catch(e){ softCatch(e); } }
    if(window.__fpEnd){ try{ window.__fpEnd.setDate(null, true, "Y-m-d"); }catch(e){ softCatch(e); } }

    setStatusSelection("");

  });

  el("btnDuplicateTask")?.addEventListener("click", ()=>{

    if(isLocked) return;

    if(!selectedProjectId || !selectedTaskId) return;

    const source = state.tasks.find(x=>x.id===selectedTaskId && x.projectId===selectedProjectId);

    if(!source) return;

    saveUndoSnapshot();

    const id = uid();

    const clone = {
      ...source,
      id,
      projectId: selectedProjectId
    };

    state.tasks.push(clone);

    selectedTaskId = id;

    markDirty();

    renderProject();

    el("btnNewTask")?.classList.remove("btn-armed");

  });

  el("btnDeleteTask")?.addEventListener("click", ()=>{

    if(isLocked) return;

    if(!selectedProjectId || !selectedTaskId) return;

    if(!confirm("Supprimer cette tâche ?")) return;

    saveUndoSnapshot();
    const taskIdSupprimee = selectedTaskId;

    state.tasks = state.tasks.filter(t=> !(t.id===selectedTaskId && t.projectId===selectedProjectId));
    state.timeLogs = (state.timeLogs || []).filter(log => log.taskId !== taskIdSupprimee);

    selectedTaskId = null;

    markDirty();

    renderProject();

    el("btnNewTask")?.classList.remove("btn-armed");

  });

  el("btnSaveTask")?.addEventListener("click", ()=>{

    if(isLocked) return;

    if(!selectedProjectId) return;

    try{
      toggleStatusMenu(false);
      document.querySelectorAll(".vendor-dropdown,.desc-dropdown").forEach(n=>n.classList.remove("open"));
      document.activeElement && document.activeElement.blur && document.activeElement.blur();
    }catch(e){ softCatch(e); }

    saveUndoSnapshot();
    const end = unformatDate(el("t_end").value);
    if (!end || !/^\d{4}-\d{2}-\d{2}$/.test(end)) {
      alert("La date de fin est obligatoire.");
      return;
    }
    const start = unformatDate(el("t_start").value);
    if (new Date(end) < new Date(start)) {
      alert("La date de fin ne peut pas être antérieure à la date de début.");
      return;
    }

    let t = state.tasks.find(x=>x.id===selectedTaskId && x.projectId===selectedProjectId);

    if(!t){

      const id=uid();

      t={id,projectId:selectedProjectId}; state.tasks.push(t); selectedTaskId=id;

    }

    t.roomNumber = el("t_room").value.trim();

    t.owner      = String(el("t_owner").value || "").toUpperCase();

    t.vendor     = el("t_vendor").value.trim();
    const taskOwnerType = ownerType(t.owner);
    if(taskOwnerType === "externe" && !t.vendor){
      alert("Prestataire externe requis : renseignez le nom du prestataire.");
      return;
    }
    if(taskOwnerType !== "externe"){
      t.vendor = "";
    }

    t.start      = unformatDate(el("t_start").value);
    t.end        = end;

    if(t.end && t.start && t.end < t.start){

      t.end = t.start;

      el("t_end").value = formatDate(t.start);

      console.warn("Date de fin ajustée  la date de début pour éviter une fin antérieure.");

    }

    updateTaskDatesWarning();

    // Bloquer les dates qui tombent le week-end
    const sDate = t.start ? new Date(t.start+"T00:00:00") : null;
    const eDate = t.end ? new Date(t.end+"T00:00:00") : null;
    if(sDate && !isNaN(sDate) && !isWeekday(sDate)){
      alert("La date de début tombe un week-end. Choisis un jour ouvré.");
      return;
    }
    if(eDate && !isNaN(eDate) && !isWeekday(eDate)){
      alert("La date de fin tombe un week-end. Choisis un jour ouvré.");
      return;
    }

    // Alerte doublon : même description + mêmes dates dans le même projet
    if(t.roomNumber && t.start && t.end){
      const dup = state.tasks.find(x=>
        x.id!==t.id &&
        x.projectId===t.projectId &&
        (x.roomNumber||"").trim().toLowerCase()===t.roomNumber.trim().toLowerCase() &&
        x.start===t.start &&
        x.end===t.end
      );
      if(dup){
        const ok = confirm("Attention : une tâche identique (même description + mêmes dates) existe déjà dans ce projet. Continuer quand même ?");
        if(!ok) return;
      }
    }

    t.status     = Array.from(selectedStatusSet).join(",");
    purgeTaskLogsByAssignedRole(t);

    markDirty();

    renderProject();

    refreshVendorsList();

    refreshDescriptionsList();

  });

  const debounce = (fn, wait=200)=>{
    let t;
    return (...args)=>{
      clearTimeout(t);
      t = setTimeout(()=> fn(...args), wait);
    };
  };
  ["filterSite","filterProject","filterStatus","filterStartAfter","filterEndBefore"].forEach(id=>{
    const n=el(id);
    if(n) n.addEventListener("input", ()=>{ 
      if(id==="filterSite") updateSitePhoto(n.value || "");
      renderMaster(); 
      saveUIState(); 
      markDirty(); 
    });
  });

  el("btnSaveTimeLog")?.addEventListener("click", ()=>{
    if(isLocked) return;
    if(!selectedProjectId || !selectedTaskId){
      alert("Sélectionne une tâche.");
      return;
    }
    const input = el("t_time_hours");
    const dateInput = el("t_time_date_input");
    if(!input){
      return;
    }
    const raw = (input.value || "").toString().replace(",",".").trim();
    if(!raw){
      alert("Saisis le temps passé (heures).");
      return;
    }
    const hours = parseFloat(raw);
    if(!isFinite(hours) || hours < 0){
      alert("Temps invalide.");
      return;
    }
    const minutes = Math.round(hours * 60);
    const t = state.tasks.find(x=>x.id===selectedTaskId && x.projectId===selectedProjectId);
    if(!t){
      return;
    }
    const selectedDate = (dateInput?.value || getSelectedLogDate()).trim();
    if(!selectedDate || !isTaskActiveOn(t, selectedDate)){
      alert("La date est hors période de la tâche.");
      return;
    }
    const dateObj = new Date(selectedDate+"T00:00:00");
    if(!isWeekday(dateObj)){
      alert("La date tombe un week-end.");
      return;
    }
    saveUndoSnapshot();
    const roleKey = getTaskRoleKey(t);
    upsertTimeLog(t.id, t.projectId, minutes, "", selectedDate, roleKey);
    markDirty();
    updateTimeLogUI(t, true);
    renderMaster();
    renderProject();
    saveState();
  });
  const search = el("filterSearch");
  if(search){
    const onSearch = debounce(()=>{ renderMaster(); saveUIState(); markDirty(); }, 250);
    search.addEventListener("input", onSearch);
  }
  el("toggleMissingOnly")?.addEventListener("change", ()=>{
    renderMaster();
    saveUIState();
    markDirty();
  });
  el("btnProcessMissingHours")?.addEventListener("click", ()=>{
    startMissingHoursFlow();
  });

  document.addEventListener("keydown",(e)=>{
    if(e.defaultPrevented) return;
    if(!(e.ctrlKey || e.metaKey) || e.shiftKey) return;
    if(String(e.key).toLowerCase()!=="z") return;
    const target = e.target;
    const tag = (target?.tagName || "").toLowerCase();
    if(tag==="input" || tag==="textarea" || target?.isContentEditable) return;
    if(restoreUndoSnapshot()) e.preventDefault();
  });

  setupVendorPicker();

  setupDescriptionPicker();

  // Affichage date du jour + copyright

  const brandSub = el("brandSub");

  if(brandSub){

    const today = new Date();

    const fmt = today.toLocaleDateString("fr-FR",{weekday:"long", day:"2-digit", month:"long", year:"numeric"});

    brandSub.innerHTML = `Tableau maître  Projets  Gantt  <span class="brand-date">${fmt}</span>`;

  }

  const brandTitle = el("brandTitle");

  if(brandTitle){

    brandTitle.innerHTML = `Suivi de Chantiers <span class="copyright"> Sébastien DUC</span>`;

  }

  // flatpickr sur les dates, week-ends interdits

  const fpOpts = {

    dateFormat:"Y-m-d",

    altInput:true,

    altFormat:"d/m/Y",

    allowInput:true,

    locale:"fr",

    disable:[ function(date){ const d=date.getDay(); return d===0 || d===6; } ]

  };

  let fpStart=null, fpEnd=null;

  if(window.flatpickr){

    const startNode = el("t_start");

    const endNode   = el("t_end");

    const todayIso = new Date().toISOString().slice(0,10);

    const startIso = startNode?.value ? toInputDate(startNode.value) : "";

    const endIso   = endNode?.value ? toInputDate(endNode.value) : "";

    let currentStartIso = startIso;
    let currentEndIso = endIso;

    if(startNode){

      fpStart = window.flatpickr(startNode, {...fpOpts,

        defaultDate: startIso || todayIso,

        onOpen: (_s,_d,inst)=>{ inst.jumpToDate(currentStartIso || todayIso); },

        onChange:(selectedDates, dateStr)=>{
          currentStartIso = dateStr || "";
          if(fpEnd){
            fpEnd.set("minDate", dateStr || null);
            if(dateStr) fpEnd.jumpToDate(dateStr);
          }
          if(endNode && dateStr){
            const startVal = startNode.value;
            const endVal = endNode.value;
            if(startVal && (!endVal || unformatDate(endVal) < unformatDate(startVal))){
              if(fpEnd) fpEnd.setDate(startVal, true);
              else endNode.value = startVal;
            }
          }
          updateTaskDatesWarning();
        }
      });
      window.__fpStart = fpStart;

    }

    if(endNode){

      fpEnd = window.flatpickr(endNode, {...fpOpts,
        defaultDate: endIso || startIso || todayIso,
        minDate: startIso || null,
        onOpen: (_s,_d,inst)=>{ const target = currentStartIso || currentEndIso || todayIso; inst.jumpToDate(target); },
        onChange:(_sel, dateStr)=>{
          currentEndIso = dateStr || "";
          updateTaskDatesWarning();
        }
      });
      window.__fpEnd = fpEnd;
    }
    updateTaskDatesWarning();
    ["filterStartAfter","filterEndBefore"].forEach(id=>{
      const node=el(id);

      if(node) window.flatpickr(node, fpOpts);

    });

  }

  // repositionnement des dropdowns flottants

  window.addEventListener("resize", ()=>{

    floatingMap.forEach((anchor, el)=> positionFloating(el, anchor));

  });

  window.addEventListener("scroll", ()=>{

    floatingMap.forEach((anchor, el)=> positionFloating(el, anchor));

  }, true);



  // Repositionnement du menu multiselect en fixed (pour qu'il reste au-dessus)

  const statusMenu = el("t_status_menu");

  const statusDisplay = el("t_status_display");

  if(statusMenu && statusDisplay){

    let portal = null;

    const ensurePortal = ()=>{

      if(portal) return portal;

      portal = document.createElement("div");

      portal.style.position="fixed";

      portal.style.zIndex="1000000";

      portal.style.left="0";

      portal.style.top="0";

      document.body.appendChild(portal);

      return portal;

    };

    const placeMenu = ()=>{

      const rect = statusDisplay.getBoundingClientRect();

      const p = ensurePortal();

      p.style.width = `${rect.width}px`;

      p.style.left = `${rect.left}px`;

      p.style.top = `${rect.bottom + 4}px`;

      statusMenu.style.width = `${rect.width}px`;

    };

    const openMenu = ()=>{

      placeMenu();

      const p = ensurePortal();

      p.appendChild(statusMenu);

      statusMenu.classList.remove("hidden");

      statusDisplay.classList.add("focus");

    };

    const closeMenu = ()=>{

      statusMenu.classList.add("hidden");

      statusDisplay.appendChild(statusMenu);

      statusDisplay.classList.remove("focus");

    };

    statusDisplay.addEventListener("click",(e)=>{ e.stopPropagation(); openMenu(); });

    document.addEventListener("click",(e)=>{

      if(!statusDisplay.contains(e.target) && !statusMenu.contains(e.target)){

        closeMenu();

      }

    });

    window.addEventListener("resize", placeMenu);

    window.addEventListener("scroll", placeMenu, true);

  }

  el("masterTable")?.addEventListener("click",(e)=>{
    const row=e.target.closest("tr[data-project]");
    if(!row) return;
    navigateTo(row.dataset.project, row.dataset.task, true);
  });
  el("masterTable")?.querySelectorAll("thead th[data-sort]")?.forEach(th=>{

    th.addEventListener("click", ()=>{

      const key = th.dataset.sort;

      if(sortMaster.key===key) sortMaster.dir = sortMaster.dir==="asc"?"desc":"asc";

      else { sortMaster.key=key; sortMaster.dir="asc"; }

      renderMaster();

      updateSortIndicators("masterTable", sortMaster);

      markDirty();

    });

  });

  el("btnResetSortMaster")?.addEventListener("click", ()=>{

    sortMaster = {key:"start", dir:"asc"};

    const ids = ["filterSite","filterProject","filterStatus","filterSearch","filterStartAfter","filterEndBefore"];
    ids.forEach(id=>{
      const n = el(id);
      if(n) n.value = "";
    });
    const toggleMissingOnly = el("toggleMissingOnly");
    if(toggleMissingOnly) toggleMissingOnly.checked = false;

    renderMaster();

    updateSortIndicators("masterTable", sortMaster);

    const fb = el("filtersBadge");

    if(fb) updateBadge(fb, false, "Tri/filtre actif", "Tri par défaut");

    saveUIState();
    markDirty();

  });

  el("projectTasksTable")?.addEventListener("click",(e)=>{

    const row=e.target.closest("tr[data-task]");

    if(!row) return;

    selectedTaskId=row.dataset.task;

    renderProject();

  });

  el("projectTasksTable")?.querySelectorAll("thead th[data-sort]")?.forEach(th=>{

    th.addEventListener("click", ()=>{

      const key = th.dataset.sort;

      if(sortProject.key===key) sortProject.dir = sortProject.dir==="asc"?"desc":"asc";

      else { sortProject.key=key; sortProject.dir="asc"; }

      renderProjectTasks(selectedProjectId);

      updateSortIndicators("projectTasksTable", sortProject);

      markDirty();

    });

  });

  el("btnResetSortProject")?.addEventListener("click", ()=>{

    sortProject = {key:"start", dir:"asc"};

    renderProjectTasks(selectedProjectId);

    updateSortIndicators("projectTasksTable", sortProject);

    markDirty();

  });



  // Alerte fermeture si modifications non sauvegardes

  window.addEventListener("beforeunload",(e)=>{

    if(unsavedChanges){

      e.preventDefault();

      e.returnValue="";

    }

  });



  // état visuel du bouton Sauvegarder au démarrage

  updateSaveButton();

  document.querySelectorAll("[data-accordion] .cfg-accordion-head").forEach(head=>{
    head.addEventListener("click", ()=>{
      const wrap = head.closest("[data-accordion]");
      if(!wrap) return;
      wrap.classList.toggle("is-open");
    });
  });

  // Ctrl+S / Cmd+S pour sauvegarder
  document.addEventListener("keydown",(e)=>{
    if(!(e.ctrlKey || e.metaKey) || e.shiftKey) return;
    if(String(e.key).toLowerCase()!=="s") return;
    e.preventDefault();
    el("btnSave")?.click();
  });

}



load();

bind();

// Auto‑sauvegarde toutes les 5 minutes
setInterval(()=>{
  if(!unsavedChanges) return;
  try{ saveState(); }catch(e){ softCatch(e); }
}, 5 * 60 * 1000);

try{
  history.replaceState({projectId:selectedProjectId, taskId:selectedTaskId}, "");
}catch(e){ softCatch(e); }
window.addEventListener("popstate",(e)=>{
  const st = e.state || {};
  selectedProjectId = st.projectId || null;
  selectedTaskId = st.taskId || null;
  closeAllOverlays();
  setTimeout(()=> closeAllOverlays(), 0);
  const view = document.querySelector(".main");
  if(view){
    view.classList.remove("view-fade");
    void view.offsetWidth;
    view.classList.add("view-fade");
  }
  renderAll();
  setTimeout(()=> scrollViewToTop(), 0);
});
renderAll();


// Préparation impression : cartouche + lgende

function setPrintPageFormat(size, margin="6mm"){
  let n = document.getElementById("printPageFormatOverride");
  if(!n){
    n = document.createElement("style");
    n.id = "printPageFormatOverride";
    document.head.appendChild(n);
  }
  n.textContent = `@page { size: ${size}; margin: ${margin}; }`;
}

function ensurePrintTemplate(){
  let tpl = document.getElementById("printTemplate");
  if(tpl) return tpl;
  tpl = document.createElement("template");
  tpl.id = "printTemplate";
  tpl.innerHTML = `
    <section id="printHeader" class="print-header">
      <h1>Export PDF</h1>
      <div id="printMeta" class="print-meta"></div>
      <div id="printLegend" class="print-legend"></div>
    </section>
    <section class="print-order"></section>
  `;
  document.body.appendChild(tpl);
  return tpl;
}
let __pdfLibsPromise = null;
function ensurePdfLibraries(){
  if(window?.html2canvas && window?.jspdf?.jsPDF) return Promise.resolve();
  if(__pdfLibsPromise) return __pdfLibsPromise;
  const loadScript = (src)=> new Promise((resolve,reject)=>{
    const s = document.createElement("script");
    let done = false;
    const t = setTimeout(()=>{
      if(done) return;
      done = true;
      try{ s.remove(); }catch(_e){}
      reject(new Error(`Timeout chargement script: ${src}`));
    }, 12000);
    s.src = src;
    s.async = true;
    s.onload = ()=>{
      if(done) return;
      done = true;
      clearTimeout(t);
      resolve();
    };
    s.onerror = ()=>{
      if(done) return;
      done = true;
      clearTimeout(t);
      reject(new Error(`Chargement impossible: ${src}`));
    };
    document.head.appendChild(s);
  });
  __pdfLibsPromise = (async()=>{
    if(!window?.html2canvas){
      await loadScript("https://cdnjs.cloudflare.com/ajax/libs/html2canvas/1.4.1/html2canvas.min.js");
    }
    if(!window?.jspdf?.jsPDF){
      await loadScript("https://cdnjs.cloudflare.com/ajax/libs/jspdf/2.5.1/jspdf.umd.min.js");
    }
  })();
  return __pdfLibsPromise;
}

function readPdfPageSetup(){
  const text = document.getElementById("printPageFormatOverride")?.textContent || "";
  const sizeMatch = text.match(/size\s*:\s*([^;]+);/i);
  const marginMatch = text.match(/margin\s*:\s*([0-9]+(?:\.[0-9]+)?)\s*mm\s*;/i);
  const size = String(sizeMatch?.[1] || "A4 landscape").toLowerCase();
  const marginMm = Number(marginMatch?.[1] || 2);
  return {
    format: size.includes("a3") ? "a3" : "a4",
    orientation: size.includes("portrait") ? "portrait" : "landscape",
    marginMm: Number.isFinite(marginMm) ? Math.max(1, marginMm) : 2
  };
}

async function openPreparedPrintInNewWindow(title="Export PDF", viewerRef=null){
  const container = document.getElementById("printInjection");
  if(!container || !container.innerHTML.trim()){
    alert("Aucun contenu d'export à imprimer.");
    return;
  }

  const viewer = viewerRef || window.open("about:blank", "_blank");
  if(!viewer){
    alert("Le popup PDF est bloqué. Autorisez les popups puis recommencez.");
    return;
  }

  viewer.document.open();
  viewer.document.write(`<!doctype html><html><head><meta charset="utf-8"><title>${attrEscape(title)}</title><style>
    :root{color-scheme:light;}
    *{box-sizing:border-box}
    body{margin:0;min-height:100vh;display:flex;align-items:center;justify-content:center;padding:18px;background:#eef3f7;font-family:"Segoe UI",Arial,sans-serif;color:#0f172a}
    .pdf-progress-card{display:inline-flex;flex-direction:column;max-width:92vw;background:linear-gradient(180deg, rgba(255,255,255,0.98) 0%, rgba(233,231,225,0.96) 100%);border:1px solid rgba(0,0,0,0.12);border-radius:10px;box-shadow:0 10px 22px rgba(0,0,0,0.18);overflow:hidden}
    .pdf-progress-body{display:flex;align-items:center;gap:20px;padding:24px 32px}
    .pdf-progress-ok{width:48px;height:48px;border-radius:8px;border:1px solid rgba(0,0,0,0.12);background:#d1fae5;display:flex;align-items:center;justify-content:center;font-weight:800;font-size:16px;color:#0f172a;flex:none}
    .pdf-progress-main{flex:1;min-width:0}
    .pdf-progress-title{font-size:18px;line-height:1.1;letter-spacing:.1px;font-weight:800;color:#0f172a;margin:0 0 4px}
    .pdf-progress-sub{font-size:16px;line-height:1.2;margin:0;color:#475569}
    .pdf-progress-step{font-size:13px;line-height:1.2;margin-top:6px;color:#64748b}
    .pdf-progress-bar-wrap{height:4px;background:rgba(148,163,184,0.25)}
    .pdf-progress-bar{height:100%;width:0%;background:linear-gradient(90deg,#34d399 0%, #22c55e 100%);transition:width .2s ease}
    @media (max-width:900px){
      .pdf-progress-body{padding:16px 16px;gap:12px}
      .pdf-progress-title{font-size:16px}
      .pdf-progress-sub{font-size:14px}
      .pdf-progress-step{font-size:12px}
    }
  </style></head><body>
    <div class="pdf-progress-card" role="status" aria-live="polite">
      <div class="pdf-progress-body">
        <div class="pdf-progress-ok">OK</div>
        <div class="pdf-progress-main">
          <h1 class="pdf-progress-title">GENERATION PDF</h1>
          <p class="pdf-progress-sub" id="pdfProgressSub">PREPARATION EN COURS</p>
          <div class="pdf-progress-step" id="pdfProgressStep">INITIALISATION...</div>
        </div>
      </div>
      <div class="pdf-progress-bar-wrap"><div class="pdf-progress-bar" id="pdfProgressBar"></div></div>
    </div>
  </body></html>`);
  viewer.document.close();
  const updateViewerProgress = (percent, stepText="", subText="")=>{
    try{
      const p = Math.max(0, Math.min(100, Number(percent) || 0));
      const bar = viewer.document.getElementById("pdfProgressBar");
      const step = viewer.document.getElementById("pdfProgressStep");
      const sub = viewer.document.getElementById("pdfProgressSub");
      if(bar) bar.style.width = `${p}%`;
      if(stepText && step) step.textContent = String(stepText).toUpperCase();
      if(subText && sub) sub.textContent = String(subText).toUpperCase();
    }catch(e){ softCatch(e); }
  };
  updateViewerProgress(4, "Initialisation...", "Préparation en cours");

  const host = document.createElement("div");
  host.style.position = "fixed";
  host.style.left = "-20000px";
  host.style.top = "0";
  host.style.width = "1800px";
  host.style.background = "#fff";
  host.style.zIndex = "-1";

  const page = document.createElement("div");
  page.className = "print-mode";
  page.style.background = "#fff";
  page.innerHTML = container.innerHTML;
  host.appendChild(page);
  document.body.appendChild(host);

  try{
    updateViewerProgress(10, "Chargement des bibliothèques PDF...", "Préparation en cours");
    await ensurePdfLibraries();
    updateViewerProgress(18, "Préparation du rendu...", "Préparation en cours");
    await new Promise(r=>setTimeout(r, 80));

    const setup = readPdfPageSetup();
    const force = document.createElement("style");
    force.textContent = `
      *{animation:none !important;transition:none !important;}
      .gantt-bar,.gantt-print-bar,.bar-wrapper,.gantt-days{opacity:1 !important;visibility:visible !important;transform:none !important;}
      body.print-mode .gantt-print-bar{
        display:block !important;
        height:10px !important;
        min-height:10px !important;
        border-radius:3px !important;
        border:1px solid rgba(0,0,0,0.35) !important;
        box-sizing:border-box !important;
        background-clip:padding-box !important;
      }
      body.print-mode .gantt-cell-inner{
        min-height:14px !important;
      }
      body.print-mode .bar-wrapper{
        min-height:12px !important;
        width:100% !important;
      }
      body.print-mode .gantt-table th.week-cell{
        text-align:center !important;
        vertical-align:middle !important;
        padding:2px 1px !important;
        line-height:1.05 !important;
      }
      body.print-mode .gantt-table th.week-cell .gantt-week-date{
        display:block !important;
        margin-top:2px !important;
        line-height:1 !important;
        font-size:9px !important;
      }
      body.print-mode .pdf-two-graphs-wrap{
        display:grid !important;
        grid-template-columns:1fr 1fr !important;
        gap:3mm !important;
        align-items:stretch !important;
        min-height:170mm !important;
        place-items:center stretch !important;
      }
      body.print-mode .pdf-two-graphs-item{
        min-width:0 !important;
        height:100% !important;
        display:flex !important;
        align-items:center !important;
      }
      body.print-mode .pdf-two-graphs-item .card{
        margin-bottom:0 !important;
        width:100% !important;
        min-height:158mm !important;
        height:158mm !important;
        border:1px solid #cbd5e1 !important;
        border-radius:8px !important;
        padding:2.5mm !important;
        box-shadow:0 2mm 5mm rgba(15,23,42,.14) !important;
        display:flex !important;
        flex-direction:column !important;
        justify-content:flex-start !important;
      }
      body.print-mode .pdf-two-graphs-item #workloadChartProjectWrap,
      body.print-mode .pdf-two-graphs-item #workloadPieProjectWrap{
        flex:1 1 auto !important;
        min-height:0 !important;
        overflow:visible !important;
        padding-bottom:3mm !important;
      }
      body.print-mode .pdf-two-graphs-item #workloadChartProject,
      body.print-mode .pdf-two-graphs-item #workloadPieProject{
        height:130mm !important;
        width:100% !important;
      }
      body.print-mode .pdf-workload-legend{
        margin-top:1.5mm !important;
        display:flex !important;
        justify-content:center !important;
        align-items:center !important;
        flex-wrap:wrap !important;
        gap:4mm !important;
      }
      body.print-mode .pdf-workload-legend-item{
        display:inline-flex !important;
        align-items:center !important;
        gap:1.5mm !important;
        font-size:10px !important;
        color:#0f172a !important;
        font-weight:700 !important;
      }
      body.print-mode .pdf-workload-legend-dot{
        width:8px !important;
        height:8px !important;
        border-radius:999px !important;
        display:inline-block !important;
      }
      body.print-mode .pdf-two-graphs-item svg{
        width:100% !important;
        max-height:146mm !important;
        overflow:visible !important;
      }
      body.print-mode .pdf-two-graphs-item .apexcharts-canvas,
      body.print-mode .pdf-two-graphs-item .apexcharts-svg,
      body.print-mode .pdf-two-graphs-item .apexcharts-inner{
        overflow:visible !important;
      }
      body.print-mode .card,
      body.print-mode .tablewrap,
      body.print-mode .gantt-table,
      body.print-mode .panel,
      body.print-mode .print-header,
      body.print-mode .print-block{
        border:none !important;
        box-shadow:none !important;
        outline:none !important;
        background:#fff !important;
      }
      body.print-mode .pdf-two-graphs-item .card{
        border:1px solid #cbd5e1 !important;
        box-shadow:0 2mm 5mm rgba(15,23,42,.14) !important;
        overflow:visible !important;
      }
      body.print-mode .card-title{
        margin-bottom:4px !important;
      }
      body.print-mode .num-badge{
        display:inline-flex !important;
        align-items:center !important;
        justify-content:center !important;
        width:26px !important;
        min-width:26px !important;
        max-width:26px !important;
        height:22px !important;
        padding:0 !important;
        margin-right:4px !important;
        border-radius:999px !important;
        line-height:1 !important;
        font-size:12px !important;
        font-weight:800 !important;
        font-variant-numeric: tabular-nums !important;
        font-feature-settings: "tnum" 1 !important;
        letter-spacing:0 !important;
        text-indent:0 !important;
        vertical-align:middle !important;
        opacity:1 !important;
        filter:none !important;
        border:none !important;
        box-shadow:0 2mm 5mm rgba(15,23,42,.14) !important;
        text-shadow:none !important;
        background-image:none !important;
        -webkit-text-stroke:0 !important;
      }
      body.print-mode .badge.owner{
        opacity:1 !important;
        filter:none !important;
      }
      body.print-mode #printHeader{
        width:100% !important;
        box-sizing:border-box !important;
      }
      body.print-mode #printMeta{
        display:grid !important;
        grid-template-columns:repeat(4,minmax(0,1fr)) !important;
        gap:2mm !important;
        width:100% !important;
        box-sizing:border-box !important;
      }
      body.print-mode #printMeta > div{
        min-width:0 !important;
        line-height:1.35 !important;
      }
      @media (max-width:1200px){
        body.print-mode #printMeta{
          grid-template-columns:repeat(2,minmax(0,1fr)) !important;
        }
      }
    `;
    host.appendChild(force);

    const { jsPDF } = window.jspdf;
    const pdf = new jsPDF({ orientation: setup.orientation, unit: "mm", format: setup.format, compress: true });

    const pageW = pdf.internal.pageSize.getWidth();
    const pageH = pdf.internal.pageSize.getHeight();
    const margin = setup.marginMm || 2;
    const drawW = pageW - (margin * 2);
    const drawH = pageH - (margin * 2);

    const blocks = Array.from(page.querySelectorAll('.print-block'));
    const targets = blocks.length ? blocks : [page];
    const totalTargets = Math.max(1, targets.length);
    updateViewerProgress(24, `${totalTargets} module(s) à rendre...`, "Rendu des pages");

    let yCursor = margin;
    let hasContentOnPage = false;

    for(let i=0; i<targets.length; i++){
      const block = targets[i];
      const loopStartPct = 24 + Math.round((i / totalTargets) * 60);
      updateViewerProgress(loopStartPct, `Rendu du module ${i+1}/${totalTargets}...`, "Rendu des pages");
      if(i>0 && block.classList && block.classList.contains("force-new-page")){
        pdf.addPage();
        yCursor = margin;
        hasContentOnPage = false;
      }
      const canvas = await window.html2canvas(block, {
        scale: 2,
        backgroundColor: "#ffffff",
        useCORS: true,
        allowTaint: true,
        logging: false,
        windowWidth: Math.max(block.scrollWidth || block.clientWidth || 1800, 1800),
        windowHeight: Math.max(block.scrollHeight || block.clientHeight || 300, 300)
      });

      const imgData = canvas.toDataURL("image/jpeg", 0.95);
      const imgH = canvas.height * drawW / canvas.width;

      if(yCursor + imgH <= (margin + drawH)){
        pdf.addImage(imgData, "JPEG", margin, yCursor, drawW, imgH, undefined, "FAST");
        yCursor += imgH + 2;
        hasContentOnPage = true;
        continue;
      }

      if(imgH <= drawH){
        if(hasContentOnPage){
          pdf.addPage();
          yCursor = margin;
          hasContentOnPage = false;
        }
        pdf.addImage(imgData, "JPEG", margin, yCursor, drawW, imgH, undefined, "FAST");
        yCursor += imgH + 2;
        hasContentOnPage = true;
        continue;
      }

      const slicePx = Math.floor((drawH * canvas.width) / drawW);
      let offsetPx = 0;
      while(offsetPx < canvas.height){
        if(hasContentOnPage){
          pdf.addPage();
          yCursor = margin;
          hasContentOnPage = false;
        }

        const hPx = Math.min(slicePx, canvas.height - offsetPx);
        const slice = document.createElement('canvas');
        slice.width = canvas.width;
        slice.height = hPx;
        const sctx = slice.getContext('2d');
        sctx.drawImage(canvas, 0, offsetPx, canvas.width, hPx, 0, 0, canvas.width, hPx);

        const sliceData = slice.toDataURL('image/jpeg', 0.95);
        const sliceH = hPx * drawW / canvas.width;
        pdf.addImage(sliceData, 'JPEG', margin, yCursor, drawW, sliceH, undefined, 'FAST');
        hasContentOnPage = true;

        offsetPx += hPx;
      }
      yCursor = margin;
      const loopEndPct = 24 + Math.round(((i + 1) / totalTargets) * 60);
      updateViewerProgress(loopEndPct, `Module ${i+1}/${totalTargets} prêt`, "Rendu des pages");
    }

    updateViewerProgress(90, "Assemblage final du PDF...", "Finalisation");
    const blob = pdf.output("blob");
    updateViewerProgress(97, "Ouverture du document...", "Finalisation");
    const blobUrl = URL.createObjectURL(blob);
    updateViewerProgress(100, "Terminé", "PDF prêt");
    viewer.location.href = blobUrl;
    setTimeout(()=>{ try{ URL.revokeObjectURL(blobUrl); }catch(e){ softCatch(e); } }, 120000);
  }catch(err){
    console.error("PDF generation failed", err);
    try{
      viewer.document.open();
      viewer.document.write(`<!doctype html><html><head><meta charset="utf-8"><title>Erreur export PDF</title></head><body style="font-family:Segoe UI,Arial,sans-serif;padding:16px"><h2>Erreur pendant la génération du PDF</h2><pre style="white-space:pre-wrap;color:#b91c1c">${attrEscape(String(err?.message || err || "Erreur inconnue"))}</pre></body></html>`);
      viewer.document.close();
    }catch(e){ softCatch(e); }
    alert("Erreur pendant la génération du PDF.");
  }finally{
    try{ host.remove(); }catch(e){ softCatch(e); }
    try{ cleanupPrint(); }catch(e){ softCatch(e); }
  }
}

function maximizePrintContainer(container){
  if(!container) return;
  container.style.zoom = 1;
  container.style.width = "100%";
  const pageW = window.innerWidth || document.documentElement.clientWidth || 1;
  const contentW = container.scrollWidth || container.offsetWidth || 1;
  let scale = pageW / contentW;
  if(!isFinite(scale) || scale <= 0) scale = 1;
  scale = Math.min(1.18, Math.max(1, scale));
  container.style.zoom = scale;
}

function buildRealHoursReportForTasks(tasksInput){
  const tasks = (tasksInput || []).filter(t=>t && t.start && t.end);
  const logs = getCanonicalTimeLogs().filter(l=>l && l.taskId && l.date && (l.minutes||0) > 0);
  const roleOrder = ["interne","rsg","ri"];
  const roleTotals = {interne:0, externe:0, rsg:0, ri:0};
  const vendorTotals = new Map();
  const internalByName = new Map();
  const externalByName = new Map();
  const detailRows = [];

  tasks.sort((a,b)=>{
    const oa=(taskOrderMap[a.id]||9999)-(taskOrderMap[b.id]||9999);
    if(oa!==0) return oa;
    return taskTitleProjectView(a).localeCompare(taskTitleProjectView(b), "fr", {sensitivity:"base"});
  });

  tasks.forEach(t=>{
    const isExternalTask = !!String(t.vendor || "").trim() || ownerType(t.owner) === "externe";
    const extName = (t.vendor || "Prestataire non renseigné").trim() || "Prestataire non renseigné";
    const taskLogs = logs.filter(l=>
      l.taskId===t.id &&
      l.date >= t.start &&
      l.date <= t.end
    );
    const perRole = new Map();
    taskLogs.forEach(l=>{
      const logRole = normalizeTimeLogRole(l);
      if(isExternalTask && logRole !== "externe") return;
      if(!isExternalTask && logRole === "externe") return;
      const rk = isExternalTask ? "externe" : logRole;
      const mins = Number(l.minutes)||0;
      if(!mins) return;
      const weightedMinutes = Math.round(mins * roleHoursMultiplier(rk));
      const prev = perRole.get(rk) || 0;
      perRole.set(rk, prev + weightedMinutes);

      if(isExternalTask || rk==="externe"){
        externalByName.set(extName, (externalByName.get(extName) || 0) + weightedMinutes);
        vendorTotals.set(extName, (vendorTotals.get(extName) || 0) + weightedMinutes);
      }else{
        const intName = roleLabel(rk);
        internalByName.set(intName, (internalByName.get(intName) || 0) + weightedMinutes);
      }
    });
    perRole.forEach((mins, rk)=>{
      if(!mins) return;
      if(roleTotals[rk] !== undefined) roleTotals[rk] += mins;
      detailRows.push({
        num: taskOrderMap[t.id] || "",
        task: taskTitleProjectView(t),
        role: rk,
        interv: rk==="externe" ? extName : roleLabel(rk),
        mins
      });
    });
  });

  const totalMinutes = Object.values(roleTotals).reduce((s,v)=>s+v,0);
  const summaryRows = roleOrder.map(rk=>({
    label: roleLabel(rk),
    mins: roleTotals[rk] || 0
  }));
  const externalVendors = Array.from(vendorTotals.entries())
    .sort((a,b)=>a[0].localeCompare(b[0], "fr", {sensitivity:"base"}));
  const internalByNameRows = Array.from(internalByName.entries())
    .sort((a,b)=>a[0].localeCompare(b[0], "fr", {sensitivity:"base"}));
  const externalByNameRows = Array.from(externalByName.entries())
    .sort((a,b)=>a[0].localeCompare(b[0], "fr", {sensitivity:"base"}));
  const internalTotalMinutes = roleTotals.interne + roleTotals.rsg + roleTotals.ri;
  const externalTotalMinutes = roleTotals.externe;

  return {
    summaryRows,
    externalVendors,
    detailRows,
    totalMinutes,
    internalByNameRows,
    externalByNameRows,
    internalTotalMinutes,
    externalTotalMinutes
  };
}

function buildProjectRealHoursReport(projectId){
  const tasks = (state.tasks || []).filter(t=>t.projectId===projectId && t.start && t.end);
  return buildRealHoursReportForTasks(tasks);
}

function buildGroupedProjectsRealHoursReport(projectIds){
  const idSet = new Set((projectIds || []).map(normId).filter(Boolean));
  const tasks = (state.tasks || []).filter(t=>idSet.has(t.projectId) && t.start && t.end);
  return buildRealHoursReportForTasks(tasks);
}

function buildMasterRealHoursReport(){
  const tasks = (filteredTasks() || []).filter(t=>t.start && t.end);
  return buildRealHoursReportForTasks(tasks);
}

function filterRealHoursReportExternal(rep, includeExternal=true){
  if(includeExternal) return rep;
  const internalKinds = new Set(["INTERNE","RSG","RI"]);
  const detailRows = (rep.detailRows || []).filter(r=> internalKinds.has(String(r.interv || "").toUpperCase()));
  const internalByNameRows = Array.isArray(rep.internalByNameRows) ? rep.internalByNameRows : [];
  const internalTotalMinutes = internalByNameRows.reduce((s, row)=> s + (Number(row?.[1]) || 0), 0);
  return {
    summaryRows: (rep.summaryRows || []).map(r=>{
      const label = String(r?.label || "");
      if(label.toUpperCase() === "EXTERNE") return {...r, mins: 0};
      return r;
    }),
    externalVendors: [],
    detailRows,
    totalMinutes: internalTotalMinutes,
    internalByNameRows,
    externalByNameRows: [],
    internalTotalMinutes,
    externalTotalMinutes: 0
  };
}

function buildRealHoursReportInnerHTML(rep, title, reportMode="master", includeExternal=true){
  const report = filterRealHoursReportExternal(rep, includeExternal);
  const summary = report.summaryRows.map(r=>
    `<tr class="report-internal-priority"><td>${r.label}</td><td style="text-align:right">${formatHoursMinutes(r.mins)}</td></tr>`
  ).join("");
  const vendors = report.externalVendors.length
    ? report.externalVendors.map(([name, mins])=>
        `<tr><td>${attrEscape(name)}</td><td style="text-align:right">${formatHoursMinutes(mins)}</td></tr>`
      ).join("")
    : `<tr><td colspan="2" class="text-muted">Aucun détail externe.</td></tr>`;
  const details = report.detailRows.length
    ? report.detailRows.map(r=>
        `<tr><td>${r.num}</td><td>${attrEscape(r.task)}</td><td>${attrEscape(r.interv)}</td><td style="text-align:right">${formatHoursMinutes(r.mins)}</td></tr>`
      ).join("")
    : `<tr><td colspan="4" class="text-muted">Aucune heure réelle saisie.</td></tr>`;
  const detailTotalMinutes = report.detailRows.reduce((s, r)=> s + (Number(r.mins) || 0), 0);
  const totalsByNameRows = [
    ...report.internalByNameRows.map(([name, mins])=>({
      kind:"Interne",
      name,
      mins
    })),
    ...report.externalByNameRows.map(([name, mins])=>({
      kind:"Externe",
      name,
      mins
    }))
  ];
  const totalsByNameHtml = totalsByNameRows.length
    ? totalsByNameRows.map(r=>
      `<tr><td>${r.kind}</td><td>${attrEscape(r.name)}</td><td style="text-align:right">${formatHoursMinutes(r.mins)}</td></tr>`
    ).join("")
    : `<tr><td colspan="3" class="text-muted">Aucune donnée.</td></tr>`;

  const vendorsBlock = includeExternal
    ? `<div>
        <div class="report-subtitle">Détail Externe (prestataires)</div>
        <table class="report-table">
          <thead><tr><th>Prestataire</th><th>Heures</th></tr></thead>
          <tbody>${vendors}</tbody>
        </table>
      </div>`
    : "";

  const totalsSubtitle = includeExternal
    ? "Totaux internes / externes par intervenant (nom)"
    : "Totaux internes par intervenant (nom)";

  return `
    <div class="row row-compact" style="justify-content:space-between;align-items:center;margin-bottom:6px;">\n      <div class="card-title">${attrEscape(title || "Analyse heures réelles")}</div>\n    </div>
    <div class="report-grid-two"${includeExternal ? "" : ' style="display:block"'}>
      <div>
        <div class="report-subtitle">Synthèse par intervenant</div>
        <table class="report-table report-table-summary">
          <thead><tr><th>Intervenant</th><th>Heures</th></tr></thead>
          <tbody>${summary}</tbody>
          <tfoot><tr><th>Total</th><th style="text-align:right">${formatHoursMinutes(report.internalTotalMinutes)}</th></tr></tfoot>
        </table>
      </div>
      ${vendorsBlock}
    </div>
    <div class="report-subtitle">Détail tâche + intervenant</div>
    <table class="report-table">
      <thead><tr><th>N</th><th>Tâche</th><th>Intervenant</th><th>Heures</th></tr></thead>
      <tbody>${details}</tbody>
      <tfoot><tr><th colspan="3">Total</th><th style="text-align:right">${formatHoursMinutes(detailTotalMinutes)}</th></tr></tfoot>
    </table>
    <div class="report-subtitle" style="margin-top:8px">${totalsSubtitle}</div>
    <table class="report-table">
      <thead><tr><th>Catégorie</th><th>Nom intervenant</th><th>Heures</th></tr></thead>
      <tbody>${totalsByNameHtml}</tbody>
      <tfoot>
        <tr><th colspan="2">Total interne</th><th style="text-align:right">${formatHoursMinutes(report.internalTotalMinutes)}</th></tr>
        ${includeExternal ? `<tr><th colspan="2">Total externe</th><th style="text-align:right">${formatHoursMinutes(report.externalTotalMinutes)}</th></tr>` : ""}
        <tr><th colspan="2">Total heures réelles</th><th style="text-align:right">${formatHoursMinutes(report.totalMinutes)}</th></tr>
      </tfoot>
    </table>
  `;
}

function buildProjectRealHoursReportInnerHTML(projectId, includeExternal=true, title="Analyse heures réelles (projet)"){
  const rep = buildProjectRealHoursReport(projectId);
  return buildRealHoursReportInnerHTML(rep, title, "project", includeExternal);
}

function buildGroupedProjectsRealHoursReportInnerHTML(projectIds, includeExternal=true, title="Analyse heures réelles (chantiers regroupés)"){
  const rep = buildGroupedProjectsRealHoursReport(projectIds);
  return buildRealHoursReportInnerHTML(rep, title, "project", includeExternal);
}

function buildProjectRepartitionExportInnerHTML(projectId, includeExternal=true, title="Répartition Interne / Externe / RSG / RI (projet)"){
  const rep = filterRealHoursReportExternal(buildProjectRealHoursReport(projectId), includeExternal);
  const totalMinutes = Number(rep.totalMinutes) || 0;
  const rowsHtml = (rep.summaryRows || []).map((r)=>{
    const label = String(r?.label || "");
    const mins = Number(r?.mins) || 0;
    const part = totalMinutes > 0 ? ((mins / totalMinutes) * 100) : 0;
    const partTxt = `${part.toFixed(1).replace(".", ",")} %`;
    return `<tr><td>${attrEscape(label)}</td><td style="text-align:right">${formatHoursMinutes(mins)}</td><td style="text-align:right">${partTxt}</td></tr>`;
  }).join("") || `<tr><td colspan="3" style="text-align:center;color:#64748b;">Aucune donnée</td></tr>`;
  const total = formatHoursMinutes(rep.totalMinutes || 0);
  return `
    <div class="card-title">${attrEscape(title)}</div>
    <table class="report-table report-table-summary" style="margin-top:6px">
      <thead><tr><th>Catégorie</th><th>Heures</th><th>Part</th></tr></thead>
      <tbody>${rowsHtml}</tbody>
      <tfoot><tr><th>Total</th><th style="text-align:right">${total}</th><th style="text-align:right">100,0 %</th></tr></tfoot>
    </table>
  `;
}

function buildGroupedProjectsRepartitionExportInnerHTML(projectIds, includeExternal=true, title="Répartition Interne / Externe / RSG / RI (chantiers regroupés)"){
  const rep = filterRealHoursReportExternal(buildGroupedProjectsRealHoursReport(projectIds), includeExternal);
  const totalMinutes = Number(rep.totalMinutes) || 0;
  const rowsHtml = (rep.summaryRows || []).map((r)=>{
    const label = String(r?.label || "");
    const mins = Number(r?.mins) || 0;
    const part = totalMinutes > 0 ? ((mins / totalMinutes) * 100) : 0;
    const partTxt = `${part.toFixed(1).replace(".", ",")} %`;
    return `<tr><td>${attrEscape(label)}</td><td style="text-align:right">${formatHoursMinutes(mins)}</td><td style="text-align:right">${partTxt}</td></tr>`;
  }).join("") || `<tr><td colspan="3" style="text-align:center;color:#64748b;">Aucune donnée</td></tr>`;
  const total = formatHoursMinutes(rep.totalMinutes || 0);
  return `
    <div class="card-title">${attrEscape(title)}</div>
    <table class="report-table report-table-summary" style="margin-top:6px">
      <thead><tr><th>Catégorie</th><th>Heures</th><th>Part</th></tr></thead>
      <tbody>${rowsHtml}</tbody>
      <tfoot><tr><th>Total</th><th style="text-align:right">${total}</th><th style="text-align:right">100,0 %</th></tr></tfoot>
    </table>
  `;
}

function buildMasterRepartitionExportInnerHTML(title="Répartition Interne / Externe / RSG / RI (tableau maître)"){
  const rep = filterRealHoursReportExternal(buildMasterRealHoursReport(), true);
  const totalMinutes = Number(rep.totalMinutes) || 0;
  const rowsHtml = (rep.summaryRows || []).map((r)=>{
    const label = String(r?.label || "");
    const mins = Number(r?.mins) || 0;
    const part = totalMinutes > 0 ? ((mins / totalMinutes) * 100) : 0;
    const partTxt = `${part.toFixed(1).replace(".", ",")} %`;
    return `<tr><td>${attrEscape(label)}</td><td style="text-align:right">${formatHoursMinutes(mins)}</td><td style="text-align:right">${partTxt}</td></tr>`;
  }).join("") || `<tr><td colspan="3" style="text-align:center;color:#64748b;">Aucune donnée</td></tr>`;
  const total = formatHoursMinutes(rep.totalMinutes || 0);
  return `
    <div class="card-title">${attrEscape(title)}</div>
    <table class="report-table report-table-summary" style="margin-top:6px">
      <thead><tr><th>Catégorie</th><th>Heures</th><th>Part</th></tr></thead>
      <tbody>${rowsHtml}</tbody>
      <tfoot><tr><th>Total</th><th style="text-align:right">${total}</th><th style="text-align:right">100,0 %</th></tr></tfoot>
    </table>
  `;
}

function buildMasterRealHoursReportInnerHTML(includeExternal=true){
  const rep = buildMasterRealHoursReport();
  return buildRealHoursReportInnerHTML(rep, "Analyse heures réelles (tableau maître)", "master", includeExternal);
}

function buildProjectRealHoursReportHTML(projectId, includeExternal=true){
  return `<div class="card print-block report-hours-card">${buildProjectRealHoursReportInnerHTML(projectId, includeExternal)}</div>`;
}
function buildMasterRealHoursReportHTML(includeExternal=true){
  return `<div class="card print-block report-hours-card">${buildMasterRealHoursReportInnerHTML(includeExternal)}</div>`;
}



function cleanupPrint(){

  document.body.classList.remove("print-mode");
  document.body.classList.remove("print-gantt-master");
  document.body.classList.remove("print-gantt-project");
  document.body.classList.remove("print-hours-report");

  const container = document.getElementById("printInjection");

  if(container) container.innerHTML = "";

}



if(typeof window !== "undefined"){

  window.onafterprint = cleanupPrint;

}

function buildMasterFiltersLabel(){
  const labels = [];
  const exportSites = (ganttExportContext==="project") ? [] : getSelectedExportSites();
  const allSites = (ganttExportContext==="project") ? [] : getAllSitesList();
  if(allSites.length){
    if(exportSites.length === 0 || exportSites.length === allSites.length){
      labels.push("Sites: Tous");
    }else{
      labels.push(`Sites: ${exportSites.join(", ")}`);
    }
  }else{
    const site = (el("filterSite")?.value || "").trim();
    if(site) labels.push(`Site: ${site}`);
  }
  const projId = el("filterProject")?.value || "";
  if(projId){
    const p = state.projects.find(x=>x.id===projId);
    if(p) labels.push(`Projet: ${p.name || "Sans nom"}`);
  }
  const statusId = el("filterStatus")?.value || "";
  if(statusId){
    const s = STATUSES.find(x=>x.v===statusId);
    labels.push(`Statut: ${(s?.label || statusId)}`);
  }
  const q = (el("filterSearch")?.value || "").trim();
  if(q) labels.push(`Recherche: "${q}"`);
  const startAfter = el("filterStartAfter")?.value || "";
  if(startAfter) labels.push(`Début après: ${formatDate(startAfter)}`);
  const endBefore = el("filterEndBefore")?.value || "";
  if(endBefore) labels.push(`Fin avant: ${formatDate(endBefore)}`);

  const rangeType = el("ganttExportRangeType")?.value || "";
  const rangeYear = el("ganttExportRangeYear")?.value || "";
  const rangeStart = el("ganttExportRangeStart")?.value || "";
  const rangeEnd = el("ganttExportRangeEnd")?.value || "";
  if(rangeType){
    if(rangeType === "all") labels.push("Période: Toutes les dates");
    else if(rangeType === "civil") labels.push(`Période: Année civile ${rangeYear || ""}`.trim());
    else if(rangeType === "school"){
      const y = parseInt(rangeYear || "", 10);
      const label = isNaN(y) ? "" : `${y}-${y+1}`;
      labels.push(`Période: Année scolaire ${label || rangeYear || ""}`.trim());
    }
    else if(rangeType === "custom"){
      if(rangeStart || rangeEnd){
        const s = rangeStart ? formatDate(rangeStart) : "—";
        const e = rangeEnd ? formatDate(rangeEnd) : "—";
        labels.push(`Période: ${s} → ${e}`);
      }
    }
  }
  return labels.length ? labels.join(" • ") : "Aucun";
}

















































function buildProjectGanttPdfStaticTable(rangeStart, rangeEnd, tasksAllOverride=null){
  const tasksAll = (tasksAllOverride || []).filter(t=>t.start && t.end);
  const tasks = (rangeStart && rangeEnd) ? tasksAll.filter(t=>{
    const s = new Date(t.start+"T00:00:00");
    const e = new Date(t.end+"T00:00:00");
    return e >= rangeStart && s <= rangeEnd;
  }) : tasksAll;

  if(tasks.length===0){
    return "<div style='padding:10px;font-size:12px;'>Aucune tâche dans cette période.</div>";
  }

  tasks.sort((a,b)=>{
    const oa=(taskOrderMap[a.id]||9999)-(taskOrderMap[b.id]||9999);
    if(oa!==0) return oa;
    const sa=Date.parse(a.start||"9999-12-31"), sb=Date.parse(b.start||"9999-12-31");
    if(sa!==sb) return sa-sb;
    return taskTitle(a).localeCompare(taskTitle(b),"fr",{sensitivity:"base"});
  });

  const displayStart = startOfWeek(rangeStart || tasks.map(t=>new Date(t.start+"T00:00:00")).reduce((a,b)=>a<b?a:b));
  const displayEnd = endOfWeek(rangeEnd || tasks.map(t=>new Date(t.end+"T00:00:00")).reduce((a,b)=>a>b?a:b));

  const weeks=[];
  for(let w=startOfWeek(displayStart); w<=addDays(startOfWeek(displayEnd),0); w=addDays(w,7)) weeks.push(new Date(w));

  let html = "<table class='pdf-gantt-table'><thead><tr>";
  html += "<th class='c-task'>Tâche</th><th class='c-vendor'>Prestataire</th><th class='c-status'>Statut</th>";
  weeks.forEach((w)=>{
    const info=isoWeekInfo(w);
    html += `<th class='c-week'>${info.week}<div class='wk-date'>${formatShortDateTwoLinesHTML(w)}</div></th>`;
  });
  html += "</tr></thead><tbody>";

  tasks.forEach((t)=>{
    const statuses = parseStatuses(t.status).map(v=>v.toUpperCase());
    const mainStatus = statuses[0] || "";
    const color = ownerColor(t.owner);
    const p = state?.projects?.find(x=>x.id===t.projectId);
    const sub = (p?.subproject || "").trim();
    const taskDesc = (t.roomNumber || "").trim();
    const label = [sub, taskDesc].filter(Boolean).join(" - ") || (taskTitle(t) || "-");

    const vendorText = (()=>{
      const typ = ownerType(t.owner);
      if(t.vendor) return t.vendor;
      if(typ === "interne") return "INTERNE";
      if(typ === "rsg") return "RSG";
      if(typ === "ri") return "RI";
      if(typ === "externe") return "Prestataire non renseigné";
      return "-";
    })();

    html += "<tr>";
    html += `<td class='c-task'>${attrEscape(label)}</td>`;
    html += `<td class='c-vendor'>${attrEscape(vendorText)}</td>`;
    html += `<td class='c-status'>${attrEscape(statusLabels(mainStatus))}</td>`;

    weeks.forEach((w)=>{
      const sDate=new Date(t.start+"T00:00:00");
      const eDate=new Date(t.end+"T00:00:00");
      const geo=barGeometry(sDate,eDate,w);
      if(geo.days>0){
        html += `<td class='c-week'><div class='pdf-bar-wrap'><div class='pdf-bar' style='width:${geo.width}%;margin-left:${geo.offset}%;background:${color};border-color:${color};'></div></div></td>`;
      }else{
        html += "<td class='c-week'></td>";
      }
    });

    html += "</tr>";
  });

  html += "</tbody></table>";
  return html;
}



