
Claude is responding
Mehema, copy karaganna lesi wenna:

rules_version = '2';
service cloud.firestore {
  match /databases/{database}/documents {

    function signedIn() { return request.auth != null; }
    function me() { return get(/databases/$(database)/documents/employees/$(request.auth.uid)).data; }
    function isSuper() { return signedIn() && me().role == 'superadmin'; }
    function isDeptAdminOf(dept) {
      return signedIn() && me().role == 'deptadmin' && dept != null && me().department == dept;
    }
    function setupDone() { return exists(/databases/$(database)/documents/system/status); }

    match /system/{docId} {
      allow read: if true;
      allow create: if signedIn() && !setupDone();
      allow update, delete: if isSuper();
    }

    match /employees/{uid} {
      allow read: if signedIn() && (
        request.auth.uid == uid || isSuper() || isDeptAdminOf(resource.data.department)
      );
      allow create: if signedIn() && (
        (request.auth.uid == uid && request.resource.data.role == 'superadmin' && !setupDone())
        || isSuper()
        || (isDeptAdminOf(request.resource.data.department) && request.resource.data.role == 'employee')
      );
      allow update: if isSuper() || (
        request.auth.uid == uid
        && request.resource.data.role == resource.data.role
        && request.resource.data.department == resource.data.department
        && request.resource.data.email == resource.data.email
      );
      allow delete: if isSuper() || (isDeptAdminOf(resource.data.department) && resource.data.role == 'employee');
    }

    match /departments/{id} {
      allow read: if signedIn();
      allow write: if isSuper();
    }

    match /leaveRequests/{id} {
      allow read: if signedIn() && (
        resource.data.employeeId == request.auth.uid || isSuper() || isDeptAdminOf(resource.data.department)
      );
      allow create: if signedIn()
        && request.resource.data.employeeId == request.auth.uid
        && request.resource.data.status == 'pending';
      allow update: if isSuper() || isDeptAdminOf(resource.data.department);
      allow delete: if signedIn() && (
        (resource.data.employeeId == request.auth.uid && resource.data.status == 'pending') || isSuper()
      );
    }

    match /attendance/{id} {
      allow read: if signedIn() && (
        resource.data.employeeId == request.auth.uid || isSuper() || isDeptAdminOf(resource.data.department)
      );
      allow create: if signedIn() && request.resource.data.employeeId == request.auth.uid;
      allow update: if signedIn() && (
        resource.data.employeeId == request.auth.uid || isSuper() || isDeptAdminOf(resource.data.department)
      );
      allow delete: if isSuper();
    }
  }
}
Meka Firebase Console → Firestore Database → Rules tab eke thiyena ewa okkoma ain karala paste karala Publish karanna. Kalin thiyena rules backup ekak gannath amathaka karanna epa.

me system ekama picme transport bill add krnn etails ekakuth dala hadala denn puluvnda puluvan nm slip eka png add krnn option ekakuth oni giye kohe idn kothanatada km gana drive details mona project ekatada and amathara detail add krnn th oni professional vidiyta full ui ekama update krnn e vagema excerl sheet apita donlode kr gnn date range ekk denna

Edited a file, ran 7 commands, and 7 more tools
Full main.js eka aluthin, "Transport Bills" module ekak (PickMe/taxi bill claim) ekathu karala dala thiyenawa. Firebase config eka, kalin data structure eka, ewa okkoma wenas karala na.

Aluth features

Transport Bills (hama kenekutama) — "From", "To", Distance (km), Amount, Date, Project, Additional details, saha slip photo (PNG/JPG) upload. Slip eka thumbnail ekak vidiyata penenawa, click karama loku karala penenawa.
Projects (Super Admin ta witharai) — Department ekak vage, project ekak add karala, bill ekak ekata tag karanna puluwan.
Transport Claims (Admin ta) — okkoma bills balanna, filter karanna (status, date range, department, project), Approve/Reject karanna (Leave requests wage), amount total ekakuth penenawa.
Excel export — date range ekak dila — "From" / "To" date dala filter karala "Export to Excel" click kalama, e range ekatama wena bills witharak file ekata yanawa. File name ekath date range eka ekka enawa (transport-bills_2026-09-01_to_2026-09-22-....xlsx).
Dashboard saha Profile ekatath transport bill stats ekathu karala thiyenawa.
Kohomada image eka store wenne

Oyage Firebase project eke thiyena Storage (office-leave-system-67784.firebasestorage.app) ekatai slip photo eka upload wenne — transport-slips/{uid}/... kiyana path ekata. Firestore document eke thiyenne image eke URL eka witharai (Firestore ekata loku file danna baha).

Wadagath — dෙන්න one deka

Firestore Rules update karanna one (firestore.rules file eka dala thiyenawa) — projects saha transportBills collections dekakata rules ekathu karala thiyenawa. Console → Firestore Database → Rules → paste → Publish.
Storage Rules aluthen dana one (storage.rules file eka) — nathinam slip upload eka "Permission denied" enawa. Console → Storage → Rules tab → paste → Publish. (Storage ekath enable karala thiyenna one — nathinam Console eken "Get Started" karanna kියලා ahanawa.)
Meka dekama publish karala, aye page eka refresh karala "New transport bill" try karanna.


Main
JS 

Firestore
RULES 

Storage
RULES 
Firestore and storage ekecode venama denna copy krgnnpuluvn vidiyta

Firestore Rules —




Claude is AI and can make mistakes.
Main · JS
import firebase from "firebase/compat/app";
import "firebase/compat/auth";
import "firebase/compat/firestore";
import "firebase/compat/storage";
import * as XLSX from "xlsx";
import "./style.css";
 
/* ============================= CONFIG ============================= */
// Firebase Console → Project settings → General → Your apps → SDK setup and configuration.
const firebaseConfig = {
  apiKey: "AIzaSyCYkMcKrFCsqQkD4y6tGUJ1mMCyE6m_Y1M",
  authDomain: "office-leave-system-67784.firebaseapp.com",
  projectId: "office-leave-system-67784",
  storageBucket: "office-leave-system-67784.firebasestorage.app",
  messagingSenderId: "516750831628",
  appId: "1:516750831628:web:f65ce20fead2e1b0497e29"
};
const APP_NAME = "LeaveDesk";
const LEAVE_TYPES = ["Annual Leave", "Casual Leave", "Sick Leave", "Half Day", "No Pay Leave"];
 
firebase.initializeApp(firebaseConfig);
const auth = firebase.auth();
const db = firebase.firestore();
const storage = firebase.storage();
const MAX_SLIP_MB = 5;
 
/* ============================= EXTRA STYLES =============================
   Injected after style.css so the look is refined without touching style.css.
   Colours reuse the existing CSS variables. */
const EXTRA_CSS = `
:root {
  --ld-radius: 12px;
  --ld-shadow: 0 1px 2px rgba(16,24,40,.05), 0 1px 3px rgba(16,24,40,.06);
  --ld-shadow-lg: 0 24px 48px -12px rgba(16,24,40,.28);
  --ld-ring: 0 0 0 3px rgba(100,116,150,.22);
  --ld-safe-t: env(safe-area-inset-top, 0px);
  --ld-safe-b: env(safe-area-inset-bottom, 0px);
}
* { -webkit-tap-highlight-color: transparent; }
body { -webkit-font-smoothing: antialiased; -moz-osx-font-smoothing: grayscale; overflow-x: hidden; }
 
.surface { border-radius: var(--ld-radius); box-shadow: var(--ld-shadow); }
 
/* buttons + inputs */
.btn { transition: background-color .15s, border-color .15s, box-shadow .15s, opacity .15s; }
.btn:active:not(:disabled) { transform: translateY(1px); }
.btn:disabled { opacity: .55; cursor: not-allowed; }
.btn:focus-visible { outline: none; box-shadow: var(--ld-ring); }
.input:focus { outline: none; border-color: var(--navy); box-shadow: var(--ld-ring); }
textarea.input { resize: vertical; min-height: 64px; }
 
/* tables */
.table-wrap { overflow-x: auto; -webkit-overflow-scrolling: touch; }
.table-wrap table { width: 100%; border-collapse: collapse; }
.table-wrap thead th {
  text-align: left; font-size: 12.5px; font-weight: 600; color: var(--text-muted);
  padding: 12px 18px; white-space: nowrap; background: rgba(15,23,42,.03);
  border-bottom: 1px solid var(--border);
}
.table-wrap tbody td { padding: 14px 18px; font-size: 14px; vertical-align: middle; border-bottom: 1px solid var(--border); }
.table-wrap tbody tr:last-child td { border-bottom: 0; }
.table-wrap tbody tr:hover td { background: rgba(15,23,42,.02); }
 
/* status badges */
.badge { display: inline-flex; align-items: center; gap: 6px; padding: 3px 10px; border-radius: 999px; font-size: 12px; font-weight: 600; line-height: 1.5; }
.badge::before { content: ""; width: 6px; height: 6px; border-radius: 50%; background: currentColor; }
.badge-pending  { background: #FEF3C7; color: #92400E; }
.badge-approved { background: #DCFCE7; color: #166534; }
.badge-rejected { background: #FEE2E2; color: #991B1B; }
.badge-open     { background: #DBEAFE; color: #1E40AF; }
.badge-closed   { background: #E2E8F0; color: #334155; }
.badge-role { background: #EEF2FF; color: #3730A3; }
.badge-role::before { display: none; }
 
/* modal */
.modal-backdrop {
  position: fixed; inset: 0; z-index: 1000; display: flex; align-items: center; justify-content: center;
  padding: 16px; padding-top: calc(16px + var(--ld-safe-t)); padding-bottom: calc(16px + var(--ld-safe-b));
  background: rgba(15,23,42,.5); backdrop-filter: blur(3px); animation: ld-fade .15s ease-out;
}
.modal-panel {
  width: 100%; max-width: 480px; max-height: 100%; overflow: auto;
  background: var(--surface, #fff); border-radius: 16px; box-shadow: var(--ld-shadow-lg); animation: ld-pop .18s ease-out;
}
@keyframes ld-fade { from { opacity: 0; } to { opacity: 1; } }
@keyframes ld-pop  { from { opacity: 0; transform: translateY(8px) scale(.985); } to { opacity: 1; transform: none; } }
 
/* toasts — always above the modal */
#toast-root {
  position: fixed; right: 12px; left: 12px; bottom: calc(12px + var(--ld-safe-b)); top: auto; z-index: 9999;
  display: flex; flex-direction: column; gap: 8px; align-items: flex-end; pointer-events: none;
}
.toast {
  position: static; inset: auto; transform: none; margin: 0; pointer-events: auto; max-width: 380px;
  background: #1F2937; color: #fff; padding: 11px 16px; border-radius: 10px; font-size: 14px;
  border-left: 4px solid #94A3B8; box-shadow: var(--ld-shadow-lg);
}
.toast.ok  { border-left-color: #22C55E; }
.toast.err { border-left-color: #EF4444; }
 
/* sidebar */
.app-sidebar { transition: transform .22s ease; }
.sidebar-link {
  display: flex; align-items: center; gap: 10px; padding: 11px 12px; border-radius: 10px;
  font-size: 14px; font-weight: 500; cursor: pointer; color: #B9C2D6; transition: background .15s, color .15s;
}
.sidebar-link:hover { background: rgba(255,255,255,.07); color: #fff; }
.sidebar-link.active { background: rgba(255,255,255,.14); color: #fff; }
.nav-count {
  margin-left: auto; min-width: 20px; height: 20px; padding: 0 6px; border-radius: 999px;
  background: #F59E0B; color: #1F2937; font-size: 11px; font-weight: 700;
  display: inline-flex; align-items: center; justify-content: center;
}
.sidebar-scrim { position: fixed; inset: 0; z-index: 40; background: rgba(15,23,42,.5); }
 
/* dashboard + blocks */
.stat { position: relative; overflow: hidden; padding: 18px 20px 18px 24px; }
.stat::before { content: ""; position: absolute; left: 0; top: 0; bottom: 0; width: 4px; background: var(--stat, #94A3B8); }
.avatar {
  width: 30px; height: 30px; border-radius: 50%; flex-shrink: 0; background: #E0E7FF; color: #3730A3;
  font-size: 12px; font-weight: 700; display: inline-flex; align-items: center; justify-content: center;
}
.avatar-lg { width: 60px; height: 60px; font-size: 20px; }
.cell-user { display: flex; align-items: center; gap: 10px; }
.empty { padding: 40px 16px; text-align: center; color: var(--text-muted); font-size: 14px; }
.attn-item { display: flex; align-items: center; justify-content: space-between; gap: 12px; padding: 12px 20px; }
.attn-item + .attn-item { border-top: 1px solid var(--border); }
.chip-count { margin-left: 6px; font-size: 11.5px; opacity: .75; }
.detail-grid { display: grid; grid-template-columns: 1fr 1fr; gap: 14px 16px; margin-bottom: 16px; }
.detail-grid .k, .k { font-size: 12.5px; color: var(--text-muted); margin-bottom: 2px; }
.detail-grid .v { font-size: 14px; font-weight: 500; }
.form-err { color: var(--danger); font-size: 14px; min-height: 20px; margin-bottom: 8px; }
.kv-row { display: flex; align-items: center; justify-content: space-between; gap: 16px; padding: 13px 20px; font-size: 14px; }
.kv-row + .kv-row { border-top: 1px solid var(--border); }
.kv-row .kv-k { color: var(--text-muted); }
.kv-row .kv-v { font-weight: 500; text-align: right; word-break: break-word; }
 
/* field check-in panel */
.punch-card { padding: 22px 20px; text-align: center; }
.punch-clock { font-size: 40px; font-weight: 700; letter-spacing: -.02em; color: var(--navy); line-height: 1.1; }
.punch-sub { font-size: 13.5px; color: var(--text-muted); margin-top: 4px; }
.punch-btn { width: 100%; max-width: 320px; margin: 18px auto 0; display: block; padding: 15px 20px; font-size: 16px; font-weight: 600; border-radius: 12px; }
.live-dot { display: inline-block; width: 8px; height: 8px; border-radius: 50%; background: #22C55E; margin-right: 7px; animation: ld-pulse 1.8s infinite; }
@keyframes ld-pulse { 0%,100% { opacity: 1; } 50% { opacity: .3; } }
.map-link { color: var(--navy); text-decoration: underline; font-size: 13px; }
 
/* transport bills */
.upload-box {
  border: 1.5px dashed var(--border); border-radius: 12px; padding: 18px; text-align: center;
  cursor: pointer; transition: border-color .15s, background .15s;
}
.upload-box:hover { border-color: var(--navy); background: rgba(15,23,42,.02); }
.upload-box input[type="file"] { display: none; }
.upload-box .icn { color: var(--text-muted); margin-bottom: 6px; }
.slip-thumb {
  width: 44px; height: 44px; border-radius: 8px; object-fit: cover; border: 1px solid var(--border);
  cursor: pointer; display: block;
}
.slip-preview { max-width: 100%; max-height: 220px; border-radius: 10px; border: 1px solid var(--border); display: block; margin: 0 auto; }
.amount { font-variant-numeric: tabular-nums; font-weight: 600; }
.route-arrow { color: var(--text-muted); margin: 0 4px; }
.lightbox {
  position: fixed; inset: 0; z-index: 1200; background: rgba(15,23,42,.85);
  display: flex; align-items: center; justify-content: center; padding: 24px; cursor: zoom-out;
}
.lightbox img { max-width: 100%; max-height: 100%; border-radius: 8px; box-shadow: var(--ld-shadow-lg); }
.progress-bar { height: 6px; border-radius: 999px; background: var(--border); overflow: hidden; margin-top: 10px; }
.progress-bar > div { height: 100%; background: var(--navy); transition: width .2s; }
 
/* ---------- responsive ---------- */
@media (max-width: 767px) {
  .app-sidebar {
    position: fixed; z-index: 50; top: 0; bottom: 0; left: 0; width: 260px;
    transform: translateX(-100%);
    padding-top: calc(16px + var(--ld-safe-t)); padding-bottom: calc(16px + var(--ld-safe-b));
    overflow-y: auto;
  }
  .app-sidebar.open { transform: translateX(0); box-shadow: var(--ld-shadow-lg); }
  .topbar { position: sticky; top: 0; z-index: 30; padding-top: calc(12px + var(--ld-safe-t)); }
  .page-head { flex-direction: column; align-items: stretch; }
  .page-head .head-actions { display: flex; gap: 8px; }
  .page-head .head-actions .btn { flex: 1; }
  .detail-grid { grid-template-columns: 1fr; }
  .punch-clock { font-size: 34px; }
  .modal-panel { max-width: 100%; }
  .filter-row { overflow-x: auto; flex-wrap: nowrap; padding-bottom: 4px; }
  .filter-row .btn, .filter-row select, .filter-row > div { flex: 0 0 auto; }
 
  /* tables become stacked cards */
  .rtable thead { display: none; }
  .rtable, .rtable tbody, .rtable tr, .rtable td { display: block; width: 100%; }
  .rtable tr { border: 1px solid var(--border); border-radius: 12px; margin: 12px; padding: 4px 0; }
  .rtable tbody td, .rtable tbody tr:hover td {
    border: 0; padding: 9px 14px; text-align: right; background: transparent;
    display: flex; align-items: center; justify-content: space-between; gap: 14px; max-width: none !important;
  }
  .rtable tbody td::before {
    content: attr(data-label); font-size: 12.5px; font-weight: 600; color: var(--text-muted);
    text-align: left; flex: 0 0 auto;
  }
  .rtable tbody td.td-empty { display: none; }
  .rtable tbody td.td-action { justify-content: flex-end; padding-top: 4px; padding-bottom: 12px; }
  .rtable tbody td.td-action::before { content: ""; }
}
 
@media (prefers-reduced-motion: reduce) {
  .modal-backdrop, .modal-panel, .app-sidebar, .live-dot { animation: none; transition: none; }
  .btn, .sidebar-link { transition: none; }
}
`;
function injectExtraCSS() {
  if (!document.getElementById("leavedesk-extra-css")) {
    const s = document.createElement("style");
    s.id = "leavedesk-extra-css";
    s.textContent = EXTRA_CSS;
    document.head.appendChild(s);
  }
  let vp = document.querySelector('meta[name="viewport"]');
  if (!vp) { vp = document.createElement("meta"); vp.name = "viewport"; document.head.appendChild(vp); }
  vp.setAttribute("content", "width=device-width, initial-scale=1, viewport-fit=cover");
}
 
/* ============================= STATE ============================= */
const state = {
  booting: true,
  setupNeeded: false,
  creatingAdmin: false,  // true while the first Super Admin is being created (pauses the auth listener)
  user: null,            // { uid, name, email, role, department, phone, designation }
  view: "login",         // login | setup | app
  tab: "dashboard",
  sidebarOpen: false,
  departments: [],
  projects: [],
  employees: [],
  leaves: [],
  attendance: [],
  bills: [],
  leaveFilter: { status: "all", dept: "all" },
  attFilter: { dept: "all", from: "", to: "" },
  billFilter: { status: "all", dept: "all", project: "all", from: "", to: "" },
  unsubs: [],
};
let reviewBusy = false;
let punchBusy = false;
let billReviewBusy = false;
let billSubmitBusy = false;
let clockTimer = null;
let pendingSlipFile = null; // File object selected in the "new bill" modal, before upload
 
/* ============================= HELPERS ============================= */
function $(id) { return document.getElementById(id); }
function esc(s) {
  return String(s ?? "").replace(/[&<>"']/g, c => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
}
function fmtDate(iso) {
  if (!iso) return "—";
  const d = new Date(iso + "T00:00:00");
  return d.toLocaleDateString(undefined, { year: "numeric", month: "short", day: "numeric" });
}
function fmtRange(l) {
  return fmtDate(l.startDate) + (l.startDate !== l.endDate ? " → " + fmtDate(l.endDate) : "");
}
function fmtDateTime(iso) {
  if (!iso) return "—";
  const d = new Date(iso);
  if (isNaN(d)) return "—";
  return d.toLocaleString(undefined, { year: "numeric", month: "short", day: "numeric", hour: "2-digit", minute: "2-digit" });
}
function fmtTime(iso) {
  if (!iso) return "—";
  const d = new Date(iso);
  if (isNaN(d)) return "—";
  return d.toLocaleTimeString(undefined, { hour: "2-digit", minute: "2-digit" });
}
function fmtDuration(mins) {
  if (mins == null || isNaN(mins)) return "—";
  const m = Math.max(0, Math.round(mins));
  const h = Math.floor(m / 60), r = m % 60;
  return h ? `${h}h ${r}m` : `${r}m`;
}
// Local date (not UTC), so the default date is correct in every timezone
function todayISO() {
  const d = new Date();
  const m = String(d.getMonth() + 1).padStart(2, "0");
  const day = String(d.getDate()).padStart(2, "0");
  return `${d.getFullYear()}-${m}-${day}`;
}
function daysBetween(start, end, half) {
  if (half) return 0.5;
  const a = new Date(start + "T00:00:00"), b = new Date(end + "T00:00:00");
  const diff = Math.round((b - a) / 86400000) + 1;
  return diff > 0 ? diff : 0;
}
function deptName(id) {
  const d = state.departments.find(x => x.id === id);
  return d ? d.name : (id ? "—" : "All departments");
}
function projectName(id) {
  const p = state.projects.find(x => x.id === id);
  return p ? p.name : (id ? "—" : "No project");
}
function fmtMoney(n) {
  const v = Number(n) || 0;
  return v.toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}
function empByUid(uid) { return state.employees.find(e => e.id === uid); }
function initials(name) {
  const parts = String(name || "?").trim().split(/\s+/).filter(Boolean);
  const a = (parts[0] || "?")[0], b = parts.length > 1 ? parts[parts.length - 1][0] : "";
  return (a + b).toUpperCase();
}
function toast(msg, type) {
  const el = document.createElement("div");
  el.className = "toast" + (type === "err" ? " err" : type === "ok" ? " ok" : "");
  el.textContent = msg;
  $("toast-root").appendChild(el);
  setTimeout(() => { el.style.opacity = "0"; el.style.transition = "opacity .3s"; setTimeout(() => el.remove(), 300); }, 3600);
}
function isSuperAdmin() { return state.user && state.user.role === "superadmin"; }
function isDeptAdmin() { return state.user && state.user.role === "deptadmin"; }
function isAdmin() { return isSuperAdmin() || isDeptAdmin(); }
function roleLabel(r) { return r === "superadmin" ? "Super Admin" : r === "deptadmin" ? "Department Admin" : "Employee"; }
 
/* Leaves an admin may see (Super Admin: all, Dept Admin: own department) */
function adminScopeLeaves() {
  return isSuperAdmin() ? state.leaves : state.leaves.filter(l => l.department === state.user.department);
}
function adminBaseList() {
  let list = adminScopeLeaves();
  if (isSuperAdmin() && state.leaveFilter.dept !== "all") list = list.filter(l => l.department === state.leaveFilter.dept);
  return list;
}
function adminFilteredList() {
  let list = adminBaseList();
  if (state.leaveFilter.status !== "all") list = list.filter(l => l.status === state.leaveFilter.status);
  return list;
}
 
/* Transport bills an admin may see (Super Admin: all, Dept Admin: own department) */
function adminScopeBills() {
  return isSuperAdmin() ? state.bills : state.bills.filter(b => b.department === state.user.department);
}
function billFilteredList() {
  let list = adminScopeBills();
  if (isSuperAdmin() && state.billFilter.dept !== "all") list = list.filter(b => b.department === state.billFilter.dept);
  if (state.billFilter.project !== "all") list = list.filter(b => b.project === state.billFilter.project);
  if (state.billFilter.status !== "all") list = list.filter(b => b.status === state.billFilter.status);
  if (state.billFilter.from) list = list.filter(b => String(b.date || "") >= state.billFilter.from);
  if (state.billFilter.to) list = list.filter(b => String(b.date || "") <= state.billFilter.to);
  return list;
}
 
/* ============================= MODAL ============================= */
function openModal(html) {
  $("modal-root").innerHTML = `
    <div class="modal-backdrop" onmousedown="if(event.target===this) closeModal()">
      <div class="modal-panel">${html}</div>
    </div>`;
}
function closeModal() { $("modal-root").innerHTML = ""; }
 
function viewSlip(url) {
  if (!url) return;
  const el = document.createElement("div");
  el.className = "lightbox";
  el.onclick = () => el.remove();
  el.innerHTML = `<img src="${esc(url)}" alt="Transport bill slip">`;
  document.body.appendChild(el);
}
 
/* ============================= RENDER DISPATCH ============================= */
function render() {
  if (state.booting) return renderShellless(`
    <div class="min-h-screen flex items-center justify-center" style="background:var(--bg)">
      <div class="text-center">
        <div class="brand text-2xl mb-2" style="color:var(--navy)">${APP_NAME}</div>
        <div style="color:var(--text-muted)" class="text-sm">Loading…</div>
      </div>
    </div>`);
 
  if (state.view === "setup") return renderSetup();
  if (state.view === "login") return renderLogin();
  return renderApp();
}
function renderShellless(html) { $("main-root").innerHTML = html; }
 
/* ============================= SETUP SCREEN ============================= */
function renderSetup() {
  renderShellless(`
  <div class="min-h-screen grid md:grid-cols-2" style="background:var(--bg)">
    <div class="hidden md:flex flex-col justify-between p-12" style="background:var(--navy); color:#fff;">
      <div class="brand text-3xl">${APP_NAME}</div>
      <div>
        <div class="brand text-4xl leading-tight mb-4" style="max-width:380px">First, let's set up your organization.</div>
        <p style="color:#B9C2D6; max-width:360px; font-size:15px;">This creates the very first Super Admin account. From there you can add departments, employees, and start approving leave.</p>
      </div>
      <div style="color:#7C89A3; font-size:13px;">One-time setup · Runs on your own Firebase project</div>
    </div>
    <div class="flex items-center justify-center p-6">
      <form onsubmit="handleSetupSubmit(event)" class="w-full" style="max-width:380px">
        <div class="brand text-2xl mb-1 md:hidden" style="color:var(--navy)">${APP_NAME}</div>
        <h1 class="text-xl font-semibold mb-1">Create your Super Admin</h1>
        <p class="text-sm mb-6" style="color:var(--text-muted)">You'll use this account to manage the whole system.</p>
 
        <label class="label">Organization name</label>
        <input class="input mb-4" name="org" placeholder="e.g. Batapola Traders (Pvt) Ltd" required>
 
        <label class="label">Your full name</label>
        <input class="input mb-4" name="name" placeholder="e.g. Gimshan Perera" required>
 
        <label class="label">Work email</label>
        <input class="input mb-4" name="email" type="email" placeholder="admin@company.com" required>
 
        <label class="label">Password</label>
        <input class="input mb-1" name="password" type="password" placeholder="At least 6 characters" minlength="6" required>
        <p class="text-xs mb-5" style="color:var(--text-muted)">Minimum 6 characters.</p>
 
        <button class="btn btn-dark w-full" type="submit" id="setup-btn">Create Super Admin & continue</button>
        <p id="setup-err" class="text-sm mt-3" style="color:var(--danger)"></p>
      </form>
    </div>
  </div>`);
}
 
async function handleSetupSubmit(e) {
  e.preventDefault();
  const f = e.target, btn = $("setup-btn"), err = $("setup-err");
  err.textContent = "";
  const org = f.org.value.trim(), name = f.name.value.trim(), email = f.email.value.trim(), password = f.password.value;
  btn.disabled = true; btn.textContent = "Creating…";
 
  // Pause the auth listener: it would otherwise fire as soon as the account is created,
  // before the employee profile below is written, and wrongly sign the new admin out.
  state.creatingAdmin = true;
  try {
    const cred = await auth.createUserWithEmailAndPassword(email, password);
    const uid = cred.user.uid;
    await db.collection("employees").doc(uid).set({
      name, email, role: "superadmin", department: null,
      createdAt: firebase.firestore.FieldValue.serverTimestamp(),
    });
    await db.collection("system").doc("status").set({
      setupDone: true, orgName: org, createdAt: firebase.firestore.FieldValue.serverTimestamp(),
    });
    state.setupNeeded = false;
    state.creatingAdmin = false;
    toast("Super Admin created. Welcome!", "ok");
    await loadUser(auth.currentUser);
  } catch (ex) {
    console.error(ex);
    const msg = friendlyAuthError(ex);
    // Roll back a half-created login so the setup can be retried with the same email
    if (auth.currentUser) { try { await auth.currentUser.delete(); } catch (_) {} }
    state.creatingAdmin = false;
    state.setupNeeded = true;
    state.view = "setup";
    render();
    toast(msg, "err");
    const e2 = $("setup-err"); if (e2) e2.textContent = msg;
  }
}
 
/* ============================= LOGIN SCREEN ============================= */
function renderLogin() {
  renderShellless(`
  <div class="min-h-screen grid md:grid-cols-2" style="background:var(--bg)">
    <div class="hidden md:flex flex-col justify-between p-12" style="background:var(--navy); color:#fff;">
      <div class="brand text-3xl">${APP_NAME}</div>
      <div>
        <div class="brand text-4xl leading-tight mb-4" style="max-width:380px">Leave, attendance, and field work — in one place.</div>
        <p style="color:#B9C2D6; max-width:360px; font-size:15px;">Submit leave, check in and out from the field, and — for admins — manage departments, staff, and approvals across the office.</p>
      </div>
      <div style="color:#7C89A3; font-size:13px;">Sign in with the email your admin set up for you</div>
    </div>
    <div class="flex items-center justify-center p-6">
      <form onsubmit="handleLogin(event)" class="w-full" style="max-width:360px">
        <div class="brand text-2xl mb-6 md:hidden" style="color:var(--navy)">${APP_NAME}</div>
        <h1 class="text-xl font-semibold mb-1">Sign in</h1>
        <p class="text-sm mb-6" style="color:var(--text-muted)">Enter your work email and password.</p>
 
        <label class="label">Email</label>
        <input class="input mb-4" name="email" type="email" placeholder="you@company.com" required autofocus>
 
        <label class="label">Password</label>
        <input class="input mb-5" name="password" type="password" placeholder="••••••••" required>
 
        <button class="btn btn-dark w-full" type="submit" id="login-btn">Sign in</button>
        <p id="login-err" class="text-sm mt-3" style="color:var(--danger)"></p>
      </form>
    </div>
  </div>`);
}
 
async function handleLogin(e) {
  e.preventDefault();
  const f = e.target, btn = $("login-btn"), err = $("login-err");
  err.textContent = "";
  btn.disabled = true; btn.textContent = "Signing in…";
  try {
    await auth.signInWithEmailAndPassword(f.email.value.trim(), f.password.value);
  } catch (ex) {
    console.error(ex);
    err.textContent = friendlyAuthError(ex);
    btn.disabled = false; btn.textContent = "Sign in";
  }
}
 
function friendlyAuthError(ex) {
  const code = ex && ex.code || "";
  if (code.includes("email-already-in-use")) return "That email is already registered.";
  if (code.includes("invalid-email")) return "That email address looks invalid.";
  if (code.includes("weak-password")) return "Password is too weak (min 6 characters).";
  if (code.includes("user-not-found") || code.includes("wrong-password") || code.includes("invalid-credential")) return "Incorrect email or password.";
  if (code.includes("too-many-requests")) return "Too many attempts. Please wait a moment and try again.";
  if (code.includes("permission-denied")) return "Permission denied — check your Firestore security rules.";
  if (code.includes("requires-recent-login")) return "Please sign out and sign in again, then retry.";
  return (ex && ex.message) || "Something went wrong. Please try again.";
}
 
async function handleLogout() {
  detachListeners();
  await auth.signOut();
  state.tab = "dashboard";
  state.sidebarOpen = false;
}
 
/* ============================= APP SHELL ============================= */
const NAV = [
  { id: "dashboard", label: "Dashboard", roles: ["superadmin", "deptadmin", "employee"], icon: "grid" },
  { id: "field", label: "Field Check-in", roles: ["superadmin", "deptadmin", "employee"], icon: "pin" },
  { id: "myleave", label: "My Leave", roles: ["superadmin", "deptadmin", "employee"], icon: "calendar" },
  { id: "leaves", label: "Leave Requests", roles: ["superadmin", "deptadmin"], icon: "inbox" },
  { id: "mybills", label: "Transport Bills", roles: ["superadmin", "deptadmin", "employee"], icon: "car" },
  { id: "attendance", label: "Attendance", roles: ["superadmin", "deptadmin"], icon: "clock" },
  { id: "billsadmin", label: "Transport Claims", roles: ["superadmin", "deptadmin"], icon: "receipt" },
  { id: "employees", label: "Employees", roles: ["superadmin", "deptadmin"], icon: "users" },
  { id: "departments", label: "Departments", roles: ["superadmin"], icon: "building" },
  { id: "projects", label: "Projects", roles: ["superadmin"], icon: "folder" },
  { id: "profile", label: "My Profile", roles: ["superadmin", "deptadmin", "employee"], icon: "user" },
];
const ICONS = {
  grid: '<svg width="17" height="17" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><rect x="3" y="3" width="8" height="8" rx="1.5"/><rect x="13" y="3" width="8" height="8" rx="1.5"/><rect x="3" y="13" width="8" height="8" rx="1.5"/><rect x="13" y="13" width="8" height="8" rx="1.5"/></svg>',
  calendar: '<svg width="17" height="17" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><rect x="3" y="5" width="18" height="16" rx="2"/><path d="M16 3v4M8 3v4M3 10h18"/></svg>',
  inbox: '<svg width="17" height="17" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M3 12h4l2 3h6l2-3h4"/><path d="M5 5h14l2 7v7a2 2 0 01-2 2H5a2 2 0 01-2-2v-7z"/></svg>',
  users: '<svg width="17" height="17" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><circle cx="9" cy="8" r="3.2"/><path d="M2.5 20c0-3.6 2.9-6.2 6.5-6.2s6.5 2.6 6.5 6.2"/><circle cx="17.2" cy="8.5" r="2.6"/><path d="M15.5 13.9c2.8.3 4.9 2.6 5 5.8"/></svg>',
  building: '<svg width="17" height="17" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><rect x="4" y="3" width="16" height="18" rx="1"/><path d="M9 8h1M14 8h1M9 12h1M14 12h1M9 16h1M14 16h1"/></svg>',
  pin: '<svg width="17" height="17" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M12 21s7-5.3 7-11a7 7 0 10-14 0c0 5.7 7 11 7 11z"/><circle cx="12" cy="10" r="2.6"/></svg>',
  clock: '<svg width="17" height="17" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><circle cx="12" cy="12" r="9"/><path d="M12 7v5l3.2 2"/></svg>',
  user: '<svg width="17" height="17" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><circle cx="12" cy="8" r="3.6"/><path d="M4.5 20c0-4 3.4-6.8 7.5-6.8s7.5 2.8 7.5 6.8"/></svg>',
  car: '<svg width="17" height="17" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M4 16V11l1.8-4.5A2 2 0 017.7 5h8.6a2 2 0 011.9 1.5L20 11v5"/><path d="M4 16h16v2.5a1 1 0 01-1 1h-1.5a1 1 0 01-1-1V17h-9v1.5a1 1 0 01-1 1H5a1 1 0 01-1-1V16z"/><circle cx="7.5" cy="16" r="1.4"/><circle cx="16.5" cy="16" r="1.4"/></svg>',
  receipt: '<svg width="17" height="17" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M6 3h12v18l-2.5-1.6L13 21l-2.5-1.6L8 21l-2-1.6z"/><path d="M9 8h6M9 12h6M9 16h3"/></svg>',
  folder: '<svg width="17" height="17" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M3 6.5a1.5 1.5 0 011.5-1.5H9l2 2.2h8a1.5 1.5 0 011.5 1.5v9.3A1.5 1.5 0 0119 19.5H4.5A1.5 1.5 0 013 18z"/></svg>',
};
 
function renderApp() {
  const items = NAV.filter(n => n.roles.includes(state.user.role));
  const pendingCount = isAdmin() ? adminScopeLeaves().filter(l => l.status === "pending").length : 0;
  const pendingBillCount = isAdmin() ? adminScopeBills().filter(b => b.status === "pending").length : 0;
  const open = openSessionFor(state.user.uid);
 
  renderShellless(`
    ${state.sidebarOpen ? `<div class="sidebar-scrim" onclick="state.sidebarOpen=false; render();"></div>` : ""}
    <div class="min-h-screen flex" style="background:var(--bg)">
      <aside class="app-sidebar ${state.sidebarOpen ? "open" : ""} w-[250px] shrink-0 flex flex-col justify-between p-4" style="background:var(--navy);">
        <div>
          <div class="flex items-center justify-between px-2 mb-6 mt-1">
            <div class="brand text-xl" style="color:#fff">${APP_NAME}</div>
            <button class="md:hidden btn btn-ghost btn-sm" style="color:#C7CFE0" onclick="state.sidebarOpen=false; render();">✕</button>
          </div>
          <nav class="flex flex-col gap-1">
            ${items.map(n => `
              <div class="sidebar-link ${state.tab === n.id ? "active" : ""}" onclick="switchTab('${n.id}')">
                ${ICONS[n.icon]}<span>${n.label}</span>
                ${n.id === "leaves" && pendingCount > 0 ? `<span class="nav-count">${pendingCount}</span>` : ""}
                ${n.id === "billsadmin" && pendingBillCount > 0 ? `<span class="nav-count">${pendingBillCount}</span>` : ""}
                ${n.id === "field" && open ? `<span class="nav-count" style="background:#22C55E;color:#062E14">●</span>` : ""}
              </div>`).join("")}
          </nav>
        </div>
        <div class="px-2 pb-1">
          <div class="flex items-center gap-2 mb-3 pt-3" style="border-top:1px solid rgba(255,255,255,0.1)">
            <div class="flex items-center justify-center rounded-full text-xs font-semibold shrink-0" style="width:32px;height:32px;background:rgba(255,255,255,0.12); color:#fff;">${esc(initials(state.user.name))}</div>
            <div class="min-w-0">
              <div class="text-sm font-medium truncate" style="color:#fff">${esc(state.user.name)}</div>
              <div class="text-xs truncate" style="color:#8D98B0">${roleLabel(state.user.role)}</div>
            </div>
          </div>
          <button class="btn btn-sm w-full" style="background:rgba(255,255,255,0.08); color:#EDEFF4;" onclick="switchTab('profile')">My profile</button>
          <button class="btn btn-sm btn-ghost w-full mt-1.5" style="color:#9AA5BC" onclick="handleLogout()">Sign out</button>
        </div>
      </aside>
 
      <div class="flex-1 min-w-0">
        <div class="topbar md:hidden flex items-center gap-3 p-3" style="border-bottom:1px solid var(--border); background:var(--surface);">
          <button class="btn btn-secondary btn-sm" onclick="state.sidebarOpen=true; render();" aria-label="Open menu">☰</button>
          <div class="brand text-lg" style="color:var(--navy)">${APP_NAME}</div>
          <button class="btn btn-sm ${open ? "btn-danger" : "btn-primary"} ml-auto" onclick="switchTab('field')">${open ? "Check out" : "Check in"}</button>
        </div>
        <main class="p-4 md:p-8 max-w-[1180px]" id="tab-content">
          ${renderTabContent()}
        </main>
      </div>
    </div>
  `);
  startClock();
}
 
function switchTab(id) { state.tab = id; state.sidebarOpen = false; render(); }
 
function renderTabContent() {
  if (state.tab === "dashboard") return tplDashboard();
  if (state.tab === "field") return tplField();
  if (state.tab === "myleave") return tplMyLeave();
  if (state.tab === "leaves") return isAdmin() ? tplLeaves() : tplNoAccess();
  if (state.tab === "mybills") return tplMyBills();
  if (state.tab === "billsadmin") return isAdmin() ? tplBillsAdmin() : tplNoAccess();
  if (state.tab === "attendance") return isAdmin() ? tplAttendance() : tplNoAccess();
  if (state.tab === "employees") return isAdmin() ? tplEmployees() : tplNoAccess();
  if (state.tab === "departments") return isSuperAdmin() ? tplDepartments() : tplNoAccess();
  if (state.tab === "projects") return isSuperAdmin() ? tplProjects() : tplNoAccess();
  if (state.tab === "profile") return tplProfile();
  return "";
}
function tplNoAccess() { return `<div class="surface p-8 text-center" style="color:var(--text-muted)">You don't have access to this section.</div>`; }
 
function pageHeader(title, sub, action) {
  return `<div class="page-head flex items-start justify-between gap-4 mb-6 flex-wrap">
    <div>
      <h1 class="brand text-[26px]" style="color:var(--navy)">${title}</h1>
      ${sub ? `<p class="text-sm mt-1" style="color:var(--text-muted)">${sub}</p>` : ""}
    </div>
    ${action ? `<div class="head-actions">${action}</div>` : ""}
  </div>`;
}
 
/* ============================= DASHBOARD ============================= */
function tplDashboard() {
  const myLeaves = state.leaves.filter(l => l.employeeId === state.user.uid);
  const visibleLeaves = isAdmin() ? adminScopeLeaves() : myLeaves;
  const pending = visibleLeaves.filter(l => l.status === "pending").length;
  const approved = visibleLeaves.filter(l => l.status === "approved").length;
  const rejected = visibleLeaves.filter(l => l.status === "rejected").length;
  const myPending = myLeaves.filter(l => l.status === "pending").length;
  const myApprovedDays = myLeaves.filter(l => l.status === "approved").reduce((s, l) => s + (l.totalDays || 0), 0);
  const outNow = isAdmin() ? attendanceScope().filter(a => a.open).length : 0;
  const myBills = myBillsScope();
  const pendingBills = isAdmin() ? adminScopeBills().filter(b => b.status === "pending").length : 0;
  const myPendingBillAmt = myBills.filter(b => b.status === "pending").reduce((s, b) => s + (Number(b.amount) || 0), 0);
 
  const cards = isAdmin()
    ? [
        ["Pending leave review", pending, "warn"],
        ["Pending transport claims", pendingBills, "warn"],
        ["In the field now", outNow, "success"],
        ["Approved leave", approved, "muted"],
        ["Rejected leave", rejected, "danger"],
      ]
    : [
        ["My pending requests", myPending, "warn"],
        ["Approved days taken", myApprovedDays, "success"],
        ["Transport claims pending", fmtMoney(myPendingBillAmt), "muted"],
        ["Field time this month", fmtDuration(myMonthMinutes()), "muted"],
      ];
  const toneColor = t => t === "warn" ? "var(--warn)" : t === "success" ? "var(--success)" : t === "danger" ? "var(--danger)" : "var(--navy)";
 
  const recent = [...(isAdmin() ? visibleLeaves : myLeaves)]
    .sort((a, b) => (b.appliedAt || "").localeCompare(a.appliedAt || "")).slice(0, 6);
  const attention = isAdmin()
    ? visibleLeaves.filter(l => l.status === "pending").sort((a, b) => (a.appliedAt || "").localeCompare(b.appliedAt || "")).slice(0, 5)
    : [];
  const open = openSessionFor(state.user.uid);
 
  return `
    ${pageHeader("Dashboard", `Welcome back, ${esc(String(state.user.name).split(" ")[0])}.`)}
 
    <div class="grid gap-4 mb-6" style="grid-template-columns:repeat(auto-fit,minmax(160px,1fr))">
      ${cards.map(([label, val, tone]) => `
        <div class="surface stat" style="--stat:${toneColor(tone)}">
          <div class="text-sm mb-2" style="color:var(--text-muted)">${label}</div>
          <div class="text-3xl font-semibold" style="color:${toneColor(tone)}">${val}</div>
        </div>`).join("")}
    </div>
 
    <div class="surface p-5 mb-6 flex items-center justify-between gap-4 flex-wrap">
      <div>
        <div class="font-semibold text-sm mb-1">${open ? "You're checked in" : "Field check-in"}</div>
        <div class="text-sm" style="color:var(--text-muted)">
          ${open ? `${open.checkInNote ? esc(open.checkInNote) + ", since " : "Since "}${fmtTime(open.checkInAt)}` : "Record your time when you start work outside the office."}
        </div>
      </div>
      <button class="btn ${open ? "btn-danger" : "btn-primary"}" onclick="switchTab('field')">${open ? "Check out" : "Check in"}</button>
    </div>
 
    ${attention.length ? `
    <div class="surface overflow-hidden mb-6">
      <div class="px-5 py-4 flex items-center justify-between gap-3" style="border-bottom:1px solid var(--border)">
        <div class="font-semibold text-sm">Waiting for your decision</div>
        <button class="btn btn-sm btn-ghost" onclick="switchTab('leaves')">See all</button>
      </div>
      ${attention.map(l => `
        <div class="attn-item">
          <div class="cell-user min-w-0">
            <div class="avatar">${esc(initials(l.employeeName))}</div>
            <div class="min-w-0">
              <div class="text-sm font-medium truncate">${esc(l.employeeName)}</div>
              <div class="text-xs" style="color:var(--text-muted)">${esc(l.leaveType)}, ${fmtRange(l)} (${l.totalDays} day${l.totalDays === 1 ? "" : "s"})</div>
            </div>
          </div>
          <button class="btn btn-sm btn-primary shrink-0" onclick="modalReview('${l.id}')">Review</button>
        </div>`).join("")}
    </div>` : ""}
 
    <div class="surface overflow-hidden">
      <div class="px-5 py-4" style="border-bottom:1px solid var(--border)">
        <div class="font-semibold text-sm">${isAdmin() ? "Recent leave requests" : "My recent requests"}</div>
      </div>
      ${recent.length === 0 ? `<div class="empty">Nothing here yet.</div>` : `
      <div class="table-wrap"><table class="rtable">
        <thead><tr><th>Employee</th><th>Type</th><th>Dates</th><th>Days</th><th>Status</th></tr></thead>
        <tbody>
        ${recent.map(l => `
          <tr>
            <td data-label="Employee"><div class="cell-user"><div class="avatar">${esc(initials(l.employeeName))}</div><span>${esc(l.employeeName)}</span></div></td>
            <td data-label="Type">${esc(l.leaveType)}</td>
            <td data-label="Dates">${fmtRange(l)}</td>
            <td data-label="Days">${l.totalDays}</td>
            <td data-label="Status">${statusBadge(l.status)}</td>
          </tr>`).join("")}
        </tbody>
      </table></div>`}
    </div>
  `;
}
function statusBadge(s) {
  const st = s || "pending";
  const label = st.charAt(0).toUpperCase() + st.slice(1);
  return `<span class="badge badge-${st}">${label}</span>`;
}
function openBadge() { return `<span class="badge badge-open">Open</span>`; }
 
/* ============================= FIELD CHECK-IN / OUT ============================= */
function openSessionFor(uid) {
  return state.attendance.find(a => a.employeeId === uid && a.open === true) || null;
}
function attendanceScope() {
  if (isSuperAdmin()) return state.attendance;
  if (isDeptAdmin()) return state.attendance.filter(a => a.department === state.user.department);
  return state.attendance.filter(a => a.employeeId === state.user.uid);
}
function myMonthMinutes() {
  const p = todayISO().slice(0, 7);
  return state.attendance
    .filter(a => a.employeeId === state.user.uid && String(a.date || "").startsWith(p))
    .reduce((s, a) => s + (a.durationMins || 0), 0);
}
function getPosition() {
  return new Promise(resolve => {
    if (!navigator.geolocation) return resolve(null);
    const done = pos => resolve({
      lat: +pos.coords.latitude.toFixed(6),
      lng: +pos.coords.longitude.toFixed(6),
      accuracy: Math.round(pos.coords.accuracy || 0),
    });
    navigator.geolocation.getCurrentPosition(done, () => resolve(null), { enableHighAccuracy: true, timeout: 12000, maximumAge: 30000 });
  });
}
function mapLink(lat, lng) {
  if (lat == null || lng == null) return "";
  return `<a class="map-link" target="_blank" rel="noopener" href="https://www.google.com/maps?q=${lat},${lng}">${lat}, ${lng}</a>`;
}
function startClock() {
  if (clockTimer) { clearInterval(clockTimer); clockTimer = null; }
  if (state.tab !== "field") return;
  clockTimer = setInterval(() => {
    const el = $("punch-elapsed");
    if (!el) { clearInterval(clockTimer); clockTimer = null; return; }
    const open = openSessionFor(state.user.uid);
    if (!open) return;
    el.textContent = fmtDuration((Date.now() - new Date(open.checkInAt).getTime()) / 60000);
  }, 20000);
}
 
function tplField() {
  const open = openSessionFor(state.user.uid);
  const mine = state.attendance
    .filter(a => a.employeeId === state.user.uid)
    .sort((a, b) => String(b.checkInAt || "").localeCompare(String(a.checkInAt || "")))
    .slice(0, 25);
  const elapsed = open ? fmtDuration((Date.now() - new Date(open.checkInAt).getTime()) / 60000) : "";
  const deptOk = isSuperAdmin() || !!state.user.department;
 
  return `
    ${pageHeader("Field Check-in", "Record when you start and finish work outside the office.")}
 
    <div class="surface punch-card mb-6">
      ${open ? `
        <div class="text-sm mb-2" style="color:var(--success)"><span class="live-dot"></span>Checked in at ${fmtTime(open.checkInAt)}</div>
        <div class="punch-clock" id="punch-elapsed">${elapsed}</div>
        <div class="punch-sub">${open.checkInNote ? esc(open.checkInNote) : "No location note"}${open.checkInLat != null ? " · " + mapLink(open.checkInLat, open.checkInLng) : ""}</div>
        <button class="btn btn-danger punch-btn" onclick="modalPunch('out')">Check out</button>
      ` : `
        <div class="punch-clock">${fmtTime(new Date().toISOString())}</div>
        <div class="punch-sub">${fmtDate(todayISO())} · You're not checked in</div>
        <button class="btn btn-primary punch-btn" ${deptOk ? "" : "disabled"} onclick="modalPunch('in')">Check in</button>
        ${deptOk ? "" : `<p class="text-sm mt-3" style="color:var(--danger)">You're not assigned to a department yet — ask your admin to fix this first.</p>`}
      `}
    </div>
 
    <div class="surface overflow-hidden">
      <div class="px-5 py-4" style="border-bottom:1px solid var(--border)">
        <div class="font-semibold text-sm">My recent field sessions</div>
      </div>
      ${mine.length === 0 ? `<div class="empty">No sessions yet. Check in when you head out.</div>` : `
      <div class="table-wrap"><table class="rtable">
        <thead><tr><th>Date</th><th>In</th><th>Out</th><th>Duration</th><th>Where / what</th></tr></thead>
        <tbody>
        ${mine.map(a => `
          <tr>
            <td data-label="Date" class="font-medium">${fmtDate(a.date)}</td>
            <td data-label="In">${fmtTime(a.checkInAt)}</td>
            <td data-label="Out">${a.open ? openBadge() : fmtTime(a.checkOutAt)}</td>
            <td data-label="Duration">${a.open ? "—" : fmtDuration(a.durationMins)}</td>
            <td data-label="Where / what">${esc(a.checkInNote) || "—"}${a.checkInLat != null ? "<br>" + mapLink(a.checkInLat, a.checkInLng) : ""}</td>
          </tr>`).join("")}
        </tbody>
      </table></div>`}
    </div>
  `;
}
 
function modalPunch(dir) {
  const inMode = dir === "in";
  const open = openSessionFor(state.user.uid);
  if (inMode && open) { toast("You're already checked in.", "err"); return; }
  if (!inMode && !open) { toast("You're not checked in right now.", "err"); return; }
 
  openModal(`
    <div class="p-6">
      <h2 class="text-lg font-semibold mb-1">${inMode ? "Check in" : "Check out"}</h2>
      <p class="text-sm mb-5" style="color:var(--text-muted)">
        ${inMode ? "Add where you're going or what you're doing. Your location is attached if you allow it." : `You checked in at ${fmtTime(open.checkInAt)}.`}
      </p>
 
      <label class="label">${inMode ? "Where / what (optional)" : "What you did (optional)"}</label>
      <input class="input mb-4" id="punch-note" placeholder="${inMode ? "e.g. Customer visit, Galle" : "e.g. Delivered samples, 3 stops"}" autofocus>
 
      <label class="flex items-center gap-2 mb-4 text-sm" style="color:var(--text-muted)">
        <input type="checkbox" id="punch-geo" checked> Attach my location
      </label>
 
      <p id="punch-err" class="form-err"></p>
 
      <div class="flex gap-2 justify-end">
        <button type="button" class="btn btn-secondary" onclick="closeModal()">Cancel</button>
        <button type="button" id="punch-go" class="btn ${inMode ? "btn-primary" : "btn-danger"}" onclick="handlePunch('${dir}')">${inMode ? "Check in" : "Check out"}</button>
      </div>
    </div>
  `);
}
 
async function handlePunch(dir) {
  if (punchBusy) return;
  const inMode = dir === "in";
  const btn = $("punch-go"), errEl = $("punch-err");
  const note = ($("punch-note") && $("punch-note").value.trim()) || null;
  const wantGeo = $("punch-geo") ? $("punch-geo").checked : false;
  if (errEl) errEl.textContent = "";
 
  punchBusy = true;
  if (btn) { btn.disabled = true; btn.textContent = wantGeo ? "Getting location…" : (inMode ? "Checking in…" : "Checking out…"); }
 
  try {
    const pos = wantGeo ? await getPosition() : null;
    if (btn) btn.textContent = inMode ? "Checking in…" : "Checking out…";
 
    if (inMode) {
      if (openSessionFor(state.user.uid)) throw new Error("You're already checked in.");
      const now = new Date().toISOString();
      await db.collection("attendance").add({
        employeeId: state.user.uid,
        employeeName: state.user.name,
        department: state.user.department || null,
        date: todayISO(),
        open: true,
        checkInAt: now,
        checkInNote: note,
        checkInLat: pos ? pos.lat : null,
        checkInLng: pos ? pos.lng : null,
        checkInAccuracy: pos ? pos.accuracy : null,
        checkOutAt: null, checkOutNote: null, checkOutLat: null, checkOutLng: null, checkOutAccuracy: null,
        durationMins: null,
        createdAt: firebase.firestore.FieldValue.serverTimestamp(),
      });
      closeModal();
      toast(pos ? "Checked in with location." : "Checked in.", "ok");
    } else {
      const open = openSessionFor(state.user.uid);
      if (!open) throw new Error("You're not checked in right now.");
      const now = new Date().toISOString();
      const mins = Math.max(0, Math.round((new Date(now) - new Date(open.checkInAt)) / 60000));
      await db.collection("attendance").doc(open.id).update({
        open: false,
        checkOutAt: now,
        checkOutNote: note,
        checkOutLat: pos ? pos.lat : null,
        checkOutLng: pos ? pos.lng : null,
        checkOutAccuracy: pos ? pos.accuracy : null,
        durationMins: mins,
      });
      closeModal();
      toast(`Checked out. ${fmtDuration(mins)} recorded.`, "ok");
    }
  } catch (ex) {
    console.error("punch failed", ex);
    if (errEl) errEl.textContent = friendlyAuthError(ex);
    if (btn) { btn.disabled = false; btn.textContent = inMode ? "Check in" : "Check out"; }
  } finally {
    punchBusy = false;
  }
}
 
/* ============================= ATTENDANCE (ADMIN) ============================= */
function attendanceFiltered() {
  let list = attendanceScope();
  if (isSuperAdmin() && state.attFilter.dept !== "all") list = list.filter(a => a.department === state.attFilter.dept);
  if (state.attFilter.from) list = list.filter(a => String(a.date || "") >= state.attFilter.from);
  if (state.attFilter.to) list = list.filter(a => String(a.date || "") <= state.attFilter.to);
  return list;
}
 
function tplAttendance() {
  const list = [...attendanceFiltered()].sort((a, b) => String(b.checkInAt || "").localeCompare(String(a.checkInAt || "")));
  const openNow = list.filter(a => a.open);
  const totalMins = list.reduce((s, a) => s + (a.durationMins || 0), 0);
 
  return `
    ${pageHeader("Attendance", isSuperAdmin() ? "Field check-ins across all departments." : `Field check-ins in ${esc(deptName(state.user.department))}.`,
      `<button class="btn btn-secondary" onclick="exportAttendanceExcel()">Export to Excel</button>`)}
 
    <div class="grid gap-4 mb-5" style="grid-template-columns:repeat(auto-fit,minmax(160px,1fr))">
      <div class="surface stat" style="--stat:var(--success)">
        <div class="text-sm mb-2" style="color:var(--text-muted)">In the field now</div>
        <div class="text-3xl font-semibold" style="color:var(--success)">${openNow.length}</div>
      </div>
      <div class="surface stat" style="--stat:var(--navy)">
        <div class="text-sm mb-2" style="color:var(--text-muted)">Sessions shown</div>
        <div class="text-3xl font-semibold" style="color:var(--navy)">${list.length}</div>
      </div>
      <div class="surface stat" style="--stat:var(--navy)">
        <div class="text-sm mb-2" style="color:var(--text-muted)">Time recorded</div>
        <div class="text-3xl font-semibold" style="color:var(--navy)">${fmtDuration(totalMins)}</div>
      </div>
    </div>
 
    <div class="filter-row flex gap-2 mb-5 flex-wrap items-end">
      <div><label class="label">From</label><input class="input" type="date" value="${state.attFilter.from}" onchange="state.attFilter.from=this.value; render();"></div>
      <div><label class="label">To</label><input class="input" type="date" value="${state.attFilter.to}" onchange="state.attFilter.to=this.value; render();"></div>
      ${isSuperAdmin() && state.departments.length ? `
      <div><label class="label">Department</label>
        <select class="input" onchange="state.attFilter.dept=this.value; render();">
          <option value="all" ${state.attFilter.dept === "all" ? "selected" : ""}>All departments</option>
          ${state.departments.map(d => `<option value="${d.id}" ${state.attFilter.dept === d.id ? "selected" : ""}>${esc(d.name)}</option>`).join("")}
        </select>
      </div>` : ""}
      <button class="btn btn-secondary" onclick="state.attFilter={dept:'all',from:'',to:''}; render();">Clear filters</button>
    </div>
 
    <div class="surface overflow-hidden">
      ${list.length === 0 ? `<div class="empty">No check-ins match these filters.</div>` : `
      <div class="table-wrap"><table class="rtable">
        <thead><tr><th>Employee</th>${isSuperAdmin() ? "<th>Dept</th>" : ""}<th>Date</th><th>In</th><th>Out</th><th>Duration</th><th>Notes</th><th>Location</th></tr></thead>
        <tbody>
        ${list.map(a => `
          <tr>
            <td data-label="Employee"><div class="cell-user"><div class="avatar">${esc(initials(a.employeeName))}</div><span class="font-medium">${esc(a.employeeName)}</span></div></td>
            ${isSuperAdmin() ? `<td data-label="Dept">${esc(deptName(a.department))}</td>` : ""}
            <td data-label="Date">${fmtDate(a.date)}</td>
            <td data-label="In">${fmtTime(a.checkInAt)}</td>
            <td data-label="Out">${a.open ? openBadge() : fmtTime(a.checkOutAt)}</td>
            <td data-label="Duration">${a.open ? "—" : fmtDuration(a.durationMins)}</td>
            <td data-label="Notes" style="max-width:220px">${esc(a.checkInNote) || "—"}${a.checkOutNote ? "<br><span style='color:var(--text-muted)'>Out: " + esc(a.checkOutNote) + "</span>" : ""}</td>
            <td data-label="Location">${a.checkInLat != null ? mapLink(a.checkInLat, a.checkInLng) : "—"}</td>
          </tr>`).join("")}
        </tbody>
      </table></div>`}
    </div>
  `;
}
 
function exportAttendanceExcel() {
  const list = attendanceFiltered();
  if (list.length === 0) { toast("Nothing to export for these filters.", "err"); return; }
  const rows = list.map(a => ({
    Employee: a.employeeName, Department: deptName(a.department), Date: a.date,
    "Check In": a.checkInAt ? a.checkInAt.slice(11, 16) : "",
    "Check Out": a.checkOutAt ? a.checkOutAt.slice(11, 16) : "",
    "Hours": a.durationMins != null ? +(a.durationMins / 60).toFixed(2) : "",
    "In Note": a.checkInNote || "", "Out Note": a.checkOutNote || "",
    "In Location": a.checkInLat != null ? `${a.checkInLat}, ${a.checkInLng}` : "",
    "Out Location": a.checkOutLat != null ? `${a.checkOutLat}, ${a.checkOutLng}` : "",
    Status: a.open ? "Open" : "Closed",
  }));
  const ws = XLSX.utils.json_to_sheet(rows);
  ws["!cols"] = [{wch:20},{wch:16},{wch:12},{wch:10},{wch:10},{wch:8},{wch:26},{wch:26},{wch:22},{wch:22},{wch:9}];
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Attendance");
  XLSX.writeFile(wb, `attendance-${todayISO()}.xlsx`);
}
 
/* ============================= MY LEAVE ============================= */
function tplMyLeave() {
  const mine = state.leaves.filter(l => l.employeeId === state.user.uid)
    .sort((a, b) => String(b.appliedAt || "").localeCompare(String(a.appliedAt || "")));
 
  return `
    ${pageHeader("My Leave", "Submit a new request or track your leave history.",
      `<button class="btn btn-primary" onclick="modalSubmitLeave()">+ New leave request</button>`)}
    <div class="surface overflow-hidden">
      ${mine.length === 0 ? `<div class="empty">You haven't submitted any leave requests yet. Use "New leave request" to file your first one.</div>` : `
      <div class="table-wrap"><table class="rtable">
        <thead><tr><th>Type</th><th>Dates</th><th>Days</th><th>Reason</th><th>Status</th><th>Note from admin</th><th></th></tr></thead>
        <tbody>
        ${mine.map(l => `
          <tr>
            <td data-label="Type" class="font-medium">${esc(l.leaveType)}</td>
            <td data-label="Dates">${fmtRange(l)}</td>
            <td data-label="Days">${l.totalDays}</td>
            <td data-label="Reason" style="max-width:220px">${esc(l.reason) || "—"}</td>
            <td data-label="Status">${statusBadge(l.status)}</td>
            <td data-label="Note from admin" style="max-width:200px; color:var(--text-muted)">${esc(l.adminNote) || "—"}</td>
            <td class="${l.status === "pending" ? "td-action" : "td-empty"}">${l.status === "pending" ? `<button class="btn btn-sm btn-ghost" onclick="handleCancelLeave('${l.id}')">Cancel</button>` : ""}</td>
          </tr>`).join("")}
        </tbody>
      </table></div>`}
    </div>
  `;
}
 
function modalSubmitLeave() {
  const deptOk = state.user.role === "superadmin" || state.user.department;
  openModal(`
    <form onsubmit="handleSubmitLeave(event)" class="p-6">
      <h2 class="text-lg font-semibold mb-1">New leave request</h2>
      <p class="text-sm mb-5" style="color:var(--text-muted)">Your admin will review and respond.</p>
 
      <label class="label">Leave type</label>
      <select class="input mb-4" name="leaveType" required>
        ${LEAVE_TYPES.map(t => `<option value="${esc(t)}">${esc(t)}</option>`).join("")}
      </select>
 
      <div class="grid grid-cols-2 gap-3 mb-1">
        <div><label class="label">From</label><input class="input" type="date" name="startDate" required value="${todayISO()}"></div>
        <div><label class="label">To</label><input class="input" type="date" name="endDate" required value="${todayISO()}"></div>
      </div>
      <label class="flex items-center gap-2 mt-3 mb-4 text-sm" style="color:var(--text-muted)">
        <input type="checkbox" name="halfDay"> This is a half-day request
      </label>
 
      <label class="label">Reason</label>
      <textarea class="input mb-5" name="reason" rows="3" placeholder="Briefly explain the reason for leave" required></textarea>
 
      ${!deptOk ? `<p class="text-sm mb-4" style="color:var(--danger)">You're not assigned to a department yet — ask your admin to fix this before submitting.</p>` : ""}
 
      <div class="flex gap-2 justify-end">
        <button type="button" class="btn btn-secondary" onclick="closeModal()">Cancel</button>
        <button type="submit" class="btn btn-primary" ${!deptOk ? "disabled" : ""} id="submit-leave-btn">Submit request</button>
      </div>
    </form>
  `);
}
 
async function handleSubmitLeave(e) {
  e.preventDefault();
  const f = e.target, btn = $("submit-leave-btn");
  const half = f.halfDay.checked;
  const startDate = f.startDate.value, endDate = half ? f.startDate.value : f.endDate.value;
  if (new Date(endDate) < new Date(startDate)) { toast("End date can't be before start date.", "err"); return; }
  btn.disabled = true; btn.textContent = "Submitting…";
  try {
    await db.collection("leaveRequests").add({
      employeeId: state.user.uid,
      employeeName: state.user.name,
      department: state.user.department,
      leaveType: f.leaveType.value,
      startDate, endDate,
      totalDays: daysBetween(startDate, endDate, half),
      reason: f.reason.value.trim(),
      status: "pending",
      appliedAt: new Date().toISOString(),
      reviewedBy: null, reviewedByName: null, reviewedAt: null, adminNote: null,
    });
    closeModal();
    toast("Leave request submitted.", "ok");
  } catch (ex) {
    console.error(ex);
    toast(friendlyAuthError(ex), "err");
    btn.disabled = false; btn.textContent = "Submit request";
  }
}
 
async function handleCancelLeave(id) {
  if (!confirm("Cancel this leave request?")) return;
  try {
    await db.collection("leaveRequests").doc(id).delete();
    toast("Request cancelled.", "ok");
  } catch (ex) { console.error(ex); toast(friendlyAuthError(ex), "err"); }
}
 
/* ============================= LEAVE REQUESTS (ADMIN) ============================= */
function tplLeaves() {
  const base = adminBaseList();
  const list = [...adminFilteredList()].sort((a, b) => String(b.appliedAt || "").localeCompare(String(a.appliedAt || "")));
  const counts = {
    all: base.length,
    pending: base.filter(l => l.status === "pending").length,
    approved: base.filter(l => l.status === "approved").length,
    rejected: base.filter(l => l.status === "rejected").length,
  };
 
  return `
    ${pageHeader("Leave Requests", isSuperAdmin() ? "Across all departments." : `For ${esc(deptName(state.user.department))}.`,
      `<button class="btn btn-secondary" onclick="exportLeavesExcel()">Export to Excel</button>`)}
 
    <div class="filter-row flex gap-2 mb-5 flex-wrap items-center">
      ${["all", "pending", "approved", "rejected"].map(s => `
        <button class="btn btn-sm ${state.leaveFilter.status === s ? "btn-dark" : "btn-secondary"}" onclick="state.leaveFilter.status='${s}'; render();">${s[0].toUpperCase() + s.slice(1)}<span class="chip-count">${counts[s]}</span></button>
      `).join("")}
      ${isSuperAdmin() && state.departments.length ? `
        <select class="input" style="width:auto;" onchange="state.leaveFilter.dept=this.value; render();">
          <option value="all" ${state.leaveFilter.dept === "all" ? "selected" : ""}>All departments</option>
          ${state.departments.map(d => `<option value="${d.id}" ${state.leaveFilter.dept === d.id ? "selected" : ""}>${esc(d.name)}</option>`).join("")}
        </select>` : ""}
    </div>
 
    <div class="surface overflow-hidden">
      ${list.length === 0 ? `<div class="empty">No requests match this filter.</div>` : `
      <div class="table-wrap"><table class="rtable">
        <thead><tr><th>Employee</th>${isSuperAdmin() ? "<th>Dept</th>" : ""}<th>Type</th><th>Dates</th><th>Days</th><th>Reason</th><th>Status</th><th></th></tr></thead>
        <tbody>
        ${list.map(l => `
          <tr>
            <td data-label="Employee"><div class="cell-user"><div class="avatar">${esc(initials(l.employeeName))}</div><span class="font-medium">${esc(l.employeeName)}</span></div></td>
            ${isSuperAdmin() ? `<td data-label="Dept">${esc(deptName(l.department))}</td>` : ""}
            <td data-label="Type">${esc(l.leaveType)}</td>
            <td data-label="Dates">${fmtRange(l)}</td>
            <td data-label="Days">${l.totalDays}</td>
            <td data-label="Reason" style="max-width:200px">${esc(l.reason) || "—"}</td>
            <td data-label="Status">${statusBadge(l.status)}</td>
            <td class="td-action" style="white-space:nowrap">${l.status === "pending"
              ? `<button class="btn btn-sm btn-primary" onclick="modalReview('${l.id}')">Review</button>`
              : `<button class="btn btn-sm btn-secondary" onclick="modalReview('${l.id}')">Details</button>`}</td>
          </tr>`).join("")}
        </tbody>
      </table></div>`}
    </div>
  `;
}
 
/* ---- Approve / Reject ----
   The buttons call handleReview(id, decision) directly (no form submit), so the click
   always reaches Firestore, and any error is shown inside the popup instead of a toast
   that could sit behind it. */
function modalReview(id) {
  const l = state.leaves.find(x => x.id === id);
  if (!l) { toast("That request could not be found.", "err"); return; }
  if (!isAdmin()) { toast("Only admins can review leave requests.", "err"); return; }
  const pendingNow = l.status === "pending";
 
  openModal(`
    <div class="p-6">
      <div class="flex items-start justify-between gap-3 mb-5">
        <div>
          <h2 class="text-lg font-semibold">${pendingNow ? "Review request" : "Leave request"}</h2>
          <p class="text-sm mt-1" style="color:var(--text-muted)">${esc(l.employeeName)}, ${esc(deptName(l.department))}</p>
        </div>
        ${statusBadge(l.status)}
      </div>
 
      <div class="detail-grid">
        <div><div class="k">Leave type</div><div class="v">${esc(l.leaveType)}</div></div>
        <div><div class="k">Duration</div><div class="v">${l.totalDays} day${l.totalDays === 1 ? "" : "s"}</div></div>
        <div><div class="k">Dates</div><div class="v">${fmtRange(l)}</div></div>
        <div><div class="k">Applied</div><div class="v">${fmtDateTime(l.appliedAt)}</div></div>
      </div>
 
      <div class="k mb-1">Reason</div>
      <div class="surface-alt rounded-lg p-3 text-sm mb-4">${esc(l.reason) || "No reason given."}</div>
 
      ${!pendingNow ? `<p class="text-sm mb-4" style="color:var(--text-muted)">${l.status === "approved" ? "Approved" : "Rejected"} by ${esc(l.reviewedByName || "—")} on ${fmtDateTime(l.reviewedAt)}. You can still change the decision below.</p>` : ""}
 
      <label class="label">Note to employee (optional)</label>
      <textarea class="input mb-3" id="review-note" rows="2" placeholder="e.g. Approved, please hand over pending tasks.">${esc(l.adminNote || "")}</textarea>
 
      <p id="review-err" class="form-err"></p>
 
      <div class="flex gap-2 justify-end flex-wrap">
        <button type="button" class="btn btn-secondary" onclick="closeModal()">Close</button>
        ${l.status !== "rejected" ? `<button type="button" data-review-btn data-decision="rejected" class="btn btn-danger" onclick="handleReview('${id}','rejected')">Reject</button>` : ""}
        ${l.status !== "approved" ? `<button type="button" data-review-btn data-decision="approved" class="btn btn-success" onclick="handleReview('${id}','approved')">Approve</button>` : ""}
      </div>
    </div>
  `);
}
 
async function handleReview(id, decision) {
  if (reviewBusy) return;
  if (decision !== "approved" && decision !== "rejected") return;
 
  const l = state.leaves.find(x => x.id === id);
  const errEl = $("review-err"), noteEl = $("review-note");
  const btns = Array.from(document.querySelectorAll("#modal-root [data-review-btn]"));
  if (errEl) errEl.textContent = "";
 
  if (!isAdmin()) { if (errEl) errEl.textContent = "Only admins can review leave requests."; return; }
  if (l && isDeptAdmin() && l.department !== state.user.department) {
    if (errEl) errEl.textContent = "You can only review requests from your own department.";
    return;
  }
 
  reviewBusy = true;
  btns.forEach(b => { b.disabled = true; });
  const clicked = btns.find(b => b.dataset.decision === decision);
  const oldLabel = clicked ? clicked.textContent : "";
  if (clicked) clicked.textContent = decision === "approved" ? "Approving…" : "Rejecting…";
 
  try {
    await db.collection("leaveRequests").doc(id).update({
      status: decision,
      adminNote: (noteEl && noteEl.value.trim()) || null,
      reviewedBy: state.user.uid,
      reviewedByName: state.user.name,
      reviewedAt: new Date().toISOString(),
    });
    closeModal();
    toast(`Request ${decision}.`, "ok");
  } catch (ex) {
    console.error("review failed", ex);
    if (errEl) errEl.textContent = friendlyAuthError(ex);
    btns.forEach(b => { b.disabled = false; });
    if (clicked) clicked.textContent = oldLabel;
  } finally {
    reviewBusy = false;
  }
}
 
function exportLeavesExcel() {
  const list = adminFilteredList();
  if (list.length === 0) { toast("Nothing to export for this filter.", "err"); return; }
 
  const rows = list.map(l => ({
    Employee: l.employeeName, Department: deptName(l.department), "Leave Type": l.leaveType,
    "Start Date": l.startDate, "End Date": l.endDate, Days: l.totalDays, Reason: l.reason || "",
    Status: l.status, "Reviewed By": l.reviewedByName || "", "Admin Note": l.adminNote || "",
    "Applied At": l.appliedAt ? l.appliedAt.slice(0, 16).replace("T", " ") : "",
  }));
  const ws = XLSX.utils.json_to_sheet(rows);
  ws["!cols"] = [{wch:20},{wch:16},{wch:14},{wch:12},{wch:12},{wch:6},{wch:30},{wch:10},{wch:16},{wch:26},{wch:17}];
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Leave Requests");
  XLSX.writeFile(wb, `leave-requests-${todayISO()}.xlsx`);
}
 
/* ============================= TRANSPORT BILLS (EMPLOYEE) ============================= */
function myBillsScope() {
  return state.bills.filter(b => b.employeeId === state.user.uid);
}
 
function tplMyBills() {
  const mine = myBillsScope().sort((a, b) => String(b.date || "").localeCompare(String(a.date || "")));
  const totalClaimed = mine.reduce((s, b) => s + (Number(b.amount) || 0), 0);
  const totalApproved = mine.filter(b => b.status === "approved").reduce((s, b) => s + (Number(b.amount) || 0), 0);
 
  return `
    ${pageHeader("Transport Bills", "Submit a PickMe / taxi bill for reimbursement, with the slip attached.",
      `<button class="btn btn-primary" onclick="modalSubmitBill()">+ New transport bill</button>`)}
 
    <div class="grid gap-4 mb-6" style="grid-template-columns:repeat(auto-fit,minmax(160px,1fr))">
      <div class="surface stat" style="--stat:var(--navy)">
        <div class="text-sm mb-2" style="color:var(--text-muted)">Bills filed</div>
        <div class="text-3xl font-semibold" style="color:var(--navy)">${mine.length}</div>
      </div>
      <div class="surface stat" style="--stat:var(--warn)">
        <div class="text-sm mb-2" style="color:var(--text-muted)">Total claimed</div>
        <div class="text-3xl font-semibold" style="color:var(--warn)">${fmtMoney(totalClaimed)}</div>
      </div>
      <div class="surface stat" style="--stat:var(--success)">
        <div class="text-sm mb-2" style="color:var(--text-muted)">Approved amount</div>
        <div class="text-3xl font-semibold" style="color:var(--success)">${fmtMoney(totalApproved)}</div>
      </div>
    </div>
 
    <div class="surface overflow-hidden">
      ${mine.length === 0 ? `<div class="empty">No transport bills yet. Use "New transport bill" to submit your first one.</div>` : `
      <div class="table-wrap"><table class="rtable">
        <thead><tr><th>Slip</th><th>Date</th><th>Route</th><th>Km</th><th>Amount</th><th>Project</th><th>Status</th><th>Note</th><th></th></tr></thead>
        <tbody>
        ${mine.map(b => `
          <tr>
            <td data-label="Slip">${b.slipUrl ? `<img class="slip-thumb" src="${esc(b.slipUrl)}" onclick="viewSlip('${esc(b.slipUrl)}')" alt="Slip">` : "—"}</td>
            <td data-label="Date">${fmtDate(b.date)}</td>
            <td data-label="Route">${esc(b.fromLocation)}<span class="route-arrow">→</span>${esc(b.toLocation)}</td>
            <td data-label="Km">${b.distanceKm ?? "—"}</td>
            <td data-label="Amount" class="amount">${fmtMoney(b.amount)}</td>
            <td data-label="Project">${esc(projectName(b.project))}</td>
            <td data-label="Status">${statusBadge(b.status)}</td>
            <td data-label="Note" style="max-width:180px; color:var(--text-muted)">${esc(b.adminNote) || "—"}</td>
            <td class="${b.status === "pending" ? "td-action" : "td-empty"}">${b.status === "pending" ? `<button class="btn btn-sm btn-ghost" onclick="handleCancelBill('${b.id}')">Cancel</button>` : ""}</td>
          </tr>`).join("")}
        </tbody>
      </table></div>`}
    </div>
  `;
}
 
function modalSubmitBill() {
  const deptOk = isSuperAdmin() || !!state.user.department;
  pendingSlipFile = null;
  const projectOptions = state.projects.length
    ? `<option value="">No specific project</option>` + state.projects.map(p => `<option value="${p.id}">${esc(p.name)}</option>`).join("")
    : `<option value="">No projects set up yet</option>`;
 
  openModal(`
    <form onsubmit="handleSubmitBill(event)" class="p-6">
      <h2 class="text-lg font-semibold mb-1">New transport bill</h2>
      <p class="text-sm mb-5" style="color:var(--text-muted)">Attach the PickMe / taxi slip as a photo or screenshot.</p>
 
      <div class="grid grid-cols-2 gap-3 mb-4">
        <div><label class="label">From</label><input class="input" name="fromLocation" required placeholder="e.g. Office, Colombo 03"></div>
        <div><label class="label">To</label><input class="input" name="toLocation" required placeholder="e.g. Client site, Rajagiriya"></div>
      </div>
 
      <div class="grid grid-cols-2 gap-3 mb-4">
        <div><label class="label">Distance (km)</label><input class="input" name="distanceKm" type="number" min="0" step="0.1" placeholder="e.g. 8.4"></div>
        <div><label class="label">Amount (Rs.)</label><input class="input" name="amount" type="number" min="0" step="0.01" required placeholder="e.g. 950.00"></div>
      </div>
 
      <div class="grid grid-cols-2 gap-3 mb-4">
        <div><label class="label">Date</label><input class="input" name="date" type="date" required value="${todayISO()}"></div>
        <div><label class="label">Project (optional)</label><select class="input" name="project">${projectOptions}</select></div>
      </div>
 
      <label class="label">Additional details (optional)</label>
      <textarea class="input mb-4" name="notes" rows="2" placeholder="e.g. Client meeting, round trip, urgent delivery"></textarea>
 
      <label class="label">Bill / slip photo</label>
      <div class="upload-box mb-1" id="upload-box" onclick="document.getElementById('slip-file').click()">
        <div id="upload-box-content">
          <div class="icn">
            <svg width="26" height="26" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" style="margin:0 auto"><path d="M12 16V4M12 4l-4 4M12 4l4 4"/><path d="M4 16v3a1 1 0 001 1h14a1 1 0 001-1v-3"/></svg>
          </div>
          <div class="text-sm font-medium">Tap to choose a photo</div>
          <div class="text-xs mt-1" style="color:var(--text-muted)">PNG or JPG, up to ${MAX_SLIP_MB}MB</div>
        </div>
        <input id="slip-file" type="file" accept="image/png,image/jpeg" onchange="handleSlipPick(event)">
      </div>
      <p class="text-xs mb-4" style="color:var(--text-muted)">The slip is optional, but helps your admin verify the claim faster.</p>
 
      ${!deptOk ? `<p class="text-sm mb-4" style="color:var(--danger)">You're not assigned to a department yet — ask your admin to fix this before submitting.</p>` : ""}
 
      <p id="bill-err" class="form-err"></p>
      <div id="bill-progress"></div>
 
      <div class="flex gap-2 justify-end mt-2">
        <button type="button" class="btn btn-secondary" onclick="closeModal()">Cancel</button>
        <button type="submit" class="btn btn-primary" ${!deptOk ? "disabled" : ""} id="submit-bill-btn">Submit bill</button>
      </div>
    </form>
  `);
}
 
function handleSlipPick(e) {
  const file = e.target.files && e.target.files[0];
  const box = $("upload-box-content");
  const errEl = $("bill-err");
  if (errEl) errEl.textContent = "";
  if (!file) { pendingSlipFile = null; return; }
  if (!/^image\/(png|jpeg)$/.test(file.type)) {
    if (errEl) errEl.textContent = "Please choose a PNG or JPG image.";
    e.target.value = ""; pendingSlipFile = null; return;
  }
  if (file.size > MAX_SLIP_MB * 1024 * 1024) {
    if (errEl) errEl.textContent = `That image is over ${MAX_SLIP_MB}MB. Please choose a smaller one.`;
    e.target.value = ""; pendingSlipFile = null; return;
  }
  pendingSlipFile = file;
  const url = URL.createObjectURL(file);
  if (box) box.innerHTML = `<img class="slip-preview" src="${url}" alt="Selected slip"><div class="text-xs mt-2" style="color:var(--text-muted)">${esc(file.name)} · tap to change</div>`;
}
 
async function uploadSlip(uid, file, onProgress) {
  const path = `transport-slips/${uid}/${Date.now()}_${file.name.replace(/[^a-zA-Z0-9._-]/g, "_")}`;
  const ref = storage.ref().child(path);
  const task = ref.put(file, { contentType: file.type });
  return new Promise((resolve, reject) => {
    task.on("state_changed",
      snap => onProgress && onProgress(Math.round((snap.bytesTransferred / snap.totalBytes) * 100)),
      err => reject(err),
      async () => { try { resolve(await task.snapshot.ref.getDownloadURL()); } catch (ex) { reject(ex); } }
    );
  });
}
 
async function handleSubmitBill(e) {
  e.preventDefault();
  const f = e.target, btn = $("submit-bill-btn"), errEl = $("bill-err"), prog = $("bill-progress");
  if (errEl) errEl.textContent = "";
  const amount = parseFloat(f.amount.value);
  if (!(amount > 0)) { if (errEl) errEl.textContent = "Enter a valid amount."; return; }
 
  btn.disabled = true; btn.textContent = "Submitting…";
  try {
    let slipUrl = null;
    if (pendingSlipFile) {
      btn.textContent = "Uploading slip…";
      slipUrl = await uploadSlip(state.user.uid, pendingSlipFile, pct => {
        if (prog) prog.innerHTML = `<div class="progress-bar"><div style="width:${pct}%"></div></div>`;
      });
    }
    btn.textContent = "Saving…";
    await db.collection("transportBills").add({
      employeeId: state.user.uid,
      employeeName: state.user.name,
      department: state.user.department || null,
      date: f.date.value,
      fromLocation: f.fromLocation.value.trim(),
      toLocation: f.toLocation.value.trim(),
      distanceKm: f.distanceKm.value ? parseFloat(f.distanceKm.value) : null,
      amount,
      project: f.project.value || null,
      notes: f.notes.value.trim() || null,
      slipUrl,
      status: "pending",
      appliedAt: new Date().toISOString(),
      reviewedBy: null, reviewedByName: null, reviewedAt: null, adminNote: null,
    });
    pendingSlipFile = null;
    closeModal();
    toast("Transport bill submitted.", "ok");
  } catch (ex) {
    console.error("bill submit failed", ex);
    if (errEl) errEl.textContent = friendlyAuthError(ex);
    btn.disabled = false; btn.textContent = "Submit bill";
    if (prog) prog.innerHTML = "";
  }
}
 
async function handleCancelBill(id) {
  if (!confirm("Cancel this transport bill?")) return;
  try {
    const b = state.bills.find(x => x.id === id);
    await db.collection("transportBills").doc(id).delete();
    if (b && b.slipUrl) { try { await storage.refFromURL(b.slipUrl).delete(); } catch (_) {} }
    toast("Bill cancelled.", "ok");
  } catch (ex) { console.error(ex); toast(friendlyAuthError(ex), "err"); }
}
 
/* ============================= TRANSPORT BILLS (ADMIN) ============================= */
function tplBillsAdmin() {
  const base = adminScopeBills();
  const list = [...billFilteredList()].sort((a, b) => String(b.date || "").localeCompare(String(a.date || "")));
  const counts = {
    all: base.length,
    pending: base.filter(b => b.status === "pending").length,
    approved: base.filter(b => b.status === "approved").length,
    rejected: base.filter(b => b.status === "rejected").length,
  };
  const totalShown = list.reduce((s, b) => s + (Number(b.amount) || 0), 0);
 
  return `
    ${pageHeader("Transport Claims", isSuperAdmin() ? "PickMe / taxi bills across all departments." : `PickMe / taxi bills for ${esc(deptName(state.user.department))}.`,
      `<button class="btn btn-secondary" onclick="exportBillsExcel()">Export to Excel</button>`)}
 
    <div class="grid gap-4 mb-5" style="grid-template-columns:repeat(auto-fit,minmax(160px,1fr))">
      <div class="surface stat" style="--stat:var(--warn)">
        <div class="text-sm mb-2" style="color:var(--text-muted)">Pending review</div>
        <div class="text-3xl font-semibold" style="color:var(--warn)">${counts.pending}</div>
      </div>
      <div class="surface stat" style="--stat:var(--navy)">
        <div class="text-sm mb-2" style="color:var(--text-muted)">Bills shown</div>
        <div class="text-3xl font-semibold" style="color:var(--navy)">${list.length}</div>
      </div>
      <div class="surface stat" style="--stat:var(--navy)">
        <div class="text-sm mb-2" style="color:var(--text-muted)">Amount shown</div>
        <div class="text-3xl font-semibold" style="color:var(--navy)">${fmtMoney(totalShown)}</div>
      </div>
    </div>
 
    <div class="filter-row flex gap-2 mb-3 flex-wrap items-center">
      ${["all", "pending", "approved", "rejected"].map(s => `
        <button class="btn btn-sm ${state.billFilter.status === s ? "btn-dark" : "btn-secondary"}" onclick="state.billFilter.status='${s}'; render();">${s[0].toUpperCase() + s.slice(1)}<span class="chip-count">${counts[s]}</span></button>
      `).join("")}
    </div>
    <div class="filter-row flex gap-2 mb-5 flex-wrap items-end">
      <div><label class="label">From</label><input class="input" type="date" value="${state.billFilter.from}" onchange="state.billFilter.from=this.value; render();"></div>
      <div><label class="label">To</label><input class="input" type="date" value="${state.billFilter.to}" onchange="state.billFilter.to=this.value; render();"></div>
      ${isSuperAdmin() && state.departments.length ? `
      <div><label class="label">Department</label>
        <select class="input" onchange="state.billFilter.dept=this.value; render();">
          <option value="all" ${state.billFilter.dept === "all" ? "selected" : ""}>All departments</option>
          ${state.departments.map(d => `<option value="${d.id}" ${state.billFilter.dept === d.id ? "selected" : ""}>${esc(d.name)}</option>`).join("")}
        </select>
      </div>` : ""}
      ${state.projects.length ? `
      <div><label class="label">Project</label>
        <select class="input" onchange="state.billFilter.project=this.value; render();">
          <option value="all" ${state.billFilter.project === "all" ? "selected" : ""}>All projects</option>
          ${state.projects.map(p => `<option value="${p.id}" ${state.billFilter.project === p.id ? "selected" : ""}>${esc(p.name)}</option>`).join("")}
        </select>
      </div>` : ""}
      <button class="btn btn-secondary" onclick="state.billFilter={status:'all',dept:'all',project:'all',from:'',to:''}; render();">Clear filters</button>
    </div>
 
    <div class="surface overflow-hidden">
      ${list.length === 0 ? `<div class="empty">No transport bills match these filters.</div>` : `
      <div class="table-wrap"><table class="rtable">
        <thead><tr><th>Slip</th><th>Employee</th>${isSuperAdmin() ? "<th>Dept</th>" : ""}<th>Date</th><th>Route</th><th>Km</th><th>Amount</th><th>Project</th><th>Status</th><th></th></tr></thead>
        <tbody>
        ${list.map(b => `
          <tr>
            <td data-label="Slip">${b.slipUrl ? `<img class="slip-thumb" src="${esc(b.slipUrl)}" onclick="viewSlip('${esc(b.slipUrl)}')" alt="Slip">` : "—"}</td>
            <td data-label="Employee"><div class="cell-user"><div class="avatar">${esc(initials(b.employeeName))}</div><span class="font-medium">${esc(b.employeeName)}</span></div></td>
            ${isSuperAdmin() ? `<td data-label="Dept">${esc(deptName(b.department))}</td>` : ""}
            <td data-label="Date">${fmtDate(b.date)}</td>
            <td data-label="Route">${esc(b.fromLocation)}<span class="route-arrow">→</span>${esc(b.toLocation)}</td>
            <td data-label="Km">${b.distanceKm ?? "—"}</td>
            <td data-label="Amount" class="amount">${fmtMoney(b.amount)}</td>
            <td data-label="Project">${esc(projectName(b.project))}</td>
            <td data-label="Status">${statusBadge(b.status)}</td>
            <td class="td-action" style="white-space:nowrap">${b.status === "pending"
              ? `<button class="btn btn-sm btn-primary" onclick="modalBillReview('${b.id}')">Review</button>`
              : `<button class="btn btn-sm btn-secondary" onclick="modalBillReview('${b.id}')">Details</button>`}</td>
          </tr>`).join("")}
        </tbody>
      </table></div>`}
    </div>
  `;
}
 
function modalBillReview(id) {
  const b = state.bills.find(x => x.id === id);
  if (!b) { toast("That bill could not be found.", "err"); return; }
  if (!isAdmin()) { toast("Only admins can review transport bills.", "err"); return; }
  const pendingNow = b.status === "pending";
 
  openModal(`
    <div class="p-6">
      <div class="flex items-start justify-between gap-3 mb-5">
        <div>
          <h2 class="text-lg font-semibold">${pendingNow ? "Review transport bill" : "Transport bill"}</h2>
          <p class="text-sm mt-1" style="color:var(--text-muted)">${esc(b.employeeName)}, ${esc(deptName(b.department))}</p>
        </div>
        ${statusBadge(b.status)}
      </div>
 
      ${b.slipUrl ? `<img class="slip-preview mb-4" src="${esc(b.slipUrl)}" onclick="viewSlip('${esc(b.slipUrl)}')" style="cursor:zoom-in" alt="Slip">` : `<div class="empty mb-4" style="padding:20px">No slip attached.</div>`}
 
      <div class="detail-grid">
        <div><div class="k">Route</div><div class="v">${esc(b.fromLocation)} → ${esc(b.toLocation)}</div></div>
        <div><div class="k">Amount</div><div class="v amount">Rs. ${fmtMoney(b.amount)}</div></div>
        <div><div class="k">Distance</div><div class="v">${b.distanceKm != null ? b.distanceKm + " km" : "Not given"}</div></div>
        <div><div class="k">Date</div><div class="v">${fmtDate(b.date)}</div></div>
        <div><div class="k">Project</div><div class="v">${esc(projectName(b.project))}</div></div>
        <div><div class="k">Submitted</div><div class="v">${fmtDateTime(b.appliedAt)}</div></div>
      </div>
 
      <div class="k mb-1">Additional details</div>
      <div class="surface-alt rounded-lg p-3 text-sm mb-4">${esc(b.notes) || "No additional details given."}</div>
 
      ${!pendingNow ? `<p class="text-sm mb-4" style="color:var(--text-muted)">${b.status === "approved" ? "Approved" : "Rejected"} by ${esc(b.reviewedByName || "—")} on ${fmtDateTime(b.reviewedAt)}. You can still change the decision below.</p>` : ""}
 
      <label class="label">Note (optional, shown to employee)</label>
      <textarea class="input mb-3" id="bill-review-note" rows="2" placeholder="e.g. Approved for reimbursement in this month's payroll.">${esc(b.adminNote || "")}</textarea>
 
      <p id="bill-review-err" class="form-err"></p>
 
      <div class="flex gap-2 justify-end flex-wrap">
        <button type="button" class="btn btn-secondary" onclick="closeModal()">Close</button>
        ${b.status !== "rejected" ? `<button type="button" data-bill-btn data-decision="rejected" class="btn btn-danger" onclick="handleBillReview('${id}','rejected')">Reject</button>` : ""}
        ${b.status !== "approved" ? `<button type="button" data-bill-btn data-decision="approved" class="btn btn-success" onclick="handleBillReview('${id}','approved')">Approve</button>` : ""}
      </div>
    </div>
  `);
}
 
async function handleBillReview(id, decision) {
  if (billReviewBusy) return;
  if (decision !== "approved" && decision !== "rejected") return;
 
  const b = state.bills.find(x => x.id === id);
  const errEl = $("bill-review-err"), noteEl = $("bill-review-note");
  const btns = Array.from(document.querySelectorAll("#modal-root [data-bill-btn]"));
  if (errEl) errEl.textContent = "";
 
  if (!isAdmin()) { if (errEl) errEl.textContent = "Only admins can review transport bills."; return; }
  if (b && isDeptAdmin() && b.department !== state.user.department) {
    if (errEl) errEl.textContent = "You can only review bills from your own department.";
    return;
  }
 
  billReviewBusy = true;
  btns.forEach(x => { x.disabled = true; });
  const clicked = btns.find(x => x.dataset.decision === decision);
  const oldLabel = clicked ? clicked.textContent : "";
  if (clicked) clicked.textContent = decision === "approved" ? "Approving…" : "Rejecting…";
 
  try {
    await db.collection("transportBills").doc(id).update({
      status: decision,
      adminNote: (noteEl && noteEl.value.trim()) || null,
      reviewedBy: state.user.uid,
      reviewedByName: state.user.name,
      reviewedAt: new Date().toISOString(),
    });
    closeModal();
    toast(`Bill ${decision}.`, "ok");
  } catch (ex) {
    console.error("bill review failed", ex);
    if (errEl) errEl.textContent = friendlyAuthError(ex);
    btns.forEach(x => { x.disabled = false; });
    if (clicked) clicked.textContent = oldLabel;
  } finally {
    billReviewBusy = false;
  }
}
 
function exportBillsExcel() {
  const list = billFilteredList();
  if (list.length === 0) { toast("Nothing to export for these filters.", "err"); return; }
  const rows = list.map(b => ({
    Employee: b.employeeName, Department: deptName(b.department), Date: b.date,
    From: b.fromLocation, To: b.toLocation, "Distance (km)": b.distanceKm ?? "",
    "Amount (Rs.)": Number(b.amount) || 0, Project: projectName(b.project),
    Notes: b.notes || "", Status: b.status,
    "Reviewed By": b.reviewedByName || "", "Admin Note": b.adminNote || "",
    "Slip URL": b.slipUrl || "", "Submitted At": b.appliedAt ? b.appliedAt.slice(0, 16).replace("T", " ") : "",
  }));
  const ws = XLSX.utils.json_to_sheet(rows);
  ws["!cols"] = [{wch:20},{wch:16},{wch:12},{wch:22},{wch:22},{wch:12},{wch:14},{wch:18},{wch:30},{wch:10},{wch:16},{wch:26},{wch:40},{wch:17}];
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Transport Bills");
  const range = (state.billFilter.from || state.billFilter.to)
    ? `_${state.billFilter.from || "start"}_to_${state.billFilter.to || "now"}`
    : "";
  XLSX.writeFile(wb, `transport-bills${range}-${todayISO()}.xlsx`);
}
 
/* ============================= EMPLOYEES ============================= */
function tplEmployees() {
  const list = (isSuperAdmin() ? state.employees : state.employees.filter(e => e.department === state.user.department))
    .slice().sort((a, b) => String(a.name).localeCompare(String(b.name)));
 
  return `
    ${pageHeader("Employees", isSuperAdmin() ? "Everyone across the organization." : `Staff in ${esc(deptName(state.user.department))}.`,
      `<button class="btn btn-secondary" onclick="exportEmployeesExcel()">Export to Excel</button>
       <button class="btn btn-primary" onclick="modalAddEmployee()">+ Add employee</button>`)}
 
    <div class="surface overflow-hidden">
      ${list.length === 0 ? `<div class="empty">No employees yet. Use "Add employee" to create the first one.</div>` : `
      <div class="table-wrap"><table class="rtable">
        <thead><tr><th>Name</th><th>Email</th><th>Phone</th>${isSuperAdmin() ? "<th>Department</th>" : ""}<th>Role</th><th>In the field</th><th></th></tr></thead>
        <tbody>
        ${list.map(e => {
          const open = openSessionFor(e.id);
          return `
          <tr>
            <td data-label="Name"><div class="cell-user"><div class="avatar">${esc(initials(e.name))}</div><span class="font-medium">${esc(e.name)}</span></div></td>
            <td data-label="Email" style="color:var(--text-muted)">${esc(e.email)}</td>
            <td data-label="Phone">${esc(e.phone) || "—"}</td>
            ${isSuperAdmin() ? `<td data-label="Department">${esc(deptName(e.department))}</td>` : ""}
            <td data-label="Role"><span class="badge badge-role">${roleLabel(e.role)}</span></td>
            <td data-label="In the field">${open ? `<span class="badge badge-open">Since ${fmtTime(open.checkInAt)}</span>` : `<span style="color:var(--text-muted)">—</span>`}</td>
            <td class="td-action">${e.id !== state.user.uid ? `<button class="btn btn-sm btn-ghost" style="color:var(--danger)" onclick="handleRemoveEmployee('${e.id}')">Remove</button>` : `<span class="text-xs" style="color:var(--text-muted)">You</span>`}</td>
          </tr>`;
        }).join("")}
        </tbody>
      </table></div>`}
    </div>
  `;
}
 
function modalAddEmployee() {
  if (!isSuperAdmin() && !state.user.department) { toast("Your account has no department yet. Ask your Super Admin.", "err"); return; }
  if (isSuperAdmin() && state.departments.length === 0) { toast("Create a department first.", "err"); return; }
  const deptOptions = isSuperAdmin()
    ? state.departments.map(d => `<option value="${d.id}">${esc(d.name)}</option>`).join("")
    : `<option value="${state.user.department}">${esc(deptName(state.user.department))}</option>`;
 
  openModal(`
    <form onsubmit="handleAddEmployee(event)" class="p-6">
      <h2 class="text-lg font-semibold mb-1">Add employee</h2>
      <p class="text-sm mb-5" style="color:var(--text-muted)">They'll be able to sign in right away with this email & password.</p>
 
      <label class="label">Full name</label>
      <input class="input mb-4" name="name" required placeholder="e.g. Nadeesha Fernando">
 
      <label class="label">Email</label>
      <input class="input mb-4" name="email" type="email" required placeholder="name@company.com">
 
      <label class="label">Phone (optional)</label>
      <input class="input mb-4" name="phone" type="tel" placeholder="e.g. 077 123 4567">
 
      <label class="label">Temporary password</label>
      <input class="input mb-4" name="password" type="text" required minlength="6" placeholder="At least 6 characters">
 
      <div class="grid ${isSuperAdmin() ? "grid-cols-2" : "grid-cols-1"} gap-3 mb-2">
        ${isSuperAdmin() ? `<div><label class="label">Department</label><select class="input" name="department" required>${deptOptions}</select></div>` : `<input type="hidden" name="department" value="${state.user.department}">`}
        <div>
          <label class="label">Role</label>
          <select class="input" name="role">
            <option value="employee">Employee</option>
            ${isSuperAdmin() ? `<option value="deptadmin">Department Admin</option>` : ""}
          </select>
        </div>
      </div>
 
      <div class="flex gap-2 justify-end mt-4">
        <button type="button" class="btn btn-secondary" onclick="closeModal()">Cancel</button>
        <button type="submit" class="btn btn-primary" id="add-emp-btn">Add employee</button>
      </div>
    </form>
  `);
}
 
async function handleAddEmployee(e) {
  e.preventDefault();
  const f = e.target, btn = $("add-emp-btn");
  btn.disabled = true; btn.textContent = "Adding…";
  const name = f.name.value.trim(), email = f.email.value.trim(), password = f.password.value;
  const phone = f.phone.value.trim() || null;
  const department = f.department.value, role = f.role.value;
 
  // Create the auth account on a secondary app instance so the admin's own session isn't replaced.
  const secondary = firebase.initializeApp(firebaseConfig, "Secondary-" + Date.now());
  try {
    const cred = await secondary.auth().createUserWithEmailAndPassword(email, password);
    const uid = cred.user.uid;
    await secondary.auth().signOut();
    await db.collection("employees").doc(uid).set({
      name, email, phone, role, department,
      createdAt: firebase.firestore.FieldValue.serverTimestamp(),
    });
    closeModal();
    toast("Employee added.", "ok");
  } catch (ex) {
    console.error(ex);
    toast(friendlyAuthError(ex), "err");
    btn.disabled = false; btn.textContent = "Add employee";
  } finally {
    secondary.delete().catch(() => {});
  }
}
 
async function handleRemoveEmployee(uid) {
  const emp = empByUid(uid);
  const name = emp ? emp.name : "this employee";
  if (!confirm(`Remove ${name} from the staff list? This won't delete their sign-in account — do that from the Firebase console if needed.`)) return;
  try {
    await db.collection("employees").doc(uid).delete();
    toast("Employee removed.", "ok");
  } catch (ex) { console.error(ex); toast(friendlyAuthError(ex), "err"); }
}
 
function exportEmployeesExcel() {
  const list = isSuperAdmin() ? state.employees : state.employees.filter(e => e.department === state.user.department);
  if (list.length === 0) { toast("No employees to export.", "err"); return; }
  const rows = list.map(e => ({ Name: e.name, Email: e.email, Phone: e.phone || "", Department: deptName(e.department), Role: roleLabel(e.role) }));
  const ws = XLSX.utils.json_to_sheet(rows);
  ws["!cols"] = [{wch:22},{wch:26},{wch:16},{wch:18},{wch:16}];
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Employees");
  XLSX.writeFile(wb, `employees-${todayISO()}.xlsx`);
}
 
/* ============================= DEPARTMENTS ============================= */
function tplDepartments() {
  const list = state.departments.slice().sort((a, b) => String(a.name).localeCompare(String(b.name)));
  return `
    ${pageHeader("Departments", "Add departments before assigning employees to them.",
      `<button class="btn btn-primary" onclick="modalAddDepartment()">+ Add department</button>`)}
    <div class="grid gap-3" style="grid-template-columns:repeat(auto-fill,minmax(220px,1fr))">
      ${list.length === 0 ? `<div class="surface empty" style="grid-column:1/-1">No departments yet. Use "Add department" to create one.</div>` :
        list.map(d => {
          const count = state.employees.filter(e => e.department === d.id).length;
          return `
          <div class="surface p-5">
            <div class="flex items-start justify-between gap-2">
              <div class="font-semibold">${esc(d.name)}</div>
              <button class="btn btn-sm btn-ghost" style="color:var(--danger)" onclick="handleDeleteDepartment('${d.id}')">Remove</button>
            </div>
            <div class="text-sm mt-1" style="color:var(--text-muted)">${count} employee${count === 1 ? "" : "s"}</div>
          </div>`;
        }).join("")}
    </div>
  `;
}
 
function modalAddDepartment() {
  openModal(`
    <form onsubmit="handleAddDepartment(event)" class="p-6">
      <h2 class="text-lg font-semibold mb-1">Add department</h2>
      <p class="text-sm mb-5" style="color:var(--text-muted)">e.g. Sales, Warehouse, HR, Finance.</p>
      <label class="label">Department name</label>
      <input class="input mb-5" name="name" required autofocus placeholder="e.g. Sales & Marketing">
      <div class="flex gap-2 justify-end">
        <button type="button" class="btn btn-secondary" onclick="closeModal()">Cancel</button>
        <button type="submit" class="btn btn-primary" id="add-dept-btn">Add department</button>
      </div>
    </form>
  `);
}
 
async function handleAddDepartment(e) {
  e.preventDefault();
  const f = e.target, btn = $("add-dept-btn");
  btn.disabled = true; btn.textContent = "Adding…";
  try {
    await db.collection("departments").add({ name: f.name.value.trim(), createdAt: firebase.firestore.FieldValue.serverTimestamp() });
    closeModal();
    toast("Department added.", "ok");
  } catch (ex) {
    console.error(ex); toast(friendlyAuthError(ex), "err");
    btn.disabled = false; btn.textContent = "Add department";
  }
}
 
async function handleDeleteDepartment(id) {
  const dept = state.departments.find(d => d.id === id);
  if (!dept) return;
  const count = state.employees.filter(e => e.department === id).length;
  if (count > 0) { toast(`Move or remove the ${count} employee(s) in ${dept.name} first.`, "err"); return; }
  if (!confirm(`Remove department "${dept.name}"?`)) return;
  try {
    await db.collection("departments").doc(id).delete();
    toast("Department removed.", "ok");
  } catch (ex) { console.error(ex); toast(friendlyAuthError(ex), "err"); }
}
 
/* ============================= PROJECTS ============================= */
function tplProjects() {
  const list = state.projects.slice().sort((a, b) => String(a.name).localeCompare(String(b.name)));
  return `
    ${pageHeader("Projects", "Add projects so employees can tag transport bills against them.",
      `<button class="btn btn-primary" onclick="modalAddProject()">+ Add project</button>`)}
    <div class="grid gap-3" style="grid-template-columns:repeat(auto-fill,minmax(220px,1fr))">
      ${list.length === 0 ? `<div class="surface empty" style="grid-column:1/-1">No projects yet. Use "Add project" to create one.</div>` :
        list.map(p => {
          const count = state.bills.filter(b => b.project === p.id).length;
          return `
          <div class="surface p-5">
            <div class="flex items-start justify-between gap-2">
              <div class="font-semibold">${esc(p.name)}</div>
              <button class="btn btn-sm btn-ghost" style="color:var(--danger)" onclick="handleDeleteProject('${p.id}')">Remove</button>
            </div>
            <div class="text-sm mt-1" style="color:var(--text-muted)">${count} transport bill${count === 1 ? "" : "s"}</div>
          </div>`;
        }).join("")}
    </div>
  `;
}
 
function modalAddProject() {
  openModal(`
    <form onsubmit="handleAddProject(event)" class="p-6">
      <h2 class="text-lg font-semibold mb-1">Add project</h2>
      <p class="text-sm mb-5" style="color:var(--text-muted)">e.g. Client name, internal project code.</p>
      <label class="label">Project name</label>
      <input class="input mb-5" name="name" required autofocus placeholder="e.g. Acme Corp Rollout">
      <div class="flex gap-2 justify-end">
        <button type="button" class="btn btn-secondary" onclick="closeModal()">Cancel</button>
        <button type="submit" class="btn btn-primary" id="add-project-btn">Add project</button>
      </div>
    </form>
  `);
}
 
async function handleAddProject(e) {
  e.preventDefault();
  const f = e.target, btn = $("add-project-btn");
  btn.disabled = true; btn.textContent = "Adding…";
  try {
    await db.collection("projects").add({ name: f.name.value.trim(), createdAt: firebase.firestore.FieldValue.serverTimestamp() });
    closeModal();
    toast("Project added.", "ok");
  } catch (ex) {
    console.error(ex); toast(friendlyAuthError(ex), "err");
    btn.disabled = false; btn.textContent = "Add project";
  }
}
 
async function handleDeleteProject(id) {
  const proj = state.projects.find(p => p.id === id);
  if (!proj) return;
  const count = state.bills.filter(b => b.project === id).length;
  if (count > 0) { toast(`${count} transport bill(s) still use "${proj.name}". Remove or reassign them first.`, "err"); return; }
  if (!confirm(`Remove project "${proj.name}"?`)) return;
  try {
    await db.collection("projects").doc(id).delete();
    toast("Project removed.", "ok");
  } catch (ex) { console.error(ex); toast(friendlyAuthError(ex), "err"); }
}
 
/* ============================= MY PROFILE ============================= */
function tplProfile() {
  const u = state.user;
  const myLeaves = state.leaves.filter(l => l.employeeId === u.uid);
  const mySessions = state.attendance.filter(a => a.employeeId === u.uid);
  const myBills = myBillsScope();
 
  return `
    ${pageHeader("My Profile", "Your details, and the settings you can change yourself.")}
 
    <div class="surface p-5 mb-6 flex items-center gap-4 flex-wrap">
      <div class="avatar avatar-lg">${esc(initials(u.name))}</div>
      <div class="min-w-0">
        <div class="text-lg font-semibold">${esc(u.name)}</div>
        <div class="text-sm" style="color:var(--text-muted)">${esc(u.email)}</div>
        <div class="mt-2 flex gap-2 flex-wrap">
          <span class="badge badge-role">${roleLabel(u.role)}</span>
          ${u.department ? `<span class="badge badge-closed">${esc(deptName(u.department))}</span>` : ""}
        </div>
      </div>
      <button class="btn btn-primary ml-auto" onclick="modalEditProfile()">Edit details</button>
    </div>
 
    <div class="grid gap-4 mb-6" style="grid-template-columns:repeat(auto-fit,minmax(170px,1fr))">
      <div class="surface stat" style="--stat:var(--navy)">
        <div class="text-sm mb-2" style="color:var(--text-muted)">Leave requests filed</div>
        <div class="text-3xl font-semibold" style="color:var(--navy)">${myLeaves.length}</div>
      </div>
      <div class="surface stat" style="--stat:var(--success)">
        <div class="text-sm mb-2" style="color:var(--text-muted)">Approved leave days</div>
        <div class="text-3xl font-semibold" style="color:var(--success)">${myLeaves.filter(l => l.status === "approved").reduce((s, l) => s + (l.totalDays || 0), 0)}</div>
      </div>
      <div class="surface stat" style="--stat:var(--navy)">
        <div class="text-sm mb-2" style="color:var(--text-muted)">Field sessions</div>
        <div class="text-3xl font-semibold" style="color:var(--navy)">${mySessions.length}</div>
      </div>
      <div class="surface stat" style="--stat:var(--navy)">
        <div class="text-sm mb-2" style="color:var(--text-muted)">Transport bills filed</div>
        <div class="text-3xl font-semibold" style="color:var(--navy)">${myBills.length}</div>
      </div>
    </div>
 
    <div class="surface overflow-hidden mb-6">
      <div class="px-5 py-4" style="border-bottom:1px solid var(--border)"><div class="font-semibold text-sm">Details</div></div>
      <div class="kv-row"><span class="kv-k">Full name</span><span class="kv-v">${esc(u.name)}</span></div>
      <div class="kv-row"><span class="kv-k">Job title</span><span class="kv-v">${esc(u.designation) || "Not set"}</span></div>
      <div class="kv-row"><span class="kv-k">Phone</span><span class="kv-v">${esc(u.phone) || "Not set"}</span></div>
      <div class="kv-row"><span class="kv-k">Email</span><span class="kv-v">${esc(u.email)}</span></div>
      <div class="kv-row"><span class="kv-k">Department</span><span class="kv-v">${esc(u.department ? deptName(u.department) : "Not assigned")}</span></div>
      <div class="kv-row"><span class="kv-k">Role</span><span class="kv-v">${roleLabel(u.role)}</span></div>
    </div>
 
    <div class="surface overflow-hidden">
      <div class="px-5 py-4" style="border-bottom:1px solid var(--border)"><div class="font-semibold text-sm">Account</div></div>
      <div class="kv-row">
        <div>
          <div class="font-medium">Password</div>
          <div class="text-sm" style="color:var(--text-muted)">Change the password you sign in with.</div>
        </div>
        <button class="btn btn-sm btn-secondary shrink-0" onclick="modalChangePassword()">Change</button>
      </div>
      <div class="kv-row">
        <div>
          <div class="font-medium">Sign out</div>
          <div class="text-sm" style="color:var(--text-muted)">End this session on this device.</div>
        </div>
        <button class="btn btn-sm btn-ghost shrink-0" style="color:var(--danger)" onclick="handleLogout()">Sign out</button>
      </div>
      <div class="kv-row">
        <div>
          <div class="font-medium">Email & department</div>
          <div class="text-sm" style="color:var(--text-muted)">Only an admin can change these for you.</div>
        </div>
        <span class="text-sm shrink-0" style="color:var(--text-muted)">Locked</span>
      </div>
    </div>
  `;
}
 
function modalEditProfile() {
  const u = state.user;
  openModal(`
    <form onsubmit="handleSaveProfile(event)" class="p-6">
      <h2 class="text-lg font-semibold mb-1">Edit my details</h2>
      <p class="text-sm mb-5" style="color:var(--text-muted)">This is what your admin and colleagues see on your requests.</p>
 
      <label class="label">Full name</label>
      <input class="input mb-4" name="name" required value="${esc(u.name)}">
 
      <label class="label">Job title (optional)</label>
      <input class="input mb-4" name="designation" value="${esc(u.designation || "")}" placeholder="e.g. Sales Executive">
 
      <label class="label">Phone (optional)</label>
      <input class="input mb-4" name="phone" type="tel" value="${esc(u.phone || "")}" placeholder="e.g. 077 123 4567">
 
      <p id="profile-err" class="form-err"></p>
 
      <div class="flex gap-2 justify-end">
        <button type="button" class="btn btn-secondary" onclick="closeModal()">Cancel</button>
        <button type="submit" class="btn btn-primary" id="profile-btn">Save changes</button>
      </div>
    </form>
  `);
}
 
async function handleSaveProfile(e) {
  e.preventDefault();
  const f = e.target, btn = $("profile-btn"), errEl = $("profile-err");
  const name = f.name.value.trim();
  if (errEl) errEl.textContent = "";
  if (!name) { if (errEl) errEl.textContent = "Enter your full name."; return; }
  btn.disabled = true; btn.textContent = "Saving…";
  try {
    await db.collection("employees").doc(state.user.uid).update({
      name,
      designation: f.designation.value.trim() || null,
      phone: f.phone.value.trim() || null,
    });
    state.user.name = name;
    state.user.designation = f.designation.value.trim() || null;
    state.user.phone = f.phone.value.trim() || null;
    closeModal();
    toast("Profile updated.", "ok");
    render();
  } catch (ex) {
    console.error(ex);
    if (errEl) errEl.textContent = friendlyAuthError(ex);
    btn.disabled = false; btn.textContent = "Save changes";
  }
}
 
function modalChangePassword() {
  openModal(`
    <form onsubmit="handleChangePassword(event)" class="p-6">
      <h2 class="text-lg font-semibold mb-1">Change password</h2>
      <p class="text-sm mb-5" style="color:var(--text-muted)">Enter your current password, then the new one.</p>
      <label class="label">Current password</label>
      <input class="input mb-4" type="password" name="current" required>
      <label class="label">New password</label>
      <input class="input mb-2" type="password" name="next" required minlength="6">
      <p class="text-xs mb-4" style="color:var(--text-muted)">Minimum 6 characters.</p>
      <p id="pw-err" class="form-err"></p>
      <div class="flex gap-2 justify-end">
        <button type="button" class="btn btn-secondary" onclick="closeModal()">Cancel</button>
        <button type="submit" class="btn btn-primary" id="pw-btn">Update password</button>
      </div>
    </form>
  `);
}
 
async function handleChangePassword(e) {
  e.preventDefault();
  const f = e.target, btn = $("pw-btn"), err = $("pw-err");
  err.textContent = "";
  btn.disabled = true; btn.textContent = "Updating…";
  try {
    const cred = firebase.auth.EmailAuthProvider.credential(state.user.email, f.current.value);
    await auth.currentUser.reauthenticateWithCredential(cred);
    await auth.currentUser.updatePassword(f.next.value);
    closeModal();
    toast("Password updated.", "ok");
  } catch (ex) {
    console.error(ex);
    err.textContent = friendlyAuthError(ex);
    btn.disabled = false; btn.textContent = "Update password";
  }
}
// kept so any older reference still works
function modalProfile() { switchTab("profile"); }
 
/* ============================= FIREBASE LISTENERS ============================= */
function attachListeners() {
  detachListeners();
 
  state.unsubs.push(db.collection("departments").onSnapshot(snap => {
    state.departments = snap.docs.map(d => ({ id: d.id, ...d.data() }));
    render();
  }, err => console.error("departments listener", err)));
 
  state.unsubs.push(db.collection("projects").onSnapshot(snap => {
    state.projects = snap.docs.map(d => ({ id: d.id, ...d.data() }));
    render();
  }, err => console.error("projects listener", err)));
 
  // Only admins need the staff list (regular employees never see it)
  if (isAdmin()) {
    const empQuery = isSuperAdmin() ? db.collection("employees") : db.collection("employees").where("department", "==", state.user.department);
    state.unsubs.push(empQuery.onSnapshot(snap => {
      state.employees = snap.docs.map(d => ({ id: d.id, ...d.data() }));
      render();
    }, err => console.error("employees listener", err)));
  }
 
  let leaveQuery;
  if (isSuperAdmin()) leaveQuery = db.collection("leaveRequests");
  else if (isDeptAdmin()) leaveQuery = db.collection("leaveRequests").where("department", "==", state.user.department);
  else leaveQuery = db.collection("leaveRequests").where("employeeId", "==", state.user.uid);
  state.unsubs.push(leaveQuery.onSnapshot(snap => {
    state.leaves = snap.docs.map(d => ({ id: d.id, ...d.data() }));
    render();
  }, err => console.error("leaves listener", err)));
 
  let attQuery;
  if (isSuperAdmin()) attQuery = db.collection("attendance");
  else if (isDeptAdmin()) attQuery = db.collection("attendance").where("department", "==", state.user.department);
  else attQuery = db.collection("attendance").where("employeeId", "==", state.user.uid);
  state.unsubs.push(attQuery.onSnapshot(snap => {
    state.attendance = snap.docs.map(d => ({ id: d.id, ...d.data() }));
    render();
  }, err => console.error("attendance listener", err)));
 
  let billQuery;
  if (isSuperAdmin()) billQuery = db.collection("transportBills");
  else if (isDeptAdmin()) billQuery = db.collection("transportBills").where("department", "==", state.user.department);
  else billQuery = db.collection("transportBills").where("employeeId", "==", state.user.uid);
  state.unsubs.push(billQuery.onSnapshot(snap => {
    state.bills = snap.docs.map(d => ({ id: d.id, ...d.data() }));
    render();
  }, err => console.error("transport bills listener", err)));
}
function detachListeners() {
  state.unsubs.forEach(u => { try { u(); } catch (e) {} });
  state.unsubs = [];
  state.departments = []; state.projects = []; state.employees = []; state.leaves = []; state.attendance = []; state.bills = [];
  if (clockTimer) { clearInterval(clockTimer); clockTimer = null; }
}
 
/* ============================= BOOT ============================= */
async function loadUser(fbUser) {
  try {
    const doc = await db.collection("employees").doc(fbUser.uid).get();
    if (!doc.exists) {
      state.booting = false; state.view = "login";
      toast("Your account isn't set up as staff yet. Contact your admin.", "err");
      await auth.signOut();
      return render();
    }
    const data = doc.data();
    state.user = {
      uid: fbUser.uid, email: fbUser.email, name: data.name, role: data.role,
      department: data.department || null, phone: data.phone || null, designation: data.designation || null,
    };
    state.booting = false;
    state.view = "app";
    attachListeners();
    render();
  } catch (ex) {
    console.error(ex);
    state.booting = false; state.view = "login";
    toast(friendlyAuthError(ex), "err");
    render();
  }
}
 
async function boot() {
  try {
    const statusDoc = await db.collection("system").doc("status").get();
    state.setupNeeded = !statusDoc.exists;
  } catch (ex) {
    console.error("status check failed", ex);
  }
 
  auth.onAuthStateChanged(async (fbUser) => {
    if (state.creatingAdmin) return; // setup flow handles its own state
    closeModal();
    if (!fbUser) {
      detachListeners();
      state.user = null;
      state.booting = false;
      state.view = state.setupNeeded ? "setup" : "login";
      return render();
    }
    await loadUser(fbUser);
  });
}
 
/* Inline onclick/onsubmit handlers in the templates need these on window (ES modules are not global) */
Object.assign(window, {
  switchTab, handleLogout, render, state, closeModal,
  modalProfile, modalEditProfile, handleSaveProfile, modalChangePassword, handleChangePassword,
  modalSubmitLeave, handleSubmitLeave, handleCancelLeave,
  modalReview, handleReview, exportLeavesExcel,
  modalPunch, handlePunch, exportAttendanceExcel,
  exportEmployeesExcel, modalAddEmployee, handleAddEmployee, handleRemoveEmployee,
  modalAddDepartment, handleAddDepartment, handleDeleteDepartment,
  modalSubmitBill, handleSubmitBill, handleSlipPick, handleCancelBill, viewSlip,
  modalBillReview, handleBillReview, exportBillsExcel,
  modalAddProject, handleAddProject, handleDeleteProject,
  handleLogin, handleSetupSubmit,
});
 
injectExtraCSS();
render();
boot();
 
