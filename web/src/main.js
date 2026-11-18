// 화면 ― 원가계산서 제비율 스마트 체크와 같은 틀(남색 머리띠 · 왼쪽 메뉴 · 흰 카드).
// 엑셀판의 입력정보 / 출력관리 시트를 메뉴 화면으로 옮겼다.
import JSZip from "jszip";
import { loadLocal, saveLocal, emptyData, normalize } from "./store/localStore.js";
import {
  linkSupported, loadLink, clearLink, permission, pickNew, pickExisting, fileTarget,
} from "./store/fileLink.js";
import {
  googleReady, loadGoogleMeta, clearGoogleMeta, googleTarget, signIn, signOut, whoAmI, tokenValid, loadGis,
  redirectSignIn, takeRedirectResult,
} from "./store/googleDrive.js";
import { createSync } from "./sync.js";
import { NEIS_KEY } from "./config.js";
import {
  refreshAuto, WORK_KEYS, WORK_NAMES, FORMS, planKind, progressOf,
  orgPositionDefaults, CONTRACT_TYPES,
} from "./domain/dictionary.js";
import { contractDays, WEEKDAYS, todayDay, dayToIso, toDay, parts } from "./domain/schedule.js";
import { Exporter, formMarked } from "./export.js";
import {
  PHASES, GROUP_LABEL, dutiesOf, stateOf, summaryOf, schoolYearOf, schoolYearRange, halfRange,
} from "./domain/duties.js";
import {
  cloneContract, rollForward, isAnnual, vendorBook, VENDOR_KEYS, shiftNameYears, shiftYears,
} from "./domain/rollover.js";

// ---------------------------------------------------------------- 상태
let data = emptyData();
let tab = "org";
let selId = null;
const checked = new Set(); // 출력 화면에서 고른 계약
const openGroups = new Set(); // 펼쳐 둔 정기 점검 묶음 "계약id:묶음"
let boardYear = null; // 진행 현황 일정표의 학년도
let rep = null; // 반기 이행 현황표 { y, h }

const $ = (sel, root = document) => root.querySelector(sel);
const esc = (v) => String(v ?? "").replace(/[&<>"']/g, (ch) =>
  ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[ch]));
const uid = () => Math.random().toString(36).slice(2, 10) + Date.now().toString(36);
const todayIso = () => dayToIso(todayDay());

function toast(msg, ms = 3500) {
  const t = $("#toast");
  t.textContent = msg;
  t.classList.add("show");
  clearTimeout(toast.timer);
  toast.timer = setTimeout(() => t.classList.remove("show"), ms);
}

function download(name, bytes, type) {
  const url = URL.createObjectURL(new Blob([bytes], { type }));
  const a = document.createElement("a");
  a.href = url; a.download = name;
  document.body.appendChild(a); a.click(); a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 10000);
}

// 버튼 고르는 창. buttons = [{ label, value, cls }] → 고른 value (Esc 는 null)
function choose(title, bodyHtml, buttons) {
  const dlg = $("#dlg");
  if (dlg.open) dlg.close();
  dlg.innerHTML = `<h2>${esc(title)}</h2><div class="dlg-body">${bodyHtml}</div>
    <div class="dlg-btns">${buttons.map((b, i) =>
      `<button class="btn ${b.cls || "btn-outline"}" data-i="${i}">${esc(b.label)}</button>`).join("")}</div>`;
  return new Promise((resolve) => {
    const done = (v) => { dlg.close(); resolve(v); };
    dlg.querySelectorAll("button[data-i]").forEach((b) =>
      b.addEventListener("click", () => done(buttons[+b.dataset.i].value)));
    dlg.oncancel = (e) => { e.preventDefault(); done(null); };
    dlg.showModal();
  });
}

// ---------------------------------------------------------------- 저장
// 1) 고칠 때마다 이 PC 브라우저(IndexedDB)에 저장
// 2) 자동 저장 대상을 켰으면 거기에도 ― 구글 드라이브(앱 숨김 폴더) 또는 내 컴퓨터 파일
let saveTimer = null;
let savePending = false;
let localOk = true;

function changed() {
  data.updatedAt = new Date().toISOString();
  savePending = true;
  clearTimeout(saveTimer);
  saveTimer = setTimeout(async () => {
    localOk = await saveLocal(data);
    savePending = false;
    sync.queue("write");
    showSaveState();
  }, 300);
}

function adopt(next) {
  data = next;
  data.contracts.forEach((c) => { c.id ??= uid(); refreshAuto(c); });
  selId = null; checked.clear();
}

function backupDownload() {
  data.backedUpAt = new Date().toISOString();
  const name = `중대재해예방서식_${(data.org.기관명 || "내정보").replace(/[\\/:*?"<>|·]/g, "")}_${todayIso()}.json`;
  download(name, JSON.stringify(data, null, 1), "application/json");
  saveLocal(data).then((ok) => { localOk = ok; showSaveState(); });
}

const sync = createSync({
  getData: () => data,
  adopt,
  persistLocal: async () => { localOk = await saveLocal(data); },
  render: () => render(),
  choose,
  toast,
  backup: backupDownload,
  onState: () => showSaveState(),
});

const hhmm = (iso) => iso ? new Date(iso).toLocaleTimeString("ko-KR", { hour: "2-digit", minute: "2-digit" }) : "";

// 머리띠 오른쪽 단추 + 화면 위 안내 띠
function showSaveState() {
  const { target, state, error } = sync.s;
  const chip = $("#syncChip");
  let cls = "", label = "", sub = "";
  if (!localOk) { cls = "is-warn"; label = "브라우저 저장 실패"; sub = "파일로 받아 두세요"; }
  else if (!target) { label = "이 컴퓨터에만 저장"; sub = "이어쓰기 켜기"; }
  else if (state === "ok") {
    cls = "is-ok";
    label = target.kind === "google" ? "구글 드라이브 저장됨" : "저장 파일에 저장됨";
    sub = hhmm(target.meta.savedAt);
  } else if (state === "busy") { cls = "is-ok"; label = "저장 중…"; }
  else if (state === "prompt") { cls = "is-warn"; label = target.kind === "google" ? "구글 계정으로 이어서 하기" : "자동 저장 다시 켜기"; }
  else { cls = "is-warn"; label = "자동 저장 멈춤"; }
  chip.className = `topbar-btn ${cls}`;
  chip.innerHTML = `<span class="dot"></span>${esc(label)}${sub ? ` <span class="sub">${esc(sub)}</span>` : ""}`;
  chip.title = target ? `${target.where} · ${target.name}` : "저장 · 이어쓰기";

  const banner = $("#banner");
  if (banner) {
    if (target && state === "prompt") {
      const g = target.kind === "google";
      banner.innerHTML = `<div class="alert is-warn alert-row">
        <span>${g ? `<b>구글 드라이브(${esc(target.name)})</b>와 아직 이어지지 않았습니다.` : `<b>저장 파일(${esc(target.name)})</b>에 자동 저장이 멈춰 있습니다.`}
          고친 내용은 이 컴퓨터에 저장되고 있습니다.</span>
        <button class="btn btn-primary btn-sm" data-sync="reconnect">${g ? "구글 계정으로 이어서 하기" : "자동 저장 다시 켜기"}</button></div>`;
    } else if (target && state === "error") {
      banner.innerHTML = `<div class="alert is-error alert-row"><span>자동 저장하지 못했습니다 — ${esc(error)}</span>
        <button class="btn btn-sm" data-sync="reconnect">다시 시도</button></div>`;
    } else if (!localOk) {
      banner.innerHTML = `<div class="alert is-error">이 브라우저에 저장하지 못했습니다(사생활 보호 창일 수 있습니다). [저장 · 이어쓰기]에서 파일로 받아 두세요.</div>`;
    } else banner.innerHTML = "";
  }
  if (tab === "save") { const p = $("#savePage"); if (p) p.outerHTML = viewSaveBody(); }
}

// ---------------------------------------------------------------- 자동 저장 켜고 끄기
async function googleConnect() {
  if (!googleReady()) { toast("구글 계정 이어쓰기는 준비 중입니다."); return; }
  try {
    await signIn(null);
  } catch (err) {
    if (err.code === "popup") offerRedirect(null); else toast(err.message, 6000);
    return;
  }
  await finishGoogleConnect();
}

// 로그인(팝업이든 이 창에서 넘어갔다 왔든) 뒤 ― 계정 확인하고 그 드라이브와 맞춘다
async function finishGoogleConnect() {
  authedThisSession = true;
  try {
    const email = await whoAmI();
    await clearLink(); // 파일 자동 저장을 쓰고 있었다면 구글로 바꾼다
    const prev = await loadGoogleMeta();
    const meta = prev?.email === email ? prev : { email, fileId: null, syncedAt: null, savedAt: null };
    const t = googleTarget(meta);
    await t.saveMeta();
    const wasEmpty = !data.org.기관명 && !data.contracts.length;
    if (await sync.attach(t)) {
      const pulled = wasEmpty && (data.org.기관명 || data.contracts.length);
      toast(pulled
        ? `구글 드라이브(${email})에서 불러왔습니다 — 계약 ${data.contracts.length}건.\n이제부터 고칠 때마다 자동 저장됩니다.`
        : `이제부터 ${email} 구글 드라이브에 자동 저장됩니다.\n다른 컴퓨터에서도 이 계정으로 로그인하면 이어서 쓸 수 있습니다.`, 6000);
    }
  } catch (err) {
    toast(err.message, 6000);
  }
}

// 팝업이 막힌 브라우저(차단 설정 · 확장 프로그램 · 기관 PC 정책): 이 창에서 구글로 넘어갔다 돌아온다
let leaving = false;
let popupBlocked = false;
async function offerRedirect(hint) {
  popupBlocked = true;
  const pick = await choose("로그인 창이 막혀 있습니다",
    `<p>이 브라우저가 구글 로그인 창(팝업)을 열지 못했습니다.
      대신 <b>이 창에서</b> 구글 로그인 화면으로 넘어갔다가, 로그인이 끝나면 이 화면으로 돌아옵니다.</p>
     <p class="hint">입력한 내용은 이 컴퓨터에 저장되어 있어 그대로 남습니다.</p>`,
    [
      { label: "이 창에서 구글 로그인", value: "go", cls: "btn-primary" },
      { label: "닫기", value: null },
    ]);
  if (pick !== "go") return;
  clearTimeout(saveTimer);
  if (savePending) { localOk = await saveLocal(data); savePending = false; }
  leaving = true;
  redirectSignIn(hint);
}

async function googleOff() {
  const pick = await choose("구글 계정 로그아웃",
    `<p>이 컴퓨터에서 구글 드라이브 자동 저장을 끕니다. 구글 드라이브에 저장된 내용은 그대로 남아,
      다시 로그인하면 이어서 쓸 수 있습니다.</p>
     <p class="hint">여럿이 함께 쓰는 컴퓨터라면 [로그아웃하고 이 컴퓨터 내용 지우기]를 고르세요.</p>`,
    [
      { label: "로그아웃", value: "out", cls: "btn-primary" },
      { label: "로그아웃하고 이 컴퓨터 내용 지우기", value: "wipe", cls: "btn-danger" },
      { label: "닫기", value: null },
    ]);
  if (!pick) return;
  await sync.queue("write"); // 남은 저장을 마저 올린다
  signOut();
  await clearGoogleMeta();
  sync.detach();
  if (pick === "wipe") {
    data = emptyData(); selId = null; checked.clear();
    localOk = await saveLocal(data);
    tab = "org";
  }
  render();
  toast(pick === "wipe" ? "로그아웃하고 이 컴퓨터의 내용을 지웠습니다." : "로그아웃했습니다.");
}

async function fileConnect(how) {
  let handle;
  try {
    handle = how === "new" ? await pickNew() : await pickExisting();
  } catch (err) {
    if (err?.name !== "AbortError") toast("파일을 고르지 못했습니다: " + err.message);
    return;
  }
  if ((await permission(handle, true)) !== "granted") { toast("파일에 저장하도록 허락해야 자동 저장이 켜집니다."); return; }
  if (sync.s.target?.kind === "google") { signOut(); await clearGoogleMeta(); }
  const link = { handle, name: handle.name, syncedAt: null, savedAt: null };
  const t = fileTarget(link);
  await t.saveMeta();
  if (await sync.attach(t)) toast(`이제부터 고칠 때마다 '${handle.name}'에 자동 저장됩니다.`);
}

async function fileOff() {
  await clearLink();
  sync.detach();
  toast("파일 자동 저장을 껐습니다. 파일은 그대로 남아 있습니다.");
}

// 끊긴 대상을 다시 잇는다(버튼 누름 안에서)
async function reconnect() {
  const t = sync.s.target;
  if (!t) return;
  if (t.kind === "google") {
    try {
      await signIn(t.meta.email);
      authedThisSession = true;
      if (!t.meta.email) { t.meta.email = await whoAmI(); await t.saveMeta(); }
    } catch (err) {
      if (err.code === "popup") offerRedirect(t.meta.email); else toast(err.message, 6000);
      return;
    }
    if (await sync.attach(t)) toast("구글 드라이브와 이어졌습니다.");
  } else {
    if ((await permission(t.meta.handle, true)) !== "granted") { toast("파일에 저장하도록 허락해 주세요."); return; }
    await sync.attach(t);
    if (sync.s.state === "error" && sync.s.error.includes("찾지 못했")) {
      toast(`${sync.s.error}\n[저장 · 이어쓰기]에서 파일을 다시 골라 주세요.`, 8000);
      await clearLink(); sync.detach();
    } else if (sync.s.state === "ok") toast("자동 저장을 다시 켰습니다.");
  }
}

async function importFile(f) {
  try {
    const next = normalize(JSON.parse(await f.text()));
    if (data.contracts.length && !confirm(`지금 입력된 내용(계약 ${data.contracts.length}건)을 이 파일 내용으로 바꿉니다. 계속할까요?`)) return;
    adopt(next);
    changed(); render();
    toast(`불러왔습니다 — 계약 ${data.contracts.length}건`);
  } catch (err) {
    toast("불러오지 못했습니다: " + err.message);
  }
}

$("#syncChip").addEventListener("click", () => {
  const st = sync.s.state;
  if (sync.s.target && (st === "prompt" || st === "error")) { reconnect(); return; }
  tab = "save"; render();
});

// 일하는 중에 구글 로그인(1시간)이 끝나 가면, 다음 클릭 때 조용히 새로 받는다
let authedThisSession = false;
document.addEventListener("pointerdown", (e) => {
  const t = sync.s.target;
  if (!t || t.kind !== "google" || !authedThisSession || popupBlocked) return;
  if (e.target.closest?.("[data-sync], #syncChip, dialog")) return;
  const st = sync.s.state;
  if ((st === "ok" && !tokenValid(5 * 60 * 1000)) || (st === "prompt" && t.meta.email)) {
    signIn(t.meta.email).then(() => sync.attach(t)).catch((err) => { if (err.code === "popup") popupBlocked = true; });
  }
}, true);

// 다른 컴퓨터에서 고쳤을 수 있으니 창으로 돌아오면 다시 본다
document.addEventListener("visibilitychange", () => {
  if (document.visibilityState === "visible") sync.queue("check");
});

// 저장이 덜 끝났는데 창을 닫으려 하면 붙잡는다
window.addEventListener("beforeunload", (e) => {
  if (leaving) return;
  if (savePending || sync.s.state === "busy") { e.preventDefault(); e.returnValue = ""; }
});

// ---------------------------------------------------------------- 메뉴
const ICONS = {
  building: `<rect x="4" y="2" width="16" height="20" rx="2"/><path d="M9 22v-4h6v4"/><path d="M8 6h.01M16 6h.01M12 6h.01M12 10h.01M12 14h.01M16 10h.01M16 14h.01M8 10h.01M8 14h.01"/>`,
  file: `<path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><polyline points="14 2 14 8 20 8"/><line x1="16" y1="13" x2="8" y2="13"/><line x1="16" y1="17" x2="8" y2="17"/>`,
  download: `<path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><polyline points="7 10 12 15 17 10"/><line x1="12" y1="15" x2="12" y2="3"/>`,
  cloud: `<path d="M17.5 19H9a7 7 0 1 1 6.71-9h1.79a4.5 4.5 0 1 1 0 9z"/>`,
  book: `<path d="M4 19.5A2.5 2.5 0 0 1 6.5 17H20"/><path d="M6.5 2H20v20H6.5A2.5 2.5 0 0 1 4 19.5v-15A2.5 2.5 0 0 1 6.5 2z"/>`,
  hdd: `<line x1="22" y1="12" x2="2" y2="12"/><path d="M5.45 5.11 2 12v6a2 2 0 0 0 2 2h16a2 2 0 0 0 2-2v-6l-3.45-6.89A2 2 0 0 0 16.76 4H7.24a2 2 0 0 0-1.79 1.11z"/><line x1="6" y1="16" x2="6.01" y2="16"/><line x1="10" y1="16" x2="10.01" y2="16"/>`,
  search: `<circle cx="11" cy="11" r="7"/><line x1="21" y1="21" x2="16.65" y2="16.65"/>`,
  users: `<path d="M17 21v-2a4 4 0 0 0-4-4H5a4 4 0 0 0-4 4v2"/><circle cx="9" cy="7" r="4"/><path d="M23 21v-2a4 4 0 0 0-3-3.87"/><path d="M16 3.13a4 4 0 0 1 0 7.75"/>`,
  briefcase: `<rect x="2" y="7" width="20" height="14" rx="2"/><path d="M16 21V5a2 2 0 0 0-2-2h-4a2 2 0 0 0-2 2v16"/>`,
  phone: `<path d="M22 16.92v3a2 2 0 0 1-2.18 2 19.8 19.8 0 0 1-8.63-3.07 19.5 19.5 0 0 1-6-6A19.8 19.8 0 0 1 2.12 4.18 2 2 0 0 1 4.11 2h3a2 2 0 0 1 2 1.72c.13.96.36 1.9.7 2.81a2 2 0 0 1-.45 2.11L8.09 9.91a16 16 0 0 0 6 6l1.27-1.27a2 2 0 0 1 2.11-.45c.91.34 1.85.57 2.81.7A2 2 0 0 1 22 16.92z"/>`,
  save: `<path d="M19 21H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h11l5 5v11a2 2 0 0 1-2 2z"/><polyline points="17 21 17 13 7 13 7 21"/><polyline points="7 3 7 8 15 8"/>`,
  chart: `<path d="M3 3v18h18"/><rect x="7" y="12" width="3" height="5" rx="1"/><rect x="12" y="8" width="3" height="9" rx="1"/><rect x="17" y="5" width="3" height="12" rx="1"/>`,
  copy: `<rect x="9" y="9" width="13" height="13" rx="2"/><path d="M5 15H4a2 2 0 0 1-2-2V4a2 2 0 0 1 2-2h9a2 2 0 0 1 2 2v1"/>`,
  forward: `<polyline points="13 17 18 12 13 7"/><polyline points="6 17 11 12 6 7"/>`,
  printer: `<polyline points="6 9 6 2 18 2 18 9"/><path d="M6 18H4a2 2 0 0 1-2-2v-5a2 2 0 0 1 2-2h16a2 2 0 0 1 2 2v5a2 2 0 0 1-2 2h-2"/><rect x="6" y="14" width="12" height="8"/>`,
  list: `<path d="M9 11l3 3L22 4"/><path d="M21 12v7a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h11"/>`,
};
const icon = (name, size = 20, cls = "sidenav-icon") =>
  `<svg class="${cls}" width="${size}" height="${size}" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">${ICONS[name]}</svg>`;

const NAV = [
  { group: "한눈에", items: [{ id: "board", label: "진행 현황", icon: "chart" }] },
  { group: "작성", items: [
    { id: "org", label: "기관정보", icon: "building" },
    { id: "contracts", label: "계약정보", icon: "file" },
    { id: "output", label: "서식 출력", icon: "download" },
  ] },
  { group: "저장", items: [{ id: "save", label: "저장 · 이어쓰기", icon: "cloud" }] },
  { group: "참고", items: [{ id: "help", label: "처음이라면 — 사용법", icon: "book" }] },
];

function renderNav() {
  const late = lateCount();
  const item = (it) => {
    const on = tab === it.id || (tab === "report" && it.id === "board");
    const badge = it.id === "contracts" && data.contracts.length ? `<span class="sidenav-badge">${data.contracts.length}</span>`
      : it.id === "board" && late ? `<span class="sidenav-badge is-late" title="기한이 지났는데 안 한 일">${late}</span>` : "";
    return `<button type="button" class="sidenav-item${on ? " is-active" : ""}" data-tab="${it.id}"${on ? ' aria-current="page"' : ""}>
      ${icon(it.icon)}<span>${it.label}</span>${badge}</button>`;
  };
  $("#sidenav").innerHTML = NAV.map((g) =>
    `<div class="sidenav-group"><div class="sidenav-label">${g.group}</div>${g.items.map(item).join("")}</div>`).join("")
    + `<a class="sidenav-foot" href="./privacy.html" target="_blank" rel="noopener">개인정보처리방침</a>`;
  $("#topnav").innerHTML = NAV.flatMap((g) => g.items).map(item).join("");
}

for (const sel of ["#sidenav", "#topnav"]) {
  $(sel).addEventListener("click", (e) => {
    const b = e.target.closest("[data-tab]");
    if (!b) return;
    tab = b.dataset.tab;
    render();
    window.scrollTo({ top: 0 });
  });
}

const pageHead = (title, desc) =>
  `<div class="page-head"><h1 class="page-title">${title}</h1>${desc ? `<p class="page-desc">${desc}</p>` : ""}</div>`;

function render() {
  renderNav();
  const v = $("#view");
  v.className = "page" + ({ org: " is-mid", contracts: " is-wide", output: " is-wide", board: " is-wide", report: " is-wide" }[tab] || "");
  const body = tab === "board" ? viewBoard()
    : tab === "report" ? viewReport()
    : tab === "org" ? viewOrg()
    : tab === "contracts" ? viewContracts()
    : tab === "output" ? viewOutput()
    : tab === "save" ? viewSave()
    : viewHelp();
  v.innerHTML = `<div id="banner"></div>${body}`;
  showSaveState();
}

// ---------------------------------------------------------------- 저장 · 이어쓰기
function viewSave() {
  if (googleReady()) loadGis().catch(() => {}); // 로그인 창이 바로 뜨게 미리 받아 둔다
  return pageHead("저장 · 이어쓰기",
    "입력한 내용은 늘 <b>이 컴퓨터의 브라우저</b>에 저장됩니다. 다른 날 다른 컴퓨터에서도 이어 쓰려면 아래 중 하나를 켜세요.")
    + viewSaveBody();
}

function viewSaveBody() {
  const { target, state, error } = sync.s;
  const g = target?.kind === "google" ? target : null;
  const f = target?.kind === "file" ? target : null;
  const when = (t) => t.meta.savedAt ? sync.fmtWhen(t.meta.savedAt) : "아직 없음";
  const stLine = (t) => state === "ok" || state === "busy"
    ? `<p class="status-line ok">✓ 자동 저장 중 · 마지막 저장 ${esc(when(t))}</p>`
    : state === "prompt" ? `<p class="status-line warn">이어지지 않음 — 아래 버튼을 한 번 눌러 주세요</p>`
    : `<p class="status-line warn">멈춤 — ${esc(error)}</p>`;

  const googleCard = `
  <section class="card save-card">
    <span class="save-icon">${icon("cloud", 24, "")}</span>
    <div class="body">
      <div class="head"><h2 class="card-title">구글 계정으로 이어쓰기</h2><span class="tag is-info">권장</span></div>
      ${g ? `<p style="margin:0 0 4px"><b>${esc(g.name)}</b> 계정의 구글 드라이브에 저장합니다.</p>${stLine(g)}
        <div class="toolbar" style="margin:0">
          ${state === "prompt" || state === "error" ? `<button class="btn btn-primary" data-sync="reconnect">구글 계정으로 이어서 하기</button>` : ""}
          <button class="btn btn-outline" data-sync="g-off">로그아웃</button>
        </div>`
      : `<p class="hint" style="margin-bottom:6px">어느 컴퓨터에서든 <b>같은 구글 계정으로 로그인</b>하면 같은 내용이 뜹니다. 프로그램을 깔 필요가 없습니다.</p>
        <ul>
          <li>내 구글 드라이브의 <b>이 앱 전용 숨김 칸</b>에 저장합니다. 드라이브의 다른 파일은 보지 않습니다.</li>
          <li>브라우저와 구글이 직접 주고받으므로 <b>개발자에게 아무것도 오지 않습니다</b>.</li>
          <li>교육청(업무용) 구글 계정은 관리자가 막아 둔 경우가 있습니다. 그때는 개인 구글 계정을 쓰세요.</li>
        </ul>
        ${googleReady()
          ? `<button class="btn btn-primary btn-lg" data-sync="g-connect">구글 계정으로 로그인</button>`
          : `<button class="btn btn-primary btn-lg" disabled>준비 중입니다</button>`}`}
    </div>
  </section>`;

  const fileCard = `
  <section class="card save-card">
    <span class="save-icon">${icon("hdd", 24, "")}</span>
    <div class="body">
      <div class="head"><h2 class="card-title">내 컴퓨터 파일에 자동 저장</h2></div>
      ${f ? `<p style="margin:0 0 4px">저장 파일: <b>${esc(f.name)}</b></p>${stLine(f)}
        <div class="toolbar" style="margin:0">
          ${state === "prompt" || state === "error" ? `<button class="btn btn-primary" data-sync="reconnect">자동 저장 다시 켜기</button>` : ""}
          <button class="btn btn-outline" data-sync="f-new">다른 파일로 바꾸기</button>
          <button class="btn btn-outline" data-sync="f-off">자동 저장 끄기</button>
        </div>`
      : `<p class="hint" style="margin-bottom:6px">구글 로그인을 쓸 수 없을 때. 고칠 때마다 고른 파일에 바로 저장됩니다.
          구글 드라이브 데스크톱 폴더(G:)나 USB 에 두면 다른 컴퓨터에서도 [있는 파일 열기]로 이어 쓸 수 있습니다.</p>
        ${linkSupported()
          ? `<div class="toolbar" style="margin:0">
              <button class="btn btn-outline" data-sync="f-new">새 저장 파일 만들기</button>
              <button class="btn btn-outline" data-sync="f-open">있는 파일 열기</button></div>`
          : `<p class="hint">이 브라우저에서는 쓸 수 없습니다(크롬 · 엣지 · 웨일에서 됩니다).</p>`}`}
    </div>
  </section>`;

  const b = data.backedUpAt ? sync.fmtWhen(data.backedUpAt) : "없음";
  const backupCard = `
  <section class="card save-card">
    <span class="save-icon">${icon("save", 24, "")}</span>
    <div class="body">
      <div class="head"><h2 class="card-title">파일로 받아 두기 · 불러오기</h2></div>
      <p class="hint" style="margin-bottom:10px">한 번씩 받아 두면 안심입니다. 다른 컴퓨터에서는 [파일 불러오기]로 엽니다. 마지막으로 받은 때: ${esc(b)}</p>
      <div class="toolbar" style="margin:0">
        <button class="btn btn-outline" data-sync="backup">내 정보 파일로 받기</button>
        <label class="btn btn-outline">파일 불러오기<input id="fileImport" type="file" accept=".json,application/json" hidden /></label>
      </div>
    </div>
  </section>`;

  return `<div id="savePage">${googleCard}${fileCard}${backupCard}
    <div class="alert is-info">이 웹페이지에는 서버가 없습니다. 기관정보 · 계약정보는 이 컴퓨터와, 직접 고른 곳(내 구글 드라이브 · 내 파일)에만 저장됩니다.
      <a href="./privacy.html" target="_blank" rel="noopener">개인정보처리방침</a></div></div>`;
}

// ---------------------------------------------------------------- 사용법
function viewHelp() {
  return pageHead("처음이라면 — 사용법", "계약 정보를 한 번 넣으면, 그 계약에 갖춰야 할 한글(hwpx) 서식이 내용이 채워진 채로 만들어집니다.") + `
  <div class="steps">
    <div class="step"><div class="step-head"><span class="step-no">1</span>기관정보</div>
      <p>학교(기관) 이름 · 결재선 · 비상연락망을 처음 한 번만 채웁니다.</p></div>
    <div class="step"><div class="step-head"><span class="step-no">2</span>계약정보</div>
      <p>계약명 · 기간 · 업체를 넣으면 계약유형과 작업유형(고소 · 화기 · 밀폐 · 전기)이 자동으로 정해집니다.</p></div>
    <div class="step"><div class="step-head"><span class="step-no">3</span>서식 출력</div>
      <p>계약을 고르고 버튼을 누르면 hwpx 파일이 내려받아집니다. 한글에서 열어 고치고 인쇄합니다.</p></div>
    <div class="step"><div class="step-head"><span class="step-no">4</span>진행 현황</div>
      <p>계약마다 할 일(계획서 · 수준평가 · 점검표 · 협의체 · 순회 · 합동점검)이 날짜와 함께 나옵니다. 한 일에 체크하면 놓친 일과 이행률이 보입니다.</p></div>
  </div>
  <section class="card">
    <h2 class="card-title">자주 묻는 것</h2>
    <dl class="qa">
      <dt>입력한 정보는 어디에 저장되나요?</dt>
      <dd>이 컴퓨터의 브라우저 안에 저장됩니다. 서버가 없어 개발자에게 오지 않습니다. 다른 컴퓨터에서도 쓰려면 [저장 · 이어쓰기]에서 구글 계정으로 로그인하세요.
        자세한 내용은 <a href="./privacy.html" target="_blank" rel="noopener">개인정보처리방침</a>에 있습니다.</dd>
      <dt>다른 컴퓨터에서 이어서 쓰려면?</dt>
      <dd>[저장 · 이어쓰기] → [구글 계정으로 로그인]. 다른 컴퓨터에서도 같은 계정으로 로그인하면 같은 내용이 뜹니다.</dd>
      <dt>한글이 없어도 되나요?</dt>
      <dd>서식 파일(hwpx)을 만드는 데는 필요 없습니다. 만든 파일을 열고 고치려면 한글(2018 이후) 또는 한컴독스가 필요합니다.</dd>
      <dt>엑셀판(xlsm)과 결과가 같나요?</dt>
      <dd>같은 규칙을 옮겨 와 같은 서식이 나옵니다. 매크로 · 보안 설정이 필요 없습니다.</dd>
      <dt>내년에 같은 용역을 또 맺으면?</dt>
      <dd>[계약정보] → [다음 학년도로 이어가기]. 연간 계약을 골라 기간은 1년 뒤로, 계약명의 연도는 다음 해로 바꿔 한꺼번에 복사합니다. 한 건만이면 그 계약의 [이 계약 복사]를 누르세요.</dd>
      <dt>할 일은 무엇을 기준으로 나오나요?</dt>
      <dd>경기도교육청 「도급 · 용역 · 위탁 사업 안전 · 보건 확보 방안」(2024.4.)의 단계입니다. 모든 도급에 작업 전 점검표, 2일 이상이면 계획서와 수준평가(작업유형에 따라 A · B · C 등급),
        30일 이상이면 매월 협의체, 7일 이상이면 매주 순회점검, 90일 이상이면 분기마다 합동점검. 서식 출력의 ○ 판정과 같은 기준입니다.</dd>
      <dt>여러 서식을 한 번에 받으면?</dt>
      <dd>zip 파일 하나로 묶여 내려받아집니다. 순회점검표 · 협의체 회의록처럼 회차가 있는 서식은 한 파일 안에 회차별 쪽으로 들어갑니다.</dd>
    </dl>
  </section>`;
}

// 칸 하나. path = "org.기관명" 꼴
function field(path, label, opts = {}) {
  const [grp, key] = path.split(".");
  const val = data[grp][key] ?? "";
  const cls = String(val).trim() ? "" : "empty";
  const wide = (opts.wide ? " wide" : "") + (opts.span ? ` span-${opts.span}` : "");
  if (opts.options) {
    return `<label class="field${wide}"><span>${label}</span><select data-path="${path}" class="${cls}">
      ${opts.options.map((o) => `<option${o === val ? " selected" : ""}>${esc(o)}</option>`).join("")}
    </select></label>`;
  }
  return `<label class="field${wide}"><span>${label}</span>
    <input data-path="${path}" value="${esc(val)}" class="${cls}" placeholder="${esc(opts.ph || "")}" /></label>`;
}

// ---------------------------------------------------------------- ① 기관정보
function viewOrg() {
  const fresh = !data.org.기관명 && !data.contracts.length && !sync.s.target;
  if (fresh && googleReady()) loadGis().catch(() => {});
  const welcome = !fresh ? "" : googleReady()
    ? `<div class="alert is-info alert-row"><span><b>다른 컴퓨터에서 쓰던 내용이 있나요?</b> 그때 쓴 구글 계정으로 로그인하면 그대로 불러옵니다.</span>
        <button class="btn btn-primary btn-sm" data-sync="g-connect">구글 계정으로 불러오기</button></div>`
    : `<div class="alert is-info">받아 둔 파일이 있으면 <b>[저장 · 이어쓰기]</b>에서 불러올 수 있습니다.</div>`;
  const kind = data.org.기관구분 || "학교";
  const isSchool = kind === "학교";
  const seg = `<div class="seg" role="radiogroup" aria-label="기관구분">
    ${["학교", "교육지원청"].map((k) => `<button type="button" role="radio" aria-checked="${k === kind}"
      class="seg-btn${k === kind ? " is-on" : ""}" data-orgkind="${k}">${k}</button>`).join("")}</div>`;

  // 기관명 칸 ― 학교면 [시도 ▾][학교명][찾기]가 한 덩어리
  const nameVal = data.org.기관명 ?? "";
  const nameField = isSchool ? `
    <div class="field span-2">
      <span>학교명</span>
      <div class="input-group">
        <select id="schoolRegion" aria-label="시도교육청">
          ${NEIS_REGIONS.map(([code, nm]) => `<option value="${code}"${code === schoolRegion() ? " selected" : ""}>${nm}</option>`).join("")}
        </select>
        <input data-path="org.기관명" value="${esc(nameVal)}" class="${nameVal.trim() ? "" : "empty"}" placeholder="○○초등학교" />
        <button type="button" class="btn btn-primary" data-act="school-find">${icon("search", 18, "")}찾기</button>
      </div>
      ${schoolPanel() || `<small class="field-note">학교명을 넣고 [찾기]를 누르면 주소 · 대표전화 · 팩스번호가 채워집니다.</small>`}
    </div>`
    : field("org.기관명", "기관명", { ph: "○○교육지원청", span: 2 });

  const noVal = data.org.고유번호 ?? "";
  const noField = `
    <div class="field span-2">
      <span>고유번호</span>
      <div class="input-group">
        <input data-path="org.고유번호" value="${esc(noVal)}" class="${noVal.trim() ? "" : "empty"}" placeholder="000-00-00000" inputmode="numeric" />
        <button type="button" class="btn btn-primary" data-act="inst-find">${icon("search", 18, "")}찾기</button>
      </div>
      ${instPanel() || `<small class="field-note">고유번호(사업자등록번호)를 넣고 [찾기]를 누르면 나라장터에서 ${isSchool ? "빈 칸을" : "기관명 · 주소 · 전화 · 팩스를"} 채웁니다.</small>`}
    </div>`;

  return pageHead("기관정보", "처음 한 번만 채우면 됩니다. <b>노란 칸</b>이 아직 빈 칸입니다.") + welcome + `
  <section class="card">
    ${cardHead("building", "기관", "서식 머리 · 끝에 들어가는 우리 기관 정보입니다.")}
    ${seg}
    <div class="form-grid">
      ${nameField}
      ${noField}
      ${field("org.주소", "주소", { span: 2 })}
      ${field("org.전화번호", "대표전화")}
      ${field("org.팩스번호", "팩스번호")}
    </div>
  </section>
  <section class="card">
    ${cardHead("users", "결재선 · 기관대표자", "결재란에는 <b>직위</b>가 들어갑니다. 기관구분을 바꾸면 직위가 다시 채워집니다(고쳐 써도 됩니다).")}
    ${roster(["직위", "성명", "전화"], [
      { role: "담당자", step: 1, cells: ["sign.담당자직위", "sign.담당자성명", "sign.담당자전화"] },
      { role: "관리자", step: 2, cells: ["sign.관리자직위", "sign.관리자성명", null] },
      { role: "부서장", step: 3, cells: ["sign.부서장직위", "sign.부서장성명", null] },
      { role: "기관대표자", cells: ["sign.대표자직위", "sign.대표자성명", null], sep: true },
    ])}
  </section>
  <section class="card">
    ${cardHead("phone", "비상연락망 · 관할기관", "안전보건관리계획서의 ⑦ 비상조치 계획에 그대로 들어갑니다.")}
    ${roster(["기관명", "전화"], [
      { role: "소방서", cells: ["emerg.소방서", "emerg.소방서전화"], ph: ["○○소방서", "119"] },
      { role: "병원", cells: ["emerg.병원", "emerg.병원전화"], ph: ["인근 병원", ""] },
      { role: "고용노동지청", cells: ["emerg.노동지청", "emerg.노동지청전화"], ph: ["○○지청", ""] },
    ])}
  </section>`;
}

function cardHead(ic, title, desc) {
  return `<div class="card-head"><span class="card-icon">${icon(ic, 20, "")}</span>
    <div><h2>${title}</h2>${desc ? `<p class="hint">${desc}</p>` : ""}</div></div>`;
}

// 역할 하나 = 한 줄. 넓은 화면은 표, 좁은 화면은 역할마다 작은 묶음(칸 이름이 칸 위에 뜬다)
function roster(cols, rows) {
  const n = cols.length;
  const head = `<div class="roster-row roster-head" aria-hidden="true"><span></span>${cols.map((c) => `<span>${c}</span>`).join("")}</div>`;
  const body = rows.map((r) => {
    const cells = r.cells.map((path, i) => {
      if (!path) return `<span class="roster-na" aria-hidden="true"></span>`;
      const [g, k] = path.split(".");
      const val = data[g][k] ?? "";
      return `<label class="roster-cell"><span class="roster-label">${cols[i]}</span>
        <input data-path="${path}" value="${esc(val)}" class="${String(val).trim() ? "" : "empty"}"
          aria-label="${esc(r.role + " " + cols[i])}" placeholder="${esc(r.ph?.[i] || "")}" /></label>`;
    }).join("");
    return `<div class="roster-row${r.sep ? " roster-sep" : ""}">
      <span class="roster-role">${r.step ? `<i class="roster-step" title="결재 ${r.step}단계">${r.step}</i>` : `<i class="roster-step is-blank"></i>`}${r.role}</span>${cells}</div>`;
  }).join("");
  return `<div class="roster" style="--cols:${n}">${head}${body}</div>`;
}

// ---------------------------------------------------------------- 학교 찾기 (교육부 NEIS 학교기본정보)
// 브라우저가 open.neis.go.kr 에 바로 묻는다(서버 없음). 인증키가 없으면 한 번에 5건까지만 온다
// ― 그래서 시도교육청을 같이 고르게 한다.
const NEIS_REGIONS = [
  ["B10", "서울"], ["C10", "부산"], ["D10", "대구"], ["E10", "인천"], ["F10", "광주"], ["G10", "대전"],
  ["H10", "울산"], ["I10", "세종"], ["J10", "경기"], ["K10", "강원"], ["M10", "충북"], ["N10", "충남"],
  ["P10", "전북"], ["Q10", "전남"], ["R10", "경북"], ["S10", "경남"], ["T10", "제주"],
];
const REGION_WORDS = { B10: "서울", C10: "부산", D10: "대구", E10: "인천", F10: "광주", G10: "대전", H10: "울산",
  I10: "세종", J10: "경기", K10: "강원", M10: "충청북", N10: "충청남", P10: "전북|전라북", Q10: "전남|전라남",
  R10: "경상북|경북", S10: "경상남|경남", T10: "제주" };
let school = { region: "", busy: false, hits: null, msg: "" };

function schoolRegion() {
  if (school.region) return school.region;
  const addr = data.org.주소 || "";
  return Object.entries(REGION_WORDS).find(([, w]) => new RegExp("^(" + w + ")").test(addr))?.[0] || "J10";
}

// 찾기 결과 목록(학교 · 교육지원청 · 업체 공통). rows = [{ title, tag, sub }]
function lookupPanel(kind, st, rows, source) {
  if (st.busy) return `<div class="lookup-panel"><p class="lookup-msg">찾는 중…</p></div>`;
  if (!st.hits && !st.msg) return "";
  const items = rows.map((r, i) => `
    <li><button type="button" data-pick="${kind}:${i}">
      <b>${esc(r.title)}</b>${r.tag ? `<span class="tag">${esc(r.tag)}</span>` : ""}
      ${r.sub ? `<span class="sub">${esc(r.sub)}</span>` : ""}</button></li>`).join("");
  return `<div class="lookup-panel">
    ${st.msg ? `<p class="lookup-msg">${st.msg}</p>` : ""}
    ${items ? `<ul class="lookup-list">${items}</ul>` : ""}
    <div class="lookup-foot"><span>자료: ${source}</span>
      <button type="button" class="link-btn" data-act="lookup-close" data-kind="${kind}">닫기</button></div></div>`;
}
const hitMsg = (n, total, more) => !n ? "" : total > n ? `${total}곳 중 ${n}곳만 보입니다. ${more}` : "맞는 것을 누르세요.";

function schoolPanel() {
  return lookupPanel("school", school, (school.hits || []).map((h) => ({
    title: h.SCHUL_NM, tag: h.JU_ORG_NM || h.ATPT_OFCDC_SC_NM,
    sub: [h.ORG_RDNMA, h.ORG_TELNO].filter(Boolean).join(" · "),
  })), "교육부 나이스 교육정보 개방 포털");
}

async function schoolFind() {
  const name = (data.org.기관명 || "").trim().replace(/\s+/g, "");
  if (!name) { school = { ...school, hits: null, msg: "학교명을 먼저 넣어 주세요." }; return keepFocus(render); }
  school = { ...school, region: schoolRegion(), busy: true, hits: null, msg: "" };
  keepFocus(render);
  const q = new URLSearchParams({ Type: "json", pIndex: "1", pSize: "10", ATPT_OFCDC_SC_CODE: school.region, SCHUL_NM: name });
  if (NEIS_KEY) q.set("KEY", NEIS_KEY);
  try {
    const r = await fetch("https://open.neis.go.kr/hub/schoolInfo?" + q);
    const j = await r.json();
    const rows = j.schoolInfo?.[1]?.row || [];
    const total = j.schoolInfo?.[0]?.head?.[0]?.list_total_count || rows.length;
    school.hits = rows;
    school.msg = rows.length ? hitMsg(rows.length, total, "학교명을 더 길게 넣어 보세요.")
      : "찾는 학교가 없습니다. 시도를 확인하거나 학교명을 줄여서(예: ○○초) 다시 찾아 보세요.";
  } catch {
    school.hits = null;
    school.msg = "나이스에 연결하지 못했습니다. 인터넷 연결을 확인하거나 직접 넣어 주세요.";
  }
  school.busy = false;
  render();
}

function schoolPick(h) {
  const tel = (s) => (s || "").trim();
  data.org.기관명 = h.SCHUL_NM;
  data.org.주소 = [h.ORG_RDNMA, h.ORG_RDNDA].filter(Boolean).join(" ").trim();
  if (tel(h.ORG_TELNO)) data.org.전화번호 = tel(h.ORG_TELNO);
  if (tel(h.ORG_FAXNO)) data.org.팩스번호 = tel(h.ORG_FAXNO);
  school = { region: school.region, busy: false, hits: null, msg: "" };
  changed(); render();
  const noHint = String(data.org.고유번호 || "").trim() ? "" : "\n고유번호는 나이스에 없습니다. 고유번호 칸에 넣고 [찾기]로 확인할 수 있습니다.";
  toast(`${h.SCHUL_NM} — 주소 · 대표전화 · 팩스번호를 채웠습니다.${noHint}`);
}

// ---------------------------------------------------------------- 나라장터 찾기 (조달청 사용자정보 서비스)
// 공공데이터포털 인증키를 감추려고 이 사이트의 /api/g2b(../api/g2b.mjs)를 거친다. 묻는 말만 넘기고 저장하지 않는다.
async function g2b(kind, q) {
  let r;
  try { r = await fetch(`/api/g2b?kind=${kind}&q=${encodeURIComponent(q)}`); }
  catch { throw new Error("인터넷 연결을 확인해 주세요."); }
  let j = {};
  try { j = await r.json(); } catch {}
  if (!r.ok) throw new Error(j.error || (r.status === 404 ? "나라장터 찾기를 쓸 수 없는 주소입니다." : `오류(${r.status})`));
  return j.items || [];
}
const bizDigits = (s) => String(s || "").replace(/\D/g, "");

// 나라장터는 사업자등록번호(고유번호)로만 찾을 수 있다 ― 이름만으로는 조달청이 받지 않는다.
// 번호 하나에 기관 · 업체는 하나뿐이라 한 건이면 바로 채우고, 여러 건이면 고르게 한다.
const NEED_NO = "사업자등록번호 10자리를 넣어 주세요. 나라장터는 이름으로는 찾을 수 없습니다.";
const filled = (labels) => labels.filter(Boolean).join(" · ");

// 기관 고유번호 → 기관명 · 주소 · 전화 · 팩스 (빈 칸만)
let inst = { busy: false, hits: null, msg: "" };
const lastWord = (s) => String(s || "").trim().split(/\s+/).pop();

function instPanel() {
  return lookupPanel("inst", inst, (inst.hits || []).map((h) => ({
    title: h.short || lastWord(h.name), tag: h.bizno, sub: h.addr,
  })), "조달청 나라장터 수요기관 정보");
}

async function instFind() {
  const no = bizDigits(data.org.고유번호);
  if (no.length !== 10) { inst = { busy: false, hits: null, msg: "고유번호 10자리를 넣어 주세요. 나라장터는 이름으로는 찾을 수 없습니다." }; return keepFocus(render); }
  inst = { busy: true, hits: null, msg: "" };
  keepFocus(render);
  try {
    const items = await g2b("inst", no);
    if (items.length === 1) { inst = { busy: false, hits: null, msg: "" }; return instPick(items[0]); }
    inst = { busy: false, hits: items, msg: items.length ? "맞는 기관을 누르세요." : "나라장터에 이 번호로 등록된 기관이 없습니다. 번호를 확인해 주세요." };
  } catch (e) { inst = { busy: false, hits: null, msg: esc(e.message) + " 직접 넣어 주세요." }; }
  render();
}

function instPick(h) {
  const o = data.org, empty = (k) => !String(o[k] || "").trim();
  const nm = h.short || lastWord(h.name);
  const got = [];
  if (nm && empty("기관명")) { o.기관명 = nm; got.push("기관명"); }
  if (h.bizno) o.고유번호 = h.bizno;
  if (h.addr && empty("주소")) { o.주소 = h.addr; got.push("주소"); }
  if (h.tel && empty("전화번호")) { o.전화번호 = h.tel; got.push("대표전화"); }
  if (h.fax && empty("팩스번호")) { o.팩스번호 = h.fax; got.push("팩스번호"); }
  inst = { busy: false, hits: null, msg: "" };
  changed(); render();
  toast(got.length ? `나라장터: ${nm} — ${filled(got)}를 채웠습니다.` : `나라장터: ${nm} — 확인했습니다(이미 채워진 칸은 그대로 두었습니다).`);
}

// 업체 사업자등록번호 → 업체명 · 대표자 · 주소 · 전화
let corp = { cid: "", busy: false, hits: null, msg: "" };

function corpPanel(c) {
  if (corp.cid !== c.id) return "";
  return lookupPanel("corp", corp, (corp.hits || []).map((h) => ({
    title: h.name, tag: h.bizno, sub: [h.ceo && "대표 " + h.ceo, h.addr, h.tel].filter(Boolean).join(" · "),
  })), "조달청 나라장터 조달업체 정보");
}

async function corpFind() {
  const c = cById(selId);
  if (!c) return;
  const no = bizDigits(c.사업자번호);
  if (no.length !== 10) { corp = { cid: c.id, busy: false, hits: null, msg: NEED_NO }; return keepFocus(render); }
  corp = { cid: c.id, busy: true, hits: null, msg: "" };
  keepFocus(render);
  try {
    const items = await g2b("corp", no);
    if (items.length === 1) { corp = { cid: c.id, busy: false, hits: null, msg: "" }; return corpPick(items[0]); }
    corp = { cid: c.id, busy: false, hits: items, msg: items.length ? "맞는 업체를 누르세요."
      : "나라장터에 이 번호로 등록된 업체가 없습니다. 번호를 확인하거나 직접 넣어 주세요." };
  } catch (e) { corp = { cid: c.id, busy: false, hits: null, msg: esc(e.message) + " 직접 넣어 주세요." }; }
  render();
}

function corpPick(h) {
  const c = cById(corp.cid);
  if (!c) return;
  if (h.name) c.업체명 = h.name;
  if (h.bizno) c.사업자번호 = h.bizno;
  if (h.addr) c.업체주소 = h.addr;
  if (h.ceo) c.업체대표 = h.ceo;
  if (h.tel && !String(c.대표자전화 || "").trim()) c.대표자전화 = h.tel;
  corp = { cid: "", busy: false, hits: null, msg: "" };
  changed(); render();
  toast(`나라장터: ${h.name} — ${filled([h.name && "업체명", h.ceo && "대표자", h.addr && "주소", h.tel && "전화"])}를 채웠습니다.
업체 담당자는 직접 넣어 주세요.`);
}

// 기관구분을 바꾸면: 지금 값을 보관 → 비움 → 그 기관구분으로 적어 둔 값이 있으면 되살림
const STASH_ORG = ["기관명", "고유번호", "주소", "전화번호", "팩스번호"];
const STASH_SIGN = ["담당자성명", "담당자전화", "관리자성명", "부서장성명", "대표자성명"];
function switchOrgKind(prev, next) {
  data.stash ||= {};
  data.stash[prev] = {
    org: Object.fromEntries(STASH_ORG.map((k) => [k, data.org[k]])),
    sign: Object.fromEntries(STASH_SIGN.map((k) => [k, data.sign[k]])),
  };
  const back = data.stash[next];
  STASH_ORG.forEach((k) => { data.org[k] = back?.org?.[k] ?? ""; });
  STASH_SIGN.forEach((k) => { data.sign[k] = back?.sign?.[k] ?? ""; });
  const [a, b, c, d] = orgPositionDefaults(next);
  Object.assign(data.sign, { 담당자직위: a, 관리자직위: b, 부서장직위: c, 대표자직위: d });
}

// ---------------------------------------------------------------- ② 계약정보
function cById(id) { return data.contracts.find((c) => c.id === id); }

function viewContracts() {
  const today = todayIso();
  const prog = data.contracts.map((c) => progressOf(c, today));
  const n진행 = prog.filter((p) => p === "진행").length, n완료 = prog.filter((p) => p === "완료").length;
  const sel = cById(selId) || data.contracts[0];
  if (sel) selId = sel.id;

  const tday = todayDay();
  const list = data.contracts.map((c, i) => {
    const p = prog[i];
    const days = contractDays(c);
    const s = summaryOf(c, tday);
    return `<li><button data-sel="${c.id}" class="${p} ${c.id === selId ? "sel" : ""}">
      <span class="nm">${i + 1}. ${esc(c.계약명 || "(계약명 없음)")}</span>
      <span class="meta">${esc(c.시작일 || "")} ~ ${esc(c.종료일 || "")}${days ? ` · ${days}일` : ""}${p ? ` · ${p}` : ""}</span>
      ${s.all ? `<span class="mini-prog"><i style="width:${Math.round((s.doneAll / s.all) * 100)}%"></i></span>
        <span class="meta">할 일 ${s.doneAll}/${s.all}${s.late ? ` · <b class="late-txt">놓친 일 ${s.late}</b>` : ""}</span>` : ""}
    </button></li>`;
  }).join("");

  return pageHead("계약정보", "계약명을 넣으면 <b>계약유형 · 작업유형 · 계획서 서식</b>이 자동으로 정해집니다.") + `
  <div class="split">
    <div class="col"><section class="card">
      <h2>계약 목록</h2>
      <p class="counts">진행 ${n진행}건 · 완료 ${n완료}건 · 총 ${data.contracts.length}건</p>
      <div class="toolbar"><button class="btn btn-primary" data-act="add">+ 계약 추가</button>
        ${data.contracts.length ? `<button class="btn btn-outline" data-act="rollover" title="연간 계약을 다음 학년도 것으로 한꺼번에 복사합니다">${icon("forward", 18, "")}다음 학년도로 이어가기</button>` : ""}</div>
      <ul class="clist">${list || `<li class="empty-note">아직 계약이 없습니다.</li>`}</ul>
    </section></div>
    <div class="col">${sel ? viewContractForm(sel) : `<section class="card"><p class="empty-note">왼쪽 [+ 계약 추가]를 누르세요.</p></section>`}</div>
  </div>`;
}

function cfield(c, key, label, opts = {}) {
  const val = c[key] ?? "";
  const cls = String(val).trim() ? "" : "empty";
  const wide = opts.wide ? " wide" : "";
  if (opts.options) {
    return `<label class="field${wide}"><span>${label}</span><select data-cfield="${key}" class="${cls}">
      ${opts.options.map((o) => `<option${o === val ? " selected" : ""}>${esc(o)}</option>`).join("")}</select></label>`;
  }
  const type = opts.type || "text";
  return `<label class="field${wide}"><span>${label}</span>
    <input type="${type}" data-cfield="${key}" value="${esc(val)}" class="${cls}" placeholder="${esc(opts.ph || "")}" ${opts.num ? 'inputmode="numeric"' : ""}${opts.list ? ` list="${opts.list}" autocomplete="off"` : ""}/>
    ${opts.note ? `<small class="field-note">${opts.note}</small>` : ""}</label>`;
}

function viewContractForm(c) {
  const days = contractDays(c);
  const typeNames = ["", ...CONTRACT_TYPES.map((t) => t.유형명)];
  const chips = WORK_KEYS.map((k, i) => {
    const w = c.work?.[k] || { v: false, manual: false };
    return `<label class="chip ${w.manual ? "manual" : "auto"}" title="${WORK_NAMES[i]}">
      <input type="checkbox" data-work="${k}" ${w.v ? "checked" : ""}/> ${k}</label>`;
  }).join("");
  const ct = c.계약유형 || { v: "", manual: false };
  return `
  <section class="card">
    ${cardHead("file", esc(c.계약명 || "새 계약"), "")}
    <div class="grid">
      ${cfield(c, "계약명", "계약명", { wide: true, ph: "2026학년도 청소·미화 위탁 용역" })}
      ${cfield(c, "시작일", "계약 시작일", { type: "date" })}
      ${cfield(c, "종료일", "계약 종료일", { type: "date" })}
      ${cfield(c, "점검요일", "점검요일", { options: WEEKDAYS })}
      ${cfield(c, "계약금액", "계약금액(원)", { num: true, ph: "15600000" })}
    </div>
    <p class="hint" style="margin-top:14px">계약기간: <b>${days ? days + "일" : "—"}</b></p>
  </section>
  <section class="card">
    ${cardHead("briefcase", "계약 업체", "<b>사업자등록번호</b>를 넣고 [나라장터에서 찾기]를 누르면 업체명 · 대표자 · 주소 · 전화가 채워집니다.")}
    <div class="grid">
      <div class="field wide">
        <span>사업자등록번호</span>
        <div class="input-group">
          <input data-cfield="사업자번호" value="${esc(c.사업자번호 ?? "")}" class="${String(c.사업자번호 ?? "").trim() ? "" : "empty"}" placeholder="000-00-00000" inputmode="numeric" />
          <button type="button" class="btn btn-primary" data-act="corp-find">${icon("search", 18, "")}나라장터에서 찾기</button>
        </div>
        ${corpPanel(c)}
      </div>
      ${vendorField(c)}
      ${cfield(c, "업체대표", "업체 대표자")}
      ${cfield(c, "업체주소", "업체 주소", { wide: true })}
      ${cfield(c, "대표자전화", "대표자 전화")}
      ${cfield(c, "안전담당자", "업체 담당자(직위 성명)")}
      ${cfield(c, "담당자전화", "업체 담당자 전화")}
    </div>
  </section>
  <section class="card">
    <h2>자동 판정</h2>
    <div class="grid">
      <label class="field"><span>계약유형 ${ct.manual ? "(직접 정함)" : "(자동)"}</span>
        <select data-ctype style="background:var(${ct.manual ? "--manual" : "--auto"})">
          ${typeNames.map((t) => `<option${t === ct.v ? " selected" : ""} value="${esc(t)}">${esc(t || "(사전에 없음)")}</option>`).join("")}
        </select></label>
      ${cfield(c, "계획서유형", `안전보건관리계획서 (지금: ${planKind(c)})`, { options: ["자동", "공사", "용역", "상세"] })}
      <label class="field"><span>비고</span>
        <select data-progress style="background:var(${c.비고?.manual ? "--manual" : "--auto"})">
          ${["자동", "진행", "완료"].map((o) => `<option${(c.비고?.manual ? c.비고.v : "자동") === o ? " selected" : ""}>${o}</option>`).join("")}
        </select></label>
    </div>
    <p class="hint" style="margin:14px 0 6px">작업유형 — 해당 점검표(4~7번)가 나갑니다</p>
    <div class="auto-row">${chips}
      <button class="btn btn-outline btn-sm" data-act="reauto">자동 판정으로 되돌리기</button></div>
    <p class="legend"><i style="background:var(--auto)"></i>자동으로 채운 값 <i style="background:var(--manual)"></i>직접 정한 값(자동 판정이 건드리지 않음)</p>
  </section>
  ${checklistCard(c)}
  <section class="card">
    <div class="toolbar" style="margin:0">
      <button class="btn btn-primary" data-act="to-output">이 계약 서식 출력하러 가기 →</button>
      <button class="btn btn-outline" data-act="copy">${icon("copy", 18, "")}이 계약 복사</button>
      <button class="btn btn-danger" data-act="delete">이 계약 삭제</button>
    </div>
  </section>`;
}

// 업체명 ― 지난 계약에 적은 업체를 고르면 나머지 칸을 채운다
function vendorField(c) {
  const book = vendorBook(data.contracts.filter((x) => x.id !== c.id));
  const opts = [...book.values()].map((v) =>
    `<option value="${esc(v.업체명)}">${esc([v.사업자번호, v.n > 1 ? `계약 ${v.n}건` : ""].filter(Boolean).join(" · "))}</option>`).join("");
  return cfield(c, "업체명", "업체명", {
    ph: "(주)○○용역", list: "vendorList",
    note: book.size ? `지난 계약 업체 ${book.size}곳 — 골라 넣으면 나머지 칸이 채워집니다.` : "",
  }) + `<datalist id="vendorList">${opts}</datalist>`;
}

function vendorFill(c) {
  const v = vendorBook(data.contracts.filter((x) => x.id !== c.id)).get(String(c.업체명 || "").trim());
  if (!v) return false;
  const got = VENDOR_KEYS.filter((k) => v[k] && !String(c[k] || "").trim());
  got.forEach((k) => { c[k] = v[k]; });
  if (got.length) toast(`지난 계약의 ${v.업체명} 정보로 빈 칸 ${got.length}개를 채웠습니다.`);
  return got.length > 0;
}

// ---------------------------------------------------------------- 할 일 (계약 하나)
const WD = ["일", "월", "화", "수", "목", "금", "토"];
function fmtDue(day, today) {
  if (day == null) return "";
  const p = parts(day), t = parts(today);
  return `${p.y !== t.y ? `${String(p.y).slice(2)}.` : ""}${p.m}.${p.d}.(${WD[p.wd]})`;
}
function whenLabel(day, today, st) {
  if (day == null) return "";
  if (st === "done") return fmtDue(day, today);
  const n = day - today;
  const rel = n < 0 ? `${-n}일 지남` : n === 0 ? "오늘" : `D-${n}`;
  return `${fmtDue(day, today)} · ${rel}`;
}
const ST_LABEL = { done: "함", late: "놓침", soon: "곧", later: "예정", none: "" };

function checklistCard(c) {
  const today = todayDay();
  const s = summaryOf(c, today);
  if (!s.all) {
    return `<section class="card">${cardHead("list", "진행 사항", "계약명과 계약기간(시작일 · 종료일)을 넣으면 이 계약에 할 일이 차례로 나옵니다.")}</section>`;
  }
  const pct = Math.round((s.doneAll / s.all) * 100);
  const row = (d) => {
    const st = stateOf(c, d, today);
    return `<li class="ck ${st}">
      <label><input type="checkbox" data-done="${c.id}|${esc(d.key)}" ${st === "done" ? "checked" : ""}/>
        <span class="ck-txt"><b>${esc(d.title)}</b>${d.sub ? `<small>${esc(d.sub)}</small>` : ""}</span></label>
      <span class="when ${st}">${whenLabel(d.due, today, st)}${st === "done" ? ` <small>✓ ${esc(c.done[d.key])}</small>` : ""}</span>
      ${d.form ? `<span class="form-no" title="서식 ${d.form}번 ${esc(FORMS[d.form].title)}">서식 ${d.form}</span>` : `<span></span>`}
    </li>`;
  };
  const group = (g, items) => {
    const sts = items.map((d) => stateOf(c, d, today));
    const nDone = sts.filter((x) => x === "done").length, nLate = sts.filter((x) => x === "late").length;
    const gid = `${c.id}:${g}`;
    const chips = items.map((d, i) => `<label class="dchip ${sts[i]}" title="${esc(d.title)}">
      <input type="checkbox" data-done="${c.id}|${esc(d.key)}" ${sts[i] === "done" ? "checked" : ""}/>${fmtDue(d.due, today)}</label>`).join("");
    return `<li class="ck-group"><details data-group="${gid}"${openGroups.has(gid) ? " open" : ""}>
      <summary><span class="ck-txt"><b>${GROUP_LABEL[g]}</b>
        <small>${nDone}/${items.length}회 함${nLate ? ` · <b class="late-txt">놓친 회차 ${nLate}</b>` : ""} · 서식 ${items[0].form}</small></span>
        <span class="mini-prog"><i style="width:${Math.round((nDone / items.length) * 100)}%"></i></span></summary>
      <div class="dchips">${chips}</div>
      <div class="toolbar" style="margin:10px 0 0">
        ${nLate ? `<button class="btn btn-sm btn-outline" data-doneall="${c.id}|${g}|late">지난 회차 ${nLate}개 모두 했음</button>` : ""}
        ${nDone ? `<button class="btn btn-sm btn-outline" data-doneall="${c.id}|${g}|clear">체크 모두 지우기</button>` : ""}
      </div></details></li>`;
  };
  const phases = PHASES.map((ph) => {
    const items = s.list.filter((d) => d.phase === ph.id);
    if (!items.length) return "";
    const singles = items.filter((d) => !d.group).map(row).join("");
    const groups = [...new Set(items.filter((d) => d.group).map((d) => d.group))]
      .map((g) => group(g, items.filter((d) => d.group === g))).join("");
    return `<h3 class="ck-phase">${ph.label}</h3><ul class="cklist">${singles}${groups}</ul>`;
  }).join("");
  return `<section class="card">
    ${cardHead("list", "진행 사항", `한 일에 체크해 두세요. 체크한 날이 함께 남습니다. 기한이 지난 일은 <b class="late-txt">빨갛게</b> 보입니다.`)}
    <div class="prog-line"><span class="prog-bar"><i style="width:${pct}%"></i></span>
      <span><b>${s.doneAll}</b>/${s.all} 함 (${pct}%)${s.late ? ` · <b class="late-txt">놓친 일 ${s.late}</b>` : ""}${s.soon ? ` · 7일 안에 ${s.soon}` : ""}</span></div>
    ${phases}
    <p class="legend">기간 · 작업유형이 바뀌면 할 일도 다시 계산됩니다. 변경계약으로 기간 · 금액이 바뀌면 계획서 · 수준평가 · 점검표를 바뀐 내용으로 다시 받으세요.</p>
  </section>`;
}

function lateCount() {
  const today = todayDay();
  return data.contracts.reduce((n, c) => n + summaryOf(c, today).late, 0);
}

function setDone(cid, key, on) {
  const c = cById(cid);
  if (!c) return;
  c.done ||= {};
  if (on) c.done[key] = todayIso(); else delete c.done[key];
}

// ---------------------------------------------------------------- ⓪ 진행 현황
function viewBoard() {
  const head = pageHead("진행 현황", "계약마다 <b>지금 해야 할 일</b>과 놓친 일을 모았습니다. 체크하면 바로 저장됩니다.");
  if (!data.contracts.length) {
    return head + `<section class="card"><p class="empty-note">[계약정보]에서 계약을 넣으면 할 일이 여기에 모입니다.</p></section>`;
  }
  const today = todayDay();
  const sums = new Map(data.contracts.map((c) => [c.id, summaryOf(c, today)]));
  const items = data.contracts.flatMap((c) => sums.get(c.id).list.map((d) => ({ c, d, st: stateOf(c, d, today) })));
  const tot = [...sums.values()].reduce((a, s) => ({ due: a.due + s.due, done: a.done + s.done }), { due: 0, done: 0 });
  const rate = tot.due ? Math.round((tot.done / tot.due) * 100) : null;
  const nLate = items.filter((x) => x.st === "late").length;
  const nSoon = items.filter((x) => x.st === "soon").length;
  const nLive = data.contracts.filter((c) => progressOf(c, dayToIso(today)) === "진행").length;

  const kpi = `<div class="kpis">
    <div class="kpi"><span>진행 중인 계약</span><b>${nLive}<small>건</small></b><em>전체 ${data.contracts.length}건</em></div>
    <div class="kpi ${nLate ? "is-late" : "is-ok"}"><span>놓친 일</span><b>${nLate}<small>개</small></b><em>${nLate ? "기한이 지났는데 체크 안 됨" : "없습니다"}</em></div>
    <div class="kpi ${nSoon ? "is-soon" : ""}"><span>7일 안에 할 일</span><b>${nSoon}<small>개</small></b><em>오늘부터 일주일</em></div>
    <div class="kpi"><span>이행률</span>
      <b>${rate == null ? "—" : rate}<small>${rate == null ? "" : "%"}</small></b>
      <em>${tot.due ? `오늘까지 할 일 ${tot.due}개 중 ${tot.done}개` : "아직 기한이 온 일이 없습니다"}</em>
      ${rate == null ? "" : `<span class="prog-bar"><i style="width:${rate}%"></i></span>`}</div>
  </div>`;

  // 놓친 일 ― 계약마다 정기 점검은 묶음별로, 계약 · 착공 전 서류는 2개 넘으면 한 줄로 접는다
  const late = items.filter((x) => x.st === "late").sort((a, b) => a.d.due - b.d.due);
  const fold = new Map();
  for (const x of late) {
    const k = `${x.c.id}|${x.d.group || "_start"}`;
    if (!fold.has(k)) fold.set(k, []);
    fold.get(k).push(x);
  }
  const lateRows = [...fold].sort((a, b) => a[1][0].d.due - b[1][0].d.due).flatMap(([k, xs]) => {
    const { c, d } = xs[0];
    const g = d.group || "_start";
    if (g === "_start" && xs.length < 3) return xs.map((x) => todoRow(x, today));
    const last = xs[xs.length - 1].d.due;
    const title = g === "_start"
      ? `서류 · 확인 ${xs.length}개 체크 안 됨`
      : `${GROUP_LABEL[g]} — ${xs.length}회 체크 안 됨`;
    const what = g === "_start"
      ? xs.map((x) => x.d.title.replace(/ — .*$/, "").replace(/ 받기$/, "")).join(" · ")
      : `${fmtDue(d.due, today)}${xs.length > 1 ? ` ~ ${fmtDue(last, today)}` : ""}`;
    return [`<li class="todo late">
      <span class="todo-ic">!</span>
      <div class="todo-main"><b>${esc(title)}</b>
        <span class="todo-c">${esc(c.계약명)} · ${esc(what)}</span></div>
      <span class="when late">${today - last}일 지남</span>
      <span class="todo-btns"><button class="btn btn-sm btn-outline" data-doneall="${c.id}|${g}|late">모두 했음</button>
        <button class="btn btn-sm btn-outline" data-open="${c.id}">보기</button></span></li>`];
  });
  const lateHtml = lateRows.join("");

  const soon = items.filter((x) => x.st === "soon").sort((a, b) => a.d.due - b.d.due);
  const next = items.filter((x) => x.st === "later" && x.d.due <= today + 30).sort((a, b) => a.d.due - b.d.due);

  const section = (title, cls, body, n, empty) => `
    <div class="todo-sec ${cls}"><h3>${title} <span class="cnt">${n}</span></h3>
      ${n ? `<ul class="todos">${body}</ul>` : `<p class="todo-empty">${empty}</p>`}</div>`;

  return head + kpi + `
  <section class="card">
    ${cardHead("list", "할 일", "체크하면 그 계약의 [진행 사항]에도 똑같이 표시됩니다.")}
    ${section("놓친 일", "is-late", lateHtml, lateRows.length, "놓친 일이 없습니다.")}
    ${section("7일 안에", "is-soon", soon.map((x) => todoRow(x, today)).join(""), soon.length, "이번 주에 기한이 오는 일이 없습니다.")}
    <details class="todo-sec"${next.length && !soon.length && !lateRows.length ? " open" : ""}>
      <summary><h3>다음 30일 <span class="cnt">${next.length}</span></h3></summary>
      ${next.length ? `<ul class="todos">${next.map((x) => todoRow(x, today)).join("")}</ul>` : `<p class="todo-empty">없습니다.</p>`}
    </details>
  </section>
  ${timelineCard(today, sums)}
  <section class="card">
    ${cardHead("printer", "반기 이행 현황표", "중대재해처벌법 시행령 §4 9호의 <b>반기 1회 이상 점검</b> 기록으로 쓸 수 있게, 반기 동안 할 일과 한 일을 계약별 표로 뽑아 인쇄합니다.")}
    <button class="btn btn-primary" data-act="report">${icon("printer", 18, "")}반기 이행 현황표 만들기</button>
  </section>`;
}

function todoRow({ c, d, st }, today) {
  return `<li class="todo ${st}">
    <label class="todo-check"><input type="checkbox" data-done="${c.id}|${esc(d.key)}" ${st === "done" ? "checked" : ""} aria-label="${esc(d.title)} 했음"/></label>
    <div class="todo-main"><b>${esc(d.title)}</b><span class="todo-c">${esc(c.계약명)}</span></div>
    <span class="when ${st}">${whenLabel(d.due, today, st)}</span>
    <span class="todo-btns">${d.form ? `<button class="btn btn-sm btn-outline" data-act="to-output" data-cid="${c.id}" title="서식 ${d.form}번 ${esc(FORMS[d.form].title)}">서식 ${d.form}</button>` : ""}
      <button class="btn btn-sm btn-outline" data-open="${c.id}">보기</button></span>
  </li>`;
}

// 학년도 일정표 ― 계약 기간 막대 + 할 일 날짜 점
const ST_RANK = { late: 4, soon: 3, later: 2, done: 1, none: 0 };
function timelineCard(today, sums) {
  const y = boardYear ?? schoolYearOf(today);
  const { from, to } = schoolYearRange(y);
  const span = to - from + 1;
  const pct = (d) => (((Math.min(Math.max(d, from), to + 1) - from) / span) * 100).toFixed(3);
  const months = Array.from({ length: 12 }, (_, i) => {
    const m = ((i + 2) % 12) + 1, yy = m >= 3 ? y : y + 1;
    return { m, at: pct(toDay(`${yy}-${String(m).padStart(2, "0")}-01`)) };
  });
  const undated = data.contracts.filter((c) => !contractDays(c)).length;
  const rows = data.contracts.map((c, i) => {
    const s = toDay(c.시작일), e = toDay(c.종료일);
    if (s == null || e == null || e < s || e < from || s > to) return "";
    const sum = sums.get(c.id);
    const byDay = new Map();
    for (const d of sum.list) {
      if (d.due == null || d.due < from || d.due > to) continue;
      const st = stateOf(c, d, today);
      const prev = byDay.get(d.due);
      if (!prev || ST_RANK[st] > ST_RANK[prev.st]) byDay.set(d.due, { st, n: (prev?.n || 0) + 1, t: d.title });
      else prev.n++;
    }
    const dots = [...byDay].map(([day, v]) =>
      `<i class="dot ${v.st}" style="left:${pct(day)}%" title="${fmtDue(day, today)} ${esc(v.t)}${v.n > 1 ? ` 외 ${v.n - 1}건` : ""} — ${ST_LABEL[v.st]}"></i>`).join("");
    const p = progressOf(c, dayToIso(today));
    return `<div class="tl-row">
      <button class="tl-label" data-open="${c.id}" title="${esc(c.계약명)}">
        <span class="nm">${i + 1}. ${esc(c.계약명)}</span>
        <span class="rate">${sum.rate == null ? "기한 전" : `이행 ${sum.rate}%`}${sum.late ? ` · <b class="late-txt">놓침 ${sum.late}</b>` : ""}</span></button>
      <div class="tl-track"><span class="tl-bar ${p}" style="left:${pct(s)}%;width:${(pct(e + 1) - pct(s)).toFixed(3)}%"></span>${dots}</div>
    </div>`;
  }).filter(Boolean).join("");
  const todayLine = today >= from && today <= to ? `<span class="tl-today" style="left:${pct(today)}%"><em>오늘</em></span>` : "";
  return `<section class="card">
    <div class="card-head tl-head">
      <span class="card-icon">${icon("chart", 20, "")}</span>
      <div><h2>학년도 일정표</h2><p class="hint">막대는 계약기간, 점은 할 일입니다. 계약 이름을 누르면 그 계약으로 갑니다.</p></div>
      <div class="year-nav"><button class="btn btn-sm btn-outline" data-act="year-prev" aria-label="이전 학년도">‹</button>
        <b>${y}학년도</b><button class="btn btn-sm btn-outline" data-act="year-next" aria-label="다음 학년도">›</button></div>
    </div>
    ${rows ? `<div class="tl">
      <div class="tl-row tl-axis"><span class="tl-label"></span><div class="tl-track">${months.map((m) => `<span class="tl-m" style="left:${m.at}%">${m.m}월</span>`).join("")}</div></div>
      <div class="tl-body">${rows}<div class="tl-overlay">${todayLine}</div></div>
    </div>
    <p class="legend"><i class="lg done"></i>함 <i class="lg late"></i>놓침 <i class="lg soon"></i>7일 안 <i class="lg later"></i>예정${undated ? ` · 기간이 비어 있는 계약 ${undated}건은 빠졌습니다` : ""}</p>`
    : `<p class="empty-note">${y}학년도(${y}.3.~${y + 1}.2.)에 걸친 계약이 없습니다.</p>`}
  </section>`;
}

// ---------------------------------------------------------------- 반기 이행 현황표 (인쇄용)
function viewReport() {
  const today = todayDay();
  if (!rep) { const p = parts(today); rep = { y: p.y, h: p.m <= 6 ? 1 : 2 }; }
  const { y, h } = rep;
  const { from, to } = halfRange(y, h);
  const until = Math.min(to, today);
  const halfName = `${y}년 ${h === 1 ? "상반기(1~6월)" : "하반기(7~12월)"}`;
  const rows = data.contracts.filter((c) => {
    const s = toDay(c.시작일), e = toDay(c.종료일);
    return s != null && e != null && e >= from && s <= to && String(c.계약명 || "").trim();
  }).map((c, i) => {
    const list = dutiesOf(c).filter((d) => d.due != null && d.due >= from && d.due <= until);
    const done = list.filter((d) => c.done?.[d.key]);
    const miss = list.filter((d) => !c.done?.[d.key]);
    const missTxt = [
      ...miss.filter((d) => !d.group).map((d) => d.title.replace(/ — .*$/, "").replace(/ 받기$/, "")),
      ...Object.keys(GROUP_LABEL).map((g) => {
        const n = miss.filter((d) => d.group === g).length;
        return n ? `${GROUP_LABEL[g].replace(/ \(.*\)$/, "")} ${n}회` : "";
      }).filter(Boolean),
    ];
    const r = list.length ? Math.round((done.length / list.length) * 100) : null;
    return `<tr><td>${i + 1}</td><td class="l">${esc(c.계약명)}</td><td class="l">${esc(c.업체명 || "")}</td>
      <td>${esc(c.시작일)}<br>~ ${esc(c.종료일)}</td><td>${list.length}</td><td>${done.length}</td>
      <td>${r == null ? "—" : r + "%"}</td><td class="l">${miss.length ? esc(missTxt.join(", ")) : "없음"}</td><td></td></tr>`;
  });
  const s = data.sign;
  const signBox = [s.담당자직위, s.관리자직위, s.부서장직위].map((p) => `<td><span>${esc(p || "")}</span></td>`).join("");
  return `<div class="no-print">${pageHead("반기 이행 현황표", "도급 · 용역 · 위탁 계약마다 그 반기에 할 일과 한 일을 모았습니다. [인쇄]를 누르면 이 표만 인쇄됩니다(PDF 로 저장도 됩니다).")}
    <div class="toolbar">
      <button class="btn btn-outline" data-act="board">← 진행 현황</button>
      <button class="btn btn-outline" data-act="half-prev" aria-label="이전 반기">‹</button>
      <b class="half-now">${halfName}</b>
      <button class="btn btn-outline" data-act="half-next" aria-label="다음 반기">›</button>
      <button class="btn btn-primary" data-act="print">${icon("printer", 18, "")}인쇄</button>
    </div>
    ${until < to && today >= from ? `<div class="alert is-info">이 반기가 아직 끝나지 않아 <b>오늘(${dayToIso(today)})까지</b> 기한이 온 일만 셉니다.</div>` : ""}
    ${today < from ? `<div class="alert is-warn">아직 오지 않은 반기입니다.</div>` : ""}
  </div>
  <section class="card report">
    <div class="rep-top">
      <div><h2 class="rep-title">${y}년 ${h === 1 ? "상반기" : "하반기"} 도급 · 용역 · 위탁 안전 · 보건 확보 이행 점검표</h2>
        <p class="rep-meta">${esc(data.org.기관명 || "")} · 점검 기간 ${dayToIso(from)} ~ ${dayToIso(until < from ? to : until)} · 작성일 ${todayIso()}</p></div>
      <table class="sign"><tr><th rowspan="2">결<br>재</th>${signBox}</tr><tr><td></td><td></td><td></td></tr></table>
    </div>
    ${rows.length ? `<div class="table-wrap"><table class="out rep">
      <thead><tr><th>연번</th><th>계약명</th><th>업체</th><th>계약기간</th><th>할 일</th><th>한 일</th><th>이행률</th><th>안 한 일</th><th>조치 사항</th></tr></thead>
      <tbody>${rows.join("")}</tbody></table></div>`
    : `<p class="empty-note">이 반기에 걸친 계약이 없습니다.</p>`}
    <p class="legend">근거: 중대재해 처벌 등에 관한 법률 시행령 §4 9호(도급 · 용역 · 위탁 시 안전 · 보건 확보 기준 · 절차, 반기 1회 이상 점검),
      경기도교육청 「중대산업재해 예방을 위한 도급 · 용역 · 위탁 사업 안전 · 보건 확보 방안」(2024.4.). 할 일은 이 프로그램에 체크한 기록으로 셉니다.</p>
  </section>`;
}

// ---------------------------------------------------------------- 복사 · 다음 학년도로 이어가기
function copyContract() {
  const c = cById(selId);
  if (!c) return;
  const x = cloneContract(c, uid());
  x.계약명 = `${c.계약명 || "새 계약"} (사본)`;
  data.contracts.splice(data.contracts.indexOf(c) + 1, 0, x);
  selId = x.id;
  changed(); render();
  toast("복사했습니다. 계약명 · 기간 · 금액을 고쳐 쓰세요.\n진행 사항 체크는 비워 두었습니다.");
  view.querySelector('[data-cfield="계약명"]')?.select();
}

async function rollover() {
  const names = new Set(data.contracts.map((c) => String(c.계약명 || "").trim()));
  const cands = data.contracts.filter((c) => String(c.계약명 || "").trim());
  if (!cands.length) { toast("계약명이 있는 계약이 없습니다."); return; }
  const lis = cands.map((c) => {
    const nextName = shiftNameYears(c.계약명, 1);
    const dup = names.has(nextName.trim()) && nextName !== c.계약명;
    return `<li><label class="roll-item${dup ? " is-dup" : ""}">
      <input type="checkbox" data-roll="${c.id}" ${isAnnual(c) && !dup ? "checked" : ""} ${dup ? "disabled" : ""}/>
      <span><b>${esc(nextName)}</b><small>${c.시작일 && c.종료일
        ? `${esc(shiftYears(c.시작일, 1))} ~ ${esc(shiftYears(c.종료일, 1))}` : "기간 없음"}${isAnnual(c) ? " · 연간" : ""}${dup ? " · 이미 있습니다" : ""}</small></span></label></li>`;
  }).join("");
  const pick = await choose("다음 학년도로 이어가기",
    `<p>고른 계약을 <b>1년 뒤 기간</b>으로 복사합니다. 계약명의 연도(예: 2026학년도 → 2027학년도)도 바꿉니다.
      업체 · 금액 · 작업유형은 그대로 두고 진행 사항 체크는 비웁니다. 원래 계약은 그대로 남습니다.</p>
     <p class="hint">기간이 300일 넘는 연간 계약을 미리 골라 두었습니다.</p>
     <ul class="roll-list">${lis}</ul>`,
    [{ label: "고른 계약 이어가기", value: "go", cls: "btn-primary" }, { label: "닫기", value: null }]);
  if (pick !== "go") return;
  const ids = [...$("#dlg").querySelectorAll("input[data-roll]:checked")].map((i) => i.dataset.roll);
  if (!ids.length) { toast("고른 계약이 없습니다."); return; }
  const made = ids.map((id) => refreshAuto(rollForward(cById(id), uid(), 1)));
  data.contracts.push(...made);
  selId = made[0].id;
  changed(); render();
  toast(`${made.length}건을 다음 학년도로 이어 만들었습니다.\n금액 · 업체 담당자가 바뀌었으면 고쳐 주세요.`, 6000);
}

// ---------------------------------------------------------------- ③ 출력
// 표 머리에 쓰는 짧은 이름 · 묶음
const SHORT = { 1: "계획서", 2: "수준평가", 3: "도급점검", 4: "고소", 5: "화기", 6: "밀폐", 7: "전기", 8: "협의체", 9: "순회", 10: "합동" };
const COL_GROUPS = [
  { label: "기본 서식", nos: [1, 2, 3] },
  { label: "작업별 점검표", nos: [4, 5, 6, 7] },
  { label: "정기 점검", nos: [8, 9, 10], cls: "g8" },
];
// 하나씩 받기 ― [버튼 이름, 한 줄 설명]
const EXPORTS = {
  forms17: { label: "기본 서식 · 작업별 점검표", sub: "계획서 · 수준평가 · 도급점검 · 고소 · 화기 · 밀폐 · 전기 중 ○ 표시된 것", tag: "1~7" },
  council: { label: "안전보건협의체 회의록", sub: "매달 1회 · 회차마다 한 쪽", tag: "8" },
  patrol: { label: "작업장 순회점검표", sub: "매달 한 장 · 그 달 점검일이 들어감", tag: "9" },
  joint: { label: "작업장 합동점검표", sub: "분기(90일)마다", tag: "10" },
};
const ON_DEMAND = [11, 12, 13, 14, 15]; // 필요할 때만 ― 번호는 서식 파일 이름(11_ … 15_)과 같게 둔다

function viewOutput() {
  const head = pageHead("서식 출력", "계약을 고르고 버튼을 누르면 <b>한글(hwpx) 파일</b>이 내려받아집니다. 여러 개면 zip 하나로 묶습니다.");
  if (!data.contracts.length) return head + `<section class="card"><p class="empty-note">[계약정보]에서 계약을 먼저 넣어 주세요.</p></section>`;
  const today = todayIso();
  const nos = COL_GROUPS.flatMap((g) => g.nos);
  const groupRow = COL_GROUPS.map((g) => `<th colspan="${g.nos.length}" class="grp ${g.cls || ""}">${g.label}</th>`).join("");
  const cols = nos.map((no) =>
    `<th class="${no >= 8 ? "g8" : ""}" title="${no}. ${esc(FORMS[no].title)}">${SHORT[no]}</th>`).join("");
  const rows = data.contracts.map((c, i) => {
    const p = progressOf(c, today);
    const marks = nos.map((no) => {
      const on = formMarked(c, no);
      const ov = typeof c.formOverride?.[no] === "boolean";
      return `<td class="mark${ov ? " override" : ""}" data-mark="${c.id}:${no}" title="${esc(FORMS[no].title)} — 눌러서 ○ 바꾸기">${on ? "○" : ""}</td>`;
    }).join("");
    return `<tr class="${p}">
      <td><input type="checkbox" data-check="${c.id}" ${checked.has(c.id) ? "checked" : ""} aria-label="${esc(c.계약명 || "계약")} 고르기" style="width:20px;height:20px"/></td>
      <td>${i + 1}</td>
      <td class="name">${esc(c.계약명)}</td>
      <td>${contractDays(c) ? contractDays(c) + "일" : ""}</td>
      <td>${esc(c.점검요일 || "")}</td>
      ${marks}</tr>`;
  }).join("");

  const nPick = data.contracts.filter((c) => checked.has(c.id)).length;
  const nForms = data.contracts.filter((c) => checked.has(c.id))
    .reduce((n, c) => n + nos.filter((no) => formMarked(c, no)).length, 0);
  const tile = (act, e) => `
    <button class="out-tile" data-export="${act}" data-label="${esc(e.label)}">
      <span class="out-no">${e.tag}</span>
      <span class="out-txt"><b>${esc(e.label)}</b><small>${esc(e.sub)}</small></span>
      ${icon("download", 18, "out-ic")}</button>`;

  return head + `
  <section class="card">
    ${cardHead("file", "① 계약 고르기", "맨 왼쪽 칸으로 출력할 계약을 고르세요. <b>○</b> 는 그 계약에 갖춰야 할 서식입니다(칸을 누르면 바꿀 수 있습니다).")}
    <div class="table-wrap"><table class="out">
      <thead>
        <tr><th rowspan="2"><input type="checkbox" data-checkall aria-label="모두 고르기" style="width:20px;height:20px"/></th>
          <th rowspan="2">연번</th><th rowspan="2">계약명</th><th rowspan="2">기간</th><th rowspan="2">점검요일</th>${groupRow}</tr>
        <tr>${cols}</tr></thead>
      <tbody>${rows}</tbody>
    </table></div>
    <p class="legend">칸 이름에 마우스를 올리면 서식의 전체 이름이 보입니다. <i style="background:var(--manual)"></i>직접 바꾼 칸</p>
  </section>

  <section class="card">
    ${cardHead("download", "② 받기", nPick ? `고른 계약 <b>${nPick}건</b> · ○ 표시된 서식 <b>${nForms}개</b>` : "먼저 위 표에서 계약을 고르세요.")}
    <button class="out-all" data-export="all" data-label="일괄출력">
      ${icon("download", 24, "")}
      <span><b>일괄 출력</b><small>고른 계약의 ○ 표시된 서식을 모두 한 번에 받습니다 (zip, 계약마다 폴더)</small></span></button>

    <h3 class="out-h">하나씩 받기</h3>
    <div class="out-tiles">${Object.entries(EXPORTS).map(([act, e]) => tile(act, e)).join("")}</div>

    <h3 class="out-h">필요할 때만 <small>표의 ○ 와 상관없이, 고른 계약마다 한 부씩</small></h3>
    <div class="out-tiles is-sm">${ON_DEMAND.map((no) => tile(`single-${no}`, { label: FORMS[no].title, sub: "", tag: String(no) })).join("")}</div>
  </section>`;
}

const exporter = new Exporter(async (name) => {
  const r = await fetch(`./templates/${encodeURIComponent(name)}`);
  if (!r.ok) throw new Error(`서식 파일을 받지 못했습니다: ${name}`);
  return new Uint8Array(await r.arrayBuffer());
}, todayDay());

const safeDir = (s) => String(s).replace(/[\\/:*?"<>|·]/g, "").trim().slice(0, 60) || "계약";

async function runExport(act, btn) {
  const picked = data.contracts.filter((c) => checked.has(c.id) && String(c.계약명 || "").trim());
  if (!picked.length) { toast("출력할 계약을 왼쪽 칸에서 골라 주세요."); return; }
  btn.disabled = true;
  const old = btn.innerHTML, label = btn.dataset.label || act;
  btn.classList.add("is-busy");
  btn.querySelector("b").textContent = "만드는 중…";
  try {
    exporter.today = todayDay();
    let r;
    if (act === "all") r = await exporter.all(data, picked);
    else if (act === "forms17") r = await exporter.forms17(data, picked);
    else if (act === "council") r = await exporter.council(data, picked);
    else if (act === "patrol") r = await exporter.patrol(data, picked);
    else if (act === "joint") r = await exporter.joint(data, picked);
    else r = await exporter.single(data, picked, +act.split("-")[1]);

    if (r.files.length === 1) {
      download(r.files[0].name, r.files[0].bytes, "application/hwp+zip");
    } else if (r.files.length > 1) {
      const z = new JSZip();
      const byDir = act === "all" && picked.length > 1;
      r.files.forEach((f) => z.file((byDir && f.dir ? safeDir(f.dir) + "/" : "") + f.name, f.bytes));
      const bytes = await z.generateAsync({ type: "uint8array" });
      const first = String(picked[0].계약명).replace(/[\\/:*?"<>|·]/g, "");
      download(`${first}${picked.length > 1 ? `_외${picked.length - 1}건` : ""}_${label.replace(/[^\w가-힣~]/g, "")}.zip`, bytes, "application/zip");
    }
    const msg = [`${r.files.length}개 파일을 만들었습니다.`];
    if (r.notes.length) msg.push("", "만들지 않은 것:", ...r.notes.map((n) => " · " + n));
    toast(msg.join("\n"), r.notes.length ? 8000 : 3500);
  } catch (err) {
    console.error(err);
    toast("출력하지 못했습니다: " + err.message, 8000);
  } finally {
    btn.disabled = false;
    btn.classList.remove("is-busy");
    btn.innerHTML = old;
  }
}

// ---------------------------------------------------------------- 이벤트
const view = $("#view");

view.addEventListener("input", (e) => {
  const t = e.target;
  if (t.dataset.path && t.tagName === "INPUT") {
    const [g, k] = t.dataset.path.split(".");
    data[g][k] = t.value;
    t.classList.toggle("empty", !t.value.trim());
    changed();
  }
  if (t.dataset.cfield && t.tagName === "INPUT" && t.type !== "date") {
    const c = cById(selId);
    c[t.dataset.cfield] = t.value;
    t.classList.toggle("empty", !t.value.trim());
    // 업체명을 지난 계약의 업체로 고르면(목록에서 고르거나 똑같이 치면) 빈 칸을 채운다
    if (t.dataset.cfield === "업체명" && vendorFill(c)) keepFocus(render);
    if (t.dataset.cfield === "계약명") {
      refreshAuto(c);
      // 자동 판정 결과만 다시 그린다 ― 입력 중인 칸의 커서를 잃지 않게
      clearTimeout(view.reTimer);
      view.reTimer = setTimeout(() => { keepFocus(render); }, 600);
    }
    changed();
  }
});

// 기관명 · 업체명 · 사업자등록번호 칸에서 Enter = [찾기]
view.addEventListener("keydown", (e) => {
  if (e.key !== "Enter" || e.isComposing) return;
  const d = e.target.dataset || {};
  if (d.path === "org.기관명" && (data.org.기관구분 || "학교") === "학교") { e.preventDefault(); schoolFind(); }
  else if (d.path === "org.고유번호") { e.preventDefault(); instFind(); }
  else if (d.cfield === "사업자번호") { e.preventDefault(); corpFind(); }
});

function keepFocus(fn) {
  const a = document.activeElement;
  const key = a?.dataset?.cfield || a?.dataset?.path;
  const pos = a?.selectionStart;
  fn();
  if (key) {
    const el = view.querySelector(`[data-cfield="${key}"],[data-path="${key}"]`);
    if (el) { el.focus(); try { el.setSelectionRange(pos, pos); } catch {} }
  }
}

view.addEventListener("change", (e) => {
  const t = e.target;
  if (t.id === "fileImport") {
    const f = t.files[0];
    t.value = "";
    if (f) importFile(f);
    return;
  }
  if (t.id === "schoolRegion") { school.region = t.value; return; }
  if (t.dataset.done) {
    const i = t.dataset.done.indexOf("|");
    setDone(t.dataset.done.slice(0, i), t.dataset.done.slice(i + 1), t.checked);
    changed(); render(); return;
  }
  if (t.dataset.path && t.tagName === "SELECT") {
    const [g, k] = t.dataset.path.split(".");
    const prev = data[g][k];
    data[g][k] = t.value;
    if (t.dataset.path === "org.기관구분" && prev !== t.value) switchOrgKind(prev, t.value);
    changed(); render();
  }
  const c = cById(selId);
  if (t.dataset.cfield && (t.tagName === "SELECT" || t.type === "date")) {
    c[t.dataset.cfield] = t.value;
    changed(); render();
  }
  if (t.dataset.work) {
    c.work[t.dataset.work] = { v: t.checked, manual: true };
    changed(); render();
  }
  if (t.hasAttribute("data-ctype")) {
    c.계약유형 = { v: t.value, manual: true };
    refreshAuto(c);
    changed(); render();
  }
  if (t.hasAttribute("data-progress")) {
    c.비고 = t.value === "자동" ? { v: "", manual: false } : { v: t.value, manual: true };
    changed(); render();
  }
  if (t.dataset.check) {
    t.checked ? checked.add(t.dataset.check) : checked.delete(t.dataset.check);
    render();
  }
  if (t.hasAttribute("data-checkall")) {
    data.contracts.forEach((x) => (t.checked ? checked.add(x.id) : checked.delete(x.id)));
    render();
  }
});

view.addEventListener("click", (e) => {
  const t = e.target.closest("button, td[data-mark]");
  if (!t) return;
  const sy = t.dataset.sync;
  if (sy) {
    if (sy === "g-connect") googleConnect();
    else if (sy === "g-off") googleOff();
    else if (sy === "f-new") fileConnect("new");
    else if (sy === "f-open") fileConnect("open");
    else if (sy === "f-off") fileOff();
    else if (sy === "reconnect") reconnect();
    else if (sy === "backup") { backupDownload(); toast("내 정보를 파일로 받았습니다.\n다른 컴퓨터에서는 [파일 불러오기]로 이 파일을 여시면 됩니다."); }
    return;
  }
  if (t.dataset.orgkind) {
    const prev = data.org.기관구분, next = t.dataset.orgkind;
    if (prev === next) return;
    data.org.기관구분 = next;
    switchOrgKind(prev, next);
    school = { ...school, hits: null, msg: "" };
    inst = { busy: false, hits: null, msg: "" };
    changed(); render(); return;
  }
  if (t.dataset.pick) {
    const [k, i] = t.dataset.pick.split(":");
    if (k === "school") schoolPick(school.hits[+i]);
    else if (k === "inst") instPick(inst.hits[+i]);
    else if (k === "corp") corpPick(corp.hits[+i]);
    return;
  }
  if (t.dataset.sel) { selId = t.dataset.sel; render(); return; }
  if (t.dataset.open) { selId = t.dataset.open; tab = "contracts"; render(); window.scrollTo({ top: 0 }); return; }
  if (t.dataset.doneall) {
    const [cid, g, how] = t.dataset.doneall.split("|");
    const c = cById(cid), today = todayDay();
    dutiesOf(c).filter((d) => (g === "_start" ? !d.group : d.group === g)).forEach((d) => {
      if (how === "clear") setDone(cid, d.key, false);
      else if (stateOf(c, d, today) === "late") setDone(cid, d.key, true);
    });
    changed(); render(); return;
  }
  if (t.dataset.mark) {
    const [id, no] = t.dataset.mark.split(":");
    const c = cById(id);
    c.formOverride ||= {};
    const auto = !!formMarked({ ...c, formOverride: {} }, +no);
    const now = formMarked(c, +no);
    if (!now === auto) delete c.formOverride[no]; else c.formOverride[no] = !now;
    changed(); render(); return;
  }
  if (t.dataset.export) { runExport(t.dataset.export, t); return; }
  const act = t.dataset.act;
  if (act === "school-find") { schoolFind(); return; }
  if (act === "inst-find") { instFind(); return; }
  if (act === "corp-find") { corpFind(); return; }
  if (act === "lookup-close") {
    const k = t.dataset.kind, off = { busy: false, hits: null, msg: "" };
    if (k === "school") school = { ...school, ...off }; else if (k === "inst") inst = off; else corp = { cid: "", ...off };
    render(); return;
  }
  if (act === "add") {
    const c = refreshAuto({ id: uid(), 계약명: "", 시작일: "", 종료일: "", 점검요일: "", 계약금액: "",
      업체명: "", 사업자번호: "", 업체주소: "", 업체대표: "", 대표자전화: "", 안전담당자: "", 담당자전화: "",
      계획서유형: "", work: {}, 계약유형: { v: "", manual: false }, 비고: { v: "", manual: false }, formOverride: {} });
    data.contracts.push(c);
    selId = c.id;
    changed(); render();
    view.querySelector('[data-cfield="계약명"]')?.focus();
  } else if (act === "delete") {
    const c = cById(selId);
    if (c && confirm(`'${c.계약명 || "새 계약"}' 을 지웁니다. 되돌릴 수 없습니다. 계속할까요?`)) {
      data.contracts = data.contracts.filter((x) => x.id !== selId);
      checked.delete(selId); selId = null;
      changed(); render();
    }
  } else if (act === "reauto") {
    const c = cById(selId);
    WORK_KEYS.forEach((k) => { c.work[k] = { v: false, manual: false }; });
    c.계약유형 = { v: "", manual: false };
    refreshAuto(c);
    changed(); render();
  } else if (act === "to-output") {
    checked.clear(); checked.add(t.dataset.cid || selId);
    tab = "output"; render(); window.scrollTo({ top: 0 });
  } else if (act === "copy") copyContract();
  else if (act === "rollover") rollover();
  else if (act === "year-prev" || act === "year-next") {
    boardYear = (boardYear ?? schoolYearOf(todayDay())) + (act === "year-next" ? 1 : -1);
    render();
  } else if (act === "report") { tab = "report"; render(); window.scrollTo({ top: 0 }); }
  else if (act === "board") { tab = "board"; render(); }
  else if (act === "half-prev" || act === "half-next") {
    const n = rep.y * 2 + (rep.h - 1) + (act === "half-next" ? 1 : -1);
    rep = { y: Math.floor(n / 2), h: (n % 2) + 1 };
    render();
  } else if (act === "print") window.print();
});

// 정기 점검 묶음을 펼쳐 둔 채로 다시 그리게 기억한다(toggle 은 거품이 일지 않아 capture 로 받는다)
view.addEventListener("toggle", (e) => {
  const g = e.target.dataset?.group;
  if (!g) return;
  e.target.open ? openGroups.add(g) : openGroups.delete(g);
}, true);

// ---------------------------------------------------------------- 시작
(async () => {
  const saved = await loadLocal();
  if (saved) {
    try { data = normalize(saved); } catch { data = emptyData(); }
    data.contracts.forEach((c) => { c.id ??= uid(); refreshAuto(c); });
  }
  if (!data.org.기관명) tab = "org";
  else tab = data.contracts.length ? "board" : "contracts";
  render();

  // 켜 둔 자동 저장 대상을 다시 잇는다. 로그인 · 허락이 살아 있으면 바로, 아니면 [이어서 하기] 한 번
  // 이 창에서 구글 로그인하고 막 돌아온 경우
  const back = googleReady() ? takeRedirectResult() : null;
  if (back?.ok) { await finishGoogleConnect(); return; }
  if (back?.error) toast(back.error, 6000);

  const gmeta = googleReady() ? await loadGoogleMeta() : null;
  if (gmeta) {
    const t = googleTarget(gmeta);
    if (tokenValid()) { authedThisSession = true; await sync.attach(t); }
    else { loadGis().catch(() => {}); sync.hold(t); }
    return;
  }
  if (linkSupported()) {
    const link = await loadLink();
    if (link?.handle) {
      const t = fileTarget(link);
      if ((await permission(link.handle)) === "granted") await sync.attach(t);
      else sync.hold(t);
    }
  }
})();
