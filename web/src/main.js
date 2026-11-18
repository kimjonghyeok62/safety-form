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
import {
  refreshAuto, WORK_KEYS, WORK_NAMES, FORMS, planKind, progressOf,
  orgPositionDefaults, CONTRACT_TYPES,
} from "./domain/dictionary.js";
import { contractDays, WEEKDAYS, todayDay, dayToIso } from "./domain/schedule.js";
import { Exporter, formMarked } from "./export.js";

// ---------------------------------------------------------------- 상태
let data = emptyData();
let tab = "org";
let selId = null;
const checked = new Set(); // 출력 화면에서 고른 계약

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
  save: `<path d="M19 21H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h11l5 5v11a2 2 0 0 1-2 2z"/><polyline points="17 21 17 13 7 13 7 21"/><polyline points="7 3 7 8 15 8"/>`,
};
const icon = (name, size = 20, cls = "sidenav-icon") =>
  `<svg class="${cls}" width="${size}" height="${size}" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">${ICONS[name]}</svg>`;

const NAV = [
  { group: "작성", items: [
    { id: "org", label: "기관정보", icon: "building" },
    { id: "contracts", label: "계약정보", icon: "file" },
    { id: "output", label: "서식 출력", icon: "download" },
  ] },
  { group: "저장", items: [{ id: "save", label: "저장 · 이어쓰기", icon: "cloud" }] },
  { group: "참고", items: [{ id: "help", label: "처음이라면 — 사용법", icon: "book" }] },
];

function renderNav() {
  const item = (it) => {
    const on = tab === it.id;
    const badge = it.id === "contracts" && data.contracts.length ? `<span class="sidenav-badge">${data.contracts.length}</span>` : "";
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
  v.className = "page" + ({ org: " is-mid", contracts: " is-wide", output: " is-wide" }[tab] || "");
  const body = tab === "org" ? viewOrg()
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
  const wide = opts.wide ? " wide" : "";
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
  return pageHead("기관정보", "처음 한 번만 채우면 됩니다. <b>노란 칸</b>이 아직 빈 칸입니다.") + welcome + `
  <section class="card">
    <h2>기관</h2>
    <div class="grid">
      ${field("org.기관구분", "기관구분", { options: ["학교", "교육지원청"] })}
      ${field("org.기관명", "기관명", { ph: "○○초등학교" })}
      ${field("org.고유번호", "고유번호")}
      ${field("org.주소", "주소", { wide: true })}
      ${field("org.전화번호", "대표전화")}
      ${field("org.팩스번호", "팩스번호")}
    </div>
  </section>
  <section class="card">
    <h2>결재선 · 기관대표자</h2>
    <p class="hint">결재란에는 <b>직위</b>가 들어갑니다. 기관구분을 고르면 직위가 다시 채워집니다(고쳐 써도 됩니다).</p>
    <div class="grid">
      ${field("sign.담당자직위", "담당자 직위")}
      ${field("sign.담당자성명", "담당자 성명")}
      ${field("sign.담당자전화", "담당자 전화")}
      ${field("sign.관리자직위", "관리자 직위")}
      ${field("sign.관리자성명", "관리자 성명")}
      ${field("sign.부서장직위", "부서장 직위")}
      ${field("sign.부서장성명", "부서장 성명")}
      ${field("sign.대표자직위", "기관대표자 직위")}
      ${field("sign.대표자성명", "기관대표자 성명")}
    </div>
  </section>
  <section class="card">
    <h2>비상연락망 · 관할기관</h2>
    <p class="hint">안전보건관리계획서의 ⑦ 비상조치 계획에 그대로 들어갑니다.</p>
    <div class="grid">
      ${field("emerg.소방서", "관할 소방서")}
      ${field("emerg.소방서전화", "소방서 전화")}
      ${field("emerg.병원", "인근 병원")}
      ${field("emerg.병원전화", "병원 전화")}
      ${field("emerg.노동지청", "관할 고용노동지청")}
      ${field("emerg.노동지청전화", "노동지청 전화")}
    </div>
  </section>`;
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

  const list = data.contracts.map((c, i) => {
    const p = prog[i];
    const days = contractDays(c);
    return `<li><button data-sel="${c.id}" class="${p} ${c.id === selId ? "sel" : ""}">
      <span class="nm">${i + 1}. ${esc(c.계약명 || "(계약명 없음)")}</span>
      <span class="meta">${esc(c.시작일 || "")} ~ ${esc(c.종료일 || "")}${days ? ` · ${days}일` : ""}${p ? ` · ${p}` : ""}</span>
    </button></li>`;
  }).join("");

  return pageHead("계약정보", "계약명을 넣으면 <b>계약유형 · 작업유형 · 계획서 서식</b>이 자동으로 정해집니다.") + `
  <div class="split">
    <div class="col"><section class="card">
      <h2>계약 목록</h2>
      <p class="counts">진행 ${n진행}건 · 완료 ${n완료}건 · 총 ${data.contracts.length}건</p>
      <div class="toolbar"><button class="btn btn-primary" data-act="add">+ 계약 추가</button></div>
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
    <input type="${type}" data-cfield="${key}" value="${esc(val)}" class="${cls}" placeholder="${esc(opts.ph || "")}" ${opts.num ? 'inputmode="numeric"' : ""}/></label>`;
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
    <h2 style="margin-bottom:14px">${esc(c.계약명 || "새 계약")}</h2>
    <div class="grid">
      ${cfield(c, "계약명", "계약명", { wide: true, ph: "2026학년도 청소·미화 위탁 용역" })}
      ${cfield(c, "시작일", "계약 시작일", { type: "date" })}
      ${cfield(c, "종료일", "계약 종료일", { type: "date" })}
      ${cfield(c, "점검요일", "점검요일", { options: WEEKDAYS })}
      ${cfield(c, "계약금액", "계약금액(원)", { num: true, ph: "15600000" })}
      ${cfield(c, "업체명", "업체명")}
      ${cfield(c, "사업자번호", "사업자등록번호")}
      ${cfield(c, "업체주소", "업체 주소", { wide: true })}
      ${cfield(c, "업체대표", "업체 대표자")}
      ${cfield(c, "대표자전화", "대표자 전화")}
      ${cfield(c, "안전담당자", "업체 담당자(직위 성명)")}
      ${cfield(c, "담당자전화", "업체 담당자 전화")}
    </div>
    <p class="hint" style="margin-top:14px">계약기간: <b>${days ? days + "일" : "—"}</b></p>
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
  <section class="card">
    <div class="toolbar" style="margin:0">
      <button class="btn btn-primary" data-act="to-output">이 계약 서식 출력하러 가기 →</button>
      <button class="btn btn-danger" data-act="delete">이 계약 삭제</button>
    </div>
  </section>`;
}

// ---------------------------------------------------------------- ③ 출력
const BTNS = [
  { act: "forms17", label: "1~7번 출력", cls: "navy" },
  { act: "council", label: "8.협의체 회의록", cls: "green" },
  { act: "patrol", label: "9.순회점검표", cls: "green" },
  { act: "joint", label: "10.합동점검표", cls: "green" },
  { act: "single-11", label: "11.아차사고", cls: "grey" },
  { act: "single-12", label: "12.관리감독자", cls: "grey" },
  { act: "single-13", label: "13.발생보고", cls: "grey" },
  { act: "single-14", label: "14.사고보고", cls: "grey" },
  { act: "single-15", label: "15.이행서약서", cls: "navy" },
];

function viewOutput() {
  const head = pageHead("서식 출력", "계약을 고르고 버튼을 누르면 <b>한글(hwpx) 파일</b>이 내려받아집니다. 여러 개면 zip 하나로 묶습니다.");
  if (!data.contracts.length) return head + `<section class="card"><p class="empty-note">[계약정보]에서 계약을 먼저 넣어 주세요.</p></section>`;
  const today = todayIso();
  const cols = FORMS.slice(1).map((f) =>
    `<th class="${f.no >= 11 ? "g11" : f.no >= 8 && f.no <= 10 ? "g8" : ""}" title="${esc(f.title)}">${f.no}</th>`).join("");
  const rows = data.contracts.map((c, i) => {
    const p = progressOf(c, today);
    const marks = FORMS.slice(1).map((f) => {
      if (f.no >= 11) return `<td class="mark" style="color:var(--muted)">·</td>`;
      const on = formMarked(c, f.no);
      const ov = typeof c.formOverride?.[f.no] === "boolean";
      return `<td class="mark${ov ? " override" : ""}" data-mark="${c.id}:${f.no}" title="눌러서 ○ 바꾸기">${on ? "○" : ""}</td>`;
    }).join("");
    return `<tr class="${p}">
      <td><input type="checkbox" data-check="${c.id}" ${checked.has(c.id) ? "checked" : ""} style="width:20px;height:20px"/></td>
      <td>${i + 1}</td>
      <td class="name">${esc(c.계약명)}</td>
      <td>${contractDays(c) ? contractDays(c) + "일" : ""}</td>
      <td>${esc(c.점검요일 || "")}</td>
      ${marks}</tr>`;
  }).join("");

  return head + `
  <section class="card">
    <p class="hint">맨 왼쪽 칸으로 계약을 고르세요. ○ 는 그 계약에 갖춰야 할 서식입니다(칸을 눌러 바꿀 수 있습니다).</p>
    <div class="toolbar">${BTNS.map((b) => `<button class="btn ${b.cls}" data-export="${b.act}">${b.label}</button>`).join("")}</div>
    <div class="table-wrap"><table class="out">
      <thead><tr><th><input type="checkbox" data-checkall style="width:20px;height:20px"/></th><th>연번</th><th>계약명</th><th>기간</th><th>점검요일</th>${cols}</tr></thead>
      <tbody>${rows}</tbody>
    </table></div>
    <p class="legend">1 계획서 · 2 수준평가 · 3 도급점검 · 4 고소 · 5 화기 · 6 밀폐 · 7 전기 · 8 협의체(월 1회) · 9 순회(달마다) · 10 합동(분기) · 11~15 필요할 때</p>
  </section>`;
}

const exporter = new Exporter(async (name) => {
  const r = await fetch(`./templates/${encodeURIComponent(name)}`);
  if (!r.ok) throw new Error(`서식 파일을 받지 못했습니다: ${name}`);
  return new Uint8Array(await r.arrayBuffer());
}, todayDay());

async function runExport(act, btn) {
  const picked = data.contracts.filter((c) => checked.has(c.id) && String(c.계약명 || "").trim());
  if (!picked.length) { toast("출력할 계약을 왼쪽 칸에서 골라 주세요."); return; }
  btn.disabled = true;
  const old = btn.textContent;
  btn.textContent = "만드는 중…";
  try {
    exporter.today = todayDay();
    let r;
    if (act === "forms17") r = await exporter.forms17(data, picked);
    else if (act === "council") r = await exporter.council(data, picked);
    else if (act === "patrol") r = await exporter.patrol(data, picked);
    else if (act === "joint") r = await exporter.joint(data, picked);
    else r = await exporter.single(data, picked, +act.split("-")[1]);

    if (r.files.length === 1) {
      download(r.files[0].name, r.files[0].bytes, "application/hwp+zip");
    } else if (r.files.length > 1) {
      const z = new JSZip();
      r.files.forEach((f) => z.file(f.name, f.bytes));
      const bytes = await z.generateAsync({ type: "uint8array" });
      const first = String(picked[0].계약명).replace(/[\\/:*?"<>|·]/g, "");
      download(`${first}${picked.length > 1 ? `_외${picked.length - 1}건` : ""}_${old.replace(/[^\w가-힣~]/g, "")}.zip`, bytes, "application/zip");
    }
    const msg = [`${r.files.length}개 파일을 만들었습니다.`];
    if (r.notes.length) msg.push("", "만들지 않은 것:", ...r.notes.map((n) => " · " + n));
    toast(msg.join("\n"), r.notes.length ? 8000 : 3500);
  } catch (err) {
    console.error(err);
    toast("출력하지 못했습니다: " + err.message, 8000);
  } finally {
    btn.disabled = false;
    btn.textContent = old;
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
    if (t.dataset.cfield === "계약명") {
      refreshAuto(c);
      // 자동 판정 결과만 다시 그린다 ― 입력 중인 칸의 커서를 잃지 않게
      clearTimeout(view.reTimer);
      view.reTimer = setTimeout(() => { keepFocus(render); }, 600);
    }
    changed();
  }
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
  if (t.dataset.sel) { selId = t.dataset.sel; render(); return; }
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
    checked.clear(); checked.add(selId);
    tab = "output"; render();
  }
});

// ---------------------------------------------------------------- 시작
(async () => {
  const saved = await loadLocal();
  if (saved) {
    try { data = normalize(saved); } catch { data = emptyData(); }
    data.contracts.forEach((c) => { c.id ??= uid(); refreshAuto(c); });
  }
  if (!data.org.기관명) tab = "org";
  else tab = data.contracts.length ? "output" : "contracts";
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
