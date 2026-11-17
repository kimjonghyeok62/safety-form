// 화면 ― 엑셀판의 입력정보 / 출력관리 시트를 세 개의 탭으로 옮겼다.
import JSZip from "jszip";
import { loadLocal, saveLocal, emptyData, normalize } from "./store/localStore.js";
import {
  FILE_NAME, linkSupported, loadLink, saveLink, clearLink, permission,
  pickNew, pickExisting, readFile, writeFile,
} from "./store/fileLink.js";
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
const checked = new Set(); // 출력 탭에서 고른 계약

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

// ---------------------------------------------------------------- 저장
let saveTimer = null;
let localOk = true;
function changed() {
  data.updatedAt = new Date().toISOString();
  clearTimeout(saveTimer);
  saveTimer = setTimeout(async () => {
    localOk = await saveLocal(data);
    queueSync("write");
    showSaveState();
  }, 300);
}

function showSaveState() {
  const el = $("#saveState");
  const btn = $("#btnLink");
  btn.hidden = !linkSupported();
  if (linkSupported()) {
    const L = {
      none: ["드라이브 폴더에 자동 저장", "btn navy"],
      prompt: ["자동 저장 다시 켜기", "btn green"],
      error: ["자동 저장 다시 연결", "btn green"],
      ok: ["자동 저장 설정", "btn grey"],
      busy: ["자동 저장 설정", "btn grey"],
    }[linkState];
    btn.textContent = L[0]; btn.className = L[1];
  }
  if (!localOk) { el.textContent = "브라우저에 저장하지 못했습니다 — 파일로 저장해 두세요"; el.className = "save-state warn"; return; }
  if (linkState === "ok" || linkState === "busy") {
    const t = link.savedAt ? new Date(link.savedAt).toLocaleTimeString("ko-KR", { hour: "2-digit", minute: "2-digit" }) : "";
    el.textContent = `✓ 자동 저장 중 · ${link.name}${t ? ` (${t})` : ""}`;
    el.className = "save-state ok";
    return;
  }
  if (linkState === "prompt") {
    el.textContent = "자동 저장이 멈춰 있습니다 — 오른쪽 버튼을 한 번 눌러 주세요";
    el.className = "save-state warn"; return;
  }
  if (linkState === "error") {
    el.textContent = `자동 저장하지 못했습니다 — ${linkError}`;
    el.className = "save-state warn"; return;
  }
  const b = data.backedUpAt ? data.backedUpAt.slice(0, 10) : null;
  const stale = !b || (data.updatedAt && data.updatedAt > data.backedUpAt);
  el.textContent = b ? `마지막 파일 저장: ${b}${stale ? " (그 뒤 고친 내용 있음)" : ""}` : "아직 파일로 저장한 적 없음";
  el.className = "save-state" + (stale && data.contracts.length ? " warn" : "");
}

// ---------------------------------------------------------------- 자동 저장 파일(드라이브 폴더)
// link = { handle, name, syncedAt(파일과 이 PC 가 마지막으로 같았던 updatedAt), savedAt }
let link = null;
let linkState = "none"; // none | prompt | ok | busy | error
let linkError = "";

// 버튼 고르는 창. buttons = [{ label, value, cls }] → 고른 value (Esc 는 null)
function choose(title, bodyHtml, buttons) {
  const dlg = $("#dlg");
  dlg.innerHTML = `<h2>${esc(title)}</h2><div class="dlg-body">${bodyHtml}</div>
    <div class="dlg-btns">${buttons.map((b, i) =>
      `<button class="btn ${b.cls || ""}" data-i="${i}">${esc(b.label)}</button>`).join("")}</div>`;
  return new Promise((resolve) => {
    const done = (v) => { dlg.close(); resolve(v); };
    dlg.querySelectorAll("button[data-i]").forEach((b) =>
      b.addEventListener("click", () => done(buttons[+b.dataset.i].value)));
    dlg.oncancel = (e) => { e.preventDefault(); done(null); };
    dlg.showModal();
  });
}

const fmtWhen = (iso) => iso ? new Date(iso).toLocaleString("ko-KR", {
  year: "numeric", month: "numeric", day: "numeric", hour: "2-digit", minute: "2-digit" }) : "모름";
const isEmptyData = (d) => !d.org.기관명 && !d.contracts.length;

function adopt(next) {
  data = next;
  data.contracts.forEach((c) => { c.id ??= uid(); refreshAuto(c); });
  selId = null; checked.clear();
}

async function writeLinked() {
  await writeFile(link.handle, data);
  link.syncedAt = data.updatedAt;
  link.savedAt = new Date().toISOString();
  await saveLink(link);
}

// 파일과 이 PC 를 맞춘다. mode: "connect"(연결할 때) | "write"(고친 뒤) | "check"(창으로 돌아왔을 때)
async function reconcile(mode) {
  let file;
  try {
    const raw = await readFile(link.handle);
    file = raw ? normalize(raw) : null;
  } catch (err) {
    if (err?.name === "NotFoundError") throw new Error("파일을 찾지 못했습니다(옮겨졌거나 지워짐)");
    if (err instanceof SyntaxError || /프로그램에서 저장한/.test(err?.message)) {
      throw new Error("이 프로그램의 저장 파일이 아닙니다");
    }
    throw err;
  }
  const fileU = file?.updatedAt ?? null;
  const localU = data.updatedAt ?? null;
  if (!file) { await writeLinked(); return; }                       // 빈 새 파일
  if (fileU === localU) {                                            // 이미 같음
    link.syncedAt = localU; link.savedAt ??= new Date().toISOString(); await saveLink(link); return;
  }
  const base = link.syncedAt;
  const fileChanged = fileU !== base;
  const localChanged = localU !== base;

  if (isEmptyData(data) || (fileChanged && !localChanged)) {        // 다른 컴퓨터에서 고친 내용 가져오기
    adopt(file);
    await saveLocal(data);
    link.syncedAt = fileU; await saveLink(link);
    render();
    if (mode !== "connect" || !isEmptyData(file)) {
      toast(`드라이브 파일에서 최신 내용을 가져왔습니다 — 계약 ${data.contracts.length}건`);
    }
    return;
  }
  if (!fileChanged || isEmptyData(file)) { await writeLinked(); return; } // 이 PC 에서만 고침

  // 양쪽 다 고쳤다 ― 사람이 고른다
  const pick = await choose("어느 내용을 쓸까요?",
    `<p>이 컴퓨터와 드라이브 파일의 내용이 서로 다릅니다. 하나를 고르면 다른 쪽은 그 내용으로 바뀝니다.</p>
     <table class="cmp"><tr><th></th><th>마지막 수정</th><th>계약</th></tr>
       <tr><td>드라이브 파일</td><td>${esc(fmtWhen(fileU))}</td><td>${file.contracts.length}건</td></tr>
       <tr><td>이 컴퓨터</td><td>${esc(fmtWhen(localU))}</td><td>${data.contracts.length}건</td></tr></table>
     <p class="hint">잘 모르겠으면 [이 컴퓨터 내용을 파일로 저장]을 먼저 눌러 두면 이 컴퓨터 내용이 따로 남습니다.</p>`,
    [
      { label: "드라이브 파일 내용 쓰기", value: "file", cls: "navy" },
      { label: "이 컴퓨터 내용 쓰기", value: "local", cls: "navy" },
      { label: "이 컴퓨터 내용을 파일로 저장", value: "backup" },
    ]);
  if (pick === "backup") { backupDownload(); return reconcile(mode); }
  if (pick === "file") {
    adopt(file);
    await saveLocal(data);
    link.syncedAt = fileU; await saveLink(link);
    render();
    toast("드라이브 파일 내용으로 바꿨습니다.");
  } else if (pick === "local") {
    await writeLinked();
    toast("이 컴퓨터 내용을 드라이브 파일에 저장했습니다.");
  } else {
    throw new Error("어느 내용을 쓸지 고르지 않아 멈췄습니다");
  }
}

let syncChain = Promise.resolve();
function queueSync(mode) {
  if (!link || (linkState !== "ok" && linkState !== "busy")) return syncChain;
  syncChain = syncChain.then(async () => {
    if (mode !== "connect") linkState = "busy";
    try {
      await reconcile(mode);
      linkState = "ok"; linkError = "";
    } catch (err) {
      linkState = "error"; linkError = err?.message || String(err);
    }
    showSaveState();
  });
  return syncChain;
}

async function connect(handle, { ask = false } = {}) {
  const p = await permission(handle, ask);
  if (p !== "granted") {
    linkState = "prompt";
    showSaveState();
    return false;
  }
  link = link?.handle === handle ? link : { handle, name: handle.name, syncedAt: null, savedAt: null };
  await saveLink(link);
  linkState = "busy";
  await queueSync("connect");
  return linkState === "ok";
}

async function setupLink() {
  const pick = await choose("드라이브 폴더에 자동 저장",
    `<p>구글 드라이브 데스크톱 프로그램을 깔면 생기는 <b>구글 드라이브 폴더</b>(보통 <b>G:\\내 드라이브</b>)에
       저장 파일을 하나 두면, 고칠 때마다 그 파일에 저장되고 드라이브가 알아서 올려 줍니다.
       다른 컴퓨터에서도 같은 파일을 열면 이어서 쓸 수 있습니다.</p>
     <p>구글 로그인은 필요 없습니다. 이 웹페이지는 고른 파일 하나에만 손을 댑니다.</p>
     <ul>
       <li><b>처음 쓰는 경우</b>: [새 저장 파일 만들기] → 왼쪽에서 <b>Google Drive</b>(G:) → <b>내 드라이브</b>를 고르고 저장</li>
       <li><b>다른 컴퓨터에서 이미 만든 경우</b>: [있는 저장 파일 열기] → 드라이브 폴더의 <b>${esc(FILE_NAME)}</b> 열기</li>
     </ul>`,
    [
      { label: "새 저장 파일 만들기", value: "new", cls: "navy" },
      { label: "있는 저장 파일 열기", value: "open", cls: "navy" },
      { label: "닫기", value: null },
    ]);
  if (!pick) return;
  let handle;
  try {
    handle = pick === "new" ? await pickNew() : await pickExisting();
  } catch (err) {
    if (err?.name !== "AbortError") toast("파일을 고르지 못했습니다: " + err.message);
    return;
  }
  link = { handle, name: handle.name, syncedAt: null, savedAt: null };
  if (await connect(handle, { ask: true })) {
    toast(`이제부터 고칠 때마다 '${handle.name}'에 자동 저장됩니다.`);
  }
}

async function manageLink() {
  const pick = await choose("자동 저장 설정",
    `<p>지금 저장 파일: <b>${esc(link.name)}</b></p>
     <p class="hint">마지막 저장: ${esc(fmtWhen(link.savedAt))}</p>`,
    [
      { label: "다른 파일로 바꾸기", value: "change", cls: "navy" },
      { label: "자동 저장 끄기", value: "off", cls: "danger" },
      { label: "닫기", value: null },
    ]);
  if (pick === "change") return setupLink();
  if (pick === "off") {
    await clearLink();
    link = null; linkState = "none";
    showSaveState();
    toast("자동 저장을 껐습니다. 드라이브의 파일은 그대로 남아 있습니다.");
  }
}

$("#btnLink").addEventListener("click", async () => {
  if (!link || linkState === "none") return setupLink();
  if (linkState === "prompt" || linkState === "error") {
    if (await connect(link.handle, { ask: true })) toast("자동 저장을 다시 켰습니다.");
    else if (linkState === "error" && /찾지 못했|저장 파일이 아닙/.test(linkError)) {
      toast(`${linkError}\n[드라이브 폴더에 자동 저장]을 눌러 파일을 다시 골라 주세요.`, 8000);
      await clearLink(); link = null; linkState = "none"; showSaveState();
    } else if (linkState === "error") toast(linkError, 8000);
    return;
  }
  return manageLink();
});

// 다른 컴퓨터에서 고쳤을 수 있으니 창으로 돌아오면 파일을 다시 본다
document.addEventListener("visibilitychange", () => {
  if (document.visibilityState === "visible" && linkState === "ok") queueSync("check");
});

function download(name, bytes, type) {
  const url = URL.createObjectURL(new Blob([bytes], { type }));
  const a = document.createElement("a");
  a.href = url; a.download = name;
  document.body.appendChild(a); a.click(); a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 10000);
}

function backupDownload() {
  data.backedUpAt = new Date().toISOString();
  const name = `중대재해예방서식_${(data.org.기관명 || "내정보").replace(/[\\/:*?"<>|·]/g, "")}_${todayIso()}.json`;
  download(name, JSON.stringify(data, null, 1), "application/json");
  saveLocal(data).then((ok) => { localOk = ok; showSaveState(); });
}

$("#btnBackup").addEventListener("click", () => {
  backupDownload();
  toast("내 정보를 파일로 저장했습니다.\n다른 컴퓨터에서는 [파일 불러오기]로 이 파일을 여시면 됩니다.");
});

$("#fileImport").addEventListener("change", async (e) => {
  const f = e.target.files[0];
  e.target.value = "";
  if (!f) return;
  try {
    const next = normalize(JSON.parse(await f.text()));
    if (data.contracts.length && !confirm(`지금 입력된 내용(계약 ${data.contracts.length}건)을 이 파일 내용으로 바꿉니다. 계속할까요?`)) return;
    data = next;
    data.contracts.forEach((c) => { c.id ??= uid(); refreshAuto(c); });
    selId = null; checked.clear();
    changed(); render();
    toast(`불러왔습니다 — 계약 ${data.contracts.length}건`);
  } catch (err) {
    toast("불러오지 못했습니다: " + err.message);
  }
});

// ---------------------------------------------------------------- 탭
document.querySelectorAll(".tab").forEach((b) =>
  b.addEventListener("click", () => { tab = b.dataset.tab; render(); }));

function render() {
  document.querySelectorAll(".tab").forEach((b) => b.classList.toggle("active", b.dataset.tab === tab));
  const v = $("#view");
  if (tab === "org") v.innerHTML = viewOrg();
  else if (tab === "contracts") v.innerHTML = viewContracts();
  else v.innerHTML = viewOutput();
  showSaveState();
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
  return `
  <section class="card">
    <h2>1. 기관정보</h2>
    <p class="hint">처음 한 번만 채우면 됩니다. <b>노란 칸</b>이 아직 빈 칸입니다.</p>
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
    <h2>2. 비상연락망 · 관할기관</h2>
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

  return `
  <div class="split">
    <section class="card">
      <h2>계약 목록</h2>
      <p class="counts">진행 ${n진행}건 · 완료 ${n완료}건 · 총 ${data.contracts.length}건</p>
      <div class="toolbar"><button class="btn navy" data-act="add">+ 계약 추가</button></div>
      <ul class="clist">${list || `<li class="empty-note">아직 계약이 없습니다.</li>`}</ul>
    </section>
    ${sel ? viewContractForm(sel) : `<section class="card"><p class="empty-note">왼쪽 [+ 계약 추가]를 누르세요.</p></section>`}
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
    <h2>${esc(c.계약명 || "새 계약")}</h2>
    <p class="hint">계약명을 넣으면 <b>계약유형 · 작업유형 · 계획서 서식</b>이 자동으로 정해집니다.</p>
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
      <button class="btn grey" data-act="reauto">자동 판정으로 되돌리기</button></div>
    <p class="legend"><i style="background:var(--auto)"></i>자동으로 채운 값 <i style="background:var(--manual)"></i>직접 정한 값(자동 판정이 건드리지 않음)</p>
  </section>
  <section class="card">
    <div class="toolbar" style="margin:0">
      <button class="btn navy" data-act="to-output">▶ 이 계약 출력하러 가기</button>
      <button class="btn danger" data-act="delete">이 계약 삭제</button>
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
  if (!data.contracts.length) return `<section class="card"><p class="empty-note">② 계약정보에서 계약을 먼저 넣어 주세요.</p></section>`;
  const today = todayIso();
  const head = FORMS.slice(1).map((f) =>
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

  return `
  <section class="card">
    <h2>출력</h2>
    <p class="hint">왼쪽 칸으로 계약을 고르고 버튼을 누르면 <b>한글(hwpx) 파일</b>이 내려받아집니다. 여러 개면 zip 하나로 묶습니다.
      ○ 는 그 계약에 갖춰야 할 서식입니다(칸을 눌러 바꿀 수 있습니다).</p>
    <div class="toolbar">${BTNS.map((b) => `<button class="btn ${b.cls}" data-export="${b.act}">${b.label}</button>`).join("")}</div>
    <div class="table-wrap"><table class="out">
      <thead><tr><th><input type="checkbox" data-checkall style="width:20px;height:20px"/></th><th>연번</th><th>계약명</th><th>기간</th><th>점검요일</th>${head}</tr></thead>
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
  if (linkSupported()) {
    const saved = await loadLink();
    if (saved?.handle) {
      link = saved;
      await connect(saved.handle); // 권한이 남아 있으면 바로 맞추고, 아니면 [다시 켜기] 버튼
    }
  }
})();
