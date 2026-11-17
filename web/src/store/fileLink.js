// 자동 저장 파일 연결 ― 구글 드라이브 데스크톱 폴더(G:\내 드라이브 등)에 있는 .json 파일 하나를
// 골라 두면, 고칠 때마다 그 파일에 바로 쓴다. 드라이브 프로그램이 알아서 클라우드로 올린다.
// 브라우저의 File System Access API 만 쓰므로 밖으로 나가는 요청이 없다(구글 로그인 없음).
// 크롬·엣지·웨일에서만 된다. 안 되는 브라우저는 [내 정보 파일로 저장]으로 대신한다.

const DB = "safety-form";
const STORE = "kv";
const KEY = "link"; // { handle, name, syncedAt }

export const FILE_NAME = "중대재해예방서식_데이터.json";
const TYPES = [{ description: "중대재해예방 서식 데이터", accept: { "application/json": [".json"] } }];

export const linkSupported = () =>
  typeof window !== "undefined" && "showSaveFilePicker" in window && "showOpenFilePicker" in window;

function openDb() {
  return new Promise((resolve, reject) => {
    const req = indexedDB.open(DB, 1);
    req.onupgradeneeded = () => req.result.createObjectStore(STORE);
    req.onsuccess = () => resolve(req.result);
    req.onerror = () => reject(req.error);
  });
}

async function idb(mode, fn) {
  const db = await openDb();
  return new Promise((resolve, reject) => {
    const tx = db.transaction(STORE, mode);
    const r = fn(tx.objectStore(STORE));
    tx.oncomplete = () => resolve(r?.result);
    tx.onerror = () => reject(tx.error);
  });
}

export async function loadLink() {
  try { return (await idb("readonly", (s) => s.get(KEY))) ?? null; } catch { return null; }
}
export async function saveLink(link) {
  try { await idb("readwrite", (s) => s.put(link, KEY)); } catch { /* 연결 정보만 못 남김 */ }
}
export async function clearLink() {
  try { await idb("readwrite", (s) => s.delete(KEY)); } catch { /* 무시 */ }
}

// 'granted' | 'prompt' | 'denied'. ask=true 면 권한 창을 띄운다(버튼 누름 안에서만 가능).
export async function permission(handle, ask = false) {
  const opts = { mode: "readwrite" };
  try {
    let p = await handle.queryPermission(opts);
    if (p !== "granted" && ask) p = await handle.requestPermission(opts);
    return p;
  } catch {
    return "denied";
  }
}

// 새 파일 만들기(처음 쓰는 컴퓨터)
export async function pickNew() {
  return window.showSaveFilePicker({ suggestedName: FILE_NAME, types: TYPES, id: "safety-form-data" });
}

// 있는 파일 열기(다른 컴퓨터에서 만든 파일)
export async function pickExisting() {
  const [h] = await window.showOpenFilePicker({ types: TYPES, multiple: false, id: "safety-form-data" });
  return h;
}

// 파일 내용. 빈 파일이면 null
export async function readFile(handle) {
  const f = await handle.getFile();
  const text = (await f.text()).trim();
  if (!text) return null;
  return JSON.parse(text);
}

export async function writeFile(handle, data) {
  const w = await handle.createWritable();
  await w.write(JSON.stringify(data, null, 1));
  await w.close();
}
