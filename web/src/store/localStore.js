// 이 PC 의 브라우저 안(IndexedDB)에 저장한다. 밖으로 나가지 않는다.
// IndexedDB 를 못 쓰는 환경(사생활 보호 창 등)이면 localStorage, 그것도 안 되면 메모리.

const DB = "safety-form";
const STORE = "kv";
const KEY = "data";

function openDb() {
  return new Promise((resolve, reject) => {
    const req = indexedDB.open(DB, 1);
    req.onupgradeneeded = () => req.result.createObjectStore(STORE);
    req.onsuccess = () => resolve(req.result);
    req.onerror = () => reject(req.error);
  });
}

let memory = null;

// 작은 설정값(자동 저장 대상 정보 등) — 같은 저장소의 다른 키
export async function kvGet(key) {
  try {
    const db = await openDb();
    return await new Promise((resolve, reject) => {
      const r = db.transaction(STORE).objectStore(STORE).get(key);
      r.onsuccess = () => resolve(r.result ?? null);
      r.onerror = () => reject(r.error);
    });
  } catch { return null; }
}
export async function kvPut(key, value) {
  try {
    const db = await openDb();
    await new Promise((resolve, reject) => {
      const tx = db.transaction(STORE, "readwrite");
      value === undefined ? tx.objectStore(STORE).delete(key) : tx.objectStore(STORE).put(value, key);
      tx.oncomplete = resolve;
      tx.onerror = () => reject(tx.error);
    });
  } catch { /* 설정값만 못 남김 */ }
}
export const kvDel = (key) => kvPut(key, undefined);

export async function loadLocal() {
  try {
    const db = await openDb();
    return await new Promise((resolve, reject) => {
      const r = db.transaction(STORE).objectStore(STORE).get(KEY);
      r.onsuccess = () => resolve(r.result ?? null);
      r.onerror = () => reject(r.error);
    });
  } catch {
    try {
      const s = localStorage.getItem(DB);
      return s ? JSON.parse(s) : memory;
    } catch {
      return memory;
    }
  }
}

export async function saveLocal(data) {
  memory = data;
  try {
    const db = await openDb();
    await new Promise((resolve, reject) => {
      const tx = db.transaction(STORE, "readwrite");
      tx.objectStore(STORE).put(data, KEY);
      tx.oncomplete = resolve;
      tx.onerror = () => reject(tx.error);
    });
    return true;
  } catch {
    try { localStorage.setItem(DB, JSON.stringify(data)); return true; } catch { return false; }
  }
}

export function emptyData() {
  return {
    version: 1,
    // 기관코드 = 나라장터 수요기관코드(계약 가져오기에 씀) · 번호기관 = 고유번호가 어느 기관 것인지(화면엔 안 보임)
    org: { 기관구분: "학교", 기관명: "", 고유번호: "", 주소: "", 전화번호: "", 팩스번호: "", 기관코드: "", 번호기관: "" },
    sign: {
      담당자직위: "주무관", 담당자성명: "", 담당자전화: "",
      관리자직위: "행정실장", 관리자성명: "",
      부서장직위: "교장", 부서장성명: "",
      대표자직위: "교장", 대표자성명: "",
    },
    emerg: { 소방서: "", 소방서전화: "", 병원: "", 병원전화: "", 노동지청: "", 노동지청전화: "" },
    stash: {},
    contracts: [],
    updatedAt: null,
    backedUpAt: null,
  };
}

// 불러온 파일이 이 프로그램의 것인지, 빠진 칸은 없는지 다듬는다
export function normalize(raw) {
  if (!raw || typeof raw !== "object" || !Array.isArray(raw.contracts)) {
    throw new Error("이 프로그램에서 저장한 파일이 아닙니다");
  }
  const base = emptyData();
  return {
    ...base,
    ...raw,
    org: { ...base.org, ...raw.org },
    sign: { ...base.sign, ...raw.sign },
    emerg: { ...base.emerg, ...raw.emerg },
    stash: raw.stash || {},
  };
}
