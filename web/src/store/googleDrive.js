// 구글 계정으로 이어쓰기 ― 그 사람 본인 구글 드라이브의 "앱 전용 숨김 폴더"(appDataFolder)에
// 파일 하나를 둔다. 브라우저 ↔ 구글 직통이라 개발자 쪽을 거치지 않는다.
//   - 권한 drive.appdata: 이 앱이 만든 숨김 폴더만 본다. 드라이브의 다른 파일은 못 본다.
//   - 권한 email: 어느 계정에 저장 중인지 보여 주고, 다음 로그인 때 계정 고르기를 건너뛰려고.
// 구글 로그인 도구(GIS)는 이 기능을 쓸 때만 받아 온다.
import { GOOGLE_CLIENT_ID } from "../config.js";
import { kvGet, kvPut, kvDel } from "./localStore.js";

const DRIVE_SCOPE = "https://www.googleapis.com/auth/drive.appdata";
const SCOPE = `${DRIVE_SCOPE} openid email`;
const NAME = "safety-form-data.json";
const KEY = "gdrive"; // { email, fileId, syncedAt, savedAt }
const TOKEN_KEY = "safety-form-gtoken"; // sessionStorage ― 새로고침해도 한 시간 안에는 다시 안 묻게

export const googleReady = () => !!GOOGLE_CLIENT_ID;
export const loadGoogleMeta = () => kvGet(KEY);
export const clearGoogleMeta = () => kvDel(KEY);

const authError = (msg) => Object.assign(new Error(msg), { code: "auth" });

// ------------------------------------------------------------ 토큰
let token = null;
let exp = 0;
try {
  const t = JSON.parse(sessionStorage.getItem(TOKEN_KEY) || "null");
  if (t && t.exp > Date.now()) { token = t.token; exp = t.exp; }
} catch { /* 무시 */ }

export const tokenValid = (marginMs = 0) => !!token && Date.now() + marginMs < exp;

function forgetToken() {
  token = null; exp = 0;
  try { sessionStorage.removeItem(TOKEN_KEY); } catch { /* 무시 */ }
}

let gisPromise = null;
export function loadGis() {
  if (window.google?.accounts?.oauth2) return Promise.resolve();
  gisPromise ||= new Promise((resolve, reject) => {
    const s = document.createElement("script");
    s.src = "https://accounts.google.com/gsi/client";
    s.async = true;
    s.onload = () => resolve();
    s.onerror = () => { gisPromise = null; reject(new Error("구글 로그인 도구를 받지 못했습니다. 인터넷 연결을 확인해 주세요.")); };
    document.head.appendChild(s);
  });
  return gisPromise;
}

let client = null;
let waiting = null;
function tokenClient() {
  client ||= window.google.accounts.oauth2.initTokenClient({
    client_id: GOOGLE_CLIENT_ID,
    scope: SCOPE,
    callback: (r) => {
      const w = waiting; waiting = null;
      if (!w) return;
      if (r.error) return w.reject(authError(r.error_description || r.error));
      if (!window.google.accounts.oauth2.hasGrantedAllScopes(r, DRIVE_SCOPE)) {
        return w.reject(authError("드라이브 저장 허락에 체크해야 저장할 수 있습니다. 다시 로그인하면서 체크해 주세요."));
      }
      token = r.access_token;
      exp = Date.now() + (Number(r.expires_in || 3600) - 120) * 1000;
      try { sessionStorage.setItem(TOKEN_KEY, JSON.stringify({ token, exp })); } catch { /* 무시 */ }
      w.resolve(token);
    },
    error_callback: (e) => {
      const w = waiting; waiting = null;
      if (!w) return;
      const msg = e?.type === "popup_closed" ? "로그인 창을 닫아 멈췄습니다."
        : e?.type === "popup_failed_to_open" ? "로그인 창이 막혔습니다. 주소창 오른쪽의 팝업 차단을 풀어 주세요."
        : "구글 로그인에 실패했습니다.";
      w.reject(authError(msg));
    },
  });
  return client;
}

// 버튼 누름(사용자 동작) 안에서 불러야 로그인 창이 막히지 않는다.
// hint(이메일)가 있으면 계정 고르기 없이 바로 닫힌다.
export async function signIn(hint) {
  await loadGis();
  return new Promise((resolve, reject) => {
    waiting?.reject(authError("다시 시도합니다"));
    waiting = { resolve, reject };
    tokenClient().requestAccessToken(hint ? { prompt: "", login_hint: hint } : { prompt: "select_account" });
  });
}

export function signOut() {
  forgetToken();
}

async function api(url, opts = {}) {
  if (!tokenValid()) throw authError("구글 로그인 시간이 지났습니다");
  // 응답이 오래 안 오면 끊는다 ― 줄이 영영 막히지 않게(멈춤 → [다시 시도])
  const ctl = new AbortController();
  const timer = setTimeout(() => ctl.abort(), 30000);
  let r;
  try {
    r = await fetch(url, { ...opts, signal: ctl.signal, cache: "no-store",
      headers: { ...(opts.headers || {}), Authorization: `Bearer ${token}` } });
  } catch (err) {
    throw new Error(err?.name === "AbortError" ? "구글 드라이브가 응답하지 않습니다(인터넷 연결 확인)" : "구글 드라이브에 연결하지 못했습니다");
  } finally {
    clearTimeout(timer);
  }
  if (r.status === 401) { forgetToken(); throw authError("구글 로그인 시간이 지났습니다"); }
  if (r.status === 404) { const e = new Error("파일 없음"); e.code = "404"; throw e; }
  if (!r.ok) {
    let msg = `구글 드라이브 오류(${r.status})`;
    try { const j = await r.json(); if (j?.error?.message) msg += `: ${j.error.message}`; } catch { /* 무시 */ }
    throw new Error(msg);
  }
  return r;
}

export async function whoAmI() {
  const r = await api("https://www.googleapis.com/oauth2/v3/userinfo");
  return (await r.json()).email || "";
}

// ------------------------------------------------------------ 파일 하나(앱 숨김 폴더)
async function findFileId() {
  const q = encodeURIComponent(`name='${NAME}' and trashed=false`);
  const r = await api(`https://www.googleapis.com/drive/v3/files?spaces=appDataFolder&q=${q}&orderBy=modifiedTime%20desc&fields=files(id)`);
  return (await r.json()).files?.[0]?.id ?? null;
}

async function createFile(json) {
  const b = "sf" + Math.random().toString(36).slice(2);
  const body =
    `--${b}\r\nContent-Type: application/json; charset=UTF-8\r\n\r\n` +
    JSON.stringify({ name: NAME, parents: ["appDataFolder"] }) +
    `\r\n--${b}\r\nContent-Type: application/json; charset=UTF-8\r\n\r\n${json}\r\n--${b}--`;
  const r = await api("https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart&fields=id", {
    method: "POST", headers: { "Content-Type": `multipart/related; boundary=${b}` }, body,
  });
  return (await r.json()).id;
}

// 자동 저장 대상 하나(sync.js 가 쓴다)
export function googleTarget(meta) {
  return {
    kind: "google",
    where: "구글 드라이브",
    get name() { return meta.email || "구글 계정"; },
    meta,
    saveMeta: () => kvPut(KEY, meta),
    async read() {
      meta.fileId ||= await findFileId();
      if (!meta.fileId) return null;
      try {
        const r = await api(`https://www.googleapis.com/drive/v3/files/${meta.fileId}?alt=media`);
        const text = (await r.text()).trim();
        if (!text) return null;
        try { return JSON.parse(text); } catch { throw new Error("구글 드라이브의 저장 파일이 깨져 있습니다"); }
      } catch (err) {
        if (err.code !== "404") throw err;
        meta.fileId = await findFileId();     // 지워졌으면 다시 찾고, 없으면 새로 만든다
        return meta.fileId ? this.read() : null;
      }
    },
    async write(data) {
      const json = JSON.stringify(data);
      if (meta.fileId) {
        try {
          await api(`https://www.googleapis.com/upload/drive/v3/files/${meta.fileId}?uploadType=media&fields=id`, {
            method: "PATCH", headers: { "Content-Type": "application/json; charset=UTF-8" }, body: json,
          });
          return;
        } catch (err) {
          if (err.code !== "404") throw err;
        }
      }
      meta.fileId = await createFile(json);
    },
  };
}
