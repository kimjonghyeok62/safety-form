// 내 컴퓨터 파일에 자동 저장 ― .json 파일 하나를 골라 두면 고칠 때마다 그 파일에 바로 쓴다.
// 구글 드라이브 데스크톱 폴더(G:\내 드라이브)나 USB 에 두면 다른 PC 에서도 이어 쓸 수 있다.
// 브라우저의 File System Access API 만 쓰므로 밖으로 나가는 요청이 없다.
// 크롬·엣지·웨일에서만 된다.
import { kvGet, kvPut, kvDel } from "./localStore.js";

const KEY = "link"; // { handle, name, syncedAt, savedAt }

export const FILE_NAME = "중대재해예방서식_데이터.json";
const TYPES = [{ description: "중대재해 예방 서식 데이터", accept: { "application/json": [".json"] } }];

export const linkSupported = () =>
  typeof window !== "undefined" && "showSaveFilePicker" in window && "showOpenFilePicker" in window;

export const loadLink = () => kvGet(KEY);
export const saveLink = (link) => kvPut(KEY, link);
export const clearLink = () => kvDel(KEY);

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

export const pickNew = () =>
  window.showSaveFilePicker({ suggestedName: FILE_NAME, types: TYPES, id: "safety-form-data" });

export async function pickExisting() {
  const [h] = await window.showOpenFilePicker({ types: TYPES, multiple: false, id: "safety-form-data" });
  return h;
}

// 자동 저장 대상 하나(sync.js 가 쓴다)
export function fileTarget(link) {
  return {
    kind: "file",
    where: "저장 파일",
    get name() { return link.name; },
    meta: link,
    saveMeta: () => saveLink(link),
    async read() {
      if ((await permission(link.handle)) !== "granted") {
        const e = new Error("파일 쓰기 허락이 풀렸습니다"); e.code = "auth"; throw e;
      }
      let text;
      try {
        text = (await (await link.handle.getFile()).text()).trim();
      } catch (err) {
        if (err?.name === "NotFoundError") {
          const e = new Error("저장 파일을 찾지 못했습니다(옮겨졌거나 지워짐)"); e.code = "gone"; throw e;
        }
        throw err;
      }
      if (!text) return null;
      try { return JSON.parse(text); } catch { throw new Error("이 프로그램의 저장 파일이 아닙니다"); }
    },
    async write(data) {
      const w = await link.handle.createWritable();
      await w.write(JSON.stringify(data, null, 1));
      await w.close();
    },
  };
}
