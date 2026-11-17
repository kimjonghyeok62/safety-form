// 자동 저장 ― 이 PC 의 내용과 자동 저장 대상(구글 드라이브 / 내 컴퓨터 파일) 하나를 맞춘다.
//
// 대상의 meta.syncedAt = 양쪽이 마지막으로 같았던 updatedAt.
//   대상만 바뀌었으면 → 가져온다(다른 컴퓨터에서 고친 것)
//   이 PC 만 바뀌었으면 → 쓴다
//   둘 다 바뀌었으면 → 사람이 고른다
// 쓰기 전에 늘 대상을 다시 읽으므로, 다른 컴퓨터가 그 사이 고친 것을 덮어쓰지 않는다.
import { normalize } from "./store/localStore.js";

export function createSync(app) {
  // app: { getData, adopt(next), persistLocal(), render(), choose(title, html, buttons), toast(msg), backup(), onState() }
  const s = { target: null, state: "none", error: "" }; // state: none | prompt | busy | ok | error
  let chain = Promise.resolve();
  const isEmpty = (d) => !d.org.기관명 && !d.contracts.length;
  const fmtWhen = (iso) => iso ? new Date(iso).toLocaleString("ko-KR", {
    year: "numeric", month: "numeric", day: "numeric", hour: "2-digit", minute: "2-digit" }) : "모름";
  const esc = (v) => String(v ?? "").replace(/[&<>"']/g, (ch) =>
    ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[ch]));

  async function writeTarget() {
    // 올리는 동안에도 사용자는 계속 고친다. 올린 그 순간의 사본과 그 시각을 기록해야
    // 다음 차례에 "파일이 바뀌었다"로 잘못 읽어 옛 내용으로 되돌리지 않는다.
    const d = JSON.parse(JSON.stringify(app.getData()));
    await s.target.write(d);
    s.target.meta.syncedAt = d.updatedAt ?? null;
    s.target.meta.savedAt = new Date().toISOString();
    await s.target.saveMeta();
  }

  async function take(file) {
    app.adopt(file);
    await app.persistLocal();
    s.target.meta.syncedAt = file.updatedAt ?? null;
    s.target.meta.savedAt = new Date().toISOString();
    await s.target.saveMeta();
    app.render();
  }

  async function reconcile() {
    const raw = await s.target.read();
    let file = null;
    if (raw) { try { file = normalize(raw); } catch { throw new Error("이 프로그램의 저장 파일이 아닙니다"); } }
    const data = app.getData();
    if (!file) return writeTarget();

    const fileU = file.updatedAt ?? null;
    const localU = data.updatedAt ?? null;
    if (fileU === localU) {
      s.target.meta.syncedAt = localU;
      s.target.meta.savedAt ??= new Date().toISOString();
      return s.target.saveMeta();
    }
    const base = s.target.meta.syncedAt ?? null;
    const fileChanged = fileU !== base;
    const localChanged = localU !== base;

    if (isEmpty(data) || (fileChanged && !localChanged)) {
      await take(file);
      if (!isEmpty(file)) app.toast(`${s.target.where}에서 최신 내용을 가져왔습니다 — 계약 ${file.contracts.length}건`);
      return;
    }
    if (!fileChanged || isEmpty(file)) return writeTarget();

    const pick = await app.choose("어느 내용을 쓸까요?",
      `<p>이 컴퓨터와 ${esc(s.target.where)}의 내용이 서로 다릅니다. 하나를 고르면 다른 쪽은 그 내용으로 바뀝니다.</p>
       <table class="cmp"><tr><th></th><th>마지막 수정</th><th>계약</th></tr>
         <tr><td>${esc(s.target.where)}</td><td>${esc(fmtWhen(fileU))}</td><td>${file.contracts.length}건</td></tr>
         <tr><td>이 컴퓨터</td><td>${esc(fmtWhen(localU))}</td><td>${data.contracts.length}건</td></tr></table>
       <p class="hint">잘 모르겠으면 [이 컴퓨터 내용을 파일로 받기]를 먼저 눌러 두면 이 컴퓨터 내용이 따로 남습니다.</p>`,
      [
        { label: `${s.target.where} 내용 쓰기`, value: "file", cls: "btn-primary" },
        { label: "이 컴퓨터 내용 쓰기", value: "local", cls: "btn-primary" },
        { label: "이 컴퓨터 내용을 파일로 받기", value: "backup", cls: "btn-outline" },
      ]);
    if (pick === "backup") { app.backup(); return reconcile(); }
    if (pick === "file") { await take(file); app.toast(`${s.target.where} 내용으로 바꿨습니다.`); return; }
    if (pick === "local") { await writeTarget(); app.toast(`이 컴퓨터 내용을 ${s.target.where}에 저장했습니다.`); return; }
    throw new Error("어느 내용을 쓸지 고르지 않아 멈췄습니다");
  }

  // mode: "connect" | "write" | "check"
  // 한 번 맞출 때마다 대상을 읽고 쓰므로(구글은 몇 초 걸리기도 한다) 아직 시작 안 한 차례가
  // 이미 줄에 있으면 새로 세우지 않는다 ― 그 차례가 시작할 때 최신 내용을 본다.
  let waiting = null;
  function queue(mode) {
    if (!s.target) return chain;
    // 로그인·허락이 필요한 동안(prompt)은 멈춘다. 일시 오류(error)는 다음 고침 때 다시 해 본다.
    if (mode !== "connect" && s.state !== "ok" && s.state !== "busy" && s.state !== "error") return chain;
    const target = s.target;
    if (mode !== "connect" && waiting === target) return chain;
    waiting = target;
    chain = chain.then(async () => {
      if (waiting === target) waiting = null;
      if (s.target !== target) return;
      s.state = "busy"; app.onState();
      try {
        await reconcile();
        s.state = "ok"; s.error = "";
      } catch (err) {
        if (s.target !== target) return;
        s.state = err?.code === "auth" ? "prompt" : "error";
        s.error = err?.message || String(err);
      }
      app.onState();
    });
    return chain;
  }

  // 대상을 붙이고 바로 맞춘다. 맞추기에 성공하면 true
  async function attach(target) {
    s.target = target;
    await queue("connect");
    return s.state === "ok";
  }

  // 대상을 붙이되 아직 맞추지 않음(로그인·허락이 필요할 때)
  function hold(target, msg) {
    s.target = target; s.state = "prompt"; s.error = msg || ""; app.onState();
  }

  function detach() {
    s.target = null; s.state = "none"; s.error = ""; app.onState();
  }

  return { s, queue, attach, hold, detach, fmtWhen };
}
