// 자동 저장(sync) 시험 ― 가짜 대상으로 "초기화했는데 드라이브 옛 내용이 되돌아오는" 일이 없는지 본다.
//   node test/sync-reset.mjs
import assert from "node:assert/strict";
import { createSync } from "../src/sync.js";
import { emptyData } from "../src/store/localStore.js";

const clone = (x) => JSON.parse(JSON.stringify(x));
let tick = 0;
const stamp = () => new Date(Date.UTC(2026, 10, 19, 0, 0, tick++)).toISOString();

function setup(local, remote, syncedAt) {
  const app = { data: clone(local), toasts: [], picks: [] };
  const target = {
    where: "구글 드라이브", meta: { syncedAt },
    file: remote ? clone(remote) : null,
    async read() { return this.file ? clone(this.file) : null; },
    async write(d) { this.file = clone(d); },
    async saveMeta() {},
  };
  const sync = createSync({
    getData: () => app.data, adopt: (n) => { app.data = n; }, persistLocal: async () => {},
    render() {}, toast: (m) => app.toasts.push(m), backup() {}, onState() {},
    choose: async (title) => { app.picks.push(title); return null; },
  });
  return { app, target, sync };
}

function withContracts(n, org = {}) {
  const d = emptyData();
  Object.assign(d.org, org);
  d.contracts = Array.from({ length: n }, (_, i) => ({ id: "c" + i, 계약명: "계약" + i }));
  d.updatedAt = stamp();
  return d;
}

// 1) 기관명이 빈 채로 계약을 쓰다가 [목록 초기화] → 드라이브에 빈 목록이 올라가야 한다
{
  const before = withContracts(2, { 기관구분: "교육지원청", 고유번호: "123-45-00000" });
  before.org.기관명 = "";
  const { app, target, sync } = setup(before, before, before.updatedAt);
  await sync.attach(target);
  app.data.trash = { contracts: { at: stamp(), why: "초기화", list: app.data.contracts } };
  app.data.contracts = [];
  app.data.updatedAt = stamp();
  await sync.queue("write");
  assert.equal(app.data.contracts.length, 0, "초기화한 목록이 드라이브 내용으로 되돌아오면 안 된다");
  assert.equal(target.file.contracts.length, 0, "드라이브에도 빈 목록이 저장되어야 한다");
  assert.ok(target.file.trash?.contracts, "되살리기용 휴지통도 함께 올라가야 한다");
  console.log("ok 1  기관명 없이 목록 초기화 → 드라이브에 저장됨");
}

// 2) 기관명까지 모두 비운 뒤(기관정보 초기화 + 목록 초기화)에도 마찬가지
{
  const before = withContracts(3, { 기관명: "예시초등학교" });
  const { app, target, sync } = setup(before, before, before.updatedAt);
  await sync.attach(target);
  app.data = emptyData();
  app.data.trash = { org: { at: stamp(), org: before.org }, contracts: { at: stamp(), list: before.contracts } };
  app.data.updatedAt = stamp();
  await sync.queue("write");
  assert.equal(app.data.contracts.length, 0);
  assert.equal(target.file.org.기관명, "");
  console.log("ok 2  기관정보 · 목록 모두 초기화 → 드라이브에 저장됨");
}

// 3) 처음 쓰는 컴퓨터(한 번도 고친 적 없음) + 로그인 → 드라이브 내용을 가져와야 한다
{
  const remote = withContracts(4, { 기관명: "예시초등학교" });
  const { app, target, sync } = setup(emptyData(), remote, null);
  await sync.attach(target);
  assert.equal(app.data.contracts.length, 4);
  console.log("ok 3  새 컴퓨터에서 로그인 → 드라이브 내용을 가져옴");
}

// 4) 다른 컴퓨터에서 고친 것만 있으면 가져온다
{
  const base = withContracts(1, { 기관명: "예시초등학교" });
  const remote = clone(base); remote.contracts.push({ id: "x", 계약명: "새 계약" }); remote.updatedAt = stamp();
  const { app, target, sync } = setup(base, remote, base.updatedAt);
  await sync.attach(target);
  assert.equal(app.data.contracts.length, 2);
  console.log("ok 4  다른 컴퓨터에서 고친 것 → 가져옴");
}

// 5) 이쪽에서만 고쳤으면 올린다
{
  const base = withContracts(2, { 기관명: "예시초등학교" });
  const { app, target, sync } = setup(base, base, base.updatedAt);
  await sync.attach(target);
  app.data.contracts.pop(); app.data.updatedAt = stamp();
  await sync.queue("write");
  assert.equal(target.file.contracts.length, 1);
  console.log("ok 5  이 컴퓨터에서 고친 것 → 올림");
}
console.log("all passed");
