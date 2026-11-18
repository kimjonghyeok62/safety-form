// 계약 한 건에 학교(도급인)가 할 일 ― 언제까지, 무엇을, 어느 서식으로.
// 근거: 경기도교육청 「중대산업재해 예방 도급·용역·위탁 사업 안전·보건 확보 방안」(2024.4.)
//       「교육현장의 안전을 위한 중대재해처벌법 업무 매뉴얼」(2025.11.) · 학교안전과-9506(2026)
// 할지 말지의 기준일수는 서식 판정(formNeeded)과 같은 settings.json 값을 쓴다.
//
// 한 일은 계약에 c.done = { [key]: "YYYY-MM-DD"(체크한 날) } 로 남긴다.
// key 는 날짜가 조금 바뀌어도 이어지게 잡는다 ― 협의체는 달, 합동점검은 회차, 순회점검은 그날.

import { WORK_KEYS, WORK_NAMES, crit, SETTINGS } from "./dictionary.js";
import { contractDays, toDay, dayToIso, inspectionDates, councilDates, parts } from "./schedule.js";

export const PHASES = [
  { id: "pre", label: "계약할 때" },
  { id: "before", label: "착공 전" },
  { id: "work", label: "작업 전" },
  { id: "during", label: "작업 중" },
  { id: "end", label: "마칠 때" },
];

// 작업유형으로 정하는 수급인 안전보건 수준 등급(확보 방안 붙임1)
// 화재·폭발 우려장소 · 밀폐공간 A(80) / 그 밖의 산재 위험장소 B(70) / 일반작업 C(60)
export function requiredGrade(c) {
  const on = (k) => !!c.work?.[k]?.v;
  if (on("화기") || on("밀폐")) return { grade: "A", score: 80 };
  if (on("고소") || on("전기")) return { grade: "B", score: 70 };
  return { grade: "C", score: 60 };
}

const pack = (s) => String(s || "").replace(/\s/g, "");
export function isConstruction(c) {
  const kind = String(c.계획서유형 || "");
  if (kind === "공사") return true;
  if (kind === "용역") return false;
  const nm = pack(c.계약명);
  const kw = (s) => String(s || "").split(",").map(pack).filter(Boolean);
  if (kw(SETTINGS.계획서키워드.용역).some((k) => nm.includes(k))) return false;
  return kw(SETTINGS.계획서키워드.공사).some((k) => nm.includes(k));
}

const money = (c) => Number(String(c.계약금액 || "").replace(/\D/g, "")) || 0;

/**
 * 계약 하나의 할 일 목록. 날짜가 없으면 날짜 칸은 null.
 * @returns {{key, phase, title, sub, due, form, group}[]}  due = 일 번호(1970-01-01 부터)
 */
export function dutiesOf(c) {
  if (!String(c.계약명 || "").trim()) return [];
  const days = contractDays(c);
  if (!days) return [];
  const s = toDay(c.시작일), e = toDay(c.종료일);
  const out = [];
  const add = (o) => out.push({ form: null, group: null, sub: "", ...o });

  // 계약할 때 ― 모든 도급(금액 · 기간 무관)
  add({ key: "pre:clause", phase: "pre", due: s, title: "과업지시서 · 특수조건에 안전 · 보건 확보 의무 4개 문구",
    sub: "법령 준수 · 유해위험 신고 시 즉시 개선 · 안전대책 후 작업 · 중대재해 선보고 후 처리" });
  add({ key: "pre:pledge", phase: "pre", due: s, form: 15, title: "계약이행 통합 서약서 받기",
    sub: "9번 「안전 및 보건 확보의무 준수」 포함" });

  // 착공 전 ― 2일 이상(연 단위 계약 포함)
  if (days >= crit("기준_계획서")) {
    add({ key: "before:plan", phase: "before", due: s, form: 1, title: "안전보건관리계획서 받기",
      sub: "업체가 작성 · 제출" });
  }
  if (days >= crit("기준_수준")) {
    const g = requiredGrade(c);
    add({ key: "before:eval", phase: "before", due: s, form: 2, title: `수급인 안전보건 수준 평가 — ${g.grade}등급(${g.score}점) 이상`,
      sub: "미달이면 계획서 보완 요구 · 산재요양확인서로 최근 2년 재해 확인" });
  }

  // 작업 전 ― 금액과 관계없이 모든 도급
  if (days >= crit("기준_도급")) {
    add({ key: "work:check", phase: "work", due: s, form: 3, title: "공사(용역) 안전 · 보건 점검표",
      sub: "작업 시작 전 · 보호구 착용, 사다리 2인 1조 · 아웃트리거 확인" });
  }
  WORK_KEYS.forEach((k, i) => {
    if (!c.work?.[k]?.v) return;
    add({ key: `work:${k}`, phase: "work", due: s, form: 4 + i, title: `${WORK_NAMES[i]} 점검표 받기`,
      sub: "업체가 작성해 작업 전에 제출" });
  });

  // 작업 중 ― 정기 점검(서식 판정과 같은 날짜 계산)
  if (days >= crit("기준_협의체")) {
    councilDates(c).forEach((d, i) => {
      const p = parts(d);
      add({ key: `council:${p.y}-${String(p.m).padStart(2, "0")}`, phase: "during", due: d, form: 8, group: "council",
        title: `안전보건협의체 ${i + 1}회차`, sub: "매월 1회 이상 · 회의록" });
    });
  }
  if (days >= crit("기준_순회")) {
    inspectionDates(c, false).forEach((d, i) => {
      add({ key: `patrol:${dayToIso(d)}`, phase: "during", due: d, form: 9, group: "patrol",
        title: `작업장 순회점검 ${i + 1}회차`, sub: "주 1회 이상" });
    });
  }
  if (days >= crit("기준_합동")) {
    inspectionDates(c, true).forEach((d, i) => {
      add({ key: `joint:${i + 1}`, phase: "during", due: d, form: 10, group: "joint",
        title: `작업장 합동 안전 · 보건 점검 ${i + 1}회차`, sub: "분기 1회 이상 · 도급인 · 수급인 · 양쪽 근로자" });
    });
  }

  // 마칠 때
  if (isConstruction(c) && money(c) >= 20000000) {
    add({ key: "end:cost", phase: "end", due: e, title: "산업안전보건관리비 실사용 확인 · 정산",
      sub: "준공검사 때 사용내역서와 증빙 확인 — 쓴 만큼만 지급" });
  }
  return out;
}

export const GROUP_LABEL = { council: "안전보건협의체 (매월)", patrol: "작업장 순회점검 (매주)", joint: "합동 점검 (분기)" };

// 할 일 하나의 상태: done · late(지났는데 안 함) · soon(7일 안) · later · none(날짜 없음)
export function stateOf(c, d, today) {
  if (c.done?.[d.key]) return "done";
  if (d.due == null) return "none";
  if (d.due < today) return "late";
  if (d.due <= today + 7) return "soon";
  return "later";
}

// 계약 하나의 요약 ― 지금까지 했어야 할 것 중 한 것의 비율
export function summaryOf(c, today) {
  const list = dutiesOf(c);
  let due = 0, done = 0, late = 0, soon = 0, all = list.length, doneAll = 0;
  for (const d of list) {
    const st = stateOf(c, d, today);
    if (st === "done") doneAll++;
    if (d.due != null && d.due <= today) { due++; if (st === "done") done++; }
    if (st === "late") late++;
    if (st === "soon") soon++;
  }
  return { list, all, doneAll, due, done, late, soon, rate: due ? Math.round((done / due) * 100) : null };
}

// 학년도(3.1.~이듬해 2월 말) ― today 가 1·2월이면 전년도 학년도
export function schoolYearOf(day) {
  const p = parts(day);
  return p.m >= 3 ? p.y : p.y - 1;
}
export function schoolYearRange(y) {
  return { from: toDay(`${y}-03-01`), to: toDay(`${y + 1}-03-01`) - 1 };
}

// 반기: 1 = 1~6월, 2 = 7~12월 (예산 · 점검 보고가 역년 반기라)
export function halfRange(y, h) {
  return h === 1
    ? { from: toDay(`${y}-01-01`), to: toDay(`${y}-06-30`) }
    : { from: toDay(`${y}-07-01`), to: toDay(`${y}-12-31`) };
}
