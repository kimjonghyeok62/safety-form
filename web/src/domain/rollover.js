// 계약 복사 · 다음 학년도로 이어가기 · 지난 계약에서 업체 정보 찾기.

import { toDay, parts, contractDays } from "./schedule.js";

const pad = (n) => String(n).padStart(2, "0");
const lastDayOf = (y, m) => new Date(Date.UTC(y, m, 0)).getUTCDate();

// 'YYYY-MM-DD' 를 n 년 뒤로. 말일이면 그 해의 말일로(2월 28일 ↔ 29일).
export function shiftYears(iso, n) {
  const d = toDay(iso);
  if (d == null) return iso || "";
  const p = parts(d);
  const y = p.y + n;
  const last = lastDayOf(y, p.m);
  const day = p.d === lastDayOf(p.y, p.m) ? last : Math.min(p.d, last);
  return `${y}-${pad(p.m)}-${pad(day)}`;
}

// 계약명 안의 연도를 n 만큼 ― "2026학년도" "2026년" "2026." 꼴의 20xx 만
export function shiftNameYears(name, n = 1) {
  return String(name || "").replace(/(20\d{2})(?=\s*(학년도|년|회계연도|\.|\s|$))/g, (y) => String(+y + n));
}

// 복사본 ― 진행 기록과 직접 고른 진행/완료는 비운다
export function cloneContract(c, newId) {
  const x = structuredClone(c);
  x.id = newId;
  x.done = {};
  x.비고 = { v: "", manual: false };
  delete x.from;
  return x;
}

export function rollForward(c, newId, years = 1) {
  const x = cloneContract(c, newId);
  x.계약명 = shiftNameYears(c.계약명, years);
  x.시작일 = shiftYears(c.시작일, years);
  x.종료일 = shiftYears(c.종료일, years);
  x.from = c.id;
  return x;
}

// 이어가기 후보: 기간이 300일 넘는 연간 계약을 먼저(기본 체크)
export const isAnnual = (c) => contractDays(c) >= 300;

// 업체 주소록 ― 계약들에 적힌 업체를 업체명으로 묶는다(가장 최근 계약의 값)
export const VENDOR_KEYS = ["사업자번호", "업체대표", "업체주소", "대표자전화", "안전담당자", "담당자전화"];
export function vendorBook(contracts) {
  const book = new Map();
  const sorted = [...contracts].sort((a, b) => String(a.시작일 || "").localeCompare(String(b.시작일 || "")));
  for (const c of sorted) {
    const nm = String(c.업체명 || "").trim();
    if (!nm) continue;
    const prev = book.get(nm) || {};
    const next = { ...prev, 업체명: nm, n: (prev.n || 0) + 1 };
    for (const k of VENDOR_KEYS) if (String(c[k] || "").trim()) next[k] = c[k];
    book.set(nm, next);
  }
  return book;
}
