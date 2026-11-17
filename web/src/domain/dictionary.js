// 계약유형 사전 · 작업유형 판정 · 서식 필요 판정 · 계획서 서식 고르기.
// 원본: modConfig.bas (FindContractType · DetectWorkTypes · DictWorkTypes · ResolvePlanTemplate)
//       modLayout.bas (ActiveWorkTypes)  build_workbook.py (form_formula)
//
// 사전 데이터는 contract_types.py → _scripts/export_web_data.py 가 만든 JSON 이다.
// 여기서 직접 고치지 말 것.

import CONTRACT_TYPES from "../data/contract-types.json" with { type: "json" };
import CONTENT_LIBRARY from "../data/content-library.json" with { type: "json" };
import SETTINGS from "../data/settings.json" with { type: "json" };
import { contractDays } from "./schedule.js";

export { CONTRACT_TYPES, CONTENT_LIBRARY, SETTINGS };

export const WORK_KEYS = ["고소", "화기", "밀폐", "전기"];
export const WORK_NAMES = ["고소작업", "화기작업", "밀폐공간작업", "전기작업"];

const typeIndex = new Map();
for (const t of CONTRACT_TYPES) if (t.유형명 && !typeIndex.has(t.유형명)) typeIndex.set(t.유형명, t);

const pack = (s) => String(s || "").trim().replace(/ /g, "");
const kwList = (csv) => String(csv || "").replace(/，/g, ",").split(",").map(pack).filter(Boolean);

// 영문 대소문자는 가리지 않는다(VBA InStr vbTextCompare 와 같게)
const has = (hay, kw) => hay.toLowerCase().includes(kw.toLowerCase());

function longestHit(packedName, csv) {
  let best = 0;
  for (const kw of kwList(csv)) if (has(packedName, kw) && kw.length > best) best = kw.length;
  return best;
}

function matchesAny(packedName, csv) {
  return kwList(csv).some((kw) => has(packedName, kw));
}

// 계약명 → 유형명. 가장 긴 키워드가 걸린 줄(같으면 위의 줄).
export function contractTypeName(contractName) {
  const nm = pack(contractName);
  if (!nm) return "";
  let best = null, bestLen = 0;
  for (const t of CONTRACT_TYPES) {
    if (!t.유형명) continue;
    const len = longestHit(nm, t.키워드);
    if (len > bestLen) { bestLen = len; best = t; }
  }
  return best ? best.유형명 : "";
}

export function typeField(typeName, field) {
  const t = typeIndex.get(String(typeName || "").trim());
  return t ? String(t[field] || "").trim() : "";
}

function workTypeIdx(name) {
  switch (pack(name)) {
    case "고소작업": case "고소": return 0;
    case "화기작업": case "화기": return 1;
    case "밀폐공간작업": case "밀폐공간": case "밀폐": return 2;
    case "전기작업": case "전기": return 3;
    default: return -1;
  }
}

// 계약명 낱말로 추론한 작업유형 [고소, 화기, 밀폐, 전기]
export function detectWorkTypes(contractName) {
  const hit = [false, false, false, false];
  const nm = pack(contractName);
  if (!nm) return hit;
  for (const row of SETTINGS.작업유형키워드) {
    const i = workTypeIdx(row.유형);
    if (i >= 0 && matchesAny(nm, row.키워드)) hit[i] = true;
  }
  return hit;
}

// 사전이 정해 둔 작업유형. 사전에 하나라도 있으면 그쪽이 이긴다(modAutoDetect.RefreshWorkTypes).
export function suggestedWorkTypes(contractName, typeName) {
  let hit = detectWorkTypes(contractName);
  const spec = typeField(typeName, "작업유형");
  if (spec) {
    const d = [false, false, false, false];
    for (const p of spec.replace(/,/g, "|").replace(/ /g, "").split("|")) {
      const i = workTypeIdx(p);
      if (i >= 0) d[i] = true;
    }
    if (d.some(Boolean)) hit = d;
  }
  return hit;
}

// "고소작업|화기작업" 꼴
export function activeWorkTypes(c) {
  return WORK_KEYS.map((k, i) => (c.work?.[k]?.v ? WORK_NAMES[i] : "")).filter(Boolean).join("|");
}

// 계약 한 건의 '자동 칸' 을 다시 계산한다. 사용자가 직접 정한 칸(manual)은 두고 본다.
export function refreshAuto(c) {
  const nm = String(c.계약명 || "").trim();
  c.계약유형 ??= { v: "", manual: false };
  if (!c.계약유형.manual) c.계약유형.v = nm ? contractTypeName(nm) : "";

  c.work ??= {};
  const hit = suggestedWorkTypes(nm, c.계약유형.v);
  WORK_KEYS.forEach((k, i) => {
    c.work[k] ??= { v: false, manual: false };
    if (!c.work[k].manual) c.work[k].v = nm ? hit[i] : false;
  });

  if (nm && !c.점검요일) c.점검요일 = SETTINGS.기본점검요일;
  if (nm && !c.계획서유형) c.계획서유형 = "자동";
  return c;
}

// 1번 계획서 서식: 직접 고른 것 > 90일↑ 또는 위험작업 2종↑ = 상세 > 공사/용역 낱말 > 상세
export function planTemplate(c) {
  switch (String(c.계획서유형 || "").trim()) {
    case "상세": return "01_계획서_상세.hwpx";
    case "공사": return "01_계획서_공사.hwpx";
    case "용역": return "01_계획서_용역.hwpx";
  }
  const types = activeWorkTypes(c);
  const typeCount = types ? types.split("|").length : 0;
  if (contractDays(c) >= 90 || typeCount >= 2) return "01_계획서_상세.hwpx";
  const nm = pack(c.계약명);
  if (matchesAny(nm, SETTINGS.계획서키워드.공사)) return "01_계획서_공사.hwpx";
  if (matchesAny(nm, SETTINGS.계획서키워드.용역)) return "01_계획서_용역.hwpx";
  return "01_계획서_상세.hwpx";
}

export function planKind(c) {
  return planTemplate(c).replace("01_계획서_", "").replace(".hwpx", "");
}

// 화면 번호 → 서식 파일
export const FORMS = [
  null,
  { no: 1, title: "안전보건관리계획서", file: null },
  { no: 2, title: "수급인 안전보건 수준평가표", file: "02_수준평가표.hwpx" },
  { no: 3, title: "공사(용역) 안전보건 점검표", file: "03_도급점검표.hwpx" },
  { no: 4, title: "고소작업 점검표", file: "04_점검표_고소.hwpx" },
  { no: 5, title: "화기작업 점검표", file: "05_점검표_화기.hwpx" },
  { no: 6, title: "밀폐공간작업 점검표", file: "06_점검표_밀폐.hwpx" },
  { no: 7, title: "전기작업 점검표", file: "07_점검표_전기.hwpx" },
  { no: 8, title: "안전보건협의체 회의록", file: "08_협의체회의록.hwpx" },
  { no: 9, title: "작업장 순회점검표", file: "09_순회점검표.hwpx" },
  { no: 10, title: "작업장 합동점검표", file: "10_합동점검표.hwpx" },
  { no: 11, title: "아차사고 사례 및 유해위험발굴 제안서", file: "11_아차사고제안서.hwpx" },
  { no: 12, title: "관리감독자 업무수행 평가표", file: "12_관리감독자평가표.hwpx" },
  { no: 13, title: "중대재해발생보고서", file: "13_중대재해발생보고서.hwpx" },
  { no: 14, title: "중대재해사고보고서", file: "14_중대재해사고보고서.hwpx" },
  { no: 15, title: "계약이행 통합 서약서", file: "15_계약이행서약서.hwpx" },
];

export function templateFor(formNo, c) {
  return formNo === 1 ? planTemplate(c) : FORMS[formNo]?.file || "";
}

export const crit = (name) => SETTINGS.기준일수[name];

// 출력관리의 ○ 판정. 11~15 는 기간으로 정할 수 없어 늘 false(버튼이 곧 선택).
export function formNeeded(formNo, c) {
  if (!String(c.계약명 || "").trim()) return false;
  const days = contractDays(c);
  switch (formNo) {
    case 1: return days >= crit("기준_계획서");
    case 2: return days >= crit("기준_수준");
    case 3: return days >= crit("기준_도급");
    case 4: case 5: case 6: case 7:
      return days >= 1 && !!c.work?.[WORK_KEYS[formNo - 4]]?.v;
    case 8: return days >= crit("기준_협의체");
    case 9: return days >= crit("기준_순회");
    case 10: return days >= crit("기준_합동");
    default: return false;
  }
}

// 진행 / 완료. 직접 고른 값이 이긴다.
export function progressOf(c, today) {
  if (c.비고?.manual) return c.비고.v || "";
  if (!String(c.계약명 || "").trim() || !c.종료일) return "";
  return c.종료일 < today ? "완료" : "진행";
}

// 기관구분별 직위 기본값: 담당자 · 관리자 · 부서장 · 기관대표자
export function orgPositionDefaults(kind) {
  return /지원청|교육청/.test(String(kind || ""))
    ? ["주무관", "팀장", "과장", "교육장"]
    : ["주무관", "행정실장", "교장", "교장"];
}
