// {{태그}} → 값.  원본: modHwpExport.bas BuildTagMap · ApplyCouncilRound · BuildContentBlock
//
// 이름이 다른데 뜻이 같은 태그는 모두 등록한다(서식마다 쓰는 이름이 조금씩 다르다).
// 값이 없는 태그도 반드시 빈 문자열로 등록한다 ― 등록하지 않으면 문서에 {{…}} 가 남는다.

import { CONTENT_LIBRARY, typeField, activeWorkTypes } from "./dictionary.js";
import {
  toDay, todayDay, contractDays, nextWorkday,
  fmtDate, fmtDateDow, parts,
} from "./schedule.js";

// 여러 항목을 한 자리에 넣을 때 쓰는 구분자(한글 찾아바꾸기 시절부터의 규칙)
export const LINE_SEP = "  ·  ";

const s = (v) => (v == null ? "" : String(v));
const trim = (v) => s(v).trim();

function joinTitleName(title, nm) {
  title = trim(title); nm = trim(nm);
  return title && nm ? `${title} ${nm}` : title + nm;
}

function orElse(a, fallback) {
  return trim(a) === "" ? fallback : trim(a);
}

// " · " 로 이은 목록에서 n 번째(1부터). 공백 없는 가운뎃점은 항목 안의 글자라 안 쪼갠다.
export function nthItem(text, n) {
  if (trim(text) === "") return "";
  let k = 0;
  for (const p of s(text).split(" · ")) {
    const v = p.trim();
    if (v) { k += 1; if (k === n) return v; }
  }
  return "";
}

function itemCount(text) {
  if (trim(text) === "") return 0;
  return s(text).split(" · ").filter((p) => p.trim()).length;
}

// 끝을 넘으면 처음으로 돌아온다
function rotateItem(text, n) {
  const cnt = itemCount(text);
  return cnt === 0 ? "" : nthItem(text, ((n - 1) % cnt) + 1);
}

function mergeBlocks(a, b) {
  a = trim(a); b = trim(b);
  if (!a) return b;
  if (!b) return a;
  return a + LINE_SEP + b;
}

// ContentLibrary 에서 '공통' + 그 계약의 작업유형 문구를 모은다
function contentBlock(types, category) {
  const wanted = new Set(["공통", ...(types ? types.split("|") : [])]);
  return CONTENT_LIBRARY
    .filter((r) => wanted.has(trim(r.작업유형)) && trim(r.구분) === category)
    .map((r) => s(r.내용))
    .join(LINE_SEP);
}

function seasonAgenda(day) {
  const key = "월" + String(parts(day).m).padStart(2, "0");
  return CONTENT_LIBRARY
    .filter((r) => trim(r.작업유형) === key && trim(r.구분) === "계절안건")
    .map((r) => s(r.내용))
    .join(LINE_SEP);
}

function roundResolution(roundNo) {
  const acc = CONTENT_LIBRARY.filter((r) => trim(r.구분) === "회차의결").map((r) => s(r.내용));
  return acc.length ? acc[(roundNo - 1) % acc.length] : "";
}

function fmtMoney(n) {
  return Math.round(n).toLocaleString("en-US");
}

/**
 * @param {object} data  저장 데이터 전체(org · sign · emerg)
 * @param {object} c     계약 한 건
 * @param {number} serial 연번(1부터)
 * @param {number} [today] 작성일(일 번호). 시험할 때 고정하려고 받는다.
 */
export function buildTagMap(data, c, serial, today = todayDay()) {
  const d = {};
  const org = data.org || {}, sign = data.sign || {}, em = data.emerg || {};
  const set = (k, v) => { d[`{{${k}}}`] = s(v); };

  // 기관 기본정보
  set("기관구분", org.기관구분);
  set("학교명", org.기관명);
  set("기관명", org.기관명);
  set("고유번호", org.고유번호);
  set("주소", org.주소);
  set("전화번호", org.전화번호);
  set("팩스번호", org.팩스번호);

  // 결재선 · 기관대표자 ― 결재란에는 직위만
  set("담당자", sign.담당자직위);
  set("관리자", sign.관리자직위);
  set("부서장", sign.부서장직위);
  set("담당자직위", sign.담당자직위);
  set("담당자성명", sign.담당자성명);
  set("담당자전화번호", sign.담당자전화);
  set("관리자직위", sign.관리자직위);
  set("관리자성명", sign.관리자성명);
  set("부서장직위", sign.부서장직위);
  set("부서장성명", sign.부서장성명);
  set("기관대표자", joinTitleName(sign.대표자직위, sign.대표자성명));
  set("기관대표자직위", sign.대표자직위);
  set("기관대표자성명", sign.대표자성명);
  // 옛 이름
  set("행정실담당자", sign.담당자성명);
  set("행정실전화번호", sign.담당자전화);
  set("계약담당자", sign.담당자성명);
  set("행정실장", sign.관리자성명);
  set("교장", sign.대표자성명);
  set("대표자", sign.대표자성명);

  // 비상연락망 · 관할기관
  set("관할소방서", em.소방서);
  set("소방서", em.소방서);
  set("소방서전화번호", em.소방서전화);
  set("인근병원", em.병원);
  set("병원", em.병원);
  set("병원전화번호", em.병원전화);
  set("관할고용노동부", em.노동지청);
  set("관할고용노동지청", em.노동지청);
  set("고용노동부전화번호", em.노동지청전화);

  // 계약정보
  set("연번", serial);
  set("계약명", c.계약명);
  set("업체명", c.업체명);
  set("사업자등록번호", c.사업자번호);
  set("업체주소", c.업체주소);
  set("업체대표자", c.업체대표);
  set("대표자명", c.업체대표);
  set("업체대표자연락처", c.대표자전화);
  set("업체전화", c.대표자전화);
  set("안전담당자", c.안전담당자);
  set("업체담당자", c.안전담당자);
  set("담당자전화", c.담당자전화);
  set("안전담당자연락처", c.담당자전화);
  set("업체담당자연락처", c.담당자전화);
  set("점검요일", c.점검요일);

  const sd = toDay(c.시작일), ed = toDay(c.종료일);
  const days = contractDays(c);
  set("계약시작일", fmtDate(sd));
  set("계약종료일", fmtDate(ed));
  set("계약일수", days > 0 ? `${days}일간` : "");
  set("계약기간", days > 0 ? `${fmtDate(sd)} ~ ${fmtDate(ed)}(${days}일간)` : "");

  const amt = Number(String(c.계약금액 ?? "").replace(/[,\s원]/g, ""));
  if (trim(c.계약금액) !== "" && Number.isFinite(amt)) {
    set("계약금액", fmtMoney(amt) + "원");
    set("계약금액천원", fmtMoney(amt / 1000) + "천원");
  } else {
    set("계약금액", "");
    set("계약금액천원", "");
  }

  // 날짜 파생
  set("작성일", fmtDateDow(today));
  set("평가일", sd == null ? "" : fmtDateDow(sd));
  set("협의체회의일", sd == null ? "" : fmtDateDow(nextWorkday(sd)));
  set("점검일자", "");
  set("회차", "");
  set("점검년", "");
  set("점검월", "");
  for (let k = 1; k <= 5; k++) { set(`점검일${k}`, ""); set(`점검요일${k}`, ""); }

  // 계약유형 사전 문구
  const types = activeWorkTypes(c);
  set("작업유형목록", types.replace(/\|/g, ", "));
  const tname = trim(c.계약유형?.v);
  const F = (field) => typeField(tname, field);
  set("계약유형", tname);
  set("작업내용", F("작업내용"));
  set("협의체안건", F("협의안건"));
  set("점검착안사항", F("점검착안"));
  set("위험요인", mergeBlocks(F("위험요인"), contentBlock(types, "위험요인")));
  set("안전대책", mergeBlocks(F("안전대책"), contentBlock(types, "안전대책")));

  set("작업절차", orElse(F("작업절차"),
    "작업 전 현장 확인 → 작업구역 통제 → 작업 시행 → 정리정돈 → 작업 후 확인"));
  set("안전점검계획", orElse(F("점검계획"), "작업장 주변 통제, 유해·위험 기계기구 점검"));
  set("보호구방호장치", orElse(F("보호구"),
    "안전모, 안전화, 보호장갑 등 개인보호구 지급 및 착용 확인"));
  set("안전보건교육", orElse(F("교육계획"),
    "작업 전 기계·기구 사용 시 주의사항 및 안전수칙 준수 교육"));
  set("대피방법", orElse(F("대피방법"),
    "작업을 즉시 멈추고 가까운 계단으로 건물 밖 안전한 곳까지 이동해 인원을 확인, " +
    "다친 사람이 있으면 보건실을 거쳐 인근 병원으로 이송"));

  set("작업권장시기", orElse(F("권장시기"), "학사일정을 확인해 수업에 지장이 없는 시간대에 시행"));
  set("위험성등급", orElse(F("위험등급"), "중"));
  set("학교확인사항", orElse(F("학교확인"),
    "작업 구역과 학생 통행로 분리 확인" + LINE_SEP +
    "작업 일정과 범위를 관련 교직원에게 안내" + LINE_SEP +
    "작업 후 정리 상태를 확인한 뒤 인수"));
  set("업체조치사항", orElse(F("업체조치"),
    "작업 전 안전교육을 하고 보호구 착용 확인" + LINE_SEP +
    "작업 구역에 통제 표지를 세우고 학생 접근 차단" + LINE_SEP +
    "작업 후 공구와 폐자재를 당일 치우고 보고"));
  set("협의체합의문", orElse(F("합의문"),
    "작업 구역과 학생 통행로를 분리하고 통제 표지를 두기로 합의함" + LINE_SEP +
    "소음과 먼지가 나는 작업은 수업시간을 피해 시행하기로 협의함" + LINE_SEP +
    "작업 후 정리 상태를 학교 담당자와 함께 확인하기로 합의함"));
  set("상시안건", contentBlock(types, "상시안건"));

  // 슬롯 태그 ― 항목을 한 줄에 하나씩
  const g = (k) => d[`{{${k}}}`];
  for (let i = 1; i <= 5; i++) {
    set(`위험요인${i}`, nthItem(g("위험요인"), i));
    set(`안전대책${i}`, nthItem(g("안전대책"), i));
  }
  for (let i = 1; i <= 3; i++) {
    set(`작업내용${i}`, nthItem(g("작업내용"), i));
    set(`협의체안건${i}`, nthItem(g("협의체안건"), i));
    set(`점검착안사항${i}`, nthItem(g("점검착안사항"), i));
    set(`사용기계기구${i}`, nthItem(F("기계기구"), i));
    set(`위험물질설비${i}`, nthItem(F("위험물질"), i));
    set(`학교확인사항${i}`, nthItem(g("학교확인사항"), i));
    set(`업체조치사항${i}`, nthItem(g("업체조치사항"), i));
    set(`협의체합의문${i}`, nthItem(g("협의체합의문"), i));
    set(`상시안건${i}`, nthItem(g("상시안건"), i));
  }
  return d;
}

// 협의체 회차 한 번 몫. 무작위가 아니라 돌려쓰기 ― 다시 뽑아도 결과가 같다.
export function applyCouncilRound(tags, roundNo, day) {
  const g = (k) => tags[`{{${k}}}`] ?? "";
  const set = (k, v) => { tags[`{{${k}}}`] = s(v); };
  set("협의체회의일", fmtDateDow(day));
  set("회차", `${roundNo}회차`);
  const off = roundNo - 1;
  set("위험요인1", rotateItem(g("위험요인"), off + 1));
  set("위험요인2", rotateItem(g("위험요인"), off + 2));
  set("안전대책1", rotateItem(g("안전대책"), off + 1));
  const season = seasonAgenda(day);
  set("상시안건1", season || rotateItem(g("상시안건"), off + 1));
  set("협의체합의문1", rotateItem(g("협의체합의문"), off + 1));
  set("협의체합의문2", rotateItem(g("협의체합의문"), off + 2));
  const res = roundResolution(roundNo);
  if (res) set("협의체합의문3", res);
  return tags;
}
