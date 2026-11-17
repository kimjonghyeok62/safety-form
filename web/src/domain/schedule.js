// 날짜 계산 ― 점검일, 공휴일 미루기, 달/분기 나누기, 날짜 표기.
// 원본: modHwpExport.bas (InspectionDates · PatrolMonths · CouncilDates · Fmt*)
//       modConfig.bas (HolidayList · NextWorkday · WeekdayNumber)
//
// 날짜는 'YYYY-MM-DD' 문자열로 받고, 안에서는 1970-01-01 부터 센 '일 번호' 로 다룬다.
// Date 객체를 그대로 쓰면 브라우저의 시간대에 따라 하루씩 밀릴 수 있다.

const DAY_MS = 86400000;

export function toDay(s) {
  if (!s) return null;
  const m = /^(\d{4})-(\d{1,2})-(\d{1,2})$/.exec(String(s).trim());
  if (!m) return null;
  const y = +m[1], mo = +m[2], d = +m[3];
  const t = Date.UTC(y, mo - 1, d);
  const back = new Date(t);
  if (back.getUTCMonth() !== mo - 1 || back.getUTCDate() !== d) return null;
  return t / DAY_MS;
}

export function parts(day) {
  const dt = new Date(day * DAY_MS);
  return { y: dt.getUTCFullYear(), m: dt.getUTCMonth() + 1, d: dt.getUTCDate(), wd: dt.getUTCDay() };
}

export function dayToIso(day) {
  const p = parts(day);
  return `${p.y}-${pad(p.m)}-${pad(p.d)}`;
}

export function todayDay() {
  const n = new Date();
  return Date.UTC(n.getFullYear(), n.getMonth(), n.getDate()) / DAY_MS;
}

const pad = (n) => String(n).padStart(2, "0");

// 0=일 … 6=토  (VBA Weekday 는 1=일 … 7=토)
export const WEEKDAYS = ["일요일", "월요일", "화요일", "수요일", "목요일", "금요일", "토요일"];

// "월요일" → 1. 모르면 수요일(3) ― modConfig.WeekdayNumber 와 같다.
export function weekdayIndex(name) {
  const i = WEEKDAYS.indexOf(String(name || "").trim());
  return i < 0 ? 3 : i;
}

// 관공서가 쉬는 날. 음력 명절이 있어 계산으로 낼 수 없어 적어 둔다(2030년까지).
// 그 뒤 해는 주말만 피한다 ― 엑셀판과 같다.
const HOLIDAYS = new Set(
  (
    "20260101 20260216 20260217 20260218 20260301 20260302 20260501 20260505 20260524 20260525 20260603 20260606 20260717 20260815 20260817 20260924 20260925 20260926 20261003 20261005 20261009 20261225 " +
    "20270101 20270206 20270207 20270208 20270209 20270301 20270501 20270503 20270505 20270513 20270606 20270717 20270719 20270815 20270816 20270914 20270915 20270916 20271003 20271004 20271009 20271011 20271225 20271227 " +
    "20280101 20280126 20280127 20280128 20280301 20280412 20280501 20280502 20280505 20280606 20280717 20280815 20281002 20281003 20281004 20281005 20281009 20281225 " +
    "20290101 20290212 20290213 20290214 20290301 20290501 20290505 20290507 20290520 20290521 20290606 20290717 20290815 20290921 20290922 20290923 20290924 20291003 20291009 20291225 " +
    "20300101 20300202 20300203 20300204 20300205 20300301 20300403 20300501 20300505 20300506 20300509 20300606 20300612 20300717 20300815 20300911 20300912 20300913 20301003 20301009 20301225"
  ).split(/\s+/).filter(Boolean)
);

export function isHoliday(day) {
  const p = parts(day);
  return HOLIDAYS.has(`${p.y}${pad(p.m)}${pad(p.d)}`);
}

// 주말이거나 공휴일이면 다음 평일로. 연휴가 길어도 계속 민다(30일에서 멈춤).
export function nextWorkday(day) {
  let i = 0;
  while (i <= 30) {
    const wd = parts(day).wd;
    if (wd !== 0 && wd !== 6 && !isHoliday(day)) break;
    day += 1;
    i += 1;
  }
  return day;
}

// 계약기간(일수). 비었거나 뒤집혀 있으면 0.
export function contractDays(c) {
  const s = toDay(c.시작일), e = toDay(c.종료일);
  if (s == null || e == null || e < s) return 0;
  return e - s + 1;
}

// 계약기간 안의 점검일. quarterly = true 이면 시작일부터 90일 구간마다 마지막 하나만.
// 구간 수는 반올림 ― 365일 계약이 5구간(꼬리 5일)으로 쪼개지지 않게 한다.
export function inspectionDates(c, quarterly) {
  const s = toDay(c.시작일), e = toDay(c.종료일);
  if (s == null || e == null) return [];
  const wd = weekdayIndex(c.점검요일);

  const total = e - s + 1;
  let spanCount = Math.floor(total / 90 + 0.5);
  if (spanCount < 1) spanCount = 1;

  const acc = [];
  let lastPeriod = -1;
  for (let d = s; d <= e; d++) {
    if (parts(d).wd !== wd) continue;
    let period = Math.floor((d - s) / 90);
    if (period > spanCount - 1) period = spanCount - 1;

    const d2 = nextWorkday(d);
    if (d2 > e) continue;
    if (quarterly && period === lastPeriod && acc.length > 0) {
      acc[acc.length - 1] = d2;
    } else if (acc.length < 501) {
      if (acc.length === 0 || acc[acc.length - 1] !== d2) acc.push(d2);
    }
    lastPeriod = period;
  }
  return acc;
}

// 달마다 한 묶음: { y, m, dates[5], dows[5] }  (없는 주차는 "")
export function patrolMonths(c) {
  const out = [];
  let cur = null, curYm = "";
  for (const d of inspectionDates(c, false)) {
    const p = parts(d);
    const ym = `${p.y}${pad(p.m)}`;
    if (ym !== curYm) {
      cur = { y: String(p.y), m: String(p.m), dates: ["", "", "", "", ""], dows: ["", "", "", "", ""], n: 0 };
      out.push(cur);
      curYm = ym;
    }
    if (cur.n < 5) {
      cur.dates[cur.n] = fmtInspectDate(d);
      cur.dows[cur.n] = fmtDow(d);
      cur.n += 1;
    }
  }
  return out;
}

// 달마다 마지막 순회점검일 = 협의체 회의일
export function councilDates(c) {
  const acc = [];
  let curYm = "";
  for (const d of inspectionDates(c, false)) {
    const p = parts(d);
    const ym = `${p.y}${pad(p.m)}`;
    if (ym === curYm && acc.length > 0) acc[acc.length - 1] = d;
    else { acc.push(d); curYm = ym; }
  }
  return acc;
}

// 계약기간용 ― 2026. 03. 01.  (요일 없음)
export function fmtDate(day) {
  if (day == null) return "";
  const p = parts(day);
  return `${p.y}. ${pad(p.m)}. ${pad(p.d)}.`;
}

// 2026.6.4.  (0 을 채우지 않는다 ― 순회점검표 칸이 좁다)
export function fmtInspectDate(day) {
  const p = parts(day);
  return `${p.y}.${p.m}.${p.d}.`;
}

// (목)
export function fmtDow(day) {
  return `(${WEEKDAYS[parts(day).wd][0]})`;
}

// 2026.6.4.(목)
export function fmtDateDow(day) {
  return day == null ? "" : fmtInspectDate(day) + fmtDow(day);
}
