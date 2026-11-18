// 계약이 지금 어디쯤인지 ― 예정 · 진행 · 완료와 계약기간 경과율(%).
// 체크할 것 없이 날짜만으로 정한다. 비고를 직접 고른 계약(진행/완료)은 그 값을 따른다.

import { toDay, contractDays } from "./schedule.js";

/** @returns {{ state: "예정"|"진행"|"완료"|"", pct: number|null, left: number|null }} left = 남은 날 */
export function statusOf(c, today) {
  const s = toDay(c.시작일), e = toDay(c.종료일);
  const days = contractDays(c);
  const manual = c.비고?.manual ? c.비고.v : "";
  if (manual === "완료") return { state: "완료", pct: 100, left: null };
  if (!String(c.계약명 || "").trim() || !days) return { state: manual || "", pct: null, left: null };
  if (!manual && today < s) return { state: "예정", pct: 0, left: s - today };
  if (!manual && today > e) return { state: "완료", pct: 100, left: null };
  const pct = Math.max(0, Math.min(100, Math.round(((today - s + 1) / days) * 100)));
  return { state: "진행", pct, left: Math.max(0, e - today) };
}
