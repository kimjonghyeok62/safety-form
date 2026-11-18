// 출력 버튼 하나 = 함수 하나. 원본: modHwpExport.bas 의 Btn_Export* 흐름.
// 결과는 { files: [{name, bytes}], notes: [문자열] } ― 파일 내려받기는 화면 쪽이 한다.

import { buildTagMap, applyCouncilRound } from "./domain/tagMap.js";
import {
  templateFor, formNeeded, crit, FORMS,
} from "./domain/dictionary.js";
import {
  contractDays, patrolMonths, councilDates, inspectionDates,
  fmtDateDow, toDay, parts,
} from "./domain/schedule.js";
import { renderHwpx, safeName, labelOf } from "./hwpx/render.js";

const pad = (n) => String(n).padStart(2, "0");
const PLEDGE = 15; // 계약이행 통합 서약서 ― 일괄 출력에 늘 넣는다

function stamp(c) {
  const d = toDay(c.시작일);
  if (d == null) return "";
  const p = parts(d);
  return `_${p.y}${pad(p.m)}${pad(p.d)}`;
}

function outName(c, label) {
  return safeName(`${String(c.계약명 || "").trim()}_${label}${stamp(c)}`) + ".hwpx";
}

// 출력관리의 ○ ― 사용자가 바꾼 것이 있으면 그것, 없으면 자동 판정
export function formMarked(c, formNo) {
  const o = c.formOverride?.[formNo];
  return typeof o === "boolean" ? o : formNeeded(formNo, c);
}

export class Exporter {
  /** @param {(name: string) => Promise<Uint8Array|ArrayBuffer>} loadTemplate */
  constructor(loadTemplate, today) {
    this.loadTemplate = loadTemplate;
    this.today = today;
    this.cache = new Map();
  }

  async tmpl(name) {
    if (!this.cache.has(name)) this.cache.set(name, await this.loadTemplate(name));
    return this.cache.get(name);
  }

  tags(data, c) {
    const serial = data.contracts.indexOf(c) + 1;
    return buildTagMap(data, c, serial, this.today);
  }

  async render(name, tagMaps, fileName, out) {
    const r = await renderHwpx(await this.tmpl(name), tagMaps);
    out.files.push({ name: fileName, bytes: r.blob });
    if (r.missing.length) out.notes.push(`${fileName}: 채우지 못한 칸 ${r.missing.join(" ")}`);
  }

  // [1~7번 출력] ― ○ 표시된 서식을 서식마다 한 파일로
  async forms17(data, contracts) {
    const out = { files: [], notes: [] };
    for (const c of contracts) {
      const tags = this.tags(data, c);
      let made = 0;
      for (let no = 1; no <= 7; no++) {
        if (!formMarked(c, no)) continue;
        const t = templateFor(no, c);
        await this.render(t, [tags], outName(c, `${pad(no)}_${labelOf(t)}`), out);
        made++;
      }
      if (!made) out.notes.push(`${c.계약명} ― ○ 표시된 서식이 없습니다`);
    }
    return out;
  }

  // [8. 협의체 회의록] ― 월 1회, 회차마다 내용이 조금씩 다르다
  async council(data, contracts) {
    const out = { files: [], notes: [] };
    for (const c of contracts) {
      const days = contractDays(c);
      if (days < crit("기준_협의체")) {
        out.notes.push(`${c.계약명} ― 계약기간 ${days}일 (${crit("기준_협의체")}일 이상일 때 필요)`);
        continue;
      }
      const dates = councilDates(c);
      if (!dates.length) { out.notes.push(`${c.계약명} ― 계약기간 안에 ${c.점검요일}이 없습니다`); continue; }
      const maps = dates.map((d, i) => applyCouncilRound(this.tags(data, c), i + 1, d));
      await this.render(FORMS[8].file, maps,
        outName(c, `안전보건협의체 회의록_총${maps.length}회차`), out);
    }
    return out;
  }

  // [9. 순회점검표] ― 달마다 1장, 1주~5주 칸에 그 달 점검일
  async patrol(data, contracts) {
    const out = { files: [], notes: [] };
    for (const c of contracts) {
      const days = contractDays(c);
      if (days < crit("기준_순회")) {
        out.notes.push(`${c.계약명} ― 계약기간 ${days}일 (${crit("기준_순회")}일 이상일 때 필요)`);
        continue;
      }
      const months = patrolMonths(c);
      if (!months.length) { out.notes.push(`${c.계약명} ― 계약기간 안에 ${c.점검요일}이 없습니다`); continue; }
      const maps = months.map((m, i) => {
        const t = this.tags(data, c);
        t["{{점검년}}"] = m.y;
        t["{{점검월}}"] = m.m;
        for (let k = 1; k <= 5; k++) {
          t[`{{점검일${k}}}`] = m.dates[k - 1];
          t[`{{점검요일${k}}}`] = m.dows[k - 1];
        }
        t["{{점검일자}}"] = m.dates[0] + m.dows[0];
        t["{{회차}}"] = `${i + 1}회차`;
        return t;
      });
      await this.render(FORMS[9].file, maps, outName(c, `순회점검표_총${maps.length}회차`), out);
    }
    return out;
  }

  // [10. 합동점검표] ― 90일 구간마다 마지막 점검일
  async joint(data, contracts) {
    const out = { files: [], notes: [] };
    for (const c of contracts) {
      const days = contractDays(c);
      if (days < crit("기준_합동")) {
        out.notes.push(`${c.계약명} ― 계약기간 ${days}일 (${crit("기준_합동")}일 이상일 때 필요)`);
        continue;
      }
      const dates = inspectionDates(c, true);
      if (!dates.length) { out.notes.push(`${c.계약명} ― 계약기간 안에 해당 요일이 없습니다`); continue; }
      const maps = dates.map((d, i) => {
        const t = this.tags(data, c);
        t["{{점검일자}}"] = fmtDateDow(d);
        t["{{회차}}"] = `${i + 1}회차`;
        return t;
      });
      await this.render(FORMS[10].file, maps,
        outName(c, `${labelOf(FORMS[10].file)}_총${maps.length}회차`), out);
    }
    return out;
  }

  // [일괄 출력] ― 계약마다 ○ 표시된 서식 전부(1~10) + 15. 계약이행 통합 서약서(계약마다 늘 1부).
  // 파일마다 dir(계약명)을 달아 zip 안에서 계약별 폴더로 나눈다
  async all(data, contracts) {
    const out = { files: [], notes: [] };
    const NONE = "○ 표시된 서식이 없습니다";
    for (const c of contracts) {
      const parts = [await this.forms17(data, [c])];
      if (formMarked(c, 8)) parts.push(await this.council(data, [c]));
      if (formMarked(c, 9)) parts.push(await this.patrol(data, [c]));
      if (formMarked(c, 10)) parts.push(await this.joint(data, [c]));
      let made = 0;
      for (const p of parts) {
        p.files.forEach((f) => out.files.push({ ...f, dir: String(c.계약명 || "").trim() }));
        made += p.files.length;
        out.notes.push(...p.notes.filter((n) => !n.endsWith(NONE)));
      }
      if (!made) out.notes.push(`${c.계약명} ― ${NONE}(계약이행 통합 서약서만 넣었습니다)`);
      const pledge = await this.single(data, [c], PLEDGE);
      pledge.files.forEach((f) => out.files.push({ ...f, name: outName(c, `${PLEDGE}_${labelOf(FORMS[PLEDGE].file)}`), dir: String(c.계약명 || "").trim() }));
      out.notes.push(...pledge.notes);
    }
    return out;
  }

  // [11~15] ― 계약마다 1부
  async single(data, contracts, formNo) {
    const out = { files: [], notes: [] };
    const t = FORMS[formNo].file;
    for (const c of contracts) {
      await this.render(t, [this.tags(data, c)], outName(c, labelOf(t)), out);
    }
    return out;
  }
}
