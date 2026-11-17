// hwpx 조립 ― 한글 프로그램 없이 서식 파일(zip + XML)을 직접 고친다.
//
// 엑셀판은 한글 COM 의 '모두 바꾸기'로 {{태그}} 를 채우고 InsertFile 로 이어붙였다.
// 브라우저에서는 그럴 수 없으므로 Contents/section*.xml 의 글자를 직접 바꾼다.
//
// 실측으로 확인한 함정(메모: hwpx-edit-gotchas)을 지킨다.
//  · 태그가 run 두 개로 쪼개져 있는 서식이 있다('{{학교' + '명}}'). 문단 단위로
//    글자를 이어 붙여 찾고, 걸친 조각은 앞 조각에 몰아 넣는다.
//  · 문단이 없는 표 칸은 한글을 죽인다 → 문단·<hp:t> 는 지우지 않고 글자만 바꾼다.
//  · 글자를 바꾼 문단의 <hp:linesegarray> 는 지운다. 옛 줄 배치가 남으면 글자가 겹친다.
//
// 같은 서식을 여러 장(순회 12장 등) 낼 때는 본문을 구역(section)으로 복제한다.
// 구역은 늘 새 쪽에서 시작하므로 쪽 나눔을 따로 넣지 않는다. 글자·문단 모양 표
// (header.xml)는 같은 서식이라 그대로 쓸 수 있다.

import JSZip from "jszip";

const TAG_RE = /\{\{[^{}]*?\}\}/g;

const esc = (v) =>
  String(v).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");

/**
 * section XML 한 장의 태그를 채운다.
 * @returns {{xml: string, missing: string[]}} missing = 값이 없어 남긴 태그
 */
export function fillSection(xml, tags) {
  const tokens = xml.match(/<[^>]*>|[^<]+/g) || [];
  const stack = [];
  const paras = [];
  const segs = []; // { i: token index, p: paragraph }
  let inT = 0;

  for (let i = 0; i < tokens.length; i++) {
    const t = tokens[i];
    if (t[0] === "<") {
      if (/^<hp:p[\s>]/.test(t)) {
        const p = { open: i, lsStart: -1, lsEnd: -1, dirty: false };
        paras.push(p);
        if (!t.endsWith("/>")) stack.push(p);
      } else if (t === "</hp:p>") {
        stack.pop();
      } else if (/^<hp:t[\s>]/.test(t)) {
        if (!t.endsWith("/>")) inT++;
      } else if (t === "</hp:t>") {
        inT--;
      } else if (/^<hp:linesegarray[\s>]/.test(t)) {
        const p = stack[stack.length - 1];
        if (p) p.lsStart = i;
        if (p && t.endsWith("/>")) p.lsEnd = i;
      } else if (t === "</hp:linesegarray>") {
        const p = stack[stack.length - 1];
        if (p) p.lsEnd = i;
      }
    } else if (inT > 0) {
      segs.push({ i, p: stack[stack.length - 1] || null });
    }
  }

  // 글자 조각을 이어 붙이고 각 조각의 시작 위치를 적어 둔다
  const starts = [];
  let all = "";
  for (const s of segs) {
    starts.push(all.length);
    all += tokens[s.i];
  }
  const segAt = (pos) => {
    // pos 를 품은 조각 번호 (이진 탐색)
    let lo = 0, hi = starts.length - 1;
    while (lo < hi) {
      const mid = (lo + hi + 1) >> 1;
      if (starts[mid] <= pos) lo = mid; else hi = mid - 1;
    }
    return lo;
  };

  const matches = [...all.matchAll(TAG_RE)];
  const missing = new Set();
  // 뒤에서부터 바꿔야 앞쪽 위치가 흐트러지지 않는다
  for (let k = matches.length - 1; k >= 0; k--) {
    const m = matches[k];
    const key = m[0];
    if (!(key in tags)) { missing.add(key); continue; }
    const a = segAt(m.index);
    const b = segAt(m.index + key.length - 1);
    // 다른 문단에 걸친 것은 태그가 아니다
    let samePara = true;
    for (let j = a; j <= b; j++) if (segs[j].p !== segs[a].p) samePara = false;
    if (!samePara) continue;

    const offA = m.index - starts[a];
    const endB = m.index + key.length - starts[b];
    const val = esc(tags[key]);
    if (a === b) {
      const tx = tokens[segs[a].i];
      tokens[segs[a].i] = tx.slice(0, offA) + val + tx.slice(endB);
    } else {
      tokens[segs[a].i] = tokens[segs[a].i].slice(0, offA) + val;
      for (let j = a + 1; j < b; j++) tokens[segs[j].i] = "";
      tokens[segs[b].i] = tokens[segs[b].i].slice(endB);
    }
    for (let j = a; j <= b; j++) if (segs[j].p) segs[j].p.dirty = true;
  }

  for (const p of paras) {
    if (p.dirty && p.lsStart >= 0 && p.lsEnd >= p.lsStart) {
      for (let i = p.lsStart; i <= p.lsEnd; i++) tokens[i] = "";
    }
  }
  return { xml: tokens.join(""), missing: [...missing] };
}

/**
 * 서식 한 개 + 태그 사전 여러 벌 → hwpx 한 파일(사전 한 벌 = 구역 하나 = 새 쪽부터).
 * @param {ArrayBuffer|Uint8Array} templateBytes
 * @param {object[]} tagMaps
 * @returns {Promise<{blob: Uint8Array, missing: string[]}>}
 */
export async function renderHwpx(templateBytes, tagMaps) {
  if (!tagMaps.length) throw new Error("채울 내용이 없습니다");
  const zip = await JSZip.loadAsync(templateBytes);

  const secNames = Object.keys(zip.files)
    .filter((n) => /^Contents\/section\d+\.xml$/.test(n))
    .sort((x, y) => +x.match(/\d+/)[0] - +y.match(/\d+/)[0]);
  if (!secNames.length) throw new Error("서식 파일 안에 본문(section)이 없습니다");

  const srcSecs = await Promise.all(secNames.map((n) => zip.file(n).async("string")));
  const missing = new Set();
  const outSecs = [];
  for (const tags of tagMaps) {
    for (const xml of srcSecs) {
      const r = fillSection(xml, tags);
      r.missing.forEach((m) => missing.add(m));
      outSecs.push(r.xml);
    }
  }

  // 구역 파일 쓰기
  outSecs.forEach((xml, i) => zip.file(`Contents/section${i}.xml`, xml));

  if (outSecs.length !== srcSecs.length) {
    // content.hpf ― 목록과 읽는 차례에 새 구역을 올린다
    let hpf = await zip.file("Contents/content.hpf").async("string");
    hpf = hpf.replace(/<opf:item id="section\d+"[^>]*\/>/g, "");
    hpf = hpf.replace(/<opf:itemref idref="section\d+"[^>]*\/>/g, "");
    const items = outSecs.map((_, i) =>
      `<opf:item id="section${i}" href="Contents/section${i}.xml" media-type="application/xml"/>`).join("");
    const refs = outSecs.map((_, i) => `<opf:itemref idref="section${i}" linear="yes"/>`).join("");
    hpf = hpf.replace(/(<opf:item id="header"[^>]*\/>)/, `$1${items}`);
    hpf = hpf.replace(/(<opf:itemref idref="header"[^>]*\/>)/, `$1${refs}`);
    zip.file("Contents/content.hpf", hpf);

    let header = await zip.file("Contents/header.xml").async("string");
    header = header.replace(/(<hh:head\b[^>]*\bsecCnt=")\d+(")/, `$1${outSecs.length}$2`);
    zip.file("Contents/header.xml", header);
  }

  // 미리보기 글자는 서식의 {{태그}} 가 그대로라 지운다(파일 목록에서 엉뚱한 글이 보인다)
  if (zip.file("Preview/PrvText.txt")) zip.file("Preview/PrvText.txt", "");

  // mimetype 은 맨 앞, 압축 없이 ― hwpx(OCF) 규칙
  const out = new JSZip();
  const mt = await zip.file("mimetype").async("uint8array");
  out.file("mimetype", mt, { compression: "STORE" });
  for (const name of Object.keys(zip.files)) {
    if (name === "mimetype" || zip.files[name].dir) continue;
    out.file(name, await zip.file(name).async("uint8array"), { compression: "DEFLATE" });
  }
  const blob = await out.generateAsync({ type: "uint8array", compression: "DEFLATE" });
  return { blob, missing: [...missing] };
}

// 파일 이름에 못 쓰는 글자를 뗀다(가운뎃점 포함 ― 엑셀판 SafeName 과 같다)
export function safeName(s) {
  return String(s).replace(/[\\/:*?"<>|·]/g, "");
}

// "09_순회점검표.hwpx" → "순회점검표"
export function labelOf(tmpl) {
  let s = tmpl.replace(/\.[^.]*$/, "");
  const i = s.indexOf("_");
  return i >= 0 ? s.slice(i + 1) : s;
}
