// 나라장터(조달청 사용자정보 서비스) 찾기 중계 ― Vercel 함수 /api/g2b (ES 모듈이라 .mjs)
//   ?kind=corp&q=업체명 또는 사업자등록번호   → 조달업체 기본정보
//   ?kind=inst&q=기관명 또는 사업자등록번호   → 수요기관 정보(교육지원청 · 학교의 고유번호)
// 공공데이터포털 인증키(DATA_GO_KR_KEY)를 웹페이지에 드러내지 않으려고 둔다.
// 묻는 말(q)만 조달청에 넘기고 아무것도 저장 · 기록하지 않는다.
// 개발 중에는 web/vite.config.js 가 같은 함수를 /api/g2b 로 붙인다(키는 web/.env.local).

const BASE = "https://apis.data.go.kr/1230000/ao/UsrInfoService02/";
const OPS = { corp: "getPrcrmntCorpBasicInfo02", inst: "getDminsttInfo02" };
const NAME_PARAM = { corp: "corpNm", inst: "dminsttNm" };

function send(res, status, body, cache) {
  res.statusCode = status;
  res.setHeader("Content-Type", "application/json; charset=utf-8");
  res.setHeader("Cache-Control", cache || "no-store");
  res.end(JSON.stringify(body));
}

// 응답 칸 이름이 문서에 영문으로 나와 있지 않아, 이름 꼴로 고른다
const val = (o, re) => {
  const k = Object.keys(o).find((x) => re.test(x));
  return k == null || o[k] == null ? "" : String(o[k]).trim();
}
const fmtBizno = (s) => {
  const d = String(s || "").replace(/\D/g, "");
  return d.length === 10 ? `${d.slice(0, 3)}-${d.slice(3, 5)}-${d.slice(5)}` : String(s || "").trim();
}

function normalize(kind, r) {
  const base = val(r, /^(adrs|addr|adres)$/i) || val(r, /^(?!dtl).*(adrs|addr)/i);
  const dtl = val(r, /^dtl.*(adrs|addr)/i);
  return {
    name: kind === "corp" ? val(r, /^corpNm$/i) || val(r, /^(?!eng).*corpNm$/i) : val(r, /^dminsttNm$/i) || val(r, /dminsttNm/i),
    bizno: fmtBizno(val(r, /bizno/i)),
    ceo: val(r, /(ceo|rprsntv).*nm/i),
    addr: [base, dtl].filter(Boolean).join(" "),
    tel: val(r, /^telno$/i) || val(r, /tel/i),
    fax: val(r, /fax/i),
    region: val(r, /^rgnNm$/i),
    type: val(r, /(insttTy|instt.*Ty|bsns.*Div).*Nm/i),
  };
}

// data.go.kr 응답 → { code, msg, rows }
function unwrap(j) {
  const err = j?.OpenAPI_ServiceResponse?.cmmMsgHeader;
  if (err) return { code: err.returnReasonCode || "ERR", msg: err.returnAuthMsg || err.errMsg || "오류", rows: [] };
  const resp = j?.response || j;
  const code = String(resp?.header?.resultCode ?? "");
  const msg = resp?.header?.resultMsg || "";
  let items = resp?.body?.items;
  if (items && !Array.isArray(items)) items = items.item ?? [];
  if (items && !Array.isArray(items)) items = [items];
  return { code, msg, rows: items || [] };
}

function stamp(d) {
  const p = (n) => String(n).padStart(2, "0");
  return `${d.getFullYear()}${p(d.getMonth() + 1)}${p(d.getDate())}${p(d.getHours())}${p(d.getMinutes())}`;
}

async function call(op, params, key) {
  const p = new URLSearchParams({ serviceKey: key, pageNo: "1", numOfRows: "10", type: "json", ...params });
  const r = await fetch(BASE + op + "?" + p, { signal: AbortSignal.timeout(9000) });
  const text = await r.text();
  try { return unwrap(JSON.parse(text)); }
  catch { return { code: "HTTP" + r.status, msg: (text.match(/<returnAuthMsg>([^<]*)/) || text.match(/<resultMsg>([^<]*)/) || [])[1] || "응답을 읽지 못함", rows: [] }; }
}

async function lookup(kind, q, key) {
  const digits = q.replace(/\D/g, "");
  const byNo = digits.length === 10 && /^[\d\s-]+$/.test(q);
  const what = byNo ? { bizno: digits } : { [NAME_PARAM[kind]]: q };
  // 조회구분(inqryDiv)의 뜻이 문서에 없다 ― 되는 조합을 차례로 묻는다
  const tries = [
    { inqryDiv: "1" },
    { inqryDiv: "1", inqryBgnDt: "190001010000", inqryEndDt: stamp(new Date()) },
    { inqryDiv: "2" },
    { inqryDiv: "3" },
  ];
  let last = { code: "", msg: "", rows: [] };
  const log = [];
  for (const t of tries) {
    last = await call(OPS[kind], { ...t, ...what }, key);
    log.push({ ...t, code: last.code, msg: last.msg, n: last.rows.length });
    if (last.rows.length) break;
    if (/^(20|22|30|31|32|33)$/.test(last.code)) break; // 키 · 권한 · 한도 문제는 다시 물어도 같다
  }
  return { last, log, byNo };
}

export default async function handler(req, res) {
  if (req.method !== "GET") return send(res, 405, { error: "GET 만 됩니다." });
  // 다른 사이트에서 이 중계를 끌어다 쓰지 못하게
  if (req.headers["sec-fetch-site"] === "cross-site") return send(res, 403, { error: "이 사이트에서만 쓸 수 있습니다." });
  const key0 = (process.env.DATA_GO_KR_KEY || "").trim();
  if (!key0) return send(res, 503, { error: "서버에 공공데이터포털 인증키가 없습니다." });
  let key = key0;
  if (/%[0-9A-F]{2}/i.test(key)) { try { key = decodeURIComponent(key); } catch {} } // Encoding 키를 넣었어도

  const url = new URL(req.url, "http://local");
  const kind = url.searchParams.get("kind");
  const q = (url.searchParams.get("q") || "").trim().slice(0, 60);
  const debug = url.searchParams.get("debug") === "1";
  if (!OPS[kind]) return send(res, 400, { error: "kind 는 corp 또는 inst 입니다." });
  if (q.replace(/\s/g, "").length < 2) return send(res, 400, { error: "두 글자 이상 넣어 주세요." });

  try {
    const { last, log } = await lookup(kind, q, key);
    const ok = last.rows.length > 0 || last.code === "00" || last.code === "03";
    if (!ok && last.code) {
      const why = /^(30|31)$/.test(last.code) ? "인증키가 아직 등록되지 않았습니다(발급 뒤 1~2시간 걸릴 수 있습니다)."
        : last.code === "22" ? "오늘 조회 한도를 다 썼습니다. 내일 다시 해 주세요."
        : last.code === "20" ? "이 인증키에 해당 기능 사용 권한이 없습니다(활용신청 상세기능 확인)."
        : `조달청 응답: ${last.msg || last.code}`;
      return send(res, 502, { error: why, ...(debug ? { log } : {}) });
    }
    const items = last.rows.map((r) => normalize(kind, r)).filter((x) => x.name || x.bizno);
    const body = { items, ...(debug ? { log, raw: last.rows.slice(0, 2) } : {}) };
    return send(res, 200, body, debug ? "no-store" : "public, s-maxage=86400, max-age=600");
  } catch (e) {
    return send(res, 504, { error: "조달청에 연결하지 못했습니다. 잠시 뒤 다시 해 주세요." });
  }
}
