// 나라장터(조달청 사용자정보 서비스) 찾기 중계 ― Vercel 함수 /api/g2b (ES 모듈이라 .mjs)
//   ?kind=corp&q=사업자등록번호   → 조달업체 기본정보(업체명 · 대표자 · 주소 · 전화 · 팩스)
//   ?kind=inst&q=사업자등록번호   → 수요기관 정보(학교 · 교육지원청의 고유번호로 기관명 · 주소 · 전화 · 팩스 · 기관코드)
//   ?kind=cntrct&biz=servc|cnstwk|thng&code=기관코드(또는 name=기관명)&from=YYYYMMDD&to=YYYYMMDD
//                                 → 그 기관이 나라장터에 등록한 계약(계약명 · 기간 · 금액 · 업체). 조달청_나라장터 계약정보서비스
// 공공데이터포털 인증키(DATA_GO_KR_KEY)를 웹페이지에 드러내지 않으려고 둔다.
// 묻는 번호만 조달청에 넘기고 아무것도 저장 · 기록하지 않는다.
// 개발 중에는 web/vite.config.js 가 같은 함수를 /api/g2b 로 붙인다(키는 web/.env.local).
//
// 2026-11-18 실제 응답으로 확인한 것:
//   inqryDiv=3 + bizno → 사업자등록번호로 한 건. 이름만으로는 못 찾는다(08 필수값 입력 에러).
//   inqryDiv=1 · 2 는 등록 · 변경 일시 구간이 필수이고 구간이 길면 07 입력범위값 초과 에러.

const BASE = "https://apis.data.go.kr/1230000/ao/UsrInfoService02/";
const OPS = { corp: "getPrcrmntCorpBasicInfo02", inst: "getDminsttInfo02" };
// 나라장터검색조건에 의한 계약현황 ― 용역 · 공사 · 물품
const CBASE = "https://apis.data.go.kr/1230000/ao/CntrctInfoService/";
const COPS = { servc: "getCntrctInfoListServcPPSSrch", cnstwk: "getCntrctInfoListCnstwkPPSSrch", thng: "getCntrctInfoListThngPPSSrch" };
const CHUNK = 7;     // 2026-11-10 부터 한 번에 7일까지만 받는다(공공데이터포털 공지)
const MAX_SPAN = 35; // 함수 한 번에 묻는 기간 ― 7일 조각 5개를 함께

function send(res, status, body, cache) {
  res.statusCode = status;
  res.setHeader("Content-Type", "application/json; charset=utf-8");
  res.setHeader("Cache-Control", cache || "no-store");
  res.end(JSON.stringify(body));
}

const str = (v) => (v == null ? "" : String(v).trim());
const fmtBizno = (s) => {
  const d = str(s).replace(/\D/g, "");
  return d.length === 10 ? `${d.slice(0, 3)}-${d.slice(3, 5)}-${d.slice(5)}` : str(s);
};

function normalize(kind, r) {
  return {
    name: str(kind === "corp" ? r.corpNm : r.dminsttNm),
    short: str(r.dminsttAbrvtNm), // 수요기관 약칭(학교 · 교육지원청 이름만)
    code: str(r.dminsttCd), // 수요기관코드 ― 계약 가져오기에 쓴다
    bizno: fmtBizno(r.bizno),
    ceo: str(r.ceoNm),
    addr: [str(r.adrs), str(r.dtlAdrs)].filter(Boolean).join(" "),
    tel: str(r.telNo),
    fax: str(r.faxNo || r.ofclFaxNo),
  };
}

// data.go.kr 응답 → { code, msg, rows }. 오류는 OpenAPI_ServiceResponse(게이트웨이) 또는
// nkoneps.com.response.ResponseError(조달청) 꼴로 온다.
function unwrap(j) {
  const gw = j?.OpenAPI_ServiceResponse?.cmmMsgHeader;
  if (gw) return { code: str(gw.returnReasonCode) || "ERR", msg: gw.returnAuthMsg || gw.errMsg || "오류", rows: [] };
  const resp = j?.response || Object.values(j || {})[0];
  const code = str(resp?.header?.resultCode);
  const msg = str(resp?.header?.resultMsg);
  let items = resp?.body?.items;
  if (items && !Array.isArray(items)) items = items.item ?? [];
  if (items && !Array.isArray(items)) items = [items];
  return { code, msg, rows: items || [] };
}

// ---------------------------------------------------------------- 계약 (kind=cntrct)
// 2026-11-19 실제 응답으로 맞춘 칸 이름(용역 · 공사). 계약기관은 insttDivCd=1 + insttCd(=수요기관코드 꼴, 경기도교육청 7530000).
//   용역: cntrctNm · cntrctPrd("2026/01/01~2026/12/31" 또는 끝날만 "2026/12/31" · "20170831") · wbgnDate(착수) · thtmScmpltDate/ttalScmpltDate(완수)
//   공사: cnstwkNm · cntrctPrd("금차 125일, 총공사 125일") · cbgnDate(착공) · thtmCcmpltDate/ttalCcmpltDate(준공)
//   공통: untyCntrctNo · cntrctCnclsDate · totCntrctAmt(장기계속이면 0) · thtmCntrctAmt · cntrctCnclsMthdNm
//   corpList: "[1^주계약업체^단독^업체명^대표자^대한민국^지분^업체명^^사업자번호],[2^…]"
const pick = (r, ...keys) => { for (const k of keys) { const v = str(r[k]); if (v) return v; } return ""; };
const DATE_RE = /(20\d{2})[-./년\s]*(\d{1,2})[-./월\s]*(\d{1,2})/;
const isoOf = (s) => {
  const m = str(s).match(DATE_RE);
  return m ? `${m[1]}-${m[2].padStart(2, "0")}-${m[3].padStart(2, "0")}` : "";
};
const addDays = (iso, n) => {
  const d = new Date(iso + "T00:00:00Z");
  if (isNaN(d)) return "";
  d.setUTCDate(d.getUTCDate() + n);
  return d.toISOString().slice(0, 10);
};

// 첫 업체(주계약업체). 칸 순서가 다르게 와도 사업자번호는 10자리 숫자로 찾는다
function firstCorp(list) {
  const first = str(list).replace(/^\[|\]$/g, "").split(/\],\s*\[/)[0] || "";
  const parts = first.split("^").map(str);
  const bizno = parts.find((p) => /^\d{3}-?\d{2}-?\d{5}$/.test(p)) || "";
  return { name: parts[3] || parts[7] || "", ceo: parts[4] || "", bizno: bizno ? fmtBizno(bizno) : "" };
}

function normalizeCntrct(biz, r) {
  const period = pick(r, "cntrctPrd");
  const dates = (period.match(new RegExp(DATE_RE.source, "g")) || []).map(isoOf);
  const signed = isoOf(pick(r, "cntrctCnclsDate", "cntrctDate"));
  const begin = isoOf(pick(r, "wbgnDate", "cbgnDate"));
  const finish = isoOf(pick(r, "ttalScmpltDate", "thtmScmpltDate", "ttalCcmpltDate", "thtmCcmpltDate"));
  let start = "", end = "";
  if (dates.length >= 2) [start, end] = dates;
  else {
    end = dates[0] || finish;               // 날짜가 하나면 그것이 끝날
    start = begin || signed;                // 착수 · 착공일, 없으면 계약일
    if (!end && start) {                    // "총공사 125일" 꼴
      const n = +(period.match(/총[^\d]*(\d+)\s*일/) || period.match(/(\d+)\s*일/) || [])[1];
      if (n > 0) end = addDays(start, n - 1);
    }
  }
  if (start && end && end < start) start = signed && signed <= end ? signed : "";
  const amt = (k) => +(str(r[k]).replace(/\D/g, "")) || 0;
  const corp = firstCorp(r.corpList);
  return {
    // 같은 확정계약번호가 다시 등록되면 통합계약번호만 새로 붙는다 ― 확정계약번호로 하나로 본다
    no: pick(r, "dcsnCntrctNo", "cntrctRefNo", "untyCntrctNo"),
    biz,
    name: pick(r, "cntrctNm", "cnstwkNm", "prdctClsfcNoNm"),
    date: signed,
    start, end,
    period,
    amount: String(amt("totCntrctAmt") || amt("thtmCntrctAmt") || ""),
    corp: corp.name, ceo: corp.ceo, bizno: corp.bizno,
    method: pick(r, "cntrctCnclsMthdNm"),
  };
}

const ymd = (d) => d.toISOString().slice(0, 10).replace(/-/g, "");
const parseYmd = (s) => /^\d{8}$/.test(s) ? new Date(Date.UTC(+s.slice(0, 4), +s.slice(4, 6) - 1, +s.slice(6, 8))) : null;

async function cntrct(res, url, key) {
  const biz = url.searchParams.get("biz") || "servc";
  if (!COPS[biz]) return send(res, 400, { error: "biz 는 servc · cnstwk · thng 입니다." });
  const code = (url.searchParams.get("code") || "").replace(/\W/g, "");
  const name = str(url.searchParams.get("name"));
  if (!code && !name) return send(res, 400, { error: "기관코드나 기관명이 필요합니다." });
  const from = parseYmd(url.searchParams.get("from") || ""), to = parseYmd(url.searchParams.get("to") || "");
  if (!from || !to || to < from) return send(res, 400, { error: "기간(from · to, YYYYMMDD)을 확인해 주세요." });
  if ((to - from) / 864e5 + 1 > MAX_SPAN) return send(res, 400, { error: `한 번에 ${MAX_SPAN}일까지만 찾습니다.` });
  const div = url.searchParams.get("div") || "1"; // 기관구분코드 1 = 계약기관(2026-11-19 확인: 학교 코드로 그 학교가 맺은 계약이 나온다)
  const peek = url.searchParams.get("peek") === "1";

  const chunks = [];
  for (let a = from; a <= to; a = new Date(a.getTime() + CHUNK * 864e5)) {
    const b = new Date(Math.min(a.getTime() + (CHUNK - 1) * 864e5, to.getTime()));
    chunks.push([ymd(a), ymd(b)]);
  }
  const ask = async ([a, b], pageNo = 1) => {
    const p = new URLSearchParams({ serviceKey: key, pageNo: String(pageNo), numOfRows: "100", type: "json",
      inqryDiv: "1", inqryBgnDate: a, inqryEndDate: b, insttDivCd: div });
    if (code) p.set("insttCd", code); else p.set("insttNm", name);
    const r = await fetch(CBASE + COPS[biz] + "?" + p, { signal: AbortSignal.timeout(8000) });
    const text = await r.text();
    let out;
    try { out = unwrap(JSON.parse(text)); }
    catch { out = { code: "HTTP" + r.status, msg: (text.match(/<returnAuthMsg>([^<]*)/) || [])[1] || "응답을 읽지 못함", rows: [] }; }
    if (!out.rows.length && out.code && out.code !== "00" && out.code !== "03") throw Object.assign(new Error(out.msg), { code: out.code });
    // 한 주에 100건 넘으면 다음 쪽까지(학교 하나로는 거의 없다)
    if (out.rows.length === 100 && pageNo < 5) out.rows.push(...(await ask([a, b], pageNo + 1)));
    return out.rows;
  };

  try {
    const rows = (await Promise.all(chunks.map((c) => ask(c)))).flat()
      .sort((a, b) => str(b.rgstDt).localeCompare(str(a.rgstDt))); // 같은 계약이 여럿이면 가장 나중 등록을 남긴다
    if (peek) return send(res, 200, { chunks, raw: rows.slice(0, 3) });
    const seen = new Set();
    const items = rows.map((x) => normalizeCntrct(biz, x)).filter((x) => {
      const k = x.no || `${x.name}|${x.start}`;
      if (!x.name || seen.has(k)) return false;
      seen.add(k); return true;
    });
    const past = to.getTime() < Date.now() - 864e5;
    return send(res, 200, { items }, past ? "public, s-maxage=86400, max-age=3600" : "public, s-maxage=600, max-age=300");
  } catch (e) {
    if (e.code) return send(res, 502, { error: whyCode(e.code, e.message, "계약정보서비스") });
    return send(res, 504, { error: "조달청에 연결하지 못했습니다. 잠시 뒤 다시 해 주세요." });
  }
}

function whyCode(code, msg, what = "") {
  return /^(30|31)$/.test(code) ? "인증키가 아직 등록되지 않았습니다(발급 뒤 1~2시간 걸릴 수 있습니다)."
    : code === "22" ? "오늘 조회 한도를 다 썼습니다. 내일 다시 해 주세요."
    : code === "20" ? `이 인증키에 ${what || "해당 기능"} 사용 권한이 없습니다(공공데이터포털에서 활용신청이 필요합니다).`
    : code === "07" ? "조회 기간이 너무 깁니다(조달청 제한)."
    : `조달청 응답: ${msg || code}`;
}

export default async function handler(req, res) {
  if (req.method !== "GET") return send(res, 405, { error: "GET 만 됩니다." });
  // 다른 사이트에서 이 중계를 끌어다 쓰지 못하게
  if (req.headers["sec-fetch-site"] === "cross-site") return send(res, 403, { error: "이 사이트에서만 쓸 수 있습니다." });

  const url = new URL(req.url, "http://local");
  const kind = url.searchParams.get("kind");
  const no = (url.searchParams.get("q") || "").replace(/\D/g, "");
  if (!OPS[kind] && kind !== "cntrct") return send(res, 400, { error: "kind 는 corp · inst · cntrct 입니다." });
  if (kind !== "cntrct" && no.length !== 10) return send(res, 400, { error: "사업자등록번호(고유번호) 10자리로만 찾을 수 있습니다." });

  let key = (process.env.DATA_GO_KR_KEY || "").trim();
  if (!key) return send(res, 503, { error: "서버에 공공데이터포털 인증키가 없습니다." });
  if (/%[0-9A-F]{2}/i.test(key)) { try { key = decodeURIComponent(key); } catch {} } // Encoding 키를 넣었어도
  if (kind === "cntrct") return cntrct(res, url, key);

  try {
    const p = new URLSearchParams({ serviceKey: key, pageNo: "1", numOfRows: "10", type: "json", inqryDiv: "3", bizno: no });
    const r = await fetch(BASE + OPS[kind] + "?" + p, { signal: AbortSignal.timeout(9000) });
    const text = await r.text();
    let out;
    try { out = unwrap(JSON.parse(text)); }
    catch { out = { code: "HTTP" + r.status, msg: (text.match(/<returnAuthMsg>([^<]*)/) || [])[1] || "응답을 읽지 못함", rows: [] }; }

    if (!out.rows.length && out.code && out.code !== "00" && out.code !== "03") {
      return send(res, 502, { error: whyCode(out.code, out.msg) });
    }
    const items = out.rows.map((x) => normalize(kind, x)).filter((x) => x.name || x.bizno);
    return send(res, 200, { items }, "public, s-maxage=86400, max-age=600");
  } catch {
    return send(res, 504, { error: "조달청에 연결하지 못했습니다. 잠시 뒤 다시 해 주세요." });
  }
}
export { normalizeCntrct }; // 시험용(web/test)
