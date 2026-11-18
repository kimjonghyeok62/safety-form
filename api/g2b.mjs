// 나라장터(조달청 사용자정보 서비스) 찾기 중계 ― Vercel 함수 /api/g2b (ES 모듈이라 .mjs)
//   ?kind=corp&q=사업자등록번호   → 조달업체 기본정보(업체명 · 대표자 · 주소 · 전화 · 팩스)
//   ?kind=inst&q=사업자등록번호   → 수요기관 정보(학교 · 교육지원청의 고유번호로 기관명 · 주소 · 전화 · 팩스)
// 공공데이터포털 인증키(DATA_GO_KR_KEY)를 웹페이지에 드러내지 않으려고 둔다.
// 묻는 번호만 조달청에 넘기고 아무것도 저장 · 기록하지 않는다.
// 개발 중에는 web/vite.config.js 가 같은 함수를 /api/g2b 로 붙인다(키는 web/.env.local).
//
// 2026-11-18 실제 응답으로 확인한 것:
//   inqryDiv=3 + bizno → 사업자등록번호로 한 건. 이름만으로는 못 찾는다(08 필수값 입력 에러).
//   inqryDiv=1 · 2 는 등록 · 변경 일시 구간이 필수이고 구간이 길면 07 입력범위값 초과 에러.

const BASE = "https://apis.data.go.kr/1230000/ao/UsrInfoService02/";
const OPS = { corp: "getPrcrmntCorpBasicInfo02", inst: "getDminsttInfo02" };

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

export default async function handler(req, res) {
  if (req.method !== "GET") return send(res, 405, { error: "GET 만 됩니다." });
  // 다른 사이트에서 이 중계를 끌어다 쓰지 못하게
  if (req.headers["sec-fetch-site"] === "cross-site") return send(res, 403, { error: "이 사이트에서만 쓸 수 있습니다." });

  const url = new URL(req.url, "http://local");
  const kind = url.searchParams.get("kind");
  const no = (url.searchParams.get("q") || "").replace(/\D/g, "");
  if (!OPS[kind]) return send(res, 400, { error: "kind 는 corp 또는 inst 입니다." });
  if (no.length !== 10) return send(res, 400, { error: "사업자등록번호(고유번호) 10자리로만 찾을 수 있습니다." });

  let key = (process.env.DATA_GO_KR_KEY || "").trim();
  if (!key) return send(res, 503, { error: "서버에 공공데이터포털 인증키가 없습니다." });
  if (/%[0-9A-F]{2}/i.test(key)) { try { key = decodeURIComponent(key); } catch {} } // Encoding 키를 넣었어도

  try {
    const p = new URLSearchParams({ serviceKey: key, pageNo: "1", numOfRows: "10", type: "json", inqryDiv: "3", bizno: no });
    const r = await fetch(BASE + OPS[kind] + "?" + p, { signal: AbortSignal.timeout(9000) });
    const text = await r.text();
    let out;
    try { out = unwrap(JSON.parse(text)); }
    catch { out = { code: "HTTP" + r.status, msg: (text.match(/<returnAuthMsg>([^<]*)/) || [])[1] || "응답을 읽지 못함", rows: [] }; }

    if (!out.rows.length && out.code && out.code !== "00" && out.code !== "03") {
      const why = /^(30|31)$/.test(out.code) ? "인증키가 아직 등록되지 않았습니다(발급 뒤 1~2시간 걸릴 수 있습니다)."
        : out.code === "22" ? "오늘 조회 한도를 다 썼습니다. 내일 다시 해 주세요."
        : out.code === "20" ? "이 인증키에 해당 기능 사용 권한이 없습니다."
        : `조달청 응답: ${out.msg || out.code}`;
      return send(res, 502, { error: why });
    }
    const items = out.rows.map((x) => normalize(kind, x)).filter((x) => x.name || x.bizno);
    return send(res, 200, { items }, "public, s-maxage=86400, max-age=600");
  } catch {
    return send(res, 504, { error: "조달청에 연결하지 못했습니다. 잠시 뒤 다시 해 주세요." });
  }
}
