// 구글 계정으로 이어쓰기에 쓰는 OAuth 클라이언트 ID.
// 웹페이지에 공개되는 값이라 비밀이 아니다(클라이언트 보안 비밀번호는 쓰지 않는다).
// 비어 있으면 [구글 계정으로 이어쓰기]가 "준비 중"으로 보이고 나머지는 그대로 동작한다.
export const GOOGLE_CLIENT_ID = "909748481776-3nnh9oeda5p2fgo0ev1mqts7vnf95ivo.apps.googleusercontent.com";

// 학교 찾기(교육부 나이스 학교기본정보)용 인증키. open.neis.go.kr 에서 무료로 받는다.
// 이것도 웹페이지에 공개되는 값이다. 비어 있어도 학교 찾기는 되지만 한 번에 5건까지만 온다.
export const NEIS_KEY = "";
