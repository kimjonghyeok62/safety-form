# 중대재해예방 서식 — 웹판

엑셀(.xlsm) 매크로 없이 브라우저에서 같은 한글(hwpx) 서식을 만든다. 서버 없음.
학교 정보는 그 PC 브라우저(IndexedDB)와 사용자가 내려받는 `.json` 백업 파일에만 있다.

화면 틀은 원가계산서 제비율 스마트 체크(cost-checker_web)와 같은 디자인 언어다(남색 머리띠 · 왼쪽 메뉴 · Pretendard).

## 저장 · 이어쓰기 (자동 저장 대상은 한 번에 하나, `src/sync.js`)
- **구글 계정으로 이어쓰기**(`src/store/googleDrive.js`): Google Identity Services 토큰 + Drive API 로
  그 사람 드라이브의 앱 전용 숨김 폴더(`appDataFolder`)에 `safety-form-data.json` 하나를 둔다.
  권한은 `drive.appdata`(+ `email`) 뿐. 브라우저 ↔ 구글 직통이라 개발자 쪽을 거치지 않는다.
  **`src/config.js` 의 `GOOGLE_CLIENT_ID` 가 비어 있으면 "준비 중"으로 보인다.**
  토큰은 1시간짜리라 다음 날·다른 PC 에서는 [구글 계정으로 이어서 하기] 한 번을 누른다.
- **내 컴퓨터 파일에 자동 저장**(`src/store/fileLink.js`): File System Access API 로 고른 `.json` 파일에 쓴다.
  크롬·엣지·웨일만 된다.
- 쓰기 전에 대상을 다시 읽어 다른 PC 가 고쳤는지 본다. 대상만 바뀌었으면 가져오고, 양쪽 다 바뀌었으면 사용자가 고른다.

### 구글 클라우드 설정(한 번)
OAuth 동의 화면: 외부 · 범위 `drive.appdata`, `openid`, `email` · 프로덕션 게시.
OAuth 클라이언트 ID(웹): 승인된 자바스크립트 원본 `https://safety-form.vercel.app`, `http://localhost:5178`.
승인된 리디렉션 URI `https://safety-form.vercel.app/`, `http://localhost:5178/` — 팝업이 막힌 브라우저에서
이 창째 구글로 넘어갔다 돌아오는 로그인에 쓴다(끝의 `/` 까지 똑같아야 한다).

## 배포
Vercel 프로젝트 `safety-form`(GitHub 연결) — `main` 에 push 하면 https://safety-form.vercel.app 으로 나간다.
설정은 저장소 루트의 `vercel.json`. 옛 GitHub Pages 주소는 새 주소로 넘겨 주는 쪽만 남겼다(`.github/workflows/pages.yml`).
CLI 로 직접 올릴 때도 `.vercelignore` 가 `web/` 밖(학교 자료 · xlsm · 백업)을 막는다.

```
npm install
npm run dev            # http://localhost:5178
npm run build          # dist/ → 정적 호스팅(GitHub Pages 등)에 그대로 올린다
npm run render-sample  # 예시 데이터로 모든 서식을 test/out/ 에 만든다
```

## 사전·서식을 고쳤을 때
사전 원본은 `_scripts/contract_types.py`, 서식 원본은 `form/templates/`. 고친 뒤:

```
python _scripts/export_web_data.py     # → src/data/*.json, public/templates/*.hwpx
```

## 엑셀판과 같은지 확인
```
python _scripts/compare_web_output.py <엑셀출력.hwpx> <웹출력.hwpx>
```

## 구조
- `src/domain/` — VBA 를 옮긴 계산(판정·날짜·태그 값). 엑셀판과 결과가 같아야 한다.
- `src/hwpx/render.js` — 한글 COM 대신 hwpx XML 을 직접 고친다. 같은 서식 여러 장은 구역으로 복제.
- `src/export.js` — 출력 버튼 하나 = 함수 하나.
- `src/main.js` — 화면.
