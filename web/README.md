# 중대재해예방 서식 — 웹판

엑셀(.xlsm) 매크로 없이 브라우저에서 같은 한글(hwpx) 서식을 만든다. 서버 없음.
학교 정보는 그 PC 브라우저(IndexedDB)와 사용자가 내려받는 `.json` 백업 파일에만 있다.

**드라이브 폴더 자동 저장**(`src/store/fileLink.js`): 구글 로그인 없이, 구글 드라이브 데스크톱 폴더
(G:\내 드라이브 등)의 `.json` 파일 하나를 File System Access API 로 골라 두고 고칠 때마다 그 파일에 쓴다.
파일 핸들은 IndexedDB 에 보관하고, 새로고침 뒤 권한이 풀리면 [자동 저장 다시 켜기] 한 번으로 이어진다.
쓰기 전에 파일을 다시 읽어 다른 PC 가 고쳤는지 보고, 양쪽 다 고쳤으면 사용자가 고른다.
크롬·엣지·웨일만 되고, 그 밖의 브라우저는 버튼이 숨고 백업 파일 저장/불러오기를 쓴다.

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
