# -*- coding: utf-8 -*-
"""웹 버전이 쓰는 사전 데이터를 JSON 으로 뽑는다.

사전 원본은 contract_types.py 한 곳이다. 엑셀(build_workbook.py)과 웹이 같은
원본에서 나와야 두 출력이 어긋나지 않는다. 사전을 고쳤으면 이것도 다시 돌린다.

    python _scripts/export_web_data.py

또 form/templates/*.hwpx 를 web/public/templates/ 로 복사한다(빈 서식만).
"""
import json
import os
import shutil
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from contract_types import CONTENT_LIBRARY, contract_rows

sys.stdout.reconfigure(encoding="utf-8")

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUT = os.path.join(ROOT, "web", "src", "data")
TSRC = os.path.join(ROOT, "form", "templates")
TDST = os.path.join(ROOT, "web", "public", "templates")

# 계약유형사전 A~T 열 순서 (modConfig.bas 의 DICT_* 상수와 같은 차례)
FIELDS = ["유형명", "키워드", "작업내용", "위험요인", "안전대책", "협의안건", "점검착안",
          "작업유형", "작업절차", "점검계획", "보호구", "교육계획", "기계기구", "위험물질",
          "대피방법", "권장시기", "학교확인", "업체조치", "위험등급", "합의문"]


def workbook_constants(*names):
    """build_workbook.py 의 상수만 읽는다. 그 파일을 import 하면 pywin32 가 있어야 해서
    문법 트리에서 값만 꺼낸다."""
    import ast
    src = open(os.path.join(os.path.dirname(os.path.abspath(__file__)), "build_workbook.py"),
               encoding="utf-8").read()
    found = {}
    for node in ast.parse(src).body:
        if isinstance(node, ast.Assign) and len(node.targets) == 1:
            t = node.targets[0]
            if isinstance(t, ast.Name) and t.id in names:
                found[t.id] = ast.literal_eval(node.value)
    missing = [n for n in names if n not in found]
    if missing:
        raise KeyError("build_workbook.py 에 없는 상수: " + ", ".join(missing))
    return [found[n] for n in names]


def main():
    WORK_KEYWORDS, PLAN_RULES, CRITERIA = workbook_constants("WORK_KEYWORDS", "PLAN_RULES", "CRITERIA")
    os.makedirs(OUT, exist_ok=True)
    rows = contract_rows()
    types = []
    for r in rows:
        if len(r) != len(FIELDS):
            raise ValueError(f"칸 수가 다릅니다: {r[0]} ({len(r)})")
        types.append({k: str(v or "").strip() for k, v in zip(FIELDS, r)})

    lib = [{"작업유형": k, "구분": c, "내용": t} for k, c, t in CONTENT_LIBRARY]

    with open(os.path.join(OUT, "contract-types.json"), "w", encoding="utf-8") as f:
        json.dump(types, f, ensure_ascii=False, indent=1)
    with open(os.path.join(OUT, "content-library.json"), "w", encoding="utf-8") as f:
        json.dump(lib, f, ensure_ascii=False, indent=1)

    settings = {
        "작업유형키워드": [{"유형": n, "키워드": k} for n, k in WORK_KEYWORDS],
        "계획서키워드": {kind: kw for kind, kw, _ in PLAN_RULES if kind in ("공사", "용역")},
        "기준일수": {nm: d for _, d, _, nm in CRITERIA},
        "기본점검요일": "수요일",
    }
    with open(os.path.join(OUT, "settings.json"), "w", encoding="utf-8") as f:
        json.dump(settings, f, ensure_ascii=False, indent=1)

    os.makedirs(TDST, exist_ok=True)
    names = sorted(n for n in os.listdir(TSRC) if n.endswith(".hwpx"))
    for n in names:
        shutil.copy2(os.path.join(TSRC, n), os.path.join(TDST, n))
    with open(os.path.join(OUT, "templates.json"), "w", encoding="utf-8") as f:
        json.dump(names, f, ensure_ascii=False, indent=1)

    print(f"계약유형 {len(types)}종 · 문구 {len(lib)}줄 · 서식 {len(names)}개")


if __name__ == "__main__":
    main()
