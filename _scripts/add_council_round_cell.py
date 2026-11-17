# -*- coding: utf-8 -*-
"""08_협의체회의록.hwpx 의 '회의일시' 값 칸에 회의일과 회차를 넣는다. (한 번만 돌린다)

왜 inject_tags_v2.py 로 하지 않는가
-----------------------------------
그 스크립트가 적어 둔 내용과 지금 서식 파일이 이미 다르다. 사용자가 그 뒤로
한글에서 직접 손봤기 때문이다. 다시 돌리면 그 손질이 통째로 날아간다.
그래서 **그 칸 하나만** 고치는 일회성 스크립트를 따로 둔다.

무엇을 고치는가
---------------
회의록 표의 '회의일시' 값 칸(colAddr=2 rowAddr=2)은 지금까지 **빈칸인 채로**
출력되고 있었다. 협의체 회의록을 월 1회 N회차로 뽑기 시작하면 열두 장이
날짜만 다른 채 나오므로, 여기에 회의일과 회차를 찍어 몇 회차인지 보이게 한다.

    {{협의체회의일}}   {{회차}}   →   2026.3.25.(수)   3회차

  · 칸 안에 문단은 이미 있고 글자만 없다(빈 run). 문단을 새로 만들지 않는다 —
    문단 없는 표 칸은 한글이 열다 오류 462 로 죽는다.
  · 글자색은 검정(charPr 49)으로 쓴다. 이 칸은 코드가 다 채우므로 사용자가
    고칠 일이 없다. 파란색은 '여기를 고치세요' 라는 뜻이라 쓰면 안 된다.
    (원래 들어 있던 charPr 71 은 71 과 모든 것이 같고 색만 파랑이다)
  · 글자 수가 달라지므로 그 문단의 <hp:linesegarray> 는 지운다. 낡은 줄 정보를
    남겨 두면 한글이 첫 줄을 엉뚱한 자리에 그린다. **그 문단 것만** 지운다.

돌린 뒤에는 반드시
------------------
    python _scripts/gen_template_tags.py

modTemplateTags 에 {{회차}} 가 없으면 RenderTemplate 이 그 태그만 건너뛰어
출력물에 '{{회차}}' 글자가 그대로 남는다.

    python _scripts/add_council_round_cell.py
"""
import os
import re
import sys
import zipfile

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
TMPL = os.path.join(ROOT, "form", "templates", "08_협의체회의록.hwpx")

SECTION = "Contents/section0.xml"
CELL_ADDR = '<hp:cellAddr colAddr="2" rowAddr="2"/>'
EMPTY_RUN = '<hp:run charPrIDRef="71"/>'
BLACK_CHARPR = "49"                  # charPr 71 과 같고 색만 검정
NEW_TEXT = "{{협의체회의일}}   {{회차}}"

LINESEG_RE = re.compile(r"<hp:linesegarray>.*?</hp:linesegarray>", re.S)


def read_entries(path):
    with zipfile.ZipFile(path) as z:
        return ({n: z.read(n) for n in z.namelist()},
                {i.filename: i for i in z.infolist()})


def write_entries(path, entries, info):
    """옆에 새 파일을 만든 뒤 바꿔치기한다. 쓰다 멈춰도 원본이 남는다."""
    tmp = path + ".new"
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as z:
        for name, data in entries.items():
            ct = zipfile.ZIP_STORED if name == "mimetype" else zipfile.ZIP_DEFLATED
            zi = zipfile.ZipInfo(name, date_time=info[name].date_time)
            zi.compress_type = ct
            zi.external_attr = info[name].external_attr
            z.writestr(zi, data)
    os.replace(tmp, path)


def target_cell(xml):
    """회의일시 값 칸의 (시작, 끝) 자리. 못 찾거나 여럿이면 None."""
    hits = []
    for m in re.finditer(r"<hp:tc\b.*?</hp:tc>", xml, re.S):
        cell = m.group(0)
        if CELL_ADDR in cell and EMPTY_RUN in cell:
            hits.append((m.start(), m.end()))
    return hits[0] if len(hits) == 1 else None


def main():
    sys.stdout.reconfigure(encoding="utf-8")

    if not os.path.exists(TMPL):
        print("서식을 찾을 수 없습니다:", TMPL)
        return 1

    entries, info = read_entries(TMPL)
    xml = entries[SECTION].decode("utf-8")

    if NEW_TEXT in xml:
        print("  이미 들어 있습니다 — 그대로 둡니다.")
        return 0

    span = target_cell(xml)
    if span is None:
        print("  ★ 회의일시 값 칸을 하나로 짚지 못했습니다. 서식이 바뀌었는지 확인하세요.")
        return 1

    a, b = span
    cell = xml[a:b]
    cell = cell.replace(
        EMPTY_RUN,
        f'<hp:run charPrIDRef="{BLACK_CHARPR}"><hp:t>{NEW_TEXT}</hp:t></hp:run>', 1)
    cell = LINESEG_RE.sub("", cell, count=1)      # 이 칸의 문단 하나뿐이다

    entries[SECTION] = (xml[:a] + cell + xml[b:]).encode("utf-8")
    write_entries(TMPL, entries, info)

    print(f"  회의일시 칸에 '{NEW_TEXT}' 를 넣었습니다.")
    print("  이제 python _scripts/gen_template_tags.py 를 돌리세요.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
