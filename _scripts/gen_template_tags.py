# -*- coding: utf-8 -*-
"""서식별 태그 목록표(modTemplateTags.bas)를 만든다.

RenderTemplate 은 서식 한 장마다 BuildTagMap 이 만든 태그를 **전부** 찾아바꾸기
했다. 태그는 126개인데 서식 하나에 실제로 박혀 있는 것은 3~37개뿐이라, 열에
아홉은 찾아도 안 나오는 것을 문서 전체에서 뒤지는 헛일이었다.

여기서 서식마다 '이 파일에 실제로 있는 태그' 를 뽑아 VBA 모듈로 써 둔다.
RenderTemplate 은 그 목록에 있는 것만 바꾼다.

태그를 뽑는 일은 verify_templates.tags_in() 이 이미 하고 있다. 문단 단위로
글자를 이어 붙인 뒤 찾으므로 여러 run 으로 쪼개진 태그도 잡는다 ― 한글
AllReplace 가 찾아내는 범위와 같다.

    python _scripts/gen_template_tags.py
"""
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from verify_templates import TDIR, tags_in

sys.stdout.reconfigure(encoding="utf-8")

OUT = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                   "vba", "modTemplateTags.bas")

# VBA 물리 줄 한도는 1023자다. 넉넉히 잡아 이보다 길어지면 " & _" 로 끊는다.
# (줄잇기는 한 문장에 25번까지 되므로 900자 × 25 = 22,000자까지 버틴다)
MAX_LINE = 900

HEAD = '''Attribute VB_Name = "modTemplateTags"
Option Explicit

'==============================================================================
' 서식별 태그 목록 ― 자동 생성 파일. 직접 고치지 말 것.
'
' _scripts/gen_template_tags.py 가 form/templates/*.hwpx 를 읽어 다시 만든다.
' 서식을 고쳐 태그가 늘거나 줄면 그 스크립트를 다시 돌려야 한다.
' (verify_templates.py 가 이 표와 서식이 어긋나 있는지 검사한다)
'
' 서식에 없는 태그까지 AllReplace 를 걸면 문서 전체를 훑는 헛일이 서식마다
' 100 번쯤 쌓인다. RenderTemplate 은 여기 적힌 것만 바꾼다.
'==============================================================================

' 서식 파일 이름 → 그 서식에 든 태그를 "|" 로 이은 것.
' 목록에 없는 서식이면 "" 를 돌려준다 ― 부르는 쪽은 그때 예전처럼 전부 바꾼다.
Public Function TemplateTagList(ByVal tmplName As String) As String
    Select Case tmplName
'''

TAIL = '''    End Select
End Function
'''


def vba_string(s):
    """VBA 문자열 리터럴. 큰따옴표는 두 번 겹쳐 쓴다."""
    return '"' + s.replace('"', '""') + '"'


def emit_case(fn, tags):
    """Case 한 줄. 길면 " & _" 로 끊어 여러 줄로 낸다."""
    indent = " " * 8
    head = f"{indent}Case {vba_string(fn)}: TemplateTagList = "

    # 끊을 때 조각 앞의 "|" 를 반드시 데려가야 한다. 두고 오면 VBA 가 조각을
    # 이어 붙였을 때 그 자리의 구분자가 사라져 두 태그가 한 덩어리로 붙는다.
    chunks, cur = [], ""
    for tag in tags:
        piece = tag if not cur else "|" + tag
        if cur and len(cur) + len(piece) > MAX_LINE:
            chunks.append(cur)
            cur = piece
        else:
            cur += piece
    chunks.append(cur)

    joiner = " & _\n" + " " * (len(indent) + 4)
    return head + joiner.join(vba_string(c) for c in chunks)


def build():
    rows = []
    for fn in sorted(os.listdir(TDIR)):
        if not fn.endswith(".hwpx"):
            continue
        # 정렬해 둔다. 순서가 흔들리면 어디가 바뀐 것인지 알기 어렵다.
        rows.append((fn, sorted(tags_in(os.path.join(TDIR, fn)))))
    return rows


def main():
    rows = build()
    if not rows:
        print(f"★ 서식을 하나도 찾지 못했습니다: {TDIR}")
        return 1

    body = "\n".join(emit_case(fn, tags) for fn, tags in rows)
    with open(OUT, "w", encoding="utf-8", newline="\r\n") as f:
        f.write(HEAD + body + "\n" + TAIL)

    total = sum(len(t) for _, t in rows)
    for fn, tags in rows:
        print(f"  {fn:<28} 태그 {len(tags):>2}개")
    root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    print(f"\n서식 {len(rows)}종, 태그 자리 모두 {total}개")
    print("생성:", os.path.relpath(OUT, root))
    return 0


if __name__ == "__main__":
    sys.exit(main())
