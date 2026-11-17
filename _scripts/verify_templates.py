# -*- coding: utf-8 -*-
"""서식 파일의 {{태그}} 가 모두 채워질 수 있는지 맞대어 본다.

한글 서식에 태그를 적어 두었는데 엑셀 쪽에 그 값을 채울 자리가 없으면,
출력한 문서에 {{관할소방서}} 같은 글자가 그대로 남는다. 눈으로 잡기 어렵고
한 번 놓치면 계속 남으므로, 여기서 기계적으로 대조한다.

  · 서식에는 있는데 엑셀이 못 채우는 태그  → ★ 고쳐야 한다
  · 엑셀은 채울 수 있는데 아무 서식도 안 쓰는 태그 → 참고만 (문제는 아니다)

한 문서로 이어붙는 서식끼리 용지·여백이 같은지도 함께 본다
(modHwpExport.bas 의 InsertFile 이 그 전제 위에 서 있다).

    python _scripts/verify_templates.py
"""
import os
import re
import sys
import zipfile

import win32com.client

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from build_workbook import BOOK, IN_FIRST_ROW

sys.stdout.reconfigure(encoding="utf-8")

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
TDIR = os.path.join(ROOT, "form", "templates")

TAG_RE = re.compile(r"\{\{[^}]*\}\}")


def tags_in(path):
    """서식 파일 안의 태그. 태그가 여러 run 으로 쪼개져 있어도 잡도록
    문단 단위로 글자를 이어 붙인 뒤 찾는다."""
    found = set()
    with zipfile.ZipFile(path) as z:
        for name in z.namelist():
            if not re.match(r"Contents/section\d+\.xml", name):
                continue
            xml = z.read(name).decode("utf-8")
            for para in re.split(r"<hp:p\b", xml):
                text = "".join(re.findall(r"<hp:t>(.*?)</hp:t>", para, re.S))
                text = (text.replace("&lt;", "<").replace("&gt;", ">")
                            .replace("&amp;", "&"))
                found.update(TAG_RE.findall(text))
    return found


# 한 문서로 이어붙는 서식 묶음. modHwpExport.bas 의 InsertFile 은 구역(용지·여백)을
# 가져오지 않고 첫 문서의 것을 따르므로, 한 묶음 안에서는 용지·여백이 같아야 한다.
# 09·10·11~15 는 같은 서식을 여러 장 찍어 붙일 뿐이라 따로 볼 것이 없다.
# 서식마다 **반드시 살아 있어야 하는 태그**.
#
# 서식의 예시 글을 태그로 갈아 끼우다 보면 그 줄에 함께 있던 {{학교명}} 같은
# 태그까지 날아갈 수 있다. 눈으로는 잘 안 보이고, 한 번 놓치면 출력물에서
# 학교 이름 자리가 계속 비어 나온다. 그래서 기계적으로 지킨다.
#
# 여기 없는 태그가 서식에 더 있는 것은 괜찮다. '이것만은 빠지면 안 된다' 는 목록이다.
REQUIRED_TAGS = {
    "01_계획서_공사.hwpx": [
        "{{학교명}}", "{{계약명}}", "{{업체명}}", "{{안전담당자}}", "{{업체전화}}",
        "{{계약금액천원}}", "{{계약기간}}",
        "{{소방서}}", "{{소방서전화번호}}", "{{병원}}", "{{병원전화번호}}",
        "{{관할고용노동부}}", "{{고용노동부전화번호}}",
        "{{작업절차}}", "{{안전점검계획}}", "{{보호구방호장치}}", "{{안전보건교육}}",
        "{{위험요인1}}", "{{위험요인2}}", "{{안전대책1}}", "{{안전대책2}}",
    ],
    "01_계획서_용역.hwpx": [
        "{{학교명}}", "{{계약명}}", "{{업체명}}", "{{안전담당자}}", "{{업체전화}}",
        "{{계약금액천원}}", "{{계약기간}}",
        "{{소방서}}", "{{소방서전화번호}}", "{{병원}}", "{{병원전화번호}}",
        "{{관할고용노동부}}", "{{고용노동부전화번호}}",
        "{{작업절차}}", "{{안전점검계획}}", "{{보호구방호장치}}", "{{안전보건교육}}",
        "{{위험요인1}}", "{{위험요인2}}", "{{안전대책1}}", "{{안전대책2}}",
    ],
    "01_계획서_상세.hwpx": [
        "{{학교명}}", "{{계약명}}", "{{업체명}}", "{{업체대표자}}", "{{업체전화}}",
        "{{안전담당자}}", "{{안전담당자연락처}}", "{{업체대표자연락처}}",
        "{{행정실담당자}}", "{{행정실전화번호}}",
        "{{계약금액천원}}", "{{계약기간}}",
        "{{관할소방서}}", "{{소방서전화번호}}", "{{인근병원}}", "{{병원전화번호}}",
        "{{관할고용노동부}}", "{{고용노동부전화번호}}",
        "{{작업내용}}",
        # ④ 는 일부러 뺐다. 학교 현장 항목을 넣으면서 두 쪽을 지키려고 ③ 까지만 쓴다.
        "{{위험요인1}}", "{{위험요인3}}", "{{안전대책1}}", "{{안전대책3}}",
        "{{사용기계기구1}}", "{{사용기계기구3}}",
        "{{위험물질설비1}}", "{{위험물질설비3}}",
        "{{작업권장시기}}", "{{위험성등급}}",
        # 3번째 항목은 일부러 안 쓴다. 두 줄로 묶어 오른쪽 칸과 줄 수를 맞췄다.
        "{{학교확인사항1}}", "{{학교확인사항2}}",
        "{{업체조치사항1}}", "{{업체조치사항2}}",
    ],
    "08_협의체회의록.hwpx": [
        "{{학교명}}", "{{업체명}}", "{{행정실전화번호}}",
        "{{업체담당자}}", "{{업체담당자연락처}}",
        "{{담당자}}", "{{관리자}}", "{{부서장}}",
        # {{계약명}} 은 일부러 뺐다. 회의록은 계약 한 건을 짚는 문서가 아니라
        # 그 자리에서 협의한 것을 적는 서식이라 계약명 칸을 두지 않기로 했다.
        # 회의일시 칸 — 월 1회 N회차로 뽑으므로 몇 회차인지 보여야 한다
        "{{협의체회의일}}", "{{회차}}", "{{대피방법}}",
        # 위험요인③·안전대책② 는 일부러 뺐다. 6번(의결사항)을 넣을 자리다.
        "{{위험요인1}}", "{{위험요인2}}", "{{안전대책1}}",
        "{{상시안건1}}",
        "{{협의체합의문1}}", "{{협의체합의문3}}",
    ],
    "09_순회점검표.hwpx": [
        "{{학교명}}", "{{점검년}}", "{{점검월}}",
        # 점검일 칸이 17mm 밖에 안 되어 날짜와 요일을 두 줄로 나눠 넣는다.
        # 요일 태그가 사라지면 '2026.6.4.(' 에서 잘리던 옛 모습으로 돌아간다.
        "{{점검일1}}", "{{점검일5}}", "{{점검요일1}}", "{{점검요일5}}",
    ],
    "12_관리감독자평가표.hwpx": [
        # 관리감독자는 부서장(학교면 교장)이다. 관리자(행정실장)가 아니다.
        "{{학교명}}", "{{부서장직위}}", "{{부서장성명}}", "{{작성일}}",
    ],
    "15_계약이행서약서.hwpx": [
        "{{계약명}}", "{{기관명}}", "{{업체명}}", "{{업체대표자}}",
        "{{사업자등록번호}}", "{{작성일}}",
    ],
}


def check_required(tags_by_file):
    """빠지면 안 되는 태그가 다 살아 있는가."""
    bad = False
    for fn, want in sorted(REQUIRED_TAGS.items()):
        gone = [t for t in want if t not in tags_by_file.get(fn, set())]
        if gone:
            bad = True
            print(f"  ★ {fn} ― 있어야 할 태그가 사라졌습니다: {' '.join(gone)}")
        else:
            print(f"  OK {fn} ― 필수 태그 {len(want)}개 모두 살아 있습니다")
    return bad


# 한 문서로 이어붙는 서식끼리는 용지·여백(구역)이 같아야 한다.
# 08_협의체회의록 은 이제 1~7 묶음에서 빠졌다 — 월 1회 N회차로 자기 문서를 따로 낸다.
# (같은 서식만 이어 붙으므로 저희끼리는 언제나 맞는다)
MERGE_GROUPS = {
    "1~7번 묶음": ["01_계획서_공사.hwpx", "01_계획서_용역.hwpx", "01_계획서_상세.hwpx",
                 "02_수준평가표.hwpx", "03_도급점검표.hwpx",
                 "04_점검표_고소.hwpx", "05_점검표_화기.hwpx",
                 "06_점검표_밀폐.hwpx", "07_점검표_전기.hwpx"],
}


TAGTABLE = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                        "vba", "modTemplateTags.bas")

# Case "01_계획서_공사.hwpx": TemplateTagList = "{{…}}|{{…}}" & _
CASE_RE = re.compile(r'^\s*Case\s+"([^"]+)":\s*TemplateTagList\s*=\s*(.+)$')


def table_tags():
    """modTemplateTags.bas 가 들고 있는 서식별 태그."""
    if not os.path.exists(TAGTABLE):
        return None

    out, cur = {}, None
    with open(TAGTABLE, encoding="utf-8") as f:
        for line in f:
            m = CASE_RE.match(line)
            if m:
                cur = m.group(1)
                out[cur] = ""
                rest = m.group(2)
            elif cur is not None and line.lstrip().startswith('"'):
                rest = line.strip()          # " & _" 로 이어진 다음 줄
            else:
                continue
            # 그 줄에 있는 문자열 조각들을 이어 붙인다 (" & _" 는 버린다)
            out[cur] += "".join(re.findall(r'"([^"]*)"', rest))
            if not rest.rstrip().endswith("_"):
                cur = None

    return {fn: set(t for t in s.split("|") if t) for fn, s in out.items()}


def check_tag_table(tags_by_file):
    """RenderTemplate 이 보는 태그 목록표가 서식과 맞는가.

    어긋나면 그 서식의 태그 일부가 치환되지 않아 출력물에 {{태그}} 가 남는다."""
    table = table_tags()
    if table is None:
        print("  ★ modTemplateTags.bas 가 없습니다."
              " python _scripts/gen_template_tags.py 를 돌리세요.")
        return True

    bad = False
    for fn in sorted(tags_by_file):
        want, got = tags_by_file[fn], table.get(fn)
        if got is None:
            # 표에 없는 서식은 '전부 치환' 으로 되돌아간다 ― 느릴 뿐 깨지지는 않는다.
            print(f"  · {fn} ― 표에 없습니다 (전부 치환으로 돌아갑니다. 느립니다)")
            continue
        if want != got:
            bad = True
            miss = sorted(want - got)
            extra = sorted(got - want)
            print(f"  ★ {fn} ― 표가 서식과 다릅니다.")
            if miss:
                print(f"     표에 빠진 태그(치환 안 됩니다): {' '.join(miss)}")
            if extra:
                print(f"     서식에 없는데 표에 있는 것: {' '.join(extra)}")

    for fn in sorted(set(table) - set(tags_by_file)):
        bad = True
        print(f"  ★ {fn} ― 표에는 있는데 서식 파일이 없습니다.")

    if not bad:
        print(f"  OK 서식 {len(tags_by_file)}종의 태그 목록표가 서식과 맞습니다")
    return bad


def page_setup(path):
    """서식의 용지 크기와 여백. 이어붙일 때 이것이 다르면 뒤 서식이 찌그러진다."""
    with zipfile.ZipFile(path) as z:
        xml = z.read("Contents/section0.xml").decode("utf-8")
    pp = re.search(r"<hp:pagePr\b[^>]*>", xml)
    mg = re.search(r"<hp:margin\b[^>]*/>", xml)
    return (pp.group(0) if pp else "", mg.group(0) if mg else "")


def check_page_setup():
    """이어붙는 서식끼리 용지·여백이 같은가."""
    bad = False
    for label, files in MERGE_GROUPS.items():
        seen = {}
        for fn in files:
            path = os.path.join(TDIR, fn)
            if not os.path.exists(path):
                print(f"  ★ {label}: {fn} 이 없습니다")
                bad = True
                continue
            seen.setdefault(page_setup(path), []).append(fn)
        if len(seen) <= 1:
            print(f"  OK {label} ― 용지·여백이 모두 같습니다")
            continue
        bad = True
        print(f"  ★ {label} ― 용지·여백이 {len(seen)}가지로 갈립니다."
              " 이대로 이어붙이면 뒤 서식의 여백이 첫 서식을 따라갑니다.")
        for i, (setup, names) in enumerate(seen.items(), 1):
            print(f"     [{i}] {' '.join(names)}")
            print(f"         {setup[0]}")
            print(f"         {setup[1]}")
    return bad


def known_tags():
    """BuildTagMap 이 만드는 태그 이름."""
    xl = win32com.client.Dispatch("Excel.Application")
    xl.Visible = False
    xl.DisplayAlerts = False
    xl.AutomationSecurity = 1
    xl.EnableEvents = False
    wb = xl.Workbooks.Open(BOOK)
    try:
        return set(str(xl.Run("GetTagList", IN_FIRST_ROW)).split("|"))
    finally:
        wb.Close(SaveChanges=False)
        xl.EnableEvents = True
        xl.Quit()


def main():
    known = known_tags()
    print(f"엑셀이 채울 수 있는 태그: {len(known)}개\n")

    used, missing_any = set(), False
    by_file = {}
    for fn in sorted(os.listdir(TDIR)):
        if not fn.endswith(".hwpx"):
            continue
        tags = tags_in(os.path.join(TDIR, fn))
        by_file[fn] = tags
        used |= tags
        missing = sorted(t for t in tags if t not in known)
        mark = "OK" if not missing else "★ 못 채움: " + " ".join(missing)
        print(f"  {fn:<28} 태그 {len(tags):>2}개  {mark}")
        if missing:
            missing_any = True

    spare = sorted(known - used)
    if spare:
        print(f"\n쓰이지 않는 태그 {len(spare)}개 (문제는 아닙니다):")
        print("  " + " ".join(spare))

    print("\n빠지면 안 되는 태그")
    required_bad = check_required(by_file)

    print("\n서식별 태그 목록표 (RenderTemplate 이 이것만 바꾼다)")
    table_bad = check_tag_table(by_file)

    print("\n이어붙는 서식의 용지·여백")
    setup_bad = check_page_setup()

    print("\n" + ("★ 채우지 못하는 태그가 있습니다. 서식이나 태그 이름을 맞추세요."
                  if missing_any else "모든 서식의 태그를 채울 수 있습니다."))
    if required_bad:
        print("★ 서식에서 태그가 사라졌습니다. 태그를 심는 과정에서 지워진 것입니다.")
    if table_bad:
        print("★ 태그 목록표가 낡았습니다."
              " python _scripts/gen_template_tags.py 를 다시 돌리세요.")
    if setup_bad:
        print("★ 이어붙는 서식끼리 용지·여백이 다릅니다.")
    return 1 if (missing_any or required_bad or table_bad or setup_bad) else 0


if __name__ == "__main__":
    sys.exit(main())
