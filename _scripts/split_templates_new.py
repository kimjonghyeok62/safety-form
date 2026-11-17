# -*- coding: utf-8 -*-
"""form/서식모음(new).hwpx 에서 11~15번 서식을 잘라 form/templates/ 에 넣는다.

1~10번 템플릿은 사용자가 직접 손본 것이라 **건드리지 않는다.**
여기서 만드는 것은 신 양식에만 있는 다섯 가지다.

    11 아차사고 사례 및 유해위험발굴 제안서   (서식 1)
    12 관리감독자 업무수행 평가표             (서식 2)
    13 중대재해발생보고서                     (서식 10)
    14 중대재해사고보고서                     (서식 11)
    15 계약이행 통합 서약서                   (참고 2, 여러 쪽)

원본(form/서식모음(new).hwpx)은 절대 수정하지 않는다.

split_templates.py 와 같은 두 단계 방식이다.
  1) 한글 COM 으로 사본에서 대상 페이지 밖을 잘라낸다. 표는 건드리지 않는다.
  2) 저장된 hwpx(zip+XML)를 열어 남은 머리표 상자와 빈 껍데기를 XML 수준에서 도려낸다.

검증은 페이지를 이미지로 구워 잉크(비백색 픽셀) 비율을 원본의 같은 페이지와 대조한다.

사용법:
    python _scripts/split_templates_new.py --dry-run   # 머리표 → 페이지 매핑만 출력
    python _scripts/split_templates_new.py             # 생성 + 검증
"""
import os
import re
import shutil
import sys
import zipfile

import win32com.client

sys.stdout.reconfigure(encoding="utf-8")

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SRC = os.path.join(ROOT, "form", "서식모음(new).hwpx")
OUT_DIR = os.path.join(ROOT, "form", "templates")
WORK = os.path.join(ROOT, "form", "_split_work_new.hwpx")

# 구간 이름 → 그 구간의 첫 쪽을 찾을 검색어.
# 머리표('서식 1')로 찾으면 안 된다. 한글 찾기는 부분일치라 '서식 1' 이 '서식 10' 에 먼저 걸린다.
# 그래서 그 서식에만 있는 제목 글자로 찾는다.
# 검색어에 '안전·보건' 처럼 가운뎃점이 든 말을 넣으면 안 된다. 이 문서는 그 자리에
# 사용자정의 글리프(U+F09E)를 쓴 곳이 있어 눈에는 같아 보여도 찾기가 실패한다.
ALL_MARKERS = [
    ("서식1 아차사고",     "아차사고 사례 및"),
    ("서식2 관리감독자",   "관리감독자 업무수행 평가표"),
    ("서식3 수준평가",     "수급인 안전"),
    ("서식10 발생보고",    "중대재해발생보고서"),
    ("서식11 사고보고",    "중대재해사고보고서"),
    ("예시1 의견청취",     "간담회 회의 개최시"),
    ("참고2 서약서",       "계약이행 통합 서약서"),
]

# (출력파일명, 이 구간, 다음 구간, 이 서식 자신의 머리표)
# 페이지는 검색으로 자동 계산한다.
#
# 머리표('서식 8' 같은 제목 상자)는 남긴다. 사용자가 손본 1~10번 템플릿에도 그대로
# 남아 있기 때문이다. 한글 문서 안의 서식 번호는 매뉴얼의 고유번호라 우리 파일명
# 번호(01~15)와 달라도 맞는 것이다 — 03_도급점검표.hwpx 안이 '서식 6' 인 식이다.
# 잘라낸 자리에 딸려 온 **다른** 서식의 머리표만 도려낸다.
SPECS = [
    ("11_아차사고제안서.hwpx",     "서식1 아차사고",  "서식2 관리감독자", "서식 1"),
    ("12_관리감독자평가표.hwpx",   "서식2 관리감독자", "서식3 수준평가",  "서식2"),
    ("13_중대재해발생보고서.hwpx", "서식10 발생보고", "서식11 사고보고", "서식 10"),
    ("14_중대재해사고보고서.hwpx", "서식11 사고보고", "예시1 의견청취",  "서식 11"),
    ("15_계약이행서약서.hwpx",     "참고2 서약서",    None,              "참고2"),
]

MARKER_RE = re.compile(r"^(서식\s*\d+(?:-\d+)?|붙임\s*\d+(?:-\d+)?|참고\s*\d+)$")

# 확인 / 확인·취소 / 예·아니오·취소 / 예·아니오 / 재시도 팝업에 긍정 응답을 자동으로 누른다.
MSGBOX_AUTO = 0x1 | 0x10 | 0x1000 | 0x10000 | 0x100000


# ---------------------------------------------------------------- 한글 COM

def connect():
    hwp = win32com.client.Dispatch("HWPFrame.HwpObject")
    hwp.RegisterModule("FilePathCheckDLL", "FilePathCheckerModule")
    hwp.SetMessageBoxMode(MSGBOX_AUTO)
    return hwp


def goto_page(hwp, n):
    hwp.HAction.GetDefault("Goto", hwp.HParameterSet.HGotoE.HSet)
    hwp.HParameterSet.HGotoE.HSet.SetItem("DialogResult", n)
    hwp.HParameterSet.HGotoE.SetSelectionIndex = 1
    hwp.HAction.Execute("Goto", hwp.HParameterSet.HGotoE.HSet)


def find_forward(hwp, text):
    hwp.HAction.Run("MoveDocBegin")
    hwp.HAction.GetDefault("RepeatFind", hwp.HParameterSet.HFindReplace.HSet)
    fr = hwp.HParameterSet.HFindReplace
    fr.FindString = text
    fr.Direction = 0
    fr.IgnoreMessage = 1
    fr.MatchCase = 0
    fr.SeveralWords = 0
    fr.UseWildCards = 0
    fr.WholeWordOnly = 0
    fr.AutoSpell = 0
    return bool(hwp.HAction.Execute("RepeatFind", fr.HSet))


def ink_ratio(hwp, pgno, tmp_dir):
    """페이지를 이미지로 굽고 흰색이 아닌 픽셀 비율을 잰다."""
    from PIL import Image
    img = os.path.join(tmp_dir, f"_ink_n{pgno}.bmp")
    hwp.CreatePageImage(img, pgno, 96, 24, "BMP")
    if not os.path.exists(img):
        return None
    with Image.open(img) as im:
        g = im.convert("L")
        dark = sum(c for v, c in zip(range(256), g.histogram()) if v < 240)
        total = g.size[0] * g.size[1]
    os.remove(img)
    return dark / total


def cut_pages(hwp, first, last):
    """사본에서 [first..last] 밖을 잘라낸다. 표 개체는 손대지 않는다."""
    shutil.copyfile(SRC, WORK)
    hwp.Open(WORK, "HWPX", "forceopen:true")
    total = hwp.PageCount

    if last < total:
        goto_page(hwp, last + 1)
        hwp.HAction.Run("MovePageBegin")
        hwp.HAction.Run("MoveSelDocEnd")
        hwp.HAction.Run("Delete")

    if first > 1:
        goto_page(hwp, first)
        hwp.HAction.Run("MovePageBegin")
        hwp.HAction.Run("MoveSelDocBegin")
        hwp.HAction.Run("Delete")


# ------------------------------------------------------------- hwpx 손질

def top_level_paragraphs(xml):
    """<hs:sec> 바로 아래 <hp:p> 들의 (시작, 끝) 바이트 위치 목록."""
    spans, depth, start = [], 0, None
    for m in re.finditer(r"<hp:p\b|</hp:p>", xml):
        if m.group(0).startswith("</"):
            depth -= 1
            if depth == 0:
                spans.append((start, m.end()))
                start = None
        else:
            if depth == 0:
                start = m.start()
            depth += 1
    return spans


def outer_table_spans(xml):
    """가장 바깥쪽 <hp:tbl> 들의 (시작, 끝) 위치."""
    spans, depth, start = [], 0, None
    for m in re.finditer(r"<hp:tbl\b|</hp:tbl>", xml):
        if m.group(0).startswith("</"):
            depth -= 1
            if depth == 0:
                spans.append((start, m.end()))
                start = None
        else:
            if depth == 0:
                start = m.start()
            depth += 1
    return spans


def para_text(seg):
    return "".join(re.findall(r"<hp:t>(.*?)</hp:t>", seg, re.S)).strip()


# 태그 이름 뒤의 \b 가 중요하다. 없으면 "<hp:line" 이 "<hp:lineseg"(빈 문단에도 늘 있다)에
# 걸려 모든 문단이 '내용 있음'으로 잘못 판정된다.
KEEP_TAG_RE = re.compile(
    r"<hp:(tbl|pic|secPr|colPr|container|rect|line|ellipse|polygon|arc|ole|chart|"
    r"equation|textart|connectLine|curve|video|formCtrl)\b")


def is_blank_para(seg):
    if para_text(seg):
        return False
    return not KEEP_TAG_RE.search(seg)


# 머리표 상자로 보는 글자. 신 양식은 '서식 1', '서식2', '참고2' 처럼 띄어쓰기가 들쭉날쭉하다.
HEAD_RE = re.compile(r"^(서식\s*\d+(?:-\d+)?|붙임\s*\d+(?:-\d+)?|참고\s*\d+)")


def norm_label(s):
    """'서식 10' 과 '서식10' 을 같은 것으로 본다."""
    return re.sub(r"\s+", "", s)


def source_xml():
    with zipfile.ZipFile(SRC) as z:
        return z.read("Contents/section0.xml").decode("utf-8")


def restore_first_para_props(xml, src_xml):
    """첫 문단의 문단모양·글자모양을 원본 것으로 되돌린다.

    앞쪽 쪽을 지우면 한글이 살아남은 머리표 문단을 '문서의 첫 문단'(구역 설정을 지고
    있는 문단) 안으로 합쳐 버린다. 이때 문단모양이 첫 문단 것으로 바뀌어 줄 높이가
    늘고, 그만큼 본문 표가 다음 쪽으로 밀린다. 표 id 로 원본의 같은 문단을 찾아
    문단모양과 글자모양을 되돌려 놓는다.
    """
    spans = top_level_paragraphs(xml)
    if not spans:
        return xml
    s, e = spans[0]
    seg = xml[s:e]

    m = re.search(r'<hp:tbl\b[^>]*\bid="(\d+)"', seg)
    if not m:
        return xml
    tid = m.group(1)

    src_seg = None
    for ss, ee in top_level_paragraphs(src_xml):
        if re.search(r'<hp:tbl\b[^>]*\bid="%s"' % tid, src_xml[ss:ee]):
            src_seg = src_xml[ss:ee]
            break
    if src_seg is None:
        return xml

    want_para = re.search(r'paraPrIDRef="(\d+)"', src_seg[:src_seg.find(">") + 1])
    want_char = re.search(r'<hp:run charPrIDRef="(\d+)"', src_seg)
    if not (want_para and want_char):
        return xml

    head_end = seg.find(">") + 1
    head = re.sub(r'paraPrIDRef="\d+"',
                  'paraPrIDRef="%s"' % want_para.group(1), seg[:head_end])
    rest = seg[head_end:]

    # 표 앞에 놓인 run 들 — 구역 설정·단 설정·쪽번호를 지고 있는 빈 run 과
    # 표를 감싸는 run — 은 모두 자기 글자 크기만큼 줄 높이를 차지한다.
    # 한글이 붙여 놓은 엉뚱한 글자모양을 원본 것으로 맞춰 준다.
    tbl_pos = rest.find("<hp:tbl")
    lead, tail = (rest, "") if tbl_pos < 0 else (rest[:tbl_pos], rest[tbl_pos:])
    lead = re.sub(r'charPrIDRef="\d+"',
                  'charPrIDRef="%s"' % want_char.group(1), lead)

    return xml[:s] + head + lead + tail + xml[e:]


def clean_hwpx(path, keep_label, src_xml):
    """다른 서식의 머리표 상자와 잘라낸 자리에 남은 빈 문단을 XML 수준에서 도려낸다.

    keep_label 과 같은 머리표(이 서식 자신의 것)는 남긴다.
    """
    with zipfile.ZipFile(path) as z:
        entries = {n: z.read(n) for n in z.namelist()}
        info = {i.filename: i for i in z.infolist()}

    sec = "Contents/section0.xml"
    xml = entries[sec].decode("utf-8")
    keep = norm_label(keep_label)

    # ① 이웃 서식의 머리표 상자를 표 단위로 도려낸다. 쪽 경계에서 딸려 온 것들이라
    #    남겨 두면 엉뚱한 서식 번호가 함께 찍힌다.
    tbl_drop, dropped_head = [], 0
    for s, e in outer_table_spans(xml):
        txt = para_text(xml[s:e])
        m = HEAD_RE.match(txt)
        if m and len(txt) < 200 and norm_label(m.group(1)) != keep:
            tbl_drop.append((s, e))
            dropped_head += 1

    if tbl_drop:
        tbl_drop.sort()
        out, cursor = [], 0
        for s, e in tbl_drop:
            out.append(xml[cursor:s])
            cursor = e
        out.append(xml[cursor:])
        xml = "".join(out)

    # ② 잘라낸 자리에 남은 빈 껍데기를 더 줄 것이 없을 때까지 걷어낸다.
    dropped_tail = 0
    while True:
        before = xml

        spans = top_level_paragraphs(xml)
        while spans and is_blank_para(xml[spans[-1][0]:spans[-1][1]]):
            s, e = spans[-1]
            xml = xml[:s] + xml[e:]
            dropped_tail += 1
            spans = top_level_paragraphs(xml)

        tspans = outer_table_spans(xml)
        if tspans:
            s, e = tspans[-1]
            if not para_text(xml[s:e]):
                xml = xml[:s] + xml[e:]
                dropped_tail += 1

        # 앞쪽 빈 문단. 첫 문단은 구역 설정(용지·여백)을 지고 있으므로 건드리지 않는다.
        spans = top_level_paragraphs(xml)
        i = 1
        while i < len(spans) and is_blank_para(xml[spans[i][0]:spans[i][1]]):
            s, e = spans[i]
            xml = xml[:s] + xml[e:]
            dropped_tail += 1
            spans = top_level_paragraphs(xml)

        if xml == before:
            break

    # ③ 표를 도려낸 자리에 남은 빈 run 을 걷어낸다. 그냥 두면 그 run 의 글자 크기만큼
    #    빈 줄이 생겨 아래 내용이 다음 쪽으로 밀린다.
    empty_run = re.compile(r"<hp:run\b[^>]*>(?:<hp:t\s*/>|<hp:t></hp:t>)?</hp:run>")
    spans = top_level_paragraphs(xml)
    pieces, cursor = [], 0
    for s, e in spans:
        seg = xml[s:e]
        head_end = seg.find("<hp:tbl")
        head, rest = (seg, "") if head_end < 0 else (seg[:head_end], seg[head_end:])
        pieces.append(xml[cursor:s])
        pieces.append(empty_run.sub("", head) + rest)
        cursor = e
    pieces.append(xml[cursor:])
    xml = "".join(pieces)

    # ④ '원본에서 이 문단이 몇 쪽 어디에 있었는지'가 박힌 줄배치 캐시를 지운다.
    #    남겨 두면 한글이 옛 위치를 그대로 써서 쪽 나눔이 어긋난다.
    xml = re.sub(r"<hp:linesegarray>.*?</hp:linesegarray>", "", xml, flags=re.S)

    # ⑤ 첫 문단의 문단·글자 모양을 원본 것으로 되돌린다 (쪽 밀림 방지)
    xml = restore_first_para_props(xml, src_xml)

    entries[sec] = xml.encode("utf-8")

    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        for name, data in entries.items():
            # mimetype 은 압축하지 않아야 한다(OCF 규약)
            ct = zipfile.ZIP_STORED if name == "mimetype" else zipfile.ZIP_DEFLATED
            zi = zipfile.ZipInfo(name, date_time=info[name].date_time)
            zi.compress_type = ct
            zi.external_attr = info[name].external_attr
            z.writestr(zi, data)

    return dropped_head, dropped_tail


# 첫 문단에 들어앉은 쪽번호 컨트롤. 앞쪽 쪽을 지우면 한글이 이걸 살아남은 문단으로
# 옮겨 넣는데, 빈 run 이지만 한 줄을 차지해 바로 아래 표를 다음 쪽으로 밀어낸다.
# 한 쪽짜리 서식에서는 이 한 줄 때문에 제목만 있는 빈 쪽이 생긴다.
PAGENUM_RUN_RE = re.compile(
    r'<hp:run charPrIDRef="\d+"><hp:ctrl><hp:pageNum[^>]*/></hp:ctrl></hp:run>')


def drop_pagenum_ctrl(path):
    """첫 문단의 쪽번호 컨트롤 run 을 지운다. 지웠으면 True."""
    with zipfile.ZipFile(path) as z:
        entries = {n: z.read(n) for n in z.namelist()}
        info = {i.filename: i for i in z.infolist()}

    sec = "Contents/section0.xml"
    xml = entries[sec].decode("utf-8")
    spans = top_level_paragraphs(xml)
    if not spans:
        return False
    s, e = spans[0]
    seg = xml[s:e]
    new = PAGENUM_RUN_RE.sub("", seg, count=1)
    if new == seg:
        return False

    entries[sec] = (xml[:s] + new + xml[e:]).encode("utf-8")
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        for name, data in entries.items():
            ct = zipfile.ZIP_STORED if name == "mimetype" else zipfile.ZIP_DEFLATED
            zi = zipfile.ZipInfo(name, date_time=info[name].date_time)
            zi.compress_type = ct
            zi.external_attr = info[name].external_attr
            z.writestr(zi, data)
    return True


def markers_in(path):
    with zipfile.ZipFile(path) as z:
        found = set()
        for n in z.namelist():
            if re.search(r"Contents/section\d+\.xml", n):
                xml = z.read(n).decode("utf-8")
                for t in re.findall(r"<hp:t>(.*?)</hp:t>", xml, re.S):
                    t = re.sub(r"<[^>]+>", "", t).strip()
                    if MARKER_RE.match(t):
                        found.add(t)
    return found


# ------------------------------------------------------------------ 본체

def marker_pages(hwp):
    """구간 이름 → 시작 페이지(1-based). 총 페이지 수도 함께 돌려준다."""
    hwp.Open(SRC, "HWPX", "forceopen:true")
    pages = {name: (hwp.KeyIndicator()[3] if find_forward(hwp, needle) else None)
             for name, needle in ALL_MARKERS}
    total = hwp.PageCount
    hwp.XHwpDocuments.Close(isDirty=False)
    return total, pages


def source_ink(hwp, wanted):
    hwp.Open(SRC, "HWPX", "forceopen:true")
    # 문서 끝까지 한 번 움직여 전체 쪽을 잡아두게 한다. 이러지 않으면
    # CreatePageImage 가 아직 배치되지 않은 쪽을 잘린 조각으로 구워 낸다.
    hwp.HAction.Run("MoveDocEnd")
    out = {p: ink_ratio(hwp, p - 1, os.path.dirname(WORK)) for p in sorted(wanted)}
    hwp.XHwpDocuments.Close(isDirty=False)
    return out


def extract(hwp, out_name, first, last, keep_label, src_ink, src_xml):
    out = os.path.join(OUT_DIR, out_name)

    cut_pages(hwp, first, last)
    hwp.SaveAs(out, "HWPX", "")
    hwp.XHwpDocuments.Close(isDirty=False)

    nh, nt = clean_hwpx(out, keep_label, src_xml)

    expect = last - first + 1

    def render():
        hwp.Open(out, "HWPX", "forceopen:true")
        hwp.HAction.Run("MoveDocEnd")
        n = hwp.PageCount
        rs = [ink_ratio(hwp, p, os.path.dirname(WORK)) for p in range(n)]
        hwp.XHwpDocuments.Close(isDirty=False)
        return n, rs

    pages, ratios = render()

    # 쪽이 하나 더 늘었으면 첫 문단의 쪽번호 컨트롤이 한 줄을 먹은 것이다. 지우고 다시 본다.
    dropped_pgnum = False
    if pages > expect and drop_pagenum_ctrl(out):
        dropped_pgnum = True
        pages, ratios = render()

    # 검증 — ① 자기 머리표만 남았는가 ② 쪽 수가 원본 구간과 같은가
    #        ③ 빈 쪽이 없는가 ④ 잉크 총량이 원본과 (지운 상자만큼만) 차이나는가
    left = markers_in(out)
    stray = sorted(m for m in left if norm_label(m) != norm_label(keep_label))
    got = sum(r for r in ratios if r is not None)
    want = sum(src_ink.get(first + i, 0) or 0 for i in range(expect))

    # 머리표는 짙게 칠한 제목 상자라 한 개만 지워도 잉크가 몇 % 씩 준다.
    # 허용치를 지운 상자 수에 비례해 늘린다.
    tolerance = 0.03 + 0.07 * nh

    problems = []
    if stray:
        problems.append(f"남의 머리표 {stray} 남음")
    if norm_label(keep_label) not in {norm_label(m) for m in left}:
        problems.append(f"자기 머리표 '{keep_label}' 가 사라짐")
    if pages != expect:
        problems.append(f"{pages}p (기대 {expect}p)")
    if any(r is not None and r < 0.02 for r in ratios):
        problems.append("거의 빈 쪽 있음")
    if got < want - tolerance:
        problems.append(f"내용 부족 (잉크 {got:.1%} vs 원본 {want:.1%})")

    tidy = f"머리표{nh}/빈문단{nt}" + ("/쪽번호1" if dropped_pgnum else "")
    print(f"  {out_name:<28} p{first}~{last} → {pages}p  "
          f"잉크={' '.join(f'{r:.1%}' for r in ratios if r is not None):<20} "
          f"손질={tidy:<22} "
          f"{'OK' if not problems else '★ ' + ' / '.join(problems)}", flush=True)
    return not problems


def resolve_ranges(total, pages):
    """SPECS 의 (이 구간, 다음 구간) → (파일명, 시작쪽, 끝쪽, 자기 머리표)"""
    out = []
    for name, here, nxt, keep in SPECS:
        a = pages.get(here)
        if a is None:
            print(f"  ! '{here}' 를 찾지 못했습니다 — {name} 건너뜀")
            continue
        b = (pages.get(nxt) - 1) if nxt and pages.get(nxt) else total
        out.append((name, a, b, keep))
    return out


def main():
    hwp = connect()
    try:
        total, pages = marker_pages(hwp)
        print(f"서식모음(new).hwpx: {total} 페이지")
        for name, _ in ALL_MARKERS:
            print(f"  {name:<18} → {pages[name]} 페이지")

        ranges = resolve_ranges(total, pages)
        print("\n잘라낼 구간:")
        for name, a, b, keep in ranges:
            print(f"  {name:<28} p{a}~{b}   (머리표 '{keep}' 유지)")

        if "--dry-run" in sys.argv:
            return 0

        os.makedirs(OUT_DIR, exist_ok=True)
        wanted = {p for _, a, b, _ in ranges for p in range(a, b + 1)}
        print("\n원본 페이지별 잉크 비율을 재는 중...", flush=True)
        src_ink = source_ink(hwp, wanted)

        sx = source_xml()
        print("\n템플릿 생성:")
        results = [extract(hwp, name, a, b, keep, src_ink, sx)
                   for name, a, b, keep in ranges]
        print(f"\n완료: {sum(results)}/{len(results)} 정상")
        if not all(results):
            print("★ 표시된 파일은 원본과 내용이 다릅니다. 한글에서 직접 확인하세요.")
            return 1
    finally:
        try:
            hwp.Quit()
        except Exception:
            pass
        if os.path.exists(WORK):
            os.remove(WORK)
    return 0


if __name__ == "__main__":
    sys.exit(main())
