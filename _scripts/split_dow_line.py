# -*- coding: utf-8 -*-
"""서식 두 곳을 그 칸만 고친다. 한 번 돌리면 그만인 스크립트다.

  ① 09_순회점검표 ― '점검일' 칸의 요일을 **날짜 아랫줄**로 내린다.
     지금은 {{점검일1}} 하나에 '2026.6.4.(목)' 이 통째로 들어가는데, 칸이 17mm
     밖에 안 되어 '2026.6.4.(' 까지만 들어가고 '목)' 만 다음 줄로 넘어간다.
     여는 괄호가 날짜 줄에 매달려 보기 흉하다. 문단을 하나 더 만들어
     {{점검요일1}} 을 심고, 날짜와 요일을 따로 넣는다.

  ② 12_관리감독자평가표 ― 관리감독자는 **부서장**이다(학교면 교장).
     {{관리자직위}}/{{관리자성명}}(행정실장) 이 들어가 있던 것을
     {{부서장직위}}/{{부서장성명}} 으로 바꾼다.

**inject_tags_v2.py 를 다시 돌리면 안 되므로** 이렇게 그 칸만 건드린다.
사용자가 한글에서 손본 나머지는 그대로 둔다. 다시 돌려도 안전하다(이미 되어
있으면 건너뛴다).
"""
import os
import re
import shutil
import sys
import zipfile

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
TPL = os.path.join(ROOT, "form", "templates")

SEC = "Contents/section0.xml"


def read_section(path):
    with zipfile.ZipFile(path) as z:
        names = z.namelist()
        return {n: z.read(n) for n in names}, names


def write_zip(path, blobs, names):
    """hwpx 는 mimetype 이 **맨 앞에 무압축**으로 있어야 한다."""
    tmp = path + ".tmp"
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as z:
        if "mimetype" in names:
            z.writestr(zipfile.ZipInfo("mimetype"), blobs["mimetype"],
                       compress_type=zipfile.ZIP_STORED)
        for n in names:
            if n == "mimetype":
                continue
            z.writestr(n, blobs[n])
    shutil.move(tmp, path)


# ---------------------------------------------------------------- ① 순회점검표

def patrol_dow_line():
    path = os.path.join(TPL, "09_순회점검표.hwpx")
    blobs, names = read_section(path)
    xml = blobs[SEC].decode("utf-8")

    if "{{점검요일1}}" in xml:
        print("  · 09_순회점검표 ― 이미 되어 있음")
        return False

    done = 0
    for k in range(1, 6):
        tag = "{{점검일%d}}" % k
        # 그 태그가 든 문단 하나를 통째로 잡는다
        m = re.search(r"<hp:p\b(?:(?!</hp:p>).)*?"
                      + re.escape(tag)
                      + r"(?:(?!</hp:p>).)*?</hp:p>", xml, re.S)
        if not m:
            print(f"  ! {tag} 문단을 찾지 못했습니다")
            continue
        para = m.group(0)

        # 날짜 문단을 복제해 요일 문단을 만든다. 글자 모양(charPrIDRef)과
        # 문단 모양(paraPrIDRef)을 그대로 물려받아야 두 줄이 한 덩어리로 보인다.
        dow = para.replace(tag, "{{점검요일%d}}" % k)
        # 문단 id 는 겹치면 안 된다.
        dow = re.sub(r'<hp:p id="\d+"', '<hp:p id="%d"' % (91000 + k), dow, count=1)
        # 줄 위치 정보는 한글이 파일을 열 때 다시 잰다. 복제본의 낡은 값을
        # 그대로 두면 두 줄이 같은 자리에 겹쳐 그려진다. 통째로 뗀다.
        dow = re.sub(r"<hp:linesegarray>.*?</hp:linesegarray>", "", dow, flags=re.S)

        xml = xml[:m.end()] + dow + xml[m.end():]
        done += 1

    if not done:
        return False
    blobs[SEC] = xml.encode("utf-8")
    write_zip(path, blobs, names)
    print(f"  + 09_순회점검표 ― 요일 줄 {done}칸에 넣음")
    return True


# ------------------------------------------------------- ② 관리감독자 평가표

def supervisor_is_head():
    path = os.path.join(TPL, "12_관리감독자평가표.hwpx")
    blobs, names = read_section(path)
    xml = blobs[SEC].decode("utf-8")

    if "{{관리자직위}}" not in xml and "{{관리자성명}}" not in xml:
        print("  · 12_관리감독자평가표 ― 이미 되어 있음")
        return False

    xml = xml.replace("{{관리자직위}}", "{{부서장직위}}")
    xml = xml.replace("{{관리자성명}}", "{{부서장성명}}")
    blobs[SEC] = xml.encode("utf-8")
    write_zip(path, blobs, names)
    print("  + 12_관리감독자평가표 ― 관리감독자를 부서장(학교면 교장)으로")
    return True


def main():
    sys.stdout.reconfigure(encoding="utf-8")
    patrol_dow_line()
    supervisor_is_head()
    print("\n끝났습니다. gen_template_tags.py 를 이어서 돌리세요.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
