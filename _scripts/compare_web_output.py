# -*- coding: utf-8 -*-
"""엑셀(한글 COM) 출력과 웹 출력의 글자를 맞대어 본다.

    python _scripts/compare_web_output.py <엑셀출력.hwpx> <웹출력.hwpx>

문단마다 글자만 뽑아 비교한다(글자 모양·배치는 보지 않는다). 엑셀판은 InsertFile 로
한 구역에 이어붙이고, 웹판은 구역을 나누므로 구역 경계는 무시한다.
"""
import difflib
import re
import sys
import zipfile

sys.stdout.reconfigure(encoding="utf-8")


def paras(path):
    out = []
    with zipfile.ZipFile(path) as z:
        secs = sorted((n for n in z.namelist() if re.match(r"Contents/section\d+\.xml$", n)),
                      key=lambda n: int(re.search(r"\d+", n).group()))
        for s in secs:
            x = z.read(s).decode("utf-8")
            for p in re.findall(r"<hp:p\b.*?</hp:p>", x, re.S):
                ts = re.findall(r"<hp:t(?:\s[^>]*)?>(.*?)</hp:t>", p, re.S)
                txt = "".join(re.sub(r"<[^>]+>", "", t) for t in ts).strip()
                if txt:
                    out.append(txt)
    return out


def main(a, b):
    pa, pb = paras(a), paras(b)
    diff = [l for l in difflib.unified_diff(pa, pb, "excel", "web", lineterm="", n=0)
            if not l.startswith("@@")]
    left = [l for l in pa if "{{" in l] + [l for l in pb if "{{" in l]
    print(f"문단 {len(pa)} / {len(pb)} · 다른 줄 {max(0, len(diff) - 2)} · 남은 태그 {len(left)}")
    for l in diff[2:60]:
        print("  ", l[:160])
    return 0 if len(diff) == 0 and not left else 1


if __name__ == "__main__":
    sys.exit(main(sys.argv[1], sys.argv[2]))
