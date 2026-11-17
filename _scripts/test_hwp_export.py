# -*- coding: utf-8 -*-
"""한글 출력 end-to-end 점검.

버튼 매크로를 실제로 돌린 뒤
  · 출력 폴더에 만들어진 치환본에 {{태그}} 가 남아 있지 않은지
  · 기관명·계약명 같은 실제 값이 들어갔는지
  · 순회점검표가 **달마다 한 장씩** 만들어졌는지 (신 양식은 한 장이 한 달치다)
  · 협의체 회의록이 **월 1회 N회차**로 나오고 회차마다 내용이 다른지
  · 날짜가 2026.3.4.(수) 꼴이고, 순회점검표만 요일이 아랫줄로 내려가는지
  · 서식마다 새 쪽에서 시작하는지
를 확인한다. 통합문서는 저장하지 않는다.

    python _scripts/test_hwp_export.py
"""
import glob
import os
import re
import shutil
import subprocess
import sys
import zipfile

import win32com.client

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from build_workbook import (BOOK, IN_FIRST_ROW, OUT_FIRST_ROW, OUT_LAST_ROW,
                            OC_서식1, NEW_COLS)

sys.stdout.reconfigure(encoding="utf-8")

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
TMP = os.path.join(ROOT, "출력")     # 통합문서 옆의 출력 폴더 (modConfig.TempFolder)

fails = []


def check(label, ok, detail=""):
    print(f"  {'OK ' if ok else '★NG'} {label}{(' — ' + detail) if detail else ''}")
    if not ok:
        fails.append(label)


def merge_info(xl):
    """마지막 병합 결과 (붙인 서식 수, 만들어진 쪽 수, 쪽 나눔 실패 수)"""
    f, p, b = str(xl.Run("GetLastMergeInfo") or "0|0|0").split("|")
    return int(f), int(p), int(b)


def doc_text(path):
    # 이어붙인 문서는 구역이 둘 이상일 수 있으므로 section*.xml 을 모두 읽는다.
    out = ""
    with zipfile.ZipFile(path) as z:
        for name in sorted(n for n in z.namelist()
                           if re.match(r"Contents/section\d+\.xml$", n)):
            xml = z.read(name).decode("utf-8")
            out += "".join(re.findall(r"<hp:t>(.*?)</hp:t>", xml, re.S))
    return re.sub(r"<[^>]+>", "", out)


# 한 줄에 들어가는 자리의 날짜 ― 2026.3.4.(수)  (modHwpExport.FmtDateDow)
DATE_RE = r"20\d\d\.\d{1,2}\.\d{1,2}\.\([월화수목금토일]\)"

# 순회점검표는 칸이 좁아 날짜와 요일을 두 줄로 나눈다 (FmtInspectDate + FmtDow)
DATE_ONLY_RE = r"20\d\d\.\d{1,2}\.\d{1,2}\.(?!\()"
DOW_RE = r"\([월화수목금토일]\)"


def safe(s):
    """modHwpExport.SafeName 과 같은 규칙 — 파일 이름에 못 쓰는 글자를 뺀다."""
    for c in '\\/:*?"<>|·':
        s = s.replace(c, "")
    return s.strip()


def snapshot():
    """출력 폴더의 결과물. '~' 로 시작하는 건 이어붙이기 전의 조각이라 뺀다."""
    return {p for p in glob.glob(os.path.join(TMP, "*.hwpx"))
            if not os.path.basename(p).startswith("~")}


def leftovers():
    """지워졌어야 할 조각(~*.hwpx)"""
    return [os.path.basename(p) for p in glob.glob(os.path.join(TMP, "~*.hwpx"))]


def since(before):
    """직전 스냅샷 이후 새로 생긴 파일. 파일 이름이 계약명·서식명으로 바뀌어
    이름 앞머리로 고를 수 없게 됐다 — 대신 '이번 실행에서 생긴 것' 으로 고른다.
    이름순으로 세우면 '10회차' 가 '1회차' 앞에 서므로 만들어진 차례(mtime)로 세운다."""
    return sorted(snapshot() - before, key=os.path.getmtime)


def inspect(prefix, files, expect_values):
    if not files:
        check(f"{prefix} 치환본 생성", False, "파일이 없습니다")
        return []
    # 값 확인은 파일별이 아니라 전체 합집합으로 한다. 서식마다 있는 칸이 달라서
    # (예: 협의체 회의록에는 계약명 칸이 아예 없다) 파일별로 요구하면 틀린다.
    bad_tag, union = [], ""
    for f in files:
        t = doc_text(f)
        union += t
        if "{{" in t:
            bad_tag.append(os.path.basename(f) + ":"
                           + str(re.findall(r"\{\{[^}]+\}\}", t)))
    missing = [v for v in expect_values if v not in union]
    check(f"{prefix} 치환본 {len(files)}개 생성", True)
    check(f"{prefix} 에 남은 {{{{태그}}}} 없음", not bad_tag, "; ".join(bad_tag[:3]))
    check(f"{prefix} 에 실제 값 들어감", not missing, f"빠진 값 {missing}")
    return files


def hwp_pids():
    """지금 떠 있는 한글 프로세스 번호."""
    try:
        out = subprocess.run(["tasklist", "/FI", "IMAGENAME eq Hwp.exe", "/FO", "CSV"],
                             capture_output=True, text=True, timeout=30).stdout
    except Exception:
        return set()
    return {int(m) for m in re.findall(r'"Hwp\.exe","(\d+)"', out)}


def close_our_hwp(before):
    """이 점검이 띄운 한글만 닫는다. **먼저 떠 있던 것은 건드리지 않는다.**

    버튼 한 번마다 새 한글 인스턴스가 뜨고(HwpConnect), 결과 문서를 보이게
    띄운 채(ShowHwp) 끝난다. 사용자가 쓸 때는 그게 맞지만 점검에서는 열 몇 개가
    쌓이고, 그것들이 출력 폴더의 파일을 붙잡고 있어 **다음 점검이 시작조차
    못 한다.** 실제로 그 때문에 한 번 넘어졌다. 우리가 늘린 것만 골라 닫는다.
    """
    ours = hwp_pids() - before
    for pid in ours:
        try:
            subprocess.run(["taskkill", "/PID", str(pid), "/F"],
                           capture_output=True, timeout=30)
        except Exception:
            pass
    if ours:
        print(f"\n(점검이 띄운 한글 {len(ours)}개를 닫았습니다. "
              f"먼저 열려 있던 것은 그대로 둡니다)")


def stash_previous():
    """출력 폴더를 비우되 **지우지는 않는다.**

    이 점검은 '이번 실행에서 새로 생긴 파일' 로 결과를 고르므로 폴더가 비어
    있어야 한다. 그런데 이 폴더는 임시폴더가 아니라 **사용자가 뽑아 둔 문서가
    쌓이는 곳**이다. 예전에는 통째로 지웠는데, 그러면 점검 한 번에 사용자가
    만들어 둔 서식이 다 날아간다. 옆으로 치워 두고 쓴다.
    """
    if not os.path.isdir(TMP):
        return None
    keep = [p for p in glob.glob(os.path.join(TMP, "*"))
            if not os.path.basename(p).startswith("~")
            and os.path.isfile(p)]
    if not keep:
        for p in glob.glob(os.path.join(TMP, "~*")):
            os.remove(p)
        return None

    box = os.path.join(TMP, "_점검전")
    os.makedirs(box, exist_ok=True)
    moved, locked = 0, []
    for p in keep:
        dst = os.path.join(box, os.path.basename(p))
        try:
            if os.path.exists(dst):
                os.remove(dst)
            shutil.move(p, dst)
            moved += 1
        except PermissionError:
            # 앞선 점검이 띄워 둔 문서를 한글이 아직 붙잡고 있다. 이 상태로
            # 이어 가면 같은 이름의 결과물을 못 써서 엉뚱한 곳에서 넘어진다.
            locked.append(os.path.basename(p))
    for p in glob.glob(os.path.join(TMP, "~*")):
        try:
            os.remove(p)
        except PermissionError:
            pass

    if locked:
        print("★ 한글이 아래 문서를 열어 두고 있어 점검을 시작할 수 없습니다.")
        for n in locked[:5]:
            print("   ·", n)
        print("  한글을 모두 닫은 뒤 다시 돌려 주세요.")
        sys.exit(2)

    print(f"(기존 출력물 {moved}개를 출력\\_점검전\\ 으로 옮겨 두었습니다)")
    return box


def main():
    hwp_before = hwp_pids()      # 우리가 띄운 한글만 뒤에 닫으려고 미리 세어 둔다
    stash_previous()

    xl = win32com.client.Dispatch("Excel.Application")
    xl.Visible = False
    xl.DisplayAlerts = False
    xl.AutomationSecurity = 1
    wb = xl.Workbooks.Open(BOOK)
    try:
        wsi = wb.Worksheets("입력정보")
        wso = wb.Worksheets("출력관리")

        xl.Run("SetSilentMode", True)      # 알림창이 뜨면 무인 실행이 거기서 멈춘다
        org = str(wsi.Cells(5, 2).Value or "")
        print(f"기관명: {org!r}")

        # 2번 계약(1년짜리 용역)을 대상으로 삼는다.
        for r in range(OUT_FIRST_ROW, OUT_LAST_ROW + 1):
            wso.Cells(r, 1).Value = False
        target_in = IN_FIRST_ROW + 1
        target_out = OUT_FIRST_ROW + 1
        wso.Cells(target_out, 1).Value = True

        wsi.Cells(target_in, NEW_COLS["시작일"]).Value = "2026-03-01"
        wsi.Cells(target_in, NEW_COLS["종료일"]).Value = "2027-02-28"   # 365일
        wsi.Cells(target_in, NEW_COLS["점검요일"]).Value = "수요일"
        contract = str(wsi.Cells(target_in, NEW_COLS["계약명"]).Value or "")
        vendor = str(wsi.Cells(target_in, NEW_COLS["업체명"]).Value or "")
        xl.CalculateFull()
        print(f"대상 계약: {contract!r}  (출력관리 {target_out}행)")

        marks = [f for f in range(1, 16)
                 if str(wso.Cells(target_out, OC_서식1 + f - 1).Value or "").strip() == "○"]
        print(f"필요 서식 번호: {marks}")

        want = [v for v in (org, contract, vendor) if v]

        print("\n[1] 1~7서식 한글 출력")
        before = snapshot()
        xl.Run("Btn_ExportHwpForms")
        err = str(xl.Run("GetLastExportError") or "")
        if err:
            print(f"  (매크로 알림) {err}")
        made = inspect("1~7서식", since(before), want)
        # 서식 여러 종이 한 파일로 나온다: 계약명_중대재해예방_계약시작일
        check("이어붙인 결과가 파일 하나인가", len(made) == 1,
              f"{len(made)}개 {[os.path.basename(q) for q in made][:3]}")
        want_name = f"{safe(contract)}_중대재해예방_20260301.hwpx"
        check("파일 이름이 계약명_중대재해예방_시작일 인가",
              bool(made) and os.path.basename(made[0]) == want_name,
              os.path.basename(made[0]) if made else "")
        check("조각 파일이 남지 않았는가", not leftovers(), str(leftovers()[:3]))
        f, p, b = merge_info(xl)
        check("서식마다 새 쪽에서 시작 (쪽 수 ≥ 서식 수)", p >= f, f"서식 {f}개 / {p}쪽")
        check("쪽 나눔이 모두 먹었는가 (앞 서식에 이어 붙지 않음)", b == 0, f"실패 {b}회")
        # 협의체 회의록은 이제 8번이라 이 묶음에서 빠졌다. 회의록에만 있는 글이
        # 여기 섞여 나오면 아직 1~7 안에 들어 있다는 뜻이다.
        check("회의록이 1~7 묶음에서 빠졌는가",
              bool(made) and "안전보건협의체 회의록" not in doc_text(made[0]))

        print("\n[2] 순회점검표 출력 — 달마다 한 장")
        before = snapshot()
        xl.Run("Btn_ExportPatrol")
        dates = xl.Run("InspectionDates", target_in, False)
        months = sorted({(d.year, d.month) for d in dates})
        # 순회점검표에는 계약명·업체명 칸이 없다(기관 전체를 도는 서식이다).
        files = inspect("순회점검표", since(before), [org])
        check("달 몫이 모두 한 파일에 담겼는가", len(files) == 1, f"{len(files)}개")
        want_name = f"{safe(contract)}_순회점검표_총{len(months)}회차_20260301.hwpx"
        check("파일 이름에 총 회차가 들어갔는가",
              bool(files) and os.path.basename(files[0]) == want_name,
              os.path.basename(files[0]) if files else "")
        check("조각 파일이 남지 않았는가", not leftovers(), str(leftovers()[:3]))
        f, p, b = merge_info(xl)
        check("달 수와 붙인 장 수가 같은가", f == len(months),
              f"장 {f} / 달 {len(months)} (점검일 {len(dates)}회)")
        check("달마다 새 쪽에서 시작하는가", f == p, f"{f}장 / {p}쪽")
        check("쪽 나눔이 모두 먹었는가 (앞 장에 이어 붙지 않음)", b == 0, f"실패 {b}회")

        if files:
            t = doc_text(files[0])
            titles = re.findall(r"\[(\d{4})년\s*(\d{1,2})월\]", t)
            check("달마다 제목에 해와 달이 채워졌는가", len(titles) == len(months),
                  f"제목 {len(titles)}개 / 달 {len(months)}개")
            days = re.findall(DATE_ONLY_RE, t)
            check("점검일이 모두 1주~5주 칸에 들어갔는가", len(days) == len(dates),
                  f"날짜 {len(days)}개 / 점검일 {len(dates)}회")
            # 요일은 날짜 **아랫줄**에 따로 들어간다. 날짜 뒤에 '(수)' 가 바로
            # 붙어 있으면 옛 한 줄짜리로 되돌아간 것이다.
            dows = re.findall(DOW_RE, t)
            check("요일이 날짜 수만큼 따로 들어갔는가", len(dows) == len(days),
                  f"요일 {len(dows)}개 / 날짜 {len(days)}개")
            check("날짜 뒤에 요일이 바로 붙지 않았는가",
                  not re.search(r"20\d\d\.\d{1,2}\.\d{1,2}\.\(", t))
            # 공휴일로 민 날은 점검요일과 다른 요일로 찍힌다 — 별표는 이제 안 쓴다.
            moved = [d for d in dates if d.weekday() != 2]
            odd = [w for w in dows if w != "(수)"]
            check("공휴일로 민 날이 다른 요일로 보이는가", len(odd) == len(moved),
                  f"수요일 아닌 요일 {len(odd)}개 / 미뤄진 점검일 {len(moved)}회")
            check("별표가 남아 있지 않은가", "* 20" not in t)

        print("\n[3] 합동점검표 출력 — 분기 1회")
        before = snapshot()
        xl.Run("Btn_ExportJoint")
        joint = xl.Run("InspectionDates", target_in, True)
        # 합동점검표에는 계약명 칸이 없다(참여 업체를 일괄로 적는 서식이다).
        files = inspect("합동점검표", since(before), [org, vendor])
        check("회차 몫이 모두 한 파일에 담겼는가", len(files) == 1, f"{len(files)}개")
        want_name = f"{safe(contract)}_합동점검표_총{len(joint)}회차_20260301.hwpx"
        check("파일 이름에 총 회차가 들어갔는가",
              bool(files) and os.path.basename(files[0]) == want_name,
              os.path.basename(files[0]) if files else "")
        f, p, b = merge_info(xl)
        check("합동 회차 수와 붙인 장 수가 같은가", f == len(joint),
              f"장 {f} / 회차 {len(joint)}")
        check("합동도 회차마다 새 쪽에서 시작하는가", f == p, f"회차 {f} / {p}쪽")
        check("쪽 나눔이 모두 먹었는가", b == 0, f"실패 {b}회")

        if files:
            t = doc_text(files[0])
            found = re.findall(DATE_RE, t)
            check("회차마다 점검일자가 채워졌는가", len(found) >= len(joint),
                  f"날짜 {len(found)}개 / 회차 {len(joint)}")

        print("\n[3-1] 협의체 회의록 출력 — 월 1회 N회차")
        before = snapshot()
        xl.Run("Btn_ExportCouncil")
        council = xl.Run("CouncilDates", target_in)
        # 회의록에는 계약명 칸이 없다(그 자리에서 협의한 것을 적는 서식이다).
        files = inspect("협의체 회의록", since(before), [org, vendor])
        check("회차 몫이 모두 한 파일에 담겼는가", len(files) == 1, f"{len(files)}개")
        want_name = (f"{safe(contract)}_안전보건협의체 회의록"
                     f"_총{len(council)}회차_20260301.hwpx")
        check("파일 이름이 계약명_안전보건협의체 회의록_총N회차_시작일 인가",
              bool(files) and os.path.basename(files[0]) == want_name,
              os.path.basename(files[0]) if files else "")
        check("조각 파일이 남지 않았는가", not leftovers(), str(leftovers()[:3]))
        f, p, b = merge_info(xl)
        check("달 수와 붙인 장 수가 같은가", f == len(council),
              f"장 {f} / 회차 {len(council)}")
        check("회차마다 새 쪽 맨 위에서 시작하는가", f == p, f"{f}장 / {p}쪽")
        check("쪽 나눔이 모두 먹었는가", b == 0, f"실패 {b}회")

        if files:
            t = doc_text(files[0])
            rounds = re.findall(r"(\d+)회차", t)
            check("회차가 1..N 으로 찍혔는가",
                  rounds == [str(i + 1) for i in range(len(council))],
                  f"찍힌 회차 {rounds[:14]}")
            found = re.findall(DATE_RE, t)
            check("회의일이 회차마다 채워졌는가", len(found) >= len(council),
                  f"날짜 {len(found)}개 / 회차 {len(council)}")

            # 열두 장이 복사본이면 회의록으로서 뜻이 없다. 회차마다 안건·의결이
            # 달라야 한다 — 쪽을 '회차' 표시로 잘라 서로 견준다.
            parts = re.split(r"\d+회차", t)[1:]
            uniq = len(set(p.strip() for p in parts))
            check("회차마다 본문이 서로 다른가", uniq == len(parts),
                  f"서로 다른 쪽 {uniq}개 / 모두 {len(parts)}쪽")

        print("\n[4] 11~15 서식 각각")
        head_title = str(wsi.Cells(6, 9).Value or "")     # 결재선 부서장 직위
        head_name = str(wsi.Cells(6, 11).Value or "")     # 결재선 부서장 성명
        for no, macro in ((11, "Btn_Export11"), (12, "Btn_Export12"),
                          (13, "Btn_Export13"), (14, "Btn_Export14"),
                          (15, "Btn_Export15")):
            before = snapshot()
            xl.Run(macro)
            got = since(before)
            ok = len(got) == 1 and "{{" not in doc_text(got[0])
            check(f"서식 {no} 출력", ok,
                  f"{len(got)}개" if len(got) != 1
                  else "남은 태그 있음" if not ok else os.path.basename(got[0]))
            # 관리감독자는 부서장(학교면 교장)이다. 관리자(행정실장)가 아니다.
            if no == 12 and len(got) == 1:
                t = doc_text(got[0])
                check("서식 12 의 관리감독자가 부서장인가",
                      bool(head_title) and head_title in t
                      and (not head_name or head_name in t),
                      f"부서장 {head_title!r} {head_name!r}")

        print("\n[5] 계약 2건을 고르면 둘 다 나오는가 (한 문서로 병합)")
        # 예전에는 계약마다 hwp.Open 을 다시 불러 앞 문서를 덮어써서
        # **마지막 계약만** 남았다. 이제 경로를 다 모아 한 번에 이어붙인다.
        second_out, second_in = OUT_FIRST_ROW + 2, IN_FIRST_ROW + 2
        wso.Cells(second_out, 1).Value = True          # 1번 계약은 아직 켜져 있다
        xl.CalculateFull()
        n1 = len([f for f in range(1, 8)
                  if str(wso.Cells(target_out, OC_서식1 + f - 1).Value or "").strip() == "○"])
        n2 = len([f for f in range(1, 8)
                  if str(wso.Cells(second_out, OC_서식1 + f - 1).Value or "").strip() == "○"])
        before = snapshot()
        xl.Run("Btn_ExportHwpForms")
        f, p, b = merge_info(xl)
        check("두 계약의 서식이 모두 한 문서에 들어갔는가", f == n1 + n2,
              f"이어붙인 서식 {f}개 (기대 {n1}+{n2}={n1 + n2})")
        check("계약 사이에도 쪽이 넘어가는가", p >= f, f"{f}개 서식 / {p}쪽")
        check("쪽 나눔이 모두 먹었는가", b == 0, f"실패 {b}회")
        # 두 계약이 한 문서로 붙으므로 이름 뒤에 "외 1건" 이 붙는다.
        made2 = since(before)
        check("두 계약 몫이 파일 하나로 나왔는가", len(made2) == 1,
              f"{len(made2)}개 {[os.path.basename(q) for q in made2][:3]}")
        check("이름에 '외 1건' 이 붙었는가",
              bool(made2) and "_외1건_" in os.path.basename(made2[0]),
              os.path.basename(made2[0]) if made2 else "")

        print("\n[6] 기준 미달 계약은 만들지 않는가 (90일 미만 → 합동 없음)")
        short_out = OUT_FIRST_ROW + 2                               # 30일 계약
        wso.Cells(target_out, 1).Value = False
        wso.Cells(short_out, 1).Value = True
        xl.CalculateFull()
        before = snapshot()
        xl.Run("Btn_ExportJoint")
        after = len(since(before))
        check("30일 계약에는 합동점검표를 만들지 않는다", after == 0, f"{after}개 생성됨")
        print(f"  (매크로 안내) {xl.Run('GetLastExportError')}")

        print("\n[7] 템플릿 원본이 그대로인가")
        broken = [os.path.basename(p)
                  for p in glob.glob(os.path.join(ROOT, "form", "templates", "*.hwpx"))
                  if "{{" not in doc_text(p)]
        check("템플릿 17종 모두 {{태그}} 그대로", not broken, str(broken))
    finally:
        wb.Close(SaveChanges=False)
        xl.Quit()
        close_our_hwp(hwp_before)

    print("\n" + ("모두 통과" if not fails else f"실패 {len(fails)}건: {fails}"))
    return 1 if fails else 0


if __name__ == "__main__":
    sys.exit(main())
