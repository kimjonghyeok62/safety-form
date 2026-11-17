# -*- coding: utf-8 -*-
"""통합문서 동작 점검 — VBA 컴파일, 계약유형 매칭, 판정식, 점검일 생성까지 확인한다.

한글을 띄우는 출력 버튼은 여기서 부르지 않는다(test_hwp_export.py 에서 확인).

    python _scripts/test_workbook.py
"""
import datetime
import os
import sys

import win32com.client

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from build_workbook import (BOOK, IN_SCHOOL_ROW, IN_SIGN_ROW, IN_FIRST_ROW,
                            IN_LAST_ROW, IN_LAST_COL, OUT_FIRST_ROW, OUT_LAST_ROW,
                            OC_계약기간, OC_점검요일, OC_서식1, OC_요약, NEW_COLS,
                            SCHOOL_FIELDS, SIGN_FIELDS, AUTO_YELLOW,
                            DONE_FILL, PROGRESS_FILL,
                            DICT_HEAD_ROW, DICT_FIRST_ROW, DICT_LAST_ROW,
                            DICT_HEADERS)
from contract_types import CONTRACT_TYPES

sys.stdout.reconfigure(encoding="utf-8")

fails = []

CC = NEW_COLS
WORK_COLS = [("고소작업", 14), ("화기작업", 15), ("밀폐공간작업", 16), ("전기작업", 17)]


def check(label, got, want):
    ok = got == want
    print(f"  {'OK ' if ok else '★NG'} {label}: {got!r}" + ("" if ok else f" (기대 {want!r})"))
    if not ok:
        fails.append(label)
    return ok


def marks(ws, row):
    return [nm for nm, c in WORK_COLS
            if str(ws.Cells(row, c).Value or "").strip() == "○"]


def main():
    xl = win32com.client.Dispatch("Excel.Application")
    xl.Visible = False
    xl.DisplayAlerts = False
    xl.AutomationSecurity = 1              # 매크로 사용
    wb = xl.Workbooks.Open(BOOK)
    try:
        wsi = wb.Worksheets("입력정보")
        wso = wb.Worksheets("출력관리")

        print("[1] VBA 컴파일 / 매크로 실행")
        # 아무 매크로나 한 번 돌면 프로젝트 전체가 컴파일된 것이다. 모듈 하나가
        # 깨져 있으면 여기서 '매크로를 실행할 수 없습니다' 로 걸린다.
        # 시트 값을 바꾸지 않는 것으로 고른다(체크박스 모양만 셀 값에 맞춘다).
        try:
            xl.Run("SyncCheckboxes")
            xl.Run("RefreshOutputWeekdays")
            print("  OK  매크로 실행됨 (컴파일 통과)")
        except Exception as e:
            print(f"  ★NG 매크로 실행 실패: {e}")
            fails.append("VBA 컴파일")
            return 1

        print("\n[2] 시트 구성")
        names = [s.Name for s in wb.Worksheets]
        check("보이는 시트", sorted(n for n in names if n != "ContentLibrary"),
              sorted(["입력정보", "출력관리", "설정", "계약유형사전"]))
        check("ContentLibrary 는 숨김", wb.Worksheets("ContentLibrary").Visible, 2)

        print("\n[3] 기관구분에 따른 결재선 직위 기본값")
        kind_col = SCHOOL_FIELDS[0][1]
        pos_cols = [SIGN_FIELDS[i][1] for i in (0, 3, 5, 7)]   # 담당/관리/부서장/대표자 직위
        # 성명·전화는 기관구분을 바꾸면 비워졌다가 되돌아온다. 검사가 사용자 값을
        # 건드리므로 직위뿐 아니라 이 칸들도 함께 저장해 두었다가 되돌린다.
        name_cols = [SIGN_FIELDS[i][1] for i in (1, 2, 4, 6, 8)]
        saved_kind = wsi.Cells(IN_SCHOOL_ROW, kind_col).Value
        saved_pos = [wsi.Cells(IN_SIGN_ROW, c).Value for c in pos_cols]
        saved_name = [wsi.Cells(IN_SIGN_ROW, c).Value for c in name_cols]

        for kind, want in (("교육지원청", ["주무관", "팀장", "과장", "교육장"]),
                           ("학교", ["주무관", "행정실장", "교장", "교장"])):
            for c in pos_cols:
                wsi.Cells(IN_SIGN_ROW, c).Value = ""
            wsi.Cells(IN_SCHOOL_ROW, kind_col).Value = kind
            xl.Run("ApplyOrgDefaults")
            got = [str(wsi.Cells(IN_SIGN_ROW, c).Value or "") for c in pos_cols]
            check(f"{kind} 기본 직위", got, want)

        # 잘못 골랐다가 되돌렸을 때 성명·전화가 돌아오는가 (지금 기관구분은 '학교')
        probe = ["홍길동", "031-000-0000", "김관리", "박부서", "박부서"]
        for c, v in zip(name_cols, probe):
            wsi.Cells(IN_SIGN_ROW, c).Value = v
        wsi.Cells(IN_SCHOOL_ROW, kind_col).Value = "교육지원청"
        xl.Run("ApplyOrgDefaults")
        wsi.Cells(IN_SCHOOL_ROW, kind_col).Value = "학교"
        xl.Run("ApplyOrgDefaults")
        got = [str(wsi.Cells(IN_SIGN_ROW, c).Value or "") for c in name_cols]
        check("학교 → 교육지원청 → 학교 성명·전화 복원", got, probe)

        # 원래 값으로 되돌린다
        wsi.Cells(IN_SCHOOL_ROW, kind_col).Value = saved_kind or "학교"
        for c, v in zip(pos_cols + name_cols, saved_pos + saved_name):
            wsi.Cells(IN_SIGN_ROW, c).Value = v

        print("\n[4] 테스트 계약 입력")
        rows = [
            ("옥상 방수공사",           "2026-05-01", "2026-05-01"),   # 1일
            ("급식실 후드 교체공사",     "2026-05-01", "2026-05-10"),   # 10일
            ("물탱크 청소 용역",         "2026-03-01", "2026-06-08"),   # 100일
            ("수배전반 정기점검 용역",   "2026-03-01", "2027-02-28"),   # 365일
        ]
        for r in range(IN_FIRST_ROW, IN_LAST_ROW + 1):
            wsi.Range(wsi.Cells(r, 2), wsi.Cells(r, 20)).ClearContents()
            wsi.Range(wsi.Cells(r, 2), wsi.Cells(r, 20)).Interior.ColorIndex = -4142
        for i, (nm, s, e) in enumerate(rows):
            r = IN_FIRST_ROW + i
            wsi.Cells(r, CC["계약명"]).Value = nm
            wsi.Cells(r, CC["시작일"]).Value = s
            wsi.Cells(r, CC["종료일"]).Value = e
        xl.Run("RefreshAllWorkTypesSilent")
        xl.CalculateFull()

        print("\n[5] 계약유형 사전 매칭 (계약명 → 유형)")
        want_types = ["옥상방수공사", "급식실 후드·덕트 청소",
                      "물탱크·저수조 청소", "수배전반·변압기 정비"]
        for i, want in enumerate(want_types):
            r = IN_FIRST_ROW + i
            check(rows[i][0], str(wsi.Cells(r, CC["계약유형"]).Value or ""), want)

        print("\n[6] 작업유형 자동 판정 (사전의 기본 작업유형)")
        want_work = [["고소작업", "화기작업"],           # 옥상방수 — 토치·용제
                     ["고소작업", "화기작업", "밀폐공간작업"],  # 급식실 후드·덕트
                     ["밀폐공간작업"],                    # 물탱크
                     ["전기작업"]]                        # 수배전반
        for i, want in enumerate(want_work):
            check(rows[i][0], sorted(marks(wsi, IN_FIRST_ROW + i)), sorted(want))

        print("\n[7] ○ / × 드롭다운")
        cel = wsi.Cells(IN_FIRST_ROW, 14)
        check("고소작업 칸 유효성 목록", str(cel.Validation.Formula1), "○,×")
        check("설명 말풍선 켜짐", bool(cel.Validation.ShowInput), True)
        check("자동 판정 결과에 × 가 찍히는가",
              str(wsi.Cells(IN_FIRST_ROW + 3, 14).Value or ""), "×")

        print("\n[8] 서식 필요 판정 (출력관리)")
        def mark(i, form_no):
            return str(wso.Cells(OUT_FIRST_ROW + i, OC_서식1 + form_no - 1).Value or "")
        # 번호는 화면 번호다. 협의체 회의록이 3번에서 8번으로 옮겨 가면서
        # 도급점검표가 3번, 작업유형 점검표가 4~7번이 되었다.
        check("1일 계약의 3.도급(모든 계약)", mark(0, 3), "○")
        check("1일 계약의 8.협의체(30일↑)", mark(0, 8), "")
        check("100일 계약의 8.협의체(30일↑)", mark(2, 8), "○")
        check("10일 계약의 9.순회(7일↑)", mark(1, 9), "○")
        check("10일 계약의 10.합동(90일↑)", mark(1, 10), "")
        check("100일 계약의 10.합동(90일↑)", mark(2, 10), "○")
        check("11~15 는 자동 판정하지 않는다", mark(3, 11), "")

        print("\n[9] 행 정렬 — 출력관리 계약명이 입력정보와 같은 줄인가")
        for i in range(4):
            check(f"출력관리 {OUT_FIRST_ROW+i}행 ↔ 입력정보 {IN_FIRST_ROW+i}행",
                  str(wso.Cells(OUT_FIRST_ROW + i, 3).Value),
                  str(wsi.Cells(IN_FIRST_ROW + i, 2).Value))

        print("\n[10] 점검요일 — 출력관리에서 바꾸면 입력정보로 되돌아가는가")
        wso.Cells(OUT_FIRST_ROW, OC_점검요일).Value = "금요일"
        check("입력정보에 반영", str(wsi.Cells(IN_FIRST_ROW, CC["점검요일"]).Value or ""),
              "금요일")
        wso.Cells(OUT_FIRST_ROW, OC_점검요일).Value = "수요일"

        print("\n[10-1] 계약기간 — 일수로 보여 주는가")
        xl.CalculateFull()
        # 계약을 몇 건 넣어 두었는지는 [4] 가 정한다. 첫 줄의 시작·종료로 직접 센다.
        s = wsi.Cells(IN_FIRST_ROW, CC["시작일"]).Value
        e = wsi.Cells(IN_FIRST_ROW, CC["종료일"]).Value
        want = f"{(e - s).days + 1:,}일"
        check("첫 계약의 계약기간", str(wso.Cells(OUT_FIRST_ROW, OC_계약기간).Value or ""), want)
        # 날짜가 비면 '1일' 이 아니라 빈칸이어야 한다 (빈 칸은 0 이라 그냥 빼면 1 이 된다)
        blank = OUT_LAST_ROW
        check("계약이 없는 줄은 빈칸",
              str(wso.Cells(blank, OC_계약기간).Value or ""), "")

        print("\n[11] 수동 수정 보존")
        r = IN_FIRST_ROW + 2                                  # 물탱크 청소 → 밀폐 ○
        wsi.Cells(r, 16).Value = "×"                          # 사람이 끔 → 회색(수동)
        wsi.Cells(r, 17).Value = "○"                          # 사람이 전기작업을 켬
        wsi.Cells(r, 2).Value = "물탱크 청소 용역 (재계약)"     # 재판정 유발
        xl.CalculateFull()
        check("손으로 끈 밀폐 칸이 그대로 × 인가", str(wsi.Cells(r, 16).Value or ""), "×")
        check("손으로 켠 전기 칸이 그대로 ○ 인가", str(wsi.Cells(r, 17).Value or ""), "○")
        check("자동으로 채운 칸은 계속 자동(노랑)인가",
              wsi.Cells(r, 14).Interior.Color, AUTO_YELLOW)

        print("\n[12] 점검일 생성 (순회=달마다 / 협의체=월 1회 / 합동=분기 1회)")
        r = IN_FIRST_ROW + 3                    # 365일 계약
        wsi.Cells(r, CC["점검요일"]).Value = "수요일"
        patrol = xl.Run("InspectionDates", r, False)
        joint = xl.Run("InspectionDates", r, True)
        council = xl.Run("CouncilDates", r)
        months = sorted({(d.year, d.month) for d in patrol})
        print(f"  점검일 {len(patrol)}회  처음={patrol[0]}  마지막={patrol[-1]}")
        check("1년 계약의 점검일(52~53회)", 52 <= len(patrol) <= 53, True)
        check("순회점검표 장수 = 달 수(12장)", len(months), 12)
        # 합동은 시작일부터 90일씩 끊은 구간마다 1회다. 365일이면 구간이 5개
        # (마지막 구간은 5일뿐이라 그 안에 점검요일이 없으면 4회가 된다).
        check("합동점검표 회차(분기 1회 → 4~5회)", 4 <= len(joint) <= 5, True)
        check("첫 점검일이 수요일", patrol[0].weekday(), 2)

        # 협의체는 달마다 마지막 순회점검일 하나다.
        check("협의체 회차 = 달 수(12회)", len(council), len(months))
        check("협의체 회의일이 모두 순회점검일 안에 있는가",
              all(d in patrol for d in council), True)
        want_last = [max(d for d in patrol if (d.year, d.month) == ym) for ym in months]
        check("달마다 마지막 점검일을 골랐는가", list(council), want_last)

        # 공휴일에 걸린 점검일은 다음 평일로 밀린다 → 요일이 수요일이 아니다.
        # 서식에 요일이 함께 찍히므로 그것만으로 밀린 날임을 알 수 있다.
        moved = [d for d in patrol if d.weekday() != 2]
        print(f"  공휴일로 미뤄진 점검일 {len(moved)}회 "
              f"{[str(d.date()) for d in moved[:4]]}")
        check("미뤄진 점검일이 하나라도 있는가 (요일 표기 검사용)", len(moved) > 0, True)

        print("\n[13] 계약유형 사전 — 열 구성과 문구")
        wsd = wb.Worksheets("계약유형사전")
        head = [str(wsd.Cells(DICT_HEAD_ROW, c).Value or "")
                for c in range(1, len(DICT_HEADERS) + 1)]
        check("사전 열", head, DICT_HEADERS)
        check("사전 줄 수", DICT_LAST_ROW - DICT_FIRST_ROW + 1, len(CONTRACT_TYPES))
        # 계획서 상세는 위험요인·안전대책을 네 줄까지 쓴다. 사전이 그만큼 주어야 한다.
        thin = [str(wsd.Cells(r, 4).Value or "")
                for r in range(DICT_FIRST_ROW, DICT_LAST_ROW + 1)
                if len(str(wsd.Cells(r, 4).Value or "").split(" · ")) < 4]
        check("위험요인이 네 항목 미만인 줄", len(thin), 0)

        print("\n[14] 줄 번호 태그 (한 줄에 한 항목씩)")
        r = IN_FIRST_ROW + 2                      # 물탱크 청소 → 밀폐공간
        tags = str(xl.Run("GetTagList", r)).split("|")
        want_tags = ["{{위험요인1}}", "{{위험요인5}}", "{{안전대책1}}", "{{안전대책5}}",
                     "{{작업내용1}}", "{{협의체안건1}}", "{{점검착안사항1}}",
                     "{{사용기계기구1}}", "{{사용기계기구3}}",
                     "{{위험물질설비1}}", "{{위험물질설비3}}",
                     "{{작업절차}}", "{{안전점검계획}}", "{{보호구방호장치}}",
                     "{{안전보건교육}}", "{{대피방법}}"]
        check("빠진 태그", [t for t in want_tags if t not in tags], [])
        # 한 항목이 통째로 들어가야 한다. 가운뎃점으로 이은 덩어리가 오면 안 된다.
        one = str(xl.Run("GetTagValue", r, "{{위험요인1}}"))
        check("{{위험요인1}} 이 항목 하나인가", " · " not in one and one != "", True)
        print(f"       {{{{위험요인1}}}} = {one!r}")

        print("\n[15] 진행 / 완료 자동 판정")
        ws = wb.Worksheets("입력정보")
        today = datetime.date.today()
        row = IN_FIRST_ROW + 8                    # 테스트 계약이 없는 빈 줄

        def remark(r):
            return str(ws.Cells(r, CC["비고"]).Value or "")

        xl.EnableEvents = True
        ws.Cells(row, CC["계약명"]).Value = "진행완료 판정 시험 계약"
        ws.Cells(row, CC["시작일"]).Value = str(today - datetime.timedelta(days=30))
        ws.Cells(row, CC["종료일"]).Value = str(today + datetime.timedelta(days=30))
        check("종료일이 남았으면 진행", remark(row), "진행")

        ws.Cells(row, CC["종료일"]).Value = str(today - datetime.timedelta(days=1))
        check("종료일이 지났으면 완료", remark(row), "완료")

        # 사용자가 손으로 고른 값은 날짜보다 우선한다
        ws.Cells(row, CC["비고"]).Value = "진행"
        ws.Cells(row, CC["종료일"]).Value = str(today - datetime.timedelta(days=2))
        check("손으로 고른 값이 날짜를 이긴다", remark(row), "진행")
        check("수동 표시가 V열에 남는가",
              str(ws.Cells(row, CC["비고수동"]).Value or ""), "수동")

        # 비우면 다시 날짜 기준으로 돌아간다
        ws.Cells(row, CC["비고"]).ClearContents()
        check("비우면 날짜 판정으로 복귀", remark(row), "완료")
        check("수동 표시도 풀리는가",
              str(ws.Cells(row, CC["비고수동"]).Value or ""), "")

        # 계약을 지우면 비고도 함께 사라진다 (빈 줄에 색만 남으면 안 된다)
        ws.Cells(row, CC["계약명"]).ClearContents()
        ws.Cells(row, CC["시작일"]).ClearContents()
        ws.Cells(row, CC["종료일"]).ClearContents()
        check("계약을 지우면 비고도 빈다", remark(row), "")
        xl.EnableEvents = False

        print("\n[16] 진행 / 완료 조건부서식")
        cf = ws.Range(ws.Cells(IN_FIRST_ROW, 1),
                      ws.Cells(IN_LAST_ROW, IN_LAST_COL)).FormatConditions
        check("입력정보 규칙 수", cf.Count, 2)
        check("완료 규칙", (cf(1).Formula1, int(cf(1).Interior.Color)),
              (f'=$T{IN_FIRST_ROW}="완료"', DONE_FILL))
        check("진행 규칙", (cf(2).Formula1, int(cf(2).Interior.Color)),
              (f'=$T{IN_FIRST_ROW}="진행"', PROGRESS_FILL))

        wso = wb.Worksheets("출력관리")
        cfo = wso.Range(wso.Cells(OUT_FIRST_ROW, 1),
                        wso.Cells(OUT_LAST_ROW, OC_요약)).FormatConditions
        check("출력관리 규칙 수", cfo.Count, 2)
        check("출력관리가 입력정보 비고를 본다", cfo(1).Formula1,
              f'=입력정보!$T{IN_FIRST_ROW}="완료"')

        print("\n[17] 5초 강조가 진행/완료 색 위에 보이는가")
        # Interior.Color 로 칠하면 조건부서식에 묻힌다. 강조도 조건부서식이어야 한다.
        target = OUT_FIRST_ROW
        rng = wso.Range(wso.Cells(target, 1), wso.Cells(target, OC_요약))
        xl.Run("HighlightOutputRow", target)
        check("강조 규칙이 얹혔는가", rng.FormatConditions.Count, 3)
        check("강조가 맨 위 우선순위인가", int(rng.FormatConditions(1).Priority),
              min(int(rng.FormatConditions(i).Priority) for i in (1, 2, 3)))
        xl.Run("ClearOutputHighlight")
        check("강조를 걷으면 진행/완료 규칙만 남는가", rng.FormatConditions.Count, 2)
    finally:
        wb.Close(SaveChanges=False)
        xl.Quit()

    print("\n" + ("모두 통과" if not fails else f"실패 {len(fails)}건: {fails}"))
    return 1 if fails else 0


if __name__ == "__main__":
    sys.exit(main())
