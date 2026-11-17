# -*- coding: utf-8 -*-
"""입력정보 시트에 예시 데이터를 채워 넣는다.

점검 스크립트가 남긴 테스트 계약을 지우고, 처음 여는 사람이 무엇을 어디에 적는지
바로 알 수 있도록 기관정보·결재선·비상연락망·계약 3건을 예시로 되돌려 놓는다.

    python _scripts/restore_sample_data.py
"""
import os
import sys

import win32com.client

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from build_workbook import (BOOK, IN_SCHOOL_ROW, IN_SIGN_ROW, IN_EMERG_ROW,
                            IN_FIRST_ROW, IN_LAST_ROW, IN_LAST_COL, NEW_COLS,
                            SCHOOL_FIELDS, SIGN_FIELDS, EMERG_FIELDS,
                            STASH_COL, STASH_STATE_ROW, STASH_LAST_ROW,
                            STASH_LAST_COL, STASH_VER_OFF, STASH_VERSION)

sys.stdout.reconfigure(encoding="utf-8")

# 배포용 기본값은 **학교** 로 둔다. 교육지원청은 기관구분만 바꾸면 직위 4칸
# (주무관·팀장·과장·교육장)이 자동으로 채워지고 성명은 빈칸으로 남는다.
SCHOOL = ["학교", "예시초등학교", "123-45-00000", "경기도 예시시 예시로 100",
          "031-760-1234", "031-760-1235"]

SIGN = ["주무관", "홍길동", "031-760-1236",
        "행정실장", "김철수",
        "교장", "이영희",
        "교장", "이영희"]

# 관할기관 전화번호는 배포 전에 한 번 확인할 것. 기관이 옮기면 번호도 바뀐다.
EMERG = ["광주소방서", "031-760-0119", "예시병원", "031-799-8114",
         "성남고용노동지청", "031-788-1500"]

# 계약 5건 — 공사 / 장기 용역 / 단기 점검 용역을 섞어 서식 판정
# (2일·30일·7일·90일)이 서로 다르게 나오도록 골랐다. 종료일이 지난 것과
# 아직 남은 것을 함께 두어 진행(연한 초록)·완료(진한 회색)가 처음 열자마자 보인다.
CONTRACTS = [
    {"계약명": "예시초등학교 계단 논슬립 교체 공사",
     "시작일": "2026-05-01", "종료일": "2026-05-15", "점검요일": "수요일",
     "계약금액": 15600000, "업체명": "㈜OO기업", "사업자번호": "165-71-00000",
     "업체주소": "경기도 광주시 회안대로 138", "업체대표": "김OO",
     "대표자전화": "031-000-0000", "안전담당자": "대리 박OO",
     "담당자전화": "010-0000-0000", "계획서유형": "자동"},
    {"계약명": "2026학년도 청소·미화 위탁 용역",
     "시작일": "2026-03-01", "종료일": "2027-02-28", "점검요일": "수요일",
     "계약금액": 45000000, "업체명": "㈜클린서비스", "사업자번호": "110-12-34567",
     "업체주소": "경기도 광주시 경안로 55", "업체대표": "이대표",
     "대표자전화": "031-111-2222", "안전담당자": "팀장 최청소",
     "담당자전화": "010-9999-8888", "계획서유형": "자동"},
    {"계약명": "2026학년도 소방시설 정기점검 용역",
     "시작일": "2026-04-01", "종료일": "2026-04-30", "점검요일": "목요일",
     "계약금액": 8500000, "업체명": "㈜세이프소방", "사업자번호": "220-88-99999",
     "업체주소": "경기도 성남시 중원구 광명로 45", "업체대표": "박방재",
     "대표자전화": "031-333-4444", "안전담당자": "점검원 정소방",
     "담당자전화": "010-5555-6666", "계획서유형": "자동"},
    {"계약명": "폐식용유 위탁용역",
     "시작일": "2026-08-02", "종료일": "2026-09-30", "점검요일": "화요일",
     "계약금액": 9500000, "업체명": "가연상사", "사업자번호": "252-99-88888",
     "업체주소": "경기도 광주시 송정대로 15", "업체대표": "김유상",
     "대표자전화": "031-211-4585", "안전담당자": "과장 설민석",
     "담당자전화": "010-8855-9596", "계획서유형": "자동"},
    {"계약명": "소방안전점검용역",
     "시작일": "2026-08-01", "종료일": "2026-12-31", "점검요일": "수요일",
     "계약금액": 12000000, "업체명": "무안소방", "사업자번호": "252-88-96565",
     "업체주소": "경기도 하남시 하남대로 20", "업체대표": "김소방",
     "대표자전화": "031-888-9955", "안전담당자": "대리 소민지",
     "담당자전화": "010-1111-5656", "계획서유형": "자동"},
]


def main():
    if not os.path.exists(BOOK):
        print("통합문서를 찾을 수 없습니다:", BOOK)
        return 1

    xl = win32com.client.Dispatch("Excel.Application")
    xl.Visible = False
    xl.DisplayAlerts = False
    xl.AutomationSecurity = 1
    xl.EnableEvents = False            # 채워 넣는 동안 자동 판정이 끼어들지 않게
    wb = xl.Workbooks.Open(BOOK)
    try:
        ws = wb.Worksheets("입력정보")

        # 값은 병합된 칸의 첫 열에 쓴다
        for row, vals, fields in ((IN_SCHOOL_ROW, SCHOOL, SCHOOL_FIELDS),
                                  (IN_SIGN_ROW, SIGN, SIGN_FIELDS),
                                  (IN_EMERG_ROW, EMERG, EMERG_FIELDS)):
            for v, (_, c1, _c2) in zip(vals, fields):
                ws.Cells(row, c1).Value = v

        # 계약 칸을 모두 비우고(연번은 남긴다) 예시를 다시 넣는다.
        # 비고·비고수동(V)까지 지워야 진행/완료가 날짜 기준으로 새로 매겨진다.
        body = ws.Range(ws.Cells(IN_FIRST_ROW, 2),
                        ws.Cells(IN_LAST_ROW, NEW_COLS["비고수동"]))
        body.ClearContents()
        body.Interior.ColorIndex = -4142        # xlNone — 자동/수동 표시도 초기화

        for i, row in enumerate(CONTRACTS):
            r = IN_FIRST_ROW + i
            for key, col in NEW_COLS.items():
                if key in row:
                    ws.Cells(r, col).Value = row[key]

        # 결재선 보관함 비우기 — 남에게 나눠 줄 파일에 남의 이름이 실려 가면 안 된다.
        # 여기에 교육지원청 성명이 남아 있으면 기관구분을 바꿀 때 되살아난다.
        # 비워 두면 직위 4칸만 자동으로 채워지고 성명은 빈칸이 된다.
        lib = wb.Worksheets("ContentLibrary")
        lib.Range(lib.Cells(STASH_STATE_ROW, STASH_COL),
                  lib.Cells(STASH_LAST_ROW, STASH_LAST_COL)).ClearContents()
        lib.Cells(STASH_STATE_ROW, STASH_COL).Value = SCHOOL[0]        # 지금 기관구분
        lib.Cells(STASH_STATE_ROW, STASH_COL + STASH_VER_OFF).Value = STASH_VERSION

        xl.EnableEvents = True
        xl.Run("RefreshAllWorkTypesSilent")
        xl.Run("RefreshOutputWeekdays")
        xl.Run("RefreshAllProgress")            # 종료일 → 비고(진행/완료)
        xl.EnableEvents = False

        wb.Save()

        print("예시 데이터를 되돌렸습니다.")
        for i in range(len(CONTRACTS)):
            r = IN_FIRST_ROW + i
            print(f"  {ws.Cells(r, 2).Value}")
            print(f"    계약유형 → {ws.Cells(r, NEW_COLS['계약유형']).Value}"
                  f"   비고 → {ws.Cells(r, NEW_COLS['비고']).Value}")
            print(f"    작업유형 → " + " ".join(
                f"{h}:{ws.Cells(r, c).Value or '×'}"
                for h, c in zip("고소 화기 밀폐 전기".split(), range(14, 18))))
    finally:
        wb.Close(SaveChanges=False)
        xl.EnableEvents = True
        xl.Quit()
    return 0


if __name__ == "__main__":
    sys.exit(main())
