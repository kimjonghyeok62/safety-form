# -*- coding: utf-8 -*-
"""중대재해예방 자동화서식.xlsm 을 새 구조로 다시 짠다.

  · '설정' 시트          — 자동판정 키워드 · 판정기준 · 경로 · 태그 사전
  · '계약유형사전' 시트   — 계약 55종의 위험요인·안전대책·협의체 안건 문구
  · 'ContentLibrary'     — 작업유형 4종 공통 문구 · 달별 계절안건 · 회차 의결문 (숨김)
  · '입력정보'           — 기관정보 / 결재선·대표자 / 비상연락망 / 계약정보 21열
  · '출력관리'           — 서식 15종, 버튼 한 줄 아홉 개(색 맞춤), 점검요일 열
  · VBA 모듈 주입 + 컴파일 검사

화면은 50대 노안 사용자를 기준으로 본문 12pt·행 높이 30px 로 키우고,
가로 스크롤로 계약을 놓치지 않도록 틀 고정을 건다.

실행 전 원본을 _backup 폴더에 복사해 둔다. 엑셀에서 파일을 닫아 두고 실행할 것.
"""
import datetime
import os
import re
import shutil
import sys

import pythoncom
import win32com.client

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from contract_types import CONTRACT_TYPES, CONTENT_LIBRARY, contract_rows

sys.stdout.reconfigure(encoding="utf-8")

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
BOOK = os.path.join(ROOT, "중대재해예방 자동화서식.xlsm")
VBA_DIR = os.path.join(ROOT, "_scripts", "vba")
BACKUP_DIR = os.path.join(ROOT, "_backup")

# --- 레이아웃 상수 (modLayout.bas 와 반드시 일치해야 한다) --------------------
# 구역 사이 빈 줄은 두지 않는다. 위쪽 세 블록이 차지하는 줄을 한 줄이라도 줄여야
# 계약정보(틀 고정 아래)가 한 화면에 더 많이 들어온다. 구역은 제목 줄 윗변의
# 가로줄로 가른다.
IN_SCHOOL_ROW = 4                            # 기관 기본정보 값 행
IN_SIGN_ROW = 6                              # 결재선·기관대표자 값 행
IN_EMERG_ROW = 9                             # 비상연락망 값 행
IN_HEAD_ROW = 11                             # 계약정보 머리글
IN_FIRST_ROW, IN_LAST_ROW = 12, 111          # 계약 100건

CC_계약명, CC_시작일, CC_종료일, CC_점검요일, CC_계약금액 = 2, 3, 4, 5, 6
CC_계획서유형, CC_계약유형, CC_비고, CC_출력버튼 = 18, 19, 20, 21
CC_비고수동 = 22                             # V 숨김 — 비고를 손으로 정했다는 표시
WORK_FIRST_COL, WORK_LAST_COL = 14, 17       # N~Q 고소·화기·밀폐·전기
IN_LAST_COL = CC_출력버튼

OUT_TITLE_ROW = 1                            # 2행은 제목 아래 빈 줄
OUT_BTN_ROW = 3                              # 버튼 한 줄 (8개)
OUT_HEAD_ROW, OUT_RULE_ROW = 4, 5
OUT_FIRST_ROW, OUT_LAST_ROW = 6, 105         # 계약 100건
OUT_ROW_OFFSET = OUT_FIRST_ROW - IN_FIRST_ROW   # 출력관리 행 = 입력정보 행 - 6
OC_계약기간, OC_점검요일, OC_서식1 = 4, 5, 6   # D, E, F
OC_서식15 = OC_서식1 + 14                    # T
OC_요약 = OC_서식1 + 15                      # U

# 결재선 보관함 (ContentLibrary 시트 AA열) — modLayout.bas 의 STASH_* 와 짝이다
STASH_COL = 27
STASH_STATE_ROW = 1
STASH_LAST_ROW = 20
STASH_LAST_COL = STASH_COL + 10       # 이름 + 기관정보 5 + 결재선 5
STASH_VER_OFF = 1                     # AB1 — 보관함 형식 번호 (칸 뜻이 바뀌면 올린다)
STASH_VERSION = 2

WEEKDAYS = "월요일,화요일,수요일,목요일,금요일,토요일,일요일"
ORG_KINDS = "학교,교육지원청"

# 사전 시트
DICT_HEAD_ROW = 3
DICT_FIRST_ROW = 4
DICT_LAST_ROW = DICT_FIRST_ROW + len(CONTRACT_TYPES) - 1

# --- 글자·크기 ---------------------------------------------------------------
FONT = "맑은 고딕"
FONT_BODY = 12          # 50대 노안 기준. 이보다 키우면 계약 21열이 한 화면에 안 들어온다.
FONT_HEAD = 12
FONT_NOTE = 11
FONT_RULE = 10
FONT_SECTION = 13          # 1·2·3 제목
FONT_TITLE = 20            # 시트 1행 제목
ROW_H_TITLE = 40
ROW_H_DATA = 40            # 12pt 두 줄이 잘리지 않는 최소선 (계약정보는 거의 다 줄바꿈)
ROW_H_HEAD = 62
ROW_H_CONTRACT_HEAD = 42   # 계약 머리글 — 두 줄이 들어가는 최소선
ROW_H_OUT_HEAD = 52        # 출력관리 머리글 — 대개 두 줄, 4번만 세 줄이라 그것에 맞춘다
ROW_H_VALUE = 25           # 상단 블록은 머리글도 값 행과 같은 높이로 맞춘다.
                           # 한 줄짜리 이름·값뿐이라 25 면 넉넉하다. 여기서 아낀
                           # 줄 높이가 그대로 계약정보가 보이는 줄 수가 된다.
ROW_H_SECTION = 28
BTN_H = 34
CHECKBOX_SIZE = 16         # 출력관리 [선택] 체크박스가 차지하는 자리.
                           # 네모 자체의 크기는 아니다 ― 그것은 정할 수 없다.


def rgb(r, g, b):
    """엑셀 Interior.Color 는 BGR 정수다."""
    return r + g * 256 + b * 65536


# --- 색 (진한 남색 계열 + 중립) -----------------------------------------------
# NAVY 는 눈으로 고른 값이 아니다. 사용자가 통합문서에서 직접 칠해 두었던 색
# (테마 accent2 #2683C6 을 25% 어둡게)을 그대로 읽어 온 것이다. 다시 빌드해도
# 같은 색이 나오게 하려고 상수로 못박았다.
NAVY = rgb(28, 98, 148)         # #1C6294  기본 머리글 · 서식 1~7 · 15
GREEN = rgb(77, 139, 63)        # 서식 9~10  순회·합동
LIGHT = rgb(222, 226, 232)      # 서식 11~14 사고 시. 흰색은 바탕에 묻혀 안 보인다
NEUTRAL = rgb(71, 85, 105)      # 이동·선택 버튼
SECTION_BG = rgb(238, 242, 245)
SECTION_FG = rgb(51, 65, 85)
SECTION_NOTE_FG = rgb(90, 100, 115)   # 제목 뒤 '— 설명' 부분
BORDER = rgb(213, 219, 225)
SECTION_LINE = rgb(160, 170, 182)     # 구역 사이 가로줄 (8행·12행 윗변)
WHITE = rgb(255, 255, 255)
RULE_BG = rgb(245, 247, 249)
STRIPE = rgb(244, 246, 248)
AUTO_YELLOW = rgb(255, 249, 224)   # modConfig.AUTO_FILL_COLOR 와 같아야 한다
INPUT_BG = rgb(255, 253, 245)      # 사람이 채우는 칸

# 조건부서식 색 — modConfig.bas 의 같은 이름 상수와 짝이다
EMPTY_YELLOW = rgb(255, 242, 153)  # #FFF299  아직 안 채운 칸
FILLED_GRAY = rgb(229, 231, 235)   # #E5E7EB  채운 칸 (눌린 느낌)
DONE_FILL = rgb(120, 127, 136)     # #787F88  계약 완료 — 진한 회색
PROGRESS_FILL = rgb(226, 243, 219)  # #E2F3DB  계약 진행 — 연한 초록형광

# 서식 번호 → 색. 열 머리글과 그 위 버튼이 **같은 색**이어야 어느 버튼이 어느 열을
# 만드는지 눈으로 이어진다. rebuild_output 이 머리글을, buttons 목록이 버튼을 이 값으로 칠한다.
#   1~7   남색   [1~7번 출력] 이 한 문서로 묶는 것
#   8~10  초록   회차를 반복해 뽑는 것 — 협의체 회의록(월 1회)·순회(달마다)·합동(분기)
#   11~14 연회색 사고가 나야 쓰는 것 — 눈에 덜 띄어야 한다
#   15    남색   계약 체결 시 (평상시 쓰는 것이므로 1~7 과 같은 무게로 둔다)
def form_color(n):
    if n <= 7:
        return NAVY
    if n <= 10:
        return GREEN
    if n <= 14:
        return LIGHT
    return NAVY


# ---------------------------------------------------------------- 설정 시트

WORK_KEYWORDS = [
    ("고소작업",
     "외벽,지붕,방수,유리창,창호,전광판,도장,페인트,천장,옥상,비계,사다리,고소,"
     "승강기,엘리베이터,조명,등기구,냉난방,에어컨,실외기,차광막,차광망,현수막,"
     "CCTV,태양광,간판,낙하물,크레인,고가,미화"),
    ("화기작업",
     "용접,용단,절단,배관,보일러,난방,급식,조리,주방,후드,철거,해체,산소,"
     "그라인더,화기,가스,열처리"),
    ("밀폐공간작업",
     "물탱크,저수조,정화조,맨홀,오수,하수,배수,집수정,탱크,지하,피트,덕트,"
     "밀폐,준설,수조,관로"),
    ("전기작업",
     "전기,수배전,변압기,배전반,분전반,통신,전산,네트워크,소방,승강기,태양광,"
     "충전,누전,정전,케이블,자동제어"),
]

PLAN_RULES = [
    ("공사", "공사,설치,교체,보수,수선,철거,해체,시공,증축,개축,도장,방수,포장",
     "01_계획서_공사.hwpx"),
    ("용역", "용역,위탁,청소,미화,경비,급식,조리,유지관리,유지보수,점검,관리,운영",
     "01_계획서_용역.hwpx"),
    ("상세", "(계약기간 90일 이상이거나 위험작업 2종 이상이면 자동으로 이것)",
     "01_계획서_상세.hwpx"),
]

CRITERIA = [
    ("1. 안전보건관리계획서", 2, "계약기간 2일 이상", "기준_계획서"),
    ("2. 수급인 안전보건 수준평가표", 2, "계약기간 2일 이상", "기준_수준"),
    ("3. 공사(용역) 안전보건 점검표", 1, "모든 도급·용역", "기준_도급"),
    ("8. 안전보건협의체 회의록", 30, "30일 이상 — 월 1회(그 달의 마지막 순회점검일)", "기준_협의체"),
    ("9. 작업장 순회점검표", 7, "7일 이상 — 계약기간을 달로 나눠 달마다 1장", "기준_순회"),
    ("10. 작업장 합동점검표", 90, "90일 이상 — 분기 1회(90일 구간의 마지막 점검요일)", "기준_합동"),
]

# 설정 시트 ⑥번에 그대로 실리는 태그 사전.
# 뜻이 같은데 이름이 다른 태그는 '같은 값' 이라고 적어 준다.
TAGS = [
    ("— 기관정보 —", ""),
    ("{{학교명}} · {{기관명}}", "기관 이름 (학교명 · 교육지원청명)"),
    ("{{기관구분}}", "학교 / 교육지원청"),
    ("{{고유번호}}", "기관 고유번호"),
    ("{{주소}}", "기관 주소"),
    ("{{전화번호}}", "기관 대표전화"),
    ("{{팩스번호}}", "기관 팩스번호"),
    ("— 결재선 (직위가 들어갑니다) —", ""),
    ("{{담당자}}", "담당자 직위 — 결재란 자리. 칸이 좁아 직위만 넣습니다"),
    ("{{관리자}}", "관리자 직위 (학교=행정실장, 교육지원청=팀장)"),
    ("{{부서장}}", "부서장 직위 (학교=교장, 교육지원청=과장)"),
    ("{{담당자직위}} {{담당자성명}} {{담당자전화번호}}", "담당자 직위 · 성명 · 전화"),
    ("{{관리자직위}} {{관리자성명}}", "관리자 직위 · 성명"),
    ("{{부서장직위}} {{부서장성명}}", "부서장 직위 · 성명"),
    ("{{기관대표자}}", "기관대표자 '직위 성명' (학교=교장, 교육지원청=교육장)"),
    ("{{기관대표자직위}} {{기관대표자성명}}", "기관대표자 직위 · 성명"),
    ("{{행정실담당자}} {{계약담당자}}", "담당자 성명과 같은 값"),
    ("{{행정실전화번호}}", "담당자 전화와 같은 값"),
    ("{{행정실장}}", "관리자 성명과 같은 값"),
    ("{{교장}} {{대표자}}", "기관대표자 성명과 같은 값"),
    ("— 비상연락망 · 관할기관 —", ""),
    ("{{관할소방서}} {{소방서}}", "관할 소방서 이름"),
    ("{{소방서전화번호}}", "관할 소방서 전화"),
    ("{{인근병원}} {{병원}}", "인근 병원 이름"),
    ("{{병원전화번호}}", "인근 병원 전화"),
    ("{{관할고용노동부}} {{관할고용노동지청}}", "관할 고용노동지청 이름"),
    ("{{고용노동부전화번호}}", "관할 고용노동지청 전화"),
    ("— 계약정보 —", ""),
    ("{{연번}}", "계약 연번"),
    ("{{계약명}}", "계약명"),
    ("{{계약유형}}", "계약유형사전에서 자동으로 걸린 유형명"),
    ("{{계약시작일}} {{계약종료일}}", "2026. 03. 01. 꼴"),
    ("{{계약기간}}", "2026. 03. 01. ~ 2027. 02. 28.(365일간)"),
    ("{{계약일수}}", "365일간"),
    ("{{계약금액}} {{계약금액천원}}", "45,000,000원 / 45,000천원"),
    ("{{점검요일}}", "순회·합동 점검을 도는 요일"),
    ("{{업체명}}", "수급업체 이름"),
    ("{{사업자등록번호}}", "수급업체 사업자등록번호"),
    ("{{업체주소}}", "수급업체 주소"),
    ("{{업체대표자}} {{대표자명}}", "수급업체 대표자"),
    ("{{업체대표자연락처}} {{업체전화}}", "수급업체 대표자 연락처"),
    ("{{안전담당자}} {{업체담당자}}", "수급업체 담당자 (직위 성명)"),
    ("{{담당자전화}} {{안전담당자연락처}} {{업체담당자연락처}}", "수급업체 담당자 연락처"),
    ("— 날짜 · 점검 —", ""),
    ("{{작성일}}", "오늘 날짜"),
    ("{{평가일}}", "계약 시작일"),
    ("{{협의체회의일}}", "협의체 회의일 — 그 달의 마지막 순회점검일 (기본값은 계약 시작일)"),
    ("{{점검년}} {{점검월}}", "순회점검표 한 장이 담는 해와 달"),
    ("{{점검일1}} ~ {{점검일5}}", "그 달의 1주~5주 점검일 (없는 주는 빈칸)"),
    ("{{점검일자}}", "합동점검표의 점검일"),
    ("{{회차}}", "회의록·점검표의 회차 (예: 3회차)"),
    ("  ※ 날짜 표기", "작성일·점검일·회의일은 2026.8.10.(월) 꼴입니다. "
                    "앞에 * 가 붙은 날은 주말·공휴일이라 다음 평일로 미룬 날입니다. "
                    "계약기간에는 요일을 붙이지 않습니다."),
    ("— 계약유형 사전에서 오는 문구 (통째로) —", ""),
    ("{{작업내용}}", "그 계약의 주요 작업내용"),
    ("{{위험요인}}", "그 계약의 유해·위험요인 (항목을 ' · ' 로 이어 붙인 것)"),
    ("{{안전대책}}", "그 계약의 안전대책 (항목을 ' · ' 로 이어 붙인 것)"),
    ("{{협의체안건}}", "안전보건협의체에서 다룰 안건"),
    ("{{점검착안사항}}", "점검할 때 눈여겨볼 것"),
    ("{{작업유형목록}}", "○ 로 표시된 작업유형 (예: 고소작업, 화기작업)"),
    ("{{작업절차}}", "공정을 화살표로 이은 한 줄 (계획서 '- 작업절차:' 자리)"),
    ("{{안전점검계획}}", "계획서 '- 안전점검계획:' 자리"),
    ("{{보호구방호장치}}", "계획서 '보호구 및 방호장치계획:' 자리"),
    ("{{안전보건교육}}", "계획서 '안전보건 교육 계획' 자리"),
    ("{{대피방법}}", "협의체 회의록 '재해발생 위험이 있는 경우 대피방법' 자리"),
    ("— 사전 문구를 한 줄에 하나씩 (줄 번호 태그) —", ""),
    ("{{위험요인1}} ~ {{위험요인5}}", "위험요인의 1·2·3…번째 항목. 없는 번호는 빈칸"),
    ("{{안전대책1}} ~ {{안전대책5}}", "안전대책의 1·2·3…번째 항목. 없는 번호는 빈칸"),
    ("{{작업내용1}} ~ {{작업내용3}}", "주요 작업내용의 1·2·3번째 항목"),
    ("{{협의체안건1}} ~ {{협의체안건3}}", "협의체 안건의 1·2·3번째 항목"),
    ("{{점검착안사항1}} ~ {{점검착안사항3}}", "점검 착안사항의 1·2·3번째 항목"),
    ("{{사용기계기구1}} ~ {{사용기계기구3}}", "사용 기계·기구 (계획서 상세 ② 자리)"),
    ("{{위험물질설비1}} ~ {{위험물질설비3}}", "위험물질·설비 (계획서 상세 ⑥ 자리)"),
    ("{{작업권장시기}}", "학사일정에 맞춘 작업 권장 시기"),
    ("{{위험성등급}}", "위험성 등급 (상이면 고위험 작업허가서 대상)"),
    ("{{학교확인사항1}} ~ {{학교확인사항3}}", "학교(행정실)가 확인·조치할 것"),
    ("{{업체조치사항1}} ~ {{업체조치사항3}}", "수급업체가 시행할 것"),
    ("{{협의체합의문1}} ~ {{협의체합의문3}}", "회의록 의결란에 그대로 쓰는 합의문"),
    ("{{상시안건1}} ~ {{상시안건3}}",
     "어느 계약에서나 다루는 상시 안건. 협의체 회의록에서는 그 달의 계절·학사일정 안건이 들어갑니다"),
]

# 계약유형사전 시트의 열. contract_types.py 의 T() 인수 차례와 반드시 같아야 한다.
# ' · ' 로 나눈 칸은 **항목 목록**이다. VBA 가 쪼개어 한 줄에 하나씩 서식에 넣는다.
DICT_HEADERS = ["유형명", "계약명 키워드 (쉼표로 구분)", "주요 작업내용",
                "위험요인", "안전대책", "협의체 안건", "점검 착안사항",
                "기본 작업유형",
                "작업절차", "안전점검계획", "보호구·방호장치", "안전보건교육",
                "사용 기계·기구", "위험물질·설비", "비상 대피방법",
                "작업 권장 시기", "학교 확인사항", "업체 조치사항",
                "위험성 등급", "협의체 합의문"]
DICT_LAST_COL_LETTER = chr(ord("A") + len(DICT_HEADERS) - 1)       # O


def sheet_title(ws, last_col, text):
    """시트 1행 제목. 색을 깔지 않고 글자 크기로만 제목임을 드러낸다."""
    ws.Range(ws.Cells(1, 1), ws.Cells(1, last_col)).Merge()
    c = ws.Cells(1, 1)
    c.Value = text
    c.Font.Name = FONT
    c.Font.Size = FONT_TITLE
    c.Font.Bold = True
    c.Font.Color = SECTION_FG
    c.Interior.ColorIndex = -4142          # xlNone
    c.HorizontalAlignment = -4131          # xlLeft
    c.VerticalAlignment = -4108            # xlCenter
    ws.Rows(1).RowHeight = ROW_H_TITLE


def characters(cel, start, length):
    """셀 안의 일부 글자만 집는다.

    Characters 는 매개변수가 붙은 속성이라 늦은 바인딩(win32com Dispatch)에서는
    cel.Characters(...) 로 부를 수 없다. GetCharacters 로 불러야 한다.
    """
    return cel.GetCharacters(start, length)


def section_title(ws, row, text, last_col=6):
    """'1. 기관정보 — 설명…' 꼴의 구역 제목.

    '—' 앞은 제목이라 크고 굵게, 뒤는 설명이라 작고 흐리게 쓴다. 한 칸 안에서
    글자마다 서식이 다르므로 Characters 로 뒷부분만 따로 손본다. 이 서식은
    나중에 셀 전체에 Font 를 다시 걸면 지워지므로 반드시 맨 마지막에 한다.
    """
    ws.Range(ws.Cells(row, 1), ws.Cells(row, last_col)).Merge()
    c = ws.Cells(row, 1)
    c.Value = text
    c.Font.Name = FONT
    c.Font.Bold = True
    c.Font.Size = FONT_SECTION
    c.Font.Color = SECTION_FG
    c.Interior.ColorIndex = -4142          # xlNone — 배경색 없음
    c.VerticalAlignment = -4108
    ws.Rows(row).RowHeight = ROW_H_SECTION

    i = text.find("—")
    if i > 0:
        note = characters(c, i + 1, len(text) - i)    # 엑셀 Characters 는 1부터
        note.Font.Size = FONT_NOTE
        note.Font.Bold = False
        note.Font.Color = SECTION_NOTE_FG


def table_header(ws, row, headers, color=NAVY, start_col=1):
    for i, h in enumerate(headers, start=start_col):
        c = ws.Cells(row, i)
        c.Value = h
        c.Font.Bold = True
        c.Font.Size = FONT_HEAD
        c.Font.Color = WHITE
        c.Interior.Color = color
        c.HorizontalAlignment = -4108
        c.VerticalAlignment = -4108
        c.WrapText = True


def thin_borders(rng):
    for idx in (7, 8, 9, 10, 11, 12):
        b = rng.Borders(idx)
        b.LineStyle = 1
        b.Color = BORDER
        b.Weight = 2                 # xlThin


XL_EXPRESSION = 2
XL_EDGE_TOP = 8


def add_cf(rng, formula, fill=None, font=None, stop_if_true=False):
    """조건부서식 규칙 하나를 건다.

    FormatConditions.Add 는 맨 뒤(우선순위 낮음)에 붙는다. 그러므로 **부르는 차례가
    곧 우선순위**다. SetFirstPriority 를 부르면 오히려 순서가 뒤집히므로 쓰지 않는다.

    두 번째 인수(Operator)는 xlExpression 일 때 쓰지 않는다. 그런데 늦은 바인딩에서는
    이름 붙인 인수(Formula1=)도, pythoncom.Missing 도 받지 않는다. 셋 중 pythoncom.Empty
    만 통한다 ― 실제로 셋 다 넣어 보고 확인했다.
    """
    fc = rng.FormatConditions.Add(XL_EXPRESSION, pythoncom.Empty, formula)
    if fill is not None:
        fc.Interior.Color = fill
    if font is not None:
        fc.Font.Color = font
    fc.StopIfTrue = stop_if_true
    return fc


def progress_cf(rng, remark_formula):
    """계약의 진행/완료에 따라 줄 전체를 칠한다.

    색의 단일 출처는 **비고 칸의 값 하나**다. 날짜는 비고를 채우고(modProgress),
    비고가 색을 정한다. 그래야 '사용자가 드롭다운으로 바꾸면 그 값이 이긴다' 가
    따로 손대지 않아도 성립한다.
    """
    add_cf(rng, f'={remark_formula}="완료"', DONE_FILL, WHITE, stop_if_true=True)
    add_cf(rng, f'={remark_formula}="진행"', PROGRESS_FILL, stop_if_true=True)


def set_list(rng, items, title="", msg=""):
    rng.Validation.Delete()
    rng.Validation.Add(3, 1, 1, items)
    rng.Validation.IgnoreBlank = True
    rng.Validation.InCellDropdown = True
    # 설명 말풍선 — 그 칸을 누를 때만 뜬다. 셀 메모처럼 빨간 삼각형이 남지 않는다.
    if msg:
        rng.Validation.ShowInput = True
        rng.Validation.InputTitle = title
        rng.Validation.InputMessage = msg
    else:
        rng.Validation.ShowInput = False
    rng.Validation.ShowError = False   # 손으로 다른 값을 넣어도 막지 않는다


BTN_PAD = 18            # 버튼 글자 양옆 여백 (합계)


def add_button(ws, anchor, caption, macro, dx, dy, h, color=NAVY, fg=None):
    """버튼 하나. 폭은 글자에 맞춰 재고 좌우 여백만 붙인다.

    폭을 손으로 정해 두면 글자보다 넉넉해 오른쪽이 휑하게 남는다. AutoSize 로
    글자에 딱 맞춘 뒤 여백을 더하는 편이 언제나 보기 좋다.
    """
    a = ws.Range(anchor)
    shp = ws.Shapes.AddShape(5, a.Left + dx, a.Top + dy, 100, h)   # 5 = 둥근 사각형
    shp.Name = "btn_" + macro
    shp.TextFrame2.TextRange.Text = caption
    shp.TextFrame2.TextRange.Font.Name = FONT
    shp.TextFrame2.TextRange.Font.Size = FONT_HEAD
    shp.TextFrame2.TextRange.Font.Bold = True
    shp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = fg if fg is not None else WHITE
    shp.TextFrame2.WordWrap = False
    shp.TextFrame2.MarginLeft = 0
    shp.TextFrame2.MarginRight = 0

    shp.TextFrame2.AutoSize = 1          # msoAutoSizeShapeToFitText
    w = shp.Width + BTN_PAD
    shp.TextFrame2.AutoSize = 0          # msoAutoSizeNone — 높이를 되돌리기 위해
    shp.Width = w
    shp.Height = h

    shp.Fill.ForeColor.RGB = color
    shp.Fill.Solid()
    if color == LIGHT:
        # 연회색 버튼은 시트 바탕과 구분이 흐리다. 테두리를 둘러 단추처럼 보이게 한다.
        shp.Line.Visible = True
        shp.Line.ForeColor.RGB = BORDER
        shp.Line.Weight = 1
    else:
        shp.Line.Visible = False
    shp.Shadow.Visible = False
    shp.OnAction = macro
    return shp


def addr(r, c):
    letters = ""
    while c:
        c, rem = divmod(c - 1, 26)
        letters = chr(65 + rem) + letters
    return f"${letters}${r}"


def col_letter(c):
    letters = ""
    while c:
        c, rem = divmod(c - 1, 26)
        letters = chr(65 + rem) + letters
    return letters


def add_name(wb, name, refers_to):
    try:
        wb.Names(name).Delete()
    except Exception:
        pass
    wb.Names.Add(Name=name, RefersTo=refers_to)


def sheet_area(ws, rows, cols):
    """쓰는 범위만 지우고 기본 글꼴을 입힌다.

    ws.Cells 전체에 걸면 100만 행 × 1만 6천 열을 훑느라 몇 분이 걸린다.
    """
    rng = ws.Range(ws.Cells(1, 1), ws.Cells(rows, cols))
    # 옛 배치에서 병합해 둔 칸이 남아 있으면 그 안쪽 칸에 값을 써도 조용히 사라진다.
    # (A5:I5 가 병합되어 있어 기관명·주소가 안 들어가던 적이 있다)
    rng.UnMerge()
    rng.Clear()
    rng.Validation.Delete()
    rng.FormatConditions.Delete()       # 옛 조건부서식이 남아 새 규칙과 겹치지 않게
    rng.Font.Name = FONT
    rng.Font.Size = FONT_BODY
    rng.WrapText = False
    return rng


def fresh_sheet(wb, name):
    """같은 이름의 시트가 있으면 지우고 새로 만든다."""
    names = [s.Name for s in wb.Worksheets]
    wb.Application.DisplayAlerts = False
    if name in names:
        # 아주 숨김(xlSheetVeryHidden) 시트는 보이게 만들어야 지울 수 있다
        wb.Worksheets(name).Visible = -1
        wb.Worksheets(name).Delete()
    wb.Application.DisplayAlerts = True
    ws = wb.Worksheets.Add(After=wb.Worksheets(wb.Worksheets.Count))
    ws.Name = name
    ws.Range("A1:Z400").Font.Name = FONT
    ws.Range("A1:Z400").Font.Size = FONT_BODY
    return ws


# ---------------------------------------------------------------- 설정 시트

def build_config(wb):
    ws = fresh_sheet(wb, "설정")

    ws.Range("A1:F1").Merge()
    ws.Range("A1").Value = "설정 — 자동 판정 사전 · 판정 기준 · 폴더 · 태그"
    ws.Range("A1").Font.Size = 18
    ws.Range("A1").Font.Bold = True
    ws.Range("A1").Font.Color = SECTION_FG
    ws.Range("A1").Interior.Color = SECTION_BG
    ws.Rows(1).RowHeight = 34
    ws.Range("A2").Value = ("여기 값을 고치면 판정과 출력이 바로 달라집니다. 코드를 고칠 필요가 없습니다.")
    ws.Range("A2").Font.Size = FONT_NOTE

    r = 4
    section_title(ws, r, "① 작업유형 자동 판정 키워드 — 계약명에 이 낱말이 들어 있으면 ○ 를 찍습니다")
    table_header(ws, r + 1, ["작업유형", "키워드 (쉼표로 구분)"])
    for i, (name, kw) in enumerate(WORK_KEYWORDS):
        ws.Cells(r + 2 + i, 1).Value = name
        ws.Cells(r + 2 + i, 2).Value = kw
    kw_range = f"$A${r+2}:$B${r+2+len(WORK_KEYWORDS)-1}"
    r += 2 + len(WORK_KEYWORDS) + 1

    section_title(ws, r, "② 안전보건관리계획서 서식 자동 선택")
    table_header(ws, r + 1, ["유형", "계약명 키워드", "쓸 한글 서식 파일"])
    for i, row in enumerate(PLAN_RULES):
        for j, v in enumerate(row, start=1):
            ws.Cells(r + 2 + i, j).Value = v
    plan_range = f"$A${r+2}:$C${r+2+len(PLAN_RULES)-1}"
    r += 2 + len(PLAN_RULES) + 1

    section_title(ws, r, "③ 서식 필요 판정 기준 (계약기간 일수)")
    table_header(ws, r + 1, ["서식", "기준(일)", "설명"])
    crit_rows = {}
    for i, (nm, days, desc, key) in enumerate(CRITERIA):
        rr = r + 2 + i
        ws.Cells(rr, 1).Value = nm
        ws.Cells(rr, 2).Value = days
        ws.Cells(rr, 3).Value = desc
        crit_rows[key] = rr
    r += 2 + len(CRITERIA) + 1

    section_title(ws, r, "④ 점검요일 기본값 — 계약을 새로 넣으면 이 요일이 자동으로 채워집니다")
    ws.Cells(r + 1, 1).Value = "기본 점검요일"
    ws.Cells(r + 1, 2).Value = "수요일"
    wd_cell = f"$B${r+1}"
    set_list(ws.Range(f"B{r+1}"), WEEKDAYS)
    ws.Cells(r + 2, 1).Value = "추천 안내창"
    ws.Cells(r + 2, 2).Value = "표시"
    ws.Cells(r + 2, 3).Value = ("계약명을 넣으면 무엇을 추천했는지 알려 주고 "
                                "잠시 뒤 저절로 닫힙니다")
    notice_cell = f"$B${r+2}"
    set_list(ws.Range(f"B{r+2}"), "표시,숨김")
    r += 4

    section_title(ws, r, "⑤ 폴더")
    ws.Cells(r + 1, 1).Value = "한글 서식 폴더"
    ws.Cells(r + 1, 2).Value = ".\\form\\templates\\"
    ws.Cells(r + 1, 3).Value = "'.\\' 는 이 엑셀 파일이 있는 폴더를 뜻합니다"
    ws.Cells(r + 2, 1).Value = "임시 폴더"
    ws.Cells(r + 2, 2).Value = ""
    ws.Cells(r + 2, 3).Value = "비워 두면 이 엑셀 파일 옆의 '출력' 폴더를 씁니다"
    tmpl_cell = f"$B${r+1}"
    temp_cell = f"$B${r+2}"
    r += 4

    section_title(ws, r, "⑥ 한글 서식에 써 넣는 치환 태그 — 서식 파일에 이 낱말을 그대로 타이핑해 두면 값이 채워집니다")
    table_header(ws, r + 1, ["태그", "채워지는 내용"])
    for i, (tag, desc) in enumerate(TAGS):
        rr = r + 2 + i
        ws.Cells(rr, 1).Value = tag
        ws.Cells(rr, 2).Value = desc
        if desc == "":                       # 소제목 줄
            ws.Cells(rr, 1).Font.Bold = True
            ws.Range(ws.Cells(rr, 1), ws.Cells(rr, 2)).Interior.Color = RULE_BG

    ws.Columns("A").ColumnWidth = 46
    ws.Columns("B").ColumnWidth = 62
    ws.Columns("C").ColumnWidth = 40
    ws.Columns("D").ColumnWidth = 12
    ws.Rows(f"4:{r + 2 + len(TAGS)}").VerticalAlignment = -4108

    # 이름 정의 — VBA 는 오직 이 이름으로만 설정을 읽는다
    add_name(wb, "SET_키워드", f"=설정!{kw_range}")
    add_name(wb, "SET_계획서", f"=설정!{plan_range}")
    add_name(wb, "SET_기본점검요일", f"=설정!{wd_cell}")
    add_name(wb, "SET_추천안내", f"=설정!{notice_cell}")
    add_name(wb, "SET_템플릿폴더", f"=설정!{tmpl_cell}")
    add_name(wb, "SET_임시폴더", f"=설정!{temp_cell}")
    for key, rr in crit_rows.items():
        add_name(wb, key, f"=설정!$B${rr}")

    ws.Rows(3).RowHeight = 30
    x = 0
    for caption, macro in (("보안모듈 등록", "RegisterHwpSecurityModule"),
                           ("서식 폴더 열기", "Btn_OpenTemplateFolder"),
                           ("계약유형사전", "Btn_GoToDict"),
                           ("다시 판정", "Btn_RefreshDetect")):
        shp = add_button(ws, "A3", caption, macro, x, 2, 26, NEUTRAL)
        x += shp.Width + 8
    return ws


# ------------------------------------------------------------ 계약유형사전

def build_dict(wb):
    ws = fresh_sheet(wb, "계약유형사전")

    ws.Range(f"A1:{DICT_LAST_COL_LETTER}1").Merge()
    ws.Range("A1").Value = "계약유형 사전 — 계약명을 넣으면 여기 한 줄이 걸려 한글 서식의 문구가 채워집니다"
    ws.Range("A1").Font.Size = 18
    ws.Range("A1").Font.Bold = True
    ws.Range("A1").Font.Color = SECTION_FG
    ws.Range("A1").Interior.Color = SECTION_BG
    ws.Rows(1).RowHeight = 34

    ws.Range(f"A2:{DICT_LAST_COL_LETTER}2").Merge()
    ws.Range("A2").Value = (
        "문구는 여기서 바로 고치시면 됩니다. 키워드가 여러 줄에 걸리면 "
        "가장 긴 낱말이 걸린 줄을 씁니다(긴 낱말일수록 구체적이므로). "
        "'기본 작업유형' 은 고소·화기·밀폐·전기를 | 로 이어 적습니다.\n"
        "위험요인·안전대책·사용 기계기구·위험물질 칸은 항목을 ' · '(앞뒤 한 칸) 로 "
        "나누어 적습니다. 한글 서식에는 한 줄에 하나씩 들어갑니다.\n"
        "P~T 열은 학교 현장 항목입니다. '학교 확인사항·업체 조치사항' 은 계획서의 "
        "역할 분담란에, '협의체 합의문' 은 회의록 의결란에 그대로 들어갑니다. "
        "위험성 등급이 '상' 이면 착수 전 고위험 작업허가서를 받는 계약입니다.")
    ws.Range("A2").Font.Size = FONT_NOTE
    ws.Rows(2).RowHeight = 46
    ws.Range("A2").WrapText = True

    table_header(ws, DICT_HEAD_ROW, DICT_HEADERS)
    ws.Rows(DICT_HEAD_ROW).RowHeight = 40

    for i, row in enumerate(contract_rows()):
        rr = DICT_FIRST_ROW + i
        for j, v in enumerate(row, start=1):
            ws.Cells(rr, j).Value = v
        ws.Rows(rr).RowHeight = 46

    body = ws.Range(ws.Cells(DICT_HEAD_ROW, 1), ws.Cells(DICT_LAST_ROW, len(DICT_HEADERS)))
    body.VerticalAlignment = -4108
    body.WrapText = True
    thin_borders(body)

    widths = [18, 30, 34, 44, 44, 30, 30, 14,
              38, 32, 32, 32, 34, 30, 32,
              34, 40, 40, 22, 48]
    for i, w in enumerate(widths, start=1):
        ws.Columns(i).ColumnWidth = w

    ws.Activate()
    wb.Windows(1).FreezePanes = False
    ws.Range(f"A{DICT_FIRST_ROW}").Select()
    wb.Windows(1).FreezePanes = True

    add_name(wb, "SET_계약유형사전",
             f"=계약유형사전!$A${DICT_FIRST_ROW}:"
             f"${DICT_LAST_COL_LETTER}${DICT_LAST_ROW}")
    add_name(wb, "SET_계약유형목록",
             f"=계약유형사전!$A${DICT_FIRST_ROW}:$A${DICT_LAST_ROW}")
    return ws


def read_stash(wb):
    """기관정보 보관함(AA열)을 읽어 둔다. 시트를 새로 만들면 지워지기 때문이다.

    기관구분을 잘못 골랐다가 되돌렸을 때 되살릴 값이 여기 들어 있다.
    빌드를 다시 돌렸다고 그 기억까지 잃으면 안 된다.
    """
    try:
        ws = wb.Worksheets("ContentLibrary")
    except Exception:
        return []
    rows = []
    for r in range(1, STASH_LAST_ROW + 1):
        rows.append([ws.Cells(r, c).Value
                     for c in range(STASH_COL, STASH_LAST_COL + 1)])
    return migrate_stash(rows)


def migrate_stash(rows):
    """옛 보관함(성명·전화 5칸)을 새 보관함(기관정보 5 + 결재선 5)으로 옮긴다.

    칸 뜻이 바뀌었는데 그냥 두면 사람 이름이 기관명·주소 칸으로 되살아난다.
    형식 번호(AB1)가 없는 것이 옛 보관함이다. 값을 뒤쪽 5칸으로 밀어 준다.
    """
    if not rows:
        return rows
    if rows[0][STASH_VER_OFF] == STASH_VERSION:
        return rows

    for i in range(1, len(rows)):
        kind, old = rows[i][0], rows[i][1:6]
        rows[i] = [kind] + [None] * 5 + list(old)
    rows[0][STASH_VER_OFF] = STASH_VERSION
    return rows


def build_content_library(wb, stash=None):
    """작업유형 4종 + 공통 문구. 계약유형 사전에 걸리지 않는 계약의 되돌림 값이다."""
    ws = fresh_sheet(wb, "ContentLibrary")
    for j, h in enumerate(["작업유형", "구분", "순서", "내용"], start=1):
        ws.Cells(1, j).Value = h
        ws.Cells(1, j).Font.Bold = True

    seq = {}
    for i, (kind, cat, txt) in enumerate(CONTENT_LIBRARY, start=2):
        seq[(kind, cat)] = seq.get((kind, cat), 0) + 1
        ws.Cells(i, 1).Value = kind
        ws.Cells(i, 2).Value = cat
        ws.Cells(i, 3).Value = seq[(kind, cat)]
        ws.Cells(i, 4).Value = txt

    ws.Columns("A:C").ColumnWidth = 14
    ws.Columns("D").ColumnWidth = 90

    # 결재선 보관함 — modLayout.bas 의 STASH_* 가 가리키는 자리
    ws.Cells(1, STASH_COL - 1).Value = "▼ 결재선 보관함 (건드리지 마세요)"
    for i, row in enumerate(stash or []):
        for j, v in enumerate(row):
            if v is not None:
                ws.Cells(1 + i, STASH_COL + j).Value = v

    # AA1 = 지금 적용 중인 기관구분. 이것이 비어 있으면 처음 기관구분을 바꿀 때
    # '전에 무엇이었는지'를 몰라 지금 적어 둔 값을 보관하지 못한다.
    if not ws.Cells(STASH_STATE_ROW, STASH_COL).Value:
        kind = wb.Worksheets("입력정보").Cells(IN_SCHOOL_ROW, 1).Value
        ws.Cells(STASH_STATE_ROW, STASH_COL).Value = str(kind or "학교").strip()
    ws.Cells(STASH_STATE_ROW, STASH_COL + STASH_VER_OFF).Value = STASH_VERSION

    ws.Visible = 2          # xlSheetVeryHidden — 시트 탭 메뉴에도 안 보인다
    return ws


# -------------------------------------------------------------- 입력정보

# 위쪽 세 블록은 아래 계약정보와 열 격자를 함께 쓴다. 계약정보에 맞춘 열 너비로는
# 주소·전화번호가 잘리므로, 칸을 여러 열에 걸쳐 병합해 쓸 만한 폭을 만든다.
# (이름, 시작 열, 끝 열)
# 이름을 길게 달지 않는다. 계약정보에 맞춘 열 너비로는 C:D 두 칸을 합쳐도 좁아
# '기관명 (학교 · 교육지원청)' 같은 이름은 옆 칸으로 삐져나간다.
# 무엇을 적는 칸인지는 바로 위 구역 제목이 이미 말해 준다.
SCHOOL_FIELDS = [("기관구분", 1, 2), ("기관명", 3, 4),
                 ("고유번호", 5, 6), ("주소", 7, 9),
                 ("대표전화", 10, 11), ("팩스번호", 12, 13)]
SIGN_FIELDS = [("담당자 직위", 1, 2), ("담당자 성명", 3, 3), ("담당자 전화", 4, 5),
               ("관리자 직위", 6, 7), ("관리자 성명", 8, 8),
               ("부서장 직위", 9, 10), ("부서장 성명", 11, 12),
               ("기관대표자 직위", 13, 14), ("기관대표자 성명", 15, 17)]
EMERG_FIELDS = [("관할 소방서", 1, 2), ("소방서 전화", 3, 4),
                ("인근 병원", 5, 6), ("병원 전화", 7, 8),
                ("관할 고용노동지청", 9, 11), ("고용노동지청 전화", 12, 13)]

# 위 칸들이 담는 값의 이름 (read_existing_school 이 돌려주는 키와 같은 순서)
SCHOOL_KEYS = ["기관구분", "기관명", "고유번호", "주소", "전화번호", "팩스번호"]
SIGN_KEYS = ["담당자직위", "담당자성명", "담당자전화", "관리자직위", "관리자성명",
             "부서장직위", "부서장성명", "대표자직위", "대표자성명"]
EMERG_KEYS = ["소방서", "소방서전화", "병원", "병원전화", "노동지청", "노동지청전화"]

CONTRACT_HEADERS = [
    "연번", "계약명", "계약 시작일\n(더블=달력)", "계약 종료일\n(더블=달력)",
    "점검요일", "계약금액", "업체명", "사업자등록번호", "업체 주소",
    "업체 대표자", "대표자 연락처", "업체 담당자\n(직위 성명)", "담당자 연락처",
    "고소작업", "화기작업", "밀폐공간", "전기작업",
    "계획서 유형", "계약유형", "비고", "출력",
]

# '(더블=달력)' 은 안내라 제목보다 작고 가늘게 쓴다. 머리글을 다 그린 뒤 손본다.
SUBLABEL = "(더블=달력)"
# 열 너비 합계는 264 다. 100% 배율 1920px 화면에서 U열 [▶ 출력] 까지 한 화면에
# 들어오는 선이다. 넉넉하게 두면 버튼이 화면 밖으로 밀려 '출력이 안 된다' 는
# 오해를 산다. 좁아진 만큼은 줄바꿈(WrapText)이 받아 준다.
# 계약유형(S)은 줄바꿈하지 않으므로 좁으면 글자가 **잘린다**. '소방시설 점검·공사'
# 가 한 줄에 들어가는 22 가 최소선이다.
# 날짜·금액(C·D·F)과 업체 주소(I)는 한 줄에 떨어지도록 넉넉히 준다.
# 합계 280 은 100% 배율에서도 U열 [▶ 출력] 이 화면 안에 남는 선이다.
CONTRACT_WIDTHS = [5.5, 26, 14, 14, 9.5, 14, 17, 15, 30, 11, 15.5, 15, 15.5,
                   7, 7, 7, 7, 9.14, 22, 8.43, 10.5]

WORK_TIP_TITLE = "자동으로 판정된 값입니다"
WORK_TIP = ("드롭다운에서 ○ / × 를 직접 고를 수 있습니다.\n"
            "한 번 고치면 회색이 되어, 계약명을 바꿔도 자동 판정이 이 칸을 건드리지 않습니다.")
TYPE_TIP_TITLE = "계약명으로 자동으로 찾은 유형입니다"
TYPE_TIP = ("드롭다운에서 직접 고를 수 있습니다.\n"
            "이 유형에 맞는 위험요인·안전대책 문구가 한글 서식에 들어갑니다.\n"
            "문구는 '계약유형사전' 시트에서 고치시면 됩니다.")
REMARK_TIP_TITLE = "계약 종료일을 보고 자동으로 정합니다"
REMARK_TIP = ("종료일이 지났으면 완료(진한 회색), 아직이면 진행(연한 초록)입니다.\n"
              "드롭다운으로 직접 고르면 날짜와 상관없이 그 값이 유지됩니다.\n"
              "칸을 비우면 다시 날짜 기준으로 돌아갑니다.")


def locate_layout(ws):
    """지금 시트에 값이 실제로 어디 있는지 찾아낸다.

    행 번호를 상수로 박아 두면 배치를 손볼 때마다 옛 통합문서를 '모르는 배치' 로
    잘못 보고 사용자가 적어 둔 계약을 통째로 날린다. 그래서 머리글 글자를 찾아
    위치를 알아낸다. 값 행은 늘 머리글 바로 아래 줄이다.

    (계약 머리글 행, 비상 값 행, 기관 값 행, 결재선 값 행) 을 돌려준다. 못 찾으면 0.
    구역 사이 빈 줄을 없애면서 네 블록이 다 같이 위로 밀렸으므로, 하나라도 상수로
    두면 옛 파일에서 엉뚱한 줄을 읽는다.
    """
    def find(col, text, lo, hi):
        for r in range(lo, hi):
            if str(ws.Cells(r, col).Value or "").strip().startswith(text):
                return r + 1                     # 값은 머리글 바로 아래 줄
        return 0

    head = 0
    for r in range(6, 30):                       # 계약 머리글 — B열이 '계약명'
        if str(ws.Cells(r, 2).Value or "").strip() == "계약명":
            head = r
            break

    return (head,
            find(1, "관할 소방서", 5, 18),
            find(1, "기관구분", 2, 12),
            find(1, "담당자 직위", 3, 14))


# 옛 배치(A~T, M 은 빈 열) → 읽어 둘 항목별 열 번호
OLD_COLS = {"계약명": 2, "시작일": 3, "종료일": 4, "계약금액": 5, "업체명": 6,
            "사업자번호": 7, "업체대표": 8, "업체주소": 9, "대표자전화": 10,
            "안전담당자": 11, "담당자전화": 12, "비고": 18, "점검요일": 19,
            "계획서유형": 20}
# 새 배치의 열 번호
# '비고수동' 도 함께 실어야 다시 빌드해도 사용자가 손으로 정한 진행/완료가 살아남는다.
NEW_COLS = {"계약명": 2, "시작일": 3, "종료일": 4, "점검요일": 5, "계약금액": 6,
            "업체명": 7, "사업자번호": 8, "업체주소": 9, "업체대표": 10,
            "대표자전화": 11, "안전담당자": 12, "담당자전화": 13,
            "계획서유형": 18, "계약유형": 19, "비고": 20, "비고수동": 22}


def read_existing_contracts(ws, head):
    """계약 데이터를 항목 이름으로 읽어 둔다. 옛 배치·새 배치 둘 다 받는다.

    행 수는 배치마다 다르므로(15건이었다가 100건이 된다) 머리글 아래부터 실제로
    쓰인 마지막 행까지 훑는다. 중간에 빈 줄이 있어도 멈추지 않는다.
    """
    if head:
        cols, first = NEW_COLS, head + 1
        last = max(first, ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1)
    else:
        cols, first, last = OLD_COLS, 9, 23      # 옛 배치 (A~T, M 은 빈 열)

    rows = []
    for r in range(first, last + 1):
        name = ws.Cells(r, cols["계약명"]).Value
        if name is None or str(name).strip() == "":
            continue
        rows.append({k: ws.Cells(r, c).Value for k, c in cols.items()})
    return rows


def field_block(ws, head_row, value_row, fields):
    """머리글 한 줄 + 값 한 줄을 그린다. 칸마다 여러 열을 병합해 폭을 확보한다."""
    for label, c1, c2 in fields:
        if c2 > c1:
            ws.Range(ws.Cells(head_row, c1), ws.Cells(head_row, c2)).Merge()
            ws.Range(ws.Cells(value_row, c1), ws.Cells(value_row, c2)).Merge()
        h = ws.Cells(head_row, c1)
        h.Value = label
        h.Font.Bold = True
        h.Font.Size = FONT_HEAD
        h.Font.Color = WHITE
        h.Interior.Color = NAVY
        h.HorizontalAlignment = -4108
        h.VerticalAlignment = -4108
        v = ws.Cells(value_row, c1)
        v.Interior.Color = INPUT_BG
        v.HorizontalAlignment = -4108
        v.VerticalAlignment = -4108
    last = fields[-1][2]
    thin_borders(ws.Range(ws.Cells(head_row, 1), ws.Cells(head_row, last)))
    thin_borders(ws.Range(ws.Cells(value_row, 1), ws.Cells(value_row, last)))
    # 머리글도 값 행과 같은 높이로. 한 줄짜리 이름들이라 더 키울 이유가 없다.
    ws.Rows(head_row).RowHeight = ROW_H_VALUE
    ws.Rows(value_row).RowHeight = ROW_H_VALUE


def read_existing_school(ws, head, emerg, school, sign):
    """기관정보를 읽어 둔다. 옛 배치·새 배치 둘 다 받는다.

    네 블록 모두 찾아낸 위치로 읽는다. 상수로 읽으면 배치를 한 줄이라도 밀 때마다
    머리글 줄을 값으로 읽어 사용자가 적어 둔 기관정보를 통째로 날린다.
    """
    if head:
        out = {}
        for row, keys, fields in ((school or IN_SCHOOL_ROW, SCHOOL_KEYS, SCHOOL_FIELDS),
                                  (sign or IN_SIGN_ROW, SIGN_KEYS, SIGN_FIELDS),
                                  (emerg or IN_EMERG_ROW, EMERG_KEYS, EMERG_FIELDS)):
            for key, (_, c1, _c2) in zip(keys, fields):
                out[key] = ws.Cells(row, c1).Value
        return out

    v = [ws.Cells(IN_SCHOOL_ROW, c).Value for c in range(1, 10)]
    return {"기관구분": "학교", "기관명": v[0], "고유번호": v[1], "주소": v[3],
            "전화번호": v[4], "팩스번호": v[5],
            "담당자직위": "주무관", "담당자성명": v[6], "담당자전화": None,
            "관리자직위": "행정실장", "관리자성명": v[7],
            "부서장직위": "교장", "부서장성명": v[8],
            "대표자직위": "교장", "대표자성명": v[8],
            "소방서": None, "소방서전화": None, "병원": None,
            "병원전화": None, "노동지청": None, "노동지청전화": None}


def rebuild_input(wb):
    ws = wb.Worksheets("입력정보")

    head, emerg, school, sign = locate_layout(ws)
    old_contracts = read_existing_contracts(ws, head)
    old_school = read_existing_school(ws, head, emerg, school, sign)

    for shp in list(ws.Shapes):
        shp.Delete()
    for cb in list(ws.CheckBoxes()):
        try:
            cb.Delete()
        except Exception:
            pass

    # 시트 전체(ws.Cells)에 서식·유효성 작업을 걸면 100만 행을 훑느라 몇 분씩 걸린다.
    # 쓰는 범위만 손댄다.
    sheet_area(ws, IN_LAST_ROW + 10, IN_LAST_COL + 4)

    # 배치가 위로 밀리면서 남게 된 아래쪽 줄들. Clear 는 행 높이를 되돌리지 않아
    # 40px 짜리 빈 줄이 표 끝에 남는다.
    ws.Rows(f"{IN_LAST_ROW + 1}:{IN_LAST_ROW + 10}").UseStandardHeight = True

    sheet_title(ws, IN_LAST_COL, "중대재해예방 안전관련 서식 자동화 — 입력정보")

    # ① 기관 기본정보 + ② 결재선 · 기관대표자 (한 덩어리로 본다)
    #
    # 결재선의 부서장과 기관대표자는 다르다. 학교는 둘 다 교장이지만
    # 교육지원청은 부서장이 과장, 기관대표자는 교육장이다.
    section_title(ws, IN_SCHOOL_ROW - 2,
                  "1. 기관정보 — 처음 한 번만 채우면 됩니다.  기관구분을 고르면 아랫줄 직위가"
                  " 자동으로 채워집니다 (학교: 행정실장 · 교장 / 교육지원청: 팀장 · 과장 · 교육장)",
                  IN_LAST_COL)
    field_block(ws, IN_SCHOOL_ROW - 1, IN_SCHOOL_ROW, SCHOOL_FIELDS)
    field_block(ws, IN_SIGN_ROW - 1, IN_SIGN_ROW, SIGN_FIELDS)

    # ③ 비상연락망
    section_title(ws, IN_EMERG_ROW - 2,
                  "2. 비상연락망 · 관할기관 — 안전보건관리계획서의 비상조치 계획에 들어갑니다",
                  IN_LAST_COL)
    field_block(ws, IN_EMERG_ROW - 1, IN_EMERG_ROW, EMERG_FIELDS)

    # ④ 계약정보
    #
    # 제목 줄 오른쪽을 떼어 건수 요약을 앉힌다. 비고(T열)가 색의 단일 출처이므로
    # 이 숫자는 화면에 칠해진 색과 언제나 같다.
    section_title(ws, IN_HEAD_ROW - 1,
                  "3. 계약정보 — 계약명을 넣으면 계약유형과 작업유형(고소·화기·밀폐·전기)이"
                  " 자동으로 추천해 주며, 사용자가 직접 수정할 수 있음",
                  IN_LAST_COL - 5)
    summary = ws.Range(ws.Cells(IN_HEAD_ROW - 1, IN_LAST_COL - 4),
                       ws.Cells(IN_HEAD_ROW - 1, IN_LAST_COL))
    summary.Merge()
    cnt = ws.Cells(IN_HEAD_ROW - 1, IN_LAST_COL - 4)
    cnt.Formula = (
        f'="진행 "&COUNTIF($T${IN_FIRST_ROW}:$T${IN_LAST_ROW},"진행")'
        f'&"건   ·   완료 "&COUNTIF($T${IN_FIRST_ROW}:$T${IN_LAST_ROW},"완료")'
        f'&"건   ·   총 "&COUNTA($B${IN_FIRST_ROW}:$B${IN_LAST_ROW})&"건"')
    cnt.Font.Name = FONT
    cnt.Font.Size = FONT_NOTE
    cnt.Font.Bold = True
    cnt.Font.Color = SECTION_FG
    cnt.HorizontalAlignment = -4152        # xlRight
    cnt.VerticalAlignment = -4108

    table_header(ws, IN_HEAD_ROW, CONTRACT_HEADERS)
    ws.Rows(IN_HEAD_ROW).RowHeight = ROW_H_CONTRACT_HEAD

    for r in range(IN_FIRST_ROW, IN_LAST_ROW + 1):
        ws.Cells(r, 1).Value = r - IN_FIRST_ROW + 1
        ws.Cells(r, 1).HorizontalAlignment = -4108
        ws.Rows(r).RowHeight = ROW_H_DATA
        ws.Range(ws.Cells(r, WORK_FIRST_COL),
                 ws.Cells(r, WORK_LAST_COL)).HorizontalAlignment = -4108
        ws.Cells(r, CC_시작일).HorizontalAlignment = -4108
        ws.Cells(r, CC_종료일).HorizontalAlignment = -4108
        ws.Cells(r, CC_점검요일).HorizontalAlignment = -4108
        ws.Cells(r, CC_계약금액).NumberFormat = "#,##0"
        ws.Cells(r, CC_시작일).NumberFormat = "yyyy-mm-dd"
        ws.Cells(r, CC_종료일).NumberFormat = "yyyy-mm-dd"

    # 되살리기 — 읽어 둔 계약을 새 열 위치로
    for i, row in enumerate(old_contracts[:IN_LAST_ROW - IN_FIRST_ROW + 1]):
        r = IN_FIRST_ROW + i
        for key, col in NEW_COLS.items():
            v = row.get(key)
            if v is not None:
                ws.Cells(r, col).Value = v

    # 기관정보 되살리기 — 값은 병합된 칸의 첫 열에 쓴다
    s = old_school
    for row, keys, fields in ((IN_SCHOOL_ROW, SCHOOL_KEYS, SCHOOL_FIELDS),
                              (IN_SIGN_ROW, SIGN_KEYS, SIGN_FIELDS),
                              (IN_EMERG_ROW, EMERG_KEYS, EMERG_FIELDS)):
        for key, (_, c1, _c2) in zip(keys, fields):
            if s.get(key) is not None:
                ws.Cells(row, c1).Value = s[key]

    body = ws.Range(ws.Cells(IN_HEAD_ROW, 1), ws.Cells(IN_LAST_ROW, IN_LAST_COL))
    body.VerticalAlignment = -4108
    thin_borders(body)

    # 열을 좁혔으므로 계약명·주소·업체명이 모두 두 줄로 감긴다. 감기지 않으면
    # 글자가 잘려 자기가 쓴 값을 못 읽는다. (sheet_area 가 WrapText=False 를
    # 깔아 두므로 반드시 그 뒤에 건다)
    #
    # 계약유형(S)만 뺀다. 사전에서 자동으로 채우는 값이라 두 줄로 접히면
    # 무엇이 걸렸는지 한눈에 안 들어온다.
    wrap = ws.Range(ws.Cells(IN_FIRST_ROW, CC_계약명),
                    ws.Cells(IN_LAST_ROW, CC_계획서유형))
    wrap.WrapText = True
    wrap.HorizontalAlignment = -4108
    wrap.VerticalAlignment = -4108

    remark = ws.Range(ws.Cells(IN_FIRST_ROW, CC_비고), ws.Cells(IN_LAST_ROW, CC_비고))
    remark.WrapText = True
    remark.HorizontalAlignment = -4108
    remark.VerticalAlignment = -4108

    types = ws.Range(ws.Cells(IN_FIRST_ROW, CC_계약유형),
                     ws.Cells(IN_LAST_ROW, CC_계약유형))
    types.WrapText = False
    types.HorizontalAlignment = -4108
    types.VerticalAlignment = -4108

    # 드롭다운 + 설명 말풍선
    set_list(ws.Cells(IN_SCHOOL_ROW, 1), ORG_KINDS, "기관구분",
             "학교 또는 교육지원청을 고르면 결재선 직위가 자동으로 채워집니다.")
    set_list(ws.Range(ws.Cells(IN_FIRST_ROW, WORK_FIRST_COL),
                      ws.Cells(IN_LAST_ROW, WORK_LAST_COL)), "○,×",
             WORK_TIP_TITLE, WORK_TIP)
    set_list(ws.Range(ws.Cells(IN_FIRST_ROW, CC_점검요일),
                      ws.Cells(IN_LAST_ROW, CC_점검요일)), WEEKDAYS,
             "점검요일", "순회점검표·합동점검표를 이 요일로 만듭니다.")
    set_list(ws.Range(ws.Cells(IN_FIRST_ROW, CC_계획서유형),
                      ws.Cells(IN_LAST_ROW, CC_계획서유형)), "자동,상세,공사,용역",
             "계획서 유형", "'자동' 이면 계약기간과 위험작업 수를 보고 알아서 고릅니다.")
    set_list(ws.Range(ws.Cells(IN_FIRST_ROW, CC_계약유형),
                      ws.Cells(IN_LAST_ROW, CC_계약유형)), "=SET_계약유형목록",
             TYPE_TIP_TITLE, TYPE_TIP)
    set_list(ws.Range(ws.Cells(IN_FIRST_ROW, CC_비고),
                      ws.Cells(IN_LAST_ROW, CC_비고)), "진행,완료",
             REMARK_TIP_TITLE, REMARK_TIP)

    for i, w in enumerate(CONTRACT_WIDTHS, start=1):
        ws.Columns(i).ColumnWidth = w

    # V열은 비고를 손으로 정했다는 표시를 담는 숨김 칸이다. 사용자가 볼 것이 아니다.
    ws.Columns(CC_비고수동).Hidden = True

    # 구역 가르기 — 예전에는 사이에 빈 줄을 한 줄씩 두었지만, 그만큼 계약정보가
    # 화면에서 밀려났다. 지금은 2·3 구역 제목 줄의 윗변에 가로줄만 긋는다.
    for row in (IN_EMERG_ROW - 2, IN_HEAD_ROW - 1):
        b = ws.Range(ws.Cells(row, 1), ws.Cells(row, IN_LAST_COL)).Borders(XL_EDGE_TOP)
        b.LineStyle = 1
        b.Color = SECTION_LINE
        b.Weight = 2                         # xlThin

    # --- 조건부서식 -----------------------------------------------------------
    #
    # ① 기관정보·비상연락망 — 안 채운 칸은 노랑으로 '여기 채우세요' 라고 말하고,
    #    채우면 회색으로 눌린 것처럼 가라앉는다. 눈이 남은 빈칸만 따라간다.
    #
    #    규칙을 **병합된 칸마다 하나씩, 그 칸의 첫 열을 절대주소로** 건다.
    #    한 줄에 한 규칙을 상대주소로 걸면 엑셀은 병합된 칸을 왼쪽 위 칸의 서식으로
    #    그려 주지만 한셀은 칸마다 따로 따져서, 값이 든 병합 칸이 반쪽만 노랗게
    #    보인다(A 는 채워졌으니 회색, B 는 비었으니 노랑). 절대주소로 걸면 병합된
    #    칸 안의 모든 셀이 같은 답을 내므로 두 프로그램에서 똑같이 보인다.
    for row, fields in ((IN_SCHOOL_ROW, SCHOOL_FIELDS),
                        (IN_SIGN_ROW, SIGN_FIELDS),
                        (IN_EMERG_ROW, EMERG_FIELDS)):
        for _label, c1, c2 in fields:
            rng = ws.Range(ws.Cells(row, c1), ws.Cells(row, c2))
            first = addr(row, c1)                      # 예: $A$4 — 절대주소
            add_cf(rng, f'={first}=""', EMPTY_YELLOW, stop_if_true=True)
            add_cf(rng, f'={first}<>""', FILLED_GRAY, stop_if_true=True)

    # ② 계약 줄 — 비고 값에 따라 줄 전체를 칠한다
    progress_cf(ws.Range(ws.Cells(IN_FIRST_ROW, 1),
                         ws.Cells(IN_LAST_ROW, IN_LAST_COL)),
                f"${col_letter(CC_비고)}{IN_FIRST_ROW}")

    # 시작일·종료일 머리글의 '(더블=달력)' 은 제목이 아니라 안내라 작고 가늘게 쓴다.
    # 머리글 전체에 Font 를 건 뒤에 해야 지워지지 않는다.
    for c in (CC_시작일, CC_종료일):
        cel = ws.Cells(IN_HEAD_ROW, c)
        i = str(cel.Value).find(SUBLABEL)
        if i >= 0:
            sub = characters(cel, i + 1, len(SUBLABEL))
            sub.Font.Size = FONT_NOTE
            sub.Font.Bold = False

    # U열 [▶ 출력관리] 버튼 — 한 번 누르면 이동·자동 체크·강조까지 한다
    for r in range(IN_FIRST_ROW, IN_LAST_ROW + 1):
        cel = ws.Cells(r, CC_출력버튼)
        shp = ws.Shapes.AddShape(5, cel.Left + 3, cel.Top + 3,
                                 cel.Width - 6, cel.Height - 6)
        shp.Name = f"goto_{r}"
        shp.TextFrame2.TextRange.Text = "▶ 출력"
        shp.TextFrame2.TextRange.Font.Name = FONT
        shp.TextFrame2.TextRange.Font.Size = 10
        shp.TextFrame2.TextRange.Font.Bold = True
        shp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = WHITE
        shp.TextFrame2.WordWrap = False
        shp.Fill.ForeColor.RGB = NEUTRAL
        shp.Fill.Solid()
        shp.Line.Visible = False
        shp.Shadow.Visible = False
        shp.OnAction = "Btn_GoToOutputRow"

    # 틀 고정 — **행만** 얼린다(첫 계약 줄). 계약 100건을 아래로 훑어도 머리글이
    # 남는다. 열은 얼리지 않는다. 가로로 넘길 때 왼쪽이 붙잡혀 있으면 오른쪽
    # 절반을 볼 때 오히려 화면이 좁아진다.
    ws.Activate()
    wb.Windows(1).FreezePanes = False
    # 화면이 아래로 밀려 있으면 그 자리에서 얼어붙는다. 반드시 맨 위로 되돌린 뒤 건다.
    wb.Windows(1).ScrollRow = 1
    wb.Windows(1).ScrollColumn = 1
    ws.Range(f"A{IN_FIRST_ROW}").Select()
    wb.Windows(1).FreezePanes = True
    ws.Range("A1").Select()

    return len(old_contracts)


# -------------------------------------------------------------- 출력관리

# 열 너비는 여기 적힌 줄 중 가장 긴 것이 정한다. 낱말 사이 빈칸 하나가 열 16개를
# 한꺼번에 넓히므로, 붙여 써도 읽히는 곳은 붙여 쓴다.
FORM_HEADERS = [
    "1.안전보건\n관리계획서", "2.수급인\n수준평가표",
    # 3번만 세 줄이다. '안전보건점검표' 는 열 너비 14.5 에 한 줄로 안 들어가
    # 어차피 감기는데, 어디서 감길지 엑셀에 맡기면 '표' 한 글자만 넘어가 볼썽사납다.
    "3.공사(용역)\n안전보건\n점검표", "4.점검표\n(고소작업)", "5.점검표\n(화기작업)",
    "6.점검표\n(밀폐공간)", "7.점검표\n(전기작업)", "8.안전보건\n협의체회의록",
    "9.작업장\n순회점검표",
    "10.작업장\n합동점검표", "11.아차사고\n제안서", "12.관리감독자\n평가표",
    "13.중대재해\n발생보고서", "14.중대재해\n사고보고서", "15.계약이행\n서약서",
]

RULE_TEXT = [
    "2일 이상", "2일 이상", "모든 계약",
    "고소작업 시", "화기작업 시", "밀폐공간 시", "전기작업 시",
    "30일 이상\n(월 1회)", "7일 이상\n(달마다 1장)", "90일 이상\n(분기 1회)",
    "연 1~2회", "연 1~2회", "사고 발생 시", "사고 발생 시", "계약 체결 시",
]


def form_formula(out_row, form_no):
    """출력관리 판정식. 입력정보 행 = 출력관리 행 - OUT_ROW_OFFSET

    번호는 화면 번호다. 8번(협의체 회의록)이 3번 자리에서 옮겨 오면서
    그 사이 넷이 하나씩 앞으로 당겨졌다 — modHwpExport.TemplateFileFor 주석 참고.
    이름 정의(기준_협의체 등)는 그대로 쓰므로 설정 시트는 손댈 것이 없다.
    """
    ir = out_row - OUT_ROW_OFFSET
    days = f"IFERROR((입력정보!$D${ir}-입력정보!$C${ir})+1,0)"
    guard = f'IF($C{out_row}="","",'

    if form_no == 1:
        cond = f"{days}>=기준_계획서"
    elif form_no == 2:
        cond = f"{days}>=기준_수준"
    elif form_no == 3:
        cond = f"{days}>=기준_도급"
    elif form_no in (4, 5, 6, 7):
        col = col_letter(WORK_FIRST_COL + form_no - 4)
        cond = f'AND({days}>=1,입력정보!${col}${ir}="○")'
    elif form_no == 8:
        cond = f"{days}>=기준_협의체"
    elif form_no == 9:
        cond = f"{days}>=기준_순회"
    elif form_no == 10:
        cond = f"{days}>=기준_합동"
    else:
        # 11~15 는 기간으로 정할 수 없다. 버튼이 곧 선택이므로 안내만 둔다.
        return ""
    return f'={guard}IF({cond},"○",""))'


def rebuild_output(wb):
    ws = wb.Worksheets("출력관리")

    for shp in list(ws.Shapes):
        shp.Delete()
    for cb in list(ws.CheckBoxes()):
        try:
            cb.Delete()
        except Exception:
            pass
    sheet_area(ws, OUT_LAST_ROW + 10, OC_요약 + 4)

    last_col = OC_요약

    sheet_title(ws, last_col,
                "출력관리 — 계약을 고르고 버튼을 누르면 한글 서식이 채워진 채 열립니다")

    # 2행 — 제목과 버튼 사이를 한 줄 띄운다. 설명은 두지 않는다(화면이 이미 직관적이다).
    ws.Rows(2).RowHeight = 26

    # 헤더 — 서식 묶음마다 색이 다르고, 위 버튼과 같은 색이다.
    # 왼쪽 다섯 칸(선택·연번·계약명·계약기간·점검요일)은 1~7 과 같은 남색으로 맞춘다.
    # 색을 하나라도 줄여야 화면이 어지럽지 않다.
    heads = ["선택", "연번", "계약명", "계약기간", "점검요일"] + FORM_HEADERS + ["필요 서식"]
    for i, h in enumerate(heads, start=1):
        c = ws.Cells(OUT_HEAD_ROW, i)
        c.Value = h
        c.Font.Bold = True
        c.Font.Size = 11
        if OC_서식1 <= i <= OC_서식1 + 14:
            bg = form_color(i - OC_서식1 + 1)
        else:
            bg = NAVY          # 왼쪽 네 칸과 '필요 서식'
        c.Interior.Color = bg
        # 연회색 바탕에 흰 글씨는 안 보인다
        c.Font.Color = SECTION_FG if bg == LIGHT else WHITE
        c.HorizontalAlignment = -4108
        c.VerticalAlignment = -4108
        c.WrapText = True
    ws.Rows(OUT_HEAD_ROW).RowHeight = ROW_H_OUT_HEAD

    # 판정기준 안내 행
    rules = ["", "", "판정 기준 →", "", ""] + RULE_TEXT + [""]
    for i, t in enumerate(rules, start=1):
        c = ws.Cells(OUT_RULE_ROW, i)
        c.Value = t
        c.Font.Size = FONT_RULE
        c.Interior.Color = RULE_BG
        c.HorizontalAlignment = -4108
        c.VerticalAlignment = -4108
        c.WrapText = True
    ws.Rows(OUT_RULE_ROW).RowHeight = 32

    # 데이터 행
    for r in range(OUT_FIRST_ROW, OUT_LAST_ROW + 1):
        ir = r - OUT_ROW_OFFSET
        ws.Cells(r, 2).Formula = f'=IF(입력정보!$B${ir}="","",입력정보!$A${ir})'
        ws.Cells(r, 3).Formula = f'=IF(입력정보!$B${ir}="","",입력정보!$B${ir})'
        # 계약기간은 '365일' 처럼 일수로만 보여 준다. 서식 필요 판정이 모두 일수
        # 기준(2일·30일·7일·90일)이라, 날짜 두 개보다 일수 하나가 바로 읽힌다.
        #
        # 빈 칸은 0 이므로 날짜가 없을 때 그냥 빼면 '1일' 이 나온다. 두 칸이
        # 실제로 채워졌는지 먼저 본다.
        ws.Cells(r, OC_계약기간).Formula = (
            f'=IF($C{r}="","",'
            f'IF(OR(입력정보!$C${ir}="",입력정보!$D${ir}=""),"",'
            f'IF(입력정보!$D${ir}<입력정보!$C${ir},"",'
            f'TEXT(입력정보!$D${ir}-입력정보!$C${ir}+1,"#,##0")&"일")))')
        for f in range(1, 16):
            fml = form_formula(r, f)
            if fml:
                ws.Cells(r, OC_서식1 + f - 1).Formula = fml
        first, last = addr(r, OC_서식1), addr(r, OC_서식1 + 9)
        ws.Cells(r, OC_요약).Formula = (
            f'=IF($C{r}="","",COUNTIF({first}:{last},"○")&"종")')

        ws.Cells(r, 1).Value = False
        # 체크박스가 연결된 셀의 TRUE/FALSE 는 감춘다. 글자색을 희게 하는 것만으로는
        # 칸이 좁을 때 '####' 가 뜬다. 표시 형식을 비워 아예 그리지 않게 한다.
        ws.Cells(r, 1).NumberFormat = ";;;"
        ws.Cells(r, 2).HorizontalAlignment = -4108
        ws.Cells(r, OC_계약기간).HorizontalAlignment = -4108
        ws.Cells(r, OC_점검요일).HorizontalAlignment = -4108
        ws.Range(ws.Cells(r, OC_서식1), ws.Cells(r, OC_요약)).HorizontalAlignment = -4108
        ws.Rows(r).RowHeight = ROW_H_DATA
        if (r - OUT_FIRST_ROW) % 2 == 1:
            ws.Range(ws.Cells(r, 1), ws.Cells(r, OC_요약)).Interior.Color = STRIPE

    set_list(ws.Range(ws.Cells(OUT_FIRST_ROW, OC_점검요일),
                      ws.Cells(OUT_LAST_ROW, OC_점검요일)), WEEKDAYS,
             "점검요일", "여기서 바꾸면 입력정보에도 그대로 반영됩니다.\n"
                         "순회점검표·합동점검표를 이 요일로 만듭니다.")

    body = ws.Range(ws.Cells(OUT_HEAD_ROW, 1), ws.Cells(OUT_LAST_ROW, OC_요약))
    body.VerticalAlignment = -4108          # 아랫맞춤이면 ○ 가 줄 바닥에 붙는다
    thin_borders(body)

    ws.Columns(1).ColumnWidth = 6
    ws.Columns(2).ColumnWidth = 6
    ws.Columns(3).ColumnWidth = 32
    ws.Columns(OC_계약기간).ColumnWidth = 11
    ws.Columns(OC_점검요일).ColumnWidth = 11

    # 서식 15종 14.5 · 필요 서식 6.
    # 예전에는 머리글을 빈 칸에 넣고 AutoFit 으로 실측해 맞췄는데, 가장 긴 머리글
    # 하나가 열 열여섯 개를 다 넓혀 버려 화면이 넘쳤다. 눈으로 확인한 값으로 못박는다.
    for c in range(OC_서식1, OC_서식15 + 1):
        ws.Columns(c).ColumnWidth = 14.5
    # 4~7(고소·화기·밀폐·전기 점검표)은 머리글이 짧아 줄일 여유가 있다.
    for formNo in (4, 5, 6, 7):
        ws.Columns(OC_서식1 + formNo - 1).ColumnWidth = 13
    ws.Columns(OC_요약).ColumnWidth = 6

    # 계약명도 길다. 줄바꿈해 한눈에 다 읽히게 한다.
    names = ws.Range(ws.Cells(OUT_FIRST_ROW, 3), ws.Cells(OUT_LAST_ROW, 3))
    names.WrapText = True
    names.HorizontalAlignment = -4108
    names.VerticalAlignment = -4108

    # 계약 줄 색칠 — 입력정보의 비고 값을 그대로 따른다. 완료된 계약이
    # 출력관리에서도 가라앉아 보여야 어느 것을 지금 출력해야 하는지 헷갈리지 않는다.
    progress_cf(ws.Range(ws.Cells(OUT_FIRST_ROW, 1),
                         ws.Cells(OUT_LAST_ROW, OC_요약)),
                f"입력정보!${col_letter(CC_비고)}{OUT_FIRST_ROW - OUT_ROW_OFFSET}")

    # 선택 체크박스
    #
    # 네모를 키워 달라는 요청이 있었지만 **키울 수가 없다.** 한셀은 폼 컨트롤
    # 체크박스의 네모를 늘 같은 크기(11px 남짓)로 그린다 ― 도형 크기를 12 로
    # 하든 40 으로 하든 네모는 그대로고 위치만 바뀐다. 글꼴로 키우는 길도
    # 엑셀이 막는다(캡션이 없는 컨트롤이라 Font.Size 를 못 세운다).
    for r in range(OUT_FIRST_ROW, OUT_LAST_ROW + 1):
        cel = ws.Cells(r, 1)
        cb = ws.CheckBoxes().Add(cel.Left + (cel.Width - CHECKBOX_SIZE) / 2,
                                 cel.Top + (cel.Height - CHECKBOX_SIZE) / 2,
                                 CHECKBOX_SIZE, CHECKBOX_SIZE)
        cb.Name = f"sel_{r}"
        cb.Caption = ""
        cb.LinkedCell = f"출력관리!{addr(r,1)}"
        cb.Value = -4146          # xlOff

    # 버튼 — 한 줄에 아홉 개. 헤더와 색을 맞춘다.
    # 서식 열이 F부터이므로 버튼도 F 위에서 시작해 머리글과 세로줄이 맞는다.
    ws.Rows(OUT_BTN_ROW).RowHeight = BTN_H + 6

    # 버튼 색은 그 서식의 머리글 색과 같다 (form_color). 11~14 는 흰 버튼이라
    # 글자를 짙게 쓰고 테두리를 두른다.
    #
    # 문구는 **짧게** 쓴다. 버튼이 여덟 개에서 아홉 개로 늘어 예전 길이로는 줄이
    # 표를 넘어간다. add_button 이 글자 폭을 재서 버튼 폭을 정하므로 문구를 줄이면
    # 그만큼 폭이 준다. 글꼴 크기(FONT_HEAD = 12)는 절대 줄이지 않는다 —
    # 쓰는 사람이 대개 50대라 작아지면 못 읽는다. 무슨 서식인지는 바로 아래
    # 열 머리글이 온전한 이름으로 말해 주고, 버튼과 머리글은 색이 같다.
    buttons = [
        ("1~7번 출력", "Btn_ExportHwpForms", NAVY),
        ("8.협의체 회의록", "Btn_ExportCouncil", GREEN),
        ("9.순회점검표", "Btn_ExportPatrol", GREEN),
        ("10.합동점검표", "Btn_ExportJoint", GREEN),
        ("11.아차사고", "Btn_Export11", form_color(11)),
        ("12.관리감독자", "Btn_Export12", form_color(12)),
        ("13.발생보고서", "Btn_Export13", form_color(13)),
        ("14.사고보고서", "Btn_Export14", form_color(14)),
        ("15.이행서약서", "Btn_Export15", form_color(15)),
    ]
    # 아홉 개를 **오른쪽 정렬**한다. 마지막 [15.이행서약서] 의 오른쪽 끝을
    # 표의 오른쪽 끝(필요 서식 열)에 맞추고, 나머지를 왼쪽으로 촘촘히 붙인다.
    #
    # 남는 자리를 똑같이 나눠 벌려 놓으면 버튼 사이가 헐거워 아홉 개가 한 덩어리로
    # 안 읽힌다. 굵은 테두리 두께만큼만 띄우면 한 줄의 도구 모음처럼 보인다.
    anchor = f"{col_letter(OC_서식1)}{OUT_BTN_ROW}"
    BTN_GAP = 3.0                       # 굵은 테두리 두께 정도

    shapes = []
    for caption, macro, color in buttons:
        fg = SECTION_FG if color == LIGHT else WHITE
        shapes.append(add_button(ws, anchor, caption, macro, 0, 3, BTN_H, color, fg))

    right = ws.Cells(OUT_BTN_ROW, OC_요약).Left + ws.Cells(OUT_BTN_ROW, OC_요약).Width
    for s in reversed(shapes):
        s.Left = right - s.Width
        right = s.Left - BTN_GAP

    # 넘침 검사 — 버튼 줄이 서식 열(F) 왼쪽을 넘어가면 계약명 위를 덮는다.
    # 나중에 문구를 다시 늘렸을 때 조용히 넘치지 않도록 여기서 알린다.
    left_edge = ws.Cells(OUT_BTN_ROW, OC_서식1).Left
    if shapes[0].Left < left_edge:
        print(f"  ! 버튼 줄이 {int(left_edge - shapes[0].Left)}pt 넘칩니다 — "
              f"buttons 목록의 문구를 더 줄이세요")
    else:
        print(f"  + 버튼 {len(buttons)}개, 여유 {int(shapes[0].Left - left_edge)}pt")

    # 틀 고정 — 가로로 넘겨도 계약명이 왼쪽에 남는다
    ws.Activate()
    wb.Windows(1).FreezePanes = False
    wb.Windows(1).ScrollRow = 1
    wb.Windows(1).ScrollColumn = 1
    ws.Range(f"{col_letter(OC_서식1)}{OUT_FIRST_ROW}").Select()
    wb.Windows(1).FreezePanes = True

    return OC_요약


# ------------------------------------------------------------------ 정리

OLD_SHEETS = ["1.계획", "2.수준", "3.협의", "4.도급", "5.고소", "6.화기", "7.밀폐",
              "8.전기", "9.순회", "10.합동", "11.아차", "12.평가", "13.발생", "14.사고"]


def drop_old_sheets(wb):
    """PDF/Word 내보내기 전용이던 시트들을 지운다. 그 경로는 없앴다."""
    names = [s.Name for s in wb.Worksheets]
    wb.Application.DisplayAlerts = False
    n = 0
    for nm in OLD_SHEETS:
        if nm in names:
            wb.Worksheets(nm).Delete()
            n += 1
    wb.Application.DisplayAlerts = True
    return n


def tidy_names(wb):
    """옛 레이아웃에서 남은 인쇄영역과 깨진 이름 정의를 정리한다."""
    for nm in list(wb.Names):
        try:
            if "#REF" in nm.RefersTo or "#NAME?" in nm.RefersTo:
                print(f"  - 깨진 이름 {nm.Name} 제거")
                nm.Delete()
        except Exception:
            pass

    in_last = col_letter(IN_LAST_COL)
    out_last = col_letter(OC_요약)
    wb.Worksheets("입력정보").PageSetup.PrintArea = f"$A$1:${in_last}${IN_LAST_ROW}"
    wb.Worksheets("출력관리").PageSetup.PrintArea = f"$A$1:${out_last}${OUT_LAST_ROW}"
    for name in ("입력정보", "출력관리", "계약유형사전"):
        ps = wb.Worksheets(name).PageSetup
        ps.Orientation = 2          # 가로
        ps.Zoom = False
        ps.FitToPagesWide = 1
        ps.FitToPagesTall = False
    print(f"  + 인쇄영역 갱신 (입력정보 A1:{in_last}{IN_LAST_ROW} / "
          f"출력관리 A1:{out_last}{OUT_LAST_ROW}, 가로 1쪽 맞춤)")


# ------------------------------------------------------------------ VBA

# modTemplateTags.bas 는 gen_template_tags.py 가 서식 파일에서 뽑아 만든다.
VBA_MODULES = ["modLayout.bas", "modConfig.bas", "modAutoDetect.bas",
               "modProgress.bas", "modTemplateTags.bas", "modHwpExport.bas",
               "modNotice.bas", "modUI.bas"]

# 출력 안내창. MsgBox 는 글자 크기를 정할 수 없어 UserForm 을 만들어 넣는다.
# modNotice 가 이 이름을 그대로 부르므로, 없으면 프로젝트가 컴파일되지 않는다.
NOTICE_FORM = "frmNotice"
NOTICE_CODE = """Option Explicit

' 안내창은 코드가 닫는다. 사용자가 X 로 닫아도 출력은 그대로 진행된다.
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        Cancel = True
        Me.Hide
    End If
End Sub
"""


def build_notice_form(proj):
    """출력 안내용 UserForm 을 만든다. 라벨 하나뿐인 단순한 창이다.

    이미 있으면 지우지 않고 **속을 비워 다시 채운다.** VBComponents.Remove 는
    곧바로 반영되지 않아서, 지운 직후 같은 이름으로 Add 하면 '이름이 이미 있다'며
    실패한다. 빌드를 두 번째 돌릴 때부터 걸린다.
    """
    frm = None
    try:
        frm = proj.VBComponents(NOTICE_FORM)
    except Exception:
        pass

    if frm is None:
        frm = proj.VBComponents.Add(3)      # vbext_ct_MSForm
        frm.Name = NOTICE_FORM
    else:
        d = frm.Designer
        for ctl in list(d.Controls):
            d.Controls.Remove(ctl.Name)
        cm = frm.CodeModule
        if cm.CountOfLines:
            cm.DeleteLines(1, cm.CountOfLines)

    frm.Properties("Caption").Value = "출력 안내"
    frm.Properties("Width").Value = 520
    frm.Properties("Height").Value = 300

    lbl = frm.Designer.Controls.Add("Forms.Label.1", "lblBody", True)
    lbl.Left, lbl.Top = 12, 12
    lbl.Width, lbl.Height = 496, 250
    lbl.WordWrap = True
    lbl.Font.Name = FONT
    lbl.Font.Size = 12

    frm.CodeModule.AddFromString(NOTICE_CODE)
    return frm

# 옛 레이아웃(체크박스·하이퍼링크 방식, PDF/Word 내보내기)을 전제로 짜인 모듈들.
# 남겨 두면 지워진 프로시저를 부르다 컴파일이 깨진다.
# modCalendar/frmCalendar 는 날짜 고르기 기능이라 그대로 둔다.
#
# modTopMost 는 안내창을 한글 창 위로 올리려고 Win32 API 를 Declare 하던 모듈이다.
# 한셀은 Declare 를 못 받아 '컴파일 오류: ( 를 찾을 수 없습니다' 로 멈추고, 그러면
# 통합문서의 **모든** 매크로가 죽는다(도형을 눌러도 '이 개체에 지정된 매크로를
# 실행할 수 없습니다' 만 나온다). 지금은 NoticeTopMost 를 부르는 곳도 없어서
# 지워도 잃는 것이 없다.
OBSOLETE = ["modWorkTypeCheckboxes", "modContentPopulation", "modSheetBuilder",
            "modSheetBuilder2", "VBA_Main", "modContentLib", "modOutputSheet",
            "modRowManager", "modLinkPopulate", "modExportOffice", "modTopMost"]

# 시트 모듈에 넣을 코드 (탭 이름 → 파일). 여기 없는 시트 모듈은 코드를 비운다.
SHEET_CODE = {"입력정보": "Sheet_입력정보.cls.txt",
              "출력관리": "Sheet_출력관리.cls.txt"}


def sheet_component(wb, proj, tab_name):
    """탭 이름으로 시트의 코드 모듈을 찾는다.

    시트 문서모듈은 VBComponents 안에서 CodeName(Sheet1 …)으로 불린다.
    Properties("Name")은 비어 있을 때가 있어 믿을 수 없으므로 CodeName 으로 찾는다.
    """
    try:
        code_name = wb.Worksheets(tab_name).CodeName
    except Exception:
        return None
    try:
        return proj.VBComponents(code_name)
    except Exception:
        return None


def set_module_code(comp, code):
    cm = comp.CodeModule
    if cm.CountOfLines:
        cm.DeleteLines(1, cm.CountOfLines)
    if code:
        cm.AddFromString(code)


def read_vba(fn):
    path = os.path.join(VBA_DIR, fn)
    if not os.path.exists(path):
        return None
    with open(path, encoding="utf-8") as f:
        return f.read()


DECL_RE = re.compile(
    r"^\s*(Public|Private)\s+(Const|Type|Enum)\b|^\s*(Public|Private|Dim)\s+\w+\s+As\s", re.I)
PROC_RE = re.compile(
    r"^\s*(Public\s+|Private\s+|Friend\s+)?(Sub|Function|Property)\s+\w", re.I)
ENDPROC_RE = re.compile(r"^\s*End\s+(Sub|Function|Property)\b", re.I)


def bad_chars(fn, lineno, line):
    """VBA 원본에 한국어 코드페이지가 담지 못하는 글자가 있는지 본다.

    VBA 는 코드를 유니코드가 아니라 시스템 코드페이지(한국어 = cp949)로 담는다.
    거기 없는 글자는 조용히 '?' 가 되어 버린다. 주석이면 보기 싫은 정도지만
    문자열 안이면 사용자에게 그대로 '?' 가 보인다.
      · em dash(—) 는 cp949 에 없다  → horizontal bar(―) 를 쓸 것
      · en dash(–), ⏎, ✎ 도 없다
    """
    out = []
    for ch in line:
        try:
            ch.encode("cp949")
        except UnicodeEncodeError:
            out.append(f"{fn}:{lineno}  cp949 에 없는 글자 U+{ord(ch):04X} {ch!r} "
                       f"— VBA 에 넣으면 '?' 가 된다")
    return out


def lint_vba():
    """모듈 수준 선언이 첫 프로시저보다 뒤에 오면 미리 잡아낸다.

    VBA 는 이걸 컴파일 때 알리지 않고 **그 코드를 실행하려 할 때** 조용히 멈춘다.
    엑셀이 보이지 않는 상태면 오류창도 안 보여서 원인을 찾는 데 한참 걸린다.
    """
    problems = []
    for fn in sorted(os.listdir(VBA_DIR)):
        path = os.path.join(VBA_DIR, fn)
        first_proc, depth = None, 0
        with open(path, encoding="utf-8") as f:
            for i, line in enumerate(f, 1):
                problems += bad_chars(fn, i, line)
                if line.lstrip().startswith("'"):
                    continue
                if PROC_RE.match(line):
                    depth = 1
                    if first_proc is None:
                        first_proc = i
                elif ENDPROC_RE.match(line):
                    depth = 0
                elif depth == 0 and first_proc and DECL_RE.match(line):
                    problems.append(f"{fn}:{i}  {line.strip()[:60]}")
    return problems


def inject_vba(wb):
    proj = wb.VBProject

    for name in OBSOLETE:
        try:
            proj.VBComponents.Remove(proj.VBComponents(name))
            print(f"  - {name} 제거")
        except Exception:
            pass

    for fn in VBA_MODULES:
        path = os.path.join(VBA_DIR, fn)
        if not os.path.exists(path):
            print(f"  ! {fn} 없음 — 건너뜀")
            continue
        mod_name = os.path.splitext(fn)[0]
        try:
            proj.VBComponents.Remove(proj.VBComponents(mod_name))
        except Exception:
            pass
        with open(path, encoding="utf-8") as f:
            code = f.read()
        # Attribute 줄은 Import 전용이라 AddFromString 에서는 빼야 한다
        code = "\n".join(l for l in code.splitlines()
                         if not l.startswith("Attribute VB_Name"))
        comp = proj.VBComponents.Add(1)
        comp.Name = mod_name
        comp.CodeModule.AddFromString(code)
        print(f"  + {mod_name} ({comp.CodeModule.CountOfLines}줄)")

    build_notice_form(proj)
    print(f"  + {NOTICE_FORM} (출력 안내창)")

    # 시트 모듈 — 옛 이벤트 코드가 지워진 프로시저를 부르면 통합문서 전체가
    # 컴파일되지 않는다. 새 코드를 넣을 시트는 넣고, 나머지는 비운다.
    keep_codenames = set()
    for tab, fn in SHEET_CODE.items():
        comp = sheet_component(wb, proj, tab)
        code = read_vba(fn)
        if comp is None:
            print(f"  ! {tab} 시트 모듈을 찾지 못함")
            continue
        set_module_code(comp, code)
        keep_codenames.add(comp.Name)
        print(f"  + {tab} 시트 모듈 ({comp.CodeModule.CountOfLines}줄)")

    for comp in proj.VBComponents:
        if comp.Type != 100 or comp.Name in keep_codenames:
            continue
        # ThisWorkbook 모듈은 이름이 언어에 따라 다르다(현재_통합_문서 등).
        # Worksheets 에서 못 찾은 문서모듈 = ThisWorkbook 으로 본다.
        cm = comp.CodeModule
        is_book = comp.Name not in {wb.Worksheets(i + 1).CodeName
                                    for i in range(wb.Worksheets.Count)}
        if is_book:
            set_module_code(comp, read_vba("ThisWorkbook.cls.txt"))
            print(f"  + {comp.Name} (통합문서 모듈) 교체")
        elif cm.CountOfLines:
            set_module_code(comp, None)
            print(f"  - {comp.Name} 시트 모듈의 옛 코드 제거")


# ------------------------------------------------------------------ main

def main():
    if not os.path.exists(BOOK):
        print("통합문서를 찾을 수 없습니다:", BOOK)
        return 1

    bad = lint_vba()
    if bad:
        print("★ VBA 모듈에 문제가 있습니다 — 모듈 수준 선언은 첫 프로시저보다 위에 두세요:")
        for b in bad:
            print("   ", b)
        return 1

    os.makedirs(BACKUP_DIR, exist_ok=True)
    stamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    backup = os.path.join(BACKUP_DIR, f"중대재해예방 자동화서식_{stamp}.xlsm")
    shutil.copyfile(BOOK, backup)
    print("백업:", os.path.relpath(backup, ROOT))

    xl = win32com.client.Dispatch("Excel.Application")
    xl.Visible = False
    xl.DisplayAlerts = False
    xl.AutomationSecurity = 1          # 매크로 사용 (끝에 컴파일 검사를 해야 한다)
    # 시트를 다시 그리는 동안 시트 이벤트가 돌면 안 된다. 아직 옛 VBA 가 들어 있는
    # 상태라 값을 쓸 때마다 옛 코드가 불려 엉뚱한 곳에서 멈춘다.
    xl.EnableEvents = False
    xl.ScreenUpdating = False
    wb = xl.Workbooks.Open(BOOK)
    try:
        print("\n[정리] 옛 서식 시트 삭제")
        n = drop_old_sheets(wb)
        print(f"  - 1.계획 ~ 14.사고 {n}개 삭제 (PDF/Word 내보내기 전용이었음)")

        print("[설정] 시트 생성")
        build_config(wb)

        print("[계약유형사전] 시트 생성")
        build_dict(wb)
        print(f"  + 계약 {len(CONTRACT_TYPES)}종")

        print("[ContentLibrary] 재구성")
        stash = read_stash(wb)          # 결재선 보관함은 시트를 새로 만들어도 남겨야 한다
        build_content_library(wb, stash)
        print(f"  + 문구 {len(CONTENT_LIBRARY)}줄 (숨김 시트) · 결재선 보관함 유지")

        print("[입력정보] 개편")
        n = rebuild_input(wb)
        print(f"  + 기관정보/결재선/비상연락망/계약 21열 × "
              f"{IN_LAST_ROW - IN_FIRST_ROW + 1}건, 옛 계약 {n}건 이어받음")

        print("[출력관리] 재구성")
        rebuild_output(wb)
        print(f"  + 서식 15종, 버튼 한 줄(남색/초록/연회색), "
              f"계약 {OUT_LAST_ROW - OUT_FIRST_ROW + 1}건")

        print("[정리] 인쇄영역 · 깨진 이름")
        tidy_names(wb)

        print("[VBA] 주입")
        inject_vba(wb)

        wb.Worksheets("입력정보").Activate()
        wb.Save()
        print("\n저장 완료:", os.path.basename(BOOK))

        # 컴파일 오류는 나중에 '엑셀이 VBE 를 띄운 채 멈춤' 으로 나타나 원인 찾기가
        # 어렵다. 여기서 매크로를 하나 불러 바로 걸러낸다.
        print("\n[검사] VBA 컴파일")
        try:
            xl.Run("GetLastExportError")
            print("  통과")
        except Exception as e:
            print(f"  ★실패: {e}")
            print("  → _scripts/vba/*.bas 를 고치고 다시 실행하세요.")
            return 1

        print("[검사] 자동 판정 한 바퀴")
        try:
            xl.EnableEvents = True          # 이제부터는 새 VBA 라 켜도 된다
            xl.Run("RefreshAllWorkTypesSilent")
            xl.Run("RefreshOutputWeekdays")
            xl.EnableEvents = False
            wb.Save()
            print("  통과")
        except Exception as e:
            print(f"  ★실패: {e}")
            return 1
    finally:
        try:
            wb.Close(SaveChanges=False)
        except Exception:
            pass
        xl.EnableEvents = True
        xl.ScreenUpdating = True
        xl.Quit()
    return 0


if __name__ == "__main__":
    sys.exit(main())
